-- SQLite. Load loans_clean.csv/raw data into table loans_raw first.
-- STEP 1: CLEANING + SEGMENTATION
DROP TABLE IF EXISTS loans_clean;
CREATE TABLE loans_clean AS
WITH dedup AS (            -- remove duplicate loan_ids
  SELECT * FROM loans_raw WHERE rowid IN (SELECT MIN(rowid) FROM loans_raw GROUP BY loan_id)
), valid AS (              -- drop impossible credit scores
  SELECT * FROM dedup WHERE credit_score BETWEEN 300 AND 850
), stats AS (              -- regional averages used to fill gaps
  SELECT region, AVG(dti) AS avg_dti,
         AVG(CASE WHEN annual_income <= 10000000 THEN annual_income END) AS avg_inc
  FROM valid GROUP BY region
), fixed AS (              -- impute missing income/DTI, treat income > 1 crore as outlier
  SELECT v.loan_id, v.application_date, v.region, v.purpose, v.credit_score,
         CAST(ROUND(CASE WHEN v.annual_income IS NULL OR v.annual_income > 10000000
                         THEN s.avg_inc ELSE v.annual_income END) AS INTEGER) AS annual_income,
         v.loan_amount, ROUND(COALESCE(v.dti, s.avg_dti), 3) AS dti, v.credit_utilization,
         v.approved, v.defaulted, v.fraud_flag, v.processing_days
  FROM valid v JOIN stats s ON s.region = v.region
)
SELECT *, strftime('%Y-%m', application_date) AS app_month,
  CASE WHEN credit_score >= 760 THEN 'A' WHEN credit_score >= 720 THEN 'B'
       WHEN credit_score >= 680 THEN 'C' WHEN credit_score >= 620 THEN 'D' ELSE 'E' END AS grade,
  CASE WHEN dti < 0.2 THEN '<20%' WHEN dti < 0.3 THEN '20-30%' WHEN dti < 0.4 THEN '30-40%' ELSE '40%+' END AS dti_band,
  CASE WHEN credit_utilization < 0.3 THEN '<30%' WHEN credit_utilization < 0.6 THEN '30-60%'
       WHEN credit_utilization < 0.8 THEN '60-80%' ELSE '80%+' END AS util_band,
  CASE WHEN annual_income < 400000 THEN 'Low (<4L)' WHEN annual_income < 800000 THEN 'Mid (4-8L)'
       WHEN annual_income < 1500000 THEN 'Upper-mid (8-15L)' ELSE 'High (15L+)' END AS income_bracket
FROM fixed;

-- ANALYSIS: kpi
SELECT COUNT(*) AS applications, SUM(approved) AS approved, SUM(defaulted) AS defaults,
  ROUND(SUM(CASE WHEN approved=1 THEN loan_amount END)/1e7,0) AS disbursed_cr,
  ROUND(SUM(CASE WHEN defaulted=1 THEN loan_amount*0.6 END)/1e7,0) AS est_loss_cr,
  ROUND(100.0*SUM(fraud_flag)/COUNT(*),2) AS fraud_pct,
  ROUND(100.0*SUM(processing_days>5)/COUNT(*),1) AS sla_breach_pct FROM loans_clean;

-- ANALYSIS: grade
SELECT grade AS segment, COUNT(*) AS applications, SUM(approved) AS approved, SUM(defaulted) AS defaults
FROM loans_clean GROUP BY grade  ORDER BY grade;

-- ANALYSIS: dti
SELECT dti_band AS segment, COUNT(*) AS applications, SUM(approved) AS approved, SUM(defaulted) AS defaults
FROM loans_clean GROUP BY dti_band  ORDER BY MIN(dti);

-- ANALYSIS: util
SELECT util_band AS segment, COUNT(*) AS applications, SUM(approved) AS approved, SUM(defaulted) AS defaults
FROM loans_clean GROUP BY util_band  ORDER BY MIN(credit_utilization);

-- ANALYSIS: income
SELECT income_bracket AS segment, COUNT(*) AS applications, SUM(approved) AS approved, SUM(defaulted) AS defaults
FROM loans_clean GROUP BY income_bracket  ORDER BY MIN(annual_income);

-- ANALYSIS: purpose
SELECT purpose AS segment, COUNT(*) AS applications, SUM(approved) AS approved, SUM(defaulted) AS defaults
FROM loans_clean GROUP BY purpose  ORDER BY purpose;

-- ANALYSIS: region
SELECT region AS segment, COUNT(*) AS applications, SUM(approved) AS approved, SUM(defaulted) AS defaults
FROM loans_clean GROUP BY region  ORDER BY region;

-- ANALYSIS: monthly
SELECT app_month AS segment, COUNT(*) AS applications, SUM(approved) AS approved, SUM(defaulted) AS defaults
FROM loans_clean GROUP BY app_month  ORDER BY app_month;

-- ANALYSIS: season
SELECT substr(app_month,6,2) AS segment, COUNT(*) AS applications, SUM(approved) AS approved, SUM(defaulted) AS defaults
FROM loans_clean GROUP BY substr(app_month,6,2)  ORDER BY substr(app_month,6,2);

-- ANALYSIS: region_ops
SELECT region AS segment, COUNT(*) AS applications,
  ROUND(100.0*SUM(approved)/COUNT(*),1) AS approval_pct,
  ROUND(100.0*SUM(defaulted)/SUM(approved),2) AS default_pct,
  ROUND(AVG(processing_days),1) AS avg_days, ROUND(100.0*SUM(processing_days>5)/COUNT(*),1) AS sla_breach_pct,
  ROUND(100.0*SUM(fraud_flag)/COUNT(*),2) AS fraud_pct FROM loans_clean GROUP BY region ORDER BY default_pct DESC;

-- ANALYSIS: sla_year
SELECT substr(app_month,1,4) AS segment, ROUND(100.0*SUM(processing_days>5)/COUNT(*),1) AS sla_breach_pct,
  ROUND(AVG(processing_days),2) AS avg_days FROM loans_clean GROUP BY 1 ORDER BY 1;

-- ANALYSIS: risk_pockets
SELECT grade || ' grade, DTI ' || dti_band AS segment, COUNT(*) AS applications, SUM(approved) AS approved,
  SUM(defaulted) AS defaults FROM loans_clean GROUP BY grade, dti_band HAVING SUM(approved) >= 150
  ORDER BY 1.0*SUM(defaulted)/SUM(approved) DESC LIMIT 6;