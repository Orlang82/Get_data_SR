SELECT 
    NVL(MAX(dt_ct_group), 'NO_TURNOVER') AS "DT  /  CT",
    NVL(SUM(turnover_amount), 0) AS Turnover
FROM (
    SELECT
        TO_CHAR(SUBSTR(d.ACCOUNT_DT, 1, 4)) || ' / ' || TO_CHAR(SUBSTR(d.ACCOUNT_CT, 1, 4)) AS dt_ct_group,
        CASE 
            WHEN (SUBSTR(d.ACCOUNT_DT, 4, 1) = '9' OR SUBSTR(d.ACCOUNT_CT, 4, 1) = '9') 
            THEN d.BASE_AMOUNT
            ELSE d.BASE_AMOUNT * -1
        END AS turnover_amount
    FROM SR_BANK.DOCUMENT d
    WHERE
        (d.ACCOUNT_CT LIKE '43%' OR d.ACCOUNT_DT LIKE '43%'
         OR d.ACCOUNT_CT LIKE '5031%' OR d.ACCOUNT_DT LIKE '5031%'
         OR d.ACCOUNT_CT LIKE '5041%' OR d.ACCOUNT_DT LIKE '5041%'
         OR d.ACCOUNT_CT LIKE '5011%' OR d.ACCOUNT_DT LIKE '5011%'
         OR d.ACCOUNT_CT LIKE '3521%' OR d.ACCOUNT_DT LIKE '3521%'
        )
        AND d.POST_DATE = TO_DATE(:date_param, 'dd.mm.yyyy')
);