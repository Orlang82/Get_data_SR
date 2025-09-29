WITH relevant_accounts AS (
    SELECT
        ca.ACCOUNT_ID,
        ca.CONTRACT_ID
    FROM
        SR_BANK.CONTRACT_ACCOUNT ca
    WHERE
        ca.CONTRACT_ID IN (:data_id_ctr) -- Сюда Python подставит чанк ID договоров
        AND ca.ACCOUNT_ID IN (:data_id_acc) -- Сюда Python подставит чанк ID счетов
)
SELECT
    ra.ACCOUNT_ID,
    a.ACCOUNT_NUMBER,
    c.CODE,
    ra.CONTRACT_ID,
    da.ACCOUNTING_TYPE,
    acs.BASE_AMOUNT AS SUM_UAH
FROM
    SR_BANK.ACCOUNT a
JOIN
    relevant_accounts ra ON a.ID = ra.ACCOUNT_ID
JOIN
    SR_BANK.ACCOUNT_SNAPSHOT acs ON a.ID = acs.ACCOUNT_ID
LEFT JOIN
    SR_BANK.CURRENCY c ON a.CURRENCY_ID = c.ID
LEFT JOIN
    SR_BANK.BALANCE_ACCOUNT ba ON a.BALANCE_ID = ba.ID
LEFT JOIN
    SR_BANK.SETUP_DOCTYPE_BACC da ON ba.ID = da.BACC_ID
WHERE
    TRUNC(acs.SNAPSHOT_DATE, 'DD') = TO_DATE(:date_param, 'dd.mm.yyyy')
    AND (
        da.ACCOUNTING_TYPE = 'RESERVE'
        OR (da.ACCOUNTING_TYPE = 'REVALUATE' AND a.AMOUNT_TYPE = 'P')
    )


