WITH employer_stats as (
SELECT
    employer,

    MAX(CASE WHEN trans_type = 'PR Deduct' THEN `date` END)   AS last_pr_deduct_date,
    MAX(CASE WHEN trans_type = 'Payment'   THEN `date` END)   AS last_payment_date,

    -- Amount from the most recent PR Deduct
    MAX(CASE WHEN trans_type = 'PR Deduct'
             AND `date` = (SELECT MAX(t2.`date`)
                           FROM financials.remit_employer_transactions t2
                           WHERE t2.employer = t1.employer
                           AND t2.trans_type = 'PR Deduct')
             THEN amount END)                                  AS last_pr_deduct_amt,

    -- Amount from the most recent Payment
    MAX(CASE WHEN trans_type = 'Payment'
             AND `date` = (SELECT MAX(t2.`date`)
                           FROM financials.remit_employer_transactions t2
                           WHERE t2.employer = t1.employer
                           AND t2.trans_type = 'Payment')
             THEN amount END)                                  AS last_payment_amt

FROM financials.remit_employer_transactions t1
GROUP BY employer
ORDER BY employer
) select * from employer_stats e
left join v_remit_employer_transactions_aging a
on e.employer = a.employer
