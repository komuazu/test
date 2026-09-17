-- export_finish_size.py と同じことを psql だけでやる版（参考・検算用）。読み取りのみ。
--
--   psql -h localhost -p 5432 -U koutei_dev -d koutei_kanri_dev -v ON_ERROR_STOP=1 \
--        -f query_finish_size.sql -o 仕上りサイズ_sql版.csv
--
-- 先に、受注番号一覧を一時テーブルに入れる（\copy はクライアント側のファイルを読む。DB には書かない）:
--   CREATE TEMP TABLE target(order_no text);
--   \copy target FROM '仕上りサイズ未確認_受注番号一覧.csv' WITH (FORMAT csv, HEADER true, ENCODING 'SJIS')
--
-- 受注番号は「前ゼロ・空白を落とした数字」で突き合わせる（PDF は 08726258、Excel は 8726258）。

SET default_transaction_read_only = on;

WITH t AS (
    SELECT DISTINCT regexp_replace(regexp_replace(order_no, '\s', '', 'g'), '^0+', '') AS k, order_no
    FROM target WHERE order_no ~ '\d'
),
ev AS (  -- タイムライン＋Waiting List（明細は data 列の JSON）
    SELECT regexp_replace(regexp_replace("orderNumber", '\s', '', 'g'), '^0+', '') AS k, data::jsonb AS d
      FROM timeline_processes WHERE pg_input_is_valid(data, 'jsonb')   -- 壊れた JSON の行は飛ばす（PG16 以降）
    UNION ALL
    SELECT regexp_replace(regexp_replace("orderNumber", '\s', '', 'g'), '^0+', ''), data::jsonb
      FROM waiting_list WHERE pg_input_is_valid(data, 'jsonb')
),
ev_items AS (  -- outsourcing_items[] を1行ずつに
    SELECT k, it->>'company' AS company, it->>'processing_content' AS content
      FROM ev, jsonb_array_elements(CASE WHEN jsonb_typeof(d->'outsourcing_items') = 'array'
                                         THEN d->'outsourcing_items' ELSE '[]'::jsonb END) it
),
sizes AS (
    SELECT k, NULLIF(trim(d->>'finish_size'), '') AS v FROM ev
    UNION SELECT regexp_replace(order_number, '^0+', ''), NULLIF(trim(finish_size), '') FROM outsourcing_list
    UNION SELECT regexp_replace(order_number, '^0+', ''), NULLIF(trim(finish_size), '') FROM delivery_schedule
),
contents AS (
    SELECT k, NULLIF(trim(d->>'finish_processing_name'), '') AS v FROM ev
    UNION SELECT k, NULLIF(trim(d->>'finish_processing'), '') FROM ev
    UNION SELECT k, NULLIF(trim(d->>'finishProcessing'), '') FROM ev
    UNION SELECT k, NULLIF(trim(d->>'processing_content'), '') FROM ev
    UNION SELECT k, NULLIF(trim(d->>'finish_process'), '') FROM ev
    UNION SELECT k, NULLIF(trim(content), '') FROM ev_items
    UNION SELECT regexp_replace(order_number, '^0+', ''), NULLIF(trim(processing_content), '') FROM outsourcing_list
    UNION SELECT regexp_replace(order_number, '^0+', ''), NULLIF(trim(processing_content), '') FROM delivery_schedule
    UNION SELECT regexp_replace(order_number, '^0+', ''), NULLIF(trim(processing_content), '') FROM processing_works
    UNION SELECT regexp_replace(order_number, '^0+', ''), NULLIF(trim(classification), '') FROM processing_works
),
companies AS (
    SELECT k, NULLIF(trim(d->>'outsourcing_company'), '') AS v FROM ev
    UNION SELECT k, NULLIF(trim(d->>'outsource_name'), '') FROM ev
    UNION SELECT k, NULLIF(trim(company), '') FROM ev_items
    UNION SELECT regexp_replace(order_number, '^0+', ''), NULLIF(trim(outsourcing_company), '') FROM outsourcing_list
    UNION SELECT regexp_replace(order_number, '^0+', ''), NULLIF(trim(outsourcing_company), '') FROM delivery_schedule
),
papers AS (  -- 用紙（本体 JSON と、部品ごとの order_paper_info[] / papers[]）
    SELECT k, NULLIF(trim(d->>'paper_type'), '') AS t, NULLIF(trim(d->>'standard_size'), '') AS z, NULLIF(trim(d->>'paper_weight'), '') AS w FROM ev
    UNION SELECT k, NULLIF(trim(pi->>'paper_type'), ''), NULLIF(trim(pi->>'standard_size'), ''), NULLIF(trim(pi->>'paper_weight'), '')
      FROM ev, jsonb_array_elements(CASE WHEN jsonb_typeof(d->'order_paper_info') = 'array' THEN d->'order_paper_info'
                                         WHEN jsonb_typeof(d->'papers') = 'array' THEN d->'papers' ELSE '[]'::jsonb END) pi
),
flags AS (
    SELECT k,
           bool_or(COALESCE((d->>'internal_work') IN ('true','1'), false) OR NULLIF(trim(d->>'work_department'), '') IS NOT NULL) AS internal,
           bool_or(COALESCE((d->>'outsourcing') IN ('true','1'), false)) AS outsourcing
      FROM ev GROUP BY k
    UNION ALL SELECT regexp_replace(order_number, '^0+', ''), true, false FROM processing_works
    UNION ALL SELECT regexp_replace(order_number, '^0+', ''), NULLIF(trim(work_department), '') IS NOT NULL, COALESCE(outsourcing, 0) <> 0 FROM delivery_schedule
    UNION ALL SELECT regexp_replace(order_number, '^0+', ''), false, true FROM outsourcing_list
),
agg AS (
    SELECT t.k, t.order_no,
           (SELECT string_agg(DISTINCT v, ' / ') FROM sizes     WHERE sizes.k = t.k AND v IS NOT NULL) AS finish_size,
           (SELECT string_agg(DISTINCT v, ' / ') FROM contents  WHERE contents.k = t.k AND v IS NOT NULL) AS content,
           (SELECT string_agg(DISTINCT v, ' / ') FROM companies WHERE companies.k = t.k AND v IS NOT NULL) AS company,
           (SELECT string_agg(DISTINCT t, ' / ') FROM papers WHERE papers.k = t.k AND t IS NOT NULL) AS paper_type,
           (SELECT string_agg(DISTINCT z, ' / ') FROM papers WHERE papers.k = t.k AND z IS NOT NULL) AS paper_size,
           (SELECT string_agg(DISTINCT w, ' / ') FROM papers WHERE papers.k = t.k AND w IS NOT NULL) AS paper_weight,
           (SELECT bool_or(internal)    FROM flags WHERE flags.k = t.k) AS internal,
           (SELECT bool_or(outsourcing) FROM flags WHERE flags.k = t.k) AS outsourcing
      FROM t
)
SELECT order_no                                   AS "受注番号",
       COALESCE(finish_size, '')                  AS "仕上りサイズ",
       COALESCE(content, '')                      AS "加工内容",
       CASE WHEN company IS NOT NULL OR outsourcing THEN
                 CASE WHEN internal THEN '内作・外注' ELSE '外注' END
            WHEN internal THEN '内作' ELSE '' END AS "内外作区分",
       COALESCE(company, '')                      AS "委託先名",
       COALESCE(paper_type, '')                   AS "用紙銘柄",
       COALESCE(paper_size, '')                   AS "用紙規格",
       COALESCE(paper_weight, '')                 AS "斤量"
  FROM agg
 ORDER BY order_no;
