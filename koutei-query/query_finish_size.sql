-- export_finish_size.py と同じことを psql だけでやる版（参考・検算用）。読み取りのみ。
--
-- 使い方（丸本PC）:
--   psql -h localhost -p 5432 -U koutei_dev -d koutei_kanri_dev -v ON_ERROR_STOP=1 --csv \
--        -c "CREATE TEMP TABLE target(order_no text)" \
--        -c "\copy target FROM '仕上りサイズ未確認_受注番号一覧.csv' WITH (FORMAT csv, HEADER true, ENCODING 'SJIS')" \
--        -f query_finish_size.sql -o 仕上りサイズ_sql版.csv
--
-- * \copy はクライアント側のファイルを一時テーブルに読むだけで、DB には何も書かない
-- * 一時テーブルを作ってから SET default_transaction_read_only = on にする順番（逆だと CREATE TEMP が拒否される）
-- * 5テーブル（timeline_processes / waiting_list / outsourcing_list / delivery_schedule / processing_works）が
--   全部ある前提。無いテーブルがあると SQL 版はエラーになる（Python 版は飛ばす）
-- * pg_input_is_valid() を使うので PostgreSQL 16 以降
--
-- 受注番号は「空白を落とし、先頭の数字だけ取り、前ゼロを落とした値」で突き合わせる
-- （PDF は 08726258、Excel は 8726258、koutei には 08726258（表紙） のような枝付きもある）。
-- 加工内容の並びは Python 版（読んだ順）と違って文字コード順になる。

SET default_transaction_read_only = on;

WITH
t AS (  -- 入力の受注番号（全角数字・「.0」・空白・前ゼロを吸収し、1キー1行に）
    SELECT k, min(order_no) AS order_no
      FROM (
        SELECT order_no,
               regexp_replace(COALESCE(substring(
                   regexp_replace(regexp_replace(translate(order_no, '０１２３４５６７８９', '0123456789'), '\s', '', 'g'), '\.0+$', '')
                   from '^\d+'), ''), '^0+', '') AS k
          FROM target
      ) x
     WHERE k <> ''
     GROUP BY k
),
ev0 AS (  -- タイムライン＋Waiting List（明細は data 列の JSON）
    SELECT id::text AS id, "orderNumber" AS onum, data::text::jsonb AS d
      FROM timeline_processes WHERE pg_input_is_valid(data::text, 'jsonb')
    UNION ALL
    SELECT id::text, "orderNumber", data::text::jsonb
      FROM waiting_list WHERE pg_input_is_valid(data::text, 'jsonb')
),
ev AS (  -- 最上位と data.events[] の要素を同じ形で並べる（MIS 取込の仕上加工名は events[] にしか無い）
    SELECT regexp_replace(COALESCE(substring(regexp_replace(translate(COALESCE(onum, ''), '０１２３４５６７８９', '0123456789'), '\s', '', 'g') from '^\d+'), ''), '^0+', '') AS k, d
      FROM ev0 WHERE jsonb_typeof(d) = 'object'
    UNION ALL
    SELECT regexp_replace(COALESCE(substring(regexp_replace(translate(COALESCE(onum, ''), '０１２３４５６７８９', '0123456789'), '\s', '', 'g') from '^\d+'), ''), '^0+', ''), e
      FROM ev0, jsonb_array_elements(CASE WHEN jsonb_typeof(d->'events') = 'array' THEN d->'events' ELSE '[]'::jsonb END) e
     WHERE jsonb_typeof(e) = 'object'
),
ev_items AS (  -- outsourcing_items[] を1行ずつに
    SELECT k, it->>'company' AS company, it->>'processing_content' AS content
      FROM ev, jsonb_array_elements(CASE WHEN jsonb_typeof(d->'outsourcing_items') = 'array'
                                         THEN d->'outsourcing_items' ELSE '[]'::jsonb END) it
     WHERE jsonb_typeof(it) = 'object'
),
ol AS (SELECT regexp_replace(COALESCE(substring(regexp_replace(translate(COALESCE(order_number, ''), '０１２３４５６７８９', '0123456789'), '\s', '', 'g') from '^\d+'), ''), '^0+', '') AS k, * FROM outsourcing_list),
ds AS (SELECT regexp_replace(COALESCE(substring(regexp_replace(translate(COALESCE(order_number, ''), '０１２３４５６７８９', '0123456789'), '\s', '', 'g') from '^\d+'), ''), '^0+', '') AS k, * FROM delivery_schedule),
pw AS (SELECT regexp_replace(COALESCE(substring(regexp_replace(translate(COALESCE(order_number, ''), '０１２３４５６７８９', '0123456789'), '\s', '', 'g') from '^\d+'), ''), '^0+', '') AS k, * FROM processing_works),
-- 値の掃除: 「-」「未設定」などは NULL（Python 版の clean() と同じ）
raw_sizes AS (
    SELECT k, d->>'finish_size' AS v FROM ev
    UNION ALL SELECT k, finish_size FROM ol
    UNION ALL SELECT k, finish_size FROM ds
),
raw_contents AS (
    SELECT k, d->>'finish_processing_name' AS v FROM ev
    UNION ALL SELECT k, d->>'finish_processing' FROM ev
    UNION ALL SELECT k, d->>'finishProcessing' FROM ev
    UNION ALL SELECT k, d->>'processing_content' FROM ev
    UNION ALL SELECT k, d->>'finish_process' FROM ev
    UNION ALL SELECT k, content FROM ev_items
    UNION ALL SELECT k, processing_content FROM ol
    UNION ALL SELECT k, processing_content FROM ds
    UNION ALL SELECT k, processing_content FROM pw
    UNION ALL SELECT k, classification FROM pw WHERE classification NOT IN ('その他', '内作', '外注')
),
raw_companies AS (
    SELECT k, d->>'outsourcing_company' AS v FROM ev
    UNION ALL SELECT k, d->>'outsource_name' FROM ev
    UNION ALL SELECT k, company FROM ev_items
    UNION ALL SELECT k, outsourcing_company FROM ol
    UNION ALL SELECT k, outsourcing_company FROM ds
),
raw_papers AS (
    SELECT k, d->>'paper_type' AS t, d->>'standard_size' AS z, d->>'paper_weight' AS w FROM ev
    UNION ALL SELECT k, pi->>'paper_type', pi->>'standard_size', pi->>'paper_weight'
      FROM ev, jsonb_array_elements(CASE WHEN jsonb_typeof(d->'order_paper_info') = 'array' THEN d->'order_paper_info' ELSE '[]'::jsonb END) pi
     WHERE jsonb_typeof(pi) = 'object'
    UNION ALL SELECT k, pi->>'paper_type', pi->>'standard_size', pi->>'paper_weight'
      FROM ev, jsonb_array_elements(CASE WHEN jsonb_typeof(d->'papers') = 'array' THEN d->'papers' ELSE '[]'::jsonb END) pi
     WHERE jsonb_typeof(pi) = 'object'
),
sizes     AS (SELECT DISTINCT k, trim(v) AS v FROM raw_sizes     WHERE lower(trim(COALESCE(v, ''))) NOT IN ('', '-', '－', '―', '未設定', '未定', 'なし', '無し', 'none', 'null', 'nan')),
contents  AS (SELECT DISTINCT k, trim(v) AS v FROM raw_contents  WHERE lower(trim(COALESCE(v, ''))) NOT IN ('', '-', '－', '―', '未設定', '未定', 'なし', '無し', 'none', 'null', 'nan')),
companies AS (SELECT DISTINCT k, trim(v) AS v FROM raw_companies WHERE lower(trim(COALESCE(v, ''))) NOT IN ('', '-', '－', '―', '未設定', '未定', 'なし', '無し', 'none', 'null', 'nan', '未手配', '内作')),
papers    AS (
    SELECT DISTINCT k, t, z, w
      FROM (
        SELECT k,
               CASE WHEN lower(trim(COALESCE(t, ''))) IN ('', '-', '－', '―', '未設定', '未定', 'なし', '無し', 'none', 'null', 'nan') THEN NULL ELSE trim(t) END AS t,
               CASE WHEN lower(trim(COALESCE(z, ''))) IN ('', '-', '－', '―', '未設定', '未定', 'なし', '無し', 'none', 'null', 'nan') THEN NULL ELSE trim(z) END AS z,
               CASE WHEN lower(trim(COALESCE(w, ''))) IN ('', '-', '－', '―', '未設定', '未定', 'なし', '無し', 'none', 'null', 'nan') THEN NULL ELSE trim(w) END AS w
          FROM raw_papers
      ) x
     WHERE t IS NOT NULL OR z IS NOT NULL OR w IS NOT NULL
),
flags AS (
    SELECT k,
           COALESCE(lower(d->>'internal_work') IN ('true', '1'), false)
             OR trim(COALESCE(d->>'work_department', '')) IN ('第二工場', '本社', 'POP課') AS internal,
           COALESCE(lower(d->>'outsourcing') IN ('true', '1'), false) AS outsourcing
      FROM ev
    UNION ALL SELECT k, true, false FROM pw
    UNION ALL SELECT k, trim(COALESCE(work_department, '')) IN ('第二工場', '本社', 'POP課'),
                        trim(COALESCE(outsourcing::text, '')) NOT IN ('', '0', 'false') FROM ds
    UNION ALL SELECT k, false, true FROM ol
),
agg AS (
    SELECT t.k, t.order_no,
           (SELECT string_agg(v, ' / ' ORDER BY v) FROM sizes     WHERE sizes.k = t.k)     AS finish_size,
           (SELECT string_agg(v, ' / ' ORDER BY v) FROM contents  WHERE contents.k = t.k)  AS content,
           (SELECT string_agg(v, ' / ' ORDER BY v) FROM companies WHERE companies.k = t.k) AS company,
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
