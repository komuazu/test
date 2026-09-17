-- export_finish_size.py の動作確認用ダミーデータ。
-- koutei-kanr30 の本番と同じ列構成（timeline_processes / waiting_list は data 列の JSON に明細が入る）。
-- 本番には絶対に流さないこと（テスト用の空 DB で使う）。

DROP TABLE IF EXISTS timeline_processes, waiting_list, outsourcing_list, delivery_schedule, processing_works;

CREATE TABLE timeline_processes (
    id TEXT PRIMARY KEY, date TEXT, machine INTEGER, "startHour" INTEGER, duration INTEGER,
    name TEXT, "orderNumber" TEXT, data TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE waiting_list (
    id TEXT PRIMARY KEY, "orderNumber" TEXT, name TEXT, data TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE outsourcing_list (
    id SERIAL PRIMARY KEY, order_number TEXT, customer_name TEXT, product_name TEXT, finish_size TEXT,
    quantity INTEGER, department_code TEXT, outsourcing_company TEXT, outsourcing_delivery_date TEXT,
    outsourcing_handover_date TEXT, processing_content TEXT, order_manager TEXT,
    order_manager_department TEXT, comment TEXT, notes TEXT
);
CREATE TABLE delivery_schedule (
    id SERIAL PRIMARY KEY, order_number TEXT, customer_name TEXT, product_name TEXT,
    final_delivery_date TEXT, delivery_method TEXT, work_department TEXT, finish_size TEXT,
    order_quantity TEXT, outsourcing INTEGER, outsourcing_company TEXT, processing_content TEXT
);
CREATE TABLE processing_works (
    id TEXT PRIMARY KEY, order_number TEXT, classification TEXT, client_name TEXT, product_name TEXT,
    processing_content TEXT, quantity INTEGER, start_date TEXT, delivery_date TEXT, progress INTEGER
);

-- 8726258: 本体＋表紙の2部品。内作（第二工場・中綴じ12P）と外注（八王子紙工・ミシン(筋)）の両方
INSERT INTO timeline_processes (id, date, machine, "startHour", duration, name, "orderNumber", data) VALUES
('tp-1', '2026-05-12', 1, 9, 3, 'ﾊﾟﾝﾌﾚｯﾄ', '08726258',
 '{"client_name":"㈱テスト","product_name":"ﾊﾟﾝﾌﾚｯﾄ","part_type":"本体","finish_size":"A4 297×210",
   "paper_type":"A2マット","standard_size":"A全判","paper_weight":"86.5",
   "internal_work":true,"outsourcing":true,"work_department":"第二工場",
   "finish_processing_name":"中綴じ12P",
   "outsourcing_items":[{"company":"八王子紙工","processing_content":"ミシン(筋)","delivery_date":"2026-05-20"}]}'),
('tp-2', '2026-05-12', 1, 12, 1, 'ﾊﾟﾝﾌﾚｯﾄ', '08726258',
 '{"client_name":"㈱テスト","product_name":"ﾊﾟﾝﾌﾚｯﾄ","part_type":"表紙","finish_size":"A4 297×210",
   "paper_type":"A2マット","standard_size":"A全判","paper_weight":"110",
   "internal_work":1,"work_department":"第二工場","finish_processing_name":"中綴じ12P"}'),
-- 返しイベント（同じ受注番号の重複）。値が二重に数えられないこと
('tp-3', '2026-05-13', 1, 9, 1, 'ﾊﾟﾝﾌﾚｯﾄ', '08726258',
 '{"part_type":"本体","finish_size":"A4 297×210","isReturnProcess":true,"internal_work":true,"work_department":"第二工場"}'),
-- 8726262: 受注番号がゼロ埋め無しで入っている。加工情報は delivery_schedule と processing_works 側
('tp-4', '2026-06-01', 2, 9, 2, 'ﾁﾗｼ', '8726262',
 '{"finish_size":"B5","internal_work":false}'),
-- 8726263: 旧形式（outsourcing_company を直接持つ）＋ MIS取込キー（finish_process / outsource_name）
('tp-5', '2026-06-02', 2, 9, 2, 'ﾘｰﾌ', '08726263',
 '{"finish_size":"A3","finish_process":"二つ折り","outsource_name":"松岡製本","outsourcing":1}'),
-- 8726264: data が壊れている行。落ちずに空扱いになること
('tp-6', '2026-06-03', 2, 9, 2, 'x', '08726264', '{broken json');

-- 8726259: Waiting List にしか無い（未配置）。仕上りサイズだけ
INSERT INTO waiting_list (id, "orderNumber", name, data) VALUES
('wl-1', '08726259', 'ﾎﾟｽﾀｰ', '{"finish_size":"B2 728×515","internal_work":false,"outsourcing":false,
   "order_paper_info":[{"partName":"本体","paper_type":"オーロラコート","standard_size":"菊全判","paper_weight":"93.5"}]}');

-- 8726260: 外注依頼書テーブルにしか無い
INSERT INTO outsourcing_list (order_number, customer_name, product_name, finish_size, outsourcing_company, processing_content) VALUES
('08726260', '㈱テスト', '会社案内', 'A4', '松岡製本', '抜き・ポケット貼り・24P中綴じ');

-- 8726262 の加工情報
INSERT INTO delivery_schedule (order_number, finish_size, work_department, outsourcing, outsourcing_company, processing_content) VALUES
('08726262', 'B5', '第二工場', 0, '', '折加工（二つ折り）');
INSERT INTO processing_works (id, order_number, classification, processing_content, quantity) VALUES
('pw-1', '08726262', '折加工1', '二つ折り', 1000);

-- 対象外の受注番号（一覧に無い）。拾わないこと
INSERT INTO timeline_processes (id, date, machine, "startHour", duration, name, "orderNumber", data) VALUES
('tp-9', '2026-06-03', 2, 9, 2, '対象外', '08799999', '{"finish_size":"A5"}');
