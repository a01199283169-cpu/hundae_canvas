-- Supabase(Postgres) 마이그레이션 — SQLite에서 전환 시 참고용
-- orders / order_items / order_images / import_logs

CREATE TABLE IF NOT EXISTS import_logs (
    id BIGSERIAL PRIMARY KEY,
    source_file TEXT NOT NULL,
    sheet_name TEXT,
    imported_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    order_count INTEGER DEFAULT 0,
    item_count INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS orders (
    id BIGSERIAL PRIMARY KEY,
    import_id BIGINT REFERENCES import_logs(id),
    source_file TEXT NOT NULL,
    sheet_name TEXT NOT NULL,
    sheet_date DATE,
    order_no TEXT NOT NULL,
    order_date DATE,
    platform TEXT,
    file_ref TEXT,
    customer TEXT,
    phone TEXT,
    address TEXT,
    sales NUMERIC(12,2),
    deduct NUMERIC(12,2),
    ship NUMERIC(12,2),
    total NUMERIC(12,2),
    pay_card TEXT,
    pay_transfer TEXT,
    pay_bank TEXT,
    remark TEXT,
    has_image BOOLEAN DEFAULT FALSE,
    order_qty NUMERIC(10,2),
    expected_ship_type TEXT,
    expected_freight TEXT,
    expected_ship_qty NUMERIC(10,2),
    payment_status TEXT CHECK (payment_status IN ('completed', 'pending')),
    deposit_date DATE,
    UNIQUE (source_file, sheet_name, order_no)
);

CREATE INDEX IF NOT EXISTS idx_orders_sheet_date ON orders(sheet_date);
CREATE INDEX IF NOT EXISTS idx_orders_platform ON orders(platform);
CREATE INDEX IF NOT EXISTS idx_orders_deposit_date ON orders(deposit_date);
CREATE INDEX IF NOT EXISTS idx_orders_payment_status ON orders(payment_status);

CREATE TABLE IF NOT EXISTS order_items (
    id BIGSERIAL PRIMARY KEY,
    order_id BIGINT NOT NULL REFERENCES orders(id) ON DELETE CASCADE,
    line_no INTEGER NOT NULL,
    frame TEXT,
    size TEXT,
    width NUMERIC(10,2),
    height NUMERIC(10,2),
    color TEXT,
    plate TEXT,
    acrylic TEXT,
    hook TEXT,
    item_note TEXT,
    qty NUMERIC(10,2),
    unit_price NUMERIC(12,2)
);

CREATE TABLE IF NOT EXISTS order_images (
    id BIGSERIAL PRIMARY KEY,
    order_id BIGINT REFERENCES orders(id) ON DELETE SET NULL,
    source_file TEXT NOT NULL,
    sheet_name TEXT NOT NULL,
    excel_row INTEGER,
    image_file TEXT NOT NULL,
    mapped BOOLEAN DEFAULT FALSE
);
