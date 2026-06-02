from src.database import connect
from src.parser import detect_columns, parse_sheet, _read_row_fields
from src.config_loader import load_config
import openpyxl
from src.main import find_latest_xlsx

c = connect()
empty = c.execute(
    "SELECT COUNT(*) FROM orders WHERE customer IS NULL OR TRIM(customer)='' "
    "OR address IS NULL OR TRIM(address)=''"
).fetchone()[0]
total = c.execute("SELECT COUNT(*) FROM orders").fetchone()[0]
print(f"missing customer/address: {empty}/{total}")
for r in c.execute(
    "SELECT sheet_name, order_no, customer, phone, address, platform FROM orders "
    "WHERE customer IS NULL OR TRIM(customer)='' OR address IS NULL OR TRIM(address)='' LIMIT 15"
):
    print(dict(r))
c.close()

p = find_latest_xlsx()
cfg = load_config()
wb = openpyxl.load_workbook(p, data_only=True)
ws = wb["0601"]
cols = detect_columns(ws, cfg)
print("0601 cols:", cols)
wb.close()
