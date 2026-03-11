import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

# ================= CONFIG =================
INPUT_FILE = "СТАРТ.xlsx"
OUTPUT_FILE = "ГОТОВО.xlsx"

DEFAULT_KG = 1.0
DEFAULT_MAIN_COEF = 3.5
DEFAULT_SEC_COEF = 1.1

FONT_ALL = Font(name="Arial Black", size=12)

FILL_WARNING = PatternFill(start_color="FFFFF2CC",
                           end_color="FFFFF2CC",
                           fill_type="solid")

FILL_RED = PatternFill(start_color="FFFF0000",
                       end_color="FFFF0000",
                       fill_type="solid")

FILL_GREEN = PatternFill(start_color="FF92D050",
                         end_color="FF92D050",
                         fill_type="solid")

# ================= HEADERS =================
headers = [
    "Номенклатура",
    "Вес, кг",
    "Розн./кг",
    "Розн./лист",
    "%",
    "МП",
    "коэффициент",
    "МП+",
    "МП+ шт",
    "МП+ шт окр",
    "Размер",
    "м2",
    "ТД ост-к",
    "Закупочная",
    "",
    "0,1 листа",
    "шт",
    "0,2 листа",
    "шт",
    "0,5 листа",
    "шт",
    "0,75 листа",
    "шт",
    "2 листа",
    "шт",
    "5 листа",
    "шт",
    "0",
    "цена",
    "шт"
]

# ================= PARSERS =================

def extract_weight(name: str):
    # Match both integer and decimal kg values: "3,0кг", "18кг", "9,45кг"
    # The "кг" label distinguishes real weights from the sale-unit ", кг" at end of name
    m = re.findall(r"(\d+(?:,\d+)?)\s*кг", name)
    if m:
        return float(m[-1].replace(",", ".")), False
    return DEFAULT_KG, True


def extract_dims(name: str):
    # Requires 3-4 digit numbers to avoid matching profile dimensions like "20х10мм"
    m = re.findall(r"(\d{3,4})\s*[xх]\s*(\d{3,4})", name)
    if m:
        return int(m[-1][0]), int(m[-1][1])
    return 0, 0


def safe(v):
    return None if pd.isna(v) else v


def parse_number(v):
    if v is None:
        return None
    s = str(v).replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except:
        return None


def normalize(text):
    return str(text).lower().replace("\n", " ").strip()


# ================= LOAD =================

raw = pd.read_excel(INPUT_FILE, header=None, dtype=str)
raw = raw.fillna("")

# ================= DETECT HEADER ROW =================
# Find the row that contains the nomenclature/name column header

header_row = None

for i in range(len(raw)):
    row_text = " ".join(str(x).lower() for x in raw.iloc[i])
    if "номенклат" in row_text:
        header_row = i
        break

if header_row is None:
    raise Exception("Header row not found (no row containing 'Номенклатура')")

# ================= DETECT DATA START ROW =================
# Scan forward from header_row to find first row with actual price numbers (>= 3 digits)

data_start_row = header_row + 1

for j in range(header_row + 1, min(header_row + 20, len(raw))):
    txt = " ".join(str(x) for x in raw.iloc[j])
    # A data row contains large numbers (prices are hundreds or thousands)
    if re.search(r"\b\d{3,}\b", txt):
        data_start_row = j
        break

# ================= DETECT HEADER BLOCK START =================
# Scan backward from header_row to include multi-row header sections.
# In 1C exports, the header can span multiple rows:
#   Row N-1: "Ценовая группа" ... "Розничная" ... "Закупочная" ...
#   Row N:   "Артикул" ... "Номенклатура..." ... "RUB" ...
#   Row N+1: "Включает НДС" ...
#   Row N+2: "Цена" ... "Остаток" ...

header_block_start = header_row

for i in range(header_row - 1, max(-1, header_row - 8), -1):
    row_vals = [str(v).strip() for v in raw.iloc[i]]
    non_empty = [v for v in row_vals if v]
    if not non_empty:
        # Empty row: stop going backward
        break
    # If the row contains large numbers it's a data row, not a header
    if any(re.search(r"\b\d{3,}\b", v) for v in non_empty):
        break
    header_block_start = i

# ================= BUILD COMBINED COLUMN HEADERS =================
# Concatenate all text from rows header_block_start..data_start_row-1, column by column.
# This ensures "Розничная" (above "Номенклатура" row) and "Остаток" (below it)
# are both visible when searching for columns.

num_cols = raw.shape[1]
combined_header = [""] * num_cols

for row_i in range(header_block_start, data_start_row):
    for col_i in range(num_cols):
        val = normalize(raw.iloc[row_i, col_i])
        if val:
            combined_header[col_i] = (combined_header[col_i] + " " + val).strip()

print(f"Header block rows: {header_block_start} – {data_start_row - 1}")
print(f"Data starts at row: {data_start_row}")
print("Combined column headers (non-empty):")
for i, h in enumerate(combined_header):
    if h:
        print(f"  col {i}: {h!r}")

# ================= DYNAMIC COLUMN DETECTION =================

def find_column(keywords, used_set):
    """
    Find best-matching column by keyword scoring against combined header values.

    keywords: list of strings (each scored 10), OR list of (keyword, score) tuples.
    Higher scores win; in a tie the lower column index wins.
    """
    if keywords and isinstance(keywords[0], str):
        kw_scores = [(kw, 10) for kw in keywords]
    else:
        kw_scores = keywords  # already (keyword, score) tuples

    best_col = None
    best_score = 0

    for i, header in enumerate(combined_header):
        if i in used_set:
            continue
        score = sum(sc for kw, sc in kw_scores if kw in header)
        if score > best_score:
            best_score = score
            best_col = i

    if best_col is not None:
        used_set.add(best_col)

    return best_col


used_columns = set()

name_col = find_column(
    ["номенклат", "товар", "продукт"],
    used_columns
)

# Retail price: "Розничная" is common; "оптовые" is fallback for files without
# a "Розничная" column (e.g. Фторлвт-type exports that use "ОптовыеRUB").
price_retail_col = find_column(
    [("рознич", 20), ("розничная", 20), ("оптов", 8)],
    used_columns
)

# Stock quantity: strongly prefer the "ТД Ростов" Остаток column over
# "Склад УСН" or similar internal warehouse columns.
# NOTE: "налич" excluded — "Наличие" contains text ("Под заказ"), not numbers.
amount_left_col = find_column(
    [("тд ростов", 20), ("остаток", 10), ("склад", 5)],
    used_columns
)

buy_price_col = find_column(
    ["закуп", "себестоим"],
    used_columns
)

if name_col is None:
    raise Exception("Name column not detected")

print("\nDetected columns:")
print("  Name col:         ", name_col)
print("  Retail price col: ", price_retail_col)
print("  Stock col:        ", amount_left_col)
print("  Buy price col:    ", buy_price_col)

# ================= BUILD DATA DATAFRAME =================

df = raw.iloc[data_start_row:].reset_index(drop=True)
df.columns = range(df.shape[1])

# ================= TRANSFORM =================

rows = []
kg_default_rows = []
width_map = {}

for _, row in df.iterrows():

    name = str(row[name_col]).strip()

    # Skip completely empty name rows
    if not name:
        continue

    # Detect numeric values in price columns
    retail_val = parse_number(row[price_retail_col]) if price_retail_col is not None else None
    buy_val    = parse_number(row[buy_price_col])    if buy_price_col    is not None else None
    stock_val  = parse_number(row[amount_left_col])  if amount_left_col  is not None else None

    # Skip category separator rows: name present but no price data at all
    # (e.g. "Слоистые пластики", "ТИП 2_35-100" etc.)
    if retail_val is None and buy_val is None:
        continue

    kg, defaulted = extract_weight(name)
    width, height = extract_dims(name)

    excel_row = len(rows) + 2  # row 1 is header

    if defaulted:
        kg_default_rows.append(excel_row)

    if width > 0:
        width_map[excel_row] = width

    new_row = [None] * len(headers)

    new_row[0] = name
    new_row[1] = kg

    new_row[2] = retail_val      # Розн./кг  (column C)
    new_row[12] = stock_val      # ТД ост-к  (column M)
    new_row[13] = buy_val        # Закупочная (column N)

    new_row[4] = DEFAULT_MAIN_COEF   # %          (column E)
    new_row[6] = DEFAULT_SEC_COEF    # коэффициент (column G)

    new_row[10] = f"{width}x{height}" if width else ""  # Размер (column K)

    rows.append(new_row)

pd.DataFrame(rows, columns=headers).to_excel(OUTPUT_FILE, index=False)

# ================= FORMULAS =================

wb = load_workbook(OUTPUT_FILE)
ws = wb.active

for r in range(2, ws.max_row + 1):

    ws[f"D{r}"] = f"=IF(B{r}=0,0,ROUNDUP(C{r}*B{r},0))"
    ws[f"F{r}"] = f"=ROUNDUP(D{r}*E{r},0)"
    ws[f"H{r}"] = f"=ROUNDUP(F{r}*G{r},0)"

    ws[f"I{r}"] = f"=IF(B{r}=0,0,M{r}/B{r})"
    ws[f"J{r}"] = f"=ROUNDDOWN(I{r},0)"

    if r in width_map:
        w = width_map[r]
        ws[f"L{r}"] = f"=ROUNDUP(H{r}*(1000/{w}),0)"
    else:
        ws[f"L{r}"] = 0

    ws[f"P{r}"] = f"=ROUNDUP(H{r}*0.1,0)"
    ws[f"Q{r}"] = f"=ROUNDDOWN(M{r}/B{r}/0.1,0)"

    ws[f"R{r}"] = f"=ROUNDUP(H{r}*0.2,0)"
    ws[f"S{r}"] = f"=ROUNDDOWN(M{r}/B{r}/0.2,0)"

    ws[f"T{r}"] = f"=ROUNDUP(H{r}*0.5,0)"
    ws[f"U{r}"] = f"=ROUNDDOWN(M{r}/B{r}/0.5,0)"

    ws[f"V{r}"] = f"=ROUNDUP(H{r}*0.75,0)"
    ws[f"W{r}"] = f"=ROUNDDOWN(M{r}/B{r}/0.75,0)"

    ws[f"X{r}"] = f"=ROUNDUP(H{r}*2,0)"
    ws[f"Y{r}"] = f"=ROUNDDOWN(M{r}/B{r}/2,0)"

    ws[f"Z{r}"] = f"=ROUNDUP(H{r}*5,0)"
    ws[f"AA{r}"] = f"=ROUNDDOWN(M{r}/B{r}/5,0)"

    ws[f"AC{r}"] = f"=IF($AB$1=0,0,ROUNDUP(H{r}*$AB$1,0))"
    ws[f"AD{r}"] = f"=IF($AB$1=0,0,ROUNDDOWN(M{r}/B{r}/$AB$1,0))"

# ================= FORMATTING =================

if price_retail_col is None:
    for r in range(2, ws.max_row + 1):
        ws[f"C{r}"].fill = FILL_RED

if amount_left_col is None:
    for r in range(2, ws.max_row + 1):
        ws[f"M{r}"].fill = FILL_RED

if buy_price_col is None:
    for r in range(2, ws.max_row + 1):
        ws[f"N{r}"].fill = FILL_RED

for col in ["C","D","F","H","L","P","R","T","V","X","Z","AC"]:
    for r in range(2, ws.max_row + 1):
        ws[f"{col}{r}"].number_format = '#,##0'

for row in ws.iter_rows():
    for cell in row:
        cell.font = FONT_ALL

for r in kg_default_rows:
    ws[f"B{r}"].fill = FILL_WARNING

ws["AB1"].fill = FILL_GREEN

# Auto column width
for col in ws.columns:
    max_length = 0
    col_letter = get_column_letter(col[0].column)

    for cell in col:
        try:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        except:
            pass

    ws.column_dimensions[col_letter].width = max_length + 2

wb.save(OUTPUT_FILE)

print(f"\nDONE — {len(rows)} rows written to {OUTPUT_FILE}")
