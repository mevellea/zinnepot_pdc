# process_plan.py
# Usage: python process_plan.py
# Produit: calendrier-2026-processed.csv (résumé avec surfaces, rendements, weekly_avail estimée)

import pandas as pd
import re
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.styles import Alignment, Font

IN_FILE = Path(r"C:\Users\mevel\Desktop\Maraichage\ASBL Zinnepot\PDC\calendrier-2026.xlsx")
OUT_FILE = Path(r"C:\Users\mevel\Desktop\Maraichage\ASBL Zinnepot\PDC\calendrier-2026-processed.xlsx")
OUT_FILE2 = Path(r"C:\Users\mevel\Desktop\Maraichage\ASBL Zinnepot\PDC\calendrier-2026-processed2.xlsx")

# ------------- réglages / hypothèses (modifie ici) -------------
# rendements approximatifs kg/m2 par culture-clé (valeurs indicatives)
DEFAULT_YIELD_PER_M2 = {
    'tomate': 8.0,
    'courgette': 6.0,
    'salade': 1.5,
    'basilic': 0.4,
    'carotte': 1.8,
    'betterave': 1.2,
    'bette': 2.0,
    'haricot': 1.2,
    'aubergine': 3.0,
    'poireau': 1.0,
    'epinard': 1.5,
    'radis': 0.8,
    'fraise': 0.8,
    'chou': 2.5,
    'kale': 2.0
}
FALLBACK_YIELD_PER_M2 = 1.0  # kg/m2 si on ne reconnait pas la culture
# ---------------------------------------------------------------

def parse_cm_to_m(s):
    if pd.isna(s): return None
    s = str(s)
    m = re.search(r'([\d\.,]+)', s)
    if not m:
        return None
    v = m.group(1).replace(',', '.')
    try:
        return float(v) / 100.0
    except:
        return None


def guess_yield(culture):
    if pd.isna(culture): return FALLBACK_YIELD_PER_M2
    c = culture.lower()
    for k,val in DEFAULT_YIELD_PER_M2.items():
        if k in c:
            return val
    return FALLBACK_YIELD_PER_M2


def main():
    xl = pd.ExcelFile(IN_FILE)
    sheet = 'plantings' if 'plantings' in xl.sheet_names else xl.sheet_names[0]
    raw = xl.parse(sheet, header=None)

    # première ligne comme header si c'est le cas
    header = raw.iloc[0].fillna('').astype(str).tolist()
    df = raw[1:].copy()
    df.columns = header

    # normaliser noms de colonnes attendus (ajoute d'autres alias si besoin)
    col_map = {
        'Calendrier': 'culture',
        'Date de Semis en pepiniere': 'sowing_date',
        "Date d'implantation": 'planting_date',
        'Recolte': 'harvest_start',
        'Fenêtre de récolte (jours)': 'harvest_window_days',
        'Type d\'emplacement': 'site_type',
        'Quantité': 'quantity',
        '# Rangs': 'num_rows',
        'Espace sur le rang': 'space_on_row',
        'Espace entre les rangs': 'space_between_rows',
        'Emplacement': 'location',
        'Notes': 'notes'
    }
    for k,v in col_map.items():
        if k in df.columns:
            df.rename(columns={k:v}, inplace=True)

    # trim culture
    if 'culture' in df.columns:
        df['culture'] = df['culture'].astype(str).str.strip()

    # parse dates (on suppose format US-ish si mois > 12 vu dans l'exemple)
    def parse_date(s):
        try:
            return pd.to_datetime(s, dayfirst=False, errors='coerce')
        except:
            return pd.to_datetime(s, dayfirst=True, errors='coerce')
    for c in ['sowing_date','planting_date','harvest_start']:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip().replace({'nan':None}).apply(lambda x: parse_date(x) if x and x.lower()!='none' else pd.NaT)

    for c in ['harvest_window_days','quantity','num_rows']:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')

    # spacings -> mètres
    for c in ['space_on_row','space_between_rows']:
        if c in df.columns:
            df[c + '_m'] = df[c].apply(parse_cm_to_m)

    def parse_bloc(q):
        matches = re.findall(r'(\d+)\s*\(([^)]+)\)', str(q))
        for jardin_str, planches_str in matches:
            return int(jardin_str) // 10
        return None

    def parse_jardin(q):
        matches = re.findall(r'(\d+)\s*\(([^)]+)\)', str(q))
        for jardin_str, planches_str in matches:
            jardin_index = int(jardin_str)
            return jardin_index % 10
        return None

    def parse_planches(q):
        matches = re.findall(r'(\d+)\s*\(([^)]+)\)', str(q))
        for jardin_str, planches_str in matches:
            planches = [int(p.strip()) for p in planches_str.split(',')]
            return planches
        return None

    df['bloc'] = df['location'].apply(parse_bloc)
    df['jardin'] = df['location'].apply(parse_jardin)
    df['planches'] = df['location'].apply(parse_planches)

    def parse_plants(q):
        if pd.isna(q): return None
        s = str(q)
        m = re.search(r'([\d\.,]+)', s)
        if m:
            return float(m.group(1).replace(',', '.'))
        try:
            return float(s)
        except:
            return None
    if 'quantity' in df.columns:
        df['plants_per_unit'] = df['quantity'].apply(parse_plants)
    else:
        df['plants_per_unit'] = 1.0

    df['plants_per_unit'] = df['plants_per_unit'].fillna(1.0)
    df['num_rows'] = df['num_rows'].fillna(1.0)

    def calc_linear_meters(row):
        if pd.isna(row.get('space_on_row_m')) or pd.isna(row.get('num_rows')):
            return None
        return row['plants_per_unit'] * row['space_on_row_m'] * row['num_rows']
    df['linear_meters'] = df.apply(calc_linear_meters, axis=1)

    def calc_area(row):
        if pd.isna(row.get('linear_meters')) or pd.isna(row.get('space_between_rows_m')):
            return None
        return row['linear_meters'] * row['space_between_rows_m']
    df['area_m2'] = df.apply(calc_area, axis=1)

    # yield estimated
    df['yield_per_m2_kg'] = df['culture'].apply(guess_yield) if 'culture' in df.columns else FALLBACK_YIELD_PER_M2
    df['estimated_total_yield_kg'] = df.apply(lambda r: (r['area_m2'] * r['yield_per_m2_kg']) if pd.notna(r.get('area_m2')) else None, axis=1)

    # harvest weeks + weekly availability
    df['harvest_weeks'] = df['harvest_window_days'].apply(lambda x: round(x/7,1) if pd.notna(x) else None)
    df['weekly_availability_kg'] = df.apply(lambda r: (r['estimated_total_yield_kg'] / r['harvest_weeks']) if pd.notna(r.get('estimated_total_yield_kg')) and pd.notna(r.get('harvest_weeks')) and r['harvest_weeks']>0 else None, axis=1)

    df['planting_week'] = df["planting_date"].dt.isocalendar().week
    df['harvest_week_start'] = df["harvest_start"].dt.isocalendar().week
    df['harvest_week_end'] = df["harvest_week_start"] + df['harvest_weeks'].round().astype(int)

    df = df.explode("planches").rename(columns={"planches": "planche"})
    df["item_id"] = (df["bloc"].astype(str)+ "-"+ df["jardin"].astype(str)+ "-"+ df["planche"].astype(str))
    df = df.sort_values(by=["bloc", "jardin", "planche"]).reset_index(drop=True)

    max_index = df["harvest_week_end"].max()
    for i in range(1, max_index + 1):
        df[i] = ""
    for idx, row in df.iterrows():
        planting_week, harvest_week_start, harvest_week_end = row["planting_week"], row["harvest_week_start"], row["harvest_week_end"]

        # X de A à B (inclus)
        if isinstance(planting_week, int):
            for col in range(planting_week, harvest_week_start):
                df.at[idx, col] = "X"

        # - de B+1 à harvest_week_end (inclus)
        for col in range(harvest_week_start, harvest_week_end):
            df.at[idx, col] = "-"

    print(df.head(5).to_string(index=False))
    df.to_excel(OUT_FILE, index=False)

    wb = load_workbook(OUT_FILE)
    ws = wb.active

    starting_label(ws)
    ws.delete_cols(1, 27)

    ws = merge_succesions(wb, ws)
    coloring(ws)
    merge_blocks(ws)
    style(ws)
    background(ws)

    wb.remove(wb['Sheet1'])
    wb.save(OUT_FILE)


def starting_label(ws):
    # Replace label
    for row in ws.iter_rows(min_row=2):  # skip header
        start_ini = True
        start = True
        for cell in row:
            if start_ini:
                label = cell.value
                start_ini = False
            if cell.value == "X":
                cell.value = label if start else cell.value
                start = False


def merge_succesions(wb, ws):
    # Read header
    headers = [cell.value for cell in ws[1]]
    item_id_col_index = headers.index("item_id") + 1
    # Group rows by item_id (ignoring header)
    groups = {}
    for row in ws.iter_rows(min_row=2, values_only=False):
        item_id = row[item_id_col_index - 1].value
        if item_id not in groups:
            groups[item_id] = []
        groups[item_id].append(row)
    # Create new workbook for output
    ws = wb.create_sheet(title="PDC")
    ws.append(headers)
    # Merge logic
    for item_id, rows in groups.items():
        merged_row = []

        for col_idx in range(len(headers)):
            values = []

            for row in rows:
                value = row[col_idx].value
                if value not in (None, ""):
                    values.append(str(value))

            if col_idx > 0 and len(values) > 1:
                raise ValueError(
                    f"Conflict detected for item_id={item_id} "
                    f"in column '{headers[col_idx]}' "
                    f"with multiple values: {values}"
                )
            merged_row.append(values[0] if values else None)

        ws.append(merged_row)
    return ws


def style(ws):
    # center, bold
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=3):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = Font(bold=True)


def merge_blocks(ws):
    # MERGE x-y-z
    # --- Locate item_id column ---
    headers = [cell.value for cell in ws[1]]
    item_col_index = headers.index("item_id") + 1
    # --- Add new headers for x, y, z ---
    x_col = 1
    y_col = 2
    z_col = 3
    ws.cell(row=1, column=1, value="Bloc")
    ws.cell(row=1, column=2, value="Jardin")
    ws.cell(row=1, column=3, value="Planche")
    # --- Split x-y-z into separate columns ---
    for row in range(2, ws.max_row + 1):
        item_value = ws.cell(row=row, column=item_col_index).value

        if item_value:
            try:
                x, y, z = item_value.split("-")
            except ValueError:
                raise ValueError(f"Invalid format at row {row}: {item_value}")

            ws.cell(row=row, column=x_col, value=x)
            ws.cell(row=row, column=y_col, value=y)
            ws.cell(row=row, column=z_col, value=z)

    background(ws)

    # --- Merge consecutive Y cells (if same Y and same X) ---
    start_row = 2
    while start_row <= ws.max_row:
        current_x = ws.cell(row=start_row, column=x_col).value
        current_y = ws.cell(row=start_row, column=y_col).value

        end_row = start_row

        while (
                end_row + 1 <= ws.max_row and
                ws.cell(row=end_row + 1, column=x_col).value == current_x and
                ws.cell(row=end_row + 1, column=y_col).value == current_y
        ):
            end_row += 1

        if end_row > start_row:
            ws.merge_cells(
                start_row=start_row,
                start_column=y_col,
                end_row=end_row,
                end_column=y_col
            )

        start_row = end_row + 1
    # --- Merge consecutive X cells (if same X) ---
    start_row = 2
    while start_row <= ws.max_row:
        current_x = ws.cell(row=start_row, column=x_col).value
        end_row = start_row

        while (
                end_row + 1 <= ws.max_row and
                ws.cell(row=end_row + 1, column=x_col).value == current_x
        ):
            end_row += 1

        if end_row > start_row:
            ws.merge_cells(
                start_row=start_row,
                start_column=x_col,
                end_row=end_row,
                end_column=x_col
            )

        start_row = end_row + 1


def background(ws):
    # color light blue
    for row in range(2, ws.max_row + 1):  # ignore header
        first_value = ws.cell(row=row, column=2).value
        for col in range(4, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)

            # Only color if no fill already applied
            if cell.fill.fill_type is None:
                if first_value in ['2', '4']:
                    cell.fill = PatternFill( start_color="dcdcdc", end_color="dcdcdc", fill_type="solid")
                else:
                    cell.fill = PatternFill( start_color="fffaf0", end_color="fffaf0", fill_type="solid")

def coloring(ws):
    light_green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    dark_green = PatternFill(start_color="006100", end_color="006100", fill_type="solid")
    # Replace label
    for row in ws.iter_rows(min_row=2):  # skip header
        start_ini = True
        for cell in row:
            if isinstance(cell.value, str) and cell.value not in ("-", "X") and not start_ini:
                cell.fill = light_green
            elif cell.value == "X":
                cell.fill = light_green
                cell.value = ""
            elif cell.value == "-":
                cell.fill = dark_green
                cell.value = ""
            start_ini = False





if __name__ == "__main__":
    main()
