from __future__ import annotations

import os
import shutil
import tempfile
import unicodedata
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import xlwings as xw

# =========================
# Helpers
# =========================
def _norm(s: Any) -> str:
    if s is None:
        return ""
    s = str(s).strip()
    s = "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    return s.upper().strip()

def _norm_line(v: Any) -> str:
    if v is None:
        return ""

    # Excel often returns numbers as float (e.g., 11.0)
    if isinstance(v, float):
        if v.is_integer():
            return str(int(v))
        return str(v).strip()

    if isinstance(v, int):
        return str(v)

    s = str(v).strip()
    # "11.0" -> "11"
    if s.endswith(".0") and s[:-2].isdigit():
        return s[:-2]
    return s

def _safe_float(v: Any) -> float:
    try:
        if v is None or v == "":
            return 0.0
        return float(v)
    except Exception:
        return 0.0

BLOCK_WIDTH = 8

# =========================
# RDP (Reporte Detalle Plantas) headers
# =========================
MANDATORY_HEADERS = [
    "Total Hours", "Mantenimiento", "Sanidades", "Paros y Arranques",
    "Cambios de producto", "Raw Material", "Problemas de calidad", "Others",
    "Waste", "Paros por SAP", "Reuniones", "Horas excedidas Mtto Prev.",
    "Horas excedidas Sanidades", "Horas excedidas cambios", "Falta Demanda (Hr)", "Flujo de APT"
]

OPTIONAL_HEADERS = [
    "Falta Materiales", "APT Lleno", "Festivos", "Cuadrillas", "Capex / Pruebas",
    "Inventario", "Otros", "Fallas Mecánicas Proceso", "Paros proceso operativos",
    "Fallas Mecánicas empaque", "Paros de empaque Operativo", "Servicios Interno",
    "Servicios Externo", "Pérdida Vel-Restr. eq.", "Pérd Vel-Mix Gram.",
    "Pérd vel- Multipack", "TL - Cambio de loop", "Pérd Vel-Pruebas", "Cocimiento TC",
    "CC Mtto", "CC Sanidad", "CC Changeover", "A005", "A003", "SKUS"
]

# =========================
# Row maps from VBA Write_Parameters
# =========================
ROWS_TEMPLATE_DEFAULT = [18, 22, 23, 24, 25, 35, 36, 39, 14, 26, 27, 30, 31, 32, 19, 32]
ROWS_TEMPLATE_TC      = [19, 23, 24, 25, 26, 36, 37, 40, 15, 27, 28, 31, 32, 33, 20, 33]
ROWS_TEMPLATE_PC      = [28, 32, 33, 34, 35, 45, 46, 49, 24, 36, 37, 40, 41, 42, 29, 42]

ROWS_OPT_COOKIES_WAFER = [
    20, 21, 22, 24, 23, 27, 28, 32, 33, 36, 37, 40, 41,
    122, 123, 124, 125, 126,
    38, 39, 40,
    27, 27, 28, 63
]
ROWS_OPT_OTHER_EXTRUDED = [
    20, 21, 22, 24, 23, 27, 28, 32, 33, 36, 37, 40, 41,
    106, 107, 108, 109, 110, 111,
    38, 39, 40,
    27, 27, 28, 63
]
ROWS_OPT_TC = [
    21, 22, 23, 25, 24, 28, 29, 33, 34, 37, 38, 41, 42,
    123, 124, 125, 126, 127,
    30, 39, 40, 41,
    28, 29, 72
]
ROWS_OPT_PC = [
    30, 31, 32, 34, 33, 37, 38, 42, 43, 46, 47, 50, 51,
    152, 153, 154, 155, 156,
    48, 49, 50,
    37, 37, 38, 91
]

# =========================
# Models
# =========================
@dataclass(frozen=True)
class Job:
    sheet: str                  # template sheet to write
    kind: str                   # "DEFAULT" | "TC" | "PC"
    cons: int                   # block index (0..)
    header_text: str            # line display name
    rdp_row: int                # row in Reporte Detalle Plantas
    pond_plant: str             # Pond col2 plant value (e.g. "PLANTA MÉXICO")
    pond_line: str              # Pond col3 line (e.g. "A3", "B1")
    type_value: Optional[str] = None  # writes into row7 type cell (Other/Extruded)
    # Sabritas Pte 128 ("Mezcladotecnia") has 18–31 condimento SKUs per line,
    # which exceeds the 12-slot product table. Set aggregate_skus=True to
    # collapse all matching SKUs into one row whose velocity is the throughput
    # weighted by production time (= total_prod / sum(prod_i / vel_i)).
    aggregate_skus: bool = False
    aggregate_name: Optional[str] = None  # SKU name to write (e.g., "CONDIMENTO L2")

@dataclass(frozen=True)
class PlantSpec:
    key: str
    jobs: List[Job]

# =========================
# Excel reading
# =========================
def _pick_sheet(wb: xw.Book, preferred: str, fallbacks: List[str]) -> xw.Sheet:
    names = [s.name for s in wb.sheets]
    if preferred in names:
        return wb.sheets[preferred]
    for fb in fallbacks:
        if fb in names:
            return wb.sheets[fb]
    raise ValueError(f"Could not find sheet '{preferred}' (or fallbacks {fallbacks}). Available: {names}")

def _rdp_header_map(ws_rdp: xw.Sheet, header_row: int = 2, max_cols: int = 2000) -> Dict[str, int]:
    last_col = min(ws_rdp.api.UsedRange.Columns.Count, max_cols)
    headers = ws_rdp.range((header_row, 1), (header_row, last_col)).value
    m: Dict[str, int] = {}
    # The RDP has multiple "Waste" columns (one in kg, one as %, sometimes a dashboard total).
    # The percentage Waste is always the one immediately following "ETE  %", so capture it
    # explicitly and override the dict (which otherwise picks the last duplicate).
    ete_idx: Optional[int] = None
    waste_after_ete: Optional[int] = None
    for idx, h in enumerate(headers, start=1):
        if isinstance(h, str) and h.strip():
            key = h.strip()
            m[key] = idx
            collapsed = " ".join(key.split())
            if collapsed == "ETE %":
                ete_idx = idx
            elif key == "Waste" and ete_idx is not None and waste_after_ete is None:
                waste_after_ete = idx
    if waste_after_ete is not None:
        m["Waste"] = waste_after_ete
    return m

def _values_by_headers(ws_rdp: xw.Sheet, row: int, hdr: Dict[str, int], names: List[str]) -> List[Any]:
    out = []
    for k in names:
        c = hdr.get(str(k).strip())
        out.append(ws_rdp.range((row, c)).value if c else None)
    return out

def _pond_last_row(ws_pond: xw.Sheet) -> int:
    # last used row based on column A
    # xlUp = -4162, but use the COM constant directly
    return ws_pond.api.Cells(ws_pond.api.Rows.Count, 1).End(-4162).Row

def _load_pond_df(ws_pond: xw.Sheet, last_col: int = 47) -> pd.DataFrame:
    """
    Load Pond as a DataFrame using ONLY positional columns.
    We read A..AU (1..47) to support PC params (col47) and material rows.
    """
    last_row = _pond_last_row(ws_pond)
    data = ws_pond.range((1, 1), (last_row, last_col)).value
    if not data or len(data) < 2:
        return pd.DataFrame()

    # DO NOT trust excel headers (often duplicated/blank). Use c1..c47
    cols = [f"c{i}" for i in range(1, last_col + 1)]
    df = pd.DataFrame(data[1:], columns=cols)
    return df

# =========================
# Pond extractors
# =========================
def pond_default_names_velprod(df: pd.DataFrame, plant: str, line: str,
                               aggregate: bool = False,
                               agg_name: Optional[str] = None) -> Tuple[List[str], List[float]]:
    """
    DEFAULT & TC use:
      - name: col5 (E) => c5
      - vel : col7 (G) => c7
      - prod: col8 (H) => c8
      Filter: plant (col2=c2), line (col3=c3), prod>0

    If aggregate=True (used for Sabritas Pte 128 / Mezcladotecnia where each
    line has dozens of condimento SKUs), the matching rows are collapsed into
    a single entry: one name, the throughput weighted by production time, and
    the total production.
    """
    if df.empty:
        return [], []

    plant_ser = df["c2"].astype(str).map(_norm)
    line_ser  = df["c3"].map(_norm_line)
    prod_ser  = pd.to_numeric(df["c8"], errors="coerce").fillna(0)

    mask = (plant_ser == _norm(plant)) & (line_ser == _norm_line(line)) & (prod_ser > 0)
    d = df.loc[mask].copy()
    if d.empty:
        return [], []

    names = d["c5"].fillna("").astype(str).tolist()
    vel = pd.to_numeric(d["c7"], errors="coerce").fillna(0).tolist()
    prod = pd.to_numeric(d["c8"], errors="coerce").fillna(0).tolist()

    if aggregate:
        total_prod = sum(prod)
        total_time = sum(p / v for p, v in zip(prod, vel) if v and v > 0)
        if total_time <= 0:
            return [], []
        weighted_vel = total_prod / total_time
        name = agg_name or f"CONDIMENTO {line}"
        return [name], [weighted_vel, total_prod]

    return names, vel + prod

def pond_pc_names_velprod(df: pd.DataFrame, ws_pond: xw.Sheet, template_location: str,
                          plant: str, line: str) -> Tuple[List[str], List[float]]:
    """
    PC needs (Pond material columns):
      - Product Name  : col5  ('Descripción')
      - Product Type  : col21 ('Product Type', e.g. 'Flats', 'Ridged')
      - % Seasoning   : col22 ('% sazonado')
      - % Oil Spray   : col23 ('% oil spray Total ')
      - Production    : col8  ('Actual Prod')
    Plus PC extra params from Pond top-table (rows 2..8). The top-table is keyed by
    Centro Planta (CECO) in col 28 and friendly name in col 29; we look up by CECO
    (which we read from the material rows) so plant-name differences like
    'Celaya' vs 'Celaya Salado' or 'Mexicali 1 PMF' vs 'Mexicali Plant' don't matter.
    The extra params come from:
      wash water temp = col32 ('Agua de lavador (°C)')
      moisture aim    = col33 ('Humedad de producto')
      surface water   = col34 ('Humedad superficial')
      gross solids    = col47 ('Sólidos gross')
    """
    if df.empty:
        return [], []

    # Material rows filter
    plant_ser = df["c2"].astype(str).map(_norm)
    line_ser  = df["c3"].map(_norm_line)
    prod_ser  = pd.to_numeric(df["c8"], errors="coerce").fillna(0)
    mask = (plant_ser == _norm(plant)) & (line_ser == _norm_line(line)) & (prod_ser > 0)
    d = df.loc[mask].copy()
    if d.empty:
        return [], []

    product_names = d["c5"].fillna("").astype(str).tolist()
    product_types = d["c21"].fillna("").astype(str).tolist()  # Product Type ('Flats', etc.)

    seasoning  = pd.to_numeric(d["c22"], errors="coerce").fillna(0).tolist()  # % sazonado
    oil_spray  = pd.to_numeric(d["c23"], errors="coerce").fillna(0).tolist()  # % oil spray Total
    production = pd.to_numeric(d["c8"],  errors="coerce").fillna(0).tolist()

    # CECO from the material rows — disambiguates the top-table row.
    ceco_series = pd.to_numeric(d["c1"], errors="coerce").dropna()
    ceco = int(ceco_series.iloc[0]) if not ceco_series.empty else None

    loc_norm = _norm(template_location)
    found = None
    # Top-table currently runs rows 2..9, but scan a few extra rows in case
    # more PC plants are added; we stop at the first blank CECO row.
    for r in range(2, 16):
        v_ceco = ws_pond.range((r, 28)).value
        v_name = ws_pond.range((r, 29)).value
        if v_ceco is None and v_name is None:
            break
        if ceco is not None and v_ceco is not None:
            try:
                if int(v_ceco) == ceco:
                    found = r
                    break
            except (ValueError, TypeError):
                pass
        if _norm(v_name) == loc_norm:
            found = r
            break
    if found is None:
        raise ValueError(
            f"PC extra params not found in Pond rows 2..8 for location "
            f"'{template_location}' (ceco={ceco})"
        )

    wash_temp = ws_pond.range((found, 32)).value   # col32 (Agua de lavador)
    moisture  = ws_pond.range((found, 33)).value   # col33 (Humedad de producto)
    surface   = ws_pond.range((found, 34)).value   # col34 (Humedad superficial)
    gross_sol = ws_pond.range((found, 47)).value   # col47 (Sólidos gross)

    # names array is 2 halves (name + type)
    names = product_names + product_types

    # vel_prod array: seasoning + oil + production + extras(4)
    vel_prod = seasoning + oil_spray + production + [wash_temp, moisture, surface, gross_sol]
    return names, vel_prod

# =========================
# Writers (match VBA layouts)
# =========================
def _write_header(ws: xw.Sheet, sheet_name: str, cons: int, header_text: str, type_value: Optional[str]):
    # PC uses row 6 headers; others use row 4
    if sheet_name == "PC":
        ws.range((6, 2 + BLOCK_WIDTH * cons)).value = header_text
    else:
        ws.range((4, 2 + BLOCK_WIDTH * cons)).value = header_text

    # type cell row 7, base col 3
    if type_value is not None:
        ws.range((7, 3 + BLOCK_WIDTH * cons)).value = type_value

def _write_mandatory(ws: xw.Sheet, cons: int, rows_template: List[int], vals: List[Any]):
    sumador = BLOCK_WIDTH * cons
    # first 15 into column C, last into column E
    for i in range(len(rows_template) - 1):
        ws.range((rows_template[i], 3 + sumador)).value = vals[i]
    ws.range((rows_template[-1], 5 + sumador)).value = vals[-1]

_UNSET = object()

def _write_optional(ws: xw.Sheet, cons: int, rows_opt: List[int], opt_vals: List[Any],
                    first_to_col7: int, second_from: int, second_to: int, second_col5: bool,
                    last_write_val=_UNSET, last_row_idx: int = 24):
    """
    Generic optional writing used across tabs.
    last_write_val: explicit value for the final col7 write. Defaults to opt[24] (SKUS).
                    Pass 0 for TC (VBA writes Optional_Parameters(25)=0).
    last_row_idx:   index into rows_opt for the final SKUS col7 write.
                    24 for COOKIES/WAFER/TC/PC (25-item arrays), 25 for OTHER/EXTRUDED (26-item array).
    """
    sumador = BLOCK_WIDTH * cons
    opt = list(opt_vals)
    if len(opt) < 25:
        opt.extend([None] * (25 - len(opt)))

    # first chunk to col7
    for i in range(0, first_to_col7 + 1):
        ws.range((rows_opt[i], 7 + sumador)).value = opt[i]

    # second chunk to col5
    if second_col5:
        for i in range(second_from, second_to + 1):
            ws.range((rows_opt[i], 5 + sumador)).value = opt[i]

    # final col7 write (SKUS destination row)
    final_v = opt[24] if last_write_val is _UNSET else last_write_val
    ws.range((rows_opt[last_row_idx], 7 + sumador)).value = final_v

def _write_default_sku(ws: xw.Sheet, sheet_name: str, cons: int, names: List[str], vel_prod: List[float]):
    sumador = BLOCK_WIDTH * cons
    # default layout
    if sheet_name in ("Other", "Extruded"):
        row_1, row_2 = 44, 60
    else:
        row_1, row_2 = 44, 68

    if not names or not vel_prod:
        return

    n = min(len(names), len(vel_prod) // 2)
    if n <= 0:
        return

    ws.range((row_1, 2 + sumador)).options(transpose=True).value = names[:n]           # names
    ws.range((row_1, 7 + sumador)).options(transpose=True).value = vel_prod[:n]        # vel
    ws.range((row_2, 3 + sumador)).options(transpose=True).value = vel_prod[n:n + n]   # prod

def _write_tc_sku(ws: xw.Sheet, cons: int, names: List[str], vel_prod: List[float]):
    """
    TC VBA layout:
      row_1=45, row_2=69, velocity goes to col (6 + sumador),
      production goes to col (3 + sumador), and it sets "Other" at row_1+idx col (3+sumador).
    """
    sumador = BLOCK_WIDTH * cons
    row_1, row_2 = 45, 69

    if not names or not vel_prod:
        return

    n = min(len(names), len(vel_prod) // 2)
    if n <= 0:
        return

    vel = vel_prod[:n]
    prod = vel_prod[n:n + n]

    for idx in range(n):
        ws.range((row_1 + idx, 3 + sumador)).value = "Other"
        ws.range((row_1 + idx, 2 + sumador)).value = names[idx]
        ws.range((row_1 + idx, 6 + sumador)).value = vel[idx]      # TC special column

        ws.range((row_2 + idx, 3 + sumador)).value = prod[idx]

def _write_pc(ws: xw.Sheet, cons: int, mandatory: List[Any], optional: List[Any],
              names: List[str], vel_prod: List[float]):
    """
    PC VBA layout:
      - mandatory rows: ROWS_TEMPLATE_PC
      - product table row_1=54:
          col2 names, col3 types, col4 seasoning, col5 oil
      - production table row_2=88:
          col3 production
      - extra params into C10,C11,C12,C17 for the block
    """
    sumador = BLOCK_WIDTH * cons

    # mandatory
    _write_mandatory(ws, cons, ROWS_TEMPLATE_PC, mandatory)

    if not names or not vel_prod:
        # optional still gets written
        _write_optional(ws, cons, ROWS_OPT_PC, optional, first_to_col7=17, second_from=18, second_to=23, second_col5=True)
        return

    n = len(names) // 2
    if n <= 0:
        _write_optional(ws, cons, ROWS_OPT_PC, optional, first_to_col7=17, second_from=18, second_to=23, second_col5=True)
        return

    # Split arrays
    prod_names = names[:n]
    prod_types = names[n:n + n]

    seasoning = vel_prod[:n]
    oil = vel_prod[n:n + n]
    production = vel_prod[2 * n: 3 * n]

    extras = vel_prod[3 * n: 3 * n + 4]  # [wash, moisture, surface, gross]
    while len(extras) < 4:
        extras.append(None)

    row_1, row_2 = 54, 88

    ws.range((row_1, 2 + sumador)).options(transpose=True).value = prod_names
    ws.range((row_1, 3 + sumador)).options(transpose=True).value = prod_types
    ws.range((row_1, 4 + sumador)).options(transpose=True).value = seasoning
    ws.range((row_1, 5 + sumador)).options(transpose=True).value = oil

    ws.range((row_2, 3 + sumador)).options(transpose=True).value = production

    # Extra operating/process parameters (same cells as VBA)
    ws.range((10, 3 + sumador)).value = extras[0]
    ws.range((11, 3 + sumador)).value = extras[1]
    ws.range((12, 3 + sumador)).value = extras[2]
    ws.range((17, 3 + sumador)).value = extras[3]

    # optional
    _write_optional(ws, cons, ROWS_OPT_PC, optional, first_to_col7=17, second_from=18, second_to=23, second_col5=True)

# =========================
# Stale-data cleanup
# =========================
_PRODUCT_SHEETS = ("PC", "TC", "Extruded", "Other", "Cookies or Biscuits", "Wafer")
_MAX_BLOCKS = 15  # templates have up to 15 (cons 0..14)

def _clear_unused_sheets(wb_tpl: xw.Book, plant: PlantSpec) -> None:
    """
    Run All overwrites templates in place. If a file was previously filled
    using a different PlantSpec (e.g. the bare "Snk_TE_Celaya" file was
    matched to CELAYA_GAMESA by an earlier code version and got Cookies/Wafer
    data, then is re-run as CELAYA_SABRITAS), the sheets the new spec doesn't
    touch keep the old data and contaminate the Summary.

    For each product sheet not used by *plant*, zero out the input cells the
    engine writes (SKU table + Total Hours per block). The template's formulas
    cascade those to 0, so the Summary reflects only the new spec's contribution.
    """
    used_sheets = {job.sheet for job in plant.jobs}
    sheets_in_wb = {s.name for s in wb_tpl.sheets}
    for sheet_name in _PRODUCT_SHEETS:
        if sheet_name in used_sheets or sheet_name not in sheets_in_wb:
            continue
        ws = wb_tpl.sheets[sheet_name]
        # Drop sheet protection so we can wipe locked input cells, then restore.
        # Pass "" so Excel doesn't pop the password dialog: if the sheet has a
        # real password, Unprotect throws and we skip — locked cells can't have
        # accumulated stale data from a previous fill anyway.
        was_protected = bool(ws.api.ProtectContents)
        if was_protected:
            try:
                ws.api.Unprotect("")
            except Exception:
                continue
        try:
            for cons in range(_MAX_BLOCKS):
                sumador = BLOCK_WIDTH * cons
                ws.range((18, 3 + sumador)).value = None  # Total Hours
                ws.range((44, 2 + sumador), (87, 2 + sumador)).value = None  # SKU names
                ws.range((44, 7 + sumador), (87, 7 + sumador)).value = None  # velocities
                ws.range((60, 3 + sumador), (87, 3 + sumador)).value = None  # productions
        finally:
            if was_protected:
                try:
                    ws.api.Protect()
                except Exception:
                    pass

# =========================
# Engine runner
# =========================
def default_out_path(blank_path: str) -> str:
    folder = os.path.dirname(blank_path)
    base = os.path.basename(blank_path)
    lower = base.lower()
    if "_blank" in lower:
        idx = lower.index("_blank")
        base = base[:idx] + "_Filled" + base[idx + 6:]
    else:
        root, ext = os.path.splitext(base)
        base = f"{root}_Filled{ext}"
    return os.path.join(folder, base)

def _is_onedrive_path(path: str) -> bool:
    """OneDrive's sync agent acquires file locks while a workbook is open in
    Excel — when an automated save fires those operations COM-throw
    "Cannot access '<file>'". Detect any OneDrive folder, both personal
    ("OneDrive") and tenant-scoped ("OneDrive - <Org>"), by substring.
    """
    return "onedrive" in (path or "").lower()


def fill_plant(db_path: str, blank_path: str, out_path: str, plant: PlantSpec, visible: bool = False) -> None:
    """Run the fill, staging through a local temp folder when any input
    lives on OneDrive. The core engine then operates on plain local files
    that the sync agent can't touch, and we copy the result back to the
    OneDrive location as a single write at the end.

    Excel COM occasionally throws "Cannot access '<file>'" after the
    workbook has already been written to disk (most often when OneDrive's
    sync agent grabs the file the instant save returns). To avoid flagging
    those runs as failed, we snapshot the output file's mtime before
    starting and treat an exception as a success if the file has been
    rewritten in the meantime.
    """
    mtime_before = (
        os.path.getmtime(out_path) if os.path.exists(out_path) else 0.0
    )

    try:
        if (_is_onedrive_path(blank_path)
                or _is_onedrive_path(db_path)
                or _is_onedrive_path(out_path)):
            with tempfile.TemporaryDirectory(prefix="te_filler_") as tmpdir:
                local_db  = os.path.join(tmpdir, "db_"  + os.path.basename(db_path))
                local_tpl = os.path.join(tmpdir, "tpl_" + os.path.basename(blank_path))
                shutil.copy2(db_path, local_db)
                shutil.copy2(blank_path, local_tpl)
                _fill_plant_core(local_db, local_tpl, local_tpl, plant, visible)
                shutil.copy2(local_tpl, out_path)
        else:
            _fill_plant_core(db_path, blank_path, out_path, plant, visible)
    except Exception:
        # If the output file's mtime advanced during this call the bytes
        # are on disk — whatever COM error fired afterwards (OneDrive lock
        # races, Excel close hiccups, etc.) shouldn't mark the run failed.
        try:
            if (os.path.exists(out_path)
                    and os.path.getmtime(out_path) > mtime_before + 0.5):
                return
        except OSError:
            pass
        raise


def _fill_plant_core(db_path: str, blank_path: str, out_path: str, plant: PlantSpec, visible: bool = False) -> None:
    app = xw.App(visible=visible, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        wb_db = app.books.open(db_path, update_links=False, read_only=True)
        wb_tpl = app.books.open(blank_path, update_links=False)

        # Excel calculation is critical
        app.api.CalculateFullRebuild()

        ws_rdp = _pick_sheet(wb_db, "Reporte Detalle Plantas", ["Reporte Detalle Plantas (2)"])
        ws_pond = _pick_sheet(wb_db, "Pond", [])

        hdr = _rdp_header_map(ws_rdp, header_row=2, max_cols=2000)
        # Only mandatory headers gate the run — optional ones (e.g. A005/A003/SKUS)
        # are dropped from some monthly DBs (FEB 2026 lost them) and the downstream
        # writers already treat a missing optional as None (leaves the cell blank).
        missing = [h for h in MANDATORY_HEADERS if str(h).strip() not in hdr]
        if missing:
            raise ValueError(f"RDP headers missing: {missing}")

        df_pond = _load_pond_df(ws_pond, last_col=47)

        # Wipe stale data left by a prior fill that used a different PlantSpec.
        _clear_unused_sheets(wb_tpl, plant)

        # PC template location is needed for PC extra params
        pc_location = None
        if "PC" in [s.name for s in wb_tpl.sheets]:
            pc_location = wb_tpl.sheets["PC"].range("D2").value

        for job in plant.jobs:
            ws = wb_tpl.sheets[job.sheet]

            # 1) header + type
            _write_header(ws, job.sheet, job.cons, job.header_text, job.type_value)

            # 2) RDP mandatory/optional
            mandatory = _values_by_headers(ws_rdp, job.rdp_row, hdr, MANDATORY_HEADERS)
            optional = _values_by_headers(ws_rdp, job.rdp_row, hdr, OPTIONAL_HEADERS)

            # 3) Pond extract + write
            if job.kind.upper() == "PC":
                if pc_location is None:
                    raise ValueError("Template PC location (PC!D2) not found but a PC job exists.")
                names, vel_prod = pond_pc_names_velprod(df_pond, ws_pond, str(pc_location), job.pond_plant, job.pond_line)
                _write_pc(ws, job.cons, mandatory, optional, names, vel_prod)

            elif job.kind.upper() == "TC":
                names, vel_prod = pond_default_names_velprod(
                    df_pond, job.pond_plant, job.pond_line,
                    aggregate=job.aggregate_skus, agg_name=job.aggregate_name,
                )
                _write_mandatory(ws, job.cons, ROWS_TEMPLATE_TC, mandatory)
                _write_optional(ws, job.cons, ROWS_OPT_TC, optional, first_to_col7=18, second_from=19, second_to=23, second_col5=True, last_write_val=0)
                _write_tc_sku(ws, job.cons, names, vel_prod)

            else:  # DEFAULT
                names, vel_prod = pond_default_names_velprod(
                    df_pond, job.pond_plant, job.pond_line,
                    aggregate=job.aggregate_skus, agg_name=job.aggregate_name,
                )
                _write_mandatory(ws, job.cons, ROWS_TEMPLATE_DEFAULT, mandatory)

                if job.sheet in ("Cookies or Biscuits", "Wafer"):
                    # 25-item array: velocity loss rows 13-17, col E second chunk starts at 18
                    _write_optional(ws, job.cons, ROWS_OPT_COOKIES_WAFER, optional,
                                    first_to_col7=17, second_from=18, second_to=24,
                                    second_col5=True, last_row_idx=24)
                else:
                    # 26-item array: velocity loss rows 13-18 (extra row 111), col E starts at 19
                    # final SKUS write goes to rows_opt[25]=63
                    _write_optional(ws, job.cons, ROWS_OPT_OTHER_EXTRUDED, optional,
                                    first_to_col7=18, second_from=19, second_to=24,
                                    second_col5=True, last_row_idx=25)

                _write_default_sku(ws, job.sheet, job.cons, names, vel_prod)

        wb_tpl.app.api.CalculateFull()
        wb_tpl.save(out_path)

        # File is on disk. Close is best-effort: when the workbook lives on
        # OneDrive the sync agent often grabs a lock the instant save finishes,
        # and Excel's Close() then COM-throws "Cannot access '<file>'". Letting
        # that propagate would mark a correctly-saved plant as Failed.
        try:
            wb_tpl.close()
        except Exception:
            pass
        try:
            wb_db.close()
        except Exception:
            pass

    finally:
        # Each call may COM-throw if Excel is in a modal/busy state from an
        # earlier failure. Swallow individually so app.quit() always runs —
        # otherwise the orphaned Excel process locks files for the next plant.
        for cleanup in (
            lambda: setattr(app, "screen_updating", True),
            lambda: setattr(app, "display_alerts", True),
            app.quit,
        ):
            try:
                cleanup()
            except Exception:
                pass