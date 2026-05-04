# plants.py
from __future__ import annotations

from typing import Dict, List
from engine import Job, PlantSpec

# -------------------------
# MONTERREY (7773 — Gamesa)
# -------------------------
MONTERREY = PlantSpec(
    key="MONTERREY",
    jobs=[
        # Cookies or Biscuits — only Merengue(0) and Rheon(7) carry a type_value
        Job("Cookies or Biscuits", "DEFAULT", 0, "11 - Merengue",       101, "PLANTA MONTERREY", "11", type_value="Merengue"),
        Job("Cookies or Biscuits", "DEFAULT", 1, "2 - Marias",          102, "PLANTA MONTERREY", "2"),
        Job("Cookies or Biscuits", "DEFAULT", 2, "3 - Alambre 3",       103, "PLANTA MONTERREY", "3"),
        Job("Cookies or Biscuits", "DEFAULT", 3, "17 - Marias Doradas", 104, "PLANTA MONTERREY", "4"),
        Job("Cookies or Biscuits", "DEFAULT", 4, "7 - Alambre 5",       105, "PLANTA MONTERREY", "5"),
        Job("Cookies or Biscuits", "DEFAULT", 5, "9 - Saladas",         106, "PLANTA MONTERREY", "6"),
        Job("Cookies or Biscuits", "DEFAULT", 6, "10 - Sandwiches",     107, "PLANTA MONTERREY", "7"),
        Job("Cookies or Biscuits", "DEFAULT", 7, "16 - Rheon",          109, "PLANTA MONTERREY", "8",  type_value="Rheon"),

        # Wafer — VBA does not write row-7 type for Monterrey Wafer
        Job("Wafer", "DEFAULT", 0, "13 - Obleas 1", 108, "PLANTA MONTERREY", "W1"),

        # Other — M3 has no type_value in the completed template
        Job("Other", "DEFAULT", 0, "18 - M1", 163, "PLANTA MONTERREY", "M1", type_value="Other"),
        Job("Other", "DEFAULT", 1, "19 - M3", 165, "PLANTA MONTERREY", "M3"),
    ]
)

# -------------------------
# VALLEJO-SABRITAS (7704)
# -------------------------
VALLEJO_SABRITAS = PlantSpec(
    key="VALLEJO_SABRITAS",
    jobs=[
        Job("PC", "PC", 0, "1 - PC 50", 3, "PLANTA MÉXICO", "A3"),

        Job("TC", "TC", 0, "3 - DT1",     5, "PLANTA MÉXICO", "B1", type_value="TC1.5"),
        Job("TC", "TC", 1, "2 - DT 2",    6, "PLANTA MÉXICO", "B2", type_value="TC1.5"),
        Job("TC", "TC", 2, "4 - DT 3000", 7, "PLANTA MÉXICO", "B3", type_value="TC2.0"),

        Job("Other", "DEFAULT", 0, "16 - Kettle",      4,  "PLANTA MÉXICO", "A2", type_value="Others"),
        Job("Other", "DEFAULT", 1, "7 - Fritos 3",     8,  "PLANTA MÉXICO", "C3", type_value="Others"),
        Job("Other", "DEFAULT", 2, "18 - CU3000",      12, "PLANTA MÉXICO", "F2", type_value="Others"),
        Job("Other", "DEFAULT", 3, "19 - Sabriton 2",  13, "PLANTA MÉXICO", "G2", type_value="Others"),
        Job("Other", "DEFAULT", 4, "20 - Sabriton 3",  14, "PLANTA MÉXICO", "G3", type_value="Others"),
        Job("Other", "DEFAULT", 5, "13 - Paloma",      15, "PLANTA MÉXICO", "M1", type_value="Others"),

        Job("Extruded", "DEFAULT", 0, "8 - Cheetos 1",   9,  "PLANTA MÉXICO", "D1", type_value="Fried"),
        Job("Extruded", "DEFAULT", 1, "9 - Cheetos 3",   10, "PLANTA MÉXICO", "D2", type_value="Fried"),
        Job("Extruded", "DEFAULT", 2, "5 - QUESABRITAS", 11, "PLANTA MÉXICO", "E1", type_value="Baked"),
    ]
)

# -------------------------
# GUADALAJARA (7702)
# -------------------------
GUADALAJARA = PlantSpec(
    key="GUADALAJARA",
    jobs=[
        Job("PC", "PC", 0, "1 - PC 32", 18, "PLANTA GUADAL", "A1"),

        Job("TC", "TC", 0, "2 - DT 2000", 19, "PLANTA GUADAL", "B3", type_value="TC 1.5"),
        Job("TC", "TC", 1, "3 - DT 3000", 20, "PLANTA GUADAL", "B4", type_value="TC 2.0"),
        Job("TC", "TC", 2, "10 - DT3002", 21, "PLANTA GUADAL", "B5", type_value="TC 2.0"),

        Job("Other", "DEFAULT", 0, "6 - Fritos 1",   22, "PLANTA GUADAL", "C1", type_value="Other"),
        Job("Other", "DEFAULT", 1, "8 - Churrumais", 25, "PLANTA GUADAL", "F1", type_value="Other"),
        Job("Other", "DEFAULT", 2, "5 - Sabriton 2", 26, "PLANTA GUADAL", "G1", type_value="Other"),

        Job("Extruded", "DEFAULT", 0, "7 - Cheetos F", 23, "PLANTA GUADAL", "D1", type_value="Fried"),
        Job("Extruded", "DEFAULT", 1, "4 - CHEETOS H", 24, "PLANTA GUADAL", "E1", type_value="Baked"),
    ]
)

# -------------------------
# VERACRUZ (7708)
# -------------------------
VERACRUZ = PlantSpec(
    key="VERACRUZ",
    jobs=[
        Job("PC", "PC", 0, "1 - PC 50-1", 29, "PLANTA VERACRUZ", "A1"),

        Job("TC", "TC", 0, "4 - DT 3000", 32, "PLANTA VERACRUZ", "B1", type_value="TC 2.0"),
        Job("TC", "TC", 1, "5 - DT 6000", 33, "PLANTA VERACRUZ", "B2", type_value="TC 2.0"),

        Job("Other", "DEFAULT", 0, "3 - CKF 1",   31, "PLANTA VERACRUZ", "A3", type_value="Others"),
        Job("Other", "DEFAULT", 1, "6 - Fritos 2", 34, "PLANTA VERACRUZ", "C1", type_value="Others"),
        Job("Other", "DEFAULT", 2, "10 - Pellet",  37, "PLANTA VERACRUZ", "G1", type_value="Others"),

        Job("Extruded", "DEFAULT", 0, "9 - Cheetos Fritos 2", 30, "PLANTA VERACRUZ", "D2", type_value="Fried"),
        Job("Extruded", "DEFAULT", 1, "7 - Cheetos F",        35, "PLANTA VERACRUZ", "D1", type_value="Fried"),
        Job("Extruded", "DEFAULT", 2, "8 - Cheetos H",        36, "PLANTA VERACRUZ", "E1", type_value="Baked"),
    ]
)

# -------------------------
# SALTILLO (7706)
# -------------------------
SALTILLO = PlantSpec(
    key="SALTILLO",
    jobs=[
        Job("PC", "PC", 0, "1 - PC 50", 39, "PLANTA SALTILLO", "A3"),

        Job("TC", "TC", 0, "3 - DT 3000",      41, "PLANTA SALTILLO", "B1", type_value="TC 2.0"),
        Job("TC", "TC", 1, "4 - DT 3001",      42, "PLANTA SALTILLO", "B2", type_value="TC 2.0"),
        Job("TC", "TC", 2, "5 - DT 3002",      43, "PLANTA SALTILLO", "B3", type_value="TC 2.0"),
        Job("TC", "TC", 3, "15 - Tortilla B4", 50, "PLANTA SALTILLO", "B4", type_value="TC 1.5"),

        Job("Other", "DEFAULT", 0, "2 - CKF 1",    40, "PLANTA SALTILLO", "A4", type_value="Other"),
        Job("Other", "DEFAULT", 1, "8 - Fritos 1",  44, "PLANTA SALTILLO", "C1", type_value="Other"),
        Job("Other", "DEFAULT", 2, "9 - Churro 1",  47, "PLANTA SALTILLO", "F1", type_value="Other"),
        Job("Other", "DEFAULT", 3, "10 - Paloma",   48, "PLANTA SALTILLO", "M1", type_value="Other"),
        Job("Other", "DEFAULT", 4, "14 - Kettle",   49, "PLANTA SALTILLO", "A5", type_value="Other"),

        Job("Extruded", "DEFAULT", 0, "7 - Cheetos",     45, "PLANTA SALTILLO", "D1", type_value="Fried"),
        Job("Extruded", "DEFAULT", 1, "6 - QUESABRITAS", 46, "PLANTA SALTILLO", "E1", type_value="Baked"),
    ]
)

# -------------------------
# OBREGÓN SABRITAS (7701)
# -------------------------
OBREGON_SABRITAS = PlantSpec(
    key="OBREGON_SABRITAS",
    jobs=[
        Job("PC", "PC", 0, "1 - PC 42", 52, "PLANTA OBREGÓN", "A1"),

        Job("TC", "TC", 0, "2 - DT 3001", 53, "PLANTA OBREGÓN", "B1", type_value="TC 2.0"),
        Job("TC", "TC", 1, "3 - DT 3002", 54, "PLANTA OBREGÓN", "B2", type_value="TC 1.5"),

        Job("Other", "DEFAULT", 0, "4 - SWL",       55, "PLANTA OBREGÓN", "B3/C1", type_value="Other"),
        Job("Other", "DEFAULT", 1, "7 - Churromais", 58, "PLANTA OBREGÓN", "F1",    type_value="Other"),

        Job("Extruded", "DEFAULT", 0, "6 - Cheetos F", 56, "PLANTA OBREGÓN", "D1", type_value="Fried"),
        Job("Extruded", "DEFAULT", 1, "5 - CHEETOS H", 57, "PLANTA OBREGÓN", "E1", type_value="Baked"),
    ]
)

# -------------------------
# MEXICALI I (7705)
# -------------------------
MEXICALI_I = PlantSpec(
    key="MEXICALI_I",
    jobs=[
        Job("PC", "PC", 0, "1 - PC 21", 60, "PLANTA MEXICALI I", "A1"),

        Job("TC", "TC", 0, "2 - DT 3000", 61, "PLANTA MEXICALI I", "B1", type_value="TC 2.0"),

        Job("Extruded", "DEFAULT", 0, "3 - Cheetos F", 62, "PLANTA MEXICALI I", "D1", type_value="Fried"),
    ]
)

# -------------------------
# CELAYA SABRITAS (7787)
# -------------------------
CELAYA_SABRITAS = PlantSpec(
    key="CELAYA_SABRITAS",
    jobs=[
        Job("PC", "PC", 0, "1 - PC-50xU", 64, "Celaya Salado", "A3"),

        Job("TC", "TC", 0, "2 - TC 4.0", 65, "Celaya Salado", "B1", type_value="TC 4.0"),

        Job("Extruded", "DEFAULT", 0, "3 - CH Fritos", 66, "Celaya Salado", "D1", type_value="Fried"),
    ]
)

# -------------------------
# MÉRIDA (7700 Sabritas + 7772 Gamesa) — same template file
# -------------------------
MERIDA = PlantSpec(
    key="MERIDA",
    jobs=[
        # Sabritas baked line (7700) — Pond: PLANTA YUCATÁN
        Job("Extruded", "DEFAULT", 0, "5 - BAKED CHEETOS", 71, "PLANTA YUCATÁN", "E1", type_value="Baked"),

        # Gamesa cookie lines (7772)
        Job("Cookies or Biscuits", "DEFAULT", 0, "1 - Marias", 122, "PLANTA MERIDA", "1", type_value="Maria"),
        Job("Cookies or Biscuits", "DEFAULT", 1, "2 - OtrosG", 123, "PLANTA MERIDA", "2", type_value="OtrosG"),
    ]
)

# -------------------------
# SABRITAS ABC (7753)
# -------------------------
SABRITAS_ABC = PlantSpec(
    key="SABRITAS_ABC",
    jobs=[
        Job("Extruded",            "DEFAULT", 0, "1 - Colmillo",    73,  "Planta ABC", "E2", type_value="Baked"),
        Job("Cookies or Biscuits", "DEFAULT", 0, "3 - Minimarias",  125, "Planta ABC", "E1", type_value="Minimarias"),
        Job("Other",               "DEFAULT", 0, "6 - N1",          141, "Planta ABC", "N1", type_value="Other"),
    ]
)

# -------------------------
# SABRITAS PTE 128 / MEZCLADOTECNIA (7709)
# -------------------------
SABRITAS_PTE_128 = PlantSpec(
    key="SABRITAS_PTE_128",
    jobs=[
        Job("Other", "DEFAULT", 0, "1 - L1", 143, "PLANTA MEZCLADOTECNIA", "L1", type_value="Other"),
        Job("Other", "DEFAULT", 1, "2 - L2", 144, "PLANTA MEZCLADOTECNIA", "L2"),
        Job("Other", "DEFAULT", 2, "3 - L3", 145, "PLANTA MEZCLADOTECNIA", "L3"),
    ]
)

# -------------------------
# TOLUCA (7707 Cacahuate + 8028 Dulce) — same template file
# -------------------------
TOLUCA = PlantSpec(
    key="TOLUCA",
    jobs=[
        # Dulce lines (8028)
        Job("Other", "DEFAULT", 0, "1 - DULCE - L", 138, "Planta toluca",           "I1", type_value="Other"),
        Job("Other", "DEFAULT", 1, "2 - DULCE - J", 139, "Planta toluca",           "J1", type_value="Other"),

        # Cacahuate lines (7707)
        Job("Other", "DEFAULT", 2, "5 - JT", 135, "Planta Toluca Cacahuate", "JT", type_value="Other"),
        Job("Other", "DEFAULT", 3, "6 - TT", 136, "Planta Toluca Cacahuate", "TT", type_value="Other"),
        Job("Other", "DEFAULT", 4, "3 - F1", 133, "Planta Toluca Cacahuate", "F1", type_value="Other"),
        Job("Other", "DEFAULT", 5, "4 - F2", 134, "Planta Toluca Cacahuate", "F2", type_value="Other"),
    ]
)

# -------------------------
# CELAYA GAMESA (7770)
# -------------------------
CELAYA_GAMESA = PlantSpec(
    key="CELAYA_GAMESA",
    jobs=[
        # Cookies or Biscuits — 12 lines (1-12; line 13 not in Jan template)
        Job("Cookies or Biscuits", "DEFAULT",  0, "1 - Alambre 1",   75, "PLANTA CELAYA", "1"),
        Job("Cookies or Biscuits", "DEFAULT",  1, "3 - Saladas 2",   76, "PLANTA CELAYA", "2",  type_value="Saladas 2"),
        Job("Cookies or Biscuits", "DEFAULT",  2, "4 - Marias",      77, "PLANTA CELAYA", "3",  type_value="Marias"),
        Job("Cookies or Biscuits", "DEFAULT",  3, "5 - Sandwiches",  78, "PLANTA CELAYA", "4",  type_value="Sandwiches"),
        Job("Cookies or Biscuits", "DEFAULT",  4, "6 - Florentinas", 79, "PLANTA CELAYA", "5",  type_value="Florentinas"),
        Job("Cookies or Biscuits", "DEFAULT",  5, "8 - Alambre 6",   80, "PLANTA CELAYA", "6",  type_value="Alambre 6"),
        Job("Cookies or Biscuits", "DEFAULT",  6, "9 - Saladas 7",   81, "PLANTA CELAYA", "7",  type_value="Saladas 7"),
        Job("Cookies or Biscuits", "DEFAULT",  7, "12 - OtrosG9",    82, "PLANTA CELAYA", "8",  type_value="OtrosG9"),
        Job("Cookies or Biscuits", "DEFAULT",  8, "13 - Crackets",   83, "PLANTA CELAYA", "9",  type_value="Crackets"),
        Job("Cookies or Biscuits", "DEFAULT",  9, "14 - Alambre 11", 84, "PLANTA CELAYA", "10", type_value="Alambre 11"),
        Job("Cookies or Biscuits", "DEFAULT", 10, "15 - Chocolate",  85, "PLANTA CELAYA", "11", type_value="Chocolate"),
        Job("Cookies or Biscuits", "DEFAULT", 11, "16 - OtrosG2",    86, "PLANTA CELAYA", "12", type_value="OtrosG2"),

        # Wafer — Gamesa Wafer lines carry type_value
        Job("Wafer", "DEFAULT", 0, "17 - Obleas 3", 88, "PLANTA CELAYA", "W3", type_value="Obleas 3"),
        Job("Wafer", "DEFAULT", 1, "18 - Obleas 4", 89, "PLANTA CELAYA", "W4", type_value="Obleas 4"),
    ]
)

# -------------------------
# VALLEJO GAMESA (7776)
# -------------------------
VALLEJO_GAMESA = PlantSpec(
    key="VALLEJO_GAMESA",
    jobs=[
        Job("Cookies or Biscuits", "DEFAULT", 0, "1 - Crackets",       91, "PLANTA VALLEJO", "1",  type_value="Crackets"),
        Job("Cookies or Biscuits", "DEFAULT", 1, "2 - Alambre",        92, "PLANTA VALLEJO", "2",  type_value="Alambre"),
        Job("Cookies or Biscuits", "DEFAULT", 2, "3 - Sandwiches 3",   93, "PLANTA VALLEJO", "3",  type_value="Sandwiches 3"),
        Job("Cookies or Biscuits", "DEFAULT", 3, "4 - OtrosG",         94, "PLANTA VALLEJO", "42", type_value="Merengue"),
        Job("Cookies or Biscuits", "DEFAULT", 4, "5 - Marias",         95, "PLANTA VALLEJO", "5",  type_value="Marias"),
        Job("Cookies or Biscuits", "DEFAULT", 5, "6 - Sandwiches 6",   96, "PLANTA VALLEJO", "6",  type_value="Sandwiches 6"),
        Job("Cookies or Biscuits", "DEFAULT", 6, "7 - Chocolate",      97, "PLANTA VALLEJO", "7",  type_value="Chocolate"),

        Job("Wafer", "DEFAULT", 0, "8 - Obleas 1", 98, "PLANTA VALLEJO", "W1", type_value="Obleas 1"),
        Job("Wafer", "DEFAULT", 1, "9 - Obleas 2", 99, "PLANTA VALLEJO", "W2", type_value="Obleas 2"),
    ]
)

# -------------------------
# OBREGÓN GAMESA (7774)
# -------------------------
OBREGON_GAMESA = PlantSpec(
    key="OBREGON_GAMESA",
    jobs=[
        Job("Cookies or Biscuits", "DEFAULT", 0, "1 - Sandwiches",       112, "PLANTA OBREGÓN", "1", type_value="Sandwiches"),
        Job("Cookies or Biscuits", "DEFAULT", 1, "3 - Crackets",         113, "PLANTA OBREGÓN", "2", type_value="Crackets"),
        Job("Cookies or Biscuits", "DEFAULT", 2, "5 - Rheon",            114, "PLANTA OBREGÓN", "3", type_value="Rheon"),
        Job("Cookies or Biscuits", "DEFAULT", 3, "6 - Saladas",          115, "PLANTA OBREGÓN", "4", type_value="Saladas"),
        Job("Cookies or Biscuits", "DEFAULT", 4, "7 - Alambre",          116, "PLANTA OBREGÓN", "5", type_value="Alambre"),
        Job("Cookies or Biscuits", "DEFAULT", 5, "8 - Marias",           117, "PLANTA OBREGÓN", "6", type_value="Marias"),
        Job("Cookies or Biscuits", "DEFAULT", 6, "11 - Linea 7 Alambre", 118, "PLANTA OBREGÓN", "7", type_value="Linea 7 Alambre"),

        Job("Wafer", "DEFAULT", 0, "10 - Obleas", 119, "PLANTA OBREGÓN", "W2", type_value="Obleas"),
    ]
)


# =============================================================================
# Filename → plant key matching (for "Run All" batch mode)
# =============================================================================
import os as _os
import unicodedata as _ud

# Ordered list: more-specific keywords must come before single-word fallbacks
# so "CelayaGamesa" is matched before the bare "Celaya" entry.
_FNAME_KEYWORDS: List[tuple] = [
    # Two-word/compound names (Sabritas vs Gamesa disambiguation)
    ("celaya_gamesa",    "CELAYA_GAMESA"),
    ("celayagamesa",     "CELAYA_GAMESA"),
    ("celaya_sabritas",  "CELAYA_SABRITAS"),
    ("celayasabritas",   "CELAYA_SABRITAS"),
    ("vallejo_gamesa",   "VALLEJO_GAMESA"),
    ("vallejogamesa",    "VALLEJO_GAMESA"),
    ("vallejo_sabritas", "VALLEJO_SABRITAS"),
    ("vallejosabritas",  "VALLEJO_SABRITAS"),
    ("obregon_gamesa",   "OBREGON_GAMESA"),
    ("obregongamesa",    "OBREGON_GAMESA"),
    ("obregon_sabritas", "OBREGON_SABRITAS"),
    ("obregonsabritas",  "OBREGON_SABRITAS"),
    ("sabritas_abc",     "SABRITAS_ABC"),
    ("sabritasabc",      "SABRITAS_ABC"),
    ("pte_128",          "SABRITAS_PTE_128"),
    ("pte128",           "SABRITAS_PTE_128"),
    ("mezcladotecnia",   "SABRITAS_PTE_128"),
    # Unambiguous single-word names
    ("monterrey",        "MONTERREY"),
    ("guadalajara",      "GUADALAJARA"),
    ("veracruz",         "VERACRUZ"),
    ("saltillo",         "SALTILLO"),
    ("mexicali",         "MEXICALI_I"),
    ("merida",           "MERIDA"),
    ("yucatan",          "MERIDA"),
    ("toluca",           "TOLUCA"),
    ("abc",              "SABRITAS_ABC"),
    # Single-word fallbacks for ambiguous plant names
    # (used only when no compound keyword matched above)
    ("celaya",           "CELAYA_GAMESA"),
    ("vallejo",          "VALLEJO_GAMESA"),
    ("obregon",          "OBREGON_GAMESA"),
]


def _norm_fname(s: str) -> str:
    """Strip diacritics and lowercase a filename string for matching."""
    s = "".join(ch for ch in _ud.normalize("NFKD", s) if not _ud.combining(ch))
    return s.lower()


def match_filename(fname: str) -> "str | None":
    """
    Return the REGISTRY key for a template filename, or None if unrecognised.
    Strips diacritics and lowercases before matching so accents don't matter.
    """
    n = _norm_fname(_os.path.basename(fname))
    for kw, key in _FNAME_KEYWORDS:
        if kw in n:
            return key
    return None


def find_templates(folder: str) -> "list[tuple[str, str]]":
    """
    Scan *folder* for TE templates.
    Any .xlsm file whose name matches a known plant keyword is included.
    Returns a list of (plant_key, absolute_filepath) pairs sorted by filename.
    """
    results = []
    try:
        for fname in sorted(_os.listdir(folder)):
            if not fname.lower().endswith(".xlsm"):
                continue
            fpath = _os.path.join(folder, fname)
            key = match_filename(fname)
            if key:
                results.append((key, fpath))
    except OSError:
        pass
    return results


# -------------------------
# Registry
# -------------------------
REGISTRY: Dict[str, PlantSpec] = {
    MONTERREY.key:        MONTERREY,
    VALLEJO_SABRITAS.key: VALLEJO_SABRITAS,
    GUADALAJARA.key:      GUADALAJARA,
    VERACRUZ.key:         VERACRUZ,
    SALTILLO.key:         SALTILLO,
    OBREGON_SABRITAS.key: OBREGON_SABRITAS,
    MEXICALI_I.key:       MEXICALI_I,
    CELAYA_SABRITAS.key:  CELAYA_SABRITAS,
    MERIDA.key:           MERIDA,
    SABRITAS_ABC.key:     SABRITAS_ABC,
    SABRITAS_PTE_128.key: SABRITAS_PTE_128,
    TOLUCA.key:           TOLUCA,
    CELAYA_GAMESA.key:    CELAYA_GAMESA,
    VALLEJO_GAMESA.key:   VALLEJO_GAMESA,
    OBREGON_GAMESA.key:   OBREGON_GAMESA,
}
