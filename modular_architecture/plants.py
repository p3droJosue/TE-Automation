# plants.py
from __future__ import annotations

from typing import Dict, List
from engine import Job, PlantSpec

# -------------------------
# Monterrey (same as before)
# -------------------------
# plants.py (MONTERREY block)

MONTERREY = PlantSpec(
    key="MONTERREY",
    jobs=[
        # Cookies or Biscuits (8 blocks)
        Job("Cookies or Biscuits", "DEFAULT", 0, "11 - Merengue",       101, "PLANTA MONTERREY", "11", type_value="Merengue"),
        Job("Cookies or Biscuits", "DEFAULT", 1, "2 - Marias",          102, "PLANTA MONTERREY", "2",  type_value="Marias"),
        Job("Cookies or Biscuits", "DEFAULT", 2, "3 - Alambre 3",       103, "PLANTA MONTERREY", "3",  type_value="Alambre 3"),
        Job("Cookies or Biscuits", "DEFAULT", 3, "17 - Marias Doradas", 104, "PLANTA MONTERREY", "4",  type_value="Marias Doradas"),
        Job("Cookies or Biscuits", "DEFAULT", 4, "7 - Alambre 5",       105, "PLANTA MONTERREY", "5",  type_value="Alambre 5"),
        Job("Cookies or Biscuits", "DEFAULT", 5, "9 - Saladas",         106, "PLANTA MONTERREY", "6",  type_value="Saladas"),
        Job("Cookies or Biscuits", "DEFAULT", 6, "10 - Sandwiches",     107, "PLANTA MONTERREY", "7",  type_value="Sandwiches"),
        Job("Cookies or Biscuits", "DEFAULT", 7, "16 - Rheon",          109, "PLANTA MONTERREY", "8",  type_value="Rheon"),

        # Wafer
        Job("Wafer", "DEFAULT", 0, "13 - Obleas 1", 108, "PLANTA MONTERREY", "W1", type_value="Obleas 1"),

        # Other
        Job("Other", "DEFAULT", 0, "18 - M1", 163, "PLANTA MONTERREY", "M1", type_value="Other"),
        Job("Other", "DEFAULT", 1, "19 - M3", 165, "PLANTA MONTERREY", "M3", type_value="Other"),
    ]
)

# ------------------------------------
# Vallejo - Sabritas (your next site)
# DB: 11_TE_NOV.xlsm
# Uses the same line codes as your VBA VallejoSalado macro.
# ------------------------------------
VALLEJO_SABRITAS = PlantSpec(
    key="VALLEJO_SABRITAS",
    jobs=[
        # PC line (A3)
        Job("PC", "PC", 0, "1 - PC 50", 3, "PLANTA MÉXICO", "A3"),

        # TC lines (B1,B2,B3)
        Job("TC", "TC", 0, "3 - DT1",    5, "PLANTA MÉXICO", "B1"),
        Job("TC", "TC", 1, "2 - DT 2",   6, "PLANTA MÉXICO", "B2"),
        Job("TC", "TC", 2, "4 - DT 3000",7, "PLANTA MÉXICO", "B3"),

        # Other lines (A2,C3,F2,G2,G3,M1) — type is constant "Others" in your completed file
        Job("Other", "DEFAULT", 0, "16 - Kettle",      4,  "PLANTA MÉXICO", "A2", type_value="Others"),
        Job("Other", "DEFAULT", 1, "7 - Fritos 3",     8,  "PLANTA MÉXICO", "C3", type_value="Others"),
        Job("Other", "DEFAULT", 2, "18 - CU3000",      12, "PLANTA MÉXICO", "F2", type_value="Others"),
        Job("Other", "DEFAULT", 3, "19 - Sabriton 2",  13, "PLANTA MÉXICO", "G2", type_value="Others"),
        Job("Other", "DEFAULT", 4, "20 - Sabriton 3",  14, "PLANTA MÉXICO", "G3", type_value="Others"),
        Job("Other", "DEFAULT", 5, "13 - Paloma",      15, "PLANTA MÉXICO", "M1", type_value="Others"),

        # Extruded lines (D1,D2,E1) — Fried/Fried/Baked
        Job("Extruded", "DEFAULT", 0, "8 - Cheetos 1",   9,  "PLANTA MÉXICO", "D1", type_value="Fried"),
        Job("Extruded", "DEFAULT", 1, "9 - Cheetos 3",   10, "PLANTA MÉXICO", "D2", type_value="Fried"),
        Job("Extruded", "DEFAULT", 2, "5 - QUESABRITAS", 11, "PLANTA MÉXICO", "E1", type_value="Baked"),
    ]
)

REGISTRY: Dict[str, PlantSpec] = {
    MONTERREY.key: MONTERREY,
    VALLEJO_SABRITAS.key: VALLEJO_SABRITAS,
}