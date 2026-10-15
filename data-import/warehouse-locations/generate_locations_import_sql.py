"""
Genera lo script SQL che associa gli articoli alle ubicazioni di magazzino
(tabella ARMGUbicazione) partendo da un file Excel. Vedi README.md.
"""
import os

import pandas as pd
import pyodbc

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

EXCEL_FILE = "ubicazioni.xlsx"
OUTPUT_SQL  = "importa_ubicazioni.sql"
DATABASE    = os.environ["DB_ARCA"]
CD_MG       = "00002"  # codice del magazzino con le ubicazioni

UBICAZIONE_COLS = [
    "Ubicazione_1", "Ubicazione_2", "Ubicazione_3", "Ubicazione_4",
    "Ubicazione_5", "Ubicazione_6", "Ubicazione_7", "Ubicazione_8"
]

def get_connection_string():
    """Stringa di connessione ODBC ad ARCA, letta dalle variabili d'ambiente (.env)."""
    conn_str = (
        "Driver={SQL Server};"
        f"Server={os.environ['SERVER_DB']};"
        f"Database={os.environ['DB_ARCA']};"
    )
    if os.getenv("DB_USER"):
        return conn_str + f"UID={os.environ['DB_USER']};PWD={os.getenv('DB_PASSWORD', '')};"
    return conn_str + "Trusted_Connection=yes;"


def pad_ubicazione(ub):
    if not ub:
        return ""
    ub = ub.replace(",", ".")
    parti = ub.split(".")
    if len(parti) == 2:
        return parti[0].zfill(3) + "." + parti[1]
    return ub

# Carica ubicazioni valide dal database
conn = pyodbc.connect(get_connection_string())
cursor = conn.cursor()
cursor.execute("SELECT Cd_MGUbicazione FROM [dbo].[MGUbicazione]")
ubicazioni_valide = set(row[0] for row in cursor.fetchall())
conn.close()

df = pd.read_excel(EXCEL_FILE, dtype=str)
df = df.fillna("")

lines = []
lines.append(f"USE [{DATABASE}];\n\n")
lines.append("DISABLE TRIGGER ALL ON [dbo].[ARMGUbicazione];\n\n")

skipped   = 0
processed = 0

for _, row in df.iterrows():
    cd_ar    = str(row["Cd_AR"]).strip()
    ub_pref  = pad_ubicazione(str(row["Ubicazione_Preferita"]).strip())

    if not cd_ar or not ub_pref:
        skipped += 1
        continue

    if ub_pref not in ubicazioni_valide:
        print(f"SKIP ubicazione non valida: Cd_AR='{cd_ar}' -> '{ub_pref}'")
        skipped += 1
        continue

    lines.append(
        f"IF NOT EXISTS (SELECT 1 FROM [dbo].[ARMGUbicazione] WHERE Cd_AR='{cd_ar}' AND Cd_MGUbicazione='{ub_pref}')\n"
        f"    INSERT INTO [dbo].[ARMGUbicazione] (Cd_AR, Cd_MGUbicazione, Cd_MG, DefaultMGUbicazione, xDp_DefaultMGUbicazione)\n"
        f"    VALUES ('{cd_ar}', '{ub_pref}', '{CD_MG}', 1, 1);\n"
    )

    for col in UBICAZIONE_COLS:
        ub = pad_ubicazione(str(row.get(col, "")).strip())
        if not ub or ub == "nan":
            continue
        if ub not in ubicazioni_valide:
            print(f"SKIP ubicazione non valida: Cd_AR='{cd_ar}' | {col} -> '{ub}'")
            continue
        lines.append(
            f"IF NOT EXISTS (SELECT 1 FROM [dbo].[ARMGUbicazione] WHERE Cd_AR='{cd_ar}' AND Cd_MGUbicazione='{ub}')\n"
            f"    INSERT INTO [dbo].[ARMGUbicazione] (Cd_AR, Cd_MGUbicazione, Cd_MG, DefaultMGUbicazione, xDp_DefaultMGUbicazione)\n"
            f"    VALUES ('{cd_ar}', '{ub}', '{CD_MG}', 0, 0);\n"
        )

    processed += 1

lines.append("\nENABLE TRIGGER ALL ON [dbo].[ARMGUbicazione];\n")

with open(OUTPUT_SQL, "w", encoding="utf-8") as f:
    f.writelines(lines)

print(f"\nArticoli processati : {processed}")
print(f"Righe skippate      : {skipped}")
print(f"File SQL generato   : {OUTPUT_SQL}")