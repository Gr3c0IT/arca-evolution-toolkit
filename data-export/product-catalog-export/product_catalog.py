"""
catalogo.py - Generazione e aggiornamento catalogo prodotti NOME_AZIENDA

Utilizzo:
  python catalogo.py --genera <file_codici.xlsx>
      Legge i codici articolo, pesca i dati da ARCA + catalogo_unificato.xlsx
      e produce catalogo_output.xlsx + missing_data.txt

  python catalogo.py --aggiorna <file_collega_compilato.xlsx>
      Aggiorna catalogo_unificato.xlsx con i dati manuali compilati dalla collega
"""

import os
import sys
import re
import warnings
import argparse
from pathlib import Path
from datetime import datetime

import pyodbc
import pandas as pd

warnings.filterwarnings("ignore")

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ---------------------------------------------------------------------------
# Configurazione
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).parent
CATALOGO_UNIFICATO = SCRIPT_DIR / "catalogo_unificato.xlsx"
OUTPUT_FILE = SCRIPT_DIR / "catalogo_output.xlsx"
DA_COMPILARE = SCRIPT_DIR / "da_compilare.xlsx"

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

# Colonne che nel catalogo_unificato vengono da compilazione manuale
COLONNE_MANUALI = [
    "Descrizione",
    "TOP_SELLER",
    "NOVITA",
    "TREND",
    "GLUTEN_FREE",
    "PROTEICO",
    "VEGANO",
    "HALAL",
]

# Colonne booleane: accettano solo TRUE/FALSE oppure VERO/FALSO (case-insensitive)
COLONNE_BOOLEANE = [
    "TOP_SELLER", "NOVITA", "TREND", "GLUTEN_FREE",
    "PROTEICO", "VEGANO", "HALAL",
]

# Ordine e nomi finali del catalogo output
COLONNE_OUTPUT = [
    "cod_art", "NOME_PRODOTTO", "Categoria", "Sottocategoria",
    "Peso", "ean_pz", "ean_cf", "pz_cf",
    "cf_ly", "cf_plt", "ly_plt", "Immagini",
    "Descrizione", "Prezzo_Rivendita_Suggerito",
    "TOP_SELLER", "NOVITA", "TREND",
    "GLUTEN_FREE", "PROTEICO", "VEGANO", "HALAL",
]

# ---------------------------------------------------------------------------
# Connessione DB
# ---------------------------------------------------------------------------
def get_connection():
    return pyodbc.connect(get_connection_string())


# ---------------------------------------------------------------------------
# Funzioni di estrazione dal DB (adattate dallo script originale)
# ---------------------------------------------------------------------------
def get_ar_data(conn, codes):
    query = f"""
    SELECT Cd_AR, Descrizione, PesoNetto, PesoFattore,
           Cd_Aliquota_V, ExtraInfo, Altezza, Lunghezza, Larghezza, DimensioniFattore
    FROM [dbo].[AR]
    WHERE Cd_AR IN ({','.join(['?'] * len(codes))})
    """
    df = pd.read_sql(query, conn, params=codes)
    # Converti peso in grammi se PesoFattore == 1
    df["PesoNetto"] = df.apply(
        lambda r: r["PesoNetto"] * 1000 if r["PesoFattore"] == 1 else r["PesoNetto"], axis=1
    )
    return df


def get_ean(conn, codes):
    query = f"""
    SELECT Cd_AR, Alias
    FROM [dbo].[ARAlias]
    WHERE Cd_AR IN ({','.join(['?'] * len(codes))}) AND TipoAlias = 'E'
    """
    return pd.read_sql(query, conn, params=codes)


def get_ean_cf(conn, codes):
    query = f"""
    WITH RankedResults AS (
        SELECT
            Cd_AR,
            Alias AS Alias_CF,
            Cd_ARMisura,
            ROW_NUMBER() OVER (
                PARTITION BY Cd_AR, TRIM(LOWER(Cd_ARMisura))
                ORDER BY CASE WHEN Alias NOT LIKE '%[^0-9]%' THEN 1 ELSE 2 END
            ) AS rn
        FROM [dbo].[ARAlias]
        WHERE Cd_AR IN ({','.join(['?'] * len(codes))})
          AND TRIM(LOWER(Cd_ARMisura)) = 'cf'
    )
    SELECT Cd_AR, Alias_CF FROM RankedResults WHERE rn = 1
    """
    df = pd.read_sql(query, conn, params=codes)
    df["Cd_AR"] = df["Cd_AR"].str.strip()
    df["Alias_CF"] = df["Alias_CF"].str.strip()
    return df


def get_price_b2c(conn, codes):
    query = f"""
    SELECT Cd_AR, Prezzo
    FROM [dbo].[LSArticolo]
    WHERE Id_LSRevisione = 1
      AND Cd_AR IN ({','.join(['?'] * len(codes))})
    """
    return pd.read_sql(query, conn, params=codes)


def get_ararmisura_data(conn, codes):
    query = f"""
    SELECT Cd_AR, Cd_ARMisura, UMFatt, xUMAltezza, xUMLunghezza, xUMLarghezza
    FROM [dbo].[ARARMisura]
    WHERE Cd_AR IN ({','.join(['?'] * len(codes))})
    """
    return pd.read_sql(query, conn, params=codes)


def get_arimg_note_b(conn, codes):
    query = f"""
    SELECT Cd_AR, Note_ARImg, Picture1OriginalFile
    FROM [dbo].[ARImg]
    WHERE Cd_AR IN ({','.join(['?'] * len(codes))}) AND TipoImg = 'B'
    """
    df = pd.read_sql(query, conn, params=codes)
    df = df.set_index("Cd_AR").reindex(codes).reset_index()
    df["Note_ARImg"] = df.apply(
        lambda r: r["Picture1OriginalFile"] if pd.isna(r["Note_ARImg"]) else r["Note_ARImg"],
        axis=1,
    )
    return df[["Cd_AR", "Note_ARImg"]]


def get_categoria(conn, codes):
    query = f"""
    SELECT AR.Cd_AR, ARG3.Descrizione AS Categoria
    FROM [dbo].[AR] AR
    LEFT JOIN [dbo].[ARGruppo3] ARG3
        ON AR.Cd_ARGruppo3 = ARG3.Cd_ARGruppo3
    WHERE AR.Cd_AR IN ({','.join(['?'] * len(codes))})
    """
    return pd.read_sql(query, conn, params=codes)


# ---------------------------------------------------------------------------
# Assemblaggio dati dal DB
# ---------------------------------------------------------------------------
def fetch_db_data(conn, codes):
    """Ritorna un DataFrame con tutti i dati disponibili dal DB per i codici dati."""

    ar = get_ar_data(conn, codes)
    ean = get_ean(conn, codes)
    ean_cf = get_ean_cf(conn, codes)
    price_b2c = get_price_b2c(conn, codes)
    ararmisura = get_ararmisura_data(conn, codes)
    arimg = get_arimg_note_b(conn, codes)
    categoria = get_categoria(conn, codes)

    # Pulizia spazi
    for df in [ar, ean, ean_cf, price_b2c, arimg, categoria]:
        df["Cd_AR"] = df["Cd_AR"].str.strip()

    # Merge progressivo
    df = ar.merge(ean.rename(columns={"Alias": "Alias_EAN"}), on="Cd_AR", how="left")
    df = df.merge(ean_cf, on="Cd_AR", how="left")
    df = df.merge(price_b2c.rename(columns={"Prezzo": "Prezzo_B2C"}), on="Cd_AR", how="left")
    df = df.merge(arimg.rename(columns={"Note_ARImg": "Percorso_Img"}), on="Cd_AR", how="left")
    df = df.merge(categoria, on="Cd_AR", how="left")

    # Misure CF / LY / PL
    arm_cf = ararmisura.query("Cd_ARMisura.str.upper() == 'CF'")
    arm_ly = ararmisura.query("Cd_ARMisura.str.upper() == 'LY'")
    arm_pl = ararmisura.query("Cd_ARMisura.str.upper() == 'PL'")

    df = df.merge(arm_cf[["Cd_AR", "UMFatt"]].rename(columns={"UMFatt": "UMFatt_CF"}),
                  on="Cd_AR", how="left")
    df = df.merge(arm_ly[["Cd_AR", "UMFatt"]].rename(columns={"UMFatt": "UMFatt_LY"}),
                  on="Cd_AR", how="left")
    df = df.merge(arm_pl[["Cd_AR", "UMFatt"]].rename(columns={"UMFatt": "UMFatt_PL"}),
                  on="Cd_AR", how="left")

    # Conversione dimensioni
    def convert_dim(row):
        if row.get("DimensioniFattore") == 0.001:
            row["Altezza"] = row["Altezza"] * 0.1
            row["Lunghezza"] = row["Lunghezza"] * 0.1
            row["Larghezza"] = row["Larghezza"] * 0.1
        return row

    df = df.apply(convert_dim, axis=1)

    # Calcoli CF×LY, CF×PL, LY×PL
    df["CF_x_LY"] = df["UMFatt_LY"] / df["UMFatt_CF"]
    df["CF_x_PL"] = df["UMFatt_PL"] / df["UMFatt_CF"]
    df["LY_x_PL"] = df["CF_x_PL"] / df["CF_x_LY"]

    # Mappatura verso nomi finali del catalogo
    df = df.rename(columns={
        "Cd_AR":        "cod_art",
        "Descrizione":  "NOME_PRODOTTO",
        "PesoNetto":    "Peso",
        "Alias_EAN":    "ean_pz",
        "Alias_CF":     "ean_cf",
        "UMFatt_CF":    "pz_cf",
        "CF_x_LY":      "cf_ly",
        "CF_x_PL":      "cf_plt",
        "LY_x_PL":      "ly_plt",
        "Percorso_Img": "Immagini",
        "Prezzo_B2C":   "Prezzo_Rivendita_Suggerito",
    })

    return df


# ---------------------------------------------------------------------------
# Validazione booleani
# ---------------------------------------------------------------------------
VALORI_BOOLEANI_VALIDI = {"true", "false", "vero", "falso"}

def is_valid_bool(val):
    if pd.isna(val):
        return False
    return str(val).strip().lower() in VALORI_BOOLEANI_VALIDI


# ---------------------------------------------------------------------------
# Caricamento catalogo_unificato
# ---------------------------------------------------------------------------
def load_catalogo_unificato():
    if not CATALOGO_UNIFICATO.exists():
        print(f"[AVVISO] {CATALOGO_UNIFICATO} non trovato. I campi manuali saranno tutti vuoti.")
        return pd.DataFrame(columns=["cod_art"] + COLONNE_MANUALI)
    df = pd.read_excel(CATALOGO_UNIFICATO, dtype=str)
    df["cod_art"] = df["cod_art"].str.strip()
    return df


# ---------------------------------------------------------------------------
# Generazione catalogo
# ---------------------------------------------------------------------------
def genera(input_file: str):
    input_path = Path(input_file)
    if not input_path.exists():
        print(f"[ERRORE] File non trovato: {input_file}")
        sys.exit(1)

    # Leggi codici articolo dal file del titolare (prima colonna, skip header)
    wb_input = pd.read_excel(input_path, dtype=str)
    codici_raw = wb_input.iloc[:, 0].dropna().str.strip().tolist()
    codici = [c for c in codici_raw if c]
    if not codici:
        print("[ERRORE] Nessun codice articolo trovato nel file di input.")
        sys.exit(1)

    print(f"  Codici da elaborare: {len(codici)}")

    # --- DB ---
    print("  Connessione al database ARCA...")
    conn = get_connection()
    db_df = fetch_db_data(conn, codici)
    conn.close()
    print(f"  Righe recuperate dal DB: {len(db_df)}")

    # --- Catalogo unificato (campi manuali) ---
    cat_df = load_catalogo_unificato()

    # Merge con i campi manuali
    cols_from_cat = ["cod_art"] + [c for c in COLONNE_MANUALI if c in cat_df.columns]
    merged = db_df.merge(cat_df[cols_from_cat], on="cod_art", how="left")

    # Aggiungi colonne mancanti con NaN
    for col in COLONNE_OUTPUT:
        if col not in merged.columns:
            merged[col] = pd.NA

    # Seleziona e ordina le colonne finali
    result = merged[COLONNE_OUTPUT].copy()

    # Normalizza booleani: converti VERO/FALSO in TRUE/FALSE maiuscolo
    for col in COLONNE_BOOLEANE:
        result[col] = result[col].apply(normalize_bool)
        result["ly_plt"] = pd.to_numeric(result["ly_plt"], errors="coerce").round(1).apply(lambda x: int(x) if pd.notna(x) and x % 1 == 0 else x)

    # --- Rilevamento dati mancanti ---
    COLONNE_COLLEGA = ["cod_art", "NOME_PRODOTTO"] + COLONNE_MANUALI
    da_compilare = []

    for _, row in result.iterrows():
        row_dict = {
            "cod_art": row.get("cod_art"),
            "NOME_PRODOTTO": row.get("NOME_PRODOTTO"),
            **{
                c: ""
                if (
                    pd.isna(row.get(c)) or
                    str(row.get(c, "")).strip() == "" or
                    (c in COLONNE_BOOLEANE and not is_valid_bool(row.get(c)))
                )
                else pd.NA
                for c in COLONNE_MANUALI
            }
        }
        if any(pd.notna(row_dict[c]) for c in COLONNE_MANUALI):
            da_compilare.append(row_dict)

    df_collega = pd.DataFrame(da_compilare, columns=COLONNE_COLLEGA)

    # Campi vuoti per articoli presenti
    for _, row in result.iterrows():
        cod = row["cod_art"]
        mancanti = []

        # Campi DB
        db_cols_check = {
            "NOME_PRODOTTO": "NOME_PRODOTTO (Descrizione DB)",
            "Categoria": "Categoria",
            "Peso": "Peso",
            "ean_pz": "ean_pz (EAN pezzo)",
            "pz_cf": "pz_cf (Pezzi per confezione)",
            "Immagini": "Immagini (Percorso img)",
            "Prezzo_Rivendita_Suggerito": "Prezzo_Rivendita_Suggerito",
        }
        for col, label in db_cols_check.items():
            if pd.isna(row.get(col)) or str(row.get(col, "")).strip() == "":
                mancanti.append(label)

        # Campi booleani
        for col in COLONNE_BOOLEANE:
            if not is_valid_bool(row.get(col)):
                mancanti.append(f"{col} (mancante o valore non valido)")

        # Descrizione marketing
        if pd.isna(row.get("Descrizione")) or str(row.get("Descrizione", "")).strip() == "":
            mancanti.append("Descrizione (testo marketing)")


    # --- Scrittura output ---
    result.to_excel(OUTPUT_FILE, index=False)
    print(f"  Catalogo salvato: {OUTPUT_FILE}")

    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if not df_collega.empty:
        df_collega.to_excel(DA_COMPILARE, index=False)
        print(f"  Articoli da compilare: {len(df_collega)} → {DA_COMPILARE}")
    else:
        print("  Nessun campo manuale mancante.")


def normalize_bool(val):
    """Normalizza VERO/FALSO → TRUE/FALSE. Restituisce NaN per valori non validi."""
    if pd.isna(val):
        return pd.NA
    s = str(val).strip().lower()
    if s in ("true", "vero"):
        return "TRUE"
    if s in ("false", "falso"):
        return "FALSE"
    return pd.NA  # valore non valido → considerato mancante


# ---------------------------------------------------------------------------
# Aggiornamento catalogo_unificato
# ---------------------------------------------------------------------------
def aggiorna(input_file: str):
    """
    Aggiorna catalogo_unificato.xlsx con i dati manuali compilati dalla collega.
    Il file in input deve avere almeno la colonna cod_art.
    I campi presenti nel file sovrascrivono quelli esistenti; i nuovi articoli
    vengono aggiunti in fondo.
    """
    input_path = Path(input_file)
    if not input_path.exists():
        print(f"[ERRORE] File non trovato: {input_file}")
        sys.exit(1)

    nuovo_df = pd.read_excel(input_path, dtype=str)
    nuovo_df.columns = [c.strip() for c in nuovo_df.columns]
    if "cod_art" not in nuovo_df.columns:
        print("[ERRORE] Il file deve contenere la colonna 'cod_art'.")
        sys.exit(1)

    nuovo_df["cod_art"] = nuovo_df["cod_art"].str.strip()

    # Normalizza booleani nel file in ingresso
    for col in COLONNE_BOOLEANE:
        if col in nuovo_df.columns:
            nuovo_df[col] = nuovo_df[col].apply(normalize_bool)

    # Carica il catalogo esistente
    if CATALOGO_UNIFICATO.exists():
        esistente = pd.read_excel(CATALOGO_UNIFICATO, dtype=str)
        esistente["cod_art"] = esistente["cod_art"].str.strip()
    else:
        esistente = pd.DataFrame(columns=COLONNE_OUTPUT)

    # Determina colonne da aggiornare (solo quelle presenti nel file della collega)
    colonne_da_aggiornare = [c for c in nuovo_df.columns if c != "cod_art"]

    # Aggiorna le righe esistenti e aggiungi le nuove
    esistente = esistente.set_index("cod_art")
    nuovo_df = nuovo_df.set_index("cod_art")

    for col in colonne_da_aggiornare:
        if col not in esistente.columns:
            esistente[col] = pd.NA
        # Sovrascrive solo le celle dove il nuovo valore è presente
        for cod in nuovo_df.index:
            nuovo_val = nuovo_df.at[cod, col] if col in nuovo_df.columns else pd.NA
            if not pd.isna(nuovo_val) and str(nuovo_val).strip() != "":
                if cod in esistente.index:
                    esistente.at[cod, col] = nuovo_val
                else:
                    esistente.loc[cod, col] = nuovo_val  # nuova riga

    esistente = esistente.reset_index()

    # Riordina colonne secondo COLONNE_OUTPUT (aggiungi quelle extra in fondo)
    cols_finali = [c for c in COLONNE_OUTPUT if c in esistente.columns]
    extra = [c for c in esistente.columns if c not in cols_finali]
    esistente = esistente[cols_finali + extra]

    esistente.to_excel(CATALOGO_UNIFICATO, index=False)
    print(f"  catalogo_unificato.xlsx aggiornato con dati manuali: {CATALOGO_UNIFICATO}")

    # Fetch dati DB per i codici del file della collega
    print("  Connessione al database ARCA per aggiornare i dati DB...")
    codici = list(nuovo_df.index.str.strip())
    conn = get_connection()
    db_df = fetch_db_data(conn, codici)
    conn.close()

    esistente = pd.read_excel(CATALOGO_UNIFICATO, dtype=str)
    esistente["cod_art"] = esistente["cod_art"].str.strip()
    esistente = esistente.set_index("cod_art")
    db_df = db_df.set_index("cod_art")

    DB_COLS = ["NOME_PRODOTTO", "Categoria", "Peso", "ean_pz", "ean_cf",
               "pz_cf", "cf_ly", "cf_plt", "ly_plt", "Immagini", "Prezzo_Rivendita_Suggerito"]

    for col in DB_COLS:
        if col not in esistente.columns:
            esistente[col] = pd.NA
        for cod in db_df.index:
            val = db_df.at[cod, col] if col in db_df.columns else pd.NA
            if pd.notna(val) and str(val).strip() != "":
                if cod in esistente.index:
                    esistente.at[cod, col] = str(val)

    if "ly_plt" in esistente.columns:
        esistente["ly_plt"] = pd.to_numeric(esistente["ly_plt"], errors="coerce").round(1).apply(lambda x: int(x) if pd.notna(x) and x % 1 == 0 else x).astype(str).replace("nan", pd.NA)

    esistente = esistente.reset_index()
    cols_finali = [c for c in COLONNE_OUTPUT if c in esistente.columns]
    extra = [c for c in esistente.columns if c not in cols_finali]
    esistente = esistente[cols_finali + extra]

    esistente.to_excel(CATALOGO_UNIFICATO, index=False)
    print(f"  catalogo_unificato.xlsx aggiornato con dati DB.")
    print(f"  Articoli nel catalogo: {len(esistente)}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="Generazione e aggiornamento catalogo prodotti NOME_AZIENDA"
    )
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument(
        "--genera",
        metavar="FILE_CODICI.xlsx",
        help="Genera catalogo_output.xlsx a partire dai codici articolo nel file dato",
    )
    group.add_argument(
        "--aggiorna",
        metavar="FILE_COLLEGA.xlsx",
        help="Aggiorna catalogo_unificato.xlsx con i dati manuali compilati dalla collega",
    )
    args = parser.parse_args()

    if args.genera:
        print(f"\n[GENERA] Input: {args.genera}")
        genera(args.genera)
        print("[GENERA] Completato.\n")
    elif args.aggiorna:
        print(f"\n[AGGIORNA] Input: {args.aggiorna}")
        aggiorna(args.aggiorna)
        print("[AGGIORNA] Completato.\n")


if __name__ == "__main__":
    main()
