import os
import pandas as pd
import urllib.parse
import re
import logging
from datetime import datetime
from sqlalchemy import create_engine, text, bindparam

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass


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

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# ============ CONFIG ============
INPUT_CSV = "woocommerce_ids.csv"
OUTPUT_EXCEL = "confronto_categorie_CF.xlsx"
OUTPUT_SQL = "update_categorie_CF.sql"
CHUNK_SIZE = 500

# Mappatura User Role (WooCommerce) -> Cd_CFGruppo1 (ARCA)
ROLE_MAPPING = {
    # Esempio: adattare ai ruoli del proprio sito e ai codici CFGruppo di ARCA
    "Altro": "OTH",
    "Ristorazione": "RIS",
    "Negozio alimentari": "ALI",
    "Vending": "VEN",
    "Ingrosso": "ING",
    "Vendita online": "WEB",
}

# Ruoli WooCommerce di sistema, NON sono categorie commerciali -> nessuna azione automatica
RUOLI_NON_MAPPATI = {"Agente/rappresentante", "Amministratore", "Cliente", "Shop manager plus"}
# =================================


class CFCategorieAnalyzer:
    def __init__(self):
        self.engine = None

    def connect_database(self):
        try:
            params = urllib.parse.quote_plus(get_connection_string())
            self.engine = create_engine("mssql+pyodbc:///?odbc_connect=%s" % params)
            with self.engine.connect() as conn:
                conn.execute(text("SELECT 1"))
                logging.info("[OK] Connessione al database stabilita con successo")
                return True
        except Exception as e:
            logging.error(f"[ERROR] Errore nella connessione al database: {str(e)}")
            return False

    def read_csv(self, path):
        df = pd.read_csv(path, sep=";", dtype=str, encoding="utf-8-sig")
        df.columns = [c.strip() for c in df.columns]

        for col in ["ID", "User Role"]:
            if col not in df.columns:
                raise ValueError(f"Colonna '{col}' non trovata nel CSV. Colonne presenti: {list(df.columns)}")

        def clean_id(raw):
            if pd.isna(raw):
                return None
            m = re.search(r"\d+", str(raw))
            return int(m.group()) if m else None

        df["ID_pulito"] = df["ID"].apply(clean_id)
        df = df.dropna(subset=["ID_pulito"])
        df["ID_pulito"] = df["ID_pulito"].astype(int)

        duplicati = df[df.duplicated(subset=["ID_pulito"], keep=False)]
        if not duplicati.empty:
            logging.warning(f"[WARN] {len(duplicati)} righe con ID duplicato nel CSV, tengo solo la prima occorrenza")
        df = df.drop_duplicates(subset=["ID_pulito"], keep="first")

        def map_role(role):
            if pd.isna(role) or str(role).strip() == "":
                return None, "VUOTO"
            role = str(role).strip()
            if role in ROLE_MAPPING:
                return ROLE_MAPPING[role], "OK"
            if role in RUOLI_NON_MAPPATI:
                return None, "RUOLO_SISTEMA_NON_MAPPABILE"
            return None, "RUOLO_SCONOSCIUTO"

        mapped = df["User Role"].apply(map_role)
        df["Codice_Mappato"] = mapped.apply(lambda x: x[0])
        df["Stato_Mappatura"] = mapped.apply(lambda x: x[1])

        logging.info(f"[OK] Letti {len(df)} utenti validi dal CSV")

        sconosciuti = df[df["Stato_Mappatura"] == "RUOLO_SCONOSCIUTO"]["User Role"].unique()
        if len(sconosciuti) > 0:
            logging.warning(f"[WARN] Ruoli non presenti in ROLE_MAPPING: {list(sconosciuti)}")

        return df

    def fetch_cf_data(self, ids):
        results = []
        query = text("""
            SELECT
                Cd_CF, Descrizione, xWooCommerce_ID_B2B,
                Cd_CFGruppo1, Cd_CFGruppo2, Cd_CFGruppo3
            FROM CF
            WHERE xWooCommerce_ID_B2B IN :ids
        """).bindparams(bindparam("ids", expanding=True))

        with self.engine.connect() as conn:
            for i in range(0, len(ids), CHUNK_SIZE):
                chunk = ids[i:i + CHUNK_SIZE]
                res = conn.execute(query, {"ids": chunk})
                results.extend(res.fetchall())

        cols = ["Cd_CF", "Descrizione", "xWooCommerce_ID_B2B", "Cd_CFGruppo1", "Cd_CFGruppo2", "Cd_CFGruppo3"]
        db_df = pd.DataFrame(results, columns=cols)
        # char(n) in SQL Server torna paddato con spazi -> strip
        db_df["Cd_CF"] = db_df["Cd_CF"].str.strip()
        db_df["Descrizione"] = db_df["Descrizione"].str.strip()

        logging.info(f"[OK] Trovati {len(db_df)} record corrispondenti in CF")
        return db_df

    @staticmethod
    def _norm(val):
        if val is None:
            return None
        v = str(val).strip()
        return v if v != "" else None

    def compara(self, csv_df, db_df):
        merged = db_df.merge(
            csv_df[["ID_pulito", "User Role", "Codice_Mappato", "Stato_Mappatura"]],
            left_on="xWooCommerce_ID_B2B", right_on="ID_pulito",
            how="outer", indicator=True
        )

        righe = []
        for _, row in merged.iterrows():
            cd_cf = row.get("Cd_CF")
            wc_id = row["xWooCommerce_ID_B2B"] if pd.notna(row.get("xWooCommerce_ID_B2B")) else row.get("ID_pulito")
            descr = row.get("Descrizione")
            user_role = row.get("User Role")
            codice_csv = row.get("Codice_Mappato")
            stato_mappatura = row.get("Stato_Mappatura")

            g1 = self._norm(row.get("Cd_CFGruppo1"))
            g2 = self._norm(row.get("Cd_CFGruppo2"))
            g3 = self._norm(row.get("Cd_CFGruppo3"))

            if row["_merge"] == "right_only":
                righe.append({
                    "Cd_CF": None, "Descrizione": None, "xWooCommerce_ID_B2B": wc_id,
                    "User_Role_CSV": user_role, "Codice_Mappato": codice_csv,
                    "Cd_CFGruppo1_DB": None, "Cd_CFGruppo2_DB": None, "Cd_CFGruppo3_DB": None,
                    "Azione": "UTENTE_NON_TROVATO_IN_CF",
                    "Note": "ID WooCommerce non presente in CF"
                })
                continue

            if stato_mappatura in ("RUOLO_SISTEMA_NON_MAPPABILE", "RUOLO_SCONOSCIUTO"):
                righe.append({
                    "Cd_CF": cd_cf, "Descrizione": descr, "xWooCommerce_ID_B2B": wc_id,
                    "User_Role_CSV": user_role, "Codice_Mappato": None,
                    "Cd_CFGruppo1_DB": g1, "Cd_CFGruppo2_DB": g2, "Cd_CFGruppo3_DB": g3,
                    "Azione": "DA_VERIFICARE_MANUALMENTE",
                    "Note": f"Ruolo CSV '{user_role}' non mappato automaticamente"
                })
                continue

            db_ha_categoria = any([g1, g2, g3])
            csv_ha_codice = codice_csv is not None

            if not csv_ha_codice and db_ha_categoria:
                azione, note = "RIMUOVI", "Ruolo assente/vuoto nel CSV ma categoria presente su CF"
            elif csv_ha_codice and not db_ha_categoria:
                azione, note = "AGGIUNGI", f"Categoria mancante su CF, da ruolo CSV '{user_role}'"
            elif csv_ha_codice and db_ha_categoria and (g1 != codice_csv or g2 != codice_csv or g3 != codice_csv):
                azione, note = "AGGIORNA", f"CF ({g1}/{g2}/{g3}) diverso da CSV ({codice_csv})"
            else:
                azione, note = "NESSUNA_AZIONE", "Allineato"

            righe.append({
                "Cd_CF": cd_cf, "Descrizione": descr, "xWooCommerce_ID_B2B": wc_id,
                "User_Role_CSV": user_role, "Codice_Mappato": codice_csv,
                "Cd_CFGruppo1_DB": g1, "Cd_CFGruppo2_DB": g2, "Cd_CFGruppo3_DB": g3,
                "Azione": azione, "Note": note
            })

        result_df = pd.DataFrame(righe)
        ordine = {"AGGIUNGI": 0, "AGGIORNA": 1, "RIMUOVI": 2, "DA_VERIFICARE_MANUALMENTE": 3,
                  "UTENTE_NON_TROVATO_IN_CF": 4, "NESSUNA_AZIONE": 5}
        result_df["_ord"] = result_df["Azione"].map(ordine)
        result_df = result_df.sort_values(["_ord", "Cd_CF"]).drop(columns="_ord")
        return result_df

    def export_excel(self, result_df, csv_df, path):
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            # Foglio 1: Confronto completo
            result_df.to_excel(writer, sheet_name="Confronto", index=False)

            # Foglio 2: Riepilogo
            riepilogo = result_df["Azione"].value_counts().reset_index()
            riepilogo.columns = ["Azione", "Conteggio"]
            riepilogo.to_excel(writer, sheet_name="Riepilogo", index=False)

            # Foglio 3: Verifica Manuale (ruoli non mappabili)
            verifica = result_df[result_df["Azione"] == "DA_VERIFICARE_MANUALMENTE"].copy()
            if not verifica.empty:
                verifica["Motivo"] = verifica["User_Role_CSV"].apply(
                    lambda r: f"Ruolo CSV '{r}' non mappato automaticamente"
                )
                verifica = verifica[[
                    "Cd_CF", "Descrizione", "xWooCommerce_ID_B2B", "User_Role_CSV",
                    "Cd_CFGruppo1_DB", "Cd_CFGruppo2_DB", "Cd_CFGruppo3_DB", "Motivo"
                ]]
            else:
                verifica = pd.DataFrame(columns=[
                    "Cd_CF", "Descrizione", "xWooCommerce_ID_B2B", "User_Role_CSV",
                    "Cd_CFGruppo1_DB", "Cd_CFGruppo2_DB", "Cd_CFGruppo3_DB", "Motivo"
                ])
            verifica.to_excel(writer, sheet_name="Verifica Manuale", index=False)

            # Foglio 4: Utente Non Trovato in CF (con dati anagrafici completi dal CSV)
            non_trovati_ids = result_df[result_df["Azione"] == "UTENTE_NON_TROVATO_IN_CF"]["xWooCommerce_ID_B2B"].tolist()
            colonne_csv_utili = [
                "ID_pulito", "First Name", "Last Name", "Società di fatturazione",
                "User Email", "User Role", "User Registered",
                "Referente commerciale", "Email referente commerciale"
            ]
            colonne_presenti = [c for c in colonne_csv_utili if c in csv_df.columns]

            non_trovati = csv_df[csv_df["ID_pulito"].isin(non_trovati_ids)][colonne_presenti].copy()
            non_trovati = non_trovati.rename(columns={"ID_pulito": "xWooCommerce_ID_B2B"})
            non_trovati["Motivo"] = "ID WooCommerce non presente in CF"
            non_trovati.to_excel(writer, sheet_name="Utente Non Trovato", index=False)

        logging.info(f"[OK] Excel esportato: {path} ({len(result_df)} righe totali)")
        logging.info(f"     - Verifica Manuale: {len(verifica)} righe")
        logging.info(f"     - Utente Non Trovato: {len(non_trovati)} righe")

    def export_sql(self, result_df, path):
        lines = [
            f"-- Script generato il {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "-- Aggiornamento categorie CF (Cd_CFGruppo1/2/3) da confronto con WooCommerce",
            "-- VERIFICARE PRIMA DI ESEGUIRE - nessuna query e' stata gia' eseguita sul DB",
            ""
        ]

        aggiunte = result_df[result_df["Azione"] == "AGGIUNGI"]
        aggiornamenti = result_df[result_df["Azione"] == "AGGIORNA"]
        rimozioni = result_df[result_df["Azione"] == "RIMUOVI"]

        if not aggiunte.empty:
            lines.append("-- ============ INSERIMENTI (categoria mancante) ============")
            for _, r in aggiunte.iterrows():
                lines.append(
                    f"UPDATE CF SET Cd_CFGruppo1 = '{r['Codice_Mappato']}', "
                    f"Cd_CFGruppo2 = '{r['Codice_Mappato']}', Cd_CFGruppo3 = '{r['Codice_Mappato']}' "
                    f"WHERE Cd_CF = '{r['Cd_CF']}'; "
                    f"-- {r['Descrizione']} (WC ID {int(r['xWooCommerce_ID_B2B'])}, ruolo: {r['User_Role_CSV']})"
                )
            lines.append("")

        if not aggiornamenti.empty:
            lines.append("-- ============ AGGIORNAMENTI (categoria diversa) ============")
            for _, r in aggiornamenti.iterrows():
                lines.append(
                    f"UPDATE CF SET Cd_CFGruppo1 = '{r['Codice_Mappato']}', "
                    f"Cd_CFGruppo2 = '{r['Codice_Mappato']}', Cd_CFGruppo3 = '{r['Codice_Mappato']}' "
                    f"WHERE Cd_CF = '{r['Cd_CF']}'; "
                    f"-- {r['Descrizione']} (WC ID {int(r['xWooCommerce_ID_B2B'])}, "
                    f"era {r['Cd_CFGruppo1_DB']}/{r['Cd_CFGruppo2_DB']}/{r['Cd_CFGruppo3_DB']}, ruolo: {r['User_Role_CSV']})"
                )
            lines.append("")

        if not rimozioni.empty:
            lines.append("-- ============ RIMOZIONI (ruolo assente su WooCommerce) ============")
            for _, r in rimozioni.iterrows():
                lines.append(
                    f"UPDATE CF SET Cd_CFGruppo1 = NULL, Cd_CFGruppo2 = NULL, Cd_CFGruppo3 = NULL "
                    f"WHERE Cd_CF = '{r['Cd_CF']}'; "
                    f"-- {r['Descrizione']} (WC ID {int(r['xWooCommerce_ID_B2B'])}, "
                    f"era {r['Cd_CFGruppo1_DB']}/{r['Cd_CFGruppo2_DB']}/{r['Cd_CFGruppo3_DB']})"
                )
            lines.append("")

        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))

        logging.info(f"[OK] SQL esportato: {path} ({len(aggiunte)} insert, {len(aggiornamenti)} update, {len(rimozioni)} rimozioni)")


def main():
    analyzer = CFCategorieAnalyzer()
    if not analyzer.connect_database():
        return

    csv_df = analyzer.read_csv(INPUT_CSV)
    ids = csv_df["ID_pulito"].tolist()

    db_df = analyzer.fetch_cf_data(ids)
    result_df = analyzer.compara(csv_df, db_df)

    analyzer.export_excel(result_df, csv_df, OUTPUT_EXCEL)
    analyzer.export_sql(result_df, OUTPUT_SQL)


if __name__ == "__main__":
    main()