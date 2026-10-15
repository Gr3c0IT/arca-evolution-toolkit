"""
cf_gruppo_updater.py
====================
Legge un file Excel con colonne [xWooCommerce_ID_B2B, Ruolo],
esegue fuzzy-matching tra i ruoli WooCommerce e le descrizioni
delle tabelle CFGruppo del database ARCA Evolution,
poi genera un file .sql con gli UPDATE sulla tabella CF.

Output:
  - cf_update_report.txt   → report pre-flight con tutti i match
  - cf_update.sql          → query da eseguire su DB_ARCA
"""

import sys
import re
from pathlib import Path
from difflib import SequenceMatcher
from datetime import datetime

try:
    import openpyxl
except ImportError:
    sys.exit("❌  Installa openpyxl:  pip install openpyxl")

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURAZIONE
# ─────────────────────────────────────────────────────────────────────────────

EXCEL_FILE = "input.xlsx"          # <- modifica con il percorso del tuo Excel
SHEET_NAME = None                  # None = primo foglio
COL_ID     = "xWooCommerce_ID_B2B"
COL_RUOLO  = "Ruolo"

OUTPUT_SQL    = "cf_update.sql"
OUTPUT_REPORT = "cf_update_report.txt"

# Soglia minima di similarità (0.0 – 1.0). Usata SOLO come fallback quando
# il ruolo WC non è presente in MAPPING_MANUALE. Sotto questa soglia il match
# viene segnalato come WARNING nel report e NON incluso nella query SQL.
FUZZY_THRESHOLD = 0.60

# Ruoli WooCommerce da ignorare completamente (non vengono importati).
# Confronto case-insensitive dopo normalizzazione.
RUOLI_ESCLUSI = {
    # versioni italiane
    "agente/rappresentante",
    "agente rappresentante",
    "amministratore",
    "cliente",
    "gestore negozio",
    "grexwne",
    "grexws",
    "nessun ruolo",
    "shop manager plus",
    # versioni inglesi (come appaiono nell'Excel WooCommerce)
    "administrator",
    "customer",
    "shop manager",
    "subscriber",
    "editor",
    "author",
    "contributor",
}

# Mappatura completa CFGruppo (codice → descrizione DB)
# Identica per CFGruppo1, CFGruppo2, CFGruppo3
# Esempio: sostituire con i codici e le descrizioni CFGruppo del proprio ARCA
CF_GRUPPI: dict[str, str] = {
    "ALI": "Negozio Alimentari",
    "ING": "Ingrosso",
    "OTH": "Altro",
    "RIS": "Ristorazione",
    "VEN": "Vending",
    "WEB": "Vendita Online",
}

# ── Mapping manuale WooCommerce → codice DB ───────────────────────────────────
# Chiave  : ruolo WooCommerce normalizzato (lowercase, separatori unificati a spazio)
# Valore  : codice in CF_GRUPPI
# Questo dict ha la PRIORITÀ sul fuzzy matching.
# Se in futuro WooCommerce aggiunge nuovi ruoli, aggiungili qui.
MAPPING_MANUALE: dict[str, str] = {
    "ristorazione":                             "RIS",
    "negozio alimentari":                       "ALI",
    "vending":                                  "VEN",
    "ingrosso":                                 "ING",
    "vendita online":                           "WEB",
    "altro":                                    "OTH",
}

# ─────────────────────────────────────────────────────────────────────────────
# UTILITY
# ─────────────────────────────────────────────────────────────────────────────

def normalizza(testo: str) -> str:
    """Lowercase, strip, sostituisce separatori / – - con spazio."""
    t = testo.lower().strip()
    t = re.sub(r"[/–\-_]+", " ", t)
    t = re.sub(r"\s+", " ", t)
    return t


def fuzzy_score(a: str, b: str) -> float:
    return SequenceMatcher(None, a, b).ratio()


def trova_codice(ruolo_wc: str) -> tuple:
    """
    Dato un ruolo WooCommerce, restituisce (codice, descrizione_db, score, metodo).
    metodo e' 'manuale' o 'fuzzy'.
    Ritorna (None, None, 0.0, 'fuzzy') se non supera la soglia.
    """
    norm_ruolo = normalizza(ruolo_wc)

    # 1. Mapping manuale (priorita' assoluta) --------------------------------
    if norm_ruolo in MAPPING_MANUALE:
        codice = MAPPING_MANUALE[norm_ruolo]
        return codice, CF_GRUPPI[codice], 1.0, "manuale"

    # 2. Fuzzy matching come fallback ----------------------------------------
    best_code  = None
    best_desc  = None
    best_score = 0.0

    for codice, descrizione in CF_GRUPPI.items():
        score = fuzzy_score(norm_ruolo, normalizza(descrizione))
        if score > best_score:
            best_score = score
            best_code  = codice
            best_desc  = descrizione

    if best_score >= FUZZY_THRESHOLD:
        return best_code, best_desc, best_score, "fuzzy"
    return None, None, best_score, "fuzzy"


# ─────────────────────────────────────────────────────────────────────────────
# LETTURA EXCEL
# ─────────────────────────────────────────────────────────────────────────────

def leggi_excel(percorso: str) -> list[dict]:
    """Restituisce lista di dict {id, ruolo} dal foglio Excel."""
    wb = openpyxl.load_workbook(percorso, read_only=True, data_only=True)
    ws = wb[SHEET_NAME] if SHEET_NAME else wb.active

    righe = list(ws.iter_rows(values_only=True))
    if not righe:
        sys.exit("❌  Il foglio Excel è vuoto.")

    intestazione = [str(c).strip() if c is not None else "" for c in righe[0]]

    try:
        idx_id    = intestazione.index(COL_ID)
        idx_ruolo = intestazione.index(COL_RUOLO)
    except ValueError as e:
        sys.exit(f"❌  Colonna non trovata: {e}\n"
                 f"   Intestazioni presenti: {intestazione}")

    dati = []
    for n, riga in enumerate(righe[1:], start=2):
        id_val    = riga[idx_id]
        ruolo_val = riga[idx_ruolo]

        if id_val is None or ruolo_val is None:
            continue
        try:
            id_int = int(id_val)
        except (ValueError, TypeError):
            print(f"⚠️   Riga {n}: ID non numerico ({id_val!r}), saltata.")
            continue

        dati.append({"riga": n, "id": id_int, "ruolo": str(ruolo_val).strip()})

    wb.close()
    return dati


# ─────────────────────────────────────────────────────────────────────────────
# GENERAZIONE SQL + REPORT
# ─────────────────────────────────────────────────────────────────────────────

def genera_output(dati: list[dict]) -> None:
    ora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    sql_lines   = []
    report_lines = []

    # Statistiche
    n_update  = 0
    n_escluso = 0
    n_warning = 0
    n_doppio  = 0

    # Per rilevare ID duplicati nell'Excel
    ids_visti: dict[int, str] = {}

    sql_lines.append(f"-- Generato da cf_gruppo_updater.py il {ora}")
    sql_lines.append(f"-- Sorgente Excel: {EXCEL_FILE}")
    sql_lines.append(f"-- Soglia fuzzy: {FUZZY_THRESHOLD:.0%}")
    sql_lines.append("-- Colonne aggiornate: Cd_CFGruppo1, Cd_CFGruppo2, Cd_CFGruppo3")
    sql_lines.append("-- Le righe già allineate vengono saltate automaticamente (WHERE differisce).")
    sql_lines.append("-- Gli ID con più di una riga in CF vengono esclusi (tabella #id_duplicati).")
    sql_lines.append("")
    sql_lines.append("-- ═══════════════════════════════════════════════════════════════")
    sql_lines.append("-- STEP 1: crea tabella temporanea con gli ID WooCommerce duplicati")
    sql_lines.append("--         (stessa xWooCommerce_ID_B2B su più anagrafiche CF)")
    sql_lines.append("--         Questi ID vengono ESCLUSI dagli UPDATE sottostanti.")
    sql_lines.append("-- ═══════════════════════════════════════════════════════════════")
    sql_lines.append("IF OBJECT_ID('tempdb..#id_duplicati') IS NOT NULL DROP TABLE #id_duplicati;")
    sql_lines.append("")
    sql_lines.append("SELECT xWooCommerce_ID_B2B")
    sql_lines.append("INTO   #id_duplicati")
    sql_lines.append("FROM   CF")
    sql_lines.append("WHERE  xWooCommerce_ID_B2B IS NOT NULL")
    sql_lines.append("GROUP  BY xWooCommerce_ID_B2B")
    sql_lines.append("HAVING COUNT(*) > 1;")
    sql_lines.append("")
    sql_lines.append("-- Verifica: questi ID verranno saltati (solo informativo)")
    sql_lines.append("-- SELECT * FROM #id_duplicati ORDER BY xWooCommerce_ID_B2B;")
    sql_lines.append("")
    sql_lines.append("-- ═══════════════════════════════════════════════════════════════")
    sql_lines.append("-- STEP 2: UPDATE categorie (solo ID univoci)")
    sql_lines.append("-- ═══════════════════════════════════════════════════════════════")
    sql_lines.append("")

    report_lines.append("=" * 72)
    report_lines.append(f"  REPORT PRE-FLIGHT  —  {ora}")
    report_lines.append(f"  Sorgente: {EXCEL_FILE}")
    report_lines.append("=" * 72)
    report_lines.append("")

    sql_lines.append("-- ═══════════════════════════════════════════════════════════════")
    sql_lines.append("-- STEP 2: disabilita trigger CLR, esegui UPDATE, riabilita")
    sql_lines.append("-- ═══════════════════════════════════════════════════════════════")
    sql_lines.append("DISABLE TRIGGER CF_atrg_brd ON CF;")
    sql_lines.append("")

    # Raccoglie tutte le coppie (wc_id, codice) valide prima di scrivere il SQL
    mapping_valido: list[tuple[int, str]] = []

    for entry in dati:
        riga  = entry["riga"]
        wc_id = entry["id"]
        ruolo = entry["ruolo"]

        # ── Duplicati ──────────────────────────────────────────────────────
        if wc_id in ids_visti:
            n_doppio += 1
            report_lines.append(
                f"⚠️  DUPLICATO  | Riga Excel {riga:>5} | ID {wc_id:>8} "
                f"| Ruolo: {ruolo!r}  (già visto a riga {ids_visti[wc_id]})"
            )
            report_lines.append("")
            continue
        ids_visti[wc_id] = riga

        # ── Ruoli esclusi ──────────────────────────────────────────────────
        if normalizza(ruolo) in RUOLI_ESCLUSI:
            n_escluso += 1
            report_lines.append(
                f"➖  ESCLUSO   | Riga Excel {riga:>5} | ID {wc_id:>8} "
                f"| Ruolo: {ruolo!r}  → ignorato per policy"
            )
            report_lines.append("")
            continue

        # ── Match ──────────────────────────────────────────────────────────
        codice, desc_db, score, metodo = trova_codice(ruolo)
        score_pct    = f"{score:.0%}"
        label_metodo = "MANUALE" if metodo == "manuale" else f"FUZZY {score_pct}"

        if codice is None:
            n_warning += 1
            report_lines.append(
                f"❌  NO MATCH  | Riga Excel {riga:>5} | ID {wc_id:>8} "
                f"| Ruolo: {ruolo!r}  → score massimo {score_pct} "
                f"(soglia {FUZZY_THRESHOLD:.0%}) — RIGA NON INCLUSA NEL SQL"
            )
            report_lines.append("")
            continue

        n_update += 1
        report_lines.append(
            f"✅  OK        | Riga Excel {riga:>5} | ID {wc_id:>8} "
            f"| {ruolo!r:45s} → {codice} ({desc_db})  [{label_metodo}]"
        )
        report_lines.append("")
        mapping_valido.append((wc_id, codice))

    # ── Genera UPDATE set-based a batch da 1000 ────────────────────────────
    BATCH_SIZE = 1000
    n_batch = (len(mapping_valido) + BATCH_SIZE - 1) // BATCH_SIZE

    for b in range(n_batch):
        slice_ = mapping_valido[b * BATCH_SIZE : (b + 1) * BATCH_SIZE]
        sql_lines.append(f"-- Batch {b + 1}/{n_batch}  ({len(slice_)} righe)")
        sql_lines.append("UPDATE CF")
        sql_lines.append("SET    Cd_CFGruppo1 = m.codice,")
        sql_lines.append("       Cd_CFGruppo2 = m.codice,")
        sql_lines.append("       Cd_CFGruppo3 = m.codice")
        sql_lines.append("FROM   CF")
        sql_lines.append("INNER JOIN (VALUES")

        values_rows = [f"    ({wc_id}, '{codice}')" for wc_id, codice in slice_]
        sql_lines.append(",\n".join(values_rows))

        sql_lines.append(") AS m(wc_id, codice) ON CF.xWooCommerce_ID_B2B = m.wc_id")
        sql_lines.append("WHERE  CF.xWooCommerce_ID_B2B NOT IN (SELECT xWooCommerce_ID_B2B FROM #id_duplicati)")
        sql_lines.append("  AND  (CF.Cd_CFGruppo1 <> m.codice OR CF.Cd_CFGruppo2 <> m.codice OR CF.Cd_CFGruppo3 <> m.codice")
        sql_lines.append("        OR CF.Cd_CFGruppo1 IS NULL  OR CF.Cd_CFGruppo2 IS NULL  OR CF.Cd_CFGruppo3 IS NULL);")
        sql_lines.append("")

    sql_lines.append("ENABLE TRIGGER CF_atrg_brd ON CF;")
    sql_lines.append("")
    # ── Riepilogo ──────────────────────────────────────────────────────────
    totale = len(dati)
    riepilogo = [
        "",
        "─" * 72,
        f"  RIEPILOGO",
        f"  Righe Excel lette       : {totale}",
        f"  UPDATE generati         : {n_update}",
        f"  Ruoli esclusi (policy)  : {n_escluso}",
        f"  Duplicati ID skippati   : {n_doppio}",
        f"  No match (⚠️  verificare): {n_warning}",
        "─" * 72,
    ]
    if n_warning:
        riepilogo.append("")
        riepilogo.append(
            f"  ⚠️  ATTENZIONE: {n_warning} ruolo/i non ha trovato corrispondenza "
            f"con score >= {FUZZY_THRESHOLD:.0%}."
        )
        riepilogo.append(
            "     Controlla il report, aggiorna RUOLI_ESCLUSI o CF_GRUPPI nello script,"
        )
        riepilogo.append(
            "     oppure abbassa FUZZY_THRESHOLD se sei sicuro del match manuale."
        )

    report_lines += riepilogo
    sql_lines.append(f"-- Fine script — {n_update} UPDATE generati su {totale} righe lette")
    sql_lines.append("")
    sql_lines.append("-- ═══════════════════════════════════════════════════════════════")
    sql_lines.append("-- STEP 3: riepilogo ID esclusi (duplicati — da gestire automaticamente)")
    sql_lines.append("-- ═══════════════════════════════════════════════════════════════")
    sql_lines.append("SELECT CF.xWooCommerce_ID_B2B, CF.Cd_CF, CF.Descrizione,")
    sql_lines.append("       CF.Cd_CFGruppo1, CF.Cd_CFGruppo2, CF.Cd_CFGruppo3")
    sql_lines.append("FROM   CF")
    sql_lines.append("INNER JOIN #id_duplicati d ON CF.xWooCommerce_ID_B2B = d.xWooCommerce_ID_B2B")
    sql_lines.append("ORDER  BY CF.xWooCommerce_ID_B2B, CF.Cd_CF;")
    sql_lines.append("")
    sql_lines.append("-- ═══════════════════════════════════════════════════════════════")
    sql_lines.append("-- STEP 4: per i duplicati, copia la categoria dalla riga")
    sql_lines.append("--         già valorizzata alle righe sorelle con NULL.")
    sql_lines.append("--         Copre automaticamente la maggior parte dei casi.")
    sql_lines.append("-- ═══════════════════════════════════════════════════════════════")
    sql_lines.append("DISABLE TRIGGER CF_atrg_brd ON CF;")
    sql_lines.append("")
    sql_lines.append("UPDATE CF_null")
    sql_lines.append("SET    CF_null.Cd_CFGruppo1 = CF_src.Cd_CFGruppo1,")
    sql_lines.append("       CF_null.Cd_CFGruppo2 = CF_src.Cd_CFGruppo2,")
    sql_lines.append("       CF_null.Cd_CFGruppo3 = CF_src.Cd_CFGruppo3")
    sql_lines.append("FROM   CF AS CF_null")
    sql_lines.append("INNER JOIN CF AS CF_src")
    sql_lines.append("       ON  CF_null.xWooCommerce_ID_B2B = CF_src.xWooCommerce_ID_B2B")
    sql_lines.append("      AND  CF_src.Cd_CFGruppo1 IS NOT NULL")
    sql_lines.append("WHERE  CF_null.Cd_CFGruppo1 IS NULL")
    sql_lines.append("  AND  CF_null.xWooCommerce_ID_B2B IN (SELECT xWooCommerce_ID_B2B FROM #id_duplicati);")
    sql_lines.append("")
    sql_lines.append("ENABLE TRIGGER CF_atrg_brd ON CF;")
    sql_lines.append("")
    sql_lines.append("-- ═══════════════════════════════════════════════════════════════")
    sql_lines.append("-- STEP 5: duplicati ancora con categoria NULL dopo lo STEP 4")
    sql_lines.append("--         → richiedono intervento manuale in ARCA")
    sql_lines.append("-- ═══════════════════════════════════════════════════════════════")
    sql_lines.append("SELECT CF.xWooCommerce_ID_B2B, CF.Cd_CF, CF.Descrizione,")
    sql_lines.append("       CF.Cd_CFGruppo1, CF.Cd_CFGruppo2, CF.Cd_CFGruppo3")
    sql_lines.append("FROM   CF")
    sql_lines.append("WHERE  CF.xWooCommerce_ID_B2B IN (SELECT xWooCommerce_ID_B2B FROM #id_duplicati)")
    sql_lines.append("  AND  CF.Cd_CFGruppo1 IS NULL")
    sql_lines.append("ORDER  BY CF.xWooCommerce_ID_B2B, CF.Cd_CF;")
    sql_lines.append("")
    sql_lines.append("-- Cleanup")
    sql_lines.append("DROP TABLE #id_duplicati;")

    # ── Scrittura file ─────────────────────────────────────────────────────
    Path(OUTPUT_REPORT).write_text("\n".join(report_lines), encoding="utf-8")
    Path(OUTPUT_SQL).write_text("\n".join(sql_lines), encoding="utf-8")

    print("\n".join(report_lines))
    print(f"\n✅  SQL salvato in:    {OUTPUT_SQL}")
    print(f"✅  Report salvato in: {OUTPUT_REPORT}")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    excel_path = sys.argv[1] if len(sys.argv) > 1 else EXCEL_FILE
    print(f"📂  Lettura file: {excel_path}")
    dati = leggi_excel(excel_path)
    print(f"📋  Righe valide trovate: {len(dati)}")
    genera_output(dati)