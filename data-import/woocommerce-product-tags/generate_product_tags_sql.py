"""
Script per generare la query SQL di import in xARARTags_WooCommerce
partendo da un file Excel con colonne:
  A -> Cd_AR
  D -> Id_xARTags_WooCommerce (valori ammessi: quelli in VALORI_AMMESSI, oppure vuoto)

NON si connette al database: legge l'Excel e produce un file .sql
da eseguire manualmente (es. in SSMS).

La tabella xARARTags_WooCommerce usa direttamente Cd_AR (codice
testuale, stesso valore presente in colonna A dell'Excel), quindi
non serve alcun lookup/JOIN: si inserisce Cd_AR + Id_xARTags_WooCommerce
così come letti dal foglio.

Uso:
    python genera_sql_woocommerce_tags.py input.xlsx output.sql
"""

import sys
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit("Manca il pacchetto openpyxl. Installa con: pip install openpyxl")

BATCH_SIZE = 900  # margine di sicurezza sotto il limite di 1000 righe per VALUES in SQL Server
# ID dei tag WooCommerce ammessi in colonna D: sostituirli con quelli del proprio sito
VALORI_AMMESSI = {101, 102}


def sql_escape(value: str) -> str:
    """Esegue l'escape degli apici singoli per l'uso in una stringa SQL."""
    return str(value).replace("'", "''")


def leggi_righe(percorso_excel: str):
    """
    Legge il file Excel e restituisce:
      - righe_valide: lista di tuple (Cd_AR, Id_xARTags_WooCommerce)
      - righe_escluse: lista di Cd_AR con cella D vuota (esclusi, comportamento atteso)
      - righe_anomale: lista di (Cd_AR, valore) con valore in D non presente in VALORI_AMMESSI
    """
    wb = load_workbook(percorso_excel, data_only=True, read_only=True)
    ws = wb.active

    righe_valide = []
    righe_escluse = []
    righe_anomale = []

    # Parte da riga 2 assumendo la riga 1 come intestazione (Cd_AR, ..., ..., Id_xARTags_WooCommerce)
    for row_idx, row in enumerate(
        ws.iter_rows(min_row=2, values_only=True), start=2
    ):
        if row is None:
            continue

        cd_ar = row[0] if len(row) > 0 else None
        id_tag = row[3] if len(row) > 3 else None

        if cd_ar is None or str(cd_ar).strip() == "":
            continue  # riga senza codice articolo, la saltiamo

        cd_ar = str(cd_ar).strip()

        if id_tag is None or str(id_tag).strip() == "":
            righe_escluse.append(cd_ar)
            continue

        try:
            id_tag_int = int(float(id_tag))
        except (ValueError, TypeError):
            righe_anomale.append((cd_ar, id_tag))
            continue

        if id_tag_int not in VALORI_AMMESSI:
            righe_anomale.append((cd_ar, id_tag_int))
            continue

        righe_valide.append((cd_ar, id_tag_int))

    wb.close()
    return righe_valide, righe_escluse, righe_anomale


def genera_batch_sql(batch):
    """Genera lo statement INSERT...SELECT (senza JOIN, Cd_AR usato direttamente) per un singolo batch."""
    values_clause = ",\n".join(
        f"        ('{sql_escape(cd_ar)}', {id_tag})" for cd_ar, id_tag in batch
    )

    return f"""INSERT INTO dbo.xARARTags_WooCommerce (Cd_AR, Id_xARTags_WooCommerce)
SELECT v.Cd_AR, v.Id_xARTags_WooCommerce
FROM (VALUES
{values_clause}
    ) AS v(Cd_AR, Id_xARTags_WooCommerce)
WHERE NOT EXISTS (
    SELECT 1
    FROM dbo.xARARTags_WooCommerce x
    WHERE x.Cd_AR = v.Cd_AR
      AND x.Id_xARTags_WooCommerce = v.Id_xARTags_WooCommerce
);
"""


def main():
    if len(sys.argv) != 3:
        sys.exit("Uso: python genera_sql_woocommerce_tags.py input.xlsx output.sql")

    percorso_excel = sys.argv[1]
    percorso_output = sys.argv[2]

    if not Path(percorso_excel).exists():
        sys.exit(f"File non trovato: {percorso_excel}")

    righe_valide, righe_escluse, righe_anomale = leggi_righe(percorso_excel)

    if not righe_valide:
        sys.exit(f"Nessuna riga valida trovata (Id_xARTags_WooCommerce in {sorted(VALORI_AMMESSI)}). Controlla il file.")

    batches = [
        righe_valide[i : i + BATCH_SIZE]
        for i in range(0, len(righe_valide), BATCH_SIZE)
    ]

    with open(percorso_output, "w", encoding="utf-8") as f:
        f.write("-- Script generato automaticamente da genera_sql_woocommerce_tags.py\n")
        f.write(f"-- Righe da inserire: {len(righe_valide)} (in {len(batches)} batch da max {BATCH_SIZE})\n")
        f.write(f"-- Righe escluse (Id_xARTags_WooCommerce vuoto): {len(righe_escluse)}\n")
        f.write(f"-- Righe anomale (valore non ammesso): {len(righe_anomale)}\n\n")
        f.write("SET NOCOUNT ON;\n")
        f.write("BEGIN TRY\n")
        f.write("    BEGIN TRANSACTION;\n\n")

        for i, batch in enumerate(batches, start=1):
            f.write(f"    -- Batch {i}/{len(batches)} ({len(batch)} righe)\n")
            sql_batch = genera_batch_sql(batch)
            # indenta ogni riga dello statement per leggibilità dentro il TRY
            indented = "\n".join("    " + line if line.strip() else line for line in sql_batch.splitlines())
            f.write(indented + "\n\n")

        f.write("    COMMIT TRANSACTION;\n")
        f.write("    PRINT 'Import completato con successo.';\n")
        f.write("END TRY\n")
        f.write("BEGIN CATCH\n")
        f.write("    IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;\n")
        f.write("    PRINT 'Errore durante import: ' + ERROR_MESSAGE();\n")
        f.write("    THROW;\n")
        f.write("END CATCH;\n")

    # Report a video
    print(f"File SQL generato: {percorso_output}")
    print(f"Righe da inserire : {len(righe_valide)}")
    print(f"Righe escluse (cella D vuota): {len(righe_escluse)}")
    if righe_anomale:
        print(f"ATTENZIONE - Righe con valore anomalo in colonna D (non ammesso): {len(righe_anomale)}")
        for cd_ar, val in righe_anomale[:20]:
            print(f"  Cd_AR={cd_ar}  valore={val}")
        if len(righe_anomale) > 20:
            print(f"  ... e altre {len(righe_anomale) - 20} righe anomale")


if __name__ == "__main__":
    main()