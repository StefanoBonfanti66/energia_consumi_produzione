
import pandas as pd
import openpyxl
import os

def crea_foglio_consolidato():
    """
    Legge i dati dai fogli delle singole macchine in 'Dati consumi e costi energetici.xlsx',
    li consolida e li salva nel foglio 'Consolidato' dello stesso file.
    """
    file_path = "Dati consumi e costi energetici.xlsx"
    
    if not os.path.exists(file_path):
        print(f"Errore: Il file '{file_path}' non è stato trovato.")
        return

    try:
        # Carica il file Excel per ispezionare i fogli
        xls = pd.ExcelFile(file_path, engine='openpyxl')
        sheet_names = xls.sheet_names

        # Filtra i fogli da elaborare (escludi quelli di sistema/riepilogo)
        sheets_to_process = [s for s in sheet_names if s not in ['Consolidato', 'Tabelle', 'ICOPOWER']]
        print(f"Fogli trovati da elaborare: {sheets_to_process}")

        if not sheets_to_process:
            print("Nessun foglio di macchina trovato da elaborare.")
            return

        # Lista per contenere i DataFrame di ogni foglio
        all_data = []

        # Itera su ogni foglio e leggi i dati
        for sheet_name in sheets_to_process:
            print(f"Leggo il foglio: {sheet_name}...")
            try:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=0)
                # Assicura che la colonna 'macchina o impianto' esista prima di procedere
                if 'macchina o impianto' in df.columns:
                    # Rimuovi le righe dove la colonna 'macchina o impianto' è vuota
                    df.dropna(subset=['macchina o impianto'], inplace=True)
                    if not df.empty:
                        all_data.append(df)
                else:
                    print(f"Attenzione: La colonna 'macchina o impianto' non è stata trovata nel foglio '{sheet_name}'.")
            except Exception as e:
                print(f"Errore durante la lettura del foglio '{sheet_name}': {e}")

        if not all_data:
            print("Nessun dato valido trovato nei fogli delle macchine.")
            return

        # Concatena tutti i DataFrame in uno solo
        consolidated_df = pd.concat(all_data, ignore_index=True)
        print("Dati da tutti i fogli uniti con successo.")

        # Pulisci i nomi delle colonne (rimuovi spazi bianchi iniziali/finali)
        consolidated_df.columns = consolidated_df.columns.str.strip()

        # Rinomina le colonne per coerenza con la dashboard
        column_mapping = {
            'macchina o impianto': 'macchina',
            'ore produzione macchina': 'ore_produzione',
            'pezzi prodotti': 'pezzi_prodotti',
            'consumo': 'consumo_kwh',
            'costo energia': 'costo_energia_per_kwh',
            'costo macchina': 'costo_macchina',
            'consumo da bolletta': 'consumo_bolletta_kwh',
            'totale bolletta': 'totale_bolletta'
        }
        consolidated_df.rename(columns=column_mapping, inplace=True)
        
        # Assicurati che tutte le colonne mappate esistano
        final_columns = [
            'anno', 'mese', 'macchina', 'ore_produzione', 'pezzi_prodotti',
            'consumo_kwh', 'lettura', 'costo_energia_per_kwh', 'costo_macchina',
            'consumo_bolletta_kwh', 'totale_bolletta'
        ]
        
        # Aggiungi colonne mancanti con None se non esistono
        for col in final_columns:
            if col not in consolidated_df.columns:
                consolidated_df[col] = None
        
        # Seleziona e ordina le colonne finali
        consolidated_df = consolidated_df[final_columns]

        # Salva il DataFrame consolidato nel foglio 'Consolidato'
        print(f"Salvataggio del foglio 'Consolidato' nel file '{file_path}'...")
        try:
            from openpyxl.utils.dataframe import dataframe_to_rows

            book = openpyxl.load_workbook(file_path)

            # Rimuovi il vecchio foglio 'Consolidato' se esiste
            if 'Consolidato' in book.sheetnames:
                old_sheet = book['Consolidato']
                book.remove(old_sheet)

            # Crea un nuovo foglio 'Consolidato'
            new_sheet = book.create_sheet('Consolidato')

            # Scrivi il DataFrame nel nuovo foglio, inclusa l'intestazione
            for r in dataframe_to_rows(consolidated_df, index=False, header=True):
                new_sheet.append(r)

            book.save(file_path)
            print("Operazione completata con successo!")

        except Exception as e:
            print(f"Errore durante il salvataggio con openpyxl: {e}")
            print("Tentativo di salvataggio alternativo con Pandas...")
            try:
                with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    consolidated_df.to_excel(writer, sheet_name='Consolidato', index=False)
                print("Salvataggio alternativo con Pandas riuscito!")
            except Exception as e2:
                print(f"Anche il salvataggio alternativo è fallito: {e2}")

    except FileNotFoundError:
        print(f"Errore critico: Il file '{file_path}' non è stato trovato.")
    except Exception as e:
        print(f"Si è verificato un errore imprevisto: {e}")

if __name__ == "__main__":
    crea_foglio_consolidato()
