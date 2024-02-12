import os
import pandas as pd
from tqdm import tqdm

def converti_dpt_a_excel(input_folder, output_folder):
    # Verifica se la cartella di output esiste, altrimenti creala
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Ottieni la lista di file .dpt nella cartella di input
    files_to_process = [file_name for file_name in os.listdir(input_folder) if file_name.endswith(".dpt")]

    # Inizializza la barra di avanzamento
    progress_bar = tqdm(total=len(files_to_process), desc='Conversione in corso')

    # Elabora tutti i file .dpt nella cartella di input
    for file_name in files_to_process:
        input_path = os.path.join(input_folder, file_name)
        output_path = os.path.join(output_folder, os.path.splitext(file_name)[0] + ".xlsx")

        # Leggi il file .dpt utilizzando pandas con la codifica 'latin-1'
        try:
            df = pd.read_csv(input_path, delimiter=',', header=None, names=['Colonna1', 'Colonna2'], encoding='latin-1')
        except Exception as e:
            print(f"Errore durante la lettura del file .dpt {file_name}: {e}")
            progress_bar.update(1)
            continue
        
        # Scrivi i dati formattati in un file Excel
        try:
            df.to_excel(output_path, index=False)
            print(f"Conversione completata con successo. File Excel salvato in: {output_path}")
        except Exception as e:
            print(f"Errore durante la scrittura del file Excel {file_name}: {e}")
        
        # Aggiorna la barra di avanzamento
        progress_bar.update(1)

    # Chiudi la barra di avanzamento
    progress_bar.close()

if __name__ == "__main__":
    # Ottieni il percorso dello script
    script_path = os.path.dirname(os.path.realpath(__file__))

    # Ottieni i percorsi delle cartelle "data" e "output"
    input_folder = os.path.join(script_path, "data")
    output_folder = os.path.join(script_path, "output")

    # Chiama la funzione di conversione
    converti_dpt_a_excel(input_folder, output_folder)
