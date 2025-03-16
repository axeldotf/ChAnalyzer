import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from tqdm import tqdm
import warnings

warnings.simplefilter(action='ignore', category=FutureWarning)

def map_operator_technology(value):
    mapping = {
        6300: ('TIM', 'L800'), 6400: ('VF', 'L800'), 6200: ('W3', 'L800'),
        1350: ('TIM', 'L1800'), 1500: ('Iliad', 'L1800'), 1650: ('W3', 'L1800'), 1850: ('VF', 'L1800'),
        2900: ('Iliad', 'L2600'), 3025: ('VF', 'L2600'), 3350: ('W3', 'L2600'), 3175: ('TIM', 'L2600'),
        125: ('W3', 'L2100'), 275: ('TIM', 'L2100'), 525: ('VF', 'L2100'), 400: ('Iliad', 'L2100'),
        2938: ('Iliad', 'U900'), 3063: ('W3', 'U900'), 10563: ('W3', 'U2100'), 100: ('W3', 'L2100')
    }
    gsm_mapping = {range(1, 26): ('TIM', 'G900'), range(1000, 1024): ('TIM', 'G900'),
                   range(27, 76): ('VF', 'G900'), range(77, 125): ('W3', 'G900')}
    if value in mapping:
        return mapping[value]
    for key, val in gsm_mapping.items():
        if value in key:
            return val
    return ('Unknown', 'Unknown')

def apply_conditional_formatting(writer, sheet_name, df):
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    
    format_green = workbook.add_format({'bg_color': '#99FF99', 'border': 1})  # Verde chiaro
    format_yellow = workbook.add_format({'bg_color': '#FFFF99', 'border': 1})  # Giallo tenue
    format_red = workbook.add_format({'bg_color': '#FFCCCC', 'border': 1})  # Rosso chiaro
    format_gray = workbook.add_format({'bg_color': '#BFBFBF', 'border': 1})  # Grigio per valori assenti o zero
    format_double_border = workbook.add_format({'top': 2})  # Bordo doppio superiore
    
    num_rows = len(df)
    num_cols = len(df.columns)
    
    for col_idx in range(num_cols):  # Adatta la larghezza in base alla prima riga
        col_letter = chr(65 + col_idx)
        max_length = max(df.iloc[:, col_idx].astype(str).apply(len).max(), len(df.columns[col_idx]))
        worksheet.set_column(f'{col_letter}:{col_letter}', max_length + 2)
    
    for col_idx in range(3, num_cols):  # Partiamo dalla quarta colonna (misure)
        col_letter = chr(65 + col_idx)
        
        worksheet.conditional_format(f'{col_letter}2:{col_letter}{num_rows+1}', 
                                     {'type': 'cell', 'criteria': 'between', 'minimum': -90, 'maximum': -1, 'format': format_green})
        worksheet.conditional_format(f'{col_letter}2:{col_letter}{num_rows+1}', 
                                     {'type': 'cell', 'criteria': 'between', 'minimum': -90, 'maximum': -80, 'format': format_yellow})
        worksheet.conditional_format(f'{col_letter}2:{col_letter}{num_rows+1}', 
                                     {'type': 'cell', 'criteria': 'less than', 'value': -90, 'format': format_red})
        worksheet.conditional_format(f'{col_letter}2:{col_letter}{num_rows+1}', 
                                     {'type': 'cell', 'criteria': 'equal to', 'value': 0, 'format': format_gray})
        worksheet.conditional_format(f'{col_letter}2:{col_letter}{num_rows+1}', 
                                     {'type': 'blanks', 'format': format_gray})
    
    # Aggiunta di bordi doppi superiori per separare i piani
    prev_area = None
    for row in range(1, num_rows + 1):
        current_area = df.iloc[row - 1]['Area Misurata'] if 'Area Misurata' in df.columns else None
        if prev_area is not None and current_area != prev_area:
            worksheet.set_row(row, None, format_double_border)
        prev_area = current_area

def process_excel_files():
    required_technologies = ['G900', 'L800', 'L1800', 'U900', 'U2100', 'L2100', 'L2600']
    summary_data = []
    
    root = tk.Tk()
    root.withdraw()
    
    output_file = filedialog.asksaveasfilename(title="Seleziona il file di output", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not output_file:
        print("Nessun file di output selezionato. Uscita...")
        return
    
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        while True:
            input_file = filedialog.askopenfilename(title="Seleziona un file Excel", filetypes=[("Excel files", "*.xlsx;*.xls")])
            if not input_file:
                break
            
            sheet_name = simpledialog.askstring("Inserisci nome foglio", "Nome del piano:")
            if not sheet_name:
                break
            
            print(f"Caricamento del file {input_file}...")
            df = pd.read_excel(input_file)
            df['Group'] = df['Ch'].fillna(df['ARFCN'])
            
            with tqdm(total=3, desc="Elaborazione dati") as pbar:
                measure_avg = df.groupby('Group', dropna=True)[['1. best RSRP', '1. best RSCP', '1. best Rx Level']].mean().mean(axis=1).round(2)
                pbar.update(3)
            
            output_df = pd.DataFrame({'Ch/ARFCN': measure_avg.index, 'Misura': measure_avg.values})
            output_df[['Operatore', 'Tecnologia']] = output_df['Ch/ARFCN'].apply(lambda x: pd.Series(map_operator_technology(x)))
            
            final_df = output_df.groupby(['Operatore', 'Tecnologia'], dropna=True).mean(numeric_only=True).round(2).reset_index()
            final_df['Ch/ARFCN'] = final_df.apply(lambda row: output_df[(output_df['Operatore'] == row['Operatore']) &
                                                                        (output_df['Tecnologia'] == row['Tecnologia'])]
                                                  .sort_values(by=['Misura'], ascending=False)
                                                  .iloc[0]['Ch/ARFCN'], axis=1)
            
            for operator in final_df['Operatore'].unique():
                for tech in required_technologies:
                    if not ((final_df['Operatore'] == operator) & (final_df['Tecnologia'] == tech)).any():
                        final_df.loc[len(final_df)] = {'Operatore': operator, 'Tecnologia': tech, 'Ch/ARFCN': None, 'Misura': None}
            
            final_df['Tecnologia'] = pd.Categorical(final_df['Tecnologia'], categories=required_technologies, ordered=True)
            final_df = final_df.sort_values(by=['Operatore', 'Tecnologia']).reset_index(drop=True)
            
            final_df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            for tech in required_technologies:
                row = {"Area Misurata": sheet_name, "Tecnologia": tech, "Listino Inwit": "VOCE" if "G900" in tech else "DATI" if "L800" in tech else "DATI PLUS" if "L1800" in tech else "Altro"}
                for op in ['TIM', 'VF', 'W3', 'Iliad']:
                    row[f"Copertura {op}"] = final_df.loc[(final_df['Operatore'] == op) & (final_df['Tecnologia'] == tech), 'Misura'].values[0] if not final_df.loc[(final_df['Operatore'] == op) & (final_df['Tecnologia'] == tech), 'Misura'].empty else "/"
                summary_data.append(row)

            # apply_conditional_formatting(writer, sheet_name, final_df)
            
            print(f"Dati per {sheet_name} salvati.\n")
            
            more_files = messagebox.askyesno("Aggiungere un altro piano?", "Vuoi selezionare un altro file?")
            if not more_files:
                break
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name="Riepilogo", index=False)
        apply_conditional_formatting(writer, "Riepilogo", summary_df)
        print(f"Dati riepilogativi salvati.")
    
    print(f"Dati elaborati salvati in {output_file}\n")

if __name__ == "__main__":
    process_excel_files()
