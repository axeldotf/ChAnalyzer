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

def add_legend_to_summary_sheet(writer, sheet_name, row_offset=2):
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    
    legend_start_row = row_offset
    legend_start_col = 8  # Posizionata fuori dalla tabella dati
    
    header_format = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': '#D9E1F2', 'border': 1})
    cell_format = workbook.add_format({'align': 'center', 'border': 1})
    color_formats = {
        'Verde': workbook.add_format({'bg_color': '#99FF99', 'border': 1, 'align': 'center'}),
        'Giallo': workbook.add_format({'bg_color': '#FFFF99', 'border': 1, 'align': 'center'}),
        'Rosso': workbook.add_format({'bg_color': '#FFCCCC', 'border': 1, 'align': 'center'}),
        'Grigio': workbook.add_format({'bg_color': '#BFBFBF', 'border': 1, 'align': 'center'})
    }
    
    worksheet.merge_range(legend_start_row, legend_start_col, legend_start_row, legend_start_col + 3, "Legenda Copertura", header_format)
    worksheet.write(legend_start_row + 1, legend_start_col, "Colore", header_format)
    worksheet.write(legend_start_row + 1, legend_start_col + 1, "Intervallo 2G", header_format)
    worksheet.write(legend_start_row + 1, legend_start_col + 2, "Intervallo 3G", header_format)
    worksheet.write(legend_start_row + 1, legend_start_col + 3, "Intervallo 4G", header_format)
    
    legend_data = [
        ("Verde", " > -80 dBm", " > -85 dBm", " > -95 dBm"),
        ("Giallo", "-80 ÷ -90 dBm", "-85 ÷ -95 dBm", "-95 ÷ -105 dBm"),
        ("Rosso", "< -90 dBm", "< -95 dBm", "< -105 dBm"),
        ("Grigio", "Assente", "Assente", "Assente")
    ]
    
    max_widths = [max(len(row[i]) for row in legend_data) for i in range(4)]
    
    for i, (color, val_2g, val_3g, val_4g) in enumerate(legend_data):
        row = legend_start_row + 2 + i
        worksheet.write(row, legend_start_col, color, color_formats[color])
        worksheet.write(row, legend_start_col + 1, val_2g, cell_format)
        worksheet.write(row, legend_start_col + 2, val_3g, cell_format)
        worksheet.write(row, legend_start_col + 3, val_4g, cell_format)
    
    for i, width in enumerate(max_widths):
        worksheet.set_column(legend_start_col + i, legend_start_col + i, width + 2)

def apply_conditional_formatting(writer, sheet_name, df):
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    format_green = workbook.add_format({'bg_color': '#99FF99', 'border': 1})
    format_yellow = workbook.add_format({'bg_color': '#FFFF99', 'border': 1})
    format_red = workbook.add_format({'bg_color': '#FFCCCC', 'border': 1})
    format_gray = workbook.add_format({'bg_color': '#BFBFBF', 'border': 1})
    format_double_border = workbook.add_format({'top': 2})

    num_rows = len(df)
    num_cols = len(df.columns)
    
    for col_idx in range(num_cols):
        max_length = max(df.iloc[:, col_idx].astype(str).apply(len).max(), len(df.columns[col_idx]))
        worksheet.set_column(col_idx, col_idx, max_length + 2)

    for col_idx in range(3, num_cols):
        col_letter = chr(65 + col_idx)
        
        for row in range(2, num_rows + 2):
            technology = df.iloc[row - 2]['Tecnologia']
            
            if 'G' in technology:
                green_min, yellow_min, red_max = -80, -90, -90
            elif 'U' in technology:
                green_min, yellow_min, red_max = -85, -95, -95
            elif 'L' in technology:
                green_min, yellow_min, red_max = -95, -105, -105
            else:
                continue

            worksheet.conditional_format(f'{col_letter}{row}:{col_letter}{row}',
                                         {'type': 'cell', 'criteria': 'between', 'minimum': green_min, 'maximum': -2, 'format': format_green})
            worksheet.conditional_format(f'{col_letter}{row}:{col_letter}{row}',
                                         {'type': 'cell', 'criteria': 'between', 'minimum': yellow_min, 'maximum': green_min, 'format': format_yellow})
            worksheet.conditional_format(f'{col_letter}{row}:{col_letter}{row}',
                                         {'type': 'cell', 'criteria': 'less than', 'value': red_max, 'format': format_red})
            worksheet.conditional_format(f'{col_letter}{row}:{col_letter}{row}',
                                         {'type': 'cell', 'criteria': 'between', 'minimum': -1, 'maximum': 1, 'format': format_gray})
            worksheet.conditional_format(f'{col_letter}{row}:{col_letter}{row}',
                                         {'type': 'blanks', 'format': format_gray})
    
    prev_area = None
    for row in range(1, num_rows + 1):
        current_area = df.iloc[row - 1]['Area Misurata'] if 'Area Misurata' in df.columns else None
        if prev_area is not None and current_area != prev_area:
            worksheet.set_row(row, None, format_double_border, {'first': 0, 'last': 6})
        prev_area = current_area

    add_legend_to_summary_sheet(writer, sheet_name)

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
