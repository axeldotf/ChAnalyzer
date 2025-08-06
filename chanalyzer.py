import pandas as pd
import os
import tkinter as tk
from tkinter import ttk, filedialog, simpledialog, messagebox
from tqdm import tqdm
import threading
from xlsxwriter.utility import xl_col_to_name

# Style constants
MAIN_BG = '#f2f2f2'
ACCENT_ORANGE = '#FFA500'
ACCENT_BLUE = '#003366'

# 5G and legacy technology mappings
MAPPING_5G = {
    643296: ('VF', 'NR3500-643296'), 645312: ('VF', 'NR3500-645312'),
    638016: ('W3', 'NR3500-638016'), 641664: ('Iliad', 'NR3500-641664'),
    636768: ('TIM', 'NR3500-636768'), 648768: ('TIM', 'NR3500-648768'), 650688: ('TIM', 'NR3500-650688')
}

MAPPING_LEGACY = {
    6300: ('TIM', 'L800'), 6400: ('VF', 'L800'), 6200: ('W3', 'L800'),
    1350: ('TIM', 'L1800'), 1500: ('Iliad', 'L1800'), 1650: ('W3', 'L1800'), 1850: ('VF', 'L1800'),
    2900: ('Iliad', 'L2600'), 3025: ('VF', 'L2600'), 3350: ('W3', 'L2600'), 3175: ('TIM', 'L2600'),
    125: ('W3', 'L2100'), 275: ('TIM', 'L2100'), 525: ('VF', 'L2100'), 400: ('Iliad', 'L2100'),
    2938: ('Iliad', 'U900'), 3063: ('W3', 'U900'), 10563: ('W3', 'U2100'), 100: ('W3', 'L2100'),
    9260: ('Iliad', 'L700'), 9360: ('TIM', 'L700'), 9460: ('VF', 'L700')
}

GSM_MAPPING = {
    range(1, 26): ('TIM', 'G900'), range(1000, 1024): ('TIM', 'G900'),
    range(27, 76): ('VF', 'G900'), range(77, 125): ('W3', 'G900')
}

def map_operator_technology(value, is_5g):
    if is_5g:
        return MAPPING_5G.get(value, ('Unknown', 'Unknown'))
    if value in MAPPING_LEGACY:
        return MAPPING_LEGACY[value]
    for key, val in GSM_MAPPING.items():
        if value in key:
            return val
    return ('Unknown', 'Unknown')

# Excel formatting helpers

def add_legend_to_summary_sheet(writer, sheet_name, n_cols):
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    protocols = ['2G', '3G', '4G', '5G']
    legend_row = 2
    # Place legend after data columns
    legend_col = n_cols + 1

    hdr = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': '#D9E1F2', 'border': 1})
    cell_fmt = workbook.add_format({'align': 'center', 'border': 1})
    colors = {
        'Verde': workbook.add_format({'bg_color': '#99FF99', 'border': 1, 'align': 'center'}),
        'Giallo': workbook.add_format({'bg_color': '#FFFF99', 'border': 1, 'align': 'center'}),
        'Rosso': workbook.add_format({'bg_color': '#FFCCCC', 'border': 1, 'align': 'center'}),
        'Grigio': workbook.add_format({'bg_color': '#BFBFBF', 'border': 1, 'align': 'center'})
    }

    # Merge header for legend
    worksheet.merge_range(legend_row, legend_col, legend_row, legend_col + len(protocols),
                          'Legenda Copertura', hdr)
    # Write legend headers
    for i, p in enumerate(['Colore'] + [f'Intervallo {t}' for t in protocols]):
        worksheet.write(legend_row + 1, legend_col + i, p, hdr)

    # Legend data
    data = [
        ('Verde', '>-80 dBm', '>-85 dBm', '>-95 dBm', '>-95 dBm'),
        ('Giallo', '-80÷-90 dBm', '-85÷-95 dBm', '-95÷-105 dBm', '-95÷-105 dBm'),
        ('Rosso', '<-90 dBm', '<-95 dBm', '<-105 dBm', '<-105 dBm'),
        ('Grigio', 'Assente', 'Assente', 'Assente', 'Assente')
    ]
    for r, row in enumerate(data):
        for c, val in enumerate(row):
            fmt = colors[row[0]] if c == 0 else cell_fmt
            worksheet.write(legend_row + 2 + r, legend_col + c, val, fmt)

    # Set column widths for legend
    for i in range(len(data[0])):
        worksheet.set_column(legend_col + i, legend_col + i, 15)


def apply_conditional_formatting(writer, sheet_name, df):
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    fmts = {
        'green': workbook.add_format({'bg_color': '#99FF99', 'border': 1}),
        'yellow': workbook.add_format({'bg_color': '#FFFF99', 'border': 1}),
        'red': workbook.add_format({'bg_color': '#FFCCCC', 'border': 1}),
        'gray': workbook.add_format({'bg_color': '#BFBFBF', 'border': 1}),
        'dborder': workbook.add_format({'top': 2})
    }
    rows, cols = df.shape

    # Apply conditional formatting per-tech
    for col in range(3, cols):
        letter = xl_col_to_name(col)
        for r in range(2, rows + 2):
            tech = df.at[r - 2, 'Tecnologia']
            if tech.startswith('NR'):
                thr = (-95, -105)
            elif tech.startswith('G'):
                thr = (-80, -90)
            elif tech.startswith('U'):
                thr = (-85, -95)
            elif tech.startswith('L'):
                thr = (-95, -105)
            else:
                continue
            worksheet.conditional_format(f'{letter}{r}',
                                         {'type': 'cell', 'criteria': 'between', 'minimum': thr[0], 'maximum': -2, 'format': fmts['green']})
            worksheet.conditional_format(f'{letter}{r}',
                                         {'type': 'cell', 'criteria': 'between', 'minimum': thr[1], 'maximum': thr[0], 'format': fmts['yellow']})
            worksheet.conditional_format(f'{letter}{r}',
                                         {'type': 'cell', 'criteria': 'less than', 'value': thr[1], 'format': fmts['red']})
            worksheet.conditional_format(f'{letter}{r}',
                                         {'type': 'cell', 'criteria': 'between', 'minimum': -1, 'maximum': 1, 'format': fmts['gray']})
            worksheet.conditional_format(f'{letter}{r}',
                                         {'type': 'blanks', 'format': fmts['gray']})

    # Add thicker border between different areas
    prev_area = None
    for r in range(1, rows + 1):
        area = df.at[r - 1, 'Area Misurata']
        if prev_area is not None and area != prev_area:
            worksheet.set_row(r, None, fmts['dborder'], {'first': 0, 'last': cols - 1})
        prev_area = area

    # Add legend, passing number of data columns
    add_legend_to_summary_sheet(writer, sheet_name, df.shape[1])


class ChAnalyzerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('ChAnalyzer')
        self.geometry('800x600')
        self.configure(bg=MAIN_BG)
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('TLabel', background=MAIN_BG, foreground=ACCENT_BLUE)
        style.configure('Footer.TLabel', background=MAIN_BG, foreground=ACCENT_BLUE, font=('Arial', 9, 'italic'))
        style.configure('TButton', background=ACCENT_BLUE, foreground='white')
        style.map('TButton', background=[('active', ACCENT_ORANGE)])
        style.configure('TCheckbutton', background=MAIN_BG, foreground=ACCENT_BLUE)
        self.files = []  # list of (path, sheet_name)
        self._stop_event = threading.Event()
        self._build_gui()

    def _build_gui(self):
        # Settings frame
        frame = ttk.Labelframe(self, text='Impostazioni', padding=10)
        frame.pack(fill='x', padx=10, pady=10)
        ttk.Label(frame, text='File di output:').grid(row=0, column=0, sticky='w')
        self.out_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.out_var).grid(row=0, column=1, sticky='ew', padx=5)
        ttk.Button(frame, text='Seleziona Output', command=self._select_output_file).grid(row=0, column=2)
        self.is_5g = tk.BooleanVar()
        ttk.Checkbutton(frame, text='Includi 5G', variable=self.is_5g).grid(row=1, column=1, sticky='w', pady=5)
        frame.columnconfigure(1, weight=1)

        # Input files frame
        fl_frame = ttk.Labelframe(self, text='File di input', padding=10)
        fl_frame.pack(fill='both', expand=True, padx=10, pady=10)
        self.listbox = tk.Listbox(fl_frame)
        self.listbox.pack(side='left', fill='both', expand=True)
        sb = ttk.Scrollbar(fl_frame, command=self.listbox.yview)
        sb.pack(side='left', fill='y')
        self.listbox.config(yscrollcommand=sb.set)
        btn_f = ttk.Frame(fl_frame)
        btn_f.pack(side='left', fill='y', padx=5)
        ttk.Button(btn_f, text='Aggiungi Piano', command=self._add_file).pack(fill='x', pady=5)
        ttk.Button(btn_f, text='Rimuovi Piano', command=self._remove_file).pack(fill='x')

        # Generate frame
        gen_frame = ttk.Frame(self)
        gen_frame.pack(fill='x', padx=10, pady=10)
        self.gen_btn = ttk.Button(gen_frame, text='Genera Riepilogo', command=self._start)
        self.gen_btn.pack(side='left')
        stop_btn = ttk.Button(gen_frame, text='Interrompi', command=self._stop)
        stop_btn.pack(side='left', padx=5)
        self.progress = ttk.Progressbar(gen_frame, mode='indeterminate')
        self.progress.pack(side='left', fill='x', expand=True, padx=10)

        # Footer
        footer = ttk.Label(self, text='Creato da Alessandro Frullo', style='Footer.TLabel')
        footer.pack(side='bottom', pady=5)

    def _select_output_file(self):
        f = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel', '*.xlsx')])
        if f:
            self.out_var.set(f)

    def _stop(self):
        self._stop_event.set()

    def _add_file(self):
        paths = filedialog.askopenfilenames(filetypes=[('Excel', '*.xlsx;*.xls')])
        for p in paths:
            name = simpledialog.askstring('Nome Foglio', f'Nome foglio per {os.path.basename(p)}:')
            if name:
                self.files.append((p, name))
                self.listbox.insert('end', f'{os.path.basename(p)} -> {name}')

    def _remove_file(self):
        sel = self.listbox.curselection()
        if sel:
            idx = sel[0]
            del self.files[idx]
            self.listbox.delete(idx)

    def _start(self):
        self._stop_event.clear()
        out = self.out_var.get().strip()
        if not out:
            messagebox.showerror('Errore', 'Seleziona file di output')
            return
        if not self.files:
            messagebox.showerror('Errore', 'Aggiungi almeno un file di input')
            return
        self.gen_btn.config(state='disabled')
        self.progress.start(10)
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        summary = []
        for path, sheet in self.files:
            if self._stop_event.is_set():
                break
            df = pd.read_excel(path)
            is5 = self.is_5g.get()
            if is5:
                df = df[['NR-ARFCN', '1. best SS-RSRP']].dropna()
                df['NR-ARFCN'] = df['NR-ARFCN'].astype(int)
                avg = df.groupby('NR-ARFCN')['1. best SS-RSRP'].mean().round(1)
                out_df = pd.DataFrame({'NR-ARFCN': avg.index, 'Misura': avg.values})
                out_df[['Operatore', 'Tecnologia']] = out_df['NR-ARFCN'].apply(
                    lambda x: pd.Series(map_operator_technology(x, True)))
                req_tech = [v[1] for v in MAPPING_5G.values()]
            else:
                df['Group'] = df['Ch'].fillna(df['ARFCN'])
                avg = df.groupby('Group')[['1. best RSRP', '1. best RSCP', '1. best Rx Level']].mean().mean(axis=1).round(1)
                out_df = pd.DataFrame({'Group': avg.index, 'Misura': avg.values})
                out_df[['Operatore', 'Tecnologia']] = out_df['Group'].astype(int).apply(
                    lambda x: pd.Series(map_operator_technology(x, False)))
                req_tech = ['G900', 'L800', 'L1800', 'U900', 'U2100', 'L2100', 'L2600', 'L700']

            final = out_df.groupby(['Operatore', 'Tecnologia']).mean(numeric_only=True).reset_index()
            final['ARFCN'] = final.apply(
                lambda r: out_df[(out_df['Operatore'] == r['Operatore']) & 
                                 (out_df['Tecnologia'] == r['Tecnologia'])]
                                .sort_values('Misura', ascending=False)
                                .iloc[0, 0], axis=1)

            # Fill missing combinations without deprecated concat
            for op in final['Operatore'].unique():
                for tech in req_tech:
                    if not ((final['Operatore'] == op) & (final['Tecnologia'] == tech)).any():
                        idx = final.shape[0]
                        final.loc[idx, 'Operatore'] = op
                        final.loc[idx, 'Tecnologia'] = tech
                        final.loc[idx, 'ARFCN'] = pd.NA
                        final.loc[idx, 'Misura'] = pd.NA

            final['Tecnologia'] = pd.Categorical(final['Tecnologia'], categories=req_tech, ordered=True)
            final = final.sort_values(['Operatore', 'Tecnologia'])

            for tech in req_tech:
                row = {
                    'Area Misurata': sheet,
                    'Tecnologia': tech,
                    'Listino Inwit': '5G' if is5 else (
                        'VOCE' if tech == 'G900' else 'DATI' if tech == 'L800' 
                        else 'DATI PLUS' if tech == 'L1800' else 'Altro')
                }
                for op in ['TIM', 'VF', 'W3', 'Iliad']:
                    val = final[(final['Operatore'] == op) & (final['Tecnologia'] == tech)]['Misura']
                    row[f'Copertura {op}'] = val.values[0] if not val.empty else '/'
                summary.append(row)

        if not self._stop_event.is_set():
            with pd.ExcelWriter(self.out_var.get(), engine='xlsxwriter') as writer:
                summary_df = pd.DataFrame(summary)
                summary_df.to_excel(writer, sheet_name='Riepilogo', index=False)
                apply_conditional_formatting(writer, 'Riepilogo', summary_df)

        self.after(0, self._done)

    def _done(self):
        self.progress.stop()
        self.gen_btn.config(state='normal')
        if self._stop_event.is_set():
            messagebox.showinfo('Interrotto', 'Elaborazione interrotta dall’utente')
        else:
            messagebox.showinfo('Successo', f'File creato: {self.out_var.get()}')

if __name__ == '__main__':
    app = ChAnalyzerApp()
    app.mainloop()
