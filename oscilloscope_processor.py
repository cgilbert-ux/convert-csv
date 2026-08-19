#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Programme de traitement de traces d'oscilloscope - Interface GUI complète
- Sélection de 1 à 4 fichiers CSV (C1, C2, C3, C4)
- Décimation automatique ou manuelle des données
- Export vers Excel avec graphique
- Sauvegarde du paramètre de décimation dans un fichier INI
- Interface entièrement graphique avec barre de progression
"""

import os
import re
import configparser
from pathlib import Path
import threading

import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference, Series
from openpyxl.styles import Font, Alignment


def get_ini_path():
    """Retourne le chemin du fichier INI"""
    script_dir = Path(__file__).parent.resolve()
    return script_dir / "oscilloscope_config.ini"


def load_decimation_from_ini():
    """Charge la valeur de décimation depuis le fichier INI"""
    ini_path = get_ini_path()
    config = configparser.ConfigParser()
    
    if ini_path.exists():
        config.read(ini_path, encoding='utf-8')
        if 'SETTINGS' in config and 'decimation' in config['SETTINGS']:
            try:
                return int(config['SETTINGS']['decimation'])
            except ValueError:
                pass
    return None


def save_decimation_to_ini(decimation):
    """Sauvegarde la valeur de décimation dans le fichier INI"""
    ini_path = get_ini_path()
    config = configparser.ConfigParser()
    
    if ini_path.exists():
        config.read(ini_path, encoding='utf-8')
    
    if 'SETTINGS' not in config:
        config['SETTINGS'] = {}
    
    config['SETTINGS']['decimation'] = str(decimation)
    
    with open(ini_path, 'w', encoding='utf-8') as f:
        config.write(f)


def extract_channel_info(filepath):
    """Extrait les informations de canal depuis le nom du fichier"""
    filename = os.path.basename(filepath)
    
    match = re.search(r'C([1-4])', filename, re.IGNORECASE)
    if match:
        channel_num = int(match.group(1))
        channel_name = f"Voie {channel_num}"
    else:
        channel_num = None
        channel_name = "Trace"
    
    base_name = re.sub(r'_?C[1-4]', '', filename, flags=re.IGNORECASE)
    base_name = re.sub(r'\.csv$', '', base_name, flags=re.IGNORECASE)
    
    return {
        'channel_num': channel_num,
        'channel_name': channel_name,
        'base_name': base_name
    }


def calculate_optimal_decimation(total_points, min_lines=50000, max_lines=100000):
    """Calcule la décimation optimale pour avoir entre min_lines et max_lines"""
    if total_points <= max_lines:
        return 1
    
    target_lines = (min_lines + max_lines) // 2
    decimation = max(1, total_points // target_lines)
    
    resulting_lines = total_points // decimation
    if resulting_lines < min_lines and decimation > 1:
        decimation -= 1
    
    return max(1, decimation)


def read_csv_file(filepath):
    """
    Lit un fichier CSV d'oscilloscope avec en-têtes multiples.
    Gère les fichiers Siglent SDS1204X avec métadonnées.
    """
    # Lire les premières lignes pour détecter le format
    with open(filepath, 'r', encoding='utf-8') as f:
        lines = []
        for i, line in enumerate(f):
            lines.append(line)
            if i > 30:  # Lire suffisamment de lignes
                break
    
    # Trouver où commencent les données
    data_start_line = 0
    for i, line in enumerate(lines):
        stripped = line.strip()
        # Chercher la ligne d'en-tête avec "Second" ou "Time"
        if 'Second' in line or 'Time' in line:
            data_start_line = i + 1
            break
        # Ou chercher une ligne qui commence par un nombre négatif (données scientifiques)
        if stripped.startswith('-') and ',' in stripped:
            data_start_line = i
            break
    
    # Essayer de lire avec pandas en sautant les lignes d'en-tête
    try:
        df = pd.read_csv(
            filepath,
            skiprows=data_start_line,
            header=None,
            names=['Time', 'Voltage'],
            usecols=[0, 1],
            engine='c'
        )
    except Exception:
        try:
            df = pd.read_csv(
                filepath,
                skiprows=data_start_line,
                header=None,
                names=['Time', 'Voltage'],
                usecols=[0, 1],
                engine='python'
            )
        except Exception:
            # Lecture manuelle en dernier recours
            times = []
            voltages = []
            with open(filepath, 'r', encoding='utf-8') as f:
                for i, line in enumerate(f):
                    if i < data_start_line:
                        continue
                    parts = line.strip().split(',')
                    if len(parts) >= 2:
                        try:
                            t = float(parts[0])
                            v = float(parts[1])
                            times.append(t)
                            voltages.append(v)
                        except ValueError:
                            continue
            df = pd.DataFrame({'Time': times, 'Voltage': voltages})
    
    return df, ',', 'utf-8'


def decimate_data(df, decimation_factor):
    """Réduit le nombre de points par décimation"""
    if decimation_factor <= 1:
        return df
    return df.iloc[::decimation_factor].reset_index(drop=True)


class OscilloscopeApp:
    """Application GUI principale"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Traitement de traces d'oscilloscope")
        self.root.geometry("700x500")
        self.root.resizable(True, True)
        
        self.selected_files = []
        self.decimation_value = tk.IntVar()
        self.processing = False
        
        self.setup_ui()
        
        # Charger la décimation sauvegardée
        saved_dec = load_decimation_from_ini()
        if saved_dec:
            self.decimation_value.set(saved_dec)
    
    def setup_ui(self):
        """Configure l'interface utilisateur"""
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # Titre
        title_label = ttk.Label(main_frame, text="Traitement de traces d'oscilloscope", 
                               font=('Arial', 14, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Section sélection de fichiers
        file_frame = ttk.LabelFrame(main_frame, text="Fichiers CSV", padding="10")
        file_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        ttk.Button(file_frame, text="Sélectionner 1-4 fichiers", 
                  command=self.select_files).grid(row=0, column=0, padx=(0, 10))
        
        self.file_list_text = tk.Text(file_frame, height=6, width=60)
        self.file_list_text.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 0))
        
        self.files_count_label = ttk.Label(file_frame, text="Aucun fichier sélectionné")
        self.files_count_label.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(5, 0))
        
        # Section décimation
        decim_frame = ttk.LabelFrame(main_frame, text="Décimation", padding="10")
        decim_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(decim_frame, text="Facteur de décimation:").grid(row=0, column=0, padx=(0, 10))
        
        self.decim_entry = ttk.Entry(decim_frame, textvariable=self.decimation_value, width=10)
        self.decim_entry.grid(row=0, column=1, sticky=tk.W)
        
        ttk.Button(decim_frame, text="Calcul automatique", 
                  command=self.calc_auto_decimation).grid(row=0, column=2, padx=(10, 0))
        
        self.decim_info_label = ttk.Label(decim_frame, text="")
        self.decim_info_label.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(5, 0))
        
        # Section progression
        progress_frame = ttk.LabelFrame(main_frame, text="Progression", padding="10")
        progress_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='indeterminate', length=400)
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        self.status_label = ttk.Label(progress_frame, text="Prêt")
        self.status_label.grid(row=1, column=0, sticky=tk.W)
        
        # Boutons d'action
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=4, column=0, columnspan=2, pady=(10, 0))
        
        ttk.Button(btn_frame, text="Traiter les fichiers", 
                  command=self.start_processing).grid(row=0, column=0, padx=(0, 10))
        
        ttk.Button(btn_frame, text="Quitter", 
                  command=self.root.quit).grid(row=0, column=1)
    
    def select_files(self):
        """Sélectionne les fichiers CSV"""
        files = filedialog.askopenfilenames(
            title="Sélectionnez 1 à 4 fichiers CSV",
            filetypes=[("Fichiers CSV", "*.csv"), ("Tous les fichiers", "*.*")]
        )
        
        if len(files) > 4:
            messagebox.showerror("Erreur", "Maximum 4 fichiers autorisés")
            return
        
        self.selected_files = list(files)
        
        self.file_list_text.delete(1.0, tk.END)
        for f in self.selected_files:
            self.file_list_text.insert(tk.END, f"  • {os.path.basename(f)}\n")
        
        count = len(self.selected_files)
        self.files_count_label.config(text=f"{count} fichier(s) sélectionné(s)")
        
        # Mettre à jour la décimation automatique
        if count > 0:
            self.calc_auto_decimation()
    
    def calc_auto_decimation(self):
        """Calcule la décimation automatique"""
        if not self.selected_files:
            messagebox.showwarning("Attention", "Sélectionnez d'abord des fichiers")
            return
        
        try:
            df, _, _ = read_csv_file(self.selected_files[0])
            total_points = len(df)
            auto_dec = calculate_optimal_decimation(total_points)
            
            self.decimation_value.set(auto_dec)
            result_lines = total_points // auto_dec
            
            self.decim_info_label.config(
                text=f"Fichier: {total_points:,} points → Décimation: {auto_dec} → {result_lines:,} lignes"
            )
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lire le fichier: {e}")
    
    def start_processing(self):
        """Démarre le traitement dans un thread"""
        if not self.selected_files:
            messagebox.showwarning("Attention", "Sélectionnez des fichiers d'abord")
            return
        
        if self.processing:
            messagebox.showwarning("Attention", "Traitement en cours...")
            return
        
        self.processing = True
        self.progress_bar.config(mode='indeterminate')
        self.progress_bar.start(10)
        
        thread = threading.Thread(target=self.process_files)
        thread.daemon = True
        thread.start()
    
    def process_files(self):
        """Traite les fichiers (dans un thread)"""
        try:
            decimation = self.decimation_value.get()
            if decimation < 1:
                decimation = 1
            
            # Sauvegarder la décimation
            save_decimation_to_ini(decimation)
            
            self.update_status("Lecture des fichiers...")
            
            channel_data = {}
            
            for i, filepath in enumerate(self.selected_files):
                self.update_status(f"Lecture ({i+1}/{len(self.selected_files)}): {os.path.basename(filepath)}")
                
                df, _, _ = read_csv_file(filepath)
                channel_info = extract_channel_info(filepath)
                
                df_decimated = decimate_data(df, decimation)
                
                time_col = df_decimated.columns[0]
                voltage_col = df_decimated.columns[1]
                
                channel_key = f"C{channel_info['channel_num']}" if channel_info['channel_num'] else filepath
                
                channel_data[channel_key] = {
                    'time': df_decimated[time_col].values,
                    'voltage': df_decimated[voltage_col].values,
                    'channel_info': channel_info,
                    'original_file': filepath
                }
            
            if not channel_data:
                self.update_status("Aucune donnée traitée")
                self.processing = False
                self.progress_bar.stop()
                return
            
            # Générer le nom de fichier Excel
            first_ch = list(channel_data.values())[0]
            base_name = first_ch['channel_info']['base_name']
            base_name = re.sub(r'[^\w\-_]', '_', base_name)
            base_name = re.sub(r'_+', '_', base_name).strip('_')
            
            if not base_name:
                base_name = "oscilloscope_data"
            
            output_filename = f"{base_name}_traces.xlsx"
            output_path = os.path.join(os.path.dirname(self.selected_files[0]), output_filename)
            
            self.update_status(f"Création Excel: {output_filename}...")
            
            # Créer le fichier Excel
            self.create_excel(channel_data, output_path)
            
            self.update_status(f"Terminé! Fichier créé: {output_filename}")
            self.progress_bar.stop()
            
            messagebox.showinfo("Succès", f"Fichier Excel créé:\n{output_filename}\n\n{len(channel_data)} trace(s) exportée(s)")
            
        except Exception as e:
            self.update_status(f"Erreur: {e}")
            self.progress_bar.stop()
            messagebox.showerror("Erreur", str(e))
        finally:
            self.processing = False
    
    def create_excel(self, channel_data, output_path):
        """Crée le fichier Excel avec graphique"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Données"
        
        channels_list = list(channel_data.keys())
        first_ch_data = channel_data[channels_list[0]]
        time_data = first_ch_data['time']
        
        # En-têtes
        headers = ["Point"]
        for ch_key in channels_list:
            ch_info = channel_data[ch_key]['channel_info']
            headers.append(f"{ch_info['channel_name']} (V)")
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')
        
        # Écriture des données
        total_rows = len(time_data)
        
        for row_idx in range(total_rows):
            ws.cell(row=row_idx + 2, column=1, value=time_data[row_idx])
            
            for col_idx, ch_key in enumerate(channels_list, 2):
                voltage = channel_data[ch_key]['voltage'][row_idx]
                ws.cell(row=row_idx + 2, column=col_idx, value=voltage)
            
            # Mise à jour de la progression tous les 1000 points
            if row_idx % 1000 == 0:
                pct = int((row_idx / total_rows) * 100)
                self.update_status(f"Écriture Excel: {pct}%")
        
        # Création du graphique
        chart = LineChart()
        chart.title = "Traces d'oscilloscope"
        chart.style = 13
        chart.y_axis.title = 'Tension (V)'
        chart.x_axis.title = 'Point de mesure'
        chart.width = 20
        chart.height = 10
        
        for col_idx in range(2, len(channels_list) + 2):
            values = Reference(ws, min_col=col_idx, min_row=2, max_row=total_rows + 1)
            series_name = ws.cell(row=1, column=col_idx).value
            series = Series(values, from_series=series_name)
            chart.series.append(series)
        
        ws.add_chart(chart, "E2")
        
        wb.save(output_path)
    
    def update_status(self, message):
        """Met à jour le statut (thread-safe)"""
        self.root.after(0, lambda: self.status_label.config(text=message))


def main():
    """Point d'entrée principal"""
    root = tk.Tk()
    app = OscilloscopeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
