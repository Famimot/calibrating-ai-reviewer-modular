#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Calibration Script v.4.7 — GUI with Visualization
Processing LLM-Reviewer runs with charts and bookmarks
Universal version - works with any article names

Fixes:
- Eliminated duplicate points in box plot
- Added top margin offset
- Ability to select number of articles
- Increased font sizes
- Wide article counter

Author: Kravtsov G.
Version: 4.7
"""

import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
from matplotlib.patches import Patch
from pathlib import Path
import re
import warnings
from datetime import datetime

warnings.filterwarnings('ignore')

# Font settings for plots - INCREASED
plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.size'] = 13
plt.rcParams['axes.labelsize'] = 14
plt.rcParams['axes.titlesize'] = 16
plt.rcParams['xtick.labelsize'] = 12
plt.rcParams['ytick.labelsize'] = 12

# Configuration
CRITERIA_FULL = {
    'Novelty (30)': ['Novelty (30)', 'Novelty', 'Новизна (30)', 'Новизна'],
    'Methodology (25)': ['Methodology (25)', 'Methodology', 'Методология (25)', 'Методология'],
    'Practical value (20)': ['Practical value (20)', 'Practical value', 'Практ. ценность (20)', 'Практическая ценность (20)', 'Практическая ценность'],
    'Visualization (15)': ['Visualization (15)', 'Visualization', 'Визуализация (15)', 'Визуализация'],
    'Reproducibility (10)': ['Reproducibility (10)', 'Reproducibility', 'Воспроизводимость (10)', 'Воспроизводимость']
}

CRITERIA_STANDARD = ['Novelty (30)', 'Methodology (25)', 'Practical value (20)', 'Visualization (15)',
                     'Reproducibility (10)']
CRITERIA_SHORT = ['Novelty', 'Methodology', 'Practical value', 'Visualization', 'Reproducibility']

# Criterion order for cumulative chart (from smallest to largest weight)
CRITERIA_CUMULATIVE_ORDER = ['Reproducibility (10)', 'Visualization (15)',
                             'Practical value (20)', 'Methodology (25)', 'Novelty (30)']
CRITERIA_CUMULATIVE_SHORT = ['Reproducibility\n(10)', 'Visualization\n(15)',
                             'Practical\nvalue (20)', 'Methodology\n(25)', 'Novelty\n(30)']

TRIM_MIN = 2
TRIM_MAX = 2
MIN_RUNS = 5

CRITERIA_MAX = {
    'Novelty (30)': 30,
    'Methodology (25)': 25,
    'Practical value (20)': 20,
    'Visualization (15)': 15,
    'Reproducibility (10)': 10,
    'Total score': 100
}


def normalize_article_name(article):
    if pd.isna(article):
        return None
    article_str = str(article).strip()
    numbers = re.findall(r'\d+', article_str)
    if not numbers:
        return article_str
    number = numbers[0]
    article_lower = article_str.lower()
    if re.match(r'^[a-zа-я]+\d+$', article_lower):
        prefix = re.match(r'^([a-zа-я]+)', article_lower).group(1)
        return f"{prefix.upper()}{number}"
    elif 'ст' in article_lower or 'article' in article_lower:
        return f"Art.{number}"
    elif article_str.isdigit():
        return f"Art.{number}"
    else:
        match = re.match(r'^([A-Za-zА-Яа-я]+)', article_str)
        if match:
            prefix = match.group(1)
            return f"{prefix}{number}"
        else:
            return f"Art.{number}"


def find_criterion_column(df_columns):
    column_mapping = {}
    for standard_crit, patterns in CRITERIA_FULL.items():
        for col in df_columns:
            col_str = str(col).strip()
            for pattern in patterns:
                if pattern in col_str or col_str == pattern:
                    column_mapping[standard_crit] = col
                    break
            if standard_crit in column_mapping:
                break
        if standard_crit not in column_mapping:
            base_name = standard_crit.split(' (')[0]
            for col in df_columns:
                col_str = str(col).strip()
                if base_name in col_str:
                    column_mapping[standard_crit] = col
                    break
    return column_mapping


class ReviewAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Calibration Script v.4.7 — LLM-Reviewer")
        self.root.geometry("1500x900")

        self.file_path = None
        self.all_runs = []
        self.final_results = {}
        self.raw_counts = {}
        self.sorted_articles_by_total = []
        self.variability_stats = {}

        self.setup_ui()

    def setup_ui(self):
        # Top panel
        top_frame = tk.Frame(self.root)
        top_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(top_frame, text="Excel file with runs:", font=('Arial', 11)).pack(side=tk.LEFT, padx=5)
        self.file_label = tk.Label(top_frame, text="No file selected", fg="red", width=50, anchor=tk.W)
        self.file_label.pack(side=tk.LEFT, padx=5)

        tk.Button(top_frame, text="Select File", command=self.load_file).pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="Process", command=self.process_data, bg="#4CAF50", fg="white").pack(side=tk.LEFT,
                                                                                                          padx=5)

        self.save_button = tk.Button(top_frame, text="💾 Save Results to Excel",
                                     command=self.save_to_excel, bg="#2196F3", fg="white",
                                     state=tk.DISABLED)
        self.save_button.pack(side=tk.LEFT, padx=5)

        self.status_label = tk.Label(self.root, text="Ready", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.pack(fill=tk.X, padx=10, pady=5)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Tabs
        self.results_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.results_frame, text="📊 Final Results")
        self.setup_results_tab()

        self.stats_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.stats_frame, text="📈 Run Statistics")
        self.setup_stats_tab()

        # Variability tab
        self.variability_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.variability_frame, text="📊 Variability (Cleaned)")
        self.setup_variability_tab()

        self.charts_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.charts_frame, text="📉 Charts")
        self.setup_charts_tab()

        self.about_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.about_frame, text="ℹ️ About")
        self.setup_about_tab()

    def setup_results_tab(self):
        columns = (
            'Article', 'Novelty', 'Methodology', 'Practical value', 'Visualization', 'Reproducibility', 'Total score')
        self.results_tree = ttk.Treeview(self.results_frame, columns=columns, show='headings', height=15)

        for col in columns:
            self.results_tree.heading(col, text=col)
            self.results_tree.column(col, width=100, anchor='center')

        scrollbar = ttk.Scrollbar(self.results_frame, orient=tk.VERTICAL, command=self.results_tree.yview)
        self.results_tree.configure(yscrollcommand=scrollbar.set)

        self.results_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.results_info = tk.Text(self.results_frame, height=8, wrap=tk.WORD)
        self.results_info.pack(fill=tk.X, pady=5)

    def setup_stats_tab(self):
        columns = ('Article', 'Criterion', 'Runs', 'Min', 'Max', 'Mean', 'Std Dev')
        self.stats_tree = ttk.Treeview(self.stats_frame, columns=columns, show='headings', height=20)

        for col in columns:
            self.stats_tree.heading(col, text=col)
            self.stats_tree.column(col, width=120, anchor='center')

        scrollbar = ttk.Scrollbar(self.stats_frame, orient=tk.VERTICAL, command=self.stats_tree.yview)
        self.stats_tree.configure(yscrollcommand=scrollbar.set)

        self.stats_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def setup_variability_tab(self):
        """Variability table on cleaned data"""
        header_frame = tk.Frame(self.variability_frame)
        header_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Label(header_frame, text="Variability Table of Article Total Scores Across Runs",
                 font=('Arial', 14, 'bold')).pack()
        tk.Label(header_frame, text=f"(based on cleaned data — removed {TRIM_MIN} min + {TRIM_MAX} max)",
                 font=('Arial', 11), fg='green').pack()

        columns = ('Article', 'Mean', 'Median', 'Min', 'Max', 'Range', 'Std', 'CV (%)')
        self.variability_tree = ttk.Treeview(self.variability_frame, columns=columns, show='headings', height=15)

        column_widths = {'Article': 120, 'Mean': 100, 'Median': 100, 'Min': 80,
                         'Max': 80, 'Range': 80, 'Std': 100, 'CV (%)': 100}

        for col in columns:
            self.variability_tree.heading(col, text=col)
            self.variability_tree.column(col, width=column_widths.get(col, 100), anchor='center')

        scrollbar = ttk.Scrollbar(self.variability_frame, orient=tk.VERTICAL, command=self.variability_tree.yview)
        self.variability_tree.configure(yscrollcommand=scrollbar.set)

        self.variability_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=5)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=5)

        self.variability_info = tk.Text(self.variability_frame, height=6, wrap=tk.WORD, font=('Arial', 10))
        self.variability_info.pack(fill=tk.X, padx=10, pady=5)

    def setup_charts_tab(self):
        btn_frame = tk.Frame(self.charts_frame, bg='#f0f0f0', height=60)
        btn_frame.pack(fill=tk.X, pady=5, padx=10)
        btn_frame.pack_propagate(False)

        btn_style = {'bg': '#4CAF50', 'fg': 'white', 'font': ('Arial', 11, 'bold'),
                     'padx': 15, 'pady': 8, 'relief': tk.RAISED, 'bd': 2}

        tk.Button(btn_frame, text="📊 Box Plot (score distribution by article)", command=self.plot_boxplots, **btn_style).pack(
            side=tk.LEFT, padx=8)
        tk.Button(btn_frame, text="🔥 Heatmap (article × run)", command=self.plot_heatmap, **btn_style).pack(
            side=tk.LEFT, padx=8)
        tk.Button(btn_frame, text="🎯 Radar Chart (article profile)", command=self.plot_radar, **btn_style).pack(
            side=tk.LEFT, padx=8)
        tk.Button(btn_frame, text="📈 Cumulative Quality Profile", command=self.plot_quality_profile,
                  **btn_style).pack(side=tk.LEFT, padx=8)

        select_frame = tk.Frame(self.charts_frame, bg='#e0e0e0')
        select_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(select_frame, text="Criterion for box plot:", font=('Arial', 11), bg='#e0e0e0').pack(
            side=tk.LEFT, padx=10)

        self.selected_criterion = tk.StringVar(value="Total score")
        criteria_options = CRITERIA_SHORT + ['Total score']

        self.criterion_menu = ttk.Combobox(select_frame, textvariable=self.selected_criterion,
                                           values=criteria_options, state='readonly', width=20)
        self.criterion_menu.pack(side=tk.LEFT, padx=10)
        self.criterion_menu.bind('<<ComboboxSelected>>', lambda e: self.plot_boxplots())

        # Display modes
        tk.Label(select_frame, text="Display mode:", font=('Arial', 11), bg='#e0e0e0').pack(side=tk.LEFT, padx=20)

        self.display_mode = tk.StringVar(value="Cleaned")
        tk.Radiobutton(select_frame, text="Raw (all runs)", variable=self.display_mode,
                       value="Raw", bg='#e0e0e0', font=('Arial', 10), command=self.plot_boxplots).pack(side=tk.LEFT,
                                                                                                       padx=5)
        tk.Radiobutton(select_frame, text="Cleaned (after removing 2 min+2 max)", variable=self.display_mode,
                       value="Cleaned", bg='#e0e0e0', font=('Arial', 10), command=self.plot_boxplots).pack(side=tk.LEFT,
                                                                                                           padx=5)

        self.canvas_frame = tk.Frame(self.charts_frame)
        self.canvas_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    def setup_about_tab(self):
        about_text = tk.Text(self.about_frame, wrap=tk.WORD, font=("Arial", 11))
        about_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        about_text.insert(tk.END, """
CALIBRATION SCRIPT v.4.7
================================

Purpose:
    Processing of multiple LLM-Reviewer runs
    with outlier removal and data visualization.

FEATURES:
    ✅ Box plot with score distribution
    ✅ Heatmap (article × run)
    ✅ Radar chart of article profiles
    ✅ Cumulative quality profile (score accumulation)
    ✅ Variability table on cleaned data
    ✅ Export results to Excel

FIXES:
    ✅ No duplicate points in box plot
    ✅ Points not clipped at boundaries
    ✅ Ability to select number of articles
    ✅ Increased font sizes

Author: Kravtsov G.
Version: 4.7
        """)
        about_text.config(state=tk.DISABLED)

    def load_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel file with runs",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.file_path = file_path
            self.file_label.config(text=Path(file_path).name, fg="green")
            self.status_label.config(text=f"File loaded: {Path(file_path).name}")

    def parse_sheet(self, df, sheet_name):
        try:
            if df.empty:
                return None

            if 'Article' not in df.columns:
                for idx, row in df.iterrows():
                    if 'Article' in str(row.values) or 'Статья' in str(row.values):
                        df.columns = row
                        df = df[idx + 1:].reset_index(drop=True)
                        break

            if 'Article' not in df.columns:
                for col in df.columns:
                    if 'article' in str(col).lower() or 'статья' in str(col).lower():
                        df.rename(columns={col: 'Article'}, inplace=True)
                        break

            if 'Article' not in df.columns:
                return None

            column_mapping = find_criterion_column(df.columns)

            run_data = {}
            for idx, row in df.iterrows():
                article = str(row['Article']).strip()
                if pd.isna(article) or article == 'nan' or article == '':
                    continue

                article_normalized = normalize_article_name(article)

                run_data[article_normalized] = {}
                for standard_crit, actual_col in column_mapping.items():
                    if actual_col in df.columns:
                        val = row[actual_col]
                        if pd.notna(val):
                            try:
                                run_data[article_normalized][standard_crit] = float(val)
                            except:
                                run_data[article_normalized][standard_crit] = None
                        else:
                            run_data[article_normalized][standard_crit] = None

            return run_data if run_data else None
        except Exception as e:
            print(f"Error: {e}")
            return None

    def trim_mean(self, values):
        valid = [v for v in values if v is not None and not np.isnan(v)]
        if len(valid) < TRIM_MIN + TRIM_MAX + 1:
            return np.mean(valid) if valid else None
        valid.sort()
        trimmed = valid[TRIM_MIN:-TRIM_MAX] if TRIM_MAX > 0 else valid[TRIM_MIN:]
        return np.mean(trimmed) if trimmed else None

    def get_cleaned_values(self, values):
        valid = [v for v in values if v is not None and not np.isnan(v)]
        if len(valid) < TRIM_MIN + TRIM_MAX + 1:
            return valid
        valid.sort()
        return valid[TRIM_MIN:-TRIM_MAX] if TRIM_MAX > 0 else valid[TRIM_MIN:]

    def calculate_variability_stats_cleaned(self):
        """Calculate variability statistics on CLEANED data"""
        self.variability_stats = {}
        for article in self.sorted_articles_by_total:
            totals = []
            for run in self.all_runs:
                if article in run:
                    total = sum([run[article].get(crit, 0) for crit in CRITERIA_STANDARD])
                    totals.append(total)

            if totals:
                cleaned_totals = self.get_cleaned_values(totals)
                if cleaned_totals:
                    mean_val = np.mean(cleaned_totals)
                    median_val = np.median(cleaned_totals)
                    min_val = np.min(cleaned_totals)
                    max_val = np.max(cleaned_totals)
                    range_val = max_val - min_val
                    std_val = np.std(cleaned_totals)
                    cv_val = (std_val / mean_val * 100) if mean_val != 0 else 0

                    self.variability_stats[article] = {
                        'Mean': mean_val, 'Median': median_val, 'Min': min_val,
                        'Max': max_val, 'Range': range_val, 'Std': std_val,
                        'CV (%)': cv_val, 'N_cleaned': len(cleaned_totals)
                    }

    def sort_articles_by_criterion(self, criterion):
        if not self.final_results:
            return []
        articles = list(self.final_results.keys())
        if criterion == 'Total score':
            return sorted(articles, key=lambda x: self.final_results[x]['Final score'] if not np.isnan(
                self.final_results[x]['Final score']) else -1, reverse=True)
        else:
            full_crit = None
            for crit in CRITERIA_STANDARD:
                if criterion in crit:
                    full_crit = crit
                    break
            if full_crit:
                return sorted(articles, key=lambda x: self.final_results[x].get(full_crit, 0) if not np.isnan(
                    self.final_results[x].get(full_crit, np.nan)) else -1, reverse=True)
            return articles

    def process_data(self):
        if not self.file_path:
            messagebox.showwarning("Warning", "Please select a file first!")
            return

        self.status_label.config(text="Processing data...")
        self.root.update()

        try:
            excel_file = pd.ExcelFile(self.file_path)
            self.all_runs = []

            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                run_data = self.parse_sheet(df, sheet_name)
                if run_data:
                    self.all_runs.append(run_data)

            if not self.all_runs:
                messagebox.showerror("Error", "No data found for processing!")
                self.status_label.config(text="Error: no data")
                return

            all_articles = set()
            for run in self.all_runs:
                all_articles.update(run.keys())

            results = {article: {crit: [] for crit in CRITERIA_STANDARD} for article in all_articles}

            for run in self.all_runs:
                for article, values in run.items():
                    if article in results:
                        for crit in CRITERIA_STANDARD:
                            if values.get(crit):
                                results[article][crit].append(values[crit])

            self.final_results = {}
            for article, criteria_values in results.items():
                self.final_results[article] = {}
                total = 0
                for crit in CRITERIA_STANDARD:
                    trimmed = self.trim_mean(criteria_values[crit])
                    self.final_results[article][crit] = trimmed if trimmed is not None else np.nan
                    if trimmed is not None:
                        total += trimmed
                self.final_results[article]['Final score'] = total

            self.raw_counts = results
            self.sorted_articles_by_total = self.sort_articles_by_criterion('Total score')

            self.update_results_tab()
            self.update_stats_tab()
            self.update_variability_tab()

            self.save_button.config(state=tk.NORMAL)

            self.status_label.config(
                text=f"Processing complete! Runs: {len(self.all_runs)}, Articles: {len(self.final_results)}")
            messagebox.showinfo("Done", f"Processed {len(self.all_runs)} runs, {len(self.final_results)} articles")

        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status_label.config(text=f"Error: {str(e)}")

    def save_to_excel(self):
        if not self.final_results:
            messagebox.showwarning("Warning", "No data to save!")
            return

        default_filename = f"calibration_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path = filedialog.asksaveasfilename(
            title="Save results to Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=default_filename
        )

        if not file_path:
            return

        try:
            self.status_label.config(text="Saving results to Excel...")
            self.root.update()

            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # Final results
                results_data = []
                for article in self.sorted_articles_by_total:
                    row = {
                        'Article': article,
                        'Novelty (30)': round(self.final_results[article].get('Novelty (30)', np.nan),
                                              2) if not np.isnan(
                            self.final_results[article].get('Novelty (30)', np.nan)) else None,
                        'Methodology (25)': round(self.final_results[article].get('Methodology (25)', np.nan),
                                                  2) if not np.isnan(
                            self.final_results[article].get('Methodology (25)', np.nan)) else None,
                        'Practical value (20)': round(self.final_results[article].get('Practical value (20)', np.nan),
                                                      2) if not np.isnan(
                            self.final_results[article].get('Practical value (20)', np.nan)) else None,
                        'Visualization (15)': round(self.final_results[article].get('Visualization (15)', np.nan),
                                                   2) if not np.isnan(
                            self.final_results[article].get('Visualization (15)', np.nan)) else None,
                        'Reproducibility (10)': round(
                            self.final_results[article].get('Reproducibility (10)', np.nan), 2) if not np.isnan(
                            self.final_results[article].get('Reproducibility (10)', np.nan)) else None,
                        'Total score (100)': round(self.final_results[article]['Final score'], 2)
                    }
                    results_data.append(row)

                df_results = pd.DataFrame(results_data)
                df_results.to_excel(writer, sheet_name='Final_Results', index=False)

                # Variability
                self.calculate_variability_stats_cleaned()
                variability_data = []
                for article in self.sorted_articles_by_total:
                    if article in self.variability_stats:
                        stats = self.variability_stats[article]
                        variability_data.append({
                            'Article': article,
                            'Mean': round(stats['Mean'], 2),
                            'Median': round(stats['Median'], 2),
                            'Min': round(stats['Min'], 2),
                            'Max': round(stats['Max'], 2),
                            'Range': round(stats['Range'], 2),
                            'Std': round(stats['Std'], 2),
                            'CV (%)': round(stats['CV (%)'], 2)
                        })

                df_variability = pd.DataFrame(variability_data)
                df_variability.to_excel(writer, sheet_name='Variability_Cleaned', index=False)

                # Information
                info_data = {
                    'Parameter': ['Source file', 'Processing date', 'Total runs', 'Total articles',
                                 'Removed min', 'Removed max'],
                    'Value': [Path(self.file_path).name, datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                                 len(self.all_runs), len(self.final_results), TRIM_MIN, TRIM_MAX]
                }
                df_info = pd.DataFrame(info_data)
                df_info.to_excel(writer, sheet_name='Information', index=False)

            self.status_label.config(text=f"Results saved: {Path(file_path).name}")
            messagebox.showinfo("Success", f"Results saved to file:\n{file_path}")

        except Exception as e:
            messagebox.showerror("Error saving", str(e))

    def update_results_tab(self):
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)

        for article in self.sorted_articles_by_total:
            prakt_value = self.final_results[article].get('Practical value (20)', np.nan)
            values = [
                article,
                f"{self.final_results[article].get('Novelty (30)', np.nan):.1f}" if not np.isnan(
                    self.final_results[article].get('Novelty (30)', np.nan)) else "N/A",
                f"{self.final_results[article].get('Methodology (25)', np.nan):.1f}" if not np.isnan(
                    self.final_results[article].get('Methodology (25)', np.nan)) else "N/A",
                f"{prakt_value:.1f}" if not np.isnan(prakt_value) else "N/A",
                f"{self.final_results[article].get('Visualization (15)', np.nan):.1f}" if not np.isnan(
                    self.final_results[article].get('Visualization (15)', np.nan)) else "N/A",
                f"{self.final_results[article].get('Reproducibility (10)', np.nan):.1f}" if not np.isnan(
                    self.final_results[article].get('Reproducibility (10)', np.nan)) else "N/A",
                f"{self.final_results[article]['Final score']:.1f}"
            ]
            self.results_tree.insert('', tk.END, values=values)

    def update_stats_tab(self):
        for item in self.stats_tree.get_children():
            self.stats_tree.delete(item)

        for article in self.sorted_articles_by_total:
            for crit in CRITERIA_STANDARD:
                vals = [v for v in self.raw_counts[article][crit] if v is not None]
                if vals:
                    self.stats_tree.insert('', tk.END, values=[
                        article, crit, len(vals),
                        f"{min(vals):.1f}", f"{max(vals):.1f}",
                        f"{np.mean(vals):.1f}", f"{np.std(vals):.2f}"
                    ])

    def update_variability_tab(self):
        for item in self.variability_tree.get_children():
            self.variability_tree.delete(item)

        self.calculate_variability_stats_cleaned()

        for article in self.sorted_articles_by_total:
            if article in self.variability_stats:
                stats = self.variability_stats[article]
                self.variability_tree.insert('', tk.END, values=[
                    article, f"{stats['Mean']:.1f}", f"{stats['Median']:.1f}",
                    f"{stats['Min']:.0f}", f"{stats['Max']:.0f}", f"{stats['Range']:.0f}",
                    f"{stats['Std']:.2f}", f"{stats['CV (%)']:.1f}%"
                ])

        if self.variability_stats:
            avg_cv = np.mean([stats['CV (%)'] for stats in self.variability_stats.values()])
            info_text = f"📊 Average coefficient of variation: {avg_cv:.1f}%\n💡 CV < 10% - low variability, CV > 20% - high variability"
            self.variability_info.delete(1.0, tk.END)
            self.variability_info.insert(tk.END, info_text)

    def get_max_score_for_criterion(self, criterion):
        if criterion == 'Total score':
            return 100
        elif criterion == 'Novelty':
            return 30
        elif criterion == 'Methodology':
            return 25
        elif criterion == 'Practical value':
            return 20
        elif criterion == 'Visualization':
            return 15
        elif criterion == 'Reproducibility':
            return 10
        else:
            return 30

    def plot_boxplots(self):
        self.clear_canvas()
        criterion = self.selected_criterion.get()
        mode = self.display_mode.get()

        if mode == "Raw":
            runs_count = len(self.all_runs)
        else:
            runs_count = len(self.all_runs) - TRIM_MIN - TRIM_MAX

        fig = Figure(figsize=(12, 8))
        ax = fig.add_subplot(111)

        articles = self.sort_articles_by_criterion(criterion)
        data = []

        if criterion == 'Total score':
            for article in articles:
                totals = []
                for run in self.all_runs:
                    if article in run:
                        total = sum([run[article].get(crit, 0) for crit in CRITERIA_STANDARD])
                        totals.append(total)
                if mode == "Cleaned":
                    totals = self.get_cleaned_values(totals)
                data.append(totals)
        else:
            full_crit = None
            for crit in CRITERIA_STANDARD:
                if criterion in crit:
                    full_crit = crit
                    break
            if full_crit:
                for article in articles:
                    vals = self.raw_counts[article].get(full_crit, [])
                    vals = [v for v in vals if v is not None]
                    if mode == "Cleaned":
                        vals = self.get_cleaned_values(vals)
                    data.append(vals)

        if not any(data):
            return

        colors = plt.cm.RdYlGn(np.linspace(0, 1, len(articles)))[::-1]

        # showfliers=False - removes duplicate points
        bp = ax.boxplot(data, labels=articles, showfliers=False, patch_artist=True, widths=0.6)

        for box, color in zip(bp['boxes'], colors):
            box.set_facecolor(color)
            box.set_edgecolor('black')
            box.set_linewidth(2)

        for i, (article, color) in enumerate(zip(articles, colors)):
            if criterion == 'Total score':
                totals = []
                for run in self.all_runs:
                    if article in run:
                        total = sum([run[article].get(crit, 0) for crit in CRITERIA_STANDARD])
                        totals.append(total)
                y_vals = self.get_cleaned_values(totals) if mode == "Cleaned" else totals
            else:
                full_crit = None
                for crit in CRITERIA_STANDARD:
                    if criterion in crit:
                        full_crit = crit
                        break
                y_vals = [v for v in self.raw_counts[article].get(full_crit, []) if v is not None]
                if mode == "Cleaned":
                    y_vals = self.get_cleaned_values(y_vals)
            if y_vals:
                x_vals = np.random.normal(i + 1, 0.08, size=len(y_vals))
                ax.scatter(x_vals, y_vals, color=color, alpha=0.7, s=80, zorder=5, marker='o', edgecolor='black',
                           linewidth=1)

        medians = [np.median(d) if d else 0 for d in data]
        # Reduced median marker size from 200 to 80
        ax.scatter(range(1, len(articles) + 1), medians, color='gold', s=80, marker='D', zorder=10, edgecolor='black',
                   linewidth=1.5, label='Median')

        # Add horizontal lines for Total score (50, 70, 90)
        if criterion == 'Total score':
            ax.axhline(y=90, color='green', linestyle='--', alpha=0.5, linewidth=1.5, label='Excellent (90)')
            ax.axhline(y=70, color='orange', linestyle='--', alpha=0.5, linewidth=1.5, label='Good (70)')
            ax.axhline(y=50, color='red', linestyle='--', alpha=0.5, linewidth=1.5, label='Satisfactory (50)')
        elif criterion == 'Novelty':
            ax.axhline(y=27, color='green', linestyle='--', alpha=0.5, linewidth=1.5, label='90% (27)')
            ax.axhline(y=21, color='orange', linestyle='--', alpha=0.5, linewidth=1.5, label='70% (21)')
            ax.axhline(y=15, color='red', linestyle='--', alpha=0.5, linewidth=1.5, label='50% (15)')
        elif criterion == 'Methodology':
            ax.axhline(y=22.5, color='green', linestyle='--', alpha=0.5, linewidth=1.5, label='90% (22.5)')
            ax.axhline(y=17.5, color='orange', linestyle='--', alpha=0.5, linewidth=1.5, label='70% (17.5)')
            ax.axhline(y=12.5, color='red', linestyle='--', alpha=0.5, linewidth=1.5, label='50% (12.5)')
        elif criterion == 'Practical value':
            ax.axhline(y=18, color='green', linestyle='--', alpha=0.5, linewidth=1.5, label='90% (18)')
            ax.axhline(y=14, color='orange', linestyle='--', alpha=0.5, linewidth=1.5, label='70% (14)')
            ax.axhline(y=10, color='red', linestyle='--', alpha=0.5, linewidth=1.5, label='50% (10)')
        elif criterion == 'Visualization':
            ax.axhline(y=13.5, color='green', linestyle='--', alpha=0.5, linewidth=1.5, label='90% (13.5)')
            ax.axhline(y=10.5, color='orange', linestyle='--', alpha=0.5, linewidth=1.5, label='70% (10.5)')
            ax.axhline(y=7.5, color='red', linestyle='--', alpha=0.5, linewidth=1.5, label='50% (7.5)')
        elif criterion == 'Reproducibility':
            ax.axhline(y=9, color='green', linestyle='--', alpha=0.5, linewidth=1.5, label='90% (9)')
            ax.axhline(y=7, color='orange', linestyle='--', alpha=0.5, linewidth=1.5, label='70% (7)')
            ax.axhline(y=5, color='red', linestyle='--', alpha=0.5, linewidth=1.5, label='50% (5)')

        # Add 5% top margin
        max_score = self.get_max_score_for_criterion(criterion)
        y_max = max_score * 1.05
        ax.set_ylim(0, y_max)

        ax.set_title(f'Distribution of "{criterion}" by Article', fontsize=16, fontweight='bold')
        ax.set_ylabel(f'Score (max {max_score})', fontsize=14)
        ax.set_xlabel('Articles', fontsize=14)
        ax.set_xticklabels(articles, rotation=0, ha='center', fontsize=11)
        ax.grid(axis='y', linestyle='--', alpha=0.3)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

        # Legend in lower left corner
        if criterion in ['Total score', 'Novelty', 'Methodology', 'Practical value', 'Visualization',
                         'Reproducibility']:
            ax.legend(loc='lower left', fontsize=9, bbox_to_anchor=(0, 0))

        self.draw_canvas(fig)

    def plot_heatmap(self):
        self.clear_canvas()
        articles = self.sort_articles_by_criterion('Total score')
        runs_count = len(self.all_runs)

        data_matrix = np.zeros((runs_count, len(articles)))
        for i, run in enumerate(self.all_runs):
            for j, article in enumerate(articles):
                if article in run:
                    total = sum([run[article].get(crit, 0) for crit in CRITERIA_STANDARD])
                    data_matrix[i, j] = total
                else:
                    data_matrix[i, j] = np.nan

        fig = Figure(figsize=(max(10, len(articles) * 0.8), max(8, runs_count * 0.5)))
        ax = fig.add_subplot(111)
        im = ax.imshow(data_matrix, cmap='RdYlGn', aspect='auto', vmin=30, vmax=100)

        ax.set_xticks(np.arange(len(articles)))
        ax.set_yticks(np.arange(runs_count))
        # X-axis labels horizontal (rotation=0)
        ax.set_xticklabels(articles, fontsize=11, rotation=0, ha='center')
        ax.set_yticklabels([f'{i + 1}' for i in range(runs_count)], fontsize=10)
        ax.set_xlabel('Articles', fontsize=13)
        ax.set_ylabel('Run', fontsize=13)
        ax.set_title('Heatmap of Total Scores', fontsize=14)

        for i in range(runs_count):
            for j in range(len(articles)):
                score = data_matrix[i, j]
                if not np.isnan(score):
                    text_color = 'white' if score < 55 else 'black'
                    ax.text(j, i, f'{score:.0f}', ha='center', va='center', color=text_color, fontsize=9)

        cbar = fig.colorbar(im, ax=ax, shrink=0.7)
        cbar.set_label('Total Score', fontsize=12)
        self.draw_canvas(fig)

    def plot_radar(self):
        self.clear_canvas()
        fig = Figure(figsize=(10, 8))
        ax = fig.add_subplot(111, projection='polar')

        angles = np.linspace(0, 2 * np.pi, len(CRITERIA_SHORT), endpoint=False).tolist()
        angles += angles[:1]

        # Use all articles
        top_articles = self.sorted_articles_by_total
        colors = plt.cm.tab20(np.linspace(0, 1, len(top_articles)))

        for idx, (article, color) in enumerate(zip(top_articles, colors)):
            values = [self.final_results[article].get(crit, 0) for crit in CRITERIA_STANDARD]
            values_max = []
            for crit, v in zip(CRITERIA_STANDARD, values):
                if 'Novelty' in crit:
                    values_max.append(v / 30)
                elif 'Methodology' in crit:
                    values_max.append(v / 25)
                elif 'Practical' in crit:
                    values_max.append(v / 20)
                elif 'Visualization' in crit:
                    values_max.append(v / 15)
                else:
                    values_max.append(v / 10)
            values_max += values_max[:1]
            ax.plot(angles, values_max, 'o-', linewidth=2, label=article, color=color)
            ax.fill(angles, values_max, alpha=0.15, color=color)

        ax.set_xticks(angles[:-1])
        ax.set_xticklabels(CRITERIA_SHORT, fontsize=11)
        ax.set_ylim(0, 1)
        ax.set_title(f'Article Profile (all {len(top_articles)} articles)', fontsize=14)
        ax.legend(loc='upper right', bbox_to_anchor=(1.3, 1.0), fontsize=9)
        self.draw_canvas(fig)

    def plot_quality_profile(self):
        """Cumulative quality profile - score accumulation from 0 to 100"""
        self.clear_canvas()
        # Use all articles
        top_articles = self.sorted_articles_by_total

        fig = Figure(figsize=(12, 7))
        ax = fig.add_subplot(111)
        colors = plt.cm.tab20(np.linspace(0, 1, len(top_articles)))

        for idx, (article, color) in enumerate(zip(top_articles, colors)):
            cumulative = 0
            x_points = []
            y_points = []

            for crit in CRITERIA_CUMULATIVE_ORDER:
                score = self.final_results[article].get(crit, 0)
                if np.isnan(score):
                    score = 0
                cumulative += score
                x_points.append(crit)
                y_points.append(cumulative)

            x_pos = range(1, len(x_points) + 1)
            ax.plot(x_pos, y_points, 'o-', linewidth=2.5, markersize=6, label=article, color=color, alpha=0.8)

        ax.set_xlim(0.5, len(CRITERIA_CUMULATIVE_ORDER) + 0.5)
        ax.set_ylim(0, 105)
        ax.set_xticks(range(1, len(CRITERIA_CUMULATIVE_SHORT) + 1))
        # Legend in upper left corner
        if len(top_articles) > 6:
            ncol = 2
        else:
            ncol = 1
        ax.legend(loc='upper left', fontsize=9, ncol=ncol)

        ax.set_xticklabels(CRITERIA_CUMULATIVE_SHORT, rotation=0, ha='center', fontsize=10)
        ax.set_yticks(np.arange(0, 101, 10))
        ax.set_yticklabels([f'{i}' for i in range(0, 101, 10)], fontsize=11)

        ax.axhline(y=90, color='green', linestyle='--', alpha=0.5, linewidth=1)
        ax.axhline(y=70, color='orange', linestyle='--', alpha=0.5, linewidth=1)
        ax.axhline(y=50, color='red', linestyle='--', alpha=0.5, linewidth=1)

        ax.set_xlabel('Evaluation Criteria (from smallest to largest weight)', fontsize=13, fontweight='bold')
        ax.set_ylabel('Accumulated Score', fontsize=13, fontweight='bold')
        ax.set_title(f'Cumulative Quality Profile of Articles (all {len(top_articles)} articles)', fontsize=14,
                     fontweight='bold')

        ax.grid(axis='y', linestyle='--', alpha=0.3)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

        plt.tight_layout()
        self.draw_canvas(fig)

    def clear_canvas(self):
        for widget in self.canvas_frame.winfo_children():
            widget.destroy()

    def draw_canvas(self, fig):
        canvas = FigureCanvasTkAgg(fig, master=self.canvas_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        toolbar = NavigationToolbar2Tk(canvas, self.canvas_frame)
        toolbar.update()


def main():
    root = tk.Tk()
    app = ReviewAnalyzerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()