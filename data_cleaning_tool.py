#############################################################
# Entwickler: Christopher Haase                             #
# Kurs: Software Developer (IHK) [xxxx]                     #
# Erstellungsdatum: 05.09.2024                              #
# Letzte Änderung: 18.11.2024                               #
# Version: 1.0                                              #
# --------------------------------------------------------- #
# Projektarbeit: Data Cleaning Tool                         #
# Beschreibung: Entwicklung einer GUI basierten Daten-      #
# bereinigungs- und Transformationspipeline in Python       #
# --------------------------------------------------------- #
# Kontakt: me@home.com                                      #
#############################################################

import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import subprocess
import webbrowser
import matplotlib.pyplot as plt
from openpyxl.drawing.image import Image
import matplotlib.ticker as ticker
import tempfile
import locale

# Konfiguration als Konstanten
WINDOW_TITLE = "Data Cleaning Tool"
WINDOW_SIZE = "520x670"
LABEL_FONT = "Helvetica Neue"
BUTTON_FONT = ("Helvetica Neue", 12, "bold")
BUTTON_COLOR = "#007AFF"
PROGRESS_BAR_WIDTH = 415


class DataCleaningApp:
    def __init__(self, root=None):
        if root is not None:
            self.root = root
            self._setup_window()
            self._create_ui()
        self.data = None
        self.cleaned_data = None  # Neue Variable für bereinigte Daten
        self.last_saved_path = None
        self.progress_value = 0

    def _setup_window(self):
        # Initialisierung Haupteigenschaften des Fensters
        if self.root:
            self.root.title(WINDOW_TITLE)
            self.root.geometry(WINDOW_SIZE)
            self.root.resizable(False, False)

    def _create_ui(self):
        # Erstellung gesamte Benutzeroberfläche
        if self.root:
            self._create_selection_menu()
            self.header_label = self._create_label("")
            self.load_button = self._create_section(
                "1. Unbereinigte Quelldatei laden:", "Datei laden", self.load_file
            )
            self.save_button = self._create_section(
                "2. Bereinigte Datei speichern:", "Datei speichern", self.save_file
            )
            self.process_button = self._create_section(
                "3. Daten bereinigen und transformieren:", "Start", self.process_data
            )
            self.open_button = self._create_section(
                "4. Bereinigte Datei öffnen:", "Datei öffnen", self.open_file
            )
            self._create_progress_bar()
            self._create_footer_buttons()

    def _create_selection_menu(self):
        # Erstellt das Auswahlmenü für die Reporting-Optionen.
        self.selection_frame = ctk.CTkFrame(self.root, corner_radius=12, fg_color="#F2F2F7", border_width=1,
                                            border_color="#D1D1D6")
        self.selection_frame.pack(pady=10, padx=20, fill="x")

        selection_label = self._create_label(
            "Reporting auswählen:", parent=self.selection_frame, font_size=14, bold=True
        )
        selection_label.pack(pady=5)

        self.selection_var = ctk.StringVar(value="")
        options = ["", "Sales Reporting"]

        selection_menu = ctk.CTkOptionMenu(
            self.selection_frame, values=options, command=self._on_selection, variable=self.selection_var,
            button_color=BUTTON_COLOR, button_hover_color="#0056D2", fg_color="#FFFFFF",
            dropdown_text_color="black", dropdown_fg_color="#F9F9F9", dropdown_hover_color="#E5E5EA",
            text_color="black"
        )
        selection_menu.pack(pady=5)

    def _create_label(self, text, parent=None, font_size=13, bold=False):
        # Hilfsfunktion zum Erstellen von Labels
        if parent is None:
            parent = self.root
        font_style = (LABEL_FONT, font_size, "bold" if bold else "normal")
        return ctk.CTkLabel(
            parent, text=text, wraplength=450, justify="center", font=font_style, text_color="#3A3A3C"
        )

    def _create_section(self, description_text, button_text, command):
        # Erstellung Abschnitt mit einer Beschreibung und einem Button
        frame = ctk.CTkFrame(self.root, corner_radius=12, fg_color="#F2F2F7", border_width=1, border_color="#D1D1D6")
        frame.pack(pady=10, padx=20, fill="x")
        label = self._create_label(description_text, parent=frame)
        label.pack(pady=5)
        button = ctk.CTkButton(
            frame, text=button_text, command=command, corner_radius=8,
            font=BUTTON_FONT, hover_color="#0056D2",
            fg_color=BUTTON_COLOR, text_color="white"
        )
        button.pack(pady=5)
        return button

    def _create_progress_bar(self):
        # Erstellung Fortschrittsanzeige und zugehörige Label
        progress_frame = ctk.CTkFrame(self.root, fg_color="#F2F2F7")
        progress_frame.pack(pady=10, padx=20, fill="x")

        self.progress_bar = ctk.CTkProgressBar(
            progress_frame, width=PROGRESS_BAR_WIDTH, height=20, progress_color=BUTTON_COLOR
        )
        self.progress_bar.set(0)
        self.progress_bar.pack(side="left", pady=10, padx=(20, 2))

        self.progress_label = self._create_label("0%", parent=progress_frame, font_size=10, bold=True)
        self.progress_label.pack(side="right", padx=(0, 10))

    def _create_footer_buttons(self):
        # Erstellung Buttons Fußbereich
        help_button = ctk.CTkButton(
            self.root, text="Hilfe", command=self._open_help_email, corner_radius=8,
            font=("Helvetica Neue", 10, "bold"), fg_color="#F2F2F7", text_color=BUTTON_COLOR,
            hover_color="#E5E5EA", width=50
        )
        help_button.pack(side="right", anchor="se", padx=20, pady=10)

        explanation_button = ctk.CTkButton(
            self.root, text="Erklärung", command=self._show_welcome_message, corner_radius=8,
            font=("Helvetica Neue", 10, "bold"), fg_color="#F2F2F7", text_color=BUTTON_COLOR,
            hover_color="#E5E5EA", width=70
        )
        explanation_button.pack(side="left", anchor="sw", padx=20, pady=10)

    def update_progress(self, value):
        # Aktualisierung Fortschrittsanzeige
        self.progress_bar.set(value)

        progress_texts = {
            0.0: "0%",
            0.33: "33%",
            0.66: "66%",
            1.0: "100%"
        }
        self.progress_label.configure(text=progress_texts.get(value, ""))

    def load_file(self):
        # Laden einer CSV- oder Excel-Datei
        if not self.selection_var.get():
            messagebox.showerror("Fehler", "Bitte Reporting auswählen.")
            return
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV Dateien", "*.csv"), ("Excel Dateien", "*.xlsx")]
        )
        if file_path:
            self.data = pd.read_csv(file_path) if file_path.endswith(".csv") else pd.read_excel(file_path)
            messagebox.showinfo("Erfolg", "Datei erfolgreich geladen.")
            self.update_progress(0.33)
            self.load_button.configure(text="Erfolgreich erledigt", fg_color="green", state="disabled")

    def save_file(self):
        # Speichern bereinigte Daten in Excel-Datei
        if self.data is None:
            messagebox.showerror("Fehler", "Keine Datei geladen.")
            return
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Dateien", "*.xlsx"), ("Alle Dateien", "*.*")]
        )
        if save_path:
            self.cleaned_data = self.clean_data(self.data)
            self.cleaned_data.to_excel(save_path, index=False)
            self.last_saved_path = save_path
            messagebox.showinfo("Erfolg", "Datei erfolgreich gespeichert.")
            self.update_progress(0.66)
            self.save_button.configure(text="Erfolgreich erledigt", fg_color="green", state="disabled")

    @staticmethod
    def _adjust_column_widths(workbook):
        # Anpassung Spaltenbreiten an Inhalt an
        for worksheet in workbook.worksheets:
            for column_cells in worksheet.columns:
                max_length = max(len(str(cell.value)) for cell in column_cells)
                adjusted_width = (max_length + 2)
                worksheet.column_dimensions[column_cells[0].column_letter].width = adjusted_width

    def process_data(self):
        # Bereinigung Daten und Erstellung Diagramme
        if self.cleaned_data is None:
            messagebox.showerror("Fehler", "Keine bereinigten Daten vorhanden. Bitte Schritt 2 ausführen.")
            return
        self.save_charts(self.cleaned_data)
        self.update_progress(1.0)
        self.process_button.configure(text="Erfolgreich erledigt", fg_color="green", state="disabled")

    @staticmethod
    def clean_data(df):
        # Bereinigung Datensatz und Korrektur Datumsformate

        # Entfernen von doppelten Einträgen, außer dem Kommentar
        df = df.drop_duplicates(
            subset=['Datum', 'Verkäufer', 'Region', 'Produkt', 'Verkaufte Menge', 'Umsatz pro Einheit', 'Gesamtumsatz'],
            keep='first'
        ).copy()

        # Bereinigung Datum und Konvertierung
        def parse_date(date):
            try:
                return pd.to_datetime(date, errors='coerce').strftime('%Y-%m-%d')
            except ValueError:
                return None

        # Direkte Zuweisung zu einer Spalte
        df.loc[:, 'Datum'] = df['Datum'].apply(parse_date)

        # Ausfüllen fehlender Datumsangaben mit nächst gültigem Datum
        df.loc[:, 'Datum'] = df['Datum'].ffill()

        # Ausfüllen fehlender Verkäufernamen basierend auf Region
        def fill_verkaeufer(row):
            if pd.isna(row['Verkäufer']):
                if row['Region'] == 'AMERICAS':
                    return 'Peter Schmidt'
                elif row['Region'] == 'EMEA':
                    return 'Stefan Berger'
                elif row['Region'] == 'ASIA':
                    return 'Dirk Donner'
            return row['Verkäufer']

        df.loc[:, 'Verkäufer'] = df.apply(fill_verkaeufer, axis=1)

        # Einheitliche Schreibweise der Verkäufernamen
        df.loc[:, 'Verkäufer'] = df['Verkäufer'].str.title()

        # Ausfüllen fehlender Regionsnamen basierend auf Verkäufer
        def fill_region(row):
            if pd.isna(row['Region']):
                if row['Verkäufer'] == 'Peter Schmidt':
                    return 'AMERICAS'
                elif row['Verkäufer'] == 'Stefan Berger':
                    return 'EMEA'
                elif row['Verkäufer'] == 'Dirk Donner':
                    return 'ASIA'
            return row['Region']

        df.loc[:, 'Region'] = df.apply(fill_region, axis=1)

        # Berechnung Monat und Gesamtumsatz
        df.loc[:, 'Monat'] = pd.to_datetime(df['Datum']).dt.to_period('M')
        df.loc[:, 'Gesamtumsatz'] = df['Verkaufte Menge'] * df['Umsatz pro Einheit']

        return df

    def save_charts(self, df):
        # Erstellung Balkendiagramme und Speicherung in Excel-Datei
        # Darstellung Monate auf Deutsch
        locale.setlocale(locale.LC_TIME, 'de_DE.UTF-8')

        df['Monat'] = pd.to_datetime(df['Datum']).dt.to_period('M')
        df['Gesamtumsatz'] = df['Verkaufte Menge'] * df['Umsatz pro Einheit']

        monthly_sales = df.groupby('Monat')['Gesamtumsatz'].sum()

        # Erstellung erstes Diagramm: Gesamtumsatz pro Monat
        fig, ax = plt.subplots(figsize=(10, 6))
        monthly_sales.plot(kind='bar', ax=ax)
        ax.set_xticklabels([period.strftime('%B %Y') for period in monthly_sales.index], rotation=45, ha="right")
        ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f'{x:,.0f} €'))
        ax.set_title("Gesamtumsatz pro Monat")
        ax.set_xlabel("Monat")
        ax.set_ylabel("Gesamtumsatz in €")
        ax.grid(True)

        # Speicherung des ersten Diagramms
        temp_file_1 = os.path.join(tempfile.gettempdir(), "chart1.png")
        plt.tight_layout()
        plt.savefig(temp_file_1)
        plt.close()

        # Erstellung zweites Diagramm: Umsatz pro Verkäufer pro Monat
        seller_monthly_sales = df.groupby(['Monat', 'Verkäufer'])['Gesamtumsatz'].sum().unstack()

        fig2, ax2 = plt.subplots(figsize=(10, 6))
        seller_monthly_sales.plot(kind='bar', stacked=True, ax=ax2)
        ax2.set_xticklabels([period.strftime('%B %Y') for period in seller_monthly_sales.index], rotation=45,
                            ha="right")
        ax2.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f'{x:,.0f} €'))
        ax2.set_title("Umsatz pro Verkäufer pro Monat")
        ax2.set_xlabel("Monat")
        ax2.set_ylabel("Umsatz in €")
        ax2.grid(True)

        # Speicherung des zweiten Diagramms
        temp_file_2 = os.path.join(tempfile.gettempdir(), "chart2.png")
        plt.tight_layout()
        plt.savefig(temp_file_2)
        plt.close()

        # Export aller Diagramme in Excel-Datei
        if self.last_saved_path:
            with pd.ExcelWriter(self.last_saved_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Bereinigte Daten', index=False)
                workbook = writer.book
                worksheet = workbook.create_sheet(title="Umsatzdiagramm")

                # Hinzufügen erstes Diagramm
                img1 = Image(temp_file_1)
                worksheet.add_image(img1, 'A1')

                # Hinzufügen zweites Diagramm
                img2 = Image(temp_file_2)
                worksheet.add_image(img2, 'A34')

                self._adjust_column_widths(workbook)
                workbook._sheets = [worksheet] + [workbook['Bereinigte Daten']]

            # Löschung temporärer Dateien
            os.remove(temp_file_1)
            os.remove(temp_file_2)

        messagebox.showinfo("Erfolg", "Daten erfolgreich bereinigt und Diagramme hinzugefügt.")

    def open_file(self):
        # Öffnen zuletzt gespeicherte Datei
        if self.last_saved_path:
            if os.name == 'nt':
                os.startfile(self.last_saved_path)
            elif os.name == 'posix':
                subprocess.call(('open', self.last_saved_path))
            else:
                messagebox.showerror("Fehler", "Das Betriebssystem wird nicht unterstützt.")
        else:
            messagebox.showerror("Fehler", "Es wurde noch keine Datei gespeichert.")

    @staticmethod
    def _open_help_email():
        # Öffne Outlook mit Adresse
        webbrowser.open("mailto:me@home.com")

    @staticmethod
    def _show_welcome_message():
        # Anzeige Willkommensnachricht
        messagebox.showinfo(
            "Erklärung",
            "Dieses Tool ermöglicht es Ihnen, CSV- oder Excel-Dateien zu laden, "
            "zu bereinigen und die Ergebnisse zu speichern. Sie können die bereinigte Datei "
            "anschließend öffnen.\n\n"
            "Bei Fragen oder Problemen verwenden Sie bitte den Hilfe Button.\n\n"
            "Software Entwickler: Christopher Haase"
        )

    def _on_selection(self, selected_option):
        # Auswahl Dropdown Menü
        if not selected_option:
            self.header_label.pack_forget()
            self.load_button.configure(
                text="Datei laden", fg_color=BUTTON_COLOR, state="normal",
                command=lambda: messagebox.showerror("Fehler", "Bitte Reporting auswählen.")
            )
            self.save_button.configure(
                text="Datei speichern", fg_color=BUTTON_COLOR, state="normal"
            )
            self.process_button.configure(
                text="Start", fg_color=BUTTON_COLOR, state="normal"
            )
            self.open_button.configure(
                text="Datei öffnen", fg_color=BUTTON_COLOR, state="normal"
            )
            self.update_progress(0.0)
        else:
            self.load_button.configure(command=self.load_file)
            if selected_option == "Sales Reporting":
                self.header_label.configure(
                    text="Sales Reporting:\n"
                         "Analyse der monatlichen Gesamtumsätze und der Einzelumsätze pro Verkäufer"
                         " im Bereich Smartphone-Vertrieb."
                )
            self.header_label.pack(pady=8, after=self.selection_frame)


if __name__ == "__main__":
    main_root = ctk.CTk()  # Erstellung Hauptfenster
    app = DataCleaningApp(main_root)
    main_root.mainloop()
