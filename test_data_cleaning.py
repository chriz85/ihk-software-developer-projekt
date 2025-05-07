#############################################################
# Entwickler: Christopher Haase                             #
# Kurs: Software Developer (IHK) [xxxx]                     #
# Erstellungsdatum: 30.10.2024                              #
# Letzte Änderung: 19.11.2024                               #
# Version: 1.0                                              #
# --------------------------------------------------------- #
# Projektarbeit: Data Cleaning and Transformation Pipeline  #
# Beschreibung: Unittest für Programm Test                  #
# --------------------------------------------------------- #
# Kontakt: me@home.com                                      #
#############################################################


import unittest
import pandas as pd
from data_cleaning_tool import DataCleaningApp  # Importiere dein Hauptprogramm


class TestDataCleaningApp(unittest.TestCase):

    def setUp(self):
        # Beispiel-Daten für die Tests
        data = {
            'Datum': ['01/25/2024', '2024-01-26', '01-27-2024'],
            'Verkäufer': ['Peter Schmidt', 'STEFAN BERGER', 'Dirk Donner'],
            'Region': [None, 'EMEA', 'ASIA'],
            'Produkt': ['Apple iPhone 16 Pro Max', 'Samsung Galaxy S24 Ultra', 'Huawei Pura 70 Ultra'],
            'Verkaufte Menge': [57, 29, 152],
            'Umsatz pro Einheit': [1449, 1239, 1499],
            'Gesamtumsatz': [82593, 35931, 227848],
            'Kommentar': ['Fehlende Region', 'Falsches Datumsformat', 'Falsches Datumsformat']
        }
        self.df = pd.DataFrame(data)  # Erstellen des DataFrames
        self.app = DataCleaningApp()  # Erstellen eines Objekts der Hauptanwendung

    def test_clean_data(self):
        # Die Funktion clean_data aufrufen
        cleaned_df = self.app.clean_data(self.df)

        # Überprüfen, ob fehlende Regionen korrekt ergänzt wurden
        self.assertEqual(cleaned_df['Region'].iloc[0], 'AMERICAS')

        # Überprüfen, ob das Datum korrekt formatiert wurde
        self.assertEqual(cleaned_df['Datum'].iloc[0], '2024-01-25')

    def test_drop_duplicates(self):
        # Duplikate im DataFrame hinzufügen
        df_with_duplicates = pd.concat([self.df, pd.DataFrame([self.df.iloc[0]])], ignore_index=True)

        # Bereinigungsfunktion aufrufen
        cleaned_df = self.app.clean_data(df_with_duplicates)

        # Sicherstellen, dass die Duplikate entfernt wurden
        self.assertEqual(len(cleaned_df), len(self.df))  # Es sollten keine zusätzlichen Zeilen geben

    def test_calculate_total_revenue(self):
        # Gesamtumsatz berechnen und prüfen, ob korrekt
        cleaned_df = self.app.clean_data(self.df)
        self.assertEqual(cleaned_df['Gesamtumsatz'].iloc[0], 57 * 1449)

    def tearDown(self):
        # Hier könntest du Aufräumarbeiten durchführen, wenn nötig
        pass


if __name__ == '__main__':
    unittest.main()
