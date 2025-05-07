# Data Cleaning Tool – Automatisierte Datenbereinigung & Reporting mit Python

Dieses Projekt wurde im Rahmen der Abschlussprüfung zum **Software Developer (IHK)** entwickelt. Es handelt sich um eine Python-basierte Anwendung mit grafischer Benutzeroberfläche (GUI), die den Prozess der Datenbereinigung und Aufbereitung für Vertriebskennzahlen automatisiert.

## Projektziel

Ziel des Tools ist es, **manuelle und fehleranfällige Datenaufbereitungsprozesse** im Vertriebsinnendienst zu ersetzen. Das Programm erkennt automatisch:

- Doppelte Einträge
- Falsche Datumsformate
- Fehlende Regionen

Es erstellt anschließend ein bereinigtes Excel-Dokument samt **visueller Auswertungen (Diagramme)** zur Unterstützung des Sales Reportings.

## Technologien

- **Python 3.12**
- **CustomTkinter** – moderne GUI-Entwicklung
- **Pandas** – Datenmanipulation
- **Matplotlib** – Diagrammerstellung
- **OpenPyXL** – Excel-Verarbeitung
- **PyInstaller** – .exe-Erstellung

## Funktionsumfang

- CSV-/Excel-Dateien einlesen
- Automatisierte Datenbereinigung
- Erstellung bereinigter Excel-Reports inkl.:
  - Monatsumsätze
  - Verkäufervergleiche
  - Diagramme als Bild in Excel integriert
- GUI mit Fortschrittsanzeige und Benutzerführung

## Projektstruktur

```
├── data_cleaning_tool.py 					# Zentrale Python-Datei
├── test_data_cleaning.py 					# Unittest Python-Datei
├── Projektarbeit_CHaase_Data_Cleaning_Tool_bereinigt.pdf	# Projektdokumentation Abschlussarbeit
```

## Projektergebnis

Das Tool wurde vollständig umgesetzt. Es erhöht die Datenqualität, reduziert manuellen Aufwand und bietet Entscheidungsgrundlagen auf Knopfdruck. Das Projekt wurde erfolgreich im Rahmen der IHK-Abschlussprüfung präsentiert.

## Autor

**Christopher Haase**  
Christopher.Haase@me.com  
[GitHub](https://github.com/chriz85) | [LinkedIn](https://www.linkedin.com/in/christopher-haase-938985129/)

---

> 📌 **Hinweis:** Die enthaltenen Dateien wurden von sensiblen Daten bereinigt. Der Quellcode ist als Demo-Version ohne produktive Unternehmensdaten veröffentlicht.
