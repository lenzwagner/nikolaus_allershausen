# Nikolaus & Krampus - Automatische Routenplanung

Vollautomatisches Python-Skript zur optimalen Planung von Nikolaus- und Krampus-Besuchen.

## 📋 Features

- ✅ **Automatische Bedarfsberechnung**: Bottleneck-Analyse mit konfigurierbarem Puffer
- ✅ **Intelligentes Geocoding**: Nominatim (OpenStreetMap) + lokale Fallbacks
- ✅ **Route-Optimierung**: K-Means Clustering + Greedy TSP
- ✅ **Interaktive Karte**: OpenStreetMap mit allen Details
- ✅ **Excel-Export**: Einsatzpläne für jeden Nikolaus/Krampus
- ✅ **Visualisierungen**: Statistiken und Routen-Karten

## 🔧 Installation

### Voraussetzungen

```bash
# Python 3.7 oder höher
python --version

# Installiere benötigte Pakete
pip install pandas openpyxl geopy folium matplotlib scikit-learn
```

## 📂 Input-Datei Format

Die Excel-Datei muss folgende Spalten enthalten:

| Spalte | Beschreibung | Beispiel |
|--------|--------------|----------|
| `ID` | Eindeutige Kind-ID | 1, 2, 3, ... |
| `Adresse` | Vollständige Adresse | "Seestraße 2, 85391 Allershausen" |
| `Tag` | Besuchstag | "5.12 Freitag", "6.12 Samstag" |
| `Uhrzeit` | Zeitslot | "17-18 Uhr", "18-19 Uhr" |
| `Krampus?` | Krampus benötigt? | "ja" oder "nein" |

## 🚀 Verwendung

### Basis-Verwendung

```bash
python nikolaus_planung_komplett.py --input besuche.xlsx
```

### Erweiterte Optionen

```bash
# Ohne Nominatim-Geocoding (schneller, aber weniger genau)
python nikolaus_planung_komplett.py --input besuche.xlsx --no-nominatim

# Mit erhöhtem Puffer
python nikolaus_planung_komplett.py --input besuche.xlsx \
    --puffer-nikolaus 2 \
    --puffer-krampus 2

# Mit anderem Output-Verzeichnis
python nikolaus_planung_komplett.py --input besuche.xlsx --output ergebnisse/
```

## 📊 Output-Dateien

1. **`zuordnung_komplett.xlsx`** - Kompletter Routenplan
2. **`nikolaus_einsatzplaene.xlsx`** - Individuelle Nikolaus-Pläne
3. **`krampus_einsatzplaene.xlsx`** - Individuelle Krampus-Pläne
4. **`routenplan_interaktiv.html`** ⭐ - Interaktive Karte (WICHTIGSTE DATEI)
5. **`statistik.png`** - Visualisierungen
6. **`routen_*.png`** - Routen-Karten pro Tag

## 🔍 Wie funktioniert das?

1. **Bedarfsanalyse**: Bottleneck-Erkennung → Nikoläuse = ⌈Max(Kinder)/3⌉ + Puffer
2. **Geocoding**: Adressen → GPS-Koordinaten (mit Cache)
3. **Clustering**: K-Means gruppiert Kinder geografisch
4. **TSP**: Optimiert Reihenfolge innerhalb jeder Gruppe
5. **Output**: Excel + Visualisierungen + interaktive Karte

## 🎯 Beispiel

```bash
# Erstplanung mit 42 Kindern
python nikolaus_planung_komplett.py --input besuche_2024.xlsx

# Ergebnis:
# → 7 Nikoläuse (6 + 1 Puffer)
# → 5 Krampusse (4 + 1 Puffer)
# → Alle Dateien in outputs/
```

## 🐛 Troubleshooting

### Geocoding schlägt fehl
```bash
python nikolaus_planung_komplett.py --input besuche.xlsx --no-nominatim
```

### Koordinaten korrigieren
1. Bearbeite `koordinaten_cache.csv`
2. Führe Skript erneut aus

### Hilfe anzeigen
```bash
python nikolaus_planung_komplett.py --help
```

## 📈 Performance

| Kinder | Mit Nominatim | Ohne Nominatim |
|--------|---------------|----------------|
| 10     | ~30 Sek       | ~5 Sek         |
| 42     | ~2 Min        | ~10 Sek        |
| 100    | ~5 Min        | ~20 Sek        |

## 🎅 Viel Erfolg!

```
      *
     /.\
    /..'\
    /'.'\
   /.''.'\
   /.'.'.\
  /'.''.'.'\
 ^^^[_]^^^
```
