#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=================================================================================
NIKOLAUS & KRAMPUS - AUTOMATISCHE ROUTENPLANUNG
=================================================================================

Dieses Skript führt eine vollständige Routenplanung für Nikolaus- und 
Krampus-Besuche durch:

1. Einlesen der Excel-Datei mit Besuchsanfragen
2. Automatische Bedarfsberechnung (Bottleneck-Analyse + Puffer)
3. Geocoding der Adressen (mit mehreren Fallback-Optionen)
4. Route-Optimierung mit K-Means Clustering + TSP
5. Erstellung aller Ausgabedateien (Excel, Visualisierungen, Karte)

VERWENDUNG:
-----------
python nikolaus_planung_komplett.py --input <excel_datei.xlsx> [Optionen]

ANFORDERUNGEN:
--------------
pip install pandas openpyxl geopy folium matplotlib scikit-learn

AUTOR: Claude (Anthropic)
DATUM: November 2024
VERSION: 2.0
=================================================================================
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import folium
from sklearn.cluster import KMeans
from datetime import datetime
import warnings
import time
import argparse
import sys
import os

warnings.filterwarnings('ignore')

# =============================================================================
# KONFIGURATION
# =============================================================================

class Config:
    """Zentrale Konfiguration"""
    
    # Geocoding-Einstellungen
    GEOCODING_TIMEOUT = 10
    GEOCODING_RETRY = 3
    GEOCODING_DELAY = 1.5  # Sekunden zwischen Anfragen (Nominatim Rate-Limit)
    
    # Optimierungs-Parameter
    MAX_KINDER_PRO_TEAM = 3
    PUFFER_NIKOLAUS = 1
    PUFFER_KRAMPUS = 1
    
    # Visualisierungs-Einstellungen
    FIGURE_DPI = 300
    MAP_ZOOM = 14
    
    # Farben für Visualisierungen
    COLORS = ['#e41a1c', '#377eb8', '#4daf4a', '#984ea3', '#ff7f00', 
              '#ffff33', '#a65628', '#f781bf', '#999999', '#66c2a5',
              '#fc8d62', '#8da0cb', '#e78ac3', '#a6d854', '#ffd92f',
              '#e5c494', '#b3b3b3', '#8dd3c7']

# =============================================================================
# GEOCODING-FUNKTIONEN
# =============================================================================

def geocode_address_nominatim(address, retry=3, delay=1.5):
    """
    Geocoding mit Nominatim (OpenStreetMap)
    
    Parameters:
    -----------
    address : str
        Vollständige Adresse
    retry : int
        Anzahl der Wiederholungsversuche
    delay : float
        Wartezeit zwischen Versuchen (Sekunden)
    
    Returns:
    --------
    tuple : (latitude, longitude) oder (None, None)
    """
    try:
        from geopy.geocoders import Nominatim
        from geopy.exc import GeocoderTimedOut, GeocoderServiceError
        
        geolocator = Nominatim(user_agent="nikolaus_planung_v2", timeout=10)
        
        for attempt in range(retry):
            try:
                # Warte zwischen Anfragen (Rate-Limiting)
                if attempt > 0:
                    time.sleep(delay * 2)
                else:
                    time.sleep(delay)
                
                location = geolocator.geocode(address + ", Germany")
                
                if location:
                    return (location.latitude, location.longitude)
                    
            except (GeocoderTimedOut, GeocoderServiceError) as e:
                if attempt == retry - 1:
                    print(f"   ⚠ Geocoding fehlgeschlagen: {address}")
                    return (None, None)
                continue
                
        return (None, None)
        
    except Exception as e:
        print(f"   ✗ Fehler beim Geocoding: {e}")
        return (None, None)


def get_fallback_coordinates(address):
    """
    Fallback-Koordinaten basierend auf Straßennamen in Allershausen
    
    Parameters:
    -----------
    address : str
        Adresse
    
    Returns:
    --------
    tuple : (latitude, longitude) oder (None, None)
    """
    
    # Basis-Koordinaten Allershausen
    BASE_LAT = 48.4167
    BASE_LON = 11.7333
    
    # Straßen-Dictionary für Allershausen (basierend auf geografischer Lage)
    street_coords = {
        # Nord

        'seestr': (48.4288453,11.5960513),
        'dominikus-käser': (48.4361731,11.5945349),
        'anton-lamprecht': (48.425052,11.5949005),
        'mühlbachstr': (48.4347749,11.5946798),
        'auenstraße': (48.4291369,11.5994042),
        'joseph-haydn': (48.4178, 11.7358),
        'franz-liszt': (48.4188, 11.7365),
        'händelstr': (48.4188, 11.7359),
        'mozart': (48.4180, 11.7365),
        
        # Zentrum
        'pfarrsaal': (48.4348473,11.586955),
        'kirchstr': (48.4172, 11.7337),
        'schulstr': (48.4165, 11.7350),
        'bonhoeffer': (48.4176, 11.7342),
        'kreuth': (48.4144, 11.7336),
        
        # Süd
        'glonntalstr': (48.4321969,11.6009219),
        'bergstr': (48.4145, 11.7315),
        'am anger': (48.4159, 11.7354),
        'kohlstattweg': (48.4148, 11.7345),
        'göttschlag': (48.4138, 11.7360),
        
        # West
        'amselweg': (48.4165, 11.7265),
        'drosselweg': (48.4168, 11.7262),
        'breimannweg': (48.4168, 11.7281),
        'am moosbolt': (48.4173, 11.7273),
        
        # Ost
        'adalharstr': (48.4162, 11.7394),
        'albert-schweitzer': (48.4160, 11.7401),
        'blumenstr': (48.4159, 11.7391),
        
        # Ortsteile
        'tünzhausen': (48.4130, 11.7247),
        'an der linde': (48.4131, 11.7247),
        'zur hochstatt': (48.4130, 11.7245),
        'amperstr': (48.4391936,11.6256344),
        
        'leonhardsbuch': (48.4103, 11.7318),
        'dorfstr': (48.4103, 11.7316),
        'am kirchberg': (48.4098, 11.7325),
    }
    
    # Suche nach Straßenname in Adresse
    address_lower = address.lower()
    
    for street_key, (lat, lon) in street_coords.items():
        if street_key in address_lower:
            # Füge kleine zufällige Variation für Hausnummern hinzu
            offset_lat = np.random.uniform(-0.0005, 0.0005)
            offset_lon = np.random.uniform(-0.0005, 0.0005)
            return (lat + offset_lat, lon + offset_lon)
    
    # Fallback: Zentrum Allershausen mit kleiner Variation
    offset_lat = np.random.uniform(-0.002, 0.002)
    offset_lon = np.random.uniform(-0.002, 0.002)
    return (BASE_LAT + offset_lat, BASE_LON + offset_lon)


def geocode_addresses(df, use_nominatim=True, cache_file='koordinaten_cache.csv'):
    """
    Geocodiert alle Adressen mit Caching und Fallback
    
    Parameters:
    -----------
    df : DataFrame
        DataFrame mit 'Adresse' Spalte
    use_nominatim : bool
        Ob Nominatim verwendet werden soll (langsam aber genau)
    cache_file : str
        Pfad zum Cache-File
    
    Returns:
    --------
    DataFrame : DataFrame mit zusätzlichen Spalten 'Latitude', 'Longitude'
    """
    
    print("\n" + "="*100)
    print("GEOCODING DER ADRESSEN")
    print("="*100)
    
    df = df.copy()
    
    # Prüfe auf manuelle Koordinaten
    has_manual_coords = 'Latitude' in df.columns and 'Longitude' in df.columns
    
    if not has_manual_coords:
        df['Latitude'] = None
        df['Longitude'] = None
    else:
        # Stelle sicher, dass leere Werte als None/NaN behandelt werden
        df['Latitude'] = pd.to_numeric(df['Latitude'], errors='coerce')
        df['Longitude'] = pd.to_numeric(df['Longitude'], errors='coerce')
    
    # Versuche Cache zu laden
    cache = {}
    if os.path.exists(cache_file):
        try:
            cache_df = pd.read_csv(cache_file)
            cache = dict(zip(cache_df['Adresse'], 
                           zip(cache_df['Latitude'], cache_df['Longitude'])))
            print(f"✓ Cache geladen: {len(cache)} Adressen")
        except:
            print("⚠ Cache konnte nicht geladen werden")
    
    total = len(df)
    success_nominatim = 0
    success_fallback = 0
    success_manual = 0
    
    for idx, row in df.iterrows():
        address = row['Adresse']
        
        # 1. Prüfe manuelle Koordinaten
        if has_manual_coords and pd.notna(row['Latitude']) and pd.notna(row['Longitude']):
            print(f"[{idx+1}/{total}] Manuell: {address[:40]}... ({row['Latitude']:.5f}, {row['Longitude']:.5f})")
            success_manual += 1
            continue
        
        # 2. Prüfe Cache
        if address in cache:
            df.at[idx, 'Latitude'] = cache[address][0]
            df.at[idx, 'Longitude'] = cache[address][1]
            continue
        
        print(f"[{idx+1}/{total}] Geocoding: {address[:60]}...")
        
        lat, lon = None, None
        
        # Versuch 1: Nominatim (falls aktiviert)
        if use_nominatim:
            lat, lon = geocode_address_nominatim(address)
            if lat and lon:
                success_nominatim += 1
                print(f"   ✓ Nominatim: {lat:.6f}, {lon:.6f}")
        
        # Versuch 2: Fallback auf Straßen-Dictionary
        if not lat or not lon:
            lat, lon = get_fallback_coordinates(address)
            success_fallback += 1
            print(f"   ⚠ Fallback: {lat:.6f}, {lon:.6f}")
        
        df.at[idx, 'Latitude'] = lat
        df.at[idx, 'Longitude'] = lon
        
        # Speichere im Cache
        cache[address] = (lat, lon)
    
    # Speichere Cache
    try:
        cache_df = pd.DataFrame([
            {'Adresse': addr, 'Latitude': coords[0], 'Longitude': coords[1]}
            for addr, coords in cache.items()
        ])
        cache_df.to_csv(cache_file, index=False)
        print(f"\n✓ Cache gespeichert: {cache_file}")
    except Exception as e:
        print(f"⚠ Cache konnte nicht gespeichert werden: {e}")
    
    print(f"\n📊 Geocoding-Statistik:")
    print(f"   • Manuell gesetzt: {success_manual}")
    print(f"   • Nominatim erfolgreich: {success_nominatim}")
    print(f"   • Fallback verwendet: {success_fallback}")
    print(f"   • Gesamt: {total}")
    
    return df


# =============================================================================
# BEDARFSANALYSE
# =============================================================================

def analyze_demand(df, puffer_nikolaus=1, puffer_krampus=1):
    """
    Analysiert den Bedarf an Nikoläusen und Krampussen
    
    Parameters:
    -----------
    df : DataFrame
        DataFrame mit Besuchsdaten
    puffer_nikolaus : int
        Puffer für Nikoläuse
    puffer_krampus : int
        Puffer für Krampusse
    
    Returns:
    --------
    dict : Bedarfs-Informationen
    """
    
    print("\n" + "="*100)
    print("BEDARFSANALYSE")
    print("="*100)
    
    # Zähle Kinder pro Zeitslot
    bedarf = df.groupby(['Tag', 'Uhrzeit']).size().reset_index(name='Anzahl_Kinder')
    
    # Berechne benötigte Teams (max 3 Kinder pro Team)
    bedarf['Teams_benötigt'] = np.ceil(bedarf['Anzahl_Kinder'] / 3).astype(int)
    
    # Finde Bottleneck
    max_teams = bedarf['Teams_benötigt'].max()
    bottleneck = bedarf[bedarf['Teams_benötigt'] == max_teams].iloc[0]
    
    print(f"\n📊 Bottleneck-Analyse:")
    print(f"   • Kritischer Zeitpunkt: {bottleneck['Tag']}, {bottleneck['Uhrzeit']}")
    print(f"   • Kinder: {bottleneck['Anzahl_Kinder']}")
    print(f"   • Teams benötigt: {max_teams}")
    
    # Berechne Gesamtbedarf
    nikolaus_bedarf = max_teams + puffer_nikolaus
    
    # Krampus-Bedarf
    krampus_besuche = df[df['Krampus?'] == 'ja']
    krampus_pro_slot = krampus_besuche.groupby(['Tag', 'Uhrzeit']).size()
    max_krampus_gleichzeitig = krampus_pro_slot.max() if len(krampus_pro_slot) > 0 else 0
    krampus_bedarf = int(np.ceil(max_krampus_gleichzeitig / 3)) + puffer_krampus
    
    print(f"\n✅ EMPFOHLENE RESSOURCEN:")
    print(f"   • Nikoläuse: {nikolaus_bedarf} (inkl. +{puffer_nikolaus} Puffer)")
    print(f"   • Krampusse: {krampus_bedarf} (inkl. +{puffer_krampus} Puffer)")
    
    print(f"\n📋 Bedarf pro Zeitslot:")
    print(bedarf.to_string(index=False))
    
    return {
        'nikolaus_bedarf': nikolaus_bedarf,
        'krampus_bedarf': krampus_bedarf,
        'bottleneck': bottleneck,
        'bedarf_detail': bedarf
    }


# =============================================================================
# DISTANZ-BERECHNUNG
# =============================================================================

def haversine_distance(lat1, lon1, lat2, lon2):
    """
    Berechnet die Distanz zwischen zwei GPS-Koordinaten in km
    
    Parameters:
    -----------
    lat1, lon1 : float
        Koordinaten Punkt 1
    lat2, lon2 : float
        Koordinaten Punkt 2
    
    Returns:
    --------
    float : Distanz in km
    """
    R = 6371  # Erdradius in km
    
    lat1_rad = np.radians(lat1)
    lat2_rad = np.radians(lat2)
    delta_lat = np.radians(lat2 - lat1)
    delta_lon = np.radians(lon2 - lon1)
    
    a = np.sin(delta_lat/2)**2 + np.cos(lat1_rad) * np.cos(lat2_rad) * np.sin(delta_lon/2)**2
    c = 2 * np.arctan2(np.sqrt(a), np.sqrt(1-a))
    
    return R * c


# =============================================================================
# ROUTE-OPTIMIERUNG
# =============================================================================

def greedy_tsp(coords, start_idx=0):
    """
    Greedy Nearest-Neighbor TSP für eine Gruppe von Koordinaten
    
    Parameters:
    -----------
    coords : list of tuples
        Liste von (lat, lon) Koordinaten
    start_idx : int
        Start-Index
    
    Returns:
    --------
    list : Optimierte Reihenfolge der Indizes
    """
    n = len(coords)
    if n <= 1:
        return list(range(n))
    
    unvisited = set(range(n))
    route = [start_idx]
    unvisited.remove(start_idx)
    
    current = start_idx
    
    while unvisited:
        nearest = min(unvisited, 
                     key=lambda x: haversine_distance(
                         coords[current][0], coords[current][1],
                         coords[x][0], coords[x][1]))
        route.append(nearest)
        unvisited.remove(nearest)
        current = nearest
    
    return route


def optimize_routes(df_slot, num_teams, max_capacity=3):
    """
    Optimiert Routen für einen Zeitslot
    
    Parameters:
    -----------
    df_slot : DataFrame
        Kinder für einen Zeitslot
    num_teams : int
        Anzahl der Teams
    max_capacity : int
        Max. Kinder pro Team
    
    Returns:
    --------
    list : Liste von Team-Zuordnungen mit Routen
    """
    
    coords = df_slot[['Latitude', 'Longitude']].values
    kind_ids = df_slot['ID'].values
    
    # K-Means Clustering
    if num_teams > 1:
        kmeans = KMeans(n_clusters=num_teams, random_state=42, n_init=10)
        clusters = kmeans.fit_predict(coords)
    else:
        clusters = np.zeros(len(coords), dtype=int)
    
    # Balanciere Cluster (max 3 Kinder pro Team)
    cluster_sizes = np.bincount(clusters, minlength=num_teams)
    
    for cluster_id in range(num_teams):
        while cluster_sizes[cluster_id] > max_capacity:
            # Finde Punkt, der am weitesten vom Centroid entfernt ist
            cluster_points = np.where(clusters == cluster_id)[0]
            if num_teams == 1:
                break
                
            centroid = coords[cluster_points].mean(axis=0)
            distances = [haversine_distance(centroid[0], centroid[1], 
                                           coords[i][0], coords[i][1])
                        for i in cluster_points]
            
            farthest_idx = cluster_points[np.argmax(distances)]
            
            # Finde Cluster mit Kapazität
            target_cluster = np.argmin([cluster_sizes[c] if c != cluster_id else max_capacity + 1 
                                       for c in range(num_teams)])
            
            if cluster_sizes[target_cluster] < max_capacity:
                clusters[farthest_idx] = target_cluster
                cluster_sizes[cluster_id] -= 1
                cluster_sizes[target_cluster] += 1
            else:
                break
    
    # Erstelle Routen für jeden Cluster
    teams = []
    for team_id in range(num_teams):
        team_indices = np.where(clusters == team_id)[0]
        
        if len(team_indices) == 0:
            continue
        
        # TSP für diesen Cluster
        team_coords = coords[team_indices]
        route_order = greedy_tsp(team_coords)
        
        # Berechne Distanzen
        distances = []
        for i in range(len(route_order)):
            if i < len(route_order) - 1:
                idx1 = team_indices[route_order[i]]
                idx2 = team_indices[route_order[i+1]]
                dist = haversine_distance(coords[idx1][0], coords[idx1][1],
                                         coords[idx2][0], coords[idx2][1])
                distances.append(dist)
            else:
                distances.append(0.0)
        
        # Erstelle Team-Daten
        team_data = []
        for i, orig_idx in enumerate(route_order):
            global_idx = team_indices[orig_idx]
            team_data.append({
                'Kind_ID': kind_ids[global_idx],
                'Latitude': coords[global_idx][0],
                'Longitude': coords[global_idx][1],
                'Besuchsreihenfolge': i + 1,
                'Distanz_zum_naechsten': distances[i]
            })
        
        teams.append({
            'team_id': team_id + 1,
            'visits': team_data,
            'total_distance': sum(distances)
        })
    
    return teams


def create_full_route_plan(df, nikolaus_bedarf, krampus_bedarf):
    """
    Erstellt vollständigen Routenplan für alle Zeitslots
    
    Parameters:
    -----------
    df : DataFrame
        Vollständige Besuchsdaten mit Koordinaten
    nikolaus_bedarf : int
        Anzahl verfügbarer Nikoläuse
    krampus_bedarf : int
        Anzahl verfügbarer Krampusse
    
    Returns:
    --------
    DataFrame : Vollständiger Routenplan
    """
    
    print("\n" + "="*100)
    print("ROUTE-OPTIMIERUNG")
    print("="*100)
    
    all_routes = []
    
    for (tag, uhrzeit), slot_group in df.groupby(['Tag', 'Uhrzeit'], sort=False):
        num_kinder = len(slot_group)
        num_teams = min(int(np.ceil(num_kinder / 3)), nikolaus_bedarf)
        
        print(f"\n📅 {tag}, {uhrzeit}: {num_kinder} Kinder → {num_teams} Teams")
        
        teams = optimize_routes(slot_group, num_teams)
        
        for team in teams:
            for visit in team['visits']:
                route_data = {
                    'Tag': tag,
                    'Uhrzeit': uhrzeit,
                    'Team_Nr': team['team_id'],
                    'Kind_ID': visit['Kind_ID'],
                    'Latitude': visit['Latitude'],
                    'Longitude': visit['Longitude'],
                    'Besuchsreihenfolge': visit['Besuchsreihenfolge'],
                    'Distanz_zum_naechsten': visit['Distanz_zum_naechsten']
                }
                all_routes.append(route_data)
            
            print(f"   Team {team['team_id']}: {len(team['visits'])} Kinder, "
                  f"{team['total_distance']:.2f} km")
    
    routes_df = pd.DataFrame(all_routes)
    
    # Merge mit Original-Daten
    routes_df = routes_df.merge(
        df[['ID', 'Adresse', 'Krampus?']], 
        left_on='Kind_ID', 
        right_on='ID', 
        how='left'
    ).drop('ID', axis=1)
    
    # Sortiere
    routes_df = routes_df.sort_values(['Tag', 'Uhrzeit', 'Team_Nr', 'Besuchsreihenfolge'])
    
    print(f"\n✅ Route-Optimierung abgeschlossen!")
    print(f"   • Gesamtdistanz: {routes_df['Distanz_zum_naechsten'].sum():.2f} km")
    
    return routes_df


def assign_staff_ids(routes_df, nikolaus_bedarf, krampus_bedarf):
    """
    Weist Nikolaus- und Krampus-IDs zu
    
    Parameters:
    -----------
    routes_df : DataFrame
        Routenplan
    nikolaus_bedarf : int
        Anzahl Nikoläuse
    krampus_bedarf : int
        Anzahl Krampusse
    
    Returns:
    --------
    DataFrame : Routenplan mit Staff-IDs
    """
    
    print("\n" + "="*100)
    print("NIKOLAUS & KRAMPUS ZUORDNUNG")
    print("="*100)
    
    routes_df = routes_df.copy()
    routes_df['Nikolaus_ID'] = ''
    routes_df['Krampus_ID'] = ''
    
    nikolaus_ids = [f'N{i+1}' for i in range(nikolaus_bedarf)]
    krampus_ids = [f'K{i+1}' for i in range(krampus_bedarf)]
    
    nik_counter = 0
    kram_counter = 0
    
    for (tag, uhrzeit, team), group in routes_df.groupby(['Tag', 'Uhrzeit', 'Team_Nr']):
        # Weise Nikolaus zu
        nik_idx = nik_counter % len(nikolaus_ids)
        nik_id = nikolaus_ids[nik_idx]
        routes_df.loc[group.index, 'Nikolaus_ID'] = nik_id
        nik_counter += 1
        
        # Weise Krampus zu (gekoppelt an Nikolaus)
        # Prüfe, ob in diesem Team überhaupt ein Krampus benötigt wird
        needs_krampus = (group['Krampus?'] == 'ja').any()
        
        if needs_krampus:
            # Kopple Krampus an Nikolaus (N1 -> K1, N2 -> K2, etc.)
            # Falls mehr Nikoläuse als Krampusse, fängt es wieder von vorne an
            kram_idx = nik_idx % len(krampus_ids)
            kram_id = krampus_ids[kram_idx]
            
            for idx in group.index:
                if routes_df.loc[idx, 'Krampus?'] == 'ja':
                    routes_df.loc[idx, 'Krampus_ID'] = kram_id
                else:
                    routes_df.loc[idx, 'Krampus_ID'] = '-'
        else:
            routes_df.loc[group.index, 'Krampus_ID'] = '-'
    
    # Statistik
    print(f"\n📊 Zuordnungs-Statistik:")
    for nik_id in nikolaus_ids:
        count = (routes_df['Nikolaus_ID'] == nik_id).sum()
        dist = routes_df[routes_df['Nikolaus_ID'] == nik_id]['Distanz_zum_naechsten'].sum()
        print(f"   • {nik_id}: {count} Besuche, {dist:.2f} km")
    
    print()
    for kram_id in krampus_ids:
        count = (routes_df['Krampus_ID'] == kram_id).sum()
        dist = routes_df[routes_df['Krampus_ID'] == kram_id]['Distanz_zum_naechsten'].sum()
        print(f"   • {kram_id}: {count} Besuche, {dist:.2f} km")
    
    return routes_df


# =============================================================================
# VISUALISIERUNGEN
# =============================================================================

def create_visualizations(routes_df, output_dir='outputs'):
    """
    Erstellt alle Visualisierungen
    
    Parameters:
    -----------
    routes_df : DataFrame
        Routenplan
    output_dir : str
        Ausgabe-Verzeichnis
    """
    
    print("\n" + "="*100)
    print("ERSTELLE VISUALISIERUNGEN")
    print("="*100)
    
    os.makedirs(output_dir, exist_ok=True)
    
    # 1. Statistik-Plot
    fig, axes = plt.subplots(2, 2, figsize=(16, 12))
    
    # Besuche pro Nikolaus
    nik_counts = routes_df.groupby('Nikolaus_ID').size()
    axes[0, 0].bar(nik_counts.index, nik_counts.values, color='#e41a1c')
    axes[0, 0].set_title('Besuche pro Nikolaus', fontsize=14, fontweight='bold')
    axes[0, 0].set_xlabel('Nikolaus ID')
    axes[0, 0].set_ylabel('Anzahl Besuche')
    axes[0, 0].grid(alpha=0.3)
    
    # Distanz pro Nikolaus
    nik_dist = routes_df.groupby('Nikolaus_ID')['Distanz_zum_naechsten'].sum()
    axes[0, 1].bar(nik_dist.index, nik_dist.values, color='#377eb8')
    axes[0, 1].set_title('Distanz pro Nikolaus (km)', fontsize=14, fontweight='bold')
    axes[0, 1].set_xlabel('Nikolaus ID')
    axes[0, 1].set_ylabel('Distanz (km)')
    axes[0, 1].grid(alpha=0.3)
    
    # Besuche pro Krampus
    kram_data = routes_df[routes_df['Krampus_ID'] != '-']
    kram_counts = kram_data.groupby('Krampus_ID').size()
    axes[1, 0].bar(kram_counts.index, kram_counts.values, color='#ff7f00')
    axes[1, 0].set_title('Besuche pro Krampus', fontsize=14, fontweight='bold')
    axes[1, 0].set_xlabel('Krampus ID')
    axes[1, 0].set_ylabel('Anzahl Besuche')
    axes[1, 0].grid(alpha=0.3)
    
    # Besuche pro Tag
    day_counts = routes_df.groupby('Tag').size()
    axes[1, 1].bar(day_counts.index, day_counts.values, color='#4daf4a')
    axes[1, 1].set_title('Besuche pro Tag', fontsize=14, fontweight='bold')
    axes[1, 1].set_xlabel('Tag')
    axes[1, 1].set_ylabel('Anzahl Besuche')
    axes[1, 1].grid(alpha=0.3)
    
    plt.tight_layout()
    plt.savefig(f'{output_dir}/statistik.png', dpi=Config.FIGURE_DPI, bbox_inches='tight')
    plt.close()
    
    print(f"✓ Statistik gespeichert: statistik.png")
    
    # 2. Route-Karten pro Tag
    for tag in routes_df['Tag'].unique():
        fig, ax = plt.subplots(figsize=(16, 12))
        
        day_data = routes_df[routes_df['Tag'] == tag]
        
        # Plot jede Route
        for (uhrzeit, team), group in day_data.groupby(['Uhrzeit', 'Team_Nr']):
            group = group.sort_values('Besuchsreihenfolge')
            
            color = Config.COLORS[team % len(Config.COLORS)]
            
            # Linie
            ax.plot(group['Longitude'], group['Latitude'], 
                   'o-', color=color, linewidth=2, markersize=8,
                   label=f'{uhrzeit}, Team {team}')
            
            # Nummern
            for idx, row in group.iterrows():
                ax.text(row['Longitude'], row['Latitude'], 
                       str(int(row['Besuchsreihenfolge'])),
                       fontsize=8, ha='center', va='center',
                       bbox=dict(boxstyle='circle', facecolor='white', alpha=0.8))
        
        ax.set_xlabel('Longitude', fontsize=12)
        ax.set_ylabel('Latitude', fontsize=12)
        ax.set_title(f'Routen: {tag}', fontsize=16, fontweight='bold')
        ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=8)
        ax.grid(alpha=0.3)
        
        plt.tight_layout()
        filename = tag.replace('.', '').replace(' ', '_').lower()
        plt.savefig(f'{output_dir}/routen_{filename}.png', 
                   dpi=Config.FIGURE_DPI, bbox_inches='tight')
        plt.close()
        
        print(f"✓ Routen-Karte gespeichert: routen_{filename}.png")


def create_interactive_map(routes_df, output_dir='outputs'):
    """
    Erstellt interaktive OpenStreetMap-Karte
    
    Parameters:
    -----------
    routes_df : DataFrame
        Routenplan
    output_dir : str
        Ausgabe-Verzeichnis
    """
    
    print("\n" + "="*100)
    print("ERSTELLE INTERAKTIVE KARTE")
    print("="*100)
    
    os.makedirs(output_dir, exist_ok=True)
    
    # Berechne Kartenzentrum
    center_lat = routes_df['Latitude'].mean()
    center_lon = routes_df['Longitude'].mean()
    
    # Erstelle Karte
    m = folium.Map(
        location=[center_lat, center_lon],
        zoom_start=Config.MAP_ZOOM,
        tiles='OpenStreetMap'
    )
    
    # Feature-Gruppen für Layer-Control
    feature_groups = {}
    
    for tag in routes_df['Tag'].unique():
        feature_groups[tag] = folium.FeatureGroup(name=tag)
    
    # Plotte Routen
    for (tag, uhrzeit, team), group in routes_df.groupby(['Tag', 'Uhrzeit', 'Team_Nr']):
        group = group.sort_values('Besuchsreihenfolge')
        
        color = Config.COLORS[team % len(Config.COLORS)]
        
        # Polyline
        coordinates = [[row['Latitude'], row['Longitude']] 
                      for _, row in group.iterrows()]
        
        folium.PolyLine(
            coordinates,
            color=color,
            weight=3,
            opacity=0.7
        ).add_to(feature_groups[tag])
        
        # Marker
        for idx, row in group.iterrows():
            # Marker-Farbe
            marker_color = 'red' if row['Krampus_ID'] != '-' else 'blue'
            icon_name = 'user-times' if row['Krampus_ID'] != '-' else 'user'
            
            # Popup
            popup_html = f"""
            <b>Stop {int(row['Besuchsreihenfolge'])}</b><br>
            <b>Team {team}</b><br>
            {tag}, {uhrzeit}<br><br>
            <b>Nikolaus:</b> {row['Nikolaus_ID']}<br>
            <b>Krampus:</b> {row['Krampus_ID']}<br><br>
            <b>Kind-ID:</b> {int(row['Kind_ID'])}<br>
            <b>Adresse:</b> {row['Adresse']}<br><br>
            <b>GPS:</b> {row['Latitude']:.6f}, {row['Longitude']:.6f}
            """
            
            folium.Marker(
                location=[row['Latitude'], row['Longitude']],
                popup=folium.Popup(popup_html, max_width=300),
                tooltip=f"ID {int(row['Kind_ID'])} - Stop {int(row['Besuchsreihenfolge'])}",
                icon=folium.Icon(color=marker_color, icon=icon_name, prefix='fa')
            ).add_to(feature_groups[tag])
            
            # Nummer
            folium.CircleMarker(
                location=[row['Latitude'], row['Longitude']],
                radius=12,
                color=color,
                fill=True,
                fillColor='white',
                fillOpacity=1,
                weight=2
            ).add_to(feature_groups[tag])
            
            folium.Marker(
                location=[row['Latitude'], row['Longitude']],
                icon=folium.DivIcon(html=f"""
                    <div style="font-size: 10pt; color: black; font-weight: bold;">
                        {int(row['Besuchsreihenfolge'])}
                    </div>
                """)
            ).add_to(feature_groups[tag])
    
    # Füge Feature-Gruppen hinzu
    for fg in feature_groups.values():
        fg.add_to(m)
    
    # Layer Control
    folium.LayerControl().add_to(m)
    
    # Legende
    legend_html = """
    <div style="position: fixed; bottom: 50px; left: 50px; width: 200px;
                background-color: white; border:2px solid grey; z-index:9999;
                font-size:14px; padding: 10px">
    <p><b>Legende</b></p>
    <p><i class="fa fa-map-marker" style="color:red"></i> Mit Krampus</p>
    <p><i class="fa fa-map-marker" style="color:blue"></i> Ohne Krampus</p>
    <p><i class="fa fa-circle" style="color:grey"></i> Besuchsreihenfolge</p>
    <p>━━━ Route</p>
    </div>
    """
    m.get_root().html.add_child(folium.Element(legend_html))
    
    # Speichere
    map_file = f'{output_dir}/routenplan_interaktiv.html'
    m.save(map_file)
    
    print(f"✓ Interaktive Karte gespeichert: routenplan_interaktiv.html")
    
    return m


# =============================================================================
# EXCEL-EXPORT
# =============================================================================

def export_to_excel(routes_df, df_original, demand_info, output_dir='outputs'):
    """
    Exportiert alle Daten nach Excel
    
    Parameters:
    -----------
    routes_df : DataFrame
        Routenplan
    df_original : DataFrame
        Original-Daten mit Koordinaten
    demand_info : dict
        Bedarfs-Informationen
    output_dir : str
        Ausgabe-Verzeichnis
    """
    
    print("\n" + "="*100)
    print("EXPORTIERE EXCEL-DATEIEN")
    print("="*100)
    
    os.makedirs(output_dir, exist_ok=True)
    
    # 1. Haupt-Zuordnung
    main_file = f'{output_dir}/zuordnung_komplett.xlsx'
    with pd.ExcelWriter(main_file, engine='openpyxl') as writer:
        routes_df.to_excel(writer, sheet_name='Routen', index=False)
        df_original.to_excel(writer, sheet_name='Koordinaten', index=False)
        demand_info['bedarf_detail'].to_excel(writer, sheet_name='Bedarf', index=False)
    
    print(f"✓ Haupt-Datei gespeichert: zuordnung_komplett.xlsx")
    
    # 2. Nikolaus-Einsatzpläne
    nikolaus_plans = []
    for (nikolaus_id, tag, uhrzeit), group in routes_df.groupby(['Nikolaus_ID', 'Tag', 'Uhrzeit']):
        group = group.sort_values('Besuchsreihenfolge')
        route_ids = ' → '.join([f"ID{int(x)}" for x in group['Kind_ID']])
        krampus_partners = group[group['Krampus_ID'] != '-']['Krampus_ID'].unique()
        krampus_str = ', '.join(krampus_partners) if len(krampus_partners) > 0 else '-'
        
        nikolaus_plans.append({
            'Nikolaus_ID': nikolaus_id,
            'Tag': tag,
            'Uhrzeit': uhrzeit,
            'Team': group.iloc[0]['Team_Nr'],
            'Krampus_Partner': krampus_str,
            'Route_IDs': route_ids,
            'Anzahl_Kinder': len(group),
            'Distanz_km': round(group['Distanz_zum_naechsten'].sum(), 2)
        })
    
    nik_df = pd.DataFrame(nikolaus_plans)
    nik_df.to_excel(f'{output_dir}/nikolaus_einsatzplaene.xlsx', index=False)
    print(f"✓ Nikolaus-Pläne gespeichert: nikolaus_einsatzplaene.xlsx")
    
    # 3. Krampus-Einsatzpläne
    krampus_plans = []
    krampus_data = routes_df[routes_df['Krampus_ID'] != '-']
    
    for (krampus_id, tag, uhrzeit), group in krampus_data.groupby(['Krampus_ID', 'Tag', 'Uhrzeit']):
        group = group.sort_values('Besuchsreihenfolge')
        route_ids = ' → '.join([f"ID{int(x)}" for x in group['Kind_ID']])
        
        krampus_plans.append({
            'Krampus_ID': krampus_id,
            'Tag': tag,
            'Uhrzeit': uhrzeit,
            'Team': group.iloc[0]['Team_Nr'],
            'Nikolaus_Partner': group.iloc[0]['Nikolaus_ID'],
            'Route_IDs': route_ids,
            'Anzahl_Kinder': len(group),
            'Distanz_km': round(group['Distanz_zum_naechsten'].sum(), 2)
        })
    
    kram_df = pd.DataFrame(krampus_plans)
    kram_df.to_excel(f'{output_dir}/krampus_einsatzplaene.xlsx', index=False)
    print(f"✓ Krampus-Pläne gespeichert: krampus_einsatzplaene.xlsx")


# =============================================================================
# HAUPTFUNKTION
# =============================================================================

def main(input_file, output_dir='outputs', use_nominatim=True, 
         puffer_nikolaus=1, puffer_krampus=1):
    """
    Hauptfunktion - führt komplette Planung durch
    
    Parameters:
    -----------
    input_file : str
        Pfad zur Input-Excel-Datei
    output_dir : str
        Ausgabe-Verzeichnis
    use_nominatim : bool
        Ob Nominatim-Geocoding verwendet werden soll
    puffer_nikolaus : int
        Puffer für Nikoläuse
    puffer_krampus : int
        Puffer für Krampusse
    """
    
    print("\n" + "="*100)
    print("NIKOLAUS & KRAMPUS - AUTOMATISCHE ROUTENPLANUNG")
    print("="*100)
    print(f"Input: {input_file}")
    print(f"Output: {output_dir}/")
    print(f"Geocoding: {'Nominatim + Fallback' if use_nominatim else 'Nur Fallback'}")
    print("="*100)
    
    start_time = time.time()
    
    # 1. Lade Daten
    print("\n📖 Lade Daten...")
    df = pd.read_excel(input_file)
    
    # Normalisiere Krampus-Spalte
    if 'Krampus?' in df.columns:
        df['Krampus?'] = df['Krampus?'].astype(str).str.lower().str.strip()
    
    print(f"✓ {len(df)} Besuche geladen")

    # Normalisiere Spaltennamen (Case-Insensitive für Latitude/Longitude)
    df.rename(columns=lambda x: 'Latitude' if x.strip().lower() == 'latitude' else x, inplace=True)
    df.rename(columns=lambda x: 'Longitude' if x.strip().lower() == 'longitude' else x, inplace=True)
    
    # 2. Bedarfsanalyse
    demand_info = analyze_demand(df, puffer_nikolaus, puffer_krampus)
    
    # 3. Geocoding
    df_geo = geocode_addresses(df, use_nominatim=use_nominatim,
                               cache_file=f'{output_dir}/koordinaten_cache.csv')
    
    # 4. Route-Optimierung
    routes_df = create_full_route_plan(
        df_geo,
        demand_info['nikolaus_bedarf'],
        demand_info['krampus_bedarf']
    )
    
    # 5. Staff-Zuordnung
    routes_df = assign_staff_ids(
        routes_df,
        demand_info['nikolaus_bedarf'],
        demand_info['krampus_bedarf']
    )
    
    # 6. Visualisierungen
    create_visualizations(routes_df, output_dir)
    
    # 7. Interaktive Karte
    create_interactive_map(routes_df, output_dir)
    
    # 8. Excel-Export
    export_to_excel(routes_df, df_geo, demand_info, output_dir)
    
    # Finale Statistik
    elapsed_time = time.time() - start_time
    
    print("\n" + "="*100)
    print("✅ PLANUNG ABGESCHLOSSEN!")
    print("="*100)
    print(f"\n📊 Zusammenfassung:")
    print(f"   • Besuche: {len(df)}")
    print(f"   • Nikoläuse: {demand_info['nikolaus_bedarf']}")
    print(f"   • Krampusse: {demand_info['krampus_bedarf']}")
    print(f"   • Gesamtdistanz: {routes_df['Distanz_zum_naechsten'].sum():.2f} km")
    print(f"   • Laufzeit: {elapsed_time:.1f} Sekunden")
    
    print(f"\n📂 Ausgabe-Dateien in: {output_dir}/")
    print(f"   • zuordnung_komplett.xlsx")
    print(f"   • nikolaus_einsatzplaene.xlsx")
    print(f"   • krampus_einsatzplaene.xlsx")
    print(f"   • routenplan_interaktiv.html")
    print(f"   • statistik.png")
    print(f"   • routen_*.png")
    
    print("\n" + "="*100)


# =============================================================================
# KOMMANDOZEILEN-INTERFACE
# =============================================================================

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Automatische Routenplanung für Nikolaus & Krampus',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Beispiele:
----------
# Standard (mit Nominatim-Geocoding)
python nikolaus_planung_komplett.py --input besuche.xlsx

# Ohne Nominatim (schneller, aber ungenauer)
python nikolaus_planung_komplett.py --input besuche.xlsx --no-nominatim

# Mit angepassten Puffern
python nikolaus_planung_komplett.py --input besuche.xlsx --puffer-nikolaus 2 --puffer-krampus 2

# Mit anderem Output-Verzeichnis
python nikolaus_planung_komplett.py --input besuche.xlsx --output ergebnisse/
        """
    )
    
    parser.add_argument('--input', '-i', required=True,
                       help='Pfad zur Input-Excel-Datei')
    
    parser.add_argument('--output', '-o', default='outputs',
                       help='Ausgabe-Verzeichnis (Standard: outputs/)')
    
    parser.add_argument('--no-nominatim', action='store_true',
                       help='Nur Fallback-Geocoding verwenden (schneller)')
    
    parser.add_argument('--puffer-nikolaus', type=int, default=1,
                       help='Puffer für Nikoläuse (Standard: 1)')
    
    parser.add_argument('--puffer-krampus', type=int, default=1,
                       help='Puffer für Krampusse (Standard: 1)')
    
    args = parser.parse_args()
    
    # Prüfe Input-Datei
    if not os.path.exists(args.input):
        print(f"❌ Fehler: Datei nicht gefunden: {args.input}")
        sys.exit(1)
    
    # Führe Planung durch
    try:
        main(
            input_file=args.input,
            output_dir=args.output,
            use_nominatim=not args.no_nominatim,
            puffer_nikolaus=args.puffer_nikolaus,
            puffer_krampus=args.puffer_krampus
        )
    except Exception as e:
        print(f"\n❌ Fehler während der Ausführung:")
        print(f"   {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
