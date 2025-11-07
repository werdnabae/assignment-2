#!/usr/bin/env python3

"""
Excel to Folium Map Generator
Parses an Excel file with multiple tabs containing different types of map data
and creates an interactive folium map.

Expected Excel tab structure:
- 'markers': Points of interest (latitude, longitude, name, description, icon, color)
- 'lines': Routes/paths (name, coordinates as list of [lat,lon] pairs, color, weight)
- 'polygons': Areas/boundaries (name, coordinates as list of [lat,lon] pairs, color, fill_color)
- 'heatmap': Heatmap data (latitude, longitude, intensity)
- 'circles': Circular markers (latitude, longitude, radius, name, color)
"""

import pandas as pd
import folium
from folium.plugins import HeatMap, MarkerCluster
import json


def parse_coordinate_string(coord_str):
    """Parse coordinate string into list of [lat, lon] pairs."""
    if pd.isna(coord_str):
        return []
    try:
        # Handle JSON format: "[[lat1,lon1],[lat2,lon2]]"
        if isinstance(coord_str, str) and coord_str.strip().startswith('['):
            return json.loads(coord_str)
        # Handle semicolon-separated format: "lat1,lon1;lat2,lon2"
        elif isinstance(coord_str, str) and ';' in coord_str:
            pairs = coord_str.split(';')
            return [[float(x) for x in pair.split(',')] for pair in pairs]
        return []
    except:
        return []


def add_markers(map_obj, df_markers, use_cluster=True):
    """Add markers from dataframe to map."""
    if df_markers is None or df_markers.empty:
        return

    marker_cluster = MarkerCluster() if use_cluster else None

    for _, row in df_markers.iterrows():
        if pd.notna(row.get('latitude')) and pd.notna(row.get('longitude')):
            # Create popup content
            popup_html = f"<b>{row.get('name', 'Marker')}</b>"
            if pd.notna(row.get('description')):
                popup_html += f"<br>{row['description']}"

            # Create marker with custom icon if specified
            icon_kwargs = {}
            if pd.notna(row.get('icon')):
                icon_kwargs['icon'] = row['icon']
            if pd.notna(row.get('color')):
                icon_kwargs['color'] = row['color']
            else:
                icon_kwargs['color'] = 'blue'

            marker = folium.Marker(
                location=[row['latitude'], row['longitude']],
                popup=folium.Popup(popup_html, max_width=300),
                tooltip=row.get('name', 'Marker'),
                icon=folium.Icon(**icon_kwargs)
            )

            if use_cluster:
                marker.add_to(marker_cluster)
            else:
                marker.add_to(map_obj)

    if use_cluster:
        marker_cluster.add_to(map_obj)


def add_lines(map_obj, df_lines):
    if df_lines is None or df_lines.empty:
        return

    for _, row in df_lines.iterrows():
        coordinates = parse_coordinate_string(row.get('coordinates'))
        if coordinates:
            folium.PolyLine(
                locations=coordinates,
                color=row.get('color', 'red'),
                weight=row.get('weight', 3),
                opacity=row.get('opacity', 0.8),
                tooltip=row.get('name', 'Line'),
                popup=row.get('name', 'Line')
            ).add_to(map_obj)


def add_polygons(map_obj, df_polygons):
    """Add polygons/areas from dataframe to map."""
    if df_polygons is None or df_polygons.empty:
        return

    for _, row in df_polygons.iterrows():
        coordinates = parse_coordinate_string(row.get('coordinates'))
        if coordinates:
            folium.Polygon(
                locations=coordinates,
                popup=row.get('name', 'Area'),
                tooltip=row.get('name', 'Area'),
                color=row.get('color', 'blue'),
                fill=True,
                fill_color=row.get('fill_color', row.get('color', 'lightblue')),
                fill_opacity=row.get('fill_opacity', 0.4),
                weight=row.get('weight', 2)
            ).add_to(map_obj)


def add_heatmap(map_obj, df_heatmap):
    """Add heatmap layer from dataframe to map."""
    if df_heatmap is None or df_heatmap.empty:
        return

    heat_data = []
    for _, row in df_heatmap.iterrows():
        if pd.notna(row.get('latitude')) and pd.notna(row.get('longitude')):
            intensity = row.get('intensity', 1)
            heat_data.append([row['latitude'], row['longitude'], intensity])

    if heat_data:
        HeatMap(heat_data).add_to(map_obj)


def add_circles(map_obj, df_circles):
    """Add circle markers from dataframe to map."""
    if df_circles is None or df_circles.empty:
        return

    for _, row in df_circles.iterrows():
        if pd.notna(row.get('latitude')) and pd.notna(row.get('longitude')):
            popup_html = f"<b>{row.get('name', 'Circle')}</b>"
            if pd.notna(row.get('description')):
                popup_html += f"<br>{row['description']}"

            folium.Circle(
                location=[row['latitude'], row['longitude']],
                radius=row.get('radius', 500),
                popup=folium.Popup(popup_html, max_width=300),
                tooltip=row.get('name', 'Circle'),
                color=row.get('color', 'blue'),
                fill=True,
                fill_color=row.get('fill_color', row.get('color', 'lightblue')),
                fill_opacity=row.get('fill_opacity', 0.4),
                weight=row.get('weight', 2)
            ).add_to(map_obj)


def create_map_from_excel(excel_file, output_file='map.html',
                          center_lat=None, center_lon=None, zoom_start=10):
    """
    Create a folium map from an Excel file with multiple tabs.

    Parameters:
    -----------
    excel_file : str
        Path to Excel file
    output_file : str
        Output HTML file name
    center_lat : float, optional
        Center latitude (auto-calculated if not provided)
    center_lon : float, optional
        Center longitude (auto-calculated if not provided)
    zoom_start : int
        Initial zoom level (default: 10)
    """

    # Read all sheets from Excel file
    excel_data = pd.read_excel(excel_file, sheet_name=None)

    print(f"Found sheets: {list(excel_data.keys())}")

    # Extract dataframes for each tab
    df_markers = excel_data.get('markers')
    df_polygons = excel_data.get('polygons')
    df_heatmap = excel_data.get('heatmap')
    df_circles = excel_data.get('circles')

    # Calculate map center if not provided
    if center_lat is None or center_lon is None:
        all_lats = []
        all_lons = []

        for df in [df_markers, df_heatmap, df_circles]:
            if df is not None and not df.empty:
                if 'latitude' in df.columns:
                    all_lats.extend(df['latitude'].dropna().tolist())
                if 'longitude' in df.columns:
                    all_lons.extend(df['longitude'].dropna().tolist())

        if all_lats and all_lons:
            center_lat = sum(all_lats) / len(all_lats)
            center_lon = sum(all_lons) / len(all_lons)
        else:
            center_lat, center_lon = 0, 0  # Default to equator

    # Create base map
    m = folium.Map(
        location=[center_lat, center_lon],
        zoom_start=zoom_start,
        tiles='OpenStreetMap'
    )

    # Add different tile layers
    folium.TileLayer('CartoDB positron', name='Light Mode').add_to(m)
    folium.TileLayer('CartoDB dark_matter', name='Dark Mode').add_to(m)

    # Add map elements from each tab
    print("Adding markers...")
    add_markers(m, df_markers, use_cluster=True)

    print("Adding polygons...")
    add_polygons(m, df_polygons)

    print("Adding circles...")
    add_circles(m, df_circles)

    print("Adding heatmap...")
    add_heatmap(m, df_heatmap)

    print("Adding lines...")
    add_lines(m, excel_data.get('lines'))

    # Add layer control
    folium.LayerControl().add_to(m)

    # Fun feature: add glowing marker for night mode
    folium.CircleMarker(
        location=[center_lat, center_lon],
        radius=12,
        color='yellow',
        fill=True,
        fill_color='yellow',
        fill_opacity=0.6,
        popup="✨ Center Glow ✨"
    ).add_to(m)

    # Save map
    m.save(output_file)
    print(f"\nMap saved to: {output_file}")

    return m


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python excel_to_folium_map.py <excel_file> [output_file]")
        print("\nExpected Excel sheet structure:")
        print("  markers: latitude, longitude, name, description, icon, color")
        print("  lines: name, coordinates, color, weight, opacity")
        print("  polygons: name, coordinates, color, fill_color, fill_opacity, weight")
        print("  heatmap: latitude, longitude, intensity")
        print("  circles: latitude, longitude, radius, name, description, color, fill_color")
        sys.exit(1)

    excel_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else 'map.html'

    create_map_from_excel(excel_file, output_file)
