# Excel to Folium Map Generator

A Python script that parses an Excel spreadsheet with multiple tabs containing different types of map data and creates an interactive Folium map.

👉 [Open the Interactive Map](https://werdnabae.github.io/assignment-2/map.html)

## Example Output
- [Map w/ Lines](map_with_lines.html)

## Features

- **Markers**: Point locations with custom icons and colors
- **Lines**: Routes and paths between locations
- **Polygons**: Areas and boundaries
- **Circles**: Circular regions with custom radius
- **Heatmap**: Intensity-based visualization
- **Marker Clustering**: Automatically groups nearby markers
- **Multiple Map Layers**: Switch between different map styles

## Installation

Install required packages:

```bash
pip install -r requirements.txt
```

## Usage

```bash
python assignment2.py <excel_file> [output_file]
```

Example:
```bash
python assignment2.py my_data.xlsx my_map.html
```

If no output file is specified, it defaults to `map.html`.

## Excel File Format

Your Excel file should contain one or more of the following sheets:

### 1. markers (Point Locations)

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| latitude | float | Yes | Latitude coordinate |
| longitude | float | Yes | Longitude coordinate |
| name | string | No | Marker name/title |
| description | string | No | Popup description |
| icon | string | No | Icon type (e.g., 'info-sign', 'star', 'home') |
| color | string | No | Marker color (blue, red, green, purple, orange, etc.) |

**Example:**
```
latitude  | longitude  | name          | description      | icon       | color
40.7128   | -74.0060   | New York City | The Big Apple    | info-sign  | blue
34.0522   | -118.2437  | Los Angeles   | City of Angels   | star       | red
```

### 2. lines (Routes/Paths)

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| name | string | No | Line name |
| coordinates | string | Yes | JSON array of [lat,lon] pairs |
| color | string | No | Line color (default: blue) |
| weight | int | No | Line thickness (default: 3) |
| opacity | float | No | Line opacity 0-1 (default: 0.8) |

**Coordinate Formats:**
- JSON: `[[40.7128,-74.0060],[34.0522,-118.2437]]`
- Semicolon: `40.7128,-74.0060;34.0522,-118.2437`

**Example:**
```
name              | coordinates                                           | color | weight | opacity
NYC to LA         | [[40.7128,-74.0060],[34.0522,-118.2437]]             | blue  | 3      | 0.7
Chicago to Miami  | [[41.8781,-87.6298],[25.7617,-80.1918]]              | red   | 4      | 0.8
```

### 3. polygons (Areas/Boundaries)

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| name | string | No | Polygon name |
| coordinates | string | Yes | JSON array of [lat,lon] pairs (must close: first=last) |
| color | string | No | Border color (default: blue) |
| fill_color | string | No | Fill color (default: lightblue) |
| fill_opacity | float | No | Fill opacity 0-1 (default: 0.4) |
| weight | int | No | Border thickness (default: 2) |

**Example:**
```
name           | coordinates                                                      | color     | fill_color  | fill_opacity | weight
Northeast Area | [[40.7128,-74.0060],[42.3601,-71.0589],[40.7128,-74.0060]]     | darkblue  | lightblue   | 0.3          | 2
```

### 4. circles (Circular Regions)

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| latitude | float | Yes | Center latitude |
| longitude | float | Yes | Center longitude |
| radius | float | Yes | Radius in meters |
| name | string | No | Circle name |
| description | string | No | Popup description |
| color | string | No | Border color (default: blue) |
| fill_color | string | No | Fill color (default: lightblue) |
| fill_opacity | float | No | Fill opacity 0-1 (default: 0.4) |
| weight | int | No | Border thickness (default: 2) |

**Example:**
```
latitude | longitude  | radius | name      | color | fill_color  | fill_opacity
40.7128  | -74.0060   | 50000  | NYC Metro | blue  | lightblue   | 0.2
```

### 5. heatmap (Intensity Data)

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| latitude | float | Yes | Latitude coordinate |
| longitude | float | Yes | Longitude coordinate |
| intensity | float | No | Heat intensity value (default: 1) |

**Example:**
```
latitude | longitude  | intensity
40.7128  | -74.0060   | 0.8
40.7500  | -73.9900   | 1.0
40.6900  | -74.0200   | 0.5
```

## Available Icon Types

Common Folium/Font Awesome icons:
- `info-sign`
- `home`
- `star`
- `plane`
- `cloud`
- `heart`
- `flag`
- `music`
- `camera`
- `shopping-cart`

## Available Colors

- `red`
- `blue`
- `green`
- `purple`
- `orange`
- `darkred`
- `lightred`
- `beige`
- `darkblue`
- `darkgreen`
- `cadetblue`
- `darkpurple`
- `white`
- `pink`
- `lightblue`
- `lightgreen`
- `gray`
- `black`
- `lightgray`

## Output

The script generates an interactive HTML map with:
- Zoom and pan controls
- Layer switcher for different map styles
- Clickable markers with popups
- Legend for different elements

## Notes

- Not all sheets are required - include only the types you need
- The map will auto-center based on your data
- Marker clustering is enabled by default for better performance
- All coordinate pairs should be in [latitude, longitude] format
- For polygons, ensure the first and last coordinate pairs are identical to close the shape

## Example Data

See `sample_map_data.xlsx` for a complete example with all sheet types populated.

## New Features
- **Lines Layer:** Added support for drawing routes or paths from the `lines` sheet in the Excel file.  
  Each line should include: `name`, `coordinates`, `color`, `weight`, and `opacity`.

- **Center Glow (Fun Feature):** Added a glowing yellow marker that highlights the map’s center — just for fun!