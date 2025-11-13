# NetStatWiz - Network Statistics Wizard

A Python tool that analyzes network connections on Windows machines, visualizes IP locations on an interactive map, and generates detailed tables showing ports and services.

## Features

- **Network Analysis**: Runs `netstat` on Windows and parses connection data
- **IP Geolocation**: Maps IP addresses to their geographic locations
- **Interactive Map**: Generates an HTML map showing where connections are coming from
- **Port & Service Tables**: Creates detailed HTML tables with port and service information
- **Service Identification**: Automatically identifies common services by port number

## Installation

1. Install Python 3.7 or higher
2. Install required dependencies:

```bash
pip install -r requirements.txt
```

Or install manually:
```bash
pip install folium pandas
```

## Usage

Simply run the script:

```bash
python NetStatWiz.py
```

The program will:
1. Run `netstat -an` to get all network connections
2. Parse the output to extract external IP addresses and ports
3. Query geolocation data for each unique IP address (this may take a while)
4. Generate two output files:
   - `network_map.html` - Interactive map showing IP locations
   - `network_tables.html` - Detailed tables with all connection information

## Output Files

### network_map.html
An interactive map (using Folium/Leaflet) showing:
- Markers for each unique IP address location
- Popup information with IP, location, ISP, and connection count
- Clickable markers to view details

### network_tables.html
HTML tables showing:
- Summary statistics
- All connections with protocol, IP, port, service, state, and location
- Port and services summary grouped by port number

## Notes

- The program filters out localhost and private IP addresses
- Geolocation data is fetched from ip-api.com (free tier: 45 requests/minute)
- The program includes a delay between API calls to respect rate limits
- Common ports are automatically mapped to service names
- The program works without folium/pandas, but map generation will be disabled

## Requirements

- Windows OS (uses Windows `netstat` command)
- Python 3.7+
- Internet connection (for geolocation API)

## License

Free to use and modify.

