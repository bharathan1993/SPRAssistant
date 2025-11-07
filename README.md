# SPR Automation - Grafana Report Generator

A Google Apps Script add-on that automatically generates Google Slides presentations with Grafana charts.

## Overview

This tool pulls charts from Grafana dashboards and inserts them into Google Slides presentations. It uses Google Drive as an intermediary to handle slow-loading images and avoid timeout issues.

## Features

- **Automated Chart Insertion**: Fetches charts from Grafana and inserts them into slides
- **Sidebar Interface**: User-friendly form to configure report parameters
- **Drive-Based Approach**: Uses Google Drive as intermediary to handle large/slow images
- **Configurable Parameters**: Support for multiple data sources, environments, and time ranges
- **Error Handling**: Comprehensive logging and error messages for troubleshooting

## Setup

### 1. Configure Script Properties

Set your Grafana API key as a script property:

1. In Google Apps Script editor, go to **Project Settings** (gear icon)
2. Scroll to **Script Properties**
3. Add property: `GRAFANA_API_KEY` with your Grafana API key value

### 2. Configure Charts

Edit the `CHARTS_TO_PULL` array in `Code.js` to specify which charts to include:

```javascript
const CHARTS_TO_PULL = [
  { 
    dashboardUid: 'LUq13bv4z',        // Grafana dashboard UID
    panelId: 27,                       // Panel ID from the dashboard
    chartTitle: 'API Overall Volume',  // Title to display in the slide
    dashboardPath: 'api-health-overall-view'  // Dashboard path in Grafana
  }
  // Add more charts as needed
];
```

### 3. Update Grafana URL

Modify the `GRAFANA_URL` constant if using a different Grafana instance:

```javascript
const GRAFANA_URL = 'https://your-grafana-instance.com';
```

## Usage

### In Google Slides

1. Open or create a Google Slides presentation
2. Create a template slide with:
   - A text placeholder containing `{{CHART_TITLE}}`
   - A rectangle shape named `ImagePlaceholder` (or the largest rectangle will be used)
3. Click **Add-ons** → **Grafana Report Generator** → **Generate Report**
4. Fill in the sidebar form:
   - **Tenant ID**: Your tenant identifier
   - **Data Source**: Select the appropriate data source
   - **Data Center**: Choose the environment
   - **Interval**: Data aggregation interval
   - **Time Frame**: Time range for the charts
5. Click **Generate Report**

The script will:
- Clear existing slides (except the first template)
- Fetch each configured chart from Grafana
- Create slides with charts positioned according to your template

## Configuration Options

### Data Sources
- Pilot Telemetry(A1-Prod)
- Pilot Telemetry(A1-Sbx)
- Pilot Telemetry(EU-Prod)
- Pilot Telemetry(US-Prod)

### Data Centers
- US Production
- NA Production
- US Sandbox
- NA Sandbox
- NA Central Sandbox

### Time Frames
- Last 1 Hour
- Last 24 Hours
- Last 7 Days

### Intervals
- 1 minute
- 1 hour
- 1 day

## Testing

Use the `testWithDriveApproach()` function to test the connection to Grafana:

1. Open Google Apps Script editor
2. Select `testWithDriveApproach` from the function dropdown
3. Click **Run**
4. Check **Execution log** (View → Execution log) for results

## Troubleshooting

### 504 Gateway Timeout
- Try a shorter time range
- Simplify the Grafana panel query
- Contact your Grafana administrator

### HTML Instead of Image
- Verify dashboard variables are correctly mapped
- Check that panel ID exists in the dashboard

### No Placeholder Found
- Ensure your template slide has a rectangle shape
- Name the shape `ImagePlaceholder` for explicit targeting

### Authorization Errors
- Verify `GRAFANA_API_KEY` is set correctly in Script Properties
- Check API key has necessary permissions in Grafana

## Files

- **Code.js**: Main application logic and Grafana integration
- **Sidebar.html**: User interface for report configuration
- **appsscript.json**: Apps Script project configuration and permissions

## Required OAuth Scopes

- `presentations.currentonly` - Modify the current presentation
- `script.external_request` - Make requests to Grafana
- `script.container.ui` - Display the sidebar
- `drive.file` - Create temporary Drive files
- `drive` - Access Drive for image handling

## How It Works

1. **Fetch**: Downloads chart as PNG from Grafana render API
2. **Store**: Temporarily saves image to Google Drive
3. **Insert**: Inserts image from Drive into the slide
4. **Clean**: Deletes temporary Drive file after 2 seconds

This approach avoids direct URL insertion which can timeout on slow-loading charts.

## License

Internal tool for Zuora SPR automation.
