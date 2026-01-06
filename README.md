# Combined Data Export - Tableau Extension

This Tableau extension allows users to combine count and percentage data from multiple sheets into a single downloadable Excel file.

## Features

- Select two different sheets (one for counts, one for percentages)
- Automatically combines data based on common dimensions
- Preview the combined data before downloading
- Export to Excel format with proper formatting
- Maintains the exact structure as seen on the dashboard

## Installation

### Option 1: Local Development Server

1. Start a local web server in this directory:
   ```bash
   # Using Python 3
   python3 -m http.server 8000
   
   # Or using Node.js
   npx http-server -p 8000
   ```

2. In Tableau Desktop:
   - Go to **Extensions** → **Add Extension**
   - Enter the URL: `http://localhost:8000/manifest.trex`
   - Or browse to the `manifest.trex` file directly

### Option 2: Host on Web Server

1. Upload all files to a web server (HTTPS required for Tableau Server)
2. In Tableau Desktop/Server:
   - Go to **Extensions** → **Add Extension**
   - Enter the URL to your hosted `manifest.trex` file

## Usage

1. Add the extension to your dashboard (drag it from the Objects pane)
2. Select the sheet containing **count data** from the first dropdown
3. Select the sheet containing **percentage data** from the second dropdown
4. Click **"Preview Combined Data"** to see the combined result
5. Click **"Download as Excel"** to export the combined table

## How It Works

1. The extension loads all worksheets from your dashboard
2. You select which sheets contain count and percentage data
3. It fetches data from both sheets using the Tableau Extensions API
4. Data is merged by matching rows on the first column (dimension/stage name)
5. Count and percentage columns are combined side by side
6. The result is exported to Excel with proper formatting

## Requirements

- Tableau Desktop 2018.3+ or Tableau Server 2018.3+
- Internet connection (for CDN resources: Tableau Extensions API and SheetJS library)
- Both sheets must have a common dimension column (typically the first column)

## File Structure

```
exte/
├── manifest.trex          # Extension manifest file
├── index.html             # Main UI and styling
├── js/
│   └── extension.js       # Extension logic and data processing
└── README.md              # This file
```

## Troubleshooting

- **Extension not loading**: Make sure you're using HTTPS (for Tableau Server) or localhost (for Tableau Desktop)
- **No sheets available**: Ensure your dashboard has worksheets added to it
- **Data not combining correctly**: Verify both sheets have the same dimension column (first column) with matching values
- **Download not working**: Check browser console for errors and ensure SheetJS library is loading

## Notes

- The extension assumes the first column in both sheets is the dimension/stage name used for matching
- Count columns will be suffixed with " (Count)" and percentage columns with " (%)"
- The exported Excel file will be named with a timestamp for easy identification

