# How to Get the Tableau Extensions API File

## Important Note

**Tableau Desktop should automatically inject the Extensions API** when you load an extension. If you're seeing errors, it might be a Tableau Desktop version issue or configuration problem.

However, if you need the file manually, here are the options:

## Option 1: Download from GitHub (If Needed)

1. Go to: https://github.com/tableau/extensions-api
2. Click "Code" → "Download ZIP"
3. Extract the ZIP file
4. Navigate to the `lib` folder inside the extracted files
5. Copy `tableau-extensions-1.latest.min.js` to the `js/` folder of this extension

## Option 2: Use npm (if you have Node.js)

```bash
npm install @tableau/extensions-api
# Then copy the file from:
# node_modules/@tableau/extensions-api/lib/tableau-extensions-1.latest.min.js
# to: js/tableau-extensions-1.latest.min.js
```

## Option 3: Clone the Repository

```bash
git clone https://github.com/tableau/extensions-api.git
cd extensions-api/lib
cp tableau-extensions-1.latest.min.js /path/to/your/extension/js/
```

## Troubleshooting

If the API still doesn't load:
1. **Check Tableau Desktop version** - You need 2018.3 or later
2. **Check browser console** (F12) - Look for specific error messages
3. **Verify extension is added correctly** - Remove and re-add the extension to the dashboard
4. **Check Tableau logs** - Look in `My Tableau Repository/Logs` for extension-related errors

