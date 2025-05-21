# Gmail Invoice Extractor

## Setup

1. Enable Gmail API at: https://console.cloud.google.com/apis
2. Download `credentials.json` and place it in this folder.
3. Run the following:

```bash
npm install googleapis cheerio xlsx
node index.js
```

4. The script will output:
   - `invoices_ws_display.json`
   - `invoices_ws_display.xlsx`

## OAuth2 Setup

The first time, you'll be prompted to authorize access. Follow the instructions printed in the console.
