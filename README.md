# 🎰 Lucky Draw — Digital Raffle Fundraising System

A complete raffle fundraising system with registration, animated draw, admin panel, and Google Sheets as the database.

---

## 📁 Project Files

| File | Description |
|---|---|
| `registration.html` | Participant registration form |
| `raffle-draw.html` | Animated raffle draw page |
| `admin.html` | Admin ticket checker & search |
| `Code.gs` | Google Apps Script backend |
| `README.md` | This guide |

---

## 🚀 Quick Setup (Step-by-Step)

### Step 1 — Create a Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com) and create a new spreadsheet.
2. Name it **"Lucky Draw Raffle"** (or any name you like).
3. Copy the **Sheet ID** from the URL bar:
   ```
   https://docs.google.com/spreadsheets/d/THIS_IS_YOUR_SHEET_ID/edit
   ```

---

### Step 2 — Set Up the Apps Script Backend

1. Inside your Google Sheet, click **Extensions → Apps Script**.
2. Delete all existing code in `Code.gs`.
3. Copy and paste the entire content of `Code.gs` from this project.
4. On **line 17**, replace `YOUR_GOOGLE_SHEET_ID` with your actual Sheet ID:
   ```javascript
   const SHEET_ID = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms'; // example
   ```
5. Click **Save** (Ctrl+S / Cmd+S).

---

### Step 3 — Initialise the Sheets (One-Time)

1. In the Apps Script editor, select `initSheets` from the function dropdown.
2. Click **▶ Run**.
3. Approve the permissions when prompted (this allows the script to access your spreadsheet).
4. You should now see two tabs in your spreadsheet: **Registrations** and **Winners**.

---

### Step 4 — Deploy the Web App

1. Click **Deploy → New Deployment**.
2. Click the gear icon ⚙ next to **Type**, select **Web App**.
3. Fill in the settings:
   - **Description:** `Lucky Draw API v1`
   - **Execute as:** `Me`
   - **Who has access:** `Anyone`
4. Click **Deploy**.
5. **Copy the Web App URL** — it looks like:
   ```
   https://script.google.com/macros/s/AKfy.../exec
   ```

> ⚠️ **Important:** Every time you edit `Code.gs`, you must create a **New Deployment** (not "Manage existing deployment") to apply changes.

---

### Step 5 — Connect the Frontend

Open each of the three HTML files and replace `YOUR_WEBAPP_URL` with your Web App URL:

**In `registration.html`** (near the bottom, inside `<script>`):
```javascript
const WEBAPP_URL = 'https://script.google.com/macros/s/YOUR.../exec';
```

**In `raffle-draw.html`** (inside `<script>`):
```javascript
const WEBAPP_URL = 'https://script.google.com/macros/s/YOUR.../exec';
```

**In `admin.html`** (inside `<script>`):
```javascript
const WEBAPP_URL = 'https://script.google.com/macros/s/YOUR.../exec';
```

---

### Step 6 — Host the HTML Files (Optional)

The HTML files can be opened directly from your computer for local testing. For public access, host them on any static file service:

- **Google Sites** — free, simple
- **GitHub Pages** — free, version-controlled
- **Netlify / Vercel** — free, fast CDN
- **Your own web host** — upload via cPanel or FTP

---

## 📋 How to Use the System

### Registration Page (`registration.html`)
1. Participants fill in their name, email, phone number.
2. They select how many tickets they want (₱100 per ticket).
3. They choose their payment method and enter their reference number.
4. Optionally upload a screenshot of their payment receipt.
5. On submission, a unique **Ticket ID** is generated (e.g., `TKT-20250521-A3BX7`).
6. The participant's details are saved to the **Registrations** tab in Google Sheets.

### Raffle Draw Page (`raffle-draw.html`)
1. Click **"⬇ Load Participants"** to fetch all registered tickets from Google Sheets.
2. The spinning ball machine shows all participants loaded.
3. Click **"🎯 Draw Winner!"** to start the draw.
4. Watch the balls spin and slow down — the selected ball drops into the chute.
5. The winner's name and ticket number are revealed with a **confetti animation**.
6. The winner is automatically saved to the **Winners** tab in Google Sheets.
7. You can draw multiple winners (e.g., 1st prize, 2nd prize) — each is logged separately.

### Admin Panel (`admin.html`)
1. The page automatically loads all records from Google Sheets on open.
2. Use the **search bar** to find any participant by:
   - Full name or partial name
   - Email address
   - Ticket ID (e.g., `TKT-20250521-A3BX7`)
   - Reference / transaction number
   - Phone number
3. Use the **filter tabs** to view: All, Winners only, or by payment method.
4. Each card shows full participant details and is highlighted in gold if they won.
5. Statistics at the top show total tickets, participants, winners, and revenue.

---

## 📥 Importing Existing Data from Excel

If you already have registrations in an Excel spreadsheet, you can import them directly.

### Method A: Copy & Paste into Google Sheets (Easiest)

1. Open your Excel file.
2. Copy all the data rows.
3. Open your Google Sheet → **Registrations** tab.
4. Click on row 2 (below the header) and paste.
5. Make sure columns match the header order:
   `Ticket ID | Full Name | First Name | Last Name | Email | Phone | Tickets Qty | Price Each | Total Amount | Payment Method | Reference No | Receipt File | Has Receipt | Registered At | Status | Winner Round | Notes`
6. For any blank **Ticket ID** cells, run the helper below.

### Method B: Upload via Admin Panel (Frontend Import)

Add the following import button to `admin.html` for a drag-and-drop experience. The frontend reads the Excel file using the [SheetJS library](https://sheetjs.com/) and sends the parsed rows to the Apps Script backend:

```html
<!-- Add this to admin.html -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<input type="file" id="excelFile" accept=".xlsx,.xls,.csv"
  onchange="importExcel(this)"/>

<script>
async function importExcel(input) {
  const file = input.files[0];
  const data = await file.arrayBuffer();
  const wb = XLSX.read(data);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws);

  const resp = await fetch(WEBAPP_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ action: 'importExcel', rows }),
    mode: 'cors',
  });
  const result = await resp.json();
  alert(result.message);
}
</script>
```

### Expected Excel Column Names (flexible)

The import function recognises these column names (case-insensitive):

| Accepted Column Name | Field |
|---|---|
| `name`, `Name`, `Full Name` | Participant name |
| `email`, `Email` | Email address |
| `phone`, `Phone` | Phone number |
| `paymentMethod`, `Payment Method`, `payment` | Payment type |
| `referenceNo`, `Reference No`, `reference` | Transaction reference |
| `ticketQty`, `tickets`, `Tickets` | Number of tickets |

---

## 🛠 Customisation

### Change the Ticket Price
In `Code.gs`, line 18:
```javascript
const TICKET_PRICE = 100; // Change to your desired price
```
Also update in `registration.html`:
```javascript
const TICKET_PRICE = 100; // Match this value
```

### Change Currency Symbol
Search for `₱` in the HTML files and replace with your currency symbol.

### Change the Raffle Name
Search for `Grand Raffle Draw 2025` in all HTML files and update to your event name.

### Add More Payment Methods
In `registration.html`, add `<option>` tags to the payment method `<select>`:
```html
<option value="BDO">BDO Bank</option>
<option value="BPI">BPI Bank</option>
```

---

## 🔒 Security Notes

- The Web App runs as **you** (the owner). Participants cannot modify the Google Sheet directly — all writes go through the Apps Script which validates input.
- For a public-facing event, consider adding a **CAPTCHA** or **rate limiting** to the registration form.
- The admin panel (`admin.html`) has no login. For real events, either:
  - Host it on a private URL
  - Add a simple password check in JavaScript
  - Use Google's built-in Apps Script HTML Service with login restrictions

---

## ❓ Troubleshooting

| Problem | Solution |
|---|---|
| "Demo mode" message | The `WEBAPP_URL` is not set or the script isn't deployed. |
| CORS error in browser console | Redeploy the script as **"Anyone"** with access. |
| Registration saves but no data in sheet | Check `SHEET_ID` in `Code.gs` is correct. |
| Approval/permissions error | Run `initSheets()` manually first to grant permissions. |
| `initSheets` ran but no tabs created | Make sure `SHEET_ID` is the correct spreadsheet. |
| Winner not saving | Check the Web App is deployed and `WEBAPP_URL` is correct in `raffle-draw.html`. |

---

## 📞 Support

This project was built with HTML, vanilla JavaScript, and Google Apps Script. No additional libraries or npm packages are required for the core system.

For Excel import, the optional [SheetJS (xlsx)](https://sheetjs.com/) library is loaded from a CDN.
