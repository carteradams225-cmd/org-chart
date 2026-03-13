# Org Chart Dashboard

A live, interactive org chart dashboard that reads from a local Excel file and serves a visual hierarchy in your browser. Anyone on the same network (or VPN) can open it via a shared link.

---

## Features

- Auto-refreshes when the Excel file is saved
- Dropdown to switch between any manager's org chart
- Toggle Job Title and Location on/off per box
- Export any chart as a standalone HTML file (shareable by email, no server needed)
- Color-coded hierarchy levels

---

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure your Excel path

Copy the example config and fill in your details:

```bash
cp config.example.json config.json
```

Then open `config.json` and set:

| Field | Description |
|---|---|
| `excel_path` | Full path to your Excel file (use forward slashes) |
| `sheet_name` | Sheet to read from — set to `null` to use the active sheet |
| `poll_interval_seconds` | How often to check for Excel changes (default: 5) |
| `port` | Port to run the server on (default: 5000) |

**Example:**
```json
{
  "excel_path": "C:/Users/YourName/Documents/employee_hierarchy.xlsx",
  "sheet_name": null,
  "poll_interval_seconds": 5,
  "port": 5000
}
```

### 3. Run the server

```bash
python server.py
```

You'll see output like:
```
Org Chart Dashboard running at:
  Local:   http://localhost:5000
  Network: http://<YOUR-IP-ADDRESS>:5000

Watching: C:/Users/YourName/Documents/employee_hierarchy.xlsx
```

### 4. Open the dashboard

Open `http://localhost:5000` in your browser.

---

## Sharing with colleagues

To share with others on your network or VPN:

1. Find your machine's IP address:
   - **Windows:** Open Command Prompt → type `ipconfig` → look for `IPv4 Address`
2. Share the link: `http://YOUR-IP-ADDRESS:5000`
3. They open it in any browser — no install needed on their end.

> The server must be running on your machine for others to access it.

---

## Exporting a static snapshot

Click the **⬇ Export** button in the dashboard to download a self-contained HTML file.
This file can be emailed or shared — it opens in any browser with no server required.

---

## Excel file requirements

The dashboard reads these columns automatically (column names are case-sensitive):

| Column | Required | Notes |
|---|---|---|
| `Employee_Name` | Yes | Full name of the employee |
| `Mgr_Emp_Name_L1` … `Mgr_Emp_Name_L6` | Yes | Management chain columns |
| `JobTitle` | No | Shown as subtitle in each box |
| `Location_Name` | No | Shown with a 📍 icon |

The dashboard auto-detects all `Mgr_Emp_Name_L*` columns. No configuration needed for column names.

---

## Switching between org charts

The **Manager** dropdown in the header is automatically populated with every person in the Excel file who manages at least one other person. Select any name to view their org chart.

---

## Notes

- `config.json` and all Excel files are in `.gitignore` — they will never be committed to GitHub.
- The server watches the Excel file every few seconds. Just save the file and the dashboard will refresh automatically.
