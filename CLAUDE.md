# Org Chart Dashboard — Project Log

## What This Project Is
An interactive org chart dashboard that reads employee hierarchy data from a local Excel file and renders it visually in a browser. Served via a local Flask server accessible to anyone on the same VPN/network.

---

## Decisions Made

### Data Source
- Excel file housed on **SharePoint**, but accessed **locally** for now (downloaded to work laptop)
- We could not register an Azure AD app (no IT permissions), so SharePoint live API connection is not currently set up
- Future goal: connect directly to SharePoint via Microsoft Graph API for true live updates without manual file downloads

### Hosting
- Server runs on user's **work laptop** (`python server.py`)
- Shared via local IP address — anyone on the **VPN** can access `http://YOUR-IP:5000`
- Limitation: dashboard goes down when the laptop is off
- Future option: move server to an always-on shared machine or VM on the network

### Sharing
- Short term: use the **⬇ Export** button to download a standalone HTML snapshot for emailing
- Long term: cloud hosting (Railway/Render) + SharePoint API for 24/7 live access

### GitHub
- Repo: `https://github.com/carteradams225-cmd/org-chart` (public)
- Made public so it can be cloned on work laptop without signing into personal GitHub
- `config.json` and all `*.xlsx` files are in `.gitignore` — no company data ever touches GitHub
- No Git installed on work laptop — user downloads ZIP from GitHub instead

### Excel Columns Used
- `Employee_Name` — required
- `Mgr_Emp_Name_L1` through `Mgr_Emp_Name_L6` — required, auto-detected dynamically
- `JobTitle` — optional, shown as subtitle in each box
- `Location_Name` — optional, shown with 📍 icon
- Direct manager logic: last non-None entry in the Mgr chain that is not the employee themselves

### Multiple Org Charts
- No names are hardcoded — everything is read dynamically from the Excel file
- The manager dropdown is auto-populated with every person who has at least one direct report
- Works for any root manager in the same Excel sheet (not just Mark Wilhelm)

### Visualization Layout
- Root manager at the top
- Root's direct reports in a **horizontal row**
- All subsequent levels in **vertical columns** under their manager
- SVG lines connect parent to children (bus-line style)
- Color-coded by depth level (dark → light blue)

---

## Current Status

### Done
- `server.py` — Flask server, reads Excel, serves JSON, polls for file changes every 5 seconds
- `org_chart.html` — Full dashboard with dropdown, job title/location toggles, export button, auto-refresh
- `config.example.json` — Template config (user fills in Excel path locally)
- `.gitignore` — Blocks config.json and all Excel files from GitHub
- `requirements.txt` — `flask`, `openpyxl`
- `README.md` — Full setup and usage instructions
- Pushed to GitHub at `carteradams225-cmd/org-chart`

### In Progress / Blockers
- Work laptop does not have Git installed — user downloads ZIP from GitHub instead
- Work laptop uses Windows Command Prompt — `cp` command does not work, use `copy` instead:
  ```
  copy config.example.json config.json
  ```
- User is currently at the step of creating `config.json` on their work laptop

### Planned / Future Work
- **Visualization enhancements** — the current dashboard is functional but will likely need improvements to the visual design, layout, and interactivity (specifics TBD)
- **SharePoint live connection** — requires Azure AD app registration (needs IT involvement)
- **Always-on hosting** — move server to a shared machine or cloud host so dashboard is available when user's laptop is off
- **Export improvements** — static HTML export works but could be enhanced

---

## How to Run (Work Laptop)

1. Download ZIP from `https://github.com/carteradams225-cmd/org-chart` → Code → Download ZIP
2. Unzip somewhere (e.g. `Documents\org-chart`)
3. Open Command Prompt in that folder (click address bar in File Explorer, type `cmd`, hit Enter)
4. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
5. Create config file (**use `copy` not `cp` on Windows**):
   ```
   copy config.example.json config.json
   ```
6. Open `config.json` and set `excel_path` to your Excel file location (use forward slashes)
7. Run the server:
   ```
   python server.py
   ```
8. Open browser at `http://localhost:5000`

---

## Notes
- Always use `copy` instead of `cp` in Windows Command Prompt
- Use forward slashes in file paths inside `config.json` (e.g. `C:/Users/...`)
- `config.json` must never be committed to GitHub
