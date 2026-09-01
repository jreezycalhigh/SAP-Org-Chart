# SAP US Account Priority Technical Outreach & Org Chart

An interactive outreach management application, org chart, and verification tool designed for tracking SAP US Account contacts, mapping their priority levels, scheduling follow-ups, and managing client-facing roles.

## 🚀 Live Interactive Portal (Real-Time)

You can interact with the live dashboard and interactive family-tree org chart directly through your web browser via **GitHub Pages**. 

*Once GitHub Pages is enabled on this repository, the link will be:*
**`https://<your-github-username>.github.io/<your-repo-name>/SAP_Account_Priority_Outreach_Hub.html`**

---

## 🛠️ Main Features & Contents

### 1. 🗂️ Interactive Org Chart & Outreach Hub (`SAP_Account_Priority_Outreach_Hub.html`)
The primary interface containing:
- **Interactive Family Tree Org Chart**: A zoomable, pannable, and collapsible visualization of the SAP Executive Board and product branches (Autonomous Suite, SuccessFactors, Concur, NS2, T&D, Security, etc.).
  - **- / + Symbols**: Click to expand or collapse portions of the org chart.
  - **Details Card Modal**: Click any card in the tree to view full notes, scheduled weeks, associated products, and direct IBM Owners.
- **Priority Call List**: A list of all 445 contacts sortable and searchable by Seller, LOB, and Priority Level.
- **Call Schedule**: A weekly calendar engine that automatically maps and spreads outreach dates based on target deadlines and priority rules.

### 2. 🔍 Contact Verification Checklist (`SAP_Verification_Checklist.html`)
A standalone helper interface mapping all 445 contacts to check their live employment details. It includes direct links to:
- 🔍 **Google Search** for the contact name at SAP.
- **in LinkedIn** search results for the contact at SAP.
- **ZI ZoomInfo** direct lookup.

### 3. 📊 Source Data Spreadsheet (`SAP_Account_Priority_Outreach_List_v16.xlsx`)
The master Excel sheet of contact records containing names, titles, organizations, priority rankings, notes, and IBM account owners.

### 4. ⚙️ Python HTML Compiler (`generate_final_html.py`)
A custom automation script that reads raw data directly from `SAP_Account_Priority_Outreach_List_v16.xlsx` and compiles it into the high-performance, responsive `SAP_Account_Priority_Outreach_Hub.html` portal.

---

## ✏️ How to Edit via the Web GUI (No coding required)

This repository is configured so that team members can edit data and update the portal completely through their web browser without installing anything locally:

### Option A: Edit directly on GitHub.com (Quick File Edits)
1. Go to your repository on GitHub.
2. Navigate to the file you want to edit (e.g., `SAP_Account_Priority_Outreach_Hub.html`).
3. Click the **Edit (pencil icon)** in the top right corner.
4. Make your edits and click **Commit changes...** at the top right to save.
5. GitHub Pages will redeploy and update the live site within seconds!

### Option B: The GitHub Web-Based VS Code Editor (Full GUI IDE)
1. While viewing your repository main page on GitHub, press the **`.` (dot) key** on your keyboard, or change the URL from `github.com/...` to `github.dev/...`.
2. This opens a fully-featured, web-based visual editor (Visual Studio Code for Web) directly in your browser.
3. You can modify any code, data arrays, or scripts.
4. Go to the **Source Control** tab on the left sidebar (shortcut: `Ctrl+Shift+G`), write a commit message, and click **Commit & Push**.
5. The live site will automatically rebuild and deploy!

---

## 💻 Local Development & Compiling

If you prefer to edit the contacts using Microsoft Excel, follow these simple steps to recompile the HTML:

1. Open `SAP_Account_Priority_Outreach_List_v16.xlsx` in Excel and update contacts, priorities, or notes. Save and close the file.
2. Open terminal in this folder and run:
   ```powershell
   python generate_final_html.py
   ```
3. The script will read your spreadsheet and instantly update `SAP_Account_Priority_Outreach_Hub.html`.
4. Commit and push the files to GitHub:
   ```bash
   git add .
   git commit -m "Update contacts from Excel spreadsheet"
   git push origin main
   ```

---
*Made with IBM Bob*
