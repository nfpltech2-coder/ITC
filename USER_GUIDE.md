# ITC Reconciliation Tool (Shakti) User Guide

## Introduction
The ITC Reconciliation Tool is a desktop application designed for Nagarkot Forwarders to reconcile GST data between **2B** (GST Portal) and **Books** (Internal Accounting). It automatically identifies mismatches, allows for bulk due-date assignment, and pushes data directly to the **Shakti** tracking system.

## How to Use

### 1. Launching the App
1. Locate the `itc_reco_app.exe` file.
2. Double-click to launch.
3. Ensure you have an active internet connection if you plan to push data to Shakti.

### 2. The Workflow (Step-by-Step)
1. **Upload Data**: Click **"Upload INPUT File"** and select your reconciliation Excel sheet.
   - *Constraint: The app specifically looks for a sheet containing "2B RECO" in the name.*
2. **Review Mismatches**: The table will automatically populate with records where `Status` is not "Matched".
3. **Filter & Toggle**: Use the sidebar to search for specific suppliers or toggle between "2B Done" and "Booking Done" records using the colored legend.
4. **Set Due Date**: Use the **Calendar Picker** in the sidebar to set the target completion date. This will immediately reflect on all records in the table.
5. **Generate Report**: Click **"Generate Sheet1 Output"** to save a local Excel copy of the mismatches.
6. **Push to Shakti**: Click **"Push to Shakti"**.
   - You can push **All** filtered records or **Select** specific rows in the table to push only those.

## Interface Reference

| Control / Input | Description | Expected Format |
| :--- | :--- | :--- |
| **Upload INPUT File** | Selects the source Excel file for processing. | .xlsx / .xls |
| **Filter by Supplier** | Searches the table by Supplier Name. | Text (Case-insensitive) |
| **Set Due Date** | Sets the Due Date for the cloud upload. | Calendar Selection (dd-MMM-yyyy) |
| **Type Legend** | Toggles visibility of 2B vs Books mismatches. | Clickable Color Tiles |
| **Push to Shakti** | Initiates bulk upload to the tracking system. | 100 records per batch |

## Troubleshooting & Validations

If you see an error, check this table:

| Message | What it means | Solution |
| :--- | :--- | :--- |
| "Could not find '2B RECO' sheet" | The Excel file doesn't have a tab named "2B RECO". | Rename the relevant tab in your Excel file to include "2B RECO". |
| "Missing 'Status' column" | The selected sheet is not in the expected format. | Ensure your sheet has a column header named exactly "Status". |
| "Zoho Refresh Token missing" | The credentials for Shakti are not set up. | Check your `.env` file and ensure the `ZOHO_REFRESH_TOKEN` is present. |
| "Failed to get Access Token" | Network error or invalid credentials. | Check your internet connection and verify Zoho Client ID/Secret. |
| "Record Error: [Message]" | Shakti rejected a specific record. | Usually caused by an invalid format or missing field in the cloud form. |

## Data Preparation Requirements
For successful processing, your input Excel file must follow these rules:
- **Sheet Name:** Must contain "2B RECO" (e.g., "2B RECO 25-26").
- **Columns Required:** 
  - `Supplier Name`
  - `Invoice Date`
  - `Invoice No.`
  - `Taxable Value`
  - `Status` (Should contain "Matched" for reconciled items).
  - `Origin` (Should contain "2B" or "BOOKS").
