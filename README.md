# ITC Reconciliation Tool (Shakti)

![Nagarkot Logo](assets/logo.png)

## Overview
The **ITC Reconciliation Tool** (internally known as **Shakti**) is a robust desktop application built for Nagarkot Forwarders to streamline the reconciliation of GST Input Tax Credit (ITC) between the **GST 2B Portal** and internal **Books of Accounts**.

The tool automates the process of identifying mismatched records (mismatches in origin, invoice numbers, or taxable values) and allows users to push these records in bulk to a cloud-based tracking system.

## Key Features
- 🚀 **Automated Mismatch Detection**: Instantly identifies records where Status is not 'Matched'.
- 📅 **Bulk Due Date Management**: Integrated calendar picker to assign due dates to hundreds of records at once.
- ☁️ **High-Speed Cloud Integration**: Optimized bulk push to Shakti using batched API requests.
- 🔍 **Live Filtering**: Search by supplier and toggle between 2B and Books records with a single click.
- 📊 **Excel Reporting**: Generate formatted reconciliation reports for offline use.

## Tech Stack
- **Language**: Python 3.10+
- **GUI Framework**: Tkinter / Custom Modern Components
- **Data Analysis**: Pandas, OpenPyXL
- **API Integration**: Requests (OAuth2 with Token Refresh)
- **Utilities**: tkcalendar, python-dotenv

## Installation & Setup

1. **Clone the repository**:
   ```bash
   git clone https://github.com/nfpltech2-coder/ITC.git
   ```

2. **Setup Virtual Environment**:
   ```bash
   python -m venv venv
   source venv/bin/activate  # venv\Scripts\activate on Windows
   ```

3. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Configure Environment**:
   Copy `.env.example` to `.env` and fill in your API credentials.

## Usage
For detailed instructions on how to use the application, please refer to the [USER_GUIDE.md](./USER_GUIDE.md).

## Project Structure
- `itc_reco_app.py`: Main application entry point.
- `assets/`: Contains UI branding elements (Logo, Icons).
- `requirements.txt`: List of required Python packages.
- `.gitignore`: Configured to ignore local data and environment secrets.

---
© 2026 Nagarkot Forwarders Pvt Ltd
