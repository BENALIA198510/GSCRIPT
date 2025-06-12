# Field Training Management System

This repository contains the Google Apps Script source code and HTML interface for a field training management system. The system stores its data in Google Sheets and exposes a web app for administrators and regular users to track field training records.

## Features

- **User authentication** with registration and login
- **Password recovery** via one‑time password (OTP)
- **CRUD operations** for training records stored in the `Data` sheet
- **Summary statistics** and filtering options
- **Export to PDF** for selected records
- **Admin and user roles** to control access

## Project Structure

```
Code.gs     # Server‑side Apps Script (functions and API)
Index.html  # Front‑end HTML/JS for the web app interface
```

### Customization

All data is stored in three Google Sheets:

- `Login` — user accounts
- `Data` — training records
- `Options` — dropdown values for filtering

Adjust sheet names in `Code.gs` if needed. The `setupApplication` function creates these sheets with default headers and adds an initial admin user (`admin@example.com` / `admin123`).

## Deployment

The code can be deployed to Google Apps Script using [clasp](https://github.com/google/clasp) or by manually copying the files.

### Using clasp

1. Install clasp and log in:
   ```bash
   npm install -g @google/clasp
   clasp login
   ```
2. Create a new Apps Script project and push the files:
   ```bash
   clasp create --title "Field Training Management"
   clasp push
   ```
3. Open the script in the Apps Script editor to configure the deployment.

### Manual copy

1. Create a new Apps Script project from [script.google.com](https://script.google.com).
2. Replace the default `Code.gs` and `Index.html` with the contents of this repository.

## Initial Setup

After importing the files, run the `setupApplication` function once from the Apps Script editor to create the required sheets and a sample admin account:

```javascript
function setupApplication() {
  // Creates sheets and a default admin user
}
```

## Quickstart

1. **Open** the Apps Script project (via clasp or manual copy).
2. **Run** `setupApplication` to initialize the sheets and sample admin.
3. **Deploy** the project as a web app (New Deployment → Web app). Choose *Execute as Me* and allow access to **Anyone**.
4. **Open** the deployment URL in your browser and log in using the sample admin credentials.

You can now add users, manage training records, and export data as needed.

