# QBO → Google Sheets Integration

Automatically pulls live financial data from QuickBooks Online (QBO) 
into Google Sheets — giving the operations team real-time visibility 
without needing direct access to QuickBooks.

## The Problem

Exporting financial reports from QuickBooks Online was a manual process 
that required accounting access and produced static snapshots. 
This integration replaced that with an automated, always-current feed 
directly inside Google Sheets.

## What It Does

- Connects to QuickBooks Online via OAuth2 (production environment)
- Pulls financial data using the QBO Reports API and Query API (QQL)
- Writes results directly into a Google Sheet for team visibility
- Supports custom queries for specific reports (e.g. open Purchase Orders)

## How It Works
```
Google Sheets (Apps Script)
        ↓
OAuth2 authentication → QuickBooks Online API
        ↓
QBO Reports API  /  QBO Query API (QQL)
        ↓
Data written back into Google Sheets
```

## Tech Stack

- Google Apps Script
- QuickBooks Online API (Reports + QQL)
- apps-script-oauth2 library
- Google Sheets

## Security

No credentials are hardcoded. All API keys and secrets are stored 
using Google Apps Script's `PropertiesService` and read at runtime.

## Status

Live in production, connected to a real QBO account.
