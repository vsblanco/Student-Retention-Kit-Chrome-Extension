# Student Retention Kit

**Version 15.0** | Last Updated: December 2025

## Overview

The Student Retention Kit is a comprehensive Chrome extension designed to help educational institutions track student engagement, identify at-risk students, and streamline outreach efforts through automated workflows.

## Core Features

### 📊 Submission Checker
- **Real-time Monitoring**: Automatically scans Canvas LMS for student submissions
- **API-Powered**: Uses Canvas API for fast, reliable data retrieval
- **Smart Filtering**: Configurable filters based on days out, failing grades, and more
- **Visual Indicators**: Color-coded heatmap showing student engagement levels

### 📞 Call Management
- **Five9 Integration**: Direct calling from the extension interface
- **Automation Mode**: Queue-based calling for multiple students
- **Call Disposition**: Track call outcomes and notes
- **Student Context**: View student details during calls

### 📁 Data Management
- **Master List**: Centralized student database with comprehensive information
- **Multiple Import Methods**: CSV, Excel, JSON clipboard, and Excel Add-in sync
- **Smart Search & Sort**: Quickly find students by name, missing count, or days out
- **Export Functionality**: Download reports for offline analysis

### 🔗 Integrations

**Excel Add-in**
- Bi-directional sync with Office Add-in
- Automatic master list updates
- Student row highlighting on submissions
- Missing assignments import

**Power Automate**
- Webhook-based automation
- Custom workflow triggers
- Real-time data push

**Canvas LMS**
- Direct API integration
- Assignment tracking
- Grade monitoring
- Embedded helper interface

**Five9 VoIP**
- Click-to-call functionality
- Call status monitoring
- Disposition tracking

## Data Cleaning Features

- **Program Version Cleanup**: Automatically removes year prefixes (e.g., "D2-Q-2024 Nursing" → "Nursing")
- **Excel Date Conversion**: Converts Excel serial dates to MM-DD-YY format
- **Field Normalization**: Handles various field name variations seamlessly

## Recent Updates

### v15.0 (Current)
- Added ProgramVersion data cleaning
- Enhanced Excel Add-in integration
- Improved hidden field payload support

### v14.4
- Added Five9 integration
- Enhanced call interface
- Daily update modal improvements

### v14.0
- Excel Add-in bi-directional sync
- Student selection automation
- Taskpane ping check

---

**© 2025 Student Retention Kit** | Built for Student Services Teams
