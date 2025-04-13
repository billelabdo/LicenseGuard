# LicenseGuard
A robust PowerShell-based solution that tracks software license expiration dates, sends proactive email alerts, and generates detailed reports. Designed for IT teams and businesses to avoid compliance risks from unexpected license expirations.

## Features

- Monitors license expiration dates from Excel files
- Sends email alerts when licenses are expiring soon or have expired
- Generates HTML reports with license status overview
- Secure credential handling
- Automated scheduling capability
- Multi-language support (French by default)
- Comprehensive logging

## Prerequisites

- Windows PowerShell 5.1 or later
- Microsoft Excel (for reading license data)
- SMTP server access (for email alerts)
- Administrative privileges (for scheduling)

## Installation

1. Clone this repository
2. Configure the settings in `LicenseGuard.ps1`:
   - Set your email credentials and SMTP settings
   - Configure paths to your license Excel file
   - Adjust alert thresholds as needed

## 📌 Ideal For

IT Administrators managing enterprise software licenses

Compliance Teams ensuring legal usage of paid software

Small Businesses avoiding unexpected renewal costs

## ✨ Why Choose This Solution?

No External Dependencies – Uses native PowerShell and Excel (no databases needed)

Customizable – Adjust thresholds, emails, and reports via simple configs

Self-Hosted – Keeps sensitive license data on your infrastructure

## 📋 Sample Use Case

"A company with 50+ software subscriptions uses LicenseGuard to track renewals. The system alerts their IT manager 90 days before expiry, preventing a $20K AutoCAD license lapse."




## Usage

### Example: Daily automated execution
```powershell .\Schedule-LicenseGuard.ps1 -Time "09:00" -Force
### Manual Execution

```powershell
.\license-monitor.ps1


