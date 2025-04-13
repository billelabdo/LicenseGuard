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

## Usage

### Manual Execution

```powershell
.\license-monitor.ps1
