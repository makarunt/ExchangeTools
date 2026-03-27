# Get-ExchangeConnectorTraffic
A PowerShell script for analyzing Microsoft Exchange SMTP Receive connector protocol logs. It identifies and reports actual mail traffic — showing sender address, recipient(s), source IP, and connector name — for any configurable time window.

## Overview
Exchange writes a new SMTP Receive Protocol log file every hour. When troubleshooting mail flow or auditing which systems are sending mail through your receive connectors, manually inspecting these logs is time-consuming. This script automates that process by parsing the raw log files, correlating SMTP session data, and producing both a detailed report and an executive summary grouped by server, connector, and source IP.

## Requirements
PowerShell 5.0 or higher
Read access to the Exchange SMTP Receive Protocol log folder
Microsoft Exchange Server 2013 / 2016 / 2019 (log format version 15.x)
## Parameters
Parameter	Required	Description
-LogPath	No	Path to the folder containing .log files. Defaults to the standard Exchange V15 SmtpReceive log path.
-Hours	No	Number of hours back to analyse. Cannot be combined with -Days. Defaults to 5 hours if neither is specified.
-Days	No	Number of days back to analyse. Cannot be combined with -Hours.
-Connector	No	Filter results to a specific connector name (partial, case-insensitive match).
-ExcludeIP	No	One or more IP addresses to exclude from the report (e.g. known Edge Transport servers).
-ExportCsv	No	Full path to export detailed results as a CSV file. A matching _summary.txt file is created automatically in the same folder.
## Usage

```powershell
### Run with default settings (last 5 hours, default log path)
.\Get-ExchangeConnectorTraffic.ps1

### Analyse the last 12 hours
.\Get-ExchangeConnectorTraffic.ps1 -Hours 12

### Analyse the last 3 days
.\Get-ExchangeConnectorTraffic.ps1 -Days 3

### Filter to a specific connector
.\Get-ExchangeConnectorTraffic.ps1 -Hours 5 -Connector "Anon Relay"

### Exclude known infrastructure IPs (e.g. Edge Transport servers)
.\Get-ExchangeConnectorTraffic.ps1 -Hours 5 -ExcludeIP "10.1.1.10","10.1.1.11"

### Export results to CSV (summary file is created automatically)
.\Get-ExchangeConnectorTraffic.ps1 -Days 7 -ExportCsv "C:\Reports\traffic.csv"

### Use a custom log path
.\Get-ExchangeConnectorTraffic.ps1 -LogPath "D:\ExchangeLogs\SmtpReceive" -Hours 24

### Combine options
.\Get-ExchangeConnectorTraffic.ps1 -Days 3 -Connector "Anon Relay" -ExcludeIP "10.1.1.10" -ExportCsv "C:\Reports\traffic.csv"
