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

Output
Console — Detailed Table
Found 3 message(s) across 2 session(s) in the last 5 hours:

Time (UTC)               Server         Connector              Source IP      From                             To
---------                ------         ---------              ---------      ----                             --
2026-03-24T15:59:44.273Z EBMMBXPROD02   Anon Relay EBMMBXPROD02  10.1.1.20  sender@example.com               recipient@gmail.com

Console — Executive Summary
============================================================
EXECUTIVE SUMMARY
============================================================
Period    : last 5 hours
From (UTC): 2026-03-24 11:00:00
To (UTC)  : 2026-03-24 16:00:00
Log path  : C:\...\SmtpReceive

Total messages : 3
Total sessions : 2

Server: EBMMBXPROD02
--------------------------------------------------------
  Connector: Anon Relay EBMMBXPROD02
    10.1.1.20                                    2 message(s)
    10.1.1.21                                    1 message(s)
============================================================

File Export (when -ExportCsv is used)
File	Content
traffic.csv	One row per message with all fields
traffic_summary.txt	Executive summary in plain text
Notes
Only sessions containing both a MAIL FROM and at least one RCPT TO command are included in the report. Connection attempts, health checks, and sessions without a completed envelope are automatically excluded.
Null sender addresses (MAIL FROM:<>) used by bounce/NDR messages may appear with an empty From field. This is expected SMTP behaviour.
Log files are selected based on their last-write time, so the time window is approximate at the boundary of each hourly file.
Progress is displayed while processing to indicate how many files have been handled and how many sessions have been found.
