# ExchangeTools

PowerShell scripts for Exchange Server administration and troubleshooting.

---

## Test-ExchangeMailFlow.ps1

Searches Exchange message tracking logs to verify outbound mail flow across
all transport servers in the organization.

Useful when testing mail flow after Exchange server migrations or connector
changes — shows which Exchange server handled the message, which Send Connector
was used, and whether the destination server accepted it.

### Requirements

- Exchange Management Shell (on-premises Exchange 2013/2016/2019)
- View-Only Organization Management or Message Tracking role

### Usage

```powershell
# Search by sender and recipient (last 15 minutes)
.\Test-ExchangeMailFlow.ps1 -From "user@company.com" -To "external@partner.com"

# Wider time range
.\Test-ExchangeMailFlow.ps1 -From "user@company.com" -To "external@partner.com" -MinutesBack 60

# Filter by subject with wildcard
.\Test-ExchangeMailFlow.ps1 -From "user@company.com" -Subject "*Invoice*"

# No filter - shows all messages from the last 5 minutes
.\Test-ExchangeMailFlow.ps1

# Show all events (not just key ones)
.\Test-ExchangeMailFlow.ps1 -From "user@company.com" -To "external@partner.com" -ShowAllEvents
