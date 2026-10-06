# Exchange SOA Conversion Tool

A PowerShell GUI tool for managing **Source of Authority (SOA)** conversion for Exchange mailboxes in a Microsoft 365 hybrid environment.

---

## Features

- Exchange mailbox SOA via `IsExchangeCloudManaged` (Exchange Online PowerShell)
- **Live search** — filter by display name or email address across all loaded data (not just the current page)
- **Hide Converted** — checkbox to hide already-converted mailboxes
- **Pagination** — 100 rows per page with Previous / Next navigation
- **Column sorting** — click any column header to sort ascending/descending
- **Batch operations** — select multiple rows (Ctrl+Click / Shift+Click) and convert in one action
- **Progress indicator** with live mailbox count while loading and per-item progress during batch conversion
- **Attribute backup** — before each conversion or rollback, the mailbox's Alias, Primary SMTP, all email addresses, HiddenFromAddressListsEnabled and custom/extension attributes are written to the log (the conversion is skipped if they can't be read)
- **Logging** — single session log file written to the script directory; open via "Open Log"

---

## Requirements

### PowerShell module
| Module | Used for |
|---|---|
| `ExchangeOnlineManagement` | Connect to EXO, `Get-Mailbox`, `Set-Mailbox` |

The tool checks for the module at startup, tries every installed version, and offers to (re)install it (CurrentUser scope) if none can be loaded.

### Permissions

Sign in interactively with an account holding one of:

- `Exchange Administrator` *(recommended)*
- `Hybrid Identity Administrator`
- `Global Administrator`

> Interactive authentication is required. Certificate-based app-only authentication is not supported.

### Sync prerequisites

- **Entra Connect Sync >= 2.5.190.0** (or **Cloud Sync**) must be deployed
- After a mailbox move or `Set-RemoteMailbox`, wait **one sync cycle + 24 hours** before converting that mailbox's SOA

---

## Usage

```powershell
.\Exchange-SOA-Conversion-Tool.ps1
```

Click **Connect to EXO** → authenticate → mailboxes load automatically.

---

## Troubleshooting

### "The specified module 'ExchangeOnlineManagement' was not loaded because no valid module file was found"

The module is visible on disk but no version can actually be imported — typically a broken/partial install, files stored as OneDrive online-only placeholders in a redirected `Documents\WindowsPowerShell\Modules` folder, or an incompatible newest version masking a working older one.

At startup the tool logs the PowerShell version, `PSModulePath`, and **every** installed `ExchangeOnlineManagement` version with its path, plus each import attempt and its error. Check the `SOAConversion_*.log` in the script directory to see exactly which versions/paths were tried.

Fix manually in PowerShell:

```powershell
Uninstall-Module ExchangeOnlineManagement -AllVersions -Force
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
```

Then restart the tool.

### Module installed but still not found (server machines)

On servers where `PSModulePath` was overridden — e.g. by Microsoft Hybrid Service or a Monitoring Agent — the per-user `Documents\WindowsPowerShell\Modules` folder may not be in the module path, so a `CurrentUser`-scoped install is invisible to `Get-Module -ListAvailable`. The tool now appends the standard user and all-users module paths to `PSModulePath` for the session and, when run elevated, installs with `-Scope AllUsers`. On such machines, install manually from an elevated prompt:

```powershell
Install-Module ExchangeOnlineManagement -Scope AllUsers -Force
```

### No sign-in prompt appears after "Connecting to Exchange Online..."

Since ExchangeOnlineManagement 3.7.0, sign-in uses Web Account Manager (WAM) by default. On older servers (e.g. Windows Server 2016) or when the tool runs under a different user (RunAs), the WAM prompt can hang or stay hidden. The tool therefore connects with `-DisableWAM` (module 3.7.2 or later) and uses the browser-based sign-in instead. See [Resolve issues in Exchange Online PowerShell after WAM integration](https://learn.microsoft.com/en-us/troubleshoot/exchange/administration/wam-integration-issues).

### .NET Framework requirement

The EXO v3 module requires **.NET Framework 4.7.2 or later** on Windows PowerShell (4.8 recommended). The startup log records the detected .NET release value; upgrade .NET first if the log flags it as too old.

---

## References

- [Exchange SOA — Cloud-based management of Exchange attributes](https://learn.microsoft.com/en-us/exchange/hybrid-deployment/enable-exchange-attributes-cloud-management)

---

## Version history

| Version | Notes |
|---|---|
| 1.00 | Initial release — Exchange SOA Conversion Tool combining Users, Groups, and Contacts SOA management |
| 1.10 | Users (mailboxes) only — Groups/Contacts tabs and related module dependencies removed; robust ExchangeOnlineManagement module loading (per-version import with logging + reinstall fallback); tenant parameter removed |

---

*MIT License — free to use and distribute, please leave author information.*  
Blog: [hcandersen.net](http://www.hcandersen.net) · LinkedIn: [Hans Chr. Andersen](https://www.linkedin.com/in/hanschrandersen/)
