# Exchange SOA Conversion Tool

A combined PowerShell GUI tool for managing **Source of Authority (SOA)** conversion across Exchange mailboxes, Groups, and Contacts in a Microsoft 365 hybrid environment.

---

## Features

| Tab | What it manages | Connection |
|---|---|---|
| **Users (Mailboxes)** | Exchange mailbox SOA via `IsExchangeCloudManaged` | Exchange Online (EXO PS) |
| **Groups** | Mail-Enabled Security Groups & Distribution Groups via Graph `onPremisesSyncBehavior` | Microsoft Graph |
| **Contacts** | Org contacts (AD-synced) via Graph `onPremisesSyncBehavior` | Microsoft Graph (shared with Groups) |

### Shared across all tabs
- **Live search** — filter by display name or email address across all loaded data (not just the current page)
- **Hide Converted** — per-tab checkbox to hide already-converted items
- **Pagination** — 100 rows per page with Previous / Next navigation
- **Column sorting** — click any column header to sort ascending/descending
- **Batch operations** — select multiple rows (Ctrl+Click / Shift+Click) and convert in one action
- **Logging** — single session log file written to the script directory; open from any tab via "Open Log"

---

## Requirements

### PowerShell modules
The tool will attempt to auto-install missing modules on first run.

| Module | Used for |
|---|---|
| `ExchangeOnlineManagement` | Users tab — connect to EXO, `Get-Mailbox`, `Set-Mailbox` |
| `Microsoft.Graph.Groups` | Groups & Contacts tabs — connect to Graph, group/contact APIs |

### Permissions

**Users tab (Exchange Online)**
- `Exchange Administrator` or `Global Administrator`

**Groups tab (Microsoft Graph)**
- `Group.ReadWrite.All`
- `Group-OnPremisesSyncBehavior.ReadWrite.All` *(may require admin consent)*

**Contacts tab (Microsoft Graph)**
- `OrgContact.Read.All`
- `Contacts-OnPremisesSyncBehavior.ReadWrite.All` *(requires admin consent — new permission)*

> **Note on Contacts permission:** `Contacts-OnPremisesSyncBehavior.ReadWrite.All` is a newer Graph permission that typically requires **tenant admin consent**. If the Groups tab connects successfully but contacts SOA calls fail, grant consent via:
> Entra admin center → Enterprise Applications → *Microsoft Graph Command Line Tools* → Permissions → Grant admin consent

---

## Usage

```powershell
# Default (connects to your home tenant)
.\SOA-Conversion-Tool.ps1

# Multi-tenant / partner scenarios
.\SOA-Conversion-Tool.ps1 -TenantId "00000000-0000-0000-0000-000000000000"
```

### Workflow

1. **Users tab** — Click **Connect to EXO** → authenticate → data loads automatically
2. **Groups tab** — Click **Connect to Graph** → authenticate + grant consent → data loads
3. **Contacts tab** — If already connected via Groups tab, click **Refresh** directly; otherwise click **Connect to Graph** (same session)

> Groups and Contacts share a single Graph session. Connecting on either tab activates both. Disconnecting on either tab disconnects both.

---

## Contacts tab — important notes

- Shows **all org contacts** (both synced and cloud-only)
- Cloud-only contacts (`On-Prem Synced = False`) display **N/A** in the Cloud Managed column and are greyed out — conversion is not applicable to them
- SOA conversion only works for contacts where `On-Prem Synced = True`
- The tool guards against attempting conversion on non-synced contacts and will warn/skip them

---

## Groups tab — nested group ordering

The Groups tab automatically detects nested group relationships and:
- Shows a **Nesting Depth** column (0 = no nested children, higher = more nesting levels)
- Warns if you try to convert a parent group before its children
- Automatically sorts the conversion batch **bottom-up** (children before parents) — the Microsoft-recommended order
- Rollback is sorted **top-down** (parents before children)

---

## References

- [Exchange SOA — Cloud-based management of Exchange attributes](https://learn.microsoft.com/en-us/exchange/hybrid-deployment/enable-exchange-attributes-cloud-management)
- [Group SOA configuration](https://learn.microsoft.com/en-us/entra/identity/hybrid/how-to-group-source-of-authority-configure)
- [Contact SOA configuration](https://learn.microsoft.com/en-us/entra/identity/hybrid/how-to-user-source-of-authority-configure#configure-contact-soa)

---

## Version history

| Version | Notes |
|---|---|
| 1.00 | Initial release — Exchange SOA Conversion Tool combining Users, Groups, and Contacts SOA management |

---

*MIT License — free to use and distribute, please leave author information.*  
Blog: [hcandersen.net](http://www.hcandersen.net) · LinkedIn: [Hans Chr. Andersen](https://www.linkedin.com/in/hanschrandersen/)
