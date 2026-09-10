
# 6az Position Grabber

Tampermonkey userscript that auto-extracts rack positions from the weekly SharePoint workload spreadsheet for **6az (IAD6)**.

## What It Does

- Opens the Excel file directly from SharePoint (no manual download needed)
- Filters rows by **AZ = 6AZ** and valid rack types (EC2, Bonsai, EBS)
- Groups positions by site with 2-row gaps
- Flags **lost reservations** and **rejected positions** with visual alerts
- One-click copy as 19-column TSV — paste directly into Col A of the workload
- Separate **Optics copy** button for the optics workload (6-col format)

## Install

1. Install [Tampermonkey](https://www.tampermonkey.net/) browser extension
2. Click the link below — Tampermonkey will prompt you to install:

   👉 **[Install 6az Position Grabber]👉 **[Install 6az Position Grabber](https://github.com/morakbri/aws-webscript/raw/refs/heads/main/6az-position-grabber.user.js)**


3. That's it — updates are automatic

## How to Use

1. Open the weekly workload `.xlsm` file in SharePoint
2. Click the blue **6az** button (bottom-right corner)
3. Select a sheet if needed, then click **Grab 6az Positions**
4. Review the results, then click **Copy (Clean)** or **Copy Optics (Clean)**
5. Paste into the destination workload at Col A

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 2.8.4 | 2026-09-10 | Added metronome TRUE/FALSE detection
| 2.8.3 | 2026-09-09 | Fixed GUID extraction bug; added WOPI fetch strategy |
| 2.8.2 | 2026-09-08 | Added HTML response guard; multi-strategy auth-resilient fetch |
| 2.8.0 | 2026-09-07 | Initial multi-strategy fetch with auth headers |

