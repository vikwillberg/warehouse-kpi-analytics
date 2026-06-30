# Picking TV Dashboard + Token-Refresh Failsafe


## Project Overview

Operational TV-display dashboard for the Site 1 (SITE1) Picking floor. The deliverable is a Power BI report rendered inside a PowerPoint slideshow that loops on a wall-mounted TV, with a scheduled-task workaround that periodically restarts PowerPoint to keep the embedded Power BI tile alive (its auth token otherwise expires and the tile goes blank).

There is **no Python ETL in this project** — data is sourced directly by the `.pbix` from the upstream IP/DS pipelines (this folder consumes their output). This folder owns only the dashboard surface and the failsafe automation.

## Folder structure

```
IN Picking Dashboard/
  Picking Dashboard.pbix                          # active dashboard (authoring file)
  Picking Dashboard for Download.pbix             # distributable copy
  Picking Dashboard Slides.pptx                   # static slides
  Picking Dashboard TV Slideshow 1.1.pptx         # active TV slideshow (embedded PBI tile)
  Picking Dashboard TV Slideshow 1.1 - Copy.pptx  # backup
  Token Refresh Failsafe/
    PowerPoint Token Refresh.xml                  # exported Windows scheduled task
    RestartPowerPointSlideshow.ps1                # restart helper invoked by the task
```

`.pbix` and `.pptx` files are binary and not parseable from this side — details below are inferred from filenames and the failsafe script. Verify visuals/data sources in Power BI Desktop / PowerPoint when changing them.

## Power BI report (inferred — verify in Power BI Desktop)

- **`Picking Dashboard.pbix`** — primary authoring file. Likely connects to the same SharePoint-synced BI data the IP / DS picking pipelines publish (e.g. `…/Shared BI/bi_reports/IP_Picking/BI Data`); confirm by opening in Power BI Desktop. TODO: verify exact data source list.
- **`Picking Dashboard for Download.pbix`** — copy intended for handoff/download (likely flattened or stripped of credentials). TODO: verify how it differs from the authoring file.

## TV slideshow

`Picking Dashboard TV Slideshow 1.1.pptx` is the file that actually runs on the wall display. It embeds the Power BI tile via the **Power BI for PowerPoint** add-in. PowerPoint is launched in slideshow mode (`/s`) and left running.

Because the embedded tile uses an OAuth token that expires after a few hours, the slideshow goes blank without intervention — handled by the failsafe below.

The `- Copy.pptx` is a manual backup; only the un-suffixed `1.1.pptx` is the live file referenced by the failsafe script.

## Token Refresh Failsafe

A Windows Task Scheduler entry (`PowerPoint Token Refresh.xml`) runs every 50 minutes (`PT50M` repetition, daily, indefinite) and invokes the helper script.

- **Trigger**: calendar trigger, repeats every 50 minutes from `2026-02-03T09:10:31` until `2050-02-03T09:10:31`. Daily schedule, runs as the interactive logged-on user (`DOMAIN\user`). Multiple-instances policy: `IgnoreNew`.
- **Action**: `powershell.exe -ExecutionPolicy Bypass -File "./RestartPowerPointSlideshow.ps1"`
  - Note: the scheduled task references `./RestartPowerPointSlideshow.ps1`, but the source-of-truth copy lives here under `Token Refresh Failsafe\`. Keep them in sync — if you edit the one in this folder, also update `./` (or re-import the task XML).
- **Restart on failure**: 3 attempts, 1-minute interval. Stops if going on batteries.

### `RestartPowerPointSlideshow.ps1`

1. `Stop-Process -Name POWERPNT -Force` — closes any running PowerPoint instance.
2. `Start-Sleep -Seconds 5` — waits for clean shutdown.
3. `Start-Process POWERPNT.EXE -ArgumentList "<path> /s"` — relaunches in slideshow mode.

The path **inside the script** is hardcoded to `…\KPI\IN Picking Dashboard\Picking Dashboard TV Slideshow.pptx` (no `1.1` suffix). The current live file is `…1.1.pptx`. **This mismatch means the failsafe is currently pointing at a non-existent file** — verify on the TV machine whether (a) the script has been edited locally to the `1.1` filename, (b) the live file is actually named without the `1.1` suffix, or (c) a renamed copy lives at the unsuffixed path. Treat this as a known issue until reconciled.

## Schedules / triggers

| Trigger | What it runs | Cadence |
|---|---|---|
| Windows Task Scheduler — `PowerPoint Token Refresh` | `RestartPowerPointSlideshow.ps1` | Every 50 min, all day, every day |

No Power Automate flow involved.

## Outputs / destinations

- TV display in the SITE1 Picking area — visual-only, no file outputs.
- No emails, no SharePoint pushes from this project.

## Configuration

- The PPTX file path is **hardcoded** in `RestartPowerPointSlideshow.ps1`.
- The scheduled task references **`./RestartPowerPointSlideshow.ps1`**, not this folder. If you move/rename the script, update the task XML or re-register the task.

## Dependencies / environment

- Microsoft PowerPoint with the **Power BI for PowerPoint** add-in installed and signed in.
- Microsoft Power BI Desktop (for editing `.pbix`).
- PowerShell (`powershell.exe`) with `-ExecutionPolicy Bypass` allowed for the failsafe task.
- Logged-on Windows user `DOMAIN\user` (the scheduled task runs as this principal — `LogonType=InteractiveToken`). Re-creating it on another account requires updating the `<UserId>` SID in the task XML.

## Known issues / gotchas

- **Filename drift**: `RestartPowerPointSlideshow.ps1` references `Picking Dashboard TV Slideshow.pptx` while the live file is `Picking Dashboard TV Slideshow 1.1.pptx`. Reconcile before relying on the failsafe.
- **Two copies of the failsafe script**: this folder vs `./`. The scheduled task only runs the latter. Don't assume edits here propagate.
- The 50-minute restart cadence is shorter than the typical Power BI token TTL (1 hour) — intentional safety margin; don't increase past 55 min.
- `desktop.ini` is a Windows artifact — ignore.
- TODO: verify the actual Power BI data sources by opening `Picking Dashboard.pbix` in Power BI Desktop.