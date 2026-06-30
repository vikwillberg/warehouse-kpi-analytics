# PowerPoint Dashboard Refresh Script
# Purpose: Reload ONLY the Power BI dashboard deck so its embedded tile gets a
#          fresh WebView + OAuth token (it otherwise goes stale / white). Any
#          OTHER decks open in PowerPoint (e.g. the safety slides) are left
#          running untouched -- this script never kills the PowerPoint process.
# Deploy:  Use Setup-TokenRefresh.ps1 (no admin needed) on EACH TV machine.

# --- Config -----------------------------------------------------------------
# The exact dashboard file to (re)open and run as a show:
$fileName = "Picking Dashboard TV Slideshow 2.0.pptx"
# A name fragment that matches the dashboard deck (any version), so we close the
# right deck even if an older version is open. Must NOT match the safety deck.
$dashMatch = "Picking Dashboard TV Slideshow"
# Seconds to let the Power BI add-in authenticate + render before starting the
# show. Tune to just past how long the tile takes to appear on a manual open.
$addinLoadSeconds = 40

$presentationPath = Join-Path $env:USERPROFILE `
    "OneDrive - Example Logistics\Desktop\KPI\IN Picking Dashboard\$fileName"

# --- Logging (user-writable, no admin) --------------------------------------
$logDir  = Join-Path $env:USERPROFILE "PowerPointTokenRefresh"
$logPath = Join-Path $logDir "TokenRefresh.log"
function Write-Log($msg) {
    $line = "{0}  {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg
    Write-Host $line
    try {
        if (-not (Test-Path -LiteralPath $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }
        Add-Content -LiteralPath $logPath -Value $line -ErrorAction Stop
    } catch {}
}

Write-Log "=== Dashboard refresh cycle starting ==="

# Verify the dashboard file exists before doing anything.
if (-not (Test-Path -LiteralPath $presentationPath)) {
    Write-Log "ERROR: dashboard not found: $presentationPath -- aborting, nothing touched."
    exit 1
}

# --- Attach to PowerPoint (or start it if it isn't running) -----------------
$ppt = $null
try { $ppt = [System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') } catch { $ppt = $null }

if ($ppt) {
    Write-Log "PowerPoint is running -- cycling ONLY the dashboard deck, leaving other decks alone."

    # 1. Exit the dashboard's slide show window (if it's currently presenting).
    #    Iterate by index, backwards, since exiting changes the collection.
    try {
        for ($i = $ppt.SlideShowWindows.Count; $i -ge 1; $i--) {
            $w = $ppt.SlideShowWindows.Item($i)
            if ($w.Presentation.Name -like "*$dashMatch*") {
                $w.View.Exit()
                Write-Log "Exited the dashboard's slide show."
            }
        }
    } catch { Write-Log "Note: could not exit dashboard show window: $($_.Exception.Message)" }

    Start-Sleep -Seconds 2

    # 2. Close ONLY the dashboard presentation(s). Saved=$true suppresses any
    #    'save changes?' prompt. Other open decks (safety slides) are skipped.
    try {
        for ($i = $ppt.Presentations.Count; $i -ge 1; $i--) {
            $p = $ppt.Presentations.Item($i)
            if ($p.Name -like "*$dashMatch*") {
                $p.Saved = $true
                $p.Close()
                Write-Log "Closed dashboard deck: $($p.Name)"
            }
        }
    } catch { Write-Log "Note: could not close dashboard deck: $($_.Exception.Message)" }

    Start-Sleep -Seconds 3
}
else {
    Write-Log "PowerPoint not running -- starting it with the dashboard."
    Start-Process -FilePath "POWERPNT.EXE" -ArgumentList ('"{0}"' -f $presentationPath)
    Start-Sleep -Seconds 10
    try { $ppt = [System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') } catch { $ppt = $null }
    if (-not $ppt) {
        Write-Log "ERROR: could not attach to PowerPoint after launch -- aborting."
        exit 1
    }
}

# --- Make sure the dashboard is open, let the add-in load, then run the show -
try {
    $msoTrue = -1; $msoFalse = 0

    # Find the dashboard among open decks; open it if it's not there.
    $dash = $null
    for ($i = 1; $i -le $ppt.Presentations.Count; $i++) {
        $p = $ppt.Presentations.Item($i)
        if ($p.Name -like "*$dashMatch*") { $dash = $p; break }
    }
    if (-not $dash) {
        # ReadOnly, not Untitled, WithWindow -- a window is required for the
        # Power BI add-in to initialize and for the show to run.
        $dash = $ppt.Presentations.Open($presentationPath, $msoTrue, $msoFalse, $msoTrue)
        Write-Log "Opened dashboard deck in normal view."
    }

    # Park the dashboard's editor window OFF-SCREEN so it doesn't cover the other
    # decks (e.g. the safety slides) while the add-in loads. Keep it in NORMAL
    # state -- the slide must stay rendered for the Power BI content add-in to
    # initialize and authenticate; minimizing would stop it loading and we'd be
    # back to a white tile. The show itself (started below) appears on its own
    # configured monitor regardless of where this editor window sits.
    try {
        $win = $dash.Windows.Item(1)
        $win.WindowState = 1          # ppWindowNormal (required before moving)
        $win.Left = 30000             # points -- far off the visible desktop
        $win.Top  = 30000
        Write-Log "Parked dashboard editor off-screen during load."
    } catch { Write-Log "Note: could not move dashboard window off-screen: $($_.Exception.Message)" }

    # Give the Power BI add-in time to authenticate + render the tile.
    Write-Log "Waiting $addinLoadSeconds s for the Power BI tile to load..."
    Start-Sleep -Seconds $addinLoadSeconds

    # Start the dashboard's slide show (uses its saved kiosk / loop settings,
    # and its configured 'show on monitor' if you're on multiple displays).
    $dash.SlideShowSettings.Run() | Out-Null
    Write-Log "Dashboard slide show started (tile had time to load). Other decks left as-is."
}
catch {
    Write-Log "ERROR reopening/running dashboard: $($_.Exception.Message)"
}
