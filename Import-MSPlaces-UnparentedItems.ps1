#Requires -Modules MicrosoftPlaces
<#
.SYNOPSIS
    Import edited CSV to parent Rooms, Desks, or Spaces in Microsoft Places.

.DESCRIPTION
    This script:
    1. Asks you for the folder where your edited CSV files are
    2. Shows you which files are available and lets you pick one
    3. Loads the reference data (Buildings, Floors, Sections)
    4. Resolves the Building/Floor/Section names you typed to GUIDs
    5. Parents each item using Set-PlaceV3
    6. Verifies remaining unparented items after import

    Building and Floor are REQUIRED. Section is OPTIONAL.
    - If Section is provided: item is parented to the Section
    - If Section is blank: item is parented to the Floor

.NOTES
    Run in PowerShell 7+
    Edit the CSV files from the Export script before running this.
#>

# ============================================================
# ASK FOR FILE LOCATION
# ============================================================
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host " Microsoft Places - Import Parented Items" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

$InputFolder = Read-Host "Enter the folder path where your edited CSV files are (e.g. C:\temp\PlacesExport)"
$InputFolder = $InputFolder.TrimEnd('\')

if (-not $InputFolder -or -not (Test-Path $InputFolder)) {
    Write-Host "Folder not found: $InputFolder" -ForegroundColor Red
    exit
}

# ============================================================
# SHOW AVAILABLE FILES
# ============================================================
Write-Host ""
Write-Host "Looking for Unparented_*.csv files in: $InputFolder" -ForegroundColor Cyan
Write-Host ""

$availableFiles = @(Get-ChildItem -Path $InputFolder -Filter "Unparented_*.csv" | Sort-Object Name)

if ($availableFiles.Count -eq 0) {
    Write-Host "No Unparented_*.csv files found in that folder." -ForegroundColor Red
    Write-Host "Make sure you've run the Export script first." -ForegroundColor Yellow
    exit
}

Write-Host "Available files:" -ForegroundColor White
Write-Host ""
for ($i = 0; $i -lt $availableFiles.Count; $i++) {
    $fileRows = @(Import-Csv -Path $availableFiles[$i].FullName)
    $filledRows = @($fileRows | Where-Object {
            $_.Building -and $_.Building -notlike "REQUIRED*" -and
            $_.Floor -and $_.Floor -notlike "REQUIRED*"
        })
    $notFilledRows = @($fileRows | Where-Object {
            -not $_.Building -or $_.Building -like "REQUIRED*" -or
            -not $_.Floor -or $_.Floor -like "REQUIRED*"
        })

    Write-Host "  [$($i + 1)] $($availableFiles[$i].Name)" -ForegroundColor White
    Write-Host "      Total: $($fileRows.Count) | Ready to import: $($filledRows.Count) | Not filled in: $($notFilledRows.Count)" -ForegroundColor Gray
}

Write-Host ""
Write-Host "  [A] Import ALL files at once" -ForegroundColor White
Write-Host ""

$choice = Read-Host "Enter your choice (1-$($availableFiles.Count), or A for all)"

# ============================================================
# DETERMINE WHICH FILES TO IMPORT
# ============================================================
$filesToImport = @()

if ($choice -eq "A" -or $choice -eq "a") {
    $filesToImport = $availableFiles
    Write-Host "`nImporting all $($filesToImport.Count) files..." -ForegroundColor Cyan
}
elseif ($choice -match '^\d+$' -and [int]$choice -ge 1 -and [int]$choice -le $availableFiles.Count) {
    $filesToImport = @($availableFiles[[int]$choice - 1])
    Write-Host "`nImporting: $($filesToImport[0].Name)" -ForegroundColor Cyan
}
else {
    Write-Host "Invalid choice. Exiting." -ForegroundColor Red
    exit
}

# ============================================================
# CHECK AND CONNECT TO MICROSOFT PLACES
# ============================================================
Write-Host ""

$placesConnected = $false
try {
    $testPlaces = Get-PlaceV3 -Type Building -ErrorAction Stop | Select-Object -First 1
    if ($testPlaces -or $true) {
        Write-Host "Microsoft Places: Already connected" -ForegroundColor Green
        $placesConnected = $true
    }
}
catch {
    # Not connected
}

if (-not $placesConnected) {
    Write-Host "Microsoft Places: Not connected" -ForegroundColor Yellow
    $connectPlaces = Read-Host "Connect to Microsoft Places now? (Y/N)"
    if ($connectPlaces -eq "Y" -or $connectPlaces -eq "y") {
        try {
            Connect-MicrosoftPlaces -ErrorAction Stop
            Write-Host "Microsoft Places: Connected successfully" -ForegroundColor Green
        }
        catch {
            Write-Host "Microsoft Places: Connection failed - $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Cannot import without a connection to Microsoft Places. Exiting." -ForegroundColor Red
            exit
        }
    }
    else {
        Write-Host "Microsoft Places is required to import. Exiting." -ForegroundColor Red
        exit
    }
}

# ============================================================
# LOAD REFERENCE DATA (Buildings, Floors, Sections)
# ============================================================
Write-Host "`n--------------------------------------------" -ForegroundColor Cyan
Write-Host " Loading Buildings, Floors, Sections from Places..." -ForegroundColor Cyan
Write-Host "--------------------------------------------" -ForegroundColor Cyan

$refBuildings = @(Get-PlaceV3 -Type Building)
$refFloors = @(Get-PlaceV3 -Type Floor)
$refSections = @(Get-PlaceV3 -Type Section)

Write-Host "  Buildings: $($refBuildings.Count)"
Write-Host "  Floors:    $($refFloors.Count)"
Write-Host "  Sections:  $($refSections.Count)"

# ============================================================
# IMPORT EACH FILE
# ============================================================
$totalSuccess = 0
$totalFailed = 0
$totalSkipped = 0

foreach ($file in $filesToImport) {
    Write-Host "`n--------------------------------------------" -ForegroundColor Cyan
    Write-Host " Importing: $($file.Name)" -ForegroundColor Cyan
    Write-Host "--------------------------------------------" -ForegroundColor Cyan

    $rows = @(Import-Csv -Path $file.FullName)

    # Building and Floor must be filled in (not empty, not instruction text)
    $rowsToImport = @($rows | Where-Object {
            $_.Building -and $_.Building -notlike "REQUIRED*" -and
            $_.Floor -and $_.Floor -notlike "REQUIRED*"
        })
    $rowsSkipped = @($rows | Where-Object {
            -not $_.Building -or $_.Building -like "REQUIRED*" -or
            -not $_.Floor -or $_.Floor -like "REQUIRED*"
        })

    if ($rowsSkipped.Count -gt 0) {
        Write-Host "Skipping $($rowsSkipped.Count) rows without Building/Floor filled in" -ForegroundColor Gray
        $totalSkipped += $rowsSkipped.Count
    }

    if ($rowsToImport.Count -eq 0) {
        Write-Host "No rows with Building and Floor filled in. Nothing to import." -ForegroundColor Yellow
        continue
    }

    Write-Host "Processing $($rowsToImport.Count) rows...`n" -ForegroundColor White

    foreach ($r in $rowsToImport) {
        $displayName = $r.DisplayName
        $buildingName = $r.Building.Trim()
        $floorName = $r.Floor.Trim()
        $sectionName = if ($r.Section -and $r.Section -notlike "OPTIONAL*") { $r.Section.Trim() } else { "" }

        # --- Resolve Building ---
        $building = $refBuildings | Where-Object { $_.DisplayName -eq $buildingName }
        if (-not $building) {
            Write-Host "  FAIL: $displayName - Building '$buildingName' not found" -ForegroundColor Red
            $totalFailed++
            continue
        }

        # --- Resolve Floor (must be under this building) ---
        $floor = $refFloors | Where-Object {
            $_.DisplayName -eq $floorName -and $_.ParentId -eq $building.PlaceId
        }
        if (-not $floor) {
            Write-Host "  FAIL: $displayName - Floor '$floorName' not found in building '$buildingName'" -ForegroundColor Red
            $totalFailed++
            continue
        }

        # --- Resolve Section (optional, must be under this floor) ---
        $parentId = $floor.PlaceId
        $parentType = "Floor"

        if ($sectionName) {
            $section = $refSections | Where-Object {
                $_.DisplayName -eq $sectionName -and $_.ParentId -eq $floor.PlaceId
            }
            if (-not $section) {
                Write-Host "  FAIL: $displayName - Section '$sectionName' not found on floor '$floorName' in '$buildingName'" -ForegroundColor Red
                $totalFailed++
                continue
            }
            $parentId = $section.PlaceId
            $parentType = "Section"
        }

        # --- Determine identity (email for real rooms, PlaceId for desks/spaces) ---
        if ($r.Type -eq "Room" -and $r.Identity -and $r.DisplayName -notlike "desk-*") {
            $identityToUse = $r.Identity
        }
        else {
            $identityToUse = $r.PlaceId
        }

        # --- Set the parent ---
        try {
            Set-PlaceV3 -Identity $identityToUse -ParentId $parentId
            $path = "$buildingName / $floorName"
            if ($sectionName) { $path += " / $sectionName" }
            Write-Host "  OK:   $displayName -> $path ($parentType)" -ForegroundColor Green
            $totalSuccess++
        }
        catch {
            Write-Host "  FAIL: $displayName - $($_.Exception.Message)" -ForegroundColor Red
            $totalFailed++
        }
    }
}

# ============================================================
# RESULTS
# ============================================================
Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host " IMPORT COMPLETE" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Success: $totalSuccess" -ForegroundColor Green
Write-Host "  Failed:  $totalFailed" -ForegroundColor $(if ($totalFailed -gt 0) { "Red" } else { "Green" })
Write-Host "  Skipped: $totalSkipped (Building/Floor not filled in)" -ForegroundColor Gray

# ============================================================
# VERIFY
# ============================================================
Write-Host ""
$verify = Read-Host "Run verification check? (Y/N)"

if ($verify -eq "Y" -or $verify -eq "y") {
    Write-Host ""

    try {
        $remainingRooms = @(Get-PlaceV3 -Type Room | Where-Object { -not $_.ParentId })
        Write-Host "Unparented Rooms:  $($remainingRooms.Count)" -ForegroundColor $(if ($remainingRooms.Count -gt 0) { "Yellow" } else { "Green" })
        if ($remainingRooms.Count -gt 0) {
            $remainingRooms | ForEach-Object { Write-Host "    - $($_.DisplayName)" -ForegroundColor Gray }
        }
    }
    catch {
        Write-Host "Could not check Rooms: $($_.Exception.Message)" -ForegroundColor Red
    }

    try {
        $remainingDesks = @(Get-PlaceV3 -Type Desk | Where-Object { -not $_.ParentId })
        Write-Host "Unparented Desks:  $($remainingDesks.Count)" -ForegroundColor $(if ($remainingDesks.Count -gt 0) { "Yellow" } else { "Green" })
        if ($remainingDesks.Count -gt 0) {
            $remainingDesks | ForEach-Object { Write-Host "    - $($_.DisplayName)" -ForegroundColor Gray }
        }
    }
    catch {
        Write-Host "Could not check Desks: $($_.Exception.Message)" -ForegroundColor Red
    }

    try {
        $remainingSpaces = @(Get-PlaceV3 -Type Space | Where-Object { -not $_.ParentId })
        Write-Host "Unparented Spaces: $($remainingSpaces.Count)" -ForegroundColor $(if ($remainingSpaces.Count -gt 0) { "Yellow" } else { "Green" })
        if ($remainingSpaces.Count -gt 0) {
            $remainingSpaces | ForEach-Object { Write-Host "    - $($_.DisplayName)" -ForegroundColor Gray }
        }
    }
    catch {
        Write-Host "Could not check Spaces: $($_.Exception.Message)" -ForegroundColor Red
    }

    $totalRemaining = $remainingRooms.Count + $remainingDesks.Count + $remainingSpaces.Count

    Write-Host ""
    if ($totalRemaining -eq 0) {
        Write-Host "All items are parented!" -ForegroundColor Green
    }
    else {
        Write-Host "$totalRemaining items still unparented." -ForegroundColor Yellow
        Write-Host "Re-run the export script to generate updated files." -ForegroundColor Yellow
    }
}
else {
    Write-Host "Skipping verification." -ForegroundColor Gray
}

Write-Host ""