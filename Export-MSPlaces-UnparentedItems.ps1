#Requires -Modules MicrosoftPlaces, ExchangeOnlineManagement
<#
.SYNOPSIS
    Export all unparented Places items into separate files with clear fill-in instructions.

.DESCRIPTION
    This script:
    1. Asks you where to save the output files
    2. Connects to Microsoft Places and Exchange Online
    3. Exports 3 reference CSVs (Buildings, Floors, Sections)
    4. Finds all unparented Rooms, Desks, Spaces from PlacesV3
    5. Finds desk-named Room/Equipment mailboxes from Exchange that are unparented
    6. Exports separate CSVs per type, each with instructions in the column headers

.NOTES
    Run in PowerShell 7+

.OUTPUTS
    Ref_Buildings.csv       - Reference: all buildings
    Ref_Floors.csv          - Reference: all floors with building names
    Ref_Sections.csv        - Reference: all sections with floor and building names
    Unparented_Rooms.csv    - Rooms to parent (fill in ParentId from Ref_Floors.csv > FloorPlaceId)
    Unparented_Desks.csv    - Desks to parent (fill in ParentId from Ref_Sections.csv > SectionPlaceId)
    Unparented_Spaces.csv   - Spaces/Desk Pools to parent (fill in ParentId from Ref_Sections.csv > SectionPlaceId)
#>

# ============================================================
# ASK FOR OUTPUT FOLDER
# ============================================================
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host " Microsoft Places - Export Unparented Items" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

$OutputFolder = Read-Host "Enter the output folder path (e.g. C:\temp\PlacesExport)"
$OutputFolder = $OutputFolder.TrimEnd('\')

if (-not $OutputFolder) {
    Write-Host "No path entered. Exiting." -ForegroundColor Red
    exit
}

if (-not (Test-Path $OutputFolder)) {
    Write-Host "Folder does not exist. Creating: $OutputFolder" -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

Write-Host "Output folder: $OutputFolder`n" -ForegroundColor Green

# ============================================================
# STEP 1: CHECK AND CONNECT TO SERVICES
# ============================================================
Write-Host "============================================" -ForegroundColor Cyan
Write-Host " STEP 1: Checking service connections" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

# --- Check Exchange Online ---
$exchangeConnected = $false
try {
    $exoConnection = Get-ConnectionInformation -ErrorAction Stop
    if ($exoConnection) {
        Write-Host "Exchange Online: Already connected as $($exoConnection.UserPrincipalName)" -ForegroundColor Green
        $exchangeConnected = $true
    }
}
catch {
    # Get-ConnectionInformation not available or no connection
}

if (-not $exchangeConnected) {
    Write-Host "Exchange Online: Not connected" -ForegroundColor Yellow
    $connectExo = Read-Host "Connect to Exchange Online now? (Y/N)"
    if ($connectExo -eq "Y" -or $connectExo -eq "y") {
        try {
            Connect-ExchangeOnline -ErrorAction Stop
            Write-Host "Exchange Online: Connected successfully" -ForegroundColor Green
            $exchangeConnected = $true
        }
        catch {
            Write-Host "Exchange Online: Connection failed - $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "The script needs Exchange Online to find desk-named mailboxes. Exiting." -ForegroundColor Red
            exit
        }
    }
    else {
        Write-Host "Exchange Online is required for this script. Exiting." -ForegroundColor Red
        exit
    }
}

# --- Check Microsoft Places ---
$placesConnected = $false
try {
    # Test with a quick PlacesV3 call
    $testPlaces = Get-PlaceV3 -Type Building -ErrorAction Stop | Select-Object -First 1
    if ($testPlaces -or $true) {
        Write-Host "Microsoft Places: Already connected" -ForegroundColor Green
        $placesConnected = $true
    }
}
catch {
    # Not connected or command failed
}

if (-not $placesConnected) {
    Write-Host "Microsoft Places: Not connected" -ForegroundColor Yellow
    $connectPlaces = Read-Host "Connect to Microsoft Places now? (Y/N)"
    if ($connectPlaces -eq "Y" -or $connectPlaces -eq "y") {
        try {
            Connect-MicrosoftPlaces -ErrorAction Stop
            Write-Host "Microsoft Places: Connected successfully" -ForegroundColor Green
            $placesConnected = $true
        }
        catch {
            Write-Host "Microsoft Places: Connection failed - $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "The script needs Microsoft Places to query rooms, desks, and spaces. Exiting." -ForegroundColor Red
            exit
        }
    }
    else {
        Write-Host "Microsoft Places is required for this script. Exiting." -ForegroundColor Red
        exit
    }
}

# ============================================================
# STEP 2: EXPORT REFERENCE HIERARCHY
# ============================================================
Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host " STEP 2: Exporting reference hierarchy" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

# --- Buildings ---
$buildings = Get-PlaceV3 -Type Building
$buildingsRef = $buildings | Select-Object `
@{N = 'BuildingPlaceId'; E = { $_.PlaceId } },
@{N = 'BuildingName'; E = { $_.DisplayName } },
@{N = 'City'; E = { $_.City } },
@{N = 'Street'; E = { $_.Street } },
@{N = 'CountryOrRegion'; E = { $_.CountryOrRegion } }

$buildingsRef | Export-Csv -Path "$OutputFolder\Ref_Buildings.csv" -NoTypeInformation
Write-Host "Buildings: $($buildingsRef.Count)" -ForegroundColor Green

# --- Floors (with Building name for easy lookup) ---
$floors = Get-PlaceV3 -Type Floor
$floorsRef = foreach ($f in $floors) {
    $bldgName = ($buildings | Where-Object { $_.PlaceId -eq $f.ParentId }).DisplayName
    [PSCustomObject]@{
        FloorPlaceId     = $f.PlaceId
        FloorName        = $f.DisplayName
        FloorLabel       = $f.FloorLabel
        BuildingName     = $bldgName
        ParentBuildingId = $f.ParentId
    }
}
$floorsRef | Export-Csv -Path "$OutputFolder\Ref_Floors.csv" -NoTypeInformation
Write-Host "Floors: $($floorsRef.Count)" -ForegroundColor Green

# --- Sections (with Floor and Building names for easy lookup) ---
$sections = Get-PlaceV3 -Type Section
$sectionsRef = foreach ($s in $sections) {
    $floor = $floorsRef | Where-Object { $_.FloorPlaceId -eq $s.ParentId }
    [PSCustomObject]@{
        SectionPlaceId = $s.PlaceId
        SectionName    = $s.DisplayName
        FloorName      = $floor.FloorName
        BuildingName   = $floor.BuildingName
        ParentFloorId  = $s.ParentId
    }
}
$sectionsRef | Export-Csv -Path "$OutputFolder\Ref_Sections.csv" -NoTypeInformation
Write-Host "Sections: $($sectionsRef.Count)" -ForegroundColor Green

Write-Host "`nReference files saved:" -ForegroundColor Cyan
Write-Host "  $OutputFolder\Ref_Buildings.csv"
Write-Host "  $OutputFolder\Ref_Floors.csv"
Write-Host "  $OutputFolder\Ref_Sections.csv"

# ============================================================
# STEP 3: FIND ALL UNPARENTED ITEMS FROM PLACEV3
# ============================================================
Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host " STEP 3: Finding unparented items in Places" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

$unparentedRooms = @()
$unparentedDesks = @()
$unparentedSpaces = @()

try {
    $allRooms = Get-PlaceV3 -Type Room
    $unparentedRooms = @($allRooms | Where-Object { -not $_.ParentId })
    Write-Host "Rooms  - Total: $($allRooms.Count), Unparented: $($unparentedRooms.Count)" -ForegroundColor Green
}
catch {
    Write-Host "Error getting Rooms: $($_.Exception.Message)" -ForegroundColor Red
}

try {
    $allDesks = Get-PlaceV3 -Type Desk
    $unparentedDesks = @($allDesks | Where-Object { -not $_.ParentId })
    Write-Host "Desks  - Total: $($allDesks.Count), Unparented: $($unparentedDesks.Count)" -ForegroundColor Green
}
catch {
    Write-Host "Error getting Desks: $($_.Exception.Message)" -ForegroundColor Red
}

try {
    $allSpaces = Get-PlaceV3 -Type Space
    $unparentedSpaces = @($allSpaces | Where-Object { -not $_.ParentId })
    Write-Host "Spaces - Total: $($allSpaces.Count), Unparented: $($unparentedSpaces.Count)" -ForegroundColor Green
}
catch {
    Write-Host "Error getting Spaces: $($_.Exception.Message)" -ForegroundColor Red
}

# ============================================================
# STEP 4: FIND DESK-NAMED MAILBOXES FROM EXCHANGE
#         These often show as Type=Room in Places but are
#         actually desks. They won't appear in Get-PlaceV3 -Type Desk
# ============================================================
Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host " STEP 4: Finding desk-named mailboxes from Exchange" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

$knownPlaceIds = @()
$knownPlaceIds += $unparentedRooms | ForEach-Object { $_.PlaceId }
$knownPlaceIds += $unparentedDesks | ForEach-Object { $_.PlaceId }
$knownPlaceIds += $unparentedSpaces | ForEach-Object { $_.PlaceId }

# Room mailboxes with desk-like names
$deskRoomMailboxes = @(Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited |
    Where-Object { $_.DisplayName -like "desk-*" -or $_.Alias -like "desk-*" })
Write-Host "Desk-named Room mailboxes in Exchange: $($deskRoomMailboxes.Count)"

# Equipment mailboxes with desk-like names
$deskEquipmentMailboxes = @(Get-Mailbox -RecipientTypeDetails EquipmentMailbox -ResultSize Unlimited |
    Where-Object { $_.DisplayName -like "desk-*" -or $_.Alias -like "desk-*" })
Write-Host "Desk-named Equipment mailboxes in Exchange: $($deskEquipmentMailboxes.Count)"

$allDeskMailboxes = @()
$allDeskMailboxes += $deskRoomMailboxes
$allDeskMailboxes += $deskEquipmentMailboxes

$exchangeMissing = @()
$exchangeAlreadyCaptured = 0
$exchangeAlreadyParented = 0
$exchangeNotInPlaces = 0

foreach ($mbx in $allDeskMailboxes) {
    try {
        $place = Get-PlaceV3 -Identity $mbx.ExternalDirectoryObjectId

        if ($place.PlaceId -in $knownPlaceIds) {
            $exchangeAlreadyCaptured++
            continue
        }

        if (-not $place.ParentId) {
            $exchangeMissing += $place
            $knownPlaceIds += $place.PlaceId
            Write-Host "  Found unparented: $($place.DisplayName)" -ForegroundColor Yellow
        }
        else {
            $exchangeAlreadyParented++
        }
    }
    catch {
        $exchangeNotInPlaces++
        Write-Host "  Not in Places: $($mbx.DisplayName) ($($mbx.PrimarySmtpAddress))" -ForegroundColor Gray
    }
}

Write-Host "`nExchange lookup results:" -ForegroundColor Cyan
Write-Host "  Already captured in Step 3:  $exchangeAlreadyCaptured"
Write-Host "  Already parented:            $exchangeAlreadyParented"
Write-Host "  Not in Places at all:        $exchangeNotInPlaces"
Write-Host "  NEW unparented found:        $($exchangeMissing.Count)" -ForegroundColor Yellow

# ============================================================
# STEP 5: SORT INTO SEPARATE FILES AND EXPORT
# ============================================================
Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host " STEP 5: Building separate export CSVs" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

# --- Separate desk-named items from real rooms ---
$realRooms = @($unparentedRooms | Where-Object { $_.DisplayName -notlike "desk-*" })
$deskNamedRooms = @($unparentedRooms | Where-Object { $_.DisplayName -like "desk-*" })

# --- Combine all desk items and deduplicate ---
$allDeskItems = @()
$allDeskItems += $unparentedDesks
$allDeskItems += $deskNamedRooms
$allDeskItems += $exchangeMissing
$allDeskItems = @($allDeskItems | Sort-Object PlaceId -Unique)

# =======================================================================
# FILE: Unparented_Rooms.csv
# =======================================================================

$roomsExport = foreach ($r in $realRooms) {
    [PSCustomObject]@{
        PlaceId     = $r.PlaceId
        Identity    = $r.Identity
        DisplayName = $r.DisplayName
        Type        = "Room"
        Building    = "REQUIRED -> Open Ref_Buildings.csv -> Type the BuildingName"
        Floor       = "REQUIRED -> Open Ref_Floors.csv -> Type the FloorName"
        Section     = "OPTIONAL -> Open Ref_Sections.csv -> Type the SectionName"
    }
}

if ($roomsExport.Count -gt 0) {
    $roomsExport | Export-Csv -Path "$OutputFolder\Unparented_Rooms.csv" -NoTypeInformation
    Write-Host "Unparented_Rooms.csv   -> $($roomsExport.Count) rooms" -ForegroundColor Green
}
else {
    Write-Host "Unparented_Rooms.csv   -> 0 rooms (none found, file not created)" -ForegroundColor Gray
}

# =======================================================================
# FILE: Unparented_Desks.csv
# =======================================================================

$desksExport = foreach ($d in $allDeskItems) {
    [PSCustomObject]@{
        PlaceId     = $d.PlaceId
        Identity    = $d.Identity
        DisplayName = $d.DisplayName
        Type        = $d.Type
        Building    = "REQUIRED -> Open Ref_Buildings.csv -> Type the BuildingName"
        Floor       = "REQUIRED -> Open Ref_Floors.csv -> Type the FloorName"
        Section     = "OPTIONAL -> Open Ref_Sections.csv -> Type the SectionName"
    }
}

if ($desksExport.Count -gt 0) {
    $desksExport | Export-Csv -Path "$OutputFolder\Unparented_Desks.csv" -NoTypeInformation
    Write-Host "Unparented_Desks.csv   -> $($desksExport.Count) desks" -ForegroundColor Green
}
else {
    Write-Host "Unparented_Desks.csv   -> 0 desks (none found, file not created)" -ForegroundColor Gray
}

# =======================================================================
# FILE: Unparented_Spaces.csv
# =======================================================================

$spacesExport = foreach ($s in $unparentedSpaces) {
    [PSCustomObject]@{
        PlaceId     = $s.PlaceId
        Identity    = $s.Identity
        DisplayName = $s.DisplayName
        Type        = "Space"
        Building    = "REQUIRED -> Open Ref_Buildings.csv -> Type the BuildingName"
        Floor       = "REQUIRED -> Open Ref_Floors.csv -> Type the FloorName"
        Section     = "OPTIONAL -> Open Ref_Sections.csv -> Type the SectionName"
    }
}

if ($spacesExport.Count -gt 0) {
    $spacesExport | Export-Csv -Path "$OutputFolder\Unparented_Spaces.csv" -NoTypeInformation
    Write-Host "Unparented_Spaces.csv  -> $($spacesExport.Count) spaces" -ForegroundColor Green
}
else {
    Write-Host "Unparented_Spaces.csv  -> 0 spaces (none found, file not created)" -ForegroundColor Gray
}

# ============================================================
# SUMMARY & INSTRUCTIONS
# ============================================================
$totalItems = $roomsExport.Count + $desksExport.Count + $spacesExport.Count

Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host " EXPORT COMPLETE - $totalItems items found" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "All files saved to: $OutputFolder" -ForegroundColor White
Write-Host ""
Write-Host "--------------------------------------------" -ForegroundColor Yellow
Write-Host " HOW TO FILL IN THE CSV FILES" -ForegroundColor Yellow
Write-Host "--------------------------------------------" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Each CSV has 3 columns to fill in:" -ForegroundColor White
Write-Host ""
Write-Host "  Building (REQUIRED)" -ForegroundColor White
Write-Host "    Open Ref_Buildings.csv and type the building name" -ForegroundColor Gray
Write-Host "    Example: Contoso HQ 1" -ForegroundColor Gray
Write-Host ""
Write-Host "  Floor (REQUIRED)" -ForegroundColor White
Write-Host "    Open Ref_Floors.csv and type the floor name" -ForegroundColor Gray
Write-Host "    Example: 1" -ForegroundColor Gray
Write-Host ""
Write-Host "  Section (OPTIONAL)" -ForegroundColor White
Write-Host "    Open Ref_Sections.csv and type the section name" -ForegroundColor Gray
Write-Host "    Example: NorthEast" -ForegroundColor Gray
Write-Host "    If you don't need a section, delete the text and leave it blank" -ForegroundColor Gray
Write-Host ""
Write-Host "  The import script will look up the names you type and" -ForegroundColor White
Write-Host "  find the correct GUIDs automatically." -ForegroundColor White
Write-Host ""
Write-Host "  IMPORTANT: Names must match EXACTLY as they appear" -ForegroundColor Yellow
Write-Host "  in the reference files (case-sensitive)." -ForegroundColor Yellow
Write-Host ""
Write-Host "  Example filled-in row:" -ForegroundColor Gray
Write-Host "    DisplayName=Conf Room Hood | Building=Contoso HQ 1 | Floor=1 | Section=" -ForegroundColor Gray
Write-Host "    DisplayName=desk-chicago.71.100 | Building=Contoso Chicago | Floor=71 | Section=NorthEast" -ForegroundColor Gray
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host " After editing, run Import-UnparentedPlaces.ps1" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""