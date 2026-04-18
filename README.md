# MS Places PowerShell Setup & Management

A comprehensive PowerShell script for setting up and managing **Microsoft Places** — Microsoft's workplace intelligence solution for hybrid work. This script walks through the full provisioning lifecycle: from installing prerequisites to creating buildings, floors, rooms, desks, and importing indoor maps.

## Prerequisites

- **PowerShell 7+** (installed via `winget`)
- **Exchange Online Management** PowerShell module
- **Microsoft Places** PowerShell module
- A Microsoft 365 tenant with Microsoft Places licensing
- Appropriate admin permissions (Exchange admin or delegated Places management role)

## What This Script Does

### Step 1 — Install Prerequisites

Installs PowerShell 7 and the Exchange Online Management module:

```powershell
winget install --id Microsoft.PowerShell --source winget
Install-Module -Name ExchangeOnlineManagement
```

### Step 2 — Create Security Group & Assign Management Roles

Connects to Exchange Online, creates a security group (`MS Places Management`), and assigns the required management roles:

- `TenantPlacesManagement`
- `Mail Recipient Creation`
- `Mail Recipients`

### Step 3 — Install & Configure Microsoft Places Module

Installs the `MicrosoftPlaces` module and enables key Places features:

| Setting | Description |
|---|---|
| `AllowInConnectionslist` | Enables Places connections |
| `EnableBuildings` | Enables building management |
| `EnablePlacesWebApp` | Enables the Places web app |
| `PlacesFinderEnabled` | Enables the Places finder experience |

### Approach 1 — Initialize Places (Guided Setup)

Uses `Initialize-Places` to bootstrap your environment with options to:

1. Export a suggested mapping CSV
2. Import mappings to auto-create Places objects
3. Export PowerShell scripts with cmdlets to create Places

### Approach 2 — Manual Building & Room Provisioning

Demonstrates creating the full Places hierarchy manually:

- **Workspaces** — shared bookable spaces (e.g., hot-desking areas)
- **Conference Rooms** — meeting rooms with Teams Rooms auto-accept and auto-release
- **Buildings** — physical building records with address and resource links
- **Floors** — floor levels within a building
- **Floor Sections** — zones/wings within a floor
- **Desks/Offices** — individual bookable desks with tags (monitor, docking station, height-adjustable, wheelchair accessible, etc.)

### Exporting Places Data

Exports all Places objects for a building to CSV for auditing or backup:

```powershell
Get-PlaceV3 -AncestorId $buildingId | Export-Csv -Path "C:\Temp\BuildingExport.csv"
```

### Indoor Mapping (IMDF)

Supports importing **Indoor Mapping Data Format (IMDF)** maps:

1. Import the IMDF zip to generate `mapfeatures.csv`
2. Correlate map features with Places objects using `xlookup`
3. Import the correlated IMDF file
4. Attach the map to a building with `New-Map`

## Usage

1. Open **PowerShell 7** as an administrator.
2. Run the script step-by-step (it's structured as a walkthrough, not an end-to-end automation).
3. Update placeholder values (aliases, building names, addresses) with your own tenant-specific information.

```powershell
# Run individual sections as needed
.\MS-Places-Cmdlets.ps1
```

## Key Cmdlets Reference

| Cmdlet | Purpose |
|---|---|
| `Connect-ExchangeOnline` | Authenticate to Exchange Online |
| `Connect-MicrosoftPlaces` | Authenticate to Microsoft Places |
| `Initialize-Places` | Guided Places setup wizard |
| `New-Place` | Create a building, floor, section, or desk |
| `Set-PlaceV3` | Update properties on a Places object |
| `Get-PlaceV3` | Query Places objects |
| `New-Mailbox -Room` | Create room/workspace/desk mailbox |
| `Set-CalendarProcessing` | Configure booking policies |
| `Import-MapCorrelations` | Import IMDF indoor map data |
| `New-Map` | Attach a map to a building |

## Export Unparented Items Script

`Export-MSPlaces-UnparentedItems.ps1` finds all **unparented** Rooms, Desks, and Spaces in your Microsoft Places environment and exports them into separate CSV files so you can assign them to the correct Building → Floor → Section hierarchy.

### What It Does

1. Connects to **Exchange Online** and **Microsoft Places**
2. Exports 3 reference CSVs (Buildings, Floors, Sections) for easy lookup
3. Finds all unparented Rooms, Desks, and Spaces from PlacesV3
4. Finds desk-named Room/Equipment mailboxes from Exchange that aren't parented
5. Exports separate CSVs per type with fill-in instructions in the column headers

### Output Files

| File | Description |
|---|---|
| `Ref_Buildings.csv` | Reference: all buildings |
| `Ref_Floors.csv` | Reference: all floors with building names |
| `Ref_Sections.csv` | Reference: all sections with floor and building names |
| `Unparented_Rooms.csv` | Rooms to parent (fill in Building & Floor from reference files) |
| `Unparented_Desks.csv` | Desks to parent (fill in Building, Floor & optionally Section) |
| `Unparented_Spaces.csv` | Spaces/Desk Pools to parent |

### Usage

```powershell
.\Export-MSPlaces-UnparentedItems.ps1
```

The script will prompt you for an output folder, connect to services, and export all unparented items with instructions on how to fill in the parent hierarchy.

## Import Unparented Items Script

`Import-MSPlaces-UnparentedItems.ps1` reads the edited CSV files from the Export script and **parents** each Room, Desk, or Space to the correct Building → Floor → Section in Microsoft Places.

### What It Does

1. Prompts for the folder containing your edited `Unparented_*.csv` files
2. Shows available files with a count of rows ready to import vs. not yet filled in
3. Loads the live Buildings, Floors, and Sections from Microsoft Places
4. Resolves the Building/Floor/Section **names** you typed in the CSV to PlaceId GUIDs
5. Parents each item using `Set-PlaceV3 -ParentId`
6. Optionally runs a verification check to show remaining unparented items

### Usage

```powershell
.\Import-MSPlaces-UnparentedItems.ps1
```

The script will prompt you to pick which file(s) to import and can process all files at once or one at a time.

> **Tip:** Building and Floor are required. Section is optional — if left blank, items are parented directly to the Floor.

## File Structure

```
MS-Places-Cmdlets.ps1               # Main PowerShell walkthrough script
Export-MSPlaces-UnparentedItems.ps1  # Export unparented Places items to CSV
Import-MSPlaces-UnparentedItems.ps1  # Import edited CSVs to parent items in Places
README.md                            # This file
```

## Resources

- [Microsoft Places Documentation](https://learn.microsoft.com/en-us/microsoft-365/places/)
- [Microsoft Places PowerShell Module](https://www.powershellgallery.com/packages/MicrosoftPlaces)
- [Exchange Online PowerShell](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell)
- [IMDF (Indoor Mapping Data Format)](https://register.apple.com/resources/imdf/)

## License

This project is provided as-is for educational and demonstration purposes.