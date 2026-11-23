# Microsoft Places PowerShell Scripts

This repository contains a step-by-step PowerShell walkthrough for configuring and provisioning **Microsoft Places**, plus separate scripts for exporting and assigning unparented Places items. The walkthrough is an example to adapt, not an end-to-end installer or tenant-safe automation.

## Prerequisites

- **PowerShell 7.4 or later**. Windows PowerShell is not supported for Microsoft Places management.
- **Exchange Online Management** PowerShell module
- **Microsoft Teams** PowerShell module (for Places administration workflows)
- **Microsoft Places** PowerShell module
- A Microsoft 365 tenant. Core Places features are available with supported Microsoft 365, Office 365, or Teams plans; some features, including individual desk booking, autorelease, and occupancy reports, require eligible space licenses.
- Places Administrator permissions, or the documented Exchange Online management roles: `TenantPlacesManagement` and `MailRecipient`. An Exchange Administrator can also manage mail recipients. Assign only the roles needed for the task.

See Microsoft's [Places license requirements](https://learn.microsoft.com/en-us/microsoft-365/places/what-is-places/license-requirements) and [Places admin setup](https://learn.microsoft.com/en-us/microsoft-365/places/what-is-places/set-up-your-account-places-admin) before making tenant changes.

## What This Script Does

### Step 1 — Install Prerequisites

Install PowerShell 7.4 or later, then install the modules used by the walkthrough:

```powershell
winget install --id Microsoft.PowerShell --source winget
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
Install-Module -Name MicrosoftTeams -Scope CurrentUser
Install-Module -Name MicrosoftPlaces -Force
```

### Step 2 — Assign Places management permissions

Assign the built-in Places Administrator role in the Microsoft 365 admin center, or assign the Places management role through Exchange Online. For example, to assign `TenantPlacesManagement` directly:

```powershell
Connect-ExchangeOnline
New-ManagementRoleAssignment -Role TenantPlacesManagement -User "<admin@contoso.com>"
```

The account also needs `MailRecipient` permissions to manage mailbox objects. Follow Microsoft's admin setup guidance for supported role groups and assignment options; avoid granting broad roles unless required.

### Step 3 — Install & Configure Microsoft Places Module

The walkthrough installs the `MicrosoftPlaces` module and demonstrates enabling supported Places features:

| Setting | Description |
|---|---|
| `EnableBuildings` | Enables building management |
| `EnablePlacesWebApp` | Enables the Places web app |
| `PlacesFinderEnabled` | Enables the Places finder experience |
| `SpaceAnalyticsEnabled` | Enables Places analytics |

Places settings can take up to 12 hours to propagate. Review the [Set-PlacesSettings reference](https://learn.microsoft.com/en-us/microsoft-365/places/powershell/set-placessettings) before changing tenant-wide defaults or scopes.

### Approach 1 — Initialize Places (Guided Setup)

Uses `Initialize-Places` to bootstrap your environment with options to:

1. Export a suggested mapping CSV
2. Import mappings to auto-create Places objects
3. Export PowerShell scripts with cmdlets to create Places

### Approach 2 — Manual Building & Room Provisioning

Demonstrates the Places hierarchy. Prefer the current `New-Place` provisioning flow where possible:

- **Workspaces** — desk pools; must be parented to a section
- **Conference rooms** — can be parented directly to a floor or to a section
- **Buildings** — physical building records with address and resource links
- **Floors** — floor levels within a building
- **Floor Sections** — zones/wings within a floor
- **Individual desks** — must be parented to a section and have a mode: `Reservable`, `DropIn`, `Assigned`, or `Unavailable`

Desk booking is licensed per space under current licensing. Autorelease for rooms and desks and occupancy reports also require eligible space licenses. See [New-Place](https://learn.microsoft.com/en-us/microsoft-365/places/powershell/new-place), [Set-PlaceV3](https://learn.microsoft.com/en-us/microsoft-365/places/powershell/set-placev3), and [license requirements](https://learn.microsoft.com/en-us/microsoft-365/places/what-is-places/license-requirements) for current details.

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

1. Open **PowerShell 7.4 or later**.
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
| `New-Place` | Create a building, floor, section, room, workspace, or desk |
| `Set-PlaceV3` | Update properties on a Places object |
| `Get-PlaceV3` | Query Places objects |
| `New-Mailbox -Room` | Create a mailbox using the legacy mailbox-based examples |
| `Set-CalendarProcessing` | Configure booking policies |
| `Import-MapCorrelations` | Import IMDF indoor map data |
| `New-Map` | Attach a map to a building |

## Export Unparented Items Script

`Export-MSPlaces-UnparentedItems.ps1` finds unparented Rooms, individual Desks, and Spaces (desk pools/workspaces) and exports separate CSV files so you can assign them to the correct Places hierarchy.

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
| `Unparented_Rooms.csv` | Rooms to parent (Building and Floor required; Section optional) |
| `Unparented_Desks.csv` | Individual desks to parent (Building, Floor, and Section required) |
| `Unparented_Spaces.csv` | Spaces/desk pools/workspaces to parent (Building, Floor, and Section required) |

### Usage

```powershell
.\Export-MSPlaces-UnparentedItems.ps1
```

The script will prompt you for an output folder, connect to services, and export all unparented items with instructions on how to fill in the parent hierarchy.

## Import Unparented Items Script

`Import-MSPlaces-UnparentedItems.ps1` reads the edited CSVs and parents each item to the selected Places hierarchy. It requires a Section for desks and workspaces, while allowing rooms to be parented directly to a Floor.

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

> Building and Floor are required for every item. Section is required for individual desks and workspaces (desk pools); it is optional only for rooms.

## File Structure

```
MS-Places-Cmdlets.ps1               # Main PowerShell walkthrough script
Export-MSPlaces-UnparentedItems.ps1  # Export unparented Places items to CSV
Import-MSPlaces-UnparentedItems.ps1  # Import edited CSVs to parent items in Places
README.md                            # This file
```

## Resources

- [Microsoft Places Documentation](https://learn.microsoft.com/en-us/microsoft-365/places/)
- [Microsoft Places overview](https://learn.microsoft.com/en-us/microsoft-365/places/what-is-places/what-is-places)
- [License requirements](https://learn.microsoft.com/en-us/microsoft-365/places/what-is-places/license-requirements)
- [Places admin setup and permissions](https://learn.microsoft.com/en-us/microsoft-365/places/what-is-places/set-up-your-account-places-admin)
- [New-Place](https://learn.microsoft.com/en-us/microsoft-365/places/powershell/new-place)
- [Get-PlaceV3](https://learn.microsoft.com/en-us/microsoft-365/places/powershell/get-placev3)
- [Set-PlaceV3](https://learn.microsoft.com/en-us/microsoft-365/places/powershell/set-placev3)
- [Set-PlacesSettings](https://learn.microsoft.com/en-us/microsoft-365/places/powershell/set-placessettings)
- [Desk booking and desk modes](https://learn.microsoft.com/en-us/microsoft-365/places/room-desk-booking/desks)
- [Microsoft Places PowerShell Module](https://www.powershellgallery.com/packages/MicrosoftPlaces)
- [Exchange Online PowerShell](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell)
- [IMDF (Indoor Mapping Data Format)](https://register.apple.com/resources/imdf/)

## License

This project is provided as-is for educational and demonstration purposes.