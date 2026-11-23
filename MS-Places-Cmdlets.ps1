#Step 1: Install PowerShell 7.4 or later and required PowerShell modules

winget install --id Microsoft.PowerShell --source winget 
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
Install-Module -Name MicrosoftTeams -Scope CurrentUser
Install-Module -Name MicrosoftPlaces -Force


#Step 2: Assign Places management permissions
# Assign the built-in Places Administrator role in the Microsoft 365 admin center,
# or assign TenantPlacesManagement directly. The account also needs MailRecipient
# permissions to manage mailbox objects; see the linked Microsoft guidance.
Connect-ExchangeOnline
New-ManagementRoleAssignment -Role 'TenantPlacesManagement' -User '<admin@contoso.com>'

#Step 3: Install or update Microsoft Places powershell module

Install-Module -Name MicrosoftPlaces -Force
Connect-MicrosoftPlaces
# These settings are tenant-wide examples. Review the current cmdlet documentation
# for supported parameters, scoping, licensing, and propagation delays.
Set-PlacesSettings -EnableBuildings 'Default:true'
Set-PlacesSettings -EnablePlacesWebApp 'Default:true'
Set-PlacesSettings -PlacesFinderEnabled 'Default:true'
Set-PlacesSettings -SpaceAnalyticsEnabled 'Default:true'


#Approach 1: Intialize Microsoft Places
Initialize-Places
#Option 1 - Export suggested mappying CSV
#Option 2 - Import mapping to automatically create 
#Option 3 - Export ps scripts with cmdlets to create places

#Approach 2: Create buildings, floors, sections, rooms, workspaces, and desks
#Create a building for Contoso in New York
New-Place -Type Building -Name "Contoso NY" -Street "11 Times Square" -City "New York" -State "New" -PostalCode "10036" -CountryorRegion "United States"
$contosonyc = (Get-PlaceV3 -Type Building | Where-Object -Property DisplayName -eq 'Contoso NY').PlaceId

#Set the resource link for NYC
Set-PlaceV3 -Identity $contosonyc -ResourceLinks @{name = "Tech Support"; Value = "www.contoso.sharepoint.com/TechSupport"; type = "URL" }

#Create floors for Contoso NYC
New-Place -Type Floor -Name "11" -SortOrder 0 -ParentId $contosonyc
$contosol11 = (Get-PlaceV3 -AncestorId $contosonyc | Where-Object -Property DisplayName -eq '11').PlaceId
New-Place -Type Floor -Name "12" -SortOrder 1 -ParentId $contosonyc
$contosol12 = (Get-PlaceV3 -AncestorId $contosonyc | Where-Object -Property DisplayName -eq '12').PlaceId

#Create a room on an existing floor. New-Place provisions the room and mailbox.
$room = New-Place -Type Room -Name "ConfRm NY/11.260" -ParentId $contosol11 -Capacity 4 -Mailbox "ConfRm-NY-11.260@contoso.com"
Set-CalendarProcessing -Identity "ConfRm-NY-11.260@contoso.com" -AutomateProcessing AutoAccept -AddOrganizerToSubject $false -AllowRecurringMeetings $true -DeleteAttachments $true -DeleteComments $false -DeleteSubject $false -ProcessExternalMeetingMessages $true -RemovePrivateProperty $false -AddAdditionalResponse $true -AdditionalResponse "This is a Microsoft Teams Meeting room!"
#Auto-release requires an eligible space license for the room.
Set-CalendarProcessing -Identity "ConfRm-NY-11.260@contoso.com" -EnableAutoRelease $true

#Create Floor sections for Contoso NYC
$sectionSW1 = (New-Place -type Section -Name "NY.11.SouthWest" -ParentId $contosol11).PlaceId

#Create a workspace (desk pool) and an individual desk under the section.
#Workspaces and desks must be parented to a section.
$workspace = New-Place -Type Workspace -Name "Workspace NY/11.160" -Capacity 17 -ParentId $sectionSW1 -Mailbox "wksp-ny-11.160@contoso.com"
Set-PlaceV3 -Identity "wksp-ny-11.160@contoso.com" -IsWheelChairAccessible $true -Tags "Monitor", "Docking Station"

#Reservable mode enables individual desk booking; a desk mailbox is provisioned automatically.
$desk1 = New-Place -Type Desk -Name "Office NY/11.190" -ParentId $sectionSW1 -Mode @{Name = 'Reservable'}
Set-PlaceV3 -Identity $desk1.PlaceId -IsWheelChairAccessible $true -Tags "Office", "Monitor", "Docking Station", "Height Adjustable Desk"

#Update an existing conference room for Contoso NYC.
Set-PlaceV3 -Identity "ConfRm-NY-12.238" -Capacity 4 -Label "ConfRm NY/12.238" -FloorLabel "12" -IsWheelChairAccessible $True -Tags "IntelliFrame Camera" -MTREnabled $true -ParentId $contosol12

#Export information and objects for Contoso HQ
$contosohq = (Get-PlaceV3 -Type Building | Where-Object -Property DisplayName -eq 'Contoso HQ').PlaceId
Get-Placev3 -AncestorId $contosohq | Export-Csv -Path "C:\Temp\ContosoHQExport.csv"


#Building your Indoor Mapping Data Format IMDF Map to create the mapfeatures.csv file
Import-MapCorrelations -MapFilePath "C:\Temp\Contoso HQ.zip"

#Once you have the mappedfeatures.csv open and save as mappedfeatures.xlsx and then open ContosoHQExport.csv file and copy the contents to a new tab in the mappedfeatures.xlsx called "contoso hq export"
#Starting in cell A5,B5,C5 we can use an xloopup command to fill in columns A , B  C

#Column A5, uses F5 as a lookup value, and searches all of Column B in ContosoHQExport tab, returns a value from Column A, and must return an exact value
#Xlookup =XLOOKUP(F6,'C:\Temp\[ContosoHQExport.csv]ContosoHQExport'!$B$1:$B$98,'C:\Temp\[ContosoHQExport.csv]ContosoHQExport'!$A$1:$A$98,1,0)


#Column B5, uses F5 as a lookup value, and searches all of Column B in ContosoHQExport tab, returns a value from Column B, and must return an exact value
#Xlookup =XLOOKUP(F6,'C:\Temp\[ContosoHQExport.csv]ContosoHQExport'!$B$1:$B$98,'C:\Temp\[ContosoHQExport.csv]ContosoHQExport'!$B$1:$B$98,1,0)

#Column C5, uses F5 as a lookup value, and searches all of Column B in ContosoHQExport tab, returns a value from Column C, and must return an exact value
#Xlookup =XLOOKUP(F6,'C:\Temp\[ContosoHQExport.csv]ContosoHQExport'!$B$1:$B$98,'C:\Temp\[ContosoHQExport.csv]ContosoHQExport'!$C$1:$C$98,1,0)

#Once the lookup is complete, save has mappedfeatures.csv

#Create the correlation file by using xlookup to populate the PlaceID, Name, Type, Feature Type from the ContosoNYExport.csv file

#Created your correlated IMdF file
Import-MapCorrelations -MapFilePath "C:\Temp\Contoso HQ.zip" -CorrelationsFilePath "C:\Temp\mapfeatures.csv"

#Import the IMDF file to create the map
New-Map -BuildingId "1b9a176e-8f65-44bd-bf20-8aceca8f395a" -FilePath "C:\Temp\imdf_correlated.zip"



