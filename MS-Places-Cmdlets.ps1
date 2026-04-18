#Step 1: Install powershell 7 and exchange online powershell module

winget install --id Microsoft.PowerShell --source winget 
Install-Module -Name ExchangeOnlineManagement


#Step 2: Create an Exchange Online security group in MS graph and then Exchange Online PowerShell and create a new management role
Connect-ExchangeOnline
New-DistributionGroup -Type Security -Name "MS Places Management" -Alias msplacesmanagement -RequireSenderAuthenticationEnabled $true 


New-ManagementRoleAssignment –Role 'TenantPlacesManagement' –SecurityGroup 'MS Places Management'
New-ManagementRoleAssignment –Role 'Mail Recipient Creation' –SecurityGroup 'MS Places Management'
New-ManagementRoleAssignment –Role 'Mail Recipients' –SecurityGroup 'MS Places Management'

#Step 3: Install or update Microsoft Places powershell module

Install-Module MicrosoftPlaces –Force -AllowClobber
Connect-MicrosoftPlaces
Set-PlacesSettings -AllowInConnectionslist 'Default:true'
Set-PlacesSettings -EnableBuildings 'Default:true'
Set-PlacesSettings -EnablePlacesWebApp 'Default:true'
Set-PlacesSettings -PlacesFinderEnabled 'Default:true'
Set-PlacesSettings -SpaceAnalyticsEnabled 'Default:true'


#Approach 1: Intialize Microsoft Places
Initialize-Places
#Option 1 - Export suggested mappying CSV
#Option 2 - Import mapping to automatically create 
#Option 3 - Export ps scripts with cmdlets to create places

#Approach 2: Create a new buildings , floors, desks manually
#Create a workspace
New-Mailbox -Room -Alias "wksp-ny-2.260" -Name "Workspace NY/11.260" | Set-Mailbox -Type Workspace
Set-MailboxCalendarConfiguration -Identity "wksp-ny-2.260" -WorkingHoursTimeZone "Pacific Standard Time" -WorkingHoursStartTime 09:00:00
Set-CalendarProcessing -Identity "wksp-ny-2.260" -EnforceCapacity $True -AllowConflicts $true

#create a room
New-Mailbox -Room -Alias "ConfRm-NY-11.260" -Name "ConfRm NY/11.260"
Set-CalendarProcessing -Identity "ConfRm-NY-11.260" -AutomateProcessing AutoAccept -AddOrganizerToSubject $false -AllowRecurringMeetings $true -DeleteAttachments $true -DeleteComments $false -DeleteSubject $false -ProcessExternalMeetingMessages $true -RemovePrivateProperty $false -AddAdditionalResponse $true -AdditionalResponse "This is a Microsoft Teams Meeting room!"
#Set Auto Release Default is 10 mins
Set-CalendarProcessing -Identity "ConfRm-NY-11.260" -EnableAutoRelease $true 

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

#Create Floor sections for Contoso NYC
$sectionSW1 = (New-Place -type Section -Name "NY.11.SouthWest" -ParentId $contosol1).PlaceId

#Create individual Desk/Office for Contoso NYC
$desk1 = New-Place -type Desk -Name "Office NY/11.190" -ParentId $sectionSW1 

#Create a mailbox for the desk and assign Tags
$mbx1 = New-Mailbox -Room -Alias "office-ny-11.190" -Name "Office NY/11.190"
Set-Mailbox $mbx1.Identity -Type Desk -HiddenFromAddressListsEnabled $true
Set-PlaceV3 $desk1.PlaceId -Mailbox $mbx1.Identity -IsWheelChairAccessible $true -Tags "Office", "Monitor", "Docking Station", "Height Adjustable Desk"

#Create a workspace for Contoso NYC
Set-PlaceV3 -Identity "wksp-NY-11.160" -Capacity 17 -Label "Workspace NY/11.160" -FloorLabel "11" -IsWheelChairAccessible $True -Tags "Monitor", "Docking Station" -ParentId $sectionSW1

#Create a conference room for Contoso NYC
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



