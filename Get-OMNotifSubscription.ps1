	param (
		# Mandatory SCOM Management Server
    	[parameter(Mandatory=$true)]
        [string]$scom,

		# Mandatory credentials for access SCOM Remotely
        [parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]$credential, 

        # Optional use this if you want to save the document produced
        [parameter(Mandatory=$false)]
        [boolean]$saveandclose,

        # Optional, but required if using the SaveandClose option
        [parameter(Mandatory=$false)]
        [string]$path,

        # Optional, but required if using the SaveandClose option
        [parameter(Mandatory=$false)]
        [string]$filename

    )


#Import the Operations Manager Powershell Module
Import-Module OperationsManager


#Create initial spreadsheet 
#get SCOM Management Group Name
$MG = get-scommanagementgroup -Computername $scom -credential $credential
$mgname = $mg.name



#to use with sheet items
$count = 1

#Create a new Excel object using COM 
#$sheetcount = 1 
$doc = New-Object -ComObject Excel.Application
$doc.visible = $True 
$doc.DisplayAlerts = $false

# "xlMaximized" Value is -4137, from XlWindowState enumeration (Excel)
# URL - https://docs.microsoft.com/en-us/office/vba/api/excel.xlwindowstate
$doc.WindowState = -4137 
$Excel = $doc.Workbooks.Add() 
$MainSheet = $Excel.Worksheets.Item($count)  
$MainSheet.Name = 'SCOM Notification Subscriptions' 


#Create a Title for the first worksheet 
$row = 1 
$Column = 1 
$MainSheet.Cells.Item($row,$column)= "Notification Subscriptions for $MGname Management Group"

$range = $MainSheet.Range("a1","s2") 
$range.Merge() | Out-Null 
$range.VerticalAlignment = -4160 
 
#Give it a nice Style so it stands out 
$range.Style = 'Title' 

#Increment row for next set of data 
$row++;$row++ 
 
#Save the initial row so it can be used later to create a border 
#Counter variable for rows 
$intRow = $row 
$xlOpenXMLWorkbook=[int]51


#Create Headers for Data
$MainSheet.Cells.Item($intRow,1)  ="Subscription Name" 
$MainSheet.Cells.Item($intRow,2)  ="Channel" 
$MainSheet.Cells.Item($intRow,3)  ="Recipients" 
$MainSheet.Cells.Item($intRow,4)  ="Enabled" 
$MainSheet.Cells.Item($intRow,5)  ="Criteria" 
$MainSheet.Cells.Item($intRow,6)  ="Group Name" 
$MainSheet.Cells.Item($intRow,7)  ="Class Name" 

for ($col = 1; $col –le 7; $col++) 
     { 
          $MainSheet.Cells.Item($intRow,$col).Font.Bold = $True 
          $MainSheet.Cells.Item($intRow,$col).Font.ColorIndex = 41 
     } 
 
$intRow++ 

#get Notification Subscritions
$notifsubs = Get-SCOMNotificationSubscription -ComputerName $scom -Credential $credential  



foreach ($object in $notifsubs){ #Begin work on main sheet
  # Testing
  #  $object = $notifsubs[0]      

        $name = $object.displayname
        $channel = $object.Actions.DisplayName
        $receipients = ($object.ToRecipients.name) -join ","
        $enabled = ($object.enabled).tostring()
        $criteria = $object.Configuration.criteria
        $groups = $object.Configuration.MonitoringObjectGroupIds
        $classes = $object.Configuration.MonitoringClassIds
 

        #Work on the main sheet that contains all distributed applications
        #adds the Distributed Application Name, Management Pack and Object ID to the Main Sheet
        $Mainsheet.Cells.Item($intRow, 1) = $name 
        $Mainsheet.Cells.Item($intRow, 2) = $channel
        $Mainsheet.Cells.Item($intRow, 3) = $receipients
        $Mainsheet.Cells.Item($intRow, 4) = $enabled 
        $Mainsheet.Cells.Item($intRow, 5) = $criteria
        $groupnames = ""
        if ($groups)
        {
            $groupnamesarray = (Get-SCOMGroup -id $object.Configuration.MonitoringObjectGroupIds).DisplayName
            $groupnames = $groupnamesarray -join ","
        }
        
        $classnames = ""         
        if ($classes)
        {
            $classnamesarrary = (Get-SCOMClass -id $object.Configuration.MonitoringClassIds).DisplayName
            $classnames = $classnamesarrary -join ","
        }

        $Mainsheet.Cells.Item($intRow, 6) = $groupnames
        $Mainsheet.Cells.Item($intRow, 7) = $classnames


        $intRow = $intRow + 1 
        $MainSheet.UsedRange.EntireColumn.AutoFit()
        
        }#end work on main sheet



#Save file to specified location if SaveandClose is True
if($saveandclose -eq $true)
{

if (!(Test-Path -path "$path")) #create it if not existing 
  { 
  New-Item "$path" -type directory | out-null }


$file = "$path$filename.xlsx" 
if (test-path $file ) { remove-item $file } #delete the file if it already exists 

$Excel.SaveAs($file, $xlOpenXMLWorkbook) #save as an XML Workbook (xslx) 
$Excel.Saved = $True
$Excel.Close() 
$doc.quit()

}

#cleanup COM Objects
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($mainsheet)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc)

Remove-Variable -Name mainsheet
Remove-Variable -Name worksheet
Remove-Variable -Name excel
Remove-Variable -Name doc
