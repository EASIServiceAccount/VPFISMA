#Log file location is here
$currpath = "$PSScriptRoot\.."
$scriptpath="$currpath\scripts"
$logFile = "$currpath/log_summary.txt"
$logDetailsFile = "$currpath/log_details.txt"

#**Set functions here**
function checkError {
  param($cmdMsg)
  $matchErrorString = 'False'
  [string[]] $errMsgArray = @("Workspace in use","*(401) Unauthorized*")
  foreach ($errMsg in $errMsgArray) {
      write-output "This is my message: $errMsg"
        if ($cmdMsg -like ($errMsg)) {
        Log-Output ("Error found: $cmdMsg")
        Log-Output ("***End process***")
        $matchErrorString = 'True'
        & "$scriptpath/8 - email_results.ps1" -status 'Failure' -body "VPFISMA script did not complete due to errors. See attachment." | Tee-Object -FilePath $logDetailsFile -Append
        write-Error "$cmdMsg" -ErrorAction Stop
        }
  }
  return $matchErrorString 
}

function getVariables{ param($strPropertiesLocation)   #Set the properties file location
    $global:properties_location = $strPropertiesLocation
    #Retrieve and prep the properties data
    $properties_content = Get-Content "$properties_location" -Raw
    $properties_content = $properties_content -replace "(\\)","/"
    $properties_content = [regex]::Escape($properties_content)
    $properties_content = $properties_content -replace "(\\r)?\\n",[Environment]::NewLine

    $props = ConvertFrom-StringData ($properties_content)

    $global:workspace_path=$props.workspace_path
    $global:VPscripts_path=$props.VPscripts_path
    $global:VPsourcefile=$props.VPsourcefile
    $global:VPexportfile=$props.VPexportfile
    $global:lkpMapFile=$props.lkpMapFile
    $global:iqyFile=$props.iqyFile
    $global:SiteURL=$props.SiteURL
    $global:FolderRelativeUrl=$props.FolderRelativeUrl 
    $global:sp_username=$props.sp_username
    $global:VPImportFile_name=$props.VPImportFile_name
    $global:VPExportFile_name=$props.VPExportFile_name
    $global:SMTPServer=$props.SMTPServer
    $global:SMTPPort=$props.SMTPPort
    $global:fromAddress=$props.fromAddress
    $global:toAddress=$props.toAddress
}

function getDateTime {
  $datetime = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
  return $datetime
}

function Log-Output {
  param($logmsg)
  Write-Output "$(getDateTime): $logmsg " | Tee-Object -FilePath $logFile -Append | Tee-Object -FilePath $logDetailsFile -Append
}

getVariables("$scriptpath\variables.properties")

#Start the Process
Write-Output "$(getDateTime): ***Start process***" | Tee-Object -FilePath $logFile | Tee-Object -FilePath $logDetailsFile

#Step 1: Update the array of teamwork files in the workspace directory
Log-Output ("Step 1 - Update of VP projects starting...")
& "$scriptpath/1 - update_VP_projects.ps1" | Tee-Object -FilePath $logDetailsFile -Append | Tee-Object -Variable cmdOutput
$isError = checkError $cmdOutput
Log-Output ("Step 1 - Update pf VP projects finishing...")

# Step 2 - Export the Visual Paradigm NCI Buisness Architecture to an Excel File
# Make a backup as well with the date timestamp
Log-Output ("Step 2 - Export of VP starting...")
& "$scriptpath/2 - export_VP.ps1" | Tee-Object -FilePath $logDetailsFile -Append | Tee-Object -Variable cmdOutput
$isError = checkError $cmdOutput
Log-Output ("Step 2 - Export of VP finishing...")

# Step 3 - Run the DoImport process in Excel
Log-Output ("Step 3 - run Excel doImport() process starting...")
& "$scriptpath/3 - doImport.ps1" | Tee-Object -FilePath $logDetailsFile -Append | Tee-Object -Variable cmdOutput
$isError = checkError $cmdOutput
Log-Output ("Step 3 - run Excel doImport() process finishing...")

# Step 4 - Import the exported info from Visual Paradigm information back to the application
Log-Output ("Step 4 - Import of map into VP starting...")
& "$scriptpath/4 - import_VP.ps1" | Tee-Object -FilePath $logDetailsFile -Append | Tee-Object -Variable cmdOutput
$isError = checkError $cmdOutput
Log-Output ("Step 4 - Import of map into VP finishing...")

# Step 5 - Commit VP workspace back to Teamwork Server
Log-Output ("Step 5 - Commit VP project into teamwork server starting...")
& "$scriptpath/5 - Commit_VP.ps1" | Tee-Object -FilePath $logDetailsFile -Append | Tee-Object -Variable cmdOutput
$isError = checkError $cmdOutput
Log-Output ("Step 5 - Commit VP project into teamwork server finishing...")

# Step 6 - Send email and finish
Log-Output ("Step 6 - Sending email to team starting...")
Log-Output ("Step 6 - Sending email to team finishing...")
Log-Output ("***End process***")
& "$scriptpath/6 - email_results.ps1" -status 'Success' | Tee-Object -FilePath $logDetailsFile -Append 