#Step 1: Update the array of teamwork files in the workspace directory

$teamwork_path="$workspace_path\teamwork_client\projects"

#$teamwork_file_array = Get-ChildItem -Path $teamwork_path\*.vpp -Recurse -Force | Where-Object -FilterScript {$_.Name -ne "$publishingproject_name"} | SELECT -expandproperty Name #-first 1
$teamwork_file_array = Get-ChildItem -Path $teamwork_path\*.vpp -Recurse -Force | Select-Object -ExpandProperty Name #-first 1

#Set the location to the VP scripts path for easy execution
Set-Location -Path $VPscripts_path #Change directory to the scripts path

foreach ($teamwork_file_name in $teamwork_file_array) {
  $teamwork_folder = $teamwork_file_name -replace ".{4}$" #Assign a folder to the same name as the .vpp but without the suffix
  $teamwork_path_full = "$teamwork_path\$teamwork_folder\$teamwork_file_name"
  .\UpdateTeamworkProject.bat -project ""$teamwork_path_full"" -workspace ""$workspace_path""
  #scan for errors and if one exists, then stop
}