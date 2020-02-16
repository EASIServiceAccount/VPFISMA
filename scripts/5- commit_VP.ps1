# Step 7 - Commit VP workspace back to Teamwork Server

#Retrieve the Visual Paradigm NCI Buisness Architecture workspace

$VPsourcefolder= "$VPsourcefile" -replace ".{4}$"

$teamwork_path="$workspace_path\teamwork_client\projects"
$VPsourcefile_full = "$teamwork_path\$VPsourcefolder\$VPsourcefile"

Set-Location $VPscripts_path

.\CommitTeamworkProject.bat -workspace ""$workspace_path"" -project ""$VPsourcefile_full""
