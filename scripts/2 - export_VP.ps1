#Export the Visual Paradigm NCI Buisness Architecture to an Excel File


$VPsourcefolder= "$VPsourcefile" -replace ".{4}$"

$teamwork_path="$workspace_path\teamwork_client\projects"
$VPsourcefile_full = "$teamwork_path\$VPsourcefolder\$VPsourcefile"
$VPexportpath = "$currpath\vp_files"
$VPexportfile_full = "$VPexportpath\$VPexportfile"

#Export the NCI Buisness Architecture diagram to a VPfile called
Set-Location $VPscripts_path

& .\ExportExcel.bat -project "$VPsourcefile_full" -out "$VPexportfile_full"

#Make a backup of the file
$backuppath = "$VPexportpath\backup"
$backupfile = (Get-Date -Format "yyyyMMdd-HH-mm").ToString() + "_$VPexportfile" 
$backupfile_full = "$backuppath\$backupfile"
Copy-Item -Path "$VPexportfile_full" -Destination "$backupfile_full"