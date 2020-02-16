#Import the exported info from Visual Paradigm information back to the application

$VPsourcefolder= "$VPsourcefile" -replace ".{4}$"
$teamwork_path="$workspace_path\teamwork_client\projects"
$VPsourcefile_full = "$teamwork_path\$VPsourcefolder\$VPsourcefile"
$VPexportpath = "$currpath\vp_files"
$VPexportfile_full = "$VPexportpath\$VPexportfile"

#Export the NCI Buisness Architecture diagram to a VPfile called
Set-Location $VPscripts_path

& .\ImportExcel.bat -project "$VPsourcefile_full" -file "$VPexportfile_full"
