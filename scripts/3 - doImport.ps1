# This script runs the doImport function from the Excel VP import
# Make sure macros are enabled


$teamwork_path="$workspace_path\teamwork_client\projects"

# Step 4 - Run the DoImport process in Excel
$excel = New-Object -comobject Excel.Application
$VPImportFile_folder="vp_files"

#Make sure that the file identified below has the proper parameters for inserting the data
#Use $VPExportFile_full and $SPDownload_path below

$VPImportFile_full="$currpath\$VPImportFile_folder\$VPImportFile_name"
$VPExportFile_full="$currpath\$VPImportFile_folder\$VPExportFile_name"
$VPDownload_path ="$currpath\$VPImportFile_folder\$VPInventoryFile_name"

#open file

$workbook = $excel.Workbooks.Open($VPImportFile_full)

#make it visible (just to check what is happening)
#$excel.Visible = $true
$UI = $workbook.Worksheets.Item(1)
$UI.Cells.Item(2,2) = $VPExportFile_full
$UI.Cells.Item(3,2) = $VPDownload_path

#access the Application object and run a macro
$app = $excel.Application
$app.Run("doImport")

$excel.quit()

While([System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)) {}
