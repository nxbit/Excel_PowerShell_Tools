<# 
	Script Creates a Combined Workbook, 
	Opens each workbook in given folder and copys contents of the first sheet, 
	and pastes to the single Combined Workbook.
 #>

Try {

<# Using Excel.Appliction Object, Initilizing Objects and Path's used #>
$ExcelObject=New-Object -ComObject excel.application
$ExcelObject.visible=$false
$ExcelFilesPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)+'\DataSet\'
$ExcelFileMerged = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)+'\MergedSet.xlsx'
<# Grabing File Count #>
$ExcelFiles = Get-ChildItem $ExcelFilesPath
$ExcelFiles.Count
<# Initilize the Wookbook with a Sheet1 #>
$Workbook=$ExcelObject.Workbooks.add()
#$Worksheet=$Workbook.Sheets.Item("Summary")

<# Open Each ExcelFile in the Folder and Copy its Data to the Combined Workbook #>
foreach($ExcelFile in $ExcelFiles){
 
$Everyexcel=$ExcelObject.Workbooks.Open($ExcelFile.FullName)
$Everysheet=$Everyexcel.sheets.item("Data")
$Everyexcel.sheets.item("Data").Copy($Worksheet)
$Everyexcel.Close()
 
}
<# Save the Merged Workbook and Close #>
$Workbook.SaveAs($ExcelFileMerged)
$ExcelObject.Quit()

Write-Output "Process Completed, Close this window and Proceed with next steps."
}catch{
Write-Output "Process could not be completed. Check you have staged all files in the nessary folder."
}