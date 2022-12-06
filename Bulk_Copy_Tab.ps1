<# 
	Script takes a group of excel files, copies out a single tab, and saves the single tab in another location
    useful when needing to strip out Data Tab's from several Reports
 #>

Try {

<# Using Excel.Appliction Object, Initilizing Objects and Path's used #>
$ExcelFilesPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)+'\DataSet\'
$ExcelFileMerged = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)+'\DataSetMerged\'
<# Grabing File Count #>
$ExcelFiles = Get-ChildItem $ExcelFilesPath
$ExcelFiles.Count


$ExcelObject=New-Object -ComObject excel.application
$ExcelObject.visible=$false


<# Open Each ExcelFile in the Folder and Copy its Data to the Combined Workbook #>
foreach($ExcelFile in $ExcelFiles){


<# Initilize the Wookbook with a Sheet1 #>
$Workbook=$ExcelObject.Workbooks.add()
$Worksheet=$Workbook.Sheets.Item("Sheet1")
 


$Everyexcel=$ExcelObject.Workbooks.Open($ExcelFile.FullName)

$Everysheet=$Everyexcel.sheets.item("Data")
$Everyexcel.sheets.item("Data").Copy($Worksheet)
$Everyexcel.Close()
$Workbook.Sheets.Item("Sheet1").Delete()
<# Save the Merged Workbook and Close #>
$Workbook.SaveAs($ExcelFileMerged+$ExcelFile.Name)
$Workbook.Close()



}

$ExcelObject.Quit() 

Write-Output "Process Completed, Close this window and Proceed with next steps."
}catch{
Write-Output "Process could not be completed. Check you have staged all files in the nessary folder."
}