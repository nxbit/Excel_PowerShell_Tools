<# 
	Script takes a group of excel files, unhides the Data, and saves the saves teh file
    useful when needing to unhide a tab on multi files
 #>

Try {

<# Using Excel.Appliction Object, Initilizing Objects and Path's used #>
$ExcelFilesPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)+'\DataSet\'
$ExcelObject=New-Object -ComObject excel.application
$ExcelObject.visible=$false
<# Grabing File Count #>
$ExcelFiles = Get-ChildItem $ExcelFilesPath
$ExcelFiles.Count


<# Open Each ExcelFile in the Folder and Copy its Data to the Combined Workbook #>
foreach($ExcelFile in $ExcelFiles){


Write-Output "Processing File "+ $ExcelFile.FullName
$Everyexcel=$ExcelObject.Workbooks.Open($ExcelFile.FullName)
$Everyexcel.sheets.Item("Data").visible = $true
$Everyexcel.Save()
$Everyexcel.Close()

}

$ExcelObject.Quit() 

Write-Output "Process Completed, Close this window and Proceed with next steps."
}catch{
Write-Output "Process could not be completed. Check you have staged all files in the nessary folder."
}