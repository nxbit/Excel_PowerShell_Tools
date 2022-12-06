$ExcelFilesPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)+'\DataSet\'
$Files = GCI $ExcelFilesPath | ?{$_.Extension -Match "xlsx?"} | select -ExpandProperty FullName


$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $True
$Excel.DisplayAlerts = $False

$fileCounter = 0
$OutputFileName = 'SingleTab'

$Dest = $Excel.Workbooks.Add()


ForEach($File in $Files){


    
    #For every 10 Files, it'll close and reopen excel. 
    if($fileCounter % 10 -eq 0)
    {
        $Dest.SaveAs([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)+"\$OutputFileName.xlsx",51)
        $Dest.close()
        $Dest = $null
        $Dest = $Excel.Workbooks.Open([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)+"\$OutputFileName.xlsx")

    }



    Write-Output "processing new File" $File
    $Source = $Excel.Workbooks.Open($File,$true,$true)

    $Source.ActiveSheet.UsedRange.Row

    $val1 = $(($Source.ActiveSheet.UsedRange.Rows|Select -Last 1).Row)
    $val2 = $val1
    $val3 = $(($Dest.ActiveSheet.UsedRange.Rows|Select -last 1).row+1)
    $val4 =  ($val3-1) + ($val1-1)

    Write-Output "row: $val3"
    
    
    If(($Dest.ActiveSheet.UsedRange.Count -eq 1) -and ([String]::IsNullOrEmpty($Dest.ActiveSheet.Range("A1").Value2))){ 
        #If the first paste
        $Excel.Application.CutCopyMode = $False

        [void]$source.ActiveSheet.Range("A1","D$val1").Copy()
        [void]$Dest.Activate()
        [void]$Dest.ActiveSheet.Range("A1").Select()
        #adding #File Path to Column F 9Last Column
        #$Dest.ActiveSheet.Range("F2:F$val1").Value = $File

    }Else{ 
        #any Subsquent Row other than the first paste
        [void]$source.ActiveSheet.Range("A2","D$val2").Copy()
        [void]$Dest.Activate()
        [void]$Dest.ActiveSheet.Range("A$val3").Select()
        #adding #File Path to Column F 9Last Column
        #$Dest.ActiveSheet.Range("F$val3","F$val4").Value = $File

    }
    
    [void]$Dest.ActiveSheet.Paste()
    $Excel.Application.CutCopyMode = $False
    $Source.Close()

    $Source = $null
    $val1 = $null
    $val2 = $null
    $val3 = $null
    $val4 = $null

    $fileCounter = $fileCounter + 1

}
$Dest.SaveAs([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)+"\$OutputFileName.xlsx",51)
$Dest.close()
$Excel.Quit()

$Files = $null
$Excel = $null
$fileCounter = $null
$Dest = $null
$OutputFileName = $null