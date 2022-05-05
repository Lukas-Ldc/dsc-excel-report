[CmdletBinding()]
Param(
    [Parameter(Mandatory=$false)]
    [Switch]$Correct #If "Test-DscConfiguration" is "False", try to apply the configuration with "Start-DscConfiguration -CimSession $cim -UseExisting -Wait"
)

$path = Split-Path -parent $PSCommandPath

#XLSX Recovery
$xlsx = $path + "\dsc_report.xlsx"
$objExcel = New-Object -ComObject Excel.Application
$objExcel.Visible = $False
$objExcel.DisplayAlerts = $False
try {$objWorkbook = $objExcel.Workbooks.Open($xlsx)}

#XLSX Creation
catch {
$objWorkbook = $objExcel.Workbooks.Add()
$objWorksheets = $objWorkbook.Worksheets.Item(1)
$objWorksheets.Name = "Global"
$objWorksheets.Range("A1:Z1").Font.Bold = $True
$objWorksheets.Range("A1:Z1").HorizontalAlignment = -4108
$objWorksheets.Cells.Item(1,1) = "DATE"
}

#Getting info
$theWorksheet = $objWorkbook.Worksheets.Item(1)
$theWkshtRow = $theWorksheet.UsedRange.Rows.Count + 1
$wkshnb = $objWorkbook.WorkSheets.Count + 1

#Getting nodes names
$nodes = (Get-Content $path/nodes.txt).Split(",")

#DSC Verification
foreach ($n in $nodes){
    if ($n.Length -gt 0) {

        #Getting node column in global
        $node_column = 0
        for ($i = 1; $i -gt 0; $i++) {
            if ($theWorksheet.Cells.Item(1,$i).Text.Length -gt 0) {
                if ($theWorksheet.Cells.Item(1,$i).Text.equals($n)) {
                    $node_column = $i
                    $i = -1
                }
            } else {
                #New node column in global
                $node_column = $i
                $theWorksheet.Cells.Item(1,$i) = $n
                $i = -1
            }
        }

        #Getting node spreadsheet
        $node_wrksht = 0
        for ($i = 2; $i -lt $wkshnb; $i++) {
            if ($objWorkbook.Worksheets.Item($i).Name.equals($n)) {
                $node_wrksht = $i
                $nodeWorksheet = $objWorkbook.Worksheets.Item($i)
            }
        }
        if ($node_wrksht -eq 0) {
            #New node spreadsheet
            $newWorksheet = $objWorkbook.Worksheets.Add()
            $newWorksheet.Name = $n
            $newWorksheet.Range("A1:Z1").Font.Bold = $True
            $newWorksheet.Range("A1:Z1").HorizontalAlignment = -4108
            $newWorksheet.Cells.Item(1,1) = "DATE"
            $newWorksheet.Cells.Item(1,2) = "DSC STATUS"
            $newWorksheet.Cells.Item(1,3) = "DSC TEST"
            $theWorksheet.Move($newWorksheet)
            $node_wrksht = $wkshnb
            $wkshnb++
            $nodeWorksheet = $objWorkbook.Worksheets.Item(2)
        }

        #Global verification
        $cim = New-CimSession $n
        $date = Get-Date -Format "HH:mm dd/MM/yyyy"
        $dsct = Test-DscConfiguration -CimSession $cim -Detailed
        $dsc_test = ($dsct).InDesiredState
        $dsc_status = (Get-DscConfigurationStatus).Status
        $r_in_state = ($dsct.ResourcesInDesiredState).ResourceId
        $r_not_state = ($dsct.ResourcesNotInDesiredState).ResourceId
        $r_state = $r_in_state + $r_not_state | Sort-Object

        #Trying to correct
        $dsc_test_2 = $dsc_test
        $dsc_status_2 = $dsc_status 
        if ($dsc_test -ne $True -and $Correct) {
            Start-DscConfiguration -CimSession $cim -UseExisting -Wait
            $dsct = Test-DscConfiguration -CimSession $cim -Detailed
            $dsc_test_2 = ($dsct).InDesiredState
            $dsc_status_2 = (Get-DscConfigurationStatus).Status
            $r_in_state_2 = ($dsct.ResourcesInDesiredState).ResourceId
            $r_not_state_2 = ($dsct.ResourcesNotInDesiredState).ResourceId

            $r_change_state = [System.Collections.ArrayList]::new()
            foreach ($n1 in $r_in_state_2) {
                foreach ($n2 in $r_not_state) {
                    if ($n1 -eq $n2) {$r_change_state.Add($n1)}
                }
            }
            foreach ($n1 in $r_not_state_2) {
                foreach ($n2 in $r_in_state) {
                    if ($n1 -eq $n2) {$r_change_state.Add($n1)}
                }
            }
        }
        
        #Writing in global worksheet
        $theWorksheet.Cells.Item($theWkshtRow,1) = $date
        $theWorksheet.Cells.Item($theWkshtRow,$node_column) = $dsc_test
        if ($dsc_test -eq $True) {$theWorksheet.Cells.Item($theWkshtRow,$node_column).Interior.ColorIndex = 4}
        elseif ($dsc_test -ne $dsc_test_2) {
            $theWorksheet.Cells.Item($theWkshtRow,$node_column) = $dsc_test_2
            $theWorksheet.Cells.Item($theWkshtRow,$node_column).Interior.ColorIndex = 5}
        else {$theWorksheet.Cells.Item($theWkshtRow,$node_column).Interior.ColorIndex = 3}

        #Writing in node worksheet
        $nodeWkshtRow = $nodeWorksheet.UsedRange.Rows.Count + 1
        $nodeWorksheet.Cells.Item($nodeWkshtRow,1) = $date
        $nodeWorksheet.Cells.Item($nodeWkshtRow,2) = $dsc_status
        $nodeWorksheet.Cells.Item($nodeWkshtRow,3) = $dsc_test

        if ($dsc_status -eq "Success") {$nodeWorksheet.Cells.Item($nodeWkshtRow,2).Interior.ColorIndex = 4}
        elseif ($dsc_status -ne $dsc_status_2) {
            $nodeWorksheet.Cells.Item($nodeWkshtRow,2) = $dsc_status_2
            $nodeWorksheet.Cells.Item($nodeWkshtRow,2).Interior.ColorIndex = 5}
        else {$nodeWorksheet.Cells.Item($nodeWkshtRow,2).Interior.ColorIndex = 3}

        if ($dsc_test -eq $True) {$nodeWorksheet.Cells.Item($nodeWkshtRow,3).Interior.ColorIndex = 4}
        elseif ($dsc_test -ne $dsc_test_2) {
            $nodeWorksheet.Cells.Item($nodeWkshtRow,3) = $dsc_test_2
            $nodeWorksheet.Cells.Item($nodeWkshtRow,3).Interior.ColorIndex = 5}
        else {$nodeWorksheet.Cells.Item($nodeWkshtRow,3).Interior.ColorIndex = 3}

        $i = 4
        foreach($r in $r_state){
            $nodeWorksheet.Cells.Item($nodeWkshtRow,$i) = $r
            foreach ($ry in $r_in_state) {if ($r -eq $ry){$nodeWorksheet.Cells.Item($nodeWkshtRow,$i).Interior.ColorIndex = 4}}
            foreach ($rn in $r_not_state) {if ($r -eq $rn){$nodeWorksheet.Cells.Item($nodeWkshtRow,$i).Interior.ColorIndex = 3}}
            foreach ($rc in $r_change_state) {if ($r -eq $rc){$nodeWorksheet.Cells.Item($nodeWkshtRow,$i).Interior.ColorIndex = 5}}
            $i++}
        
        #Node finished
        Get-CimSession -ComputerName $n | Remove-CimSession
        $r_change_state = $null | Out-Null
    }
}

#XLSX saving
$objWorkbook.SaveAs($xlsx)
$objWorkbook.Close()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel) | Out-Null
