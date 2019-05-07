#####################################################################################
#
# Set-XLSX_Output.psm1
# This function creates the Excel header body and then formats it.
#
# Parameters Used: 
# $IssueKey
#
# Author: https://github.com/CBHue/
#  
# Version: 1.0
#
#####################################################################################

function Create-XLSX {
    # Set up the Excel workbook ... 
    $script:xlsx=New-Object -ComObject "Excel.Application"
    $xlsx.SheetsInNewWorkbook = 1
    $wBook=$xlsx.Workbooks.Add()
    $script:wSheet=$wBook.ActiveSheet
    $wSheet.Name = “Security POAM Matrix”
    $script:cells=$wSheet.Cells
}

function Insert-Header {
    # Define some variables to control navigation
    $row=1
    $col=1
    
    # insert and format headers
    $h = 0
    $hSize = "20", "10", "10", "20"
    $a = @(("IP Address", "Port", "Protocol", "Service Name"))
    $range = $wSheet.Range("A$row" , "M$row")    
    $range.Value2 = $a
    
    ($hSize) | foreach {
        $cells.item($row,$col).columnWidth = $hSize[$h]
        $h++
        $col++
    }
}

function Insert-Body {
    if ((get-pscallstack |select -last 2 |select -expa arguments -first 1) -match "debug"){ $debugpreference="continue"}
    write-output "Welcome to the XLSX Body writer ... "
    $row = $wSheet.UsedRange.Rows.Count

    foreach ($x in $IssueKey.GetEnumerator()| sort -Property name){
        Write-Debug "+ $($sw.Elapsed) Risk $($xy): inserting $($loopER.count) rows ..."
        #    # Create the array
        $array = New-Object 'object[,]' 1,4
        $array[0,0]  = $IssueKey[$x.key]."ip"
        $array[0,1]  = $IssueKey[$x.key]."port"
        $array[0,2]  = $IssueKey[$x.key]."proto"
        $array[0,3]  = $IssueKey[$x.key]."svc_name"

        $row++
        $range = $wSheet.Range("A$row" , "M$row")
        #    # This is the bottle neck ... maybe convert from CSV is faster ...    
        $range.Value2 = $array
            
        #    $t = $IssueThreat[$x].tostring()
        #    $c = $cells.item($row,2).interior

        #    switch -regex  ($t){ 
        #        Critical { $cells.item($row,2).font.colorIndex=2
        #                   $c.colorIndex=9;  break } 
        #        High     { $c.colorIndex=3;  break } 
        #        Medium   { $c.colorIndex=45; break }  
        #        Low      { $c.colorIndex=6;  break } 
        #        FAILED   { $c.colorIndex=3;  break }  
        #        Error    { $c.colorIndex=3;  break }  
        #        PASSED   { $c.colorIndex=4;  break } 
        #        default  { $c.colorIndex=40        } 
        #    }
        #}
    }
}
    
function Format-XLSX {    
    # Formating Sheet         
    $page = $wSheet.UsedRange 
    $rows = $wSheet.UsedRange.Rows.Count
    $col  = $wSheet.UsedRange.Columns.Count
    $page.WrapText = $True
    $lineStyle = "microsoft.office.interop.excel.xlLineStyle" -as [type]
    $page.borders.LineStyle = $lineStyle::xlContinuous
    $XLTypes = Add-Type -AssemblyName "Microsoft.Office.Interop.Excel" -PassThru
    $VAlign = $xltypes | ? {$_.name -eq "XlVAlign"} # xlVAlignTop, xlVAlignJustify, xlVAlignCenter, xlVAlignBottom
    $page.VerticalAlignment = $VAlign::XlVAlignTop
    $page.rowHeight = 20
    
    # Format Body
    $wSheet.Range("B2","B$rows").font.bold=$True
    $wSheet.Range("C2","C$rows").font.bold=$True
    
    # Format Header
    $head = $wSheet.Cells.Item(1,1).EntireRow
    $head.font.bold=$True
    $head.rowHeight = 20    
    $head.font.size=12
    $head.font.colorIndex=1
    $head.interior.colorIndex=10
    $head.Font.Name = "Calibri"
    $xlConstants = "microsoft.office.interop.excel.Constants" -as [type] 
    $head.HorizontalAlignment = $xlConstants::xlCenter
    $xlsx.Activewindow.Zoom = 90      
    $xlsx.Rows.Item("2:2").Select() | Out-Null;
    $xlsx.ActiveWindow.FreezePanes = $True
    # ok all the data is inserted show the doc
    $xlsx.Visible=$True
}
#
# End: Set-XLSX_Output.psm1
#