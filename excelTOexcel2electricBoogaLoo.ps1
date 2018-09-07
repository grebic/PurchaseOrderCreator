$ErrorActionPreference = "Stop"
$drives = Get-PSDrive
$drvExits = ($drives|where name -EQ p).count 
if (!$drvExists){$drive = New-PSDrive -PSProvider FileSystem -Name P -Root \\ANGDC01\General}

cd "P:\ANG_System_Files"

function Load-Dll
{
    param(
        [string]$assembly
    )

    $driver = $assembly
    $fileStream = ([System.IO.FileInfo] (Get-Item $driver)).OpenRead();
    $assemblyBytes = new-object byte[] $fileStream.Length
    $fileStream.Read($assemblyBytes, 0, $fileStream.Length) | Out-Null;
    $fileStream.Close();
    $assemblyLoaded = [System.Reflection.Assembly]::Load($assemblyBytes);
}

function Get-DatacomparisonObjects
{
    param([Smartsheet.Api.Models.Sheet]$sheet)

    $data = $sheet.Rows | foreach {
        
        [pscustomobject]@{
            
            RowId = $_.Id;
            RowNumber = $_.RowNumber;
            JobNumCol = $_.Cells[0].ColumnId;
            JobNum = $_.Cells[0].Value;
            JobNameCol = $_.Cells[5].ColumnId;
            JobName = $_.Cells[5].Value;
            ProjManCol = $_.Cells[15].ColumnId;
            ProjMan = $_.Cells[15].Value;
            FilePathCol = $_.Cells[45].ColumnId;
            FilePath = $_.Cells[45].Value;
        }                                                  
    } | where {![string]::IsNullOrWhiteSpace($_.JobNum)} 

    return $data                                       
}

function Save-AttachmentToSheetRow
{
    param(
        [long]$sheetId,
        [long]$rowId,
        [System.IO.FileInfo]$file,
        [string]$mimeType
    )

    $result = $client.SheetResources.RowResources.AttachmentResources.AttachFile($sheetId, $rowId, $file.FullName, $mimeType)

    return $result
}

function Merge-ExcelIntoSSpo
{
    $poPoCol            = $poSheet.Columns | where {$_.Title -eq ("PO/WO #")}
    $poJobsCol          = $poSheet.Columns | where {$_.Title -eq ("Job")}
    $poDescCol          = $poSheet.Columns | where {$_.Title -eq ("Description")}
    $poVendor           = $poSheet.Columns | where {$_.Title -eq ("Vendor")}
    $poAssignCol        = $poSheet.Columns | where {$_.Title -eq ("Assigned To")}
    $poDestinationCol   = $poSheet.Columns | where {$_.Title -eq ("Destination")}
    $poGlCol            = $poSheet.Columns | where {$_.Title -eq ("GL No.")}
    $poCostCodeCol      = $poSheet.Columns | where {$_.Title -eq ("Cost Code")}
    $poPoDateCol        = $poSheet.Columns | where {$_.Title -eq ("PO Date")}
    $poPoAmountCol      = $poSheet.Columns | where {$_.Title -eq ("PO Amount")}
    $poInvoiceNumCol    = $poSheet.Columns | where {$_.Title -eq ("Invoice Number")}
    $poInvoiceAmountCol = $poSheet.Columns | where {$_.Title -eq ("Invoice Amount")}


    $poCell = [Smartsheet.Api.Models.Cell]::new()
    $poCell.ColumnId     = $poPoCol.Id
    $poCell.Value        = $fullPO
    
    $jobsCell = [Smartsheet.Api.Models.Cell]::new()
    $jobsCell.ColumnId   = $poJobsCol.Id
    $JobsCell.Value      =  if ($thisJobName -ne $null){$thisJobName} else {[string]::Empty}
    
    $descCell = [Smartsheet.Api.Models.Cell]::new()
    $descCell.ColumnId   = $poDescCol.Id
    $descCell.Value      =  if ($object.Desc -ne $null){$object.Desc} else {[string]::Empty}
    
    $vendCell = [Smartsheet.Api.Models.Cell]::new()
    $vendCell.COlumnId    = $poVendor.Id
    $vendCell.Value       = $object.Vendor

    $AssignCell = [Smartsheet.Api.Models.Cell]::new()
    $AssignCell.ColumnId = $poAssignCol.Id
    $AssignCell.Value    =  if ($object.ProjMan -ne $null){$object.ProjMan} else {[string]::Empty}

    #$DestinationCell = [Smartsheet.Api.Models.Cell]::new()
    #$DestinationCell.ColumnId = $poDestinationCol.Id
    #$DestinationCell.Value    =  if ( -ne $null){} else {[string]::Empty}

    #$gLCell = [Smartsheet.Api.Models.Cell]::new()
    #$gLCell.ColumnId = $poGlCol.Id
    #$gLCell.Value    =  if ($object.Gl -ne $null){$object.Gl} else {[string]::Empty}

    $costCodeCell = [Smartsheet.Api.Models.Cell]::new()
    $costCodeCell.ColumnId = $poCostCodeCol.Id
    $costCodeCell.Value    =  if ($object.CostCode -ne $null){$object.CostCode} else {[string]::Empty}

    #$poDateCell = [Smartsheet.Api.Models.Cell]::new()
    #$poDateCell.ColumnId = $poPoDateCol.Id
    #$poDateCell.Value    =  if ( -ne $null){} else {[string]::Empty}

    $poAmountCell = [Smartsheet.Api.Models.Cell]::new()
    $poAmountCell.ColumnId = $poPoAmountCol.Id
    $poAmountCell.Value    =  if ($object.PoAmmoun -ne $null){$object.PoAmmoun} else {[string]::Empty}

    #$invoiceNumCell = [Smartsheet.Api.Models.Cell]::new()
    #$invoiceNumCell.ColumnId = $poInvoiceNumCol.Id
    #$invoiceNumCell.Value    =  if ( -ne $null){} else {[string]::Empty}

    #$invoiceAmountCell = [Smartsheet.Api.Models.Cell]::new()
    #$invoiceAmountCell.ColumnId = $poInvoiceAmountCol.Id
    #$invoiceAmountCell.Value    =  if ($object.InvAmm -ne $null){$object.InvAmm} else {[string]::Empty}

    $row = [Smartsheet.Api.Models.Row]::new()
    $row.ToTop = $true 
    $row.Cells = [Smartsheet.Api.Models.Cell[]]@($poCell,$jobsCell,$descCell,$vendCell,$AssignCell,$costCodeCell,$poAmountCell) 
    
    $newRow = $client.SheetResources.RowResources.AddRows($poId, [Smartsheet.Api.Models.Row[]]@($row))

    $pdfFile = "$PDFpath.pdf"

    $result = Save-AttachmentToSheetRow -sheetId $poId -rowId $newRow.Id -file $pdfFile -mimeType "application/pdf"
}

Load-Dll ".\smartsheet-csharp-sdk.dll"                     
Load-Dll ".\RestSharp.dll"
Load-Dll ".\Newtonsoft.Json.dll"
Load-Dll ".\NLog.dll"
$token      = "e41266qmwuasa15w9rwe5321ob"
$smartsheet = [Smartsheet.Api.SmartSheetBuilder]::new()
$builder    = $smartsheet.SetAccessToken($token)
$client     = $builder.Build()
$includes   =  @([Smartsheet.Api.Models.SheetLevelInclusion]::ATTACHMENTS)
$includes   = [System.Collections.Generic.List[Smartsheet.Api.Models.SheetLevelInclusion]]$includes 
$dataId     = "1549079680444292"
$dataSheet  = $client.SheetResources.GetSheet($dataId, $includes, $null, $null, $null, $null, $null, $null);
$poId       = "5299005023381380"
$poSheet    = $client.SheetResources.GetSheet($poId, $includes, $null, $null, $null, $null, $null, $null);
$dataSheetCOs = Get-DatacomparisonObjects $dataSheet

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form1 = New-Object System.Windows.Forms.Form  
$form1.Text = 'All New Glass'
$form1.Size = [System.Drawing.Size]::new(360,225)
$form1.StartPosition = 'CenterScreen'

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = [System.Drawing.Point]::new(130,140)
$OKButton.Size = [System.Drawing.Size]::new(75,23)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form1.AcceptButton = $OKButton
$form1.Controls.Add($OKButton)

$label1 = New-Object System.Windows.Forms.Label
$label1.Location = [System.Drawing.Point]::new(10,20)
$label1.Size = [System.Drawing.Size]::new(380,20)
$label1.Text = 'To create a Purchase Order, please enter the Job Number below:'
$form1.Controls.Add($label1)

$textBox1 = New-Object System.Windows.Forms.TextBox
$textBox1.Location = [System.Drawing.Point]::new(10,40)
$textBox1.Size = [System.Drawing.Size]::new(320,20)
$form1.Controls.Add($textBox1)

$label2 = New-Object System.Windows.Forms.Label
$label2.Location = [System.Drawing.Point]::new(10,80)
$label2.Size = [System.Drawing.Size]::new(380,20)
$label2.Text = 'Enter a brief description of this Purchase Order below:'
$form1.Controls.Add($label2)

$textBox2 = New-Object System.Windows.Forms.TextBox
$textBox2.Location = [System.Drawing.Point]::new(10,100)
$textBox2.Size = [System.Drawing.Size]::new(320,20)
$form1.Controls.Add($textBox2)

$form1.Topmost = $true

$form1.Add_Shown({$textBox1.Select()})
$result1 = $form1.ShowDialog()

if ($result1 -eq [System.Windows.Forms.DialogResult]::OK)
{
    [int]$firstEnteredJobNum = $textBox1.Text
    [int]$firstEnteredJobNum

    $firstEnteredDescription = $textBox2.Text
    $firstEnteredDescription
}

foreach ($dataSheetCO in $dataSheetCOs)
{
    $found = $false

    $newData = $([string]$dataSheetCO.JobNum).split("-",2)[0]
    if ($newData -eq $firstEnteredJobNum)
    {
       $thisJobName = $dataSheetCO.JobName

       if (![string]::IsNullOrWhiteSpace($dataSheetCO.FilePath))
       {
            $found = $true
       }

       if($found)
       {
            $chosenPOlog = $dataSheetCO.FilePath
            break
       }
    }
}

if (!$found)  
{
    $form2 = New-Object System.Windows.Forms.Form
    $form2.Text = 'All New Glass'
    $form2.Size = [System.Drawing.Size]::new(625,100)
    $form2.StartPosition = 'CenterScreen'

    $label3 = New-Object System.Windows.Forms.Label
    $label3.Location = [System.Drawing.Point]::new(10,20)
    $label3.Size = [System.Drawing.Size]::new(600,20)
    $label3.Text = "PO NUMBER $firstEnteredJobNum IS NOT IN THE SYSTEM YET.  PLEASE CONTACT ERIC GRECHKO TO HAVE IT ADDED."
    $form2.Controls.Add($label3)

    $form2.Topmost = $true
    $result2 = $form2.ShowDialog()


        $ol = New-Object -comObject Outlook.Application
        $mail = $ol.CreateItem(0)
        $mail.To = "ericg@allnewglass.com"
        $mail.Subject = "Missing PO In System"
        $mail.Body = "PO# $firstEnteredJobNum needs to be added to the PO system.`n`n Enter any further coments or questions below:"
        $inspector = $mail.GetInspector
        $inspector.Activate()
        
        start https://app.smartsheet.com/b/home?lx=eNtoesd7A6eqcIAMsROIUA

    exit 
}

#ACTIVATING THE TWO EXCEL INSTANCES
    $visibleExcel = New-Object -ComObject excel.application 
    $visibleExcel.visible = $true 
    $visibleExcel.Left = 0
    $visibleExcel.Top = 0
    $visibleExcel.DisplayFullScreen = $true

$PoLogWB = $visibleExcel.Workbooks.open("$chosenPOlog")
$PoLogWS = $PoLogWB.Worksheets.item("po log") 
$PoLogWS.activate() | Out-Null

$PoFormWB = $visibleExcel.Workbooks.open("P:\ANG_System_Files\commonFormsUsedInScripts\ANG PO form-master.xlsx") 
$PoFormWS = $PoFormWB.WorkSheets.item("po") 
$PoFormWS.activate() | Out-Null 
    
    $visibleExcel.DisplayFullScreen = $true

$sheet2ranges = $PoLogWS.Range("C1","C$($PoLogWS.UsedRange.Rows.Count)")
$columnMax = $visibleExcel.WorksheetFunction.Max($sheet2ranges)
$newGeneratedPO = if([string]::IsNullOrWhiteSpace($columnMax)){1} else {$columnMax + 1}
$newUsedPO = if ($newGeneratedPO -gt "99"){"-$newGeneratedPO"} elseif ($newGeneratedPO -le "09"){"-00$newGeneratedPO"} else {"-0$newGeneratedPO"} 

$PoFormWS.Cells.Item(2, 8) = "$firstEnteredJobNum" 
$PoFormWS.Cells.Item(2, 9) = "$newUsedPO"
$PoFormWS.Cells.Item(4, 8) = "$thisJobName"
$fullPO  = ($PoFormWS.Rows[2].Columns[8].Text) + ($PoFormWS.Rows[2].Columns[9].Text) + ($PoFormWS.Rows[2].Columns[10].Text) 

    $hiddenExcel = New-Object -ComObject excel.application 
    $hiddenExcel.visible = $false

##BEGIN BUTTON FOR FIRST SHEET
$form = New-Object System.Windows.Forms.Form
$form.FormBorderStyle = "None"
$form.StartPosition = "Manual"
$form.Location.X = 0
$form.Location.Y = 0
$form.TopLevel = $false
$form.Topmost = $false
$form.Text = 'All New Glass'

$form.Size = New-Object System.Drawing.Size(550,75)

$EMAILButton = New-Object System.Windows.Forms.Button
$EMAILButton.Location = [System.Drawing.Point]::new(35,40)
$EMAILButton.Size = [System.Drawing.Size]::new(150,25)
$EMAILButton.Text = "SAVE and EMAIL PDF"
$EMAILButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $EMAILButton
$form.Controls.Add($EMAILButton)

$PDFButton = New-Object System.Windows.Forms.Button
$PDFButton.Location = [System.Drawing.Point]::new(205,40)
$PDFButton.Size = [System.Drawing.Size]::new(150,25)
$PDFButton.Text = "SAVE and EDIT PDF"
$PDFButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $PDFButton
$form.Controls.Add($PDFButton)

$CANCELButton = New-Object System.Windows.Forms.Button
$CANCELButton.Location = [System.Drawing.Point]::new(375,40)
$CANCELButton.Size = [System.Drawing.Size]::new(150,25)
$CANCELButton.Text = "CANCEL"
$CANCELButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $CANCELButton
$form.Controls.Add($CANCELButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = [System.Drawing.Point]::new(10,20)
$label.Size = [System.Drawing.Size]::new(550,30)
$label.Text = 'CLICK HERE WHEN YOU HAVE COMPLETED FILLING OUT THIS FORM OR CLICK CANCEL TO EXIT.'
$form.Controls.Add($label)

$action = $null

$EMAILButton.Add_Click({$script:action = 'email'})
$PDFButton.Add_Click({$script:action = 'pdf'})
$CANCELButton.Add_Click({
    $form.Close()
    $PoLogWB.Close($false) | Out-Null 
    $PoFormWB.Close($false) | Out-Null  
    $visibleExcel.Quit() | Out-Null
    $hiddenExcel.Quit() | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($PoLogWB)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($PoFormWB)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($visibleExcel)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($hiddenExcel)
    Stop-Process -Id $pid
})

$form.TopLevel = $true
$form.Topmost = $true

$result = $form.ShowDialog()

$PoFormWB.SaveAs("P:\ANG_System_Files\commonFormsUsedInScripts\TEMP\$fullPO.xlsx")
$PoLogWB.Close() | Out-Null 
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($PoLogWB)
$PoFormWB.Close() | Out-Null  
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($PoFormWB)

$PoFormWB = $hiddenExcel.Workbooks.open("P:\ANG_System_Files\commonFormsUsedInScripts\TEMP\$fullPO.xlsx") 
$PoFormWS = $PoFormWB.WorkSheets.item("po") 
$PoFormWS.activate() | Out-Null 

$POnumCollected  = ($PoFormWS.Rows[2].Columns[8].Text)+($PoFormWS.Rows[2].Columns[9].Text)+($PoFormWS.Rows[2].Columns[10].Text)
$VendorCollected = $PoFormWS.Rows[10].Columns[2].Text
$firstPOnum  = $POnumCollected + "_"
$firstVendor = $VendorCollected + "_"
$poPath = $chosenPOlog.Substring(0, $chosenPOlog.lastIndexOf('\'))

$NewPoFromWSname = ($firstPOnum) + ($firstVendor) + ("$firstEnteredDescription") 
New-Item -ItemType directory -Path "$poPath\$NewPoFromWSname" 

$PoFormWB.SaveAs("$poPath\$NewPoFromWSname\$NewPoFromWSname.xlsx")
##END BUTTON FOR FIRST SHEET

#CONVERT TO PDF
$xlFixedFormat = "Microsoft.Office.Interop.Excel.xlFixedFormatType" -as [type]
$PDFpath = "$poPath\$NewPoFromWSname\$NewPoFromWSname"
$PoFormWB.WorkSheets.Item(1).ExportAsFixedFormat($xlFixedFormat::xlTypePDF,$PDFpath) 
$PoFormWB.Close() | Out-Null  
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($PoFormWB)

#END OF CONVERT TO PDF

#THIS RE-OPENS THE FIRST EXCEL SHEET TO TAKE DATA FROM IT FOR SECOND SHEET
$ReOpenPoFormWB = $hiddenExcel.Workbooks.open("$poPath\$NewPoFromWSname\$NewPoFromWSname.xlsx") 
$ReOpenPoFormWS = $ReOpenPoFormWB.WorkSheets.item("po") 
$ReOpenPoFormWS.activate() | Out-Null 

$logPOnum = ($newUsedPO).Split("-")[-1]

$object = [pscustomobject] @{
    CurrDate = (Get-Date).Date.ToString('MM/dd/yy'); 
    Job      = $firstEnteredJobNum;                         
    Po       = [int]$logPOnum;  
    Initials = $ReOpenPoFormWS.Rows[2].Columns[10].Text;    
    Gl       = $ReOpenPoFormWS.Rows[19].Columns[7].Text;
    CostCode = $ReOpenPoFormWS.Rows[19].Columns[1].Text;
    Vendor   = $ReOpenPoFormWS.Rows[10].Columns[2].Text;
    Desc     = $firstEnteredDescription; 
    ProjMan  = $ReOpenPoFormWS.Rows[16].Columns[4].Text;
    PoAmmoun = $ReOpenPoFormWS.Rows[52].Columns[10].Text;
    InvDate  = $ReOpenPoFormWS.Rows[16].Columns[1].Text;
    InvAmm   = $ReOpenPoFormWS.Rows[52].Columns[10].Text;
    }

$object | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation | Clip

$ReOpenPoFormWB.Close() | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ReOpenPoFormWB)
#END OF THIS RE-OPENS THE FIRST EXCEL SHEET TO TAKE DATA FROM IT FOR SECOND SHEET


$PoLogWB = $hiddenExcel.Workbooks.open("$chosenPOlog")
$PoLogWS = $PoLogWB.Worksheets.item("po log") 
$PoLogWS.activate() | Out-Null

$lastRow = $($hiddenExcel.ActiveSheet.UsedRange.Rows)[-1]
$firstBlankRow = $($hiddenExcel.ActiveSheet.UsedRange.Rows)[-1].Row + 1
$hiddenExcel.ActiveSheet.Range("A$firstBlankRow").Activate()

$PoLogWS.Cells.Item($firstBlankRow, 1)  = if ($object.CurrDate -ne $null){$object.CurrDate} else {[string]::Empty}  
$PoLogWS.Cells.Item($firstBlankRow, 2)  = if ($object.Job -ne $null){$object.Job} else {[string]::Empty}
$PoLogWS.Cells.Item($firstBlankRow, 3)  = if ($object.Po -ne $null){$object.Po} else {[string]::Empty}
$PoLogWS.Cells.Item($firstBlankRow, 4)  = if ($object.Initials -ne $null){$object.Initials} else {[string]::Empty}
$PoLogWS.Cells.Item($firstBlankRow, 5)  = if ($object.Gl -ne $null){$object.Gl} else {[string]::Empty}
$PoLogWS.Cells.Item($firstBlankRow, 6)  = if ($object.CostCode -ne $null){$object.CostCode} else {[string]::Empty}
$PoLogWS.Cells.Item($firstBlankRow, 7)  = if ($object.Vendor -ne $null){$object.Vendor} else {[string]::Empty}
$PoLogWS.Cells.Item($firstBlankRow, 8)  = if ($object.Desc -ne $null){$object.Desc} else {[string]::Empty}
$PoLogWS.Cells.Item($firstBlankRow, 9)  = if ($object.PoAmmoun -ne $null){$object.PoAmmoun} else {[string]::Empty}
#$PoLogWS.Cells.Item($firstBlankRow, 10) = if ($object.InvDate -ne $null){$object.InvDate} else {[string]::Empty}
#$PoLogWS.Cells.Item($firstBlankRow, 12) = if ($object.InvAmm -ne $null){$object.InvAmm} else {[string]::Empty}

$PoLogWB.Save() | Out-Null
$PoLogWB.Close() | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($PoLogWB) 

#CLOSING TWO EXCEL INSTANCES
    $visibleExcel.Quit() | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($visibleExcel) 
    $hiddenExcel.Quit() | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($hiddenExcel) 

Merge-ExcelIntoSSpo

start "$poPath\$NewPoFormWSname"

switch($action)
{
    'email' {
        $ol1 = New-Object -comObject Outlook.Application
        $attchPath1 = "$PDFpath.pdf"
        $mail1 = $ol1.CreateItem(0)
        $mail1.Subject = "Purchase Order $fullPO"
        $mail1.Attachments.Add($attchPath1)
        $inspector1 = $mail1.GetInspector
        $inspector1.Activate();
        break
    }
    'pdf' {
        start "$PDFpath.pdf"; break
    }
}

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($PoFormWS)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ReOpenPoFormWS)