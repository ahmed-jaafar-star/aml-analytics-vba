Attribute VB_Name = "DashboardToWord"
Option Explicit

'==============================================================================
' EXPORT DASHBOARD TO WORD ("WU Accounts" report)
'------------------------------------------------------------------------------
' Reads the headline figures and two tables from the active analysis sheet
' and fills a Word template:
'   (Today) (From-To) (Total$) (Trx.no.) (TUC) (TB/W) (TB) (TW) (TopG) (TopG$)
'   (Governorate Table) (Country Table)
' Saves "WU Accounts <period>.docx" next to the workbook.
'==============================================================================

Sub ExportDashboardToWord()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim wdApp As Object, wdDoc As Object, objRange As Object
    Dim templatePath As String, savePath As String
    Dim wsTemp As Worksheet
    
    ' --- 1. CHOOSE THE WORD TEMPLATE ---
    Dim fd As Object
    Set fd = Application.FileDialog(3) ' 3 = msoFileDialogFilePicker
    With fd
        .Title = "Select the Word Template"
        .Filters.Clear
        .Filters.Add "Word Documents", "*.doc; *.docx"
        .AllowMultiSelect = False
        .InitialFileName = "C:\Automated Reports Templates\"
        If .Show = -1 Then
            templatePath = .SelectedItems(1)
        Else
            MsgBox "No template selected. Export cancelled.", vbExclamation
            Exit Sub
        End If
    End With
    
    ' --- VARIABLES TO HOLD THE EXTRACTED DATA ---
    Dim vToday As String, vFromTo As String
    Dim vTotalSusAmt As String, vTrxNo As String, vTUC As String
    Dim vTBW As String, vTB As String, vTW As String
    Dim vTopG As String, vTopGAmt As String
    
    Dim rngFind As Range, rowTotal As Long
    Dim rngCountryTable As Range, rngGovTable As Range
    
    ' Safety variables for the Find method
    Dim rngTotalCountry As Range, rngTotalGov As Range, rngTotalSus As Range
    
    ' Initialize default empty values
    vToday = Format(Date, "dd/MM/yyyy")
    vFromTo = ws.Range("A1").Value
    vTotalSusAmt = "0": vTrxNo = "0": vTUC = "0"
    vTBW = "0": vTB = "0": vTW = "0"
    vTopG = "N/A": vTopGAmt = "0"
    
    Application.ScreenUpdating = False
    
    ' --- 2. EXTRACT HIGH RISK COUNTRIES INFO ---
    Set rngFind = ws.Cells.Find(What:="HIGH RISK COUNTRIES", LookAt:=xlPart, MatchCase:=False)
    If Not rngFind Is Nothing Then
        Set rngTotalCountry = ws.Columns(rngFind.Column).Find(What:="Total", After:=rngFind, LookAt:=xlWhole)
        If Not rngTotalCountry Is Nothing Then
            rowTotal = rngTotalCountry.Row
            Set rngCountryTable = ws.Range(ws.Cells(rngFind.Row, rngFind.Column), ws.Cells(rowTotal, rngFind.Column + 10))
            vTrxNo = Format(ws.Cells(rowTotal, rngFind.Column + 8).Value, "#,##0")
            vTotalSusAmt = Format(ws.Cells(rowTotal, rngFind.Column + 10).Value, "#,##0")
        End If
    End If
    
    ' --- 3. EXTRACT HIGH RISK GOVERNORATES INFO ---
    Set rngFind = ws.Cells.Find(What:="HIGH RISK GOVERNORATES", LookAt:=xlPart, MatchCase:=False)
    If Not rngFind Is Nothing Then
        Set rngTotalGov = ws.Columns(rngFind.Column).Find(What:="Total", After:=rngFind, LookAt:=xlWhole)
        If Not rngTotalGov Is Nothing Then
            rowTotal = rngTotalGov.Row
            Set rngGovTable = ws.Range(ws.Cells(rngFind.Row, rngFind.Column), ws.Cells(rowTotal, rngFind.Column + 6))
            vTopG = ws.Cells(rngFind.Row + 2, rngFind.Column).Value
            vTopGAmt = Format(ws.Cells(rngFind.Row + 2, rngFind.Column + 6).Value, "#,##0")
        End If
    End If
    
    ' --- 4. EXTRACT SUSPICIOUS ACCOUNTS ---
    Set rngFind = ws.Cells.Find(What:="SUSPICIOUS ACCOUNTS (>=", LookAt:=xlPart, MatchCase:=False)
    If Not rngFind Is Nothing Then
        Set rngTotalSus = ws.Columns(rngFind.Column).Find(What:="Total (", After:=rngFind, LookAt:=xlPart)
        If Not rngTotalSus Is Nothing Then
            rowTotal = rngTotalSus.Row
            Dim bankMobileStr As String
            bankMobileStr = ws.Cells(rowTotal, rngFind.Column + 5).Value
            On Error Resume Next
            vTB = Trim(Split(Split(bankMobileStr, ",")(0), "=")(1))
            vTW = Trim(Split(Split(bankMobileStr, ",")(1), "=")(1))
            vTBW = Format(CLng(vTB) + CLng(vTW), "#,##0")
            On Error GoTo 0
        End If
    End If
    
    ' --- 5. EXTRACT UNIQUE SUSPICIOUS SENDERS ---
    Set rngFind = ws.Cells.Find(What:="SUSPICIOUS UNIQUE SENDERS", LookAt:=xlPart, MatchCase:=False)
    If Not rngFind Is Nothing Then
        Dim lastSenderRow As Long
        lastSenderRow = ws.Cells(rngFind.Row + 1, rngFind.Column).End(xlDown).Row
        If lastSenderRow > 1000000 Then lastSenderRow = rngFind.Row + 1
        vTUC = Format(lastSenderRow - (rngFind.Row + 1), "#,##0")
    End If
    
    ' ========================================================
    ' OPEN WORD AND REPLACE TAGS
    ' ========================================================
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then Set wdApp = CreateObject("Word.Application")
    On Error GoTo 0
    
    wdApp.Visible = True
    Set wdDoc = wdApp.Documents.Open(templatePath)
    
    Dim arrFind As Variant, arrRep As Variant
    arrFind = Array("(Today)", "(From-To)", "(Total$)", "(Trx.no.)", "(TUC)", "(TB/W)", "(TB)", "(TW)", "(TopG)", "(TopG$)")
    arrRep = Array(vToday, vFromTo, vTotalSusAmt, vTrxNo, vTUC, vTBW, vTB, vTW, vTopG, vTopGAmt)
    
    Dim f As Long
    For f = LBound(arrFind) To UBound(arrFind)
        Set objRange = wdDoc.Content
        With objRange.Find
            .Text = CStr(arrFind(f))
            .Replacement.Text = CStr(arrRep(f))
            .Forward = True
            .Wrap = 1 ' wdFindContinue
            .Execute Replace:=2 ' wdReplaceAll
        End With
    Next f
    
    ' --- PASTE GOVERNORATE TABLE ---
    Set objRange = wdDoc.Content
    With objRange.Find
        .Text = "(Governorate Table)"
        .Execute
    End With
    If objRange.Find.Found Then
        If Not rngGovTable Is Nothing Then
            Set wsTemp = ws.Parent.Sheets.Add
            rngGovTable.Copy
            wsTemp.Range("A1").PasteSpecial Paste:=xlPasteColumnWidths
            wsTemp.Range("A1").PasteSpecial Paste:=xlPasteAll
            
            ' Delete the top merged header
            wsTemp.Rows(1).Delete Shift:=xlUp
            
            ' Delete Columns B, C, D
            wsTemp.Range("B:D").Delete Shift:=xlToLeft
            
            ' RENAME HEADERS FOR WORD
            wsTemp.Range("B1").Value = "Transactions Count"
            wsTemp.Range("C1").Value = "Unique Senders"
            wsTemp.Range("D1").Value = "Amount"
            
            ' Strip colors, un-bold everything, force black text/borders
            With wsTemp.UsedRange
                .Interior.ColorIndex = xlNone
                .Font.Color = vbBlack
                .Borders.Color = vbBlack
                .Font.Bold = False
            End With
            
            ' Re-apply Bold ONLY to the Header and Footer rows
            wsTemp.Rows(1).Font.Bold = True
            wsTemp.Rows(wsTemp.UsedRange.Rows.count).Font.Bold = True
            
            wsTemp.Columns.AutoFit
            
            wsTemp.UsedRange.Copy
            objRange.PasteExcelTable False, False, False
            
            If objRange.Tables.count > 0 Then
                objRange.Tables(1).AutoFitBehavior 1 ' wdAutoFitContent
                objRange.Tables(1).AutoFitBehavior 2 ' wdAutoFitWindow
            End If
            
            Application.DisplayAlerts = False
            wsTemp.Delete
            Application.DisplayAlerts = True
        Else
            objRange.Text = "No High Risk Governorates identified."
        End If
    End If
    
    ' --- PASTE COUNTRY TABLE ---
    Set objRange = wdDoc.Content
    With objRange.Find
        .Text = "(Country Table)"
        .Execute
    End With
    If objRange.Find.Found Then
        If Not rngCountryTable Is Nothing Then
            Set wsTemp = ws.Parent.Sheets.Add
            rngCountryTable.Copy
            wsTemp.Range("A1").PasteSpecial Paste:=xlPasteColumnWidths
            wsTemp.Range("A1").PasteSpecial Paste:=xlPasteAll
            
            ' Delete the top merged header
            wsTemp.Rows(1).Delete Shift:=xlUp
            
            ' Delete Columns B to G
            wsTemp.Range("B:G").Delete Shift:=xlToLeft
            
            ' RENAME HEADERS FOR WORD
            wsTemp.Range("B1").Value = "Accounts Count"
            wsTemp.Range("C1").Value = "Transactions Count"
            wsTemp.Range("D1").Value = "Unique Senders"
            wsTemp.Range("E1").Value = "Amount"
            
            ' Strip colors, un-bold everything, force black text/borders
            With wsTemp.UsedRange
                .Interior.ColorIndex = xlNone
                .Font.Color = vbBlack
                .Borders.Color = vbBlack
                .Font.Bold = False
            End With
            
            ' Re-apply Bold ONLY to the Header and Footer rows
            wsTemp.Rows(1).Font.Bold = True
            wsTemp.Rows(wsTemp.UsedRange.Rows.count).Font.Bold = True
            
            wsTemp.Columns.AutoFit
            
            wsTemp.UsedRange.Copy
            objRange.PasteExcelTable False, False, False
            
            If objRange.Tables.count > 0 Then
                objRange.Tables(1).AutoFitBehavior 1 ' wdAutoFitContent
                objRange.Tables(1).AutoFitBehavior 2 ' wdAutoFitWindow
            End If
            
            Application.DisplayAlerts = False
            wsTemp.Delete
            Application.DisplayAlerts = True
        Else
            objRange.Text = "No High Risk Countries identified."
        End If
    End If
    
    Application.CutCopyMode = False
    
    ' --- CLEAN DATE FOR FILE NAMING ---
    Dim cleanDate As String
    cleanDate = Replace(vFromTo, "/", "-")
    
    ' --- SAVE THE NEW WORD DOCUMENT ---
    If ActiveWorkbook.Path <> "" Then
        savePath = ActiveWorkbook.Path & "\WU Accounts " & cleanDate & ".docx"
    Else
        savePath = Application.DefaultFilePath & "\WU Accounts " & cleanDate & ".docx"
    End If
    
    wdDoc.SaveAs2 fileName:=savePath, FileFormat:=12  ' 12 = wdFormatDocumentDefault (.docx)
    Application.ScreenUpdating = True
    
    MsgBox "Report Completed !", vbInformation, "Export Complete"
End Sub
