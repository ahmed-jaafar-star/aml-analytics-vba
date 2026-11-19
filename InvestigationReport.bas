Attribute VB_Name = "InvestigationReport"
Option Explicit

'==============================================================================
' GENERATE INVESTIGATION REPORT (Word) FROM RED-FLAGGED MSISDNs
'------------------------------------------------------------------------------
' - Scans ALL sheets named "Analysis - XXXX" in the ACTIVE workbook
' - Collects flagged MSISDNs: Font.Color = 255 (red) OR pink fill
'   (Interior.Color = 13551615 [RGB 255,199,206] or 13408767)
' - If the COUNTRY NAME itself is red in the HIGH RISK COUNTRIES table
'   (e.g. VIETNAM / CAMBODIA), the WHOLE corridor is treated as suspicious
' - Pulls all SEND_MONEY / SENDMONEY transactions from Sheet1:
'     Col 10 = Governorate | Col 12 = Sender MSISDN | Col 13 = Type
'     Col 14 = Amount      | Col 16 = Destination Country
' - Per country: Total vs Suspicious (Amount + Transactions)
'   + Governorate table (Total Trx / Unique Senders / Total Amount)
'     sorted DESC by transactions, with a bold Total row
' - Key Observations + Conclusion left EMPTY for manual filling
' - Action Taken section with placeholder bullets (auto count inserted)
' - Saves .docx in the SAME folder as the Excel file
'
' NOTE: Uses ActiveWorkbook (macros live in PERSONAL.XLSB)
'       Word late-binding (no reference needed)
'       Never uses negative numbers for Font.Color
'==============================================================================

Private Const RED_FONT As Long = 255            ' pure red font
Private Const PINK_1 As Long = 13551615         ' RGB(255,199,206) - actual pink in file
Private Const PINK_2 As Long = 13408767         ' legacy pink value (checked too)

' Names printed in the report header table
Private Const AML_MANAGER As String = "____________"
Private Const DEPUTY_MANAGER As String = "____________"

' Word constants (late binding)
Private Const wdCollapseEnd As Long = 0
Private Const wdAlignCenter As Long = 1
Private Const wdAlignLeft As Long = 0
Private Const wdFormatXMLDocument As Long = 12
Private Const wdPreferredWidthPoints As Long = 3

Public Sub GenerateInvestigationReport()

    Dim wb As Workbook, wsData As Worksheet, sh As Worksheet
    Dim dataArr As Variant
    Dim lastRow As Long, r As Long, i As Long, j As Long, ci As Long, idx As Long
    Dim nC As Long
    Dim countries() As String
    Dim flagsByCountry As Object, corridorFlag As Object
    Dim flags As Object, isCorr As Boolean
    Dim cel As Range, s As String, v As Variant, fc As Variant
    Dim cName As String, uCountry As String
    Dim tType As String, msisdn As String, gv As String
    Dim amt As Double
    Dim totTrx As Long, susTrx As Long
    Dim totAmt As Double, susAmt As Double
    Dim govTrx As Object, govAmt As Object, govSenders As Object, allSenders As Object
    Dim totalFlagged As Object
    Dim k As Variant
    Dim nG As Long
    Dim gNames() As String, gT() As Long, gU() As Long, gA() As Double
    Dim tmpS As String, tmpL As Long, tmpD As Double
    Dim wdApp As Object, wdDoc As Object, rng As Object, tbl As Object
    Dim invName As String
    Dim periodRaw As String, periodTxt As String
    Dim d1 As Date, d2 As Date, okDates As Boolean, posTo As Long
    Dim fileName As String, savePath As String
    Dim allSus As Boolean

    On Error GoTo ErrHandler

    Set wb = ActiveWorkbook

    On Error Resume Next
    Set wsData = wb.Worksheets("Sheet1")
    On Error GoTo ErrHandler
    If wsData Is Nothing Then
        MsgBox "Sheet1 not found in the active workbook.", vbCritical, "Investigation Report"
        Exit Sub
    End If
    If Len(wb.Path) = 0 Then
        MsgBox "Please save the Excel workbook first (a folder is needed to save the report).", vbCritical, "Investigation Report"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    '--------------------------------------------------------------------------
    ' 1) LOAD SHEET1 INTO MEMORY (columns 1..17)
    '--------------------------------------------------------------------------
    lastRow = wsData.Cells(wsData.Rows.count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Sheet1 has no data.", vbCritical, "Investigation Report"
        GoTo CleanExit
    End If
    dataArr = wsData.Range(wsData.Cells(2, 1), wsData.Cells(lastRow, 17)).Value2

    '--------------------------------------------------------------------------
    ' 2) SCAN ALL "Analysis - XXXX" SHEETS FOR FLAGGED MSISDNs
    '--------------------------------------------------------------------------
    Set flagsByCountry = CreateObject("Scripting.Dictionary")
    Set corridorFlag = CreateObject("Scripting.Dictionary")
    ReDim countries(1 To wb.Worksheets.count)
    nC = 0

    For Each sh In wb.Worksheets
        If UCase$(sh.Name) Like "ANALYSIS - *" Then
            nC = nC + 1
            cName = Trim$(Mid$(sh.Name, InStr(1, sh.Name, "-") + 1))
            countries(nC) = cName

            Set flags = CreateObject("Scripting.Dictionary")
            isCorr = False
            Application.StatusBar = "Scanning: " & sh.Name & " ..."

            For Each cel In sh.UsedRange.Cells
                v = cel.Value2
                If Not IsEmpty(v) Then
                    If Not IsError(v) Then
                        s = CleanDigits(v)
                        If Len(s) >= 10 And Len(s) <= 15 Then
                            ' candidate MSISDN -> check red font / pink fill
                            If IsFlaggedCell(cel) Then
                                If Not flags.Exists(s) Then flags.Add s, True
                            End If
                        ElseIf VarType(v) = vbString Then
                            ' country name written in RED = whole corridor suspicious
                            If UCase$(Trim$(CStr(v))) = UCase$(cName) Then
                                fc = cel.Font.Color
                                If Not IsNull(fc) Then
                                    If fc = RED_FONT Then isCorr = True
                                End If
                            End If
                        End If
                    End If
                End If
            Next cel

            Set flagsByCountry(cName) = flags
            corridorFlag(cName) = isCorr
        End If
    Next sh

    If nC = 0 Then
        MsgBox "No sheets named 'Analysis - XXXX' were found.", vbCritical, "Investigation Report"
        GoTo CleanExit
    End If

    '--------------------------------------------------------------------------
    ' 2b) COUNTRY SELECTION (JOKER MODE)
    '     Empty = ALL countries | or numbers like: 1,3,5
    '--------------------------------------------------------------------------
    Dim pickMsg As String, pickInput As String
    Dim selCountries() As String, nSel As Long
    Dim parts() As String, p As Variant, pn As Long

    pickMsg = "Detected " & nC & " Analysis sheets:" & vbCrLf & vbCrLf
    For i = 1 To nC
        pickMsg = pickMsg & i & " - " & countries(i) & vbCrLf
    Next i
    pickMsg = pickMsg & vbCrLf & "Press OK (empty) = ALL countries" & vbCrLf & _
              "Or type numbers separated by commas (e.g. 1,3,5):"

    pickInput = Trim$(InputBox(pickMsg, "Select Countries", ""))

    ReDim selCountries(1 To nC)
    nSel = 0
    If Len(pickInput) = 0 Then
        For i = 1 To nC
            nSel = nSel + 1
            selCountries(nSel) = countries(i)
        Next i
    Else
        parts = Split(pickInput, ",")
        For Each p In parts
            If IsNumeric(Trim$(CStr(p))) Then
                pn = CLng(Trim$(CStr(p)))
                If pn >= 1 And pn <= nC Then
                    nSel = nSel + 1
                    selCountries(nSel) = countries(pn)
                End If
            End If
        Next p
        If nSel = 0 Then
            MsgBox "No valid country numbers were entered. Report cancelled.", vbExclamation, "Investigation Report"
            GoTo CleanExit
        End If
    End If

    ' Replace the working list with the selected countries only
    For i = 1 To nSel
        countries(i) = selCountries(i)
    Next i
    nC = nSel

    '--------------------------------------------------------------------------
    ' 3) INVESTIGATOR NAME + PERIOD + FILE NAME
    '--------------------------------------------------------------------------
    invName = InputBox("Investigator Name:", "Investigation Report", "")
    If Len(Trim$(invName)) = 0 Then invName = "____________"

    periodRaw = ""
    For Each sh In wb.Worksheets
        If UCase$(sh.Name) Like "ANALYSIS - *" Then
            periodRaw = Trim$(CStr(sh.Range("A1").Value))
            Exit For
        End If
    Next sh

    okDates = False
    periodTxt = periodRaw
    posTo = InStr(1, periodRaw, " to ", vbTextCompare)
    If posTo > 0 Then
        On Error Resume Next
        d1 = CDate(Replace(Trim$(Left$(periodRaw, posTo - 1)), "-", " "))
        d2 = CDate(Replace(Trim$(Mid$(periodRaw, posTo + 4)), "-", " "))
        If Err.Number = 0 Then okDates = True
        Err.Clear
        On Error GoTo ErrHandler
    End If
    If okDates Then
        periodTxt = Day(d1) & " " & MonthName(Month(d1)) & " " & Year(d1) & _
                    " to " & Day(d2) & " " & MonthName(Month(d2)) & " " & Year(d2)
        fileName = "Specific_Countries_" & Day(d1) & "_" & MonthName(Month(d1)) & "_" & Year(d1) & _
                   "_to_" & Day(d2) & "_" & MonthName(Month(d2)) & "_" & Year(d2) & ".docx"
    Else
        fileName = "Specific_Countries_Investigation_Report_" & Format$(Date, "yyyy-mm-dd") & ".docx"
    End If

    '--------------------------------------------------------------------------
    ' 4) OPEN WORD AND BUILD THE DOCUMENT
    '--------------------------------------------------------------------------
    Application.StatusBar = "Building Word report ..."
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    Set wdDoc = wdApp.Documents.Add
    wdDoc.Content.Font.Name = "Calibri"
    wdDoc.Content.Font.Size = 11

    Set rng = wdDoc.Range
    rng.Collapse wdCollapseEnd

    '---------------- Header info table (6 x 2) ----------------
    Set tbl = wdDoc.Tables.Add(rng, 6, 2)
    tbl.Borders.Enable = True
    tbl.Range.Font.Size = 11
    tbl.Columns(1).PreferredWidthType = wdPreferredWidthPoints
    tbl.Columns(1).PreferredWidth = 140
    tbl.Columns(2).PreferredWidthType = wdPreferredWidthPoints
    tbl.Columns(2).PreferredWidth = 320

    SetCell tbl, 1, 1, "Investigator Name", True
    SetCell tbl, 1, 2, invName
    SetCell tbl, 2, 1, "AML Manager", True
    SetCell tbl, 2, 2, AML_MANAGER
    SetCell tbl, 3, 1, "Deputy AML Manager", True
    SetCell tbl, 3, 2, DEPUTY_MANAGER
    SetCell tbl, 4, 1, "Date of Investigation", True
    SetCell tbl, 4, 2, Format$(Date, "dd/mm/yyyy")
    SetCell tbl, 5, 1, "Investigation Period", True
    SetCell tbl, 5, 2, periodTxt
    SetCell tbl, 6, 1, "Investigated Countries", True
    SetCell tbl, 6, 2, nC & " Countries"

    Set rng = wdDoc.Range
    rng.Collapse wdCollapseEnd
    AddPara rng, ""

    '---------------- One section per country ----------------
    Set totalFlagged = CreateObject("Scripting.Dictionary")

    For ci = 1 To nC
        cName = countries(ci)
        uCountry = UCase$(cName)
        Set flags = flagsByCountry(cName)
        allSus = corridorFlag(cName)

        Application.StatusBar = "Writing section: " & cName & " (" & ci & "/" & nC & ") ..."

        '----- compute stats from Sheet1 -----
        totTrx = 0: susTrx = 0: totAmt = 0: susAmt = 0
        Set govTrx = CreateObject("Scripting.Dictionary")
        Set govAmt = CreateObject("Scripting.Dictionary")
        Set govSenders = CreateObject("Scripting.Dictionary")
        Set allSenders = CreateObject("Scripting.Dictionary")

        For r = 1 To UBound(dataArr, 1)
            tType = UCase$(Trim$(CStr(dataArr(r, 13))))
            If tType = "SEND_MONEY" Or tType = "SENDMONEY" Then
                If UCase$(Trim$(CStr(dataArr(r, 16)))) = uCountry Then

                    msisdn = CleanDigits(dataArr(r, 12))
                    If Len(msisdn) = 0 Then msisdn = Trim$(CStr(dataArr(r, 12)))

                    gv = Trim$(CStr(dataArr(r, 10)))
                    If Len(gv) = 0 Then gv = "Unknown"

                    If IsNumeric(dataArr(r, 14)) Then
                        amt = CDbl(dataArr(r, 14))
                    Else
                        amt = 0
                    End If

                    totTrx = totTrx + 1
                    totAmt = totAmt + amt

                    If Not govTrx.Exists(gv) Then
                        govTrx.Add gv, 0
                        govAmt.Add gv, CDbl(0)
                        govSenders.Add gv, CreateObject("Scripting.Dictionary")
                    End If
                    govTrx(gv) = govTrx(gv) + 1
                    govAmt(gv) = govAmt(gv) + amt
                    If Not govSenders(gv).Exists(msisdn) Then govSenders(gv).Add msisdn, True
                    If Not allSenders.Exists(msisdn) Then allSenders.Add msisdn, True

                    If allSus Or flags.Exists(msisdn) Then
                        susTrx = susTrx + 1
                        susAmt = susAmt + amt
                        If Not totalFlagged.Exists(msisdn) Then totalFlagged.Add msisdn, True
                    End If
                End If
            End If
        Next r

        '----- country title + WU analysis lines -----
        AddPara rng, StrConv(cName, vbProperCase), True, 13
        AddPara rng, ""
        AddPara rng, "Western Union Transactions Analysis:", True
        AddPara rng, ""
        AddPara rng, "Total Amount Sent: " & Format$(totAmt, "#,##0") & " IQD."
        AddPara rng, "Suspicious Amount Sent: " & Format$(susAmt, "#,##0") & " IQD."
        AddPara rng, "Total Transactions: " & totTrx & " transactions."
        AddPara rng, "Suspicious Transactions: " & susTrx & " transactions."
        AddPara rng, ""

        '----- governorate table (sorted DESC by transactions) -----
        nG = govTrx.count
        If nG > 0 Then
            ReDim gNames(1 To nG): ReDim gT(1 To nG)
            ReDim gU(1 To nG): ReDim gA(1 To nG)
            idx = 0
            For Each k In govTrx.Keys
                idx = idx + 1
                gNames(idx) = CStr(k)
                gT(idx) = govTrx(k)
                gU(idx) = govSenders(k).count
                gA(idx) = govAmt(k)
            Next k

            For i = 1 To nG - 1
                For j = 1 To nG - i
                    If gT(j) < gT(j + 1) Then
                        tmpS = gNames(j): gNames(j) = gNames(j + 1): gNames(j + 1) = tmpS
                        tmpL = gT(j): gT(j) = gT(j + 1): gT(j + 1) = tmpL
                        tmpL = gU(j): gU(j) = gU(j + 1): gU(j + 1) = tmpL
                        tmpD = gA(j): gA(j) = gA(j + 1): gA(j + 1) = tmpD
                    End If
                Next j
            Next i

            Set tbl = wdDoc.Tables.Add(rng, nG + 2, 4)
            tbl.Borders.Enable = True
            tbl.Range.Font.Size = 10
            tbl.Rows(1).Shading.BackgroundPatternColor = RGB(217, 217, 217)

            SetCell tbl, 1, 1, "Governorate", True
            SetCell tbl, 1, 2, "Total Transactions", True
            SetCell tbl, 1, 3, "Unique Senders", True
            SetCell tbl, 1, 4, "Total Amount", True

            For i = 1 To nG
                SetCell tbl, i + 1, 1, gNames(i)
                SetCell tbl, i + 1, 2, CStr(gT(i))
                SetCell tbl, i + 1, 3, CStr(gU(i))
                SetCell tbl, i + 1, 4, Format$(gA(i), "#,##0")
            Next i

            SetCell tbl, nG + 2, 1, "Total", True
            SetCell tbl, nG + 2, 2, CStr(totTrx), True
            SetCell tbl, nG + 2, 3, CStr(allSenders.count), True
            SetCell tbl, nG + 2, 4, Format$(totAmt, "#,##0"), True

            Set rng = wdDoc.Range
            rng.Collapse wdCollapseEnd
        Else
            AddPara rng, "(No transactions found for this country in Sheet1)"
        End If

        AddPara rng, ""
        AddPara rng, "This table highlights WU transactions originating from each governorate.", , , True

        '----- Key Observations (EMPTY bullets for manual filling) -----
        AddPara rng, "Key Observations:", True
        AddPara rng, "", , , , True
        AddPara rng, "", , , , True
        AddPara rng, "", , , , True
        AddPara rng, "", , , , True

        '----- Conclusion (EMPTY for manual filling) -----
        AddPara rng, "Conclusion:", True
        AddPara rng, ""
        AddPara rng, ""
    Next ci

    '---------------- Action Taken (placeholder bullets) ----------------
    AddPara rng, "Action Taken:", True
    AddPara rng, "Newly flagged suspicious customers across the investigated countries: " & _
                 totalFlagged.count & ".", , , , True
    AddPara rng, "", , , , True
    AddPara rng, "", , , , True

    '--------------------------------------------------------------------------
    ' 5) SAVE AS .DOCX IN THE SAME FOLDER AS THE EXCEL FILE
    '--------------------------------------------------------------------------
    savePath = wb.Path & Application.PathSeparator & fileName

    ' Anti-overwrite: if a report with the same name exists, append _2, _3 ...
    Dim verN As Long, baseName As String
    verN = 1
    baseName = Left$(fileName, Len(fileName) - 5)   ' strip ".docx"
    Do While Len(Dir$(savePath)) > 0
        verN = verN + 1
        savePath = wb.Path & Application.PathSeparator & baseName & "_" & verN & ".docx"
    Loop

    wdDoc.SaveAs2 fileName:=savePath, FileFormat:=wdFormatXMLDocument

    Application.StatusBar = False
    Application.ScreenUpdating = True
    wdApp.Activate

    MsgBox "Investigation report generated successfully:" & vbCrLf & vbCrLf & savePath & _
           vbCrLf & vbCrLf & "Countries: " & nC & vbCrLf & _
           "Flagged suspicious customers: " & totalFlagged.count, _
           vbInformation, "Investigation Report"
    Exit Sub

CleanExit:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Investigation Report"
End Sub

'==============================================================================
' HELPERS
'==============================================================================

' Returns the value as a digits-only string ("" if it contains non-digits)
Private Function CleanDigits(ByVal v As Variant) As String
    Dim s As String, i As Long, ch As String
    If IsNumeric(v) And VarType(v) <> vbString Then
        s = Format$(v, "0")
    Else
        s = Trim$(CStr(v))
    End If
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    CleanDigits = s
End Function

' True if the cell has red font (255) or pink fill (either known pink value)
Private Function IsFlaggedCell(cel As Range) As Boolean
    Dim fc As Variant, ic As Variant
    fc = cel.Font.Color
    If Not IsNull(fc) Then
        If fc = RED_FONT Then
            IsFlaggedCell = True
            Exit Function
        End If
    End If
    ic = cel.Interior.Color
    If Not IsNull(ic) Then
        If ic = PINK_1 Or ic = PINK_2 Then IsFlaggedCell = True
    End If
End Function

' Appends one formatted paragraph at the current insertion range, then collapses
Private Sub AddPara(rng As Object, ByVal txt As String, _
                    Optional ByVal b As Boolean = False, _
                    Optional ByVal sz As Single = 11, _
                    Optional ByVal centered As Boolean = False, _
                    Optional ByVal bullet As Boolean = False)
    rng.InsertAfter txt & vbCr
    rng.Font.Bold = b
    rng.Font.Size = sz
    If centered Then
        rng.ParagraphFormat.Alignment = 1   ' wdAlignParagraphCenter
    Else
        rng.ParagraphFormat.Alignment = 0   ' wdAlignParagraphLeft
    End If
    If bullet Then
        rng.ListFormat.ApplyBulletDefault
    Else
        rng.ListFormat.RemoveNumbers
        rng.ParagraphFormat.LeftIndent = 0
    End If
    rng.Collapse 0                          ' wdCollapseEnd
End Sub

' Writes text into a Word table cell with optional bold
Private Sub SetCell(tbl As Object, ByVal r As Long, ByVal c As Long, _
                    ByVal txt As String, Optional ByVal b As Boolean = False)
    With tbl.cell(r, c).Range
        .Text = txt
        .Font.Bold = b
    End With
End Sub
