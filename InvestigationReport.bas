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
' - Key Observations + Conclusion AUTO-DRAFTED from the data
'   (summary %, top governorates, max-limit pattern [WU_LIMIT const,
'    currently 2,000,000 IQD] with real
'    examples, sender nationality, unique bank/wallet accounts, purpose
'    concentration) - always REVIEW before sending
'   + network indicators over UNIQUE customers: new-wallet mule onboarding
'     (KYC age < 30 days, incl. same-day first transfer), recruiter
'     concentration (system self-registration IDs excluded via
'     EXCLUDED_RECRUITERS const), common-receiver funnel, round-amount
'     pattern, cross-governorate coordination, per-customer intensity
'   + all tables centered on page with vertically centered cells
' - Action Taken section with placeholder bullets (auto count inserted)
' - Saves .docx in the SAME folder as the Excel file
'
' NOTE: Uses ActiveWorkbook (macros live in PERSONAL.XLSB)
'       Word late-binding (no reference needed)
'       Never uses negative numbers for Font.Color
'==============================================================================

' >>> TRANSFER LIMIT SETTING - change here when the transfer limit changes <<<
Private Const WU_LIMIT As Double = 2000000      ' max allowed transfer limit (IQD)
Private Const WU_NEAR_BAND As Double = 50000    ' "immediately below the limit" band

' >>> SYSTEM RECRUITER IDs (app self-registration wallets, NOT real agents) <<<
' These are ignored by the recruiter-concentration observation. To add more,
' keep the format: comma before and after every ID.
Private Const EXCLUDED_RECRUITERS As String = ","   ' e.g. ",11111111,22222222,"

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
    Dim nc As Long
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

    ' --- Key Observations engine ---
    Dim susMaxCnt As Long, susNearCnt As Long
    Dim natIraq As Long, natDest As Long, natOther As Long
    Dim bankTrx As Long, walTrx As Long, cashTrx As Long
    Dim bankAccts As Object, walAccts As Object, exNear As Object, purpCnt As Object
    Dim natS As String, delivS As String, acctS As String, purpS As String
    Dim exNearTxt As String, obsTxt As String, conclTxt As String
    Dim topPurp As String, topPurpN As Long
    Dim sNat As Object, sKyc As Object, sMinSend As Object, sRecr As Object
    Dim recvSenders As Object, recvName As Object, recrCust As Object
    Dim uniqSus As Long, newWallets As Long, topRecrN As Long, topRecvN As Long
    Dim topRecr As String, recvDisplay As String, recvKey As String, rnameS As String
    Dim kycV As Variant, sendV As Variant, recrS As String, mKey As Variant
    Dim susGovs As Object
    Dim roundCnt As Long, sameDayW As Long
    Dim avgSus As Double, avgAll As Double

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
    dataArr = wsData.Range(wsData.Cells(2, 1), wsData.Cells(lastRow, 29)).Value2   ' 3=KYC date, 4=Recruiter, 5=Nationality, 15=Receiver, 18=SendAt, 22=Purpose, 25=Delivery, 26=Bank acct, 29=Metaphone

    '--------------------------------------------------------------------------
    ' 2) SCAN ALL "Analysis - XXXX" SHEETS FOR FLAGGED MSISDNs
    '--------------------------------------------------------------------------
    Set flagsByCountry = CreateObject("Scripting.Dictionary")
    Set corridorFlag = CreateObject("Scripting.Dictionary")
    ReDim countries(1 To wb.Worksheets.count)
    nc = 0

    For Each sh In wb.Worksheets
        If UCase$(sh.Name) Like "ANALYSIS - *" Then
            nc = nc + 1
            cName = Trim$(Mid$(sh.Name, InStr(1, sh.Name, "-") + 1))
            countries(nc) = cName

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

    If nc = 0 Then
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

    pickMsg = "Detected " & nc & " Analysis sheets:" & vbCrLf & vbCrLf
    For i = 1 To nc
        pickMsg = pickMsg & i & " - " & countries(i) & vbCrLf
    Next i
    pickMsg = pickMsg & vbCrLf & "Press OK (empty) = ALL countries" & vbCrLf & _
              "Or type numbers separated by commas (e.g. 1,3,5):"

    pickInput = Trim$(InputBox(pickMsg, "Select Countries", ""))

    ReDim selCountries(1 To nc)
    nSel = 0
    If Len(pickInput) = 0 Then
        For i = 1 To nc
            nSel = nSel + 1
            selCountries(nSel) = countries(i)
        Next i
    Else
        parts = Split(pickInput, ",")
        For Each p In parts
            If IsNumeric(Trim$(CStr(p))) Then
                pn = CLng(Trim$(CStr(p)))
                If pn >= 1 And pn <= nc Then
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
    nc = nSel

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
    tbl.Rows.Alignment = 1                       ' table centered on the page
    tbl.Range.Cells.VerticalAlignment = 1        ' cell content vertically centered
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
    SetCell tbl, 6, 2, nc & " Countries"

    Set rng = wdDoc.Range
    rng.Collapse wdCollapseEnd
    AddPara rng, ""

    '---------------- One section per country ----------------
    Set totalFlagged = CreateObject("Scripting.Dictionary")

    For ci = 1 To nc
        cName = countries(ci)
        uCountry = UCase$(cName)
        Set flags = flagsByCountry(cName)
        allSus = corridorFlag(cName)

        Application.StatusBar = "Writing section: " & cName & " (" & ci & "/" & nc & ") ..."

        '----- compute stats from Sheet1 -----
        totTrx = 0: susTrx = 0: totAmt = 0: susAmt = 0
        susMaxCnt = 0: susNearCnt = 0
        natIraq = 0: natDest = 0: natOther = 0
        bankTrx = 0: walTrx = 0: cashTrx = 0
        Set bankAccts = CreateObject("Scripting.Dictionary")
        Set walAccts = CreateObject("Scripting.Dictionary")
        Set exNear = CreateObject("Scripting.Dictionary")
        Set purpCnt = CreateObject("Scripting.Dictionary")
        Set sNat = CreateObject("Scripting.Dictionary")
        Set sKyc = CreateObject("Scripting.Dictionary")
        Set sMinSend = CreateObject("Scripting.Dictionary")
        Set sRecr = CreateObject("Scripting.Dictionary")
        Set recvSenders = CreateObject("Scripting.Dictionary")
        Set recvName = CreateObject("Scripting.Dictionary")
        Set susGovs = CreateObject("Scripting.Dictionary")
        roundCnt = 0: sameDayW = 0
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

                        ' --- observation stats (suspicious transactions only) ---
                        ' (a) max-limit pattern (limit = WU_LIMIT constant)
                        If amt = WU_LIMIT Then
                            susMaxCnt = susMaxCnt + 1
                        ElseIf amt >= WU_LIMIT - WU_NEAR_BAND And amt < WU_LIMIT Then
                            susNearCnt = susNearCnt + 1
                            If exNear.count < 2 Then
                                If Not exNear.Exists(Format$(amt, "#,##0")) Then
                                    exNear.Add Format$(amt, "#,##0"), True
                                End If
                            End If
                        End If

                        ' (b) per-customer profile (UNIQUE customers, not transactions)
                        If Not sNat.Exists(msisdn) Then
                            sNat.Add msisdn, UCase$(Trim$(CStr(dataArr(r, 5))))
                            sKyc.Add msisdn, dataArr(r, 3)      ' KYC created (serial date)
                            sRecr.Add msisdn, Trim$(CStr(dataArr(r, 4)))
                            sMinSend.Add msisdn, dataArr(r, 18) ' first suspicious SEND_AT
                        Else
                            If IsNumeric(dataArr(r, 18)) And IsNumeric(sMinSend(msisdn)) Then
                                If dataArr(r, 18) < sMinSend(msisdn) Then sMinSend(msisdn) = dataArr(r, 18)
                            End If
                        End If

                        ' (b2) receiver funnel key: bank acct -> metaphone -> receiver name
                        recvKey = Trim$(CStr(dataArr(r, 26)))
                        If Len(recvKey) = 0 Then recvKey = Trim$(CStr(dataArr(r, 29)))
                        If Len(recvKey) = 0 Then recvKey = UCase$(Trim$(CStr(dataArr(r, 15))))
                        If Len(recvKey) > 0 Then
                            If Not recvSenders.Exists(recvKey) Then
                                recvSenders.Add recvKey, CreateObject("Scripting.Dictionary")
                                rnameS = Trim$(CStr(dataArr(r, 15)))
                                If Len(rnameS) = 0 Then rnameS = recvKey
                                recvName.Add recvKey, rnameS
                            End If
                            If Not recvSenders(recvKey).Exists(msisdn) Then recvSenders(recvKey).Add msisdn, True
                        End If

                        ' (c) delivery channel (col 25) + unique accounts (col 26)
                        delivS = UCase$(Trim$(CStr(dataArr(r, 25))))
                        acctS = Trim$(CStr(dataArr(r, 26)))
                        If InStr(1, delivS, "BANK") > 0 Then
                            bankTrx = bankTrx + 1
                            If Len(acctS) > 0 Then
                                If Not bankAccts.Exists(acctS) Then bankAccts.Add acctS, True
                            End If
                        ElseIf InStr(1, delivS, "MOBILE") > 0 Then
                            walTrx = walTrx + 1
                            If Len(acctS) > 0 Then
                                If Not walAccts.Exists(acctS) Then walAccts.Add acctS, True
                            End If
                        ElseIf InStr(1, delivS, "MINUTES") > 0 Or InStr(1, delivS, "MINUTOS") > 0 Then
                            cashTrx = cashTrx + 1   ' Money in Minutes = cash pickup
                        End If

                        ' (d2) dispersion / round-amount stats
                        If Not susGovs.Exists(gv) Then susGovs.Add gv, True
                        If amt >= 250000 And amt <> WU_LIMIT And amt < 2000000000 Then
                            If CLng(amt) Mod 250000 = 0 Then roundCnt = roundCnt + 1
                        End If

                        ' (d) declared purpose concentration (col 22)
                        purpS = Trim$(CStr(dataArr(r, 22)))
                        If Len(purpS) > 0 Then
                            If Not purpCnt.Exists(purpS) Then purpCnt.Add purpS, 0
                            purpCnt(purpS) = purpCnt(purpS) + 1
                        End If
                    End If
                End If
            End If
        Next r

        '----- network indicators (computed over UNIQUE flagged customers) -----
        uniqSus = sNat.count
        natIraq = 0: natDest = 0: natOther = 0: newWallets = 0
        For Each mKey In sNat.Keys
            If sNat(mKey) = "IRAQ" Then
                natIraq = natIraq + 1
            ElseIf sNat(mKey) = uCountry Then
                natDest = natDest + 1
            Else
                natOther = natOther + 1
            End If
            kycV = sKyc(mKey)
            sendV = sMinSend(mKey)
            If IsNumeric(kycV) And IsNumeric(sendV) Then
                If kycV > 0 And sendV > 0 Then
                    If sendV - kycV >= 0 And sendV - kycV <= 30 Then
                        newWallets = newWallets + 1
                        If Int(sendV) = Int(kycV) Then sameDayW = sameDayW + 1
                    End If
                End If
            End If
        Next mKey

        ' top recruiter by number of unique flagged customers
        Set recrCust = CreateObject("Scripting.Dictionary")
        topRecr = "": topRecrN = 0
        For Each mKey In sRecr.Keys
            recrS = sRecr(mKey)
            If Len(recrS) > 0 Then
                ' skip system/self-registration recruiter IDs (not real agents)
                If InStr(1, EXCLUDED_RECRUITERS, "," & recrS & ",") = 0 Then
                    If Not recrCust.Exists(recrS) Then recrCust.Add recrS, 0
                    recrCust(recrS) = recrCust(recrS) + 1
                End If
            End If
        Next mKey
        For Each k In recrCust.Keys
            If recrCust(k) > topRecrN Then
                topRecrN = recrCust(k)
                topRecr = CStr(k)
            End If
        Next k

        ' top receiver by number of unique senders (funnel indicator)
        recvDisplay = "": topRecvN = 0
        For Each k In recvSenders.Keys
            If recvSenders(k).count > topRecvN Then
                topRecvN = recvSenders(k).count
                recvDisplay = recvName(k)
            End If
        Next k


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
            tbl.Rows.Alignment = 1                       ' table centered on the page
            tbl.Range.Cells.VerticalAlignment = 1        ' cell content vertically centered
            tbl.Rows.AllowBreakAcrossPages = False       ' rows never split across pages
            tbl.Rows(1).HeadingFormat = True             ' header repeats on long tables

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

        '----- Key Observations (AUTO-GENERATED - review before sending) -----
        AddPara rng, "Key Observations:", True

        If susTrx = 0 Then
            AddPara rng, "No transactions were flagged as suspicious for this corridor during the review period.", , , , True
        Else
            ' (1) summary rule
            obsTxt = "Out of " & totTrx & " total transactions, " & susTrx & _
                     " are flagged as suspicious, accounting for a highly suspicious volume of " & _
                     Format$(susAmt, "#,##0") & " IQD"
            If totAmt > 0 Then
                obsTxt = obsTxt & " (" & Format$(susAmt / totAmt * 100, "0") & "% of the total amount sent)"
            End If
            AddPara rng, obsTxt & ".", , , , True

            ' (2) top governorates rule
            If nG >= 2 Then
                AddPara rng, "The majority of the overall transactions originate from " & gNames(1) & _
                             " (" & gT(1) & " transactions) and " & gNames(2) & " (" & gT(2) & _
                             " transactions).", , , , True
            ElseIf nG = 1 Then
                AddPara rng, "All transactions originate from " & gNames(1) & " (" & gT(1) & _
                             " transactions).", , , , True
            End If

            ' (3) max-limit pattern rule (only when the pattern exists)
            If susMaxCnt + susNearCnt > 0 Then
                exNearTxt = Join(exNear.Keys, ", ")
                If susMaxCnt > 0 Then
                    obsTxt = "The transfer amounts are consistently high, with " & susMaxCnt & _
                             " suspicious transactions deliberately reaching the maximum allowed transfer limit of exactly " & _
                             Format$(WU_LIMIT, "#,##0") & " IQD"
                    If susNearCnt > 0 Then
                        obsTxt = obsTxt & " and " & susNearCnt & " transactions at values immediately below it"
                        If Len(exNearTxt) > 0 Then obsTxt = obsTxt & " (e.g., " & exNearTxt & ")"
                    End If
                Else
                    obsTxt = "The transfer amounts are consistently high, with " & susNearCnt & _
                             " suspicious transactions at values immediately below the maximum allowed transfer limit of " & _
                             Format$(WU_LIMIT, "#,##0") & " IQD"
                    If Len(exNearTxt) > 0 Then obsTxt = obsTxt & " (e.g., " & exNearTxt & ")"
                End If
                AddPara rng, obsTxt & ".", , , , True
            End If

            ' (4) unique customers + nationality rule (UNIQUE customers, not trx)
            obsTxt = "The " & susTrx & " suspicious transactions were executed by " & uniqSus & _
                     " unique customers: " & natIraq & " Iraqi nationals"
            If natDest > 0 Then obsTxt = obsTxt & ", " & natDest & " " & StrConv(cName, vbProperCase) & " nationals"
            If natOther > 0 Then obsTxt = obsTxt & ", and " & natOther & " of other nationalities"
            AddPara rng, obsTxt & ".", , , , True

            ' (5) receiving channel rule (bank / wallet / cash pickup - only non-zero parts)
            If bankTrx + walTrx + cashTrx > 0 Then
                obsTxt = ""
                If bankTrx > 0 Then
                    obsTxt = bankAccts.count & " unique bank accounts (" & bankTrx & " transactions)"
                End If
                If walTrx > 0 Then
                    If Len(obsTxt) > 0 Then obsTxt = obsTxt & ", "
                    obsTxt = obsTxt & walAccts.count & " unique mobile wallets (" & walTrx & " transactions)"
                End If
                If cashTrx > 0 Then
                    If Len(obsTxt) > 0 Then obsTxt = obsTxt & ", "
                    obsTxt = obsTxt & cashTrx & " Money in Minutes (cash-pickup transactions)"
                End If
                AddPara rng, "The suspicious funds were received through " & obsTxt & ".", , , , True
            End If

            ' (6) declared purpose concentration (only when one purpose >= 60%)
            topPurp = "": topPurpN = 0
            For Each k In purpCnt.Keys
                If purpCnt(k) > topPurpN Then
                    topPurpN = purpCnt(k)
                    topPurp = CStr(k)
                End If
            Next k
            If topPurpN > 0 And susTrx > 0 Then
                If topPurpN / susTrx >= 0.6 Then
                    AddPara rng, "The declared purpose of transaction is heavily concentrated in '" & topPurp & _
                                 "' (" & topPurpN & " of " & susTrx & " suspicious transactions).", , , , True
                End If
            End If

            ' (7) new-wallet / mule onboarding indicator (shown when >=3 or >=20%)
            If newWallets >= 3 Or (uniqSus > 0 And newWallets / uniqSus >= 0.2) Then
                If newWallets > 0 Then
                    obsTxt = newWallets & " of the " & uniqSus & _
                             " flagged customers registered their wallets less than 30 days before their first suspicious transfer, a pattern consistent with mule-account onboarding"
                    If sameDayW > 0 Then
                        obsTxt = obsTxt & " - including " & sameDayW & _
                                 " wallets that executed their very first transfer on the same day they were registered"
                    End If
                    AddPara rng, obsTxt & ".", , , , True
                End If
            End If

            ' (8) recruiter concentration indicator (shown when one recruiter >=5 customers)
            If topRecrN >= 5 Then
                AddPara rng, topRecrN & " of the " & uniqSus & _
                             " flagged customers were registered under the same recruiter (Recruiter ID: " & topRecr & _
                             "); the involvement of this cash-in agent should be reviewed.", , , , True
            End If

            ' (9) common-receiver funnel indicator (shown when one receiver has >=3 senders)
            If topRecvN >= 3 Then
                AddPara rng, topRecvN & " unique senders transferred funds to the same receiver ('" & recvDisplay & _
                             "'), a strong indicator of a funnel network.", , , , True
            End If


            ' (11) round-amount indicator (>=50% in exact multiples of 250,000, excluding the limit itself)
            If susTrx >= 5 And roundCnt > 0 Then
                If roundCnt / susTrx >= 0.5 Then
                    AddPara rng, "Most transfer values are clean round figures (" & roundCnt & _
                                 " transfers in exact multiples of 250,000 IQD), a hallmark of instructed transfers rather than genuine personal remittances.", , , , True
                End If
            End If

            ' (12) geographic dispersion + identical behavior (coordination indicator)
            If susGovs.count >= 4 And susTrx >= 5 Then
                If (susMaxCnt + susNearCnt) / susTrx >= 0.5 Then
                    AddPara rng, "Although the flagged senders are dispersed across " & susGovs.count & _
                                 " different governorates, their transfer behavior is nearly identical in amounts and timing, indicating centrally coordinated activity rather than independent remittances.", , , , True
                End If
            End If

            ' (13) per-customer intensity vs corridor average (>=1.5x)
            If uniqSus > 0 And allSenders.count > 0 Then
                avgSus = susAmt / uniqSus
                avgAll = totAmt / allSenders.count
                If avgAll > 0 Then
                    If avgSus / avgAll >= 1.5 Then
                        AddPara rng, "On average, each flagged customer moved " & Format$(avgSus, "#,##0") & _
                                     " IQD within this single week - approximately " & Format$(avgSus / avgAll, "0.0") & _
                                     " times the volume of a typical sender in this corridor.", , , , True
                    End If
                End If
            End If
        End If

        '----- Conclusion (AUTO-DRAFT - review before sending) -----
        AddPara rng, "Conclusion:", True
        If susTrx = 0 Then
            conclTxt = "No suspicious activity was identified for this corridor during the investigation period, and no action is required at this stage."
        ElseIf totTrx > 0 And susTrx / totTrx >= 0.8 Then
            conclTxt = "The volume and structure of the flagged transfers strongly indicate an organized money-funneling network operating through this corridor. Services for the flagged customers should be restricted pending valid supporting evidence."
        Else
            conclTxt = "The pattern of structured, near-limit transfers executed by the flagged senders is consistent with known money-funneling typologies. Services for the flagged customers should be restricted pending valid supporting evidence."
        End If
        AddPara rng, conclTxt
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
           vbCrLf & vbCrLf & "Countries: " & nc & vbCrLf & _
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
