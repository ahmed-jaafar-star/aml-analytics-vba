Attribute VB_Name = "WUReportAnalyzer"
' ============================================================================
'  WU REPORT ANALYZER - main module
'  Input : active sheet = WU transactions export (columns A:AB, SEND_MONEY rows)
'          Sheet2       = cash-in data (optional, for the agent analysis)
'  Output: one sheet per target country ("Analysis - <COUNTRY>") or
'          "All Countries", with the risk tables described in README.md
'  Needs : UserForm "frmSettings" and class "clsBtn" in the same project
' ============================================================================
Sub WUReportAnalysis_Final_Version()
    Dim wsSource As Worksheet, wsDest As Worksheet, wsCashIn As Worksheet
    Dim lastRow As Long, lastRow2 As Long, i As Long, cIdx As Long, mIdx As Long
    Dim dataArr As Variant, itemArr As Variant, cArr As Variant, gArr As Variant, sArr As Variant
    Dim acct As String, sender As String, amt As Double, isRedSus As Boolean
    Dim ctry As String, gov As String, del As String, trxType As String, modelID As String
    Dim uniqueAcctKey As String, kVar As Variant, subK As Variant
    Dim recName As String, mtcn As String, createdAt As Variant, nat As String
    Dim r5 As Long, tRow As Long, startR As Long
    Dim currentModel As String, iterModel As String, chkSender As String

    Dim combKey As String, cTrx As Long, trxData As String, sKey As String
    Dim sp() As String, trxPs() As String, d() As String
    Dim cl As Range

    Dim redBank As Long, redMobile As Long
    Dim yelBank As Long, yelMobile As Long

    Dim r As Long, cr As Long, natRow As Long, r4 As Long
    Dim redS As Long, yelS As Long, hRS As Long, gRS As Long, nStart As Long, hG As Long, gg As Long
    Dim gCtryStartRow As Long
    Dim nArrGlobal As Variant
    Dim recMeta As String, nmKey As String

    ' --- ADJUST THESE COLUMN NUMBERS IF NEEDED ---
    Dim colMTCN As Long: colMTCN = 17   ' MTCN is in column Q (Actor ID is column A = 1)
    Dim colNat As Long: colNat = 5
    Dim colRecName As Long: colRecName = 15
    Dim colCreatedAt As Long: colCreatedAt = 20
    ' ---------------------------------------------

    Set wsSource = ActiveSheet
    lastRow = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ' --- DETECT RECEIVER-NAME & MOTHER-NAME SOURCE COLUMNS (by header) ---
    Dim recNameSrcCol As Long, motherSrcCol As Long, hh As String
    recNameSrcCol = 0: motherSrcCol = 0
    For cIdx = 1 To 28
        hh = UCase(Trim(CStr(wsSource.Cells(1, cIdx).Value)))
        If recNameSrcCol = 0 And (hh = "RECEIVER_NAME_CONCATENATED_Q1_2" Or InStr(hh, "RECEIVER_NAME") > 0) Then recNameSrcCol = cIdx
        If motherSrcCol = 0 And InStr(hh, "MOTHER") > 0 Then motherSrcCol = cIdx
    Next cIdx
    If recNameSrcCol = 0 Then recNameSrcCol = colRecName

    ' ========================================================
    ' SETTINGS - single pop-up form
    ' ========================================================
    Dim f As frmSettings
    Set f = New frmSettings
    f.Show vbModal
    Dim wasCancel As Boolean: wasCancel = f.Cancelled
    Dim userInput As String: userInput = f.countries
    Dim useUnique As Boolean: useUnique = f.useUnique
    Dim redThresh As Long: redThresh = f.redThresh
    Dim useAgent As Boolean: useAgent = f.useAgent
    Dim agtUseUnique As Boolean: agtUseUnique = f.agtUseUnique
    Dim agtThresh As Long: agtThresh = f.agtThresh
    Dim useMeta As Boolean: useMeta = f.UseMetaphone
    Dim noteAllTrx As Boolean: noteAllTrx = f.noteAllTrx
    Unload f
    Set f = Nothing
    If wasCancel Then Exit Sub
    If Trim(userInput) = "" Then Exit Sub
    If redThresh < 2 Then redThresh = 2
    If agtThresh < 1 Then agtThresh = 1

    Dim targetArr() As String
    targetArr = Split(userInput, ",")

    ' --- AGENT SHEET CHECK ---
    Dim hasCashInSheet As Boolean
    hasCashInSheet = False
    If useAgent Then
        On Error Resume Next
        Set wsCashIn = wsSource.Parent.Sheets("Sheet2")
        On Error GoTo 0
        hasCashInSheet = Not (wsCashIn Is Nothing)
        If Not hasCashInSheet Then
            MsgBox "'Sheet2' (Cash-In data) was not found!" & vbCrLf & "The Agent Risk analysis will be skipped.", vbExclamation, "Missing Data"
            useAgent = False
        End If
    End If

    ' ========================================================
    ' RAW-DATA COLUMN PLAN (computed once)
    '   Injections: Name Count (after receiver name), Account Count (after 26),
    '   Model ID Count (after 28), Mother Name Count (after mother col).
    '   + 3 agent columns at the end if useAgent.
    ' ========================================================
    Dim outCol(1 To 28) As Long
    Dim posNameCount As Long, posAcctCount As Long, posModelCount As Long, posMotherCount As Long
    Dim posAgent1 As Long, posAgent2 As Long, posAgent3 As Long
    posNameCount = 0: posAcctCount = 0: posModelCount = 0: posMotherCount = 0
    Dim nc As Long: nc = 1
    For cIdx = 1 To 28
        outCol(cIdx) = nc: nc = nc + 1
        If cIdx = recNameSrcCol Then posNameCount = nc: nc = nc + 1
        If cIdx = 26 Then posAcctCount = nc: nc = nc + 1
        If cIdx = 28 Then posModelCount = nc: nc = nc + 1
        If motherSrcCol > 0 And cIdx = motherSrcCol Then posMotherCount = nc: nc = nc + 1
    Next cIdx
    Dim baseCols As Long: baseCols = nc - 1
    Dim rawCols As Long: rawCols = baseCols
    If useAgent Then
        posAgent1 = baseCols + 1: posAgent2 = baseCols + 2: posAgent3 = baseCols + 3
        rawCols = baseCols + 3
    End If

    Application.ScreenUpdating = False
    dataArr = wsSource.Range("A1:AB" & lastRow).Value

    ' --- METAPHONE CACHE + WRITE TO SOURCE COLUMN AC (29) ---
    Dim dictNameMeta As Object: Set dictNameMeta = CreateObject("Scripting.Dictionary")
    If useMeta Then
        wsSource.Cells(1, 29).Value = "RECEIVER_METAPHONE"
        Dim metaOut() As Variant: ReDim metaOut(1 To lastRow - 1, 1 To 1)
        For i = 2 To lastRow
            nmKey = UCase(Trim(CStr(dataArr(i, recNameSrcCol))))
            If dictNameMeta.Exists(nmKey) Then
                recMeta = dictNameMeta(nmKey)
            Else
                recMeta = Metaphone(nmKey): dictNameMeta.Add nmKey, recMeta
            End If
            metaOut(i - 1, 1) = recMeta
        Next i
        wsSource.Cells(2, 29).Resize(lastRow - 1, 1).Value = metaOut
        wsSource.Cells(1, 29).Font.Bold = True
    End If

    ' --- GLOBAL DATE SCANNER ---
    Dim globalMin As Date, globalMax As Date
    Dim globalDateFound As Boolean
    globalDateFound = False
    For i = 2 To lastRow
        Dim gdt As Variant
        gdt = ParseDate(dataArr(i, colCreatedAt))
        If IsDate(gdt) Then
            If Not globalDateFound Then
                globalMin = gdt: globalMax = gdt: globalDateFound = True
            Else
                If gdt < globalMin Then globalMin = gdt
                If gdt > globalMax Then globalMax = gdt
            End If
        End If
    Next i

    ' --- MASTER LOOP ---
    Dim tIdx As Long
    For tIdx = LBound(targetArr) To UBound(targetArr)

        Dim currentTarget As String
        currentTarget = UCase(Trim(targetArr(tIdx)))
        If currentTarget = "" Then GoTo NextTarget

        Dim isGlobal As Boolean
        isGlobal = (currentTarget = "ALL COUNTRIES")

        ' RESET DICTIONARIES
        Dim dictAcctCtryCount As Object: Set dictAcctCtryCount = CreateObject("Scripting.Dictionary")
        Dim dictPrimaryCtry As Object: Set dictPrimaryCtry = CreateObject("Scripting.Dictionary")
        Dim dictMaxCtry As Object: Set dictMaxCtry = CreateObject("Scripting.Dictionary")
        Dim dictAcct As Object: Set dictAcct = CreateObject("Scripting.Dictionary")
        Dim dictSenders As Object: Set dictSenders = CreateObject("Scripting.Dictionary")
        Dim dictRedAccts As Object: Set dictRedAccts = CreateObject("Scripting.Dictionary")
        Dim dictYellowAccts As Object: Set dictYellowAccts = CreateObject("Scripting.Dictionary")
        Dim dictCountry As Object: Set dictCountry = CreateObject("Scripting.Dictionary")
        Dim dictCountrySenders As Object: Set dictCountrySenders = CreateObject("Scripting.Dictionary")
        Dim dictCountrySusSenders As Object: Set dictCountrySusSenders = CreateObject("Scripting.Dictionary")
        Dim dictGov As Object: Set dictGov = CreateObject("Scripting.Dictionary")
        Dim dictGovSenders As Object: Set dictGovSenders = CreateObject("Scripting.Dictionary")
        Dim dictGovSusSenders As Object: Set dictGovSusSenders = CreateObject("Scripting.Dictionary")
        Dim dictSenderAnalysis As Object: Set dictSenderAnalysis = CreateObject("Scripting.Dictionary")
        Dim dictSenderActor As Object: Set dictSenderActor = CreateObject("Scripting.Dictionary")
        Dim dictNoteAll As Object: Set dictNoteAll = CreateObject("Scripting.Dictionary")
        Dim dictNoteSus As Object: Set dictNoteSus = CreateObject("Scripting.Dictionary")
        Dim dictModelCount As Object: Set dictModelCount = CreateObject("Scripting.Dictionary")
        Dim dictModelAmt As Object: Set dictModelAmt = CreateObject("Scripting.Dictionary")
        Dim dictMetaCount As Object: Set dictMetaCount = CreateObject("Scripting.Dictionary")
        Dim dictMetaAmt As Object: Set dictMetaAmt = CreateObject("Scripting.Dictionary")
        Dim dictSusSendersCheck As Object: Set dictSusSendersCheck = CreateObject("Scripting.Dictionary")
        Dim dictRedCtry As Object: Set dictRedCtry = CreateObject("Scripting.Dictionary")
        Dim dictYelCtry As Object: Set dictYelCtry = CreateObject("Scripting.Dictionary")
        Dim dictNatGlobal As Object: Set dictNatGlobal = CreateObject("Scripting.Dictionary")
        Dim dictNatSendersGlobal As Object: Set dictNatSendersGlobal = CreateObject("Scripting.Dictionary")
        Dim dictRecNameCtry As Object: Set dictRecNameCtry = CreateObject("Scripting.Dictionary")
        Dim dictMotherCtry As Object: Set dictMotherCtry = CreateObject("Scripting.Dictionary")

        ' AGENT TRACKERS
        Dim dictCustGov As Object: Set dictCustGov = CreateObject("Scripting.Dictionary")
        Dim dictSusSenderData As Object: Set dictSusSenderData = CreateObject("Scripting.Dictionary")
        Dim dictSenderCtryDict As Object: Set dictSenderCtryDict = CreateObject("Scripting.Dictionary")
        Dim dictFlaggedAgents As Object: Set dictFlaggedAgents = CreateObject("Scripting.Dictionary")
        Dim dictFlaggedAgentStrings As Object: Set dictFlaggedAgentStrings = CreateObject("Scripting.Dictionary")
        Dim dictFlaggedAgentMetrics As Object: Set dictFlaggedAgentMetrics = CreateObject("Scripting.Dictionary")
        Dim dictCustToAllAgents As Object: Set dictCustToAllAgents = CreateObject("Scripting.Dictionary")

        redBank = 0: redMobile = 0: yelBank = 0: yelMobile = 0

        ' --- PASS 0: Primary Country ---
        For i = 2 To lastRow
            trxType = UCase(Replace(Trim(CStr(dataArr(i, 13))), " ", "_"))
            If trxType = "SEND_MONEY" Or trxType = "SENDMONEY" Then
                acct = CleanMSISDN(CStr(dataArr(i, 26)))
                ctry = UCase(Trim(CStr(dataArr(i, 16))))
                If ctry = "" Then ctry = "UNKNOWN"
                If acct <> "" Then
                    combKey = acct & "|" & ctry
                    If Not dictAcctCtryCount.Exists(combKey) Then dictAcctCtryCount.Add combKey, 1 Else dictAcctCtryCount(combKey) = dictAcctCtryCount(combKey) + 1
                End If
            End If
        Next i

        For Each kVar In dictAcctCtryCount.Keys
            acct = Split(kVar, "|")(0): ctry = Split(kVar, "|")(1): cTrx = dictAcctCtryCount(kVar)
            If Not dictPrimaryCtry.Exists(acct) Then
                dictPrimaryCtry.Add acct, ctry: dictMaxCtry.Add acct, cTrx
            Else
                If cTrx > dictMaxCtry(acct) Then dictPrimaryCtry(acct) = ctry: dictMaxCtry(acct) = cTrx
            End If
        Next kVar

        ' --- PASS 1: Categorize Accounts (+ count Model IDs and Metaphones) ---
        For i = 2 To lastRow
            trxType = UCase(Replace(Trim(CStr(dataArr(i, 13))), " ", "_"))
            If trxType = "SEND_MONEY" Or trxType = "SENDMONEY" Then
                acct = CleanMSISDN(CStr(dataArr(i, 26)))
                If acct <> "" Then ctry = dictPrimaryCtry(acct) Else ctry = UCase(Trim(CStr(dataArr(i, 16))))
                If ctry = "" Then ctry = "UNKNOWN"

                If isGlobal Or ctry = currentTarget Then
                    modelID = Trim(CStr(dataArr(i, 28)))
                    If modelID <> "" Then dictModelCount(modelID) = dictModelCount(modelID) + 1

                    If useMeta Then
                        nmKey = UCase(Trim(CStr(dataArr(i, recNameSrcCol))))
                        If dictNameMeta.Exists(nmKey) Then recMeta = dictNameMeta(nmKey) Else recMeta = Metaphone(nmKey): dictNameMeta.Add nmKey, recMeta
                        If recMeta <> "" Then dictMetaCount(recMeta) = dictMetaCount(recMeta) + 1
                    End If

                    If acct <> "" Then
                        sender = CleanMSISDN(CStr(dataArr(i, 12)))
                        uniqueAcctKey = acct
                        If Not dictAcct.Exists(uniqueAcctKey) Then dictAcct.Add uniqueAcctKey, Array(0, 0, CDbl(0), ctry, Trim(CStr(dataArr(i, 27))), UCase(Trim(CStr(dataArr(i, 25)))), acct)
                        itemArr = dictAcct(uniqueAcctKey)
                        itemArr(0) = itemArr(0) + 1
                        itemArr(2) = itemArr(2) + Val(dataArr(i, 14))
                        If Not dictSenders.Exists(uniqueAcctKey & "||" & sender) Then
                            dictSenders.Add uniqueAcctKey & "||" & sender, True
                            itemArr(1) = itemArr(1) + 1
                        End If
                        dictAcct(uniqueAcctKey) = itemArr
                    End If
                End If
            End If
        Next i

        Dim evalMetric As Long
        For Each kVar In dictAcct.Keys
            If useUnique Then evalMetric = dictAcct(kVar)(1) Else evalMetric = dictAcct(kVar)(0)
            If evalMetric >= redThresh Then
                dictRedAccts.Add kVar, True
            ElseIf evalMetric >= 3 Then
                dictYellowAccts.Add kVar, True
            End If
        Next kVar

        ' --- PASS 2: Aggregate Tables & Capture Raw Data ---
        Dim maxTrx As Long: maxTrx = 1
        Dim t5Row As Long: t5Row = 0
        Dim t7Row As Long: t7Row = 0
        Dim arrT5() As Variant: ReDim arrT5(1 To lastRow, 1 To 12)
        Dim arrT7() As Variant: ReDim arrT7(1 To lastRow, 1 To 12)

        Dim rawArr() As Variant
        Dim rawCount As Long
        If Not isGlobal Then
            ReDim rawArr(1 To lastRow, 1 To rawCols)
            rawCount = 0
        End If

        For i = 2 To lastRow
            trxType = UCase(Replace(Trim(CStr(dataArr(i, 13))), " ", "_"))
            If trxType = "SEND_MONEY" Or trxType = "SENDMONEY" Then
                acct = CleanMSISDN(CStr(dataArr(i, 26)))
                If acct <> "" Then ctry = dictPrimaryCtry(acct) Else ctry = UCase(Trim(CStr(dataArr(i, 16))))
                If ctry = "" Then ctry = "UNKNOWN"

                ' ---- RAW CAPTURE (country only) ----
                If Not isGlobal And ctry = currentTarget Then
                    rawCount = rawCount + 1
                    For cIdx = 1 To 28
                        Select Case cIdx
                            Case 12
                                rawArr(rawCount, outCol(cIdx)) = NumIf(CStr(dataArr(i, cIdx)))   ' MSISDN as number (no E+, no triangle)
                            Case 1, colMTCN, 26, 28
                                rawArr(rawCount, outCol(cIdx)) = CStr(dataArr(i, cIdx))           ' actor id, MTCN, account, model as text
                            Case Else
                                rawArr(rawCount, outCol(cIdx)) = dataArr(i, cIdx)
                        End Select
                    Next cIdx
                    ' Account count
                    Dim tmpAcct As String: tmpAcct = CleanMSISDN(CStr(dataArr(i, 26)))
                    If posAcctCount > 0 Then If dictAcct.Exists(tmpAcct) Then rawArr(rawCount, posAcctCount) = dictAcct(tmpAcct)(0) Else rawArr(rawCount, posAcctCount) = 0
                    ' Model id count
                    Dim tmpModel As String: tmpModel = Trim(CStr(dataArr(i, 28)))
                    If posModelCount > 0 Then If dictModelCount.Exists(tmpModel) Then rawArr(rawCount, posModelCount) = dictModelCount(tmpModel) Else rawArr(rawCount, posModelCount) = 0

                    ' tally receiver-name & mother-name counts for this country
                    Dim rnV As String: rnV = UCase(Trim(CStr(dataArr(i, recNameSrcCol))))
                    If Not dictRecNameCtry.Exists(rnV) Then dictRecNameCtry.Add rnV, 0
                    dictRecNameCtry(rnV) = dictRecNameCtry(rnV) + 1
                    If motherSrcCol > 0 Then
                        Dim moV As String: moV = UCase(Trim(CStr(dataArr(i, motherSrcCol))))
                        If Not dictMotherCtry.Exists(moV) Then dictMotherCtry.Add moV, 0
                        dictMotherCtry(moV) = dictMotherCtry(moV) + 1
                    End If
                End If

                ' ---- AGGREGATIONS ----
                If isGlobal Or ctry = currentTarget Then
                    sender = CleanMSISDN(CStr(dataArr(i, 12)))
                    amt = Val(dataArr(i, 14))
                    gov = Trim(CStr(dataArr(i, 10)))
                    del = UCase(Trim(CStr(dataArr(i, 25))))
                    modelID = Trim(CStr(dataArr(i, 28)))
                    mtcn = Trim(CStr(dataArr(i, colMTCN)))
                    recName = Trim(CStr(dataArr(i, recNameSrcCol)))
                    createdAt = ParseDate(dataArr(i, colCreatedAt))
                    nat = UCase(Trim(CStr(dataArr(i, colNat))))
                    If nat = "" Then nat = "UNKNOWN"
                    If gov = "" Then gov = "UNKNOWN"

                    recMeta = ""
                    If useMeta Then
                        nmKey = UCase(Trim(CStr(dataArr(i, recNameSrcCol))))
                        If dictNameMeta.Exists(nmKey) Then recMeta = dictNameMeta(nmKey) Else recMeta = Metaphone(nmKey): dictNameMeta.Add nmKey, recMeta
                    End If

                    If Not dictCustGov.Exists(sender) Then dictCustGov.Add sender, UCase(gov)

                    If Not dictNatGlobal.Exists(nat) Then dictNatGlobal.Add nat, Array(0, 0)
                    nArrGlobal = dictNatGlobal(nat)
                    nArrGlobal(0) = nArrGlobal(0) + 1
                    If Not dictNatSendersGlobal.Exists(nat & "|" & sender) Then
                        dictNatSendersGlobal.Add nat & "|" & sender, True
                        nArrGlobal(1) = nArrGlobal(1) + 1
                    End If
                    dictNatGlobal(nat) = nArrGlobal

                    If Not dictCountry.Exists(ctry) Then dictCountry.Add ctry, Array(0, 0, 0, 0, 0, CDbl(0), 0, 0, 0, CDbl(0))
                    cArr = dictCountry(ctry): cArr(0) = cArr(0) + 1: cArr(5) = cArr(5) + amt
                    If del = "DIRECT TO BANK" Then cArr(2) = cArr(2) + 1
                    If del = "MOBILE MONEY TRANSFER" Then cArr(3) = cArr(3) + 1
                    If del = "MONEY IN MINUTES" Then cArr(4) = cArr(4) + 1
                    If Not dictCountrySenders.Exists(ctry & "|" & sender) Then dictCountrySenders.Add ctry & "|" & sender, True: cArr(1) = cArr(1) + 1

                    If Not dictGov.Exists(gov) Then dictGov.Add gov, Array(0, 0, CDbl(0), 0, 0, CDbl(0))
                    gArr = dictGov(gov): gArr(0) = gArr(0) + 1: gArr(2) = gArr(2) + amt
                    If Not dictGovSenders.Exists(gov & "|" & sender) Then dictGovSenders.Add gov & "|" & sender, True: gArr(1) = gArr(1) + 1

                    ' --- NOTE fragment (Country - Receiver - Amount - Date) ---
                    Dim noteDate As String
                    If IsDate(createdAt) Then noteDate = Format(createdAt, "yyyy-mm-dd") Else noteDate = "N/A"
                    Dim noteFrag As String
                    noteFrag = ctry & " - " & recName & " - " & Format(amt, "#,##0") & " - " & noteDate
                    Dim nKey As String: nKey = sender & "|" & gov
                    If Not dictNoteAll.Exists(nKey) Then dictNoteAll.Add nKey, ""
                    If dictNoteAll(nKey) = "" Then dictNoteAll(nKey) = noteFrag Else dictNoteAll(nKey) = dictNoteAll(nKey) & vbLf & noteFrag

                    isRedSus = dictRedAccts.Exists(acct)
                    If isRedSus Then
                        Dim sKeyCtry As String: sKeyCtry = sender & "|" & ctry
                        If Not dictSusSenderData.Exists(sKeyCtry) Then dictSusSenderData.Add sKeyCtry, 0
                        dictSusSenderData(sKeyCtry) = dictSusSenderData(sKeyCtry) + 1
                        If Not dictSenderCtryDict.Exists(sender) Then Set dictSenderCtryDict(sender) = CreateObject("Scripting.Dictionary")
                        If Not dictSenderCtryDict(sender).Exists(ctry) Then dictSenderCtryDict(sender).Add ctry, True

                        cArr(7) = cArr(7) + 1: cArr(9) = cArr(9) + amt
                        gArr(3) = gArr(3) + 1: gArr(5) = gArr(5) + amt
                        If Not dictSusSendersCheck.Exists(sender) Then dictSusSendersCheck.Add sender, True
                        If Not dictCountrySusSenders.Exists(ctry & "|" & sender) Then dictCountrySusSenders.Add ctry & "|" & sender, True: cArr(8) = cArr(8) + 1
                        If Not dictGovSusSenders.Exists(gov & "|" & sender) Then dictGovSusSenders.Add gov & "|" & sender, True: gArr(4) = gArr(4) + 1

                        Dim cSer As String
                        If IsDate(createdAt) Then cSer = Trim(Str(CDbl(createdAt))) Else cSer = ""
                        trxData = CStr(amt) & "^" & acct & "^" & ctry & "^" & recName & "^" & mtcn & "^" & cSer
                        sKey = sender & "|" & gov
                        If Not dictSenderActor.Exists(sKey) Then dictSenderActor.Add sKey, Trim(CStr(dataArr(i, 1)))
                        If Not dictNoteSus.Exists(sKey) Then dictNoteSus.Add sKey, ""
                        If dictNoteSus(sKey) = "" Then dictNoteSus(sKey) = noteFrag Else dictNoteSus(sKey) = dictNoteSus(sKey) & vbLf & noteFrag
                        If Not dictSenderAnalysis.Exists(sKey) Then
                            dictSenderAnalysis.Add sKey, Array(1, trxData, ctry)
                        Else
                            sArr = dictSenderAnalysis(sKey): sArr(0) = sArr(0) + 1: sArr(1) = sArr(1) & "|" & trxData: dictSenderAnalysis(sKey) = sArr
                        End If
                        If dictSenderAnalysis(sKey)(0) > maxTrx Then maxTrx = dictSenderAnalysis(sKey)(0)
                    End If

                    dictCountry(ctry) = cArr
                    dictGov(gov) = gArr

                    ' Many-to-one by MODEL ID
                    If modelID <> "" And dictModelCount(modelID) >= 5 Then
                        dictModelAmt(modelID) = dictModelAmt(modelID) + amt
                        t5Row = t5Row + 1
                        arrT5(t5Row, 1) = NumIf(sender): arrT5(t5Row, 2) = gov: arrT5(t5Row, 3) = recName: arrT5(t5Row, 4) = amt
                        arrT5(t5Row, 5) = ctry: arrT5(t5Row, 6) = modelID: arrT5(t5Row, 7) = dictModelCount(modelID): arrT5(t5Row, 8) = del
                        arrT5(t5Row, 9) = acct: arrT5(t5Row, 10) = mtcn: arrT5(t5Row, 11) = createdAt
                    End If

                    ' Many-to-one by NAME SOUND (metaphone)
                    If useMeta And recMeta <> "" And dictMetaCount(recMeta) >= 5 Then
                        dictMetaAmt(recMeta) = dictMetaAmt(recMeta) + amt
                        t7Row = t7Row + 1
                        arrT7(t7Row, 1) = NumIf(sender): arrT7(t7Row, 2) = gov: arrT7(t7Row, 3) = recName: arrT7(t7Row, 4) = amt
                        arrT7(t7Row, 5) = ctry: arrT7(t7Row, 6) = recMeta: arrT7(t7Row, 7) = dictMetaCount(recMeta): arrT7(t7Row, 8) = del
                        arrT7(t7Row, 9) = acct: arrT7(t7Row, 10) = mtcn: arrT7(t7Row, 11) = createdAt
                    End If
                End If
            End If
        Next i

        ' --- PASS 3: AGENT LOGIC ---
        Dim dictAgtCustDict As Object: Set dictAgtCustDict = CreateObject("Scripting.Dictionary")
        If useAgent Then
            lastRow2 = wsCashIn.Cells(wsCashIn.Rows.count, "A").End(xlUp).Row
            If lastRow2 > 1 Then
                Dim data2 As Variant
                data2 = wsCashIn.Range("A1:Q" & lastRow2).Value
                Dim dictAgtTrx As Object: Set dictAgtTrx = CreateObject("Scripting.Dictionary")
                Dim dictAgtAmt As Object: Set dictAgtAmt = CreateObject("Scripting.Dictionary")
                Dim dictAgtCity As Object: Set dictAgtCity = CreateObject("Scripting.Dictionary")
                Dim cust As Variant, agt As Variant
                For i = 2 To lastRow2
                    cust = CleanMSISDN(CStr(data2(i, 5)))
                    agt = CleanMSISDN(CStr(data2(i, 10)))
                    If cust <> "" And agt <> "" Then
                        If Not dictCustToAllAgents.Exists(cust) Then
                            dictCustToAllAgents.Add cust, agt
                        Else
                            If InStr(dictCustToAllAgents(cust), agt) = 0 Then dictCustToAllAgents(cust) = dictCustToAllAgents(cust) & ", " & agt
                        End If
                        If dictSusSendersCheck.Exists(cust) Then
                            If Not dictAgtCustDict.Exists(agt) Then Set dictAgtCustDict(agt) = CreateObject("Scripting.Dictionary")
                            If Not dictAgtCustDict(agt).Exists(cust) Then dictAgtCustDict(agt).Add cust, True
                            If Not dictAgtTrx.Exists(agt) Then dictAgtTrx.Add agt, 0
                            dictAgtTrx(agt) = dictAgtTrx(agt) + 1
                            If Not dictAgtAmt.Exists(agt) Then dictAgtAmt.Add agt, CDbl(0)
                            dictAgtAmt(agt) = dictAgtAmt(agt) + Val(data2(i, 7))
                            If Not dictAgtCity.Exists(agt) Then dictAgtCity.Add agt, UCase(Trim(CStr(data2(i, 17))))
                        End If
                    End If
                Next i
                Dim evalAgt As Long
                Dim ctryStats As Object, arr As Variant, strB As String
                For Each agt In dictAgtCustDict.Keys
                    If agtUseUnique Then evalAgt = dictAgtCustDict(agt).count Else evalAgt = dictAgtTrx(agt)
                    If evalAgt >= agtThresh Then
                        dictFlaggedAgents.Add agt, evalAgt
                        Dim aCity As String: aCity = dictAgtCity(agt)
                        Dim sameC As Double: sameC = 0
                        Dim otherC As Double: otherC = 0
                        For Each cust In dictAgtCustDict(agt).Keys
                            If dictCustGov.Exists(cust) Then
                                Dim cGov As String: cGov = dictCustGov(cust)
                                If cGov <> "" And aCity <> "" And (cGov = aCity Or InStr(1, cGov, aCity, vbTextCompare) > 0 Or InStr(1, aCity, cGov, vbTextCompare) > 0) Then
                                    sameC = sameC + 1
                                Else
                                    otherC = otherC + 1
                                End If
                            Else
                                otherC = otherC + 1
                            End If
                        Next cust
                        Dim totCust As Double: totCust = dictAgtCustDict(agt).count
                        Dim pSame As Double: pSame = 0: Dim pOther As Double: pOther = 0
                        If totCust > 0 Then pSame = sameC / totCust: pOther = otherC / totCust
                        dictFlaggedAgentMetrics.Add agt, Array(dictAgtAmt(agt), aCity, pSame, pOther)
                        Set ctryStats = CreateObject("Scripting.Dictionary")
                        For Each cust In dictAgtCustDict(agt).Keys
                            If dictSenderCtryDict.Exists(cust) Then
                                For Each kVar In dictSenderCtryDict(cust).Keys
                                    Dim tTrx As Long: tTrx = dictSusSenderData(cust & "|" & kVar)
                                    If Not ctryStats.Exists(kVar) Then ctryStats.Add kVar, Array(0, 0)
                                    arr = ctryStats(kVar): arr(0) = arr(0) + tTrx: arr(1) = arr(1) + 1: ctryStats(kVar) = arr
                                Next kVar
                            End If
                        Next cust
                        strB = ""
                        For Each kVar In ctryStats.Keys
                            strB = strB & kVar & " (" & ctryStats(kVar)(0) & " Trx / " & ctryStats(kVar)(1) & " Senders), "
                        Next kVar
                        If Len(strB) > 0 Then strB = Left(strB, Len(strB) - 2)
                        dictFlaggedAgentStrings.Add agt, strB
                    End If
                Next agt
            End If
        End If

        ' --- PASS 4: AGENT INJECT INTO RAW ---
        If useAgent And Not isGlobal And rawCount > 0 Then
            Dim cleanSender As String
            Dim dictLocalAgtTrx As Object: Set dictLocalAgtTrx = CreateObject("Scripting.Dictionary")
            Dim dictLocalAgtCust As Object: Set dictLocalAgtCust = CreateObject("Scripting.Dictionary")
            For i = 1 To rawCount
                cleanSender = CleanMSISDN(CStr(rawArr(i, outCol(12))))
                If dictCustToAllAgents.Exists(cleanSender) Then
                    Dim aList() As String: aList = Split(dictCustToAllAgents(cleanSender), ", ")
                    Dim kA As Variant
                    For Each kA In aList
                        If Not dictLocalAgtTrx.Exists(CStr(kA)) Then dictLocalAgtTrx.Add CStr(kA), 0
                        dictLocalAgtTrx(CStr(kA)) = dictLocalAgtTrx(CStr(kA)) + 1
                        If Not dictLocalAgtCust.Exists(CStr(kA)) Then Set dictLocalAgtCust(CStr(kA)) = CreateObject("Scripting.Dictionary")
                        If Not dictLocalAgtCust(CStr(kA)).Exists(cleanSender) Then dictLocalAgtCust(CStr(kA)).Add cleanSender, True
                    Next kA
                End If
            Next i
            For i = 1 To rawCount
                cleanSender = CleanMSISDN(CStr(rawArr(i, outCol(12))))
                If dictCustToAllAgents.Exists(cleanSender) Then
                    Dim allAgts As String: allAgts = dictCustToAllAgents(cleanSender)
                    rawArr(i, posAgent1) = allAgts
                    Dim agtList() As String, strLocal As String, strGlobal As String, kAgt As Variant
                    agtList = Split(allAgts, ", ")
                    strLocal = "": strGlobal = ""
                    For Each kAgt In agtList
                        strLocal = strLocal & kAgt & " (" & dictLocalAgtTrx(CStr(kAgt)) & " Trx / " & dictLocalAgtCust(CStr(kAgt)).count & " Senders) | "
                        If dictFlaggedAgentStrings.Exists(CStr(kAgt)) Then strGlobal = strGlobal & kAgt & ": " & dictFlaggedAgentStrings(CStr(kAgt)) & " | "
                    Next kAgt
                    If Len(strLocal) > 0 Then rawArr(i, posAgent2) = Left(strLocal, Len(strLocal) - 3) Else rawArr(i, posAgent2) = ""
                    If Len(strGlobal) > 0 Then rawArr(i, posAgent3) = Left(strGlobal, Len(strGlobal) - 3) Else rawArr(i, posAgent3) = ""
                Else
                    rawArr(i, posAgent1) = "": rawArr(i, posAgent2) = "": rawArr(i, posAgent3) = ""
                End If
            Next i
        End If

        ' --- FILL NAME / MOTHER COUNT COLUMNS ---
        If Not isGlobal And rawCount > 0 Then
            For i = 1 To rawCount
                If posNameCount > 0 Then
                    rnV = UCase(Trim(CStr(rawArr(i, outCol(recNameSrcCol)))))
                    If dictRecNameCtry.Exists(rnV) Then rawArr(i, posNameCount) = dictRecNameCtry(rnV) Else rawArr(i, posNameCount) = 0
                End If
                If posMotherCount > 0 And motherSrcCol > 0 Then
                    moV = UCase(Trim(CStr(rawArr(i, outCol(motherSrcCol)))))
                    If dictMotherCtry.Exists(moV) Then rawArr(i, posMotherCount) = dictMotherCtry(moV) Else rawArr(i, posMotherCount) = 0
                End If
            Next i
        End If

        For Each subK In dictRedAccts.Keys
            ctry = dictAcct(subK)(3)
            If isGlobal Or ctry = currentTarget Then
                If dictCountry.Exists(ctry) Then cArr = dictCountry(ctry): cArr(6) = cArr(6) + 1: dictCountry(ctry) = cArr
            End If
        Next subK

        If t5Row > 0 Then For r5 = 1 To t5Row: arrT5(r5, 12) = dictModelAmt(Trim(CStr(arrT5(r5, 6)))): Next r5
        If t7Row > 0 Then For r5 = 1 To t7Row: arrT7(r5, 12) = dictMetaAmt(Trim(CStr(arrT7(r5, 6)))): Next r5

        ' --- DYNAMIC COLUMNS ---
        Dim hasHRCtry As Boolean: hasHRCtry = False
        Dim hasGenCtry As Boolean: hasGenCtry = False
        For Each kVar In dictCountry.Keys
            If dictCountry(kVar)(7) > 0 Then hasHRCtry = True
            If dictCountry(kVar)(7) = 0 Then hasGenCtry = True
        Next kVar
        Dim hasHRGov As Boolean: hasHRGov = False
        Dim hasGenGov As Boolean: hasGenGov = False
        For Each kVar In dictGov.Keys
            If dictGov(kVar)(3) > 0 Then hasHRGov = True
            If dictGov(kVar)(3) = 0 Then hasGenGov = True
        Next kVar

        Dim curCol As Long: curCol = 1
        Dim cAcct As Long, cGeo As Long, cSnd As Long, cAgent As Long, cM2O As Long, cMeta As Long, cNat As Long
        cAcct = 0: cGeo = 0: cSnd = 0: cAgent = 0: cM2O = 0: cMeta = 0: cNat = 0
        If dictRedAccts.count > 0 Or dictYellowAccts.count > 0 Then cAcct = curCol: curCol = curCol + 8
        If hasHRCtry Or hasGenCtry Or hasHRGov Or hasGenGov Or dictNatGlobal.count > 0 Then cGeo = curCol: curCol = curCol + 12
        If dictSenderAnalysis.count > 0 Then cSnd = curCol: curCol = curCol + 4 + (maxTrx * 6) + 2
        If dictFlaggedAgents.count > 0 Then cAgent = curCol: curCol = curCol + 8
        If t5Row > 0 Then cM2O = curCol: curCol = curCol + 13
        If t7Row > 0 Then cMeta = curCol: curCol = curCol + 13

        ' ========================================================
        ' OUTPUT
        ' ========================================================
        Dim safeSheetName As String
        If isGlobal Then safeSheetName = "All Countries" Else safeSheetName = Left("Analysis - " & currentTarget, 31)
        On Error Resume Next
        wsSource.Parent.Sheets(safeSheetName).Delete
        On Error GoTo 0
        Set wsDest = wsSource.Parent.Sheets.Add(After:=wsSource.Parent.Sheets(wsSource.Parent.Sheets.count))
        wsDest.Name = safeSheetName

        With wsDest
            r = 3: cr = 3

            ' --- T1: ACCOUNTS SPLIT ---
            If cAcct > 0 Then
                .Columns(cAcct).NumberFormat = "@"   ' account column as text, no apostrophe
                If dictRedAccts.count > 0 Then
                    .Range(.Cells(r, cAcct), .Cells(r, cAcct + 6)).Interior.Color = RGB(255, 199, 206)
                    .Cells(r, cAcct).Value = "SUSPICIOUS ACCOUNTS (>= " & redThresh & ")"
                    .Cells(r, cAcct).Font.Bold = True
                    r = r + 1
                    .Range(.Cells(r, cAcct), .Cells(r, cAcct + 6)).Value = Array("Receiver Account", "Total Transactions", "Unique Senders", "Total Amount", "Country", "Delivery Method", "Provider")
                    .Range(.Cells(r, cAcct), .Cells(r, cAcct + 6)).Font.Bold = True
                    redS = r + 1
                    For Each kVar In dictRedAccts.Keys
                        itemArr = dictAcct(kVar)
                        .Cells(redS, cAcct).Value = CStr(itemArr(6)): .Cells(redS, cAcct).Font.Color = vbRed: .Cells(redS, cAcct).Font.Bold = True
                        .Cells(redS, cAcct + 1).Value = itemArr(0): .Cells(redS, cAcct + 2).Value = itemArr(1)
                        .Cells(redS, cAcct + 3).Value = itemArr(2): .Cells(redS, cAcct + 4).Value = itemArr(3)
                        .Cells(redS, cAcct + 5).Value = itemArr(5): .Cells(redS, cAcct + 6).Value = itemArr(4)
                        If Not dictRedCtry.Exists(itemArr(3)) Then dictRedCtry.Add itemArr(3), True
                        If itemArr(5) = "DIRECT TO BANK" Then redBank = redBank + 1
                        If itemArr(5) = "MOBILE MONEY TRANSFER" Then redMobile = redMobile + 1
                        redS = redS + 1
                    Next kVar
                    If redS > r + 1 Then
                        With .Sort
                            .SortFields.Clear: .SortFields.Add Key:=wsDest.Cells(r + 1, cAcct + 2), Order:=xlDescending
                            .SetRange wsDest.Range(wsDest.Cells(r, cAcct), wsDest.Cells(redS - 1, cAcct + 6)): .Header = xlYes: .Apply
                        End With
                        .Cells(redS, cAcct).Value = "Total (" & dictRedAccts.count & ")"
                        .Cells(redS, cAcct + 1).Value = Application.WorksheetFunction.Sum(.Range(.Cells(r + 1, cAcct + 1), .Cells(redS - 1, cAcct + 1)))
                        .Cells(redS, cAcct + 2).Value = Application.WorksheetFunction.Sum(.Range(.Cells(r + 1, cAcct + 2), .Cells(redS - 1, cAcct + 2)))
                        .Cells(redS, cAcct + 3).Value = Application.WorksheetFunction.Sum(.Range(.Cells(r + 1, cAcct + 3), .Cells(redS - 1, cAcct + 3)))
                        .Cells(redS, cAcct + 4).Value = dictRedCtry.count & " Countries"
                        .Cells(redS, cAcct + 5).Value = "Bank = " & redBank & ", Mobile = " & redMobile
                        .Range(.Cells(redS, cAcct), .Cells(redS, cAcct + 6)).Font.Bold = True
                        .Range(wsDest.Cells(r, cAcct), wsDest.Cells(redS, cAcct + 6)).Borders.LineStyle = xlContinuous
                        r = redS + 2
                    End If
                End If
                If dictYellowAccts.count > 0 Then
                    .Range(.Cells(r, cAcct), .Cells(r, cAcct + 6)).Interior.Color = RGB(255, 235, 156)
                    Dim yText As String
                    If redThresh > 4 Then yText = "REVIEW ACCOUNTS (3-" & (redThresh - 1) & ")" Else yText = "REVIEW ACCOUNTS (< " & redThresh & ")"
                    .Cells(r, cAcct).Value = yText
                    .Cells(r, cAcct).Font.Bold = True
                    r = r + 1
                    .Range(.Cells(r, cAcct), .Cells(r, cAcct + 6)).Value = Array("Receiver Account", "Total Transactions", "Unique Senders", "Total Amount", "Country", "Delivery Method", "Provider")
                    .Range(.Cells(r, cAcct), .Cells(r, cAcct + 6)).Font.Bold = True
                    yelS = r + 1
                    For Each kVar In dictYellowAccts.Keys
                        itemArr = dictAcct(kVar)
                        .Cells(yelS, cAcct).Value = CStr(itemArr(6)): .Cells(yelS, cAcct + 1).Value = itemArr(0): .Cells(yelS, cAcct + 2).Value = itemArr(1)
                        .Cells(yelS, cAcct + 3).Value = itemArr(2): .Cells(yelS, cAcct + 4).Value = itemArr(3): .Cells(yelS, cAcct + 5).Value = itemArr(5): .Cells(yelS, cAcct + 6).Value = itemArr(4)
                        If Not dictYelCtry.Exists(itemArr(3)) Then dictYelCtry.Add itemArr(3), True
                        If itemArr(5) = "DIRECT TO BANK" Then yelBank = yelBank + 1
                        If itemArr(5) = "MOBILE MONEY TRANSFER" Then yelMobile = yelMobile + 1
                        yelS = yelS + 1
                    Next kVar
                    If yelS > r + 1 Then
                        With .Sort
                            .SortFields.Clear: .SortFields.Add Key:=wsDest.Cells(r + 1, cAcct + 2), Order:=xlDescending
                            .SetRange wsDest.Range(wsDest.Cells(r, cAcct), wsDest.Cells(yelS - 1, cAcct + 6)): .Header = xlYes: .Apply
                        End With
                        .Cells(yelS, cAcct).Value = "Total (" & dictYellowAccts.count & ")"
                        .Cells(yelS, cAcct + 1).Value = Application.WorksheetFunction.Sum(.Range(.Cells(r + 1, cAcct + 1), .Cells(yelS - 1, cAcct + 1)))
                        .Cells(yelS, cAcct + 2).Value = Application.WorksheetFunction.Sum(.Range(.Cells(r + 1, cAcct + 2), .Cells(yelS - 1, cAcct + 2)))
                        .Cells(yelS, cAcct + 3).Value = Application.WorksheetFunction.Sum(.Range(.Cells(r + 1, cAcct + 3), .Cells(yelS - 1, cAcct + 3)))
                        .Cells(yelS, cAcct + 4).Value = dictYelCtry.count & " Countries"
                        .Cells(yelS, cAcct + 5).Value = "Bank = " & yelBank & ", Mobile = " & yelMobile
                        .Range(.Cells(yelS, cAcct), .Cells(yelS, cAcct + 6)).Font.Bold = True
                        .Range(wsDest.Cells(r, cAcct), wsDest.Cells(yelS, cAcct + 6)).Borders.LineStyle = xlContinuous
                        r = yelS + 2
                    End If
                End If
            End If

            ' --- T2 & T3 GEOGRAPHIC ---
            If cGeo > 0 Then
                If hasHRCtry Then
                    .Range(.Cells(cr, cGeo), .Cells(cr, cGeo + 10)).Interior.Color = RGB(255, 199, 206)
                    .Cells(cr, cGeo).Value = "HIGH RISK COUNTRIES": .Cells(cr, cGeo).Font.Bold = True
                    cr = cr + 1
                    .Range(.Cells(cr, cGeo), .Cells(cr, cGeo + 10)).Value = Array("Country", "Total Transactions", "Unique Senders", "Direct To Bank", "Mobile Money", "Money In Minutes", "Total Amount", "Suspicious Accounts", "Suspicious Transactions", "Suspicious Unique Senders", "Suspicious Amount")
                    .Range(.Cells(cr, cGeo), .Cells(cr, cGeo + 10)).Font.Bold = True
                    hRS = cr + 1
                    For Each kVar In dictCountry.Keys
                        cArr = dictCountry(kVar)
                        If cArr(7) > 0 Then
                            .Cells(hRS, cGeo).Value = kVar
                            For mIdx = 0 To 9: .Cells(hRS, cGeo + 1 + mIdx).Value = cArr(mIdx): Next mIdx
                            .Cells(hRS, cGeo).Font.Color = vbRed: .Cells(hRS, cGeo).Font.Bold = True
                            .Range(.Cells(hRS, cGeo + 7), .Cells(hRS, cGeo + 10)).Font.Color = vbRed: .Range(.Cells(hRS, cGeo + 7), .Cells(hRS, cGeo + 10)).Font.Bold = True
                            hRS = hRS + 1
                        End If
                    Next kVar
                    If hRS > cr + 1 Then
                        With .Sort
                            .SortFields.Clear: .SortFields.Add Key:=wsDest.Cells(cr + 1, cGeo + 8), Order:=xlDescending
                            .SetRange wsDest.Range(wsDest.Cells(cr, cGeo), wsDest.Cells(hRS - 1, cGeo + 10)): .Header = xlYes: .Apply
                        End With
                        .Cells(hRS, cGeo).Value = "Total"
                        For cIdx = 1 To 10: .Cells(hRS, cGeo + cIdx).Value = Application.WorksheetFunction.Sum(.Range(.Cells(cr + 1, cGeo + cIdx), .Cells(hRS - 1, cGeo + cIdx))): Next cIdx
                        .Range(.Cells(hRS, cGeo), .Cells(hRS, cGeo + 10)).Font.Color = vbBlack: .Range(.Cells(hRS, cGeo), .Cells(hRS, cGeo + 10)).Font.Bold = True
                        .Range(.Cells(cr, cGeo), .Cells(hRS, cGeo + 10)).Borders.LineStyle = xlContinuous
                        cr = hRS + 2
                    End If
                End If
                gCtryStartRow = cr
                If hasGenCtry Then
                    .Range(.Cells(cr, cGeo), .Cells(cr, cGeo + 6)).Interior.Color = RGB(221, 235, 247)
                    .Cells(cr, cGeo).Value = "GENERAL COUNTRY OVERVIEW": .Cells(cr, cGeo).Font.Bold = True
                    cr = cr + 1
                    .Range(.Cells(cr, cGeo), .Cells(cr, cGeo + 6)).Value = Array("Country", "Total Transactions", "Unique Senders", "Direct To Bank", "Mobile Money", "Money In Minutes", "Total Amount")
                    .Range(.Cells(cr, cGeo), .Cells(cr, cGeo + 6)).Font.Bold = True
                    gRS = cr + 1
                    For Each kVar In dictCountry.Keys
                        cArr = dictCountry(kVar)
                        If cArr(7) = 0 Then
                            .Cells(gRS, cGeo).Value = kVar
                            For mIdx = 0 To 5: .Cells(gRS, cGeo + 1 + mIdx).Value = cArr(mIdx): Next mIdx
                            gRS = gRS + 1
                        End If
                    Next kVar
                    If gRS > cr + 1 Then
                        With .Sort
                            .SortFields.Clear: .SortFields.Add Key:=wsDest.Cells(cr + 1, cGeo + 1), Order:=xlDescending
                            .SetRange wsDest.Range(wsDest.Cells(cr, cGeo), wsDest.Cells(gRS - 1, cGeo + 6)): .Header = xlYes: .Apply
                        End With
                        .Cells(gRS, cGeo).Value = "Total"
                        For cIdx = 1 To 6: .Cells(gRS, cGeo + cIdx).Value = Application.WorksheetFunction.Sum(.Range(.Cells(cr + 1, cGeo + cIdx), .Cells(gRS - 1, cGeo + cIdx))): Next cIdx
                        .Range(.Cells(gRS, cGeo), .Cells(gRS, cGeo + 6)).Font.Color = vbBlack: .Range(.Cells(gRS, cGeo), .Cells(gRS, cGeo + 6)).Font.Bold = True
                        .Range(.Cells(cr, cGeo), .Cells(gRS, cGeo + 6)).Borders.LineStyle = xlContinuous
                        cr = gRS + 2
                    End If
                End If
                If dictNatGlobal.count > 0 Then
                    If Not hasHRCtry And Not hasGenCtry Then cNat = cGeo Else cNat = cGeo + 8
                    natRow = gCtryStartRow
                    .Range(.Cells(natRow, cNat), .Cells(natRow, cNat + 2)).Interior.Color = RGB(226, 239, 218)
                    .Cells(natRow, cNat).Value = "GLOBAL NATIONALITY OVERVIEW": .Cells(natRow, cNat).Font.Bold = True
                    natRow = natRow + 1
                    .Range(.Cells(natRow, cNat), .Cells(natRow, cNat + 2)).Value = Array("Nationality", "Total Transactions", "Unique Senders")
                    .Range(.Cells(natRow, cNat), .Cells(natRow, cNat + 2)).Font.Bold = True
                    nStart = natRow + 1
                    For Each kVar In dictNatGlobal.Keys
                        nArrGlobal = dictNatGlobal(kVar): .Cells(nStart, cNat).Value = kVar
                        .Cells(nStart, cNat + 1).Value = nArrGlobal(0): .Cells(nStart, cNat + 2).Value = nArrGlobal(1): nStart = nStart + 1
                    Next kVar
                    If nStart > natRow + 1 Then
                        With .Sort
                            .SortFields.Clear: .SortFields.Add Key:=wsDest.Cells(natRow + 1, cNat + 1), Order:=xlDescending
                            .SetRange wsDest.Range(wsDest.Cells(natRow, cNat), wsDest.Cells(nStart - 1, cNat + 2)): .Header = xlYes: .Apply
                        End With
                        .Cells(nStart, cNat).Value = "Total"
                        .Cells(nStart, cNat + 1).Value = Application.WorksheetFunction.Sum(.Range(.Cells(natRow + 1, cNat + 1), .Cells(nStart - 1, cNat + 1)))
                        .Cells(nStart, cNat + 2).Value = Application.WorksheetFunction.Sum(.Range(.Cells(natRow + 1, cNat + 2), .Cells(nStart - 1, cNat + 2)))
                        .Range(.Cells(nStart, cNat), .Cells(nStart, cNat + 2)).Font.Bold = True
                        .Range(.Cells(natRow, cNat), .Cells(nStart, cNat + 2)).Borders.LineStyle = xlContinuous
                        natRow = nStart + 2
                    End If
                Else
                    natRow = cr
                End If
                If natRow > cr Then cr = natRow
                If hasHRGov Then
                    .Range(.Cells(cr, cGeo), .Cells(cr, cGeo + 6)).Interior.Color = RGB(255, 199, 206)
                    .Cells(cr, cGeo).Value = "HIGH RISK GOVERNORATES": .Cells(cr, cGeo).Font.Bold = True
                    cr = cr + 1
                    .Range(.Cells(cr, cGeo), .Cells(cr, cGeo + 6)).Value = Array("Governorate", "Total Transactions", "Unique Senders", "Total Amount", "Suspicious Transactions", "Suspicious Unique Senders", "Suspicious Amount")
                    .Range(.Cells(cr, cGeo), .Cells(cr, cGeo + 6)).Font.Bold = True
                    hG = cr + 1
                    For Each kVar In dictGov.Keys
                        gArr = dictGov(kVar)
                        If gArr(3) > 0 Then
                            .Cells(hG, cGeo).Value = kVar
                            For mIdx = 0 To 5: .Cells(hG, cGeo + 1 + mIdx).Value = gArr(mIdx): Next mIdx
                            .Cells(hG, cGeo).Font.Color = vbRed: .Cells(hG, cGeo).Font.Bold = True
                            .Range(.Cells(hG, cGeo + 4), .Cells(hG, cGeo + 6)).Font.Color = vbRed: .Range(.Cells(hG, cGeo + 4), .Cells(hG, cGeo + 6)).Font.Bold = True
                            hG = hG + 1
                        End If
                    Next kVar
                    If hG > cr + 1 Then
                        With .Sort
                            .SortFields.Clear: .SortFields.Add Key:=wsDest.Cells(cr + 1, cGeo + 4), Order:=xlDescending
                            .SetRange wsDest.Range(wsDest.Cells(cr, cGeo), wsDest.Cells(hG - 1, cGeo + 6)): .Header = xlYes: .Apply
                        End With
                        .Cells(hG, cGeo).Value = "Total"
                        For cIdx = 1 To 6: .Cells(hG, cGeo + cIdx).Value = Application.WorksheetFunction.Sum(.Range(.Cells(cr + 1, cGeo + cIdx), .Cells(hG - 1, cGeo + cIdx))): Next cIdx
                        .Range(.Cells(hG, cGeo), .Cells(hG, cGeo + 6)).Font.Color = vbBlack: .Range(.Cells(hG, cGeo), .Cells(hG, cGeo + 6)).Font.Bold = True
                        .Range(.Cells(cr, cGeo), .Cells(hG, cGeo + 6)).Borders.LineStyle = xlContinuous
                        cr = hG + 2
                    End If
                End If
                If hasGenGov Then
                    .Range(.Cells(cr, cGeo), .Cells(cr, cGeo + 3)).Interior.Color = RGB(221, 235, 247)
                    .Cells(cr, cGeo).Value = "GENERAL GOVERNORATE OVERVIEW": .Cells(cr, cGeo).Font.Bold = True
                    cr = cr + 1
                    .Range(.Cells(cr, cGeo), .Cells(cr, cGeo + 3)).Value = Array("Governorate", "Total Transactions", "Unique Senders", "Total Amount")
                    .Range(.Cells(cr, cGeo), .Cells(cr, cGeo + 3)).Font.Bold = True
                    gg = cr + 1
                    For Each kVar In dictGov.Keys
                        gArr = dictGov(kVar)
                        If gArr(3) = 0 Then
                            .Cells(gg, cGeo).Value = kVar
                            For mIdx = 0 To 2: .Cells(gg, cGeo + 1 + mIdx).Value = gArr(mIdx): Next mIdx
                            gg = gg + 1
                        End If
                    Next kVar
                    If gg > cr + 1 Then
                        With .Sort
                            .SortFields.Clear: .SortFields.Add Key:=wsDest.Cells(cr + 1, cGeo + 1), Order:=xlDescending
                            .SetRange wsDest.Range(wsDest.Cells(cr, cGeo), wsDest.Cells(gg - 1, cGeo + 3)): .Header = xlYes: .Apply
                        End With
                        .Cells(gg, cGeo).Value = "Total"
                        For cIdx = 1 To 3: .Cells(gg, cGeo + cIdx).Value = Application.WorksheetFunction.Sum(.Range(.Cells(cr + 1, cGeo + cIdx), .Cells(gg - 1, cGeo + cIdx))): Next cIdx
                        .Range(.Cells(gg, cGeo), .Cells(gg, cGeo + 3)).Font.Color = vbBlack: .Range(.Cells(gg, cGeo), .Cells(gg, cGeo + 3)).Font.Bold = True
                        .Range(.Cells(cr, cGeo), .Cells(gg, cGeo + 3)).Borders.LineStyle = xlContinuous
                        cr = gg + 2
                    End If
                End If
            End If

            ' --- T4: SUSPICIOUS UNIQUE SENDERS ---
            ' Layout: Sender | Actor ID | Governorate | Primary Country | [6-col blocks] | Note
            If cSnd > 0 Then
                Dim noteCol As Long: noteCol = cSnd + 4 + (maxTrx * 6)
                .Columns(cSnd).NumberFormat = "0"                ' Sender MSISDN as number (no E+, no triangle)
                For cIdx = 1 To maxTrx
                    .Columns(cSnd + 5 + (cIdx - 1) * 6).NumberFormat = "@"   ' Account
                    .Columns(cSnd + 8 + (cIdx - 1) * 6).NumberFormat = "@"   ' MTCN
                Next cIdx
                .Range(.Cells(3, cSnd), .Cells(3, noteCol)).Interior.Color = RGB(226, 239, 218)
                .Cells(3, cSnd).Value = "SUSPICIOUS UNIQUE SENDERS": .Cells(3, cSnd).Font.Bold = True
                .Cells(4, cSnd).Value = "Sender MSISDN": .Cells(4, cSnd + 1).Value = "Actor ID"
                .Cells(4, cSnd + 2).Value = "Governorate": .Cells(4, cSnd + 3).Value = "Primary Country"
                For cIdx = 1 To maxTrx
                    .Cells(4, cSnd + 4 + (cIdx - 1) * 6).Value = "Amount " & cIdx
                    .Cells(4, cSnd + 5 + (cIdx - 1) * 6).Value = "Account " & cIdx
                    .Cells(4, cSnd + 6 + (cIdx - 1) * 6).Value = Ordinal(cIdx) & " Destination Country"
                    .Cells(4, cSnd + 7 + (cIdx - 1) * 6).Value = Ordinal(cIdx) & " Receiver Name"
                    .Cells(4, cSnd + 8 + (cIdx - 1) * 6).Value = "MTCN " & cIdx
                    .Cells(4, cSnd + 9 + (cIdx - 1) * 6).Value = "Date " & cIdx
                Next cIdx
                .Cells(4, noteCol).Value = "Note"
                r4 = 5
                For Each kVar In dictSenderAnalysis.Keys
                    sp = Split(kVar, "|")
                    .Cells(r4, cSnd).Value = NumIf(sp(0))
                    If dictSenderActor.Exists(kVar) Then .Cells(r4, cSnd + 1).Value = dictSenderActor(kVar)
                    .Cells(r4, cSnd + 2).Value = sp(1)
                    sArr = dictSenderAnalysis(kVar): .Cells(r4, cSnd + 3).Value = sArr(2)
                    trxPs = Split(sArr(1), "|")
                    For mIdx = 0 To UBound(trxPs)
                        d = Split(trxPs(mIdx), "^")
                        .Cells(r4, cSnd + 4 + (mIdx * 6)).Value = d(0)
                        .Cells(r4, cSnd + 5 + (mIdx * 6)).Value = d(1)
                        .Cells(r4, cSnd + 6 + (mIdx * 6)).Value = d(2)
                        .Cells(r4, cSnd + 7 + (mIdx * 6)).Value = d(3)
                        .Cells(r4, cSnd + 8 + (mIdx * 6)).Value = d(4)
                        If d(5) <> "" Then .Cells(r4, cSnd + 9 + (mIdx * 6)).Value = CDate(Val(d(5))) Else .Cells(r4, cSnd + 9 + (mIdx * 6)).Value = ""
                    Next mIdx
                    ' Note: all transactions or only suspicious, per user choice
                    Dim noteBody As String
                    If noteAllTrx Then
                        If dictNoteAll.Exists(kVar) Then noteBody = dictNoteAll(kVar) Else noteBody = ""
                    Else
                        If dictNoteSus.Exists(kVar) Then noteBody = dictNoteSus(kVar) Else noteBody = ""
                    End If
                    If Len(noteBody) > 30000 Then noteBody = Left(noteBody, 30000) & vbLf & "...(truncated)"
                    .Cells(r4, noteCol).Value = ChrW(1575) & ChrW(1604) & ChrW(1586) & ChrW(1576) & ChrW(1608) & ChrW(1606) & " " & _
                        ChrW(1602) & ChrW(1575) & ChrW(1605) & " " & ChrW(1576) & ChrW(1575) & ChrW(1604) & ChrW(1578) & ChrW(1581) & ChrW(1608) & ChrW(1610) & ChrW(1604) & " " & _
                        ChrW(1575) & ChrW(1604) & ChrW(1609) & ":" & vbLf & noteBody
                    .Cells(r4, noteCol).WrapText = True
                    .Cells(r4, noteCol).VerticalAlignment = xlTop
                    .Cells(r4, noteCol).HorizontalAlignment = xlRight
                    r4 = r4 + 1
                Next kVar
                If r4 > 5 Then
                    .Range(.Cells(4, cSnd), .Cells(r4 - 1, noteCol)).Borders.LineStyle = xlContinuous
                    For cIdx = 1 To maxTrx
                        .Columns(cSnd + 9 + (cIdx - 1) * 6).NumberFormat = "yyyy-mm-dd hh:mm:ss"
                        .Columns(cSnd + 4 + (cIdx - 1) * 6).NumberFormat = "#,##0"
                    Next cIdx
                End If
            End If

            ' --- T5: AGENT CASH-IN RISK ---
            If cAgent > 0 Then
                .Columns(cAgent).NumberFormat = "@"
                .Range(.Cells(3, cAgent), .Cells(3, cAgent + 6)).Interior.Color = RGB(252, 213, 180)
                .Cells(3, cAgent).Value = "AGENT CASH-IN RISK (>= " & agtThresh & " " & IIf(agtUseUnique, "Customers", "Transactions") & ")"
                .Cells(3, cAgent).Font.Bold = True
                .Cells(4, cAgent).Value = "Agent MSISDN": .Cells(4, cAgent + 1).Value = "Hit Count": .Cells(4, cAgent + 2).Value = "Sus Cash-In Amount"
                .Cells(4, cAgent + 3).Value = "Agent City": .Cells(4, cAgent + 4).Value = "% Same City": .Cells(4, cAgent + 5).Value = "% Other City": .Cells(4, cAgent + 6).Value = "Associated Countries Stats"
                .Range(.Cells(4, cAgent), .Cells(4, cAgent + 6)).Font.Bold = True
                Dim agtRow As Long: agtRow = 5
                For Each kVar In dictFlaggedAgents.Keys
                    Dim aMetrics As Variant: aMetrics = dictFlaggedAgentMetrics(kVar)
                    .Cells(agtRow, cAgent).Value = kVar: .Cells(agtRow, cAgent + 1).Value = dictFlaggedAgents(kVar)
                    .Cells(agtRow, cAgent + 2).Value = aMetrics(0): .Cells(agtRow, cAgent + 3).Value = aMetrics(1)
                    .Cells(agtRow, cAgent + 4).Value = aMetrics(2): .Cells(agtRow, cAgent + 5).Value = aMetrics(3)
                    .Cells(agtRow, cAgent + 6).Value = dictFlaggedAgentStrings(kVar)
                    agtRow = agtRow + 1
                Next kVar
                If agtRow > 5 Then
                    With .Sort
                        .SortFields.Clear: .SortFields.Add Key:=wsDest.Cells(5, cAgent + 1), Order:=xlDescending
                        .SetRange wsDest.Range(wsDest.Cells(4, cAgent), wsDest.Cells(agtRow - 1, cAgent + 6)): .Header = xlYes: .Apply
                    End With
                    .Range(.Cells(4, cAgent), .Cells(agtRow - 1, cAgent + 6)).Borders.LineStyle = xlContinuous
                    .Columns(cAgent + 2).NumberFormat = "#,##0": .Columns(cAgent + 4).NumberFormat = "0%": .Columns(cAgent + 5).NumberFormat = "0%"
                End If
            End If

            ' --- T6: MANY TO ONE (BY MODEL ID) ---
            If cM2O > 0 Then
                .Columns(cM2O).NumberFormat = "0": .Columns(cM2O + 5).NumberFormat = "@": .Columns(cM2O + 8).NumberFormat = "@": .Columns(cM2O + 9).NumberFormat = "@"
                .Range(.Cells(3, cM2O), .Cells(3, cM2O + 11)).Interior.Color = RGB(255, 230, 153)
                .Cells(3, cM2O).Value = "MANY TO ONE RISK (BY MODEL ID)": .Cells(3, cM2O).Font.Bold = True
                .Range(.Cells(4, cM2O), .Cells(4, cM2O + 11)).Value = Array("Sender MSISDN", "Governorate", "Receiver's Name", "Amount", "Country", "Model ID", "Hit Count", "Delivery Method", "Bank/Wallet Account", "MTCN", "Date of Transfer", "Total Amount")
                .Range(.Cells(5, cM2O), .Cells(4 + t5Row, cM2O + 11)).Value = arrT5
                With .Sort
                    .SortFields.Clear
                    .SortFields.Add Key:=wsDest.Range(wsDest.Cells(5, cM2O + 11), wsDest.Cells(4 + t5Row, cM2O + 11)), Order:=xlDescending
                    .SortFields.Add Key:=wsDest.Range(wsDest.Cells(5, cM2O + 5), wsDest.Cells(4 + t5Row, cM2O + 5)), Order:=xlAscending
                    .SetRange wsDest.Range(wsDest.Cells(4, cM2O), wsDest.Cells(4 + t5Row, cM2O + 11)): .Header = xlYes: .Apply
                End With
                Application.DisplayAlerts = False
                startR = 5: currentModel = CStr(.Cells(5, cM2O + 5).Value)
                For r5 = 5 To 4 + t5Row + 1
                    If r5 <= 4 + t5Row Then iterModel = CStr(.Cells(r5, cM2O + 5).Value) Else iterModel = "END_OF_LOOP"
                    If iterModel <> currentModel Then
                        If r5 - 1 > startR Then .Range(.Cells(startR, cM2O + 11), .Cells(r5 - 1, cM2O + 11)).Merge
                        .Cells(startR, cM2O + 11).VerticalAlignment = xlCenter
                        startR = r5: currentModel = iterModel
                    End If
                Next r5
                Application.DisplayAlerts = True
                .Range(.Cells(4, cM2O), .Cells(4 + t5Row, cM2O + 11)).Borders.LineStyle = xlContinuous
                .Columns(cM2O + 10).NumberFormat = "yyyy-mm-dd hh:mm:ss": .Columns(cM2O + 3).NumberFormat = "#,##0": .Columns(cM2O + 11).NumberFormat = "#,##0"
                For tRow = 5 To 4 + t5Row
                    chkSender = CleanMSISDN(CStr(.Cells(tRow, cM2O).Value))
                    If Not dictSusSendersCheck.Exists(chkSender) Then .Cells(tRow, cM2O).Font.Color = vbRed Else .Cells(tRow, cM2O).Font.Color = vbBlack
                Next tRow
            End If

            ' --- T7: MANY TO ONE (BY NAME SOUND / METAPHONE) ---
            If cMeta > 0 Then
                .Columns(cMeta).NumberFormat = "0": .Columns(cMeta + 8).NumberFormat = "@": .Columns(cMeta + 9).NumberFormat = "@"
                .Range(.Cells(3, cMeta), .Cells(3, cMeta + 11)).Interior.Color = RGB(204, 236, 255)
                .Cells(3, cMeta).Value = "MANY TO ONE RISK (BY NAME SOUND / METAPHONE)": .Cells(3, cMeta).Font.Bold = True
                .Range(.Cells(4, cMeta), .Cells(4, cMeta + 11)).Value = Array("Sender MSISDN", "Governorate", "Receiver's Name", "Amount", "Country", "Metaphone", "Hit Count", "Delivery Method", "Bank/Wallet Account", "MTCN", "Date of Transfer", "Total Amount")
                .Range(.Cells(5, cMeta), .Cells(4 + t7Row, cMeta + 11)).Value = arrT7
                With .Sort
                    .SortFields.Clear
                    .SortFields.Add Key:=wsDest.Range(wsDest.Cells(5, cMeta + 11), wsDest.Cells(4 + t7Row, cMeta + 11)), Order:=xlDescending
                    .SortFields.Add Key:=wsDest.Range(wsDest.Cells(5, cMeta + 5), wsDest.Cells(4 + t7Row, cMeta + 5)), Order:=xlAscending
                    .SetRange wsDest.Range(wsDest.Cells(4, cMeta), wsDest.Cells(4 + t7Row, cMeta + 11)): .Header = xlYes: .Apply
                End With
                Application.DisplayAlerts = False
                startR = 5: currentModel = CStr(.Cells(5, cMeta + 5).Value)
                For r5 = 5 To 4 + t7Row + 1
                    If r5 <= 4 + t7Row Then iterModel = CStr(.Cells(r5, cMeta + 5).Value) Else iterModel = "END_OF_LOOP"
                    If iterModel <> currentModel Then
                        If r5 - 1 > startR Then .Range(.Cells(startR, cMeta + 11), .Cells(r5 - 1, cMeta + 11)).Merge
                        .Cells(startR, cMeta + 11).VerticalAlignment = xlCenter
                        startR = r5: currentModel = iterModel
                    End If
                Next r5
                Application.DisplayAlerts = True
                .Range(.Cells(4, cMeta), .Cells(4 + t7Row, cMeta + 11)).Borders.LineStyle = xlContinuous
                .Columns(cMeta + 10).NumberFormat = "yyyy-mm-dd hh:mm:ss": .Columns(cMeta + 3).NumberFormat = "#,##0": .Columns(cMeta + 11).NumberFormat = "#,##0"
                For tRow = 5 To 4 + t7Row
                    chkSender = CleanMSISDN(CStr(.Cells(tRow, cMeta).Value))
                    If Not dictSusSendersCheck.Exists(chkSender) Then .Cells(tRow, cMeta).Font.Color = vbRed Else .Cells(tRow, cMeta).Font.Color = vbBlack
                Next tRow
            End If

            ' --- NUMBER FORMATTING ---
            If cAcct > 0 Then .Columns(cAcct + 3).NumberFormat = "#,##0"
            If cGeo > 0 Then .Columns(cGeo + 3).NumberFormat = "#,##0": .Columns(cGeo + 6).NumberFormat = "#,##0": .Columns(cGeo + 10).NumberFormat = "#,##0"
            If cNat > 0 Then .Columns(cNat + 1).NumberFormat = "#,##0": .Columns(cNat + 2).NumberFormat = "#,##0"

            For Each cl In .UsedRange
                If CStr(cl.Value) = "0" Then cl.ClearContents
            Next cl

            If globalDateFound Then .Range("A1").Value = Format(globalMin, "dd-MMM-yyyy") & " to " & Format(globalMax, "dd-MMM-yyyy") Else .Range("A1").Value = "Date Range Not Found"
            .Range("A1").Font.Bold = True: .Range("A1").Font.Color = vbBlack

            ' ========================================================
            ' RAW "SUSPICIOUS DATA" TABLE
            ' ========================================================
            If Not isGlobal And rawCount > 0 Then
                Dim finalMaxRow As Long
                On Error Resume Next
                finalMaxRow = .Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
                On Error GoTo 0
                If finalMaxRow < 3 Then finalMaxRow = 3
                Dim rawStartRow As Long: rawStartRow = finalMaxRow + 4

                .Range(.Cells(rawStartRow, 1), .Cells(rawStartRow, 5)).Interior.Color = RGB(226, 239, 218)
                .Cells(rawStartRow, 1).Value = "SUSPICIOUS DATA - " & currentTarget
                .Cells(rawStartRow, 1).Font.Bold = True: .Cells(rawStartRow, 1).Font.Size = 14

                ' Headers (with renames + injected counts)
                Dim rawHeaders() As Variant: ReDim rawHeaders(1 To rawCols)
                Dim hd As String
                nc = 1
                For cIdx = 1 To 28
                    hd = CStr(wsSource.Cells(1, cIdx).Value)
                    If cIdx = recNameSrcCol Then hd = "RECEIVER_NAME"
                    If motherSrcCol > 0 And cIdx = motherSrcCol Then hd = "Mother Name"
                    If InStr(UCase(hd), "DESTINATION_COUNTRY") > 0 Then hd = "DESTINATION_COUNTRY"
                    If InStr(UCase(hd), "GOVERNORATE") > 0 Then hd = "Governorate"
                    rawHeaders(nc) = hd: nc = nc + 1
                    If cIdx = recNameSrcCol Then rawHeaders(nc) = "Name Count (" & currentTarget & ")": nc = nc + 1
                    If cIdx = 26 Then rawHeaders(nc) = "Account Count": nc = nc + 1
                    If cIdx = 28 Then rawHeaders(nc) = "Model ID Count": nc = nc + 1
                    If motherSrcCol > 0 And cIdx = motherSrcCol Then rawHeaders(nc) = "Mother Name Count (" & currentTarget & ")": nc = nc + 1
                Next cIdx
                If useAgent Then
                    rawHeaders(posAgent1) = "Cash-In Agent MSISDN"
                    rawHeaders(posAgent2) = "Agent Count in " & currentTarget
                    rawHeaders(posAgent3) = "Agent Global Stats (If Flagged)"
                End If

                .Cells(rawStartRow + 2, 1).Resize(1, rawCols).Value = rawHeaders
                .Cells(rawStartRow + 2, 1).Resize(1, rawCols).Font.Bold = True
                .Cells(rawStartRow + 2, 1).Resize(1, rawCols).Interior.Color = RGB(191, 191, 191)
                .Range(.Cells(rawStartRow + 2, 1), .Cells(rawStartRow + 2, rawCols)).AutoFilter

                ' Key ID columns formatting by POSITION
                .Cells(rawStartRow + 3, outCol(colMTCN)).Resize(rawCount, 1).NumberFormat = "@"
                .Cells(rawStartRow + 3, outCol(12)).Resize(rawCount, 1).NumberFormat = "0"   ' MSISDN as number, full digits, no triangle
                .Cells(rawStartRow + 3, outCol(26)).Resize(rawCount, 1).NumberFormat = "@"
                .Cells(rawStartRow + 3, outCol(28)).Resize(rawCount, 1).NumberFormat = "@"
                If useAgent Then
                    .Cells(rawStartRow + 3, posAgent1).Resize(rawCount, 1).NumberFormat = "@"
                    .Cells(rawStartRow + 3, posAgent3).Resize(rawCount, 1).NumberFormat = "@"
                End If
                ' Also text-format any other ID-like header columns
                Dim hdName As String
                For cIdx = 1 To rawCols
                    hdName = UCase(Trim(.Cells(rawStartRow + 2, cIdx).Value))
                    If cIdx <> outCol(12) And InStr(hdName, "DATE") = 0 And InStr(hdName, "COUNT") = 0 And InStr(hdName, "AMOUNT") = 0 Then
                        If InStr(hdName, "MTCN") > 0 Or InStr(hdName, "ACCOUNT") > 0 _
                           Or InStr(hdName, "RECRUITER") > 0 Or InStr(hdName, "MODEL ID") > 0 Or InStr(hdName, "ID") > 0 Or InStr(hdName, "AGENT") > 0 Then
                            .Cells(rawStartRow + 3, cIdx).Resize(rawCount, 1).NumberFormat = "@"
                        End If
                    End If
                Next cIdx

                ' Dump data
                .Cells(rawStartRow + 3, 1).Resize(rawCount, rawCols).Value = rawArr
                .Range(.Cells(rawStartRow + 2, 1), .Cells(rawStartRow + 2 + rawCount, rawCols)).Borders.LineStyle = xlContinuous
                .Range(.Cells(rawStartRow + 3, 1), .Cells(rawStartRow + 2 + rawCount, rawCols)).Font.Bold = False
                .Range(.Cells(rawStartRow + 3, 1), .Cells(rawStartRow + 2 + rawCount, rawCols)).Font.Color = vbBlack

                ' Dates -> real dates + format ; amounts -> #,##0
                Dim rr As Long, pv As Variant, colV As Variant
                For cIdx = 1 To rawCols
                    hdName = UCase(Trim(.Cells(rawStartRow + 2, cIdx).Value))
                    If hdName = "DATE_ONLY" Then
                        colV = .Cells(rawStartRow + 3, cIdx).Resize(rawCount, 1).Value
                        For rr = 1 To rawCount
                            pv = ParseDate(colV(rr, 1)): If IsDate(pv) Then colV(rr, 1) = pv
                        Next rr
                        .Cells(rawStartRow + 3, cIdx).Resize(rawCount, 1).Value = colV
                        .Cells(rawStartRow + 3, cIdx).Resize(rawCount, 1).NumberFormat = "yyyy-mm-dd"
                    ElseIf InStr(hdName, "DATE") > 0 Or hdName = "SEND_AT" Or hdName = "CREATED_AT" Then
                        colV = .Cells(rawStartRow + 3, cIdx).Resize(rawCount, 1).Value
                        For rr = 1 To rawCount
                            pv = ParseDate(colV(rr, 1)): If IsDate(pv) Then colV(rr, 1) = pv
                        Next rr
                        .Cells(rawStartRow + 3, cIdx).Resize(rawCount, 1).Value = colV
                        .Cells(rawStartRow + 3, cIdx).Resize(rawCount, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
                    ElseIf hdName = "AMOUNT" Or InStr(hdName, "TOTAL AMOUNT") > 0 Or InStr(hdName, "PRINCIPAL AMOUNT") > 0 Then
                        .Cells(rawStartRow + 3, cIdx).Resize(rawCount, 1).NumberFormat = "#,##0"
                    End If
                Next cIdx
            End If

            .UsedRange.HorizontalAlignment = xlCenter
            .Columns.AutoFit
            ' Keep the Note column readable (wrapped, fixed width, right-aligned for Arabic)
            If cSnd > 0 Then
                Dim noteColFix As Long: noteColFix = cSnd + 4 + (maxTrx * 6)
                .Columns(noteColFix).ColumnWidth = 55
                .Columns(noteColFix).WrapText = True
                .Columns(noteColFix).HorizontalAlignment = xlRight
            End If
        End With
NextTarget:
    Next tIdx

    Application.ScreenUpdating = True
    MsgBox "Analysis & Extractions Completed!", vbInformation
End Sub

' --- SMART MSISDN CLEANER (strips only + and ' ; keeps leading zeros) ---
Function CleanMSISDN(ByVal num As String) As String
    num = Trim(num)
    Do While Left(num, 1) = "+" Or Left(num, 1) = "'"
        num = Mid(num, 2)
    Loop
    CleanMSISDN = num
End Function

' --- ROBUST DATE PARSER : real dates/serials AND text like "6/7/26 8:00 AM" (mm/dd/yyyy) ---
Function ParseDate(ByVal v As Variant) As Variant
    Dim s As String, dpart As String, tpart As String
    Dim parts() As String, sp() As String, tp() As String
    Dim mo As Integer, dy As Integer, yr As Integer
    Dim hh As Integer, mn As Integer, ss As Integer
    Dim isPM As Boolean, isAM As Boolean, up As String
    ParseDate = Empty
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then ParseDate = CDate(v): Exit Function
    If IsNumeric(v) And VarType(v) <> vbString Then
        On Error Resume Next
        ParseDate = CDate(v)
        On Error GoTo 0
        Exit Function
    End If
    s = Trim(CStr(v))
    If s = "" Then Exit Function
    up = UCase(s)
    If InStr(up, "PM") > 0 Then isPM = True
    If InStr(up, "AM") > 0 Then isAM = True
    s = Trim(Replace(Replace(up, "AM", ""), "PM", ""))
    sp = Split(s, " ")
    dpart = sp(0): tpart = ""
    If UBound(sp) >= 1 Then tpart = Trim(sp(UBound(sp)))
    If InStr(dpart, "/") > 0 Then
        parts = Split(dpart, "/")
    ElseIf InStr(dpart, "-") > 0 And InStr(dpart, ":") = 0 Then
        parts = Split(dpart, "-")
    Else
        Exit Function
    End If
    If UBound(parts) < 2 Then Exit Function
    mo = CInt(Val(parts(0))): dy = CInt(Val(parts(1))): yr = CInt(Val(parts(2)))
    If yr < 100 Then yr = 2000 + yr
    If mo < 1 Or mo > 12 Or dy < 1 Or dy > 31 Then Exit Function
    hh = 0: mn = 0: ss = 0
    If tpart <> "" And InStr(tpart, ":") > 0 Then
        tp = Split(tpart, ":")
        If UBound(tp) >= 0 Then hh = CInt(Val(tp(0)))
        If UBound(tp) >= 1 Then mn = CInt(Val(tp(1)))
        If UBound(tp) >= 2 Then ss = CInt(Val(tp(2)))
    End If
    If isPM And hh < 12 Then hh = hh + 12
    If isAM And hh = 12 Then hh = 0
    On Error Resume Next
    ParseDate = DateSerial(yr, mo, dy) + TimeSerial(hh, mn, ss)
    On Error GoTo 0
End Function

' --- METAPHONE (practical, deterministic ; Latin/English names) ---
Function Metaphone(ByVal Word As String) As String
    Dim u As String, a As String, c As String
    Dim i As Long, L As Long, p As Long
    Dim curr As String, prev As String, nxt As String, nxt2 As String
    Dim m As String
    Const VWL As String = "AEIOU"

    u = UCase$(Word): a = ""
    For i = 1 To Len(u)
        c = Mid$(u, i, 1)
        If c >= "A" And c <= "Z" Then a = a & c
    Next i
    If Len(a) = 0 Then Metaphone = "": Exit Function

    Select Case Left$(a, 2)
        Case "AE", "GN", "KN", "PN", "WR": a = Mid$(a, 2)
    End Select
    If Left$(a, 1) = "X" Then a = "S" & Mid$(a, 2)
    If Left$(a, 2) = "WH" Then a = "W" & Mid$(a, 3)

    L = Len(a): m = "": p = 1
    Do While p <= L
        curr = Mid$(a, p, 1)
        If p > 1 Then prev = Mid$(a, p - 1, 1) Else prev = ""
        If p < L Then nxt = Mid$(a, p + 1, 1) Else nxt = ""
        If p < L - 1 Then nxt2 = Mid$(a, p + 2, 1) Else nxt2 = ""

        If curr = prev And curr <> "C" Then
            p = p + 1
        Else
            Select Case curr
                Case "A", "E", "I", "O", "U"
                    If p = 1 Then m = m & curr
                Case "B"
                    If Not (p = L And prev = "M") Then m = m & "B"
                Case "C"
                    If nxt = "I" And nxt2 = "A" Then
                        m = m & "X"
                    ElseIf nxt = "H" Then
                        If prev = "S" Then m = m & "K" Else m = m & "X"
                        p = p + 1
                    ElseIf nxt = "I" Or nxt = "E" Or nxt = "Y" Then
                        If prev <> "S" Then m = m & "S"
                    Else
                        m = m & "K"
                    End If
                Case "D"
                    If nxt = "G" And (nxt2 = "E" Or nxt2 = "I" Or nxt2 = "Y") Then
                        m = m & "J": p = p + 1
                    Else
                        m = m & "T"
                    End If
                Case "F": m = m & "F"
                Case "G"
                    If nxt = "H" Then
                        If p > 1 And InStr(VWL, prev) = 0 Then
                        ElseIf nxt2 <> "" And InStr(VWL, nxt2) = 0 Then
                        Else
                            m = m & "K"
                        End If
                        p = p + 1
                    ElseIf nxt = "N" Then
                        ' silent
                    ElseIf nxt = "I" Or nxt = "E" Or nxt = "Y" Then
                        m = m & "J"
                    Else
                        m = m & "K"
                    End If
                Case "H"
                    If InStr(VWL, prev) > 0 And InStr(VWL, nxt) = 0 Then
                    ElseIf prev = "C" Or prev = "S" Or prev = "P" Or prev = "T" Or prev = "G" Then
                    Else
                        m = m & "H"
                    End If
                Case "J": m = m & "J"
                Case "K"
                    If prev <> "C" Then m = m & "K"
                Case "L": m = m & "L"
                Case "M": m = m & "M"
                Case "N": m = m & "N"
                Case "P"
                    If nxt = "H" Then m = m & "F": p = p + 1 Else m = m & "P"
                Case "Q": m = m & "K"
                Case "R": m = m & "R"
                Case "S"
                    If nxt = "H" Then
                        m = m & "X": p = p + 1
                    ElseIf nxt = "I" And (nxt2 = "O" Or nxt2 = "A") Then
                        m = m & "X"
                    Else
                        m = m & "S"
                    End If
                Case "T"
                    If nxt = "H" Then
                        m = m & "0": p = p + 1
                    ElseIf nxt = "I" And (nxt2 = "O" Or nxt2 = "A") Then
                        m = m & "X"
                    Else
                        m = m & "T"
                    End If
                Case "V": m = m & "F"
                Case "W", "Y"
                    If InStr(VWL, nxt) > 0 Then m = m & curr
                Case "X": m = m & "KS"
                Case "Z": m = m & "S"
            End Select
            p = p + 1
        End If
    Loop
    Metaphone = m
End Function

' --- Convert a numeric string to a real number (else return the string) ---
Function NumIf(ByVal s As String) As Variant
    s = Trim(s)
    If Len(s) > 0 And Len(s) <= 15 And IsNumeric(s) Then
        NumIf = CDbl(s)
    Else
        NumIf = s
    End If
End Function

' --- Ordinal word for column headers (1->First, 2->Second, ...) ---
Function Ordinal(ByVal n As Long) As String
    Dim w As Variant
    w = Array("First", "Second", "Third", "Fourth", "Fifth", "Sixth", "Seventh", "Eighth", "Ninth", "Tenth", _
              "Eleventh", "Twelfth", "Thirteenth", "Fourteenth", "Fifteenth", "Sixteenth", "Seventeenth", _
              "Eighteenth", "Nineteenth", "Twentieth")
    If n >= 1 And n <= 20 Then
        Ordinal = w(n - 1)
    Else
        Ordinal = CStr(n) & "th"
    End If
End Function
