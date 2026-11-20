Attribute VB_Name = "NameScreening"
Option Explicit

' ============================================================================
' AML Beneficiary Name Screening - Gibberish / Random Name Detector
' ----------------------------------------------------------------------------
' Purpose: Supports the beneficiary review by scoring every beneficiary
'          "full_name" entry and separating likely GIBBERISH / RANDOM names
'          from VALID ones, so the Screening Team can focus manual effort
'          where it's actually needed.
'
' IMPORTANT: This is a TRIAGE AID, not a final decision-maker. It flags
'            patterns (digits, symbols, keyboard-mash, no vowels, repeated
'            characters, known placeholder words). Some legitimate but
'            unusual names (e.g. certain Slavic, African, or transliterated
'            names) may land in the "REVIEW" band. A human must always make
'            the final call - never auto-reject purely on this score.
'
' IMPLEMENTATION NOTES:
'   - Run-time error 6 "Overflow" solved. In VBA, "A And B" evaluates BOTH
'     sides (no short-circuit), so "(vowelCount / letterCount)" was executed
'     even when letterCount = 0 -> division by zero -> Overflow. The vowel-ratio
'     test is now inside a nested If that only runs when letterCount > 0.
'   - All counters changed from Integer to Long to avoid overflow on long text.
'   - Ratio comparison done by cross-multiplication (vowelCount * 100 <
'     letterCount * 15) so no division is needed at all.
'   - Accented vowels are built with ChrW (see AccentedVowels), so the
'     letter lists stay correct on any Windows language / code page.
'
' HOW TO USE:
'   1. Open the workbook, go to the sheet with your beneficiary list.
'   2. Alt+F11, File > Import File, choose NameScreening.bas.
'   3. F5 (or run "ScreenBeneficiaryNames") with the target sheet active.
'   4. Two columns are added: Name_Status and Suspicion_Score.
'      A new sheet "Gibberish_Review" lists every flagged name for quick review.
'   5. Re-running is safe: it updates the same columns instead of duplicating.
'
' TUNE HERE:
'   - REVIEW_THRESHOLD / HIGH_RISK_THRESHOLD below: score cut-offs.
'   - exactJunkWords / substrJunkPatterns inside GetSuspicionScore: add any
'     placeholder text your team has actually seen in real submissions.
'   - NAME_COLUMN_HEADER: header text used to auto-find the name column.
' ============================================================================

Const REVIEW_THRESHOLD As Long = 2      ' score >= this -> REVIEW
Const HIGH_RISK_THRESHOLD As Long = 5   ' score >= this -> GIBBERISH - HIGH RISK
Const NAME_COLUMN_HEADER As String = "full_name"
Const DEFAULT_NAME_COLUMN As Long = 8   ' column H - used only if header not found
Const REVIEW_SHEET_NAME As String = "Gibberish_Review"


Sub ScreenBeneficiaryNames()

    Dim ws As Worksheet
    Dim reviewWs As Worksheet
    Dim lastRow As Long
    Dim nameCol As Long
    Dim outCol As Long
    Dim scoreCol As Long
    Dim statusCol As Long          ' existing system status column on the main sheet
    Dim i As Long
    Dim fullName As String
    Dim mainStatusVal As String    ' holds the pending/approved status
    Dim score As Long
    Dim status As String
    Dim reasons As String
    Dim reviewRow As Long
    Dim flaggedCount As Long

    Set ws = ActiveSheet

    ' 1) Find the name column automatically by header text; fall back to default
    nameCol = FindColumnByHeader(ws, NAME_COLUMN_HEADER)
    If nameCol = 0 Then nameCol = DEFAULT_NAME_COLUMN

    ' 1.5) Find the existing 'status' column on the main sheet
    statusCol = FindColumnByHeader(ws, "status")

    lastRow = ws.Cells(ws.Rows.count, nameCol).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data found under column " & nameCol & ".", vbExclamation
        Exit Sub
    End If

    ' 2) Result columns - reuse them if this macro already ran before (no duplicates)
    outCol = FindColumnByHeader(ws, "Name_Status")
    If outCol = 0 Then
        outCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column + 1
    End If
    scoreCol = outCol + 1

    ws.Cells(1, outCol).Value = "Name_Status"
    ws.Cells(1, scoreCol).Value = "Suspicion_Score"
    ws.Cells(1, outCol).Font.Bold = True
    ws.Cells(1, scoreCol).Font.Bold = True

    ' 3) Separate review sheet - actual separation of flagged names, not just coloring
    Application.DisplayAlerts = False
    On Error Resume Next
    Worksheets(REVIEW_SHEET_NAME).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set reviewWs = Worksheets.Add(After:=Worksheets(Worksheets.count))
    reviewWs.Name = REVIEW_SHEET_NAME

    reviewWs.Range("A1:E1").Value = Array("Source_Row", "Full_Name", "System_Status", "Suspicion_Score", "Reasons")
    reviewWs.Range("A1:E1").Font.Bold = True

    reviewRow = 2
    flaggedCount = 0

    For i = 2 To lastRow
        fullName = Trim(CStr(ws.Cells(i, nameCol).Value))
        reasons = ""

        ' Grab the status from the main sheet if the column exists
        If statusCol > 0 Then
            mainStatusVal = Trim(CStr(ws.Cells(i, statusCol).Value))
        Else
            mainStatusVal = "Not Found"
        End If

        If fullName = "" Then
            status = "EMPTY"
            score = 0
        Else
            score = GetSuspicionScore(fullName, reasons)
            If score >= HIGH_RISK_THRESHOLD Then
                status = "GIBBERISH - HIGH RISK"
            ElseIf score >= REVIEW_THRESHOLD Then
                status = "REVIEW - POSSIBLE GIBBERISH"
            Else
                status = "VALID"
            End If
        End If

        ws.Cells(i, outCol).Value = status
        ws.Cells(i, scoreCol).Value = score

        Select Case status
            Case "GIBBERISH - HIGH RISK"
                ws.Cells(i, outCol).Interior.Color = RGB(255, 199, 206)
            Case "REVIEW - POSSIBLE GIBBERISH"
                ws.Cells(i, outCol).Interior.Color = RGB(255, 235, 156)
            Case "VALID"
                ws.Cells(i, outCol).Interior.Color = RGB(198, 239, 206)
            Case Else
                ws.Cells(i, outCol).Interior.ColorIndex = xlColorIndexNone
        End Select

        If status = "GIBBERISH - HIGH RISK" Or status = "REVIEW - POSSIBLE GIBBERISH" Then
            reviewWs.Cells(reviewRow, 1).Value = i
            reviewWs.Cells(reviewRow, 2).Value = fullName
            reviewWs.Cells(reviewRow, 3).Value = mainStatusVal
            reviewWs.Cells(reviewRow, 4).Value = score
            reviewWs.Cells(reviewRow, 5).Value = reasons
            reviewRow = reviewRow + 1
            flaggedCount = flaggedCount + 1
        End If
    Next i

    reviewWs.Columns("A:E").AutoFit
    ws.Columns(outCol).AutoFit
    ws.Columns(scoreCol).AutoFit

    MsgBox flaggedCount & " of " & (lastRow - 1) & " name(s) flagged for review." & vbCrLf & _
           "Full list is on the '" & REVIEW_SHEET_NAME & "' sheet.", vbInformation, "Name Screening Complete"
End Sub


Function GetSuspicionScore(nameStr As String, ByRef reasons As String) As Long
    ' Accented Latin letters are built with ChrW so they survive any code page
    Dim VOWELS As String, EXTRA_LETTERS As String
    VOWELS = "aeiou" & AccentedVowels()
    EXTRA_LETTERS = AccentedVowels() & ChrW(241) & ChrW(231) & ChrW(353)   ' n-tilde, c-cedilla, s-caron

    Dim score As Long
    Dim cleanName As String
    Dim words() As String
    Dim w As String
    Dim wordIdx As Long
    Dim i As Long
    Dim ch As String
    Dim vowelCount As Long
    Dim letterCount As Long
    Dim consonantRun As Long
    Dim maxConsonantRun As Long
    Dim repeatRun As Long
    Dim maxRepeatRun As Long
    Dim exactJunkWords As Variant
    Dim substrJunkPatterns As Variant
    Dim j As Long

    score = 0
    reasons = ""
    cleanName = LCase(Trim(nameStr))
    words = Split(cleanName, " ")

    ' Digits inside a name field - strong signal
    If HasDigits(nameStr) Then
        score = score + 3
        reasons = reasons & "Contains digits; "
    End If

    ' Unexpected symbols (letters incl. accented ones, space, hyphen, apostrophe, period are OK)
    If HasInvalidSymbols(nameStr) Then
        score = score + 2
        reasons = reasons & "Contains unexpected symbols; "
    End If

    ' Overall length too short
    If Len(cleanName) < 3 Then
        score = score + 3
        reasons = reasons & "Too short overall; "
    End If

    ' Known placeholder/junk words - exact match per word (not substring), to avoid
    ' false hits on real names that merely contain these letters (e.g. "na" inside "Nadia")
    exactJunkWords = Array("test", "n/a", "na", "unknown", "none", "null", "xxx", _
                            "sample", "dummy", "noname", "nameless", "unavailable")
    For wordIdx = LBound(words) To UBound(words)
        For j = LBound(exactJunkWords) To UBound(exactJunkWords)
            If words(wordIdx) = exactJunkWords(j) Then
                score = score + 5
                reasons = reasons & "Word '" & words(wordIdx) & "' matches known placeholder; "
            End If
        Next j
    Next wordIdx

    ' Longer junk/keyboard-mash patterns - substring match is safe here (low chance
    ' of accidental overlap with a real name)
    substrJunkPatterns = Array("asdfgh", "qwerty", "asdasd", "qazwsx", "notavailable")
    For j = LBound(substrJunkPatterns) To UBound(substrJunkPatterns)
        If InStr(cleanName, substrJunkPatterns(j)) > 0 Then
            score = score + 5
            reasons = reasons & "Contains junk pattern '" & substrJunkPatterns(j) & "'; "
        End If
    Next j

    ' Same character repeated 3+ times in a row (e.g. "aaa", "kkkk")
    maxRepeatRun = 1
    repeatRun = 1
    For i = 2 To Len(cleanName)
        If Mid(cleanName, i, 1) = Mid(cleanName, i - 1, 1) And Mid(cleanName, i, 1) <> " " Then
            repeatRun = repeatRun + 1
            If repeatRun > maxRepeatRun Then maxRepeatRun = repeatRun
        Else
            repeatRun = 1
        End If
    Next i
    If maxRepeatRun >= 3 Then
        score = score + 4
        reasons = reasons & "Repeated character run (" & maxRepeatRun & "x); "
    End If

    ' Per word: vowel ratio + longest run of consecutive consonants
    For wordIdx = LBound(words) To UBound(words)
        w = words(wordIdx)
        If Len(w) >= 3 Then
            vowelCount = 0
            letterCount = 0
            consonantRun = 0
            maxConsonantRun = 0
            For i = 1 To Len(w)
                ch = Mid(w, i, 1)
                If ch Like "[a-z]" Or InStr(EXTRA_LETTERS, ch) > 0 Then
                    letterCount = letterCount + 1
                    If InStr(VOWELS, ch) > 0 Then
                        vowelCount = vowelCount + 1
                        consonantRun = 0
                    Else
                        consonantRun = consonantRun + 1
                        If consonantRun > maxConsonantRun Then maxConsonantRun = consonantRun
                    End If
                End If
            Next i

            ' *** FIX: only evaluate the vowel ratio when letterCount > 0 ***
            ' (nested If avoids the division-by-zero Overflow; ratio checked by
            '  cross-multiplication so no division is performed at all)
            If letterCount > 0 Then
                If vowelCount = 0 Then
                    score = score + 3
                    reasons = reasons & "'" & w & "' has no vowels; "
                ElseIf letterCount >= 4 Then
                    If vowelCount * 100 < letterCount * 15 Then   ' same as ratio < 0.15
                        score = score + 2
                        reasons = reasons & "'" & w & "' very low vowel ratio; "
                    End If
                End If
            End If

            If maxConsonantRun >= 5 Then
                score = score + 2
                reasons = reasons & "'" & w & "' has long consonant run; "
            End If
        End If
    Next wordIdx

    GetSuspicionScore = score
End Function


Function HasDigits(s As String) As Boolean
    Dim i As Long
    For i = 1 To Len(s)
        If Mid(s, i, 1) Like "#" Then
            HasDigits = True
            Exit Function
        End If
    Next i
    HasDigits = False
End Function


Function HasInvalidSymbols(s As String) As Boolean
    Const BAD_CHARS As String = "!@#$%^&*()_+={}[]|\:;""<>,?/~`"
    Dim i As Long
    For i = 1 To Len(s)
        If InStr(BAD_CHARS, Mid(s, i, 1)) > 0 Then
            HasInvalidSymbols = True
            Exit Function
        End If
    Next i
    HasInvalidSymbols = False
End Function


Function FindColumnByHeader(ws As Worksheet, headerText As String) As Long
    Dim lastCol As Long
    Dim c As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        If LCase(Trim(CStr(ws.Cells(1, c).Value))) = LCase(headerText) Then
            FindColumnByHeader = c
            Exit Function
        End If
    Next c
    FindColumnByHeader = 0
End Function


' Lower-case accented vowels (grave, acute, circumflex, tilde, diaeresis, ring)
Private Function AccentedVowels() As String
    Dim codes As Variant, i As Long, s As String
    codes = Array(224, 225, 226, 227, 228, 229, 232, 233, 234, 235, 236, 237, 238, 239, _
                  242, 243, 244, 245, 246, 249, 250, 251, 252, 253, 255)
    For i = LBound(codes) To UBound(codes)
        s = s & ChrW(codes(i))
    Next i
    AccentedVowels = s
End Function
