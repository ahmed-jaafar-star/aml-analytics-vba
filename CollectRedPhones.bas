Attribute VB_Name = "CollectRedPhones"
Option Explicit

'==============================================================================
' COLLECT RED PHONES
'------------------------------------------------------------------------------
' Gathers every MSISDN marked in red or pink (column M) on the country sheets
' into one "Red_Phones" sheet, with its destination country (column R).
' - Skips the "Analysis - ..." sheets and the output sheet itself
' - Reads DisplayFormat, so conditional-formatting colors count too
' - IsRedOrPink accepts any red/pink shade, not one exact color code
'==============================================================================

Sub Collect_Red_Phones_Correctly()

    Dim wb As Workbook
    Dim ws As Worksheet, outWs As Worksheet
    Dim hdr As Range, c As Range
    Dim lastRow As Long, r As Long, outRow As Long
    Dim phone As String, country As String
    Dim scanned As Long, total As Long
    Dim fc As Long, ic As Long
    Dim isRed As Boolean

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Set wb = ActiveWorkbook

    On Error Resume Next
    Set outWs = wb.Worksheets("Red_Phones")
    On Error GoTo ErrHandler

    If outWs Is Nothing Then
        Set outWs = wb.Worksheets.Add( _
            After:=wb.Worksheets(wb.Worksheets.count))
        outWs.Name = "Red_Phones"
    Else
        outWs.Cells.Clear
    End If

    outWs.Range("A1:B1").Value = _
        Array("DESTINATION_COUNTRY", "MSISDN")
    outWs.Range("A1:B1").Font.Bold = True
    outWs.Columns("B").NumberFormat = "@"

    outRow = 2
    scanned = 0
    total = 0

    For Each ws In wb.Worksheets

        If ws.Name <> outWs.Name _
           And LCase(Left(Trim(ws.Name), 8)) <> "analysis" Then

            Set hdr = Nothing

            Set hdr = ws.Cells.Find( _
                What:="MSISDN", _
                LookIn:=xlValues, _
                LookAt:=xlWhole, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlNext, _
                MatchCase:=False, _
                SearchFormat:=False)

            If Not hdr Is Nothing Then

                scanned = scanned + 1

                lastRow = ws.Cells(ws.Rows.count, "M") _
                            .End(xlUp).Row

                For r = hdr.Row + 1 To lastRow

                    Set c = ws.Cells(r, "M")

                    If Not IsError(c.Value2) Then

                        If IsNumeric(c.Value2) And _
                           Len(Trim(CStr(c.Value2))) > 0 Then

                            If IsNumeric(c.Value2) Then
                                phone = Format$(c.Value2, "0")
                            End If

                            If Len(phone) >= 10 And _
                               Len(phone) <= 15 And _
                               phone Like String(Len(phone), "#") Then

                                fc = c.DisplayFormat.Font.Color
                                ic = c.DisplayFormat.Interior.Color

                                isRed = IsRedOrPink(fc) Or _
                                        IsRedOrPink(ic)

                                If isRed Then

                                    country = Trim(CStr( _
                                        ws.Cells(r, "R").Value2))

                                    If Len(country) > 0 Then
                                        outWs.Cells(outRow, "A").Value = country
                                        outWs.Cells(outRow, "B").Value = phone
                                        outRow = outRow + 1
                                        total = total + 1
                                    End If

                                End If
                            End If
                        End If
                    End If

                    phone = vbNullString

                Next r
            End If
        End If
    Next ws

    outWs.Columns("A:B").AutoFit
    Application.ScreenUpdating = True

    MsgBox "Country sheets scanned: " & scanned & vbCrLf & _
           "Valid red phones collected: " & total, vbInformation

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbExclamation

End Sub

Private Function IsRedOrPink(ByVal clr As Long) As Boolean

    Dim rr As Long, gg As Long, bb As Long

    rr = clr Mod 256
    gg = (clr \ 256) Mod 256
    bb = (clr \ 65536) Mod 256

    IsRedOrPink = (rr >= 180 And gg <= 200 And _
                   bb <= 200 And rr > gg * 1.1)

End Function
