VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "UserForm1"
   ClientHeight    =   2424
   ClientLeft      =   60
   ClientTop       =   276
   ClientWidth     =   2928
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================
'  USERFORM  -  must be named exactly:  frmSettings
'  (Insert > UserForm, leave it blank, set (Name) = frmSettings,
'   then paste ALL of this into the form's code window.
'   Controls build themselves automatically - no manual drawing.)
' ============================================================
Private mHandlers As Collection
Private txtCountries As MSForms.TextBox
Private cboRecMetric As MSForms.ComboBox
Private txtRedThresh As MSForms.TextBox
Private chkAgent As MSForms.CheckBox
Private cboAgtMetric As MSForms.ComboBox
Private txtAgtThresh As MSForms.TextBox
Private chkMeta As MSForms.CheckBox
Private cboNoteScope As MSForms.ComboBox

' ---- values the macro reads after the form closes ----
Public Cancelled As Boolean
Public countries As String
Public useUnique As Boolean
Public redThresh As Long
Public useAgent As Boolean
Public agtUseUnique As Boolean
Public agtThresh As Long
Public UseMetaphone As Boolean
Public noteAllTrx As Boolean

Private Sub UserForm_Initialize()
    Set mHandlers = New Collection
    Me.Caption = "WU Report Analyzer - Settings"
    Me.Width = 360
    Me.Height = 355

    AddLabel "Countries (comma separated, or 'All Countries'):", 12, 8
    Set txtCountries = Me.Controls.Add("Forms.TextBox.1", "txtCountries")
    PlaceCtl txtCountries, 12, 24, 332, 18
    txtCountries.Text = "All Countries"

    AddLabel "Receiver metric:", 12, 50
    Set cboRecMetric = Me.Controls.Add("Forms.ComboBox.1", "cboRecMetric")
    PlaceCtl cboRecMetric, 170, 48, 174, 18
    cboRecMetric.AddItem "Unique Senders"
    cboRecMetric.AddItem "Total Transactions"
    cboRecMetric.Style = fmStyleDropDownList
    cboRecMetric.ListIndex = 0

    AddLabel "Receiver RED threshold (>=):", 12, 76
    Set txtRedThresh = Me.Controls.Add("Forms.TextBox.1", "txtRedThresh")
    PlaceCtl txtRedThresh, 220, 74, 60, 18
    txtRedThresh.Text = "10"

    Set chkAgent = Me.Controls.Add("Forms.CheckBox.1", "chkAgent")
    PlaceCtl chkAgent, 12, 100, 332, 18
    chkAgent.Caption = "Include Agent Cash-In analysis (needs Sheet2)"
    chkAgent.Value = False

    AddLabel "Agent metric:", 12, 124
    Set cboAgtMetric = Me.Controls.Add("Forms.ComboBox.1", "cboAgtMetric")
    PlaceCtl cboAgtMetric, 170, 122, 174, 18
    cboAgtMetric.AddItem "Unique Customers"
    cboAgtMetric.AddItem "Total Transactions"
    cboAgtMetric.Style = fmStyleDropDownList
    cboAgtMetric.ListIndex = 0

    AddLabel "Agent threshold (>=):", 12, 148
    Set txtAgtThresh = Me.Controls.Add("Forms.TextBox.1", "txtAgtThresh")
    PlaceCtl txtAgtThresh, 220, 146, 60, 18
    txtAgtThresh.Text = "5"

    Set chkMeta = Me.Controls.Add("Forms.CheckBox.1", "chkMeta")
    PlaceCtl chkMeta, 12, 172, 332, 18
    chkMeta.Caption = "Generate Metaphone + Name-Sound Many-to-One table"
    chkMeta.Value = True

    AddLabel "Note shows:", 12, 198
    Set cboNoteScope = Me.Controls.Add("Forms.ComboBox.1", "cboNoteScope")
    PlaceCtl cboNoteScope, 120, 196, 224, 18
    cboNoteScope.AddItem "Only suspicious transactions"
    cboNoteScope.AddItem "All transactions in the data"
    cboNoteScope.Style = fmStyleDropDownList
    cboNoteScope.ListIndex = 0

    Dim btnOK As MSForms.CommandButton, btnCancel As MSForms.CommandButton
    Set btnOK = Me.Controls.Add("Forms.CommandButton.1", "btnOK")
    PlaceCtl btnOK, 80, 235, 90, 28
    btnOK.Caption = "Run Analysis"
    btnOK.Default = True
    WireButton btnOK, True

    Set btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel")
    PlaceCtl btnCancel, 190, 235, 90, 28
    btnCancel.Caption = "Cancel"
    btnCancel.Cancel = True
    WireButton btnCancel, False
End Sub

Private Sub AddLabel(ByVal txt As String, ByVal x As Single, ByVal y As Single)
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1")
    lbl.Caption = txt
    lbl.Left = x: lbl.Top = y: lbl.Width = 340: lbl.Height = 14
End Sub

Private Sub PlaceCtl(ByVal ctl As Object, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single)
    ctl.Left = x: ctl.Top = y: ctl.Width = w: ctl.Height = h
End Sub

Private Sub WireButton(ByVal b As MSForms.CommandButton, ByVal okFlag As Boolean)
    Dim h As clsBtn
    Set h = New clsBtn
    Set h.Btn = b
    Set h.FormRef = Me
    h.IsOK = okFlag
    mHandlers.Add h
End Sub

' Called by clsBtn when a button is clicked
Public Sub HandleButton(ByVal okPressed As Boolean)
    If Not okPressed Then
        Cancelled = True
        Me.Hide
        Exit Sub
    End If
    If Trim(txtCountries.Text) = "" Then
        MsgBox "Please enter at least one country (or 'All Countries').", vbExclamation
        Exit Sub
    End If
    countries = txtCountries.Text
    useUnique = (cboRecMetric.ListIndex = 0)
    If IsNumeric(txtRedThresh.Text) Then redThresh = CLng(txtRedThresh.Text) Else redThresh = 10
    If redThresh < 2 Then redThresh = 2
    useAgent = (chkAgent.Value = True)
    agtUseUnique = (cboAgtMetric.ListIndex = 0)
    If IsNumeric(txtAgtThresh.Text) Then agtThresh = CLng(txtAgtThresh.Text) Else agtThresh = 5
    If agtThresh < 1 Then agtThresh = 1
    UseMetaphone = (chkMeta.Value = True)
    noteAllTrx = (cboNoteScope.ListIndex = 1)
    Cancelled = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancelled = True
        Me.Hide
        Cancel = True   ' keep loaded so the macro can read the values, then it unloads
    End If
End Sub

