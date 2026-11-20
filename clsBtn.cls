VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBtn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================
'  CLASS MODULE  -  must be named exactly:  clsBtn
'  (Insert > Class Module, then in Properties set (Name) = clsBtn)
' ============================================================
Public WithEvents Btn As MSForms.CommandButton
Attribute Btn.VB_VarHelpID = -1
Public FormRef As Object
Public IsOK As Boolean

Private Sub Btn_Click()
    FormRef.HandleButton IsOK
End Sub

