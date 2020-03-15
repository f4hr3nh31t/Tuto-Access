VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Close()
    Logging "logoff"
    
    TempVars.RemoveAll
    
    closeAllAgentCard
End Sub

Private Sub Form_Load()
    If IsNull(TempVars!UserId) Then
        TempVars.Add "UserId", 1
        TempVars.Add "UserName", "SuperAdmin"
        TempVars.Add "UserGroup", "4"
    End If
    
    Logging "logon"
    
    Me.btModifications.Visible = TempVars!UserGroup > 2
    Me.btActivity.Visible = TempVars!UserGroup > 2
    Me.btUsers.Visible = TempVars!UserGroup > 2

End Sub
