VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLoginScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub pbConnexion_Click()
    Dim rs As Recordset
    
    Set rs = CurrentDb.OpenRecordset("USERS", dbOpenSnapshot, dbReadOnly)
    
    rs.FindFirst "Username = '" & Me.txtUsername & "'"
    
    If rs.NoMatch Then
        Me.lblUsername.Visible = True
        Me.txtUsername.SetFocus
        Me.txtUsername.SelStart = Nz(Len(Me.txtUsername), 0)
        Exit Sub
    End If
    
    If rs!Password <> Encrypt(Nz(Me.txtPassword, "")) Then
        Me.lblPassword.Visible = True
        Me.txtPassword.SetFocus
        Me.txtPassword.SelStart = Nz(Len(Me.txtPassword), 0)
        Exit Sub
    End If
    
    TempVars.Add "UserName", rs!Username.Value
    TempVars.Add "UserId", rs!Id.Value
    TempVars.Add "UserGroup", rs!Group.Value
    
    DoCmd.OpenForm "frmMainWindow"
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub txtPassword_Change()
    If Me.lblPassword.Visible Then Me.lblPassword.Visible = False
End Sub

Private Sub txtUsername_Change()
    If Me.lblUsername.Visible Then Me.lblUsername.Visible = False
End Sub
