VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMainWindow_Acount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdModifierPass_Click()
    Dim rs As Recordset
    
    Set rs = CurrentDb.OpenRecordset("Users", dbOpenSnapshot, dbReadOnly)
    
    rs.FindFirst "Id = " & Me.Id
    
    If rs.NoMatch Then
        MsgBox ("Utilisateur non trouvé")
        Exit Sub
    End If
    
    If rs!Password <> Encrypt(Nz(Me.txtOldPass, "")) Then
        Me.lblOldPass.Visible = True
        Me.txtOldPass.SetFocus
        If Not IsNull(Me.txtOldPass) Then Me.txtOldPass.SelStart = Len(Me.txtOldPass)
        Exit Sub
    End If
    If IsNullOrEmpty(Me.txtPass1) Then
        Me.lblPass1.Visible = True
        Me.txtPass1.SetFocus
        Exit Sub
    End If
    If IsNullOrEmpty(Me.txtPass2) Then
        Me.lblPass2.Visible = True
        Me.txtPass2.SetFocus
        Exit Sub
    End If
    If Me.txtPass1 <> Me.txtPass2 Then
        Me.lblPass2.Visible = True
        Me.txtPass2.SetFocus
        Me.txtPass2.SelStart = Len(Me.txtPass2)
        Exit Sub
    End If
    
    runQuery "UPDATE USERS SET Password = '" & Encrypt(Me.txtPass2) & "' WHERE Id = " & Me.Id & ";"
    
    Me.txtOldPass = ""
    Me.txtPass1 = ""
    Me.txtPass2 = ""
    Me.lblmaj.Visible = True
End Sub

Private Sub Form_Load()
    Me.Recordset.FindFirst "Id = " & TempVars!UserId
    
    Me.Group.Enabled = TempVars!UserGroup > 2

End Sub

Private Sub pbDeco_Click()
    DoCmd.Close
End Sub
