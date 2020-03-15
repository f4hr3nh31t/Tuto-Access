VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMainWindow_Versions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Détail_DblClick(Cancel As Integer)
    openAgentCard Me.AGENTS_Id
End Sub

Private Sub pbEffacerAnciennesVerions_Click()
    If MsgBox("Effacer toutes les anciennes versions ? Cette action est irréversible.", vbYesNo) = vbYes Then
        runQueryLogging " DELETE FROM AGENTS WHERE LastVersion = 0 "
        actualize "frmMainWindow"
    End If
End Sub

Private Sub pbOuvrirAgent_Click()
    openAgentCard Me.AGENTS_Id
End Sub

Private Sub pbOuvrirUser_Click()
    openUserCard Me.USERS_Id
End Sub
