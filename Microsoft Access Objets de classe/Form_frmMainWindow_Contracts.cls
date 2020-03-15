VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMainWindow_Contracts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub lstType_AfterUpdate()
    Me.FilterOn = False
    Me.filter = "[Type] LIKE '" & Me.lstType & "*'"
    Me.FilterOn = True
    
End Sub

Private Sub pbOuvrirAgent_Click()
    openAgentCard Me.AGENTS_Id
End Sub
