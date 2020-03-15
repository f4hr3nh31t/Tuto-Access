VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmUserCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Close()
    'Purpose : Remove this instance from clnUserCard collection
    Dim obj As Object   'Object in clnUserCard
    
    Dim blnRemove As Boolean    'Flag to remove it
    
    'Check if this instance is in the collection
    For Each obj In clnUserCard
        If obj.Hwnd = Me.Hwnd Then
            blnRemove = True
            Exit For
        End If
    Next
    
    'Deassign the object and remove from collection
    Set obj = Nothing
    If blnRemove Then
        clnUserCard.Remove CStr(Me.Hwnd)
    End If
End Sub

Private Sub pbEffacerPassword_Click()
    Me.Password = ""
End Sub

Private Sub pbOk_Click()
    DoCmd.Close
End Sub
