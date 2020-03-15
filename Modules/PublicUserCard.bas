Attribute VB_Name = "PublicUserCard"
Option Compare Database
Option Explicit

Public clnUserCard As New Collection   'Instances of frmUserCard

Function openUserCard(Optional ByVal UserId As Integer = 0)
    'Purpose : Open an independent instance of form frmUserCard
    Dim frm As Form
    
    'Open a new instance, show it and set a caption
    Set frm = New Form_frmUserCard
    frm.Visible = True
    frm.Caption = frm.Hwnd & ", opened " & Now()
    
    
    
    If UserId = 0 Then
        DoCmd.GoToRecord , , acNewRec
    Else
        'Find record
        frm.Recordset.FindFirst "Id = " & str(UserId)
        If frm.Recordset.NoMatch Then MsgBox "Can't find the record"
    End If
    
    'Append the new instance to our collection
    clnAgentCard.Add Item:=frm, Key:=CStr(frm.Hwnd)
    
    Set frm = Nothing
End Function

Function closeAllUserCard()
    'Purpose : Close all instances in the clnUserCard collection
    
    While clnUserCard.Count > 0
        clnUserCard.Remove 1
    Wend
End Function
