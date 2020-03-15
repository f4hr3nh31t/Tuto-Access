Attribute VB_Name = "PublicAgentCard"
Option Compare Database
Option Explicit

Public clnAgentCard As New Collection   'Instances of frmAgentCard

Function openAgentCard(Optional ByVal AgentId As Integer = 0)
    'Purpose : Open an independent instance of form frmAgentCard
    Dim frm As Form
    
    'Open a new instance, show it and set a caption
    Set frm = New Form_frmAgentCard
    frm.Visible = True
    frm.Caption = frm.Hwnd & ", opened " & Now()

    If TempVars!UserGroup > 0 Then
        frm.RecordsetType = 0
    Else
        frm.RecordsetType = 2
    End If
    
    frm.AllowAdditions = TempVars!UserGroup > 0
    frm.pbClearAgent.Enabled = TempVars!UserGroup > 0
    frm.CtlTab.Pages(1).Visible = TempVars!UserGroup > 0
    frm.pbNewAgent.Enabled = TempVars!UserGroup > 0
    frm.pbDupplicateAgent.Enabled = TempVars!UserGroup > 0
    frm.pbDeleteAgent.Enabled = TempVars!UserGroup > 0

    frm.CtlTab.Pages(2).Visible = TempVars!UserGroup > 1
    frm.CtlTab.Pages(3).Visible = TempVars!UserGroup > 1

    frm.pbEffacerVersionsPrecedentes.Enabled = TempVars!UserGroup > 2
    frm.pbEffacerVersionsSuivantes.Enabled = TempVars!UserGroup > 2
    
    If AgentId = 0 Then
        DoCmd.GoToRecord , , acNewRec
    Else
        'Find record
        frm.Recordset.FindFirst "[AGENTS_Id] = " & str(AgentId)
        If frm.Recordset.NoMatch Then MsgBox "Can't find the record"
    End If
    
    'Append the new instance to our collection
    clnAgentCard.Add Item:=frm, Key:=CStr(frm.Hwnd)
    
    Set frm = Nothing
End Function

Function closeAllAgentCard()
    'Purpose : Close all instances in the clnAgentCard collection
    
    While clnAgentCard.Count > 0
        clnAgentCard.Remove 1
    Wend
End Function
