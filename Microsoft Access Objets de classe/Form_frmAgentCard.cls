VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAgentCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function saveChanges()
    Dim newAgentVersion As Integer, newAgentNum As Integer
    Dim strSql As String
    
    newAgentNum = Nz(Me.AgentNum, 0)
    newAgentVersion = Nz(DMax("SELECT MAX(AgentVersion) FROM AGENTS WHERE AgentNum = " & newAgentNum, "AGENTS"), 0) + 1
    If newAgentNum = 0 Then newAgentNum = DMax("SELECT MAX(AgentNum) FROM AGENTS", "AGENTS") + 1
    
    strSql = "INSERT INTO AGENTS ( AgentNum, AgentVersion, FirstName, LastName, Function, Department, AuthorId ) "
    strSql = strSql & "VALUES ( "
    strSql = strSql & newAgentNum & ", "
    strSql = strSql & newAgentVersion & ", '"
    strSql = strSql & Me.FirstName & "', '"
    strSql = strSql & Me.LastName & "', '"
    strSql = strSql & Me.Function & "', '"
    strSql = strSql & Me.Department & "', "
    strSql = strSql & Nz(TempVars!UserId, 1) & " );"

    runQueryLogging strSql
    runQueryLogging "UPDATE AGENTS SET LastVersion = 0 WHERE AgentNum = " & newAgentNum & " AND AgentVersion <> " & newAgentVersion & ";"
    
    Me.Undo
    
    Me.Requery
    Me.Recordset.FindFirst "AgentNum = " & newAgentNum & " AND AgentVersion = " & newAgentVersion
    actualize "frmMainWindow"
    
    Me.pbAppliquer.Enabled = False
End Function

Private Sub Form_BeforeUpdate(Cancel As Integer)
    If Me.Dirty Then
        If MsgBox("Save changes ?", 4) = vbYes Then
            saveChanges
        Else
            Me.Undo
        End If
    End If
End Sub

Private Sub Form_Close()
    'Purpose : Remove this instance from clnAgentCard collection
    Dim obj As Object   'Object in clnAgentCard
    
    Dim blnRemove As Boolean    'Flag to remove it
    
    'Check if this instance is in the collection
    For Each obj In clnAgentCard
        If obj.Hwnd = Me.Hwnd Then
            blnRemove = True
            Exit For
        End If
    Next
    
    'Deassign the object and remove from collection
    Set obj = Nothing
    If blnRemove Then
        clnAgentCard.Remove CStr(Me.Hwnd)
    End If
End Sub

Private Sub Form_Dirty(Cancel As Integer)
    Me.pbAppliquer.Enabled = True
    
End Sub

Private Sub pbAnnuler_Click()
    If Me.Dirty Then Me.Undo
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub pbAppliquer_Click()
    If Me.Dirty Then saveChanges
End Sub

Private Sub pbClearAgent_Click()
    Me.FirstName = " "
    Me.LastName = " "
    Me.Function = Null
    Me.Department = ""
End Sub

Private Sub pbDeleteAgent_Click()
    If MsgBox("Supprimer cet Agent ?", vbYesNo) = vbYes Then
        runQueryLogging "UPDATE AGENTS SET LastVersion = 0 WHERE AgentNum = " & Me.AgentNum & ";"
        
        Me.Requery
        actualize "frmMainWindow"
    End If
End Sub

Private Sub pbEffacerVersionsPrecedentes_Click()
    If MsgBox("Effacer toutes les anciennes versions ? Cette action est irréversible.", vbYesNo) = vbYes Then
        Dim AgentNum As String, AgentVersion As String
        AgentNum = Me.AgentNum
        AgentVersion = Me.subfrmAgentVersions("AgentVersion")
        
        runQueryLogging " DELETE FROM AGENTS WHERE AgentNum = " & Me.AgentNum & " AND Me.AgentVersion < " & Me.subfrmAgentVersions("AgentVersion") & ";"
        
        Me.Recordset.FindFirst "AgentNum = " & AgentNum & " AND LastVersion = 1"
        Me.subfrmAgentVersions.Requery
        actualize "frmMainWindow"
    End If
End Sub

Private Sub pbEffacerVersionsSuivantes_Click()
    If MsgBox("Effacer toutes les versions ultérieures ? Cette action est irréversible.", vbYesNo) = vbYes Then
        Dim AgentNum As String, AgentVersion As String
        AgentNum = Me.AgentNum
        AgentVersion = Me.subfrmAgentVersions("AgentVersion")
        
        runQueryLogging " DELETE FROM AGENTS WHERE AgentNum = " & AgentNum & " AND AgentVersion > " & Me.subfrmAgentVersions("AgentVersion") & ";"
        runQueryLogging "UPDATE AGENTS SET LastVersion = 0 WHERE AgentNum = " & AgentNum & " AND AgentVersion <> " & AgentVersion & ";"
        runQueryLogging "UPDATE AGENTS SET LastVersion = 1 WHERE AgentNum = " & AgentNum & " AND AgentVersion = " & AgentVersion & ";"

        Me.Recordset.FindFirst "AgentNum = " & AgentNum & " AND LastVersion = 1"
        Me.subfrmAgentVersions.Requery
        actualize "frmMainWindow"
    End If
End Sub

Private Sub pbNewAgent_Click()
    openAgentCard
End Sub

Private Sub pbOk_Click()
    If Me.Dirty Then saveChanges
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub pbRestaurer_Click()
    Dim strSql As String
    Dim AgentNum As String, AgentVersion As String, newAgentVersion As String
    Dim AuthorId As Integer
    
    AuthorId = Nz(TempVars!UserId, 1)
    
    AgentNum = Me.AgentNum
    AgentVersion = Me.subfrmAgentVersions("AgentVersion")
    newAgentVersion = Nz(DMax("SELECT MAX(AgentVersion) FROM AGENTS WHERE AgentNum = " & AgentNum, "AGENTS"), 0) + 1
    
    strSql = "INSERT INTO AGENTS ( AgentNum, AgentVersion, FirstName, LastName, Function, Department, AuthorId ) "
    strSql = strSql & "SELECT AgentNum, " & newAgentVersion & ", FirstName, LastName, Function, Department, " & AuthorId & " "
    strSql = strSql & "FROM AGENTS "
    strSql = strSql & "WHERE AgentNum = " & AgentNum & " AND AgentVersion = " & AgentVersion & ";"
    
    runQueryLogging strSql
    runQueryLogging "UPDATE AGENTS SET LastVersion = 0 WHERE AgentNum = " & AgentNum & " AND AgentVersion <> " & newAgentVersion & ";"

    Me.Recordset.FindFirst "AgentNum = " & AgentNum & " AND LastVersion = 1"
    Me.subfrmAgentVersions.Requery
    actualize "frmMainWindow"
End Sub
