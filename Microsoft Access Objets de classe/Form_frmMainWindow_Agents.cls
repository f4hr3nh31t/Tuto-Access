VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMainWindow_Agents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private SpaceBarPress As Boolean

Private Sub Détail_DblClick(Cancel As Integer)
    openAgentCard Me.AGENTS_Id
End Sub

Private Sub lblRecherche_Click()
    Me.txtSearch.SetFocus
End Sub

Private Sub lstDepartement_AfterUpdate()
    updateFilter Nz(Me.lstDepartement, ""), Nz(Me.txtSearch, "")
End Sub

Private Sub pbEraseFilter_Click()
    'On efface le texte de la barre de recherche
    Me.txtSearch = vbNullString
    updateFilter Nz(Me.lstDepartement, ""), Nz(Me.txtSearch, "")
    
    Me.lblRecherche.Visible = True
    Me.pbEraseFilter.Enabled = False
End Sub

Private Sub pbNouvelAgent_Click()
    openAgentCard
End Sub

Private Sub pbOuvrirAgent_Click()
    openAgentCard Me.AGENTS_Id
End Sub

Private Sub txtSearch_Change()
    
    'La propriete TEXT ne peut etre appelee sur un recordset vide si AllowAdditions est faux
    Me.AllowAdditions = True
    
    'Access removes trailing spaces when requerying
    'On ajoute un espace lorsque cela est necessaire
    If SpaceBarPress Then
        Me.txtSearch = Me.txtSearch & " "
        'On reset l'indicateur
        SpaceBarPress = False
    End If
    
    'La propriété TEXT a besoin du focus pour etre appelee
    Me.txtSearch.SetFocus
    updateFilter Nz(Me.lstDepartement, ""), Me.txtSearch.Text
    
    'On replace le curseur dans la barre de recherche
    Me.txtSearch.SetFocus
    If Not IsNullOrEmpty(Me.txtSearch.Text) Then
        'On replace le curseur a la fin de la barre de recherche
        Me.txtSearch.SelStart = Len(Me.txtSearch.Text)
    End If
    
    'On desactive AllowAdditions
    Me.AllowAdditions = False
    
End Sub

Private Sub txtSearch_GotFocus()
    Me.pbEraseFilter.Enabled = True
    Me.lblRecherche.Visible = False
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 Then SpaceBarPress = True
End Sub

Private Sub txtSearch_LostFocus()
    If IsNullOrEmpty(txtSearch.Value) Then
        Me.pbEraseFilter.Enabled = False
        Me.lblRecherche.Visible = True
    End If
End Sub

Public Function updateFilter(Department As String, LastName As String)
    Dim filter As String
    
    Me.filter = ""
    Me.FilterOn = False
    
    If Not IsNullOrEmpty(Department) Or Not IsNullOrEmpty(LastName) Then
        filter = ""
        If Not IsNullOrEmpty(Department) Then filter = "[Department] LIKE '" & Department & "*'"
        If Not IsNullOrEmpty(Department) And Not IsNullOrEmpty(LastName) Then filter = filter & " AND "
        If Not IsNullOrEmpty(LastName) Then filter = filter & "[LastName] LIKE '" & LastName & "*'"
        
        Me.filter = filter
        Me.FilterOn = True
    End If

End Function

