Attribute VB_Name = "PublicActualizefrm"
Option Compare Database
Option Explicit

Function actualize(frm As String)
    If CurrentProject.AllForms(frm).IsLoaded Then
        Forms(frm).Requery
        Forms(frm).Refresh
    End If
End Function
