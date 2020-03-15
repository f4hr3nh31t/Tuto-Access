Attribute VB_Name = "Globals"
Option Compare Database
Option Explicit

Function IsNullOrEmpty(str As Variant) As Integer
    IsNullOrEmpty = (Nz(str, "") = "") 'same as IsNullOrEmpty = (Len(str & "") = 0)
    
End Function
