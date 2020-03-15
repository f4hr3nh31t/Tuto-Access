Attribute VB_Name = "PublicRunQuery"
Option Compare Database
Option Explicit

Public Function runQueryLogging(str As String)
    runQuery str
    Logging str
End Function

Public Function Logging(str As String)
    runQuery "INSERT INTO activity (AuthorId, Activity) VALUES(" & Nz(TempVars!UserId, 1) & ", """ & str & """);"
End Function

Public Function runQuery(str As String)
    DoCmd.SetWarnings False
    Debug.Print str
    DoCmd.RunSQL str
    DoCmd.SetWarnings True
End Function
