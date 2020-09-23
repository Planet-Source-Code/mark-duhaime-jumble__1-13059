Attribute VB_Name = "Common"
Public Sub DBCloseAll()
    Dim WSTemp As Workspace
    Dim DBTemp As Database
    Dim RSTemp As Recordset
    '
    On Error Resume Next
    For Each WSTemp In Workspaces
        For Each DBTemp In WSTemp.Databases
            For Each RSTemp In DBTemp.Recordsets
                RSTemp.Close
            Next
            DBTemp.Close
        Next
    Next
    '
End Sub
