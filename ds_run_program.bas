Attribute VB_Name = "ds_run_program"
Public Sub ds_run_program() 'Run DraftSight App if not open yet
    Dim x As Variant
    Dim Path As String
    
    Path = "C:\Program Files\Dassault Systemes\DraftSight\bin\DraftSight.exe"
         
    Dim dsApp As DraftSight.Application
    On Error Resume Next
    Set dsApp = GetObject(, "DraftSight.Application")
    
    If dsApp Is Nothing Then
        x = Shell(Path, vbNormalFocus)
    End If
 
End Sub

