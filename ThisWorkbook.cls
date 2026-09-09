VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error GoTo SafeExit
    ExportSourceModules
SafeExit:
End Sub

Private Sub Workbook_Open()
    On Error GoTo CleanFail

    RunGoveeImport

    If Len(Dir(GetGoveeFolder() & "\debug.txt")) > 0 Then
        Application.VBE.MainWindow.Visible = True
    Else
        Application.DisplayAlerts = False
        Application.Quit
    End If

    Exit Sub

CleanFail:
    Application.DisplayAlerts = True
    MsgBox "Workbook_Open failed: " & Err.Description, vbExclamation
End Sub

