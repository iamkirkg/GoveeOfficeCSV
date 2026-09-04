Attribute VB_Name = "SourceControl"
Option Explicit

Public Sub ExportSourceModules()
    Dim vbComp As Object
    Dim exportFolder As String
    Dim filePath As String
    Dim compName As String
    Dim ext As String
    
    exportFolder = "e:\Bitbucket\GoveeOfficeCSV"
    
    If Dir(exportFolder, vbDirectory) = "" Then
        MkDir exportFolder
    End If
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        compName = vbComp.Name
        
        Select Case vbComp.Type
            Case 1   ' Standard module
                ext = ".bas"
            Case 2   ' Class module
                ext = ".cls"
            Case 3   ' UserForm
                ext = ".frm"
            Case 100 ' Document module: ThisWorkbook, Sheet1, etc.
                ext = ".cls"
            Case Else
                ext = ""
        End Select
        
        If ext <> "" Then
            filePath = exportFolder & "\" & compName & ext
            
            On Error Resume Next
            Kill filePath
            On Error GoTo 0
            
            vbComp.Export filePath
        End If
    Next vbComp
End Sub

