Attribute VB_Name = "Govee"
Sub HelloGovee()
    MsgBox "Module is working"
End Sub

Option Explicit

' Public gLogPath As String

Public Sub RunGoveeImport()
    Dim folderPath As String
    Dim dateToken As String
    Dim gLogPath As String
    Dim fileName As String
    Dim fullPath As String
    Dim sensorNum As Long
    Dim matchedCount As Long
    
    On Error GoTo FatalError
    
    folderPath = GetGoveeFolder()
    dateToken = Format(Date - 1, "yyyymmdd")
    gLogPath = folderPath & "\Govee" & dateToken & ".txt"
    
    InitializeLog gLogPath
    WriteLogLine gLogPath, "START"
    WriteLogLine gLogPath, "Folder: " & folderPath
    WriteLogLine gLogPath, "Date token: " & dateToken
    
    fileName = Dir(folderPath & "\*" & dateToken & "*.csv")
    
    If fileName = "" Then
        WriteLogLine gLogPath, "No CSV files found."
    End If
    
    Do While fileName <> ""
        fullPath = folderPath & "\" & fileName
        WriteLogLine gLogPath, "Found: " & fileName
        
        If TryMatchGoveeFile(fileName, dateToken, sensorNum) Then
            matchedCount = matchedCount + 1
            WriteLogLine gLogPath, "Accepted: sensor " & Format(sensorNum, "00") & " -> " & fileName
        Else
            WriteLogLine gLogPath, "Skipped: " & fileName
        End If
        
        fileName = Dir()
    Loop
    
    WriteLogLine gLogPath, "Matched files: " & matchedCount
    WriteLogLine gLogPath, "END"
    
    MsgBox "Scan complete. Log written to:" & vbCrLf & gLogPath, vbInformation
    
    Exit Sub

FatalError:
    WriteLogLine gLogPath, "FATAL ERROR: " & Err.Number & " - " & Err.Description
    MsgBox "Error: " & Err.Description, vbExclamation
End Sub

Private Function GetGoveeFolder() As String
    Dim oneDrivePath As String
    Dim folderPath As String
    
    oneDrivePath = Environ$("OneDrive")
    
    If Len(oneDrivePath) = 0 Then
        Err.Raise vbObjectError + 1000, , "OneDrive environment variable not found."
    End If
    
    folderPath = oneDrivePath & "\Govee"
    
    If Dir(folderPath, vbDirectory) = "" Then
        Err.Raise vbObjectError + 1001, , "Folder not found: " & folderPath
    End If
    
    GetGoveeFolder = folderPath
End Function

Private Function TryMatchGoveeFile(ByVal fileName As String, ByVal dateToken As String, ByRef sensorNum As Long) As Boolean
    Dim re As Object
    Dim matches As Object
    Dim sensorText As String
    Dim dateText As String
    
    Set re = CreateObject("VBScript.RegExp")
    
    re.Pattern = "^Glerum Govee .* #([0-9]{1,2})_export_([0-9]{8}).*\.csv$"
    re.IgnoreCase = True
    re.Global = False
    
    If Not re.Test(fileName) Then
        TryMatchGoveeFile = False
        Exit Function
    End If
    
    Set matches = re.Execute(fileName)
    
    sensorText = matches(0).SubMatches(0)
    dateText = matches(0).SubMatches(1)
    sensorNum = CLng(sensorText)
    
    If sensorNum < 1 Or sensorNum > 19 Then
        TryMatchGoveeFile = False
        Exit Function
    End If
    
    If dateText <> dateToken Then
        TryMatchGoveeFile = False
        Exit Function
    End If
    
    TryMatchGoveeFile = True
End Function

Private Sub InitializeLog(argLog As String)
    Dim fileNum As Integer
    
    fileNum = FreeFile
    Open argLog For Output As #fileNum
    Print #fileNum, "Govee import log"
    Close #fileNum
End Sub

Private Sub WriteLogLine(argLog As String, ByVal message As String)
    Dim fileNum As Integer
    
    fileNum = FreeFile
    Open argLog For Append As #fileNum
    Print #fileNum, Format(Now, "yyyy-mm-dd hh:nn:ss"); "  "; message
    Close #fileNum
End Sub

