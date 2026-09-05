Attribute VB_Name = "Govee"
' Alt-F11 opens VBA

Option Explicit

' Public szLog As String

Public Sub RunGoveeImport()
    Dim folderPath As String
    Dim dateToken As String
    Dim szLog As String
    Dim szFile As String
    Dim fullPath As String
    Dim sensorNum As Long
    Dim dictTimes As Object
    Dim dictSensors As Object
    Dim matchedCount As Long
    Dim arrTimes() As String
    
    On Error GoTo FatalError
    
    folderPath = GetGoveeFolder()
    dateToken = Format(Date - 1, "yyyymmdd")
    szLog = folderPath & "\Govee" & dateToken & ".txt"
    
    Set dictTimes = CreateObject("Scripting.Dictionary")
    Set dictSensors = CreateObject("Scripting.Dictionary")
    
    InitializeLog szLog
    WriteLogLine szLog, "START"
    WriteLogLine szLog, "Folder: " & folderPath
    WriteLogLine szLog, "Date token: " & dateToken
    
    szFile = Dir(folderPath & "\*" & dateToken & "*.csv")
    
    If szFile = "" Then
        WriteLogLine szLog, "No CSV files found."
    End If
    
    Do While szFile <> ""
        fullPath = folderPath & "\" & szFile
        WriteLogLine szLog, "Found: " & szFile
        
        If TryMatchGoveeFile(szFile, dateToken, sensorNum) Then
            matchedCount = matchedCount + 1
            WriteLogLine szLog, "Accepted: sensor " & Format(sensorNum, "00") & " -> " & szFile
            ParseOneCsvFile fullPath, sensorNum, dictTimes, dictSensors, szLog
        Else
            WriteLogLine szLog, "Skipped: " & szFile
        End If
        
        szFile = Dir()
    Loop
    
    WriteLogLine szLog, "Matched files: " & matchedCount
    WriteLogLine szLog, "Distinct timestamps across all files: " & dictTimes.Count
    
    If dictTimes.Count > 0 Then
        arrTimes = GetSortedTimeArray(dictTimes)

        LogGlobalTimeRange dictTimes, szLog

        WriteLogLine szLog, "Master timestamp rows: " & (UBound(arrTimes) - LBound(arrTimes) + 1)
        WriteLogLine szLog, "Master first timestamp: " & arrTimes(LBound(arrTimes))
        WriteLogLine szLog, "Master last timestamp: " & arrTimes(UBound(arrTimes))
    End If

    WriteLogLine szLog, "Sensors loaded: " & dictSensors.Count

    If dictTimes.Count > 0 And dictSensors.Count > 0 Then
        WriteGoveeWorkbook folderPath, dateToken, arrTimes, dictSensors, szLog
    Else
        WriteLogLine szLog, "Workbook not written because no data was available."
    End If

    Exit Sub

FatalError:
    WriteLogLine szLog, "FATAL ERROR: " & Err.Number & " - " & Err.Description
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

Private Function TryMatchGoveeFile(ByVal szFile As String, ByVal dateToken As String, ByRef sensorNum As Long) As Boolean
    Dim re As Object
    Dim matches As Object
    Dim sensorText As String
    Dim dateText As String
    
    Set re = CreateObject("VBScript.RegExp")
    
    re.Pattern = "^Glerum Govee .* #([0-9]{1,2})_export_([0-9]{8}).*\.csv$"
    re.IgnoreCase = True
    re.Global = False
    
    If Not re.Test(szFile) Then
        TryMatchGoveeFile = False
        Exit Function
    End If
    
    Set matches = re.Execute(szFile)
    
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

