Attribute VB_Name = "ParseCSV"
Option Explicit

' dictSensors("07")("2026-09-03 08:11:00") = Array("75.38", "56.4")

Public Sub ParseOneCsvFile(ByVal filePath As String, ByVal sensorNum As Long, ByRef dictTimes As Object, ByRef dictSensors As Object, ByVal szLog As String)
    Dim fileNum As Integer
    Dim lineText As String
    Dim lineNumber As Long
    Dim rowCount As Long
    Dim dataRowCount As Long
    Dim malformedCount As Long
    Dim duplicateCount As Long
    Dim differingDuplicateCount As Long
    Dim blankTimestampCount As Long
    Dim nonNumericCount As Long
    Dim timestamp As String
    Dim tempText As String
    Dim humidityText As String
    Dim parts() As String
    Dim sensorTimes As Object
    Dim firstTimestamp As String
    Dim lastTimestamp As String
    Dim priorValue As Variant
    Dim sensorKey As String
    
    On Error GoTo FileError
    
    sensorKey = Format(sensorNum, "00")

    If dictSensors.Exists(sensorKey) Then
        Set sensorTimes = dictSensors(sensorKey)
    Else
        Set sensorTimes = CreateObject("Scripting.Dictionary")
        dictSensors.Add sensorKey, sensorTimes
    End If
    
    WriteLogLine szLog, "Parsing sensor " & sensorKey & ": " & filePath
    
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineNumber = lineNumber + 1
        
        If lineNumber = 1 Then
            WriteLogLine szLog, "Header sensor " & sensorKey & ": " & lineText
        Else
            rowCount = rowCount + 1
            
            parts = Split(lineText, ",")
            
            If UBound(parts) < 2 Then
                malformedCount = malformedCount + 1
                WriteLogLine szLog, "Malformed row, sensor " & sensorKey & ", line " & lineNumber & ": " & lineText
                GoTo NextLine
            End If
            
            timestamp = Trim$(parts(0))
            tempText = Trim$(parts(1))
            humidityText = Trim$(parts(2))
            
            If Len(timestamp) = 0 Then
                blankTimestampCount = blankTimestampCount + 1
                WriteLogLine szLog, "Blank timestamp, sensor " & sensorKey & ", line " & lineNumber
                GoTo NextLine
            End If
            
            If Not IsNumeric(tempText) Or Not IsNumeric(humidityText) Then
                nonNumericCount = nonNumericCount + 1
                WriteLogLine szLog, "Non-numeric value, sensor " & sensorKey & ", line " & lineNumber & ": " & lineText
            End If
            
            If Not dictTimes.Exists(timestamp) Then
                dictTimes.Add timestamp, True
            End If
            
            If sensorTimes.Exists(timestamp) Then
                duplicateCount = duplicateCount + 1
                priorValue = sensorTimes(timestamp)
                
                If CStr(priorValue(0)) <> tempText Or CStr(priorValue(1)) <> humidityText Then
                    differingDuplicateCount = differingDuplicateCount + 1
                    WriteLogLine szLog, "Duplicate timestamp with different values, sensor " & sensorKey & ", time " & timestamp & _
                        ", prior=" & CStr(priorValue(0)) & "|" & CStr(priorValue(1)) & _
                        ", new=" & tempText & "|" & humidityText
                Else
                    WriteLogLine szLog, "Duplicate timestamp with same values, sensor " & sensorKey & ", time " & timestamp
                End If
            Else
                sensorTimes.Add timestamp, Array(tempText, humidityText)
                dataRowCount = dataRowCount + 1
                
                If Len(firstTimestamp) = 0 Then
                    firstTimestamp = timestamp
                End If
                lastTimestamp = timestamp
            End If
        End If
        
NextLine:
    Loop
    
    Close #fileNum
    
    WriteLogLine szLog, "Summary sensor " & sensorKey & _
        ": rows=" & rowCount & _
        ", distinctTimes=" & sensorTimes.Count & _
        ", duplicates=" & duplicateCount & _
        ", differingDuplicates=" & differingDuplicateCount & _
        ", malformed=" & malformedCount & _
        ", blankTimestamps=" & blankTimestampCount & _
        ", nonNumeric=" & nonNumericCount
    
    If sensorTimes.Count > 0 Then
        WriteLogLine szLog, "Range sensor " & sensorKey & ": first=" & firstTimestamp & ", last=" & lastTimestamp
    Else
        WriteLogLine szLog, "Range sensor " & sensorKey & ": no valid data rows"
    End If
    
    Exit Sub

FileError:
    On Error Resume Next
    Close #fileNum
    On Error GoTo 0
    WriteLogLine szLog, "FILE ERROR sensor " & Format(sensorNum, "00") & ": " & Err.Number & " - " & Err.Description & " [" & filePath & "]"

End Sub

Public Function GetSortedTimeArray(ByRef dictTimes As Object) As String()
    Dim keys As Variant
    Dim sortedKeys() As String

    keys = dictTimes.keys
    sortedKeys = VariantKeysToStringArray(keys)
    SortStringArray sortedKeys

    GetSortedTimeArray = sortedKeys
End Function

Public Sub LogGlobalTimeRange(ByRef dictTimes As Object, ByVal szLog As String)
    Dim keys As Variant
    Dim sortedKeys() As String
    
    keys = dictTimes.keys
    sortedKeys = VariantKeysToStringArray(keys)
    SortStringArray sortedKeys
    
    WriteLogLine szLog, "Global range: first=" & sortedKeys(LBound(sortedKeys)) & ", last=" & sortedKeys(UBound(sortedKeys))
End Sub

Private Function VariantKeysToStringArray(ByVal keys As Variant) As String()
    Dim arr() As String
    Dim i As Long
    
    ReDim arr(LBound(keys) To UBound(keys))
    
    For i = LBound(keys) To UBound(keys)
        arr(i) = CStr(keys(i))
    Next i
    
    VariantKeysToStringArray = arr
End Function

Private Sub SortStringArray(ByRef arr() As String)
    Dim i As Long
    Dim j As Long
    Dim temp As String
    
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) < arr(i) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub


