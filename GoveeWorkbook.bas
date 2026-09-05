Attribute VB_Name = "GoveeWorkbook"
Option Explicit

Public Sub CloseWorkbookIfOpen(ByVal fullPath As String, ByVal szLog As String)
    Dim wb As Excel.Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            WriteLogLine szLog, "Closing already-open workbook: " & fullPath
            wb.Close SaveChanges:=False
            Exit For
        End If
    Next wb
End Sub

Public Sub WriteGoveeWorkbook(ByVal folderPath As String, ByVal dateToken As String, ByRef arrTimes() As String, ByRef dictSensors As Object, ByVal szLog As String)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim outputPath As String
    Dim sensorKeys() As String
    Dim sensorKey As String
    Dim sensorTimes As Object
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim i As Long
    Dim valuePair As Variant
    
    On Error GoTo WriteError
    
    outputPath = folderPath & "\Govee" & dateToken & ".xlsx"
    
    CloseWorkbookIfOpen outputPath, szLog

    Set wb = Application.Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "Data"
    
    WriteLogLine szLog, "Writing workbook: " & outputPath
    
    ws.Cells(1, 1).Value = "TimeStamp"
    
    For i = LBound(arrTimes) To UBound(arrTimes)
        ws.Cells(i - LBound(arrTimes) + 2, 1).Value = arrTimes(i)
    Next i
    
    sensorKeys = GetSortedSensorKeys(dictSensors)
    colIndex = 2
    
    For i = LBound(sensorKeys) To UBound(sensorKeys)
        sensorKey = sensorKeys(i)
        Set sensorTimes = dictSensors(sensorKey)
        
        ws.Cells(1, colIndex).Value = sensorKey & "Temp"
        ws.Cells(1, colIndex + 1).Value = sensorKey & "Humidity"
        
        For rowIndex = LBound(arrTimes) To UBound(arrTimes)
            If sensorTimes.Exists(arrTimes(rowIndex)) Then
                valuePair = sensorTimes(arrTimes(rowIndex))
                ws.Cells(rowIndex - LBound(arrTimes) + 2, colIndex).Value = valuePair(0)
                ws.Cells(rowIndex - LBound(arrTimes) + 2, colIndex + 1).Value = valuePair(1)
            Else
                ws.Cells(rowIndex - LBound(arrTimes) + 2, colIndex).Value = ""
                ws.Cells(rowIndex - LBound(arrTimes) + 2, colIndex + 1).Value = ""
            End If
        Next rowIndex
        
        WriteLogLine szLog, "Wrote columns for sensor " & sensorKey & _
            ": col " & colIndex & "=" & sensorKey & "Temp, col " & (colIndex + 1) & "=" & sensorKey & "Humidity"
        
        colIndex = colIndex + 2
    Next i
    
    ws.Columns.AutoFit
    
    Application.DisplayAlerts = False
    wb.SaveAs fileName:=outputPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True

    WriteLogLine szLog, "Saved workbook: " & outputPath
    wb.Close SaveChanges:=False

    Set ws = Nothing
    Set wb = Nothing

    Exit Sub

WriteError:
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    Set ws = Nothing
    Set wb = Nothing
    On Error GoTo 0
    
    WriteLogLine szLog, "WRITE ERROR: " & Err.Number & " - " & Err.Description
    MsgBox "Workbook write error: " & Err.Description, vbExclamation
End Sub

Public Function GetSortedSensorKeys(ByRef dictSensors As Object) As String()
    Dim keys As Variant
    Dim arr() As String
    
    keys = dictSensors.keys
    arr = VariantKeysToStringArray(keys)
    SortStringArray arr
    
    GetSortedSensorKeys = arr
End Function


