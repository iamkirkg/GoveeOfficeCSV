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
    Dim wsData As Worksheet
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
    Set wsData = wb.Worksheets(1)
    wsData.Name = "Data"
    
    WriteLogLine szLog, "Writing workbook: " & outputPath
    
    wsData.Cells(1, 1).Value = "TimeStamp"
    
    For i = LBound(arrTimes) To UBound(arrTimes)
        wsData.Cells(i - LBound(arrTimes) + 2, 1).Value = arrTimes(i)
    Next i
    
    sensorKeys = GetSortedSensorKeys(dictSensors)
    colIndex = 2
    
    For i = LBound(sensorKeys) To UBound(sensorKeys)
        sensorKey = sensorKeys(i)
        Set sensorTimes = dictSensors(sensorKey)
        
        wsData.Cells(1, colIndex).Value = sensorKey & "Temp"
        wsData.Cells(1, colIndex + 1).Value = sensorKey & "Humidity"
        
        For rowIndex = LBound(arrTimes) To UBound(arrTimes)
            If sensorTimes.Exists(arrTimes(rowIndex)) Then
                valuePair = sensorTimes(arrTimes(rowIndex))
                wsData.Cells(rowIndex - LBound(arrTimes) + 2, colIndex).Value = valuePair(0)
                wsData.Cells(rowIndex - LBound(arrTimes) + 2, colIndex + 1).Value = valuePair(1)
            Else
                wsData.Cells(rowIndex - LBound(arrTimes) + 2, colIndex).Value = ""
                wsData.Cells(rowIndex - LBound(arrTimes) + 2, colIndex + 1).Value = ""
            End If
        Next rowIndex
        
        WriteLogLine szLog, "Wrote columns for sensor " & sensorKey & _
            ": col " & colIndex & "=" & sensorKey & "Temp, col " & (colIndex + 1) & "=" & sensorKey & "Humidity"
        
        colIndex = colIndex + 2
    Next i

    wsData.Activate
    wsData.Range("B2").Select
    ActiveWindow.FreezePanes = True

    wsData.Columns.AutoFit

    Dim wsMainHouse As Worksheet
    Set wsMainHouse = wb.Worksheets.Add(After:=wsData)
    wsMainHouse.Name = "Main House"
    AddChart wsData, wsMainHouse, dateToken, _
        Array("Boiler", "Blowoff", "B1.input", "B1.dump", "B1.feed", "B1.handler", "B1.vent", "B1.wall"), _
        "Main House Sensors", szLog

    Dim wsPool As Worksheet
    Set wsPool = wb.Worksheets.Add(After:=wsMainHouse)
    wsPool.Name = "Pool"
    AddChart wsData, wsPool, dateToken, _
        Array("Boiler", "Blowoff", "Pool.vent", "Pool.feed", "Pool.dump"), _
        "Pool Sensors", szLog

    Application.DisplayAlerts = False
    wb.SaveAs fileName:=outputPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True

    WriteLogLine szLog, "Saved workbook: " & outputPath
    wb.Close SaveChanges:=False

    Set wsData = Nothing
    Set wsMainHouse = Nothing
    Set wb = Nothing

    Exit Sub

WriteError:
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    Set wsData = Nothing
    Set wsMainHouse = Nothing
    Set wb = Nothing
    On Error GoTo 0
    
    WriteLogLine szLog, "WRITE ERROR: " & Err.Number & " - " & Err.Description
    MsgBox "Workbook write error: " & Err.Description, vbExclamation
End Sub

Private Function DateTokenToDate(ByVal dateToken As String) As Date
    DateTokenToDate = DateSerial(CInt(Left$(dateToken, 4)), _
                                 CInt(Mid$(dateToken, 5, 2)), _
                                 CInt(Right$(dateToken, 2)))
End Function

Public Function GetSortedSensorKeys(ByRef dictSensors As Object) As String()
    Dim keys As Variant
    Dim arr() As String
    
    keys = dictSensors.keys
    arr = VariantKeysToStringArray(keys)
    SortStringArray arr
    
    GetSortedSensorKeys = arr
End Function

Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim lastCol As Long
    Dim col As Long

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For col = 1 To lastCol
        If CStr(ws.Cells(1, col).Value) = headerText Then
            FindHeaderColumn = col
            Exit Function
        End If
    Next col

    FindHeaderColumn = 0
End Function

Private Sub AddChart(ByVal wsData As Worksheet, ByVal wsChart As Worksheet, ByVal dateToken As String, _
                     ByVal chartSensors As Variant, ByVal chartTitle As String, ByVal szLog As String)
    Dim chartObj As ChartObject
    Dim ch As Chart
    Dim colTime As Long
    Dim lastRow As Long
    Dim targetDate As Date
    Dim i As Long
    Dim label As String
    Dim s As SensorDef
    Dim colSeries As Long

    targetDate = DateTokenToDate(dateToken)
    colTime = FindHeaderColumn(wsData, "TimeStamp")
    If colTime = 0 Then
        WriteLogLine szLog, "Chart not added (no TimeStamp column): " & chartTitle
        Exit Sub
    End If
    lastRow = wsData.Cells(wsData.Rows.Count, colTime).End(xlUp).Row

    Set chartObj = wsChart.ChartObjects.Add(Left:=0, Top:=0, Width:=900, Height:=450)
    Set ch = chartObj.Chart
    ch.ChartType = xlXYScatterSmooth

    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop

    For i = LBound(chartSensors) To UBound(chartSensors)
        label = CStr(chartSensors(i))

        If Not HasSensor(label) Then
            WriteLogLine szLog, "Chart " & chartTitle & ": unknown sensor '" & label & "' - skipped."
            GoTo NextSensor
        End If

        s = GetSensor(label)
        colSeries = FindHeaderColumn(wsData, s.govee & "Temp")

        If colSeries = 0 Then
            WriteLogLine szLog, "Chart " & chartTitle & ": no data for " & s.label & " (" & s.govee & "Temp) - skipped."
            GoTo NextSensor
        End If

        With ch.SeriesCollection.NewSeries
            .Name = s.label
            .XValues = wsData.Range(wsData.Cells(2, colTime), wsData.Cells(lastRow, colTime))
            .Values = wsData.Range(wsData.Cells(2, colSeries), wsData.Cells(lastRow, colSeries))
            .MarkerStyle = xlMarkerStyleNone
            .Format.Line.Weight = s.lineWeight
            .Format.Line.ForeColor.RGB = s.lineColor
            .Format.Line.DashStyle = s.lineDash
        End With

        WriteLogLine szLog, "Chart " & chartTitle & ": added " & s.label & " (col " & colSeries & ")"
NextSensor:
    Next i

    ch.Axes(xlCategory).MinimumScale = CDbl(targetDate)
    ch.Axes(xlCategory).MaximumScale = CDbl(targetDate + 1)
    ch.Axes(xlCategory).TickLabels.NumberFormat = "m/d h:mm"

    ch.Axes(xlValue).MinimumScale = 50
    ch.Axes(xlValue).MaximumScale = 130

    ch.HasTitle = True
    ch.chartTitle.Text = chartTitle & " : " & targetDate

    ch.Axes(xlCategory).HasTitle = True
    ch.Axes(xlCategory).AxisTitle.Text = "TimeStamp"

    ch.Axes(xlValue).HasTitle = True
    ch.Axes(xlValue).AxisTitle.Text = "Temperature (F)"

    WriteLogLine szLog, "Added chart: " & chartTitle
End Sub

