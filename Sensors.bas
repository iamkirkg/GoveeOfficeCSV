Attribute VB_Name = "Sensors"
Option Explicit

Public Type SensorDef
    label As String
    govee As String
    lineColor As Long
    lineWeight As Double
    lineDash As Long
End Type

Private gSensors() As SensorDef
Private gSensorIndex As Object
Private gCatalogReady As Boolean

' ---------------------------------------------------------------
' Build the catalog once per run.
' Call this at the start of RunGoveeImport (or ProcessRequestedDates).

' msoLineSolid        ' solid
' msoLineDash         ' dashed
' msoLineDot          ' dotted
' msoLineRoundDot     ' round dots
' msoLineDashDot      ' dash-dot
' msoLineDashDotDot   ' dash-dot-dot
' msoLineLongDash
' msoLineLongDashDot

' ---------------------------------------------------------------
Public Sub BuildSensorCatalog()
    Dim n As Long

    Set gSensorIndex = CreateObject("Scripting.Dictionary")
    ReDim gSensors(0 To 18)   ' room for 19 sensors, 01..19

    n = 0

    AddSensor n, "Boiler", "01", clrBlue, 1, msoLineDash
    AddSensor n, "Blowoff", "05", clrOlive, 1, msoLineDash

    AddSensor n, "B1.input", "12", clrRed, 1, msoLineSolid
    AddSensor n, "B1.dump", "17", clrGreen, 1, msoLineRoundDot
    AddSensor n, "B1.feed", "19", clrPurple, 1, msoLineSolid
    AddSensor n, "B1.handler", "18", clrNavy, 1, msoLineSolid
    AddSensor n, "B1.vent", "10", clrOrange, 1, msoLineSolid
    AddSensor n, "B1.wall", "03", clrBlue, 1, msoLineSolid

    AddSensor n, "B3.feed", "16", clrPurple, 1, msoLineSolid

    AddSensor n, "ADU.vent", "02", clrOrange, 1, msoLineSolid
    AddSensor n, "ADU.wall", "07", clrBlue, 1, msoLineSolid

    AddSensor n, "Pool.vent", "06", clrOrange, 1, msoLineSolid
    AddSensor n, "Pool.feed", "11", clrPurple, 1, msoLineSolid
    AddSensor n, "Pool.dump", "14", clrGreen, 1, msoLineDash

    AddSensor n, "Freezer", "04", clrPink, 1, msoLineSolid
    AddSensor n, "Fridge", "08", clrTeal, 1, msoLineSolid
    AddSensor n, "SquirrelCage", "09", clrRed, 1, msoLineSolid
    AddSensor n, "KitchenPatio", "13", clrOlive, 1, msoLineSolid
    
    ReDim Preserve gSensors(0 To n - 1)
    gCatalogReady = True
End Sub

' n is the slot index we're filling; pass a ByRef counter.
Private Sub AddSensor(ByRef n As Long, ByVal label As String, ByVal govee As String, _
                      ByVal lineColor As Long, ByVal lineWeight As Double, ByVal lineDash As Long)
    gSensors(n).label = label
    gSensors(n).govee = govee
    gSensors(n).lineColor = lineColor
    gSensors(n).lineWeight = lineWeight
    gSensors(n).lineDash = lineDash

    If Not gSensorIndex.Exists(label) Then
        gSensorIndex.Add label, n
    End If

    n = n + 1
End Sub

Public Function HasSensor(ByVal label As String) As Boolean
    If Not gCatalogReady Then Err.Raise vbObjectError + 2000, , "Sensor catalog not built."
    HasSensor = gSensorIndex.Exists(label)
End Function

' Returns a copy of the SensorDef for the given label.
Public Function GetSensor(ByVal label As String) As SensorDef
    Dim idx As Long

    If Not gCatalogReady Then Err.Raise vbObjectError + 2001, , "Sensor catalog not built."

    If Not gSensorIndex.Exists(label) Then
        Err.Raise vbObjectError + 2002, , "Unknown sensor label: " & label
    End If

    idx = gSensorIndex(label)
    GetSensor = gSensors(idx)
End Function

' Convenience: sensor number only, e.g. "01".
Public Function GetSensorGovee(ByVal label As String) As String
    GetSensorGovee = GetSensor(label).govee
End Function


