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
' ---------------------------------------------------------------
Public Sub BuildSensorCatalog()
    Dim n As Long

    Set gSensorIndex = CreateObject("Scripting.Dictionary")
    ReDim gSensors(0 To 18)   ' room for 19 sensors, 01..19

    n = 0
    AddSensor n, "Boiler", "1", RGB(0, 0, 255), 2, xlSolid
    AddSensor n, "Unassigned", "2", RGB(0, 0, 255), 2, xlSolid
    AddSensor n, "B1.wall", "3", RGB(0, 0, 255), 2, xlSolid
    AddSensor n, "Freezer", "4", RGB(0, 0, 255), 2, xlSolid
    AddSensor n, "Blowoff", "5", RGB(0, 0, 255), 2, xlSolid
    AddSensor n, "Pool.vent", "6", RGB(0, 0, 255), 2, xlSolid
    AddSensor n, "ADU.wall", "7", RGB(0, 0, 255), 2, xlSolid
    AddSensor n, "Fridge", "8", RGB(0, 0, 255), 2, xlSolid
    AddSensor n, "SquirrelCage", "9", RGB(0, 0, 255), 2, xlSolid
    AddSensor n, "B1.vent", "10", RGB(0, 0, 255), 2, xlSolid
    AddSensor n, "Pool.feed", "11", RGB(0, 0, 255), 2, xlSolid
    AddSensor n, "B1.input", "12", RGB(0, 0, 255), 2, xlSolid
    AddSensor n, "KitchenPatio", "13", RGB(0, 0, 255), 2, xlSolid
    AddSensor n, "Pool.dump", "14", RGB(0, 0, 255), 2, xlSolid
    AddSensor n, "B3.feed", "16", RGB(0, 0, 255), 2, xlSolid
    AddSensor n, "B1.dump", "17", RGB(0, 0, 255), 2, xlSolid
    AddSensor n, "B1.handler", "18", RGB(0, 0, 255), 2, xlSolid
    AddSensor n, "B1.feed", "19", RGB(0, 0, 255), 2, xlSolid
    
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


