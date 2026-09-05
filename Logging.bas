Attribute VB_Name = "Logging"

Public Sub WriteLogLine(argLog As String, ByVal message As String)
    Dim fileNum As Integer
    
    fileNum = FreeFile
    Open argLog For Append As #fileNum
    Print #fileNum, Format(Now, "yyyy-mm-dd hh:nn:ss"); "  "; message
    Close #fileNum
End Sub


