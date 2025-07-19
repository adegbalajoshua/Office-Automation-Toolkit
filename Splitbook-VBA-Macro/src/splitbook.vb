Sub Splitbook()
    Dim xPath As String
    Dim xWs As Worksheet
    xPath = Application.ActiveWorkbook.Path
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    For Each xWs In ThisWorkbook.Sheets
        On Error Resume Next
        With xWs.PageSetup
            .Orientation = xlLandscape
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
        End With
        xWs.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=xPath & "\" & xWs.Name & ".pdf", _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
        On Error GoTo 0
    Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub