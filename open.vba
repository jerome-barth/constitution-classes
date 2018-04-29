Private Sub Workbook_Open()
    Application.Calculate
    Dim ws As Worksheet
    On Error GoTo Oops
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        ws.EnableOutlining = True
        ws.Protect UserInterfaceOnly:=True, AllowInsertingRows:=True, AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True
    Next ws
    Worksheets("Patates").Activate
Oops:
End Sub