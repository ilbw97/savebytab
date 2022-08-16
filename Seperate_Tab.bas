Attribute VB_Name = "Module1"
Sub Separate_Tab()

Dim Directory_Path As String
Directory_Path = Application.ActiveWorkbook.Path

Application.ScreenUpdating = False
Application.DisplayAlerts = False

FileName = ActiveWorkbook.Name
If InStr(FileName, ".") > 0 Then
   FileName = Left(FileName, InStr(FileName, ".") - 1)
End If

For Each Tab_name In ThisWorkbook.Sheets
    Tab_name.Copy
    Application.ActiveWorkbook.SaveAs FileName:=Directory_Path & "/" & FileName & "_" & Tab_name.Name & ".xlsx"
    Application.ActiveWorkbook.Close False
Next

Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub
