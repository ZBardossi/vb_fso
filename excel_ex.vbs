'Create an excel file

Dim ofso, objExcel, filePath
filePath = "C:\Work\vb_fso\demo.xlsx"

Set ofso = CreateObject("Scripting.FileSystemObject")
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

If ofso.FileExists(filePath) Then
    objExcel.Workbooks.Open(filePath)
    objExcel.Worksheets(1).cells(1,2)="New Text"
    objExcel.Worksheets(1).cells(2,1) = "Another new text"
    objExcel.ActiveWorkbook.Save

Else
    objExcel.Workbooks.Add
    objExcel.Worksheets(1).cells(1,1) = "Text2"
    objExcel.ActiveWorkbook.SaveAs(filePath)

End If

objExcel.Quit
Set ofso = Nothing
Set objExcel = Nothing

