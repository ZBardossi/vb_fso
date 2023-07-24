Dim objExcel, filePath
filePath = "C:\Work\vb_fso\demo.xlsx"

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set obj1 = objExcel.Workbooks.Open(filePath)
Set obj2 = obj1.Worksheets("Munka1")

MsgBox(obj2.cells(2,1).Value)

obj1.Close
objExcel.Quit

Set obj1 = Nothing
Set obj2 = Nothing
Set objExcel = Nothing

