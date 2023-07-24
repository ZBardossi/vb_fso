Dim ofso
Dim path
path = "C:\Work\vb_fso\szovegfile1.txt"

Set ofso = CreateObject("Scripting.FileSystemObject")
Const forReading = 1
Const forWriting = 2

Set file1 = ofso.OpenTextFile(path, forReading)
'MsgBox(file1.ReadAll) 'Entire content
MsgBox(file1.ReadLine) 'First line 
'MsgBox(file1.Read(20)) 'first n character
file1.Close

Set file1 = Nothing
Set ofso = Nothing
