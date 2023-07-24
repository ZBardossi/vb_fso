Dim ofso

'Create a file system object
Set ofso = CreateObject("Scripting.FileSystemObject")
'ofso.CreateTextFile("sample.txt")

'Opening the text file
Set file1 = ofso.OpenTextFile("sample.txt", 2) '2 means open for writing
file1.WriteLine("It is the content of the file, file system object practice")
file1.Write("This is another line")
file1.Close

'Release file
Set file1 = Nothing
'Release object
Set ofso = Nothing


