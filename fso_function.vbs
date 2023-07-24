

    Dim ofso, file1
    Set ofso = CreateObject("Scripting.FileSystemObject")

    Set file1 = ofso.OpenTextFile("C:\Work\vb_fso\szovegfile1.txt", 8, True, 0) 'append some text to each textfile
    file1.WriteLine("This is an addition to each file!!!!")
    file1.Close
    
    Set file1 = Nothing
    Set ofso = Nothing

