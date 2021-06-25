Attribute VB_Name = "Module1"
Public Sub Main()
    
    Dim iFileNo As Integer
    iFileNo = FreeFile
    Open "C:\test\project\output.txt" For Output As #iFileNo
    Write #iFileNo, "something"
    Close #iFileNo
End Sub
