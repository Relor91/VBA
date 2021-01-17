Sub MakeDir()
On Error GoTo ErrorHandler:
MkDir (Environ("Userprofile") & "\Downloads\TempFiles")
ErrorHandler:
    Resume Next
    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    FSO.DeleteFolder Environ("Userprofile") & "\Downloads\TempFiles", False
MkDir (Environ("Userprofile") & "\Downloads\TempFiles")
End Sub
