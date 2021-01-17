Sub Unzip_File()
Call UnzipAFile(Environ("Userprofile") & "\Downloads\TempFiles\Region_Mobility_Report_CSVs.zip", Environ("Userprofile") & "\Downloads\TempFiles\")
End Sub

Sub UnzipAFile(zippedFileFullName As Variant, unzipToPath As Variant)
Dim ShellApp As Object

Set ShellApp = CreateObject("shell.Application")
ShellApp.Namespace(unzipToPath).copyHere ShellApp.Namespace(zippedFileFullName).Items

End Sub
