Attribute VB_Name = "xlwt"
Option Explicit


Public fs As Object
Public stTimeStamp As String


Sub CreateFileSystemObject()

    Set fs = CreateObject("Scripting.FileSystemObject")

End Sub


Sub SetTimeStamp()

    stTimeStamp = Format(Now, "YYYY-MM-DD HH-MM-SS")

End Sub


