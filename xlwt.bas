Attribute VB_Name = "xlwt"
Option Explicit


Public fs As Object
Public fsLog As Object
Public stTimeStamp As String


Sub CreateFileSystemObject()

    Set fs = CreateObject("Scripting.FileSystemObject")

End Sub


Sub SetTimeStamp()

    stTimeStamp = Format(Now, "YYYY-MM-DD HH-MM-SS")

End Sub


Sub CreateLog(stLogNameExt As String)

    Call SetTimeStamp

    Set fsLog = fs.CreateTextFile( _
        Filename:= _
            Environ("USERPROFILE") & "\Desktop\" _
            & stTimeStamp & " - " & stLogNameExt, _
        Overwrite:=False _
    )

    fsLog.WriteLine (stTimeStamp & " - Created log.")
    Debug.Print (stTimeStamp & " - Created log.")

End Sub

    
