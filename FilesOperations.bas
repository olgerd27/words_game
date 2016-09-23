Attribute VB_Name = "FilesOperations"
' FileSystem's data
Dim FileLastDate As Date

' FileSystem's operations
Function IsFileExists(FileName$) As Boolean
    Set FS = CreateObject("Scripting.FileSystemObject")
    IsFileExists = FS.FileExists(FileName)
End Function

Sub CreateOfferFile()
    Open Settings.FilePathName For Output As #1
    Print #1, Settings.StartWord_File
    Print #1, Settings.Player1Name_File
    Print #1, Settings.Player2Name_File
    Close #1
End Sub

Sub LoadOfferFile()
    Dim FileData(1 To 3) As String
    Dim I As Integer
    
    Open Settings.FilePathName For Input As #1
    I = 1
    Do While Not EOF(1)
        Line Input #1, FileData(I)
        I = I + 1
    Loop
    Close #1
    
    Settings.SetStartWord (GetParam(FileData(1), Settings.GetSep_MaskData))
    ' if file was created by a some player 2, then 2-th line in the file is its name, 3-th - player 1 name
    Settings.SetPlayer2Name (GetParam(FileData(2), Settings.GetSep_MaskData))
    Settings.SetPlayer1Name (GetParam(FileData(3), Settings.GetSep_MaskData))
End Sub

Function GetParam(Str$, Sep$) As String
' Return string that is part of initial Str from first occurence the Sep to the Str end
    If Str = "" Then
        GetParam = ""
    Else
        GetParam = Mid(Str, InStr(Str, Sep) + 1)
    End If
End Function

Sub OutPlayer1Name()
    Dim StrLine As String

    Open Settings.FilePathName For Random As #1
    Get #1, 1, StrLine
    Settings.SetPlayer1Name (StrLine)
    Close #1
    MsgBox "OutPl: " & Settings.GetPlayer1Name
End Sub

Sub RemoveFile()
    If IsFileExists(Settings.FilePathName) Then
        Kill Settings.FilePathName
    End If
End Sub
