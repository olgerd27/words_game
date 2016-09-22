Attribute VB_Name = "Settings"
' >>> The settigs (global variables) module. $ -> As String
Dim GamePath$ ' path to the data file
Dim FileName$ ' name of the data file

Dim Player1Name$ ' a name of player #1
Dim Player2Name$ ' a name of player #2
Dim StartWord$ ' the start word

Dim Player1Name_Mask$ ' mask in the file for recognition name of player 1
Dim Player2Name_Mask$ ' mask in the file for recognition name of player 2
Dim StartWord_Mask$ ' mask in the file for recognition start word
Dim Sep_MaskData$ ' separator Mask - Data

Dim SheetName$ ' the game sheet name
Dim StartWord_SSAddr$ ' spreadsheet address of the start word
Dim Player1Name_SSAddr$ ' spreadsheet address of the player 1 name
Dim Player2Name_SSAddr$ ' spreadsheet address of the player 2 name

' Getters and Setters of the global variable (for access from others modules)
' GamePath
Function GetGamePath() As String
    GetGamePath = GamePath
End Function

Sub SetGamePath(Path$)
    GamePath = Path
End Sub

' FileName
Function GetFileName() As String
    GetFileName = FileName
End Function

Sub SetFileName(FName$)
    FileName = FName
End Sub

' Player1Name
Function GetPlayer1Name() As String
    GetPlayer1Name = Player1Name
End Function

Sub SetPlayer1Name(Name$)
    Player1Name = Name
End Sub

' Player2Name
Function GetPlayer2Name() As String
    GetPlayer2Name = Player2Name
End Function

Sub SetPlayer2Name(Name$)
    Player2Name = Name
End Sub

' StartWord
Function GetStartWord() As String
    GetStartWord = StartWord
End Function

Sub SetStartWord(StWord$)
    StartWord = StWord
End Sub

' Player1Name_Mask
Function GetPlayer1Name_Mask() As String
    GetPlayer1Name_Mask = Player1Name_Mask
End Function

Sub SetPlayer1Name_Mask(Mask$)
    Player1Name_Mask = Mask
End Sub

' Player2Name_Mask
Function GetPlayer2Name_Mask() As String
    GetPlayer2Name_Mask = Player2Name_Mask
End Function

Sub SetPlayer2Name_Mask(Mask$)
    Player2Name_Mask = Mask
End Sub

' StartWord_Mask
Function GetStartWord_Mask() As String
    GetStartWord_Mask = StartWord_Mask
End Function

Sub SetStartWord_Mask(Mask$)
    StartWord_Mask = Mask
End Sub

' Sep_MaskData
Function GetSep_MaskData() As String
    GetSep_MaskData = Sep_MaskData
End Function

Sub SetSep_MaskData(Mask$)
    Sep_MaskData = Mask
End Sub

' SheetName
Function GetSheetName() As String
    GetSheetName = SheetName
End Function

Sub SetSheetName(ShName$)
    SheetName = ShName
End Sub

' StartWord_SSAddr
Function GetStartWord_SSAddr() As String
    GetStartWord_SSAddr = StartWord_SSAddr
End Function

Sub SetStartWord_SSAddr(Addr$)
    StartWord_SSAddr = Addr
End Sub

' Player1Name_SSAddr
Function GetPlayer1Name_SSAddr() As String
    GetPlayer1Name_SSAddr = Player1Name_SSAddr
End Function

Sub SetPlayer1Name_SSAddr(Addr$)
    Player1Name_SSAddr = Addr
End Sub

' Player2Name_SSAddr
Function GetPlayer2Name_SSAddr() As String
    GetPlayer2Name_SSAddr = Player2Name_SSAddr
End Function

Sub SetPlayer2Name_SSAddr(Addr$)
    Player2Name_SSAddr = Addr
End Sub

' >>> Data manipulation functions
Function FilePathName() As String
    FilePathName = GamePath & FileName
End Function

Function StartWord_File() As String
    StartWord_File = StartWord_Mask & Sep_MaskData & StartWord
End Function

Function Player1Name_File() As String
    Player1Name_File = Player1Name_Mask & Sep_MaskData & Player1Name
End Function

Function Player2Name_File() As String
    Player2Name_File = Player2Name_Mask & Sep_MaskData & Player2Name
End Function
