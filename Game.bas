Attribute VB_Name = "Game"
' The Game logic
Sub PrepareNewGame()
    ' Check - is offer to play exists (check existing of offer file)
    '     If YES - show created game info and connect to this game
    '     else - create new game (open dialog for input player name and start word, and create new offer file)
    ' Clear spreadsheet for new game
    ' Make "Start game" button disabled

    If FilesOperations.IsFileExists(Settings.FilePathName) Then
        FilesOperations.LoadOfferFile
        'MsgBox "Start Word: " & Settings.GetStartWord & ", Player 1: " & Settings.GetPlayer1Name & ", Player 2: " & Settings.GetPlayer2Name
        ShowCreatedGameDialog ' open connect dialog for showing created game info and asking player 1 name
        FilesOperations.OutPlayer1Name
        SSOperations.PutNewGameInfo
    Else
        form_CreateGame.Show ' open create new dialog for asking the player name and the start word
        If IsGameByPlayer1Created Then
            FilesOperations.CreateOfferFile
            SSOperations.PutNewGameInfo
        End If
    End If
    ' uninterrupted checking of the file changes: Do ... While two players connection was not established
End Sub

Sub ShowCreatedGameDialog()
    form_ConnectGame.tb_StartWord.Text = Settings.GetStartWord
    form_ConnectGame.tb_Player2Name = Settings.GetPlayer2Name
    form_ConnectGame.Show
End Sub

' Status application getters
Function IsGameByPlayer1Created() As Boolean
    IsGameByPlayer1Created = (Settings.GetStartWord <> "" And Settings.GetPlayer1Name <> "" And Settings.GetPlayer2Name = "")
End Function

Function IsGameByPlayer2Created() As Boolean
    IsGameByPlayer2Created = (Settings.GetStartWord <> "" And Settings.GetPlayer1Name = "" And Settings.GetPlayer2Name <> "")
End Function

Function IsGameNobodyCreated() As Boolean
    IsGameNobodyCreated = (Settings.GetStartWord = "" And Settings.GetPlayer1Name = "" And Settings.GetPlayer2Name = "")
End Function

