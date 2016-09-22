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
        AskNewGameData
        SSOperations.PutNewGameInfo
    Else
        'form_NewGame.Show ' ask the player name and the start word
        AskNewGameData
        FilesOperations.CreateOfferFile
    End If
    ' uninterrupted checking of the file changes: Do ... While two players connection was not established
End Sub

Sub AskNewGameData()
    ' If the player 1 name is not set -> show window for asking it
    If Settings.GetStartWord = "" And Settings.GetPlayer1Name = "" Then
        form_NewGame.tb_StartWord.Locked = False
    ElseIf Settings.GetStartWord <> "" And Settings.GetPlayer1Name = "" Then
        form_NewGame.tb_StartWord.Text = Settings.GetStartWord
        form_NewGame.tb_StartWord.Locked = True
    End If
    form_NewGame.Show ' open dialog
End Sub
