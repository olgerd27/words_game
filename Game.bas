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
        AskNewGameData ' ask only player name
        SSOperations.PutNewGameInfo
    Else
        AskNewGameData ' ask the player name and the start word
        If IsGameByPlayer1Created Then
            FilesOperations.CreateOfferFile
        End If
    End If
    ' uninterrupted checking of the file changes: Do ... While two players connection was not established
End Sub

Sub AskNewGameData()
    ' If the player 1 name is not set -> show window for asking it
    If IsGameNobodyCreated Then
        form_NewGame.lbl_Title = "Создание новой"
        form_NewGame.lbl_StartWord = "Введите начальное слово:"
        form_NewGame.tb_StartWord.Locked = False
    ElseIf IsGameByPlayer2Created Then
        form_NewGame.lbl_Title = "Подключение к существующей"
        form_NewGame.lbl_StartWord = "Начальное слово:"
        form_NewGame.tb_StartWord.Text = Settings.GetStartWord
        form_NewGame.tb_StartWord.Locked = True
    End If
    form_NewGame.Show ' open dialog
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

