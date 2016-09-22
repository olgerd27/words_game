Attribute VB_Name = "SSOperations"
' SpreadSheet's operations
Sub PutNewGameInfo()
    'ClearGameInfo ' clear previous game info
    Set WrkSheet = Worksheets(Settings.GetSheetName)
    WrkSheet.Range(Settings.GetStartWord_SSAddr).Value = Settings.GetStartWord ' set start word
    WrkSheet.Range(Settings.GetPlayer1Name_SSAddr).Value = Settings.GetPlayer1Name ' set player 1 name
    WrkSheet.Range(Settings.GetPlayer2Name_SSAddr).Value = Settings.GetPlayer2Name ' set player 2 name
    ' TODO: place the start word at the game field center - run some Sub for this purpose
End Sub

Sub ClearGameInfo()
    Set WrkSheet = Worksheets(Settings.GetSheetName)
    WrkSheet.Range(Settings.GetStartWord_SSAddr).Value = "" ' clear start word
    WrkSheet.Range(Settings.GetPlayer1Name_SSAddr).ClearContents ' clear player 1 name
    WrkSheet.Range(Settings.GetPlayer2Name_SSAddr).ClearContents ' clear player 2 name
    ' TODO: clear the game field
End Sub
