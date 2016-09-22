Attribute VB_Name = "Main"
' The main module
Sub NewGame()
    InitGame
    Game.PrepareNewGame
End Sub

Sub InitGame()
    Settings.SetGamePath (ActiveWorkbook.Path & "\")
    Settings.SetFileName ("words.doc")
    Settings.SetStartWord ("")
    Settings.SetPlayer1Name ("")
    Settings.SetPlayer2Name ("")
    Settings.SetPlayer1Name_Mask ("[pl1]")
    Settings.SetPlayer2Name_Mask ("[pl2]")
    Settings.SetStartWord_Mask ("[stword]")
    Settings.SetSep_MaskData ("=")
    Settings.SetSheetName ("Game")
    Settings.SetStartWord_SSAddr ("A1")
    Settings.SetPlayer1Name_SSAddr ("AA1")
    Settings.SetPlayer2Name_SSAddr ("AF1")
End Sub

Sub MakeCourse()
    MsgBox "The course made"
    
    ' The course file has:
    ' - game field with data
    ' - player 1 name
    ' - player 1 online status
    ' - player 2 name
    ' - player 2 online status
End Sub

Sub EndGame()
    InitGame
    FilesOperations.RemoveFile
    SSOperations.ClearGameInfo
End Sub
