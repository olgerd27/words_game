VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_CreateGame 
   Caption         =   "�������� �����"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6795
   OleObjectBlob   =   "form_CreateGame.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_CreateGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Ok_Click()
    Call Settings.SetPlayer1Name(form_CreateGame.tb_PlayerName.Text)
    Call Settings.SetStartWord(LCase(form_CreateGame.tb_StartWord.Text))
    form_CreateGame.Hide
    ClearTextBoxes
End Sub

Private Sub cmd_Cancel_Click()
    form_CreateGame.Hide
    ClearTextBoxes
End Sub

Private Sub ClearTextBoxes()
    form_CreateGame.tb_PlayerName.Text = ""
    form_CreateGame.tb_StartWord.Text = ""
End Sub

