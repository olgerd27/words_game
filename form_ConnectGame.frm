VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_ConnectGame 
   Caption         =   "Подключение"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6660
   OleObjectBlob   =   "form_ConnectGame.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_ConnectGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Cancel_Click()
    Call Settings.SetPlayer1Name(form_CreateGame.tb_Player1Name.Text)
    form_ConnectGame.Hide
'    ClearTextBoxes
End Sub

Private Sub cmd_Ok_Click()
    form_ConnectGame.Hide
'    ClearTextBoxes
End Sub

