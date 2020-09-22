VERSION 5.00
Begin VB.Form frmPrompt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atwinda Network Chat - Prompt"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "frmPrompt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtPrompt 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSend_Click()
If txtPrompt <> "" Then
    Call frmChat.wskClient.SendData("Prompt " & txtPrompt.Text)
    DoEvents
    txtPrompt.Text = ""
End If
End Sub

Private Sub txtPrompt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSend_Click
    KeyAscii = 0
End If
End Sub
