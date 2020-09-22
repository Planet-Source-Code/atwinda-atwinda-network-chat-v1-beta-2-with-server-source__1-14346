VERSION 5.00
Begin VB.Form frmPM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atwinda Chat - Private Message"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   Icon            =   "frmPM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSendPM 
      Caption         =   "Send PM"
      Default         =   -1  'True
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtPM 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4935
   End
   Begin VB.Label lblUserPM 
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Send to:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Whatever you type in this window only the user you selected will see."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmPM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub cmdSendPM_Click()
If txtPM.Text <> "" Then
Dim allText As String
allText = frmChat.txtUser.Text & "|" & txtPM.Text
Call frmChat.wskClient.SendData("PMMessage " & lblUserPM.Caption & ":" & allText)
txtPM.Text = ""
DoEvents
End If
End Sub

Private Sub txtPM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSendPM_Click
    KeyAscii = 0
End If
End Sub
