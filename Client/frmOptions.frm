VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atwinda Chat - Options"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSoundOpt 
      Caption         =   "Sound Options"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdDisplayOpt 
      Caption         =   "Display Options"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.Frame famSound 
      Caption         =   "Sound Options"
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   4575
      Begin VB.CheckBox chkSound 
         Caption         =   "Enabled"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Please leave the box check for sounds, and to turn them off, uncheck the box."
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame famDisplay 
      Caption         =   "Display Options"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4575
      Begin VB.CheckBox chkLine 
         Caption         =   "Display Correcting"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   $"frmOptions.frx":0CCA
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DisplayC As Boolean, SoundC As Boolean

Private Sub chkLine_Click()
If chkLine.Value = 1 Then
    DisplayC = True
ElseIf chkLine.Value = 0 Then
    DisplayC = False
End If
End Sub

Private Sub chkSound_Click()
If chkSound.Value = 1 Then
    chkSound.Caption = "Enabled"
    SoundC = True
ElseIf chkSound.Value = 0 Then
    chkSound.Caption = "Disabled"
    SoundC = False
End If
End Sub

Private Sub cmdApply_Click()
If DisplayC = True Then
    Call WriteINI("DisplayCorrection", "On", "True", App.Path & "\Options.ini")
    DisplayCorr = True
ElseIf DisplayC = False Then
    Call WriteINI("DisplayCorrection", "On", "False", App.Path & "\Options.ini")
    DisplayCorr = False
End If

If SoundC = True Then
    Call WriteINI("Sounds", "Value", "On", App.Path & "\Options.ini")
    PlaySnd = True
ElseIf SoundC = False Then
    Call WriteINI("Sounds", "Value", "Off", App.Path & "\Options.ini")
    PlaySnd = False
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDisplayOpt_Click()
famDisplay.Visible = True
famSound.Visible = False
End Sub

Private Sub cmdoK_Click()
cmdApply_Click
DoEvents
Unload Me
End Sub

Private Sub cmdSoundOpt_Click()
famSound.Visible = True
famDisplay.Visible = False
End Sub

Private Sub Form_Load()
famDisplay.Visible = True
famSound.Visible = False

If DisplayCorr = True Then
    chkLine.Value = 1
End If

If PlaySnd = True Then
    chkSound.Value = 1
End If
End Sub
