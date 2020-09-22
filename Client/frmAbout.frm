VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Atwinda Chat"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdoK 
      Caption         =   "OK"
      Height          =   255
      Left            =   5640
      TabIndex        =   1
      Top             =   2160
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.Label Label5 
         Caption         =   $"frmAbout.frx":0CCA
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "Atwinda Software © 2001"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label lblWebsite 
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Website:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblVersion 
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Version:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblAuthor 
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Author:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   120
      Picture         =   "frmAbout.frx":0D57
      Top             =   120
      Width           =   2850
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdoK_Click()
Unload Me
End Sub

Private Sub Form_Load()
lblAuthor.Caption = "Andy Stagg"
lblVersion = App.Major & "." & App.Minor & "." & App.Revision
lblWebsite.Caption = "http://atwindasoft.myqth.com"
End Sub
