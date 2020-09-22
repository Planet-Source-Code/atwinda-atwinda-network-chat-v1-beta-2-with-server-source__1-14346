VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChat 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4680
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8145
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrConnection 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   4800
   End
   Begin VB.TextBox txtText 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   7935
   End
   Begin MSWinsockLib.Winsock wskClient 
      Left            =   120
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox lstUsers 
      Height          =   3375
      Left            =   6120
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin RichTextLib.RichTextBox txtDisplay 
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5953
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmChat.frx":0CCA
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   7935
      Begin VB.CommandButton cmdSend 
         Caption         =   "&Send"
         Height          =   375
         Left            =   6720
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblStatus 
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   290
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   290
         Width           =   615
      End
   End
   Begin RichTextLib.RichTextBox rt 
      Height          =   255
      Left            =   4320
      TabIndex        =   9
      Top             =   4200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmChat.frx":0D4C
   End
   Begin RichTextLib.RichTextBox rt2 
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   4080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmChat.frx":0DCE
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuDiscon 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuPrompt 
         Caption         =   "Prompt"
         Visible         =   0   'False
      End
      Begin VB.Menu spacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuAbout0 
      Caption         =   "About"
      Begin VB.Menu mnuAAC 
         Caption         =   "About Atwinda Chat"
      End
   End
   Begin VB.Menu mnu 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuPM 
         Caption         =   "Private Message"
      End
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nTime As Long
Private SvrIPAdd As String

Function Connect(svrIP As String)
wskClient.Close
wskClient.Connect svrIP, "1290"
SvrIPAdd = svrIP
lblStatus.Caption = "Connecting..."
End Function

Private Sub cmdSend_Click()
If cmdSend.Enabled = True Then
If wskClient.State = sckConnected Then
    If txtText.Text <> "" Then
        Dim allText As String
        allText = txtUser.Text & ":" & txtText.Text
        DoEvents
        Call wskClient.SendData("Message " & allText)
        If PlaySnd = True Then
        Call PlayWav("Send.wav")
        End If
    Else
        Call AddText("  ~// You must enter text to send it")
    End If
Else
    Call AddText("  ~// You must be connected to some to send text")
End If
txtText.Text = ""
End If
End Sub

Private Sub Form_Load()
nTime = "0"
tmrConnection.Enabled = True
If Prompt = True Then
    mnuPrompt.Visible = True
End If
cmdSend.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If wskClient.State = sckConnected Then
Call wskClient.SendData("Leave " & txtUser.Text)
DoEvents
End If
End Sub

Private Sub lstUsers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lstUsers.SelCount = "0" Then
Exit Sub
Else
If Button = 1 Then
    Exit Sub
Else
    Call Me.PopupMenu(mnu, , lstUsers.Left + 50)
End If
End If
End Sub

Private Sub mnuAAC_Click()
frmAbout.Show
End Sub

Private Sub mnuDiscon_Click()
Call wskClient.SendData("Leave " & txtUser.Text)
DoEvents
wskClient.Close
DoEvents
frmSignon.txtUser.Text = txtUser.Text
frmSignon.Show
DoEvents
Unload Me
End Sub

Private Sub mnuExit_Click()
If wskClient.State = sckConnected Then
Call wskClient.SendData("Leave " & txtUser.Text)
DoEvents
End If
End
End Sub

Private Sub mnuOpt_Click()
frmOptions.Show
End Sub

Private Sub mnuPM_Click()
If lstUsers.Text = txtUser.Text Then
    Call AddText("  ~// You can't PM yourself")
Else
frmPM.lblUserPM.Caption = lstUsers.Text
frmPM.Show
End If
End Sub

Private Sub mnuPrompt_Click()
frmPrompt.Show
End Sub

Private Sub txtText_Change()
If txtText.Text > "" Then
    cmdSend.Enabled = True
Else
    cmdSend.Enabled = False
End If
End Sub

Private Sub txtText_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSend_Click
    KeyAscii = 0
End If
End Sub

Public Sub wskClient_Close()
lblStatus.Caption = "Closed"
Call AddText("  ~// The connection was unexpectedly droped.")
lstUsers.Clear
txtDisplay.Text = ""
End Sub

Private Sub wskClient_Connect()
lblStatus.Caption = "Connected..."
Call wskClient.SendData("Join " & txtUser.Text)
End Sub

Private Sub wskClient_DataArrival(ByVal bytesTotal As Long)
Dim strData As String

wskClient.GetData strData$

Call DataParsing(strData)
End Sub

Function DataParsing(strData As String)
On Error Resume Next
Dim Command As String, Info As String, User As String, Text As String

Command$ = Left$(strData$, InStr(strData$, " ") - 1)
Info$ = Right$(strData$, Len(strData$) - InStr(strData$, " "))

Select Case Command

Case "Message":
    User$ = Left$(Info$, InStr(Info$, ":") - 1)
    Text$ = Right$(Info$, Len(Info$) - InStr(Info$, ":"))
    DoEvents
    
    Call AddChat(User, Text)
    DoEvents
    If Not User = txtUser.Text Then
        If PlaySnd = True Then
            Call PlayWav("Recv.wav")
        End If
    End If
    
Case "PMMessage":
    User$ = Left$(Info$, InStr(Info$, "|") - 1)
    Text$ = Right$(Info$, Len(Info$) - InStr(Info$, "|"))
    DoEvents
    
    Call AddChatPM(User, Text)
    DoEvents
    
Case "ErrUsername":
    Dim Answer As VbMsgBoxResult
    Answer = MsgBox("The username " & Info$ & " is in use." & vbCrLf & "Would you like to enter a new one.", vbYesNo, "Username Error")
        If Answer = vbYes Then
            Dim strUser As String
            strUser = InputBox("Please enter a new username.", "New Username")
            If strUser <> "" Then
            Call wskClient.SendData("Join " & strUser)
            DoEvents
            txtUser.Text = strUser
            Me.Caption = "Atwinda Chat - (" & txtUser.Text & ")"
            End If
        ElseIf Answer = vbNo Then
            wskClient.Close
            frmSignon.Show
            DoEvents
            Unload Me
        End If

Case "PMError":
    Call AddText("  ~// User: " & Info & " doesn't exist, or has left")
    DoEvents

Case "UserList":
    Call lstUsers.AddItem(Info)
    DoEvents
    
Case "Joined":
    Call lstUsers.AddItem(Info)
    DoEvents
    Call AddText("  ~// User: " & Info & " has joined the chat")
    
Case "Left":
    For i = 0 To lstUsers.ListCount
        If lstUsers.List(i) = Info Then
            Call lstUsers.RemoveItem(i)
            DoEvents
        End If
        DoEvents
    Next i
    
    Call AddText("  ~// User: " & Info & " left the chat room")
    
Case "Kicked":
    For i = 0 To lstUsers.ListCount
        If lstUsers.List(i) = Info Then
            Call lstUsers.RemoveItem(i)
            DoEvents
        End If
        DoEvents
    Next i
    
    Call AddText("  ~// User: " & Info & " was kicked")
    
Case "UKicked":
    Call AddText("  ~// You have been kicked by: " & Info)
   
End Select
End Function

Function AddText(Text As String)

rt2.SelStart = 0
rt2.SelLength = 0
rt2.TextRTF = Text
rt2.SelStart = 2
rt2.SelLength = Len(Text)
rt2.SelColor = &H8000&
rt2.SelStart = 0
rt2.SelLength = 0
DoEvents
txtDisplay.SelStart = Len(txtDisplay)
DoEvents
txtDisplay.SelRTF = rt2.TextRTF & vbCrLf
If DisplayCorr = True Then
    txtDisplay.SelStart = Len(txtDisplay)
    txtDisplay.SelRTF = vbCrLf
End If
DoEvents
txtDisplay.SelStart = Len(txtDisplay)

rt2.TextRTF = ""
End Function

Function AddChat(User As String, Text As String)
rt.SelBold = False
rt.SelColor = vbBlack

rt.SelStart = 0
rt.SelLength = 0
rt.SelRTF = User & ": " & Text
rt.SelStart = 0
rt.SelLength = Len(User) + 1
If User = txtUser.Text Then
rt.SelColor = vbRed
Else
rt.SelColor = vbBlue
End If
rt.SelBold = True
rt.SelStart = 0
rt.SelLength = 0
DoEvents
txtDisplay.SelStart = Len(txtDisplay)
DoEvents
txtDisplay.SelRTF = rt.TextRTF & vbCrLf

If DisplayCorr = True Then
    txtDisplay.SelStart = Len(txtDisplay)
    txtDisplay.SelRTF = vbCrLf
End If

DoEvents
txtDisplay.SelStart = Len(txtDisplay)

rt.TextRTF = ""
End Function

Function AddChatPM(User As String, Text As String)
rt.SelBold = False
rt.SelColor = vbBlack

rt.SelStart = 0
rt.SelLength = 0
rt.SelRTF = User & ": " & Text
rt.SelStart = 0
rt.SelLength = Len(User) + 1
rt.SelColor = &H8000&
rt.SelBold = True
rt.SelStart = 0
rt.SelLength = 0
DoEvents
txtDisplay.SelStart = Len(txtDisplay)
DoEvents
txtDisplay.SelRTF = rt.TextRTF & vbCrLf
If DisplayCorr = True Then
    txtDisplay.SelStart = Len(txtDisplay)
    txtDisplay.SelRTF = vbCrLf
End If
DoEvents
txtDisplay.SelStart = Len(txtDisplay)

rt.TextRTF = ""
End Function
