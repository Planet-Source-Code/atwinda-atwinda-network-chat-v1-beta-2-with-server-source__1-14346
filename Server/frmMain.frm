VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atwinda Chat - Server"
   ClientHeight    =   2040
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6705
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstUnKickable 
      Height          =   255
      Left            =   6960
      TabIndex        =   12
      Top             =   360
      Width           =   135
   End
   Begin MSWinsockLib.Winsock wskServer 
      Index           =   0
      Left            =   3360
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSysTray 
      Caption         =   "&To System Tray"
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "&Start Server"
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Users"
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2055
      Begin VB.ListBox lstUsers 
         Height          =   1425
         IntegralHeight  =   0   'False
         ItemData        =   "frmMain.frx":0CCA
         Left            =   120
         List            =   "frmMain.frx":0CCC
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Server Info"
      Height          =   1335
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.Label lblStatus 
         Height          =   255
         Left            =   1360
         TabIndex        =   6
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Server Status:"
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
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblIpPort 
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Server IP + Port:"
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
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblCount 
         Height          =   255
         Left            =   1155
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "User Count:"
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
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version:"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Menu mnu 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuGetIP 
         Caption         =   "GetIP"
      End
      Begin VB.Menu mnuKick 
         Caption         =   "Kick User"
      End
      Begin VB.Menu mnuUsrIndex 
         Caption         =   "Users Indexs"
      End
   End
   Begin VB.Menu mnu2 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Show Server"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Num As Long
Dim StartTime As String
Dim strUsers() As String
Const SvrPort = "1290"

Private Sub cmdListen_Click()
On Error Resume Next
If cmdListen.Caption = "&Start Server" Then
    wskServer(0).Close
    wskServer(0).Listen
    lblStatus.Caption = "Waiting for Connections..."
    cmdListen.Caption = "&Close Server"
ElseIf cmdListen.Caption = "&Close Server" Then
    For i = 0 To wskServer().UBound
        wskServer(i).Close
        strUsers(i) = ""
    Next i
    lblStatus.Caption = "Closed..."
    cmdListen.Caption = "&Start Server"
    lstUsers.Clear
    Num = "0"
    lblCount.Caption = Num
End If
End Sub

Private Sub cmdSysTray_Click()
Me.Hide
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
    Call MsgBox("Another instance of Atwinda Network" & vbCrLf & "Chat Server Is already running." & vbCrLf & "This instance will be closed.", vbExclamation, "App already open")
    End
End If
    
wskServer(0).LocalPort = SvrPort
lblIpPort.Caption = wskServer(0).LocalIP & ":" & wskServer(0).LocalPort
Num = "0"
lblCount.Caption = Num
lblStatus.Caption = "Closed..."
StartTime = Time
lblVersion.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
Call AddIcon(Me, "Click to restore Server window")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call DeleteIcon
DoEvents
End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim Result As Long, msg As Long
If Me.ScaleMode = vbPixels Then
    msg = X
Else
    msg = X / Screen.TwipsPerPixelX
End If
    
Select Case msg
Case 517
Me.PopupMenu mnu2
        
Case 514
Result = SetForegroundWindow(Me.hwnd)
Me.Show
            
End Select
End Sub

Private Sub lstUsers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lstUsers.SelCount = "0" Then
Exit Sub
Else
    If Button = 1 Then
        Exit Sub
    Else
        Call Me.PopupMenu(mnu, , lstUsers.Left + 160)
    End If
End If
End Sub

Private Sub mnuGetIP_Click()
Call MsgBox(lstUsers.Text & "'s IP is: " & GetIP(lstUsers.Text), vbInformation, "User's IP")
End Sub

Private Sub mnuKick_Click()
Call KickUser(Trim(lstUsers.Text), "Server")
End Sub

Private Sub mnuShow_Click()
Me.Show
End Sub

Private Sub mnuUsrIndex_Click()
Dim strList As String
For i = LBound(strUsers()) To UBound(strUsers())
    If Not strUsers(i) = "" Then
        strList = strList & strUsers(i) & "(" & i & ")" & vbCrLf
        DoEvents
    End If
Next i
MsgBox strList, vbInformation, "Users Index's"
End Sub

Private Sub wskServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
Dim lngIndex As Long, blnFlag As Boolean
   
For lngIndex& = 1 To wskServer().UBound
    If wskServer(lngIndex&).State = sckClosed Then
        blnFlag = True
        Exit For
    End If
Next lngIndex&
      
If blnFlag = False Then
    lngIndex& = wskServer().UBound + 1
    Load wskServer(lngIndex&)
    ReDim Preserve strUsers(lngIndex&) As String
    strUsers(lngIndex&) = ""
End If

Call wskServer(lngIndex&).Accept(requestID&)

End Sub

Private Sub wskServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strData As String

wskServer(Index).GetData strData$

Call DataParsing(strData, Index)
DoEvents
End Sub

Function DataParsing(strData As String, Index As Integer)
On Error Resume Next
Dim Command As String, Info As String, i As Integer

Command$ = Left$(strData$, InStr(strData$, " ") - 1)
Info$ = Right$(strData$, Len(strData$) - InStr(strData$, " "))

Select Case Command

Case "Message":
    For i = 1 To wskServer().UBound
        Call wskServer(i).SendData("Message " & Info)
        DoEvents
    Next i
    
Case "PMMessage":
    Dim User As String, UserText As String, UserE As Boolean
    User$ = Left$(Info$, InStr(Info$, ":") - 1)
    UserText$ = Right$(Info$, Len(Info$) - InStrRev(Info$, ":"))
    DoEvents
    
    For i = LBound(strUsers()) To UBound(strUsers())
        If strUsers(i) = User Then
            UserE = True
            Call wskServer(i).SendData("PMMessage " & UserText)
            DoEvents
            Call wskServer(Index).SendData("PMMessage " & UserText)
            DoEvents
        End If
    Next i

    If Not UserE = True Then
        Call wskServer(Index).SendData("PMError " & User)
        DoEvents
    End If
    
Case "Join":
    If CheckUsername(Info) = True Then
        Call wskServer(Index).SendData("ErrUsername " & Info)
        DoEvents
        Exit Function
    End If
    
    Num = Num + 1
    lblCount.Caption = Num
    
    Call lstUsers.AddItem(Info)
    strUsers(Index) = Info
    DoEvents
    
    For i = 0 To lstUsers.ListCount
        If lstUsers.List(i) <> "" Then
            Call wskServer(Index).SendData("UserList " & lstUsers.List(i))
            DoEvents
        End If
        DoEvents
    Next i
    
    For i = 1 To wskServer().UBound
        If Not i = Index Then
            Call wskServer(i).SendData("Joined " & Info)
            DoEvents
        End If
        DoEvents
    Next i
    
Case "Leave":
    For i = 1 To wskServer().UBound
        If Not i = Index Then
            Call wskServer(i).SendData("Left " & Info)
            DoEvents
        End If
        DoEvents
    Next i
    
    For i = 0 To lstUsers.ListCount
        If lstUsers.List(i) = Info Then
            Call lstUsers.RemoveItem(i)
            DoEvents
        End If
        DoEvents
    Next i
    DoEvents
    
    If CheckUnKick(Info) = True Then
        For i = 0 To lstUnKickable.ListCount
            If lstUnKickable.List(i) = Info Then
                Call lstUnKickable.RemoveItem(i)
                DoEvents
            End If
        Next i
    End If
    
    strUsers(Index) = ""
    
    Num = Num - 1
    If Num < "0" Then Num = "0"
    lblCount.Caption = Num

    wskServer(Index).Close
    
Case "Prompt":
    Call Prompt(Info$, Index)

End Select
End Function

Function Prompt(Info As String, Index As Integer)
    Dim Command As String, Data As String
    Command$ = Left$(Info$, InStr(Info$, ":") - 1)
    Data$ = Right$(Info$, Len(Info$) - InStr(Info$, ":"))
    DoEvents
    
    If Command = "KickUser" Then
        Call KickUser(Data, strUsers(Index))
        DoEvents
    ElseIf Command = "ServerInfo" Then
        Dim strText As String
        strText = "Server IP/Port: " & lblIpPort.Caption & vbCrLf
        strText = strText & "User Count: " & lblCount.Caption & vbCrLf
        strText = strText & "Server Vesion: " & lblVersion.Caption & vbCrLf
        strText = strText & "Server Start Time: " & StartTime
        DoEvents
        Call wskServer(Index).SendData("PMMessage " & "Server|" & strText)
        DoEvents
    ElseIf Command = "GetIP" Then
        Call wskServer(Index).SendData("PMMessage " & "Server|" & GetIP(Data))
        DoEvents
    ElseIf Command = "MakemeUnKickable" Then
        Call lstUnKickable.AddItem(Data)
        DoEvents
        Call wskServer(Index).SendData("PMMessage " & "Server|" & "You are unkickable")
        DoEvents
    ElseIf Command = "Help" Then
        Dim sText As String
        sText = "KickUser:Username" & vbCrLf
        sText = sText & "ServerInfo:username" & vbCrLf
        sText = sText & "GetIP:username" & vbCrLf
        sText = sText & "Help:username"
        DoEvents
        Call wskServer(Index).SendData("PMMessage " & "Server|" & sText)
        DoEvents
    End If
End Function

Function GetIP(Username As String) As String
For i = 0 To UBound(strUsers())
    If strUsers(i) = Username Then
    GetIP = wskServer(i).RemoteHostIP
    End If
Next i
If GetIP = "" Then
GetIP = "None"
End If
End Function

Function KickUser(Username As String, By As String)
On Error Resume Next
Dim i As Integer
For i = 0 To UBound(strUsers())
    
    If CheckUnKick(Username) = True Then
        If Not By = "Server" Then
            wskServer(UIndex(By)).SendData ("PMMessage " & "Server|" & "User " & Username & "is unkickable")
            DoEvents
            Exit Function
        End If
    End If
    
    If strUsers(i) = Username Then
        Call wskServer(i).SendData("UKicked " & By)
        DoEvents
        Call wskServer(i).Close
        DoEvents
            
    For z = 1 To wskServer().UBound
        If Not i = Index Then
            Call wskServer(z).SendData("Kicked " & Username)
            DoEvents
        End If
        DoEvents
    Next z
    DoEvents
    
    For z = 0 To lstUsers.ListCount
        If lstUsers.List(z) = Username Then
            Call lstUsers.RemoveItem(z)
            DoEvents
        End If
        DoEvents
    Next z
    DoEvents
    
    If CheckUnKick(Username) = True Then
        For a = 0 To lstUnKickable.ListCount
            If lstUnKickable.List(a) = Username Then
                Call lstUnKickable.RemoveItem(i)
                DoEvents
            End If
        Next a
    End If
    
    strUsers(i) = ""
    
    Num = Num - 1
    If Num < "0" Then Num = "0"
    lblCount.Caption = Num
        
    End If
Next i

End Function

Function UIndex(Username As String) As Integer
Dim i As Integer
For i = o To UBound(strUsers())
    If strUsers(i) = Username Then
        UIndex = i
        Exit Function
    End If
Next i
End Function

Function CheckUnKick(Username As String) As Boolean
For i = 0 To lstUnKickable.ListCount
If lstUnKickable.List(i) = Username Then
    CheckUnKick = True
    Exit Function
End If
Next i
CheckUnKick = False
End Function

Function CheckUsername(Username As String) As Boolean
For i = 0 To lstUsers.ListCount
    If lstUsers.List(i) = Username Then
        CheckUsername = True
        Exit Function
    End If
Next i
CheckUsername = False
End Function
