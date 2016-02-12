VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form network 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Network Communication"
   ClientHeight    =   5085
   ClientLeft      =   10050
   ClientTop       =   2910
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.Frame local_remote_frame 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   495
         Left            =   480
         TabIndex        =   10
         Top             =   1200
         Width           =   4815
         Begin VB.OptionButton Local_Option 
            Caption         =   "LOCAL: Connected directly to a Porometer"
            Height          =   195
            Left            =   0
            TabIndex        =   1
            Top             =   0
            Value           =   -1  'True
            Width           =   3975
         End
         Begin VB.OptionButton Remote_Option 
            Caption         =   "REMOTE: Will connect through another computer"
            Height          =   195
            Left            =   0
            TabIndex        =   2
            Top             =   240
            Width           =   4335
         End
      End
      Begin VB.Frame client_server_frame 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   495
         Left            =   480
         TabIndex        =   9
         Top             =   1800
         Width           =   4815
         Begin VB.OptionButton Server_Option 
            Caption         =   "SERVER: Waits for client computer to initiate connection"
            Height          =   195
            Left            =   0
            TabIndex        =   4
            Top             =   240
            Width           =   4695
         End
         Begin VB.OptionButton Client_Option 
            Caption         =   "CLIENT: Initiates connection to a server"
            Height          =   195
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Value           =   -1  'True
            Width           =   3975
         End
      End
      Begin VB.TextBox port_text 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Text            =   "23"
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox server_ip_text 
         Height          =   285
         Left            =   3600
         TabIndex        =   6
         Text            =   "0.0.0.0"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton go_button 
         Caption         =   "Connect To Server"
         Height          =   615
         Left            =   360
         TabIndex        =   7
         Top             =   2880
         Width           =   3015
      End
      Begin VB.CommandButton stop_button 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   615
         Left            =   3600
         TabIndex        =   8
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   5040
         Top             =   4320
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   5040
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Caption         =   "This will hold the local IP address"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label settings_status_label 
         Caption         =   "This computer is a LOCAL CLIENT"
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Port:"
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label server_ip_label 
         Alignment       =   1  'Right Justify
         Caption         =   "Server IP address:"
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label connection_status_label 
         Caption         =   "Connection Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3600
         Width           =   5295
      End
      Begin VB.Label response_label 
         Caption         =   "no response"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   3960
         Width           =   5295
      End
      Begin VB.Label data_sent_label 
         Caption         =   "no data sent"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   4320
         Width           =   5415
      End
   End
End
Attribute VB_Name = "network"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private waiting_for_local_to_connect As Boolean

Private Sub Client_Option_Click()
update_status_line
End Sub

Private Sub Form_Load()
Label1.Caption = "Local IP Address is: " & Winsock1.LocalIP
waiting_for_local_to_connect = False
'edc 12-11-06 alter border color and caption
'Me.Caption = Me.Caption & "    " & SubCaption
Me.BackColor = lngBorderColor
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
If network_connected Then
    cancel = 1
    Hide
End If
End Sub

Private Sub go_button_Click()
If Server_Option.value Then
    Winsock1.LocalPort = port_text.Text
    Winsock1.Listen
    If Winsock1.state <> sckListening Then Exit Sub
Else
    Winsock1.RemotePort = port_text.Text
    Winsock1.RemoteHost = server_ip_text.Text
    On Error Resume Next
    Winsock1.Connect
    If Err.Number <> 0 Then
        MsgBox "Error trying to connect"
        Winsock1.Close
        Exit Sub
    End If
    On Error GoTo 0
End If
enable_stuff False
waiting_for_local_to_connect = Remote_Option.value
End Sub

Private Sub Local_Option_Click()
update_status_line
End Sub

Private Sub Remote_Option_Click()
update_status_line
End Sub

Private Sub Server_Option_Click()
update_status_line
End Sub

Private Sub update_status_line()
settings_status_label.Caption = "This computer is a " + _
    IIf(Local_Option.value, "LOCAL", "REMOTE") + " " + _
    IIf(Client_Option.value, "CLIENT", "SERVER")
server_ip_label.Visible = Client_Option.value
server_ip_text.Visible = Client_Option.value
go_button.Caption = IIf(Client_Option.value, "Connect To Server", "Start Listening for Connection")
End Sub

Private Sub enable_stuff(enable As Boolean)
local_remote_frame.Enabled = enable
client_server_frame.Enabled = enable
go_button.Enabled = enable
server_ip_text.Enabled = enable
port_text.Enabled = enable
go_button.Enabled = enable
stop_button.Enabled = Not enable
End Sub

Private Sub stop_button_Click()
Winsock1.Close
enable_stuff True
network_connected = False
End Sub

Private Sub Timer1_Timer()
Dim a$
If Winsock1.state = sckClosed Then
    a$ = "Closed"
ElseIf Winsock1.state = sckOpen Then
    a$ = "Open"
ElseIf Winsock1.state = sckListening Then
    a$ = "Listening"
ElseIf Winsock1.state = sckConnectionPending Then
    a$ = "Connection Pending"
ElseIf Winsock1.state = sckResolvingHost Then
    a$ = "Resolving Host"
ElseIf Winsock1.state = sckHostResolved Then
    a$ = "Host Resolved"
ElseIf Winsock1.state = sckConnecting Then
    a$ = "Connecting"
ElseIf Winsock1.state = sckConnected Then
    a$ = "Connected"
ElseIf Winsock1.state = sckClosing Then
    a$ = "Closing"
ElseIf Winsock1.state = sckError Then
    a$ = "Error"
Else
    a$ = "Unknown #" + str$(Winsock1.state)
End If
connection_status_label.Caption = "Connection Status: " & a$
network_connected = (Winsock1.state = sckConnected)
If network_connected And waiting_for_local_to_connect Then
    waiting_for_local_to_connect = False
    Me.Hide
End If
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
If Winsock1.state = sckListening Then
    Winsock1.Close
    Winsock1.Accept requestID
End If
End Sub

'Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'Dim a$, db$
'Dim L(5) As Long
'Dim i As Integer
'Dim j As Integer
'Dim r As Long
'Winsock1.GetData a$
'response_label.Caption = a$
'If Local_Option.value Then
'    ' we are a local computer, so we need to send this command to the actual machine and get a
'    ' response
'    If ComLoc% = -1 And TitleScrn.MainComm.PortOpen = False Then
'        TitleScrn.MainComm.PortOpen = True
'    End If
'    ' parse a$ as 5 numeric values with commas in between
'    j = 1
'    For i = 1 To 5
'        L(i) = val(Mid$(a$, j))
'        j = InStr(j, a$, ",") + 1
'    Next i
'    If L(1) <= 255 Then
'        db$ = Chr$(L(1))
'    ElseIf L(1) < 256& * 256 Then
'        db$ = Chr$(L(1) And 255) & Chr$(L(1) \ 256)
'    Else
'        db$ = Chr$(L(1) And 255) & Chr$((L(1) \ 256) And 255) & Chr$(L(1) \ 256 \ 256)
'    End If
'    r = RSEcho_New(db$, CByte(L(2)), Chr$(L(3)), CByte(L(4)), L(5))
'    ' during the above command, it is possible that the network connection could have been broken
'    ' and we don't want to try to send data over a broken connection, do we?
'    If network_connected Then
'        Winsock1.SendData str$(r) & vbCrLf
'        data_sent_label.Caption = str$(r)
'    End If
'Else
'    ' we have been waiting for this information
'    abort_wait = True
'End If
'End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "Connection error: " & Description
End Sub
