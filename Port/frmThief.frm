VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmThief 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "IP-Thief"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock file 
      Left            =   1680
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Connect 
      Left            =   1200
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin MSWinsockLib.Winsock ThiefSck 
      Left            =   2880
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   80
   End
   Begin VB.TextBox txtYourIP 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   180
      Width           =   2175
   End
   Begin VB.CommandButton cmdStartStop 
      Caption         =   "Start"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   975
   End
   Begin VB.ListBox lstIPs 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   1005
      ItemData        =   "frmThief.frx":0000
      Left            =   40
      List            =   "frmThief.frx":0002
      TabIndex        =   0
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Your IP:"
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
      Width           =   735
   End
End
Attribute VB_Name = "frmThief"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This IP Thief I wrote real quick just for users wanting a IP from AIM or something.
'Nothing special just listen's on port 80 for http requests then list's the IP of
'the person trying to connect to it.

'Declare variables for usage in subs
Dim listenip As Long
Dim listenport As Integer
Dim onoff As Boolean
Private Sub cmdCancel_Click()
'On close close the socket and show the main form and unload this one.
    ThiefSck.Close
    frmMain.Show
    Unload Me
End Sub
Private Sub cmdStartStop_Click()
'On error keep going
On Error Resume Next
'This if statement determains if the program needs to start or stop the IP Thief
    If onoff = False Then
        cmdStartStop.Caption = "Stop"
        onoff = True
        startIPThief
    Else
        cmdStartStop.Caption = "Start"
        onoff = False
        stopIPThief
    End If
End Sub
Private Sub startIPThief()
'Set the port to listen on then start the listening
    Connect.Close
    Connect.LocalPort = 4443
    Connect.Listen
    file.Close
    file.LocalPort = 5190
    file.Listen
    ThiefSck.Close
    ThiefSck.LocalPort = 80
    ThiefSck.Listen
End Sub
Private Sub stopIPThief()
'Close the socket
    Unload Connect
    Unload file
    Unload ThiefSck
End Sub
Private Sub Connect_ConnectionRequest(ByVal requestID As Long)
'On connection for the direct connect in aim add the IP to the listbox.
    lstIPs.AddItem Connect.RemoteHostIP
End Sub
Private Sub file_ConnectionRequest(ByVal requestID As Long)
'When sending a file through aim it will add the IP to the listbox.
    lstIPs.AddItem file.RemoteHostIP
End Sub
Private Sub Form_Load()
'Resize the main form then hide it so this form has focus
    frmMain.Width = 3795
    frmMain.Hide
    txtYourIP.Text = ThiefSck.LocalIP
End Sub
Private Sub Form_Unload(Cancel As Integer)
'Make sure that the main form is shown
    frmMain.Show
End Sub
Private Sub lblHelp_Click()
        MsgBox "To use this just send the person a weblink with your IP as the target." & vbCrLf _
        & "Or if on aim just direct connect to them or send a file.", vbOKOnly, "IP Thief"
End Sub
Private Sub ThiefSck_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
'As a connection comes in close the socket and request the ID then add teh remote IP
'To the list box.
    ThiefSck.Close
    ThiefSck.Accept requestID
    lstIPs.AddItem ThiefSck.RemoteHostIP
End Sub
Private Sub ThiefSck_DataArrival(ByVal bytesTotal As Long)
'Once requestID has been done we can send our data, wich is a customer 404 page.
    ThiefSck.SendData "<html>"
    ThiefSck.SendData "<head>"
    ThiefSck.SendData "<Title>404 Error</title>"
    ThiefSck.SendData "</head>"
    ThiefSck.SendData "<body>"
    ThiefSck.SendData "<center><h1>404 - Error</h1></center><br>"
    ThiefSck.SendData "<center>The file you have requested cannot be found on the server.<br>"
    ThiefSck.SendData "Please contact the system administrator immediatly.<br><br>"
    ThiefSck.SendData "<font size=1>Byte#: 111246686</font></center>"
End Sub
Private Sub ThiefSck_SendComplete()
'Webpage has been sent so close the winsock control, then reopen it.
    ThiefSck.Close
    ThiefSck.Listen
End Sub
