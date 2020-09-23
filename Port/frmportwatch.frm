VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmportwatch 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Port Spy"
   ClientHeight    =   4125
   ClientLeft      =   9600
   ClientTop       =   4125
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   3720
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   30
      TabIndex        =   6
      Top             =   1320
      Width           =   3945
      Begin VB.ListBox StatusBox 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000001&
         Height          =   2175
         ItemData        =   "frmportwatch.frx":0000
         Left            =   45
         List            =   "frmportwatch.frx":0002
         TabIndex        =   8
         Top             =   120
         Width           =   3855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3600
         TabIndex        =   10
         Top             =   2400
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   3945
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   9
         Text            =   "Watching Ports:"
         Top             =   210
         Width           =   1455
      End
      Begin VB.TextBox txtPort 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   285
         Left            =   50
         TabIndex        =   4
         Top             =   450
         Width           =   1260
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "<<"
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   ">>"
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   480
         Width           =   375
      End
      Begin VB.ListBox lstWatchPorts 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         ItemData        =   "frmportwatch.frx":0004
         Left            =   2040
         List            =   "frmportwatch.frx":0006
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Port to add:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   50
         TabIndex        =   5
         Top             =   210
         Width           =   975
      End
   End
   Begin MSWinsockLib.Winsock sckListen 
      Index           =   0
      Left            =   1920
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmportwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declare the variable for use
Dim prts As Integer
Private Sub cmdAdd_Click()
On Error Resume Next
'If the port to add is blank then do nothing
    If txtPort.Text = "" Then
        Exit Sub
    Else
'Add the port to the list box then kill duplicates so you cant add two of them.
        lstWatchPorts.AddItem txtPort.Text
        ListKillDupes lstWatchPorts
        GoTo makelisten
    End If
    Exit Sub

'This is the routine that spawns new winsock per port to watch
makelisten:
    prts = txtPort.Text
    Load sckListen(prts)
    sckListen(prts).LocalPort = prts
    sckListen(prts).Listen
    txtPort.Text = ""
    Exit Sub
End Sub
Private Sub cmdClose_Click()
'When close is pressed show the main window and unload this form.
    frmMain.Show
    Unload Me
End Sub
Private Sub cmdDel_Click()
'Declare the variable onport for use as integer (numbers)
Dim onport As Integer
    If lstWatchPorts.ListIndex = -1 Then
        cmdDel.Enabled = False
        Exit Sub
    Else
'This routine find wich port the user wants to delete and stops the winsock
'It then deletes the item from the list.
        onport = lstWatchPorts.Text
        sckListen(onport).Close
        lstWatchPorts.RemoveItem lstWatchPorts.ListIndex
        cmdDel.Enabled = False
    End If
End Sub
Private Sub Form_Load()
'When loading reset the main windows width and change the expand caption then hide it.
    frmMain.Width = 3795
    frmMain.Hide
End Sub
Private Sub Form_Unload(Cancel As Integer)
'Show the main windows again then uload the port watcher.
    frmMain.Show
    sckListen(prts).Close
    Unload Me
End Sub
Private Sub Label3_Click()
'Message explaining how to function the port scanner
MsgBox "Usage: To begin watching a port just type the port in the Port to add: box and click >>. Once pressed monitoring begins on that port. To stop monitoring just click the one to stop in the list and click <<.", vbOKOnly, "Port Watcher Help"
End Sub
Private Sub lstWatchPorts_Click()
'Dont enable the delete button unless a item is clicked
    cmdDel.Enabled = True
End Sub
Private Sub sckListen_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
'On a connection then accept it send a message.
    If sckListen(Index).State <> sckClosed Then sckListen(Index).Close
    
    sckListen(Index).Accept requestID
    
    sckListen(Index).SendData "Logged IP: " & sckListen(Index).RemoteHostIP & vbCrLf & vbCrLf & _
        "This computer is being monitored for all connections by KwikScan" & vbCrLf & vbCrLf & _
        "Your IP address has been recorded to the system logs." & vbCrLf & vbCrLf & _
        "Please discontinue all attempts to access this computer or legal actions may be taken." & vbCrLf & vbCrLf & _
        "This connect will now be terminated." & vbCrLf
    
    StatusBox.AddItem sckListen(Index).RemoteHostIP & " on port " & Index & ": " & Time$ & " / " & Date$
End Sub
Private Sub sckListen_SendComplete(Index As Integer)
'After message was sent in connection request, close the winsock to disconnect the
'user.
    sckListen(Index).Close
    sckListen(Index).Listen
End Sub
