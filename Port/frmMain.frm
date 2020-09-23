VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KwikScan"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   45
      TabIndex        =   19
      Top             =   3630
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   1440
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Misc:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      Left            =   3720
      TabIndex        =   12
      Top             =   0
      Width           =   1095
      Begin VB.CommandButton cmdThief 
         Caption         =   "IP-Thief"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   16
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdPortSpy 
         Caption         =   "Port Spy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton save 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton portlist 
         Caption         =   "Port List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   13
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblAbout 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "About"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   3600
         Width           =   855
      End
   End
   Begin MSComctlLib.StatusBar statbar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   3900
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "Idle"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "date$"
            TextSave        =   "date$"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrOut 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2040
      Top             =   2880
   End
   Begin MSWinsockLib.Winsock sck1 
      Index           =   0
      Left            =   960
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox lstResults 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   1785
      ItemData        =   "frmMain.frx":08CA
      Left            =   40
      List            =   "frmMain.frx":08CC
      TabIndex        =   9
      Top             =   1860
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   40
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton cmdClear 
         Appearance      =   0  'Flat
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdScan 
         Appearance      =   0  'Flat
         Caption         =   "Scan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtPortTo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   5
         Text            =   "66000"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtPortFrom 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtIP 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Text            =   "127.0.0.1"
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "KwikScan - Ip Scanner"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label expand 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "to"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "Ports:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "IP Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status:"
      ForeColor       =   &H80000006&
      Height          =   255
      Left            =   40
      TabIndex        =   17
      Top             =   1620
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------
'This is a portscanner coded by SkinCrop
'-----------------------------------------------------------------------------
'This code took me 4 days off and on to finish, I realize its not as advanced
'as most but I think its fast and good coding. I wrote everthing from scratch.
'
'Only borrowed code is in the module wich is the remove dups in listbox wich
'all credit goes to "Source" My favorite part of the program is the port watcher,
'I coded it completely from scratch and made it pretty dynamic as it spawns new
'winsock controls as ports are added.
'
'To contact me please email me at: skincrop@hotmail.com
'
'PLEASE DO NOT RIP MY CODE WITHOUT PROPER CREDIT
'I WORKED HARD TO MAKE THIS CODE PLEASE DONT STEAL IT
'I WILL BE ON THE LOOKOUT FOR DUPLICATE CODE!
'-----------------------------------------------------------------------------
'
'Please take no offence to the aim sniffer code I wrote, it was just for fun
'and to show my port watcher works. I setup the watcher to listen on 1 - 8000
'then tried to direct connect and tried to filesend and other things aim can do.
'I dont guarantee that the code is fully correct or will work always coz in my trials
'it worked 7 out of 16 times. So hopefully my system is just messed up.
'
'-------------------------------------------------------------------------------
Option Explicit
'Declare variables for usage.
Dim port As Long
Dim started As Boolean
Dim openports As Long
Dim stopport As Long
Private Sub cmdClear_Click()
'This clears the results and resets the variables for the scan.
    lstResults.Clear
    openports = 0
    stopport = 0
    statbar.SimpleText = "Idle"
    ProgressBar1.Value = 0
End Sub
Private Sub cmdPortSpy_Click()
'Show the port watch form
    frmportwatch.Show
End Sub
Private Sub cmdScan_Click()
    If started = False Then 'Checks to see if it is running or not to determain to start or stop the scan.
        On Error Resume Next
        lstResults.Clear 'This routine is ran if the scan hasnt been started yet
        lstResults.AddItem "Scanning started on " & Date$ & " at " & Time$
        lstResults.AddItem "----------------------------------------------------------------------"
        txtIP.Enabled = False
        txtPortFrom.Enabled = False
        txtPortTo.Enabled = False
        cmdClear.Enabled = False
        expand.Enabled = False
        cmdScan.Caption = "Cancel"
        started = True
        port = txtPortFrom.Text + 1
        ProgressBar1.Value = 0
        ProgressBar1.Max = txtPortTo.Text
        startScan
    Else
        started = False 'This routine is ran if it has been started.
        txtIP.Enabled = True
        txtPortFrom.Enabled = True
        txtPortTo.Enabled = True
        cmdClear.Enabled = True
        expand.Enabled = True
        cmdScan.Caption = "Scan"
        tmrOut.Enabled = False
        sck1(0).Close
        port = txtPortTo.Text
        lstResults.AddItem "----------------------------------------------------------------------"
        lstResults.AddItem "Scan completed - found " & openports & " open port(s)."
        openports = 0
        statbar.SimpleText = "Idle - Stopped on port " & stopport
        ProgressBar1.Value = 0
    End If
End Sub
Private Sub cmdThief_Click()
frmThief.Show
End Sub
Private Sub expand_Click()
'This is my funny idea for cool special effects, Makes the form wider to
'Show the options ive included.
    If Me.Width = "3795" Then
      frmMain.Width = 4965
    Else
      frmMain.Width = 3795
    End If
End Sub
Private Sub Form_Load()
'Set the width of the form at startup
    frmMain.Width = 3795
End Sub

Private Sub lblAbout_Click()
MsgBox "This program was written by: Michael Gunderson" & vbCrLf & _
        "You may contact me at: skincrop@hotmail.com", vbOKOnly, "About KwikScan"
End Sub
Private Sub portlist_Click()
'Shows the listing of ports
    frmPortList.Show
End Sub
Private Sub save_Click()
'This snipplet of code will open a save dialog and save the listbox results to the
'File of your choice.
    Dim List As Long
    cmdSave.Filter = "Text File (*.txt)|*.txt"
    cmdSave.ShowSave
    If cmdSave.FileTitle = "" Then
        frmMain.Width = 3795
        Exit Sub
    Else
        On Error Resume Next
        Open cmdSave.FileName For Output As #1
        For List& = 0 To lstResults.ListCount - 1
            Print #1, lstResults.List(List&)
            Next List&
        Close #1
        lstResults.AddItem "[ Saved results to: " & cmdSave.FileTitle & " ]"
        frmMain.Width = 3795
    End If
End Sub
Private Sub sck1_Connect(Index As Integer)
On Error Resume Next
'If a winsock is able to connect to a port it will mark it as open in the list,
'Then set the counters to increase by one and move to the next port.
    lstResults.AddItem "Port: " & sck1(0).RemotePort & " - Open"
    port = port + 1
    openports = openports + 1
End Sub
Private Sub startScan()
'If port variable is less than or equal to the end port then continue by
'closing the winsock control, setting the remote IP and port to start on
'Then connect using the timer to space out connection attempts to reduce CPU
'Usage.
On Error Resume Next

Dim intI As Integer
    If port < txtPortTo Then
        ProgressBar1.Value = port
        sck1(0).Close
        sck1(0).RemoteHost = txtIP.Text
        sck1(0).RemotePort = port
        sck1(0).Connect
        tmrOut.Interval = 10
        tmrOut.Enabled = True
        DoEvents
    Else
         'If the from port is higher than the to port then it will
         'jump to the routine ran as if clicking scan to stop the scan.
        Call cmdScan_Click
    End If
    Exit Sub
End Sub
Private Sub sck1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'If the winsock runs into a error on a port it is instructed to ignore
'and continue scanning.
On Error Resume Next
End Sub
Private Sub tmrOut_Timer()
'This is the time that increased the port to scan variable and continue the scan.
    port = port + 1
    stopport = port
    DoEvents
    statbar.SimpleText = "Scanning - Port " & port
    startScan
End Sub
