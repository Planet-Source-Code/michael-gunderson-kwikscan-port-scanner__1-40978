VERSION 5.00
Begin VB.Form frmPortList 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Port List"
   ClientHeight    =   4260
   ClientLeft      =   5295
   ClientTop       =   3735
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   40
      TabIndex        =   2
      Top             =   0
      Width           =   5895
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   5280
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   6
         Top             =   200
         Width           =   495
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   410
         Width           =   855
      End
      Begin VB.TextBox txtPort 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   370
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Port Number:"
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
         Top             =   450
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   3960
      Width           =   1695
   End
   Begin VB.ListBox lstPorts 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000001&
      Height          =   3150
      ItemData        =   "Form1.frx":04F5
      Left            =   40
      List            =   "Form1.frx":04F7
      TabIndex        =   0
      Top             =   740
      Width           =   5895
   End
End
Attribute VB_Name = "frmPortList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
'Clear the port, reset the width for the main window and change the expand caption.
    lstPorts.Clear
    frmMain.Width = 3795
'Unload PortList
    Unload Me
End Sub
Private Sub cmdSearch_Click()
'Declare returnval as a variable for usage.
Dim returnval As Long
'Set returnval's value by running the scan function
    returnval = FindInList(lstPorts, txtPort.Text)
'If returnval equals a valid port then highlight the port in the listbox
    If returnval >= 0 Then
        lstPorts.ListIndex = returnval
    Else
'If it doesnt find a valid port then display a messagebox
        MsgBox "Port: " & txtPort.Text & " cannot be found!", vbOKOnly, "Search"
    End If
End Sub
Private Sub Form_Load()
On Error GoTo errorhandle
'Hide the main form to focus on the PortList Form
    frmMain.Hide
'Opens ports.txt and reads the file line by line and adds each line to the listbox
'untill it reaches the end of the file.
    Open App.Path & "\ports.txt" For Input As #1
    Do Until EOF(1)
        Line Input #1, lineoftext$
        lstPorts.AddItem lineoftext$
    Loop
    Close #1
    Exit Sub

'Error handling I wrote in just in case ports.txt is not found it wont crash.
errorhandle:
    lstPorts.AddItem ">-----<Port defenitions could not be found.>-----<"
    lstPorts.AddItem "Please make sure ports.txt is in the directory."
Exit Sub
End Sub
Private Sub Form_Unload(Cancel As Integer)
'When form is closed show the main form again.
    frmMain.Show
End Sub
