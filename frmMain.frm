VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SoF2 Voice Manager"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   10515
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   10515
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command28 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Right"
      Height          =   495
      Left            =   1680
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command27 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Down"
      Height          =   495
      Left            =   960
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command26 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Up"
      Height          =   495
      Left            =   960
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Left"
      Height          =   495
      Left            =   240
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   4575
      Left            =   5880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   27
      Top             =   480
      Width           =   4455
   End
   Begin VB.CommandButton Command24 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter"
      Height          =   1095
      Left            =   4800
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H00FFFFFF&
      Caption         =   "+"
      Height          =   1095
      Left            =   4800
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H00FFFFFF&
      Caption         =   "--"
      Height          =   495
      Left            =   4800
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H00FFFFFF&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3"
      Height          =   495
      Left            =   4080
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00FFFFFF&
      Caption         =   "6"
      Height          =   495
      Left            =   4080
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00FFFFFF&
      Caption         =   "9"
      Height          =   495
      Left            =   4080
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
      Height          =   495
      Left            =   4080
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   495
      Left            =   3360
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "8"
      Height          =   495
      Left            =   3360
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "5"
      Height          =   495
      Left            =   3360
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2"
      Height          =   495
      Left            =   3360
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   495
      Left            =   2640
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1"
      Height          =   495
      Left            =   2640
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "4"
      Height          =   495
      Left            =   2640
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "7"
      Height          =   495
      Left            =   2640
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Num Lock"
      Height          =   495
      Left            =   2640
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Page Down"
      Height          =   495
      Left            =   1680
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "End"
      Height          =   495
      Left            =   960
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete"
      Height          =   495
      Left            =   240
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Page Up"
      Height          =   495
      Left            =   1680
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Home"
      Height          =   495
      Left            =   960
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Insert"
      Height          =   495
      Left            =   240
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmMain.frx":6852
      Left            =   240
      List            =   "frmMain.frx":6916
      TabIndex        =   0
      Top             =   480
      Width           =   5175
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   135
      Left            =   4440
      TabIndex        =   30
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "cfg Script"
      Height          =   255
      Left            =   5880
      TabIndex        =   28
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   1200
      Width           =   5175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Phrase:"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose a phrase from the list below"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSave 
         Caption         =   "&Apply All"
      End
      Begin VB.Menu mnuResetall 
         Caption         =   "Reset All"
      End
      Begin VB.Menu mnuViewScript 
         Caption         =   "View CFG Script"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuViewReadme 
         Caption         =   "View Readme"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Phrase As String
Dim AutoExec_Contents
Dim Voices_Contents
Dim UnRec As Boolean
Dim button As String
Dim ClearItem As String

Private Sub cmdApply_Click()
Open Sof2dir + "\base\mp\voices.cfg" For Output As #1
    Print #1, Text1.Text
Close #1
MsgBox "Voices applied!  Enjoy!", vbInformation, "Done!"
End Sub

Private Sub Command1_Click()
button = " LEFTARROW"
If Combo1.Text = "" Then
    ClearItem = "LEFTARROW"
    Command1.Tag = Combo1.Text
    Command1.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind LEFTARROW """ + Phrase + """" + vbCrLf
Command1.Tag = Combo1.Text
Command1.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command1_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command1.Tag
End Sub

Private Sub Command10_Click()
button = "KP_LEFTARROW"
If Combo1.Text = "" Then
    ClearItem = "KP_LEFTARROW"
    Command10.Tag = Combo1.Text
    Command10.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind KP_LEFTARROW """ + Phrase + """" + vbCrLf
Command10.Tag = Combo1.Text
Command10.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command10_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command10.Tag
End Sub

Private Sub Command11_Click()
button = "KP_END"
If Combo1.Text = "" Then
    ClearItem = "KP_END"
    Command11.Tag = Combo1.Text
    Command11.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind KP_END """ + Phrase + """" + vbCrLf
Command11.Tag = Combo1.Text
Command11.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command11_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command11.Tag
End Sub

Private Sub Command12_Click()
button = "KP_INS"
If Combo1.Text = "" Then
    ClearItem = "KP_INS"
    Command12.Tag = Combo1.Text
    Command12.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind KP_INS """ + Phrase + """" + vbCrLf
Command12.Tag = Combo1.Text
Command12.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command12_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command12.Tag
End Sub

Private Sub Command13_Click()
button = "KP_DOWNARROW"
If Combo1.Text = "" Then
    ClearItem = "KP_DOWNARROW"
    Command13.Tag = Combo1.Text
    Command13.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind KP_DOWNARROW """ + Phrase + """" + vbCrLf
Command13.Tag = Combo1.Text
Command13.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command13_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command13.Tag
End Sub

Private Sub Command14_Click()
button = "KP_5"
If Combo1.Text = "" Then
    ClearItem = "KP_5"
    Command14.Tag = Combo1.Text
    Command14.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind KP_5 """ + Phrase + """" + vbCrLf
Command14.Tag = Combo1.Text
Command14.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command14_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command14.Tag
End Sub

Private Sub Command15_Click()
button = "KP_UPARROW"
If Combo1.Text = "" Then
    ClearItem = "KP_UPARROW"
    Command15.Tag = Combo1.Text
    Command15.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind KP_UPARROW """ + Phrase + """" + vbCrLf
Command15.Tag = Combo1.Text
Command15.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command15_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command15.Tag
End Sub

Private Sub Command16_Click()
button = "KP_SLASH"
If Combo1.Text = "" Then
    ClearItem = "KP_SLASH"
    Command16.Tag = Combo1.Text
    Command16.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind KP_SLASH """ + Phrase + """" + vbCrLf
Command16.Tag = Combo1.Text
Command16.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command16_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command16.Tag
End Sub

Private Sub Command17_Click()
button = "*"
If Combo1.Text = "" Then
    ClearItem = "*"
    Command17.Tag = Combo1.Text
    Command17.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind * """ + Phrase + """" + vbCrLf
Command17.Tag = Combo1.Text
Command17.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command17_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command17.Tag
End Sub

Private Sub Command18_Click()
button = "KP_PGUP"
If Combo1.Text = "" Then
    ClearItem = "KP_PGUP"
    Command18.Tag = Combo1.Text
    Command18.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind KP_PGUP """ + Phrase + """" + vbCrLf
Command18.Tag = Combo1.Text
Command18.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command18_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command18.Tag
End Sub

Private Sub Command19_Click()
button = "KP_RIGHTARROW"
If Combo1.Text = "" Then
    ClearItem = "KP_RIGHTARROW"
    Command19.Tag = Combo1.Text
    Command19.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind KP_RIGHTARROW """ + Phrase + """" + vbCrLf
Command19.Tag = Combo1.Text
Command19.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command19_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command19.Tag
End Sub

Private Sub Command2_Click()
button = " INS"
If Combo1.Text = "" Then
    ClearItem = "INS"
    Command2.Tag = Combo1.Text
    Command2.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind INS """ + Phrase + """" + vbCrLf
Command2.Tag = Combo1.Text
Command2.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command2_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command2.Tag
End Sub

Private Sub Command20_Click()
button = "KP_PGDN"
If Combo1.Text = "" Then
    ClearItem = "KP_PGDN"
    Command20.Tag = Combo1.Text
    Command20.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind KP_PGDN """ + Phrase + """" + vbCrLf
Command20.Tag = Combo1.Text
Command20.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command20_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command20.Tag
End Sub

Private Sub Command21_Click()
button = "KP_DEL"
If Combo1.Text = "" Then
    ClearItem = "KP_DEL"
    Command21.Tag = Combo1.Text
    Command21.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind KP_DEL """ + Phrase + """" + vbCrLf
Command21.Tag = Combo1.Text
Command21.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command21_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command21.Tag
End Sub

Private Sub Command22_Click()
button = "KP_MINUS"
If Combo1.Text = "" Then
    ClearItem = "KP_MINUS"
    Command22.Tag = Combo1.Text
    Command22.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind KP_MINUS """ + Phrase + """" + vbCrLf
Command22.Tag = Combo1.Text
Command22.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command22_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command22.Tag
End Sub

Private Sub Command23_Click()
button = "KP_PLUS"
If Combo1.Text = "" Then
    ClearItem = "KP_PLUS"
    Command23.Tag = Combo1.Text
    Command23.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind KP_PLUS """ + Phrase + """" + vbCrLf
Command23.Tag = Combo1.Text
Command23.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command23_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command23.Tag
End Sub

Private Sub Command24_Click()
button = "KP_ENTER"
If Combo1.Text = "" Then
    ClearItem = "KP_ENTER"
    Command24.Tag = Combo1.Text
    Command24.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind KP_ENTER """ + Phrase + """" + vbCrLf
Command24.Tag = Combo1.Text
Command24.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command24_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command24.Tag
End Sub

Private Sub Command26_Click()
button = " UPARROW"
If Combo1.Text = "" Then
    ClearItem = "UPARROW"
    Command26.Tag = Combo1.Text
    Command26.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind UPARROW """ + Phrase + """" + vbCrLf
Command26.Tag = Combo1.Text
Command26.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command26_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command26.Tag
End Sub

Private Sub Command27_Click()
button = " DOWNARROW"
If Combo1.Text = "" Then
    ClearItem = "DOWNARROW"
    Command27.Tag = Combo1.Text
    Command27.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind DOWNARROW """ + Phrase + """" + vbCrLf
Command27.Tag = Combo1.Text
Command27.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command27_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command27.Tag
End Sub

Private Sub Command28_Click()
button = " RIGHTARROW"
If Combo1.Text = "" Then
    ClearItem = "RIGHTARROW"
    Command28.Tag = Combo1.Text
    Command28.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind RIGHTARROW """ + Phrase + """" + vbCrLf
Command28.Tag = Combo1.Text
Command28.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command28_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command28.Tag
End Sub

Private Sub Command3_Click()
button = " HOME"
If Combo1.Text = "" Then
    ClearItem = "HOME"
    Command3.Tag = Combo1.Text
    Command3.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind HOME """ + Phrase + """" + vbCrLf
Command3.Tag = Combo1.Text
Command3.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command3_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command3.Tag
End Sub

Private Sub Command4_Click()
button = " PGUP"
If Combo1.Text = "" Then
    ClearItem = "PGUP"
    Command4.Tag = Combo1.Text
    Command4.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind PGUP """ + Phrase + """" + vbCrLf
Command4.Tag = Combo1.Text
Command4.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command4_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command4.Tag
End Sub

Private Sub Command5_Click()
button = " DEL"
If Combo1.Text = "" Then
    ClearItem = "DEL"
    Command5.Tag = Combo1.Text
    Command5.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind DEL """ + Phrase + """" + vbCrLf
Command5.Tag = Combo1.Text
Command5.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command5_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command5.Tag
End Sub

Private Sub Command6_Click()
button = " END"
If Combo1.Text = "" Then
    ClearItem = "END"
    Command6.Tag = Combo1.Text
    Command6.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind END """ + Phrase + """" + vbCrLf
Command6.Tag = Combo1.Text
Command6.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command6_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command6.Tag
End Sub

Private Sub Command7_Click()
button = " PGDN"
If Combo1.Text = "" Then
    ClearItem = "PGDN"
    Command7.Tag = Combo1.Text
    Command7.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind PGDN """ + Phrase + """" + vbCrLf
Command7.Tag = Combo1.Text
Command7.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command7_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command7.Tag
End Sub

Private Sub Command8_Click()
button = "KP_NUMLOCK"
If Combo1.Text = "" Then
    ClearItem = "KP_NUMLOCK"
    Command8.Tag = Combo1.Text
    Command8.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind KP_NUMLOCK """ + Phrase + """" + vbCrLf
Command8.Tag = Combo1.Text
Command8.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command8_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command8.Tag
End Sub

Private Sub Command9_Click()
button = "KP_HOME"
If Combo1.Text = "" Then
    ClearItem = "KP_HOME"
    Command9.Tag = Combo1.Text
    Command9.BackColor = vbWhite
    Checker (button)
    Exit Sub
End If
GetPhrase
If UnRec = True Then
    UnRec = False
    MsgBox "Phrase not recognized.  Please select a valid phrase from the list.", vbExclamation + vbSystemModal, "Oops!"
    Exit Sub
End If
Text1.Text = Text1.Text + "bind KP_HOME """ + Phrase + """" + vbCrLf
Command9.Tag = Combo1.Text
Command9.BackColor = vbGreen
Checker (button)
End Sub

Private Sub Command9_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Command9.Tag
End Sub

Private Sub Form_Load()
Text1.Text = vbNullString
ALL = vbNullString
Me.Width = 5700
Me.Height = 6255
MediaPlayer1.FileName = App.Path + "\mrt.wav"
If FileExists(App.Path + "\sof2loc.dat") = True Then
    Open App.Path + "\sof2loc.dat" For Input As #1
        Input #1, FileLoc
    Close #1
    If FileExists(FileLoc + "\SoF2MP.exe") = True Then
        Sof2dir = FileLoc
        GoTo WereOK
    End If
Else
    If FileExists("C:\Program Files\Soldier of Fortune II - Double Helix\SoF2MP.exe") = True Then
        MsgBox "C:\Program Files\Soldier of Fortune II - Double Helix\SoF2MP.exe was located successfully!", vbInformation, "SoF2MP.exe found"
        fPath = "C:\Program Files\Soldier of Fortune II - Double Helix\"
        Remloc
        Sof2dir = fPath
    Else
        MsgBox "SoF2MP.exe was not found.  Please locate it manually.", vbExclamation, "File not found"
        frmLocate.Show
        Unload Me
        Exit Sub
    End If
End If
WereOK:
'check if autoexec.cfg exists:
If FileExists(Sof2dir + "\base\mp\autoexec.cfg") = True Then
'if it exists, check if "exec voices.cfg has been added already (i.e. the prog has been used before)
    Open Sof2dir + "\base\mp\autoexec.cfg" For Input As #1
        Do Until EOF(1)
            Input #1, AutoExec_Contents
        Loop
    Close #1
    If InStr(AutoExec_Contents, "exec voices.cfg") > 0 Then 'yes, everythings set.. clear up
        'ClearUp
        LoadPresets
    Else
        'ok, we gotta call "exec voices.cfg" cuz it's not there yet, and also create voices.cfg
        Open Sof2dir + "\base\mp\autoexec.cfg" For Append As #1
            Print #1, vbCrLf + "exec voices.cfg"
        Close #1
        'create voices.cfg
        Open Sof2dir + "\base\mp\voices.cfg" For Output As #1
            Print #1, ""
        Close #1
    End If
Else
'if it doesn't exist, create it and call "exec voices.cfg" and also create voices.cfg
    Open Sof2dir + "\base\mp\autoexec.cfg" For Output As #1
        Print #1, "exec voices.cfg"
    Close #1
    'create voices.cfg
    Open Sof2dir + "\base\mp\voices.cfg" For Output As #1
        Print #1, ""
    Close #1
End If
If Not FileExists(Sof2dir + "\base\mp_SoF2_VM.pk3") Then
    FileCopy App.Path + "\mp_SoF2_VM.pk3", Sof2dir + "\base\mp_SoF2_VM.pk3"
End If
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = ""
End Sub

Private Sub mnuAbout_Click()
On Error Resume Next
frmAbout.Show
MediaPlayer1.Play
End Sub

Private Sub mnuExit_Click()
Unload Me
Unload frmLocate
Unload frmAbout
End
End Sub

Private Sub mnuHelp_Click()
MsgBox "-Select a phrase from the list." + vbCrLf + vbCrLf + "-Then click the button corresponding to the keyboard key you'd like to assign that phrase to." + vbCrLf + vbCrLf + "-You can clear a key by deleting all text in the ""Choose a phrase"" box and assiging a blank phrase to a button." + vbCrLf + vbCrLf + "-You can see what phrase is currently set to a button by moving the mouse over the button." + vbCrLf + vbCrLf + "-When you're all done, click the ""Apply"" button." + vbCrLf + vbCrLf + "-Now when you play SoF2, you can hear the phrases you selected by pressing the assigned keys.  Enjoy!", vbInformation + vbSystemModal, "How to:"
End Sub

Private Sub mnuResetall_Click()
If Text1.Text = "" Then Exit Sub
If MsgBox("Are you sure you want to clear all keybinds?", vbQuestion + vbYesNo, "Really?") = vbYes Then
Text1.Text = ""
Command1.Tag = "": Command2.Tag = "": Command3.Tag = "": Command4.Tag = ""
Command5.Tag = "": Command6.Tag = "": Command7.Tag = "": Command8.Tag = ""
Command9.Tag = "": Command10.Tag = "": Command11.Tag = "": Command12.Tag = ""
Command13.Tag = "": Command14.Tag = "": Command15.Tag = "": Command16.Tag = ""
Command17.Tag = "": Command18.Tag = "": Command19.Tag = "": Command20.Tag = ""
Command21.Tag = "": Command22.Tag = "": Command23.Tag = "": Command24.Tag = ""
Command26.Tag = "": Command27.Tag = "": Command28.Tag = ""
'
Command1.BackColor = vbWhite: Command2.BackColor = vbWhite
Command3.BackColor = vbWhite: Command4.BackColor = vbWhite
Command5.BackColor = vbWhite: Command6.BackColor = vbWhite
Command7.BackColor = vbWhite: Command8.BackColor = vbWhite
Command9.BackColor = vbWhite: Command10.BackColor = vbWhite
Command11.BackColor = vbWhite: Command12.BackColor = vbWhite
Command13.BackColor = vbWhite: Command14.BackColor = vbWhite
Command15.BackColor = vbWhite: Command16.BackColor = vbWhite
Command17.BackColor = vbWhite: Command18.BackColor = vbWhite
Command19.BackColor = vbWhite: Command20.BackColor = vbWhite
Command21.BackColor = vbWhite: Command22.BackColor = vbWhite
Command23.BackColor = vbWhite: Command24.BackColor = vbWhite
Command26.BackColor = vbWhite: Command27.BackColor = vbWhite
Command28.BackColor = vbWhite
Beep
End If
End Sub

Function GetPhrase()
Select Case Combo1.ListIndex
    Case 0: Phrase = "vsay_team clear"
    Case 1: Phrase = "vsay_team east"
    Case 2: Phrase = "vsay_team holding"
    Case 3: Phrase = "vsay_team hot"
    Case 4: Phrase = "vsay_team north"
    Case 5: Phrase = "vsay_team notargets"
    Case 6: Phrase = "vsay_team secure"
    Case 7: Phrase = "vsay_team south"
    Case 8: Phrase = "vsay_team targets_cold"
    Case 9: Phrase = "vsay_team targets_hot"
    Case 10: Phrase = "vsay_team multiple_targets"
    Case 11: Phrase = "vsay_team one_target"
    Case 12: Phrase = "vsay_team west"
    Case 13: Phrase = "vsay_team come_with_me"
    Case 14: Phrase = "vsay_team advance"
    Case 15: Phrase = "vsay_team cover_me"
    Case 16: Phrase = "vsay_team got_bastard"
    Case 17: Phrase = "vsay_team hold_position"
    Case 18: Phrase = "vsay_team need_backup"
    Case 19: Phrase = "vsay_team over_here"
    Case 20: Phrase = "vsay_team report_targets"
    Case 21: Phrase = "vsay_team secure_the_area"
    Case 22: Phrase = "vsay_team slaughtered"
    Case 23: Phrase = "vsay_team surround_him"
    Case 24: Phrase = "vsay_team target_eliminated"
    Case 25: Phrase = "vsay_team man_down"
    Case 26: Phrase = "vsay_team they_went_this_way"
    Case 27: Phrase = "vsay_team wet_pants"
    Case 28: Phrase = "vsay_team action_start"
    Case 29: Phrase = "vsay_team scared"
    Case 30: Phrase = "vsay_team coming_through"
    Case 31: Phrase = "vsay_team those_men_have_guns"
    Case 32: Phrase = "vsay_team im_bleeding"
    Case 33: Phrase = "vsay_team cry_to_mama"
    Case 34: Phrase = "vsay_team get_some"
    Case 35: Phrase = "vsay_team got_him"
    Case 36: Phrase = "vsay_team cover_left"
    Case 37: Phrase = "vsay_team cover_right"
    Case 38: Phrase = "vsay_team eyes_open"
    Case 39: Phrase = "vsay_team get_down"
    Case 40: Phrase = "vsay_team get_moving"
    Case 41: Phrase = "vsay_team go_check"
    Case 42: Phrase = "vsay_team good_work"
    Case 43: Phrase = "vsay_team hold_position"
    Case 44: Phrase = "vsay_team incoming"
    Case 45: Phrase = "vsay_team move_out"
    Case 46: Phrase = "vsay_team nice_shot"
    Case 47: Phrase = "vsay_team take_out"
    Case 48: Phrase = "vsay_team take_that"
    Case 49: Phrase = "vsay_team underfire"
    Case 50: Phrase = "vsay_team lets_go"
    Case 51: Phrase = "vsay_team this_place"
    Case 52: Phrase = "vsay_team shuddup"
    Case 53: Phrase = "vsay_team close"
    Case 54: Phrase = "vsay_team eat_lead"
    Case 55: Phrase = "vsay_team kicking_ass"
    Case 56: Phrase = "vsay_team want_some"
    Case 57: Phrase = "vsay_team you_like"
    Case 58: Phrase = "vsay_team that_guy"
    Case 59: Phrase = "vsay_team negative"
    Case 60: Phrase = "vsay_team bastards_out"
    Case 61: Phrase = "vsay_team 123_go"
    Case 62: Phrase = "vsay_team take_cover"
    Case 63: Phrase = "vsay_team affirmative"
    Case Else: UnRec = True
End Select
End Function

Sub Checker(CMD As String)
    Dim ALL
    Dim ALL2
    Dim Temp As String
    Dim i As Integer
    i = 0
    ALL = Text1.Text
    numbrks = CharCount(ALL, vbCrLf, 2)
    'If numbrks = 1 Then Exit Sub 'no need to do it.....
    Dim LItem()
    
    ReDim LItem(numbrks)
    ALL2 = ALL 'for later usage...
    Text1.Text = "" 'it's stored in ALL, so it's ALL good.
        'idea:
        'Store ALL in an array LItem(), then cross-reference ALL2 with the array
        'and see if there's been duplicates.  If so, store Nullstring in the
        'infected Array item(s), then load The array back into the text box
        'all reset all vars for next time.  Should work.
        'Ok, so let's store ALL into array LIten() by means of finding VbCrLf:
        'Find vbCrLf
    
    For p = 1 To numbrks 'repeat once for each line break.....
        For j = 1 To Len(ALL)
            Temp = Mid$(ALL, j, 2)
            If Temp = vbCrLf Then 'ok, we found a Line Break!
                'now we gotta store the text before the linebreak into an array item
                'to make this easier, we need to delete this same stored text from ALL,
                'so next time we can start at the beginning again.  dont worry, we still
                'have the original text stored in ALL2 for future use!
                i = i + 1
                LItem(i) = Mid(ALL, 1, j + 1) 'store it in LItem(i)
                'If CharCount(ALL, LItem(i), Len(LItem(i))) > 1 Then
                    'simple duplicate...
                    'replace first occurance...
                    Temp2$ = Replace(ALL, LItem(i), "", 1, 1)
                    ALL = Temp2$
                    If Len(ALL) > j Then Exit For
                'End If
                'ok, now repeat!
            End If
        Next j
    Next p
            '
            'ok, ALL is stored line by line into the array LItem().  now
            'it's time to cross-check it with ALL2 to see for duplicates!
            '
        
blah:
        For j = 1 To UBound(LItem())
            If CharCount(ALL2, LItem(j), Len(LItem(j))) > 1 Then
                'well, we found a duplicate.  so store nullstring to this array item!
                LItem(j) = vbNullString
                Exit For
                'ok this array item is gone.
            End If
            If CharCount(ALL2, button, Len(button)) > 1 Then
                For i = 1 To UBound(LItem())
                    If CharCount(LItem(i), button, Len(button)) Then
                        LItem(i) = vbNullString
                        GoTo esc:
                    End If
                Next i
            End If
        Next j
esc:


Select Case ClearItem
                Case "": GoTo bleh:
                Case "LEFTARROW"
                    For i = 1 To UBound(LItem())
                        If CharCount(LItem(i), " LEFTARROW", Len(" LEFTARROW")) > 0 Then LItem(i) = vbNullString: Command1.BackColor = vbWhite
                    Next i
                Case "UPARROW"
                    For i = 1 To UBound(LItem())
                        If CharCount(LItem(i), " UPARROW", Len(" UPARROW")) > 0 Then LItem(i) = vbNullString: Command26.BackColor = vbWhite
                    Next i
                Case "DOWNARROW"
                    For i = 1 To UBound(LItem())
                        If CharCount(LItem(i), " DOWNARROW", Len(" DOWNARROW")) > 0 Then LItem(i) = vbNullString: Command27.BackColor = vbWhite
                    Next i
                Case "RIGHTARROW"
                    For i = 1 To UBound(LItem())
                        If CharCount(LItem(i), " RIGHTARROW", Len(" RIGHTARROW")) > 0 Then LItem(i) = vbNullString: Command28.BackColor = vbWhite
                    Next i
                Case "INS"
                    For i = 1 To UBound(LItem())
                        If CharCount(LItem(i), " INS", Len(" INS")) > 0 Then LItem(i) = vbNullString: Command2.BackColor = vbWhite
                    Next i
                Case "HOME"
                    For i = 1 To UBound(LItem())
                        If CharCount(LItem(i), " HOME", Len(" HOME")) > 0 Then LItem(i) = vbNullString: Command3.BackColor = vbWhite
                    Next i
                Case "PGUP"
                    For i = 1 To UBound(LItem())
                        If CharCount(LItem(i), " PGUP", Len(" PGUP")) > 0 Then LItem(i) = vbNullString: Command4.BackColor = vbWhite
                    Next i
                Case "DEL"
                    For i = 1 To UBound(LItem())
                        If CharCount(LItem(i), " DEL", Len(" DEL")) > 0 Then LItem(i) = vbNullString: Command5.BackColor = vbWhite
                    Next i
                Case "END"
                    For i = 1 To UBound(LItem())
                        If CharCount(LItem(i), " END", Len(" END")) > 0 Then LItem(i) = vbNullString: Command6.BackColor = vbWhite
                    Next i
                Case "PGDN"
                    For i = 1 To UBound(LItem())
                        If CharCount(LItem(i), " PGDN", Len(" PGDN")) > 0 Then LItem(i) = vbNullString: Command7.BackColor = vbWhite
                    Next i
                Case "KP_NUMLOCK"
                    For i = 1 To UBound(LItem())
                        If InStr(1, LItem(i), "KP_NUMLOCK") > 0 Then LItem(i) = vbNullString: Command8.BackColor = vbWhite
                    Next i
                Case "KP_SLASH"
                    For i = 1 To UBound(LItem())
                        If InStr(1, LItem(i), "KP_SLASH") > 0 Then LItem(i) = vbNullString: Command16.BackColor = vbWhite
                    Next i
                Case "KP_MINUS"
                    For i = 1 To UBound(LItem())
                        If InStr(1, LItem(i), "KP_MINUS") > 0 Then LItem(i) = vbNullString: Command22.BackColor = vbWhite
                    Next i
                Case "KP_HOME"
                    For i = 1 To UBound(LItem())
                        If InStr(1, LItem(i), "KP_HOME") > 0 Then LItem(i) = vbNullString: Command9.BackColor = vbWhite
                    Next i
                Case "KP_UPARROW"
                    For i = 1 To UBound(LItem())
                        If InStr(1, LItem(i), "KP_UPARROW") > 0 Then LItem(i) = vbNullString: Command15.BackColor = vbWhite
                    Next i
                Case "KP_PGUP"
                    For i = 1 To UBound(LItem())
                        If InStr(1, LItem(i), "KP_PGUP") > 0 Then LItem(i) = vbNullString: Command18.BackColor = vbWhite
                    Next i
                Case "KP_PLUS"
                    For i = 1 To UBound(LItem())
                        If InStr(1, LItem(i), "KP_PLUS") > 0 Then LItem(i) = vbNullString: Command23.BackColor = vbWhite
                    Next i
                Case "KP_LEFTARROW"
                    For i = 1 To UBound(LItem())
                        If InStr(1, LItem(i), "KP_LEFTARROW") > 0 Then LItem(i) = vbNullString: Command10.BackColor = vbWhite
                    Next i
                Case "KP_5"
                    For i = 1 To UBound(LItem())
                        If InStr(1, LItem(i), "KP_5") > 0 Then LItem(i) = vbNullString: Command14.BackColor = vbWhite
                    Next i
                Case "KP_RIGHTARROW"
                    For i = 1 To UBound(LItem())
                        If InStr(1, LItem(i), "KP_RIGHTARROW") > 0 Then LItem(i) = vbNullString: Command19.BackColor = vbWhite
                    Next i
                Case "KP_END"
                    For i = 1 To UBound(LItem())
                        If InStr(1, LItem(i), "KP_END") > 0 Then LItem(i) = vbNullString: Command11.BackColor = vbWhite
                    Next i
                Case "KP_DOWNARROW"
                    For i = 1 To UBound(LItem())
                        If InStr(1, LItem(i), "KP_DOWNARROW") > 0 Then LItem(i) = vbNullString: Command13.BackColor = vbWhite
                    Next i
                Case "KP_PGDN"
                    For i = 1 To UBound(LItem())
                        If InStr(1, LItem(i), "KP_PGDN") > 0 Then LItem(i) = vbNullString: Command20.BackColor = vbWhite
                    Next i
                Case "*"
                    For i = 1 To UBound(LItem())
                        If InStr(1, LItem(i), "*") > 0 Then LItem(i) = vbNullString: Command17.BackColor = vbWhite
                    Next i
                Case "KP_INS"
                    For i = 1 To UBound(LItem())
                        If InStr(1, LItem(i), "KP_INS") > 0 Then LItem(i) = vbNullString: Command12.BackColor = vbWhite
                    Next i
                Case "KP_DEL"
                    For i = 1 To UBound(LItem())
                        If InStr(1, LItem(i), "KP_DEL") > 0 Then LItem(i) = vbNullString: Command21.BackColor = vbWhite
                    Next i
                Case "KP_ENTER"
                    For i = 1 To UBound(LItem())
                        If InStr(1, LItem(i), "KP_ENTER") > 0 Then LItem(i) = vbNullString: Command24.BackColor = vbWhite
                    Next i
            End Select


bleh:

ClearItem = ""


        '
        'now that we have our NEW nice and clean arrays, with duplicates scratched,
        'we can finally restore the array items to text1.text.
        '
        For j = 1 To UBound(LItem())
            Text1.Text = Text1.Text + LItem(j)
        Next j
        '
        'clear vars: (not needed)
        For i = 1 To UBound(LItem())
            LItem(i) = vbNullString
        Next i
        ALL = ""
        ALL2 = ""
        Temp = ""
        Temp2$ = ""
        numbrks = 0
        button = ""
End Sub

Function ClearUp()
    On Error GoTo err 'user deleted voices.cfg?
    'clears voices.cfg upon restarting prog
    'should be deleted if you ever get around to reloading the presets
    Kill Sof2dir + "\base\mp\voices.cfg"
    Open Sof2dir + "\base\mp\voices.cfg" For Output As #1
        Print #1, ""
    Close #1
    Text1.Text = ""
    Exit Function
err:
    If MsgBox(err.Description + vbCrLf + vbCrLf + "voices.cfg was not found.  This file should be found in the ""sof2/base/mp/"" folder along with autoexec.cfg.  Would you like to to create it?", vbCritical + vbSystemModal + vbYesNo, "Error:") = vbNo Then
        MsgBox "voices.cfg needs to exist for this program to function properly.", vbCritical + vbSystemModal, "Error:"
        End
    Else
        Open Sof2dir + "\base\mp\voices.cfg" For Output As #1
            Print #1, ""
        Close #1
    End If
End Function

Private Sub mnuSave_Click()
cmdApply_Click
End Sub

Private Sub mnuViewReadme_Click()
Shell "Notepad " + App.Path + "\readme.txt", vbNormalFocus
End Sub

Private Sub mnuViewScript_Click()
If mnuViewScript.Checked = False Then mnuViewScript.Checked = True Else mnuViewScript.Checked = False
If mnuViewScript.Checked = True Then
    Me.Width = 10725
Else
    Me.Width = 5700
End If
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2
End Sub

Function CharCount(ByVal Start As String, ByVal Search As String, ByVal NumCount As Integer) As Integer
    Dim CCount As Integer
    CCount = 0
    For i = 1 To Len(Start)
        Temp$ = Mid$(Start, i, NumCount)
        If Temp$ = Search Then CCount = CCount + 1
    Next i
    CharCount = CCount
End Function

Function LoadPresets()
'First we must load contents of voices.cfg into text1.text.
'Of course if it's empty, then no need to.
Dim X As String, Y As String
Open Sof2dir + "\base\mp\voices.cfg" For Input As #1
    Do Until EOF(1)
        Input #1, X
        Y = Y + X + vbCrLf
    Loop
    Text1.Text = Y
Close #1
If Text1.Text = "" Then Exit Function
' 1.Load voices.cfg into a string variable
' 2.Put it line-by-line into an array LItem()
'  Diagnose each array item by means of--
'--3.What key is it assigned from?
'----Once we figure that out, we further diagnose it to determine--
'--4. What taunt is set to that key?
'----One long algorithm will be used to determine this.
'
'1:
Dim VoicesCFG As String, Voices_TEMP As String
VoicesCFG = Text1.Text
Voices_TEMP = Text1.Text
'
'2:
Dim Temp As String, Temp2 As String
Dim Count As Integer
Dim LItem() As String
Dim NumBreaks As Integer
NumBreaks = CharCount(Voices_TEMP, vbCrLf, 2)
ReDim LItem(NumBreaks) As String
For i = 1 To NumBreaks
    For j = 1 To Len(Voices_TEMP)
        Temp = Mid$(Voices_TEMP, j, 2)
        If Temp = vbCrLf Then
            Count = Count + 1
            LItem(Count) = Mid(Voices_TEMP, 1, j + 1)
            Temp2 = Replace(Voices_TEMP, LItem(Count), vbNullString, 1, 1)
            Voices_TEMP = Temp2
            If Len(Voices_TEMP) > j Then Exit For
        End If
    Next j
Next i
Count = 0
'
'3:
For i = 1 To UBound(LItem())
    If InStr(1, LItem(i), " INS") > 0 Then Diagnose "INS", LItem(i)
    If InStr(1, LItem(i), " HOME") > 0 Then Diagnose "HOME", LItem(i)
    If InStr(1, LItem(i), " PGUP") > 0 Then Diagnose "PGUP", LItem(i)
    If InStr(1, LItem(i), " DEL") > 0 Then Diagnose "DEL", LItem(i)
    If InStr(1, LItem(i), " END") > 0 Then Diagnose "END", LItem(i)
    If InStr(1, LItem(i), " PGDN") > 0 Then Diagnose "PGDN", LItem(i)
    If InStr(1, LItem(i), " LEFTARROW") > 0 Then Diagnose "LEFTARROW", LItem(i)
    If InStr(1, LItem(i), " UPARROW") > 0 Then Diagnose "UPARROW", LItem(i)
    If InStr(1, LItem(i), " DOWNARROW") > 0 Then Diagnose "DOWNARROW", LItem(i)
    If InStr(1, LItem(i), " RIGHTARROW") > 0 Then Diagnose "RIGHTARROW", LItem(i)
    If InStr(1, LItem(i), "KP_NUMLOCK") > 0 Then Diagnose "KP_NUMLOCK", LItem(i)
    If InStr(1, LItem(i), "KP_SLASH") > 0 Then Diagnose "KP_SLASH", LItem(i)
    If InStr(1, LItem(i), "*") > 0 Then Diagnose "*", LItem(i)
    If InStr(1, LItem(i), "KP_MINUS") > 0 Then Diagnose "KP_MINUS", LItem(i)
    If InStr(1, LItem(i), "KP_HOME") > 0 Then Diagnose "KP_HOME", LItem(i)
    If InStr(1, LItem(i), "KP_UPARROW") > 0 Then Diagnose "KP_UPARROW", LItem(i)
    If InStr(1, LItem(i), "KP_PGUP") > 0 Then Diagnose "KP_PGUP", LItem(i)
    If InStr(1, LItem(i), "KP_PLUS") > 0 Then Diagnose "KP_PLUS", LItem(i)
    If InStr(1, LItem(i), "KP_LEFTARROW") > 0 Then Diagnose "KP_LEFTARROW", LItem(i)
    If InStr(1, LItem(i), "KP_5") > 0 Then Diagnose "KP_5", LItem(i)
    If InStr(1, LItem(i), "KP_RIGHTARROW") > 0 Then Diagnose "KP_RIGHTARROW", LItem(i)
    If InStr(1, LItem(i), "KP_END") > 0 Then Diagnose "KP_END", LItem(i)
    If InStr(1, LItem(i), "KP_DOWNARROW") > 0 Then Diagnose "KP_DOWNARROW", LItem(i)
    If InStr(1, LItem(i), "KP_PGDN") > 0 Then Diagnose "KP_PGDN", LItem(i)
    If InStr(1, LItem(i), "KP_ENTER") > 0 Then Diagnose "KP_ENTER", LItem(i)
    If InStr(1, LItem(i), "KP_INS") > 0 Then Diagnose "KP_INS", LItem(i)
    If InStr(1, LItem(i), "KP_DEL") > 0 Then Diagnose "KP_DEL", LItem(i)
Next i
End Function

Function Diagnose(Key As String, Whole As String)
Dim Phrase As String
marker# = Len("bind ") + Len(Key) + Len(" 'vsay_team ") + 1
Phrase = Mid$(Whole, marker, Len(Whole) - marker - 2)
If Key = "INS" Then Command2.Tag = TranslatePhrase(Phrase): Command2.BackColor = vbGreen
If Key = "HOME" Then Command3.Tag = TranslatePhrase(Phrase): Command3.BackColor = vbGreen
If Key = "PGUP" Then Command4.Tag = TranslatePhrase(Phrase): Command4.BackColor = vbGreen
If Key = "DEL" Then Command5.Tag = TranslatePhrase(Phrase): Command5.BackColor = vbGreen
If Key = "END" Then Command6.Tag = TranslatePhrase(Phrase): Command6.BackColor = vbGreen
If Key = "PGDN" Then Command7.Tag = TranslatePhrase(Phrase): Command7.BackColor = vbGreen
If Key = "UPARROW" Then Command26.Tag = TranslatePhrase(Phrase): Command26.BackColor = vbGreen
If Key = "LEFTARROW" Then Command1.Tag = TranslatePhrase(Phrase): Command1.BackColor = vbGreen
If Key = "DOWNARROW" Then Command27.Tag = TranslatePhrase(Phrase): Command27.BackColor = vbGreen
If Key = "RIGHTARROW" Then Command28.Tag = TranslatePhrase(Phrase): Command28.BackColor = vbGreen
If Key = "KP_NUMLOCK" Then Command8.Tag = TranslatePhrase(Phrase): Command8.BackColor = vbGreen
If Key = "KP_SLASH" Then Command16.Tag = TranslatePhrase(Phrase): Command16.BackColor = vbGreen
If Key = "*" Then Command17.Tag = TranslatePhrase(Phrase): Command17.BackColor = vbGreen
If Key = "KP_MINUS" Then Command22.Tag = TranslatePhrase(Phrase): Command22.BackColor = vbGreen
If Key = "KP_HOME" Then Command9.Tag = TranslatePhrase(Phrase): Command9.BackColor = vbGreen
If Key = "KP_UPARROW" Then Command15.Tag = TranslatePhrase(Phrase): Command15.BackColor = vbGreen
If Key = "KP_PGUP" Then Command18.Tag = TranslatePhrase(Phrase): Command18.BackColor = vbGreen
If Key = "KP_PLUS" Then Command23.Tag = TranslatePhrase(Phrase): Command23.BackColor = vbGreen
If Key = "KP_LEFTARROW" Then Command10.Tag = TranslatePhrase(Phrase): Command10.BackColor = vbGreen
If Key = "KP_5" Then Command14.Tag = TranslatePhrase(Phrase): Command14.BackColor = vbGreen
If Key = "KP_RIGHTARROW" Then Command19.Tag = TranslatePhrase(Phrase): Command19.BackColor = vbGreen
If Key = "KP_END" Then Command11.Tag = TranslatePhrase(Phrase): Command11.BackColor = vbGreen
If Key = "KP_DOWNARROW" Then Command13.Tag = TranslatePhrase(Phrase): Command13.BackColor = vbGreen
If Key = "KP_PGDN" Then Command20.Tag = TranslatePhrase(Phrase): Command20.BackColor = vbGreen
If Key = "KP_ENTER" Then Command24.Tag = TranslatePhrase(Phrase): Command24.BackColor = vbGreen
If Key = "KP_INS" Then Command12.Tag = TranslatePhrase(Phrase): Command12.BackColor = vbGreen
If Key = "KP_DEL" Then Command21.Tag = TranslatePhrase(Phrase): Command21.BackColor = vbGreen
End Function

Function TranslatePhrase(What As String) As String
Dim NewP As String
Select Case What
Case "clear": NewP = "Clear!"
Case "east": NewP = "To your East!"
Case "holding": NewP = "Holding..."
Case "hot": NewP = "hot!"
Case "north": NewP = "To your North!"
Case "notargets": NewP = "No targets to report."
Case "secure": NewP = "Secure!"
Case "south": NewP = "To your South!"
Case "targets_cold": NewP = "Targets are cold!"
Case "targets_hot": NewP = "Targets are hot!"
Case "multiple_targets": NewP = "Multiple targets."
Case "one_target": NewP = "One target."
Case "west": NewP = "To your West!"
Case "come_with_me": NewP = "You: come with me!"
Case "advance": NewP = "Advance!"
Case "cover_me": NewP = "Cover Me"
Case "got_bastard": NewP = "I got that bastard!"
Case "hold_position": NewP = "Hold position!"
Case "need_backup": NewP = "I need backup, now!"
Case "over_here": NewP = "He's over here!"
Case "report_targets": NewP = "Report any targets."
Case "secure_the_area": NewP = "Secure the area!"
Case "slaughtered": NewP = "We're getting slaughtered!"
Case "surround_him": NewP = "Surround him!"
Case "target_eliminated": NewP = "Target has been eliminated."
Case "man_down": NewP = "Man Down!"
Case "they_went_this_way": NewP = "Sir, I think they went this way."
Case "wet_pants": NewP = "Hey!  Did you see that guy?  He wet his pants! Hahah!"
Case "action_start": NewP = "I'm sick of all this guarding!  When's the action gonna start?"
Case "scared": NewP = "You scared?  You should be!"
Case "coming_through": NewP = "Coming through!"
Case "those_men_have_guns": NewP = "Those men have guns!"
Case "im_bleeding": NewP = "Oh God... I'm bleeding!"
Case "cry_to_mama": NewP = "Go cry to mama!"
Case "get_some": NewP = "Come get some!"
Case "got_him": NewP = "Got him!"
Case "cover_left": NewP = "Cover my left!"
Case "cover_right": NewP = "Cover my right!"
Case "eyes_open": NewP = "Keep your eyes open, boys."
Case "get_down": NewP = "Get down!"
Case "get_moving": NewP = "Let's get moving!"
Case "go_check": NewP = "Go check that out."
Case "good_work": NewP = "Good work, Marines!"
Case "hold_position": NewP = "Hold this position!"
Case "incoming": NewP = "Incoming!"
Case "move_out": NewP = "Let's move out."
Case "nice_shot": NewP = "Nice shot!"
Case "take_out": NewP = "Take him out!"
Case "take_that": NewP = "Take that!"
Case "underfire": NewP = "We're under fire!"
Case "lets_go": NewP = "Let's go!  Move move move!"
Case "this_place": NewP = "Man, this place is really getting on my nerves!"
Case "shuddup": NewP = "Shut up, man!"
Case "close": NewP = "Damn, that was close!"
Case "eat_lead": NewP = "Eat lead, sucka!"
Case "kicking_ass": NewP = "Kicking ass and taking names!"
Case "want_some": NewP = "You want some of this?"
Case "you_like": NewP = "How'd you like that?"
Case "that_guy": NewP = "Glad I'm not that guy."
Case "negative": NewP = "Negative!"
Case "bastards_out": NewP = "I'm gonna take these bastards out."
Case "123_go": NewP = "Alright.  1, 2, 3, GO!"
Case "take_cover": NewP = "Take Cover!"
Case "affirmative": NewP = "Affirmative!"
End Select
TranslatePhrase = NewP
End Function

