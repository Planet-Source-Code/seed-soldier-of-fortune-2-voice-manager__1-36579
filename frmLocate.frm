VERSION 5.00
Begin VB.Form frmLocate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please Locate SoF2MP.exe"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7140
   Icon            =   "frmLocate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   7140
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   3720
      Width           =   1575
   End
   Begin VB.FileListBox File1 
      Height          =   3210
      Left            =   3840
      TabIndex        =   2
      Top             =   240
      Width           =   3015
   End
   Begin VB.DirListBox Dir1 
      Height          =   3465
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmLocate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If File1.FileName = "SoF2MP.exe" Then
  Me.Hide
  fPath = Dir1
  Remloc
  Sof2dir = fPath
  MsgBox "SoF2MP.exe was located successfully!", vbInformation, "SoF2MP.exe found"
  Load frmMain
  frmMain.Show
Else
  MsgBox "Please locate SoF2MP.exe.  This file is located install directory of SoF2.", vbExclamation, "Error:"
End If
End Sub

Private Sub Dir1_Change()
File1 = Dir1
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1 = Drive1
End Sub

Private Sub File1_DblClick()
Command1_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

