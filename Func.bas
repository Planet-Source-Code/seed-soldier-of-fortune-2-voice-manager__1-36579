Attribute VB_Name = "Func"
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public fPath As String
Public FileLoc As String
Public Sof2dir As String

Public Function Remloc()
    Open App.Path + "\sof2loc.dat" For Output As #1
        Print #1, fPath
        Sof2dir = fPath
    Close #1
End Function

Public Function FileExists(sFilename As String)
    Dim Files As String
    Files = Dir(sFilename, vbHidden + vbSystem + vbNormal)
    If Files = "" Then
        FileExists = False
    Else
        FileExists = True
    End If
End Function

Public Function PlayWav(WavFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(WavFile$)
    If SafeFile$ <> "" Then
       Call sndPlaySound(WavFile$, SND_FLAG)
    End If
End Function

Public Function Pause(interval As Integer)
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Function
