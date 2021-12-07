VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Hourglass 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hourglass"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   173.353
   ScaleMode       =   0  'User
   ScaleTop        =   -50
   ScaleWidth      =   101.658
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Stop Alarm"
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   7920
      Width           =   975
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   6960
      Top             =   7920
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   7920
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set Alarm"
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   7920
      Width           =   975
   End
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   2160
      Top             =   4920
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3840
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7320
      Top             =   600
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   7440
      TabIndex        =   3
      Top             =   7920
      Visible         =   0   'False
      Width           =   615
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1085
      _cy             =   873
   End
   Begin VB.Image Image1 
      Height          =   5370
      Left            =   120
      Picture         =   "Form1.frx":0000
      Top             =   2280
      Width           =   7935
   End
   Begin VB.Shape dd 
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   3120
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Shape uu 
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   1815
      Left            =   3120
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   3720
      TabIndex        =   0
      Top             =   240
      Width           =   600
   End
   Begin VB.Shape sss 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   4080
      Shape           =   2  'Oval
      Top             =   5040
      Width           =   135
   End
End
Attribute VB_Name = "Hourglass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As Integer
Dim aaa As String
Dim bbb As String

Private Sub Command1_Click()
aaa = Text1.Text
Text1.Enabled = False
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
WindowsMediaPlayer1.Controls.stop
Text1.Enabled = True
Command1.Enabled = True
End Sub


Private Sub Timer1_Timer()
Label1.Caption = Format(Now(), "HH:MM:SS AMPM")

End Sub

Private Sub Timer2_Timer()
t = Format(Now(), "SS")

If t = 0 Then
uu.Top = 14.6 ' +1.19
uu.Height = 36.22 '-1.20


dd.Top = 88 '-1.19
dd.Height = 2.6 '+ 1.15
End If

If t = 2 Then
uu.Top = 15.79
uu.Height = 35.02


dd.Top = 86.19
dd.Height = 3.75
End If

If t = 4 Then
uu.Top = 16.98
uu.Height = 33.82


dd.Top = 85.65
dd.Height = 4.9
End If

If t = 6 Then
uu.Top = 18.17
uu.Height = 32.62


dd.Top = 84.43
dd.Height = 6.05
End If
If t = 8 Then
uu.Top = 19.36
uu.Height = 31.42


dd.Top = 83.24
dd.Height = 7.2
End If

If t = 10 Then
uu.Top = 20.55
uu.Height = 30.22


dd.Top = 82.05
dd.Height = 8.35
End If

If t = 12 Then
uu.Top = 21.74
uu.Height = 29.02


dd.Top = 80.86
dd.Height = 9.5
End If

If t = 14 Then
uu.Top = 22.93
uu.Height = 27.82


dd.Top = 79.67
dd.Height = 10.65
End If
If t = 16 Then
uu.Top = 24.12
uu.Height = 26.62


dd.Top = 78.48
dd.Height = 11.8
End If

If t = 18 Then
uu.Top = 25.31 '1.19
uu.Height = 25.42


dd.Top = 77.29
dd.Height = 12.95
End If
If t = 20 Then
uu.Top = 26.5
uu.Height = 24.22


dd.Top = 76.1
dd.Height = 14.1
End If
If t = 22 Then
uu.Top = 27.69
uu.Height = 23.02

dd.Top = 74.91
dd.Height = 15.25
End If
If t = 24 Then
uu.Top = 28.88
uu.Height = 21.82


dd.Top = 73.72
dd.Height = 16.4
End If
If t = 26 Then
uu.Top = 30.7
uu.Height = 20.62


dd.Top = 72.53
dd.Height = 17.55
End If
If t = 28 Then
uu.Top = 31.26
uu.Height = 19.42


dd.Top = 71.34
dd.Height = 18.7
End If

If t = 30 Then
uu.Top = 32.45
uu.Height = 18.22


dd.Top = 70.15
dd.Height = 19.85
End If

If t = 32 Then
uu.Top = 33.64
uu.Height = 17.02


dd.Top = 68.96
dd.Height = 21
End If

If t = 34 Then
uu.Top = 34.83
uu.Height = 15.82


dd.Top = 67.77
dd.Height = 22.15
End If

If t = 36 Then
uu.Top = 36.2
uu.Height = 14.62


dd.Top = 66.58
dd.Height = 23.3
End If

If t = 38 Then
uu.Top = 37.39
uu.Height = 13.42


dd.Top = 65.39
dd.Height = 24.45
End If

If t = 40 Then
uu.Top = 38.58
uu.Height = 12.22


dd.Top = 64.2
dd.Height = 25.6
End If

If t = 42 Then
uu.Top = 39.77
uu.Height = 11.02


dd.Top = 63.01
dd.Height = 26.75
End If

If t = 44 Then
uu.Top = 40.96
uu.Height = 9.82


dd.Top = 61.82
dd.Height = 27.9
End If

If t = 46 Then
uu.Top = 42.15
uu.Height = 8.62


dd.Top = 60.63
dd.Height = 29.05
End If

If t = 48 Then
uu.Top = 43.34
uu.Height = 7.42


dd.Top = 59.44
dd.Height = 30.2
End If

If t = 50 Then
uu.Top = 44.53
uu.Height = 6.22


dd.Top = 58.25
dd.Height = 31.35
End If

If t = 52 Then
uu.Top = 45.57
uu.Height = 5.22


dd.Top = 57.06
dd.Height = 32.5
End If

If t = 54 Then
uu.Top = 46.91
uu.Height = 3.82


dd.Top = 55.87
dd.Height = 33.65
End If

If t = 56 Then
uu.Top = 48.1
uu.Height = 2.62


dd.Top = 54.68
dd.Height = 34.8
End If

If t = 58 Then
uu.Top = 49.29
uu.Height = 1


dd.Top = 53.49
dd.Height = 35.95
End If

If t = 59 Then
uu.Top = 50.29
uu.Height = 0.5


dd.Top = 52.3
dd.Height = 36.62
End If

End Sub

Private Sub Timer3_Timer()
If sss.Top > 49 And sss.Top < 87 Then
sss.Top = sss.Top + 1
Else
sss.Top = 50
End If
End Sub

Private Sub Timer4_Timer()

bbb = Format(Now(), "HH:MM:SS AMPM")

If aaa = bbb Then
WindowsMediaPlayer1.URL = "1.mp3"
End If

End Sub
