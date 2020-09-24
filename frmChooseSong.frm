VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmChooseSong 
   Caption         =   "Choose A Song To Play"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6510
   LinkTopic       =   "Form2"
   ScaleHeight     =   4140
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLyricLocalPath 
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   3840
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox txtSongName 
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   3360
      Visible         =   0   'False
      Width           =   3855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdChooseLyric 
      Caption         =   "Browse"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   2220
      Width           =   1455
   End
   Begin VB.TextBox txtLyricPath 
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2220
      Width           =   3735
   End
   Begin VB.TextBox txtSongPath 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   900
      Width           =   3735
   End
   Begin VB.CommandButton cmdChooseSong 
      Caption         =   "Browse"
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   900
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "3- Press Play"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "2- Choose Lyric of the Song (TEXT File)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "1- Choose Song (WAV Or MP3)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmChooseSong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SongName
Private Sub cmdChooseSong_Click()
On Error GoTo 10
CommonDialog1.Filter = "WaveFiles|*.wav;|MP3|*.mp3;|ALLFiles|*.*"
CommonDialog1.Action = 1
txtSongPath.Text = CommonDialog1.filename
Dim pos
pos = 0
For i = 1 To Len(txtSongPath.Text) - 1
If Mid(txtSongPath.Text, i, 1) = "\" Then
pos = i
End If
Next i
SongName = Mid(txtSongPath.Text, pos + 1)
SongName = Left(SongName, Len(SongName) - 4)
txtSongName.Text = SongName
Exit Sub
10:
MsgBox "You did not choose any Song"
Exit Sub
End Sub

Private Sub cmdChooseLyric_Click()
CommonDialog1.Filter = "TextFiles|*.txt;|AllFiles|*.*"
CommonDialog1.Action = 1
txtLyricPath.Text = CommonDialog1.filename
Dim Source, Destination
Source = CommonDialog1.filename
Clipboard.SetText Source
Destination = App.Path & "\Lyrics\" & SongName & ".txt"
Dim MyFile
MyFile = Dir(Destination)
If MyFile = "" Then
FileCopy Source, Destination
txtLyricLocalPath = Destination
Exit Sub
Else
GoTo 10
End If
10:
Dim Z
Z = MsgBox(Destination & " already Exist-Replace?", vbYesNo)
    If Z = vbYes Then
    'Kill Destination
    FileCopy Source, Destination
    txtLyricLocalPath = Destination
    Else
    Exit Sub
    End If

End Sub

Private Sub cmdPlay_Click()
Form1.Show
Me.Hide
End Sub

Private Sub Form_Load()
frmChooseSong.Icon = LoadPicture(App.Path & "\Dorami2.ico")
frmChooseSong.BackColor = RGB(242, 195, 255)

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
