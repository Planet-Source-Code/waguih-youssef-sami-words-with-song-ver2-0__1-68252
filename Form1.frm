VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   4080
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   7.197
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   11.748
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   480
      Top             =   2760
   End
   Begin VB.TextBox txtVolume 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   4920
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtVolume 
      BackColor       =   &H00404040&
      Height          =   195
      Index           =   2
      Left            =   4920
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtVolume 
      BackColor       =   &H80000006&
      Height          =   225
      Index           =   1
      Left            =   4920
      TabIndex        =   3
      Top             =   810
      Width           =   975
   End
   Begin VB.TextBox txtVolume 
      BackColor       =   &H00404040&
      Height          =   195
      Index           =   0
      Left            =   4920
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      UpdateInterval  =   3000
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1920
      Width           =   6375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Click here to Change Volume ==>>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   4455
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   8.678
      X2              =   8.678
      Y1              =   0.212
      Y2              =   2.328
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      Height          =   1695
      Left            =   120
      Top             =   0
      Width           =   6375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuChoose 
         Caption         =   "Choose An Other Song"
      End
      Begin VB.Menu mnuExt 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function waveOutSetVolume Lib "Winmm" _
(ByVal wDeviceID As Integer, ByVal dwVolume As Long) As Integer
Dim i

Private Sub Form_Activate()
MMControl1.UpdateInterval = 3000
'****************Loading Song.wav*********************
'MMControl1.DeviceType = "WaveAudio"

Dim MySong
MySong = frmChooseSong.txtSongPath
If MySong = "" Then GoTo 20
MMControl1.filename = MySong
MMControl1.Command = "Open"
MMControl1.Command = "Play"

'*****************Loading Lyric File.txt**************
On Error GoTo 10
Dim MytxtFile
MytxtFile = frmChooseSong.txtLyricLocalPath.Text
Open MytxtFile For Input As #1
Exit Sub
'--------------------
10:
MsgBox "The Lyric text File Does not Exists in your Program Path", vbOKOnly
MMControl1.UpdateInterval = 0
frmChooseSong.Show
Me.Hide
Exit Sub
'--------------------
20:
MsgBox "You Should Choose A Song First From File Menue", vbOKOnly
frmChooseSong.Show
MMControl1.Command = "Close"
MMControl1.UpdateInterval = 0
frmChooseSong.Show
Me.Hide
End Sub

Private Sub Form_Load()
i = 0
Form1.BackColor = RGB(152, 208, 232)
Form1.Icon = LoadPicture(App.Path & "\Dorami2.ico")
End Sub

Private Sub Form_Unload(Cancel As Integer)
MMControl1.Command = "Close"
End
End Sub

Private Sub MMControl1_StatusUpdate()
 Dim MyString
 Line Input #1, MyString
 If MyString <> "" Then
 Text1.Text = MyString
 Else
Close #1
End If

End Sub

Private Sub mnuChoose_Click()
MMControl1.Command = "Close"
Close #1
Text1.Text = ""

frmChooseSong.Show
Me.Hide

End Sub

Private Sub mnuExt_Click()
End
End Sub

Private Sub txtVolume_Click(Index As Integer)
Dim vol As Long
    Select Case Index
    Case 0
        vol = -1673487296
    Case 1
        vol = 1310740000
    Case 2
        vol = 655370000
    Case 3
        vol = 327685000
    End Select
    waveOutSetVolume 0, vol

End Sub

Private Sub Timer1_Timer()
i = i + 1
Dim MyFormCaption, MyLength
MyFormCaption = " Click The File Menu to Choose another Song"
MyLength = Len(MyFormCaption)

Form1.Caption = Mid(MyFormCaption, 1, i)

If i = MyLength Then
i = 1
End If

End Sub

