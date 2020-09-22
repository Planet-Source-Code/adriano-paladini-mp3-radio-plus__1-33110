VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSample 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sample Player"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Mp3RadioPlus.PlayButton cmdPLAY 
      Height          =   270
      Index           =   0
      Left            =   840
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   476
   End
   Begin VB.TextBox Vt 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3360
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Vt 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   9
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   4440
      Width           =   4935
   End
   Begin VB.TextBox Vt 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   8
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3960
      Width           =   4935
   End
   Begin VB.TextBox Vt 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   7
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3480
      Width           =   4935
   End
   Begin VB.TextBox Vt 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   6
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3000
      Width           =   4935
   End
   Begin VB.TextBox Vt 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   5
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2520
      Width           =   4935
   End
   Begin VB.TextBox Vt 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   4
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2040
      Width           =   4935
   End
   Begin VB.TextBox Vt 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   3
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   4935
   End
   Begin VB.TextBox Vt 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   316
      Index           =   2
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   4935
   End
   Begin VB.TextBox Vt 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   4935
   End
   Begin Mp3RadioPlus.PlayButton cmdPLAY 
      Height          =   270
      Index           =   1
      Left            =   840
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   476
   End
   Begin Mp3RadioPlus.PlayButton cmdPLAY 
      Height          =   270
      Index           =   2
      Left            =   840
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   476
   End
   Begin Mp3RadioPlus.PlayButton cmdPLAY 
      Height          =   270
      Index           =   3
      Left            =   840
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   476
   End
   Begin Mp3RadioPlus.PlayButton cmdPLAY 
      Height          =   270
      Index           =   4
      Left            =   840
      TabIndex        =   15
      Top             =   2040
      Visible         =   0   'False
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   476
   End
   Begin Mp3RadioPlus.PlayButton cmdPLAY 
      Height          =   270
      Index           =   5
      Left            =   840
      TabIndex        =   16
      Top             =   2520
      Visible         =   0   'False
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   476
   End
   Begin Mp3RadioPlus.PlayButton cmdPLAY 
      Height          =   270
      Index           =   6
      Left            =   840
      TabIndex        =   17
      Top             =   3000
      Visible         =   0   'False
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   476
   End
   Begin Mp3RadioPlus.PlayButton cmdPLAY 
      Height          =   270
      Index           =   7
      Left            =   840
      TabIndex        =   18
      Top             =   3480
      Visible         =   0   'False
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   476
   End
   Begin Mp3RadioPlus.PlayButton cmdPLAY 
      Height          =   270
      Index           =   8
      Left            =   840
      TabIndex        =   19
      Top             =   3960
      Visible         =   0   'False
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   476
   End
   Begin Mp3RadioPlus.PlayButton cmdPLAY 
      Height          =   270
      Index           =   9
      Left            =   840
      TabIndex        =   20
      Top             =   4440
      Visible         =   0   'False
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   476
   End
   Begin Mp3RadioPlus.PicScroll VBalance1 
      Height          =   225
      Left            =   360
      TabIndex        =   21
      Top             =   5310
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   397
      Min             =   -100
   End
   Begin VB.Image cmdCancel 
      Height          =   330
      Left            =   6000
      Top             =   5160
      Width           =   975
   End
   Begin VB.Image cmdOk 
      Height          =   330
      Left            =   4920
      Top             =   5160
      Width           =   975
   End
   Begin VB.Image cmdOpen 
      Height          =   330
      Left            =   1320
      Top             =   5160
      Width           =   975
   End
   Begin VB.Image cmdSave 
      Height          =   330
      Left            =   2400
      Top             =   5160
      Width           =   975
   End
   Begin VB.Image cmdCLEARALL 
      Height          =   330
      Left            =   3480
      Top             =   5160
      Width           =   975
   End
   Begin VB.Image cmdCLEAR 
      Height          =   270
      Index           =   9
      Left            =   6240
      Top             =   4440
      Width           =   315
   End
   Begin VB.Image cmdCLEAR 
      Height          =   270
      Index           =   8
      Left            =   6240
      Top             =   3960
      Width           =   315
   End
   Begin VB.Image cmdCLEAR 
      Height          =   270
      Index           =   7
      Left            =   6240
      Top             =   3480
      Width           =   315
   End
   Begin VB.Image cmdCLEAR 
      Height          =   270
      Index           =   6
      Left            =   6240
      Top             =   3000
      Width           =   315
   End
   Begin VB.Image cmdCLEAR 
      Height          =   270
      Index           =   5
      Left            =   6240
      Top             =   2520
      Width           =   315
   End
   Begin VB.Image cmdCLEAR 
      Height          =   270
      Index           =   4
      Left            =   6240
      Top             =   2040
      Width           =   315
   End
   Begin VB.Image cmdCLEAR 
      Height          =   270
      Index           =   3
      Left            =   6240
      Top             =   1560
      Width           =   315
   End
   Begin VB.Image cmdCLEAR 
      Height          =   270
      Index           =   2
      Left            =   6240
      Top             =   1080
      Width           =   315
   End
   Begin VB.Image cmdCLEAR 
      Height          =   270
      Index           =   1
      Left            =   6240
      Top             =   600
      Width           =   315
   End
   Begin VB.Image cmdCLEAR 
      Height          =   270
      Index           =   0
      Left            =   6240
      Top             =   120
      Width           =   315
   End
   Begin VB.Image cmdLoad 
      Height          =   270
      Index           =   9
      Left            =   480
      Top             =   4440
      Width           =   315
   End
   Begin VB.Image cmdLoad 
      Height          =   270
      Index           =   8
      Left            =   480
      Top             =   3960
      Width           =   315
   End
   Begin VB.Image cmdLoad 
      Height          =   270
      Index           =   7
      Left            =   480
      Top             =   3480
      Width           =   315
   End
   Begin VB.Image cmdLoad 
      Height          =   270
      Index           =   6
      Left            =   480
      Top             =   3000
      Width           =   315
   End
   Begin VB.Image cmdLoad 
      Height          =   270
      Index           =   5
      Left            =   480
      Top             =   2520
      Width           =   315
   End
   Begin VB.Image cmdLoad 
      Height          =   270
      Index           =   4
      Left            =   480
      Top             =   2040
      Width           =   315
   End
   Begin VB.Image cmdLoad 
      Height          =   270
      Index           =   3
      Left            =   480
      Top             =   1560
      Width           =   315
   End
   Begin VB.Image cmdLoad 
      Height          =   270
      Index           =   2
      Left            =   480
      Top             =   1080
      Width           =   315
   End
   Begin VB.Image cmdLoad 
      Height          =   270
      Index           =   1
      Left            =   480
      Top             =   600
      Width           =   315
   End
   Begin VB.Image cmdLoad 
      Height          =   270
      Index           =   0
      Left            =   480
      Top             =   120
      Width           =   315
   End
   Begin MediaPlayerCtl.MediaPlayer MP4 
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   495
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
      ClickToPlay     =   0   'False
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   0   'False
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
      ShowControls    =   0   'False
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
      Volume          =   -410
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "frmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AIndex As Integer
Dim Playing As Boolean
Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub cmdCLEAR_Click(Index As Integer)
Vt(Index).Text = ""
cmdPLAY(Index).Visible = False
End Sub
Private Sub cmdCLEARALL_Click()
For I = 0 To 9
    Vt(I).Text = ""
    cmdPLAY(I).Visible = False
Next I
End Sub
Private Sub cmdLOAD_Click(Index As Integer)
CD1.Filter = "MP3; WAV|*.WAV;*.MP3"
CD1.ShowOpen
If CD1.FileName <> "" Then
    Vt(Index).Text = CD1.FileName
    cmdPLAY(Index).Visible = True
End If
End Sub
Private Sub cmdOk_Click()
For I = 0 To 9
    SaveSetting "ACP Software", "Radio Plus", "Sample" & I, Vt(I).Text
    If Vt(I).Text = "" Then
        frmMain.cmdSAMPLE(I).ForeColor = Hex2VB(ReadINI(Arq0, "Colors", "SampleD"))
    Else
        frmMain.MPV(I).FileName = Vt(I).Text
        frmMain.cmdSAMPLE(I).ForeColor = Hex2VB(ReadINI(Arq0, "Colors", "SampleE"))
    End If
    frmMain.MPV(I).Balance = VBalance1.Value * 100
Next I
Unload Me
End Sub
Private Sub cmdOpen_Click()
CD1.Filter = "SMP|*.SMP"
CD1.ShowOpen
If CD1.FileName <> "" Then
    For I = 0 To 9
        Vt(I).Text = Trim(ReadINI(CD1.FileName, "SMP", "" & I))
        If Vt(I).Text <> "" Then cmdPLAY(Index).Visible = True
    Next I
End If
End Sub
Private Sub cmdPLAY_Click(Index As Integer)
If Playing = True Then
    If Index = AIndex Then
        MP4.Stop
        Playing = False
        cmdPLAY(AIndex).PStop = False
    Else
        MP4.FileName = Vt(Index).Text
        MP4.Play
        cmdPLAY(AIndex).PStop = False
        AIndex = Index
        cmdPLAY(Index).PStop = True
    End If
Else
    Playing = True
    MP4.FileName = Vt(Index).Text
    MP4.Play
    AIndex = Index
    cmdPLAY(AIndex).PStop = True
End If
End Sub
Private Sub cmdSave_Click()
CD1.Filter = "SMP|*.SMP"
CD1.ShowSave
If CD1.FileName <> "" Then
    For I = 0 To 9
        WriteINI CD1.FileName, "SMP", "" & I, Trim(Vt(I).Text) & ""
    Next I
End If
End Sub
Private Sub Form_Load()
Dim Arq As String
Dim Arq0 As String
Dim Arq1 As String
Arq = ReadINI(App.Path & "\RdPlus.ini", "Skin", "Skin")
Arq0 = App.Path & "\Skins\" & Arq & ".ini"
Arq1 = App.Path & "\Skins\" & ReadINI(Arq0, "Images", "Sample")
Color0 = Hex2VB(ReadINI(Arq0, "Colors", "ListBk"))
Color1 = Hex2VB(ReadINI(Arq0, "Colors", "ListTx"))
Me.Picture = LoadPicture(Arq1)

VBalance1.PaintPicture1 Me.Picture, 0, 0, , , 1215, 4455, 105, 225
VBalance1.PaintPicture2 Me.Picture, 0, 0, , , 360, 5310, 660, 225
VBalance1.PaintPicture3 Me.Picture, 0, 0, , , 360, 5310, 660, 225

For I = 0 To 9
    Vt(I).BackColor = Color0
    Vt(I).ForeColor = Color1
    cmdPLAY(I).PaintPicturePlay Me.Picture, 0, 0, , , 1215, 135, 315, 270
    cmdPLAY(I).PaintPictureStop Me.Picture, 0, 0, , , 1545, 135, 315, 270
    StrSample = GetSetting("ACP Software", "Radio Plus", "Sample" & I, "")
    If StrSample <> "" Then
        Vt(I).Text = StrSample
        cmdPLAY(I).Visible = True
    Else
        cmdPLAY(I).Visible = False
    End If
Next I
VBalance1.Value = frmMain.MPV(0).Balance / 100
End Sub
Private Sub MP4_EndOfStream(ByVal Result As Long)
cmdPLAY(AIndex).PStop = False
Playing = False
End Sub
Private Sub VBalance1_LeftClick()
VBalance1.Value = 0
End Sub
