VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2C1EC115-F1BA-11D3-BF43-00A0CC32BE58}#9.1#0"; "DMC2.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Radio Plus"
   ClientHeight    =   8655
   ClientLeft      =   5985
   ClientTop       =   330
   ClientWidth     =   11925
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   11925
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer4 
      Interval        =   50
      Left            =   4680
      Top             =   2160
   End
   Begin VB.PictureBox PicSpectrum1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   250
      ScaleHeight     =   450
      ScaleWidth      =   2340
      TabIndex        =   75
      Top             =   600
      Visible         =   0   'False
      Width           =   2345
   End
   Begin VB.PictureBox PicSpectrum2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   7090
      ScaleHeight     =   450
      ScaleWidth      =   2340
      TabIndex        =   76
      Top             =   600
      Visible         =   0   'False
      Width           =   2345
   End
   Begin VB.Timer Timer3 
      Interval        =   50
      Left            =   4680
      Top             =   1800
   End
   Begin DMC2.DMC DMC 
      Left            =   4680
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer TimerManualFade 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   4920
      Top             =   3600
   End
   Begin VB.Timer TimerCross 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   6600
      Top             =   3600
   End
   Begin VB.Timer Horas 
      Interval        =   300
      Left            =   5760
      Top             =   360
   End
   Begin Mp3RadioPlus.PicVScroll Volume1 
      Height          =   2295
      Left            =   7180
      TabIndex        =   0
      Top             =   5600
      Width           =   120
      _ExtentX        =   5583
      _ExtentY        =   13996
      Min             =   -360
      Max             =   0
      Value           =   -360
   End
   Begin Mp3RadioPlus.PicScroll Position1 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      Max             =   1
   End
   Begin Mp3RadioPlus.PicScroll Balance1 
      Height          =   225
      Left            =   7000
      TabIndex        =   5
      Top             =   8090
      Width           =   660
      _ExtentX        =   13785
      _ExtentY        =   4260
      Min             =   -100
   End
   Begin Mp3RadioPlus.PicVScroll Volume2 
      Height          =   2295
      Left            =   8200
      TabIndex        =   6
      Top             =   5600
      Width           =   120
      _ExtentX        =   5583
      _ExtentY        =   13996
      Min             =   -360
      Max             =   0
      Value           =   -360
   End
   Begin Mp3RadioPlus.PicScroll Position2 
      Height          =   375
      Left            =   7060
      TabIndex        =   7
      Top             =   3000
      Width           =   3030
      _ExtentX        =   13785
      _ExtentY        =   4260
      Max             =   1
   End
   Begin Mp3RadioPlus.PicScroll Balance2 
      Height          =   225
      Left            =   8020
      TabIndex        =   8
      Top             =   8090
      Width           =   660
      _ExtentX        =   13785
      _ExtentY        =   4260
      Min             =   -100
   End
   Begin Mp3RadioPlus.PicScroll Volume3 
      Height          =   615
      Left            =   5280
      TabIndex        =   14
      Top             =   2760
      Width           =   1365
      _ExtentX        =   13785
      _ExtentY        =   4260
      Min             =   -360
      Max             =   0
      Value           =   -180
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10080
      Top             =   1440
   End
   Begin VB.Timer tmLED2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10440
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3240
      Top             =   1440
   End
   Begin VB.Timer tmLED1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3600
      Top             =   1440
   End
   Begin Mp3RadioPlus.PicScroll Mixer 
      Height          =   600
      Left            =   4560
      TabIndex        =   27
      Top             =   3555
      Width           =   2835
      _ExtentX        =   13785
      _ExtentY        =   4260
      Min             =   -100
   End
   Begin Mp3RadioPlus.PicScroll SegM 
      Height          =   600
      Left            =   7630
      TabIndex        =   28
      Top             =   3555
      Width           =   2820
      _ExtentX        =   13785
      _ExtentY        =   4260
      Max             =   20
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3495
      Left            =   120
      TabIndex        =   30
      Top             =   5040
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Artist"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Genre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "File"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.PictureBox cmdPLAY2On 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   9800
      ScaleHeight     =   855
      ScaleWidth      =   1020
      TabIndex        =   50
      Top             =   2040
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox cmdPLAY2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   9800
      ScaleHeight     =   855
      ScaleWidth      =   1020
      TabIndex        =   52
      Top             =   2040
      Width           =   1020
   End
   Begin VB.PictureBox cmdCUE2On 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   8720
      ScaleHeight     =   855
      ScaleWidth      =   1020
      TabIndex        =   53
      Top             =   2040
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox cmdCUE2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   8720
      ScaleHeight     =   855
      ScaleWidth      =   1020
      TabIndex        =   56
      Top             =   2040
      Width           =   1020
   End
   Begin VB.PictureBox cmdCUE1On 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   1880
      ScaleHeight     =   855
      ScaleWidth      =   1020
      TabIndex        =   54
      Top             =   2040
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox cmdCUE1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   1880
      ScaleHeight     =   855
      ScaleWidth      =   1020
      TabIndex        =   55
      Top             =   2040
      Width           =   1020
   End
   Begin VB.PictureBox Mixt 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Index           =   5
      Left            =   2445
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   57
      Top             =   3720
      Width           =   465
   End
   Begin VB.PictureBox Mixt 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Index           =   4
      Left            =   1980
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   58
      Top             =   3720
      Width           =   465
   End
   Begin VB.PictureBox Mixt 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Index           =   3
      Left            =   1510
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   59
      Top             =   3720
      Width           =   465
   End
   Begin VB.PictureBox Mixt 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Index           =   2
      Left            =   1050
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   60
      Top             =   3720
      Width           =   465
   End
   Begin VB.PictureBox Mixt 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Index           =   1
      Left            =   585
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   61
      Top             =   3720
      Width           =   465
   End
   Begin VB.PictureBox Mixt 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Index           =   0
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   62
      Top             =   3720
      Width           =   465
   End
   Begin VB.PictureBox CAuto 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   10940
      ScaleHeight     =   480
      ScaleWidth      =   900
      TabIndex        =   63
      Top             =   3615
      Width           =   900
   End
   Begin VB.PictureBox loopOFF1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3440
      ScaleHeight     =   285
      ScaleWidth      =   540
      TabIndex        =   65
      Top             =   3030
      Width           =   540
   End
   Begin VB.PictureBox loopOFF2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   10280
      ScaleHeight     =   285
      ScaleWidth      =   540
      TabIndex        =   66
      Top             =   3030
      Width           =   540
   End
   Begin VB.PictureBox cmdPLAY1On 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2960
      ScaleHeight     =   855
      ScaleWidth      =   1020
      TabIndex        =   49
      Top             =   2040
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox cmdPLAY1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2960
      ScaleHeight     =   855
      ScaleWidth      =   1020
      TabIndex        =   51
      Top             =   2040
      Width           =   1020
   End
   Begin Mp3RadioPlus.PicVScroll Pitch1 
      Height          =   2295
      Left            =   4380
      TabIndex        =   71
      Top             =   630
      Width           =   120
      _ExtentX        =   5583
      _ExtentY        =   13996
      Min             =   -80
      Max             =   80
   End
   Begin Mp3RadioPlus.PicVScroll Pitch2 
      Height          =   2295
      Left            =   11220
      TabIndex        =   73
      Top             =   630
      Width           =   120
      _ExtentX        =   5583
      _ExtentY        =   13996
      Min             =   -80
      Max             =   80
   End
   Begin VB.PictureBox VolR1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00008000&
      ForeColor       =   &H00008000&
      Height          =   100
      Left            =   250
      ScaleHeight     =   105
      ScaleWidth      =   2295
      TabIndex        =   68
      Top             =   700
      Width           =   2295
   End
   Begin VB.PictureBox VolL1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00008000&
      ForeColor       =   &H00008000&
      Height          =   100
      Left            =   250
      ScaleHeight     =   105
      ScaleWidth      =   2295
      TabIndex        =   67
      Top             =   850
      Width           =   2295
   End
   Begin VB.PictureBox VolR2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00008000&
      ForeColor       =   &H00008000&
      Height          =   100
      Left            =   7090
      ScaleHeight     =   105
      ScaleWidth      =   2295
      TabIndex        =   69
      Top             =   700
      Width           =   2295
   End
   Begin VB.PictureBox VolL2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00008000&
      ForeColor       =   &H00008000&
      Height          =   100
      Left            =   7090
      ScaleHeight     =   105
      ScaleWidth      =   2295
      TabIndex        =   70
      Top             =   850
      Width           =   2295
   End
   Begin VB.Image cmdPitchB2 
      Height          =   270
      Left            =   11250
      Top             =   240
      Width           =   270
   End
   Begin VB.Image cmdPitchA2 
      Height          =   270
      Left            =   10980
      Top             =   240
      Width           =   270
   End
   Begin VB.Image cmdPitchC2 
      Height          =   270
      Left            =   11515
      Top             =   240
      Width           =   270
   End
   Begin VB.Image cmdPitchC1 
      Height          =   270
      Left            =   4680
      Top             =   240
      Width           =   270
   End
   Begin VB.Image cmdPitchA1 
      Height          =   265
      Left            =   4130
      Top             =   240
      Width           =   265
   End
   Begin VB.Label lblPitch2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0.0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   11070
      TabIndex        =   74
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblPitch1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0.0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   4230
      TabIndex        =   72
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image cmdSkin 
      Height          =   450
      Left            =   11360
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label CautoON 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Height          =   480
      Left            =   10940
      TabIndex        =   64
      Top             =   3615
      Width           =   900
   End
   Begin VB.Label lblRemaining2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8290
      TabIndex        =   32
      Top             =   140
      Width           =   1215
   End
   Begin VB.Label lblElapsed2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Elapsed"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6970
      TabIndex        =   31
      Top             =   140
      Width           =   1215
   End
   Begin VB.Label ledREM2 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "-0:00:00"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8290
      TabIndex        =   33
      Top             =   290
      Width           =   1215
   End
   Begin VB.Label ledELA2 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "0:00:00"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6970
      TabIndex        =   34
      Top             =   290
      Width           =   1215
   End
   Begin VB.Label ledNAME2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Player 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7090
      TabIndex        =   35
      Top             =   1095
      Width           =   3615
   End
   Begin VB.Label LedVERSION2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   9490
      TabIndex        =   36
      Top             =   860
      Width           =   1215
   End
   Begin VB.Label LedMODO2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   9490
      TabIndex        =   37
      Top             =   140
      Width           =   1215
   End
   Begin VB.Label LedKHZ2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   9490
      TabIndex        =   38
      Top             =   620
      Width           =   1215
   End
   Begin VB.Label LedKBPS2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   9490
      TabIndex        =   39
      Top             =   380
      Width           =   1215
   End
   Begin VB.Label ledNAME1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Player 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   250
      TabIndex        =   40
      Top             =   1100
      Width           =   3585
   End
   Begin VB.Label ledELA1 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "0:00:00"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   130
      TabIndex        =   41
      Top             =   290
      Width           =   1215
   End
   Begin VB.Label ledREM1 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "-0:00:00"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1450
      TabIndex        =   42
      Top             =   290
      Width           =   1215
   End
   Begin VB.Label lblElapsed1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Elapsed"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblRemaining1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1450
      TabIndex        =   48
      Top             =   140
      Width           =   1215
   End
   Begin VB.Label LedKBPS1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   2650
      TabIndex        =   47
      Top             =   380
      Width           =   1215
   End
   Begin VB.Label LedKHZ1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   2655
      TabIndex        =   46
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label LedMODO1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   2650
      TabIndex        =   45
      Top             =   140
      Width           =   1215
   End
   Begin VB.Label LedVERSION1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   2650
      TabIndex        =   44
      Top             =   860
      Width           =   1215
   End
   Begin VB.Image srtCol 
      Height          =   255
      Index           =   3
      Left            =   5685
      Top             =   4800
      Width           =   1050
   End
   Begin VB.Image srtCol 
      Height          =   255
      Index           =   2
      Left            =   4905
      Top             =   4800
      Width           =   780
   End
   Begin VB.Image srtCol 
      Height          =   255
      Index           =   1
      Left            =   2880
      Top             =   4800
      Width           =   2025
   End
   Begin VB.Image srtCol 
      Height          =   255
      Index           =   0
      Left            =   120
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Image cmdDOWN 
      Height          =   450
      Left            =   6240
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image cmdUP 
      Height          =   450
      Left            =   5640
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image cmdINF 
      Height          =   450
      Left            =   3720
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image cmdSELALL 
      Height          =   450
      Left            =   3120
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image cmdREMOVE 
      Height          =   450
      Left            =   2520
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image cmdADD 
      Height          =   450
      Left            =   1920
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image cmdSAVE 
      Height          =   450
      Left            =   1320
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image cmdLOAD 
      Height          =   450
      Left            =   720
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image cmdCLEAR 
      Height          =   450
      Left            =   120
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image cmdFade 
      Height          =   390
      Left            =   3120
      Top             =   3705
      Width           =   1185
   End
   Begin VB.Label lblFadeTime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   10530
      TabIndex        =   29
      Top             =   3750
      Width           =   255
   End
   Begin VB.Label cmdSAMPLE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   9
      Left            =   6240
      TabIndex        =   26
      Top             =   1210
      Width           =   420
   End
   Begin VB.Label cmdSAMPLE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   8
      Left            =   5770
      TabIndex        =   25
      Top             =   1210
      Width           =   420
   End
   Begin VB.Label cmdSAMPLE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   7
      Left            =   5290
      TabIndex        =   24
      Top             =   1210
      Width           =   420
   End
   Begin VB.Label cmdSAMPLE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   6
      Left            =   6240
      TabIndex        =   23
      Top             =   1570
      Width           =   420
   End
   Begin VB.Label cmdSAMPLE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   5
      Left            =   5770
      TabIndex        =   22
      Top             =   1570
      Width           =   420
   End
   Begin VB.Label cmdSAMPLE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   4
      Left            =   5290
      TabIndex        =   21
      Top             =   1570
      Width           =   420
   End
   Begin VB.Label cmdSAMPLE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   3
      Left            =   6240
      TabIndex        =   20
      Top             =   1930
      Width           =   420
   End
   Begin VB.Label cmdSAMPLE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   2
      Left            =   5770
      TabIndex        =   19
      Top             =   1930
      Width           =   420
   End
   Begin VB.Label cmdSAMPLE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   1
      Left            =   5290
      TabIndex        =   18
      Top             =   1930
      Width           =   420
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackColor       =   &H00789678&
      BackStyle       =   0  'Transparent
      Caption         =   "31/12/02"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5320
      TabIndex        =   17
      Top             =   200
      Width           =   1335
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00789678&
      BackStyle       =   0  'Transparent
      Caption         =   "23:59"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5325
      TabIndex        =   16
      Top             =   550
      Width           =   1335
   End
   Begin VB.Label cmdLOAD3 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   5760
      TabIndex        =   15
      Top             =   2280
      Width           =   900
   End
   Begin VB.Label cmdSAMPLE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   0
      Left            =   5290
      TabIndex        =   13
      Top             =   2290
      Width           =   420
   End
   Begin VB.Image cmdLOAD2 
      Height          =   390
      Left            =   6960
      Top             =   1560
      Width           =   1185
   End
   Begin VB.Image cmdCUP2 
      Height          =   600
      Left            =   7740
      Top             =   2160
      Width           =   780
   End
   Begin VB.Label ledPLAY2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   10160
      TabIndex        =   11
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label ledCUE2 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   9060
      TabIndex        =   10
      Top             =   1800
      Width           =   255
   End
   Begin VB.Image cmdSERBACK2 
      Height          =   390
      Left            =   6960
      Top             =   2040
      Width           =   585
   End
   Begin VB.Image cmdSERFORW2 
      Height          =   390
      Left            =   6960
      Top             =   2520
      Width           =   585
   End
   Begin VB.Label lblVol2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "-Inf"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   8020
      TabIndex        =   9
      Top             =   5330
      Width           =   615
   End
   Begin VB.Label lblVol1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "-Inf"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   7000
      TabIndex        =   4
      Top             =   5330
      Width           =   615
   End
   Begin VB.Image cmdSERFORW1 
      Height          =   390
      Left            =   120
      Top             =   2520
      Width           =   585
   End
   Begin VB.Image cmdSERBACK1 
      Height          =   390
      Left            =   120
      Top             =   2040
      Width           =   585
   End
   Begin VB.Label ledCUE1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   2235
      TabIndex        =   3
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label ledPLAY1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3315
      TabIndex        =   2
      Top             =   1800
      Width           =   255
   End
   Begin VB.Image cmdCUP1 
      Height          =   600
      Left            =   900
      Top             =   2160
      Width           =   780
   End
   Begin VB.Image cmdLOAD1 
      Height          =   390
      Left            =   120
      Top             =   1560
      Width           =   1185
   End
   Begin VB.Image loopON1 
      Height          =   285
      Left            =   3440
      Top             =   3030
      Width           =   540
   End
   Begin VB.Image loopON2 
      Height          =   285
      Left            =   10280
      Top             =   3030
      Width           =   540
   End
   Begin MediaPlayerCtl.MediaPlayer MPV 
      Height          =   495
      Index           =   0
      Left            =   6240
      TabIndex        =   12
      Top             =   2040
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
   Begin VB.Image cmdPitchB1 
      Height          =   270
      Left            =   4400
      Top             =   240
      Width           =   270
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mud As Boolean
Dim MixType As Integer
Dim Play1 As Boolean
Dim Play2 As Boolean
Dim Loop1 As Boolean
Dim Loop2 As Boolean
Dim CrossAuto As Boolean
Dim Loaded1 As Boolean
Dim Loaded2 As Boolean
Dim Mix1 As Integer
Dim Mix2 As Integer
Dim A2B As Boolean
Dim StCA As Boolean
Dim ColorL As Long
Dim ColorR As Long
Dim STFreq As Long
Dim strFile1 As String
Dim strFile2 As String
Dim ActiveStream1 As Boolean
Dim ActiveStream2 As Boolean
Dim SampleData(1000) As Integer, nDataSize As Integer
Dim itmX As ListItem
Dim id3MP3 As New MP3Info
Dim CUE1
Dim CUE2
Function Filtro(Texto) As String
    Pos = Len(Texto) - InStrRev(Texto, "\")
    tmpfiltro = Right(Texto, Pos)
    Filtro = tmpfiltro
End Function
Private Sub Balance1_Change()
DMC.StreamPan = Balance1.Value * 100
End Sub
Private Sub Balance1_LeftClick()
Balance1.Value = 0
DMC.StreamPan = 0
End Sub
Private Sub Balance2_Change()
DMC.Stream2Pan = Balance2.Value * 100
End Sub
Private Sub Balance2_LeftClick()
Balance2.Value = 0
DMC.Stream2Pan = 0
End Sub
Private Sub CAuto_Click()
If Play1 = Not Play2 Then
    If Play1 = True Then
        Mixer.Value = -100
    Else
        Mixer.Value = 100
    End If
    Mixer_Change
    CAuto.Visible = False
    CrossAuto = True
End If
End Sub
Private Sub CautoON_Click()
CAuto.Visible = True
CrossAuto = False
End Sub
Private Sub cmdADD_Click()
CommonDialog1.MaxFileSize = 16384
CommonDialog1.FileName = ""
CommonDialog1.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
Dim DirInicial As String
Dim Ngenre As Integer
DirInicial = GetSetting("ACP Software", "Radio Plus", "Dir", "C:\")
If Dir(DirInicial, vbDirectory) = "" Then DirInicial = "C:\"
CommonDialog1.InitDir = DirInicial
CommonDialog1.Filter = "MP3|*.mp3"
CommonDialog1.ShowOpen
MousePointer = 11
If CommonDialog1.FileName <> "" Then
    Dim vFiles As Variant
    Dim lFile As Long
    vFiles = Split(CommonDialog1.FileName, Chr(0))
    If UBound(vFiles) = 0 Then
        SaveSetting "ACP Software", "Radio Plus", "Dir", FiltroDir(CommonDialog1.FileName)
        id3MP3.FileName = CommonDialog1.FileName
        id3MP3.GetMPEGInfo
        If GetId3(CommonDialog1.FileName) = True Then
            Ngenre = Int(Trim(id3Info.Genre))
            If Ngenre > 147 Then Ngenre = 13
            Set itmX = ListView1.ListItems.Add(, , Trim(id3Info.Title))
            itmX.SubItems(1) = Trim(id3Info.Artist)
            itmX.SubItems(2) = GenreArray(Ngenre)
            itmX.SubItems(3) = Pos2Time(id3MP3.Seconds)
            itmX.SubItems(4) = CommonDialog1.FileName
        Else
            Set itmX = ListView1.ListItems.Add(, , Filtro(CommonDialog1.FileName))
            itmX.SubItems(1) = ""
            itmX.SubItems(2) = ""
            itmX.SubItems(3) = Pos2Time(id3MP3.Seconds)
            itmX.SubItems(4) = CommonDialog1.FileName
        End If
    Else
        Dim DirE As String
        If Right(vFiles(0), 1) = "\" Then
            DirE = vFiles(0)
        Else
            DirE = vFiles(0) & "\"
        End If
        SaveSetting "ACP Software", "Radio Plus", "Dir", DirE
        For lFile = 1 To UBound(vFiles)
            id3MP3.FileName = DirE & vFiles(lFile)
            id3MP3.GetMPEGInfo
            If GetId3(DirE & vFiles(lFile)) = True Then
                Ngenre = Int(Trim(id3Info.Genre))
                If Ngenre > 147 Then Ngenre = 13
                Set itmX = ListView1.ListItems.Add(, , Trim(id3Info.Title))
                itmX.SubItems(1) = Trim(id3Info.Artist)
                itmX.SubItems(2) = GenreArray(Ngenre)
                itmX.SubItems(3) = Pos2Time(id3MP3.Seconds)
                itmX.SubItems(4) = DirE & vFiles(lFile)
            Else
                Set itmX = ListView1.ListItems.Add(, , Filtro(DirE & vFiles(lFile)))
                itmX.SubItems(1) = ""
                itmX.SubItems(2) = ""
                itmX.SubItems(3) = Pos2Time(id3MP3.Seconds)
                itmX.SubItems(4) = DirE & vFiles(lFile)
            End If
        Next lFile
    End If
End If
MousePointer = 0
End Sub
Private Sub cmdCLEAR_Click()
ListView1.ListItems.Clear
End Sub
Private Sub cmdCUE1_Click()
If Loaded1 = True Then
    If Play1 = True Then
        DMC.StopStream
        DMC.StreamPos = CUE1
        Mud = True
        Position1.Value = DMC.StreamPos / 100
        Mud = False
        Play1 = False
        PicSpectrum1.Cls
        VolL1.Cls
        VolR1.Cls
    Else
        CUE1 = DMC.StreamPos
        tmLED1.Enabled = False
    End If
    cmdPLAY1On.Visible = False
    cmdCUE1On.Visible = True
    ledPLAY1.BackStyle = 0
    ledCUE1.BackStyle = 1
End If
End Sub
Private Sub cmdCUE1On_Click()
If Loaded1 = True Then cmdCUE1_Click
End Sub
Private Sub cmdCUE2_Click()
If Loaded2 = True Then
    If Play2 = True Then
        DMC.StopStream2
        DMC.Stream2Pos = CUE2
        Mud = True
        Position2.Value = DMC.Stream2Pos / 100
        Mud = False
        Play2 = False
        PicSpectrum2.Cls
        VolL2.Cls
        VolR2.Cls
    Else
        CUE2 = DMC.Stream2Pos
        tmLED2.Enabled = False
    End If
    cmdPLAY2On.Visible = False
    cmdCUE2On.Visible = True
    ledPLAY2.BackStyle = 0
    ledCUE2.BackStyle = 1
End If
End Sub
Private Sub cmdCUE2On_Click()
If Loaded2 = True Then cmdCUE2_Click
End Sub
Private Sub cmdCUP1_Click()
If Loaded1 = True Then
    DMC.StreamPos = CUE1
    If Play1 = False Then
        cmdPLAY1_Click
    End If
End If
End Sub
Private Sub cmdCUP2_Click()
If Loaded2 = True Then
    DMC.Stream2Pos = CUE2
    If Play2 = False Then
        cmdPLAY2_Click
    End If
End If
End Sub
Private Sub cmdDOWN_Click()
If ListView1.SelectedItem Is Nothing Then Exit Sub
If ListView1.SelectedItem.Index < ListView1.ListItems.Count Then
    dcx = ListView1.SelectedItem.Index + 1
    Set itmX = ListView1.ListItems.Add(dcx + 1, , ListView1.SelectedItem)
    itmX.SubItems(1) = ListView1.SelectedItem.SubItems(1)
    itmX.SubItems(2) = ListView1.SelectedItem.SubItems(2)
    itmX.SubItems(3) = ListView1.SelectedItem.SubItems(3)
    itmX.SubItems(4) = ListView1.SelectedItem.SubItems(4)
    ListView1.ListItems.Remove ListView1.SelectedItem.Index
    ListView1.ListItems(dcx).Selected = True
End If
End Sub
Private Sub cmdFade_Click()
If Mixer.Value < 0 Then
    A2B = True
Else
    A2B = False
End If
TimerManualFade.Interval = (SegM.Value * 1000) / (200 / 5)
TimerManualFade.Enabled = True
End Sub
Private Sub cmdINF_Click()
If ListView1.SelectedItem Is Nothing Then
    Exit Sub
Else
    If LCase(strFile1) = LCase(ListView1.SelectedItem.SubItems(4)) Or LCase(strFile2) = LCase(ListView1.SelectedItem.SubItems(4)) Then
        MsgBox "The selected file '" & ListView1.SelectedItem.SubItems(4) & "' is playing and our ID3 cannot be edited!", vbOKOnly, "ID3"
    Else
        frmID3.GID3 ListView1.SelectedItem.SubItems(4)
    End If
End If
End Sub
Private Sub cmdLOAD_Click()
CommonDialog1.Filter = "PLS|*.PLS"
Dim DirInicial As String
Dim strFile As String
DirInicial = GetSetting("ACP Software", "Radio Plus", "Dir", "C:\")
If Dir(DirInicial, vbDirectory) = "" Then DirInicial = "C:\"
CommonDialog1.InitDir = DirInicial
CommonDialog1.ShowOpen
MousePointer = 11
If CommonDialog1.FileName <> "" Then
    SaveSetting "ACP Software", "Radio Plus", "Dir", FiltroDir(CommonDialog1.FileName)
    ListView1.ListItems.Clear
    For I = 1 To Int(ReadINI(CommonDialog1.FileName, "playlist", "NumberOfEntries"))
        strFile = ReadINI(CommonDialog1.FileName, "playlist", "File" & I)
        If Dir(strFile) <> "" Then
            id3MP3.FileName = strFile
            id3MP3.GetMPEGInfo
            If GetId3(strFile) = True Then
                Ngenre = Int(Trim(id3Info.Genre))
                If Ngenre > 147 Then Ngenre = 13
                Set itmX = ListView1.ListItems.Add(, , Trim(id3Info.Title))
                itmX.SubItems(1) = Trim(id3Info.Artist)
                itmX.SubItems(2) = GenreArray(Ngenre)
                itmX.SubItems(3) = Pos2Time(id3MP3.Seconds)
                itmX.SubItems(4) = strFile
            Else
                Set itmX = ListView1.ListItems.Add(, , Filtro(strFile))
                itmX.SubItems(1) = ""
                itmX.SubItems(2) = ""
                itmX.SubItems(3) = Pos2Time(id3MP3.Seconds)
                itmX.SubItems(4) = strFile
            End If
        End If
    Next I
End If
MousePointer = 0
End Sub
Private Sub cmdLOAD1_Click()
If ListView1.SelectedItem Is Nothing Then Exit Sub
Timer1.Enabled = True
Play1 = False
strFile1 = ListView1.SelectedItem.SubItems(4)
id3MP3.FileName = strFile1
id3MP3.GetMPEGInfo
If Trim(ListView1.SelectedItem.SubItems(1)) <> "" Then
    ledNAME1.Caption = ListView1.SelectedItem & "   #   " & ListView1.SelectedItem.SubItems(1)
Else
    ledNAME1.Caption = ListView1.SelectedItem
End If
LedMODO1.Caption = id3MP3.Mode
LedKBPS1.Caption = id3MP3.BitRate & " kbps"
LedKHZ1.Caption = id3MP3.Frequency & " khz"
LedVERSION1.Caption = id3MP3.VersionLayer
Loaded1 = True
DMC.OpenStream strFile1
DMC.PlayStream False
DMC.PauseStream
cmdCUE1_Click
cmdCUE1On.Visible = True
Position1.Value = 0
If Mud = True Then
    Mud = False
End If
ActiveStream1 = True
Position1.Max = DMC.StreamLen / 100
End Sub
Private Sub cmdLOAD2_Click()
If ListView1.SelectedItem Is Nothing Then Exit Sub
Timer2.Enabled = True
Play2 = False
strFile2 = ListView1.SelectedItem.SubItems(4)
id3MP3.FileName = strFile2
id3MP3.GetMPEGInfo
If Trim(ListView1.SelectedItem.SubItems(1)) <> "" Then
    ledNAME2.Caption = ListView1.SelectedItem & "   #   " & ListView1.SelectedItem.SubItems(1)
Else
    ledNAME2.Caption = ListView1.SelectedItem
End If
LedMODO2.Caption = id3MP3.Mode
LedKBPS2.Caption = id3MP3.BitRate & " kbps"
LedKHZ2.Caption = id3MP3.Frequency & " khz"
LedVERSION2.Caption = id3MP3.VersionLayer
Loaded2 = True
DMC.OpenStream2 strFile2
DMC.PlayStream2 False
DMC.PauseStream2
cmdCUE2_Click
Position2.Value = 0
If Mud = True Then
    Mud = False
End If
ActiveStream2 = True
Position2.Max = DMC.Stream2Len / 100
End Sub
Private Sub cmdLOAD3_Click()
frmSample.Show vbModal
End Sub
Private Sub cmdPitchA1_Click()
If Pitch1.Value > 80 Then Pitch1.Value = 80
If Pitch1.Value < -80 Then Pitch1.Value = -80
Pitch1.Max = 80
Pitch1.Min = -80
Pitch1_Change
End Sub
Private Sub cmdPitchA2_Click()
If Pitch2.Value > 80 Then Pitch2.Value = 80
If Pitch2.Value < -80 Then Pitch2.Value = -80
Pitch2.Max = 80
Pitch2.Min = -80
Pitch2_Change
End Sub
Private Sub cmdPitchB1_Click()
If Pitch1.Value > 160 Then Pitch1.Value = 160
If Pitch1.Value < -160 Then Pitch1.Value = -160
Pitch1.Max = 160
Pitch1.Min = -160
Pitch1_Change
End Sub
Private Sub cmdPitchB2_Click()
If Pitch2.Value > 160 Then Pitch2.Value = 160
If Pitch2.Value < -160 Then Pitch2.Value = -160
Pitch2.Max = 160
Pitch2.Min = -160
Pitch2_Change
End Sub
Private Sub cmdPitchC1_Click()
Pitch1.Max = 240
Pitch1.Min = -240
Pitch1_Change
End Sub
Private Sub cmdPitchC2_Click()
Pitch2.Max = 240
Pitch2.Min = -240
Pitch2_Change
End Sub
Private Sub cmdPLAY1_Click()
If Loaded1 = True Then
    cmdPLAY1On.Visible = True
    cmdCUE1On.Visible = False
    If Play1 = True Then
        Play1 = False
        tmLED1.Enabled = True
        DMC.PauseStream
    Else
        Play1 = True
        tmLED1.Enabled = False
        ledPLAY1.BackStyle = 1
        ledCUE1.BackStyle = 0
        If DMC.StreamIsPaused Then
            DMC.ResumeStream
        Else
            DMC.PlayStream False
        End If
        Pitch1_Change
        Balance1_Change
        Volume1_Change
    End If
End If
End Sub
Private Sub cmdPLAY1On_Click()
If Loaded1 = True Then cmdPLAY1_Click
End Sub
Private Sub cmdPLAY2_Click()
If Loaded2 = True Then
    cmdPLAY2On.Visible = True
    cmdCUE2On.Visible = False
    If Play2 = True Then
        Play2 = False
        tmLED2.Enabled = True
        DMC.PauseStream2
    Else
        Play2 = True
        tmLED2.Enabled = False
        ledPLAY2.BackStyle = 1
        ledCUE2.BackStyle = 0
        If DMC.Stream2IsPaused Then
            DMC.ResumeStream2
        Else
            DMC.PlayStream2 False
        End If
        Pitch2_Change
        Balance2_Change
        Volume2_Change
    End If
End If
End Sub
Private Sub cmdPLAY2On_Click()
If Loaded2 = True Then cmdPLAY2_Click
End Sub
Private Sub cmdREMOVE_Click()
For I = ListView1.ListItems.Count To 1 Step -1
    If ListView1.ListItems(I).Selected = True Then
        ListView1.ListItems.Remove I
    End If
Next I
End Sub
Private Sub cmdSave_Click()
CommonDialog1.Filter = "PLS|*.PLS"
Dim DirInicial As String
DirInicial = GetSetting("ACP Software", "Radio Plus", "Dir", "C:\")
If Dir(DirInicial, vbDirectory) = "" Then DirInicial = "C:\"
CommonDialog1.InitDir = DirInicial
CommonDialog1.ShowSave
MousePointer = 11
If CommonDialog1.FileName <> "" Then
    SaveSetting "ACP Software", "Radio Plus", "Dir", FiltroDir(CommonDialog1.FileName)
    For I = 1 To ListView1.ListItems.Count
        id3MP3.FileName = ListView1.ListItems(I).SubItems(4)
        id3MP3.GetMPEGInfo
        WriteINI CommonDialog1.FileName, "playlist", "File" & I, ListView1.ListItems(I).SubItems(4)
        WriteINI CommonDialog1.FileName, "playlist", "Title" & I, ListView1.ListItems(I)
        WriteINI CommonDialog1.FileName, "playlist", "Length" & I, id3MP3.Seconds
    Next I
    WriteINI CommonDialog1.FileName, "playlist", "NumberOfEntries", "" & ListView1.ListItems.Count
    WriteINI CommonDialog1.FileName, "playlist", "Version", "2"
End If
MousePointer = 0
End Sub
Private Sub cmdSELALL_Click()
For I = 1 To ListView1.ListItems.Count
    ListView1.ListItems(I).Selected = True
Next I
End Sub

Private Sub cmdSkin_Click()
frmSkin.Show vbModal
End Sub

''NOT FINISHED YET
'Private Sub cmdSERBACK1_Click()
'If Loaded1 = True Then
'
'End If
'End Sub
'Private Sub cmdSERBACK2_Click()
'If Loaded2 = True Then
'
'End If
'End Sub
'Private Sub cmdSERFORW1_Click()
'If Loaded1 = True Then
'
'End If
'End Sub
'Private Sub cmdSERFORW2_Click()
'If Loaded2 = True Then
'
'End If
'End Sub
''NOT FINISHED YET
Private Sub cmdUP_Click()
If ListView1.SelectedItem Is Nothing Then Exit Sub
If ListView1.SelectedItem.Index > 1 Then
    dcx = ListView1.SelectedItem.Index - 1
    Set itmX = ListView1.ListItems.Add(dcx, , ListView1.SelectedItem)
    itmX.SubItems(1) = ListView1.SelectedItem.SubItems(1)
    itmX.SubItems(2) = ListView1.SelectedItem.SubItems(2)
    itmX.SubItems(3) = ListView1.SelectedItem.SubItems(3)
    itmX.SubItems(4) = ListView1.SelectedItem.SubItems(4)
    ListView1.ListItems.Remove ListView1.SelectedItem.Index
    ListView1.ListItems(dcx).Selected = True
End If
End Sub
Private Sub cmdSample_Click(Index As Integer)
If cmdSAMPLE(Index).ForeColor <> Hex2VB(ReadINI(Arq0, "Colors", "SampleD")) Then
    If cmdSAMPLE(Index).ForeColor = Hex2VB(ReadINI(Arq0, "Colors", "SampleE")) Then
        MPV(Index).Play
        cmdSAMPLE(Index).ForeColor = Hex2VB(ReadINI(Arq0, "Colors", "SampleP"))
    Else
        MPV(Index).Stop
        MPV(Index).CurrentPosition = 0
        cmdSAMPLE(Index).ForeColor = Hex2VB(ReadINI(Arq0, "Colors", "SampleE"))
    End If
End If
End Sub
Private Sub DMC_Stream2Stoped(ByVal paused As Boolean)
If paused = False Then
    If Loop2 = True Then
        DMC.Stream2Pos = CUE2
        DMC.PlayStream2 False
    Else
        cmdPLAY2On.Visible = False
        cmdCUE2On.Visible = True
        DMC.Stream2Pos = CUE2
        Mud = True
        Position2.Value = DMC.Stream2Pos / 100
        Mud = False
        ledPLAY2.BackStyle = 0
        ledCUE2.BackStyle = 1
        Play2 = False
        If CrossAuto = True Then
            If SegM.Value = 0 Then
                If Loaded1 = True Then cmdPLAY1_Click
                Mixer.Value = 100
                Mixer_Change
                cmdCUE2_Click
            End If
        End If
    End If
End If
End Sub
Private Sub DMC_StreamStoped(ByVal paused As Boolean)
If paused = False Then
    If Loop1 = True Then
        DMC.StreamPos = CUE1
        DMC.PlayStream False
    Else
        cmdPLAY1On.Visible = False
        cmdCUE1On.Visible = True
        DMC.StreamPos = CUE1
        Mud = True
        Position1.Value = DMC.StreamPos / 100
        Mud = False
        ledPLAY1.BackStyle = 0
        ledCUE1.BackStyle = 1
        Play1 = False
        If CrossAuto = True Then
            If SegM.Value = 0 Then
                If Loaded2 = True Then cmdPLAY2_Click
                Mixer.Value = 100
                Mixer_Change
                cmdCUE1_Click
            End If
        End If
    End If
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    cmdPLAY1_Click
    cmdPLAY2_Click
End If
If KeyCode = vbKeyEscape Then
    cmdCUE1_Click
    cmdCUE2_Click
End If
If KeyCode = vbKeyNumpad0 Then cmdSample_Click (0)
If KeyCode = vbKeyNumpad1 Then cmdSample_Click (1)
If KeyCode = vbKeyNumpad2 Then cmdSample_Click (2)
If KeyCode = vbKeyNumpad3 Then cmdSample_Click (3)
If KeyCode = vbKeyNumpad4 Then cmdSample_Click (4)
If KeyCode = vbKeyNumpad5 Then cmdSample_Click (5)
If KeyCode = vbKeyNumpad6 Then cmdSample_Click (6)
If KeyCode = vbKeyNumpad7 Then cmdSample_Click (7)
If KeyCode = vbKeyNumpad8 Then cmdSample_Click (8)
If KeyCode = vbKeyNumpad9 Then cmdSample_Click (9)
End Sub
Public Sub SetSkin()
Dim Arq As String
Dim Arq1 As String
Arq = ReadINI(App.Path & "\RdPlus.ini", "Skin", "Skin")
Arq0 = App.Path & "\Skins\" & Arq & ".ini"
Arq1 = App.Path & "\Skins\" & ReadINI(Arq0, "Images", "Main")
Color0 = Hex2VB(ReadINI(Arq0, "Colors", "ListBk"))
Color1 = Hex2VB(ReadINI(Arq0, "Colors", "ListTx"))
Color2 = Hex2VB(ReadINI(Arq0, "Colors", "Elapsed"))
Color3 = Hex2VB(ReadINI(Arq0, "Colors", "Remaining"))
Color4 = Hex2VB(ReadINI(Arq0, "Colors", "NameBk"))
Color5 = Hex2VB(ReadINI(Arq0, "Colors", "NameTx"))
Color6 = Hex2VB(ReadINI(Arq0, "Colors", "Values"))
Color7 = Hex2VB(ReadINI(Arq0, "Colors", "Info"))
Color8 = Hex2VB(ReadINI(Arq0, "Colors", "Date"))
ColorR = Hex2VB(ReadINI(Arq0, "Colors", "SpectrumR"))
ColorL = Hex2VB(ReadINI(Arq0, "Colors", "SpectrumL"))
Me.Picture = LoadPicture(Arq1)
Pitch1.PaintPicture1 Me.Picture, 0, 0, , , 1575, 6390, 285, 150
Pitch1.PaintPicture2 Me.Picture, 0, 0, , , 4380, 630, 285, 2295
Pitch1.PaintPicture3 Me.Picture, 0, 0, , , 11220, 630, 285, 2295
Pitch2.PaintPicture1 Me.Picture, 0, 0, , , 1575, 6390, 285, 150
Pitch2.PaintPicture2 Me.Picture, 0, 0, , , 4380, 630, 285, 2295
Pitch2.PaintPicture3 Me.Picture, 0, 0, , , 11220, 630, 285, 2295
Volume1.PaintPicture1 Me.Picture, 0, 0, , , 1575, 6390, 285, 150
Volume1.PaintPicture2 Me.Picture, 0, 0, , , 7180, 5600, 285, 2295
Volume1.PaintPicture3 Me.Picture, 0, 0, , , 8200, 5600, 285, 2295
Volume2.PaintPicture1 Me.Picture, 0, 0, , , 1575, 6390, 285, 150
Volume2.PaintPicture2 Me.Picture, 0, 0, , , 7180, 5600, 285, 2295
Volume2.PaintPicture3 Me.Picture, 0, 0, , , 8200, 5600, 285, 2295
Volume3.PaintPicture1 Me.Picture, 0, 0, , , 2820, 6015, 105, 225
Volume3.PaintPicture2 Me.Picture, 0, 0, , , 5280, 2760, 1365, 615
Volume3.PaintPicture3 Me.Picture, 0, 0, , , 195, 6390, 1365, 615
Position1.PaintPicture1 Me.Picture, 0, 0, , , 3000, 5130, 150, 285
Position1.PaintPicture2 Me.Picture, 0, 0, , , 240, 3000, 3030, 375
Position1.PaintPicture3 Me.Picture, 0, 0, , , 7060, 3000, 6030, 5505
Position2.PaintPicture1 Me.Picture, 0, 0, , , 3000, 5130, 150, 285
Position2.PaintPicture2 Me.Picture, 0, 0, , , 240, 3000, 3030, 375
Position2.PaintPicture3 Me.Picture, 0, 0, , , 7060, 3000, 6030, 5505
Balance1.PaintPicture1 Me.Picture, 0, 0, , , 2820, 6015, 105, 225
Balance1.PaintPicture2 Me.Picture, 0, 0, , , 7000, 8090, 660, 225
Balance1.PaintPicture3 Me.Picture, 0, 0, , , 7000, 8090, 660, 225
Balance2.PaintPicture1 Me.Picture, 0, 0, , , 2820, 6015, 105, 225
Balance2.PaintPicture2 Me.Picture, 0, 0, , , 7000, 8090, 660, 225
Balance2.PaintPicture3 Me.Picture, 0, 0, , , 7000, 8090, 660, 225
Mixer.PaintPicture1 Me.Picture, 0, 0, , , 3000, 5130, 150, 285
Mixer.PaintPicture2 Me.Picture, 0, 0, , , 4560, 3555, 2835, 600
Mixer.PaintPicture3 Me.Picture, 0, 0, , , 4560, 3555, 2835, 600
SegM.PaintPicture1 Me.Picture, 0, 0, , , 3000, 5130, 150, 285
SegM.PaintPicture2 Me.Picture, 0, 0, , , 7630, 3555, 2820, 600
SegM.PaintPicture3 Me.Picture, 0, 0, , , 7630, 3555, 2820, 600
cmdPLAY1.PaintPicture Me.Picture, 0, 0, , , 2960, 2040, 1020, 855
cmdPLAY2.PaintPicture Me.Picture, 0, 0, , , 2960, 2040, 1020, 855
cmdPLAY1On.PaintPicture Me.Picture, 0, 0, , , 1230, 5520, 1020, 855
cmdPLAY2On.PaintPicture Me.Picture, 0, 0, , , 1230, 5520, 1020, 855
cmdCUE1.PaintPicture Me.Picture, 0, 0, , , 1880, 2040, 1020, 855
cmdCUE2.PaintPicture Me.Picture, 0, 0, , , 1880, 2040, 1020, 855
cmdCUE1On.PaintPicture Me.Picture, 0, 0, , , 195, 5520, 1020, 855
cmdCUE2On.PaintPicture Me.Picture, 0, 0, , , 195, 5520, 1020, 855
CAuto.PaintPicture Me.Picture, 0, 0, , , 2265, 5520, 900, 480
loopOFF1.PaintPicture Me.Picture, 0, 0, , , 2265, 6015, 540, 285
loopOFF2.PaintPicture Me.Picture, 0, 0, , , 2265, 6015, 540, 285
PicSpectrum1.PaintPicture Me.Picture, 0, 0, , , 195, 7015, 2345, 450
PicSpectrum2.PaintPicture Me.Picture, 0, 0, , , 195, 7015, 2345, 450
PicSpectrum1.Picture = PicSpectrum1.Image
PicSpectrum2.Picture = PicSpectrum2.Image
VolR1.PaintPicture Me.Picture, 0, 0, , , VolR2.Left, VolR2.Top, VolR2.Width, VolR2.Height
VolL1.PaintPicture Me.Picture, 0, 0, , , VolL2.Left, VolL2.Top, VolL2.Width, VolL2.Height
VolR2.PaintPicture Me.Picture, 0, 0, , , VolR2.Left, VolR2.Top, VolR2.Width, VolR2.Height
VolL2.PaintPicture Me.Picture, 0, 0, , , VolL2.Left, VolL2.Top, VolL2.Width, VolL2.Height
VolR1.Picture = VolR1.Image
VolL1.Picture = VolL1.Image
VolR2.Picture = VolR2.Image
VolL2.Picture = VolL2.Image
For I = 0 To 5
    Mixt(I).PaintPicture Me.Picture, 0, 0, , , ((31 * I) + 13) * 15, 5130, 465, 345
Next I
ListView1.BackColor = Color0
ListView1.ForeColor = Color1
lblElapsed1.ForeColor = Color2
lblElapsed2.ForeColor = Color2
lblRemaining1.ForeColor = Color3
lblRemaining2.ForeColor = Color3
lblDate.ForeColor = Color8
lblTime.ForeColor = Color8
lblVol1.ForeColor = Color6
lblVol2.ForeColor = Color6
lblPitch1.ForeColor = Color6
lblPitch2.ForeColor = Color6
lblFadeTime.ForeColor = Color6
ledELA1.ForeColor = Color2
ledELA2.ForeColor = Color2
ledREM1.ForeColor = Color3
ledREM2.ForeColor = Color3
ledNAME1.BackColor = Color4
ledNAME1.ForeColor = Color5
ledNAME2.BackColor = Color4
ledNAME2.ForeColor = Color5
LedKBPS1.ForeColor = Color7
LedKBPS2.ForeColor = Color7
LedKHZ1.ForeColor = Color7
LedKHZ2.ForeColor = Color7
LedMODO1.ForeColor = Color7
LedMODO2.ForeColor = Color7
LedVERSION1.ForeColor = Color7
LedVERSION2.ForeColor = Color7
End Sub

Private Sub Form_Load()
STFreq = 44100
DMC.DeviceToUse = -1
DMC.InitBASS Me.hWnd, STFreq, False, False
DMC.BufferLenInSeconds = 1#


SetSkin
GenreArray = Split(sGenreMatrix, "|")
ListView1.ColumnHeaders(5).Width = 0
ListView1.ColumnHeaders(4).Width = 800
ListView1.ColumnHeaders(3).Width = 800
ListView1.ColumnHeaders(2).Width = 2000
ListView1.ColumnHeaders(1).Width = 2760
If (Screen.Height / Screen.TwipsPerPixelY) = 600 Then Me.WindowState = 2
KeyPreview = True
For I = 1 To 9
    Load MPV(I)
Next I
Dim StrSample As String
For I = 0 To 9
    StrSample = GetSetting("ACP Software", "Radio Plus", "Sample" & I, "")
    If StrSample <> "" Then
        MPV(I).FileName = StrSample
        cmdSAMPLE(I).ForeColor = Hex2VB(ReadINI(Arq0, "Colors", "SampleE"))
    Else
        cmdSAMPLE(I).ForeColor = Hex2VB(ReadINI(Arq0, "Colors", "SampleD"))
    End If
Next I
SegM.Value = Int(GetSetting("ACP Software", "Radio Plus", "FadeTime", "10"))
SegM_Change
Volume1.Value = Volume1.Min / 2
Volume2.Value = Volume2.Min / 2
MixType = Int(GetSetting("ACP Software", "Radio Plus", "MixType", "0"))
Mixt(MixType).Visible = False
Mixer_Change
Volume3.Value = Volume3.Min / 2
Volume3_Change
Balance1.Value = 0
Balance2.Value = 0
Balance1_Change
Balance2_Change
Loaded1 = False
Loaded2 = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
DMC.TerminateBASS
End Sub

Private Sub Horas_Timer()
lblTime.Caption = Format(Now, "hh:mm")
lblDate.Caption = Format(Now, "dd/mm/yy")
'Me.Caption = "Radio Plus   -   CPU: " & Format$(DMC.Info_UsedCPU(), "#0.00") & "%"
End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ListView1.Sorted = True
ListView1.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    For I = 1 To ListView1.ListItems.Count 'Goes through all items In the listView
'        'checks to see if the mouse is over the
'        '     current listView item
'        If (x > ListView1.ListItems.Item(I).Left) And _
'        (x < (ListView1.ListItems.Item(I).Left + ListView1.ListItems.Item(I).Width)) _
'        And (y > ListView1.ListItems.Item(I).Top) And _
'        (y < ListView1.ListItems.Item(I).Top + ListView1.ListItems.Item(I).Height) Then
'        'if it is, set all to default, in this c
'        '     ase, black
'
'            If ListView1.ListItems.Item(I).Bold = False Then
'                For b = 1 To ListView1.ListItems.Count
'                    ListView1.ListItems.Item(b).Bold = False
'                Next b
'                'sets the one that the mouse is over to
'                '     Blue, can be changed.
'                ListView1.ListItems.Item(I).Bold = True
'            End If
'        End If
'    Next I
End Sub
Private Sub loopOFF1_Click()
loopOFF1.Visible = False
Loop1 = True
End Sub
Private Sub loopOFF2_Click()
loopOFF2.Visible = False
Loop2 = True
End Sub
Private Sub loopON1_Click()
loopOFF1.Visible = True
Loop1 = False
End Sub
Private Sub loopON2_Click()
loopOFF2.Visible = True
Loop2 = False
End Sub
Private Sub Mixer_Change()
Select Case MixType
    Case 0
        If Mixer.Value = 0 Then
            Mix1 = 100
            Mix2 = 100
        ElseIf Mixer.Value < 0 Then
            Mix1 = 100
            Mix2 = 0
        Else
            Mix1 = 0
            Mix2 = 100
        End If
    Case 1
        Mix1 = Int((Mixer.Value - 100) / 2)
        If Mix1 < 0 Then Mix1 = -Mix1
        Mix2 = 100 - Mix1
    Case 2
        Mix1 = -Mixer.Value
        Mix2 = Mixer.Value
        If Mix1 < 0 Then Mix1 = 0
        If Mix2 < 0 Then Mix2 = 0
    Case 3
        Mix1 = -(Mixer.Value - 100)
        Mix2 = (Mixer.Value + 100)
        If Mix1 > 100 Then Mix1 = 100
        If Mix2 > 100 Then Mix2 = 100
    Case 4
        If Mixer.Value = 100 Then
            Mix1 = 0
        Else
            Mix1 = 100
        End If
        Mix2 = (Mixer.Value + 100) / 2
    Case 5
        If Mixer.Value = -100 Then
            Mix2 = 0
        Else
            Mix2 = 100
        End If
        Mix1 = -((Mixer.Value - 100) / 2)
End Select
Volume1_Change
Volume2_Change
End Sub
Private Sub Mixer_ErroValue()
TimerCross.Enabled = False
StCA = False
If A2B = True Then
    A2B = False
    cmdCUE1_Click
Else
    A2B = True
    cmdCUE2_Click
End If
End Sub
Private Sub Mixer_LeftClick()
Mixer.Value = 0
Mixer_Change
End Sub
Private Sub Mixt_Click(Index As Integer)
SaveSetting "ACP Software", "Radio Plus", "MixType", Index
Mixt(MixType).Visible = True
Mixt(Index).Visible = False
MixType = Index
Mixer.Value = 0
Mixer_Change
End Sub
Private Sub MPV_EndOfStream(Index As Integer, ByVal Result As Long)
cmdSAMPLE(Index).ForeColor = Hex2VB(ReadINI(Arq0, "Colors", "SampleE"))
End Sub

Private Sub PicSpectrum1_Click()
PicSpectrum1.Visible = False
End Sub

Private Sub PicSpectrum2_Click()
PicSpectrum2.Visible = False
End Sub

Private Sub Pitch1_Change()
Ptv = Pitch1.Value / 10
If InStr(Ptv, ",") = 0 Then
    lblPitch1.Caption = Ptv & ",0"
Else
    lblPitch1.Caption = Ptv
End If
DMC.StreamFreq = STFreq + ((Ptv * STFreq) / 100)
End Sub
Private Sub Pitch1_LeftClick()
Pitch1.Value = 0
Pitch1_Change
End Sub
Private Sub Pitch2_Change()
Ptv = Pitch2.Value / 10
If InStr(Ptv, ",") = 0 Then
    lblPitch2.Caption = Ptv & ",0"
Else
    lblPitch2.Caption = Ptv
End If
DMC.Stream2Freq = STFreq + ((Ptv * STFreq) / 100)
End Sub
Private Sub Pitch2_LeftClick()
Pitch2.Value = 0
Pitch2_Change
End Sub
Private Sub Position1_Change()
If Mud = False Then DMC.StreamPos = Position1.Value * 100
End Sub
Private Sub Position2_Change()
If Mud = False Then DMC.Stream2Pos = Position2.Value * 100
End Sub
Private Sub SegM_Change()
lblFadeTime.Caption = SegM.Value
End Sub
Private Sub SegM_LeftClick()
SegM.Value = 10
lblFadeTime.Caption = SegM.Value
End Sub
Private Sub SegM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
SaveSetting "ACP Software", "Radio Plus", "FadeTime", SegM.Value
End Sub
Private Sub srtCol_Click(Index As Integer)
If ListView1.Sorted = True Then
    If ListView1.SortKey = Index Then
        If ListView1.SortOrder = lvwAscending Then
            ListView1.SortOrder = lvwDescending
        Else
            ListView1.SortOrder = lvwAscending
        End If
    Else
        ListView1.SortKey = Index
        ListView1.SortOrder = lvwAscending
    End If
Else
    ListView1.Sorted = True
    ListView1.SortKey = Index
    ListView1.SortOrder = lvwAscending
End If
End Sub
Private Sub Timer1_Timer()
ledELA1.Caption = Pos2Time(DMC.StreamPosInSeconds)
ledREM1.Caption = "-" & Pos2Time(DMC.StreamLenInSeconds - DMC.StreamPosInSeconds)
If CrossAuto = True Then
    If (DMC.StreamLenInSeconds - DMC.StreamPosInSeconds) <= SegM.Value Then
    If SegM.Value <> 0 Then
        If StCA = False Then
            StCA = True
            A2B = True
            Mixer.Value = -100
            Mixer_Change
            If Loaded2 = True Then cmdPLAY2_Click
            TimerCross.Enabled = True
        End If
    End If
    End If
End If
End Sub
Function Pos2Time(Pos)
Dim Hour As Integer
Dim Min As String
Dim Sec As String
Hour = Int((Pos \ 60) \ 60)
Min = Int(Pos \ 60)
Sec = Int(Pos - (Min * 60))
If Hour > 0 Then
    Min = Min - (Hour * 60)
End If
If Sec = "-1" Then Sec = "0"
If Min < 10 Then Min = "0" & Min
If Sec < 10 Then Sec = "0" & Sec
If Hour = 0 Then
    Pos2Time = Min & ":" & Sec
Else
    Pos2Time = Hour & ":" & Min & ":" & Sec
End If
End Function
Private Sub Timer2_Timer()
ledELA2.Caption = Pos2Time(DMC.Stream2PosInSeconds)
ledREM2.Caption = "-" & Pos2Time(DMC.Stream2LenInSeconds - DMC.Stream2PosInSeconds)
If CrossAuto = True Then
    If (DMC.Stream2LenInSeconds - DMC.Stream2PosInSeconds) <= SegM.Value Then
    If SegM.Value <> 0 Then
        If StCA = False Then
            StCA = True
            A2B = False
            Mixer.Value = 100
            Mixer_Change
            If Loaded1 = True Then cmdPLAY1_Click
            TimerCross.Enabled = True
        End If
    End If
    End If
End If
End Sub
Private Sub Timer3_Timer()
If DMC.StreamIsActive = True Then
    Mud = True
    Position1.Value = DMC.StreamPos / 100
    Mud = False
    If PicSpectrum1.Visible = True Then
        DrawScope1
    Else
        If ActiveStream1 = True Then
            VolL1.Cls
            VolR1.Cls
            posL = Int((VolL1.Width / 128) * DMC.StreamLeftLevel)
            posR = Int((VolR1.Width / 128) * DMC.StreamRightLevel)
            If posL = 0 Then
                VolL1.Cls
            Else
                VolL1.PaintPicture Me.Picture, 0, 0, , , VolL1.Left, VolL1.Top, posL, VolL1.Height
            End If
            If posR = 0 Then
                VolR1.Cls
            Else
                VolR1.PaintPicture Me.Picture, 0, 0, , , VolR1.Left, VolR1.Top, posR, VolR1.Height
            End If
        End If
    End If
End If
End Sub
Private Sub Timer4_Timer()
If DMC.Stream2IsActive = True Then
    Mud = True
    Position2.Value = DMC.Stream2Pos / 100
    Mud = False
    If PicSpectrum2.Visible = True Then
        DrawScope2
    Else
        If ActiveStream2 = True Then
            VolL2.Cls
            VolR2.Cls
            posL = Int((VolL2.Width / 128) * DMC.Stream2LeftLevel)
            posR = Int((VolR2.Width / 128) * DMC.Stream2RightLevel)
            If posL = 0 Then
                VolL2.Cls
            Else
                VolL2.PaintPicture Me.Picture, 0, 0, , , VolL1.Left, VolL1.Top, posL, VolL1.Height
            End If
            If posR = 0 Then
                VolR2.Cls
            Else
                VolR2.PaintPicture Me.Picture, 0, 0, , , VolR1.Left, VolR1.Top, posR, VolR1.Height
            End If
        End If
    End If
End If
End Sub
Private Sub TimerCross_Timer()
If A2B = True Then
    If Mixer.Value = 100 Then
        A2B = False
        cmdCUE1_Click
        TimerCross.Enabled = False
        StCA = False
    Else
        Mvalue = -Int((((DMC.StreamLenInSeconds - DMC.StreamPosInSeconds) * 100 / SegM.Value) * 2) - 100)
        If Mvalue > 100 Then
            Mixer.Value = 100
        Else
            Mixer.Value = Mvalue
        End If
        Mixer_Change
    End If
Else
    If Mixer.Value = -100 Then
        A2B = True
        cmdCUE2_Click
        TimerCross.Enabled = False
        StCA = False
    Else
        Mvalue = Int((((DMC.Stream2LenInSeconds - DMC.Stream2PosInSeconds) * 100 / SegM.Value) * 2) - 100)
        If Mvalue < -100 Then
            Mixer.Value = -100
        Else
            Mixer.Value = Mvalue
        End If
        Mixer_Change
    End If
End If
End Sub
Private Sub TimerManualFade_Timer()
If A2B = True Then
    If Mixer.Value = 100 Then
        TimerManualFade.Enabled = False
    Else
        If (Mixer.Value + 5) > 100 Then
            Mixer.Value = 100
        Else
            Mixer.Value = Mixer.Value + 5
        End If
        Mixer_Change
    End If
Else
    If Mixer.Value = -100 Then
        TimerManualFade.Enabled = False
    Else
        If (Mixer.Value - 5) < -100 Then
            Mixer.Value = 100
        Else
            Mixer.Value = Mixer.Value - 5
        End If
        Mixer_Change
    End If
End If
End Sub
Private Sub tmLED1_Timer()
If ledPLAY1.BackStyle = 1 Then
    ledPLAY1.BackStyle = 0
Else
    ledPLAY1.BackStyle = 1
End If
End Sub
Private Sub tmLED2_Timer()
If ledPLAY2.BackStyle = 1 Then
    ledPLAY2.BackStyle = 0
Else
    ledPLAY2.BackStyle = 1
End If
End Sub
Private Sub VolL1_Click()
PicSpectrum1.Visible = True
End Sub
Private Sub VolL2_Click()
PicSpectrum2.Visible = True
End Sub
Private Sub VolR1_Click()
PicSpectrum1.Visible = True
End Sub
Private Sub VolR2_Click()
PicSpectrum2.Visible = True
End Sub
Private Sub Volume1_Change()
Dim dB As String
If Volume1.Value = 0 Then
    lblVol1.Caption = "-Inf"
Else
    dB = ((Volume1.Value + 360) / 10)
    If dB = 0 Then
        lblVol1.Caption = "0"
    Else
        If InStr(dB, ",") = 0 Then
            lblVol1.Caption = -dB & ",0"
        Else
            lblVol1.Caption = -dB
        End If
    End If
End If
DMC.StreamVol = -((Volume1.Value * Mix1) / 360)
End Sub
Private Sub Volume2_Change()
Dim dB As String
If Volume2.Value = 0 Then
    lblVol2.Caption = "-Inf"
Else
    dB = ((Volume2.Value + 360) / 10)
    If dB = 0 Then
        lblVol2.Caption = "0"
    Else
        If InStr(dB, ",") = 0 Then
            lblVol2.Caption = -dB & ",0"
        Else
            lblVol2.Caption = -dB
        End If
    End If
End If
DMC.Stream2Vol = -((Volume2.Value * Mix2) / 360)
End Sub
Private Sub Volume3_Change()
If Volume3.Value = -360 Then
    For I = 0 To MPV.UBound
        MPV(I).Mute = True
    Next I
Else
    For I = 0 To MPV.UBound
        MPV(I).Mute = False
        MPV(I).Volume = Volume3.Value * 10
    Next I
End If
End Sub
Private Sub Volume3_LeftClick()
Volume3.Value = Volume3.Min / 2
Volume3_Change
End Sub
Function FiltroDir(Texto) As String
    Dim Pos As Integer
    Pos = InStrRev(Texto, "\")
    FiltroDir = Left(Texto, Pos)
End Function
Private Sub DrawScope1()
Static X As Integer, vh As Integer, h As Integer
Static picWidth%, picHeight%
picWidth = PicSpectrum1.ScaleWidth
picHeight = PicSpectrum1.ScaleHeight
If DMC.StreamIsMono Then
    nDataSize = 500
Else
    nDataSize = 1000
End If
DMC.StreamData SampleData, nDataSize
PicSpectrum1.Cls
vh = picHeight
If DMC.StreamIsMono Then
    'left channel
    For I = 0 To 499 Step 1
       h = ((SampleData(I) + 32768) / 65535 * vh)
       X = (picWidth * I * 2) / nDataSize
       If I = 0 Then PicSpectrum1.PSet (0, h)
       PicSpectrum1.Line -(X, h), ColorL
    Next
Else
    'left channel
    For I = 0 To 499 Step 1
       h = ((SampleData(I) + 32768) / 65535 * vh)
       X = (picWidth * I * 2) / nDataSize
       If I = 0 Then PicSpectrum1.PSet (0, h)
       PicSpectrum1.Line -(X, h), ColorL
    Next
    'right channel
    For I = 1 To 499 Step 2
       h = ((SampleData(I) + 32768) / 65535 * vh)
       X = (picWidth * I) / 500
       If I = 1 Then PicSpectrum1.PSet (0, h)
       PicSpectrum1.Line -(X, h), ColorR
    Next
End If
End Sub
Private Sub DrawScope2()
Static X As Integer, vh As Integer, h As Integer
Static picWidth%, picHeight%
picWidth = PicSpectrum2.ScaleWidth
picHeight = PicSpectrum2.ScaleHeight
If DMC.Stream2IsMono Then
    nDataSize = 500
Else
    nDataSize = 1000
End If
DMC.Stream2Data SampleData, nDataSize
PicSpectrum2.Cls
vh = picHeight
If DMC.Stream2IsMono Then
    'left channel
    For I = 0 To 499 Step 1
       h = ((SampleData(I) + 32768) / 65535 * vh)
       X = (picWidth * I * 2) / nDataSize
       If I = 0 Then PicSpectrum2.PSet (0, h)
       PicSpectrum2.Line -(X, h), ColorL
    Next
Else
    'left channel
    For I = 0 To 499 Step 1
       h = ((SampleData(I) + 32768) / 65535 * vh)
       X = (picWidth * I * 2) / nDataSize
       If I = 0 Then PicSpectrum2.PSet (0, h)
       PicSpectrum2.Line -(X, h), ColorL
    Next
    'right channel
    For I = 1 To 499 Step 2
       h = ((SampleData(I) + 32768) / 65535 * vh)
       X = (picWidth * I) / 500
       If I = 1 Then PicSpectrum2.PSet (0, h)
       PicSpectrum2.Line -(X, h), ColorR
    Next
End If
End Sub

