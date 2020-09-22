VERSION 5.00
Begin VB.Form frmSkin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Skin"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   4440
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.PictureBox picPreview 
      BackColor       =   &H8000000C&
      Height          =   2460
      Left            =   3720
      ScaleHeight     =   2400
      ScaleWidth      =   3300
      TabIndex        =   0
      Top             =   1440
      Width           =   3360
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      TabIndex        =   8
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label lblName 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label lblTitulo2 
      BackStyle       =   0  'Transparent
      Caption         =   "Info:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblTitulo1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblTitulo3 
      BackStyle       =   0  'Transparent
      Caption         =   "Preview:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
End
Attribute VB_Name = "frmSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If List1.ListIndex > -1 Then
    WriteINI App.Path & "\RdPlus.ini", "Skin", "Skin", List1.Text
    Me.Visible = False
    DoEvents
    frmMain.SetSkin
    Unload Me
End If
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Command3_Click()
frmAbout.Show vbModal
End Sub
Private Sub Form_Load()
MyPath = App.Path & "\Skins\"
MyName = Dir(MyPath)
Do While MyName <> ""
    If LCase(Right(MyName, 4)) = ".ini" Then
       List1.AddItem Left(MyName, Len(MyName) - 4)
    End If
   MyName = Dir
Loop
End Sub
Private Sub List1_Click()
If List1.ListIndex > -1 Then
    lblName.Caption = ReadINI(App.Path & "\Skins\" & List1.Text & ".ini", "Skin", "Name")
    lblInfo.Caption = ReadINI(App.Path & "\Skins\" & List1.Text & ".ini", "Skin", "Info")
    pPreview = ReadINI(App.Path & "\Skins\" & List1.Text & ".ini", "Images", "Preview")
    If pPreview <> "" Then picPreview.Picture = LoadPicture(App.Path & "\Skins\" & pPreview)
End If
End Sub
