VERSION 5.00
Begin VB.Form frmID3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ID3"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
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
      Height          =   330
      Left            =   1320
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtFILE 
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
      Left            =   720
      TabIndex        =   5
      Top             =   120
      Width           =   3735
   End
   Begin VB.TextBox txtARTIST 
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
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   4
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox txtALBUM 
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
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   3
      Top             =   1680
      Width           =   3015
   End
   Begin VB.TextBox txtYEAR 
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
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   2
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txtCOMMENTS 
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
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   1
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox txtTITLE 
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
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   0
      Top             =   720
      Width           =   3015
   End
   Begin VB.Image cmdSave 
      Height          =   330
      Left            =   1200
      Top             =   3720
      Width           =   975
   End
   Begin VB.Image cmdOk 
      Height          =   330
      Left            =   2520
      Top             =   3720
      Width           =   975
   End
End
Attribute VB_Name = "frmID3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub GID3(strFile As String)
txtFILE.Text = strFile
If GetId3(txtFILE.Text) = True Then
    txtTITLE.Text = Trim(id3Info.Title)
    txtARTIST.Text = Trim(id3Info.Artist)
    txtALBUM.Text = Trim(id3Info.Album)
    txtYEAR.Text = Trim(id3Info.sYear)
    txtCOMMENTS.Text = Trim(id3Info.Comments)
    If id3Info.Genre > (Combo1.ListCount - 2) Then
        Combo1.ListIndex = 0
    Else
        For I = 0 To (Combo1.ListCount - 1)
            If Trim(Combo1.List(I)) = Trim(GenreArray(id3Info.Genre)) Then
                Combo1.ListIndex = I
                Exit For
            End If
        Next I
    End If
End If
frmID3.Show vbModal
End Sub
Private Sub cmdOk_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
id3Info.Title = txtTITLE.Text
id3Info.Artist = txtARTIST.Text
id3Info.Album = txtALBUM.Text
id3Info.sYear = txtYEAR.Text
id3Info.Comments = txtCOMMENTS.Text
If Combo1.ListIndex = 0 Then
    id3Info.Genre = 255
Else
    For I = 0 To UBound(GenreArray)
        If Trim(Combo1.Text) = Trim(GenreArray(I)) Then
            id3Info.Genre = I
            Exit For
        End If
    Next I
End If
SaveId3 txtFILE.Text, id3Info
frmMain.ListView1.SelectedItem = txtTITLE.Text
frmMain.ListView1.SelectedItem.SubItems(1) = txtARTIST.Text
frmMain.ListView1.SelectedItem.SubItems(2) = Combo1.Text
End Sub
Private Sub Form_Load()
Dim Arq As String
Dim Arq0 As String
Dim Arq1 As String
Arq = ReadINI(App.Path & "\RdPlus.ini", "Skin", "Skin")
Arq0 = App.Path & "\Skins\" & Arq & ".ini"
Arq1 = App.Path & "\Skins\" & ReadINI(Arq0, "Images", "Id3")
Color0 = Hex2VB(ReadINI(Arq0, "Colors", "ListBk"))
Color1 = Hex2VB(ReadINI(Arq0, "Colors", "ListTx"))
txtFILE.BackColor = Color0
txtTITLE.BackColor = Color0
txtARTIST.BackColor = Color0
txtALBUM.BackColor = Color0
txtYEAR.BackColor = Color0
txtCOMMENTS.BackColor = Color0
Combo1.BackColor = Color0
txtFILE.ForeColor = Color1
txtTITLE.ForeColor = Color1
txtARTIST.ForeColor = Color1
txtALBUM.ForeColor = Color1
txtYEAR.ForeColor = Color1
txtCOMMENTS.ForeColor = Color1
Combo1.ForeColor = Color1
Me.Picture = LoadPicture(Arq1)

Combo1.AddItem "", 0
For I = 0 To UBound(GenreArray)
    Combo1.AddItem GenreArray(I)
Next I
End Sub
