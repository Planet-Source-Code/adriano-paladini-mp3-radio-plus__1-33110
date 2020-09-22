VERSION 5.00
Begin VB.UserControl PicVScroll 
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10140
   ScaleHeight     =   9195
   ScaleWidth      =   10140
   Begin VB.PictureBox picBack1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   3480
      ScaleHeight     =   7815
      ScaleWidth      =   3765
      TabIndex        =   2
      Top             =   120
      Width           =   3765
   End
   Begin VB.PictureBox picBar 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   7560
      ScaleHeight     =   2295
      ScaleWidth      =   2205
      TabIndex        =   1
      Top             =   120
      Width           =   2205
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   7935
      Left            =   0
      ScaleHeight     =   7935
      ScaleWidth      =   3165
      TabIndex        =   0
      Top             =   0
      Width           =   3165
   End
End
Attribute VB_Name = "PicVScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

' Declarations
Dim iY As Long
Dim bDrag As Boolean
Dim iMin As Long
Dim iMax As Long
Dim iValue As Long

' Events
Event Change()
Event LeftClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Property Get BackColor() As OLE_COLOR
    BackColor = picBack.BackColor
End Property

Public Property Let BackColor(ByVal New_Color As OLE_COLOR)
    picBack.BackColor = New_Color
    Call DrawBar
    
    PropertyChanged "BackColor"
End Property



Public Property Get Bar() As Picture
    Set Bar = picBar.Picture
End Property

Public Property Set Bar(ByVal New_Bar As Picture)
    Set picBar.Picture = New_Bar
    UserControl_Resize
    
    
    Call DrawBar
    PropertyChanged "Bar"
End Property


Private Sub CalcValue()
    iValue = iY / (ToPixels(picBack.Height) - ToPixels(picBar.Height)) * (iMax - iMin) + iMin
End Sub

Private Sub CalcX()
    iY = (iValue - iMin) / (iMax - iMin) * (ToPixels(picBack.Height) - ToPixels(picBar.Height))
End Sub

Private Sub DrawBar(Optional CalculateX As Boolean = True)
    If CalculateX Then Call CalcX
    
    picBack.Cls
    Call BitBlt(picBack.hDC, 0, iY, picBack1.ScaleWidth, picBack1.ScaleHeight, picBack1.hDC, 0, iY, vbSrcCopy)
    Call BitBlt(picBack.hDC, 0, iY, picBar.ScaleWidth, picBar.ScaleHeight, picBar.hDC, 0, 0, vbSrcCopy)
    picBack.Refresh
    
    
    UserControl.Refresh
End Sub
Public Property Get Max() As Long
    Max = iMax
End Property

Public Property Let Max(New_Max As Long)
    If New_Max < iValue Then
        MsgBox "Maximum exceeds value!", vbOKOnly + vbExclamation, "Error"
        Exit Property
    End If
    
    iMax = New_Max
    Call DrawBar
    
    PropertyChanged "Max"
End Property

Public Property Get Min() As Long
    Min = iMin
End Property

Public Property Let Min(New_Min As Long)
    If iMin > iValue Then
        MsgBox "Minimum exceeds value!"
        Exit Property
    End If
    
    iMin = New_Min
    Call DrawBar
    
    PropertyChanged "Min"
End Property

Public Property Get Picture() As Picture
    Set Picture = picBack.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set picBack.Picture = New_Picture
    UserControl_Resize
    Call DrawBar
    
    PropertyChanged "Picture"
End Property
Public Property Get Picture2() As Picture
    Set Picture2 = picBack1.Picture
End Property

Public Property Set Picture2(ByVal New_Picture2 As Picture)
    Set picBack1.Picture = New_Picture2
    UserControl_Resize
    Call DrawBar
    
    PropertyChanged "Picture2"
End Property


Private Function ToPixels(ByVal nTwips As Long) As Long
    ToPixels = nTwips / Screen.TwipsPerPixelY
End Function

Public Property Get Value() As Long
    Value = iValue
End Property

Public Property Let Value(New_Value As Long)
    If New_Value < iMin Or New_Value > iMax Then
        MsgBox "Value exceeds limits!", vbOKOnly + vbExclamation, "Error"
        Exit Property
    End If
    
    iValue = New_Value
    Call DrawBar
    
    PropertyChanged "Value"
End Property
Private Sub picBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    RaiseEvent LeftClick
Else
    If ToPixels(Y) >= iY And ToPixels(Y) <= iY + ToPixels(picBar.Height) And Button = 1 Then
        bDrag = True
    Else
        bDrag = True
        iY = ToPixels(Y)
        
        If iY > ToPixels(picBack.Height) - (ToPixels(picBar.Height) / 2) Then iY = ToPixels(picBack.Height) - (ToPixels(picBar.Height) / 2)
        If iY < ToPixels(picBar.Height) / 2 Then iY = ToPixels(picBar.Height) / 2
        
        iY = iY - ToPixels(picBar.Height) / 2
        
        Call DrawBar(False)
        Call CalcValue
        Value = iValue
        RaiseEvent Change
        
    End If
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
End If
End Sub


Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bDrag Then
        iY = ToPixels(Y)
        
        If iY > ToPixels(picBack.Height) - (ToPixels(picBar.Height) / 2) Then iY = ToPixels(picBack.Height) - (ToPixels(picBar.Height) / 2)
        If iY < ToPixels(picBar.Height) / 2 Then iY = ToPixels(picBar.Height) / 2
        
        iY = iY - ToPixels(picBar.Height) / 2
        
        Call DrawBar(False)
        Call CalcValue
        Value = iValue
        RaiseEvent Change
    End If
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picBack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bDrag = False
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
Private Sub UserControl_Initialize()
    If iMax = 0 Then iMax = 100
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    picBack.Picture = PropBag.ReadProperty("Picture", Nothing)
    picBack1.Picture = PropBag.ReadProperty("Picture2", Nothing)
    picBar.Picture = PropBag.ReadProperty("Bar", Nothing)
    picBack.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    iMin = PropBag.ReadProperty("Min", 0)
    iMax = PropBag.ReadProperty("Max", 100)
    iValue = PropBag.ReadProperty("Value", 0)
    
    Call DrawBar
End Sub

Private Sub UserControl_Resize()
    picBar.Top = UserControl.Height + picBar.Height
    picBar.Left = 0
    picBack1.Left = 0
    picBack1.Top = UserControl.Height + picBack1.Height
    'UserControl.Width = picBack.Width
    'UserControl.Height = picBack.Height
    ' Resize picBack
'    With picBack
'        .Top = 0
'        .Left = 0
'        .Width = UserControl.Width
'        .Height = UserControl.Height
'    End With
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Picture", picBack.Picture, Nothing)
    Call PropBag.WriteProperty("Picture2", picBack1.Picture, Nothing)
    Call PropBag.WriteProperty("Bar", picBar.Picture, Nothing)
    Call PropBag.WriteProperty("BackColor", picBack.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Min", iMin, 0)
    Call PropBag.WriteProperty("Max", iMax, 100)
    Call PropBag.WriteProperty("Value", iValue, 0)
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picBar,picBack,-1,PaintPicture
Public Sub PaintPicture1(ByVal Picture As Picture, ByVal X1 As Single, ByVal Y1 As Single, Optional ByVal Width1 As Variant, Optional ByVal Height1 As Variant, Optional ByVal X2 As Variant, Optional ByVal Y2 As Variant, Optional ByVal Width2 As Variant, Optional ByVal Height2 As Variant, Optional ByVal Opcode As Variant)
Attribute PaintPicture1.VB_Description = "Draws the contents of a graphics file on a Form, PictureBox, or Printer object."
    picBar.PaintPicture Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2, Opcode
    picBar.Picture = picBar.Image
    picBar.Height = Height2
    picBar.Width = Width2
    UserControl_Resize
    Call DrawBar
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picBack,picBack,-1,PaintPicture
Public Sub PaintPicture2(ByVal Picture As Picture, ByVal X1 As Single, ByVal Y1 As Single, Optional ByVal Width1 As Variant, Optional ByVal Height1 As Variant, Optional ByVal X2 As Variant, Optional ByVal Y2 As Variant, Optional ByVal Width2 As Variant, Optional ByVal Height2 As Variant, Optional ByVal Opcode As Variant)
    picBack.PaintPicture Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2, Opcode
    picBack.Picture = picBack.Image
    picBack.Height = Height2
    picBack.Width = Width2
    UserControl.Width = picBack.Width
    UserControl.Height = picBack.Height
    UserControl_Resize
    Call DrawBar
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picBack1,picBack,-1,PaintPicture
Public Sub PaintPicture3(ByVal Picture As Picture, ByVal X1 As Single, ByVal Y1 As Single, Optional ByVal Width1 As Variant, Optional ByVal Height1 As Variant, Optional ByVal X2 As Variant, Optional ByVal Y2 As Variant, Optional ByVal Width2 As Variant, Optional ByVal Height2 As Variant, Optional ByVal Opcode As Variant)
    picBack1.PaintPicture Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2, Opcode
    picBack1.Picture = picBack1.Image
    picBack1.Height = Height2
    picBack1.Width = Width2
    UserControl_Resize
    Call DrawBar
End Sub

