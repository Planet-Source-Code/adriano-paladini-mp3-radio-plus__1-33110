VERSION 5.00
Begin VB.UserControl PlayButton 
   ClientHeight    =   1635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2325
   ScaleHeight     =   1635
   ScaleWidth      =   2325
   Begin VB.PictureBox Image2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox Image1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   0
      Width           =   315
   End
End
Attribute VB_Name = "PlayButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Private Sub Image1_Click()
    RaiseEvent Click
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Image1,Image1,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = Image1.Visible
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Image1.Visible() = New_Enabled
    PropertyChanged "Enabled"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Image2,Image2,-1,Enabled
Public Property Get PStop() As Boolean
    PStop = Image2.Visible
End Property
Public Property Let PStop(ByVal New_Stop As Boolean)
    Image2.Visible() = New_Stop
    PropertyChanged "Stop"
End Property
Private Sub Image2_Click()
RaiseEvent Click
End Sub
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Image1.Enabled = PropBag.ReadProperty("Enabled", True)
    Image2.Enabled = PropBag.ReadProperty("PStop", False)
End Sub
Private Sub UserControl_Resize()
UserControl.Height = Image1.Height
UserControl.Width = Image1.Width
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", Image1.Enabled, True)
    Call PropBag.WriteProperty("PStop", Image2.Enabled, False)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picBack,picBack,-1,PaintPicture
Public Sub PaintPictureStop(ByVal Picture As Picture, ByVal X1 As Single, ByVal Y1 As Single, Optional ByVal Width1 As Variant, Optional ByVal Height1 As Variant, Optional ByVal X2 As Variant, Optional ByVal Y2 As Variant, Optional ByVal Width2 As Variant, Optional ByVal Height2 As Variant, Optional ByVal Opcode As Variant)
    Image2.PaintPicture Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2, Opcode
    Image2.Picture = Image2.Image
    Image2.Height = Height2
    Image2.Width = Width2
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picBack,picBack,-1,PaintPicture
Public Sub PaintPicturePlay(ByVal Picture As Picture, ByVal X1 As Single, ByVal Y1 As Single, Optional ByVal Width1 As Variant, Optional ByVal Height1 As Variant, Optional ByVal X2 As Variant, Optional ByVal Y2 As Variant, Optional ByVal Width2 As Variant, Optional ByVal Height2 As Variant, Optional ByVal Opcode As Variant)
    Image1.PaintPicture Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2, Opcode
    Image1.Picture = Image1.Image
    Image1.Height = Height2
    Image1.Width = Width2
End Sub
