VERSION 5.00
Begin VB.UserControl checkbox 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1755
      Top             =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   345
      TabIndex        =   0
      Top             =   0
      Width           =   1530
   End
   Begin VB.Image Image5 
      Height          =   315
      Left            =   0
      Picture         =   "checkbox.ctx":0000
      Top             =   0
      Width           =   315
   End
   Begin VB.Image Image4 
      Height          =   315
      Left            =   2340
      Picture         =   "checkbox.ctx":0582
      Top             =   2625
      Width           =   315
   End
   Begin VB.Image Image3 
      Height          =   315
      Left            =   1365
      Picture         =   "checkbox.ctx":0B04
      Top             =   2280
      Width           =   315
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   480
      Picture         =   "checkbox.ctx":1086
      Top             =   2580
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   405
      Picture         =   "checkbox.ctx":1608
      Top             =   2040
      Width           =   315
   End
End
Attribute VB_Name = "checkbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Event Declarations:
Event Click() 'MappingInfo=Image5,Image5,-1,Click
'Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
'Default Property Values:
Const m_def_ColourScheme = 5
Const m_def_Value = 0
'Property Variables:
Dim m_ColourScheme As Variant
Dim m_Value As Variant
Public Enum ole_sch
red = 1
blue = 2
green = 3
black = 4
custom = 5
End Enum
Public value2 As Integer







Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Value = 1 Then
Value = 3
Else
Value = 2
End If
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Value = 3 Then
Value = 1
Else
Value = 0
End If
End Sub

Private Sub Label1_Click()
If UserControl.Enabled = True Then
RaiseEvent Click
Clicked
End If
End Sub



Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Value = 1 Then
Value = 3
Else
Value = 2
End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Value = 3 Then
Value = 1
Else
Value = 0
End If
End Sub

Private Sub Timer1_Timer()

If Value = 0 Then
Image5.Picture = Image1.Picture
End If
If Value = 1 Then
Image5.Picture = Image4.Picture
End If
If Value = 2 Then
Image5.Picture = Image2.Picture
End If
If Value = 3 Then
Image5.Picture = Image3.Picture
End If

End Sub

Private Sub UserControl_Initialize()
UserControl.BackColor = RGB(127, 127, 127)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    If New_Enabled = False Then
    Label1.ForeColor = RGB(192, 192, 192)
    Else
    Label1.ForeColor = ForeColor
    End If
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub
'
'Private Sub UserControl_Click()
'    RaiseEvent Click
'End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    
    End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Label1.Caption = PropBag.ReadProperty("LabelText", "Label1")
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_ColourScheme = PropBag.ReadProperty("ColourScheme", m_def_ColourScheme)
End Sub

Private Sub UserControl_Resize()
UserControl.Height = Image5.Height
Label1.Width = UserControl.Width - (Image5.Left + Image5.Width + Label1.Left)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("LabelText", Label1.Caption, "Label1")
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("ColourScheme", m_ColourScheme, m_def_ColourScheme)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Caption
Public Property Get LabelText() As String
Attribute LabelText.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    LabelText = Label1.Caption
End Property

Public Property Let LabelText(ByVal New_LabelText As String)
    Label1.Caption() = New_LabelText
    PropertyChanged "LabelText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Value() As Variant
Attribute Value.VB_Description = "If it is checked or not.\r\n1 = checked\r\n2 = crossed"
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Variant)
    m_Value = New_Value
    PropertyChanged "Value"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Value = m_def_Value
    m_ColourScheme = m_def_ColourScheme
End Sub

Private Sub Image5_Click()
If UserControl.Enabled = True Then
    RaiseEvent Click
Clicked
End If
End Sub

Private Function Clicked()
If Value = 0 Then
Value = 1
Else
Value = 0
End If
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get ColourScheme() As ole_sch
Attribute ColourScheme.VB_Description = "Whether the control is a preset colour or not"
    ColourScheme = m_ColourScheme
End Property

Public Property Let ColourScheme(ByVal New_ColourScheme As ole_sch)
    m_ColourScheme = New_ColourScheme
    If New_ColourScheme = black Then
    Label1.ForeColor = vbBlack
    End If
        If New_ColourScheme = blue Then
    Label1.ForeColor = &HFFC0C0
    End If
    If New_ColourScheme = custom Then
    Label1.ForeColor = ForeColor
    End If
    If New_ColourScheme = green Then
    Label1.ForeColor = &HC0FFC0
    End If
    If New_ColourScheme = red Then
    Label1.ForeColor = &HC0C0FF
    End If
    PropertyChanged "ColourScheme"
End Property

