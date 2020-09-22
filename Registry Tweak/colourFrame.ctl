VERSION 5.00
Begin VB.UserControl colourFrame 
   BackColor       =   &H8000000B&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   PropertyPages   =   "colourFrame.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H000080FF&
      Height          =   585
      Left            =   240
      TabIndex        =   0
      Top             =   45
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      FillColor       =   &H0080FFFF&
      Height          =   2400
      Left            =   15
      Shape           =   4  'Rounded Rectangle
      Top             =   15
      Width           =   2895
   End
End
Attribute VB_Name = "colourFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Default Property Values:
Const m_def_BackStyle = 0
Const m_def_LabelText = "Your Text Here"
'Property Variables:
Dim m_BackStyle As Integer

'Event Declarations:
'Event GetColour()
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."



Private Sub UserControl_Resize()
On Error GoTo gh
Shape1.Width = UserControl.Width - 50
Shape1.Height = UserControl.Height - 50
Label1.Top = UserControl.Height / 40
Label1.Left = UserControl.Width / 20
Label1.Width = (UserControl.Width - 50) - (UserControl.Width / 20)
Label1.Height = 1080
Exit Sub
gh:
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background colour used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape1,Shape1,-1,BorderColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the colour of a shapes border."
    ForeColor = Shape1.BorderColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Shape1.BorderColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape1,Shape1,-1,BorderColor
Public Property Get LabelColour() As OLE_COLOR
Attribute LabelColour.VB_Description = "Returns/sets the labels colour."
    LabelColour = Label1.ForeColor
End Property

Public Property Let LabelColour(ByVal New_LabelColour As OLE_COLOR)
    Label1.ForeColor = New_LabelColour
    PropertyChanged "LabelColour"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "LabelText"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Label1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
     
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,Your Text Here


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_LabelText = m_def_LabelText
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000008)
    Shape1.BorderColor = PropBag.ReadProperty("ForeColor", 65535)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Shape1.FillColor = PropBag.ReadProperty("FillColour", &H80FFFF)
    Shape1.FillStyle = PropBag.ReadProperty("FillStyle", 1)
    Label1.Caption = PropBag.ReadProperty("LabelText", "Label1")
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000008)
    Call PropBag.WriteProperty("ForeColor", Shape1.BorderColor, 65535)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)

    Call PropBag.WriteProperty("FillColour", Shape1.FillColor, &H80FFFF)
    Call PropBag.WriteProperty("FillStyle", Shape1.FillStyle, 1)
    Call PropBag.WriteProperty("LabelText", Label1.Caption, "Label1")
End Sub



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape1,Shape1,-1,FillColor
Public Property Get FillColour() As OLE_COLOR
Attribute FillColour.VB_Description = "Returns/sets the color used to fill in shapes, circles, and boxes."
    FillColour = Shape1.FillColor
End Property

Public Property Let FillColour(ByVal New_FillColour As OLE_COLOR)
    Shape1.FillColor() = New_FillColour
    PropertyChanged "FillColour"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape1,Shape1,-1,FillStyle
Public Property Get FillStyle() As Integer
Attribute FillStyle.VB_Description = "Returns/sets the fill style of a shape."
Attribute FillStyle.VB_ProcData.VB_Invoke_Property = "ltext"
    FillStyle = Shape1.FillStyle
End Property

Public Property Let FillStyle(ByVal New_FillStyle As Integer)
    Shape1.FillStyle() = New_FillStyle
    PropertyChanged "FillStyle"
End Property

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

