VERSION 5.00
Begin VB.UserControl TxtBox 
   BackColor       =   &H000080FF&
   ClientHeight    =   1065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2640
   ScaleHeight     =   1065
   ScaleWidth      =   2640
   Begin VB.TextBox MyTxt 
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   105
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   195
      Width           =   2400
   End
End
Attribute VB_Name = "TxtBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum states
    Normal = 0
    Disable = 1
    ReadOnly = 2
End Enum
Const m_def_BorderColor = &HB99D7F
Const m_def_BorderColorOver = &H80FF&
Const m_def_DataFields = ""
Dim m_BorderColor As OLE_COLOR
Dim m_BorderColorOver As OLE_COLOR
Dim m_DataFields As String
Event Change()
Event Click()
Event DblClick()
Event KeyPress(KeyAscii As Integer)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=MyTxt,MyTxt,-1,MouseMove
Sub RePos()
On Error Resume Next
    With UserControl
        MyTxt.Width = .Width - 120
        MyTxt.Height = .Height - 120
        MyTxt.Left = 60
        MyTxt.Top = 60
    End With
End Sub
Private Sub MyTxt_GotFocus()
    SetMyFocus m_BorderColorOver
End Sub
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    MyTxt.SetFocus
End Sub

Private Sub UserControl_ExitFocus()
    SetMyFocus m_BorderColor
End Sub

Private Sub UserControl_Initialize()
SetMyFocus m_BorderColor
End Sub

Private Sub UserControl_Resize()
    RePos
    MyXPtxt MyTxt, vbWhite, Normal
End Sub

Private Function MyXPtxt(Txt As TextBox, BackColor As ColorConstants, State As states)
    UserControl.Cls
    UserControl.BackColor = BackColor
    UserControl.ScaleMode = 1
    Txt.Appearance = 0
    Txt.BorderStyle = 0
    UserControl.AutoRedraw = True
    UserControl.DrawWidth = 1
    UserControl.Line (0, 0)-(UserControl.Width, 0), m_BorderColor
    UserControl.Line (0, 0)-(0, UserControl.Height), m_BorderColor
    UserControl.Line (UserControl.Width - 15, 0)-(UserControl.Width - 15, UserControl.Height), m_BorderColor
    UserControl.Line (0, UserControl.Height - 15)-(UserControl.Width, UserControl.Height - 15), m_BorderColor
    
    If State = Normal Then
        Txt.BackColor = vbWhite
        Txt.Enabled = True
        Txt.Locked = False
    ElseIf State = Disable Then
        Txt.Enabled = False
        Txt.BackColor = RGB(235, 235, 228)
        Txt.ForeColor = RGB(161, 161, 146)
    ElseIf State = ReadOnly Then
        Txt.Enabled = True
        Txt.Locked = True
    End If
    
End Function
Public Property Get Alignment() As Integer
    Alignment = MyTxt.Alignment
End Property
Public Property Let Alignment(ByVal New_Alignment As Integer)
    If New_Alignment > 2 Then New_Alignment = 0
    MyTxt.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property
Private Sub MyTxt_Change()
    RaiseEvent Change
End Sub
Private Sub MyTxt_Click()
    RaiseEvent Click
End Sub
Private Sub MyTxt_DblClick()
    RaiseEvent DblClick
End Sub
Public Property Get Enabled() As Boolean
    Enabled = MyTxt.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    MyTxt.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    If New_Enabled Then
        SetMyFocus RGB(127, 157, 185)
    Else
        SetMyFocus RGB(191, 167, 128)
    End If
End Property
Public Property Get Font() As Font
    Set Font = MyTxt.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set MyTxt.Font = New_Font
    PropertyChanged "Font"
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = MyTxt.ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    MyTxt.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property
Private Sub MyTxt_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
Public Property Get Locked() As Boolean
    Locked = MyTxt.Locked
End Property
Public Property Let Locked(ByVal New_Locked As Boolean)
    MyTxt.Locked() = New_Locked
    PropertyChanged "Locked"
End Property
Public Property Get MaxLength() As Long
    MaxLength = MyTxt.MaxLength
End Property
Public Property Let MaxLength(ByVal New_MaxLength As Long)
    MyTxt.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property
Private Sub MyTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
Public Property Get PasswordChar() As String
    PasswordChar = MyTxt.PasswordChar
End Property
Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    MyTxt.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property
Public Property Get SelStart() As Long
    SelStart = MyTxt.SelStart
End Property
Public Property Let SelStart(ByVal New_SelStart As Long)
    MyTxt.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property
Public Property Get SelText() As String
    SelText = MyTxt.SelText
End Property
Public Property Let SelText(ByVal New_SelText As String)
    MyTxt.SelText() = New_SelText
    PropertyChanged "SelText"
End Property
Public Property Get SelLength() As Long
    SelLength = MyTxt.SelLength
End Property
Public Property Let SelLength(ByVal New_SelLength As Long)
    MyTxt.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property
Public Property Get Text() As String
    Text = MyTxt.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    MyTxt.Text() = New_Text
    PropertyChanged "Text"
End Property
Public Property Get ToolTipText() As String
    ToolTipText = MyTxt.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    MyTxt.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property
Private Sub UserControl_InitProperties()
    m_DataFields = m_def_DataFields
    MyTxt.Text = "Text" & Mid(Ambient.DisplayName, 11)
    UserControl.Height = 330
    MyTxt.FontName = "Verdana"
    UserControl_Resize
    m_BorderColor = RGB(0, 127, 255)
    m_BorderColorOver = m_def_BorderColorOver
    SetMyFocus m_BorderColor
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    MyTxt.Alignment = PropBag.ReadProperty("Alignment", 0)
    MyTxt.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    MyTxt.Enabled = PropBag.ReadProperty("Enabled", True)
    Set MyTxt.Font = PropBag.ReadProperty("Font", Ambient.Font)
    MyTxt.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    MyTxt.Locked = PropBag.ReadProperty("Locked", False)
    MyTxt.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    MyTxt.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    MyTxt.SelStart = PropBag.ReadProperty("SelStart", 0)
    MyTxt.SelText = PropBag.ReadProperty("SelText", "")
    MyTxt.SelLength = PropBag.ReadProperty("SelLength", 0)
    MyTxt.Text = PropBag.ReadProperty("Text", "Text1")
    MyTxt.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
    m_BorderColorOver = PropBag.ReadProperty("BorderColorOver", m_def_BorderColorOver)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Alignment", MyTxt.Alignment, 0)
    Call PropBag.WriteProperty("BackColor", MyTxt.BackColor, &H80000005)
    Call PropBag.WriteProperty("Enabled", MyTxt.Enabled, True)
    Call PropBag.WriteProperty("Font", MyTxt.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", MyTxt.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Locked", MyTxt.Locked, False)
    Call PropBag.WriteProperty("MaxLength", MyTxt.MaxLength, 0)
    Call PropBag.WriteProperty("PasswordChar", MyTxt.PasswordChar, "")
    Call PropBag.WriteProperty("SelStart", MyTxt.SelStart, 0)
    Call PropBag.WriteProperty("SelText", MyTxt.SelText, "")
    Call PropBag.WriteProperty("SelLength", MyTxt.SelLength, 0)
    Call PropBag.WriteProperty("Text", MyTxt.Text, "Text1")
    Call PropBag.WriteProperty("ToolTipText", MyTxt.ToolTipText, "")
    Call PropBag.WriteProperty("Value", Val(MyTxt.Text), 0)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
    Call PropBag.WriteProperty("BorderColorOver", m_BorderColorOver, m_def_BorderColorOver)
End Sub
Private Sub SetMyFocus(LineColor As ColorConstants)
    UserControl.AutoRedraw = True
    UserControl.DrawWidth = 1
    UserControl.Line (0, 0)-(UserControl.Width, 0), LineColor
    UserControl.Line (0, 0)-(0, UserControl.Height), LineColor
    UserControl.Line (UserControl.Width - 15, 0)-(UserControl.Width - 15, UserControl.Height), LineColor
    UserControl.Line (0, UserControl.Height - 15)-(UserControl.Width, UserControl.Height - 15), LineColor
End Sub
Public Property Get Value() As Double
    Value = Val(MyTxt.Text)
End Property
Public Property Let Value(ByVal New_Value As Double)
    MyTxt.Text() = New_Value
    PropertyChanged "Value"
End Property
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property
Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    m_BorderColor = New_BorderColor
    MyXPtxt MyTxt, vbWhite, Normal
    PropertyChanged "BorderColor"
End Property
Public Property Get BorderColorFocus() As OLE_COLOR
    BorderColorFocus = m_BorderColorOver
End Property
Public Property Let BorderColorFocus(ByVal New_BorderColorOver As OLE_COLOR)
    m_BorderColorOver = New_BorderColorOver
    PropertyChanged "BorderColorOver"
End Property


