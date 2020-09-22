VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   4650
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   8295
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Tweak.checkbox checkbox1 
      Height          =   315
      Left            =   375
      TabIndex        =   3
      Top             =   4080
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   556
      LabelText       =   "Show at Start-Up"
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   3885
      Top             =   3960
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Click to Enter"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   6960
      TabIndex        =   4
      Top             =   60
      Width           =   2115
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   570
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Width           =   2685
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Warning"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   735
      TabIndex        =   2
      ToolTipText     =   "Warning Agreement"
      Top             =   2025
      Width           =   6840
   End
   Begin VB.Label LbNote 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSplash.frx":000C
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1125
      Left            =   750
      TabIndex        =   1
      ToolTipText     =   "Warning Agreement"
      Top             =   2355
      Width           =   6780
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00B99D7F&
      X1              =   1005
      X2              =   7440
      Y1              =   3540
      Y2              =   3540
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00B99D7F&
      X1              =   1005
      X2              =   7440
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2040
      Left            =   135
      Shape           =   4  'Rounded Rectangle
      Top             =   1740
      Width           =   7875
   End
   Begin VB.Image Image2 
      Height          =   390
      Left            =   4425
      Top             =   3870
      Width           =   540
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Programmed by Craig Phillips"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   5220
      TabIndex        =   0
      ToolTipText     =   "Author"
      Top             =   4185
      Width           =   3675
   End
   Begin VB.Image Image1 
      Height          =   1650
      Left            =   60
      Picture         =   "frmSplash.frx":00E9
      ToolTipText     =   "Tweak Machine Logo"
      Top             =   60
      Width           =   3750
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim j As Integer


Private Sub checkbox1_Click()
SetKeyValue HKEY_LOCAL_MACHINE, "Software\" & Software_Name & "\Main", "Splash", 0, REG_SZ
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Timer1.Enabled = False
 ChangeRes 1024, 768
    frmMain.Show 1
    Unload Me

End Sub




Private Sub Form_Load()
On Error GoTo hand
  Dim Add As Long
  Dim Sum As Long

  Dim X As Single
  Dim Y As Single
  Dim sp As String


    Hive_Key = HKEY_LOCAL_MACHINE
    Sub_Key = "Software\" & Software_Name & "\Main"
    Open_SubKey Hive_Key, Sub_Key
    Query_Value REG_SZ, "Splash"
    sp = S_Value
    If sp = 0 Then
    ChangeRes 1024, 768
    frmMain.Show 1
    Unload Me
    End If
    X = Me.Width / Screen.TwipsPerPixelX   'Registers the size of the
    Y = Me.Height / Screen.TwipsPerPixelY  'form in pixels
Me.AutoRedraw = True
    Sum = CreateRectRgn(5, 0, X - 5, 1)
    CombineRgn Sum, Sum, CreateRectRgn(3, 1, X - 3, 2), 2
    CombineRgn Sum, Sum, CreateRectRgn(2, 2, X - 2, 3), 2
    CombineRgn Sum, Sum, CreateRectRgn(1, 3, X - 1, 4), 2
    CombineRgn Sum, Sum, CreateRectRgn(1, 4, X - 1, 5), 2
    CombineRgn Sum, Sum, CreateRectRgn(0, 5, X, Y), 2
    SetWindowRgn hwnd, Sum, True   'Sets corners transparent
Me.Line (Me.Width - 15, 0)-(Me.Width - 15, Me.Height - 1), RGB(90, 190, 255)
Me.Line (0, 0)-(0, Me.Height - 15), RGB(90, 190, 255)
Me.Line (0, Me.Height - 15)-(Me.Width - 15, Me.Height - 15), RGB(90, 190, 255)
Me.Line (0, 0)-(Me.Width, 0), RGB(90, 190, 255)
Me.Line (Me.Width - 30, 0)-(Me.Width - 30, 75), RGB(90, 190, 255)
Me.Line (Me.Width - 45, 0)-(Me.Width - 45, 45), RGB(90, 190, 255)
Me.Line (Me.Width - 75, 15)-(Me.Width - 45, 15), RGB(90, 190, 255)
Me.Line (15, 15)-(75, 15), RGB(90, 190, 255)
Me.Line (15, 0)-(15, 75), RGB(90, 190, 255)
Me.Line (30, 0)-(30, 45), RGB(90, 190, 255)
checkbox1.Value = sp
Exit Sub
hand:
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = False
 ChangeRes 1024, 768
    frmMain.Show 1
    Unload Me
End Sub



Private Sub Image1_Click()
Timer1.Enabled = False
 ChangeRes 1024, 768
    frmMain.Show 1
    Unload Me
End Sub

Private Sub Label1_Click()
Timer1.Enabled = False
 ChangeRes 1024, 768
    frmMain.Show 1
    Unload Me
End Sub

Private Sub Label2_Click()
Timer1.Enabled = False
 ChangeRes 1024, 768
    frmMain.Show 1
    Unload Me
End Sub

Private Sub Label8_Click()
Timer1.Enabled = False
 ChangeRes 1024, 768
    frmMain.Show 1
    Unload Me
End Sub

Private Sub LbNote_Click()
Timer1.Enabled = False
 ChangeRes 1024, 768
    frmMain.Show 1
    Unload Me
End Sub

Private Sub Timer1_Timer()
j = j + 1
If j > 95 Then
j = 1
    
End If

Image2.Picture = frmSettings.IMG.ListImages(j).Picture

End Sub
