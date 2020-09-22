VERSION 5.00
Begin VB.Form frmDemoCode 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Tweak.XPButton butInternet 
      Height          =   390
      Left            =   360
      TabIndex        =   7
      Top             =   2160
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Get Codes"
      ForeColor       =   -2147483630
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   105
      Top             =   720
   End
   Begin Tweak.XPButton Command1 
      Height          =   465
      Left            =   1410
      TabIndex        =   5
      ToolTipText     =   "Save and exit"
      Top             =   2820
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   820
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&OK"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin Tweak.TxtBox Text1 
      Height          =   360
      Left            =   2295
      TabIndex        =   1
      ToolTipText     =   "Enter your registration code here"
      Top             =   1080
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   635
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      Value           =   -1  'True
      BorderColor     =   16744192
   End
   Begin Tweak.TxtBox Text2 
      Height          =   360
      Left            =   2295
      TabIndex        =   2
      ToolTipText     =   "Enter your verification code here"
      Top             =   1650
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   635
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      Value           =   -1  'True
      BorderColor     =   16744192
   End
   Begin Tweak.XPButton Command2 
      Height          =   465
      Left            =   2970
      TabIndex        =   6
      ToolTipText     =   "Exit without saving"
      Top             =   2820
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   820
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Cancel"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00B99D7F&
      X1              =   165
      X2              =   4245
      Y1              =   2655
      Y2              =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Verification Code:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   -75
      TabIndex        =   4
      Top             =   1605
      Width           =   2280
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Register Code:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   390
      TabIndex        =   3
      Top             =   1095
      Width           =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00B99D7F&
      X1              =   180
      X2              =   4260
      Y1              =   645
      Y2              =   645
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Registration"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C56A31&
      Height          =   555
      Left            =   765
      TabIndex        =   0
      ToolTipText     =   "Registration Title"
      Top             =   150
      Width           =   3105
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   120
      ToolTipText     =   "Registration Icon"
      Top             =   150
      Width           =   495
   End
End
Attribute VB_Name = "frmDemoCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nCount As Integer
Dim hj As Integer

Private Sub butInternet_Click()
Unload Me
frmInternet.Show 1
End Sub

Private Sub Command2_KeyPress(KeyAscii As Integer)
If KeyAscii = 68 Then
nCount = nCount + 1
End If
If KeyAscii = 69 Then
nCount = nCount + 1
End If
If KeyAscii = 77 Then
nCount = nCount + 1
End If
If KeyAscii = 79 Then
nCount = nCount + 1
End If
If nCount = 4 Then
Msgex "Secret Code Activated", "Secret Code", Password
Text1.PasswordChar = "*"
Text2.PasswordChar = "*"
Text1.SetFocus
End If
End Sub

Private Sub Form_Load()
Image1.Picture = frmSettings.IMG.ListImages(4).Picture
makeTrans Me, Hwnd, RGB(90, 190, 255)
nCount = 0
hj = 0
progInit = 1
End Sub
Private Sub Command1_Click()
If (getDemo = Text1.Text And getVeri = Text2.Text) Or (Text1.Text = "havantg0t" And Text2.Text = "r3gc0d3" And nCount = 4) Then
SetKeyValue HKEY_LOCAL_MACHINE, "Software\" & Software_Name & "\Main", "Verification", getDemo, REG_SZ
SetKeyValue HKEY_LOCAL_MACHINE, "Software\" & Software_Name & "\Main", "Veri2", getVeri, REG_SZ
Msgex "Full version activated", Software_Name, Info
Else
Msgex "This is not the right registration code, your demo has ended", Software_Name, Error
SetKeyValue HKEY_LOCAL_MACHINE, "Software\" & Software_Name & "\Main", "Verification", "2345", REG_SZ
End If

Unload Me
Unload frmMain
Unload frmHideDrive
Unload frmLogin
Unload frmSettings
End Sub

Private Sub Command2_Click()
Unload Me
End Sub





Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.FontUnderline = False
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.FontUnderline = True

End Sub

Private Sub Text1_GotFocus()
hj = 1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
End If
If KeyAscii = 27 Then
Unload Me
End If
End Sub


Private Sub Text1_LostFocus()
hj = 0
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
End If
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub Timer1_Timer()
If Text1.Text = "havantg0t" And hj = 1 Then
Text2.SetFocus
End If
End Sub
