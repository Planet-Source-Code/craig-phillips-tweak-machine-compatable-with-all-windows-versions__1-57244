VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Tweak.XPButton XPButton1 
      Height          =   525
      Left            =   1515
      TabIndex        =   9
      Top             =   6660
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   926
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
   Begin Tweak.checkbox chkUser 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   556
      ForeColor       =   12632319
      LabelText       =   "Don't display the last Username logged on"
   End
   Begin Tweak.checkbox chkDisablePrinterAddition 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   2805
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   556
      ForeColor       =   16761024
      LabelText       =   "Disable the Addition of Printers"
      ColourScheme    =   2
   End
   Begin Tweak.checkbox chkRestrictPassword 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   4050
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   556
      ForeColor       =   0
      LabelText       =   "Restrict Access to the Passwords Applet"
      ColourScheme    =   4
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "These are examples of checkboxes and will not change any of the settings."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   570
      Left            =   210
      TabIndex        =   10
      Top             =   6000
      Width           =   4755
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "If an option requires for you to reset or log off your computer it will prompt you to do so."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   570
      Left            =   210
      TabIndex        =   8
      Top             =   5340
      Width           =   4755
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":0000
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   795
      Left            =   480
      TabIndex        =   6
      Top             =   4395
      Width           =   4740
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":0089
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   795
      Left            =   480
      TabIndex        =   4
      Top             =   3150
      Width           =   4740
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":011C
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   795
      Left            =   480
      TabIndex        =   3
      Top             =   1905
      Width           =   4740
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Indicators"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   105
      TabIndex        =   2
      Top             =   1080
      Width           =   4185
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00B99D7F&
      X1              =   120
      X2              =   5040
      Y1              =   930
      Y2              =   930
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
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
      Left            =   975
      TabIndex        =   0
      Top             =   270
      Width           =   3105
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()

Image1.Picture = frmSettings.IMG.ListImages(65).Picture

BackColor = RGB(127, 127, 127)
makeTrans Me, hwnd, RGB(255, 0, 255)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
drag Me
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
drag Me
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
drag Me
End Sub


Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
drag Me
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
drag Me
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
drag Me
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
drag Me
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
drag Me
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
drag Me
End Sub

Private Sub XPButton1_Click()
Unload Me
End Sub

