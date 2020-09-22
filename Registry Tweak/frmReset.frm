VERSION 5.00
Begin VB.Form frmReset 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Tweak.XPButton XPButton1 
      Height          =   420
      Left            =   1245
      TabIndex        =   3
      Top             =   2535
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   741
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
   Begin Tweak.XPButton XPButton2 
      Height          =   420
      Left            =   3015
      TabIndex        =   4
      Top             =   2535
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Later"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.Image ImgPic 
      Height          =   600
      Left            =   195
      Picture         =   "frmReset.frx":0000
      ToolTipText     =   "Message Icon"
      Top             =   150
      Width           =   600
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
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
      Left            =   855
      TabIndex        =   2
      Top             =   210
      Width           =   3105
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reset Computer"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   105
      TabIndex        =   1
      ToolTipText     =   "Message Title"
      Top             =   900
      Width           =   5700
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmReset.frx":1302
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   75
      TabIndex        =   0
      ToolTipText     =   "Message"
      Top             =   1230
      Width           =   5700
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00B99D7F&
      X1              =   195
      X2              =   5505
      Y1              =   795
      Y2              =   795
   End
End
Attribute VB_Name = "frmReset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()

ImgPic.Picture = frmSettings.IMG.ListImages(29).Picture
makeTrans Me, hwnd, RGB(90, 190, 255)
PlaySound 102
End Sub

Private Sub ImgPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
drag Me
End Sub

Private Sub XPButton1_Click()
ExitWindowsEx EWX_REBOOT, 0&
End Sub

Private Sub XPButton2_Click()
Unload Me
End Sub

