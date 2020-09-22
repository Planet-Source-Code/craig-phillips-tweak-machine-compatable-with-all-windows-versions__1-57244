VERSION 5.00
Begin VB.Form frmLogoff 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Tweak.XPButton XPButton1 
      Height          =   420
      Left            =   1260
      TabIndex        =   0
      Top             =   2550
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
      Left            =   3030
      TabIndex        =   1
      Top             =   2550
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
   Begin VB.Line Line1 
      BorderColor     =   &H00B99D7F&
      X1              =   210
      X2              =   5520
      Y1              =   810
      Y2              =   810
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmLogoff.frx":0000
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
      Left            =   90
      TabIndex        =   4
      ToolTipText     =   "Message"
      Top             =   1245
      Width           =   5700
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Logoff Computer"
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
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Message Title"
      Top             =   915
      Width           =   5700
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
      Left            =   870
      TabIndex        =   2
      Top             =   225
      Width           =   3105
   End
   Begin VB.Image ImgPic 
      Height          =   600
      Left            =   210
      Picture         =   "frmLogoff.frx":009C
      ToolTipText     =   "Message Icon"
      Top             =   165
      Width           =   600
   End
End
Attribute VB_Name = "frmLogoff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

ImgPic.Picture = frmSettings.IMG.ListImages(29).Picture
makeTrans Me, hwnd, RGB(90, 190, 255)
PlaySound 102
End Sub

Private Sub XPButton1_Click()
ExitWindowsEx EWX_LOGOFF, 0&
End Sub

Private Sub XPButton2_Click()
Unload Me
End Sub

