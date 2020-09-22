VERSION 5.00
Begin VB.Form frmMsg 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Tweak.XPButton XPButton1 
      Height          =   480
      Left            =   1995
      TabIndex        =   2
      Top             =   2655
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Ok"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   1035
      Picture         =   "frmMsg.frx":0000
      Top             =   1095
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00B99D7F&
      X1              =   345
      X2              =   5655
      Y1              =   735
      Y2              =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   225
      TabIndex        =   3
      ToolTipText     =   "Message"
      Top             =   1170
      Width           =   5700
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
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
      Left            =   255
      TabIndex        =   1
      ToolTipText     =   "Message Title"
      Top             =   840
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
      Left            =   1005
      TabIndex        =   0
      Top             =   150
      Width           =   3105
   End
   Begin VB.Image ImgPic 
      Height          =   600
      Left            =   345
      Picture         =   "frmMsg.frx":1302
      ToolTipText     =   "Message Icon"
      Top             =   90
      Width           =   600
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()

progInit = 1
makeTrans Me, hwnd, RGB(90, 190, 255)
PlaySound 102
End Sub



Private Sub ImgPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
drag Me
End Sub

Private Sub XPButton1_Click()

ImgPic.Picture = Image1.Picture
Unload Me
If Not No_Login = 4 Then
frmLogin.Show 1
DoEvents
End If

Exit Sub
hand:
MsgBox "Eror"
End Sub

