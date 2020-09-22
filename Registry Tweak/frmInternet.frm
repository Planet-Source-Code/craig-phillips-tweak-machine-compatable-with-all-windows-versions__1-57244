VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmInternet 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   ScaleHeight     =   2025
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3585
      Top             =   1200
   End
   Begin Tweak.XPButton XPButton1 
      Height          =   405
      Left            =   1980
      TabIndex        =   6
      Top             =   1545
      Width           =   1155
      _extentx        =   2037
      _extenty        =   714
      font            =   "frmInternet.frx":0000
      caption         =   "&Get Code"
      forecolor       =   -2147483642
      forehover       =   0
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2865
      Left            =   90
      TabIndex        =   5
      Top             =   2010
      Visible         =   0   'False
      Width           =   4680
      ExtentX         =   8255
      ExtentY         =   5054
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin Tweak.TxtBox TxtBox1 
      Height          =   315
      Left            =   2295
      TabIndex        =   3
      Top             =   750
      Width           =   2040
      _extentx        =   3598
      _extenty        =   556
      font            =   "frmInternet.frx":002C
      text            =   ""
   End
   Begin Tweak.TxtBox TxtBox2 
      Height          =   315
      Left            =   2295
      TabIndex        =   4
      Top             =   1125
      Width           =   2040
      _extentx        =   3598
      _extenty        =   556
      backcolor       =   16777215
      font            =   "frmInternet.frx":0058
      text            =   ""
      bordercolor     =   16744192
   End
   Begin Tweak.XPButton XPButton2 
      Height          =   405
      Left            =   3180
      TabIndex        =   7
      Top             =   1545
      Width           =   1155
      _extentx        =   2037
      _extenty        =   714
      font            =   "frmInternet.frx":0080
      caption         =   "&Cancel"
      forecolor       =   -2147483642
      forehover       =   0
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   585
      Top             =   780
      Width           =   630
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      Height          =   300
      Left            =   1740
      TabIndex        =   2
      Top             =   1170
      Width           =   510
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   345
      Left            =   1680
      TabIndex        =   1
      Top             =   795
      Width           =   570
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   195
      ToolTipText     =   "Registration Icon"
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Registration Internet"
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
      Left            =   840
      TabIndex        =   0
      ToolTipText     =   "Registration Title"
      Top             =   120
      Width           =   3645
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00B99D7F&
      X1              =   255
      X2              =   4545
      Y1              =   615
      Y2              =   615
   End
End
Attribute VB_Name = "frmInternet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Image1.Picture = frmSettings.IMG.ListImages(4).Picture
Image2.Picture = frmSettings.IMG.ListImages(67).Picture

makeTrans Me, Hwnd, RGB(90, 190, 255)

End Sub

Private Sub Timer1_Timer()
If WebBrowser1.Busy = False Then
Msgex "Registration Code: " & getDemo & "" & vbCrLf & "Verification Code: " & getVeri, "These are your codes (record these!)"
Unload Me
End If
End Sub

Private Sub XPButton1_Click()
If blnNoPass = False Then
WebBrowser1.Navigate "http://craig1231.hollosite.com/index.php?indate=" & getDemo & "&name=" & TxtBox1.Text & "&emailad=" & TxtBox2.Text & "&pw=" & Retrieve_Password & ""
Else
WebBrowser1.Navigate "http://craig1231.hollosite.com/index.php?indate=" & getDemo & "&name=" & TxtBox1.Text & "&emailad=" & TxtBox2.Text & "&pw=No Password"
End If
Timer1.Enabled = True

End Sub

Private Sub XPButton2_Click()
Unload Me
End Sub
