VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2370
   ClientLeft      =   1290
   ClientTop       =   3345
   ClientWidth     =   4065
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Tweak.XPButton vmdOk 
      Height          =   465
      Left            =   960
      TabIndex        =   3
      ToolTipText     =   "Click this button to save changes and exit"
      Top             =   1755
      Width           =   1320
      _ExtentX        =   2328
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
   Begin Tweak.TxtBox txtPassword 
      Height          =   345
      Left            =   960
      TabIndex        =   1
      ToolTipText     =   "Enter your password in here"
      Top             =   1170
      Width           =   2835
      _extentx        =   5001
      _extenty        =   609
      backcolor       =   16777215
      font            =   "frmLogin.frx":0CCE
      passwordchar    =   "*"
      text            =   ""
      value           =   -1  'True
      bordercolor     =   16744192
   End
   Begin Tweak.XPButton cmdExit 
      Height          =   465
      Left            =   2460
      TabIndex        =   4
      ToolTipText     =   "Click this button to exit without saving changes"
      Top             =   1755
      Width           =   1320
      _ExtentX        =   2328
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
      Caption         =   "&Exit"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00B99D7F&
      X1              =   150
      X2              =   3930
      Y1              =   825
      Y2              =   825
   End
   Begin VB.Label lblPassword 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password to Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   765
      TabIndex        =   2
      ToolTipText     =   "To login you must enter a password"
      Top             =   855
      Width           =   3120
   End
   Begin VB.Image imgPassword 
      Height          =   480
      Left            =   300
      Picture         =   "frmLogin.frx":0CF6
      Top             =   180
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
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
      TabIndex        =   0
      ToolTipText     =   "Login title"
      Top             =   240
      Width           =   3105
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub Form_Activate()
If blnNoPass = True Then
No_Login = 4
      
    frmSplash.Show 1
    Unload Me
End If
txtPassword.SetFocus
Select_On_Focus

End Sub

Private Sub Form_Load()

makeTrans Me, hwnd, RGB(90, 190, 255)

    Init_Password
   
End Sub


Private Sub cmdExit_Click()
End
End Sub


Private Sub Select_On_Focus()
    On Error Resume Next
    With ActiveControl
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub





Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
drag Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set cShadow = Nothing
End Sub



Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
drag Me
End Sub

Private Sub imgPassword_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
drag Me
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
drag Me
End Sub

Private Sub lblPassword_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
drag Me
End Sub

Private Sub txtPassword_GotFocus()
Select_On_Focus

End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
vmdOk_Click
End If

End Sub

Private Sub vmdOk_Click()
On Error Resume Next
    'check for correct password
    txtPassword.Text = UCase(txtPassword.Text)
    If txtPassword.Text = Retrieve_Password Then
    No_Login = 4

        Unload Me
        PlaySound 102
        frmSplash.Show
    Else
        No_Login = No_Login + 1
        If No_Login = 3 Then
        Unload Me
            Msgex "Sorry !! You Are Not Authorized to Continue... Exiting...", "Unauthorized Access"
            If demono = True Then
            SetKeyValue HKEY_LOCAL_MACHINE, "Software\" & Software_Name & "\Main", "Verification", "2345", REG_SZ
            End If
            Unload Me
            End
        Else
        Hide
            Msgex "Invalid Password, try again!", "Login", Error
            
        End If
    End If
   
End Sub
