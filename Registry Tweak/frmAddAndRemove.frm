VERSION 5.00
Begin VB.Form frmAddAndRemove 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Tweak.XPButton cmdDelete 
      Height          =   405
      Left            =   1665
      TabIndex        =   4
      Top             =   5655
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Erase"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.ListBox lstAddRemove 
      Height          =   2595
      Left            =   195
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2295
      Width           =   5220
   End
   Begin Tweak.XPButton XPButton1 
      Height          =   405
      Left            =   2850
      TabIndex        =   5
      Top             =   5655
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "E&xit"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00B99D7F&
      X1              =   225
      X2              =   5400
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAddAndRemove.frx":0000
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   225
      TabIndex        =   3
      Top             =   930
      Width           =   5190
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To remove a program from the list you can simply select the program and click Erase."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   180
      TabIndex        =   2
      Top             =   5010
      Width           =   5235
   End
   Begin VB.Image Image1 
      Height          =   465
      Left            =   240
      Top             =   180
      Width           =   510
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Add and Remove Programs"
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
      Left            =   915
      TabIndex        =   0
      Top             =   255
      Width           =   4950
   End
End
Attribute VB_Name = "frmAddAndRemove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdDelete_Click()
    If Trim(lstAddRemove.List(lstAddRemove.ListIndex)) <> "" Then
        User_Res = MsgBox("Are you sure you want to Delete the Selected Key " & _
                        "from Add/Remove ?", vbYesNo + vbDefaultButton2 + vbCritical, _
                        "Confirm Delete from Add/Remove")
        If User_Res = vbYes Then
            If Delete_AddRemove = True Then
                Msgex "Key Deleted from Add/Remove !!", "Deletion Successfull", Info
            Else
                Msgex "Cannot Delete Selected Key from Add/Remove !!", "Deletion Failed", Error
            End If
        End If
        Read_AddRemove_Programs
    End If
    cmdDelete.Enabled = False

End Sub

Private Sub Form_Load()

makeTrans Me, hwnd, RGB(90, 190, 255)
Image1.Picture = frmSettings.IMG.ListImages(17).Picture
Read_AddRemove_Programs
End Sub
Private Function Delete_AddRemove() As Boolean
On Error GoTo ErrHand
    
    Dim lKeyNum As Long, lIndex As Long
    Dim sKeyName As String, lKeyNameLen As Long
    Dim DontAdd As Boolean
    Dim Soft_Name As String
    
    Hive_Key = HKEY_LOCAL_MACHINE
    Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Uninstall"
    
    Re_Init_Val
    
    Open_SubKey Hive_Key, Sub_Key 'Opening Uninstall Subkey
   
    'Retrieving the number of Subkeys Present in Uninstall
    Call RegQueryInfoKey(HRegKey, vbNullString, 0&, 0&, lKeyNum, 0&, 0&, 0&, 0&, 0&, 0&, tFT)
           
    
    If lKeyNum > 0 Then 'If there are any Softwares Installed
        For lIndex = 0 To lKeyNum - 1 'Loop Until End of Sub Keys
        
            sKeyName = String$(255, 0) 'Declare a 255 length string
            lKeyNameLen = 255
            
            'Retrieve the Sub Key Name
            Call RegEnumKeyEx(HRegKey, lIndex, sKeyName, lKeyNameLen, 0&, vbNullString, 0&, tFT)
                                           
            sKeyName = Left$(sKeyName, lKeyNameLen) 'Reduce String Length
            
            'Opening the Specified Key
            RegOpenKey Hive_Key, Sub_Key & "\" & sKeyName, HRegKey2 'Different Handle to key
            lngBuffer = 0
            
            'Check if there are 'DisplayName' & 'UninstallString' Values in the Key
            If RegQueryValueEx(HRegKey2, "DisplayName", 0&, REG_SZ, ByVal 0&, lngBuffer) = ERROR_SUCCESS Then
                lngBuffer = 256
                Soft_Name = Space$(lngBuffer)
                
                'Retrieve the Software Name ('DisplayName')
                RetVal = RegQueryValueEx(HRegKey2, "DisplayName", 0&, REG_SZ, ByVal Soft_Name, lngBuffer)
                Soft_Name = Left(Soft_Name, lngBuffer - 1) 'drop null-terminator
                
                'If DisplayName = List
                If Soft_Name = lstAddRemove.List(lstAddRemove.ListIndex) Then
                    RegCloseKey HRegKey2
                    Delete_Key (sKeyName) ' Delete Sub Key
                    Delete_AddRemove = True
                    GoTo END_Delete
                End If
            End If
                                                
            RegCloseKey HRegKey2 'Close Registry Handle
            
        Next lIndex
    End If

END_Delete:
    
    RegCloseKey HRegKey
    ''
    
Exit Function
ErrHand:
    MsgBox "Error Occurred while Deleting Selected Uninstall String !!" & vbCrLf & "Function : Delete_AddRemove " & vbCrLf & Err.Description, vbCritical
End Function


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

Private Sub XPButton1_Click()
Unload Me
End Sub
