VERSION 5.00
Begin VB.Form frmHideDrive 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstDrive 
      Height          =   2535
      Left            =   1575
      Style           =   1  'Checkbox
      TabIndex        =   5
      ToolTipText     =   "This is where the drives are displayed"
      Top             =   1575
      Width           =   2775
   End
   Begin Tweak.XPButton cmdOK 
      Height          =   525
      Left            =   510
      TabIndex        =   1
      ToolTipText     =   "Exit and Save"
      Top             =   4260
      Width           =   1650
      _ExtentX        =   2910
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
      Caption         =   "&Ok"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin Tweak.XPButton cmdCancel 
      Height          =   525
      Left            =   2220
      TabIndex        =   2
      ToolTipText     =   "Exit without saving"
      Top             =   4260
      Width           =   1650
      _ExtentX        =   2910
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
      Caption         =   "Ca&ncel"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin Tweak.XPButton cmdApply 
      Height          =   525
      Left            =   3930
      TabIndex        =   3
      ToolTipText     =   "Save changes"
      Top             =   4260
      Width           =   1650
      _ExtentX        =   2910
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
      Caption         =   "&Apply"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00B99D7F&
      X1              =   435
      X2              =   5445
      Y1              =   765
      Y2              =   765
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Remove the check-mark from a drive to prevent the drive from being displayed in My Computer."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   780
      TabIndex        =   4
      ToolTipText     =   "To hide a drive from My Computer, uncheck the check mark in the box next to it."
      Top             =   855
      Width           =   4320
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   240
      Top             =   180
      Width           =   1080
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hide Drives"
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
      Height          =   645
      Left            =   750
      TabIndex        =   0
      ToolTipText     =   "Hide Drive Title"
      Top             =   180
      Width           =   2205
   End
End
Attribute VB_Name = "frmHideDrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Apply_Hide_Drives()
On Error GoTo ErrHand

    Dim lTot_Hide_Number As Long
    Dim i As Integer
    lTot_Hide_Number = 0
    
    For i = 0 To lstDrive.ListCount - 1
        If lstDrive.Selected(i) = False Then
            lTot_Hide_Number = lTot_Hide_Number + _
             get_Hide_Drive_Number(Left(lstDrive.List(i), 1))
        End If
    Next
    
    Hive_Key = HKEY_CURRENT_USER
    Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    Re_Init_Val
    
    If lTot_Hide_Number <> 0 Then
        Open_SubKey Hive_Key, Sub_Key
        Create_Value REG_DWORD, "NoDrives", lTot_Hide_Number
    Else
        Open_SubKey Hive_Key, Sub_Key
        Delete_Value "NoDrives"
    End If
    
    
Exit Sub
ErrHand:
    MsgBox "Error Occurred while Hiding Drives !!" & vbCrLf & "Procedure : Apply_Hide_Drives " & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub Calc_Hidden_Drives()
On Error GoTo ErrHand
Dim lLastNum As Long
Dim sHide_Drives(26) As String
Dim iNo_HDrives As Integer

Re_Init_Val
Hive_Key = HKEY_CURRENT_USER
Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"

Open_SubKey Hive_Key, Sub_Key
Query_Value REG_DWORD, "NoDrives"
lTotHiddenNumber = BD_Value

lLastNum = BD_Value
iNo_HDrives = 0

For i = 65 To 90
    
    If get_Hide_Drive_Number(Chr(i)) = lTotHiddenNumber Then
        sHide_Drives(iNo_HDrives) = Chr(i)
        sHide_Drives(iNo_HDrives) = Left(sHide_Drives(iNo_HDrives), 1)
        iNo_HDrives = iNo_HDrives + 1
        Exit For
    ElseIf get_Hide_Drive_Number(Chr(i)) > lTotHiddenNumber And lLastNum < lTotHiddenNumber Then
        sHide_Drives(iNo_HDrives) = Chr(i - 1)
        lTotHiddenNumber = lTotHiddenNumber - get_Hide_Drive_Number(Chr(i - 1))
        sHide_Drives(iNo_HDrives) = Left(sHide_Drives(iNo_HDrives), 1)
        iNo_HDrives = iNo_HDrives + 1
        i = 64
    End If
    lLastNum = get_Hide_Drive_Number(Chr(i))
    
Next

For i = 0 To iNo_HDrives - 1
    For Ctr = 0 To lstDrive.ListCount - 1
        If Left(lstDrive.List(Ctr), 1) = sHide_Drives(i) Then
            lstDrive.Selected(Ctr) = False
            Exit For
        End If
    Next
Next

Exit Sub
ErrHand:
    MsgBox "Error Occurred while getting Hidden Drives !!" & vbCrLf & "Procedure : Calc_Hidden_Drives " & vbCrLf & Err.Description, vbCritical
End Sub

Private Function get_Hide_Drive_Number(sDriveLetter As String) As Long
On Error GoTo ErrHand

    Dim lDriveLetter As Integer
    Dim lDriveNumber As Long
    
    lDriveLetter = Asc(sDriveLetter) - 64
    
    lDriveNumber = 1
    For Ctr = 1 To lDriveLetter - 1
        lDriveNumber = lDriveNumber + lDriveNumber
    Next
    
    get_Hide_Drive_Number = lDriveNumber

Exit Function
ErrHand:
    MsgBox "Error Occurred while Getting Drive Hide Number !!" & vbCrLf & "Function : get_Hide_Drive_Number " & vbCrLf & Err.Description, vbCritical
End Function


Private Sub cmdApply_Click()
Apply_Hide_Drives
Get_All_Drives
cmdApply.Enabled = False
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Get_All_Drives()
On Error GoTo ErrHand

    Dim sDriveName As String
    sDriveName = String$(1, 0) 'Declare a Single length string
    
    sDriveName = "A"
    lstDrive.Clear
    For i = 0 To 25
        Select Case GetDriveType(sDriveName & ":\")
            'Case DRIVE_UNKNOWN

            'Case DRIVE_DOES_NOT_EXIST
                
            Case DRIVE_REMOVABLE
                lstDrive.AddItem (sDriveName) & vbTab & "-" & vbTab & "Floppy Drive"
            Case DRIVE_FIXED
                lstDrive.AddItem (sDriveName) & vbTab & "-" & vbTab & "Hard Disk Drive"
            'Case DRIVE_REMOTE
           
            Case DRIVE_CDROM
                lstDrive.AddItem (sDriveName) & vbTab & "-" & vbTab & "CD-ROM Drive"
                
            'Case DRIVE_RAMDISK
            
        End Select
        lstDrive.Selected(lstDrive.ListCount - 1) = True
        sDriveName = Chr(Asc(sDriveName) + 1)
    Next
    
    Calc_Hidden_Drives
    lstDrive.ListIndex = 0
Exit Sub
ErrHand:
    MsgBox "Error Occurred while Retrieving Drive Details !!" & vbCrLf & "Procedure : Get_All_Drives " & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub cmdOK_Click()
cmdApply_Click
Unload Me
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
drag Me
End Sub

Private Sub lstDrive_Click()
cmdApply.Enabled = True
logo = 1
End Sub


Private Sub Form_Load()

progInit = 1
makeTrans Me, hwnd, RGB(90, 190, 255)
Get_All_Drives
cmdApply.Enabled = False
Image1.Picture = frmSettings.IMG.ListImages(27).Picture
End Sub

Private Sub lstDrive_GotFocus()
logo = 1
End Sub
