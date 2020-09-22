VERSION 5.00
Begin VB.Form frmHideControlPanelApplet 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstHideApplet 
      Height          =   2535
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   1005
      Width           =   5055
   End
   Begin Tweak.XPButton cmdCancel 
      Height          =   450
      Left            =   2205
      TabIndex        =   2
      Top             =   3660
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   794
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
   Begin Tweak.XPButton cmdApply 
      Height          =   450
      Left            =   3705
      TabIndex        =   3
      Top             =   3660
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   794
      Enabled         =   0   'False
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
   Begin Tweak.XPButton cmdOK 
      Height          =   450
      Left            =   705
      TabIndex        =   4
      Top             =   3660
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   794
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
   Begin VB.Line Line1 
      BorderColor     =   &H00B99D7F&
      X1              =   165
      X2              =   4590
      Y1              =   690
      Y2              =   690
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Control Panel Applets"
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
      Left            =   630
      TabIndex        =   1
      Top             =   135
      Width           =   3780
   End
   Begin VB.Image Image1 
      Height          =   465
      Left            =   180
      Top             =   120
      Width           =   510
   End
End
Attribute VB_Name = "frmHideControlPanelApplet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()

makeTrans Me, hwnd, RGB(90, 190, 255)
Image1.Picture = frmSettings.IMG.ListImages(11).Picture
    Read_ControlPanel_Applets
    progInit = 1
End Sub
Private Sub cmdApply_Click()
Write_ControlPanel_Applets
Read_ControlPanel_Applets
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Read_ControlPanel_Applets()
On Error GoTo ErrHand
        
    Dim blnFound As Boolean
    Dim LstCtr As Integer
    Dim Dont_Load As Integer
    Dim ESCEQ As Long
    Dim blnDont_Load As Boolean
    
    Dont_Load = 0
    LstCtr = 0
    
    lstHideApplet.Clear
    
    File_Name = Dir(WINSYSDIR & "\*.cpl", vbNormal + vbHidden)
    
    Do While File_Name <> ""
        Dont_Load = 0
        blnFound = False
        For Ctr = 1 To NO_CON_APP
            If LCase$(Con_App(Ctr).Name) = LCase$(File_Name) Then
                blnFound = True
                Exit For
            End If
        Next
        If blnFound = True Then
            lstHideApplet.AddItem (Con_App(Ctr).Description)
        Else
            lstHideApplet.AddItem (File_Name)
        End If
        
        Open WINDIR & "\Control.ini" For Input As #1
            Do While Not EOF(1)   ' Loop until end of file.
                blnDont_Load = False
                Input #1, FString ' Read data
                If LCase$(FString) = "[don't load]" Then
                    Dont_Load = 1
                ElseIf Left(Trim(FString), 1) = "[" Then
                    If Dont_Load = 1 Then
                        Exit Do
                    End If
                Else
                    ESCEQ = InStr(1, FString, "=")
                    If ESCEQ <> 0 Then
                        If LCase$(Left(FString, ESCEQ - 1)) = LCase$(File_Name) Then
                            If Dont_Load = 1 And Trim(Right(FString, Len(FString) - ESCEQ)) <> "" Then
                                blnDont_Load = True
                                Exit Do
                            End If
                        End If
                    End If
                End If
            Loop
        Close #1
        If blnDont_Load = False Then
            lstHideApplet.Selected(LstCtr) = True
        Else
            lstHideApplet.Selected(LstCtr) = False
        End If
        File_Name = Dir
        LstCtr = LstCtr + 1
    Loop
    
    lstHideApplet.ListIndex = 0
    cmdApply.Enabled = False

Exit Sub
ErrHand:
    MsgBox "Error Occurred while Reading Control Panel Applets !!" & vbCrLf & "Procedure : Read_ControlPanel_Applets " & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub Write_ControlPanel_Applets()
On Error GoTo ErrHand

    Dim blnDontLoad As Boolean
    Dim blnWritten As Boolean
    Dim blnFound As Boolean
    Dim ESCEQ As Long
    Dim i As Integer
    
    Open WINDIR & "\Control.ini" For Input As #1
        On Error Resume Next
        Kill WINDIR & "\Control.bak"
        On Error GoTo ErrHand
        Open WINDIR & "\Control.bak" For Output As #2
            Do While Not EOF(1)   ' Loop until end of file.
                Input #1, FString ' Read data
                If LCase$(FString) = "[don't load]" And blnWritten = False Then
                    blnDontLoad = True
                    Print #2, "[don't load]"
                    For Ctr = 0 To lstHideApplet.ListCount - 1
                        If lstHideApplet.Selected(Ctr) = False Then
                            
                            blnFound = False
                            
                            For i = 1 To NO_CON_APP
                                If LCase$(Con_App(i).Description) = LCase$(lstHideApplet.List(Ctr)) Then
                                    blnFound = True
                                    Exit For
                                End If
                            Next
                                                        
                            If blnFound = True Then
                                Print #2, Con_App(i).Name & "=no"
                            Else
                                Print #2, lstHideApplet.List(Ctr) & "=no"
                            End If
                            
                        End If
                    Next
                    blnWritten = True
                ElseIf Left(Trim(FString), 1) = "[" Then
                    blnDontLoad = False
                    Print #2, FString
                Else
                    If blnDontLoad = True Then
                        ESCEQ = InStr(1, FString, "=")
                        If ESCEQ <> 0 Then
                            
                            blnFound = False
                            For i = 1 To NO_CON_APP
                                If LCase$(Con_App(i).Name) = LCase$(Left(FString, ESCEQ - 1)) Then
                                    blnFound = True
                                    Exit For
                                End If
                            Next
                            
                            If blnFound = False Then
                                For Ctr = 0 To lstHideApplet.ListCount - 1
                                    If LCase$(Left(FString, ESCEQ - 1)) = LCase$(lstHideApplet.List(Ctr)) Then
                                        blnFound = True
                                        Exit For
                                    End If
                                Next
                            End If
                            
                            If blnFound = False Then
                                Print #2, FString
                            End If
                        Else
                            Print #2, FString
                        End If
                    Else
                        Print #2, FString
                    End If
                End If
            Loop
            If blnWritten = False Then
                Dim Printed_DontLoad As Boolean
                
                For Ctr = 0 To lstHideApplet.ListCount - 1
                    If lstHideApplet.Selected(Ctr) = False Then
                        
                        blnFound = False
                        If Printed_DontLoad = False Then
                            Print #2, "[don't load]"
                            Printed_DontLoad = True
                        End If
                        For i = 1 To NO_CON_APP
                            If LCase$(Con_App(i).Description) = LCase$(lstHideApplet.List(Ctr)) Then
                                blnFound = True
                                Exit For
                            End If
                        Next
                                                    
                        If blnFound = True Then
                            Print #2, Con_App(i).Name & "=no"
                        Else
                            Print #2, lstHideApplet.List(Ctr) & "=no"
                        End If
                        
                    End If
                Next
            End If
        Close #2
    Close #1
    Kill WINDIR & "\Control.ini"
    Name WINDIR & "\Control.bak" As WINDIR & "\Control.ini"
Exit Sub
ErrHand:
    MsgBox "Error Occurred while Writing Control Panel Settings !!" & vbCrLf & "Procedure : Write_ControlPanel_Applets " & vbCrLf & Err.Description, vbCritical
End Sub


Private Sub cmdOK_Click()
cmdApply_Click
Unload Me
End Sub




Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
drag Me
End Sub

Private Sub lstHideApplet_Click()
cmdApply.Enabled = True
End Sub

