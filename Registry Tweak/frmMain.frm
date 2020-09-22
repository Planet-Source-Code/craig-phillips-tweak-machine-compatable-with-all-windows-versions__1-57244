VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000006&
   BorderStyle     =   0  'None
   ClientHeight    =   11490
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   15330
   ControlBox      =   0   'False
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11490
   ScaleWidth      =   15330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1320
      Left            =   780
      ScaleHeight     =   1320
      ScaleWidth      =   7305
      TabIndex        =   52
      Top             =   8580
      Visible         =   0   'False
      Width           =   7305
      Begin Tweak.checkbox chkHideVirtualMemory 
         Height          =   315
         Left            =   0
         TabIndex        =   58
         Top             =   780
         Width           =   6300
         _ExtentX        =   11113
         _ExtentY        =   556
         LabelText       =   "Hide the Virtual Memory Button"
      End
      Begin Tweak.checkbox chkHideFileSystem 
         Height          =   315
         Left            =   0
         TabIndex        =   57
         Top             =   390
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   556
         LabelText       =   "Hide the File System Button"
      End
      Begin Tweak.checkbox chkTime 
         Height          =   315
         Left            =   0
         TabIndex        =   53
         Top             =   0
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   556
         ForeColor       =   0
         LabelText       =   "Ability to configure the System Time and Date"
         ColourScheme    =   4
      End
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   3690
      Left            =   8895
      ScaleHeight     =   3690
      ScaleWidth      =   5655
      TabIndex        =   46
      Top             =   2910
      Visible         =   0   'False
      Width           =   5655
      Begin Tweak.checkbox chkHideSettings 
         Height          =   315
         Left            =   0
         TabIndex        =   47
         Top             =   390
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         ForeColor       =   0
         LabelText       =   "Hide the Display - Settings, Web && Effects Tab"
         ColourScheme    =   4
      End
      Begin Tweak.checkbox chkHideAppearance 
         Height          =   315
         Left            =   0
         TabIndex        =   48
         Top             =   1560
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   556
         LabelText       =   "Hide the Display - Appearance Tab"
      End
      Begin Tweak.checkbox chkDenyDisplay 
         Height          =   315
         Left            =   0
         TabIndex        =   49
         Top             =   0
         Width           =   5610
         _ExtentX        =   9895
         _ExtentY        =   556
         ForeColor       =   0
         LabelText       =   "Deny Access to Display Settings Page"
         ColourScheme    =   4
      End
      Begin Tweak.checkbox chkHideScreenSaver 
         Height          =   315
         Left            =   0
         TabIndex        =   50
         Top             =   780
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   556
         ForeColor       =   0
         LabelText       =   "Hide the Display - Screen Saver Tab"
         ColourScheme    =   4
      End
      Begin Tweak.checkbox chkHideBackground 
         Height          =   315
         Left            =   0
         TabIndex        =   51
         Top             =   1170
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
         ForeColor       =   0
         LabelText       =   "Hide the Display - Background Tab"
         ColourScheme    =   4
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   4365
      Left            =   795
      ScaleHeight     =   4365
      ScaleWidth      =   6855
      TabIndex        =   36
      Top             =   2940
      Visible         =   0   'False
      Width           =   6855
      Begin Tweak.checkbox chkDisablePrinterDeletion 
         Height          =   315
         Left            =   0
         TabIndex        =   37
         Top             =   390
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   556
         ForeColor       =   16761024
         LabelText       =   "Disable the Deletion of Printers"
         ColourScheme    =   2
      End
      Begin Tweak.checkbox chkRestrictPassword 
         Height          =   315
         Left            =   0
         TabIndex        =   38
         Top             =   1560
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   556
         ForeColor       =   0
         LabelText       =   "Restrict Access to the Passwords Applet"
         ColourScheme    =   4
      End
      Begin Tweak.checkbox chkRestrictUserProfile 
         Height          =   315
         Left            =   0
         TabIndex        =   39
         Top             =   1950
         Width           =   5970
         _ExtentX        =   10530
         _ExtentY        =   556
         ForeColor       =   0
         LabelText       =   "Restrict Access to the User Profiles Page"
         ColourScheme    =   4
      End
      Begin Tweak.checkbox chkHidePrinterGeneralDetails 
         Height          =   315
         Left            =   0
         TabIndex        =   40
         Top             =   780
         Width           =   7020
         _ExtentX        =   12383
         _ExtentY        =   556
         ForeColor       =   16761024
         LabelText       =   "Hide the General and Details Printer Pages"
         ColourScheme    =   2
      End
      Begin Tweak.checkbox chkHideChangePassword 
         Height          =   315
         Left            =   0
         TabIndex        =   41
         Top             =   1170
         Width           =   7080
         _ExtentX        =   12488
         _ExtentY        =   556
         ForeColor       =   0
         LabelText       =   "Hide the Change Passwords Tab"
         ColourScheme    =   4
      End
      Begin Tweak.checkbox chkDisablePrinterAddition 
         Height          =   315
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   556
         ForeColor       =   16761024
         LabelText       =   "Disable the Addition of Printers"
         ColourScheme    =   2
      End
      Begin Tweak.XPButton XPButton4 
         Height          =   465
         Left            =   225
         TabIndex        =   43
         Top             =   2445
         Width           =   2130
         _ExtentX        =   3757
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
         Caption         =   "&Control Panel Applets"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   795
      ScaleHeight     =   975
      ScaleWidth      =   7305
      TabIndex        =   27
      Top             =   8565
      Width           =   7305
      Begin Tweak.checkbox chkAddCopyFolder 
         Height          =   315
         Left            =   0
         TabIndex        =   29
         Top             =   390
         Width           =   5790
         _ExtentX        =   10213
         _ExtentY        =   556
         LabelText       =   "Add a Menu Option to Copy Folders"
      End
      Begin Tweak.checkbox chkAddCommandPrompt 
         Height          =   315
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   556
         ForeColor       =   0
         LabelText       =   "Add Command Prompt Option to Every Folder"
         ColourScheme    =   4
      End
   End
   Begin Tweak.colourFrame colourFrame3 
      Height          =   2445
      Left            =   315
      TabIndex        =   26
      Top             =   7695
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   4313
      BackColor       =   -2147483625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FillStyle       =   0
      LabelText       =   "Explorer"
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4260
      Top             =   1710
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   3690
      Left            =   8895
      ScaleHeight     =   3690
      ScaleWidth      =   5655
      TabIndex        =   15
      Top             =   2910
      Width           =   5655
      Begin Tweak.checkbox chkDisableRecDoc 
         Height          =   315
         Left            =   0
         TabIndex        =   21
         Top             =   390
         Width           =   4980
         _ExtentX        =   8784
         _ExtentY        =   556
         ForeColor       =   16761024
         LabelText       =   "Disable Recent Documents History"
      End
      Begin Tweak.checkbox chkDisableRegistry 
         Height          =   315
         Left            =   0
         TabIndex        =   20
         Top             =   1560
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   556
         LabelText       =   "Disable Registry Editing Tools"
      End
      Begin Tweak.XPButton Command2 
         Height          =   465
         Left            =   255
         TabIndex        =   19
         Top             =   2190
         Width           =   2130
         _ExtentX        =   3757
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
         Caption         =   "&Hide Drives"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin Tweak.checkbox chkDisableRightClick 
         Height          =   315
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   5610
         _ExtentX        =   9895
         _ExtentY        =   556
         ForeColor       =   16761024
         LabelText       =   "Disable Right Click on Desktop && Explorer"
      End
      Begin Tweak.checkbox chkHideInternetExplorer 
         Height          =   315
         Left            =   0
         TabIndex        =   17
         Top             =   780
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   556
         ForeColor       =   16761024
         LabelText       =   "Hide the Internet Explorer Icon"
      End
      Begin Tweak.checkbox chkHideAllDesktop 
         Height          =   315
         Left            =   0
         TabIndex        =   16
         Top             =   1170
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
         ForeColor       =   16761024
         LabelText       =   "Hide All Items on the Desktop"
      End
   End
   Begin Tweak.colourFrame colourFrame2 
      Height          =   5580
      Left            =   8490
      TabIndex        =   14
      Top             =   1890
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   9843
      BackColor       =   -2147483642
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FillStyle       =   0
      LabelText       =   "Security"
   End
   Begin Tweak.XPButton cmdCancel 
      Height          =   450
      Left            =   11115
      TabIndex        =   11
      Top             =   10290
      Width           =   1995
      _ExtentX        =   3519
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
   Begin VB.PictureBox frameGen 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   4365
      Left            =   795
      ScaleHeight     =   4365
      ScaleWidth      =   6855
      TabIndex        =   1
      Top             =   2925
      Width           =   6855
      Begin Tweak.checkbox chkDisableTaskbarConMenu 
         Height          =   315
         Left            =   0
         TabIndex        =   23
         Top             =   2340
         Width           =   4170
         _ExtentX        =   7355
         _ExtentY        =   556
         ForeColor       =   16761024
         LabelText       =   "Disable Taskbar Context Menus"
      End
      Begin Tweak.checkbox chkDisableNetwork 
         Height          =   315
         Left            =   0
         TabIndex        =   22
         Top             =   3495
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   556
         LabelText       =   "Disable Network Control Panel"
      End
      Begin Tweak.checkbox chkUser 
         Height          =   315
         Left            =   0
         TabIndex        =   10
         Top             =   390
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   556
         ForeColor       =   12632319
         LabelText       =   "Don't display the last Username logged on"
      End
      Begin Tweak.checkbox chkDisableNewMenu 
         Height          =   315
         Left            =   0
         TabIndex        =   9
         Top             =   3105
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   556
         LabelText       =   "Disable the New Menu Item"
      End
      Begin Tweak.checkbox chkHideFolderOptions 
         Height          =   315
         Left            =   0
         TabIndex        =   8
         Top             =   1560
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   556
         ForeColor       =   16761024
         LabelText       =   "Hide the Folder Options from the Tools/View Menu"
      End
      Begin Tweak.checkbox chkHideControlPanel 
         Height          =   315
         Left            =   0
         TabIndex        =   7
         Top             =   1950
         Width           =   5205
         _ExtentX        =   9181
         _ExtentY        =   556
         ForeColor       =   16761024
         LabelText       =   "Hide Control Panel and Printers"
      End
      Begin Tweak.checkbox chkDisableStart 
         Height          =   315
         Left            =   0
         TabIndex        =   2
         Top             =   780
         Width           =   7020
         _ExtentX        =   12383
         _ExtentY        =   556
         ForeColor       =   16761024
         LabelText       =   "Disable Menu Bars and the Start Button "
         ColourScheme    =   2
      End
      Begin Tweak.checkbox chkDisableSave 
         Height          =   315
         Left            =   0
         TabIndex        =   3
         Top             =   1170
         Width           =   7080
         _ExtentX        =   12488
         _ExtentY        =   556
         ForeColor       =   16761024
         LabelText       =   "Disable Saving of Windows Settings at Shut Down or Restart"
      End
      Begin Tweak.checkbox chkExpandingControlPanel 
         Height          =   315
         Left            =   0
         TabIndex        =   4
         Top             =   2730
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   556
         LabelText       =   "Add an Expanding Control Panel to Start Menu"
      End
      Begin Tweak.checkbox chkRemoveShortCutTo 
         Height          =   315
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   556
         ForeColor       =   12632319
         LabelText       =   "Remove the 'Shortcut to...' Prefix on Shortcuts"
      End
      Begin Tweak.checkbox chkUseNotepad 
         Height          =   315
         Left            =   0
         TabIndex        =   6
         Top             =   3885
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   556
         LabelText       =   "Easily Use Notepad to Open a File"
      End
   End
   Begin Tweak.colourFrame colourFrame1 
      Height          =   5580
      Left            =   330
      TabIndex        =   0
      Top             =   1890
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   9843
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FillColour      =   16744448
      FillStyle       =   0
      LabelText       =   "General"
   End
   Begin Tweak.XPButton cmdApply 
      Height          =   450
      Left            =   13155
      TabIndex        =   12
      Top             =   10290
      Width           =   1995
      _ExtentX        =   3519
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
      Left            =   9090
      TabIndex        =   13
      Top             =   10290
      Width           =   1995
      _ExtentX        =   3519
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
   Begin Tweak.XPButton XPButton1 
      Height          =   450
      Left            =   4635
      TabIndex        =   24
      Top             =   10290
      Width           =   1995
      _ExtentX        =   3519
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
      Caption         =   "&Change Password"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin Tweak.XPButton cmdReg 
      Height          =   450
      Left            =   4650
      TabIndex        =   25
      Top             =   10290
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
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
      Caption         =   "&Register"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin Tweak.XPButton XPButton2 
      Height          =   450
      Left            =   2205
      TabIndex        =   34
      Top             =   10290
      Width           =   1995
      _ExtentX        =   3519
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
      Caption         =   "&Next >>>"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin Tweak.XPButton XPButton3 
      Height          =   450
      Left            =   150
      TabIndex        =   35
      Top             =   10290
      Width           =   1995
      _ExtentX        =   3519
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
      Caption         =   "<<< &Previous"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   8895
      ScaleHeight     =   1410
      ScaleWidth      =   5925
      TabIndex        =   54
      Top             =   8655
      Visible         =   0   'False
      Width           =   5925
      Begin Tweak.checkbox chkHideHardwareProfiles 
         Height          =   315
         Left            =   0
         TabIndex        =   55
         Top             =   390
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   556
         LabelText       =   "Hide the Hardware Profiles Page"
      End
      Begin Tweak.checkbox chkHideDeviceManager 
         Height          =   315
         Left            =   0
         TabIndex        =   56
         Top             =   0
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   556
         LabelText       =   "Hide the Device Manager Page"
      End
      Begin Tweak.XPButton XPButton5 
         Height          =   450
         Left            =   330
         TabIndex        =   59
         Top             =   840
         Width           =   1995
         _ExtentX        =   3519
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
         Caption         =   "R&emove Programs"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   8895
      ScaleHeight     =   975
      ScaleWidth      =   5925
      TabIndex        =   31
      Top             =   8580
      Width           =   5925
      Begin Tweak.checkbox chkEnableRenameDelete 
         Height          =   315
         Left            =   0
         TabIndex        =   33
         Top             =   390
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   556
         LabelText       =   "Enable Recycle Bin Rename and Delete"
      End
      Begin Tweak.checkbox chkAddMoveFolder 
         Height          =   315
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   556
         LabelText       =   "Add a Menu Option to Move Folders"
      End
   End
   Begin Tweak.colourFrame colourFrame4 
      Height          =   2445
      Left            =   8490
      TabIndex        =   30
      Top             =   7695
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   4313
      BackColor       =   -2147483625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FillStyle       =   0
      LabelText       =   "Explorer Cont."
   End
   Begin Tweak.XPButton XPButton6 
      Height          =   450
      Left            =   6660
      TabIndex        =   60
      Top             =   10290
      Width           =   1995
      _ExtentX        =   3519
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
      Caption         =   "&Help"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tweak Machine is compatable with"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1275
      Left            =   3900
      TabIndex        =   45
      Top             =   30
      Width           =   11010
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1275
      Left            =   3900
      TabIndex        =   44
      Top             =   360
      Width           =   11010
   End
   Begin VB.Image Image1 
      Height          =   1650
      Left            =   135
      Picture         =   "frmMain.frx":0000
      Top             =   90
      Width           =   3750
   End
   Begin VB.Menu try 
      Caption         =   "Tray"
      Visible         =   0   'False
      Begin VB.Menu showe 
         Caption         =   "&Show"
      End
      Begin VB.Menu hidee 
         Caption         =   "&Hide"
      End
      Begin VB.Menu poo 
         Caption         =   "-"
      End
      Begin VB.Menu exite 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pageNo As Integer
Private Sub chkAddCommandPrompt_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkAddCopyFolder_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkAddMoveFolder_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkDenyDisplay_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkDisableNetwork_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkDisableNewMenu_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkDisablePrinterAddition_Click()
cmdApply.Enabled = True
logo = 1
End Sub

Private Sub chkDisablePrinterDeletion_Click()
cmdApply.Enabled = True
logo = 1
End Sub

Private Sub chkDisableRecDoc_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkDisableRegistry_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkDisableRightClick_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkDisableSave_Click()
cmdApply.Enabled = True
logo = 1
End Sub

Private Sub chkDisableStart_Click()
cmdApply.Enabled = True
logo = 1
End Sub

Private Sub chkDisableTaskbarConMenu_Click()
cmdApply.Enabled = True
logo = 1
End Sub

Private Sub chkEnableRenameDelete_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkExpandingControlPanel_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkHideAllDesktop_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkHideAppearance_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkHideBackground_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkHideChangePassword_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkHideControlPanel_Click()
cmdApply.Enabled = True
logo = 1
End Sub

Private Sub chkHideFolderOptions_Click()
cmdApply.Enabled = True
logo = 1
End Sub

Private Sub chkHideInternetExplorer_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkHidePrinterGeneralDetails_Click()
cmdApply.Enabled = True
logo = 1
End Sub

Private Sub chkHideScreenSaver_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkHideSettings_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkRemoveShortCutTo_Click()
cmdApply.Enabled = True
rese = 1
End Sub

Private Sub chkRestrictPassword_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkRestrictUserProfile_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkTime_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkUseNotepad_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkUser_Click()
cmdApply.Enabled = True
rese = 1
End Sub

Private Sub cmdApply_Click()
Apply_Security_Page
Read_Security_Page
Apply_General_Page
Read_General_Page
Apply_Explorer_Page
Read_Explorer_Page
Apply_ControlPanel_Settings
Read_ControlPanel_Settings
Apply_Properties_Page
Read_Properties_Page
cmdApply.Enabled = False
If rese = 1 Or logo = 1 Then
ask = 1
End If

End Sub

Private Sub cmdOK_Click()
cmdApply_Click
cmdCancel_Click
End Sub

Private Sub cmdReg_Click()
frmDemoCode.Show 1
End Sub

Private Sub Command2_Click()
frmHideDrive.Show 1
End Sub

Private Sub exite_Click()
cmdCancel_Click
End Sub

Private Sub Form_GotFocus()
progInit = 1
End Sub

Private Sub Form_Load()
On Error GoTo ghty
progInit = 1
  Read_AddRemove_Programs
  Read_ControlPanel_Settings
  Read_Explorer_Page
  Read_General_Page
  Read_Properties_Page
  Read_Security_Page
colourFrame1.FillColour = RGB(128, 128, 128)
colourFrame1.ForeColor = RGB(255, 0, 255)
colourFrame1.LabelColour = RGB(255, 255, 255)
colourFrame2.FillColour = RGB(128, 128, 128)
colourFrame2.ForeColor = RGB(255, 0, 255)
colourFrame2.LabelColour = RGB(255, 255, 255)
colourFrame3.FillColour = RGB(128, 128, 128)
colourFrame3.ForeColor = RGB(255, 0, 255)
colourFrame3.LabelColour = RGB(255, 255, 255)
colourFrame4.FillColour = RGB(128, 128, 128)
colourFrame4.ForeColor = RGB(255, 0, 255)
colourFrame4.LabelColour = RGB(255, 255, 255)

pageNo = 1
Label1.Caption = GetWindowsVersion
Label1.ForeColor = RGB(90, 190, 255)

    Exit Sub
ghty:
End Sub
Private Sub END_Tweakui()
End
End Sub
Private Sub cmdCancel_Click()
Unload Me
    End

End Sub

Private Sub Form_LostFocus()
progInit = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim pooe As Integer
pooe = 0
If rese = 1 And logo = 1 And ask = 1 Then
frmReset.Show 1
pooe = 1
End If
If rese = 1 And ask = 1 And pooe = 0 Then
frmReset.Show 1
End If
If logo = 1 And ask = 1 And pooe = 0 Then
frmLogoff.Show 1
End If
ChangeRes screenResW, screenResH
Cls
End Sub



Private Sub hidee_Click()
Me.WindowState = 1
End Sub

Private Sub showe_Click()
Me.WindowState = 0
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Unload Me
Unload frmMsg
Unload frmLogin
Unload frmHideDrive
Unload frmSettings
End Sub





Private Sub XPButton1_Click()
frmSettings.Show 1
End Sub
Public Sub Apply_General_Page()
On Error GoTo ErrHand

    Hive_Key = HKEY_CURRENT_USER
    Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"

'General
    'Disable Save Settings at Exit
    Write_DWORD Hive_Key, Sub_Key, "NoSaveSettings", chkDisableSave.Value

'General
    'Disable Network Control Panel
    Write_DWORD Hive_Key, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoNetSetup", frmMain.chkDisableNetwork.Value
    
    'Hide Control Panel on Start Menu
    Write_DWORD Hive_Key, Sub_Key, "NoSetFolders", frmMain.chkHideControlPanel.Value
    
    'Hide Folder Options
    Write_DWORD Hive_Key, Sub_Key, "NoFolderOptions", frmMain.chkHideFolderOptions.Value
    
    'Disable the New Menu Item in Context Menu
    Disable_NewMenu (chkDisableNewMenu.Value)
    'Display last Username logon
SetKeyValue HKEY_LOCAL_MACHINE, "Network\Logon", "DontShowLastUser", frmMain.chkUser.Value, REG_DWORD
'General
    'Disable Menu Bars and the Start Button
    Disable_StartMenu (chkDisableStart.Value)
    
    Hive_Key = HKEY_CURRENT_USER
    Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    'Add an Expanding Control Panel to Start Menu
    Dim sCPanelExt As String
    Dim sStartControl As String
    sCPanelExt = "{21EC2020-3AEA-1069-A2DD-08002B30309D}"
    If chkExpandingControlPanel.Value = 0 Then 'disabled
        If Dir(StartMenu_DIR & "\*." & sCPanelExt, vbDirectory Or vbHidden Or vbNormal) <> "" Then 'previous enabled
            sStartControl = Dir(StartMenu_DIR & "\*." & sCPanelExt, vbDirectory Or vbHidden Or vbNormal)
            RmDir StartMenu_DIR & "\" & sStartControl
        End If
    Else 'enabled
        If Dir(StartMenu_DIR & "\*." & sCPanelExt, vbDirectory Or vbHidden Or vbNormal) = "" Then 'previous disabled
            MkDir (StartMenu_DIR & "\Control Panel." & sCPanelExt)
        End If
    End If
    
    
    'Disable Taskbar Context Menus
    Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    Write_DWORD Hive_Key, Sub_Key, "NoTrayContextMenu", chkDisableTaskbarConMenu.Value

    'General
    'Remove the 'Shortcut to...' Prefix on Shortcuts
    Hive_Key = HKEY_CURRENT_USER
    Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Explorer"
    
    Open_SubKey Hive_Key, Sub_Key
    Query_Value REG_BINARY, "link"
    lShortcut = BD_Value
    

    
    If frmMain.chkRemoveShortCutTo.Value = 0 Then 'Tweak Disabled
        If lShortcut = 0 Then 'Previously Enabled
            Open_SubKey HKEY_LOCAL_MACHINE, "Software\" & Software_Name & "\Main" 'My Software Location
            If Query_Value(REG_BINARY, "Shortcut") = ERROR_SUCCESS Then
                Open_SubKey Hive_Key, Sub_Key
                Create_Value REG_BINARY, "link", BD_Value
            Else
                Open_SubKey Hive_Key, Sub_Key
                Create_Value REG_BINARY, "link", CLng(7)
            End If
        End If
    Else 'Tweak Enabled
        If lShortcut <> 0 Then 'Previously Disabled
            Open_SubKey HKEY_LOCAL_MACHINE, "Software\" & Software_Name & "\Main"
            Create_Value REG_BINARY, "Shortcut", lShortcut
            
            Open_SubKey Hive_Key, Sub_Key
            Create_Value REG_BINARY, "link", 0 'tweak Enabled
        End If
    End If
    RegCloseKey HRegKey
    
    'Easily Use Notepad to Open a File
    Hive_Key = HKEY_CLASSES_ROOT
    Sub_Key = "*\Shell\Open\command"
    blnFlag = False
    sData = "notepad.exe %1"
    
    If frmMain.chkUseNotepad.Value = 0 Then 'disabled
        If Open_SubKey(Hive_Key, Sub_Key) = ERROR_SUCCESS Then 'Previously Enabled
            Query_Value REG_SZ, ""
            If LCase$(S_Value) = sData Then
                RegDeleteKey Hive_Key, "*\Shell\Open"
            End If
        End If
    Else 'enabled
        If Open_SubKey(Hive_Key, Sub_Key) = ERROR_SUCCESS Then 'Previously Enabled
            Query_Value REG_SZ, ""
            If LCase$(S_Value) <> sData Then
                blnFlag = True
            End If
        Else
            blnFlag = True
        End If
        
        If blnFlag = True Then
            Create_SubKey Hive_Key, "*\shell", "Open\command"
            Open_SubKey Hive_Key, "*\Shell\Open"
            Create_Value REG_SZ, "", "Open With Notepad"
            Open_SubKey Hive_Key, "*\Shell\Open\command"
            Create_Value REG_SZ, "", sData
        End If
        RegCloseKey HRegKey
    End If
Exit Sub
ErrHand:
    Msgex "Error Occurred while Applying General Information Page !! " & vbCrLf & "Procedure : Apply_General_Page" & vbCrLf & Err.Description, , Error
End Sub
Private Sub Apply_Security_Page()


Hive_Key = HKEY_CURRENT_USER
Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
'Type: REG_DWORD (DWORD Value)
'Value: (0 = disabled, 1 = enabled)

    
'Desktop
    Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    'Hide All Items on the Desktop
    Write_DWORD Hive_Key, Sub_Key, "NoDesktop", frmMain.chkHideAllDesktop.Value
    
    'Hide the Internet Explorer Icon
    Write_DWORD Hive_Key, Sub_Key, "NoInternetIcon", frmMain.chkHideInternetExplorer.Value

    'Disable the Ability to Right Click on the Desktop && Explorer
    Write_DWORD Hive_Key, Sub_Key, "NoViewContextMenu", frmMain.chkDisableRightClick.Value
    
'Registry
    Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    'Disable Registry Editing Tools
    Write_DWORD Hive_Key, Sub_Key, "DisableRegistryTools", frmMain.chkDisableRegistry.Value
'History
    'Disable Recent Documents History
    Write_DWORD Hive_Key, Sub_Key, "NoRecentDocsHistory", chkDisableRecDoc.Value
        
Exit Sub
ErrHand:
    MsgBox "Error Occurred while Applying Security Settings !!" & vbCrLf & "Procedure : Apply_Security_Page " & vbCrLf & Err.Description, vbCritical
End Sub
Private Sub Disable_NewMenu(blnEnabled As Boolean)
On Error GoTo ErrHand
            
    Dim hKey As Long
    Dim sKey As String
    Dim sEKey As String
    hKey = HKEY_CLASSES_ROOT
    
    If blnEnabled = True Then
        sKey = "CLSID\{D969A300-E7FF-11d0-A93B-00A0C90F2719}"
        sEKey = "CLSID\{-D969A300-E7FF-11d0-A93B-00A0C90F2719}"
    Else
        sKey = "CLSID\{-D969A300-E7FF-11d0-A93B-00A0C90F2719}"
        sEKey = "CLSID\{D969A300-E7FF-11d0-A93B-00A0C90F2719}"
    End If
    
    Re_Init_Val
    If RegOpenKey(hKey, sKey, HRegKey) = ERROR_SUCCESS Then
        Create_SubKey hKey, sEKey, "InProcServer32"
        
        Open_SubKey hKey, sKey
        Query_Value REG_DWORD, "flags"
        Open_SubKey hKey, sEKey
        Create_Value REG_DWORD, "flags", BD_Value
        
        Open_SubKey hKey, sKey
        Query_Value REG_SZ, ""
        Open_SubKey hKey, sEKey
        Create_Value REG_SZ, "", S_Value
        
        Open_SubKey hKey, sKey & "\InProcServer32"
        Query_Value REG_SZ, "ThreadingModel"
        Open_SubKey hKey, sEKey & "\InProcServer32"
        Create_Value REG_SZ, "ThreadingModel", S_Value
        
        Open_SubKey hKey, sKey & "\InProcServer32"
        Query_Value REG_SZ, ""
        Open_SubKey hKey, sEKey & "\InProcServer32"
        Create_Value REG_SZ, "", S_Value
        
        'CopyKeys hKey, sKey, sEKey
        RegDeleteKey hKey, sKey 'Delete old key
    End If
    
Exit Sub
ErrHand:
    Msgex "Error Occurred while Enabling/Disabling New Menu !! " & vbCrLf & "Procedure : Disable_NewMenu " & vbCrLf & Err.Description, , Error
End Sub
Private Sub Disable_StartMenu(blnEnabled As Boolean)
On Error GoTo ErrHand
    
    Dim hKey As Long
    Dim sKey As String
    Dim sEKey As String
    hKey = HKEY_CLASSES_ROOT
    
    If blnEnabled = True Then
        sKey = "CLSID\{5b4dae26-b807-11d0-9815-00c04fd91972}"
        sEKey = "CLSID\{-5b4dae26-b807-11d0-9815-00c04fd91972}"
    Else
        sKey = "CLSID\{-5b4dae26-b807-11d0-9815-00c04fd91972}"
        sEKey = "CLSID\{5b4dae26-b807-11d0-9815-00c04fd91972}"
    End If
    
    Re_Init_Val
    If RegOpenKey(hKey, sKey, HRegKey) = ERROR_SUCCESS Then
        
        Create_SubKey hKey, sEKey, "InProcServer32"
        
        Open_SubKey hKey, sKey & "\InProcServer32"
        Query_Value REG_SZ, ""
        Open_SubKey hKey, sEKey & "\InProcServer32"
        Create_Value REG_SZ, "", S_Value
        
        Open_SubKey hKey, sKey & "\InProcServer32"
        Query_Value REG_SZ, "ThreadingModel"
        Open_SubKey hKey, sEKey & "\InProcServer32"
        Create_Value REG_SZ, "ThreadingModel", S_Value
        
        Open_SubKey hKey, sKey
        Query_Value REG_SZ, ""
        Open_SubKey hKey, sEKey
        Create_Value REG_SZ, "", S_Value
        
        RegDeleteKey hKey, sKey
        
    End If
    
Exit Sub
ErrHand:
    Msgex "Error Occurred while Enabling/Disabling Start Menu Settings !! " & vbCrLf & "Procedure : Disable_StartMenu " & vbCrLf & Err.Description, , Error
End Sub
Private Sub Apply_Explorer_Page()
On Error GoTo ErrHand

    Dim blnFlag As Boolean
    Dim lShortcut As Long
    Dim sData As String
    Dim hKey As Long
    Dim sKey As String
    
'Folder
    'Add Command Prompt Option to Every Folder
    Hive_Key = HKEY_CLASSES_ROOT
    Sub_Key = "Directory\shell\Command\Command"
    blnFlag = False
    sData = "command.com /k cd " & """" & "%1" & """"
    
    If frmMain.chkAddCommandPrompt.Value = 0 Then 'disabled
         If Open_SubKey(Hive_Key, Sub_Key) = ERROR_SUCCESS Then 'Previously Enabled
            Query_Value REG_SZ, ""
            If LCase$(S_Value) = sData Then
                RegDeleteKey Hive_Key, "Directory\shell\Command"
            End If
        End If
    Else 'enabled
        If Open_SubKey(Hive_Key, Sub_Key) = ERROR_SUCCESS Then 'Previously Enabled
            Query_Value REG_SZ, ""
            If LCase$(S_Value) <> sData Then
                blnFlag = True
            End If
        Else
            blnFlag = True
        End If
        If blnFlag = True Then
            Create_SubKey Hive_Key, "Directory\shell", "Command\Command"
            
            Open_SubKey Hive_Key, "Directory\shell\Command"
            Create_Value REG_SZ, "", "Open Command Prompt Here"
            
            Open_SubKey Hive_Key, "Directory\shell\Command\Command"
            Create_Value REG_SZ, "", sData
        End If
        RegCloseKey HRegKey
    End If
    
    'Add a Menu Option to Copy Folders
    Hive_Key = HKEY_CLASSES_ROOT
    Sub_Key = "Directory\shellex\ContextMenuHandlers\Copy to Folder"
    blnFlag = False
    sData = "{C2FBB630-2971-11d1-A18C-00C04FD75D13}"
    
    If frmMain.chkAddCopyFolder.Value = 0 Then 'disabled
        If Open_SubKey(Hive_Key, Sub_Key) = ERROR_SUCCESS Then 'Previously Enabled
            Query_Value REG_SZ, ""
            If S_Value = sData Then
                RegDeleteKey Hive_Key, Sub_Key
            End If
        End If
    Else 'enabled
        If Open_SubKey(Hive_Key, Sub_Key) = ERROR_SUCCESS Then 'Previously Enabled
            Query_Value REG_SZ, ""
            If S_Value <> sData Then
                blnFlag = True
            End If
        Else
            blnFlag = True
        End If
        If blnFlag = True Then
            Create_SubKey Hive_Key, Sub_Key, ""
            Open_SubKey Hive_Key, Sub_Key
            Create_Value REG_SZ, "", sData
        End If
        RegCloseKey HRegKey
    End If
    
    'Add a Menu Option to Move Folders
    Hive_Key = HKEY_CLASSES_ROOT
    Sub_Key = "Directory\shellex\ContextMenuHandlers\Move to Folder"
    blnFlag = False
    sData = "{C2FBB631-2971-11d1-A18C-00C04FD75D13}"
    
    If frmMain.chkAddMoveFolder.Value = 0 Then
        If Open_SubKey(Hive_Key, Sub_Key) = ERROR_SUCCESS Then
            Query_Value REG_SZ, ""
            If S_Value = sData Then
                RegDeleteKey Hive_Key, Sub_Key
            End If
        End If
    Else
        If Open_SubKey(Hive_Key, Sub_Key) = ERROR_SUCCESS Then
            Query_Value REG_SZ, ""
            If S_Value <> sData Then
                blnFlag = True
            End If
        Else
            blnFlag = True
        End If
        If blnFlag = True Then
            Create_SubKey Hive_Key, Sub_Key, ""
            Open_SubKey Hive_Key, Sub_Key
            Create_Value REG_SZ, "", sData
        End If
        RegCloseKey HRegKey
    End If

    'Enable Recycle Bin Rename and Delete
    Hive_Key = HKEY_CLASSES_ROOT
    Sub_Key = "CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\ShellFolder"
    'Enable Delete  - 536871280
    'Disable Delete - 536871232
    blnFlag = False
    
    If frmMain.chkEnableRenameDelete.Value = 0 Then 'disabled
        If Open_SubKey(Hive_Key, Sub_Key) = ERROR_SUCCESS Then 'Previously Enabled
            Query_Value REG_BINARY, "Attributes"
            If BD_Value = 536871280 Then
                Open_SubKey Hive_Key, Sub_Key
                Create_Value REG_BINARY, "Attributes", CLng(536871232)
            End If
        End If
    Else
        If Open_SubKey(Hive_Key, Sub_Key) = ERROR_SUCCESS Then
            Query_Value REG_BINARY, "Attributes"
            If BD_Value <> 536871280 Then
                blnFlag = True
            End If
        Else
            blnFlag = True
        End If
        If blnFlag = True Then
            Create_SubKey Hive_Key, Sub_Key, ""
            Open_SubKey Hive_Key, Sub_Key
            Create_Value REG_BINARY, "Attributes", CLng(536871280)
        End If
    End If
        
Exit Sub
ErrHand:
    MsgBox "Error Occurred while Applying Explorer Settings !!" & vbCrLf & "Procedure : Apply_Explorer_Page " & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub XPButton2_Click()
pageNo = pageNo + 1
If pageNo > 1 Then
XPButton3.Enabled = True
XPButton2.Enabled = False
End If
If pageNo = 2 Then
colourFrame1.LabelText = "Control Panel"
frameGen.Visible = False
Picture4.Visible = True
colourFrame2.LabelText = "Display"
Picture1.Visible = False
Picture5.Visible = True
colourFrame3.LabelText = "System"
Picture2.Visible = False
Picture6.Visible = True
colourFrame4.LabelText = "System Cont."
Picture3.Visible = False
Picture7.Visible = True
End If
End Sub

Private Sub XPButton3_Click()
pageNo = pageNo - 1
If pageNo = 1 Then
XPButton3.Enabled = False
XPButton2.Enabled = True
colourFrame1.LabelText = "General"
frameGen.Visible = True
Picture4.Visible = False
colourFrame2.LabelText = "Security"
Picture1.Visible = True
Picture5.Visible = False
colourFrame3.LabelText = "Explorer"
Picture2.Visible = True
Picture6.Visible = False
colourFrame4.LabelText = "Explorer Cont."
Picture3.Visible = True
Picture7.Visible = False
End If
End Sub

Private Sub XPButton4_Click()
frmHideControlPanelApplet.Show 1
End Sub
Private Sub Apply_ControlPanel_Settings()
On Error GoTo ErrHand

    Hive_Key = HKEY_CURRENT_USER
    Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"

'Printers
    'Disable the Addition of Printers
    Write_DWORD Hive_Key, Sub_Key, "NoAddPrinter", chkDisablePrinterAddition.Value
    
    'Disable the Deletion of Printers
    Write_DWORD Hive_Key, Sub_Key, "NoDeletePrinter", chkDisablePrinterDeletion.Value
    
    'Hide the General and Details Printer Pages
    Write_DWORD Hive_Key, Sub_Key, "NoPrinterTabs", chkHidePrinterGeneralDetails.Value

'Passwords
    Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    'Hide the Change Passwords Tab
    Write_DWORD Hive_Key, Sub_Key, "NoPwdPage", frmMain.chkHideChangePassword.Value
    
    'Restrict Access to the Passwords Applet
    Write_DWORD Hive_Key, Sub_Key, "NoSecCPL", frmMain.chkRestrictPassword.Value
    
    'Restrict Access to the User Profiles Page
    Write_DWORD Hive_Key, Sub_Key, "NoProfilePage", frmMain.chkRestrictUserProfile.Value

Exit Sub
ErrHand:
    MsgBox "Error Occurred while Applying Printer Settings !! " & vbCrLf & "Procedure : Apply_ControlPanel_Settings " & vbCrLf & Err.Description, vbCritical
End Sub
Private Sub Apply_Properties_Page()
On Error GoTo Errorh

Hive_Key = HKEY_CURRENT_USER
Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
'Type: REG_DWORD (DWORD Value)
'Value: (0 = disabled, 1 = enabled)

'Display Properties
    'Deny Access to Display Settings
    Write_DWORD Hive_Key, Sub_Key, "NoDispCPL", chkDenyDisplay.Value
    
    'Settings Tab
    Write_DWORD Hive_Key, Sub_Key, "NoDispSettingsPage", chkHideSettings.Value

    'Screen Saver Tab
    Write_DWORD Hive_Key, Sub_Key, "NoDispScrSavPage", chkHideScreenSaver.Value
    
    'Background Tab
    Write_DWORD Hive_Key, Sub_Key, "NoDispBackgroundPage", chkHideBackground.Value
    
    'Appearance Tab
    Write_DWORD Hive_Key, Sub_Key, "NoDispAppearancePage", chkHideAppearance.Value
'System Time
    'Ability to configure System Time
    Write_DWORD HKEY_LOCAL_MACHINE, "Software\" & Software_Name & "\Main", "timeCon", chkTime.Value
    If (Read_DWORD(HKEY_LOCAL_MACHINE, "Software\" & Software_Name & "\Main", "timeCon") = 0) Then
    FileSystem.MkDir (WINSYSDIR & "\timedate")
    FileSystem.FileCopy WINSYSDIR & "\timedate.cpl", WINSYSDIR & "\timedate\timedate.cpl"
    FileSystem.Kill WINSYSDIR & "\timedate.cpl"
    Open_SubKey HKEY_LOCAL_MACHINE, Software_Name & "\Main"
    Create_Value REG_SZ, "timePath", WINSYSDIR & "\timedate\timedate.cpl"
    Else
    FileSystem.FileCopy WINSYSDIR & "\timedate\timedate.cpl", WINSYSDIR & "\timedate.cpl"
    FileSystem.Kill WINSYSDIR & "\timedate\timedate.cpl"
    FileSystem.RmDir WINSYSDIR & "\timedate\"
    Open_SubKey HKEY_LOCAL_MACHINE, Software_Name & "\Main"
    Create_Value REG_SZ, "timePath", WINSYSDIR & "\timedate.cpl"
    End If
    
    'System Properties
    'Device Manager Page
    Write_DWORD Hive_Key, Sub_Key, "NoDevMgrPage", chkHideDeviceManager.Value
    
    'Hardware Profiles Page
    Write_DWORD Hive_Key, Sub_Key, "NoConfigPage", chkHideHardwareProfiles.Value
    
        'File System Button
    Write_DWORD Hive_Key, Sub_Key, "NoFileSysPage", chkHideFileSystem.Value

    'Virtual Memory Button
    Write_DWORD Hive_Key, Sub_Key, "NoVirtMemPage", chkHideVirtualMemory.Value
Exit Sub
Errorh:

End Sub

Private Sub XPButton5_Click()

frmAddAndRemove.Show 1
End Sub
Private Function poov()
PopupMenu try
End Function

Private Sub XPButton6_Click()
frmHelp.Show 1
End Sub
