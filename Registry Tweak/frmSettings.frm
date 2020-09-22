VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "SYSTEMMSCOMCTL.OCX"
Begin VB.Form frmSettings 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   105
      ScaleHeight     =   555
      ScaleWidth      =   615
      TabIndex        =   10
      Top             =   2640
      Width           =   615
   End
   Begin Tweak.TxtBox txtOldPassword 
      Height          =   330
      Left            =   2685
      TabIndex        =   7
      ToolTipText     =   "Type your old password here"
      Top             =   750
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PasswordChar    =   "*"
      Text            =   ""
      Value           =   -1  'True
      BorderColor     =   16744192
   End
   Begin Tweak.XPButton cmdPassOK 
      Height          =   390
      Left            =   2175
      TabIndex        =   5
      ToolTipText     =   "Click this to save changes"
      Top             =   2805
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   688
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
   Begin MSComctlLib.ImageList IMG 
      Left            =   825
      Top             =   2565
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   95
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":0CDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":15B4
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":2406
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":2CE0
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":35BA
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":3E94
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":476E
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":4A88
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":5362
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":5C3C
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":6516
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":6DF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":7AD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":87B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":9490
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":A170
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":AE50
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":BB30
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":C810
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":D4F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":E1D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":EEB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":FB90
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":10870
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":11550
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":12230
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":12F10
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":13BF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":148D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":155B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":16290
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":16F70
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":17C50
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":18930
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":19610
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":19EEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":1ABCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":1B8AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":1C58C
            Key             =   "exclamation"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":1D268
            Key             =   "critical"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":1DF44
            Key             =   "information"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":1EC20
            Key             =   "question"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":1F8FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":201D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":20AB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":2138A
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":21C64
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":2253E
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":22E18
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":236F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":23A0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":242E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":24BC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":2549A
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":25D74
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":2662A
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":26F04
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":277DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":280B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":28992
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":2986C
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":2A746
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":2B620
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":2C4FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":2D3D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":2E2AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":2F188
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":30062
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":30F3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":31E16
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":32CF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":33BCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":34AA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":3597E
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":36858
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":37732
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":3860C
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":394E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":3A3C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":3B29A
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":3C174
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":3D04E
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":3DF28
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":3EE02
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":3FCDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":40BB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":41A90
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":4296A
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":43844
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":4471E
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":455F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":464D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":473AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":48286
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Tweak.XPButton cmdCancel 
      Height          =   390
      Left            =   3705
      TabIndex        =   6
      ToolTipText     =   "Click this to exit change password without saving"
      Top             =   2805
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   688
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
   Begin Tweak.TxtBox txtNewPassword 
      Height          =   330
      Left            =   2685
      TabIndex        =   8
      ToolTipText     =   "Type your New password here"
      Top             =   1245
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PasswordChar    =   "*"
      Text            =   ""
      Value           =   -1  'True
      BorderColor     =   16744192
   End
   Begin Tweak.TxtBox txtConfirmPassword 
      Height          =   330
      Left            =   2685
      TabIndex        =   9
      ToolTipText     =   "Type to confirm your New Password"
      Top             =   1710
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PasswordChar    =   "*"
      Text            =   ""
      Value           =   -1  'True
      BorderColor     =   16744192
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00B99D7F&
      X1              =   210
      X2              =   4920
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Change Password"
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
      Left            =   930
      TabIndex        =   4
      ToolTipText     =   "Change Password Title"
      Top             =   45
      Width           =   3105
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   285
      ToolTipText     =   "Change Password Logo"
      Top             =   45
      Width           =   540
   End
   Begin VB.Label lblConfirmPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "&Confirm New Password"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   870
      TabIndex        =   3
      Top             =   1785
      Width           =   1695
   End
   Begin VB.Label lblNewPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "&New Password"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   870
      TabIndex        =   2
      Top             =   1305
      Width           =   1335
   End
   Begin VB.Label lblOldPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "&Old Password"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   870
      TabIndex        =   1
      Top             =   825
      Width           =   1635
   End
   Begin VB.Image imgPassword 
      Height          =   480
      Left            =   270
      Picture         =   "frmSettings.frx":49160
      ToolTipText     =   "Password Icon"
      Top             =   1230
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Leave Password Field Blank, To Disable Password"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   105
      TabIndex        =   0
      ToolTipText     =   "If you want to disable the password dialog box leave New Password and Confirm New field blank"
      Top             =   2205
      Width           =   4815
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Image1.Picture = IMG.ListImages(7).Picture
makeTrans Me, hwnd, RGB(90, 190, 255)
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Select_On_Focus()
    On Error Resume Next
    With ActiveControl
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub cmdPassOK_Click()
'Confirm Password
txtOldPassword.Text = UCase(txtOldPassword.Text)
If txtConfirmPassword.Text <> txtNewPassword.Text Then
    Msgex "New Password and Confirmed Password Do Not Match !!", "Password", Info
    txtConfirmPassword.SetFocus
    Select_On_Focus
    Exit Sub
End If

If OldPassword_Validate = False Then
    Msgex "Incorrect Password !!", "Password", Info
    txtConfirmPassword.Text = ""
    txtNewPassword.Text = ""
    txtOldPassword.Text = ""
    txtOldPassword.SetFocus
    Select_On_Focus
    Exit Sub
End If

If txtConfirmPassword.Text = "" Then
    If Save_Password(True) = True Then
        Msgex "Password Disabled !!", , Info
    Else
        Msgex "Password Cannot be Updated !!", , Error
    End If
Else
    If Save_Password(False) = True Then
        Msgex "Password Updated Successfully !!", , Info
    Else
        Msgex "Password Cannot be Updated !!", , Error
    End If
End If
txtConfirmPassword.Text = ""
txtNewPassword.Text = ""
txtOldPassword.Text = ""
If blnInitial_Password = True Then
    txtOldPassword.Enabled = True
    blnInitial_Password = False
End If
cmdCancel_Click
End Sub

Private Function Save_Password(NoPass As Boolean) As Boolean
On Error GoTo ErrHand
    
    Dim blnFailed As Boolean
    blnFailed = True
    
    Hive_Key = HKEY_LOCAL_MACHINE
    Sub_Key = "Software\" & Software_Name & "\Main"
    
    If Open_SubKey(Hive_Key, Sub_Key) = ERROR_SUCCESS Then
        If NoPass = True Then
            Delete_Value ("TPass")
            
            Open_SubKey Hive_Key, Sub_Key
            If Query_Value(REG_SZ, "TPass") = ERROR_SUCCESS Then
                blnFailed = False
            End If
        Else
            If Create_Value(REG_SZ, "TPass", Create_Password) <> ERROR_SUCCESS Then
                blnFailed = False
            End If
        End If
    Else
        blnFailed = False
    End If
    Save_Password = blnFailed
    
Exit Function
ErrHand:
    Msgex "Error Occurred while Saving Password !!" & vbCrLf & "Function : Save_Password " & vbCrLf & Err.Description, , Error
End Function

Public Function Create_Password() As String
On Error GoTo ErrHand
    
    Dim sPassword As String
    
    Hive_Key = HKEY_LOCAL_MACHINE
    Sub_Key = "Software\" & Software_Name & "\Main"
    
    sPassword = ""
    If val(UCase(txtConfirmPassword.Text)) = True Then
    txtConfirmPassword.Text = UCase(txtConfirmPassword.Text)
    For Ctr = 1 To Len(txtConfirmPassword.Text)
        sPassword = sPassword + Encrypt_Password(Mid(txtConfirmPassword.Text, Ctr, 1), Ctr)
    Next
    Create_Password = sPassword
End If
Exit Function
ErrHand:
    Msgex "Error Occurred while Retrieving Password !!" & vbCrLf & "Function : Create_Password " & vbCrLf & Err.Description, , Error
End Function

Public Function Encrypt_Password(sChar As String, lPos As Long) As String
On Error GoTo ErrHand
   
        Encrypt_Password = Asc(sChar) + lPos

Exit Function
ErrHand:
    Msgex "Error Occurred while Encrypting Password !!" & vbCrLf & "Function : Encrypt_Password " & vbCrLf & Err.Description, , Error
End Function

Private Sub Form_Activate()
progInit = 1
If blnNoPass = True Then
txtOldPassword.Visible = False
lblOldPassword.Visible = False
txtNewPassword.SetFocus
Else
    txtOldPassword.SetFocus
Select_On_Focus
End If

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
drag Me
End Sub

Private Sub txtConfirmPassword_GotFocus()
Select_On_Focus
End Sub

Private Sub txtConfirmPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdPassOK_Click
End If
End Sub

Private Sub txtNewPassword_GotFocus()
Select_On_Focus
End Sub

Private Sub txtNewPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdPassOK_Click
End If
End Sub

Private Sub txtOldPassword_GotFocus()
Select_On_Focus
End Sub

Private Function OldPassword_Validate() As Boolean
On Error GoTo ErrHand
Init_Password
If blnNoPass = False Then
    If txtOldPassword.Text <> Retrieve_Password Then
        OldPassword_Validate = False
    Else
        OldPassword_Validate = True
    End If
Else
    If txtOldPassword.Text <> "" Then
        OldPassword_Validate = False
    Else
        OldPassword_Validate = True
    End If
End If

Exit Function
ErrHand:
    Msgex "Error Occurred while Validating Password !!" & vbCrLf & "Function : OldPassword_Validate " & vbCrLf & Err.Description, , Error
End Function

Private Sub txtOldPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdPassOK_Click
End If
End Sub
