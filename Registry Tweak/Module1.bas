Attribute VB_Name = "Module1"
Option Explicit
'API Structures
Public Declare Function PlaySoundData Lib "winmm.dll" Alias "PlaySoundA" _
(lpData As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Const EWX_FORCE = 4
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1
Public m_snd() As Byte
Public findTimer As Integer
Public whereTime As String
Public ve As String
Public No_Login
Public newer As Integer
Public demono As Boolean
Public screenResW As Single
Public screenResH As Single
Public progInit As Integer
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_NOTOPMOST = -2
Public Const LB_ITEMFROMPOINT = &H1A9
Public shad As Integer
Public rese As Integer
Public logo As Integer
Public ask As Integer
Const REG_EXPAND_SZ = 2
Private Type SECURITY_ATTRIBUTES
nLength As Long
lpSecurityDescriptor As Long
bInheritHandle As Boolean
End Type
Dim GetErrorMsg As String
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long


Dim hKey As Long, MainKeyHandle As Long
Dim rtn As Long, lBuffer As Long, sBuffer As String
Dim lBufferSize As Long
Dim lDataSize As Long
Dim ByteArray() As Byte

Public Declare Function Setwindowpos Lib "user32" Alias "SetWindowPos" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function RegSetValueExB Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long


' Create Registrykeys-Typs...
Const REG_OPTION_NON_VOLATILE = 0 ' Key is exists on Systemstart

' Registrykeys Securityoptions...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Const KEY_EXECUTE = KEY_READ


'For Registry Handlings
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
 (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" _
 Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
 ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal _
 hKey As Long) As Long

Public Declare Function RegSetValueEx Lib "advapi32.dll" _
 Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName _
 As String, ByVal Reserved As Long, ByVal dwType As Long, _
 lpData As Any, ByVal cbData As Long) As Long

Public Declare Function RegSetValueExLong Lib "advapi32.dll" _
 Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName _
 As String, ByVal Reserved As Long, ByVal dwType As Long, _
 lpData As Long, ByVal cbData As Long) As Long
 
Public Declare Function RegCreateKey Lib "advapi32.dll" _
 Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, _
 phkResult As Long) As Long

Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
 (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
 (ByVal hKey As Long, ByVal lpValueName As String) As Long

Public Declare Function RegReplaceKey Lib "advapi32.dll" Alias "RegReplaceKeyA" _
 (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpNewFile As String, _
 ByVal lpOldFile As String) As Long

Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
 (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, _
  ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, _
  lpftLastWriteTime As FILETIME) As Long
  
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" _
 (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
  ByVal cbName As Long) As Long
  
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
 (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
 lpcbValueName As Long, ByVal lpReserved As Long, ByVal lpType As Long, _
 ByVal lpData As Byte, ByVal lpcbData As Long) As Long

'Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
' (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
'  lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, _
'  lpData As Byte, lpcbData As Long) As Long

Public Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" _
 (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, _
 lpcbValue As Long) As Long

Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
 (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
 lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Public Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" _
 (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, _
  ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, _
  lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, _
  lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, _
  lpftLastWriteTime As FILETIME) As Long

'For Getting Windows Directory
Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias _
 "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'Type Decalaration
Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

'For Getting Windows Version
Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long

Public Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" _
 (lpVersionInformation As OSVERSIONINFO) As Long
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

'For Getting Drive Type
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

'API Constants
Public Const DRIVE_UNKNOWN = 0
Public Const DRIVE_DOES_NOT_EXIST = 1
Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_FIXED = 3
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_CDROM = 5
Public Const DRIVE_RAMDISK = 6

'Constants Definition
'Registry Hives
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Public Const KEY_ALL_ACCESS = &H3F
'Registry Data Types
Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD As Long = 4

'Public Const KEY_ALL_ACCESS = &H3F
'Public Const REG_OPTION_NON_VOLATILE = 0
'Public Const KEY_WRITE = &H20006
Public Const ERROR_SUCCESS = 0&

Public Const Software_Name = "Tweak Machine"

'Custom Variables
Public HRegKey As Long
Public HRegKey2 As Long
Public Sub_Key As String
Public StringBuffer As String
Public RetVal As Variant
Public Hive_Key As Long
Public lngBuffer As Long
Public S_Value As String
Public BD_Value As Long
Public Value_Type As Long
Public New_Key As String
Public WINDIR As String
Public WINSYSDIR As String
Public StartMenu_DIR As String
Public Desktop_DIR As String
Public tFT As FILETIME
Public File_Name As String
Public Ctr As Long
Public FString As String
Public blnNoPass As Boolean
Public blnInitial_Password As Boolean
Public tp As Integer
Public B As Long
Public i As Long
Public c, d, H As Long
Public vf, vg As Integer
Public Type CPanel_Applets
    Name As String
    Description As String
End Type

Public Con_App(26) As CPanel_Applets
Public Const NO_CON_APP = 26
Public Enum eIcon
Alert = 1
Error = 2
Password = 3
Info = 4
Question = 5
Security = 6
demo = 7
End Enum
Public gk As Integer
Public timeup As Integer
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean


Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
    Const CCDEVICENAME = 32
    Const CCFORMNAME = 32
    Const DM_PELSWIDTH = &H80000
    Const DM_PELSHEIGHT = &H100000


Public Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    End Type
    Dim DevM As DEVMODE


'Sub Main
Private Sub Main()
gk = 0
Get_System_DIRS
        findTime
    Init_My_Software

    If gk = 0 Then
    
    If timeup = 0 Then
    getScreenRes
        

       
       
       frmLogin.Show
       Init_Settings
       End If
        'Load all Settings
       Else
       Unload frmLogin
       Unload frmMain
       Unload frmHideDrive
       Unload frmMsg
       Unload frmSettings
       End If
End Sub

'Load all Initial Settings
Private Sub Init_Settings()
Init_Control_Panel_Applet_Desc

End Sub
Public Function Create_SubKey(hKey As Long, sKey As String, NKey As String) As Long
On Error GoTo ErrHand
    Dim sSubKey As String
    
    If Trim(NKey) <> "" Then
        sSubKey = sKey & "\" & NKey
    Else
        sSubKey = sKey
    End If
    Create_SubKey = RegCreateKey(hKey, sSubKey, HRegKey)
    RegCloseKey HRegKey
    
    Exit Function
ErrHand:
    Create_SubKey = 1
    RegCloseKey HRegKey
    Msgex "Cannot Create Sub Key !!" & vbCrLf & Err.Description
End Function

'Function to Create a New Value - REG_SZ / REG_BINARY / REG_DWORD
Public Function Create_Value(vType As Long, RValue As String, RVData As Variant) As Long
On Error GoTo ErrHand
        
    Select Case vType
        Case REG_SZ
            StringBuffer = CStr(RVData)
            Create_Value = RegSetValueEx(HRegKey, RValue, 0&, vType, ByVal StringBuffer, Len(StringBuffer))
        Case REG_BINARY
            Create_Value = RegSetValueExLong(HRegKey, RValue, 0&, vType, RVData, 4&)
        Case REG_DWORD
            Create_Value = RegSetValueExLong(HRegKey, RValue, 0&, vType, RVData, 4&)
    End Select
    RegCloseKey HRegKey
    
    Exit Function
ErrHand:
    Create_Value = 1
    RegCloseKey HRegKey
    Msgex "Cannot Create Value !!" & vbCrLf & Err.Description
End Function

'Function to Open a Key
Public Function Open_SubKey(hKey As Long, sKey As String) As Long
On Error GoTo ErrHand
            
    Open_SubKey = RegOpenKey(hKey, sKey, HRegKey)
    
    Exit Function
ErrHand:
    Open_SubKey = 1
    RegCloseKey HRegKey
    Msgex "Cannot Open Sub Key !!" & vbCrLf & Err.Description
End Function

'Precedure to Query a Value
Public Function Query_Value(vType As Long, RValue As String) As Long
On Error GoTo ErrHand
    
    lngBuffer = 0
    If RegQueryValueEx(HRegKey, RValue, 0&, vType, ByVal 0&, lngBuffer) = ERROR_SUCCESS Then
        Select Case vType
            Case REG_SZ
                lngBuffer = 256
                S_Value = Space$(lngBuffer)
                RetVal = RegQueryValueEx(HRegKey, RValue, 0&, vType, ByVal S_Value, lngBuffer)
                'drop null-terminator
                S_Value = Left(S_Value, lngBuffer - 1)
            Case REG_BINARY
                'MsgBox Hex(22)
                RetVal = RegQueryValueEx(HRegKey, RValue, 0&, vType, BD_Value, 4&) ' 4& = 4-byte word (long integer)
            Case REG_DWORD
                RetVal = RegQueryValueEx(HRegKey, RValue, 0&, vType, BD_Value, 4&) ' 4& = 4-byte word (long integer)
        End Select
    Else
        Query_Value = 1
    End If
    RegCloseKey HRegKey
    
    Exit Function
ErrHand:
    Query_Value = 1
    RegCloseKey HRegKey
    Msgex "Cannot Query Value !!" & vbCrLf & Err.Description
End Function

'Precedure to Delete a Key
Public Function Delete_Key(RKey As String) As Long
On Error GoTo ErrHand

    Delete_Key = RegDeleteKey(HRegKey, RKey)
    RegCloseKey HRegKey
    
    Exit Function
ErrHand:
    Delete_Key = 1
    RegCloseKey HRegKey
    Msgex "Cannot Delete Key !!" & vbCrLf & Err.Description
End Function

'Function to Delete a Value
Public Function Delete_Value(RValue As String) As Long
On Error GoTo ErrHand
    
    Delete_Value = RegDeleteValue(HRegKey, RValue)
    RegCloseKey HRegKey
    
    Exit Function
ErrHand:
    Delete_Value = 1
    RegCloseKey HRegKey
    Msgex "Cannot Delete Value !!" & vbCrLf & Err.Description
End Function
Public Function Read_DWORD(hKey As Long, sKey As String, RValue As String) As Integer
On Error GoTo ErrHand
    
    Re_Init_Val
    If Open_SubKey(hKey, sKey) = ERROR_SUCCESS Then
        Query_Value REG_DWORD, RValue
    End If
    Read_DWORD = BD_Value
    
    Exit Function
ErrHand:
    Msgex "Error Occurred while Reading Values !!" & vbCrLf & "Procedure : Read_DWORD " & vbCrLf & Err.Description
    
End Function

Public Sub Write_DWORD(hKey As Long, sKey As String, RValue As String, RData As Long)
On Error GoTo ErrHand

    If Read_DWORD(hKey, sKey, RValue) = 1 And RData = 0 Then
        Open_SubKey hKey, sKey
        Delete_Value RValue
        Exit Sub
    End If
    
    If RData = 1 Then
        Re_Init_Val
        If Open_SubKey(hKey, sKey) <> ERROR_SUCCESS Then
            Create_SubKey hKey, sKey, ""
            Open_SubKey hKey, sKey
        End If
        
        Create_Value REG_DWORD, RValue, RData
    End If
    
    Exit Sub
ErrHand:
    Msgex "Error Occurred while Writing Values !!" & vbCrLf & "Procedure : Write_DWORD " & vbCrLf & Err.Description
End Sub
Public Sub Read_Form()

   
    'Display the last Username logon
    Hive_Key = HKEY_LOCAL_MACHINE
    Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Winlogon"
    Open_SubKey Hive_Key, Sub_Key
    Query_Value REG_SZ, "DontDisplayLastUserName"
    frmMain.chkUser.Value = S_Value
Hive_Key = HKEY_CURRENT_USER
Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
'Type: REG_DWORD (DWORD Value)
'Value: (0 = disabled, 1 = enabled)

'General
    'Disable Network Control Panel
    frmMain.chkDisableNetwork.Value = Read_DWORD(Hive_Key, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoNetSetup")
    
    'Hide Control Panel on Start Menu
    frmMain.chkHideControlPanel.Value = Read_DWORD(Hive_Key, Sub_Key, "NoSetFolders")
    
    'Hide Folder Options
    frmMain.chkHideFolderOptions.Value = Read_DWORD(Hive_Key, Sub_Key, "NoFolderOptions")

        
Exit Sub
ErrHand:
    Msgex "Error Occurred while Reading Form Settings !! " & vbCrLf & "Procedure : Read_Security_Page " & vbCrLf & Err.Description
End Sub
Public Sub Init_Password()
On Error GoTo ErrHand
    
    blnNoPass = False
    
    Hive_Key = HKEY_LOCAL_MACHINE
    Sub_Key = "Software\" & Software_Name & "\Main"
    
    If Open_SubKey(Hive_Key, Sub_Key) = ERROR_SUCCESS Then
        If Query_Value(REG_SZ, "TPass") = ERROR_SUCCESS Then
            If Trim(S_Value) = "" Then
                blnNoPass = True
            End If
        Else
            blnNoPass = True
        End If
    Else
        blnNoPass = True
    End If
    
Exit Sub

ErrHand:
    Msgex "Error Occurred while Initializing Password !!" & vbCrLf & "Procedure : Init_Password " & vbCrLf & Err.Description
    
End Sub
Public Function Retrieve_Password() As String
On Error GoTo ErrHand

    Dim sTPass As String
    Dim sPassword As String
    
    Hive_Key = HKEY_LOCAL_MACHINE
    Sub_Key = "Software\" & Software_Name & "\Main"
    
    sPassword = ""
    sTPass = ""
    
    Open_SubKey Hive_Key, Sub_Key
    Query_Value REG_SZ, "TPass"
    sTPass = S_Value
    
    For Ctr = 1 To Len(sTPass) Step 2
        sPassword = sPassword + Decrypt_Password(Mid(sTPass, Ctr, 2), Ctr)
    Next
    Retrieve_Password = sPassword

Exit Function
ErrHand:
    Msgex "Error Occurred while Retrieving Password !!" & vbCrLf & "Function : Retrieve_Password " & vbCrLf & Err.Description
End Function

Public Function Decrypt_Password(sChar As String, lPos As Long) As String
Dim poo As Integer
On Error GoTo ErrHand
       poo = (lPos - 1) / 2
    Decrypt_Password = Chr(sChar - poo - 1)

Exit Function
ErrHand:
    Msgex "Error Occurred while Decrypting Password !!" & vbCrLf & "Function : Decrypt_Password " & vbCrLf & Err.Description
End Function

Public Function SetKeyValue(lPredefinedKey As Long, sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)
' Description:
'   This Function will set the data field of a value
'
' Syntax:
'   QueryValue Location, KeyName, ValueName, ValueSetting, ValueType
'
'   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'   , HKEY_USERS
'
'   KeyName is the key that the value is under (example: "Key1\SubKey1")
'
'   ValueName is the name of the value you want create, or set the value of (example: "ValueTest")
'
'   ValueSetting is what you want the value to equal
'
'   ValueType must equal either REG_SZ (a string) Or REG_DWORD (an integer)

       Dim lRetVal As Long         'result of the SetValueEx function
       Dim hKey As Long         'handle of open key

       'open the specified key

       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
       RegCloseKey (hKey)

End Function
Public Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String

    Select Case lType
        Case REG_SZ
            sValue = vValue
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
        End Select

End Function
Public Function Msgex(mess As String, Optional title As String, Optional icon As eIcon)

frmMsg.Label3.Caption = mess
frmMsg.Label2.Caption = title
If icon = Alert Then
frmMsg.ImgPic.Picture = frmSettings.IMG.ListImages(40).Picture
End If
If icon = Error Then
frmMsg.ImgPic.Picture = frmSettings.IMG.ListImages(41).Picture
End If
If icon = Info Then
frmMsg.ImgPic.Picture = frmSettings.IMG.ListImages(42).Picture
End If
If icon = Password Then
frmMsg.ImgPic.Picture = frmSettings.IMG.ListImages(7).Picture
End If
If icon = Question Then
frmMsg.ImgPic.Picture = frmSettings.IMG.ListImages(43).Picture
End If
If icon = Security Then
frmMsg.ImgPic.Picture = frmSettings.IMG.ListImages(59).Picture
End If
If icon = demo Then
frmMsg.ImgPic.Picture = frmSettings.IMG.ListImages(1).Picture
End If
frmMsg.Show 1

End Function
Public Sub Re_Init_Val()
    BD_Value = 0
    S_Value = ""
    lngBuffer = 0
End Sub
Public Sub Read_General_Page()
On Error GoTo ErrHand

    Hive_Key = HKEY_CURRENT_USER
    Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"

'General
    'Disable Saving of Windows Settings at Shut Down or Restart
    frmMain.chkDisableSave.Value = Read_DWORD(Hive_Key, Sub_Key, "NoSaveSettings")

    'Disable Network Control Panel
    frmMain.chkDisableNetwork.Value = Read_DWORD(Hive_Key, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoNetSetup")
    
    'Hide Control Panel on Start Menu
    frmMain.chkHideControlPanel.Value = Read_DWORD(Hive_Key, Sub_Key, "NoSetFolders")
    
    'Hide Folder Options
    frmMain.chkHideFolderOptions.Value = Read_DWORD(Hive_Key, Sub_Key, "NoFolderOptions")
    
    'Disable the New Menu Item in Context Menu
    Read_Disable_NewMenu
        
            'Remove the 'Shortcut to...' Prefix on Shortcuts
    Hive_Key = HKEY_CURRENT_USER
    Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Explorer"
    
    Open_SubKey Hive_Key, Sub_Key
    Query_Value REG_BINARY, "link"
    If BD_Value = 0 Then
        frmMain.chkRemoveShortCutTo.Value = 1
    Else
        frmMain.chkRemoveShortCutTo.Value = 0
    End If
    
    'Easily Use Notepad to Open a File
    Hive_Key = HKEY_CLASSES_ROOT
    Sub_Key = "*\Shell\Open\command"
    
    If Open_SubKey(Hive_Key, Sub_Key) = ERROR_SUCCESS Then
        Query_Value REG_SZ, ""
        If LCase$(S_Value) = "notepad.exe %1" Then
            frmMain.chkUseNotepad.Value = 1
        Else
            frmMain.chkUseNotepad.Value = 0
        End If
    Else
        frmMain.chkUseNotepad.Value = 0
    End If
    
        'Disable Menu Bars and the Start Button
    Hive_Key = HKEY_CLASSES_ROOT
    Sub_Key = "CLSID\{-5b4dae26-b807-11d0-9815-00c04fd91972}"
        
    If RegOpenKey(Hive_Key, Sub_Key, HRegKey) = ERROR_SUCCESS Then
        frmMain.chkDisableStart.Value = 1
        RegCloseKey HRegKey
    Else
        frmMain.chkDisableStart.Value = 0
    End If
          
    Hive_Key = HKEY_CURRENT_USER
    Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    'Add an Expanding Control Panel to Start Menu
    Dim sCPanelExt As String
    sCPanelExt = "{21EC2020-3AEA-1069-A2DD-08002B30309D}"
    
    If Dir(StartMenu_DIR & "\*." & sCPanelExt, vbDirectory Or vbHidden Or vbNormal) <> "" Then
        frmMain.chkExpandingControlPanel.Value = 1
    Else
        frmMain.chkExpandingControlPanel.Value = 0
    End If
    
    'Disable Taskbar Context Menus
    Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    frmMain.chkDisableTaskbarConMenu.Value = Read_DWORD(Hive_Key, Sub_Key, "NoTrayContextMenu")

    
'Display the last Username logon
    frmMain.chkUser.Value = Read_DWORD(HKEY_LOCAL_MACHINE, "Network\Logon", "DontShowLastUser")

Exit Sub
ErrHand:
    Msgex "Error Occurred while Reading General Information Page !! " & vbCrLf & "Procedure : Read_General_Page " & vbCrLf & Err.Description
End Sub
Public Sub Read_Security_Page()
On Error GoTo ErrHand
Hive_Key = HKEY_CURRENT_USER
Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
'Type: REG_DWORD (DWORD Value)
'Value: (0 = disabled, 1 = enabled)

'Desktop
    Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    'Hide All Items on the Desktop
    frmMain.chkHideAllDesktop.Value = Read_DWORD(Hive_Key, Sub_Key, "NoDesktop")
    
    'Hide the Internet Explorer Icon
    frmMain.chkHideInternetExplorer.Value = Read_DWORD(Hive_Key, Sub_Key, "NoInternetIcon")

    'Disable the Ability to Right Click on the Desktop && Explorer
    frmMain.chkDisableRightClick.Value = Read_DWORD(Hive_Key, Sub_Key, "NoViewContextMenu")
'Registry
    Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    'Disable Registry Editing Tools
    frmMain.chkDisableRegistry.Value = Read_DWORD(Hive_Key, Sub_Key, "DisableRegistryTools")
'History
    'Disable Recent Documents History
    frmMain.chkDisableRecDoc.Value = Read_DWORD(Hive_Key, Sub_Key, "NoRecentDocsHistory")
Exit Sub
ErrHand:
    Msgex "Error Occurred while Reading Security Settings !! " & vbCrLf & "Procedure : Read_Security_Page " & vbCrLf & Err.Description
End Sub
Private Sub Read_Disable_NewMenu()
On Error GoTo ErrHand
    Dim hKey As Long
    Dim sKey As String
    hKey = HKEY_CLASSES_ROOT
    sKey = "CLSID\{-D969A300-E7FF-11d0-A93B-00A0C90F2719}"
    
    If RegOpenKey(hKey, sKey, HRegKey) = ERROR_SUCCESS Then
        frmMain.chkDisableNewMenu.Value = 1
        RegCloseKey HRegKey
    Else
        frmMain.chkDisableNewMenu.Value = 0
    End If
    
Exit Sub
ErrHand:
    Msgex "Error Occurred while Reading New Menu Item !! " & vbCrLf & "Procedure : Read_Disable_NewMenu " & vbCrLf & Err.Description
End Sub
'Initialize My Software
Private Sub Init_My_Software()


    Hive_Key = HKEY_LOCAL_MACHINE
    Sub_Key = "Software\" & Software_Name
    
    If Open_SubKey(Hive_Key, Sub_Key) <> ERROR_SUCCESS Then
gk = 1
'FileCopy App.Path & "Mscomctl.ocx", WINSYSDIR & "Mscomctl.ocx"
'FileCopy App.Path & "Mscomctl.dep", WINSYSDIR & "Mscomctl.dep"
'FileCopy App.Path & "Mscomctl.oca", WINSYSDIR & "Mscomctl.oca"
'FileCopy App.Path & "Mscomctl.srg", WINSYSDIR & "Mscomctl.srg"
'Shell "regsvr32 /s " & WINSYSDIR & "Mscomctl.ocx"
        'Create Password
        blnInitial_Password = True
        Create_SubKey Hive_Key, Sub_Key, "Main"
        Open_SubKey Hive_Key, Sub_Key & "\Main"
        
        Create_Value REG_SZ, "Software Name", Software_Name
        
        Open_SubKey Hive_Key, Sub_Key & "\Main"
        Create_Value REG_SZ, "Software Path", App.Path & "\" & App.EXEName
        
                
        Open_SubKey Hive_Key, Sub_Key & "\Main"
        Create_Value REG_DWORD, "timeCon", findTimer
        
        Open_SubKey Hive_Key, Sub_Key & "\Main"
        Create_Value REG_SZ, "timePath", whereTime

        Open_SubKey Hive_Key, Sub_Key & "\Main"
        Create_Value REG_SZ, "Splash", 1

        Open_SubKey Hive_Key, Sub_Key & "\Main"

        Create_Value REG_SZ, "Installed Date", Date & " " & Time
             
        End
       
    End If
    
    Hive_Key = HKEY_LOCAL_MACHINE
    Sub_Key = "Software\Microsoft\Windows\CurrentVersion\App Paths"
    
    If Open_SubKey(Hive_Key, Sub_Key & "\Tweak.exe") <> ERROR_SUCCESS Then
        Create_SubKey Hive_Key, Sub_Key, "Tweak.exe"
        Open_SubKey Hive_Key, Sub_Key & "\Tweak.exe"
        Create_Value REG_SZ, "", WINDIR & "\Tweak.exe"
    Else
        Query_Value REG_SZ, ""
        If S_Value = "" Then
            Open_SubKey Hive_Key, Sub_Key & "\Tweak.exe"
            Create_Value REG_SZ, "", WINDIR & "\Tweak.exe"
        End If
    End If

Exit Sub
ErrHand:
    If Err.Number = 58 Then
        Kill (Desktop_DIR & "\Shortcut to Tweak.exe.lnk")
    Else
        Msgex "Error Occurred while Initializing Software !!" & vbCrLf & "Procedure : Init_My_Software " & vbCrLf & Err.Description, , Error
    End If
End Sub
Public Function makeTrans(frm As Form, hwnd As String, col As OLE_COLOR)
On Error GoTo hand
  Dim Add As Long
  Dim Sum As Long

  Dim X As Single
  Dim Y As Single

    X = frm.Width / Screen.TwipsPerPixelX   'Registers the size of the
    Y = frm.Height / Screen.TwipsPerPixelY  'form in pixels
frm.AutoRedraw = True
    Sum = CreateRectRgn(5, 0, X - 5, 1)
    CombineRgn Sum, Sum, CreateRectRgn(3, 1, X - 3, 2), 2
    CombineRgn Sum, Sum, CreateRectRgn(2, 2, X - 2, 3), 2
    CombineRgn Sum, Sum, CreateRectRgn(1, 3, X - 1, 4), 2
    CombineRgn Sum, Sum, CreateRectRgn(1, 4, X - 1, 5), 2
    CombineRgn Sum, Sum, CreateRectRgn(0, 5, X, Y), 2
    SetWindowRgn frm.hwnd, Sum, True   'Sets corners transparent
frm.Line (frm.Width - 15, 0)-(frm.Width - 15, frm.Height - 1), col
frm.Line (0, 0)-(0, frm.Height - 15), col
frm.Line (0, frm.Height - 15)-(frm.Width - 15, frm.Height - 15), col
frm.Line (0, 0)-(frm.Width, 0), col
frm.Line (frm.Width - 30, 0)-(frm.Width - 30, 75), col
frm.Line (frm.Width - 45, 0)-(frm.Width - 45, 45), col
frm.Line (frm.Width - 75, 15)-(frm.Width - 45, 15), col
frm.Line (15, 15)-(75, 15), col
frm.Line (15, 0)-(15, 75), col
frm.Line (30, 0)-(30, 45), col
Exit Function
hand:
End Function

Public Sub Read_Explorer_Page()
On Error GoTo ErrHand
'Folder
    'Add Command Prompt Option to Every Folder
    Hive_Key = HKEY_CLASSES_ROOT
    Sub_Key = "Directory\shell\Command\Command"
    Dim sComm As String
        
    sComm = "command.com /k cd " & """" & "%1" & """" 'command.com /k cd "%1"
    
    If Open_SubKey(Hive_Key, Sub_Key) = ERROR_SUCCESS Then
        Query_Value REG_SZ, ""
        If LCase$(S_Value) = sComm Then
            frmMain.chkAddCommandPrompt.Value = 1
        Else
            frmMain.chkAddCommandPrompt.Value = 0
        End If
    Else
        frmMain.chkAddCommandPrompt.Value = 0
    End If
    
    'Add a Menu Option to Copy Folders
    Hive_Key = HKEY_CLASSES_ROOT
    Sub_Key = "Directory\shellex\ContextMenuHandlers\Copy to Folder"
    
    If Open_SubKey(Hive_Key, Sub_Key) = ERROR_SUCCESS Then
        Query_Value REG_SZ, ""
        If S_Value = "{C2FBB630-2971-11d1-A18C-00C04FD75D13}" Then
            frmMain.chkAddCopyFolder.Value = 1
        Else
            frmMain.chkAddCopyFolder.Value = 0
        End If
    Else
        frmMain.chkAddCopyFolder.Value = 0
    End If
    
    'Add a Menu Option to Move Folders
    Hive_Key = HKEY_CLASSES_ROOT
    Sub_Key = "Directory\shellex\ContextMenuHandlers\Move to Folder"
    
    If Open_SubKey(Hive_Key, Sub_Key) = ERROR_SUCCESS Then
        Query_Value REG_SZ, ""
        If S_Value = "{C2FBB631-2971-11d1-A18C-00C04FD75D13}" Then
            frmMain.chkAddMoveFolder.Value = 1
        Else
            frmMain.chkAddMoveFolder.Value = 0
        End If
    Else
        frmMain.chkAddMoveFolder.Value = 0
    End If
    
    'Enable Recycle Bin Rename and Delete
    Hive_Key = HKEY_CLASSES_ROOT
    Sub_Key = "CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\ShellFolder"
    
    If Open_SubKey(Hive_Key, Sub_Key) = ERROR_SUCCESS Then
        Query_Value REG_BINARY, "Attributes"
        If BD_Value = 536871280 Then
            frmMain.chkEnableRenameDelete.Value = 1
        Else
            frmMain.chkEnableRenameDelete.Value = 0
        End If
    Else
        frmMain.chkEnableRenameDelete.Value = 0
    End If
    
Exit Sub
ErrHand:
    Msgex "Error Occurred while Reading Explorer Settings !! " & vbCrLf & "Procedure : Read_Explorer_Page " & vbCrLf & Err.Description, "Error: " & Err.Number, Error
End Sub
Public Sub Get_System_DIRS()
'Get_System_DIRS
On Error GoTo ErrHand
    
    Dim worked As Long
    
    WINDIR = String$(144, 0)
    worked = GetWindowsDirectory(WINDIR, Len(WINDIR))
    
    If worked = 0 Then
        Msgex "Cannot Get WINDOWS Directory !!", "Directory", Error
        GoTo ErrHand
    Else
        WINDIR = Left(WINDIR, worked)
    End If
    If WINDIR = "C:\WINNT" Then
    WINSYSDIR = WINDIR & "\SYSTEM32"
    Else
    If GetWindowsVersion = "Windows XP" Then
    WINSYSDIR = WINDIR & "\SYSTEM32"
    Else
    WINSYSDIR = WINDIR & "\SYSTEM"
    End If
    End If
    
    
    Hive_Key = HKEY_CURRENT_USER
    Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    
    Open_SubKey Hive_Key, Sub_Key
    Query_Value REG_SZ, "Start Menu"
    StartMenu_DIR = S_Value
    
    If Trim(StartMenu_DIR) = "" Then
        StartMenu_DIR = WINDIR & "\Start Menu"
    End If
    
    Open_SubKey Hive_Key, Sub_Key
    Query_Value REG_SZ, "Desktop"
    Desktop_DIR = S_Value
    
    If Trim(Desktop_DIR) = "" Then
        Desktop_DIR = WINDIR & "\Desktop"
    End If
    
    Exit Sub
ErrHand:
    Msgex "Error Occurred while Searching Windows Directory !!" & "Software Cannot be Loaded !!" & vbCrLf & "Function : Get_System_DIRS " & vbCrLf & Err.Description, "Error: " & Err.Number, Error
    End
End Sub
Public Sub Read_ControlPanel_Settings()
On Error GoTo ErrHand

    Hive_Key = HKEY_CURRENT_USER
    Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"

'Printers
    'Disable the Addition of Printers
    frmMain.chkDisablePrinterAddition.Value = Read_DWORD(Hive_Key, Sub_Key, "NoAddPrinter")
    
    'Disable the Deletion of Printers
    frmMain.chkDisablePrinterDeletion.Value = Read_DWORD(Hive_Key, Sub_Key, "NoDeletePrinter")
    
    'Hide the General and Details Printer Pages
    frmMain.chkHidePrinterGeneralDetails.Value = Read_DWORD(Hive_Key, Sub_Key, "NoPrinterTabs")

'Passwords
    Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    'Hide the Change Passwords Tab
    frmMain.chkHideChangePassword.Value = Read_DWORD(Hive_Key, Sub_Key, "NoPwdPage")
    
    'Restrict Access to the Passwords Applet
    frmMain.chkRestrictPassword.Value = Read_DWORD(Hive_Key, Sub_Key, "NoSecCPL")
    
    'Restrict Access to the User Profiles Page
    frmMain.chkRestrictUserProfile.Value = Read_DWORD(Hive_Key, Sub_Key, "NoProfilePage")

Exit Sub
ErrHand:
    Msgex "Error Occurred while Reading Printer Settings Page !! " & vbCrLf & "Procedure : Read_ControlPanel_Settings " & vbCrLf & Err.Description, "Error: " & Err.Number, Error
End Sub
Public Function GetWindowsVersion() As String
On Error GoTo ErrHand
    Dim TheOS As OSVERSIONINFO

    TheOS.dwOSVersionInfoSize = Len(TheOS)
    
    GetVersionEx TheOS 'Get Operating System
    
    Select Case TheOS.dwPlatformId
        Case VER_PLATFORM_WIN32_WINDOWS
            If TheOS.dwMinorVersion >= 10 Then
                GetWindowsVersion = "Windows 98"
            Else
                GetWindowsVersion = "Windows 95"
            End If
        Case VER_PLATFORM_WIN32_NT
        If Not TheOS.dwBuildNumber = 2600 Then
        GetWindowsVersion = "Windows NT"
        Else
        GetWindowsVersion = "Windows XP"
        End If
    End Select
   
    '& " (Build " & TheOS.dwBuildNumber & strCSDVersion & ")"
    
Exit Function

ErrHand:
    Msgex "Error Occurred while Getting Windows Version !!" & vbCrLf & "Function : GetWindowsVersion " & vbCrLf & Err.Description, "Error: " & Err.Number, Error
End Function
Public Sub Read_Properties_Page()
On Error GoTo ErrHand

Hive_Key = HKEY_CURRENT_USER
Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
'Type: REG_DWORD (DWORD Value)
'Value: (0 = disabled, 1 = enabled)

'Display Properties
    'Deny Access to Display Settings
    frmMain.chkDenyDisplay.Value = Read_DWORD(Hive_Key, Sub_Key, "NoDispCPL")
    
    'Settings Tab
    frmMain.chkHideSettings.Value = Read_DWORD(Hive_Key, Sub_Key, "NoDispSettingsPage")
    
    'Screen Saver Tab
    frmMain.chkHideScreenSaver.Value = Read_DWORD(Hive_Key, Sub_Key, "NoDispScrSavPage")
    
    'Background Tab
    frmMain.chkHideBackground.Value = Read_DWORD(Hive_Key, Sub_Key, "NoDispBackgroundPage")
    
    'Appearance Tab
    frmMain.chkHideAppearance.Value = Read_DWORD(Hive_Key, Sub_Key, "NoDispAppearancePage")
'System Time
    'Ability to configure System Time
    frmMain.chkTime.Value = Read_DWORD(HKEY_LOCAL_MACHINE, "Software\" & Software_Name & "\Main", "timeCon")
    
    'System Properties
    'Device Manager Page
    frmMain.chkHideDeviceManager.Value = Read_DWORD(Hive_Key, Sub_Key, "NoDevMgrPage")
    
    'Hardware Profiles Page
    frmMain.chkHideHardwareProfiles.Value = Read_DWORD(Hive_Key, Sub_Key, "NoConfigPage")
    
        'File System Button
    frmMain.chkHideFileSystem.Value = Read_DWORD(Hive_Key, Sub_Key, "NoFileSysPage")
    
    'Virtual Memory Button
    frmMain.chkHideVirtualMemory.Value = Read_DWORD(Hive_Key, Sub_Key, "NoVirtMemPage")

Exit Sub
ErrHand:
    Msgex "Error Occurred while Retrieving Display Settings !!" & vbCrLf & "Procedure : Read_Properties_Page " & vbCrLf & Err.Description, "Error: " & Err.Number, Error
End Sub
Public Function findTime() As Integer
On Error GoTo notime
If FileSystem.FileLen(WINSYSDIR & "\timedate.cpl") > 0 Then
findTimer = 1
whereTime = WINSYSDIR & "\timedate.cpl"
End If

Exit Function
notime:
findTime2

End Function
Public Function findTime2()
On Error GoTo abnotime
If FileSystem.FileLen(WINSYSDIR & "\timedate\timedate.cpl") > 0 Then
findTimer = 1
whereTime = WINSYSDIR & "\timedate\timedate.cpl"
End If

Exit Function
abnotime:
Msgex "Could not locate timedate.cpl" & vbCrLf & "Please put the timedate.cpl in " & WINSYSDIR & " folder to make the time and date function to work.", "File not Found", Error
frmMain.chkTime.Enabled = False
findTimer = 0
whereTime = ""
End Function

Public Sub ChangeRes(iWidth As Single, iHeight As Single)

    Dim a As Boolean
    Dim i&
    i = 0


    Do
        a = EnumDisplaySettings(0&, i&, DevM)
        i = i + 1
    Loop Until (a = False)

    Dim B&
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
    DevM.dmPelsWidth = iWidth
    DevM.dmPelsHeight = iHeight
    B = ChangeDisplaySettings(DevM, 0)
End Sub
Public Function getScreenRes()
screenResH = Screen.Height / 15
screenResW = Screen.Width / 15
End Function
Public Function SetWinPos(iPos As Integer, lHWnd As Long) As Boolean
Dim lwinpos As Long
iPos = 1

Select Case iPos
    Case 1
        lwinpos = HWND_TOPMOST
    End Select
If Setwindowpos(lHWnd, lwinpos, 0, 0, 0, 0, SWP_NOMOVE _
                                    + SWP_NOSIZE) Then
SetWinPos = True
End If
End Function
Public Function val(tex As String) As Boolean
Dim hj As Integer
If Len(tex) < 8 Then
For c = 1 To Len(tex)
If (Asc(Mid(tex, c, 1)) < 91 And Asc(Mid(tex, c, 1)) > 64) Or (Asc(Mid(tex, c, 1)) > 47 And Asc(Mid(tex, c, 1)) < 58) Then
hj = hj + 1
End If
Next
Else
val = False
End If
If hj = Len(tex) Then
val = True
End If
End Function
Public Sub Read_AddRemove_Programs()
On Error GoTo ErrHand

Dim lKeyNum As Long, lIndex As Long
Dim sKeyName As String, lKeyNameLen As Long
Dim DontAdd As Boolean

Hive_Key = HKEY_LOCAL_MACHINE
Sub_Key = "Software\Microsoft\Windows\CurrentVersion\Uninstall"

Re_Init_Val
    
    
    Open_SubKey Hive_Key, Sub_Key 'Opening Uninstall Subkey
   
    'Retrieving the number of Subkeys Present in Uninstall
    Call RegQueryInfoKey(HRegKey, vbNullString, 0&, 0&, lKeyNum, 0&, 0&, 0&, 0&, 0&, 0&, tFT)
    
    frmAddAndRemove.lstAddRemove.Clear
    
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
            If RegQueryValueEx(HRegKey2, "DisplayName", 0&, REG_SZ, ByVal 0&, lngBuffer) <> ERROR_SUCCESS Then
                DontAdd = True
            ElseIf RegQueryValueEx(HRegKey2, "UninstallString", 0&, REG_SZ, ByVal 0&, lngBuffer) <> ERROR_SUCCESS Then
                DontAdd = True
            End If
                                    
            'If Exits
            If DontAdd = False Then
                lngBuffer = 256
                sKeyName = Space$(lngBuffer)
                
                'Retrieve the Software Name ('DisplayName')
                RetVal = RegQueryValueEx(HRegKey2, "DisplayName", 0&, REG_SZ, ByVal sKeyName, lngBuffer)
                sKeyName = Left(sKeyName, lngBuffer - 1) 'drop null-terminator
                    
                frmAddAndRemove.lstAddRemove.AddItem (sKeyName) 'Add to List in Add/Remove
            End If
            
            DontAdd = False
            RegCloseKey HRegKey2 'Close Registry Handle
            
        Next lIndex
    End If
    RegCloseKey HRegKey

Exit Sub
ErrHand:
    Msgex "Error Occurred while Reading Add/Remove Program List !! " & vbCrLf & "Procedure : Read_AddRemove_Programs " & vbCrLf & Err.Description, "Error", Error
End Sub

Public Function getDemo()
Dim iq As String
Dim d, f, g As Integer
    Hive_Key = HKEY_LOCAL_MACHINE
    Sub_Key = "Software\" & Software_Name & "\Main"
    Open_SubKey Hive_Key, Sub_Key
    Query_Value REG_SZ, "Installed Date"
    iq = S_Value
    f = Left(iq, 2)
    g = Right(iq, 2)
    If f < 10 Then
    f = f + 10
    End If
    If g < 10 Then
    g = g + 10
    End If
    
    d = f & g + 5
    getDemo = d - 30
End Function

Public Function getVeri() As String
Dim vers As String
vers = App.Major & App.Minor & App.Revision
getVeri = getDemo * -1 + 423 + vers
End Function

Public Function PlaySound(ByVal SndID As Long) As Long
      Const flags = &H1 Or &H4
      m_snd = LoadResData(SndID, "WAVE")
      PlaySoundData m_snd(0), 0, flags
End Function

Public Sub Init_Control_Panel_Applet_Desc()
On Error GoTo ErrHand
    
    Con_App(1).Name = "Access.cpl"
    Con_App(1).Description = "Accessibility Options"
    
    Con_App(2).Name = "Appwiz.cpl"
    Con_App(2).Description = "Add/Remove Programs"
    
    Con_App(3).Name = "Audiohq.cpl"
    Con_App(3).Description = "AudioHQ"
    
    Con_App(4).Name = "CTDetect.cpl"
    Con_App(4).Description = "Disk Detector"
    
    Con_App(5).Name = "Desk.cpl"
    Con_App(5).Description = "Display"
    
    Con_App(6).Name = "Inetcpl.cpl"
    Con_App(6).Description = "Internet Options"
    
    Con_App(7).Name = "Input98.cpl"
    Con_App(7).Description = "Text Services"
                
    Con_App(8).Name = "Intl.cpl"
    Con_App(8).Description = "Regional Settings"
        
    Con_App(9).Name = "Joy.cpl"
    Con_App(9).Description = "Gaming Options"
        
    Con_App(10).Name = "Main.cpl"
    Con_App(10).Description = "Mouse"
  
    Con_App(11).Name = "Mmsys.cpl"
    Con_App(11).Description = "Multimedia"
    
    Con_App(12).Name = "Modem.cpl"
    Con_App(12).Description = "Modems"
    
    Con_App(13).Name = "Netcpl.cpl"
    Con_App(13).Description = "Network"
    
    Con_App(14).Name = "Odbccp32.cpl"
    Con_App(14).Description = "ODBC Data Sources [32bit]"
    
    Con_App(15).Name = "Password.cpl"
    Con_App(15).Description = "Passwords"
    
    Con_App(16).Name = "Powercfg.cpl"
    Con_App(16).Description = "Power Management"
    
    Con_App(17).Name = "Sticpl.cpl"
    Con_App(17).Description = "Scanners and Cameras"
    
    Con_App(18).Name = "Sysdm.cpl"
    Con_App(18).Description = "System"
    
    Con_App(19).Name = "Telephon.cpl"
    Con_App(19).Description = "Telephony"
    
    Con_App(20).Name = "Themes.cpl"
    Con_App(20).Description = "Desktop Themes"
    
    Con_App(21).Name = "Timedate.cpl"
    Con_App(21).Description = "Date/Time"
        
    Con_App(22).Name = "tweakmanager.cpl"
    Con_App(22).Description = "Winguides Tweak Manager"
    
    Con_App(23).Name = "Tweakui.cpl"
    Con_App(23).Description = "Tweak UI"
    
    Con_App(24).Name = "XQXSetup.cpl"
    Con_App(24).Description = "Xteq X-Setup"
    
    Con_App(25).Name = "Qtw32.cpl"
    Con_App(25).Description = "Quick Time 32"
    
    Con_App(26).Name = "Adobe Gamma.cpl"
    Con_App(26).Description = "Adobe Gamma"
        
Exit Sub
ErrHand:
    Msgex "Error Occurred while Initializing Control Panel Applets !!" & vbCrLf & "Procedure : Init_Control_Panel_Applet_Desc " & vbCrLf & Err.Description, "Error: " & Err.Number, Error
End Sub
