Attribute VB_Name = "Mod_MastEr_by_CyborgX"

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Author: mostly CyborgX, some functions or subs were created by someone
'  else--but not many.
' CyborgX@Dangerous-Minds.Com
' Please, don't be a lamer, give credit where credit is due.
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '

Global Const chrq$ = """"

Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
    Public Const TH32CS_SNAPPROCESS As Long = 2&
    Public Const MAX_PATH As Integer = 260
    Public Type PROCESSENTRY32

dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Declare Function EnumWindows Lib "user32" (ByVal wndenmprc As Long, ByVal lParam As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long
    Public Const INTERNET_CONNECTION_MODEM As Long = &H1
    Public Const INTERNET_CONNECTION_LAN As Long = &H2
    Public Const INTERNET_CONNECTION_PROXY As Long = &H4
    Public Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8
    Public Const INTERNET_RAS_INSTALLED As Long = &H10
    Public Const INTERNET_CONNECTION_OFFLINE As Long = &H20
    Public Const INTERNET_CONNECTION_CONFIGURED As Long = &H40

Public Const WM_CLOSE = &H10

Private Target As String

Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Sub SHAddToRecentDocs Lib "shell32.dll" (ByVal uFlags As Long, ByVal pv As String)
Declare Function SetCursorPos& Lib "user32" (ByVal X As Long, ByVal Y As Long)

Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
Public Const SPI_SCREENSAVERRUNNING = 97

Public Type mnuCommands
	Captions As New Collection
	Commands As New Collection
End Type

Public Type filetype
	Commands As mnuCommands
	Extension As String
	ProperName As String
	FullName As String
	ContentType As String
	IconPath As String
	IconIndex As Integer
End Type

Public Const REG_SZ = 1
Public Const HKEY_CLASSES_ROOT = &H80000000

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpszValueName As String, ByVal dwReserved As Long, ByVal fdwType As Long, lpbData As Any, ByVal cbData As Long) As Long

Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Const EWX_FORCE = 4
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long

Public Const RSP_SIMPLE_SERVICE = 1
Public Const RSP_UNREGISTER_SERVICE = 0

Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal Filename$)

Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Public Const ERROR_SUCCESS = 0&

Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

#If Win16 Then
    Type RECT
        Left As Integer
        Top As Integer
        Right As Integer
        Bottom As Integer
    End Type
#Else
    Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type
#End If

#If Win16 Then
    Declare Sub GetWindowRect Lib "User" (ByVal hwnd As Integer, lpRect As RECT)
    Declare Function GetDC Lib "User" (ByVal hwnd As Integer) As Integer
    Declare Function ReleaseDC Lib "User" (ByVal hwnd As Integer, ByVal hDC As Integer) As Integer
    Declare Sub SetBkColor Lib "GDI" (ByVal hDC As Integer, ByVal crColor As Long)
    Declare Sub Rectangle Lib "GDI" (ByVal hDC As Integer, ByVal x1 As Integer, ByVal Y1 As Integer, ByVal x2 As Integer, ByVal Y2 As Integer)
    Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
    Declare Function SelectObject Lib "GDI" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
    Declare Sub DeleteObject Lib "GDI" (ByVal hObject As Integer)
#Else
    Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
    Declare Function SetBkColor Lib "GDI32" (ByVal hDC As Long, ByVal crColor As Long) As Long
    Declare Function Rectangle Lib "GDI32" (ByVal hDC As Long, ByVal x1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
    Declare Function CreateSolidBrush Lib "GDI32" (ByVal crColor As Long) As Long
    Declare Function SelectObject Lib "user32" (ByVal hDC As Long, ByVal hObject As Long) As Long
    Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
#End If

Public i As Integer
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function OSGetLongPathName Lib "VB5STKIT.DLL" Alias "GetLongPathName" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long

Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1

Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT

Public Const SW_HIDE = 0
Public Const SW_SHOW = 5

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26

Public Const WM_CHAR = &H102
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOVE = &HF012
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

Public Const ENTER_KEY = 13
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Type POINTAPI
	X As Long
	Y As Long
End Type

Public Const APINULL = 0&
Public ReturnCode As Long

Public dTextSrc As Object
Public StatusChange As Boolean
Public TimEachLine As Integer
Public Tim4Line As Double

Global cdtoc() As toc
Global totaltr As Integer
Global r As String * 40

Type toc
	min As Long
	sec As Long
	fram As Long
	offset As Long
End Type

Function app_path() As String
Dim obj As String
obj = App.Path
If Right(obj, 1) <> "\" Then
    app_path = GetLongPathName(obj) & "\"
Else
    app_path = GetLongPathName(obj)
End If
End Function

Function GetLongPathName(ByVal strShortPath As String) As String
    Const cchBuffer = 300
    Dim strLongPath As String
    Dim lResult As Long
    
    On Error GoTo 0
    
    strLongPath = String(cchBuffer, Chr$(0))
    lResult = OSGetLongPathName(strShortPath, strLongPath, cchBuffer)
    If lResult = 0 Then
        Error 53
    Else
        GetLongPathName = StripTerminator(strLongPath)
    End If
End Function

Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Sub Window_Hide(hwnd)
Dim X
X = ShowWindow(hwnd, SW_HIDE)
End Sub

Sub Window_Show(hwnd)
Dim X
X = ShowWindow(hwnd, SW_SHOW)
End Sub

Function TimeouT(Duration)
Dim StartTime
StartTime = Timer
Do While Timer - StartTime < Duration
    DoEvents
Loop
End Function

Public Function GetUser() As String
Dim AOL As Long, MDI As Long, welcome As Long
Dim child As Long, UserString As String
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
UserString$ = GetCaption(child&)
If InStr(UserString$, "Welcome, ") = 1 Then
    UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
    GetUser$ = UserString$
    Exit Function
Else
    Do
        child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
        UserString$ = GetCaption(child&)
        If InStr(UserString$, "Welcome, ") = 1 Then
            UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
            GetUser$ = UserString$
            Exit Function
        End If
    Loop Until child& = 0&
End If
GetUser$ = ""
End Function

Public Sub Pause(Duration As Double)
Dim Current As Long
Current = Timer
Do Until Timer - Current >= Duration
    DoEvents
Loop
End Sub

Public Function GetCaption(WindowHandle As Long) As String
Dim buffer As String, TextLength As Long
TextLength& = GetWindowTextLength(WindowHandle&)
buffer$ = String(TextLength&, 0&)
Call GetWindowText(WindowHandle&, buffer$, TextLength& + 1)
GetCaption$ = buffer$
End Function

Public Function ActiveConnection() As Boolean
Dim hKey As Long
Dim lpSubKey As String
Dim phkResult As Long
Dim lpValueName As String
Dim lpReserved As Long
Dim lpType As Long
Dim lpData As Long
Dim lpcbData As Long
ActiveConnection = False
lpSubKey = "System\CurrentControlSet\Services\RemoteAccess"
ReturnCode = RegOpenKey(HKEY_LOCAL_MACHINE, lpSubKey, phkResult)

If ReturnCode = ERROR_SUCCESS Then
    hKey = phkResult
    lpValueName = "Remote Connection"
    lpReserved = APINULL
    lpType = APINULL
    lpData = APINULL
    lpcbData = APINULL
    ReturnCode = RegQueryValueEx(hKey, lpValueName, lpReserved, lpType, ByVal lpData, lpcbData)
    lpcbData = Len(lpData)
    ReturnCode = RegQueryValueEx(hKey, lpValueName, lpReserved, lpType, lpData, lpcbData)

    If ReturnCode = ERROR_SUCCESS Then
        If lpData = 0 Then

            ActiveConnection = False
        Else
            ActiveConnection = True
        End If
    End If
    RegCloseKey (hKey)
End If
End Function

Function New_File(FilePathAndFileName)
Dim NewFile As String
NewFile = FilePathAndFileName
Call SHAddToRecentDocs(2, NewFile)
End Function

Function chkWin(xHWND As Long) As Boolean
    Dim obj
    obj = IsWindowEnabled(xHWND)
        If obj = "1" Then
            chkWin = True
        ElseIf obj = "0" Then
            chkWin = False
        End If
End Function

Public Sub DisCtrAltDel()
Dim Ret As Integer
Dim dis As Boolean
Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, dis, 0)
End Sub

Public Sub EnbCtrAltDel()
Dim Ret As Integer
Dim dis As Boolean
Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, dis, 0)
End Sub

Public Function getuser_loggedon() As String
Dim s As String
Dim cnt As Long
Dim dl As Long
Dim CurUser As String
cnt = 199
s = String$(200, 0)
dl = GetUserName(s, cnt)
If dl <> 0 Then CurUser = Left$(s, cnt) Else CurUser = ""
getuser_loggedon = CurUser
End Function

Sub Delete_File(FilePathAndName)
Dim strPath As String
strPath = FilePathAndName
Kill strPath
End Sub

Sub LoadText(txtLoad As TextBox, Path As String)
    Dim TextString As String
    On Error Resume Next
    Open Path$ For Input As #1
    TextString$ = Input(LOF(1), #1)
    Close #1
    txtLoad.Text = TextString$
End Sub

Sub SaveText(txtSave As TextBox, Path As String)
    Dim TextString As String
    On Error Resume Next
    TextString$ = txtSave.Text
    Open Path$ For Output As #1
    Print #1, TextString$
    Close #1
End Sub

Public Sub ReverseFile(FromFile As String, ToFile As String)
Dim ReverseByte As Long
Open FromFile For Binary As #1
    Open ToFile For Binary As #2
        ReDim MyByte(1 To LOF(1)) As Byte
        ReDim ReversedByte(1 To LOF(1)) As Byte
        Get #1, , MyByte
        For ReverseByte = UBound(MyByte) To 1 Step -1
            ReversedByte(ReverseByte) = MyByte(UBound(MyByte) - ReverseByte + 1)
        Next
        Put #2, , ReversedByte
    Close #2
Close #1
End Sub

Public Function EnumCallback(ByVal app_hWnd As Long, ByVal param As Long) As Long
Dim buf As String * 256
Dim title As String
Dim Length As Long
Length = GetWindowText(app_hWnd, buf, Len(buf))
title = Left$(buf, Length)
If InStr(title, Target) <> 0 Then
    SendMessage app_hWnd, WM_CLOSE, 0, 0
End If
EnumCallback = 1
End Function

Public Sub TerminateTask(app_name As String)
Target = app_name
EnumWindows AddressOf EnumCallback, 0
End Sub

Public Function net() As Boolean
    net = InternetGetConnectedState(0&, 0&)
End Function

Public Sub FormLoc(frmForm As Form)
   With frmForm
      .Left = (Screen.Width - .Width)
      .Top = (Screen.Height - .Height) - 420
   End With
End Sub

Public Sub FormLoc2(frmForm As Form)
   With frmForm
      .Left = 0
      .Top = (Screen.Height - .Height) - 420
   End With
End Sub

Public Sub CreateKey(hKey As Long, strPath As String)
Dim hCurKey As Long
Dim lRegResult As Long
lRegResult = RegCreateKey(hKey, strPath, hCurKey)
If lRegResult <> ERROR_SUCCESS Then
End If
lRegResult = RegCloseKey(hCurKey)
End Sub

Public Sub DeleteKey(ByVal hKey As Long, ByVal strPath As String)
Dim lRegResult As Long
lRegResult = RegDeleteKey(hKey, strPath)
End Sub

Public Sub DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
Dim hCurKey As Long
Dim lRegResult As Long
lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lRegResult = RegDeleteValue(hCurKey, strValue)
lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function GetSettingString(hKey As Long, strPath As String, strValue As String, Optional Default As String) As String
Dim hCurKey As Long
Dim lValueType As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim intZeroPos As Integer
Dim lRegResult As Long
If Not IsEmpty(Default) Then
    GetSettingString = Default
Else
    GetSettingString = ""
End If
lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)
If lRegResult = ERROR_SUCCESS Then
    If lValueType = REG_SZ Then
        strBuffer = String(lDataBufferSize, " ")
        lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
        intZeroPos = InStr(strBuffer, Chr$(0))
        If intZeroPos > 0 Then
            GetSettingString = Left$(strBuffer, intZeroPos - 1)
        Else
            GetSettingString = strBuffer
        End If
    End If
Else
    
End If

lRegResult = RegCloseKey(hCurKey)
End Function

Public Sub SaveSettingString(hKey As Long, strPath As String, strValue As String, strData As String)
Dim hCurKey As Long
Dim lRegResult As Long
lRegResult = RegCreateKey(hKey, strPath, hCurKey)
lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))
If lRegResult <> ERROR_SUCCESS Then
    
End If
lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function GetSettingLong(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, Optional Default As Long) As Long
Dim lRegResult As Long
Dim lValueType As Long
Dim lBuffer As Long
Dim lDataBufferSize As Long
Dim hCurKey As Long
If Not IsEmpty(Default) Then
    GetSettingLong = Default
Else
    GetSettingLong = 0
End If
lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lDataBufferSize = 4
lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, lBuffer, lDataBufferSize)
If lRegResult = ERROR_SUCCESS Then
    If lValueType = REG_DWORD Then
        GetSettingLong = lBuffer
    End If
Else
    
End If
lRegResult = RegCloseKey(hCurKey)
End Function

Public Sub SaveSettingLong(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, ByVal lData As Long)
Dim hCurKey As Long
Dim lRegResult As Long
lRegResult = RegCreateKey(hKey, strPath, hCurKey)
lRegResult = RegSetValueEx(hCurKey, strValue, 0&, REG_DWORD, lData, 4)
If lRegResult <> ERROR_SUCCESS Then
    
End If
lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function GetSettingByte(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, Optional Default As Variant) As Variant
Dim lValueType As Long
Dim byBuffer() As Byte
Dim lDataBufferSize As Long
Dim lRegResult As Long
Dim hCurKey As Long
If Not IsEmpty(Default) Then
    If VarType(Default) = vbArray + vbByte Then
        GetSettingByte = Default
    Else
        GetSettingByte = 0
    End If
Else
    GetSettingByte = 0
End If
lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lRegResult = RegQueryValueEx(hCurKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufferSize)
If lRegResult = ERROR_SUCCESS Then
    If lValueType = REG_BINARY Then
        ReDim byBuffer(lDataBufferSize - 1) As Byte
        lRegResult = RegQueryValueEx(hCurKey, strValueName, 0&, lValueType, byBuffer(0), lDataBufferSize)
        GetSettingByte = byBuffer
    End If
Else
    
End If
lRegResult = RegCloseKey(hCurKey)
End Function

Public Sub SaveSettingByte(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, byData() As Byte)
Dim lRegResult As Long
Dim hCurKey As Long
lRegResult = RegCreateKey(hKey, strPath, hCurKey)
lRegResult = RegSetValueEx(hCurKey, strValueName, 0&, REG_BINARY, byData(0), UBound(byData()) + 1)
lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function GetAllKeys(hKey As Long, strPath As String) As Variant
Dim lRegResult As Long
Dim lCounter As Long
Dim hCurKey As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim strNames() As String
Dim intZeroPos As Integer
lCounter = 0
lRegResult = RegOpenKey(hKey, strPath, hCurKey)
Do
    lDataBufferSize = 255
    strBuffer = String(lDataBufferSize, " ")
    lRegResult = RegEnumKey(hCurKey, lCounter, strBuffer, lDataBufferSize)
    If lRegResult = ERROR_SUCCESS Then
        ReDim Preserve strNames(lCounter) As String
        intZeroPos = InStr(strBuffer, Chr$(0))
        If intZeroPos > 0 Then
            strNames(UBound(strNames)) = Left$(strBuffer, intZeroPos - 1)
        Else
            strNames(UBound(strNames)) = strBuffer
        End If
        lCounter = lCounter + 1
    Else
        Exit Do
    End If
Loop
GetAllKeys = strNames
End Function

Public Function GetAllValues(hKey As Long, strPath As String) As Variant
Dim lRegResult As Long
Dim hCurKey As Long
Dim lValueNameSize As Long
Dim strValueName As String
Dim lCounter As Long
Dim byDataBuffer(4000) As Byte
Dim lDataBufferSize As Long
Dim lValueType As Long
Dim strNames() As String
Dim lTypes() As Long
Dim intZeroPos As Integer
lRegResult = RegOpenKey(hKey, strPath, hCurKey)
Do
    lValueNameSize = 255
    strValueName = String$(lValueNameSize, " ")
    lDataBufferSize = 4000
    lRegResult = RegEnumValue(hCurKey, lCounter, strValueName, lValueNameSize, 0&, lValueType, byDataBuffer(0), lDataBufferSize)
    If lRegResult = ERROR_SUCCESS Then
        ReDim Preserve strNames(lCounter) As String
        ReDim Preserve lTypes(lCounter) As Long
        lTypes(UBound(lTypes)) = lValueType
        intZeroPos = InStr(strValueName, Chr$(0))
        If intZeroPos > 0 Then
            strNames(UBound(strNames)) = Left$(strValueName, intZeroPos - 1)
        Else
            strNames(UBound(strNames)) = strValueName
        End If
        lCounter = lCounter + 1
    Else
        Exit Do
    End If
Loop
Dim Finisheddata() As Variant
ReDim Finisheddata(UBound(strNames), 0 To 1) As Variant
For lCounter = 0 To UBound(strNames)
    Finisheddata(lCounter, 0) = strNames(lCounter)
    Finisheddata(lCounter, 1) = lTypes(lCounter)
Next
GetAllValues = Finisheddata
End Function

Function GetIni(Section As String, Key As String, IniFile As String)
Dim RetVal As String, Worked As Integer
RetVal = String$(255, 0)
Worked = GetPrivateProfileString(Section, Key, "", RetVal, Len(RetVal), IniFile)
If Worked = 0 Then
    GetIni = ""
Else
    GetIni = Left(RetVal, InStr(RetVal, Chr(0)) - 1)
End If
End Function

Function AddToINI(Section As String, Key As String, Value As String, IniFile As String) As Boolean
Dim Worked As Integer
Worked = WritePrivateProfileString(Section, Key, Value, IniFile)
If Worked = 0 Then
    AddToINI = False
Else
    AddToINI = True
End If
End Function

Public Sub CreateExtension(newfiletype As filetype)
Dim IconString As String
Dim Result As Long, Result2 As Long, ResultX As Long
Dim ReturnValue As Long, HKeyX As Long
Dim cmdloop As Integer
IconString = newfiletype.IconPath & "," & newfiletype.IconIndex
If Left$(newfiletype.Extension, 1) <> "." Then newfiletype.Extension = "." & newfiletype.Extension
RegCreateKey HKEY_CLASSES_ROOT, newfiletype.Extension, Result
ReturnValue = RegSetValueEx(Result, "", 0, REG_SZ, ByVal newfiletype.ProperName, LenB(StrConv(newfiletype.ProperName, vbFromUnicode))) 'Set up content type
If newfiletype.ContentType <> "" Then
    ReturnValue = RegSetValueEx(Result, "Content Type", 0, REG_SZ, ByVal CStr(newfiletype.ContentType), LenB(StrConv(newfiletype.ContentType, vbFromUnicode)))
End If
RegCreateKey HKEY_CLASSES_ROOT, newfiletype.ProperName, Result
If Not IconString = ",0" Then
    RegCreateKey Result, "DefaultIcon", Result2
    ReturnValue = RegSetValueEx(Result2, "", 0, REG_SZ, ByVal IconString, LenB(StrConv(IconString, vbFromUnicode))) 'Set The Default Value for the Key
End If
ReturnValue = RegSetValueEx(Result, "", 0, REG_SZ, ByVal newfiletype.FullName, LenB(StrConv(newfiletype.FullName, vbFromUnicode)))
RegCreateKey Result, ByVal "Shell", ResultX
For cmdloop = 1 To newfiletype.Commands.Captions.Count
    RegCreateKey ResultX, ByVal newfiletype.Commands.Captions(cmdloop), Result
    RegCreateKey Result, ByVal "Command", Result2
    Dim CurrentCommand$
    CurrentCommand = newfiletype.Commands.Commands(cmdloop)
    ReturnValue = RegSetValueEx(Result2, "", 0, REG_SZ, ByVal CurrentCommand$, LenB(StrConv(CurrentCommand$, vbFromUnicode)))
    RegCloseKey Result
    RegCloseKey Result2
Next

RegCloseKey Result2
End Sub

Public Sub FormOnTop(FormName As Form)
Call SetWindowPos(FormName.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub

Public Sub FormNotOnTop(FormName As Form)
Call SetWindowPos(FormName.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub

Sub ExplodeForm(F As Form, Movement As Integer)
Dim myRect As RECT
Dim formWidth%, formHeight%, i%, X%, Y%, CX%, cy%
Dim TheScreen As Long
Dim Brush As Long
GetWindowRect F.hwnd, myRect
formWidth = (myRect.Right - myRect.Left)
formHeight = myRect.Bottom - myRect.Top
TheScreen = GetDC(0)
Brush = CreateSolidBrush(F.BackColor)
For i = 1 To Movement
    CX = formWidth * (i / Movement)
    cy = formHeight * (i / Movement)
    X = myRect.Left + (formWidth - CX) / 2
    Y = myRect.Top + (formHeight - cy) / 2
    Rectangle TheScreen, X, Y, X + CX, Y + cy
Next i
X = ReleaseDC(0, TheScreen)
DeleteObject (Brush)
End Sub

Public Sub ImplodeForm(F As Form, Direction As Integer, Movement As Integer, ModalState As Integer)
Dim myRect As RECT
Dim formWidth%, formHeight%, i%, X%, Y%, CX%, cy%
Dim TheScreen As Long
Dim Brush As Long
GetWindowRect F.hwnd, myRect
formWidth = (myRect.Right - myRect.Left)
formHeight = myRect.Bottom - myRect.Top
TheScreen = GetDC(0)
Brush = CreateSolidBrush(F.BackColor)
For i = Movement To 1 Step -1
    CX = formWidth * (i / Movement)
    cy = formHeight * (i / Movement)
    X = myRect.Left + (formWidth - CX) / 2
    Y = myRect.Top + (formHeight - cy) / 2
    Rectangle TheScreen, X, Y, X + CX, Y + cy
Next i
X = ReleaseDC(0, TheScreen)
DeleteObject (Brush)
End Sub

Public Sub FormDrag(TheForm As Form)
Call ReleaseCapture
Call SendMessage(TheForm.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub

Public Function FileExists(sFileName As String) As Boolean
If Len(sFileName$) = 0 Then
    FileExists = False
    Exit Function
End If
If Len(Dir$(sFileName$)) Then
    FileExists = True
Else
    FileExists = False
End If
End Function

Public Sub Loadlistbox(Directory As String, TheList As ListBox)
    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
        TheList.AddItem MyString$
    Wend
    Close #1
End Sub

Private Function rString(ArgNum As Integer, srchstr As String, Delim As String) As String
    On Error GoTo Err_rString
    Dim ArgCount As Integer
    Dim LastPos As Integer
    Dim Pos As Integer
    Dim Arg As String
    Arg = ""
    LastPos = 1
    If ArgNum = 1 Then Arg = srchstr
    Do While InStr(srchstr, Delim) > 0
        Pos = InStr(LastPos, srchstr, Delim)
        If Pos = 0 Then
            If ArgCount = ArgNum - 1 Then Arg = Mid(srchstr, LastPos)
            Exit Do
        Else
            ArgCount = ArgCount + 1
            If ArgCount = ArgNum Then
                Arg = Mid(srchstr, LastPos, Pos - LastPos)
                Exit Do
            End If
        End If
        LastPos = Pos + 1
    Loop
    rString = Arg
    Exit Function
Err_rString:
    MsgBox "Error " & Err & ": " & Error
    Resume Next
End Function

Sub RND_NUM_LET(txtName As TextBox)
Randomize
Dim num As String
num = Int((62 * Rnd) + 1)
If num = 1 Then
    txtName.Text = txtName.Text & "1"
ElseIf num = 2 Then
    txtName.Text = txtName.Text & "2"
ElseIf num = 3 Then
    txtName.Text = txtName.Text & "3"
ElseIf num = 4 Then
    txtName.Text = txtName.Text & "4"
ElseIf num = 5 Then
    txtName.Text = txtName.Text & "5"
ElseIf num = 6 Then
    txtName.Text = txtName.Text & "6"
ElseIf num = 7 Then
    txtName.Text = txtName.Text & "7"
ElseIf num = 8 Then
    txtName.Text = txtName.Text & "8"
ElseIf num = 9 Then
    txtName.Text = txtName.Text & "9"
ElseIf num = 10 Then
    txtName.Text = txtName.Text & "0"

ElseIf num = 11 Then
    txtName.Text = txtName.Text & "a"
ElseIf num = 12 Then
    txtName.Text = txtName.Text & "b"
ElseIf num = 13 Then
    txtName.Text = txtName.Text & "c"
ElseIf num = 14 Then
    txtName.Text = txtName.Text & "d"
ElseIf num = 15 Then
    txtName.Text = txtName.Text & "e"
ElseIf num = 16 Then
    txtName.Text = txtName.Text & "f"
ElseIf num = 17 Then
    txtName.Text = txtName.Text & "g"
ElseIf num = 18 Then
    txtName.Text = txtName.Text & "h"
ElseIf num = 19 Then
    txtName.Text = txtName.Text & "i"
ElseIf num = 20 Then
    txtName.Text = txtName.Text & "j"
ElseIf num = 21 Then
    txtName.Text = txtName.Text & "k"
ElseIf num = 22 Then
    txtName.Text = txtName.Text & "l"
ElseIf num = 23 Then
    txtName.Text = txtName.Text & "m"
ElseIf num = 24 Then
    txtName.Text = txtName.Text & "n"
ElseIf num = 25 Then
    txtName.Text = txtName.Text & "o"
ElseIf num = 26 Then
    txtName.Text = txtName.Text & "p"
ElseIf num = 27 Then
    txtName.Text = txtName.Text & "q"
ElseIf num = 28 Then
    txtName.Text = txtName.Text & "r"
ElseIf num = 29 Then
    txtName.Text = txtName.Text & "s"
ElseIf num = 30 Then
    txtName.Text = txtName.Text & "t"
ElseIf num = 31 Then
    txtName.Text = txtName.Text & "u"
ElseIf num = 32 Then
    txtName.Text = txtName.Text & "v"
ElseIf num = 33 Then
    txtName.Text = txtName.Text & "w"
ElseIf num = 34 Then
    txtName.Text = txtName.Text & "x"
ElseIf num = 35 Then
    txtName.Text = txtName.Text & "y"
ElseIf num = 36 Then
    txtName.Text = txtName.Text & "z"

ElseIf num = 37 Then
    txtName.Text = txtName.Text & "A"
ElseIf num = 38 Then
    txtName.Text = txtName.Text & "B"
ElseIf num = 39 Then
    txtName.Text = txtName.Text & "C"
ElseIf num = 40 Then
    txtName.Text = txtName.Text & "D"
ElseIf num = 41 Then
    txtName.Text = txtName.Text & "E"
ElseIf num = 42 Then
    txtName.Text = txtName.Text & "F"
ElseIf num = 43 Then
    txtName.Text = txtName.Text & "G"
ElseIf num = 44 Then
    txtName.Text = txtName.Text & "H"
ElseIf num = 45 Then
    txtName.Text = txtName.Text & "I"
ElseIf num = 46 Then
    txtName.Text = txtName.Text & "J"
ElseIf num = 47 Then
    txtName.Text = txtName.Text & "K"
ElseIf num = 48 Then
    txtName.Text = txtName.Text & "L"
ElseIf num = 49 Then
    txtName.Text = txtName.Text & "M"
ElseIf num = 50 Then
    txtName.Text = txtName.Text & "N"
ElseIf num = 51 Then
    txtName.Text = txtName.Text & "O"
ElseIf num = 52 Then
    txtName.Text = txtName.Text & "P"
ElseIf num = 53 Then
    txtName.Text = txtName.Text & "Q"
ElseIf num = 54 Then
    txtName.Text = txtName.Text & "R"
ElseIf num = 55 Then
    txtName.Text = txtName.Text & "S"
ElseIf num = 56 Then
    txtName.Text = txtName.Text & "T"
ElseIf num = 57 Then
    txtName.Text = txtName.Text & "U"
ElseIf num = 58 Then
    txtName.Text = txtName.Text & "V"
ElseIf num = 59 Then
    txtName.Text = txtName.Text & "W"
ElseIf num = 60 Then
    txtName.Text = txtName.Text & "X"
ElseIf num = 61 Then
    txtName.Text = txtName.Text & "Y"
ElseIf num = 62 Then
    txtName.Text = txtName.Text & "Z"
End If
End Sub

Public Function GetPathName(Filename As String) As String
Dim intPathPos As Integer
Dim intExtPos As Integer
Dim i As Integer
Dim J As Integer
    For i = Len(Filename) To 1 Step -1
    If Mid(Filename, i, 1) = "." Then
        intExtPos = i
        For J = Len(Filename) To 1 Step -1
        If Mid(Filename, J, 1) = "\" Then
            intPathPos = J
        Exit For
    End If
        Next J
    Exit For
    End If
    Next i
    If intPathPos > intExtPos Then
        Exit Function
    Else
        If intExtPos = 0 Then Exit Function
        GetPathName = Mid(Filename, 1, intPathPos)
    End If
End Function

Public Function GetFileName(Filename As String) As String
Dim intPathPos As Integer
Dim intExtPos As Integer
Dim i As Integer
Dim J As Integer
Dim X As Integer
    For i = Len(Filename) To 1 Step -1
    If Mid(Filename, i, 1) = "." Then
        intExtPos = i
        For J = Len(Filename) To 1 Step -1
        If Mid(Filename, J, 1) = "\" Then
            intPathPos = J
        Exit For
    End If
        Next J
    Exit For
    End If
    Next i
    If intPathPos > intExtPos Then
        Exit Function
    Else
        If intExtPos = 0 Then Exit Function
        Dim tempStr, intExtPos2, xVal As String
            tempStr = Mid(Filename, intPathPos + 1)
            For X = Len(tempStr) To 1 Step -1
                If Mid(tempStr, X, 1) = "." Then
                    intExtPos2 = X
                    Exit For
                End If
            Next X
                xVal = Len(tempStr) - intExtPos2 + 1
            GetFileName = Mid(tempStr, 1, Len(tempStr) - xVal)
    End If
End Function

Public Function GetExtension(Filename As String) As String
Dim intPathPos As Integer
Dim intExtPos As Integer
Dim i As Integer
Dim J As Integer
    For i = Len(Filename) To 1 Step -1
        If Mid(Filename, i, 1) = "." Then
            intExtPos = i
            For J = Len(Filename) To 1 Step -1
                If Mid(Filename, J, 1) = "\" Then
                    intPathPos = J
                    Exit For
                End If
            Next J
        Exit For
        End If
    Next i
    If intPathPos > intExtPos Then
        Exit Function
    Else
        If intExtPos = 0 Then Exit Function
        GetExtension = Mid(Filename, intExtPos + 1, Len(Filename) - intExtPos)
    End If
End Function

Public Function GetFileNameAndExtension(Filename As String) As String
Dim intPathPos As Integer
Dim intExtPos As Integer
Dim i As Integer
Dim J As Integer
    For i = Len(Filename) To 1 Step -1
        If Mid(Filename, i, 1) = "." Then
            intExtPos = i
            For J = Len(Filename) To 1 Step -1
                If Mid(Filename, J, 1) = "\" Then
                    intPathPos = J
                    Exit For
                End If
            Next J
            Exit For
        End If
    Next i
    If intPathPos > intExtPos Then
        Exit Function
    Else
        If intExtPos = 0 Then Exit Function
        GetFileNameAndExtension = Mid(Filename, intPathPos + 1, Len(Filename) + intPathPos - intExtPos + 1)
    End If
End Function

'commented because you need an Inet control on the form, or you get an error.
'Public Function GetWebServerFile(InetName As String, webAddress As String, destPathName As String) As Boolean
'On Error GoTo hdl_error_description
'Dim b() As Byte
'InetName.Cancel
'InetName.Protocol = icHTTP
'InetName.URL = webAddress
'b() = InetName.OpenURL(, icByteArray)
'Open destPathName For Binary Access Write As #1
'    Put #1, , b()
'Close #1
'GetWebServerFile = True
'
'Exit Function
'hdl_error_description:
'GetWebServerFile = False
'MsgBox "Error number '" & Err.Number & vbCrLf & Err.Description, vbCritical + vbSystemModal, "Error"
'End Function

Public Function myRound(Number, Optional DecimalPlaces = 0) As String
    myRound = Int(Number * 10 ^ DecimalPlaces + 0.5 / 10 ^ DecimalPlaces)
End Function

Public Function chkCommandLine(cmdLine As String) As Boolean
If Len(cmdLine) > 0 Then
    chkCommandLine = True
Else
    chkCommandLine = False
End If
End Function

Public Function split_str(strText As String, sepChar As String, lstListBox As ListBox)
Dim i As Integer
Dim curPos As Integer
Dim curChar As String
Dim xString1 As String
Dim xString2 As String
i = 1
curPos = 1
Do Until i = Len(strText)
    curChar = Mid(strText, i, 1)
    If curChar = sepChar Then
        xString1 = Mid(strText, 1, i - 1)
        xString2 = Mid(strText, i + 1, Len(strText))
        lstListBox.AddItem xString1
        lstListBox.AddItem xString2
    End If
    i = i + 1
Loop
End Function

Public Function DirExists(ByVal strDirName As String) As Boolean
    Dim strDummy As String
    On Error Resume Next
    strDummy = Dir$(strDirName & "*.*", vbDirectory)
    If Len(strDummy) <= 0 Then
        DirExists = False
    Else
        DirExists = True
    End If
    Err = 0
End Function

Public Function current_tasks(lstListBox As ListBox)
Dim hSnapShot As Long
Dim uProcess As PROCESSENTRY32
Dim r As Long
hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
If hSnapShot = 0 Then
    Exit Function
End If
uProcess.dwSize = Len(uProcess)
r = ProcessFirst(hSnapShot, uProcess)
    Do While r
        lstListBox.AddItem LCase(uProcess.szExeFile)
        r = ProcessNext(hSnapShot, uProcess)
    Loop
Call CloseHandle(hSnapShot)
End Function

Public Function get_cd_id(frmForm As Form) As String
mciSendString "close all", 0, 0, 0
If (SendMCIString(frmForm, "open cdaudio alias cd69 wait shareable", True) = False) Then
    Exit Function
End If
SendMCIString frmForm, "set cd69 time format tmsf wait", True
readcdtoc
get_cd_id = cddbdiscid(totaltr)
End Function

Public Function SendMCIString(frmForm As Form, cmd As String, fShowError As Boolean) As Boolean
Static rc As Long
Static errStr As String * 200
rc = mciSendString(cmd, 0, 0, frmForm.hwnd)
If (fShowError And rc <> 0) Then
    mciGetErrorString rc, errStr, Len(errStr)
    MsgBox errStr
End If
SendMCIString = (rc = 0)
End Function

Public Function readcdtoc() As Integer
mciSendString "status cd69 number of tracks wait", r, Len(r), 0
On Error Resume Next
totaltr = CInt(Mid$(r, 1, 2))
ReDim cdtoc(totaltr + 1) As toc
mciSendString "set cd69 time format msf", 0, 0, 0
For i = 1 To totaltr
    cmd = "status cd69 position track " & i
    mciSendString cmd, r, Len(r), 0
    cdtoc(i - 1).min = CInt(Mid$(r, 1, 2))
    cdtoc(i - 1).sec = CInt(Mid$(r, 4, 2))
    cdtoc(i - 1).fram = CInt(Mid$(r, 7, 2))
    cdtoc(i - 1).offset = (cdtoc(i - 1).min * 60 * 75) + (cdtoc(i - 1).sec * 75) + cdtoc(i - 1).fram
Next
End Function

Public Function cddbsum(n) As Integer
Ret = 0
m = n
For i = 1 To m
    Ret = Ret + (n Mod 10)
    n = n / 10
Next
cddbsum = Ret
End Function

Public Function cddbdiscid(tr) As String
Dim n As Long
Dim tm As Long
For i = 0 To tr - 1
    tm = ((cdtoc(i).min * 60) + cdtoc(i).sec)
    Do While tm > 0
        n = n + (tm Mod 10)
        tm = tm \ 10
    Loop
Next
mciSendString "status cd69 length wait", r, Len(r), 0
On Error Resume Next
t = (CInt(Mid$(r, 1, 2)) * 60) + CInt(Mid$(r, 4, 2))
cddbdiscid = LCase$(Zeros(Hex$(n Mod &HFF), 2) & Zeros(Hex$(t), 4) & Zeros(Hex$(tr), 2))
End Function

Private Function Zeros(s As String, n As Integer) As String
If Len(s) < n Then
    Zeros = String$(n - Len(s), "0") & s
Else
    Zeros = s
End If
End Function

Private Function ChangeSpaces(cString As String) As String
On Error Resume Next
Dim cChar As String
Dim cReturn As String
Dim nLoop As Long
cReturn = ""
For nLoop = 1 To Len(cString)
    cChar = Mid(cString, nLoop, 1)
    If cChar = " " Then
        cChar = "+"
    End If
    cReturn = cReturn + cChar
Next
ChangeSpaces = cReturn
End Function

'commented for above reason
'Public Function icq_pager(wskPager As Winsock, strUIN As String, strMessage As String, strSubject As String, strName As String, strEMail As String)
'On Error Resume Next
'Dim cMessage As String
'Dim cSubject As String
'Dim cSend As String
'Dim cData As String
'' Verify datas
'If Not IsNumeric(txtUIN) Then
'    Exit Function
'End If
'If Trim(strMessage) = "" Then
'    Exit Function
'End If
'' Status
'LabelStatus.Caption = "Starting..."
'' Close Socket
'SockPager.Close
'' Change the " " for "+"
'cSubject = ChangeSpaces(strSubject)
'cMessage = ChangeSpaces(strMessage)
'' Fill the String
'    cData = "from=" & strName & "&fromemail=" & strEMail & "&subject=" & cSubject & "&body=" & cMessage & "&to=" & Trim(strUIN) & "&Send=" & """"
'    cSend = "POST /scripts/WWPMsg.dll HTTP/1.0" & vbCrLf
'    cSend = cSend & "Referer: http://wwp.mirabilis.com" & vbCrLf
'    cSend = cSend & "User-Agent: Mozilla/4.06 (Win95; I)" & vbCrLf
'    cSend = cSend & "Connection: Keep-Alive" & vbCrLf
'    cSend = cSend & "Host: wwp.mirabilis.com:80" & vbCrLf
'    cSend = cSend & "Content-type: application/x-www-form-urlencoded" & vbCrLf
'    cSend = cSend & "Content-length: " & Len(cData) & vbCrLf
'    cSend = cSend & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, */*" & vbCrLf & vbCrLf
'    cSend = cSend & cData & vbCrLf & vbCrLf & vbCrLf & vbCrLf
'' Send Message
'wskPager.Tag = cSend
'wskPager.Connect "wwp.mirabilis.com", 80
'Dim i As Integer
'i = 0
'HDL_RESEND:
'i = i + 1
'On Error GoTo HDL_ERROR
'    wskPager.SendData wskPager.Tag
'
'Exit Function
'HDL_ERROR:
'If i = 10 Then
'    Exit Function
'Else
'    TimeouT 0.5
'    GoTo HDL_RESEND
'End If
'End Function

'this function should be called from a text_change sub:
Public Function short_string(cControl As Label, lngMax As Long)
If cControl.Width >= lngMax Then
    On Error Resume Next
    cControl.Caption = Mid(cControl.Caption, 1, Len(cControl.Caption) - 3) & "..."
    Do Until cControl.Width <= lngMax
        cControl.Caption = Mid(cControl.Caption, 1, Len(cControl.Caption) - 4) & "..."
    Loop
End If
cControl.Refresh
End Function

Public Function identify_char(strChar As String) As Integer
'-2=invalid entry,-1=unknown,0=number,1=letter,2=symbol
If strChar = "0" Then identify_char = 0: Exit Function
If strChar = "1" Then identify_char = 0: Exit Function
If strChar = "2" Then identify_char = 0: Exit Function
If strChar = "3" Then identify_char = 0: Exit Function
If strChar = "4" Then identify_char = 0: Exit Function
If strChar = "5" Then identify_char = 0: Exit Function
If strChar = "6" Then identify_char = 0: Exit Function
If strChar = "7" Then identify_char = 0: Exit Function
If strChar = "8" Then identify_char = 0: Exit Function
If strChar = "9" Then identify_char = 0: Exit Function
If LCase(strChar) = "a" Then identify_char = 1: Exit Function
If LCase(strChar) = "b" Then identify_char = 1: Exit Function
If LCase(strChar) = "c" Then identify_char = 1: Exit Function
If LCase(strChar) = "d" Then identify_char = 1: Exit Function
If LCase(strChar) = "e" Then identify_char = 1: Exit Function
If LCase(strChar) = "f" Then identify_char = 1: Exit Function
If LCase(strChar) = "g" Then identify_char = 1: Exit Function
If LCase(strChar) = "h" Then identify_char = 1: Exit Function
If LCase(strChar) = "i" Then identify_char = 1: Exit Function
If LCase(strChar) = "j" Then identify_char = 1: Exit Function
If LCase(strChar) = "k" Then identify_char = 1: Exit Function
If LCase(strChar) = "l" Then identify_char = 1: Exit Function
If LCase(strChar) = "m" Then identify_char = 1: Exit Function
If LCase(strChar) = "n" Then identify_char = 1: Exit Function
If LCase(strChar) = "o" Then identify_char = 1: Exit Function
If LCase(strChar) = "p" Then identify_char = 1: Exit Function
If LCase(strChar) = "q" Then identify_char = 1: Exit Function
If LCase(strChar) = "r" Then identify_char = 1: Exit Function
If LCase(strChar) = "s" Then identify_char = 1: Exit Function
If LCase(strChar) = "t" Then identify_char = 1: Exit Function
If LCase(strChar) = "u" Then identify_char = 1: Exit Function
If LCase(strChar) = "v" Then identify_char = 1: Exit Function
If LCase(strChar) = "w" Then identify_char = 1: Exit Function
If LCase(strChar) = "x" Then identify_char = 1: Exit Function
If LCase(strChar) = "y" Then identify_char = 1: Exit Function
If LCase(strChar) = "z" Then identify_char = 1: Exit Function
If strChar = "!" Then identify_char = 2: Exit Function
If strChar = "@" Then identify_char = 2: Exit Function
If strChar = "#" Then identify_char = 2: Exit Function
If strChar = "$" Then identify_char = 2: Exit Function
If strChar = "%" Then identify_char = 2: Exit Function
If strChar = "^" Then identify_char = 2: Exit Function
If strChar = "&" Then identify_char = 2: Exit Function
If strChar = "*" Then identify_char = 2: Exit Function
If strChar = "(" Then identify_char = 2: Exit Function
If strChar = ")" Then identify_char = 2: Exit Function
If strChar = "_" Then identify_char = 2: Exit Function
If strChar = "+" Then identify_char = 2: Exit Function
If strChar = "{" Then identify_char = 2: Exit Function
If strChar = "}" Then identify_char = 2: Exit Function
If strChar = "|" Then identify_char = 2: Exit Function
If strChar = ":" Then identify_char = 2: Exit Function
If strChar = """" Then identify_char = 2: Exit Function
If strChar = "<" Then identify_char = 2: Exit Function
If strChar = ">" Then identify_char = 2: Exit Function
If strChar = "?" Then identify_char = 2: Exit Function
If strChar = "~" Then identify_char = 2: Exit Function
If strChar = "-" Then identify_char = 2: Exit Function
If strChar = "=" Then identify_char = 2: Exit Function
If strChar = "[" Then identify_char = 2: Exit Function
If strChar = "]" Then identify_char = 2: Exit Function
If strChar = "\" Then identify_char = 2: Exit Function
If strChar = ";" Then identify_char = 2: Exit Function
If strChar = "'" Then identify_char = 2: Exit Function
If strChar = "," Then identify_char = 2: Exit Function
If strChar = "." Then identify_char = 2: Exit Function
If strChar = "/" Then identify_char = 2: Exit Function
If strChar = "`" Then identify_char = 2: Exit Function
If Len(strChar) > 1 Then identify_char = -2: Exit Function
identify_char = -1: Exit Function
End Function

Public Function rename_file(strFullPath As String, strOldName As String, strNewName As String)
Name strFullPath & strOldName As strFullPath & strNewName
End Function

Public Function return_percent(tasknum As Integer, taskcount As Integer) As Integer
return_percent = Int((tasknum * 100) / taskcount)
End Function

Public Function sep_string(m_string As String, m_sepchar As String, m_listbox As ListBox)
saved = m_string
i = 1
res = 1
def = 1
Do While res > 0
    res = InStr(def, saved, m_sepchar)
    If InStr(def + 1, saved, m_sepchar) = 0 Then
        counted = Len(saved)
    Else
        counted = InStr(def + 1, saved, m_sepchar) - def
    End If
    m_listbox.AddItem Mid(saved, def, counted)
    def = res + 1
    i = i + 1
Loop
End Function
