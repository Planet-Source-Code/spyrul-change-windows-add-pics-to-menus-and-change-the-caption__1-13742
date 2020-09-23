Attribute VB_Name = "modMain"
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Public Declare Function EnableWindow& Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long)
Public Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal dwMessage As Long) As Integer
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal ProcessID As Long, ByVal ServiceFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long

Global DontRemove As Boolean
Global RefreshD As Boolean

Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const MAX_PATH& = 260
Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_SHOW = 5
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const WM_CLOSE = &H10
Public Const FLAGS = SWP_NOSIZE Or SWP_NOMOVE

Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * MAX_PATH
End Type
Public Sub FreezeComputer()
On Error Resume Next

a = frmMain.hWnd
b = frmMain.hWnd

Freeze = SetParent(a, b)
End Sub
Sub WindowHandle(hWindow, mCase As Long)
Select Case mCase
    Case 0
        X% = SendMessage(hWindow, WM_CLOSE, 0, 0)
    Case 1
        X = ShowWindow(hWindow, SW_SHOW)
    Case 2
        X = ShowWindow(hWindow, SW_HIDE)
    Case 3
        X = ShowWindow(hWindow, SW_MAXIMIZE)
    Case 4
        X = ShowWindow(hWindow, SW_MINIMIZE)
End Select
End Sub
Public Function GetWindowTitle(ByVal hWnd As Long) As String
On Error Resume Next
Dim s As String

l = GetWindowTextLength(hWnd)
s = Space(l + 1)

GetWindowText hWnd, s, l + 1
GetWindowTitle = Left$(s, l)
End Function
Public Sub CloseWin(hWnd)
On Error Resume Next

rReturn = SendMessageByString(hWnd, &H10, 0, 0&)
End Sub
