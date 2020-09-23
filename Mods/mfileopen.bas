Attribute VB_Name = "mfileopen"
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const ERROR_ALREADY_EXISTS = 183&
Private m_hMutex As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Const SMTO_NORMAL = &H0
Public Const WM_COPYDATA = &H4A
Public Type COPYDATASTRUCT
   dwData As Long
   cbData As Long
   lpData As Long
End Type
Private Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_RESTORE = &HF120
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long


Private m_hWndPrevious As Long
Private m_bInDevelopment As Boolean

Private Const mcTHISAPPID = "pbaspediter2000"



Public Sub RestoreAndActivate(ByVal hwnd As Long)
   If (IsIconic(hwnd)) Then
      SendMessageByLong hwnd, WM_SYSCOMMAND, SC_RESTORE, 0
   End If
   ActivateWindow hwnd
End Sub

Public Sub TagWindow(ByVal hwnd As Long)

   SetProp hwnd, mcTHISAPPID & "_APPLICATION", 1
End Sub

Private Function IsThisApp(ByVal hwnd As Long) As Boolean

   If GetProp(hwnd, mcTHISAPPID & "_APPLICATION") = 1 Then
      IsThisApp = True
   End If
End Function


Public Function EnumWindowsProc( _
        ByVal hwnd As Long, _
        ByVal lParam As Long _
    ) As Long
Dim bStop As Boolean

   bStop = False
   If IsThisApp(hwnd) Then
      EnumWindowsProc = 0
      m_hWndPrevious = hwnd
   Else
      EnumWindowsProc = 1
   End If
End Function

Public Function EnumerateWindows() As Boolean
   EnumWindows AddressOf EnumWindowsProc, 0
End Function

Public Sub ActivateWindow(ByVal lHwnd As Long)
    SetForegroundWindow lHwnd
End Sub
Public Function InDevelopment() As Boolean

   Debug.Assert InDevelopmentHack() = True
   InDevelopment = m_bInDevelopment
End Function

Private Function InDevelopmentHack() As Boolean
   m_bInDevelopment = True
   InDevelopmentHack = m_bInDevelopment
End Function

Private Function WeAreAlone(ByVal sMutex As String) As Boolean
On Error Resume Next
   If InDevelopment Then
      WeAreAlone = Not (App.PrevInstance)
   Else

      m_hMutex = CreateMutex(ByVal 0&, 1, sMutex)
      If (Err.LastDllError = ERROR_ALREADY_EXISTS) Then
         CloseHandle m_hMutex
      Else
         WeAreAlone = True
      End If
   End If
End Function

Public Function EndApp()
On Error Resume Next
   If (m_hMutex <> 0) Then
      CloseHandle m_hMutex
   End If
   m_hMutex = 0
End Function

