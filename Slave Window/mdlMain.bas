Attribute VB_Name = "mdlMain"
Option Explicit

'Windows messages that might be used in SendMessage() and/or
'processed in our window Callback routine
Public Const WM_MOVE = &H3&
'...
Public Const WM_ACTIVATEAPP = &H1C&
'...
Public Const WM_NCACTIVATE = &H86&
'...
Public Const WM_NCLBUTTONDOWN = &HA1&
'...
Public Const WM_SYSCOMMAND = &H112&
'...
Public Const WM_EXITSIZEMOVE = &H232&

'wParam values for WM_SYSCOMMAND argument
Public Const SC_SIZE = &HF000&
Public Const SC_MOVE = &HF010&
Public Const SC_MINIMIZE = &HF020&
Public Const SC_MAXIMIZE = &HF030&
Public Const SC_NEXTWINDOW = &HF040&
Public Const SC_PREVWINDOW = &HF050&
Public Const SC_CLOSE = &HF060&
Public Const SC_VSCROLL = &HF070&
Public Const SC_HSCROLL = &HF080&
Public Const SC_MOUSEMENU = &HF090&
Public Const SC_KEYMENU = &HF100&
Public Const SC_ARRANGE = &HF110&
Public Const SC_RESTORE = &HF120&
Public Const SC_TASKLIST = &HF130&
Public Const SC_SCREENSAVE = &HF140&
Public Const SC_HOTKEY = &HF150&

'Constants used by the CombineRgn() API function.
Public Const RGN_AND = 1&
Public Const RGN_OR = 2&
Public Const RGN_XOR = 3&
Public Const RGN_DIFF = 4&
Public Const RGN_COPY = 5&

'Constants used by the GetVersionEx() API function.
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

'Constant used by SetWindowLong() to establish a new address for the window procedure.
Public Const GWL_STYLE = -16&
Public Const GWL_WNDPROC = -4&

'Constants used by the ModifyMenu() API function.
Public Const MF_BYCOMMAND = &H0&
Public Const MF_ENABLED = &H0&
Public Const MF_GRAYED = &H1&
Public Const MF_DISABLED = &H2&
Public Const MF_BYPOSITION = &H400&

'Flags used by SetWindowLong() API function.
Public Const WS_SYSMENU = &H80000

'Flags used by SetWindowPos() to prevent the form from jumping to the top of the Z order.
Public Const SWP_NOSIZE = &H1&
Public Const SWP_NOMOVE = &H2&

'Flags used by GetSysColor() API function.
Public Const COLOR_ACTIVECAPTION = 2&
Public Const COLOR_INACTIVECAPTION = 3&
Public Const COLOR_CAPTIONTEXT = 9&
Public Const COLOR_INACTIVECAPTIONTEXT = 19&
Public Const COLOR_GRADIENTACTIVECAPTION = 27&
Public Const COLOR_GRADIENTINACTIVECAPTION = 28&

'Type used by GetWindowRect().
Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

'Type used by CreatePolygonRgn().
Public Type POINTAPI
    X As Long
    Y As Long
End Type

'Type used by the GetVersionEx() API function.
Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long 'value should be set to 148
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
    szCSDVersion As String * 128
End Type

'API Sub declarations.
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'API Functions used to shape forms.
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

'Other API function declarations.
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Function WndCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim rMaster As RECT
    Dim rSlave As RECT
    Dim OrigWndProc As Long

    Select Case hWnd
        Case frmMaster.hWnd
        'Trap messages for the "Master" form.
        
            'Set the value for the window's original hook procedure.
            OrigWndProc = frmMaster.OrigWndProc
        
            Select Case uMsg
                Case WM_MOVE
                'This message is received after the window has been moved.
                'We need to know this so we can move the "slave" window to match.
                
                    'Find out where the "master" window is.
                    Call GetWindowRect(frmMaster.hWnd, rMaster)
                    'Find out where the "slave" window is.
                    Call GetWindowRect(frmSlave.hWnd, rSlave)
                    'Move the "slave" window along with the "master" and make sure it stays the same size.
                    Call MoveWindow(frmSlave.hWnd, rMaster.Left, rMaster.Bottom + frmSlave.YOffset, rSlave.Right - rSlave.Left, rSlave.Bottom - rSlave.Top, 1)
                
                Case WM_ACTIVATEAPP
                'This message is received when the application is about to be
                'activated or deactivated. The wParam value will tell us whether
                'it is being activated or deactivated. We need to monitor this
                'message so we know when to redraw our title bar.
                
                    Call frmMaster.DrawTitleBar(wParam)
                    
                Case WM_SYSCOMMAND
                'This message is sent when the user chooses a command from the
                'system menu (from the taskbar). We need to watch for an SC_CLOSE
                'message so we know to exit the application.
                
                    If (wParam = SC_CLOSE) Then
                        Unload frmSlave
                        Unload frmMaster
                    End If
                
                Case WM_EXITSIZEMOVE
                'This message is received when the window exits the moving modal loop.
                'This is how we can tell that the user has stopped moving the window.
                
                    'Modify the system menu to disable the "Move" item since we can't respond
                    'to this command correctly with a borderless window.
                    Call ModifyMenu(GetSystemMenu(frmMaster.hWnd, False), SC_MOVE, MF_BYCOMMAND Or MF_GRAYED, 0&, "&Move")
                
                Case Else
                    'Nothing to do here
                    
            End Select
        
        Case frmSlave.hWnd
        'Trap messages for the "Slave" form.
        
            'Set the value for the window's original hook procedure.
            OrigWndProc = frmSlave.OrigWndProc
        
            Select Case uMsg
                Case WM_NCACTIVATE
                'This message is received when the window is about to be activated
                'or deactivated. We need to monitor this message to make sure the
                '"slave" form doesn't jump in front of the "master" form which
                'would cause an ugly and undesirable visual glitch.
                    
                    Call SetWindowPos(frmSlave.hWnd, frmMaster.hWnd, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
                    
                Case Else
                    'Nothing to do here
                    
            End Select
        
        Case Else
            'Nothing to do here
            
    End Select

    'Pass the message to the original window procedure so we don't hang the program.
    WndCallbackProc = CallWindowProc(OrigWndProc, hWnd, uMsg, wParam, lParam)

End Function

Public Function LowWord(ByVal dwValue As Long) As Long
    'Returns the low order word of a dWord value
    CopyMemory LowWord, dwValue, 2
End Function

Public Function HighWord(ByVal dwValue As Long) As Long
    'Returns the high order word of a dWord value
    CopyMemory HighWord, ByVal VarPtr(dwValue) + 2, 2
End Function

Public Function MSByte(ByVal dwValue As Long) As Long
    'Returns the most significant byte of a dWord value
    CopyMemory MSByte, ByVal VarPtr(dwValue) + 3, 1
End Function

Public Function NMSByte(ByVal dwValue As Long) As Long
    'Returns the next most significant byte of a dWord value
    CopyMemory NMSByte, ByVal VarPtr(dwValue) + 2, 1
End Function

Public Function NLSByte(ByVal dwValue As Long) As Long
    'Returns the next least significant byte of a dWord value
    CopyMemory NLSByte, ByVal VarPtr(dwValue) + 1, 1
End Function

Public Function LSByte(ByVal dwValue As Long) As Long
    'Returns the least significant byte of a dWord value
    CopyMemory LSByte, dwValue, 1
End Function

Public Function OSType() As Long

    'This function determines the current operating system version.
    'Returns 95 for Windows 95, 98 for Windows 98, 4 for Windows NT, 5 for Windows 2000.

    Dim VersionInfo As OSVERSIONINFO
    
    VersionInfo.dwOSVersionInfoSize = 148
    Call GetVersionEx(VersionInfo)
    If VersionInfo.dwPlatformID = VER_PLATFORM_WIN32_WINDOWS Then
        If VersionInfo.dwMinorVersion = 0 Then OSType = 95 Else OSType = 98
    ElseIf VersionInfo.dwPlatformID = VER_PLATFORM_WIN32_NT Then
        OSType = VersionInfo.dwMajorVersion
    End If

End Function


