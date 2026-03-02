Imports System.Runtime.InteropServices
Imports System.Text

Namespace MacroAutoControl
    ''' <summary>
    ''' Win32 API 선언 모음
    ''' </summary>
    Module NativeMethods
        ' 마우스 이벤트 플래그
        Public Const MOUSEEVENTF_LEFTDOWN As Integer = &H2
        Public Const MOUSEEVENTF_LEFTUP As Integer = &H4
        Public Const MOUSEEVENTF_RIGHTDOWN As Integer = &H8
        Public Const MOUSEEVENTF_RIGHTUP As Integer = &H10

        ' SendMessage 상수
        Public Const WM_LBUTTONDOWN As Integer = &H201
        Public Const WM_LBUTTONUP As Integer = &H202
        Public Const MK_LBUTTON As Integer = &H1
        Public Const WM_RBUTTONDOWN As Integer = &H204
        Public Const WM_RBUTTONUP As Integer = &H205
        Public Const MK_RBUTTON As Integer = &H2

        ' 키보드 메시지 상수
        Public Const WM_KEYDOWN As Integer = &H100
        Public Const WM_KEYUP As Integer = &H101
        Public Const WM_CHAR As Integer = &H102
        Public Const KEYEVENTF_KEYUP As Integer = &H2

        ' Window 상태 상수
        Public Const SW_RESTORE As Integer = 9
        Public Const SW_SHOW As Integer = 5

        ' RECT 구조체
        <StructLayout(LayoutKind.Sequential)>
        Public Structure RECT
            Public Left As Integer
            Public Top As Integer
            Public Right As Integer
            Public Bottom As Integer
        End Structure

        ' POINT 구조체 (Win32 API용)
        <StructLayout(LayoutKind.Sequential)>
        Public Structure APIPOINT
            Public X As Integer
            Public Y As Integer
        End Structure

        ' 콜백 대리자
        Public Delegate Function EnumWindowsProc(hWnd As IntPtr, lParam As IntPtr) As Boolean

        <DllImport("user32.dll", SetLastError:=True)>
        Public Function EnumWindows(lpEnumFunc As EnumWindowsProc, lParam As IntPtr) As Boolean
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Public Function GetWindowText(hWnd As IntPtr, lpString As StringBuilder, nMaxCount As Integer) As Integer
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Function GetWindowTextLength(hWnd As IntPtr) As Integer
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Function IsWindowVisible(hWnd As IntPtr) As Boolean
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Function GetWindowRect(hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Function GetClientRect(hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Function ClientToScreen(hWnd As IntPtr, ByRef lpPoint As APIPOINT) As Boolean
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Function SetForegroundWindow(hWnd As IntPtr) As Boolean
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Function ShowWindow(hWnd As IntPtr, nCmdShow As Integer) As Boolean
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Function IsIconic(hWnd As IntPtr) As Boolean
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Function SetCursorPos(X As Integer, Y As Integer) As Boolean
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Function GetCursorPos(ByRef lpPoint As APIPOINT) As Boolean
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Sub mouse_event(dwFlags As Integer, dx As Integer, dy As Integer, dwData As Integer, dwExtraInfo As IntPtr)
        End Sub

        <DllImport("user32.dll", SetLastError:=True)>
        Public Sub keybd_event(bVk As Byte, bScan As Byte, dwFlags As Integer, dwExtraInfo As IntPtr)
        End Sub

        <DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Public Function SendMessage(hWnd As IntPtr, Msg As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Public Function PostMessage(hWnd As IntPtr, Msg As Integer, wParam As IntPtr, lParam As IntPtr) As Boolean
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Function PrintWindow(hWnd As IntPtr, hDC As IntPtr, nFlags As UInteger) As Boolean
        End Function

        ''' <summary>
        ''' lParam 값 생성 (좌표를 LPARAM으로 변환)
        ''' </summary>
        Public Function MakeLParam(x As Integer, y As Integer) As IntPtr
            Return New IntPtr((y << 16) Or (x And &HFFFF))
        End Function

        ' =============================================
        ' 글로벌 키보드 훅 (Low-Level Keyboard Hook)
        ' =============================================
        Public Const WH_KEYBOARD_LL As Integer = 13
        Public Const HC_ACTION As Integer = 0

        Public Delegate Function LowLevelKeyboardProc(nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr

        <DllImport("user32.dll", SetLastError:=True)>
        Public Function SetWindowsHookEx(idHook As Integer, lpfn As LowLevelKeyboardProc, hMod As IntPtr, dwThreadId As UInteger) As IntPtr
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Function UnhookWindowsHookEx(hhk As IntPtr) As Boolean
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Function CallNextHookEx(hhk As IntPtr, nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
        End Function

        <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Public Function GetModuleHandle(lpModuleName As String) As IntPtr
        End Function

        <StructLayout(LayoutKind.Sequential)>
        Public Structure KBDLLHOOKSTRUCT
            Public vkCode As Integer
            Public scanCode As Integer
            Public flags As Integer
            Public time As Integer
            Public dwExtraInfo As IntPtr
        End Structure
    End Module
End Namespace
