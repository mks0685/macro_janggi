Imports System.Drawing
Imports System.Threading

Namespace MacroAutoControl
    Public Enum ClickButton
        Left = 0
        Right = 1
    End Enum

    ''' <summary>
    ''' 찾은 버튼을 클릭하는 클래스
    ''' </summary>
    Public Class ButtonClicker
        ''' <summary>
        ''' 화면 좌표로 마우스 클릭 (창을 전면에 가져온 후 클릭)
        ''' </summary>
        Public Shared Sub ClickAtScreen(screenX As Integer, screenY As Integer, Optional button As ClickButton = ClickButton.Left)
            NativeMethods.SetCursorPos(screenX, screenY)
            Thread.Sleep(10)
            If button = ClickButton.Right Then
                NativeMethods.mouse_event(NativeMethods.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, IntPtr.Zero)
                Thread.Sleep(10)
                NativeMethods.mouse_event(NativeMethods.MOUSEEVENTF_RIGHTUP, 0, 0, 0, IntPtr.Zero)
            Else
                NativeMethods.mouse_event(NativeMethods.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, IntPtr.Zero)
                Thread.Sleep(10)
                NativeMethods.mouse_event(NativeMethods.MOUSEEVENTF_LEFTUP, 0, 0, 0, IntPtr.Zero)
            End If
        End Sub

        ''' <summary>
        ''' 창 내부의 상대 좌표에 PostMessage로 클릭 (백그라운드 클릭)
        ''' </summary>
        Public Shared Sub ClickAtWindow(hWnd As IntPtr, clientX As Integer, clientY As Integer, Optional button As ClickButton = ClickButton.Left)
            Dim lParam = NativeMethods.MakeLParam(clientX, clientY)
            If button = ClickButton.Right Then
                Dim wParam = New IntPtr(NativeMethods.MK_RBUTTON)
                NativeMethods.PostMessage(hWnd, NativeMethods.WM_RBUTTONDOWN, wParam, lParam)
                Thread.Sleep(10)
                NativeMethods.PostMessage(hWnd, NativeMethods.WM_RBUTTONUP, IntPtr.Zero, lParam)
            Else
                Dim wParam = New IntPtr(NativeMethods.MK_LBUTTON)
                NativeMethods.PostMessage(hWnd, NativeMethods.WM_LBUTTONDOWN, wParam, lParam)
                Thread.Sleep(10)
                NativeMethods.PostMessage(hWnd, NativeMethods.WM_LBUTTONUP, IntPtr.Zero, lParam)
            End If
        End Sub

        ''' <summary>
        ''' 창의 상대 좌표(GetWindowRect 기준)를 변환 후 클릭
        ''' CaptureWindow가 GetWindowRect(타이틀바+테두리 포함) 기준으로 캡처하므로,
        ''' 클릭 시 클라이언트 영역 오프셋을 보정해야 합니다.
        ''' </summary>
        Public Shared Sub ClickInWindow(hWnd As IntPtr, relativeX As Integer, relativeY As Integer, useBackground As Boolean, Optional button As ClickButton = ClickButton.Left)
            ' 창 프레임(타이틀바, 테두리) 오프셋 계산
            Dim rect As NativeMethods.RECT
            NativeMethods.GetWindowRect(hWnd, rect)
            Dim pt As NativeMethods.APIPOINT
            pt.X = 0 : pt.Y = 0
            NativeMethods.ClientToScreen(hWnd, pt)
            Dim frameLeft = pt.X - rect.Left
            Dim frameTop = pt.Y - rect.Top

            If useBackground Then
                ' PostMessage는 클라이언트 영역 좌표 사용
                Dim clientX = relativeX - frameLeft
                Dim clientY = relativeY - frameTop
                ClickAtWindow(hWnd, clientX, clientY, button)
            Else
                Dim screenX = rect.Left + relativeX
                Dim screenY = rect.Top + relativeY

                WindowFinder.BringToFront(hWnd)
                Thread.Sleep(100)
                ClickAtScreen(screenX, screenY, button)
            End If
        End Sub
    End Class
End Namespace
