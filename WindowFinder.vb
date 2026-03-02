Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Text

Namespace MacroAutoControl
    ''' <summary>
    ''' 카카오톡 장기 게임 창을 찾는 클래스
    ''' </summary>
    Public Class WindowFinder
        ' DPI 관련 API
        <DllImport("user32.dll")>
        Private Shared Function SetProcessDPIAware() As Boolean
        End Function

        <DllImport("user32.dll")>
        Private Shared Function GetDC(hWnd As IntPtr) As IntPtr
        End Function

        <DllImport("user32.dll")>
        Private Shared Function ReleaseDC(hWnd As IntPtr, hDC As IntPtr) As Integer
        End Function

        <DllImport("gdi32.dll")>
        Private Shared Function BitBlt(hdcDest As IntPtr, xDest As Integer, yDest As Integer,
                                        wDest As Integer, hDest As Integer,
                                        hdcSrc As IntPtr, xSrc As Integer, ySrc As Integer,
                                        rop As Integer) As Boolean
        End Function

        Private Const SRCCOPY As Integer = &HCC0020

        ' PrintWindow 플래그
        Private Const PW_CLIENTONLY As UInteger = 1
        Private Const PW_RENDERFULLCONTENT As UInteger = 2

        Shared Sub New()
            ' DPI 인식 활성화
            Try
                SetProcessDPIAware()
            Catch
            End Try
        End Sub

        Public Class WindowInfo
            Public Property Handle As IntPtr
            Public Property Title As String
            Public Property Bounds As Rectangle
        End Class

        ''' <summary>
        ''' 제목에 특정 키워드가 포함된 창 목록을 반환
        ''' </summary>
        Public Shared Function FindWindowsByTitle(keywords As String()) As List(Of WindowInfo)
            Dim results As New List(Of WindowInfo)

            NativeMethods.EnumWindows(
                Function(hWnd, lParam)
                    If Not NativeMethods.IsWindowVisible(hWnd) Then Return True

                    Dim length = NativeMethods.GetWindowTextLength(hWnd)
                    If length = 0 Then Return True

                    Dim sb As New StringBuilder(length + 1)
                    NativeMethods.GetWindowText(hWnd, sb, sb.Capacity)
                    Dim title = sb.ToString()

                    For Each keyword In keywords
                        If title.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 Then
                            Dim rect As NativeMethods.RECT
                            NativeMethods.GetWindowRect(hWnd, rect)

                            results.Add(New WindowInfo() With {
                                .Handle = hWnd,
                                .Title = title,
                                .Bounds = New Rectangle(
                                    rect.Left, rect.Top,
                                    rect.Right - rect.Left,
                                    rect.Bottom - rect.Top)
                            })
                            Exit For
                        End If
                    Next

                    Return True
                End Function,
                IntPtr.Zero)

            Return results
        End Function

        ''' <summary>
        ''' 카카오 장기 관련 창 찾기
        ''' </summary>
        Public Shared Function FindJanggiWindow() As WindowInfo
            Dim keywords = {"장기", "Janggi", "janggi", "KakaoGame", "한게임"}
            Dim windows = FindWindowsByTitle(keywords)

            If windows.Count = 0 Then Return Nothing
            If windows.Count = 1 Then Return windows(0)

            Return windows.OrderByDescending(Function(w) w.Bounds.Width * w.Bounds.Height).First()
        End Function

        ''' <summary>
        ''' 모든 보이는 창 목록 반환
        ''' </summary>
        Public Shared Function GetAllVisibleWindows() As List(Of WindowInfo)
            Dim results As New List(Of WindowInfo)

            NativeMethods.EnumWindows(
                Function(hWnd, lParam)
                    If Not NativeMethods.IsWindowVisible(hWnd) Then Return True

                    Dim length = NativeMethods.GetWindowTextLength(hWnd)
                    If length = 0 Then Return True

                    Dim sb As New StringBuilder(length + 1)
                    NativeMethods.GetWindowText(hWnd, sb, sb.Capacity)
                    Dim title = sb.ToString()

                    If String.IsNullOrWhiteSpace(title) Then Return True

                    Dim rect As NativeMethods.RECT
                    NativeMethods.GetWindowRect(hWnd, rect)

                    Dim w = rect.Right - rect.Left
                    Dim h = rect.Bottom - rect.Top
                    ' 너무 작은 창 제외
                    If w < 50 OrElse h < 50 Then Return True

                    results.Add(New WindowInfo() With {
                        .Handle = hWnd,
                        .Title = title,
                        .Bounds = New Rectangle(rect.Left, rect.Top, w, h)
                    })

                    Return True
                End Function,
                IntPtr.Zero)

            Return results
        End Function

        ''' <summary>
        ''' 창을 전면으로 가져오기
        ''' </summary>
        Public Shared Sub BringToFront(hWnd As IntPtr)
            If NativeMethods.IsIconic(hWnd) Then
                NativeMethods.ShowWindow(hWnd, NativeMethods.SW_RESTORE)
            End If
            NativeMethods.SetForegroundWindow(hWnd)
        End Sub

        ''' <summary>
        ''' 창 캡처 - 여러 방법을 시도
        ''' </summary>
        Public Shared Function CaptureWindow(hWnd As IntPtr) As Bitmap
            Dim rect As NativeMethods.RECT
            NativeMethods.GetWindowRect(hWnd, rect)

            Dim width = rect.Right - rect.Left
            Dim height = rect.Bottom - rect.Top

            If width <= 0 OrElse height <= 0 Then Return Nothing

            ' 방법 1: PrintWindow (PW_RENDERFULLCONTENT)
            Dim bmp = TryPrintWindow(hWnd, width, height, PW_RENDERFULLCONTENT)
            If bmp IsNot Nothing AndAlso Not IsBlackImage(bmp) Then Return bmp
            bmp?.Dispose()

            ' 방법 2: PrintWindow (기본)
            bmp = TryPrintWindow(hWnd, width, height, 0)
            If bmp IsNot Nothing AndAlso Not IsBlackImage(bmp) Then Return bmp
            bmp?.Dispose()

            ' 방법 3: PrintWindow (클라이언트 영역만)
            bmp = TryPrintWindow(hWnd, width, height, PW_CLIENTONLY)
            If bmp IsNot Nothing AndAlso Not IsBlackImage(bmp) Then Return bmp
            bmp?.Dispose()

            ' 방법 4: BitBlt (DC 복사)
            bmp = TryBitBlt(hWnd, width, height)
            If bmp IsNot Nothing AndAlso Not IsBlackImage(bmp) Then Return bmp
            bmp?.Dispose()

            ' 방법 5: CopyFromScreen (화면에서 직접 캡처 - 창이 앞에 있어야 함)
            BringToFront(hWnd)
            Threading.Thread.Sleep(500)
            ' 좌표 다시 가져오기 (Restore로 위치가 바뀔 수 있음)
            NativeMethods.GetWindowRect(hWnd, rect)
            width = rect.Right - rect.Left
            height = rect.Bottom - rect.Top

            bmp = New Bitmap(width, height)
            Using g = Graphics.FromImage(bmp)
                g.CopyFromScreen(rect.Left, rect.Top, 0, 0, New Size(width, height), CopyPixelOperation.SourceCopy)
            End Using
            Return bmp
        End Function

        ''' <summary>
        ''' PrintWindow 시도
        ''' </summary>
        Private Shared Function TryPrintWindow(hWnd As IntPtr, width As Integer, height As Integer, flags As UInteger) As Bitmap
            Try
                Dim bmp As New Bitmap(width, height)
                Using g = Graphics.FromImage(bmp)
                    Dim hdc = g.GetHdc()
                    Dim success = NativeMethods.PrintWindow(hWnd, hdc, flags)
                    g.ReleaseHdc(hdc)
                    If success Then Return bmp
                End Using
                bmp.Dispose()
                Return Nothing
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' BitBlt 시도
        ''' </summary>
        Private Shared Function TryBitBlt(hWnd As IntPtr, width As Integer, height As Integer) As Bitmap
            Try
                Dim srcDC = GetDC(hWnd)
                If srcDC = IntPtr.Zero Then Return Nothing

                Dim bmp As New Bitmap(width, height)
                Using g = Graphics.FromImage(bmp)
                    Dim destDC = g.GetHdc()
                    BitBlt(destDC, 0, 0, width, height, srcDC, 0, 0, SRCCOPY)
                    g.ReleaseHdc(destDC)
                End Using
                ReleaseDC(hWnd, srcDC)
                Return bmp
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 이미지가 전부 검은색인지 확인
        ''' </summary>
        Private Shared Function IsBlackImage(bmp As Bitmap) As Boolean
            ' 샘플링으로 빠르게 확인
            Dim data = bmp.LockBits(
                New Rectangle(0, 0, bmp.Width, bmp.Height),
                Imaging.ImageLockMode.ReadOnly,
                Imaging.PixelFormat.Format32bppArgb)

            Dim bytes(data.Stride * data.Height - 1) As Byte
            Runtime.InteropServices.Marshal.Copy(data.Scan0, bytes, 0, bytes.Length)
            bmp.UnlockBits(data)

            Dim stepX = Math.Max(1, bmp.Width \ 10)
            Dim stepY = Math.Max(1, bmp.Height \ 10)
            Dim nonBlackCount = 0

            For y = 0 To bmp.Height - 1 Step stepY
                For x = 0 To bmp.Width - 1 Step stepX
                    Dim idx = y * data.Stride + x * 4
                    If bytes(idx) > 10 OrElse bytes(idx + 1) > 10 OrElse bytes(idx + 2) > 10 Then
                        nonBlackCount += 1
                        If nonBlackCount > 5 Then Return False
                    End If
                Next
            Next

            Return True
        End Function

        ''' <summary>
        ''' 화면 전체 캡처
        ''' </summary>
        Public Shared Function CaptureFullScreen() As Bitmap
            Dim bounds = System.Windows.Forms.Screen.PrimaryScreen.Bounds
            Dim bmp As New Bitmap(bounds.Width, bounds.Height)
            Using g = Graphics.FromImage(bmp)
                g.CopyFromScreen(bounds.Left, bounds.Top, 0, 0, bounds.Size, CopyPixelOperation.SourceCopy)
            End Using
            Return bmp
        End Function
    End Class
End Namespace
