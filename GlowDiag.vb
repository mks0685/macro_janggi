' 빛남 감지 진단 도구 v4 - V>220 기준
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports OpenCvSharp
Imports OpenCvSharp.Extensions
Imports MacroAutoControl.Constants
Imports MacroAutoControl.Engine
Imports MacroAutoControl.Capture

Namespace MacroAutoControl
    Public Module GlowDiag

        Public Sub RunDiagnostic()
            Dim log As New List(Of String)
            Dim baseDir = IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
            Dim outPath = IO.Path.Combine(baseDir, "glow_diag.txt")
            Dim imgPath = IO.Path.Combine(baseDir, "glow_diag.png")

            log.Add($"=== 빛남 감지 진단 v4 ({DateTime.Now:yyyy-MM-dd HH:mm:ss}) ===")
            log.Add($"기준: H:15~30, S>50, V>220")

            ' 1. 장기 창 찾기
            Dim windows = WindowFinder.GetAllVisibleWindows()
            Dim janggiWin = windows.FirstOrDefault(Function(w) w.Title.Contains("장기"))
            If janggiWin Is Nothing Then
                log.Add("[오류] 장기 창을 찾을 수 없습니다.")
                IO.File.WriteAllLines(outPath, log, System.Text.Encoding.UTF8)
                Return
            End If
            log.Add($"창: {janggiWin.Title}")

            ' 2. 캡처
            Dim screenshot = WindowFinder.CaptureWindow(janggiWin.Handle)
            If screenshot Is Nothing Then
                log.Add("[오류] 캡처 실패")
                IO.File.WriteAllLines(outPath, log, System.Text.Encoding.UTF8)
                Return
            End If
            screenshot.Save(imgPath, System.Drawing.Imaging.ImageFormat.Png)

            Dim mat = BitmapConverter.ToMat(screenshot)
            If mat.Channels() = 4 Then
                Dim bgr As New Mat()
                Cv2.CvtColor(mat, bgr, ColorConversionCodes.BGRA2BGR)
                mat.Dispose()
                mat = bgr
            End If

            ' 3. 보드 인식
            Dim recognizer As New BoardRecognizer()
            recognizer.LoadTemplates()
            Dim board = recognizer.GetBoardState(mat)
            If board Is Nothing Then
                log.Add("[오류] 보드 인식 실패")
                IO.File.WriteAllLines(outPath, log, System.Text.Encoding.UTF8)
                mat.Dispose() : screenshot.Dispose()
                Return
            End If

            Dim gridPos = recognizer.GetGridPositions()
            Dim cellSize = recognizer.GetCellSize()
            If gridPos Is Nothing Then
                log.Add("[오류] 그리드 좌표 없음")
                IO.File.WriteAllLines(outPath, log, System.Text.Encoding.UTF8)
                mat.Dispose() : screenshot.Dispose()
                Return
            End If

            Dim hsv As New Mat()
            Cv2.CvtColor(mat, hsv, ColorConversionCodes.BGR2HSV)

            Dim baseSize = Math.Max(cellSize.Item1, cellSize.Item2)
            Dim radii = New Integer() {CInt(baseSize * 0.55), CInt(baseSize * 0.65), CInt(baseSize * 0.75)}
            For ri = 0 To radii.Length - 1
                If radii(ri) < 24 Then radii(ri) = 24 + ri * 6
            Next
            log.Add($"반경: {radii(0)}, {radii(1)}, {radii(2)}")

            ' 4. 기물별 빛남 검사 (새 기준: H15~30, S>50, V>220)
            log.Add("")
            log.Add(String.Format("{0,-8} {1,-6} {2,7} {3}", "위치", "기물", "빛남수", "판정 + V>220 픽셀"))
            log.Add(New String("-"c, 80))

            For r = 0 To BOARD_ROWS - 1
                For c = 0 To BOARD_COLS - 1
                    Dim piece = board.Grid(r)(c)
                    If piece = EMPTY Then Continue For

                    Dim idx = r * BOARD_COLS + c
                    Dim cx = gridPos(idx)(0), cy = gridPos(idx)(1)
                    Dim glowCount = 0
                    Dim glowPixels As New List(Of String)

                    For Each radius In radii
                        Dim cos45v = CInt(radius * 0.707)
                        Dim dx = New Integer() {radius, cos45v, 0, -cos45v, -radius, -cos45v, 0, cos45v}
                        Dim dy = New Integer() {0, cos45v, radius, cos45v, 0, -cos45v, -radius, -cos45v}
                        For i = 0 To 7
                            Dim px = cx + dx(i), py = cy + dy(i)
                            If px < 0 OrElse px >= hsv.Cols OrElse py < 0 OrElse py >= hsv.Rows Then Continue For
                            Dim pixel(2) As Byte
                            Marshal.Copy(hsv.Data + (py * CInt(hsv.Step()) + px * 3), pixel, 0, 3)
                            Dim h = CInt(pixel(0)), s = CInt(pixel(1)), v = CInt(pixel(2))
                            If h >= 15 AndAlso h <= 30 AndAlso s > 50 AndAlso v > 220 Then
                                glowCount += 1
                                glowPixels.Add($"H{h}S{s}V{v}")
                            End If
                        Next
                    Next

                    Dim pieceName = ""
                    BoardRecognizer.PIECE_NAMES.TryGetValue(piece, pieceName)
                    Dim verdict = If(glowCount >= 3, " ★빛남", "")
                    Dim pixelStr = If(glowPixels.Count > 0, String.Join(" ", glowPixels), "-")
                    log.Add(String.Format("({0},{1})  {2,-6} {3,5}  {4} {5}", r, c, pieceName, glowCount, verdict, pixelStr))
                Next
            Next

            hsv.Dispose()
            mat.Dispose()
            screenshot.Dispose()

            IO.File.WriteAllLines(outPath, log, System.Text.Encoding.UTF8)
        End Sub

    End Module
End Namespace
