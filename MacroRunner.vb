Imports System.Drawing
Imports System.Threading
Imports System.Windows.Forms
Imports OpenCvSharp
Imports OpenCvSharp.Extensions
Imports MacroAutoControl.Constants
Imports MacroAutoControl.Engine
Imports MacroAutoControl.AI
Imports System.Runtime.InteropServices
Imports MacroAutoControl.Capture

Namespace MacroAutoControl
    ''' <summary>
    ''' 매크로 시퀀스 실행 엔진
    ''' </summary>
    Public Class MacroRunner
        Public Event ProgressChanged(currentIndex As Integer, totalCount As Integer, itemName As String, status As String, currentRepeat As Integer, totalRepeat As Integer)
        Public Event MacroCompleted(success As Boolean, message As String)
        Public Event AIMoveVisualize(screenshot As Bitmap, fromX As Integer, fromY As Integer, toX As Integer, toY As Integer, moveInfo As String)

        Private _cancelRequested As Boolean
        Private _isRunning As Boolean

        ' 현재 실행 중인 AI 항목 (깊이 조절용)
        Private _currentAIItem As MacroItem = Nothing

        ' AI 보드 인식기 (재사용)
        Private _recognizer As BoardRecognizer = Nothing

        ' 게임 종료 사유 (팝업 감지 시 설정)
        Private _gameEndReason As String = Nothing

        ' 내 차례 스크린샷 저장
        Private _watchDir As String = Nothing
        Private _watchCounter As Integer = 0

        ' 게임 결과 팝업 템플릿
        Private Shared ReadOnly TEMPLATES_DIR As String =
            IO.Path.Combine(IO.Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().Location), "templates")

        Private Shared ReadOnly RESULT_TEMPLATES As String()() = {
            New String() {"result_invalid.png", "무효대국"},
            New String() {"result_win.png", "승리"},
            New String() {"result_lose.png", "패배"},
            New String() {"result_timeout_win.png", "시간승"},
            New String() {"result_timeout_lose.png", "시간패"},
            New String() {"result_resign_win.png", "기권승"},
            New String() {"result_resign_lose.png", "기권패"}
        }

        Public ReadOnly Property IsRunning As Boolean
            Get
                Return _isRunning
            End Get
        End Property

        Public Sub Cancel()
            _cancelRequested = True
            Search.CancelSearch()
        End Sub

        ''' <summary>
        ''' AI 탐색 깊이 조절 (+1 또는 -1). 실행 중인 AI 항목에 즉시 반영.
        ''' </summary>
        Public Sub AdjustDepth(delta As Integer)
            Dim item = _currentAIItem
            If item Is Nothing Then Return
            Dim newDepth = Math.Max(1, Math.Min(10, item.AIDepth + delta))
            item.AIDepth = newDepth
        End Sub

        Public ReadOnly Property CurrentAIDepth As Integer
            Get
                Dim item = _currentAIItem
                If item Is Nothing Then Return 0
                Return item.AIDepth
            End Get
        End Property

        Private Function GetRecognizer() As BoardRecognizer
            If _recognizer Is Nothing Then
                _recognizer = New BoardRecognizer()
                _recognizer.LoadTemplates()
            End If
            Return _recognizer
        End Function

        ''' <summary>
        ''' 내 기물 주변에 빛남(하이라이트) 효과가 있는지 감지
        ''' 내 차례이면 기물 주변이 빛남 → True, 상대방 차례이면 빛남 없음 → False
        ''' </summary>
        ''' <summary>
        ''' Mat에서 특정 픽셀의 최대 밝기(B,G,R 중 최댓값) 반환
        ''' </summary>
        Private Shared Function GetPixelBrightness(mat As Mat, x As Integer, y As Integer) As Integer
            If x < 0 OrElse x >= mat.Cols OrElse y < 0 OrElse y >= mat.Rows Then Return 0
            Dim pixel(2) As Byte
            Marshal.Copy(mat.Data + (y * CInt(mat.Step()) + x * 3), pixel, 0, 3)
            Return Math.Max(Math.Max(CInt(pixel(0)), CInt(pixel(1))), CInt(pixel(2)))
        End Function

        ''' <summary>
        ''' HSV 기반 기물 주변 빛남 픽셀 수 측정
        ''' 보드 바닥색: H20~24, S90~135, V200~218
        ''' 빛남 효과: 보드보다 밝은 노란/금색 (H:15~30, S>50, V>220)
        ''' </summary>
        Private Function CountGlowPixels(hsv As Mat, cx As Integer, cy As Integer, radii As Integer()) As Integer
            Dim glowCount = 0
            For Each radius In radii
                Dim cos45 = CInt(radius * 0.707)
                Dim dx = New Integer() {radius, cos45, 0, -cos45, -radius, -cos45, 0, cos45}
                Dim dy = New Integer() {0, cos45, radius, cos45, 0, -cos45, -radius, -cos45}
                For i = 0 To 7
                    Dim px = cx + dx(i)
                    Dim py = cy + dy(i)
                    If px < 0 OrElse px >= hsv.Cols OrElse py < 0 OrElse py >= hsv.Rows Then Continue For
                    Dim pixel(2) As Byte
                    Marshal.Copy(hsv.Data + (py * CInt(hsv.Step()) + px * 3), pixel, 0, 3)
                    Dim h = CInt(pixel(0)), s = CInt(pixel(1)), v = CInt(pixel(2))
                    ' 빛남 효과: 노란~금색 (H:15~30) + 채도 있음 + 보드보다 밝음 (V>220)
                    If h >= 15 AndAlso h <= 30 AndAlso s > 50 AndAlso v > 220 Then
                        glowCount += 1
                    End If
                Next
            Next
            Return glowCount
        End Function

        Private Function HasGlowAroundMyPieces(mat As Mat, board As Board, gridPos As Integer()(), mySide As String, cellW As Integer, cellH As Integer, boardFlipped As Boolean, Optional ByRef glowPieceName As String = Nothing) As Boolean
            Dim myPrefix = If(mySide = Constants.CHO, "C", "H")
            Dim enemyPrefix = If(mySide = Constants.CHO, "H", "C")

            ' 기물 바깥쪽 빛남 감지: 셀 크기의 55%~75% 범위
            Dim baseSize = Math.Max(cellW, cellH)
            Dim radii = New Integer() {
                CInt(baseSize * 0.55),
                CInt(baseSize * 0.65),
                CInt(baseSize * 0.75)
            }
            For ri = 0 To radii.Length - 1
                If radii(ri) < 24 Then radii(ri) = 24 + ri * 6
            Next

            ' BGR → HSV 변환
            Dim hsv As New Mat()
            Cv2.CvtColor(mat, hsv, ColorConversionCodes.BGR2HSV)

            ' 각 기물의 빛남 픽셀 수 측정
            Dim myMaxGlow = 0, myMaxPiece = ""
            Dim enemyMaxGlow = 0, enemyMaxPiece = ""

            For r = 0 To BOARD_ROWS - 1
                For c = 0 To BOARD_COLS - 1
                    Dim piece = board.Grid(r)(c)
                    If piece = Constants.EMPTY Then Continue For
                    Dim screenR = If(boardFlipped, 9 - r, r)
                    Dim idx = screenR * BOARD_COLS + c
                    Dim cx = gridPos(idx)(0)
                    Dim cy = gridPos(idx)(1)
                    Dim glow = CountGlowPixels(hsv, cx, cy, radii)

                    If piece.StartsWith(myPrefix) Then
                        If glow > myMaxGlow Then myMaxGlow = glow : myMaxPiece = $"{piece}({r},{c})"
                    ElseIf piece.StartsWith(enemyPrefix) Then
                        If glow > enemyMaxGlow Then enemyMaxGlow = glow : enemyMaxPiece = $"{piece}({r},{c})"
                    End If
                Next
            Next
            hsv.Dispose()

            ' 빛남 판정 기준: 24개 샘플 중 3개 이상이면 빛남
            Dim glowMin = 3

            ' 1. 내 기물 중 빛남 있으면 → 상대 차례 (대기)
            If myMaxGlow >= glowMin Then
                Return False
            End If

            ' 2. 상대 기물 중 빛남 있으면 → 내 차례
            If enemyMaxGlow >= glowMin Then
                glowPieceName = $"상대빛남:{enemyMaxPiece} glow={enemyMaxGlow}"
                Return True
            End If

            ' 3. 양쪽 모두 빛남 없음 → 내 차례 (첫 수 등)
            glowPieceName = $"양쪽무빛 my={myMaxGlow} enemy={enemyMaxGlow}"
            Return True
        End Function

        ''' <summary>
        ''' 캡처 화면에서 게임 결과 팝업(무효대국/승/패 등)을 감지
        ''' </summary>
        Private Function DetectGameResult(screenshot As Bitmap) As String
            For Each tpl In RESULT_TEMPLATES
                Dim templatePath = IO.Path.Combine(TEMPLATES_DIR, tpl(0))
                If Not IO.File.Exists(templatePath) Then Continue For

                Dim result = ButtonFinder.FindByTemplate(screenshot, templatePath, 0.75)
                If result.Found Then
                    Return tpl(1)
                End If
            Next
            Return Nothing
        End Function

        ''' <summary>
        ''' 내 차례 감지: 내 기물 주변에 빛남 효과가 나타날 때까지 대기
        ''' </summary>
        Private Function WaitForMyTurn(item As MacroItem,
                                       targetWindow As WindowFinder.WindowInfo,
                                       ownerForm As Form,
                                       index As Integer, totalCount As Integer,
                                       currentRepeat As Integer, totalRepeat As Integer) As Board
            Dim recognizer = GetRecognizer()
            Dim pollCount = 0

            While Not _cancelRequested
                pollCount += 1
                RaiseEvent ProgressChanged(index + 1, totalCount, item.Name,
                    $"AI: 내 차례 대기 중... ({pollCount})", currentRepeat, totalRepeat)
                Application.DoEvents()

                ' 1초 대기 (100ms 단위로 취소 체크)
                Dim waited = 0
                While waited < 1000 AndAlso Not _cancelRequested
                    Thread.Sleep(100)
                    waited += 100
                    Application.DoEvents()
                End While

                If _cancelRequested Then Return Nothing

                ' 캡처 (PrintWindow는 숨김 불필요)
                Dim screenshot = WindowFinder.CaptureWindow(targetWindow.Handle)

                If screenshot Is Nothing Then Continue While

                ' 게임 결과 팝업 감지
                Dim gameResult = DetectGameResult(screenshot)
                If gameResult IsNot Nothing Then
                    screenshot.Dispose()
                    _gameEndReason = gameResult
                    Return Nothing
                End If

                ' Mat 변환
                Dim mat As Mat = Nothing
                Try
                    mat = BitmapConverter.ToMat(screenshot)
                    If mat.Channels() = 4 Then
                        Dim bgr As New Mat()
                        Cv2.CvtColor(mat, bgr, ColorConversionCodes.BGRA2BGR)
                        mat.Dispose()
                        mat = bgr
                    End If
                Catch ex As Exception
                    screenshot.Dispose()
                    Continue While
                End Try

                ' 보드 인식
                Dim board = recognizer.GetBoardState(mat)

                If board Is Nothing Then
                    mat.Dispose()
                    screenshot.Dispose()
                    Continue While
                End If

                ' 진영 자동 감지
                ' 보드가 뒤집혔으면 원래 화면에서 한(HAN)이 아래에 있었으므로 내 진영 = 한
                Dim mySide = item.AISide
                If mySide = "AUTO" Then
                    If recognizer.IsBoardFlipped Then
                        mySide = Constants.HAN
                    Else
                        For r = 7 To 9
                            For c = 3 To 5
                                Dim p = board.Grid(r)(c)
                                If p = Constants.CK Then mySide = Constants.CHO : Exit For
                                If p = Constants.HK Then mySide = Constants.HAN : Exit For
                            Next
                            If mySide <> "AUTO" Then Exit For
                        Next
                    End If
                    If mySide = "AUTO" Then mySide = Constants.CHO
                End If
                Dim dbgSide = If(mySide = Constants.CHO, "초", "한")
                RaiseEvent ProgressChanged(index + 1, totalCount, item.Name,
                    $"AI: 진영={dbgSide} 대기 중... ({pollCount})", currentRepeat, totalRepeat)
                Application.DoEvents()

                ' 빛남 감지
                Dim gridPos = recognizer.GetGridPositions()
                Dim cellSize = recognizer.GetCellSize()
                Dim glowPiece As String = Nothing
                If gridPos IsNot Nothing AndAlso HasGlowAroundMyPieces(mat, board, gridPos, mySide, cellSize.Item1, cellSize.Item2, recognizer.IsBoardFlipped, glowPiece) Then
                    ' 내 차례 감지! 스크린샷 저장
                    If _watchDir IsNot Nothing Then
                        Try
                            _watchCounter += 1
                            Dim savePath = IO.Path.Combine(_watchDir, $"{_watchCounter.ToString("D4")}.jpg")
                            screenshot.Save(savePath, System.Drawing.Imaging.ImageFormat.Jpeg)
                        Catch
                        End Try
                    End If

                    ' 상대 수 애니메이션 완료 대기 후 재캡처
                    mat.Dispose()
                    screenshot.Dispose()

                    Dim sideDbg = If(mySide = Constants.CHO, "초", "한")
                    RaiseEvent ProgressChanged(index + 1, totalCount, item.Name,
                        $"AI: 내 차례 감지 ({sideDbg}, 빛남:{glowPiece}) → 대기...", currentRepeat, totalRepeat)
                    Application.DoEvents()

                    ' 1.5초 대기 (상대 기물 이동 애니메이션 완료)
                    Dim w2 = 0
                    While w2 < 1500 AndAlso Not _cancelRequested
                        Thread.Sleep(100)
                        w2 += 100
                        Application.DoEvents()
                    End While
                    If _cancelRequested Then Return Nothing

                    ' 재캡처 + 재인식 → 0.5초 후 한번 더 확인
                    Dim finalBoard As Board = Nothing
                    For confirm = 1 To 2
                        Dim ss2 = WindowFinder.CaptureWindow(targetWindow.Handle)
                        If ss2 Is Nothing Then Continue For

                        Dim mat2 As Mat = Nothing
                        Try
                            mat2 = BitmapConverter.ToMat(ss2)
                            If mat2.Channels() = 4 Then
                                Dim bgr2 As New Mat()
                                Cv2.CvtColor(mat2, bgr2, ColorConversionCodes.BGRA2BGR)
                                mat2.Dispose()
                                mat2 = bgr2
                            End If
                        Catch
                            ss2.Dispose()
                            Continue For
                        End Try

                        finalBoard = recognizer.GetBoardState(mat2)
                        mat2.Dispose()
                        ss2.Dispose()

                        If confirm = 1 AndAlso finalBoard IsNot Nothing Then
                            ' 0.5초 후 한번 더 확인
                            RaiseEvent ProgressChanged(index + 1, totalCount, item.Name,
                                $"AI: 상대 수 확인 중...", currentRepeat, totalRepeat)
                            Application.DoEvents()
                            Dim w3 = 0
                            While w3 < 500 AndAlso Not _cancelRequested
                                Thread.Sleep(100)
                                w3 += 100
                                Application.DoEvents()
                            End While
                            If _cancelRequested Then Return Nothing
                        End If
                    Next

                    If finalBoard IsNot Nothing Then
                        Return finalBoard
                    Else
                        Return board
                    End If
                End If

                mat.Dispose()
                screenshot.Dispose()
            End While

            Return Nothing
        End Function

        ''' <summary>
        ''' 매크로 리스트를 순차 실행 (repeatCount=0이면 무한 반복)
        ''' </summary>
        Public Sub RunSequence(items As List(Of MacroItem),
                               targetWindow As WindowFinder.WindowInfo,
                               ownerForm As Form,
                               useBackground As Boolean,
                               Optional repeatCount As Integer = 1)
            If items.Count = 0 Then
                RaiseEvent MacroCompleted(False, "매크로 항목이 없습니다.")
                Return
            End If

            If targetWindow Is Nothing Then
                RaiseEvent MacroCompleted(False, "대상 창이 선택되지 않았습니다.")
                Return
            End If

            _cancelRequested = False
            _isRunning = True

            _gameEndReason = Nothing

            ' 내 차례 스크린샷 저장 폴더 초기화
            _watchCounter = 0
            Dim exeDir = IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
            ' 기존 watch_ 폴더 중 마지막 번호 찾기
            Dim dateStr = DateTime.Now.ToString("yyyy-MM-dd")
            Dim watchBase = $"watch_{dateStr}"
            Dim maxNum = 0
            If IO.Directory.Exists(exeDir) Then
                For Each d In IO.Directory.GetDirectories(exeDir, $"{watchBase}_*")
                    Dim name = IO.Path.GetFileName(d)
                    Dim parts = name.Split("_"c)
                    Dim num As Integer
                    If parts.Length > 0 AndAlso Integer.TryParse(parts(parts.Length - 1), num) Then
                        If num > maxNum Then maxNum = num
                    End If
                Next
            End If
            _watchDir = IO.Path.Combine(exeDir, $"{watchBase}_{(maxNum + 1).ToString("D3")}")
            IO.Directory.CreateDirectory(_watchDir)
            Console.WriteLine($"[알림] 스크린샷 저장 폴더: {_watchDir}")

            ' 마우스 원래 위치 저장
            Dim savedCursorPos As NativeMethods.APIPOINT
            NativeMethods.GetCursorPos(savedCursorPos)

            Dim infinite = (repeatCount = 0)
            Dim totalRepeat = If(infinite, 0, repeatCount)
            Dim currentRepeat = 0

            ' 첫 AI 항목 인덱스 찾기 (반복 재시작 시 AI부터 시작)
            Dim firstAIIndex = 0
            For idx = 0 To items.Count - 1
                If items(idx).IsAI Then firstAIIndex = idx : Exit For
            Next
            Dim startIndex = 0  ' 첫 실행은 처음부터

            Try
                While infinite OrElse currentRepeat < totalRepeat
                    currentRepeat += 1

                    For i = startIndex To items.Count - 1
                        If _cancelRequested Then
                            RaiseEvent MacroCompleted(False, $"매크로 중지됨 (반복 {currentRepeat}, {i}/{items.Count})")
                            Return
                        End If

                        Dim item = items(i)

                        If item.IsAI Then
                            ' AI 항목 실행
                            Dim gameEnded = ExecuteAIItem(item, i, items.Count, targetWindow, ownerForm, useBackground, currentRepeat, totalRepeat)
                            If _cancelRequested Then
                                RaiseEvent MacroCompleted(False, $"매크로 중지됨 (반복 {currentRepeat}, AI)")
                                Return
                            End If

                            If gameEnded AndAlso infinite Then
                                ' 게임 종료 → 그리드 초기화 후 매크로 처음부터 다시 시작
                    
                                _gameEndReason = Nothing
                                _recognizer?.ResetGrid()
                                RaiseEvent ProgressChanged(0, items.Count, "", $"게임 종료 → 재시작 대기 (3초)...", currentRepeat, totalRepeat)
                                Application.DoEvents()
                                Dim w = 0
                                While w < 3000 AndAlso Not _cancelRequested
                                    Thread.Sleep(100)
                                    w += 100
                                    Application.DoEvents()
                                End While
                                startIndex = 0  ' 게임 종료 시 매크로 처음부터
                                Exit For  ' For 루프 탈출 → While에서 다시
                            ElseIf gameEnded Then
                                RaiseEvent MacroCompleted(True, $"게임 종료 ({currentRepeat}회 실행)")
                                Return
                            End If
                        ElseIf item.Threshold > 0 Then
                            ' 이미지 찾기+클릭 항목 (실패 시 다음 항목으로)
                            RaiseEvent ProgressChanged(i + 1, items.Count, item.Name, "찾는 중...", currentRepeat, totalRepeat)
                            Application.DoEvents()

                            Dim screenshot = WindowFinder.CaptureWindow(targetWindow.Handle)

                            If screenshot IsNot Nothing AndAlso Not _cancelRequested Then
                                Dim result = ButtonFinder.FindByTemplate(screenshot, item.Image, item.Threshold)
                                screenshot.Dispose()

                                If result.Found Then
                                    RaiseEvent ProgressChanged(i + 1, items.Count, item.Name, "클릭 중...", currentRepeat, totalRepeat)
                                    Application.DoEvents()

                                    Dim clickX As Integer
                                    Dim clickY As Integer
                                    If item.ClickOffsetX >= 0 Then
                                        clickX = result.MatchRect.X + item.ClickOffsetX
                                        clickY = result.MatchRect.Y + item.ClickOffsetY
                                    Else
                                        clickX = result.Location.X
                                        clickY = result.Location.Y
                                    End If

                                    ButtonClicker.ClickInWindow(
                                        targetWindow.Handle,
                                        clickX,
                                        clickY,
                                        useBackground,
                                        item.Button)
                                Else
                                    RaiseEvent ProgressChanged(i + 1, items.Count, item.Name, $"못 찾음 ({result.Confidence:P0}) → 다음", currentRepeat, totalRepeat)
                                    Application.DoEvents()
                                End If
                            Else
                                screenshot?.Dispose()
                            End If
                        End If

                        ' 대기 (AI 항목은 ExecuteAIItem 내부에서 처리)
                        If item.DelayAfterClick > 0 AndAlso Not item.IsAI Then
                            RaiseEvent ProgressChanged(i + 1, items.Count, item.Name, $"대기 중 ({item.DelayAfterClick}ms)...", currentRepeat, totalRepeat)
                            Application.DoEvents()

                            Dim waited = 0
                            While waited < item.DelayAfterClick AndAlso Not _cancelRequested
                                Dim chunk = Math.Min(100, item.DelayAfterClick - waited)
                                Thread.Sleep(chunk)
                                waited += chunk
                                Application.DoEvents()
                            End While
                        End If

                        ' 키보드 전송
                        If Not String.IsNullOrEmpty(item.SendKeys) AndAlso Not _cancelRequested Then
                            RaiseEvent ProgressChanged(i + 1, items.Count, item.Name, $"키 전송: {item.SendKeys}", currentRepeat, totalRepeat)
                            Application.DoEvents()

                            SendKeysToWindow(targetWindow.Handle, item.SendKeys, useBackground)
                            Thread.Sleep(100)
                        End If
                    Next

                    ' 다음 반복 준비
                    If (infinite OrElse currentRepeat < totalRepeat) AndAlso Not _cancelRequested Then
                        startIndex = firstAIIndex  ' 다음 반복은 AI부터
                    End If
                End While

                RaiseEvent MacroCompleted(True, $"매크로 완료! ({items.Count}개 항목 x {currentRepeat}회 실행)")

            Finally
                _isRunning = False
                _currentAIItem = Nothing
                ' 마우스 원래 위치로 복귀
                NativeMethods.SetCursorPos(savedCursorPos.X, savedCursorPos.Y)
                ' 취소로 인한 종료 시에도 UI 복원 보장
                If _cancelRequested Then
                    RaiseEvent MacroCompleted(False, "매크로 중지됨")
                End If
            End Try
        End Sub

        ''' <summary>
        ''' AI 항목 실행: 캡처 → 보드 인식 → AI 탐색 → 출발지/도착지 클릭
        ''' 반환값: True=게임 종료 (무한반복 시 다시 시작 필요)
        ''' </summary>
        Private Function ExecuteAIItem(item As MacroItem, index As Integer, totalCount As Integer,
                                   targetWindow As WindowFinder.WindowInfo,
                                   ownerForm As Form, useBackground As Boolean,
                                   currentRepeat As Integer, totalRepeat As Integer) As Boolean

            _currentAIItem = item
            Dim board As Board = Nothing
            Dim recognizer = GetRecognizer()

            ' 0. 내 차례 감지: 기물 주변 빛남이 나타날 때까지 대기
            board = WaitForMyTurn(item, targetWindow, ownerForm, index, totalCount,
                                 currentRepeat, totalRepeat)
            If board Is Nothing Then
                If _gameEndReason IsNot Nothing Then
                    RaiseEvent ProgressChanged(index + 1, totalCount, item.Name, $"게임 종료: {_gameEndReason}", currentRepeat, totalRepeat)
                    Return True  ' 게임 종료
                End If
                Return False
            End If

            ' 시각화용 캡처 (WaitForMyTurn 후에도 현재 화면 캡처)
            Dim capturedBmp As Bitmap = WindowFinder.CaptureWindow(targetWindow.Handle)
            Dim retryCount = 0
            While board Is Nothing AndAlso Not _cancelRequested
                retryCount += 1
                If retryCount > 10 Then
                    RaiseEvent MacroCompleted(False, $"AI 보드 인식 실패: 10회 재시도 후 중단")
                    Return False
                End If

                If retryCount > 1 Then
                    RaiseEvent ProgressChanged(index + 1, totalCount, item.Name, $"AI: 재시도 {retryCount}/10...", currentRepeat, totalRepeat)
                    Application.DoEvents()
                    Thread.Sleep(1000)
                Else
                    RaiseEvent ProgressChanged(index + 1, totalCount, item.Name, "AI: 캡처 중...", currentRepeat, totalRepeat)
                    Application.DoEvents()
                End If

                Dim screenshot = WindowFinder.CaptureWindow(targetWindow.Handle)
                If screenshot Is Nothing Then Continue While

                If _cancelRequested Then
                    screenshot.Dispose()
                    Return False
                End If

                ' 게임 결과 팝업 감지
                Dim gameResult = DetectGameResult(screenshot)
                If gameResult IsNot Nothing Then
                    screenshot.Dispose()
                    RaiseEvent ProgressChanged(index + 1, totalCount, item.Name, $"게임 종료: {gameResult}", currentRepeat, totalRepeat)
                    Return True  ' 게임 종료
                End If

                Dim mat As Mat = Nothing
                Try
                    mat = BitmapConverter.ToMat(screenshot)
                    If mat.Channels() = 4 Then
                        Dim bgr As New Mat()
                        Cv2.CvtColor(mat, bgr, ColorConversionCodes.BGRA2BGR)
                        mat.Dispose()
                        mat = bgr
                    End If
                Catch ex As Exception
                    screenshot.Dispose()
                    Continue While
                End Try

                board = recognizer.GetBoardState(mat)
                mat.Dispose()

                If board IsNot Nothing Then
                    capturedBmp = screenshot  ' 시각화용 보존
                Else
                    screenshot.Dispose()
                End If
            End While

            If _cancelRequested Then
                Return False
            End If

            ' 2. 진영 자동 감지
            ' 보드가 뒤집혔으면 원래 화면에서 한(HAN)이 아래에 있었으므로 내 진영 = 한
            Dim aiSide = item.AISide
            If aiSide = "AUTO" Then
                If recognizer.IsBoardFlipped Then
                    aiSide = Constants.HAN
                Else
                    For r = 7 To 9
                        For c = 3 To 5
                            Dim piece = board.Grid(r)(c)
                            If piece = Constants.CK Then aiSide = Constants.CHO : Exit For
                            If piece = Constants.HK Then aiSide = Constants.HAN : Exit For
                        Next
                        If aiSide <> "AUTO" Then Exit For
                    Next
                End If
                If aiSide = "AUTO" Then aiSide = Constants.CHO
            End If

            ' 3. AI 최적수 계산
            Dim sideText = If(aiSide = Constants.CHO, "초", "한")
            RaiseEvent ProgressChanged(index + 1, totalCount, item.Name, $"AI: {sideText} 탐색 중 (깊이:{item.AIDepth}, 시간:{item.AITime:F0}s)...", currentRepeat, totalRepeat)
            Application.DoEvents()

            Dim result = Search.FindBestMove(board, aiSide, item.AIDepth, item.AITime)

            If Not result.BestMove.HasValue Then
                capturedBmp?.Dispose()
                RaiseEvent ProgressChanged(index + 1, totalCount, item.Name, $"게임 종료 (착수 불가)", currentRepeat, totalRepeat)
                Return True  ' 게임 종료
            End If

            Dim bestMove = result.BestMove.Value
            Dim fromRow = bestMove.Item1.Item1
            Dim fromCol = bestMove.Item1.Item2
            Dim toRow = bestMove.Item2.Item1
            Dim toCol = bestMove.Item2.Item2

            ' 5. 그리드 좌표 → 화면 좌표 변환 (반전 보정)
            Dim gridPos = recognizer.GetGridPositions()
            If gridPos Is Nothing Then
                RaiseEvent MacroCompleted(False, "AI 그리드 좌표 없음")
                Return False
            End If

            Dim actualFromRow = recognizer.TranslateRow(fromRow)
            Dim actualToRow = recognizer.TranslateRow(toRow)
            Dim fromIdx = actualFromRow * BOARD_COLS + fromCol
            Dim toIdx = actualToRow * BOARD_COLS + toCol
            Dim fromScreenX = gridPos(fromIdx)(0)
            Dim fromScreenY = gridPos(fromIdx)(1)
            Dim toScreenX = gridPos(toIdx)(0)
            Dim toScreenY = gridPos(toIdx)(1)

            If _cancelRequested Then
                Return False
            End If

            ' 6. 시각화 이벤트 발생
            Dim movePiece = board.Grid(fromRow)(fromCol)
            Dim pieceName = ""
            BoardRecognizer.PIECE_NAMES.TryGetValue(movePiece, pieceName)
            Dim capturedPiece = board.Grid(toRow)(toCol)
            Dim captureInfo = ""
            If capturedPiece <> EMPTY Then
                Dim capName = ""
                BoardRecognizer.PIECE_NAMES.TryGetValue(capturedPiece, capName)
                captureInfo = $" (잡기:{capName})"
            End If
            Dim moveInfo = $"{sideText} {pieceName} ({fromRow},{fromCol})→({toRow},{toCol}){captureInfo} | 점수:{result.Score} 깊이:{result.Depth}"

            RaiseEvent AIMoveVisualize(capturedBmp, fromScreenX, fromScreenY, toScreenX, toScreenY, moveInfo)
            Application.DoEvents()

            ' 7. 출발지 클릭 → 대기 → 도착지 클릭
            RaiseEvent ProgressChanged(index + 1, totalCount, item.Name, $"AI: {moveInfo} 클릭...", currentRepeat, totalRepeat)
            Application.DoEvents()

            ' 출발지 클릭
            ButtonClicker.ClickInWindow(targetWindow.Handle, fromScreenX, fromScreenY, useBackground, ClickButton.Left)
            Thread.Sleep(300)

            If _cancelRequested Then
                capturedBmp?.Dispose()
                Return False
            End If

            ' 도착지 클릭
            ButtonClicker.ClickInWindow(targetWindow.Handle, toScreenX, toScreenY, useBackground, ClickButton.Left)

            RaiseEvent ProgressChanged(index + 1, totalCount, item.Name, $"AI: 수 완료 ({moveInfo}) → 상대 차례 대기...", currentRepeat, totalRepeat)
            Application.DoEvents()

            ' 상대 차례 대기 (1초)
            Dim dw = 0
            While dw < 1000 AndAlso Not _cancelRequested
                Thread.Sleep(100)
                dw += 100
                Application.DoEvents()
            End While

            Return False
        End Function

        ''' <summary>
        ''' 매크로를 .macro 파일 + 동명 폴더(템플릿)로 저장
        ''' </summary>
        Public Shared Sub SaveToFile(macroFilePath As String, items As List(Of MacroItem), Optional windowTitle As String = Nothing)
            Dim dir = IO.Path.GetDirectoryName(macroFilePath)
            Dim baseName = IO.Path.GetFileNameWithoutExtension(macroFilePath)
            Dim imgFolder = IO.Path.Combine(dir, baseName)

            If Not IO.Directory.Exists(imgFolder) Then
                IO.Directory.CreateDirectory(imgFolder)
            End If

            Dim lines As New List(Of String)
            ' 첫 줄: 창 이름 (WINDOW| 접두사)
            If Not String.IsNullOrEmpty(windowTitle) Then
                lines.Add($"WINDOW|{windowTitle}")
            End If
            For i = 0 To items.Count - 1
                Dim item = items(i)
                If item.IsAI Then
                    ' AI 항목: AI|이름|대기|임계값|버튼|키|AI진영|AI깊이|AI시간
                    lines.Add($"AI|{item.Name}|{item.DelayAfterClick}|{item.Threshold:F2}|0||{item.AISide}|{item.AIDepth}|{item.AITime:F1}")
                Else
                    Dim imgFileName = $"{(i + 1):D2}_{SanitizeFileName(item.Name)}.png"
                    lines.Add($"{imgFileName}|{item.Name}|{item.DelayAfterClick}|{item.Threshold:F2}|{CInt(item.Button)}|{item.SendKeys}|{item.ClickOffsetX}|{item.ClickOffsetY}")
                    ' 이미지를 폴더에 저장
                    item.Image.Save(IO.Path.Combine(imgFolder, imgFileName), Imaging.ImageFormat.Png)
                End If
            Next

            IO.File.WriteAllLines(macroFilePath, lines.ToArray(), System.Text.Encoding.UTF8)
        End Sub

        ''' <summary>
        ''' .macro 파일에서 매크로 리스트 로드
        ''' </summary>
        Public Shared Function LoadFromFile(macroFilePath As String, Optional ByRef windowTitle As String = Nothing) As List(Of MacroItem)
            Dim items As New List(Of MacroItem)
            windowTitle = Nothing

            If Not IO.File.Exists(macroFilePath) Then
                Return items
            End If

            Dim dir = IO.Path.GetDirectoryName(macroFilePath)
            Dim baseName = IO.Path.GetFileNameWithoutExtension(macroFilePath)
            Dim imgFolder = IO.Path.Combine(dir, baseName)

            Dim lines = IO.File.ReadAllLines(macroFilePath, System.Text.Encoding.UTF8)
            For Each line In lines
                If String.IsNullOrWhiteSpace(line) Then Continue For

                Dim parts = line.Split("|"c)

                ' 창 이름 행
                If parts(0).Trim() = "WINDOW" Then
                    If parts.Length >= 2 Then windowTitle = parts(1).Trim()
                    Continue For
                End If
                If parts.Length < 4 Then Continue For

                ' AI 항목 감지: 첫 필드가 "AI"
                If parts(0).Trim() = "AI" Then
                    Dim aiItem As New MacroItem()
                    aiItem.IsAI = True
                    aiItem.Name = parts(1).Trim()
                    Integer.TryParse(parts(2).Trim(), aiItem.DelayAfterClick)
                    aiItem.Threshold = -1
                    aiItem.Image = New Bitmap(1, 1)
                    If parts.Length >= 7 Then
                        aiItem.AISide = parts(6).Trim()
                    End If
                    If parts.Length >= 8 Then
                        Integer.TryParse(parts(7).Trim(), aiItem.AIDepth)
                    End If
                    If parts.Length >= 9 Then
                        Double.TryParse(parts(8).Trim(), aiItem.AITime)
                    End If
                    items.Add(aiItem)
                    Continue For
                End If

                ' 일반 항목
                Dim imgPath = IO.Path.Combine(imgFolder, parts(0).Trim())
                If Not IO.File.Exists(imgPath) Then Continue For

                Dim item As New MacroItem()
                item.Name = parts(1).Trim()
                item.Image = New Bitmap(imgPath)
                Integer.TryParse(parts(2).Trim(), item.DelayAfterClick)
                Double.TryParse(parts(3).Trim(), item.Threshold)
                If parts.Length >= 5 Then
                    Dim btnVal As Integer
                    If Integer.TryParse(parts(4).Trim(), btnVal) Then
                        item.Button = CType(btnVal, ClickButton)
                    End If
                End If
                If parts.Length >= 6 Then
                    item.SendKeys = parts(5).Trim()
                End If
                If parts.Length >= 8 Then
                    Integer.TryParse(parts(6).Trim(), item.ClickOffsetX)
                    Integer.TryParse(parts(7).Trim(), item.ClickOffsetY)
                End If

                items.Add(item)
            Next

            Return items
        End Function

        ''' <summary>
        ''' 창에 키보드 입력 전송
        ''' </summary>
        Private Shared Sub SendKeysToWindow(hWnd As IntPtr, keys As String, useBackground As Boolean)
            If useBackground Then
                Dim i = 0
                While i < keys.Length
                    If keys(i) = "{"c Then
                        Dim endIdx = keys.IndexOf("}"c, i)
                        If endIdx > i Then
                            Dim keyName = keys.Substring(i + 1, endIdx - i - 1).ToUpper()
                            Dim vk = GetVirtualKey(keyName)
                            If vk > 0 Then
                                NativeMethods.PostMessage(hWnd, NativeMethods.WM_KEYDOWN, New IntPtr(vk), IntPtr.Zero)
                                Thread.Sleep(30)
                                NativeMethods.PostMessage(hWnd, NativeMethods.WM_KEYUP, New IntPtr(vk), IntPtr.Zero)
                                Thread.Sleep(30)
                            End If
                            i = endIdx + 1
                            Continue While
                        End If
                    End If

                    Dim charCode = AscW(keys(i))
                    NativeMethods.PostMessage(hWnd, NativeMethods.WM_CHAR, New IntPtr(charCode), IntPtr.Zero)
                    Thread.Sleep(30)
                    i += 1
                End While
            Else
                WindowFinder.BringToFront(hWnd)
                Thread.Sleep(200)
                System.Windows.Forms.SendKeys.SendWait(keys)
            End If
        End Sub

        Private Shared Function GetVirtualKey(keyName As String) As Integer
            Select Case keyName
                Case "ENTER", "RETURN" : Return &HD
                Case "TAB" : Return &H9
                Case "ESC", "ESCAPE" : Return &H1B
                Case "SPACE" : Return &H20
                Case "BS", "BACKSPACE" : Return &H8
                Case "DEL", "DELETE" : Return &H2E
                Case "UP" : Return &H26
                Case "DOWN" : Return &H28
                Case "LEFT" : Return &H25
                Case "RIGHT" : Return &H27
                Case "HOME" : Return &H24
                Case "END" : Return &H23
                Case "PGUP" : Return &H21
                Case "PGDN" : Return &H22
                Case "F1" : Return &H70
                Case "F2" : Return &H71
                Case "F3" : Return &H72
                Case "F4" : Return &H73
                Case "F5" : Return &H74
                Case "F6" : Return &H75
                Case "F7" : Return &H76
                Case "F8" : Return &H77
                Case "F9" : Return &H78
                Case "F10" : Return &H79
                Case "F11" : Return &H7A
                Case "F12" : Return &H7B
                Case Else : Return 0
            End Select
        End Function

        Private Shared Function SanitizeFileName(name As String) As String
            Dim invalid = IO.Path.GetInvalidFileNameChars()
            Dim result = name
            For Each c In invalid
                result = result.Replace(c, "_"c)
            Next
            Return result
        End Function
    End Class
End Namespace
