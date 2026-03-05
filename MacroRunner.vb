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
Imports System.Linq

Namespace MacroAutoControl
    ''' <summary>
    ''' 매크로 시퀀스 실행 엔진
    ''' </summary>
    Public Class MacroRunner
        Public Event ProgressChanged(currentIndex As Integer, totalCount As Integer, itemName As String, status As String, currentRepeat As Integer, totalRepeat As Integer)
        Public Event MacroCompleted(success As Boolean, message As String)
        Public Event AIMoveVisualize(screenshot As Bitmap, fromX As Integer, fromY As Integer, toX As Integer, toY As Integer, moveInfo As String)
        Public Event BoardPreviewUpdate(preview As Bitmap)

        Private _cancelRequested As Boolean
        Private _isRunning As Boolean

        ' 현재 실행 중인 AI 항목 (깊이 조절용)
        Private _currentAIItem As MacroItem = Nothing

        ' AI 보드 인식기 (재사용)
        Private _recognizer As BoardRecognizer = Nothing

        ' 게임 종료 사유 (팝업 감지 시 설정)
        Private _gameEndReason As String = Nothing

        ' 보드/그리드 인식 연속 실패 카운터
        Private _recognitionFailCount As Integer = 0

        ' 동일구간 반복 금지를 위한 게임 수준 해시 히스토리
        Private _gameHashHistory As New List(Of ULong)

        ' 동일구간 반복 금지를 위한 AI 수 히스토리 (from, to)
        Private _gameAIMoves As New List(Of ((Integer, Integer), (Integer, Integer)))

        ' 내 차례 스크린샷 저장
        Private _watchDir As String = Nothing
        Private _watchDateDir As String = Nothing
        Private _watchCounter As Integer = 0
        Private _watchPreviousBoard As Board = Nothing

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

        ' 내차례 강제 지정 플래그
        Private _forceMyTurn As Boolean = False

        ''' <summary>내 차례 대기를 강제로 건너뛰고 즉시 AI 탐색 진행</summary>
        Public Sub ForceMyTurn()
            _forceMyTurn = True
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
        ''' 정규화된 보드 하단(row 7~9) 궁 기물로 내 진영 판단.
        ''' 정규화 후 항상 초=아래이므로, 보드 반전 여부와 화면 하단 궁 위치로 판단.
        ''' </summary>
        Private Function DetectMySideByBoard(board As Board) As String
            If board Is Nothing Then Return Constants.CHO
            ' 정규화된 보드: row 7~9 = 초 궁성, row 0~2 = 한 궁성
            ' IsBoardFlipped = True → 화면에서 초궁이 위에 있었음 → 사용자는 한
            ' IsBoardFlipped = False → 화면에서 초궁이 아래 → 사용자는 초
            Dim recognizer = GetRecognizer()
            If recognizer.IsBoardFlipped Then
                Return Constants.HAN
            Else
                Return Constants.CHO
            End If
        End Function

        ''' <summary>
        ''' BoardRecognizer.DetectHighlightedPiece를 사용하여 내 차례인지 판정.
        ''' 하이라이트된 기물이 상대 기물이면 → 내 차례 (True)
        ''' 하이라이트된 기물이 내 기물이면 → 상대 차례 (False)
        ''' 하이라이트 없으면 → 내 차례 (첫 수 등)
        ''' </summary>
        Private Function HasGlowAroundMyPieces(mat As Mat, board As Board, gridPos As Integer()(), mySide As String, cellW As Integer, cellH As Integer, boardFlipped As Boolean, targetWindow As WindowFinder.WindowInfo, Optional ByRef glowPieceName As String = Nothing) As Boolean
            Dim recognizer = GetRecognizer()
            Dim highlighted = recognizer.DetectHighlightedPiece(mat, board)

            If highlighted Is Nothing Then
                ' 빛남 없음: 초(선공)이면 내 차례, 한(후공)이면 대기
                If mySide = Constants.CHO Then
                    glowPieceName = "빛남없음(초선공)"
                    Return True
                Else
                    glowPieceName = "빛남없음(한후공대기)"
                    Return False
                End If
            End If

            Dim hlSide = highlighted.Value.Side
            Dim hlGlow = highlighted.Value.GlowCount
            Dim piece = board.Grid(highlighted.Value.Row)(highlighted.Value.Col)
            Dim pieceName = If(BoardRecognizer.PIECE_NAMES.ContainsKey(piece), BoardRecognizer.PIECE_NAMES(piece), piece)

            If hlSide = mySide Then
                ' 내 기물이 하이라이트 → 상대 차례
                glowPieceName = $"내빛남:{pieceName} glow={hlGlow}"
                Return False
            Else
                ' 상대 기물이 하이라이트 → 내 차례
                glowPieceName = $"상대빛남:{pieceName}({highlighted.Value.Row},{highlighted.Value.Col}) glow={hlGlow}"
                Return True
            End If
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
        ''' 매크로 목록의 패턴매칭 항목을 스크린샷에 대해 검사.
        ''' 매칭되면 클릭 후 항목 이름 반환, 없으면 Nothing.
        ''' </summary>
        Private Function CheckMacroPatterns(screenshot As Bitmap, items As List(Of MacroItem),
                                             targetWindow As WindowFinder.WindowInfo,
                                             useBackground As Boolean) As String
            For Each mi In items
                If Not mi.IsAI Then Continue For
                If mi.Threshold <= 0 Then Continue For
                If mi.Image Is Nothing OrElse mi.Image.Width <= 1 Then Continue For

                Dim result = ButtonFinder.FindByTemplate(screenshot, mi.Image, mi.Threshold, mi.Mask)
                If result.Found Then
                    Dim clickX As Integer
                    Dim clickY As Integer
                    If mi.ClickOffsetX >= 0 Then
                        clickX = result.MatchRect.X + mi.ClickOffsetX
                        clickY = result.MatchRect.Y + mi.ClickOffsetY
                    Else
                        clickX = result.Location.X
                        clickY = result.Location.Y
                    End If

                    ButtonClicker.ClickInWindow(targetWindow.Handle, clickX, clickY, useBackground, mi.Button)
                    Thread.Sleep(500)

                    ' 키전송이 있으면 실행
                    If Not String.IsNullOrEmpty(mi.SendKeys) Then
                        SendKeysToWindow(targetWindow.Handle, mi.SendKeys, useBackground)
                    End If

                    Return mi.Name
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
                                       currentRepeat As Integer, totalRepeat As Integer,
                                       Optional items As List(Of MacroItem) = Nothing,
                                       Optional useBackground As Boolean = False) As Board
            Dim recognizer = GetRecognizer()
            Dim pollCount = 0
            Dim mySide As String = Nothing  ' 첫 감지 후 고정

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

                ' 캡처
                Dim screenshot = WindowFinder.CaptureWindow(targetWindow.Handle)
                If screenshot Is Nothing Then Continue While

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
                    _recognitionFailCount += 1

                    ' 보드 미감지 시 매크로 목록의 패턴매칭 확인 → 클릭 후 계속 대기
                    If items IsNot Nothing AndAlso _recognitionFailCount >= 2 Then
                        Dim matched = CheckMacroPatterns(screenshot, items, targetWindow, useBackground)
                        If matched IsNot Nothing Then
                            ' "종료" 포함 패턴 → AI엔진 종료
                            If matched.Contains("종료") Then
                                mat.Dispose()
                                screenshot.Dispose()
                                _gameEndReason = $"패턴매칭 종료: {matched}"
                                Return Nothing
                            End If
                            _recognitionFailCount = 0
                            RaiseEvent ProgressChanged(index + 1, totalCount, item.Name,
                                $"AI: 패턴매칭 클릭: {matched} → 계속 대기...", currentRepeat, totalRepeat)
                            Application.DoEvents()
                        End If
                    End If

                    If _recognitionFailCount > 10 Then
                        mat.Dispose()
                        screenshot.Dispose()
                        _gameEndReason = "보드/그리드 인식 실패 10회 초과"
                        Return Nothing
                    End If
                    mat.Dispose()
                    screenshot.Dispose()
                    Continue While
                End If

                ' 진영 자동 감지 (첫 감지 후 고정) - 정규화된 보드 하단 궁 기물로 판단
                If mySide Is Nothing Then
                    mySide = item.AISide
                    If mySide = "AUTO" Then
                        mySide = DetectMySideByBoard(board)
                    End If
                End If
                Dim dbgSide = If(mySide = Constants.CHO, "초", "한")
                RaiseEvent ProgressChanged(index + 1, totalCount, item.Name,
                    $"AI: 진영={dbgSide} 대기 중... ({pollCount})", currentRepeat, totalRepeat)
                Application.DoEvents()

                ' 보드 프리뷰 표시
                Dim gridPos2 = recognizer.GetGridPositions()
                If gridPos2 IsNot Nothing Then
                    Try
                        Dim hlPiece = recognizer.DetectHighlightedPiece(mat, board)
                        Dim allGlows = recognizer.GetAllGlowValues(mat, board)
                        Dim glowMap As New Dictionary(Of (Integer, Integer), Integer)
                        For Each gv In allGlows
                            glowMap((gv.Row, gv.Col)) = gv.Glow
                        Next

                        Dim myPieces = If(mySide = Constants.CHO, Constants.CHO_PIECES, Constants.HAN_PIECES)
                        Dim sideText = If(mySide = Constants.CHO, "초", "한")
                        Dim myCount = 0
                        Dim enemyCount = 0
                        Dim moveInfo = ""

                        Dim preview = New Bitmap(screenshot)
                        Using g = Graphics.FromImage(preview)
                            g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias

                            ' 마지막 이동 기물 표시 (검정 원)
                            If hlPiece.HasValue Then
                                Dim hr = hlPiece.Value.Row
                                Dim hc = hlPiece.Value.Col
                                Dim sr = recognizer.TranslateRow(hr)
                                Dim si = sr * Constants.BOARD_COLS + hc
                                If si >= 0 AndAlso si < gridPos2.Length Then
                                    Dim hx = gridPos2(si)(0), hy = gridPos2(si)(1)
                                    Using pen As New Pen(Color.FromArgb(220, 0, 0, 0), 4)
                                        g.DrawEllipse(pen, hx - 26, hy - 26, 52, 52)
                                    End Using
                                End If
                            End If

                            ' glow 검출 링 범위 계산
                            Dim cellSz = recognizer.GetCellSize()
                            Dim baseSize = Math.Max(cellSz.Item1, cellSz.Item2)
                            If baseSize <= 0 Then baseSize = 60
                            Dim innerR = CInt(baseSize * 0.42)
                            Dim outerR = CInt(baseSize * 0.52)
                            If innerR < 18 Then innerR = 18
                            If outerR < 22 Then outerR = 22
                            Dim smallInnerR = CInt(baseSize * 0.32)
                            Dim smallOuterR = CInt(baseSize * 0.42)
                            If smallInnerR < 14 Then smallInnerR = 14
                            If smallOuterR < 18 Then smallOuterR = 18
                            Dim bigInnerR = CInt(baseSize * 0.42)
                            Dim bigOuterR = CInt(baseSize * 0.572)
                            If bigInnerR < 18 Then bigInnerR = 18
                            If bigOuterR < 24 Then bigOuterR = 24

                            ' 기물 표시 (적 기물 먼저, 내 기물 나중에)
                            For pass = 0 To 1
                                For r = 0 To Constants.BOARD_ROWS - 1
                                    For c = 0 To Constants.BOARD_COLS - 1
                                        Dim piece = board.Grid(r)(c)
                                        If piece = Constants.EMPTY Then Continue For
                                        Dim isMine = myPieces.Contains(piece)
                                        If pass = 0 AndAlso isMine Then Continue For
                                        If pass = 1 AndAlso Not isMine Then Continue For

                                        Dim screenR = recognizer.TranslateRow(r)
                                        Dim idx2 = screenR * Constants.BOARD_COLS + c
                                        If idx2 < 0 OrElse idx2 >= gridPos2.Length Then Continue For
                                        Dim px = gridPos2(idx2)(0)
                                        Dim py = gridPos2(idx2)(1)

                                        ' glow 검출 범위 (내원/외원: 갈색 점선)
                                        Dim drawInner As Integer = innerR
                                        Dim drawOuter As Integer = outerR
                                        If piece = Constants.CJ OrElse piece = Constants.HB OrElse piece = Constants.CS OrElse piece = Constants.HS Then
                                            drawInner = smallInnerR
                                            drawOuter = smallOuterR
                                        ElseIf piece = Constants.CK OrElse piece = Constants.HK Then
                                            drawInner = bigInnerR
                                            drawOuter = bigOuterR
                                        End If
                                        Using pen As New Pen(Color.FromArgb(150, 139, 90, 43), 1)
                                            pen.DashStyle = Drawing2D.DashStyle.Dot
                                            g.DrawEllipse(pen, px - drawInner, py - drawInner, drawInner * 2, drawInner * 2)
                                            g.DrawEllipse(pen, px - drawOuter, py - drawOuter, drawOuter * 2, drawOuter * 2)
                                        End Using

                                        ' 카운트
                                        Dim pName = ""
                                        BoardRecognizer.PIECE_NAMES.TryGetValue(piece, pName)
                                        If pass = 1 AndAlso isMine Then
                                            myCount += 1
                                        ElseIf pass = 0 AndAlso Not isMine Then
                                            enemyCount += 1
                                        End If

                                        ' 초=파란원, 한=빨간원 (졸/병/사는 작게)
                                        Dim circleR As Integer = 22
                                        If piece = Constants.CJ OrElse piece = Constants.HB OrElse piece = Constants.CS OrElse piece = Constants.HS Then
                                            circleR = 16
                                        End If
                                        If Constants.CHO_PIECES.Contains(piece) Then
                                            Using pen As New Pen(Color.FromArgb(200, 0, 120, 255), 3)
                                                g.DrawEllipse(pen, px - circleR, py - circleR, circleR * 2, circleR * 2)
                                            End Using
                                        ElseIf Constants.HAN_PIECES.Contains(piece) Then
                                            Using pen As New Pen(Color.FromArgb(200, 255, 60, 60), 3)
                                                g.DrawEllipse(pen, px - circleR, py - circleR, circleR * 2, circleR * 2)
                                            End Using
                                        End If

                                        ' 내 궁에 남색 사각형 표시
                                        If isMine AndAlso (piece = Constants.CK OrElse piece = Constants.HK) Then
                                            Using pen As New Pen(Color.FromArgb(255, 0, 0, 128), 3)
                                                g.DrawRectangle(pen, px - circleR, py - circleR, circleR * 2, circleR * 2)
                                            End Using
                                        End If

                                        ' 기물 이름 + glow 값
                                        Dim shortName = If(pName.Length >= 2, pName.Substring(1), pName)
                                        Dim glowVal = 0
                                        glowMap.TryGetValue((r, c), glowVal)
                                        Dim label = If(glowVal > 0, $"{shortName} {glowVal}", shortName)
                                        Using font As New Font("맑은 고딕", 8, FontStyle.Bold)
                                            Dim sz = g.MeasureString(label, font)
                                            Dim txtX = px - sz.Width / 2
                                            Dim txtY = py - 24 - sz.Height
                                            Using bgBrush As New SolidBrush(Color.FromArgb(180, 0, 0, 0))
                                                g.FillRectangle(bgBrush, txtX - 1, txtY - 1, sz.Width + 2, sz.Height + 1)
                                            End Using
                                            Dim textColor = If(glowVal > 0, Brushes.Yellow, Brushes.White)
                                            g.DrawString(label, font, textColor, txtX, txtY)
                                        End Using
                                    Next
                                Next
                            Next

                            ' 장군 감지
                            Dim checkInfo = ""
                            Dim inCheck = False
                            If hlPiece.HasValue Then
                                Dim movedPiece = board.Grid(hlPiece.Value.Row)(hlPiece.Value.Col)
                                Dim movedName = ""
                                BoardRecognizer.PIECE_NAMES.TryGetValue(movedPiece, movedName)
                                moveInfo = $"  |  마지막: {movedName}"
                                Dim movedSide = If(Constants.CHO_PIECES.Contains(movedPiece), Constants.CHO, Constants.HAN)
                                Dim targetSide = If(movedSide = Constants.CHO, Constants.HAN, Constants.CHO)
                                If board.IsInCheck(targetSide) Then
                                    If targetSide = mySide Then
                                        inCheck = True
                                        checkInfo = "  ★ 장군! 멍군필요"
                                    Else
                                        checkInfo = "  ★ 장군!"
                                    End If
                                End If
                            Else
                                moveInfo = "  |  마지막 이동 기물 없음"
                            End If
                            Dim summary = $"[{sideText}] 내 {myCount}개  적 {enemyCount}개{moveInfo}{checkInfo}"
                            Dim summaryColor = If(inCheck, Color.FromArgb(255, 255, 80, 80), Color.FromArgb(255, 100, 255, 100))
                            Using font As New Font("맑은 고딕", 10, FontStyle.Bold)
                                Dim sz = g.MeasureString(summary, font)
                                Using bgBrush As New SolidBrush(Color.FromArgb(200, 0, 0, 0))
                                    g.FillRectangle(bgBrush, 4, 4, sz.Width + 8, sz.Height + 4)
                                End Using
                                Using textBrush As New SolidBrush(summaryColor)
                                    g.DrawString(summary, font, textBrush, 8, 6)
                                End Using
                            End Using
                        End Using
                        RaiseEvent BoardPreviewUpdate(preview)
                        Application.DoEvents()
                    Catch
                    End Try
                End If

                ' 빛남 감지
                Dim gridPos = recognizer.GetGridPositions()
                Dim cellSize = recognizer.GetCellSize()
                Dim glowPiece As String = Nothing
                Dim isMyTurn = False
                If gridPos IsNot Nothing Then
                    isMyTurn = HasGlowAroundMyPieces(mat, board, gridPos, mySide, cellSize.Item1, cellSize.Item2, recognizer.IsBoardFlipped, targetWindow, glowPiece)
                End If
                ' 강제 내차례: 상대 기물 확인된 경우만 허용, 빛남 없으면 강제 진행
                If _forceMyTurn Then
                    If Not isMyTurn AndAlso glowPiece IsNot Nothing AndAlso glowPiece.StartsWith("내빛남") Then
                        ' 내 기물이 빛나면 강제 지정 무시
                        glowPiece = $"강제지정 거부 ({glowPiece})"
                    Else
                        isMyTurn = True
                        glowPiece = If(glowPiece, "강제지정")
                    End If
                    _forceMyTurn = False
                End If

                ' 기물 위치 변동 시 스크린샷 저장 (내 차례/상대 차례 무관)
                SaveWatchIfChanged(screenshot, board)

                If Not isMyTurn Then
                    mat.Dispose()
                    screenshot.Dispose()
                    Continue While
                End If

            ' 상대 수 애니메이션 완료 대기 후 재캡처
            mat.Dispose()
            screenshot.Dispose()

            Dim sideDbg2 = If(mySide = Constants.CHO, "초", "한")
            RaiseEvent ProgressChanged(index + 1, totalCount, item.Name,
                $"AI: 내 차례 감지 ({sideDbg2}, 빛남:{glowPiece}) → 대기...", currentRepeat, totalRepeat)
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

            ' 최종 확인: 하이라이트된 기물이 여전히 상대 기물인지 검증
            Dim verifyBoard = If(finalBoard, board)
            Dim stillMyTurn2 = True
            Dim verifySs = WindowFinder.CaptureWindow(targetWindow.Handle)
            If verifySs IsNot Nothing Then
                Dim verifyMat As Mat = Nothing
                Try
                    verifyMat = BitmapConverter.ToMat(verifySs)
                    If verifyMat.Channels() = 4 Then
                        Dim bgr3 As New Mat()
                        Cv2.CvtColor(verifyMat, bgr3, ColorConversionCodes.BGRA2BGR)
                        verifyMat.Dispose()
                        verifyMat = bgr3
                    End If
                    Dim vGridPos = recognizer.GetGridPositions()
                    Dim vCellSize = recognizer.GetCellSize()
                    If vGridPos IsNot Nothing Then
                        Dim vGlowPiece As String = Nothing
                        stillMyTurn2 = HasGlowAroundMyPieces(verifyMat, verifyBoard, vGridPos, mySide, vCellSize.Item1, vCellSize.Item2, recognizer.IsBoardFlipped, targetWindow, vGlowPiece)
                        If Not stillMyTurn2 Then
                            RaiseEvent ProgressChanged(index + 1, totalCount, item.Name, $"AI: 내 차례 재확인 실패 ({vGlowPiece}) → 재대기...", currentRepeat, totalRepeat)
                            Application.DoEvents()
                        End If
                    End If
                Catch
                Finally
                    verifyMat?.Dispose()
                    verifySs.Dispose()
                End Try
            End If

            If Not stillMyTurn2 Then Continue While

            If finalBoard IsNot Nothing Then
                Return finalBoard
            Else
                Return board
            End If
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
                               Optional repeatCount As Integer = 1,
                               Optional initialStartIndex As Integer = 0)
            If _isRunning Then
                Return
            End If

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
            _watchDir = Nothing
            _watchPreviousBoard = Nothing
            Dim baseDir = "D:\images\macro_janggi"
            Dim dateStr = DateTime.Now.ToString("yyyy-MM-dd")
            _watchDateDir = IO.Path.Combine(baseDir, $"watch_{dateStr}")
            IO.Directory.CreateDirectory(_watchDateDir)

            Dim defaultWindow = targetWindow

            ' 첫줄 항목에 창이 지정되어 있으면 그것을 기본 창으로 사용
            If items.Count > 0 AndAlso Not String.IsNullOrEmpty(items(0).WindowTitle) Then
                Dim windows = WindowFinder.GetAllVisibleWindows()
                Dim matched = windows.FirstOrDefault(Function(w) w.Title = items(0).WindowTitle)
                If matched Is Nothing Then
                    matched = windows.FirstOrDefault(Function(w) w.Title.Contains(items(0).WindowTitle))
                End If
                If matched IsNot Nothing Then
                    defaultWindow = matched
                End If
            End If

            Dim infinite = (repeatCount = 0)
            Dim totalRepeat = If(infinite, 0, repeatCount)
            Dim currentRepeat = 0

            ' 첫 AI 항목 인덱스 찾기 (반복 재시작 시 AI부터 시작)
            Dim firstAIIndex = 0
            For idx = 0 To items.Count - 1
                If items(idx).IsAI Then firstAIIndex = idx : Exit For
            Next
            Dim startIndex = Math.Max(0, Math.Min(initialStartIndex, items.Count - 1))

            Try
                While infinite OrElse currentRepeat < totalRepeat
                    currentRepeat += 1

                    For i = startIndex To items.Count - 1
                        If _cancelRequested Then
                            RaiseEvent MacroCompleted(False, $"매크로 중지됨 (반복 {currentRepeat}, {i}/{items.Count})")
                            Return
                        End If

                        Dim item = items(i)

                        ' 항목별 대상 창 해결 (없으면 첫줄 윈도우 사용)
                        If Not String.IsNullOrEmpty(item.WindowTitle) Then
                            Dim windows = WindowFinder.GetAllVisibleWindows()
                            Dim matched = windows.FirstOrDefault(Function(w) w.Title = item.WindowTitle)
                            If matched Is Nothing Then
                                matched = windows.FirstOrDefault(Function(w) w.Title.Contains(item.WindowTitle))
                            End If
                            If matched IsNot Nothing Then
                                targetWindow = matched
                            End If
                        Else
                            targetWindow = defaultWindow
                        End If

                        If item.IsAI AndAlso item.Image IsNot Nothing AndAlso item.Image.Width > 1 Then
                            ' AI+패턴 항목: 실행 시 스킵 (CheckMacroPatterns에서 사용)
                            Continue For
                        ElseIf item.IsAI Then
                            ' AI 항목 실행
                            Dim gameEnded = ExecuteAIItem(item, i, items.Count, targetWindow, ownerForm, useBackground, currentRepeat, totalRepeat, items)
                            If _cancelRequested Then
                                RaiseEvent MacroCompleted(False, $"매크로 중지됨 (반복 {currentRepeat}, AI)")
                                Return
                            End If

                            If gameEnded AndAlso infinite Then
                                ' 게임 종료 → 그리드 초기화 후 매크로 처음부터 다시 시작
                    
                                _gameEndReason = Nothing
                                _recognizer?.ResetGrid()
                                _gameHashHistory.Clear()
                    _gameAIMoves.Clear()
                    _watchPreviousBoard = Nothing
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
                            ' 이미지 찾기+클릭 항목
                            If item.DelayAfterClick > 0 Then
                                ' 폴링 모드: DelayAfterClick 시간 내에 0.5초 간격으로 반복 탐색
                                Dim elapsed = 0
                                Dim imageFound = False
                                While elapsed < item.DelayAfterClick AndAlso Not _cancelRequested
                                    Dim remaining = item.DelayAfterClick - elapsed
                                    RaiseEvent ProgressChanged(i + 1, items.Count, item.Name, $"찾는 중... ({remaining}ms 남음)", currentRepeat, totalRepeat)
                                    Application.DoEvents()

                                    Dim screenshot = WindowFinder.CaptureWindow(targetWindow.Handle)
                                    If screenshot IsNot Nothing Then
                                        Dim result = ButtonFinder.FindByTemplate(screenshot, item.Image, item.Threshold, item.Mask)
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

                                            ' 클릭 후 1초 대기
                                            RaiseEvent ProgressChanged(i + 1, items.Count, item.Name, "클릭 완료 → 1초 대기...", currentRepeat, totalRepeat)
                                            Application.DoEvents()
                                            Dim cw1 = 0
                                            While cw1 < 1000 AndAlso Not _cancelRequested
                                                Thread.Sleep(100)
                                                cw1 += 100
                                                Application.DoEvents()
                                            End While

                                            imageFound = True
                                            Exit While
                                        End If
                                    Else
                                        screenshot?.Dispose()
                                    End If

                                    ' 0.5초 대기 후 재시도
                                    Dim pollWait = Math.Min(500, item.DelayAfterClick - elapsed)
                                    Dim pw = 0
                                    While pw < pollWait AndAlso Not _cancelRequested
                                        Dim chunk = Math.Min(100, pollWait - pw)
                                        Thread.Sleep(chunk)
                                        pw += chunk
                                        Application.DoEvents()
                                    End While
                                    elapsed += pollWait
                                End While

                                If Not imageFound AndAlso Not _cancelRequested Then
                                    RaiseEvent ProgressChanged(i + 1, items.Count, item.Name, "시간 초과 → 다음", currentRepeat, totalRepeat)
                                    Application.DoEvents()
                                End If
                            Else
                                ' 원샷 모드: DelayAfterClick = 0이면 기존 1회 탐색
                                RaiseEvent ProgressChanged(i + 1, items.Count, item.Name, "찾는 중...", currentRepeat, totalRepeat)
                                Application.DoEvents()

                                Dim screenshot = WindowFinder.CaptureWindow(targetWindow.Handle)
                                If screenshot IsNot Nothing AndAlso Not _cancelRequested Then
                                    Dim result = ButtonFinder.FindByTemplate(screenshot, item.Image, item.Threshold, item.Mask)
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

                                        ' 클릭 후 1초 대기
                                        RaiseEvent ProgressChanged(i + 1, items.Count, item.Name, "클릭 완료 → 1초 대기...", currentRepeat, totalRepeat)
                                        Application.DoEvents()
                                        Dim cw2 = 0
                                        While cw2 < 1000 AndAlso Not _cancelRequested
                                            Thread.Sleep(100)
                                            cw2 += 100
                                            Application.DoEvents()
                                        End While
                                    Else
                                        RaiseEvent ProgressChanged(i + 1, items.Count, item.Name, $"못 찾음 ({result.Confidence:P0}) → 다음", currentRepeat, totalRepeat)
                                        Application.DoEvents()
                                    End If
                                Else
                                    screenshot?.Dispose()
                                End If
                            End If
                        End If

                        ' 대기 (AI 항목은 ExecuteAIItem 내부에서 처리, 이미지 항목은 폴링에서 처리)
                        If item.DelayAfterClick > 0 AndAlso Not item.IsAI AndAlso item.Threshold <= 0 Then
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
                        RaiseEvent ProgressChanged(0, items.Count, "", $"반복 대기 (1초)...", currentRepeat, totalRepeat)
                        Application.DoEvents()
                        Dim rw = 0
                        While rw < 1000 AndAlso Not _cancelRequested
                            Thread.Sleep(100)
                            rw += 100
                            Application.DoEvents()
                        End While
                        startIndex = 0  ' 처음부터
                    End If
                End While

                RaiseEvent MacroCompleted(True, $"매크로 완료! ({items.Count}개 항목 x {currentRepeat}회 실행)")

            Finally
                _isRunning = False
                _currentAIItem = Nothing
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
                                   currentRepeat As Integer, totalRepeat As Integer,
                                   Optional items As List(Of MacroItem) = Nothing) As Boolean

            _currentAIItem = item
            Dim recognizer = GetRecognizer()
            Dim aiSide As String = Nothing  ' 첫 감지 후 고정

            While Not _cancelRequested
            Dim board As Board = Nothing

            ' 0. 내 차례 감지: 기물 주변 빛남이 나타날 때까지 대기
            board = WaitForMyTurn(item, targetWindow, ownerForm, index, totalCount,
                                 currentRepeat, totalRepeat, items, useBackground)
            If board Is Nothing Then
                If _gameEndReason IsNot Nothing Then
                    RaiseEvent ProgressChanged(index + 1, totalCount, item.Name, $"팝업 감지: {_gameEndReason} → 클릭 후 다음", currentRepeat, totalRepeat)
                    _gameEndReason = Nothing
                    _recognizer?.ResetGrid()
                    _gameHashHistory.Clear()
                    _gameAIMoves.Clear()
                    _watchPreviousBoard = Nothing
                    Return False  ' 다음 매크로 항목으로
                End If
                Return False
            End If

            ' 시각화용 캡처 (WaitForMyTurn 후에도 현재 화면 캡처)
            Dim capturedBmp As Bitmap = WindowFinder.CaptureWindow(targetWindow.Handle)
            While board Is Nothing AndAlso Not _cancelRequested
                _recognitionFailCount += 1
                If _recognitionFailCount > 10 Then
                    RaiseEvent MacroCompleted(False, $"AI 보드/그리드 인식 실패: 10회 재시도 후 중단")
                    Return False
                End If

                If _recognitionFailCount > 1 Then
                    Dim failDelay = Math.Max(1000, item.DelayAfterClick)
                    RaiseEvent ProgressChanged(index + 1, totalCount, item.Name, $"AI: 재시도 {_recognitionFailCount}/10 ({failDelay / 1000.0:F1}초)...", currentRepeat, totalRepeat)
                    Application.DoEvents()
                    Dim fd = 0
                    While fd < failDelay AndAlso Not _cancelRequested
                        Thread.Sleep(100)
                        fd += 100
                        Application.DoEvents()
                    End While
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
                    _recognitionFailCount = 0
                    capturedBmp = screenshot  ' 시각화용 보존
                Else
                    screenshot.Dispose()
                End If
            End While

            If _cancelRequested Then
                Return False
            End If

            ' 2. 진영 자동 감지 (첫 감지 후 고정)
            If aiSide Is Nothing Then
                aiSide = item.AISide
                If aiSide = "AUTO" Then
                    aiSide = DetectMySideByBoard(board)
                End If
                Dim detSideText = If(aiSide = Constants.CHO, "초", "한")
                Console.WriteLine($"[진영감지] AI 진영: {detSideText} (flipped={recognizer.IsBoardFlipped})")
                RaiseEvent ProgressChanged(index + 1, totalCount, item.Name, $"AI: 진영={detSideText} 감지 완료", currentRepeat, totalRepeat)
                Application.DoEvents()
            End If

            ' 3. AI 최적수 계산
            Dim sideText = If(aiSide = Constants.CHO, "초", "한")
            RaiseEvent ProgressChanged(index + 1, totalCount, item.Name, $"AI: {sideText} 탐색 중 (깊이:{item.AIDepth}, 시간:{item.AITime:F0}s)...", currentRepeat, totalRepeat)
            Application.DoEvents()

            ' 게임 수준 해시 히스토리 주입 (동일구간 반복 금지)
            _gameHashHistory.Add(board.ZobristHash)
            board.InjectGameHistory(_gameHashHistory)

            Dim result = Search.FindBestMove(board, aiSide, item.AIDepth, item.AITime, _gameAIMoves)

            If Not result.BestMove.HasValue Then
                capturedBmp?.Dispose()
                Dim retryDelay = Math.Max(1000, item.DelayAfterClick)
                RaiseEvent ProgressChanged(index + 1, totalCount, item.Name, $"착수 불가 → {retryDelay / 1000.0:F1}초 후 재시도...", currentRepeat, totalRepeat)
                Application.DoEvents()
                Dim rd = 0
                While rd < retryDelay AndAlso Not _cancelRequested
                    Thread.Sleep(100)
                    rd += 100
                    Application.DoEvents()
                End While
                If _cancelRequested Then Return False

                ' 재캡처 + 재인식 + 재탐색
                Dim ss = WindowFinder.CaptureWindow(targetWindow.Handle)
                If ss IsNot Nothing Then
                    Dim m As Mat = Nothing
                    Try
                        m = BitmapConverter.ToMat(ss)
                        If m.Channels() = 4 Then
                            Dim bgr As New Mat()
                            Cv2.CvtColor(m, bgr, ColorConversionCodes.BGRA2BGR)
                            m.Dispose()
                            m = bgr
                        End If
                        Dim retryBoard = recognizer.GetBoardState(m)
                        If retryBoard IsNot Nothing Then
                            board = retryBoard
                            board.InjectGameHistory(_gameHashHistory)
                            result = Search.FindBestMove(board, aiSide, item.AIDepth, item.AITime, _gameAIMoves)
                        End If
                    Catch
                    Finally
                        m?.Dispose()
                        ss.Dispose()
                    End Try
                End If

                If Not result.BestMove.HasValue Then
                    RaiseEvent ProgressChanged(index + 1, totalCount, item.Name, $"착수 불가 → 다음", currentRepeat, totalRepeat)
                    _recognizer?.ResetGrid()
                    _gameHashHistory.Clear()
                    _gameAIMoves.Clear()
                    _watchPreviousBoard = Nothing
                    Return False  ' 다음 매크로 항목으로
                End If
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

            ' 5.5. 클릭 전 상대 기물 하이라이트 재확인
            Dim preClickSs = WindowFinder.CaptureWindow(targetWindow.Handle)
            If preClickSs IsNot Nothing Then
                Dim preClickMat As Mat = Nothing
                Try
                    preClickMat = BitmapConverter.ToMat(preClickSs)
                    If preClickMat.Channels() = 4 Then
                        Dim bgr4 As New Mat()
                        Cv2.CvtColor(preClickMat, bgr4, ColorConversionCodes.BGRA2BGR)
                        preClickMat.Dispose()
                        preClickMat = bgr4
                    End If
                    Dim preBoard = recognizer.GetBoardState(preClickMat)
                    If preBoard IsNot Nothing Then
                        ' 보드 상태 변경 감지 (상대가 AI 계산 중 이동)
                        Dim boardChanged = False
                        For br = 0 To Constants.BOARD_ROWS - 1
                            For bc = 0 To Constants.BOARD_COLS - 1
                                If preBoard.Grid(br)(bc) <> board.Grid(br)(bc) Then
                                    boardChanged = True
                                    Exit For
                                End If
                            Next
                            If boardChanged Then Exit For
                        Next
                        If boardChanged Then
                            preClickMat.Dispose()
                            preClickSs.Dispose()
                            capturedBmp?.Dispose()
                            _gameHashHistory.RemoveAt(_gameHashHistory.Count - 1)
                            RaiseEvent ProgressChanged(index + 1, totalCount, item.Name, $"AI: 보드 변경 감지 → 재계산...", currentRepeat, totalRepeat)
                            Application.DoEvents()
                            Continue While
                        End If

                        ' 하이라이트 재확인
                        Dim pcGridPos = recognizer.GetGridPositions()
                        Dim pcCellSize = recognizer.GetCellSize()
                        If pcGridPos IsNot Nothing Then
                            Dim pcGlow As String = Nothing
                            Dim pcMyTurn = HasGlowAroundMyPieces(preClickMat, preBoard, pcGridPos, aiSide, pcCellSize.Item1, pcCellSize.Item2, recognizer.IsBoardFlipped, targetWindow, pcGlow)
                            If Not pcMyTurn Then
                                preClickMat.Dispose()
                                preClickSs.Dispose()
                                capturedBmp?.Dispose()
                                _gameHashHistory.RemoveAt(_gameHashHistory.Count - 1)
                                RaiseEvent ProgressChanged(index + 1, totalCount, item.Name, $"AI: 클릭 전 재확인 실패 ({pcGlow}) → 재대기...", currentRepeat, totalRepeat)
                                Application.DoEvents()
                                Continue While
                            End If
                        End If
                    End If
                Catch
                Finally
                    preClickMat?.Dispose()
                    preClickSs.Dispose()
                End Try
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

            ' AI 수 히스토리에 기록 (동일구간 반복 방지)
            _gameAIMoves.Add((bestMove.Item1, bestMove.Item2))

            RaiseEvent ProgressChanged(index + 1, totalCount, item.Name, $"AI: 수 완료 ({moveInfo}) → 상대 차례 대기...", currentRepeat, totalRepeat)
            Application.DoEvents()

            capturedBmp?.Dispose()

            ' AI 수 실행 후 대기 (1초) → 재캡처 진행
            Dim dw = 0
            While dw < 1000 AndAlso Not _cancelRequested
                Thread.Sleep(100)
                dw += 100
                Application.DoEvents()
            End While

            End While ' 다시 WaitForMyTurn부터
            Return False
        End Function

        ''' <summary>
        ''' 매크로를 .macro 파일 + 동명 폴더(템플릿)로 저장
        ''' </summary>
        Public Shared Sub SaveToFile(macroFilePath As String, items As List(Of MacroItem), Optional windowTitle As String = Nothing)
            Dim dir = IO.Path.GetDirectoryName(macroFilePath)
            Dim baseName = IO.Path.GetFileNameWithoutExtension(macroFilePath)
            Dim imgFolder = IO.Path.Combine(dir, baseName)

            ' 이미지 폴더 생성 (기존 폴더는 유지)
            IO.Directory.CreateDirectory(imgFolder)

            Dim lines As New List(Of String)
            ' 첫 줄: 창 이름 (WINDOW| 접두사)
            If Not String.IsNullOrEmpty(windowTitle) Then
                lines.Add($"WINDOW|{windowTitle}")
            End If
            For i = 0 To items.Count - 1
                Dim item = items(i)
                If item.IsAI AndAlso item.Image IsNot Nothing AndAlso item.Image.Width > 1 Then
                    ' AI패턴 항목: 이미지파일|이름|AI패턴|임계값|버튼|키|클릭X|클릭Y|창
                    Dim imgFileName = $"{SanitizeFileName(item.Name)}.png"
                    lines.Add($"{imgFileName}|{item.Name}|AI패턴|{item.Threshold:F2}|{CInt(item.Button)}|{item.SendKeys}|{item.ClickOffsetX}|{item.ClickOffsetY}|{item.WindowTitle}")
                    item.Image.Save(IO.Path.Combine(imgFolder, imgFileName), Imaging.ImageFormat.Png)
                    If item.Mask IsNot Nothing Then
                        Dim maskFileName = $"{SanitizeFileName(item.Name)}_mask.png"
                        item.Mask.Save(IO.Path.Combine(imgFolder, maskFileName), Imaging.ImageFormat.Png)
                    End If
                ElseIf item.IsAI Then
                    ' 순수 AI 항목: AI|이름|대기(초)|임계값|버튼|키|AI진영|AI깊이|AI시간|창
                    Dim aiDelaySec = item.DelayAfterClick / 1000.0
                    lines.Add($"AI|{item.Name}|{aiDelaySec:F1}|{item.Threshold:F2}|0||{item.AISide}|{item.AIDepth}|{item.AITime:F1}|{item.WindowTitle}")
                Else
                    Dim imgFileName = $"{SanitizeFileName(item.Name)}.png"
                    Dim delaySec = item.DelayAfterClick / 1000.0
                    lines.Add($"{imgFileName}|{item.Name}|{delaySec:F1}|{item.Threshold:F2}|{CInt(item.Button)}|{item.SendKeys}|{item.ClickOffsetX}|{item.ClickOffsetY}|{item.WindowTitle}")
                    ' 이미지를 폴더에 저장
                    item.Image.Save(IO.Path.Combine(imgFolder, imgFileName), Imaging.ImageFormat.Png)
                    ' 마스크 저장
                    If item.Mask IsNot Nothing Then
                        Dim maskFileName = $"{SanitizeFileName(item.Name)}_mask.png"
                        item.Mask.Save(IO.Path.Combine(imgFolder, maskFileName), Imaging.ImageFormat.Png)
                    End If
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
                    Dim aiDelayVal As Double
                    Double.TryParse(parts(2).Trim(), aiDelayVal)
                    aiItem.DelayAfterClick = If(aiDelayVal < 100, CInt(aiDelayVal * 1000), CInt(aiDelayVal))
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
                    If parts.Length >= 10 Then
                        aiItem.WindowTitle = parts(9).Trim()
                    End If
                    items.Add(aiItem)
                    Continue For
                End If

                ' AI패턴 항목 감지: 3번째 필드가 "AI패턴"
                Dim isAIPattern = (parts.Length >= 3 AndAlso parts(2).Trim() = "AI패턴")

                Dim item As New MacroItem()
                item.Name = parts(1).Trim()

                ' 이미지 로드: 이름 기반 파일명 우선, 없으면 저장된 파일명(구형식 호환)
                Dim imgPath = IO.Path.Combine(imgFolder, $"{SanitizeFileName(item.Name)}.png")
                If Not IO.File.Exists(imgPath) Then
                    imgPath = IO.Path.Combine(imgFolder, parts(0).Trim())
                End If

                ' 파일 잠금 방지: 메모리로 복사 후 파일 해제
                If IO.File.Exists(imgPath) Then
                    Using tmp As New Bitmap(imgPath)
                        item.Image = New Bitmap(tmp)
                    End Using
                Else
                    ' 이미지 파일 없으면 더미 이미지로 대체 (항목 유지)
                    item.Image = New Bitmap(1, 1)
                End If
                ' 마스크 로드
                Dim maskPath = IO.Path.Combine(imgFolder, $"{SanitizeFileName(item.Name)}_mask.png")
                If IO.File.Exists(maskPath) Then
                    Using tmpMask As New Bitmap(maskPath)
                        item.Mask = New Bitmap(tmpMask)
                    End Using
                End If

                If isAIPattern Then
                    ' AI패턴 항목: IsAI=True, 이미지 유지
                    item.IsAI = True
                    item.DelayAfterClick = 1000
                    item.AISide = "AUTO"
                    item.AIDepth = 1
                    item.AITime = 1.0
                    Double.TryParse(parts(3).Trim(), item.Threshold)
                Else
                    Dim delayVal As Double
                    Double.TryParse(parts(2).Trim(), delayVal)
                    item.DelayAfterClick = If(delayVal < 100, CInt(delayVal * 1000), CInt(delayVal))
                    Double.TryParse(parts(3).Trim(), item.Threshold)
                End If
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
                If parts.Length >= 9 Then
                    item.WindowTitle = parts(8).Trim()
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

        ''' <summary>
        ''' 기물 위치 변동 시 스크린샷 저장. 이전 보드와 비교하여 변동이 있을 때만 저장.
        ''' </summary>
        Private Sub SaveWatchIfChanged(screenshot As Bitmap, board As Board)
            If _watchDateDir Is Nothing Then Return
            If board Is Nothing Then Return

            ' 이전 보드와 비교: 변동 없으면 저장하지 않음
            If _watchPreviousBoard IsNot Nothing Then
                Dim changed = False
                For r = 0 To Constants.BOARD_ROWS - 1
                    For c = 0 To Constants.BOARD_COLS - 1
                        If board.Grid(r)(c) <> _watchPreviousBoard.Grid(r)(c) Then
                            changed = True
                            Exit For
                        End If
                    Next
                    If changed Then Exit For
                Next
                If Not changed Then Return
            End If

            Try
                If _watchDir Is Nothing Then
                    Dim maxNum = 0
                    For Each d In IO.Directory.GetDirectories(_watchDateDir)
                        Dim dn = IO.Path.GetFileName(d)
                        Dim num As Integer
                        If Integer.TryParse(dn, num) Then
                            If num > maxNum Then maxNum = num
                        End If
                    Next
                    _watchDir = IO.Path.Combine(_watchDateDir, (maxNum + 1).ToString("D3"))
                    IO.Directory.CreateDirectory(_watchDir)
                    Console.WriteLine($"[알림] 스크린샷 저장 폴더: {_watchDir}")
                End If
                _watchCounter += 1
                Dim savePath = IO.Path.Combine(_watchDir, $"{_watchCounter.ToString("D4")}.jpg")
                screenshot.Save(savePath, System.Drawing.Imaging.ImageFormat.Jpeg)
            Catch
            End Try

            _watchPreviousBoard = board
        End Sub

        Friend Shared Function SanitizeFileName(name As String) As String
            Dim invalid = IO.Path.GetInvalidFileNameChars()
            Dim result = name
            For Each c In invalid
                result = result.Replace(c, "_"c)
            Next
            Return result
        End Function
    End Class
End Namespace
