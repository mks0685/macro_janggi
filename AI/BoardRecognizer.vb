' 보드 상태 인식 (동적 보드 감지 + 템플릿 매칭)
Imports OpenCvSharp
Imports MacroAutoControl.Constants
Imports MacroAutoControl.Engine

Namespace MacroAutoControl.Capture

    Public Class BoardRecognizer

        Private _templates As New Dictionary(Of String, Mat)
        Private _templatesGray As New Dictionary(Of String, Mat)
        Private _matchThreshold As Double = 0.55
        Private _gridPositions As Integer()() = Nothing
        Private _cellSize As (Integer, Integer) = (0, 0)
        Private _boardRegion As (Integer, Integer, Integer, Integer) = (0, 0, 0, 0)
        Private _lastImage As Mat = Nothing
        Private _gridLocked As Boolean = False
        Private _boardFlipped As Boolean = False
        Private _buttonTemplates As New Dictionary(Of String, (Color As Mat, Gray As Mat))
        Private _resultTemplates As New Dictionary(Of String, (Color As Mat, Gray As Mat))

        Private Shared ReadOnly TEMPLATES_DIR As String =
            IO.Path.Combine(IO.Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().Location), "templates")

        Private Shared ReadOnly ALL_PIECES As String() =
            {CK, CS, CC, CM, CE, CP, CJ, HK, HS, HC, HM, HE, HP, HB}

        Public Shared ReadOnly PIECE_NAMES As New Dictionary(Of String, String) From {
            {CK, "초궁"}, {CS, "초사"}, {CC, "초차"}, {CM, "초마"},
            {CE, "초상"}, {CP, "초포"}, {CJ, "초졸"},
            {HK, "한궁"}, {HS, "한사"}, {HC, "한차"}, {HM, "한마"},
            {HE, "한상"}, {HP, "한포"}, {HB, "한병"}
        }

        Private Shared ReadOnly _crossSideMap As New Dictionary(Of String, String) From {
            {CC, HC}, {HC, CC}, {CS, HS}, {HS, CS},
            {CM, HM}, {HM, CM}, {CE, HE}, {HE, CE},
            {CP, HP}, {HP, CP}
        }

        Private Shared ReadOnly _pieceSide As New Dictionary(Of String, String) From {
            {CK, CHO}, {CS, CHO}, {CC, CHO}, {CM, CHO}, {CE, CHO}, {CP, CHO}, {CJ, CHO},
            {HK, HAN}, {HS, HAN}, {HC, HAN}, {HM, HAN}, {HE, HAN}, {HP, HAN}, {HB, HAN}
        }

        Private Shared ReadOnly BUTTON_NAMES As New Dictionary(Of String, String) From {
            {"btn_refresh", "최신정보갱신"},
            {"btn_match", "대국신청"}
        }

        Private Shared ReadOnly RESULT_NAMES As New Dictionary(Of String, String) From {
            {"result_win", "승리"},
            {"result_lose", "패배"},
            {"result_draw", "무승부"},
            {"result_resign_win", "기권승"},
            {"result_resign_lose", "기권패"},
            {"result_timeout_win", "시간승"},
            {"result_timeout_lose", "시간패"},
            {"result_invalid", "무효"}
        }

        Public Function LoadTemplates(Optional templatesDir As String = Nothing) As Integer
            If templatesDir Is Nothing Then templatesDir = TEMPLATES_DIR
            If Not IO.Directory.Exists(templatesDir) Then
                Console.WriteLine($"[오류] 템플릿 디렉토리가 없습니다: {templatesDir}")
                Return 0
            End If

            Dim loaded = 0
            For Each pieceCode In ALL_PIECES
                Dim path = IO.Path.Combine(templatesDir, $"{pieceCode}.png")
                If IO.File.Exists(path) Then
                    Dim tmpl = Cv2.ImRead(path, ImreadModes.Color)
                    If tmpl IsNot Nothing AndAlso Not tmpl.Empty() Then
                        _templates(pieceCode) = tmpl
                        Dim gray As New Mat()
                        Cv2.CvtColor(tmpl, gray, ColorConversionCodes.BGR2GRAY)
                        _templatesGray(pieceCode) = gray
                        loaded += 1
                    End If
                End If
            Next
            Console.WriteLine($"[알림] 템플릿 {loaded}/{ALL_PIECES.Length}개 로드됨")
            Return loaded
        End Function

        Private Function DetectBoardRect(image As Mat) As OpenCvSharp.Rect?
            Dim gray As New Mat()
            Cv2.CvtColor(image, gray, ColorConversionCodes.BGR2GRAY)
            Dim thresh As New Mat()
            Cv2.Threshold(gray, thresh, 150, 255, ThresholdTypes.Binary)
            Dim kernel = Cv2.GetStructuringElement(MorphShapes.Rect, New OpenCvSharp.Size(10, 10))
            Cv2.MorphologyEx(thresh, thresh, MorphTypes.Close, kernel)
            Cv2.MorphologyEx(thresh, thresh, MorphTypes.Open, kernel)

            Dim contours As Point()()
            Dim hierarchy As HierarchyIndex()
            Cv2.FindContours(thresh, contours, hierarchy, RetrievalModes.External, ContourApproximationModes.ApproxSimple)

            Dim h = image.Rows, w = image.Cols
            Dim minArea = w * h * 0.1
            Dim largeContours = contours.Where(Function(cnt) Cv2.ContourArea(cnt) > minArea).ToArray()
            If largeContours.Length = 0 Then
                Return Nothing
            End If

            Dim largest = largeContours.OrderByDescending(Function(cnt) Cv2.ContourArea(cnt)).First()
            Return Cv2.BoundingRect(largest)
        End Function

        Private Function FindGridLines(image As Mat, boardRect As OpenCvSharp.Rect) As (HLines As Integer(), VLines As Integer())
            Dim margin = 10
            Dim crop = image(New OpenCvSharp.Rect(boardRect.X + margin, boardRect.Y + margin,
                                       boardRect.Width - 2 * margin, boardRect.Height - 2 * margin))
            Dim gray As New Mat()
            Cv2.CvtColor(crop, gray, ColorConversionCodes.BGR2GRAY)

            Dim sobelX As New Mat()
            Cv2.Sobel(gray, sobelX, MatType.CV_64F, 1, 0, 3)
            Dim absSobelX As New Mat()
            Cv2.ConvertScaleAbs(sobelX, absSobelX)
            Dim vProfile = GetColumnProfile(absSobelX)
            Dim vDist = Math.Max(20, boardRect.Width \ 12)
            Dim vPeaks = FindPeaks(vProfile, vDist, 3)
            Dim vLines = vPeaks.Select(Function(p) p + boardRect.X + margin).ToArray()

            Dim sobelY As New Mat()
            Cv2.Sobel(gray, sobelY, MatType.CV_64F, 0, 1, 3)
            Dim absSobelY As New Mat()
            Cv2.ConvertScaleAbs(sobelY, absSobelY)
            Dim hProfile = GetRowProfile(absSobelY)
            Dim hDist = Math.Max(20, boardRect.Height \ 14)
            Dim hPeaks = FindPeaks(hProfile, hDist, 3)
            Dim hLines = hPeaks.Select(Function(p) p + boardRect.Y + margin).ToArray()

            Return (hLines, vLines)
        End Function

        Private Function GetColumnProfile(mat As Mat) As Double()
            Dim profile(mat.Cols - 1) As Double
            For c = 0 To mat.Cols - 1
                Dim sum As Double = 0
                For r = 0 To mat.Rows - 1
                    sum += mat.Get(Of Byte)(r, c)
                Next
                profile(c) = sum / mat.Rows
            Next
            Return profile
        End Function

        Private Function GetRowProfile(mat As Mat) As Double()
            Dim profile(mat.Rows - 1) As Double
            For r = 0 To mat.Rows - 1
                Dim sum As Double = 0
                For c = 0 To mat.Cols - 1
                    sum += mat.Get(Of Byte)(r, c)
                Next
                profile(r) = sum / mat.Cols
            Next
            Return profile
        End Function

        Private Function FindPeaks(profile As Double(), distance As Integer, prominence As Double) As Integer()
            Dim peaks As New List(Of Integer)
            For i = 1 To profile.Length - 2
                If profile(i) > profile(i - 1) AndAlso profile(i) > profile(i + 1) AndAlso profile(i) >= prominence Then
                    peaks.Add(i)
                End If
            Next
            Dim filtered As New List(Of Integer)
            For Each p In peaks
                If filtered.Count = 0 OrElse p - filtered.Last() >= distance Then
                    filtered.Add(p)
                Else
                    If profile(p) > profile(filtered.Last()) Then
                        filtered(filtered.Count - 1) = p
                    End If
                End If
            Next
            Return filtered.ToArray()
        End Function

        Private Function FitUniformGrid(lines As Integer(), n As Integer, expectedSpacing As Double) As Integer()
            If lines.Length < 2 Then Return Nothing
            If lines.Length < n Then
                Dim first = lines(0), last = lines(lines.Length - 1)
                Return Enumerable.Range(0, n).Select(Function(i) CInt(first + i * (last - first) / (n - 1))).ToArray()
            End If

            Dim bestScore = Double.MaxValue
            Dim bestResult As Integer() = Nothing

            For i = 0 To lines.Length - 1
                For j = i + n - 1 To lines.Length - 1
                    Dim spacing = (lines(j) - lines(i)) / CDbl(n - 1)
                    If Math.Abs(spacing - expectedSpacing) > expectedSpacing * 0.3 Then Continue For

                    Dim expected = Enumerable.Range(0, n).Select(Function(k) lines(i) + CInt(k * spacing)).ToArray()
                    Dim totalErr As Double = 0
                    For Each exp In expected
                        totalErr += lines.Min(Function(l) Math.Abs(l - exp))
                    Next

                    If totalErr < bestScore Then
                        bestScore = totalErr
                        bestResult = expected
                    End If
                Next
            Next
            Return bestResult
        End Function

        Public Function DetectGrid(image As Mat) As Boolean
            Dim rectResult = DetectBoardRect(image)
            If Not rectResult.HasValue Then Return False

            Dim boardRect = rectResult.Value
            Dim h = image.Rows, w = image.Cols
            If boardRect.Width * boardRect.Height < w * h * 0.1 Then Return False

            If _gridPositions IsNot Nothing AndAlso
               (_boardRegion.Item1 <> 0 OrElse _boardRegion.Item2 <> 0 OrElse _boardRegion.Item3 <> 0 OrElse _boardRegion.Item4 <> 0) Then
                If Math.Abs(boardRect.X - _boardRegion.Item1) < 10 AndAlso
                   Math.Abs(boardRect.Y - _boardRegion.Item2) < 10 AndAlso
                   Math.Abs(boardRect.Width - _boardRegion.Item3) < 10 AndAlso
                   Math.Abs(boardRect.Height - _boardRegion.Item4) < 10 Then
                    Return True
                End If
            End If

            Dim gridLines = FindGridLines(image, boardRect)
            Dim hLines = gridLines.HLines
            Dim vLines = gridLines.VLines

            Dim expectedH = boardRect.Height / 9.0
            Dim expectedV = boardRect.Width / 8.0
            Dim hGrid = FitUniformGrid(hLines, 10, expectedH)
            Dim vGrid = FitUniformGrid(vLines, 9, expectedV)
            If hGrid Is Nothing OrElse vGrid Is Nothing Then
                Return False
            End If

            Dim grid As New List(Of Integer())
            For r = 0 To BOARD_ROWS - 1
                For c = 0 To BOARD_COLS - 1
                    grid.Add(New Integer() {vGrid(c), hGrid(r)})
                Next
            Next

            Dim cw = CInt((vGrid.Last() - vGrid.First()) / 8.0)
            Dim ch = CInt((hGrid.Last() - hGrid.First()) / 9.0)

            If cw < 40 OrElse ch < 40 OrElse cw > 100 OrElse ch > 100 Then
                Return False
            End If
            If Math.Abs(cw - ch) > Math.Max(cw, ch) * 0.5 Then
                Return False
            End If

            _gridPositions = grid.ToArray()
            _cellSize = (cw, ch)
            _boardRegion = (boardRect.X, boardRect.Y, boardRect.Width, boardRect.Height)
            _gridLocked = True
            Return True
        End Function

        Public Sub ResetGrid()
            _gridPositions = Nothing
            _gridLocked = False
            _boardRegion = (0, 0, 0, 0)
            _boardFlipped = False
        End Sub

        ''' <summary>
        ''' 보드가 상하 반전되었는지 여부 (원래 화면에서 초궁이 위에 있었음)
        ''' </summary>
        Public ReadOnly Property IsBoardFlipped As Boolean
            Get
                Return _boardFlipped
            End Get
        End Property

        ''' <summary>
        ''' 정규화된 행 번호를 원래 화면 행 번호로 변환.
        ''' 보드가 뒤집혔으면 9-row, 아니면 row 그대로 반환.
        ''' </summary>
        Public Function TranslateRow(row As Integer) As Integer
            If _boardFlipped Then Return 9 - row
            Return row
        End Function

        ''' <summary>
        ''' 화면 하단 궁의 HSV 색상으로 내 진영 판단.
        ''' 하단 궁 영역(row 8~9, col 3~5)의 실제 픽셀 색상을 검사.
        ''' 빨강이면 CHO, 파랑/녹색이면 HAN.
        ''' </summary>
        Public Function DetectMySideByBottomKingColor(image As Mat) As String
            If _gridPositions Is Nothing Then Return Nothing

            ' 화면 하단 = 정규화된 row 8~9 (뒤집혔으면 screenR = 9-r, 아니면 r 그대로)
            ' 화면 하단 궁 위치: screenRow 8~9, col 3~5
            For screenR = 9 To 8 Step -1
                For c = 3 To 5
                    Dim idx = screenR * BOARD_COLS + c
                    If idx < 0 OrElse idx >= _gridPositions.Length Then Continue For
                    Dim cx = _gridPositions(idx)(0)
                    Dim cy = _gridPositions(idx)(1)

                    Dim half = 26
                    If _cellSize.Item1 > 0 Then
                        half = Math.Max(26, CInt(Math.Min(_cellSize.Item1, _cellSize.Item2) * 0.45))
                    End If

                    Dim imgH = image.Rows, imgW = image.Cols
                    Dim x1 = Math.Max(0, cx - half)
                    Dim y1 = Math.Max(0, cy - half)
                    Dim x2 = Math.Min(imgW, cx + half)
                    Dim y2 = Math.Min(imgH, cy + half)
                    If x2 <= x1 OrElse y2 <= y1 Then Continue For

                    Dim cellImg = image(New OpenCvSharp.Rect(x1, y1, x2 - x1, y2 - y1))
                    Dim color = DetectCellColor(cellImg)
                    If color IsNot Nothing Then Return color
                Next
            Next
            Return Nothing
        End Function

        Public Function GetGridPositions() As Integer()()
            Return _gridPositions
        End Function

        Public Function GetCellSize() As (Integer, Integer)
            Return _cellSize
        End Function

        Public Function GetCellImage(image As Mat, row As Integer, col As Integer) As Mat
            If _gridPositions Is Nothing Then Return Nothing
            Dim idx = row * BOARD_COLS + col
            Dim cx = _gridPositions(idx)(0)
            Dim cy = _gridPositions(idx)(1)

            Dim half = 26
            If _cellSize.Item1 > 0 Then
                half = Math.Max(26, CInt(Math.Min(_cellSize.Item1, _cellSize.Item2) * 0.45))
            End If

            Dim imgH = image.Rows, imgW = image.Cols
            Dim x1 = Math.Max(0, cx - half)
            Dim y1 = Math.Max(0, cy - half)
            Dim x2 = Math.Min(imgW, cx + half)
            Dim y2 = Math.Min(imgH, cy + half)
            If x2 <= x1 OrElse y2 <= y1 Then Return Nothing

            Return image(New OpenCvSharp.Rect(x1, y1, x2 - x1, y2 - y1))
        End Function

        Private Function DetectCellColor(cellImage As Mat) As String
            Dim hsv As New Mat()
            Cv2.CvtColor(cellImage, hsv, ColorConversionCodes.BGR2HSV)

            Dim mr1 As New Mat(), mr2 As New Mat()
            Cv2.InRange(hsv, New Scalar(0, 40, 60), New Scalar(12, 255, 255), mr1)
            Cv2.InRange(hsv, New Scalar(158, 40, 60), New Scalar(180, 255, 255), mr2)
            Dim redPx = Cv2.CountNonZero(mr1) + Cv2.CountNonZero(mr2)

            Dim mb As New Mat()
            Cv2.InRange(hsv, New Scalar(85, 40, 60), New Scalar(135, 255, 255), mb)
            Dim bluePx = Cv2.CountNonZero(mb)

            Dim total = cellImage.Rows * cellImage.Cols
            Dim minPx = Math.Max(10, total * 0.02)

            If redPx > minPx AndAlso redPx > bluePx * 1.3 Then Return CHO
            If bluePx > minPx AndAlso bluePx > redPx * 1.3 Then Return HAN
            Return Nothing
        End Function

        Public Function RecognizePiece(cellImage As Mat) As String
            Try
                Return RecognizePieceInternal(cellImage)
            Catch
                Return EMPTY
            End Try
        End Function

        Private Function RecognizePieceInternal(cellImage As Mat) As String
            If cellImage Is Nothing OrElse cellImage.Empty() Then Return EMPTY
            If _templates.Count = 0 Then Return EMPTY

            ' 채널 보정: 4채널(BGRA) → 3채널(BGR)
            If cellImage.Channels() = 4 Then
                Dim bgr As New Mat()
                Cv2.CvtColor(cellImage, bgr, ColorConversionCodes.BGRA2BGR)
                cellImage = bgr
            ElseIf cellImage.Channels() = 1 Then
                Dim bgr As New Mat()
                Cv2.CvtColor(cellImage, bgr, ColorConversionCodes.GRAY2BGR)
                cellImage = bgr
            End If

            Dim bestMatch = EMPTY
            Dim bestScore = _matchThreshold
            Dim ch = cellImage.Rows, cw = cellImage.Cols

            For Each kvp In _templates
                Dim th = kvp.Value.Rows, tw = kvp.Value.Cols
                If th > ch OrElse tw > cw Then Continue For
                Dim result As New Mat()
                Cv2.MatchTemplate(cellImage, kvp.Value, result, TemplateMatchModes.CCoeffNormed)
                Dim maxVal As Double
                result.MinMaxLoc(Nothing, maxVal)
                If maxVal > bestScore Then
                    bestScore = maxVal
                    bestMatch = kvp.Key
                End If
            Next
            If bestScore >= 0.7 Then
                bestMatch = VerifyColorSide(cellImage, bestMatch)
                Return bestMatch
            End If

            For Each scale In {0.88, 0.94, 1.06, 1.12}
                For Each kvp In _templates
                    Dim th = CInt(kvp.Value.Rows * scale)
                    Dim tw = CInt(kvp.Value.Cols * scale)
                    If th > ch OrElse tw > cw OrElse th < 10 OrElse tw < 10 Then Continue For
                    Dim tmplR As New Mat()
                    Cv2.Resize(kvp.Value, tmplR, New OpenCvSharp.Size(tw, th), 0, 0, InterpolationFlags.Area)
                    Dim result As New Mat()
                    Cv2.MatchTemplate(cellImage, tmplR, result, TemplateMatchModes.CCoeffNormed)
                    Dim maxVal As Double
                    result.MinMaxLoc(Nothing, maxVal)
                    If maxVal > bestScore Then
                        bestScore = maxVal
                        bestMatch = kvp.Key
                    End If
                Next
            Next
            If bestScore >= 0.65 Then
                bestMatch = VerifyColorSide(cellImage, bestMatch)
                Return bestMatch
            End If

            Dim cellGray As New Mat()
            Cv2.CvtColor(cellImage, cellGray, ColorConversionCodes.BGR2GRAY)
            Dim cellColor = DetectCellColor(cellImage)
            Dim grayMatch = EMPTY
            Dim grayScore = _matchThreshold

            For Each scale In {0.88, 0.94, 1.0, 1.06, 1.12}
                For Each kvp In _templatesGray
                    Dim th = CInt(kvp.Value.Rows * scale)
                    Dim tw = CInt(kvp.Value.Cols * scale)
                    If th > ch OrElse tw > cw OrElse th < 10 OrElse tw < 10 Then Continue For
                    Dim gr As New Mat()
                    Cv2.Resize(kvp.Value, gr, New OpenCvSharp.Size(tw, th), 0, 0, InterpolationFlags.Area)
                    Dim result As New Mat()
                    Cv2.MatchTemplate(cellGray, gr, result, TemplateMatchModes.CCoeffNormed)
                    Dim maxVal As Double
                    result.MinMaxLoc(Nothing, maxVal)
                    If maxVal > grayScore Then
                        grayScore = maxVal
                        grayMatch = kvp.Key
                    End If
                Next
            Next

            If grayMatch <> EMPTY AndAlso grayScore > bestScore Then
                Dim matchedSide As String = Nothing
                _pieceSide.TryGetValue(grayMatch, matchedSide)
                If cellColor IsNot Nothing AndAlso matchedSide <> cellColor Then
                    Dim cross As String = Nothing
                    If _crossSideMap.TryGetValue(grayMatch, cross) Then
                        grayMatch = cross
                    End If
                End If
                bestMatch = grayMatch
                bestScore = grayScore
            End If

            Return bestMatch
        End Function

        ''' <summary>
        ''' 템플릿 매칭 결과의 진영을 HSV 색상으로 검증하고, 불일치 시 교차 진영으로 교정.
        ''' 궁(K)/졸(J)/병(B)은 고유 기물이므로 교정 대상이 아님.
        ''' </summary>
        Private Function VerifyColorSide(cellImage As Mat, matchedPiece As String) As String
            If matchedPiece = EMPTY Then Return matchedPiece

            ' 궁/졸/병은 진영 교차 매핑이 없으므로 그대로 반환
            Dim cross As String = Nothing
            If Not _crossSideMap.TryGetValue(matchedPiece, cross) Then Return matchedPiece

            Dim cellColor = DetectCellColor(cellImage)
            If cellColor Is Nothing Then Return matchedPiece  ' 색상 판별 불가 시 그대로

            Dim matchedSide As String = Nothing
            _pieceSide.TryGetValue(matchedPiece, matchedSide)
            If matchedSide <> cellColor Then
                Return cross
            End If
            Return matchedPiece
        End Function

        Public Function GetBoardState(image As Mat) As Board
            If Not DetectGrid(image) Then
                Return Nothing
            End If
            _lastImage = image

            Dim grid As String()() = New String(BOARD_ROWS - 1)() {}
            For r = 0 To BOARD_ROWS - 1
                grid(r) = New String(BOARD_COLS - 1) {}
                For c = 0 To BOARD_COLS - 1
                    Dim cellImg = GetCellImage(image, r, c)
                    grid(r)(c) = RecognizePiece(cellImg)
                Next
            Next

            ' 후처리: 궁성 내 사(士)/궁(宮) 진영 보정
            ' 사와 궁은 궁성(宮城) 밖으로 나갈 수 없으므로
            ' 같은 궁성 안의 궁(King) 진영에 맞춰 보정
            ValidatePalacePieces(grid)

            ' 정규화: 초궁(CK)이 행 0~2에 있으면 보드를 상하 반전
            ' → 항상 "초=아래(행7~9), 한=위(행0~2)"로 만들어 AI 코드와 일치시킴
            _boardFlipped = False
            For r = 0 To 2
                For c = 3 To 5
                    If grid(r)(c) = CK Then
                        _boardFlipped = True
                        Exit For
                    End If
                Next
                If _boardFlipped Then Exit For
            Next
            If _boardFlipped Then
                FlipGrid(grid)
                Console.WriteLine("[정규화] 보드 상하 반전 (초궁이 위에 있었음)")
            End If

            Return New Board(grid)
        End Function

        ''' <summary>
        ''' 궁성 내 사/궁 진영 보정: 같은 궁성의 궁(King) 진영과 일치하도록 수정
        ''' </summary>
        Private Shared Sub ValidatePalacePieces(grid As String()())
            ' 상단 궁성 (행 0~2, 열 3~5) 궁 진영 파악
            Dim topKingSide As String = Nothing
            For r = 0 To 2
                For c = 3 To 5
                    If grid(r)(c) = CK Then topKingSide = CHO
                    If grid(r)(c) = HK Then topKingSide = HAN
                Next
            Next

            ' 하단 궁성 (행 7~9, 열 3~5) 궁 진영 파악
            Dim bottomKingSide As String = Nothing
            For r = 7 To 9
                For c = 3 To 5
                    If grid(r)(c) = CK Then bottomKingSide = CHO
                    If grid(r)(c) = HK Then bottomKingSide = HAN
                Next
            Next

            ' 상단 궁성: 사 진영 보정
            If topKingSide IsNot Nothing Then
                For r = 0 To 2
                    For c = 3 To 5
                        If grid(r)(c) = CS AndAlso topKingSide = HAN Then grid(r)(c) = HS
                        If grid(r)(c) = HS AndAlso topKingSide = CHO Then grid(r)(c) = CS
                    Next
                Next
            End If

            ' 하단 궁성: 사 진영 보정
            If bottomKingSide IsNot Nothing Then
                For r = 7 To 9
                    For c = 3 To 5
                        If grid(r)(c) = CS AndAlso bottomKingSide = HAN Then grid(r)(c) = HS
                        If grid(r)(c) = HS AndAlso bottomKingSide = CHO Then grid(r)(c) = CS
                    Next
                Next
            End If

            ' 궁성 밖의 사/궁은 인식 오류 → 빈 칸으로 처리
            For r = 0 To BOARD_ROWS - 1
                For c = 0 To BOARD_COLS - 1
                    Dim p = grid(r)(c)
                    If p = CS OrElse p = HS OrElse p = CK OrElse p = HK Then
                        Dim inTopPalace = (r >= 0 AndAlso r <= 2 AndAlso c >= 3 AndAlso c <= 5)
                        Dim inBottomPalace = (r >= 7 AndAlso r <= 9 AndAlso c >= 3 AndAlso c <= 5)
                        If Not inTopPalace AndAlso Not inBottomPalace Then
                            grid(r)(c) = EMPTY
                        End If
                    End If
                Next
            Next
        End Sub

        ''' <summary>
        ''' 보드 그리드를 상하 반전 (행 0↔9, 1↔8, 2↔7, 3↔6, 4↔5)
        ''' </summary>
        Private Shared Sub FlipGrid(grid As String()())
            For i = 0 To 4
                Dim temp = grid(i)
                grid(i) = grid(9 - i)
                grid(9 - i) = temp
            Next
        End Sub

        ' ───── 보드 유효성 검증 ─────

        ''' <summary>
        ''' 인식된 보드 상태의 구조적 유효성을 검증.
        ''' 궁 수, 기물 유형별 최대 수 등을 확인.
        ''' </summary>
        Public Function ValidateGridWithPieces(board As Board) As Boolean
            If board Is Nothing Then Return False

            Dim pieceCounts As New Dictionary(Of String, Integer)
            For Each p In ALL_PIECES
                pieceCounts(p) = 0
            Next

            For r = 0 To BOARD_ROWS - 1
                For c = 0 To BOARD_COLS - 1
                    Dim piece = board.Grid(r)(c)
                    If piece = EMPTY Then Continue For
                    If pieceCounts.ContainsKey(piece) Then
                        pieceCounts(piece) += 1
                    Else
                        Console.WriteLine($"  [검증] 알 수 없는 기물: '{piece}' at ({r},{c})")
                        Return False
                    End If
                Next
            Next

            ' 궁은 각각 정확히 1개
            If pieceCounts(CK) <> 1 OrElse pieceCounts(HK) <> 1 Then
                Console.WriteLine($"  [검증] 궁 수 이상: 초궁={pieceCounts(CK)}, 한궁={pieceCounts(HK)}")
                Return False
            End If

            ' 기물 유형별 최대 수 검증
            Dim maxCounts As New Dictionary(Of String, Integer) From {
                {CC, 2}, {HC, 2}, {CS, 2}, {HS, 2},
                {CM, 2}, {HM, 2}, {CE, 2}, {HE, 2},
                {CP, 2}, {HP, 2}, {CJ, 5}, {HB, 5}
            }
            For Each kvp In maxCounts
                If pieceCounts(kvp.Key) > kvp.Value Then
                    Dim name As String = ""
                    PIECE_NAMES.TryGetValue(kvp.Key, name)
                    Console.WriteLine($"  [검증] {name} 수 초과: {pieceCounts(kvp.Key)}개 (최대 {kvp.Value}개)")
                    Return False
                End If
            Next

            Return True
        End Function

        ' ───── 색상 감지 ─────

        Public Function DetectPieceColor(row As Integer, col As Integer) As String
            If _lastImage Is Nothing OrElse _gridPositions Is Nothing Then Return Nothing
            Dim cellImg = GetCellImage(_lastImage, row, col)
            If cellImg Is Nothing OrElse cellImg.Empty() Then Return Nothing
            Return DetectCellColor(cellImg)
        End Function

        ''' <summary>
        ''' 색상 분포 기반 내 진영 감지: 하단에 초(빨강) 많으면 CHO, 한(파랑) 많으면 HAN
        ''' </summary>
        Public Function DetectSideByColor() As String
            If _lastImage Is Nothing OrElse _gridPositions Is Nothing Then Return Nothing

            Dim bottomCho = 0, bottomHan = 0
            Dim topCho = 0, topHan = 0

            For r = 0 To 3
                For c = 0 To 8
                    Dim color = DetectPieceColor(r, c)
                    If color = CHO Then topCho += 1
                    If color = HAN Then topHan += 1
                Next
            Next
            For r = 6 To 9
                For c = 0 To 8
                    Dim color = DetectPieceColor(r, c)
                    If color = CHO Then bottomCho += 1
                    If color = HAN Then bottomHan += 1
                Next
            Next

            Console.WriteLine($"[색상감지] 하단: 초={bottomCho}, 한={bottomHan} | 상단: 초={topCho}, 한={topHan}")

            If bottomCho > bottomHan AndAlso bottomCho >= 2 Then Return CHO
            If bottomHan > bottomCho AndAlso bottomHan >= 2 Then Return HAN
            If topHan > topCho AndAlso topHan >= 2 Then Return CHO
            If topCho > topHan AndAlso topCho >= 2 Then Return HAN
            Return Nothing
        End Function

        ' ───── 보드 비교 ─────

        Public Function CompareBoards(oldBoard As Board, newBoard As Board) As List(Of (Integer, Integer, String, String))
            Dim changes As New List(Of (Integer, Integer, String, String))
            For r = 0 To BOARD_ROWS - 1
                For c = 0 To BOARD_COLS - 1
                    Dim oldPiece = oldBoard.Grid(r)(c)
                    Dim newPiece = newBoard.Grid(r)(c)
                    If oldPiece <> newPiece Then
                        changes.Add((r, c, oldPiece, newPiece))
                    End If
                Next
            Next
            Return changes
        End Function

        Public Function CountPieces(board As Board) As Integer
            Dim count = 0
            For r = 0 To BOARD_ROWS - 1
                For c = 0 To BOARD_COLS - 1
                    If board.Grid(r)(c) <> EMPTY Then count += 1
                Next
            Next
            Return count
        End Function

        ' ───── 하이라이트 감지 ─────

        ''' <summary>
        ''' HSV 기반으로 기물 주변의 금색 빛남 픽셀 수를 측정.
        ''' 링 영역(innerR ~ outerR)을 촘촘하게 샘플링하여 정확도 향상.
        ''' </summary>
        Private Function CountRingGlowPixels(hsv As Mat, cx As Integer, cy As Integer,
                                              innerR As Integer, outerR As Integer) As Integer
            Dim glowCount = 0
            Dim imgH = hsv.Rows, imgW = hsv.Cols

            ' 링 영역을 2px 간격으로 촘촘하게 샘플링
            For dy = -outerR To outerR Step 2
                For dx = -outerR To outerR Step 2
                    Dim dist = Math.Sqrt(dx * dx + dy * dy)
                    If dist < innerR OrElse dist > outerR Then Continue For

                    Dim px = cx + dx
                    Dim py = cy + dy
                    If px < 0 OrElse px >= imgW OrElse py < 0 OrElse py >= imgH Then Continue For

                    Dim pixel(2) As Byte
                    Runtime.InteropServices.Marshal.Copy(
                        hsv.Data + (py * CInt(hsv.Step()) + px * 3), pixel, 0, 3)
                    Dim h = CInt(pixel(0)), s = CInt(pixel(1)), v = CInt(pixel(2))

                    ' 빛남 효과: 노란~금색 (H:15~35) + 채도 있음(S>40) + 밝음(V>200)
                    If h >= 15 AndAlso h <= 35 AndAlso s > 40 AndAlso v > 200 Then
                        glowCount += 1
                    End If
                Next
            Next
            Return glowCount
        End Function

        ''' <summary>
        ''' 보드에서 금색 빛남 효과가 가장 강한 기물의 위치와 진영을 반환.
        ''' 카카오장기에서 마지막 착수 기물에 금색 하이라이트가 표시됨.
        ''' HSV 색상 기반으로 정확하게 검출.
        ''' </summary>
        Public Function DetectHighlightedPiece(image As Mat, board As Board) As (Row As Integer, Col As Integer, Side As String, GlowCount As Integer)?
            If _gridPositions Is Nothing Then Return Nothing

            ' BGR → HSV 변환
            Dim hsv As New Mat()
            Cv2.CvtColor(image, hsv, ColorConversionCodes.BGR2HSV)

            ' 링 크기 계산 (기물 바깥 영역)
            Dim baseSize = Math.Max(_cellSize.Item1, _cellSize.Item2)
            If baseSize <= 0 Then baseSize = 60
            Dim innerR = CInt(baseSize * 0.42)
            Dim outerR = CInt(baseSize * 0.72)
            If innerR < 18 Then innerR = 18
            If outerR < 30 Then outerR = 30

            Dim bestGlow = 0
            Dim bestRow = -1, bestCol = -1
            Dim secondGlow = 0

            For r = 0 To BOARD_ROWS - 1
                For c = 0 To BOARD_COLS - 1
                    If board.Grid(r)(c) = EMPTY Then Continue For

                    Dim screenR = TranslateRow(r)
                    Dim idx = screenR * BOARD_COLS + c
                    Dim cx = _gridPositions(idx)(0)
                    Dim cy = _gridPositions(idx)(1)

                    Dim glow = CountRingGlowPixels(hsv, cx, cy, innerR, outerR)

                    If glow > bestGlow Then
                        secondGlow = bestGlow
                        bestGlow = glow
                        bestRow = r
                        bestCol = c
                    ElseIf glow > secondGlow Then
                        secondGlow = glow
                    End If
                Next
            Next
            hsv.Dispose()

            ' 최소 빛남 임계값: 10픽셀 이상 + 2위 대비 2배 이상 차이
            If bestRow < 0 OrElse bestGlow < 10 Then Return Nothing
            If secondGlow > 0 AndAlso bestGlow < secondGlow * 2 Then
                ' 1위와 2위 차이가 불분명하면 신뢰도 낮음 → 그래도 1위 반환 (임계값은 충족)
                ' 단, 1위가 충분히 강하면 (20 이상) 그대로 반환
                If bestGlow < 20 Then Return Nothing
            End If

            Dim piece = board.Grid(bestRow)(bestCol)
            Dim side As String = Nothing
            If CHO_PIECES.Contains(piece) Then side = CHO
            If HAN_PIECES.Contains(piece) Then side = HAN

            Return (bestRow, bestCol, side, bestGlow)
        End Function

        ' ───── 대기 화면 / 로비 감지 ─────

        Public Function LoadButtonTemplates(Optional templatesDir As String = Nothing) As Integer
            If templatesDir Is Nothing Then templatesDir = TEMPLATES_DIR
            Dim loaded = 0
            For Each btnKey In BUTTON_NAMES.Keys
                Dim path = IO.Path.Combine(templatesDir, $"{btnKey}.png")
                If IO.File.Exists(path) Then
                    Dim tmpl = Cv2.ImRead(path, ImreadModes.Color)
                    If tmpl IsNot Nothing AndAlso Not tmpl.Empty() Then
                        Dim gray As New Mat()
                        Cv2.CvtColor(tmpl, gray, ColorConversionCodes.BGR2GRAY)
                        _buttonTemplates(btnKey) = (tmpl, gray)
                        loaded += 1
                    End If
                End If
            Next
            If loaded > 0 Then
                Dim names = _buttonTemplates.Keys.Select(Function(k) BUTTON_NAMES(k))
                Console.WriteLine($"[알림] 버튼 템플릿 {loaded}개 로드됨: {String.Join(", ", names)}")
            End If
            Return loaded
        End Function

        Public Function DetectLobby(image As Mat, Optional threshold As Double = 0.7) As List(Of String)
            Dim found As New List(Of String)
            If _buttonTemplates.Count = 0 Then Return found

            Dim gray As New Mat()
            Cv2.CvtColor(image, gray, ColorConversionCodes.BGR2GRAY)

            For Each kvp In _buttonTemplates
                Dim result As New Mat()
                Cv2.MatchTemplate(gray, kvp.Value.Gray, result, TemplateMatchModes.CCoeffNormed)
                Dim maxVal As Double
                result.MinMaxLoc(Nothing, maxVal)
                If maxVal >= threshold Then
                    found.Add(BUTTON_NAMES(kvp.Key))
                End If
            Next
            Return found
        End Function

        ' ───── 게임 결과 감지 ─────

        Public Function LoadResultTemplates(Optional templatesDir As String = Nothing) As Integer
            If templatesDir Is Nothing Then templatesDir = TEMPLATES_DIR
            Dim loaded = 0
            For Each key In RESULT_NAMES.Keys
                Dim path = IO.Path.Combine(templatesDir, $"{key}.png")
                If IO.File.Exists(path) Then
                    Dim tmpl = Cv2.ImRead(path, ImreadModes.Color)
                    If tmpl IsNot Nothing AndAlso Not tmpl.Empty() Then
                        Dim gray As New Mat()
                        Cv2.CvtColor(tmpl, gray, ColorConversionCodes.BGR2GRAY)
                        _resultTemplates(key) = (tmpl, gray)
                        loaded += 1
                    End If
                End If
            Next
            If loaded > 0 Then
                Dim names = _resultTemplates.Keys.Select(Function(k) RESULT_NAMES(k))
                Console.WriteLine($"[알림] 결과 템플릿 {loaded}개 로드됨: {String.Join(", ", names)}")
            End If
            Return loaded
        End Function

        ''' <summary>
        ''' 게임 결과 메시지 창을 감지. 매칭된 결과 이름을 반환 (없으면 Nothing).
        ''' </summary>
        Public Function DetectGameResult(image As Mat, Optional threshold As Double = 0.7) As String
            If _resultTemplates.Count = 0 Then Return Nothing

            Dim gray As New Mat()
            Cv2.CvtColor(image, gray, ColorConversionCodes.BGR2GRAY)

            Dim bestName As String = Nothing
            Dim bestScore As Double = threshold

            For Each kvp In _resultTemplates
                Dim tmplC = kvp.Value.Color
                Dim tmplG = kvp.Value.Gray

                If tmplC.Rows > image.Rows OrElse tmplC.Cols > image.Cols Then Continue For

                ' 컬러 매칭
                Dim resultC As New Mat()
                Cv2.MatchTemplate(image, tmplC, resultC, TemplateMatchModes.CCoeffNormed)
                Dim maxC As Double
                resultC.MinMaxLoc(Nothing, maxC)

                ' 그레이 매칭
                Dim resultG As New Mat()
                Cv2.MatchTemplate(gray, tmplG, resultG, TemplateMatchModes.CCoeffNormed)
                Dim maxG As Double
                resultG.MinMaxLoc(Nothing, maxG)

                Dim maxVal = Math.Max(maxC, maxG)
                If maxVal > bestScore Then
                    bestScore = maxVal
                    bestName = RESULT_NAMES(kvp.Key)
                End If
            Next

            Return bestName
        End Function

        Public ReadOnly Property IsCalibrated As Boolean
            Get
                Return _templates.Count > 0
            End Get
        End Property

    End Class

End Namespace
