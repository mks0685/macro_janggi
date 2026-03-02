' NegaMax + Alpha-Beta 탐색 엔진 (강화판)
Imports MacroAutoControl.Constants
Imports MacroAutoControl.Engine
Imports MacroAutoControl.AI.Evaluator

Namespace MacroAutoControl.AI

    ' TT 엔트리 (Age 필드 추가)
    Public Structure TTEntry
        Public Depth As Integer
        Public Score As Integer
        Public Flag As Integer
        Public BestMove As ((Integer, Integer), (Integer, Integer))?
        Public Age As Integer
    End Structure

    Public Module Search

        Private Const INF As Integer = 999999
        Private Const TT_EXACT As Integer = 0
        Private Const TT_ALPHA As Integer = 1
        Private Const TT_BETA As Integer = 2
        Private Const TT_MAX_SIZE As Integer = 1000000
        Private Const QS_MAX_DEPTH As Integer = 6
        Private Const DELTA_MARGIN As Integer = 200
        Private ReadOnly FUTILITY_MARGINS As Integer() = {0, 200, 450}
        Private Const HISTORY_MAX As Integer = 16000

        Private _tt As New Dictionary(Of ULong, TTEntry)
        Private _killers As New Dictionary(Of Integer, (((Integer, Integer), (Integer, Integer))?, ((Integer, Integer), (Integer, Integer))?))
        Private _historyTable As New Dictionary(Of (Integer, Integer, Integer, Integer), Integer)
        ' Countermove: 직전 수(from,to) → 반격 수
        Private _countermoves As New Dictionary(Of (Integer, Integer, Integer, Integer), ((Integer, Integer), (Integer, Integer)))

        Private _deadline As DateTime
        Private _softDeadline As DateTime
        Private _cancelRequested As Boolean = False
        Private _searchAge As Integer = 0

        ''' <summary>외부에서 탐색 중단 요청</summary>
        Public Sub CancelSearch()
            _cancelRequested = True
        End Sub

        Private Function TimeExpired() As Boolean
            Return _cancelRequested OrElse DateTime.Now > _deadline
        End Function

        Private Function SoftTimeExpired() As Boolean
            Return _cancelRequested OrElse DateTime.Now > _softDeadline
        End Function

        Public Function Quiescence(board As Board, alpha As Integer, beta As Integer,
                                    sideToMove As String, ply As Integer,
                                    Optional qsDepth As Integer = 0) As Integer
            If TimeExpired() Then
                Dim raw = Evaluate(board)
                Return If(sideToMove = CHO, raw, -raw)
            End If

            If Not board.FindKing(sideToMove).HasValue Then
                Return -(INF - ply)
            End If

            Dim rawEval = Evaluate(board)
            Dim standPat = If(sideToMove = CHO, rawEval, -rawEval)

            If standPat >= beta Then Return beta
            If standPat > alpha Then alpha = standPat
            If qsDepth >= QS_MAX_DEPTH Then Return standPat

            Dim moves = board.GetAllMoves(sideToMove)
            Dim grid = board.Grid
            Dim captures As New List(Of (Integer, ((Integer, Integer), (Integer, Integer))))

            For Each m In moves
                Dim tr = m.Item2.Item1, tc = m.Item2.Item2
                Dim target = grid(tr)(tc)
                If target <> EMPTY Then
                    Dim victimVal = 0
                    PIECE_VALUES.TryGetValue(target, victimVal)
                    ' Delta pruning
                    If standPat + victimVal + DELTA_MARGIN < alpha Then Continue For
                    ' SEE < 0인 캡처는 QS에서 제외
                    Dim seeVal = board.SEE(m.Item1.Item1, m.Item1.Item2, tr, tc, sideToMove)
                    If seeVal < 0 Then Continue For
                    captures.Add((-seeVal, m))
                End If
            Next

            captures.Sort(Function(a, b) a.Item1.CompareTo(b.Item1))

            Dim nextSide = If(sideToMove = CHO, HAN, CHO)

            For Each cap In captures
                Dim move = cap.Item2
                board.MakeMove(move.Item1, move.Item2)

                If board.IsInCheck(sideToMove) Then
                    board.UndoMove()
                    Continue For
                End If

                Dim score = -Quiescence(board, -beta, -alpha, nextSide, ply + 1, qsDepth + 1)
                board.UndoMove()

                If score >= beta Then Return beta
                If score > alpha Then alpha = score
            Next

            Return alpha
        End Function

        Public Function Negamax(board As Board, depth As Integer, alpha As Integer, beta As Integer,
                                 sideToMove As String, allowNull As Boolean, ply As Integer) As Integer
            If TimeExpired() Then
                Dim raw = Evaluate(board)
                Return If(sideToMove = CHO, raw, -raw)
            End If

            If Not board.FindKing(sideToMove).HasValue Then
                Return -(INF - ply)
            End If

            ' 반복 수 감지
            If ply > 0 AndAlso board.IsRepetition(2) Then Return 0

            Dim origAlpha = alpha
            Dim ttMove As ((Integer, Integer), (Integer, Integer))? = Nothing
            Dim ttKey = board.ZobristHash
            Dim ttEntry As TTEntry
            If _tt.TryGetValue(ttKey, ttEntry) Then
                If ttEntry.Depth >= depth Then
                    If ttEntry.Flag = TT_EXACT Then Return ttEntry.Score
                    If ttEntry.Flag = TT_BETA Then alpha = Math.Max(alpha, ttEntry.Score)
                    If ttEntry.Flag = TT_ALPHA Then beta = Math.Min(beta, ttEntry.Score)
                    If alpha >= beta Then Return ttEntry.Score
                End If
                ttMove = ttEntry.BestMove
            End If

            If depth <= 0 Then
                Return Quiescence(board, alpha, beta, sideToMove, ply)
            End If

            Dim nextSide = If(sideToMove = CHO, HAN, CHO)

            Dim inCheck = board.IsInCheck(sideToMove)
            Dim extension = If(inCheck, 1, 0)

            ' staticEval 한 번만 계산하여 재사용 (Phase 1C)
            Dim staticEval As Integer = 0
            Dim staticEvalComputed = False
            If Not inCheck AndAlso depth <= 2 AndAlso Math.Abs(beta) < INF - 100 Then
                Dim rawEval = Evaluate(board)
                staticEval = If(sideToMove = CHO, rawEval, -rawEval)
                staticEvalComputed = True
                ' Reverse futility pruning
                If staticEval - FUTILITY_MARGINS(depth) >= beta Then Return beta
            End If

            ' Null Move Pruning (Phase 2A) - 장기는 전술 시퀀스가 길어 보수적으로
            If allowNull AndAlso depth >= 3 AndAlso Not inCheck Then
                Dim R = If(depth >= 8, 3, 2)  ' depth 8+ 에서만 R=3
                board.NullMoveHash()
                Dim nullScore = -Negamax(board, depth - 1 - R, -beta, -beta + 1, nextSide, False, ply + 1)
                board.NullMoveHash()
                If nullScore >= beta Then
                    ' Verification re-search for deep depths (Phase 2A)
                    If depth >= 8 Then
                        Dim vScore = Negamax(board, depth - 1 - R, beta - 1, beta, sideToMove, False, ply)
                        If vScore >= beta Then Return beta
                    Else
                        Return beta
                    End If
                End If
            End If

            ' IID: Internal Iterative Deepening (Phase 2B)
            If Not ttMove.HasValue AndAlso depth >= 5 AndAlso Not inCheck Then
                Negamax(board, depth - 3, alpha, beta, sideToMove, False, ply)
                If _tt.TryGetValue(ttKey, ttEntry) Then
                    ttMove = ttEntry.BestMove
                End If
            End If

            Dim moves = board.GetAllMoves(sideToMove)
            If moves.Count = 0 Then Return -(INF - ply)

            ' Futility pruning 조건 (staticEval 재사용)
            Dim futilityOk = False
            If depth <= 2 AndAlso Not inCheck AndAlso Math.Abs(alpha) < INF - 100 Then
                If Not staticEvalComputed Then
                    Dim rawEval = Evaluate(board)
                    staticEval = If(sideToMove = CHO, rawEval, -rawEval)
                    staticEvalComputed = True
                End If
                If staticEval + FUTILITY_MARGINS(depth) <= alpha Then futilityOk = True
            End If

            Dim grid = board.Grid
            Dim plyKillers As (((Integer, Integer), (Integer, Integer))?, ((Integer, Integer), (Integer, Integer))?) = (Nothing, Nothing)
            _killers.TryGetValue(ply, plyKillers)

            ' Countermove 조회
            Dim cmMove As ((Integer, Integer), (Integer, Integer))? = Nothing
            If board.History.Count > 0 Then
                Dim lastH = board.History(board.History.Count - 1)
                Dim cmKey = (lastH.FromPos.Item1, lastH.FromPos.Item2, lastH.ToPos.Item1, lastH.ToPos.Item2)
                Dim cmVal As ((Integer, Integer), (Integer, Integer))
                If _countermoves.TryGetValue(cmKey, cmVal) Then
                    cmMove = cmVal
                End If
            End If

            moves.Sort(Function(a, b) SortKey(a, ttMove, grid, plyKillers, cmMove, board, nextSide).CompareTo(SortKey(b, ttMove, grid, plyKillers, cmMove, board, nextSide)))

            Dim bestScore = -(INF + 1)
            Dim bestMove As ((Integer, Integer), (Integer, Integer))? = Nothing
            Dim legalCount = 0
            Dim newDepth = depth - 1 + extension

            For Each move In moves
                Dim fromPos = move.Item1
                Dim toPos = move.Item2
                Dim tr = toPos.Item1, tc = toPos.Item2
                Dim isCapture = grid(tr)(tc) <> EMPTY

                ' SEE 계산 (MakeMove 전에 수행)
                Dim seeValue = 0
                Dim isLosingCapture = False
                If isCapture Then
                    seeValue = board.SEE(fromPos.Item1, fromPos.Item2, tr, tc, sideToMove)
                    isLosingCapture = (seeValue < 0)
                End If

                board.MakeMove(fromPos, toPos)

                If board.IsInCheck(sideToMove) Then
                    board.UndoMove()
                    Continue For
                End If

                legalCount += 1

                ' 자살 수 감지: quiet move에서 이동한 기물이 더 싼 적 기물에 잡히는지
                Dim isSuicidal = False
                If Not isCapture Then
                    Dim movingPiece = grid(tr)(tc)
                    Dim myVal = 0
                    PIECE_VALUES.TryGetValue(movingPiece, myVal)
                    If myVal >= 350 Then
                        Dim minAtt = board.GetMinAttackerValue(tr, tc, nextSide)
                        If minAtt > 0 AndAlso minAtt < myVal Then
                            isSuicidal = True
                        End If
                    End If
                End If

                ' 나쁜 캡처 가지치기: depth <= 3이고 손해 캡처이면 스킵
                If isLosingCapture AndAlso depth <= 3 AndAlso Not inCheck AndAlso Not board.IsInCheck(nextSide) Then
                    board.UndoMove()
                    Continue For
                End If

                ' 자살 수 가지치기: depth <= 3이고 자살 수이면 스킵 (체크가 아닐 때)
                If isSuicidal AndAlso depth <= 3 AndAlso Not inCheck AndAlso Not board.IsInCheck(nextSide) Then
                    board.UndoMove()
                    Continue For
                End If

                If futilityOk AndAlso Not isCapture Then
                    If Not board.IsInCheck(nextSide) Then
                        board.UndoMove()
                        Continue For
                    End If
                End If

                ' LMR 개선: 로그 공식 (나쁜 캡처도 LMR 대상)
                Dim doLmr = depth >= 2 AndAlso legalCount >= 3 AndAlso (Not isCapture OrElse isLosingCapture) AndAlso
                            Not (ttMove.HasValue AndAlso move.Item1.Equals(ttMove.Value.Item1) AndAlso move.Item2.Equals(ttMove.Value.Item2)) AndAlso
                            Not inCheck
                If doLmr AndAlso plyKillers.Item1.HasValue Then
                    If (move.Item1.Equals(plyKillers.Item1.Value.Item1) AndAlso move.Item2.Equals(plyKillers.Item1.Value.Item2)) OrElse
                       (plyKillers.Item2.HasValue AndAlso move.Item1.Equals(plyKillers.Item2.Value.Item1) AndAlso move.Item2.Equals(plyKillers.Item2.Value.Item2)) Then
                        doLmr = False
                    End If
                End If

                Dim score As Integer
                If legalCount = 1 Then
                    score = -Negamax(board, newDepth, -beta, -alpha, nextSide, True, ply + 1)
                Else
                    If doLmr Then
                        ' 로그 공식 (장기에 맞게 보수적: / 3.5)
                        Dim R = CInt(Math.Floor(0.5 + Math.Log(depth) * Math.Log(legalCount) / 3.5))
                        If R < 1 Then R = 1
                        ' history 값 높은 수는 reduction 1 감소
                        Dim hKey = (fromPos.Item1, fromPos.Item2, tr, tc)
                        Dim hVal = 0
                        _historyTable.TryGetValue(hKey, hVal)
                        If hVal > HISTORY_MAX \ 2 Then R = Math.Max(1, R - 1)
                        ' 자살 수 / 나쁜 캡처는 추가 감소 +1
                        If isSuicidal OrElse isLosingCapture Then R += 1
                        Dim reducedDepth = Math.Max(1, depth - 1 - R)
                        score = -Negamax(board, reducedDepth, -alpha - 1, -alpha, nextSide, True, ply + 1)
                    Else
                        score = -Negamax(board, newDepth, -alpha - 1, -alpha, nextSide, True, ply + 1)
                    End If

                    If score > alpha AndAlso score < beta Then
                        score = -Negamax(board, newDepth, -beta, -alpha, nextSide, True, ply + 1)
                    End If
                End If

                board.UndoMove()

                If score > bestScore Then
                    bestScore = score
                    bestMove = move
                End If
                If score > alpha Then alpha = score
                If alpha >= beta Then
                    If Not isCapture Then
                        ' Killer 업데이트
                        If Not _killers.ContainsKey(ply) Then
                            _killers(ply) = (move, Nothing)
                        Else
                            Dim k = _killers(ply)
                            If Not (k.Item1.HasValue AndAlso k.Item1.Value.Item1.Equals(move.Item1) AndAlso k.Item1.Value.Item2.Equals(move.Item2)) Then
                                _killers(ply) = (move, k.Item1)
                            End If
                        End If
                        ' History 업데이트 (최대값 제한)
                        Dim hKey2 = (fromPos.Item1, fromPos.Item2, tr, tc)
                        Dim hVal2 = 0
                        _historyTable.TryGetValue(hKey2, hVal2)
                        _historyTable(hKey2) = Math.Min(hVal2 + depth * depth, HISTORY_MAX)
                        ' Countermove 업데이트 (Phase 4A)
                        If board.History.Count > 0 Then
                            Dim lastH2 = board.History(board.History.Count - 1)
                            Dim cmKey2 = (lastH2.FromPos.Item1, lastH2.FromPos.Item2, lastH2.ToPos.Item1, lastH2.ToPos.Item2)
                            _countermoves(cmKey2) = move
                        End If
                    End If
                    Exit For
                End If
            Next

            If legalCount = 0 Then Return -(INF - ply)

            ' TT 저장 (Age 기반 교체 정책, Phase 2F)
            Dim ttFlag As Integer
            If bestScore <= origAlpha Then
                ttFlag = TT_ALPHA
            ElseIf bestScore >= beta Then
                ttFlag = TT_BETA
            Else
                ttFlag = TT_EXACT
            End If

            Dim shouldReplace = True
            Dim existingEntry As TTEntry
            If _tt.TryGetValue(ttKey, existingEntry) Then
                ' 깊이+나이 기반 교체: 새 항목이 더 깊거나, 나이가 다르면 교체
                If existingEntry.Age = _searchAge AndAlso existingEntry.Depth > depth AndAlso existingEntry.Flag = TT_EXACT Then
                    shouldReplace = False
                End If
            End If
            If shouldReplace Then
                _tt(ttKey) = New TTEntry With {
                    .Depth = depth, .Score = bestScore, .Flag = ttFlag,
                    .BestMove = bestMove, .Age = _searchAge
                }
            End If

            Return bestScore
        End Function

        Private Function SortKey(m As ((Integer, Integer), (Integer, Integer)),
                                  ttMove As ((Integer, Integer), (Integer, Integer))?,
                                  grid As String()(),
                                  plyKillers As (((Integer, Integer), (Integer, Integer))?, ((Integer, Integer), (Integer, Integer))?),
                                  cmMove As ((Integer, Integer), (Integer, Integer))?,
                                  Optional board As Board = Nothing,
                                  Optional nextSide As String = Nothing) As Integer
            If ttMove.HasValue AndAlso m.Item1.Equals(ttMove.Value.Item1) AndAlso m.Item2.Equals(ttMove.Value.Item2) Then
                Return -1000000
            End If
            Dim tr = m.Item2.Item1, tc = m.Item2.Item2
            Dim cap = grid(tr)(tc)
            If cap <> EMPTY Then
                ' SEE 기반 캡처 정렬
                If board IsNot Nothing AndAlso nextSide IsNot Nothing Then
                    Dim capturingSide = If(nextSide = CHO, HAN, CHO)
                    Dim seeVal = board.SEE(m.Item1.Item1, m.Item1.Item2, tr, tc, capturingSide)
                    If seeVal >= 0 Then
                        Return -(seeVal + 10000)  ' 좋은 캡처: 앞으로
                    Else
                        Return -seeVal  ' 나쁜 캡처: 양수가 되어 뒤로 밀림
                    End If
                End If
                Dim capVal = 0
                PIECE_VALUES.TryGetValue(cap, capVal)
                Return -(capVal + 10000)
            End If
            If plyKillers.Item1.HasValue Then
                If m.Item1.Equals(plyKillers.Item1.Value.Item1) AndAlso m.Item2.Equals(plyKillers.Item1.Value.Item2) Then Return -5000
                If plyKillers.Item2.HasValue AndAlso m.Item1.Equals(plyKillers.Item2.Value.Item1) AndAlso m.Item2.Equals(plyKillers.Item2.Value.Item2) Then Return -5000
            End If
            ' Countermove heuristic (Phase 4A)
            If cmMove.HasValue AndAlso m.Item1.Equals(cmMove.Value.Item1) AndAlso m.Item2.Equals(cmMove.Value.Item2) Then
                Return -4000
            End If
            Dim hKey = (m.Item1.Item1, m.Item1.Item2, tr, tc)
            Dim hVal = 0
            _historyTable.TryGetValue(hKey, hVal)
            Dim baseScore = -hVal

            ' 자살 수 감지: 이동할 칸이 더 싼 적 기물에 공격받으면 뒤로 밀기
            If board IsNot Nothing AndAlso nextSide IsNot Nothing Then
                Dim movingPiece = grid(m.Item1.Item1)(m.Item1.Item2)
                Dim myVal = 0
                PIECE_VALUES.TryGetValue(movingPiece, myVal)
                If myVal > 0 Then
                    Dim minAttacker = board.GetMinAttackerValue(tr, tc, nextSide)
                    If minAttacker > 0 AndAlso minAttacker < myVal Then
                        ' 자살 수: 큰 패널티로 정렬 후순위
                        baseScore += (myVal - minAttacker)
                    End If
                End If
            End If

            Return baseScore
        End Function

        Public Function FindBestMove(board As Board, side As String,
                                      Optional depth As Integer = 0,
                                      Optional maxTime As Double = 0) As (BestMove As ((Integer, Integer), (Integer, Integer))?, Score As Integer, Depth As Integer)
            If depth = 0 Then depth = DEFAULT_SEARCH_DEPTH
            If maxTime = 0 Then maxTime = MAX_SEARCH_TIME

            _cancelRequested = False
            _deadline = DateTime.Now.AddSeconds(maxTime * 0.9)       ' hard deadline 90%
            _softDeadline = DateTime.Now.AddSeconds(maxTime * 0.4)   ' soft deadline 40%

            _searchAge += 1

            ' TT: 자연 퇴출 (전체 삭제 대신), Phase 2F
            If _tt.Count > TT_MAX_SIZE Then
                Dim toRemove As New List(Of ULong)
                For Each kvp In _tt
                    If kvp.Value.Age < _searchAge - 2 Then toRemove.Add(kvp.Key)
                Next
                For Each k In toRemove
                    _tt.Remove(k)
                Next
                ' 여전히 크면 절반 제거
                If _tt.Count > TT_MAX_SIZE Then
                    toRemove.Clear()
                    Dim cnt = 0
                    For Each kvp In _tt
                        cnt += 1
                        If cnt Mod 2 = 0 Then toRemove.Add(kvp.Key)
                    Next
                    For Each k In toRemove
                        _tt.Remove(k)
                    Next
                End If
            End If

            _killers.Clear()
            ' History aging: 절반으로 감소 (Phase 4B)
            Dim oldHistory As New Dictionary(Of (Integer, Integer, Integer, Integer), Integer)(_historyTable)
            _historyTable.Clear()
            For Each kvp In oldHistory
                Dim halfVal = kvp.Value \ 2
                If halfVal > 0 Then _historyTable(kvp.Key) = halfVal
            Next
            _countermoves.Clear()

            Dim bestMove As ((Integer, Integer), (Integer, Integer))? = Nothing
            Dim bestScore = 0
            Dim searchedDepth = 0
            Dim prevScore = 0
            Dim prevIterScore = 0  ' 이전 반복의 점수 (급락 감지용)
            Dim stableBestMoveCount = 0
            Dim prevBestMove As ((Integer, Integer), (Integer, Integer))? = Nothing

            For d = 1 To depth
                If TimeExpired() Then Exit For

                Dim score As Integer

                ' Aspiration Window 개선 (Phase 2E)
                If d > 1 AndAlso Math.Abs(prevScore) < INF - 100 Then
                    Dim window = 25  ' 초기 25pt
                    Dim a = prevScore - window
                    Dim b = prevScore + window
                    score = Negamax(board, d, a, b, side, True, 0)

                    ' 실패시 점진적 2배 확장 루프
                    While (score <= a OrElse score >= b) AndAlso Not TimeExpired()
                        window *= 2
                        If window > 800 Then
                            score = Negamax(board, d, -INF, INF, side, True, 0)
                            Exit While
                        End If
                        a = prevScore - window
                        b = prevScore + window
                        score = Negamax(board, d, a, b, side, True, 0)
                    End While
                Else
                    score = Negamax(board, d, -INF, INF, side, True, 0)
                End If

                If TimeExpired() AndAlso d > 1 Then Exit For

                prevIterScore = prevScore  ' 급락 비교용: 이전 반복 점수 저장
                prevScore = score

                Dim ttEntry As TTEntry
                If _tt.TryGetValue(board.ZobristHash, ttEntry) AndAlso ttEntry.BestMove.HasValue Then
                    Dim currentBest = ttEntry.BestMove
                    ' Best move 안정성 추적 (Phase 2D)
                    If prevBestMove.HasValue AndAlso
                       currentBest.Value.Item1.Equals(prevBestMove.Value.Item1) AndAlso
                       currentBest.Value.Item2.Equals(prevBestMove.Value.Item2) Then
                        stableBestMoveCount += 1
                    Else
                        stableBestMoveCount = 0
                    End If
                    prevBestMove = currentBest
                    bestMove = currentBest
                    bestScore = If(side = CHO, score, -score)
                    searchedDepth = d
                End If

                If Math.Abs(score) > INF - 100 Then Exit For

                ' 시간 관리 (Phase 2D): soft deadline에서 조기 종료
                If d >= 4 AndAlso SoftTimeExpired() Then
                    ' 점수 급락시 soft 연장 (70%) - prevIterScore와 비교
                    If d > 1 AndAlso score < prevIterScore - 50 Then
                        _softDeadline = DateTime.Now.AddSeconds(maxTime * 0.3)  ' 추가 30%
                    ElseIf stableBestMoveCount >= 3 Then
                        ' best move 3회 이상 안정이면 조기 종료
                        Exit For
                    Else
                        Exit For
                    End If
                End If
            Next

            Return (bestMove, bestScore, searchedDepth)
        End Function

    End Module

End Namespace
