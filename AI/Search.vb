' NegaMax + Alpha-Beta 탐색 엔진
' 최적화: TT, Killer Move, History Heuristic, Null Move Pruning,
'         LMR, Aspiration Windows, Check Extension, Futility/Reverse Futility, PVS
Imports MacroAutoControl.Constants
Imports MacroAutoControl.Engine
Imports MacroAutoControl.AI.Evaluator

Namespace MacroAutoControl.AI

    ' TT 엔트리
    Public Structure TTEntry
        Public Depth As Integer
        Public Score As Integer
        Public Flag As Integer
        Public BestMove As ((Integer, Integer), (Integer, Integer))?
    End Structure

    Public Module Search

        Private Const INF As Integer = 999999
        Private Const REPETITION_PENALTY As Integer = 900000
        Private Const TT_EXACT As Integer = 0
        Private Const TT_ALPHA As Integer = 1
        Private Const TT_BETA As Integer = 2
        Private Const TT_MAX_SIZE As Integer = 500000
        Private Const QS_MAX_DEPTH As Integer = 6
        Private Const DELTA_MARGIN As Integer = 200
        Private ReadOnly FUTILITY_MARGINS As Integer() = {0, 200, 450}

        ' 모듈 레벨 탐색 테이블
        Private _tt As New Dictionary(Of ULong, TTEntry)
        Private _killers As New Dictionary(Of Integer, (((Integer, Integer), (Integer, Integer))?, ((Integer, Integer), (Integer, Integer))?))
        Private _historyTable As New Dictionary(Of (Integer, Integer, Integer, Integer), Integer)

        Private _deadline As DateTime
        Private _cancelRequested As Boolean = False

        ' 동일구간 반복 방지: 게임 수준 AI 수 히스토리
        Private _gameMoves As List(Of ((Integer, Integer), (Integer, Integer)))
        Private Const MOVE_REPEAT_PENALTY As Integer = 500000

        ''' <summary>외부에서 탐색 중단 요청</summary>
        Public Sub CancelSearch()
            _cancelRequested = True
        End Sub

        Private Function TimeExpired() As Boolean
            Return _cancelRequested OrElse DateTime.Now > _deadline
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

            ' Stand pat
            Dim rawEval = Evaluate(board)
            Dim standPat = If(sideToMove = CHO, rawEval, -rawEval)

            If standPat >= beta Then Return beta
            If standPat > alpha Then alpha = standPat
            If qsDepth >= QS_MAX_DEPTH Then Return standPat

            ' 잡는 수만 생성
            Dim moves = board.GetAllMoves(sideToMove)
            Dim grid = board.Grid
            Dim captures As New List(Of (Integer, ((Integer, Integer), (Integer, Integer))))

            For Each m In moves
                Dim tr = m.Item2.Item1, tc = m.Item2.Item2
                Dim target = grid(tr)(tc)
                If target <> EMPTY Then
                    Dim victimVal = 0
                    PIECE_VALUES.TryGetValue(target, victimVal)
                    Dim attackerVal = 0
                    PIECE_VALUES.TryGetValue(grid(m.Item1.Item1)(m.Item1.Item2), attackerVal)
                    ' Delta pruning
                    If standPat + victimVal + DELTA_MARGIN < alpha Then Continue For
                    captures.Add((-(victimVal * 10 - attackerVal), m))
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

            ' TT 조회
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

            ' 깊이 0 → 정지 탐색
            If depth <= 0 Then
                Return Quiescence(board, alpha, beta, sideToMove, ply)
            End If

            Dim nextSide = If(sideToMove = CHO, HAN, CHO)

            ' 장군 상태 캐싱
            Dim inCheck = board.IsInCheck(sideToMove)
            Dim extension = If(inCheck, 1, 0)

            ' Reverse Futility Pruning
            If Not inCheck AndAlso depth <= 2 AndAlso Math.Abs(beta) < INF - 100 Then
                Dim rawEval = Evaluate(board)
                Dim staticEval = If(sideToMove = CHO, rawEval, -rawEval)
                If staticEval - FUTILITY_MARGINS(depth) >= beta Then Return beta
            End If

            ' Null Move Pruning
            If allowNull AndAlso depth >= 3 AndAlso Not inCheck Then
                Dim R = 2
                board.NullMoveHash()
                Dim nullScore = -Negamax(board, depth - 1 - R, -beta, -beta + 1, nextSide, False, ply + 1)
                board.NullMoveHash()
                If nullScore >= beta Then Return beta
            End If

            ' 수 생성
            Dim moves = board.GetAllMoves(sideToMove)
            If moves.Count = 0 Then Return -(INF - ply)

            ' Futility Pruning 준비
            Dim futilityOk = False
            If depth <= 2 AndAlso Not inCheck AndAlso Math.Abs(alpha) < INF - 100 Then
                Dim rawEval = Evaluate(board)
                Dim staticEval = If(sideToMove = CHO, rawEval, -rawEval)
                If staticEval + FUTILITY_MARGINS(depth) <= alpha Then futilityOk = True
            End If

            ' 수 정렬
            Dim grid = board.Grid
            Dim plyKillers As (((Integer, Integer), (Integer, Integer))?, ((Integer, Integer), (Integer, Integer))?) = (Nothing, Nothing)
            _killers.TryGetValue(ply, plyKillers)

            moves.Sort(Function(a, b) SortKey(a, ttMove, grid, plyKillers).CompareTo(SortKey(b, ttMove, grid, plyKillers)))

            Dim bestScore = -(INF + 1)
            Dim bestMove As ((Integer, Integer), (Integer, Integer))? = Nothing
            Dim bestIsRepetition = False
            Dim legalCount = 0
            Dim newDepth = depth - 1 + extension

            For Each move In moves
                Dim fromPos = move.Item1
                Dim toPos = move.Item2
                Dim tr = toPos.Item1, tc = toPos.Item2
                Dim isCapture = grid(tr)(tc) <> EMPTY

                board.MakeMove(fromPos, toPos)

                If board.IsInCheck(sideToMove) Then
                    board.UndoMove()
                    Continue For
                End If

                legalCount += 1

                ' 루트 레벨 동일구간 반복 감지: 같은 수를 다시 두면 페널티
                If ply = 0 AndAlso _gameMoves IsNot Nothing AndAlso _gameMoves.Count > 0 Then
                    Dim moveRepeatPenalty = GetMoveRepeatPenalty(fromPos, toPos)
                    If moveRepeatPenalty > 0 Then
                        Dim repScore = -(moveRepeatPenalty)
                        board.UndoMove()
                        If repScore > bestScore Then
                            bestScore = repScore
                            bestMove = move
                            bestIsRepetition = True
                        End If
                        Continue For
                    End If
                End If

                ' 반복 수 감지: 히스토리에 1회 이상 등장 = 2번째 동일 국면 = 반복 금지
                If board.IsRepetition(1) Then
                    Dim repScore = -(REPETITION_PENALTY - ply)
                    board.UndoMove()
                    If repScore > bestScore Then
                        bestScore = repScore
                        bestMove = move
                        bestIsRepetition = True
                    End If
                    Continue For
                End If

                ' Futility Pruning
                If futilityOk AndAlso Not isCapture Then
                    If Not board.IsInCheck(nextSide) Then
                        board.UndoMove()
                        Continue For
                    End If
                End If

                ' LMR 판단
                Dim doLmr = depth >= 3 AndAlso legalCount >= 4 AndAlso Not isCapture AndAlso
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
                        Dim R = If(legalCount >= 8, 2, 1)
                        score = -Negamax(board, depth - 1 - R, -alpha - 1, -alpha, nextSide, True, ply + 1)
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
                    bestIsRepetition = False
                End If
                If score > alpha Then alpha = score
                If alpha >= beta Then
                    If Not isCapture Then
                        ' Killer move 저장
                        If Not _killers.ContainsKey(ply) Then
                            _killers(ply) = (move, Nothing)
                        Else
                            Dim k = _killers(ply)
                            If Not (k.Item1.HasValue AndAlso k.Item1.Value.Item1.Equals(move.Item1) AndAlso k.Item1.Value.Item2.Equals(move.Item2)) Then
                                _killers(ply) = (move, k.Item1)
                            End If
                        End If
                        ' History heuristic
                        Dim hKey = (fromPos.Item1, fromPos.Item2, tr, tc)
                        Dim hVal = 0
                        _historyTable.TryGetValue(hKey, hVal)
                        _historyTable(hKey) = hVal + depth * depth
                    End If
                    Exit For
                End If
            Next

            If legalCount = 0 Then Return -(INF - ply)

            ' TT 저장 (반복 점수는 경로 의존적이므로 TT에 저장하지 않음)
            If Not bestIsRepetition Then
                Dim ttFlag As Integer
                If bestScore <= origAlpha Then
                    ttFlag = TT_ALPHA
                ElseIf bestScore >= beta Then
                    ttFlag = TT_BETA
                Else
                    ttFlag = TT_EXACT
                End If
                _tt(ttKey) = New TTEntry With {
                    .Depth = depth, .Score = bestScore, .Flag = ttFlag, .BestMove = bestMove
                }
            End If

            Return bestScore
        End Function

        ''' <summary>
        ''' 동일구간 반복 페널티 계산.
        ''' - 마지막 수의 역수(왕복): MOVE_REPEAT_PENALTY (최대)
        ''' - 이전에 같은 수를 둔 적 있음: MOVE_REPEAT_PENALTY / 2
        ''' - 반복 아님: 0
        ''' </summary>
        Private Function GetMoveRepeatPenalty(fromPos As (Integer, Integer), toPos As (Integer, Integer)) As Integer
            If _gameMoves Is Nothing OrElse _gameMoves.Count = 0 Then Return 0

            ' 마지막 AI 수의 역수인지 (왕복 패턴: A→B 다음에 B→A)
            Dim lastMove = _gameMoves(_gameMoves.Count - 1)
            If lastMove.Item1.Equals(toPos) AndAlso lastMove.Item2.Equals(fromPos) Then
                Return MOVE_REPEAT_PENALTY
            End If

            ' 같은 수를 이전에 둔 적 있는지
            For i = _gameMoves.Count - 1 To 0 Step -1
                Dim gm = _gameMoves(i)
                If gm.Item1.Equals(fromPos) AndAlso gm.Item2.Equals(toPos) Then
                    Return MOVE_REPEAT_PENALTY \ 2
                End If
            Next

            Return 0
        End Function

        Private Function SortKey(m As ((Integer, Integer), (Integer, Integer)),
                                  ttMove As ((Integer, Integer), (Integer, Integer))?,
                                  grid As String()(),
                                  plyKillers As (((Integer, Integer), (Integer, Integer))?, ((Integer, Integer), (Integer, Integer))?)) As Integer
            If ttMove.HasValue AndAlso m.Item1.Equals(ttMove.Value.Item1) AndAlso m.Item2.Equals(ttMove.Value.Item2) Then
                Return -1000000
            End If
            Dim tr = m.Item2.Item1, tc = m.Item2.Item2
            Dim cap = grid(tr)(tc)
            If cap <> EMPTY Then
                Dim capVal = 0
                PIECE_VALUES.TryGetValue(cap, capVal)
                Return -(capVal + 10000)
            End If
            If plyKillers.Item1.HasValue Then
                If m.Item1.Equals(plyKillers.Item1.Value.Item1) AndAlso m.Item2.Equals(plyKillers.Item1.Value.Item2) Then Return -5000
                If plyKillers.Item2.HasValue AndAlso m.Item1.Equals(plyKillers.Item2.Value.Item1) AndAlso m.Item2.Equals(plyKillers.Item2.Value.Item2) Then Return -5000
            End If
            Dim hKey = (m.Item1.Item1, m.Item1.Item2, tr, tc)
            Dim hVal = 0
            _historyTable.TryGetValue(hKey, hVal)
            Return -hVal
        End Function

        Public Function FindBestMove(board As Board, side As String,
                                      Optional depth As Integer = 0,
                                      Optional maxTime As Double = 0,
                                      Optional gameMoves As List(Of ((Integer, Integer), (Integer, Integer))) = Nothing) As (BestMove As ((Integer, Integer), (Integer, Integer))?, Score As Integer, Depth As Integer)
            If depth = 0 Then depth = DEFAULT_SEARCH_DEPTH
            If maxTime = 0 Then maxTime = MAX_SEARCH_TIME

            _cancelRequested = False
            _deadline = DateTime.Now.AddSeconds(maxTime)

            ' 동일구간 반복 방지용 게임 수 히스토리 설정
            _gameMoves = gameMoves

            ' TT 크기 초과 시 정리
            If _tt.Count > TT_MAX_SIZE Then _tt.Clear()

            ' killers와 history 초기화
            _killers.Clear()
            _historyTable.Clear()

            Dim bestMove As ((Integer, Integer), (Integer, Integer))? = Nothing
            Dim bestScore = 0
            Dim searchedDepth = 0
            Dim prevScore = 0

            ' Iterative Deepening
            For d = 1 To depth
                If TimeExpired() Then Exit For

                Dim score As Integer

                ' Aspiration Windows
                If d > 1 AndAlso Math.Abs(prevScore) < INF - 100 Then
                    Dim window = 50
                    Dim a = prevScore - window
                    Dim b = prevScore + window
                    score = Negamax(board, d, a, b, side, True, 0)

                    If score <= a OrElse score >= b Then
                        Dim window2 = window * 4
                        Dim a2 = prevScore - window2
                        Dim b2 = prevScore + window2
                        score = Negamax(board, d, a2, b2, side, True, 0)

                        If score <= a2 OrElse score >= b2 Then
                            score = Negamax(board, d, -INF, INF, side, True, 0)
                        End If
                    End If
                Else
                    score = Negamax(board, d, -INF, INF, side, True, 0)
                End If

                If TimeExpired() AndAlso d > 1 Then Exit For

                prevScore = score

                ' TT에서 루트 최선수 추출
                Dim ttEntry As TTEntry
                If _tt.TryGetValue(board.ZobristHash, ttEntry) AndAlso ttEntry.BestMove.HasValue Then
                    bestMove = ttEntry.BestMove
                    bestScore = If(side = CHO, score, -score)
                    searchedDepth = d
                End If

                ' 외통수 발견시 즉시 반환
                If Math.Abs(score) > INF - 100 Then Exit For
            Next

            Return (bestMove, bestScore, searchedDepth)
        End Function

    End Module

End Namespace
