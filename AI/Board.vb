' 보드 표현 및 수 실행/되돌리기
Imports MacroAutoControl.Constants
Imports MacroAutoControl.Engine.Pieces

Namespace MacroAutoControl.Engine

    ' 역방향 공격 탐지 오프셋
    Public Structure HorseOffset
        Public HorseDr As Integer
        Public HorseDc As Integer
        Public WaypointDr As Integer
        Public WaypointDc As Integer
        Public Sub New(hdr As Integer, hdc As Integer, wdr As Integer, wdc As Integer)
            HorseDr = hdr : HorseDc = hdc : WaypointDr = wdr : WaypointDc = wdc
        End Sub
    End Structure

    Public Structure ElephantOffset
        Public Edr As Integer
        Public Edc As Integer
        Public W1dr As Integer
        Public W1dc As Integer
        Public W2dr As Integer
        Public W2dc As Integer
        Public Sub New(edr As Integer, edc As Integer, w1dr As Integer, w1dc As Integer, w2dr As Integer, w2dc As Integer)
            Me.Edr = edr : Me.Edc = edc : Me.W1dr = w1dr : Me.W1dc = w1dc : Me.W2dr = w2dr : Me.W2dc = w2dc
        End Sub
    End Structure

    ' 히스토리 항목
    Public Structure MoveHistory
        Public FromPos As (Integer, Integer)
        Public ToPos As (Integer, Integer)
        Public Captured As String
        Public MovingPiece As String
    End Structure

    Public Class Board

        Public Grid As String()()
        Private _history As New List(Of MoveHistory)
        Public ReadOnly Property History As IReadOnlyList(Of MoveHistory)
            Get
                Return _history
            End Get
        End Property
        Private _kingPos As New Dictionary(Of String, (Integer, Integer)?)
        Public ZobristHash As ULong

        ' 기물 리스트: piece code → List of (row, col)
        Private _pieceLists As New Dictionary(Of String, List(Of (Integer, Integer)))

        ' 반복 수 감지용 해시 히스토리
        Private _hashHistory As New List(Of ULong)

        ' 진영별 전진 방향 (왕 위치 기반 동적 결정)
        Public ChoForward As Integer = -1  ' 초: 기본 위로
        Public HanForward As Integer = 1   ' 한: 기본 아래로

        ' Zobrist 테이블 (공유)
        Private Shared ReadOnly _allPieces As String() = {
            "CK", "CS", "CC", "CM", "CE", "CP", "CJ",
            "HK", "HS", "HC", "HM", "HE", "HP", "HB"
        }
        Private Shared ReadOnly _pieceIndex As New Dictionary(Of String, Integer)
        Private Shared ReadOnly _zobristTable As ULong(,)
        Private Shared ReadOnly _zobristSide As ULong

        ' 역방향 공격 오프셋 테이블
        Private Shared ReadOnly _horseAttackOffsets As HorseOffset() = {
            New HorseOffset(2, 1, 1, 1), New HorseOffset(2, -1, 1, -1),
            New HorseOffset(-2, 1, -1, 1), New HorseOffset(-2, -1, -1, -1),
            New HorseOffset(1, 2, 1, 1), New HorseOffset(-1, 2, -1, 1),
            New HorseOffset(1, -2, 1, -1), New HorseOffset(-1, -2, -1, -1)
        }
        Private Shared ReadOnly _elephantAttackOffsets As ElephantOffset() = {
            New ElephantOffset(3, 2, 2, 2, 1, 1), New ElephantOffset(3, -2, 2, -2, 1, -1),
            New ElephantOffset(-3, 2, -2, 2, -1, 1), New ElephantOffset(-3, -2, -2, -2, -1, -1),
            New ElephantOffset(2, 3, 2, 2, 1, 1), New ElephantOffset(-2, 3, -2, 2, -1, 1),
            New ElephantOffset(2, -3, 2, -2, 1, -1), New ElephantOffset(-2, -3, -2, -2, -1, -1)
        }

        Shared Sub New()
            For i = 0 To _allPieces.Length - 1
                _pieceIndex(_allPieces(i)) = i
            Next
            Dim rng As New Random(42)
            _zobristTable = New ULong(13, 89) {}
            For i = 0 To 13
                For j = 0 To 89
                    Dim bytes(7) As Byte
                    rng.NextBytes(bytes)
                    _zobristTable(i, j) = BitConverter.ToUInt64(bytes, 0)
                Next
            Next
            Dim sideBytes(7) As Byte
            rng.NextBytes(sideBytes)
            _zobristSide = BitConverter.ToUInt64(sideBytes, 0)
        End Sub

        Public Sub New(Optional grid As String()() = Nothing)
            If grid IsNot Nothing Then
                Me.Grid = New String(BOARD_ROWS - 1)() {}
                For r = 0 To BOARD_ROWS - 1
                    Me.Grid(r) = New String(BOARD_COLS - 1) {}
                    Array.Copy(grid(r), Me.Grid(r), BOARD_COLS)
                Next
            Else
                Me.Grid = GetInitialBoard()
            End If

            _history = New List(Of MoveHistory)
            _kingPos(CHO) = Nothing
            _kingPos(HAN) = Nothing
            ZobristHash = 0

            ' 기물 리스트 초기화
            For Each p In _allPieces
                _pieceLists(p) = New List(Of (Integer, Integer))
            Next
            _hashHistory = New List(Of ULong)

            For r = 0 To BOARD_ROWS - 1
                For c = 0 To BOARD_COLS - 1
                    Dim piece = Me.Grid(r)(c)
                    If piece <> EMPTY Then
                        Dim idx As Integer
                        If _pieceIndex.TryGetValue(piece, idx) Then
                            ZobristHash = ZobristHash Xor _zobristTable(idx, r * 9 + c)
                        End If
                        If piece = CK Then _kingPos(CHO) = (r, c)
                        If piece = HK Then _kingPos(HAN) = (r, c)
                        Dim lst As List(Of (Integer, Integer)) = Nothing
                        If _pieceLists.TryGetValue(piece, lst) Then lst.Add((r, c))
                    End If
                Next
            Next

            ' 왕 위치 기반 전진 방향 결정
            If _kingPos(CHO).HasValue Then
                ChoForward = If(_kingPos(CHO).Value.Item1 >= 5, -1, 1)
            End If
            If _kingPos(HAN).HasValue Then
                HanForward = If(_kingPos(HAN).Value.Item1 >= 5, -1, 1)
            End If
        End Sub

        Public Function Copy() As Board
            Dim newBoard As New Board(Me.Grid)
            newBoard._history = New List(Of MoveHistory)(_history)
            newBoard._kingPos = New Dictionary(Of String, (Integer, Integer)?)(_kingPos)
            newBoard.ZobristHash = ZobristHash
            newBoard._hashHistory = New List(Of ULong)(_hashHistory)
            For Each kvp In _pieceLists
                newBoard._pieceLists(kvp.Key) = New List(Of (Integer, Integer))(kvp.Value)
            Next
            Return newBoard
        End Function

        Public Function [Get](r As Integer, c As Integer) As String
            Return Grid(r)(c)
        End Function

        Public Sub [Set](r As Integer, c As Integer, piece As String)
            Dim old = Grid(r)(c)
            If old <> EMPTY Then
                Dim idx As Integer
                If _pieceIndex.TryGetValue(old, idx) Then
                    ZobristHash = ZobristHash Xor _zobristTable(idx, r * 9 + c)
                End If
                If old = CK Then _kingPos(CHO) = Nothing
                If old = HK Then _kingPos(HAN) = Nothing
            End If
            Grid(r)(c) = piece
            If piece <> EMPTY Then
                Dim idx As Integer
                If _pieceIndex.TryGetValue(piece, idx) Then
                    ZobristHash = ZobristHash Xor _zobristTable(idx, r * 9 + c)
                End If
                If piece = CK Then _kingPos(CHO) = (r, c)
                If piece = HK Then _kingPos(HAN) = (r, c)
            End If
        End Sub

        Public Function MakeMove(fromPos As (Integer, Integer), toPos As (Integer, Integer)) As String
            Dim fr = fromPos.Item1, fc = fromPos.Item2
            Dim tr = toPos.Item1, tc = toPos.Item2
            Dim movingPiece = Grid(fr)(fc)
            Dim captured = Grid(tr)(tc)

            ' 해시 히스토리 push
            _hashHistory.Add(ZobristHash)

            _history.Add(New MoveHistory With {
                .FromPos = fromPos, .ToPos = toPos,
                .Captured = captured, .MovingPiece = movingPiece
            })

            Dim pi As Integer
            If _pieceIndex.TryGetValue(movingPiece, pi) Then
                ZobristHash = ZobristHash Xor _zobristTable(pi, fr * 9 + fc)
                ZobristHash = ZobristHash Xor _zobristTable(pi, tr * 9 + tc)
            End If
            If captured <> EMPTY Then
                Dim ci As Integer
                If _pieceIndex.TryGetValue(captured, ci) Then
                    ZobristHash = ZobristHash Xor _zobristTable(ci, tr * 9 + tc)
                End If
            End If
            ZobristHash = ZobristHash Xor _zobristSide

            Grid(tr)(tc) = movingPiece
            Grid(fr)(fc) = EMPTY

            If movingPiece = CK Then _kingPos(CHO) = (tr, tc)
            If movingPiece = HK Then _kingPos(HAN) = (tr, tc)
            If captured = CK Then _kingPos(CHO) = Nothing
            If captured = HK Then _kingPos(HAN) = Nothing

            ' 기물 리스트 점진 업데이트
            Dim moveLst As List(Of (Integer, Integer)) = Nothing
            If _pieceLists.TryGetValue(movingPiece, moveLst) Then
                Dim idx = moveLst.IndexOf(fromPos)
                If idx >= 0 Then moveLst(idx) = toPos
            End If
            If captured <> EMPTY Then
                Dim capLst As List(Of (Integer, Integer)) = Nothing
                If _pieceLists.TryGetValue(captured, capLst) Then
                    capLst.Remove(toPos)
                End If
            End If

            Return captured
        End Function

        Public Function UndoMove() As (FromPos As (Integer, Integer), ToPos As (Integer, Integer), Captured As String)
            If _history.Count = 0 Then Throw New InvalidOperationException("되돌릴 수가 없습니다")

            Dim last = _history(_history.Count - 1)
            _history.RemoveAt(_history.Count - 1)
            Dim fr = last.FromPos.Item1, fc = last.FromPos.Item2
            Dim tr = last.ToPos.Item1, tc = last.ToPos.Item2

            ' 해시 히스토리 pop → 복원
            If _hashHistory.Count > 0 Then
                ZobristHash = _hashHistory(_hashHistory.Count - 1)
                _hashHistory.RemoveAt(_hashHistory.Count - 1)
            Else
                ZobristHash = ZobristHash Xor _zobristSide
                If last.Captured <> EMPTY Then
                    Dim ci As Integer
                    If _pieceIndex.TryGetValue(last.Captured, ci) Then
                        ZobristHash = ZobristHash Xor _zobristTable(ci, tr * 9 + tc)
                    End If
                End If
                Dim pi As Integer
                If _pieceIndex.TryGetValue(last.MovingPiece, pi) Then
                    ZobristHash = ZobristHash Xor _zobristTable(pi, tr * 9 + tc)
                    ZobristHash = ZobristHash Xor _zobristTable(pi, fr * 9 + fc)
                End If
            End If

            Grid(fr)(fc) = last.MovingPiece
            Grid(tr)(tc) = last.Captured

            If last.MovingPiece = CK Then _kingPos(CHO) = (fr, fc)
            If last.MovingPiece = HK Then _kingPos(HAN) = (fr, fc)
            If last.Captured = CK Then _kingPos(CHO) = (tr, tc)
            If last.Captured = HK Then _kingPos(HAN) = (tr, tc)

            ' 기물 리스트 복원
            Dim moveLst As List(Of (Integer, Integer)) = Nothing
            If _pieceLists.TryGetValue(last.MovingPiece, moveLst) Then
                Dim idx = moveLst.IndexOf(last.ToPos)
                If idx >= 0 Then moveLst(idx) = last.FromPos
            End If
            If last.Captured <> EMPTY Then
                Dim capLst As List(Of (Integer, Integer)) = Nothing
                If _pieceLists.TryGetValue(last.Captured, capLst) Then
                    capLst.Add(last.ToPos)
                End If
            End If

            Return (last.FromPos, last.ToPos, last.Captured)
        End Function

        Public Sub NullMoveHash()
            ZobristHash = ZobristHash Xor _zobristSide
        End Sub

        ''' <summary>기물 리스트 접근 (Evaluator용)</summary>
        Public Function GetPieceList(piece As String) As List(Of (Integer, Integer))
            Dim lst As List(Of (Integer, Integer)) = Nothing
            If _pieceLists.TryGetValue(piece, lst) Then Return lst
            Return New List(Of (Integer, Integer))
        End Function

        ''' <summary>반복 수 감지: 현재 해시가 history에 threshold 이상 등장하면 True</summary>
        Public Function IsRepetition(Optional threshold As Integer = 2) As Boolean
            Dim h = ZobristHash
            Dim count = 0
            For i = _hashHistory.Count - 1 To 0 Step -1
                If _hashHistory(i) = h Then
                    count += 1
                    If count >= threshold Then Return True
                End If
            Next
            Return False
        End Function

        ''' <summary>게임 수준 해시 히스토리를 보드에 주입 (MacroRunner에서 호출)</summary>
        Public Sub InjectGameHistory(gameHashes As IEnumerable(Of ULong))
            _hashHistory = New List(Of ULong)(gameHashes)
        End Sub

        ''' <summary>
        ''' (r,c) 칸을 attackerSide가 공격하는 기물 중 가장 싼 기물 가치를 반환.
        ''' 공격자가 없으면 0. 궁=1(명목), 졸=200, 사=300, 상=350, 마=450, 포=800, 차=1300
        ''' </summary>
        Public Function GetMinAttackerValue(r As Integer, c As Integer, attackerSide As String) As Integer
            Dim minVal = Integer.MaxValue

            ' 공격측 기물 코드 결정
            Dim aKing, aCha, aPo, aMa, aSang, aSol, aSa As String
            Dim solFwd As Integer
            If attackerSide = CHO Then
                aKing = CK : aCha = CC : aPo = CP : aMa = CM : aSang = CE : aSol = CJ : aSa = CS
                solFwd = ChoForward
            Else
                aKing = HK : aCha = HC : aPo = HP : aMa = HM : aSang = HE : aSol = HB : aSa = HS
                solFwd = HanForward
            End If

            ' 0) 궁(왕) 공격 (명목 가치 1) - 궁성 내 인접 칸
            Dim kingPalace = If(attackerSide = CHO, CHO_PALACE, HAN_PALACE)
            Dim kingEnemyPalace = If(attackerSide = CHO, HAN_PALACE, CHO_PALACE)
            ' 직선 인접
            For Each d In {(-1, 0), (1, 0), (0, -1), (0, 1)}
                Dim kr = r + d.Item1, kc = c + d.Item2
                If kr >= 0 AndAlso kr < 10 AndAlso kc >= 0 AndAlso kc < 9 Then
                    If Grid(kr)(kc) = aKing Then
                        ' 왕이 (kr,kc)에서 (r,c)로 이동 가능한지: 둘 다 같은 궁성 내
                        If (kingPalace.Contains((kr, kc)) AndAlso kingPalace.Contains((r, c))) OrElse
                           (kingEnemyPalace.Contains((kr, kc)) AndAlso kingEnemyPalace.Contains((r, c))) Then
                            minVal = 1 : GoTo Done
                        End If
                    End If
                End If
            Next
            ' 궁성 대각선
            Dim kingDiag As List(Of (Integer, Integer)) = Nothing
            If PALACE_DIAG_MOVES.TryGetValue((r, c), kingDiag) Then
                For Each t In kingDiag
                    If Grid(t.Item1)(t.Item2) = aKing Then
                        If (kingPalace.Contains(t) AndAlso kingPalace.Contains((r, c))) OrElse
                           (kingEnemyPalace.Contains(t) AndAlso kingEnemyPalace.Contains((r, c))) Then
                            minVal = 1 : GoTo Done
                        End If
                    End If
                Next
            End If

            ' 1) 졸/병 공격 (가치 200) - 졸이 (r,c)를 공격하려면 졸이 특정 위치에 있어야 함
            ' 졸은 전진+좌우로 이동하므로, (r,c)를 공격하는 졸은:
            '   - (r - solFwd, c): 전진해서 (r,c)에 도달
            '   - (r, c-1), (r, c+1): 좌우에서 (r,c)에 도달
            Dim sr = r - solFwd
            If sr >= 0 AndAlso sr < 10 AndAlso Grid(sr)(c) = aSol Then minVal = Math.Min(minVal, 200)
            If minVal = 200 Then GoTo Done
            If c > 0 AndAlso Grid(r)(c - 1) = aSol Then minVal = Math.Min(minVal, 200)
            If minVal = 200 Then GoTo Done
            If c < 8 AndAlso Grid(r)(c + 1) = aSol Then minVal = Math.Min(minVal, 200)
            If minVal = 200 Then GoTo Done

            ' 2) 사 공격 (가치 300) - 궁성 대각선 이동
            Dim diagTargets As List(Of (Integer, Integer)) = Nothing
            If PALACE_DIAG_MOVES.TryGetValue((r, c), diagTargets) Then
                For Each t In diagTargets
                    If Grid(t.Item1)(t.Item2) = aSa Then minVal = Math.Min(minVal, 300)
                Next
            End If
            ' 사의 직선 이동 (궁성 내)
            Dim palace = If(attackerSide = CHO, CHO_PALACE, HAN_PALACE)
            Dim enemyPalace = If(attackerSide = CHO, HAN_PALACE, CHO_PALACE)
            If palace.Contains((r, c)) OrElse enemyPalace.Contains((r, c)) Then
                For Each d In {(-1, 0), (1, 0), (0, -1), (0, 1)}
                    Dim nr2 = r + d.Item1, nc2 = c + d.Item2
                    If nr2 >= 0 AndAlso nr2 < 10 AndAlso nc2 >= 0 AndAlso nc2 < 9 Then
                        Dim targetPalace = If(palace.Contains((r, c)), palace, enemyPalace)
                        If targetPalace.Contains((nr2, nc2)) AndAlso Grid(nr2)(nc2) = aSa Then
                            minVal = Math.Min(minVal, 300)
                        End If
                    End If
                Next
            End If

            ' 3) 상 공격 (가치 350)
            For Each e In _elephantAttackOffsets
                Dim er = r + e.Edr, ec = c + e.Edc
                If er >= 0 AndAlso er < 10 AndAlso ec >= 0 AndAlso ec < 9 AndAlso Grid(er)(ec) = aSang Then
                    If Grid(r + e.W1dr)(c + e.W1dc) = EMPTY Then
                        If Grid(r + e.W2dr)(c + e.W2dc) = EMPTY Then
                            minVal = Math.Min(minVal, 350)
                        End If
                    End If
                End If
            Next

            ' 4) 마 공격 (가치 450)
            For Each h In _horseAttackOffsets
                Dim hr2 = r + h.HorseDr, hc2 = c + h.HorseDc
                If hr2 >= 0 AndAlso hr2 < 10 AndAlso hc2 >= 0 AndAlso hc2 < 9 AndAlso Grid(hr2)(hc2) = aMa Then
                    If Grid(r + h.WaypointDr)(c + h.WaypointDc) = EMPTY Then
                        minVal = Math.Min(minVal, 450)
                    End If
                End If
            Next

            ' 5) 포 공격 (가치 800) - 직선 방향, 포대 1개 넘어서
            Dim lineDirs = {(-1, 0), (1, 0), (0, -1), (0, 1)}
            For Each d In lineDirs
                Dim nr = r + d.Item1, nc = c + d.Item2
                Dim screen = False
                While nr >= 0 AndAlso nr < 10 AndAlso nc >= 0 AndAlso nc < 9
                    Dim p = Grid(nr)(nc)
                    If Not screen Then
                        If p <> EMPTY Then
                            If p = CP OrElse p = HP Then Exit While  ' 포는 포대가 될 수 없음
                            screen = True
                        End If
                    Else
                        If p <> EMPTY Then
                            If p = aPo Then minVal = Math.Min(minVal, 800)
                            Exit While
                        End If
                    End If
                    nr += d.Item1 : nc += d.Item2
                End While
            Next

            ' 6) 차 공격 (가치 1300) - 직선 방향
            For Each d In lineDirs
                Dim nr = r + d.Item1, nc = c + d.Item2
                While nr >= 0 AndAlso nr < 10 AndAlso nc >= 0 AndAlso nc < 9
                    Dim p = Grid(nr)(nc)
                    If p <> EMPTY Then
                        If p = aCha Then minVal = Math.Min(minVal, 1300)
                        Exit While
                    End If
                    nr += d.Item1 : nc += d.Item2
                End While
            Next

Done:
            If minVal = Integer.MaxValue Then Return 0
            Return minVal
        End Function

        ''' <summary>
        ''' Static Exchange Evaluation: (fromR,fromC) → (toR,toC) 캡처 교환의 최종 가치.
        ''' 양수=이득, 음수=손해. side = 캡처를 시작하는 측.
        ''' </summary>
        Public Function SEE(fromR As Integer, fromC As Integer, toR As Integer, toC As Integer, side As String) As Integer
            ' 피해자 가치
            Dim target = Grid(toR)(toC)
            If target = EMPTY Then Return 0
            Dim victimVal = 0
            PIECE_VALUES.TryGetValue(target, victimVal)

            ' 공격자 가치
            Dim attacker = Grid(fromR)(fromC)
            Dim attackerVal = 0
            PIECE_VALUES.TryGetValue(attacker, attackerVal)
            ' 왕 공격자는 명목 가치 사용
            If attackerVal = 0 AndAlso (attacker = CK OrElse attacker = HK) Then attackerVal = 1

            ' gain 배열: 교환 단계별 이득
            Dim gain(31) As Integer
            Dim d = 0
            gain(0) = victimVal

            ' 공격자를 칸에서 임시 제거하며 교대로 최소 공격자 탐색
            Dim currentAttackerVal = attackerVal
            Dim currentSide = If(side = CHO, HAN, CHO)  ' 다음 응수 측

            ' 공격자 임시 제거 (SEE용)
            Dim savedPieces As New List(Of (Integer, Integer, String))  ' (r, c, original)
            Grid(fromR)(fromC) = EMPTY
            savedPieces.Add((fromR, fromC, attacker))

            Do
                d += 1
                gain(d) = currentAttackerVal - gain(d - 1)  ' 이전 기물 가치 - 이전 이득
                ' stand pat: 잡지 않을 수 있으므로 (음수 이득은 0 처리는 후처리에서)
                If Math.Max(-gain(d - 1), gain(d)) < 0 Then Exit Do

                ' 다음 최소 공격자 탐색
                Dim minAttVal = GetMinAttackerValue(toR, toC, currentSide)
                If minAttVal = 0 Then Exit Do  ' 더 이상 공격자 없음

                currentAttackerVal = minAttVal

                ' 해당 최소 공격자 찾아서 임시 제거
                Dim found = RemoveMinAttacker(toR, toC, currentSide, minAttVal, savedPieces)
                If Not found Then Exit Do

                currentSide = If(currentSide = CHO, HAN, CHO)
            Loop

            ' NegaMax 방식으로 최적값 계산 (뒤에서 앞으로)
            While d > 0
                gain(d - 1) = -Math.Max(-gain(d - 1), gain(d))
                d -= 1
            End While

            ' 임시 제거한 기물 복원 (역순)
            For i = savedPieces.Count - 1 To 0 Step -1
                Dim sp = savedPieces(i)
                Grid(sp.Item1)(sp.Item2) = sp.Item3
            Next

            Return gain(0)
        End Function

        ''' <summary>
        ''' SEE 내부용: (toR,toC)를 공격하는 attackerSide의 최소 가치 공격자를 찾아 제거.
        ''' 제거한 기물 정보를 savedPieces에 추가.
        ''' </summary>
        Private Function RemoveMinAttacker(toR As Integer, toC As Integer, attackerSide As String,
                                            targetVal As Integer,
                                            savedPieces As List(Of (Integer, Integer, String))) As Boolean
            Dim minVal = targetVal
            Dim aKing, aCha, aPo, aMa, aSang, aSol, aSa As String
            Dim solFwd As Integer
            If attackerSide = CHO Then
                aKing = CK : aCha = CC : aPo = CP : aMa = CM : aSang = CE : aSol = CJ : aSa = CS
                solFwd = ChoForward
            Else
                aKing = HK : aCha = HC : aPo = HP : aMa = HM : aSang = HE : aSol = HB : aSa = HS
                solFwd = HanForward
            End If

            ' 왕 (명목 1)
            If minVal <= 1 Then
                Dim kingPalace = If(attackerSide = CHO, CHO_PALACE, HAN_PALACE)
                Dim kingEnemyPalace = If(attackerSide = CHO, HAN_PALACE, CHO_PALACE)
                For Each d In {(-1, 0), (1, 0), (0, -1), (0, 1)}
                    Dim kr = toR + d.Item1, kc = toC + d.Item2
                    If kr >= 0 AndAlso kr < 10 AndAlso kc >= 0 AndAlso kc < 9 Then
                        If Grid(kr)(kc) = aKing Then
                            If (kingPalace.Contains((kr, kc)) AndAlso kingPalace.Contains((toR, toC))) OrElse
                               (kingEnemyPalace.Contains((kr, kc)) AndAlso kingEnemyPalace.Contains((toR, toC))) Then
                                Grid(kr)(kc) = EMPTY
                                savedPieces.Add((kr, kc, aKing))
                                Return True
                            End If
                        End If
                    End If
                Next
                Dim kingDiag As List(Of (Integer, Integer)) = Nothing
                If PALACE_DIAG_MOVES.TryGetValue((toR, toC), kingDiag) Then
                    For Each t In kingDiag
                        If Grid(t.Item1)(t.Item2) = aKing Then
                            Dim kp2 = If(attackerSide = CHO, CHO_PALACE, HAN_PALACE)
                            Dim kep2 = If(attackerSide = CHO, HAN_PALACE, CHO_PALACE)
                            If (kp2.Contains(t) AndAlso kp2.Contains((toR, toC))) OrElse
                               (kep2.Contains(t) AndAlso kep2.Contains((toR, toC))) Then
                                Grid(t.Item1)(t.Item2) = EMPTY
                                savedPieces.Add((t.Item1, t.Item2, aKing))
                                Return True
                            End If
                        End If
                    Next
                End If
            End If

            ' 졸 (200)
            If minVal <= 200 Then
                Dim sr = toR - solFwd
                If sr >= 0 AndAlso sr < 10 AndAlso Grid(sr)(toC) = aSol Then
                    Grid(sr)(toC) = EMPTY
                    savedPieces.Add((sr, toC, aSol))
                    Return True
                End If
                If toC > 0 AndAlso Grid(toR)(toC - 1) = aSol Then
                    Grid(toR)(toC - 1) = EMPTY
                    savedPieces.Add((toR, toC - 1, aSol))
                    Return True
                End If
                If toC < 8 AndAlso Grid(toR)(toC + 1) = aSol Then
                    Grid(toR)(toC + 1) = EMPTY
                    savedPieces.Add((toR, toC + 1, aSol))
                    Return True
                End If
            End If

            ' 사 (300)
            If minVal <= 300 Then
                Dim diagTargets As List(Of (Integer, Integer)) = Nothing
                If PALACE_DIAG_MOVES.TryGetValue((toR, toC), diagTargets) Then
                    For Each t In diagTargets
                        If Grid(t.Item1)(t.Item2) = aSa Then
                            Grid(t.Item1)(t.Item2) = EMPTY
                            savedPieces.Add((t.Item1, t.Item2, aSa))
                            Return True
                        End If
                    Next
                End If
                Dim palace = If(attackerSide = CHO, CHO_PALACE, HAN_PALACE)
                Dim enemyPalace = If(attackerSide = CHO, HAN_PALACE, CHO_PALACE)
                If palace.Contains((toR, toC)) OrElse enemyPalace.Contains((toR, toC)) Then
                    For Each d In {(-1, 0), (1, 0), (0, -1), (0, 1)}
                        Dim nr = toR + d.Item1, nc = toC + d.Item2
                        If nr >= 0 AndAlso nr < 10 AndAlso nc >= 0 AndAlso nc < 9 Then
                            Dim targetPalace = If(palace.Contains((toR, toC)), palace, enemyPalace)
                            If targetPalace.Contains((nr, nc)) AndAlso Grid(nr)(nc) = aSa Then
                                Grid(nr)(nc) = EMPTY
                                savedPieces.Add((nr, nc, aSa))
                                Return True
                            End If
                        End If
                    Next
                End If
            End If

            ' 상 (350)
            If minVal <= 350 Then
                For Each e In _elephantAttackOffsets
                    Dim er = toR + e.Edr, ec = toC + e.Edc
                    If er >= 0 AndAlso er < 10 AndAlso ec >= 0 AndAlso ec < 9 AndAlso Grid(er)(ec) = aSang Then
                        If Grid(toR + e.W1dr)(toC + e.W1dc) = EMPTY AndAlso
                           Grid(toR + e.W2dr)(toC + e.W2dc) = EMPTY Then
                            Grid(er)(ec) = EMPTY
                            savedPieces.Add((er, ec, aSang))
                            Return True
                        End If
                    End If
                Next
            End If

            ' 마 (450)
            If minVal <= 450 Then
                For Each h In _horseAttackOffsets
                    Dim hr = toR + h.HorseDr, hc = toC + h.HorseDc
                    If hr >= 0 AndAlso hr < 10 AndAlso hc >= 0 AndAlso hc < 9 AndAlso Grid(hr)(hc) = aMa Then
                        If Grid(toR + h.WaypointDr)(toC + h.WaypointDc) = EMPTY Then
                            Grid(hr)(hc) = EMPTY
                            savedPieces.Add((hr, hc, aMa))
                            Return True
                        End If
                    End If
                Next
            End If

            ' 포 (800)
            If minVal <= 800 Then
                For Each d In {(-1, 0), (1, 0), (0, -1), (0, 1)}
                    Dim nr = toR + d.Item1, nc = toC + d.Item2
                    Dim screen = False
                    While nr >= 0 AndAlso nr < 10 AndAlso nc >= 0 AndAlso nc < 9
                        Dim p = Grid(nr)(nc)
                        If Not screen Then
                            If p <> EMPTY Then
                                If p = CP OrElse p = HP Then Exit While
                                screen = True
                            End If
                        Else
                            If p <> EMPTY Then
                                If p = aPo Then
                                    Grid(nr)(nc) = EMPTY
                                    savedPieces.Add((nr, nc, aPo))
                                    Return True
                                End If
                                Exit While
                            End If
                        End If
                        nr += d.Item1 : nc += d.Item2
                    End While
                Next
            End If

            ' 차 (1300)
            If minVal <= 1300 Then
                For Each d In {(-1, 0), (1, 0), (0, -1), (0, 1)}
                    Dim nr = toR + d.Item1, nc = toC + d.Item2
                    While nr >= 0 AndAlso nr < 10 AndAlso nc >= 0 AndAlso nc < 9
                        Dim p = Grid(nr)(nc)
                        If p <> EMPTY Then
                            If p = aCha Then
                                Grid(nr)(nc) = EMPTY
                                savedPieces.Add((nr, nc, aCha))
                                Return True
                            End If
                            Exit While
                        End If
                        nr += d.Item1 : nc += d.Item2
                    End While
                Next
            End If

            Return False
        End Function

        Public Function GetAllMoves(side As String) As List(Of ((Integer, Integer), (Integer, Integer)))
            Dim pieceNames As String()
            If side = CHO Then
                pieceNames = New String() {CK, CS, CC, CM, CE, CP, CJ}
            Else
                pieceNames = New String() {HK, HS, HC, HM, HE, HP, HB}
            End If
            Dim moves As New List(Of ((Integer, Integer), (Integer, Integer)))
            For Each pName In pieceNames
                Dim lst As List(Of (Integer, Integer)) = Nothing
                If Not _pieceLists.TryGetValue(pName, lst) Then Continue For
                For Each pos In lst
                    Dim targets = GetMovesForPiece(Grid, pos.Item1, pos.Item2, ChoForward, HanForward)
                    For Each t In targets
                        moves.Add((pos, t))
                    Next
                Next
            Next
            Return moves
        End Function

        Public Function GetLegalMoves(side As String) As List(Of ((Integer, Integer), (Integer, Integer)))
            Dim allMoves = GetAllMoves(side)
            Dim legalMoves As New List(Of ((Integer, Integer), (Integer, Integer)))
            For Each m In allMoves
                MakeMove(m.Item1, m.Item2)
                If Not IsInCheck(side) Then
                    legalMoves.Add(m)
                End If
                UndoMove()
            Next
            Return legalMoves
        End Function

        Public Function FindKing(side As String) As (Integer, Integer)?
            Return _kingPos(side)
        End Function

        Public Function IsInCheck(side As String) As Boolean
            Dim kingPos = _kingPos(side)
            If Not kingPos.HasValue Then Return True

            Dim kr = kingPos.Value.Item1
            Dim kc = kingPos.Value.Item2

            Dim eCha, ePo, eMa, eSang, eSol As String
            Dim solFwd As Integer
            If side = CHO Then
                eCha = HC : ePo = HP : eMa = HM : eSang = HE : eSol = HB
                solFwd = HanForward  ' 적(한) 졸의 전진 방향
            Else
                eCha = CC : ePo = CP : eMa = CM : eSang = CE : eSol = CJ
                solFwd = ChoForward  ' 적(초) 졸의 전진 방향
            End If

            Dim lineDirs = {(-1, 0), (1, 0), (0, -1), (0, 1)}
            For Each d In lineDirs
                Dim r = kr + d.Item1, c = kc + d.Item2
                While r >= 0 AndAlso r < 10 AndAlso c >= 0 AndAlso c < 9
                    Dim p = Grid(r)(c)
                    If p <> EMPTY Then
                        If p = eCha Then Return True
                        Exit While
                    End If
                    r += d.Item1 : c += d.Item2
                End While
            Next

            For Each d In lineDirs
                Dim r = kr + d.Item1, c = kc + d.Item2
                Dim screen = False
                While r >= 0 AndAlso r < 10 AndAlso c >= 0 AndAlso c < 9
                    Dim p = Grid(r)(c)
                    If Not screen Then
                        If p <> EMPTY Then
                            If p = CP OrElse p = HP Then Exit While
                            screen = True
                        End If
                    Else
                        If p <> EMPTY Then
                            If p = ePo Then Return True
                            Exit While
                        End If
                    End If
                    r += d.Item1 : c += d.Item2
                End While
            Next

            For Each h In _horseAttackOffsets
                Dim hr = kr + h.HorseDr, hc = kc + h.HorseDc
                If hr >= 0 AndAlso hr < 10 AndAlso hc >= 0 AndAlso hc < 9 AndAlso Grid(hr)(hc) = eMa Then
                    If Grid(kr + h.WaypointDr)(kc + h.WaypointDc) = EMPTY Then Return True
                End If
            Next

            For Each e In _elephantAttackOffsets
                Dim er = kr + e.Edr, ec = kc + e.Edc
                If er >= 0 AndAlso er < 10 AndAlso ec >= 0 AndAlso ec < 9 AndAlso Grid(er)(ec) = eSang Then
                    If Grid(kr + e.W1dr)(kc + e.W1dc) = EMPTY Then
                        If Grid(kr + e.W2dr)(kc + e.W2dc) = EMPTY Then Return True
                    End If
                End If
            Next

            Dim solDirs = {(-solFwd, 0), (0, -1), (0, 1)}
            For Each d In solDirs
                Dim sr = kr + d.Item1, sc = kc + d.Item2
                If sr >= 0 AndAlso sr < 10 AndAlso sc >= 0 AndAlso sc < 9 AndAlso Grid(sr)(sc) = eSol Then
                    Return True
                End If
            Next

            Dim diagTargets As List(Of (Integer, Integer)) = Nothing
            If PALACE_DIAG_MOVES.TryGetValue((kr, kc), diagTargets) Then
                For Each t In diagTargets
                    Dim p = Grid(t.Item1)(t.Item2)
                    If p = eSol Then
                        Dim dd = kr - t.Item1
                        If dd = solFwd OrElse dd = 0 Then Return True
                    End If
                Next
            End If

            For Each line In PALACE_DIAG_LINES
                Dim kIdx = -1
                For i = 0 To line.Length - 1
                    If line(i).Item1 = kr AndAlso line(i).Item2 = kc Then kIdx = i : Exit For
                Next
                If kIdx < 0 Then Continue For
                For i = kIdx + 1 To line.Length - 1
                    Dim lp = Grid(line(i).Item1)(line(i).Item2)
                    If lp = EMPTY Then Continue For
                    If lp = eCha Then Return True
                    Exit For
                Next
                For i = kIdx - 1 To 0 Step -1
                    Dim lp = Grid(line(i).Item1)(line(i).Item2)
                    If lp = EMPTY Then Continue For
                    If lp = eCha Then Return True
                    Exit For
                Next
            Next

            ' 궁성 중심을 왕 실제 위치 기반으로 결정
            Dim pctrR, pctrC As Integer
            Dim isCorner As Boolean
            If kr >= 5 Then  ' 아래쪽 궁성
                pctrR = 8 : pctrC = 4
                isCorner = (kr = 7 OrElse kr = 9) AndAlso (kc = 3 OrElse kc = 5)
            Else  ' 위쪽 궁성
                pctrR = 1 : pctrC = 4
                isCorner = (kr = 0 OrElse kr = 2) AndAlso (kc = 3 OrElse kc = 5)
            End If
            If isCorner Then
                Dim oppR = pctrR * 2 - kr
                Dim oppC = pctrC * 2 - kc
                If Grid(oppR)(oppC) = ePo Then
                    Dim ctrP = Grid(pctrR)(pctrC)
                    If ctrP <> EMPTY AndAlso ctrP <> CP AndAlso ctrP <> HP Then Return True
                End If
            End If

            Return False
        End Function

        Public Function IsGameOver(sideToMove As String) As Boolean
            If FindKing(sideToMove) Is Nothing Then Return True
            Dim legalMoves = GetLegalMoves(sideToMove)
            Return legalMoves.Count = 0
        End Function

        Public Function IsBikjang() As Boolean
            Dim choKing = _kingPos(CHO)
            Dim hanKing = _kingPos(HAN)
            If Not choKing.HasValue OrElse Not hanKing.HasValue Then Return False
            If choKing.Value.Item2 <> hanKing.Value.Item2 Then Return False

            Dim col = choKing.Value.Item2
            Dim minRow = Math.Min(choKing.Value.Item1, hanKing.Value.Item1)
            Dim maxRow = Math.Max(choKing.Value.Item1, hanKing.Value.Item1)
            For r = minRow + 1 To maxRow - 1
                If Grid(r)(col) <> EMPTY Then Return False
            Next
            Return True
        End Function

        ''' <summary>
        ''' 장군 원인을 상세히 반환 (디버그용, 검색에서는 사용하지 말 것)
        ''' </summary>
        Public Function DescribeCheck(side As String) As String
            Dim kingPos = _kingPos(side)
            If Not kingPos.HasValue Then Return "궁 없음"

            Dim kr = kingPos.Value.Item1
            Dim kc = kingPos.Value.Item2

            Dim eCha, ePo, eMa, eSang, eSol As String
            Dim solFwd As Integer
            If side = CHO Then
                eCha = HC : ePo = HP : eMa = HM : eSang = HE : eSol = HB
                solFwd = HanForward  ' 적(한) 졸의 전진 방향
            Else
                eCha = CC : ePo = CP : eMa = CM : eSang = CE : eSol = CJ
                solFwd = ChoForward  ' 적(초) 졸의 전진 방향
            End If

            ' 1. 차
            Dim lineDirs = {(-1, 0), (1, 0), (0, -1), (0, 1)}
            For Each d In lineDirs
                Dim r = kr + d.Item1, c = kc + d.Item2
                While r >= 0 AndAlso r < 10 AndAlso c >= 0 AndAlso c < 9
                    Dim p = Grid(r)(c)
                    If p <> EMPTY Then
                        If p = eCha Then Return $"차({p}) at ({r},{c}) → 궁({kr},{kc}) 직선 공격"
                        Exit While
                    End If
                    r += d.Item1 : c += d.Item2
                End While
            Next

            ' 2. 포
            For Each d In lineDirs
                Dim r = kr + d.Item1, c = kc + d.Item2
                Dim screen = False
                Dim screenPos As (Integer, Integer) = (0, 0)
                Dim screenPiece As String = ""
                While r >= 0 AndAlso r < 10 AndAlso c >= 0 AndAlso c < 9
                    Dim p = Grid(r)(c)
                    If Not screen Then
                        If p <> EMPTY Then
                            If p = CP OrElse p = HP Then Exit While
                            screen = True
                            screenPos = (r, c)
                            screenPiece = p
                        End If
                    Else
                        If p <> EMPTY Then
                            If p = ePo Then Return $"포({p}) at ({r},{c}) → 포대: {screenPiece} at ({screenPos.Item1},{screenPos.Item2}) → 궁({kr},{kc})"
                            Exit While
                        End If
                    End If
                    r += d.Item1 : c += d.Item2
                End While
            Next

            ' 3. 마
            For Each h In _horseAttackOffsets
                Dim hr = kr + h.HorseDr, hc = kc + h.HorseDc
                If hr >= 0 AndAlso hr < 10 AndAlso hc >= 0 AndAlso hc < 9 AndAlso Grid(hr)(hc) = eMa Then
                    Dim wr = kr + h.WaypointDr, wc = kc + h.WaypointDc
                    If Grid(wr)(wc) = EMPTY Then Return $"마({eMa}) at ({hr},{hc}) → 경유({wr},{wc}) 빈칸 → 궁({kr},{kc})"
                End If
            Next

            ' 4. 상
            For Each e In _elephantAttackOffsets
                Dim er = kr + e.Edr, ec = kc + e.Edc
                If er >= 0 AndAlso er < 10 AndAlso ec >= 0 AndAlso ec < 9 AndAlso Grid(er)(ec) = eSang Then
                    If Grid(kr + e.W1dr)(kc + e.W1dc) = EMPTY Then
                        If Grid(kr + e.W2dr)(kc + e.W2dc) = EMPTY Then
                            Return $"상({eSang}) at ({er},{ec}) → 궁({kr},{kc})"
                        End If
                    End If
                End If
            Next

            ' 5. 졸/병
            Dim solDirs = {(-solFwd, 0), (0, -1), (0, 1)}
            For Each d In solDirs
                Dim sr = kr + d.Item1, sc = kc + d.Item2
                If sr >= 0 AndAlso sr < 10 AndAlso sc >= 0 AndAlso sc < 9 AndAlso Grid(sr)(sc) = eSol Then
                    Return $"졸/병({eSol}) at ({sr},{sc}) → 궁({kr},{kc})"
                End If
            Next

            ' 6. 궁성 대각선 졸/병
            Dim diagTargets As List(Of (Integer, Integer)) = Nothing
            If PALACE_DIAG_MOVES.TryGetValue((kr, kc), diagTargets) Then
                For Each t In diagTargets
                    Dim p = Grid(t.Item1)(t.Item2)
                    If p = eSol Then
                        Dim dd = kr - t.Item1
                        If dd = solFwd OrElse dd = 0 Then Return $"졸/병({eSol}) at ({t.Item1},{t.Item2}) → 궁성 대각선 → 궁({kr},{kc})"
                    End If
                Next
            End If

            ' 6b. 궁성 대각선 차
            For Each line In PALACE_DIAG_LINES
                Dim kIdx = -1
                For i = 0 To line.Length - 1
                    If line(i).Item1 = kr AndAlso line(i).Item2 = kc Then kIdx = i : Exit For
                Next
                If kIdx < 0 Then Continue For
                For i = kIdx + 1 To line.Length - 1
                    Dim lp = Grid(line(i).Item1)(line(i).Item2)
                    If lp = EMPTY Then Continue For
                    If lp = eCha Then Return $"차({eCha}) at ({line(i).Item1},{line(i).Item2}) → 궁성 대각선 → 궁({kr},{kc})"
                    Exit For
                Next
                For i = kIdx - 1 To 0 Step -1
                    Dim lp = Grid(line(i).Item1)(line(i).Item2)
                    If lp = EMPTY Then Continue For
                    If lp = eCha Then Return $"차({eCha}) at ({line(i).Item1},{line(i).Item2}) → 궁성 대각선 → 궁({kr},{kc})"
                    Exit For
                Next
            Next

            ' 7. 궁성 대각선 포 (왕 위치 기반 궁성 결정)
            Dim pctrR, pctrC As Integer
            Dim isCorner As Boolean
            If kr >= 5 Then  ' 아래쪽 궁성
                pctrR = 8 : pctrC = 4
                isCorner = (kr = 7 OrElse kr = 9) AndAlso (kc = 3 OrElse kc = 5)
            Else  ' 위쪽 궁성
                pctrR = 1 : pctrC = 4
                isCorner = (kr = 0 OrElse kr = 2) AndAlso (kc = 3 OrElse kc = 5)
            End If
            If isCorner Then
                Dim oppR = pctrR * 2 - kr
                Dim oppC = pctrC * 2 - kc
                If Grid(oppR)(oppC) = ePo Then
                    Dim ctrP = Grid(pctrR)(pctrC)
                    If ctrP <> EMPTY AndAlso ctrP <> CP AndAlso ctrP <> HP Then
                        Return $"포({ePo}) at ({oppR},{oppC}) → 궁성 대각 포대: {ctrP} at ({pctrR},{pctrC}) → 궁({kr},{kc})"
                    End If
                End If
            End If

            Return "장군 아님"
        End Function

        Public Sub Display()
            Dim colHeader = "  " & String.Join("  ", Enumerable.Range(0, BOARD_COLS).Select(Function(i) i.ToString()))
            Console.WriteLine(colHeader)
            Console.WriteLine("  " & New String("─"c, BOARD_COLS * 3 - 1))
            For r = 0 To BOARD_ROWS - 1
                Dim rowStr = $"{r}│"
                For c = 0 To BOARD_COLS - 1
                    Dim piece = Grid(r)(c)
                    If piece = EMPTY Then
                        rowStr &= " · "
                    Else
                        rowStr &= $"{piece} "
                    End If
                Next
                Console.WriteLine(rowStr)
            Next
            Console.WriteLine()
        End Sub

        Public Overrides Function ToString() As String
            Dim lines As New List(Of String)
            For r = 0 To BOARD_ROWS - 1
                Dim rowStr = ""
                For c = 0 To BOARD_COLS - 1
                    Dim piece = Grid(r)(c)
                    If piece = EMPTY Then
                        rowStr &= " .. "
                    Else
                        rowStr &= $" {piece} "
                    End If
                Next
                lines.Add(rowStr)
            Next
            Return String.Join(Environment.NewLine, lines)
        End Function

    End Class

End Namespace
