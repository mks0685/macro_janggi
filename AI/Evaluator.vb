' 형세 판단 (평가 함수) - 강화판: Tapered Eval + 추가 항목
Imports MacroAutoControl.Constants
Imports MacroAutoControl.Engine
Imports MacroAutoControl.Engine.Pieces

Namespace MacroAutoControl.AI

    Public Module Evaluator

        ' ===== MG (중반) PST =====
        Private ReadOnly SOLDIER_MG_PST As Integer(,) = {
            {0, 0, 0, 0, 0, 0, 0, 0, 0},
            {0, 0, 0, 0, 0, 0, 0, 0, 0},
            {0, 0, 0, 0, 0, 0, 0, 0, 0},
            {20, 10, 30, 20, 40, 20, 30, 10, 20},
            {30, 15, 40, 25, 50, 25, 40, 15, 30},
            {30, 15, 40, 25, 50, 25, 40, 15, 30},
            {10, 5, 20, 10, 30, 10, 20, 5, 10},
            {5, 5, 5, 5, 10, 5, 5, 5, 5},
            {0, 0, 0, 0, 0, 0, 0, 0, 0},
            {0, 0, 0, 0, 0, 0, 0, 0, 0}
        }

        ' ===== EG (종반) PST: 졸 - 전진 보너스 강화 =====
        Private ReadOnly SOLDIER_EG_PST As Integer(,) = {
            {0, 0, 0, 0, 0, 0, 0, 0, 0},
            {0, 0, 0, 0, 0, 0, 0, 0, 0},
            {0, 0, 0, 0, 0, 0, 0, 0, 0},
            {40, 30, 50, 40, 60, 40, 50, 30, 40},
            {50, 40, 60, 50, 70, 50, 60, 40, 50},
            {50, 40, 60, 50, 70, 50, 60, 40, 50},
            {25, 20, 35, 25, 45, 25, 35, 20, 25},
            {10, 10, 15, 10, 20, 10, 15, 10, 10},
            {0, 0, 0, 0, 5, 0, 0, 0, 0},
            {0, 0, 0, 0, 0, 0, 0, 0, 0}
        }

        Private ReadOnly HORSE_MG_PST As Integer(,) = {
            {0, 10, 10, 10, 10, 10, 10, 10, 0},
            {10, 20, 20, 20, 20, 20, 20, 20, 10},
            {10, 20, 30, 30, 30, 30, 30, 20, 10},
            {10, 20, 30, 40, 40, 40, 30, 20, 10},
            {10, 20, 30, 40, 50, 40, 30, 20, 10},
            {10, 20, 30, 40, 50, 40, 30, 20, 10},
            {10, 20, 30, 30, 30, 30, 30, 20, 10},
            {10, 20, 20, 20, 20, 20, 20, 20, 10},
            {0, 10, 10, 10, 10, 10, 10, 10, 0},
            {0, 0, 0, 10, 10, 10, 0, 0, 0}
        }

        Private ReadOnly HORSE_EG_PST As Integer(,) = {
            {0, 5, 5, 10, 10, 10, 5, 5, 0},
            {5, 15, 15, 20, 20, 20, 15, 15, 5},
            {5, 15, 25, 30, 30, 30, 25, 15, 5},
            {10, 20, 30, 35, 40, 35, 30, 20, 10},
            {10, 20, 30, 40, 45, 40, 30, 20, 10},
            {10, 20, 30, 40, 45, 40, 30, 20, 10},
            {5, 15, 25, 30, 30, 30, 25, 15, 5},
            {5, 15, 15, 20, 20, 20, 15, 15, 5},
            {0, 5, 5, 10, 10, 10, 5, 5, 0},
            {0, 0, 0, 5, 5, 5, 0, 0, 0}
        }

        ' 포 MG: 대포대 위치에서 보너스
        Private ReadOnly CANNON_MG_PST As Integer(,) = {
            {0, 0, 15, 20, 25, 20, 15, 0, 0},
            {0, 0, 15, 25, 30, 25, 15, 0, 0},
            {0, 10, 20, 30, 35, 30, 20, 10, 0},
            {0, 10, 25, 30, 35, 30, 25, 10, 0},
            {0, 5, 15, 25, 30, 25, 15, 5, 0},
            {0, 5, 10, 20, 25, 20, 10, 5, 0},
            {0, 0, 10, 15, 15, 15, 10, 0, 0},
            {0, 0, 0, 0, 0, 0, 0, 0, 0},
            {0, 0, 0, 0, 0, 0, 0, 0, 0},
            {0, 0, 0, 0, 0, 0, 0, 0, 0}
        }

        ' 포 EG: 종반에서 약화
        Private ReadOnly CANNON_EG_PST As Integer(,) = {
            {0, 0, 5, 10, 15, 10, 5, 0, 0},
            {0, 0, 5, 10, 15, 10, 5, 0, 0},
            {0, 5, 10, 15, 20, 15, 10, 5, 0},
            {0, 5, 10, 15, 20, 15, 10, 5, 0},
            {0, 0, 5, 10, 15, 10, 5, 0, 0},
            {0, 0, 5, 10, 10, 10, 5, 0, 0},
            {0, 0, 5, 5, 5, 5, 5, 0, 0},
            {0, 0, 0, 0, 0, 0, 0, 0, 0},
            {0, 0, 0, 0, 0, 0, 0, 0, 0},
            {0, 0, 0, 0, 0, 0, 0, 0, 0}
        }

        Private ReadOnly CHARIOT_MG_PST As Integer(,) = {
            {50, 55, 60, 70, 80, 70, 60, 55, 50},
            {40, 45, 50, 55, 65, 55, 50, 45, 40},
            {30, 35, 40, 45, 50, 45, 40, 35, 30},
            {20, 20, 25, 30, 35, 30, 25, 20, 20},
            {10, 10, 15, 20, 25, 20, 15, 10, 10},
            {5, 5, 10, 10, 15, 10, 10, 5, 5},
            {0, 0, 0, 0, 5, 0, 0, 0, 0},
            {-5, -5, 0, 0, 0, 0, 0, -5, -5},
            {-15, -10, -5, -10, -15, -10, -5, -10, -15},
            {-20, -15, -10, -15, -20, -15, -10, -15, -20}
        }

        Private ReadOnly CHARIOT_EG_PST As Integer(,) = {
            {60, 60, 65, 75, 85, 75, 65, 60, 60},
            {50, 50, 55, 60, 70, 60, 55, 50, 50},
            {40, 40, 45, 50, 55, 50, 45, 40, 40},
            {30, 30, 35, 40, 45, 40, 35, 30, 30},
            {20, 20, 25, 30, 35, 30, 25, 20, 20},
            {10, 10, 15, 20, 25, 20, 15, 10, 10},
            {5, 5, 5, 10, 15, 10, 5, 5, 5},
            {0, 0, 5, 5, 10, 5, 5, 0, 0},
            {-5, 0, 0, 0, 0, 0, 0, 0, -5},
            {-10, -5, -5, -5, -10, -5, -5, -5, -10}
        }

        Private ReadOnly ELEPHANT_MG_PST As Integer(,) = {
            {0, 10, 10, 10, 10, 10, 10, 10, 0},
            {5, 15, 15, 15, 20, 15, 15, 15, 5},
            {5, 15, 20, 20, 20, 20, 20, 15, 5},
            {5, 10, 15, 20, 20, 20, 15, 10, 5},
            {0, 5, 10, 15, 15, 15, 10, 5, 0},
            {0, 5, 10, 15, 15, 15, 10, 5, 0},
            {0, 5, 10, 10, 10, 10, 10, 5, 0},
            {0, 5, 5, 5, 5, 5, 5, 5, 0},
            {0, 0, 0, 0, 5, 0, 0, 0, 0},
            {0, 0, 0, 0, 0, 0, 0, 0, 0}
        }

        Private ReadOnly ELEPHANT_EG_PST As Integer(,) = {
            {0, 5, 5, 5, 5, 5, 5, 5, 0},
            {5, 10, 10, 10, 10, 10, 10, 10, 5},
            {5, 10, 15, 15, 15, 15, 15, 10, 5},
            {5, 10, 10, 15, 15, 15, 10, 10, 5},
            {0, 5, 10, 10, 10, 10, 10, 5, 0},
            {0, 5, 5, 10, 10, 10, 5, 5, 0},
            {0, 0, 5, 5, 5, 5, 5, 0, 0},
            {0, 0, 0, 0, 5, 0, 0, 0, 0},
            {0, 0, 0, 0, 0, 0, 0, 0, 0},
            {0, 0, 0, 0, 0, 0, 0, 0, 0}
        }

        Private Function GetPosBonusMG(piece As String, r As Integer, c As Integer) As Integer
            Dim side = piece(0).ToString()
            Dim ptype = piece(1).ToString()
            Dim row = If(side = CHO, r, BOARD_ROWS - 1 - r)

            Select Case ptype
                Case "J", "B" : Return SOLDIER_MG_PST(row, c)
                Case "M" : Return HORSE_MG_PST(row, c)
                Case "P" : Return CANNON_MG_PST(row, c)
                Case "C" : Return CHARIOT_MG_PST(row, c)
                Case "E" : Return ELEPHANT_MG_PST(row, c)
                Case Else : Return 0
            End Select
        End Function

        Private Function GetPosBonusEG(piece As String, r As Integer, c As Integer) As Integer
            Dim side = piece(0).ToString()
            Dim ptype = piece(1).ToString()
            Dim row = If(side = CHO, r, BOARD_ROWS - 1 - r)

            Select Case ptype
                Case "J", "B" : Return SOLDIER_EG_PST(row, c)
                Case "M" : Return HORSE_EG_PST(row, c)
                Case "P" : Return CANNON_EG_PST(row, c)
                Case "C" : Return CHARIOT_EG_PST(row, c)
                Case "E" : Return ELEPHANT_EG_PST(row, c)
                Case Else : Return 0
            End Select
        End Function

        ''' <summary>게임 단계 계산: 비졸 기물의 남은 가치 합 (양 진영)</summary>
        Private Function CalcPhase(board As Board) As Integer
            Dim phase = 0
            Dim grid = board.Grid
            For r = 0 To BOARD_ROWS - 1
                For c = 0 To BOARD_COLS - 1
                    Dim piece = grid(r)(c)
                    If piece = EMPTY Then Continue For
                    Dim ptype = piece(1).ToString()
                    If ptype = "J" OrElse ptype = "B" OrElse ptype = "K" Then Continue For
                    Dim v = 0
                    PIECE_VALUES.TryGetValue(piece, v)
                    phase += v
                Next
            Next
            ' 양 진영 합이므로 TOTAL_PHASE * 2가 최대
            Return Math.Min(phase, TOTAL_PHASE * 2)
        End Function

        Private Function ChariotFileBonus(grid As String()(), r As Integer, c As Integer, side As String) As Integer
            Dim myPawn = If(side = CHO, CJ, HB)
            Dim enemyPawn = If(side = CHO, HB, CJ)
            Dim hasMyPawn = False
            Dim hasEnemyPawn = False
            For row = 0 To 9
                Dim p = grid(row)(c)
                If p = myPawn Then hasMyPawn = True
                If p = enemyPawn Then hasEnemyPawn = True
            Next
            If Not hasMyPawn AndAlso Not hasEnemyPawn Then Return 30
            If Not hasMyPawn Then Return 15
            Return 0
        End Function

        Private Function CannonMountBonus(grid As String()(), r As Integer, c As Integer) As Integer
            Dim mounts = 0
            Dim dirs = {(0, 1), (0, -1), (1, 0), (-1, 0)}
            For Each d In dirs
                Dim nr = r + d.Item1, nc = c + d.Item2
                While nr >= 0 AndAlso nr < 10 AndAlso nc >= 0 AndAlso nc < 9
                    Dim p = grid(nr)(nc)
                    If p <> EMPTY Then
                        If p(1) <> "P"c Then mounts += 1
                        Exit While
                    End If
                    nr += d.Item1 : nc += d.Item2
                End While
            Next
            Dim bonuses = {-30, 0, 20, 30, 30}
            Return bonuses(Math.Min(mounts, 4))
        End Function

        Private Function ChariotCannonSynergy(chariotPos As List(Of (Integer, Integer)),
                                               cannonPos As List(Of (Integer, Integer)),
                                               side As String) As Integer
            Dim bonus = 0
            For Each cPos In chariotPos
                Dim inEnemyHalf = If(side = CHO, cPos.Item1 <= 4, cPos.Item1 >= 5)
                If Not inEnemyHalf Then Continue For

                For Each pp In cannonPos
                    If cPos.Item2 = pp.Item2 Then bonus += 40
                    If cPos.Item1 = pp.Item1 Then bonus += 30
                Next
            Next
            If chariotPos.Count = 2 Then
                If chariotPos(0).Item2 = chariotPos(1).Item2 Then bonus += 25
            End If
            Return bonus
        End Function

        Private Function PalacePressure(grid As String()(), side As String,
                                         chariotPos As List(Of (Integer, Integer)),
                                         cannonPos As List(Of (Integer, Integer))) As Integer
            Dim enemyPalaceRows As IEnumerable(Of Integer)
            Dim enemyKing As String
            If side = CHO Then
                enemyPalaceRows = Enumerable.Range(0, 3)
                enemyKing = HK
            Else
                enemyPalaceRows = Enumerable.Range(7, 3)
                enemyKing = CK
            End If

            Dim bonus = 0
            Dim enemyKingR As Integer = -1, enemyKingC As Integer = -1
            For Each r In enemyPalaceRows
                For c = 3 To 5
                    If grid(r)(c) = enemyKing Then
                        enemyKingR = r : enemyKingC = c
                    End If
                Next
            Next

            Dim enemyHalf = If(side = CHO, Enumerable.Range(0, 5), Enumerable.Range(5, 5))
            For Each cPos In chariotPos
                If enemyHalf.Contains(cPos.Item1) AndAlso cPos.Item2 >= 3 AndAlso cPos.Item2 <= 5 Then
                    bonus += 30
                End If
            Next

            For Each pp In cannonPos
                If pp.Item2 >= 3 AndAlso pp.Item2 <= 5 Then
                    Dim dr = If(side = CHO, -1, 1)
                    Dim nr = pp.Item1 + dr
                    Dim foundMount = False
                    While nr >= 0 AndAlso nr < BOARD_ROWS
                        Dim p = grid(nr)(pp.Item2)
                        If p <> EMPTY Then
                            If p(1) <> "P"c Then foundMount = True
                            Exit While
                        End If
                        nr += dr
                    End While
                    If foundMount Then bonus += 20
                End If
            Next

            If enemyKingR >= 0 Then
                Dim hasChariotOnKing = False
                Dim hasCannonOnKing = False
                For Each cPos In chariotPos
                    If cPos.Item1 = enemyKingR OrElse cPos.Item2 = enemyKingC Then
                        hasChariotOnKing = True : Exit For
                    End If
                Next
                For Each pp In cannonPos
                    If pp.Item1 = enemyKingR OrElse pp.Item2 = enemyKingC Then
                        hasCannonOnKing = True : Exit For
                    End If
                Next
                If hasChariotOnKing AndAlso hasCannonOnKing Then bonus += 40
            End If

            Return bonus
        End Function

        ''' <summary>
        ''' 기물 취약성 패널티: 각 기물이 적에게 공격받는지 검사하고,
        ''' 공격자/방어자 가치 비교로 패널티 부과
        ''' </summary>
        Private Function PieceVulnerabilityPenalty(board As Board, side As String) As Integer
            Dim penalty = 0
            Dim enemySide = If(side = CHO, HAN, CHO)

            ' 사/상/마/포/차 대상 (졸/왕 제외)
            Dim pieceNames As String()
            If side = CHO Then
                pieceNames = New String() {CS, CE, CM, CP, CC}
            Else
                pieceNames = New String() {HS, HE, HM, HP, HC}
            End If

            For Each pName In pieceNames
                Dim lst = board.GetPieceList(pName)
                For Each pos In lst
                    Dim r = pos.Item1, c = pos.Item2
                    Dim pieceValue = 0
                    PIECE_VALUES.TryGetValue(pName, pieceValue)

                    ' 적의 최소 공격자 가치
                    Dim attackerValue = board.GetMinAttackerValue(r, c, enemySide)
                    If attackerValue = 0 Then Continue For  ' 공격 없음

                    ' 아군 방어자 가치
                    Dim defenderValue = board.GetMinAttackerValue(r, c, side)

                    If defenderValue = 0 Then
                        ' 공격받고 방어 없음 → 기물 가치의 40% 패널티
                        penalty += pieceValue * 40 \ 100
                    ElseIf attackerValue < pieceValue Then
                        ' 공격자가 더 쌈 + 방어 있음 → 차액의 25% 패널티
                        penalty += (pieceValue - attackerValue) * 25 \ 100
                    End If
                    ' 공격자가 같거나 비싸고 방어 있음 → 패널티 0
                Next
            Next
            Return penalty
        End Function

        Private Function KingSafety(board As Board, side As String) As Integer
            Dim kingPos = board.FindKing(side)
            If Not kingPos.HasValue Then Return -9999

            Dim kr = kingPos.Value.Item1
            Dim kc = kingPos.Value.Item2
            Dim grid = board.Grid
            Dim safety = 0
            Dim myPieces = If(side = CHO, CHO_PIECES, HAN_PIECES)
            Dim palace = If(side = CHO, CHO_PALACE, HAN_PALACE)

            Dim myChariot = If(side = CHO, CC, HC)
            For dr = -1 To 1
                For dc = -1 To 1
                    If dr = 0 AndAlso dc = 0 Then Continue For
                    Dim nr = kr + dr, nc = kc + dc
                    If nr >= 0 AndAlso nr < BOARD_ROWS AndAlso nc >= 0 AndAlso nc < BOARD_COLS Then
                        Dim adj = grid(nr)(nc)
                        If myPieces.Contains(adj) AndAlso adj <> myChariot Then safety += 15
                    End If
                Next
            Next

            If board.IsInCheck(side) Then safety -= 50

            Dim kingMoves = GetMovesForPiece(grid, kr, kc)
            If kingMoves.Count <= 1 Then safety -= 40

            Dim enemyChariot = If(side = CHO, HC, CC)
            For col = 0 To BOARD_COLS - 1
                If col <> kc AndAlso grid(kr)(col) = enemyChariot Then safety -= 25
            Next
            For row = 0 To BOARD_ROWS - 1
                If row <> kr AndAlso grid(row)(kc) = enemyChariot Then safety -= 25
            Next

            Dim palaceCenterR = If(side = CHO, 8, 1)
            If kr = palaceCenterR AndAlso kc = 4 Then safety += 20

            Dim myGuard = If(side = CHO, CS, HS)
            For dr = -1 To 1
                For dc = -1 To 1
                    If dr = 0 AndAlso dc = 0 Then Continue For
                    Dim nr = kr + dr, nc = kc + dc
                    If palace.Contains((nr, nc)) AndAlso
                       nr >= 0 AndAlso nr < BOARD_ROWS AndAlso nc >= 0 AndAlso nc < BOARD_COLS AndAlso
                       grid(nr)(nc) = myGuard Then
                        safety += 20
                    End If
                Next
            Next

            Return safety
        End Function

        ' ===== Phase 3C: 추가 평가 항목 =====

        ''' <summary>연결된 졸 보너스: 인접한 아군 졸이 있으면 보너스</summary>
        Private Function ConnectedSoldierBonus(grid As String()(), r As Integer, c As Integer, side As String) As Integer
            Dim myPawn = If(side = CHO, CJ, HB)
            Dim bonus = 0
            For Each dc In {-1, 1}
                Dim nc = c + dc
                If nc >= 0 AndAlso nc < BOARD_COLS AndAlso grid(r)(nc) = myPawn Then
                    bonus += 25
                End If
            Next
            Return bonus
        End Function

        ''' <summary>통과 졸 감지: 전방에 적 졸이 없으면 보너스</summary>
        Private Function PassedSoldierBonus(grid As String()(), r As Integer, c As Integer, side As String, forward As Integer) As Integer
            Dim enemyPawn = If(side = CHO, HB, CJ)
            ' 전방 3열(자기 열 + 양 옆) 검사
            Dim sr = r + forward
            While sr >= 0 AndAlso sr < BOARD_ROWS
                For dc = -1 To 1
                    Dim nc = c + dc
                    If nc >= 0 AndAlso nc < BOARD_COLS AndAlso grid(sr)(nc) = enemyPawn Then
                        Return 0  ' 적 졸이 있으면 통과 졸 아님
                    End If
                Next
                sr += forward
            End While
            ' 적진에 가까울수록 보너스 증가
            Dim distToEnd = If(forward < 0, r, BOARD_ROWS - 1 - r)
            If distToEnd <= 2 Then Return 80
            If distToEnd <= 4 Then Return 60
            Return 40
        End Function

        ''' <summary>포 대포대 부족 패널티: 잔여 기물 < 8이면 포에 패널티</summary>
        Private Function CannonLackOfMountPenalty(board As Board, side As String) As Integer
            ' 양 진영의 총 기물 수 (왕, 졸 제외 비졸 기물)
            Dim totalPieces = 0
            Dim grid = board.Grid
            For r = 0 To BOARD_ROWS - 1
                For c = 0 To BOARD_COLS - 1
                    Dim p = grid(r)(c)
                    If p = EMPTY Then Continue For
                    Dim pt = p(1).ToString()
                    If pt <> "K" AndAlso pt <> "J" AndAlso pt <> "B" Then totalPieces += 1
                Next
            Next
            If totalPieces < 8 Then
                ' 포 개수 × 패널티
                Dim myCannon = If(side = CHO, CP, HP)
                Dim cannonCount = board.GetPieceList(myCannon).Count
                Return cannonCount * 80
            End If
            Return 0
        End Function

        ''' <summary>마 전초 보너스: 적 졸이 공격할 수 없는 적진 위치</summary>
        Private Function HorseOutpostBonus(grid As String()(), r As Integer, c As Integer, side As String, forward As Integer) As Integer
            Dim enemyPawn = If(side = CHO, HB, CJ)
            ' 적 진영에 있는지 확인
            Dim inEnemyHalf = If(side = CHO, r <= 4, r >= 5)
            If Not inEnemyHalf Then Return 0
            ' 적 졸이 좌/우 전방에서 공격 가능한지 확인
            Dim enemyFwd = -forward
            For Each dc In {-1, 1}
                Dim nr = r + enemyFwd, nc = c + dc
                If nr >= 0 AndAlso nr < BOARD_ROWS AndAlso nc >= 0 AndAlso nc < BOARD_COLS Then
                    If grid(nr)(nc) = enemyPawn Then Return 0
                End If
            Next
            Return 30
        End Function

        ''' <summary>상 차단 패널티: 상의 양쪽 경로가 모두 막힘</summary>
        Private Function ElephantBlockedPenalty(grid As String()(), r As Integer, c As Integer) As Integer
            ' 상이 이동하려면 직선 1칸 + 대각 2칸 경로가 필요
            ' 4방향 중 최소 1개라도 열려있으면 패널티 없음
            Dim dirs = {(-1, 0), (1, 0), (0, -1), (0, 1)}
            Dim blockedDirs = 0
            Dim totalDirs = 0
            For Each d In dirs
                Dim r1 = r + d.Item1, c1 = c + d.Item2
                If r1 < 0 OrElse r1 >= BOARD_ROWS OrElse c1 < 0 OrElse c1 >= BOARD_COLS Then Continue For
                totalDirs += 1
                If grid(r1)(c1) <> EMPTY Then blockedDirs += 1
            Next
            If totalDirs > 0 AndAlso blockedDirs = totalDirs Then Return 20
            Return 0
        End Function

        ''' <summary>왕 접근도: 차/포의 적 왕 Manhattan 거리 기반 보너스</summary>
        Private Function KingProximityBonus(piecePos As List(Of (Integer, Integer)), enemyKingPos As (Integer, Integer)?) As Integer
            If Not enemyKingPos.HasValue Then Return 0
            Dim ekr = enemyKingPos.Value.Item1
            Dim ekc = enemyKingPos.Value.Item2
            Dim bonus = 0
            For Each pos In piecePos
                Dim dist = Math.Abs(pos.Item1 - ekr) + Math.Abs(pos.Item2 - ekc)
                ' 거리가 가까울수록 보너스
                bonus += Math.Max(0, (14 - dist) * 3)
            Next
            Return bonus
        End Function

        ' ===== 메인 평가 함수 =====
        Public Function Evaluate(board As Board) As Integer
            Dim mgScore = 0
            Dim egScore = 0
            Dim choMobility = 0
            Dim hanMobility = 0
            Dim grid = board.Grid

            Dim choChariots As New List(Of (Integer, Integer))
            Dim hanChariots As New List(Of (Integer, Integer))
            Dim choCannons As New List(Of (Integer, Integer))
            Dim hanCannons As New List(Of (Integer, Integer))

            ' 게임 단계 계산
            Dim phase = CalcPhase(board)
            Dim totalPhase2 = TOTAL_PHASE * 2

            For r = 0 To BOARD_ROWS - 1
                For c = 0 To BOARD_COLS - 1
                    Dim piece = grid(r)(c)
                    If piece = EMPTY Then Continue For

                    Dim side = piece(0).ToString()
                    Dim ptype = piece(1).ToString()
                    Dim mgValue = 0
                    Dim egValue = 0
                    PIECE_VALUES.TryGetValue(piece, mgValue)
                    PIECE_VALUES_EG.TryGetValue(piece, egValue)
                    Dim mgPos = GetPosBonusMG(piece, r, c)
                    Dim egPos = GetPosBonusEG(piece, r, c)

                    If ptype = "C" Then
                        Dim fileBonus = ChariotFileBonus(grid, r, c, side)
                        mgPos += fileBonus
                        egPos += fileBonus
                        If side = CHO Then choChariots.Add((r, c)) Else hanChariots.Add((r, c))
                    ElseIf ptype = "P" Then
                        mgPos += CannonMountBonus(grid, r, c)
                        If side = CHO Then choCannons.Add((r, c)) Else hanCannons.Add((r, c))
                    End If

                    Dim mgTotal = mgValue + mgPos
                    Dim egTotal = egValue + egPos

                    ' 추가 평가 항목 (Phase 3C)
                    If ptype = "J" OrElse ptype = "B" Then
                        Dim fwd = If(side = CHO, board.ChoForward, board.HanForward)
                        Dim connected = ConnectedSoldierBonus(grid, r, c, side)
                        Dim passed = PassedSoldierBonus(grid, r, c, side, fwd)
                        mgTotal += connected
                        egTotal += connected + passed  ' 통과 졸은 종반에서 더 가치
                    ElseIf ptype = "M" Then
                        Dim fwd = If(side = CHO, board.ChoForward, board.HanForward)
                        Dim outpost = HorseOutpostBonus(grid, r, c, side, fwd)
                        mgTotal += outpost
                        egTotal += outpost
                    ElseIf ptype = "E" Then
                        Dim blocked = ElephantBlockedPenalty(grid, r, c)
                        mgTotal -= blocked
                        egTotal -= blocked
                    End If

                    If side = CHO Then
                        mgScore += mgTotal
                        egScore += egTotal
                    Else
                        mgScore -= mgTotal
                        egScore -= egTotal
                    End If

                    If ptype = "C" OrElse ptype = "P" OrElse ptype = "M" Then
                        Dim moves = GetMovesForPiece(grid, r, c)
                        If side = CHO Then choMobility += moves.Count Else hanMobility += moves.Count
                    End If
                Next
            Next

            ' 포 대포대 부족 패널티
            Dim choCannonPenalty = CannonLackOfMountPenalty(board, CHO)
            Dim hanCannonPenalty = CannonLackOfMountPenalty(board, HAN)
            egScore -= choCannonPenalty
            egScore += hanCannonPenalty

            ' 왕 접근도 (종반 가중)
            Dim choKingProx = KingProximityBonus(choChariots, board.FindKing(HAN)) +
                              KingProximityBonus(choCannons, board.FindKing(HAN))
            Dim hanKingProx = KingProximityBonus(hanChariots, board.FindKing(CHO)) +
                              KingProximityBonus(hanCannons, board.FindKing(CHO))
            egScore += choKingProx - hanKingProx

            Dim mobilityBonus = (choMobility - hanMobility) * 3
            mgScore += mobilityBonus
            egScore += mobilityBonus

            Dim synergy = ChariotCannonSynergy(choChariots, choCannons, CHO) -
                          ChariotCannonSynergy(hanChariots, hanCannons, HAN)
            mgScore += synergy
            egScore += synergy \ 2  ' 종반에서 시너지 약화

            Dim pressure = PalacePressure(grid, CHO, choChariots, choCannons) -
                           PalacePressure(grid, HAN, hanChariots, hanCannons)
            mgScore += pressure
            egScore += pressure

            Dim hanging = PieceVulnerabilityPenalty(board, CHO) - PieceVulnerabilityPenalty(board, HAN)
            mgScore -= hanging
            egScore -= hanging

            Dim safety = KingSafety(board, CHO) - KingSafety(board, HAN)
            mgScore += safety
            egScore += safety \ 2  ' 종반에서 왕 안전 덜 중요

            ' Tapered Eval: score = (mgScore * phase + egScore * (total - phase)) / total
            Dim score = (mgScore * phase + egScore * (totalPhase2 - phase)) \ totalPhase2

            Return score
        End Function

    End Module

End Namespace
