' 형세 판단 (평가 함수)
' 양수: 초(CHO)에 유리, 음수: 한(HAN)에 유리
Imports MacroAutoControl.Constants
Imports MacroAutoControl.Engine
Imports MacroAutoControl.Engine.Pieces

Namespace MacroAutoControl.AI

    Public Module Evaluator

        ' 위치 보너스 테이블 (초 기준, 한은 행을 뒤집어 사용)

        ' 졸/병: 전진할수록 가치 증가, 중앙 선호
        Private ReadOnly SOLDIER_POS_TABLE As Integer(,) = {
            {0, 0, 0, 0, 0, 0, 0, 0, 0},
            {0, 0, 0, 0, 0, 0, 0, 0, 0},
            {0, 0, 0, 0, 0, 0, 0, 0, 0},
            {20, 0, 30, 0, 40, 0, 30, 0, 20},
            {30, 0, 40, 0, 50, 0, 40, 0, 30},
            {30, 0, 40, 0, 50, 0, 40, 0, 30},
            {10, 0, 20, 0, 30, 0, 20, 0, 10},
            {0, 0, 0, 0, 0, 0, 0, 0, 0},
            {0, 0, 0, 0, 0, 0, 0, 0, 0},
            {0, 0, 0, 0, 0, 0, 0, 0, 0}
        }

        ' 마: 중앙과 적진 선호
        Private ReadOnly HORSE_POS_TABLE As Integer(,) = {
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

        ' 포: 중앙 열 + 적진 방향 강화
        Private ReadOnly CANNON_POS_TABLE As Integer(,) = {
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

        ' 차: 적진 침투 강력 보너스 + 자진 후방 패널티
        Private ReadOnly CHARIOT_POS_TABLE As Integer(,) = {
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

        ' 상: 중앙 배치 선호
        Private ReadOnly ELEPHANT_POS_TABLE As Integer(,) = {
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

        Private Function GetPosBonus(piece As String, r As Integer, c As Integer) As Integer
            Dim side = piece(0).ToString()
            Dim ptype = piece(1).ToString()
            Dim row = If(side = CHO, r, BOARD_ROWS - 1 - r)

            Select Case ptype
                Case "J", "B" : Return SOLDIER_POS_TABLE(row, c)
                Case "M" : Return HORSE_POS_TABLE(row, c)
                Case "P" : Return CANNON_POS_TABLE(row, c)
                Case "C" : Return CHARIOT_POS_TABLE(row, c)
                Case "E" : Return ELEPHANT_POS_TABLE(row, c)
                Case Else : Return 0
            End Select
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
                ' 차가 적진(상대 절반)에 있을 때만 시너지 보너스
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

        Private Function HangingPenalty(grid As String()(), side As String) As Integer
            Dim penalty = 0
            Dim myPieces = If(side = CHO, CHO_PIECES, HAN_PIECES)
            Dim enemyPawn = If(side = CHO, HB, CJ)
            ' 적 졸이 공격할 수 있는 위치: 적 졸의 전진 반대 방향
            ' CHO 기물 → 적 HAN 졸은 forward=+1, 졸이 (r-1,c)에서 (r,c)로 공격 → 방향 -1
            ' HAN 기물 → 적 CHO 졸은 forward=-1, 졸이 (r+1,c)에서 (r,c)로 공격 → 방향 +1
            Dim pawnFwd = If(side = CHO, -1, 1)

            For r = 0 To BOARD_ROWS - 1
                For c = 0 To BOARD_COLS - 1
                    Dim piece = grid(r)(c)
                    If Not myPieces.Contains(piece) Then Continue For
                    Dim ptype = piece(1).ToString()
                    If ptype <> "C" AndAlso ptype <> "P" AndAlso ptype <> "M" Then Continue For
                    Dim value = 0
                    PIECE_VALUES.TryGetValue(piece, value)

                    Dim attackDirs = {(pawnFwd, 0), (0, -1), (0, 1)}
                    For Each d In attackDirs
                        Dim ar = r + d.Item1, ac = c + d.Item2
                        If ar >= 0 AndAlso ar < BOARD_ROWS AndAlso ac >= 0 AndAlso ac < BOARD_COLS Then
                            If grid(ar)(ac) = enemyPawn Then
                                penalty += value \ 4
                                Exit For
                            End If
                        End If
                    Next
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

            ' 인접 아군 기물에 의한 방어 (차 제외 — 차는 공격용)
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

            ' 궁이 장군 상태인지
            If board.IsInCheck(side) Then safety -= 50

            ' 궁 이동 가능 칸 수
            Dim kingMoves = GetMovesForPiece(grid, kr, kc)
            If kingMoves.Count <= 1 Then safety -= 40

            ' 적 차가 같은 행/열에 있으면 위협
            Dim enemyChariot = If(side = CHO, HC, CC)
            For col = 0 To BOARD_COLS - 1
                If col <> kc AndAlso grid(kr)(col) = enemyChariot Then safety -= 25
            Next
            For row = 0 To BOARD_ROWS - 1
                If row <> kr AndAlso grid(row)(kc) = enemyChariot Then safety -= 25
            Next

            ' 궁이 궁성 중앙에 있으면 보너스
            Dim palaceCenterR = If(side = CHO, 8, 1)
            If kr = palaceCenterR AndAlso kc = 4 Then safety += 20

            ' 사 보호 보너스
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

        Public Function Evaluate(board As Board) As Integer
            Dim score = 0
            Dim choMobility = 0
            Dim hanMobility = 0
            Dim grid = board.Grid

            Dim choChariots As New List(Of (Integer, Integer))
            Dim hanChariots As New List(Of (Integer, Integer))
            Dim choCannons As New List(Of (Integer, Integer))
            Dim hanCannons As New List(Of (Integer, Integer))

            For r = 0 To BOARD_ROWS - 1
                For c = 0 To BOARD_COLS - 1
                    Dim piece = grid(r)(c)
                    If piece = EMPTY Then Continue For

                    Dim side = piece(0).ToString()
                    Dim ptype = piece(1).ToString()
                    Dim value = 0
                    PIECE_VALUES.TryGetValue(piece, value)
                    Dim posBonus = GetPosBonus(piece, r, c)

                    If ptype = "C" Then
                        posBonus += ChariotFileBonus(grid, r, c, side)
                        If side = CHO Then choChariots.Add((r, c)) Else hanChariots.Add((r, c))
                    ElseIf ptype = "P" Then
                        posBonus += CannonMountBonus(grid, r, c)
                        If side = CHO Then choCannons.Add((r, c)) Else hanCannons.Add((r, c))
                    End If

                    If side = CHO Then
                        score += value + posBonus
                    Else
                        score -= value + posBonus
                    End If

                    ' 기동력: 차·포·마만 계산
                    If ptype = "C" OrElse ptype = "P" OrElse ptype = "M" Then
                        Dim moves = GetMovesForPiece(grid, r, c)
                        If side = CHO Then choMobility += moves.Count Else hanMobility += moves.Count
                    End If
                Next
            Next

            ' 기동력 보너스
            score += (choMobility - hanMobility) * 3

            ' 차·포 합동공격 보너스
            score += ChariotCannonSynergy(choChariots, choCannons, CHO)
            score -= ChariotCannonSynergy(hanChariots, hanCannons, HAN)

            ' 궁성 압박 보너스
            score += PalacePressure(grid, CHO, choChariots, choCannons)
            score -= PalacePressure(grid, HAN, hanChariots, hanCannons)

            ' 위험 기물 패널티
            score -= HangingPenalty(grid, CHO)
            score += HangingPenalty(grid, HAN)

            ' 궁 안전도
            score += KingSafety(board, CHO)
            score -= KingSafety(board, HAN)

            Return score
        End Function

    End Module

End Namespace
