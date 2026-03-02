' 기물별 이동 규칙 정의
Imports MacroAutoControl.Constants

Namespace MacroAutoControl.Engine

    Public Module Pieces

        Private Function InBounds(r As Integer, c As Integer) As Boolean
            Return r >= 0 AndAlso r < BOARD_ROWS AndAlso c >= 0 AndAlso c < BOARD_COLS
        End Function

        Private Function SideOf(piece As String) As String
            If piece = EMPTY Then Return Nothing
            Return piece(0).ToString()
        End Function

        Private Function IsEnemy(board As String()(), r As Integer, c As Integer, side As String) As Boolean
            Dim piece = board(r)(c)
            If piece = EMPTY Then Return False
            Return piece(0).ToString() <> side
        End Function

        Private Function IsEmptyOrEnemy(board As String()(), r As Integer, c As Integer, side As String) As Boolean
            Dim piece = board(r)(c)
            Return piece = EMPTY OrElse piece(0).ToString() <> side
        End Function

        Private Function PalaceFor(side As String) As HashSet(Of (Integer, Integer))
            If side = CHO Then Return CHO_PALACE Else Return HAN_PALACE
        End Function

        Public Function GetKingMoves(board As String()(), r As Integer, c As Integer, side As String) As List(Of (Integer, Integer))
            Dim palace = PalaceFor(side)
            Dim moves As New List(Of (Integer, Integer))
            Dim dirs = {(-1, 0), (1, 0), (0, -1), (0, 1)}
            For Each d In dirs
                Dim nr = r + d.Item1
                Dim nc = c + d.Item2
                If palace.Contains((nr, nc)) AndAlso IsEmptyOrEnemy(board, nr, nc, side) Then
                    moves.Add((nr, nc))
                End If
            Next
            Dim diagTargets As List(Of (Integer, Integer)) = Nothing
            If PALACE_DIAG_MOVES.TryGetValue((r, c), diagTargets) Then
                For Each t In diagTargets
                    If palace.Contains(t) AndAlso IsEmptyOrEnemy(board, t.Item1, t.Item2, side) Then
                        moves.Add(t)
                    End If
                Next
            End If
            Return moves
        End Function

        Public Function GetGuardMoves(board As String()(), r As Integer, c As Integer, side As String) As List(Of (Integer, Integer))
            Return GetKingMoves(board, r, c, side)
        End Function

        Public Function GetChariotMoves(board As String()(), r As Integer, c As Integer, side As String) As List(Of (Integer, Integer))
            Dim moves As New List(Of (Integer, Integer))
            Dim dirs = {(-1, 0), (1, 0), (0, -1), (0, 1)}
            For Each d In dirs
                Dim nr = r + d.Item1
                Dim nc = c + d.Item2
                While InBounds(nr, nc)
                    If board(nr)(nc) = EMPTY Then
                        moves.Add((nr, nc))
                    ElseIf IsEnemy(board, nr, nc, side) Then
                        moves.Add((nr, nc))
                        Exit While
                    Else
                        Exit While
                    End If
                    nr += d.Item1
                    nc += d.Item2
                End While
            Next
            For Each line In PALACE_DIAG_LINES
                Dim idx = -1
                For i = 0 To line.Length - 1
                    If line(i).Item1 = r AndAlso line(i).Item2 = c Then idx = i : Exit For
                Next
                If idx < 0 Then Continue For
                For i = idx + 1 To line.Length - 1
                    If board(line(i).Item1)(line(i).Item2) = EMPTY Then
                        moves.Add(line(i))
                    ElseIf IsEnemy(board, line(i).Item1, line(i).Item2, side) Then
                        moves.Add(line(i))
                        Exit For
                    Else
                        Exit For
                    End If
                Next
                For i = idx - 1 To 0 Step -1
                    If board(line(i).Item1)(line(i).Item2) = EMPTY Then
                        moves.Add(line(i))
                    ElseIf IsEnemy(board, line(i).Item1, line(i).Item2, side) Then
                        moves.Add(line(i))
                        Exit For
                    Else
                        Exit For
                    End If
                Next
            Next
            Return moves
        End Function

        Public Function GetCannonMoves(board As String()(), r As Integer, c As Integer, side As String) As List(Of (Integer, Integer))
            Dim moves As New List(Of (Integer, Integer))
            Dim cannonPieces As New HashSet(Of String) From {CP, HP}
            Dim dirs = {(-1, 0), (1, 0), (0, -1), (0, 1)}
            For Each d In dirs
                Dim nr = r + d.Item1
                Dim nc = c + d.Item2
                Dim jumped = False
                While InBounds(nr, nc)
                    Dim piece = board(nr)(nc)
                    If Not jumped Then
                        If piece <> EMPTY Then
                            If cannonPieces.Contains(piece) Then Exit While
                            jumped = True
                        End If
                    Else
                        If piece = EMPTY Then
                            moves.Add((nr, nc))
                        ElseIf cannonPieces.Contains(piece) Then
                            Exit While
                        ElseIf IsEnemy(board, nr, nc, side) Then
                            moves.Add((nr, nc))
                            Exit While
                        Else
                            Exit While
                        End If
                    End If
                    nr += d.Item1
                    nc += d.Item2
                End While
            Next
            Dim palace = PalaceFor(side)
            Dim enemyPalace = If(side = CHO, HAN_PALACE, CHO_PALACE)
            For Each pal In {palace, enemyPalace}
                If pal.Contains((r, c)) Then
                    CannonPalaceDiag(board, r, c, side, pal, moves, cannonPieces)
                End If
            Next
            Return moves
        End Function

        Private Sub CannonPalaceDiag(board As String()(), r As Integer, c As Integer, side As String,
                                      palace As HashSet(Of (Integer, Integer)),
                                      moves As List(Of (Integer, Integer)),
                                      cannonPieces As HashSet(Of String))
            Dim center As (Integer, Integer)
            Dim corners() As (Integer, Integer)

            If palace Is CHO_PALACE Then
                center = (8, 4)
                corners = {(7, 3), (7, 5), (9, 3), (9, 5)}
            Else
                center = (1, 4)
                corners = {(0, 3), (0, 5), (2, 3), (2, 5)}
            End If

            If Not palace.Contains((r, c)) Then Return

            If r = center.Item1 AndAlso c = center.Item2 Then
                Return
            End If

            If corners.Any(Function(x) x.Item1 = r AndAlso x.Item2 = c) Then
                Dim mid = board(center.Item1)(center.Item2)
                If mid = EMPTY OrElse cannonPieces.Contains(mid) Then Return
                Dim oppR = center.Item1 * 2 - r
                Dim oppC = center.Item2 * 2 - c
                If palace.Contains((oppR, oppC)) Then
                    Dim target = board(oppR)(oppC)
                    If target = EMPTY Then
                        moves.Add((oppR, oppC))
                    ElseIf Not cannonPieces.Contains(target) AndAlso IsEnemy(board, oppR, oppC, side) Then
                        moves.Add((oppR, oppC))
                    End If
                End If
            End If
        End Sub

        Public Function GetHorseMoves(board As String()(), r As Integer, c As Integer, side As String) As List(Of (Integer, Integer))
            Dim moves As New List(Of (Integer, Integer))
            Dim paths = {
                ((-1, 0), (-1, -1)), ((-1, 0), (-1, 1)),
                ((1, 0), (1, -1)), ((1, 0), (1, 1)),
                ((0, -1), (-1, -1)), ((0, -1), (1, -1)),
                ((0, 1), (-1, 1)), ((0, 1), (1, 1))
            }
            For Each p In paths
                Dim dr1 = p.Item1.Item1 : Dim dc1 = p.Item1.Item2
                Dim dr2 = p.Item2.Item1 : Dim dc2 = p.Item2.Item2
                Dim mr = r + dr1 : Dim mc = c + dc1
                If Not InBounds(mr, mc) Then Continue For
                If board(mr)(mc) <> EMPTY Then Continue For
                Dim nr = mr + dr2 : Dim nc = mc + dc2
                If Not InBounds(nr, nc) Then Continue For
                If IsEmptyOrEnemy(board, nr, nc, side) Then
                    moves.Add((nr, nc))
                End If
            Next
            Return moves
        End Function

        Public Function GetElephantMoves(board As String()(), r As Integer, c As Integer, side As String) As List(Of (Integer, Integer))
            Dim moves As New List(Of (Integer, Integer))
            Dim paths = {
                ((-1, 0), (-1, -1), (-1, -1)), ((-1, 0), (-1, 1), (-1, 1)),
                ((1, 0), (1, -1), (1, -1)), ((1, 0), (1, 1), (1, 1)),
                ((0, -1), (-1, -1), (-1, -1)), ((0, -1), (1, -1), (1, -1)),
                ((0, 1), (-1, 1), (-1, 1)), ((0, 1), (1, 1), (1, 1))
            }
            For Each p In paths
                Dim dr1 = p.Item1.Item1 : Dim dc1 = p.Item1.Item2
                Dim dr2 = p.Item2.Item1 : Dim dc2 = p.Item2.Item2
                Dim dr3 = p.Item3.Item1 : Dim dc3 = p.Item3.Item2
                Dim r1 = r + dr1 : Dim c1 = c + dc1
                If Not InBounds(r1, c1) OrElse board(r1)(c1) <> EMPTY Then Continue For
                Dim r2 = r1 + dr2 : Dim c2 = c1 + dc2
                If Not InBounds(r2, c2) OrElse board(r2)(c2) <> EMPTY Then Continue For
                Dim r3 = r2 + dr3 : Dim c3 = c2 + dc3
                If Not InBounds(r3, c3) Then Continue For
                If IsEmptyOrEnemy(board, r3, c3, side) Then
                    moves.Add((r3, c3))
                End If
            Next
            Return moves
        End Function

        Public Function GetSoldierMoves(board As String()(), r As Integer, c As Integer, side As String, Optional choFwd As Integer = -1, Optional hanFwd As Integer = 1) As List(Of (Integer, Integer))
            Dim moves As New List(Of (Integer, Integer))
            Dim forward = If(side = CHO, choFwd, hanFwd)

            Dim nr = r + forward
            If InBounds(nr, c) AndAlso IsEmptyOrEnemy(board, nr, c, side) Then
                moves.Add((nr, c))
            End If
            For Each dc In {-1, 1}
                Dim nc = c + dc
                If InBounds(r, nc) AndAlso IsEmptyOrEnemy(board, r, nc, side) Then
                    moves.Add((r, nc))
                End If
            Next
            Dim enemyPalace = If(side = CHO, HAN_PALACE, CHO_PALACE)
            Dim myPalace = PalaceFor(side)
            For Each pal In {myPalace, enemyPalace}
                If pal.Contains((r, c)) Then
                    Dim diagTargets As List(Of (Integer, Integer)) = Nothing
                    If PALACE_DIAG_MOVES.TryGetValue((r, c), diagTargets) Then
                        For Each t In diagTargets
                            If pal.Contains(t) Then
                                If (t.Item1 - r) = forward OrElse (t.Item1 - r) = 0 Then
                                    If IsEmptyOrEnemy(board, t.Item1, t.Item2, side) Then
                                        If Not moves.Contains(t) Then
                                            moves.Add(t)
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            Next
            Return moves
        End Function

        Public Function GetMovesForPiece(board As String()(), r As Integer, c As Integer, Optional choFwd As Integer = -1, Optional hanFwd As Integer = 1) As List(Of (Integer, Integer))
            Dim piece = board(r)(c)
            If piece = EMPTY Then Return New List(Of (Integer, Integer))

            Dim side = piece(0).ToString()
            Dim pieceType = piece(1).ToString()

            Select Case pieceType
                Case "K" : Return GetKingMoves(board, r, c, side)
                Case "S" : Return GetGuardMoves(board, r, c, side)
                Case "C" : Return GetChariotMoves(board, r, c, side)
                Case "P" : Return GetCannonMoves(board, r, c, side)
                Case "M" : Return GetHorseMoves(board, r, c, side)
                Case "E" : Return GetElephantMoves(board, r, c, side)
                Case "J" : Return GetSoldierMoves(board, r, c, side, choFwd, hanFwd)
                Case "B" : Return GetSoldierMoves(board, r, c, side, choFwd, hanFwd)
                Case Else : Return New List(Of (Integer, Integer))
            End Select
        End Function

    End Module

End Namespace
