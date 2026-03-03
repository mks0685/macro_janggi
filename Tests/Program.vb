''' 상 뜨면서 포장 반복 금지 검증 테스트
''' 시나리오: CE(5,6) → (2,4) 이동 시 CP(5,4)가 HK(1,4)를 포장(장군)
'''           HK 피하고 CE 돌아오면 동일 국면 반복 → AI가 회피해야 함

Imports MacroAutoControl.Constants
Imports MacroAutoControl.Engine
Imports MacroAutoControl.AI

Module Program

    Private Function GetPieceName(piece As String) As String
        Select Case piece
            Case CK : Return "초궁"
            Case CS : Return "초사"
            Case CC : Return "초차"
            Case CM : Return "초마"
            Case CE : Return "초상"
            Case CP : Return "초포"
            Case CJ : Return "초졸"
            Case HK : Return "한궁"
            Case HS : Return "한사"
            Case HC : Return "한차"
            Case HM : Return "한마"
            Case HE : Return "한상"
            Case HP : Return "한포"
            Case HB : Return "한병"
            Case Else : Return piece
        End Select
    End Function

    ''' <summary>빈 보드 생성</summary>
    Private Function EmptyGrid() As String()()
        Dim grid(BOARD_ROWS - 1)() As String
        For r = 0 To BOARD_ROWS - 1
            grid(r) = New String(BOARD_COLS - 1) {}
            For c = 0 To BOARD_COLS - 1
                grid(r)(c) = EMPTY
            Next
        Next
        Return grid
    End Function

    Sub Main()
        Console.OutputEncoding = System.Text.Encoding.UTF8
        Console.WriteLine("=== 상 뜨면서 포장 반복 금지 테스트 ===")
        Console.WriteLine()

        ' ── 테스트 1: IsRepetition 기본 동작 검증 ──
        Test1_IsRepetition()
        Console.WriteLine()

        ' ── 테스트 2: 상 뜨면서 포장 시나리오 ──
        Test2_ElephantCannonCheck()
        Console.WriteLine()

        ' ── 테스트 3: InjectGameHistory로 반복 감지 ──
        Test3_InjectAndDetect()
        Console.WriteLine()

        ' ── 테스트 4: Search에서 반복 수 회피 ──
        Test4_SearchAvoidsRepetition()
    End Sub

    Sub Test1_IsRepetition()
        Console.WriteLine("── 테스트 1: IsRepetition 기본 동작 ──")

        Dim grid = EmptyGrid()
        grid(1)(4) = HK : grid(0)(3) = HS : grid(0)(5) = HS
        grid(8)(4) = CK : grid(9)(3) = CS : grid(9)(5) = CS
        grid(5)(4) = CP : grid(5)(6) = CE
        grid(3)(0) = HB : grid(6)(0) = CJ

        Dim board As New Board(grid)
        Dim h0 = board.ZobristHash
        Console.WriteLine($"  초기 해시: {h0}")

        ' CE (5,6) → (2,4): 상이 뜨면서 포장
        board.MakeMove((5, 6), (2, 4))
        Dim h1 = board.ZobristHash
        Console.WriteLine($"  CE(5,6)→(2,4) 후 해시: {h1}")
        Console.WriteLine($"  포장(HK 장군)? {board.IsInCheck(HAN)}")

        ' HK (1,4) → (0,4): 왕 피함
        board.MakeMove((1, 4), (0, 4))
        Dim h2 = board.ZobristHash

        ' CE (2,4) → (5,6): 상 복귀
        board.MakeMove((2, 4), (5, 6))
        Dim h3 = board.ZobristHash

        ' HK (0,4) → (1,4): 왕 복귀
        board.MakeMove((0, 4), (1, 4))
        Dim h4 = board.ZobristHash

        Console.WriteLine($"  4수 후 해시: {h4} (초기={h0})")
        Console.WriteLine($"  동일 국면 복귀? {h0 = h4}")

        ' 이제 CE(5,6)→(2,4) 다시 하면 h1과 동일
        board.MakeMove((5, 6), (2, 4))
        Dim h5 = board.ZobristHash
        Console.WriteLine($"  CE(5,6)→(2,4) 재시도 해시: {h5} (첫 포장={h1})")
        Console.WriteLine($"  동일 국면? {h1 = h5}")
        Console.WriteLine($"  IsRepetition(1)? {board.IsRepetition(1)}")  ' 히스토리에 h1이 있으므로 True

        Dim result = If(h0 = h4 AndAlso h1 = h5 AndAlso board.IsRepetition(1), "PASS ✓", "FAIL ✗")
        Console.WriteLine($"  결과: {result}")
        board.UndoMove() ' h5
    End Sub

    Sub Test2_ElephantCannonCheck()
        Console.WriteLine("── 테스트 2: 상 뜨면서 포장 확인 ──")

        Dim grid = EmptyGrid()
        ' 한(HAN) 진영
        grid(1)(4) = HK
        grid(0)(3) = HS : grid(0)(5) = HS
        grid(3)(0) = HB : grid(3)(8) = HB
        ' 초(CHO) 진영
        grid(8)(4) = CK
        grid(9)(3) = CS : grid(9)(5) = CS
        grid(5)(4) = CP   ' 포: 4열에서 왕 조준
        grid(5)(6) = CE   ' 상: (5,6)에서 (2,4)로 이동 가능
        grid(6)(0) = CJ : grid(6)(8) = CJ

        Dim board As New Board(grid)

        ' 상 이동 전: 포장 아님 (CP와 HK 사이에 screen 없음)
        Console.WriteLine($"  상 이동 전 HK 장군? {board.IsInCheck(HAN)} (기대: False)")

        ' CE (5,6) → (2,4) 이동 (상 뜨면서 포장)
        ' 경유지: (4,5), (3,4) 모두 빈칸
        board.MakeMove((5, 6), (2, 4))

        ' 이제 CP(5,4) → screen CE(2,4) → HK(1,4) 포장!
        Dim isCheck = board.IsInCheck(HAN)
        Console.WriteLine($"  CE(5,6)→(2,4) 후 HK 장군? {isCheck} (기대: True)")

        Dim result = If(Not board.IsInCheck(HAN) = False AndAlso isCheck, "PASS ✓", "FAIL ✗")
        Console.WriteLine($"  결과: {result}")

        board.UndoMove()
    End Sub

    Sub Test3_InjectAndDetect()
        Console.WriteLine("── 테스트 3: InjectGameHistory 반복 감지 ──")

        Dim grid = EmptyGrid()
        grid(1)(4) = HK : grid(0)(3) = HS : grid(0)(5) = HS
        grid(8)(4) = CK : grid(9)(3) = CS : grid(9)(5) = CS
        grid(5)(4) = CP : grid(5)(6) = CE
        grid(3)(0) = HB : grid(6)(0) = CJ

        Dim board As New Board(grid)

        ' 게임 히스토리 시뮬레이션: 외부에서 해시 기록
        Dim gameHistory As New List(Of ULong)
        gameHistory.Add(board.ZobristHash) ' 초기 국면

        ' 1수: CE(5,6)→(2,4) 포장
        board.MakeMove((5, 6), (2, 4))
        gameHistory.Add(board.ZobristHash) ' 포장 국면

        ' 2수: HK(1,4)→(0,4) 피함
        board.MakeMove((1, 4), (0, 4))
        gameHistory.Add(board.ZobristHash)

        ' 3수: CE(2,4)→(5,6) 복귀
        board.MakeMove((2, 4), (5, 6))
        gameHistory.Add(board.ZobristHash)

        ' 4수: HK(0,4)→(1,4) 복귀
        board.MakeMove((0, 4), (1, 4))
        ' 이 시점에서 초기 국면과 동일

        Console.WriteLine($"  4수 후 = 초기 국면? {board.ZobristHash = gameHistory(0)}")

        ' 새 보드에 게임 히스토리 주입 (MacroRunner처럼)
        Dim board2 As New Board(board.Grid) ' 현재 상태로 새 보드
        board2.InjectGameHistory(gameHistory)

        ' CE(5,6)→(2,4) 다시 시도
        board2.MakeMove((5, 6), (2, 4))
        Dim isRep = board2.IsRepetition(1)
        Console.WriteLine($"  InjectGameHistory 후 CE(5,6)→(2,4) 반복? {isRep} (기대: True)")

        Dim result = If(isRep, "PASS ✓", "FAIL ✗")
        Console.WriteLine($"  결과: {result}")
    End Sub

    Sub Test4_SearchAvoidsRepetition()
        Console.WriteLine("── 테스트 4: Search에서 상 포장 반복 회피 ──")

        ' 보드 설정: CE가 포장을 만들 수 있지만, 포가 왕을 직접 잡을 수 없는 위치
        ' CP(7,4)는 HK(1,4)까지 거리가 멀고 사이에 screen이 없어서 직접 공격 불가
        ' CE(5,6) → (2,4) 이동 시 CE가 screen이 되어 포장 발생
        Dim grid = EmptyGrid()
        ' 한 진영: 왕은 궁성 안, 사로 보호
        grid(1)(4) = HK
        grid(0)(3) = HS : grid(0)(5) = HS
        grid(3)(0) = HB : grid(3)(8) = HB
        grid(0)(0) = HC  ' 한 차: 초의 공격을 어렵게 함
        ' 초 진영
        grid(8)(4) = CK
        grid(9)(3) = CS : grid(9)(5) = CS
        grid(5)(4) = CP   ' 포: 4열, 왕과 같은 열 (screen 없으면 공격 불가)
        grid(5)(6) = CE   ' 상: (5,6) ↔ (2,4) 왕복 가능
        grid(6)(0) = CJ : grid(6)(8) = CJ

        ' ── A: 히스토리 없이 탐색 ──
        Dim boardA As New Board(grid)
        Console.WriteLine($"  [히스토리 없음] CHO 최선수 탐색 (depth=5)...")
        Dim resultA = Search.FindBestMove(boardA, CHO, 5, 10)
        If resultA.BestMove.HasValue Then
            Dim m = resultA.BestMove.Value
            Dim piece = boardA.Grid(m.Item1.Item1)(m.Item1.Item2)
            Dim pName = GetPieceName(piece)
            Console.WriteLine($"  최선수: {pName} ({m.Item1.Item1},{m.Item1.Item2})→({m.Item2.Item1},{m.Item2.Item2}) 점수={resultA.Score} 깊이={resultA.Depth}")
            Dim isElephantCheck = (m.Item1.Item1 = 5 AndAlso m.Item1.Item2 = 6 AndAlso m.Item2.Item1 = 2 AndAlso m.Item2.Item2 = 4)
            Console.WriteLine($"  CE(5,6)→(2,4) 포장? {isElephantCheck}")
        Else
            Console.WriteLine($"  최선수 없음!")
        End If

        ' ── B: 포장 반복 히스토리 주입 후 탐색 ──
        Console.WriteLine()
        Console.WriteLine($"  [포장 1회 반복 히스토리 주입]")

        ' 히스토리 구성: 초기→포장→피함→복귀→복귀 (1사이클 완료)
        Dim boardSim As New Board(grid)
        Dim gameHistory As New List(Of ULong)
        gameHistory.Add(boardSim.ZobristHash)

        boardSim.MakeMove((5, 6), (2, 4))  ' CE 포장
        gameHistory.Add(boardSim.ZobristHash)

        boardSim.MakeMove((1, 4), (0, 4))  ' HK 피함
        gameHistory.Add(boardSim.ZobristHash)

        boardSim.MakeMove((2, 4), (5, 6))  ' CE 복귀
        gameHistory.Add(boardSim.ZobristHash)

        boardSim.MakeMove((0, 4), (1, 4))  ' HK 복귀
        gameHistory.Add(boardSim.ZobristHash)

        Console.WriteLine($"  히스토리 크기: {gameHistory.Count}")
        Console.WriteLine($"  초기 국면 = 현재? {gameHistory(0) = boardSim.ZobristHash}")

        ' 새 보드에 히스토리 주입
        Dim boardB As New Board(grid)
        boardB.InjectGameHistory(gameHistory)

        Console.WriteLine($"  CHO 최선수 탐색 (depth=5, 반복 히스토리 주입)...")
        Dim resultB = Search.FindBestMove(boardB, CHO, 5, 10)
        If resultB.BestMove.HasValue Then
            Dim m = resultB.BestMove.Value
            Dim piece = boardB.Grid(m.Item1.Item1)(m.Item1.Item2)
            Dim pName = GetPieceName(piece)
            Console.WriteLine($"  최선수: {pName} ({m.Item1.Item1},{m.Item1.Item2})→({m.Item2.Item1},{m.Item2.Item2}) 점수={resultB.Score} 깊이={resultB.Depth}")
            Dim isElephantCheck = (m.Item1.Item1 = 5 AndAlso m.Item1.Item2 = 6 AndAlso m.Item2.Item1 = 2 AndAlso m.Item2.Item2 = 4)
            Console.WriteLine($"  CE(5,6)→(2,4) 포장? {isElephantCheck}")

            ' 수동 반복 확인
            boardB.MakeMove(m.Item1, m.Item2)
            Console.WriteLine($"  선택한 수 후 IsRepetition(1)? {boardB.IsRepetition(1)}")

            If isElephantCheck Then
                Console.WriteLine($"  결과: FAIL ✗ (반복 수를 선택함)")
            Else
                Console.WriteLine($"  결과: PASS ✓ (반복 수를 회피함)")
            End If
        Else
            Console.WriteLine($"  최선수 없음!")
            Console.WriteLine($"  결과: FAIL ✗")
        End If

        ' ── C: 직접 반복 감지 확인 (포장 수 수동 실행) ──
        Console.WriteLine()
        Console.WriteLine($"  [직접 반복 감지 확인]")
        Dim boardC As New Board(grid)
        boardC.InjectGameHistory(gameHistory)
        boardC.MakeMove((5, 6), (2, 4))  ' CE 포장 수 수동 실행
        Console.WriteLine($"  CE(5,6)→(2,4) 수동 실행 후:")
        Console.WriteLine($"    IsRepetition(1)? {boardC.IsRepetition(1)} (기대: True)")
        Console.WriteLine($"    IsInCheck(HAN)? {boardC.IsInCheck(HAN)} (기대: True, 포장)")
    End Sub

End Module
