' 장기 상수 정의
Namespace MacroAutoControl

    Public Module Constants
        ' 보드 크기
        Public Const BOARD_ROWS As Integer = 10
        Public Const BOARD_COLS As Integer = 9

        ' 진영 식별
        Public Const CHO As String = "C"  ' 초 (아래쪽, red)
        Public Const HAN As String = "H"  ' 한 (위쪽, green/blue)

        ' 빈 칸
        Public Const EMPTY As String = "."

        ' 기물 코드 (진영 접두사 + 기물 코드)
        ' 초 기물
        Public Const CK As String = "CK"  ' 궁 (King)
        Public Const CS As String = "CS"  ' 사 (Guard/Advisor)
        Public Const CC As String = "CC"  ' 차 (Chariot/Rook)
        Public Const CM As String = "CM"  ' 마 (Horse/Knight)
        Public Const CE As String = "CE"  ' 상 (Elephant/Bishop)
        Public Const CP As String = "CP"  ' 포 (Cannon)
        Public Const CJ As String = "CJ"  ' 졸 (Soldier/Pawn)

        ' 한 기물
        Public Const HK As String = "HK"  ' 궁
        Public Const HS As String = "HS"  ' 사
        Public Const HC As String = "HC"  ' 차
        Public Const HM As String = "HM"  ' 마
        Public Const HE As String = "HE"  ' 상
        Public Const HP As String = "HP"  ' 포
        Public Const HB As String = "HB"  ' 병 (Soldier/Pawn)

        ' 모든 기물 목록
        Public ReadOnly CHO_PIECES As New HashSet(Of String) From {CK, CS, CC, CM, CE, CP, CJ}
        Public ReadOnly HAN_PIECES As New HashSet(Of String) From {HK, HS, HC, HM, HE, HP, HB}

        ' 기물 가치 (중반)
        Public ReadOnly PIECE_VALUES As New Dictionary(Of String, Integer) From {
            {CK, 0}, {HK, 0},
            {CC, 1300}, {HC, 1300},
            {CP, 800}, {HP, 800},
            {CM, 450}, {HM, 450},
            {CE, 350}, {HE, 350},
            {CS, 300}, {HS, 300},
            {CJ, 200}, {HB, 200}
        }

        ' 기물 가치 (종반) - Phase 3A
        Public ReadOnly PIECE_VALUES_EG As New Dictionary(Of String, Integer) From {
            {CK, 0}, {HK, 0},
            {CC, 1400}, {HC, 1400},
            {CP, 600}, {HP, 600},
            {CM, 480}, {HM, 480},
            {CE, 300}, {HE, 300},
            {CS, 280}, {HS, 280},
            {CJ, 280}, {HB, 280}
        }

        ' Tapered Eval: 비졸 기물의 초기 총 가치 (한 진영 기준)
        ' 차2(2600) + 포2(1600) + 마2(900) + 상2(700) + 사2(600) = 6400
        Public Const TOTAL_PHASE As Integer = 6400

        ' 궁성 좌표 (row, col)
        ' 초 궁성: 행 7~9, 열 3~5
        Public ReadOnly CHO_PALACE As New HashSet(Of (Row As Integer, Col As Integer)) From {
            (7, 3), (7, 4), (7, 5),
            (8, 3), (8, 4), (8, 5),
            (9, 3), (9, 4), (9, 5)
        }
        ' 한 궁성: 행 0~2, 열 3~5
        Public ReadOnly HAN_PALACE As New HashSet(Of (Row As Integer, Col As Integer)) From {
            (0, 3), (0, 4), (0, 5),
            (1, 3), (1, 4), (1, 5),
            (2, 3), (2, 4), (2, 5)
        }

        ' 궁성 대각선 이동이 허용되는 위치와 방향
        Public ReadOnly PALACE_DIAG_MOVES As New Dictionary(Of (Integer, Integer), List(Of (Integer, Integer))) From {
            {(0, 3), New List(Of (Integer, Integer)) From {(1, 4)}},
            {(0, 5), New List(Of (Integer, Integer)) From {(1, 4)}},
            {(1, 4), New List(Of (Integer, Integer)) From {(0, 3), (0, 5), (2, 3), (2, 5)}},
            {(2, 3), New List(Of (Integer, Integer)) From {(1, 4)}},
            {(2, 5), New List(Of (Integer, Integer)) From {(1, 4)}},
            {(7, 3), New List(Of (Integer, Integer)) From {(8, 4)}},
            {(7, 5), New List(Of (Integer, Integer)) From {(8, 4)}},
            {(8, 4), New List(Of (Integer, Integer)) From {(7, 3), (7, 5), (9, 3), (9, 5)}},
            {(9, 3), New List(Of (Integer, Integer)) From {(8, 4)}},
            {(9, 5), New List(Of (Integer, Integer)) From {(8, 4)}}
        }

        ' 궁성 대각선 라인 (차 슬라이딩용)
        Public ReadOnly PALACE_DIAG_LINES As (Integer, Integer)()() = {
            New (Integer, Integer)() {(7, 3), (8, 4), (9, 5)},
            New (Integer, Integer)() {(7, 5), (8, 4), (9, 3)},
            New (Integer, Integer)() {(0, 3), (1, 4), (2, 5)},
            New (Integer, Integer)() {(0, 5), (1, 4), (2, 3)}
        }

        ' 초기 배치 (기본 상마상마 배치)
        Public Function GetInitialBoard() As String()()
            Dim board As String()() = New String(9)() {}
            board(0) = New String() {HC, HM, HE, HS, EMPTY, HS, HM, HE, HC}
            board(1) = New String() {EMPTY, EMPTY, EMPTY, EMPTY, HK, EMPTY, EMPTY, EMPTY, EMPTY}
            board(2) = New String() {EMPTY, HP, EMPTY, EMPTY, EMPTY, EMPTY, EMPTY, HP, EMPTY}
            board(3) = New String() {HB, EMPTY, HB, EMPTY, HB, EMPTY, HB, EMPTY, HB}
            board(4) = New String() {EMPTY, EMPTY, EMPTY, EMPTY, EMPTY, EMPTY, EMPTY, EMPTY, EMPTY}
            board(5) = New String() {EMPTY, EMPTY, EMPTY, EMPTY, EMPTY, EMPTY, EMPTY, EMPTY, EMPTY}
            board(6) = New String() {CJ, EMPTY, CJ, EMPTY, CJ, EMPTY, CJ, EMPTY, CJ}
            board(7) = New String() {EMPTY, CP, EMPTY, EMPTY, EMPTY, EMPTY, EMPTY, CP, EMPTY}
            board(8) = New String() {EMPTY, EMPTY, EMPTY, EMPTY, CK, EMPTY, EMPTY, EMPTY, EMPTY}
            board(9) = New String() {CC, CM, CE, CS, EMPTY, CS, CM, CE, CC}
            Return board
        End Function

        ' 탐색 관련
        Public Const DEFAULT_SEARCH_DEPTH As Integer = 7
        Public Const MAX_SEARCH_TIME As Double = 30.0
    End Module

End Namespace
