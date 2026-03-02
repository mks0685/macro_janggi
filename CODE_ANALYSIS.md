# MacroAutoControl - 카카오 장기 AI 매크로 코드 분석

## 1. 프로젝트 개요

카카오 장기 게임을 자동으로 플레이하는 Windows 데스크톱 애플리케이션. 화면 캡처 기반 이미지 매칭으로 장기판을 인식하고, 내장 AI 엔진이 최적의 수를 계산하여 자동으로 착수한다.

| 항목 | 내용 |
|------|------|
| **언어** | VB.NET (.NET 8.0) |
| **UI** | Windows Forms |
| **이미지 처리** | OpenCvSharp 4.9.0 |
| **프로젝트 타입** | WinExe (독립 실행형) |
| **네임스페이스** | `MacroAutoControl` |

---

## 2. 프로젝트 구조

```
kakao_JangGi/
├── Program.vb                 # 엔트리 포인트
├── MainForm.vb                # 메인 UI (1,518줄)
├── MacroRunner.vb             # 매크로 실행 엔진 (743줄)
├── MacroItem.vb               # 매크로 항목 데이터 클래스
├── ButtonClicker.vb           # 마우스 클릭 처리
├── ButtonFinder.vb            # 이미지 템플릿 매칭
├── WindowFinder.vb            # 윈도우 탐색/캡처
├── NativeMethods.vb           # Win32 API 선언
├── MacroAutoControl.vbproj    # 프로젝트 파일
├── AI/
│   ├── Constants.vb           # 장기 상수 (기물, 보드 크기, 궁성 좌표)
│   ├── Pieces.vb              # 기물별 이동 규칙
│   ├── Board.vb               # 보드 상태 관리 + Zobrist 해싱
│   ├── Search.vb              # NegaMax + Alpha-Beta 탐색 엔진
│   ├── Evaluator.vb           # 형세 평가 함수
│   └── BoardRecognizer.vb     # 보드 이미지 인식 (OpenCV)
├── templates/                 # 기물 템플릿 이미지 (14개)
│   ├── CK.png ~ CJ.png       # 초 기물 (궁/사/차/마/상/포/졸)
│   └── HK.png ~ HB.png       # 한 기물 (궁/사/차/마/상/포/병)
└── bin/Debug/net8.0-windows/
    └── templates/             # 빌드 출력 (기물 + 결과 팝업 템플릿)
        ├── btn_match.png      # 대국신청 버튼
        ├── btn_refresh.png    # 새로고침 버튼
        └── result_*.png       # 게임 결과 팝업 (승/패/시간승/기권 등)
```

---

## 3. 핵심 모듈 상세 분석

### 3.1 MainForm.vb - 메인 UI

메인 폼은 **좌/우 분할 레이아웃**으로 구성된다.

- **왼쪽**: 캡처 미리보기 (`PictureBox`, Dock=Fill)
- **오른쪽**: 제어 패널 (380px, 스크롤 가능)

#### 우측 패널 구성
1. **대상 선택** - 창 목록 / 모니터 목록 (더블클릭으로 캡처)
2. **매크로 리스트** - 템플릿 미리보기, 클릭/키전송/AI 항목 추가, 저장/불러오기
3. **실행** - 반복 횟수, 무한 반복, 실행/중지 버튼

#### 주요 기능
- **드래그 → 템플릿 선택**: 캡처 이미지에서 마우스 드래그로 영역 선택 → 템플릿 이미지 생성
- **클릭 위치 지정**: 템플릿 선택 후 클릭으로 세밀한 클릭 지점 설정
- **AI 테스트**: 현재 화면에서 보드 인식 → AI 탐색 → 결과 시각화 (출발지 파란 원, 도착지 빨간 원, 화살표)
- **장기 창 자동 탐색**: 제목에 "장기"가 포함된 창을 자동 선택
- **윈도우 설정 저장/복원**: 폼 위치/크기를 `window_settings.txt`에 저장

---

### 3.2 MacroRunner.vb - 매크로 실행 엔진

매크로 시퀀스를 순차 실행하는 핵심 엔진.

#### 실행 흐름
```
RunSequence() → 항목별 반복 {
   ├── AI 항목  → ExecuteAIItem()
   │     ├── WaitForMyTurn() (기물 빛남 감지로 내 차례 대기)
   │     ├── 캡처 → 보드 인식 (최대 10회 재시도)
   │     ├── 진영 자동 감지 (아래쪽 궁 색상)
   │     ├── AI 탐색 (Search.FindBestMove)
   │     ├── 시각화 이벤트 발생
   │     └── 출발지 클릭 → 300ms 대기 → 도착지 클릭
   ├── 이미지 매칭+클릭 항목 → 캡처 → FindByTemplate → 클릭
   └── 키전송 항목 → SendKeysToWindow()
}
```

#### 내 차례 감지 (`HasGlowAroundMyPieces`)
- 내 기물 외곽(셀 크기의 45% 반경)을 **8방향 샘플링**
- 빈 칸 평균 밝기 + 40 이상이면 "빛남"으로 판단
- 샘플의 절반 이상이 밝으면 해당 기물에 빛남 효과 있음 → 내 차례

#### 게임 종료 감지 (`DetectGameResult`)
- 결과 팝업 템플릿 7종 매칭: 무효대국, 승리, 패배, 시간승, 시간패, 기권승, 기권패

#### 매크로 저장 형식 (`.macro` 파일)
```
# 일반 항목
이미지파일|이름|대기ms|임계값|버튼|키|클릭X|클릭Y

# AI 항목
AI|이름|대기ms|임계값|버튼|키|진영|깊이|시간
```
- 이미지는 동명 폴더에 PNG로 저장

---

### 3.3 AI 엔진 (AI/ 폴더)

#### 3.3.1 Constants.vb - 장기 상수

| 분류 | 내용 |
|------|------|
| 보드 | 10행 x 9열 |
| 진영 | 초(C, 아래쪽/빨강), 한(H, 위쪽/파랑) |
| 기물 코드 | `CK`(초궁), `CS`(초사), `CC`(초차), `CM`(초마), `CE`(초상), `CP`(초포), `CJ`(초졸), `HK`~`HB`(한) |
| 기물 가치 | 차:1300, 포:800, 마:450, 상:350, 사:300, 졸/병:200, 궁:0 |
| 궁성 | 초: (7~9, 3~5), 한: (0~2, 3~5) |
| 궁성 대각선 | 꼭짓점 ↔ 중앙 이동 허용 |

#### 3.3.2 Pieces.vb - 기물 이동 규칙

각 기물별 합법적 이동 위치 계산:

| 기물 | 이동 방식 |
|------|-----------|
| **궁/사** | 궁성 내 상하좌우 + 대각선(꼭짓점↔중앙) |
| **차** | 직선 슬라이딩 (상하좌우 무제한) + 궁성 대각선 슬라이딩 |
| **포** | 하나를 넘어서 이동 (포끼리는 넘지 못함) + 궁성 대각선 뛰기 |
| **마** | 직선1칸 + 대각선1칸 (길목 차단 적용) |
| **상** | 직선1칸 + 대각선2칸 (경유지 2곳 차단 적용) |
| **졸/병** | 앞/좌/우 1칸 + 궁성 내 대각선 이동 |

#### 3.3.3 Board.vb - 보드 상태 관리

- **Zobrist 해싱**: 14기물 x 90칸 랜덤 테이블 + 진영 전환 XOR
- **MakeMove/UndoMove**: O(1) 수 실행/복귀 (히스토리 스택)
- **장군 체크 (`IsInCheck`)**: 역방향 공격 탐지 (차/포/마/상/졸 각각 최적화)
  - 마: 8방향 역오프셋 + 경유지 확인
  - 상: 8방향 역오프셋 + 경유지 2곳 확인
  - 포: 궁성 대각선 포함 판정
- **합법수 생성**: 의사합법수 생성 → MakeMove → 자가 장군 필터링
- **빅장 판정 (`IsBikjang`)**: 두 궁이 같은 열에서 중간에 기물 없이 마주보는지 확인

#### 3.3.4 Search.vb - 탐색 엔진

**NegaMax + Alpha-Beta Pruning** 기반의 탐색 엔진.

##### 적용된 최적화 기법
| 기법 | 설명 |
|------|------|
| **Iterative Deepening** | 깊이 1부터 점진적 심화 |
| **Aspiration Window** | 이전 깊이 점수 ±50으로 윈도우 설정, 실패 시 4배 확장 후 전체 탐색 |
| **Transposition Table** | Zobrist 해시 기반 50만 엔트리 (EXACT/ALPHA/BETA 플래그) |
| **Null Move Pruning** | 깊이 ≥ 3, 장군 아닐 때 R=2로 축소 탐색 |
| **Futility Pruning** | 깊이 ≤ 2에서 정적 평가 + 마진이 beta 이상이면 컷 |
| **Late Move Reduction** | 깊이 ≥ 3, 4번째 수부터 축소 탐색 (R=1~2) |
| **Principal Variation Search** | 첫 수만 full-window, 나머지는 zero-window → 재탐색 |
| **Killer Move Heuristic** | 플라이별 2개의 킬러 무브 저장 |
| **History Heuristic** | 성공한 비캡처 수에 depth² 가산 |
| **Quiescence Search** | 리프 노드에서 캡처 수만 최대 6겹 추가 탐색 |
| **Delta Pruning** | 정지 탐색 시 기물 가치 + 200 마진 미만이면 건너뜀 |
| **MVV-LVA 정렬** | 피해자 가치 x 10 - 공격자 가치로 캡처 수 정렬 |
| **Check Extension** | 장군 상태면 탐색 깊이 1 연장 |
| **시간 제한** | `DateTime.Now` 기반 데드라인 (기본 30초) |

#### 3.3.5 Evaluator.vb - 형세 평가 함수

**초 기준 양수, 한 기준 음수** 반환.

| 평가 요소 | 설명 |
|-----------|------|
| **기물 가치** | 차:1300, 포:800, 마:450, 상:350, 사:300, 졸/병:200 |
| **위치 보너스 (PST)** | 졸/마/포/차/상 각각 10x9 위치 테이블 (중앙/적진 진출 보너스) |
| **차 열린 줄** | 아군 졸 없는 열 +15, 양측 졸 없는 열 +30 |
| **포 포대 수** | 4방향 포대(넘을 기물) 개수에 따라 -30~+30 |
| **차+포 시너지** | 같은 행/열에 위치하면 추가 보너스 (+30~40) |
| **궁성 압박** | 적 궁성에 차/포 투입 시 보너스 (차+포 동시 +40) |
| **기동력** | 차/포/마의 가능한 수 개수 × 3 |
| **낙하 패널티** | 적 졸/병에 인접한 고가 기물 가치의 25% 차감 |
| **궁 안전도** | 인접 아군 +15, 장군 -50, 이동칸 ≤ 1이면 -40, 궁성 중앙 +20, 인접 사 +20 |

#### 3.3.6 BoardRecognizer.vb - 보드 이미지 인식

OpenCV 기반 장기판 인식 파이프라인:

```
1. DetectBoardRect()     → 윤곽선 기반 장기판 영역 감지
2. FindGridLines()       → Sobel 에지 → 프로파일 분석 → 피크 탐지
3. FitUniformGrid()      → 10행 × 9열 균등 그리드 피팅
4. GetCellImage()        → 교차점 중심으로 셀 이미지 추출
5. RecognizePiece()      → 3단계 템플릿 매칭
     a. 컬러 원본 매칭 (임계값 0.7)
     b. 다중 스케일 컬러 매칭 (0.88~1.12, 임계값 0.65)
     c. 그레이스케일 + HSV 색상 보정 매칭
6. ValidatePalacePieces() → 궁성 내 사/궁 진영 후처리 보정
```

##### 템플릿 매칭 전략
- 원본 크기 → 실패 시 4개 스케일(±6%, ±12%) 추가 시도
- 그레이스케일 매칭 시 HSV 색상 분석으로 초(빨강)/한(파랑) 진영 판별
- 궁성 밖의 사/궁 인식은 오류로 간주하여 빈 칸 처리

---

### 3.4 입력 처리

#### ButtonClicker.vb
두 가지 클릭 모드 지원:

| 모드 | 방식 | 특징 |
|------|------|------|
| **포그라운드** | `SetCursorPos` + `mouse_event` | 창을 전면으로 가져온 후 실제 마우스 이동 |
| **백그라운드** | `PostMessage` (WM_LBUTTONDOWN/UP) | 창이 뒤에 있어도 클릭 가능 |

- 프레임 오프셋 보정: `GetWindowRect` → `ClientToScreen` → 타이틀바/테두리 크기 계산

#### WindowFinder.vb
- `EnumWindows`로 모든 보이는 창 열거
- `PrintWindow` (PW_RENDERFULLCONTENT → 기본 → 클라이언트) → `BitBlt` → `CopyFromScreen` 순서로 캡처 시도
- `IsBlackImage`로 빈 캡처 필터링

#### ButtonFinder.vb
- **템플릿 매칭**: 샘플링 기반 고속 이미지 매칭 (√(총픽셀/500) 간격, 스캔 스텝 2)
- **색상 매칭**: 특정 색상 영역 탐지 (미사용 유틸리티)
- 조기 종료: 매칭률 < 임계값 × 0.7이면 스킵

---

### 3.5 NativeMethods.vb - Win32 API

사용되는 API 목록:

| 카테고리 | API |
|----------|-----|
| 윈도우 열거 | `EnumWindows`, `GetWindowText`, `IsWindowVisible` |
| 윈도우 정보 | `GetWindowRect`, `GetClientRect`, `ClientToScreen` |
| 윈도우 제어 | `SetForegroundWindow`, `ShowWindow`, `IsIconic` |
| 마우스 입력 | `SetCursorPos`, `GetCursorPos`, `mouse_event` |
| 키보드 입력 | `keybd_event` |
| 메시지 전송 | `SendMessage`, `PostMessage` |
| 화면 캡처 | `PrintWindow` |

---

## 4. 데이터 흐름

```
[카카오 장기 창]
    ↓ PrintWindow/BitBlt
[스크린샷 (Bitmap)]
    ↓ OpenCV Mat 변환
[BoardRecognizer.GetBoardState()]
    ├── DetectGrid() → 10x9 그리드 좌표
    └── RecognizePiece() × 90칸 → Board 객체
         ↓
[Search.FindBestMove()]
    ├── Iterative Deepening (1 ~ maxDepth)
    ├── NegaMax + Alpha-Beta
    ├── Quiescence Search (캡처 수)
    └── Evaluate() → 형세 점수
         ↓
[최적수 (fromRow,fromCol) → (toRow,toCol)]
    ↓ 그리드 좌표 → 화면 좌표 변환
[ButtonClicker.ClickInWindow()]
    ├── 출발지 클릭 (300ms 대기)
    └── 도착지 클릭
```

---

## 5. 설정 및 파일

| 파일 | 용도 |
|------|------|
| `window_settings.txt` | 폼 위치/크기/상태 저장 (X, Y, W, H, WindowState) |
| `last_macro.txt` | 마지막 사용 매크로 파일 경로 |
| `*.macro` | 매크로 시퀀스 정의 (파이프 구분자) |
| `templates/*.png` | 기물 인식용 템플릿 이미지 14개 |
| `templates/result_*.png` | 게임 결과 팝업 인식용 7개 |
| `templates/btn_*.png` | UI 버튼 인식용 |

---

## 6. 기술적 특징 요약

1. **완전 자립형 AI**: 외부 엔진 없이 VB.NET 내장 장기 AI 구현
2. **실시간 보드 인식**: OpenCV 기반 동적 장기판 감지 + 템플릿 매칭
3. **자동 차례 감지**: 기물 주변 빛남 효과 분석으로 내 차례 판별
4. **다중 캡처 전략**: PrintWindow → BitBlt → CopyFromScreen 순차 시도
5. **매크로 시스템**: 이미지 클릭 + 키전송 + AI 수를 조합한 시퀀스 매크로
6. **백그라운드 지원**: PostMessage 기반 백그라운드 클릭 가능
7. **게임 결과 자동 감지**: 7종 결과 팝업 템플릿 매칭으로 자동 종료
