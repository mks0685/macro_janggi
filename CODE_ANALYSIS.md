# 카카오 장기 AI 매크로 - 소스 분석

**VB.NET / .NET 8.0 WinForms** 애플리케이션으로, 카카오 장기 게임 화면을 캡처하여 보드 상태를 자동 인식하고, AI로 최적수를 계산한 뒤 마우스 클릭으로 자동 착수하는 매크로 프로그램.

---

## 아키텍처 (15개 소스 파일, ~4,500 LOC)

```
MacroAutoControl/
├── Program.vb              - 진입점 (--diag 모드 지원)
├── MainForm.vb             - UI 메인폼 (~1,926줄, 가장 큼)
├── MacroRunner.vb          - 매크로 실행 엔진 (~1,082줄)
├── MacroItem.vb            - 매크로 항목 데이터 클래스
├── WindowFinder.vb         - Win32 API로 창 탐색/캡처 (5가지 캡처 방법)
├── ButtonFinder.vb         - 이미지 템플릿 매칭 (픽셀 비교 방식)
├── ButtonClicker.vb        - 마우스 클릭 (포그라운드/백그라운드)
├── NativeMethods.vb        - Win32 P/Invoke 선언
├── GlowDiag.vb             - 빛남 감지 진단 도구
│
├── AI/
│   ├── Constants.vb        - 보드 크기, 기물 코드, 궁성, 기물 가치
│   ├── Board.vb            - 보드 표현, 수 실행/되돌리기, Zobrist 해싱
│   ├── Pieces.vb           - 기물별 이동 규칙 (궁/사/차/포/마/상/졸)
│   ├── Evaluator.vb        - 형세 판단 (위치 보너스, 기동력, 궁 안전도 등)
│   ├── Search.vb           - NegaMax + Alpha-Beta 탐색 엔진
│   └── BoardRecognizer.vb  - OpenCV 기반 보드 인식 (템플릿 매칭)
│
└── templates/              - 기물 이미지 템플릿 14개
                              (CC, CE, CJ, CK, CM, CP, CS, HB, HC, HE, HK, HM, HP, HS)
```

---

## 핵심 모듈별 분석

### 1. 화면 캡처 & 보드 인식

#### WindowFinder.vb
- **창 탐색**: `EnumWindows` Win32 API로 제목 키워드("장기", "Janggi", "KakaoGame", "한게임") 포함 창 탐색
- **창 캡처**: 5단계 폴백 전략
  1. `PrintWindow(PW_RENDERFULLCONTENT)` - 최신 렌더링 포함
  2. `PrintWindow(기본)` - 표준 캡처
  3. `PrintWindow(PW_CLIENTONLY)` - 클라이언트 영역만
  4. `BitBlt(DC 복사)` - GDI 기반 캡처
  5. `CopyFromScreen` - 화면 직접 캡처 (최후 수단, 창 전면 필요)
- **검은 이미지 감지**: 샘플링으로 빠르게 확인, 검은 이미지면 다음 방법 시도
- **DPI 인식**: `SetProcessDPIAware()` 호출

#### BoardRecognizer.vb (~1,107줄)
- **보드 영역 감지**: 이진화 → 모폴로지(Close/Open) → 윤곽선 검출 → 최대 영역 선택
- **격자선 감지**: Sobel 에지 검출(수직/수평) → 피크 검출 → 등간격 피팅(10행 x 9열)
- **기물 인식**: OpenCvSharp `MatchTemplate(CCoeffNormed)` 사용
  - 1단계: 원본 스케일 컬러 매칭 (임계값 0.7)
  - 2단계: 4가지 스케일(0.88, 0.94, 1.06, 1.12) 컬러 매칭 (임계값 0.65)
  - 3단계: 5가지 스케일 그레이스케일 매칭 (임계값 0.55)
  - 스케일별 템플릿 사전 캐싱으로 반복 리사이즈 방지
- **진영 판별**: HSV 색상 분석
  - 빨강(H:0~12 또는 158~180, S>40, V>60) → 한(HAN)
  - 파랑(H:85~135, S>40, V>60) → 초(CHO)
  - 기물 중앙 30% 영역만 사용하여 배경 간섭 최소화
- **진영 보정**: 템플릿 매칭 결과와 색상 분석 결과가 불일치하면 교차 진영으로 교정
- **궁성 후처리**: 사/궁은 궁성 밖에 있을 수 없으므로, 궁성 밖 인식 결과는 빈 칸으로 처리
- **보드 정규화**: 초궁이 상단(행 0~2)에 있으면 보드 상하 반전 → 항상 "초=하단(행7~9), 한=상단(행0~2)"
- **유효성 검증**: 궁 각 1개, 기물 유형별 최대 수 확인
- **하이라이트 감지**: HSV(H:15~30, S>50, V>220) 금색 글로우 픽셀을 링 영역(innerR~outerR)에서 2px 간격 샘플링
  - 기물 크기별 링 크기 조절 (졸/사: 축소, 궁: 확대, 나머지: 기본)
  - 최소 100 픽셀 이상이어야 하이라이트로 판정
  - 가장 강한 글로우 기물 → 마지막 착수 기물

#### ButtonFinder.vb
- **템플릿 매칭**: BitmapData 직접 접근으로 고속 픽셀 비교
  - 샘플링 간격(stepSize)으로 전수 비교 대신 ~500 픽셀만 비교
  - 스캔 간격(scanStep=2)으로 탐색 속도 향상
  - 조기 종료: 매칭률이 임계값×0.7 미만이면 스킵
  - RGB 각 채널 차이 30 이내이면 매칭으로 판정
- **색상 범위 탐색**: 특정 색상의 연속 영역을 찾아 버튼 위치 추정

---

### 2. AI 엔진

#### Constants.vb
- 보드 크기: 10행 × 9열
- 기물 코드: 진영 접두사(C/H) + 기물 코드(K/S/C/M/E/P/J/B)
- 기물 가치 (중반): 차 1300, 포 800, 마 450, 상 350, 사 300, 졸 200
- 기물 가치 (종반): 차 1400, 포 600, 마 480, 상 300, 사 280, 졸 280
- 궁성 좌표: 초(행7~9, 열3~5), 한(행0~2, 열3~5)
- 궁성 대각선 이동 테이블, 궁성 대각선 라인(차 슬라이딩용)
- 기본 탐색 깊이: 7, 최대 탐색 시간: 30초

#### Board.vb (~1,070줄)
- **보드 표현**: `String()()` 2D 배열 (각 칸 = 기물 코드 또는 ".")
- **Zobrist 해싱**: 14개 기물 × 90칸 랜덤 테이블 + 수번(side) 해시
  - MakeMove/UndoMove 시 점진적 해시 업데이트
  - 해시 히스토리 저장으로 UndoMove 시 정확한 복원
- **기물 리스트**: 기물 코드 → 위치 리스트 딕셔너리로 빠른 기물 위치 조회
  - MakeMove/UndoMove 시 점진 업데이트
- **전진 방향**: 왕 위치 기반 동적 결정 (행 5 이상이면 위로, 미만이면 아래로)
- **장군 감지 (`IsInCheck`)**: 역방향 공격 탐지
  - 차: 직선 4방향 슬라이딩
  - 포: 직선 4방향 + 포대(1개 기물 넘기)
  - 마: 8방향 역방향 오프셋 테이블 + 경유지 빈칸 체크
  - 상: 8방향 역방향 오프셋 테이블 + 2개 경유지 빈칸 체크
  - 졸/병: 전진 역방향 + 좌우
  - 궁성 대각선: 졸/사/차의 궁성 대각 공격
  - 궁성 대각선 포: 코너→중앙 포대→대각 반대 코너
- **합법수 생성 (`GetLegalMoves`)**: 모든 수 생성 → MakeMove → 장군 체크 → UndoMove
- **빅장 감지 (`IsBikjang`)**: 두 궁이 같은 열에 사이 기물 없이 마주보는 상태
- **반복 수 감지 (`IsRepetition`)**: 해시 히스토리에서 동일 해시 등장 횟수 확인
- **Static Exchange Evaluation (SEE)**:
  - 교환 시 최소 가치 공격자 순서로 교대 캡처 시뮬레이션
  - NegaMax 방식 역산으로 최적 교환 가치 계산
  - 기물 가치 순서: 궁(1) < 졸(200) < 사(300) < 상(350) < 마(450) < 포(800) < 차(1300)

#### Pieces.vb
- **궁/사**: 궁성 내 직선 + 대각선 이동 (PALACE_DIAG_MOVES 테이블 사용)
- **차**: 직선 4방향 무한 슬라이딩 + 궁성 대각선 슬라이딩 (PALACE_DIAG_LINES)
- **포**: 직선 4방향 포대 넘기 + 궁성 대각선 포격
  - 포는 포를 포대로 사용 불가, 포로 포를 잡을 수 없음
  - 궁성 대각선: 코너에서 중앙 포대 넘어 반대 코너로 이동/캡처
- **마**: 日자 이동 (직선 1칸 → 대각 1칸), 경유지 빈칸 필수
- **상**: 用자 이동 (직선 1칸 → 대각 2칸), 2개 경유지 빈칸 필수
- **졸/병**: 전진 + 좌우 + 궁성 대각선 (전진 또는 수평 방향만)

#### Evaluator.vb
- **위치 보너스 테이블** (초 기준, 한은 행 반전):
  - 졸: 전진할수록 + 중앙 선호 (최대 50)
  - 마: 중앙 + 적진 선호 (최대 50)
  - 포: 중앙 열 + 적진 방향 (최대 35)
  - 차: 적진 침투 강력 보너스, 자진 후방 패널티 (범위 -20~80)
  - 상: 중앙 배치 선호 (최대 20)
- **차 열린 줄 보너스**: 같은 열에 아군/적 졸 없으면 +30/+15
- **포 포대 보너스**: 4방향에 포대 역할 기물 수에 따라 -30~+30
- **차-포 합동 공격 시너지**: 같은 행/열에 차+포 → +30~40, 쌍차 같은 열 → +25
- **궁성 압박**: 적 궁성 근처 차/포 배치 보너스, 차+포 동시 궁 행/열 → +40
- **위험 기물 패널티**: 적 졸에 의해 공격받는 차/포/마 가치의 1/4 감산
- **궁 안전도**: 인접 아군 보호(+15), 장군(-50), 이동 제한(-40), 적 차 위협(-25), 궁성 중앙(+20), 사 보호(+20)
- **기동력**: 차/포/마 이동 가능 수 × 3

#### Search.vb
- **Iterative Deepening**: 깊이 1부터 최대 깊이까지 점진 탐색
- **Aspiration Windows**: 이전 깊이 점수 ± 50 → 실패 시 ± 200 → 풀 윈도우
- **NegaMax + Alpha-Beta**: 기본 탐색 프레임워크
- **Transposition Table (TT)**:
  - Dictionary 기반, 최대 50만 엔트리 (초과 시 전체 클리어)
  - EXACT/ALPHA/BETA 3가지 플래그
  - TT 최선수를 수 정렬에서 최우선 사용
- **수 정렬 우선순위**:
  1. TT 최선수 (-1,000,000)
  2. 캡처 수 (피해자 가치 × 10 - 공격자 가치)
  3. Killer Move (-5,000)
  4. History Heuristic (depth² 누적)
- **Null Move Pruning**: 깊이 3+, 비장군 상태에서 R=2 감소 탐색
- **Late Move Reduction (LMR)**: 깊이 3+, 4번째 수부터, 비캡처/비킬러/비TT수에 적용
  - R=1 (기본), R=2 (8번째 수 이후)
- **Futility Pruning**: 깊이 1~2, 정적 평가 + 마진 ≤ alpha이면 비캡처/비장군 수 스킵
  - 마진: 깊이 1 = 200, 깊이 2 = 450
- **Reverse Futility Pruning**: 깊이 1~2, 정적 평가 - 마진 ≥ beta이면 조기 컷
- **Principal Variation Search (PVS)**: 첫 수 풀 윈도우, 이후 제로 윈도우 → 실패 시 재탐색
- **Check Extension**: 장군 상태면 깊이 +1
- **Quiescence Search**:
  - 최대 6수 깊이
  - 캡처 수만 생성, MVV-LVA 정렬
  - Delta Pruning: standPat + 피해자 가치 + 200 < alpha이면 스킵
- **시간 관리**: DateTime 기반 데드라인, 외부 취소 요청 지원
- **외통수 감지**: |score| > INF - 100이면 즉시 반환

---

### 3. 매크로 실행 엔진 (MacroRunner.vb)

**실행 흐름**:
```
RunSequence (무한/유한 반복)
  ├── 이미지 항목: 캡처 → 템플릿 매칭 → 클릭 → 대기
  ├── 키전송 항목: PostMessage/SendKeys → 대기
  └── AI 항목: ExecuteAIItem
       └── While 루프:
            ├── WaitForMyTurn (글로우 폴링, 1초 간격)
            │    ├── 캡처 → 보드 인식 → 진영 자동 감지
            │    ├── 보드 프리뷰 표시 (기물 원/글로우 값/장군 감지)
            │    ├── 게임 결과 팝업 감지 (승/패/무효 등)
            │    └── 금색 글로우 감지 → 상대 기물이 빛나면 내 차례
            ├── 1.5초 대기 (상대 애니메이션 완료)
            ├── 재캡처 + 재인식 (0.5초 후 한번 더 확인)
            ├── AI 탐색 (FindBestMove)
            ├── 시각화 이벤트 (출발/도착 화살표)
            ├── 출발지 클릭 → 300ms → 도착지 클릭
            └── 2초 대기 → 다시 WaitForMyTurn
```

**주요 기능**:
- **내 차례 감지**: 금색 글로우가 상대 기물에 있으면 내 차례, 내 기물에 있으면 상대 차례
- **진영 자동 감지 (AUTO)**: 보드 반전 여부 + 하단 궁 색상으로 판별
- **게임 결과 팝업 감지**: 7가지 결과 템플릿 매칭 (무효/승리/패배/시간승패/기권승패)
- **연속 실패 감지**: 보드 인식 10회 연속 실패 시 중단
- **스크린샷 자동 저장**: `D:\images\macro_janggi\watch_yyyy-MM-dd\NNN\NNNN.jpg`
- **매크로 파일 직렬화**: `.macro` 텍스트 파일 + 동명 폴더(이미지 템플릿)
  - AI 항목: `AI|이름|대기|임계값|버튼|키|진영|깊이|시간`
  - 일반 항목: `이미지파일|이름|대기|임계값|버튼|키|클릭X|클릭Y`
  - 첫 줄: `WINDOW|창 이름`

---

### 4. UI (MainForm.vb)

**레이아웃**:
- **좌측**: 캡처 미리보기 (PictureBox, Dock=Fill)
  - 드래그로 템플릿 영역 선택 (점선 사각형 + 실시간 크기 표시)
  - 클릭으로 템플릿 내 클릭 위치 지정 (십자 + 원 표시)
  - 드래그&드롭으로 이미지 파일 로드 → 자동 보드 인식
- **우측 패널** (380px, 스크롤 가능):
  1. 대상 선택: 창 목록 + 모니터 목록 (더블클릭으로 선택/캡처)
  2. 매크로 리스트: 템플릿 미리보기, 옵션(대기/임계/버튼), 추가/삭제/정렬, AI 항목(대기/깊이/시간), 이미지 테스트/AI 테스트, 저장/불러오기
  3. 실행: 반복 횟수(무한 지원), 실행/중지 버튼, 진행 상태 표시

**글로벌 키보드 훅**:
- ESC: 매크로 중지
- +/-: AI 탐색 깊이 실시간 조절 (실행 중)

**윈도우 설정 저장**: 위치/크기/상태를 `window_settings.txt`에 저장, 복원 시 화면 내 확인

---

### 5. 입력 처리

#### ButtonClicker.vb
두 가지 클릭 모드 지원:

| 모드 | 방식 | 특징 |
|------|------|------|
| **포그라운드** | `SetCursorPos` + `mouse_event` | 창을 전면으로 가져온 후 실제 마우스 이동 |
| **백그라운드** | `PostMessage` (WM_LBUTTONDOWN/UP) | 창이 뒤에 있어도 클릭 가능 |

- 프레임 오프셋 보정: `GetWindowRect` → `ClientToScreen` → 타이틀바/테두리 크기 계산

#### NativeMethods.vb
사용되는 Win32 API:

| 카테고리 | API |
|----------|-----|
| 윈도우 열거 | `EnumWindows`, `GetWindowText`, `IsWindowVisible` |
| 윈도우 정보 | `GetWindowRect`, `GetClientRect`, `ClientToScreen` |
| 윈도우 제어 | `SetForegroundWindow`, `ShowWindow`, `IsIconic` |
| 마우스 입력 | `SetCursorPos`, `GetCursorPos`, `mouse_event` |
| 키보드 입력 | `keybd_event` |
| 메시지 전송 | `SendMessage`, `PostMessage` |
| 화면 캡처 | `PrintWindow` |
| 키보드 훅 | `SetWindowsHookEx`, `UnhookWindowsHookEx`, `CallNextHookEx` |

---

## 데이터 흐름

```
[카카오 장기 창]
    ↓ PrintWindow/BitBlt/CopyFromScreen
[스크린샷 (Bitmap)]
    ↓ OpenCvSharp BitmapConverter.ToMat
[BoardRecognizer.GetBoardState()]
    ├── DetectBoardRect()      → 장기판 영역 감지
    ├── FindGridLines()        → Sobel 에지 → 격자선 검출
    ├── FitUniformGrid()       → 10행 × 9열 등간격 피팅
    ├── RecognizePiece() × 90  → 기물 인식 (3단계 템플릿 매칭)
    ├── ValidatePalacePieces() → 궁성 내 진영 보정
    └── FlipGrid()             → 보드 정규화 (초=하단)
         ↓
[Board 객체 (10×9 그리드)]
         ↓
[Search.FindBestMove()]
    ├── Iterative Deepening (1 ~ maxDepth)
    ├── NegaMax + Alpha-Beta + PVS
    ├── TT / Killer / History / Null Move / LMR / Futility
    ├── Quiescence Search (캡처 수 + Delta Pruning)
    └── Evaluate() → 형세 점수
         ↓
[최적수 (fromRow,fromCol) → (toRow,toCol)]
    ↓ TranslateRow (보드 반전 보정)
    ↓ gridPositions[idx] → 화면 픽셀 좌표
[ButtonClicker.ClickInWindow()]
    ├── 출발지 클릭 (300ms 대기)
    └── 도착지 클릭
```

---

## 기술 스택

| 항목 | 기술 |
|------|------|
| 언어 | VB.NET (.NET 8.0) |
| UI | Windows Forms |
| 이미지 처리 | OpenCvSharp4 (4.9.0) |
| 화면 캡처 | Win32 API (PrintWindow, BitBlt, CopyFromScreen) |
| 마우스/키보드 | Win32 API (mouse_event, PostMessage, keybd_event) |
| 키보드 훅 | Low-Level Keyboard Hook (WH_KEYBOARD_LL) |

---

## 설정 및 파일

| 파일 | 용도 |
|------|------|
| `window_settings.txt` | 폼 위치/크기/상태 저장 |
| `last_macro.txt` | 마지막 사용 매크로 파일 경로 |
| `*.macro` | 매크로 시퀀스 정의 (파이프 구분자) |
| `templates/*.png` | 기물 인식용 템플릿 이미지 14개 |
| `templates/result_*.png` | 게임 결과 팝업 인식용 7개 |
| `templates/btn_*.png` | UI 버튼 인식용 |

---

## 특기 사항

- **순수 VB.NET 구현**: 외부 장기 엔진 없이 완전한 장기 AI를 자체 구현
- **Zobrist 해싱**: 보드 상태를 64비트 해시로 관리, TT 및 반복 수 감지에 활용
- **기물 리스트 점진 업데이트**: MakeMove/UndoMove 시 기물 위치 리스트를 효율적으로 유지
- **멀티스케일 템플릿 매칭**: 해상도 차이에 대응하는 5단계 스케일 + 컬러/그레이 이중 매칭
- **금색 글로우 감지**: 카카오장기 특유의 하이라이트 효과를 HSV 색상 분석으로 정밀 검출하여 차례 판정
- **5단계 캡처 폴백**: 다양한 윈도우 렌더링 방식에 대응
- **SEE (Static Exchange Evaluation)**: 캡처 교환의 최종 이득/손해를 정확히 계산
