Imports System.Diagnostics
Imports System.Drawing
Imports System.Windows.Forms

Namespace MacroAutoControl
    Public Class MainForm
        Inherits Form

        ' UI 컨트롤 - 창 선택
        Private lstWindows As ListBox
        Private pnlRight As Panel
        Private grpWindows As GroupBox
        Private splitter As Splitter

        ' UI 컨트롤 - 매크로 리스트
        Private grpMacro As GroupBox
        Private lstMacro As ListBox
        Private picTemplate As PictureBox
        Private lblTemplate As Label
        Private WithEvents btnMacroAdd As Button
        Private WithEvents btnMacroDelete As Button
        Private WithEvents btnMacroUp As Button
        Private WithEvents btnMacroDown As Button
        Private WithEvents btnMacroSave As Button
        Private WithEvents btnMacroLoad As Button
        Private nudDelay As NumericUpDown
        Private nudThreshold As NumericUpDown
        Private cboMouseButton As ComboBox
        Private nudKeyDelay As NumericUpDown
        Private txtSendKeys As TextBox
        Private WithEvents btnKeyAdd As Button
        Private WithEvents chkBackground As CheckBox

        ' UI 컨트롤 - AI 항목
        Private cboAISide As ComboBox
        Private nudAIDelay As NumericUpDown
        Private nudAIDepth As NumericUpDown
        Private nudAITime As NumericUpDown
        Private WithEvents btnAIAdd As Button
        Private WithEvents btnAITest As Button

        ' AI 보드 인식기 (테스트용)
        Private _boardRecognizer As MacroAutoControl.Capture.BoardRecognizer = Nothing

        ' UI 컨트롤 - 모니터 선택
        Private lstMonitors As ListBox

        ' UI 컨트롤 - 실행
        Private grpActions As GroupBox
        Private WithEvents btnMacroRun As Button
        Private WithEvents btnMacroStop As Button
        Private lblMacroProgress As Label
        Private nudRepeatCount As NumericUpDown
        Private WithEvents chkInfinite As CheckBox

        ' UI 컨트롤 - 상태
        Private lblStatus As Label
        Private lblWindowInfo As Label
        Private picPreview As PictureBox

        ' 상태 변수
        Private _targetWindow As WindowFinder.WindowInfo
        Private _screenshot As Bitmap
        Private _templateImage As Bitmap
        Private _templateRect As Rectangle  ' 스크린샷 내 템플릿 영역
        Private _clickOffset As Point = New Point(-1, -1)  ' 템플릿 내 클릭 위치 (-1이면 중앙)

        ' 매크로 관련
        Private _macroItems As New List(Of MacroItem)
        Private _macroRunner As New MacroRunner()

        Private Const RIGHT_PANEL_WIDTH As Integer = 380
        Private ReadOnly _settingsPath As String = IO.Path.Combine(Application.StartupPath, "window_settings.txt")
        Private ReadOnly _lastMacroPath As String = IO.Path.Combine(Application.StartupPath, "last_macro.txt")
        Private _currentMacroFile As String = ""

        ' 윈도우 복원용 (Normal 상태의 위치/크기 저장)
        Private _lastNormalBounds As Rectangle

        ' 글로벌 키보드 훅 (ESC로 매크로 중지)
        Private _hookHandle As IntPtr = IntPtr.Zero
        Private _hookProc As NativeMethods.LowLevelKeyboardProc

        Public Sub New()
            InitializeComponent()
            AddHandler _macroRunner.ProgressChanged, AddressOf MacroRunner_ProgressChanged
            AddHandler _macroRunner.MacroCompleted, AddressOf MacroRunner_MacroCompleted
            AddHandler _macroRunner.AIMoveVisualize, AddressOf MacroRunner_AIMoveVisualize
            RestoreWindowSettings()
            InstallKeyboardHook()
        End Sub

        Private Sub InitializeComponent()
            Me.Text = "매크로 자동 제어"
            Me.WindowState = FormWindowState.Maximized
            Me.MinimumSize = New Size(900, 700)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.FormBorderStyle = FormBorderStyle.Sizable
            Me.KeyPreview = True

            ' ============================================================
            ' 오른쪽 패널 (스크롤 가능)
            ' ============================================================
            pnlRight = New Panel() With {
                .Dock = DockStyle.Right,
                .Width = RIGHT_PANEL_WIDTH,
                .Padding = New Padding(5),
                .BackColor = Color.FromArgb(240, 240, 240),
                .AutoScroll = True
            }
            Me.Controls.Add(pnlRight)

            splitter = New Splitter() With {
                .Dock = DockStyle.Right,
                .Width = 5,
                .BackColor = Color.FromArgb(200, 200, 200)
            }
            Me.Controls.Add(splitter)

            Dim pw = RIGHT_PANEL_WIDTH - 16
            Dim currentY = 5

            ' --- 1. 대상 선택 그룹 ---
            grpWindows = New GroupBox() With {
                .Text = "1. 대상 선택 (더블클릭으로 선택+캡처)",
                .Location = New Point(5, currentY),
                .Size = New Size(pw, 210),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            }
            pnlRight.Controls.Add(grpWindows)

            ' 창 목록
            Dim lblWindows As New Label() With {
                .Text = "창 목록:",
                .Location = New Point(10, 20),
                .Size = New Size(pw - 22, 16),
                .AutoSize = False
            }
            grpWindows.Controls.Add(lblWindows)

            lstWindows = New ListBox() With {
                .Location = New Point(10, 37),
                .Size = New Size(pw - 22, 90),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            }
            grpWindows.Controls.Add(lstWindows)
            AddHandler lstWindows.DoubleClick, AddressOf lstWindows_DoubleClick

            ' 모니터 목록
            Dim lblMonitors As New Label() With {
                .Text = "모니터 (더블클릭으로 전체화면 캡처):",
                .Location = New Point(10, 132),
                .Size = New Size(pw - 22, 16),
                .AutoSize = False
            }
            grpWindows.Controls.Add(lblMonitors)

            lstMonitors = New ListBox() With {
                .Location = New Point(10, 149),
                .Size = New Size(pw - 22, 52),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            }
            grpWindows.Controls.Add(lstMonitors)
            AddHandler lstMonitors.DoubleClick, AddressOf lstMonitors_DoubleClick

            currentY += grpWindows.Height + 5

            ' --- 2. 매크로 리스트 그룹 ---
            grpMacro = New GroupBox() With {
                .Text = "2. 매크로 리스트",
                .Location = New Point(5, currentY),
                .Size = New Size(pw, 564),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            }
            pnlRight.Controls.Add(grpMacro)

            Dim gy = 20

            ' 현재 드래그 템플릿 미리보기
            lblTemplate = New Label() With {
                .Text = "템플릿: 캡처 후 미리보기에서 드래그",
                .Location = New Point(10, gy),
                .Size = New Size(pw - 22, 18),
                .AutoSize = False
            }
            grpMacro.Controls.Add(lblTemplate)
            gy += 20

            picTemplate = New PictureBox() With {
                .Location = New Point(10, gy),
                .Size = New Size(pw - 22, 120),
                .BorderStyle = BorderStyle.FixedSingle,
                .SizeMode = PictureBoxSizeMode.Zoom,
                .BackColor = Color.White,
                .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            }
            grpMacro.Controls.Add(picTemplate)
            gy += 124

            ' 공통 레이아웃 상수 (pw=364 기준, 겹침 없이 배치)
            '  c1  c2        c3   c4       c5         btnX
            '  |대기(ms):|[  100]  |임계:|[75]  |[좌클릭]  |[추가 ▼]|
            Dim halfW = CInt((pw - 26) / 2)
            Dim rowH = 27       ' 행 높이
            Dim btnW = 65       ' 우측 버튼 너비
            Dim btnX = pw - 12 - btnW  ' 우측 버튼 X
            Dim c1 = 8          ' 첫번째 라벨 X
            Dim c2 = 62         ' 첫번째 입력 X
            Dim c3 = 130        ' 두번째 라벨 X
            Dim c4 = 162        ' 두번째 입력 X
            Dim c5 = 216        ' 세번째 시작 X

            ' === 옵션 행 1: 클릭 항목 추가 ===
            ' [대기(ms): 54px][4px][nud 64px][4px][임계: 28px][4px][nud 48px][4px][cbo 63px][4px][btn 65px]
            Dim lblDelay As New Label() With {
                .Text = "대기(ms):",
                .Location = New Point(c1, gy + 4),
                .Size = New Size(52, 18),
                .AutoSize = False
            }
            grpMacro.Controls.Add(lblDelay)

            nudDelay = New NumericUpDown() With {
                .Location = New Point(c2, gy),
                .Size = New Size(c3 - c2 - 4, 25),
                .Minimum = 0,
                .Maximum = 30000,
                .Value = 1000,
                .Increment = 100
            }
            grpMacro.Controls.Add(nudDelay)
            AddHandler nudDelay.ValueChanged, AddressOf nudDelay_ValueChanged

            Dim lblThresh As New Label() With {
                .Text = "임계:",
                .Location = New Point(c3, gy + 4),
                .Size = New Size(30, 18),
                .AutoSize = False
            }
            grpMacro.Controls.Add(lblThresh)

            nudThreshold = New NumericUpDown() With {
                .Location = New Point(c4, gy),
                .Size = New Size(c5 - c4 - 4, 25),
                .Minimum = 50,
                .Maximum = 100,
                .Value = 75,
                .Increment = 5,
                .DecimalPlaces = 0
            }
            grpMacro.Controls.Add(nudThreshold)

            cboMouseButton = New ComboBox() With {
                .Location = New Point(c5, gy),
                .Size = New Size(btnX - c5 - 4, 25),
                .DropDownStyle = ComboBoxStyle.DropDownList
            }
            cboMouseButton.Items.AddRange({"좌클릭", "우클릭"})
            cboMouseButton.SelectedIndex = 0
            grpMacro.Controls.Add(cboMouseButton)

            btnMacroAdd = New Button() With {
                .Text = "추가 ▼",
                .Location = New Point(btnX, gy),
                .Size = New Size(btnW, 25),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Right
            }
            grpMacro.Controls.Add(btnMacroAdd)
            gy += rowH

            ' === 옵션 행 2: 키전송 항목 추가 ===
            Dim lblKeyDelay As New Label() With {
                .Text = "대기(ms):",
                .Location = New Point(c1, gy + 4),
                .Size = New Size(52, 18),
                .AutoSize = False
            }
            grpMacro.Controls.Add(lblKeyDelay)

            nudKeyDelay = New NumericUpDown() With {
                .Location = New Point(c2, gy),
                .Size = New Size(c3 - c2 - 4, 25),
                .Minimum = 0,
                .Maximum = 30000,
                .Value = 1000,
                .Increment = 100
            }
            grpMacro.Controls.Add(nudKeyDelay)

            Dim lblKeys As New Label() With {
                .Text = "키:",
                .Location = New Point(c3, gy + 4),
                .Size = New Size(20, 18),
                .AutoSize = False
            }
            grpMacro.Controls.Add(lblKeys)

            txtSendKeys = New TextBox() With {
                .Location = New Point(c3 + 22, gy),
                .Size = New Size(btnX - (c3 + 22) - 4, 23),
                .PlaceholderText = "{ENTER}, abc, {F5}",
                .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            }
            grpMacro.Controls.Add(txtSendKeys)

            btnKeyAdd = New Button() With {
                .Text = "키추가 ▼",
                .Location = New Point(btnX, gy),
                .Size = New Size(btnW, 25),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Right
            }
            grpMacro.Controls.Add(btnKeyAdd)
            gy += rowH

            ' === 옵션 행 3: AI 항목 추가 (행1과 열 정렬) ===
            ' [대기(ms): c1~c2][nud c2~c3][깊이: c3~c4][nud c4~c5][시간: c5~][nud ~btnX][btn]
            Dim lblAIDelay As New Label() With {
                .Text = "대기(ms):", .Location = New Point(c1, gy + 4), .Size = New Size(52, 18), .AutoSize = False
            }
            grpMacro.Controls.Add(lblAIDelay)

            nudAIDelay = New NumericUpDown() With {
                .Location = New Point(c2, gy), .Size = New Size(c3 - c2 - 4, 25),
                .Minimum = 0, .Maximum = 10000, .Value = 1000, .Increment = 100
            }
            grpMacro.Controls.Add(nudAIDelay)

            Dim lblAIDepth As New Label() With {
                .Text = "깊이:", .Location = New Point(c3, gy + 4), .Size = New Size(30, 18), .AutoSize = False
            }
            grpMacro.Controls.Add(lblAIDepth)

            nudAIDepth = New NumericUpDown() With {
                .Location = New Point(c4, gy), .Size = New Size(c5 - c4 - 4, 25),
                .Minimum = 1, .Maximum = 10, .Value = 9, .Increment = 1
            }
            grpMacro.Controls.Add(nudAIDepth)

            Dim lblAITime As New Label() With {
                .Text = "시간:", .Location = New Point(c5, gy + 4), .Size = New Size(30, 18), .AutoSize = False
            }
            grpMacro.Controls.Add(lblAITime)

            nudAITime = New NumericUpDown() With {
                .Location = New Point(c5 + 32, gy), .Size = New Size(btnX - c5 - 32 - 4, 25),
                .Minimum = 1, .Maximum = 60, .Value = 10, .Increment = 5
            }
            grpMacro.Controls.Add(nudAITime)

            btnAIAdd = New Button() With {
                .Text = "AI추가 ▼",
                .Location = New Point(btnX, gy),
                .Size = New Size(btnW, 25),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Right
            }
            grpMacro.Controls.Add(btnAIAdd)
            gy += rowH

            ' 옵션 행 4: AI 테스트
            btnAITest = New Button() With {
                .Text = "AI 테스트 ▶",
                .Location = New Point(10, gy),
                .Size = New Size(pw - 22, 25),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right,
                .BackColor = Color.FromArgb(200, 220, 255),
                .Font = New Font(Me.Font, FontStyle.Bold)
            }
            grpMacro.Controls.Add(btnAITest)
            gy += rowH

            ' 매크로 항목 리스트
            lstMacro = New ListBox() With {
                .Location = New Point(10, gy),
                .Size = New Size(pw - 22, 130),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            }
            grpMacro.Controls.Add(lstMacro)
            AddHandler lstMacro.SelectedIndexChanged, AddressOf lstMacro_SelectedIndexChanged
            AddHandler lstMacro.DoubleClick, AddressOf lstMacro_DoubleClick
            gy += 134

            ' 삭제, 위로, 아래로
            Dim thirdW = CInt((pw - 30) / 3)

            btnMacroDelete = New Button() With {
                .Text = "삭제",
                .Location = New Point(10, gy),
                .Size = New Size(thirdW, 26),
                .Enabled = False
            }
            grpMacro.Controls.Add(btnMacroDelete)

            btnMacroUp = New Button() With {
                .Text = "▲ 위로",
                .Location = New Point(10 + thirdW + 4, gy),
                .Size = New Size(thirdW, 26),
                .Enabled = False
            }
            grpMacro.Controls.Add(btnMacroUp)

            btnMacroDown = New Button() With {
                .Text = "▼ 아래로",
                .Location = New Point(10 + (thirdW + 4) * 2, gy),
                .Size = New Size(thirdW, 26),
                .Enabled = False
            }
            grpMacro.Controls.Add(btnMacroDown)
            gy += 30

            ' 저장/불러오기
            btnMacroSave = New Button() With {
                .Text = "매크로 저장",
                .Location = New Point(10, gy),
                .Size = New Size(halfW, 28),
                .Enabled = False
            }
            grpMacro.Controls.Add(btnMacroSave)

            btnMacroLoad = New Button() With {
                .Text = "매크로 불러오기",
                .Location = New Point(10 + halfW + 4, gy),
                .Size = New Size(halfW, 28)
            }
            grpMacro.Controls.Add(btnMacroLoad)
            gy += 32

            ' 백그라운드 클릭 체크박스
            chkBackground = New CheckBox() With {
                .Text = "백그라운드 클릭 (PostMessage)",
                .Location = New Point(10, gy),
                .Size = New Size(pw - 22, 25),
                .Checked = False
            }
            grpMacro.Controls.Add(chkBackground)

            currentY += grpMacro.Height + 5

            ' --- 3. 실행 그룹 ---
            grpActions = New GroupBox() With {
                .Text = "3. 실행",
                .Location = New Point(5, currentY),
                .Size = New Size(pw, 125),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            }
            pnlRight.Controls.Add(grpActions)

            ' 반복 옵션
            Dim lblRepeat As New Label() With {
                .Text = "반복:",
                .Location = New Point(10, 24),
                .Size = New Size(35, 18),
                .AutoSize = False
            }
            grpActions.Controls.Add(lblRepeat)

            nudRepeatCount = New NumericUpDown() With {
                .Location = New Point(45, 21),
                .Size = New Size(55, 25),
                .Minimum = 1,
                .Maximum = 9999,
                .Value = 100,
                .Increment = 1
            }
            grpActions.Controls.Add(nudRepeatCount)

            Dim lblRepeatSuffix As New Label() With {
                .Text = "회",
                .Location = New Point(102, 24),
                .Size = New Size(18, 18),
                .AutoSize = False
            }
            grpActions.Controls.Add(lblRepeatSuffix)

            chkInfinite = New CheckBox() With {
                .Text = "무한 반복",
                .Location = New Point(125, 22),
                .Size = New Size(80, 22),
                .Checked = False
            }
            grpActions.Controls.Add(chkInfinite)

            ' 매크로 실행/중지
            Dim btnHalfW = CInt((pw - 26 - 210) / 2)
            btnMacroRun = New Button() With {
                .Text = "▶ 실행",
                .Location = New Point(210, 19),
                .Size = New Size(btnHalfW, 28),
                .Enabled = False,
                .BackColor = Color.FromArgb(255, 200, 200),
                .Font = New Font(Me.Font, FontStyle.Bold),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            }
            grpActions.Controls.Add(btnMacroRun)

            btnMacroStop = New Button() With {
                .Text = "■ 중지",
                .Location = New Point(210 + btnHalfW + 4, 19),
                .Size = New Size(btnHalfW, 28),
                .Enabled = False,
                .BackColor = Color.FromArgb(200, 200, 200),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Right
            }
            grpActions.Controls.Add(btnMacroStop)

            ' 매크로 진행 상태
            lblMacroProgress = New Label() With {
                .Text = "",
                .Location = New Point(10, 52),
                .Size = New Size(pw - 22, 65),
                .ForeColor = Color.DarkGreen,
                .Font = New Font("맑은 고딕", 9.0F, FontStyle.Bold),
                .AutoSize = False,
                .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            }
            grpActions.Controls.Add(lblMacroProgress)

            currentY += grpActions.Height + 5

            ' --- 상태 표시 ---
            lblWindowInfo = New Label() With {
                .Text = "",
                .Location = New Point(5, currentY),
                .Size = New Size(pw, 20),
                .ForeColor = Color.DarkBlue,
                .AutoSize = False,
                .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            }
            pnlRight.Controls.Add(lblWindowInfo)

            currentY += 22

            lblStatus = New Label() With {
                .Text = "상태: 대기 중",
                .Location = New Point(5, currentY),
                .Size = New Size(pw, 60),
                .AutoSize = False,
                .Font = New Font("맑은 고딕", 9.5F, FontStyle.Regular),
                .ForeColor = Color.FromArgb(30, 30, 30),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            }
            pnlRight.Controls.Add(lblStatus)

            ' ============================================================
            ' 왼쪽: 캡처 미리보기 (나머지 공간 전부)
            ' ============================================================
            picPreview = New PictureBox() With {
                .Dock = DockStyle.Fill,
                .BorderStyle = BorderStyle.FixedSingle,
                .SizeMode = PictureBoxSizeMode.Zoom,
                .BackColor = Color.FromArgb(35, 35, 38),
                .AllowDrop = True
            }
            AddHandler picPreview.MouseDown, AddressOf PicPreview_MouseDown
            AddHandler picPreview.MouseMove, AddressOf PicPreview_MouseMove
            AddHandler picPreview.MouseUp, AddressOf PicPreview_MouseUp
            AddHandler picPreview.Paint, AddressOf PicPreview_Paint
            AddHandler picPreview.DragEnter, AddressOf PicPreview_DragEnter
            AddHandler picPreview.DragDrop, AddressOf PicPreview_DragDrop
            Me.Controls.Add(picPreview)

            picPreview.BringToFront()

            ' 템플릿 자동 로드
            Dim defaultTemplatePath = IO.Path.Combine(Application.StartupPath, "button_template.png")
            If IO.File.Exists(defaultTemplatePath) Then
                LoadTemplate(defaultTemplatePath)
            End If
        End Sub

        ' =============================================
        ' 드래그로 템플릿 영역 선택
        ' =============================================
        Private _dragStart As Point
        Private _dragCurrent As Point
        Private _isDragging As Boolean

        Private Sub PicPreview_MouseDown(sender As Object, e As MouseEventArgs)
            If _screenshot Is Nothing Then Return
            _dragStart = e.Location
            _dragCurrent = e.Location
            _isDragging = True
        End Sub

        Private Sub PicPreview_MouseMove(sender As Object, e As MouseEventArgs)
            If Not _isDragging Then Return
            _dragCurrent = e.Location
            picPreview.Invalidate()
        End Sub

        Private Sub PicPreview_Paint(sender As Object, e As PaintEventArgs)
            If Not _isDragging Then Return

            Dim x = Math.Min(_dragStart.X, _dragCurrent.X)
            Dim y = Math.Min(_dragStart.Y, _dragCurrent.Y)
            Dim w = Math.Abs(_dragCurrent.X - _dragStart.X)
            Dim h = Math.Abs(_dragCurrent.Y - _dragStart.Y)

            If w < 2 OrElse h < 2 Then Return

            Dim rect As New Rectangle(x, y, w, h)

            ' 반투명 채우기
            Using brush As New SolidBrush(Color.FromArgb(40, 0, 120, 255))
                e.Graphics.FillRectangle(brush, rect)
            End Using

            ' 점선 테두리
            Using pen As New Pen(Color.FromArgb(200, 0, 120, 255), 2)
                pen.DashStyle = Drawing2D.DashStyle.Dash
                e.Graphics.DrawRectangle(pen, rect)
            End Using

            ' 크기 표시
            If _screenshot IsNot Nothing Then
                Dim imgRect = GetImageRect(picPreview, _screenshot)
                If imgRect.Width > 0 AndAlso imgRect.Height > 0 Then
                    Dim scaleX = CDbl(_screenshot.Width) / imgRect.Width
                    Dim scaleY = CDbl(_screenshot.Height) / imgRect.Height
                    Dim realW = CInt(w * scaleX)
                    Dim realH = CInt(h * scaleY)
                    Dim sizeText = $"{realW}x{realH}"
                    Using font As New Font("맑은 고딕", 9, FontStyle.Bold)
                        Dim textSize = e.Graphics.MeasureString(sizeText, font)
                        Dim tx = x + (w - textSize.Width) / 2
                        Dim ty = y + (h - textSize.Height) / 2
                        If ty < y Then ty = y + 2
                        e.Graphics.DrawString(sizeText, font, Brushes.White, tx + 1, ty + 1)
                        e.Graphics.DrawString(sizeText, font, Brushes.Blue, tx, ty)
                    End Using
                End If
            End If
        End Sub

        Private Sub PicPreview_MouseUp(sender As Object, e As MouseEventArgs)
            If Not _isDragging Then Return
            _isDragging = False

            If _screenshot Is Nothing OrElse picPreview.Image Is Nothing Then Return

            Try
                Dim imgRect = GetImageRect(picPreview, _screenshot)
                If imgRect.Width <= 0 OrElse imgRect.Height <= 0 Then Return

                Dim scaleX = CDbl(_screenshot.Width) / imgRect.Width
                Dim scaleY = CDbl(_screenshot.Height) / imgRect.Height

                Dim x1 = CInt((_dragStart.X - imgRect.X) * scaleX)
                Dim y1 = CInt((_dragStart.Y - imgRect.Y) * scaleY)
                Dim x2 = CInt((e.X - imgRect.X) * scaleX)
                Dim y2 = CInt((e.Y - imgRect.Y) * scaleY)

                Dim left = Math.Min(x1, x2)
                Dim top = Math.Min(y1, y2)
                Dim w = Math.Abs(x2 - x1)
                Dim h = Math.Abs(y2 - y1)

                ' 작은 클릭 = 클릭 위치 지정 (템플릿 선택된 상태에서)
                If w < 10 OrElse h < 5 Then
                    If _templateImage IsNot Nothing AndAlso _templateRect.Width > 0 Then
                        Dim clickImgX = CInt((_dragStart.X - imgRect.X) * scaleX)
                        Dim clickImgY = CInt((_dragStart.Y - imgRect.Y) * scaleY)
                        _clickOffset = New Point(clickImgX - _templateRect.X, clickImgY - _templateRect.Y)
                        ShowPreviewWithClick(_templateRect, _clickOffset)
                        lblTemplate.Text = $"템플릿: {_templateRect.Width}x{_templateRect.Height} px, 클릭:{_clickOffset.X},{_clickOffset.Y}"
                        UpdateStatus($"클릭 위치 설정: ({_clickOffset.X},{_clickOffset.Y}) - 템플릿 좌상단 기준")
                    End If
                    Return
                End If

                left = Math.Max(0, Math.Min(left, _screenshot.Width - 1))
                top = Math.Max(0, Math.Min(top, _screenshot.Height - 1))
                w = Math.Min(w, _screenshot.Width - left)
                h = Math.Min(h, _screenshot.Height - top)

                _templateImage?.Dispose()
                _templateImage = _screenshot.Clone(New Rectangle(left, top, w, h), _screenshot.PixelFormat)
                _templateRect = New Rectangle(left, top, w, h)
                _clickOffset = New Point(-1, -1)  ' 기본 중앙
                picTemplate.Image = _templateImage
                lblTemplate.Text = $"템플릿: {w}x{h} px (클릭으로 위치 지정 가능)"

                ShowPreviewWithHighlight(New Rectangle(left, top, w, h), Color.Blue)
                UpdateStatus("템플릿 선택 완료. 클릭으로 위치 지정, '추가'로 매크로에 등록.")
            Catch
            End Try
        End Sub

        Private Function GetImageRect(pic As PictureBox, img As Bitmap) As Rectangle
            Dim ratioX = CDbl(pic.Width) / img.Width
            Dim ratioY = CDbl(pic.Height) / img.Height
            Dim ratio = Math.Min(ratioX, ratioY)

            Dim newW = CInt(img.Width * ratio)
            Dim newH = CInt(img.Height * ratio)
            Dim offsetX = (pic.Width - newW) \ 2
            Dim offsetY = (pic.Height - newH) \ 2

            Return New Rectangle(offsetX, offsetY, newW, newH)
        End Function

        ' =============================================
        ' 매크로 리스트 관리
        ' =============================================
        Private Sub btnMacroAdd_Click(sender As Object, e As EventArgs) Handles btnMacroAdd.Click
            If _templateImage Is Nothing Then
                UpdateStatus("먼저 캡처 후 미리보기에서 드래그하여 템플릿을 선택하세요.")
                Return
            End If

            Dim name = InputBox("매크로 항목 이름을 입력하세요:", "매크로 추가", $"항목{_macroItems.Count + 1}")
            If String.IsNullOrWhiteSpace(name) Then Return

            Dim clickBtn = If(cboMouseButton.SelectedIndex = 1, ClickButton.Right, ClickButton.Left)
            Dim item As New MacroItem(
                name,
                _templateImage,
                CInt(nudDelay.Value),
                nudThreshold.Value / 100.0,
                clickBtn,
                "",
                _clickOffset.X,
                _clickOffset.Y)

            _macroItems.Add(item)
            RefreshMacroList()
            lstMacro.SelectedIndex = lstMacro.Items.Count - 1

            btnMacroRun.Enabled = (_macroItems.Count > 0 AndAlso _targetWindow IsNot Nothing)
            btnMacroSave.Enabled = (_macroItems.Count > 0)

            ' 원본 스크린샷으로 복원하여 추가 드래그 가능하게
            If _screenshot IsNot Nothing Then
                SetPreviewImage(_screenshot)
            End If

            UpdateStatus($"매크로 항목 추가: {name}. 계속 드래그하여 다음 항목을 추가하세요.")
        End Sub

        Private Sub btnKeyAdd_Click(sender As Object, e As EventArgs) Handles btnKeyAdd.Click
            Dim keys = txtSendKeys.Text.Trim()
            If String.IsNullOrEmpty(keys) Then
                UpdateStatus("전송할 키를 입력하세요. 예: {ENTER}, abc, {F5}")
                Return
            End If

            ' 키 전송 전용 항목: 1x1 투명 더미 이미지 사용
            Dim dummyImg As New Bitmap(1, 1)
            Dim item As New MacroItem(
                $"키전송: {keys}",
                dummyImg,
                CInt(nudKeyDelay.Value),
                0.0,
                ClickButton.Left,
                keys)
            item.Threshold = 0  ' 임계값 0 = 이미지 찾기 건너뜀
            dummyImg.Dispose()

            _macroItems.Add(item)
            RefreshMacroList()
            lstMacro.SelectedIndex = lstMacro.Items.Count - 1

            btnMacroRun.Enabled = (_macroItems.Count > 0 AndAlso _targetWindow IsNot Nothing)
            btnMacroSave.Enabled = (_macroItems.Count > 0)
            UpdateStatus($"키전송 항목 추가: {keys}")
        End Sub

        Private Sub btnAIAdd_Click(sender As Object, e As EventArgs) Handles btnAIAdd.Click
            Dim delay = CInt(nudAIDelay.Value)
            Dim depth = CInt(nudAIDepth.Value)
            Dim time = CDbl(nudAITime.Value)

            ' AI 항목 추가
            Dim name = "AI수두기 (자동)"
            Dim item = MacroItem.CreateAI(name, delay, "AUTO", depth, time)
            _macroItems.Add(item)
            RefreshMacroList()
            lstMacro.SelectedIndex = lstMacro.Items.Count - 1

            btnMacroRun.Enabled = (_macroItems.Count > 0 AndAlso _targetWindow IsNot Nothing)
            btnMacroSave.Enabled = (_macroItems.Count > 0)
            UpdateStatus($"AI 항목 추가: {name} (깊이:{depth}, 시간:{time:F0}s)")
        End Sub

        Private Sub btnAITest_Click(sender As Object, e As EventArgs) Handles btnAITest.Click
            If Not AutoSelectJanggiWindow() Then Return

            btnAITest.Enabled = False
            Try
                Dim bmp = WindowFinder.CaptureWindow(_targetWindow.Handle)
                If bmp Is Nothing Then
                    UpdateStatus("AI 테스트: 캡처 실패")
                    Return
                End If

                ' 스크린샷 업데이트
                picPreview.Image = Nothing
                _screenshot?.Dispose()
                _screenshot = bmp
                picPreview.Image = _screenshot

                RunAITestOnCurrentImage()
            Finally
                btnAITest.Enabled = True
            End Try
        End Sub

        ''' <summary>
        ''' 현재 _screenshot 이미지로 AI 테스트 실행
        ''' </summary>
        Private Sub RunAITestOnCurrentImage()
            If _screenshot Is Nothing Then
                UpdateStatus("AI 테스트: 이미지 없음")
                Return
            End If

            Dim mat As OpenCvSharp.Mat = Nothing
            Try
                ' Mat 변환 (BGRA → BGR)
                mat = OpenCvSharp.Extensions.BitmapConverter.ToMat(_screenshot)
                If mat.Channels() = 4 Then
                    Dim bgr As New OpenCvSharp.Mat()
                    OpenCvSharp.Cv2.CvtColor(mat, bgr, OpenCvSharp.ColorConversionCodes.BGRA2BGR)
                    mat.Dispose()
                    mat = bgr
                End If

                ' BoardRecognizer 초기화
                If _boardRecognizer Is Nothing Then
                    _boardRecognizer = New MacroAutoControl.Capture.BoardRecognizer()
                    Dim loaded = _boardRecognizer.LoadTemplates()
                    If loaded = 0 Then
                        UpdateStatus("AI 테스트: 템플릿 로드 실패 (templates 폴더 확인)")
                        Return
                    End If
                End If

                ' 보드 인식
                UpdateStatus("AI 테스트: 보드 인식 중...")
                Application.DoEvents()

                Dim board = _boardRecognizer.GetBoardState(mat)
                If board Is Nothing Then
                    UpdateStatus("AI 테스트: 보드 인식 실패 - 장기판을 찾을 수 없습니다")
                    Return
                End If

                ' AI 설정 - 내 진영 자동 감지
                ' 보드가 뒤집혔으면 원래 화면에서 한(HAN)이 아래에 있었으므로 내 진영 = 한
                Dim side As String = Nothing
                If _boardRecognizer.IsBoardFlipped Then
                    side = HAN
                Else
                    For r = 7 To 9
                        For c = 3 To 5
                            Dim piece = board.Grid(r)(c)
                            If piece = CK Then
                                side = CHO
                                Exit For
                            ElseIf piece = HK Then
                                side = HAN
                                Exit For
                            End If
                        Next
                        If side IsNot Nothing Then Exit For
                    Next
                End If
                If side Is Nothing Then
                    side = CHO
                End If
                Dim sideText = If(side = CHO, "초", "한")
                Dim depth = CInt(nudAIDepth.Value)
                Dim time = CDbl(nudAITime.Value)

                ' 내 기물 확인 표시
                Dim gridPos = _boardRecognizer.GetGridPositions()
                If gridPos IsNot Nothing Then
                    Dim preview = New Bitmap(_screenshot)
                    Using g = Graphics.FromImage(preview)
                        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                        Dim myPieces = If(side = CHO, CHO_PIECES, HAN_PIECES)
                        Dim enemyPieces = If(side = CHO, HAN_PIECES, CHO_PIECES)
                        Dim myCount = 0
                        Dim enemyCount = 0
                        Dim myPieceNames As New List(Of String)

                        For r = 0 To BOARD_ROWS - 1
                            For c = 0 To BOARD_COLS - 1
                                Dim piece = board.Grid(r)(c)
                                If piece = EMPTY Then Continue For
                                Dim screenR = _boardRecognizer.TranslateRow(r)
                                Dim idx = screenR * BOARD_COLS + c
                                Dim px = gridPos(idx)(0)
                                Dim py = gridPos(idx)(1)
                                Dim pName = ""
                                MacroAutoControl.Capture.BoardRecognizer.PIECE_NAMES.TryGetValue(piece, pName)

                                If myPieces.Contains(piece) Then
                                    myCount += 1
                                    If piece(1) <> "K"c Then myPieceNames.Add(pName)
                                    ' 내 기물: 파란 테두리
                                    Using pen As New Pen(Color.FromArgb(200, 0, 120, 255), 3)
                                        g.DrawEllipse(pen, px - 22, py - 22, 44, 44)
                                    End Using
                                ElseIf enemyPieces.Contains(piece) Then
                                    enemyCount += 1
                                    ' 적 기물: 빨간 테두리
                                    Using pen As New Pen(Color.FromArgb(150, 255, 60, 60), 2)
                                        g.DrawEllipse(pen, px - 22, py - 22, 44, 44)
                                    End Using
                                End If

                                ' 기물 이름 라벨
                                Dim shortName = If(pName.Length >= 2, pName.Substring(1), pName)
                                Using font As New Font("맑은 고딕", 8, FontStyle.Bold)
                                    Dim sz = g.MeasureString(shortName, font)
                                    Dim tx = px - sz.Width / 2
                                    Dim ty = py - 24 - sz.Height
                                    Using bgBrush As New SolidBrush(Color.FromArgb(180, 0, 0, 0))
                                        g.FillRectangle(bgBrush, tx - 1, ty - 1, sz.Width + 2, sz.Height + 1)
                                    End Using
                                    Using textBrush As New SolidBrush(Color.White)
                                        g.DrawString(shortName, font, textBrush, tx, ty)
                                    End Using
                                End Using
                            Next
                        Next

                        ' 상단에 기물 요약 표시
                        Dim summary = $"[{sideText}] 내 기물 {myCount}개: {String.Join(" ", myPieceNames)}  |  적 기물 {enemyCount}개"
                        Using font As New Font("맑은 고딕", 10, FontStyle.Bold)
                            Dim sz = g.MeasureString(summary, font)
                            Using bgBrush As New SolidBrush(Color.FromArgb(200, 0, 0, 0))
                                g.FillRectangle(bgBrush, 4, 4, sz.Width + 8, sz.Height + 4)
                            End Using
                            Using textBrush As New SolidBrush(Color.FromArgb(255, 100, 255, 100))
                                g.DrawString(summary, font, textBrush, 8, 6)
                            End Using
                        End Using
                    End Using

                    SetPreviewImage(preview)
                    UpdateStatus($"AI 테스트: {sideText} 기물 확인 완료 - 탐색 시작...")
                    Application.DoEvents()
                End If

                ' AI 탐색
                UpdateStatus($"AI 테스트: {sideText} 탐색 중 (깊이:{depth}, 시간:{time:F0}s)...")
                Application.DoEvents()

                Dim result = AI.Search.FindBestMove(board, side, depth, time)

                If result.BestMove.HasValue AndAlso gridPos IsNot Nothing Then
                    Dim bestMove = result.BestMove.Value
                    Dim fromRow = bestMove.Item1.Item1
                    Dim fromCol = bestMove.Item1.Item2
                    Dim toRow = bestMove.Item2.Item1
                    Dim toCol = bestMove.Item2.Item2

                    Dim actualFromRow = _boardRecognizer.TranslateRow(fromRow)
                    Dim actualToRow = _boardRecognizer.TranslateRow(toRow)
                    Dim fromIdx = actualFromRow * BOARD_COLS + fromCol
                    Dim toIdx = actualToRow * BOARD_COLS + toCol
                    Dim fromX = gridPos(fromIdx)(0)
                    Dim fromY = gridPos(fromIdx)(1)
                    Dim toX = gridPos(toIdx)(0)
                    Dim toY = gridPos(toIdx)(1)

                    ' 결과 시각화
                    Dim preview = New Bitmap(_screenshot)
                    Using g = Graphics.FromImage(preview)
                        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias

                        Using brush As New SolidBrush(Color.FromArgb(100, 0, 100, 255))
                            g.FillEllipse(brush, fromX - 24, fromY - 24, 48, 48)
                        End Using
                        Using pen As New Pen(Color.Blue, 3)
                            g.DrawEllipse(pen, fromX - 24, fromY - 24, 48, 48)
                        End Using

                        Using brush As New SolidBrush(Color.FromArgb(100, 255, 50, 0))
                            g.FillEllipse(brush, toX - 24, toY - 24, 48, 48)
                        End Using
                        Using pen As New Pen(Color.Red, 3)
                            g.DrawEllipse(pen, toX - 24, toY - 24, 48, 48)
                        End Using

                        Using pen As New Pen(Color.FromArgb(220, 255, 220, 0), 4)
                            pen.EndCap = Drawing2D.LineCap.ArrowAnchor
                            g.DrawLine(pen, fromX, fromY, toX, toY)
                        End Using
                    End Using

                    SetPreviewImage(preview)

                    Dim pieceName = ""
                    MacroAutoControl.Capture.BoardRecognizer.PIECE_NAMES.TryGetValue(board.Grid(fromRow)(fromCol), pieceName)
                    Dim capturedPiece = board.Grid(toRow)(toCol)
                    Dim captureInfo = ""
                    If capturedPiece <> EMPTY Then
                        Dim capName = ""
                        MacroAutoControl.Capture.BoardRecognizer.PIECE_NAMES.TryGetValue(capturedPiece, capName)
                        captureInfo = $" (잡기:{capName})"
                    End If

                    UpdateStatus($"AI 테스트: {sideText} {pieceName} ({fromRow},{fromCol})→({toRow},{toCol}){captureInfo} | 점수:{result.Score} 깊이:{result.Depth}")
                Else
                    UpdateStatus("AI 테스트: 최적수 없음 (게임 종료 또는 착수 불가)")
                End If
            Catch ex As Exception
                UpdateStatus($"AI 테스트 오류: {ex.Message}")
            Finally
                mat?.Dispose()
            End Try
        End Sub

        Private Sub btnMacroDelete_Click(sender As Object, e As EventArgs) Handles btnMacroDelete.Click
            Dim idx = lstMacro.SelectedIndex
            If idx < 0 OrElse idx >= _macroItems.Count Then Return

            If picTemplate.Image Is _macroItems(idx).Image Then picTemplate.Image = Nothing
            _macroItems(idx).Dispose()
            _macroItems.RemoveAt(idx)
            RefreshMacroList()

            If _macroItems.Count > 0 Then
                lstMacro.SelectedIndex = Math.Min(idx, _macroItems.Count - 1)
            End If

            btnMacroRun.Enabled = (_macroItems.Count > 0 AndAlso _targetWindow IsNot Nothing)
            btnMacroSave.Enabled = (_macroItems.Count > 0)
            UpdateStatus("매크로 항목 삭제됨")
        End Sub

        Private Sub btnMacroUp_Click(sender As Object, e As EventArgs) Handles btnMacroUp.Click
            Dim idx = lstMacro.SelectedIndex
            If idx <= 0 Then Return

            Dim temp = _macroItems(idx)
            _macroItems(idx) = _macroItems(idx - 1)
            _macroItems(idx - 1) = temp
            RefreshMacroList()
            lstMacro.SelectedIndex = idx - 1
        End Sub

        Private Sub btnMacroDown_Click(sender As Object, e As EventArgs) Handles btnMacroDown.Click
            Dim idx = lstMacro.SelectedIndex
            If idx < 0 OrElse idx >= _macroItems.Count - 1 Then Return

            Dim temp = _macroItems(idx)
            _macroItems(idx) = _macroItems(idx + 1)
            _macroItems(idx + 1) = temp
            RefreshMacroList()
            lstMacro.SelectedIndex = idx + 1
        End Sub

        Private _updatingFromSelection As Boolean = False

        Private Sub lstMacro_SelectedIndexChanged(sender As Object, e As EventArgs)
            Dim hasSelection = (lstMacro.SelectedIndex >= 0)
            btnMacroDelete.Enabled = hasSelection
            btnMacroUp.Enabled = (lstMacro.SelectedIndex > 0)
            btnMacroDown.Enabled = (hasSelection AndAlso lstMacro.SelectedIndex < _macroItems.Count - 1)

            ' 선택한 항목의 템플릿 미리보기 + 대기시간 표시
            If hasSelection AndAlso lstMacro.SelectedIndex < _macroItems.Count Then
                Dim item = _macroItems(lstMacro.SelectedIndex)

                ' 대기시간을 nudDelay에 반영
                _updatingFromSelection = True
                nudDelay.Value = Math.Max(nudDelay.Minimum, Math.Min(nudDelay.Maximum, item.DelayAfterClick))
                _updatingFromSelection = False

                Try
                    If item.IsAI Then
                        picTemplate.Image = Nothing
                        Dim sideText = If(item.AISide = "C", "초", "한")
                        lblTemplate.Text = $"[{lstMacro.SelectedIndex + 1}] AI: {sideText} 깊이:{item.AIDepth} 시간:{item.AITime:F0}s"
                    ElseIf item.Image IsNot Nothing Then
                        picTemplate.Image = item.Image
                        Dim clickInfo = If(item.ClickOffsetX >= 0, $", 클릭:{item.ClickOffsetX},{item.ClickOffsetY}", ", 클릭:중앙")
                        lblTemplate.Text = $"[{lstMacro.SelectedIndex + 1}] {item.Name} ({item.Image.Width}x{item.Image.Height}{clickInfo})"
                    Else
                        picTemplate.Image = Nothing
                        lblTemplate.Text = $"[{lstMacro.SelectedIndex + 1}] {item.Name}"
                    End If
                Catch
                    picTemplate.Image = Nothing
                    lblTemplate.Text = $"[{lstMacro.SelectedIndex + 1}] {item.Name}"
                End Try
            End If
        End Sub

        Private Sub nudDelay_ValueChanged(sender As Object, e As EventArgs)
            If _updatingFromSelection Then Return
            Dim idx = lstMacro.SelectedIndex
            If idx < 0 OrElse idx >= _macroItems.Count Then Return
            _macroItems(idx).DelayAfterClick = CInt(nudDelay.Value)
            ' 리스트 텍스트 갱신 (SelectedIndexChanged 재발 방지)
            _updatingFromSelection = True
            lstMacro.Items(idx) = $"{idx + 1}. {_macroItems(idx)}"
            _updatingFromSelection = False
        End Sub

        ''' <summary>
        ''' 장기 창 자동 탐색. 성공 시 True, 실패 시 False
        ''' </summary>
        Private Function AutoSelectJanggiWindow() As Boolean
            If _targetWindow IsNot Nothing Then Return True

            Dim windows = WindowFinder.GetAllVisibleWindows()
            Dim janggiWindow = windows.FirstOrDefault(Function(w) w.Title.Contains("장기"))
            If janggiWindow IsNot Nothing Then
                _targetWindow = janggiWindow
                lblWindowInfo.Text = $"선택: {janggiWindow.Title}"
                btnMacroRun.Enabled = (_macroItems.Count > 0)
                UpdateStatus($"장기 창 자동 선택: {janggiWindow.Title}")
                Application.DoEvents()
                Return True
            End If

            UpdateStatus("장기 창을 찾을 수 없습니다. 대상 창을 선택하세요.")
            Return False
        End Function

        Private Sub lstMacro_DoubleClick(sender As Object, e As EventArgs)
            Dim idx = lstMacro.SelectedIndex
            If idx < 0 OrElse idx >= _macroItems.Count Then Return

            If Not AutoSelectJanggiWindow() Then Return

            If _macroRunner.IsRunning Then Return

            Dim singleItem As New List(Of MacroItem) From {_macroItems(idx)}
            SetMacroRunningUI(True)
            _macroRunner.RunSequence(singleItem, _targetWindow, Me, chkBackground.Checked, 1)
        End Sub

        Private Sub RefreshMacroList()
            lstMacro.Items.Clear()
            For i = 0 To _macroItems.Count - 1
                lstMacro.Items.Add($"{i + 1}. {_macroItems(i)}")
            Next
        End Sub

        ' =============================================
        ' 매크로 저장/불러오기
        ' =============================================
        ''' <summary>선택된 항목에 nudDelay 값 강제 반영</summary>
        Private Sub CommitDelayToSelected()
            Dim idx = lstMacro.SelectedIndex
            If idx >= 0 AndAlso idx < _macroItems.Count Then
                _macroItems(idx).DelayAfterClick = CInt(nudDelay.Value)
            End If
        End Sub

        Private Sub btnMacroSave_Click(sender As Object, e As EventArgs) Handles btnMacroSave.Click
            If _macroItems.Count = 0 Then Return
            CommitDelayToSelected()
            RefreshMacroList()

            Using dlg As New SaveFileDialog()
                dlg.Title = "매크로 저장"
                dlg.Filter = "매크로 파일 (*.macro)|*.macro"
                dlg.DefaultExt = "macro"
                dlg.FileName = "매크로1"
                If Not String.IsNullOrEmpty(_currentMacroFile) Then
                    dlg.InitialDirectory = IO.Path.GetDirectoryName(_currentMacroFile)
                    dlg.FileName = IO.Path.GetFileNameWithoutExtension(_currentMacroFile)
                End If

                If dlg.ShowDialog() = DialogResult.OK Then
                    Try
                        Dim winTitle = If(_targetWindow IsNot Nothing, _targetWindow.Title, Nothing)
                        MacroRunner.SaveToFile(dlg.FileName, _macroItems, winTitle)
                        _currentMacroFile = dlg.FileName
                        SaveLastMacroPath()
                        Dim name = IO.Path.GetFileNameWithoutExtension(dlg.FileName)
                        UpdateStatus($"매크로 저장 완료: {name}.macro + {name}/ ({_macroItems.Count}개 항목)")
                    Catch ex As Exception
                        UpdateStatus($"매크로 저장 실패: {ex.Message}")
                    End Try
                End If
            End Using
        End Sub

        Private Sub btnMacroLoad_Click(sender As Object, e As EventArgs) Handles btnMacroLoad.Click
            Using dlg As New OpenFileDialog()
                dlg.Title = "매크로 불러오기"
                dlg.Filter = "매크로 파일 (*.macro)|*.macro"
                If Not String.IsNullOrEmpty(_currentMacroFile) Then
                    dlg.InitialDirectory = IO.Path.GetDirectoryName(_currentMacroFile)
                End If

                If dlg.ShowDialog() = DialogResult.OK Then
                    LoadMacroFromFile(dlg.FileName)
                End If
            End Using
        End Sub

        Private Sub LoadMacroFromFile(macroFilePath As String)
            Try
                ' picTemplate이 매크로 항목 이미지를 참조하고 있으면 해제
                picTemplate.Image = Nothing
                For Each item In _macroItems
                    item.Dispose()
                Next
                _macroItems.Clear()

                Dim savedWindowTitle As String = Nothing
                _macroItems = MacroRunner.LoadFromFile(macroFilePath, savedWindowTitle)
                RefreshMacroList()

                _currentMacroFile = macroFilePath
                SaveLastMacroPath()

                ' 저장된 창 이름으로 자동 선택
                If Not String.IsNullOrEmpty(savedWindowTitle) Then
                    Dim windows = WindowFinder.GetAllVisibleWindows()
                    Dim matched = windows.FirstOrDefault(Function(w) w.Title = savedWindowTitle)
                    If matched IsNot Nothing Then
                        SelectWindow(matched)
                    End If
                End If

                btnMacroRun.Enabled = (_macroItems.Count > 0 AndAlso _targetWindow IsNot Nothing)
                btnMacroSave.Enabled = (_macroItems.Count > 0)
                Dim name = IO.Path.GetFileNameWithoutExtension(macroFilePath)
                Dim winInfo = If(Not String.IsNullOrEmpty(savedWindowTitle) AndAlso _targetWindow IsNot Nothing, $", 창:{savedWindowTitle}", "")
                UpdateStatus($"매크로 로드 완료: {_macroItems.Count}개 항목 ({name}){winInfo}")
            Catch ex As Exception
                UpdateStatus($"매크로 로드 실패: {ex.Message}")
            End Try
        End Sub

        Private Sub SaveLastMacroPath()
            Try
                IO.File.WriteAllText(_lastMacroPath, _currentMacroFile)
            Catch
            End Try
        End Sub

        Private Sub LoadLastMacro()
            Try
                If Not IO.File.Exists(_lastMacroPath) Then Return
                Dim filePath = IO.File.ReadAllText(_lastMacroPath).Trim()
                If String.IsNullOrEmpty(filePath) OrElse Not IO.File.Exists(filePath) Then Return
                LoadMacroFromFile(filePath)
            Catch
            End Try
        End Sub

        ' =============================================
        ' 매크로 실행/중지
        ' =============================================
        Private Sub chkInfinite_CheckedChanged(sender As Object, e As EventArgs) Handles chkInfinite.CheckedChanged
            nudRepeatCount.Enabled = Not chkInfinite.Checked
        End Sub

        Private Sub btnMacroRun_Click(sender As Object, e As EventArgs) Handles btnMacroRun.Click
            If _macroItems.Count = 0 Then
                UpdateStatus("매크로 항목이 없습니다.")
                Return
            End If

            If Not AutoSelectJanggiWindow() Then Return

            ' UI 상태 변경
            SetMacroRunningUI(True)

            Dim repeatCount = If(chkInfinite.Checked, 0, CInt(nudRepeatCount.Value))
            _macroRunner.RunSequence(_macroItems, _targetWindow, Me, chkBackground.Checked, repeatCount)
        End Sub

        Private Sub btnMacroStop_Click(sender As Object, e As EventArgs) Handles btnMacroStop.Click
            _macroRunner.Cancel()
            UpdateStatus("매크로 중지 요청...")
        End Sub

        Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
            If e.KeyCode = Keys.Escape AndAlso _macroRunner.IsRunning Then
                _macroRunner.Cancel()
                UpdateStatus("ESC: 매크로 중지 요청...")
                e.Handled = True
            End If
            MyBase.OnKeyDown(e)
        End Sub

        Private Sub MacroRunner_ProgressChanged(currentIndex As Integer, totalCount As Integer, itemName As String, status As String, currentRepeat As Integer, totalRepeat As Integer)
            Dim repeatInfo = If(totalRepeat = 0, $"[반복 {currentRepeat}회/무한]", $"[반복 {currentRepeat}/{totalRepeat}회]")
            lblMacroProgress.Text = $"{repeatInfo} {currentIndex}/{totalCount}: {itemName} {status}"
            Application.DoEvents()
        End Sub

        Private Sub MacroRunner_MacroCompleted(success As Boolean, message As String)
            SetMacroRunningUI(False)
            lblMacroProgress.Text = ""
            UpdateStatus(message)
            ' 실행 중 변경된 항목(깊이 등) 리스트 표시 갱신
            RefreshMacroList()
        End Sub

        Private Sub MacroRunner_AIMoveVisualize(screenshot As Bitmap, fromX As Integer, fromY As Integer, toX As Integer, toY As Integer, moveInfo As String)
            If screenshot Is Nothing Then Return

            Try
                Dim preview = New Bitmap(screenshot)
                Using g = Graphics.FromImage(preview)
                    g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias

                    ' 출발지 (파란 원)
                    Using brush As New SolidBrush(Color.FromArgb(100, 0, 100, 255))
                        g.FillEllipse(brush, fromX - 24, fromY - 24, 48, 48)
                    End Using
                    Using pen As New Pen(Color.Blue, 3)
                        g.DrawEllipse(pen, fromX - 24, fromY - 24, 48, 48)
                    End Using

                    ' 도착지 (빨간 원)
                    Using brush As New SolidBrush(Color.FromArgb(100, 255, 50, 0))
                        g.FillEllipse(brush, toX - 24, toY - 24, 48, 48)
                    End Using
                    Using pen As New Pen(Color.Red, 3)
                        g.DrawEllipse(pen, toX - 24, toY - 24, 48, 48)
                    End Using

                    ' 화살표
                    Using pen As New Pen(Color.FromArgb(220, 255, 220, 0), 4)
                        pen.EndCap = Drawing2D.LineCap.ArrowAnchor
                        g.DrawLine(pen, fromX, fromY, toX, toY)
                    End Using
                End Using

                SetPreviewImage(preview)
                screenshot.Dispose()
                UpdateStatus($"AI: {moveInfo}")
            Catch ex As Exception
                screenshot?.Dispose()
            End Try
        End Sub

        Private Sub SetMacroRunningUI(running As Boolean)
            btnMacroRun.Enabled = Not running
            btnMacroStop.Enabled = running
            btnMacroAdd.Enabled = Not running
            btnMacroDelete.Enabled = Not running
            btnMacroUp.Enabled = Not running
            btnMacroDown.Enabled = Not running
            btnMacroSave.Enabled = Not running AndAlso _macroItems.Count > 0
            btnMacroLoad.Enabled = Not running
            grpWindows.Enabled = Not running
        End Sub

        ' =============================================
        ' 창 관련 이벤트
        ' =============================================
        Private Sub RefreshWindowList()
            lstWindows.Items.Clear()
            Dim windows = WindowFinder.GetAllVisibleWindows()

            For Each w In windows.OrderBy(Function(x) x.Title)
                If w.Title.Contains("설정") Then Continue For
                If w.Title.Contains("매크로 자동 제어") Then Continue For
                lstWindows.Items.Add($"[{w.Handle}] {w.Title} ({w.Bounds.Width}x{w.Bounds.Height})")
            Next

            UpdateStatus($"총 {lstWindows.Items.Count}개 창을 찾았습니다.")
        End Sub

        Private Sub lstWindows_DoubleClick(sender As Object, e As EventArgs)
            If lstWindows.SelectedItem Is Nothing Then Return

            Dim text = lstWindows.SelectedItem.ToString()
            Dim handleStr = text.Substring(1, text.IndexOf("]") - 1)
            Dim handle = New IntPtr(Long.Parse(handleStr))

            Dim windows = WindowFinder.GetAllVisibleWindows()
            Dim selected = windows.FirstOrDefault(Function(w) w.Handle = handle)

            If selected IsNot Nothing Then
                SelectWindow(selected)
            End If
        End Sub

        Private Sub SelectWindow(info As WindowFinder.WindowInfo)
            _targetWindow = info
            lblWindowInfo.Text = $"선택: {info.Title}"
            btnMacroRun.Enabled = (_macroItems.Count > 0)
            UpdateStatus($"창 선택: {info.Title} ({info.Bounds.Width}x{info.Bounds.Height})")

            CaptureTargetWindow()
        End Sub

        ' =============================================
        ' 캡처 이벤트
        ' =============================================
        Private Sub CaptureTargetWindow()
            If _targetWindow Is Nothing Then
                UpdateStatus("먼저 대상 창을 선택하세요.")
                Return
            End If

            Try
                picPreview.Image = Nothing
                _screenshot?.Dispose()
                _screenshot = WindowFinder.CaptureWindow(_targetWindow.Handle)

                If _screenshot Is Nothing Then
                    UpdateStatus("캡처 실패. 창을 다시 선택해보세요.")
                    Return
                End If

                picPreview.Image = _screenshot
                RefreshWindowList()
                UpdateStatus($"캡처 완료 ({_screenshot.Width}x{_screenshot.Height}). 드래그로 버튼 영역을 선택하세요.")
            Catch ex As Exception
                UpdateStatus($"캡처 오류: {ex.Message}")
            End Try
        End Sub

        Protected Overrides Sub OnLoad(e As EventArgs)
            MyBase.OnLoad(e)
            RefreshWindowList()
            RefreshMonitorList()
            LoadLastMacro()
        End Sub

        Private Sub SetPreviewImage(img As Bitmap)
            Dim old = picPreview.Image
            picPreview.Image = img
            If old IsNot Nothing AndAlso old IsNot _screenshot Then
                old.Dispose()
            End If
        End Sub

        Private Sub ShowPreviewWithHighlight(rect As Rectangle, color As Color)
            If _screenshot Is Nothing Then Return

            Dim preview = New Bitmap(_screenshot)
            Using g = Graphics.FromImage(preview)
                Using pen As New Pen(color, 3)
                    g.DrawRectangle(pen, rect)
                End Using
                Dim cx = rect.X + rect.Width \ 2
                Dim cy = rect.Y + rect.Height \ 2
                g.FillEllipse(Brushes.Red, cx - 5, cy - 5, 10, 10)
            End Using

            SetPreviewImage(preview)
        End Sub

        Private Sub ShowPreviewWithClick(rect As Rectangle, clickOff As Point)
            If _screenshot Is Nothing Then Return

            Dim preview = New Bitmap(_screenshot)
            Using g = Graphics.FromImage(preview)
                Using pen As New Pen(Color.Blue, 3)
                    g.DrawRectangle(pen, rect)
                End Using
                ' 클릭 위치 표시 (십자 + 원)
                Dim cx = rect.X + clickOff.X
                Dim cy = rect.Y + clickOff.Y
                g.FillEllipse(Brushes.Red, cx - 6, cy - 6, 12, 12)
                Using pen As New Pen(Color.Yellow, 2)
                    g.DrawLine(pen, cx - 10, cy, cx + 10, cy)
                    g.DrawLine(pen, cx, cy - 10, cx, cy + 10)
                End Using
            End Using

            SetPreviewImage(preview)
        End Sub

        ' =============================================
        ' 모니터 관련
        ' =============================================
        Private Sub RefreshMonitorList()
            lstMonitors.Items.Clear()
            For i = 0 To Screen.AllScreens.Length - 1
                Dim scr = Screen.AllScreens(i)
                Dim primary = If(scr.Primary, " [주]", "")
                lstMonitors.Items.Add($"모니터 {i + 1}{primary}: {scr.Bounds.Width}x{scr.Bounds.Height} ({scr.Bounds.X},{scr.Bounds.Y})")
            Next
        End Sub

        Private Sub lstMonitors_DoubleClick(sender As Object, e As EventArgs)
            If lstMonitors.SelectedIndex < 0 Then Return
            Dim scrIdx = lstMonitors.SelectedIndex
            If scrIdx >= Screen.AllScreens.Length Then Return

            Dim scr = Screen.AllScreens(scrIdx)

            Try
                ' 모니터 캡처(CopyFromScreen)는 프로그램이 같은 모니터에 있을 때만 숨김
                Dim myScreen = Screen.FromControl(Me)
                Dim needHide = (myScreen.DeviceName = scr.DeviceName)

                If needHide Then
                    Me.Hide()
                    Threading.Thread.Sleep(500)
                End If

                picPreview.Image = Nothing
                _screenshot?.Dispose()

                _screenshot = New Bitmap(scr.Bounds.Width, scr.Bounds.Height)
                Using g = Graphics.FromImage(_screenshot)
                    g.CopyFromScreen(scr.Bounds.Location, Point.Empty, scr.Bounds.Size)
                End Using

                If needHide Then
                    Me.Show()
                    Me.Activate()
                End If

                picPreview.Image = _screenshot
                RefreshWindowList()
                UpdateStatus($"모니터 {scrIdx + 1} 캡처 완료 ({scr.Bounds.Width}x{scr.Bounds.Height}). 드래그로 버튼 영역을 선택하세요.")
            Catch ex As Exception
                Me.Show()
                Me.Activate()
                UpdateStatus($"모니터 캡처 오류: {ex.Message}")
            End Try
        End Sub

        ' =============================================
        ' 드래그 앤 드롭 이미지 가져오기
        ' =============================================
        Private Sub PicPreview_DragEnter(sender As Object, e As DragEventArgs)
            If e.Data.GetDataPresent(DataFormats.FileDrop) Then
                e.Effect = DragDropEffects.Copy
            Else
                e.Effect = DragDropEffects.None
            End If
        End Sub

        Private Sub PicPreview_DragDrop(sender As Object, e As DragEventArgs)
            Try
                Dim files = CType(e.Data.GetData(DataFormats.FileDrop), String())
                If files Is Nothing OrElse files.Length = 0 Then Return

                Dim path = files(0)
                Dim ext = IO.Path.GetExtension(path).ToLower()
                If ext <> ".png" AndAlso ext <> ".jpg" AndAlso ext <> ".jpeg" AndAlso ext <> ".bmp" Then
                    UpdateStatus("지원하지 않는 파일 형식입니다. (png/jpg/bmp)")
                    Return
                End If

                Dim bmp As New Bitmap(path)
                picPreview.Image = Nothing
                _screenshot?.Dispose()
                _screenshot = bmp
                picPreview.Image = _screenshot

                UpdateStatus($"이미지 로드: {IO.Path.GetFileName(path)} ({bmp.Width}x{bmp.Height}) → AI 테스트 시작...")
                Application.DoEvents()

                ' AI 테스트 실행
                RunAITestOnCurrentImage()
            Catch ex As Exception
                UpdateStatus($"이미지 로드 오류: {ex.Message}")
            End Try
        End Sub

        ' =============================================
        ' 템플릿 불러오기 (단일 파일)
        ' =============================================
        Private Sub LoadTemplate(path As String)
            Try
                Dim bmp As New Bitmap(path)
                _templateImage?.Dispose()
                _templateImage = bmp
                picTemplate.Image = _templateImage
                lblTemplate.Text = $"템플릿: {_templateImage.Width}x{_templateImage.Height} px"
                UpdateStatus($"템플릿 로드: {IO.Path.GetFileName(path)}")
            Catch ex As Exception
                UpdateStatus($"템플릿 로드 실패: {ex.Message}")
            End Try
        End Sub

        Private Sub UpdateStatus(msg As String)
            lblStatus.Text = $"상태: {msg}"
        End Sub

        ' =============================================
        ' 글로벌 ESC 키보드 훅
        ' =============================================
        Private Sub InstallKeyboardHook()
            _hookProc = New NativeMethods.LowLevelKeyboardProc(AddressOf KeyboardHookCallback)
            Using curProcess = Process.GetCurrentProcess()
                Using curModule = curProcess.MainModule
                    _hookHandle = NativeMethods.SetWindowsHookEx(
                        NativeMethods.WH_KEYBOARD_LL,
                        _hookProc,
                        NativeMethods.GetModuleHandle(curModule.ModuleName),
                        0)
                End Using
            End Using
        End Sub

        Private Function KeyboardHookCallback(nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
            If nCode >= NativeMethods.HC_ACTION AndAlso wParam.ToInt32() = NativeMethods.WM_KEYDOWN Then
                Dim hookStruct = Runtime.InteropServices.Marshal.PtrToStructure(Of NativeMethods.KBDLLHOOKSTRUCT)(lParam)
                If hookStruct.vkCode = Keys.Escape Then
                    If _macroRunner.IsRunning Then
                        _macroRunner.Cancel()
                    End If
                ElseIf _macroRunner.IsRunning Then
                    ' +키(OemPlus/Add)로 깊이 증가, -키(OemMinus/Subtract)로 깊이 감소
                    If hookStruct.vkCode = Keys.Oemplus OrElse hookStruct.vkCode = Keys.Add Then
                        _macroRunner.AdjustDepth(1)
                        UpdateStatus($"AI 깊이 → {_macroRunner.CurrentAIDepth}")
                    ElseIf hookStruct.vkCode = Keys.OemMinus OrElse hookStruct.vkCode = Keys.Subtract Then
                        _macroRunner.AdjustDepth(-1)
                        UpdateStatus($"AI 깊이 → {_macroRunner.CurrentAIDepth}")
                    End If
                End If
            End If
            Return NativeMethods.CallNextHookEx(_hookHandle, nCode, wParam, lParam)
        End Function

        Private Sub UninstallKeyboardHook()
            If _hookHandle <> IntPtr.Zero Then
                NativeMethods.UnhookWindowsHookEx(_hookHandle)
                _hookHandle = IntPtr.Zero
            End If
        End Sub

        Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
            UninstallKeyboardHook()
            SaveWindowSettings()
            ' 매크로가 있고 저장 폴더가 있으면 자동 저장
            CommitDelayToSelected()
            RefreshMacroList()
            If _macroItems.Count > 0 AndAlso Not String.IsNullOrEmpty(_currentMacroFile) Then
                Try
                    Dim winTitle2 = If(_targetWindow IsNot Nothing, _targetWindow.Title, Nothing)
                    MacroRunner.SaveToFile(_currentMacroFile, _macroItems, winTitle2)
                Catch
                End Try
            End If
            _screenshot?.Dispose()
            _templateImage?.Dispose()
            For Each item In _macroItems
                item.Dispose()
            Next
            MyBase.OnFormClosing(e)
        End Sub

        Protected Overrides Sub OnResize(e As EventArgs)
            MyBase.OnResize(e)
            If Me.WindowState = FormWindowState.Normal Then
                _lastNormalBounds = Me.Bounds
            End If
        End Sub

        Protected Overrides Sub OnMove(e As EventArgs)
            MyBase.OnMove(e)
            If Me.WindowState = FormWindowState.Normal Then
                _lastNormalBounds = Me.Bounds
            End If
        End Sub

        ' =============================================
        ' 윈도우 위치/크기 저장·복원
        ' =============================================
        Private Sub SaveWindowSettings()
            Try
                Dim bounds = If(Me.WindowState = FormWindowState.Normal, Me.Bounds, _lastNormalBounds)
                Dim lines As String() = {
                    bounds.X.ToString(),
                    bounds.Y.ToString(),
                    bounds.Width.ToString(),
                    bounds.Height.ToString(),
                    CInt(Me.WindowState).ToString()
                }
                IO.File.WriteAllLines(_settingsPath, lines)
            Catch
            End Try
        End Sub

        Private Sub RestoreWindowSettings()
            Try
                If Not IO.File.Exists(_settingsPath) Then Return

                Dim lines = IO.File.ReadAllLines(_settingsPath)
                If lines.Length < 5 Then Return

                Dim x = Integer.Parse(lines(0))
                Dim y = Integer.Parse(lines(1))
                Dim w = Integer.Parse(lines(2))
                Dim h = Integer.Parse(lines(3))
                Dim state = CType(Integer.Parse(lines(4)), FormWindowState)

                ' 저장된 위치가 화면 내에 있는지 확인
                Dim savedRect As New Rectangle(x, y, w, h)
                Dim onScreen = False
                For Each scr In Screen.AllScreens
                    If scr.WorkingArea.IntersectsWith(savedRect) Then
                        onScreen = True
                        Exit For
                    End If
                Next

                If Not onScreen Then Return

                Me.StartPosition = FormStartPosition.Manual
                Me.Bounds = savedRect
                _lastNormalBounds = savedRect

                If state = FormWindowState.Maximized Then
                    Me.WindowState = FormWindowState.Maximized
                Else
                    Me.WindowState = FormWindowState.Normal
                End If
            Catch
            End Try
        End Sub
    End Class
End Namespace
