Imports System.Drawing

Namespace MacroAutoControl
    ''' <summary>
    ''' 매크로 항목 데이터 클래스
    ''' </summary>
    Public Class MacroItem
        Public Property Name As String
        Public Property Image As Bitmap
        Public Property DelayAfterClick As Integer = 1000
        Public Property Threshold As Double = 0.75
        Public Property Button As ClickButton = ClickButton.Left
        Public Property SendKeys As String = ""
        ''' <summary>클릭 위치 오프셋 (템플릿 좌상단 기준, -1이면 중앙)</summary>
        Public Property ClickOffsetX As Integer = -1
        Public Property ClickOffsetY As Integer = -1
        ''' <summary>마스크 비트맵 (빨간 픽셀=마스킹 영역, 매칭에서 제외)</summary>
        Public Property Mask As Bitmap
        ''' <summary>이 항목이 추가될 때 선택된 대상 창 이름</summary>
        Public Property WindowTitle As String = ""

        ' AI 항목 속성
        ''' <summary>AI 항목 여부 (Threshold=-1로 구분)</summary>
        Public Property IsAI As Boolean = False
        ''' <summary>AI 진영 ("C" 초 또는 "H" 한)</summary>
        Public Property AISide As String = "C"
        ''' <summary>AI 탐색 깊이</summary>
        Public Property AIDepth As Integer = 5
        ''' <summary>AI 탐색 시간 제한 (초)</summary>
        Public Property AITime As Double = 10.0

        Public Sub New()
        End Sub

        Public Sub New(name As String, image As Bitmap, Optional delay As Integer = 1000, Optional threshold As Double = 0.75, Optional button As ClickButton = ClickButton.Left, Optional sendKeys As String = "", Optional clickOffX As Integer = -1, Optional clickOffY As Integer = -1)
            Me.Name = name
            Me.Image = CType(image.Clone(), Bitmap)
            Me.DelayAfterClick = delay
            Me.Threshold = threshold
            Me.Button = button
            Me.SendKeys = If(sendKeys, "")
            Me.ClickOffsetX = clickOffX
            Me.ClickOffsetY = clickOffY
        End Sub

        ''' <summary>AI 항목 생성용 팩토리 메서드</summary>
        Public Shared Function CreateAI(name As String, delay As Integer, side As String, depth As Integer, time As Double) As MacroItem
            Dim dummyImg As New Bitmap(1, 1)
            Dim item As New MacroItem() With {
                .Name = name,
                .Image = dummyImg,
                .DelayAfterClick = delay,
                .Threshold = -1,
                .IsAI = True,
                .AISide = side,
                .AIDepth = depth,
                .AITime = time
            }
            Return item
        End Function

        Private ReadOnly Property ButtonText As String
            Get
                Return If(Button = ClickButton.Right, "우클릭", "좌클릭")
            End Get
        End Property

        Public Overrides Function ToString() As String
            Dim winTag = If(String.IsNullOrEmpty(WindowTitle), "", $" [{WindowTitle}]")
            If IsAI Then
                Return $"{Name} (대기:{DelayAfterClick}ms, 깊이:{AIDepth}, 시간:{AITime:F0}s){winTag}"
            End If
            ' 키전송 전용 항목 (Threshold=0, SendKeys 있음)
            If Threshold <= 0 AndAlso Not String.IsNullOrEmpty(SendKeys) Then
                Return $"{Name} (대기:{DelayAfterClick}ms, {SendKeys}){winTag}"
            End If
            Dim keyInfo = If(String.IsNullOrEmpty(SendKeys), "", $", 키:{SendKeys}")
            Dim clickPos = If(ClickOffsetX >= 0, $", 클릭:{ClickOffsetX},{ClickOffsetY}", "")
            Return $"{Name} ({ButtonText}, 대기:{DelayAfterClick}ms{clickPos}{keyInfo}){winTag}"
        End Function

        Public Sub Dispose()
            Image?.Dispose()
            Image = Nothing
            Mask?.Dispose()
            Mask = Nothing
        End Sub
    End Class
End Namespace
