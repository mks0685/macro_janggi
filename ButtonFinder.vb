Imports System.Drawing
Imports System.Drawing.Imaging

Namespace MacroAutoControl
    ''' <summary>
    ''' 화면에서 "대국신청" 버튼을 이미지 매칭으로 찾는 클래스
    ''' </summary>
    Public Class ButtonFinder
        ''' <summary>
        ''' 매칭 결과
        ''' </summary>
        Public Class MatchResult
            Public Property Found As Boolean
            Public Property Location As Drawing.Point
            Public Property Confidence As Double
            Public Property MatchRect As Rectangle
        End Class

        ''' <summary>
        ''' 템플릿 이미지를 사용하여 버튼 위치 찾기
        ''' </summary>
        Public Shared Function FindByTemplate(screenshot As Bitmap, templatePath As String, Optional threshold As Double = 0.85) As MatchResult
            If Not IO.File.Exists(templatePath) Then
                Return New MatchResult() With {.Found = False}
            End If

            Using template As New Bitmap(templatePath)
                Return FindByTemplate(screenshot, template, threshold)
            End Using
        End Function

        ''' <summary>
        ''' 템플릿 이미지를 사용하여 버튼 위치 찾기 (Bitmap 버전)
        ''' </summary>
        Public Shared Function FindByTemplate(screenshot As Bitmap, template As Bitmap, Optional threshold As Double = 0.85) As MatchResult
            Dim bestMatch As New MatchResult() With {.Found = False, .Confidence = 0}

            Dim srcWidth = screenshot.Width
            Dim srcHeight = screenshot.Height
            Dim tplWidth = template.Width
            Dim tplHeight = template.Height

            If tplWidth > srcWidth OrElse tplHeight > srcHeight Then
                Return bestMatch
            End If

            ' BitmapData로 빠른 픽셀 접근
            Dim srcData = screenshot.LockBits(
                New Rectangle(0, 0, srcWidth, srcHeight),
                ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb)
            Dim tplData = template.LockBits(
                New Rectangle(0, 0, tplWidth, tplHeight),
                ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb)

            Try
                Dim srcStride = srcData.Stride
                Dim tplStride = tplData.Stride
                Dim srcBytes(srcStride * srcHeight - 1) As Byte
                Dim tplBytes(tplStride * tplHeight - 1) As Byte

                Runtime.InteropServices.Marshal.Copy(srcData.Scan0, srcBytes, 0, srcBytes.Length)
                Runtime.InteropServices.Marshal.Copy(tplData.Scan0, tplBytes, 0, tplBytes.Length)

                Dim totalPixels = tplWidth * tplHeight
                ' 샘플링 간격 (성능 향상을 위해 모든 픽셀을 비교하지 않음)
                Dim stepSize = Math.Max(1, CInt(Math.Sqrt(totalPixels / 500)))

                Dim bestScore As Double = 0

                ' 스캔 간격 (속도 향상)
                Dim scanStep = 2

                For y = 0 To srcHeight - tplHeight - 1 Step scanStep
                    For x = 0 To srcWidth - tplWidth - 1 Step scanStep
                        Dim matchCount = 0
                        Dim sampleCount = 0
                        Dim earlyExit = False

                        ' 샘플 픽셀로 유사도 계산
                        For ty = 0 To tplHeight - 1 Step stepSize
                            For tx = 0 To tplWidth - 1 Step stepSize
                                sampleCount += 1

                                Dim srcIdx = (y + ty) * srcStride + (x + tx) * 4
                                Dim tplIdx = ty * tplStride + tx * 4

                                Dim db = Math.Abs(CInt(srcBytes(srcIdx)) - CInt(tplBytes(tplIdx)))
                                Dim dg = Math.Abs(CInt(srcBytes(srcIdx + 1)) - CInt(tplBytes(tplIdx + 1)))
                                Dim dr = Math.Abs(CInt(srcBytes(srcIdx + 2)) - CInt(tplBytes(tplIdx + 2)))

                                ' 색상 차이가 허용 범위 내이면 매칭
                                If db <= 30 AndAlso dg <= 30 AndAlso dr <= 30 Then
                                    matchCount += 1
                                End If
                            Next
                            ' 조기 종료: 현재까지 매칭률이 너무 낮으면 스킵
                            If sampleCount > 20 AndAlso (matchCount / CDbl(sampleCount)) < threshold * 0.7 Then
                                earlyExit = True
                                Exit For
                            End If
                        Next

                        If earlyExit Then Continue For

                        Dim score = matchCount / CDbl(sampleCount)
                        If score > bestScore Then
                            bestScore = score
                            bestMatch.Location = New Point(x + tplWidth \ 2, y + tplHeight \ 2)
                            bestMatch.MatchRect = New Rectangle(x, y, tplWidth, tplHeight)
                            bestMatch.Confidence = score
                        End If
                    Next
                Next

                bestMatch.Found = bestScore >= threshold

            Finally
                screenshot.UnlockBits(srcData)
                template.UnlockBits(tplData)
            End Try

            Return bestMatch
        End Function

        ''' <summary>
        ''' 템플릿 이미지를 사용하여 버튼 위치 찾기 (마스크 지원 버전)
        ''' 마스크의 빨간(R>200) 픽셀은 비교에서 제외
        ''' </summary>
        Public Shared Function FindByTemplate(screenshot As Bitmap, template As Bitmap, threshold As Double, mask As Bitmap) As MatchResult
            If mask Is Nothing Then Return FindByTemplate(screenshot, template, threshold)

            Dim bestMatch As New MatchResult() With {.Found = False, .Confidence = 0}

            Dim srcWidth = screenshot.Width
            Dim srcHeight = screenshot.Height
            Dim tplWidth = template.Width
            Dim tplHeight = template.Height

            If tplWidth > srcWidth OrElse tplHeight > srcHeight Then
                Return bestMatch
            End If

            ' BitmapData로 빠른 픽셀 접근
            Dim srcData = screenshot.LockBits(
                New Rectangle(0, 0, srcWidth, srcHeight),
                ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb)
            Dim tplData = template.LockBits(
                New Rectangle(0, 0, tplWidth, tplHeight),
                ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb)

            ' 마스크 크기를 템플릿에 맞춰 처리
            Dim maskBmp = mask
            If mask.Width <> tplWidth OrElse mask.Height <> tplHeight Then
                maskBmp = New Bitmap(mask, tplWidth, tplHeight)
            End If
            Dim maskData = maskBmp.LockBits(
                New Rectangle(0, 0, tplWidth, tplHeight),
                ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb)

            Try
                Dim srcStride = srcData.Stride
                Dim tplStride = tplData.Stride
                Dim maskStride = maskData.Stride
                Dim srcBytes(srcStride * srcHeight - 1) As Byte
                Dim tplBytes(tplStride * tplHeight - 1) As Byte
                Dim maskBytes(maskStride * tplHeight - 1) As Byte

                Runtime.InteropServices.Marshal.Copy(srcData.Scan0, srcBytes, 0, srcBytes.Length)
                Runtime.InteropServices.Marshal.Copy(tplData.Scan0, tplBytes, 0, tplBytes.Length)
                Runtime.InteropServices.Marshal.Copy(maskData.Scan0, maskBytes, 0, maskBytes.Length)

                Dim totalPixels = tplWidth * tplHeight
                Dim stepSize = Math.Max(1, CInt(Math.Sqrt(totalPixels / 500)))

                Dim bestScore As Double = 0
                Dim scanStep = 2

                For y = 0 To srcHeight - tplHeight - 1 Step scanStep
                    For x = 0 To srcWidth - tplWidth - 1 Step scanStep
                        Dim matchCount = 0
                        Dim sampleCount = 0
                        Dim earlyExit = False

                        For ty = 0 To tplHeight - 1 Step stepSize
                            For tx = 0 To tplWidth - 1 Step stepSize
                                ' 마스크 확인: R채널 > 200이면 스킵
                                Dim maskIdx = ty * maskStride + tx * 4
                                If maskBytes(maskIdx + 2) > 200 Then Continue For

                                sampleCount += 1

                                Dim srcIdx = (y + ty) * srcStride + (x + tx) * 4
                                Dim tplIdx = ty * tplStride + tx * 4

                                Dim db = Math.Abs(CInt(srcBytes(srcIdx)) - CInt(tplBytes(tplIdx)))
                                Dim dg = Math.Abs(CInt(srcBytes(srcIdx + 1)) - CInt(tplBytes(tplIdx + 1)))
                                Dim dr = Math.Abs(CInt(srcBytes(srcIdx + 2)) - CInt(tplBytes(tplIdx + 2)))

                                If db <= 30 AndAlso dg <= 30 AndAlso dr <= 30 Then
                                    matchCount += 1
                                End If
                            Next
                            If sampleCount > 20 AndAlso (matchCount / CDbl(sampleCount)) < threshold * 0.7 Then
                                earlyExit = True
                                Exit For
                            End If
                        Next

                        If earlyExit Then Continue For
                        If sampleCount = 0 Then Continue For

                        Dim score = matchCount / CDbl(sampleCount)
                        If score > bestScore Then
                            bestScore = score
                            bestMatch.Location = New Point(x + tplWidth \ 2, y + tplHeight \ 2)
                            bestMatch.MatchRect = New Rectangle(x, y, tplWidth, tplHeight)
                            bestMatch.Confidence = score
                        End If
                    Next
                Next

                bestMatch.Found = bestScore >= threshold

            Finally
                screenshot.UnlockBits(srcData)
                template.UnlockBits(tplData)
                maskBmp.UnlockBits(maskData)
                If maskBmp IsNot mask Then maskBmp.Dispose()
            End Try

            Return bestMatch
        End Function

        ''' <summary>
        ''' 색상 범위로 버튼 영역 찾기 (대국신청 버튼의 특징적인 색상 사용)
        ''' </summary>
        Public Shared Function FindByColor(screenshot As Bitmap,
                                           targetColor As Color,
                                           Optional colorTolerance As Integer = 40,
                                           Optional minWidth As Integer = 60,
                                           Optional minHeight As Integer = 20) As MatchResult

            Dim width = screenshot.Width
            Dim height = screenshot.Height

            Dim srcData = screenshot.LockBits(
                New Rectangle(0, 0, width, height),
                ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb)

            Dim result As New MatchResult() With {.Found = False}

            Try
                Dim srcBytes(srcData.Stride * height - 1) As Byte
                Runtime.InteropServices.Marshal.Copy(srcData.Scan0, srcBytes, 0, srcBytes.Length)

                ' 색상 매칭 맵 생성
                Dim matchMap(width - 1, height - 1) As Boolean

                For y = 0 To height - 1
                    For x = 0 To width - 1
                        Dim idx = y * srcData.Stride + x * 4
                        Dim b = srcBytes(idx)
                        Dim g = srcBytes(idx + 1)
                        Dim r = srcBytes(idx + 2)

                        If Math.Abs(CInt(r) - CInt(targetColor.R)) <= colorTolerance AndAlso
                           Math.Abs(CInt(g) - CInt(targetColor.G)) <= colorTolerance AndAlso
                           Math.Abs(CInt(b) - CInt(targetColor.B)) <= colorTolerance Then
                            matchMap(x, y) = True
                        End If
                    Next
                Next

                ' 연속된 색상 영역 찾기 (가장 큰 사각형 영역)
                Dim bestArea = 0
                Dim bestRect As Rectangle

                For y = 0 To height - minHeight - 1 Step 3
                    For x = 0 To width - minWidth - 1 Step 3
                        If Not matchMap(x, y) Then Continue For

                        ' 이 위치에서 연속 매칭 영역 크기 측정
                        Dim matchW = 0
                        Dim matchH = 0

                        ' 가로 연속 측정
                        While x + matchW < width AndAlso matchMap(x + matchW, y)
                            matchW += 1
                        End While

                        If matchW < minWidth Then Continue For

                        ' 세로 연속 측정
                        matchH = 1
                        Dim validH = True
                        While y + matchH < height AndAlso validH
                            Dim rowMatch = 0
                            For cx = x To x + matchW - 1
                                If matchMap(cx, y + matchH) Then rowMatch += 1
                            Next
                            If rowMatch >= matchW * 0.6 Then
                                matchH += 1
                            Else
                                validH = False
                            End If
                        End While

                        If matchH < minHeight Then Continue For

                        Dim area = matchW * matchH
                        If area > bestArea Then
                            bestArea = area
                            bestRect = New Rectangle(x, y, matchW, matchH)
                        End If
                    Next
                Next

                If bestArea > 0 Then
                    result.Found = True
                    result.MatchRect = bestRect
                    result.Location = New Point(
                        bestRect.X + bestRect.Width \ 2,
                        bestRect.Y + bestRect.Height \ 2)
                    result.Confidence = 0.7
                End If

            Finally
                screenshot.UnlockBits(srcData)
            End Try

            Return result
        End Function
    End Class
End Namespace
