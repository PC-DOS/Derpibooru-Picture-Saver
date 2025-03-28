Imports System.Drawing
Imports System.Drawing.Imaging
Module BitmapHelper
    Public Function PasteBitmap(ByRef Target As Bitmap, ByVal Source As Bitmap, XPos As Integer, YPos As Integer) As Bitmap
        For x = XPos To XPos + Source.Width - 1
            For y = YPos To YPos + Source.Height - 1
                If x >= 0 And x < Target.Width Then
                    If y >= 0 And y < Target.Height Then
                        Target.SetPixel(x, y, Source.GetPixel(x - XPos, y - YPos))
                    End If
                End If
            Next
        Next
        Return Target
    End Function
    Public Function PadBitmap(ByVal Source As Bitmap, NewWidth As Integer, NewHeight As Integer, Optional BackgroundColor As Color = Nothing, Optional SourcePosition As System.Drawing.ContentAlignment = ContentAlignment.MiddleCenter) As Bitmap
        If IsNothing(BackgroundColor) Then
            BackgroundColor = Color.Transparent
        End If

        Dim PaddedBitmap As Bitmap = New Bitmap(NewWidth, NewHeight)
        '初始化
        For y As Integer = 0 To PaddedBitmap.Height - 1
            For x As Integer = 0 To PaddedBitmap.Width - 1
                PaddedBitmap.SetPixel(x, y, BackgroundColor)
            Next
        Next

        '放置影像
        Dim Padder As Graphics = Graphics.FromImage(PaddedBitmap)
        Padder.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        Select Case SourcePosition
            Case ContentAlignment.TopLeft
                PasteBitmap(PaddedBitmap, Source, 0, 0)
            Case ContentAlignment.TopCenter
                PasteBitmap(PaddedBitmap, Source, Int((NewWidth - Source.Width) / 2), 0)
            Case ContentAlignment.TopRight
                PasteBitmap(PaddedBitmap, Source, NewWidth - Source.Width, 0)
            Case ContentAlignment.MiddleLeft
                PasteBitmap(PaddedBitmap, Source, 0, Int((NewHeight - Source.Height) / 2))
            Case ContentAlignment.MiddleCenter
                PasteBitmap(PaddedBitmap, Source, Int((NewWidth - Source.Width) / 2), Int((NewHeight - Source.Height) / 2))
            Case ContentAlignment.MiddleRight
                PasteBitmap(PaddedBitmap, Source, NewWidth - Source.Width, Int((NewHeight - Source.Height) / 2))
            Case ContentAlignment.BottomLeft
                PasteBitmap(PaddedBitmap, Source, 0, NewHeight - Source.Height)
            Case ContentAlignment.BottomCenter
                PasteBitmap(PaddedBitmap, Source, Int((NewWidth - Source.Width) / 2), NewHeight - Source.Height)
            Case ContentAlignment.BottomRight
                PasteBitmap(PaddedBitmap, Source, NewWidth - Source.Width, NewHeight - Source.Height)
        End Select
        Padder.Dispose()

        Return PaddedBitmap
    End Function
    Public Function RescaleBitmap(ByVal Source As Bitmap, NewWidth As Integer, NewHeight As Integer) As Bitmap
        Dim RescaledBitmap As Bitmap = New Bitmap(NewWidth, NewHeight)
        Dim Rescaler As Graphics = Graphics.FromImage(RescaledBitmap)
        Rescaler.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        Rescaler.DrawImage(Source, 0, 0, NewWidth, NewHeight)
        Rescaler.Dispose()
        Return RescaledBitmap
    End Function
    Public Function GetImageEncoderInfo(Format As ImageFormat) As ImageCodecInfo
        For Each CodecInfo As ImageCodecInfo In ImageCodecInfo.GetImageEncoders()
            If CodecInfo.FormatID = Format.Guid Then
                Return CodecInfo
            End If
        Next
        Return Nothing
    End Function
End Module
