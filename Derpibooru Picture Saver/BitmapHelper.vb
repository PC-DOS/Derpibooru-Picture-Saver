Imports System.Drawing
Imports System.Drawing.Imaging
Module BitmapHelper
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
