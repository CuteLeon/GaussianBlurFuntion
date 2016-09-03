Option Explicit On
Public Class Form1
    Dim Time(1) As Double

    Private Sub TestButton_Click(sender As Object, e As EventArgs) Handles Test1.Click, Test2.Click, Test3.Click, Test4.Click
        '记录事件开始时间（精确到1x10^(-6)秒）
        Time(0) = My.Computer.Clock.LocalTime.TimeOfDay.TotalSeconds

        PictureBox1.Image = CType(sender, Button).Image
        Dim TempPicture As Bitmap = PictureBox1.Image
        GaussianBlur(PictureBox1.Image, PictureBox1, SetPixel.Value)

        '记录事件结束时间
        Time(1) = My.Computer.Clock.LocalTime.TimeOfDay.TotalSeconds

        '显示处理事件耗时
        MsgBox("Task finished!" & vbCrLf & "共用时： 【" & Time(1) - Time(0) & "】 秒。", MsgBoxStyle.Information)
    End Sub

    Private Sub GaussianBlur(ByVal PictureRes As Bitmap, PictureBoxName As PictureBox, ByVal Pixel As Long)
        Dim GetPicture As Bitmap = PictureRes
        Dim PicWidth As Long = PictureRes.Width, PicHeight As Long = PictureRes.Height
        Dim PositionX, PositionY As Long, RoundX, RoundY As Integer
        Dim SumA, SumR, SumG, SumB As Double
        Dim Gaussian As Double, GaussianCount As Double, CenterColor As Color

        For PositionY = 0 To PicHeight - 1
            For PositionX = 0 To PicWidth - 1
                SumA = 0.0 : SumR = 0.0 : SumG = 0.0 : SumB = 0.0 : GaussianCount = 0.0
                For RoundY = PositionY - Pixel To PositionY + Pixel
                    For RoundX = PositionX - Pixel To PositionX + Pixel
                        If (0 <= RoundX And RoundX < PicWidth) And (0 <= RoundY And RoundY < PicHeight) Then
                            Gaussian = Math.Abs((1 / (2 * Math.PI * (Pixel + 1 / 2) ^ 2)) * Math.Exp(-((RoundX - PositionX) ^ 2 + (RoundY - PositionY) ^ 2) / (2 * (Pixel + 1 / 2) ^ 2)))
                            GaussianCount += Gaussian
                            'Alpha通道  模糊透明度
                            SumA += PictureRes.GetPixel(RoundX, RoundY).A * Gaussian
                            'RGB
                            SumR += PictureRes.GetPixel(RoundX, RoundY).R * Gaussian
                            SumG += PictureRes.GetPixel(RoundX, RoundY).G * Gaussian
                            SumB += PictureRes.GetPixel(RoundX, RoundY).B * Gaussian
                        End If
                    Next
                Next

                SumA /= GaussianCount
                SumR /= GaussianCount
                SumG /= GaussianCount
                SumB /= GaussianCount
                CenterColor = Color.FromArgb(SumA, SumR, SumG, SumB)
                GetPicture.SetPixel(PositionX, PositionY, CenterColor)
            Next
        Next

        PictureBoxName.Image = GetPicture
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Gaussian As Double, Index, X, Y As Integer
        Dim Pixel As Integer = 70
        Dim Count As Long = 270 * 360
        Dim Time1 As Double = My.Computer.Clock.LocalTime.TimeOfDay.TotalSeconds
        For Index = 0 To Count
            For X = -Pixel To Pixel
                For Y = -Pixel To Pixel
                    Gaussian = Math.Abs((1 / (2 * Math.PI * (Pixel + 1 / 2) ^ 2)) * Math.Exp(-((X) ^ 2 + (Y) ^ 2) / (2 * (Pixel + 1 / 2) ^ 2)))
                Next
            Next
        Next
        Dim Time2 As Double = My.Computer.Clock.LocalTime.TimeOfDay.TotalSeconds
        MsgBox(Time2 - Time1)
    End Sub

End Class
