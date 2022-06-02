Public Class Loading
    Private duration As Decimal
    Private total_duration As Decimal
    Private maximum_peak As Integer = GetRandom(1, 5)
    Private url As String = "https://www.million.my/wp-content/uploads/2016/11/small-logo-1024x435-1.png"
    Public returnType As Integer = 0
    Private Sub Loading_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim img As Image = GetImageFromURL(url)
        If img IsNot Nothing Then
            PictureBox1.Image = img
        End If
        startTimer()
    End Sub
    Private Sub startTimer()
        Dim time As Integer = GetRandom(600, 1200)
        Me.total_duration = time
        Me.duration = 0
        Timer1.Start()
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If duration > total_duration Then
            ProgressBar1.Value = 100
            maximum_peak -= 1
            Timer1.Stop()
            Dim waitDuration As New Timer()
            waitDuration.Interval = 500
            waitDuration.Start()
            If maximum_peak <= 0 Then
                AddHandler waitDuration.Tick, AddressOf closeForm
            Else
                AddHandler waitDuration.Tick, AddressOf myTickEvent
            End If
        Else
            duration += Timer1.Interval
            Dim process As Integer = (duration / total_duration) * 100
            If process > 100 Then
                ProgressBar1.Value = 100
            Else
                ProgressBar1.Value = process
            End If
        End If
    End Sub
    Private Sub closeForm(sender As Object, e As EventArgs)
        sender.Stop()
        Me.Close()
    End Sub
    Private Sub myTickEvent(sender As Object, e As EventArgs)
        sender.Stop()
        startTimer()
    End Sub
    Private Function GetRandom(ByVal Min As Integer, ByVal Max As Integer) As Integer
        Dim Generator As System.Random = New System.Random()
        Return Generator.Next(Min, Max)
    End Function
    Private Function GetImageFromURL(ByVal url As String) As Image

        Dim retVal As Image = Nothing

        Try
            If Not String.IsNullOrWhiteSpace(url) Then
                Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(url.Trim)

                Using request As System.Net.WebResponse = req.GetResponse
                    Using stream As System.IO.Stream = request.GetResponseStream
                        retVal = New Bitmap(System.Drawing.Image.FromStream(stream))
                    End Using
                End Using
            End If

        Catch ex As Exception
            'MessageBox.Show(String.Format("An error occurred:{0}{0}{1}",
            '                              vbCrLf, ex.Message),
            '                              "Exception Thrown",
            '                              MessageBoxButtons.OK,
            '                              MessageBoxIcon.Warning)

        End Try

        Return retVal
    End Function
    Public Function getReturnType() As Boolean
        If returnType = -1 Then
            Return False
        End If
        Return True
    End Function
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        returnType = -1
        Me.Close()
    End Sub
End Class