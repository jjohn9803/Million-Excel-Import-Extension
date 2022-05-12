Public Class Form1
    Private msg As String = ""
    Dim DrawingFont As New Font("Arial", 25)
    Dim CaptchaImage As New Bitmap(140, 40)
    Dim CaptchaGraf As Graphics = Graphics.FromImage(CaptchaImage)
    Dim Alphabet As String = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz"
    Dim CaptchaString, TickRandom As String
    Dim ProcessNumber As Integer

    Private Sub txtInput_TextChanged(sender As Object, e As EventArgs) Handles txtInput.TextChanged
        msg = txtInput.Text
        If msg.Equals(CaptchaString) Then
            Threading.Thread.Sleep(500)
            Me.Close()
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        GenerateCaptcha()
        txtInput.Focus()
    End Sub

    Private Sub txtInput_KeyDown(sender As Object, e As KeyEventArgs) Handles txtInput.KeyDown
        If e.KeyValue = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    'Private Async Function Form1_KeyDown(sender As Object, e As KeyEventArgs) As Task Handles Me.KeyDown
    '    If e.KeyValue = Keys.Escape Then
    '        Me.Close()
    '    End If
    'End Function

    Private Sub GenerateCaptcha()
        ProcessNumber = My.Computer.Clock.LocalTime.Millisecond
        If ProcessNumber < 521 Then
            ProcessNumber = ProcessNumber \ 10
            CaptchaString = Alphabet.Substring(ProcessNumber, 1)
        Else
            CaptchaString = CStr(My.Computer.Clock.LocalTime.Second \ 6)
        End If
        ProcessNumber = My.Computer.Clock.LocalTime.Second
        If ProcessNumber < 30 Then
            ProcessNumber = Math.Abs(ProcessNumber - 8)
            CaptchaString += Alphabet.Substring(ProcessNumber, 1)
        Else
            CaptchaString += CStr(My.Computer.Clock.LocalTime.Minute \ 6)
        End If
        ProcessNumber = My.Computer.Clock.LocalTime.DayOfYear
        If ProcessNumber Mod 2 = 0 Then
            ProcessNumber = ProcessNumber \ 8
            CaptchaString += Alphabet.Substring(ProcessNumber, 1)
        Else
            CaptchaString += CStr(ProcessNumber \ 37)
        End If
        TickRandom = My.Computer.Clock.TickCount.ToString
        ProcessNumber = Val(TickRandom.Substring(TickRandom.Length - 1, 1))
        If ProcessNumber Mod 2 = 0 Then
            CaptchaString += CStr(ProcessNumber)
        Else
            ProcessNumber = Math.Abs(Int(Math.Cos(Val(TickRandom)) * 51))
            CaptchaString += Alphabet.Substring(ProcessNumber, 1)
        End If
        ProcessNumber = My.Computer.Clock.LocalTime.Hour
        If ProcessNumber Mod 2 = 0 Then
            ProcessNumber = Math.Abs(Int(Math.Sin(Val(My.Computer.Clock.LocalTime.Year)) * 51))
            CaptchaString += Alphabet.Substring(ProcessNumber, 1)
        Else
            CaptchaString += CStr(ProcessNumber \ 3)
        End If
        ProcessNumber = My.Computer.Clock.LocalTime.Millisecond
        If ProcessNumber > 521 Then
            ProcessNumber = Math.Abs((ProcessNumber \ 10) - 52)
            CaptchaString += Alphabet.Substring(ProcessNumber, 1)
        Else
            CaptchaString += CStr(My.Computer.Clock.LocalTime.Second \ 6)
        End If
        CaptchaGraf.Clear(Color.White)

        For hasher As Integer = 0 To 5
            CaptchaGraf.DrawString(CaptchaString.Substring(hasher, 1), DrawingFont, Brushes.Black, hasher * 20 + hasher + ProcessNumber \ 200, (hasher Mod 3) * (ProcessNumber \ 200))
        Next
        PictureBox1.Image = CaptchaImage
    End Sub
End Class