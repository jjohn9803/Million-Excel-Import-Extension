Public Class Form1
    Private msg As String = ""
    Private Async Function Form1_KeyDown(sender As Object, e As KeyEventArgs) As Task Handles Me.KeyDown
        If e.KeyValue = Keys.Escape Then
            Me.Close()
        ElseIf e.KeyValue = Keys.Back Then
            If msg.Length = 0 Then
                Return
            End If
            msg = msg.Substring(0, msg.Length - 1)
        ElseIf e.KeyValue = Keys.Space Then
            If msg.Length > 17 Then
                Return
            End If
            msg += " "
        Else
            If msg.Length > 17 Then
                Return
            End If
            Dim allowedChars As String = "abcdefghijklmnopqrstuvwxyz"
            If Not allowedChars.Contains(e.KeyCode.ToString.ToLower) Then
                Return
                e.Handled = True
            End If
            msg += e.KeyCode.ToString.ToLower
        End If
        lblInput.Text = msg
        If msg.Equals("rockbell software") Then
            Threading.Thread.Sleep(500)
            Me.Close()
        End If
    End Function
End Class