Public Class Login
    Private ReadOnly staffID As String = My.Resources.staffID
    Private ReadOnly staffPassword As String = My.Resources.staffPassword
    Private Sub btnSetting_Click(sender As Object, e As EventArgs) Handles btnSetting.Click
        If staffID.Equals(getStaffID) AndAlso staffPassword.Equals(getPassword) Then
            Main_Form.showSettingForm()
            Me.Close()
        Else
            MsgBox("Staff Id or password is incorrect!", MsgBoxStyle.Critical)
        End If
    End Sub
    Private Sub txtStaffID_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtStaffID.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            txtPassword.Focus()
            e.Handled = True
        End If
    End Sub
    Private Sub txtPassword_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPassword.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            btnSetting.Focus()
            e.Handled = True
        End If
    End Sub
    Private Sub txtPassword_LostFocus(sender As Object, e As EventArgs) Handles txtPassword.LostFocus
        lblCapsLock.Visible = False
    End Sub
    Private Sub txtPassword_GotFocus(sender As Object, e As EventArgs) Handles txtPassword.GotFocus
        If Control.IsKeyLocked(Keys.CapsLock) Then
            lblCapsLock.Visible = True
        Else
            lblCapsLock.Visible = False
        End If
    End Sub

    Private Sub txtPassword_KeyUp(sender As Object, e As KeyEventArgs) Handles txtPassword.KeyUp, txtPassword.KeyDown
        If Control.IsKeyLocked(Keys.CapsLock) Then
            lblCapsLock.Visible = True
        Else
            lblCapsLock.Visible = False
        End If
    End Sub

    Private Function getStaffID() As String
        Return Encryption.Encrypt(txtStaffID.Text, My.Resources.myPassword)
    End Function
    Private Function getPassword() As String
        Return Encryption.Encrypt(txtPassword.Text, My.Resources.myPassword)
    End Function
End Class