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
    Private Function getStaffID() As String
        Return Encryption.Encrypt(txtStaffID.Text, My.Resources.myPassword)
    End Function
    Private Function getPassword() As String
        Return Encryption.Encrypt(txtPassword.Text, My.Resources.myPassword)
    End Function
End Class