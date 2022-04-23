Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System.IO
Imports System.ComponentModel
Public Class Setting
    Private myConnn As SqlConnection
    Private pwd_mode As Integer
    Private uid As String
    Private pwd As String
    Private Sub SQL_Connection_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.StartPosition = FormStartPosition.CenterScreen
        cbPasswordOption.SelectedIndex = 0
        pwd_mode = 0
        getFileInfo()
    End Sub
    Public Sub getFileInfo()
        'Try
        '    Dim fileReader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(getConnectionSetting())
        'Catch ex As System.IO.FileNotFoundException
        '    Return
        'End Try
        Try
            If FileLen(getConnectionSetting()) = 0 Then
                Return
            End If
            Dim lineOneFromFile As String = Encryption.Decrypt(IO.File.ReadAllLines(getConnectionSetting)(0), My.Resources.myPassword)
            Dim objReader As New System.IO.StringReader(lineOneFromFile)
            Dim settingBoolean As Boolean = False
            Do While objReader.Peek() <> -1
                settingBoolean = True
                'Dim EncryptedResult = objReader.ReadLine
                'Dim result = Encryption.Decrypt(EncryptedResult, My.Resources.myPassword)
                'Dim result2 = result.ToString.Split(":")(1).Trim
                Dim result = objReader.ReadLine
                Dim result2 = result.Split(":")(1)
                If result.Contains("Server Name:") Then
                    If result2 <> String.Empty Then
                        cbServerList.Text() = result2
                        Main_Form.setServerName(result2)
                    End If
                End If
                If result.Contains("Database:") Then
                    If result2 <> String.Empty Then
                        Main_Form.setDatabase(result2)
                    End If
                End If
                If result.Contains("pwd_option:") Then
                    If result2 <> String.Empty Then
                        cbPasswordOption.SelectedIndex = CInt(result2)
                        pwd_mode = result2
                    End If
                End If
                If result.Contains("user_id:") Then
                    If result2 <> String.Empty Then
                        txtUserId.Text() = result2
                        uid = result2
                    End If
                End If
                If result.Contains("pwd:") Then
                    If result2 <> String.Empty Then
                        txtPassword.Text() = result2
                        pwd = result2
                    End If
                End If
            Loop
            If settingBoolean Then
                findDatabase()
                If Main_Form.getDatabase <> String.Empty Then
                    cbDatabase.Text() = Main_Form.getDatabase
                End If
            End If
            objReader.Close()
        Catch ex As System.IO.FileNotFoundException

        End Try
    End Sub
    Private Function getConnectionSetting() As String
        Return "setting.dll"
    End Function
    Public Function returnUpperFolder(filePath As String, times As Integer)
        For index As Integer = 1 To times
            filePath = Path.GetFullPath(Path.Combine(filePath, "..\"))
        Next
        Return filePath
    End Function
    Private Sub setPasswordOption()
        If pwd_mode = 0 Then
            Main_Form.setPwd_query("Trusted_Connection=yes;")
        Else
            uid = txtUserId.Text.ToString
            pwd = txtPassword.Text.ToString
            Main_Form.setPwd_query("user id = " + uid + ";password=" + pwd)
        End If
    End Sub
    Private Sub getServerList()
        Try
            'Dim dt As DataTable = SqlDataSourceEnumerator.Instance.GetDataSources
            cbServerList.Items.Clear()
            Dim dt As DataTable = SqlClientFactory.Instance.CreateDataSourceEnumerator.GetDataSources()
            For Each dr As DataRow In dt.Rows
                cbServerList.Items.Add(String.Concat(dr("ServerName"), "\", dr("InstanceName")))
                'cbServerList.Items.Add(dr(4))
            Next
            If dt.Rows.Count <= 1 Then
                MsgBox("Found " + cbServerList.Items.Count.ToString + " server!", MsgBoxStyle.Information)
            Else
                MsgBox("Found " + cbServerList.Items.Count.ToString + " servers!", MsgBoxStyle.Information)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, Me.Text)
        End Try
    End Sub
    Private Sub btnSql_Click(sender As Object, e As EventArgs) Handles btnSql.Click
        findDatabase()
    End Sub
    Private Sub findDatabase()
        Main_Form.setServerName(cbServerList.Text().Trim)
        setPasswordOption()
        If (Main_Form.getServerName.Equals(String.Empty)) Then
            updateSQLStatus(0, "Server name must be filled out!")
            Return
        End If
        Try
            Main_Form.setMyConn(New SqlConnection("Data Source=" + Main_Form.getServerName + ";" &
                                    "Initial Catalog=master;" + Main_Form.getPwd_query))
            Main_Form.myConn.Open()
            Using Main_Form.myConn
                Dim cmd As New SqlCommand("Select name from master.dbo.sysdatabases WHERE dbid > 4;", Main_Form.myConn)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    cbDatabase.Items.Clear()
                    While reader.Read()
                        cbDatabase.Items.Add(reader.GetValue(0))
                    End While
                End Using
                If (cbDatabase.Items.Count > 0) Then
                    updateSQLStatus(1, "")
                Else
                    updateSQLStatus(0, "Could not find any database from the server!")
                End If
            End Using
            Main_Form.myConn.Close()
        Catch ex As Exception
            If (ex.GetHashCode = 49205706) Then
                updateSQLStatus(0, "Failed to connect the server!")
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, Me.Text)
            End If
        End Try
    End Sub
    Private Sub updateSQLStatus(mode As Integer, message As String)
        If mode = 0 Then
            If message.Equals(String.Empty) = False Then
                MsgBox(message, MsgBoxStyle.Critical)
                cbServerList.Text = String.Empty
            End If
            cbDatabase.Items.Clear()
            cbDatabase.Enabled = False
            Main_Form.setStatusConnection(False)
        ElseIf mode = 1 Then
            cbDatabase.Enabled = True
            Main_Form.setStatusConnection(True)
        End If
    End Sub
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            Dim content = "Server Name:" + Main_Form.getServerName +
                        vbNewLine + "Database:" + Main_Form.getDatabase +
                        vbNewLine + "pwd_option:" + pwd_mode.ToString +
                        vbNewLine + "user_id:" + uid +
                        vbNewLine + "pwd:" + pwd
            System.IO.File.WriteAllText(getConnectionSetting, Encryption.Encrypt(content, My.Resources.myPassword))
            MsgBox("'" + content + "'" + vbNewLine + vbNewLine + "Setting Saved!", MsgBoxStyle.Information)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, Me.Text)
        End Try

    End Sub
    Private Sub btnTestConnection_Click(sender As Object, e As EventArgs) Handles btnTestConnection.Click
        Me.Cursor = Cursors.WaitCursor
        Try
            If Main_Form.getServerName.Equals(String.Empty) Or Main_Form.getDatabase.Equals(String.Empty) Or Not Main_Form.getStatusConnection Then
                MsgBox("Failed to connect the server!", MsgBoxStyle.Critical)
                Me.Cursor = Cursors.Default
                Return
            End If
            Main_Form.setMyConn(New SqlConnection("Data Source=" + Main_Form.getServerName + ";" &
                                    "Initial Catalog=" + Main_Form.getDatabase + ";" + Main_Form.getPwd_query))
            Main_Form.myConn.Open()
            Dim connectStatus = False
            Dim cmd As New SqlCommand("SELECT * FROM INFORMATION_SCHEMA.COLUMNS", Main_Form.getMyConn)
            Using reader As SqlDataReader = cmd.ExecuteReader()
                While reader.Read()
                    connectStatus = True
                End While
            End Using
            If Main_Form.getMyConn.State = ConnectionState.Open And connectStatus Then
                MsgBox("Connect successfully!", MsgBoxStyle.Information)
            Else
                MsgBox("Failed to connect the server!", MsgBoxStyle.Critical)
            End If
            Main_Form.myConn.Close()
        Catch ex As Exception
            If (ex.GetHashCode = 49205706) Then
                updateSQLStatus(0, "Failed to connect the server!")
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, Me.Text)
            End If
        End Try
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub cbDatabase_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbDatabase.SelectedIndexChanged
        Main_Form.setDatabase(cbDatabase.Text.ToString)
    End Sub

    Private Sub cbPasswordOption_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbPasswordOption.SelectedIndexChanged
        pwd_mode = cbPasswordOption.SelectedIndex
        cbShowPassword.Checked = False
        If cbPasswordOption.SelectedIndex = 1 Then
            txtUserId.Enabled = True
            txtPassword.Enabled = True
            cbShowPassword.Enabled = True
        Else
            txtUserId.Enabled = False
            txtPassword.Enabled = False
            cbShowPassword.Enabled = False
        End If
    End Sub

    Private Sub cbServerList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbServerList.SelectedIndexChanged
        updateSQLStatus(0, "")
    End Sub

    Private Sub cbShowPassword_CheckedChanged(sender As Object, e As EventArgs) Handles cbShowPassword.CheckedChanged
        If cbShowPassword.Checked Then
            txtPassword.PasswordChar = ""
        Else
            txtPassword.PasswordChar = "*"
        End If
    End Sub
    'Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
    '    If TabControl1.SelectedTab.TabIndex = 1 Then
    '        form2.Visible = True
    '        Me.Size = form2.Size
    '        'Me.WindowState = FormWindowState.Maximized
    '        Me.WindowState = FormWindowState.Normal
    '        form2.WindowState = FormWindowState.Maximized
    '    ElseIf TabControl1.SelectedTab.TabIndex = 2 Then
    '        form3.Visible = True
    '        Me.Size = form3.Size
    '        Me.WindowState = FormWindowState.Normal
    '        form3.WindowState = FormWindowState.Maximized
    '        form3.readMaintainSetting()
    '    Else
    '        form2.Visible = False
    '        form3.Visible = False
    '        Me.Size = form_default_size
    '        Me.WindowState = FormWindowState.Normal
    '    End If
    'End Sub
    Public Sub closeForm()
        Function_Form.Close()
        'ExcelImporter.Close()
        Maintainance_Form.Close()
        Me.Close()
    End Sub
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        closeForm()
    End Sub

    Private Sub btnSearchServer_Click(sender As Object, e As EventArgs) Handles btnSearchServer.Click
        Me.Cursor = Cursors.WaitCursor
        getServerList()
        Me.Cursor = Cursors.Default
    End Sub
End Class