Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System.IO
Imports System.ComponentModel

Public Class SQL_Connection_Form
    Public serverName As String
    Public database As String
    Public myConn As SqlConnection
    Public statusConnection As Boolean
    Public pwd_query As String
    Public import_type As String
    Private pwd_mode As Integer
    Private uid As String
    Private pwd As String
    'Public form2 As New ExcelImporter
    Private form_default_size As New Size
    Public form2 As New Function_Form
    Public form3 As New Maintainance_Form
    Private Sub SQL_Connection_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.StartPosition = FormStartPosition.CenterScreen
        cbPasswordOption.SelectedIndex = 0
        pwd_mode = 0
        getFileInfo()
        getServerList()
        tabFormLoad()
    End Sub
    Private Sub tabFormLoad()
        form_default_size = Me.Size

        form2.TopLevel = False
        form2.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        form2.Location = New System.Drawing.Point(0, 0)
        form2.Visible = False
        TabPageExcel.Controls.Add(form2)

        form3.TopLevel = False
        form3.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        form3.Location = New System.Drawing.Point(0, 0)
        form3.Visible = False
        TabPageMaintain.Controls.Add(form3)
    End Sub
    Private Sub getFileInfo()
        'Try
        '    Dim fileReader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(getConnectionSetting())
        'Catch ex As System.IO.FileNotFoundException
        '    Return
        'End Try
        Try
            Dim objReader As New System.IO.StreamReader(getConnectionSetting)
            Dim settingBoolean As Boolean = False
            Do While objReader.Peek() <> -1
                settingBoolean = True
                Dim result = objReader.ReadLine
                Dim result2 = (result.Substring(result.IndexOf(":") + 1)).Trim
                If result.Contains("Server Name:") Then
                    If result2 <> String.Empty Then
                        cbServerList.Text() = result2
                        serverName = result2
                    End If
                End If
                If result.Contains("Database:") Then
                    If result2 <> String.Empty Then
                        database = result2
                    End If
                End If
                If result.Contains("pwd_option:") Then
                    If result2 <> String.Empty Then
                        cbPasswordOption.SelectedIndex = result2
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
                Me.btnSql.PerformClick()
                If database <> String.Empty Then
                    cbDatabase.Text() = database
                End If
            End If
            objReader.Close()
        Catch ex As System.IO.FileNotFoundException

        End Try
    End Sub
    Private Function getConnectionSetting() As String
        Return returnUpperFolder(Application.StartupPath(), 2) + "setting.dll"
    End Function
    Public Function returnUpperFolder(filePath As String, times As Integer)
        For index As Integer = 1 To times
            filePath = Path.GetFullPath(Path.Combine(filePath, "..\"))
        Next
        Return filePath
    End Function
    Private Sub setPasswordOption()
        If pwd_mode = 0 Then
            pwd_query = "Trusted_Connection=yes;"
        Else
            uid = txtUserId.Text.ToString
            pwd = txtPassword.Text.ToString
            pwd_query = "user id = " + uid + ";password=" + pwd
        End If
    End Sub
    Private Sub getServerList()
        Try
            Dim dt As DataTable = SqlDataSourceEnumerator.Instance.GetDataSources
            For Each dr As DataRow In dt.Rows
                cbServerList.Items.Add(String.Concat(dr("ServerName"), "\", dr("InstanceName")))
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, Me.Text)
        End Try
    End Sub
    Private Sub btnSql_Click(sender As Object, e As EventArgs) Handles btnSql.Click
        serverName = cbServerList.Text().Trim
        setPasswordOption()

        If (serverName.Equals(String.Empty)) Then
            updateSQLStatus(0, "Server name must be filled out!")
        End If
        Try
            myConn = New SqlConnection("Data Source=" + serverName + ";" &
                                    "Initial Catalog=master;" + pwd_query)
            myConn.Open()
            Using myConn
                Dim cmd As New SqlCommand("Select name from master.dbo.sysdatabases WHERE dbid > 4;", myConn)
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
            myConn.Close()
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
            lblSQLStatus.Text = "Disconnected"
            lblSQLStatus.ForeColor = Color.Red
            statusConnection = False
        ElseIf mode = 1 Then
            cbDatabase.Enabled = True
            lblSQLStatus.Text = "Connected"
            lblSQLStatus.ForeColor = Color.Green
            statusConnection = True
        End If
    End Sub
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            Dim content = "Server Name: " + serverName +
                        vbNewLine + "Database: " + database +
                        vbNewLine + "pwd_option: " + pwd_mode.ToString +
                        vbNewLine + "user_id: " + uid +
                        vbNewLine + "pwd: " + pwd
            System.IO.File.WriteAllText(getConnectionSetting, content)
            MsgBox("Setting Saved!")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, Me.Text)
        End Try

    End Sub
    Private Sub btnTestConnection_Click(sender As Object, e As EventArgs) Handles btnTestConnection.Click
        Try
            myConn = New SqlConnection("Data Source=" + serverName + ";" &
                                   "Initial Catalog=" + database + ";" + pwd_query)
            myConn.Open()
            'Using myConn
            '    Dim cmd As New SqlCommand("Select * from master.dbo.sysdatabases WHERE dbid > 4;", myConn)
            '    'Dim cmd As New SqlCommand("Select * from master.dbo.systemdatabases where name='" + database.Trim + "'", myConn)
            '    Using reader As SqlDataReader = cmd.ExecuteReader()
            '        While reader.Read()
            '            cbDatabase.Items.Add(reader.GetValue(0))
            '        End While
            '    End Using
            '    If (cbDatabase.Items.Count > 0) Then
            '        cbDatabase.Enabled = True
            '        lblSQLStatus.Text = "Connected"
            '        lblSQLStatus.ForeColor = Color.Green
            '    Else
            '        cbDatabase.Enabled = False
            '        lblSQLStatus.Text = "Disconnected"
            '        lblSQLStatus.ForeColor = Color.Red
            '    End If
            'End Using
            If myConn.State = 1 Then
                MsgBox("Connect successfully!", MsgBoxStyle.Information)
            End If
            myConn.Close()
        Catch ex As Exception
            If (ex.GetHashCode = 49205706) Then
                updateSQLStatus(0, "Failed to connect the server!")
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, Me.Text)
            End If
        End Try
    End Sub

    Private Sub cbDatabase_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbDatabase.SelectedIndexChanged
        database = cbDatabase.Text.ToString
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
    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedTab.TabIndex = 1 Then
            form2.Visible = True
            Me.Size = form2.Size
            'Me.WindowState = FormWindowState.Maximized
            Me.WindowState = FormWindowState.Normal
            form2.WindowState = FormWindowState.Maximized
        ElseIf TabControl1.SelectedTab.TabIndex = 2 Then
            form3.Visible = True
            Me.Size = form3.Size
            Me.WindowState = FormWindowState.Normal
            form3.WindowState = FormWindowState.Maximized
            form3.readMaintainSetting()
        Else
            form2.Visible = False
            form3.Visible = False
            Me.Size = form_default_size
            Me.WindowState = FormWindowState.Normal
        End If
    End Sub
    Public Sub closeForm()
        Function_Form.Close()
        'ExcelImporter.Close()
        Maintainance_Form.Close()
        Me.Close()
    End Sub
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        closeForm()
    End Sub
End Class