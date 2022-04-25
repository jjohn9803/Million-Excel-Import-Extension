Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System.IO
Imports System.ComponentModel

Public Class Main_Form
    Public serverName As String = ""
    Public database As String = ""
    Public myConn As SqlConnection
    Public statusConnection As Boolean
    Public pwd_query As String
    Public import_type As String
    Public features As New ArrayList
    Private Sub SQL_Connection_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        mainFormLoad()
    End Sub
    Private Sub mainFormLoad()
        Dim form As New Function_Form
        form.TopLevel = False
        Me.WindowState = FormWindowState.Normal
        form.WindowState = FormWindowState.Maximized
        form.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        form.Visible = True

        panelMain.Controls.Add(form)

        Setting.getFileInfo()
    End Sub
    Public Function getServerName() As String
        Return Me.serverName
    End Function
    Public Function getDatabase() As String
        Return Me.database
    End Function
    Public Function getMyConn() As SqlConnection
        Return Me.myConn
    End Function
    Public Function getStatusConnection() As Boolean
        Return Me.statusConnection
    End Function
    Public Function getPwd_query() As String
        Return Me.pwd_query
    End Function
    Public Function getImport_type() As String
        Return Me.import_type
    End Function
    Public Function getFeatures() As ArrayList
        Return Me.features
    End Function
    Public Sub setServerName(ByVal serverName As String)
        Me.serverName = serverName
        txtServerName.Text = serverName
    End Sub
    Public Sub setDatabase(ByVal database As String)
        Me.database = database
        txtDatabaseName.Text = database
    End Sub
    Public Sub setMyConn(ByVal myConn As SqlConnection)
        Me.myConn = myConn
    End Sub
    Public Sub setStatusConnection(ByVal statusConnection As Boolean)
        If statusConnection Then
            lblSQLStatus.Text = "Connected"
            lblSQLStatus.ForeColor = Color.Green
        Else
            lblSQLStatus.Text = "Disconnected"
            lblSQLStatus.ForeColor = Color.Red
        End If
        Me.statusConnection = statusConnection
    End Sub
    Public Sub setPwd_query(ByVal pwd_query As String)
        Me.pwd_query = pwd_query
    End Sub
    Public Sub setImport_type(ByVal import_type As String)
        Me.import_type = import_type
    End Sub
    Public Sub setFeatures(ByVal features As ArrayList)
        Me.features = features
    End Sub
    Public Sub clearFeature()
        Me.features.Clear()
    End Sub
    Public Sub appendFeature(ByVal feature As String)
        Me.features.add(feature)
    End Sub
    Private Sub btnSetting_Click(sender As Object, e As EventArgs) Handles btnSetting.Click
        Dim loginForm As New Login
        loginForm.ShowDialog()
    End Sub
    Public Sub showSettingForm()
        Dim settingForm As New Setting
        settingForm.ShowDialog()
    End Sub
End Class