Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System.IO
Imports ExcelDataReader
Public Class Maintainance_Form
    Private serverName As String
    Private database As String
    Private myConn As SqlConnection
    Private statusConnection As Boolean
    Private pwd_query As String
    Dim tables As DataTableCollection
    Private Sub Maintainance_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Private Sub init()
        serverName = SQL_Connection_Form.serverName
        database = SQL_Connection_Form.database
        myConn = SQL_Connection_Form.myConn
        statusConnection = SQL_Connection_Form.statusConnection
        pwd_query = SQL_Connection_Form.pwd_query
    End Sub
    Public Sub readMaintainSetting()
        Try
            Using stream = File.Open(getMaintainSetting, FileMode.Open, FileAccess.Read)
                Using reader As IExcelDataReader = ExcelReaderFactory.CreateReader(stream)
                    Dim result As DataSet = reader.AsDataSet(New ExcelDataSetConfiguration() With {
                                                                     .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
                                                                     .UseHeaderRow = True}})
                    tables = result.Tables
                    cbManage.Items.Clear()
                    For Each table As DataTable In tables
                        cbManage.Items.Add(table.TableName)
                    Next
                End Using
            End Using
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
        'Try
        '    myConn = New SqlConnection("Data Source=" + serverName + ";" &
        '                            "Initial Catalog=" + database + ";" + pwd_query)
        '    myConn.Open()
        '    Using myConn
        '        Dim cmd As New SqlCommand("Select name from master.dbo.sysdatabases WHERE dbid > 4;", myConn)
        '        Using reader As SqlDataReader = cmd.ExecuteReader()
        '            cbManage.Items.Clear()
        '            While reader.Read()
        '                cbManage.Items.Add(reader.GetValue(0))
        '            End While
        '        End Using
        '        If (cbManage.Items.Count > 0) Then

        '        Else

        '        End If
        '    End Using
        '    myConn.Close()
        'Catch ex As Exception
        '    If (ex.GetHashCode = 49205706) Then
        '        'updateSQLStatus(0, "Failed to connect the server!")
        '    Else
        '        MsgBox(ex.Message, MsgBoxStyle.Critical, Me.Text)
        '    End If
        'End Try
    End Sub
    Private Function getMaintainSetting() As String
        Return "maintain.xls"
    End Function

    Private Sub cbManage_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbManage.SelectedIndexChanged
        Dim dt As DataTable = tables(cbManage.SelectedItem.ToString())
        'For Each row As DataRow In dt.Rows
        '    'MsgBox(row.Item(0))
        'Next row
        DataGridViewManage.Rows.Clear()

        For Each row As DataRow In dt.Rows
            Dim newRow As DataGridViewRow = DataGridViewManage.Rows(DataGridViewManage.Rows.Add())
            newRow.Cells(0).Value = row.Item("sql_format").ToString
            newRow.Cells(1).Value = row.Item("excel_format").ToString
            newRow.Cells(2).Value = row.Item("default_value").ToString
            newRow.Cells(3).Value = row.Item("formula").ToString
            newRow.Cells(4).Value = row.Item("data_type").ToString
        Next row

    End Sub
End Class