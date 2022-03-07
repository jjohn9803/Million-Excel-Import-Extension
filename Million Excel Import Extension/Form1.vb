Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports ExcelDataReader
Public Class Form1
    Dim deviceName As String = Environment.MachineName
    Dim myConn As SqlConnection
    Dim tableName As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnFileOpen_Click(sender As Object, e As EventArgs) Handles btnFileOpen.Click
        'https://www.youtube.com/watch?v=pQ1PpcIcHno
        Using ofd As OpenFileDialog = New OpenFileDialog() With {.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"}
            If ofd.ShowDialog() = DialogResult.OK Then
                lblFileLocation.Text = ofd.FileName
                Using stream = File.Open()
            End If
        End Using




        'If OpenFileDialog1.ShowDialog = DialogResult.OK Then
        '    Dim cn As OleDbConnection
        '    Dim cmd As OleDbCommand
        '    Dim da As OleDbDataAdapter
        '    Dim dt As DataTable
        '    da = New OleDbDataAdapter
        '    dt = New DataTable
        '    Dim fileLocation As String = OpenFileDialog1.FileName
        '    Dim fileExtension As String = New IO.FileInfo(fileLocation).Extension
        '    If (fileExtension.Equals(".xlsx")) Then
        '        'cn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileLocation +
        '        '                     ";Extended Properties='Excel 12.0; IMEX=1; HDR=Yes;'")
        '        MsgBox("here")
        '        cn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileLocation + ";Extended Properties=Excel 12.0;")
        '    Else
        '        cn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileLocation +
        '                             ";Extended Properties='Excel 8.0; IMEX=1; HDR=Yes;'")
        '    End If
        '    cmd = New OleDbCommand("select * from [sheet1$]", cn)
        '    da.SelectCommand = cmd
        '    da.Fill(dt)

        '    DataGridView1.DataSource = dt
        '    lblFileLocation.Text = OpenFileDialog1.FileName

        'End If
    End Sub

    Private Sub btnSql_Click(sender As Object, e As EventArgs) Handles btnSql.Click
        Try
            lblSQLStatus.ForeColor = Color.Green
            lblSQLStatus.Text = "Connected"
            tableName = "sys.Tables"
            bindGrid(tableName, "name")
        Catch ex As SqlException
            MsgBox(ex.Message, MsgBoxStyle.Critical, Me.Text)
        End Try
    End Sub
    Private Sub bindGrid(table As String, rows As String)
        myConn = New SqlConnection("Data Source=" + deviceName + "\ML001; " &
                                   "Initial Catalog=demo;user id=sa;password=<ml$a>")
        'Open the connection.
        myConn.Open()
        Using myConn
            Using cmd As New SqlCommand("SELECT " + rows + " FROM " + table, myConn)
                cmd.CommandType = CommandType.Text
                Using sda As New SqlDataAdapter(cmd)
                    Using ds As New DataSet()
                        sda.Fill(ds)
                        DataGridViewTables.DataSource = ds.Tables(0)
                    End Using
                End Using
            End Using
        End Using
        tableName = table
        myConn.Close()
    End Sub

    Private Sub DataGridViewTables_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewTables.CellDoubleClick
        If (e.RowIndex <> -1 And e.ColumnIndex <> -1) Then
            Dim cellValue = DataGridViewTables.CurrentRow.Cells(e.ColumnIndex).Value.ToString
            If tableName.Equals("sys.Tables") Then
                bindGrid(cellValue, "*")
            End If
        End If


    End Sub
End Class
