Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnFileOpen_Click(sender As Object, e As EventArgs) Handles btnFileOpen.Click
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            Dim cn As OleDbConnection
            Dim cmd As OleDbCommand
            Dim da As OleDbDataAdapter
            Dim dt As DataTable
            da = New OleDbDataAdapter
            dt = New DataTable
            Dim fileLocation = OpenFileDialog1.FileName
            Debug.Write("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileLocation +
                                     ";Extension Properties='Excel 8.0; IMEX=1; HDR=Yes;'")
            'cn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileLocation +
            '                         ";Extension Properties='Excel 8.0; IMEX=1; HDR=Yes;'")
            cn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileLocation +
                                     ";Extended Properties='Excel 8.0; IMEX=1; HDR=Yes;'")

            cmd = New OleDbCommand("select * from [sheet1$]", cn)
            da.SelectCommand = cmd
            da.Fill(dt)

            DataGridView1.DataSource = dt
            lblFileLocation.Text = OpenFileDialog1.FileName

        End If
    End Sub

    Private Sub btnSql_Click(sender As Object, e As EventArgs) Handles btnSql.Click
        Dim myConn As SqlConnection
        Dim myCmd As SqlCommand
        Dim myReader As SqlDataReader
        Dim results As String
        'Create a Connection object.
        myConn = New SqlConnection("Initial Catalog=Northwind;" &
        "Data Source=DESKTOP-GAUUIC1;Initial Catalog=test;Integrated Security=True")
        'Create a Command object.
        myCmd = myConn.CreateCommand
        myCmd.CommandText = "SELECT name, description FROM Product"

        'Open the connection.
        myConn.Open()
        myReader = myCmd.ExecuteReader()
        'Concatenate the query result into a string.
        Do While myReader.Read()
            results = results & myReader.GetString(0) & vbTab &
            myReader.GetString(1) & vbLf
        Loop
        'Display results.
        MsgBox(results)
        'Close the reader and the database connection.
        myReader.Close()
        myConn.Close()
    End Sub
End Class
