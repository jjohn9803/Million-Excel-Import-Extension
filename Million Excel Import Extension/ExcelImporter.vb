Imports System.Data.SqlClient
Imports ExcelDataReader
Imports System.IO
Imports ClosedXML.Excel

Public Class ExcelImporter
    Dim tables As DataTableCollection
    Private serverName As String
    Private database As String
    Private myConn As SqlConnection
    Private statusConnection As Boolean
    Private pwd_query As String
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Private Sub init()
        serverName = SQL_Connection_Form.serverName
        database = SQL_Connection_Form.database
        myConn = SQL_Connection_Form.myConn
        statusConnection = SQL_Connection_Form.statusConnection
        pwd_query = SQL_Connection_Form.pwd_query
    End Sub
    Private Sub cbSheet_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbSheet.SelectedIndexChanged
        Dim dt As DataTable = tables(cbSheet.SelectedItem.ToString())
        dgvExcel.DataSource = dt
    End Sub
    Private Sub txtFileName_MouseClick(sender As Object, e As MouseEventArgs) Handles txtFileName.MouseClick
        Try
            Using ofd As OpenFileDialog = New OpenFileDialog() With {.Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls"}
                If ofd.ShowDialog() = DialogResult.OK Then
                    txtFileName.Text = ofd.FileName
                    Using stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read)
                        Try
                            Using reader As IExcelDataReader = ExcelReaderFactory.CreateReader(stream)
                                Dim result As DataSet = reader.AsDataSet(New ExcelDataSetConfiguration() With {
                                                                         .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
                                                                         .UseHeaderRow = True}})
                                tables = result.Tables
                                cbSheet.Items.Clear()
                                For Each table As DataTable In tables
                                    cbSheet.Items.Add(table.TableName)
                                Next
                            End Using
                        Catch ex As Exceptions.HeaderException
                            MsgBox("The file is invalid! Please try another file!", MsgBoxStyle.Critical)
                            txtFileName.Text = String.Empty
                        End Try
                    End Using
                End If
            End Using
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub CreateTemplate_Click(sender As Object, e As EventArgs) Handles btnCreateTemplate.Click
        excelTemplate(cbType.Text.ToString)
    End Sub
    Private Sub excelTemplate(templateType As String)
        Using workbook As New XLWorkbook
            Dim worksheet As IXLWorksheet = workbook.Worksheets.Add(templateType)
            'Dim worksheetLR As Integer = 1
            Dim tableExcelSetting As DataTableCollection
            Try
                Using stream = File.Open(getMaintainSetting, FileMode.Open, FileAccess.Read)
                    Using reader As IExcelDataReader = ExcelReaderFactory.CreateReader(stream)
                        Dim result As DataSet = reader.AsDataSet(New ExcelDataSetConfiguration() With {
                                                                     .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
                                                                     .UseHeaderRow = True}})
                        tableExcelSetting = result.Tables
                        Dim excelUnique As New ArrayList
                        arrayExtendFromExcelSheet(tableExcelSetting, "Quotation", excelUnique, worksheet)
                        arrayExtendFromExcelSheet(tableExcelSetting, "Quotation Desc", excelUnique, worksheet)

                    End Using
                End Using
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical)
            End Try
            Using sfd As SaveFileDialog = New SaveFileDialog() With {.Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls"}
                If sfd.ShowDialog() = DialogResult.OK Then
                    worksheet.Columns.Width = 25
                    workbook.SaveAs(sfd.FileName)
                    MsgBox("File saved in " + sfd.FileName)
                End If
            End Using
        End Using
    End Sub
    Private Sub arrayExtendFromExcelSheet(tableExcelSetting As DataTableCollection, sheet As String, excelUnique As ArrayList, worksheet As IXLWorksheet)
        Dim dt As DataTable = tableExcelSetting(sheet)
        For Each row As DataRow In dt.Rows
            Dim sql_value = row(0).ToString
            Dim excel_value = row(1).ToString
            Dim default_value = row(2).ToString
            Dim formula = row(3).ToString
            If excel_value <> String.Empty Then
                Dim repeated As Boolean = False
                For Each al As String In excelUnique
                    If excel_value.Equals(al) Then
                        repeated = True
                    End If
                Next
                If repeated = False Then
                    excelUnique.Add(excel_value)
                    worksheet.Cell(1, excelUnique.Count).Value = excel_value
                End If
            End If
        Next
    End Sub
    Private Function getMaintainSetting() As String
        Return SQL_Connection_Form.returnUpperFolder(Application.StartupPath(), 2) + "maintain.xls"
    End Function
    Private Sub cbType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbType.SelectedIndexChanged
        If Not (cbType.SelectedItem.ToString.Equals(String.Empty)) Then
            btnCreateTemplate.Enabled = True
        Else
            btnCreateTemplate.Enabled = False
        End If
    End Sub

    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        Dim importType = "Quotation"
        Dim tableExcelSetting As DataTableCollection
        Try
            Using stream = File.Open(getMaintainSetting, FileMode.Open, FileAccess.Read)
                Using reader As IExcelDataReader = ExcelReaderFactory.CreateReader(stream)
                    Dim result As DataSet = reader.AsDataSet(New ExcelDataSetConfiguration() With {
                                                                 .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
                                                                 .UseHeaderRow = True}})
                    tableExcelSetting = result.Tables
                    Dim quotationTable As New ArrayList()
                    Dim quotationDescTable As New ArrayList()
                    For i As Integer = 0 To dgvExcel.ColumnCount - 1
                        Dim dgvHeader = dgvExcel.Columns(i).HeaderText
                        Dim dtTemp As DataTable = tableExcelSetting("Quotation")
                        For Each row As DataRow In dtTemp.Rows
                            Dim sql_value = row(0).ToString
                            Dim excel_value = row(1).ToString
                            If excel_value.Equals(dgvHeader) Then
                                quotationTable.Add(sql_value)
                            End If
                        Next
                        Dim dtTemp2 As DataTable = tableExcelSetting("Quotation Desc")
                        For Each row As DataRow In dtTemp2.Rows
                            Dim sql_value = row(0).ToString
                            Dim excel_value = row(1).ToString
                            If excel_value.Equals(dgvHeader) Then
                                quotationDescTable.Add(sql_value)
                            End If
                        Next
                    Next
                    Dim strs As String = ""
                    For Each str As String In quotationDescTable
                        strs += str + vbTab
                    Next
                    MsgBox(strs)
                End Using
            End Using
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
        'Try
        '    myConn = New SqlConnection("Data Source=" + serverName + ";" &
        '                            "Initial Catalog=master;" + pwd_query)
        '    myConn.Open()
        '    Using myConn
        '        Dim cmd As New SqlCommand("Select name from master.dbo.sysdatabases WHERE dbid > 4;", myConn)
        '        Using reader As SqlDataReader = cmd.ExecuteReader()
        '            cbDatabase.Items.Clear()
        '            While reader.Read()
        '                cbDatabase.Items.Add(reader.GetValue(0))
        '            End While
        '        End Using
        '        If (cbDatabase.Items.Count > 0) Then
        '            updateSQLStatus(1, "")
        '        Else
        '            updateSQLStatus(0, "Could not find any database from the server!")
        '        End If
        '    End Using
        '    myConn.Close()
        'Catch ex As Exception
        '    If (ex.GetHashCode = 49205706) Then
        '        updateSQLStatus(0, "Failed to connect the server!")
        '    Else
        '        MsgBox(ex.Message, MsgBoxStyle.Critical, Me.Text)
        '    End If
        'End Try
    End Sub
End Class
