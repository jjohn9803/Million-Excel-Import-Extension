Imports System.Data.SqlClient
Imports ExcelDataReader
Imports System.IO
Imports OfficeOpenXml
Public Class ExcelImporter
    Dim tables As DataTableCollection

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Private Sub cbSheet_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbSheet.SelectedIndexChanged
        Dim dt As DataTable = tables(cbSheet.SelectedItem.ToString())
        dgvExcel.DataSource = dt
    End Sub
    Private Sub txtFileName_MouseClick(sender As Object, e As MouseEventArgs) Handles txtFileName.MouseClick
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
    End Sub

    Private Sub CreateTemplate_Click(sender As Object, e As EventArgs) Handles btnCreateTemplate.Click
        Dim excelColumn As String()
        readExcelSetting()
        'Using excelPackage As New ExcelPackage = New ExcelPackage(New FileInfo(""))
        '    Dim workSheet As ExcelWorksheet = excelPackage.Workbook.Worksheets.Add("Sheet1")
        '    workSheet.Cells.LoadFromDataReader()
        'End Using


    End Sub
    Private Sub readExcelSetting()
        Dim array As String()
        Dim tableExcelSetting As DataTableCollection
        Try
            Using stream = File.Open(getMaintainSetting, FileMode.Open, FileAccess.Read)
                Using reader As IExcelDataReader = ExcelReaderFactory.CreateReader(stream)
                    Dim result As DataSet = reader.AsDataSet(New ExcelDataSetConfiguration() With {
                                                                     .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
                                                                     .UseHeaderRow = True}})
                    tableExcelSetting = result.Tables
                    Dim dt As DataTable = tableExcelSetting("Quotation")
                    For Each row As DataRow In dt.Rows
                        Dim sql_value = row(0).ToString
                        Dim excel_value = row(1).ToString
                        If excel_value <> String.Empty Then
                            Dim repeated As Boolean = False
                            For Each previous_array As String In array.Where(array.Length > 0)
                                If previous_array.Equals(excel_value) Then
                                    repeated = True
                                    MsgBox("lmao" + excel_value)
                                End If
                            Next
                            If repeated = False Then
                                array(array.Length) = excel_value
                            End If
                        End If
                    Next
                End Using
            End Using
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub
    Private Function getMaintainSetting() As String
        Return SQL_Connection_Form.returnUpperFolder(Application.StartupPath(), 2) + "maintain.xls"
    End Function
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedItem.ToString.Equals("Quotation") Then
            btnCreateTemplate.Enabled = True
        End If
    End Sub
End Class
