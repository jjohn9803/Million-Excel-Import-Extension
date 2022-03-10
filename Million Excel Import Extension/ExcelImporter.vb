Imports System.Data.SqlClient
Imports ExcelDataReader
Imports System.IO
Public Class ExcelImporter
    Dim tables As DataTableCollection

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Private Sub cbSheet_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbSheet.SelectedIndexChanged
        Dim dt As DataTable = tables(cbSheet.SelectedItem.ToString())
        dgvExcel.DataSource = dt
    End Sub
    Private Sub txtFileName_MouseClick(sender As Object, e As MouseEventArgs) Handles txtFileName.MouseClick
        Using ofd As OpenFileDialog = New OpenFileDialog() With {.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"}
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
End Class
