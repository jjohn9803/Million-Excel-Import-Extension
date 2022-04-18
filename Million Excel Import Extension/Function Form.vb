Imports ClosedXML.Excel
Public Class Function_Form
    Private Sub openExcelImport(sender As Object, e As EventArgs) Handles btnQuotation.Click, btnDeliveryOrder.Click, btnSalesOrder.Click
        Dim form
        If sender.text.Equals("Quotation") Then
            form = New Quotation_Form
        ElseIf sender.text.Equals("Sales Order") Then
            form = New Sales_Order_Form
        ElseIf sender.text.Equals("Delivery Order") Then
            form = New Delivery_Order_Form
        Else

        End If
        If SQL_Connection_Form.statusConnection And Not SQL_Connection_Form.database.Equals(String.Empty) Then
            SQL_Connection_Form.import_type = sender.text
            form.Text = sender.text
            form.ShowDialog()
        Else
            MsgBox("Please connect to server and database before get into it!", MsgBoxStyle.Critical)
        End If
        'Dim form As ExcelImporter = New ExcelImporter
        'SQL_Connection_Form.import_type = sender.text
        'form.Text = sender.text
        'form.ShowDialog()
    End Sub

    Friend Sub printExcelResult(filename As String, queryTable As ArrayList, value_arraylist As ArrayList, sql_format_arraylist As ArrayList, dgvExcel As DataGridView)
        Using workbook As New XLWorkbook
            Dim sheets As New ArrayList

            For i As Integer = 0 To queryTable.Count - 1
                Dim sheetName As String = queryTable(i)(0)
                Dim worksheet As IXLWorksheet = workbook.Worksheets.Add(sheetName)
                Dim writableRow = 1

                For row As Integer = 0 To dgvExcel.RowCount - 1
                    If Not (value_arraylist(i)(row)(0).Equals("{INVALID ARRAY}")) Then
                        For j As Integer = 0 To value_arraylist(i)(row).count - 1
                            worksheet.Cell(writableRow, (j + 1)).Value = sql_format_arraylist(i)(j)
                            worksheet.Cell(writableRow, (j + 1)).Style.Font.Bold = True
                        Next
                        writableRow += 1
                        For j As Integer = 0 To value_arraylist(i)(row).count - 1
                            worksheet.Cell(writableRow, (j + 1)).Value = value_arraylist(i)(row)(j)
                        Next
                        writableRow += 2
                    End If
                Next
                worksheet.Columns.Width = 25
            Next
            'C:\Users\RBADM07\Desktop\Generated Result.xlsx
            workbook.SaveAs(filename)
            'Using sfd As SaveFileDialog = New SaveFileDialog() With {.Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls"}
            '    sfd.FileName = "Generated Result"
            '    If sfd.ShowDialog() = DialogResult.OK Then
            '        workbook.SaveAs(sfd.FileName)
            '    End If
            'End Using
        End Using
    End Sub
End Class