﻿Imports System.IO
Imports ClosedXML.Excel
Imports ExcelDataReader

Public Class Function_Form
    Private Sub openExcelImport(sender As Object, e As EventArgs) Handles btnQuotation.Click, btnDeliveryOrder.Click, btnSalesOrder.Click, btnSalesInvoice.Click
        Dim formName = sender.text
        If Not Main_Form.getFeatures.Contains(sender.text) Then
            MsgBox("You have not enough privilege to access!", MsgBoxStyle.Critical)
            Return
        End If
        Dim form
        If formName.Equals("Quotation") Then
            form = New Quotation_Form
        ElseIf formName.Equals("Sales Order") Then
            form = New Sales_Order_Form
        ElseIf formName.Equals("Delivery Order") Then
            form = New Delivery_Order_Form
        ElseIf formName.Equals("Sales Invoice") Then
            form = New Sales_Invoice_Form
        End If
        If Main_Form.getStatusConnection And Not Main_Form.getDatabase.Equals(String.Empty) Then
            Main_Form.setImport_type(formName)
            form.Text = formName
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
            'workbook.SaveAs(filename)
            Using sfd As SaveFileDialog = New SaveFileDialog() With {.Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls"}
                Dim saveFile As String = Application.StartupPath + "\Report\ReportMELE_" + filename + "_" + Date.Now.Year.ToString + Date.Now.Month.ToString("00") + Date.Now.Day.ToString("00") + "_" + Date.Now.Hour.ToString("00") + Date.Now.Minute.ToString("00") + Date.Now.Second.ToString("00") + ".xlsx"
                sfd.FileName = saveFile
                workbook.SaveAs(saveFile)
                MsgBox("The report has been saved in " + saveFile, MsgBoxStyle.Information)

                'If sfd.ShowDialog() = DialogResult.OK Then

                'End If
            End Using
        End Using
    End Sub
    Public Shared Sub validateExcelDateFormat(dgvExcel As DataGridView, validateDateFormatArray As String())
        Dim result As Boolean = True
        For cell As Integer = 0 To dgvExcel.Rows(0).Cells.Count - 1
            Dim headerText As String = dgvExcel.Columns(cell).HeaderText.Trim
            If validateDateFormatArray.Contains(headerText) Then
                For row As Integer = 0 To dgvExcel.RowCount - 1
                    Dim cellText As String = dgvExcel.Rows(row).Cells(cell).Value.ToString.Trim
                    If Not cellText.Equals(String.Empty) Then
                        Try
                            Dim a = Convert.ToDateTime(cellText).ToString("dd-MMM-yy HH:mm:ss")
                        Catch ex As Exception
                            result = False
                        End Try

                    End If

                Next
            End If
        Next
        If result = False Then
            dgvExcel.DataSource = Nothing
            dgvExcel.Refresh()
            MsgBox("The imported excel format does not correct!", MsgBoxStyle.Critical)
        End If
    End Sub
    'Public Shared Function convertCharToHex16(text As String, char16 As Char) As String
    '    Dim convertedText As String = ""
    '    Dim bts As New List(Of Byte)
    '    For Each ch As Char In text
    '        If Not ch.Equals(char16) Then
    '            bts += Convert.ToString(Convert.ToInt32(ch), 16)
    '        End If
    '    Next
    '    Return System.Text.Encoding.ASCII.GetString(convertedText)
    'End Function
End Class