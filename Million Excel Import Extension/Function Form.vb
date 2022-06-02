Imports System.IO
Imports ClosedXML.Excel
Imports ExcelDataReader
Imports System.Data.SqlClient

Public Class Function_Form
    Private Shared serverName As String
    Private Shared database As String
    Private Shared myConn As SqlConnection
    Private Shared statusConnection As Boolean
    Private Shared pwd_query As String
    Private Shared import_type As String
    Private Sub openExcelImport(sender As Object, e As EventArgs) Handles btnCashSales.Click, btnCreditNote.Click, btnDebitNote.Click, btnDeliveryOrder.Click, btnDeliveryReturn.Click, btnQuotation.Click, btnSalesInvoice.Click, btnSalesOrder.Click, btnStockAdjustment.Click, btnStockIssue.Click, btnStockReceive.Click, btnStockTrasfer.Click, btnPayBill.Click, btnReceivePayment.Click
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
        ElseIf formName.Equals("Cash Sales") Then
            form = New Cash_Sales_Form
        ElseIf formName.Equals("Debit Note") Then
            form = New Debit_Note_Form
        ElseIf formName.Equals("Credit Note") Then
            form = New Credit_Note_Form
        ElseIf formName.Equals("Delivery Return") Then
            form = New Delivery_Return_Form
        ElseIf formName.Equals("Stock Receive") Then
            form = New Stock_Receive_Form
        ElseIf formName.Equals("Stock Issue") Then
            form = New Stock_Issue_Form
        ElseIf formName.Equals("Stock Adjustment") Then
            form = New Stock_Adjustment_Form
        ElseIf formName.Equals("Stock Transfer") Then
            form = New Stock_Transfer_Form
        ElseIf formName.Equals("Receive Payment") Then
            form = New Receive_Payment
        ElseIf formName.Equals("Pay Bills") Then
            form = New Pay_Bills
        Else
            MsgBox("You have not enough privilege to access!", MsgBoxStyle.Critical)
            Return
        End If
        If Not Main_Form.validationProduct Then
            Threading.Thread.Sleep(3333)
        End If
        If Main_Form.getStatusConnection And Not Main_Form.getDatabase.Equals(String.Empty) Then
            Main_Form.setImport_type(formName)
            If formName.Contains("Stock") Then
                form.Text = formName + " (Stock)"
            ElseIf formName.Contains("Pay Bills") Then
                form.Text = formName + " (Creditor)"
            ElseIf formName.Contains("Receive Payment") Then
                form.Text = formName + " (Debtor)"
            Else
                form.Text = formName + " (Sales)"
            End If
            form.StartPosition = FormStartPosition.CenterScreen
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
        Return
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
                Dim saveFolder As String = Application.StartupPath + "\Report\"
                Dim saveFile As String = "ReportMELE_" + filename + "_" + Date.Now.Year.ToString + Date.Now.Month.ToString("00") + Date.Now.Day.ToString("00") + "_" + Date.Now.Hour.ToString("00") + Date.Now.Minute.ToString("00") + Date.Now.Second.ToString("00") + ".xlsx"
                If Not System.IO.Directory.Exists(saveFolder) Then
                    System.IO.Directory.CreateDirectory(saveFolder)
                End If
                sfd.FileName = saveFile
                workbook.SaveAs(saveFolder + saveFile)
                MsgBox("The report has been saved in " + saveFolder + saveFile, MsgBoxStyle.Information)

                'If sfd.ShowDialog() = DialogResult.OK Then

                'End If
            End Using
        End Using
    End Sub
    Public Shared Function validateExcelFormat(dgvExcel As DataGridView) As Boolean
        Dim integerColumn() As String = {"Credit Terms", "Batch Group", "Batch Code", "Quantity"}
        Dim doubleColumn() As String = {"Exchange Rate", "Deposit", "Mode of Payment 1 Amount", "Mode of Payment 2 Amount", "Mode of Payment 3 Amount", "Mode of Payment 4 Amount", "Debt Amount", "Price", "Gross Amount", "Discount Percentage 1", "Discount Percentage 2", "Discount Percentage 3", "Discount Amount Total", "Amount", "Tax Amount Adjustment", "Cost", "Unit Cost"}
        Dim dateColumn() As String = {"Date", "Delivery Date"}
        Dim result As Boolean = True
        Dim resultDate As Boolean = True
        Dim msg As String = ""
        For cell As Integer = 0 To dgvExcel.Rows(0).Cells.Count - 1
            Dim headerText As String = dgvExcel.Columns(cell).HeaderText.Trim
            If dateColumn.Contains(headerText) Then
                For row As Integer = 0 To dgvExcel.RowCount - 1
                    Dim cellText As String = dgvExcel.Rows(row).Cells(cell).Value.ToString.Trim
                    If Not cellText.Equals(String.Empty) Then
                        Try
                            Dim a = Convert.ToDateTime(cellText).ToString("yyyy-MM-dd HH:mm:ss")
                        Catch ex As Exception
                            result = False
                            msg += cellText + " (" + headerText + ")" + vbTab
                            resultDate = False
                        End Try
                    End If
                Next
            End If
            If doubleColumn.Contains(headerText) Then
                For row As Integer = 0 To dgvExcel.RowCount - 1
                    Dim cellText As String = dgvExcel.Rows(row).Cells(cell).Value.ToString.Trim
                    If Not cellText.Equals(String.Empty) Then
                        Try
                            Dim b = Convert.ToDecimal(cellText)
                        Catch ex As Exception
                            result = False
                            msg += cellText + " (" + headerText + ")" + vbTab
                        End Try
                    End If
                Next
            End If
            If integerColumn.Contains(headerText) Then
                For row As Integer = 0 To dgvExcel.RowCount - 1
                    Dim cellText As String = dgvExcel.Rows(row).Cells(cell).Value.ToString.Trim
                    If Not cellText.Equals(String.Empty) Then
                        Try
                            Dim c = Convert.ToInt32(cellText)
                        Catch ex As Exception
                            result = False
                            msg += cellText + " (" + headerText + ")" + vbTab
                        End Try
                    End If
                Next
            End If
        Next
        If result = False Then
            dgvExcel.DataSource = Nothing
            dgvExcel.Refresh()
            msg = "The following column(s) in excel files does not match the requirement!" + vbNewLine + msg
            If resultDate = False Then
                msg += vbNewLine + vbNewLine + "Tips: If date is not matching requirement, go to excel file, select all columns of date, and change format to Date."
            End If
            MsgBox(msg, MsgBoxStyle.Critical)
            Return False
        End If
        Return True
    End Function
    'Public Shared Function validateExcelDateFormat(dgvExcel As DataGridView, validateDateFormatArray As String()) As Boolean
    '    Dim result As Boolean = True
    '    For cell As Integer = 0 To dgvExcel.Rows(0).Cells.Count - 1
    '        Dim headerText As String = dgvExcel.Columns(cell).HeaderText.Trim
    '        If validateDateFormatArray.Contains(headerText) Then
    '            For row As Integer = 0 To dgvExcel.RowCount - 1
    '                Dim cellText As String = dgvExcel.Rows(row).Cells(cell).Value.ToString.Trim
    '                If Not cellText.Equals(String.Empty) Then
    '                    Try
    '                        Dim a = Convert.ToDateTime(cellText).ToString("yyyy-MM-dd HH:mm:ss")
    '                    Catch ex As Exception
    '                        result = False
    '                    End Try

    '                End If

    '            Next
    '        End If
    '    Next
    '    If result = False Then
    '        dgvExcel.DataSource = Nothing
    '        dgvExcel.Refresh()
    '        MsgBox("The imported excel format does not correct!", MsgBoxStyle.Critical)
    '        Return False
    '    End If
    '    Return True
    'End Function
    Public Shared Function getNull(ByVal type As Integer) As String
        'type: 0 = " ", 1 = "new Date", 2 = "current Date", 3="0"
        Dim result As String = ""
        If type = 0 Then
            result = "   "
        ElseIf type = 1 Then
            result = New Date(1900, 1, 1).ToString("yyyy-MM-dd HH:mm:ss")
        ElseIf type = 2 Then
            result = Date.Now.ToString("yyyy-MM-dd HH:mm:ss")
        ElseIf type = 3 Then
            result = "0"
        End If
        Return result
    End Function
    Public Shared Function convertDateFormat(ByVal datetime As String) As String
        Dim convertedDate As String = Convert.ToDateTime(datetime).ToString("yyyy-MM-dd HH:mm:ss")
        If convertedDate.Equals("2000-01-01 00:00:00") Then
            convertedDate = Convert.ToDateTime(New Date(1900, 1, 1)).ToString("yyyy-MM-dd HH:mm:ss")
        End If
        Return convertedDate
    End Function
    Public Shared Function queryValue(ByVal value) As String
        Return "'" + value.ToString + "',"
    End Function
    Public Shared Sub promptImportSuccess(ByVal insert As Integer, ByVal update As Integer)
        Dim result As String = "Data Import Sucessfully!"
        If insert > 0 Then
            result += vbNewLine + "Row Inserted: " + insert.ToString
        End If
        If update > 0 Then
            result += vbNewLine + "Row Updated: " + update.ToString
        End If
        MsgBox(result, MsgBoxStyle.Information)
    End Sub
    Public Shared Function repeatedExcelCell(ByVal dgvExcel As DataGridView, ByVal table_at As String, ByVal value As String, ByVal myrow As Integer) As Boolean
        Dim repeat_valid = False
        For i As Integer = 0 To dgvExcel.RowCount - 1
            If i <> myrow Then
                Dim value_in_row As String = dgvExcel.Rows(i).Cells(table_at).Value.ToString.Trim
                If value.ToString.Trim.Equals(value_in_row) Then
                    repeat_valid = True
                End If
            End If
        Next
        Return repeat_valid
    End Function
    Public Shared Function getCurrDoc(doc_type As String) As String
        serverName = Main_Form.getServerName
        database = Main_Form.getDatabase
        pwd_query = Main_Form.getPwd_query
        myConn = New SqlConnection("Data Source=" + serverName + ";" & "Initial Catalog=" + database + ";" + pwd_query)
        Dim curr_doc = ""
        Dim prefix As String = ""
        Dim curr_no As Integer = 0
        Dim doc_width As String = ""
        myConn.Open()
        Dim cmd As New SqlCommand
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "select prefix,curr_no,doc_width from defdocno WHERE doc_type='" + doc_type + "'"
        cmd.Connection = myConn
        Dim cmdreader As SqlDataReader = cmd.ExecuteReader
        While cmdreader.Read()
            prefix = cmdreader.GetValue(0).ToString.Trim
            curr_no = CInt(cmdreader.GetValue(1).ToString.Trim)
            doc_width = cmdreader.GetValue(2).ToString.Trim
        End While
        curr_doc = prefix + curr_no.ToString("D" + doc_width)
        myConn.Close()
        myConn.Open()
        Dim cmd2 As New SqlCommand
        cmd2.CommandType = CommandType.Text
        cmd2.CommandText = "UPDATE defdocno SET curr_no='" + (curr_no + 1).ToString + "' WHERE doc_type='" + doc_type + "'"
        cmd2.Connection = myConn
        cmd2.ExecuteNonQuery() '
        myConn.Close()
        Return curr_doc
    End Function
    Public Shared Function getCurrDoc(prefix As String, curr_no As Integer, doc_width As Integer) As String
        Dim prefix_temp As String = prefix
        Dim curr_no_temp As Integer = curr_no
        Dim doc_width_temp As Integer = doc_width

        serverName = Main_Form.getServerName
        database = Main_Form.getDatabase
        pwd_query = Main_Form.getPwd_query
        myConn = New SqlConnection("Data Source=" + serverName + ";" & "Initial Catalog=" + database + ";" + pwd_query)

        Try
            myConn.Open()
            Dim cmd As New SqlCommand
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "UPDATE defdocno SET curr_no='" + (curr_no + 1) + "' WHERE prefix='" + prefix_temp + "'"
            cmd.Connection = myConn
            cmd.ExecuteNonQuery() '
            myConn.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
        Return prefix_temp + curr_no_temp.ToString("D" + doc_width_temp.ToString)
    End Function
End Class