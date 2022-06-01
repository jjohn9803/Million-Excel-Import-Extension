Imports System.Data.SqlClient
Imports System.IO
Imports ClosedXML.Excel
Imports ExcelDataReader
Public Class Pay_Bills
    Dim tables As DataTableCollection
    Private serverName As String
    Private database As String
    Private myConn As SqlConnection
    Private statusConnection As Boolean
    Private pwd_query As String
    Private import_type As String
    Private Sub Pay_Bills_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        init()
    End Sub
    Private Sub init()
        serverName = Main_Form.getServerName
        database = Main_Form.getDatabase
        pwd_query = Main_Form.getPwd_query
        myConn = New SqlConnection("Data Source=" + serverName + ";" & "Initial Catalog=" + database + ";" + pwd_query)
        statusConnection = Main_Form.getStatusConnection
        import_type = Main_Form.getImport_type
        txtType.Text = import_type
    End Sub
    Private Sub cbSheet_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbSheet.SelectedIndexChanged
        Dim dt As DataTable = tables(cbSheet.SelectedItem.ToString())
        dgvExcel.DataSource = dt
        For Each column As DataGridViewColumn In dgvExcel.Columns
            column.SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        If Function_Form.validateExcelFormat(dgvExcel) = False Then
            txtFileName.Text = String.Empty
            cbSheet.Items.Clear()
            dgvExcel.DataSource = Nothing
            dgvExcel.Refresh()
            btnImport.Enabled = False
        Else
            btnImport.Enabled = True
        End If
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
                                dgvExcel.DataSource = Nothing
                                dgvExcel.Refresh()
                                btnImport.Enabled = False
                            End Using
                        Catch ex As Exceptions.HeaderException
                            MsgBox("The file is invalid! Please try another file!", MsgBoxStyle.Critical)
                            txtFileName.Text = String.Empty
                            cbSheet.Items.Clear()
                            dgvExcel.DataSource = Nothing
                            dgvExcel.Refresh()
                            btnImport.Enabled = False
                        End Try
                    End Using
                End If
            End Using
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            txtFileName.Text = String.Empty
            cbSheet.Items.Clear()
            dgvExcel.DataSource = Nothing
            dgvExcel.Refresh()
            btnImport.Enabled = False
        End Try
    End Sub
    Private Function getMaintainSetting() As String
        Return "maintain.xls"
    End Function
    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        Dim importType = "Pay Bills"
        Dim tableExcelSetting As DataTableCollection
        Dim stream As FileStream
        Try
            stream = File.Open(getMaintainSetting, FileMode.Open, FileAccess.Read)
        Catch ex As Exception
            MsgBox(ex.Message + vbNewLine + ex.StackTrace, MsgBoxStyle.Critical)
            Return
        End Try
        Using reader As IExcelDataReader = ExcelReaderFactory.CreateReader(stream)
            Dim result As DataSet = reader.AsDataSet(New ExcelDataSetConfiguration() With {
                                                                                .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
                                                                                .UseHeaderRow = True}})
            tableExcelSetting = result.Tables
            Dim queryTable As New ArrayList
            For i = 0 To 2
                queryTable.Add(New ArrayList)
            Next
            queryTable(0).add("PB AP") '0
            queryTable(0).add("ap") '1
            queryTable(1).add("PB GL")
            queryTable(1).add("gl")
            queryTable(2).add("PB GL Off")
            queryTable(2).add("gloff")
            quotationWriteIntoSQL(tableExcelSetting, queryTable)
        End Using
    End Sub
    Private Sub quotationWriteIntoSQL(tableExcelSetting As DataTableCollection, queryTable As ArrayList)
        Dim value_arraylist = New ArrayList
        Dim sql_format_arraylist = New ArrayList
        Dim excel_format_arraylist = New ArrayList
        Dim default_arraylist = New ArrayList
        Dim formula_arraylist = New ArrayList
        Dim data_type_arraylist = New ArrayList
        For j As Integer = 0 To queryTable.Count - 1
            value_arraylist.Add(New ArrayList)
            sql_format_arraylist.Add(New ArrayList)
            excel_format_arraylist.Add(New ArrayList)
            default_arraylist.Add(New ArrayList)
            formula_arraylist.Add(New ArrayList)
            data_type_arraylist.Add(New ArrayList)
            Dim dtTemp As DataTable = tableExcelSetting(queryTable(j)(0))
            For Each row As DataRow In dtTemp.Rows
                Dim sql_value = row(0).ToString
                Dim excel_value = row(1).ToString
                Dim default_value = row(2).ToString
                Dim formula_value = row(3).ToString
                Dim data_type = row(4).ToString
                sql_format_arraylist(j).add(sql_value)
                excel_format_arraylist(j).add(excel_value)
                default_arraylist(j).add(default_value)
                formula_arraylist(j).add(formula_value)
                data_type_arraylist(j).add(data_type)
            Next
        Next
        For m As Integer = 0 To queryTable.Count - 1
            Dim query As String = "INSERT INTO " + queryTable(m)(1) + " ("
            Dim values As String = ""
            For n As Integer = 0 To sql_format_arraylist(m).Count - 1
                If Not (default_arraylist(m)(n).Equals("PK_int")) Then
                    values += "," + sql_format_arraylist(m)(n)
                End If
            Next
            query += values.Substring(1, values.Length - 1) + ") VALUES ("
            queryTable(m).add(query)
        Next
        For i As Integer = 0 To queryTable.Count - 1
            For row As Integer = 0 To dgvExcel.RowCount - 1
                Using r As DataGridViewRow = dgvExcel.Rows(row)
                    Dim Queryable = False
                    value_arraylist(i).Add(New ArrayList)
                    For u As Integer = 0 To sql_format_arraylist(i).Count - 1
                        Dim sql_temp = sql_format_arraylist(i)(u).ToString.Trim
                        Dim data_type_temp = data_type_arraylist(i)(u).ToString.Trim
                        Dim default_temp = default_arraylist(i)(u).ToString.Trim
                        Dim formula_temp = formula_arraylist(i)(u).ToString.Trim
                        Dim add_value As Boolean = False
                        Dim value = ""
                        For q As Integer = 0 To r.Cells.Count - 1
                            Dim cellValue As Object = r.Cells(q).Value.ToString.Trim
                            Dim headerText As String = dgvExcel.Columns(q).HeaderText.Trim
                            Dim excel_temp = excel_format_arraylist(i)(u).ToString.Trim
                            If headerText.Equals(excel_temp) Then
                                If Not (cellValue.Equals(String.Empty)) Then
                                    q = r.Cells.Count
                                    value_arraylist(i)(row).add(cellValue)
                                    add_value = True
                                    value = cellValue
                                    If data_type_temp.Contains("PK") Then
                                        Queryable = True
                                    End If
                                End If
                            End If
                        Next
                        If add_value = False Then
                            If Not (default_temp.Equals(String.Empty)) Then
                                value = "{DEFAULT_VALUE}"
                                value_arraylist(i)(row).add("{DEFAULT_VALUE}")
                            ElseIf Not (formula_temp.Equals(String.Empty)) Then
                                value = "{FORMULA_VALUE}"
                                value_arraylist(i)(row).add("{FORMULA_VALUE}")
                            Else
                                If data_type_temp.ToString.Contains("char") Or data_type_temp.ToString.Contains("text") Then
                                    value = "   "
                                    value_arraylist(i)(row).add("   ")
                                ElseIf data_type_temp.ToString.Contains("date") Or data_type_temp.ToString.Contains("time") Then
                                    value = Function_Form.getNull(1)
                                    value_arraylist(i)(row).add(Function_Form.getNull(1))
                                Else
                                    value = "0"
                                    value_arraylist(i)(row).add("0")
                                End If
                            End If
                        End If
                    Next
                    If Queryable = False Then
                        value_arraylist(i)(row).clear()
                        value_arraylist(i)(row).add("{INVALID ARRAY}")
                    End If
                End Using
            Next
        Next
        'For i As Integer = 0 To queryTable.Count - 1
        '    For row As Integer = 0 To dgvExcel.RowCount - 1
        '        Dim strs = ""
        '        For Each str As String In value_arraylist(i)(row)
        '            strs += str + vbTab
        '        Next
        '        MsgBox("Row " + row.ToString + vbNewLine + strs)
        '    Next
        'Next

        Dim rangeQuo As New ArrayList
        Dim rangeEnd = -1
        For i = 0 To dgvExcel.RowCount - 1
            Dim str = value_arraylist(0)(i)(0)
            If Not str.Equals("{INVALID ARRAY}") Then
                rangeEnd = i
            End If
            rangeQuo.Add(rangeEnd.ToString + "." + i.ToString)
        Next

        For default_value_checker As Integer = 0 To 10
            For i As Integer = 0 To queryTable.Count - 1
                For row As Integer = 0 To dgvExcel.RowCount - 1
                    Using r As DataGridViewRow = dgvExcel.Rows(row)
                        If Not value_arraylist(i)(row)(0).Equals("{INVALID ARRAY}") Then
                            For g As Integer = 0 To value_arraylist(i)(row).count - 1
                                Dim value_temp As String = value_arraylist(i)(row)(g).ToString.Trim
                                If value_temp.Equals("{DEFAULT_VALUE}") Then
                                    Dim default_temp = default_arraylist(i)(g).ToString.Trim
                                    If default_temp.Equals("FK") Then
                                        If row > 0 Then
                                            Dim fk_value_temp = ""
                                            For h = row To 0 Step -1
                                                fk_value_temp = value_arraylist(i)(h)(g).ToString.Trim
                                                If Not fk_value_temp.Equals(String.Empty) AndAlso Not fk_value_temp.Equals("{DEFAULT_VALUE}") Then
                                                    Exit For
                                                End If
                                            Next
                                            If fk_value_temp.Equals("{DEFAULT_VALUE}") Then
                                                fk_value_temp = "   "
                                            End If
                                            'MsgBox(sql_format_arraylist(i)(g))
                                            value_arraylist(i)(row)(g) = fk_value_temp
                                        Else
                                            Dim data_type_temp = data_type_arraylist(i)(g).ToString.Trim
                                            If data_type_temp.ToString.Contains("char") Or data_type_temp.ToString.Contains("text") Then
                                                value_arraylist(i)(row)(g) = "   "
                                            ElseIf data_type_temp.ToString.Contains("date") Or data_type_temp.ToString.Contains("time") Then
                                                value_arraylist(i)(row)(g) = Function_Form.getNull(1)
                                            Else
                                                value_arraylist(i)(row)(g) = "0"
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End Using
                Next
            Next
        Next

        For default_value_checker As Integer = 0 To 10
            For i As Integer = 0 To queryTable.Count - 1
                For row As Integer = 0 To dgvExcel.RowCount - 1
                    Using r As DataGridViewRow = dgvExcel.Rows(row)
                        If Not value_arraylist(i)(row)(0).Equals("{INVALID ARRAY}") Then
                            For g As Integer = 0 To value_arraylist(i)(row).count - 1
                                Dim value_temp As String = value_arraylist(i)(row)(g).ToString.Trim
                                If value_temp.Equals("{DEFAULT_VALUE}") Then
                                    Dim default_temp = default_arraylist(i)(g).ToString.Trim
                                    If default_temp.Contains("getstringonly_") Then
                                        Dim default_index As Integer = Convert.ToInt32(default_temp.Substring(14))
                                        Dim stringonly_char As Char() = value_arraylist(i)(row)(default_index).ToString.ToCharArray()
                                        Dim stringonly_str = ""
                                        For Each ch As Char In stringonly_char
                                            If Not Char.IsDigit(ch) Then
                                                stringonly_str += ch
                                            End If
                                        Next
                                        value_arraylist(i)(row)(g) = stringonly_str
                                    ElseIf default_temp.Contains("increase") Then
                                        Dim getSource = rangeQuo(row).ToString.Split(".")(0)
                                        Dim order = 1
                                        Dim increasement = CInt(default_temp.ToString.Split(".")(1))
                                        For nm = 0 To rangeQuo.Count - 1
                                            If rangeQuo(nm).Split(".")(0).ToString.Trim.Equals(getSource) Then
                                                If rangeQuo(nm).Split(".")(1).ToString.Trim.Equals(row.ToString) Then
                                                    Exit For
                                                End If
                                                order += increasement
                                            End If
                                        Next
                                        value_arraylist(i)(row)(g) = order.ToString
                                    ElseIf default_temp.Contains("PK_int") Then
                                        value_arraylist(i)(row)(g) = "{._!@#$%^&*()}"
                                    ElseIf default_temp.Contains("=") Then

                                        Dim source_sql_value = default_temp.Split("=")(0)
                                        Dim destination_table = (default_temp.Split("=")(1)).Split(".")(0)
                                        Dim destination_sql_value = (default_temp.Split("=")(1)).Split(".")(1)

                                        Dim source_sql_index = sql_format_arraylist(i).IndexOf(source_sql_value)
                                        Dim source_value = value_arraylist(i)(row)(source_sql_index)
                                        Dim result_value = ""
                                        'MsgBox("SELECT " + destination_sql_value + " FROM " + destination_table + " WHERE " + source_sql_value + "='" + source_value + "'")
                                        'MsgBox(source_value)
                                        If source_value.ToString.Trim.Equals(String.Empty) Then
                                            value_arraylist(i)(row)(g) = 0
                                        Else
                                            init()
                                            'myConn = New SqlConnection("Data Source=" + serverName + ";" & "Initial Catalog=" + database + ";" + pwd_query)

                                            Dim command = New SqlCommand("SELECT " + destination_sql_value + " FROM " + destination_table + " WHERE " + source_sql_value + "='" + source_value + "'", myConn)
                                            myConn.Open()
                                            Dim reader As SqlDataReader = command.ExecuteReader
                                            While reader.Read()
                                                result_value += reader.GetValue(0).ToString
                                            End While

                                            value_arraylist(i)(row)(g) = result_value

                                            'MsgBox("I guess I put it as " + result_value + " with " + source_value)
                                            myConn.Close()
                                        End If
                                    Else
                                        value_arraylist(i)(row)(g) = default_temp
                                    End If
                                End If
                            Next
                        End If
                    End Using
                Next
            Next
        Next


        'Hardcode Formula
        For row As Integer = 0 To dgvExcel.RowCount - 1
            If Not value_arraylist(0)(row)(0).Equals("{INVALID ARRAY}") Then
                Dim myTarget As New ArrayList
                For Each target In rangeQuo
                    If target.ToString.Split(".")(0).Equals(row.ToString) Then
                        Dim targetRow = CInt(target.ToString.Split(".")(1))
                        myTarget.Add(targetRow)
                    End If
                Next

                'ap
                'paid
                If value_arraylist(0)(row)(22).Equals("{FORMULA_VALUE}") Then
                    Dim paid As Double = 0
                    For Each targetRow As Integer In myTarget
                        Dim paid_amt = value_arraylist(2)(targetRow)(11)
                        paid += paid_amt
                    Next
                    value_arraylist(0)(row)(22) = Math.Round(paid, 2)
                End If
            End If
        Next
        'End Hardcode Formula

        'hardcore exist checker
        Dim execute_valid As Boolean = True
        Dim exist_result As String = ""
        init()
        For row As Integer = 0 To dgvExcel.RowCount - 1
            Dim table As String
            Dim value_name As String
            Dim value As String
            'ap
            If Not value_arraylist(0)(row)(0).Equals("{INVALID ARRAY}") Then
                Dim myTarget As New ArrayList
                For Each target In rangeQuo
                    If target.ToString.Split(".")(0).Equals(row.ToString) Then
                        Dim targetRow = CInt(target.ToString.Split(".")(1))
                        myTarget.Add(targetRow)
                    End If
                Next

                'gldata.accno / exist
                table = "gldata"
                value_name = "accno"
                value = dgvExcel.Rows(row).Cells("Bank / Cash Acc No.").Value.ToString.Trim
                If Not value.Trim.Equals(String.Empty) Then
                    If Not existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                End If

                'supplier.suppcode / exist
                table = "supplier"
                value_name = "suppcode"
                value = value_arraylist(0)(row)(0)
                If Not value.Trim.Equals(String.Empty) Then
                    If Not existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                End If

                'currency.curr_code / exist
                table = "currency"
                value_name = "curr_code"
                value = value_arraylist(0)(row)(18)
                If Not value.Trim.Equals(String.Empty) Then
                    If Not existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                End If

                'project.projcode / exist
                table = "project"
                value_name = "projcode"
                value = value_arraylist(0)(row)(31)
                If Not value.Trim.Equals(String.Empty) Then
                    If Not existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                End If

                '//
                For Each targetRow As Integer In myTarget
                    'gloff
                    'pkdoc_type,pkdoc_date,pkfx_rate
                    myConn.Open()
                    Dim command = New SqlCommand("select doc_type,doc_date,fx_rate from pinv WHERE doc_no='" + value_arraylist(2)(targetRow)(6) + "'", myConn)
                    Dim reader As SqlDataReader = command.ExecuteReader
                    While reader.Read()
                        value_arraylist(2)(targetRow)(5) = reader.GetValue(0).ToString.Trim()
                        value_arraylist(2)(targetRow)(8) = reader.GetValue(1).ToString.Trim()
                        value_arraylist(2)(targetRow)(10) = reader.GetValue(2).ToString.Trim()
                    End While
                    myConn.Close()

                    'pinv.doc_no / duplicate
                    table = "pinv"
                    value_name = "doc_no"
                    value = value_arraylist(2)(targetRow)(6)
                    If Not value.Trim.Equals(String.Empty) Then
                        myConn.Open()
                        Dim sncommand = New SqlCommand("select * from " + table + " WHERE suppcode='" + value_arraylist(0)(row)(0) + "' AND doc_no='" + value + "'", myConn)
                        Dim snreader As SqlDataReader = sncommand.ExecuteReader
                        Dim exist_acc = False
                        While snreader.Read()
                            exist_acc = True
                        End While
                        If Not exist_acc Then
                            execute_valid = False
                            exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                        End If
                        myConn.Close()
                    End If

                    'gloff.pkdoc_no / duplicate
                    table = "gloff"
                    value_name = "pkdoc_no"
                    value = value_arraylist(2)(targetRow)(6)
                    If Not value.Trim.Equals(String.Empty) Then
                        If existed_checker(table, value_name, value) Then
                            execute_valid = False
                            exist_result += value_name + " '" + value + "' has been used in the database (" + table + ")!" + vbNewLine
                        End If
                    End If
                Next
            End If

            'gl
            If Not value_arraylist(1)(row)(0).Equals("{INVALID ARRAY}") Then
                'batchno
                If value_arraylist(1)(row)(20).length = 2 Then
                    Dim date1 = value_arraylist(1)(row)(4)
                    Dim batchno = value_arraylist(1)(row)(20)
                    value_arraylist(1)(row)(20) = Year(date1).ToString.Substring(2) + Month(date1).ToString("00") + batchno
                End If

                'desc1 supplier
                myConn.Open()
                Dim command = New SqlCommand("select name from supplier WHERE suppcode='" + value_arraylist(0)(row)(0) + "'", myConn)
                Dim reader As SqlDataReader = command.ExecuteReader
                While reader.Read()
                    value_arraylist(1)(row)(8) = reader.GetValue(0).ToString.Trim()
                End While
                myConn.Close()

                'desc2 bank/cash
                myConn.Open()
                Dim command2 = New SqlCommand("select accdesp from gldata WHERE accno='" + dgvExcel.Rows(row).Cells("Bank / Cash Acc No.").Value.ToString.Trim + "'", myConn)
                Dim reader2 As SqlDataReader = command2.ExecuteReader
                While reader2.Read()
                    value_arraylist(1)(row)(9) = reader2.GetValue(0).ToString.Trim()
                End While
                myConn.Close()
            End If
        Next
        'pay not exceed than total amt
        For row As Integer = 0 To dgvExcel.RowCount - 1
            If Not value_arraylist(0)(row)(0).Equals("{INVALID ARRAY}") Then
                Dim startRange As Integer = -1
                Dim endRange As Integer = -1
                For Each range As String In rangeQuo
                    If range.Split(".")(0).ToString.Trim.Equals(row.ToString.Trim) Then
                        startRange = CInt(range.Split(".")(0).ToString.Trim)
                        endRange = CInt(range.Split(".")(1).ToString.Trim)
                    End If
                Next
                Dim amt As Double = value_arraylist(0)(row)(21)
                Dim pay = 0
                For i As Integer = startRange To endRange
                    pay += CDbl(value_arraylist(2)(i)(11))
                Next
                If pay > amt Then
                    execute_valid = False
                    exist_result += "Payment cannot exceed total amount! (Row " + (row + 1).ToString + ")" + vbNewLine
                End If
            End If
        Next

        If execute_valid = False Then
            MsgBox(exist_result + vbNewLine + "The operation has been stopped!", MsgBoxStyle.Exclamation)
            Return
        End If

        'endhardcore exist checker
        init()
        Dim confirmImport As DialogResult = MsgBox("Are you sure to import data?", MsgBoxStyle.YesNo, "")
        If confirmImport = DialogResult.No Then
            Return
        End If

        Dim rowInsertNum = 0
        Dim rowUpdateNum = 0

        Dim gloffseq As New ArrayList
        'gl
        For row As Integer = 0 To dgvExcel.RowCount - 1
            If Not value_arraylist(0)(row)(0).Equals("{INVALID ARRAY}") Then
                Dim query_temp As New ArrayList
                query_temp.Add(value_arraylist(0)(row)(19)) 'fx_rate
                query_temp.Add(value_arraylist(0)(row)(18)) 'curr_code
                query_temp.Add(value_arraylist(1)(row)(20)) 'batchno
                query_temp.Add(value_arraylist(0)(row)(31)) 'projcode
                query_temp.Add(Function_Form.getNull(0)) 'deptcode
                query_temp.Add(Function_Form.getNull(0)) 'taxcode
                query_temp.Add(Function_Form.getNull(3)) 'taxable
                query_temp.Add(Function_Form.getNull(3)) 'fx_taxable
                query_temp.Add(Function_Form.getNull(3)) 'link_seq
                query_temp.Add("P") 'billtype
                query_temp.Add(value_arraylist(0)(row)(13)) 'remark1
                query_temp.Add(value_arraylist(0)(row)(14)) 'remark2
                query_temp.Add(value_arraylist(0)(row)(15)) 'cheque_no
                If value_arraylist(0)(row)(15).ToString.Trim.Equals(String.Empty) Then
                    query_temp.Add(Function_Form.getNull(1)) 'chqrc_date
                Else
                    query_temp.Add(Function_Form.convertDateFormat(value_arraylist(0)(row)(4))) 'chqrc_date
                End If
                query_temp.Add(Function_Form.getNull(1)) 'koff_date
                query_temp.Add(Function_Form.getNull(1)) 'recon_date
                query_temp.Add(Function_Form.getNull(3)) 'recon_flag
                query_temp.Add(Function_Form.getNull(0)) 'accmgr_id
                query_temp.Add(Function_Form.getNull(0)) 'lkdoc_type
                query_temp.Add(Function_Form.getNull(0)) 'lkdoc_no
                query_temp.Add(Function_Form.getNull(3)) 'lkseq
                query_temp.Add(Function_Form.getNull(3)) 'lock
                query_temp.Add(Function_Form.getNull(3)) 'void
                query_temp.Add(Function_Form.getNull(3)) 'exported
                query_temp.Add("CRP") 'entry
                query_temp.Add(Function_Form.getNull(3)) 'fastentry
                query_temp.Add(Function_Form.getNull(3)) 'followdesp
                query_temp.Add(Function_Form.getNull(3)) 'tsid
                query_temp.Add(Function_Form.getNull(0)) 'spcode
                query_temp.Add(Function_Form.getNull(1)) 'taxdate
                query_temp.Add(Function_Form.getNull(1)) 'taxdate_bt
                query_temp.Add(value_arraylist(1)(row)(49)) 'createdby
                query_temp.Add(value_arraylist(1)(row)(50)) 'updatedby
                query_temp.Add(Function_Form.getNull(2)) 'createdate
                query_temp.Add(Function_Form.getNull(2)) 'lastupdate
                Dim cmd_last = ""
                For j = 0 To query_temp.Count - 1
                    cmd_last += "'" + query_temp(j) + "',"
                Next
                cmd_last = cmd_last.Substring(0, cmd_last.Length - 1) + ")"

                Dim queryAL As New ArrayList
                Dim amount As Double = 0
                Dim debit As Double = 0
                Dim credit As Double = 0
                Dim fx_rate As Double = 0
                Dim fx_amount As Double = 0
                Dim fx_debit As Double = 0
                Dim fx_credit As Double = 0
                Dim curr_doc As String = Function_Form.getCurrDoc("GL")
                gloffseq.Add(value_arraylist(0)(row)(0) + "." + curr_doc + ".")

                'line 1
                queryAL.Clear()
                queryAL.Add(value_arraylist(0)(row)(0)) 'accno
                queryAL.Add("GL") 'doc_type
                queryAL.Add(curr_doc) 'doc_no
                queryAL.Add(1) 'seq
                queryAL.Add(Function_Form.convertDateFormat(value_arraylist(0)(row)(4))) 'doc_date
                queryAL.Add(value_arraylist(0)(row)(6)) 'refno
                queryAL.Add(value_arraylist(0)(row)(7)) 'refno2
                queryAL.Add(Function_Form.getNull(0)) 'refno3
                queryAL.Add(value_arraylist(1)(row)(8)) 'desp
                queryAL.Add(value_arraylist(0)(row)(10)) 'desp2
                queryAL.Add(Function_Form.getNull(0)) 'desp3
                queryAL.Add(Function_Form.getNull(0)) 'desp4

                amount = Math.Round(CDbl(value_arraylist(0)(row)(21)) * -1, 2)
                queryAL.Add(amount) 'amount
                debit = 0
                credit = 0
                If amount < 0 Then
                    credit = amount * -1
                Else
                    debit = amount
                End If
                queryAL.Add(debit) 'debit
                queryAL.Add(credit) 'credit
                fx_rate = CDbl(value_arraylist(0)(row)(19))
                fx_amount = Math.Round(amount * fx_rate * -1, 2)
                fx_debit = 0
                fx_credit = 0
                If fx_amount < 0 Then
                    fx_credit = fx_amount * -1
                Else
                    fx_debit = fx_amount
                End If
                fx_debit = fx_debit
                fx_credit = fx_credit

                queryAL.Add(fx_amount) 'fx_amount
                queryAL.Add(fx_debit) 'fx_debit
                queryAL.Add(fx_credit) 'fx_credit

                Dim execmd As String = queryTable(1)(2)
                For j = 0 To queryAL.Count - 1
                    Try
                        execmd += "'" + queryAL(j) + "',"
                    Catch ex As InvalidCastException
                        execmd += "'" + queryAL(j).ToString + "',"
                    End Try
                Next
                execmd = execmd + cmd_last
                myConn.Open()
                Dim cmd1 = New SqlCommand(execmd, myConn)
                cmd1.ExecuteNonQuery()
                rowInsertNum += 1
                myConn.Close()

                'line2
                queryAL.Clear()
                queryAL.Add(dgvExcel.Rows(row).Cells("Bank / Cash Acc No.").Value.ToString.Trim) 'accno
                queryAL.Add("GL") 'doc_type
                queryAL.Add(curr_doc) 'doc_no
                queryAL.Add(2) 'seq
                queryAL.Add(Function_Form.convertDateFormat(value_arraylist(0)(row)(4))) 'doc_date
                queryAL.Add(value_arraylist(0)(row)(6)) 'refno
                queryAL.Add(value_arraylist(0)(row)(7)) 'refno2
                queryAL.Add(Function_Form.getNull(0)) 'refno3
                queryAL.Add(value_arraylist(1)(row)(9)) 'desp
                queryAL.Add(value_arraylist(0)(row)(10)) 'desp2
                queryAL.Add(Function_Form.getNull(0)) 'desp3
                queryAL.Add(Function_Form.getNull(0)) 'desp4

                amount = Math.Round(CDbl(value_arraylist(0)(row)(21)), 2)
                queryAL.Add(amount) 'amount
                debit = 0
                credit = 0
                If amount < 0 Then
                    credit = amount * -1
                Else
                    debit = amount
                End If
                queryAL.Add(debit) 'debit
                queryAL.Add(credit) 'credit
                fx_rate = CDbl(value_arraylist(0)(row)(19))
                fx_amount = Math.Round(amount * fx_rate, 2)
                fx_debit = 0
                fx_credit = 0
                If fx_amount < 0 Then
                    fx_credit = fx_amount * -1
                Else
                    fx_debit = fx_amount
                End If
                fx_debit = fx_debit
                fx_credit = fx_credit

                queryAL.Add(fx_amount) 'fx_amount
                queryAL.Add(fx_debit) 'fx_debit
                queryAL.Add(fx_credit) 'fx_credit

                execmd = queryTable(1)(2)
                For j = 0 To queryAL.Count - 1
                    Try
                        execmd += "'" + queryAL(j) + "',"
                    Catch ex As InvalidCastException
                        execmd += "'" + queryAL(j).ToString + "',"
                    End Try
                Next
                execmd = execmd + cmd_last
                myConn.Open()
                Dim cmd2 = New SqlCommand(execmd, myConn)
                cmd2.ExecuteNonQuery()
                rowInsertNum += 1
                myConn.Close()
            End If
        Next

        'gloff
        For row As Integer = 0 To dgvExcel.RowCount - 1
            Dim queryGLOff As New ArrayList
            Dim doc_no = ""
            For Each seq_temp As String In gloffseq
                If seq_temp.Split(".")(0).Trim.Equals(value_arraylist(0)(row)(0)) Then
                    doc_no = seq_temp.Split(".")(1).Trim
                End If
            Next

            queryGLOff.Add(value_arraylist(0)(row)(0)) 'accno
            queryGLOff.Add("GL") 'doc_type
            queryGLOff.Add(doc_no) 'doc_no
            queryGLOff.Add(2) 'seq
            queryGLOff.Add(Function_Form.convertDateFormat(value_arraylist(0)(row)(4))) 'doc_date
            queryGLOff.Add(value_arraylist(0)(row)(0)) 'pkdoc_type
            queryGLOff.Add(value_arraylist(2)(row)(5)) 'pkdoc_no
            queryGLOff.Add(2) 'pkseq
            queryGLOff.Add(Function_Form.convertDateFormat(value_arraylist(2)(row)(8))) 'pkdoc_date
            queryGLOff.Add(-1) 'sign
            queryGLOff.Add(value_arraylist(2)(row)(10)) 'pkfx_rate
            queryGLOff.Add(value_arraylist(2)(row)(11)) 'paid
            Dim fx_rate As Double = CDbl(value_arraylist(0)(row)(19))
            Dim local_paid As Double = Math.Round(CDbl(value_arraylist(2)(row)(11)) * fx_rate, 2)
            queryGLOff.Add(fx_rate) 'fx_rate
            queryGLOff.Add(local_paid) 'local_paid
            queryGLOff.Add(local_paid) 'local_amount
            queryGLOff.Add(Function_Form.getNull(3)) 'fx_gainloss
            queryGLOff.Add(Function_Form.getNull(3)) 'v_gainloss

            If CDbl(value_arraylist(2)(row)(11)) = 0 Then
                Continue For
            End If

            Dim gloff_cmd As String = queryTable(2)(2)
            For j = 0 To queryGLOff.Count - 1
                gloff_cmd += "'" + queryGLOff(j).ToString + "',"
            Next
            gloff_cmd = gloff_cmd.Substring(0, gloff_cmd.Length - 1) + ")"

            myConn.Open()
            Dim cmd_gloff = New SqlCommand(gloff_cmd, myConn)
            cmd_gloff.ExecuteNonQuery()
            rowInsertNum += 1
            myConn.Close()
        Next

        'ap
        For row As Integer = 0 To dgvExcel.RowCount - 1
            If Not value_arraylist(0)(row)(0).Equals("{INVALID ARRAY}") Then
                Dim startRange As Integer = -1
                Dim endRange As Integer = -1
                For Each range As String In rangeQuo
                    If range.Split(".")(0).ToString.Trim.Equals(row.ToString.Trim) Then
                        startRange = CInt(range.Split(".")(0).ToString.Trim)
                        endRange = CInt(range.Split(".")(1).ToString.Trim)
                    End If
                Next

                Dim queryAL As New ArrayList
                Dim doc_no = ""
                For Each seq_temp As String In gloffseq
                    If seq_temp.Split(".")(0).Trim.Equals(value_arraylist(0)(row)(0)) Then
                        doc_no = seq_temp.Split(".")(1).Trim
                    End If
                Next

                queryAL.Clear()
                queryAL.Add(value_arraylist(0)(row)(0)) 'suppcode
                queryAL.Add("GL") 'doc_type
                queryAL.Add(doc_no) 'doc_no
                queryAL.Add(2) 'seq
                queryAL.Add(Function_Form.convertDateFormat(value_arraylist(0)(row)(4))) 'doc_date
                queryAL.Add(Function_Form.convertDateFormat(value_arraylist(0)(row)(4))) 'due_date
                queryAL.Add(value_arraylist(0)(row)(6)) 'refno
                queryAL.Add(value_arraylist(0)(row)(7)) 'refno2
                queryAL.Add(Function_Form.getNull(0)) 'refno3
                queryAL.Add(value_arraylist(1)(row)(9)) 'desp
                queryAL.Add(value_arraylist(0)(row)(10)) 'desp2
                queryAL.Add(Function_Form.getNull(0)) 'desp3
                queryAL.Add(Function_Form.getNull(0)) 'desp4
                queryAL.Add(Function_Form.getNull(0)) 'remark1
                queryAL.Add(Function_Form.getNull(0)) 'remark2
                If value_arraylist(0)(row)(15).ToString.Trim.Equals(String.Empty) Then
                    queryAL.Add(Function_Form.getNull(1)) 'chqrc_date
                Else
                    queryAL.Add(Function_Form.convertDateFormat(value_arraylist(0)(row)(4))) 'chqrc_date
                End If
                queryAL.Add(Function_Form.getNull(1)) 'koff_date
                queryAL.Add(value_arraylist(0)(row)(18)) 'curr_code
                queryAL.Add(value_arraylist(0)(row)(19)) 'fx_rate
                queryAL.Add(Function_Form.getNull(3)) 'fx_gainloss
                queryAL.Add(Function_Form.getNull(1)) 'koff_date
                Dim amount = Math.Round(CDbl(value_arraylist(0)(row)(21)), 2)
                queryAL.Add(amount) 'amount
                Dim paid As Double = 0
                For i = startRange To endRange
                    paid += CDbl(value_arraylist(2)(i)(11))
                Next
                paid = Math.Round(CDbl(paid), 2)
                queryAL.Add(paid) 'paid
                queryAL.Add(Math.Round(amount * CDbl(value_arraylist(0)(row)(19)), 2)) 'local_amount
                queryAL.Add(Math.Round(paid * CDbl(value_arraylist(0)(row)(19)), 2)) 'local_paid
                queryAL.Add(Function_Form.getNull(3)) 'taxable
                queryAL.Add(Function_Form.getNull(3)) 'fx_taxable
                queryAL.Add(Function_Form.getNull(3)) 'fx_tax
                If paid >= amount Then
                    queryAL.Add(1) 'knockoff
                Else
                    queryAL.Add(0) 'knockoff
                End If
                queryAL.Add(Function_Form.getNull(0)) 'accmgr_id
                queryAL.Add(value_arraylist(0)(row)(31)) 'projcode
                queryAL.Add(Function_Form.getNull(0)) 'deptcode
                queryAL.Add("P") 'bill_type
                queryAL.Add(Function_Form.getNull(0)) 'spcode
                queryAL.Add(Function_Form.getNull(0)) 'source
                queryAL.Add(Function_Form.getNull(0)) 'taxcode
                queryAL.Add(Function_Form.getNull(1)) 'taxdate
                queryAL.Add(Function_Form.getNull(1)) 'taxdate_bt
                queryAL.Add(Function_Form.getNull(0)) 'tax_basis
                queryAL.Add(Function_Form.getNull(0)) 'lkdoc_type
                queryAL.Add(Function_Form.getNull(0)) 'lkdoc_no
                queryAL.Add(Function_Form.getNull(3)) 'lkseq

                Dim execmd As String = queryTable(0)(2)
                For j = 0 To queryAL.Count - 1
                    Try
                        execmd += "'" + queryAL(j) + "',"
                    Catch ex As InvalidCastException
                        execmd += "'" + queryAL(j).ToString + "',"
                    End Try
                Next
                execmd = execmd.Substring(0, execmd.Length - 1) + ")"
                myConn.Open()
                Dim cmd1 = New SqlCommand(execmd, myConn)
                cmd1.ExecuteNonQuery()
                rowInsertNum += 1
                myConn.Close()
            End If
        Next


        Function_Form.promptImportSuccess(rowInsertNum, rowUpdateNum)
        Function_Form.printExcelResult("Pay_Bills", queryTable, value_arraylist, sql_format_arraylist, dgvExcel)
    End Sub
    Private Function existed_checker(table As String, sql_value As String, value As String)
        myConn.Open()
        Dim exist_value As Boolean = False
        Dim command = New SqlCommand("SELECT * FROM " + table + " WHERE " + sql_value + " ='" + value + "'", myConn)
        Dim reader As SqlDataReader = command.ExecuteReader
        While reader.Read()
            exist_value = True
        End While
        myConn.Close()
        Return exist_value
    End Function
End Class