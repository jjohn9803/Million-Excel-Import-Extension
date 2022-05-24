Imports System.Data.SqlClient
Imports System.IO
Imports ClosedXML.Excel
Imports ExcelDataReader
Public Class Stock_Adjustment_Form
    Dim tables As DataTableCollection
    Private serverName As String
    Private database As String
    Private myConn As SqlConnection
    Private statusConnection As Boolean
    Private pwd_query As String
    Private import_type As String
    Private Sub Stock_Adjustment_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        Dim importType = "Stock Adjustment"
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
            For i = 0 To 4
                queryTable.Add(New ArrayList)
            Next
            queryTable(0).add("SA ts") '0
            queryTable(0).add("stkts") '1
            queryTable(1).add("SA ts Desc")
            queryTable(1).add("stktsdet")
            queryTable(2).add("SA Stock")
            queryTable(2).add("stock")
            queryTable(3).add("SA Product Serial No")
            queryTable(3).add("prodsn")
            queryTable(4).add("SA Stock Serial No")
            queryTable(4).add("stocksn")
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
            'stktsdet
            'amt
            If value_arraylist(1)(row)(11).Equals("{FORMULA_VALUE}") Then
                Dim qty = value_arraylist(1)(row)(9)
                Dim price = value_arraylist(1)(row)(10)
                Dim amt = qty * price
                value_arraylist(1)(row)(11) = Math.Round(amt, 2)
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
            'Stkts
            If Not value_arraylist(0)(row)(0).Equals("{INVALID ARRAY}") Then
                'stkts.doc_no / duplicate
                table = "stkts"
                value_name = "doc_no"
                value = value_arraylist(0)(row)(1)
                If Not value.Trim.Equals(String.Empty) Then
                    If existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' already existed in the database (" + table + ")!" + vbNewLine
                    End If
                    If Function_Form.repeatedExcelCell(dgvExcel, excel_format_arraylist(0)(1), value, row) Then
                        execute_valid = False
                        If Not exist_result.Contains(excel_format_arraylist(0)(1) + " '" + value + "' is repeated!") Then
                            exist_result += excel_format_arraylist(0)(1) + " '" + value + "' is repeated!" + vbNewLine
                        End If
                    End If
                End If

                'project.projcode / exist
                table = "project"
                value_name = "projcode"
                value = value_arraylist(0)(row)(5)
                If Not value.Trim.Equals(String.Empty) Then
                    If Not existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                End If

                'deptment.deptcode / exist
                table = "deptment"
                value_name = "deptcode"
                value = value_arraylist(0)(row)(6)
                If Not value.Trim.Equals(String.Empty) Then
                    If Not existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                End If
            End If

            'Stkts Desc
            If Not value_arraylist(1)(row)(0).Equals(String.Empty) Then
                'product.prodcode / exist
                table = "product"
                value_name = "prodcode"
                value = value_arraylist(1)(row)(4)
                If Not value.Trim.Equals(String.Empty) Then
                    If Not existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                End If

                'stockloc.location / exist
                table = "stockloc"
                value_name = "location"
                value = value_arraylist(1)(row)(15)
                If Not value.Trim.Equals(String.Empty) Then
                    If Not existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                End If

                'prodbatch.batchcode / exist
                table = "prodbatch"
                value_name = "batchcode"
                value = value_arraylist(1)(row)(16)
                If Not value.Trim.Equals(String.Empty) Then
                    If Not existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                End If
            End If

            'Serial No
            If Not value_arraylist(3)(row)(0).Equals(String.Empty) Then
                'prodsn.serialno / exist
                table = "prodsn"
                value_name = "serialno"
                Dim values As New List(Of String)(value_arraylist(3)(row)(1).ToString.Trim.Split(","c))
                For Each value_sn As String In values
                    If Not value_sn.Trim.Equals(String.Empty) Then
                        If Not existed_checker(table, value_name, value_sn) Then
                            execute_valid = False
                            exist_result += value_name + " '" + value_sn + "' is not found in the database (" + table + ")!" + vbNewLine
                        End If
                    End If
                Next

                'stocksn.serialno / exist
                table = "stocksn"
                value_name = "serialno"
                Dim values2 As New List(Of String)(value_arraylist(3)(row)(1).ToString.Trim.Split(","c))
                For Each value_sn As String In values2
                    If Not value_sn.Trim.Equals(String.Empty) Then
                        If Not existed_checker(table, value_name, value_sn) Then
                            execute_valid = False
                            exist_result += value_name + " '" + value_sn + "' is not found in the database (" + table + ")!" + vbNewLine
                        End If
                    End If
                Next
            End If
        Next
        If execute_valid = False Then
            MsgBox(exist_result + vbNewLine + "The operation has been stopped!", MsgBoxStyle.Exclamation)
            Return
        End If

        'endhardcore exist checker

        init()
        Dim exist_serial As Boolean = False
        Dim msg_serial As New ArrayList
        For row As Integer = 0 To dgvExcel.RowCount - 1
            If Not value_arraylist(3)(row)(1).ToString.Trim.Equals(String.Empty) Then
                Dim serialnos As New List(Of String)(value_arraylist(3)(row)(1).ToString.Trim.Split(","c))
                If serialnos.Count <> CInt(value_arraylist(1)(row)(9)) Then
                    MsgBox("The serial no quantity (" + serialnos.Count.ToString + ") does not matched with items quantity(" + value_arraylist(1)(row)(9).ToString + ")" + vbNewLine + "The operation has been stopped!", MsgBoxStyle.Exclamation)
                    Return
                End If
                For sn = 0 To serialnos.Count - 1
                    Dim serialno As String = serialnos(sn)
                    myConn.Open()
                    Dim sncommand = New SqlCommand("SELECT * FROM stocksn WHERE serialno ='" + serialno + "' AND qty='-1'", myConn)
                    Dim snreader As SqlDataReader = sncommand.ExecuteReader
                    While snreader.Read()
                        If Not msg_serial.Contains(snreader.GetValue(snreader.GetOrdinal("serialno")).ToString.Trim) Then
                            msg_serial.Add(snreader.GetValue(snreader.GetOrdinal("serialno")).ToString.Trim)
                        End If
                        exist_serial = True
                    End While
                    myConn.Close()
                    myConn.Open()
                    Dim prodsncommand = New SqlCommand("SELECT * FROM prodsn WHERE qty='-1' AND serialno ='" + serialno + "'", myConn)
                    Dim prodsnreader As SqlDataReader = prodsncommand.ExecuteReader
                    While prodsnreader.Read()
                        If Not msg_serial.Contains(prodsnreader.GetValue(prodsnreader.GetOrdinal("serialno")).ToString.Trim) Then
                            msg_serial.Add(prodsnreader.GetValue(prodsnreader.GetOrdinal("serialno")).ToString.Trim)
                        End If
                        exist_serial = True
                    End While
                    myConn.Close()
                Next
            End If
        Next
        If exist_serial Then
            Dim prompt_exist_serial As String = "The following serial no has been used:"
            For Each serial As String In msg_serial
                prompt_exist_serial += vbNewLine + serial
            Next
            MsgBox(prompt_exist_serial + vbNewLine + "The operation has been stopped!", MsgBoxStyle.Exclamation)
            Return
        End If

        Dim confirmImport As DialogResult = MsgBox("Are you sure to import data?", MsgBoxStyle.YesNo, "")
        If confirmImport = DialogResult.No Then
            Return
        End If

        Dim rowInsertNum = 0
        Dim rowUpdateNum = 0

        'prodsn(3) + stocksn(4)
        For row As Integer = 0 To dgvExcel.RowCount - 1
            If Not value_arraylist(3)(row)(1).ToString.Trim.Equals(String.Empty) Then
                'Dim mySource As Integer = -1
                'For Each target In rangeQuo
                '    If target.ToString.Split(".")(1).Equals(row.ToString) Then
                '        Dim sourceRow = CInt(target.ToString.Split(".")(0))
                '        mySource = sourceRow
                '    End If
                'Next
                Dim serialnos As New List(Of String)(value_arraylist(3)(row)(1).ToString.Trim.Split(","c))
                For sn = 0 To serialnos.Count - 1
                    Dim serialno As String = serialnos(sn)
                    Dim qty = "1"
                    Dim location = value_arraylist(3)(row)(4)
                    Dim doc_no = value_arraylist(3)(row)(8)
                    Dim line_no = value_arraylist(3)(row)(9)
                    Dim doc_date = Function_Form.convertDateFormat(value_arraylist(3)(row)(10))
                    Dim procode = value_arraylist(3)(row)(0)
                    Dim serialNoProdCommand As String = "UPDATE prodsn SET "
                    Dim serialNoColumns = "qty='" + qty + "',"
                    serialNoColumns += "location='" + location + "',"
                    serialNoColumns += "doc_type='SA',"
                    serialNoColumns += "doc_no='" + doc_no + "',"
                    serialNoColumns += "line_no='" + line_no + "',"
                    serialNoColumns += "doc_date='" + doc_date + "' "
                    serialNoColumns += "WHERE prodcode='" + procode + "' AND serialno='" + serialno + "'"
                    serialNoProdCommand += serialNoColumns
                    Dim command = New SqlCommand(serialNoProdCommand, myConn)
                    myConn.Open()
                    command.ExecuteNonQuery()
                    rowUpdateNum += 1
                    Dim serialNoStockdCommand As String = "INSERT INTO stocksn (prodcode,serialno,doc_type,doc_no,line_no,doc_date,qty,location) VALUES ('"
                    serialNoStockdCommand += procode + "','" + serialno + "','SA','" + doc_no + "','" + line_no + "','" + doc_date + "','" + qty + "','" + location + "')"
                    Dim command2 = New SqlCommand(serialNoStockdCommand, myConn)
                    command2.ExecuteNonQuery()
                    rowInsertNum += 1
                    myConn.Close()
                Next
            End If
        Next

        For i As Integer = 0 To 1
            init()
            Using command As New SqlCommand("", myConn)
                For row As Integer = 0 To dgvExcel.RowCount - 1
                    Using r As DataGridViewRow = dgvExcel.Rows(row)
                        If Not value_arraylist(i)(row)(0).Equals("{INVALID ARRAY}") Then
                            Dim query = ""
                            For g As Integer = 0 To value_arraylist(i)(row).count - 1
                                Dim value_temp As String = value_arraylist(i)(row)(g).ToString
                                If sql_format_arraylist(i)(g).ToString.Trim.Equals("createdate") Or sql_format_arraylist(i)(g).ToString.Trim.Equals("lastupdate") Then
                                    query += "'" + Function_Form.convertDateFormat(Date.Now.ToString) + "',"
                                    value_arraylist(i)(row)(g) = Function_Form.convertDateFormat(Date.Now.ToString)
                                ElseIf data_type_arraylist(i)(g).ToString.Trim.Contains("date") Then
                                    query += "'" + Function_Form.convertDateFormat(value_temp) + "',"
                                    value_arraylist(i)(row)(g) = Function_Form.convertDateFormat(value_temp)
                                ElseIf Not (value_temp.Equals("{._!@#$%^&*()}")) Then
                                    query += "'" + value_temp + "',"
                                End If
                            Next
                            Dim command_text As String = queryTable(i)(2) + query
                            command_text = command_text.Substring(0, command_text.Length - 1) + ")"
                            command.CommandText = command_text
                            myConn.Open()
                            Try
                                rowInsertNum += 1
                                command.ExecuteNonQuery()
                            Catch ex As Exception
                                MsgBox(ex.Message + vbNewLine + command_text, MsgBoxStyle.Exclamation)
                            End Try
                            myConn.Close()
                        End If

                    End Using
                Next
            End Using
        Next

        'stock
        For row As Integer = 0 To dgvExcel.RowCount - 1
            If Not value_arraylist(1)(row)(4).ToString.Trim.Equals(String.Empty) Then
                Dim mySource As Integer = -1
                For Each target In rangeQuo
                    If target.ToString.Split(".")(1).Equals(row.ToString) Then
                        Dim sourceRow = CInt(target.ToString.Split(".")(0))
                        mySource = sourceRow
                    End If
                Next

                Dim insertArray As New ArrayList
                insertArray.Add(value_arraylist(1)(row)(4)) 'prodcode
                insertArray.Add(value_arraylist(1)(row)(1)) 'doc_type
                insertArray.Add(value_arraylist(1)(row)(2)) 'doc_no

                Dim dkeyFromDO As String = ""
                Dim command_temp = New SqlCommand("SELECT TOP 1 dkey FROM stktsdet WHERE doc_no ='" + value_arraylist(2)(row)(2) + "' AND line_no ='" + value_arraylist(2)(row)(4) + "'", myConn)
                myConn.Open()
                Dim reader_temp As SqlDataReader = command_temp.ExecuteReader
                While reader_temp.Read()
                    dkeyFromDO = reader_temp.GetValue(0).ToString
                End While
                myConn.Close()
                value_arraylist(2)(row)(3) = dkeyFromDO
                insertArray.Add(dkeyFromDO) 'dkey

                insertArray.Add(value_arraylist(1)(row)(3)) 'line_no
                insertArray.Add(value_arraylist(0)(mySource)(2)) 'doc_date
                insertArray.Add(value_arraylist(0)(mySource)(3)) 'doc_desp
                insertArray.Add(value_arraylist(0)(mySource)(4)) 'doc_desp2
                insertArray.Add(Function_Form.getNull(0)) 'custcode
                insertArray.Add(Function_Form.getNull(0)) 'suppcode
                insertArray.Add(Function_Form.getNull(0)) 'refno
                insertArray.Add(Function_Form.getNull(0)) 'refno2
                insertArray.Add(value_arraylist(1)(row)(9)) 'qty
                insertArray.Add(value_arraylist(1)(row)(10)) 'cost
                insertArray.Add(Function_Form.getNull(3)) 'price
                insertArray.Add(value_arraylist(1)(row)(11)) 'local_amount
                insertArray.Add(Function_Form.getNull(3)) 'utd_cost
                insertArray.Add(value_arraylist(1)(row)(15)) 'location
                insertArray.Add(value_arraylist(1)(row)(16)) 'batchcode
                insertArray.Add(value_arraylist(0)(mySource)(5)) 'projcode
                insertArray.Add(value_arraylist(0)(mySource)(6)) 'deptcode
                insertArray.Add(Function_Form.getNull(0)) 'pkdoc_type
                insertArray.Add(Function_Form.getNull(0)) 'pkdoc_no
                insertArray.Add(Function_Form.getNull(3)) 'pkdkey
                insertArray.Add(Function_Form.getNull(3)) 'bfseq
                Dim stockCommand As String = queryTable(2)(2)
                For Each query In insertArray
                    stockCommand += "'" + query.ToString + "',"
                Next
                stockCommand = stockCommand.Substring(0, stockCommand.Length - 1) + ")"
                Dim command = New SqlCommand(stockCommand, myConn)
                myConn.Open()
                'Clipboard.SetText(stockCommand)
                'MsgBox(stockCommand)
                command.ExecuteNonQuery()
                rowInsertNum += 1
                myConn.Close()
            End If
        Next

        Function_Form.promptImportSuccess(rowInsertNum, rowUpdateNum)
        Function_Form.printExcelResult("Stock_Adjustment", queryTable, value_arraylist, sql_format_arraylist, dgvExcel)
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