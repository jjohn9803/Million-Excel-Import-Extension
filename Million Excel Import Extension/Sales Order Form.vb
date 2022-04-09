Imports System.Data.SqlClient
Imports System.IO
Imports ClosedXML.Excel
Imports ExcelDataReader

Public Class Sales_Order_Form
    Dim tables As DataTableCollection
    Private serverName As String
    Private database As String
    Private myConn As SqlConnection
    Private statusConnection As Boolean
    Private pwd_query As String
    Private import_type As String
    Private Sub Sales_Order_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        init()
    End Sub
    Private Sub init()
        serverName = SQL_Connection_Form.serverName
        database = SQL_Connection_Form.database
        myConn = SQL_Connection_Form.myConn
        statusConnection = SQL_Connection_Form.statusConnection
        pwd_query = SQL_Connection_Form.pwd_query
        import_type = SQL_Connection_Form.import_type
        txtType.Text = import_type
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
    Private Function getMaintainSetting() As String
        Return SQL_Connection_Form.returnUpperFolder(Application.StartupPath(), 2) + "maintain.xls"
    End Function
    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        Dim importType = "Sales Order"
        Dim tableExcelSetting As DataTableCollection
        'Try
        Using stream = File.Open(getMaintainSetting, FileMode.Open, FileAccess.Read)
                Using reader As IExcelDataReader = ExcelReaderFactory.CreateReader(stream)
                    Dim result As DataSet = reader.AsDataSet(New ExcelDataSetConfiguration() With {
                                                                         .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
                                                                         .UseHeaderRow = True}})
                    tableExcelSetting = result.Tables
                    Dim queryTable As New ArrayList
                    queryTable.Add(New ArrayList)
                    queryTable.Add(New ArrayList)
                    queryTable(0).add("Sales Order") '0
                    queryTable(0).add("sorder") '1
                    queryTable(1).add("Sales Order Desc")
                    queryTable(1).add("sorderdet")
                    quotationWriteIntoSQL(tableExcelSetting, queryTable)
                End Using
            End Using
        'Catch ex As Exception
        '    MsgBox(ex.Message + vbNewLine + ex.StackTrace, MsgBoxStyle.Critical)
        'End Try
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
                                    value = New Date.ToString
                                    value_arraylist(i)(row).add(New Date.ToString)
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
                                                value_arraylist(i)(row)(g) = New Date.ToString
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
                                            myConn = New SqlConnection("Data Source=" + serverName + ";" & "Initial Catalog=" + database + ";" + pwd_query)

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
                'batchno
                If value_arraylist(0)(row)(48).length = 2 Then
                    Dim date1 = value_arraylist(0)(row)(2)
                    Dim batchno = value_arraylist(0)(row)(48)
                    value_arraylist(0)(row)(48) = Year(date1).ToString.Substring(2) + Month(date1).ToString("00") + batchno
                End If
            End If

            'qtyrp
            If value_arraylist(1)(row)(10).Equals("{FORMULA_VALUE}") Then
                Dim qty = value_arraylist(1)(row)(9)
                value_arraylist(1)(row)(10) = qty
            End If

            'gross_amt
            If value_arraylist(1)(row)(26).Equals("{FORMULA_VALUE}") Then
                Dim qty = CDbl(value_arraylist(1)(row)(9))
                Dim price = CDbl(value_arraylist(1)(row)(13))
                Dim gross_amt = qty * price
                value_arraylist(1)(row)(26) = Math.Round(gross_amt, 2)
            End If

            'disamt1
            If value_arraylist(1)(row)(14).Equals("{FORMULA_VALUE}") Then
                Dim gross_amt = CDbl(value_arraylist(1)(row)(26))
                Dim disp1 = CDbl(value_arraylist(1)(row)(18))
                Dim disamt1 = gross_amt * (disp1 * 0.01)
                value_arraylist(1)(row)(14) = Math.Round(disamt1, 2)
            End If

            'disamt2
            If value_arraylist(1)(row)(15).Equals("{FORMULA_VALUE}") Then
                Dim gross_amt = CDbl(value_arraylist(1)(row)(26))
                Dim disamt1 = CDbl(value_arraylist(1)(row)(14))
                Dim disp2 = CDbl(value_arraylist(1)(row)(19))
                Dim disamt2 = (gross_amt - disamt1) * (disp2 * 0.01)
                value_arraylist(1)(row)(15) = Math.Round(disamt2, 2)
            End If

            'disamt3
            If value_arraylist(1)(row)(16).Equals("{FORMULA_VALUE}") Then
                Dim gross_amt = CDbl(value_arraylist(1)(row)(26))
                Dim disamt1 = CDbl(value_arraylist(1)(row)(14))
                Dim disamt2 = CDbl(value_arraylist(1)(row)(15))
                Dim disp3 = CDbl(value_arraylist(1)(row)(20))
                Dim disamt3 = (gross_amt - disamt1 - disamt2) * (disp3 * 0.01)
                value_arraylist(1)(row)(16) = Math.Round(disamt3, 2)
            End If

            'disamt
            If value_arraylist(1)(row)(17).Equals("{FORMULA_VALUE}") Then
                Dim disamt1 = CDbl(value_arraylist(1)(row)(14))
                Dim disamt2 = CDbl(value_arraylist(1)(row)(15))
                Dim disamt3 = CDbl(value_arraylist(1)(row)(16))
                Dim disamt = disamt1 + disamt2 + disamt3
                value_arraylist(1)(row)(17) = Math.Round(disamt, 2)
            End If

            'nett_amt
            If value_arraylist(1)(row)(27).Equals("{FORMULA_VALUE}") Then
                Dim gross_amt = CDbl(value_arraylist(1)(row)(26))
                Dim disamt = CDbl(value_arraylist(1)(row)(17))
                Dim nett_amt = gross_amt - disamt
                value_arraylist(1)(row)(27) = Math.Round(nett_amt, 2)
            End If

            'amt
            If value_arraylist(1)(row)(28).Equals("{FORMULA_VALUE}") Then
                Dim gross_amt = CDbl(value_arraylist(1)(row)(26))
                Dim disamt = CDbl(value_arraylist(1)(row)(17))
                Dim amt = gross_amt - disamt
                value_arraylist(1)(row)(28) = Math.Round(amt, 2)
            End If

            'taxamt1
            If value_arraylist(1)(row)(21).Equals("{FORMULA_VALUE}") Then
                Dim nett_amt = CDbl(value_arraylist(1)(row)(27))
                Dim taxp1 = 0
                If Not value_arraylist(1)(row)(25).Equals(String.Empty) Then
                    taxp1 = CDbl(value_arraylist(1)(row)(25))
                End If
                Dim taxamt1 = nett_amt * (taxp1 * 0.01)
                value_arraylist(1)(row)(21) = Math.Round(taxamt1, 2)
            End If

            'taxamt
            If value_arraylist(1)(row)(24).Equals("{FORMULA_VALUE}") Then
                Dim taxamt1 = CDbl(value_arraylist(1)(row)(21))
                Dim taxamt2 = CDbl(value_arraylist(1)(row)(22))
                Dim taxamt = taxamt1 + taxamt2
                value_arraylist(1)(row)(24) = Math.Round(taxamt, 2)
            End If

            'local converter
            Dim find_local = row
            For f = find_local To 0 Step -1
                If Not value_arraylist(0)(f)(0).Equals("{INVALID ARRAY}") Then
                    find_local = f
                    Exit For
                End If
            Next
            Dim fx_rate = CDbl(value_arraylist(0)(find_local)(18))

            'local_price
            If value_arraylist(1)(row)(29).Equals("{FORMULA_VALUE}") Then
                Dim price = CDbl(value_arraylist(1)(row)(13))
                Dim local_price = price * fx_rate
                value_arraylist(1)(row)(29) = local_price
            End If

            'local_gamt
            If value_arraylist(1)(row)(30).Equals("{FORMULA_VALUE}") Then
                Dim gross_amt = CDbl(value_arraylist(1)(row)(26))
                Dim local_gamt = gross_amt * fx_rate
                value_arraylist(1)(row)(30) = Math.Round(local_gamt, 2)
            End If

            'local_disamt
            If value_arraylist(1)(row)(31).Equals("{FORMULA_VALUE}") Then
                Dim disamt = CDbl(value_arraylist(1)(row)(17))
                Dim local_disamt = disamt * fx_rate
                value_arraylist(1)(row)(31) = Math.Round(local_disamt, 2)
            End If

            'local_namt
            If value_arraylist(1)(row)(32).Equals("{FORMULA_VALUE}") Then
                Dim local_gamt = CDbl(value_arraylist(1)(row)(30))
                Dim local_disamt = CDbl(value_arraylist(1)(row)(31))
                Dim local_namt = local_gamt - local_disamt
                value_arraylist(1)(row)(32) = Math.Round(local_namt, 2)
            End If

            'local_taxamt1
            If value_arraylist(1)(row)(33).Equals("{FORMULA_VALUE}") Then
                Dim taxamt1 = CDbl(value_arraylist(1)(row)(21))
                Dim local_taxamt1 = taxamt1 * fx_rate
                value_arraylist(1)(row)(33) = Math.Round(local_taxamt1, 2)
            End If

            'local_taxamt2
            If value_arraylist(1)(row)(34).Equals("{FORMULA_VALUE}") Then
                Dim taxamt2 = CDbl(value_arraylist(1)(row)(22))
                Dim local_taxamt2 = taxamt2 * fx_rate
                value_arraylist(1)(row)(34) = Math.Round(local_taxamt2, 2)
            End If

            'local_taxamtadj1
            If value_arraylist(1)(row)(35).Equals("{FORMULA_VALUE}") Then
                Dim taxamtadj1 = CDbl(value_arraylist(1)(row)(23))
                Dim local_taxamtadj1 = taxamtadj1 * fx_rate
                value_arraylist(1)(row)(35) = Math.Round(local_taxamtadj1, 2)
            End If

            'local_taxamt
            If value_arraylist(1)(row)(36).Equals("{FORMULA_VALUE}") Then
                Dim taxamt = CDbl(value_arraylist(1)(row)(24))
                Dim local_taxamt = taxamt * fx_rate
                value_arraylist(1)(row)(36) = Math.Round(local_taxamt, 2)
            End If

            'local_amt
            If value_arraylist(1)(row)(37).Equals("{FORMULA_VALUE}") Then
                Dim amt = CDbl(value_arraylist(1)(row)(28))
                Dim local_amt = amt * fx_rate
                value_arraylist(1)(row)(37) = Math.Round(local_amt, 2)
            End If

            'local_amtrp
            If value_arraylist(1)(row)(38).Equals("{FORMULA_VALUE}") Then
                Dim local_amt = value_arraylist(1)(row)(37)
                value_arraylist(1)(row)(38) = Math.Round(local_amt, 2)
            End If

            'local_mcamt1
            If value_arraylist(1)(row)(40).Equals("{FORMULA_VALUE}") Then
                Dim mcamt1 = CDbl(value_arraylist(1)(row)(39))
                Dim local_mcamt1 = mcamt1 * fx_rate
                value_arraylist(1)(row)(40) = Math.Round(local_mcamt1, 2)
            End If

        Next

        'End Hardcode Formula
        'For i As Integer = 0 To 1
        '    Dim str = ""
        '    For g As Integer = 0 To sql_format_arraylist(i).count - 1
        '        str += sql_format_arraylist(i)(g) + vbTab
        '    Next
        '    MsgBox(str)
        'Next

        For i As Integer = 0 To 0
            For row As Integer = 0 To dgvExcel.RowCount - 1
                For g As Integer = 0 To value_arraylist(i)(row).count - 1
                    Dim value_temp As String = value_arraylist(i)(row)(g).ToString.Trim
                    If value_temp.Equals("{FORMULA_VALUE}") Then
                        Dim formula_temp = formula_arraylist(i)(g).ToString.Trim
                        Dim finalized_temp = New List(Of String)(formula_temp.Split("?"c))
                        Dim cal_result = ""
                        Dim cal_type = ""
                        For Each local_formula As String In finalized_temp
                            If local_formula.Contains("+") Or local_formula.Contains("-") Or local_formula.Contains("*") Then
                                Dim expression = New List(Of String)(local_formula.Split(New [Char]() {"+"c, "*"c, "-"c}))
                                Dim calculation = local_formula
                                'MsgBox(local_formula)
                                For Each express As String In expression
                                    Dim table_name_index = -1
                                    Dim table_value_index = -1
                                    Dim express_temp = (express.Replace("(", "")).Replace(")", "").Trim
                                    Dim cal_value_temp = ""
                                    If express_temp.Contains("~") Then
                                        Dim table_name = express_temp.Split("~")(0).Trim
                                        Dim table_value_name = express_temp.Split("~")(1).Trim
                                        For yu = 0 To queryTable.Count - 1
                                            Dim search_table_name = queryTable(yu)(1)
                                            If table_name.Equals(search_table_name) Then
                                                table_name_index = yu
                                            End If
                                        Next
                                        'table_name_index=找到quo还是quo_desc OUTPUT:0或者1 同辈：query.i value(这里)()()
                                        'getDescSource=找到这个desc是属于哪一个quo的(row) 例如：quo1拥有product1,2 quo2拥有product3,4,5 OUTPUT:0或者2 同辈：row value()(这里)()
                                        'table_value_index=从sql的0或1找到符合formula express value名字在valuearraylist的位置 value()()(这里)
                                        'value_arraylist=(type)(row)(value)
                                        Try
                                            table_value_index = sql_format_arraylist(table_name_index).IndexOf(table_value_name)
                                        Catch ex As Exception
                                            MsgBox("ROW: " + row.ToString + vbNewLine + "table_name_index: " + table_name_index.ToString + vbNewLine + "table_value_name: " + table_value_name)
                                        End Try

                                        Dim myTargets As New List(Of String)
                                        For Each target As String In rangeQuo
                                            Dim rangeStart = target.Split(".")(0)
                                            If rangeStart.Equals(row.ToString) Then
                                                myTargets.Add(target.Split(".")(1))
                                                'MsgBox(row.ToString + " is adding " + target.Split(".")(1) + " as its target")
                                            End If
                                        Next
                                        For Each target As String In myTargets
                                            Dim getValueOfTarget = value_arraylist(table_name_index)(CInt(target))(table_value_index).ToString.Trim
                                            cal_value_temp += getValueOfTarget + "+"
                                        Next
                                        cal_value_temp = "(" + cal_value_temp.Substring(0, cal_value_temp.Length - 1) + ")"

                                        If cal_value_temp.Contains("{FORMULA_VALUE}") Then
                                            MsgBox("No way formula_value gonna exists here!")
                                            cal_value_temp = 0
                                        End If

                                        cal_type = "1"
                                    Else
                                        'valuearraylist(querytable,i)(table_name,0,1 % row)(table_index,values_index)
                                        'Dim getDescSource = CInt(rangeQuo(row).ToString.Split(".")(0))
                                        'table_name_index = 1
                                        table_name_index = i
                                        Dim table_value_name = express_temp.Trim
                                        table_value_index = sql_format_arraylist(table_name_index).IndexOf(table_value_name)
                                        If table_value_index <> -1 Then
                                            cal_value_temp = value_arraylist(table_name_index)(row)(table_value_index).ToString.Trim
                                            If cal_value_temp.Equals("{FORMULA_VALUE}") Then
                                                MsgBox("WHY ARE YOU A FORMULA")
                                                cal_value_temp = 0
                                            End If
                                        Else
                                            'MsgBox("404 not found" + vbNewLine + table_name_index.ToString + vbNewLine + row.ToString + vbNewLine + table_value_index.ToString + vbNewLine + table_value_name)
                                        End If

                                        cal_type = "2"
                                    End If
                                    'MsgBox("From " + express_temp.ToString + " TO " + cal_value_temp.ToString)
                                    If Not express_temp.Contains(".") Then
                                        calculation = calculation.Replace(express_temp, cal_value_temp)
                                    End If
                                Next
                                Try
                                    cal_result = New DataTable().Compute(calculation, Nothing)
                                    'MsgBox("After calculation and the result: " + cal_result)
                                Catch ex As Exception
                                    MsgBox("This calculation has problem -> " + calculation + vbNewLine + local_formula, MsgBoxStyle.Exclamation)
                                    MsgBox(value_arraylist(0)(row)(33))
                                    'Dim str = ""
                                    'For Each temp As String In value_arraylist(0)(row)
                                    '    str += temp + vbTab
                                    'Next
                                    'MsgBox(str)
                                End Try

                                'MsgBox("Calculation:" + vbTab + calculation + vbNewLine + local_formula)
                            Else
                                '没有加减乘除的单个value
                                Dim table_name_index = -1
                                Dim table_value_index = -1
                                Dim cal_value_temp = ""
                                Dim calculation = local_formula
                                If calculation.Contains("~") Then
                                    Dim table_name = calculation.Split("~")(0).Trim
                                    Dim table_value_name = calculation.Split("~")(1).Trim
                                    For yu = 0 To queryTable.Count - 1
                                        Dim search_table_name = queryTable(yu)(1)
                                        If table_name.Equals(search_table_name) Then
                                            table_name_index = yu
                                        End If
                                    Next
                                    'table_name_index=找到quo还是quo_desc OUTPUT:0或者1 同辈：query.i value(这里)()()
                                    'getDescSource=找到这个desc是属于哪一个quo的(row) 例如：quo1拥有product1,2 quo2拥有product3,4,5 OUTPUT:0或者2 同辈：row value()(这里)()
                                    'table_value_index=从sql的0或1找到符合formula express value名字在valuearraylist的位置 value()()(这里)
                                    'value_arraylist=(type)(row)(value)
                                    table_value_index = sql_format_arraylist(table_name_index).IndexOf(table_value_name)

                                    Dim myTargets As New List(Of String)
                                    For Each target As String In rangeQuo
                                        Dim rangeStart = target.Split(".")(0)
                                        If rangeStart.Equals(row.ToString) Then
                                            myTargets.Add(target.Split(".")(1))
                                            'MsgBox(row.ToString + " is adding " + target.Split(".")(1) + " as its target")
                                        End If
                                    Next
                                    For Each target As String In myTargets
                                        Dim getValueOfTarget = value_arraylist(table_name_index)(CInt(target))(table_value_index).ToString.Trim
                                        cal_value_temp += getValueOfTarget + "+"
                                    Next
                                    cal_value_temp = "(" + cal_value_temp.Substring(0, cal_value_temp.Length - 1) + ")"
                                    'MsgBox("Expresses: " + expresses.Substring(0, expresses.Length - 1) + vbNewLine + local_formula)
                                    'Dim getDescSource = CInt(rangeQuo(row).ToString.Split(".")(0))

                                    'cal_value_temp = value_arraylist(table_name_index)(getDescSource)(table_value_index).ToString.Trim
                                    If cal_value_temp.Contains("{FORMULA_VALUE}") Then
                                        MsgBox("No way formula_value gonna exists here!")
                                        cal_value_temp = 0
                                    End If
                                    cal_type = "B1"
                                Else
                                    'table_name_index=找到quo还是quo_desc OUTPUT:0或者1 同辈：query.i value(这里)()()
                                    'getDescSource=找到这个desc是属于哪一个quo的(row) 例如：quo1拥有product1,2 quo2拥有product3,4,5 OUTPUT:0或者2 同辈：row value()(这里)()
                                    'table_value_index=从sql的0或1找到符合formula express value名字在valuearraylist的位置 value()()(这里)
                                    'value_arraylist=(type)(row)(value)
                                    table_name_index = i
                                    Dim table_value_name = local_formula.Trim
                                    table_value_index = sql_format_arraylist(table_name_index).IndexOf(table_value_name)
                                    If table_value_index <> -1 Then
                                        cal_value_temp = value_arraylist(table_name_index)(row)(table_value_index).ToString.Trim
                                        If cal_value_temp.Equals("{FORMULA_VALUE}") Then
                                            MsgBox("WHY ARE YOU A FORMULA")
                                            cal_value_temp = 0
                                        End If
                                    Else
                                        MsgBox("单品 404 not found" + vbNewLine + table_name_index.ToString + vbNewLine + row.ToString + vbNewLine + table_value_index.ToString + vbNewLine + table_value_name)
                                    End If
                                    cal_type = "B2"
                                End If
                                If Not calculation.Contains(".") Then
                                    calculation = cal_value_temp
                                End If
                                Try
                                    cal_result = New DataTable().Compute(calculation, Nothing)
                                    'MsgBox("After calculation and the result: " + cal_result)
                                Catch ex As Exception
                                    MsgBox("This calculation has problem -> " + calculation + vbNewLine + local_formula, MsgBoxStyle.Exclamation)
                                    MsgBox(value_arraylist(0)(row)(33))
                                    'Dim str = ""
                                    'For Each temp As String In value_arraylist(0)(row)
                                    '    str += temp + vbTab
                                    'Next
                                    'MsgBox(str)
                                End Try
                            End If

                        Next
                        If cal_type = "" Then
                            MsgBox("Empty -> " + formula_temp + vbNewLine + finalized_temp(0).ToString)
                        End If
                        'MsgBox("Upgrade value: " + value_arraylist(i)(row)(g) + vbNewLine + "This calculation -> " + cal_result + vbNewLine + formula_temp + vbNewLine + "Cal Type: " + cal_type, MsgBoxStyle.Exclamation)
                        'value_arraylist(i)(row)(g) = Format(CDec(cal_result), "0.00").ToString
                        If CDec(cal_result) <> 0 Then
                            value_arraylist(i)(row)(g) = Format(CDec(cal_result), "0.00").ToString
                        Else
                            value_arraylist(i)(row)(g) = cal_result.ToString
                        End If
                    End If
                Next
            Next
        Next

        'Hardcore Taxable
        For row As Integer = 0 To dgvExcel.RowCount - 1
            If Not value_arraylist(0)(row)(0).Equals("{INVALID ARRAY}") Then
                'taxable
                If value_arraylist(0)(row)(26).ToString.Trim.Equals("0") Then
                    'get query target
                    Dim taxable As Double = 0
                    Dim myTarget As New ArrayList
                    For Each target In rangeQuo
                        If target.ToString.Split(".")(0).Equals(row.ToString) Then
                            Dim targetRow = CInt(target.ToString.Split(".")(1))
                            'myTarget.Add(target.ToString.Split(".")(1))
                            Dim nett_amt = value_arraylist(1)(targetRow)(25)
                            Dim taxcode = value_arraylist(1)(targetRow)(52)
                            If taxcode.ToString.Trim.ToCharArray.Count > 0 Then
                                taxable += nett_amt
                            End If
                        End If
                    Next
                    value_arraylist(0)(row)(26) = Math.Round(taxable, 2)
                End If
            End If
        Next
        'End Hardcore Taxable

        'hardcore exist checker
        Dim execute_valid As Boolean = True
        Dim exist_result As String = ""
        init()
        myConn = New SqlConnection("Data Source=" + serverName + ";" & "Initial Catalog=" + database + ";" + pwd_query)
        For row As Integer = 0 To dgvExcel.RowCount - 1
            Dim table As String
            Dim value_name As String
            Dim value As String
            'sorder
            If Not value_arraylist(0)(row)(0).Equals("{INVALID ARRAY}") Then
                'squote.doc_no / duplicate
                table = "sorder"
                value_name = "doc_no"
                value = value_arraylist(0)(row)(1)
                If existed_checker(table, value_name, value) Then
                    execute_valid = False
                    exist_result += value_name + " '" + value + "' already existed in the database (" + table + ")!" + vbNewLine
                End If

                'customer.custcode / exist
                table = "customer"
                value_name = "custcode"
                value = value_arraylist(0)(row)(9)
                If Not value.Trim.Equals(String.Empty) Then
                    If Not existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                End If

                'currency.curr_code / exist
                table = "currency"
                value_name = "curr_code"
                value = value_arraylist(0)(row)(16)
                If Not value.Trim.Equals(String.Empty) Then
                    If Not existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                End If

                'accmgr.accmgr_id / exist
                table = "accmgr"
                value_name = "accmgr_id"
                value = value_arraylist(0)(row)(16)
                If Not value.Trim.Equals(String.Empty) Then
                    If Not existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                End If

                'glbatch.batchno / exist
                table = "glbatch"
                value_name = "batchno"
                value = value_arraylist(0)(row)(48)
                If Not value.Trim.Equals(String.Empty) Then
                    If Not existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                End If

                'project.projcode / exist
                table = "project"
                value_name = "projcode"
                value = value_arraylist(0)(row)(49)
                If Not value.Trim.Equals(String.Empty) Then
                    If Not existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                End If

                'deptment.deptcode / exist
                table = "deptment"
                value_name = "deptcode"
                value = value_arraylist(0)(row)(50)
                If Not value.Trim.Equals(String.Empty) Then
                    If Not existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                End If
            End If

            'sorderdet
            If Not value_arraylist(1)(row)(0).Equals(String.Empty) Then
                'squotedet.doc_no / duplicate
                table = "sorderdet"
                value_name = "doc_no"
                value = value_arraylist(1)(row)(2)
                If Not value.Trim.Equals(String.Empty) Then
                    If existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' already existed in the database (" + table + ")!" + vbNewLine
                    End If
                End If

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

                'gltax.taxcode / exist
                table = "gltax"
                value_name = "taxcode"
                value = value_arraylist(1)(row)(54)
                If Not value.Trim.Equals(String.Empty) Then
                    If Not existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                End If

                'gldata.accno / exist
                table = "gldata"
                value_name = "accno"
                value = value_arraylist(1)(row)(57)
                If Not value.Trim.Equals(String.Empty) Then
                    If Not existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                End If

                'project.projcode / exist
                table = "project"
                value_name = "projcode"
                value = value_arraylist(1)(row)(58)
                If Not value.Trim.Equals(String.Empty) Then
                    If Not existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                End If

                'deptment.deptcode / exist
                table = "deptment"
                value_name = "deptcode"
                value = value_arraylist(1)(row)(59)
                If Not value.Trim.Equals(String.Empty) Then
                    If Not existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                End If

            End If

        Next
        If execute_valid = False Then
            MsgBox(exist_result + vbNewLine + "The operation has been stopped!", MsgBoxStyle.Exclamation)
            Return
        End If
        'endhardcore exist checker

        'For i As Integer = 0 To queryTable.Count - 1
        '    For row As Integer = 0 To dgvExcel.RowCount - 1
        '        Dim strs = ""
        '        For Each str As String In value_arraylist(i)(row)
        '            strs += str + vbTab
        '        Next
        '        MsgBox("Row " + row.ToString + vbNewLine + strs)
        '    Next
        'Next
        'Quotation only end
        Dim rowInsertNum = 0
        For i As Integer = 0 To queryTable.Count - 1
            init()
            Using myConn = New SqlConnection("Data Source=" + serverName + ";" & "Initial Catalog=" + database + ";" + pwd_query)
                Using command As New SqlCommand("", myConn)
                    For row As Integer = 0 To dgvExcel.RowCount - 1
                        Using r As DataGridViewRow = dgvExcel.Rows(row)
                            If Not value_arraylist(i)(row)(0).Equals("{INVALID ARRAY}") Then
                                Dim query = ""
                                For g As Integer = 0 To value_arraylist(i)(row).count - 1
                                    Dim value_temp As String = value_arraylist(i)(row)(g).ToString
                                    If sql_format_arraylist(i)(g).ToString.Trim.Equals("createdate") Or sql_format_arraylist(i)(g).ToString.Trim.Equals("lastupdate") Then
                                        query += "'" + Date.Now.ToString + "',"
                                        value_arraylist(i)(row)(g) = Date.Now.ToString
                                    ElseIf Not (value_temp.Equals("{._!@#$%^&*()}")) Then
                                        query += "'" + value_temp + "',"
                                    End If
                                Next
                                Dim command_text As String = queryTable(i)(2) + query
                                command_text = command_text.Substring(0, command_text.Length - 1) + ")"
                                'MsgBox(command_text)
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
            End Using
        Next
        MsgBox("Data Import Sucessfully!" + vbNewLine + "Row Inserted: " + rowInsertNum.ToString)
        Function_Form.printExcelResult("C:\Users\RBADM07\Desktop\Generated Result SO.xlsx", queryTable, value_arraylist, sql_format_arraylist, dgvExcel)
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