Imports System.Data.SqlClient
Imports System.IO
Imports ClosedXML.Excel
Imports ExcelDataReader
Public Class Quotation_Form
    Dim tables As DataTableCollection
    Private serverName As String
    Private database As String
    Private myConn As SqlConnection
    Private statusConnection As Boolean
    Private pwd_query As String
    Private import_type As String
    Private Sub Quotation_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        'MsgBox(serverName + vbTab + database + vbTab + myConn.ToString + vbTab + statusConnection.ToString + vbTab + pwd_query)
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
                                    For Each row As DataRow In table.Rows
                                        'If Not (row.ItemArray(1).ToString.Equals(String.Empty)) Then
                                        '    If row.ItemArray(1).GetType.ToString = "System.DateTime" Then
                                        '        MsgBox(DateTime.FromOADate(CDbl(Val(row.ItemArray(1).ToString))))
                                        '    End If

                                        'End If
                                    Next
                                    'Dim strs As String = ""
                                    'For Each row As DataRow In table.Rows
                                    '    'strDetail = row.Item("Detail")
                                    '    If Not (row.Item(1).ToString.Equals(String.Empty)) Then
                                    '        MsgBox(row.Item(1).ToString)
                                    '    End If

                                    '    'strs += CStr(row.Item(1)) + vbTab
                                    'Next row
                                    'MsgBox(strs)
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
        Dim importType = "Quotation"
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
                    queryTable(0).add("Quotation") '0
                    queryTable(0).add("squote") '1
                    queryTable(1).add("Quotation Desc")
                    queryTable(1).add("squotedet")
                    quotationWriteIntoSQLQuotation2(tableExcelSetting, queryTable)
                End Using
            End Using
        'Catch ex As Exception
        '    MsgBox(ex.Message + vbNewLine + ex.StackTrace, MsgBoxStyle.Critical)
        'End Try
    End Sub
    'Private Sub quotationWriteIntoSQLQuotation(tableExcelSetting As DataTableCollection, queryTable As ArrayList)
    '    Dim value_arraylist = New ArrayList
    '    Dim head_arraylist = New ArrayList
    '    For row As Integer = 0 To dgvExcel.RowCount - 1
    '        Using r As DataGridViewRow = dgvExcel.Rows(row)
    '            value_arraylist.Add(New ArrayList)
    '            For cell As Integer = 0 To r.Cells.Count - 1
    '                Dim headerText As String = dgvExcel.Columns(cell).HeaderText.Trim
    '                Dim cellValue As String = r.Cells(cell).Value.ToString.Trim
    '                value_arraylist(row).Add(cellValue)
    '                head_arraylist(row).Add(headerText)
    '                'MsgBox(headerText + vbTab + cellValue)
    '            Next
    '        End Using
    '    Next

    'End Sub
    Private Sub quotationWriteIntoSQLQuotation2(tableExcelSetting As DataTableCollection, queryTable As ArrayList)
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
                                    value = (New DateTime).ToString
                                    value_arraylist(i)(row).add((New DateTime).ToString)
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
                                    ElseIf default_temp.Contains("FK") Then
                                        If row > 0 Then
                                            Dim fk_value_temp = ""
                                            For h = row To 0 Step -1
                                                fk_value_temp = value_arraylist(i)(h)(g).ToString.Trim
                                                If fk_value_temp.ToCharArray.Count > 0 And Not (fk_value_temp.Equals("{DEFAULT_VALUE}")) Then
                                                    h = 0
                                                End If
                                            Next
                                            value_arraylist(i)(row)(g) = fk_value_temp
                                        Else
                                            Dim data_type_temp = data_type_arraylist(i)(g).ToString.Trim
                                            If data_type_temp.ToString.Contains("char") Or data_type_temp.ToString.Contains("text") Then
                                                value_arraylist(i)(row)(g) = "   "
                                            ElseIf data_type_temp.ToString.Contains("date") Or data_type_temp.ToString.Contains("time") Then
                                                value_arraylist(i)(row)(g) = (New DateTime).ToString
                                            Else
                                                value_arraylist(i)(row)(g) = "0"
                                            End If
                                        End If
                                    ElseIf default_temp.Contains("PK_int") Then
                                        value_arraylist(i)(row)(g) = "{._!@#$%^&*()}"
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

        'Quotation only
        'Quetdesc fill formula first then
        'find product 1,2 belong to quet 1
        'find product 2,3,4 belong to quet 2
        'quet 1 do formula, total of product 1 + 2
        'quet 2 do formula, total of products 2+3+4
        'INSERT into quet 1 formula field
        'Define the range between Quoatation and Desc
        Dim rangeQuo As New ArrayList
        Dim rangeEnd = -1
        For i = 0 To dgvExcel.RowCount - 1
            Dim str = value_arraylist(0)(i)(0)
            If Not str.Equals("{INVALID ARRAY}") Then
                rangeEnd = i
            End If
            rangeQuo.Add(rangeEnd.ToString + "." + i.ToString)
        Next
        For i As Integer = 1 To 1
            For row As Integer = 0 To dgvExcel.RowCount - 1
                For g As Integer = 0 To value_arraylist(i)(row).count - 1
                    Dim value_temp As String = value_arraylist(i)(row)(g).ToString.Trim
                    If value_temp.Equals("{FORMULA_VALUE}") Then
                        Dim formula_temp = formula_arraylist(i)(g).ToString.Trim
                        Dim finalized_temp = New List(Of String)(formula_temp.Split("?"c))
                        Dim cal_result = ""
                        For Each local_formula As String In finalized_temp
                            Dim expression = New List(Of String)(local_formula.Split(New [Char]() {"+"c, "*"c, "-"c}))
                            Dim calculation = local_formula
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

                                    Dim getDescSource = CInt(rangeQuo(row).ToString.Split(".")(0))
                                    table_value_index = sql_format_arraylist(table_name_index).IndexOf(table_value_name)
                                    cal_value_temp = value_arraylist(table_name_index)(getDescSource)(table_value_index).ToString.Trim
                                    If cal_value_temp.Equals("{FORMULA_VALUE}") Then
                                        cal_value_temp = 0
                                    End If

                                    'Dim getDescSource = rangeQuo(table_name_index)
                                    'table_value_index = sql_format_arraylist(table_name_index).IndexOf(table_value_name)
                                    'cal_value_temp = value_arraylist(0)(table_name_index)(table_value_index).ToString.Trim
                                    'If cal_value_temp.Equals("{FORMULA_VALUE}") Then
                                    '    cal_value_temp = 0
                                    'End If
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

                                    ''valuearraylist(querytable,i)(table_name,0,1 % row)(table_index,values_index)
                                    'Dim getDescSource = CInt(rangeQuo(row).ToString.Split(".")(0))
                                    ''table_name_index = 1
                                    'table_name_index = getDescSource
                                    'Dim table_value_name = express_temp.Trim
                                    'table_value_index = sql_format_arraylist(i).IndexOf(table_value_name)
                                    'If table_value_index <> -1 Then
                                    '    Try
                                    '        cal_value_temp = value_arraylist(i)(table_name_index)(table_value_index).ToString.Trim
                                    '    Catch ex As Exception
                                    '        MsgBox("Problem: " + table_value_index.ToString)
                                    '    End Try

                                    '    If cal_value_temp.Equals("{FORMULA_VALUE}") Then
                                    '        cal_value_temp = 0
                                    '    End If
                                    'End If
                                    'If getDescSource = 2 Then
                                    '    MsgBox("BAKA")
                                    'End If

                                End If
                                'MsgBox("From " + express_temp.ToString + " TO " + cal_value_temp.ToString)
                                If Not express_temp.Contains(".") Then
                                    calculation = calculation.Replace(express_temp, cal_value_temp)
                                End If
                            Next
                            cal_result = New DataTable().Compute(calculation, Nothing)
                            'MsgBox("Calculation:" + vbTab + calculation + vbNewLine + local_formula)

                        Next
                        If CDec(cal_result) <> 0 Then
                            value_arraylist(i)(row)(g) = Format(CDec(cal_result), "0.00").ToString
                        Else
                            value_arraylist(i)(row)(g) = cal_result.ToString
                        End If
                        'Try
                        '    value_arraylist(i)(row)(g) = Format(CDec(cal_result), "0.00").ToString
                        'Catch ex As Exception
                        '    MsgBox("Check it out : " + cal_result + vbNewLine + formula_temp)
                        'End Try

                    End If
                Next
            Next
        Next
        'Return
        'Dim te = ""
        'For Each temp As String In rangeQuo
        '    te += temp + vbTab
        'Next
        'MsgBox(te)
        'Return
        'For i As Integer = 0 To queryTable.Count - 1
        '    For row As Integer = 0 To dgvExcel.RowCount - 1
        '        Dim strs = ""
        '        For Each str As String In value_arraylist(i)(row)
        '            strs += str + vbTab
        '        Next
        '        MsgBox("Row " + row.ToString + vbNewLine + strs)
        '    Next
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
                                        'If Not express_temp.Contains(".") Then
                                        '    calculation = calculation.Replace(express_temp, cal_value_temp)
                                        'End If
                                        'Dim getDescSource = rangeQuo(table_name_index)
                                        'table_value_index = sql_format_arraylist(table_name_index).IndexOf(table_value_name)
                                        'cal_value_temp = value_arraylist(0)(table_name_index)(table_value_index).ToString.Trim
                                        'If cal_value_temp.Equals("{FORMULA_VALUE}") Then
                                        '    cal_value_temp = 0
                                        'End If
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

                                        ''valuearraylist(querytable,i)(table_name,0,1 % row)(table_index,values_index)
                                        'Dim getDescSource = CInt(rangeQuo(row).ToString.Split(".")(0))
                                        ''table_name_index = 1
                                        'table_name_index = getDescSource
                                        'Dim table_value_name = express_temp.Trim
                                        'table_value_index = sql_format_arraylist(i).IndexOf(table_value_name)
                                        'If table_value_index <> -1 Then
                                        '    Try
                                        '        cal_value_temp = value_arraylist(i)(table_name_index)(table_value_index).ToString.Trim
                                        '    Catch ex As Exception
                                        '        MsgBox("Problem: " + table_value_index.ToString)
                                        '    End Try

                                        '    If cal_value_temp.Equals("{FORMULA_VALUE}") Then
                                        '        cal_value_temp = 0
                                        '    End If
                                        'End If
                                        'If getDescSource = 2 Then
                                        '    MsgBox("BAKA")
                                        'End If
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
        For i As Integer = 0 To queryTable.Count - 1
            init()
            Using myConn = New SqlConnection("Data Source=" + serverName + ";" & "Initial Catalog=" + database + ";" + pwd_query)
                Using command As New SqlCommand("", myConn)
                    For row As Integer = 0 To dgvExcel.RowCount - 2
                        Using r As DataGridViewRow = dgvExcel.Rows(row)
                            If Not value_arraylist(i)(row)(0).Equals("{INVALID ARRAY}") Then
                                Dim query = ""
                                For g As Integer = 0 To value_arraylist(i)(row).count - 1
                                    Dim value_temp As String = value_arraylist(i)(row)(g).ToString
                                    If Not (value_temp.Equals("{._!@#$%^&*()}")) Then
                                        query += "'" + value_temp + "',"
                                    End If
                                Next
                                Dim command_text As String = queryTable(i)(2) + query
                                command_text = command_text.Substring(0, command_text.Length - 1) + ")"
                                MsgBox(command_text)
                                command.CommandText = command_text
                                myConn.Open()
                                command.ExecuteNonQuery()
                                myConn.Close()
                            End If

                        End Using
                    Next
                End Using
            End Using
        Next
    End Sub
End Class