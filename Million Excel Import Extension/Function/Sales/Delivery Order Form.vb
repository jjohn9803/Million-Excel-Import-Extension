Imports System.Data.SqlClient
Imports System.IO
Imports ClosedXML.Excel
Imports ExcelDataReader

Public Class Delivery_Order_Form
    Dim tables As DataTableCollection
    Private serverName As String
    Private database As String
    Private myConn As SqlConnection
    Private statusConnection As Boolean
    Private pwd_query As String
    Private import_type As String
    Private Sub Delivery_Order_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        Dim importType = "Delivery Order"
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
            queryTable.Add(New ArrayList)
            queryTable.Add(New ArrayList)
            queryTable.Add(New ArrayList)
            queryTable.Add(New ArrayList)
            queryTable.Add(New ArrayList)
            queryTable(0).add("Delivery Order") '0
            queryTable(0).add("sdo") '1
            queryTable(1).add("Delivery Order Desc")
            queryTable(1).add("sdodet")
            queryTable(2).add("Delivery Order Stock")
            queryTable(2).add("stock")
            queryTable(3).add("DO Product Serial No")
            queryTable(3).add("prodsn")
            queryTable(4).add("DO Stock Serial No")
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
            If Not value_arraylist(0)(row)(0).Equals("{INVALID ARRAY}") Then
                'batchno
                If value_arraylist(0)(row)(47).length = 2 Then
                    Dim date1 = value_arraylist(0)(row)(2)
                    Dim batchno = value_arraylist(0)(row)(47)
                    value_arraylist(0)(row)(47) = Year(date1).ToString.Substring(2) + Month(date1).ToString("00") + batchno
                End If
            End If

            'qtyrp
            If value_arraylist(1)(row)(10).Equals("{FORMULA_VALUE}") Then
                Dim qty = value_arraylist(1)(row)(9)
                value_arraylist(1)(row)(10) = qty
            End If

            'gross_amt
            If value_arraylist(1)(row)(25).Equals("{FORMULA_VALUE}") Then
                Dim qty = CDbl(value_arraylist(1)(row)(9))
                Dim price = CDbl(value_arraylist(1)(row)(12))
                Dim gross_amt = qty * price
                value_arraylist(1)(row)(25) = Math.Round(gross_amt, 2)
            End If

            'disamt1
            If value_arraylist(1)(row)(13).Equals("{FORMULA_VALUE}") Then
                Dim gross_amt = CDbl(value_arraylist(1)(row)(25))
                Dim disp1 = CDbl(value_arraylist(1)(row)(17))
                Dim disamt1 = gross_amt * (disp1 * 0.01)
                value_arraylist(1)(row)(13) = Math.Round(disamt1, 2)
            End If

            'disamt2
            If value_arraylist(1)(row)(14).Equals("{FORMULA_VALUE}") Then
                Dim gross_amt = CDbl(value_arraylist(1)(row)(25))
                Dim disamt1 = CDbl(value_arraylist(1)(row)(13))
                Dim disp2 = CDbl(value_arraylist(1)(row)(18))
                Dim disamt2 = (gross_amt - disamt1) * (disp2 * 0.01)
                value_arraylist(1)(row)(14) = Math.Round(disamt2, 2)
            End If

            'disamt3
            If value_arraylist(1)(row)(15).Equals("{FORMULA_VALUE}") Then
                Dim gross_amt = CDbl(value_arraylist(1)(row)(25))
                Dim disamt1 = CDbl(value_arraylist(1)(row)(13))
                Dim disamt2 = CDbl(value_arraylist(1)(row)(14))
                Dim disp3 = CDbl(value_arraylist(1)(row)(19))
                Dim disamt3 = (gross_amt - disamt1 - disamt2) * (disp3 * 0.01)
                value_arraylist(1)(row)(15) = Math.Round(disamt3, 2)
            End If

            'disamt
            If value_arraylist(1)(row)(16).Equals("{FORMULA_VALUE}") Then
                Dim disamt1 = CDbl(value_arraylist(1)(row)(13))
                Dim disamt2 = CDbl(value_arraylist(1)(row)(14))
                Dim disamt3 = CDbl(value_arraylist(1)(row)(15))
                Dim disamt = disamt1 + disamt2 + disamt3
                value_arraylist(1)(row)(16) = Math.Round(disamt, 2)
            End If

            'nett_amt
            If value_arraylist(1)(row)(26).Equals("{FORMULA_VALUE}") Then
                Dim gross_amt = CDbl(value_arraylist(1)(row)(25))
                Dim disamt = CDbl(value_arraylist(1)(row)(16))
                Dim nett_amt = gross_amt - disamt
                value_arraylist(1)(row)(26) = Math.Round(nett_amt, 2)
            End If

            'taxamt1
            If value_arraylist(1)(row)(20).Equals("{FORMULA_VALUE}") Then
                Dim nett_amt = CDbl(value_arraylist(1)(row)(26))
                Dim taxp1 As Double = 0
                If Not value_arraylist(1)(row)(24).Equals(String.Empty) Then
                    taxp1 = CDbl(value_arraylist(1)(row)(24))
                End If
                Dim taxamt1 = nett_amt * (taxp1 * 0.01)
                value_arraylist(1)(row)(20) = Math.Round(taxamt1, 2)
            End If

            'taxamt
            If value_arraylist(1)(row)(23).Equals("{FORMULA_VALUE}") Then
                Dim taxamt1 = CDbl(value_arraylist(1)(row)(20))
                Dim taxamt2 = CDbl(value_arraylist(1)(row)(21))
                Dim taxamt = taxamt1 + taxamt2
                value_arraylist(1)(row)(23) = Math.Round(taxamt, 2)
            End If

            'amt
            If value_arraylist(1)(row)(27).Equals("{FORMULA_VALUE}") Then
                Dim taxcode = value_arraylist(1)(row)(54)
                Dim taxinclude As Boolean = False
                If Not taxcode.Equals(String.Empty) Then
                    init()
                    myConn.Open()
                    Dim command = New SqlCommand("select taxmethod from gltax WHERE taxcode='" + taxcode + "'", myConn)
                    Dim reader As SqlDataReader = command.ExecuteReader
                    While reader.Read()
                        If reader.GetValue(0).ToString = "I" Then
                            taxinclude = True
                        End If
                    End While
                    myConn.Close()
                End If

                Dim gross_amt = CDbl(value_arraylist(1)(row)(25))
                Dim disamt = CDbl(value_arraylist(1)(row)(16))
                Dim amt = gross_amt - disamt
                If taxinclude Then
                    Dim taxamt = CDbl(value_arraylist(1)(row)(23))
                    amt -= taxamt
                End If
                value_arraylist(1)(row)(27) = Math.Round(amt, 2)
            End If

            'local converter
            Dim find_local = row
            For f = find_local To 0 Step -1
                If Not value_arraylist(0)(f)(0).Equals("{INVALID ARRAY}") Then
                    find_local = f
                    Exit For
                End If
            Next
            Dim fx_rate = CDbl(value_arraylist(0)(find_local)(17))

            'local_price
            If value_arraylist(1)(row)(28).Equals("{FORMULA_VALUE}") Then
                Dim price = CDbl(value_arraylist(1)(row)(12))
                Dim local_price = price * fx_rate
                value_arraylist(1)(row)(28) = local_price
            End If

            'local_gamt
            If value_arraylist(1)(row)(29).Equals("{FORMULA_VALUE}") Then
                Dim gross_amt = CDbl(value_arraylist(1)(row)(25))
                Dim local_gamt = gross_amt * fx_rate
                value_arraylist(1)(row)(29) = Math.Round(local_gamt, 2)
            End If

            'local_disamt
            If value_arraylist(1)(row)(30).Equals("{FORMULA_VALUE}") Then
                Dim disamt = CDbl(value_arraylist(1)(row)(16))
                Dim local_disamt = disamt * fx_rate
                value_arraylist(1)(row)(30) = Math.Round(local_disamt, 2)
            End If

            'local_namt
            If value_arraylist(1)(row)(31).Equals("{FORMULA_VALUE}") Then
                Dim local_gamt = CDbl(value_arraylist(1)(row)(29))
                Dim local_disamt = CDbl(value_arraylist(1)(row)(30))
                Dim local_namt = local_gamt - local_disamt
                value_arraylist(1)(row)(31) = Math.Round(local_namt, 2)
            End If

            'local_taxamt1
            If value_arraylist(1)(row)(32).Equals("{FORMULA_VALUE}") Then
                Dim taxamt1 = CDbl(value_arraylist(1)(row)(20))
                Dim local_taxamt1 = taxamt1 * fx_rate
                value_arraylist(1)(row)(32) = Math.Round(local_taxamt1, 2)
            End If

            'local_taxamt2
            If value_arraylist(1)(row)(33).Equals("{FORMULA_VALUE}") Then
                Dim taxamt2 = CDbl(value_arraylist(1)(row)(21))
                Dim local_taxamt2 = taxamt2 * fx_rate
                value_arraylist(1)(row)(33) = Math.Round(local_taxamt2, 2)
            End If

            'local_taxamtadj1
            If value_arraylist(1)(row)(34).Equals("{FORMULA_VALUE}") Then
                Dim taxamtadj1 = CDbl(value_arraylist(1)(row)(22))
                Dim local_taxamtadj1 = taxamtadj1 * fx_rate
                value_arraylist(1)(row)(34) = Math.Round(local_taxamtadj1, 2)
            End If

            'local_taxamt
            If value_arraylist(1)(row)(35).Equals("{FORMULA_VALUE}") Then
                Dim taxamt = CDbl(value_arraylist(1)(row)(23))
                Dim local_taxamt = taxamt * fx_rate
                value_arraylist(1)(row)(35) = Math.Round(local_taxamt, 2)
            End If

            'local_amt
            If value_arraylist(1)(row)(36).Equals("{FORMULA_VALUE}") Then
                Dim amt = CDbl(value_arraylist(1)(row)(27))
                Dim local_amt = amt * fx_rate
                value_arraylist(1)(row)(36) = Math.Round(local_amt, 2)
            End If

            'local_amtrp
            If value_arraylist(1)(row)(37).Equals("{FORMULA_VALUE}") Then
                Dim local_amt = value_arraylist(1)(row)(36)
                value_arraylist(1)(row)(37) = Math.Round(local_amt, 2)
            End If

            'local_mcamt1
            If value_arraylist(1)(row)(39).Equals("{FORMULA_VALUE}") Then
                Dim mcamt1 = CDbl(value_arraylist(1)(row)(38))
                Dim local_mcamt1 = mcamt1 * fx_rate
                value_arraylist(1)(row)(39) = Math.Round(local_mcamt1, 2)
            End If

            'stock.qty
            If value_arraylist(2)(row)(12).Equals("{FORMULA_VALUE}") Then
                Dim qty1 = CDbl(value_arraylist(1)(row)(9))
                Dim qty = qty1 * -1
                value_arraylist(2)(row)(12) = Math.Round(qty, 2)
            End If

            'stock.local_amount
            If value_arraylist(2)(row)(15).Equals("{FORMULA_VALUE}") Then
                Dim qty = CDbl(value_arraylist(2)(row)(12))
                Dim cost = CDbl(value_arraylist(2)(row)(13))
                Dim local_amount = qty * cost
                value_arraylist(2)(row)(15) = Math.Round(local_amount, 2)
            End If

        Next

        'Hardcore Formula sdo
        For row As Integer = 0 To dgvExcel.RowCount - 1
            If Not value_arraylist(0)(row)(0).Equals("{INVALID ARRAY}") Then

                Dim myTarget As New ArrayList
                For Each target In rangeQuo
                    If target.ToString.Split(".")(0).Equals(row.ToString) Then
                        Dim targetRow = CInt(target.ToString.Split(".")(1))
                        myTarget.Add(targetRow)
                    End If
                Next

                'sdo.address set address to default delivery address
                If value_arraylist(0)(row)(11).ToString.Trim.Equals(String.Empty) Then
                    Dim custcode = value_arraylist(0)(row)(9).ToString.Trim
                    myConn.Open()
                    Dim addrcommand = New SqlCommand("select daddr from customer WHERE custcode='" + custcode + "'", myConn)
                    Dim addrreader As SqlDataReader = addrcommand.ExecuteReader
                    While addrreader.Read()
                        value_arraylist(0)(row)(11) = addrreader.GetValue(0)
                    End While
                    myConn.Close()
                End If

                'sdo.discount
                If value_arraylist(0)(row)(20).Equals("{FORMULA_VALUE}") Then
                    Dim discount1 = CDbl(value_arraylist(0)(row)(18))
                    Dim discount2 = CDbl(value_arraylist(0)(row)(19))
                    Dim discount = discount1 * discount2
                    value_arraylist(0)(row)(20) = Math.Round(discount, 2)
                End If

                'sdo.taxable
                If value_arraylist(0)(row)(26).Equals("{FORMULA_VALUE}") Then
                    'get query target

                    Dim taxable As Double = 0
                    For Each targetRow As Integer In myTarget
                        Dim nett_amt = value_arraylist(1)(targetRow)(26)
                        Dim taxcode = value_arraylist(1)(targetRow)(54)
                        If taxcode.ToString.Trim.ToCharArray.Count > 0 Then
                            taxable += nett_amt
                        End If
                    Next
                    value_arraylist(0)(row)(26) = Math.Round(taxable, 2)
                End If

                'sdo.tax
                If value_arraylist(0)(row)(30).Equals("{FORMULA_VALUE}") Then
                    Dim tax As Double = 0
                    For Each targetRow As Integer In myTarget
                        Dim taxamt1 = value_arraylist(1)(targetRow)(20)
                        Dim taxamt2 = value_arraylist(1)(targetRow)(21)
                        tax += taxamt1 + taxamt2
                    Next
                    value_arraylist(0)(row)(30) = Math.Round(tax, 2)
                End If

                'sdo.subtotal
                If value_arraylist(0)(row)(32).Equals("{FORMULA_VALUE}") Then
                    Dim subtotal As Double = 0
                    For Each targetRow As Integer In myTarget
                        Dim amt = value_arraylist(1)(targetRow)(27)
                        subtotal += amt
                    Next
                    value_arraylist(0)(row)(32) = Math.Round(subtotal, 2)
                End If

                'sdo.nett
                If value_arraylist(0)(row)(33).Equals("{FORMULA_VALUE}") Then
                    Dim nett As Double = 0
                    For Each targetRow As Integer In myTarget
                        Dim nett_amt = value_arraylist(1)(targetRow)(26)
                        nett += nett_amt
                    Next
                    value_arraylist(0)(row)(33) = Math.Round(nett, 2)
                End If

                'sdo.total
                If value_arraylist(0)(row)(34).Equals("{FORMULA_VALUE}") Then
                    Dim subtotal As Double = value_arraylist(0)(row)(32)
                    Dim tax As Double = value_arraylist(0)(row)(30)
                    'MsgBox(subtotal.ToString + vbTab + tax.ToString)
                    Dim total = subtotal + tax
                    value_arraylist(0)(row)(34) = Math.Round(total, 2)
                End If

                'sdo.local_gross
                If value_arraylist(0)(row)(35).Equals("{FORMULA_VALUE}") Then
                    Dim local_gross As Double = 0
                    For Each targetRow As Integer In myTarget
                        Dim local_amt = value_arraylist(1)(targetRow)(36)
                        local_gross += local_amt
                    Next
                    value_arraylist(0)(row)(35) = Math.Round(local_gross, 2)
                End If

                Dim fx_rate = value_arraylist(0)(row)(17)

                'sdo.local_discount
                If value_arraylist(0)(row)(36).Equals("{FORMULA_VALUE}") Then
                    Dim discount = value_arraylist(0)(row)(20)
                    Dim local_discount = discount * fx_rate
                    value_arraylist(0)(row)(36) = Math.Round(local_discount, 2)
                End If

                'sdo.local_nett
                If value_arraylist(0)(row)(37).Equals("{FORMULA_VALUE}") Then
                    Dim nett = value_arraylist(0)(row)(33)
                    Dim local_nett = nett * fx_rate
                    value_arraylist(0)(row)(37) = Math.Round(local_nett, 2)
                End If

                'sdo.local_tax1
                If value_arraylist(0)(row)(38).Equals("{FORMULA_VALUE}") Then
                    Dim tax1 = value_arraylist(0)(row)(27)
                    Dim local_tax1 = tax1 * fx_rate
                    value_arraylist(0)(row)(38) = Math.Round(local_tax1, 2)
                End If

                'sdo.local_tax2
                If value_arraylist(0)(row)(39).Equals("{FORMULA_VALUE}") Then
                    Dim tax2 = value_arraylist(0)(row)(28)
                    Dim local_tax2 = tax2 * fx_rate
                    value_arraylist(0)(row)(39) = Math.Round(local_tax2, 2)
                End If

                'sdo.local_tax
                If value_arraylist(0)(row)(40).Equals("{FORMULA_VALUE}") Then
                    Dim tax = value_arraylist(0)(row)(30)
                    Dim local_tax = tax * fx_rate
                    value_arraylist(0)(row)(40) = Math.Round(local_tax, 2)
                End If

                'sdo.local_rndoff
                If value_arraylist(0)(row)(41).Equals("{FORMULA_VALUE}") Then
                    Dim rndoff = value_arraylist(0)(row)(31)
                    Dim local_rndoff = rndoff * fx_rate
                    value_arraylist(0)(row)(41) = Math.Round(local_rndoff, 2)
                End If

                'sdo.local_total
                If value_arraylist(0)(row)(42).Equals("{FORMULA_VALUE}") Then
                    Dim total = value_arraylist(0)(row)(34)
                    Dim local_total = total * fx_rate
                    value_arraylist(0)(row)(42) = Math.Round(local_total, 2)
                End If

                'sdo.local_totalrp
                If value_arraylist(0)(row)(43).Equals("{FORMULA_VALUE}") Then
                    Dim subtotal = value_arraylist(0)(row)(32)
                    Dim local_totalrp = subtotal * fx_rate
                    value_arraylist(0)(row)(43) = Math.Round(local_totalrp, 2)
                End If

            End If
        Next

        'End Hardcode Formula

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
                            Dim nett_amt = value_arraylist(1)(targetRow)(26)
                            Dim taxcode = value_arraylist(1)(targetRow)(54)
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
        'myConn = New SqlConnection("Data Source=" + serverName + ";" & "Initial Catalog=" + database + ";" + pwd_query)
        For row As Integer = 0 To dgvExcel.RowCount - 1
            Dim table As String
            Dim value_name As String
            Dim value As String
            'Delivery Order
            If Not value_arraylist(0)(row)(0).Equals("{INVALID ARRAY}") Then
                'sdo.doc_no / duplicate
                table = "sdo"
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

                'custaddr.addr / exist
                table = "custaddr"
                value_name = "addr"
                value = dgvExcel.Rows(row).Cells("Delivery Address").Value.ToString
                Dim value2 = value_arraylist(0)(row)(9) 'custcode
                If Not value.Trim.Equals(String.Empty) And Not value2.Trim.Equals(String.Empty) Then
                    value = value.Replace(vbLf, vbCrLf)
                    myConn.Open()
                    Dim exist_value As Boolean = False
                    Dim command = New SqlCommand("SELECT * FROM " + table + " WHERE cast(" + value_name + " as varchar(MAX)) ='" + value + "' AND custcode ='" + value2 + "'", myConn)
                    Dim reader As SqlDataReader = command.ExecuteReader
                    While reader.Read()
                        exist_value = True
                        value_arraylist(0)(row)(12) = reader.GetValue(0).ToString 'update addrkey
                    End While
                    myConn.Close()
                    If Not exist_value Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "'(" + value2 + ") is not found in the database (" + table + ")!" + vbNewLine
                    End If
                End If

                'accmgr.accmgr_id / exist
                table = "accmgr"
                value_name = "accmgr_id"
                value = value_arraylist(0)(row)(15)
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

                'glbatch.batchno / exist
                table = "glbatch"
                value_name = "batchno"
                value = value_arraylist(0)(row)(47)
                If Not value.Trim.Equals(String.Empty) Then
                    If Not existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                End If

                'project.projcode / exist
                table = "project"
                value_name = "projcode"
                value = value_arraylist(0)(row)(48)
                If Not value.Trim.Equals(String.Empty) Then
                    If Not existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                End If

                'deptment.deptcode / exist
                table = "deptment"
                value_name = "deptcode"
                value = value_arraylist(0)(row)(49)
                If Not value.Trim.Equals(String.Empty) Then
                    If Not existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                End If
            End If

            'Delivery Order Desc
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

                'prodbatch.batchcode / exist
                table = "prodbatch"
                value_name = "batchcode"
                value = value_arraylist(1)(row)(56)
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

            'Delivery Order Stock
            If Not value_arraylist(2)(row)(0).Equals(String.Empty) Then
                'product.prodcode / exist
                table = "product"
                value_name = "prodcode"
                value = value_arraylist(2)(row)(0)
                If Not value.Trim.Equals(String.Empty) Then
                    If Not existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                End If

                'defdocno.doc_type / exist
                table = "defdocno"
                value_name = "doc_type"
                value = value_arraylist(2)(row)(1)
                If Not value.Trim.Equals(String.Empty) Then
                    If Not existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                End If
            End If

            'Delivery Order Serial No
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

        'For i As Integer = 0 To queryTable.Count - 1
        '    For row As Integer = 0 To dgvExcel.RowCount - 1
        '        Dim strs = ""
        '        For Each str As String In value_arraylist(i)(row)
        '            strs += str + vbTab
        '        Next
        '        MsgBox("Row " + row.ToString + vbNewLine + strs)
        '    Next
        'Next

        init()
        'myConn = New SqlConnection("Data Source=" + serverName + ";" & "Initial Catalog=" + database + ";" + pwd_query)
        Dim exist_serial As Boolean = False
        Dim msg_serial As String = ""
        For row As Integer = 0 To dgvExcel.RowCount - 1
            If Not value_arraylist(3)(row)(1).ToString.Trim.Equals(String.Empty) Then
                Dim serialnos As New List(Of String)(value_arraylist(3)(row)(1).ToString.Trim.Split(","c))
                If row = 0 Then
                    If serialnos.Count <> CInt(value_arraylist(1)(row)(9)) Then
                        MsgBox("The serial no quantity (" + serialnos.Count.ToString + ") does not matched with items quantity(" + value_arraylist(1)(row)(9).ToString + ")" + vbNewLine + "The operation has been stopped!", MsgBoxStyle.Exclamation)
                        Return
                    End If
                End If
                For sn = 0 To serialnos.Count - 1
                    Dim serialno As String = serialnos(sn)
                    myConn.Open()
                    Dim sncommand = New SqlCommand("select stocksn.serialno from stocksn LEFT JOIN prodsn ON stocksn.serialno = prodsn.serialno WHERE (stocksn.serialno = '" + serialno + "' AND stocksn.qty='-1') OR (prodsn.serialno = '" + serialno + "' AND prodsn.qty='-1')", myConn)
                    Dim snreader As SqlDataReader = sncommand.ExecuteReader
                    While snreader.Read()
                        If Not msg_serial.Contains(snreader.GetValue(0).ToString.Trim) Then
                            msg_serial += snreader.GetValue(0).ToString.Trim + vbTab
                        End If
                        exist_serial = True
                    End While
                    myConn.Close()
                Next
            End If
        Next
        If exist_serial Then
            MsgBox("The following serial no has been used:" + vbNewLine + msg_serial + vbNewLine + "The operation has been stopped!", MsgBoxStyle.Exclamation)
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
                Dim serialnos As New List(Of String)(value_arraylist(3)(row)(1).ToString.Trim.Split(","c))
                For sn = 0 To serialnos.Count - 1
                    Dim serialno As String = serialnos(sn)
                    Dim qty = "-1"
                    Dim location = value_arraylist(3)(row)(4)
                    Dim doc_no = value_arraylist(3)(row)(8)
                    Dim line_no = value_arraylist(3)(row)(9)
                    Dim doc_date = Convert.ToDateTime(value_arraylist(3)(row)(10)).ToString("yyyy-MM-dd HH:mm:ss")
                    Dim procode = value_arraylist(3)(row)(0)
                    Dim serialNoProdCommand As String = "UPDATE prodsn SET "
                    Dim serialNoColumns = "qty='" + qty + "',"
                    serialNoColumns += "location='" + location + "',"
                    serialNoColumns += "doc_type='DO',"
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
                    serialNoStockdCommand += procode + "','" + serialno + "','DO','" + doc_no + "','" + line_no + "','" + doc_date + "','" + qty + "','" + location + "')"
                    Dim command2 = New SqlCommand(serialNoStockdCommand, myConn)
                    command2.ExecuteNonQuery()
                    'MsgBox(serialNoStockdCommand)
                    rowInsertNum += 1
                    myConn.Close()
                Next
            End If
        Next

        'Quotation only end
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
                                    query += "'" + Date.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                                    value_arraylist(i)(row)(g) = Date.Now.ToString("yyyy-MM-dd HH:mm:ss")
                                ElseIf data_type_arraylist(i)(g).ToString.Trim.Contains("date") Then
                                    query += "'" + Convert.ToDateTime(value_temp).ToString("yyyy-MM-dd HH:mm:ss") + "',"
                                    value_arraylist(i)(row)(g) = Convert.ToDateTime(value_temp).ToString("yyyy-MM-dd HH:mm:ss")
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
        Next

        'Delivery Order stock
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
                Dim command_temp = New SqlCommand("SELECT TOP 1 dkey FROM sdodet WHERE doc_no ='" + value_arraylist(2)(row)(2) + "' AND line_no ='" + value_arraylist(2)(row)(4) + "'", myConn)
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
                insertArray.Add(value_arraylist(0)(mySource)(10)) 'doc_desp
                insertArray.Add(Function_Form.getNull(0)) 'doc_desp2
                insertArray.Add(value_arraylist(0)(mySource)(9)) 'custcode
                insertArray.Add(Function_Form.getNull(0)) 'suppcode
                insertArray.Add(value_arraylist(0)(mySource)(7)) 'refno
                insertArray.Add(value_arraylist(0)(mySource)(8)) 'refno2
                insertArray.Add(value_arraylist(1)(row)(9)) 'qty
                insertArray.Add(Function_Form.getNull(3)) 'cost
                insertArray.Add(value_arraylist(1)(row)(12)) 'price
                insertArray.Add(Function_Form.getNull(3)) 'local_amount
                insertArray.Add(Function_Form.getNull(3)) 'utd_cost
                insertArray.Add(value_arraylist(1)(row)(55)) 'location
                insertArray.Add(value_arraylist(1)(row)(56)) 'batchcode
                insertArray.Add(value_arraylist(1)(row)(58)) 'projcode
                insertArray.Add(value_arraylist(1)(row)(59)) 'deptcode
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
                command.ExecuteNonQuery()
                rowInsertNum += 1
                myConn.Close()
            End If
        Next

        Function_Form.promptImportSuccess(rowInsertNum, rowUpdateNum)
        Function_Form.printExcelResult("Delivery_Order", queryTable, value_arraylist, sql_format_arraylist, dgvExcel)
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