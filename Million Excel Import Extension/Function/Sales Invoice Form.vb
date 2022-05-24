Imports System.Data.SqlClient
Imports System.IO
Imports ClosedXML.Excel
Imports ExcelDataReader
Public Class Sales_Invoice_Form
    Dim tables As DataTableCollection
    Private serverName As String
    Private database As String
    Private myConn As SqlConnection
    Private statusConnection As Boolean
    Private pwd_query As String
    Private import_type As String
    Private Sub Sales_Invoice_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        Dim importType = "Sales Invoice"
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
            For i = 0 To 8
                queryTable.Add(New ArrayList)
            Next
            queryTable(0).add("Sales Invoice") '0
            queryTable(0).add("sinv") '1
            queryTable(1).add("Sales Invoice Desc")
            queryTable(1).add("sinvdet")
            queryTable(2).add("Sales Invoice Stock")
            queryTable(2).add("stock")
            queryTable(3).add("Sales Invoice AR") '0
            queryTable(3).add("ar") '1
            queryTable(4).add("Sales Invoice GL")
            queryTable(4).add("gl")
            queryTable(5).add("Sales Invoice GL Off")
            queryTable(5).add("gloff")
            queryTable(6).add("Sales Invoice GL Audit")
            queryTable(6).add("glaudit")
            queryTable(7).add("SI Product Serial No")
            queryTable(7).add("prodsn")
            queryTable(8).add("SI Stock Serial No")
            queryTable(8).add("stocksn")
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
            'sinv
            If Not value_arraylist(0)(row)(0).Equals("{INVALID ARRAY}") Then
                'batchno
                If value_arraylist(0)(row)(48).length = 2 Then
                    Dim date1 = value_arraylist(0)(row)(2)
                    Dim batchno = value_arraylist(0)(row)(48)
                    value_arraylist(0)(row)(48) = Year(date1).ToString.Substring(2) + Month(date1).ToString("00") + batchno
                End If
            End If

            'sinvdet
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
            If value_arraylist(1)(row)(28).Equals("{FORMULA_VALUE}") Then
                Dim taxcode = dgvExcel.Rows(row).Cells("Tax Code").Value.ToString.Trim
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
            Dim fx_rate = CDbl(value_arraylist(0)(find_local)(18))

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

        'Hardcore Formula sinv
        For row As Integer = 0 To dgvExcel.RowCount - 1
            If Not value_arraylist(0)(row)(0).Equals("{INVALID ARRAY}") Then

                Dim myTarget As New ArrayList
                For Each target In rangeQuo
                    If target.ToString.Split(".")(0).Equals(row.ToString) Then
                        Dim targetRow = CInt(target.ToString.Split(".")(1))
                        myTarget.Add(targetRow)
                    End If
                Next

                'sinv.discount
                If value_arraylist(0)(row)(21).Equals("{FORMULA_VALUE}") Then
                    Dim discount1 = CDbl(value_arraylist(0)(row)(19))
                    Dim discount2 = CDbl(value_arraylist(0)(row)(20))
                    Dim discount = discount1 * discount2
                    value_arraylist(0)(row)(21) = Math.Round(discount, 2)
                End If

                'sinv.taxable
                If value_arraylist(0)(row)(27).Equals("{FORMULA_VALUE}") Then
                    'get query target

                    Dim taxable As Double = 0
                    For Each targetRow As Integer In myTarget
                        Dim nett_amt = value_arraylist(1)(targetRow)(26)
                        Dim taxcode = value_arraylist(1)(targetRow)(54)
                        If taxcode.ToString.Trim.ToCharArray.Count > 0 Then
                            taxable += nett_amt
                        End If
                    Next
                    value_arraylist(0)(row)(27) = Math.Round(taxable, 2)
                End If

                'sinv.tax
                If value_arraylist(0)(row)(31).Equals("{FORMULA_VALUE}") Then
                    Dim tax As Double = 0
                    For Each targetRow As Integer In myTarget
                        Dim taxamt1 = value_arraylist(1)(targetRow)(20)
                        Dim taxamt2 = value_arraylist(1)(targetRow)(21)
                        tax += taxamt1 + taxamt2
                    Next
                    value_arraylist(0)(row)(31) = Math.Round(tax, 2)
                End If

                'sinv.subtotal
                If value_arraylist(0)(row)(33).Equals("{FORMULA_VALUE}") Then
                    Dim subtotal As Double = 0
                    For Each targetRow As Integer In myTarget
                        Dim amt = value_arraylist(1)(targetRow)(27)
                        subtotal += amt
                    Next
                    value_arraylist(0)(row)(33) = Math.Round(subtotal, 2)
                End If

                'sinv.nett
                If value_arraylist(0)(row)(34).Equals("{FORMULA_VALUE}") Then
                    Dim nett As Double = 0
                    For Each targetRow As Integer In myTarget
                        Dim nett_amt = value_arraylist(1)(targetRow)(26)
                        nett += nett_amt
                    Next
                    value_arraylist(0)(row)(34) = Math.Round(nett, 2)
                End If

                'sinv.total
                If value_arraylist(0)(row)(35).Equals("{FORMULA_VALUE}") Then
                    Dim subtotal As Double = value_arraylist(0)(row)(33)
                    Dim tax As Double = value_arraylist(0)(row)(31)
                    'MsgBox(subtotal.ToString + vbTab + tax.ToString)
                    Dim total = subtotal + tax
                    value_arraylist(0)(row)(35) = Math.Round(total, 2)
                End If

                'sinv.local_gross
                If value_arraylist(0)(row)(36).Equals("{FORMULA_VALUE}") Then
                    Dim local_gross As Double = 0
                    For Each targetRow As Integer In myTarget
                        Dim local_amt = value_arraylist(1)(targetRow)(36)
                        local_gross += local_amt
                    Next
                    value_arraylist(0)(row)(36) = Math.Round(local_gross, 2)
                End If

                Dim fx_rate = value_arraylist(0)(row)(18)

                'sinv.local_discount
                If value_arraylist(0)(row)(37).Equals("{FORMULA_VALUE}") Then
                    Dim discount = value_arraylist(0)(row)(21)
                    Dim local_discount = discount * fx_rate
                    value_arraylist(0)(row)(37) = Math.Round(local_discount, 2)
                End If

                'sinv.local_nett
                If value_arraylist(0)(row)(38).Equals("{FORMULA_VALUE}") Then
                    Dim nett = value_arraylist(0)(row)(34)
                    Dim local_nett = nett * fx_rate
                    value_arraylist(0)(row)(38) = Math.Round(local_nett, 2)
                End If

                'sinv.local_tax1
                If value_arraylist(0)(row)(39).Equals("{FORMULA_VALUE}") Then
                    Dim tax1 = value_arraylist(0)(row)(28)
                    Dim local_tax1 = tax1 * fx_rate
                    value_arraylist(0)(row)(39) = Math.Round(local_tax1, 2)
                End If

                'sinv.local_tax2
                If value_arraylist(0)(row)(40).Equals("{FORMULA_VALUE}") Then
                    Dim tax2 = value_arraylist(0)(row)(29)
                    Dim local_tax2 = tax2 * fx_rate
                    value_arraylist(0)(row)(40) = Math.Round(local_tax2, 2)
                End If

                'sinv.local_tax
                If value_arraylist(0)(row)(41).Equals("{FORMULA_VALUE}") Then
                    Dim tax = value_arraylist(0)(row)(31)
                    Dim local_tax = tax * fx_rate
                    value_arraylist(0)(row)(41) = Math.Round(local_tax, 2)
                End If

                'sinv.local_rndoff
                If value_arraylist(0)(row)(42).Equals("{FORMULA_VALUE}") Then
                    Dim rndoff = value_arraylist(0)(row)(32)
                    Dim local_rndoff = rndoff * fx_rate
                    value_arraylist(0)(row)(42) = Math.Round(local_rndoff, 2)
                End If

                'sinv.local_total
                If value_arraylist(0)(row)(43).Equals("{FORMULA_VALUE}") Then
                    Dim total = value_arraylist(0)(row)(35)
                    Dim local_total = total * fx_rate
                    value_arraylist(0)(row)(43) = Math.Round(local_total, 2)
                End If

                'sinv.local_totalrp
                If value_arraylist(0)(row)(44).Equals("{FORMULA_VALUE}") Then
                    Dim subtotal = value_arraylist(0)(row)(33)
                    Dim local_totalrp = subtotal * fx_rate
                    value_arraylist(0)(row)(44) = Math.Round(local_totalrp, 2)
                End If

            End If
        Next

        'End Hardcode Formula

        'hardcore exist checker
        Dim execute_valid As Boolean = True
        Dim exist_result As String = ""
        init()
        'myConn = New SqlConnection("Data Source=" + serverName + ";" & "Initial Catalog=" + database + ";" + pwd_query)
        For row As Integer = 0 To dgvExcel.RowCount - 1
            Dim table As String
            Dim value_name As String
            Dim value As String
            'Sales Invoice
            If Not value_arraylist(0)(row)(0).Equals("{INVALID ARRAY}") Then
                'sinv.doc_no / duplicate
                table = "sinv"
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
                value = value_arraylist(0)(row)(10)
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
                Dim value2 = value_arraylist(0)(row)(10) 'custcode
                If Not value.Trim.Equals(String.Empty) And Not value2.Trim.Equals(String.Empty) Then
                    value = value.Replace(vbLf, vbCrLf)
                    myConn.Open()
                    Dim exist_value As Boolean = False
                    Dim command = New SqlCommand("SELECT * FROM " + table + " WHERE cast(" + value_name + " as varchar(MAX)) ='" + value + "' AND custcode ='" + value2 + "'", myConn)
                    Dim reader As SqlDataReader = command.ExecuteReader
                    While reader.Read()
                        exist_value = True
                        value_arraylist(0)(row)(13) = reader.GetValue(0).ToString 'update addrkey
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
                value = value_arraylist(0)(row)(16)
                If Not value.Trim.Equals(String.Empty) Then
                    If Not existed_checker(table, value_name, value) Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                End If

                'currency.curr_code / exist
                table = "currency"
                value_name = "curr_code"
                value = value_arraylist(0)(row)(17)
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

                'gldata.accno / exist (Cash, Bank) Acc_No1
                If Not value_arraylist(0)(row)(73).ToString.Trim.Equals(String.Empty) Then
                    table = "gldata"
                    value_name = "accno"
                    value = value_arraylist(0)(row)(73)
                    myConn.Open()
                    Dim sncommand = New SqlCommand("select * from " + table + " WHERE (classify = 'C' OR classify = 'B') AND accno='" + value + "'", myConn)
                    Dim snreader As SqlDataReader = sncommand.ExecuteReader
                    Dim exist_acc = False
                    While snreader.Read()
                        exist_acc = True
                    End While
                    If exist_acc = False Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                    myConn.Close()
                Else
                    'payment set empty
                    value_arraylist(0)(row)(73) = Function_Form.getNull(0)
                    value_arraylist(0)(row)(74) = Function_Form.getNull(0)
                    value_arraylist(0)(row)(75) = Function_Form.getNull(3)
                End If

                'gldata.accno / exist (Cash, Bank) Acc_No2
                If Not value_arraylist(0)(row)(78).ToString.Trim.Equals(String.Empty) Then
                    table = "gldata"
                    value_name = "accno"
                    value = value_arraylist(0)(row)(78)
                    myConn.Open()
                    Dim sncommand = New SqlCommand("select * from " + table + " WHERE (classify = 'C' OR classify = 'B') AND accno='" + value + "'", myConn)
                    Dim snreader As SqlDataReader = sncommand.ExecuteReader
                    Dim exist_acc = False
                    While snreader.Read()
                        exist_acc = True
                    End While
                    If exist_acc = False Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                    myConn.Close()
                Else
                    'payment set empty
                    value_arraylist(0)(row)(78) = Function_Form.getNull(0)
                    value_arraylist(0)(row)(79) = Function_Form.getNull(0)
                    value_arraylist(0)(row)(80) = Function_Form.getNull(3)
                End If

                'gldata.accno / exist (Cash, Bank) Acc_No3
                If Not value_arraylist(0)(row)(83).ToString.Trim.Equals(String.Empty) Then
                    table = "gldata"
                    value_name = "accno"
                    value = value_arraylist(0)(row)(83)
                    myConn.Open()
                    Dim sncommand = New SqlCommand("select * from " + table + " WHERE (classify = 'C' OR classify = 'B') AND accno='" + value + "'", myConn)
                    Dim snreader As SqlDataReader = sncommand.ExecuteReader
                    Dim exist_acc = False
                    While snreader.Read()
                        exist_acc = True
                    End While
                    If exist_acc = False Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                    myConn.Close()
                Else
                    'payment set empty
                    value_arraylist(0)(row)(83) = Function_Form.getNull(0)
                    value_arraylist(0)(row)(84) = Function_Form.getNull(0)
                    value_arraylist(0)(row)(85) = Function_Form.getNull(3)
                End If

                'gldata.accno / exist (Cash, Bank) Acc_No4
                If Not value_arraylist(0)(row)(86).ToString.Trim.Equals(String.Empty) Then
                    table = "gldata"
                    value_name = "accno"
                    value = value_arraylist(0)(row)(86)
                    myConn.Open()
                    Dim sncommand = New SqlCommand("select * from " + table + " WHERE (classify = 'C' OR classify = 'B') AND accno='" + value + "'", myConn)
                    Dim snreader As SqlDataReader = sncommand.ExecuteReader
                    Dim exist_acc = False
                    While snreader.Read()
                        exist_acc = True
                    End While
                    If exist_acc = False Then
                        execute_valid = False
                        exist_result += value_name + " '" + value + "' is not found in the database (" + table + ")!" + vbNewLine
                    End If
                    myConn.Close()
                Else
                    'payment set empty
                    value_arraylist(0)(row)(86) = Function_Form.getNull(0)
                    value_arraylist(0)(row)(87) = Function_Form.getNull(0)
                    value_arraylist(0)(row)(88) = Function_Form.getNull(3)
                End If
            End If

            'Sales Invoice Desc
            If Not value_arraylist(1)(row)(0).Equals(String.Empty) Then
                'sinvdet.doc_no / duplicate
                'table = "sinvdet"
                'value_name = "doc_no"
                'value = value_arraylist(1)(row)(2)
                'If Not value.Trim.Equals(String.Empty) Then
                '    If existed_checker(table, value_name, value) Then
                '        execute_valid = False
                '        exist_result += value_name + " '" + value + "' already existed in the database (" + table + ")!" + vbNewLine
                '    End If
                '    If Function_Form.repeatedExcelCell(dgvExcel, excel_format_arraylist(1)(2), value, row) Then
                '        execute_valid = False
                '        If Not exist_result.Contains(excel_format_arraylist(1)(2) + " '" + value + "' is repeated!") Then
                '            exist_result += excel_format_arraylist(1)(2) + " '" + value + "' is repeated!" + vbNewLine
                '        End If
                '    End If
                'End If

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

            'Sales Invoice Stock
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
            If Not value_arraylist(7)(row)(0).Equals(String.Empty) Then
                'prodsn.serialno / exist
                table = "prodsn"
                value_name = "serialno"
                Dim values As New List(Of String)(value_arraylist(7)(row)(1).ToString.Trim.Split(","c))
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
                Dim values2 As New List(Of String)(value_arraylist(7)(row)(1).ToString.Trim.Split(","c))
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
        'myConn = New SqlConnection("Data Source=" + serverName + ";" & "Initial Catalog=" + database + ";" + pwd_query)
        Dim exist_serial As Boolean = False
        Dim msg_serial As New ArrayList
        For row As Integer = 0 To dgvExcel.RowCount - 1
            If Not value_arraylist(7)(row)(1).ToString.Trim.Equals(String.Empty) Then
                Dim serialnos As New List(Of String)(value_arraylist(7)(row)(1).ToString.Trim.Split(","c))
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
                        'MsgBox("hey1" + snreader.GetValue(snreader.GetOrdinal("serialno")))
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
                        'MsgBox("hey2" + prodsnreader.GetValue(prodsnreader.GetOrdinal("serialno")))
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
        'For i As Integer = 0 To 1
        '    For row As Integer = 0 To dgvExcel.RowCount - 1
        '        Dim strs = ""
        '        For Each str As String In value_arraylist(i)(row)
        '            strs += str + vbTab
        '        Next
        '        MsgBox("Row " + row.ToString + vbNewLine + strs)
        '    Next
        'Next

        Dim confirmImport As DialogResult = MsgBox("Are you sure to import data?", MsgBoxStyle.YesNo, "")
        If confirmImport = DialogResult.No Then
            Return
        End If

        Dim rowInsertNum = 0
        Dim arseq As New ArrayList
        Dim gloffseq As New ArrayList
        'gl
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

                Dim query_temp As New ArrayList
                query_temp.Add("I") 'billtype
                query_temp.Add(Function_Form.getNull(0)) 'remark1
                query_temp.Add(Function_Form.getNull(0)) 'remark2
                query_temp.Add(Function_Form.getNull(0)) 'cheque_no
                query_temp.Add(Function_Form.getNull(1)) 'chqrc_date
                query_temp.Add(Function_Form.getNull(1)) 'koff_date
                query_temp.Add(Function_Form.getNull(1)) 'recon_date
                query_temp.Add(Function_Form.getNull(3)) 'recon_flag
                query_temp.Add(value_arraylist(0)(row)(16)) 'accmgr_id
                query_temp.Add(Function_Form.getNull(0)) 'lkdoc_type
                query_temp.Add(Function_Form.getNull(0)) 'lkdoc_no
                query_temp.Add(Function_Form.getNull(3)) 'lkseq
                query_temp.Add("1") 'lock
                query_temp.Add(Function_Form.getNull(3)) 'void
                query_temp.Add(Function_Form.getNull(3)) 'exported
                query_temp.Add(Function_Form.getNull(0)) 'entry
                query_temp.Add(Function_Form.getNull(3)) 'fastentry
                query_temp.Add(Function_Form.getNull(3)) 'followdesp
                query_temp.Add(Function_Form.getNull(3)) 'tsid
                query_temp.Add(Function_Form.getNull(0)) 'spcode
                query_temp.Add(Function_Form.getNull(1)) 'taxdate
                query_temp.Add(Function_Form.getNull(1)) 'taxdate_bt
                query_temp.Add(value_arraylist(0)(row)(112)) 'createdby
                query_temp.Add(value_arraylist(0)(row)(112)) 'updatedby
                query_temp.Add(Function_Form.getNull(2)) 'createdate
                query_temp.Add(Function_Form.getNull(2)) 'lastupdate
                Dim cmd_last = ""
                For j = 0 To query_temp.Count - 1
                    cmd_last += "'" + query_temp(j) + "',"
                Next
                cmd_last = cmd_last.Substring(0, cmd_last.Length - 1) + ")"

                Dim seq = 1
                Dim queryAL As New ArrayList
                Dim amount As Double = 0
                Dim debit As Double = 0
                Dim credit As Double = 0
                Dim fx_rate As Double = 0
                Dim fx_amount As Double = 0
                Dim fx_debit As Double = 0
                Dim fx_credit As Double = 0
                For i = startRange To endRange
                    queryAL.Clear()
                    'Product
                    queryAL.Add(value_arraylist(1)(i)(57)) 'accno
                    queryAL.Add("SI") 'doc_type
                    queryAL.Add(value_arraylist(0)(row)(1)) 'doc_no
                    queryAL.Add(seq) 'seq
                    queryAL.Add(Function_Form.convertDateFormat(value_arraylist(0)(row)(2))) 'doc_date
                    queryAL.Add(value_arraylist(0)(row)(1)) 'refno
                    queryAL.Add(value_arraylist(0)(row)(9)) 'refno2
                    queryAL.Add(Function_Form.getNull(0)) 'refno3
                    queryAL.Add(value_arraylist(0)(row)(11)) 'desp
                    queryAL.Add(value_arraylist(0)(row)(4)) 'desp2
                    queryAL.Add(Function_Form.getNull(0)) 'desp3
                    queryAL.Add(Function_Form.getNull(0)) 'desp4

                    amount = Math.Round(CDbl(value_arraylist(1)(i)(31)) * -1, 2)
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
                    fx_rate = CDbl(value_arraylist(0)(row)(18))
                    fx_amount = Math.Round(CDbl(value_arraylist(1)(i)(26)) * -1, 2)
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
                    queryAL.Add(fx_rate) 'fx_rate
                    queryAL.Add(value_arraylist(0)(row)(17)) 'curr_code
                    queryAL.Add(value_arraylist(0)(row)(48)) 'batchno
                    queryAL.Add(value_arraylist(0)(row)(49)) 'projcode
                    queryAL.Add(value_arraylist(0)(row)(50)) 'deptcode
                    queryAL.Add(Function_Form.getNull(0)) 'taxcode
                    queryAL.Add(Function_Form.getNull(3)) 'taxable
                    queryAL.Add(Function_Form.getNull(3)) 'fx_taxable
                    queryAL.Add(Function_Form.getNull(3)) 'link_seq

                    Dim execmd As String = queryTable(4)(2)
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
                    seq += 1

                    'If product has tax
                    Dim taxcode = value_arraylist(1)(i)(54).ToString.Trim
                    If Not taxcode.Equals(String.Empty) Then
                        queryAL.Clear()
                        Dim accno = String.Empty
                        myConn.Open()
                        Dim gltaxcommand = New SqlCommand("SELECT accno FROM gltax WHERE taxcode ='" + taxcode + "'", myConn)
                        Dim gltaxreader As SqlDataReader = gltaxcommand.ExecuteReader
                        While gltaxreader.Read()
                            accno = gltaxreader.GetValue(0)
                        End While
                        myConn.Close()
                        If accno.Equals(String.Empty) Then
                            MsgBox("Tax Code '" + taxcode + "' is not found in database!" + vbNewLine + " The operation has been stopped!", MsgBoxStyle.Critical)
                            Return
                        Else
                            queryAL.Add(accno) 'accno
                        End If
                        queryAL.Add("SI") 'doc_type
                        queryAL.Add(value_arraylist(0)(row)(1)) 'doc_no
                        queryAL.Add(seq) 'seq
                        queryAL.Add(Function_Form.convertDateFormat(value_arraylist(0)(row)(2))) 'doc_date
                        queryAL.Add(value_arraylist(0)(row)(1)) 'refno
                        queryAL.Add(value_arraylist(0)(row)(9)) 'refno2
                        queryAL.Add(Function_Form.getNull(0)) 'refno3
                        queryAL.Add(value_arraylist(0)(row)(11)) 'desp
                        queryAL.Add(value_arraylist(0)(row)(4)) 'desp2
                        queryAL.Add(Function_Form.getNull(0)) 'desp3
                        queryAL.Add(Function_Form.getNull(0)) 'desp4

                        amount = Math.Round(CDbl(value_arraylist(1)(i)(35)) * -1, 2)
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
                        fx_rate = CDbl(value_arraylist(0)(row)(18))
                        fx_amount = Math.Round(CDbl(value_arraylist(1)(i)(23)) * -1, 2)
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
                        queryAL.Add(fx_rate) 'fx_rate

                        queryAL.Add(value_arraylist(0)(row)(17)) 'curr_code
                        queryAL.Add(value_arraylist(0)(row)(48)) 'batchno
                        queryAL.Add(value_arraylist(0)(row)(49)) 'projcode
                        queryAL.Add(value_arraylist(0)(row)(50)) 'deptcode
                        queryAL.Add(taxcode) 'taxcode
                        Dim taxable = Math.Round(CDbl(value_arraylist(1)(i)(26)) * -1, 2)
                        queryAL.Add(taxable * fx_rate) 'taxable
                        queryAL.Add(taxable) 'fx_taxable
                        queryAL.Add((seq - 1).ToString) 'link_seq

                        Dim execmd2 = queryTable(4)(2)
                        For j = 0 To queryAL.Count - 1
                            execmd2 += "'" + queryAL(j).ToString + "',"
                        Next
                        execmd2 = execmd2 + cmd_last
                        myConn.Open()
                        Dim cmd2 = New SqlCommand(execmd2, myConn)
                        cmd2.ExecuteNonQuery()
                        rowInsertNum += 1
                        myConn.Close()
                        seq += 1
                    End If
                Next

                queryAL.Clear()
                'Subtotal of Product
                queryAL.Add(value_arraylist(0)(row)(10)) 'accno
                queryAL.Add("SI") 'doc_type
                queryAL.Add(value_arraylist(0)(row)(1)) 'doc_no
                queryAL.Add(seq) 'seq
                arseq.Add(value_arraylist(0)(row)(1) + "." + seq.ToString + ".0." + row.ToString) 'AR get doc_no,seq,knockoff,row
                gloffseq.Add(value_arraylist(0)(row)(1) + "." + seq.ToString + "." + (seq + 1).ToString) 'GLOff get doc_no,seq
                queryAL.Add(Function_Form.convertDateFormat(value_arraylist(0)(row)(2))) 'doc_date
                queryAL.Add(value_arraylist(0)(row)(1)) 'refno
                queryAL.Add(value_arraylist(0)(row)(9)) 'refno2
                queryAL.Add(Function_Form.getNull(0)) 'refno3
                queryAL.Add(value_arraylist(0)(row)(3)) 'desp
                queryAL.Add(value_arraylist(0)(row)(4)) 'desp2
                queryAL.Add(Function_Form.getNull(0)) 'desp3
                queryAL.Add(Function_Form.getNull(0)) 'desp4

                amount = Math.Round(CDbl(value_arraylist(0)(row)(43)), 2)
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
                fx_rate = CDbl(value_arraylist(0)(row)(18))
                fx_amount = Math.Round(CDbl(value_arraylist(0)(row)(35)), 2)
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
                queryAL.Add(fx_rate) 'fx_rate

                queryAL.Add(value_arraylist(0)(row)(17)) 'curr_code
                queryAL.Add(value_arraylist(0)(row)(48)) 'batchno
                queryAL.Add(value_arraylist(0)(row)(49)) 'projcode
                queryAL.Add(value_arraylist(0)(row)(50)) 'deptcode
                queryAL.Add(Function_Form.getNull(0)) 'taxcode
                queryAL.Add(Function_Form.getNull(3)) 'taxable
                queryAL.Add(Function_Form.getNull(3)) 'fx_taxable
                queryAL.Add(Function_Form.getNull(3)) 'link_seq

                Dim subtotalcmd As String = queryTable(4)(2)
                For j = 0 To queryAL.Count - 1
                    subtotalcmd += "'" + queryAL(j).ToString + "',"
                Next
                subtotalcmd = subtotalcmd + cmd_last
                myConn.Open()
                Dim cmd3 = New SqlCommand(subtotalcmd, myConn)
                cmd3.ExecuteNonQuery()
                rowInsertNum += 1
                myConn.Close()
                seq += 1

                Dim acc_p1 = value_arraylist(0)(row)(73)
                Dim ref_p1 = value_arraylist(0)(row)(74)
                Dim amt_p1 = value_arraylist(0)(row)(75)
                Dim acc_p2 = value_arraylist(0)(row)(78)
                Dim ref_p2 = value_arraylist(0)(row)(79)
                Dim amt_p2 = value_arraylist(0)(row)(80)
                Dim acc_p3 = value_arraylist(0)(row)(83)
                Dim ref_p3 = value_arraylist(0)(row)(84)
                Dim amt_p3 = value_arraylist(0)(row)(85)
                Dim acc_p4 = value_arraylist(0)(row)(86)
                Dim ref_p4 = value_arraylist(0)(row)(87)
                Dim amt_p4 = value_arraylist(0)(row)(88)
                Dim p_def = -1
                Dim p_def_accno = ""
                Dim p_def_refno = ""
                If Not acc_p1.ToString.Trim.Equals(String.Empty) And Not amt_p1.ToString.Trim.Equals("0") Then
                    p_def = 1
                    p_def_accno = acc_p1
                    p_def_refno = ref_p1
                Else
                    If Not acc_p2.ToString.Trim.Equals(String.Empty) And Not amt_p2.ToString.Trim.Equals("0") Then
                        p_def = 2
                        p_def_accno = acc_p2
                        p_def_refno = ref_p2
                    Else
                        If Not acc_p3.ToString.Trim.Equals(String.Empty) And Not amt_p3.ToString.Trim.Equals("0") Then
                            p_def = 3
                            p_def_accno = acc_p3
                            p_def_refno = ref_p3
                        Else
                            If Not acc_p4.ToString.Trim.Equals(String.Empty) And Not amt_p4.ToString.Trim.Equals("0") Then
                                p_def = 4
                                p_def_accno = acc_p4
                                p_def_refno = ref_p4
                            End If
                        End If
                    End If
                End If
                If p_def = -1 Then
                    Continue For
                End If
                myConn.Open()
                Dim gldatacommand = New SqlCommand("SELECT accdesp,accdesp2 FROM gldata WHERE accno ='" + p_def_accno + "'", myConn)
                Dim gldatareader As SqlDataReader = gldatacommand.ExecuteReader
                Dim p_def_desp = ""
                Dim p_def_desp2 = ""
                While gldatareader.Read()
                    p_def_desp = gldatareader.GetValue(0)
                    p_def_desp2 = gldatareader.GetValue(1)
                End While
                myConn.Close()

                queryAL.Clear()
                'Subtotal Of Default Payment
                queryAL.Add(value_arraylist(0)(row)(10)) 'accno
                queryAL.Add("SI") 'doc_type
                queryAL.Add(value_arraylist(0)(row)(1)) 'doc_no
                queryAL.Add(seq) 'seq
                arseq.Add(value_arraylist(0)(row)(1) + "." + seq.ToString + ".1." + row.ToString) 'AR get doc_no,seq,knockoff,row
                queryAL.Add(Function_Form.convertDateFormat(value_arraylist(0)(row)(2))) 'doc_date
                queryAL.Add(value_arraylist(0)(row)(1)) 'refno
                queryAL.Add(value_arraylist(0)(row)(9)) 'refno2
                queryAL.Add(Function_Form.getNull(0)) 'refno3
                queryAL.Add(p_def_desp) 'desp
                queryAL.Add(p_def_desp2) 'desp2
                queryAL.Add(Function_Form.getNull(0)) 'desp3
                queryAL.Add(Function_Form.getNull(0)) 'desp4
                If Not acc_p1.ToString.Trim.Equals(String.Empty) Then
                    amt_p1 = Math.Round(CDbl(amt_p1), 2)
                Else
                    amt_p1 = 0
                End If
                If Not acc_p2.ToString.Trim.Equals(String.Empty) Then
                    amt_p2 = Math.Round(CDbl(amt_p2), 2)
                Else
                    amt_p2 = 0
                End If
                If Not acc_p3.ToString.Trim.Equals(String.Empty) Then
                    amt_p3 = Math.Round(CDbl(amt_p3), 2)
                Else
                    amt_p3 = 0
                End If
                If Not acc_p4.ToString.Trim.Equals(String.Empty) Then
                    amt_p4 = Math.Round(CDbl(amt_p4), 2)
                Else
                    amt_p4 = 0
                End If

                fx_rate = CDbl(value_arraylist(0)(row)(18))
                amount = ((CDbl(amt_p1) + CDbl(amt_p2) + CDbl(amt_p3) + CDbl(amt_p4)) * -1) * fx_rate
                debit = 0
                credit = 0
                If amount < 0 Then
                    credit = amount * -1
                Else
                    debit = amount
                End If
                fx_amount = (CDbl(amt_p1) + CDbl(amt_p2) + CDbl(amt_p3) + CDbl(amt_p4)) * -1
                fx_debit = 0
                fx_credit = 0
                If fx_amount < 0 Then
                    fx_credit = fx_amount * -1
                Else
                    fx_debit = fx_amount
                End If
                queryAL.Add(amount) 'amount
                queryAL.Add(debit) 'debit
                queryAL.Add(credit) 'credit
                queryAL.Add(fx_amount) 'fx_amount
                queryAL.Add(fx_debit) 'fx_debit
                queryAL.Add(fx_credit) 'fx_credit
                queryAL.Add(fx_rate) 'fx_rate

                queryAL.Add(value_arraylist(0)(row)(17)) 'curr_code
                queryAL.Add(value_arraylist(0)(row)(48)) 'batchno
                queryAL.Add(value_arraylist(0)(row)(49)) 'projcode
                queryAL.Add(value_arraylist(0)(row)(50)) 'deptcode
                queryAL.Add(Function_Form.getNull(0)) 'taxcode
                queryAL.Add(Function_Form.getNull(3)) 'taxable
                queryAL.Add(Function_Form.getNull(3)) 'fx_taxable
                queryAL.Add(Function_Form.getNull(3)) 'link_seq
                queryAL.Add("I") 'billtype
                queryAL.Add(Function_Form.getNull(0)) 'remark1
                queryAL.Add(Function_Form.getNull(0)) 'remark2
                queryAL.Add(p_def_refno) 'cheque_no
                queryAL.Add(Function_Form.convertDateFormat(value_arraylist(0)(row)(2))) 'chqrc_date
                queryAL.Add(Function_Form.getNull(1)) 'koff_date
                queryAL.Add(Function_Form.getNull(1)) 'recon_date
                queryAL.Add(Function_Form.getNull(3)) 'recon_flag
                queryAL.Add(value_arraylist(0)(row)(16)) 'accmgr_id
                queryAL.Add(Function_Form.getNull(0)) 'lkdoc_type
                queryAL.Add(Function_Form.getNull(0)) 'lkdoc_no
                queryAL.Add(Function_Form.getNull(3)) 'lkseq
                queryAL.Add("1") 'lock
                queryAL.Add(Function_Form.getNull(3)) 'void
                queryAL.Add(Function_Form.getNull(3)) 'exported
                queryAL.Add(Function_Form.getNull(0)) 'entry
                queryAL.Add(Function_Form.getNull(3)) 'fastentry
                queryAL.Add(Function_Form.getNull(3)) 'followdesp
                queryAL.Add(Function_Form.getNull(3)) 'tsid
                queryAL.Add(Function_Form.getNull(0)) 'spcode
                queryAL.Add(Function_Form.getNull(1)) 'taxdate
                queryAL.Add(Function_Form.getNull(1)) 'taxdate_bt
                queryAL.Add(value_arraylist(0)(row)(112)) 'createdby
                queryAL.Add(value_arraylist(0)(row)(112)) 'updatedby
                queryAL.Add(Function_Form.getNull(2)) 'createdate
                queryAL.Add(Function_Form.getNull(2)) 'lastupdate

                Dim subtotalpaycmd As String = queryTable(4)(2)
                For j = 0 To queryAL.Count - 1
                    subtotalpaycmd += "'" + queryAL(j).ToString + "',"
                Next
                subtotalpaycmd = subtotalpaycmd.Substring(0, subtotalpaycmd.Length - 1) + ")"
                myConn.Open()
                Dim cmd4 = New SqlCommand(subtotalpaycmd, myConn)
                cmd4.ExecuteNonQuery()
                rowInsertNum += 1
                myConn.Close()
                seq += 1

                'Subtotal Of Payment1
                If Not acc_p1.ToString.Trim.Equals(String.Empty) Then
                    queryAL.Clear()
                    queryAL.Add(acc_p1) 'accno
                    queryAL.Add("SI") 'doc_type
                    queryAL.Add(value_arraylist(0)(row)(1)) 'doc_no
                    queryAL.Add(seq) 'seq
                    queryAL.Add(Function_Form.convertDateFormat(value_arraylist(0)(row)(2))) 'doc_date
                    queryAL.Add(value_arraylist(0)(row)(1)) 'refno
                    queryAL.Add(value_arraylist(0)(row)(9)) 'refno2
                    queryAL.Add(Function_Form.getNull(0)) 'refno3
                    queryAL.Add(value_arraylist(0)(row)(11)) 'desp
                    queryAL.Add(Function_Form.getNull(0)) 'desp2
                    queryAL.Add(Function_Form.getNull(0)) 'desp3
                    queryAL.Add(Function_Form.getNull(0)) 'desp4

                    fx_rate = CDbl(value_arraylist(0)(row)(18))
                    fx_amount = Math.Round(amt_p1, 2)
                    fx_debit = 0
                    fx_credit = 0
                    If fx_amount < 0 Then
                        fx_credit = fx_amount * -1
                    Else
                        fx_debit = fx_amount
                    End If
                    amt_p1 = Math.Round(amt_p1 * fx_rate, 2)
                    debit = 0
                    credit = 0
                    If amt_p1 < 0 Then
                        credit = amt_p1 * -1
                    Else
                        debit = amt_p1
                    End If
                    queryAL.Add(amt_p1) 'amount
                    queryAL.Add(debit) 'debit
                    queryAL.Add(credit) 'credit
                    queryAL.Add(fx_amount) 'fx_amount
                    queryAL.Add(fx_debit) 'fx_debit
                    queryAL.Add(fx_credit) 'fx_credit
                    queryAL.Add(fx_rate) 'fx_rate

                    queryAL.Add(value_arraylist(0)(row)(17)) 'curr_code
                    queryAL.Add(value_arraylist(0)(row)(48)) 'batchno
                    queryAL.Add(value_arraylist(0)(row)(49)) 'projcode
                    queryAL.Add(value_arraylist(0)(row)(50)) 'deptcode
                    queryAL.Add(Function_Form.getNull(0)) 'taxcode
                    queryAL.Add(Function_Form.getNull(3)) 'taxable
                    queryAL.Add(Function_Form.getNull(3)) 'fx_taxable
                    queryAL.Add(Function_Form.getNull(3)) 'link_seq
                    queryAL.Add("I") 'billtype
                    queryAL.Add(Function_Form.getNull(0)) 'remark1
                    queryAL.Add(Function_Form.getNull(0)) 'remark2
                    queryAL.Add(ref_p1) 'cheque_no
                    queryAL.Add(Function_Form.convertDateFormat(value_arraylist(0)(row)(2))) 'chqrc_date
                    queryAL.Add(Function_Form.getNull(1)) 'koff_date
                    queryAL.Add(Function_Form.getNull(1)) 'recon_date
                    queryAL.Add(Function_Form.getNull(3)) 'recon_flag
                    queryAL.Add(value_arraylist(0)(row)(16)) 'accmgr_id
                    queryAL.Add(Function_Form.getNull(0)) 'lkdoc_type
                    queryAL.Add(Function_Form.getNull(0)) 'lkdoc_no
                    queryAL.Add(Function_Form.getNull(3)) 'lkseq
                    queryAL.Add("1") 'lock
                    queryAL.Add(Function_Form.getNull(3)) 'void
                    queryAL.Add(Function_Form.getNull(3)) 'exported
                    queryAL.Add(Function_Form.getNull(0)) 'entry
                    queryAL.Add(Function_Form.getNull(3)) 'fastentry
                    queryAL.Add(Function_Form.getNull(3)) 'followdesp
                    queryAL.Add(Function_Form.getNull(3)) 'tsid
                    queryAL.Add(Function_Form.getNull(0)) 'spcode
                    queryAL.Add(Function_Form.getNull(1)) 'taxdate
                    queryAL.Add(Function_Form.getNull(1)) 'taxdate_bt
                    queryAL.Add(value_arraylist(0)(row)(112)) 'createdby
                    queryAL.Add(value_arraylist(0)(row)(112)) 'updatedby
                    queryAL.Add(Function_Form.getNull(2)) 'createdate
                    queryAL.Add(Function_Form.getNull(2)) 'lastupdate

                    Dim pay1cmd As String = queryTable(4)(2)
                    For j = 0 To queryAL.Count - 1
                        pay1cmd += "'" + queryAL(j).ToString + "',"
                    Next
                    pay1cmd = pay1cmd.Substring(0, pay1cmd.Length - 1) + ")"
                    myConn.Open()
                    Dim cmd_p1 = New SqlCommand(pay1cmd, myConn)
                    cmd_p1.ExecuteNonQuery()
                    rowInsertNum += 1
                    myConn.Close()
                    seq += 1
                End If

                'Subtotal Of Payment2
                If Not acc_p2.ToString.Trim.Equals(String.Empty) Then
                    queryAL.Clear()
                    queryAL.Add(acc_p2) 'accno
                    queryAL.Add("SI") 'doc_type
                    queryAL.Add(value_arraylist(0)(row)(1)) 'doc_no
                    queryAL.Add(seq) 'seq
                    queryAL.Add(Function_Form.convertDateFormat(value_arraylist(0)(row)(2))) 'doc_date
                    queryAL.Add(value_arraylist(0)(row)(1)) 'refno
                    queryAL.Add(value_arraylist(0)(row)(9)) 'refno2
                    queryAL.Add(Function_Form.getNull(0)) 'refno3
                    queryAL.Add(value_arraylist(0)(row)(11)) 'desp
                    queryAL.Add(Function_Form.getNull(0)) 'desp2
                    queryAL.Add(Function_Form.getNull(0)) 'desp3
                    queryAL.Add(Function_Form.getNull(0)) 'desp4

                    fx_rate = CDbl(value_arraylist(0)(row)(18))
                    fx_amount = Math.Round(amt_p2, 2)
                    fx_debit = 0
                    fx_credit = 0
                    If fx_amount < 0 Then
                        fx_credit = fx_amount * -1
                    Else
                        fx_debit = fx_amount
                    End If
                    amt_p2 = Math.Round(amt_p2 * fx_rate, 2)
                    debit = 0
                    credit = 0
                    If amt_p2 < 0 Then
                        credit = amt_p2 * -1
                    Else
                        debit = amt_p2
                    End If
                    queryAL.Add(amt_p2) 'amount
                    queryAL.Add(debit) 'debit
                    queryAL.Add(credit) 'credit
                    queryAL.Add(fx_amount) 'fx_amount
                    queryAL.Add(fx_debit) 'fx_debit
                    queryAL.Add(fx_credit) 'fx_credit
                    queryAL.Add(fx_rate) 'fx_rate

                    queryAL.Add(value_arraylist(0)(row)(17)) 'curr_code
                    queryAL.Add(value_arraylist(0)(row)(48)) 'batchno
                    queryAL.Add(value_arraylist(0)(row)(49)) 'projcode
                    queryAL.Add(value_arraylist(0)(row)(50)) 'deptcode
                    queryAL.Add(Function_Form.getNull(0)) 'taxcode
                    queryAL.Add(Function_Form.getNull(3)) 'taxable
                    queryAL.Add(Function_Form.getNull(3)) 'fx_taxable
                    queryAL.Add(Function_Form.getNull(3)) 'link_seq
                    queryAL.Add("I") 'billtype
                    queryAL.Add(Function_Form.getNull(0)) 'remark1
                    queryAL.Add(Function_Form.getNull(0)) 'remark2
                    queryAL.Add(ref_p2) 'cheque_no
                    queryAL.Add(Function_Form.convertDateFormat(value_arraylist(0)(row)(2))) 'chqrc_date
                    queryAL.Add(Function_Form.getNull(1)) 'koff_date
                    queryAL.Add(Function_Form.getNull(1)) 'recon_date
                    queryAL.Add(Function_Form.getNull(3)) 'recon_flag
                    queryAL.Add(value_arraylist(0)(row)(16)) 'accmgr_id
                    queryAL.Add(Function_Form.getNull(0)) 'lkdoc_type
                    queryAL.Add(Function_Form.getNull(0)) 'lkdoc_no
                    queryAL.Add(Function_Form.getNull(3)) 'lkseq
                    queryAL.Add("1") 'lock
                    queryAL.Add(Function_Form.getNull(3)) 'void
                    queryAL.Add(Function_Form.getNull(3)) 'exported
                    queryAL.Add(Function_Form.getNull(0)) 'entry
                    queryAL.Add(Function_Form.getNull(3)) 'fastentry
                    queryAL.Add(Function_Form.getNull(3)) 'followdesp
                    queryAL.Add(Function_Form.getNull(3)) 'tsid
                    queryAL.Add(Function_Form.getNull(0)) 'spcode
                    queryAL.Add(Function_Form.getNull(1)) 'taxdate
                    queryAL.Add(Function_Form.getNull(1)) 'taxdate_bt
                    queryAL.Add(value_arraylist(0)(row)(112)) 'createdby
                    queryAL.Add(value_arraylist(0)(row)(112)) 'updatedby
                    queryAL.Add(Function_Form.getNull(2)) 'createdate
                    queryAL.Add(Function_Form.getNull(2)) 'lastupdate

                    Dim paycmd As String = queryTable(4)(2)
                    For j = 0 To queryAL.Count - 1
                        paycmd += "'" + queryAL(j).ToString + "',"
                    Next
                    paycmd = paycmd.Substring(0, paycmd.Length - 1) + ")"
                    myConn.Open()
                    Dim cmd_p = New SqlCommand(paycmd, myConn)
                    cmd_p.ExecuteNonQuery()
                    rowInsertNum += 1
                    myConn.Close()
                    seq += 1
                End If

                'Subtotal Of Payment3
                If Not acc_p3.ToString.Trim.Equals(String.Empty) Then
                    queryAL.Clear()
                    queryAL.Add(acc_p3) 'accno
                    queryAL.Add("SI") 'doc_type
                    queryAL.Add(value_arraylist(0)(row)(1)) 'doc_no
                    queryAL.Add(seq) 'seq
                    queryAL.Add(Function_Form.convertDateFormat(value_arraylist(0)(row)(2))) 'doc_date
                    queryAL.Add(value_arraylist(0)(row)(1)) 'refno
                    queryAL.Add(value_arraylist(0)(row)(9)) 'refno2
                    queryAL.Add(Function_Form.getNull(0)) 'refno3
                    queryAL.Add(value_arraylist(0)(row)(11)) 'desp
                    queryAL.Add(Function_Form.getNull(0)) 'desp2
                    queryAL.Add(Function_Form.getNull(0)) 'desp3
                    queryAL.Add(Function_Form.getNull(0)) 'desp4

                    fx_rate = CDbl(value_arraylist(0)(row)(18))
                    fx_amount = Math.Round(amt_p3, 2)
                    fx_debit = 0
                    fx_credit = 0
                    If fx_amount < 0 Then
                        fx_credit = fx_amount * -1
                    Else
                        fx_debit = fx_amount
                    End If
                    amt_p3 = Math.Round(amt_p3 * fx_rate, 2)
                    debit = 0
                    credit = 0
                    If amt_p3 < 0 Then
                        credit = amt_p3 * -1
                    Else
                        debit = amt_p3
                    End If
                    queryAL.Add(amt_p3) 'amount
                    queryAL.Add(debit) 'debit
                    queryAL.Add(credit) 'credit
                    queryAL.Add(fx_amount) 'fx_amount
                    queryAL.Add(fx_debit) 'fx_debit
                    queryAL.Add(fx_credit) 'fx_credit
                    queryAL.Add(fx_rate) 'fx_rate

                    queryAL.Add(value_arraylist(0)(row)(17)) 'curr_code
                    queryAL.Add(value_arraylist(0)(row)(48)) 'batchno
                    queryAL.Add(value_arraylist(0)(row)(49)) 'projcode
                    queryAL.Add(value_arraylist(0)(row)(50)) 'deptcode
                    queryAL.Add(Function_Form.getNull(0)) 'taxcode
                    queryAL.Add(Function_Form.getNull(3)) 'taxable
                    queryAL.Add(Function_Form.getNull(3)) 'fx_taxable
                    queryAL.Add(Function_Form.getNull(3)) 'link_seq
                    queryAL.Add("I") 'billtype
                    queryAL.Add(Function_Form.getNull(0)) 'remark1
                    queryAL.Add(Function_Form.getNull(0)) 'remark2
                    queryAL.Add(ref_p3) 'cheque_no
                    queryAL.Add(Function_Form.convertDateFormat(value_arraylist(0)(row)(2))) 'chqrc_date
                    queryAL.Add(Function_Form.getNull(1)) 'koff_date
                    queryAL.Add(Function_Form.getNull(1)) 'recon_date
                    queryAL.Add(Function_Form.getNull(3)) 'recon_flag
                    queryAL.Add(value_arraylist(0)(row)(16)) 'accmgr_id
                    queryAL.Add(Function_Form.getNull(0)) 'lkdoc_type
                    queryAL.Add(Function_Form.getNull(0)) 'lkdoc_no
                    queryAL.Add(Function_Form.getNull(3)) 'lkseq
                    queryAL.Add("1") 'lock
                    queryAL.Add(Function_Form.getNull(3)) 'void
                    queryAL.Add(Function_Form.getNull(3)) 'exported
                    queryAL.Add(Function_Form.getNull(0)) 'entry
                    queryAL.Add(Function_Form.getNull(3)) 'fastentry
                    queryAL.Add(Function_Form.getNull(3)) 'followdesp
                    queryAL.Add(Function_Form.getNull(3)) 'tsid
                    queryAL.Add(Function_Form.getNull(0)) 'spcode
                    queryAL.Add(Function_Form.getNull(1)) 'taxdate
                    queryAL.Add(Function_Form.getNull(1)) 'taxdate_bt
                    queryAL.Add(value_arraylist(0)(row)(112)) 'createdby
                    queryAL.Add(value_arraylist(0)(row)(112)) 'updatedby
                    queryAL.Add(Function_Form.getNull(2)) 'createdate
                    queryAL.Add(Function_Form.getNull(2)) 'lastupdate

                    Dim paycmd As String = queryTable(4)(2)
                    For j = 0 To queryAL.Count - 1
                        paycmd += "'" + queryAL(j).ToString + "',"
                    Next
                    paycmd = paycmd.Substring(0, paycmd.Length - 1) + ")"
                    myConn.Open()
                    Dim cmd_p = New SqlCommand(paycmd, myConn)
                    cmd_p.ExecuteNonQuery()
                    rowInsertNum += 1
                    myConn.Close()
                    seq += 1
                End If

                'Subtotal Of Payment4
                If Not acc_p4.ToString.Trim.Equals(String.Empty) Then
                    queryAL.Clear()
                    queryAL.Add(acc_p4) 'accno
                    queryAL.Add("SI") 'doc_type
                    queryAL.Add(value_arraylist(0)(row)(1)) 'doc_no
                    queryAL.Add(seq) 'seq
                    queryAL.Add(Function_Form.convertDateFormat(value_arraylist(0)(row)(2))) 'doc_date
                    queryAL.Add(value_arraylist(0)(row)(1)) 'refno
                    queryAL.Add(value_arraylist(0)(row)(9)) 'refno2
                    queryAL.Add(Function_Form.getNull(0)) 'refno3
                    queryAL.Add(value_arraylist(0)(row)(11)) 'desp
                    queryAL.Add(Function_Form.getNull(0)) 'desp2
                    queryAL.Add(Function_Form.getNull(0)) 'desp3
                    queryAL.Add(Function_Form.getNull(0)) 'desp4

                    fx_rate = CDbl(value_arraylist(0)(row)(18))
                    fx_amount = Math.Round(amt_p4, 2)
                    fx_debit = 0
                    fx_credit = 0
                    If fx_amount < 0 Then
                        fx_credit = fx_amount * -1
                    Else
                        fx_debit = fx_amount
                    End If
                    amt_p4 = Math.Round(amt_p4 * fx_rate, 2)
                    debit = 0
                    credit = 0
                    If amt_p4 < 0 Then
                        credit = amt_p4 * -1
                    Else
                        debit = amt_p4
                    End If
                    queryAL.Add(amt_p4) 'amount
                    queryAL.Add(debit) 'debit
                    queryAL.Add(credit) 'credit
                    queryAL.Add(fx_amount) 'fx_amount
                    queryAL.Add(fx_debit) 'fx_debit
                    queryAL.Add(fx_credit) 'fx_credit
                    queryAL.Add(fx_rate) 'fx_rate

                    queryAL.Add(value_arraylist(0)(row)(17)) 'curr_code
                    queryAL.Add(value_arraylist(0)(row)(48)) 'batchno
                    queryAL.Add(value_arraylist(0)(row)(49)) 'projcode
                    queryAL.Add(value_arraylist(0)(row)(50)) 'deptcode
                    queryAL.Add(Function_Form.getNull(0)) 'taxcode
                    queryAL.Add(Function_Form.getNull(3)) 'taxable
                    queryAL.Add(Function_Form.getNull(3)) 'fx_taxable
                    queryAL.Add(Function_Form.getNull(3)) 'link_seq
                    queryAL.Add("I") 'billtype
                    queryAL.Add(Function_Form.getNull(0)) 'remark1
                    queryAL.Add(Function_Form.getNull(0)) 'remark2
                    queryAL.Add(ref_p4) 'cheque_no
                    queryAL.Add(Function_Form.convertDateFormat(value_arraylist(0)(row)(2))) 'chqrc_date
                    queryAL.Add(Function_Form.getNull(1)) 'koff_date
                    queryAL.Add(Function_Form.getNull(1)) 'recon_date
                    queryAL.Add(Function_Form.getNull(3)) 'recon_flag
                    queryAL.Add(value_arraylist(0)(row)(16)) 'accmgr_id
                    queryAL.Add(Function_Form.getNull(0)) 'lkdoc_type
                    queryAL.Add(Function_Form.getNull(0)) 'lkdoc_no
                    queryAL.Add(Function_Form.getNull(3)) 'lkseq
                    queryAL.Add("1") 'lock
                    queryAL.Add(Function_Form.getNull(3)) 'void
                    queryAL.Add(Function_Form.getNull(3)) 'exported
                    queryAL.Add(Function_Form.getNull(0)) 'entry
                    queryAL.Add(Function_Form.getNull(3)) 'fastentry
                    queryAL.Add(Function_Form.getNull(3)) 'followdesp
                    queryAL.Add(Function_Form.getNull(3)) 'tsid
                    queryAL.Add(Function_Form.getNull(0)) 'spcode
                    queryAL.Add(Function_Form.getNull(1)) 'taxdate
                    queryAL.Add(Function_Form.getNull(1)) 'taxdate_bt
                    queryAL.Add(value_arraylist(0)(row)(112)) 'createdby
                    queryAL.Add(value_arraylist(0)(row)(112)) 'updatedby
                    queryAL.Add(Function_Form.getNull(2)) 'createdate
                    queryAL.Add(Function_Form.getNull(2)) 'lastupdate

                    Dim paycmd As String = queryTable(4)(2)
                    For j = 0 To queryAL.Count - 1
                        paycmd += "'" + queryAL(j).ToString + "',"
                    Next
                    paycmd = paycmd.Substring(0, paycmd.Length - 1) + ")"
                    myConn.Open()
                    Dim cmd_p = New SqlCommand(paycmd, myConn)
                    cmd_p.ExecuteNonQuery()
                    rowInsertNum += 1
                    myConn.Close()
                    seq += 1
                End If
            End If
        Next

        'ar
        For Each ar As String In arseq
            Dim row As String = ar.Split(".")(3)
            Dim custcode As String = ""
            Dim doc_type As String = "SI"
            Dim doc_no As String = ar.Split(".")(0)
            Dim seq As String = ar.Split(".")(1)
            Dim knockoff As String = ar.Split(".")(2)
            Dim doc_date As String = Convert.ToDateTime(dgvExcel.Rows(row).Cells("Date").Value.ToString).ToString("yyyy-MM-dd HH:mm:ss")

            'due_date
            Dim credit_terms As Integer = CInt(value_arraylist(0)(row)(15))
            Dim due_date As String = Convert.ToDateTime(doc_date).AddDays(credit_terms).ToString("yyyy-MM-dd HH:mm:ss") '2

            Dim refno As String = ""
            Dim refno2 As String = ""
            Dim refno3 As String = ""
            Dim desp As String = ""
            Dim desp2 As String = ""
            Dim desp3 As String = ""
            Dim desp4 As String = ""
            Dim remark1 As String = ""
            Dim remark2 As String = ""
            Dim cheque_no As String = ""
            Dim chqrc_date As String = ""
            Dim koff_date As String = ""
            Dim curr_code As String = ""
            Dim fx_rate As String = value_arraylist(0)(row)(18)
            Dim fx_gainloss As String = Function_Form.getNull(3)
            Dim amount As String = ""

            'paid
            Dim sum_p As Double = 0
            Dim amt_p1 = value_arraylist(0)(row)(75)
            Dim amt_p2 = value_arraylist(0)(row)(80)
            Dim amt_p3 = value_arraylist(0)(row)(85)
            Dim amt_p4 = value_arraylist(0)(row)(88)
            If Not value_arraylist(0)(row)(73).Equals(String.Empty) Then
                sum_p += CDbl(amt_p1)
            End If
            If Not value_arraylist(0)(row)(78).Equals(String.Empty) Then
                sum_p += CDbl(amt_p2)
            End If
            If Not value_arraylist(0)(row)(83).Equals(String.Empty) Then
                sum_p += CDbl(amt_p3)
            End If
            If Not value_arraylist(0)(row)(86).Equals(String.Empty) Then
                sum_p += CDbl(amt_p4)
            End If
            Dim paid As String
            If knockoff.Equals("0") Then
                paid = sum_p.ToString
            Else
                paid = (sum_p * -1).ToString
            End If

            Dim local_amount As String = ""
            Dim local_paid As String
            If knockoff.Equals("0") Then
                local_paid = (sum_p * CDbl(fx_rate)).ToString '2
            Else
                local_paid = (sum_p * CDbl(fx_rate) * -1).ToString '2
            End If
            Dim taxable As String
            Dim tax As String
            Dim fx_taxable As String
            Dim fx_tax As String
            If knockoff.Equals("0") Then
                taxable = (CDbl(value_arraylist(0)(row)(27)) * CDbl(fx_rate)).ToString '2
                tax = (CDbl(value_arraylist(0)(row)(31)) * CDbl(fx_rate)).ToString '2
                fx_taxable = value_arraylist(0)(row)(27) '2
                fx_tax = value_arraylist(0)(row)(31) '2
            Else
                taxable = Function_Form.getNull(3)
                tax = Function_Form.getNull(3)
                fx_taxable = Function_Form.getNull(3)
                fx_tax = Function_Form.getNull(3)
            End If
            Dim accmgr_id As String = ""
            Dim projcode As String = ""
            Dim deptcode As String = ""
            Dim billtype As String = ""
            Dim spcode As String = ""
            Dim source As String = Function_Form.getNull(0)
            Dim taxcode As String = ""
            Dim taxdate As String = ""
            Dim taxdate_bt As String = ""
            Dim taxdate_ds As String = Function_Form.getNull(1)
            Dim tax_basis As String = "" 'SELECT tax_basis FROM gltaxgrp (If P then P else empty)
            Dim lkdoc_type As String = ""
            Dim lkdoc_no As String = ""
            Dim lkseq As String = ""

            myConn.Open()
            Dim glcommand = New SqlCommand("SELECT * FROM gl WHERE doc_no ='" + doc_no + "' AND seq='" + seq + "'", myConn)
            Dim glreader As SqlDataReader = glcommand.ExecuteReader
            While glreader.Read()
                custcode = glreader.GetValue(glreader.GetOrdinal("accno")).ToString.Trim
                doc_type = glreader.GetValue(glreader.GetOrdinal("doc_type")).ToString.Trim
                doc_date = Function_Form.convertDateFormat(glreader.GetValue(glreader.GetOrdinal("doc_date")))
                refno = glreader.GetValue(glreader.GetOrdinal("refno")).ToString.Trim
                refno2 = glreader.GetValue(glreader.GetOrdinal("refno2")).ToString.Trim
                refno3 = glreader.GetValue(glreader.GetOrdinal("refno3")).ToString.Trim
                desp = glreader.GetValue(glreader.GetOrdinal("desp")).ToString.Trim
                desp2 = glreader.GetValue(glreader.GetOrdinal("desp2")).ToString.Trim
                desp3 = glreader.GetValue(glreader.GetOrdinal("desp3")).ToString.Trim
                desp4 = glreader.GetValue(glreader.GetOrdinal("desp4")).ToString.Trim
                remark1 = glreader.GetValue(glreader.GetOrdinal("remark1")).ToString.Trim
                remark2 = glreader.GetValue(glreader.GetOrdinal("remark2")).ToString.Trim
                cheque_no = glreader.GetValue(glreader.GetOrdinal("cheque_no")).ToString.Trim
                chqrc_date = Function_Form.convertDateFormat(glreader.GetValue(glreader.GetOrdinal("chqrc_date")))
                koff_date = Function_Form.convertDateFormat(glreader.GetValue(glreader.GetOrdinal("koff_date")))
                curr_code = glreader.GetValue(glreader.GetOrdinal("curr_code")).ToString.Trim
                refno = glreader.GetValue(glreader.GetOrdinal("refno")).ToString.Trim
                amount = glreader.GetValue(glreader.GetOrdinal("fx_amount")).ToString.Trim
                local_amount = glreader.GetValue(glreader.GetOrdinal("amount")).ToString.Trim
                accmgr_id = glreader.GetValue(glreader.GetOrdinal("accmgr_id")).ToString.Trim
                projcode = glreader.GetValue(glreader.GetOrdinal("projcode")).ToString.Trim
                deptcode = glreader.GetValue(glreader.GetOrdinal("deptcode")).ToString.Trim
                billtype = glreader.GetValue(glreader.GetOrdinal("billtype")).ToString.Trim
                spcode = glreader.GetValue(glreader.GetOrdinal("spcode")).ToString.Trim
                taxcode = glreader.GetValue(glreader.GetOrdinal("taxcode")).ToString.Trim
                taxdate = Function_Form.convertDateFormat(glreader.GetValue(glreader.GetOrdinal("taxdate")))
                taxdate_bt = Function_Form.convertDateFormat(glreader.GetValue(glreader.GetOrdinal("taxdate_bt")))
                lkdoc_type = glreader.GetValue(glreader.GetOrdinal("lkdoc_type")).ToString.Trim
                lkdoc_no = glreader.GetValue(glreader.GetOrdinal("lkdoc_no")).ToString.Trim
                lkseq = glreader.GetValue(glreader.GetOrdinal("lkseq")).ToString.Trim
            End While
            myConn.Close()

            If knockoff.Equals("0") Then
                myConn.Open()
                Dim taxgrpcommand = New SqlCommand("SELECT taxgrp FROM gltaxgrp WHERE tax_basis = 'P'", myConn)
                Dim taxgrpreader As SqlDataReader = taxgrpcommand.ExecuteReader
                Dim taxgrp As New ArrayList
                While taxgrpreader.Read()
                    taxgrp.Add(taxgrpreader.GetValue(0).ToString)
                End While
                If taxgrp.Count > 0 Then
                    Dim myTarget As New ArrayList
                    For Each range As String In rangeQuo
                        If range.Split(".")(0).ToString.Trim.Equals(row.ToString.Trim) Then
                            myTarget.Add(CInt(range.Split(".")(1).ToString.Trim))
                        End If
                    Next
                    For Each targetRow As Integer In myTarget
                        Dim targetTax As String = value_arraylist(1)(targetRow)(54).ToString.Trim
                        For Each tax_in_grp In taxgrp
                            If targetTax.Contains(tax_in_grp.ToString.Trim) Then
                                tax_basis = "P"
                                Exit For
                            End If
                        Next
                    Next
                End If
                If Not tax_basis.Equals("P") Then
                    tax_basis = Function_Form.getNull(0)
                End If
                myConn.Close()
            End If

            myConn.Open()
            Dim arcmd As String = queryTable(3)(2) + Function_Form.queryValue(custcode) +
            Function_Form.queryValue(doc_type) + Function_Form.queryValue(doc_no) +
            Function_Form.queryValue(seq) + Function_Form.queryValue(doc_date) +
            Function_Form.queryValue(due_date) + Function_Form.queryValue(refno) +
            Function_Form.queryValue(refno2) + Function_Form.queryValue(refno3) +
            Function_Form.queryValue(desp) + Function_Form.queryValue(desp2) +
            Function_Form.queryValue(desp3) + Function_Form.queryValue(desp4) +
            Function_Form.queryValue(remark1) + Function_Form.queryValue(remark2) +
            Function_Form.queryValue(cheque_no) + Function_Form.queryValue(chqrc_date) +
            Function_Form.queryValue(koff_date) + Function_Form.queryValue(curr_code) +
            Function_Form.queryValue(fx_rate) + Function_Form.queryValue(fx_gainloss) +
            Function_Form.queryValue(amount) + Function_Form.queryValue(paid) +
            Function_Form.queryValue(local_amount) + Function_Form.queryValue(local_paid) +
            Function_Form.queryValue(taxable) + Function_Form.queryValue(tax) +
            Function_Form.queryValue(fx_taxable) + Function_Form.queryValue(fx_tax) +
            Function_Form.queryValue(knockoff) + Function_Form.queryValue(accmgr_id) +
            Function_Form.queryValue(projcode) + Function_Form.queryValue(deptcode) +
            Function_Form.queryValue(billtype) + Function_Form.queryValue(spcode) +
            Function_Form.queryValue(source) + Function_Form.queryValue(taxcode) +
            Function_Form.queryValue(taxdate) + Function_Form.queryValue(taxdate_bt) +
            Function_Form.queryValue(taxdate_ds) + Function_Form.queryValue(tax_basis) +
            Function_Form.queryValue(lkdoc_type) + Function_Form.queryValue(lkdoc_no) +
            Function_Form.queryValue(lkseq)
            arcmd = arcmd.Substring(0, arcmd.Length - 1) + ")"

            Dim cmd_ar = New SqlCommand(arcmd, myConn)
            cmd_ar.ExecuteNonQuery()
            rowInsertNum += 1
            myConn.Close()
        Next

        'gloff
        For row As Integer = 0 To dgvExcel.RowCount - 1
            If Not value_arraylist(0)(row)(0).Equals("{INVALID ARRAY}") Then
                Dim queryGLOff As New ArrayList
                Dim doc_no = ""
                Dim seq = ""
                Dim pkseq = ""
                For Each seq_temp As String In gloffseq
                    If seq_temp.Split(".")(0).Trim.Equals(value_arraylist(0)(row)(1)) Then
                        doc_no = seq_temp.Split(".")(0).Trim
                        seq = seq_temp.Split(".")(2).Trim
                        pkseq = seq_temp.Split(".")(1).Trim
                    End If
                Next
                queryGLOff.Add(value_arraylist(0)(row)(10)) 'accno
                queryGLOff.Add(value_arraylist(0)(row)(0)) 'doc_type
                queryGLOff.Add(doc_no) 'doc_no
                queryGLOff.Add(seq) 'seq
                queryGLOff.Add(Function_Form.convertDateFormat(value_arraylist(0)(row)(2))) 'doc_date
                queryGLOff.Add(value_arraylist(0)(row)(0)) 'pkdoc_type
                queryGLOff.Add(doc_no) 'pkdoc_no
                queryGLOff.Add(pkseq) 'pkseq
                queryGLOff.Add(Function_Form.convertDateFormat(value_arraylist(0)(row)(2))) 'pkdoc_date
                queryGLOff.Add(1) 'sign
                queryGLOff.Add(value_arraylist(0)(row)(18)) 'pkfx_rate

                'paid
                Dim sum_p As Double = 0
                Dim amt_p1 = value_arraylist(0)(row)(75)
                Dim amt_p2 = value_arraylist(0)(row)(80)
                Dim amt_p3 = value_arraylist(0)(row)(85)
                Dim amt_p4 = value_arraylist(0)(row)(88)
                If Not value_arraylist(0)(row)(73).Equals(String.Empty) Then
                    sum_p += CDbl(amt_p1)
                End If
                If Not value_arraylist(0)(row)(78).Equals(String.Empty) Then
                    sum_p += CDbl(amt_p2)
                End If
                If Not value_arraylist(0)(row)(83).Equals(String.Empty) Then
                    sum_p += CDbl(amt_p3)
                End If
                If Not value_arraylist(0)(row)(86).Equals(String.Empty) Then
                    sum_p += CDbl(amt_p4)
                End If
                Dim paid As String = sum_p.ToString
                queryGLOff.Add(paid) 'paid

                Dim fx_rate As Double = CDbl(value_arraylist(0)(row)(18))
                Dim local_paid As Double = sum_p * fx_rate
                queryGLOff.Add(fx_rate) 'fx_rate
                queryGLOff.Add(local_paid) 'local_paid
                queryGLOff.Add(local_paid) 'local_amount
                queryGLOff.Add(Function_Form.getNull(3)) 'fx_gainloss
                queryGLOff.Add(Function_Form.getNull(3)) 'v_gainloss

                If sum_p = 0 Then
                    Continue For
                End If

                Dim gloff_cmd As String = queryTable(5)(2)
                For j = 0 To queryGLOff.Count - 1
                    gloff_cmd += "'" + queryGLOff(j).ToString + "',"
                Next
                gloff_cmd = gloff_cmd.Substring(0, gloff_cmd.Length - 1) + ")"

                myConn.Open()
                Dim cmd_gloff = New SqlCommand(gloff_cmd, myConn)
                cmd_gloff.ExecuteNonQuery()
                rowInsertNum += 1
                myConn.Close()
            End If
        Next

        Dim rowUpdateNum = 0
        'prodsn(7) + stocksn(8)
        For row As Integer = 0 To dgvExcel.RowCount - 1
            If Not value_arraylist(7)(row)(1).ToString.Trim.Equals(String.Empty) Then
                Dim serialnos As New List(Of String)(value_arraylist(7)(row)(1).ToString.Trim.Split(","c))
                For sn = 0 To serialnos.Count - 1
                    Dim serialno As String = serialnos(sn)
                    Dim qty = "-1"
                    Dim location = value_arraylist(7)(row)(4)
                    Dim doc_no = value_arraylist(7)(row)(8)
                    Dim line_no = value_arraylist(7)(row)(9)
                    Dim doc_date = Convert.ToDateTime(value_arraylist(7)(row)(10)).ToString("yyyy-MM-dd HH:mm:ss")
                    Dim procode = value_arraylist(7)(row)(0)
                    Dim serialNoProdCommand As String = "UPDATE prodsn SET "
                    Dim serialNoColumns = "qty='" + qty + "',"
                    serialNoColumns += "location='" + location + "',"
                    serialNoColumns += "doc_type='SI',"
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
                    serialNoStockdCommand += procode + "','" + serialno + "','SI','" + doc_no + "','" + line_no + "','" + doc_date + "','" + qty + "','" + location + "')"
                    Dim command2 = New SqlCommand(serialNoStockdCommand, myConn)
                    command2.ExecuteNonQuery()
                    'MsgBox(serialNoStockdCommand)
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
                Dim command_temp = New SqlCommand("SELECT TOP 1 dkey FROM sinvdet WHERE doc_no ='" + value_arraylist(2)(row)(2) + "' AND line_no ='" + value_arraylist(2)(row)(4) + "'", myConn)
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
                insertArray.Add(value_arraylist(0)(mySource)(10)) 'custcode
                insertArray.Add(Function_Form.getNull(0)) 'suppcode
                insertArray.Add(value_arraylist(0)(mySource)(8)) 'refno
                insertArray.Add(value_arraylist(0)(mySource)(9)) 'refno2
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
                'Clipboard.SetText(stockCommand)
                'MsgBox(stockCommand)
                command.ExecuteNonQuery()
                rowInsertNum += 1
                myConn.Close()
            End If
        Next

        Function_Form.promptImportSuccess(rowInsertNum, rowUpdateNum)
        Function_Form.printExcelResult("Sales_Invoice", queryTable, value_arraylist, sql_format_arraylist, dgvExcel)
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