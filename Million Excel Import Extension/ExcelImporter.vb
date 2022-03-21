Imports System.Data.SqlClient
Imports System.IO
Imports ClosedXML.Excel
Imports ExcelDataReader

Public Class ExcelImporter
    Dim tables As DataTableCollection
    Private serverName As String
    Private database As String
    Private myConn As SqlConnection
    Private statusConnection As Boolean
    Private pwd_query As String
    Private import_type As String
    Private data_type_group As New List(Of String) From {"one", "two", "three"}
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

    Private Sub CreateTemplate_Click(sender As Object, e As EventArgs) Handles btnCreateTemplate.Click
        test()
        'excelTemplate(cbType.Text.ToString)
    End Sub
    Private Sub excelTemplate(templateType As String)
        Using workbook As New XLWorkbook
            Dim worksheet As IXLWorksheet = workbook.Worksheets.Add(templateType)
            'Dim worksheetLR As Integer = 1
            Dim tableExcelSetting As DataTableCollection
            Try
                Using stream = File.Open(getMaintainSetting, FileMode.Open, FileAccess.Read)
                    Using reader As IExcelDataReader = ExcelReaderFactory.CreateReader(stream)
                        Dim result As DataSet = reader.AsDataSet(New ExcelDataSetConfiguration() With {
                                                                     .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
                                                                     .UseHeaderRow = True}})
                        tableExcelSetting = result.Tables
                        Dim excelUnique As New ArrayList
                        arrayExtendFromExcelSheet(tableExcelSetting, "Quotation", excelUnique, worksheet)
                        arrayExtendFromExcelSheet(tableExcelSetting, "Quotation Desc", excelUnique, worksheet)

                    End Using
                End Using
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical)
            End Try
            Using sfd As SaveFileDialog = New SaveFileDialog() With {.Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls"}
                If sfd.ShowDialog() = DialogResult.OK Then
                    worksheet.Columns.Width = 25
                    workbook.SaveAs(sfd.FileName)
                    MsgBox("File saved in " + sfd.FileName)
                End If
            End Using
        End Using
    End Sub
    Private Sub arrayExtendFromExcelSheet(tableExcelSetting As DataTableCollection, sheet As String, excelUnique As ArrayList, worksheet As IXLWorksheet)
        Dim dt As DataTable = tableExcelSetting(sheet)
        For Each row As DataRow In dt.Rows
            Dim sql_value = row(0).ToString
            Dim excel_value = row(1).ToString
            Dim default_value = row(2).ToString
            Dim formula = row(3).ToString
            If excel_value <> String.Empty Then
                Dim repeated As Boolean = False
                For Each al As String In excelUnique
                    If excel_value.Equals(al) Then
                        repeated = True
                    End If
                Next
                If repeated = False Then
                    excelUnique.Add(excel_value)
                    worksheet.Cell(1, excelUnique.Count).Value = excel_value
                End If
            End If
        Next
    End Sub
    Private Function getMaintainSetting() As String
        Return SQL_Connection_Form.returnUpperFolder(Application.StartupPath(), 2) + "maintain.xls"
    End Function
    'Private Sub cbType_SelectedIndexChanged(sender As Object, e As EventArgs)
    '    If Not (cbType.SelectedItem.ToString.Equals(String.Empty)) Then
    '        btnCreateTemplate.Enabled = True
    '        btnImport.Enabled = True
    '    Else
    '        btnCreateTemplate.Enabled = False
    '        btnImport.Enabled = False
    '    End If
    'End Sub

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
                If txtType.Text.Equals("Quotation") Then
                    queryTable.Add(New ArrayList)
                    queryTable.Add(New ArrayList)
                    queryTable(0).add("Quotation") '0
                    queryTable(0).add("squote") '1
                    queryTable(1).add("Quotation Desc")
                    queryTable(1).add("squotedet")
                ElseIf txtType.Text.Equals("SalesOrder") Then

                End If
                quotationWriteIntoSQL(tableExcelSetting, queryTable)
            End Using
        End Using
    End Sub
    Private Sub quotationWriteIntoSQL(tableExcelSetting As DataTableCollection, queryTable As ArrayList)
        Dim sql_format_arraylist = New ArrayList
        Dim excel_format_arraylist = New ArrayList
        Dim data_type_arraylist = New ArrayList
        Dim default_arraylist = New ArrayList
        Dim format_arraylist = New ArrayList
        Dim value_arraylist = New ArrayList
        For j As Integer = 0 To queryTable.Count - 1
            sql_format_arraylist.Add(New ArrayList)
            excel_format_arraylist.Add(New ArrayList)
            data_type_arraylist.Add(New ArrayList)
            default_arraylist.Add(New ArrayList)
            value_arraylist.Add(New ArrayList)
            format_arraylist.Add(New ArrayList)
            Dim dtTemp As DataTable = tableExcelSetting(queryTable(j)(0))
            For Each row As DataRow In dtTemp.Rows
                Dim sql_value = row(0).ToString
                Dim excel_value = row(1).ToString
                Dim data_type = row(4).ToString
                Dim default_value = row(2).ToString
                Dim format_value = row(3).ToString
                sql_format_arraylist(j).add(sql_value)
                data_type_arraylist(j).add(data_type)
                excel_format_arraylist(j).add(excel_value)
                default_arraylist(j).add(default_value)
                format_arraylist(j).add(format_value)
                'For i As Integer = 0 To dgvExcel.ColumnCount - 1
                '    Dim dgvHeader = dgvExcel.Columns(i).HeaderText
                '    If excel_value.Equals(dgvHeader) Then
                '        'queryTable(j).Add(sql_value)
                '        'sql_format_arraylist(j).add(sql_value)
                '        'excel_format_arraylist(j).add(excel_value)
                '        'quotationTable.Add(sql_value)
                '    End If
                'Next
            Next
        Next
        'For i As Integer = 0 To sql_format_arraylist.Count - 1
        '    Dim v = 0
        '    Dim strs = ""
        '    For Each str As String In sql_format_arraylist(i)
        '        strs += str + vbTab
        '        v += 1
        '    Next
        '    MsgBox(v.ToString + " (sql Count) : " + strs)
        'Next
        'For i As Integer = 0 To default_arraylist.Count - 1
        '    Dim v = 0
        '    Dim strs = ""
        '    For Each str As String In default_arraylist(i)
        '        strs += str + vbTab
        '        v += 1
        '    Next
        '    MsgBox(v.ToString + " (default_arraylist Count) : " + strs)
        'Next
        For m As Integer = 0 To queryTable.Count - 1
            Dim query As String = "INSERT INTO " + queryTable(m)(1) + " ("
            Dim values As String = ""
            For n As Integer = 0 To sql_format_arraylist(m).Count - 1
                values += "," + sql_format_arraylist(m)(n)
            Next
            query += values.Substring(1, values.Length - 1) + ") VALUES ("
            'values = ""
            'For b As Integer = 0 To sql_format_arraylist(m).Count - 1
            '    values += ",@" + sql_format_arraylist(m)(b)
            'Next
            'query += values.Substring(1, values.Length - 1) + ")"
            'MsgBox(query)
            queryTable(m).add(query)
        Next
        'MsgBox(queryTable(0)(2))
        'For Each al As ArrayList In queryTable
        '    Dim strs = ""
        '    For Each str As String In al
        '        strs += str + vbNewLine
        '    Next
        '    MsgBox(strs)
        'Next
        'Return
        For i As Integer = 0 To queryTable.Count - 1
            init()
            Using myConn = New SqlConnection("Data Source=" + serverName + ";" & "Initial Catalog=" + database + ";" + pwd_query)
                Using command As New SqlCommand("", myConn)
                    For row As Integer = 0 To dgvExcel.RowCount - 2
                        Using r As DataGridViewRow = dgvExcel.Rows(row)
                            value_arraylist(i).Add(New ArrayList)
                            Dim query = ""
                            Dim queryable = False
                            For u As Integer = 0 To sql_format_arraylist(i).Count - 1
                                Dim sql_temp = sql_format_arraylist(i)(u).ToString.Trim
                                Dim data_type_temp = data_type_arraylist(i)(u).ToString.Trim
                                Dim default_temp = default_arraylist(i)(u).ToString.Trim
                                'Dim value_temp = value_arraylist(i)(u).ToString.Trim
                                Dim add_value As Boolean = False
                                Dim value = ""
                                For q As Integer = 0 To r.Cells.Count - 1
                                    Dim cellValue As Object = r.Cells(q).Value.ToString.Trim
                                    Dim headerText As String = dgvExcel.Columns(q).HeaderText.Trim
                                    Dim excel_temp = excel_format_arraylist(i)(u).ToString.Trim
                                    If headerText.Equals(excel_temp) Then
                                        'q = r.Cells.Count
                                        'value_arraylist(i)(row).add(cellValue)
                                        'add_value = True
                                        'value = cellValue
                                        If Not (cellValue.Equals(String.Empty)) Then
                                            q = r.Cells.Count
                                            value_arraylist(i)(row).add(cellValue)
                                            add_value = True
                                            value = cellValue
                                            If data_type_temp.Contains("PK") Then
                                                queryable = True
                                            End If
                                        End If
                                    End If
                                Next
                                If add_value = False Then
                                    If default_temp.Equals(String.Empty) Then
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
                                    Else
                                        'If default_temp.Contains("getstringonly_") Then
                                        '    Dim default_index As Integer = Convert.ToInt32(default_temp.Substring(14))
                                        '    MsgBox(default_index.ToString)
                                        '    MsgBox(value_arraylist(i).count.ToString)
                                        '    'MsgBox("getonly" + vbTab + value_arraylist(i)(default_index).ToString)
                                        '    'value = value_arraylist(i)(default_index)
                                        '    'value_arraylist(i).add("DEFAULT_VALUE")
                                        'End If
                                        value = "{DEFAULT_VALUE}"
                                        value_arraylist(i)(row).add("{DEFAULT_VALUE}")
                                    End If
                                End If
                                'MsgBox("sql_temp:" + vbTab + sql_temp + vbTab + "value:" + vbTab + value_arraylist(i).count.ToString + vbTab + "count:" + vbTab + u.ToString)
                                'query += "'" + value + "',"
                            Next
                            'MsgBox("sql: " + vbTab + (sql_format_arraylist(i).count).ToString + " default: " + default_arraylist(i).count.ToString)
                            For g As Integer = 0 To value_arraylist(i)(row).count - 1
                                Dim value_temp As String = value_arraylist(i)(row)(g).ToString.Trim
                                If value_temp.Equals("{DEFAULT_VALUE}") Then
                                    'MsgBox("i=" + i.ToString + vbTab + "g=" + g.ToString)
                                    Dim default_temp = default_arraylist(i)(g).ToString.Trim
                                    'Dim default_temp = default_arraylist(i)(g).ToString.Trim
                                    If default_temp.Contains("getstringonly_") Then
                                        Dim default_index As Integer = Convert.ToInt32(default_temp.Substring(14))
                                        Dim stringonly_char As Char() = value_arraylist(i)(row)(default_index).ToString.ToCharArray()
                                        Dim stringonly_str = ""
                                        For Each ch As Char In stringonly_char
                                            If Not Char.IsDigit(ch) Then
                                                stringonly_str += ch
                                            End If
                                        Next
                                        'MsgBox("COUNT: " + i.ToString + vbTab + value_arraylist(i)(default_index))
                                        'value = value_arraylist(i)(default_index)
                                        value_arraylist(i)(row)(g) = stringonly_str
                                    ElseIf default_temp.Contains("FK") Then
                                        If row > 0 Then
                                            Dim fk_value_temp = ""
                                            For h = row - 1 To 0
                                                If fk_value_temp.Equals(String.Empty) Then
                                                    fk_value_temp = value_arraylist(i)(h)(g).ToString.Trim
                                                    h = 0
                                                End If
                                            Next
                                            MsgBox("Value in table " + i.ToString + " with row " + row.ToString + " in value " + g.ToString + " => " + fk_value_temp)
                                            value_arraylist(i)(row)(g) = fk_value_temp + "LMAO"
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
                                    Else
                                        value_arraylist(i)(row)(g) = default_temp
                                    End If
                                End If
                                query += "'" + value_arraylist(i)(row)(g) + "',"
                            Next
                            'MsgBox(queryTable(i)(2) + query)
                            'MsgBox("complete")
                            If queryable Then
                                'For b = 0 To default_arraylist(i).Count - 1
                                '    Dim default_value_temp = default_arraylist(i)(b)
                                '    If Not (default_value_temp.Equals(String.Empty)) Then

                                '    End If
                                'Next
                                Dim command_text As String = queryTable(i)(2) + query
                                command_text = command_text.Substring(0, command_text.Length - 1) + ")"
                                MsgBox(command_text)
                                'queryTable(i)(2) += ")"
                                'MsgBox(queryTable(i)(2) + query)
                                'MsgBox(command.CommandText)
                                command.CommandText = command_text
                                'myConn.Open()
                                'command.ExecuteNonQuery()
                                'myConn.Close()
                            Else
                                'Dim command_text As String = queryTable(i)(2) + query
                                'command_text = command_text.Substring(0, command_text.Length - 1) + ")"
                                'MsgBox(command_text)
                            End If
                            'MsgBox("Queryable: " + queryable.ToString + vbNewLine + queryTable(i)(2) + query)
                        End Using
                    Next
                End Using
            End Using
        Next
    End Sub
    Private Sub test()
        'dgvExcel.Columns(1).DefaultCellStyle.Format = "d"
        'init()
        'myConn = New SqlConnection("Data Source=" + serverName + ";" & "Initial Catalog=test_database;" + pwd_query)
        'Dim command As New SqlCommand("INSERT INTO table_1(name,lastupdate,price) VALUES (@name,@lastupdate,@price)", myConn)
        'command.Parameters.Add("@name", SqlDbType.Char).Value = "Name1"
        'command.Parameters.Add("@lastupdate", SqlDbType.Char).Value = "2022-04-05 00:00:00.000"
        'command.Parameters.Add("@price", SqlDbType.Char).Value = 585
        'myConn.Open()
        'If command.ExecuteNonQuery = 1 Then
        '    MsgBox("win")
        'Else
        '    MsgBox("lose")
        'End If
        'myConn.Close()
    End Sub

    'Private Sub btnSync_Click(sender As Object, e As EventArgs) Handles btnSync.Click
    '    init()
    '    myConn = New SqlConnection("Data Source=" + serverName + ";" & "Initial Catalog=" + database + ";" + pwd_query)
    '    Dim command As SqlCommand
    '    command = myConn.CreateCommand
    '    If txtType.Text.Equals("Quotation") Then
    '        command.CommandText = "SELECT DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'squote'"
    '    End If
    '    myConn.Open()
    '    Dim myReader As SqlDataReader
    '    Dim al_dataType As New ArrayList
    '    myReader = command.ExecuteReader()
    '    Do While myReader.Read()
    '        al_dataType.Add(myReader.GetString(0))
    '    Loop
    '    myConn.Close()
    '    Using workbook As New XLWorkbook
    '        Dim worksheet As IXLWorksheet = workbook.Worksheets.Add("Quotation")
    '        'Dim worksheetLR As Integer = 1
    '        Dim tableExcelSetting As DataTableCollection
    '        Try
    '            Using stream = File.Open(getMaintainSetting, FileMode.Open, FileAccess.Read)
    '                Using reader As IExcelDataReader = ExcelReaderFactory.CreateReader(stream)
    '                    Dim result As DataSet = reader.AsDataSet(New ExcelDataSetConfiguration() With {
    '                                                                 .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
    '                                                                 .UseHeaderRow = True}})
    '                    tableExcelSetting = result.Tables
    '                    Dim excelUnique As New ArrayList
    '                    arrayExtendFromExcelSheet(tableExcelSetting, "Quotation", excelUnique, worksheet)
    '                    arrayExtendFromExcelSheet(tableExcelSetting, "Quotation Desc", excelUnique, worksheet)

    '                End Using
    '            End Using
    '        Catch ex As Exception
    '            MsgBox(ex.Message, MsgBoxStyle.Critical)
    '        End Try
    '        Using sfd As SaveFileDialog = New SaveFileDialog() With {.Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls"}
    '            If sfd.ShowDialog() = DialogResult.OK Then
    '                worksheet.Columns.Width = 25
    '                workbook.SaveAs(sfd.FileName)
    '                MsgBox("File saved in " + sfd.FileName)
    '            End If
    '        End Using
    '    End Using

    'End Sub
    Private Sub runQueryTest()
        Dim str = "INSERT INTO [dbo].[squote]
           ([doc_type]
           ,[doc_no]
           ,[doc_date]
           ,[doc_desp]
           ,[doc_desp2]
           ,[doc_status]
           ,[doc_set]
           ,[refno]
           ,[refno2]
           ,[custcode]
           ,[name]
           ,[addr]
           ,[addrkey]
           ,[attn2]
           ,[terms]
           ,[accmgr_id]
           ,[curr_code]
           ,[fx_rate]
           ,[discount1]
           ,[discount2]
           ,[discount]
           ,[dispec1]
           ,[dispec2]
           ,[taxcode]
           ,[taxpec1]
           ,[taxpec2]
           ,[taxable]
           ,[tax1]
           ,[tax2]
           ,[taxadj]
           ,[tax]
           ,[rndoff]
           ,[subtotal]
           ,[nett]
           ,[total]
           ,[local_gross]
           ,[local_discount]
           ,[local_nett]
           ,[local_tax1]
           ,[local_tax2]
           ,[local_tax]
           ,[local_rndoff]
           ,[local_total]
           ,[local_totalrp]
           ,[deposit]
           ,[sign]
           ,[signqty]
           ,[batchno]
           ,[projcode]
           ,[deptcode]
           ,[shipmethod]
           ,[hremark1]
           ,[hremark2]
           ,[hremark3]
           ,[hremark4]
           ,[hremark5]
           ,[hremark6]
           ,[hremark7]
           ,[hremark8]
           ,[hremark9]
           ,[hremark10]
           ,[hremark11]
           ,[hremark12]
           ,[hremark13]
           ,[hremark14]
           ,[hremark15]
           ,[remark1]
           ,[counter_no]
           ,[cashier_id]
           ,[buyer_id]
           ,[buyer_name]
           ,[cpdoc_type]
           ,[cpdoc_no]
           ,[cpdoc_date]
           ,[printcnt]
           ,[approvedby]
           ,[approveddate]
           ,[tsdate]
           ,[createdby]
           ,[updatedby]
           ,[createdate]
           ,[lastupdate])
     VALUES
           ('SQ','SQ000004','4-2-2022 12:00:00 AM','Quotation','   ','2','1','   ','   ','3000/B01','Best Beauty Sdn Bhd','4, Jln. Kuning 2, Tmn. Pelangi, 80400 JB, Johor','0','   ','0','   ','MYR','1','0','0','0','0','0','   ','0','0','2117','0','0','0','211.7','0','2117','2117','2328.7','2117','0','2117','0','0','211.7','0','2328.7','2117','0','1','-1','220410','   ','   ','   ','   ','   ','   ','   ','   ','   ','   ','   ','   ','   ','   ','   ','   ','   ','   ','   ','   ','   ','   ','   ','   ','   ','1-1-1900 12:00:00 AM','0','   ','1-1-1900 12:00:00 AM','1-1-1900 12:00:00 AM','admin','admin','3-1-2022 3:47:50 PM','3-1-2022 3:47:50 PM')"
        Using myConn = New SqlConnection("Data Source=" + serverName + ";" & "Initial Catalog=" + database + ";" + pwd_query)
            Using command As New SqlCommand(str, myConn)
                myConn.Open()
                command.ExecuteNonQuery()
                myConn.Close()
            End Using
        End Using
    End Sub
End Class

