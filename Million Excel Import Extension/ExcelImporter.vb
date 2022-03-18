Imports System.Data.SqlClient
Imports ExcelDataReader
Imports System.IO
Imports ClosedXML.Excel

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
                'writeIntoQueryTable(tableExcelSetting, queryTable)
                'Dim strs As String = ""
                'For i As Integer = 1 To queryTable(0).count - 1
                '    strs += queryTable(0)(i) + vbNewLine
                'Next
                'MsgBox(strs)



                'MsgBox()
                'Command = strs.Substring(1, strs.Length - 1)
                'Using myConn
                '    Using command As New SqlCommand("INSERT INTO table_master(item, price) VALUES(@item, @price)",
                '                                    myConn)

                '        myConn.Open()

                '        ' Create and add the parameters, just one time here with dummy values or'
                '        ' use the full syntax to create each single the parameter'
                '        command.Parameters.AddWithValue("@item", "")
                '        command.Parameters.AddWithValue("@price", 0)

                '        For Each r As DataGridViewRow In dgvMain.Rows
                '            If (Not String.IsNullOrWhiteSpace(r.Cells(1).Value)) Then

                '                command.Parameters("@item").Value = r.Cells(1).Value.Trim
                '                command.Parameters("@price").Value = r.Cells(2).Value
                '                command.ExecuteNonQuery()
                '            End If
                '        Next

                '    End Using
                'End Using
            End Using
        End Using
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical)
        'End Try
        'Try
        '    myConn = New SqlConnection("Data Source=" + serverName + ";" &
        '                            "Initial Catalog=master;" + pwd_query)
        '    myConn.Open()
        '    Using myConn
        '        Dim cmd As New SqlCommand("Select name from master.dbo.sysdatabases WHERE dbid > 4;", myConn)
        '        Using reader As SqlDataReader = cmd.ExecuteReader()
        '            cbDatabase.Items.Clear()
        '            While reader.Read()
        '                cbDatabase.Items.Add(reader.GetValue(0))
        '            End While
        '        End Using
        '        If (cbDatabase.Items.Count > 0) Then
        '            updateSQLStatus(1, "")
        '        Else
        '            updateSQLStatus(0, "Could not find any database from the server!")
        '        End If
        '    End Using
        '    myConn.Close()
        'Catch ex As Exception
        '    If (ex.GetHashCode = 49205706) Then
        '        updateSQLStatus(0, "Failed to connect the server!")
        '    Else
        '        MsgBox(ex.Message, MsgBoxStyle.Critical, Me.Text)
        '    End If
        'End Try
    End Sub
    Private Sub quotationWriteIntoSQL(tableExcelSetting As DataTableCollection, queryTable As ArrayList)
        Dim sql_format_arraylist = New ArrayList
        Dim excel_format_arraylist = New ArrayList
        Dim data_type_arraylist = New ArrayList
        For j As Integer = 0 To queryTable.Count - 1
            sql_format_arraylist.Add(New ArrayList)
            excel_format_arraylist.Add(New ArrayList)
            data_type_arraylist.Add(New ArrayList)
            Dim dtTemp As DataTable = tableExcelSetting(queryTable(j)(0))
            For Each row As DataRow In dtTemp.Rows
                Dim sql_value = row(0).ToString
                Dim excel_value = row(1).ToString
                Dim data_type = row(4).ToString
                sql_format_arraylist(j).add(sql_value)
                data_type_arraylist(j).add(data_type)
                For i As Integer = 0 To dgvExcel.ColumnCount - 1
                    Dim dgvHeader = dgvExcel.Columns(i).HeaderText
                    If excel_value.Equals(dgvHeader) Then
                        'queryTable(j).Add(sql_value)
                        'sql_format_arraylist(j).add(sql_value)
                        excel_format_arraylist(j).add(excel_value)
                        'quotationTable.Add(sql_value)
                    End If
                Next
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
        'For i As Integer = 0 To excel_format_arraylist.Count - 1
        '    Dim v = 0
        '    Dim strs = ""
        '    For Each str As String In excel_format_arraylist(i)
        '        strs += str + vbTab
        '        v += 1
        '    Next
        '    MsgBox(v.ToString + " (excel Count) : " + strs)
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
        For Each al As ArrayList In queryTable
            Dim strs = ""
            For Each str As String In al
                strs += str + vbNewLine
            Next
            'MsgBox(strs)
        Next
        For i As Integer = 0 To queryTable.Count - 1
            init()
            Using myConn = New SqlConnection("Data Source=" + serverName + ";" & "Initial Catalog=" + database + ";" + pwd_query)
                Using command As New SqlCommand("", myConn)
                    'command.CommandType = Text
                    ' Create and add the parameters, just one time here with dummy values or'
                    ' use the full syntax to create each single the parameter'
                    For y As Integer = 0 To excel_format_arraylist.Count - 1
                        'MsgBox("Excel:" + vbTab + excel_format_arraylist(i)(y))
                    Next
                    For j As Integer = 0 To sql_format_arraylist(i).count - 1
                        'command.Parameters.AddWithValue("@" + sql_format_arraylist(i)(j), "")
                    Next
                    For u As Integer = 0 To sql_format_arraylist.Count - 1
                        Dim sql_temp = sql_format_arraylist(i)(u).ToString.Trim
                        'Dim cmd As New SqlCommand("INSERT INTO " + queryTable(i)(1) + " (" + sql_temp + ")VALUES('" + +"')", myConn)
                        'Dim cmd As New SqlCommand
                        'Dim cmdText As String = "INSERT INTO " + queryTable(i)(1) + " (" + sql_temp + ")VALUES('"
                        For row As Integer = 0 To dgvExcel.RowCount - 2
                            Using r As DataGridViewRow = dgvExcel.Rows(row)
                                Dim queryable As Boolean
                                queryable = False
                                'Dim strtemp As String = ""
                                For q As Integer = 0 To r.Cells.Count - 1
                                    Dim cellValue As Object = r.Cells(q).Value.ToString.Trim
                                    Dim headerText As String = dgvExcel.Columns(q).HeaderText.Trim
                                    Dim matched_with_header = False
                                    For y As Integer = 0 To excel_format_arraylist(i).Count - 1
                                        Dim excel_temp = excel_format_arraylist(i)(y).ToString.Trim
                                        'MsgBox(excel_temp + vbTab + headerText)
                                        If excel_temp.Equals(headerText) Then
                                            If Not (cellValue.Equals(String.Empty)) Then
                                                queryTable(i)(2) += "'" + cellValue + "',"
                                                'cmdText += cellValue
                                                'cmd.CommandText = "INSERT INTO " + queryTable(i)(1) + " (" + sql_temp + ")VALUES('" + +"')"
                                                'command.Parameters.Add("@" + sql_temp, SqlDbType.Char).Value = cellValue
                                            Else
                                                If data_type_arraylist(i)(y).ToString.Contains("char") Or data_type_arraylist(i)(y).ToString.Contains("text") Then
                                                    'command.Parameters.Add("@" + sql_temp, SqlDbType.Char).Value = " "
                                                    queryTable(i)(2) += "' ',"
                                                    'cmdText += " "
                                                    'MsgBox("CHAR->" + excel_temp)
                                                ElseIf data_type_arraylist(i)(y).ToString.Contains("date") Or data_type_arraylist(i)(y).ToString.Contains("time") Then
                                                    'command.Parameters.Add("@" + sql_temp, SqlDbType.Char).Value = New DateTime
                                                    queryTable(i)(2) += "'" + (New DateTime).ToString + "',"
                                                    'cmdText += (New DateTime).ToString
                                                    'MsgBox("DATETIME->" + excel_temp)
                                                Else
                                                    'command.Parameters.Add("@" + sql_temp, SqlDbType.Char).Value = 0
                                                    queryTable(i)(2) += "'0',"
                                                    'cmdText += "0"
                                                    'MsgBox("NUMBER->" + excel_temp)
                                                End If
                                            End If
                                            'MsgBox(cmdText)
                                            'cmdText += +"')"
                                            'Dim cmd As New SqlCommand(cmdText, myConn)
                                            ''cmd.CommandText = cmdText
                                            ''cmd.CommandType = myConn
                                            'myConn.Open()
                                            'cmd.ExecuteNonQuery()
                                            'myConn.Close()
                                            queryable = True
                                        Else
                                            If data_type_arraylist(i)(y).ToString.Contains("char") Or data_type_arraylist(i)(y).ToString.Contains("text") Then
                                                'command.Parameters.Add("@" + sql_temp, SqlDbType.Char).Value = " "
                                                queryTable(i)(2) += "' ',"
                                                'MsgBox("CHAR->" + excel_temp)
                                            ElseIf data_type_arraylist(i)(y).ToString.Contains("date") Or data_type_arraylist(i)(y).ToString.Contains("time") Then
                                                'command.Parameters.Add("@" + sql_temp, SqlDbType.Char).Value = New DateTime
                                                queryTable(i)(2) += "'" + (New DateTime).ToString + "',"
                                                'MsgBox("DATETIME->" + excel_temp)
                                            Else
                                                'command.Parameters.Add("@" + sql_temp, SqlDbType.Char).Value = 0
                                                queryTable(i)(2) += "'0',"
                                                'MsgBox("NUMBER->" + excel_temp)
                                            End If
                                            'cmdText += +"')"
                                            'Dim cmd As New SqlCommand(cmdText, myConn)
                                            ''cmd.CommandText = cmdText
                                            ''cmd.CommandType = myConn
                                            'myConn.Open()
                                            'cmd.ExecuteNonQuery()
                                            'myConn.Close()
                                        End If
                                    Next
                                Next
                                If queryable Then
                                    queryTable(i)(2) += ")"
                                    MsgBox(queryTable(i)(2))
                                    'MsgBox(command.CommandText)
                                    'myConn.Open()
                                    'command.ExecuteNonQuery()
                                    'myConn.Close()
                                End If
                            End Using
                        Next
                    Next


                End Using
            End Using
        Next
    End Sub
    Private Sub writeIntoQueryTable(tableExcelSetting As DataTableCollection, queryTable As ArrayList)
        'writeIntoQueryTable(tableExcelSetting, queryTable)
        Dim sql_format_arraylist = New ArrayList
        Dim excel_format_arraylist = New ArrayList
        Dim data_type_arraylist = New ArrayList
        For j As Integer = 0 To queryTable.Count - 1
            sql_format_arraylist.Add(New ArrayList)
            excel_format_arraylist.Add(New ArrayList)
            data_type_arraylist.Add(New ArrayList)
            Dim dtTemp As DataTable = tableExcelSetting(queryTable(j)(0))
            For Each row As DataRow In dtTemp.Rows
                Dim sql_value = row(0).ToString
                Dim excel_value = row(1).ToString
                Dim data_type = row(4).ToString
                For i As Integer = 0 To dgvExcel.ColumnCount - 1
                    Dim dgvHeader = dgvExcel.Columns(i).HeaderText
                    If excel_value.Equals(dgvHeader) Then
                        'queryTable(j).Add(sql_value)
                        sql_format_arraylist(j).add(sql_value)
                        excel_format_arraylist(j).add(excel_value)
                        data_type_arraylist(j).add(data_type)
                        'quotationTable.Add(sql_value)
                    End If
                Next
            Next
        Next
        'For i As Integer = 0 To sql_format_arraylist.Count - 1
        '    Dim strs = ""
        '    For Each str As String In sql_format_arraylist(i)
        '        strs += str + vbTab
        '    Next
        '    MsgBox(strs)
        'Next
        'For Each al As ArrayList In data_type_arraylist
        '    Dim strs = ""
        '    For Each str As String In al
        '        strs += str + vbTab
        '    Next
        '    MsgBox(strs)
        'Next
        For m As Integer = 0 To queryTable.Count - 1
            Dim query As String = "INSERT INTO " + queryTable(m)(1) + " ("
            Dim values As String = ""
            For n As Integer = 0 To sql_format_arraylist(m).Count - 1
                values += "," + sql_format_arraylist(m)(n)
            Next
            query += values.Substring(1, values.Length - 1) + ") VALUES ("
            values = ""
            For b As Integer = 0 To sql_format_arraylist(m).Count - 1
                values += ",@" + sql_format_arraylist(m)(b)
            Next
            query += values.Substring(1, values.Length - 1) + ")"
            'MsgBox(query)
            queryTable(m).add(query)
        Next
        For Each al As ArrayList In queryTable
            Dim strs = ""
            For Each str As String In al
                strs += str + vbNewLine
            Next
            'MsgBox(strs)
        Next
        For i As Integer = 0 To queryTable.Count - 1
            init()
            Using myConn = New SqlConnection("Data Source=" + serverName + ";" & "Initial Catalog=" + database + ";" + pwd_query)
                Using command As New SqlCommand(queryTable(i)(2), myConn)
                    'command.CommandType = Text
                    ' Create and add the parameters, just one time here with dummy values or'
                    ' use the full syntax to create each single the parameter'
                    For y As Integer = 0 To excel_format_arraylist.Count - 1
                        'MsgBox("Excel:" + vbTab + excel_format_arraylist(i)(y))
                    Next
                    For j As Integer = 0 To sql_format_arraylist(i).count - 1
                        'command.Parameters.AddWithValue("@" + sql_format_arraylist(i)(j), "")
                    Next
                    For row As Integer = 0 To dgvExcel.RowCount - 2
                        Using r As DataGridViewRow = dgvExcel.Rows(row)
                            Dim queryable As Boolean
                            queryable = False
                            'Dim strtemp As String = ""
                            For q As Integer = 0 To r.Cells.Count - 1
                                Dim cellValue As Object = r.Cells(q).Value.ToString.Trim
                                Dim headerText As String = dgvExcel.Columns(q).HeaderText.Trim
                                'If Not (cellValue.Equals(String.Empty)) Then
                                '    For y As Integer = 0 To excel_format_arraylist(i).Count - 1
                                '        Dim excel_temp = excel_format_arraylist(i)(y).ToString.Trim
                                '        'MsgBox(excel_temp + vbTab + headerText)
                                '        If excel_temp.Equals(headerText) Then
                                '            Dim sql_temp = sql_format_arraylist(i)(y).ToString.Trim
                                '            If cellValue.ToString.Equals(String.Empty) Then
                                '                MsgBox("Empty: " + headerText)
                                '                command.Parameters.Add("@" + sql_format_arraylist(i)(y), SqlDbType.Char).Value = 0
                                '            Else
                                '                command.Parameters.Add("@" + sql_format_arraylist(i)(y), SqlDbType.Char).Value = cellValue
                                '            End If
                                '            'command.Parameters.Add("@" + sql_format_arraylist(i)(y), SqlDbType.Char).Value = cellValue
                                '            MsgBox(excel_temp + vbTab + headerText)
                                '            'MsgBox(headerText + vbTab + cellValue)
                                '            'command.Parameters("@" + sql_temp).Value = cellValue
                                '            queryable = True
                                '        End If
                                '    Next
                                'Else

                                'End If
                                For y As Integer = 0 To excel_format_arraylist(i).Count - 1
                                    Dim excel_temp = excel_format_arraylist(i)(y).ToString.Trim
                                    'MsgBox(excel_temp + vbTab + headerText)
                                    If excel_temp.Equals(headerText) Then
                                        Dim sql_temp = sql_format_arraylist(i)(y).ToString.Trim
                                        If Not (cellValue.Equals(String.Empty)) Then
                                            command.Parameters.Add("@" + sql_format_arraylist(i)(y), SqlDbType.Char).Value = cellValue
                                        Else
                                            'MsgBox(excel_temp + vbTab + "(" + headerText + ")" + vbNewLine + "Data Type: " + data_type_arraylist(i)(y))
                                            'MsgBox()
                                            If data_type_arraylist(i)(y).ToString.Contains("char") Or data_type_arraylist(i)(y).ToString.Contains("text") Then
                                                command.Parameters.Add("@" + sql_format_arraylist(i)(y), SqlDbType.Char).Value = " "
                                                MsgBox("CHAR->" + excel_temp)
                                            ElseIf data_type_arraylist(i)(y).ToString.Contains("date") Or data_type_arraylist(i)(y).ToString.Contains("time") Then
                                                command.Parameters.Add("@" + sql_format_arraylist(i)(y), SqlDbType.Char).Value = New DateTime
                                                MsgBox("DATETIME->" + excel_temp)
                                            Else
                                                command.Parameters.Add("@" + sql_format_arraylist(i)(y), SqlDbType.Char).Value = 0
                                                MsgBox("NUMBER->" + excel_temp)
                                            End If

                                            'command.Parameters.Add("@" + sql_format_arraylist(i)(y), data_type_arraylist(i)(y)).Value = 0
                                            'command.Parameters.Add("@" + sql_format_arraylist(i)(y), SqlDbType.Char).Value = 0
                                        End If
                                        'command.Parameters.Add("@" + sql_format_arraylist(i)(y), SqlDbType.Char).Value = cellValue
                                        'MsgBox(excel_temp + vbTab + headerText)
                                        'MsgBox(headerText + vbTab + cellValue)
                                        'command.Parameters("@" + sql_temp).Value = cellValue
                                        queryable = True
                                    End If
                                Next
                            Next
                            'MsgBox(strtemp)
                            If queryable Then
                                'MsgBox(command.c)
                                myConn.Open()
                                command.ExecuteNonQuery()
                                myConn.Close()
                            End If
                        End Using
                    Next

                End Using
            End Using
        Next
        'For i As Integer = 0 To queryTable.Count - 1
        '    init()
        '    Using myConn = New SqlConnection("Data Source=" + serverName + ";" & "Initial Catalog=" + database + ";" + pwd_query)
        '        Using command As New SqlCommand(queryTable(i)(2), myConn)
        '            'command.CommandType = Text
        '            myConn.Open()
        '            ' Create and add the parameters, just one time here with dummy values or'
        '            ' use the full syntax to create each single the parameter'
        '            For j As Integer = 0 To sql_format_arraylist(i).count - 1
        '                'command.Parameters.AddWithValue("@" + sql_format_arraylist(i)(j), "")
        '            Next
        '            For row As Integer = 0 To dgvExcel.RowCount - 2
        '                Using r As DataGridViewRow = dgvExcel.Rows(row)
        '                    Dim queryable As Boolean
        '                    queryable = False
        '                    'Dim strtemp As String = ""
        '                    For q As Integer = 0 To r.Cells.Count - 1
        '                        Dim cellValue As Object = r.Cells(q).Value
        '                        Dim headerText As String = dgvExcel.Columns(q).HeaderText
        '                        If Not (cellValue.ToString.Trim.Equals(String.Empty)) Then
        '                            For y As Integer = 0 To excel_format_arraylist.Count - 1
        '                                Dim excel_temp = excel_format_arraylist(i)(y).ToString.Trim
        '                                If excel_temp.Equals(headerText) Then
        '                                    Dim sql_temp = sql_format_arraylist(i)(y).ToString.Trim

        '                                    command.Parameters.AddWithValue("@" + sql_format_arraylist(i)(y), cellValue.ToString)
        '                                    command.Parameters("@" + sql_temp).Value = cellValue
        '                                    queryable = True
        '                                End If
        '                            Next
        '                        End If
        '                    Next
        '                    'MsgBox(strtemp)
        '                    If queryable Then
        '                        'MsgBox(command.c)
        '                        command.ExecuteNonQuery()
        '                    End If
        '                End Using
        '            Next
        '            'For Each r As DataGridViewRow In dgvExcel.Rows
        '            '    'MsgBox("rock")
        '            '    'Dim strtemp As String = ""
        '            '    'For q As Integer = 0 To r.Cells.Count - 1
        '            '    '    Dim cellValue = r.Cells(q).Value.ToString.Trim
        '            '    '    If Not (cellValue.Equals(String.Empty)) Then
        '            '    '        strtemp += cellValue + vbTab
        '            '    '    End If
        '            '    'Next
        '            '    'If Not (strtemp.Equals(String.Empty)) Then
        '            '    '    MsgBox(strtemp)
        '            '    'End If
        '            '    'MsgBox(r.Cells.Count)
        '            '    'If (Not String.IsNullOrWhiteSpace(r.Cells(1).Value)) Then
        '            '    '    For k As Integer = 0 To excel_format_arraylist(i).count - 1
        '            '    '        command.Parameters.AddWithValue("@" + sql_format_arraylist(i)(k), "")
        '            '    '        command.Parameters("@" + sql_format_arraylist(i)(k)).Value = r.Cells(1).Value.Trim
        '            '    '    Next
        '            '    '    command.Parameters("@item").Value = r.Cells(1).Value.Trim
        '            '    '    command.Parameters("@price").Value = r.Cells(2).Value
        '            '    '    command.ExecuteNonQuery()
        '            '    'End If
        '            'Next
        '            myConn.Close()
        '        End Using
        '    End Using
        'Next
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

