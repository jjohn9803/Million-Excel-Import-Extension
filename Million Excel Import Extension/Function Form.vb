Public Class Function_Form
    Private Sub openExcelImport(sender As Object, e As EventArgs) Handles btnQuotation.Click, btnDeliveryOrder.Click, btnSalesOrder.Click
        Dim form
        If sender.text.Equals("Quotation") Then
            form = New Quotation_Form
        Else

        End If
        SQL_Connection_Form.import_type = sender.text
        form.Text = sender.text
        form.ShowDialog()
        'Dim form As ExcelImporter = New ExcelImporter
        'SQL_Connection_Form.import_type = sender.text
        'form.Text = sender.text
        'form.ShowDialog()
    End Sub
End Class