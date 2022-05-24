<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Function_Form
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.FlowLayoutPanel2 = New System.Windows.Forms.FlowLayoutPanel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnStockReceive = New System.Windows.Forms.Button()
        Me.btnStockTrasfer = New System.Windows.Forms.Button()
        Me.btnStockAdjustment = New System.Windows.Forms.Button()
        Me.btnStockIssue = New System.Windows.Forms.Button()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.btnQuotation = New System.Windows.Forms.Button()
        Me.btnSalesOrder = New System.Windows.Forms.Button()
        Me.btnDeliveryOrder = New System.Windows.Forms.Button()
        Me.btnSalesInvoice = New System.Windows.Forms.Button()
        Me.btnCashSales = New System.Windows.Forms.Button()
        Me.btnDebitNote = New System.Windows.Forms.Button()
        Me.btnCreditNote = New System.Windows.Forms.Button()
        Me.btnDeliveryReturn = New System.Windows.Forms.Button()
        Me.FlowLayoutPanel2.SuspendLayout()
        Me.FlowLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'FlowLayoutPanel2
        '
        Me.FlowLayoutPanel2.Controls.Add(Me.Label1)
        Me.FlowLayoutPanel2.Controls.Add(Me.btnStockReceive)
        Me.FlowLayoutPanel2.Controls.Add(Me.btnStockIssue)
        Me.FlowLayoutPanel2.Controls.Add(Me.btnStockAdjustment)
        Me.FlowLayoutPanel2.Controls.Add(Me.btnStockTrasfer)
        Me.FlowLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.FlowLayoutPanel2.Location = New System.Drawing.Point(0, 286)
        Me.FlowLayoutPanel2.Margin = New System.Windows.Forms.Padding(4)
        Me.FlowLayoutPanel2.Name = "FlowLayoutPanel2"
        Me.FlowLayoutPanel2.Size = New System.Drawing.Size(1014, 186)
        Me.FlowLayoutPanel2.TabIndex = 16
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft YaHei", 13.8!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(50, 0)
        Me.Label1.Margin = New System.Windows.Forms.Padding(50, 0, 1000, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(79, 31)
        Me.Label1.TabIndex = 15
        Me.Label1.Text = "Stock"
        '
        'btnStockReceive
        '
        Me.btnStockReceive.Font = New System.Drawing.Font("Microsoft YaHei", 12.0!)
        Me.btnStockReceive.Location = New System.Drawing.Point(67, 68)
        Me.btnStockReceive.Margin = New System.Windows.Forms.Padding(67, 37, 4, 37)
        Me.btnStockReceive.Name = "btnStockReceive"
        Me.btnStockReceive.Size = New System.Drawing.Size(175, 73)
        Me.btnStockReceive.TabIndex = 11
        Me.btnStockReceive.TabStop = False
        Me.btnStockReceive.Text = "Stock Receive"
        Me.btnStockReceive.UseVisualStyleBackColor = True
        '
        'btnStockTrasfer
        '
        Me.btnStockTrasfer.Font = New System.Drawing.Font("Microsoft YaHei", 12.0!)
        Me.btnStockTrasfer.Location = New System.Drawing.Point(805, 68)
        Me.btnStockTrasfer.Margin = New System.Windows.Forms.Padding(67, 37, 4, 37)
        Me.btnStockTrasfer.Name = "btnStockTrasfer"
        Me.btnStockTrasfer.Size = New System.Drawing.Size(175, 73)
        Me.btnStockTrasfer.TabIndex = 14
        Me.btnStockTrasfer.TabStop = False
        Me.btnStockTrasfer.Text = "Stock Transfer"
        Me.btnStockTrasfer.UseVisualStyleBackColor = True
        '
        'btnStockAdjustment
        '
        Me.btnStockAdjustment.Font = New System.Drawing.Font("Microsoft YaHei", 12.0!)
        Me.btnStockAdjustment.Location = New System.Drawing.Point(559, 68)
        Me.btnStockAdjustment.Margin = New System.Windows.Forms.Padding(67, 37, 4, 37)
        Me.btnStockAdjustment.Name = "btnStockAdjustment"
        Me.btnStockAdjustment.Size = New System.Drawing.Size(175, 73)
        Me.btnStockAdjustment.TabIndex = 13
        Me.btnStockAdjustment.TabStop = False
        Me.btnStockAdjustment.Text = "Stock Adjustment"
        Me.btnStockAdjustment.UseVisualStyleBackColor = True
        '
        'btnStockIssue
        '
        Me.btnStockIssue.Font = New System.Drawing.Font("Microsoft YaHei", 12.0!)
        Me.btnStockIssue.Location = New System.Drawing.Point(313, 68)
        Me.btnStockIssue.Margin = New System.Windows.Forms.Padding(67, 37, 4, 37)
        Me.btnStockIssue.Name = "btnStockIssue"
        Me.btnStockIssue.Size = New System.Drawing.Size(175, 73)
        Me.btnStockIssue.TabIndex = 12
        Me.btnStockIssue.TabStop = False
        Me.btnStockIssue.Text = "Stock Issue"
        Me.btnStockIssue.UseVisualStyleBackColor = True
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.Controls.Add(Me.btnQuotation)
        Me.FlowLayoutPanel1.Controls.Add(Me.btnSalesOrder)
        Me.FlowLayoutPanel1.Controls.Add(Me.btnDeliveryOrder)
        Me.FlowLayoutPanel1.Controls.Add(Me.btnSalesInvoice)
        Me.FlowLayoutPanel1.Controls.Add(Me.btnCashSales)
        Me.FlowLayoutPanel1.Controls.Add(Me.btnDebitNote)
        Me.FlowLayoutPanel1.Controls.Add(Me.btnCreditNote)
        Me.FlowLayoutPanel1.Controls.Add(Me.btnDeliveryReturn)
        Me.FlowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.FlowLayoutPanel1.Margin = New System.Windows.Forms.Padding(4)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(1014, 286)
        Me.FlowLayoutPanel1.TabIndex = 17
        '
        'btnQuotation
        '
        Me.btnQuotation.Font = New System.Drawing.Font("Microsoft YaHei", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnQuotation.Location = New System.Drawing.Point(67, 37)
        Me.btnQuotation.Margin = New System.Windows.Forms.Padding(67, 37, 4, 37)
        Me.btnQuotation.Name = "btnQuotation"
        Me.btnQuotation.Size = New System.Drawing.Size(175, 73)
        Me.btnQuotation.TabIndex = 3
        Me.btnQuotation.TabStop = False
        Me.btnQuotation.Text = "Quotation"
        Me.btnQuotation.UseVisualStyleBackColor = True
        '
        'btnSalesOrder
        '
        Me.btnSalesOrder.Font = New System.Drawing.Font("Microsoft YaHei", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSalesOrder.Location = New System.Drawing.Point(313, 37)
        Me.btnSalesOrder.Margin = New System.Windows.Forms.Padding(67, 37, 4, 37)
        Me.btnSalesOrder.Name = "btnSalesOrder"
        Me.btnSalesOrder.Size = New System.Drawing.Size(175, 73)
        Me.btnSalesOrder.TabIndex = 4
        Me.btnSalesOrder.TabStop = False
        Me.btnSalesOrder.Text = "Sales Order"
        Me.btnSalesOrder.UseVisualStyleBackColor = True
        '
        'btnDeliveryOrder
        '
        Me.btnDeliveryOrder.Font = New System.Drawing.Font("Microsoft YaHei", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDeliveryOrder.Location = New System.Drawing.Point(559, 37)
        Me.btnDeliveryOrder.Margin = New System.Windows.Forms.Padding(67, 37, 4, 37)
        Me.btnDeliveryOrder.Name = "btnDeliveryOrder"
        Me.btnDeliveryOrder.Size = New System.Drawing.Size(175, 73)
        Me.btnDeliveryOrder.TabIndex = 5
        Me.btnDeliveryOrder.TabStop = False
        Me.btnDeliveryOrder.Text = "Delivery Order"
        Me.btnDeliveryOrder.UseVisualStyleBackColor = True
        '
        'btnSalesInvoice
        '
        Me.btnSalesInvoice.Font = New System.Drawing.Font("Microsoft YaHei", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSalesInvoice.Location = New System.Drawing.Point(805, 37)
        Me.btnSalesInvoice.Margin = New System.Windows.Forms.Padding(67, 37, 4, 37)
        Me.btnSalesInvoice.Name = "btnSalesInvoice"
        Me.btnSalesInvoice.Size = New System.Drawing.Size(175, 73)
        Me.btnSalesInvoice.TabIndex = 6
        Me.btnSalesInvoice.TabStop = False
        Me.btnSalesInvoice.Text = "Sales Invoice"
        Me.btnSalesInvoice.UseVisualStyleBackColor = True
        '
        'btnCashSales
        '
        Me.btnCashSales.Font = New System.Drawing.Font("Microsoft YaHei", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCashSales.Location = New System.Drawing.Point(67, 184)
        Me.btnCashSales.Margin = New System.Windows.Forms.Padding(67, 37, 4, 37)
        Me.btnCashSales.Name = "btnCashSales"
        Me.btnCashSales.Size = New System.Drawing.Size(175, 73)
        Me.btnCashSales.TabIndex = 7
        Me.btnCashSales.TabStop = False
        Me.btnCashSales.Text = "Cash Sales"
        Me.btnCashSales.UseVisualStyleBackColor = True
        '
        'btnDebitNote
        '
        Me.btnDebitNote.Font = New System.Drawing.Font("Microsoft YaHei", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDebitNote.Location = New System.Drawing.Point(313, 184)
        Me.btnDebitNote.Margin = New System.Windows.Forms.Padding(67, 37, 4, 37)
        Me.btnDebitNote.Name = "btnDebitNote"
        Me.btnDebitNote.Size = New System.Drawing.Size(175, 73)
        Me.btnDebitNote.TabIndex = 8
        Me.btnDebitNote.TabStop = False
        Me.btnDebitNote.Text = "Debit Note"
        Me.btnDebitNote.UseVisualStyleBackColor = True
        '
        'btnCreditNote
        '
        Me.btnCreditNote.Font = New System.Drawing.Font("Microsoft YaHei", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCreditNote.Location = New System.Drawing.Point(559, 184)
        Me.btnCreditNote.Margin = New System.Windows.Forms.Padding(67, 37, 4, 37)
        Me.btnCreditNote.Name = "btnCreditNote"
        Me.btnCreditNote.Size = New System.Drawing.Size(175, 73)
        Me.btnCreditNote.TabIndex = 9
        Me.btnCreditNote.TabStop = False
        Me.btnCreditNote.Text = "Credit Note"
        Me.btnCreditNote.UseVisualStyleBackColor = True
        '
        'btnDeliveryReturn
        '
        Me.btnDeliveryReturn.Font = New System.Drawing.Font("Microsoft YaHei", 12.0!)
        Me.btnDeliveryReturn.Location = New System.Drawing.Point(805, 184)
        Me.btnDeliveryReturn.Margin = New System.Windows.Forms.Padding(67, 37, 4, 37)
        Me.btnDeliveryReturn.Name = "btnDeliveryReturn"
        Me.btnDeliveryReturn.Size = New System.Drawing.Size(175, 73)
        Me.btnDeliveryReturn.TabIndex = 10
        Me.btnDeliveryReturn.TabStop = False
        Me.btnDeliveryReturn.Text = "Delivery Return"
        Me.btnDeliveryReturn.UseVisualStyleBackColor = True
        '
        'Function_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1014, 472)
        Me.Controls.Add(Me.FlowLayoutPanel1)
        Me.Controls.Add(Me.FlowLayoutPanel2)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "Function_Form"
        Me.Text = "Function_Form"
        Me.FlowLayoutPanel2.ResumeLayout(False)
        Me.FlowLayoutPanel2.PerformLayout()
        Me.FlowLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents FlowLayoutPanel2 As FlowLayoutPanel
    Friend WithEvents Label1 As Label
    Friend WithEvents btnStockReceive As Button
    Friend WithEvents btnStockTrasfer As Button
    Friend WithEvents btnStockAdjustment As Button
    Friend WithEvents btnStockIssue As Button
    Friend WithEvents FlowLayoutPanel1 As FlowLayoutPanel
    Friend WithEvents btnQuotation As Button
    Friend WithEvents btnSalesOrder As Button
    Friend WithEvents btnDeliveryOrder As Button
    Friend WithEvents btnSalesInvoice As Button
    Friend WithEvents btnCashSales As Button
    Friend WithEvents btnDebitNote As Button
    Friend WithEvents btnCreditNote As Button
    Friend WithEvents btnDeliveryReturn As Button
End Class
