﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
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
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.btnDeliveryOrder = New System.Windows.Forms.Button()
        Me.btnSalesOrder = New System.Windows.Forms.Button()
        Me.btnQuotation = New System.Windows.Forms.Button()
        Me.FlowLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.Controls.Add(Me.btnQuotation)
        Me.FlowLayoutPanel1.Controls.Add(Me.btnSalesOrder)
        Me.FlowLayoutPanel1.Controls.Add(Me.btnDeliveryOrder)
        Me.FlowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(654, 268)
        Me.FlowLayoutPanel1.TabIndex = 3
        '
        'btnDeliveryOrder
        '
        Me.btnDeliveryOrder.Font = New System.Drawing.Font("Microsoft YaHei", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDeliveryOrder.Location = New System.Drawing.Point(418, 30)
        Me.btnDeliveryOrder.Margin = New System.Windows.Forms.Padding(50, 30, 3, 30)
        Me.btnDeliveryOrder.Name = "btnDeliveryOrder"
        Me.btnDeliveryOrder.Size = New System.Drawing.Size(131, 59)
        Me.btnDeliveryOrder.TabIndex = 5
        Me.btnDeliveryOrder.Text = "Delivery Order"
        Me.btnDeliveryOrder.UseVisualStyleBackColor = True
        '
        'btnSalesOrder
        '
        Me.btnSalesOrder.Font = New System.Drawing.Font("Microsoft YaHei", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSalesOrder.Location = New System.Drawing.Point(234, 30)
        Me.btnSalesOrder.Margin = New System.Windows.Forms.Padding(50, 30, 3, 30)
        Me.btnSalesOrder.Name = "btnSalesOrder"
        Me.btnSalesOrder.Size = New System.Drawing.Size(131, 59)
        Me.btnSalesOrder.TabIndex = 4
        Me.btnSalesOrder.Text = "Sales Order"
        Me.btnSalesOrder.UseVisualStyleBackColor = True
        '
        'btnQuotation
        '
        Me.btnQuotation.Font = New System.Drawing.Font("Microsoft YaHei", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnQuotation.Location = New System.Drawing.Point(50, 30)
        Me.btnQuotation.Margin = New System.Windows.Forms.Padding(50, 30, 3, 30)
        Me.btnQuotation.Name = "btnQuotation"
        Me.btnQuotation.Size = New System.Drawing.Size(131, 59)
        Me.btnQuotation.TabIndex = 3
        Me.btnQuotation.Text = "Quotation"
        Me.btnQuotation.UseVisualStyleBackColor = True
        '
        'Function_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(654, 268)
        Me.Controls.Add(Me.FlowLayoutPanel1)
        Me.Name = "Function_Form"
        Me.Text = "Function_Form"
        Me.FlowLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents FlowLayoutPanel1 As FlowLayoutPanel
    Friend WithEvents btnDeliveryOrder As Button
    Friend WithEvents btnSalesOrder As Button
    Friend WithEvents btnQuotation As Button
End Class
