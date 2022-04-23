<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Sales_Order_Form
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
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Sales_Order_Form))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtFileName = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cbSheet = New System.Windows.Forms.ComboBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.FlowLayoutPanel2 = New System.Windows.Forms.FlowLayoutPanel()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtType = New System.Windows.Forms.TextBox()
        Me.btnImport = New System.Windows.Forms.Button()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.dgvExcel = New System.Windows.Forms.DataGridView()
        Me.Panel1.SuspendLayout()
        Me.FlowLayoutPanel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.FlowLayoutPanel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        CType(Me.dgvExcel, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.FlowLayoutPanel1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(884, 73)
        Me.Panel1.TabIndex = 17
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.Controls.Add(Me.Label2)
        Me.FlowLayoutPanel1.Controls.Add(Me.txtFileName)
        Me.FlowLayoutPanel1.Controls.Add(Me.Label1)
        Me.FlowLayoutPanel1.Controls.Add(Me.cbSheet)
        Me.FlowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(884, 73)
        Me.FlowLayoutPanel1.TabIndex = 14
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Ebrima", 11.0!)
        Me.Label2.Location = New System.Drawing.Point(30, 15)
        Me.Label2.Margin = New System.Windows.Forms.Padding(30, 15, 3, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Padding = New System.Windows.Forms.Padding(0, 5, 0, 5)
        Me.Label2.Size = New System.Drawing.Size(72, 30)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Filename:"
        '
        'txtFileName
        '
        Me.txtFileName.Font = New System.Drawing.Font("Ebrima", 11.0!)
        Me.txtFileName.Location = New System.Drawing.Point(108, 15)
        Me.txtFileName.Margin = New System.Windows.Forms.Padding(3, 15, 3, 3)
        Me.txtFileName.Name = "txtFileName"
        Me.txtFileName.Size = New System.Drawing.Size(387, 27)
        Me.txtFileName.TabIndex = 9
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Ebrima", 11.0!)
        Me.Label1.Location = New System.Drawing.Point(528, 15)
        Me.Label1.Margin = New System.Windows.Forms.Padding(30, 15, 3, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Padding = New System.Windows.Forms.Padding(0, 5, 0, 5)
        Me.Label1.Size = New System.Drawing.Size(49, 30)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Sheet:"
        '
        'cbSheet
        '
        Me.cbSheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbSheet.Font = New System.Drawing.Font("Ebrima", 11.0!)
        Me.cbSheet.FormattingEnabled = True
        Me.cbSheet.Location = New System.Drawing.Point(583, 15)
        Me.cbSheet.Margin = New System.Windows.Forms.Padding(3, 15, 3, 3)
        Me.cbSheet.Name = "cbSheet"
        Me.cbSheet.Size = New System.Drawing.Size(261, 28)
        Me.cbSheet.TabIndex = 6
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.FlowLayoutPanel2)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel2.Location = New System.Drawing.Point(0, 401)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(884, 60)
        Me.Panel2.TabIndex = 18
        '
        'FlowLayoutPanel2
        '
        Me.FlowLayoutPanel2.Controls.Add(Me.Label4)
        Me.FlowLayoutPanel2.Controls.Add(Me.txtType)
        Me.FlowLayoutPanel2.Controls.Add(Me.btnImport)
        Me.FlowLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FlowLayoutPanel2.Location = New System.Drawing.Point(0, 0)
        Me.FlowLayoutPanel2.Name = "FlowLayoutPanel2"
        Me.FlowLayoutPanel2.Size = New System.Drawing.Size(884, 60)
        Me.FlowLayoutPanel2.TabIndex = 18
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Ebrima", 11.0!)
        Me.Label4.Location = New System.Drawing.Point(50, 15)
        Me.Label4.Margin = New System.Windows.Forms.Padding(50, 15, 3, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Padding = New System.Windows.Forms.Padding(0, 5, 0, 5)
        Me.Label4.Size = New System.Drawing.Size(48, 30)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "Type: "
        '
        'txtType
        '
        Me.txtType.Font = New System.Drawing.Font("Ebrima", 11.0!)
        Me.txtType.Location = New System.Drawing.Point(111, 17)
        Me.txtType.Margin = New System.Windows.Forms.Padding(10, 17, 3, 0)
        Me.txtType.Name = "txtType"
        Me.txtType.ReadOnly = True
        Me.txtType.Size = New System.Drawing.Size(275, 27)
        Me.txtType.TabIndex = 23
        '
        'btnImport
        '
        Me.btnImport.Enabled = False
        Me.btnImport.Font = New System.Drawing.Font("Ebrima", 11.0!)
        Me.btnImport.Location = New System.Drawing.Point(539, 15)
        Me.btnImport.Margin = New System.Windows.Forms.Padding(150, 15, 3, 3)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(175, 33)
        Me.btnImport.TabIndex = 18
        Me.btnImport.Text = "Import"
        Me.btnImport.UseVisualStyleBackColor = True
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.dgvExcel)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(0, 73)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(884, 328)
        Me.Panel3.TabIndex = 20
        '
        'dgvExcel
        '
        Me.dgvExcel.AllowUserToAddRows = False
        Me.dgvExcel.AllowUserToDeleteRows = False
        DataGridViewCellStyle4.NullValue = Nothing
        Me.dgvExcel.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle4
        Me.dgvExcel.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells
        Me.dgvExcel.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.ControlLight
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Ebrima", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvExcel.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle5
        Me.dgvExcel.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Ebrima", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle6.NullValue = Nothing
        DataGridViewCellStyle6.Padding = New System.Windows.Forms.Padding(0, 3, 0, 3)
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvExcel.DefaultCellStyle = DataGridViewCellStyle6
        Me.dgvExcel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvExcel.Location = New System.Drawing.Point(0, 0)
        Me.dgvExcel.Name = "dgvExcel"
        Me.dgvExcel.ReadOnly = True
        Me.dgvExcel.Size = New System.Drawing.Size(884, 328)
        Me.dgvExcel.TabIndex = 2
        '
        'Sales_Order_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(884, 461)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel2)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Sales_Order_Form"
        Me.Text = "Sales Order"
        Me.Panel1.ResumeLayout(False)
        Me.FlowLayoutPanel1.ResumeLayout(False)
        Me.FlowLayoutPanel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.FlowLayoutPanel2.ResumeLayout(False)
        Me.FlowLayoutPanel2.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        CType(Me.dgvExcel, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents FlowLayoutPanel1 As FlowLayoutPanel
    Friend WithEvents Label2 As Label
    Friend WithEvents txtFileName As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents cbSheet As ComboBox
    Friend WithEvents Panel2 As Panel
    Friend WithEvents FlowLayoutPanel2 As FlowLayoutPanel
    Friend WithEvents Label4 As Label
    Friend WithEvents txtType As TextBox
    Friend WithEvents btnImport As Button
    Friend WithEvents Panel3 As Panel
    Friend WithEvents dgvExcel As DataGridView
End Class
