<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Good = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.btnFileOpen = New System.Windows.Forms.Button()
        Me.lblFileLocation = New System.Windows.Forms.Label()
        Me.btnSql = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblSQLStatus = New System.Windows.Forms.Label()
        Me.DataGridViewTables = New System.Windows.Forms.DataGridView()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridViewTables, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedHeaders
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Good})
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(0, 0)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(400, 350)
        Me.DataGridView1.TabIndex = 0
        '
        'Good
        '
        Me.Good.HeaderText = "Column1"
        Me.Good.Name = "Good"
        Me.Good.Width = 327
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'btnFileOpen
        '
        Me.btnFileOpen.Location = New System.Drawing.Point(3, 3)
        Me.btnFileOpen.Name = "btnFileOpen"
        Me.btnFileOpen.Size = New System.Drawing.Size(117, 23)
        Me.btnFileOpen.TabIndex = 1
        Me.btnFileOpen.Text = "Import From.. (.xlsx)"
        Me.btnFileOpen.UseVisualStyleBackColor = True
        '
        'lblFileLocation
        '
        Me.lblFileLocation.AutoSize = True
        Me.lblFileLocation.Location = New System.Drawing.Point(126, 8)
        Me.lblFileLocation.Name = "lblFileLocation"
        Me.lblFileLocation.Size = New System.Drawing.Size(16, 13)
        Me.lblFileLocation.TabIndex = 2
        Me.lblFileLocation.Text = "..."
        '
        'btnSql
        '
        Me.btnSql.Location = New System.Drawing.Point(485, 3)
        Me.btnSql.Name = "btnSql"
        Me.btnSql.Size = New System.Drawing.Size(147, 23)
        Me.btnSql.TabIndex = 3
        Me.btnSql.Text = "Connect to SQL Server"
        Me.btnSql.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(638, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(43, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Status: "
        '
        'lblSQLStatus
        '
        Me.lblSQLStatus.AutoSize = True
        Me.lblSQLStatus.BackColor = System.Drawing.Color.Transparent
        Me.lblSQLStatus.ForeColor = System.Drawing.Color.DarkRed
        Me.lblSQLStatus.Location = New System.Drawing.Point(687, 8)
        Me.lblSQLStatus.Name = "lblSQLStatus"
        Me.lblSQLStatus.Size = New System.Drawing.Size(73, 13)
        Me.lblSQLStatus.TabIndex = 5
        Me.lblSQLStatus.Text = "Disconnected"
        '
        'DataGridViewTables
        '
        Me.DataGridViewTables.AllowUserToAddRows = False
        Me.DataGridViewTables.AllowUserToDeleteRows = False
        Me.DataGridViewTables.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells
        Me.DataGridViewTables.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridViewTables.Location = New System.Drawing.Point(0, 0)
        Me.DataGridViewTables.Name = "DataGridViewTables"
        Me.DataGridViewTables.ReadOnly = True
        Me.DataGridViewTables.Size = New System.Drawing.Size(398, 350)
        Me.DataGridViewTables.TabIndex = 6
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnFileOpen)
        Me.Panel1.Controls.Add(Me.lblFileLocation)
        Me.Panel1.Controls.Add(Me.lblSQLStatus)
        Me.Panel1.Controls.Add(Me.btnSql)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(800, 100)
        Me.Panel1.TabIndex = 7
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.DataGridViewTables)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel2.Location = New System.Drawing.Point(402, 100)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(398, 350)
        Me.Panel2.TabIndex = 8
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.DataGridView1)
        Me.Panel3.Location = New System.Drawing.Point(0, 100)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(400, 350)
        Me.Panel3.TabIndex = 9
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "Form1"
        Me.Text = "Million Excel Import Extension"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridViewTables, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents btnFileOpen As Button
    Friend WithEvents lblFileLocation As Label
    Friend WithEvents Good As DataGridViewTextBoxColumn
    Friend WithEvents btnSql As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents lblSQLStatus As Label
    Friend WithEvents DataGridViewTables As DataGridView
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Panel3 As Panel
End Class
