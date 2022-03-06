<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Good = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.btnFileOpen = New System.Windows.Forms.Button()
        Me.lblFileLocation = New System.Windows.Forms.Label()
        Me.btnSql = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedHeaders
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Good})
        Me.DataGridView1.Location = New System.Drawing.Point(12, 57)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(776, 338)
        Me.DataGridView1.TabIndex = 0
        '
        'Good
        '
        Me.Good.HeaderText = "Column1"
        Me.Good.Name = "Good"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'btnFileOpen
        '
        Me.btnFileOpen.Location = New System.Drawing.Point(12, 12)
        Me.btnFileOpen.Name = "btnFileOpen"
        Me.btnFileOpen.Size = New System.Drawing.Size(117, 23)
        Me.btnFileOpen.TabIndex = 1
        Me.btnFileOpen.Text = "Import From.. (.xlsx)"
        Me.btnFileOpen.UseVisualStyleBackColor = True
        '
        'lblFileLocation
        '
        Me.lblFileLocation.AutoSize = True
        Me.lblFileLocation.Location = New System.Drawing.Point(172, 17)
        Me.lblFileLocation.Name = "lblFileLocation"
        Me.lblFileLocation.Size = New System.Drawing.Size(16, 13)
        Me.lblFileLocation.TabIndex = 2
        Me.lblFileLocation.Text = "..."
        '
        'btnSql
        '
        Me.btnSql.Location = New System.Drawing.Point(251, 11)
        Me.btnSql.Name = "btnSql"
        Me.btnSql.Size = New System.Drawing.Size(147, 23)
        Me.btnSql.TabIndex = 3
        Me.btnSql.Text = "Get From SQL"
        Me.btnSql.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.btnSql)
        Me.Controls.Add(Me.lblFileLocation)
        Me.Controls.Add(Me.btnFileOpen)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents btnFileOpen As Button
    Friend WithEvents lblFileLocation As Label
    Friend WithEvents Good As DataGridViewTextBoxColumn
    Friend WithEvents btnSql As Button
End Class
