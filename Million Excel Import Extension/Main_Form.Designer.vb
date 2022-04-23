<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Main_Form
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Main_Form))
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtServerName = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtDatabaseName = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblSQLStatus = New System.Windows.Forms.Label()
        Me.btnSetting = New System.Windows.Forms.Button()
        Me.panelMain = New System.Windows.Forms.Panel()
        Me.FlowLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.BackColor = System.Drawing.SystemColors.ControlLight
        Me.FlowLayoutPanel1.Controls.Add(Me.Label7)
        Me.FlowLayoutPanel1.Controls.Add(Me.txtServerName)
        Me.FlowLayoutPanel1.Controls.Add(Me.Label2)
        Me.FlowLayoutPanel1.Controls.Add(Me.txtDatabaseName)
        Me.FlowLayoutPanel1.Controls.Add(Me.Label1)
        Me.FlowLayoutPanel1.Controls.Add(Me.lblSQLStatus)
        Me.FlowLayoutPanel1.Controls.Add(Me.btnSetting)
        Me.FlowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(0, 304)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(594, 102)
        Me.FlowLayoutPanel1.TabIndex = 34
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Yu Gothic Medium", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(10, 15)
        Me.Label7.Margin = New System.Windows.Forms.Padding(10, 15, 3, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(106, 17)
        Me.Label7.TabIndex = 30
        Me.Label7.Text = "Server Name: "
        '
        'txtServerName
        '
        Me.txtServerName.Location = New System.Drawing.Point(122, 12)
        Me.txtServerName.Margin = New System.Windows.Forms.Padding(3, 12, 3, 3)
        Me.txtServerName.Name = "txtServerName"
        Me.txtServerName.ReadOnly = True
        Me.txtServerName.Size = New System.Drawing.Size(149, 20)
        Me.txtServerName.TabIndex = 0
        Me.txtServerName.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Yu Gothic Medium", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(294, 15)
        Me.Label2.Margin = New System.Windows.Forms.Padding(20, 15, 3, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(128, 17)
        Me.Label2.TabIndex = 35
        Me.Label2.Text = "Database Name: "
        '
        'txtDatabaseName
        '
        Me.txtDatabaseName.Location = New System.Drawing.Point(428, 12)
        Me.txtDatabaseName.Margin = New System.Windows.Forms.Padding(3, 12, 3, 3)
        Me.txtDatabaseName.Name = "txtDatabaseName"
        Me.txtDatabaseName.ReadOnly = True
        Me.txtDatabaseName.Size = New System.Drawing.Size(149, 20)
        Me.txtDatabaseName.TabIndex = 36
        Me.txtDatabaseName.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Yu Gothic Medium", 9.75!, System.Drawing.FontStyle.Bold)
        Me.Label1.Location = New System.Drawing.Point(50, 45)
        Me.Label1.Margin = New System.Windows.Forms.Padding(50, 10, 3, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 17)
        Me.Label1.TabIndex = 37
        Me.Label1.Text = "Status: "
        '
        'lblSQLStatus
        '
        Me.lblSQLStatus.AutoSize = True
        Me.lblSQLStatus.BackColor = System.Drawing.Color.Transparent
        Me.lblSQLStatus.Font = New System.Drawing.Font("Yu Gothic Medium", 9.75!, System.Drawing.FontStyle.Bold)
        Me.lblSQLStatus.ForeColor = System.Drawing.Color.DarkRed
        Me.lblSQLStatus.Location = New System.Drawing.Point(125, 45)
        Me.lblSQLStatus.Margin = New System.Windows.Forms.Padding(8, 10, 3, 0)
        Me.lblSQLStatus.Name = "lblSQLStatus"
        Me.lblSQLStatus.Size = New System.Drawing.Size(104, 17)
        Me.lblSQLStatus.TabIndex = 38
        Me.lblSQLStatus.Text = "Disconnected"
        '
        'btnSetting
        '
        Me.btnSetting.Font = New System.Drawing.Font("Yu Gothic Medium", 9.75!, System.Drawing.FontStyle.Bold)
        Me.btnSetting.Location = New System.Drawing.Point(372, 42)
        Me.btnSetting.Margin = New System.Windows.Forms.Padding(140, 7, 3, 3)
        Me.btnSetting.Name = "btnSetting"
        Me.btnSetting.Size = New System.Drawing.Size(120, 30)
        Me.btnSetting.TabIndex = 39
        Me.btnSetting.TabStop = False
        Me.btnSetting.Text = "Setting"
        Me.btnSetting.UseVisualStyleBackColor = True
        '
        'panelMain
        '
        Me.panelMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.panelMain.Location = New System.Drawing.Point(0, 0)
        Me.panelMain.Name = "panelMain"
        Me.panelMain.Size = New System.Drawing.Size(594, 304)
        Me.panelMain.TabIndex = 35
        '
        'Main_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(594, 406)
        Me.Controls.Add(Me.panelMain)
        Me.Controls.Add(Me.FlowLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "Main_Form"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Million Excel Import Extension"
        Me.FlowLayoutPanel1.ResumeLayout(False)
        Me.FlowLayoutPanel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents FlowLayoutPanel1 As FlowLayoutPanel
    Friend WithEvents Label7 As Label
    Friend WithEvents txtServerName As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents txtDatabaseName As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents lblSQLStatus As Label
    Friend WithEvents panelMain As Panel
    Friend WithEvents btnSetting As Button
End Class
