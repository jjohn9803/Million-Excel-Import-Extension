<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Setting
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Setting))
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbShowPassword = New System.Windows.Forms.CheckBox()
        Me.cbServerList = New System.Windows.Forms.ComboBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.btnSql = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.txtUserId = New System.Windows.Forms.TextBox()
        Me.cbPasswordOption = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cbDatabase = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnTestConnection = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnSearchServer = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Ebrima", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(53, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(104, 21)
        Me.Label2.TabIndex = 29
        Me.Label2.Text = "Server Name:"
        '
        'cbShowPassword
        '
        Me.cbShowPassword.AutoSize = True
        Me.cbShowPassword.Enabled = False
        Me.cbShowPassword.Font = New System.Drawing.Font("Ebrima", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbShowPassword.Location = New System.Drawing.Point(389, 113)
        Me.cbShowPassword.Name = "cbShowPassword"
        Me.cbShowPassword.Size = New System.Drawing.Size(139, 25)
        Me.cbShowPassword.TabIndex = 5
        Me.cbShowPassword.Text = "Show password"
        Me.cbShowPassword.UseVisualStyleBackColor = True
        '
        'cbServerList
        '
        Me.cbServerList.Font = New System.Drawing.Font("Ebrima", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbServerList.FormattingEnabled = True
        Me.cbServerList.Location = New System.Drawing.Point(163, 6)
        Me.cbServerList.Name = "cbServerList"
        Me.cbServerList.Size = New System.Drawing.Size(278, 29)
        Me.cbServerList.TabIndex = 0
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Ebrima", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(77, 114)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(80, 21)
        Me.Label6.TabIndex = 40
        Me.Label6.Text = "Password:"
        '
        'btnSql
        '
        Me.btnSql.Font = New System.Drawing.Font("Ebrima", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSql.Location = New System.Drawing.Point(197, 147)
        Me.btnSql.Name = "btnSql"
        Me.btnSql.Size = New System.Drawing.Size(150, 35)
        Me.btnSql.TabIndex = 6
        Me.btnSql.Text = "Connect"
        Me.btnSql.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Ebrima", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(91, 79)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(64, 21)
        Me.Label5.TabIndex = 39
        Me.Label5.Text = "User ID:"
        '
        'txtPassword
        '
        Me.txtPassword.Enabled = False
        Me.txtPassword.Font = New System.Drawing.Font("Ebrima", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPassword.Location = New System.Drawing.Point(163, 111)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(220, 29)
        Me.txtPassword.TabIndex = 4
        '
        'txtUserId
        '
        Me.txtUserId.Enabled = False
        Me.txtUserId.Font = New System.Drawing.Font("Ebrima", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUserId.Location = New System.Drawing.Point(163, 76)
        Me.txtUserId.Name = "txtUserId"
        Me.txtUserId.Size = New System.Drawing.Size(220, 29)
        Me.txtUserId.TabIndex = 3
        '
        'cbPasswordOption
        '
        Me.cbPasswordOption.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbPasswordOption.Font = New System.Drawing.Font("Ebrima", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbPasswordOption.FormattingEnabled = True
        Me.cbPasswordOption.Items.AddRange(New Object() {"Window Authentication", "Sql Server Authentication"})
        Me.cbPasswordOption.Location = New System.Drawing.Point(163, 41)
        Me.cbPasswordOption.Name = "cbPasswordOption"
        Me.cbPasswordOption.Size = New System.Drawing.Size(278, 29)
        Me.cbPasswordOption.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Ebrima", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(25, 44)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(132, 21)
        Me.Label4.TabIndex = 36
        Me.Label4.Text = "Password Option:"
        '
        'cbDatabase
        '
        Me.cbDatabase.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbDatabase.Enabled = False
        Me.cbDatabase.Font = New System.Drawing.Font("Ebrima", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbDatabase.FormattingEnabled = True
        Me.cbDatabase.Location = New System.Drawing.Point(163, 188)
        Me.cbDatabase.Name = "cbDatabase"
        Me.cbDatabase.Size = New System.Drawing.Size(278, 29)
        Me.cbDatabase.TabIndex = 7
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Ebrima", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(78, 188)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(77, 21)
        Me.Label3.TabIndex = 30
        Me.Label3.Text = "Database:"
        '
        'btnTestConnection
        '
        Me.btnTestConnection.Font = New System.Drawing.Font("Ebrima", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTestConnection.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnTestConnection.Location = New System.Drawing.Point(41, 233)
        Me.btnTestConnection.Name = "btnTestConnection"
        Me.btnTestConnection.Size = New System.Drawing.Size(150, 40)
        Me.btnTestConnection.TabIndex = 8
        Me.btnTestConnection.Text = "Test Connection"
        Me.btnTestConnection.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Font = New System.Drawing.Font("Ebrima", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Location = New System.Drawing.Point(197, 233)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(150, 40)
        Me.btnSave.TabIndex = 9
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Font = New System.Drawing.Font("Ebrima", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(353, 233)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(150, 40)
        Me.btnClose.TabIndex = 10
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnSearchServer
        '
        Me.btnSearchServer.Location = New System.Drawing.Point(447, 6)
        Me.btnSearchServer.Name = "btnSearchServer"
        Me.btnSearchServer.Size = New System.Drawing.Size(32, 29)
        Me.btnSearchServer.TabIndex = 1
        Me.btnSearchServer.Text = "..."
        Me.btnSearchServer.UseVisualStyleBackColor = True
        '
        'Setting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(545, 291)
        Me.Controls.Add(Me.btnSearchServer)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cbShowPassword)
        Me.Controls.Add(Me.cbServerList)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.btnSql)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.txtUserId)
        Me.Controls.Add(Me.cbPasswordOption)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cbDatabase)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnTestConnection)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnClose)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "Setting"
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Setting"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label2 As Label
    Friend WithEvents cbShowPassword As CheckBox
    Friend WithEvents cbServerList As ComboBox
    Friend WithEvents Label6 As Label
    Friend WithEvents btnSql As Button
    Friend WithEvents Label5 As Label
    Friend WithEvents txtPassword As TextBox
    Friend WithEvents txtUserId As TextBox
    Friend WithEvents cbPasswordOption As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents cbDatabase As ComboBox
    Friend WithEvents Label3 As Label
    Friend WithEvents btnTestConnection As Button
    Friend WithEvents btnSave As Button
    Friend WithEvents btnClose As Button
    Friend WithEvents btnSearchServer As Button
End Class
