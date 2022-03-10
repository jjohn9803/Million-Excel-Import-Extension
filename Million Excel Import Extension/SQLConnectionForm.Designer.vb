<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SQL_Connection_Form
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
        Me.cbDatabase = New System.Windows.Forms.ComboBox()
        Me.lblSQLStatus = New System.Windows.Forms.Label()
        Me.btnSql = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cbServerList = New System.Windows.Forms.ComboBox()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnTestConnection = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cbPasswordOption = New System.Windows.Forms.ComboBox()
        Me.txtUserId = New System.Windows.Forms.TextBox()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cbShowPassword = New System.Windows.Forms.CheckBox()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.Setting = New System.Windows.Forms.TabPage()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.TabControl1.SuspendLayout()
        Me.Setting.SuspendLayout()
        Me.SuspendLayout()
        '
        'cbDatabase
        '
        Me.cbDatabase.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbDatabase.Enabled = False
        Me.cbDatabase.FormattingEnabled = True
        Me.cbDatabase.Location = New System.Drawing.Point(88, 174)
        Me.cbDatabase.Name = "cbDatabase"
        Me.cbDatabase.Size = New System.Drawing.Size(278, 21)
        Me.cbDatabase.TabIndex = 12
        '
        'lblSQLStatus
        '
        Me.lblSQLStatus.AutoSize = True
        Me.lblSQLStatus.BackColor = System.Drawing.Color.Transparent
        Me.lblSQLStatus.ForeColor = System.Drawing.Color.DarkRed
        Me.lblSQLStatus.Location = New System.Drawing.Point(264, 136)
        Me.lblSQLStatus.Name = "lblSQLStatus"
        Me.lblSQLStatus.Size = New System.Drawing.Size(73, 13)
        Me.lblSQLStatus.TabIndex = 11
        Me.lblSQLStatus.Text = "Disconnected"
        '
        'btnSql
        '
        Me.btnSql.Location = New System.Drawing.Point(115, 131)
        Me.btnSql.Name = "btnSql"
        Me.btnSql.Size = New System.Drawing.Size(94, 23)
        Me.btnSql.TabIndex = 9
        Me.btnSql.Text = "Connect"
        Me.btnSql.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(215, 136)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(43, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Status: "
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 6)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 13)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Server Name:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(26, 177)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 13)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "Database:"
        '
        'cbServerList
        '
        Me.cbServerList.FormattingEnabled = True
        Me.cbServerList.Location = New System.Drawing.Point(84, 3)
        Me.cbServerList.Name = "cbServerList"
        Me.cbServerList.Size = New System.Drawing.Size(278, 21)
        Me.cbServerList.TabIndex = 16
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(34, 206)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 17
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(287, 206)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 18
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnTestConnection
        '
        Me.btnTestConnection.Location = New System.Drawing.Point(146, 206)
        Me.btnTestConnection.Name = "btnTestConnection"
        Me.btnTestConnection.Size = New System.Drawing.Size(109, 23)
        Me.btnTestConnection.TabIndex = 19
        Me.btnTestConnection.Text = "Test Connection"
        Me.btnTestConnection.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(19, 33)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(90, 13)
        Me.Label4.TabIndex = 21
        Me.Label4.Text = "Password Option:"
        '
        'cbPasswordOption
        '
        Me.cbPasswordOption.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbPasswordOption.FormattingEnabled = True
        Me.cbPasswordOption.Items.AddRange(New Object() {"Window Authentication", "Sql Server Authentication"})
        Me.cbPasswordOption.Location = New System.Drawing.Point(115, 30)
        Me.cbPasswordOption.Name = "cbPasswordOption"
        Me.cbPasswordOption.Size = New System.Drawing.Size(164, 21)
        Me.cbPasswordOption.TabIndex = 20
        '
        'txtUserId
        '
        Me.txtUserId.Enabled = False
        Me.txtUserId.Location = New System.Drawing.Point(115, 60)
        Me.txtUserId.Name = "txtUserId"
        Me.txtUserId.Size = New System.Drawing.Size(131, 20)
        Me.txtUserId.TabIndex = 22
        '
        'txtPassword
        '
        Me.txtPassword.Enabled = False
        Me.txtPassword.Location = New System.Drawing.Point(115, 91)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(131, 20)
        Me.txtPassword.TabIndex = 23
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(63, 60)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(46, 13)
        Me.Label5.TabIndex = 24
        Me.Label5.Text = "User ID:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(53, 94)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(56, 13)
        Me.Label6.TabIndex = 25
        Me.Label6.Text = "Password:"
        '
        'cbShowPassword
        '
        Me.cbShowPassword.AutoSize = True
        Me.cbShowPassword.Enabled = False
        Me.cbShowPassword.Location = New System.Drawing.Point(265, 93)
        Me.cbShowPassword.Name = "cbShowPassword"
        Me.cbShowPassword.Size = New System.Drawing.Size(101, 17)
        Me.cbShowPassword.TabIndex = 26
        Me.cbShowPassword.Text = "Show password"
        Me.cbShowPassword.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.Setting)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(396, 280)
        Me.TabControl1.TabIndex = 27
        '
        'Setting
        '
        Me.Setting.Controls.Add(Me.Label2)
        Me.Setting.Controls.Add(Me.cbShowPassword)
        Me.Setting.Controls.Add(Me.cbServerList)
        Me.Setting.Controls.Add(Me.Label6)
        Me.Setting.Controls.Add(Me.btnSql)
        Me.Setting.Controls.Add(Me.Label5)
        Me.Setting.Controls.Add(Me.lblSQLStatus)
        Me.Setting.Controls.Add(Me.txtPassword)
        Me.Setting.Controls.Add(Me.Label1)
        Me.Setting.Controls.Add(Me.txtUserId)
        Me.Setting.Controls.Add(Me.cbPasswordOption)
        Me.Setting.Controls.Add(Me.Label4)
        Me.Setting.Controls.Add(Me.cbDatabase)
        Me.Setting.Controls.Add(Me.Label3)
        Me.Setting.Controls.Add(Me.btnTestConnection)
        Me.Setting.Controls.Add(Me.btnSave)
        Me.Setting.Controls.Add(Me.btnClose)
        Me.Setting.Location = New System.Drawing.Point(4, 22)
        Me.Setting.Name = "Setting"
        Me.Setting.Padding = New System.Windows.Forms.Padding(3)
        Me.Setting.Size = New System.Drawing.Size(388, 254)
        Me.Setting.TabIndex = 0
        Me.Setting.Text = "Setting"
        Me.Setting.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(388, 254)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Import Excel"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'SQL_Connection_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(396, 280)
        Me.Controls.Add(Me.TabControl1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "SQL_Connection_Form"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "SQL Connection Form"
        Me.TabControl1.ResumeLayout(False)
        Me.Setting.ResumeLayout(False)
        Me.Setting.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents cbDatabase As ComboBox
    Friend WithEvents lblSQLStatus As Label
    Friend WithEvents btnSql As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents cbServerList As ComboBox
    Friend WithEvents btnSave As Button
    Friend WithEvents btnClose As Button
    Friend WithEvents btnTestConnection As Button
    Friend WithEvents Label4 As Label
    Friend WithEvents cbPasswordOption As ComboBox
    Friend WithEvents txtUserId As TextBox
    Friend WithEvents txtPassword As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents cbShowPassword As CheckBox
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents Setting As TabPage
    Friend WithEvents TabPage2 As TabPage
End Class
