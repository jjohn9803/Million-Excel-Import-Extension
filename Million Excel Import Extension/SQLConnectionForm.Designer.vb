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
        Me.btnSql = New System.Windows.Forms.Button()
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
        Me.TabPageSetting = New System.Windows.Forms.TabPage()
        Me.TabPageExcel = New System.Windows.Forms.TabPage()
        Me.TabPageMaintain = New System.Windows.Forms.TabPage()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblSQLStatus = New System.Windows.Forms.Label()
        Me.TabControl1.SuspendLayout()
        Me.TabPageSetting.SuspendLayout()
        Me.Panel1.SuspendLayout()
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
        'btnSql
        '
        Me.btnSql.Location = New System.Drawing.Point(115, 131)
        Me.btnSql.Name = "btnSql"
        Me.btnSql.Size = New System.Drawing.Size(94, 23)
        Me.btnSql.TabIndex = 9
        Me.btnSql.Text = "Connect"
        Me.btnSql.UseVisualStyleBackColor = True
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
        Me.TabControl1.Controls.Add(Me.TabPageSetting)
        Me.TabControl1.Controls.Add(Me.TabPageExcel)
        Me.TabControl1.Controls.Add(Me.TabPageMaintain)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(396, 284)
        Me.TabControl1.TabIndex = 27
        '
        'TabPageSetting
        '
        Me.TabPageSetting.Controls.Add(Me.Label2)
        Me.TabPageSetting.Controls.Add(Me.cbShowPassword)
        Me.TabPageSetting.Controls.Add(Me.cbServerList)
        Me.TabPageSetting.Controls.Add(Me.Label6)
        Me.TabPageSetting.Controls.Add(Me.btnSql)
        Me.TabPageSetting.Controls.Add(Me.Label5)
        Me.TabPageSetting.Controls.Add(Me.txtPassword)
        Me.TabPageSetting.Controls.Add(Me.txtUserId)
        Me.TabPageSetting.Controls.Add(Me.cbPasswordOption)
        Me.TabPageSetting.Controls.Add(Me.Label4)
        Me.TabPageSetting.Controls.Add(Me.cbDatabase)
        Me.TabPageSetting.Controls.Add(Me.Label3)
        Me.TabPageSetting.Controls.Add(Me.btnTestConnection)
        Me.TabPageSetting.Controls.Add(Me.btnSave)
        Me.TabPageSetting.Controls.Add(Me.btnClose)
        Me.TabPageSetting.Location = New System.Drawing.Point(4, 22)
        Me.TabPageSetting.Name = "TabPageSetting"
        Me.TabPageSetting.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPageSetting.Size = New System.Drawing.Size(388, 258)
        Me.TabPageSetting.TabIndex = 0
        Me.TabPageSetting.Text = "Setting"
        Me.TabPageSetting.UseVisualStyleBackColor = True
        '
        'TabPageExcel
        '
        Me.TabPageExcel.Location = New System.Drawing.Point(4, 22)
        Me.TabPageExcel.Name = "TabPageExcel"
        Me.TabPageExcel.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPageExcel.Size = New System.Drawing.Size(388, 258)
        Me.TabPageExcel.TabIndex = 1
        Me.TabPageExcel.Text = "Import Excel"
        Me.TabPageExcel.UseVisualStyleBackColor = True
        '
        'TabPageMaintain
        '
        Me.TabPageMaintain.Location = New System.Drawing.Point(4, 22)
        Me.TabPageMaintain.Name = "TabPageMaintain"
        Me.TabPageMaintain.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPageMaintain.Size = New System.Drawing.Size(388, 258)
        Me.TabPageMaintain.TabIndex = 2
        Me.TabPageMaintain.Text = "Maintainance"
        Me.TabPageMaintain.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.lblSQLStatus)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 264)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(396, 20)
        Me.Panel1.TabIndex = 29
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label1.Location = New System.Drawing.Point(280, 0)
        Me.Label1.Margin = New System.Windows.Forms.Padding(3, 1, 3, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(43, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Status: "
        '
        'lblSQLStatus
        '
        Me.lblSQLStatus.AutoSize = True
        Me.lblSQLStatus.BackColor = System.Drawing.Color.Transparent
        Me.lblSQLStatus.Dock = System.Windows.Forms.DockStyle.Right
        Me.lblSQLStatus.ForeColor = System.Drawing.Color.DarkRed
        Me.lblSQLStatus.Location = New System.Drawing.Point(323, 0)
        Me.lblSQLStatus.Name = "lblSQLStatus"
        Me.lblSQLStatus.Size = New System.Drawing.Size(73, 13)
        Me.lblSQLStatus.TabIndex = 11
        Me.lblSQLStatus.Text = "Disconnected"
        '
        'SQL_Connection_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(396, 284)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.TabControl1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "SQL_Connection_Form"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "SQL Connection Form"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPageSetting.ResumeLayout(False)
        Me.TabPageSetting.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents cbDatabase As ComboBox
    Friend WithEvents btnSql As Button
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
    Friend WithEvents TabPageSetting As TabPage
    Friend WithEvents TabPageExcel As TabPage
    Friend WithEvents TabPageMaintain As TabPage
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Label1 As Label
    Friend WithEvents lblSQLStatus As Label
End Class
