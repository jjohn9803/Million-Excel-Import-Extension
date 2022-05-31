<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Login
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Login))
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtStaffID = New System.Windows.Forms.TextBox()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.btnSetting = New System.Windows.Forms.Button()
        Me.lblCapsLock = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Yu Gothic Medium", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(44, 30)
        Me.Label7.Margin = New System.Windows.Forms.Padding(13, 18, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(89, 22)
        Me.Label7.TabIndex = 31
        Me.Label7.Text = "Staff ID: "
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Yu Gothic Medium", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(25, 98)
        Me.Label1.Margin = New System.Windows.Forms.Padding(13, 18, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(107, 22)
        Me.Label1.TabIndex = 32
        Me.Label1.Text = "Password: "
        '
        'txtStaffID
        '
        Me.txtStaffID.Font = New System.Drawing.Font("Yu Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStaffID.Location = New System.Drawing.Point(169, 26)
        Me.txtStaffID.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtStaffID.Name = "txtStaffID"
        Me.txtStaffID.Size = New System.Drawing.Size(240, 34)
        Me.txtStaffID.TabIndex = 0
        '
        'txtPassword
        '
        Me.txtPassword.Font = New System.Drawing.Font("Yu Gothic", 9.75!)
        Me.txtPassword.Location = New System.Drawing.Point(169, 95)
        Me.txtPassword.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(240, 34)
        Me.txtPassword.TabIndex = 1
        '
        'btnSetting
        '
        Me.btnSetting.Font = New System.Drawing.Font("Yu Gothic Medium", 9.75!, System.Drawing.FontStyle.Bold)
        Me.btnSetting.Location = New System.Drawing.Point(140, 159)
        Me.btnSetting.Margin = New System.Windows.Forms.Padding(187, 9, 4, 4)
        Me.btnSetting.Name = "btnSetting"
        Me.btnSetting.Size = New System.Drawing.Size(160, 37)
        Me.btnSetting.TabIndex = 2
        Me.btnSetting.Text = "Login"
        Me.btnSetting.UseVisualStyleBackColor = True
        '
        'lblCapsLock
        '
        Me.lblCapsLock.AutoSize = True
        Me.lblCapsLock.Font = New System.Drawing.Font("Yu Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCapsLock.ForeColor = System.Drawing.SystemColors.WindowFrame
        Me.lblCapsLock.Location = New System.Drawing.Point(166, 133)
        Me.lblCapsLock.Name = "lblCapsLock"
        Me.lblCapsLock.Size = New System.Drawing.Size(128, 19)
        Me.lblCapsLock.TabIndex = 33
        Me.lblCapsLock.Text = "Caps Lock is on!"
        Me.lblCapsLock.Visible = False
        '
        'Login
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(456, 226)
        Me.Controls.Add(Me.lblCapsLock)
        Me.Controls.Add(Me.btnSetting)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.txtStaffID)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label7)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.Name = "Login"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label7 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents txtStaffID As TextBox
    Friend WithEvents txtPassword As TextBox
    Friend WithEvents btnSetting As Button
    Friend WithEvents lblCapsLock As Label
End Class
