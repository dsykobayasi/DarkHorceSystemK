<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgLogin
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
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

    'Windows フォーム デザイナで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
    'Windows フォーム デザイナを使用して変更できます。  
    'コード エディタを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(dlgLogin))
        Me.lblNotPass = New System.Windows.Forms.Label
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me.chkPassSave = New System.Windows.Forms.CheckBox
        Me.txtMailAddress = New System.Windows.Forms.TextBox
        Me.lblPassword = New System.Windows.Forms.Label
        Me.lblMail = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnLogin = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'lblNotPass
        '
        Me.lblNotPass.AutoSize = True
        Me.lblNotPass.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblNotPass.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblNotPass.ForeColor = System.Drawing.Color.Blue
        Me.lblNotPass.Location = New System.Drawing.Point(78, 141)
        Me.lblNotPass.Name = "lblNotPass"
        Me.lblNotPass.Size = New System.Drawing.Size(152, 12)
        Me.lblNotPass.TabIndex = 14
        Me.lblNotPass.Text = "パスワードを忘れてしまった場合"
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(89, 56)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(143, 19)
        Me.txtPassword.TabIndex = 8
        '
        'chkPassSave
        '
        Me.chkPassSave.AutoSize = True
        Me.chkPassSave.Location = New System.Drawing.Point(89, 84)
        Me.chkPassSave.Name = "chkPassSave"
        Me.chkPassSave.Size = New System.Drawing.Size(136, 16)
        Me.chkPassSave.TabIndex = 9
        Me.chkPassSave.Text = "ログイン情報を保存する"
        Me.chkPassSave.UseVisualStyleBackColor = True
        '
        'txtMailAddress
        '
        Me.txtMailAddress.Location = New System.Drawing.Point(89, 21)
        Me.txtMailAddress.Name = "txtMailAddress"
        Me.txtMailAddress.Size = New System.Drawing.Size(194, 19)
        Me.txtMailAddress.TabIndex = 7
        '
        'lblPassword
        '
        Me.lblPassword.AutoSize = True
        Me.lblPassword.Location = New System.Drawing.Point(29, 59)
        Me.lblPassword.Name = "lblPassword"
        Me.lblPassword.Size = New System.Drawing.Size(54, 12)
        Me.lblPassword.TabIndex = 12
        Me.lblPassword.Text = "パスワード:"
        '
        'lblMail
        '
        Me.lblMail.AutoSize = True
        Me.lblMail.Location = New System.Drawing.Point(12, 24)
        Me.lblMail.Name = "lblMail"
        Me.lblMail.Size = New System.Drawing.Size(71, 12)
        Me.lblMail.TabIndex = 10
        Me.lblMail.Text = "メールアドレス:"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(157, 106)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 13
        Me.btnCancel.Text = "キャンセル"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnLogin
        '
        Me.btnLogin.Location = New System.Drawing.Point(76, 106)
        Me.btnLogin.Name = "btnLogin"
        Me.btnLogin.Size = New System.Drawing.Size(75, 23)
        Me.btnLogin.TabIndex = 11
        Me.btnLogin.Text = "ログイン"
        Me.btnLogin.UseVisualStyleBackColor = True
        '
        'dlgLogin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(304, 162)
        Me.Controls.Add(Me.lblNotPass)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.chkPassSave)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnLogin)
        Me.Controls.Add(Me.txtMailAddress)
        Me.Controls.Add(Me.lblPassword)
        Me.Controls.Add(Me.lblMail)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dlgLogin"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ログイン"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblNotPass As System.Windows.Forms.Label
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents chkPassSave As System.Windows.Forms.CheckBox
    Friend WithEvents txtMailAddress As System.Windows.Forms.TextBox
    Friend WithEvents lblPassword As System.Windows.Forms.Label
    Friend WithEvents lblMail As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnLogin As System.Windows.Forms.Button

End Class
