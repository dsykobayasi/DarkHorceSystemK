<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgVersionInfo
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
        Me.btnOk = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.lstVersionHistory = New System.Windows.Forms.ListBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.picVersion = New System.Windows.Forms.PictureBox
        Me.lblVersion = New System.Windows.Forms.Label
        CType(Me.picVersion, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(323, 251)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(75, 23)
        Me.btnOk.TabIndex = 0
        Me.btnOk.Text = "OK"
        Me.btnOk.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(203, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(154, 12)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "互當穴ノ守　十周年記念ソフト"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(203, 78)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(154, 12)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Copyright © 2014 互當穴ノ守"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(203, 111)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(61, 12)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "互當穴ノ守"
        '
        'lstVersionHistory
        '
        Me.lstVersionHistory.BackColor = System.Drawing.Color.White
        Me.lstVersionHistory.FormattingEnabled = True
        Me.lstVersionHistory.ItemHeight = 12
        Me.lstVersionHistory.Location = New System.Drawing.Point(205, 172)
        Me.lstVersionHistory.Name = "lstVersionHistory"
        Me.lstVersionHistory.Size = New System.Drawing.Size(193, 64)
        Me.lstVersionHistory.TabIndex = 6
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(203, 157)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(98, 12)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "バージョン更新履歴"
        '
        'picVersion
        '
        Me.picVersion.Image = Global.AnaSoftVerK.My.Resources.Resources.versionImg
        Me.picVersion.Location = New System.Drawing.Point(13, 13)
        Me.picVersion.Name = "picVersion"
        Me.picVersion.Size = New System.Drawing.Size(172, 250)
        Me.picVersion.TabIndex = 1
        Me.picVersion.TabStop = False
        '
        'lblVersion
        '
        Me.lblVersion.AutoSize = True
        Me.lblVersion.Location = New System.Drawing.Point(204, 134)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.Size = New System.Drawing.Size(0, 12)
        Me.lblVersion.TabIndex = 8
        '
        'dlgVersionInfo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(435, 291)
        Me.Controls.Add(Me.lblVersion)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.lstVersionHistory)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.picVersion)
        Me.Controls.Add(Me.btnOk)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dlgVersionInfo"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "互當穴ノ守 十周年記念ソフト バージョン情報"
        CType(Me.picVersion, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents picVersion As System.Windows.Forms.PictureBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lstVersionHistory As System.Windows.Forms.ListBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents lblVersion As System.Windows.Forms.Label

End Class
