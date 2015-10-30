<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgTimer
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
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
        Me.OK_Button = New System.Windows.Forms.Button
        Me.Cancel_Button = New System.Windows.Forms.Button
        Me.chkTimerOnOff = New System.Windows.Forms.CheckBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.dtpToTime = New System.Windows.Forms.DateTimePicker
        Me.dtpFromTime = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.nudUpdate = New System.Windows.Forms.NumericUpDown
        Me.lblExecuteTime = New System.Windows.Forms.Label
        Me.lblUpdate = New System.Windows.Forms.Label
        Me.TableLayoutPanel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.nudUpdate, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(134, 136)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 27)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Location = New System.Drawing.Point(3, 3)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(67, 21)
        Me.OK_Button.TabIndex = 5
        Me.OK_Button.Text = "OK"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(76, 3)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(67, 21)
        Me.Cancel_Button.TabIndex = 6
        Me.Cancel_Button.Text = "キャンセル"
        '
        'chkTimerOnOff
        '
        Me.chkTimerOnOff.AutoSize = True
        Me.chkTimerOnOff.Location = New System.Drawing.Point(13, 13)
        Me.chkTimerOnOff.Name = "chkTimerOnOff"
        Me.chkTimerOnOff.Size = New System.Drawing.Size(145, 16)
        Me.chkTimerOnOff.TabIndex = 1
        Me.chkTimerOnOff.Text = "タイマー機能を有効にする"
        Me.chkTimerOnOff.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.dtpToTime)
        Me.GroupBox1.Controls.Add(Me.dtpFromTime)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.nudUpdate)
        Me.GroupBox1.Controls.Add(Me.lblExecuteTime)
        Me.GroupBox1.Controls.Add(Me.lblUpdate)
        Me.GroupBox1.Location = New System.Drawing.Point(13, 36)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(267, 94)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        '
        'dtpToTime
        '
        Me.dtpToTime.CustomFormat = "HH:mm"
        Me.dtpToTime.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpToTime.Location = New System.Drawing.Point(205, 49)
        Me.dtpToTime.Name = "dtpToTime"
        Me.dtpToTime.ShowUpDown = True
        Me.dtpToTime.Size = New System.Drawing.Size(50, 19)
        Me.dtpToTime.TabIndex = 6
        Me.dtpToTime.Value = New Date(2010, 7, 29, 0, 0, 0, 0)
        '
        'dtpFromTime
        '
        Me.dtpFromTime.CustomFormat = "HH:mm"
        Me.dtpFromTime.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFromTime.Location = New System.Drawing.Point(126, 49)
        Me.dtpFromTime.Name = "dtpFromTime"
        Me.dtpFromTime.ShowUpDown = True
        Me.dtpFromTime.Size = New System.Drawing.Size(50, 19)
        Me.dtpFromTime.TabIndex = 3
        Me.dtpFromTime.Value = New Date(2010, 7, 29, 0, 0, 0, 0)
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(182, 52)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(17, 12)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "～"
        '
        'nudUpdate
        '
        Me.nudUpdate.Location = New System.Drawing.Point(126, 17)
        Me.nudUpdate.Maximum = New Decimal(New Integer() {60, 0, 0, 0})
        Me.nudUpdate.Minimum = New Decimal(New Integer() {15, 0, 0, 0})
        Me.nudUpdate.Name = "nudUpdate"
        Me.nudUpdate.Size = New System.Drawing.Size(50, 19)
        Me.nudUpdate.TabIndex = 2
        Me.nudUpdate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.nudUpdate.Value = New Decimal(New Integer() {15, 0, 0, 0})
        '
        'lblExecuteTime
        '
        Me.lblExecuteTime.AutoSize = True
        Me.lblExecuteTime.Location = New System.Drawing.Point(36, 52)
        Me.lblExecuteTime.Name = "lblExecuteTime"
        Me.lblExecuteTime.Size = New System.Drawing.Size(86, 12)
        Me.lblExecuteTime.TabIndex = 1
        Me.lblExecuteTime.Text = "実行する時間帯:"
        '
        'lblUpdate
        '
        Me.lblUpdate.AutoSize = True
        Me.lblUpdate.Location = New System.Drawing.Point(7, 19)
        Me.lblUpdate.Name = "lblUpdate"
        Me.lblUpdate.Size = New System.Drawing.Size(115, 12)
        Me.lblUpdate.TabIndex = 0
        Me.lblUpdate.Text = "更新間隔（15～60分）:"
        '
        'dlgTimer
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(292, 173)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.chkTimerOnOff)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dlgTimer"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "タイマー設定"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.nudUpdate, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents chkTimerOnOff As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents nudUpdate As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblExecuteTime As System.Windows.Forms.Label
    Friend WithEvents lblUpdate As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtpFromTime As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpToTime As System.Windows.Forms.DateTimePicker

End Class
