<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgKakoDelete
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
        Me.rbDelList0 = New System.Windows.Forms.RadioButton
        Me.rbDelList1 = New System.Windows.Forms.RadioButton
        Me.rbDelList2 = New System.Windows.Forms.RadioButton
        Me.rbDelList3 = New System.Windows.Forms.RadioButton
        Me.rbDelList4 = New System.Windows.Forms.RadioButton
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.TableLayoutPanel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(127, 165)
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
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "OK"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(76, 3)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(67, 21)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.Text = "キャンセル"
        '
        'rbDelList0
        '
        Me.rbDelList0.AutoSize = True
        Me.rbDelList0.Checked = True
        Me.rbDelList0.Location = New System.Drawing.Point(17, 18)
        Me.rbDelList0.Name = "rbDelList0"
        Me.rbDelList0.Size = New System.Drawing.Size(79, 16)
        Me.rbDelList0.TabIndex = 1
        Me.rbDelList0.TabStop = True
        Me.rbDelList0.Text = "過去２年分"
        Me.rbDelList0.UseVisualStyleBackColor = True
        '
        'rbDelList1
        '
        Me.rbDelList1.AutoSize = True
        Me.rbDelList1.Location = New System.Drawing.Point(17, 40)
        Me.rbDelList1.Name = "rbDelList1"
        Me.rbDelList1.Size = New System.Drawing.Size(79, 16)
        Me.rbDelList1.TabIndex = 2
        Me.rbDelList1.Text = "過去１年分"
        Me.rbDelList1.UseVisualStyleBackColor = True
        '
        'rbDelList2
        '
        Me.rbDelList2.AutoSize = True
        Me.rbDelList2.Location = New System.Drawing.Point(17, 62)
        Me.rbDelList2.Name = "rbDelList2"
        Me.rbDelList2.Size = New System.Drawing.Size(83, 16)
        Me.rbDelList2.TabIndex = 3
        Me.rbDelList2.Text = "過去半年分"
        Me.rbDelList2.UseVisualStyleBackColor = True
        '
        'rbDelList3
        '
        Me.rbDelList3.AutoSize = True
        Me.rbDelList3.Location = New System.Drawing.Point(17, 84)
        Me.rbDelList3.Name = "rbDelList3"
        Me.rbDelList3.Size = New System.Drawing.Size(89, 16)
        Me.rbDelList3.TabIndex = 4
        Me.rbDelList3.Text = "過去１か月分"
        Me.rbDelList3.UseVisualStyleBackColor = True
        '
        'rbDelList4
        '
        Me.rbDelList4.AutoSize = True
        Me.rbDelList4.Location = New System.Drawing.Point(17, 106)
        Me.rbDelList4.Name = "rbDelList4"
        Me.rbDelList4.Size = New System.Drawing.Size(87, 16)
        Me.rbDelList4.TabIndex = 5
        Me.rbDelList4.Text = "過去１５日分"
        Me.rbDelList4.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rbDelList4)
        Me.GroupBox1.Controls.Add(Me.rbDelList3)
        Me.GroupBox1.Controls.Add(Me.rbDelList2)
        Me.GroupBox1.Controls.Add(Me.rbDelList1)
        Me.GroupBox1.Controls.Add(Me.rbDelList0)
        Me.GroupBox1.Location = New System.Drawing.Point(21, 10)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(225, 140)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "データ保持期間"
        '
        'dlgKakoDelete
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(285, 203)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dlgKakoDelete"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "過去レース情報削除"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents rbDelList0 As System.Windows.Forms.RadioButton
    Friend WithEvents rbDelList1 As System.Windows.Forms.RadioButton
    Friend WithEvents rbDelList2 As System.Windows.Forms.RadioButton
    Friend WithEvents rbDelList3 As System.Windows.Forms.RadioButton
    Friend WithEvents rbDelList4 As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox

End Class
