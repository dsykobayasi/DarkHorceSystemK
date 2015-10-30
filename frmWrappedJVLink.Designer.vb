<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmWrappedJVLink
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmWrappedJVLink))
        Me.axJVLink = New AxJVDTLabLib.AxJVLink
        CType(Me.axJVLink, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'axJVLink
        '
        Me.axJVLink.Enabled = True
        Me.axJVLink.Location = New System.Drawing.Point(12, 12)
        Me.axJVLink.Name = "axJVLink"
        Me.axJVLink.OcxState = CType(resources.GetObject("axJVLink.OcxState"), System.Windows.Forms.AxHost.State)
        Me.axJVLink.Size = New System.Drawing.Size(192, 192)
        Me.axJVLink.TabIndex = 0
        '
        'frmWrappedJVLink
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(220, 221)
        Me.Controls.Add(Me.axJVLink)
        Me.Name = "frmWrappedJVLink"
        Me.ShowInTaskbar = False
        Me.Text = "JVLink"
        Me.TopMost = True
        Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
        CType(Me.axJVLink, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents axJVLink As AxJVDTLabLib.AxJVLink
End Class
