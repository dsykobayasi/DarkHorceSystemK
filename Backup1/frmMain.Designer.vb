<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.MenuStrip = New System.Windows.Forms.MenuStrip
        Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuExit = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuJVLink = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuSetting = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuCacheClr = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuData = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuKakoDelete = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuAbout = New System.Windows.Forms.ToolStripMenuItem
        Me.TabControl = New System.Windows.Forms.TabControl
        Me.tabUpdate = New System.Windows.Forms.TabPage
        Me.btnExcel = New System.Windows.Forms.Button
        Me.Panel326 = New System.Windows.Forms.Panel
        Me.TimerSetup = New System.Windows.Forms.Label
        Me.lstUpdateInfo = New System.Windows.Forms.ListBox
        Me.prgDownload = New System.Windows.Forms.ProgressBar
        Me.btnDataUpdate = New System.Windows.Forms.Button
        Me.tmrDownload = New System.Windows.Forms.Timer(Me.components)
        Me.tmrTimerSet = New System.Windows.Forms.Timer(Me.components)
        Me.pdPrintDoc = New System.Drawing.Printing.PrintDocument
        Me.pdDialog = New System.Windows.Forms.PrintDialog
        Me.AxJVLink1 = New AxJVDTLabLib.AxJVLink
        Me.MenuStrip.SuspendLayout()
        Me.TabControl.SuspendLayout()
        Me.tabUpdate.SuspendLayout()
        Me.Panel326.SuspendLayout()
        CType(Me.AxJVLink1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip
        '
        Me.MenuStrip.BackColor = System.Drawing.SystemColors.Control
        Me.MenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile, Me.mnuJVLink, Me.mnuData, Me.mnuHelp})
        Me.MenuStrip.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip.Name = "MenuStrip"
        Me.MenuStrip.Size = New System.Drawing.Size(746, 26)
        Me.MenuStrip.TabIndex = 1
        Me.MenuStrip.Text = "MenuStrip"
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuExit})
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(85, 22)
        Me.mnuFile.Text = "ファイル(&F)"
        '
        'mnuExit
        '
        Me.mnuExit.Name = "mnuExit"
        Me.mnuExit.Size = New System.Drawing.Size(118, 22)
        Me.mnuExit.Text = "終了(&X)"
        '
        'mnuJVLink
        '
        Me.mnuJVLink.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuSetting, Me.mnuCacheClr})
        Me.mnuJVLink.Name = "mnuJVLink"
        Me.mnuJVLink.Size = New System.Drawing.Size(77, 22)
        Me.mnuJVLink.Text = "JV-Link(&J)"
        '
        'mnuSetting
        '
        Me.mnuSetting.Name = "mnuSetting"
        Me.mnuSetting.Size = New System.Drawing.Size(190, 22)
        Me.mnuSetting.Text = "設定(&S)"
        '
        'mnuCacheClr
        '
        Me.mnuCacheClr.Name = "mnuCacheClr"
        Me.mnuCacheClr.Size = New System.Drawing.Size(190, 22)
        Me.mnuCacheClr.Text = "キャッシュクリア(&C)"
        '
        'mnuData
        '
        Me.mnuData.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuKakoDelete})
        Me.mnuData.Name = "mnuData"
        Me.mnuData.Size = New System.Drawing.Size(75, 22)
        Me.mnuData.Text = "データ(&D)"
        '
        'mnuKakoDelete
        '
        Me.mnuKakoDelete.Name = "mnuKakoDelete"
        Me.mnuKakoDelete.Size = New System.Drawing.Size(202, 22)
        Me.mnuKakoDelete.Text = "過去レース情報削除(&R)"
        '
        'mnuHelp
        '
        Me.mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuAbout})
        Me.mnuHelp.Name = "mnuHelp"
        Me.mnuHelp.Size = New System.Drawing.Size(75, 22)
        Me.mnuHelp.Text = "ヘルプ(&H)"
        '
        'mnuAbout
        '
        Me.mnuAbout.Name = "mnuAbout"
        Me.mnuAbout.Size = New System.Drawing.Size(190, 22)
        Me.mnuAbout.Text = "この製品について(&A)"
        '
        'TabControl
        '
        Me.TabControl.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl.Controls.Add(Me.tabUpdate)
        Me.TabControl.Location = New System.Drawing.Point(0, 26)
        Me.TabControl.Name = "TabControl"
        Me.TabControl.SelectedIndex = 0
        Me.TabControl.Size = New System.Drawing.Size(746, 329)
        Me.TabControl.TabIndex = 2
        '
        'tabUpdate
        '
        Me.tabUpdate.Controls.Add(Me.btnExcel)
        Me.tabUpdate.Controls.Add(Me.Panel326)
        Me.tabUpdate.Controls.Add(Me.lstUpdateInfo)
        Me.tabUpdate.Controls.Add(Me.prgDownload)
        Me.tabUpdate.Controls.Add(Me.btnDataUpdate)
        Me.tabUpdate.Location = New System.Drawing.Point(4, 22)
        Me.tabUpdate.Name = "tabUpdate"
        Me.tabUpdate.Padding = New System.Windows.Forms.Padding(3)
        Me.tabUpdate.Size = New System.Drawing.Size(738, 303)
        Me.tabUpdate.TabIndex = 1
        Me.tabUpdate.Text = "データ更新"
        '
        'btnExcel
        '
        Me.btnExcel.Location = New System.Drawing.Point(178, 7)
        Me.btnExcel.Name = "btnExcel"
        Me.btnExcel.Size = New System.Drawing.Size(130, 29)
        Me.btnExcel.TabIndex = 6
        Me.btnExcel.Text = "ソフト起動"
        Me.btnExcel.UseVisualStyleBackColor = True
        '
        'Panel326
        '
        Me.Panel326.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel326.Controls.Add(Me.TimerSetup)
        Me.Panel326.Location = New System.Drawing.Point(556, 6)
        Me.Panel326.Name = "Panel326"
        Me.Panel326.Size = New System.Drawing.Size(174, 30)
        Me.Panel326.TabIndex = 5
        '
        'TimerSetup
        '
        Me.TimerSetup.AutoSize = True
        Me.TimerSetup.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TimerSetup.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.TimerSetup.Location = New System.Drawing.Point(54, 9)
        Me.TimerSetup.Margin = New System.Windows.Forms.Padding(0)
        Me.TimerSetup.Name = "TimerSetup"
        Me.TimerSetup.Size = New System.Drawing.Size(93, 12)
        Me.TimerSetup.TabIndex = 5
        Me.TimerSetup.Text = "タイマーを設定する"
        Me.TimerSetup.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lstUpdateInfo
        '
        Me.lstUpdateInfo.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstUpdateInfo.FormattingEnabled = True
        Me.lstUpdateInfo.ItemHeight = 12
        Me.lstUpdateInfo.Location = New System.Drawing.Point(8, 63)
        Me.lstUpdateInfo.Name = "lstUpdateInfo"
        Me.lstUpdateInfo.Size = New System.Drawing.Size(722, 196)
        Me.lstUpdateInfo.TabIndex = 3
        '
        'prgDownload
        '
        Me.prgDownload.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prgDownload.Location = New System.Drawing.Point(8, 42)
        Me.prgDownload.Name = "prgDownload"
        Me.prgDownload.Size = New System.Drawing.Size(722, 15)
        Me.prgDownload.TabIndex = 2
        '
        'btnDataUpdate
        '
        Me.btnDataUpdate.BackColor = System.Drawing.SystemColors.Control
        Me.btnDataUpdate.Location = New System.Drawing.Point(8, 6)
        Me.btnDataUpdate.Name = "btnDataUpdate"
        Me.btnDataUpdate.Size = New System.Drawing.Size(150, 30)
        Me.btnDataUpdate.TabIndex = 1
        Me.btnDataUpdate.Text = "データを更新"
        Me.btnDataUpdate.UseVisualStyleBackColor = False
        '
        'tmrDownload
        '
        Me.tmrDownload.Interval = 500
        '
        'tmrTimerSet
        '
        Me.tmrTimerSet.Interval = 1000
        '
        'pdPrintDoc
        '
        Me.pdPrintDoc.DocumentName = "大穴忠臣蔵"
        '
        'pdDialog
        '
        Me.pdDialog.UseEXDialog = True
        '
        'AxJVLink1
        '
        Me.AxJVLink1.Enabled = True
        Me.AxJVLink1.Location = New System.Drawing.Point(542, 27)
        Me.AxJVLink1.Name = "AxJVLink1"
        Me.AxJVLink1.OcxState = CType(resources.GetObject("AxJVLink1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxJVLink1.Size = New System.Drawing.Size(192, 192)
        Me.AxJVLink1.TabIndex = 0
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Black
        Me.ClientSize = New System.Drawing.Size(746, 355)
        Me.Controls.Add(Me.TabControl)
        Me.Controls.Add(Me.AxJVLink1)
        Me.Controls.Add(Me.MenuStrip)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "大穴忠臣蔵verZ"
        Me.MenuStrip.ResumeLayout(False)
        Me.MenuStrip.PerformLayout()
        Me.TabControl.ResumeLayout(False)
        Me.tabUpdate.ResumeLayout(False)
        Me.Panel326.ResumeLayout(False)
        Me.Panel326.PerformLayout()
        CType(Me.AxJVLink1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip As System.Windows.Forms.MenuStrip
    Friend WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuJVLink As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuSetting As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuCacheClr As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuAbout As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TabControl As System.Windows.Forms.TabControl
    Friend WithEvents tabUpdate As System.Windows.Forms.TabPage
    Friend WithEvents lstUpdateInfo As System.Windows.Forms.ListBox
    Friend WithEvents prgDownload As System.Windows.Forms.ProgressBar
    Friend WithEvents btnDataUpdate As System.Windows.Forms.Button
    Friend WithEvents tmrDownload As System.Windows.Forms.Timer
    Friend WithEvents tmrTimerSet As System.Windows.Forms.Timer
    Friend WithEvents pdPrintDoc As System.Drawing.Printing.PrintDocument
    Friend WithEvents Panel326 As System.Windows.Forms.Panel
    Friend WithEvents TimerSetup As System.Windows.Forms.Label
    Friend WithEvents mnuData As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuKakoDelete As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents pdDialog As System.Windows.Forms.PrintDialog
    Friend WithEvents AxJVLink1 As AxJVDTLabLib.AxJVLink
    Friend WithEvents btnExcel As System.Windows.Forms.Button

End Class
