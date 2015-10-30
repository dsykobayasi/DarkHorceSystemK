<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DoubleBufferedSplitContainer
    Inherits System.Windows.Forms.SplitContainer

    Public Sub New()
        MyBase.New()
        Call InitializeComponent()
        '以下の３つをTrueにすることでダブルバッファリングを有効にする
        Me.SetStyle(ControlStyles.UserPaint, True)
        Me.SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        Me.SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
    End Sub

    Protected Overrides Property DoubleBuffered() As Boolean
        Get
            Return True
        End Get
        Set(ByVal value As Boolean)
        End Set
    End Property

    'UserControl はコンポーネント一覧をクリーンアップするために dispose をオーバーライドします。
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
        components = New System.ComponentModel.Container()
    End Sub
End Class
