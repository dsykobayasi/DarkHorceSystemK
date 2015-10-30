Option Explicit On
Public Class frmWrappedJVLink
    '
    '   JVLinkがインストールされているかどうかの判定の為の隠しフォーム
    '
    '   Visible=Falseでユーザーからは隠匿されます。
    '

    '---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
    '   プロパティ
    '---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

    '
    '   機能: JVLinkオブジェクトを返す
    '
    '   備考: なし
    '
    Public Property JVlink() As AxJVDTLabLib.AxJVLink
        Get
            JVlink = axJVLink
        End Get
        Set(ByVal value As AxJVDTLabLib.AxJVLink)

        End Set
    End Property

    '
    '   機能: JVLInkオブジェクトのバージョンを返す
    '
    '   備考: なし
    '
    Public Property m_JVLinkVersion() As String
        Get
            m_JVLinkVersion = axJVLink.m_JVLinkVersion
        End Get
        Set(ByVal value As String)

        End Set
    End Property

    Private Sub frmWrappedJVLink_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New Point(-300, -300)
    End Sub
End Class