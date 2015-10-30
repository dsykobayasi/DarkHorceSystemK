Imports System.Windows.Forms

Public Class dlgVersionInfo


    Private Sub VersionInfo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim strVer As String = String.Empty

        ' ClickOnceでインストールされている場合のみ取得
        If System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed Then
            Dim ver As Version = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion
            strVer = ver.Major.ToString() + "."
            strVer += ver.Minor.ToString() + "."
            strVer += ver.Build.ToString() + "."
            strVer += ver.Revision.ToString()

            ' バージョンを表示（メジャー、マイナー、リビジョン、ビルド）
            Me.lblVersion.Text = "Ver " & ver.ToString
            Me.lstVersionHistory.Items.Add("2014/07/2 新発売 Ver" & ver.ToString)
        End If
        '' 自分自身のAssemblyを取得
        'Dim asm As System.Reflection.Assembly = _
        '    System.Reflection.Assembly.GetExecutingAssembly()
        '' バージョンの取得
        'Dim ver As System.Version = asm.GetName().Version

        '' バージョンを表示（メジャー、マイナー、リビジョン、ビルド）
        'Me.lblVersion.Text = "Ver " & ver.ToString
        'Me.lstVersionHistory.Items.Add("2010/9/15 新発売 Ver" & ver.ToString)
    End Sub

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

End Class
