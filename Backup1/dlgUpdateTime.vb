Imports System.Windows.Forms

Public Class dlgUpdateTime

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        gSelUpdateTime = New ArrayList
        Dim CheckCount As Integer = clbUpdateTime.Items.Count - 1
        Dim umaNoIndex As Integer = -1
        Dim tanOddsIndex As Integer = 0
        Dim fukuOddsIndex As Integer = 1
        Dim addint As Integer = 3
        Dim BeforeTime As String = ""
        Dim selectCounter As Integer = 0
        Dim blnCheck As Boolean     'チェック有無フラグ

        blnCheck = False
        For i = 0 To CheckCount
            If clbUpdateTime.GetItemChecked(i) Then
                selectCounter = selectCounter + 1
                '選択内容（発表時間）を取得
                Dim DataTime As String = clbUpdateTime.SelectedItem
                'MsgBox("チェックされている")
                gSelUpdateTime.Add(gSelAllUpdateList(i))
                'チェック有
                blnCheck = True
            End If
        Next

        'チェック無の場合
        If Not blnCheck Then
            MsgBox("更新時間をチェックしてください。", MsgBoxStyle.Exclamation, CommonConstant.AppTitle)
            Exit Sub
        End If

        '選択数が1つ以下の場合
        'If selectCounter <= CommonConstant.Const1 Then
        '    MsgBox("更新時間は2個以上選択してください。", MsgBoxStyle.Exclamation, CommonConstant.AppTitle)
        '    Exit Sub
        'End If

        If gAuthority = CommonConstant.AUTHORITY Then
            '選択数が4つ以上の場合
            If selectCounter >= CommonConstant.Const4 Then
                MsgBox("選択できる更新時間は最大3個までです。", MsgBoxStyle.Exclamation, CommonConstant.AppTitle)
                Exit Sub
            End If
        Else
            '選択数が3つ以上の場合
            If selectCounter >= CommonConstant.Const3 Then
                MsgBox("選択できる更新時間は最大2個までです。", MsgBoxStyle.Exclamation, CommonConstant.AppTitle)
                Exit Sub
            End If
        End If

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    '画面初期化処理
    Private Sub dlgUpdateTime_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ' リストボックスが0件の場合
        If Me.clbUpdateTime.Items.Count = CommonConstant.Const0 Then
            ' OKボタンを使用不可にする
            Me.OK_Button.Enabled = False
        Else
            ' OKボタンを使用可にする
            Me.OK_Button.Enabled = True
        End If
    End Sub
End Class
