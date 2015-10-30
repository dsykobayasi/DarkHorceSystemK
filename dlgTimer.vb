Imports System.Windows.Forms

Public Class dlgTimer
    'Private strIniPath As String    ' INIファイルパス

    'OKボタン押下時処理
    Private Sub OK_Button_Click(ByVal sender As System.Object, _
                                ByVal e As System.EventArgs) Handles OK_Button.Click

        Dim dtNowTime As DateTime
        Dim dtFromTime As DateTime
        Dim dtToTime As DateTime
        gFromTimeLabel = ""

        ' 設定された内容をチェックする
        ' 更新間隔
        If (Me.nudUpdate.Value.ToString = "") Or _
            (Me.nudUpdate.Value.ToString < CommonConstant.MIN_15) Or _
            (Me.nudUpdate.Value.ToString > CommonConstant.MIN_60) Then
            MsgBox("更新間隔は15～60の間で入力してください。", _
                   MsgBoxStyle.Exclamation, CommonConstant.AppTitle)
            Exit Sub
        End If

        '' 設定された内容をINIファイルに更新する
        'strIniPath = Application.StartupPath & CommonConstant.INIFile

        ' タイマーフラグを取得()
        PutIni(CommonConstant.INI_Ap_Timer, CommonConstant.INI_Key_TimerFlg, _
               Me.chkTimerOnOff.Checked, gIniFile)
        ' 更新間隔を取得
        PutIni(CommonConstant.INI_Ap_Timer, CommonConstant.INI_Key_UpdateTime, _
               Me.nudUpdate.Value, gIniFile)

        ' 実行する時間帯を調べる
        ' 現在時刻を取得
        dtNowTime = New DateTime(Now.Ticks)
        dtFromTime = DateTime.Parse(Me.dtpFromTime.Value)
        dtToTime = DateTime.Parse(Me.dtpToTime.Value)
        Dim strNowMinute As String = IIf(dtNowTime.Minute.ToString.Length = CommonConstant.Const1, _
                                CommonConstant.Const0 & dtNowTime.Minute, dtNowTime.Minute)
        Dim strFromMinute As String = IIf(dtFromTime.Minute.ToString.Length = CommonConstant.Const1, _
                                CommonConstant.Const0 & dtFromTime.Minute, dtFromTime.Minute)
        Dim strToMinute As String = IIf(dtToTime.Minute.ToString.Length = CommonConstant.Const1, _
                                CommonConstant.Const0 & dtToTime.Minute, dtToTime.Minute)
        ' 現在時刻が実行する時間帯のFROM～TO以内なら実行時間とする 
        If (CInt(dtNowTime.Hour & strNowMinute) >= CInt(dtFromTime.Hour & strFromMinute) And _
            (CInt(dtNowTime.Hour & strNowMinute)) <= CInt(dtToTime.Hour & strToMinute)) Then
            'If (dtNowTime.Hour & dtNowTime.Minute >= DateTime.Parse(Me.dtpFromTime.Value).Hour & DateTime.Parse(Me.dtpFromTime.Value).Minute) And _
            '    (dtNowTime.Hour & dtNowTime.Minute <= DateTime.Parse(Me.dtpToTime.Value).Hour & DateTime.Parse(Me.dtpToTime.Value).Minute) Then
            gFromTime = Format(dtNowTime.AddMinutes(Me.nudUpdate.Value), CommonConstant.TimeHHmm)
            gFromTimeLabel = Format(dtNowTime.AddMinutes(Me.nudUpdate.Value), CommonConstant.TimeHHmm)
        Else
            '　開始時間をクリア
            gFromTime = ""
        End If

        ' 実行する時間帯（FROM）を取得
        PutIni(CommonConstant.INI_Ap_Timer, CommonConstant.INI_Key_FromTime, _
               Me.dtpFromTime.Value.ToString, gIniFile)
        ' 実行する時間帯（TO）を取得
        PutIni(CommonConstant.INI_Ap_Timer, CommonConstant.INI_Key_ToTime, _
               Me.dtpToTime.Value.ToString, gIniFile)
        ''　開始時間をクリア
        'gFromTime = ""

        ' バージョン確認
        If Me.chkTimerOnOff.Checked And _
            gVersion < CommonConstant.CompVersion250 Then
            MsgBox("インストールされているJVLinkのバージョンはVer2.4以前です。（Ver2.5以降推奨）" & vbCrLf & vbCrLf & _
                    "現在インストールされているJVLinkの場合、データ更新を実行する度に新バージョン告知のメッセージが表示されますのでご注意下さい。" & vbCrLf & _
                    "告知メッセージの表示を停止する場合はVer2.5以降のJVLinkをインストールして下さい。" & vbCrLf & _
                    "※影響についてはヘルプ（タイマー機能を使用する）をご参照下さい。", _
                    MsgBoxStyle.Exclamation + MsgBoxStyle.SystemModal, CommonConstant.AppTitle)
        End If

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    'キャンセルボタン押下時処理
    Private Sub Cancel_Button_Click(ByVal sender As System.Object, _
                                    ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    'タイマーを有効にするチェック
    Private Sub chkTimerOnOff_CheckedChanged(ByVal sender As System.Object, _
                                             ByVal e As System.EventArgs) Handles chkTimerOnOff.CheckedChanged
        If Me.chkTimerOnOff.Checked Then
            Me.nudUpdate.Enabled = True
            Me.dtpFromTime.Enabled = True
            Me.dtpToTime.Enabled = True
        Else
            Me.nudUpdate.Enabled = False
            Me.dtpFromTime.Enabled = False
            Me.dtpToTime.Enabled = False
        End If
    End Sub

    '画面初期処理
    Private Sub dlgTimer_Load(ByVal sender As System.Object, _
                              ByVal e As System.EventArgs) Handles MyBase.Load
        Dim blnTimerFlg As Boolean
        Dim intTimer As Integer
        Dim strFromTime As String
        Dim strToTime As String

        '' INIファイルパスを設定
        'strIniPath = Application.StartupPath & CommonConstant.INIFile

        '2011/04/07 d-kobayashi update start
        ' INIファイルが存在するか
        'If Dir(gIniFile) <> "" Then
        ' タイマーフラグを取得()
        ''blnTimerFlg = GetIni(CommonConstant.INI_Ap_Timer, CommonConstant.INI_Key_TimerFlg, _
        ''                     "", gIniFile)
        ' 更新間隔を取得
        'intTimer = GetIni(CommonConstant.INI_Ap_Timer, CommonConstant.INI_Key_UpdateTime, _
        '                  "", gIniFile)
        ' 実行する時間帯（FROM）を取得
        'strFromTime = GetIni(CommonConstant.INI_Ap_Timer, CommonConstant.INI_Key_FromTime, _
        '                     "", gIniFile)
        ' 実行する時間帯（TO）を取得
        'strToTime = GetIni(CommonConstant.INI_Ap_Timer, CommonConstant.INI_Key_ToTime, _
        '                   "", gIniFile)
        ' タイマーフラグを取得()
        blnTimerFlg = GetIni(CommonConstant.INI_Ap_Timer, CommonConstant.INI_Key_TimerFlg, _
                             CommonConstant.INI_TIMERFLG, gIniFile)
        ' 更新間隔を取得
        intTimer = GetIni(CommonConstant.INI_Ap_Timer, CommonConstant.INI_Key_UpdateTime, _
                          CommonConstant.MIN_15, gIniFile)
        ' 実行する時間帯（FROM）を取得
        strFromTime = GetIni(CommonConstant.INI_Ap_Timer, CommonConstant.INI_Key_FromTime, _
                             Format(Now, "yyyy/MM/dd") & CommonConstant.INI_FROMTIME, gIniFile)
        ' 実行する時間帯（TO）を取得
        strToTime = GetIni(CommonConstant.INI_Ap_Timer, CommonConstant.INI_Key_ToTime, _
                           Format(Now, "yyyy/MM/dd") & CommonConstant.INI_TOTIME, gIniFile)
        '2011/04/07 d-kobayashi update end

        ' INIファイルから取得した情報を画面に設定
        Me.chkTimerOnOff.Checked = blnTimerFlg
        Me.nudUpdate.Value = intTimer
        Me.dtpFromTime.Value = DateTime.Parse(strFromTime)
        Me.dtpToTime.Value = DateTime.Parse(strToTime)
        'Else
        ' 初期情報を画面に設定
        '2011/04/07 d-kobayashi delete
        'INIファイルが無い場合、もしくはキー・セクション情報が無い場合でも取得可能にしたため、削除
        'Me.chkTimerOnOff.Checked = False
        'Me.nudUpdate.Value = 15
        ''Me.dtpFromTime.Value = New DateTime(Now().Ticks)
        ''Me.dtpToTime.Value = New DateTime(Now().Ticks)
        'Me.dtpFromTime.Value = DateTime.Parse(Format(Now, "yyyy/MM/dd") & CommonConstant.INI_FROMTIME)
        'Me.dtpToTime.Value = DateTime.Parse(Format(Now, "yyyy/MM/dd") & CommonConstant.INI_TOTIME)
        'End If

        ' チェックイベントを実行
        Call chkTimerOnOff_CheckedChanged(sender, e)
    End Sub
End Class
