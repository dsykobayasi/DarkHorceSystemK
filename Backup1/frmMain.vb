Imports System.IO

Public Class frmMain
    Private lDownloadCount As Long          ' JVOpen: 総ダウンロードファイル数
    Private JVOpenFlg As Long               ' JVOpen状態フラグ Open時:True
    Private objCommonUtil As clsCommonUtil  ' CommonUtilクラスインスタンス
    Private strJVFromTime As String         ' JVOpen: ダウンロード開始ポイント時間

    Private aRaceList As New ArrayList      ' レース詳細のデータリスト
    Private aHorseList As New ArrayList     ' 馬毎レース情報のデータリスト
    Private aOdds1List As New ArrayList     ' オッズ1（単複枠）のデータリスト
    Private aOdds2List As New ArrayList     ' オッズ2（馬連）のデータリスト
    Private aOdds4List As New ArrayList     ' オッズ4（馬単）のデータリスト
    Private strRAFileName As String         ' レース詳細JVDファイル名
    Private strSEFileName As String         ' 馬毎レース情報JVDファイル名
    Private strO1FileName As String         ' オッズ1情報JVDファイル名
    Private strO2FileName As String         ' オッズ2情報JVDファイル名
    Private strO4FileName As String         ' オッズ4情報JVDファイル名

    Private RaceCheckIndexList As ArrayList = New ArrayList 'レース一覧のキー情報保有リスト
    Private RaceDataFile As String = String.Empty
    Private OddsDataFile As String = String.Empty
    'Dim SelUpdateList As ArrayList = New ArrayList
    'Private ComUtil As clsCommonUtil = New clsCommonUtil 'ユーティリティクラス
    Private Ksort As clsKeySort = New clsKeySort 'ソートクラス
    Private Housoku As clsHousoku = New clsHousoku '法則演算クラス

    Private TanshouMaxNinkiUmaNo As String = String.Empty '単勝オッズ人気No1馬番号格納エリア
    '2011/11/10 d-kobayashi add start
    Private FukushouMaxNinkiUmaNo As String = String.Empty
    Private EightOddsList As New ArrayList '8時時点オッズ情報リスト
    Private EightMzList As New ArrayList '8時時点マジックゾーンリスト
    Private EightHousokuPointList As New ArrayList '8時時点法則該当リスト
    Private NineOddsList As New ArrayList '9時時点オッズ情報リスト
    Private NineMzList As New ArrayList '9時時点マジックゾーンリスト
    Private NineHousokuPointList As New ArrayList '8時時点法則該当リスト
    Private TenOddsList As New ArrayList '10時時点オッズ情報リスト
    Private TenMzList As New ArrayList '9時時点マジックゾーンリスト
    Private TenHousokuPointList As New ArrayList '10時時点法則該当リスト
    Private HousokuPointList As New ArrayList '馬毎ポイント合計リスト

    Private KeyData As String 'キー情報格納エリア
    Private UpdateOddsList As ArrayList = New ArrayList 'オッズ情報保持ワークリスト
    Private UpdateMzList As ArrayList = New ArrayList 'マジックゾーン情報保持ワークリスト
    Private selTime As String = String.Empty '選択時間保持エリア
    Private selTimeCount As Integer = 0
    Private OpenFile As String() '発表時間選択数保持エリア
    Private EightKeyOdds As String = String.Empty '８時オッズエリアのソートキーオッズ区分
    Private NineKeyOdds As String = String.Empty '９時オッズエリアのソートキーオッズ区分
    Private TenKeyOdds As String = String.Empty '１０時オッズエリアのソートキーオッズ区分


    ' 機能     : メイン画面初期表示処理
    ' 機能説明 : メイン画面起動時の処理を行う
    '
    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim lReturnCode As Long
        Dim strFromTime As String
        Dim strToTime As String
        Dim intUpdateTime As Integer

        Dim dtFromTime As DateTime
        Dim dtToTime As DateTime
        Dim dtNowTime As DateTime

        '********************************************************
        '今後はここのソートキー順を修正するだけで、次期バージョンに対応できます。
        ' バージョンA販売時は、
        '   CommonConstant.TanshouOddsKbn（単勝）
        '   CommonConstant.FukushouOddsKbn（複勝）
        '   CommonConstant.UmatanOddsKbn（馬単）

        ''ソートキーオッズを初期化
        ''８時を単勝オッズ情報でソート
        'EightKeyOdds = CommonConstant.TanshouOddsKbn
        ''９時を複勝オッズ情報でソート
        'NineKeyOdds = CommonConstant.FukushouOddsKbn
        ''１０時を馬単オッズ情報でソート
        'TenKeyOdds = CommonConstant.UmatanOddsKbn
        '********************************************************

        ' JV-Link初期化処理を呼び出す
        lReturnCode = Me.AxJVLink1.JVInit(gJVLinkSID)

        ' 戻り値のエラー判定
        If lReturnCode <> 0 Then
            MsgBox("JV-Link初期化時にエラーが発生しました。" & vbCrLf & _
                   "エラーコード：" & lReturnCode, _
                    MsgBoxStyle.Exclamation, CommonConstant.AppTitle)
            ' アプリケーションを終了する
            Application.Exit()
        End If

        ' コード変換インスタンス生成
        objCodeConv = New clsCodeConv
        ' コードファイルを読み込む
        objCodeConv.FileName = Application.StartupPath & CommonConstant.CodeFile

        ' CommonUtilクラスインスタンス生成
        objCommonUtil = New clsCommonUtil

        ' INIファイルのパス設定
        'gIniFile = Application.StartupPath & CommonConstant.INIFile
        '2011/04/07 d-kobayashi add 
        InitIniFileProc()


        '2011/04/07 d-kobayashi update start
        ' INIファイルが存在するか
        'If Dir(gIniFile) <> "" Then
        '    ' タイマーフラグを取得
        '    gTimerFlg = GetIni(CommonConstant.INI_Ap_Timer, CommonConstant.INI_Key_TimerFlg, "", gIniFile)
        'Else
        '    gTimerFlg = False
        'End If
        'INIファイルの取得を変更
        gTimerFlg = GetIni(CommonConstant.INI_Ap_Timer, CommonConstant.INI_Key_TimerFlg, CommonConstant.INI_TIMERFLG, gIniFile)
        '2011/04/07 d-kobayashi update end

        '' JVLink設定をINIファイルから読み込む
        'strJVFromTime = GetIni(CommonConstant.INI_Ap_JVLink, CommonConstant.INI_Key_JVFromTIme, "", gIniFile)
        'If strJVFromTime = "" Then
        '    strJVFromTime = CommonConstant.DataFromTime
        'End If

        ' タイマーフラグが設定されている場合、タイマー処理を起動する
        If gTimerFlg Then
            '2011/04/07 d-kobayashi update start
            '' 実行する時間帯（FROM）を取得
            'strFromTime = GetIni(CommonConstant.INI_Ap_Timer, CommonConstant.INI_Key_FromTime, "", gIniFile)
            '' 実行する時間帯（TO）を取得
            'strToTime = GetIni(CommonConstant.INI_Ap_Timer, CommonConstant.INI_Key_ToTime, "", gIniFile)
            '' 更新間隔を取得
            'intUpdateTime = GetIni(CommonConstant.INI_Ap_Timer, CommonConstant.INI_Key_UpdateTime, "", gIniFile)
            ' 実行する時間帯（FROM）を取得
            strFromTime = GetIni(CommonConstant.INI_Ap_Timer, CommonConstant.INI_Key_FromTime, Format(Now, "yyyy/MM/dd") & CommonConstant.INI_FROMTIME, gIniFile)
            ' 実行する時間帯（TO）を取得
            strToTime = GetIni(CommonConstant.INI_Ap_Timer, CommonConstant.INI_Key_ToTime, Format(Now, "yyyy/MM/dd") & CommonConstant.INI_TOTIME, gIniFile)
            ' 更新間隔を取得
            intUpdateTime = GetIni(CommonConstant.INI_Ap_Timer, CommonConstant.INI_Key_UpdateTime, CommonConstant.MIN_15, gIniFile)
            '2011/04/07 d-kobayashi update en

            ' 実行する時間帯を調べる
            ' 現在時刻を取得
            dtNowTime = New DateTime(Now.Ticks)
            dtFromTime = DateTime.Parse(strFromTime)
            dtToTime = DateTime.Parse(strToTime)
            Dim strNowMinute As String = IIf(dtNowTime.Minute.ToString.Length = CommonConstant.Const1, _
                                    CommonConstant.Const0 & dtNowTime.Minute, dtNowTime.Minute)
            Dim strFromMinute As String = IIf(dtFromTime.Minute.ToString.Length = CommonConstant.Const1, _
                                    CommonConstant.Const0 & dtFromTime.Minute, dtFromTime.Minute)
            Dim strToMinute As String = IIf(dtToTime.Minute.ToString.Length = CommonConstant.Const1, _
                                    CommonConstant.Const0 & dtToTime.Minute, dtToTime.Minute)
            ' 現在時刻が実行する時間帯のFROM～TO以内なら実行時間とする 
            If (CInt(dtNowTime.Hour & strNowMinute) >= CInt(dtFromTime.Hour & strFromMinute) And _
                (CInt(dtNowTime.Hour & strNowMinute)) <= CInt(dtToTime.Hour & strToMinute)) Then
                'If (dtNowTime.Hour & dtNowTime.Minute >= dtFromTime.Hour & dtFromTime.Minute) And _
                '    (dtNowTime.Hour & dtNowTime.Minute <= dtToTime.Hour & dtToTime.Minute) Then
                gFromTime = Format(dtNowTime.AddMinutes(intUpdateTime), CommonConstant.TimeHHmm)
                gFromTimeLabel = Format(dtNowTime.AddMinutes(intUpdateTime), CommonConstant.TimeHHmm)
                strFromTime = Strings.Left(gFromTimeLabel, CommonConstant.Const2) & ":" & _
                                Strings.Right(gFromTimeLabel, CommonConstant.Const2)
            Else
                strFromTime = FormatDateTime(DateTime.Parse(strFromTime), DateFormat.ShortTime)
            End If

            Me.TimerSetup.Text = "次回更新予定 " & strFromTime & "頃"
            ' タイマー開始
            Me.tmrTimerSet.Enabled = True
        Else
            Me.TimerSetup.Text = "タイマーを設定する"
            ' タイマー終了
            Me.tmrTimerSet.Enabled = False
        End If

        'OS判定
        ' プラットフォームの取得
        Dim os As System.OperatingSystem = System.Environment.OSVersion
        Select Case os.Platform
            Case System.PlatformID.Win32Windows
                GoTo OS_Err
                Exit Select
            Case System.PlatformID.Win32NT
                Select Case os.Version.Major
                    Case 3
                        GoTo OS_Err
                        Exit Select
                    Case 4
                        GoTo OS_Err
                        Exit Select
                    Case 5
                        Select Case os.Version.Minor
                            Case 0
                                GoTo OS_Err
                                Exit Select
                            Case 1
                                ' CSVファイル保存パス設定
                                gPrgFilesPath = System.Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) _
                                                & CommonConstant.JravanPath
                                ' キャッシュファイル保存パス設定
                                gCacheFilePath = System.Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) _
                                                & CommonConstant.JravanCachePath
                                Exit Select
                            Case 2
                                GoTo OS_Err
                                Exit Select
                        End Select
                        Exit Select
                    Case 6
                        Select Case os.Version.Minor
                            Case 0
                                ' CSVファイル保存パス設定
                                gPrgFilesPath = CommonConstant.JravanPath64
                                ' フォルダが存在しな場合は作成する
                                If Not System.IO.Directory.Exists(gPrgFilesPath) Then
                                    System.IO.Directory.CreateDirectory(gPrgFilesPath)
                                End If
                                ' キャッシュファイル保存パス設定
                                gCacheFilePath = CommonConstant.JravanCachePath64
                                ' フォルダが存在しな場合は作成する
                                If Not System.IO.Directory.Exists(gCacheFilePath) Then
                                    System.IO.Directory.CreateDirectory(gCacheFilePath)
                                End If
                                Exit Select
                            Case 1
                                ' CSVファイル保存パス設定
                                gPrgFilesPath = CommonConstant.JravanPath64
                                ' フォルダが存在しな場合は作成する
                                If Not System.IO.Directory.Exists(gPrgFilesPath) Then
                                    System.IO.Directory.CreateDirectory(gPrgFilesPath)
                                End If
                                ' キャッシュファイル保存パス設定
                                gCacheFilePath = CommonConstant.JravanCachePath64
                                ' フォルダが存在しな場合は作成する
                                If Not System.IO.Directory.Exists(gCacheFilePath) Then
                                    System.IO.Directory.CreateDirectory(gCacheFilePath)
                                End If
                                Exit Select
                            Case 2
                                ' CSVファイル保存パス設定
                                gPrgFilesPath = CommonConstant.JravanPath64
                                ' フォルダが存在しな場合は作成する
                                If Not System.IO.Directory.Exists(gPrgFilesPath) Then
                                    System.IO.Directory.CreateDirectory(gPrgFilesPath)
                                End If
                                ' キャッシュファイル保存パス設定
                                gCacheFilePath = CommonConstant.JravanCachePath64
                                ' フォルダが存在しな場合は作成する
                                If Not System.IO.Directory.Exists(gCacheFilePath) Then
                                    System.IO.Directory.CreateDirectory(gCacheFilePath)
                                End If
                                Exit Select
                        End Select
                        Exit Select
                End Select
                Exit Select
        End Select
        ''Vista、7対応
        '' アプリが32ビット、64ビットで動作しているかを判定する
        'If IntPtr.Size = 4 Then
        '    ' 32ビットで動作している
        '    ' CSVファイル保存パス設定
        '    gPrgFilesPath = System.Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) _
        '                    & CommonConstant.JravanPath

        '    ' キャッシュファイル保存パス設定
        '    gCacheFilePath = System.Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) _
        '                    & CommonConstant.JravanCachePath
        'ElseIf IntPtr.Size = 8 Then
        '    ' 64ビットで動作している
        '    ' CSVファイル保存パス設定
        '    gPrgFilesPath = Application.StartupPath & CommonConstant.JravanPath64

        '    ' キャッシュファイル保存パス設定
        '    gCacheFilePath = System.Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) _
        '                    & CommonConstant.JravanCachePath
        'End If
        ' '' CSVファイル保存パス設定
        ''gPrgFilesPath = System.Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) _
        ''                & CommonConstant.JravanPath

        ' '' キャッシュファイル保存パス設定
        ''gCacheFilePath = System.Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) _
        ''                & CommonConstant.JravanCachePath

        ' ファイルオープンフラグを初期化
        gFileOpenFlg = True
        ' レース情報ファイルオープン
        Call objCommonUtil.RaceFileOpen()
        'オッズ情報ファイルオープン
        Call objCommonUtil.OddsFileOpen()


        '' 期限切れデータの削除
        'Call InvalidityDataDelete()


        ''レース情報ファイルオープン
        'If (ComUtil.RaceFileOpen() = False) Then
        '    ''ファイルオープンエラーの場合
        '    'MsgBox("レース情報が存在しません")
        '    'Exit Sub
        'End If

        ''日付のコンボボックスへレース日一覧を設定
        ''ComUtil.setRaceDate(AllRaceInfo, cboDate)
        'ComUtil.setRaceDate(cboDate)

        ' 画面初期化
        ' 画面タイトル
        If gAuthority = CommonConstant.AUTHORITY Then
            Me.Text = CommonConstant.Admin_AppTitle
        Else
            Me.Text = CommonConstant.AppTitle
        End If
        ' 画面サイズを設定する
        Me.Width = 1024
        Me.Height = 738
        ' 画面表示位置を設定する
        Me.StartPosition = FormStartPosition.CenterScreen

        ''テスト用
        '' 予想結果を画面に反映
        'Call Form_ResultReflect()

        '' ダブルバッファリング
        'Me.SetStyle(ControlStyles.DoubleBuffer, True)
        'Me.SetStyle(ControlStyles.UserPaint, True)
        'Me.SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        ''Me.DoubleBuffered = True
        ''Me.SetStyle(ControlStyles.UserPaint, True)
        ''Me.SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        ''Me.SetStyle(ControlStyles.OptimizedDoubleBuffer, True)

        ' 画面を再描画する
        Me.Refresh()

        '過去1か月より前のファイルを削除する
        'If gAuthority <> CommonConstant.AUTHORITY Then
        '    InvalidityDataDelete(3)
        'End If

        'ファイルを開く
        openExcelFile()

        Exit Sub

OS_Err:
        MsgBox("このOSはサポート対象外です。", MsgBoxStyle.Exclamation, CommonConstant.AppTitle)
    End Sub

    ' 機能     : 画面終了時処理
    ' 機能説明 : 画面が閉じられた場合、アプリケーションを終了する
    '
    Private Sub frmMain_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'SplashScreen2.SetInitMessage("過去レース情報を更新中です・・・")
        dlgQuitMessage.lblMessage.Text = "過去レース情報を更新中です・・・"
        Me.Visible = False

        dlgQuitMessage.Show()
        System.Threading.Thread.Sleep(1000) ' ダミー
        Application.DoEvents()
        '過去レース情報の更新を行う
        '2010/11/15 d-kobayashi add
        '構造体をクリアする
        aRaceList = Nothing
        aHorseList = Nothing
        aOdds1List = Nothing
        aOdds2List = Nothing
        aOdds4List = Nothing
        Call dlgLogin.AddOldData()

        ' アプリケーションを終了する
        Application.Exit()
    End Sub

    ' 機能     : JV-Linkメニュー - 「設定」選択時処理
    ' 機能説明 : JV-Link設定画面を表示する
    '
    Private Sub mnuSetting_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSetting.Click
        Try
            Dim lReturnCode As Long
            ' JV-Link設定画面を表示する
            lReturnCode = AxJVLink1.JVSetUIProperties()
            If lReturnCode <> 0 Then
                MsgBox("JV-Link設定画面表示時にエラーが発生しました。" & vbCrLf & _
                       "エラーコード：" & lReturnCode, _
                       MsgBoxStyle.Exclamation, CommonConstant.AppTitle)
            End If
        Catch ex As Exception
        End Try
    End Sub

    ' 機能     : JVLinkメニュー - 「キャッシュクリア」選択時処理
    ' 機能説明 : JVLinkのキャッシュをクリアする
    '
    Private Sub mnuCacheClr_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuCacheClr.Click
        ' JVDファイルのファイルリストを取得する
        Dim OpenJvdFile As String() = System.IO.Directory.GetFiles(gCacheFilePath & "\", "*.jvd")
        ' RTDファイルのファイルリストを取得する
        Dim OpenRtdFile As String() = System.IO.Directory.GetFiles(gCacheFilePath & "\", "*.rtd")
        Try

        
            ' 2010/08/20 記述
            ' キャッシュクリア
            'jvdファイル削除
            If OpenJvdFile.Length > CommonConstant.Const0 Then

                For i = 0 To OpenJvdFile.Length - 1
                    Dim delFile As String = OpenJvdFile(i)
                    System.IO.File.Delete(delFile)
                Next

            End If

            'rtdファイル削除
            If OpenRtdFile.Length > CommonConstant.Const0 Then

                For i = 0 To OpenRtdFile.Length - 1
                    Dim delFile As String = OpenRtdFile(i)
                    System.IO.File.Delete(delFile)
                Next

            End If

            MsgBox("キャッシュファイルを削除しました。")

        Catch ex As Exception
            MsgBox("キャッシュファイルを削除しました。", MsgBoxStyle.Exclamation, CommonConstant.AppTitle)

        End Try

    End Sub

    ' 機能     : ヘルプメニュー - 「使い方」選択時処理
    ' 機能説明 : 使い方を開く（URL参照）
    '
    Private Sub mnuManual_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' 使い方のWebページを標準のブラウザで開いて表示する
        System.Diagnostics.Process.Start(CommonConstant.HelpURL)
    End Sub

    '' 機能     : ファイルメニュー - 「画面キャプチャ」選択時処理
    '' 機能説明 : アクティブ画面のキャプチャを行う
    ''
    'Private Sub mnuCapture_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    '' 画面全体のイメージをクリップボードにコピーする
    '    'SendKeys.SendWait("^{PRTSC}")

    '    ' アクティブなウィンドウのイメージをクリップボードにコピーする
    '    SendKeys.SendWait("{PRTSC}")
    '    SendKeys.SendWait("%{PRTSC}")
    '    Application.DoEvents()
    'End Sub

    ' 機能     : ファイルメニュー - 「終了」選択時処理
    ' 機能説明 : アプリケーションを終了する
    '
    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        ' アプリケーションを終了する
        Application.Exit()
    End Sub

    ' 機能     : データ更新タブ - データを更新ボタン押下時処理
    ' 機能説明 : JVデータのダウンロード、CSVファイル出力処理を行う
    '
    Private Sub btnDataUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDataUpdate.Click
        Dim i, j As Integer
        Dim lReturnCode As Long
        Dim strDataSpec As String           ' 引数 JVOpen:ファイル識別子
        Dim strKeyFromTime As String        ' 引数 JVOpen:データ提供日付
        Dim strRaseInfo() As String         ' レース詳細レコード
        Dim RaseInfoList As New ArrayList   ' レース詳細レコードリスト
        Dim SearchFileList As New ArrayList ' 見つかったファイルリスト
        '2011/03/28 d-kobayashi add start
        'リトライカウンタ(5回)
        Const intRetryCounter = 4
        Dim k As Integer
        '2011/03/28 d-kobayashi add end


        ' フォームを無効にする
        Me.Enabled = False
        ' データ更新ボタンを無効にする
        Me.btnDataUpdate.Enabled = False
        ' マウスカーソルを砂時計にする
        Me.Cursor = Cursors.WaitCursor

        ' リストボックスに作業状況を表示する
        Me.lstUpdateInfo.Items.Clear()
        Me.lstUpdateInfo.Items.Add("レースデータを更新します...")
        Me.lstUpdateInfo.Items.Add("データのダウンロードを開始します...")
        Me.lstUpdateInfo.SelectedIndex = Me.lstUpdateInfo.Items.Count - 1

        '' オッズ情報ファイルオープン
        'Call ComUtil.OddsFileOpen()

        ' 蓄積系JVデータ取得時の初期設定
        strDataSpec = CommonConstant.DataSpec       'レース詳細
        strJVFromTime = CommonConstant.DataFromTime '取得開始日付
        'If strJVFromTime = "" Then
        '    strJVFromTime = CommonConstant.DataFromTime
        'End If

        ' 蓄積系JVデータ取得処理を呼び出す
        '2011/03/28 d-kobayashi update start
        'エラー発生時、リトライを行う。
        For i = 0 To intRetryCounter
            lReturnCode = GetJVData(strDataSpec, strJVFromTime)
            If (lReturnCode = CommonConstant.ReturnCd_0 _
            Or lReturnCode = CommonConstant.ReturnCd_1 _
            Or lReturnCode = CommonConstant.ReturnCd_504) Then Exit For
        Next
        '2011/03/28 d-kobayashi update end

        If lReturnCode = CommonConstant.ReturnCd_504 Then
            GoTo JVLink_MenteEnd
        ElseIf lReturnCode < 0 Then
            If lReturnCode <> CommonConstant.ReturnCd_1 And _
                lReturnCode <> CommonConstant.ReturnCd_504 Then
                MsgBox("JVデータ取得処理でエラーが発生しました。" & vbCrLf & _
                        "システム管理者に連絡してください。" & vbCrLf & vbCrLf & _
                        "エラーコード：" & lReturnCode, _
                        MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, CommonConstant.AppTitle)
                GoTo JVLink_Err
            End If
        End If
        'If lReturnCode = CommonConstant.ReturnCd_504 Then
        '    GoTo JVLink_MenteEnd
        'ElseIf lReturnCode = CommonConstant.ReturnCd_1 Then
        '    GoTo JVLink_UpdateEnd
        'ElseIf lReturnCode < 0 Then
        '    MsgBox("JVデータ取得処理でエラーが発生しました。" & vbCrLf & _
        '           "システム管理者に連絡してください。" & vbCrLf & vbCrLf & _
        '           "エラーコード：" & lReturnCode, _
        '           MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, CommonConstant.AppTitle)
        '    GoTo JVLink_Err
        'End If

        ' 蓄積系JVデータのCSVファイルが存在するか検索する
        Call GetAllFiles(gPrgFilesPath, CommonConstant.RaseID & _
                         CommonConstant.ASTERISK & _
                         CommonConstant.CSV, SearchFileList)

        ' 蓄積系JVデータのCSVファイルが存在する場合
        If SearchFileList.Count > 0 Then
            ' 変数初期化
            strRAFileName = ""
            strSEFileName = ""
            strO1FileName = ""
            strO2FileName = ""
            strO4FileName = ""
            ' リスト初期化
            aRaceList = New ArrayList
            aHorseList = New ArrayList
            aOdds1List = New ArrayList
            aOdds2List = New ArrayList
            aOdds4List = New ArrayList

            For i = 0 To SearchFileList.Count - 1
                ' レース詳細CSVファイルを読み込む
                objCommonUtil.setFileToList(SearchFileList.Item(i), RaseInfoList)

                ' 作業状況を更新する
                Me.lstUpdateInfo.Items.Add("速報系データを読み込んでいます...")
                Me.lstUpdateInfo.SelectedIndex = Me.lstUpdateInfo.Items.Count - 1

                ' レース詳細レコードの件数分、処理を繰り返す
                For j = 0 To RaseInfoList.Count - 1
                    ' レース詳細データからキー情報を取得し、速報系JVデータを取得する
                    strRaseInfo = RaseInfoList.Item(j)

                    ' キー情報を取得する
                    strKeyFromTime = strRaseInfo(CommonConstant.IndexPos0).ToString
                    '' キー情報から先頭12桁（YYYYMMDDJJRR）を取得する
                    'strFromTime = Mid(strRaseInfo(0), 1, 12)

                    ' 速報系JVデータ取得時の設定（単複枠取得モード）
                    strDataSpec = CommonConstant.RtDataSpec_O1
                    ' 速報オッズ（単複枠）取得処理を呼び出す
                    For k = 0 To intRetryCounter
                        lReturnCode = GetRtJVData(strDataSpec, strKeyFromTime)
                        If (lReturnCode = CommonConstant.ReturnCd_0 _
                        Or lReturnCode = CommonConstant.ReturnCd_1 _
                        Or lReturnCode = CommonConstant.ReturnCd_504) Then Exit For
                    Next

                    If lReturnCode = CommonConstant.ReturnCd_504 Then
                        GoTo JVLink_MenteEnd
                    ElseIf lReturnCode < 0 Then
                        If lReturnCode <> CommonConstant.ReturnCd_1 And _
                            lReturnCode <> CommonConstant.ReturnCd_504 Then
                            MsgBox("速報オッズ（単複枠）取得処理でエラーが発生しました。" & vbCrLf & _
                                       "システム管理者に連絡してください。" & vbCrLf & vbCrLf & _
                                       "エラーコード：" & lReturnCode, _
                                       MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, CommonConstant.AppTitle)
                            GoTo JVLink_Err
                        End If
                    End If

                    ' 速報系JVデータ取得時の設定（馬連取得モード）
                    strDataSpec = CommonConstant.RtDataSpec_O2
                    ' 速報オッズ（馬連）取得処理を呼び出す
                    '2011/03/28 d-kobayashi update start
                    'エラー時、リトライを行う。
                    For k = 0 To intRetryCounter
                        lReturnCode = GetRtJVData(strDataSpec, strKeyFromTime)
                        If (lReturnCode = CommonConstant.ReturnCd_0 _
                        Or lReturnCode = CommonConstant.ReturnCd_1 _
                        Or lReturnCode = CommonConstant.ReturnCd_504) Then Exit For
                    Next
                    '2011/03/28 d-kobayashi update end

                    If lReturnCode = CommonConstant.ReturnCd_504 Then
                        GoTo JVLink_MenteEnd
                    ElseIf lReturnCode < 0 Then
                        If lReturnCode <> CommonConstant.ReturnCd_1 And _
                            lReturnCode <> CommonConstant.ReturnCd_504 Then
                            MsgBox("速報オッズ（馬連）取得処理でエラーが発生しました。" & vbCrLf & _
                                       "システム管理者に連絡してください。" & vbCrLf & vbCrLf & _
                                       "エラーコード：" & lReturnCode, _
                                       MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, CommonConstant.AppTitle)
                            GoTo JVLink_Err
                        End If
                    End If

                    ' 速報系JVデータ取得時の設定（馬単取得モード）
                    strDataSpec = CommonConstant.RtDataSpec_O4
                    ' 速報オッズ（馬単）取得処理を呼び出す
                    '2011/03/28 d-kobayashi update start
                    'エラー時、リトライを行う。
                    For k = 0 To intRetryCounter
                        lReturnCode = GetRtJVData(strDataSpec, strKeyFromTime)
                        If (lReturnCode = CommonConstant.ReturnCd_0 _
                        Or lReturnCode = CommonConstant.ReturnCd_504 _
                        Or lReturnCode = CommonConstant.ReturnCd_1) Then Exit For
                    Next
                    '2011/03/28 d-kobayashi update end
                    If lReturnCode = CommonConstant.ReturnCd_504 Then
                        GoTo JVLink_MenteEnd
                    ElseIf lReturnCode < 0 Then
                        If lReturnCode <> CommonConstant.ReturnCd_1 And _
                            lReturnCode <> CommonConstant.ReturnCd_504 Then
                            MsgBox("速報オッズ（馬単）取得処理でエラーが発生しました。" & vbCrLf & _
                                       "システム管理者に連絡してください。" & vbCrLf & vbCrLf & _
                                       "エラーコード：" & lReturnCode, _
                                       MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, CommonConstant.AppTitle)
                            GoTo JVLink_Err
                        End If
                    End If
                Next
            Next

            ' 速報系JVデータCSV出力処理を呼び出す
            '2011/03/28 d-kobayashi update start
            'エラー時、リトライを行う
            For k = 0 To intRetryCounter
                lReturnCode = JVData_CSVOutput()
                If (lReturnCode = CommonConstant.ReturnCd_0 _
                Or lReturnCode = CommonConstant.ReturnCd_1 _
                Or lReturnCode = CommonConstant.ReturnCd_504) Then Exit For
            Next
            '2011/03/28 d-kobayashi update end
            If lReturnCode < 0 Then
                MsgBox("JVデータ作成処理でエラーが発生しました。" & vbCrLf & _
                       "システム管理者に連絡してください。" & vbCrLf & vbCrLf & _
                       "エラーコード：" & lReturnCode, _
                       MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, CommonConstant.AppTitle)
                GoTo JVLink_Err
            End If
        End If

        ' Downloadタイマー有効時は、無効化する
        If tmrDownload.Enabled = True Then
            tmrDownload.Enabled = False
        End If

        ' 作業状況を更新する
        prgDownload.Value = prgDownload.Maximum
        If Not gFileOpenFlg Then
            Me.lstUpdateInfo.Items.Add("レース情報を読み込んでいます...")
            Me.lstUpdateInfo.SelectedIndex = Me.lstUpdateInfo.Items.Count - 1
            ' 最新レース情報ファイルオープン
            Call objCommonUtil.RaceFileOpen()
            Me.lstUpdateInfo.Items.Add("オッズ情報を読み込んでいます...")
            Me.lstUpdateInfo.SelectedIndex = Me.lstUpdateInfo.Items.Count - 1
            ' 最新オッズ情報ファイルオープン
            Call objCommonUtil.OddsFileOpen()
        End If
        Me.lstUpdateInfo.Items.Add("データの更新が完了しました...")
        Me.lstUpdateInfo.SelectedIndex = Me.lstUpdateInfo.Items.Count - 1

        ' データ更新ボタンを有効にする
        Me.btnDataUpdate.Enabled = True
        ' フォームを有効にする
        Me.Enabled = True
        ' マウスカーソルを標準に戻す
        Me.Cursor = Cursors.Default

        Exit Sub

JVLink_Err:
        ' Downloadタイマー有効時は、無効化する
        If tmrDownload.Enabled = True Then
            tmrDownload.Enabled = False
        End If

        ' 作業状況を更新する
        prgDownload.Value = prgDownload.Maximum
        Me.lstUpdateInfo.Items.Add("データ更新時にエラーが発生しました...")
        Me.lstUpdateInfo.SelectedIndex = Me.lstUpdateInfo.Items.Count - 1

        ' データ更新ボタンを有効にする
        Me.btnDataUpdate.Enabled = True
        ' フォームを有効にする
        Me.Enabled = True
        ' マウスカーソルを標準に戻す
        Me.Cursor = Cursors.Default

        Exit Sub

JVLink_MenteEnd:
        ' Downloadタイマー有効時は、無効化する
        If tmrDownload.Enabled = True Then
            tmrDownload.Enabled = False
        End If

        ' 作業状況を更新する
        prgDownload.Value = prgDownload.Maximum
        Me.lstUpdateInfo.Items.Add("サーバメンテナンス中のためデータ更新を終了します...")
        Me.lstUpdateInfo.SelectedIndex = Me.lstUpdateInfo.Items.Count - 1

        ' データ更新ボタンを有効にする
        Me.btnDataUpdate.Enabled = True
        ' フォームを有効にする
        Me.Enabled = True
        ' マウスカーソルを標準に戻す
        Me.Cursor = Cursors.Default

        Exit Sub

JVLink_UpdateEnd:
        ' Downloadタイマー有効時は、無効化する
        If tmrDownload.Enabled = True Then
            tmrDownload.Enabled = False
        End If

        ' 作業状況を更新する
        prgDownload.Value = prgDownload.Maximum
        Me.lstUpdateInfo.Items.Add("最新のJVデータが存在しない、又はJVLinkのダウンロードが選択されたためデータ更新を終了します...")
        Me.lstUpdateInfo.SelectedIndex = Me.lstUpdateInfo.Items.Count - 1

        ' データ更新ボタンを有効にする
        Me.btnDataUpdate.Enabled = True
        ' フォームを有効にする
        Me.Enabled = True
        ' マウスカーソルを標準に戻す
        Me.Cursor = Cursors.Default

        Exit Sub
    End Sub

    ' 機能　　 : 蓄積系JVデータ取得処理
    ' 引き数　 : strDataSpec - データ取得モード
    '            strFromTime - データ取得開始日時
    ' 返り値　 : Long - 処理結果
    ' 機能説明 : 蓄積系のJVデータをダウンロードし、CSVファイルに出力する
    '
    Private Function GetJVData(ByVal strDataSpec As String, ByVal strFromTime As String) As Long
        Dim lReturnCode As Long
        Dim lOption As Long                 ' 引数 JVOpen:オプション
        Dim lReadCount As Long              ' JVLink 戻り値
        Dim strLastFileTimestamp As String  ' JVOpen: 最新ファイルのタイムスタンプ
        Const lBuffSize As Long = 102901    ' JVRead:データ格納バッファサイズ
        Const lNameSize As Integer = 256    ' JVRead:ファイル名サイズ
        Dim strBuff As String               ' JVRead:データ格納バッファ
        Dim strFileName As String           ' JVRead:ダウンロードファイル名
        Dim csvOutput As New clsCSVOutput() ' CSV出力クラス
        Dim strCSVFileName As String        ' CSVファイル名
        Dim SearchFileList As New ArrayList ' 見つかったファイルリスト

        ' 各種フラグ
        Dim blnDlFlg As Boolean
        Dim blnRaFlg As Boolean
        Dim blnSeFlg As Boolean
        Dim blnO1Flg As Boolean
        Dim blnO2Flg As Boolean
        Dim blno4Flg As Boolean
        Dim blnErrFlg As Boolean

        GetJVData = CommonConstant.ReturnCd_0

        Try
            ' フラグ初期化
            blnDlFlg = False
            blnRaFlg = False
            blnSeFlg = False
            blnO1Flg = False
            blnO2Flg = False
            blno4Flg = False
            ' 変数初期化
            strLastFileTimestamp = ""
            strRAFileName = ""
            strSEFileName = ""
            strO1FileName = ""
            strO2FileName = ""
            strO4FileName = ""

            ' 進捗表示初期設定
            tmrDownload.Enabled = False         'タイマー停止
            prgDownload.Value = 0               'JVData ダウンロード進捗

            ' JVLinkオプション設定
            lOption = CommonConstant.WeekMode   '今週データモード
            
            ' JVLink - JVデータオープン処理を実行する（JVデータダウンロード）
            lReturnCode = Me.AxJVLink1.JVOpen(strDataSpec, strFromTime, lOption, _
                            lReadCount, lDownloadCount, strLastFileTimestamp)

            ' JVデータオープン処理の結果判定
            Select Case lReturnCode
                Case CommonConstant.ReturnCd_0
                    ' 続行
                Case CommonConstant.ReturnCd_1
                    ''続行
                    'MsgBox("データは最新です。", MsgBoxStyle.Information, CommonConstant.AppTitle)
                    'prgDownload.Value = prgDownload.Maximum
                    'Me.lstUpdateInfo.Items.Add("データの更新が完了しました...")
                    'Me.lstUpdateInfo.SelectedIndex = Me.lstUpdateInfo.Items.Count - 1
                    Me.AxJVLink1.JVClose()
                    GetJVData = lReturnCode
                    Exit Function
                Case CommonConstant.ReturnCd_504    ' メンテナンス中
                    ' メンテナンス中のメッセージはJVLinkが表示するので
                    ' 何も表示しないで終了する。
                    Me.AxJVLink1.JVClose()
                    GetJVData = lReturnCode
                    Exit Function
                Case Else
                    'MsgBox("JV-Linkへ接続失敗しました。" & vbCrLf _
                    '    & "JV-Linkからのエラーメッセージ: " & vbCrLf _
                    '    & ErrMsgJVOpen(lReturnCode), vbInformation, CommonConstant.AppTitle)
                    Me.AxJVLink1.JVClose()
                    GetJVData = lReturnCode
                    Exit Function
            End Select

            ' 進捗表示プログレスバー最大値を設定
            If lDownloadCount = 0 Then
                prgDownload.Maximum = lReadCount
                tmrDownload.Enabled = True  ' タイマー開始
            Else
                prgDownload.Maximum = lDownloadCount + lReadCount
                tmrDownload.Enabled = True  ' タイマー開始
            End If

            ' JVデータが1件でも存在する場合
            If lReadCount > 0 Then
                blnErrFlg = False
                Do
                    ' バックグラウンドでの処理を実行
                    Application.DoEvents()
                    ' JVデータを格納するバッファ作成
                    strBuff = New String(vbNullChar, lBuffSize)
                    strFileName = New String(vbNullChar, lNameSize)
                    ' JVLink - JVデータ読み込み処理を実行する
                    lReturnCode = Me.AxJVLink1.JVRead(strBuff, lBuffSize, strFileName)
                    ' JVデータ読み込み処理により処理を分枝
                    Select Case lReturnCode
                        Case CommonConstant.ReturnCd_0          ' 全ファイル読み込み終了
                            prgDownload.Value = prgDownload.Maximum     ' 進捗表示
                            Exit Do
                        Case CommonConstant.ReturnCd_1          ' ファイル切り替わり
                            prgDownload.Value = prgDownload.Value + 1
                        Case CommonConstant.ReturnCd_3          ' ダウンロード中
                            Me.lstUpdateInfo.Items.Add("レースデータをダウンロードしています...")
                            Me.lstUpdateInfo.SelectedIndex = Me.lstUpdateInfo.Items.Count - 1
                        Case CommonConstant.ReturnCd_201        ' Init されてない
                            GetJVData = lReturnCode
                            GoTo JVLink_JVClose
                            'MsgBox("JVInit が行われていません。", MsgBoxStyle.Exclamation, CommonConstant.AppTitle)
                            'blnErrFlg = True
                            'Exit Do
                        Case CommonConstant.ReturnCd_203        ' Open されてない
                            GetJVData = lReturnCode
                            GoTo JVLink_JVClose
                            'MsgBox("JVOpen が行われていません。", MsgBoxStyle.Exclamation, CommonConstant.AppTitle)
                            'blnErrFlg = True
                            'Exit Do
                        Case CommonConstant.ReturnCd_502        ' ダウンロード失敗
                            GetJVData = lReturnCode
                            GoTo JVLink_JVClose
                            'MsgBox("ダウンロード中にエラーが発生しました。", MsgBoxStyle.Exclamation, CommonConstant.AppTitle)
                            'blnErrFlg = True
                            'Exit Do
                        Case CommonConstant.ReturnCd_503        ' ファイルがない
                            GetJVData = lReturnCode
                            GoTo JVLink_JVClose
                            'MsgBox(strFileName & " ファイルが存在しません。", MsgBoxStyle.Exclamation, CommonConstant.AppTitle)
                            'blnErrFlg = True
                            'Exit Do
                        Case Is > CommonConstant.ReturnCd_0     ' 正常読み込み
                            ' レコード種別IDの識別
                            If Mid(strBuff, 1, 2) = CommonConstant.RaseID Then
                                ' 初回処理された場合のみ作業状況を更新する
                                If Not blnRaFlg Then
                                    Me.lstUpdateInfo.Items.Add("レース詳細を読み込んでいます...")
                                    Me.lstUpdateInfo.SelectedIndex = Me.lstUpdateInfo.Items.Count - 1
                                    blnRaFlg = True
                                End If

                                ' 最新のファイル名を保持する
                                strRAFileName = strFileName
                                ' レース詳細リスト作成処理を呼び出す
                                csvOutput.RaceInfoMakeList(strBuff, aRaceList)
                            ElseIf Mid(strBuff, 1, 2) = CommonConstant.HorseID Then
                                ' 初回処理された場合のみ作業状況を更新する
                                If Not blnSeFlg Then
                                    Me.lstUpdateInfo.Items.Add("馬毎レース情報を読み込んでいます...")
                                    Me.lstUpdateInfo.SelectedIndex = Me.lstUpdateInfo.Items.Count - 1
                                    blnSeFlg = True
                                End If

                                ' 最新のファイル名を保持する
                                strSEFileName = strFileName
                                ' 馬毎レース情報リスト作成処理を呼び出す
                                csvOutput.HorseMakeList(strBuff, aHorseList)
                            End If
                    End Select
                Loop While (1)

                Application.DoEvents()

                ' JVデータ最終取得時間をINIファイルに保存
                PutIni(CommonConstant.INI_Ap_JVLink, CommonConstant.INI_Key_JVFromTime, _
                       strLastFileTimestamp, gIniFile)

                '' JVデータ読み込み時にエラーが発生した場合
                'If blnErrFlg Then
                '    GetJVData = False
                '    Exit Function
                'End If

                ' レース詳細リストが存在する場合
                If aRaceList.Count <> 0 Then
                    ' 作業状況を更新する
                    Me.lstUpdateInfo.Items.Add("レース詳細を設定しています...")
                    Me.lstUpdateInfo.SelectedIndex = Me.lstUpdateInfo.Items.Count - 1

                    ' レース詳細CSVファイル名を設定
                    strCSVFileName = CommonConstant.RA_FileName
                    'strCSVFileName = Mid(strRAFileName, 1, InStr(strRAFileName, ".") - 1) & CommonConstant.CSV

                    '' すでにファイルが存在するか検索する
                    'Call GetAllFiles(gPrgFilesPath, CommonConstant.RaseID & _
                    '                 CommonConstant.ASTERISK & _
                    '                 CommonConstant.CSV, SearchFileList)

                    '' レース詳細CSVファイルが存在した場合
                    'If SearchFileList.Count > 0 Then
                    '    For i = 0 To SearchFileList.Count - 1
                    '        ' 現在のファイルと異なるか判定
                    '        If gPrgFilesPath & "\" & strCSVFileName <> _
                    '            SearchFileList.Item(i).ToString Then
                    '            ' 異なる場合、現在のファイル名に"old_"を付ける
                    '            File.Move(SearchFileList.Item(i).ToString, _
                    '                      gPrgFilesPath & "\" & CommonConstant.OLDCSV & _
                    '                      Path.GetFileName(SearchFileList.Item(i).ToString))
                    '        End If
                    '    Next
                    '    ' ファイルリスト初期化
                    '    SearchFileList = New ArrayList
                    'End If

                    ' レース詳細CSVファイル出力処理を呼び出す
                    csvOutput.RaceInfoOutput(strCSVFileName, gPrgFilesPath, aRaceList)
                End If

                ' 馬毎レース情報リストが存在する場合
                If aHorseList.Count <> 0 Then
                    ' 作業状況を更新する
                    Me.lstUpdateInfo.Items.Add("馬毎レース情報を設定しています...")
                    Me.lstUpdateInfo.SelectedIndex = Me.lstUpdateInfo.Items.Count - 1

                    ' 馬毎レース情報CSVファイル名を設定
                    strCSVFileName = CommonConstant.SE_FileName
                    'strCSVFileName = Mid(strSEFileName, 1, InStr(strSEFileName, ".") - 1) & CommonConstant.CSV

                    '' すでにファイルが存在するか検索する
                    'Call GetAllFiles(gPrgFilesPath, CommonConstant.HorseID & _
                    '                 CommonConstant.ASTERISK & _
                    '                 CommonConstant.CSV, SearchFileList)

                    '' 馬毎レース情報CSVファイルが存在した場合
                    'If SearchFileList.Count > 0 Then
                    '    For i = 0 To SearchFileList.Count - 1
                    '        ' 現在のファイルと異なるか判定
                    '        If gPrgFilesPath & "\" & strCSVFileName <> _
                    '            SearchFileList.Item(i).ToString Then
                    '            ' 異なる場合、現在のファイル名に"old_"を付ける
                    '            File.Move(SearchFileList.Item(i).ToString, _
                    '                      gPrgFilesPath & "\" & CommonConstant.OLDCSV & _
                    '                      Path.GetFileName(SearchFileList.Item(i).ToString))
                    '        End If
                    '    Next
                    '    ' ファイルリスト初期化
                    '    SearchFileList = New ArrayList
                    'End If

                    ' 馬毎レース情報CSVファイル出力処理を呼び出す
                    csvOutput.SeInfoOutput(strCSVFileName, gPrgFilesPath, aHorseList)
                End If
            End If
        Catch
            Debug.WriteLine(Err.Description)
            GetJVData = Err.Number
        End Try

        GoTo JVLink_JVClose

JVLink_JVClose:
        ' JVLink - 終了処理を実行する
        lReturnCode = Me.AxJVLink1.JVClose()
        If lReturnCode <> 0 Then
            'MsgBox("JVClose エラー：" & lReturnCode, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
            GetJVData = lReturnCode
        End If
    End Function


    ' 機能　　 : 速報系JVデータ取得処理（JVデータリストのみ作成）
    ' 引き数　 : strDataSpec - データ取得モード
    '            strFromTime - データ取得開始日時
    ' 返り値　 : Long - 処理結果
    ' 機能説明 : 速報系JVデータをダウンロードし、データをリストに追加する
    '
    Private Function GetRtJVData(ByVal strDataSpec As String, ByVal strFromTime As String) As Long
        Dim lReturnCode As Long
        'Dim lOption As Long                 ' 引数 JVOpen:オプション
        Dim lReadCount As Long              ' JVLink 戻り値
        Dim strLastFileTimestamp As String  ' JVOpen: 最新ファイルのタイムスタンプ
        Const lBuffSize As Long = 102901    ' JVRead:データ格納バッファサイズ
        Const lNameSize As Integer = 256    ' JVRead:ファイル名サイズ
        Dim strBuff As String               ' JVRead:データ格納バッファ
        Dim strFileName As String           ' JVRead:ダウンロードファイル名
        Dim csvOutput As New clsCSVOutput() ' CSV出力クラス
        Dim SearchFileList As New ArrayList ' 見つかったファイルリスト

        ' 各種フラグ
        Dim blnDlFlg As Boolean
        Dim blnRaFlg As Boolean
        Dim blnSeFlg As Boolean
        Dim blnO1Flg As Boolean
        Dim blnO2Flg As Boolean
        Dim blno4Flg As Boolean
        Dim blnErrFlg As Boolean

        GetRtJVData = CommonConstant.ReturnCd_0

        Try
            ' フラグ初期化
            blnDlFlg = False
            blnRaFlg = False
            blnSeFlg = False
            blnO1Flg = False
            blnO2Flg = False
            blno4Flg = False
            ' 変数初期化
            strLastFileTimestamp = ""

            ' 進捗表示初期設定
            tmrDownload.Enabled = False         'タイマー停止
            prgDownload.Value = 0               'JVData ダウンロード進捗

            ' 引数設定
            'lOption = CommonConstant.WeekMode       '今週データモード

            ' JVLink - JV速報データオープン処理を実行する（JVRTデータダウンロード）
            lReturnCode = Me.AxJVLink1.JVRTOpen(strDataSpec, strFromTime)

            ' JVデータオープン処理の結果判定
            Select Case lReturnCode
                Case CommonConstant.ReturnCd_0
                    ' 続行
                Case CommonConstant.ReturnCd_1
                    prgDownload.Maximum = prgDownload.Maximum
                    '' 該当データなし
                    'Me.lstUpdateInfo.Items.Add("速報オッズが存在しません...")
                    'Me.lstUpdateInfo.SelectedIndex = Me.lstUpdateInfo.Items.Count - 1
                    Me.AxJVLink1.JVClose()
                    GetRtJVData = lReturnCode
                    Exit Function
                Case CommonConstant.ReturnCd_504    ' メンテナンス中
                    ' メンテナンス中のメッセージはJVLinkが表示するので
                    ' 何も表示しないで終了する。
                    Me.AxJVLink1.JVClose()
                    GetRtJVData = lReturnCode
                    Exit Function
                Case Else
                    'MsgBox("JV-Linkへ接続失敗しました。" & vbCrLf _
                    '    & "JV-Linkからのエラーメッセージ: " & vbCrLf _
                    '    & ErrMsgJVOpen(lReturnCode), vbInformation, CommonConstant.AppTitle)
                    Me.AxJVLink1.JVClose()
                    GetRtJVData = lReturnCode
                    Exit Function
            End Select

            ' 進捗表示プログレスバー最大値を設定
            If lDownloadCount = 0 Then
                prgDownload.Maximum = lReadCount
                tmrDownload.Enabled = True  ' タイマー開始
            Else
                prgDownload.Maximum = lDownloadCount + lReadCount
                tmrDownload.Enabled = True  ' タイマー開始
            End If

            blnErrFlg = False
            Do
                ' バックグラウンドでの処理を実行
                Application.DoEvents()
                ' JVデータを格納するバッファ作成
                strBuff = New String(vbNullChar, lBuffSize)
                strFileName = New String(vbNullChar, lNameSize)
                ' JVLink - JVデータ読み込み処理を実行する
                lReturnCode = Me.AxJVLink1.JVRead(strBuff, lBuffSize, strFileName)
                ' JVデータ読み込み処理により処理を分枝
                Select Case lReturnCode
                    Case CommonConstant.ReturnCd_0          ' 全ファイル読み込み終了
                        prgDownload.Value = prgDownload.Maximum     ' 進捗表示
                        Exit Do
                    Case CommonConstant.ReturnCd_1          ' ファイル切り替わり
                        prgDownload.Value = prgDownload.Value + 1
                    Case CommonConstant.ReturnCd_3          ' ダウンロード中
                        Me.lstUpdateInfo.Items.Add("速報レースデータをダウンロードしています...")
                        Me.lstUpdateInfo.SelectedIndex = Me.lstUpdateInfo.Items.Count - 1
                    Case CommonConstant.ReturnCd_201        ' Init されてない
                        GetRtJVData = lReturnCode
                        GoTo JVLink_JVClose
                        'MsgBox("JVInit が行われていません。", MsgBoxStyle.Exclamation, CommonConstant.AppTitle)
                        'blnErrFlg = True
                        'Exit Do
                    Case CommonConstant.ReturnCd_203        ' Open されてない
                        GetRtJVData = lReturnCode
                        GoTo JVLink_JVClose
                        'MsgBox("JVOpen が行われていません。", MsgBoxStyle.Exclamation, CommonConstant.AppTitle)
                        'blnErrFlg = True
                        'Exit Do
                    Case CommonConstant.ReturnCd_502        ' ダウンロード失敗
                        GetRtJVData = lReturnCode
                        GoTo JVLink_JVClose
                        'MsgBox("ダウンロード中にエラーが発生しました。", MsgBoxStyle.Exclamation, CommonConstant.AppTitle)
                        'blnErrFlg = True
                        'Exit Do
                    Case CommonConstant.ReturnCd_503        ' ファイルがない
                        GetRtJVData = lReturnCode
                        GoTo JVLink_JVClose
                        'MsgBox(strFileName & " ファイルが存在しません。", MsgBoxStyle.Exclamation, CommonConstant.AppTitle)
                        'blnErrFlg = True
                        'Exit Do
                    Case Is > CommonConstant.ReturnCd_0     ' 正常読み込み
                        ' レコード種別IDの識別
                        If Mid(strBuff, 1, 2) = CommonConstant.RaseID Then
                            ' レース詳細リスト作成処理を呼び出す
                            csvOutput.RaceInfoMakeList(strBuff, aRaceList)
                        ElseIf Mid(strBuff, 1, 2) = CommonConstant.HorseID Then
                            ' 馬毎レース情報リスト作成処理を呼び出す
                            csvOutput.HorseMakeList(strBuff, aHorseList)
                        ElseIf Mid(strBuff, 1, 2) = CommonConstant.Odds1ID Then
                            ' オッズ（単複枠）情報リスト作成処理を呼び出す
                            csvOutput.Odds1MakeList(strBuff, aOdds1List)
                        ElseIf Mid(strBuff, 1, 2) = CommonConstant.Odds2ID Then
                            ' オッズ（馬連）情報リスト作成処理を呼び出す
                            csvOutput.Odds2MakeList(strBuff, aOdds2List)
                        ElseIf Mid(strBuff, 1, 2) = CommonConstant.Odds4ID Then
                            ' オッズ（馬単）情報リスト作成処理を呼び出す
                            csvOutput.Odds4MakeList(strBuff, aOdds4List)
                        End If
                End Select
            Loop While (1)

            Application.DoEvents()
        Catch
            Debug.WriteLine(Err.Description)
            GetRtJVData = Err.Number
        End Try

        GoTo JVLink_JVClose

JVLink_JVClose:
        ' JVLink - 終了処理を実行する
        lReturnCode = Me.AxJVLink1.JVClose()
        If lReturnCode <> 0 Then
            'MsgBox("JVClose エラー：" & lReturnCode, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
            GetRtJVData = lReturnCode
        End If
    End Function


    ' 機能　　 : 速報系JVデータCSVファイル出力処理
    ' 引数　　 : なし
    ' 返り値　 : Long - 処理結果
    ' 機能説明 : 速報系JVデータリストを読み込み、CSVファイルに出力する
    '
    Private Function JVData_CSVOutput() As Long
        Dim strLastFileTimestamp As String  ' JVOpen: 最新ファイルのタイムスタンプ
        Dim csvOutput As New clsCSVOutput() ' CSV出力クラス
        Dim strCSVFileName As String        ' CSVファイル名
        Dim SearchFileList As New ArrayList ' 見つかったファイルリスト

        JVData_CSVOutput = CommonConstant.ReturnCd_0

        Try
            ' 変数初期化
            strLastFileTimestamp = ""

            ' レース詳細リストが存在する場合
            If aRaceList.Count <> 0 Then
                ' 作業状況を更新する
                Me.lstUpdateInfo.Items.Add("速報レース詳細を設定しています...")
                Me.lstUpdateInfo.SelectedIndex = Me.lstUpdateInfo.Items.Count - 1

                ' レース詳細CSVファイル名を設定
                strCSVFileName = CommonConstant.RA_FileName
                'strCSVFileName = Mid(strRAFileName, 1, InStr(strRAFileName, ".") - 1) & CommonConstant.CSV

                '' すでにファイルが存在するか検索する
                'Call GetAllFiles(gPrgFilesPath, CommonConstant.RaseID & _
                '                 CommonConstant.ASTERISK & _
                '                 CommonConstant.CSV, SearchFileList)

                '' レース詳細CSVファイルが存在した場合
                'If SearchFileList.Count > 0 Then
                '    For i = 0 To SearchFileList.Count - 1
                '        ' 現在のファイルと異なるか判定
                '        If gPrgFilesPath & "\" & strCSVFileName <> _
                '            SearchFileList.Item(i).ToString Then
                '            ' 異なる場合、現在のファイル名に"old_"を付ける
                '            File.Move(SearchFileList.Item(i).ToString, _
                '                      gPrgFilesPath & CommonConstant.OLDCSV & _
                '                      Path.GetFileName(SearchFileList.Item(i).ToString))
                '        End If
                '    Next
                '    ' ファイルリスト初期化
                '    SearchFileList = New ArrayList
                'End If

                ' レース詳細CSVファイル出力処理を呼び出す
                csvOutput.RaceInfoOutput(strCSVFileName, gPrgFilesPath, aRaceList)
            End If

            ' 馬毎レース情報リストが存在する場合
            If aHorseList.Count <> 0 Then
                ' 作業状況を更新する
                Me.lstUpdateInfo.Items.Add("速報馬毎レース情報を設定しています...")
                Me.lstUpdateInfo.SelectedIndex = Me.lstUpdateInfo.Items.Count - 1

                ' 馬毎レース情報CSVファイル名を設定
                strCSVFileName = CommonConstant.SE_FileName
                'strCSVFileName = Mid(strSEFileName, 1, InStr(strSEFileName, ".") - 1) & CommonConstant.CSV

                '' すでにファイルが存在するか検索する
                'Call GetAllFiles(gPrgFilesPath, CommonConstant.HorseID & _
                '                 CommonConstant.ASTERISK & _
                '                 CommonConstant.CSV, SearchFileList)

                '' 馬毎レース情報CSVファイルが存在した場合
                'If SearchFileList.Count > 0 Then
                '    For i = 0 To SearchFileList.Count - 1
                '        ' 現在のファイルと異なるか判定
                '        If gPrgFilesPath & "\" & strCSVFileName <> _
                '            SearchFileList.Item(i).ToString Then
                '            ' 異なる場合、現在のファイル名に"old_"を付ける
                '            File.Move(SearchFileList.Item(i).ToString, _
                '                      gPrgFilesPath & CommonConstant.OLDCSV & _
                '                      Path.GetFileName(SearchFileList.Item(i).ToString))
                '        End If
                '    Next
                '    ' ファイルリスト初期化
                '    SearchFileList = New ArrayList
                'End If

                ' 馬毎レース情報CSVファイル出力処理を呼び出す
                csvOutput.SeInfoOutput(strCSVFileName, gPrgFilesPath, aHorseList)
            End If

            ' オッズ1情報リストが存在する場合
            If aOdds1List.Count <> 0 Then
                ' 作業状況を更新する
                Me.lstUpdateInfo.Items.Add("速報オッズを設定しています...")
                Me.lstUpdateInfo.SelectedIndex = Me.lstUpdateInfo.Items.Count - 1

                ' オッズ1情報CSVファイル名を設定
                strCSVFileName = CommonConstant.O1_FileName

                '' すでにファイルが存在するか検索する
                'Call GetAllFiles(gPrgFilesPath, CommonConstant.Odds1ID & _
                '                 CommonConstant.ASTERISK & _
                '                 CommonConstant.CSV, SearchFileList)

                '' オッズ1情報CSVファイルが存在した場合
                'If SearchFileList.Count > 0 Then
                '    For i = 0 To SearchFileList.Count - 1
                '        ' 現在のファイルと異なるか判定
                '        If gPrgFilesPath & "\" & strCSVFileName <> _
                '            SearchFileList.Item(i).ToString Then
                '            ' 異なる場合、現在のファイル名に"old_"を付ける
                '            File.Move(SearchFileList.Item(i).ToString, _
                '                      gPrgFilesPath & "\" & CommonConstant.OLDCSV & _
                '                      Path.GetFileName(SearchFileList.Item(i).ToString))
                '        End If
                '    Next
                '    ' ファイルリスト初期化
                '    SearchFileList = New ArrayList
                'End If

                ' オッズ1情報CSVファイル出力処理を呼び出す
                csvOutput.OddsOutput(strCSVFileName, gPrgFilesPath, aOdds1List, CommonConstant.TanshouOddsKbn)
            End If

            ' オッズ2情報リストが存在する場合
            If aOdds2List.Count <> 0 Then
                ' 作業状況を更新する
                Me.lstUpdateInfo.Items.Add("速報オッズを設定しています...")
                Me.lstUpdateInfo.SelectedIndex = Me.lstUpdateInfo.Items.Count - 1

                ' オッズ2情報CSVファイル名を設定
                strCSVFileName = CommonConstant.O2_FileName

                '' すでにファイルが存在するか検索する
                'Call GetAllFiles(gPrgFilesPath, CommonConstant.Odds2ID & _
                '                 CommonConstant.ASTERISK & _
                '                 CommonConstant.CSV, SearchFileList)

                '' オッズ2情報CSVファイルが存在した場合
                'If SearchFileList.Count > 0 Then
                '    For i = 0 To SearchFileList.Count - 1
                '        ' 現在のファイルと異なるか判定
                '        If gPrgFilesPath & "\" & strCSVFileName <> _
                '            SearchFileList.Item(i).ToString Then
                '            ' 異なる場合、現在のファイル名に"old_"を付ける
                '            File.Move(SearchFileList.Item(i).ToString, _
                '                      gPrgFilesPath & "\" & CommonConstant.OLDCSV & _
                '                      Path.GetFileName(SearchFileList.Item(i).ToString))
                '        End If
                '    Next
                '    ' ファイルリスト初期化
                '    SearchFileList = New ArrayList
                'End If

                ' オッズ2情報CSVファイル出力処理を呼び出す
                csvOutput.OddsOutput(strCSVFileName, gPrgFilesPath, aOdds2List, CommonConstant.UmarenOddsKbn)
            End If

            ' オッズ4情報リストが存在する場合
            If aOdds4List.Count <> 0 Then
                ' 作業状況を更新する
                Me.lstUpdateInfo.Items.Add("速報オッズを設定しています...")
                Me.lstUpdateInfo.SelectedIndex = Me.lstUpdateInfo.Items.Count - 1

                ' オッズ4情報CSVファイル名を設定
                strCSVFileName = CommonConstant.O4_FileName

                '' すでにファイルが存在するか検索する
                'Call GetAllFiles(gPrgFilesPath, CommonConstant.Odds4ID & _
                '                 CommonConstant.ASTERISK & _
                '                 CommonConstant.CSV, SearchFileList)

                '' オッズ4情報CSVファイルが存在した場合
                'If SearchFileList.Count > 0 Then
                '    For i = 0 To SearchFileList.Count - 1
                '        ' 現在のファイルと異なるか判定
                '        If gPrgFilesPath & "\" & strCSVFileName <> _
                '            SearchFileList.Item(i).ToString Then
                '            ' 異なる場合、現在のファイル名に"old_"を付ける
                '            File.Move(SearchFileList.Item(i).ToString, _
                '                      gPrgFilesPath & "\" & CommonConstant.OLDCSV & _
                '                      Path.GetFileName(SearchFileList.Item(i).ToString))
                '        End If
                '    Next
                '    ' ファイルリスト初期化
                '    SearchFileList = New ArrayList
                'End If

                ' オッズ4情報CSVファイル出力処理を呼び出す
                csvOutput.OddsOutput(strCSVFileName, gPrgFilesPath, aOdds4List, CommonConstant.UmatanOddsKbn)
            End If

            ' タイマー有効時は、無効化する
            If tmrDownload.Enabled = True Then
                tmrDownload.Enabled = False
            End If
        Catch
            Debug.WriteLine(Err.Description)
            JVData_CSVOutput = Err.Number
        End Try
    End Function

    ' 機能　　 : タイマー設定機能
    ' 機能説明 : 設定された時刻に「データ更新」処理を実行する
    '
    Private Sub tmrTimerSet_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrTimerSet.Tick
        Dim dtNowTime As New DateTime
        Dim dtFromTime As New DateTime
        Dim dtToTime As New DateTime
        Dim strNowTime As String
        Dim strFromTime As String
        Dim strToTime As String
        Dim intUpdateTime As Integer

        ' 現在時刻を取得
        dtNowTime = New DateTime(Now.Ticks)
        strNowTime = Format(dtNowTime, CommonConstant.TimeHHmm)

        ' 更新間隔を取得
        intUpdateTime = GetIni(CommonConstant.INI_Ap_Timer, CommonConstant.INI_Key_UpdateTime, CommonConstant.MIN_15, gIniFile)

        ' 実行する時間帯（FROM）を取得
        strFromTime = GetIni(CommonConstant.INI_Ap_Timer, CommonConstant.INI_Key_FromTime, Format(Now, "yyyy/MM/dd") & CommonConstant.INI_FROMTIME, gIniFile)
        dtFromTime = DateTime.Parse(strFromTime)
        strFromTime = Format(dtFromTime, CommonConstant.TimeHHmm)
        ' 開始時間を設定
        If gFromTime = "" Then
            gFromTime = strFromTime
        End If

        ' 実行する時間帯（TO）を取得
        strToTime = GetIni(CommonConstant.INI_Ap_Timer, CommonConstant.INI_Key_ToTime, Format(Now, "yyyy/MM/dd") & CommonConstant.INI_TOTIME, gIniFile)
        dtToTime = DateTime.Parse(strToTime)
        strToTime = Format(dtToTime, CommonConstant.TimeHHmm)

        ' 現在時刻 > 実行時間（TO）の場合、処理終了
        If strNowTime > strToTime Then
            ' 実行する時間帯（FROM）を開始時間に設定
            gFromTime = strFromTime
            Exit Sub
        End If

        ' 現在時刻と実行時間を比較
        If strNowTime = gFromTime Then
            ' 文字列を日付型のフォーマットに変換する
            ' ここでは"HHmm"（複数指定可能）
            Dim dt1 As DateTime = DateTime.ParseExact(gFromTime, _
                            CommonConstant.TimeHHmm, _
                            System.Globalization.DateTimeFormatInfo.InvariantInfo, _
                            System.Globalization.DateTimeStyles.None)

            ' 時間を加算し、次回更新時間を設定
            gFromTime = Format(dt1.AddMinutes(intUpdateTime), CommonConstant.TimeHHmm)

            ' データ取得ボタン押下処理を呼び出す
            Call btnDataUpdate_Click(sender, e)

            ' 実行時間（FROM < TO）の場合、「タイマー設定」ラベルの文字列を更新
            If gFromTime < strToTime Then
                Me.TimerSetup.Text = "次回更新予定 " & _
                                    Strings.Left(gFromTime, CommonConstant.Const2) & _
                                    ":" & _
                                    Strings.Right(gFromTime, CommonConstant.Const2) & _
                                    "頃"
            End If
        End If
    End Sub

    ' 機能　　 : JVデータダウンロードタイマー処理
    ' 機能説明 : JVデータダウンロード時の進捗状況を更新する
    '
    Private Sub tmrDownload_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrDownload.Tick
        Dim lReturnCode As Long     ' JVLink返値

        ' JVLinkダウンロード進捗率
        lReturnCode = AxJVLink1.JVStatus    ' ダウンロード済のファイル数を返す

        ' エラー判定
        If lReturnCode < 0 Then
            '' エラー
            'MsgBox("JVStatusエラー:" & lReturnCode, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
            ' タイマー停止
            tmrDownload.Enabled = False
            ' JVLink終了処理
            lReturnCode = Me.AxJVLink1.JVClose()
            'If lReturnCode <> 0 Then
            '    MsgBox("JVClseエラー：" & lReturnCode, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
            'End If
        ElseIf lReturnCode < lDownloadCount Then
            ' ダウンロード中
            ' プログレス表示
            prgDownload.Value = lReturnCode
        ElseIf lReturnCode = lDownloadCount Then
            ' ダウンロード完了
            ' タイマー停止
            tmrDownload.Enabled = False
            ' プログレス表示
            prgDownload.Value = lReturnCode
        End If
    End Sub

    ' 機能　　 : タイマー設定ラベル押下時処理
    ' 機能説明 : タイマー設定ダイアログを開き、タイマー起動有無およびタイマー起動時刻等の設定を行う
    '
    Private Sub TimerSetup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerSetup.Click
        Dim fTimer As New dlgTimer
        Dim strFromTime As String = ""

        ' タイマー設定ダイアログ起動
        fTimer.ShowDialog()
        fTimer.Dispose()

        ' タイマー設定結果
        If fTimer.DialogResult = DialogResult.OK Then
            If fTimer.chkTimerOnOff.Checked Then
                If gFromTimeLabel = "" Then
                    strFromTime = FormatDateTime(DateTime.Parse(fTimer.dtpFromTime.Value), DateFormat.ShortTime)
                Else
                    strFromTime = Strings.Left(gFromTimeLabel, CommonConstant.Const2) & ":" & _
                                    Strings.Right(gFromTimeLabel, CommonConstant.Const2)
                End If

                Me.TimerSetup.Text = "次回更新予定 " & strFromTime & "頃"
                ' タイマー開始
                Me.tmrTimerSet.Enabled = True
            Else
                Me.TimerSetup.Text = "タイマーを設定する"
                ' タイマー終了
                Me.tmrTimerSet.Enabled = False
            End If
        ElseIf fTimer.DialogResult = DialogResult.Cancel Then
            ' 何もしない
        End If
    End Sub

    'フォームの再描画を行う
    Private Sub FormReDisplay()
        ' フォームを無効にする
        Me.Enabled = False


        ' フォームを有効にする
        Me.Enabled = True

        ' フォーカスセット処理を呼び出す
        ' SplitContainerの下段パネルにフォーカスを設定
    End Sub

    ' 機能　　 : フォーカスセット処理
    ' 引き数　 : ctrl - 対象コントロール
    ' 返り値　 : なし
    ' 機能説明 : 引数で指定されたコントロールにフォーカスを設定する
    '
    Private Sub Control_SetFocus(ByVal ctrl As Control)
        ' 指定のコントロールにフォーカスを設定
        ctrl.Focus()
    End Sub

    ' 機能     : SplitContainer - マウスホイールイベント
    ' 機能説明 : マウスホイールでPanel内のスクロールバーをスクロールさせる
    '
    Private Sub SplitContainer1_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Dim numberOfTextLinesToMove As Integer = CInt(e.Delta * SystemInformation.MouseWheelScrollLines / 120)
        Dim numberOfPixelsToMove As Integer = numberOfTextLinesToMove * 12
        Dim mousePath As New System.Drawing.Drawing2D.GraphicsPath

        If numberOfPixelsToMove <> 0 Then
            Dim translateMatrix As New System.Drawing.Drawing2D.Matrix()
            translateMatrix.Translate(0, numberOfPixelsToMove)
            mousePath.Transform(translateMatrix)
        End If

    End Sub
    Private Sub mnuAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAbout.Click
        'バージョン情報ダイアログを表示
        dlgVersionInfo.Show()

    End Sub
    '
    '過去1年以上前のレース情報を削除する
    '
    Private Sub mnuKakoDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuKakoDelete.Click
        Dim dlgDelete As New dlgKakoDelete
        Dim dlgDelete2 As New dlgKakoDelete

        dlgDelete.ShowDialog()

        If dlgDelete.DialogResult = Windows.Forms.DialogResult.OK Then
            If InvalidityDataDelete(dlgDelete.intSelectCode) Then
                MsgBox("データの削除中にエラーが発生しました。" & vbCrLf & "管理者にご連絡ください。", MsgBoxStyle.Critical)
            Else
                MsgBox("過去レース情報を削除しました。")
            End If
        End If
    End Sub
    '

    '
    '期限切れデータの削除を行う
    '
    '2010/11/15 d-kobayashi update
    'Private Function InvalidityDataDelete() As Boolean
    Private Function InvalidityDataDelete(ByVal pDeleteIndex As Integer) As Boolean

        Dim ExceptionFlg As Boolean = False

        Try

            ' 日付と時刻を格納するための変数を宣言する 
            Dim nowDate As DateTime = (Date.Now).ToString("yyyy/MM/dd")
            '2010/11/15 d-kobayashi update start
            ' ２年減算する 
            'Dim DateOfRecord As String = (nowDate.AddYears(-2)).ToString("yyyyMMdd")
            Dim DateOfRecord As String = ""
            ''ダイアログで選択した範囲に合わせて削除対象日を決める
            Select Case pDeleteIndex
                Case 0  '2年
                    DateOfRecord = (nowDate.AddYears(-2)).ToString("yyyyMMdd")
                Case 1  '1年
                    DateOfRecord = (nowDate.AddYears(-1)).ToString("yyyyMMdd")
                Case 2  '半年
                    DateOfRecord = (nowDate.AddMonths(-6)).ToString("yyyyMMdd")
                Case 3  '1ヶ月
                    DateOfRecord = (nowDate.AddMonths(-1)).ToString("yyyyMMdd")
                Case 4  '15日
                    DateOfRecord = (nowDate.AddDays(-15)).ToString("yyyyMMdd")
            End Select
            '2010/11/15 d-kobayashi update end

            'レース／騎手データリスト展開
            objCommonUtil.RaceFileOpen()

            'オッズデータリスト展開
            objCommonUtil.OddsFileOpen()

            '日付リストを取得
            objCommonUtil.RaceFileOpen(False)
            objCommonUtil.OddsFileOpen(False)

            'レース情報削除
            'RecordDelete(gPrgFilesPath & "\" & CommonConstant.Old_RA_FileName, _
            '             RaceInfo, _
            '             DateOfRecord)
            RecordDelete(RaceInfo, DateOfRecord)

            '騎手／馬名情報削除
            'RecordDelete(gPrgFilesPath & "\" & CommonConstant.Old_SE_FileName, _
            '            AllHorseInfo, _
            '            DateOfRecord)
            RecordDelete(AllHorseInfo, DateOfRecord)

            '単勝オッズ情報削除
            'RecordDelete(gPrgFilesPath & "\" & CommonConstant.Old_O1_FileName, _
            '            TanFukuAllOddsInfo, _
            '            DateOfRecord)
            RecordDelete(TanFukuAllOddsInfo, DateOfRecord)

            '馬単オッズ情報削除
            'RecordDelete(gPrgFilesPath & "\" & CommonConstant.Old_O4_FileName, _
            '            UmatanAllOddsInfo, _
            '            DateOfRecord)
            RecordDelete(UmatanAllOddsInfo, DateOfRecord)

            '馬連オッズ情報削除
            'RecordDelete(gPrgFilesPath & "\" & CommonConstant.Old_O2_FileName, _
            '             UmarenAllOddsInfo, _
            '             DateOfRecord)
            RecordDelete(UmarenAllOddsInfo, DateOfRecord)

        Catch ex As Exception

            ExceptionFlg = True

        End Try


        Return ExceptionFlg

    End Function



    '基準日切れデータ削除実処理
    '引数
    ' 1：削除ファイル名
    ' 2：データリスト（レース or 騎手／馬名 or 各種オッズ情報）
    ' 3：削除基準日
    'Private Sub RecordDelete(ByVal DeleteDataFileName As String, ByVal DataList As ArrayList, ByVal DateOfRecord As String)
    Private Sub RecordDelete(ByVal Datalist As ArrayList, ByVal DateOfRecord As String)

        Try
            'Dim Writer As New IO.StreamWriter(DeleteDataFileName, _
            '                              False, _
            '                             System.Text.Encoding.GetEncoding("shift_jis"))

            For i = 0 To Datalist.Count - 1

                Application.DoEvents()

                'データリストから１レコード取得
                Dim readDataH As String() = Datalist(i)


                'キー情報部分からレース日を取得
                'Dim readDataRaceDate As String = readDataH(CommonConstant.KeyData).Substring(CommonConstant.Const0, CommonConstant.Const8)
                Dim readDataRaceDate As String = readDataH(CommonConstant.IndexPos0)


                'If (readDataRaceDate >= DateOfRecord) Then
                If (readDataRaceDate < DateOfRecord) Then
                    '基準日以前のファイルは削除する
                    Dim strFname As String = readDataH(CommonConstant.IndexPos1)
                    System.IO.File.Delete(strFname)
                    '基準日と同日か、それより未来の場合、出力ファイルに書き込む
                    'Dim readData As String = Join(readDataH, ",")
                    'Writer.WriteLine(readData)
                End If

            Next

            'Writer.Close()

        Catch ex As Exception


        End Try

    End Sub

    Private Sub btnExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExcel.Click
        openExcelFile()
    End Sub

    Private Sub openExcelFile()
        Dim commonUtil As New clsCommonUtil
        Dim strFilePath As String = ""

        If gAuthority = CommonConstant.AUTHORITY Then
            strFilePath = System.Windows.Forms.Application.StartupPath & "\" & CommonConstant.strExcelFileAdmin
        Else
            strFilePath = System.Windows.Forms.Application.StartupPath & "\" & CommonConstant.strExcelFile
        End If
        If commonUtil.CheckFileUsing(strFilePath) = False Then
            Process.Start("excel", """" & strFilePath & """")
        End If
    End Sub
End Class