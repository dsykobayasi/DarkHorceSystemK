Module basCommonUtil
    ' <共通変数>
    Public objCodeConv As Object            ' コード変換インスタンス
    Public gJVLinkSID As String             ' SID
    Public gTimerFlg As Boolean             ' タイマーフラグ
    Public gIniFile As String               ' INIファイルパス
    Public gFromTime As String              ' タイマー開始時間
    Public gFromTimeLabel As String         ' タイマーを実行する時間ラベル
    Public gPrgFilesPath As String          ' CSVファイルパス
    Public gCacheFilePath As String         ' Cacheファイルパス
    Public gReturnCode As Integer           ' JVLinkリターンコード
    Public gFileOpenFlg As Boolean          ' ファイルオープンフラグ
    Public gVersion As Long                 ' JVLinkのバージョン

    Declare Function WritePrivateProfileString _
        Lib "KERNEL32.DLL" Alias "WritePrivateProfileStringA" (ByVal lpAppName As String, _
                                                               ByVal lpKeyName As String, _
                                                               ByVal lpString As String, _
                                                               ByVal lpFileName As String) As Integer
    Declare Function GetPrivateProfileString _
        Lib "KERNEL32.DLL" Alias "GetPrivateProfileStringA" (ByVal lpAppName As String, _
                                                             ByVal lpKeyName As String, _
                                                             ByVal lpDefault As String, _
                                                             ByVal lpReturnedString As String, _
                                                             ByVal nSize As Integer, _
                                                             ByVal lpFileName As String) As Integer

    Public RaceInfo As ArrayList = New ArrayList 'レース日該当レース情報格納リスト
    Public AllHorseInfo As ArrayList = New ArrayList '全騎手／馬名リスト
    Public HorseData As ArrayList = New ArrayList '騎手／馬名情報

    Public EightHousokuTable As Hashtable = New Hashtable  '8時時点の法則成立カウント保持
    Public NineHousokuTable As Hashtable = New Hashtable    '9時時点の法則成立カウント保持
    Public TenHousokuTable As Hashtable = New Hashtable     '10時時点の法則成立カウント保持

    Public gSelAllUpdateList As ArrayList '全発表時間を保持するリスト
    Public gSelUpdateTime As ArrayList '選択された発表時間を保持するリスト

    Public TanFukuAllOddsInfo As ArrayList = New ArrayList '全単勝／複勝オッズ情報格納リスト
    Public UmatanAllOddsInfo As ArrayList = New ArrayList '全馬単オッズ情報格納リスト
    Public UmarenAllOddsInfo As ArrayList = New ArrayList '全馬連オッズ情報格納リスト

    Public gMailAddress As String = String.Empty '認証情報：メールアドレス
    Public gPassword As String = String.Empty '認証情報：パスワード
    Public gAuthority As String = String.Empty '認証情報：使用権限

    Public gRaceFileName As String = String.Empty 'レース情報ファイル
    Public gJockeyHorseNameFile As String = String.Empty '騎手／馬名情報ファイル
    Public gTanshouOddsFile As String = String.Empty '単勝オッズ情報ファイル
    Public gUmatanOddsFile As String = String.Empty '馬単オッズ情報ファイル
    Public gUmarenOddsFile As String = String.Empty '馬連オッズ情報ファイル

    Public OldRaceInfo As ArrayList = New ArrayList
    Public OldHorseInfo As ArrayList = New ArrayList
    Public OldOdds1Info As ArrayList = New ArrayList
    Public OldOdds2Info As ArrayList = New ArrayList
    Public OldOdds4Info As ArrayList = New ArrayList

    '20140810 アクティベーション用変数
    Public gActivationFileName As String = String.Empty 'ファイル名



    ' 機能     : INIファイル情報取得
    ' 引き数   : ApName - セクション名
    '            KeyName - 項目名
    '            Default - 項目が存在しない場合の初期値
    '            FileName - 参照ファイル名
    ' 返り値   : String - 取得キー値
    ' 機能説明 : INIファイルから参照したいキーの値を取得する
    '
    Public Function GetIni(ByVal ApName As String, _
                           ByVal KeyName As String, _
                           ByVal Defaults As String, _
                           ByVal Filename As String) As String

        Dim strResult As String = Space(255)
        Call GetPrivateProfileString(ApName, KeyName, Defaults, _
                                     strResult, Len(strResult), _
                                     Filename)
        GetIni = Microsoft.VisualBasic.Left(strResult, _
                                            InStr(strResult, Chr(0)) - 1)
    End Function

    ' 機能     : INIファイル初期処理
    ' 引き数   : なし
    ' 返り値   : なし
    ' 機能説明 : INIファイルの初期処理
    '            INIファイルのディレクトリ存在チェック(エラーなら作成)
    '
    Public Sub InitIniFileProc()
        Dim i As Integer = 0

        'INIファイルのディレクトリ存在チェック
        If Not System.IO.Directory.Exists(CommonConstant.INI_DIR) Then
            IO.Directory.CreateDirectory(CommonConstant.INI_DIR)

        End If

        Dim di As New IO.DirectoryInfo(CommonConstant.INI_DIR)

        '隠し属性を追加する
        If Not (di.Attributes And IO.FileAttributes.Hidden) Then
            di.Attributes = di.Attributes Or IO.FileAttributes.Hidden
        End If

        gIniFile = CommonConstant.INI_DIR & CommonConstant.INIFile

    End Sub

    ' 機能     : INIファイル書き込み処理
    ' 引き数   : ApName - セクション名
    '            KeyName - 項目名
    '            Param - 更新する値
    '            FileName - 書出ファイル名
    ' 返り値   : なし
    ' 機能説明 : INIファイルに新たなキーの値を書込む
    '            ※既存のキーがあれば更新・なければ新規作成する
    '
    Public Sub PutIni(ByVal ApName As String, _
                      ByVal KeyName As String, _
                      ByVal Param As String, _
                      ByVal Filename As String)

        Call WritePrivateProfileString(ApName, KeyName, _
                                       Param, Filename)
    End Sub

    '
    '   機能: JVLinkがインストールされているか調べる
    '
    '   備考: なし
    '
    Public Function JVLinkCheck() As Boolean
        On Error GoTo ErrorHandler
        Dim JVlink As frmWrappedJVLink
        JVlink = New frmWrappedJVLink
        JVlink.Show()
        JVlink.Close()
        JVLinkCheck = True
        Exit Function
ErrorHandler:
        JVLinkCheck = False
    End Function


    '
    '   機能: JVLinkのバージョンを調べる
    '
    '   備考: なし
    '
    Public Function JVLinkVersion2Up() As Boolean
        On Error GoTo ErrorHandler

        Dim JVlink As frmWrappedJVLink
        Dim lngVersion As Long
        Dim strDummy As String
        Dim lngDummy1 As Long
        Dim lngDummy2 As Long

        JVlink = New frmWrappedJVLink
        JVlink.Show()

        gVersion = CLng(JVlink.m_JVLinkVersion)

        ' JVInit
        gReturnCode = JVlink.axJVLink.JVInit(gJVLinkSID)
        If gReturnCode = CommonConstant.ReturnCd_0 Then
            ' JVOpen
            gReturnCode = JVlink.axJVLink.JVOpen(CommonConstant.JVChkData, _
                                                   CommonConstant.JVChkTime, _
                                                   2, lngDummy1, lngDummy2, strDummy)

            If gReturnCode < CommonConstant.ReturnCd_1 Then
                If gReturnCode <> CommonConstant.ReturnCd_504 Then
                    MsgBox(ErrMsgJVOpen(gReturnCode), MsgBoxStyle.Exclamation + MsgBoxStyle.SystemModal, _
                           CommonConstant.AppTitle)
                    'MsgBox("JVLink - JVOpenエラー", MsgBoxStyle.Exclamation, CommonConstant.AppTitle)
                End If
            End If

            ' JVClose
            JVlink.axJVLink.JVClose()
        End If

        JVlink.Close()
        If gVersion < CommonConstant.CompVersion Then
            JVLinkVersion2Up = False
        Else
            JVLinkVersion2Up = True
        End If

        Exit Function
ErrorHandler:
        JVLinkVersion2Up = False
    End Function


    ' 機能     : ファイル検索処理
    ' 引き数   : strFolder - 検索対象のフォルダ名
    '            strSearchPattern - ファイル名検索文字列
    '            FileList - 見つかったファイル名リスト
    ' 返り値   : なし
    ' 機能説明 : 指定されたフォルダ以下にあるすべてのファイルを取得する
    '
    Public Sub GetAllFiles(ByVal strFolder As String, _
            ByVal strSearchPattern As String, ByRef FileList As ArrayList)
        ' strFolderにあるファイルを取得する
        Dim fs As String() = _
            System.IO.Directory.GetFiles(strFolder, strSearchPattern)
        ' ArrayListに追加する
        FileList.AddRange(fs)

        '' strFolderのサブフォルダを取得する
        'Dim ds As String() = System.IO.Directory.GetDirectories(strFolder)
        '' サブフォルダにあるファイルも調べる
        'Dim d As String
        'For Each d In ds
        '    GetAllFiles(d, strSearchPattern, FileList)
        'Next d
    End Sub

    '法則成立カウント保持テーブルへのセッターファンクション
    Public Sub HousokuTableSetter(ByVal UmaNo As String, ByVal OddsKbn As String, ByVal TimeKbn As String, ByVal Data As String())
        '設定キー生成（馬番号＋オッズ区分）
        Dim setKey As String = UmaNo & OddsKbn

        '時系列区分で更新するハッシュテーブルを切り替える
        If (TimeKbn = CommonConstant.EightTime) Then
            If (EightHousokuTable.ContainsKey(setKey)) Then
                'ハッシュテーブルにキーが存在した場合、対象のデータを削除
                EightHousokuTable.Remove(setKey)
            End If

            EightHousokuTable.Add(setKey, Data)

        ElseIf (TimeKbn = CommonConstant.NineTime) Then
            If (NineHousokuTable.ContainsKey(setKey)) Then
                'ハッシュテーブルにキーが存在した場合、対象のデータを削除
                NineHousokuTable.Remove(setKey)
            End If

            NineHousokuTable.Add(setKey, Data)

        Else
            If (TenHousokuTable.ContainsKey(setKey)) Then
                'ハッシュテーブルにキーが存在した場合、対象のデータを削除
                TenHousokuTable.Remove(setKey)
            End If

            TenHousokuTable.Add(setKey, Data)

        End If

    End Sub
    '法則成立カウント保持テーブルへのゲッターファンクション
    'キー存在しない場合、空の配列を返却する
    Public Function HousokuTableGetter(ByVal UmaNo As String, ByVal OddsKbn As String, ByVal TimeKbn As String)
        '取得キー生成（馬番号＋オッズ区分）
        Dim getKey As String = UmaNo & OddsKbn
        Dim retData As String() = {}
        '時系列区分で更新するハッシュテーブルを切り替える
        If (TimeKbn = CommonConstant.EightTime) Then
            If (EightHousokuTable.ContainsKey(getKey)) Then
                'ハッシュテーブルにキーが存在した場合、対象のデータを取得
                retData = CType(EightHousokuTable(getKey), String())

            End If


        ElseIf (TimeKbn = CommonConstant.NineTime) Then
            If (NineHousokuTable.ContainsKey(getKey)) Then
                'ハッシュテーブルにキーが存在した場合、対象のデータを削除
                retData = CType(NineHousokuTable(getKey), String())
            End If



        Else
            If (TenHousokuTable.ContainsKey(getKey)) Then
                'ハッシュテーブルにキーが存在した場合、対象のデータを削除
                retData = CType(TenHousokuTable(getKey), String())
            End If



        End If

        Return retData
    End Function

    '
    '   機能: JVOpenのエラーメッセージを変換する
    '
    '   備考: なし
    '
    Public Function ErrMsgJVOpen(ByVal lngRet As Long) As String
        Select Case lngRet
            'Case CommonConstant.ReturnCd_0
            '    ErrMsgJVOpen = "正常" & vbCrLf & ""
            'Case CommonConstant.ReturnCd_1
            '    ErrMsgJVOpen = "該当データ無し" & vbCrLf & "指定されたパラメータに合致する新しいデータがサーバーに存在しない｡又は､最新バージョンが公開され､ユーザーが最新バージョンのダウンロードを選択しました｡JVCloseを呼び出して取り込み処理を終了してください｡"
            'Case CommonConstant.ReturnCd_2
            '    ErrMsgJVOpen = "セットアップダイアログでキャンセルが押された" & vbCrLf & "セットアップ用データの取り込み時にユーザーがダイアログでキャンセルを押しました｡JVCloseを呼び出して取り込み処理を終了してください｡ "
            'Case CommonConstant.ReturnCd_111
            '    ErrMsgJVOpen = "dataspecパラメータが不正" & vbCrLf & "パラメータの渡し方かパラメータの内容に問題があると思われます｡サンプルプログラム等を参照し､正しくパラメータがJV -Linkに渡っているか確認してください｡ "
            'Case CommonConstant.ReturnCd_112
            '    ErrMsgJVOpen = "fromdateパラメータが不正" & vbCrLf & "パラメータの渡し方かパラメータの内容に問題があると思われます｡サンプルプログラム等を参照し､正しくパラメータがJV -Linkに渡っているか確認してください｡ "
            'Case CommonConstant.ReturnCd_114
            '    ErrMsgJVOpen = "keyパラメータが不正" & vbCrLf & "パラメータの渡し方かパラメータの内容に問題があると思われます｡サンプルプログラム等を参照し､正しくパラメータがJV -Linkに渡っているか確認してください｡ "
            'Case CommonConstant.ReturnCd_115
            '    ErrMsgJVOpen = "optionパラメータが不正" & vbCrLf & "パラメータの渡し方かパラメータの内容に問題があると思われます｡サンプルプログラム等を参照し､正しくパラメータがJV -Linkに渡っているか確認してください｡ "
            'Case CommonConstant.ReturnCd_116
            '    ErrMsgJVOpen = "dataspecとoptionの組み合わせが不正" & vbCrLf & "パラメータの渡し方かパラメータの内容に問題があると思われます｡サンプルプログラム等を参照し､正しくパラメータがJV -Linkに渡っているか確認してください｡ "
            'Case CommonConstant.ReturnCd_201
            '    ErrMsgJVOpen = "ＪＶＩｎｉｔが行なわれていない" & vbCrLf & "JVOpen/JVRTOpenに先立ってJVInitが呼ばれていないと思われます｡必ずJVInitを先に呼び出してください｡ "
            'Case CommonConstant.ReturnCd_202
            '    ErrMsgJVOpen = "前回のJVOpen/JVRTOpenに対してJVCloseが呼ばれていない（オープン中）" & vbCrLf & "前回呼び出したJVOpen/JVRTOpenがJVCloseによってクローズされていないと思われます｡JVOpen/JVRTOpenを呼び出した後は次に呼び出すまでの間にJVCloseを必ず呼び出してください｡ "
            'Case CommonConstant.ReturnCd_211
            '    ErrMsgJVOpen = "レジストリ内容が不正（レジストリ内容が不正に変更された）" & vbCrLf & "JV-Linkはレジストリに値をセットする際に値のチェックを行います（例えばサービスキーの桁数など）が、レジストリから値を読み出して使用する際に問題が発生するとこのエラーが発生します｡レジストリが直接書き換えられたなどの状況が考えられない場合にはJRA-VANへご連絡ください。"
            'Case CommonConstant.ReturnCd_301
            '    ErrMsgJVOpen = "認証エラー" & vbCrLf & "サービスキーが正しくない。あるいは複数のマシンで同一サービスキーを使用した場合に発生します。複数のマシンで同じサービスキーをしようした場合には、このエラーが発生したマシンのJV-Linkをアンインストールし、再インストール後、利用キーの再発行が必要となります。"
            Case CommonConstant.ReturnCd_302
                ErrMsgJVOpen = "サービスキーの有効期限切れ" & vbCrLf & "Data Lab.サービスの有効期限が切れています。サービス権の自動延長が停止していると思われます｡解消するにはサービス権の再購入が必要です｡ "
            Case CommonConstant.ReturnCd_303
                ErrMsgJVOpen = "サービスキーが設定されていない（サービスキーが空値）" & vbCrLf & "サービスキーを設定していないと思われます。JVLinkインストール直後はサービスキーが空なので必ず設定する必要があります｡ "
            Case CommonConstant.ReturnCd_401
                ErrMsgJVOpen = "JV-Link内部エラー" & vbCrLf & "JV-Link内部でエラーが発生したと思われます。JRAVANへご連絡ください｡ "
            Case CommonConstant.ReturnCd_411
                ErrMsgJVOpen = "サーバーエラー（ HTTP ステータス404NotFount）" & vbCrLf & "レジストリが直接変更されたか、Data Lab.用サーバーに問題が発生したと思われます｡JRA -VANのメンテナンス中でない場合で､このエラーが続く場合はJRA-VANへご連絡ください。"
            Case CommonConstant.ReturnCd_412
                ErrMsgJVOpen = "サーバーエラー（ HTTP ステータス403Forbidden）" & vbCrLf & "Data Lab.用サーバーに問題が発生したと思われます｡このエラーが続く場合はJRA -VANへご連絡ください｡ "
            Case CommonConstant.ReturnCd_413
                ErrMsgJVOpen = "サーバーエラー（HTTPステータス200,403,404以外）" & vbCrLf & "Data Lab.用サーバーに問題が発生したと思われます｡このエラーが続く場合はJRA -VANへご連絡ください｡ "
            Case CommonConstant.ReturnCd_421
                ErrMsgJVOpen = "サーバーエラー（サーバーの応答が不正）" & vbCrLf & "Data Lab.用サーバーに問題が発生したと思われます｡このエラーが続く場合はJRA -VANへご連絡ください｡ "
            Case CommonConstant.ReturnCd_431
                ErrMsgJVOpen = "サーバーエラー（サーバーアプリケーション内部エラー）" & vbCrLf & "Data Lab.用サーバーに問題が発生したと思われます｡このエラーが続く場合はJRA -VANへご連絡ください｡ "
            Case CommonConstant.ReturnCd_501
                ErrMsgJVOpen = "セットアップ処理においてＣＤ－ＲＯＭが無効" & vbCrLf & "JRA-VANが提供した正しいCD-ROMをセットしていないと思われます｡正しいCD -ROMをセットしてください｡ "
            Case CommonConstant.ReturnCd_504
                ErrMsgJVOpen = "サーバーメンテナンス中" & vbCrLf & "サーバーがメンテナンス中です。"
            Case Else
                ErrMsgJVOpen = "想定外のエラーが発生しました。" & vbCrLf & ""
        End Select
    End Function
End Module
