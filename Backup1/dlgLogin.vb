Imports System.Windows.Forms
Imports Microsoft.Win32
Imports System.IO

Public Class dlgLogin

    'HttpWebRequestの作成 
    Private webreq As System.Net.HttpWebRequest
    'サーバーからの応答を受信するためのHttpWebResponseを取得 
    Private webres As System.Net.HttpWebResponse
    '応答データを受信するためのStreamを取得 
    Private st As System.IO.Stream
    '文字コード(EUC)を指定する 
    Private enc As System.Text.Encoding = _
        System.Text.Encoding.GetEncoding("euc-jp")
    ' 操作するレジストリ・キーの名前 
    Private rKeyName As String = "Software\anauma"
    Private rKeyNamek As String = "Software\verk"
    ' 設定処理を行う対象となるレジストリの値の名前 
    Private rSetValueName As String = "saveId"
    ' 設定する値のデータ 
    Private setData As String = ""
    '入力情報保存要否判定フラグ
    Private check As String = CommonConstant.rCheckOff
    'メインサーバー接続フラグ
    Private MainServerConnectFlg As Boolean = CommonConstant.serverConnectOff
    'サブサーバー接続フラグ
    Private SubServerConnectFlg As Boolean = CommonConstant.serverConnectOff

    Private AddMode As Boolean = True
    Private NewMode As Boolean = False

    Private outRaceInfoList As New ArrayList
    Private outHorseInfoList As New ArrayList
    Private outTanshouOddsList As New ArrayList
    Private outUmatanList As New ArrayList
    Private outUmarenList As New ArrayList
    Private NinshouDataList As New ArrayList

    Private ActivationDirLocal As String = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal) & "\"

    ' CommonUtilクラスインスタンス
    Private objCommonUtil As New clsCommonUtil

    '画面初期処理
    Private Sub dlgLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' SIDを生成
        gJVLinkSID = CommonConstant.SID

        SplashScreen.SetInitMessage("JV-Linkをチェックしています...")
        System.Threading.Thread.Sleep(1000) ' ダミー
        If Not JVLinkCheck() Then
            MsgBox("JV-Linkがインストールされていません。", MsgBoxStyle.Exclamation + MsgBoxStyle.SystemModal, CommonConstant.AppTitle)
            Exit Sub
        End If

        SplashScreen.SetInitMessage("JV-Linkのバージョンをチェックしています...")
        System.Threading.Thread.Sleep(1000) ' ダミー
        If Not JVLinkVersion2Up() Then
            MsgBox("インストールされているJVLinkのバージョンでは、データの更新を行う事ができません。" _
                & vbCrLf & "データの更新を行う場合は、Ver2.1以降のJVLinkをインストールして下さい。", _
                MsgBoxStyle.Exclamation + MsgBoxStyle.SystemModal, CommonConstant.AppTitle)
            Exit Sub
        Else
            If gReturnCode < -1 Then
                ' アプリケーションを終了する
                Application.Exit()
            End If

            ' バージョン確認
            If gVersion < CommonConstant.CompVersion250 Then
                MsgBox("インストールされているJVLinkのバージョンはVer2.4以前です。（Ver2.5以降推奨）" & vbCrLf & vbCrLf & _
                        "現在インストールされているJVLinkの場合、データ更新を実行する度に新バージョン告知のメッセージが表示されますのでご注意下さい。" & vbCrLf & _
                        "告知メッセージの表示を停止する場合はVer2.5以降のJVLinkをインストールして下さい。" & vbCrLf & _
                        "※影響についてはヘルプ（タイマー機能を使用する）をご参照下さい。", _
                        MsgBoxStyle.Exclamation + MsgBoxStyle.SystemModal, CommonConstant.AppTitle)
            End If
        End If
        SplashScreen.SetInitMessage("データの読み込み中です。この処理は数分かかる場合があります...")

        Try
            'メインサーバーから認証情報を取得する
            'リクエスト情報を設定 
            webreq = _
                CType(System.Net.WebRequest.Create(CommonConstant.MainServerUrl & CommonConstant.NInshouFileName),  _
                    System.Net.HttpWebRequest)
            'サーバーログインＩＤ／パスワード設定
            webreq.Credentials = New System.Net.NetworkCredential(CommonConstant.MainServerHttpUserId, _
                                                                  CommonConstant.MainServerHttpUserPass)


            'レスポンスデータ受信情報を設定 
            webres = _
                CType(webreq.GetResponse(), System.Net.HttpWebResponse)

            st = webres.GetResponseStream()

            Dim sr As New System.IO.StreamReader(st, enc)

            '認証情報のリスト化
            While True
                '認証情報を展開する
                Dim readData As String() = sr.ReadLine().Split(","c)
                NinshouDataList.Add(readData)
                '最終レコードの場合、処理を終了
                If (sr.EndOfStream) Then
                    Exit While
                End If
            End While

            'ストリームを閉じる 
            sr.Close()

            'メインサーバー接続フラグを設定
            MainServerConnectFlg = CommonConstant.serverConnectOn

        Catch ex As System.Net.WebException
            'メインサーバー接続不可フラグを設定
            MainServerConnectFlg = CommonConstant.serverConnectOff
        End Try

        If Not MainServerConnectFlg Then
            Try
                'サブサーバーから認証情報を取得する
                'リクエスト情報を設定 
                webreq = _
                    CType(System.Net.WebRequest.Create(CommonConstant.SubServerUrl & CommonConstant.NInshouFileName),  _
                        System.Net.HttpWebRequest)
                'サーバーログインＩＤ／パスワード設定
                webreq.Credentials = New System.Net.NetworkCredential(CommonConstant.SubServerHttpUserId, _
                                                                      CommonConstant.SubServerHttpUserPass)

                'レスポンスデータ受信情報を設定 
                webres = _
                    CType(webreq.GetResponse(), System.Net.HttpWebResponse)

                st = webres.GetResponseStream()

                Dim sr As New System.IO.StreamReader(st, enc)

                '認証情報のリスト化
                While True
                    '認証情報を展開する
                    Dim readData As String() = sr.ReadLine().Split(","c)
                    NinshouDataList.Add(readData)
                    '最終レコードの場合、処理を終了
                    If (sr.EndOfStream) Then
                        Exit While
                    End If
                End While

                'ストリームを閉じる 
                sr.Close()

                'サブサーバー接続フラグ設定
                SubServerConnectFlg = CommonConstant.serverConnectOn
            Catch ex As System.Net.WebException
                'サブサーバー接続不可フラグ設定
                SubServerConnectFlg = CommonConstant.serverConnectOff
            End Try
        End If

        'メインサーバー＆サブサーバー共に認証情報が取得できなかった場合
        If Not MainServerConnectFlg And Not SubServerConnectFlg Then
            MsgBox("サーバーに接続できませんでした。" & vbCrLf & _
                   "管理者にお問い合わせください。", MsgBoxStyle.Exclamation, CommonConstant.AppTitle)
            btnLogin.Enabled = False
        Else

            Dim saveid As String = String.Empty
            Dim savepass As String = String.Empty
            Dim savecheck As String = String.Empty
            Try
                ' レジストリ・キーを読み込みモードで開く 
                '20141216 変更 start VerAと被っていたレジストリキーを変更
                'Dim rKey As RegistryKey = Registry.CurrentUser.OpenSubKey(rKeyName, False)
                Dim rKey As RegistryKey = Registry.CurrentUser.OpenSubKey(rKeyNamek, False)
                '20141216 変更 end
                '' レジストリ・キーを読み込みモードで開く 
                'Dim rKey As RegistryKey = Registry.LocalMachine.OpenSubKey(rKeyName, False)
                ' レジストリの値を設定（すべてのデータ型を） 
                Dim saveData As String() = rKey.GetValue(rSetValueName).split(","c)
                ' 開いたレジストリを閉じる 
                rKey.Close()

                '
                If saveData.Length = 3 Then
                    saveid = saveData(CommonConstant.rId)
                    savepass = saveData(CommonConstant.rPass)
                    savecheck = saveData(CommonConstant.rAuthority)

                End If

                'レジストリ登録内容を有効とする場合、ID／PASSを表示し、チェックを付加する
                If (savecheck = CommonConstant.rCheckOn) Then
                    txtMailAddress.Text = saveid
                    txtPassword.Text = savepass
                    chkPassSave.Checked = True

                End If


            Catch ex As Exception
                ' レジストリ・キーを新規作成 
                '20141216 VerAと被っていたレジストリキーを変更 start
                'Dim rKey As RegistryKey = Registry.CurrentUser.CreateSubKey(rKeyName)
                Dim rKey As RegistryKey = Registry.CurrentUser.CreateSubKey(rKeyNamek)
                '20141216 VerAと被っていたレジストリキーを変更 end
                '' レジストリ・キーを新規作成 
                'Dim rKey As RegistryKey = Registry.LocalMachine.CreateSubKey(rKeyName)
                ' 開いたレジストリを閉じる 
                rKey.Close()
            End Try
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

        '過去レース情報の更新を行う
        '2010/11/15 d-kobayashi del
        'Call AddOldData()

        '2010/12/16 d-kobayashi add start
        '過去情報の分割
        objCommonUtil.SplitOldDataFile(CommonConstant.Old_RA_FileName, CommonConstant.Old_RA_FileNameDate)
        objCommonUtil.SplitOldDataFile(CommonConstant.Old_SE_FileName, CommonConstant.Old_SE_FileNameDate)
        objCommonUtil.SplitOldDataFile(CommonConstant.Old_O1_FileName, CommonConstant.Old_O1_FileNameDate)
        objCommonUtil.SplitOldDataFile(CommonConstant.Old_O2_FileName, CommonConstant.Old_O2_FileNameDate)
        objCommonUtil.SplitOldDataFile(CommonConstant.Old_O4_FileName, CommonConstant.Old_O4_FileNameDate)

        '2010/12/16 d-kobayashi add end

        Exit Sub
OS_Err:
        MsgBox("このOSはサポート対象外です。", MsgBoxStyle.Exclamation, CommonConstant.AppTitle)
    End Sub

    '製品のアクティベーションを行う。
    Private Function Activation() As Boolean
        Dim blnLocal As Boolean
        Dim blnServer As Boolean
        'アクティベーション確認
        'ファイル名をメールアドレスから作成
        gActivationFileName = Replace(gMailAddress, "@", "_")
        gActivationFileName = Replace(gActivationFileName, ".", "_")
        gActivationFileName = Replace(gActivationFileName, ";", "")

        'ローカルファイルの存在チェック
        blnLocal = ActivationLocalFileFind()
        'サーバファイルの存在チェック
        blnServer = ActivationServerFileFind()

        'debug start
        System.Net.WebRequest.DefaultWebProxy = Nothing
        'debug end

        'ローカルとサーバにファイルがない場合初めての認証となるため、両方にファイルを作成する
        If Not blnLocal And Not blnServer Then
            'ローカルにファイルを作成
            ActivationFileMake()
            'サーバにファイルを作成
            ActivationFileUpload()
            Return True
        ElseIf blnLocal And blnServer Then
            'ローカルとサーバにファイルがある場合、認証成功
            Return True
        ElseIf blnLocal And Not blnServer Then
            'ローカルにしかファイルがない場合、サーバにファイルをアップする
            ActivationFileUpload()
            Return True
        Else
            'それ以外は認証失敗で終了する
            Return False
        End If
    End Function

    'アクティベーション用ファイルがローカルに存在するか確認する。
    Private Function ActivationLocalFileFind() As Boolean
        Return System.IO.File.Exists(ActivationDirLocal & gActivationFileName)
    End Function

    'アクティベーション用ファイルがサーバに存在するか確認する。
    Private Function ActivationServerFileFind() As Boolean
        '20141216 変更 start ファイル読み込み方法を FtpWebRequest → WebRequest に変更
        'Dim req As System.Net.FtpWebRequest = System.Net.WebRequest.Create(CommonConstant.FtpUpload & CommonConstant.ActivationDir & gActivationFileName)

        'req.Credentials = New System.Net.NetworkCredential(CommonConstant.MainServerFtpId, CommonConstant.MainServerFtpPass)
        'req.Method = System.Net.WebRequestMethods.Ftp.ListDirectory
        ''要求の完了後に接続を閉じる
        'req.KeepAlive = False
        ''PASSIVEモードを無効にする
        'req.UsePassive = False

        ''FTPWebResponseを取得
        'Dim ftpres As System.Net.FtpWebResponse = CType(req.GetResponse(), System.Net.FtpWebResponse)

        ''FTPサーバから送信されたデータを取得
        'Dim sr As New System.IO.StreamReader(ftpres.GetResponseStream())
        'Dim res As String = sr.ReadToEnd()
        ''ファイル一覧を表示
        'sr.Close()

        'メインサーバーから認証情報を取得する
        'リクエスト情報を設定 
        Try
            webreq = _
                CType(System.Net.WebRequest.Create(CommonConstant.ActivationDirWeb & gActivationFileName),  _
                    System.Net.HttpWebRequest)
            'サーバーログインＩＤ／パスワード設定
            webreq.Credentials = New System.Net.NetworkCredential(CommonConstant.MainServerHttpUserId, _
                                                                  CommonConstant.MainServerHttpUserPass)


            'レスポンスデータ受信情報を設定 
            webres = _
                CType(webreq.GetResponse(), System.Net.HttpWebResponse)

            st = webres.GetResponseStream()

            Dim sr As New System.IO.StreamReader(st, enc)

            Dim res As String = sr.ReadToEnd

            'ストリームを閉じる 
            sr.Close()

        Catch ex As System.Net.WebException
            If ex.Status = System.Net.WebExceptionStatus.ProtocolError Then
                'HttpWebResponseを取得
                Dim errres As System.Net.HttpWebResponse = _
                    CType(ex.Response, System.Net.HttpWebResponse)
                '応答したURIを表示する
                Console.WriteLine(errres.ResponseUri)
                '応答ステータスコードを表示する
                Console.WriteLine("{0}:{1}", _
                    errres.StatusCode, errres.StatusDescription)

                If errres.StatusCode = Net.HttpStatusCode.NotFound Then
                    Return False
                End If
                Console.WriteLine(ex.Message)
            End If
        End Try
        '20141216 変更 end

        Return True


    End Function

    'ローカル用アクティベーションファイルを作成する
    Private Sub ActivationFileMake()
        Dim writer As New IO.StreamWriter(ActivationDirLocal & gActivationFileName)

        writer.WriteLine(gActivationFileName)

        writer.Close()
    End Sub

    'アクティベーション用ファイルをサーバにアップする
    Private Function ActivationFileUpload() As Boolean
        'アップロードするファイル
        Dim upFile As String = "C:\test.txt"
        'アップロード先のURI
        Dim u As New Uri(CommonConstant.FtpUpload & CommonConstant.ActivationDir & gActivationFileName)

        '送信元ファイル 
        Dim filePath As String = ActivationDirLocal & gActivationFileName
        '送信先ファイル 
        Dim url As String = CommonConstant.FtpUpload & CommonConstant.ActivationDir & gActivationFileName
        'サーバー接続ユーザーID格納エリア
        Dim setId As String = String.Empty
        'サーバー接続パスワード格納エリア
        Dim setPass As String = String.Empty

        '処理結果の戻り値
        Dim retCode As Boolean = True

        'サーバー接続ID／パスワードの設定
        'メインサーバーの場合
        setId = CommonConstant.MainServerFtpId
        setPass = CommonConstant.MainServerFtpPass

        'FtpWebRequestの作成
        Dim ftpReq As System.Net.FtpWebRequest = _
            CType(System.Net.WebRequest.Create(u), System.Net.FtpWebRequest)
        'ログインユーザー名とパスワードを設定
        ftpReq.Credentials = New System.Net.NetworkCredential(setId, setPass)
        'MethodにWebRequestMethods.Ftp.UploadFile("STOR")を設定
        ftpReq.Method = System.Net.WebRequestMethods.Ftp.UploadFile
        '要求の完了後に接続を閉じる
        ftpReq.KeepAlive = False
        'ASCIIモードで転送する
        ftpReq.UseBinary = False
        'PASVモードを無効にする
        ftpReq.UsePassive = False

        'ファイルをアップロードするためのStreamを取得
        Dim reqStrm As System.IO.Stream = ftpReq.GetRequestStream()
        'アップロードするファイルを開く
        Try
            Dim fs As New System.IO.FileStream( _
                filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read)
            'アップロードStreamに書き込む
            Dim buffer(1023) As Byte
            While True
                Dim readSize As Integer = fs.Read(buffer, 0, buffer.Length)
                If readSize = 0 Then
                    Exit While
                End If
                reqStrm.Write(buffer, 0, readSize)
            End While
            fs.Close()
            reqStrm.Close()

            'FtpWebResponseを取得
            Dim ftpRes As System.Net.FtpWebResponse = _
                CType(ftpReq.GetResponse(), System.Net.FtpWebResponse)
            'FTPサーバーから送信されたステータスを表示
            Console.WriteLine("{0}: {1}", ftpRes.StatusCode, ftpRes.StatusDescription)
            '閉じる
            ftpRes.Close()
        Catch ex As Exception
            'アップロードエラーの場合
            retCode = False

        End Try


        Return retCode

    End Function



    'ログインボタン押下時処理
    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click

        Dim NinshouFlg As Boolean = False

        '入力情報を取得
        Dim MailAddress As String = Me.txtMailAddress.Text
        Dim Password As String = Me.txtPassword.Text

        'メールアドレス内容チェック
        If MailAddress = String.Empty Then
            MsgBox("メールアドレスが入力されていません。", MsgBoxStyle.Exclamation, CommonConstant.AppTitle)
            Me.txtMailAddress.Focus()
            Exit Sub
        End If

        'パスワード内容チェック
        If Password = String.Empty Then
            MsgBox("パスワードが入力されていません。", MsgBoxStyle.Exclamation, CommonConstant.AppTitle)
            Me.txtPassword.Focus()
            Exit Sub
        End If


        '受信して表示 
        'Dim html As String = sr.ReadToEnd() 
        'Dim sr As New System.IO.StreamReader(st, enc)
        For i = 0 To NinshouDataList.Count - 1
            '認証情報を展開する
            Dim readData As String() = NinshouDataList(i)
            Dim readMailAddress As String = readData(CommonConstant.IndexPos0)
            Dim readPassowrd As String = readData(CommonConstant.IndexPos1)
            Dim readAuthority As String = readData(CommonConstant.IndexPos2)

            'メールアドレス／パスワードの比較を行う
            If (MailAddress = readMailAddress And Password = readPassowrd) Then
                '認証情報に該当した場合
                'グローバル変数へユーザー情報を設定
                gMailAddress = MailAddress 'メールアドレス
                gPassword = Password 'パスワード
                gAuthority = readAuthority '使用権限

                NinshouFlg = True
                Exit For

            End If

        Next


        '認証ＯＫの場合
        If NinshouFlg Then
            '一般権限のユーザーのみアクティベーションを行う。
            If gAuthority <> CommonConstant.AUTHORITY Then
                If Activation() = False Then
                    MsgBox("このメールアドレスは既にほかの端末で使用されているため、この端末で使用できません。", MsgBoxStyle.Exclamation, CommonConstant.AppTitle)
                    Exit Sub
                End If
            End If

            If chkPassSave.Checked Then
                check = "1"
                ' 設定する値のデータ 
                setData = MailAddress & "," & Password & "," & check ' REG_SZ型 

                ' レジストリの設定と削除 
                Try
                    ' レジストリ・キーを書き込みモードで開く 
                    '20141216 VerAと被っていたレジストリキーを変更 start
                    'Dim rKey As RegistryKey = Registry.CurrentUser.OpenSubKey(rKeyName, True)
                    Dim rKey As RegistryKey = Registry.CurrentUser.OpenSubKey(rKeyNamek, True)
                    '20141216 VerAと被っていたレジストリキーを変更 end
                    '' レジストリ・キーを書き込みモードで開く 
                    'Dim rKey As RegistryKey = Registry.LocalMachine.OpenSubKey(rKeyName, True)
                    ' レジストリの値を設定（すべてのデータ型を） 
                    rKey.SetValue(rSetValueName, setData)
                    ' 開いたレジストリを閉じる 
                    rKey.Close()

                Catch ex As Exception
                    ' レジストリ・キーが存在しない 
                    Console.WriteLine(ex.Message)
                End Try
            Else
                check = "0"
                Try
                    ' レジストリ・キーを書き込みモードで開く 
                    '20141216 VerAと被っていたレジストリキーを変更 start
                    'Dim rKey As RegistryKey = Registry.CurrentUser.OpenSubKey(rKeyName, True)
                    Dim rKey As RegistryKey = Registry.CurrentUser.OpenSubKey(rKeyNamek, True)
                    '20141216 VerAと被っていたレジストリキーを変更 end
                    '' レジストリ・キーを書き込みモードで開く 
                    'Dim rKey As RegistryKey = Registry.LocalMachine.OpenSubKey(rKeyName, True)
                    ' 削除の例（キー無し時の例外発生抑止（FALSE）指定）
                    rKey.DeleteValue(rSetValueName, False)
                    ' 開いたレジストリを閉じる 
                    rKey.Close()

                Catch ex As Exception
                    ' レジストリ・キーが存在しない 
                    Console.WriteLine(ex.Message)
                End Try

            End If


            ' ログイン画面を閉じる
            Me.Visible = False

            ' メイン画面を表示
            frmMain.Show()
        Else
            MsgBox("認証情報に登録がありません。" & vbCrLf & "管理者にお問い合わせください。", MsgBoxStyle.Exclamation, CommonConstant.AppTitle)
            ' アプリケーションを終了する
            'Application.Exit()
        End If

    End Sub


    'キャンセルボタン押下時処理
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        ' アプリケーションを終了する
        Application.Exit()
    End Sub


    '
    '昨日以前のデータを過去レース蓄積データへ追加する
    '
    Public Function AddOldData() As Boolean

        Dim ExceptionFlg As Boolean = False
        Dim transFlg As Boolean = False

        Try
            '更新前の情報を一時退避する
            objCommonUtil.backupToTempFolder()


            transFlg = True

            ' 日付と時刻を格納するための変数を宣言する 
            Dim nowDate As String = (Date.Now).ToString("yyyyMMdd")
            ' ２年減算する 
            'Dim DateOfRecord As String = (nowDate.AddYears(-2)).ToString("yyyyMMdd")

            'レース／騎手データリスト展開
            'objCommonUtil.RaceFileOpen()

            'オッズデータリスト展開
            'objCommonUtil.OddsFileOpen()

            'メモリ領域確保のため、レース情報を開放する
            RaceInfo = Nothing
            AllHorseInfo = Nothing
            TanFukuAllOddsInfo = Nothing
            UmatanAllOddsInfo = Nothing
            UmarenAllOddsInfo = Nothing
            HorseData.Clear()




            '2010/12/15 d-kobayashi update start
            '使用メモリ節約のため、処理順序変更
            '更新前：データを全てメモリに格納後、更新
            '更新後：データの種類ごとに処理を行う。
            '  レース／騎手情報・オッズ１・オッズ２・オッズ４情報ごとに
            '  レース／騎手情報 1.データ読み込み→2.更新→3.メモリ開放 → オッズ１ 1.データ読み込み・・・
            '  を繰り返す。

            '過去レース／騎手データリスト展開
            'Dim OldRaceDataFile As String = gPrgFilesPath & "\" & CommonConstant.Old_RA_FileName
            ''過去レース詳細CSVファイルの存在チェック
            'If Not System.IO.File.Exists(OldRaceDataFile) Then
            '    'ファイルが存在しなければ空ファイルを作成する
            '    Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding(CommonConstant.EncType)
            '    Dim sr As New System.IO.StreamWriter(OldRaceDataFile, True, enc)
            '    sr.Close()
            'End If

            'objCommonUtil.setOldFileToList(OldRaceDataFile, OldRaceInfo)
            '過去データの日付の情報を取得する
            objCommonUtil.setOldFileToList(CommonConstant.Old_RA_FileNameDate, OldRaceInfo)

            objCommonUtil.UpdateRaceInfo(nowDate, CommonConstant.RA_FileName, _
                                         OldRaceInfo, CommonConstant.Old_RA_FileNameDate)
            '過去レース情報へデータ追加
            '2010/12/17 d-kobayashi update start
            'FileWriter(GetAddData(nowDate, RaceInfo, OldRaceInfo, outRaceInfoList), _
            '            gPrgFilesPath & "\" & CommonConstant.Old_RA_FileName, _
            '            AddMode)
            ''レース情報マスタ更新
            'FileWriter(outRaceInfoList, _
            '           gPrgFilesPath & "\" & CommonConstant.RA_FileName, _
            '           NewMode)
            '2010/12/17 d-kobayashi update end
            'メモリ開放
            OldRaceInfo = Nothing
            RaceInfo = Nothing
            outRaceInfoList = Nothing


            '過去騎手データリスト展開
            'Dim OldHorseDataFile As String = gPrgFilesPath & "\" & CommonConstant.Old_SE_FileName
            ''過去騎手情報CSVファイルの存在チェック
            'If Not System.IO.File.Exists(OldHorseDataFile) Then
            '    'ファイルが存在しなければ空ファイルを作成する
            '    Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding(CommonConstant.EncType)
            '    Dim sr As New System.IO.StreamWriter(OldHorseDataFile, True, enc)
            '    sr.Close()
            'End If
            'objCommonUtil.setFileToList(OldHorseDataFile, OldHorseInfo)
            objCommonUtil.setOldFileToList(CommonConstant.Old_SE_FileNameDate, OldHorseInfo)
            objCommonUtil.UpdateRaceInfo(nowDate, CommonConstant.SE_FileName, _
                                         OldHorseInfo, CommonConstant.Old_SE_FileNameDate)
            'UpdateRaceInfo(nowDate, AllHorseInfo, OldHorseInfo, outHorseInfoList, CommonConstant.Old_SE_FileNameDate)
            '過去騎手／馬名情報追加
            'FileWriter(GetAddData(nowDate, AllHorseInfo, OldHorseInfo, outHorseInfoList), _
            '           gPrgFilesPath & "\" & CommonConstant.Old_SE_FileName, _
            '           AddMode)
            ''騎手／馬名マスタ更新
            'FileWriter(outHorseInfoList, _
            '           gPrgFilesPath & "\" & CommonConstant.SE_FileName, _
            '           NewMode)
            'メモリ開放
            OldHorseInfo = Nothing
            AllHorseInfo = Nothing
            outHorseInfoList = Nothing


            ''過去オッズデータリスト展開
            'Dim OldOdds1DataFile As String = gPrgFilesPath & "\" & CommonConstant.Old_O1_FileName
            ''過去オッズ情報1情報CSVファイルの存在チェック
            'If Not System.IO.File.Exists(OldOdds1DataFile) Then
            '    'ファイルが存在しなければ空ファイルを作成する
            '    Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding(CommonConstant.EncType)
            '    Dim sr As New System.IO.StreamWriter(OldOdds1DataFile, True, enc)
            '    sr.Close()
            'End If
            'objCommonUtil.setFileToList(OldOdds1DataFile, OldOdds1Info)
            objCommonUtil.setOldFileToList(CommonConstant.Old_O1_FileNameDate, OldOdds1Info)
            objCommonUtil.UpdateRaceInfo(nowDate, CommonConstant.O1_FileName, _
                                         OldOdds1Info, CommonConstant.Old_O1_FileNameDate)

            'UpdateRaceInfo(nowDate, TanFukuAllOddsInfo, OldOdds1Info, outTanshouOddsList, CommonConstant.Old_O1_FileNameDate)
            '単勝／複勝オッズ情報追加
            'FileWriter(GetAddData(nowDate, TanFukuAllOddsInfo, OldOdds1Info, outTanshouOddsList), _
            '           gPrgFilesPath & "\" & CommonConstant.Old_O1_FileName, _
            '           AddMode)
            ''単勝／複勝オッズマスタ更新
            'FileWriter(outTanshouOddsList, _
            '           gPrgFilesPath & "\" & CommonConstant.O1_FileName, _
            'NewMode)
            'メモリ開放
            OldOdds1Info = Nothing
            TanFukuAllOddsInfo = Nothing
            outTanshouOddsList = Nothing


            '過去オッズ情報2データリスト展開
            'Dim OldOdds2DataFile As String = gPrgFilesPath & "\" & CommonConstant.Old_O2_FileName
            ''過去オッズ情報2CSVファイルの存在チェック
            'If Not System.IO.File.Exists(OldOdds2DataFile) Then
            '    'ファイルが存在しなければ空ファイルを作成する
            '    Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding(CommonConstant.EncType)
            '    Dim sr As New System.IO.StreamWriter(OldOdds2DataFile, True, enc)
            '    sr.Close()
            'End If
            'objCommonUtil.setFileToList(OldOdds2DataFile, OldOdds2Info)
            objCommonUtil.setOldFileToList(CommonConstant.Old_O2_FileNameDate, OldOdds2Info)
            objCommonUtil.UpdateRaceInfo(nowDate, CommonConstant.O2_FileName, _
                                         OldOdds2Info, CommonConstant.Old_O2_FileNameDate)
            'UpdateRaceInfo(nowDate, UmatanAllOddsInfo, OldOdds2Info, outUmatanList, CommonConstant.Old_O2_FileNameDate)
            '馬単オッズ情報追加
            'FileWriter(GetAddData(nowDate, UmatanAllOddsInfo, OldOdds2Info, outUmatanList), _
            '           gPrgFilesPath & "\" & CommonConstant.Old_O4_FileName, _
            '           AddMode)
            ''馬単オッズマスタ更新
            'FileWriter(outUmatanList, _
            '           gPrgFilesPath & "\" & CommonConstant.O4_FileName, _
            '           NewMode)
            'メモリ開放

            OldOdds2Info = Nothing
            UmatanAllOddsInfo = Nothing
            outUmatanList = Nothing


            '過去オッズ情報4データリスト展開
            'Dim OldOdds4DataFile As String = gPrgFilesPath & "\" & CommonConstant.Old_O4_FileName
            ''過去オッズ情報4CSVファイルの存在チェック
            'If Not System.IO.File.Exists(OldOdds4DataFile) Then
            '    'ファイルが存在しなければ空ファイルを作成する
            '    Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding(CommonConstant.EncType)
            '    Dim sr As New System.IO.StreamWriter(OldOdds4DataFile, True, enc)
            '    sr.Close()
            'End If
            'objCommonUtil.setFileToList(OldOdds4DataFile, OldOdds4Info)
            objCommonUtil.setOldFileToList(CommonConstant.Old_O4_FileNameDate, OldOdds4Info)
            objCommonUtil.UpdateRaceInfo(nowDate, CommonConstant.O4_FileName, _
                                         OldOdds4Info, CommonConstant.Old_O4_FileNameDate)
            'UpdateRaceInfo(nowDate, UmarenAllOddsInfo, OldOdds4Info, outUmarenList, CommonConstant.Old_O4_FileNameDate)
            '馬連オッズ情報追加
            'FileWriter(GetAddData(nowDate, UmarenAllOddsInfo, OldOdds4Info, outUmarenList), _
            '           gPrgFilesPath & "\" & CommonConstant.Old_O2_FileName, _
            '           AddMode)
            ''馬連オッズマスタ更新
            'FileWriter(outUmarenList, _
            '           gPrgFilesPath & "\" & CommonConstant.O2_FileName, _
            '           NewMode)
            'メモリ開放
            OldOdds4Info = Nothing
            UmarenAllOddsInfo = Nothing
            outUmarenList = Nothing
            '2010/12/15 d-kobayashi update end

            '一時ファイルを削除する
            objCommonUtil.deleteToTempFolder()


        Catch ex As Exception

            'ファイルをロールバックする
            If transFlg Then objCommonUtil.rollbackToTempFolder()
            ExceptionFlg = True

        End Try



    End Function
    '
    '追加対象データリストを返却する
    '
    Private Function GetAddData(ByVal nowDate As String, ByVal DataList As ArrayList, _
                                ByVal OldDataList As ArrayList, ByRef outList As ArrayList) As ArrayList
        Dim retList As New ArrayList
        Dim blnEqualsFlg As Boolean

        For i = 0 To DataList.Count - 1
            Dim readData As String() = DataList(i)
            'データレコードからレース日を取得する
            Dim readYMD As String = readData(CommonConstant.IndexPos0).Substring(CommonConstant.Const0, CommonConstant.Const8)

            blnEqualsFlg = False

            '現在年月日とレース日の比較
            If (nowDate > readYMD) Then
                '過去レース情報に追加対象レコードが存在するか調べる
                For j = 0 To OldDataList.Count - 1
                    Dim readOldData As String() = OldDataList(j)
                    If Join(readData, ",") = Join(readOldData, ",") Then
                        blnEqualsFlg = True
                        Exit For
                    End If
                Next

                'レコードが存在しなければ、過去レース情報に追加する
                If Not blnEqualsFlg Then
                    '先月以前のデータの場合
                    retList.Add(Join(readData, ","))
                End If
            Else
                outList.Add(Join(readData, ","))
            End If
            Application.DoEvents()
        Next

        Return retList

    End Function

    'リストに格納された情報を出力ファイルに追加する
    '
    Private Sub FileWriter(ByVal outDataList As ArrayList, ByVal OutPutFile As String, ByVal Mode As Boolean)

        Dim sw As New System.IO.StreamWriter(OutPutFile, _
                                             Mode, _
                                             System.Text.Encoding.GetEncoding("shift_jis"))
        'リストの内容を1行ずつ書き込む 
        For i = 0 To outDataList.Count - 1
            Dim readData = outDataList(i)
            sw.WriteLine(readData)
            Application.DoEvents()
        Next
        '閉じる 
        sw.Close()

    End Sub

    ' パスワードを忘れた場合
    Private Sub lblNotPass_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblNotPass.Click
        'パスワードを忘れた場合のURLを標準のブラウザで開いて表示する
        System.Diagnostics.Process.Start(CommonConstant.ContactURL)
    End Sub
End Class
