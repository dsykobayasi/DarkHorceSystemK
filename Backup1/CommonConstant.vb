Public Class CommonConstant
    '共通定数定義クラス

    'システム設定関連定数
    'アプリタイトル
    Public Const Admin_AppTitle As String = "互當穴ノ守 十周年記念ソフト(R)（管理者）"
    'アプリタイトル
    Public Const AppTitle As String = "互當穴ノ守 十周年記念ソフト(R)"
    'JRA-VANデータ取得モード（今週モード）
    Public Const DataSpec As String = "RACE"
    'Public Const DataSpec As String = "TOKURACE"
    'Public Const DataSpec As String = "TOKURACETCOVRCOV"
    Public Const WeekMode As String = "2"
    Public Const RtDataSpec_RA As String = "0B15"
    Public Const RtDataSpec_O1 As String = "0B31"
    Public Const RtDataSpec_O2 As String = "0B32"
    Public Const RtDataSpec_O4 As String = "0B34"
    'JRA-VANデータ取得開始日時
    Public Const DataFromTime As String = "20100701000000"
    'Public Const DataFromTime As String = "00000000000000"
    'JRA-VANチェック
    Public Const JVChkTime As String = "99999999999999"
    Public Const JVChkData As String = "RCOV"
    'レコード種別ID
    Public Const RaseID As String = "RA"
    Public Const HorseID As String = "SE"
    Public Const Odds1ID As String = "O1"
    Public Const Odds2ID As String = "O2"
    Public Const Odds4ID As String = "O4"
    'CSVファイルタイムスタンプ
    Public Const TimeStamp As String = "yyyyMMddHHmmss"
    'JRA-VANデータ取得SID
    Public Const SID As String = "UNKNOWN"
    'エンコード
    Public Const EncType As String = "Shift_JIS"
    'JVLink結果コード
    Public Const ReturnCd_0 As Integer = 0
    Public Const ReturnCd_1 As Integer = -1
    Public Const ReturnCd_2 As Integer = -2
    Public Const ReturnCd_3 As Integer = -3
    Public Const ReturnCd_111 As Integer = -111
    Public Const ReturnCd_112 As Integer = -112
    Public Const ReturnCd_114 As Integer = -114
    Public Const ReturnCd_115 As Integer = -115
    Public Const ReturnCd_116 As Integer = -116
    Public Const ReturnCd_201 As Integer = -201
    Public Const ReturnCd_202 As Integer = -202
    Public Const ReturnCd_203 As Integer = -203
    Public Const ReturnCd_211 As Integer = -211
    Public Const ReturnCd_301 As Integer = -301
    Public Const ReturnCd_302 As Integer = -302
    Public Const ReturnCd_303 As Integer = -303
    Public Const ReturnCd_401 As Integer = -401
    Public Const ReturnCd_411 As Integer = -411
    Public Const ReturnCd_412 As Integer = -412
    Public Const ReturnCd_413 As Integer = -413
    Public Const ReturnCd_421 As Integer = -421
    Public Const ReturnCd_431 As Integer = -431
    Public Const ReturnCd_501 As Integer = -501
    Public Const ReturnCd_502 As Integer = -502
    Public Const ReturnCd_503 As Integer = -503
    Public Const ReturnCd_504 As Integer = -504
    'JVLink比較バージョンコード
    Public Const CompVersion250 As Integer = 250
    Public Const CompVersion As Integer = 210
    'タイマーフォーマット
    Public Const TimeHHmm As String = "HHmm"
    '馬番（未決定）
    Public Const No_HorseNo As String = "00"

    '各種パス
    'コードファイル
    Public Const CodeFile As String = "\CodeTable.csv"
    'INIファイル
    Public Const INIFile As String = "\verk.ini"
    '2011/04/07 d-kobayashi add start
    'INIファイル保存パス
    Public Const INI_DIR As String = "C:\AnasoftK"
    Public Const INI_FROMTIME As String = " 7:00:00"
    Public Const INI_TOTIME As String = " 13:00:00"
    Public Const INI_TIMERFLG As Boolean = False

    '2011/04/07 d-kobayashi add end

    'CSV保存パス（OS：XP以前）
    Public Const JravanPath As String = "\JRA-VAN\Data Lab\data"
    'CSV保存パス（OS：Vista、7）
    Public Const JravanPath64 As String = "C:\ProgramData\JRA-VAN\Data Lab\data"
    'キャッシュ保存パス（OS：XP以前）
    Public Const JravanCachePath As String = "\JRA-VAN\Data Lab\Cache"
    'キャッシュ保存パス（OS：Vista、7）
    Public Const JravanCachePath64 As String = "C:\ProgramData\JRA-VAN\Data Lab\Cache"

    'CSVファイル拡張子
    Public Const CSV As String = ".csv"
    '過去CSVファイル
    Public Const OLDCSV As String = "old_"
    'レース情報CSVファイル
    Public Const RA_FileName As String = "RA_INFODATA.csv"
    Public Const SE_FileName As String = "SE_INFODATA.csv"
    'オッズ情報CSVファイル
    Public Const O1_FileName As String = "O1_INFODATA.csv"
    Public Const O2_FileName As String = "O2_INFODATA.csv"
    Public Const O4_FileName As String = "O4_INFODATA.csv"
    '過去レース情報CSVファイル
    Public Const Old_RA_FileName As String = "OLD_RA_INFODATA.csv"
    Public Const Old_SE_FileName As String = "OLD_SE_INFODATA.csv"
    '2010/12/16 d-kobayashi add start
    '日付単位のファイル名分割用
    Public Const Old_RA_FileNameDate As String = "OLD_RA_INFODATA"
    Public Const Old_SE_FileNameDate As String = "OLD_SE_INFODATA"
    '2010/12/16 d-kobayashi add end

    '過去オッズ情報CSVファイル
    Public Const Old_O1_FileName As String = "OLD_O1_INFODATA.csv"
    Public Const Old_O2_FileName As String = "OLD_O2_INFODATA.csv"
    Public Const Old_O4_FileName As String = "OLD_O4_INFODATA.csv"
    '2010/12/16 d-kobayashi add start
    '日付単位のファイル名分割用
    Public Const Old_O1_FileNameDate As String = "OLD_O1_INFODATA"
    Public Const Old_O2_FileNameDate As String = "OLD_O2_INFODATA"
    Public Const Old_O4_FileNameDate As String = "OLD_O4_INFODATA"
    '2010/12/16 d-kobayashi add end
    '画像ファイル
    Public Const IMG_BLACK As String = "\img\black.jpg"
    Public Const IMG_BLUE As String = "\img\blue.jpg"
    Public Const IMG_BROWN As String = "\img\brown.jpg"
    Public Const IMG_GLAY As String = "\img\glay.jpg"
    Public Const IMG_GREEN As String = "\img\green.jpg"
    Public Const IMG_LIGHTBLUE As String = "\img\lightblue.jpg"
    Public Const IMG_ORANGE As String = "\img\orange.jpg"
    Public Const IMG_PINK As String = "\img\pink.jpg"
    Public Const IMG_PURPLE As String = "\img\purple.jpg"
    Public Const IMG_RED As String = "\img\red.jpg"
    Public Const IMG_YELLOW As String = "\img\yellow.jpg"

    'INIファイルセクション名 [TIMER]
    Public Const INI_Ap_Timer As String = "TIMER"
    'INIファイルキー項目 [TIMER]
    Public Const INI_Key_TimerFlg As String = "timerflg"
    Public Const INI_Key_UpdateTime As String = "updatetime"
    Public Const INI_Key_FromTime As String = "fromtime"
    Public Const INI_Key_ToTime As String = "totime"
    'INIファイルセクション名 [JVLink]
    Public Const INI_Ap_JVLink As String = "JVLink"
    'INIファイルキー項目 [JVLink]
    Public Const INI_Key_JVFromTime As String = "jvfromtime"

    '2010/11/15 d-kobayashi add start
    'INIファイルセクション名［KakoDelete]
    Public Const INI_Ap_KakoDelete As String = "KakoDelete"
    'INIファイルキー項目［KakoDelete]
    Public Const INI_Key_KakoDelete As String = "KakoDelete"
    'INIファイルキー項目［KakoDelete］デフォルト値
    Public Const INI_DefaultValue_KakoDelete As String = "0"
    '2010/11/15 d-kobayashi add end

    ' 変動線初期化の数値
    Public Const LINE_X As Integer = 0
    Public Const LINE_Y As Integer = 18

    'アスタリスク
    Public Const ASTERISK As String = "*"

    ' 更新間隔（15分）
    Public Const MIN_15 As Integer = 15
    ' 更新間隔（60分）
    Public Const MIN_60 As Integer = 60

    ' ユーザ権限
    Public Const AUTHORITY As String = "1"

    'サーバー関連
    'メインサーバーＩＰアドレス
    Public Const MainServerAddressIP As String = "153.121.59.92/ninshou/"
    'メインサーバーＵＲＬ
    Public Const MainServerUrl As String = "http://www10318uo.sakura.ne.jp/ninshou/"
    'メインサーバーＨＴＴＰログインユーザーＩＤ
    Public Const MainServerHttpUserId As String = "root"
    'メインサーバーＨＴＴＰログインパスワード
    Public Const MainServerHttpUserPass As String = "vsmyvn3yt3"
    'メインサーバーＦＴＰログインユーザーＩＤ
    Public Const MainServerFtpId As String = "root"
    'メインサーバーＦＴＰログインパスワード
    Public Const MainServerFtpPass As String = "vsmyvn3yt3"


    'サブサーバーアドレス
    Public Const SubServerAddressIP As String = "203.105.84.56/www/htdocs/anasoft/Data/"
    'サブサーバーＵＲＬ
    Public Const SubServerUrl As String = "http://203.105.84.56/anasoft/Data/"
    'サブサーバーＨＴＴＰログインユーザーＩＤ
    Public Const SubServerHttpUserId As String = "anasoft"
    'サブサーバーＨＴＴＰログインパスワード
    Public Const SubServerHttpUserPass As String = "8cdg7Kmg"
    'サブサーバーＦＴＰログインユーザーＩＤ
    Public Const SubServerFtpId As String = "anasoft"
    'サブサーバーＦＴＰログインパスワード
    Public Const SubServerFtpPass As String = "Aw2dKiu4"

    'サーバー接続フラグ
    Public Const serverConnectOff As Boolean = False
    Public Const serverConnectOn As Boolean = True
    '伝送方式
    Public Const FtpUpload As String = "ftp://"
    Public Const HttpUpload As String = "http://"

    'パスワードを忘れた場合のURL
    Public Const ContactURL As String = "http://www10318uo.sakura.ne.jp/help/k/contact.html"
    'ヘルプファイルのURL
    Public Const HelpURL As String = "http://www.ananokami38.com/~anasoft2/help/a/index.html"

    '認証ファイル名称
    Public Const NInshouFileName As String = "NinshouDataK.csv"

    '20140810 アクティベーション用
    Public Const ActivationDir As String = "153.121.59.92/ninshou/activation/"
    'debug start
    Public Const ActivationDirWeb As String = "http://www10318uo.sakura.ne.jp/ninshou/activation/"
    'debug end


    'レジストリデータ保持位置情報
    Public Const rId As String = "0"
    Public Const rPass As String = "1"
    Public Const rAuthority As String = "2"

    'レジスト登録有効有無
    Public Const rCheckOff As String = "0"
    Public Const rCheckOn As String = "1"

    '出走馬データインデックス位置情報定数
    'キー情報格納位置
    Public Const KeyData As String = "0"
    '馬番号／馬名称格納位置
    Public Const UmaNo As String = "1"
    Public Const UmaName As String = "2"
    Public Const KishuName As String = "3"

    'レース情報データインデックス位置
    Public Const RaceDataKey As String = "0"
    Public Const RaceDataRaceDate As String = "1"

    'オールオッズデータ
    Public Const AllOdKeyData As String = "0"
    Public Const AllOdUpTime As String = "1"

    'オッズデータ
    '馬番号位置
    Public Const OdUmaNo As String = "0"

    'オッズ格納位置
    Public Const OdTanshouOddsPos As String = "1"
    Public Const OdFukushouOddsPos As String = "2"
    Public Const OdUmatanOddsPos As String = "3"
    Public Const OdUmarenOddsPos As String = "4"

    '法則成立カウントリスト設定位置
    Public Const HCsetUmaNo As String = "0"
    Public Const HCSetCountKawaPos As String = "1"
    Public Const HCSetCountKawa03Pos As String = "2"
    Public Const HCSetCountYamaPos As String = "3"
    Public Const HCSetCountUchiiriPos As String = "4"
    Public Const HCSetCountRourouPos As String = "5"
    Public Const HCSetCountJikeiUchiiriPos As String = "6"
    Public Const HCSetCountJikeiRourouPos As String = "7"
    Public Const HCSetCountJikeiOogake As String = "8"
    Public Const HCSetCountOogakePos As String = "9"
    Public Const HCSetCountLast80Pos As String = "10"
    Public Const HCSetCountGyakutenPos As String = "11"


    'オッズリスト格納位置
    Public Const OlTnashouOddsPos As String = "1"
    Public Const OlFukushouOddsPos As String = "2"
    Public Const OlUmatanOddsPos As String = "3"
    Public Const OlUmarenOddsPos As String = "4"


    '騎手／馬名リスト設定位置
    Public Const HorseDataKey As String = "0"
    Public Const HorseDataUmaNo As String = "1"
    Public Const HorseDataHorseName As String = "2"
    Public Const HorseDataKishuName As String = "3"


    '単勝／複勝オッズデータインデックス
    Public Const setTFKeyPos As Integer = 0
    Public Const setTFUmaNo As Integer = 1
    Public Const setTFTanOdds As Integer = 2
    Public Const setTFFukuOdds As Integer = 3

    '馬別ポイント保持データインデックス位置情報定数
    '総ポイント設定位置
    Public Const SetPointPos As String = "0"
    '馬番号
    Public Const PointUmaNo As String = "1"

    'オッズ区分
    Public Const TanshouOddsKbn As String = "1"
    Public Const FukushouOddsKbn As String = "2"
    Public Const UmatanOddsKbn As String = "3"
    Public Const UmarenOddsKbn As String = "4"
    '2011/11/08 d-kobayashi add start
    'マジックゾーンのソート順のための新区分
    Public Const UmarenOddsKbn2 As String = "5"
    Public Const UmarenOddsKbn3 As String = "6"
    '2011/11/08 d-kobayashi add end

    '時系列区分
    Public Const EightTime As String = "1"
    Public Const NineTime As String = "2"
    Public Const TenTime As String = "3"



    '法則毎加算ポイント
    Public Const addPointKawa As Integer = 11
    Public Const addPointKawa03 As Integer = 5
    Public Const addPointYama As Integer = 7
    Public Const addPointUchiiri As Integer = 6
    Public Const addPointRourou As Integer = 5
    Public Const addPointJikeiUchiiri As Integer = 8
    Public Const addPointJikeiRourou As Integer = 8
    Public Const addPointJikeiOogake As Integer = 8
    Public Const addPointOogake As Integer = 5
    Public Const addPointLast80 As Integer = 8
    Public Const addPointGyakuten As Integer = 5

    '特別加算ポイント
    Public Const AddSpecialPoint As Integer = 3

    'フラグセット値
    Public Const FlgOn As String = "1"
    Public Const FlgOff As String = "0"

    'インデックス値
    Public Const IndexPos0 As String = "0"
    Public Const IndexPos1 As String = "1"
    Public Const IndexPos2 As String = "2"
    Public Const IndexPos3 As String = "3"
    Public Const IndexPos4 As String = "4"
    Public Const IndexPos5 As String = "5"
    Public Const IndexPos6 As String = "6"
    Public Const IndexPos7 As String = "7"
    Public Const IndexPos8 As String = "8"
    Public Const IndexPos9 As String = "9"
    Public Const IndexPos10 As String = "10"
    Public Const IndexPos11 As String = "11"
    Public Const IndexPos12 As String = "12"
    Public Const IndexPos13 As String = "13"
    Public Const IndexPos14 As String = "14"
    Public Const IndexPos15 As String = "15"
    Public Const IndexPos16 As String = "16"
    Public Const IndexPos17 As String = "17"
    Public Const IndexPos18 As String = "18"
    Public Const IndexPos19 As String = "19"
    Public Const IndexPos20 As String = "20"
    Public Const IndexPos86 As String = "86"

    '数値定数定義
    Public Const Const0 As Integer = 0
    Public Const Const1 As Integer = 1
    Public Const Const2 As Integer = 2
    Public Const Const3 As Integer = 3
    Public Const Const4 As Integer = 4
    Public Const Const5 As Integer = 5
    Public Const Const6 As Integer = 6
    Public Const Const7 As Integer = 7
    Public Const Const8 As Integer = 8
    Public Const Const9 As Integer = 9
    Public Const Const10 As Integer = 10
    Public Const Const11 As Integer = 11
    Public Const Const12 As Integer = 12
    Public Const Const13 As Integer = 13
    Public Const Const14 As Integer = 14
    Public Const Const15 As Integer = 15
    Public Const Const16 As Integer = 16
    Public Const Const17 As Integer = 17
    Public Const Const18 As Integer = 18
    Public Const Const19 As Integer = 19
    Public Const Const20 As Integer = 20
    Public Const Const25 As Integer = 25
    Public Const Const80 As Integer = 800
    Public Const ConstM3 As Integer = -3
    Public Const Const03 As Double = 3.0

    Public Const ConstBairitu As Double = 1.5

    Public Const NotRunOdds As String = "xxxx"

    'Excelファイル名定義
    Public Const strExcelFileAdmin = "verkadmin.xls"
    Public Const strExcelFile = "verk.xls"

End Class