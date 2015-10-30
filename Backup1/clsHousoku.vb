Public Class clsHousoku
    '共通ユーティティークラスのインスタンス
    Dim commonUtil As clsCommonUtil = New clsCommonUtil

    '単独で成立する法則の該当確認を行う制御ファンクション
    '＜該当法則＞
    '　川の法則
    '　0.3川の法則
    '　山の法則
    '　ラスト80倍の法則
    '　逆転の法則

    Public Sub TandokuHousokuGaitouCheck(ByRef RunHorseList As ArrayList, _
                                         ByRef PointList As ArrayList, _
                                         ByVal OddsKbn As String, _
                                         ByVal TimeKbn As String)

        '法則チェックに使用するオッズ情報を取得
        Dim OddsPos As String = commonUtil.GetOddsPosition(OddsKbn)

        '川の法則該当チェック
        KawanoHousokuGaitouCheck(RunHorseList, PointList, OddsPos, OddsKbn, TimeKbn)

        '0.3川の法則該当チェック
        KawanoHousoku03GaitouCheck(RunHorseList, PointList, OddsPos, OddsKbn, TimeKbn)

        '山の法則該当チェック
        YamanoHousokuGaitouCheck(RunHorseList, PointList, OddsPos, OddsKbn, TimeKbn)

        'マジックゾーンオッズの場合、ラスト80倍及び逆転の法則を実施しない
        If (OddsKbn <> CommonConstant.UmarenOddsKbn) Then
            '単勝オッズの場合、ラスト80倍を実施
            If (OddsKbn = CommonConstant.TanshouOddsKbn) Then
                'ラスト80倍の法則該当チェック
                TanshouOddLast80BaiHousokuGaitouCheck(RunHorseList, PointList, OddsPos, OddsKbn, TimeKbn)
            End If

            '逆転の法則該当チェック
            GyakutenHousokuGaitouCheck(RunHorseList, PointList, OddsPos, OddsKbn, TimeKbn)

        End If



    End Sub

    'オッズ-マジックゾーン間で成立する法則の該当確認を行う制御ファンクション
    '＜該当法則＞
    '　討入りの法則
    '　浪々の法則
    '　大駆けの法則
    Public Sub MzHousokuGaitouCheck(ByRef RunHorseList As ArrayList, _
                                    ByRef MagicZoneList As ArrayList, _
                                    ByRef PointList As ArrayList, _
                                    ByVal OddsKbn As String, _
                                    ByVal TimeKbn As String)

        '討入りの法則該当チェック
        UchiiriHousokuGaitouCheck(RunHorseList, MagicZoneList, PointList, OddsKbn, TimeKbn)

        '浪々の法則該当チェック
        RourouHousokuGaitouCheck(RunHorseList, MagicZoneList, PointList, OddsKbn, TimeKbn)

        '大駆けの法則該当チェック
        OogakeHousokuGaitouCheck(RunHorseList, MagicZoneList, PointList, OddsKbn, TimeKbn)



    End Sub

    '主軸オッズ-次時間軸オッズ間で成立する法則の該当確認を行う制御ファンクション
    '＜該当法則＞
    '　時系列討入りの法則
    '　時系列浪々の法則
    '　時系列大駆けの法則
    Public Sub JikeiretuHousokuGaitouCheck(ByRef RunHorseList As ArrayList, _
                                           ByRef NextRunHorseList As ArrayList, _
                                           ByRef PointList As ArrayList, _
                                           ByVal OddsKbn As String, _
                                           ByVal TimeKbn As String)

        '法則チェックに使用するオッズ情報を取得
        Dim OddsPos As String = commonUtil.GetOddsPosition(OddsKbn)

        '時系列討入りの法則該当チェック
        JikeiretuUchiiriHousokuGaitouCheck(RunHorseList, NextRunHorseList, PointList, OddsKbn, TimeKbn)

        '時系列浪々の法則該当チェック
        JikeiretuRourouHousokuGaitouCheck(RunHorseList, NextRunHorseList, PointList, OddsKbn, TimeKbn, OddsPos)

        '時系列大駆けの法則該当チェック
        JikeiretuOogakeHousokuGaitouCheck(RunHorseList, NextRunHorseList, PointList, OddsKbn, TimeKbn, OddsPos)

    End Sub

    '川の法則該当チェック
    '引数
    ' - 出馬リスト
    ' - ポイント保持リスト
    ' - 予測オッズ位置情報

    Private Sub KawanoHousokuGaitouCheck(ByVal RunHorseList As ArrayList, _
                                         ByRef PointList As ArrayList, _
                                         ByVal OddsPos As String, _
                                         ByVal OddsKbn As String, _
                                         ByVal TimeKbn As String)


        Dim AnaumaDataHairetu As String() = {}
        Dim RunHorseListCount As Integer = RunHorseList.Count - 1
        Dim SeirituFlg As String = CommonConstant.FlgOff
        Dim ReadOdds As String = String.Empty
        Dim ReadHorseNo As String = String.Empty
        Dim BeforeOdds As String = String.Empty
        Dim BeforeHorseNo As String = String.Empty

        '出馬リストの件数分処理を繰り返し
        For i = 0 To RunHorseListCount

            '出馬リストから１レコード取得し、配列へ格納する（デリミタ：","）
            AnaumaDataHairetu = RunHorseList(i).Split(","c)
            'オッズを取得
            ReadOdds = AnaumaDataHairetu(OddsPos)
            '馬番号を取得
            ReadHorseNo = AnaumaDataHairetu(CommonConstant.OdUmaNo)

            '初回レコードの場合、前回情報へセットして次レコードを処理する
            If i = CommonConstant.Const0 Then

                BeforeHorseNo = ReadHorseNo
                BeforeOdds = ReadOdds

                Continue For

            End If

            'オッズの比較
            If (BeforeOdds = ReadOdds) Then
                '同一オッズの場合
                'フラグセット＆ポイント加算を実施する
                '引数
                ' - 法則成立馬番号
                ' - オッズ区分
                ' - 時系列区分
                ' - 法則更新位置情報
                ' - 馬別法則成立ポイント保持リスト
                ' - 加算ポイント
                SetHousokuFlgAndAddPoint(BeforeHorseNo, _
                                         OddsKbn, _
                                         TimeKbn, _
                                         CommonConstant.HCSetCountKawaPos, _
                                         PointList, _
                                         CommonConstant.addPointKawa)

                '法律成立フラグをオンに設定
                SeirituFlg = CommonConstant.FlgOn

                '最終レコードの場合、今回レコードを処理する。
                If (i = RunHorseListCount) Then
                    '同一オッズの場合
                    'フラグセット＆ポイント加算を実施する
                    '引数
                    ' - 法則成立馬番号
                    ' - オッズ区分
                    ' - 時系列区分
                    ' - 法則更新位置情報
                    ' - 馬別法則成立ポイント保持リスト
                    ' - 加算ポイント
                    SetHousokuFlgAndAddPoint(ReadHorseNo, _
                                             OddsKbn, _
                                             TimeKbn, _
                                             CommonConstant.HCSetCountKawaPos, _
                                             PointList, _
                                             CommonConstant.addPointKawa)

                End If



            Else
                'オッズ相違の場合
                If SeirituFlg Then
                    '前オッズと法則が成立している場合

                    '同一オッズの場合
                    'フラグセット＆ポイント加算を実施する
                    '引数
                    ' - 法則成立馬番号
                    ' - オッズ区分
                    ' - 時系列区分
                    ' - 法則更新位置情報
                    ' - 馬別法則成立ポイント保持リスト
                    ' - 加算ポイント
                    SetHousokuFlgAndAddPoint(BeforeHorseNo, _
                                             OddsKbn, _
                                             TimeKbn, _
                                             CommonConstant.HCSetCountKawaPos, _
                                             PointList, _
                                             CommonConstant.addPointKawa)

                    '法則成立フラグをオフに設定
                    SeirituFlg = CommonConstant.FlgOff

                End If

            End If

            '今回情報を前回情報に設定
            BeforeHorseNo = ReadHorseNo
            BeforeOdds = ReadOdds

        Next

    End Sub
    '0.3川の法則
    '引数
    ' - 出馬リスト
    ' - ポイント保持リスト
    ' - 予測オッズ位置情報

    Private Sub KawanoHousoku03GaitouCheck(ByVal RunHorseList As ArrayList, _
                                          ByRef PointList As ArrayList, _
                                          ByVal OddsPos As String, _
                                          ByVal OddsKbn As String, _
                                          ByVal TimeKbn As String)


        Dim AnaumaDataHairetu As String() = {}
        Dim CloneAnaumaDataHairetu As String() = {}
        Dim SetAnaumaData As String = String.Empty
        Dim RunHorseListCount As Integer = RunHorseList.Count - 1
        Dim SeirituFlg As String = CommonConstant.FlgOff

        Dim ReadOdds As Integer = 0
        Dim ReadHorseNo As String = String.Empty
        Dim BeforeOdds As Integer = 0
        Dim BeforeHorseNo As String = String.Empty

        Dim GaitouList As New ArrayList

        Dim CloneRunHorseList As ArrayList = RunHorseList.Clone

        For i = 0 To RunHorseListCount

            '出馬リストから１レコード取得し、配列へ格納する（デリミタ：","）
            AnaumaDataHairetu = RunHorseList(i).Split(","c)
            'オッズ未設定の場合は次レコードを処理
            If Not IsNumeric(AnaumaDataHairetu(OddsPos)) Then

                '今回データが欠馬の場合
                If SeirituFlg Then
                    '法則が成立していれば、前回レコードを書き込む
                    'フラグセット＆ポイント加算を実施する
                    '引数
                    ' - 法則成立馬番号
                    ' - オッズ区分
                    ' - 時系列区分
                    ' - 法則更新位置情報
                    ' - 馬別法則成立ポイント保持リスト
                    ' - 加算ポイント
                    'MsgBox("前回データ更新In")

                    SetHousokuFlgAndAddPoint(BeforeHorseNo, _
                                             OddsKbn, _
                                             TimeKbn, _
                                             CommonConstant.HCSetCountKawa03Pos, _
                                             PointList, _
                                             CommonConstant.addPointKawa03)


                    ReadOdds = CommonConstant.Const0
                    BeforeOdds = CommonConstant.Const0

                    SeirituFlg = CommonConstant.FlgOff

                End If

                Continue For

            End If

            'オッズを取得
            ReadOdds = AnaumaDataHairetu(OddsPos)
            '馬番号を取得
            ReadHorseNo = AnaumaDataHairetu(CommonConstant.OdUmaNo)

            '初回レコードの場合、前回情報へセットして次レコードを処理する
            If i = CommonConstant.Const0 Then

                BeforeHorseNo = ReadHorseNo
                BeforeOdds = ReadOdds

                Continue For

            End If



            Dim Anser As Integer = ReadOdds - BeforeOdds

            '前回オッズ－今回オッズ差が－3～3の間且つ0以外であるかを判定
            If (Anser >= CommonConstant.ConstM3 And Anser <= CommonConstant.Const3 And Anser <> CommonConstant.Const0) Then
                'オッズ差が－3～3の間且つ0以外の場合
                'フラグセット＆ポイント加算を実施する
                '引数
                ' - 法則成立馬番号
                ' - オッズ区分
                ' - 時系列区分
                ' - 法則更新位置情報
                ' - 馬別法則成立ポイント保持リスト
                ' - 加算ポイント
                'MsgBox("前回データ更新In")

                SetHousokuFlgAndAddPoint(BeforeHorseNo, _
                                         OddsKbn, _
                                         TimeKbn, _
                                         CommonConstant.HCSetCountKawa03Pos, _
                                         PointList, _
                                         CommonConstant.addPointKawa03)

                '法則成立フラグをオンに設定
                SeirituFlg = CommonConstant.FlgOn

                '最終レコードの場合、今回レコードを処理する。
                If (i = RunHorseListCount) Then
                    '同一オッズの場合
                    'フラグセット＆ポイント加算を実施する
                    '引数
                    ' - 法則成立馬番号
                    ' - オッズ区分
                    ' - 時系列区分
                    ' - 法則更新位置情報
                    ' - 馬別法則成立ポイント保持リスト
                    ' - 加算ポイント
                    'MsgBox("前回データ更新In")

                    SetHousokuFlgAndAddPoint(ReadHorseNo, _
                                             OddsKbn, _
                                             TimeKbn, _
                                             CommonConstant.HCSetCountKawa03Pos, _
                                             PointList, _
                                             CommonConstant.addPointKawa03)

                End If

            Else

                'オッズ相違の場合
                If SeirituFlg Then
                    '前オッズと法則が成立している場合

                    'フラグセット＆ポイント加算を実施する
                    '引数
                    ' - 法則成立馬番号
                    ' - オッズ区分
                    ' - 時系列区分
                    ' - 法則更新位置情報
                    ' - 馬別法則成立ポイント保持リスト
                    ' - 加算ポイント
                    'MsgBox("前回データ更新In")

                    SetHousokuFlgAndAddPoint(BeforeHorseNo, _
                                             OddsKbn, _
                                             TimeKbn, _
                                             CommonConstant.HCSetCountKawa03Pos, _
                                             PointList, _
                                             CommonConstant.addPointKawa03)

                    '法則成立フラグをオンに設定
                    SeirituFlg = CommonConstant.FlgOff


                End If



            End If

            BeforeHorseNo = ReadHorseNo
            BeforeOdds = ReadOdds




            'For j = 0 To RunHorseListCount

            '    'クローンの出馬リストから１レコード取得
            '    CloneAnaumaDataHairetu = CloneRunHorseList(j).split(","c)

            '    'オッズ未設定の場合は次レコードを処理
            '    If Not IsNumeric(CloneAnaumaDataHairetu(OddsPos)) Then

            '        Continue For

            '    End If

            '    'クローンからオッズを取得
            '    CloneReadOdds = CloneAnaumaDataHairetu(OddsPos)
            '    'クローンから馬番号を取得
            '    CloneReadHorseNo = CloneAnaumaDataHairetu(CommonConstant.OdUmaNo)

            '    Dim Anser As Double = ReadOdds - CloneReadOdds

            '    If (Anser >= Double.Parse(CommonConstant.ConstM3) And Anser <= CommonConstant.Const03 And Anser <> CommonConstant.Const0) Then

            '        'MsgBox("今回RecOdds:" & AnaumaDataHairetu(OddsPos) & "　前回RecOdds:" & AnaumaDataWorkHairetu(OddsPos))

            '        '前回レコードと今回レコードのオッズが同一の場合は「0.3川の法則」の成立とする
            '        '一度法則が成立しているものに関しては処理を実施しない
            '        If Not GaitouList.Contains(CloneReadHorseNo) Then

            '            'フラグセット＆ポイント加算を実施する
            '            '引数
            '            ' - 法則成立馬番号
            '            ' - オッズ区分
            '            ' - 時系列区分
            '            ' - 法則更新位置情報
            '            ' - 馬別法則成立ポイント保持リスト
            '            ' - 加算ポイント
            '            'MsgBox("前回データ更新In")

            '            SetHousokuFlgAndAddPoint(CloneReadHorseNo, _
            '                                     OddsKbn, _
            '                                     TimeKbn, _
            '                                     CommonConstant.HCSetCountKawa03Pos, _
            '                                     PointList, _
            '                                     CommonConstant.addPointKawa03)
            '            'リストへ法則の成立した馬番号を設定
            '            GaitouList.Add(CloneReadHorseNo)

            '            SeirituFlg = CommonConstant.FlgOn

            '        End If

            '    End If

            'Next




        Next

    End Sub
    '山の法則該当チェック
    '引数
    ' - 出馬リスト
    ' - ポイント保持リスト
    ' - 予測オッズ位置情報

    Private Sub YamanoHousokuGaitouCheck(ByVal RunHorseList As ArrayList, _
                                        ByRef PointList As ArrayList, _
                                        ByVal OddsPos As String, _
                                        ByVal OddsKbn As String, _
                                        ByVal TimeKbn As String)

        Dim AnaumaData As String = ""
        Dim AnaumaDataHairetu As String() = {}
        Dim AnaumaDataWorkHairetu1 As String() = {}
        Dim AnaumaDataWorkHairetu2 As String() = {}
        Dim SetAnaumaData1 As String = ""
        Dim SetAnaumaData2 As String = ""
        Dim RunHorseListCount As Integer = RunHorseList.Count - 1

        Dim readOdds As Double = 0
        Dim readWorkOdds As Double = 0

        For i = 0 To RunHorseListCount

            '出馬リストから１レコード取得し、配列へ納する（デリミタ：","）
            AnaumaData = RunHorseList(i)
            AnaumaDataHairetu = AnaumaData.Split(","c)
            'MsgBox("山の法則　馬情報配列数：" & AnaumaDataHairetu.Length)

            If (i = CommonConstant.Const0) Then

                AnaumaDataWorkHairetu1 = AnaumaDataHairetu

                Continue For

            End If


            ''馬単／馬連オッズの場合、一番人気馬にオッズが設定されていない為、0として計算する
            If Not IsNumeric(AnaumaDataHairetu(OddsPos)) Then
                AnaumaDataWorkHairetu2 = AnaumaDataWorkHairetu1
                AnaumaDataWorkHairetu1 = AnaumaDataHairetu

                Continue For

            Else
                readOdds = Double.Parse(AnaumaDataHairetu(OddsPos))

            End If

            '馬単／馬連オッズの場合、一番人気馬にオッズが設定されていない為、0として計算する
            If Not IsNumeric(AnaumaDataWorkHairetu1(OddsPos)) Then
                AnaumaDataWorkHairetu2 = AnaumaDataWorkHairetu1
                AnaumaDataWorkHairetu1 = AnaumaDataHairetu
                'readWorkOdds = 0

                Continue For

            Else
                readWorkOdds = Double.Parse(AnaumaDataWorkHairetu1(OddsPos))
            End If



            Dim Anser As Double = readOdds / readWorkOdds

            'MsgBox("割算結果 ： " & Anser)

            If (Anser >= CommonConstant.ConstBairitu) Then
                '今回レコード÷前回レコードの結果が１．５倍以上の場合は「山の法則」の成立とする
                '前々回レコード及び前回レコードに対して更新を行う

                '前回レコードのフラグセット＆ポイント加算を実施する
                '引数
                ' - 法則成立馬番号
                ' - オッズ区分
                ' - 時系列区分
                ' - 法則更新位置情報
                ' - 馬別法則成立ポイント保持リスト
                ' - 加算ポイント
                SetHousokuFlgAndAddPoint(AnaumaDataWorkHairetu1(CommonConstant.OdUmaNo), _
                         OddsKbn, _
                         TimeKbn, _
                         CommonConstant.HCSetCountYamaPos, _
                         PointList, _
                         CommonConstant.addPointYama)


                If (i > CommonConstant.Const1) Then
                    '前々回レコードのフラグセット＆ポイント加算を実施する
                    '引数
                    ' - 法則成立馬番号
                    ' - オッズ区分
                    ' - 時系列区分
                    ' - 法則更新位置情報
                    ' - 馬別法則成立ポイント保持リスト
                    ' - 加算ポイント
                    SetHousokuFlgAndAddPoint(AnaumaDataWorkHairetu2(CommonConstant.OdUmaNo), _
                                             OddsKbn, _
                                             TimeKbn, _
                                             CommonConstant.HCSetCountYamaPos, _
                                             PointList, _
                                             CommonConstant.addPointYama)

                End If

            End If

            '前回レコードを前々回レコード格納エリアへ代入する
            '今回レコードを前回レコード格納エリアへ代入する
            AnaumaDataWorkHairetu2 = AnaumaDataWorkHairetu1
            AnaumaDataWorkHairetu1 = AnaumaDataHairetu

        Next


    End Sub
    'ラスト80倍の法則該当チェック
    '引数
    ' - 出馬リスト
    ' - ポイント保持リスト
    Private Sub TanshouOddLast80BaiHousokuGaitouCheck(ByVal RunHorseList As ArrayList, _
                                                     ByRef PointList As ArrayList, _
                                                     ByVal OddsPos As String, _
                                                     ByVal OddsKbn As String, _
                                                     ByVal TimeKbn As String)

        Dim RunHorseListLastIndex As Integer = RunHorseList.Count - 1

        Dim SetAnaumaData As String = ""
        Dim AnaumaData As String = ""
        Dim AnaumaDataHairetu As String() = {}
        Dim AnaumaNextData As String = ""
        Dim AnaumaNextDataHairetu As String() = {}

        Do While (1)
            '出走馬リストループ
            '出馬リストから１レコード取得し、配列へ納する（デリミタ：","）
            AnaumaData = RunHorseList(RunHorseListLastIndex)
            AnaumaDataHairetu = AnaumaData.Split(","c)
            'MsgBox("ラスト８０倍の法則　馬情報配列数：" & AnaumaDataHairetu.Length)

            If IsNumeric(AnaumaDataHairetu(CommonConstant.OdTanshouOddsPos)) Then
                'オッズが数値の場合、処理を抜ける
                Exit Do
            End If
            '最終馬のオッズが数値以外だった場合、その上の馬を対象とする
            RunHorseListLastIndex = RunHorseListLastIndex - 1
        Loop

        If (AnaumaDataHairetu(CommonConstant.OdTanshouOddsPos) >= CommonConstant.Const80) Then
            '最終人気馬の単勝オッズが80倍以上の場合、「ラスト80倍の法則」成立
            'フラグセット＆ポイント加算を実施する
            '引数
            ' - 法則成立馬番号
            ' - オッズ区分
            ' - 時系列区分
            ' - 法則更新位置情報
            ' - 馬別法則成立ポイント保持リスト
            ' - 加算ポイント
            SetHousokuFlgAndAddPoint(AnaumaDataHairetu(CommonConstant.OdUmaNo), _
                         OddsKbn, _
                         TimeKbn, _
                         CommonConstant.HCSetCountLast80Pos, _
                         PointList, _
                         CommonConstant.addPointLast80)

        End If


    End Sub

    '逆転の法則該当チェック
    '引数
    ' - 出馬リスト
    ' - ポイント保持リスト
    Private Sub GyakutenHousokuGaitouCheck(ByVal RunHorseList As ArrayList, _
                                          ByRef PointList As ArrayList, _
                                          ByVal OddsPos As String, _
                                          ByVal OddsKbn As String, _
                                          ByVal TimeKbn As String)

        Dim RunHorseListCount As Integer = RunHorseList.Count - 1

        Dim SetAnaumaData As String = ""
        Dim AnaumaData As String = ""
        Dim AnaumaDataHairetu As String() = {}
        Dim AnaumaNextData As String = ""
        Dim AnaumaNextDataHairetu As String() = {}
        '出走馬リストループ
        For i = 0 To RunHorseListCount

            '出馬リストから１レコード取得し、配列へ納する（デリミタ：","）
            AnaumaData = RunHorseList(i)
            AnaumaDataHairetu = AnaumaData.Split(","c)
            'MsgBox("逆転の法則　馬情報配列数：" & AnaumaDataHairetu.Length)

            If Not IsNumeric(AnaumaDataHairetu(CommonConstant.OdTanshouOddsPos)) Or _
               Not IsNumeric(AnaumaDataHairetu(CommonConstant.OdFukushouOddsPos)) Then
                'オッズに数値以外の値が設定されていた場合
                Continue For

            End If


            If (AnaumaDataHairetu(CommonConstant.OdTanshouOddsPos) <= AnaumaDataHairetu(CommonConstant.OdFukushouOddsPos)) Then
                '単勝オッズが<=複勝オッズの場合、「逆転の法則」成立

                'フラグセット＆ポイント加算を実施する
                '引数
                ' - 法則成立馬番号
                ' - オッズ区分
                ' - 時系列区分
                ' - 法則更新位置情報
                ' - 馬別法則成立ポイント保持リスト
                ' - 加算ポイント
                SetHousokuFlgAndAddPoint(AnaumaDataHairetu(CommonConstant.OdUmaNo), _
                                         OddsKbn, _
                                         TimeKbn, _
                                         CommonConstant.HCSetCountGyakutenPos, _
                                         PointList, _
                                         CommonConstant.addPointGyakuten)

            End If

            'MsgBox(AnaumaDataHairetu.Length)

            'MsgBox("Index0：" & AnaumaDataHairetu(0) & "Index1：" & AnaumaDataHairetu(1) & "Index2：" & _
            '       AnaumaDataHairetu(2) & "Index3：" & AnaumaDataHairetu(3))

            If Not IsNumeric(AnaumaDataHairetu(CommonConstant.OdUmarenOddsPos)) Or _
               Not IsNumeric(AnaumaDataHairetu(CommonConstant.OdUmatanOddsPos)) Then
                'オッズに数値以外の値が設定されていた場合
                Continue For

            End If

            If (AnaumaDataHairetu(CommonConstant.OdTanshouOddsPos) <= AnaumaDataHairetu(CommonConstant.OdUmarenOddsPos) And _
                AnaumaDataHairetu(CommonConstant.OdUmarenOddsPos) <= AnaumaDataHairetu(CommonConstant.OdUmatanOddsPos)) Then
                '単勝オッズ<=馬連オッズ<=馬単オッズの場合、「逆転の法則」成立

                'フラグセット＆ポイント加算を実施する
                '引数
                ' - 法則成立馬番号
                ' - オッズ区分
                ' - 時系列区分
                ' - 法則更新位置情報
                ' - 馬別法則成立ポイント保持リスト
                ' - 加算ポイント
                SetHousokuFlgAndAddPoint(AnaumaDataHairetu(CommonConstant.OdUmaNo), _
                                         OddsKbn, _
                                         TimeKbn, _
                                         CommonConstant.HCSetCountGyakutenPos, _
                                         PointList, _
                                         CommonConstant.addPointGyakuten)

            End If

        Next

    End Sub

    '討ち入りの法則該当チェック
    '引数
    ' - 出馬リスト
    ' - マジックゾーンリスト
    ' - ポイント保持リスト
    ' - 予測オッズ位置情報

    Private Sub UchiiriHousokuGaitouCheck(ByVal RunHorseList As ArrayList, _
                                         ByVal MagicZoneList As ArrayList, _
                                         ByRef PointList As ArrayList, _
                                         ByVal OddsKbn As String, _
                                         ByVal TimeKbn As String)

        Dim RunHorseListCount As Integer = RunHorseList.Count - 1

        Dim SetAnaumaData As String = ""
        Dim AnaumaData As String = ""
        Dim AnaumaDataHairetu As String() = {}
        Dim AnaumaNextData As String = ""
        Dim AnaumaNextDataHairetu As String() = {}
        For i = 4 To RunHorseListCount
            '出馬リストから１レコード取得し、配列へ納する（デリミタ：","）
            AnaumaData = RunHorseList(i)
            AnaumaDataHairetu = AnaumaData.Split(","c)

            '次時系列データリストループ
            For j = 0 To CommonConstant.Const3
                AnaumaNextData = MagicZoneList(j)
                AnaumaNextDataHairetu = AnaumaNextData.Split(","c)

                If (AnaumaDataHairetu(CommonConstant.OdUmaNo) = AnaumaNextDataHairetu(CommonConstant.OdUmaNo)) Then
                    '馬番号が一致した場合、次時系列にて上位４位以内に入っているとして、「討ち入りの法則」成立

                    'フラグセット＆ポイント加算を実施する
                    '引数
                    ' - 法則成立馬番号
                    ' - オッズ区分
                    ' - 時系列区分
                    ' - 法則更新位置情報
                    ' - 馬別法則成立ポイント保持リスト
                    ' - 加算ポイント
                    SetHousokuFlgAndAddPoint(AnaumaDataHairetu(CommonConstant.OdUmaNo), _
                                             OddsKbn, _
                                             TimeKbn, _
                                             CommonConstant.HCSetCountUchiiriPos, _
                                             PointList, _
                                             CommonConstant.addPointUchiiri)

                End If

            Next

        Next
    End Sub

    '浪々の法則該当チェック
    '引数
    ' - 出馬リスト
    ' - マジックゾーンリスト
    ' - ポイント保持リスト
    ' - 予測オッズ位置情報
    Private Sub RourouHousokuGaitouCheck(ByVal RunHorseList As ArrayList, _
                                         ByVal MagicZoneList As ArrayList, _
                                         ByRef PointList As ArrayList, _
                                         ByVal OddsKbn As String, _
                                         ByVal TimeKbn As String)

        Dim MagicZoneListCount As Integer = MagicZoneList.Count - 1

        Dim SetAnaumaData As String = ""
        Dim AnaumaData As String = ""
        Dim AnaumaDataHairetu As String() = {}
        Dim AnaumaNextData As String = ""
        Dim AnaumaNextDataHairetu As String() = {}
        For i = 0 To CommonConstant.Const3
            '出馬リストから１レコード取得し、配列へ納する（デリミタ：","）
            AnaumaData = RunHorseList(i)
            AnaumaDataHairetu = AnaumaData.Split(","c)

            'マジックゾーンデータリストループ
            For j = 4 To MagicZoneListCount
                AnaumaNextData = MagicZoneList(j)
                AnaumaNextDataHairetu = AnaumaNextData.Split(","c)

                If (AnaumaDataHairetu(CommonConstant.OdUmaNo) = AnaumaNextDataHairetu(CommonConstant.OdUmaNo)) Then
                    '馬番号が一致した場合、マジックゾーンにて５位以下に入っているとして、「浪々の法則」成立

                    '法則成立カウント＆ポイント加算を実施する
                    'フラグセット＆ポイント加算を実施する
                    '引数
                    ' - 法則成立馬番号
                    ' - オッズ区分
                    ' - 時系列区分
                    ' - 法則更新位置情報
                    ' - 馬別法則成立ポイント保持リスト
                    ' - 加算ポイント
                    SetHousokuFlgAndAddPoint(AnaumaDataHairetu(CommonConstant.OdUmaNo), _
                                             OddsKbn, _
                                             TimeKbn, _
                                             CommonConstant.HCSetCountRourouPos, _
                                             PointList, _
                                             CommonConstant.addPointRourou)

                End If

            Next

        Next

    End Sub

    '大駆けの法則該当チェック
    '引数
    ' - 出馬リスト
    ' - マジックゾーンリスト
    ' - ポイント保持リスト
    Private Sub OogakeHousokuGaitouCheck(ByVal RunHorseList As ArrayList, _
                                         ByVal MagicZoneList As ArrayList, _
                                         ByRef PointList As ArrayList, _
                                         ByVal OddsKbn As String, _
                                         ByVal TimeKbn As String)

        Dim RunHorseListCount As Integer = RunHorseList.Count - 1
        Dim MagicZoneListCount As Integer = MagicZoneList.Count - 1

        Dim SetAnaumaData As String = ""
        Dim AnaumaData As String = ""
        Dim AnaumaDataHairetu As String() = {}
        Dim AnaumaNextData As String = ""
        Dim AnaumaNextDataHairetu As String() = {}
        '出走馬リストループ
        For i = 0 To RunHorseListCount

            '出馬リストから１レコード取得し、配列へ納する（デリミタ：","）
            AnaumaData = RunHorseList(i)
            AnaumaDataHairetu = AnaumaData.Split(","c)

            'マジックゾーンリストループ
            For j = 0 To MagicZoneListCount

                AnaumaNextData = MagicZoneList(j)
                AnaumaNextDataHairetu = AnaumaNextData.Split(","c)

                If (AnaumaDataHairetu(CommonConstant.OdUmaNo) = AnaumaNextDataHairetu(CommonConstant.OdUmaNo)) Then
                    '馬番号が一致した場合、ポイント差を判定する

                    If ((j - i) >= CommonConstant.Const3 Or (j - i) <= CommonConstant.ConstM3) Then
                        'ポイント差が３以上または－３以下だった場合、「大駆けの法則」成立

                        '法則成立カウント＆ポイント加算を実施する
                        'フラグセット＆ポイント加算を実施する
                        '引数
                        ' - 法則成立馬番号
                        ' - オッズ区分
                        ' - 時系列区分
                        ' - 法則更新位置情報
                        ' - 馬別法則成立ポイント保持リスト
                        ' - 加算ポイント
                        SetHousokuFlgAndAddPoint(AnaumaDataHairetu(CommonConstant.OdUmaNo), _
                                                 OddsKbn, _
                                                 TimeKbn, _
                                                 CommonConstant.HCSetCountOogakePos, _
                                                 PointList, _
                                                 CommonConstant.addPointOogake)

                    End If

                End If

            Next

        Next

    End Sub


    '時系列討ち入りの法則該当チェック
    '引数
    ' - 出馬リスト
    ' - 次時系列出馬リスト
    ' - ポイント保持リスト
    ' - 予測オッズ位置情報

    Private Sub JikeiretuUchiiriHousokuGaitouCheck(ByRef RunHorseList As ArrayList, _
                                                  ByRef NextRunHorseList As ArrayList, _
                                                  ByRef PointList As ArrayList, _
                                                  ByVal OddsKbn As String, _
                                                  ByVal TimeKbn As String)


        Dim RunHorseListCount As Integer = RunHorseList.Count - 1

        Dim SetAnaumaData As String = ""
        Dim AnaumaData As String = ""
        Dim AnaumaDataHairetu As String() = {}
        Dim AnaumaNextData As String = ""
        Dim AnaumaNextDataHairetu As String() = {}
        For i = 4 To RunHorseListCount
            '出馬リストから１レコード取得し、配列へ納する（デリミタ：","）
            AnaumaData = RunHorseList(i)
            AnaumaDataHairetu = AnaumaData.Split(","c)

            '次時系列データリストループ
            For j = 0 To CommonConstant.Const3
                AnaumaNextData = NextRunHorseList(j)
                AnaumaNextDataHairetu = AnaumaNextData.Split(","c)

                If (AnaumaDataHairetu(CommonConstant.OdUmaNo) = AnaumaNextDataHairetu(CommonConstant.OdUmaNo)) Then
                    '馬番号が一致した場合、次時系列にて上位４位以内に入っているとして、「時系列討ち入りの法則」成立

                    'フラグセット＆ポイント加算を実施する
                    '引数
                    ' - 法則成立馬番号
                    ' - オッズ区分
                    ' - 時系列区分
                    ' - 法則更新位置情報
                    ' - 馬別法則成立ポイント保持リスト
                    ' - 加算ポイント
                    SetHousokuFlgAndAddPoint(AnaumaDataHairetu(CommonConstant.OdUmaNo), _
                                             OddsKbn, _
                                             TimeKbn, _
                                             CommonConstant.HCSetCountJikeiUchiiriPos, _
                                             PointList, _
                                             CommonConstant.addPointJikeiUchiiri)
                End If

            Next

        Next
    End Sub


    '時系列浪々の法則該当チェック
    '引数
    ' - 出馬リスト
    ' - 次時系列出馬リスト
    ' - ポイント保持リスト
    ' - 予測オッズ位置情報
    Private Sub JikeiretuRourouHousokuGaitouCheck(ByRef RunHorseList As ArrayList, _
                                                 ByRef NextRunHorseList As ArrayList, _
                                                 ByRef PointList As ArrayList, _
                                                 ByVal OddsKbn As String, _
                                                 ByVal TimeKbn As String, _
                                                 ByVal OddsPos As String)

        Dim RunHorseNextListCount As Integer = NextRunHorseList.Count - 1

        Dim SetAnaumaData As String = ""
        Dim AnaumaData As String = ""
        Dim AnaumaDataHairetu As String() = {}
        Dim AnaumaNextData As String = ""
        Dim AnaumaNextDataHairetu As String() = {}
        For i = 0 To CommonConstant.Const3
            '出馬リストから１レコード取得し、配列へ納する（デリミタ：","）
            AnaumaData = RunHorseList(i)
            AnaumaDataHairetu = AnaumaData.Split(","c)

            'マジックゾーンデータリストループ
            For j = 4 To RunHorseNextListCount
                AnaumaNextData = NextRunHorseList(j)
                AnaumaNextDataHairetu = AnaumaNextData.Split(","c)

                If ((AnaumaDataHairetu(CommonConstant.OdUmaNo) = AnaumaNextDataHairetu(CommonConstant.OdUmaNo)) And _
                    IsNumeric(AnaumaDataHairetu(OddsPos)) And IsNumeric(AnaumaNextDataHairetu(OddsPos))) Then
                    '馬番号が一致した場合、マジックゾーンにて５位以下に入っているとして、「時系列浪々の法則」成立

                    '法則成立カウント＆ポイント加算を実施する
                    '引数
                    ' - 法則成立馬番号
                    ' - オッズ区分
                    ' - 時系列区分
                    ' - 法則更新位置情報
                    ' - 馬別法則成立ポイント保持リスト
                    ' - 加算ポイント
                    SetHousokuFlgAndAddPoint(AnaumaDataHairetu(CommonConstant.OdUmaNo), _
                                             OddsKbn, _
                                             TimeKbn, _
                                             CommonConstant.HCSetCountJikeiRourouPos, _
                                             PointList, _
                                             CommonConstant.addPointJikeiRourou)

                End If

            Next

        Next

    End Sub

    '時系列大駆けの法則該当チェック
    '引数
    ' - 出馬リスト
    ' - 次時系列出馬リスト
    ' - ポイント保持リスト
    Private Sub JikeiretuOogakeHousokuGaitouCheck(ByRef RunHorseList As ArrayList, _
                                                 ByRef NextRunHorseList As ArrayList, _
                                                 ByRef PointList As ArrayList, _
                                                 ByVal OddsKbn As String, _
                                                 ByVal TimeKbn As String, _
                                                 ByVal OddsPos As String)

        Dim RunHorseListCount As Integer = RunHorseList.Count - 1
        Dim NextRunHorseListCount As Integer = NextRunHorseList.Count - 1

        Dim SetAnaumaData As String = String.Empty
        Dim AnaumaData As String = String.Empty
        Dim AnaumaDataHairetu As String() = {}
        Dim AnaumaNextData As String = String.Empty
        Dim AnaumaNextDataHairetu As String() = {}

        '出走馬リストループ
        For i = 0 To RunHorseListCount

            '出馬リストから１レコード取得し、配列へ納する（デリミタ：","）
            AnaumaData = RunHorseList(i)
            AnaumaDataHairetu = AnaumaData.Split(","c)

            'マジックゾーンリストループ
            For j = 0 To NextRunHorseListCount

                AnaumaNextData = NextRunHorseList(j)
                AnaumaNextDataHairetu = AnaumaNextData.Split(","c)

                If (AnaumaDataHairetu(CommonConstant.OdUmaNo) = AnaumaNextDataHairetu(CommonConstant.OdUmaNo)) Then
                    '馬番号が一致した場合、ポイント差を判定する

                    If (((j - i) >= CommonConstant.Const3 Or (j - i) <= CommonConstant.ConstM3) And _
                        IsNumeric(AnaumaDataHairetu(OddsPos)) And IsNumeric(AnaumaNextDataHairetu(OddsPos))) Then
                        'ポイント差が３以上または－３以下だった場合、「時系列大駆けの法則」成立

                        '法則成立カウント＆ポイント加算を実施する
                        '引数
                        ' - 法則成立馬番号
                        ' - オッズ区分
                        ' - 時系列区分
                        ' - 法則更新位置情報
                        ' - 馬別法則成立ポイント保持リスト
                        ' - 加算ポイント
                        SetHousokuFlgAndAddPoint(AnaumaDataHairetu(CommonConstant.OdUmaNo), _
                                                 OddsKbn, _
                                                 TimeKbn, _
                                                 CommonConstant.HCSetCountJikeiOogake, _
                                                 PointList, _
                                                 CommonConstant.addPointJikeiOogake)

                    End If

                End If

            Next

        Next

    End Sub

    '法則該当時のフラグセット＆ポイント加算処理を実施
    '引数
    ' - 馬番号
    ' - オッズ区分
    ' - 時系列区分
    ' - カウンター加算位置情報
    ' - 穴馬情報（配列）
    Private Sub SetHousokuFlgAndAddPoint(ByVal UmaNo As String, _
                                         ByVal OddsKbn As String, _
                                         ByVal TimeKbn As String, _
                                         ByVal CountPos As String, _
                                         ByRef PointList As ArrayList, _
                                         ByVal addPoint As Integer _
                                         )



        Dim HousokuCounter As String() = {}


        'ハッシュテーブルから該当のデータを取得
        HousokuCounter = HousokuTableGetter(UmaNo, OddsKbn, TimeKbn)

        '配列が空の場合
        If (HousokuCounter.Length = 0) Then
            '法則成立カウンターを初期化
            HousokuCounter = setHairetuInit()
            '馬番号を設定
            HousokuCounter(CommonConstant.HCsetUmaNo) = UmaNo

        End If

        '該当法則カウンターを加算
        HousokuCounter(CountPos) = Integer.Parse(HousokuCounter(CountPos)) + CommonConstant.Const1
        'ハッシュテーブルへ設定する
        HousokuTableSetter(UmaNo, OddsKbn, TimeKbn, HousokuCounter)


        '馬別ポイントリストの更新
        'ポイント保持リストから該当のデータを取得
        For i = 0 To PointList.Count - 1
            Dim PointData As String() = PointList(i)
            If (UmaNo = PointData(0)) Then
                'MsgBox("ポイントリスト　index0：" & PointData(0) & "index1：" & PointData(1))

                'ポイントを加算
                PointData(CommonConstant.IndexPos1) = Integer.Parse(PointData(CommonConstant.IndexPos1)) + addPoint
                'リストを更新
                PointList.Item(i) = PointData
                Exit For
            End If


        Next


    End Sub

    '法則カウンターの初期化
    Private Function setHairetuInit() As String()

        Dim retHairetu As String() = {"", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"}

        Return retHairetu
    End Function

End Class
