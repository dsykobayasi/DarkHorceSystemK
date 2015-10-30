Public Class clsKeySort
    '対象リストのソート処理を実施する
    '   ListSort:ソート処理コントロール
    '   GetSortKey：ソートキーの設定
    '   SortExecute：リストのソート及びソートキーの削除
    Public Function ListSort(ByVal RunningHorseMstList As ArrayList, ByVal OddsKbn As String) As ArrayList

        '共通ユーティリティークラスのインスタンスを生成
        Dim Util As clsCommonUtil = New clsCommonUtil

        '出走場マスタリストのクローンを返却値リストに設定
        Dim retList As ArrayList = RunningHorseMstList.Clone

        'ソートキー位置情報取得
        Dim SortKey As String() = Util.GetSortKeyPos(OddsKbn)

        'ソートキー設定
        GetSortKey(retList, SortKey, OddsKbn)

        'ソート処理＆ソートキー除外
        SortExecute(retList)

        Return retList

    End Function

    Private Sub GetSortKey(ByRef SortList As ArrayList, ByVal SortKey As String(), ByVal OddsKbn As String)

        'ソートキー付加

        For i = 0 To SortList.Count - 1
            Dim Swork As String = SortList(i)
            Dim SworkH As String() = Swork.Split(","c)

            'ソートキーを設定し、区切り文字"|"（パイプ）でリスト内の情報と接続する
            '2011/11/08 d-kobayashi update
            'If (OddsKbn = CommonConstant.UmarenOddsKbn) Then
            If (OddsKbn = CommonConstant.UmarenOddsKbn) Or (OddsKbn = CommonConstant.UmarenOddsKbn2) Or (OddsKbn = CommonConstant.UmarenOddsKbn3) Then
                '2011/11/08 d-kobayasahi update end
                'マジックゾーンデータに使用するリストの場合、固定で以下の順でソートする
                '１：馬連人気　２：単勝人気　３：複勝人気　４：馬単人気
                SortList.Item(i) = SworkH(SortKey(CommonConstant.Const0)) & "," & _
                                   SworkH(SortKey(CommonConstant.Const1)) & "," & _
                                   SworkH(SortKey(CommonConstant.Const2)) & "," & _
                                   SworkH(SortKey(CommonConstant.Const3)) & "," & "|" & Swork

            Else
                'MsgBox(Swork)
                SortList.Item(i) = SworkH(SortKey(CommonConstant.Const0)) & "," & _
                                   SworkH(SortKey(CommonConstant.Const1)) & "," & _
                                   SworkH(SortKey(CommonConstant.Const2)) & "," & "|" & Swork

            End If
            'MsgBox("ソートキー付加：" & SortList.Item(i))
        Next

    End Sub

    Private Sub SortExecute(ByRef sortList As ArrayList)
        'キーソート
        sortList.Sort()

        'ソートキー除外
        For i = 0 To sortList.Count - 1
            Dim Swork As String = sortList(i)
            Dim SworkH As String() = Swork.Split("|"c)
            sortList.Item(i) = SworkH(1)
            'MsgBox("ソートキー除外後 " & sortList(i))

        Next


    End Sub




End Class
