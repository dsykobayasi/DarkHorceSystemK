Friend Class clsCSVOutput

    ' 機能　　 : レース詳細のListを作成する
    ' 引き数　 : strBuff - JVDデータから取得したレース詳細レコード
    '            aRaceList - レース詳細リスト
    ' 返り値　 : なし
    ' 機能説明 : レース詳細レコードの情報を編集し、レース詳細リストに設定する
    '
    Public Sub RaceInfoMakeList(ByVal strBuff As String, ByRef aRaceList As ArrayList)
        Dim bBuff As Byte()
        Dim bSize As Long
        Dim strRace As String
        Dim strRaceInfo() As String
        Dim strFromKey As String
        Dim strNewKey As String

        Try
            bSize = 1272
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            ' リストの1件目にキー情報を追加する
            ' <競走識別情報>
            strRace = MidB2S(bBuff, 12, 4) & _
                    MidB2S(bBuff, 16, 4) & _
                    MidB2S(bBuff, 20, 2) & _
                    MidB2S(bBuff, 22, 2) & _
                    MidB2S(bBuff, 24, 2) & _
                    MidB2S(bBuff, 26, 2) & ","

            ' <競走識別情報>
            strRace = strRace & MidB2S(bBuff, 12, 4) & _
                    MidB2S(bBuff, 16, 4) & "," & _
                    objCodeConv.GetCodeName("2001", MidB2S(bBuff, 20, 2), "3") & "," & _
                    MidB2S(bBuff, 22, 2) & "," & _
                    MidB2S(bBuff, 24, 2) & "," & _
                    MidB2S(bBuff, 26, 2) & "," & _
                    objCodeConv.GetCodeName("2002", MidB2S(bBuff, 28, 1), "2") & "," & _
                    objCodeConv.GetCodeName("2008", MidB2S(bBuff, 622, 1), "1") & "," & _
                    MidB2S(bBuff, 698, 4) & "," & _
                    objCodeConv.GetCodeName("2009", MidB2S(bBuff, 706, 2), "2") & "," & _
                    MidB2S(bBuff, 874, 4) & "," & _
                    MidB2S(bBuff, 882, 2) & "," & _
                    Trim(MidB2S(bBuff, 33, 60)) & "," & _
                    Trim(MidB2S(bBuff, 93, 60)) & "," & _
                    Trim(MidB2S(bBuff, 153, 60)) 
            '出走頭数をやめて登録頭数を取得するよう修正
            'MidB2S(bBuff, 884, 2) & "," & _

            ' レース情報を判定し、更新があった場合はリストを更新する
            For i = 0 To RaceInfo.Count - 1
                ' 各レース情報から先頭25文字をキーとして取得
                strRaceInfo = RaceInfo(i)
                strFromKey = strRaceInfo(CommonConstant.Const0) & "," & _
                             strRaceInfo(CommonConstant.Const1)
                strNewKey = Strings.Left(strRace, CommonConstant.Const25)
                ' 現在レース情報と取得したレース情報のキーを比較する
                If strFromKey = strNewKey Then
                    ' レース情報を比較して異なる場合、現在レース情報を更新する
                    If Join(strRaceInfo, ",") <> strRace Then
                        RaceInfo.Item(i) = strRace.Split(","c)
                        Exit For
                    End If
                End If
            Next

            ' レース情報がすでに追加済みかチェック
            If Not aRaceList.Contains(strRace) Then
                aRaceList.Add(strRace)
            End If
            bBuff = Nothing
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    ' 機能　　 : レース詳細のCSVファイルを出力する
    ' 引き数　 : strFileName - ファイル名
    '            strFilePath - 出力パス
    '            aRaceList - レース詳細リスト
    ' 返り値　 : なし
    ' 機能説明 : 指定されたパスにレース詳細のCSVファイルを出力する
    '
    Public Sub RaceInfoOutput(ByVal strFileName As String, ByVal strFilePath As String, ByVal aRaceList As ArrayList)
        Try
            Dim i As Integer
            Dim strCsvPath As String    ' 保存先のCSVファイルパス
            Dim strCsvName As String    ' 保存先のCSVファイル名
            Dim workRaceList As New ArrayList

            ' 引数fileNameからCSVファイル名を作成
            'strCsvName = strFileName
            strCsvName = Mid(strFileName, 1, InStr(strFileName, ".") - 1) & CommonConstant.CSV
            strCsvPath = strFilePath & "\" & strCsvName

            If System.IO.File.Exists(strCsvPath) Then
                ' ファイルを削除
                Kill(strCsvPath)
            End If

            Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding(CommonConstant.EncType)
            Dim sr As New System.IO.StreamWriter(strCsvPath, True, enc)

            ' 取得した全レース詳細をCSVに出力する（20100912修正）
            For i = 0 To aRaceList.Count - 1
                ' <競走識別情報>
                sr.Write(aRaceList.Item(i).ToString)
                sr.Write(vbCrLf)
            Next

            '' レース詳細をArrayList<String>に変換する
            'For i = 0 To RaceInfo.Count - 1
            '    strWorkRace = Join(RaceInfo(i), ",")
            '    workRaceList.Add(strWorkRace)
            'Next

            'For i = 0 To workRaceList.Count - 1
            '    ' CSVファイルにレコードが存在するかチェックする
            '    ' 取得したレース詳細リストにデータが存在する場合はリストに追加しない
            '    If Not aRaceList.Contains(workRaceList.Item(i)) Then
            '        ' <競走識別情報>
            '        sr.Write(workRaceList.Item(i).ToString)
            '        sr.Write(vbCrLf)
            '    End If
            'Next

            'For i = 0 To aRaceList.Count - 1
            '    ' CSVファイルにレコードが存在するかチェックする
            '    ' レース詳細リストに取得データが存在する場合はリストに追加しない
            '    If Not workRaceList.Contains(aRaceList.Item(i)) Then
            '        ' <競走識別情報>
            '        sr.Write(aRaceList.Item(i).ToString)
            '        sr.Write(vbCrLf)
            '    End If
            'Next

            ''For i = 0 To aRaceList.Count - 1
            ''    ' CSVファイルにレコードが存在するかチェックする
            ''    ' レース詳細リストに取得データが存在する場合はリストに追加しない
            ''    If ((RaceInfo.Count = 0) Or _
            ''         (Not workRaceList.Contains(aRaceList.Item(i)))) Then
            ''        '' <競走識別情報>
            ''        sr.Write(aRaceList.Item(i).ToString)
            ''        sr.Write(vbCrLf)

            ''        ' ファイルオープンフラグをFalseにする
            ''        gFileOpenFlg = False
            ''    End If
            ''    '' <競走識別情報>
            ''    'sr.Write(aRaceList.Item(i).ToString)
            ''    'sr.Write(vbCrLf)
            ''Next

            ' 閉じる
            sr.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    ' 機能　　 : 馬毎レース情報のListを作成する
    ' 引き数　 : strBuff - JVDデータから取得した馬毎レース情報レコード
    '            aHorseList - 馬毎レース情報リスト
    ' 返り値　 : なし
    ' 機能説明 : 馬毎レース情報レコードの情報を編集し、馬毎レース情報リストに設定する
    '
    Public Sub HorseMakeList(ByVal strBuff As String, ByRef aHorseList As ArrayList)
        Dim bBuff As Byte()
        Dim bSize As Long
        Dim strHorse As String          ' レース情報
        Dim strHorseInfo() As String
        Dim strFromKey As String
        Dim strNewKey As String

        Try
            bSize = 2042
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)
            strHorse = ""

            ' 馬番が未決定の場合は処理終了
            If MidB2S(bBuff, 29, 2) = CommonConstant.No_HorseNo Then
                Exit Sub
            End If

            ' リストの1件目にキー情報を追加する
            ' <競走識別情報>
            strHorse = MidB2S(bBuff, 12, 4) & _
                    MidB2S(bBuff, 16, 4) & _
                    MidB2S(bBuff, 20, 2) & _
                    MidB2S(bBuff, 22, 2) & _
                    MidB2S(bBuff, 24, 2) & _
                    MidB2S(bBuff, 26, 2) & ","

            ' <馬毎レース情報>
            strHorse = strHorse & _
                    MidB2S(bBuff, 29, 2) & "," & _
                    Trim(MidB2S(bBuff, 41, 36)) & "," & _
                    Trim(MidB2S(bBuff, 307, 8))

            ' 馬毎レース情報を判定し、更新があった場合はリストを更新する
            For i = 0 To AllHorseInfo.Count - 1
                ' 各レース情報から先頭19文字をキーとして取得
                strHorseInfo = AllHorseInfo(i)
                strFromKey = strHorseInfo(CommonConstant.Const0) & "," & _
                             strHorseInfo(CommonConstant.Const1)
                strNewKey = Strings.Left(strHorse, CommonConstant.Const19)
                ' 現在レース情報と取得したレース情報のキーを比較する
                If strFromKey = strNewKey Then
                    ' レース情報を比較して異なる場合、現在レース情報を更新する
                    If Join(strHorseInfo, ",") <> strHorse Then
                        AllHorseInfo.Item(i) = strHorse.Split(","c)
                        Exit For
                    End If
                End If
            Next

            ' レース情報がすでに追加済みかチェック
            If Not aHorseList.Contains(strHorse) Then
                aHorseList.Add(strHorse)
            End If
            bBuff = Nothing
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    ' 機能　　 : オッズ（単複枠）のListを作成する
    ' 引き数　 : strBuff - JVDデータから取得したオッズ（単複枠）レコード
    '            aOdds1List - オッズ（単複枠）リスト
    ' 返り値　 : なし
    ' 機能説明 : オッズ（単複枠）レコードの情報を編集し、オッズ（単複枠）リストに設定する
    '
    Public Sub Odds1MakeList(ByVal strBuff As String, ByRef aOdds1List As ArrayList)
        Dim i As Integer
        Dim bBuff As Byte()
        Dim bSize As Long
        Dim bOddsInfo As Byte()         ' オッズ詳細情報
        Dim strOdds As String           ' オッズ情報

        Try
            bSize = 962
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)
            strOdds = ""

            ' リストの1件目にキー情報を追加する
            ' <競走識別情報>
            strOdds = MidB2S(bBuff, 12, 4) & _
                    MidB2S(bBuff, 16, 4) & _
                    MidB2S(bBuff, 20, 2) & _
                    MidB2S(bBuff, 22, 2) & _
                    MidB2S(bBuff, 24, 2) & _
                    MidB2S(bBuff, 26, 2) & ","

            ' <発表月日時分>
            bOddsInfo = MidB2B(bBuff, 28, 8)
            strOdds = strOdds & _
                        MidB2S(bOddsInfo, 1, 2) & _
                        MidB2S(bOddsInfo, 3, 2) & _
                        MidB2S(bOddsInfo, 5, 2) & _
                        MidB2S(bOddsInfo, 7, 2) & ","

            ' 馬番（28頭立て）分繰り返す
            For i = 0 To 27
                ' <単勝オッズ>
                ' 馬番、オッズ
                bOddsInfo = MidB2B(bBuff, 44 + (8 * i), 8)
                strOdds = strOdds & _
                            MidB2S(bOddsInfo, 1, 2) & "," & _
                            MidB2S(bOddsInfo, 3, 4) & ","

                ' <複勝オッズ>
                ' 最低オッズ
                bOddsInfo = MidB2B(bBuff, 268 + (12 * i), 12)
                strOdds = strOdds & _
                            MidB2S(bOddsInfo, 3, 4) & ","
            Next

            ' <登録頭数>
            strOdds = strOdds & MidB2S(bBuff, 36, 2)

            '出走頭数をやめて登録頭数を取得するよう修正
            '' <出走頭数>
            'strOdds = strOdds & MidB2S(bBuff, 38, 2)

            aOdds1List.Add(strOdds)
            bBuff = Nothing
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    ' 機能　　 : オッズ2（馬連）のListを作成する
    ' 引き数　 : strBuff - JVDデータから取得したオッズ2（馬連）レコード
    '            aOdds2List - オッズ2（馬連）リスト
    ' 返り値　 : なし
    ' 機能説明 : オッズ2（馬連）レコードの情報を編集し、オッズ2（馬連）リストに設定する
    '
    Public Sub Odds2MakeList(ByVal strBuff As String, ByRef aOdds2List As ArrayList)
        Dim i As Integer
        Dim bBuff As Byte()
        Dim bSize As Long
        Dim bOddsInfo As Byte()         ' オッズ詳細情報
        Dim strOdds As String           ' オッズ情報

        Try
            bSize = 2042
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)
            strOdds = ""

            ' リストの1件目にキー情報を追加する
            ' <競走識別情報>
            strOdds = MidB2S(bBuff, 12, 4) & _
                    MidB2S(bBuff, 16, 4) & _
                    MidB2S(bBuff, 20, 2) & _
                    MidB2S(bBuff, 22, 2) & _
                    MidB2S(bBuff, 24, 2) & _
                    MidB2S(bBuff, 26, 2) & ","

            ' <発表月日時分>
            bOddsInfo = MidB2B(bBuff, 28, 8)
            strOdds = strOdds & _
                        MidB2S(bOddsInfo, 1, 2) & _
                        MidB2S(bOddsInfo, 3, 2) & _
                        MidB2S(bOddsInfo, 5, 2) & _
                        MidB2S(bOddsInfo, 7, 2) & ","

            For i = 0 To 152
                ' <馬連オッズ>
                ' 組番、オッズ
                bOddsInfo = MidB2B(bBuff, 41 + (13 * i), 13)
                strOdds = strOdds & _
                            MidB2S(bOddsInfo, 1, 4) & "," & _
                            MidB2S(bOddsInfo, 5, 6) & ","
            Next
            aOdds2List.Add(strOdds)
            bBuff = Nothing
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    ' 機能　　 : オッズ4（馬単）のListを作成する
    ' 引き数　 : strBuff - JVDデータから取得したオッズ4（馬単）レコード
    '            aOdds4List - オッズ4（馬単）リスト
    ' 返り値　 : なし
    ' 機能説明 : オッズ4（馬単）レコードの情報を編集し、オッズ4（馬単）リストに設定する
    '
    Public Sub Odds4MakeList(ByVal strBuff As String, ByRef aOdds4List As ArrayList)
        Dim i As Integer
        Dim bBuff As Byte()
        Dim bSize As Long
        Dim bOddsInfo As Byte()         ' オッズ詳細情報
        Dim strOdds As String           ' オッズ情報

        Try
            bSize = 4031
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)
            strOdds = ""

            ' リストの1件目にキー情報を追加する
            ' <競走識別情報>
            strOdds = MidB2S(bBuff, 12, 4) & _
                    MidB2S(bBuff, 16, 4) & _
                    MidB2S(bBuff, 20, 2) & _
                    MidB2S(bBuff, 22, 2) & _
                    MidB2S(bBuff, 24, 2) & _
                    MidB2S(bBuff, 26, 2) & ","

            ' <発表月日時分>
            bOddsInfo = MidB2B(bBuff, 28, 8)
            strOdds = strOdds & _
                        MidB2S(bOddsInfo, 1, 2) & _
                        MidB2S(bOddsInfo, 3, 2) & _
                        MidB2S(bOddsInfo, 5, 2) & _
                        MidB2S(bOddsInfo, 7, 2) & ","

            For i = 0 To 305
                ' <馬連オッズ>
                bOddsInfo = MidB2B(bBuff, 41 + (13 * i), 13)
                strOdds = strOdds & _
                            MidB2S(bOddsInfo, 1, 4) & "," & _
                            MidB2S(bOddsInfo, 5, 6) & ","
            Next
            aOdds4List.Add(strOdds)
            bBuff = Nothing
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    ' 機能　　 : 馬毎レース情報のCSVファイルを出力する
    ' 引き数　 : strFileName - ファイル名
    '            strFilePath - 出力パス
    '            aHorseList - 馬毎レース情報のリスト
    ' 返り値　 : なし
    ' 機能説明 : 指定されたパスにCSVファイルを出力する
    '
    Public Sub SeInfoOutput(ByVal strFileName As String, ByVal strFilePath As String, ByVal aHorseList As ArrayList)
        Try
            Dim i As Integer
            Dim strCsvPath As String    ' 保存先のCSVファイルパス
            Dim strCsvName As String    ' 保存先のCSVファイル名
            Dim workRaceList As New ArrayList

            ' 引数fileNameからCSVファイル名を作成
            'strCsvName = strFileName
            strCsvName = Mid(strFileName, 1, InStr(strFileName, ".") - 1) & CommonConstant.CSV
            strCsvPath = strFilePath & "\" & strCsvName

            If System.IO.File.Exists(strCsvPath) Then
                ' ファイルを削除
                Kill(strCsvPath)
            End If

            Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding(CommonConstant.EncType)
            Dim sr As New System.IO.StreamWriter(strCsvPath, True, enc)

            ' 取得した全馬毎レース情報をCSVに出力する（20100912修正）
            For i = 0 To aHorseList.Count - 1
                ' <競走識別情報>
                sr.Write(aHorseList.Item(i).ToString)
                sr.Write(vbCrLf)
            Next

            '' レース詳細をArrayList<String>に変換する
            'For i = 0 To AllHorseInfo.Count - 1
            '    strWorkRace = Join(AllHorseInfo(i), ",")
            '    workRaceList.Add(strWorkRace)
            'Next

            'For i = 0 To workRaceList.Count - 1
            '    ' CSVファイルにレコードが存在するかチェックする
            '    ' 取得したレース詳細リストにデータが存在する場合はリストに追加しない
            '    If Not aHorseList.Contains(workRaceList.Item(i)) Then
            '        ' <馬毎レース情報>
            '        sr.Write(workRaceList.Item(i).ToString)
            '        sr.Write(vbCrLf)
            '    End If
            'Next

            'For i = 0 To aHorseList.Count - 1
            '    ' CSVファイルにレコードが存在するかチェックする
            '    ' レース詳細リストに取得データが存在する場合はリストに追加しない
            '    If Not workRaceList.Contains(aHorseList.Item(i)) Then
            '        ' <馬毎レース情報>
            '        sr.Write(aHorseList.Item(i).ToString)
            '        sr.Write(vbCrLf)
            '    End If
            'Next

            ''For i = 0 To aHorseList.Count - 1
            ''    ' CSVファイルにレコードが存在するかチェックする
            ''    ' 馬毎レース情報リストに取得データが存在する場合はリストに追加しない
            ''    If ((AllHorseInfo.Count = 0) Or _
            ''         (Not workRaceList.Contains(aHorseList.Item(i)))) Then
            ''        ' <馬毎レース情報>
            ''        sr.Write(aHorseList.Item(i).ToString)
            ''        sr.Write(","c)
            ''        sr.Write(vbCrLf)
            ''    End If
            ''    '' <馬毎レース情報>
            ''    'sr.Write(aHorseList.Item(i).ToString)
            ''    'sr.Write(","c)
            ''    'sr.Write(vbCrLf)
            ''Next

            ' 閉じる
            sr.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    ' 機能　　 : オッズ情報のCSVファイルを出力する
    ' 引き数　 : strFileName - ファイル名
    '            strFilePath - 出力パス
    '            aOddsList  - オッズ情報のリスト
    '            strDataKbn - データ区分（1：単複枠、2：馬連、3：馬単）
    ' 返り値　 : なし
    ' 機能説明 : 指定されたパスにCSVファイルを出力する
    '
    Public Sub OddsOutput(ByVal strFileName As String, ByVal strFilePath As String, _
                          ByVal aOddsList As ArrayList, ByVal strDataKbn As String)
        Try
            Dim i As Integer
            Dim strCsvPath As String    ' 保存先のCSVファイルパス
            Dim strCsvName As String    ' 保存先のCSVファイル名

            ' 引数fileNameからCSVファイル名を作成
            strCsvName = strFileName
            'strCsvName = Mid(strFileName, 1, InStr(strFileName, ".") - 1) & CommonConstant.CSV
            strCsvPath = strFilePath & "\" & strCsvName

            Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding(CommonConstant.EncType)
            Dim sr As New System.IO.StreamWriter(strCsvPath, True, enc)

            For i = 0 To aOddsList.Count - 1
                ' CSVファイルにレコードが存在するかチェックする
                ' オッズ情報1リストに取得データが存在する場合はリストに追加しない
                If (strDataKbn = CommonConstant.TanshouOddsKbn) And _
                    (TanFukuAllOddsInfo.Count = 0 Or _
                     Not TanFukuAllOddsInfo.Contains(aOddsList.Item(i))) Then
                    ' <オッズ情報>
                    sr.Write(aOddsList.Item(i).ToString)
                    sr.Write(","c)
                    sr.Write(vbCrLf)
                End If

                ' CSVファイルにレコードが存在するかチェックする
                ' オッズ情報2リストに取得データが存在する場合はリストに追加しない
                If (strDataKbn = CommonConstant.UmarenOddsKbn) And _
                    (UmarenAllOddsInfo.Count = 0 Or _
                     Not UmarenAllOddsInfo.Contains(aOddsList.Item(i))) Then
                    ' <オッズ情報>
                    sr.Write(aOddsList.Item(i).ToString)
                    sr.Write(","c)
                    sr.Write(vbCrLf)
                End If

                ' CSVファイルにレコードが存在するかチェックする
                ' オッズ情報4リストに取得データが存在する場合はリストに追加しない
                If (strDataKbn = CommonConstant.UmatanOddsKbn) And _
                    (UmatanAllOddsInfo.Count = 0 Or _
                     Not UmatanAllOddsInfo.Contains(aOddsList.Item(i))) Then
                    ' <オッズ情報>
                    sr.Write(aOddsList.Item(i).ToString)
                    sr.Write(","c)
                    sr.Write(vbCrLf)
                End If
            Next

            ' 閉じる
            sr.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
