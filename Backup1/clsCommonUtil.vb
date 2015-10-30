Imports System.IO
Public Class clsCommonUtil

    Public Function CheckFileUsing(ByVal vstrFilePath As String) As Boolean

        Dim f As New FileInfo(vstrFilePath)
        Dim strFileNameBK As String = f.Name

        If Not f.Exists Then
            'ファイルが存在しなければ、使用中ではない
            Return False
        End If

        Try
            'ファイル名を変更して、使用中かチェックする
            f.MoveTo(Path.Combine(f.DirectoryName, f.Name & ".BK"))
            'ファイル名を元に戻す
            f.MoveTo(Path.Combine(f.DirectoryName, strFileNameBK))

            'ファイル名の変更が成功したので、使用中ではない
            Return False
        Catch ex As Exception
            'ファイル名が変更できないので、使用中とする
            Return True
        End Try
    End Function


    ' 機能 　　：テンポラリフォルダへレース情報を格納する。
    Public Sub backupToTempFolder()
        Dim tempPath As String = System.IO.Path.GetTempPath()
        'テンポラリフォルダへ一時保存していたファイルを削除する
        Dim fs As String()
        Dim fName(9) As String
        Dim i As Integer
        Dim j As Integer

        '\マークの付加
        tempPath = IIf(Right(tempPath, 1) = "\", tempPath, tempPath & "\")

        fName(0) = CommonConstant.RA_FileName
        fName(1) = CommonConstant.Old_RA_FileNameDate & "*" & CommonConstant.CSV
        fName(2) = CommonConstant.SE_FileName
        fName(3) = CommonConstant.Old_SE_FileNameDate & "*" & CommonConstant.CSV
        fName(4) = CommonConstant.O1_FileName
        fName(5) = CommonConstant.Old_O1_FileNameDate & "*" & CommonConstant.CSV
        fName(6) = CommonConstant.O2_FileName
        fName(7) = CommonConstant.Old_O2_FileNameDate & "*" & CommonConstant.CSV
        fName(8) = CommonConstant.O4_FileName
        fName(9) = CommonConstant.Old_O4_FileNameDate & "*" & CommonConstant.CSV


        'ファイルを削除する
        For i = 0 To fName.Count - 1
            'ファイル名配列をクリア
            If i > 0 Then Array.Clear(fs, 0, fs.Length)
            fs = System.IO.Directory.GetFiles(gPrgFilesPath, fName(i))
            'ファイルを削除
            For j = 0 To fs.Count - 1
                Try
                    System.IO.File.Copy(fs(j), tempPath & IO.Path.GetFileName(fs(j)), True)
                Catch fex As IO.FileLoadException

                End Try
            Next
        Next
    End Sub

    ' 機能  　：テンポラリフォルダへ一時保存していたファイルを削除する
    Public Sub deleteToTempFolder()
        Dim tempPath As String = System.IO.Path.GetTempPath()
        'テンポラリフォルダへ一時保存していたファイルを削除する
        Dim fs As String()
        Dim fName(9) As String
        Dim i As Integer
        Dim j As Integer

        '\マークの付加
        tempPath = IIf(Right(tempPath, 1) = "\", tempPath, tempPath & "\")

        fName(0) = CommonConstant.RA_FileName
        fName(1) = CommonConstant.Old_RA_FileNameDate & "*" & CommonConstant.CSV
        fName(2) = CommonConstant.SE_FileName
        fName(3) = CommonConstant.Old_SE_FileNameDate & "*" & CommonConstant.CSV
        fName(4) = CommonConstant.O1_FileName
        fName(5) = CommonConstant.Old_O1_FileNameDate & "*" & CommonConstant.CSV
        fName(6) = CommonConstant.O2_FileName
        fName(7) = CommonConstant.Old_O2_FileNameDate & "*" & CommonConstant.CSV
        fName(8) = CommonConstant.O4_FileName
        fName(9) = CommonConstant.Old_O4_FileNameDate & "*" & CommonConstant.CSV


        'ファイルを削除する
        For i = 0 To fName.Count - 1
            'ファイル名配列をクリア
            If i > 0 Then Array.Clear(fs, 0, fs.Length)
            fs = System.IO.Directory.GetFiles(tempPath, fName(i))
            'ファイルを削除
            For j = 0 To fs.Count - 1
                Try
                    System.IO.File.Delete(fs(j))
                Catch fex As Exception

                End Try
            Next
        Next

    End Sub

    ' 機能    : テンポラリフォルダへ一時保存していたファイルを戻す
    Public Sub rollbackToTempFolder()
        Dim tempPath As String = System.IO.Path.GetTempPath()
        'テンポラリフォルダへ一時保存していたファイルを削除する
        Dim fs As String()
        Dim fName(9) As String
        Dim i As Integer
        Dim j As Integer
        Dim strmsg As String

        '\マークの付加
        tempPath = IIf(Right(tempPath, 1) = "\", tempPath, tempPath & "\")

        fName(0) = CommonConstant.RA_FileName
        fName(1) = CommonConstant.Old_RA_FileNameDate & "*" & CommonConstant.CSV
        fName(2) = CommonConstant.SE_FileName
        fName(3) = CommonConstant.Old_SE_FileNameDate & "*" & CommonConstant.CSV
        fName(4) = CommonConstant.O1_FileName
        fName(5) = CommonConstant.Old_O1_FileNameDate & "*" & CommonConstant.CSV
        fName(6) = CommonConstant.O2_FileName
        fName(7) = CommonConstant.Old_O2_FileNameDate & "*" & CommonConstant.CSV
        fName(8) = CommonConstant.O4_FileName
        fName(9) = CommonConstant.Old_O4_FileNameDate & "*" & CommonConstant.CSV


        'ファイルを削除する
        For i = 0 To fName.Count - 1
            'ファイル名配列をクリア
            If i > 0 Then Array.Clear(fs, 0, fs.Length)
            fs = System.IO.Directory.GetFiles(tempPath, IO.Path.GetFileName(fName(i)))
            'ファイルを削除
            For j = 0 To fs.Count - 1
                Try
                    System.IO.File.Copy(fs(j), gPrgFilesPath & IO.Path.GetFileName(fs(j)), True)
                Catch fex As Exception
                    strmsg = fex.Message
                End Try
            Next
        Next
    End Sub

    Public Function RaceFileOpen() As Boolean
        Dim RetCode As Boolean = True

        Dim OpenFile As String() = System.IO.Directory.GetFiles(gPrgFilesPath & "\", CommonConstant.RA_FileName)
        Application.DoEvents()
        If (OpenFile.Length = CommonConstant.Const0 Or OpenFile.Length > CommonConstant.Const1) Then
            RetCode = False
            Return RetCode
        Else
            'レースファイル名を取得する
            Dim RaceDataFile As String = OpenFile(CommonConstant.IndexPos0)
            setFileToList(RaceDataFile, RaceInfo)

        End If

        '騎手／馬名情報ファイル名を取得する
        Application.DoEvents()
        OpenFile = System.IO.Directory.GetFiles(gPrgFilesPath & "\", CommonConstant.SE_FileName)
        If (OpenFile.Length = CommonConstant.Const0 Or OpenFile.Length > CommonConstant.Const1) Then
            RetCode = False
            Return RetCode
        Else
            'レースファイル名を取得する
            Dim HorseDataFile As String = OpenFile(CommonConstant.IndexPos0)
            '騎手／馬名情報ファイルをリスト＜配列＞に展開する
            setFileToList(HorseDataFile, AllHorseInfo)
        End If

        Return RetCode

    End Function

    ' 機能　　：過去データファイル分割処理
    ' 引き数　：なし
    ' 返り値　：Boolean - 処理結果
    ' 機能説明：過去データファイルを、日付単位に分散する Old_XX_INFODATA.csv → Old_XX_INFODATA_YYYYMMDD.csv
    Public Function SplitOldDataFile(ByVal pBeforeOldFPath As String, ByVal pOldFName As String) As Boolean

        Dim RetCode As Boolean = True
        Dim OpenOldFile As String()
        Dim OpenOldDateFile As String()
        Dim OldYMD As String = ""

        Dim exceptionflg As Boolean = False


        Try
            '分割されていない過去データファイルが存在するときのみ処理を行う
            OpenOldFile = System.IO.Directory.GetFiles(gPrgFilesPath & "\", pOldFName & CommonConstant.CSV)
            If OpenOldFile.Length > 0 Then
                '分割前レース情報ファイルを変数に格納
                'setFileToList(OpenOldFile(CommonConstant.IndexPos0), OldInfo)

                '過去レース情報ファイルのサイズを取得する
                Dim fi As New System.IO.FileInfo(gPrgFilesPath & "\" & pBeforeOldFPath)
                Dim l As Long = fi.Length

                '過去レース情報ファイルサイズが0バイト以上なら処理する
                If l > 0 Then
                    'ファイルを開く
                    Dim sr As New IO.StreamReader(gPrgFilesPath & "\" & pBeforeOldFPath, System.Text.Encoding.GetEncoding("Shift-JIS"))

                    '内容を一行ずつ読み込む
                    While sr.Peek() > -1
                        Dim strLine As String = sr.ReadLine()
                        '行の先頭8文字を取得(日付)
                        OldYMD = strLine.Substring(CommonConstant.Const0, CommonConstant.Const8)


                        '日付ファイルが存在するかチェック
                        OpenOldDateFile = System.IO.Directory.GetFiles(gPrgFilesPath & "\", pOldFName & "_" & OldYMD & CommonConstant.CSV)
                        If Not (OpenOldDateFile.Length > CommonConstant.Const0) Then
                            'なければ新規作成
                            Dim sw As New IO.StreamWriter(gPrgFilesPath & "\" & pOldFName & "_" & OldYMD & CommonConstant.CSV, True, _
                                                          System.Text.Encoding.GetEncoding("Shift-JIS"))

                            sw.Close()
                            sw = Nothing
                            'System.IO.File.Create(gPrgFilesPath & "\" & pOldFName & "_" & OldYMD & CommonConstant.CSV)
                            OpenOldDateFile = System.IO.Directory.GetFiles(gPrgFilesPath & "\", pOldFName & "_" & OldYMD & CommonConstant.CSV)
                        End If

                        FileWriter(OpenOldDateFile(CommonConstant.IndexPos0), strLine)

                    End While
                    '閉じる
                    sr.Close()
                    sr = Nothing

                    '分割前のファイルを削除する
                    'System.IO.File.Delete(OpenOldFile(CommonConstant.IndexPos0))
                    System.IO.File.Move(OpenOldFile(CommonConstant.IndexPos0), gPrgFilesPath & "\" & pOldFName & ".old")

                End If

                '                If OldInfo.Count > 0 Then
                '                    '分割前レース情報を１行ずつ読む
                '                    For i = 0 To OldInfo.Count - 1
                '                        OldLine = OldInfo(i)
                '                        '行の先頭8文字を取得(日付)
                '                        OldYMD = OldLine(CommonConstant.IndexPos0).Substring(CommonConstant.Const0, CommonConstant.Const8)

                '                        '日付ファイルが存在するかチェック
                '                        OpenOldDateFile = System.IO.Directory.GetFiles(gPrgFilesPath & "\", pOldFName & "_" & OldYMD & CommonConstant.CSV)
                '                        If Not (OpenOldDateFile.Length > CommonConstant.Const0) Then
                '                            'なければ新規作成
                '                            Dim sw As New IO.StreamWriter(gPrgFilesPath & "\" & pOldFName & "_" & OldYMD & CommonConstant.CSV, True, _
                '                                                          System.Text.Encoding.GetEncoding("Shift-JIS"))

                '                            sw.Close()
                '                            sw = Nothing
                '                            'System.IO.File.Create(gPrgFilesPath & "\" & pOldFName & "_" & OldYMD & CommonConstant.CSV)
                '                            OpenOldDateFile = System.IO.Directory.GetFiles(gPrgFilesPath & "\", pOldFName & "_" & OldYMD & CommonConstant.CSV)
                '                        End If
                '#If False Then
                '                        'ファイルを変数に格納する
                '                        Dim sr As New IO.StreamReader(OpenOldDateFile(CommonConstant.IndexPos0), _
                '                                                      System.Text.Encoding.GetEncoding("Shift-JIS"))
                '                        'ファイル全体を読む

                '                        OldDateData = sr.ReadToEnd()
                '                        sr.Close()
                '                        sr = Nothing

                '                        blnEqualsFlg = False
                '                        '同じレコードが無いか確認する
                '                        If InStr(OldDateData, Join(OldLine, ",")) <> 0 Then
                '                            blnEqualsFlg = True
                '                        End If
                '                        '同じレコードが無い場合、追加する
                '                        If Not blnEqualsFlg Then
                '                            FileWriter(OpenOldDateFile(CommonConstant.IndexPos0), Join(OldLine, ","))
                '                        End If
                '#End If
                '                        FileWriter(OpenOldDateFile(CommonConstant.IndexPos0), Join(OldLine, ","))

                '                    Next
                '                End If
                '                '処理が終わったら、ファイルを削除する
                '                System.IO.File.Delete(OpenOldFile(CommonConstant.IndexPos0))
            End If

        Catch ex As Exception
            exceptionflg = True
        End Try

    End Function

    Public Function UpdateRaceInfo(ByVal newDate As String, ByVal NewFname As String, _
                                    ByVal olddatalist As ArrayList, ByVal oldFnameDate As String) As Boolean
        Dim retList As New ArrayList
        Dim openOldDateFile As String()
        Dim newFPath As String = gPrgFilesPath & "\" & NewFname


        Try


            '(無い訳は無いと思うが)ファイルが存在するときのみ処理を行う
            If IO.File.Exists(newFPath) Then
                '直近情報更新後用テンポラリファイルパス
                Dim tempNewFilepath As String = gPrgFilesPath & "\" & "temp" & CommonConstant.CSV
                If IO.File.Exists(tempNewFilepath) Then
                    IO.File.Delete(tempNewFilepath)
                End If

                Dim swt As New IO.StreamWriter(tempNewFilepath, True, System.Text.Encoding.GetEncoding("Shift-JIS"))
                swt.Close()
                swt = Nothing

                'レース情報ファイルのサイズを取得する
                Dim fi As New System.IO.FileInfo(newFPath)
                Dim l As Long = fi.Length

                'レース情報ファイルサイズが0バイト以上なら処理する
                If l > 0 Then
                    'ファイルを開く
                    Dim sr As New IO.StreamReader(newFPath, System.Text.Encoding.GetEncoding("Shift-JIS"))

                    '内容を一行ずつ読み込む
                    While sr.Peek() > -1
                        Dim strLine As String = sr.ReadLine()
                        '行の先頭8文字を取得(日付)
                        Dim readYMD = strLine.Substring(CommonConstant.Const0, CommonConstant.Const8)


                        '日付ファイルが存在するかチェック
                        openOldDateFile = System.IO.Directory.GetFiles(gPrgFilesPath & "\", oldFnameDate & "_" & readYMD & CommonConstant.CSV)
                        If Not (openOldDateFile.Length > CommonConstant.Const0) Then
                            'なければ新規作成
                            Dim sw As New IO.StreamWriter(gPrgFilesPath & "\" & oldFnameDate & "_" & readYMD & CommonConstant.CSV, True, _
                                                          System.Text.Encoding.GetEncoding("Shift-JIS"))

                            sw.Close()
                            sw = Nothing
                            openOldDateFile = System.IO.Directory.GetFiles(gPrgFilesPath & "\", oldFnameDate & "_" & readYMD & CommonConstant.CSV)
                        End If

                        FileWriter(openOldDateFile(CommonConstant.IndexPos0), strLine)
                    End While
                    '閉じる
                    sr.Close()
                    sr = Nothing

                    IO.File.Delete(newFPath)
                    FileSystem.Rename(tempNewFilepath, newFPath)

                End If
            End If


        Catch ex As Exception
            Return False
        End Try

    End Function


    Private Function FileWriter(ByVal pfilepath As String, ByVal pData As String) As Boolean
        Dim sw As New IO.StreamWriter(pfilepath, True, System.Text.Encoding.GetEncoding("Shift-JIS"))

        sw.WriteLine(pData)

        sw.Close()

    End Function

    ' 機能　　 : レース情報ファイル読み込み処理
    ' 引き数　 : blnOpenFlg - ファイル切り替えフラグ（True：直近、False：過去）
    ' 返り値　 : Boolean - 処理結果
    ' 機能説明 : レース詳細、馬毎レース情報のCSVファイルを読み込み、リストに展開する
    '
    '2010/12/17 d-kobayashi  update start
    '過去データ分割後は日付のコンボを選択するたびにデータを取得させる
    'Public Function RaceFileOpen(ByVal blnOpenFlg As Boolean) As Boolean
    Public Function RaceFileOpen(ByVal blnOpenFlg As Boolean, Optional ByVal pFileDate As String = "0") As Boolean
        '2010/12/17 update end
        Dim RetCode As Boolean = True
        Dim OpenFile As String()

        '
        If pFileDate = "0" Then
            ' ファイル切り替えフラグにより読み込むCSVファイルを切り替える
            If blnOpenFlg Then
                OpenFile = System.IO.Directory.GetFiles(gPrgFilesPath & "\", CommonConstant.RA_FileName)
                If (OpenFile.Length > CommonConstant.Const0) Then
                    'レースファイル名を取得する
                    Dim RaceDataFile As String = OpenFile(CommonConstant.IndexPos0)
                    setFileToList(RaceDataFile, RaceInfo)

                End If

                '騎手／馬名情報ファイル名を取得する
                OpenFile = System.IO.Directory.GetFiles(gPrgFilesPath & "\", CommonConstant.SE_FileName)
                If (OpenFile.Length > CommonConstant.Const0) Then
                    'レースファイル名を取得する
                    Dim HorseDataFile As String = OpenFile(CommonConstant.IndexPos0)
                    '騎手／馬名情報ファイルをリスト＜配列＞に展開する
                    setFileToList(HorseDataFile, AllHorseInfo)
                End If
            Else
                OpenFile = System.IO.Directory.GetFiles(gPrgFilesPath & "\", CommonConstant.Old_RA_FileNameDate & "_????????.csv")
                If (OpenFile.Length > CommonConstant.Const0) Then
                    'レースファイル名を取得する
                    Dim RaceDataFile As String = OpenFile(CommonConstant.IndexPos0)
                    '2010/12/16 d-kobayashi update start
                    'setFileToList(RaceDataFile, RaceInfo)
                    setOldFileToList(CommonConstant.Old_RA_FileNameDate, RaceInfo)
                    '2010/12/16 d-kobayashi update end
                Else
                    RaceInfo.Clear()
                End If

                '騎手／馬名情報ファイル名を取得する
                OpenFile = System.IO.Directory.GetFiles(gPrgFilesPath & "\", CommonConstant.Old_SE_FileNameDate & "_????????.csv")
                If (OpenFile.Length > CommonConstant.Const0) Then
                    'レースファイル名を取得する
                    Dim HorseDataFile As String = OpenFile(CommonConstant.IndexPos0)
                    '騎手／馬名情報ファイルをリスト＜配列＞に展開する
                    '2010/12/16 d-kobayashi update start
                    'setFileToList(HorseDataFile, AllHorseInfo)
                    setOldFileToList(CommonConstant.Old_SE_FileNameDate, AllHorseInfo)
                    '2010/12/16 d-kobayashi update end
                Else
                    AllHorseInfo.Clear()
                End If
            End If
        Else
            OpenFile = System.IO.Directory.GetFiles(gPrgFilesPath & "\", CommonConstant.Old_RA_FileNameDate & "_" & pFileDate & CommonConstant.CSV)
            If (OpenFile.Length > CommonConstant.Const0) Then
                'レースファイル名を取得する
                Dim RaceDataFile As String = OpenFile(CommonConstant.IndexPos0)
                '騎手／馬名情報ファイルをリスト＜配列＞に展開する
                setFileToList(RaceDataFile, RaceInfo)
            Else
                RaceInfo.Clear()

            End If

            OpenFile = System.IO.Directory.GetFiles(gPrgFilesPath & "\", CommonConstant.Old_SE_FileNameDate & "_" & pFileDate & CommonConstant.CSV)
            If (OpenFile.Length > CommonConstant.Const0) Then
                'レースファイル名を取得する
                Dim horsedatafile As String = OpenFile(CommonConstant.IndexPos0)
                '騎手／馬名情報ファイルをリスト＜配列＞に展開する。
                setFileToList(horsedatafile, AllHorseInfo)
            Else
                AllHorseInfo.Clear()
            End If
        End If

        Return RetCode

    End Function


    Public Function OddsFileOpen() As Boolean

        Dim RetCode As Boolean = True

        'MsgBox("OddsFileOpen() IN")
        Dim OpenFile As String()
        Dim OddsDataFile As String


        '単勝／複勝オッズファイル名を取得する
        Application.DoEvents()
        OpenFile = System.IO.Directory.GetFiles(gPrgFilesPath & "\", CommonConstant.O1_FileName)
        If (OpenFile.Length = CommonConstant.Const0 Or OpenFile.Length > CommonConstant.Const1) Then
            'ファイル読み込みフラグをオフに設定
            RetCode = False
            Return RetCode
        Else
            'オッズファイル名を取得する
            OddsDataFile = OpenFile(CommonConstant.IndexPos0)
            '単勝／複勝オッズファイルをリスト＜配列＞に展開する
            setFileToList(OddsDataFile, TanFukuAllOddsInfo)
        End If

        '馬単オッズファイル名を取得する
        Application.DoEvents()
        OpenFile = System.IO.Directory.GetFiles(gPrgFilesPath & "\", CommonConstant.O4_FileName)
        If (OpenFile.Length = CommonConstant.Const0 Or OpenFile.Length > CommonConstant.Const1) Then
            'ファイル読み込みフラグをオフに設定
            RetCode = False
            Return RetCode
        Else
            'オッズファイル名を取得する
            OddsDataFile = OpenFile(CommonConstant.IndexPos0)
            '馬単オッズファイルをリスト＜配列＞に展開する
            setFileToList(OddsDataFile, UmatanAllOddsInfo)
        End If


        '馬連オッズファイル名を取得する
        Application.DoEvents()
        OpenFile = System.IO.Directory.GetFiles(gPrgFilesPath & "\", CommonConstant.O2_FileName)
        If (OpenFile.Length = CommonConstant.Const0 Or OpenFile.Length > CommonConstant.Const1) Then
            'ファイル読み込みフラグをオフに設定
            RetCode = False
            Return RetCode
        Else

            'オッズファイル名を取得する
            OddsDataFile = OpenFile(CommonConstant.IndexPos0)
            '馬連オッズファイルをリスト＜配列＞に展開する
            setFileToList(OddsDataFile, UmarenAllOddsInfo)
        End If

        Return RetCode

    End Function

    ' 機能　　 : オッズ情報ファイル読み込み処理
    ' 引き数　 : blnOpenFlg - ファイル切り替えフラグ（True：直近、False：過去）
    ' 返り値　 : Boolean - 処理結果
    ' 機能説明 : オッズ情報のCSVファイルを読み込み、リストに展開する
    '
    '2010/12/17 d-kobayashi update start
    'Public Sub OddsFileOpen(ByVal blnOpenFile As Boolean)
    'ファイルの日付を追加
    Public Sub OddsFileOpen(ByVal blnOpenFile As Boolean, Optional ByVal pFileDate As String = "0")
        '2010/12/17 d-kobayashi update end

        Dim OpenFile As String()
        Dim OddsDataFile As String

        If pFileDate = "0" Then
            ' ファイル切り替えフラグにより読み込むCSVファイルを切り替える
            If blnOpenFile Then
                '単勝／複勝オッズファイル名を取得する
                OpenFile = System.IO.Directory.GetFiles(gPrgFilesPath & "\", CommonConstant.O1_FileName)
                If (OpenFile.Length > CommonConstant.Const0) Then
                    'オッズファイル名を取得する
                    OddsDataFile = OpenFile(CommonConstant.IndexPos0)
                    '単勝／複勝オッズファイルをリスト＜配列＞に展開する
                    setFileToList(OddsDataFile, TanFukuAllOddsInfo)
                End If

                '馬単オッズファイル名を取得する
                OpenFile = System.IO.Directory.GetFiles(gPrgFilesPath & "\", CommonConstant.O4_FileName)
                If (OpenFile.Length > CommonConstant.Const0) Then
                    'オッズファイル名を取得する
                    OddsDataFile = OpenFile(CommonConstant.IndexPos0)
                    '馬単オッズファイルをリスト＜配列＞に展開する
                    setFileToList(OddsDataFile, UmatanAllOddsInfo)
                End If


                '馬連オッズファイル名を取得する
                OpenFile = System.IO.Directory.GetFiles(gPrgFilesPath & "\", CommonConstant.O2_FileName)
                If (OpenFile.Length > CommonConstant.Const0) Then
                    'オッズファイル名を取得する
                    OddsDataFile = OpenFile(CommonConstant.IndexPos0)
                    '馬連オッズファイルをリスト＜配列＞に展開する
                    setFileToList(OddsDataFile, UmarenAllOddsInfo)
                End If
            Else
                '単勝／複勝オッズファイル名を取得する
                OpenFile = System.IO.Directory.GetFiles(gPrgFilesPath & "\", CommonConstant.Old_O1_FileNameDate & "_????????.csv")
                If (OpenFile.Length > CommonConstant.Const0) Then
                    'オッズファイル名を取得する
                    OddsDataFile = OpenFile(CommonConstant.IndexPos0)
                    '単勝／複勝オッズファイルをリスト＜配列＞に展開する
                    '2010/12/17 d-kobayashi update start
                    'setFileToList(OddsDataFile, TanFukuAllOddsInfo)
                    setOldFileToList(CommonConstant.Old_O1_FileNameDate, TanFukuAllOddsInfo)
                    '2010/12/17 d-kobayashi update end
                Else
                    TanFukuAllOddsInfo.Clear()
                End If

                '馬単オッズファイル名を取得する
                OpenFile = System.IO.Directory.GetFiles(gPrgFilesPath & "\", CommonConstant.Old_O4_FileNameDate & "_????????.csv")
                If (OpenFile.Length > CommonConstant.Const0) Then
                    'オッズファイル名を取得する
                    OddsDataFile = OpenFile(CommonConstant.IndexPos0)
                    '馬単オッズファイルをリスト＜配列＞に展開する
                    '2010/12/17 d-kobayashi update start
                    'setFileToList(OddsDataFile, UmatanAllOddsInfo)
                    setOldFileToList(CommonConstant.Old_O4_FileNameDate, UmatanAllOddsInfo)
                    '2010/12/17 d-kobayashi update end
                Else
                    UmatanAllOddsInfo.Clear()
                End If

                '馬連オッズファイル名を取得する
                OpenFile = System.IO.Directory.GetFiles(gPrgFilesPath & "\", CommonConstant.Old_O2_FileNameDate & "_????????.csv")
                If (OpenFile.Length > CommonConstant.Const0) Then
                    'オッズファイル名を取得する
                    OddsDataFile = OpenFile(CommonConstant.IndexPos0)
                    '馬連オッズファイルをリスト＜配列＞に展開する
                    '2010/12/17 d-kobayashi update start
                    'setFileToList(OddsDataFile, UmarenAllOddsInfo)
                    setOldFileToList(CommonConstant.Old_O2_FileNameDate, UmarenAllOddsInfo)
                    '2010/12/17 d-kobayashi updat end
                Else
                    UmarenAllOddsInfo.Clear()
                End If
            End If
        Else
            '単勝／複勝オッズファイル名を取得する
            OpenFile = System.IO.Directory.GetFiles(gPrgFilesPath & "\", CommonConstant.Old_O1_FileNameDate & "_" & pFileDate & CommonConstant.CSV)
            If (OpenFile.Length > CommonConstant.Const0) Then
                'オッズファイル名を取得する
                OddsDataFile = OpenFile(CommonConstant.IndexPos0)
                '単勝／複勝オッズファイルをリスト＜配列＞に展開する
                setFileToList(OddsDataFile, TanFukuAllOddsInfo)
            End If

            '馬単オッズファイル名を取得する
            OpenFile = System.IO.Directory.GetFiles(gPrgFilesPath & "\", CommonConstant.Old_O4_FileNameDate & "_" & pFileDate & CommonConstant.CSV)
            If (OpenFile.Length > CommonConstant.Const0) Then
                'オッズファイル名を取得する
                OddsDataFile = OpenFile(CommonConstant.IndexPos0)
                '馬単オッズファイルをリスト＜配列＞に展開する
                setFileToList(OddsDataFile, UmatanAllOddsInfo)
            End If


            '馬連オッズファイル名を取得する
            OpenFile = System.IO.Directory.GetFiles(gPrgFilesPath & "\", CommonConstant.Old_O2_FileNameDate & "_" & pFileDate & CommonConstant.CSV)
            If (OpenFile.Length > CommonConstant.Const0) Then
                'オッズファイル名を取得する
                OddsDataFile = OpenFile(CommonConstant.IndexPos0)
                '馬連オッズファイルをリスト＜配列＞に展開する
                setFileToList(OddsDataFile, UmarenAllOddsInfo)
            End If
        End If

    End Sub

    Public Function GetOddsPosition(ByVal OddsKbn As String) As String
        '穴馬予想基準のオッズを判定し、オッズ情報が格納された配列位置を返却する

        '戻り値格納変数（配列）
        Dim retOddsPos As String

        If (OddsKbn = CommonConstant.TanshouOddsKbn) Then
            'オッズ区分が"1"（単勝オッズ）の場合
            retOddsPos = CommonConstant.OdTanshouOddsPos

        ElseIf (OddsKbn = CommonConstant.FukushouOddsKbn) Then
            'オッズ区分が"2"（複勝オッズ）の場合
            retOddsPos = CommonConstant.OdFukushouOddsPos

        ElseIf (OddsKbn = CommonConstant.UmatanOddsKbn) Then
            'オッズ区分が"3"（馬単オッズ）の場合
            retOddsPos = CommonConstant.OdUmatanOddsPos

        Else
            'オッズ区分が"4"（馬連オッズ）の場合
            retOddsPos = CommonConstant.OdUmarenOddsPos

        End If

        '値の返却
        Return retOddsPos

    End Function

    Public Function GetSortKeyPos(ByVal OddsKbn As String) As String()
        '穴馬予想基準のオッズを判定し、ソートキーとなるオッズ情報格納位置を返却する

        '戻り値格納変数（配列）
        Dim retSortOddsDataPos As String() = {}

        If (OddsKbn = CommonConstant.TanshouOddsKbn) Then
            'オッズ区分が"1"（単勝オッズ）の場合
            Dim WorkPos As String() = {CommonConstant.OlTnashouOddsPos, _
                                       CommonConstant.OlFukushouOddsPos, _
                                       CommonConstant.OlUmatanOddsPos}
            retSortOddsDataPos = WorkPos


        ElseIf (OddsKbn = CommonConstant.FukushouOddsKbn) Then
            'オッズ区分が"2"（複勝オッズ）の場合
            Dim WorkPos As String() = {CommonConstant.OlFukushouOddsPos, _
                                       CommonConstant.OlTnashouOddsPos, _
                                       CommonConstant.OlUmatanOddsPos}
            retSortOddsDataPos = WorkPos

        ElseIf (OddsKbn = CommonConstant.UmatanOddsKbn) Then
            'オッズ区分が"3"（馬単オッズ）の場合
            Dim WorkPos As String() = {CommonConstant.OlUmatanOddsPos, _
                                       CommonConstant.OlTnashouOddsPos, _
                                       CommonConstant.OlFukushouOddsPos}
            retSortOddsDataPos = WorkPos
            '2011/11/08 d-kobayashi update start
            'Else
        ElseIf (OddsKbn = CommonConstant.UmarenOddsKbn) Then
            '2011/11/08 d-kobayashi update end
            'オッズ区分が"4"（馬連オッズ）の場合
            Dim WorkPos As String() = {CommonConstant.OlUmarenOddsPos, _
                                       CommonConstant.OlTnashouOddsPos, _
                                       CommonConstant.OlFukushouOddsPos, _
                                       CommonConstant.OlUmatanOddsPos}
            retSortOddsDataPos = WorkPos
            '2011/11/08 add start
            '新しいソート順を追加
        ElseIf (OddsKbn = CommonConstant.UmarenOddsKbn2) Then
            'オッズ区分が"5"（馬連オッズ2）の場合
            Dim WorkPos As String() = {CommonConstant.OlUmarenOddsPos, _
                                       CommonConstant.OlFukushouOddsPos, _
                                       CommonConstant.OlTnashouOddsPos, _
                                       CommonConstant.OlUmatanOddsPos}
            retSortOddsDataPos = WorkPos
        ElseIf (OddsKbn = CommonConstant.UmarenOddsKbn3) Then
            'オッズ区分が"6"（馬連オッズ3）の場合
            Dim WorkPos As String() = {CommonConstant.OlUmarenOddsPos, _
                                       CommonConstant.OlUmatanOddsPos, _
                                       CommonConstant.OlTnashouOddsPos, _
                                       CommonConstant.OlFukushouOddsPos}
            retSortOddsDataPos = WorkPos
            '2011/11/08 d-kobayashi add end

        End If



        Return retSortOddsDataPos


    End Function

    '過去データファイル名をリストへ(ArrayList<String()>へ展開する
    Public Sub setOldFileToList(ByVal InputFile As String, ByRef OutputList As ArrayList)

        Dim fs As String() = System.IO.Directory.GetFiles(gPrgFilesPath, InputFile & "_????????.csv")
        Dim i As Integer
        Dim strYMD As String
        'ArrayListに追加する
        OutputList.Clear()
        'OutputList.AddRange(fs)
        If fs.Count > 0 Then
            For i = 0 To fs.Count - 1
                strYMD = StrConv(fs(i), VbStrConv.Uppercase).Replace(InputFile & "_", "")
                strYMD = strYMD.Replace(".CSV", "")
                strYMD = Right(strYMD, 8)
                'strYMD = Left(strYMD, 4) & "/" & Mid(strYMD, 5, 2) & "/" & Right(strYMD, 2)
                'strYMD = Format(DateValue(strYMD), "yyyy年MM月dd日(ddd)")
                Dim strList As String() = {strYMD, fs(i)}
                OutputList.Add(strList)

            Next
        End If


    End Sub

    'オッズ情報及びレース情報をリスト(ArrayList<String()>)へ展開する
    Public Sub setFileToList(ByVal InputFile As String, ByRef OutputList As ArrayList)
        'Dim f As Integer
        'Dim readData As String
        'Dim FirstRec As Boolean = True

        'リストの初期化
        OutputList.Clear()

        'f = FreeFile()

        'FileOpen(f, InputFile, OpenMode.Input)

        'Do Until EOF(f)
        '    readData = LineInput(f)
        '    Dim readDataH As String() = readData.Split(","c)
        '    OutputList.Add(readDataH)
        'Loop

        'FileClose(f)

        ' StreamReader の新しいインスタンスを生成する 
        Dim cReader As New System.IO.StreamReader(InputFile, System.Text.Encoding.Default)

        ' 読み込んだ結果をすべて格納するための変数を宣言する 
        Dim stResult As String = String.Empty

        ' 読み込みできる文字がなくなるまで繰り返す 
        While (cReader.Peek() >= CommonConstant.Const0)
            ' ファイルを 1 行ずつ読み込む 
            Dim stBuffer As String() = cReader.ReadLine().Split(","c)
            ' 読み込んだものを追加で格納する 
            OutputList.Add(stBuffer)
            Application.DoEvents()
        End While

        ' cReader を閉じる (正しくは オブジェクトの破棄を保証する を参照) 
        cReader.Close()

    End Sub

    'レース日毎の情報をメインフォームのコンボボックスにセットする 
    '2010/12/17 d-kobayashi update start
    '引数を変更
    'Public Sub setRaceDate(ByRef cboDate As System.Windows.Forms.ComboBox)
    Public Sub setRaceDate(ByRef cboDate As System.Windows.Forms.ComboBox, Optional ByVal cboSelType As Integer = 0)
        '2010/12/17 d-kobayashi update end

        Dim BeforeDate As String = ""
        Dim ReadDate As String = ""
        Dim FirstRec As Boolean = True
        Dim cboDateList As New ArrayList
        Dim FirstDispIndex As Integer = CommonConstant.Const0
        Dim IndexSetFlg As Boolean = False


        ' 日付と時刻を格納するための変数を宣言する 
        Dim ToDay As String = (Date.Now).ToString("yyyyMMdd")
        Dim AddDateList As New ArrayList

        For i = 0 To RaceInfo.Count - 1

            Dim readDataH As String() = RaceInfo(i)

            If cboSelType = 1 Then
                ReadDate = readDataH(CommonConstant.IndexPos0)
            Else
                ReadDate = readDataH(CommonConstant.IndexPos1)
            End If

            'If (BeforeDate <> ReadDate) Then 
            If Not AddDateList.Contains(ReadDate) Then
                'cboDateList.Add(ReadDate.Clone) 
                ' 日付 
                Dim dateP As String = Format(DateValue(Format(CInt(ReadDate), "0000/00/00")), "yyyy/MM/dd(ddd)")

                ' 曜日 
                'dateP = dateP & "(" & readDataH(CommonConstant.IndexPos6) & ")"
                cboDateList.Add(dateP)

                AddDateList.Add(ReadDate.Clone)
            End If
            'BeforeDate = ReadDate.Clone 

        Next

        '昇順ソート 
        cboDateList.Sort()
        '降順に並べ替え 
        cboDateList.Reverse()

        'コンボボックスに表示 
        For i = 0 To cboDateList.Count - 1
            cboDate.Items.Add(cboDateList(i))
        Next

        If cboDate.Items.Count > 0 Then
            cboDate.SelectedIndex = CommonConstant.Const0
        End If
        'For i = 0 To cboDateList.Count - 1 
        '    Dim AddCboDate As String = cboDateList(i) 

        '    If (ToDay <= AddCboDate) Then 
        '        FirstDispIndex = i 
        '        Exit For 
        '    End If 
        'Next 

        'If (cboDate.Items.Count > CommonConstant.Const0) Then 
        '    If IndexSetFlg Then 
        '        'インデックスをセットしている（当日以降のレース日がある）場合 
        '        cboDate.SelectedIndex = FirstDispIndex 
        '    Else 
        '        'インデックスをセットしている（過去のレース日がある）場合 
        '        cboDate.SelectedIndex = cboDateList.Count - 1 

        '    End If 

        'End If 



    End Sub


    '-----------------------------------------------------------------
    '全オッズ情報から該当レースのオッズ情報を取得し、リストへ展開する
    '<INPUT>
    ' TanFukuAllOddsInfo：全単勝／複勝オッズリスト
    ' UmatanAllOddsInfo：全馬単オッズリスト
    ' UmarenAllOddsInfo：全馬連オッズリスト
    ' SelTime：選択オッズ発表時間
    ' CheckKey：該当レースキー情報
    '
    '<OUTPUT>
    ' OddsList：オッズリスト（要素フォーマット：馬番号・単勝オッズ・複勝オッズ・馬単オッズ）
    ' MzList：マジックゾーンリスト（要素フォーマット：馬番号・馬連オッズ・単勝オッズ・複勝オッズ・馬単オッズ）


    Public Sub GetOddsData(ByVal TanFukuAllOddsInfo As ArrayList, _
                           ByVal UmatanAllOddsInfo As ArrayList, _
                           ByVal UmarenAllOddsInfo As ArrayList, _
                           ByVal CheckKey As String, _
                           ByVal SelTime As String, _
                           ByRef TanshouMaxNinkiUmaNo As String, _
                           ByRef OddsList As ArrayList, _
                           ByRef MzList As ArrayList, _
                           ByVal oddsKbn As Integer)

        Dim umaNoIndex As Integer = -1
        Dim tanOddsIndex As Integer = 0
        Dim fukuOddsIndex As Integer = 1
        Dim addint As Integer = 3
        Dim BeforeTime As String = ""
        Dim selectCounter As Integer = 0
        Dim kumiawaseIndex As Integer = 2
        Dim umatanOddsIndex As Integer = 3
        Dim umatanWorkList As ArrayList = New ArrayList
        Dim umarenWorkList As ArrayList = New ArrayList

        '人気馬初期化
        TanshouMaxNinkiUmaNo = String.Empty

        '単勝／複勝オッズの分解＆再設定を実施
        For i = 0 To TanFukuAllOddsInfo.Count - 1
            Dim OddsData As String() = TanFukuAllOddsInfo(i)
            'キー情報を取得
            Dim OddsKeyData As String = OddsData(CommonConstant.AllOdKeyData)

            '選択されたキー情報と一致するかを判定
            If (CheckKey = OddsKeyData) Then

                Dim OddsTime As String = OddsData(CommonConstant.AllOdUpTime)

                'オッズデータの発表時間とリストボックスで選択された発表時間を比較
                If (SelTime = OddsTime) Then

                    addint = 3
                    '一致した場合
                    For k = 1 To 18
                        '馬番号がなくなったらリスト設定を終了
                        If (Replace(OddsData(umaNoIndex + addint), " ", "") = "") Then
                            Exit For
                        End If

                        'MsgBox("OddsDataBeforeへ設定")
                        OddsList.Add(OddsData(umaNoIndex + addint) & "," & _
                                           OddsData(tanOddsIndex + addint) & "," & _
                                           OddsData(fukuOddsIndex + addint))
                        addint = addint + 3
                    Next
                    Exit For
                End If
            End If
        Next

        'For x = 0 To OddsList.Count - 1
        '    Dim a As String = OddsList(x)

        '    'MsgBox("単複分解後" & a)

        'Next

        '欠馬オッズの書き換えを実施
        ChangeNotRunOdds(OddsList)

        '単勝／複勝オッズ情報から一番人気の馬番号を取得
        'If TanshouMaxNinkiUmaNo = String.Empty Then
        GetNinkiUmaNo(TanshouMaxNinkiUmaNo, OddsList, oddsKbn)
        'End If


        '馬単オッズ情報の分解＆設定を実施
        For i = 0 To UmatanAllOddsInfo.Count - 1

            'Dim OddsData As String() = UmatanAllOddsInfo(i).Split(","c)
            Dim OddsData As String() = UmatanAllOddsInfo(i)
            'キー情報を取得
            Dim OddsKeyData As String = OddsData(CommonConstant.AllOdKeyData)

            '選択されたキー情報と一致するかを判定
            If (CheckKey = OddsKeyData) Then

                Dim OddsTime As String = OddsData(CommonConstant.AllOdUpTime)

                'オッズデータの発表時間とリストボックスで選択された発表時間を比較
                If (SelTime <= OddsTime) Then

                    addint = 0
                    '一致した場合
                    Do While 1

                        If (OddsData.Length - 1 < umatanOddsIndex + addint) Then

                            Exit Do

                        End If


                        '組み合わせを取得
                        Dim Kumiawase As String = OddsData(kumiawaseIndex + addint)
                        '空の場合、処理を終了
                        If (Kumiawase.Length = 0) Then
                            Exit Do
                        End If

                        Dim KumiawaseJikuUmaNo As String = Kumiawase.Substring(0, 2)
                        Dim KumiawaseFukuUmaNo As String = Kumiawase.Substring(2, 2)
                        Dim KumiawaseOdds As String = OddsData(umatanOddsIndex + addint)

                        '主軸馬番号と一番人気馬番号を比較する
                        If (KumiawaseJikuUmaNo = TanshouMaxNinkiUmaNo) Then
                            '人気馬と一致した場合、複馬及び組み合わせのオッズをセット
                            umatanWorkList.Add(KumiawaseFukuUmaNo & "," & KumiawaseOdds)

                        End If


                        addint = addint + 2
                    Loop

                    Exit For

                End If
            End If
        Next


        '馬単オッズ情報をオッズリストに追加する
        setAddOdds(TanshouMaxNinkiUmaNo, umatanWorkList, OddsList)


        'For x = 0 To OddsList.Count - 1
        '    Dim a As String = OddsList(x)

        '    MsgBox("馬単結合後" & a)

        'Next


        'マジックゾーンリスト作成

        'オッズリストをベースに作成する為、リストのクローンを設定する
        'MzList = OddsList.Clone

        For i = 0 To UmarenAllOddsInfo.Count - 1

            'Dim OddsData As String() = UmarenAllOddsInfo(i).Split(","c)
            Dim OddsData As String() = UmarenAllOddsInfo(i)

            'キー情報を取得
            Dim OddsKeyData As String = OddsData(CommonConstant.AllOdKeyData)

            '選択されたキー情報と一致するかを判定
            If (CheckKey = OddsKeyData) Then

                Dim OddsTime As String = OddsData(CommonConstant.AllOdUpTime)

                'オッズデータの発表時間とリストボックスで選択された発表時間を比較
                If (SelTime <= OddsTime) Then

                    addint = 0
                    '一致した場合
                    Do While 1

                        If (OddsData.Length - 1 < umatanOddsIndex + addint) Then

                            Exit Do

                        End If


                        '組み合わせを取得
                        Dim Kumiawase As String = OddsData(kumiawaseIndex + addint)
                        '空の場合、処理を終了
                        If (Kumiawase.Length = 0) Then
                            Exit Do
                        End If
                        Dim KumiawaseJikuUmaNo As String = Kumiawase.Substring(0, 2)
                        Dim KumiawaseFukuUmaNo As String = Kumiawase.Substring(2, 2)
                        Dim KumiawaseOdds As String = OddsData(umatanOddsIndex + addint)

                        '主軸馬番号／複馬番号と一番人気馬番号を比較する
                        If (KumiawaseJikuUmaNo = TanshouMaxNinkiUmaNo Or KumiawaseFukuUmaNo = TanshouMaxNinkiUmaNo) Then
                            '一番人気馬が組合せに入っていた場合、人気馬ではないほうの馬番号及び組み合わせのオッズをセット

                            If (KumiawaseJikuUmaNo <> TanshouMaxNinkiUmaNo) Then
                                '軸馬<>一番人気馬の場合、軸馬番号及び組合せオッズをセット
                                umarenWorkList.Add(KumiawaseJikuUmaNo & "," & KumiawaseOdds)

                            Else
                                '軸馬=一番人気馬の場合、複馬番号及び組合せオッズをセット
                                umarenWorkList.Add(KumiawaseFukuUmaNo & "," & KumiawaseOdds)

                            End If


                        End If
                        addint = addint + 2
                    Loop

                    Exit For

                End If
            End If
        Next

        '馬連オッズ情報をオッズリストに追加する
        setAddOdds(TanshouMaxNinkiUmaNo, umarenWorkList, OddsList)

        'For x = 0 To OddsList.Count - 1
        '    Dim a As String = OddsList(x)

        '    MsgBox("馬連結合後" & a)

        'Next


        'マジックゾーンリストにクローンを設定する
        MzList = OddsList.Clone

    End Sub

    '単勝／複勝オッズリストから一番人気の馬番号を取得する
    Private Sub GetNinkiUmaNo(ByRef TanshouMaxNinkiUmaNo As String, ByVal OddsList As ArrayList, ByVal oddsKbn As Integer)

        '単勝一番人気の馬を取得する
        Dim tanMinOddsDataH As String() = {}
        Dim readOddsData As String
        Dim readOddsDataH As String()
        For i = 0 To OddsList.Count - 1

            readOddsData = OddsList(i)
            readOddsDataH = readOddsData.Split(","c)

            If (i = 0) Then
                tanMinOddsDataH = readOddsDataH

                Continue For

            End If

            '2011/11/08 d-kobayashi update start
            '2番目の表のみ複勝の人気を取得する
            If oddsKbn = 1 Then
                If (tanMinOddsDataH(CommonConstant.OdFukushouOddsPos) > readOddsDataH(CommonConstant.OdFukushouOddsPos)) Then
                    '今回読み込みレコードと複勝オッズマスタのオッズを比較し、
                    '今回読み込みレコードのオッズが低ければ、複勝オッズマスタの内容を更新する
                    tanMinOddsDataH = readOddsDataH

                ElseIf (tanMinOddsDataH(CommonConstant.OdFukushouOddsPos) = readOddsDataH(CommonConstant.OdFukushouOddsPos)) Then
                    '今回読み込みレコードと複勝オッズマスタのオッズが同一の場合、
                    '単勝オッズをで比較する
                    If (tanMinOddsDataH(CommonConstant.OdTanshouOddsPos) > readOddsDataH(CommonConstant.OdTanshouOddsPos)) Then

                        tanMinOddsDataH = readOddsDataH

                    End If
                End If
            Else
                If (tanMinOddsDataH(CommonConstant.OdTanshouOddsPos) > readOddsDataH(CommonConstant.OdTanshouOddsPos)) Then
                    '今回読み込みレコードと単勝オッズマスタのオッズを比較し、
                    '今回読み込みレコードのオッズが低ければ、単勝オッズマスタの内容を更新する
                    tanMinOddsDataH = readOddsDataH

                ElseIf (tanMinOddsDataH(CommonConstant.OdTanshouOddsPos) = readOddsDataH(CommonConstant.OdTanshouOddsPos)) Then
                    '今回読み込みレコードと単勝オッズマスタのオッズが同一の場合、
                    '複勝オッズをで比較する
                    If (tanMinOddsDataH(CommonConstant.OdFukushouOddsPos) > readOddsDataH(CommonConstant.OdFukushouOddsPos)) Then

                        tanMinOddsDataH = readOddsDataH

                    End If
                End If
            End If
            '2011/11/08 d-kobayashi update end
        Next

        '一番人気馬の馬番号を設定
        TanshouMaxNinkiUmaNo = tanMinOddsDataH(0)

    End Sub

    '馬単／馬連オッズの結合を行う
    '<INPUT>
    ' TanshouMaxNinkiUmaNo：一番人気馬番号
    ' UTOddsData：追加オッズリスト（馬番号,オッズ）
    '
    '<OUTPUT>
    ' OddsList：オッズ情報
    '

    Private Sub setAddOdds(ByVal TanshouMaxNinkiUmaNo As String, ByVal AddOddsData As ArrayList, ByRef OddsList As ArrayList)


        Dim UmatanBaseData As String() = {}
        Dim readOddsDataH As String() = {}
        Dim GitouFlg As Boolean = False


        'オッズリストの件数分、処理を繰り返す
        For j = 0 To OddsList.Count - 1
            'オッズデータを取得
            Dim TFOddsDataH As String() = OddsList(j).Split(","c)
            'オッズデータから馬番号を取得
            Dim TFOddsUmaNo As String = TFOddsDataH(0)

            If (TFOddsUmaNo = TanshouMaxNinkiUmaNo) Then
                '一番人気馬の場合はオッズに空を設定する
                'オッズデータの内容を更新する
                OddsList.Item(j) = Join(TFOddsDataH, ",") & "," & ""
                'MsgBox("一番人気馬： " & OddsList.Item(j))

                Continue For

            End If

            For i = 0 To AddOddsData.Count - 1
                '追加オッズ情報リストから１件取得
                Dim UmatanDataH As String() = AddOddsData(i).split(","c)
                '追加オッズ情報の馬番を取得
                Dim UmatanUmaNo As String = UmatanDataH(0)
                Dim UmatanOdds As String = UmatanDataH(1)

                If Not IsNumeric(UmatanOdds) Then
                    UmatanOdds = CommonConstant.NotRunOdds
                End If

                '追加オッズ情報-オッズデータの馬番号を比較
                If (UmatanUmaNo = TFOddsUmaNo) Then
                    '馬番が一致した場合
                    'オッズデータの内容を更新する
                    OddsList.Item(j) = Join(TFOddsDataH, ",") & "," & UmatanOdds
                    'MsgBox("馬単オッズ設定： " & OddsList.Item(j))
                End If
            Next

        Next


    End Sub

    '各リストに法則該当カウントを設定するエリアを追加する

    'Public Sub setHousokuCountArea(ByRef setList As ArrayList)

    '    Dim OddsSetArea As String() = {"0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"}
    '    Dim MzSetArea As String() = {"0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"}

    '    For i = 1 To setList.Count - 1

    '        Dim readData As String() = setList(i).split(","c)

    '        If (readData.Length > 4) Then
    '            '要素数が５以上の場合はマジックゾーンリスト
    '            setList.Item(i) = Join(readData, ",") & "," & Join(MzSetArea, ",")
    '        Else
    '            '要素数が４以下の場合はオッズリスト
    '            setList.Item(i) = Join(readData, ",") & "," & Join(OddsSetArea, ",")

    '        End If
    '    Next

    'End Sub

    '各時系列毎の馬別ポイントを保有するリストを生成する。
    'ポイントリストはインデックスと馬番号を紐付け、ソートしても馬番号でポイントを引けるようにする
    Public Sub initPointList(ByVal OddsList As ArrayList, ByRef PointList As ArrayList)

        Dim OddsCount As Integer = OddsList.Count
        Dim initValue As String() = {"", "0"}

        PointList = New ArrayList

        'ポイントリストに初期値を設定する
        'For i = 0 To OddsCount
        '    PointList.Add(initValue)

        'Next

        For i = 0 To OddsCount - 1

            Dim readData As String() = OddsList(i).split(","c)
            '馬番号を取得
            Dim UmaNo As String = readData(CommonConstant.OdUmaNo)
            Dim setValue As String() = {UmaNo, "0"}
            'ポイントリストを更新
            PointList.Add(setValue)

        Next


    End Sub
    '馬毎の合計ポイントを算出する
    Public Sub PointCalculation(ByVal selTimeCount As Integer, _
                                ByVal EightHousokuPointList As ArrayList, _
                                ByVal NineHousokuPointList As ArrayList, _
                                ByVal TenHousokuPointList As ArrayList, _
                                ByRef HousokuPointList As ArrayList)

        Dim counter As Integer = EightHousokuPointList.Count - 1
        Dim UmaNo As String = String.Empty
        Dim PointE As Integer = 0
        Dim PointN As Integer = 0
        Dim PointT As Integer = 0

        '法則ポイントリストの要素削除
        HousokuPointList.Clear()

        For i = 0 To counter
            '８時時点データ
            Dim readDataE As String() = EightHousokuPointList(i)
            UmaNo = readDataE(CommonConstant.OdUmaNo)
            PointE = Integer.Parse(readDataE(CommonConstant.IndexPos1))

            '発表時間が２つ選択されていた場合
            If (selTimeCount >= CommonConstant.Const1) Then
                '９時時点データ
                For j = 0 To NineHousokuPointList.Count - 1
                    Dim readDataN As String() = NineHousokuPointList(j)
                    '９時の馬番号を取得
                    Dim NUmaNo As String = readDataN(CommonConstant.OdUmaNo)
                    If (UmaNo = NUmaNo) Then
                        PointN = Integer.Parse(readDataN(CommonConstant.IndexPos1))
                    End If
                Next

            End If

            '発表時間が３つ以上選択されていた場合
            If (selTimeCount >= CommonConstant.Const2) Then
                '１０時時点データ
                For j = 0 To TenHousokuPointList.Count - 1
                    Dim readDataT As String() = TenHousokuPointList(j)
                    '１０時の馬番号を取得
                    Dim TUmaNo As String = readDataT(CommonConstant.OdUmaNo)
                    If (UmaNo = TUmaNo) Then
                        PointT = Integer.Parse(readDataT(CommonConstant.IndexPos1))
                    End If
                Next
            End If

            Dim AllPoint As Integer = PointE + PointN + PointT

            '法則ポイントリストに設定する
            HousokuPointList.Add(AllPoint & "," & UmaNo)


        Next

    End Sub

    '該当レースの騎手／馬名情報を抽出し、リストへ設定する
    Public Sub setHorseData(ByVal checkKey As String, ByVal AllHorseDataList As ArrayList, ByRef HorseDataList As ArrayList)

        '騎手／馬名情報リストを初期化
        HorseDataList.Clear()

        For i = 0 To AllHorseDataList.Count - 1

            Dim readData As String() = AllHorseDataList(i)
            '騎手／馬名情報のキーを取得
            Dim readDataKey As String = readData(CommonConstant.IndexPos0)

            If (checkKey = readDataKey) Then
                'キーが一致した場合、「馬番号」、「馬名」、「騎手名」をリストに設定する
                Dim setData As String() = {readData(CommonConstant.HorseDataUmaNo), _
                                           readData(CommonConstant.HorseDataHorseName), _
                                           readData(CommonConstant.HorseDataKishuName)}

                HorseDataList.Add(setData)

            End If


        Next


    End Sub
    '出馬しない馬のオッズの再設定を実施する
    Public Sub ChangeNotRunOdds(ByRef OddsList As ArrayList)
        'オッズリストの件数を取得
        Dim OddsListLoopCounter As Integer = OddsList.Count - 1

        'オッズリストの件数分繰り返し
        For i = 0 To OddsListLoopCounter

            Dim readData As String() = OddsList(i).split(","c)
            '単勝オッズを判定し、数値以外が設定されていた場合、欠馬と判断する
            If Not IsNumeric(readData(CommonConstant.OdTanshouOddsPos)) Then
                '
                Dim setData As String() = {readData(CommonConstant.OdUmaNo), CommonConstant.NotRunOdds, CommonConstant.NotRunOdds}
                OddsList.Item(i) = Join(setData, ",")

            End If

        Next

    End Sub

    '順位が５～７位の馬に固定で５ポイント加算する
    Public Sub AddSpecialPoint(ByVal OddsList As ArrayList, ByRef PointList As ArrayList)

        For i = 4 To OddsList.Count - 1
            Dim readData As String() = OddsList(i).split(","c)
            '馬番号を取得
            Dim HorseNo As String = readData(CommonConstant.OdUmaNo)

            For j = 0 To PointList.Count - 1
                Dim PointReadData As String() = PointList(j)
                '馬番号を取得
                Dim PointHorseNo As String = PointReadData(CommonConstant.IndexPos0)
                If (HorseNo = PointHorseNo) Then

                    PointReadData(CommonConstant.IndexPos1) = Integer.Parse(PointReadData(CommonConstant.IndexPos1)) + CommonConstant.AddSpecialPoint
                    'リストを更新
                    PointList.Item(j) = PointReadData
                    Exit For
                End If

            Next

            '８頭以降は処理不要
            If (i >= CommonConstant.Const6) Then
                Exit For
            End If

        Next

    End Sub
End Class

