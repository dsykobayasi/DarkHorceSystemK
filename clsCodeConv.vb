Option Strict Off
Option Explicit On
Friend Class clsCodeConv
    ' コード名称取得モジュール
    ' コードから名称を取得する
    '
    Private Structure mudtCodeLine ''コード名称構造
        Dim strCodeNo As String ''コードNo.
        Dim StrCode As String ''コード
        Dim strNames As String ''名称列（複数ある場合，連続した文字列のまま格納）
    End Structure

    Private mFileName As String ''入力ファイル名
    Private mArrData() As clsCodeConv.mudtCodeLine ''コード名称データ
    Private blnFlag As Boolean ''データ読込確認フラグ


    ' @(f)
    '
    ' 機能　　 : データの格納
    '
    ' 引き数　 : ARG1 - ファイル名
    '
    ' 返り値　 : なし
    '
    ' 機能説明 : 指定されたファイルのデータをメモリ上に格納する
    '
    Public WriteOnly Property FileName() As String
        Set(ByVal Value As String)
            On Error GoTo err_Renamed
            mFileName = Value
            Call SetData()
ext:
            Exit Property
err_Renamed:
            MsgBox(Err.Description)
            Resume ext
        End Set
    End Property

    ' @(f)
    '
    ' 機能　　 : 名称の取得
    '
    ' 引き数　 : ARG1 - コードNo.
    ' 　　　　   ARG2 - コード
    '
    ' 返り値　 : 名称
    '
    ' 機能説明 : メモリ上に格納したデータをコードにより検索し名称を取得する
    '
    Public Function GetCodeName(ByVal strCodeNo As String, ByVal StrCode As String, Optional ByVal intNo As Short = 1) As String
        On Error GoTo err_Renamed
        Dim i As Short 'ループカウンタ
        Dim j As Short 'ループカウンタ
        Dim ct As Short '名称取得用カウンタ
        Dim strName As String '名称

        'データが読み込めていない場合

        If Not blnFlag Then
            GetCodeName = ""
            GoTo ext
        End If

        '名称文字列から指定番目の名称を返す
        For i = 0 To UBound(mArrData, 1)
            If mArrData(i).strCodeNo = strCodeNo And mArrData(i).StrCode = StrCode Then
                ct = 1
                For j = 1 To Len(mArrData(i).strNames)
                    If Mid(mArrData(i).strNames, j, 1) = "," Then
                        ct = ct + 1
                        If ct > intNo Then Exit For
                    ElseIf ct = intNo Then
                        strName = strName & Mid(mArrData(i).strNames, j, 1)
                    End If
                Next j
                Exit For
            End If
        Next i
        GetCodeName = strName

ext:
        Exit Function
err_Renamed:
        MsgBox(Err.Description)
        Resume ext
    End Function

    ' @(f)
    '
    ' 機能　　 : データの開放
    '
    ' 引き数　 : なし
    '
    ' 返り値　 : なし
    '
    ' 機能説明 : メモリ上に格納したデータを開放する
    '
    Private Sub Class_Terminate_Renamed()
        On Error GoTo err_Renamed
        Erase mArrData
ext:
        Exit Sub
err_Renamed:
        MsgBox(Err.Description)
        Resume ext
    End Sub
    Protected Overrides Sub Finalize()
        Class_Terminate_Renamed()
        MyBase.Finalize()
    End Sub

    ' @(f)
    '
    ' 機能　　 : データを1行ずつ処理
    '
    ' 引き数　 : なし
    '
    ' 返り値　 : なし
    '
    ' 機能説明 : CSVデータを1行分ずつ区切って処理する
    '
    Private Function SetData() As Object
        On Error GoTo err_Renamed
        Dim strRt As String '改行文字
        Dim lngLnRt As Integer '改行文字の文字数
        Dim strData As String 'CSVファイルを受ける文字列
        Dim lnglenData As Integer 'strDataの文字数
        Dim lngRt As Integer 'strData中のstrRtの位置
        Dim lngCt As Integer 'Rtのカウンタ，mArrDataの行数
        Dim lngBeforeRt As Integer 'ひとつ前のlngRt
        Dim strLine As String 'CSVファイル一行分
        Dim intFileNo As Short '使用するファイルNo.
        Dim bytData() As Byte 'ファイルのデータ格納先

        blnFlag = True

        '改行文字の決定
        strRt = vbCrLf
        lngLnRt = Len(strRt)

        'ファイルの中身を文字列として取得
        intFileNo = FreeFile()
        FileOpen(intFileNo, mFileName, OpenMode.Binary, OpenAccess.Read)
        ReDim bytData(LOF(intFileNo) - 1)
        FileGet(intFileNo, bytData)
        FileClose(intFileNo)

        'エンコード
        strData = System.Text.Encoding.GetEncoding(932).GetString(bytData)

        '配列クリア
        Erase bytData

        lnglenData = Len(strData)

        '中身が空，もしくはファイルが存在しない場合
        If Len(strData) = 0 Then
            blnFlag = False
            GoTo ext
        End If

        '一行ずつ処理
        lngBeforeRt = 1 - lngLnRt '一行目の前に改行があると仮定
        Do While lngRt < lnglenData
            lngRt = InStr(lngRt + 1, strData, strRt, CompareMethod.Binary)
            If lngRt = 0 Then Exit Do
            ReDim Preserve mArrData(lngCt)
            strLine = Mid(strData, lngBeforeRt + lngLnRt, lngRt - lngBeforeRt - lngLnRt)
            SetLine(strLine, lngCt)
            lngCt = lngCt + 1
            lngBeforeRt = lngRt
        Loop

ext:
        Exit Function
err_Renamed:
        blnFlag = False
        MsgBox(Err.Description)
        Resume ext
    End Function

    ' @(f)
    '
    ' 機能　　 : 配列に格納
    '
    ' 引き数　 : ARG1 - 一行分の文字列
    ' 　　　　 : ARG2 - 現在の行番号
    '
    ' 返り値　 : なし
    '
    ' 機能説明 : 1行分を構造体に変換して配列に格納する
    '
    Private Function SetLine(ByRef strLine As String, ByRef lngCt As Integer) As Object
        On Error GoTo err_Renamed
        Dim bytFieldCt As Byte 'フィールド（列）数
        Dim strDelimiter As String '区切り子
        Dim lngDelimiter As Integer '区切り子の位置
        Dim lngBeforeDel As Integer '前の区切り子の位置
        Dim strWord As String 'フィールド1つ分の文字列
        Dim udtWords As clsCodeConv.mudtCodeLine '一行分のstrWordを格納

        '区切り子の決定
        strDelimiter = ","

        'ユーザ定義型mudtCodeLineに変換
        Do While bytFieldCt <= 2
            If bytFieldCt < 2 Then
                lngDelimiter = InStr(lngDelimiter + 1, strLine, strDelimiter, CompareMethod.Binary)
            Else
                lngDelimiter = Len(strLine) + 1
            End If

            'フィールドが2以下の場合
            If lngDelimiter = 0 Then MsgBox("CSVファイルが不正です") : blnFlag = False : GoTo ext

            strWord = Mid(strLine, lngBeforeDel + 1, lngDelimiter - lngBeforeDel - 1)

            Select Case bytFieldCt
                Case 0
                    udtWords.strCodeNo = strWord
                Case 1
                    udtWords.StrCode = strWord
                Case 2
                    udtWords.strNames = strWord
                Case Else
                    GoTo ext
            End Select

            bytFieldCt = bytFieldCt + 1
            lngBeforeDel = lngDelimiter
        Loop

        'ユーザ定義型mudtCodeLineを配列に代入
        mArrData(lngCt) = udtWords

ext:
        Exit Function
err_Renamed:
        MsgBox(Err.Description)
        Resume ext
    End Function
End Class