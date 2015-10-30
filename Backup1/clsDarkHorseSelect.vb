Public Class clsDarkHorseSelect

    Public Function PointCalculation(ByVal RunHorseList_10 As ArrayList, ByVal PointList As ArrayList) As ArrayList
        '穴馬候補となる６頭を選定する。ただし、同一ポイントの馬が複数存在した場合はそれすべてで１頭とする


        '穴馬候補リスト
        Dim AnaUmaList As New ArrayList
        '穴馬数カウンター
        Dim AnaUmaCount As Integer
        '前回ポイント格納エリア
        Dim WorkPoint As String = "0"

        'ポイント順にソートする
        PointList.Sort()

        'ポイントを降順に直す
        PointList.Reverse()

 
        For i = 0 To PointList.Count - 1

            Dim GetPointData As String = PointList(i)
            Dim PointData As String() = GetPointData.Split(","c)

            For j = 4 To RunHorseList_10.Count - 1
                '軸馬以外の出馬リストを作成
                Dim WorkRunHorseData As String = RunHorseList_10(j)
                Dim RunHorseData As String() = WorkRunHorseData.Split(","c)

                '馬番号を比較
                If (PointData(1) = RunHorseData(0)) Then
                    '一致した場合、穴馬候補とする（設定形式：馬番号（ポイント））
                    AnaUmaList.Add(PointData(1) & "(" & PointData(0) & ")")

                End If
            Next

            '前回ポイントと今回ポイントが相違する場合、カウントアップ
            If (WorkPoint <> PointData(0)) Then

                AnaUmaCount = AnaUmaCount + 1

            End If

            'カウントが6未満の場合は繰り返し処理を継続、6以上になった場合は処理を終了
            If (AnaUmaCount < 6) Then

                WorkPoint = PointData(0)

            Else

                Exit For

            End If
        Next

        '穴馬候補リスト返却
        Return AnaUmaList

    End Function

    
End Class
