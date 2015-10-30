Imports System.Windows.Forms

Public Class dlgKakoDelete
    Public intSelectCode As Integer
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Dim strmsg As String
        Dim i As Integer
        Dim ctl As Object
        Dim selectedIndex As Integer
        Const MSG_DEL_0 As String = "２年"
        Const MSG_DEL_1 As String = "１年"
        Const MSG_DEL_2 As String = "半年"
        Const MSG_DEL_3 As String = "１ヶ月"
        Const MSG_DEL_4 As String = "１５日"

        strmsg = "過去"
        For i = 0 To 4
            ctl = Me.Controls.Find("rbDelList" & CStr(i), True)(0)
            If ctl.checked Then
                selectedIndex = i
                Exit For
            End If
        Next

        Select Case selectedIndex
            Case 0
                strmsg = strmsg & MSG_DEL_0
            Case 1
                strmsg = strmsg & MSG_DEL_1
            Case 2
                strmsg = strmsg & MSG_DEL_2
            Case 3
                strmsg = strmsg & MSG_DEL_3
            Case 4
                strmsg = strmsg & MSG_DEL_4
        End Select
        strmsg = strmsg & "以前のレース情報を削除します。"

        If MsgBox(strmsg & vbCrLf & "よろしいでしょうか？", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then

            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            PutIni( _
                CommonConstant.INI_Ap_KakoDelete _
              , CommonConstant.INI_Key_KakoDelete _
              , selectedIndex _
              , gIniFile)

            intSelectCode = selectedIndex
        Else
            Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        End If

        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    ''2010/11/15 d-kobayashi add
    ''INIファイルの内容チェック

    Private Sub IniCheck()

        'セクションとキーの存在チェックを行う。
        If GetIni( _
            CommonConstant.INI_Ap_KakoDelete _
          , CommonConstant.INI_Key_KakoDelete _
          , "-1" _
          , gIniFile) = "-1" Then

            ''セクションとキーが見つからなかった場合、追加する。
            Call PutIni( _
                    CommonConstant.INI_Ap_KakoDelete _
                  , CommonConstant.INI_Key_KakoDelete _
                  , CommonConstant.INI_DefaultValue_KakoDelete _
                  , gIniFile)

        End If
    End Sub

    Private Sub dlgKakoDelete2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim retIniValue As String
        Dim ctl As Object
        ''INIファイルの内容チェック
        'IniCheck()

        ''INIファイルからコンボボックス選択値を取得し、コンボボックスに反映させる。
        retIniValue = GetIni( _
                        CommonConstant.INI_Ap_KakoDelete _
                      , CommonConstant.INI_Key_KakoDelete _
                      , CommonConstant.INI_DefaultValue_KakoDelete _
                      , gIniFile)

        ctl = Me.Controls.Find("rbDelList" & retIniValue, True)(0)
        ctl.checked = True
    End Sub
End Class
