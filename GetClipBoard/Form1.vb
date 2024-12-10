Imports GetClipBoard.CMyNumAPI2
Imports GetClipBoard.CJACCESS3

Public Class Form1

    Private viewer As CGetClipBoard.MyClipboardViewer

    Public Sub New()

        Clipboard.SetDataObject(New DataObject())

        viewer = New CGetClipBoard.MyClipboardViewer(Me)
        ' イベントハンドラを登録
        AddHandler viewer.ClipboardHandler, AddressOf OnClipBoardChanged

        ' この呼び出しは、Windows フォーム デザイナで必要です。
        InitializeComponent()
    End Sub

    ' クリップボードにテキストがコピーされると呼び出される
    Private Sub OnClipBoardChanged(ByVal sender As Object, ByVal args As CGetClipBoard.ClipboardEventArgs)

        Dim Res As New CClipData

        Dim S As String

        S = args.Text

        If Left_(S, 4) <> "資格情報" Then
            Exit Sub
        End If

        TextBox1.Text = S

        Res.RawData = S

        Dim S1() As String
        S1 = Split(S, vbCrLf)

        Dim S2() As String

        Dim Tag As String = ""

        With Res

            Dim I As Integer
            For I = 0 To UBound(S1)
                S2 = Split(S1(I), vbTab)
                Select Case S2(0)

                    Case "資格情報", "裏面記載情報", "高齢受給者証"
                        Tag = S2(0)

                    Case "確認日 : "
                        .QualificationConfirmationDate = S2(1)

                    Case "保険者番号"
                        .InsurerNumber = S2(1)

                    Case "保険者名"
                        .InsurerName = S2(1)

                    Case "記号"
                        .InsuredCardSymbol = S2(1)

                    Case "番号"
                        .InsuredIdentificationNumber = S2(1)

                    Case "枝番"
                        .InsuredBranchNumber = S2(1)

                    Case "フリガナ"
                        .NameKana = S2(1)

                    Case "氏名"
                        If Tag = "資格情報" Then
                            .Name = S2(1)
                        ElseIf Tag = "裏面記載情報" Then
                            .NameOfOther = S2(1)
                        End If

                    Case "氏名カナ"
                        .NameOfOtherKana = S2(1)

                    Case "生年月日"
                        .Birthdate = S2(1)

                    Case "性別"
                        If Tag = "資格情報" Then
                            .Sex1 = S2(1)
                        ElseIf Tag = "裏面記載情報" Then
                            .Sex2 = S2(1)
                        End If

                    Case "証区分"
                        .InsuredCardClassification = S2(1)

                    Case "有効開始日"
                        .InsuredCardValidDate = S2(1)

                    Case "有効終了日"
                        .InsuredCardExpirationDate = S2(1)

                    Case "資格取得年月日"
                        .QualificationDate = S2(1)

                    Case "負担割合"
                        .InsuredPartialContributionRatio = S2(1)

                    Case "本人・家族の別"
                        .PersonalFamilyClassification = S2(1)

                    Case "被保険者氏名"
                        .InsuredName = S2(1)

                    Case "有効開始日"
                        .ElderlyRecipientValidStartDate = S2(1)

                    Case "有効終了日"
                        .ElderlyRecipientValidEndDate = S2(1)

                    Case "負担割合"
                        .ElderlyRecipientContributionRatio = S2(1)

                End Select

            Next I

            '処理実行日時
            .ProcessExecutionTime = Now

        End With

        Add_ClipData(Res)

    End Sub

    Function Seireki_To_MyNum(D As String) As String

        Return Left_(D, 4) + Mid(D, 6, 2) + Mid(D, 9, 2)

    End Function

    Public Shared Function Left_(ByVal S As String, ByVal I As Integer) As String
        'Microsoft.VisualBasic.Left()と同じ
        'Inports Microsoft.VisualBasicを書いてもMicrosoft.VisualBasicの
        '名前空間を指定しなければならない場合に名前が長くならないようにするため

        Return Microsoft.VisualBasic.Left(S, I)

    End Function

    Public Shared Function Right_(ByVal S As String, ByVal I As Integer) As String
        'Microsoft.VisualBasic.Left()と同じ
        'Inports Microsoft.VisualBasicを書いてもMicrosoft.VisualBasicの
        '名前空間を指定しなければならない場合に名前が長くならないようにするため

        Return Microsoft.VisualBasic.Right(S, I)

    End Function

    '*2019/06/12*
    Public Const NewGengou = "令和"     '新元号
    Public Const LastGengou = "令和"    '期間の終わりが未定の場合の元号
    Public Const NewGengouRoma = "R"    '新元号ローマ字
    Public Const NewGengouStartY = 2019 '新元号開始年(西暦)
    Public Const NewGengouStartM = 5    '新元号開始月
    Public Const NewGengouStartD = 1    '新元号開始日
    Public Const LastGengouEndY = 2019  '旧元号終了年(西暦)
    Public Const LastGengouEndM = 4     '旧元号終了月
    Public Const LastGengouEndD = 30    '旧元号終了日
    Public Const NewGengouDefYM = " 1年 5月"        'YM_Check3()で使用
    Public Const NewGengouDefYMD = " 1年 5月 1日"   'YM_Check3()で使用

    Public Const Max_Date& = 5373485    ' 日付通し番号の最大値(未定の値)

    <Serializable()> Structure Date_
        '/// 日付を取り扱うための構造体 ///
        '/// 32bit対応に伴いGengouを可変長にした ///
        '/// Rec_Data_SetがDate_型を含むのでRec_Data_Setは可変長になる ///
        Public Year As Integer        '西暦
        Public Gengou As String       '元号
        Public WYear As Integer       '和暦
        Public Month As Integer       '月

        Public Day As Integer         '日

        'Sub New()

        '    Year = 0
        '    Gengou = ""
        '    WYear = 0
        '    Month = 0
        '    Day = 0

        'End Sub

    End Structure

    Public Shared Function TrimX(ByVal S As String) As String
        '機能   :文字列の両端からスペースとTABとNullを外す
        '引渡:S        :文字列
        '戻り:S        :保存

        '    :TrimX    :変換結果

        Dim L As Integer
        Dim I As Integer
        Dim D As String
        Dim S0 As String

        S0 = S
        L = Len(S0)
        For I = 1 To L
            D = Mid$(S0, I, 1)
            If D <> " " And D <> "　" And D <> Chr(9) And D <> Chr(0) Then Exit For
        Next I
        If I = L + 1 Then
            Return ""
        End If

        S0 = Mid$(S0, I)

        L = Len(S0)
        For I = L To 1 Step -1
            D = Mid$(S0, I, 1)
            If D <> " " And D <> "　" And D <> Chr(9) And D <> Chr(0) Then Exit For
        Next I
        S0 = Left_(S0, I)
        Return S0

    End Function

    Public Shared Function Date_Maple_VB(ByVal L As Integer) As Date
        '
        '機能:柔整システムで使っている日付の通し番号をVBのDate型に変換する

        '引渡:L             :柔整システムで使っている日付の通し番号

        '戻り:L             :保存

        '    :Date_Maple_Jet:VBのDate型

        '                   :元が未定の時はNothingを返す
        '
        Dim D As Date
        Dim DSet As New Date_

        Select Case L
            Case 0, Max_Date&
                Date_Maple_VB = Nothing '西暦1年1月1日 00:00:00
            Case Else
                Date_N_Set(L, DSet)
                D = New Date(DSet.Year, DSet.Month, DSet.Day)
                Return D
                'Date_Maple_VB = CDate(L - 2415019)
        End Select

    End Function

    Public Shared Sub Date_N_Set(ByVal N As Integer, ByRef Da As Date_)

        '機能:通し番号から日付のセット（Date_)を得る
        '引渡:N         :通し番号
        '    :Date_     :
        '戻り:N         :保存

        '    :Date_     :ユーザー定義変数に日付を設定


        Dim Ye As Integer
        Dim N0 As Integer
        Dim A As Integer
        Dim I As Integer
        Dim M As Integer

        Const Y400& = 400& * 365& + 100& - 4& + 1&
        Const Y100& = 100& * 365& + 25& - 1&
        Const Y4& = 4& * 365& + 1&
        Const Y1& = 365&

        If N < 1721426 Or 5373484 < N Then

            '/// 通し番号が異常なときの処理 ///
            Da.Gengou = ""
            Da.WYear = 0
            Da.Month = 0
            Da.Day = 0
            Da.Year = 0
            Exit Sub

        End If

        N0 = N - 1721426        '西暦1年1月1日を0とする

        If N0 >= Y400 Then A = N0 \ Y400 : Ye = A * 400 : N0 = N0 - A * Y400
        If N0 >= Y100 Then
            A = N0 \ Y100
            If A > 3 Then A = 3
            Ye = Ye + A * 100
            N0 = N0 - A * Y100
        End If
        If N0 >= Y4 Then A = N0 \ Y4 : Ye = Ye + A * 4 : N0 = N0 - A * Y4
        If N0 >= Y1 Then
            A = N0 \ Y1
            If A > 3 Then A = 3
            Ye = Ye + A
            N0 = N0 - A * Y1
        End If
        Da.Year = Ye + 1

        A = 0
        For I = 1 To 12
            M = Max_Day_Month(Da.Year, I)
            A = A + M
            If N0 < A Then
                Da.Month = I
                Da.Day = N0 - A + M + 1
                Exit For
            End If
        Next I

        If Seireki_Wareki(Da) Then
            Da.Gengou = ""
            Da.WYear = 0
        End If

        Exit Sub

    End Sub

    ''' <summary>
    ''' 西暦から和暦への変換
    ''' 1868(M.1)/9/8～2087(H.99)/12/31内のみ有効
    ''' 無効のとき-1を返す
    ''' *2017/12/29*
    ''' </summary>
    ''' <param name="DD">
    ''' 引き渡し:DD.Gengou DD.WYear は参照されない
    ''' 戻り    :DD.Gengou DD.WYear がセットされる
    ''' </param>
    ''' <returns>=0 正常に変換   =-1 変換できない</returns>
    ''' <remarks></remarks>
    Public Shared Function Seireki_Wareki(ByRef DD As Date_) As Integer

        Dim D As Integer
        Const M0 = 1868 '明治元年

        D = (DD.Year - M0) * 10000& + DD.Month * 100 + DD.Day

        Select Case D
            Case (1868 - M0) * 10000& + 9 * 100 + 8 To (1912 - M0) * 10000& + 7 * 100 + 29
                DD.WYear = DD.Year - 1867 : DD.Gengou = "明治"
            Case (1912 - M0) * 10000& + 7 * 100 + 30 To (1926 - M0) * 10000& + 12 * 100 + 24
                DD.WYear = DD.Year - 1911 : DD.Gengou = "大正"
            Case (1926 - M0) * 10000& + 12 * 100 + 25 To (1989 - M0) * 10000& + 1 * 100 + 7
                DD.WYear = DD.Year - 1925 : DD.Gengou = "昭和"
            Case (1989 - M0) * 10000& + 1 * 100 + 8 To (LastGengouEndY - M0) * 10000& + LastGengouEndM * 100 + LastGengouEndD
                DD.WYear = DD.Year - 1988 : DD.Gengou = "平成"
            Case (NewGengouStartY - M0) * 10000& + NewGengouStartM * 100 + NewGengouStartD To (2087 - M0) * 10000& + 12 * 100 + 31
                DD.WYear = DD.Year - NewGengouStartY + 1 : DD.Gengou = NewGengou
            Case Else
                Seireki_Wareki = -1 : Exit Function
        End Select

        Seireki_Wareki = 0

    End Function

    Public Shared Function Date_Check1(ByRef G As String, ByRef S As String) As Integer
        '機能:日付の有効性のチェックと整理

        '    :[未定の処理]
        '    :G=何等かの元号 S="  年  月  日"は未定として有効とする
        '引渡:G         :元号
        '    :S         :"YY年MM月DD日"
        '戻り:G         :保存

        '    :S         :整理された日付(例:"01年2 月03日"->" 1年 2月 3日")
        '    :Date_Check:=0 有効 =-1 無効

        Dim DD As New Date_

        If S = "  年  月  日" And Trim(G) = "" Then Date_Check1 = 0 : Exit Function
        If S = "  年  月  日" And (InStr("明治 大正 昭和 平成 " + NewGengou + " ", G + " ") Mod 3) = 1 Then Date_Check1 = 0 : Exit Function

        If Len(S) <> 9 Then Date_Check1 = -1 : Exit Function

        DD.Gengou = G
        DD.WYear = Val(Left_(S, 2))
        DD.Month = Val(Mid$(S, 4, 2))
        DD.Day = Val(Mid$(S, 7, 2))

        If Date_Check2(DD) Then Date_Check1 = -1 : Exit Function

        S = Right_(Str$(DD.WYear), 2) + "年" + Right_(Str$(DD.Month), 2) + "月" + Right_(Str$(DD.Day), 2) + "日"
        Date_Check1 = 0

    End Function

    Public Shared Function Date_Check2(ByVal DD As Date_) As Integer
        '機能:日付の有効性のチェック
        '    :元号が""のときは西暦で判断する それ以外では和暦で判断する
        '    :西暦、和暦の年で使用しないほうは何が入っていても構わない

        '    :1868(M.1)/9/8～2087(H.99)/12/31内のみ有効
        '引渡:DD        :日付を表わす構造体

        '戻り:DD        :保存

        '    :Date_Check:=0 有効 =-1 無効

        'Dim Ye as integer

        If DD.Month < 1 Or DD.Month > 12 Then Date_Check2 = -1 : Exit Function
        If DD.Day < 1 Then Date_Check2 = -1 : Exit Function

        If TrimX(DD.Gengou) = "" Then

            If DD.Year < 1868 Or DD.Year > 2087 Then Date_Check2 = -1 : Exit Function

        Else

            If Wareki_Seireki(DD) Then Date_Check2 = -1 : Exit Function

        End If

        If DD.Day > Max_Day_Month(DD.Year, DD.Month) Then Date_Check2 = -1 : Exit Function

        Date_Check2 = 0

    End Function

    Public Shared Function Max_Day_Month(ByVal Ye As Integer, ByVal Mo As Integer) As Integer
        '機能   :指定年月の最終日を求める

        '引渡   :Ye             :年
        '       :Mo             :月

        '戻り   :Ye             :保存

        '       :Mo             :保存

        '       :Max_Day_Month  :最終日

        Dim Da As Integer

        Select Case Mo

            Case 1, 3, 5, 7, 8, 10, 12
                Da = 31
            Case 4, 6, 9, 11
                Da = 30
            Case 2
                Da = 28
                If (Ye Mod 4) = 0 Then Da = 29
                If (Ye Mod 100) = 0 Then Da = 28
                If (Ye Mod 400) = 0 Then Da = 29
        End Select

        Max_Day_Month = Da

    End Function

    ''' <summary>
    ''' 和暦から西暦への変換
    ''' 1868(M.1)/9/8～2087(H.99)/12/31内のみ有効
    ''' 無効のとき-1を返す
    ''' *2019/06/12*
    ''' </summary>
    ''' <param name="DD">
    ''' 引渡   :日付構造体 DD.Gengou DD.WYear DD.Month DD.Day をセットしておく
    ''' 戻り   :DD.Year がセットされる
    ''' </param>
    ''' <returns>=0 正常に変換   =1 変換失敗</returns>
    ''' <remarks></remarks>
    Public Shared Function Wareki_Seireki(ByRef DD As Date_) As Integer

        Dim d As Integer
        Dim S As Integer
        Dim E As Integer
        Dim Ye As Integer

        d = DD.WYear * 10000& + DD.Month * 100 + DD.Day

        Select Case DD.Gengou
            Case "明治"
                S = 1 * 10000 + 9 * 100 + 8 : E = 45 * 10000& + 7 * 100 + 29
                Ye = DD.WYear + 1867
            Case "大正"
                S = 1 * 10000 + 7 * 100 + 30 : E = 15 * 10000& + 12 * 100 + 24
                Ye = DD.WYear + 1911
            Case "昭和"
                S = 1 * 10000 + 12 * 100 + 25 : E = 64 * 10000& + 1 * 100 + 7
                Ye = DD.WYear + 1925
            Case "平成"
                S = 1 * 10000 + 1 * 100 + 8 : E = (LastGengouEndY - 1988) * 10000& + LastGengouEndM * 100 + LastGengouEndD
                Ye = DD.WYear + 1988
            Case NewGengou
                S = 1 * 10000 + NewGengouStartM * 100 + NewGengouStartD : E = 69 * 10000& + 12 * 100 + 31
                Ye = DD.WYear + NewGengouStartY - 1
            Case Else
                Wareki_Seireki = -1 : Exit Function
        End Select

        If d < S Or E < d Then Wareki_Seireki = -1 : Exit Function

        DD.Year = Ye
        Wareki_Seireki = 0

    End Function

    Public Shared Function Dnum(ByVal G As String, ByVal D As String) As Integer
        '機能:日付の文字列から通し番号を返す
        '    :[未定の処理]
        '    :元号如何に係わらずD="  年  月  日"ならば0を返す
        '    :未定が0のときはそのまま使えるが未定がMax_Date&のときは後で利用者が変更する
        '    :Date_Check1でチェックしてから渡すこと
        '引渡:G             :元号(例:"平成")
        '    :D             :日付(例:"01年2 月 3日")
        '戻り:G             :保存
        '　　:D             :保存
        '    :DNum          :通し番号  =0 未入力


        Dim DD As New Date_

        If Date_Check1(G, D) Then Throw New Exception()
        If D = "  年  月  日" Then Dnum = 0 : Exit Function

        DD.Gengou = G
        DD.WYear = Val(Left_(D, 2))
        DD.Month = Val(Mid$(D, 4, 2))
        DD.Day = Val(Mid$(D, 7, 2))

        Dnum = Date_Set_N(DD)

    End Function

    Public Shared Function Date_Set_N(ByRef Da As Date_) As Integer
        '機能:日付のセット（Date_)から通し番号を得る
        '    :Date_Check2でチェックしてから渡すこと
        '    :未定は扱えない

        '引渡:Date_     :ユーザー定義変数に日付を設定 和暦を優先

        '               :元号が"    "のときは西暦を優先

        '戻り:Date_     :保存

        '    :Date_Set_N:通し番号

        Dim N As Integer
        Dim Y0 As Integer
        Dim I As Integer

        'If Date_Check2(Da) Then Throw New Exception()

        If TrimX(Da.Gengou) <> "" Then

            If Wareki_Seireki(Da) Then Throw New Exception()

        End If

        Y0 = Da.Year - 1
        N = Y0 * 365& + (Y0 \ 400&) - (Y0 \ 100&) + (Y0 \ 4&)

        For I = 1 To Da.Month - 1
            N = N + Max_Day_Month(Da.Year, I)
        Next I

        N = N + Da.Day

        N = N + 1721426 - 1 ' 1/1/1を1721426&とする

        Date_Set_N = N

    End Function

    Public Shared Function Dnum(ByVal D As String) As Integer
        '機能:日付の文字列から通し番号を返す
        '    :[未定の処理]
        '    :元号如何に係わらずD="  年  月  日"ならば0を返す
        '    :未定が0のときはそのまま使えるが未定がMax_Date&のときは後で利用者が変更する
        '    :Date_Check1でチェックしてから渡すこと
        '引渡:D             :日付(例:"平成01年2 月 3日")
        '戻り:D             :保存
        '    :DNum          :通し番号  =0 未入力

        Dim G As String
        Dim S As String

        G = Left_(D, 2)
        S = Mid(D, 3)

        Return Dnum(G, S)

    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        Clipboard.SetDataObject(New DataObject())

        'viewer = New CGetClipBoard.MyClipboardViewer(Me)
        '' イベントハンドラを登録
        'AddHandler viewer.ClipboardHandler, AddressOf OnClipBoardChanged

        Open_Connection()

    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

        Close_Connection()

    End Sub
End Class
