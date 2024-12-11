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
            Return
        End If

        txtRes.Text = S

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
                        If Tag = "資格情報" Then
                            .InsuredCardValidDate = S2(1)
                        ElseIf Tag = "高齢受給者証" Then
                            .ElderlyRecipientValidStartDate = S2(1)
                        End If

                    Case "有効終了日"
                        If Tag = "資格情報" Then
                            .InsuredCardExpirationDate = S2(1)
                        ElseIf Tag = "高齢受給者証" Then
                            .ElderlyRecipientValidEndDate = S2(1)
                        End If

                    Case "資格取得年月日"
                        .QualificationDate = S2(1)

                    Case "負担割合"
                        If Tag = "資格情報" Then
                            .InsuredPartialContributionRatio = S2(1)
                        ElseIf Tag = "高齢受給者証" Then
                            .ElderlyRecipientContributionRatio = S2(1)
                        End If

                    Case "本人・家族の別"
                        .PersonalFamilyClassification = S2(1)

                    Case "被保険者氏名"
                        .InsuredName = S2(1)

                End Select

            Next I

            '処理実行日時
            .ProcessExecutionTime = Now

            If .QualificationConfirmationDate = "" Then
                Me.WindowState = FormWindowState.Normal
                MsgBox("確認日不明の処理できないデータを受け取りました。", vbOKOnly Or vbCritical, Me.Text)
                Return
            End If

        End With

        Select Case Add_ClipData(Res)
            Case 0
            Case 1
                Me.WindowState = FormWindowState.Normal
                MsgBox("「カルテ入力」とのデータベースの共有で問題が発生しています。", vbOKOnly Or vbCritical, Me.Text)
            Case 2
                Me.WindowState = FormWindowState.Normal
                MsgBox("エラーのためマイナ資格確認アプリからのコピーが受け取れません。", vbOKOnly Or vbCritical, Me.Text)
        End Select

    End Sub

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

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        Open_Connection()

        If CheckRunningApp("JSY96.EXE") = False Then
            'JSY96.EXEが起動していなければLockを初期化する
            'Lockが不完全状態になったときにJSY96.EXEを終了してGetClipBoard.exeを走らせれば復旧するように
            LockOFF(1)
        End If

    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

        Close_Connection()

    End Sub

    ''' <summary>
    ''' 指定されたプログラムが起動中かどうか調べる
    ''' </summary>
    ''' <param name="FNa"></param>
    ''' <returns></returns>
    Function CheckRunningApp(FNa As String) As Boolean

        'ローカルコンピュータ上で実行されているすべてのプロセスを取得
        Dim ps As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcesses()

        '配列から1つずつ取りして比較する
        For Each p As System.Diagnostics.Process In ps

            Try
                If System.IO.Path.GetFileName(p.MainModule.FileName).ToUpper = FNa.ToUpper Then
                    Return True
                End If
            Catch ex As Exception

            End Try

        Next

        Return False

    End Function

End Class
