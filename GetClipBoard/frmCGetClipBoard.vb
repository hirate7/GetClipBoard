Imports GetClipBoard.CMyNumAPI2
Imports GetClipBoard.CJACCESS3
Imports GetClipBoard.CGetClipBoard
Imports Microsoft.Win32
Imports System.IO
Imports System.Text

Public Class frmCGetClipBoard

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

        Try

            Dim Res As New CClipData

            Dim S As String

            S = args.Text

            If mnuTransferToMaple.Checked Then
                txtRes.Text = S
                TranserToMaple(S + vbLf + "///" + vbLf + "転送")
                Return
            End If

            If Left_(S, 4) <> "資格情報" And Left_(S, 4) <> "照会番号" And Left_(S, 5) <> "未就学区分" Then
                Return
            End If

            If Left_(S, 8) = "資格情報" + vbTab + "END" Then
                Me.Close()
                Clipboard.Clear()
                Return
            End If

            txtRes.Text = S

            Res.RawData = S

            Dim S1() As String
            If InStr(S, vbCrLf) > 0 Then
                S1 = Split(S, vbCrLf)
            Else
                S1 = Split(S, vbLf)
            End If

            Dim S2() As String

            Dim Tag As String = ""

            With Res

                Dim I As Integer
                For I = 0 To UBound(S1)
                    S2 = Split(S1(I), vbTab)
                    If UBound(S2) = 0 Then
                        ReDim Preserve S2(1)
                        S2(1) = ""
                    End If
                    Select Case S2(0)

                        Case "資格情報", "裏面記載情報", "高齢受給者証", "資格情報(医療保険)", "裏面記載情報(医療保険)", "資格情報(医療扶助)", "裏面記載情報(医療扶助)"
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
                            Select Case Tag
                                Case "資格情報", "資格情報(医療保険)"
                                    If .NameKana IsNot Nothing Then
                                        UnexpectedError(S, S2(0), Tag)
                                        Return
                                    End If
                                    .NameKana = S2(1)
                                Case "資格情報(医療扶助)"

                                Case Else
                                    UnexpectedError(S, S2(0), Tag)
                                    Return
                            End Select

                        Case "氏名"
                            Select Case Tag
                                Case "資格情報", "資格情報(医療保険)"
                                    If .Name IsNot Nothing Then
                                        UnexpectedError(S, S2(0), Tag)
                                        Return
                                    End If
                                    .Name = S2(1)
                                Case "裏面記載情報", "裏面記載情報(医療保険)"
                                    If .NameOfOther IsNot Nothing Then
                                        UnexpectedError(S, S2(0), Tag)
                                        Return
                                    End If
                                    .NameOfOther = S2(1)
                                Case "資格情報(医療扶助)", "裏面記載情報(医療扶助)"

                                Case Else
                                    UnexpectedError(S, S2(0), Tag)
                                    Return
                            End Select

                        Case "氏名カナ"
                            Select Case Tag
                                Case "裏面記載情報", "裏面記載情報(医療保険)"
                                    If .NameOfOtherKana IsNot Nothing Then
                                        UnexpectedError(S, S2(0), Tag)
                                        Return
                                    End If
                                    .NameOfOtherKana = S2(1)
                                Case "裏面記載情報(医療扶助)"

                                Case Else
                                    UnexpectedError(S, S2(0), Tag)
                                    Return
                            End Select

                        Case "生年月日"
                            Select Case Tag
                                Case "資格情報", "資格情報(医療保険)"
                                    If .Birthdate IsNot Nothing Then
                                        UnexpectedError(S, S2(0), Tag)
                                        Return
                                    End If
                                    .Birthdate = S2(1)
                                Case "資格情報(医療扶助)"

                                Case Else
                                    UnexpectedError(S, S2(0), Tag)
                                    Return
                            End Select

                        Case "性別"
                            Select Case Tag
                                Case "資格情報", "資格情報(医療保険)"
                                    If .Sex1 IsNot Nothing Then
                                        UnexpectedError(S, S2(0), Tag)
                                        Return
                                    End If
                                    .Sex1 = S2(1)
                                Case "裏面記載情報", "裏面記載情報(医療保険)"
                                    If .Sex2 IsNot Nothing Then
                                        UnexpectedError(S, S2(0), Tag)
                                        Return
                                    End If
                                    .Sex2 = S2(1)
                                Case "資格情報(医療扶助)", "裏面記載情報(医療扶助)"

                                Case Else
                                    UnexpectedError(S, S2(0), Tag)
                                    Return
                            End Select

                        Case "郵便番号"
                            Select Case Tag
                                Case "資格情報", "資格情報(医療保険)"
                                    If .PostNumber IsNot Nothing Then
                                        UnexpectedError(S, S2(0), Tag)
                                        Return
                                    End If
                                    .PostNumber = S2(1)
                                Case "資格情報(医療扶助)"

                                Case Else
                                    UnexpectedError(S, S2(0), Tag)
                                    Return
                            End Select

                        Case "住所"
                            Select Case Tag
                                Case "資格情報", "資格情報(医療保険)"
                                    If .Address IsNot Nothing Then
                                        UnexpectedError(S, S2(0), Tag)
                                        Return
                                    End If
                                    .Address = S2(1)
                                Case "資格情報(医療扶助)"

                                Case Else
                                    UnexpectedError(S, S2(0), Tag)
                                    Return
                            End Select

                        Case "証区分"
                            Select Case Tag
                                Case "資格情報", "資格情報(医療保険)"
                                    If .InsuredCardClassification IsNot Nothing Then
                                        UnexpectedError(S, S2(0), Tag)
                                        Return
                                    End If
                                    .InsuredCardClassification = S2(1)
                                Case "資格情報(医療扶助)"

                                Case Else
                                    UnexpectedError(S, S2(0), Tag)
                                    Return
                            End Select

                        Case "有効開始日"
                            Select Case Tag
                                Case "資格情報", "資格情報(医療保険)"
                                    If .InsuredCardValidDate IsNot Nothing Then
                                        UnexpectedError(S, S2(0), Tag)
                                        Return
                                    End If
                                    .InsuredCardValidDate = S2(1)
                                Case "高齢受給者証"
                                    If .ElderlyRecipientValidStartDate IsNot Nothing Then
                                        UnexpectedError(S, S2(0), Tag)
                                        Return
                                    End If
                                    .ElderlyRecipientValidStartDate = S2(1)
                                Case Else
                                    UnexpectedError(S, S2(0), Tag)
                                    Return
                            End Select

                        Case "有効終了日"
                            Select Case Tag
                                Case "資格情報", "資格情報(医療保険)"
                                    If .InsuredCardExpirationDate IsNot Nothing Then
                                        UnexpectedError(S, S2(0), Tag)
                                        Return
                                    End If
                                    .InsuredCardExpirationDate = S2(1)
                                Case "高齢受給者証"
                                    If .ElderlyRecipientValidEndDate IsNot Nothing Then
                                        UnexpectedError(S, S2(0), Tag)
                                        Return
                                    End If
                                    .ElderlyRecipientValidEndDate = S2(1)
                                Case Else
                                    UnexpectedError(S, S2(0), Tag)
                                    Return
                            End Select

                        Case "資格取得年月日"
                            .QualificationDate = S2(1)

                        Case "負担割合"
                            Select Case Tag
                                Case "資格情報", "資格情報(医療保険)"
                                    If .InsuredPartialContributionRatio IsNot Nothing Then
                                        UnexpectedError(S, S2(0), Tag)
                                        Return
                                    End If
                                    .InsuredPartialContributionRatio = S2(1)
                                Case "高齢受給者証"
                                    If .ElderlyRecipientContributionRatio IsNot Nothing Then
                                        UnexpectedError(S, S2(0), Tag)
                                        Return
                                    End If
                                    .ElderlyRecipientContributionRatio = S2(1)
                                Case Else
                                    UnexpectedError(S, S2(0), Tag)
                                    Return
                            End Select

                        Case "本人・家族の別"
                            .PersonalFamilyClassification = S2(1)

                        Case "被保険者氏名"
                            .InsuredName = S2(1)

                        Case "照会番号"
                            .ReferenceNumber = S2(1)

                        Case "未就学区分"
                            If S2(1) = "義務教育就学前" Then
                                .PreschoolClassification = "1"
                            Else
                                UnexpectedError(S, S2(0), Tag)
                                Return
                            End If

                    End Select

                Next I

                '処理実行日時
                .ProcessExecutionTime = Now

                If .QualificationConfirmationDate = "" Then
                    Me.WindowState = FormWindowState.Normal
                    MsgBox("確認日不明の処理できないデータを受け取りました。", vbOKOnly Or vbCritical, Me.Text)
                    Me.WindowState = FormWindowState.Minimized
                    Return
                End If

            End With

            Open_Connection()
            Dim ResNo As Integer = Add_ClipData(Res)
            Close_Connection()
            Select Case ResNo
                Case 0
                    If mnuClearClipBoard.Checked Then
                        Clipboard.Clear()
                    End If
                Case Else
                    Me.WindowState = FormWindowState.Normal
                    MsgBox("コピー内容のデータベースへの書き込みでエラーが発生しました。" + vbCrLf + "エラー番号:" + CStr(ResNo), vbOKOnly Or vbCritical, Me.Text)
                    Me.WindowState = FormWindowState.Minimized
            End Select

        Catch ex As Exception
            'エラーメッセージを表示する
            MsgBox(ex.ToString, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical)
        End Try

    End Sub

    ''' <summary>
    ''' 想定外のエラー
    ''' </summary>
    ''' <param name="S"></param>
    Private Sub UnexpectedError(S As String, S0 As String, Tag As String)

        Me.WindowState = FormWindowState.Normal
        If MsgBox("想定外のデータのため処理できません。" + vbCrLf + "データをメープルに送信しますか？", vbYesNo Or MsgBoxStyle.DefaultButton1 Or vbCritical, Me.Text) = vbYes Then
            TranserToMaple(S + vbLf + "///" + vbLf + S0 + vbLf + Tag + vbLf + "想定外")
        End If
        Me.WindowState = FormWindowState.Minimized

    End Sub

    Private m_DoubleRun As Integer

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        ''イベントをイベントハンドラに関連付ける
        ''フォームコンストラクタなどの適当な位置に記述してもよい
        'AddHandler SystemEvents.SessionEnding,
        'AddressOf SystemEvents_SessionEnding

        JPath = New CJPath
        JPath.Read()

        Personal = New CPersonal
        Personal.Read()

        Me.Text = Me.Text + " - " + Personal.J_Yago

        ToolStripStatusLabel1.Text = JPath.CURRENT
        ToolStripStatusLabel2.Text = JPath.KDATA

        mnuAutoStartup.Checked = RaedStartupState()

        JEnv = New CJEnv(JPath.MPROG, "[設定]")

        If Val(JEnv.Get("コピー後クリップボード削除", "1")) = 1 Then
            mnuClearClipBoard.Checked = True
        Else
            mnuClearClipBoard.Checked = False
        End If

        'Open_Connection()

        'If CheckRunningApp("JSY96.EXE") = False Then
        '    'JSY96.EXEが起動していなければLockを初期化する
        '    'Lockが不完全状態になったときにJSY96.EXEを終了してGetClipBoard.exeを走らせれば復旧するように
        '    LockOFF(1)
        'End If

    End Sub

    ''ログオフ、シャットダウンしようとしているとき
    'Private Sub SystemEvents_SessionEnding(
    '        ByVal sender As Object,
    '        ByVal e As SessionEndingEventArgs)
    '    Dim s As String
    '    If e.Reason = SessionEndReasons.Logoff Then
    '        s = "ログオフしようとしています。"
    '    ElseIf e.Reason = SessionEndReasons.SystemShutdown Then
    '        s = "シャットダウンしようとしています。"
    '    End If
    '    If MessageBox.Show(s + vbNewLine + "キャンセルしますか？",
    '            "質問", MessageBoxButtons.YesNo) = DialogResult.Yes Then
    '        'キャンセルする
    '        e.Cancel = True
    '    End If
    '    Me.Close()

    'End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

        'If m_DoubleRun = False Then
        '    Close_Connection()
        'End If

        ''イベントを解放する
        ''フォームDisposeメソッド内の基本クラスのDisposeメソッド呼び出しの前に
        ''記述してもよい
        'RemoveHandler SystemEvents.SessionEnding,
        '    AddressOf SystemEvents_SessionEnding

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

    Private Sub frmCGetClipBoard_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        If m_DoubleRun = False Then
            'FormClosedイベントで終了するとシャットダウン時にエラーが出るのでここに置く必要がある
            'Close_Connection()
            'FileOpen(1, JPath.KDATA + "Success Close.txt", OpenMode.Output)
            'FileClose(1)
        End If

    End Sub

    Private Sub RegistStartup()

        'Runキーを開く
        Dim regkey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
            "Software\Microsoft\Windows\CurrentVersion\Run", True)
        '値の名前に製品名、値のデータに実行ファイルのパスを指定し、書き込む
        regkey.SetValue("GetClipBoard", JPath.CURRENT + "GetClipBoard.EXE " + JPath.KDATA0)
        '閉じる
        regkey.Close()

    End Sub

    Private Function RaedStartupState() As Boolean

        Dim S As String

        S = Microsoft.Win32.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", "GetClipBoard", "")

        If S <> "" Then
            Return True
        Else
            Return False
        End If

    End Function

    Private Sub DeleteStartup()

        'スタートアップのレジストリから削除
        'キーを書き込み許可で開く
        Dim regkey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Run", True)

        '指定したキーが見つからなくてもエラーは出ない
        regkey.DeleteValue("GetClipBoard", True)

        '閉じる
        regkey.Close()

    End Sub

    Private Sub mnuAutoStartup_Click(sender As Object, e As EventArgs) Handles mnuAutoStartup.Click

        If mnuAutoStartup.Checked Then
            RegistStartup()
        Else
            DeleteStartup()
        End If

    End Sub

    Private Sub TranserToMaple(ByVal postData As String)

        Dim url As String = "http://mtry.main.jp/sub/upload.php"

        Dim FNa As String

        FNa = Replace(Replace(Replace(Now.ToString, " ", ""), "/", ""), ":", "")

        FNa = "ClipBoard" + FNa + ".txt"

        Dim wc As New System.Net.WebClient()
        Dim enc As System.Text.Encoding = System.Text.Encoding.UTF8
        wc.Encoding = enc
        'wc.Headers.Add("Content-Length", CStr(System.Text.Encoding.GetEncoding("utf-8").GetByteCount(postData)))
        wc.Headers.Add("x-file-name", FNa)


        postData = postData + vbLf + Personal.J_Yago

        Dim res As String = ""
        Try

            res = wc.UploadString(url, postData)

        Catch Ex As Exception

        End Try

        If wc IsNot Nothing Then
            wc.Dispose()
        End If

    End Sub

    Private Sub mnuClearClipBoard_Click(sender As Object, e As EventArgs) Handles mnuClearClipBoard.Click

        If mnuClearClipBoard.Checked Then
            JEnv.Put("コピー後クリップボード削除", "1")
        Else
            JEnv.Put("コピー後クリップボード削除", "0")
        End If

    End Sub
End Class
