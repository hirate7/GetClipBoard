Imports GetClipBoard.CMyNumAPI2
Imports GetClipBoard.CJACCESS3
Imports GetClipBoard.CGetClipBoard
Imports Microsoft.Win32

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

            If Left_(S, 4) <> "資格情報" And Left_(S, 4) <> "照会番号" Then
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

                        Case "照会番号"
                            .ReferenceNumber = S2(1)

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

End Class
