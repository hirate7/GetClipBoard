Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.IO
Imports Microsoft.Win32

Public Class CGetClipBoard

    Public Shared Sub Main()

        '二重起動をチェックする
        If Diagnostics.Process.GetProcessesByName(
        Diagnostics.Process.GetCurrentProcess.ProcessName).Length > 1 Then
            'すでに起動していると判断して終了
            Return
        End If

        'Dim ps As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("jsy96")
        'If 0 < ps.Length Then

        '    '見つかった時は、アクティブにする
        '    'Application.DoEvents()
        '    'Dim DT As DateTime
        '    'DT = Now + New TimeSpan(0, 0, 1)
        '    'Do
        '    '    If Now > DT Then
        '    '        Me.Visible = True
        '    '        Exit Do
        '    '    End If
        '    'Loop

        '    'Microsoft.VisualBasic.Interaction.AppActivate(ps(0).Id)
        '    'Application.DoEvents()
        '    'Me.BackColor = Color.AliceBlue

        '    'Me.Hide()
        '    'Application.DoEvents()

        '    'Application.DoEvents()

        '    'Dim DT As DateTime
        '    'DT = Now + New TimeSpan(0, 0, 2)
        '    'Do
        '    '    If Now > DT Then
        '    '        'Me.Visible = True
        '    '        Exit Do
        '    '    End If
        '    'Loop

        'End If

        Dim frmGetClipBoard As New frmCGetClipBoard

        frmGetClipBoard.ShowDialog()

    End Sub

    Public Class ClipboardEventArgs
        Inherits EventArgs
        Private m_text As String

        Public ReadOnly Property Text() As String
            Get
                Return Me.m_text
            End Get
        End Property

        Public Sub New(ByVal str As String)
            Me.m_text = str
        End Sub
    End Class

    Public Delegate Sub cbEventHandler(ByVal sender As Object,
        ByVal ev As ClipboardEventArgs)

    <System.Security.Permissions.PermissionSet(
        System.Security.Permissions.SecurityAction.Demand,
        Name:="FullTrust")>
    Friend Class MyClipboardViewer
        Inherits NativeWindow

        <DllImport("user32")>
        Public Shared Function SetClipboardViewer(
                ByVal hWndNewViewer As IntPtr) As IntPtr
        End Function

        <DllImport("user32")>
        Public Shared Function ChangeClipboardChain(
                ByVal hWndRemove As IntPtr,
                ByVal hWndNewNext As IntPtr) As Boolean
        End Function

        <DllImport("user32")>
        Public Shared Function SendMessage(
                ByVal hWnd As IntPtr, ByVal Msg As Integer,
                ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
        End Function

        Private Const WM_DRAWCLIPBOARD As Integer = &H308
        Private Const WM_CHANGECBCHAIN As Integer = &H30D
        Private nextHandle As IntPtr

        Private parent As Form
        Public Event ClipboardHandler As cbEventHandler

        Public Sub New(ByVal parent As Form)
            AddHandler parent.HandleCreated, AddressOf Me.OnHandleCreated
            AddHandler parent.HandleDestroyed, AddressOf Me.OnHandleDestroyed
            Me.parent = parent

        End Sub

        Friend Sub OnHandleCreated(ByVal sender As Object, ByVal e As EventArgs)
            AssignHandle(DirectCast(sender, Form).Handle)
            ' ビューアを登録
            nextHandle = SetClipboardViewer(Me.Handle)
        End Sub

        Friend Sub OnHandleDestroyed(ByVal sender As Object, ByVal e As EventArgs)
            ' ビューアを解除
            Dim sts As Boolean = ChangeClipboardChain(Me.Handle, nextHandle)
            ReleaseHandle()
        End Sub

        Protected Overloads Overrides Sub WndProc(ByRef msg As Message)
            Select Case msg.Msg

                Case WM_DRAWCLIPBOARD
                    ' クリップボードの内容がテキストの場合
                    If Clipboard.ContainsText() Then
                        ' クリップボードの内容を取得してハンドラを呼び出す
                        RaiseEvent ClipboardHandler(
                                Me, New ClipboardEventArgs(Clipboard.GetText()))
                    End If

                    If CInt(nextHandle) <> 0 Then
                        SendMessage(nextHandle, msg.Msg, msg.WParam, msg.LParam)
                    End If
                    Exit Select

                ' クリップボード・ビューア・チェーンが更新された
                Case WM_CHANGECBCHAIN
                    If msg.WParam = nextHandle Then
                        nextHandle = msg.LParam
                    ElseIf CInt(nextHandle) <> 0 Then
                        SendMessage(nextHandle, msg.Msg, msg.WParam, msg.LParam)
                    End If
                    Exit Select

            End Select
            MyBase.WndProc(msg)
        End Sub

    End Class

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

    Public Shared JPath As CJPath

    Public Class CJPath

        '☆開発環境でのファイルのパス☆
        Private Const CURRENT0 = "c:\MAPLEO\"

        Public CURRENT As String

        '相対パス
        Public KDATA0 As String

        '絶対パス
        Dim m_Path_KDATA As String

        Public ReadOnly Property KDATA() As String
            Get
                Return m_Path_KDATA
            End Get
        End Property

        Public ReadOnly Property PERSONAL() As String
            Get
                Return KDATA + "PERSONAL.INI"
            End Get
        End Property

        Public ReadOnly Property MyNum_Connect() As String
            Get
                Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Me.KDATA + "MyNumClip.mdb;Jet OLEDB:Database Password=4300365;Persist Security Info=False"
            End Get
        End Property

        'Public MyNum_Connect As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Me.KDATA + "MyNumClip.mdb;Jet OLEDB:Database Password=4300365;Persist Security Info=False"

        Public Sub Read()
            '機能   :グロ－バル変数(PATH_?????)にファイル名、ディレクトリ－名をセットする

            '/// カレントドライブ設定 ///
            Dim Fi As New FileInfo("INI\MPROG.INI")
            If Fi.Exists Then
                CURRENT = CurDir() & "\"
            Else
                CURRENT = CURRENT0
            End If

            If Environment.CommandLine.Count = 1 Then
                KDATA0 = Environment.GetCommandLineArgs(1)
            End If
            If KDATA0 = "" Then
                KDATA0 = "KDATA\"
            ElseIf Right_(KDATA0, 1) <> "\" Then
                KDATA0 = KDATA0 + "\"
            End If
            m_Path_KDATA = CURRENT + KDATA0

        End Sub

    End Class

    Public Shared Personal As CPersonal

    ''' <summary>
    ''' ユ－ザ-情報格納構造体
    ''' </summary>
    ''' <remarks></remarks>
    Class CPersonal

        Public J_Yago As String
        Public MedicalInstitutionCode As String    '医療機関等コード 10桁
        Public LoginID As String                   'ログインID 8桁

        ''' <summary>
        ''' CPersonalにユーザーデーターを読み込む
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Read()

            Dim Stat As Integer
            Dim D() As String
            ReDim D(0 To 10)

            Do
                Stat = 0
                Read_INI(JPath.PERSONAL, "ユーザー", D, Stat)
                If Stat Then Exit Do

                Select Case D(0)

                    Case "屋号"

                        Personal.J_Yago = D(1)

                    Case "医療機関等コード"

                        Personal.MedicalInstitutionCode = D(1)

                    Case "ログインID"

                        Personal.LoginID = D(1)

                End Select

            Loop

        End Sub

    End Class

    Public Shared Sub Read_INI(ByVal F As String, ByVal Section0 As String, ByRef D() As String, ByRef ST As Integer)
        '機能   :INI ファイルの順次読み込み 指定したセクションの内容を順次呼び出すのに使用する
        '       :セクションの最後まで呼ぶとST=-1を返しファイルはクローズされる
        '       :途中で止めるとファイルがオープンされたままになるので強制クローズする必要がある
        '引渡   :F      :ファイル名
        '       :Section:セクション名
        '       :D()    :１次元配列を指定する 添え字は０から
        '       :ST     :=0 順次読み込み
        '               :=1 強制クローズ
        '戻り   :F      :保存
        '       :Section:保存
        '       :D()    :読み込まれたINIファイルの中身      ST=1のときは保存
        '       :ST     :=0 正常 =-1 読み込めるデーターなし

        'Static File_Section(1 To 5) As String
        'Static FNo(1 To 5) As Integer
        Static File_Section(0 To 5) As String
        Static FNo(0 To 5) As Integer
        Static F_Pointer As Integer
        Dim I As Integer
        Dim DD As String

        '/// ST値異常 ///
        If ST <> 0 And ST <> 1 Then ST = -1 : Exit Sub

        '/// ファイルポインタ-検索 ///
        F = Trim$(F)
        Section0 = Trim$(Section0)
        If Left$(Section0, 1) = "[" Or Left$(Section0, 1) = "［" Then Section0 = Mid$(Section0, 2)
        If Right$(Section0, 1) = "]" Or Left$(Section0, 1) = "］" Then Section0 = Left$(Section0, Len(Section0) - 1)

        If F = "" Or Section0 = "" Then Throw New Exception()

        For F_Pointer = 1 To 5
            If F + "[" + Section0 + "]" = File_Section(F_Pointer) Then Exit For
        Next F_Pointer

        '/// 強制クローズ ///
        If ST = 1 Then
            If F_Pointer > 5 Then Throw New Exception()
            FileClose(FNo(F_Pointer)) : File_Section(F_Pointer) = "" : FNo(F_Pointer) = 0
            ST = 0 : Exit Sub
        End If

        '/// １つめ読み込み ///
        If F_Pointer > 5 Then
            For F_Pointer = 1 To 5
                If FNo(F_Pointer) = 0 Then

                    File_Section(F_Pointer) = F + "[" + Section0 + "]"
                    FNo(F_Pointer) = FreeFile()

                    FileOpen(FNo(F_Pointer), F, OpenMode.Input)

                    Do
                        If EOF(FNo(F_Pointer)) Then
                            FileClose(FNo(F_Pointer))
                            File_Section(F_Pointer) = ""
                            FNo(F_Pointer) = 0
                            ST = -1
                            Exit Sub
                        End If

                        DD = LineInput(FNo(F_Pointer))
                        I = InStr(DD, ";")
                        If I Then DD = Left$(DD, I - 1)

                        I = InStr(DD, "[")
                        If I = 0 Then I = InStr(DD, "［")
                        If I > 0 And Len(DD) > I Then
                            DD = Mid$(DD, I + 1)
                            I = InStr(DD, "]")
                            If I = 0 Then I = InStr(DD, "］")
                            If I > 1 Then
                                DD = Mid$(DD, 1, I - 1)
                                If Trim$(DD) = Trim$(Section0) Then Exit Do
                            End If
                        End If
                    Loop
                    Exit For
                End If
            Next F_Pointer
            If F_Pointer > 5 Then Throw New Exception()
        End If

        '/// ２つめ以降読み込み ///
        Do

            If EOF(FNo(F_Pointer)) Then
                FileClose(FNo(F_Pointer))
                File_Section(F_Pointer) = ""
                FNo(F_Pointer) = 0
                ST = -1
                Exit Sub
            End If

            DD = LineInput(FNo(F_Pointer))
            I = InStr(DD, ";")
            If I Then DD = Left$(DD, I - 1)
            DD = TrimX(DD)
            If Len(DD) > 0 Then
                If InStr(DD, "[") Or InStr(DD, "［") Then
                    FileClose(FNo(F_Pointer))
                    File_Section(F_Pointer) = ""
                    FNo(F_Pointer) = 0
                    F_Pointer = 0
                    ST = -1
                    Exit Sub
                End If
                I = InStr(DD, "=")
                If I > 0 Then Exit Do
                If I = 0 Then DD = DD + "=" : Exit Do
            End If

        Loop

        An_INI(DD, D)

        ST = 0

    End Sub

    Private Shared Sub An_INI(ByVal L As String, ByRef LL() As String)
        '機能:INI File の１行を分解 (,で区切られたもの)
        '引渡:L          :分解対象文字列
        '    :LL()       :
        '戻り:L          :不定
        '    :LL()       :分解結果
        Dim p, I As Integer
        For p = 0 To UBound(LL, 1)
            LL(p) = ""
        Next p

        p = InStr(L, "=")
        If p = 0 Then LL(0) = "" : Exit Sub
        LL(0) = TrimX(Left$(L, p - 1))
        L = Mid$(L, p + 1)
        I = 1

        Do
            If I > UBound(LL) Then ReDim Preserve LL(0 To I)
            p = InStr(L, ",")
            If p = 0 Then LL(I) = TrimX(L) : Exit Do
            LL(I) = TrimX(Left$(L, p - 1))
            L = Mid$(L, p + 1)
            I = I + 1
        Loop

    End Sub

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
        S0 = Left$(S0, I)
        Return S0

    End Function

End Class
