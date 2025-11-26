Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.IO
Imports Microsoft.Win32

Public Class CGetClipBoard

    ''☆開発環境ではTrueにしてエラーをその場で止める☆
    Public Shared DvevelopMode As Boolean = False

    Public Shared Sub Main()

        If DvevelopMode Then
            Main2()
        Else
            Try
                Main2()
            Catch ex As Exception
                'エラーメッセージを表示する
                MsgBox(ex.ToString, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical)
            End Try
        End If

    End Sub

    Public Shared Sub Main2()

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
        Private Const CURRENT0 = "c:\MAPLEP\"

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

        Public MPROG As String

        Public ReadOnly Property MyNum_Connect() As String
            Get
                Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Me.KDATA + "MyNumClip.mdb;Jet OLEDB:Database Password=4300365;Persist Security Info=False"
            End Get
        End Property

        'Public MyNum_Connect As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Me.KDATA + "MyNumClip.mdb;Jet OLEDB:Database Password=4300365;Persist Security Info=False"

        Public Sub Read()
            '機能   :グロ－バル変数(PATH_?????)にファイル名、ディレクトリ－名をセットする

            Dim StrPath As String = System.Windows.Forms.Application.StartupPath

            ChDir(StrPath)

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

            MPROG = CURRENT + "INI\MPROG.INI"

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

    'Mprog.iniの設定を保持するオブジェクト
    Public Shared JEnv As CJEnv

    Class CJEnv
        '柔整システムの設定MPROG.INI、Rprog.iniを保持するクラス

        Private m_Path As String
        Private m_Section As String

        Class CMprog_Env
            'Mprog.iniの設定を保持するための構造体

            Public Entry As String 'エントリー
            Public value As String '値
            Public Chenge_F As Boolean '変更が有ったことを示すフラグ

        End Class

        'Mprog.iniの設定を保持する配列
        Private Mprog_Env() As CMprog_Env

        Public Sub New(ByVal Path As String, ByVal Section As String)

            m_Path = Path
            m_Section = Section

        End Sub

        ''' <summary>
        ''' デフォルト値の設定
        ''' </summary>
        ''' <param name="Entry"></param>
        ''' <param name="Def_Value"></param>
        ''' <remarks></remarks>
        Public Sub SetDefVal(ByVal Entry As String, ByVal Def_Value As String)

            If Def_Value Is Nothing Then Throw New Exception("デフォルト値にNothingは指定できません。")

            [Get](Entry, Def_Value)

        End Sub

        Public Function [Get](ByVal Entry As String, Optional ByVal Def_Value As String = Nothing) As String
            '機能   :Mprog.iniの設定の呼び出し

            '引渡   :Entry           :項目名

            '       :Def_Value       :デフォルト値(Mprog.iniにない場合)
            '戻り   :Entry           :保存

            '       :Def_Value       :保存

            '       :Get_Env　　     :設定値(INIファイルの仕様に拘わらず複数の値はもてないので注意)

            Dim Stat As Integer
            Dim Wk() As String
            ReDim Wk(0)
            Dim I As Integer

            Static Readed As Boolean
            If Readed = False Then

                I = 0

                Do

                    Stat = 0
                    Read_INI(m_Path, m_Section, Wk, Stat)

                    If Stat Then Readed = True : Exit Do

                    ReDim Preserve Mprog_Env(0 To I)
                    Mprog_Env(I) = New CMprog_Env
                    Mprog_Env(I).Entry = Wk(0)
                    Mprog_Env(I).value = Wk(1)

                    I = I + 1

                Loop

            End If

            For I = 0 To UBound(Mprog_Env)
                If Mprog_Env(I).Entry = Entry Then
                    Exit For
                End If
            Next I

            If UBound(Mprog_Env) < I Then
                If Def_Value Is Nothing Then Throw New Exception("Get_Env()でデフォルト値が決まっていません。")
                ReDim Preserve Mprog_Env(0 To I)
                Mprog_Env(I) = New CMprog_Env
                Mprog_Env(I).Entry = Entry
                Mprog_Env(I).value = Def_Value
            End If

            [Get] = Mprog_Env(I).value

        End Function

        Public Function Get_B(ByVal Entry As String, Optional ByVal Def_Value As Boolean = False) As Boolean

            Dim N As String
            Dim F As String

            If Def_Value = True Then
                N = "1"
            Else
                N = "0"
            End If

            F = [Get](Entry, N)
            Return Trim(F) = "1"

        End Function

        Public Sub Put(ByVal Entry As String, ByVal Value As String, Optional ByVal Save_F As Integer = 0)
            '機能   :Mprog.iniの設定の書き込み
            '       :Mprog_Env()を書き換えて、指定されたエントリーのみMprog.iniにその都度書き込む
            '       :Mprog.INIになければ追加
            '引渡   :Entry           :項目名

            '       :Value           :値
            '       :Save_F          :0:ファイルへ即書き込む 1:ファイルへの書き込みはしない

            '戻り   :Entry           :保存

            '       :Value           :保存

            '       :Save_F          :保存


            Dim Stat As Integer
            Dim Wk() As String
            ReDim Wk(0 To 2)
            Dim I As Integer

            For I = 0 To UBound(Mprog_Env)
                If Mprog_Env(I).Entry = Entry Then
                    Mprog_Env(I).value = Value
                    Mprog_Env(I).Chenge_F = True
                    Exit For
                End If
            Next I

            If UBound(Mprog_Env) < I Then Throw New Exception()

            If Save_F Then Exit Sub

            Wk(0) = Entry
            Wk(1) = Value

            Stat = 1
            EditIni(m_Path, m_Section, Wk, 1, Stat)

        End Sub

        Public Sub Put(ByVal Entry As String, ByVal Value As Boolean, Optional ByVal Save_F As Integer = 0)

            Dim N As String

            If Value = True Then
                N = "1"
            Else
                N = "0"
            End If

            Put(Entry, N, Save_F)

        End Sub

        Public Sub Save()
            '機能   :Mprog.iniの設定のファイルへの書き込み

            Dim Stat As Integer
            Dim Wk() As String
            ReDim Wk(0 To 1)
            Dim I As Integer
            Dim U As Integer

            U = -1
            On Error Resume Next
            U = UBound(Mprog_Env)
            On Error GoTo 0

            For I = 0 To U
                If Mprog_Env(I).Chenge_F = True Then
                    Wk(0) = Mprog_Env(I).Entry
                    Wk(1) = Mprog_Env(I).value
                    Stat = 1
                    EditIni(m_Path, m_Section, Wk, 1, Stat)
                End If
            Next I

        End Sub

    End Class

    Public Shared Sub EditIni(ByVal F As String, ByVal Sec As String, ByVal D() As String, ByVal DMax As Integer, ByRef ST As Integer)
        'Sub EditIni(F As String, Sec As String, D() As String, DMax As Integer, ST As Integer)
        '機能   iniファイルに項目を追加、または既存の項目を削除する
        '引数   F       ファイル名
        '       Sec     セクション名
        '       D()     書き込む内容　D(0) = D(1),D(2),D(3)...
        '       DMax    D()の添え字の最大値
        '       st      処理モード  0..項目がなければ追加する。既に項目があるときは追加も上書きもしない。
        '                           1..項目がなければ追加する。既に項目があれば上書きする。
        '                           9..削除する。
        '戻値   st      -1..エラー　他..正常終了
        '
        'EX)
        'D(0)="A":D(1)="10",D(2)="20"
        'EditIni "d.ini" ,"[設定]", D(), 2, 1
        '
        '[設定]
        '"A = 10,20"

        Dim I As Integer      '添え字
        Dim J As Integer      '添え字
        Dim K As Integer      '添え字
        Dim FNo As Integer      'ファイル番号
        Dim SS As String       'INIファイル処理用
        Dim LMax As Integer      'INIファイル処理用
        Dim L() As String       'INIファイル処理用
        Dim WrtFlg As Boolean      'True..該当項目書込済　False..未書込
        Dim SecFlg As Boolean      'True..該当セクションを処理中　False..該当セクションでない
        Dim EdtStr As String       '書込文字列作成用
        '
        Dim FTbl() As String       '更新ファイルテーブル
        Dim FTCnt As Integer      '更新ファイルテーブル件数

        Dim F1 As String       '一時ファイル(.BAK)

        FTCnt = 0
        ReDim FTbl(FTCnt)

        LMax = 0
        ReDim L(LMax)
        SecFlg = False
        WrtFlg = False

        '書き込みする文字列作成
        If DMax > 0 Then
            EdtStr = D(0) + " = "
            If DMax > 1 Then
                For I = 1 To (DMax - 1)
                    EdtStr = EdtStr + D(I) + ","
                Next I
            End If
            EdtStr = EdtStr + D(DMax)
        Else
            ST = -1
            Exit Sub
        End If

        If Trim(Sec) = "" Then
            'エラー終了　セクション名なし
            ST = -1
            Exit Sub
        End If

        'セクション名に[]を付ける
        If InStr(Sec, "[") <= 0 Then
            Sec = "[" + Trim(Sec)
        End If
        If InStr(Sec, "]") <= 0 Then
            Sec = Trim(Sec) + "]"
        End If

        'INIファイル
        FNo = FreeFile()
        FileOpen(FNo, F, OpenMode.Input)

        'INIファイル読込　ループ
        While EOF(FNo) = False
            SS = ""
            SS = LineInput(FNo)
            SepStr(SS, L, LMax)
            If InStr(L(0), "[") > 0 And InStr(L(0), "]") > 0 And Left$(L(0), 1) <> ";" Then
                'セクション検出
                If TrimX(L(0)) = TrimX(Sec) Then
                    '該当セクション検出
                    SecFlg = True
                Else
                    If SecFlg = True Then
                        '項目が検出されないまま、該当セクションが終了
                        '項目を追加
                        If ST <> 9 Then
                            FTCnt = FTCnt + 1
                            ReDim Preserve FTbl(FTCnt)
                            FTbl(FTCnt) = EdtStr
                        End If
                        SecFlg = False
                        WrtFlg = True
                    End If
                End If
            End If
            If SecFlg = True And TrimX(L(0)) = TrimX(D(0)) Then
                '該当項目を検出
                SecFlg = False
                WrtFlg = True
                Select Case ST
                    Case 0  '追加・上書きしない
                        'SS = SS　*読み込んだ内容のまま*
                    Case 1  '追加・上書きする
                        SS = EdtStr
                    Case 9  '削除
                        SS = "dElEtE"
                End Select
            End If
            If SS <> "dElEtE" Then
                'ST = 9 以外を書き込む
                FTCnt = FTCnt + 1
                ReDim Preserve FTbl(FTCnt)
                FTbl(FTCnt) = SS
            End If
            '    WrtFlg = True
        End While
        FileClose(FNo)

        '書込が行われないまま、INIファイル読込終了したときの処理
        If WrtFlg = False Then
            Select Case ST
                Case 0  '追加・上書きしない
                    If SecFlg = True Then
                        '項目を追加
                        FTCnt = FTCnt + 1
                        ReDim Preserve FTbl(FTCnt)
                        FTbl(FTCnt) = EdtStr
                    Else
                        'セクション、項目を追加
                        FTCnt = FTCnt + 1
                        ReDim Preserve FTbl(FTCnt)
                        FTbl(FTCnt) = ""
                        FTCnt = FTCnt + 1
                        ReDim Preserve FTbl(FTCnt)
                        FTbl(FTCnt) = Sec
                        FTCnt = FTCnt + 1
                        ReDim Preserve FTbl(FTCnt)
                        FTbl(FTCnt) = EdtStr
                    End If
                Case 1  '追加・上書きする
                    If SecFlg = True Then
                        '項目を追加
                        FTCnt = FTCnt + 1
                        ReDim Preserve FTbl(FTCnt)
                        FTbl(FTCnt) = EdtStr
                    Else
                        'セクション、項目を追加
                        FTCnt = FTCnt + 1
                        ReDim Preserve FTbl(FTCnt)
                        FTbl(FTCnt) = ""
                        FTCnt = FTCnt + 1
                        ReDim Preserve FTbl(FTCnt)
                        FTbl(FTCnt) = Sec
                        FTCnt = FTCnt + 1
                        ReDim Preserve FTbl(FTCnt)
                        FTbl(FTCnt) = EdtStr
                    End If
                Case 9  '削除
            End Select
        End If

        'INIファイルを更新する
        'For I = 1 To 500 : DoEvents() : Next I

        If FTCnt > 0 Then

            I = InStr(F, ".")
            F1 = Left$(F, I) + "BAK"
            FNo = FreeFile()
            FileOpen(FNo, F1, OpenMode.Output)
            For I = 1 To FTCnt
                PrintLine(FNo, FTbl(I))
            Next I
            FileClose(FNo)

            Kill(F)
            Dim FI As New System.IO.FileInfo(F)
            My.Computer.FileSystem.RenameFile(F1, FI.Name)
            'Name F1 As F

        End If

        'For I = 1 To 500 : DoEvents() : Next I

    End Sub

    Public Shared Sub SepStr(ByVal S As String, ByRef D() As String, ByRef DMax As Integer)
        '機能   文字列を "," , "=" で分割し、D()に納める

        '       先頭が";"の場合はコメントと見なし、

        '       分割せずにD(0)に納める

        '       文字列が空、もしくはスペースの場合も
        '       D(0)に納める

        '引数
        '       S   分割したい文字列
        '       D()
        '       DMax
        '戻値
        '       S   保存

        '       D() 分割された文字列
        '       DMax 分割された数

        Dim I As Integer
        Dim SS As String
        Dim S1 As String

        ReDim D(0)
        DMax = 0

        '空白、またはコメント

        If Trim(S) = "" Then
            D(0) = S
            Exit Sub
        End If
        If Left$(Trim(S), 1) = ";" Then
            D(0) = S
            Exit Sub
        End If

        '文字列を分割
        SS = ""
        For I = 1 To Len(S)
            S1 = Mid$(S, I, 1)
            If S1 = "," Or S1 = "=" Then
                D(DMax) = SS
                DMax = DMax + 1
                ReDim Preserve D(DMax)
                SS = ""
            Else
                SS = SS + S1
            End If
        Next I
        D(DMax) = SS

    End Sub

End Class
