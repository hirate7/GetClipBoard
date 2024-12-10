Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Public Class CGetClipBoard

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

    Public Shared Path_KDATA As String = "C:\Mapleo\KDATA\"

End Class
