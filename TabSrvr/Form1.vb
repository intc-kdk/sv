'通信関連
Imports System.Net
Imports System.Net.Sockets
Imports System.Threading

'JSON関連
Imports Newtonsoft.Json
Imports System.Security.Permissions

Imports System.Web.Script.Serialization

'Imports Microsoft.Office.Core
'Imports Microsoft.Office.Interop

Public Class Form1

#Region "ローカルメンバ"

    ''' <summary>各種設定値</summary>
    Dim 設定値 As New c設定値

    ''' <summary>送信を管理するタスクリストオブジェクト</summary>
    Dim SendTaskList As cSendTaskList

    ''' <summary>受信を管理するタスクリストオブジェクト</summary>
    Dim RecieveTaskList As cRecieveTaskList

    ''' <summary>コマンド処理を管理するタスクリストオブジェクト</summary>
    Dim CommandPool As cCommandPool

    ''' <summary>受信の待機・接続を行うスレッド管理オブジェクト</summary>
    Dim Accept As cAccept

#End Region

#Region "イベント"

    ''' <summary>起動時</summary>
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            DoubleProcessCheck()    '二重起動チェック

            With 設定値
                '各種設定ファイルを読込む
                .ReadPort = CLng(ReadTextFileFirstLine(Application.StartupPath & "\ReadPort.ini"))              '受信ポート番号
                .RemotePort = CLng(ReadTextFileFirstLine(Application.StartupPath & "\RemotePort.ini"))          '接続相手ポート番号
                .ReadTimeOut = CLng(ReadTextFileFirstLine(Application.StartupPath & "\ReadTimeOut.ini"))        '受信タイムアウト
                .WriteTimeOut = CLng(ReadTextFileFirstLine(Application.StartupPath & "\WriteTimeOut.ini"))      '送信タイムアウト
                .LogPath = ReadTextFileFirstLine(Application.StartupPath & "\LogPath.ini")                      'ログ保存先
                .ConnectTimeOut = CLng(ReadTextFileFirstLine(Application.StartupPath & "\ConnectTimeOut.ini"))  '接続タイムアウト
                .MaxSendCount = CLng(ReadTextFileFirstLine(Application.StartupPath & "\MaxSendCount.ini"))      '最大接続試行回数
                .DebugMode = CBool(ReadTextFileFirstLine(Application.StartupPath & "\DebugMode.ini"))           'デバッグモード
                .DebugMenuMode = CBool(ReadTextFileFirstLine(Application.StartupPath & "\DebugMenuMode.ini"))   'デバッグメニューモード
                .ResetTejunCmd = ReadTextFileFirstLine(Application.StartupPath & "\ResetTejunCmd.ini")          '手順書リセットコマンド
                .ReportDirPath = ReadTextFileFirstLine(Application.StartupPath & "\ReportDirPath.ini")          '作業レポートファイルのあるディレクトリのパス
                .ReportFileName = ReadTextFileFirstLine(Application.StartupPath & "\ReportFileName.ini")        '作業レポートファイルのファイル名
                .ReportSheetName = ReadTextFileFirstLine(Application.StartupPath & "\ReportSheetName.ini")      '作業レポートファイル(エクセル)のシート名
                .DumpFileName = ReadTextFileFirstLine(Application.StartupPath & "\DumpFileName.ini")            'Dumpファイル名
                .DumpCmd = ReadTextFileFirstLine(Application.StartupPath & "\DumpCmd.ini")                      'Dumpコマンド

                '内部設定を定める
                '.ConnectText = "Server=127.0.0.1;Port=5432;Database=kdk_ban;Encoding=UNICODE;User Id=postgres;Password=y3pevd36;"   'PostgreSQLへの接続文字列。
                .ConnectText = "Server=127.0.0.1;Port=5432;Database=kdk_ban2;Encoding=UNICODE;User Id=postgres;Password=y3pevd36;"   'PostgreSQLへの接続文字列。
                .ReadEndByte = 36       '受信終了を判断するbyte値 「$」
                .ReadSeparateByte = 64  'コマンドとパラメータの境目を判断するbyte値「@」

                'タブレットの開発上の初期値を設定する
                Dim tmpTablet As cComBase.cTabletList.cTablet
                tmpTablet = New cComBase.cTabletList.cTablet : With tmpTablet : .Name = "INTC開発環境" : .Type = 10 : .LastAccess = Now : .IP = "192.168.10.30" : End With : .TabletList.AddTablet(tmpTablet)
                tmpTablet = New cComBase.cTabletList.cTablet : With tmpTablet : .Name = "川本開発環境" : .Type = 10 : .LastAccess = Now : .IP = "127.0.0.1" : End With : .TabletList.AddTablet(tmpTablet)
                tmpTablet = New cComBase.cTabletList.cTablet : With tmpTablet : .Name = "鈴木開発環境" : .Type = 10 : .LastAccess = Now : .IP = "192.168.0.30" : End With : .TabletList.AddTablet(tmpTablet)
                tmpTablet = New cComBase.cTabletList.cTablet : With tmpTablet : .Name = "INTC開発環境" : .Type = 20 : .LastAccess = Now : .IP = "192.168.10.40" : End With : .TabletList.AddTablet(tmpTablet)
                tmpTablet = New cComBase.cTabletList.cTablet : With tmpTablet : .Name = "川本開発環境" : .Type = 20 : .LastAccess = Now : .IP = "127.0.0.1" : End With : .TabletList.AddTablet(tmpTablet)
                tmpTablet = New cComBase.cTabletList.cTablet : With tmpTablet : .Name = "鈴木開発環境" : .Type = 20 : .LastAccess = Now : .IP = "192.168.0.40" : End With : .TabletList.AddTablet(tmpTablet)

                'デバッグメニューモードを各メニューに反映する
                'ログに追記
                TSMI_MM_編集.Visible = .DebugMenuMode
                TSMI_TTM_編集.Visible = .DebugMenuMode
                TSS_TTM_編集.Visible = .DebugMenuMode
                'DBのバックアップ
                TSMI_MM_DBのバックアップ.Visible = .DebugMenuMode
                TSMI_TTM_DBのバックアップ.Visible = .DebugMenuMode

            End With

            '設定値を画面に表示する
            DrawConfig()

            '送受信関連オブジェクトのインスタンスを作成
            SendTaskList = New cSendTaskList(設定値, RecieveTaskList)
            CommandPool = New cCommandPool(設定値, SendTaskList)
            RecieveTaskList = New cRecieveTaskList(設定値, CommandPool)
            Accept = New cAccept(設定値, RecieveTaskList)
            Accept.StartAcceptLoop()

            ErrMsgBox(Nothing, 設定値.DebugMode, "サーバプログラムを起動しました。(設定値:" & 設定値.toString & ")", 0, 設定値.LogPath)

            Return
        Catch ex As Exception
            ErrMsgBox(ex, 設定値.DebugMode, "起動時に例外が発生しました。", 2)
            MsgBox("起動時に例外エラーが発生しました。終了します。" & ex.Message)
            Application.Exit()  '終了
        End Try
    End Sub

    ''' <summary>終了時</summary>
    Private Sub TSML_終了_Click(sender As Object, e As EventArgs) Handles TSMI_MM_終了.Click, TSMI_TTM_終了.Click
        Try
            Application.Exit()  '終了
        Catch ex As Exception
            ErrMsgBox(ex, 設定値.DebugMode, "終了時に例外が発生しました。", 2, 設定値.LogPath)
            Application.Exit()  '終了
        End Try
    End Sub

    ''' <summary>ウインドウメッセージ処理</summary>
    ''' <remarks>
    '''     主にフォームを閉じる、最小化時に
    '''     終了させずにフォームを非表示化する。
    '''     （タスクトレイアプリの挙動実現の処理）
    ''' </remarks>
    <SecurityPermission(SecurityAction.Demand, Flags:=SecurityPermissionFlag.UnmanagedCode)> _
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        Try
            Const WM_SYSCOMMAND As Integer = &H112
            Const SC_CLOSE As Long = &HF060L
            Const WM_SIZE As Integer = &H5
            Const SIZE_RESTORED As Integer = &H0
            Const SIZE_MINIMIZED As Integer = &H1
            Const SIZE_MAXIMIZED As Integer = &H2

            Select Case m.Msg
                Case WM_SYSCOMMAND
                    Select Case m.WParam.ToInt32()
                        Case SC_CLOSE
                            Me.Hide()
                            Me.WindowState = FormWindowState.Minimized
                            Return
                    End Select
                Case WM_SIZE
                    Select Case m.WParam.ToInt32()
                        Case SIZE_MINIMIZED
                            Me.Hide()
                            Me.WindowState = FormWindowState.Minimized
                            Return
                    End Select
            End Select

            MyBase.WndProc(m)
        Catch ex As Exception
            ErrMsgBox(ex, 設定値.DebugMode, "ウインドウ最小化時に例外が発生しました。", 2, 設定値.LogPath)
        End Try
    End Sub

    ''' <summary>「ウインドウを開く」実行時</summary>
    Private Sub TSMI_ウインドウを開く_Click(sender As Object, e As EventArgs) Handles TSMI_TTM_ウインドウを開く.Click, NotifyIcon1.DoubleClick
        Try
            Me.Visible = True
            Me.WindowState = FormWindowState.Normal
        Catch ex As Exception
            ErrMsgBox(ex, 設定値.DebugMode, "ウインドウ可視化時に例外が発生しました。", 2, 設定値.LogPath)
        End Try
    End Sub

    ''' <summary>「手順書をリセット」実行時</summary>
    Private Sub TSMI_手順書をリセット_Click(sender As Object, e As EventArgs) Handles TSMI_TTM_手順書をリセット.Click, TSMI_MM_手順書をリセット.Click
        Try
            If MsgBoxResult.No = MsgBox("手順書をリセットしますか？" & "リセットするとこれまでのタブレット操作履歴が失われ、最初から操作する事になります。", MsgBoxStyle.YesNo, Me.Text) Then Return

            '手順書リセットを行う。コンソールを開いて、閉じない。
            Dim psInfo As New ProcessStartInfo()
            psInfo.FileName = "cmd.exe"
            psInfo.Arguments = 設定値.ResetTejunCmd '"/K c:\php\php.exe c:\ban\reset_cmd.php"
            Process.Start(psInfo)
            設定値.LastCommandList.AllRemove()    '直近のコマンドを全て削除(タブレット側の再起動時の変な挙動への対処)
            ErrMsgBox(Nothing, 設定値.DebugMode, "コマンドラインで手順書のリセットをかけました。", 0, 設定値.LogPath)
            Return
        Catch ex As Exception
            ErrMsgBox(ex, 設定値.DebugMode, "手順書のリセットで例外が発生しました。", 2, 設定値.LogPath)
            Return
        End Try
    End Sub

    ''' <summary>「DBのバックアップ」(アイコンメニュー)実行時</summary>
    Private Sub TSMI_DBのバックアップ_Click(sender As Object, e As EventArgs) Handles TSMI_TTM_DBのバックアップ.Click, TSMI_MM_DBのバックアップ.Click
        Dim SFDialog As SaveFileDialog
        Dim tmpComBase As cComBase
        Try
            '出力先のファイル名を得る
            SFDialog = New SaveFileDialog()
            With SFDialog
                .FileName = "新しいファイル.sql"                 'ファイル名
                '.InitialDirectory = "C:\"                       'ディレクトリ
                .Filter = "SQLファイル(*.sql)|*.sql|テキストファイル(*.txt)|*.txt|すべてのファイル(*.*)|*.*"   'ファイルの種類
                .FilterIndex = 1                                 '最初に選択されているファイルの種類
                .Title = "保存先のファイルを選択してください"    'ダイアログのタイトル
                .RestoreDirectory = True                         'ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
                .OverwritePrompt = True                          '既に存在するファイル名を指定したとき警告する
                .CheckPathExists = True                          '存在しないパスが指定されたとき警告を表示する
                If Not (.ShowDialog() = DialogResult.OK) Then Return 'OKが押されなかったときは抜ける
            End With
            tmpComBase = New cComBase
            tmpComBase.SetPrivateProperty("", 設定値, Nothing, Nothing)
            tmpComBase.DoneDump(設定値.DumpCmd, SFDialog.FileName)
            Return
        Catch ex As Exception
            ErrMsgBox(ex, 設定値.DebugMode, "DBのバックアップ(GUI)で例外が発生しました。", 2, 設定値.LogPath)
        Finally
            '解放する
            If Not (SFDialog Is Nothing) Then SFDialog.Dispose()
            SFDialog = Nothing
            tmpComBase = Nothing
        End Try
    End Sub

    ''' <summary>「ログに追記」(メインメニュー)実行時</summary>
    Private Sub TSMI_ログに追記_Click(sender As Object, e As EventArgs) Handles TSMI_MM_ログに追記.Click, TSMI_TTM_ログに追記.Click
        Try
            Dim InputText As String = InputBox("ログに記録する文字列を入力して下さい。", Me.Text, "ユーザ追記:")
            If InputText Is Environ(255) Then Return 'キャンセル時は抜ける
            ErrMsgBox(Nothing, 設定値.DebugMode, InputText, 0, 設定値.LogPath)    'ログに記述する
            Return
        Catch ex As Exception
            ErrMsgBox(ex, 設定値.DebugMode, "「ログに追記」で例外が発生しました。", 2, 設定値.LogPath)
        End Try
    End Sub

    ''' <summary>発呼実験</summary>
    Private Sub 発呼実験ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 発呼実験ToolStripMenuItem1.Click
        Try
            Dim SocketState As New cSocketState(設定値)
            Dim Communicate As New cCommunicateEventArgs()
            With Communicate
                .WriteCommand = 99
                .WriteParam = Now.ToString
            End With

            SendTaskList.AddSend(SocketState, Communicate, True, "127.0.0.1", 1235)
        Catch ex As Exception
            ErrMsgBox(ex, 設定値.DebugMode, "発呼実験時に例外が発生しました。", 2, 設定値.LogPath)
        End Try
    End Sub

    Private Sub 雑ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 雑ToolStripMenuItem.Click
        Try
            Return
        Catch ex As Exception
            Return
        End Try
    End Sub

#End Region

#Region "その他関数"

    ''' <summary>設定を画面に表示する</summary>
    Private Sub DrawConfig()
        Try
            TB_設定.Text = 設定値.toString
            Return
        Catch ex As Exception
            ErrMsgBox(ex, 設定値.DebugMode, "設定の画面表示に失敗しました。", 2, 設定値.LogPath)
        End Try
    End Sub

#End Region

End Class

#Region "通信関連クラス"

#Region "送受信関連クラス"

''' <summary>サーバ発呼の実行・管理を行うクラス</summary>
''' <remarks>インスタンス作成時に必要な情報を登録する。Doneメソッドで送信が開始され、受信まで行う。</remarks>
Public Class cServerStart

#Region "内部メンバ"

    ''' <summary>各種設定値</summary>
    Dim _設定値 As New c設定値

    ''' <summary>通信相手のIPアドレス</summary>
    Dim _RemoteIP As String

    ''' <summary>通信相手のポート</summary>
    Dim _RemotePort As String

#End Region

#Region "プロパティ"

    ''' <summary>通信データ管理クラスのインスタンス</summary>
    Public Property CommunicateEventArgs As cCommunicateEventArgs

    ''' <summary>処理が成功したかどうか</summary>
    Private _Success As Boolean = False
    Public ReadOnly Property Success As Boolean
        Get
            Return _Success
        End Get
    End Property

#End Region

#Region "イベント"

    ''' <summary>コンストラクタ</summary>
    Public Sub New(設定値 As c設定値, ByRef _CommunicateEventArgs As cCommunicateEventArgs, ByRef RemoteIP As String, ByRef RemotePort As Integer)
        Try
            _設定値 = 設定値
            CommunicateEventArgs = _CommunicateEventArgs
            _RemoteIP = RemoteIP
            _RemotePort = RemotePort
            Return
        Catch ex As Exception
            ErrMsgBox(ex, 設定値.DebugMode, "サーバ発呼クラス(cServerStart)のコンストラクタで例外が発生しました。", 2, 設定値.LogPath)
            Return
        End Try
    End Sub

#End Region

#Region "パブリックメソッド"

    ''' <summary>サーバ発呼の一連の処理</summary>
    ''' <remarks></remarks>
    Public Sub Done()
        Try
            _Success = False
            Dim SocketState As New cSocketState(_設定値)

            Dim insSend As New cSend(_設定値, SocketState, CommunicateEventArgs, True, Nothing, _RemoteIP, _RemotePort)
            insSend.Done()      '送信を行う
            If Not (insSend.Success) Then Return '送信に失敗した時は抜ける
            Dim insRecieve As New cRecieve(_設定値, SocketState, CommunicateEventArgs, True, Nothing)
            insRecieve.Done()   '受信を行う
            If Not (insSend.Success) Then Return '受信に失敗した時は抜ける
            _Success = True '通信成功フラグを立てる
            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "サーバ発呼処理で例外が発生しました。", 2, _設定値.LogPath)
            Return
        End Try
    End Sub

#End Region

End Class


''' <summary>クライアント発呼時の返信を行うクラス</summary>
''' <remarks>返信のみを行う。主にコマンド処理から呼び出される。</remarks>
Public Class cClientStart

#Region "内部メンバ"

    ''' <summary>各種設定値</summary>
    Dim _設定値 As New c設定値

#End Region

#Region "プロパティ"

    ''' <summary>通信データ管理クラスのインスタンス</summary>
    Public Property CommunicateEventArgs As cCommunicateEventArgs

    ''' <summary>ソケット管理クラスのインスタンス</summary>
    Public Property SocketState As cSocketState

#End Region

#Region "イベント"

    ''' <summary>コンストラクタ</summary>
    Public Sub New(設定値 As c設定値, ByRef _SocketState As cSocketState, ByRef _CommunicateEventArgs As cCommunicateEventArgs)
        Try
            _設定値 = 設定値
            SocketState = _SocketState
            CommunicateEventArgs = _CommunicateEventArgs
            Return
        Catch ex As Exception
            ErrMsgBox(ex, 設定値.DebugMode, "クライアント発呼クラス(cClientStart)のコンストラクタで例外が発生しました。", 2, 設定値.LogPath)
            Return
        End Try
    End Sub

#End Region

#Region "パブリックメソッド"

    ''' <summary>クライアントへの返信の一連の処理</summary>
    ''' <remarks></remarks>
    Public Sub Done()
        Try
            Dim insSend As New cSend(_設定値, SocketState, CommunicateEventArgs, False, Nothing, "", _設定値.RemotePort)
            insSend.Done()
            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "クライアント発呼処理で例外が発生しました。", 2, _設定値.LogPath)
            Return
        End Try
    End Sub

#End Region

End Class

#End Region

#Region "受信関連クラス"

'''<summary>受信処理の実体のクラス</summary>
Public Class cRecieve

#Region "内部変数メンバ"

    '''<summary>上位から渡されるグローバル設定値</summary>
    Private _設定値 As c設定値

    ''' <summary>ソケット関連引数オブジェクト</summary>
    Private _SocketState As cSocketState

    ''' <summary>コマンド情報関連引数オブジェクト</summary>
    Private _Command As cCommunicateEventArgs

    ''' <summary>この通信がサーバー発呼かどうか</summary>
    Private _IsServerStart As Boolean

    ''' <summary>受信後にデータを送るコマンドタスクリストのオブジェクト</summary>
    Private _CommandTaskList As cCommandPool

    ''' <summary>受信完了まで受信スレッド実行をブロックする</summary>
    Private Shared RecieveDone As New ManualResetEvent(False)

    ''' <summary>受信が成功したかどうかのフラグ</summary>
    Private Shared RecieveSuccess As Boolean = False

    ''' <summary>受信データ量。直近のコールバックの値</summary>
    Private Shared RecieveByte_CollBack As Long

    ''' <summary>受信した文字列</summary>
    Private Shared RecieveText As String = ""

    ''' <summary>受信したデータ</summary>
    Private Shared RecieveBytes As Byte()

#End Region

#Region "プロパティ"

    ''' <summary>処理が完了したかどうか</summary>
    Private _Complete As Boolean = False
    Public ReadOnly Property Complete As Boolean
        Get
            Return _Complete
        End Get
    End Property

    ''' <summary>処理が成功したかどうか</summary>
    Private _Success As Boolean = False
    Public ReadOnly Property Success As Boolean
        Get
            Return _Success
        End Get
    End Property

    ''' <summary>処理開始の日時</summary>
    Private _StartDateTime As Date = Date.MinValue
    Public ReadOnly Property StartDateTime As Date
        Get
            Return _StartDateTime
        End Get
    End Property

    ''' <summary>処理終了の日時</summary>
    Private _EndDateTime As Date = Date.MaxValue
    Public ReadOnly Property EndDateTime As Date
        Get
            Return _EndDateTime
        End Get
    End Property

    ''' <summary>スレッドプールに追加された日時</summary>
    Public AddDateTime As Date = Date.MaxValue

#End Region

#Region "イベント"

    ''' <summary                     >コンストラクタ                                          </summary>
    ''' <param name="Arg設定値"      >グローバル設定値                                        </param  >
    ''' <param name="SocketState"    >受信に用いるソケット関連引数                            </param  >
    ''' <param name="Command"        >コマンド処理に用いるコマンド関連引数                    </param  >
    ''' <param name="ShutDown"       >受信後にシャットダウンを行うかどうか                    </param  >
    ''' <param name="CommandTaskList">受信後に受信データを送るコマンドタスクリストオブジェクト</param  >
    Public Sub New(Arg設定値 As c設定値, ByRef SocketState As cSocketState, ByRef Command As cCommunicateEventArgs, IsServerStart As Boolean, ByRef CommandTaskList As cCommandPool)
        Try
            _設定値 = Arg設定値           '設定値
            _SocketState = SocketState    '
            _Command = Command
            _IsServerStart = IsServerStart
            _CommandTaskList = CommandTaskList
            RecieveDone = New ManualResetEvent(False)
            Return
        Catch ex As Exception
            ErrMsgBox(ex, Arg設定値.DebugMode, "受信処理クラス(cRecieve)のコンストラクタで例外が発生しました。", 2, Arg設定値.LogPath)
            Return
        End Try
    End Sub

    ''' <summary>デストラクタ</summary>
    Protected Overrides Sub Finalize()
        Try
            '必要な開放処理をここに記述していたが、移設したｗ
            MyBase.Finalize()
            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "受信処理クラス(cRecieve)のデストラクタで例外が発生しました。", 2, _設定値.LogPath)
            Return
        End Try
    End Sub

#End Region

#Region "受信関連関数"

    ''' <summary                 >データ受信を行う関数。            </summary>
    ''' <remarks                 >受信、コマンド判定、ログ出力を行う</remarks>
    Public Sub Done()
        RecieveText = ""
        Dim ErrMessage As String = ""
        Dim LogReadText As String = "未受信"
        Try
            '処理前の初期化を行う
            _Complete = False       '処理が終わったかどうか
            _Success = False        '処理が成功したか
            _StartDateTime = Now    '処理開始日時

            '実行条件を調べる
            If _SocketState Is Nothing Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド受信に失敗しました。受信ソケットが与えられていません。(_SocketState is nothing)", 1, _設定値.LogPath)
                ErrMessage = "受信ソケットが与えられていません。(_SocketState is nothing)"
                Return
            End If
            If _SocketState.WorkSocket Is Nothing Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド受信に失敗しました。受信ソケットが与えられていません。(_SocketState.WorkSocket is nothing)", 1, _設定値.LogPath)
                ErrMessage = "受信ソケットが与えられていません。(_SocketState.WorkSocket is nothing)"
                Return
            End If
            'If Not (_SocketState.WorkSocket.Connected) Then
            If Not (IsConnecting(_SocketState.WorkSocket)) Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド受信に失敗しました。受信ソケットが未接続です。", 1, _設定値.LogPath)
                ErrMessage = "受信ソケットが未接続です。"
                Return
            End If

            '受信を行う
            If Not (TryRecieve(_SocketState)) Then
                '受信失敗時は
                ErrMsgBox(Nothing, _設定値.DebugMode, "データ受信に失敗しました。", 1, _設定値.LogPath)
                ErrMessage = "データ受信に失敗しました。"
                Return        '偽を返す
            End If
            '受信成功時はコマンド文字列として適当かどうか調べる
            If Not (ConvRecieveDataToCP(_Command)) Then
                'コマンド文字ではないときは
                ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド条件に違反しています。", 1, _設定値.LogPath)
                ErrMessage = "コマンド条件に違反しています。"
                Return         '偽を返す
            End If
            _Success = True

            'サーバ発呼時でなくクライアント発呼ならばコマンド処理タスクリストへ追加する
            If Not _IsServerStart Then
                _CommandTaskList.AddSend(_SocketState, _Command)
            End If

            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "受信処理に失敗しました。(例外)", 2, _設定値.LogPath)
            Return
        Finally
            'ログ記述
            If _Success Then ErrMessage = ""

            'ログに記述する接続相手を求める
            Dim LogRemoteText As String = "接続先不明"
            Try
                If Not (_SocketState Is Nothing) Then
                    If Not (_SocketState.WorkSocket Is Nothing) Then
                        If Not (_SocketState.WorkSocket.RemoteEndPoint Is Nothing) Then
                            LogRemoteText = _SocketState.WorkSocket.RemoteEndPoint.ToString()
                        End If
                    End If
                End If
            Catch ex As Exception
                '接続先を求める際にエラーになってもなにもせず
            End Try

            If Not ("" = RecieveText) Then LogReadText = RecieveText
            If Not ("" = ErrMessage) Then ErrMessage = "," & ErrMessage 'エラーが有る時はエラーメッセージの先頭にカンマを追加

            Dim ShutDownFlag As Boolean = False  'コネクションのシャットダウン・ソケットの解放を行うかどうかのフラグ
            If Not (Success) Then ShutDownFlag = True '通信に失敗した時は解放処理を行う
            If _IsServerStart Then ShutDownFlag = True 'サーバ発呼ならば即遮断する

            If ShutDownFlag Then    '解放処理を行う場合
                'ソケットの解放をする
                If Not (_SocketState Is Nothing) Then
                    With _SocketState
                        'If Not (.WorkSocket Is Nothing) Then If .WorkSocket.Connected Then .WorkSocket.Shutdown(SocketShutdown.Both) '通信を閉じる
                        If Not (.WorkSocket Is Nothing) Then
                            If IsConnecting(.WorkSocket) Then
                                '解放後にコールバック関数が呼ばれないようにする為、以下の方法を取る。
                                .WorkSocket.Shutdown(SocketShutdown.Both)   '通信を閉じる
                                .WorkSocket.Disconnect(False)               'ソケット接続を閉じる。
                                ErrMsgBox(Nothing, _設定値.DebugMode, "受信後、サーバからコネクションを閉じました。リモート:" & LogRemoteText, 0, _設定値.LogPath)
                            Else
                                ErrMsgBox(Nothing, _設定値.DebugMode, "受信後、相手がコネクションを閉じました。リモート:" & LogRemoteText, 0, _設定値.LogPath)
                            End If
                        End If
                        'If Not (.WorkSocket Is Nothing) Then .WorkSocket.Close()   'これを実行するとコールバック関数が呼ばれてしまう！のでコメントアウト
                        If Not (.WorkSocket Is Nothing) Then .WorkSocket.Dispose() '開放
                        If Not (.WorkSocket Is Nothing) Then .WorkSocket = Nothing '開放
                    End With
                End If
            End If

            '他のオブジェクトを明示的に解放する
            If Not (RecieveDone Is Nothing) Then RecieveDone.Dispose()
            If Not (RecieveDone Is Nothing) Then RecieveDone = Nothing

            _EndDateTime = Now  '処理終了日時
            _Complete = True    '処理が終わったかどうか

            WriteLog(_設定値.LogPath, "input", LogRemoteText, LogReadText & ErrMessage) 'ログを出力

        End Try
    End Sub

    ''' <summary               >データ受信を試みる関数。           </summary>
    ''' <param name="SocketState">接続に使うする引数オブジェクト     </param  >
    ''' <returns               >送信の正否を返す。                 </returns>
    ''' <remarks               >ここからコールバック関数を呼び出す。</remarks>
    Private Function TryRecieve(ByRef SocketState As cSocketState) As Boolean
        Try
            '実行条件を判断する
            With SocketState
                If .WorkSocket Is Nothing Then : ErrMsgBox(Nothing, _設定値.DebugMode, "受信に失敗しました。受信ソケットインスタンスがNothingです。", 1, _設定値.LogPath) : Return False : End If 'ソケットが無い時は抜ける
                'If Not (.WorkSocket.Connected) Then : ErrMsgBox(Nothing, _設定値.DebugMode, "受信に失敗しました。受信ソケットのコネクションが接続状態ではありません。", 1, _設定値.LogPath) : Return False : End If '未接続時は抜ける
                If Not (IsConnecting(.WorkSocket)) Then : ErrMsgBox(Nothing, _設定値.DebugMode, "受信に失敗しました。受信ソケットのコネクションが接続状態ではありません。", 1, _設定値.LogPath) : Return False : End If '未接続時は抜ける
                'データが分割されて送られてくる事を考慮して受信ループを作る
                '試行回数は無効。タイムアウトだけ。
                Do
                    Dim tmpIAsyncResult As IAsyncResult
                    Try
                        RecieveSuccess = False      '受信成功フラグをリセット
                        RecieveByte_CollBack = 0    'コールバックでの受信バイト数をリセット
                        RecieveDone.Reset()         'このスレッドをブロックできるようにフラグをリセット
                        tmpIAsyncResult = SocketState.WorkSocket.BeginReceive(SocketState.ReadBuffer, 0, SocketState.ReadBuffer.Length, System.Net.Sockets.SocketFlags.None, New System.AsyncCallback(AddressOf RecieveCallback), SocketState) 'コールバック関数を指定して非同期で受信
                        RecieveDone.WaitOne(_設定値.ReadTimeOut) 'タイムアウトを設定してブロック
                    Catch ex As Exception
                        ErrMsgBox(Nothing, _設定値.DebugMode, "受信試行に失敗しました。受信ソケットのコネクションが接続状態ではありません。", 1, _設定値.LogPath)
                        RecieveSuccess = False      'エラー時は受信失敗とする
                        'SocketState.WorkSocket.EndReceive(tmpIAsyncResult) '本当はここでも受信コールバックを止めたいが応答が帰ってこない仕様なのでコメントアウト
                    End Try

                    '受信正否を判断する
                    Dim SuccessFlag As Boolean = True
                    If Not (RecieveSuccess) Then
                        SuccessFlag = False '受信コールバック処理が行われていない時・コールバック関数内での失敗受信失敗とする
                    End If

                    If SuccessFlag Then
                        '受信成功時は受信データをバッファに蓄積する
                        .AddReadAddBuffer(RecieveByte_CollBack)         '受信バッファを蓄積配列に追加する
                        If 0 < .InReadEnd(RecieveByte_CollBack) Then    '直近の受信内容に終端文字があった時は
                            '受信完了とする
                            RecieveBytes = .ReadData                                       '蓄積したデータを確保
                            Try 'UTF8への返還を試みる
                                If .ReadData Is Nothing Then    'データがNull時は
                                    RecieveText = ""   '空文字を作成
                                Else    'データがある時は
                                    RecieveText = System.Text.Encoding.UTF8.GetString(.ReadData)   '蓄積したデータを文字列に変換する
                                End If
                                ErrMsgBox(Nothing, _設定値.DebugMode, "データ受信に成功しました。(リモート：" & .WorkSocket.RemoteEndPoint.ToString & "」)", 0, _設定値.LogPath)
                                Return True '真を返す
                            Catch ex As Exception
                                '変換に失敗した時は
                                RecieveText = ""   '空文字を作成
                            End Try
                            ErrMsgBox(Nothing, _設定値.DebugMode, "受信に失敗しました。受信データのUTF-8文字列への変換に失敗しました。", 1, _設定値.LogPath)
                            Return False 'UTF-8への変換に失敗した時は偽を返す
                        Else    '終端文字が見つからなかったときは
                            'If .WorkSocket.Connected Then   'コネクションが切られていた時はループから出る
                            If IsConnecting(.WorkSocket) Then   'コネクションが切られていた時はループから出る
                                Exit Do '受信失敗時はループを抜ける
                            End If
                        End If
                        '受信が途中ならばさらに受信を続ける
                    Else
                        Exit Do '受信失敗時はループを抜ける
                    End If
                Loop
                '受信に失敗した時は
                'これまで受信した文字列をデコード
                RecieveBytes = .ReadData                                       '蓄積したデータを確保
                Try 'UTF8への返還を試みる
                    If .ReadData Is Nothing Then    'データがNull時は
                        RecieveText = ""   '空文字を作成
                    Else    'データがある時は
                        RecieveText = System.Text.Encoding.UTF8.GetString(.ReadData)   '蓄積したデータを文字列に変換する
                    End If
                Catch ex As Exception
                    '変換に失敗した時は
                    RecieveText = ""   '空文字を作成
                End Try
            End With
            ErrMsgBox(Nothing, _設定値.DebugMode, "受信に失敗しました。(受信試行文字列「" & RecieveText & "」)", 1, _設定値.LogPath)
            Return False '偽を返す
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "受信に失敗しました。(例外)", 2, _設定値.LogPath)
            Return False '偽を返す
        End Try
    End Function

    ''' <summary         >データ受信のコールバック関数          </summary>
    ''' <param name="Arg">引数コンテナクラス(cSocketState)      </param  >
    ''' <remarks         >タブレット発呼・サーバ発呼の両方で使う</remarks>
    Private Sub RecieveCallback(ByVal Arg As IAsyncResult)
        Dim Listener As cSocketState
        Try
            Listener = CType(Arg.AsyncState, cSocketState)  '引数を変換
            RecieveByte_CollBack = 0
            Try
                RecieveByte_CollBack = Listener.WorkSocket.EndReceive(Arg)  '受信を完了する
            Catch ex As Exception
                'Socketの
            End Try
            RecieveSuccess = True  '受信成功のフラグを上げる
            Return  '抜ける
        Catch ex As Exception 'エラー発生時は
            ErrMsgBox(ex, _設定値.DebugMode, "非同期受信関数内で例外が発生しました。", 2, _設定値.LogPath)
            RecieveSuccess = False '受信成功のフラグを下げる
        Finally
            If Not (RecieveDone Is Nothing) Then RecieveDone.Set() '受信完了待機を解除
        End Try
    End Sub

    ''' <summary                 >受信時の文字列処理を行う。受信データをコマンドとパラメータに分ける</summary>
    ''' <param name="CommandArgs">解析結果                                                          </param  >
    ''' <returns                 >コマンドとして有効かどうか                                        </returns>
    ''' <remarks                 >                                                                  </remarks>
    Private Function ConvRecieveDataToCP(ByRef CommandArgs As cCommunicateEventArgs) As Boolean
        Try
            Dim Source As Byte() = RecieveBytes
            Dim Dest As Boolean
            '受信データが条件を満たしているかを調べる
            If Source.Length < 4 Then : ErrMsgBox(Nothing, _設定値.DebugMode, "受信データが4Byteに足りません(" & Source.Length & "Byte)。", 1, _設定値.LogPath) : Return False : End If
            If Not (_設定値.ReadSeparateByte = Source(2)) Then : ErrMsgBox(Nothing, _設定値.DebugMode, "受信データにセパレータ「" & Chr(_設定値.ReadSeparateByte) & "」が見つかりません(" & Source(2) & ")。", 1, _設定値.LogPath) : Return False : End If '3文字目がセパレータでない時は抜ける
            If Not ((47 < Source(0) Or Source(0) < 58) Or (64 < Source(0) Or Source(0) < 91)) Then : ErrMsgBox(Nothing, _設定値.DebugMode, "受信データの1byte目が正しい範囲にありません(" & Source(0) & ")。", 1, _設定値.LogPath) : Return False : End If '1文字目が0-Zでない時は抜ける
            If Not ((47 < Source(1) Or Source(1) < 58) Or (64 < Source(1) Or Source(1) < 91)) Then : ErrMsgBox(Nothing, _設定値.DebugMode, "受信データの2byte目が正しい範囲にありません(" & Source(1) & ")。", 1, _設定値.LogPath) : Return False : End If '2文字目が0-Zでない時は抜ける
            If Not (_設定値.ReadEndByte = Source(UBound(Source))) Then : ErrMsgBox(Nothing, _設定値.DebugMode, "受信データに終端文字「" & Chr(_設定値.ReadEndByte) & "」が見つかりません(" & Source(UBound(Source)) & ")。", 1, _設定値.LogPath) : Return False : End If '最終文字が終端文字でない時は抜ける

            'コマンドとパラメータを分ける、文字列に直す
            With CommandArgs
                .OrgReadBytes = Source.Clone                                    '受信データ

                Try 'UTF8への返還を試みる
                    If Source Is Nothing Then    'データがNull時は
                        .OrgReadText = ""   '空文字を作成
                    Else    'データがある時は
                        .OrgReadText = System.Text.Encoding.UTF8.GetString(Source)      '受信文字列
                    End If
                Catch ex As Exception
                    '変換に失敗した時は
                    .OrgReadText = ""   '空文字を作成
                End Try
                .OrgReadText = System.Text.Encoding.UTF8.GetString(Source)      '受信文字列
                .ReadCommand = .OrgReadText.Substring(0, 2)                     'コマンド文字列(先頭2文字)
                .ReadParam = .OrgReadText.Substring(3, .OrgReadText.Length - 4) 'コマンド文字列(先頭3文字と終端文字以外)
            End With
            Return True
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "受信文字列の調査判定で例外が発生しました。", 2, _設定値.LogPath)
            Return False
        End Try
    End Function

#End Region

End Class

''' <summary>受信のTaskList（疑似スレッドプール）を管理するクラス</summary>
Public Class cRecieveTaskList

#Region "内部メンバ"

    '''<summary>上位から渡されるグローバル設定値</summary>
    Private _設定値 As c設定値

    ''' <summary>タスクを管理するList</summary>
    Private _RecieverList As List(Of cRecieveTask) 'タスクを管理するList

    ''' <summary>受信データを送るコマンドタスクリストオブジェクト。</summary>
    Private _CommandTaskList As cCommandPool

#End Region

#Region "イベント"

    ''' <summary                     >コンストラクタ                                          </summary>
    ''' <param name="Arg設定値"      >グローバル設定値                                        </param  >
    ''' <param name="CommandTaskList">受信後に受信データを送るコマンドタスクリストオブジェクト</param  >
    Public Sub New(Arg設定値 As c設定値, ByRef CommandTaskList As cCommandPool)
        Try
            _設定値 = Arg設定値                       '設定値
            _CommandTaskList = CommandTaskList        'コマンドタスクリスト
            _RecieverList = New List(Of cRecieveTask) 'インスタンス作成
            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "受信タスクリストのコンストラクタで例外が発生しました。", 2, _設定値.LogPath)
            Return  '抜ける
        End Try
    End Sub

    ''' <summary>デストラクタ</summary>
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#End Region

#Region "パブリックメソッド"

    ''' <summary                 >受信コネクションをタスクのListに登録して実行する</summary>
    ''' <param name="SocketState">受信に用いるソケット関連引数                    </param  >
    ''' <param name="Command"    >コマンド処理に用いるコマンド関連引数            </param  >
    ''' <param name="ShutDown"   >受信後にシャットダウンを行うかどうか            </param  >
    Public Sub AddRecieve(ByRef SocketState As cSocketState, ByRef Command As cCommunicateEventArgs, ShutDown As Boolean)
        Try
            'タスクと受信クラスを作成
            Dim AddRecieveTask As New cRecieveTask(_設定値, SocketState, Command, ShutDown, _CommandTaskList)
            _RecieverList.Add(AddRecieveTask)   'リストに登録

            'ここで処理の終わった受信インスタンスを解除する
            'For Eachではエラーになる
            For Index As Long = _RecieverList.Count - 1 To 0 Step -1
                Dim DeleteFlag As Boolean = False   '削除フラグ
                With _RecieverList(Index).TaskObj
                    If .IsCompleted Then DeleteFlag = True 'タスク側で完全終了していれば削除対象
                End With

                '削除するかどうかを調べる
                With _RecieverList(Index).TaskObj
                    'スレッドの動作で判断
                    If .IsCompleted Then DeleteFlag = True
                    If .IsFaulted Then DeleteFlag = True
                    If .IsCanceled Then DeleteFlag = True
                End With

                'タスクが終了してない時はキャンセルをかけるべきだが未実装
                'With _RecieverList(Index).Recieve
                '    '削除するかどうかを調べる
                '    If .Complete Then   '終了時
                '        If .Success Then    '成功終了時は削除
                '            DeleteFlag = True
                '        Else                '失敗で終了時は削除（再送する？）
                '            DeleteFlag = True
                '        End If
                '    Else    '未完了の時
                '    End If
                'End With

                If DeleteFlag Then  '削除対象なら削除する
                    _RecieverList.Remove(_RecieverList(Index))  'リストから削除する
                End If

            Next
            Return  '抜ける
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "受信タスクリストの登録・実行で例外が発生しました。", 2, _設定値.LogPath)
            Return  '抜ける
        End Try
    End Sub

#End Region

    ''' <summary>タスクと処理を配列で管理する為のクラス</summary>
    Private Class cRecieveTask

#Region "内部メンバ"

        ''' <summary>上位から渡される設定値</summary>
        Private _設定値 As c設定値

#End Region

#Region "プロパティ"

        ''' <summary>受信を処理するクラス</summary>
        Private _Reciever As cRecieve  '受信クラス
        Public ReadOnly Property Recieve As cRecieve
            Get
                Return _Reciever
            End Get
        End Property

        ''' <summary>受信を別スレッドで処理するタスク</summary>
        Private _Task As Task         'タスククラス
        Public ReadOnly Property TaskObj As Task
            Get
                Return _Task
            End Get
        End Property

#End Region

#Region "イベント"

        ''' <summary                     >コンストラクタ                                        </summary>
        ''' <param name="Arg設定値"      >上位から渡される設定値                                </param  >
        ''' <param name="SocketState"    >受信に用いるソケット関連引数                          </param  >
        ''' <param name="Command"        >コマンド処理に用いるコマンド関連引数                  </param  >
        ''' <param name="ShutDown"       >受信後にシャットダウンを行うかどうか                  </param  >
        ''' <param name="CommandTaskList">受信後にコマンド処理を送るコマンドプールのインスタンス</param  >
        ''' <remarks                     >インスタンス作成と同時に受信処理を開始する            </remarks>
        Public Sub New(Arg設定値 As c設定値, ByRef SocketState As cSocketState, ByRef Command As cCommunicateEventArgs, ShutDown As Boolean, ByRef CommandTaskList As cCommandPool)
            Try
                _設定値 = Arg設定値           '設定値
                _Reciever = New cRecieve(_設定値, SocketState, Command, ShutDown, CommandTaskList) '受信タスクのインスタンスを作成
                _Reciever.AddDateTime = Now                     '現在日時を登録
                _Task = New Task(AddressOf _Reciever.Done)       'タスクに処理を登録
                _Task.Start()                                    'さっさと開始
                Return
            Catch ex As Exception
                ErrMsgBox(ex, _設定値.DebugMode, "受信タスクのコンストラクタで例外が発生しました。", 2, _設定値.LogPath)
                Return
            End Try
        End Sub

        ''' <summary>デストラクタ</summary>
        Protected Overrides Sub Finalize()
            Try
                _Task.Dispose()
                MyBase.Finalize()
                Return
            Catch ex As Exception
                ErrMsgBox(ex, _設定値.DebugMode, "受信タスクのデストラクタで例外が発生しました。", 2, _設定値.LogPath)
                Return
            End Try
        End Sub

#End Region

    End Class

End Class

#End Region

#Region "送信関連クラス"

''' <summary>送信処理の実体のクラス</summary>
Public Class cSend

#Region "内部変数メンバ"

    '''<summary>上位から渡されるグローバル設定値</summary>
    Private _設定値 As c設定値

    ''' <summary>ソケット関連引数オブジェクト</summary>
    Private _SocketState As cSocketState

    ''' <summary>コマンド情報関連引数オブジェクト</summary>
    Private _Command As cCommunicateEventArgs

    ''' <summary>この通信がサーバー発呼かどうか</summary>
    Private _IsServerStart As Boolean

    ''' <summary>送信後に受信を行う受信タスクリストのオブジェクト</summary>
    Private _RecieveTaskList As cRecieveTaskList

    ''' <summary>接続相手のIPアドレス、未接続時に使用</summary>
    Private _RemoteIP As String = ""

    ''' <summary>接続相手のポート、未接続時に使用</summary>
    Private _RemotePort As Integer = 0

    ''' <summary>送信完了まで受信スレッド実行をブロックする 非同期通信用</summary>
    Private Shared SendDone As New ManualResetEvent(False)

    ''' <summary>送信が成功したかどうかのフラグ 非同期通信用</summary>
    Private Shared SendSuccess As Boolean = False

    ''' <summary>接続完了まで接続スレッド実行をブロックする。非同期通信用</summary>
    Private Shared ConnectDone As New ManualResetEvent(False)

    ''' <summary>接続が成功したかどうかのフラグ 非同期通信用</summary>
    Private Shared ConnectSuccess As Boolean = False

#End Region

#Region "プロパティ"

    ''' <summary>スレッドプールに追加された日時</summary>
    Public AddDateTime As Date = Date.MaxValue

    ''' <summary>処理が完了したかどうか</summary>
    Private _Complete As Boolean = False
    Public ReadOnly Property Complete As Boolean
        Get
            Return _Complete
        End Get
    End Property

    ''' <summary>処理が成功したかどうか</summary>
    Private _Success As Boolean = False
    Public ReadOnly Property Success As Boolean
        Get
            Return _Success
        End Get
    End Property

    ''' <summary>処理開始の日時</summary>
    Private _StartDateTime As Date = Date.MinValue
    Public ReadOnly Property StartDateTime As Date
        Get
            Return _StartDateTime
        End Get
    End Property

    ''' <summary>処理終了の日時</summary>
    Private _EndDateTime As Date = Date.MaxValue
    Public ReadOnly Property EndDateTime As Date
        Get
            Return _EndDateTime
        End Get
    End Property

#End Region

#Region "イベント"

    ''' <summary                >コンストラクタ  </summary>
    ''' <param name="Arg設定値" >グローバル設定値</param  >
    Public Sub New(Arg設定値 As c設定値, ByRef SocketState As cSocketState, ByRef Command As cCommunicateEventArgs, IsServerStart As Boolean, ByRef RecieveTaskList As cRecieveTaskList, Optional ByVal RemoteIP As String = "", Optional ByVal RemotePort As Integer = 0)
        Try
            _設定値 = Arg設定値           '設定値
            _SocketState = SocketState    '
            _Command = Command
            _IsServerStart = IsServerStart
            _RecieveTaskList = RecieveTaskList
            SendDone = New ManualResetEvent(False)
            _RemoteIP = RemoteIP
            _RemotePort = RemotePort
        Catch ex As Exception
            ErrMsgBox(ex, Arg設定値.DebugMode, "送信処理クラス(cSend)のコンストラクタで例外が発生しました。", 2, Arg設定値.LogPath)
        End Try
    End Sub

    ''' <summary>デストラクタ</summary>
    Protected Overrides Sub Finalize()
        Try
            '必要な開放処理をここに記述していたが、移設したｗ
            MyBase.Finalize()
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "送信処理クラス(cSend)のデストラクタで例外が発生しました。", 2, _設定値.LogPath)
        End Try
    End Sub

#End Region

#Region "送信関連関数"

    ''' <summary>データ送信を行う関数。            </summary>
    ''' <remarks>コマンド判定、送信、ログ出力を行う</remarks>
    Public Sub Done()
        Dim ErrMessage As String = ""
        Dim LogSendText As String = ""
        Try

            '処理前の初期化を行う
            _Complete = False       '処理が終わったかどうか
            _Success = False        '処理が成功したか
            _StartDateTime = Now    '処理開始日時

            '実行条件を調べる
            If _Command Is Nothing Then '送信情報が無い時は抜ける
                ErrMessage = "送信に失敗しました。送信情報が不明です。(_Command is nothing)"
                ErrMsgBox(Nothing, _設定値.DebugMode, ErrMessage, 1, _設定値.LogPath)
                LogSendText = ""
                Return
            End If

            '仮のログの送信文字列を作る
            LogSendText = _Command.WriteCommand & "@" & _Command.WriteParam & "$(仮)"

            'ソケットが存在しない時は用意する
            If _SocketState Is Nothing Then _SocketState = New cSocketState(_設定値)
            If _SocketState.WorkSocket Is Nothing Then
                _SocketState.WorkSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp) 'ソケットを作成する
            End If

            'コマンド文字列を作る
            If Not (ConvCPToWriteData(_Command)) Then    '失敗した時は
                ErrMessage = "送信に失敗しました。コマンド条件に違反しています。(コマンド値:" & _Command.WriteCommand & ")"
                ErrMsgBox(Nothing, _設定値.DebugMode, ErrMessage, 1, _設定値.LogPath)
                Return
            End If
            _SocketState.SendText = _Command.OrgWriteText       '送信文字列
            _SocketState.SendBuffer = _Command.OrgWriteBytes    '送信データ
            LogSendText = _SocketState.SendText                 'ログ記述用

            'If Not (_SocketState.WorkSocket.Connected) Then
            If Not (IsConnecting(_SocketState.WorkSocket)) Then
                'まだ接続状態でないなら接続を試みる
                '接続条件をチェック
                If "" = _RemoteIP Then  '相手IPが未設定の時は抜ける
                    ErrMessage = "送信に失敗しました。通信相手との接続に失敗しました(相手IP不明)。"
                    ErrMsgBox(Nothing, _設定値.DebugMode, ErrMessage, 1, _設定値.LogPath)
                    Return
                End If
                If _RemotePort < 1 Then '相手ポートが未設定の時は抜ける
                    ErrMessage = "送信に失敗しました。通信相手との接続に失敗しました(相手Port不明)。"
                    ErrMsgBox(Nothing, _設定値.DebugMode, ErrMessage, 1, _設定値.LogPath)
                    Return
                End If

                'リモートエンドポイントを作成する
                Dim RemoteEndPoint As New IPEndPoint(IPAddress.Parse(_RemoteIP), _RemotePort) '送信相手のアドレスとポートを指定する
                Dim LocalEndPoint As New IPEndPoint(System.Net.IPAddress.Any, 0)                         '受信デバイスを指定する(ポートは指定しない)←送信時は自分のポートを固定すると再接続に4分を要する仕様なので固定せず使う
                _SocketState.WorkSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp) 'ソケットを作成する
                With _SocketState.WorkSocket
                    .ReceiveTimeout = _設定値.ReadTimeOut   '受信タイムアウト
                    .SendTimeout = _設定値.WriteTimeOut     '送信タイムアウト
                    '.Bind(LocalEndPoint)                    '自分側のデバイスを指定を指定
                End With

                '送信相手と接続する
                If Not (TryConnect_同期(_SocketState.WorkSocket, RemoteEndPoint)) Then
                    '接続失敗時は
                    ErrMessage = "送信に失敗しました。通信相手との接続に失敗しました。"
                    ErrMsgBox(Nothing, _設定値.DebugMode, ErrMessage, 1, _設定値.LogPath)
                    Return              '抜ける
                End If

                ErrMsgBox(Nothing, _設定値.DebugMode, "セッション確立。送信開始。(リモート:" & RemoteEndPoint.ToString & ")", 0, _設定値.LogPath)

                '接続成功時は送信へ進む
            End If

            '送信する
            If Not (TrySend(_SocketState)) Then
                '送信失敗時は
                ErrMessage = "送信に失敗しました。データ送信に失敗しました。"
                ErrMsgBox(Nothing, _設定値.DebugMode, ErrMessage, 1, _設定値.LogPath)
                Return     '偽を返す
            End If

            ErrMsgBox(Nothing, _設定値.DebugMode, "送信完了。(リモート:" & _SocketState.WorkSocket.RemoteEndPoint.ToString & ")", 0, _設定値.LogPath)

            '送信成功時は
            _Success = True
            'If Not (_IsServerStart) Then    'サーバ発呼の通信ならば
            '   _RecieveTaskList.AddRecieve(_SocketState, _Command, True)  '受信スレッドプールへ追加する
            'End If
            Return
        Catch ex As Exception
            Dim RemoteText As String = _RemoteIP & ":" & _RemotePort
            If Not (_SocketState Is Nothing) Then
                If Not (_SocketState.WorkSocket Is Nothing) Then
                    If Not (_SocketState.WorkSocket.RemoteEndPoint Is Nothing) Then
                        RemoteText = _SocketState.WorkSocket.RemoteEndPoint.ToString()
                    End If
                End If
            End If
            ErrMessage = "送信に失敗しました。データ送信に失敗しました。(例外)" & ex.Message
            ErrMsgBox(ex, _設定値.DebugMode, ErrMessage, 2, _設定値.LogPath)
            If _設定値.DebugMode Then Throw '再度エラーを上げる
            Return
        Finally
            'ログ記述
            If _Success Then ErrMessage = ""
            'ログに記述する接続相手を求める
            Dim LogRemoteText As String = _RemoteIP & ":" & _RemotePort
            Try
                If Not (_SocketState Is Nothing) Then
                    If Not (_SocketState.WorkSocket Is Nothing) Then
                        If Not (_SocketState.WorkSocket.RemoteEndPoint Is Nothing) Then
                            LogRemoteText = _SocketState.WorkSocket.RemoteEndPoint.ToString()
                        End If
                    End If
                End If
            Catch ex As Exception
                'エラーの時は何もしない
            End Try
            If Not ("" = ErrMessage) Then ErrMessage = "," & ErrMessage 'エラーが有る時はエラーメッセージの先頭にカンマを追加

            'クライアント側が切断するので送信ではCloseしないが、タイムアウトを過ぎても接続中ならば切断しちゃう
            Try
                Dim ShutDownFlag As Boolean = False  'コネクションのシャットダウン・ソケットの解放を行うかどうかのフラグ
                If Not (Success) Then ShutDownFlag = True '通信に失敗した時は解放処理を行う
                If Not (_IsServerStart) Then ShutDownFlag = True 'クライアント発呼ならば受信しないので開放処理を行う

                If ShutDownFlag Then    '解放処理を行う場合

                    'タイムアウト設定時間の間、コネクションが切れたかどうか確認する
                    Dim ConnectOutStart As DateTime = Now   'コネクション切断監視開始の時間
                    Do
                        If _SocketState Is Nothing Then Exit Do 'ソケット管理オブジェクトが無い時は抜ける
                        With _SocketState
                            If .WorkSocket Is Nothing Then Exit Do 'ソケットオブジェクトが無い時は抜ける
                            If Not (IsConnecting(.WorkSocket)) Then Exit Do 'コネクションが切断されているときは抜ける
                        End With
                        Dim tmpSpan As TimeSpan = Now - ConnectOutStart
                        If _設定値.WriteTimeOut < tmpSpan.TotalMilliseconds Then Exit Do '送信タイムアウトで設定された時間を超えたら抜ける
                        System.Threading.Thread.Sleep(_設定値.ConnectTimeOut)    'タイムアウトで設定された時間だけ待機
                    Loop

                    'ソケットの解放をする
                    If Not (_SocketState Is Nothing) Then
                        With _SocketState
                            'If Not (.WorkSocket Is Nothing) Then If .WorkSocket.Connected Then .WorkSocket.Shutdown(SocketShutdown.Both) '通信を閉じる
                            If Not (.WorkSocket Is Nothing) Then
                                If IsConnecting(.WorkSocket) Then
                                    .WorkSocket.Shutdown(SocketShutdown.Both) '通信を閉じる
                                    ErrMsgBox(Nothing, _設定値.DebugMode, "受信後、サーバからコネクションを閉じました。(リモート:" & LogRemoteText & ")", 0, _設定値.LogPath)
                                Else
                                    ErrMsgBox(Nothing, _設定値.DebugMode, "受信後、相手がコネクションを閉じました。(リモート:" & LogRemoteText & ")", 0, _設定値.LogPath)
                                End If
                            End If
                            If Not (.WorkSocket Is Nothing) Then .WorkSocket.Close() '開放
                            If Not (.WorkSocket Is Nothing) Then .WorkSocket.Dispose() '開放
                            If Not (.WorkSocket Is Nothing) Then .WorkSocket = Nothing '開放
                        End With
                    End If

                End If

                '他のオブジェクトを明示的に解放する
                If Not (SendDone Is Nothing) Then SendDone.Dispose()
                If Not (SendDone Is Nothing) Then SendDone = Nothing

                WriteLog(_設定値.LogPath, "output", LogRemoteText, LogSendText & ErrMessage) 'ログを出力

                _EndDateTime = Now  '処理終了日時
                _Complete = True    '処理が終わったかどうか

            Catch ex As Exception
            End Try
        End Try
    End Sub

    ''' <summary                    >送信開始前に通信相手との接続を試みる関数。                    </summary>
    ''' <param name="SrcSocket"     >接続に使うするソケットオブジェクト                            </param  >
    ''' <param name="RemoteEndPoint">接続相手の情報を格納したIPEndPointオブジェクト                </param  >
    ''' <returns                    >接続の正否を返す。                                            </returns>
    ''' <remarks                    >サーバ発呼用。相手先設定の不備でも判断せず試行して結果を返す。</remarks>
    Private Function TryConnect_同期(ByRef SrcSocket As Socket, ByRef RemoteEndPoint As IPEndPoint) As Boolean
        'Dim mreSock As ManualResetEvent
        Try
            'mreSock = New ManualResetEvent(False)
            '試行回数分だけ接続を試みる
            Dim TryCount As Long = _設定値.MaxSendCount    '接続試行回数を指定する
            Do
                If TryCount < 1 Then Exit Do '接続試行回数が尽きたらループを抜ける
                Try
                    If SrcSocket Is Nothing Then : ErrMsgBox(Nothing, _設定値.DebugMode, "(送信)接続に失敗しました。送信ソケットインスタンスがNothingです。", 1, _設定値.LogPath) : Return False : End If
                    ErrMsgBox(Nothing, _設定値.DebugMode, "サーバ発呼のセッション作成を行います(残り回数：" & TryCount & ")(リモート:" & RemoteEndPoint.ToString & ")", 0, _設定値.LogPath)
                    If Not (IsConnecting(SrcSocket)) Then SrcSocket.Connect(RemoteEndPoint) '送信相手と接続する
                    If IsConnecting(SrcSocket) Then
                        ErrMsgBox(Nothing, _設定値.DebugMode, "サーバ発呼のセッションを確立しました(リモート:" & RemoteEndPoint.ToString & ")", 0, _設定値.LogPath)
                        Return True '接続できたら真を返す
                    End If
                    ErrMsgBox(Nothing, _設定値.DebugMode, "サーバ発呼のセッション確立に失敗しました(タイムアウト)(残り回数：" & TryCount & ")(リモート:" & RemoteEndPoint.ToString & ")", 1, _設定値.LogPath)
                Catch ex As Exception
                    ErrMsgBox(ex, _設定値.DebugMode, "サーバ発呼のセッション確立に失敗しました(例外)(残り回数：" & TryCount & ")(リモート:" & RemoteEndPoint.ToString & ")", 1, _設定値.LogPath)
                    '接続失敗でも何もせず次の試行へ移る。
                End Try
                'System.Threading.Thread.Sleep(_設定値.ConnectTimeOut)    'タイムアウトで設定された時間だけ待機
                TryCount -= 1 '接続試行回数を減らす
            Loop
            ErrMsgBox(Nothing, _設定値.DebugMode, "(送信)接続に失敗しました。試行回数が設定した最大数を超えました。", 1, _設定値.LogPath)
            Return False '接続試行回数が尽きたら偽を返す
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "(送信)接続に失敗しました。(例外)", 2, _設定値.LogPath)
            Return False '他のエラー失敗があっても特に何もせず偽を返す
        Finally
            'If Not (mreSock Is Nothing) Then mreSock.Dispose()
        End Try
    End Function

    ''' <summary                    >送信開始前に通信相手との接続を試みる関数。(非同期)            </summary>
    ''' <param name="SrcSocket"     >接続に使うするソケットオブジェクト                            </param  >
    ''' <param name="RemoteEndPoint">接続相手の情報を格納したIPEndPointオブジェクト                </param  >
    ''' <returns                    >接続の正否を返す。                                            </returns>
    ''' <remarks                    >サーバ発呼用。相手先設定の不備でも判断せず試行して結果を返す。</remarks>
    Private Function TryConnect_非同期(ByRef SrcSocket As Socket, ByRef RemoteEndPoint As IPEndPoint) As Boolean
        'Dim mreSock As ManualResetEvent
        Try
            'mreSock = New ManualResetEvent(False)
            '試行回数分だけ接続を試みる
            Dim TryCount As Long = _設定値.MaxSendCount    '接続試行回数を指定する
            Do
                If TryCount < 1 Then Exit Do '接続試行回数が尽きたらループを抜ける
                Try
                    'If SrcSocket Is Nothing Then : ErrMsgBox(Nothing, _設定値.DebugMode, "(送信)接続に失敗しました。送信ソケットインスタンスがNothingです。", 1, _設定値.LogPath) : Return False : End If
                    ErrMsgBox(Nothing, _設定値.DebugMode, "サーバ発呼のセッション作成を行います(残り回数：" & TryCount & ")(リモート:" & RemoteEndPoint.ToString & ")", 0, _設定値.LogPath)
                    '一度接続試行に使ったソケットは使いまわせないので、接続試行する度に作り直す
                    Dim DestSocket As New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)

                    ConnectSuccess = False '送信成功のフラグを下げる
                    ConnectDone = New ManualResetEvent(False)   'フラグのインスタンスを作る
                    ConnectDone.Reset()     'スレッドをブロックできるようフラグをリセット
                    If Not (IsConnected(DestSocket)) Then '                        If Not (SrcSocket.Connected) Then
                        'SrcSocket.MulticastLoopback = True
                        DestSocket = CType(DestSocket.BeginConnect(RemoteEndPoint, New AsyncCallback(AddressOf ConnectCallback), DestSocket).AsyncState, Socket)   'コールバック関数を設定して非同期で接続
                    End If
                    ConnectDone.WaitOne(_設定値.ConnectTimeOut, False)
                    If Not (DestSocket Is Nothing) Then
                        SrcSocket = DestSocket
                    End If
                    If ConnectSuccess Or IsConnected(DestSocket) Then ' SrcSocket.Connected Then 'If IsConnecting(SrcSocket) Then
                        ErrMsgBox(Nothing, _設定値.DebugMode, "サーバ発呼のセッションを確立しました(リモート:" & RemoteEndPoint.ToString & ")", 0, _設定値.LogPath)
                        Return True '接続できたら真を返す
                    End If
                    DestSocket.Close()
                    'Else
                    'ErrMsgBox(Nothing, _設定値.DebugMode, "意図しないタイミングでセッションが存在しています(リモート:" & RemoteEndPoint.ToString & ")", 0, _設定値.LogPath)
                    'Return False
                    'End If
                    ErrMsgBox(Nothing, _設定値.DebugMode, "サーバ発呼のセッション確立に失敗しました(残り回数：" & TryCount & ")(リモート:" & RemoteEndPoint.ToString & ")", 1, _設定値.LogPath)
                Catch ex As Exception
                    ErrMsgBox(ex, _設定値.DebugMode, "サーバ発呼のセッション確立に失敗しました(フックされた例外)(残り回数：" & TryCount & ")(リモート:" & RemoteEndPoint.ToString & ")", 1, _設定値.LogPath)
                    '接続失敗でも何もせず次の試行へ移る。
                End Try
                TryCount -= 1 '接続試行回数を減らす
            Loop
            ErrMsgBox(Nothing, _設定値.DebugMode, "(送信)接続に失敗しました。試行回数が設定した最大数を超えました。", 1, _設定値.LogPath)
            Return False '接続試行回数が尽きたら偽を返す
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "(送信)接続に失敗しました。(例外)", 2, _設定値.LogPath)
            Return False '他のエラー失敗があっても特に何もせず偽を返す
        Finally
            'If Not (ConnectDone Is Nothing) Then ConnectDone.Dispose()
        End Try
    End Function

    ''' <summary>非同期でサーバ発呼による接続を行う為のコールバック関数</summary>
    ''' <param name="Arg">引数コンテナクラス(Socket)      </param  >
    ''' <remarks>旧来は同期処理で接続していたが、タイムアウトの監視ができないので非同期処理とした</remarks>
    Private Sub ConnectCallback(ByVal Arg As IAsyncResult)
        Dim Speaker As Socket
        Try
            'ErrMsgBox(Nothing, _設定値.DebugMode, "非同期接続関数内開始", 0, _設定値.LogPath)
            '送信処理
            Speaker = CType(Arg.AsyncState, Socket)               '引数を変換
            Try
                Speaker.EndConnect(Arg)
                ConnectSuccess = True '接続成功のフラグを上げる
            Catch ex As Exception
                'Close時にSocketがNothingで呼び出される事があるのでその時は何もしない
            End Try
            Return
        Catch ex As Exception
            '接続失敗時は
            ErrMsgBox(ex, _設定値.DebugMode, "非同期接続関数内で例外が発生しました。", 2, _設定値.LogPath)
            ConnectSuccess = False '送信成功のフラグを下げる
        Finally
            'ErrMsgBox(Nothing, _設定値.DebugMode, "非同期接続関数内でフラグ解除", 0, _設定値.LogPath)
            If Not (ConnectDone Is Nothing) Then ConnectDone.Set() '接続完了待機を解除
        End Try
    End Sub

    ''' <summary               >データ送信を試みる関数。            </summary>
    ''' <param name="SocketState">接続に使うする引数オブジェクト      </param  >
    ''' <returns               >送信の正否を返す。                  </returns>
    ''' <remarks               >ここからコールバック関数を呼び出す。</remarks>
    Private Function TrySend(ByRef SocketState As cSocketState) As Boolean
        Try
            '実行条件を判断する
            With SocketState
                If .WorkSocket Is Nothing Then : ErrMsgBox(Nothing, _設定値.DebugMode, "送信に失敗しました。送信ソケットインスタンスがNothingです。", 1, _設定値.LogPath) : Return False : End If 'ソケットが無い時は抜ける
                'If Not (.WorkSocket.Connected) Then : ErrMsgBox(Nothing, _設定値.DebugMode, "送信に失敗しました。送信ソケットのコネクションが接続状態ではありません。", 1, _設定値.LogPath) : Return False : End If '未接続時は抜ける
                If Not (IsConnecting(.WorkSocket)) Then : ErrMsgBox(Nothing, _設定値.DebugMode, "送信に失敗しました。送信ソケットのコネクションが接続状態ではありません。", 1, _設定値.LogPath) : Return False : End If '未接続時は抜ける
                SendSuccess = False     '送信成功フラグをリセット
                SendDone = New ManualResetEvent(False)  'ここでフラグのインスタンスを作る。
                SendDone.Reset()        'このスレッドをブロックできるようにフラグをリセット
                .WorkSocket.BeginSend(.SendBuffer, 0, .SendBuffer.Length, 0, New AsyncCallback(AddressOf SendCallback), SocketState)   'コールバック関数を指定して非同期で送信
                SendDone.WaitOne(_設定値.WriteTimeOut)  'タイムアウトを設定してブロック
                If SendSuccess Then Return True '送信成功時は真を返す
            End With
            ErrMsgBox(Nothing, _設定値.DebugMode, "送信に失敗しました。タイムアウトと思われます。", 1, _設定値.LogPath)
            Return False
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "送信に失敗しました。(例外)", 2, _設定値.LogPath)
            Return False    'エラー時もタイムアウト時も偽を返す
        End Try
    End Function

    ''' <summary         >データ送信のコールバック関数          </summary>
    ''' <param name="Arg">引数コンテナクラス(cSocketState)      </param  >
    ''' <remarks         >タブレット発呼・サーバ発呼の両方で使う</remarks>
    Private Sub SendCallback(ByVal Arg As IAsyncResult)
        Dim Speaker As cSocketState
        Try
            '送信処理
            Speaker = CType(Arg.AsyncState, cSocketState)               '引数を変換
            Dim BytesSend As Integer = Speaker.WorkSocket.EndSend(Arg)  '送信を完了する
            SendSuccess = True  '送信成功フラグ
            Return
        Catch ex As Exception
            '送信失敗時は
            ErrMsgBox(ex, _設定値.DebugMode, "非同期送信関数内で例外が発生しました。", 2, _設定値.LogPath)
            SendSuccess = False '送信成功のフラグを下げる
        Finally
            If Not (SendDone Is Nothing) Then SendDone.Set() '受信完了待機を解除
        End Try
    End Sub

    ''' <summary                 >送信時の文字列処理を行う</summary>
    ''' <param name="CommandArgs">条件と結果のオブジェクト</param  >
    ''' <returns                 >処理が成功したかどうか  </returns>
    ''' <remarks                 >                        </remarks>
    Private Function ConvCPToWriteData(ByRef CommandArgs As cCommunicateEventArgs) As Boolean
        Try
            '送信データがが条件を満たしているか調べる
            With CommandArgs
                Dim Param As String = .WriteParam.ToString.Trim
                Dim CmdCode As String = .WriteCommand.ToString.Trim

                If CmdCode Is Nothing Then : ErrMsgBox(Nothing, _設定値.DebugMode, "返信コマンドが存在しません。", 1, _設定値.LogPath) : Return False : End If
                If Not (CmdCode.Length = 2) Then : ErrMsgBox(Nothing, _設定値.DebugMode, "返信コマンドの文字数が正しくありません(" & CmdCode.Length & "文字)。", 1, _設定値.LogPath) : Return False : End If

                '送信データを作る
                .OrgWriteText = CmdCode & Chr(_設定値.ReadSeparateByte) & Param & Chr(_設定値.ReadEndByte) '送信文字列を作る
                .OrgWriteBytes = System.Text.Encoding.UTF8.GetBytes(.OrgWriteText)                         'Byte型配列に変換する

            End With
            Return True
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "送信文字列の処理で例外が発生しました。", 2, _設定値.LogPath)
            Return False
        End Try
    End Function

#End Region

End Class

'''<summary>送信のTaskList(疑似スレッドプール)を管理するクラス</summary>
Public Class cSendTaskList

#Region "内部メンバ"

    '''<summary>上位から渡されるグローバル設定値</summary>
    Private _設定値 As c設定値

    ''' <summary>スレッドプールに登録した送信オブジェクトのリスト</summary>
    Private _SenderList As List(Of cSenderTask)   'スレッドプールに登録した送信オブジェクトのリスト

    ''' <summary>受信用スレッドプールクラス</summary>
    Private _RecieveTaskList As cRecieveTaskList    '受信用スレッドプールクラス

#End Region

#Region "イベント"

    ''' <summary                     >コンストラクタ                                </summary>
    ''' <param name="Arg設定値"      >グローバル設定値                              </param  >
    ''' <param name="RecieveTaskList">送信後に受信を行う受信タスクリストオブジェクト</param  >
    Public Sub New(Arg設定値 As c設定値, ByRef RecieveTaskList As cRecieveTaskList)
        Try
            _設定値 = Arg設定値                     '設定値
            _RecieveTaskList = RecieveTaskList      '受信タスクリスト
            _SenderList = New List(Of cSenderTask)  'インスタンス作成
            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "送信タスクリストのコンストラクタで例外が発生しました。", 2, _設定値.LogPath)
            Return
        End Try
    End Sub

    ''' <summary>デストラクタ</summary>
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#End Region

#Region "パブリックメソッド"

    ''' <summary                 >送信コネクションをタスクのListに登録して実行する</summary>
    ''' <param name="SocketState">送信に用いるソケット関連引数                    </param  >
    ''' <param name="Command"    >コマンド処理に用いるコマンド関連引数            </param  >
    ''' <param name="ShutDown"   >送信後にシャットダウンを行うかどうか            </param  >
    ''' <param name="RemoteIP"   >接続先のIPアドレス(SocketStateが未接続時に使用) </param  >
    ''' <param name="RemotePort" >接続先のポート番号(SocketStateが未接続時に使用) </param  >
    Public Sub AddSend(ByRef SocketState As cSocketState, ByRef Command As cCommunicateEventArgs, ShutDown As Boolean, Optional ByVal RemoteIP As String = "", Optional ByVal RemotePort As Integer = 0)
        Try
            'タスクと送信クラスを作成
            Dim AddSenderTask As New cSenderTask(_設定値, SocketState, Command, ShutDown, _RecieveTaskList, RemoteIP, RemotePort)
            _SenderList.Add(AddSenderTask)   'リストに登録

            'ここで処理の終わった送信インスタンスを解除する
            'For Eachではエラーになる
            For Index As Long = _SenderList.Count - 1 To 0 Step -1
                Dim DeleteFlag As Boolean = False   '削除フラグ
                With _SenderList(Index).TaskObj
                    If .IsCompleted Then DeleteFlag = True 'タスク側で完全終了していれば削除対象
                    If .IsFaulted Then DeleteFlag = True 'タスク側で失敗していれば削除対象
                    If .IsCanceled Then DeleteFlag = True 'タスク側でキャンセルされていれば削除対象
                End With

                'キャンセルを行うかどうかの判断だが、未実装
                'With _SenderList(Index).Sender
                '    '削除するかどうかを調べる
                '    If .Complete Then   '終了時
                '        If .Success Then    '成功終了時は削除
                '            DeleteFlag = True
                '        Else                '失敗で終了時は削除（再送する？）
                '            DeleteFlag = True
                '        End If
                '    Else    '未完了の時
                '    End If
                'End With

                If DeleteFlag Then  '削除対象なら削除する
                    '_SenderList.Remove(_SenderList(Index))  'リストから削除する
                End If
            Next
            Return  '抜ける
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "送信タスクリストの登録・実行で例外が発生しました。", 2, _設定値.LogPath)
            Return
        End Try
    End Sub

#End Region

    ''' <summary>タスクと処理を配列で管理する為のクラス</summary>
    Private Class cSenderTask

#Region "内部メンバ"

        ''' <summary>上位から渡される設定値</summary>
        Private _設定値 As c設定値

#End Region

#Region "プロパティ"

        ''' <summary>送信を処理するクラス</summary>
        Private _Sender As cSend  '送信クラス
        Public ReadOnly Property Sender As cSend
            Get
                Return _Sender
            End Get
        End Property

        ''' <summary>送信を別スレッドで処理するタスク</summary>
        Private _Task As Task         'タスククラス
        Public ReadOnly Property TaskObj As Task
            Get
                Return _Task
            End Get
        End Property

#End Region

#Region "イベント"

        ''' <summary                     >コンストラクタ                              </summary>
        ''' <param name="Arg設定値"      >上位から渡される設定値                      </param  >
        ''' <param name="SocketState"    >送信に用いるソケット関連引数                </param  >
        ''' <param name="Command"        >コマンド処理に用いるコマンド関連引数        </param  >
        ''' <param name="ShutDown"       >送信後にシャットダウンを行うかどうか        </param  >
        ''' <param name="RecieveTaskList">送信後に受信を行うタスクリストのインスタンス</param  >
        ''' <remarks                     >インスタンス作成と同時に送信処理を開始する  </remarks>
        Public Sub New(Arg設定値 As c設定値, ByRef SocketState As cSocketState, ByRef Command As cCommunicateEventArgs, ShutDown As Boolean, ByRef RecieveTaskList As cRecieveTaskList, Optional ByVal RemoteIP As String = "", Optional ByVal RemotePort As Integer = 0)
            Try
                _設定値 = Arg設定値                       '設定値
                _Sender = New cSend(_設定値, SocketState, Command, ShutDown, RecieveTaskList, RemoteIP, RemotePort) '受信タスクのインスタンスを作成
                _Sender.AddDateTime = Now                 '現在日時を登録
                _Task = New Task(AddressOf _Sender.Done)  'タスクに処理を登録
                _Task.Start()                             'さっさと開始
                Return
            Catch ex As Exception
                ErrMsgBox(ex, _設定値.DebugMode, "送信タスクのコンストラクタで例外が発生しました。", 2, _設定値.LogPath)
                Return
            End Try
        End Sub

        ''' <summary>デストラクタ</summary>
        Protected Overrides Sub Finalize()
            Try
                _Task.Dispose()
                MyBase.Finalize()
                Return
            Catch ex As Exception
                ErrMsgBox(ex, _設定値.DebugMode, "送信タスクのデストラクタで例外が発生しました。", 2, _設定値.LogPath)
                Return
            End Try
        End Sub

#End Region

    End Class

End Class

' '''<summary>処理のTaskList(疑似スレッドプール)を管理する(基底)クラス</summary>
'Public Class cTaskList

'#Region "内部メンバ"

'    '''<summary>上位から渡されるグローバル設定値</summary>
'    Private _設定値 As c設定値

'    ''' <summary>スレッドプールに登録した送信オブジェクトのリスト</summary>
'    Private _List As List(Of cTaskProcess)   'スレッドプールに登録した処理オブジェクトのリスト

'    ''' <summary>受信用スレッドプールクラス</summary>
'    Private _NextTaskList As cTaskList    'ここでの処理が終わった次に実行する処理のクラス

'#End Region

'#Region "イベント"

'    ''' <summary               >コンストラクタ                                        </summary>
'    ''' <param name="Arg設定値">グローバル設定値                                      </param  >
'    ''' <param name="NextTaskList" >ここでの処理後に次の処理を行うタスクリストオブジェクト</param  >
'    Public Sub New(Arg設定値 As c設定値, ByRef NextTaskList As cTaskList)
'        Try
'            _設定値 = Arg設定値   '設定値
'            _NextTaskList = NextTaskList  'タスクリスト
'            Return
'        Catch ex As Exception
'            ErrMsgBox(ex, _設定値.DebugMode, "送信タスクリストのコンストラクタで例外が発生しました。", 2, _設定値.LogPath)
'            Return
'        End Try
'    End Sub

'    ''' <summary>デストラクタ</summary>
'    Protected Overrides Sub Finalize()
'        MyBase.Finalize()
'    End Sub

'#End Region

'#Region "パブリックメソッド"

'    ''' <summary             >送信コネクションをタスクのListに登録して実行する</summary>
'    ''' <param name="Process">実行するクラスインスタンス                      </param  >
'    Public Sub AddSend(ByRef Process As Object)
'        Try
'            '動作条件をチェック
'            If Process Is Nothing Then Return '処理インスタンスが無い時は抜ける
'            If IsNothing(Process.GetType.GetMethod("Done")) Then Return 'メソッドが無い時は抜ける
'            If IsNothing(Process.GetType.GetProperty("AddDateTime")) Then Return 'プロパティが無い時は抜ける
'            If IsNothing(Process.GetType.GetProperty("Complete")) Then Return 'プロパティが無い時は抜ける
'            If IsNothing(Process.GetType.GetProperty("Success")) Then Return 'プロパティが無い時は抜ける

'            Dim AddTaskProcess As New cTaskProcess(_設定値)
'            'タスクと送信クラスを作成
'            AddTaskProcess.TaskStart()
'            _List.Add(AddTaskProcess)   'リストに登録

'            'ここで処理の終わった送信インスタンスを解除する
'            'For Eachではエラーになる
'            For Index As Long = _List.Count - 1 To 0 Step -1
'                Dim DeleteFlag As Boolean = False   '削除フラグ
'                With _List(Index).TaskObj
'                    If .IsCompleted Then DeleteFlag = True 'タスク側で完全終了していれば削除対象
'                End With
'                With _List(Index).Process
'                    '削除するかどうかを調べる
'                    If .Complete Then   '終了時
'                        If .Success Then    '成功終了時は削除
'                            DeleteFlag = True
'                        Else                '失敗で終了時は削除（再送する？）
'                            DeleteFlag = True
'                        End If
'                    Else    '未完了の時
'                    End If
'                End With

'                If DeleteFlag Then  '削除対象なら削除する
'                    _List.Remove(_List(Index))  'リストから削除する
'                End If
'            Next
'            Return  '抜ける
'        Catch ex As Exception
'            ErrMsgBox(ex, _設定値.DebugMode, "送信タスクリストへの追加時に例外が発生しました。", 2, _設定値.LogPath)
'            Return
'        End Try
'    End Sub

'#End Region

'#Region "内部関数"

'    ''' <summary>タスクと処理を配列で管理する為のクラス</summary>
'    Private Class cTaskProcess

'#Region "内部変数メンバ"

'        ''' <summary>上位から渡される設定値</summary>
'        Private _設定値 As c設定値

'#End Region

'#Region "プロパティメンバ"

'        ''' <summary>処理を行うクラス</summary>
'        Public Process As Object  '処理を行う

'        ''' <summary>_Processを別スレッドで実行するタスク</summary>
'        Private _Task As Task       'タスククラス
'        Public ReadOnly Property TaskObj As Task
'            Get
'                Return _Task
'            End Get
'        End Property

'#End Region

'#Region "イベント"

'        ''' <summary                     >コンストラクタ                              </summary>
'        ''' <param name="Arg設定値"      >上位から渡される設定値                      </param  >
'        ''' <remarks                     >インスタンス作成と同時に送信処理を開始する  </remarks>
'        Public Sub New(Arg設定値 As c設定値)
'            Try
'                _設定値 = Arg設定値                       '設定値
'                Return
'            Catch ex As Exception
'                ErrMsgBox(ex, _設定値.DebugMode, "送信タスクリストアイテムのコンストラクタで例外が発生しました。", 2, _設定値.LogPath)
'                Return
'            End Try
'        End Sub

'        ''' <summary>デストラクタ</summary>
'        Protected Overrides Sub Finalize()
'            Try
'                _Task.Dispose()     'タスクを解放
'                MyBase.Finalize()
'                Return
'            Catch ex As Exception
'                ErrMsgBox(ex, _設定値.DebugMode, "送信タスクリストアイテムのデストラクタで例外が発生しました。", 2, _設定値.LogPath)
'                Return
'            End Try
'        End Sub

'#End Region

'#Region "パブリックメソッド"

'        ''' <summary>処理を開始する</summary>
'        Public Sub TaskStart()
'            Try
'                '実行条件を調べる
'                If Process Is Nothing Then Return '処理インスタンスが無い時は抜ける
'                If IsNothing(Process.GetType.GetMethod("Done")) Then Return 'Doneメソッドが無い時は抜ける
'                If IsNothing(Process.GetType.GetProperty("AddDateTime")) Then Return 'AddDateTimeプロパティが無い時は抜ける
'                'Taskに登録して実行
'                Process.AddDateTime = Now  '現在日時を登録
'                _Task = New Task(AddressOf Process.Done)  'タスクに処理を登録
'                _Task.Start()   '開始
'                Return
'            Catch ex As Exception
'                ErrMsgBox(ex, _設定値.DebugMode, "送信タスクの開始時に例外が発生しました。", 2, _設定値.LogPath)
'                Return
'            End Try
'        End Sub

'#End Region

'    End Class

'#End Region

'End Class

#End Region

#Region "受信接続関連クラス"

''' <summary>受信接続を待ち続けるクラス</summary>
Public Class cAccept

#Region "内部メンバ"

    '''<summary>上位から渡されるグローバル設定値</summary>
    Private _設定値 As c設定値

    ''' <summary>受信完了まで受信スレッド実行をブロックする</summary>
    Private Shared RecieveDone As New ManualResetEvent(False)

    ''' <summary>受信が成功したかどうかのフラグ</summary>
    Private Shared RecieveSuccess As Boolean = False

    ''' <summary>受信用疑似スレッドプールクラス</summary>
    Private _RecieveTaskList As cRecieveTaskList

    ''' <summary>受信待機のコールバック(AcceptCallback)で得たソケットオブジェクト</summary>
    Private Shared AcceptSocket As Socket

    ''' <summary>受信待機のループを止めるフラグ</summary>
    Private StopFlag As Boolean = False

    ''' <summary>受信待機を繰り返し続けるスレッド</summary>
    Dim AcceptLoopThread As System.Threading.Thread

#End Region

#Region "イベント"

    ''' <summary                     >コンストラクタ              </summary>
    ''' <param name="Arg設定値"      >グローバル設定値            </param  >
    ''' <param name="RecieveTaskList">受信タスクリストオブジェクト</param  >
    Public Sub New(Arg設定値 As c設定値, ByRef RecieveTaskList As cRecieveTaskList)
        Try
            _設定値 = Arg設定値                 '設定値
            _RecieveTaskList = RecieveTaskList  '受信タスクリスト
            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "受信接続クラスのコンストラクタで例外が発生しました。", 2, _設定値.LogPath)
            Return
        End Try
    End Sub

    ''' <summary>デストラクタ</summary>
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#End Region

#Region "パブリックメソッド"

    ''' <summary>受信待機のループ・スレッドを開始する関数</summary>
    Public Sub StartAcceptLoop()
        Try
            StopFlag = False    '停止フラグを下げる
            'スレッドを実行        
            AcceptLoopThread = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf AcceptLoop))
            AcceptLoopThread.IsBackground = True   'バックグラウンドで実行
            AcceptLoopThread.Start()               'スレッドを開始する
            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "受信接続クラスのループスレッド開始で例外が発生しました。", 2, _設定値.LogPath)
            Return
        End Try
    End Sub

    ''' <summary>受信待機のループ・スレッドを停止する関数          </summary>
    ''' <remarks>ループ停止のフラグを上げ、スレッドの強制停止を行う</remarks>
    Public Sub StopAcceptLoop()
        Try
            StopFlag = True             '停止フラグを上げる
            System.Threading.Thread.Sleep(1000)    '待機
            AcceptLoopThread.Abort()    'スレッド強制停止
            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "受信待機中ループの停止時に例外が発生しました。", 2, _設定値.LogPath)
        End Try
    End Sub

#End Region

#Region "内部関数"

    ''' <summary>受信待機のループを行う関数</summary>
    Private Sub AcceptLoop()
        Try
            '受信待機する 
            Dim LocalEndPoint As New IPEndPoint(System.Net.IPAddress.Any, _設定値.ReadPort)     '受信デバイスを指定する
            Dim SocketObject As New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
            SocketObject.Bind(LocalEndPoint)
            SocketObject.Listen(10000)
            Do
                Try
                    RecieveSuccess = False     '受信成功フラグをリセット
                    RecieveDone.Reset()        'このスレッドをブロックできるようにフラグをリセット
                    SocketObject.BeginAccept(New AsyncCallback(AddressOf AcceptCallback), SocketObject)
                    RecieveDone.WaitOne()      '接続できるまでここでずっとブロックするよ！
                    If RecieveSuccess Then
                        Dim SocketState As New cSocketState(_設定値) '引数オブジェクトを作る
                        SocketState.WorkSocket = AcceptSocket
                        _RecieveTaskList.AddRecieve(SocketState, New cCommunicateEventArgs, False)
                    Else
                        '接続失敗時は
                        WriteLog(_設定値.LogPath, "input", "接続相手不明", "受信文字列なし" & ", 通信相手との接続に失敗しました。") 'エラーログを出力
                    End If
                Catch ex As Exception
                    'エラーでも止まらない
                End Try
                If StopFlag Then Exit Do
            Loop
            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "受信待機中に例外が発生しました。受信待ちループが止まってる恐れがあります。", 2, _設定値.LogPath)
        End Try
    End Sub

    ''' <summary         >受信待機のコールバック関数      </summary>
    ''' <param name="Arg">引数コンテナクラス(cSocketState)</param  >
    ''' <remarks         >タブレット発呼の受信待機を行う  </remarks>
    Private Sub AcceptCallback(ByVal Arg As IAsyncResult)
        Dim Listener As Socket
        Try
            Listener = CType(Arg.AsyncState, Socket)              '引数を変換
            Dim Handler As Socket = Listener.EndAccept(Arg)  '接続を受け入れる
            AcceptSocket = Handler
            RecieveSuccess = True  '受信成功のフラグを上げる
            Return  '抜ける
        Catch ex As Exception
            RecieveSuccess = False  '受信成功のフラグを下げる
            ErrMsgBox(ex, _設定値.DebugMode, "非同期受信接続関数内で例外が発生しました。", 2, _設定値.LogPath)
            Return
        Finally
            If Not (RecieveDone Is Nothing) Then RecieveDone.Set() '受信完了待機を解除
        End Try
    End Sub

#End Region

End Class

#End Region

#Region "引数コンテナクラス"

''' <summary>非同期通信のコールバック用引数コンテナクラス</summary>
''' <remarks>                                            </remarks>
Public Class cSocketState

#Region "プロパティ"

    ''' <summary>クライアントソケット</summary>
    Public WorkSocket As Socket = Nothing
    ''' <summary>送受信バッファサイズ</summary>
    Public Const BufferSize As Integer = 1024
    ''' <summary>送信バッファ</summary>
    Public SendBuffer(BufferSize) As Byte
    ''' <summary>受信バッファ</summary>
    Public ReadBuffer(BufferSize) As Byte
    ''' <summary>送信文字列</summary>
    Public SendText As String = ""
    ' ''' <summary>この通信がサーバ発呼かどうか</summary>
    ' Public IsServerStart As Boolean = False
    '''<summary>送信の試行回数</summary>
    '''<remarks>送信を試みる度にデクリメントされ、0でエラーを発呼する為のカウンタ</remarks>
    Public SendTryCount As Long = 3
    '''<summary>受信の試行回数</summary>
    '''<remarks>受信を試みる度にデクリメントされ、0でエラーを発呼する為のカウンタ</remarks>
    Public ReadTryCount As Long = 3


    '''<summary>上位から渡されるグローバル設定値</summary>
    Private _設定値 As c設定値
    '''<summary>上位から渡されるグローバル設定値</summary>
    Public ReadOnly Property 設定値() As c設定値
        Get
            Return _設定値
        End Get
    End Property

    ' ''' <summary>受信蓄積データ</summary>
    Public _ReadData() As Byte
    ' ''' <summary>受信蓄積データ</summary>
    Public ReadOnly Property ReadData() As Byte()
        Get
            Return _ReadData
        End Get
    End Property

#End Region

    ''' <summary                >コンストラクタ  </summary>
    ''' <param name="Arg設定値" >グローバル設定値</param  >
    Public Sub New(Arg設定値 As c設定値)
        _設定値 = Arg設定値
        ResetSendTryCount() '送信試行回数を代入
        ResetReadTryCount() '受信試行回数を代入
    End Sub

#Region "その他パブリックメソッド"

    ''' <summary>受信試行回数をリセットする関数</summary>
    Public Sub ResetReadTryCount()
        ReadTryCount = _設定値.MaxSendCount  '受信試行回数を代入
    End Sub

    ''' <summary>送信試行回数をリセットする関数</summary>
    Public Sub ResetSendTryCount()
        SendTryCount = _設定値.MaxSendCount  '送信試行回数を代入
    End Sub

    ''' <summary>受信バッファをリセットする関数</summary>
    Public Sub ResetReadBuffer()
        ReadBuffer = New Byte(BufferSize) {}
    End Sub

    ''' <summary>蓄積した受信データ配列をリセットする関数</summary>
    Public Sub ResetReadData()
        Dim tmp() As Byte
        _ReadData = tmp
    End Sub

    ''' <summary            >ReadBufferを_ReadDataに追加する関数                       </summary>
    ''' <param name="Length">追加するサイズ　　　　　　   　　　                       </param  >
    ''' <remarks            >受信データを受信バッファから受信データ配列に加える為に使う</remarks>
    Public Sub AddReadAddBuffer(Length As Long)
        Try
            If ReadBuffer.Length < Length Then Err.Raise(9999, Nothing, "追加したいデータサイズが追加元のバッファのサイズを超えています。")
            If Length < 1 Then Return
            '追加処理する
            Dim Dest As Byte()
            If _ReadData Is Nothing Then    '元々のデータがNUllなら
                Dest = New Byte(Length - 1) {}                      'メモリ確保
                Buffer.BlockCopy(ReadBuffer, 0, Dest, 0, Length)    '追加するデータをコピー
                _ReadData = Dest
                Return                                              '抜ける
            End If
            If _ReadData.Length < 1 Then    '元々のデータサイズがゼロなら
                Dest = New Byte(Length - 1) {}                      'メモリ確保
                Buffer.BlockCopy(ReadBuffer, 0, Dest, 0, Length)    '追加するデータをコピー
                _ReadData = Dest
                Return                                              '抜ける
            End If
            'データの追加を行う
            Dest = New Byte(_ReadData.Length + (Length - 1)) {}   'メモリ確保
            Buffer.BlockCopy(_ReadData, 0, Dest, 0, _ReadData.Length)           '元々のデータをコピー
            Buffer.BlockCopy(ReadBuffer, 0, Dest, _ReadData.Length, Length)     '追加するデータをコピー
            _ReadData = Dest    '値を置き換え
            Return
        Catch ex As Exception
            Err.Raise(9999, Nothing, "/バッファの蓄積処理に失敗しました。/" & ex.Message)
        End Try
    End Sub

    ''' <summary            >受信範囲のReadBufferに終端文字が含まれている位置を返す    </summary>
    ''' <param name="Length">受信したデータ長                                          </param  >
    ''' <returns            >終端文字の位置(何文字目か)を返す。存在しない時は-1を返す。</returns>
    Public Function InReadEnd(Length As Long) As Long
        Try
            If Length = 0 Then Return -1
            For Index As Long = 0 To Length - 1
                If _設定値.ReadEndByte = ReadBuffer(Index) Then Return Index + 1 '終端ならばその位置を返す
            Next
            Return -1
        Catch ex As Exception
            Err.Raise(9999, Nothing, "/受信終端の判断に失敗しました。/" & ex.Message)
            Return -1
        End Try
    End Function

#End Region

End Class

''' <summary>cCommunicateのイベント引数クラス</summary>
''' <remarks>                                </remarks>
Public Class cCommunicateEventArgs
    Inherits EventArgs

#Region "プロパティ"

    ''' <summary>オリジナルの受信文字列</summary>
    Public Property OrgReadText As String

    ''' <summary>オリジナルの受信データ(Byte型配列)</summary>
    Public Property OrgReadBytes As Byte()

    ''' <summary>受信コマンド文字列</summary>
    Public Property ReadCommand As String

    ''' <summary>受信パラメータ文字列</summary>
    Public Property ReadParam As String

    ''' <summary>オリジナルの送信文字列</summary>
    Public Property OrgWriteText As String

    ''' <summary>オリジナルの送信データ(Byte型配列)</summary>
    Public Property OrgWriteBytes As Byte()

    ''' <summary>送信コマンド文字列</summary>
    Public Property WriteCommand As String

    ''' <summary>送信パラメータ文字列</summary>
    Public Property WriteParam As String

#End Region

End Class

#End Region

#End Region

#Region "コマンド処理関連"

#Region "スレッド・分岐関連"

''' <summary>コマンド処理を行うクラス</summary>
Public Class cCommand

#Region "内部メンバ"

    '''<summary>上位から渡されるグローバル設定値</summary>
    Private _設定値 As c設定値

    ''' <summary>コマンド情報関連引数オブジェクト</summary>
    Private _Command As cCommunicateEventArgs

    ''' <summary>コマンド処理後に送信を行うタスクリストオブジェクト</summary>
    Private _SendTaskList As cSendTaskList

#End Region

#Region "プロパティ"

    ''' <summary>ソケット関連引数オブジェクト</summary>
    Public _SocketState As cSocketState

    ''' <summary>処理が完了したかどうか</summary>
    Private _Complete As Boolean = False
    Public ReadOnly Property Complete As Boolean
        Get
            Return _Complete
        End Get
    End Property

    ''' <summary>処理が成功したかどうか</summary>
    Private _Success As Boolean = False
    Public ReadOnly Property Success As Boolean
        Get
            Return _Success
        End Get
    End Property

    ''' <summary>処理開始の日時</summary>
    Private _StartDateTime As Date = Date.MinValue
    Public ReadOnly Property StartDateTime As Date
        Get
            Return _StartDateTime
        End Get
    End Property

    ''' <summary>処理終了の日時</summary>
    Private _EndDateTime As Date = Date.MaxValue
    Public ReadOnly Property EndDateTime As Date
        Get
            Return _EndDateTime
        End Get
    End Property

    ''' <summary>スレッドプールに追加された日時</summary>
    Public _AddDateTime As Date = Date.MaxValue

#End Region

#Region "イベント"

    ''' <summary                  >コンストラクタ                                        </summary>
    ''' <param name="Arg設定値"   >グローバル設定値                                      </param  >
    ''' <param name="SocketState" >送信に用いるソケット関連引数                          </param  >
    ''' <param name="Command"     >コマンド処理に用いるコマンド関連引数                  </param  >
    ''' <param name="SendTaskList">コマンド処理後に送信を行う送信タスクリストオブジェクト</param  >
    Public Sub New(Arg設定値 As c設定値, ByRef SocketState As cSocketState, ByRef Command As cCommunicateEventArgs, ByRef SendTaskList As cSendTaskList)
        Try
            _設定値 = Arg設定値           '設定値
            _SocketState = SocketState    '
            _Command = Command
            _SendTaskList = SendTaskList
            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "コマンド処理クラスのコンストラクタで例外が発生しました。", 2, _設定値.LogPath)
            Return
        End Try
    End Sub

    ''' <summary>デストラクタ</summary>
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#End Region

#Region "パブリックメソッド"

    ''' <summary>コマンド処理を実行する</summary>
    ''' <remarks>スレッド・タスクで実行する関係でこの形をとっています。</remarks>
    Public Sub CommandDone()
        Try
            _Complete = False   '完了フラグ
            _Success = False    '成功フラグ

            Dim tmpCommand As String() = distributeFunc(_Command.ReadCommand, _Command.ReadParam)   'コマンド分岐

            ''レスポンスの置換
            '_Command.WriteCommand = tmpCommand(0)   'コマンド
            '_Command.WriteParam = tmpCommand(1)     'パラメータ

            '_SendTaskList.AddSend(_SocketState, _Command, False) '送信タスクリストへ送る

            _Success = True
            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "コマンド処理開始で例外が発生しました。", 2, _設定値.LogPath)
            Return
        Finally
            _Complete = True
        End Try
    End Sub

#End Region

#Region "内部関数"

    ''' <summary              >コマンド毎に処理を割振る              </summary>
    ''' <param name="CmdCode" >コマンド文字列(2文字)                 </param>
    ''' <param name="Param"   >パラメータ文字列                      </param>
    ''' <returns              >返答送信文字列                        </returns>
    ''' <remarks              >受信条件のチェックはここでは行わない。</remarks>
    Private Function distributeFunc(CmdCode As String, Param As String) As String()
        Dim Dest(1) As String   '返答用の文字列配列(0:コマンド,1:パラメータ)
        Try
            Dim ConnectText As String = _設定値.ConnectText
            'コマンド条件の最低限のチェックは呼び出し元で行っているのでここではしない
            Dim ResCmd As String = ""       '返答コマンド
            Dim ResParam As String = ""     '返答パラメータ(JSON)
            Select Case CmdCode
                Case "10", "11"    '盤タブレット設置状態を送信
                    Dim oParam As New cCom10        'JSON変換が必要無い時
                    oParam.SetPrivateProperty(Param, _設定値, _SendTaskList, _SocketState)
                    oParam.DoCom()                 'コマンド処理を実行
                    ResCmd = oParam.GetSendCom     'コマンド
                    ResParam = oParam.GetSendParam 'パラメータ
                    oParam = Nothing
                Case "12"           '手順書データを送信
                    Dim oParam As New cCom12        'JSON変換が必要無い時
                    oParam.SetPrivateProperty(Param, _設定値, _SendTaskList, _SocketState)
                    oParam.DoCom()                 'コマンド処理を実行
                    ResCmd = oParam.GetSendCom     'コマンド
                    ResParam = oParam.GetSendParam 'パラメータ
                Case "13"   '指示済とする。確認者・盤へのステータス更新命令を送信
                    Dim oParam As cCom13 = JsonConvert.DeserializeObject(Of cCom13)(Param)  '受信データをデシリアライズする
                    oParam.SetPrivateProperty(Param, _設定値, _SendTaskList, _SocketState)
                    oParam.DoCom()                 'コマンド処理を実行
                    ResCmd = oParam.GetSendCom     'コマンド
                    ResParam = oParam.GetSendParam 'パラメータ
                Case "14"   '現場差異指示済とする。確認者へのステータス更新命令を送信
                    Dim oParam As cCom14 = JsonConvert.DeserializeObject(Of cCom14)(Param)  '受信データをデシリアライズする
                    oParam.SetPrivateProperty(Param, _設定値, _SendTaskList, _SocketState)
                    oParam.DoCom()                 'コマンド処理を実行
                    ResCmd = oParam.GetSendCom     'コマンド
                    ResParam = oParam.GetSendParam 'パラメータ
                Case "16"   '終了処理。全タブレットへ画面OFF命令を送信
                    Dim oParam As New cCom16        'JSON変換が必要無い時
                    oParam.SetPrivateProperty(Param, _設定値, _SendTaskList, _SocketState)
                    oParam.DoCom()                 'コマンド処理を実行
                    ResCmd = oParam.GetSendCom     'コマンド
                    ResParam = oParam.GetSendParam 'パラメータ
                Case "17"   '開始画面から手順書画面への移行命令
                    Dim oParam As New cCom17        'JSON変換が必要無い時
                    oParam.SetPrivateProperty(Param, _設定値, _SendTaskList, _SocketState)
                    oParam.DoCom()                 'コマンド処理を実行
                    ResCmd = oParam.GetSendCom     'コマンド
                    ResParam = oParam.GetSendParam 'パラメータ
                Case "20"   '初期画面を判断し、指示情報を送信
                    Dim oParam As New cCom20        'JSON変換が必要無い時
                    oParam.SetPrivateProperty(Param, _設定値, _SendTaskList, _SocketState)
                    oParam.DoCom()                 'コマンド処理を実行
                    ResCmd = oParam.GetSendCom     'コマンド
                    ResParam = oParam.GetSendParam 'パラメータ
                Case "21"   '手順書データを配列を送信
                    Dim oParam As New cCom21        'JSON変換が必要無い時
                    oParam.SetPrivateProperty(Param, _設定値, _SendTaskList, _SocketState)
                    oParam.DoCom()                 'コマンド処理を実行
                    ResCmd = oParam.GetSendCom     'コマンド
                    ResParam = oParam.GetSendParam 'パラメータ
                Case "22"   '確認済みとする。指示者・確認者・盤へのステータス更新命令を送信
                    Dim oParam As cCom22 = JsonConvert.DeserializeObject(Of cCom22)(Param)  '受信データをデシリアライズする
                    oParam.SetPrivateProperty(Param, _設定値, _SendTaskList, _SocketState)
                    oParam.DoCom()                 'コマンド処理を実行
                    ResCmd = oParam.GetSendCom     'コマンド
                    ResParam = oParam.GetSendParam 'パラメータ
                Case "23"   '現場差異指示済みとする。確認者へのステータス更新命令を送信
                    Dim oParam As cCom23 = JsonConvert.DeserializeObject(Of cCom23)(Param)  '受信データをデシリアライズする
                    'Dim oParam As New cCom23        'JSON変換が必要無い時
                    oParam.SetPrivateProperty(Param, _設定値, _SendTaskList, _SocketState)
                    oParam.DoCom()                 'コマンド処理を実行
                    ResCmd = oParam.GetSendCom     'コマンド
                    ResParam = oParam.GetSendParam 'パラメータ
                Case "30"   '盤情報の一覧を返す
                    Dim oParam As New cCom30        'JSON変換が必要無い時
                    oParam.SetPrivateProperty(Param, _設定値, _SendTaskList, _SocketState)
                    oParam.DoCom()                 'コマンド処理を実行
                    ResCmd = oParam.GetSendCom     'コマンド
                    ResParam = oParam.GetSendParam 'パラメータ
                Case "31"   '盤の現状を返す
                    Dim oParam As cCom31 = JsonConvert.DeserializeObject(Of cCom31)(Param)  '受信データをデシリアライズする
                    oParam.SetPrivateProperty(Param, _設定値, _SendTaskList, _SocketState)
                    oParam.DoCom()                 'コマンド処理を実行
                    ResCmd = oParam.GetSendCom     'コマンド
                    ResParam = oParam.GetSendParam 'パラメータ
                Case "32"   '盤の再選択を行う
                    Dim oParam As New cCom32
                    oParam.SetPrivateProperty(Param, _設定値, _SendTaskList, _SocketState)
                    oParam.DoCom()                 'コマンド処理を実行
                    ResCmd = oParam.GetSendCom     'コマンド
                    ResParam = oParam.GetSendParam 'パラメータ
                Case "90"   'コマンド再送要求
                    Dim oParam As New cCom90       'JSON変換が必要無い時
                    oParam.SetPrivateProperty(Param, _設定値, _SendTaskList, _SocketState)
                    oParam.DoCom()                 'コマンド処理を実行
                    ResCmd = oParam.GetSendCom     'コマンド
                    ResParam = oParam.GetSendParam 'パラメータ
                Case Else   'どのコマンドでも無い時
                    'エラーコードを返す
                    ResCmd = "5Q"       'コマンド
                    ResParam = ""       'パラメータ
            End Select

            '返答文字列配列に渡す
            Dest(0) = ResCmd        'コマンド
            Dest(1) = ResParam      'パラメータ
            Return Dest             '返す
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "コマンド分岐で例外が発生しました。", 2, _設定値.LogPath)
            Dest(0) = "5Q"    'コマンド
            Dest(1) = ""      'パラメータ
            Return Dest
        End Try
    End Function

#End Region

End Class

''' <summary>コマンド処理のスレッドプールを管理するクラス</summary>
Public Class cCommandPool

#Region "内部メンバ"

    '''<summary>上位から渡されるグローバル設定値</summary>
    Private _設定値 As c設定値

    ''' <summary>タスクを管理するList</summary>
    Private _CommandList As List(Of cCommandTask)

    ''' <summary>'送信用タスクリストクラス</summary>
    Private _SendTaskList As cSendTaskList

#End Region

#Region "イベント"

    ''' <summary                  >コンストラクタ                                        </summary>
    ''' <param name="Arg設定値"   >グローバル設定値                                      </param  >
    ''' <param name="SendTaskList">コマンド処理後に送信を行う送信タスクリストオブジェクト</param  >
    Public Sub New(Arg設定値 As c設定値, ByRef SendTaskList As cSendTaskList)
        Try
            _設定値 = Arg設定値                         '設定値
            _SendTaskList = SendTaskList                '送信タスクリスト
            _CommandList = New List(Of cCommandTask)    'インスタンス作成
        Catch ex As Exception
            ErrMsgBox(ex, Arg設定値.DebugMode, "コマンドタスクリストのコンストラクタで例外が発生しました。", 2, Arg設定値.LogPath)
            Return
        End Try
    End Sub

    ''' <summary>デストラクタ</summary>
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#End Region

#Region "パブリックメソッド"

    ''' <summary                 >コマンド処理をスレッドプールに登録して実行する</summary>
    ''' <param name="SocketState">送信に用いるソケット関連引数                  </param  >
    ''' <param name="Command"    >コマンド処理に用いるコマンド関連引数          </param  >
    Public Sub AddSend(ByRef SocketState As cSocketState, ByRef Command As cCommunicateEventArgs)
        Try
            'タスクと受信クラスを作成
            Dim AddCommandTask As New cCommandTask(_設定値, SocketState, Command, _SendTaskList)
            _CommandList.Add(AddCommandTask)   'リストに登録

            'ここで処理の終わった受信インスタンスを解除する
            For Index As Long = _CommandList.Count - 1 To 0 Step -1
                Dim DeleteFlag As Boolean = False   '削除フラグ
                With _CommandList(Index).TaskObj
                    If .IsCompleted Then DeleteFlag = True 'タスク側で完全終了していれば削除対象
                End With
                With _CommandList(Index).Command
                    '削除するかどうかを調べる
                    If .Complete Then   '終了時
                        If .Success Then    '成功終了時は削除
                            DeleteFlag = True
                        Else                '失敗で終了時は削除（再送する？）
                            DeleteFlag = True
                        End If
                    Else    '未完了の時
                    End If
                End With

                If DeleteFlag Then  '削除対象なら削除する
                    '通信を切断する
                    If Not (_CommandList(Index) Is Nothing) Then
                        If Not (_CommandList(Index).Command Is Nothing) Then
                            If Not (_CommandList(Index).Command._SocketState Is Nothing) Then
                                If Not (_CommandList(Index).Command._SocketState.WorkSocket Is Nothing) Then
                                    'If _CommandList(Index).Command._SocketState.WorkSocket.Connected Then
                                    If IsConnecting(_CommandList(Index).Command._SocketState.WorkSocket) Then
                                        _CommandList(Index).Command._SocketState.WorkSocket.Shutdown(SocketShutdown.Both)
                                    End If
                                    _CommandList(Index).Command._SocketState.WorkSocket.Close()
                                    _CommandList(Index).Command._SocketState.WorkSocket.Dispose()
                                End If
                            End If
                        End If
                    End If
                    _CommandList.Remove(_CommandList(Index))  'リストから削除する
                End If
            Next
            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "コマンドタスクリストへの登録・実行で例外が発生しました。", 2, _設定値.LogPath)
            Return
        End Try
    End Sub

#End Region

    ''' <summary>タスクと処理を配列で管理する為のクラス</summary>
    Private Class cCommandTask

#Region "内部メンバ"

        ''' <summary>上位から渡される設定値</summary>
        Private _設定値 As c設定値

#End Region

#Region "プロパティ"

        ''' <summary>コマンドを処理するクラス</summary>
        Private _Command As cCommand  'コマンドクラス
        Public ReadOnly Property Command As cCommand
            Get
                Return _Command
            End Get
        End Property

        ''' <summary>コマンド処理を別スレッドで処理するタスク</summary>
        Private _Task As Task         'タスククラス
        Public ReadOnly Property TaskObj As Task
            Get
                Return _Task
            End Get
        End Property

#End Region

#Region "イベント"

        ''' <summary                  >コンストラクタ                                          </summary>
        ''' <param name="Arg設定値"   >上位から渡される設定値                                  </param  >
        ''' <param name="SocketState" >コマンド処理後の送信に用いるソケット関連引数            </param  >
        ''' <param name="Command"     >コマンド処理に用いるコマンド関連引数                    </param  >
        ''' <param name="SendTaskList">コマンド処理後に送信を行う送信タスクリストのインスタンス</param  >
        ''' <remarks                  >インスタンス作成と同時にコマンド処理を開始する          </remarks>
        Public Sub New(Arg設定値 As c設定値, ByRef SocketState As cSocketState, ByRef Command As cCommunicateEventArgs, ByRef SendTaskList As cSendTaskList)
            Try
                _設定値 = Arg設定値           '設定値
                _Command = New cCommand(_設定値, SocketState, Command, SendTaskList)   'コマンド処理のインスタンスを作成
                _Command._AddDateTime = Now                      '現在日時を登録
                _Task = New Task(AddressOf _Command.CommandDone) 'タスクに処理を登録
                _Task.Start()                                    'さっさと開始
                Return
            Catch ex As Exception
                ErrMsgBox(ex, _設定値.DebugMode, "コマンドタスクのコンストラクタで例外が発生しました。", 2, _設定値.LogPath)
                Return
            End Try
        End Sub

        ''' <summary>デストラクタ</summary>
        Protected Overrides Sub Finalize()
            Try
                _Task.Dispose()
                Return
            Catch ex As Exception
                ErrMsgBox(ex, _設定値.DebugMode, "コマンドタスクのデストラクタで例外が発生しました。", 2, _設定値.LogPath)
                Return
            Finally
                MyBase.Finalize()
            End Try
        End Sub

#End Region

    End Class

End Class

#End Region

#Region "各コマンド処理クラス"

'JSON変換に関するメモ

''JSON書式宣言クラスに関する調査結果
''今回、コマンドパラメータやその結果返答にJSONを用い、その型宣言にクラスが必要だが、
''コマンド処理等もこのクラスに含めたい。
''その際、JSON処理と実際のコマンドとの間にどのような制限が発生するかを調査した。
' ''実験結果
' ''クラス⇒JSON変換
' ''   Public Propertyのみが対象。
' ''   Private変数、Public関数等は変換対象ではない
' ''   クラス内に自クラス⇒JSON変換変換を置ける
' ''JSON⇒クラス変換
' ''   ①最初にNewが実行されてる
' ''   ②New実行の後でJSONの値が代入される
' ''   なのでNewにコマンド処理を置けない。
' ''   またクラス内にJSON⇒自クラス変換を置けない。
''以上の事が分かった。



'変換用の型を宣言

''' <summary>コマンド処理クラスの基底クラス</summary>
''' <remarks>
'''     共通して使う命令をここで実装する
'''     そもそもJSON文字列⇒値変換用のクラス。
'''     ●使い方
'''       ①宣言
'''       ②必要に応じて外部でJSON⇒値変換を行いパブリック変数を代入
'''       ③必要なパラメータ類を代入（DB接続文字列等）
'''       ③DoCom関数を実行して返信コマンド・パラメータを作成
'''       ④外部からGetSendCom関数とGetSendParam関数を用いて返信コマンド、返信パラメータを取得する
''' </remarks>
Public Class cComBase

#Region "内部メンバ"

#Region "受信情報"

    '''<summary>受信したパラメータ文字列</summary>
    Protected _SrcParam As String

    ''' <summary>ソケット関連引数オブジェクト</summary>
    Protected _SocketState As cSocketState

#End Region

#Region "返信情報"

    '''<summary>返信するコマンド文字列</summary>
    Protected _SendCom As String

    '''<summary>返信するパラメータ文字列</summary>
    Protected _SendParam As String

#End Region

#Region "その他インスタンス"


    ''' <summary>上位から渡される設定値</summary>
    Protected _設定値 As c設定値

    ''' <summary>DB操作のインスタンス</summary>
    Protected DB As cDB

    ''' <summary>コマンド処理が終わった後の送信用タスクリストインスタンス</summary>
    Protected _SendTaskList As cSendTaskList

#End Region

#Region "定数"

    '''<summary>返信するパラメータ文字列(エラー時)</summary>
    Protected _ErrCom As String

    '''<summary>返信するパラメータ文字列(無視・問題無時)</summary>
    Protected _ThroughCom As String

    '''<summary>現場差異処理のスキップ時の表示文字列</summary>
    Protected _GS_TBL_5 As String

    '''<summary>現場差異処理の追加時の表示文字列</summary>
    Protected _GS_TBL_6 As String

    '''<summary>確認者タブレットで、固定件数表示の件数</summary>
    Protected _SLINES As Long

    '''<summary>確認者タブレットで、何手順前から表示するか</summary>
    Protected _SLINES_BEFORE As Long

#End Region

#End Region

#Region "イベント"

    ''' <summary>コンストラクタ</summary>
    ''' <remarks>              </remarks>
    Public Sub New()
        'ここは設定を得る前なのでログ出力しない
        MyBase.New()
        _SrcParam = ""              '受信したパラメータ文字列を確保
        _ErrCom = "5Q"              '返信するパラメータ文字列(エラー時)
        _ThroughCom = "50"          '返信するパラメータ文字列(無視・問題無時)
        _SendCom = _ErrCom          '返信するコマンド文字列(初期値はエラー状態)
        _SendParam = ""             '返信するパラメータ文字列
        _GS_TBL_5 = "スキップ"      '現場差異処理のスキップ時の表示文字列
        _GS_TBL_6 = "追加"          '現場差異処理の追加時の表示文字列
        _SLINES = 10                '確認者タブレットで、固定件数表示の件数
        _SLINES_BEFORE = 0          '確認者タブレットで、何手順前から表示するか
        Return
    End Sub

#End Region

    ''' <summary>
    '''     プライマリ変数への代入を行う
    ''' </summary>
    ''' <remarks>
    '''     プロパティではJSON変換の対象となり、
    '''     コンストラクタ渡しではJSON変換時に実態化されない為。
    ''' </remarks>
    Public Sub SetPrivateProperty(SrcParam As String, ByRef 設定値 As c設定値, ByRef SendTaskList As cSendTaskList, ByRef SocketState As cSocketState)
        Try
            _SrcParam = SrcParam         '受信したパラメータ文字列を確保
            _設定値 = 設定値
            _SendTaskList = SendTaskList
            _SocketState = SocketState
            DB = New cDB(_設定値)
            Return
        Catch ex As Exception
            ErrMsgBox(ex, 設定値.DebugMode, "コマンド基底クラスの内部メンバへの各値登録で例外が発生しました。", 2, 設定値.LogPath)
            Return
        End Try
    End Sub

    ''' <summary>
    '''     コマンドを実行する。
    '''     この関数は派生クラス側で上書きする事。
    '''     派生クラスのこのコマンドで_SendComと_SendParamを更新する事。
    ''' </summary>
    ''' <remarks>  </remarks>
    Public Overridable Sub DoCom()

    End Sub

#Region "エクセル関連"

    ''' <summary>報告書ファイルのフルパスを得る</summary>
    ''' <returns>報告書のフルパス</returns>
    ''' <remarks>失敗した時は空の文字列を返す</remarks>
    Private Function GetReportFilePath() As String
        Try
            '元の手順書ファイルのファイル名を取得する
            Dim SrcFileName As String = GetTxSFile()
            If "" = SrcFileName Then Return "" 'ファイル名取得失敗時は空文字を返す
            Dim DirName As String = SrcFileName.Split(".")(0)
            Dim DestFilePath As String = System.IO.Path.Combine(_設定値.ReportDirPath, DirName)
            DestFilePath = System.IO.Path.Combine(DestFilePath, _設定値.ReportFileName)
            Return DestFilePath
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "報告書ファイルパス取得時に例外が発生しました。", 2, _設定値.LogPath)
            Return ""
        End Try
    End Function

    ''' <summary>DBから手順書ファイルのファイル名を得る</summary>
    ''' <returns>報告書のフルパス</returns>
    ''' <remarks>失敗した時は空の文字列を返す</remarks>
    Private Function GetTxSFile() As String
        Try
            '元の手順書ファイルのファイル名を取得する
            '現状を確認する
            Dim Sql1 As String = "SELECT ""tx_sfile"" FROM ""t_sno"""
            Dim dt1 As DataTable = GetDT_Select(Sql1)
            If 0 = dt1.Rows.Count Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "手順書ファイル名取得時にクエリ「" & Sql1 & "」にヒットしませんでした。", 1, _設定値.LogPath)
                Return ""
            End If
            Return dt1.Rows(0)("tx_sfile").ToString
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "手順書ファイル名取得時に例外が発生しました。", 2, _設定値.LogPath)
            Return ""
        End Try
    End Function

    ''' <summary>レポートファイルに値を記入する</summary>
    ''' <param name="RowIndex">対象の行位置(1から開始)</param>
    ''' <param name="ColIndex">対象の列位置(1から開始)</param>
    ''' <param name="WriteValue">記述する文字列</param>
    ''' <returns>処理の正否</returns>
    Protected Function WriteReport(RowIndex As Integer, ColIndex As Integer, WriteValue As String) As Boolean
        Try
            '書き込み先のファイルパスを得る
            '元の手順書のファイルパスはDBから取得する必要がある
            Dim FilePath As String = GetReportFilePath()
            If "" = FilePath Then Return False '記述先が取得できなかったときは抜ける(エラーログ記述済み)
            Return WriteExcellCell(FilePath, _設定値.ReportSheetName, RowIndex, ColIndex, WriteValue)    'ファイルに記述する
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "レポートファイルへの記入時に例外が発生しました。", 2, _設定値.LogPath)
            Return False
        End Try
    End Function

    ''' <summary>指定したエクセルファイルの指定したシートのセルに文字列を記述する</summary>
    ''' <param name="FilePath">対象のファイルパス</param>
    ''' <param name="SheetName">対象のシート名</param>
    ''' <param name="RowIndex">対象の行位置(1から開始)</param>
    ''' <param name="ColIndex">対象の列位置(1から開始)</param>
    ''' <param name="WriteValue">記述する文字列</param>
    ''' <returns>処理の正否</returns>
    ''' <remarks>実行時バインドバージョン(遅延バインディング)</remarks>
    Private Function WriteExcellCell(FilePath As String, SheetName As String, RowIndex As Integer, ColIndex As Integer, WriteValue As String) As Boolean
        Dim ExcelApp As Object      'Excelオブジェクト
        Dim WorkBook As Object      'Workbookオブジェクト
        Dim WorkSheet As Object     'WorkSheetオブジェクト
        Dim TargetRange As Object   '書き込み先のRangeオブジェクト
        Try
            '起動条件を調べる
            If FilePath Is Nothing Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "エクセルファイルへの記入でファイルパスが与えられませんでした。", 1, _設定値.LogPath)
                Return False
            End If
            If "" = FilePath Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "エクセルファイルへの記入でファイルパスが空でした。", 1, _設定値.LogPath)
                Return False
            End If
            If SheetName Is Nothing Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "エクセルファイルへの記入でSheets名が与えられませんでした。", 1, _設定値.LogPath)
                Return False
            End If
            If "" = SheetName Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "エクセルファイルへの記入でSheets名が空でした。", 1, _設定値.LogPath)
                Return False
            End If
            If RowIndex < 1 Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "エクセルファイルへの記入でRowIndexが範囲内にありません。(" & RowIndex & ")", 1, _設定値.LogPath)
                Return False
            End If
            If ColIndex < 1 Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "エクセルファイルへの記入でColIndexが範囲内にありません。(" & ColIndex & ")", 1, _設定値.LogPath)
                Return False
            End If
            If Not (System.IO.File.Exists(FilePath)) Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "エクセルファイルへの記入で存在しないファイルを指定されました。(" & FilePath & ")", 1, _設定値.LogPath)
                Return False
            End If

            'ファイルを開く
            ExcelApp = CreateObject("Excel.Application")        'エクセルを起動する
            ExcelApp.Visible = False                            'エクセルは非表示する
            ExcelApp.DisplayAlerts = False                      'ダイアログを表示しない
            WorkBook = ExcelApp.Workbooks.Open(FilePath)        'ファイルを開く

            'Sheetsオブジェクトをループして対象のシートを見つける
            WorkSheet = Nothing
            For Each tmpSheet As Object In WorkBook.Sheets
                If SheetName = tmpSheet.Name Then   'ヒットしたら返す
                    WorkSheet = tmpSheet
                    Exit For
                End If
            Next
            If WorkSheet Is Nothing Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "エクセルファイルへの記入で該当するSheetが見つかりませんでした。(シート名：" & SheetName & ")", 1, _設定値.LogPath)
                Return False
            End If

            TargetRange = WorkSheet.Cells(RowIndex, ColIndex)   '該当セルを得る
            TargetRange.Value = WriteValue                      '値を記入
            WorkBook.Save()                                     '保存する

            Return True
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "エクセルファイルへの記入時に例外が発生しました。", 2, _設定値.LogPath)
            Return False
        Finally
            Try
                If Not (WorkBook Is Nothing) Then WorkBook.Close() '閉じる
                If Not (ExcelApp Is Nothing) Then ExcelApp.Quit() '終了する
            Catch ex As Exception
            End Try

            '解放
            MRComObject(TargetRange)
            MRComObject(WorkSheet)
            MRComObject(WorkBook)
            MRComObject(ExcelApp)
        End Try
    End Function

    ' ''' <summary>指定したエクセルファイルの指定したシートのセルに文字列を記述する</summary>
    ' ''' <param name="FilePath">対象のファイルパス</param>
    ' ''' <param name="SheetName">対象のシート名</param>
    ' ''' <param name="RowIndex">対象の行位置(1から開始)</param>
    ' ''' <param name="ColIndex">対象の列位置(1から開始)</param>
    ' ''' <param name="WriteValue">記述する文字列</param>
    ' ''' <returns>処理の正否</returns>
    ' ''' <remarks>エクセル参照バージョン(事前バインディング)</remarks>
    'Protected Function WriteExcellCell_old(FilePath As String, SheetName As String, RowIndex As Integer, ColIndex As Integer, WriteValue As String) As Boolean
    '    Dim ExcelApp As Excel.Application   'Excelオブジェクト
    '    Dim WorkBook As Excel.Workbook      'Workbookオブジェクト
    '    Dim WorkSheet As Excel.Worksheet    'WorkSheetオブジェクト
    '    Dim TargetRange As Excel.Range      '書き込み先のRangeオブジェクト
    '    Try
    '        '起動条件を調べる
    '        If FilePath Is Nothing Then
    '            ErrMsgBox(Nothing, _設定値.DebugMode, "エクセルファイルへの記入でファイルパスが与えられませんでした。", 1, _設定値.LogPath)
    '            Return False
    '        End If
    '        If "" = FilePath Then
    '            ErrMsgBox(Nothing, _設定値.DebugMode, "エクセルファイルへの記入でファイルパスが空でした。", 1, _設定値.LogPath)
    '            Return False
    '        End If
    '        If SheetName Is Nothing Then
    '            ErrMsgBox(Nothing, _設定値.DebugMode, "エクセルファイルへの記入でSheets名が与えられませんでした。", 1, _設定値.LogPath)
    '            Return False
    '        End If
    '        If "" = SheetName Then
    '            ErrMsgBox(Nothing, _設定値.DebugMode, "エクセルファイルへの記入でSheets名が空でした。", 1, _設定値.LogPath)
    '            Return False
    '        End If
    '        If RowIndex < 1 Then
    '            ErrMsgBox(Nothing, _設定値.DebugMode, "エクセルファイルへの記入でRowIndexが範囲内にありません。(" & RowIndex & ")", 1, _設定値.LogPath)
    '            Return False
    '        End If
    '        If ColIndex < 1 Then
    '            ErrMsgBox(Nothing, _設定値.DebugMode, "エクセルファイルへの記入でColIndexが範囲内にありません。(" & ColIndex & ")", 1, _設定値.LogPath)
    '            Return False
    '        End If
    '        If Not (System.IO.File.Exists(FilePath)) Then
    '            ErrMsgBox(Nothing, _設定値.DebugMode, "エクセルファイルへの記入で存在しないファイルを指定されました。(" & FilePath & ")", 1, _設定値.LogPath)
    '            Return False
    '        End If

    '        'ファイルを開く
    '        ExcelApp = New Excel.Application()  'エクセルを起動する
    '        ExcelApp.Visible = False            'エクセルは非表示する
    '        ExcelApp.DisplayAlerts = False      'ダイアログを表示しない
    '        WorkBook = DirectCast(ExcelApp.Workbooks.Open(FilePath), Excel.Workbook)    'ファイルを開く
    '        WorkSheet = GetSheet(WorkBook.Sheets, SheetName)                            '該当シートを得る
    '        TargetRange = DirectCast(WorkSheet.Cells(RowIndex, ColIndex), Excel.Range)  '該当セルを得る
    '        TargetRange.Value = WriteValue                                              '値を記入
    '        WorkBook.Save()     '保存する

    '        Return True
    '    Catch ex As Exception
    '        ErrMsgBox(ex, _設定値.DebugMode, "エクセルファイルへの記入時に例外が発生しました。", 2, _設定値.LogPath)
    '        Return False
    '    Finally
    '        If Not (WorkBook Is Nothing) Then WorkBook.Close() '閉じる
    '        If Not (ExcelApp Is Nothing) Then ExcelApp.Quit() '終了する

    '        '解放
    '        MRComObject(TargetRange)
    '        MRComObject(WorkSheet)
    '        MRComObject(WorkBook)
    '        MRComObject(ExcelApp)
    '    End Try
    'End Function

    ' ''' <summary>ワークシート名から該当するエクセルワークシートを返す</summary>
    ' ''' <param name="SrcSheets">対象のExcel.Sheetsオブジェクト</param>
    ' ''' <param name="sheetName">検索するワークシート名</param>
    ' ''' <returns>ヒットしたワークシート</returns>
    ' ''' <remarks>参照バージョン(事前バインディング)</remarks>
    'Private Function GetSheet_old(ByRef SrcSheets As Excel.Sheets, ByVal SheetName As String) As Excel.Worksheet
    '    Try
    '        '動作条件をチェック
    '        If SrcSheets Is Nothing Then
    '            ErrMsgBox(Nothing, _設定値.DebugMode, "エクセルのワークシート検索でSheetsオブジェクトが与えられませんでした。", 1, _設定値.LogPath)
    '            Return Nothing
    '        End If
    '        If SheetName Is Nothing Then
    '            ErrMsgBox(Nothing, _設定値.DebugMode, "エクセルのワークシート検索でSheets名が与えられませんでした。", 1, _設定値.LogPath)
    '            Return Nothing
    '        End If
    '        If "" = SheetName Then
    '            ErrMsgBox(Nothing, _設定値.DebugMode, "エクセルのワークシート検索でSheets名が空でした。", 1, _設定値.LogPath)
    '            Return Nothing
    '        End If

    '        'Sheetsオブジェクトをループする
    '        For Each SrcSheet As Microsoft.Office.Interop.Excel.Worksheet In SrcSheets
    '            If SheetName = SrcSheet.Name Then Return SrcSheet 'ヒットしたら返す
    '        Next
    '        Return Nothing  'ヒットしなかったときは抜ける
    '    Catch ex As Exception
    '        ErrMsgBox(ex, _設定値.DebugMode, "エクセルのワークシート検索時に例外が発生しました。", 2, _設定値.LogPath)
    '        Return Nothing
    '    End Try
    'End Function

#End Region

#Region "dump出力関連"

    ''' <summary>Dumpファイルのフルパスを得る</summary>
    ''' <returns>報告書のフルパス</returns>
    ''' <remarks>失敗した時は空の文字列を返す</remarks>
    Private Function GetDumpFilePath() As String
        Try
            '元の手順書ファイルのファイル名を取得する
            Dim SrcFileName As String = GetTxSFile()
            Dim DirName As String = SrcFileName.Split(".")(0)
            Dim DestFilePath As String = System.IO.Path.Combine(_設定値.ReportDirPath, DirName)
            DestFilePath = System.IO.Path.Combine(DestFilePath, _設定値.DumpFileName)
            Return DestFilePath
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "Dumpファイルパス取得時に例外が発生しました。", 2, _設定値.LogPath)
            Return ""
        End Try
    End Function

    ''' <summary>DBのバックアップ(dump)を行う</summary>
    ''' <returns>処理の正否</returns>
    Protected Function DoneDump() As Boolean
        Try
            '書き込み先のファイルパスを得る
            '元の手順書のファイルパスはDBから取得する必要がある
            Dim FilePath As String = GetDumpFilePath()
            If "" = FilePath Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "Dump実行時に出力先ファイルパスの取得ができませんでした。", 1, _設定値.LogPath)
                Return False
            End If

            'Dumpを実行
            Return DoneDump(_設定値.DumpCmd, FilePath)
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "Dump実行時に例外が発生しました。", 2, _設定値.LogPath)
            Return False
        End Try
    End Function

    ''' <summary>DBのバックアップ(dump)を行う</summary>
    ''' <returns>処理の正否</returns>
    Public Function DoneDump(CmdText As String, WriteFilePath As String) As Boolean
        Try
            'Dumpを実行
            Dim psInfo As New ProcessStartInfo()
            psInfo.FileName = "cmd.exe"
            psInfo.Arguments = CmdText & " > " & WriteFilePath
            Process.Start(psInfo)

            ErrMsgBox(Nothing, _設定値.DebugMode, "DBのバックアップ(dump)を行いました。", 0, _設定値.LogPath)
            Return True
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "Dump実行時に例外が発生しました。", 2, _設定値.LogPath)
            Return False
        End Try
    End Function

#End Region

#Region "DB操作関連"

    ''' <summary>返信するコマンド文字列を得る</summary>
    ''' <returns>返信するコマンド文字列      </returns>
    ''' <remarks>コマンド処理関数実行後に有効</remarks>
    Public Function GetSendCom() As String
        Return _SendCom
    End Function

    ''' <summary>返信するパラメータ文字列を得る</summary>
    ''' <returns>返信するパラメータ文字列      </returns>
    ''' <remarks>コマンド処理関数実行後に有効  </remarks>
    Public Function GetSendParam() As String
        Return _SendParam
    End Function

    ''' <summary         >DBへSELECTクエリを出して結果のDataTableを得る</summary>
    ''' <param name="SQL">DBへ要求するSELECT文                         </param>
    ''' <remarks         >cDBクラスへ移管した際の残骸ラップ関数        </remarks>
    Protected Function GetDT_Select(SQL As String) As DataTable
        Return DB.GetDT_Select(SQL)
    End Function

    ''' <summary         >DBへSELECTコマンド出して結果の値(レコード件数等)を得る</summary>
    ''' <param name="SQL">DBへ要求するクエリ文                                  </param>
    ''' <remarks         >cDBクラスへ移管した際の残骸ラップ関数    　　　　　   </remarks>
    Protected Function DBNonQuery(SQL As String) As Integer
        Return DB.DBNonQuery(SQL)
    End Function

    ''' <summary>DBへトランザクションのBEGINを要求し、Transakutionオブジェクトを得る</summary>
    Protected Sub DB_BeginTransaction(ByRef Transaction As Npgsql.NpgsqlTransaction)
        DB.DB_BeginTransaction(Transaction)
        Return 'DB.DB_BeginTransaction
    End Sub

    ''' <summary>DB操作関連のクラス</summary>
    ''' <remarks></remarks>
    Protected Class cDB

#Region "内部メンバ"

        ''' <summary>外部から与えられた設定値</summary>
        Private _設定値 As c設定値

#End Region

#Region "イベント"

        ''' <summary>コンストラクタ</summary>
        ''' <remarks>              </remarks>
        Public Sub New(設定値 As c設定値)
            MyBase.New()
            Try
                _設定値 = 設定値
                Return
            Catch ex As Exception
                ErrMsgBox(ex, _設定値.DebugMode, "DBクラスのコンストラクタで例外が発生しました。", 2, _設定値.LogPath)
                Return
            End Try
        End Sub

#End Region

#Region "パブリックメソッド"

        ''' <summary         >DBへSELECTクエリを出して結果のDataTableを得る</summary>
        ''' <param name="SQL">DBへ要求するSELECT文                         </param>
        ''' <remarks         >                                　　　　　   </remarks>
        Public Function GetDT_Select(SQL As String) As DataTable
            Dim PGConnection As Npgsql.NpgsqlConnection 'DBコネクション
            Try
                GetPGSQLConnection(PGConnection)                        'コネクション取得
                Dim DataAdapter As Npgsql.NpgsqlDataAdapter = New Npgsql.NpgsqlDataAdapter(SQL.Trim, PGConnection)    'データアダプタを取得
                Dim DataTable As New DataTable
                DataAdapter.Fill(DataTable)   'データテーブルを作成
                Return DataTable
            Catch ex As Exception
                ErrMsgBox(ex, _設定値.DebugMode, "SQL文「" & SQL & "」で例外が発生しました。(GetDT_Select)", 2, _設定値.LogPath)
                Return Nothing
            Finally
                If Not (PGConnection Is Nothing) Then PGConnection.Close() 'コネクション開放
            End Try
        End Function

        ''' <summary         >DBへ更新クエリを出して結果の値(レコード件数等)を得る</summary>
        ''' <param name="SQL">DBへ要求するクエリ文                                </param>
        ''' <remarks         >                                         　　　　　 </remarks>
        Public Function DBNonQuery(SQL As String) As Integer
            Dim PGConnection As Npgsql.NpgsqlConnection 'DBコネクション
            Dim Command As Npgsql.NpgsqlCommand         'DBコマンド
            Try
                GetPGSQLConnection(PGConnection)                        'コネクション取得
                Command = New Npgsql.NpgsqlCommand(SQL, PGConnection)   'コマンド登録
                Return Command.ExecuteNonQuery()                        'コマンド実行
            Catch ex As Exception
                ErrMsgBox(ex, _設定値.DebugMode, "SQL文「" & SQL & "」で例外が発生しました。(GetDT_NonQuery)", 2, _設定値.LogPath)
                Return Nothing
            Finally
                If Not (PGConnection Is Nothing) Then PGConnection.Close() 'コネクション開放
                If Not (Command Is Nothing) Then Command.Dispose() 'コマンド開放
            End Try
        End Function

        ''' <summary>DBへトランザクションの開始を指示する</summary>
        ''' <remarks>                                    </remarks>
        Public Sub DB_BeginTransaction(ByRef Transaction As Npgsql.NpgsqlTransaction)
            'Public Function DB_BeginTransaction(ByRef Transaction As Npgsql.NpgsqlTransaction) As Npgsql.NpgsqlTransaction
            Dim PGConnection As Npgsql.NpgsqlConnection 'DBコネクション
            Try
                GetPGSQLConnection(PGConnection)        'コネクション取得
                Transaction = PGConnection.BeginTransaction()
                Return 'PGConnection.BeginTransaction()  'トランザクション開始
            Catch ex As Exception
                ErrMsgBox(ex, _設定値.DebugMode, "トランザクションの開始で例外が発生しました。(GetDT_NonQuery)", 2, _設定値.LogPath)
                Return 'Nothing
            Finally
                If Not (PGConnection Is Nothing) Then PGConnection.Close() 'コネクション開放
            End Try
        End Sub


#End Region

#Region "内部関数"

        ''' <summary          >DBへのコネクションを取得する    </summary>
        ''' <param name="Dest">取得したDBへコネクション応答引数</param>
        ''' <remarks          >                                </remarks>
        Protected Sub GetPGSQLConnection(ByRef Dest As Npgsql.NpgsqlConnection)
            Dim Connection As Npgsql.NpgsqlConnection
            Try
                Connection = New Npgsql.NpgsqlConnection(_設定値.ConnectText)
                Connection.Open()
                Dest = Connection
                Return
            Catch ex As Exception
                ErrMsgBox(ex, _設定値.DebugMode, "PostgreSQLとのConnection取得で例外が発生しました。", 2, _設定値.LogPath)
                Return
            Finally
                Dest = Connection
            End Try
        End Sub

#End Region

    End Class

#End Region

#Region "送信関連"

    ''' <summary>リモートアドレスを指定してサーバ発呼でコマンドを送信する</summary>
    ''' <param name="RemoteIP">送信先のIPアドレス</param>
    ''' <param name="RemotePort">送信先のポート番号</param>
    ''' <param name="Command">コマンド文字列</param>
    ''' <param name="Param">パラメータ文字列</param>
    ''' <returns>送信が正常に完了したかどうか</returns>
    Protected Function Send(RemoteIP As String, RemotePort As Integer, Command As String, Param As String) As Boolean
        Try
            'サーバ発呼クラスに必要な引数のインスタンスを用意する

            'ソケット関連のクラスインスタンス
            Dim SocketState As New cSocketState(_設定値)
            Dim Communicate As New cCommunicateEventArgs()
            With Communicate
                .WriteCommand = Command
                .WriteParam = Param
            End With

            'コマンド情報のクラスインスタンス
            Dim CommunicateEventArgs As New cCommunicateEventArgs
            With CommunicateEventArgs
                .ReadCommand = ""
                .ReadCommand = _SrcParam
                .WriteCommand = Command
                .WriteParam = Param
            End With

            '直近コマンドを保存
            Dim LastCommand As New cLastCommandList.cLastCommand
            With LastCommand
                .IP = RemoteIP.Trim
                .Command = Command
                .Param = Param
            End With
            _設定値.LastCommandList.AddLastCommand(LastCommand)

            'サーバ発呼を行う
            Dim ServerStart As New cServerStart(_設定値, CommunicateEventArgs, RemoteIP, RemotePort)   'クラスのインスタンスを作成
            ServerStart.Done()  '送信開始
            '_SendTaskList.AddSend(SocketState, Communicate, True, RemoteIP, RemotePort)
            Return ServerStart.Success
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "コマンド送信時に例外が発生しました。", 2, _設定値.LogPath)
            Return False
        End Try
    End Function

    ''' <summary>登録されている全ての指示者タブレットへコマンドを送信する</summary>
    ''' <param name="Command">コマンド文字列</param>
    ''' <param name="Param">パラメータ文字列</param>
    ''' <returns>送信が正常に完了したかどうか</returns>
    Protected Function SendTo指示者Tablet(Command As String, Param As String) As Boolean
        Try
            Dim Tablets As List(Of cTabletList.cTablet) = _設定値.TabletList.GetTabletList("10") '指示者タブレットのIPを全て得る
            For Each Tablet As cTabletList.cTablet In Tablets   '得たIPへコマンドを送る
                If Send(Tablet.IP, _設定値.RemotePort, Command, Param) Then Return True '通信が成功したらTrueを返して抜ける
            Next
            Return False
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "指示者タブレットへのコマンド送信時に例外が発生しました。", 2, _設定値.LogPath)
            Return False
        End Try
    End Function

    ''' <summary>登録されている全ての確認者タブレットへコマンドを送信する</summary>
    ''' <param name="Command">コマンド文字列</param>
    ''' <param name="Param">パラメータ文字列</param>
    ''' <returns>送信が正常に完了したかどうか</returns>
    Protected Function SendTo確認者Tablet(Command As String, Param As String) As Boolean
        Try
            '確認者タブレットのIPを得る
            Dim Tablets As List(Of cTabletList.cTablet) = _設定値.TabletList.GetTabletList("20") '確認者タブレットの一覧を得る
            For Each Tablet As cTabletList.cTablet In Tablets
                If Send(Tablet.IP, _設定値.RemotePort, Command, Param) Then Return True '通信が成功したらTrueを返して抜ける
            Next
            'Dim dt As DataTable = GetDT_Select("SELECT ""tx_rip"" FROM ""m_kakunin""")
            'If 0 = dt.Rows.Count Then Return 'データが無かった時は抜ける
            ''データがあった時は0
            'Dim RemoteIP As String = dt.Rows(0)("tx_rip").ToString
            'Send(RemoteIP, _設定値.RemotePort, Command, Param)
            Return False
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "確認者タブレットへのコマンド送信時に例外が発生しました。", 2, _設定値.LogPath)
            Return False
        End Try
    End Function

    ''' <summary>登録されていない全ての盤タブレットへコマンドを送信する</summary>
    ''' <param name="Command">コマンド文字列</param>
    ''' <param name="Param">パラメータ文字列</param>
    ''' <param name="Registed">登録済みの盤タブレットかどうか</param>
    ''' <returns>送信が正常に完了したかどうか</returns>
    Protected Function SendTo盤Tablet(Command As String, Param As String, Registed As Boolean) As Boolean
        Try
            '盤タブレットのIPを得る
            Dim Tablets As List(Of cTabletList.cTablet) = _設定値.TabletList.GetTabletList("30") '盤タブレットの一覧を得る
            For Each Tablet As cTabletList.cTablet In Tablets
                If (Registed = _設定値.TabletList.Is盤登録済みIP(Tablet.IP)) Then
                    Send(Tablet.IP, _設定値.RemotePort, Command, Param)
                End If
            Next
            Return True
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "未登録盤タブレットへのコマンド送信時に例外が発生しました。", 2, _設定値.LogPath)
            Return False
        End Try
    End Function

#End Region

#Region "タブレットシステム共通関数"

    ''' <summary            >盤タブレットの黄色と点滅をキャンセルする</summary>
    ''' <param name="In_Bno">対象の盤を示す番号                      </param>
    ''' <remarks            >旧PHP版のfunc.ini内set_ban_clear関数参照</remarks>
    Protected Sub Set_Ban_Clear(In_Bno As Long)
        Try
            If In_Bno = -1 Then Return '実行条件確認
            Dim tbname = "ban" & In_Bno.ToString
            DBNonQuery("UPDATE " & tbname & " SET ""bo_active""='f',""in_disp_hi1""=0,""in_disp_hi2""=0,""in_disp_hi3""=0,""in_disp_hi4""=0,""in_disp_hi5""=0")
            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "タブレットの表示を初期化する操作で例外が発生しました。", 2, _設定値.LogPath)
            Return
        End Try
    End Sub

    ''' <summary>次の手順がコメントなら次へ進める       </summary>
    ''' <remarks>旧PHP版のfunc.ini内skip_comment関数参照</remarks>
    Protected Function Skip_Comment() As Boolean
        Try
            '現状を確認する
            Dim Sql1 As String = "SELECT ""in_csno"" FROM ""t_sno"""
            Dim dt1 As DataTable = GetDT_Select(Sql1)
            If 0 = dt1.Rows.Count Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "コメント行のスキップに失敗しました。t_snoテーブルが空です。クエリ「" & Sql1 & "」にヒットしませんでした。", 1, _設定値.LogPath)
                Return False
            End If
            Dim in_csno As Long
            If IsDBNull(dt1.Rows(0)("in_csno")) Then    'Null対策
                ErrMsgBox(Nothing, _設定値.DebugMode, "コメント行のスキップに失敗しました。t_snoテーブルのin_csnoがNULLです。クエリは「" & Sql1 & "」です。", 1, _設定値.LogPath)
                Return False
            Else
                in_csno = dt1.Rows(0)("in_csno").ToString '要Null対策
            End If

            Dim Sql2 As String = "SELECT count(*) FROM ""t_order"""
            Dim dt2 As DataTable = GetDT_Select(Sql2)
            If 0 = dt2.Rows.Count Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "コメント行のスキップに失敗しました。t_orderテーブルのレコードカウントが行えません。クエリ「" & Sql2 & "」にヒットしませんでした。", 1, _設定値.LogPath)
                Return False
            End If
            Dim co_t_order As String = dt2.Rows(0)(0).ToString

            For I As Long = in_csno To co_t_order
                Dim Sql3 As String = "SELECT ""in_sno"",""tx_sno"" FROM ""t_order"" WHERE ""in_sno""=" & I.ToString
                Dim dt3 As DataTable = GetDT_Select(Sql3)
                Dim Rows As Long = dt3.Rows.Count
                ''以下のコメントアウトは必要無さそうなのでコメントにした。
                ''手順書データt_orderのin_snoに番号の抜けがあった時は正しくコメントスキップしない恐れがあるため。
                'If 0 = Rows Then
                '    ErrMsgBox(Nothing, _設定値.DebugMode, "コメント行のスキップに失敗しました。t_orderテーブルに存在するはずのin_snoのレコードが見つかりません。クエリ「" & Sql3 & "」にヒットしませんでした。", 1, _設定値.LogPath)
                '    Return False
                'End If
                If 1 = Rows Then
                    Dim in_sno As String = dt3.Rows(0)("in_sno").ToString
                    Dim tx_sno As String = dt3.Rows(0)("tx_sno").ToString '文字列型からNULL対策は必要
                    If "C" = tx_sno Then
                        DBNonQuery("UPDATE ""t_order"" SET ""bo_fin""='t',""cd_status""='7' WHERE ""in_sno""=" & in_sno)
                        DBNonQuery("UPDATE ""t_sno"" SET ""in_csno""=""in_csno""+1")
                    Else
                        Exit For
                    End If
                End If
            Next
            Return True
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "コメント手順行のスキップで例外が発生しました。", 2, _設定値.LogPath)
            Return False
        End Try
    End Function

    ''' <summary>盤の状態をリセットする関数だと思います    </summary>
    ''' <remarks>旧PHP版のfunc.ini内set_ban_tables2関数参照</remarks>
    Protected Sub Set_Ban_Tables2()
        Try
            '盤マスタ(m_ban)をループする
            Dim Sql盤 As String = "SELECT * FROM ""m_ban"""  '"SELECT ""in_bno"",""tx_tbl_name"" FROM ""m_ban"""
            Dim dt盤 As DataTable = GetDT_Select(Sql盤)
            If 0 = dt盤.Rows.Count Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "クエリ「" & Sql盤 & "」にヒットしませんでした。要確認", 1, _設定値.LogPath)
                Return
            End If
            'Dim in_csno As Long = dt盤.Rows(0)("in_csno")
            For Each Row盤 As DataRow In dt盤.Rows
                Dim in_bno As Long = Row盤("in_bno")
                Dim tbname As String = Row盤("tx_tbl_name").ToString

                '盤テーブルへの各値を初期化しておく
                Dim tx_lb1 As String = "" : Dim tx_clr1 As String = "" : Dim in_disp_blink1 As Long = 0
                Dim tx_lb2 As String = "" : Dim tx_clr2 As String = "" : Dim in_disp_blink2 As Long = 0
                Dim tx_lb3 As String = "" : Dim tx_clr3 As String = "" : Dim in_disp_blink3 As Long = 0
                Dim tx_lb4 As String = "" : Dim tx_clr4 As String = "" : Dim in_disp_blink4 As Long = 0
                Dim tx_lb5 As String = "" : Dim tx_clr5 As String = "" : Dim in_disp_blink5 As Long = 0

                '盤内の機器(m_device2)をループする
                Dim Sql機器 As String = "SELECT * FROM ""m_device2"" WHERE ""in_bno""=" & in_bno & " ORDER BY ""no"""
                Dim dt機器 As DataTable = GetDT_Select(Sql機器)
                If 0 = dt機器.Rows.Count Then
                    ErrMsgBox(Nothing, _設定値.DebugMode, "クエリ「" & Sql機器 & "」にヒットしませんでした。要確認", 1, _設定値.LogPath)
                Else
                    Dim CountJ As Long = 1
                    For Each Row機器 As DataRow In dt機器.Rows
                        '値の初期値を決める
                        Dim 状態 As String = Row機器("jotai_name1").ToString
                        Dim 状態色 As String = Row機器("jotai_clr1").ToString
                        '各フラグを調整する
                        If "0" = Row機器("cd_flip").ToString Then
                            状態 = Row機器("jotai_name0").ToString
                            状態色 = Row機器("jotai_clr0").ToString
                        End If
                        Dim tx_lbl As String = Row機器("swname") & vbCrLf & 状態

                        ''接地取付ならin_disp_blinkX=1,接地取外なら0を立てる
                        Dim 点滅 As Long = 0
                        Select Case Row機器("cd_flip")
                            Case "0"
                                Select Case Row機器("jotai_name0")
                                    Case "接地取付", "接地器具取付" : 点滅 = 1
                                    Case "接地取外", "接地器具取外" : 点滅 = 0
                                    Case Else : 点滅 = 0
                                End Select
                            Case "1"
                                Select Case Row機器("jotai_name1")
                                    Case "接地取付", "接地器具取付" : 点滅 = 1
                                    Case "接地取外", "接地器具取外" : 点滅 = 0
                                    Case Else : 点滅 = 0
                                End Select
                            Case Else
                                : 点滅 = 0
                        End Select

                        Select Case CountJ
                            Case 1 : tx_lb1 = tx_lbl : tx_clr1 = 状態色 : in_disp_blink1 = 点滅.ToString
                            Case 2 : tx_lb2 = tx_lbl : tx_clr2 = 状態色 : in_disp_blink2 = 点滅.ToString
                            Case 3 : tx_lb3 = tx_lbl : tx_clr3 = 状態色 : in_disp_blink3 = 点滅.ToString
                            Case 4 : tx_lb4 = tx_lbl : tx_clr4 = 状態色 : in_disp_blink4 = 点滅.ToString
                            Case 5 : tx_lb5 = tx_lbl : tx_clr5 = 状態色 : in_disp_blink5 = 点滅.ToString
                        End Select

                        CountJ += 1
                    Next

                    Dim sql As String = "UPDATE """ & tbname & """ SET "
                    sql &= """tx_lb1""='" & tx_lb1 & "',""tx_clr1""='" & tx_clr1 & "',""in_disp_blink1""=" & in_disp_blink1 & ","
                    sql &= """tx_lb2""='" & tx_lb2 & "',""tx_clr2""='" & tx_clr2 & "',""in_disp_blink2""=" & in_disp_blink2 & ","
                    sql &= """tx_lb3""='" & tx_lb3 & "',""tx_clr3""='" & tx_clr3 & "',""in_disp_blink3""=" & in_disp_blink3 & ","
                    sql &= """tx_lb4""='" & tx_lb4 & "',""tx_clr4""='" & tx_clr4 & "',""in_disp_blink4""=" & in_disp_blink4 & ","
                    sql &= """tx_lb5""='" & tx_lb5 & "',""tx_clr5""='" & tx_clr5 & "',""in_disp_blink5""=" & in_disp_blink5 & " "
                    DBNonQuery(sql)
                End If
            Next
            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "盤の状態をリセットする処理で例外が発生しました。", 2, _設定値.LogPath)
            Return
        End Try
    End Sub

#End Region

#Region "タブレットへの送信関数"

    ''' <summary>コマンド57を行う</summary>
    ''' <remarks>各タブレットの状態を纏めて返信まで行う。</remarks>
    Protected Sub SendCom57()
        Try
            Dim Param57 As New cCom57 '返答に用いるパラメータオブジェクト

            Param57.tablet = Getタブレット設置状況() 'タブレットの設置状況を得る

            Dim Param57Json As String = JsonConvert.SerializeObject(Param57)

            SendTo指示者Tablet("57", Param57Json)
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "コマンド57の実行時に例外が発生しました。", 2, _設定値.LogPath)
            Return
        End Try
    End Sub

    ''' <summary>盤タブレットへの現状を返すコマンド73やそれに類するコマンドを実行し送信まで行う。</summary>
    ''' <param name="Command">コマンド</param>
    ''' <param name="in_bno">盤の通し番号</param>
    ''' <remarks>                    </remarks>
    ''' <returns>コマンド処理が成功したかどうか</returns>
    Protected Function SendCom7x(Command As String, in_bno As String) As Boolean
        Try

            If in_bno Is Nothing Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "引数の盤ID「in_bno」がnothingです。パラメータを確認して下さい。", 1, _設定値.LogPath)
                Return False
            End If
            If "" = in_bno Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "引数の盤ID「in_bno」が与えられていません(空文字列)。", 1, _設定値.LogPath)
                Return False
            End If
            If CLng(in_bno) < 0 Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "引数の盤ID「in_bno」が0未満です。", 1, _設定値.LogPath)
                Return False
            End If

            '盤が存在するかどうか調べる
            Dim dt盤マスタ1 As DataTable = GetDT_Select("SELECT * FROM ""m_ban"" WHERE ""in_bno""=" & in_bno)
            If dt盤マスタ1.Rows.Count < 1 Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "与えられた盤ID「" & in_bno & "」の盤が見つかりませんでした。", 0, _設定値.LogPath)
                Return True
            End If

            If "" = dt盤マスタ1.Rows(0)("tx_rip").ToString Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "盤（" & in_bno & "）にタブレットが登録されていないので盤への送信は行いません。", 0, _設定値.LogPath)
                Return True
            End If

            ''盤タブレットへコマンド7xを送信
            ''盤情報を取得する
            'Dim TableName As String = "ban" & in_bno
            'Dim dt盤1 As DataTable = GetDT_Select("SELECT * FROM """ & TableName & """")
            'If dt盤1.Rows.Count < 1 Then
            '    ErrMsgBox(Nothing, _設定値.DebugMode, "盤テーブル（" & TableName & "）にレコードがありません。", 1, _設定値.LogPath)
            '    Return False
            'End If
            ''値を割振る
            'Dim 盤s As New c盤表示詳細
            'With 盤s
            '    .lines = dt盤マスタ1.Rows.Count.ToString
            '    .bo_active = dt盤1.Rows(0)("bo_active").ToString
            '    ReDim .m_device(4)
            '    For Index As Long = 0 To 4
            '        Dim tmpM_device As New c機器表示詳細
            '        With tmpM_device
            '            .tx_lb = dt盤1.Rows(0)("tx_lb" & (Index + 1).ToString).ToString
            '            .tx_clr = dt盤1.Rows(0)("tx_clr" & (Index + 1).ToString).ToString
            '            .in_disp_hi = dt盤1.Rows(0)("in_disp_hi" & (Index + 1).ToString).ToString
            '            .in_disp_blink = dt盤1.Rows(0)("in_disp_blink" & (Index + 1).ToString).ToString
            '        End With
            '        .m_device(Index) = tmpM_device
            '    Next
            'End With

            Dim 盤s As c盤表示詳細 = Get盤タブレット表示情報(dt盤マスタ1.Rows(0)("tx_rip"))
            If 盤s Is Nothing Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "盤タブレットへの送信情報の作成に失敗しました。", 1, _設定値.LogPath)
                Return False
            End If

            Send(dt盤マスタ1.Rows(0)("tx_rip"), _設定値.RemotePort, Command, JsonConvert.SerializeObject(盤s)) '送信

            Return True
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "コマンド73類の実行時に例外が発生しました。", 2, _設定値.LogPath)
            Return False
        End Try
    End Function

    'クライアント発呼に対する返答を行う
    Protected Sub ReturnCom(Command As String, Param As String, 受信コマンド As String)
        Try
            '引数を調整する
            Dim CommunicateEventArgs As New cCommunicateEventArgs
            With CommunicateEventArgs
                .ReadCommand = 受信コマンド
                .ReadCommand = _SrcParam
                .WriteCommand = Command
                .WriteParam = Param
            End With
            Dim ClientStart As New cClientStart(_設定値, _SocketState, CommunicateEventArgs)
            ClientStart.Done()  'クライアント発呼への返信を行う
            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "クライアント発呼への返答で例外が発生しました。", 2, _設定値.LogPath)
            Return
        End Try
    End Sub

#End Region

#Region "サーバ上のデータを返すのに必要な関数"

    ''' <summary>手順の総数を返す関数</summary>
    ''' <returns>登録されている手順の総数、実行できなかった時は-1を返す</returns>
    Protected Function Get手順総数() As Long
        Try
            Dim Sql1 As String = "SELECT count(*) FROM ""t_order"""
            Dim dt1 As DataTable = GetDT_Select(Sql1)
            If 0 = dt1.Rows.Count Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "手順の総数の取得でクエリ「" & Sql1 & "」にヒットしませんでした。", 1, _設定値.LogPath)
                Return -1
            End If
            If IsDBNull(dt1.Rows(0)(0)) Then Return -1 'NULL値なら-1を返す
            Return CLng(dt1.Rows(0)(0).ToString)
        Catch ex As Exception
            ErrMsgBox(Nothing, _設定値.DebugMode, "手順の総数の取得で例外が発生しました。", 2, _設定値.LogPath)
            Return -1
        End Try
    End Function

    ''' <summary>全手順書データを配列で返す関数</summary>
    ''' <returns>登録されている全手順書データ。失敗時は空の配列を返す。</returns>
    Protected Function Get全手順データ() As c手順書プレビュー()
        Try
            '手順書を得る
            Dim Dest As c手順書プレビュー() '返答配列
            Dim dt手順 As DataTable = GetDT_Select("SELECT * FROM ""t_order"" order by ""in_sno""")
            If dt手順.Rows.Count < 1 Then       'データが無かった時は
                ErrMsgBox(Nothing, _設定値.DebugMode, "全手順データの取得で手順が一つもありませんでした。", 1, _設定値.LogPath)
                Return Dest
            End If
            'データがあった時は
            'メモリを確保
            Dim tmp手順s As c手順書プレビュー()
            ReDim tmp手順s(dt手順.Rows.Count - 1)

            Dim Index As Long = 0
            For Each Row As DataRow In dt手順.Rows  '全行をループ
                tmp手順s(Index) = New c手順書プレビュー
                With tmp手順s(Index)
                    .in_sno = Row("in_sno").ToString
                    .tx_sno = Row("tx_sno").ToString
                    .tx_basho = Row("tx_basho").ToString
                    .tx_bname = Row("tx_bname").ToString
                    .tx_swname = Row("tx_swname").ToString
                    .tx_action = Row("tx_action").ToString
                    .tx_biko = Row("tx_biko").ToString
                    .dotime = Row("ts_b").ToString
                    .tx_gs = Row("tx_gs").ToString
                    .tx_com = Row("tx_com").ToString
                    .tx_s_l = Row("tx_s_l").ToString
                    .tx_s_r = Row("tx_s_r").ToString
                    .tx_b_l = Row("tx_b_l").ToString
                    .tx_b_r = Row("tx_b_r").ToString
                    .tx_clr1 = Row("tx_clr1").ToString
                    .tx_clr2 = Row("tx_clr2").ToString
                    .ts_b = Row("ts_b").ToString
                    .cd_status = Row("cd_status").ToString
                    .bo_gs = Row("bo_gs").ToString
                    .in_swno = Row("in_swno").ToString
                    .in_bno = Row("in_bno").ToString
                    .cd_pair = Row("cd_pair").ToString
                End With
                Index += 1
            Next
            Dest = tmp手順s
            Return Dest
        Catch ex As Exception
            ErrMsgBox(Nothing, _設定値.DebugMode, "全手順データの取得で例外が発生しました。", 2, _設定値.LogPath)
            Dim Dest As c手順書プレビュー() '返答配列
            Return Dest
        End Try
    End Function

    ''' <summary>タブレットの設置状況を返す</summary>
    ''' <returns>現在のタブレット設置状況データ。失敗時は空の配列を返す</returns>
    Protected Function Getタブレット設置状況() As cタブレット状況()
        Try
            '現状のタブレット状況を取得する
            Dim Dest As cタブレット状況()

            '確認者タブレットを確認する
            Dim 確認タブ状況s As cタブレット状況()
            ReDim 確認タブ状況s(0)
            確認タブ状況s(0) = New cタブレット状況
            With 確認タブ状況s(0) '初期値を指定する
                .名称 = "確認者タブレット"
                .状況 = ""
            End With
            Dim dt確認タブ As DataTable = GetDT_Select("SELECT ""bo_set"" FROM ""m_kakunin""")
            If 0 < dt確認タブ.Rows.Count Then       'データが有った時
                If CBool(dt確認タブ(0)("bo_set")) Then 確認タブ状況s(0).状況 = "○"
            End If

            '盤タブレット状況を得る
            Dim 盤タブ状況s As cタブレット状況()
            Dim dtタブ As DataTable = GetDT_Select("SELECT ""tx_basho"",""tx_bname"",""bo_set"" FROM ""m_ban"" ORDER BY ""tx_basho"",""tx_bname""")
            If 0 = dtタブ.Rows.Count Then       'データが無かった時は
                盤タブ状況s = Nothing
            Else                                'データがあった時は
                'メモリを確保
                ReDim 盤タブ状況s(dtタブ.Rows.Count - 1)
                Dim Index As Long = 0
                For Each Row As DataRow In dtタブ.Rows  '全行をループ
                    盤タブ状況s(Index) = New cタブレット状況
                    With 盤タブ状況s(Index)
                        .名称 = Row("tx_basho").ToString & "　" & Row("tx_bname").ToString
                        .状況 = "" : If CBool(Row("bo_set")) Then .状況 = "○"
                    End With
                    Index += 1
                Next
            End If

            '確認者タブレットの状況と盤タブレットの状況を合体する
            ReDim Dest(盤タブ状況s.Count + 確認タブ状況s.Count - 1)
            Array.Copy(確認タブ状況s, 0, Dest, 0, 確認タブ状況s.Count)
            Array.Copy(盤タブ状況s, 0, Dest, 確認タブ状況s.Count, 盤タブ状況s.Count)

            Return Dest
        Catch ex As Exception
            ErrMsgBox(Nothing, _設定値.DebugMode, "タブレット設置状況の取得で例外が発生しました。", 2, _設定値.LogPath)
            Dim Dest As cタブレット状況()
            Return Dest
        End Try
    End Function

    ''' <summary>現在の手順の情報を返す</summary>
    ''' <returns></returns>
    ''' <remarks>t_snoのデータを返す</remarks>
    Protected Function Get現在手順情報() As c指示現状
        Try
            '手順書を得る
            Dim Dest As c指示現状 '返答配列
            Dim dt1 As DataTable = GetDT_Select("SELECT * FROM ""t_sno""")
            If dt1.Rows.Count < 1 Then       'データが無かった時は
                ErrMsgBox(Nothing, _設定値.DebugMode, "指示の現状データの取得でレコードが一件もありませんでした。", 1, _設定値.LogPath)
                Return Dest
            End If
            'データがあった時は
            Dest = New c指示現状
            With Dest
                .in_tsno = dt1.Rows(0)("in_tsno").ToString
                .in_csno = dt1.Rows(0)("in_csno").ToString
                .in_bno = dt1.Rows(0)("in_bno").ToString
                .cd_status = dt1.Rows(0)("cd_status").ToString
                .cd_pctl = dt1.Rows(0)("cd_pctl").ToString
                .in_snoco = dt1.Rows(0)("in_snoco").ToString
                .cd_runmode = dt1.Rows(0)("cd_runmode").ToString
                .bo_stnojump = dt1.Rows(0)("bo_stnojump").ToString
                .tx_bgcolor = dt1.Rows(0)("tx_bgcolor").ToString
                .cd_finmode = dt1.Rows(0)("cd_finmode").ToString
                .cd_gsmode = dt1.Rows(0)("cd_gsmode").ToString
                .bo_owari = dt1.Rows(0)("bo_owari").ToString
                .tx_sfile = dt1.Rows(0)("tx_sfile").ToString
                .bo_reset = dt1.Rows(0)("bo_reset").ToString
                .bo_poff = dt1.Rows(0)("bo_poff").ToString
            End With
            Return Dest
        Catch ex As Exception
            ErrMsgBox(Nothing, _設定値.DebugMode, "指示の現状データの取得で例外が発生しました。", 2, _設定値.LogPath)
            Dim Dest As c指示現状
            Return Dest
        End Try
    End Function

    ''' <summary>現在手順情報クラス        </summary>
    ''' <remarks>手順の現状(t_snoテーブルの内容)を格納するクラス</remarks>
    Public Class c指示現状

        ''' <summary>？手順書番号？</summary>
        Public Property in_tsno As String

        ''' <summary>手順書(通し)番号</summary>
        ''' <returns>t_orderテーブルのin_snoの値</returns>
        Public Property in_csno As String

        ''' <summary>盤(通し)番号</summary>
        ''' <remarks>m_banテーブルのin_bnoの値</remarks>
        Public Property in_bno As String

        ''' <summary>手順内での指示・確認等のステータス</summary>
        ''' <remarks>0:指示前、1:指示後、確認前、7:当該手順終了(確認済み、スキップ後)</remarks>
        Public Property cd_status As String

        ''' <summary>？？</summary>
        Public Property cd_pctl As String

        ''' <summary>？手順書末尾のin_sno番号？</summary>
        Public Property in_snoco As String

        ''' <summary>？？</summary>
        Public Property cd_runmode As String

        ''' <summary>？？</summary>
        Public Property bo_stnojump As String

        ''' <summary>？テキストの背景色？</summary>
        Public Property tx_bgcolor As String

        ''' <summary>終了状態</summary>
        ''' <remarks>U:手順開始状態、N:手順開始前、1:手順終了状態</remarks>
        Public Property cd_finmode As String

        ''' <summary>この手順の現場差異状態</summary>
        ''' <remarks>0:キャンセル・通常状態、5:スキップ、6:この手順の前に操作を追加、9:追加後の確認済み</remarks>
        Public Property cd_gsmode As String

        ''' <summary>？？</summary>
        Public Property bo_owari As String

        ''' <summary>？エクセルのファイル名？</summary>
        Public Property tx_sfile As String

        ''' <summary>？？</summary>
        Public Property bo_reset As String

        ''' <summary>？全手順終了フラグ？</summary>
        Public Property bo_poff As String

    End Class

    ''' <summary>指定されたリモートIPの盤タブレット表示情報を返す</summary>
    ''' <returns>c盤表示詳細クラスを返す。</returns>
    ''' <remarks>エラー時はnothingを返す</remarks>
    Protected Function Get盤タブレット表示情報(RemoteIP As String) As c盤表示詳細
        Try
            If RemoteIP Is Nothing Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "盤タブレットの表示情報取得をする際に与えられたリモートIPがNothingでした。", 1, _設定値.LogPath)
                Return Nothing
            End If

            If "" = RemoteIP Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "盤タブレットの表示情報取得をする際に与えられたリモートIPが空文字でした。", 1, _設定値.LogPath)
                Return Nothing
            End If

            '盤が存在するかどうか調べる
            Dim dt盤マスタ1 As DataTable = GetDT_Select("SELECT * FROM ""m_ban"" WHERE ""tx_rip""='" & RemoteIP & "';")
            If dt盤マスタ1.Rows.Count < 1 Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "盤タブレットの表示情報取得をする際に与えられたリモートIPの盤が見つかりませんでした。", 1, _設定値.LogPath)
                Return Nothing
            End If

            '盤内情報を取得する
            Dim TableName As String = "ban" & dt盤マスタ1(0)("in_bno").ToString
            Dim dt盤内 As DataTable = GetDT_Select("SELECT * FROM """ & TableName & """;")
            If dt盤内.Rows.Count < 1 Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "盤タブレットの表示情報取得をする際に「" & TableName & "」テーブルの中身がありませんでした。", 1, _設定値.LogPath)
                Return Nothing
            End If

            '機器数を得る
            Dim SQL機器マスタ As String = "SELECT * FROM ""m_device2"" WHERE ""in_bno""=" & dt盤マスタ1(0)("in_bno").ToString & ";"
            Dim dt機器マスタ As DataTable = GetDT_Select(SQL機器マスタ)
            'If dt機器マスタ.Rows.Count < 1 Then '機器の存在しない盤操作も存在すると言う事でこれはエラー扱いしない。
            '    ErrMsgBox(Nothing, _設定値.DebugMode, "盤タブレットの表示情報取得をする際に「" & SQL機器マスタ & "」にヒットするレコードがありませんでした。", 1, _設定値.LogPath)
            '    Return Nothing
            'End If

            '値を割振る
            Dim Dest As New c盤表示詳細
            With Dest
                .in_bno = dt盤マスタ1.Rows(0)("in_bno").ToString
                .tx_bname = dt盤マスタ1.Rows(0)("tx_bname").ToString
                .lines = dt機器マスタ.Rows.Count.ToString
                .bo_active = dt盤内.Rows(0)("bo_active").ToString
                If 0 < .lines Then
                    ReDim .m_device(4)
                    For Index As Long = 0 To 4
                        .m_device(Index) = New c機器表示詳細
                        With .m_device(Index)
                            .tx_lb = dt盤内.Rows(0)("tx_lb" & (Index + 1).ToString).ToString
                            .tx_clr = dt盤内.Rows(0)("tx_clr" & (Index + 1).ToString).ToString
                            .in_disp_hi = dt盤内.Rows(0)("in_disp_hi" & (Index + 1).ToString).ToString
                            .in_disp_blink = dt盤内.Rows(0)("in_disp_blink" & (Index + 1).ToString).ToString
                        End With
                    Next
                End If
            End With
            Return Dest
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "盤タブレットの表示情報取得をする際に例外が発生しました。", 2, _設定値.LogPath)
            Return Nothing
        End Try
    End Function

    ''' <summary>盤タブレットの盤選択一覧情報を返す</summary>
    ''' <returns>c盤表示詳細クラスを返す。</returns>
    ''' <remarks>エラー時はnothingを返す</remarks>
    Protected Function Get盤選択一覧情報() As c盤一覧情報
        Try
            '現在登録されている盤情報を全て返す
            '盤タブレット状況を得る
            Dim 盤s As New c盤一覧情報
            Dim dt盤s As DataTable = GetDT_Select("SELECT * FROM ""m_ban""")
            If 0 < dt盤s.Rows.Count Then       'データがあった時は
                Dim dt場所 As DataTable = GetDT_Select("SELECT DISTINCT ""tx_basho"" FROM ""m_ban"" ORDER BY ""tx_basho""")
                For Each 場所 As DataRow In dt場所.Rows  '全行をループ
                    Dim 場所名 As String = 場所("tx_basho")
                    '場所毎に盤を取得する
                    Dim dt盤in場所 As DataTable = GetDT_Select("SELECT * FROM ""m_ban"" WHERE ""tx_basho""='" & 場所名 & "' ORDER BY ""tx_bname""")
                    If 0 < dt盤in場所.Rows.Count Then
                        Dim Add盤s() As c盤簡易情報
                        ReDim Add盤s(dt盤in場所.Rows.Count - 1)
                        Dim Add盤Index As Long = 0
                        For Each 盤 As DataRow In dt盤in場所.Rows  '全行をループ
                            Dim dt機器マスタ As DataTable = GetDT_Select("SELECT * FROM ""m_device2"" WHERE ""in_bno""=" & 盤("in_bno").ToString & ";")
                            Dim tmpAdd盤s As New c盤簡易情報
                            With tmpAdd盤s
                                .in_bno = 盤("in_bno").ToString
                                .tx_bname = 盤("tx_bname").ToString
                                .tx_rip = 盤("tx_rip").ToString
                                .DeviceCount = dt機器マスタ.Rows.Count.ToString
                            End With
                            Add盤s(Add盤Index) = tmpAdd盤s '配列に追加する
                            Add盤Index += 1
                        Next
                        盤s.盤情報.Add(場所("tx_basho").ToString, Add盤s)
                    End If
                Next
            End If
            Return 盤s
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "盤選択一覧情報を作成する際に例外が発生しました。", 2, _設定値.LogPath)
            Return Nothing
        End Try
    End Function

    ''' <summary>手順が全て終わって電源OFF処理されているかどうかを返す</summary>
    ''' <returns>手順が全て終わって電源OFF処理されているかどうか</returns>
    ''' <remarks>DBのt_snoテーブル、cd_finmodeフィールドが2かどうかを返す。処理失敗時は偽を返す。</remarks>
    Protected Function Is手順完全終了() As Boolean
        Try
            Dim Sql1 As String = "SELECT ""cd_finmode"" FROM ""t_sno"" ;"
            Dim dt1 As DataTable = GetDT_Select(Sql1)
            If 0 = dt1.Rows.Count Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "手順の完全終了確認に失敗しました。t_snoテーブルが空です。クエリ「" & Sql1 & "」にヒットしませんでした。", 1, _設定値.LogPath)
                Return False
            End If
            If "2" = dt1.Rows(0)("cd_finmode") Then Return True
            Return False
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "手順の完全終了確認処理で例外が発生しました。", 2, _設定値.LogPath)
            Return False
        End Try
    End Function

#End Region

#Region "返答時のJSON変換用クラス"

    ''' <summary>手順書指示行クラス                                        </summary>
    ''' <remarks>一般的に手順書の情報を渡すクラス。開始時のプレビューとは別</remarks>
    Public Class c手順書
        Public Property No As String
        Public Property 盤名 As String
        Public Property 機器名 As String
        Public Property 操作 As String
        Public Property 色1 As String
        Public Property 色2 As String
        Public Property 時刻 As String
        Public Property コメント As String
        Public Property 差異種別 As String
    End Class

    ''' <summary>手順書プレビュークラス</summary>
    ''' <remarks>設置状況確認画面(S-01)で用いる</remarks>
    Public Class c手順書プレビュー
        ''' <summary>No.手順番号(本システム上の記述)</summary>
        Public Property in_sno As String
        ''' <summary>No.手順番号(エクセル上の記述)</summary>
        Public Property tx_sno As String
        ''' <summary>場所名</summary>
        Public Property tx_basho As String
        ''' <summary>盤名称</summary>
        Public Property tx_bname As String
        ''' <summary>機器名称</summary>
        Public Property tx_swname As String
        ''' <summary>指示内容</summary>
        Public Property tx_action As String
        ''' <summary>備考</summary>
        Public Property tx_biko As String
        ''' <summary>実行時刻</summary>
        Public Property dotime As String
        ''' <summary>現場差異の種類</summary>
        Public Property tx_gs As String
        ''' <summary>コメント</summary>
        Public Property tx_com As String
        ''' <summary>左ボタン　文字(指示者タブレット)</summary>
        Public Property tx_s_l As String
        ''' <summary>右ボタン　文字(指示者タブレット)</summary>
        Public Property tx_s_r As String
        ''' <summary>左ボタン　文字(確認者タブレット)</summary>
        Public Property tx_b_l As String
        ''' <summary>右ボタン　文字(確認者タブレット)</summary>
        Public Property tx_b_r As String
        ''' <summary>左ボタン　背景色</summary>
        Public Property tx_clr1 As String
        ''' <summary>右ボタン　背景色</summary>
        Public Property tx_clr2 As String
        ''' <summary>実行時刻</summary>
        ''' <remarks>書式「YYYY-MM- DD HH:MM:SS.00」</remarks>
        Public Property ts_b As String
        ''' <summary>手順指示のステータス</summary>
        ''' <remarks>0：指示前、1：指示後、7：確認済</remarks>
        Public Property cd_status As String
        ''' <summary>現場差異フラグ</summary>
        ''' <remarks>'true：現場差異あり、false：現場差異なし</remarks>
        Public Property bo_gs As String
        ''' <summary>★操作紐づけ値？★</summary>
        ''' <remarks>意味不明・要確認</remarks>
        Public Property in_swno As String
        ''' <summary>盤の(通し)番号</summary>
        Public Property in_bno As String
        ''' <summary>この操作が対の操作かどうか</summary>
        ''' <remarks>0:対の操作ではない、1:対の操作、それ以外:コメント</remarks>
        Public Property cd_pair As String
    End Class

    ''' <summary>タブレット状況クラス</summary>
    ''' <remarks>設置状況確認画面(S-01)で用いる</remarks>
    Public Class cタブレット状況
        '各々は特に説明は必要無いかと
        Public Property 名称 As String
        Public Property 状況 As String
    End Class

    ''' <summary>コマンド57クラス                                        </summary>
    ''' <remarks>タブレットの状況と手順書配列(プレビュー)を格納するクラス</remarks>
    Private Class cCom57
        ''' <summary>タブレットの状況を複数格納する配列</summary>
        Public Property tablet As cタブレット状況()
    End Class

    ''' <summary>盤タブレット用表示詳細クラス            </summary>
    ''' <remarks>盤タブレットで盤の詳細の表示に必要な情報を格納するクラス</remarks>
    Public Class c盤表示詳細

        ''' <summary>盤の番号</summary>
        Public Property in_bno As String

        ''' <summary>盤名</summary>
        Public Property tx_bname As String

        ''' <summary>機器数</summary>
        Public Property lines As String

        ''' <summary>盤を開いて操作中のフラグ？</summary>
        Public Property bo_active As String

        ''' <summary>盤内の機器の表示情報</summary>
        Public Property m_device As c機器表示詳細()

    End Class

    ''' <summary>機器情報クラス                      </summary>
    ''' <remarks>一つの機器の表示情報を格納するクラス</remarks>
    Public Class c機器表示詳細

        ''' <summary>機器の名称</summary>
        Public Property tx_lb As String

        ''' <summary>色</summary>
        Public Property tx_clr As String

        ''' <summary>太線枠が点滅状態か</summary>
        Public Property in_disp_hi As String

        ''' <summary>点滅状態</summary>
        Public Property in_disp_blink As String

    End Class

    ''' <summary>盤一覧クラス          </summary>
    ''' <remarks>盤一覧を格納するクラス</remarks>
    Public Class c盤一覧情報

        ''' <summary>盤一覧</summary>
        ''' <remarks>c盤簡易情報クラスを格納する連想配列</remarks>
        Public Property 盤情報 As New Hashtable

    End Class

    ''' <summary>盤簡易情報クラス              </summary>
    ''' <remarks>一つの盤の簡易情報を格納するクラス</remarks>
    Public Class c盤簡易情報

        ''' <summary>盤名</summary>
        Public Property tx_bname As String

        ''' <summary>盤の通し番号</summary>
        Public Property in_bno As String

        ''' <summary>盤タブレットのIPアドレス</summary>
        Public Property tx_rip As String

        ''' <summary>機器の数</summary>
        Public Property DeviceCount As String

    End Class

    ''' <summary>コマンド5Rクラス          </summary>
    ''' <remarks>正しい手順番号を格納するクラス</remarks>
    Public Class cCom5R
        Public Property in_sno As String
    End Class

#End Region

    ''' <summary>タブレットを管理するクラス</summary>
    ''' <remarks></remarks>
    Public Class cTabletList

#Region "内部メンバ"

        ''' <summary>このクラスで管理するタブレットのリスト</summary>
        Dim _List As New List(Of cTablet)

        ''' <summary>送信タスクリストクラスのインスタンス</summary>
        Dim _SendTaskList As cSendTaskList

        ''' <summary>外部から与えられた設定値</summary>
        Dim _設定値 As c設定値

#End Region

#Region "イベント"

        ''' <summary>コンストラクタ</summary>
        ''' <param name="設定値">外部から与えられる設定値</param>
        Public Sub New(ByRef 設定値 As c設定値)
            Try
                _設定値 = 設定値
            Catch ex As Exception
                ErrMsgBox(ex, _設定値.DebugMode, "タブレット管理クラスのコンストラクタで例外が発生しました。", 2, _設定値.LogPath)
                Return
            End Try
        End Sub

#End Region

#Region "パブリックメソッド"

        ''' <summary>タブレットをリストに追加登録する</summary>
        ''' <param name="Tablet">登録するタブレットの情報</param>
        Public Sub AddTablet(ByRef Tablet As cTablet)
            Try
                _List.Add(Tablet)
                Return
            Catch ex As Exception
                ErrMsgBox(ex, _設定値.DebugMode, "タブレット情報追加で例外が発生しました。", 2, _設定値.LogPath)
                Return
            End Try
        End Sub

        ''' <summary>指定されたTypeの相手にコマンドを送信する</summary>
        ''' <param name="Type">送信するタブレットのタイプ、10:指示者、20:確認者</param>
        ''' <param name="Command">送信するコマンド</param>
        ''' <param name="Param">送信するパラメータ</param>
        ''' <remarks>現状では送信の結果は返さない。</remarks>
        Public Sub Send(Type As String, Command As String, Param As String)
            Try
                For Each Tablet As cTablet In _List
                    If Not (Tablet Is Nothing) Then 'インスタンスが存在するなら
                        With Tablet
                            If Type = Tablet.Type Then Send(.IP, Command, Param) '送信する
                        End With
                    End If
                Next
                Return
            Catch ex As Exception
                ErrMsgBox(ex, _設定値.DebugMode, "指定タイプのタブレットへの一斉送信で例外が発生しました。", 2, _設定値.LogPath)
                Return
            End Try
        End Sub

        ''' <summary>コマンドを受信した時にタブレット情報の更新を行う</summary>
        ''' <param name="Command">受信したコマンド</param>
        ''' <param name="SocketState">受信で得たソケット管理オブジェクト</param>
        ''' <remarks>指示者タブレットはこのクラスで管理する。確認者タブレットはDBを更新する。実行に失敗してもエラーは返さない。</remarks>
        Public Sub TabletUpdate(Command As String, ByRef SocketState As cSocketState)
            Try
                '動作条件を調べる
                If SocketState Is Nothing Then
                    ErrMsgBox(Nothing, _設定値.DebugMode, "    ", 1, _設定値.LogPath)
                    Return
                End If
                If SocketState.WorkSocket Is Nothing Then
                    ErrMsgBox(Nothing, _設定値.DebugMode, "タブレット情報更新ができません。SocketインスタンスがNothingです。", 1, _設定値.LogPath)
                    Return
                End If
                If SocketState.WorkSocket.RemoteEndPoint Is Nothing Then
                    ErrMsgBox(Nothing, _設定値.DebugMode, "タブレット情報更新ができません。接続先情報がNothingです。", 1, _設定値.LogPath)
                    Return
                End If
                If "" = Command Then
                    ErrMsgBox(Nothing, _設定値.DebugMode, "タブレット情報更新ができません。コマンドが不明です。", 1, _設定値.LogPath)
                    Return
                End If
                If Command.Length < 2 Then
                    ErrMsgBox(Nothing, _設定値.DebugMode, "タブレット情報更新ができません。コマンドが不正です。", 1, _設定値.LogPath)
                    Return
                End If

                '相手のIPを得る
                Dim RemoteIP As String = CType(SocketState.WorkSocket.RemoteEndPoint, IPEndPoint).Address.ToString

                Select Case Command.Substring(0, 1)
                    Case "1"    '指示者コマンドだった時
                        RemoveOfType("10")  '既に登録されている指示者タブレットを削除する
                        '新しく登録する
                        Dim tmpTablet As New cTablet
                        With tmpTablet
                            .IP = RemoteIP
                            .LastAccess = Now
                            .Name = "実機"
                            .Type = "10"
                        End With
                        _List.Add(tmpTablet)    '追加する
                    Case "2"    '確認者コマンドだった時
                        RemoveOfType("20")  '既に登録されている確認者タブレットを削除する
                        Dim tmp As New cComBase()

                        'DBに登録
                        Dim tmpDB As New cDB(_設定値) ': tmpDB._ConnectText = _設定値.ConnectText ' ConnectText
                        tmpDB.DBNonQuery("UPDATE ""m_kakunin"" SET ""tx_rip""='" & RemoteIP & "', ""bo_set""='t';")
                    Case "3"    '盤タブレットコマンドだった時
                        '既に存在するタブレット情報を消す
                        RemoveOfRemoteIP(RemoteIP)  '既に登録されている同じIPの盤タブレットを削除する
                        '新しく登録する
                        Dim tmpTablet As New cTablet
                        With tmpTablet
                            .IP = RemoteIP
                            .LastAccess = Now
                            .Name = "実機"
                            .Type = "30"
                        End With
                        _List.Add(tmpTablet)    '追加する
                    Case Else
                End Select
                Return
            Catch ex As Exception
                ErrMsgBox(ex, _設定値.DebugMode, "タブレット情報更新で例外が発生しました。", 2, _設定値.LogPath)
                Return
            End Try
        End Sub

        ''' <summary>指定されたTypeのタブレット情報をLISTで返す</summary>
        ''' <param name="Type">タブレットのタイプ、10:指示者、20:確認者</param>
        ''' <returns>該当するタブレットを複数収めたListオブジェクト。</returns>
        Public Function GetTabletList(Type As String) As List(Of cTablet)
            Dim Dest As New List(Of cTablet)
            Try
                'DBに登録されたIPを優先
                'このLIST以外にタブレットがあるなら追加する
                Select Case Type
                    Case "20"   '確認者タブレットの場合
                        Dim tmpDB As New cDB(_設定値) ': tmpDB._ConnectText = _設定値.ConnectText ' ConnectText
                        Dim dt As DataTable = tmpDB.GetDT_Select("SELECT ""tx_rip"" FROM ""m_kakunin"";") ' WHERE ""bo_set""='t'")
                        If Not (0 = dt.Rows.Count) Then '登録済みの確認者タブレットがあったなら
                            Dim tx_rip As String = dt.Rows(0)("tx_rip").ToString
                            If Not ("" = tx_rip) Then   'DBにIPが登録されているなら
                                Dim tmpTablet As New cTablet
                                With tmpTablet
                                    .Type = "20"
                                    .Name = "DB登録機"
                                    .LastAccess = Now
                                    .IP = tx_rip
                                End With
                                Dest.Add(tmpTablet) '返答Listに登録
                            End If
                        End If
                        Exit Select
                    Case "30"   '盤タブレットの場合
                        Dim tmpDB As New cDB(_設定値) ': tmpDB._ConnectText = _設定値.ConnectText ' ConnectText
                        Dim dt As DataTable = tmpDB.GetDT_Select("SELECT ""tx_rip"" FROM ""m_ban"" WHERE ""tx_rip"" != '' ;")
                        If Not (0 = dt.Rows.Count) Then '登録済みの盤タブレットがあったなら
                            For Each RIP As DataRow In dt.Rows
                                Dim tx_rip As String = RIP("tx_rip").ToString
                                If Not ("" = tx_rip) Then   'DBにIPが登録されているなら
                                    Dim tmpTablet As New cTablet
                                    With tmpTablet
                                        .Type = "30"
                                        .Name = "DB登録機"
                                        .LastAccess = Now
                                        .IP = tx_rip
                                    End With
                                    Dest.Add(tmpTablet)
                                End If
                            Next
                        End If
                        Exit Select
                End Select

                '現在このクラスで保持しているタブレットデータを最後に追加
                For Each srcTablet As cTablet In _List
                    If Not (srcTablet Is Nothing) Then 'リストアイテムのインスタンスが有る時（エラー対策）
                        If srcTablet.Type = Type Then  '同じタイプなら
                            '既に同じIPが登録されていないか確認
                            Dim AddFlag As Boolean = True
                            For Each destTablet As cTablet In Dest              '登録先Listをループ
                                If srcTablet.IP.Trim = destTablet.IP.Trim Then  '一致するIPが見つかったら
                                    AddFlag = False                             '追加フラグを下げて
                                    Exit For                                    'ループから抜ける
                                End If
                            Next
                            If AddFlag Then Dest.Add(srcTablet) '返答Listに登録する
                        End If
                    End If
                Next

                '指示者・確認者タブレット情報が失われていた時の回復処理
                Select Case Type
                    Case "10", "20"
                        If Dest.Count < 1 Then
                            '新しく登録する
                            Dim tmpTablet As New cTablet
                            With tmpTablet
                                Select Case Type
                                    Case "10" : .IP = "192.168.10.30"   '指示者
                                    Case "20" : .IP = "192.168.10.40"   '確認者
                                End Select
                                .LastAccess = Now
                                .Name = "実機"
                                .Type = Type
                            End With
                            Dest.Add(tmpTablet)    '追加する
                        End If
                End Select
                
                Return Dest
            Catch ex As Exception
                ErrMsgBox(ex, _設定値.DebugMode, "タブレット検索で例外が発生しました。", 2, _設定値.LogPath)
                Return Dest
            End Try
        End Function

        ''' <summary>指定されたリモートIPが盤に登録されているかどうかを返す</summary>
        ''' <param name="RemoteIP">調べたいリモートIP</param>
        ''' <remarks>エラー時にはFALSEを返す</remarks>
        Public Function Is盤登録済みIP(RemoteIP As String) As Boolean
            Try
                '実行条件を確認
                If RemoteIP Is Nothing Then Return False '登録対象のIPではないので偽を返す
                If "" = RemoteIP Then Return False '登録対象のIPではないので偽を返す

                '現在盤に登録されているタブレットのIPを取得する
                Dim tmpDB As New cDB(_設定値)
                Dim dt盤マスタ1 As DataTable = tmpDB.GetDT_Select("SELECT DISTINCT ""tx_rip"" FROM ""m_ban"" WHERE ""tx_rip"" != '' ;")
                If dt盤マスタ1.Rows.Count < 1 Then Return False '登録IPが無いので偽を返す

                For Each 盤IP As DataRow In dt盤マスタ1.Rows  '全行をループ
                    If RemoteIP = 盤IP("tx_rip") Then Return True '一致するIPが有ったら真を返す
                Next

                '見つからなかったときは偽を返す
                Return False
            Catch ex As Exception
                ErrMsgBox(ex, _設定値.DebugMode, "指定されたリモートIPが盤に登録されているかを調べる処理で例外が発生しました。", 2, _設定値.LogPath)
                Return False
            End Try
        End Function

#End Region

#Region "内部関数"

        ''' <summary>指定されたTypeのタブレット情報を全て削除する</summary>
        ''' <param name="Type">削除したいタブレットのType</param>
        Private Sub RemoveOfType(Type As String)
            Try
                '実行条件を確認
                If _List Is Nothing Then Return
                If _List.Count < 1 Then Return

                For Index As Long = _List.Count - 1 To 0 Step -1
                    If _List(Index) Is Nothing Then 'リストアイテムのインスタンスが無い時
                        'エラーの原因になるので削除する
                        _List.Remove(_List(Index))
                    Else    'リストアイテムのインスタンスがあるとき
                        If _List(Index).Type = Type Then    'タイプが一致するときは
                            _List.Remove(_List(Index))      '削除
                        End If
                    End If
                Next
                Return
            Catch ex As Exception
                ErrMsgBox(ex, _設定値.DebugMode, "タイプ指定によるタブレット情報削除で例外が発生しました。", 2, _設定値.LogPath)
                Return
            End Try
        End Sub

        ''' <summary>指定されたリモートIPのタブレット情報を全て削除する</summary>
        ''' <param name="RemoteIP">削除したいリモートIP</param>
        Private Sub RemoveOfRemoteIP(RemoteIP As String)
            Try
                '実行条件を確認
                If _List Is Nothing Then Return
                If _List.Count < 1 Then Return
                If RemoteIP Is Nothing Then Return
                If "" = RemoteIP Then Return

                For Index As Long = _List.Count - 1 To 0 Step -1
                    If _List(Index) Is Nothing Then 'リストアイテムのインスタンスが無い時
                        'エラーの原因になるので削除する
                        _List.Remove(_List(Index))
                    Else    'リストアイテムのインスタンスがあるとき
                        If _List(Index).IP = RemoteIP Then  'IPが一致するときは
                            _List.Remove(_List(Index))      '削除
                        End If
                    End If
                Next
                Return
            Catch ex As Exception
                ErrMsgBox(ex, _設定値.DebugMode, "リモートIP指定によるタブレット情報削除で例外が発生しました。", 2, _設定値.LogPath)
                Return
            End Try
        End Sub

#End Region

        ''' <summary>タブレットの情報を収めたクラス</summary>
        Public Class cTablet

            ''' <summary>IPアドレス</summary>
            Public IP As String

            ''' <summary>タブレットのタイプ</summary>
            Public Type As String

            ''' <summary>名前</summary>
            ''' <remarks>わかりやすい名前でOK</remarks>
            Public Name As String

            ''' <summary>最終アクセス日時</summary>
            ''' <remarks>現在未使用・未対応</remarks>
            Public LastAccess As DateTime

        End Class

    End Class

    ''' <summary>各タブレットへ送信した直近のコマンドを管理するクラス</summary>
    Public Class cLastCommandList

#Region "内部メンバ"

        ''' <summary>このクラスで管理する直近コマンドのリスト</summary>
        Dim _List As New List(Of cLastCommand)

        ''' <summary>外部から与えられた設定値</summary>
        Dim _設定値 As c設定値

#End Region

#Region "イベント"

        ''' <summary>コンストラクタ</summary>
        ''' <param name="設定値">外部から与えられる設定値</param>
        Public Sub New(ByRef 設定値 As c設定値)
            Try
                _設定値 = 設定値
            Catch ex As Exception
                ErrMsgBox(ex, _設定値.DebugMode, "直近コマンド管理クラスのコンストラクタで例外が発生しました。", 2, _設定値.LogPath)
                Return
            End Try
        End Sub

#End Region

#Region "パブリックメソッド"

        ''' <summary>コマンドをリストに追加登録する</summary>
        ''' <param name="Tablet">登録するタブレットの情報</param>
        Public Sub AddLastCommand(LastCommand As cLastCommand)
            Try
                RemoveOfIP(LastCommand.IP.Trim)   '指定されたIPのコマンド情報を削除する
                _List.Add(LastCommand)              '追加する
                Return
            Catch ex As Exception
                ErrMsgBox(ex, _設定値.DebugMode, "直近コマンド追加で例外が発生しました。", 2, _設定値.LogPath)
                Return
            End Try
        End Sub

        ''' <summary>指定されたIPの直近コマンド情報を返す</summary>
        ''' <param name="IP">目的のIPアドレス</param>
        ''' <returns>該当する直近コマンド情報。無い時はNothingを返す</returns>
        Public Function GetLastCommand(IP As String) As cLastCommand
            Dim Dest As New cLastCommand
            Try
                '実行条件を確認
                If _List Is Nothing Then Return Nothing
                If _List.Count < 1 Then Return Nothing

                For Index As Long = _List.Count - 1 To 0 Step -1
                    If _List(Index) Is Nothing Then 'リストアイテムのインスタンスが無い時
                        'エラーの原因になるので削除する
                        _List.Remove(_List(Index))
                    Else    'リストアイテムのインスタンスがあるとき
                        If _List(Index).IP.Trim = IP.Trim Then    'IPが一致するときは
                            Return _List(Index)                     '返す
                        End If
                    End If
                Next
                Return Nothing
            Catch ex As Exception
                ErrMsgBox(ex, _設定値.DebugMode, "直近コマンド検索で例外が発生しました。", 2, _設定値.LogPath)
                Return Dest
            End Try
        End Function

        ''' <summary>指示者・確認者の直近コマンドをクリアする</summary>
        Public Sub Remove指示者確認者Command()
            Try
                '指示者のIPを得る
                Dim Target As New List(Of cTabletList.cTablet)
                Dim tmpTarget As List(Of cTabletList.cTablet)
                tmpTarget = _設定値.TabletList.GetTabletList("10") : If Not (tmpTarget Is Nothing) Then If 0 < tmpTarget.Count Then Target.AddRange(tmpTarget) '指示者
                tmpTarget = _設定値.TabletList.GetTabletList("20") : If Not (tmpTarget Is Nothing) Then If 0 < tmpTarget.Count Then Target.AddRange(tmpTarget) '確認者

                '削除対象のListをループして削除
                For TargetIndex As Long = Target.Count - 1 To 0 Step -1
                    If Target(TargetIndex) Is Nothing Then '削除対象リストのアイテムのインスタンスが無い時
                        'エラーの原因になるので削除する
                        Target.Remove(Target(TargetIndex))
                    Else    '削除対象リストのアイテムのインスタンスがあるとき
                        RemoveOfIP(Target(TargetIndex).IP.Trim) 'そのIPの直近コマンドをクリアする
                    End If
                Next
                Return
            Catch ex As Exception
                ErrMsgBox(ex, _設定値.DebugMode, "指示者・確認者の直近コマンド削除で例外が発生しました。", 2, _設定値.LogPath)
                Return
            End Try
        End Sub

        ''' <summary>直近コマンドを全てクリアする</summary>
        Public Sub AllRemove()
            Try
                _List = New List(Of cLastCommand)
                Return
            Catch ex As Exception
                ErrMsgBox(ex, _設定値.DebugMode, "直近コマンド全削除で例外が発生しました。", 2, _設定値.LogPath)
                Return
            End Try
        End Sub

#End Region

#Region "内部関数"

        ''' <summary>指定されたIPのコマンド情報を全て削除する</summary>
        ''' <param name="IP">削除したいIP</param>
        Private Sub RemoveOfIP(IP As String)
            Try
                '実行条件を確認
                If _List Is Nothing Then Return
                If _List.Count < 1 Then Return

                For Index As Long = _List.Count - 1 To 0 Step -1
                    If _List(Index) Is Nothing Then 'リストアイテムのインスタンスが無い時
                        'エラーの原因になるので削除する
                        _List.Remove(_List(Index))
                    Else    'リストアイテムのインスタンスがあるとき
                        If _List(Index).IP.Trim = IP.Trim Then    'IPが一致するときは
                            _List.Remove(_List(Index))      '削除
                        End If
                    End If
                Next
                Return
            Catch ex As Exception
                ErrMsgBox(ex, _設定値.DebugMode, "IP指定による直近コマンド情報削除で例外が発生しました。", 2, _設定値.LogPath)
                Return
            End Try
        End Sub

#End Region

        ''' <summary>直近のコマンド情報を収めたクラス</summary>
        Public Class cLastCommand

            ''' <summary>IPアドレス</summary>
            Public IP As String

            ''' <summary>コマンド</summary>
            Public Command As String

            ''' <summary>コマンドパラメータ</summary>
            Public Param As String

        End Class

    End Class

End Class

'実装
#Region "指示者タブレット⇒サーバ"

''' <summary>コマンド10受信処理クラス</summary>
''' <remarks>
'''     各タブレットの設置状況要求コマンド。
'''     パラメータなし。
'''     コマンド51を返す。
'''     ResParam関数で返答データを返す。
'''     旧php版confmenu.php参考。
''' </remarks>
Public Class cCom10
    Inherits cComBase   '基底の継承を宣言

    ''' <summary>コマンドを実行する。</summary>
    ''' <remarks>                    </remarks>
    Public Overrides Sub DoCom()
        Try
            '指示者タブレット情報を更新する
            _設定値.TabletList.TabletUpdate("10", _SocketState)

            '現状のタブレット状況を確認する
            Dim Param As New cCom51 '返答に用いるパラメータオブジェクト

            Param.tablet = Getタブレット設置状況()   'タブレットの状況を得る

            Param.tejun = Get全手順データ()   '手順書を得る

            '手順の現状を得る
            Param.t_sno = Get現在手順情報()

            ReturnCom("51", JsonConvert.SerializeObject(Param), "10")   '返信 

            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "コマンド10処理で例外が発生しました。", 2, _設定値.LogPath)
            ReturnCom(_ErrCom, "", "10")   '返信 
        End Try
    End Sub

#Region "返答用JSONクラス"

    ''' <summary>コマンド51クラス                                        </summary>
    ''' <remarks>タブレットの状況と手順書配列(プレビュー)を格納するクラス</remarks>
    Private Class cCom51
        ''' <summary>タブレットの状況を複数格納する配列</summary>
        Public Property tablet As cタブレット状況()
        ''' <summary>プレビュー用の手順書の手順を複数格納する配列</summary>
        Public Property tejun As c手順書プレビュー()
        ''' <summary>手順の現状を格納した型</summary>
        Public Property t_sno As c指示現状
    End Class

#End Region

End Class

''' <summary>コマンド12受信処理クラス</summary>
''' <remarks>
'''     手順書要求コマンド。
'''     パラメータなし。
'''     コマンド53を返す。
'''     ResParam関数で返答データを返す。
'''     旧php版s1.php参考。
''' </remarks>
Public Class cCom12
    Inherits cComBase   '基底の継承を宣言

    ''' <summary>コマンドを実行する。</summary>
    ''' <remarks>                    </remarks>
    Public Overrides Sub DoCom()
        Try
            '指示者タブレット情報を更新する
            _設定値.TabletList.TabletUpdate("12", _SocketState)

            '現状のタブレット状況を確認する
            Dim Param As New cCom53 '返答に用いるパラメータオブジェクト
            '検索する

            '手順書を得る
            Dim dt手順 As DataTable = GetDT_Select("SELECT * FROM ""t_order"" WHERE ""bo_tf""='f' order by ""in_sno""")
            If 0 = dt手順.Rows.Count Then       'データが無かった時は
                Param.tejun = Nothing
            Else                                'データがあった時は
                'メモリを確保
                Dim tmp手順s As c手順書()
                ReDim tmp手順s(dt手順.Rows.Count - 1)
                Dim Index As Long = 0
                For Each Row As DataRow In dt手順.Rows  '全行をループ
                    tmp手順s(Index) = New c手順書
                    With tmp手順s(Index)
                        .No = Row("tx_sno").ToString
                        .盤名 = Row("tx_bname").ToString
                        .機器名 = Row("tx_swname").ToString
                        .操作 = Row("tx_s_r").ToString
                        .色1 = Row("tx_clr1").ToString
                        .色2 = Row("tx_clr2").ToString
                        .時刻 = Row("ts_b").ToString '("yyyy-MM-dd hh:mm:ss.00")
                        .コメント = Row("tx_com").ToString
                        .差異種別 = Row("tx_gs").ToString
                    End With
                    Index += 1
                Next
                Param.tejun = tmp手順s
            End If

            ReturnCom("53", JsonConvert.SerializeObject(Param), "12")   '返信 

            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "コマンド12処理で例外が発生しました。", 2, _設定値.LogPath)
            ReturnCom(_ErrCom, "", "12")   '返信 
        End Try
    End Sub

#Region "返答用JSONクラス"

    ''' <summary>コマンド53クラス          </summary>
    ''' <remarks>手順書配列を格納するクラス</remarks>
    Public Class cCom53
        ''' <summary>手順書の手順を複数格納する配列</summary>
        Public Property tejun As c手順書()
    End Class

#End Region

End Class

''' <summary>コマンド13受信処理クラス</summary>
''' <remarks>
'''     手順指示発令コマンド。
'''     パラメータあり。
'''     コマンド55を返す。
'''     ResParam関数で返答データを返す。
'''     旧php版set_order.php参考。
''' </remarks>
Public Class cCom13
    Inherits cComBase   '基底の継承を宣言

    ''' <summary>手順書番号</summary>
    Public Property 手順書番号 As String

    ''' <summary>コマンドを実行する。</summary>
    ''' <remarks>                    </remarks>
    Public Overrides Sub DoCom()
        Dim トランザクション中 As Boolean = False
        Try
            '指示者タブレット情報を更新する
            _設定値.TabletList.TabletUpdate("13", _SocketState)


            '受信した手順書番号が数値に変換できるか確認
            Try
                Dim tmpLong As Long = CLng(手順書番号.Trim)
            Catch ex As Exception
                ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド13の処理に失敗しました。受信パラメータが数値に変換できません。「" & 手順書番号 & "」", 1, _設定値.LogPath)
                ReturnCom(_ErrCom, "", "13")     'エラーコマンドを返信
                Return
            End Try

            '手順書番号 = _SrcParam 'JSON変換しない時
            DBNonQuery("BEGIN") : トランザクション中 = True

            '現在の手順の指示情報を得る
            Dim Sql1 As String = "SELECT * FROM ""t_sno"""
            Dim dt1 As DataTable = GetDT_Select(Sql1)
            If 0 = dt1.Rows.Count Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド13の処理に失敗しました。クエリ「" & Sql1 & "」にヒットしませんでした。", 1, _設定値.LogPath)
                ReturnCom(_ErrCom, "", "13")     'エラーを返信
                Return
            End If

            '手順書番号のチェックを行う
            Dim in_csno As String = dt1.Rows(0)("in_csno").ToString.Trim
            If Not (in_csno.Trim = 手順書番号.Trim) Then    '手順書番号が正しくなければ
                'If CLng(手順書番号.Trim) < CLng(in_csno.Trim) Then    '受信した手順書番号がDB側の手順番号より小さい時は
                '    ReturnCom("54", "", "13")   '返信 
                '    ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド13で受信した手順書の番号がDBがもつ手順書の番号より小さかったので、コマンド54でタブレット側の手順を進めました。", 0, _設定値.LogPath)

                '    SendTo指示者Tablet("55", ResParam) '指示者へ送信

                '    Return
                'Else
                ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド13で受信した手順書の番号がDBと一致しません。", 0, _設定値.LogPath)

                '正しい手順書番号を返す
                Dim RetParam As New cCom5R
                RetParam.in_sno = in_csno.Trim

                ReturnCom("5R", JsonConvert.SerializeObject(RetParam), "13")    '指示者タブレットへ返信
                Return                                  '抜ける
                'End If
            End If

            '現在のステータスが0以外なら
            If Not ("0" = dt1.Rows(0)("cd_status").ToString) Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "手順管理のステータス、t_sno.cd_statusの値が0ではありません。", 0, _設定値.LogPath)
                ReturnCom(_ThroughCom, "", "13")   '返信 
                Return
            End If

            '手順書の指示を更新する
            Dim Sql2 As String = "SELECT * FROM ""t_order"" WHERE ""in_sno""=" & dt1.Rows(0)("in_csno").ToString
            Dim dt2 As DataTable = GetDT_Select(Sql2)
            If 0 = dt2.Rows.Count Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド13の処理に失敗しました。クエリ「" & Sql2 & "」にヒットしませんでした。", 1, _設定値.LogPath)
                ReturnCom(_ErrCom, "", "13")     'エラーを返信
                Return
            End If

            DBNonQuery("UPDATE ""t_sno"" SET ""in_tsno""=" & dt1.Rows(0)("in_csno").ToString & ",""in_csno""=" & dt1.Rows(0)("in_csno").ToString & ",""in_bno""=" & dt2.Rows(0)("in_bno").ToString & ",""cd_status""='1',""cd_pctl""='1',""cd_finmode""='U'")
            DBNonQuery("UPDATE ""t_order"" SET ""cd_status""='1' WHERE ""in_sno""=" & dt1.Rows(0)("in_csno").ToString)

            '★php版ではここに調査用LOG出力がある

            '該当する盤情報を取得する
            Dim in_bno As String = dt2.Rows(0)("in_bno").ToString
            Dim tBName As String = "ban" & in_bno
            Dim Sql3 As String = "SELECT * FROM """ & tBName & """"
            Dim dt3 As DataTable = GetDT_Select(Sql3)
            If 0 = dt3.Rows.Count Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド13の処理に失敗しました。クエリ「" & Sql3 & "」にヒットしませんでした。", 1, _設定値.LogPath)
                ReturnCom(_ErrCom, "", "13")     'エラーを返信
                Return
            End If

            'ハイライトする機器のフラグを立てる
            '機器は複数登録されている事があるのでそれをここで比較確認してフラグを操作する
            Dim Sql4 As String = "UPDATE """ & tBName & """ SET bo_active='t' "

            '指示書の機器名を半角にして比較する
            Dim s_swname As String = StrConv(dt2.Rows(0)("tx_swname").ToString(), VbStrConv.Narrow)
            Dim tx_lb(5) As String
            tx_lb(1) = StrConv(dt3.Rows(0)("tx_lb1").ToString(), VbStrConv.Narrow)
            tx_lb(2) = StrConv(dt3.Rows(0)("tx_lb2").ToString(), VbStrConv.Narrow)
            tx_lb(3) = StrConv(dt3.Rows(0)("tx_lb3").ToString(), VbStrConv.Narrow)
            tx_lb(4) = StrConv(dt3.Rows(0)("tx_lb4").ToString(), VbStrConv.Narrow)
            tx_lb(5) = StrConv(dt3.Rows(0)("tx_lb5").ToString(), VbStrConv.Narrow)

            '★旧PHP版では「半角英数、全角カナ」にして比較とあったが
            '★これを同時に行う関数がVBには無い為暫定的に全て半角にして比較する
            '★もし忠実に移植するなら一文字ずつ比較変換するか、正規表現で変換する
            Dim s機器名s As String() = s_swname.Split(",") '分割する
            For Each s機器名 As String In s機器名s
                Dim s代入値 As String = StrConv(s機器名 & vbCrLf & dt2.Rows(0)("tx_action2").ToString(), VbStrConv.Narrow)
                Select Case s代入値 's機器名
                    Case tx_lb(1) : Sql4 &= ",""in_disp_hi1""=1"
                    Case tx_lb(2) : Sql4 &= ",""in_disp_hi2""=1"
                    Case tx_lb(3) : Sql4 &= ",""in_disp_hi3""=1"
                    Case tx_lb(4) : Sql4 &= ",""in_disp_hi4""=1"
                    Case tx_lb(5) : Sql4 &= ",""in_disp_hi5""=1"
                End Select
            Next
            DBNonQuery(Sql4)

            '★php版ではここに調査用LOG出力がある

            'ReturnCom(_ThroughCom, "", "13")   '返信 '55@は確認者タブレットで確認された時指示者タブレットが受け取るコマンド
            _設定値.LastCommandList.Remove指示者確認者Command()  '直近のコマンドをクリア
            ReturnCom("54", "", "13")   '返信 '55@は確認者タブレットで確認された時指示者タブレットが受け取るコマンド
            SendTo確認者Tablet("63", "")       '確認者タブレットへ送信
            SendCom7x("73", in_bno)            '盤タブレットへ送信

            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "コマンド13の実行時に例外が発生しました。", 2, _設定値.LogPath)
            ReturnCom(_ErrCom, "", "13")     'エラーを返信
        Finally
            If トランザクション中 Then DBNonQuery("COMMIT")
        End Try
    End Sub

End Class

''' <summary>コマンド14受信処理クラス</summary>
''' <remarks>
'''     差異画面コマンド。
'''     パラメータあり。
'''     コマンド56を返す。
'''     ResParam関数で返答データを返す。
'''     旧php版set_gs.php参考。
''' </remarks>
Public Class cCom14
    Inherits cComBase   '基底の継承を宣言

    ''' <summary>手順書番号</summary>
    Public Property in_sno As String
    ''' <summary>コマンド(0:キャンセル、1:この手順をスキップ、2:この手順の前に操作を追加)</summary>
    Public Property Com As String

    ''' <summary>コマンドを実行する。</summary>
    ''' <remarks>                    </remarks>
    Public Overrides Sub DoCom()
        Dim トランザクション中 As Boolean = False
        Dim Transaction As Npgsql.NpgsqlTransaction

        Try
            トランザクション中 = True
            DB_BeginTransaction(Transaction)
            'DBNonQuery("BEGIN") : トランザクション中 = True
            If Transaction Is Nothing Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド14でトランザクションを開始できませんでした。(Nothing)", 2, _設定値.LogPath)
                ReturnCom(_ErrCom, "", "14")     'エラーを返信
                Return
            End If
            'If Transaction.Connection Is Nothing Then
            '    ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド14でトランザクションを開始できませんでした。(Connectionレス)", 2, _設定値.LogPath)
            '    ReturnCom(_ErrCom, "", "14")     'エラーを返信
            '    Return
            'End If


            '指示者タブレット情報を更新する
            _設定値.TabletList.TabletUpdate("14", _SocketState)


            '手順書番号を確認する
            Dim Sql1 As String = "SELECT * FROM ""t_sno"""
            Dim dt1 As DataTable = GetDT_Select(Sql1)
            If 0 = dt1.Rows.Count Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド14の処理に失敗しました。クエリ「" & Sql1 & "」にヒットしませんでした。", 2, _設定値.LogPath)
                ReturnCom(_ErrCom, "", "14")     'エラーを返信
                Return
            End If
            Dim in_csno As String = dt1.Rows(0)("in_csno").ToString
            If Not (in_csno.Trim = in_sno.Trim) Then    '手順書番号が正しくなければ
                ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド14で受信した手順書の番号がDBと一致しません。", 0, _設定値.LogPath)

                '正しい手順書番号を返す
                Dim RetParam As New cCom5R
                RetParam.in_sno = in_csno.Trim

                ReturnCom("5R", JsonConvert.SerializeObject(RetParam), "14")    '支持者タブレットへ返信

                Return                                  '抜ける
            End If

            '現場差異のフラグを変更する
            _SendCom = "50" 'コマンド
            Select Case Com
                Case 0  'キャンセル
                    DBNonQuery("UPDATE ""t_sno"" SET ""cd_gsmode""='0'")
                Case 1  'スキップ
                    DBNonQuery("UPDATE ""t_sno"" SET ""cd_gsmode""='5'")
                Case 2  'この手順の前に操作を追加
                    DBNonQuery("UPDATE ""t_sno"" SET ""cd_gsmode""='6'")
                Case Else   'それ以外
                    _SendCom = _ErrCom  'コマンド'とりあえずエラーを返す
            End Select

            _設定値.LastCommandList.Remove指示者確認者Command()  '直近のコマンドをクリア
            ReturnCom(_SendCom, "", "14")   '返信

            If _SendCom = _ErrCom Then Return 'エラー時は抜ける

            'エラーじゃない時は確認者へ送信
            Dim Com64 As New cCom64
            Com64.Com = Com
            SendTo確認者Tablet("64", JsonConvert.SerializeObject(Com64))       '確認者タブレットへ送信

            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "コマンド14の実行時に例外が発生しました。", 2, _設定値.LogPath)
            ReturnCom(_ErrCom, "", "14")     'エラーを返信
        Finally
            Try
                If トランザクション中 Then
                    'DBNonQuery("COMMIT")
                    If Not (Transaction Is Nothing) Then
                        If Not (Transaction.Connection Is Nothing) Then Transaction.Commit()
                        Transaction.Dispose()
                        Transaction = Nothing
                    End If
                End If
            Catch ex As Npgsql.NpgsqlException
            End Try
        End Try
    End Sub

#Region "返答用JSONクラス"

    ''' <summary>コマンド64クラス          </summary>
    ''' <remarks>現場差異の受信を伝える情報を格納するクラス</remarks>
    Public Class cCom64
        ''' <summary>現場差異の種類</summary>
        Public Property Com As String
    End Class

#End Region

End Class

' ''' <summary>コマンド15受信処理クラス</summary>
' ''' <remarks>
' '''     終了方法選択コマンド。「指示者確認者のタブレットを終了」
' '''     パラメータなし。
' '''     返答の必要無し。
' '''     旧php版finish.php参考。
' '''     廃止との事。
' ''' </remarks>
'Public Class cCom15
'    Inherits cComBase   '基底の継承を宣言

'    ''' <summary>コマンドを実行する。</summary>
'    ''' <remarks>                    </remarks>
'    Public Overrides Sub DoCom()
'        Try
'            DBNonQuery("UPDATE ""t_sno"" SET ""cd_finmode""=1")

'            '★ここでDBバックアップ処理が入る

'            _SendCom = _ThroughCom 'コマンド
'            _SendParam = "" 'パラメータ
'            Return
'        Catch ex As Exception
'            _SendCom = _ErrCom
'            _SendParam = ""
'            Err.Raise(9999, Nothing, "コマンド15処理に失敗しました。" & vbCrLf & ex.Message)
'        End Try
'    End Sub

'End Class

''' <summary>コマンド16受信処理クラス</summary>
''' <remarks>
'''     終了方法選択コマンド。「全てのタブレットを終了」
'''     パラメータなし。
'''     電源OFF画面コマンド9Cを返す。
'''     旧php版finish.php参考。
''' </remarks>
Public Class cCom16
    Inherits cComBase   '基底の継承を宣言

    ''' <summary>コマンドを実行する。</summary>
    ''' <remarks>                    </remarks>
    Public Overrides Sub DoCom()
        Dim 確認者Tablets As List(Of cTabletList.cTablet)
        Dim 盤Tablets As List(Of cTabletList.cTablet)
        Dim 登録済盤Tablets As List(Of cTabletList.cTablet)
        Dim 未登録盤Tablets As List(Of cTabletList.cTablet)
        Try
            '指示者タブレット情報を更新する
            _設定値.TabletList.TabletUpdate("16", _SocketState)

            'まずリセット情報を送る送信先情報を確保する
            'DBをリセットすると盤の登録情報が消える為
            確認者Tablets = _設定値.TabletList.GetTabletList("20")  '確認者タブレットの一覧を得る
            盤Tablets = _設定値.TabletList.GetTabletList("30")      '盤タブレットの一覧を得る
            登録済盤Tablets = New List(Of cTabletList.cTablet)
            未登録盤Tablets = New List(Of cTabletList.cTablet)
            For Each 盤Tablet As cTabletList.cTablet In 盤Tablets
                If _設定値.TabletList.Is盤登録済みIP(盤Tablet.IP) Then
                    登録済盤Tablets.Add(盤Tablet)
                Else
                    未登録盤Tablets.Add(盤Tablet)
                End If
            Next

            '手順が既に終わっているかどうかを確認する
            '現在の手順位置を得る
            Dim Sql1 As String = "SELECT * FROM ""t_sno"""
            Dim dt1 As DataTable = GetDT_Select(Sql1)
            If 0 = dt1.Rows.Count Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド16の処理に失敗しました。t_snoテーブルが空です。クエリ「" & Sql1 & "」にヒットしませんでした。", 1, _設定値.LogPath)
                ReturnCom(_ErrCom, "", "30")     'エラーコマンドを返信
                Return
            End If

            '残りの手順数を得る
            Dim in_csno As String = dt1(0)("in_csno").ToString
            Dim Sql2 As String = "SELECT COUNT(*) FROM ""t_order"" WHERE ""in_sno"" >= " & in_csno & ";"
            Dim dt2 As DataTable = GetDT_Select(Sql2)
            If 0 = dt2.Rows.Count Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド16の処理に失敗しました。検索結果が不正です。クエリ「" & Sql2 & "」にヒットしませんでした。", 1, _設定値.LogPath)
                ReturnCom(_ErrCom, "", "30")     'エラーコマンドを返信
                Return
            End If
            If dt2.Rows(0)(0) < 1 Then  '手順が終わっているなら
                'DBのリセット処理を行う
                DBNonQuery("UPDATE ""t_sno"" SET ""cd_finmode""='2' ;") '手順が終わった上での電源OFF済みとする
            End If

            '確認者タブレットは常にリセットとする
            DBNonQuery("UPDATE ""m_kakunin"" SET ""bo_poff""='t',""bo_set""='f'")
            '盤タブレットはIPはそのままフラグだけリセット
            DBNonQuery("UPDATE ""m_ban"" SET ""bo_poff""='t',""bo_set""='f'")

            ReturnCom(_ThroughCom, "", "16")   '返信 

            '各タブレットへ電源段を送信
            Dim DestCom As String = "9C"
            '確認者へ送る
            For Each 確認者tablet As cTabletList.cTablet In 確認者Tablets
                Send(確認者tablet.IP, _設定値.RemotePort, DestCom, "")
            Next
            '登録済盤タブレットへ送る
            For Each 盤tablet As cTabletList.cTablet In 登録済盤Tablets
                Send(盤tablet.IP, _設定値.RemotePort, DestCom, "")
            Next
            '未登録盤タブレットへ送る
            For Each 盤tablet As cTabletList.cTablet In 未登録盤Tablets
                Send(盤tablet.IP, _設定値.RemotePort, DestCom, "")
            Next

            _設定値.LastCommandList.AllRemove()    '直近のコマンドを全て削除(タブレット側の再起動時の変な挙動への対処)

            'DoneDump()  'DBのバックアップを保存

            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "コマンド16の実行時に例外が発生しました。", 2, _設定値.LogPath)
            ReturnCom(_ErrCom, "", "16")     'エラーを返信
        Finally
            If Not (確認者Tablets Is Nothing) Then 確認者Tablets = Nothing
            If Not (盤Tablets Is Nothing) Then 盤Tablets = Nothing
            If Not (登録済盤Tablets Is Nothing) Then 登録済盤Tablets = Nothing
            If Not (未登録盤Tablets Is Nothing) Then 未登録盤Tablets = Nothing
        End Try
    End Sub

End Class

''' <summary>コマンド17受信処理クラス</summary>
''' <remarks>
'''     開始画面から手順書画面への移行を指示するコマンド
'''     パラメータなし。
'''     コマンド50を返す。
'''     旧php版file_sel.php参考。
''' </remarks>
Public Class cCom17
    Inherits cComBase   '基底の継承を宣言

    ''' <summary>コマンドを実行する。</summary>
    ''' <remarks>                    </remarks>
    Public Overrides Sub DoCom()
        Try
            '指示者タブレット情報を更新する
            _設定値.TabletList.TabletUpdate("17", _SocketState)

            '現在のcd_finmodeを確認する。
            Dim cd_finmode As String = "N"
            Dim dt As DataTable = GetDT_Select("SELECT ""cd_finmode"" FROM ""t_sno""")
            If 0 = dt.Rows.Count Then   'データが無かった時は
                ErrMsgBox(Nothing, _設定値.DebugMode, "t_snoテーブルの検索結果がゼロ件です（コマンド17）。暫定的に「手順開始前」として応答します。", 1, _設定値.LogPath)
            Else                        'データがあった時は
                cd_finmode = dt.Rows(0)("cd_finmode").ToString()
            End If

            'まず、開始指示が出た事を記録する
            DBNonQuery("UPDATE ""t_sno"" SET ""cd_finmode""='U'")

            ReturnCom(_ThroughCom, "", "17")   '指示者へ返信 

            '次現状の手順状態で応答を変える
            Select Case cd_finmode
                Case "N" '手順開始前
                    '確認者タブレットは指示者からの手順開始を待機している状態なので手順開始指示の61@を送る。
                    SendTo確認者Tablet("61", "1")  '確認者タブレットへ送信
                    Exit Select
                Case "U" '手順開始状態 
                    '既に手順が始まっているときは途中からの再開状態を意味する。
                    'この時の管理者タブレットは起動時に直接手順書画面へ移行する動作をするので
                    '指示者からの手順書開始指示を待つ状態に無いため、61@を送る必要は無い。
                    Exit Select
            End Select

            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "コマンド17の実行時に例外が発生しました。", 2, _設定値.LogPath)
            ReturnCom(_ErrCom, "", "17")     'エラーを返信
        End Try
    End Sub

End Class

#End Region

#Region "確認者タブレット⇒サーバ"

''' <summary>コマンド20受信処理クラス</summary>
''' <remarks>
'''     確認者タブレットの起動通知コマンド。
'''     パラメータなし。
'''     コマンド61を返す。
'''     ResParam関数で返答データを返す。
'''     旧php版output_wait.php参考。
''' </remarks>
Public Class cCom20
    Inherits cComBase   '基底の継承を宣言

    ''' <summary>コマンドを実行する。</summary>
    ''' <remarks>                    </remarks>
    Public Overrides Sub DoCom()
        Try
            '指示者タブレット情報を更新する
            _設定値.TabletList.TabletUpdate("20", _SocketState)

            '完全終了状態かどうかを確認する
            If Is手順完全終了() Then          '完全終了状態ならば
                ReturnCom("9C", "", "30")     '電源OFF画面コマンドを返信 
                Return
            End If

            'コマンド61処理を行う
            '画面モードを移行する
            Dim Parram61 As String = 0

            '画面モードを確認
            Dim dt As DataTable = GetDT_Select("SELECT ""cd_finmode"" FROM ""t_sno""")
            If 0 = dt.Rows.Count Then   'データが無かった時は
                ErrMsgBox(Nothing, _設定値.DebugMode, "t_snoテーブルの検索結果がゼロ件です（コマンド20）。暫定的に「手順開始前」として応答します。", 1, _設定値.LogPath)
                Parram61 = "0"
            Else                        'データがあった時は
                'Param.画面番号 = dt.Rows(0)("cd_finmode").ToString  '先頭行だけ返す
                Select Case dt.Rows(0)("cd_finmode").ToString
                    Case "U" : Parram61 = "1"   '手順開始
                    Case "N"                    '手順開始前
                End Select
            End If

            If Not (1 = Parram61) Then    '手順開始前なら
            End If

            ReturnCom("61", Parram61, "20")   '返信 

            '指示者へコマンド57を送信
            SendCom57()

            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "コマンド20の実行時に例外が発生しました。", 2, _設定値.LogPath)
            ReturnCom(_ErrCom, "", "20")     'エラーを返信
        End Try
    End Sub

#Region "返答用JSONクラス"

    ''' <summary>コマンド61クラス        </summary>
    ''' <remarks>画面番号を格納するクラス</remarks>
    Public Class cCom61 '未使用
        ''' <summary>画面番号</summary>
        Public Property 画面番号 As String
    End Class

    ' ''' <summary>コマンド57クラス                                        </summary>
    ' ''' <remarks>タブレットの状況と手順書配列(プレビュー)を格納するクラス</remarks>
    'Private Class cCom57
    '    ''' <summary>タブレットの状況を複数格納する配列</summary>
    '    Public Property tablet As cタブレット状況()
    'End Class

#End Region

End Class

''' <summary>コマンド21受信処理クラス</summary>
''' <remarks>
'''     手順書データ要求コマンド。
'''     パラメータなし。
'''     コマンド62を返す。
'''     ResParam関数で返答データを返す。
'''     旧php版K1.php参考。
''' </remarks>
Public Class cCom21
    Inherits cComBase   '基底の継承を宣言

    ''' <summary>コマンドを実行する。</summary>
    ''' <remarks>                    </remarks>
    Public Overrides Sub DoCom()
        Try
            '指示者タブレット情報を更新する
            _設定値.TabletList.TabletUpdate("21", _SocketState)

            '完全終了状態かどうかを確認する
            If Is手順完全終了() Then          '完全終了状態ならば
                ReturnCom("9C", "", "30")     '電源OFF画面コマンドを返信 
                Return
            End If

            Dim Param As New cCom62 '返答に用いるパラメータオブジェクト

            Param.tejun = Get全手順データ()   '手順書を得る

            '手順の現状を得る
            Param.t_sno = Get現在手順情報()

            ReturnCom("62", JsonConvert.SerializeObject(Param), "21")   '返信 

            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "コマンド21の実行時に例外が発生しました。", 2, _設定値.LogPath)
            ReturnCom(_ErrCom, "", "21")     'エラーを返信
        End Try
    End Sub

#Region "返答用JSONクラス"

    ''' <summary>コマンド62クラス          </summary>
    ''' <remarks>手順書配列を格納するクラス</remarks>
    Public Class cCom62
        ''' <summary>手順書の配列</summary>
        Public Property tejun As c手順書プレビュー()
        ''' <summary>手順の現状を格納した型</summary>
        Public Property t_sno As c指示現状
    End Class

#End Region

End Class

''' <summary>コマンド22受信処理クラス</summary>
''' <remarks>
'''     確認操作コマンド。
'''     パラメータあり。
'''     コマンド6Nを返す。
'''     ResParam関数で返答データを返す。
'''     旧php版set_kakunin.php参考。
''' </remarks>
Public Class cCom22
    Inherits cComBase   '基底の継承を宣言

    ''' <summary>手順書番号</summary>
    Public Property 手順書番号 As String

    ''' <summary>コマンドを実行する。</summary>
    ''' <remarks>                    </remarks>
    Public Overrides Sub DoCom()
        Dim トランザクション中 As Boolean = False
        Try
            '指示者タブレット情報を更新する
            _設定値.TabletList.TabletUpdate("22", _SocketState)

            '受信した手順書番号が数値に変換できるか確認
            Try
                Dim tmpLong As Long = CLng(手順書番号.Trim)
            Catch ex As Exception
                ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド22の処理に失敗しました。受信パラメータが数値に変換できません。「" & 手順書番号 & "」", 1, _設定値.LogPath)
                ReturnCom(_ErrCom, "", "22")     'エラーコマンドを返信
                Return
            End Try

            '手順書番号 = _SrcParam 'JSON変換しない場合

            DBNonQuery("BEGIN")
            トランザクション中 = True

            Dim Sql1 As String = "SELECT * FROM ""t_sno"""
            Dim dt1 As DataTable = GetDT_Select(Sql1)
            If 0 = dt1.Rows.Count Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド22の処理に失敗しました。t_snoテーブルが空です。クエリ「" & Sql1 & "」にヒットしませんでした。", 1, _設定値.LogPath)
                ReturnCom(_ErrCom, "", "22")     'エラーコマンドを返信
                Return
            End If
            Dim in_sno As String = dt1.Rows(0)("in_csno").ToString
            Dim in_bno As String = dt1.Rows(0)("in_bno").ToString
            Dim fn As String = dt1.Rows(0)("tx_sfile").ToString

            ''手順書番号のチェックを行う
            'If Not (in_sno.Trim = 手順書番号) Then    '手順書番号が正しくなければ
            '    '前回受信した手順書番号と同じならば重複コマンドとして登録されているタイムスタンプを返す
            '    If _設定値.StaticCom22手順書番号.Trim = 手順書番号.Trim Then
            '        '既に登録されている日時を60@で返す
            '        Dim Sql4 As String = "SELECT ""ts_b"",""in_line"" FROM ""t_order"" WHERE ""in_sno""=" & 手順書番号.Trim
            '        Dim dt4 As DataTable = GetDT_Select(Sql4)
            '        If 0 = dt4.Rows.Count Then
            '            ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド22の処理に失敗しました。(重複コマンド処理時)コマンド60で返す時刻を得る際にt_orderテーブルに該当手順が見つかりませんでした。クエリ「" & Sql4 & "」にヒットしませんでした。", 1, _設定値.LogPath)
            '            ReturnCom(_ErrCom, "", "22")     'エラーコマンドを返信
            '            Return
            '        End If
            '        Dim Com60_2 As New cCom60
            '        Com60_2.ts_b = dt4.Rows(0)("ts_b").ToString

            '        Dim ResParam_2 As String = JsonConvert.SerializeObject(Com60_2)
            '        ReturnCom("60", ResParam_2, "22")    '確認者へ返信 
            '        ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド22で受信した手順書の番号が前回と重複していたので、DBを変更せずにコマンド60を返しました。", 0, _設定値.LogPath)
            '        Return
            '    End If

            '    'If CLng(手順書番号.Trim) < CLng(in_sno.Trim) Then    '受信した手順書番号がDB側の手順番号より小さい時は
            '    '    '既に登録されている日時を60@で返す
            '    '    Dim Sql4 As String = "SELECT ""ts_b"",""in_line"" FROM ""t_order"" WHERE ""in_sno""=" & 手順書番号.Trim
            '    '    Dim dt4 As DataTable = GetDT_Select(Sql4)
            '    '    If 0 = dt4.Rows.Count Then
            '    '        ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド22の処理に失敗しました。(手順不一致処理時)コマンド60で返す時刻を得る際にt_orderテーブルに該当手順が見つかりませんでした。クエリ「" & Sql4 & "」にヒットしませんでした。", 1, _設定値.LogPath)
            '    '        ReturnCom(_ErrCom, "", "22")     'エラーコマンドを返信
            '    '        Return
            '    '    End If
            '    '    Dim Com60_2 As New cCom60
            '    '    Com60_2.ts_b = dt4.Rows(0)("ts_b").ToString

            '    '    Dim ResParam_2 As String = JsonConvert.SerializeObject(Com60_2)
            '    '    ReturnCom("60", ResParam_2, "22")    '確認者へ返信 
            '    '    ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド22で受信した手順書の番号がDBがもつ手順書の番号より小さかったので、コマンド60でタブレット側の手順を進めました。", 0, _設定値.LogPath)
            '    '    Return
            '    'Else
            '    ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド22で受信した手順書の番号がDBと一致しません。", 0, _設定値.LogPath)
            '    ReturnCom("6R", "", "22")               '確認者タブレットへ送信
            '    Return                                  '抜ける
            '    'End If
            'End If

            '指示前なら何もせずに終了
            If "0" = dt1.Rows(0)("cd_status") Then
                ReturnCom(_ThroughCom, "", "22")   '返信 
                Return
            End If

            Dim Sql2 As String = "SELECT ""in_bno"",""in_swno"",""in_jotai_no"",""cd_flip_after"",""cd_pair"",""bo_gs"",""ts_gs"",""in_line"" FROM ""t_order"" WHERE ""in_sno""=" & in_sno
            Dim dt2 As DataTable = GetDT_Select(Sql2)
            If 0 = dt2.Rows.Count Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド22の処理に失敗しました。t_order上に該当手順が見つかりません。クエリ「" & Sql2 & "」にヒットしませんでした。", 1, _設定値.LogPath)
                ReturnCom(_ErrCom, "", "22")     'エラーコマンドを返信
                Return
            End If

            Dim in_bno2 As String = dt2.Rows(0)("in_bno").ToString
            Dim in_swno As String = dt2.Rows(0)("in_swno").ToString
            Dim in_jotai_no As String = dt2.Rows(0)("in_jotai_no").ToString
            Dim cd_flip_after As String = dt2.Rows(0)("cd_flip_after").ToString
            Dim cd_pair As String = dt2.Rows(0)("cd_pair").ToString
            Dim bo_gs As Boolean = ("t" = dt2.Rows(0)("bo_gs").ToString)
            'Dim ts_gs As DateTime = CDate(dt2.Rows(0)("ts_gs"))
            Dim in_line As String = dt2.Rows(0)("in_line").ToString

            Dim in_sno_next As Long = CLng(in_sno) + 1    '次の手順番号

            DBNonQuery("UPDATE ""t_sno"" SET ""in_csno""=" & in_sno_next & ",""in_tsno""=-1,""in_bno""=-1,""cd_status""='0',""cd_pctl""='0',""cd_gsmode""='0'")
            DBNonQuery("UPDATE ""t_order"" SET ""cd_status""='7',""ts_b""='now',""bo_fin""='t' WHERE ""in_sno""=" & in_sno)

            '60@で返す日時を得る
            Dim Sql3 As String = "SELECT ""ts_b"",""in_line"" FROM ""t_order"" WHERE ""in_sno""=" & in_sno
            Dim dt3 As DataTable = GetDT_Select(Sql3)
            If 0 = dt3.Rows.Count Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド22の処理に失敗しました。コマンド60で返す時刻を得る際にt_orderテーブルに該当手順が見つかりませんでした。クエリ「" & Sql3 & "」にヒットしませんでした。", 1, _設定値.LogPath)
                ReturnCom(_ErrCom, "", "22")     'エラーコマンドを返信
                Return
            End If
            Dim Com60 As New cCom60
            Com60.ts_b = dt3.Rows(0)("ts_b").ToString

            '対のある動作なら、この機器の状態フラグを指示書の通りに書き換える

            If 1 = cd_pair Then
                Dim Sql4 As String = "SELECT ""in_bno"",""in_sw_no"",""in_pair_no"" FROM ""t2"" WHERE ""in_line""=" & in_line
                Dim dt4 As DataTable = GetDT_Select(Sql4)
                If 0 = dt4.Rows.Count Then
                    ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド22の処理に失敗しました。機器の状態を変更する際に目的の機器がt2テーブルに見つかりませんでした。クエリ「" & Sql4 & "」にヒットしませんでした。", 1, _設定値.LogPath)
                    ReturnCom(_ErrCom, "", "22")     'エラーコマンドを返信
                    Return
                End If
                For Each e As DataRow In dt4.Rows
                    DBNonQuery("UPDATE ""m_device2"" SET ""cd_flip""='" & cd_flip_after & "' WHERE ""in_bno""=" & e("in_bno").ToString & " AND ""in_sw_no""=" & e("in_sw_no").ToString & " AND ""in_jotai_no""=" & e("in_pair_no").ToString)
                Next
            End If

            Set_Ban_Tables2()       '盤の状態をリセットする？

            Set_Ban_Clear(in_bno2)  '盤の背景色を元に戻す

            '★PHP版ではここに調査用ログ出力あり

            If Not (Skip_Comment()) Then 'コメント手順行をスキップ
                '失敗時はエラーを返す
                ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド22の処理に失敗しました。スキップ処理(Skip_Comment関数)内の問題です。", 1, _設定値.LogPath)
                ReturnCom(_ErrCom, "", "22")     'エラーコマンドを返信
                Return
            End If

            _設定値.LastCommandList.Remove指示者確認者Command()  '直近のコマンドをクリア

            Dim ResParam As String = JsonConvert.SerializeObject(Com60)
            ReturnCom("60", ResParam, "22")    '確認者へ返信 
            SendTo指示者Tablet("55", ResParam) '指示者へ送信

            '盤タブレットへコマンドを送信
            SendCom7x("73", in_bno) '操作前の盤タブレットへ送信

            'レポートにタイムスタンプを記述
            WriteReport(CInt(dt3.Rows(0)("in_line").ToString), 4, DateTime.Parse(dt3.Rows(0)("ts_b").ToString()).ToString("yyyy/MM/dd HH:mm"))

            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "コマンド22で例外が発生しました。", 2, _設定値.LogPath)
            ReturnCom(_ErrCom, "", "22")     'エラーコマンドを返信
        Finally
            If トランザクション中 Then DBNonQuery("COMMIT")
            _設定値.StaticCom22手順書番号 = 手順書番号.Trim  '今回受信した手順書番号を保存
        End Try
    End Sub

#Region "返答用JSONクラス"

    ''' <summary>コマンド60クラス                </summary>
    ''' <remarks>確認を行った時刻を格納するクラス</remarks>
    Private Class cCom60
        ''' <summary>確認時刻＝手順完了時刻</summary>
        Public Property ts_b As String
    End Class

#End Region

End Class

''' <summary>コマンド23受信処理クラス</summary>
''' <remarks>
'''     現場差異コマンド。
'''     パラメータなし。
'''     コマンド6Sを返す。
'''     ResParam関数で返答データを返す。
'''     旧php版set_gs.php、func.ini参考。
''' </remarks>
Public Class cCom23
    Inherits cComBase   '基底の継承を宣言

    '''' <summary>手順書番号</summary>
    Public Property in_sno As String

    ''' <summary>コマンドを実行する。</summary>
    ''' <remarks>                    </remarks>
    Public Overrides Sub DoCom()
        Dim トランザクション中 As Boolean = False
        Dim Transaction As Npgsql.NpgsqlTransaction
        Try
            トランザクション中 = True
            DB_BeginTransaction(Transaction)
            'DBNonQuery("BEGIN") : トランザクション中 = True
            If Transaction Is Nothing Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド23でトランザクションを開始できませんでした。", 2, _設定値.LogPath)
                ReturnCom(_ErrCom, "", "23")     'エラーを返信
                Return
            End If

            '指示者タブレット情報を更新する
            _設定値.TabletList.TabletUpdate("23", _SocketState)

            
            '★ここにエクセル関連の処理

            '●ここでは現場差異の確認操作のみなので
            '●旧php版のgs_skip関数(func.ini参照)を実行する(要確認)

            
            Dim Sql1 As String = "SELECT * FROM ""t_sno"""
            Dim dt1 As DataTable = GetDT_Select(Sql1)
            If 0 = dt1.Rows.Count Then Err.Raise(9999, Nothing, "コマンド23の処理に失敗しました。クエリ「" & Sql1 & "」にヒットしませんでした。")
            Dim in_csno As String = dt1.Rows(0)("in_csno").ToString
            Dim cd_status As String = dt1.Rows(0)("cd_status").ToString
            Dim in_bno As String = dt1.Rows(0)("in_bno").ToString
            Dim cd_gsmode As String = dt1.Rows(0)("cd_gsmode").ToString

            If Not (in_csno.Trim = in_sno.Trim) Then    '手順書番号が正しくなければ
                ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド23で受信した手順書の番号がDBと一致しません。", 0, _設定値.LogPath)
                ReturnCom("6R", "", "23")               '確認者タブレットへ送信
                Return                                  '抜ける
            End If

            Dim tx_gs As String = ""
            Select Case cd_gsmode
                Case "0"  'キャンセル
                    '指示者がキャンセルした現場差異指示が確認者タブレットに反映されないまま確認者が発呼した現場差異確認コマンドと見なす事ができるので
                    ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド23で現場差異指示が既にキャンセルされているので現場差異確認コマンドを中断します。", 0, _設定値.LogPath)
                    ReturnCom("6R", "", "23")               '確認者タブレットへ送信
                    Return                                  '抜ける
                Case "5"  'スキップ
                    tx_gs = _GS_TBL_5
                    DBNonQuery("UPDATE ""t_order"" SET ""bo_fin""='t',""cd_status""='7',""bo_gs""='t',""ts_gs""='now',""tx_gs""='" & tx_gs & "' WHERE ""in_sno""=" & in_csno)
                    Set_Ban_Clear(in_bno) '該当している版タブレットの黄色と点滅をキャンセルする

                    'ポインタを薦める
                    in_csno += 1
                    DBNonQuery("UPDATE ""t_sno"" SET ""in_csno""=" & in_csno & ",""in_tsno""=-1,""in_bno""=-1,""cd_status""='0',""cd_pctl""='0',""cd_gsmode""='0'")

                    Skip_Comment()  'コメント行を飛ばす
                Case "6"  '前に追加
                    tx_gs = _GS_TBL_6
                    DBNonQuery("UPDATE ""t_order"" SET ""bo_gs""='t',""ts_gs""='now',""tx_gs""='" & tx_gs & "' WHERE ""in_sno""=" & in_csno)
                    DBNonQuery("UPDATE ""t_sno"" SET ""cd_gsmode""='9' ;")
            End Select

            Dim TimeStampDone As Date = Now

            _設定値.LastCommandList.Remove指示者確認者Command()  '直近のコマンドをクリア

            ReturnCom("6N", "", "23")    '返信 
            SendTo指示者Tablet("56", "") '指示者へ送信する

            '盤タブレットへコマンドを送信
            SendCom7x("73", in_bno) '操作前の盤タブレットへ送信

            '通信が終わってからレポートにタイムスタンプを記述する
            Dim Sql3 As String = "SELECT ""ts_b"",""in_line"" FROM ""t_order"" WHERE ""in_sno""=" & in_csno
            Dim dt3 As DataTable = GetDT_Select(Sql3)
            If 0 = dt3.Rows.Count Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド23の処理に失敗しました。レポートの記述位置を取得する際にt_orderテーブルに該当手順が見つかりませんでした。クエリ「" & Sql3 & "」にヒットしませんでした。", 1, _設定値.LogPath)
                Return
            End If
            WriteReport(CInt(dt3.Rows(0)("in_line").ToString), 11, TimeStampDone.ToString("yyyy/MM/dd HH:mm"))

            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "コマンド23で例外が発生しました。", 2, _設定値.LogPath)
            ReturnCom(_ErrCom, "", "23")     'エラーを返信
        Finally
            Try
                If トランザクション中 Then
                    'DBNonQuery("COMMIT")
                    If Not (Transaction Is Nothing) Then
                        If Not (Transaction.Connection Is Nothing) Then Transaction.Commit()
                        Transaction.Dispose()
                        Transaction = Nothing
                    End If
                End If
            Catch ex As Npgsql.NpgsqlException
            End Try
        End Try
    End Sub

#Region "返答用JSONクラス"

    ''' <summary>コマンド6Nクラス                </summary>
    ''' <remarks>確認を行った時刻を格納するクラス</remarks>
    Private Class cCom6N
        ''' <summary>タブレットの状況を複数格納する配列</summary>
        Public Property ts_gs As String
    End Class

#End Region

End Class

#End Region

#Region "盤タブレット⇒サーバ"

''' <summary>コマンド30受信処理クラス</summary>
''' <remarks>
'''     盤タブレットの盤リスト要求コマンド。
'''     パラメータなし。
'''     盤に登録済みならコマンド72を返す。そうでないなら71@を返す。
'''     ResParam関数で返答データを返す。
'''     旧php版ban_Sel.php参考。
''' </remarks>
Public Class cCom30
    Inherits cComBase   '基底の継承を宣言

    ''' <summary>コマンドを実行する。</summary>
    ''' <remarks>                    </remarks>
    Public Overrides Sub DoCom()
        Try
            '今回の盤タブレット情報を更新する
            _設定値.TabletList.TabletUpdate("30", _SocketState)

            '完全終了状態かどうかを確認する
            If Is手順完全終了() Then          '完全終了状態ならば
                ReturnCom("9C", "", "30")     '電源OFF画面コマンドを返信 
                Return
            End If

            'リモートIPが既に登録されているかどうかを確認する
            '既に盤にタブレットが登録されているか調べる
            Dim RemoteIP As String = CType(_SocketState.WorkSocket.RemoteEndPoint, IPEndPoint).Address.ToString '今回の盤タブレットのIP
            Dim dt盤マスタ1 As DataTable = GetDT_Select("SELECT * FROM ""m_ban"" WHERE ""tx_rip""='" & RemoteIP & "';")
            If dt盤マスタ1.Rows.Count < 1 Then
                '盤が登録済みでないなら
                '盤選択用に71@を返す。
                Dim Dest As c盤一覧情報 = Get盤選択一覧情報()
                If Dest Is Nothing Then
                    ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド30で盤タブレットへの送信情報(盤一覧情報)の作成に失敗しました。", 1, _設定値.LogPath)
                    ReturnCom(_ErrCom, "", "31")     'エラーを返信
                    Return
                End If
                ReturnCom("71", JsonConvert.SerializeObject(Dest), "30")   '返信 
                Return  '戻る
            Else
                '盤にIPが登録済みなら
                DBNonQuery("UPDATE ""m_ban"" SET ""bo_set""='t' WHERE ""tx_rip""='" & RemoteIP & "' ;") '盤タブレットの起動を登録
                '盤詳細表示用に72@を返す
                '盤内情報を取得する
                Dim Dest As c盤表示詳細 = Get盤タブレット表示情報(dt盤マスタ1.Rows(0)("tx_rip").ToString)
                If Dest Is Nothing Then
                    ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド30で盤タブレットへの送信情報の作成(盤内情報)に失敗しました。", 1, _設定値.LogPath)
                    ReturnCom(_ErrCom, "", "30")     'エラーを返信
                    Return
                End If
                ReturnCom("72", JsonConvert.SerializeObject(Dest), "31")   '返信 

                '指示者タブレットへコマンド57を送信
                SendCom57()

                Return  '戻る
            End If
            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "コマンド30で例外が発生しました。", 2, _設定値.LogPath)
            ReturnCom(_ErrCom, "", "30")     'エラーを返信
        End Try
    End Sub

    '#Region "返答用JSONクラス"

    '    ''' <summary>コマンド71クラス        </summary>
    '    ''' <remarks>盤一覧を格納するクラス</remarks>
    '    Public Class cCom71

    '        ''' <summary>盤一覧</summary>
    '        Public Property 盤情報 As New Hashtable

    '    End Class

    '    ''' <summary>盤情報クラス                  </summary>
    '    ''' <remarks>一つの盤の情報を格納するクラス</remarks>
    '    Public Class c盤情報

    '        ''' <summary>盤名</summary>
    '        Public Property tx_bname As String

    '        ''' <summary>盤の通し番号</summary>
    '        Public Property in_bno As String

    '        ''' <summary>盤タブレットのIPアドレス</summary>
    '        Public Property tx_rip As String

    '    End Class

    '#End Region

End Class

''' <summary>コマンド31受信処理クラス</summary>
''' <remarks>
'''     盤タブレットの単一盤情報要求コマンド。
'''     パラメータなし。
'''     コマンド72を返す。
'''     ResParam関数で返答データを返す。
'''     旧php版？？？？？？？？？.php参考。
''' </remarks>
Public Class cCom31
    Inherits cComBase   '基底の継承を宣言

    ''' <summary>盤の通し番号</summary>
    Public Property in_bno As String

    ''' <summary>コマンドを実行する。</summary>
    ''' <remarks>                    </remarks>
    Public Overrides Sub DoCom()
        Try
            '今回の盤タブレット情報を更新する
            _設定値.TabletList.TabletUpdate("31", _SocketState)

            '完全終了状態かどうかを確認する
            If Is手順完全終了() Then          '完全終了状態ならば
                ReturnCom("9C", "", "30")     '電源OFF画面コマンドを返信 
                Return
            End If

            '実行条件を確認する
            If in_bno Is Nothing Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "盤ID「in_bno」がnothingです。パラメータを確認して下さい。", 1, _設定値.LogPath)
                ReturnCom(_ErrCom, "", "31")     'エラーを返信
                Return
            End If
            If "" = in_bno Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "盤ID「in_bno」が与えられていません。", 1, _設定値.LogPath)
                ReturnCom(_ErrCom, "", "31")     'エラーを返信
                Return
            End If

            Dim RemoteIP As String = CType(_SocketState.WorkSocket.RemoteEndPoint, IPEndPoint).Address.ToString '今回の盤タブレットのIP

            '盤に既にタブレットが登録されているか調べる
            '盤が存在するかどうか調べる
            Dim dt盤マスタ0 As DataTable = GetDT_Select("SELECT * FROM ""m_ban"" WHERE ""in_bno""=" & in_bno)
            If dt盤マスタ0.Rows.Count < 1 Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "与えられた盤ID「" & in_bno & "」の盤が見つかりませんでした。(IP確認)", 1, _設定値.LogPath)
                ReturnCom(_ErrCom, "", "31")     'エラーを返信
                Return
            End If
            Dim Old_tx_rip As String = dt盤マスタ0(0)("tx_rip").ToString   '更新前のIPアドレス
            Dim Update_RemoteIp As Boolean = False  'このタブレットのIPを登録するかどうかのフラグ
            If "" = Old_tx_rip Then Update_RemoteIp = True '旧IPが空なら未登録の盤として更新対象とする
            If Not (Old_tx_rip = RemoteIP) Then Update_RemoteIp = True 'このタブレットが旧IPと違うなら更新対象とする

            If Update_RemoteIp Then '登録IPを更新するなら
                '既に盤にタブレットが登録されているなら上書きする
                DBNonQuery("UPDATE ""m_ban"" SET  ""tx_rip""='' , ""bo_set""='f' WHERE ""tx_rip""='" & RemoteIP & "';")  '既に同じIPで登録されている盤があるならそのIPをクリアする
                If DBNonQuery("UPDATE ""m_ban"" SET ""tx_rip""='" & RemoteIP & "', ""bo_set""='t' WHERE ""in_bno""=" & in_bno & ";") < 1 Then
                    ErrMsgBox(Nothing, _設定値.DebugMode, "盤情報に盤タブレットのIPを登録できませんでした。与えられた盤IDが存在しないと思われます。", 1, _設定値.LogPath)
                    ReturnCom(_ErrCom, "", "31")     'エラーを返信
                    Return
                End If
            End If

            '盤内情報を取得する
            Dim Dest As c盤表示詳細 = Get盤タブレット表示情報(RemoteIP)
            If Dest Is Nothing Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド31で盤タブレットへの送信情報の作成に失敗しました。", 1, _設定値.LogPath)
                ReturnCom(_ErrCom, "", "31")     'エラーを返信
                Return
            End If

            Dim DestParam As String = JsonConvert.SerializeObject(Dest)

            ReturnCom("72", DestParam, "31")   '返信 

            '指示者タブレットへコマンド57を送信
            SendCom57()

            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "コマンド31で例外が発生しました。", 2, _設定値.LogPath)
            ReturnCom(_ErrCom, "", "31")     'エラーを返信
        End Try
    End Sub

    '#Region "返答用JSONクラス"

    '    ''' <summary>コマンド72クラス            </summary>
    '    ''' <remarks>盤の詳細情報を格納するクラス</remarks>
    '    Public Class cCom72

    '        ''' <summary>盤の番号</summary>
    '        Public Property in_bno As String

    '        ''' <summary>盤名</summary>
    '        Public Property tx_bname As String

    '        ''' <summary>機器数</summary>
    '        Public Property lines As String

    '        ''' <summary>盤を開いて操作中？のフラグ？</summary>
    '        Public Property bo_active As String

    '        ''' <summary>盤を開いて操作中？のフラグ？</summary>
    '        Public Property m_device As cM_device()

    '    End Class

    '    ''' <summary>盤情報クラス                  </summary>
    '    ''' <remarks>一つの盤の情報を格納するクラス</remarks>
    '    Public Class cM_device

    '        ''' <summary>機器の名称</summary>
    '        Public Property tx_lb As String

    '        ''' <summary>色</summary>
    '        Public Property tx_clr As String

    '        ''' <summary>太線枠が点滅状態か</summary>
    '        Public Property in_disp_hi As String

    '        ''' <summary>点滅状態</summary>
    '        Public Property in_disp_blink As String

    '    End Class

    '#End Region

End Class

''' <summary>コマンド32受信処理クラス</summary>
''' <remarks>
'''     盤タブレットの盤再選択コマンド。
'''     パラメータなし。
'''     コマンド71を返す。
'''     ResParam関数で返答データを返す。
'''     旧php版？？？？？？？？？.php参考。
''' </remarks>
Public Class cCom32
    Inherits cComBase   '基底の継承を宣言

    ''' <summary>コマンドを実行する。</summary>
    ''' <remarks>                    </remarks>
    Public Overrides Sub DoCom()
        Try
            '今回の盤タブレット情報を更新する
            _設定値.TabletList.TabletUpdate("32", _SocketState)

            '実行条件を確認する
            '既に盤にタブレットが登録されているならクリアする
            Dim RemoteIP As String = CType(_SocketState.WorkSocket.RemoteEndPoint, IPEndPoint).Address.ToString '今回の盤タブレットのIP
            DBNonQuery("UPDATE ""m_ban"" SET  ""tx_rip""='' , ""bo_set""='f' WHERE ""tx_rip""='" & RemoteIP & "';")  '既に同じIPで登録されている盤があるならそのIPをクリアする

            '盤一覧を返す
            Dim Dest As c盤一覧情報 = Get盤選択一覧情報()
            If Dest Is Nothing Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド32で盤タブレットへの送信情報(盤一覧情報)の作成に失敗しました。", 1, _設定値.LogPath)
                ReturnCom(_ErrCom, "", "32")     'エラーを返信
                Return
            End If
            Dim DestParam As String = JsonConvert.SerializeObject(Dest) 'JSONシリアライズする
            ReturnCom("71", DestParam, "32")   '返信 

            '指示者タブレットへコマンド57を送信
            SendCom57()

            '未登録の盤タブレットへ71を送信する
            'SendTo未登録盤Tablet("71", DestParam)

            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "コマンド32で例外が発生しました。", 2, _設定値.LogPath)
            ReturnCom(_ErrCom, "", "32")     'エラーを返信
        End Try
    End Sub

End Class

#End Region

#Region "各タブレット⇒サーバ"

''' <summary>コマンド90受信処理クラス</summary>
''' <remarks>
'''     各タブレットからの直近コマンド再送要求コマンド。
'''     パラメータなし。
'''     コマンド9N@または9Q@を返す。
'''     ResParam関数で返答データを返す。
''' </remarks>
Public Class cCom90
    Inherits cComBase   '基底の継承を宣言

    ''' <summary>コマンドを実行する。</summary>
    ''' <remarks>                    </remarks>
    Public Overrides Sub DoCom()
        Try
            Dim RemoteIP As String = CType(_SocketState.WorkSocket.RemoteEndPoint, IPEndPoint).Address.ToString '今回のタブレットのIP

            '直近のコマンドを得る
            Dim LastCommand As cLastCommandList.cLastCommand = _設定値.LastCommandList.GetLastCommand(RemoteIP.Trim)
            If LastCommand Is Nothing Then
                ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド90で直近のコマンドがヒットしなかったので再送要求に応えられませんでした。", 0, _設定値.LogPath)
                ReturnCom("9Q", "", "90")     '再送忌避を返信
                Return
            End If

            ''実行条件を確認する
            ''現在の手順の指示情報を得る
            'Dim Sql1 As String = "SELECT * FROM ""t_sno"""
            'Dim dt1 As DataTable = GetDT_Select(Sql1)
            'If 0 = dt1.Rows.Count Then
            '    ErrMsgBox(Nothing, _設定値.DebugMode, "コマンド90でクエリ「" & Sql1 & "」にヒットしませんでした。", 1, _設定値.LogPath)
            '    ReturnCom(_ErrCom, "", "90")     'エラーを返信
            '    Return
            'End If

            ''現在の値
            'Dim cd_status As String = dt1.Rows(0)("cd_status").ToString
            'Dim cd_gsmode As String = dt1.Rows(0)("cd_gsmode").ToString

            ''再送に不適切かどうかを判断する
            'If Not ("0" = cd_gsmode) Then '現場差異の指示後確認前の時
            '    'DBが現場差異の指示後確認前の時に指示者タブレットに確認完了コマンドを再送するべきではない。
            '    Select Case LastCommand.Command '直近のコマンドが
            '        Case "55", "56"
            '            ReturnCom("9Q", "", "90")   '返信 
            '            Return
            '    End Select
            'Else    '現場差異モードではない時
            '    Select Case cd_status
            '        Case "0"    '確認後、次の指示前(初期値)
            '            'DBが確認後指示前の時に確認者タブレットに指示済みコマンドを再送するべきではない。
            '            Select Case LastCommand.Command '直近のコマンドが
            '                Case "63", "64"   '63:指示済み、64:現場差異指示済み
            '                    ReturnCom("9Q", "", "90")   '返信 
            '                    Return
            '            End Select
            '        Case "1"    '指示後、確認前
            '            'DBが指示後確認前の時に指示者タブレットに確認完了コマンドを再送するべきではない。
            '            Select Case LastCommand.Command '直近のコマンドが
            '                Case "55", "56"   '55:手順の確認完了、56:現場差異の確認完了
            '                    ReturnCom("9Q", "", "90")   '返信 
            '                    Return
            '            End Select
            '    End Select
            'End If

            ReturnCom("9N", "", "90")   '返信 

            Send(RemoteIP, _設定値.RemotePort, LastCommand.Command, LastCommand.Param) '再送

            Return
        Catch ex As Exception
            ErrMsgBox(ex, _設定値.DebugMode, "コマンド90で例外が発生しました。", 2, _設定値.LogPath)
            ReturnCom(_ErrCom, "", "32")     'エラーを返信
        End Try
    End Sub

End Class





#End Region

#End Region

#End Region

#Region "その他"

''' <summary>各種設定値をパッケージしたクラス</summary>
''' <remarks>                                </remarks>
Public Class c設定値

#Region "値"

#Region "設定ファイルから設定される値"

    '''<summary>ポート番号</summary>
    Public ReadPort As Integer

    '''<summary>リモートのポート番号</summary>
    Public RemotePort As Integer

    '''<summary>受信タイムアウト(ミリ秒)</summary>
    Public ReadTimeOut As Integer

    '''<summary>送信タイムアウト(ミリ秒)</summary>
    Public WriteTimeOut As Integer

    '''<summary>ログ保存先パス</summary>
    Public LogPath As String

    '''<summary>接続タイムアウト(ミリ秒)</summary>
    Public ConnectTimeOut As Integer

    '''<summary>最大接続試行回数</summary>
    Public MaxSendCount As Long

    '''<summary>デバッグモードかどうか</summary>
    '''<remarks></remarks>
    Public DebugMode As Boolean = False

    '''<summary>デバッグメニューモード</summary>
    Public DebugMenuMode As Boolean = False

    '''<summary>手順書をリセットするコマンド</summary>
    '''<remarks></remarks>
    Public ResetTejunCmd As String

    '''<summary>タイムスタンプを記述する手順書ファイルの存在するディレクトリのパス</summary>
    Public ReportDirPath As String

    '''<summary>タイムスタンプを記述する手順書ファイル名</summary>
    Public ReportFileName As String

    '''<summary>タイムスタンプを記述する手順書ファイル（エクセル）のシート名</summary>
    Public ReportSheetName As String

    '''<summary>dumpファイル名</summary>
    Public DumpFileName As String

    '''<summary>dumpを行うコマンド</summary>
    Public DumpCmd As String

#End Region

#Region "内部だけで扱われる値"

    '''<summary>PostgreSQLへの接続文字列</summary>
    Public ConnectText As String

    '''<summary>受信の終わりを表すByte値</summary>
    '''<remarks>初期値はヌルで</remarks>
    Public ReadEndByte As Byte = 0

    '''<summary>コマンド文字列とパラメータ文字列の境界を示すByte値</summary>
    '''<remarks>初期値は「@」ヌルで</remarks>
    Public ReadSeparateByte As Byte = 64

    ''' <summary>タブレットを管理するクラス</summary>
    Public TabletList As New cComBase.cTabletList(Me)

    ''' <summary>直近のコマンドを管理するクラス</summary>
    Public LastCommandList As New cComBase.cLastCommandList(Me)

    ''' <summary>前回実行されたコマンド22@で受信した手順書番号</summary>
    ''' <remarks>重複送信対策</remarks>
    Public StaticCom22手順書番号 As String

#End Region

#End Region

#Region "関数"

    ''' <summary>設定を文字列にする            </summary>
    ''' <returns>このクラスの各値のサマリ文字列</returns>
    ''' <remarks>                              </remarks>
    Public Function toString() As String
        Dim Dest As String = ""
        Try
            Dest &= "受信ポート番号　　　　:" & ReadPort.ToString & vbCrLf
            Dest &= "接続相手ポート番号　　:" & RemotePort.ToString & vbCrLf
            Dest &= "受信タイムアウト(ﾐﾘ秒):" & ReadTimeOut.ToString & vbCrLf
            Dest &= "送信タイムアウト(ﾐﾘ秒):" & WriteTimeOut.ToString & vbCrLf
            Dest &= "ログ保存先パス　　　　:" & LogPath.ToString & vbCrLf
            Dest &= "接続タイムアウト(ﾐﾘ秒):" & ConnectTimeOut.ToString & vbCrLf
            Dest &= "通信接続最大試行回数　:" & MaxSendCount.ToString & vbCrLf
            Dest &= "デバッグモード        :" & DebugMode.ToString & vbCrLf
            Dest &= "デバッグメニューモード:" & DebugMenuMode.ToString & vbCrLf
            Dest &= "手順書リセットコマンド:" & ResetTejunCmd.ToString & vbCrLf
            Dest &= "レポートのディレクトリ:" & ReportDirPath.ToString & vbCrLf
            Dest &= "レポートのファイル名　:" & ReportFileName.ToString & vbCrLf
            Dest &= "レポートのシート名　　:" & ReportSheetName.ToString & vbCrLf
            Dest &= "Dumpファイル名　　　　:" & DumpFileName.ToString & vbCrLf
            Dest &= "Dumpコマンド　　　　　:" & DumpCmd.ToString & vbCrLf
            Return Dest
        Catch ex As Exception
            ErrMsgBox(ex, DebugMode, "設定の文字列変換に失敗しました。", 2, LogPath)
            Return Dest
        End Try
    End Function

#End Region

End Class

#End Region

#Region "メモ類"

'PostgreSQLのレストア時のコマンド
'C:\Program Files\PostgreSQL\9.6\bin\psql -U postgres -d kdk_ban -f C:\Users\gekkoji\Desktop\short.sql
'C:\Program Files\PostgreSQL\9.6\bin\psql -U postgres -d kdk_ban -f C:\Users\gekkoji\Desktop\kdk_ban.sql
'C:\Program Files\PostgreSQL\9.6\bin\psql -U postgres -d kdk_ban -f C:\Users\gekkoji\Desktop\full20161014.sql
'パスワードはy3pevd36

'手順書のリセットコマンド
'c:\php\php.exe c:\ban\reset_cmd.php



#End Region


