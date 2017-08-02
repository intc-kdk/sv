Imports System.Net.Sockets

Module Module1
#Region "ファイルIO"

    ''' <summary>テキストファイルを読込む</summary>
    ''' <param name="FilePath">対象のファイルパス</param>
    ''' <returns>ファイルの全体の文字列</returns>
    ''' <remarks></remarks>
    Public Function ReadTextFile(ByVal FilePath As String) As String
        Dim StreamReader As IO.StreamReader
        Try
            'ファイルの存在を確認
            If Not (System.IO.File.Exists(FilePath)) Then Err.Raise(9999, Nothing, "指定されたファイル「" & FilePath & "」が存在しません。確認して下さい。")

            'ファイルを読み込む
            StreamReader = New IO.StreamReader(FilePath, System.Text.Encoding.Default)
            If StreamReader.Peek() = -1 Then Err.Raise(9999, Nothing, "指定されたファイル「" & FilePath & "」が見つかりましたが何らかの原因で読み込めません。確認してください。") '読み取れない時
            Return StreamReader.ReadToEnd   '全て返す
        Catch ex As Exception
            Err.Raise(9999, Nothing, "テキストファイルの読込に失敗しました。" & vbCrLf & ex.Message)
        Finally
            If Not (StreamReader Is Nothing) Then StreamReader.Close() 'ファイルを閉じる
            StreamReader = Nothing  'ファイルシステムオブジェクトを解放
        End Try
    End Function

    ''' <summary>指定されたテキストファイルの先頭行だけ読む</summary>
    ''' <param name="FilePath">対象のファイルパス</param>
    ''' <returns>先頭の文字列を返す</returns>
    ''' <remarks></remarks>
    Public Function ReadTextFileFirstLine(ByVal FilePath As String) As String
        Try
            Dim Source As String = ReadTextFile(FilePath)       'ファイルを読込む
            Dim StrArray() As String = Split(Source, vbCrLf)    '文字列を改行で配列に分割
            Return StrArray(0)
        Catch ex As Exception
            Err.Raise(9999, Nothing, "テキストファイルの先頭行の読込に失敗しました。" & vbCrLf & ex.Message)
        End Try
    End Function

    ''' <summary>指定されたテキストファイルの末尾に文字列を追記する</summary>
    ''' <param name="FilePath">追記するファイルのパス</param>
    ''' <param name="Src">追記する文字列</param>
    ''' <remarks></remarks>
    Public Sub WriteLineTextFile(FilePath As String, Src As String)
        Dim StreamWriter As System.IO.StreamWriter
        Try
            '保存先ディレクトリの確認
            Dim DirPath As String = System.IO.Path.GetDirectoryName(FilePath)   'ファイルパスからディレクトリパスを得る
            If Not (System.IO.Directory.Exists(DirPath)) Then System.IO.Directory.CreateDirectory(DirPath) 'ディレクトリが存在しないなら作成する
            Dim StartTime As Date = Now
            Do
                'ファイルを書き込む
                Try
                    StreamWriter = New System.IO.StreamWriter(FilePath, True, System.Text.Encoding.GetEncoding("Shift-JIS")) '指定したファイルの末尾に追加するストリームを作る
                    StreamWriter.WriteLine(Src) '文字列を書き込む
                    Exit Do                     'ループを抜ける
                Catch ex As Exception
                    'エラー時はもう一度試す
                End Try
                Dim tmpSpan As TimeSpan = Now - StartTime
                If 1000 < tmpSpan.TotalMilliseconds Then Exit Do '1秒以上記入できないときは抜ける
                System.Threading.Thread.Sleep(100)    '待機
            Loop
            Return
        Catch ex As Exception
            Err.Raise(9999, Nothing, "テキストファイルへの(一行追加)書込みに失敗しました。" & vbCrLf & ex.Message)
        Finally
            If Not (StreamWriter Is Nothing) Then StreamWriter.Close() 'ストリームを閉じる
        End Try
    End Sub

    ''' <summary>ログファイルに追記する関数</summary>
    ''' <param name="DirPath">ログファイルの保存先のディレクトリパス</param>
    ''' <param name="ModeStr">ログに記述する処理の種類</param>
    ''' <param name="RemoteIP">通信相手のIP</param>
    ''' <param name="Src">通信データ、エラー情報の文字列</param>
    ''' <remarks></remarks>
    Public Sub WriteLog(DirPath As String, ModeStr As String, RemoteIP As String, Src As String)
        Try
            '通信ログを保存する

            '保存先を決める
            Dim FileName As String = Today.ToString("yyyyMMdd") & ".log"
            Dim FilePath As String = System.IO.Path.Combine(DirPath, FileName)
            '保存内容を決める
            Dim ModeText As String = "S"
            If "input" = ModeStr Then ModeText = "R"
            Dim LogText As String = Now.ToString("yyyy/MM/dd HH:mm:ss.fff") & "," & ModeText & "," & RemoteIP & "," & Src
            WriteLineTextFile(FilePath, LogText)

            'システムログを保存する
            Dim FileName2 As String = "sys" & Today.ToString("yyyyMMdd") & ".log"
            Dim FilePath2 As String = System.IO.Path.Combine(DirPath, FileName2)
            WriteLineTextFile(FilePath2, LogText)
        Catch ex As Exception
            Err.Raise(9999, Nothing, "ログの書込みに失敗しました。" & vbCrLf & ex.Message)
        End Try
    End Sub

#End Region

#Region "その他"

    ''' <summary>二重起動をチェックする</summary>
    ''' <remarks>二重起動時は終了する</remarks>
    Public Sub DoubleProcessCheck()
        If UBound(Diagnostics.Process.GetProcessesByName(Diagnostics.Process.GetCurrentProcess.ProcessName)) > 0 Then
            'すでに起動していると判断する
            MsgBox("多重起動はできません。")
            '終了させるコードを書いてください
            System.Windows.Forms.Application.Exit()
            End
        End If
    End Sub

    'エラー関連処理
    ''' <summary>エラー関連処理</summary>
    ''' <param name="exc">エラー情報</param>
    ''' <param name="DebugMode">デバッグモードかどうか</param>
    ''' <param name="Message">エラーメッセージ</param>
    ''' <param name="Type">緊急度(1:警告,2:エラー)</param>
    ''' <param name="DirPath">保存先フォルダパス</param>
    ''' <remarks>デバッグモード時、エラーならばメッセージボックスを表示する</remarks>
    Public Sub ErrMsgBox(exc As Exception, DebugMode As Boolean, Message As String, Optional ByVal Type As Long = 0, Optional ByVal DirPath As String = "c:\var\logs\")
        'エラー・警告等の処理
        Try
            Dim ErrMessage As String = Message & vbCrLf & ""

            If Not (exc Is Nothing) Then
                '例外コード取得
                Dim hResult As Integer = Convert.ToInt32(GetType(Exception).InvokeMember( _
                "HResult", Reflection.BindingFlags.NonPublic Or _
                System.Reflection.BindingFlags.Instance Or _
                System.Reflection.BindingFlags.Public Or Reflection.BindingFlags.GetProperty, _
                Nothing, exc, Nothing, Globalization.CultureInfo.InvariantCulture))

                ErrMessage = ""
                ErrMessage &= "エラー番号：" & hResult & vbCrLf
                ErrMessage &= "ex.メンバ名：" & exc.TargetSite.Name & vbCrLf
                ErrMessage &= "メッセージ：" & Message & vbCrLf
                ErrMessage &= "ex.メッセージ：" & exc.Message & vbCrLf
            End If

            Dim DrawMsg As String = "エラー情報" & vbCrLf & ErrMessage

            Dim TypeText As String = "?"
            Select Case Type
                Case 0 : TypeText = "N" '通常動作
                Case 1 : TypeText = "W" '警告
                Case 2 : TypeText = "E" 'エラー
            End Select

            Dim LogText As String = Now.ToString("yyyy/MM/dd HH:mm:ss.fff") & "," & TypeText & "," & ErrMessage.Replace(vbCrLf, ",")

            'ログファイルに記述する
            '保存先を決める
            Dim FileName As String = "sys" & Today.ToString("yyyyMMdd") & ".log"
            Dim FilePath As String = System.IO.Path.Combine(DirPath, FileName)
            '保存内容を決める
            WriteLineTextFile(FilePath, LogText)

            If (DebugMode And (Type = 2)) Then MsgBox(DrawMsg)
        Catch ex As Exception
        End Try
    End Sub


    ''' <summary>ソケットが接続中かどうかを返す</summary>
    ''' <remarks></remarks>
    Public Function IsConnecting(ByRef SrcSocket As Socket) As Boolean
        ''Socket.Connectedはクラスが行った直近の接続結果のみを返す。
        ''なので相手が切断した時等、厳密なコネクションの現状を返すわけではない。
        ''その対応として以下のルーチンがMSDNで推奨されている
        'Dim blockingState As Boolean = SrcSocket.Blocking
        'Try
        '    Dim tmp(0) As Byte
        '    SrcSocket.Blocking = False
        '    SrcSocket.Send(tmp, 0, 0)
        'Catch e As SocketException
        '    If Not (e.NativeErrorCode.Equals(10035)) Then Return False
        'Finally
        '    SrcSocket.Blocking = blockingState
        'End Try
        'Return SrcSocket.Connected
        ''しかしこれは検証した結果正しく動作しない事がわかった。
        ''SrcSocket.Connectedは接続は認知するが、他者からの切断は認知しないのでこれを反映する
        Try
            If Not (SrcSocket.Connected) Then Return False 'まだ接続状態ではないと言う事でFALSEを返す
            If SocketShutDowned(SrcSocket) Then Return False '接続後の切断状態だと言う事でFALSEを返す
            Return True '接続後、切断前と言う事でTrueを返す
        Catch ex As Exception
            Err.Raise(9999, Nothing, "ソケットの接続判断で例外が発生しました。" & ex.Message)
        End Try
    End Function

    ''' <summary>ソケットが切断されたかどうかを返す</summary>
    ''' <remarks>接続中かどうかではなく接続されていたコネクションが切断されたかどうかを返す関数。未接続の物は</remarks>
    Public Function SocketShutDowned(ByRef SrcSocket As Socket) As Boolean
        'MSDNにはSocketConnectedの方法が紹介されているが、これでは切断を取得できなかった。
        'なので以下の方法をとる
        Try
            Dim part1 As Boolean = SrcSocket.Poll(1000, SelectMode.SelectRead)
            Dim part2 As Boolean = (SrcSocket.Available = 0)
            If part1 And part2 Then Return Not (False)
            Return Not (True)
        Catch ex As Exception
            Err.Raise(9999, Nothing, "ソケットの切断判断で例外が発生しました。" & ex.Message)
        End Try
    End Function

    Public Function IsConnected(ByRef SrcSocket As Socket) As Boolean
        ''Socket.Connectedはクラスが行った直近の接続結果のみを返す。
        ''なので相手が切断した時等、厳密なコネクションの現状を返すわけではない。
        ''その対応として以下のルーチンがMSDNで推奨されている
        Dim blockingState As Boolean = SrcSocket.Blocking
        Try
            Dim tmp(0) As Byte
            SrcSocket.Blocking = False
            SrcSocket.Send(tmp, 0, 0)
            Return True
        Catch e As SocketException
            If e.NativeErrorCode.Equals(10035) Then Return True
            Return False
        Finally
            SrcSocket.Blocking = blockingState
        End Try
        Return SrcSocket.Connected
        ''しかしこれは検証した結果正しく動作しない事がわかった。
        ''SrcSocket.Connectedは接続は認知するが、他者からの切断は認知しないのでこれを反映する
    End Function

    Public Sub MRComObject(ByRef objCom As Object)
        'COM オブジェクトの使用後、明示的に COM オブジェクトへの参照を解放する
        Try
            '提供されたランタイム呼び出し可能ラッパーの参照カウントをデクリメントします
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objCom)
        Catch
        Finally
            '参照を解除する
            objCom = Nothing
        End Try
    End Sub


#End Region



End Module

