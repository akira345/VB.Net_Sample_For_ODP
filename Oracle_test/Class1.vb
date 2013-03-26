'ODPを参照するために必要なクラスを読み込む
Imports Oracle.DataAccess.Client
Imports Oracle.DataAccess.Types



Public Class Oracle_Function
    'Oracle10gR2へODPを用いてＤＢを操作するクラス
    Dim conn As New OracleConnection    '接続子

    Public Function Db_Connect(ByVal DSN As String, ByVal user_name As String, ByVal passwd As String)
        'Oracleに接続する。
        '引数：DSN：データソースネーム,user_name: ユーザー,passwd:パスワード

        '戻り値の設定
        Db_Connect = False


        '接続文字列の構築
        conn.ConnectionString = "User ID = " & user_name & ";Password = " & passwd & ";Data Source = " & DSN
        ' Debug.Print(conn.ConnectionString)

        '接続
        Try
            conn.Open()
        Catch
            '何かしらのエラー発生時処理
            MsgBox("Oracle接続に失敗")
        End Try
        MsgBox("成功！！!")

        '閉じる
        'conn.Close()

        'リターン
        Db_Connect = conn

    End Function
    Public Function Db_Close()
        'Oracleから切り離す
        conn.Close()
        conn.Dispose()

        Db_Close = True
    End Function
    Public Function Run_Stoado()
        'ストアドプロシージャ実行
        Dim cmd As New OracleCommand
        cmd.Connection = conn
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = "testproc"    '実行するプロシージャ名
        'プロシージャのパラメタをセットする
        cmd.Parameters.Add(New OracleParameter("in_val", OracleDbType.Int32, ParameterDirection.Input))
        'パラメタを入れる
        cmd.Parameters("in_val").Value = 987

        'ストアド実行
        cmd.ExecuteNonQuery()

        '終了
        cmd.Dispose()
        Run_Stoado = True

    End Function
    Public Function exec_update_sql(ByVal in_sql As String, ByVal trn_sw As Integer)

        Dim cmd As New OracleCommand
        Dim rows As Decimal
        Dim trn As OracleTransaction


        cmd.Connection = conn

        cmd.CommandText = in_sql
        'トランザクション処理かどうか
        If trn_sw = True Then
            'トランザクション処理開始

            trn = conn.BeginTransaction
            Try
                rows = cmd.ExecuteNonQuery
                If (rows > 0) Then
                    trn.Commit()
                    exec_update_sql = rows
                End If
            Catch ex As Exception
                trn.Rollback()
                'trn.Commit()
                exec_update_sql = False
            End Try
        Else
            '直更新
            rows = cmd.ExecuteNonQuery
            exec_update_sql = rows
        End If


        cmd.Dispose()

    End Function
    Public Function exec_sql(ByVal in_sql As String) As OracleDataReader
        Dim cmd As New OracleCommand
        Dim in_rs As OracleDataReader
        cmd.Connection = conn
        cmd.CommandText = in_sql
        in_rs = cmd.ExecuteReader
        exec_sql = in_rs


    End Function
End Class
