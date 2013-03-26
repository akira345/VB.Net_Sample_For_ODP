Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim a As New Oracle_Function
        a.Db_Connect("localhost/testdb", "scott", "tiger") 'DB接続子
        'ストアドを実行
        a.Run_Stoado()
        'Selectを実行
        Dim rs = a.exec_sql("select * from testtable order by id")
        While rs.read
            Debug.Print(rs("ID").ToString)
            'Insertを実行
            If (a.exec_update_sql("insert into testtable values(1,'test')", True) <> True) Then
                MsgBox("エラー発生")
                Exit While
            End If
        End While

        a.Db_Close()

    End Sub
End Class
