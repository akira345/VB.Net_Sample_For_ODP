Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim a As New Oracle_Function
        a.Db_Connect("localhost/testdb", "scott", "tiger") 'DB�ڑ��q
        '�X�g�A�h�����s
        a.Run_Stoado()
        'Select�����s
        Dim rs = a.exec_sql("select * from testtable order by id")
        While rs.read
            Debug.Print(rs("ID").ToString)
            'Insert�����s
            If (a.exec_update_sql("insert into testtable values(1,'test')", True) <> True) Then
                MsgBox("�G���[����")
                Exit While
            End If
        End While

        a.Db_Close()

    End Sub
End Class
