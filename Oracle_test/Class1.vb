'ODP���Q�Ƃ��邽�߂ɕK�v�ȃN���X��ǂݍ���
Imports Oracle.DataAccess.Client
Imports Oracle.DataAccess.Types



Public Class Oracle_Function
    'Oracle10gR2��ODP��p���Ăc�a�𑀍삷��N���X
    Dim conn As New OracleConnection    '�ڑ��q

    Public Function Db_Connect(ByVal DSN As String, ByVal user_name As String, ByVal passwd As String)
        'Oracle�ɐڑ�����B
        '�����FDSN�F�f�[�^�\�[�X�l�[��,user_name: ���[�U�[,passwd:�p�X���[�h

        '�߂�l�̐ݒ�
        Db_Connect = False


        '�ڑ�������̍\�z
        conn.ConnectionString = "User ID = " & user_name & ";Password = " & passwd & ";Data Source = " & DSN
        ' Debug.Print(conn.ConnectionString)

        '�ڑ�
        Try
            conn.Open()
        Catch
            '��������̃G���[����������
            MsgBox("Oracle�ڑ��Ɏ��s")
        End Try
        MsgBox("�����I�I!")

        '����
        'conn.Close()

        '���^�[��
        Db_Connect = conn

    End Function
    Public Function Db_Close()
        'Oracle����؂藣��
        conn.Close()
        conn.Dispose()

        Db_Close = True
    End Function
    Public Function Run_Stoado()
        '�X�g�A�h�v���V�[�W�����s
        Dim cmd As New OracleCommand
        cmd.Connection = conn
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = "testproc"    '���s����v���V�[�W����
        '�v���V�[�W���̃p�����^���Z�b�g����
        cmd.Parameters.Add(New OracleParameter("in_val", OracleDbType.Int32, ParameterDirection.Input))
        '�p�����^������
        cmd.Parameters("in_val").Value = 987

        '�X�g�A�h���s
        cmd.ExecuteNonQuery()

        '�I��
        cmd.Dispose()
        Run_Stoado = True

    End Function
    Public Function exec_update_sql(ByVal in_sql As String, ByVal trn_sw As Integer)

        Dim cmd As New OracleCommand
        Dim rows As Decimal
        Dim trn As OracleTransaction


        cmd.Connection = conn

        cmd.CommandText = in_sql
        '�g�����U�N�V�����������ǂ���
        If trn_sw = True Then
            '�g�����U�N�V���������J�n

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
            '���X�V
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
