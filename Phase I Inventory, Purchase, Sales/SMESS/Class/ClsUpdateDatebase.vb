Public Class ClsUpdateDatebase
    Public Function UpdateDatabase(ByVal vOldVersion As String) As Boolean
        Try
            If vOldVersion = "1.6.01" Then
                Me.UpdatingTo_v_1_6_02()
                Me.UpdatingTo_v_1_6_03()

                Return True
            ElseIf vOldVersion = "1.6.02" Then
                Me.UpdatingTo_v_1_6_03()

                Return True
            Else
                MsgBox("Your current database version is not up to date with the software requirement. To guarantee maximum software performance please contact your vendor to settle this problem.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub UpdatingTo_v_1_6_03()
        Dim ObjDB As New Core_Module.ClsDBVersion

        Try
            'antrian, jangan sampai ada 2 user yang melakukan proses ini sekaligus

            'back up database

            'lakukan semua perubahan database secara otomatis

            'update versi di dalam database
            ObjDB.SetDBVersion("1.6.03")
            'lepaskan antrian

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjDB = Nothing
        End Try
    End Sub

    Private Sub UpdatingTo_v_1_6_02()
        Dim ObjDB As New Core_Module.ClsDBVersion

        Try
            'antrian, jangan sampai ada 2 user yang melakukan proses ini sekaligus

            'back up database

            'lakukan semua perubahan database secara otomatis

            'update versi di dalam database
            ObjDB.SetDBVersion("1.6.02")
            'lepaskan antrian

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjDB = Nothing
        End Try
    End Sub
End Class
