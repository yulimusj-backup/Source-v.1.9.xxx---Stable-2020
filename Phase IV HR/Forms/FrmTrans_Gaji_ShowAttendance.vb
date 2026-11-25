Public Class FrmTrans_Gaji_ShowAttendance

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmTrans_Gaji_ShowAttendance_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            'Pakai coding Isi DGV Item di ClsTransaksiAttendance_Item, berdasarkan Tanggal Dari sampai
            Me.ShowIcon = False
            Me.ColNo.Frozen = True
            Me.ColDate.Frozen = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class