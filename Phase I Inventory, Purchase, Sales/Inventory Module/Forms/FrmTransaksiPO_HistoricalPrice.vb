Option Explicit On
Option Strict Off

Public Class FrmTransaksiPO_HistoricalPrice

    Dim objCD As New Core_Module.ClsCekData

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub FrmTransaksiPO_HistoricalPrice_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.dgvPrice.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
        Me.ShowIcon = False
    End Sub

    Private Sub dgvPrice_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvPrice.Sorted
        Try
            objCD.SetNumber(Me.dgvPrice)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class