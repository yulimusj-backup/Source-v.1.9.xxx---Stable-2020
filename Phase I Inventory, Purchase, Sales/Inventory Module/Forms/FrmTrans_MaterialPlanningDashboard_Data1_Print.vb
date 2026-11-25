Public Class FrmTrans_MaterialPlanningDashboard_Data1_Print
    Public vPrimary As Double = -1
    Public Obj As ClsTransaksiMaterialPlanningGroup

    Dim DS As New DseTransaksiMaterialPlanningGroup
    Dim rpt As New CrpTransaksiMaterialPlanningGroup_Cetak

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_Data1_Print_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.Obj.Cetak(vPrimary, Me.CrvLaporan, Me.rpt, Me.DS)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class