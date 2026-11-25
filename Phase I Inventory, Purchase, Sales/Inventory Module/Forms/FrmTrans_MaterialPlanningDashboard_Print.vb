Public Class FrmTrans_MaterialPlanningDashboard_Print
    Public vPrimaryTransaksi As Double = -1
    Public ObjMP As New ClsTransaksiMaterialPlanning

    Dim DS As New DseTransaksiMaterialPlanning2
    Dim vReport As New CrpTrans_MaterialPlanning
    Dim vReport2 As New CrpTrans_MaterialPlanning2

    Public Enum EnumGroupingCetakan As Integer
        NoGrouping
        GroupByCategory
        GroupBySubBrand
        GroupByBrand
        GroupByProductGroup
    End Enum

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_Print_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim i As New FrmTrans_MaterialPlanningDashboard_PrintFilter
        Dim vRefDate As Date = #1/1/1950#
        Dim ShowCodeNo, ShowDesc As Boolean
        Dim Group2 As EnumGroupingCetakan = EnumGroupingCetakan.GroupByBrand

        Try
            If i.ShowDialog = Windows.Forms.DialogResult.OK Then
                If i.chkMP_RequiredOnly.Checked = True Then
                    vRefDate = i.dtpMP_Filter.Value
                End If
                If i.cboProductDetail.SelectedIndex = 0 Then
                    ShowCodeNo = True
                    ShowDesc = False
                ElseIf i.cboProductDetail.SelectedIndex = 1 Then
                    ShowCodeNo = False
                    ShowDesc = True
                ElseIf i.cboProductDetail.SelectedIndex = 2 Then
                    ShowCodeNo = True
                    ShowDesc = True
                End If
                If i.cboGrouping.SelectedIndex = 0 Then
                    Group2 = EnumGroupingCetakan.NoGrouping
                ElseIf i.cboGrouping.SelectedIndex = 1 Then
                    Group2 = EnumGroupingCetakan.GroupByProductGroup
                ElseIf i.cboGrouping.SelectedIndex = 2 Then
                    Group2 = EnumGroupingCetakan.GroupByBrand
                ElseIf i.cboGrouping.SelectedIndex = 3 Then
                    Group2 = EnumGroupingCetakan.GroupBySubBrand
                ElseIf i.cboGrouping.SelectedIndex = 4 Then
                    Group2 = EnumGroupingCetakan.GroupByCategory
                End If

                Me.ObjMP.Cetak(Me.vPrimaryTransaksi, i.chkRequiredOnly.Checked, vRefDate, Me.CrvLaporan, Me.DS, Me.vReport, _
                i.cboFilter_ProductGroupLevel.Text, i.TxtFilter_ProductGroupCode.Text, i.chkSOBOnly.Checked, _
                i.chkPrOBOnly.Checked, i.chkFollowUpOnly.Checked, ShowCodeNo, ShowDesc, Group2, vReport2, i.chkShowLegend.Checked, _
                i.chkFilter_Goods.Checked, i.chkFilter_Service.Checked)

                Me.CrvLaporan.Refresh()
            Else
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub
End Class