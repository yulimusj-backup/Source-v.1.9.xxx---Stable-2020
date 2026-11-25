Public Class FrmTrans_MaterialPlanningDashboard_SF_PrintDetail
    Dim ObjSettings As New Core_Module.ClsTransaksiSettingUser
    Public UserName As String = ""

    Private Sub FrmTrans_MaterialPlanningDashboard_SF_PrintDetail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Dummy As String = ""

        Try
            Dummy = Me.ObjSettings.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ShippingForecast, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ShippingForecast_PrintOrderNo, Me.UserName)
            If Dummy = "" Then
                Me.cboOrderNo.SelectedIndex = 0
            Else
                Me.cboOrderNo.SelectedIndex = CInt(Dummy)
            End If

            Dummy = Me.ObjSettings.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ShippingForecast, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ShippingForecast_PrintGroupBy, Me.UserName)
            If Dummy = "" Then
                Me.cboGroupBy.SelectedIndex = 0
            Else
                Me.cboGroupBy.SelectedIndex = CInt(Dummy)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.ObjSettings.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ShippingForecast, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ShippingForecast_PrintGroupBy, Me.cboGroupBy.SelectedIndex.ToString, Me.UserName)
            Me.ObjSettings.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ShippingForecast, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ShippingForecast_PrintOrderNo, Me.cboOrderNo.SelectedIndex.ToString, Me.UserName)
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class