Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_SF_SetVisible
    Dim LangsungTutup As Boolean = False
    Dim ObjSetting As New Core_Module.ClsTransaksiSettingUser

    Public vUserName As String = ""

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim dummy As MsgBoxResult

        Try
            dummy = MsgBox("Save changes?", MsgBoxStyle.YesNoCancel, "Information")
            If dummy = MsgBoxResult.Yes Then
                'update nilai setting
                If Me.chkCustomerPONumber.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ShippingForecast, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ShippingForecast_ShowCustPO, 1, Me.vUserName)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ShippingForecast, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ShippingForecast_ShowCustPO, 0, Me.vUserName)
                End If
                If Me.chkSONumber.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ShippingForecast, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ShippingForecast_ShowSONumber, 1, Me.vUserName)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ShippingForecast, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ShippingForecast_ShowSONumber, 0, Me.vUserName)
                End If
                If Me.chkProductDesc.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ShippingForecast, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ShippingForecast_ShowProductDesc, 1, Me.vUserName)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ShippingForecast, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ShippingForecast_ShowProductDesc, 0, Me.vUserName)
                End If

                LangsungTutup = True
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            ElseIf dummy = MsgBoxResult.No Then
                LangsungTutup = True
                Me.DialogResult = Windows.Forms.DialogResult.No
                Me.Close()
            ElseIf dummy = MsgBoxResult.Cancel Then
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_SF_SetVisible_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim dummy As MsgBoxResult

        Try
            If Me.LangsungTutup = False Then
                dummy = MsgBox("Save changes?", MsgBoxStyle.YesNoCancel, "Information")
                If dummy = MsgBoxResult.Yes Then
                    LangsungTutup = True
                    Me.DialogResult = Windows.Forms.DialogResult.OK
                    Me.Close()
                ElseIf dummy = MsgBoxResult.No Then
                    LangsungTutup = True
                    Me.DialogResult = Windows.Forms.DialogResult.No
                    Me.Close()
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_SF_SetVisible_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Dummy As String = ""

        Try
            Me.chkSONumber.Text = "Document Number"
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ShippingForecast, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ShippingForecast_ShowSONumber, _
            vUserName)
            If Dummy = "" Then
                Me.chkSONumber.Checked = False
            Else
                If Dummy = 1 Then
                    Me.chkSONumber.Checked = True
                Else
                    Me.chkSONumber.Checked = False
                End If
            End If

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ShippingForecast, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ShippingForecast_ShowCustPO, _
            vUserName)
            If Dummy = "" Then
                Me.chkCustomerPONumber.Checked = False
            Else
                If Dummy = 1 Then
                    Me.chkCustomerPONumber.Checked = True
                Else
                    Me.chkCustomerPONumber.Checked = False
                End If
            End If

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ShippingForecast, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ShippingForecast_ShowProductDesc, _
            vUserName)
            If Dummy = "" Then
                Me.chkProductDesc.Checked = False
            Else
                If Dummy = 1 Then
                    Me.chkProductDesc.Checked = True
                Else
                    Me.chkProductDesc.Checked = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class