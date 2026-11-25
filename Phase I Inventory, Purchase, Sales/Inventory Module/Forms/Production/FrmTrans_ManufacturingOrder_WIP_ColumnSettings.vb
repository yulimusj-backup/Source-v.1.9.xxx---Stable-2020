Option Explicit On
Option Strict Off
Public Class FrmTrans_ManufacturingOrder_WIP_ColumnSettings

    Public UserCode As String = String.Empty
    Dim objUserSetting As New Core_Module.ClsTransaksiSettingUser
    Dim ShowWO, ShowProcess, ShowCodeNo, ShowDesc As String

    Private Sub FrmTrans_ManufacturingOrder_WIP_ColumnSettings_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.DialogResult = Windows.Forms.DialogResult.None
            ShowWO = Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_WIP_ShowWO, Me.UserCode)
            ShowProcess = Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_WIP_ShowProcessCode, Me.UserCode)
            ShowCodeNo = Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_WIP_ShowCodeNo, Me.UserCode)
            ShowDesc = Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_WIP_ShowDesc, Me.UserCode)
            If ShowWO = "1" Or ShowWO = "" Then
                Me.ChkWO.Checked = True
            Else
                Me.ChkWO.Checked = False
            End If
            If ShowProcess = "1" Or ShowProcess = "" Then
                Me.ChkProcessCode.Checked = True
            Else
                Me.ChkProcessCode.Checked = False
            End If
            If ShowCodeNo = "1" Or ShowCodeNo = "" Then
                Me.ChkCodeNo.Checked = True
            Else
                Me.ChkCodeNo.Checked = False
            End If
            If ShowDesc = "1" Or ShowDesc = "" Then
                Me.ChkDesc.Checked = True
            Else
                Me.ChkDesc.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        Try
            If MsgBox("Save columns setting?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                If Me.ChkWO.Checked Then
                    ShowWO = "1"
                Else
                    ShowWO = "0"
                End If
                If Me.ChkProcessCode.Checked Then
                    ShowProcess = "1"
                Else
                    ShowProcess = "0"
                End If
                If Me.ChkCodeNo.Checked Then
                    ShowCodeNo = "1"
                Else
                    ShowCodeNo = "0"
                End If
                If Me.ChkDesc.Checked Then
                    ShowDesc = "1"
                Else
                    ShowDesc = "0"
                End If
                Me.objUserSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_WIP_ShowWO, ShowWO, Me.UserCode)
                Me.objUserSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_WIP_ShowProcessCode, ShowProcess, Me.UserCode)
                Me.objUserSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_WIP_ShowCodeNo, ShowCodeNo, Me.UserCode)
                Me.objUserSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_WIP_ShowDesc, ShowDesc, Me.UserCode)
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        Try
            If MsgBox("Are you sure to cancel and discard data changes?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class