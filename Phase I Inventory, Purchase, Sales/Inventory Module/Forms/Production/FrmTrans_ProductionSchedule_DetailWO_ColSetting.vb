Option Explicit On
Option Strict Off
Public Class FrmTrans_ProductionSchedule_DetailWO_ColSetting

    Dim ObjSetting As New Core_Module.ClsTransaksiSettingUser
    Dim ShowCodeNo, ShowDesc, ShowProcess, ShowProcessDesc As String
    Public UserCode As String

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If MsgBox("Are you sure to cancel and discard data changes?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        End If
    End Sub

    Private Sub FrmTrans_ProductionSchedule_DetailWO_ColSetting_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.CancelButton = Me.Button2
            Me.AcceptButton = Me.Button1
            Me.DialogResult = Windows.Forms.DialogResult.None
            ShowCodeNo = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_DetailWO, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_DetailWO_ShowCodeNo, Me.UserCode)
            ShowDesc = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_DetailWO, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_DetailWO_ShowDesc, Me.UserCode)
            ShowProcess = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_DetailWO, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_DetailWO_ShowProcess, Me.UserCode)
            ShowProcessDesc = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_DetailWO, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_DetailWO_ShowProcessDesc, Me.UserCode)
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
            If ShowProcess = "1" Or ShowProcess = "" Then
                Me.ChkProcess.Checked = True
            Else
                Me.ChkProcess.Checked = False
            End If
            If ShowProcessDesc = "1" Or ShowProcessDesc = "" Then
                Me.ChkProcessDesc.Checked = True
            Else
                Me.ChkProcessDesc.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If MsgBox("Are you sure to save columns settings?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
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
                If Me.ChkProcess.Checked Then
                    ShowProcess = "1"
                Else
                    ShowProcess = "0"
                End If
                If Me.ChkProcessDesc.Checked Then
                    ShowProcessDesc = "1"
                Else
                    ShowProcessDesc = "0"
                End If
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_DetailWO, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_DetailWO_ShowCodeNo, ShowCodeNo, Me.UserCode)
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_DetailWO, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_DetailWO_ShowDesc, ShowDesc, Me.UserCode)
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_DetailWO, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_DetailWO_ShowProcess, ShowProcess, Me.UserCode)
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_DetailWO, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_DetailWO_ShowProcessDesc, ShowProcessDesc, Me.UserCode)
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class