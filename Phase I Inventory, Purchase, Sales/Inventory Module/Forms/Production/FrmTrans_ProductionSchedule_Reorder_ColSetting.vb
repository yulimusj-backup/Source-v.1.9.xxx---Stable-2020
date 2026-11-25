Option Explicit On
Option Strict Off
Public Class FrmTrans_ProductionSchedule_Reorder_ColSetting
    Public UserCode As String
    Dim ObjSetting As New Core_Module.ClsTransaksiSettingUser

    Dim ShowMO, ShowJO, ShowWO, ShowProcess, ShowCodeNo, ShowDesc, ShowRate, ShowQty, ShowUnit, ShowWH, ShowDayStart, ShowDayFinish, ShowProgress As String
    Private Sub FrmTrans_ProductionSchedule_Reorder_ColSetting_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.DialogResult = Windows.Forms.DialogResult.None
            ShowMO = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowMO, Me.UserCode)
            ShowJO = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowJO, Me.UserCode)
            ShowWO = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowWO, Me.UserCode)
            ShowProcess = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowProcess, Me.UserCode)
            ShowCodeNo = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowCodeNo, Me.UserCode)
            ShowDesc = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowDesc, Me.UserCode)
            ShowRate = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowRate, Me.UserCode)
            ShowQty = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowQty, Me.UserCode)
            ShowUnit = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowUnit, Me.UserCode)
            ShowWH = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowWH, Me.UserCode)
            ShowDayStart = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowDayStart, Me.UserCode)
            ShowDayFinish = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowDayFinish, Me.UserCode)
            ShowProgress = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowProgress, Me.UserCode)
            If ShowMO = "1" Or ShowMO = "" Then
                Me.ChkMO.Checked = True
            Else
                Me.ChkMO.Checked = False
            End If
            If ShowJO = "1" Or ShowJO = "" Then
                Me.ChkJO.Checked = True
            Else
                Me.ChkJO.Checked = False
            End If
            If ShowWO = "1" Or ShowWO = "" Then
                Me.ChkWONo.Checked = True
            Else
                Me.ChkWONo.Checked = False
            End If
            If ShowProcess = "1" Or ShowProcess = "" Then
                Me.ChkProcess.Checked = True
            Else
                Me.ChkProcess.Checked = False
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
            If ShowRate = "1" Or ShowRate = "" Then
                Me.ChkRate.Checked = True
            Else
                Me.ChkRate.Checked = False
            End If
            If ShowQty = "1" Or ShowQty = "" Then
                Me.ChkQty.Checked = True
            Else
                Me.ChkQty.Checked = False
            End If
            If ShowUnit = "1" Or ShowUnit = "" Then
                Me.ChkUnit.Checked = True
            Else
                Me.ChkUnit.Checked = False
            End If
            If ShowWH = "1" Or ShowWH = "" Then
                Me.ChkWorkHour.Checked = True
            Else
                Me.ChkWorkHour.Checked = False
            End If
            If ShowDayStart = "1" Or ShowDayStart = "" Then
                Me.ChkDayStart.Checked = True
            Else
                Me.ChkDayStart.Checked = False
            End If
            If ShowDayFinish = "1" Or ShowDayFinish = "" Then
                Me.ChkDayFinish.Checked = True
            Else
                Me.ChkDayFinish.Checked = False
            End If
            If ShowProgress = "1" Or ShowProgress = "" Then
                Me.ChkProgress.Checked = True
            Else
                Me.ChkProgress.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        Try
            If MsgBox("Are you sure to save columns settings?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                If Me.ChkMO.Checked Then
                    ShowMO = "1"
                Else
                    ShowMO = "0"
                End If
                If Me.ChkJO.Checked Then
                    ShowJO = "1"
                Else
                    ShowJO = "0"
                End If
                If Me.ChkWONo.Checked Then
                    ShowWO = "1"
                Else
                    ShowWO = "0"
                End If
                If Me.ChkProcess.Checked Then
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
                If Me.ChkRate.Checked Then
                    ShowRate = "1"
                Else
                    ShowRate = "0"
                End If
                If Me.ChkQty.Checked Then
                    ShowQty = "1"
                Else
                    ShowQty = "0"
                End If
                If Me.ChkUnit.Checked Then
                    ShowUnit = "1"
                Else
                    ShowUnit = "0"
                End If
                If Me.ChkWorkHour.Checked Then
                    ShowWH = "1"
                Else
                    ShowWH = "0"
                End If
                If Me.ChkDayStart.Checked Then
                    ShowDayStart = "1"
                Else
                    ShowDayStart = "0"
                End If
                If Me.ChkDayFinish.Checked Then
                    ShowDayFinish = "1"
                Else
                    ShowDayFinish = "0"
                End If
                If Me.ChkProgress.Checked Then
                    ShowProgress = "1"
                Else
                    ShowProgress = "0"
                End If
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowMO, ShowMO, Me.UserCode)
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowJO, ShowJO, Me.UserCode)
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowWO, ShowWO, Me.UserCode)
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowProcess, ShowProcess, Me.UserCode)
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowCodeNo, ShowCodeNo, Me.UserCode)
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowDesc, ShowDesc, Me.UserCode)
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowRate, ShowRate, Me.UserCode)
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowQty, ShowQty, Me.UserCode)
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowUnit, ShowUnit, Me.UserCode)
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowWH, ShowWH, Me.UserCode)
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowDayStart, ShowDayStart, Me.UserCode)
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowDayFinish, ShowDayFinish, Me.UserCode)
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowProgress, ShowProgress, Me.UserCode)
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