''rosadin-r16
Option Explicit On
Option Strict Off
Public Class FrmTrans_ProductionSchedule_SplitMerge_ColSetting
    Public UserCode As String = String.Empty
    Dim objUserSetting As New Core_Module.ClsTransaksiSettingUser

    Private Sub FrmTrans_ProductionSchedule_ChangeWU_ColSetting_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim MO, JO, WO, Process, Code, Desc, Rate, Qty, Unit, WH, DayStart, DayEnd, Progress As String
        Try
            Me.DialogResult = Windows.Forms.DialogResult.None
            If Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowMO, Me.UserCode) <> "0" Then
                MO = "1"
            Else
                MO = "0"
            End If
            If Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowJO, Me.UserCode) <> "0" Then
                JO = "1"
            Else
                JO = "0"
            End If
            If Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowWO, Me.UserCode) <> "0" Then
                WO = "1"
            Else
                WO = "0"
            End If
            If Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowProcess, Me.UserCode) <> "0" Then
                Process = "1"
            Else
                Process = "0"
            End If
            If Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowCodeNo, Me.UserCode) <> "0" Then
                Code = "1"
            Else
                Code = "0"
            End If
            If Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowDesc, Me.UserCode) <> "0" Then
                Desc = "1"
            Else
                Desc = "0"
            End If
            If Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowRate, Me.UserCode) <> "0" Then
                Rate = "1"
            Else
                Rate = "0"
            End If
            If Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowQty, Me.UserCode) <> "0" Then
                Qty = "1"
            Else
                Qty = "0"
            End If
            If Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowUnit, Me.UserCode) <> "0" Then
                Unit = "1"
            Else
                Unit = "0"
            End If
            If Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowWH, Me.UserCode) <> "0" Then
                WH = "1"
            Else
                WH = "0"
            End If
            If Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowStartDay, Me.UserCode) <> "0" Then
                DayStart = "1"
            Else
                DayStart = "0"
            End If
            If Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowEndDay, Me.UserCode) <> "0" Then
                DayEnd = "1"
            Else
                DayEnd = "0"
            End If
            If Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowProgress, Me.UserCode) <> "0" Then
                Progress = "1"
            Else
                Progress = "0"
            End If
            If MO = "1" Then
                Me.ChkMO.Checked = True
            Else
                Me.ChkMO.Checked = False
            End If
            If JO = "1" Then
                Me.ChkJO.Checked = True
            Else
                Me.ChkJO.Checked = False
            End If
            If WO = "1" Then
                Me.ChkWO.Checked = True
            Else
                Me.ChkWO.Checked = False
            End If
            If Process = "1" Then
                Me.ChkProcess.Checked = True
            Else
                Me.ChkProcess.Checked = False
            End If
            If Code = "1" Then
                Me.ChkCodeNo.Checked = True
            Else
                Me.ChkCodeNo.Checked = False
            End If
            If Desc = "1" Then
                Me.ChkDesc.Checked = True
            Else
                Me.ChkDesc.Checked = False
            End If
            If Rate = "1" Then
                Me.ChkRate.Checked = True
            Else
                Me.ChkRate.Checked = False
            End If
            If Qty = "1" Then
                Me.ChkQty.Checked = True
            Else
                Me.ChkQty.Checked = False
            End If
            If Unit = "1" Then
                Me.ChkUnit.Checked = True
            Else
                Me.ChkUnit.Checked = False
            End If
            If WH = "1" Then
                Me.ChkWorkHour.Checked = True
            Else
                Me.ChkWorkHour.Checked = False
            End If
            If DayStart = "1" Then
                Me.ChkDayStart.Checked = True
            Else
                Me.ChkDayStart.Checked = False
            End If
            If DayEnd = "1" Then
                Me.ChkDayFinish.Checked = True
            Else
                Me.ChkDayFinish.Checked = False
            End If
            If Progress = "1" Then
                Me.ChkProgress.Checked = True
            Else
                Me.ChkProgress.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        Try
            Me.DialogResult = Windows.Forms.DialogResult.No
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        Dim MO, JO, WO, Process, Code, Desc, Rate, Qty, Unit, WH, DayStart, DayEnd, Progress As String
        Try
            If MsgBox("Save columns setting?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                If Me.ChkMO.Checked = True Then
                    MO = "1"
                Else
                    MO = "0"
                End If
                If Me.ChkJO.Checked = True Then
                    JO = "1"
                Else
                    JO = "0"
                End If
                If Me.ChkWO.Checked = True Then
                    WO = "1"
                Else
                    WO = "0"
                End If
                If Me.ChkProcess.Checked = True Then
                    Process = "1"
                Else
                    Process = "0"
                End If
                If Me.ChkCodeNo.Checked = True Then
                    Code = "1"
                Else
                    Code = "0"
                End If
                If Me.ChkDesc.Checked = True Then
                    Desc = "1"
                Else
                    Desc = "0"
                End If
                If Me.ChkRate.Checked = True Then
                    Rate = "1"
                Else
                    Rate = "0"
                End If
                If Me.ChkQty.Checked = True Then
                    Qty = "1"
                Else
                    Qty = "0"
                End If
                If Me.ChkUnit.Checked = True Then
                    Unit = "1"
                Else
                    Unit = "0"
                End If
                If Me.ChkWorkHour.Checked = True Then
                    WH = "1"
                Else
                    WH = "0"
                End If
                If Me.ChkDayStart.Checked = True Then
                    DayStart = "1"
                Else
                    DayStart = "0"
                End If
                If Me.ChkDayFinish.Checked = True Then
                    DayEnd = "1"
                Else
                    DayEnd = "0"
                End If
                If Me.ChkProgress.Checked = True Then
                    Progress = "1"
                Else
                    Progress = "0"
                End If
                Me.objUserSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowMO, MO, Me.UserCode)
                Me.objUserSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowJO, JO, Me.UserCode)
                Me.objUserSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowWO, WO, Me.UserCode)
                Me.objUserSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowProcess, Process, Me.UserCode)
                Me.objUserSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowCodeNo, Code, Me.UserCode)
                Me.objUserSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowDesc, Desc, Me.UserCode)
                Me.objUserSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowRate, Rate, Me.UserCode)
                Me.objUserSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowQty, Qty, Me.UserCode)
                Me.objUserSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowUnit, Unit, Me.UserCode)
                Me.objUserSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowWH, WH, Me.UserCode)
                Me.objUserSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowStartDay, DayStart, Me.UserCode)
                Me.objUserSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowEndDay, DayEnd, Me.UserCode)
                Me.objUserSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowProgress, Progress, Me.UserCode)
                Me.DialogResult = Windows.Forms.DialogResult.OK
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class