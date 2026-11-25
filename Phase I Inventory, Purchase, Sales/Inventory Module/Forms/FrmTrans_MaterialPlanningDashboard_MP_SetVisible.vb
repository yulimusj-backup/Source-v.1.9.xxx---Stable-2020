Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_MP_SetVisible
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
                If Me.chkProcurementStrategy.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanning, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanning_ShowProcurementStrategy, 1, Me.vUserName)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanning, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanning_ShowProcurementStrategy, 0, Me.vUserName)
                End If
                If Me.chkOrderStrategy.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanning, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanning_ProcurementType, 1, Me.vUserName)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanning, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanning_ProcurementType, 0, Me.vUserName)
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
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanning, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanning_ShowProcurementStrategy, _
            vUserName)
            If Dummy = "" Then
                Me.chkProcurementStrategy.Checked = False
            Else
                If Dummy = 1 Then
                    Me.chkProcurementStrategy.Checked = True
                Else
                    Me.chkProcurementStrategy.Checked = False
                End If
            End If

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanning, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanning_ProcurementType, _
            vUserName)
            If Dummy = "" Then
                Me.chkOrderStrategy.Checked = False
            Else
                If Dummy = 1 Then
                    Me.chkOrderStrategy.Checked = True
                Else
                    Me.chkOrderStrategy.Checked = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class