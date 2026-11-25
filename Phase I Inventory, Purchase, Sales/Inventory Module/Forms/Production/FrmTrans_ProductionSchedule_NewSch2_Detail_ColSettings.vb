Option Explicit On
Option Strict Off

Public Class FrmTrans_ProductionSchedule_NewSch2_Detail_ColSettings

    Public UserCode As String = ""
    Dim ObjSetting As New Core_Module.ClsTransaksiSettingUser

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmTrans_ProductionSchedule_NewSch2_Detail_ColSettings_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ShowJO, ShowProcess, ShowCodeNo, ShowDesc As String
        Try
            ShowJO = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2_Detail, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_Detail_ShowJO, Me.UserCode)
            ShowProcess = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2_Detail, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_Detail_ShowProcess, Me.UserCode)
            ShowCodeNo = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2_Detail, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_Detail_ShowCodeNo, Me.UserCode)
            ShowDesc = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2_Detail, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_Detail_ShowDesc, Me.UserCode)

            If ShowJO = "1" Then
                Me.ChkJONumber.Checked = True
            Else
                Me.ChkJONumber.Checked = False
            End If
            If ShowProcess = "1" Then
                Me.ChkProcessCode.Checked = True
            Else
                Me.ChkProcessCode.Checked = False
            End If
            If ShowCodeNo = "1" Then
                Me.ChkCodeNo.Checked = True
            Else
                Me.ChkCodeNo.Checked = False
            End If
            If ShowDesc = "1" Then
                Me.ChkDesc.Checked = True
            Else
                Me.ChkDesc.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If Me.ChkCodeNo.Checked = False And Me.ChkDesc.Checked = False Then
                MsgBox("Please select one colom Code No or Description", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                If Me.ChkJONumber.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2_Detail, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_Detail_ShowJO, "1", Me.UserCode)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2_Detail, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_Detail_ShowJO, "0", Me.UserCode)
                End If
                If Me.ChkProcessCode.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2_Detail, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_Detail_ShowProcess, "1", Me.UserCode)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2_Detail, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_Detail_ShowProcess, "0", Me.UserCode)
                End If
                If Me.ChkCodeNo.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2_Detail, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_Detail_ShowCodeNo, "1", Me.UserCode)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2_Detail, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_Detail_ShowCodeNo, "0", Me.UserCode)
                End If
                If Me.ChkDesc.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2_Detail, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_Detail_ShowDesc, "1", Me.UserCode)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2_Detail, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_Detail_ShowDesc, "0", Me.UserCode)
                End If

                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class