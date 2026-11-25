Public Class FrmTrans_ProductionSchedule_NewSch_Setting
    Public UserCode As String = ""
    Dim ObjSetting As New Core_Module.ClsTransaksiSettingUser

    Private Sub FrmTrans_ProductionSchedule_NewSch_Setting_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ShowJO, ShowMO, ShowCodeNo, ShowDesc, ShowQtyTarget As String

        Try
            ShowJO = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch_ShowJO, Me.UserCode)
            ShowMO = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch_ShowMO, Me.UserCode)
            ShowCodeNo = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch_ShowCodeNo, Me.UserCode)
            ShowDesc = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch_ShowDesc, Me.UserCode)
            ShowQtyTarget = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch_ShowQtyTarget, Me.UserCode)

            If ShowJO = "1" Then
                Me.chkJONUmber.Checked = True
            Else
                Me.chkJONUmber.Checked = False
            End If
            If ShowMO = "1" Then
                Me.chkMONumber.Checked = True
            Else
                Me.chkMONumber.Checked = False
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
            If ShowQtyTarget = "1" Then
                Me.ChkQtyTarget.Checked = True
            Else
                Me.ChkQtyTarget.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            'code no atau desc harus di cek
            If Me.ChkCodeNo.Checked = False And Me.ChkDesc.Checked = False Then
                MsgBox("Please check at least code no or description.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            'mo number atau jo number harus di cek
            If Me.chkJONUmber.Checked = False And Me.chkMONumber.Checked = False Then
                MsgBox("Please check at least MO Number or JO Number.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                If Me.ChkCodeNo.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch_ShowCodeNo, "1", Me.UserCode)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch_ShowCodeNo, "0", Me.UserCode)
                End If
                If Me.ChkDesc.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch_ShowDesc, "1", Me.UserCode)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch_ShowDesc, "0", Me.UserCode)
                End If
                If Me.chkJONUmber.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch_ShowJO, "1", Me.UserCode)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch_ShowJO, "0", Me.UserCode)
                End If
                If Me.chkMONumber.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch_ShowMO, "1", Me.UserCode)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch_ShowMO, "0", Me.UserCode)
                End If
                If Me.ChkQtyTarget.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch_ShowQtyTarget, "1", Me.UserCode)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch_ShowQtyTarget, "0", Me.UserCode)
                End If
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class