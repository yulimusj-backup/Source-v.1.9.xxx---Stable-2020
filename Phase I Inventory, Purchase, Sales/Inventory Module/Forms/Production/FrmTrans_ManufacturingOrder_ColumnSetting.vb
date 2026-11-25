Option Explicit On
Option Strict Off

Public Class FrmTrans_ManufacturingOrder_ColumnSetting

    Public UserCode As String = ""
    Dim ObjSetting As New Core_Module.ClsTransaksiSettingUser

    Private Sub FrmTrans_ProductionSchedule_NewSch2_ColumnSetting_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ShowCodeNo, ShowDesc, ShowProcessCode, ShowProcessDesc As String
        Try
            ShowCodeNo = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowCodeNo, Me.UserCode)
            ShowDesc = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowDescription, Me.UserCode)
            ShowProcessCode = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowProcessCode, Me.UserCode)
            ShowProcessDesc = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowProcessDesc, Me.UserCode)

            If ShowCodeNo = "1" Then
                Me.ChkCodeNo.Checked = True
            Else
                Me.ChkCodeNo.Checked = False
            End If
            If ShowDesc = "1" Then
                Me.ChkDescription.Checked = True
            Else
                Me.ChkDescription.Checked = False
            End If
            If ShowProcessCode = "1" Then
                Me.ChkProcessCode.Checked = True
            Else
                Me.ChkProcessCode.Checked = False
            End If
            If ShowProcessDesc = "1" Then
                Me.ChkProcessDesc.Checked = True
            Else
                Me.ChkProcessDesc.Checked = False
            End If 
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                
                If Me.ChkCodeNo.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowCodeNo, "1", Me.UserCode)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowCodeNo, "0", Me.UserCode)
                End If
                If Me.ChkDescription.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowDescription, "1", Me.UserCode)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowDescription, "0", Me.UserCode)
                End If
                If Me.ChkProcessCode.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowProcessCode, "1", Me.UserCode)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowProcessCode, "0", Me.UserCode)
                End If
                If Me.ChkProcessDesc.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowProcessDesc, "1", Me.UserCode)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowProcessDesc, "0", Me.UserCode)
                End If

                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class