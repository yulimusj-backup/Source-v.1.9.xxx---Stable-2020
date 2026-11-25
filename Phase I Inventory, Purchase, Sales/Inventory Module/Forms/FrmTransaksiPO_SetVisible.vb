Option Explicit On
Option Strict Off

Public Class FrmTransaksiPO_SetVisible
    Dim LangsungTutup As Boolean = False
    Dim ObjSetting As New Core_Module.ClsTransaksiSettingUser

    Public vUserName As String = ""
    Public IsMultiUnit As Boolean = False

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
                If Me.chkQtyDefault.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.PurchaseOrder, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.PurchaseOrder_ShowQtyDefault, 1, Me.vUserName)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.PurchaseOrder, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.PurchaseOrder_ShowQtyDefault, 0, Me.vUserName)
                End If

                If Me.chkUnitCodeDefault.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.PurchaseOrder, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.PurchaseOrder_ShowUnitCodeDefault, 1, Me.vUserName)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.PurchaseOrder, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.PurchaseOrder_ShowUnitCodeDefault, 0, Me.vUserName)
                End If

                If Me.chkToleranceDown.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.PurchaseOrder, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.PurchaseOrder_ShowToleranceDown, 1, Me.vUserName)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.PurchaseOrder, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.PurchaseOrder_ShowToleranceDown, 0, Me.vUserName)
                End If

                If Me.chkToleranceUp.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.PurchaseOrder, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.PurchaseOrder_ShowToleranceUp, 1, Me.vUserName)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.PurchaseOrder, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.PurchaseOrder_ShowToleranceUp, 0, Me.vUserName)
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

    Private Sub FrmTransaksiPO_SetVisible_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub FrmTransaksiPO_SetVisible_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Dummy As String = ""

        Try
            Me.ShowIcon = False
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.PurchaseOrder, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.PurchaseOrder_ShowQtyDefault, vUserName)
            If Dummy = "" Then
                Me.chkQtyDefault.Checked = False
            Else
                If Dummy = 1 Then
                    Me.chkQtyDefault.Checked = True
                Else
                    Me.chkQtyDefault.Checked = False
                End If
            End If

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.PurchaseOrder, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.PurchaseOrder_ShowUnitCodeDefault, vUserName)
            If Dummy = "" Then
                Me.chkUnitCodeDefault.Checked = False
            Else
                If Dummy = 1 Then
                    Me.chkUnitCodeDefault.Checked = True
                Else
                    Me.chkUnitCodeDefault.Checked = False
                End If
            End If

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.PurchaseOrder, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.PurchaseOrder_ShowToleranceDown, vUserName)
            If Dummy = "" Then
                Me.chkToleranceDown.Checked = False
            Else
                If Dummy = 1 Then
                    Me.chkToleranceDown.Checked = True
                Else
                    Me.chkToleranceDown.Checked = False
                End If
            End If

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.PurchaseOrder, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.PurchaseOrder_ShowToleranceUp, vUserName)
            If Dummy = "" Then
                Me.chkToleranceUp.Checked = False
            Else
                If Dummy = 1 Then
                    Me.chkToleranceUp.Checked = True
                Else
                    Me.chkToleranceUp.Checked = False
                End If
            End If

            If Me.IsMultiUnit = False Then
                Me.chkQtyDefault.Checked = False
                Me.chkQtyDefault.Enabled = False
                Me.chkUnitCodeDefault.Checked = False
                Me.chkUnitCodeDefault.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class