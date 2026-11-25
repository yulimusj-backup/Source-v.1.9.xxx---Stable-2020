Option Explicit On
Option Strict Off

Public Class FrmTransaksiSO_SetVisible
    Dim LangsungTutup As Boolean = False
    Dim ObjSetting As New Core_Module.ClsTransaksiSettingUser

    Public vUserName As String = ""
    Public IsMultiUnit As Boolean = False

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Err")
        End Try
    End Sub

    Private Sub FrmTransaksiSO_SetVisible_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub FrmTransaksiSO_SetVisible_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Dummy As String = ""
        Try
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.SalesOrder, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.SalesOrder_ShowUnitsDefault, vUserName)
            If Dummy = "" Then
                Me.chkUnitsDefault.Checked = False
            Else
                If Dummy = 1 Then
                    Me.chkUnitsDefault.Checked = True
                Else
                    Me.chkUnitsDefault.Checked = False
                End If
            End If

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.SalesOrder, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.SalesOrder_ShowTolerance, vUserName)
            If Dummy = "" Then
                Me.chkTolerance.Checked = False
            Else
                If Dummy = 1 Then
                    Me.chkTolerance.Checked = True
                Else
                    Me.chkTolerance.Checked = False
                End If
            End If

            If Me.IsMultiUnit = False Then
                Me.chkUnitsDefault.Checked = False
                Me.chkUnitsDefault.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim dummy As MsgBoxResult
        Try
            dummy = MsgBox("Save changes?", MsgBoxStyle.YesNoCancel, "Information")
            If dummy = MsgBoxResult.Yes Then
                If Me.chkUnitsDefault.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.SalesOrder, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.SalesOrder_ShowUnitsDefault, 1, Me.vUserName)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.SalesOrder, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.SalesOrder_ShowUnitsDefault, 0, Me.vUserName)
                End If

                If Me.chkTolerance.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.SalesOrder, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.SalesOrder_ShowTolerance, 1, Me.vUserName)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.PurchaseOrder, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.SalesOrder_ShowTolerance, 0, Me.vUserName)
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
End Class