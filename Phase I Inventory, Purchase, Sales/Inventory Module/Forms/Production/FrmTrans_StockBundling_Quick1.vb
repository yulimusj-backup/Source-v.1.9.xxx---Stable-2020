Imports System.Windows.Forms

Public Class FrmTrans_StockBundling_Quick1
    Public UserCode As String = "" 

    Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Dim dummy As String = ""

        Try
            If MsgBox("Save settings?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                If Me.chkShowCodeNo.Checked = True Then
                    dummy = "1"
                Else
                    dummy = "0"
                End If
                Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_StockBundling_Quick, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_StockBundling_Quick_ShowCodeNo, dummy, Me.UserCode)

                If Me.chkShowDesc.Checked = True Then
                    dummy = "1"
                Else
                    dummy = "0"
                End If
                Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_StockBundling_Quick, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_StockBundling_Quick_ShowDesc, dummy, Me.UserCode)

                If Me.chkShowProcessCode.Checked = True Then
                    dummy = "1"
                Else
                    dummy = "0"
                End If
                Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_StockBundling_Quick, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_StockBundling_Quick_ShowProcessCode, dummy, Me.UserCode)

                If Me.chkShowProcessDesc.Checked = True Then
                    dummy = "1"
                Else
                    dummy = "0"
                End If
                Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_StockBundling_Quick, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_StockBundling_Quick_ShowProcessDesc, dummy, Me.UserCode)

                If Me.chkShowStart.Checked = True Then
                    dummy = "1"
                Else
                    dummy = "0"
                End If
                Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_StockBundling_Quick, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_StockBundling_Quick_ShowStartTime, dummy, Me.UserCode)

                If Me.chkShowEndTime.Checked = True Then
                    dummy = "1"
                Else
                    dummy = "0"
                End If
                Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_StockBundling_Quick, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_StockBundling_Quick_ShowEndTime, dummy, Me.UserCode)

                Me.DialogResult = System.Windows.Forms.DialogResult.OK
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub FrmTrans_StockBundling_Quick1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Dummy As String = ""

        Try
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_StockBundling_Quick, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_StockBundling_Quick_ShowCodeNo, Me.UserCode)
            If Dummy = "" Then
                Dummy = "1"
            End If
            If Dummy = "1" Then
                Me.chkShowCodeNo.Checked = True
            Else
                Me.chkShowCodeNo.Checked = False
            End If
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_StockBundling_Quick, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_StockBundling_Quick_ShowDesc, Me.UserCode)
            If Dummy = "" Then
                Dummy = "1"
            End If
            If Dummy = "1" Then
                Me.chkShowDesc.Checked = True
            Else
                Me.chkShowDesc.Checked = False
            End If
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_StockBundling_Quick, _
           Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_StockBundling_Quick_ShowProcessCode, Me.UserCode)
            If Dummy = "" Then
                Dummy = "1"
            End If
            If Dummy = "1" Then
                Me.chkShowProcessCode.Checked = True
            Else
                Me.chkShowProcessCode.Checked = False
            End If
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_StockBundling_Quick, _
          Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_StockBundling_Quick_ShowProcessDesc, Me.UserCode)
            If Dummy = "" Then
                Dummy = "1"
            End If
            If Dummy = "1" Then
                Me.chkShowProcessDesc.Checked = True
            Else
                Me.chkShowProcessDesc.Checked = False
            End If
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_StockBundling_Quick, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_StockBundling_Quick_ShowStartTime, Me.UserCode)
            If Dummy = "" Then
                Dummy = "1"
            End If
            If Dummy = "1" Then
                Me.chkShowStart.Checked = True
            Else
                Me.chkShowStart.Checked = False
            End If
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_StockBundling_Quick, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_StockBundling_Quick_ShowEndTime, Me.UserCode)
            If Dummy = "" Then
                Dummy = "1"
            End If
            If Dummy = "1" Then
                Me.chkShowEndTime.Checked = True
            Else
                Me.chkShowEndTime.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
