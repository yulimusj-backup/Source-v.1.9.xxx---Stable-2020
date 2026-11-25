Option Explicit On
Option Strict Off
Partial Public Class FrmSettingAP_AR
    Dim ObjSch As New Inventory_Module.ClsTrans_Schedule

    Private Sub btnSet_AllocatedWS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSet_AllocatedWS.Click
        Dim ij As New Windows.Forms.ColorDialog

        Try
            If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.pb_AllocatedWS.BackColor = ij.Color
                Me.SettingProductionBerubah = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanDataSettingProduction()
        Try
            If Me.SettingProductionBerubah = True Then
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
               Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_AllocatedWS_a, CInt(Me.pb_AllocatedWS.BackColor.A))
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_AllocatedWS_r, CInt(Me.pb_AllocatedWS.BackColor.R))
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_AllocatedWS_g, CInt(Me.pb_AllocatedWS.BackColor.G))
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_AllocatedWS_b, CInt(Me.pb_AllocatedWS.BackColor.B))

                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchHor_FirmPlanned1, Me.TxtFirmPlan.Text)

                If Me.chkProd_UseCustomMOPrint.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ManufacturingOrder, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.ManufacturingOrder_CustomPrintMO, "1")
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ManufacturingOrder, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.ManufacturingOrder_CustomPrintMO, "0")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeFieldSettingProduksi()
        Dim Dummy As String = ""
        Dim DefPrecQuot As Integer = 0
        Dim a, r, g, b As String

        Try
            'salin data custom invoice print
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                        Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchHor_FirmPlanned1)
            If Dummy = "" Then
                Dummy = Me.ObjSch.GetDefaultSchHor_FirmPlanned1
            End If
            Me.TxtFirmPlan.Text = Dummy

            a = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_AllocatedWS_a)
            If a = "" Then
                'load default color
                Me.pb_AllocatedWS.BackColor = Me.ObjSch.GetDefaultColor_AllocatedWS
            Else
                'load dari db
                r = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_AllocatedWS_r)
                g = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_AllocatedWS_g)
                b = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_AllocatedWS_b)
                Me.pb_AllocatedWS.BackColor = Color.FromArgb(a, r, g, b)
            End If

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ManufacturingOrder, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ManufacturingOrder_CustomPrintJO)
            If Dummy = "" Or Dummy = "0" Then
                Me.chkProd_UseCustomMOPrint.Checked = False
            Else
                Me.chkProd_UseCustomMOPrint.Checked = True
            End If
            Me.SettingProductionBerubah = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TextBox_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFirmPlan.Enter
        Dim TxtBox As Windows.Forms.TextBox

        Try
            TxtBox = TryCast(sender, Windows.Forms.TextBox)
            If IsNothing(TxtBox) = False Then
                Me.Buffer = Me.TxtFirmPlan.Text
                Me.ObjCD.Number_Enter(TxtBox, sender, e)
                TxtBox.DeselectAll()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TextBox_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFirmPlan.GotFocus
        Dim TxtBox As Windows.Forms.TextBox

        Try
            TxtBox = TryCast(sender, Windows.Forms.TextBox)
            If IsNothing(TxtBox) = False Then
                TxtBox.BackColor = Me.BackColorAll
                TxtBox.DeselectAll()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtFirmPlan.KeyDown
        Dim TxtBox As Windows.Forms.TextBox

        Try
            TxtBox = TryCast(sender, Windows.Forms.TextBox)
            If IsNothing(TxtBox) = False Then
                Me.ObjCD.Number_KeyDown(TxtBox, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TextBox1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtFirmPlan.KeyUp
        Dim TxtBox As Windows.Forms.TextBox

        Try
            TxtBox = TryCast(sender, Windows.Forms.TextBox)
            If IsNothing(TxtBox) = False Then
                Me.ObjCD.Number_Leave(TxtBox, sender, e, 0, Me.ObjCD)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TextBox_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFirmPlan.Leave
        Dim TxtBox As Windows.Forms.TextBox

        Try
            TxtBox = TryCast(sender, Windows.Forms.TextBox)
            If IsNothing(TxtBox) = False Then
                Me.ObjCD.Number_Leave(TxtBox, sender, e, 0, Me.ObjCD)
                If CDec(Me.Buffer) <> CDec(Me.TxtFirmPlan.Text) Then
                    Me.SettingProductionBerubah = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Textbox_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFirmPlan.LostFocus
        Dim TxtBox As Windows.Forms.TextBox

        Try
            TxtBox = TryCast(sender, Windows.Forms.TextBox)
            If IsNothing(TxtBox) = False Then
                TxtBox.BackColor = Color.White
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkProd_UseCustomMOPrint_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkProd_UseCustomMOPrint.CheckedChanged
        Try
            Me.SettingProductionBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
