Option Explicit On
Option Strict Off

Public Class FrmSettingProduksi
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser
    Dim ObjSch As New ClsTrans_Schedule
    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub btnSet_AllocatedWS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSet_AllocatedWS.Click
        Try
            If Me.ColorDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.pb_AllocatedWS.BackColor = Me.ColorDialog1.Color
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim Dummy As String = ""

        Try
            'berikan peringatan kalo ada yang sama
            If Me.pb_AllocatedWS.BackColor = Me.pb_Offtime.BackColor Or _
            Me.pb_AllocatedWS.BackColor = Me.pb_SchDT.BackColor Or _
            Me.pb_AllocatedWS.BackColor = Me.pb_SchOT.BackColor Or _
            Me.pb_Offtime.BackColor = Me.pb_SchDT.BackColor Or _
            Me.pb_Offtime.BackColor = Me.pb_SchOT.BackColor Then
                If MsgBox("There is one or more color use more than once, continue anyway?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Exit Sub
                End If
            End If

            'firm planned dan end date harus positif
            If CInt(Me.TxtEndDate.Text) <= 0 Then
                MsgBox("Please specify a positive number form end date data.", MsgBoxStyle.Information, "Information")
                Me.TxtEndDate.Focus()
                Exit Sub
            End If
            'start date harus negatif
            If CInt(Me.TxtStartDate.Text) >= 0 Then
                MsgBox("Please specify a negative number form start date data.", MsgBoxStyle.Information, "Information")
                Me.TxtStartDate.Focus()
                Exit Sub
            End If

            'simpan ke dalam db
            If MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                If Me.chkShowFreeTime.Checked = True Then
                    Dummy = "1"
                Else
                    Dummy = "0"
                End If
                Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ProductionSetting_ShowFreeTime, Dummy, Me.UserCode)
                If Me.chkShowOffTime.Checked = True Then
                    Dummy = "1"
                Else
                    Dummy = "0"
                End If
                Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ProductionSetting_ShowOffTime, Dummy, Me.UserCode)

                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchHor_EndDate, Me.TxtEndDate.Text)
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchHor_StartDate, Me.TxtStartDate.Text)
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchHor_FirmPlanned1, Me.TxtFirmPlan.Text)

                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_AllocatedWS_a, CInt(Me.pb_AllocatedWS.BackColor.A))
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_AllocatedWS_r, CInt(Me.pb_AllocatedWS.BackColor.R))
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_AllocatedWS_g, CInt(Me.pb_AllocatedWS.BackColor.G))
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_AllocatedWS_b, CInt(Me.pb_AllocatedWS.BackColor.B))

                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_Offtime_a, CInt(Me.pb_Offtime.BackColor.A))
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_Offtime_r, CInt(Me.pb_Offtime.BackColor.R))
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_Offtime_g, CInt(Me.pb_Offtime.BackColor.G))
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_Offtime_b, CInt(Me.pb_Offtime.BackColor.B))

                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchDT_a, CInt(Me.pb_SchDT.BackColor.A))
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchDT_r, CInt(Me.pb_SchDT.BackColor.R))
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchDT_g, CInt(Me.pb_SchDT.BackColor.G))
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchDT_b, CInt(Me.pb_SchDT.BackColor.B))

                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchOT_a, CInt(Me.pb_SchOT.BackColor.A))
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchOT_r, CInt(Me.pb_SchOT.BackColor.R))
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchOT_g, CInt(Me.pb_SchOT.BackColor.G))
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchOT_b, CInt(Me.pb_SchOT.BackColor.B))

                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSet_Offtime_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSet_Offtime.Click
        Try
            If Me.ColorDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.pb_Offtime.BackColor = Me.ColorDialog1.Color
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSet_SchDT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSet_SchDT.Click
        Try
            If Me.ColorDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.pb_SchDT.BackColor = Me.ColorDialog1.Color
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSet_OT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSet_OT.Click
        Try
            If Me.ColorDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.pb_SchOT.BackColor = Me.ColorDialog1.Color
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmSettingProduksi_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim a, r, g, b As String
        Dim Dummy As String = ""

        Try
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ProductionSetting, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ProductionSetting_ShowFreeTime, Me.UserCode)
            If Dummy = "" Then
                If Me.ObjSch.GetDefault_ShowFreeTime = True Then
                    Dummy = "1"
                Else
                    Dummy = "0"
                End If
                Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ProductionSetting_ShowFreeTime, Dummy, Me.UserCode)
            End If
            If Dummy = "1" Then
                Me.chkShowFreeTime.Checked = True
            Else
                Me.chkShowFreeTime.Checked = False
            End If
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ProductionSetting, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ProductionSetting_ShowOffTime, Me.UserCode)
            If Dummy = "" Then
                If Me.ObjSch.GetDefault_ShowOffTime = True Then
                    Dummy = "1"
                Else
                    Dummy = "0"
                End If
                Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ProductionSetting_ShowOffTime, Dummy, Me.UserCode)
            End If
            If Dummy = "1" Then
                Me.chkShowOffTime.Checked = True
            Else
                Me.chkShowOffTime.Checked = False
            End If

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchHor_EndDate)
            If Dummy = "" Then
                Dummy = Me.ObjSch.GetDefaultSchHor_EndDate
            End If
            Me.TxtEndDate.Text = Dummy
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchHor_StartDate)
            If Dummy = "" Then
                Dummy = Me.ObjSch.GetDefaultSchHor_StartDate
            End If
            Me.TxtStartDate.Text = Dummy
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

            a = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_Offtime_a)
            If a = "" Then
                'load default color
                Me.pb_Offtime.BackColor = Me.ObjSch.GetDefaultColor_OffTime
            Else
                'load dari db
                r = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_Offtime_r)
                g = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_Offtime_g)
                b = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_Offtime_b)
                Me.pb_Offtime.BackColor = Color.FromArgb(a, r, g, b)
            End If

            a = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchDT_a)
            If a = "" Then
                'load default color
                Me.pb_SchDT.BackColor = Me.ObjSch.GetDefaultColor_SchDT
            Else
                'load dari db
                r = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchDT_r)
                g = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchDT_g)
                b = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchDT_b)
                Me.pb_SchDT.BackColor = Color.FromArgb(a, r, g, b)
            End If

            a = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchOT_a)
            If a = "" Then
                'load default color
                Me.pb_SchOT.BackColor = Me.ObjSch.GetDefaultColor_SchOT
            Else
                'load dari db
                r = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchOT_r)
                g = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchOT_g)
                b = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchOT_b)
                Me.pb_SchOT.BackColor = Color.FromArgb(a, r, g, b)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TextBox_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStartDate.Enter, TxtEndDate.Enter, TxtFirmPlan.Enter
        Dim TxtBox As Windows.Forms.TextBox

        Try
            TxtBox = TryCast(sender, Windows.Forms.TextBox)
            If IsNothing(TxtBox) = False Then
                Me.ObjCD.Number_Enter(TxtBox, sender, e)
                TxtBox.DeselectAll()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TextBox_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStartDate.GotFocus, TxtEndDate.GotFocus, TxtFirmPlan.GotFocus
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

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtStartDate.KeyDown, TxtEndDate.KeyDown, TxtFirmPlan.KeyDown
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

    Private Sub TextBox1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtStartDate.KeyUp, TxtEndDate.KeyUp, TxtFirmPlan.KeyUp
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

    Private Sub TextBox_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStartDate.Leave, TxtEndDate.Leave, TxtFirmPlan.Leave
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

    Private Sub Textbox_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStartDate.LostFocus, TxtEndDate.LostFocus, TxtFirmPlan.LostFocus
        Dim TxtBox As Windows.Forms.TextBox

        Try
            TxtBox = TryCast(sender, Windows.Forms.TextBox)
            If IsNothing(TxtBox) = False Then
                TxtBox.BackColor = Me.TxtRef.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class