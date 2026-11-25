Public Class FrmTrans_SPL_ColumnsSettings
    Dim ObjSettingsUser As New Core_Module.ClsTransaksiSettingUser
    Dim FromBtnOK As Boolean = False

    Public UserCode As String = ""

    Private Sub FrmTrans_Attendance_ColumnsSettings_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.FromBtnOK = False Then
                If MsgBox("Save settings before close?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.SimpanSettings()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanSettings()
        Try
            'simpan data ke dalam database
            'OT
            If Me.ChkShowEarlyOT.Checked = True Then
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowEarlyOT, "1", Me.UserCode)
            Else
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowEarlyOT, "0", Me.UserCode)
            End If
            If Me.chkShowOT.Checked = True Then
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowOT, "1", Me.UserCode)
            Else
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowOT, "0", Me.UserCode)
            End If
            If Me.ChkShowHolidayOT.Checked = True Then
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowHolidayOT, "1", Me.UserCode)
            Else
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowHolidayOT, "0", Me.UserCode)
            End If

            'Delay
            If Me.ChkShowEarlyDelay.Checked = True Then
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowEarlyDelay, "1", Me.UserCode)
            Else
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowEarlyDelay, "0", Me.UserCode)
            End If
            If Me.ChkShowDelay.Checked = True Then
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowDelay, "1", Me.UserCode)
            Else
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowDelay, "0", Me.UserCode)
            End If
            If Me.ChkShowHolidayDelay.Checked = True Then
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowHolidayDelay, "1", Me.UserCode)
            Else
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowHolidayDelay, "0", Me.UserCode)
            End If

            'Break
            If Me.ChkShowEarlyBreak.Checked = True Then
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowEarlyBreak, "1", Me.UserCode)
            Else
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowEarlyBreak, "0", Me.UserCode)
            End If
            If Me.ChkShowBreak.Checked = True Then
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowBreak, "1", Me.UserCode)
            Else
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowBreak, "0", Me.UserCode)
            End If
            If Me.ChkShowHolidayBreak.Checked = True Then
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowHolidayBreak, "1", Me.UserCode)
            Else
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowHolidayBreak, "0", Me.UserCode)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Attendance_ColumnsSettings_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Text = "Columns Settings Schedule Setup"
            Me.ChkShowHolidayOT.Text = "Show Holiday OT"
            'Early OT
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowEarlyOT, Me.UserCode) = "1" Then
                Me.ChkShowEarlyOT.Checked = True

                'Delay
                If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                   Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowEarlyDelay, Me.UserCode) = "1" Then
                    Me.ChkShowEarlyDelay.Checked = True
                Else
                    Me.ChkShowEarlyDelay.Checked = False
                End If
                'Break
                If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowEarlyBreak, Me.UserCode) = "1" Then
                    Me.ChkShowEarlyBreak.Checked = True
                Else
                    Me.ChkShowEarlyBreak.Checked = False
                End If
            Else
                Me.ChkShowEarlyOT.Checked = False
                Me.ChkShowEarlyDelay.Checked = False
                Me.ChkShowEarlyBreak.Checked = False
                Me.ChkShowEarlyDelay.Enabled = False
                Me.ChkShowEarlyBreak.Enabled = False
            End If

            'OT
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowOT, Me.UserCode) = "1" Then
                Me.chkShowOT.Checked = True

                'Delay
                If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowDelay, Me.UserCode) = "1" Then
                    Me.ChkShowDelay.Checked = True
                Else
                    Me.ChkShowDelay.Checked = False
                End If
                'Break
                If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowBreak, Me.UserCode) = "1" Then
                    Me.ChkShowBreak.Checked = True
                Else
                    Me.ChkShowBreak.Checked = False
                End If
            Else
                Me.chkShowOT.Checked = False
                Me.ChkShowDelay.Checked = False
                Me.ChkShowBreak.Checked = False
                Me.ChkShowDelay.Enabled = False
                Me.ChkShowBreak.Enabled = False
            End If

            'Holiday OT
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowHolidayOT, Me.UserCode) = "1" Then
                Me.ChkShowHolidayOT.Checked = True

                'Delay
                If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowHolidayDelay, Me.UserCode) = "1" Then
                    Me.ChkShowHolidayDelay.Checked = True
                Else
                    Me.ChkShowHolidayDelay.Checked = False
                End If
                'Break
                If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowHolidayBreak, Me.UserCode) = "1" Then
                    Me.ChkShowHolidayBreak.Checked = True
                Else
                    Me.ChkShowHolidayBreak.Checked = False
                End If
            Else
                Me.ChkShowHolidayOT.Checked = False
                Me.ChkShowHolidayDelay.Checked = False
                Me.ChkShowHolidayBreak.Checked = False
                Me.ChkShowHolidayDelay.Enabled = False
                Me.ChkShowHolidayBreak.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If MsgBox("Save settings?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                Me.Close()
                Exit Sub
            End If
            Me.SimpanSettings()
            Me.FromBtnOK = True
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkShowEarlyOT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkShowEarlyOT.CheckedChanged
        Try
            If Me.ChkShowEarlyOT.Checked = True Then
                Me.ChkShowEarlyDelay.Enabled = True
                Me.ChkShowEarlyBreak.Enabled = True
                Me.ChkShowEarlyDelay.Checked = True
                Me.ChkShowEarlyBreak.Checked = True
            Else
                Me.ChkShowEarlyDelay.Checked = False
                Me.ChkShowEarlyBreak.Checked = False
                Me.ChkShowEarlyDelay.Enabled = False
                Me.ChkShowEarlyBreak.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkShowOT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowOT.CheckedChanged
        Try
            If Me.chkShowOT.Checked = True Then
                Me.ChkShowDelay.Enabled = True
                Me.ChkShowBreak.Enabled = True
                Me.ChkShowDelay.Checked = True
                Me.ChkShowBreak.Checked = True
            Else
                Me.ChkShowDelay.Checked = False
                Me.ChkShowBreak.Checked = False
                Me.ChkShowDelay.Enabled = False
                Me.ChkShowBreak.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkShowHolidayOT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkShowHolidayOT.CheckedChanged
        Try
            If Me.ChkShowHolidayOT.Checked = True Then
                Me.ChkShowHolidayDelay.Enabled = True
                Me.ChkShowHolidayBreak.Enabled = True
                Me.ChkShowHolidayDelay.Checked = True
                Me.ChkShowHolidayBreak.Checked = True
            Else
                Me.ChkShowHolidayDelay.Checked = False
                Me.ChkShowHolidayBreak.Checked = False
                Me.ChkShowHolidayDelay.Enabled = False
                Me.ChkShowHolidayBreak.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class