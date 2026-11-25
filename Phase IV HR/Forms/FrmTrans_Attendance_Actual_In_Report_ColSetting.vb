Public Class FrmTrans_Attendance_Actual_In_Report_ColSetting
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
            If Me.ChkShowSch.Checked = True Then
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_ClockInOut_ShowSchedule, "1", Me.UserCode)
            Else
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_ClockInOut_ShowSchedule, "0", Me.UserCode)
            End If
            If Me.ChkShowSchBreak.Checked = True Then
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_ClockInOut_ShowScheduleBreak, "1", Me.UserCode)
            Else
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_ClockInOut_ShowScheduleBreak, "0", Me.UserCode)
            End If

            If Me.chkShowAct.Checked = True Then
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_ClockInOut_ShowActual, "1", Me.UserCode)
            Else
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_ClockInOut_ShowActual, "0", Me.UserCode)
            End If
            If Me.ChkShowActBreak.Checked = True Then
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_ClockInOut_ShowActualBreak, "1", Me.UserCode)
            Else
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_ClockInOut_ShowActualBreak, "0", Me.UserCode)
            End If

            If Me.ChkShowShift.Checked = True Then
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_ClockInOut_ShowShift, "1", Me.UserCode)
            Else
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_ClockInOut_ShowShift, "0", Me.UserCode)
            End If

            If Me.ChkShowDept.Checked = True Then
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_ClockInOut_ShowDept, "1", Me.UserCode)
            Else
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_ClockInOut_ShowDept, "0", Me.UserCode)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Attendance_ColumnsSettings_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Text = "Columns Settings Format 2"

            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_ClockInOut_ShowSchedule, Me.UserCode) = "1" Then
                Me.ChkShowSch.Checked = True

                If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_ClockInOut_ShowScheduleBreak, Me.UserCode) = "1" Then
                    Me.ChkShowSchBreak.Checked = True
                Else
                    Me.ChkShowSchBreak.Checked = False
                End If
            Else
                Me.ChkShowSch.Checked = False
                Me.ChkShowSchBreak.Checked = False
                Me.ChkShowSchBreak.Enabled = False
            End If

            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_ClockInOut_ShowActual, Me.UserCode) = "1" Then
                Me.chkShowAct.Checked = True

                If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_ClockInOut_ShowActualBreak, Me.UserCode) = "1" Then
                    Me.ChkShowActBreak.Checked = True
                Else
                    Me.ChkShowActBreak.Checked = False
                End If
            Else
                Me.chkShowAct.Checked = False
                Me.ChkShowActBreak.Checked = False
                Me.ChkShowActBreak.Enabled = False
            End If

            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_ClockInOut_ShowShift, Me.UserCode) = "1" Then
                Me.ChkShowShift.Checked = True
            Else
                Me.ChkShowShift.Checked = False
            End If

            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_ClockInOut_ShowDept, Me.UserCode) = "1" Then
                Me.ChkShowDept.Checked = True
            Else
                Me.ChkShowDept.Checked = False
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

    Private Sub ChkShowSch_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkShowSch.CheckedChanged
        Try
            If Me.ChkShowSch.Checked = True Then
                Me.ChkShowSchBreak.Enabled = True
                Me.ChkShowSchBreak.Checked = True
            Else
                Me.ChkShowSchBreak.Checked = False
                Me.ChkShowSchBreak.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkShowAct_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowAct.CheckedChanged
        Try
            If Me.chkShowAct.Checked = True Then
                Me.ChkShowActBreak.Enabled = True
                Me.ChkShowActBreak.Checked = True
            Else
                Me.ChkShowActBreak.Checked = False
                Me.ChkShowActBreak.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkShowDept_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkShowDept.CheckedChanged

    End Sub
End Class