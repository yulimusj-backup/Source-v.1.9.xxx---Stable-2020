Public Class FrmTrans_Attendance_ColumnsSettings
    Dim ObjSettingsUser As New Core_Module.ClsTransaksiSettingUser
    Dim FromBtnOK As Boolean = False

    Public UserName As String = ""

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
            If Me.ChkShowActualBreak.Checked = True Then
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowActualBreak, "1", Me.UserName)
            Else
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowActualBreak, "0", Me.UserName)
            End If
            If Me.ChkShowEarlyOT.Checked = True Then
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowEartlyOT, "1", Me.UserName)
            Else
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowEartlyOT, "0", Me.UserName)
            End If
            If Me.chkShowEarlyOut.Checked = True Then
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowEarlyOut, "1", Me.UserName)
            Else
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowEarlyOut, "0", Me.UserName)
            End If
            If Me.chkShowLate.Checked = True Then
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowLate, "1", Me.UserName)
            Else
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowLate, "0", Me.UserName)
            End If
            If Me.chkShowOT.Checked = True Then
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowOT, "1", Me.UserName)
            Else
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowOT, "0", Me.UserName)
            End If
            If Me.chkShowSchedule.Checked = True Then
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowSchedule, "1", Me.UserName)
            Else
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowSchedule, "0", Me.UserName)
            End If

            ''@rosadinnn
            ''@20150908 0902
            If Me.ChkShowCalculateEarlyOT.Checked = True Then
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowCalculateEarlyOT, "1", Me.UserName)
            Else
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowCalculateEarlyOT, "0", Me.UserName)
            End If
            If Me.ChkShowCalculateLate.Checked = True Then
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowCalculateLate, "1", Me.UserName)
            Else
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowCalculateLate, "0", Me.UserName)
            End If
            If Me.ChkShowCalculateEarlyOut.Checked = True Then
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowCalculateEarlyOut, "1", Me.UserName)
            Else
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowCalculateEarlyOut, "0", Me.UserName)
            End If
            If Me.ChkShowCalculateOT.Checked = True Then
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowCalculateOT, "1", Me.UserName)
            Else
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowCalculateOT, "0", Me.UserName)
            End If

            'aul - penambahan outstation-
            If Me.ChkShowOutstation.Checked = True Then
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowOutstation, "1", Me.UserName)
            Else
                Me.ObjSettingsUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowOutstation, "0", Me.UserName)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Attendance_ColumnsSettings_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowSchedule, Me.UserName) = "1" Then
                Me.chkShowSchedule.Checked = True
            Else
                Me.chkShowSchedule.Checked = False
            End If
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
           Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowActualBreak, Me.UserName) = "1" Then
                Me.ChkShowActualBreak.Checked = True
            Else
                Me.ChkShowActualBreak.Checked = False
            End If
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
           Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowEartlyOT, Me.UserName) = "1" Then
                Me.ChkShowEarlyOT.Checked = True
            Else
                Me.ChkShowEarlyOT.Checked = False
            End If
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
           Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowOT, Me.UserName) = "1" Then
                Me.chkShowOT.Checked = True
            Else
                Me.chkShowOT.Checked = False
            End If
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowEarlyOut, Me.UserName) = "1" Then
                Me.chkShowEarlyOut.Checked = True
            Else
                Me.chkShowEarlyOut.Checked = False
            End If
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowLate, Me.UserName) = "1" Then
                Me.chkShowLate.Checked = True
            Else
                Me.chkShowLate.Checked = False
            End If

            ''@rosadinnn
            ''@20150908 0859
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowCalculateEarlyOT, Me.UserName) = "1" Then
                Me.ChkShowCalculateEarlyOT.Checked = True
            Else
                Me.ChkShowCalculateEarlyOT.Checked = False
            End If
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowCalculateLate, Me.UserName) = "1" Then
                Me.ChkShowCalculateLate.Checked = True
            Else
                Me.ChkShowCalculateLate.Checked = False
            End If
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowCalculateEarlyOut, Me.UserName) = "1" Then
                Me.ChkShowCalculateEarlyOut.Checked = True
            Else
                Me.ChkShowCalculateEarlyOut.Checked = False
            End If
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowCalculateOT, Me.UserName) = "1" Then
                Me.ChkShowCalculateOT.Checked = True
            Else
                Me.ChkShowCalculateOT.Checked = False
            End If

            'aul - penambahan outstation-
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowOutstation, Me.UserName) = "1" Then
                Me.ChkShowOutstation.Checked = True
            Else
                Me.ChkShowOutstation.Checked = False
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
End Class