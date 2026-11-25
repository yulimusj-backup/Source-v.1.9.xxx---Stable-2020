Public Class FrmTrans_Attendance_SendEmail
    Dim tblEmail As New Core_Module.DseEmailSettings.DataTable5DataTable
    Dim rowEmail As Core_Module.DseEmailSettings.DataTable5Row
    Dim obj As New Core_Module.ClsEmailSettings
    Dim ObjSetting As New Core_Module.ClsTransaksiSettingUser
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjCD As New Core_Module.ClsCekData
    Dim IsSendingEmail As Boolean = False
    Dim IsSendingCompleted As Boolean = False
    Dim vFromBtnClose As Boolean = False
    Dim ObjAttendance As New ClsTrans_Attendance
    Dim ObjKary As New Core_Module.ClsMasterKaryawan
    Dim ObjEmail As New Core_Module.ClsEmailSettings

    Public UserCode As String = ""
    Public ListPrimaryAttendance As New List(Of Double)
    Public BackColorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll

    Private Sub FrmTrans_Attendance_SendEmail_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.vFromBtnClose = False Then
                If Me.IsSendingCompleted = False Then
                    If Me.IsSendingEmail = False Then
                        'jika tidak sedang kirim email:
                        'konfirmasi, lalu keluar
                        If MsgBox("Cancel sending email?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                            e.Cancel = True
                        End If
                    Else
                        'jika sedang kirim email,
                        'konfirmasi, stop kirim email, keluar
                        If MsgBox("Stop sending email?", MsgBoxStyle.YesNo, "Confirmed") = MsgBoxResult.Yes Then
                            Me.IsSendingEmail = False
                            Me.lblProgress.Text = Me.lblProgress.Text + "(stoping process after this user)"
                            e.Cancel = True
                            Me.vFromBtnClose = True
                            Me.ControlBox = False
                            Me.btnCancel.Enabled = False
                        Else
                            e.Cancel = True
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Attendance_SendEmail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Dummy As String = ""

        Try
            'load data email account
            Me.obj.LoadEmailAccount(Me.UserCode, Me.tblEmail, Core_Module.ClsEmailSettings.EnumTipeTransaksiEmail.Attendance)
            Me.tblEmail.DefaultView.Sort = "[Account Code]"
            Me.cboEmailList.DataSource = Me.tblEmail.DefaultView
            Me.cboEmailList.DisplayMember = "Account Code"

            'load data history
            Me.txtSignByName.Text = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.AttEmail_SignBy, Me.UserCode)
            Me.txtSignByPosition.Text = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.AttEmail_SignPosition, Me.UserCode)
            Me.txtSubject.Text = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.AttEmail_Subject, Me.UserCode)

            'cek llb settings
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmEmailSettings") = False Then
                Me.llbSettingEmail.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSend.Click
        Try
            'pastikan ada akun email pengirim yang dipilih
            If Me.cboEmailList.Text = "" Then
                MsgBox("Please specify email address for sending email.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            'cek subject, sign by, position, dan email yang digunakan,
            'pastikan semua ada datanya
            Me.txtSubject.Text = Me.txtSubject.Text.Trim
            If Me.txtSubject.Text = "" Then
                MsgBox("Please specify email subject", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            Me.txtSignByName.Text = Me.txtSignByName.Text.Trim
            If Me.txtSignByName.Text = "" Then
                MsgBox("Please specify email sender name.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            Me.txtSignByPosition.Text = Me.txtSignByPosition.Text.Trim
            If Me.txtSignByPosition.Text = "" Then
                MsgBox("Please specify email sender position.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            Me.RepairEmail()
            If Me.ListPrimaryAttendance.Count > 0 Then
                If MsgBox("Start sending email?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False
                    Me.IsSendingEmail = True
                    Me.BackgroundWorker1.RunWorkerAsync()
                    Me.btnSend.Enabled = False
                    Me.btnCancel.Enabled = True
                    'Me.SendEmail()

                    'If Me.IsSendingCompleted = True Then
                    '    MsgBox("Sending email completed.", MsgBoxStyle.Information, "Information")
                    'Else
                    '    MsgBox("Email sending process interupted by user.", MsgBoxStyle.Information, "Information")
                    'End If

                    'Me.Close()
                End If
            Else
                Me.IsSendingCompleted = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RepairEmail()
        Dim ListPrimaryKary As New List(Of Double)
        Dim ij As New Core_Module.FrmMasterKaryawan_CekEmail

        Try
            'cek email, akan mengembalikan list email
            Me.CekEmail(ListPrimaryKary)

            'jika listnya ada isi, maka repair email melalui master form karyawan repair email
            If ListPrimaryKary.Count > 0 Then
                ij.UserCode = Me.UserCode
                ij.BackColorAll = Me.BackColorAll
                ij.ListPrimaryKaryawan = ListPrimaryKary
                ij.ShowDialog()

                'cek email 1x lagi, jika masih ada maka untuk setiap baris yang menggunakan karyawan ybs, tidak jadi dikirimkan emailnya
                Me.RemoveAttendance()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryKary = Nothing
            ij = Nothing
        End Try
    End Sub

    Private Sub RemoveAttendance()
        Dim PrimaryAttendance As Double = -1
        Dim PrimaryKary As Double = -1
        Dim DTA As New Core_Module.DseMasterKaryawanTableAdapters.tblmaster_karyawan3TableAdapter
        Dim TBL As New Core_Module.DseMasterKaryawan.tblmaster_karyawan3DataTable
        Dim row As Core_Module.DseMasterKaryawan.tblmaster_karyawan3Row
        Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
        Dim ListRemove As New List(Of Double)
        Dim Dummy As String = ""
        Dim DummyInt As Double = 0

        Try
            Core_Module.SetConstring_DTAKary(DTA)
            DTA.Fill(TBL)

            'untuk setiap baris dalam dgv, yang kirim email, cek apakah email karyawannya bener
            For Each PrimaryAttendance In Me.ListPrimaryAttendance
                PrimaryKary = Me.ObjAttendance.GetPrimaryEmployeeAttendance(PrimaryAttendance)
                If PrimaryKary <> -1 Then
                    'cek emailnya
                    TBL.DefaultView.RowFilter = "primary = " + PrimaryKary.ToString
                    If TBL.DefaultView.Count > 0 Then
                        row = TBL.DefaultView(0).Row
                        row.email = row.email.Trim
                        If row.email = "" Then
                            ListRemove.Add(PrimaryAttendance)
                        Else
                            If Expression.IsMatch(row.email) = False Then
                                ListRemove.Add(PrimaryAttendance)
                            End If
                        End If
                    End If
                End If
            Next

            For Each PrimaryAttendance In ListRemove
                ListPrimaryAttendance.Remove(PrimaryAttendance)
                DummyInt += 1
                If Dummy = "" Then
                    Dummy = DummyInt.ToString + ". " + Me.ObjAttendance.GetNoTransByPrimary(PrimaryAttendance) + _
                    " - " + Me.ObjKary.GetFullNameByPrimary(Me.ObjAttendance.GetPrimaryEmployeeAttendance(PrimaryAttendance))
                Else
                    Dummy = Dummy + vbCrLf + DummyInt.ToString + ". " + Me.ObjAttendance.GetNoTransByPrimary(PrimaryAttendance) + _
                    " - " + Me.ObjKary.GetFullNameByPrimary(Me.ObjAttendance.GetPrimaryEmployeeAttendance(PrimaryAttendance))
                End If
            Next

            If Dummy <> "" Then
                MsgBox("The following attendace data won't be send due to invalid email address:" + _
                vbCrLf + Dummy, MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ListRemove = Nothing
        End Try
    End Sub

    Private Sub CekEmail(ByVal ListPrimaryKary As List(Of Double))
        Dim PrimaryAttendance As Double = -1
        Dim PrimaryKary As Double = -1
        Dim DTA As New Core_Module.DseMasterKaryawanTableAdapters.tblmaster_karyawan3TableAdapter
        Dim TBL As New Core_Module.DseMasterKaryawan.tblmaster_karyawan3DataTable
        Dim row As Core_Module.DseMasterKaryawan.tblmaster_karyawan3Row
        Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")

        Try
            Core_Module.SetConstring_DTAKary(DTA)
            DTA.Fill(TBL)
            ListPrimaryKary.Clear()

            'untuk setiap baris dalam dgv, yang kirim email, cek apakah email karyawannya bener
            For Each PrimaryAttendance In Me.ListPrimaryAttendance
                PrimaryKary = Me.ObjAttendance.GetPrimaryEmployeeAttendance(PrimaryAttendance)
                If PrimaryKary <> -1 Then
                    If ListPrimaryKary.Contains(PrimaryKary) = False Then
                        'cek emailnya
                        TBL.DefaultView.RowFilter = "primary = " + PrimaryKary.ToString
                        If TBL.DefaultView.Count > 0 Then
                            row = TBL.DefaultView(0).Row
                            row.email = row.email.Trim
                            If row.email = "" Then
                                ListPrimaryKary.Add(PrimaryKary)
                            Else
                                If Expression.IsMatch(row.email) = False Then
                                    ListPrimaryKary.Add(PrimaryKary)
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub SendEmail()
        Dim FrmCetak As New FrmTrans_Attendance_Cetak
        Dim fileExists As Boolean = False
        Dim ListAdd As New List(Of String)
        Dim ListCC As New List(Of String)
        Dim ListFileName As New List(Of String)
        Dim DTA1 As New DseTrans_AttendanceTableAdapters.DataTable8TableAdapter
        Dim TBL1 As New DseTrans_Attendance.DataTable8DataTable
        Dim row1 As DseTrans_Attendance.DataTable8Row
        Dim DummyMessage As String = ""
        Dim UseSSL As Boolean = False
        Dim NoTrans As String = ""

        Try
            'simpan data ke tabel historical
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.AttEmail_SignBy, _
            Me.txtSignByName.Text, Me.UserCode)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.AttEmail_SignPosition, _
            Me.txtSignByPosition.Text, Me.UserCode)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.AttEmail_Subject, _
            Me.txtSubject.Text, Me.UserCode)

            'start pengiriman email
            'cc address
            Me.rowEmail = Me.tblEmail.DefaultView(Me.cboEmailList.SelectedIndex).Row
            Dim ObjEncrypt As New Core_Module.ClsEncrypt(Me.ObjEmail.GetPassPhrase + Me.rowEmail.Account_Code + Me.rowEmail.Primary_User.ToString, Me.obj.GetIniVector)
            Me.rowEmail.Password = ObjEncrypt.Decrypt(Me.rowEmail.Password)
            ObjEncrypt = Nothing

            ListCC.Clear()
            ListCC.Add(Me.rowEmail.CC)
           

            Me.prgAll.Value = 0
            Me.prgAll.Maximum = Me.ListPrimaryAttendance.Count

            For Each vPrimary As Double In Me.ListPrimaryAttendance
                If Me.IsSendingEmail = False Then
                    Me.IsSendingCompleted = False
                    'Me.IsSendingEmail = True
                    Exit For
                End If
                Me.prgAll.Value += 1
                DTA1.Fill(TBL1, vPrimary)
                If TBL1.Rows.Count > 0 Then
                    row1 = TBL1.Rows(0)
                    Me.lblProgress.Text = "Data " + Me.prgAll.Value.ToString + " of " + Me.prgAll.Maximum.ToString + " : " + row1.no_transaksi + " - " + row1.Full_Name
                    ListAdd.Clear()
                    ListAdd.Add(row1.email)

                    FrmCetak.FromAttendanceEmail = True
                    'FrmCetak.Show()
                    FrmCetak.CreatePDF(vPrimary, row1.Full_Name)
                    fileExists = My.Computer.FileSystem.FileExists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\Attendance-" + row1.Full_Name + ".pdf")
                    ListFileName.Clear()
                    ListFileName.Add(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\Attendance-" + row1.Full_Name + ".pdf")
                    If fileExists = True Then
                        'kirim email
                        DummyMessage = Me.rowEmail.Message
                        DummyMessage = DummyMessage.Replace("<<Date>>", Format(Me.dtpDue.Value, "dd MMM yyyy"))
                        DummyMessage = DummyMessage.Replace("<<Subject>>", Me.txtSubject.Text)
                        DummyMessage = DummyMessage.Replace("<<EmpName>>", row1.Full_Name)
                        DummyMessage = DummyMessage.Replace("<<SignBy>>", Me.txtSignByName.Text)
                        DummyMessage = DummyMessage.Replace("<<Position>>", Me.txtSignByPosition.Text)
                        If Me.rowEmail.use_ssl = 1 Then
                            UseSSL = True
                        Else
                            UseSSL = False
                        End If
                        Me.obj.SendEmail(ListAdd, ListCC, Me.txtSubject.Text, DummyMessage, ListFileName, Me.rowEmail.Email_From, _
                        Me.rowEmail.Reply_To, Me.rowEmail.Username, Me.rowEmail.Password, Me.rowEmail.SMTP_Server, Me.rowEmail.SMTP_Port, UseSSL)
                    End If
                    'Me.obj = Nothing
                    'hapus filenya
                    My.Computer.FileSystem.DeleteFile(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\Attendance-" + row1.Full_Name + ".pdf", FileIO.UIOption.OnlyErrorDialogs, _
                    FileIO.RecycleOption.DeletePermanently)
                End If
            Next

            If Me.IsSendingEmail = True Then
                Me.IsSendingCompleted = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListAdd = Nothing
            ListCC = Nothing
            ListFileName = Nothing
            DTA1.Dispose()
            TBL1.Dispose()
            FrmCetak.Dispose()
        End Try
    End Sub

    Private Sub Keluar()
        Try
            If Me.IsSendingEmail = False Then
                'jika tidak sedang kirim email:
                'konfirmasi, lalu keluar
                If MsgBox("Cancel sending email?") = MsgBoxResult.Yes Then
                    Me.Close()
                End If
            Else
                'jika sedang kirim email,
                'konfirmasi, stop kirim email, keluar
                If MsgBox("Stop sending email?", MsgBoxStyle.YesNo, "Confirmed") = MsgBoxResult.Yes Then
                    Me.IsSendingEmail = False
                    Me.lblProgress.Text = Me.lblProgress.Text + "(stoping process after this user)"
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If Me.IsSendingEmail = False Then
                'jika tidak sedang kirim email:
                'konfirmasi, lalu keluar
                If MsgBox("Cancel sending email?") = MsgBoxResult.Yes Then
                    Me.Close()
                End If
            Else
                'jika sedang kirim email,
                'konfirmasi, stop kirim email, keluar
                If MsgBox("Stop sending email?", MsgBoxStyle.YesNo, "Confirmed") = MsgBoxResult.Yes Then
                    Me.IsSendingEmail = False
                    Me.lblProgress.Text = Me.lblProgress.Text + "(stoping process after this user)"
                    Me.vFromBtnClose = True
                    Me.btnCancel.Enabled = False
                    Me.ControlBox = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbSettingEmail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbSettingEmail.LinkClicked
        Dim frm As New Core_Module.FrmEmailSettings
        Try
            frm.UserCode = Me.UserCode
            frm.BackColorAll = Me.BackColorAll
            frm.ShowDialog()
            Me.obj.LoadEmailAccount(Me.UserCode, Me.tblEmail, Core_Module.ClsEmailSettings.EnumTipeTransaksiEmail.Attendance)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
        End Try
    End Sub

    Private Sub txt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSubject.GotFocus, txtSignByName.GotFocus, txtSignByPosition.GotFocus
        Dim txt As Windows.Forms.TextBox
        Try
            txt = TryCast(sender, Windows.Forms.TextBox)
            If IsNothing(txt) = False Then
                If txt.Enabled = True Then
                    If txt.ReadOnly = True Then
                        txt.BackColor = Me.TxtReadOnly.BackColor
                    Else
                        txt.BackColor = Me.BackColorAll
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSubject.LostFocus, txtSignByName.LostFocus, txtSignByPosition.LostFocus
        Dim txt As Windows.Forms.TextBox
        Try
            txt = TryCast(sender, Windows.Forms.TextBox)
            If IsNothing(txt) = False Then
                If txt.Enabled = True Then
                    If txt.ReadOnly = True Then
                        txt.BackColor = Me.TxtReadOnly.BackColor
                    Else
                        txt.BackColor = Me.TxtAcuan.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Me.ObjCD.UpdateThreading()
            Me.SendEmail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.IsSendingCompleted = True Then
                MsgBox("Sending email completed.", MsgBoxStyle.Information, "Information")
            Else
                MsgBox("Email sending process interupted by user.", MsgBoxStyle.Information, "Information")
            End If

            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class