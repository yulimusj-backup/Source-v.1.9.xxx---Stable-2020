Option Explicit On
Option Strict Off
Imports System.Windows.Forms

Public Class FrmTransaksiCuti_Pengajuan
    Public UserCode As String = ""
    Public BackColorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll
    Public TBLDuration As New DseTrans_Cuti_Pengajuan.Tbl_Explain_DurationDataTable
    Public ij As New FrmTransaksiCuti_Pengajuan_Balance
    Public FrmExplain As New FrmTransaksiCuti_Pengajuan_ExplainDuration
    Public TblList_StatusExplain As New DseTrans_Cuti_Pengajuan.List_StatusExplainDataTable  '-
    Dim BufferDate As String = ""
    Dim Buffer As String = ""
    Dim Baru As Boolean = True
    Dim AktifForm As Boolean = True
    Public ObjCekData As New Core_Module.ClsCekData
    Dim ObjPengajuanCuti As New ClsTrans_Cuti_Pengajuan
    Dim ObjKaryawan As New Core_Module.ClsMasterKaryawan
    Dim ObjJob As New Core_Module.ClsMasterJobPosition
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjCuti As New Core_Module.ClsMasterCuti
    Dim FrmStatus As New FrmTransaksiCuti_Approval_StatusNotes
    'Dim tbl As New DseTrans_Cuti_Pengajuan.tblmaster_cutiDataTable
    Dim PrivilegeUser As ClsTrans_Cuti_Pengajuan.EnumPrivilagePengajuanCuti
    Dim mulai As Boolean = False
    Dim ObjAttendanceItem As New ClsTrans_Attendance_Item
    Dim ObjItem As New ClsTrans_Cuti_Pengajuan_Item

    Public Sub SetBalance()
        Dim ObjMasterCutiItem As New Core_Module.ClsMasterCuti_Item
        Dim NilaiBalance As Decimal = 0

        Try
            If Me.TxtPrimaryLeave.Text <> "" Then
                'balance baru = balance db - (txtduration - txtduration_db)
                If Me.TxtStatus.Text.ToLower = "Rejected".ToLower Then
                    NilaiBalance = ObjMasterCutiItem.GetNilaiBalance(Me.TxtPrimaryLeave.Text, Me.TxtPrimaryEmployee.Text)
                Else
                    NilaiBalance = ObjMasterCutiItem.GetNilaiBalance(Me.TxtPrimaryLeave.Text, Me.TxtPrimaryEmployee.Text) - _
                    (CDbl(Me.TxtDuration.Text) - Me.ObjPengajuanCuti.GetDuration(Me.TxtPrimary.Text, Me.TxtPrimaryEmployee.Text, Me.TxtPrimaryLeave.Text))
                End If
                Me.TxtBalance.Text = Me.ObjCekData.DigitGrouping(NilaiBalance, 1)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjMasterCutiItem = Nothing
        End Try
    End Sub

    Private Sub SetEmployeeCode(ByVal vPrimaryUser As Double)
        Dim PrimaryKaryawan As Double = -1
        Try
            'cari primary karyawan dari primary user
            PrimaryKaryawan = Me.ObjPengajuanCuti.GetPrimaryKaryawabByUser(vPrimaryUser)
            Me.TxtPrimaryEmployee.Text = PrimaryKaryawan
            Me.TxtEmployeeCode.Text = Me.ObjKaryawan.GetCodeByPrimary(PrimaryKaryawan)
            Me.TxtEmployeeName.Text = Me.ObjKaryawan.GetFullNameByPrimary(PrimaryKaryawan)
            Me.TxtPrimaryPosition.Text = Me.ObjKaryawan.GetJobByPrimary(PrimaryKaryawan)
            Me.TxtPosition.Text = Me.ObjJob.GetCode(Me.TxtPrimaryPosition.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function BlokSimbol(ByVal obj As System.Windows.Forms.TextBox, ByVal e As Windows.Forms.KeyPressEventArgs) As Boolean
        Try
            If e.KeyChar = "'" Or _
            e.KeyChar = "%" Or _
            e.KeyChar = "*" Or _
            e.KeyChar = "#" Then
                obj.ReadOnly = True
                obj.BackColor = Me.BackColorAll
            Else
                obj.ReadOnly = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim DTA As New DseTrans_Cuti_PengajuanTableAdapters.tblmaster_workshift_itemallocateTableAdapter
        Dim TBL As New DseTrans_Cuti_Pengajuan.tblmaster_workshift_itemallocateDataTable
        'Dim row As DseTrans_Cuti_Pengajuan.tblmaster_workshift_itemallocateRow
        Dim CekDate As Date = Me.dtpStart.Value.Date
        Dim ListDate As New List(Of Date)
        Dim CekDate_Str As String = ""
        Dim ListTglAttendance As New List(Of Date)
        Dim DummyStartDate As Date
        Dim Dummy As String = ""
        Dim JmlAttendance As Integer = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'tidak boleh kosong
            KoleksiNamaData.Add("Employee code")
            KoleksiTextBox.Add(Me.TxtEmployeeCode)
            KoleksiNamaData.Add("Reason")
            KoleksiTextBox.Add(Me.TxtReason)
            If Me.ObjCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'harus ada alokasi workshift untuk employee yang bersangkutan
            ListDate.Clear()
            DTA.FillBy_PrimaryEmployee(TBL, Me.TxtPrimaryEmployee.Text)
            While CekDate <= Me.DtpEnd.Value.Date
                TBL.DefaultView.RowFilter = "[start date] <=#" + CekDate.ToString + _
                               "# AND [end date] >=#" + CekDate.ToString + "#"
                If TBL.DefaultView.Count = 0 Then
                    ListDate.Add(CekDate)
                End If
                CekDate += TimeSpan.FromDays(1)
            End While
            If ListDate.Count > 0 Then
                For Each CekDate In ListDate
                    If CekDate_Str = "" Then
                        CekDate_Str = Format(CekDate, "dd MMM yyyy")
                    Else
                        CekDate_Str = CekDate_Str + ", " + Format(CekDate, "dd MMM yyyy")
                    End If
                Next

                MsgBox("Please allocate workshift first for the following date: " + CekDate_Str, MsgBoxStyle.Information, "Information")
                Return False
            End If

            If Me.TxtPrimaryLeave.Text = "" Or Me.TxtPrimaryLeave.Text = "-1" Then
                MsgBox("Please specify leave data form this employee.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If Me.ObjCekData.CekBilanganPositif(Me.TxtDuration.Text) = False Then
                MsgBox("Please input positive number in duration", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If Me.Baru = True Then
                If Me.TxtStatus.Text.ToLower <> "waiting for approval" Then
                    MsgBox("Aplication status isn't waiting of approval, you can't save this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If Me.TxtPrimaryLeave.Text = -1 Or Me.cboLeave.Text = "" Then
                MsgBox("Please select leave code.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If Me.dtpStart.Value.Year <> Me.DtpEnd.Value.Year Then
                MsgBox("Start date and End date must have same year", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If dtpStart.Value > DtpEnd.Value Then
                MsgBox("Start date must be smaller than end date", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'tidak boleh diajukan pada pengajuan cuti lainnya
            If Me.ObjPengajuanCuti.CekDurasi(Me.TxtPrimaryEmployee.Text, Me.TxtPrimaryLeave.Text, dtpStart.Value, DtpEnd.Value, Me.Baru, Me.TxtPrimary.Text) = False Then
                Return False
            End If

            'jika sudah ada item attendance untuk salah satu tanggal ybs, maka tidak bisa simpan
            Dummy = ""
            DummyStartDate = Me.dtpStart.Value
            While DummyStartDate.Date <= Me.DtpEnd.Value.Date
                JmlAttendance = Me.ObjAttendanceItem.GetJml(Me.TxtPrimaryEmployee.Text, DummyStartDate)
                If JmlAttendance > 0 Then
                    If Dummy = "" Then
                        Dummy = Format(DummyStartDate, "dd MMM yyyy")
                    Else
                        Dummy = Dummy + vbNewLine + Format(DummyStartDate, "dd MMM yyyy")
                    End If
                End If
                DummyStartDate = DummyStartDate + TimeSpan.FromDays(1)
            End While
            'MsgBox("Tolong direpair")
            If Dummy <> "" Then
                MsgBox("There are attendance items for the following date: " + vbNewLine + Dummy, MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ListTglAttendance = Nothing
        End Try
    End Function

    Private Sub Aktifkandetail()
        Try
            Me.ObjCekData.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCekData.KendalikanControl(Me.GboApproval.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.TxtReason.Enabled = True
            Me.AktifForm = True
            Me.llbBalance.Enabled = True
            Me.llbNotes.Enabled = True
            Me.llbDuration.Enabled = True
            Me.CancelButton = Me.btnCancel
            If Me.PrivilegeUser = ClsTrans_Cuti_Pengajuan.EnumPrivilagePengajuanCuti.Concerned_Employee Then
                Me.TxtEmployeeCode.ReadOnly = True
                Me.BtnFindEmployeeCode.Enabled = False
                Me.SetEmployeeCode(Me.ObjUser.GetPrimary(Me.UserCode))
                Me.ObjPengajuanCuti.IsiCobLeave(Me)
            End If

            Me.DtpApp.Enabled = False
            Me.llbGoToApproval.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCekData.KendalikanControl(Me.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCekData.KendalikanControl(Me.GboApproval.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.AktifForm = False
            Me.TxtReason.Enabled = False
            Me.CancelButton = Me.btnExit
            Me.llbGoToApproval.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.ObjCekData.KosongkanControl(Me.gboDetail.Controls, True, True, True)
            Me.ObjCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, True)
            Me.ObjCekData.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, True)
            Me.ObjCekData.KosongkanControl(Me.GboApproval.Controls, True, True, True)
            Me.DseTrans_Cuti_Pengajuan.tblmaster_cuti.Rows.Clear()
            Me.TxtPrimary.Text = -1
            Me.TxtPrimaryEmployee.Text = -1
            'Me.TxtPrimaryLeave.Text = -1
            Me.TxtPrimaryPosition.Text = -1
            Me.TxtDuration.Text = 0
            Me.TxtBalance.Text = 0
            Me.TxtReason.Clear()
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            'If Me.PrivilegeUser = ClsTrans_Cuti_Pengajuan.EnumPrivilagePengajuanCuti.Concerned_Employee Then
            '    Me.TxtEmployeeCode.ReadOnly = True
            '    Me.BtnFindEmployeeCode.Enabled = False
            '    Me.SetEmployeeCode(Me.ObjUser.GetPrimary(Me.UserCode))
            '    Me.ObjPengajuanCuti.IsiCobLeave(Me.cboLeave, Me.TxtPrimaryEmployee.Text, Me.DtpDate.Value, Me.tbl, Me.TxtPrimaryLeave)
            'End If
            Me.ChkHalfDayEnd.Checked = False
            Me.ChkHalfDayStart.Checked = False
            Me.FrmStatus.TxtStatusNotes.Text = ""
            Me.TblList_StatusExplain.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.ObjPengajuanCuti
                Me.TxtPrimary.Text = .Primary
                Me.TxtDocNumber.Text = .NoTransaksi
                Me.DtpDate.Value = .Tanggal
                Me.TxtPrimaryEmployee.Text = .PrimaryKaryawan
                Me.TxtEmployeeCode.Text = Me.ObjKaryawan.GetCodeByPrimary(.PrimaryKaryawan)
                Me.TxtEmployeeName.Text = Me.ObjKaryawan.GetFullNameByPrimary(.PrimaryKaryawan)
                Me.TxtPrimaryPosition.Text = .Primaryjjob
                Me.TxtPosition.Text = Me.ObjJob.GetCode(.Primaryjjob)
                Me.ObjPengajuanCuti.IsiCobLeave(Me)
                Me.cboLeave.SelectedValue = .PrimaryCuti
                ''Me.TxtPrimaryLeave.Text = .PrimaryCuti
                If .Status = ClsTrans_Cuti_Pengajuan.EnumStatus.Waiting Then
                    Me.TxtStatus.Text = "Waiting For Approval"
                ElseIf .Status = ClsTrans_Cuti_Pengajuan.EnumStatus.App Then
                    Me.TxtStatus.Text = "Approved"
                ElseIf .Status = ClsTrans_Cuti_Pengajuan.EnumStatus.Rejected Then
                    Me.TxtStatus.Text = "Rejected"
                End If
                Me.FrmStatus.TxtStatusNotes.Text = .StatusNote
                Me.TxtAppBy.Text = .AppBy
                Me.DtpApp.Value = .AppDate
                Me.dtpStart.Value = .StartDate
                If .HalfDayStart = ClsTrans_Cuti_Pengajuan.EnumHalfDay.HalfDay Then
                    Me.ChkHalfDayStart.Checked = True
                Else
                    Me.ChkHalfDayStart.Checked = False
                End If
                Me.DtpEnd.Value = .EndDate
                If .HalfDayEnd = ClsTrans_Cuti_Pengajuan.EnumHalfDay.HalfDay Then
                    Me.ChkHalfDayEnd.Checked = True
                Else
                    Me.ChkHalfDayEnd.Checked = False
                End If
                Me.TxtDuration.Text = Me.ObjCekData.DigitGrouping(.Duration, 1)
                Me.TxtReason.Text = .Reason
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString

                Me.SetBalance()
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeObject()

        Try
            With Me.ObjPengajuanCuti
                .Primary = Me.TxtPrimary.Text
                .NoTransaksi = Me.TxtDocNumber.Text
                .Urutan = Me.TxtDocNumber.Text.Substring(0, 5)
                .Tanggal = Me.DtpDate.Value
                .PrimaryKaryawan = Me.TxtPrimaryEmployee.Text
                .Primaryjjob = Me.TxtPrimaryPosition.Text
                .PrimaryCuti = Me.TxtPrimaryLeave.Text
                If Me.TxtStatus.Text.ToLower = "waiting for approval" Then
                    .Status = ClsTrans_Cuti_Pengajuan.EnumStatus.Waiting
                ElseIf Me.TxtStatus.Text.ToLower = "approved" Then
                    .Status = ClsTrans_Cuti_Pengajuan.EnumStatus.App
                ElseIf Me.TxtStatus.Text.ToLower = "rejected" Then
                    .Status = ClsTrans_Cuti_Pengajuan.EnumStatus.Rejected
                End If
                .StatusNote = Me.FrmStatus.TxtStatusNotes.Text
                .AppBy = Me.TxtAppBy.Text
                .AppDate = Me.DtpDate.Value
                .StartDate = Me.dtpStart.Value
                .EndDate = Me.DtpEnd.Value
                If Me.ChkHalfDayStart.Checked = True Then
                    .HalfDayStart = ClsTrans_Cuti_Pengajuan.EnumHalfDay.HalfDay
                Else
                    .HalfDayStart = ClsTrans_Cuti_Pengajuan.EnumHalfDay.FullDay
                End If
                If Me.ChkHalfDayEnd.Checked = True Then
                    .HalfDayEnd = ClsTrans_Cuti_Pengajuan.EnumHalfDay.HalfDay
                Else
                    .HalfDayEnd = ClsTrans_Cuti_Pengajuan.EnumHalfDay.FullDay
                End If
                .Duration = Me.TxtDuration.Text
                .Reason = Me.TxtReason.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbBalance.LinkClicked
        'Dim DummyStartDate As Date
        Dim Dummy As String = ""
        Dim JmlAttendance As Integer = 0

        Try
            If Me.TxtReason.Enabled = False Then
                If Me.TxtPrimary.Text <> "-1" Then
                    If Me.TxtPrimaryLeave.Text = "" Or Me.TxtPrimaryLeave.Text = "-1" Then
                        Exit Sub
                    End If

                    Me.ij.mPrimaryEmployee = Me.TxtPrimaryEmployee.Text
                    Me.ij.mPrimaryLeave = Me.TxtPrimaryLeave.Text
                    Me.ij.mPeriod = Me.DtpDate.Value
                    Me.ij.mBaru = Me.Baru
                    Me.ij.mFrm = Me
                    Me.ij.ObjCd = Me.ObjCekData
                    Me.ij.Balance = 0
                    Me.ij.mPrimaryTransaksiPengajuan = Me.TxtPrimary.Text

                    Me.ij.ShowDialog()
                End If
            Else
                If Me.TxtPrimaryLeave.Text <> "-1" And Me.TxtPrimaryEmployee.Text <> "-1" Then
                    If Me.TxtPrimaryLeave.Text = "" Or Me.TxtPrimaryLeave.Text = "-1" Then
                        MsgBox("Please specify leave data form this employee.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If

                    Me.ij.mPrimaryEmployee = Me.TxtPrimaryEmployee.Text
                    Me.ij.mPrimaryLeave = Me.TxtPrimaryLeave.Text
                    Me.ij.mPeriod = Me.DtpDate.Value
                    Me.ij.mBaru = Me.Baru
                    Me.ij.mFrm = Me
                    Me.ij.ObjCd = Me.ObjCekData
                    Me.ij.Balance = 0
                    Me.ij.mPrimaryTransaksiPengajuan = Me.TxtPrimary.Text

                    Me.ij.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiCuti_Pengajuan_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Dim ObjUser As New Core_Module.ClsUser

        Try
            Me.SetBalance()
            If ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiCuti_Approval1") = False Then
                Me.llbGoToApproval.Visible = False
            Else
                Me.llbGoToApproval.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjUser = Nothing
        End Try
    End Sub

    Private Sub FrmTransaksiCuti_Pengajuan_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtDocNumber.Text <> "" And Me.TxtDocNumber.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObject()
                        If Baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjPengajuanCuti.SimpanBaru(Me.UserCode, Me.FrmExplain, Me.FrmExplain.dgvExplain, Me.TblList_StatusExplain, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjPengajuanCuti.SimpanPerubahan(Me.UserCode, Me.FrmExplain, Me.FrmExplain.dgvExplain, Me.TblList_StatusExplain, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            Else
                                e.Cancel = True
                            End If
                        End If
                    Else
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiCuti_Pengajuan_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim PrimaryUser As Double = -1

        Try
            Control.CheckForIllegalCrossThreadCalls = False
            PrimaryUser = Me.ObjUser.GetPrimary(Me.UserCode)
            Me.NonAktifkanDetail()
            'cari privilege nya
            If Me.ObjPengajuanCuti.GetPrivilegeUser(PrimaryUser) = _
            ClsTrans_Cuti_Pengajuan.EnumPrivilagePengajuanCuti.Concerned_Employee Then
                Me.PrivilegeUser = ClsTrans_Cuti_Pengajuan.EnumPrivilagePengajuanCuti.Concerned_Employee
                Me.SetEmployeeCode(PrimaryUser)
            Else
                Me.PrivilegeUser = ClsTrans_Cuti_Pengajuan.EnumPrivilagePengajuanCuti.All_Employee
            End If
            Me.FrmExplain.BackColorAll = Me.BackColorAll
            Me.ShowIcon = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.Baru = True
            Me.KosongkanField()
            Me.Aktifkandetail()
            Me.TxtDocNumber.Text = Me.ObjPengajuanCuti.GetNoBaru(Me.DtpDate.Value)
            Me.DtpDate.Focus()
            Me.TxtStatus.Text = "Waiting For Approval"
            Me.TxtAppBy.Text = "-"
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            'cek data kosong
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObject()
                hasil = MsgBox("Save data ?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Baru = True Then
                        Windows.Forms.Cursor.Current = Windows.Forms.Cursors.WaitCursor
                        If Me.ObjPengajuanCuti.SimpanBaru(Me.UserCode, Me.FrmExplain, Me.FrmExplain.dgvExplain, Me.TblList_StatusExplain, Me) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                        End If
                    Else
                        If Me.ObjPengajuanCuti.SimpanPerubahan(Me.UserCode, Me.FrmExplain, Me.FrmExplain.dgvExplain, Me.TblList_StatusExplain, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReason_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReason.GotFocus
        Try
            If Me.btnOK.Enabled = True Then
                If Me.TxtReason.ReadOnly = False Then
                    Me.TxtReason.BackColor = Me.BackColorAll
                Else
                    Me.TxtReason.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReason_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtReason.KeyPress
        Me.BlokSimbol(Me.TxtReason, e)
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtDocNumber.Text <> "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Baru = True Then
                        Me.NonAktifkanDetail()
                        Me.KosongkanField()
                    Else
                        Me.NonAktifkanDetail()
                        Me.ObjPengajuanCuti.CariDetail(Me.TxtPrimary.Text, Me.FrmExplain, Me.FrmExplain.dgvExplain, Me.TblList_StatusExplain, Me)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.NonAktifkanDetail()
                Me.KosongkanField()
            End If
            Me.btnNew.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmTransaksiCuti_Pengajuan_Cari
        Try
            hj.BackColorAll = Me.BackColorAll
            If Me.PrivilegeUser = ClsTrans_Cuti_Pengajuan.EnumPrivilagePengajuanCuti.All_Employee Then
                hj.AllFilterData = True
            ElseIf Me.PrivilegeUser = ClsTrans_Cuti_Pengajuan.EnumPrivilagePengajuanCuti.Concerned_Employee Then
                hj.AllFilterData = False
                hj.PrimaryKaryawan = Me.TxtPrimaryEmployee.Text
            End If
            hj.ShowDialog()
            If Not hj.Primary = -1 Then
                If Me.ObjPengajuanCuti.CariDetail(hj.Primary, Me.FrmExplain, Me.FrmExplain.dgvExplain, Me.TblList_StatusExplain, Me) = True Then
                    Me.SalinKeField()
                    Me.Baru = False
                    Me.btnEdit.Focus()
                    Me.NonAktifkanDetail()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            Me.Baru = False
            If Me.TxtDocNumber.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtDocNumber.Text <> "" Then
                If Me.ObjPengajuanCuti.ApakahBolehEdit(Me.UserCode) = True Then
                    Me.Baru = False
                    Me.ObjPengajuanCuti.CariDetail(Me.TxtPrimary.Text, Me.FrmExplain, Me.FrmExplain.dgvExplain, Me.TblList_StatusExplain, Me)
                    Me.SalinKeField()
                    Me.Aktifkandetail()
                    Me.TxtDocNumber.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtDocNumber.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtDocNumber.Text <> "" Then
                If Me.ObjPengajuanCuti.ApakahBolehHapus(Me.UserCode) Then
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ObjPengajuanCuti.Hapus(Me.TxtPrimary.Text, Me.UserCode) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub TxtEmployeeCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmployeeCode.GotFocus
        Me.Buffer = Me.TxtPrimaryEmployee.Text
        If Me.TxtEmployeeCode.ReadOnly = False Then
            Me.TxtEmployeeCode.BackColor = Me.BackColorAll
        Else
            Me.TxtEmployeeCode.BackColor = Me.TxtReadOnly.BackColor
        End If
    End Sub

    Private Sub TxtEmployeeCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtEmployeeCode.KeyDown
        Try
            Me.Buffer = Me.TxtPrimaryEmployee.Text
            If e.KeyValue = 13 Then
                If Me.ObjKaryawan.ApakahSudahAda(Me.TxtEmployeeCode.Text) = True Then
                    Me.TxtPrimaryEmployee.Text = Me.ObjKaryawan.GetPrimaryByCode(Me.TxtEmployeeCode.Text)
                    Me.TxtEmployeeCode.Text = Me.ObjKaryawan.GetCodeByPrimary(Me.TxtPrimaryEmployee.Text)
                    Me.TxtEmployeeName.Text = Me.ObjKaryawan.GetFullNameByPrimary(Me.TxtPrimaryEmployee.Text)
                    Me.TxtPrimaryPosition.Text = Me.ObjKaryawan.GetJobByPrimary(Me.TxtPrimaryEmployee.Text)
                    Me.TxtPosition.Text = Me.ObjJob.GetCode(Me.TxtPrimaryPosition.Text)
                    Me.cboLeave.Focus()
                    If dtpStart.Value > DtpEnd.Value Then
                        Me.TxtDuration.Text = 0
                    Else
                        Me.ObjPengajuanCuti.HitungDurasi(Me.TxtPrimaryEmployee.Text, Me.dtpStart.Value, Me.DtpEnd.Value, _
                        Me.TxtDuration, Me.ChkHalfDayStart, Me.ChkHalfDayEnd, Me.FrmExplain, Me.TblList_StatusExplain, Me)
                    End If
                Else
                    Me.BtnFindEmployeeCode.PerformClick()
                End If
                'If Me.Buffer <> Me.TxtPrimaryEmployee.Text Then
                Me.ObjPengajuanCuti.IsiCobLeave(Me)
                'End If
                Me.SetBalance()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtEmployeeCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmployeeCode.Leave
        Dim TglAttendanceMax As Date = Today

        Try
            If Me.TxtEmployeeCode.Text <> "" Then
                If Me.ObjKaryawan.ApakahSudahAda(Me.TxtEmployeeCode.Text) = True Then
                    Me.TxtPrimaryEmployee.Text = Me.ObjKaryawan.GetPrimaryByCode(Me.TxtEmployeeCode.Text)
                    Me.TxtEmployeeCode.Text = Me.ObjKaryawan.GetCodeByPrimary(Me.TxtPrimaryEmployee.Text)
                    Me.TxtEmployeeName.Text = Me.ObjKaryawan.GetFullNameByPrimary(Me.TxtPrimaryEmployee.Text)
                    Me.TxtPrimaryPosition.Text = Me.ObjKaryawan.GetJobByPrimary(Me.TxtPrimaryEmployee.Text)
                    Me.TxtPosition.Text = Me.ObjJob.GetCode(Me.TxtPrimaryPosition.Text)
                    If dtpStart.Value > DtpEnd.Value Then
                        Me.TxtDuration.Text = 0
                    Else
                        Me.ObjPengajuanCuti.HitungDurasi(Me.TxtPrimaryEmployee.Text, Me.dtpStart.Value, Me.DtpEnd.Value, _
                        Me.TxtDuration, Me.ChkHalfDayStart, Me.ChkHalfDayEnd, Me.FrmExplain, Me.TblList_StatusExplain, Me)
                    End If
                Else
                    Me.BtnFindEmployeeCode.PerformClick()
                End If
            Else
                Me.TxtPrimaryEmployee.Text = -1
                Me.TxtEmployeeCode.Text = ""
                Me.TxtEmployeeName.Text = ""
                Me.TxtPrimaryPosition.Text = -1
                Me.TxtPosition.Text = ""
                If dtpStart.Value > DtpEnd.Value Then
                    Me.TxtDuration.Text = 0
                Else
                    Me.ObjPengajuanCuti.HitungDurasi(Me.TxtPrimaryEmployee.Text, Me.dtpStart.Value, Me.DtpEnd.Value, _
                    Me.TxtDuration, Me.ChkHalfDayStart, Me.ChkHalfDayEnd, Me.FrmExplain, Me.TblList_StatusExplain, Me)
                End If
            End If
            If Me.Buffer <> Me.TxtPrimaryEmployee.Text Then
                Me.ObjPengajuanCuti.IsiCobLeave(Me)
            End If

            'jika ada yang memiliki attendance, maka geser langsung ke tanggal attendance selanjutnya
            TglAttendanceMax = Me.ObjAttendanceItem.GetMaxAttendanceDate(Me.TxtPrimaryEmployee.Text)
            If Me.DtpEnd.Value <= TglAttendanceMax Or Me.dtpStart.Value <= TglAttendanceMax Then
                Me.dtpStart.Value = TglAttendanceMax + TimeSpan.FromDays(1)
                Me.DtpEnd.Value = TglAttendanceMax + TimeSpan.FromDays(1)
                Me.ObjPengajuanCuti.HitungDurasi(Me.TxtPrimaryEmployee.Text, Me.dtpStart.Value, Me.DtpEnd.Value, _
                Me.TxtDuration, Me.ChkHalfDayStart, Me.ChkHalfDayEnd, Me.FrmExplain, Me.TblList_StatusExplain, Me)
            End If

            Me.SetBalance()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindEmployeeCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindEmployeeCode.Click
        Dim FrmCariEmployee As New Core_Module.FrmMasterKaryawanCari
        Try
            Me.Buffer = Me.TxtPrimaryEmployee.Text
            FrmCariEmployee.BackColorAll = Me.BackColorAll
            If Me.TxtPrimaryEmployee.Text = -1 And Me.TxtEmployeeCode.Text = "" Then
                FrmCariEmployee.ShowDialog()
            Else
                FrmCariEmployee.BackColorAll = Me.BackColorAll
                FrmCariEmployee.txtCari.Text = Me.TxtEmployeeCode.Text
                FrmCariEmployee.ShowDialog()
            End If
            If CDbl(FrmCariEmployee.PrimaryEmploye) > 0 Then
                Me.TxtPrimaryEmployee.Text = FrmCariEmployee.PrimaryEmploye
                Me.TxtEmployeeCode.Text = Me.ObjKaryawan.GetCodeByPrimary(FrmCariEmployee.PrimaryEmploye)
                Me.TxtEmployeeName.Text = Me.ObjKaryawan.GetFullNameByPrimary(FrmCariEmployee.PrimaryEmploye)
                Me.TxtPrimaryPosition.Text = Me.ObjKaryawan.GetJobByPrimary(Me.TxtPrimaryEmployee.Text)
                Me.TxtPosition.Text = Me.ObjJob.GetCode(Me.TxtPrimaryPosition.Text)
                Me.cboLeave.Focus()
            Else
                Me.TxtPrimaryEmployee.Text = -1
                Me.TxtEmployeeCode.Text = ""
                Me.TxtEmployeeName.Text = ""
                Me.TxtPrimaryPosition.Text = -1
                Me.TxtPosition.Text = ""
                Me.cboLeave.Focus()
            End If
            'If Me.Buffer <> Me.TxtPrimaryEmployee.Text Then
            Me.ObjPengajuanCuti.IsiCobLeave(Me)
            'End If
            If dtpStart.Value > DtpEnd.Value Then
                Me.TxtDuration.Text = 0
            Else
                Me.ObjPengajuanCuti.HitungDurasi(Me.TxtPrimaryEmployee.Text, Me.dtpStart.Value, Me.DtpEnd.Value, _
                Me.TxtDuration, Me.ChkHalfDayStart, Me.ChkHalfDayEnd, Me.FrmExplain, Me.TblList_StatusExplain, Me)
            End If
            Me.SetBalance()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmCariEmployee.Dispose()
        End Try
    End Sub

    Private Sub DtpDate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpDate.Enter
        Me.BufferDate = Me.DtpDate.Value.Year.ToString
    End Sub

    Private Sub DtpDate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpDate.Leave
        Try
            If Me.DtpDate.Value.Year.ToString <> Me.BufferDate Then
                If Me.TxtPrimaryEmployee.Text <> "-1" Then
                    Me.ObjPengajuanCuti.IsiCobLeave(Me)
                End If
            End If
            Me.TxtDocNumber.Text = Me.ObjPengajuanCuti.GetNoBaru(Me.DtpDate.Value)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbNotes.LinkClicked
        Try
            If Me.TxtReason.Enabled = False Then
                If Me.TxtPrimary.Text <> "-1" Then
                    Me.FrmStatus.TxtStatusNotes.Dock = DockStyle.Fill
                    Me.FrmStatus.btnOK.SendToBack()
                    Me.FrmStatus.btnCancel.SendToBack()
                    Me.FrmStatus.btnOK.TabStop = False
                    Me.FrmStatus.btnCancel.TabStop = False
                    Me.FrmStatus.MaximizeBox = False
                    Me.FrmStatus.MinimizeBox = False
                    Me.FrmStatus.TxtStatusNotes.ReadOnly = True
                    Me.FrmStatus.ShowDialog()
                End If
            Else
                Me.FrmStatus.TxtStatusNotes.Dock = DockStyle.Fill
                Me.FrmStatus.btnOK.SendToBack()
                Me.FrmStatus.btnCancel.SendToBack()
                Me.FrmStatus.btnOK.TabStop = False
                Me.FrmStatus.btnCancel.TabStop = False
                Me.FrmStatus.MaximizeBox = False
                Me.FrmStatus.MinimizeBox = False
                Me.FrmStatus.TxtStatusNotes.ReadOnly = True
                Me.FrmStatus.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel3_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbDuration.LinkClicked
        Try
            ObjItem.Add_To_List(Me.TblList_StatusExplain, Me.FrmExplain, Me)
            If Me.btnOK.Enabled = False Then
                Me.FrmExplain.dgvExplain.ReadOnly = True
            Else
                Me.FrmExplain.dgvExplain.ReadOnly = False
            End If
            If Me.TxtReason.Enabled = False Then
                If Me.TxtPrimary.Text <> "-1" Then
                    If Me.TxtPrimaryLeave.Text = "" Or Me.TxtPrimaryLeave.Text = "-1" Then
                        Exit Sub
                    End If
                    Me.ObjItem.IsiDGVItem(FrmExplain.dgvExplain, Me.TxtPrimary.Text, Me.FrmExplain, Me.TblList_StatusExplain, Me)
                    Me.FrmExplain.ShowDialog()
                End If
            Else
                If Me.TxtPrimaryLeave.Text = "" Or Me.TxtPrimaryLeave.Text = "-1" Then
                    MsgBox("Please specify leave data form this employee.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                If CDbl(Me.TxtDuration.Text) <> 0 Or Me.TblList_StatusExplain.Rows.Count > 0 Then
                    If Me.ChkHalfDayStart.Checked = True Then Me.FrmExplain.Half1 = True
                    If Me.ChkHalfDayEnd.Checked = True Then Me.FrmExplain.Half2 = True
                    Me.FrmExplain.ShowDialog()
                    ObjItem.Add_To_List(Me.TblList_StatusExplain, Me.FrmExplain, Me)
                    ObjPengajuanCuti.HitungDurasi(Me.TxtPrimaryEmployee.Text, Me.dtpStart.Value, Me.DtpEnd.Value, Me.TxtDuration, Me.ChkHalfDayStart, Me.ChkHalfDayEnd, _
                     Me.FrmExplain, Me.TblList_StatusExplain, Me)
                    Me.SetBalance()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpEnd_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpEnd.Enter
        Try
            Me.BufferDate = Me.DtpEnd.Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpEnd_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpEnd.Leave
        Try
            Me.BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpStart_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpStart.Enter
        Try
            Me.BufferDate = Me.dtpStart.Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpStart_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpStart.Leave
        Try
            Me.BackgroundWorker1.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkHalfDayEnd_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkHalfDayEnd.CheckedChanged
        Try
            If dtpStart.Value > DtpEnd.Value Then
                Me.TxtDuration.Text = 0
            Else
                Me.ObjPengajuanCuti.HitungDurasi(Me.TxtPrimaryEmployee.Text, Me.dtpStart.Value, Me.DtpEnd.Value, _
                Me.TxtDuration, Me.ChkHalfDayStart, Me.ChkHalfDayEnd, Me.FrmExplain, Me.TblList_StatusExplain, Me)
            End If
            If Me.TxtReason.Enabled = True Then
                Me.SetBalance()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkHalfDayStart_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkHalfDayStart.CheckedChanged
        Try
            If dtpStart.Value > DtpEnd.Value Then
                Me.TxtDuration.Text = 0
            Else
                Me.ObjPengajuanCuti.HitungDurasi(Me.TxtPrimaryEmployee.Text, Me.dtpStart.Value, Me.DtpEnd.Value, _
                Me.TxtDuration, Me.ChkHalfDayStart, Me.ChkHalfDayEnd, Me.FrmExplain, Me.TblList_StatusExplain, Me)
            End If
            If Me.TxtReason.Enabled = True Then
                Me.SetBalance()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboLeave_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboLeave.SelectedIndexChanged
        Dim dta As New DseTrans_Cuti_PengajuanTableAdapters.tblmaster_cutiTableAdapter
        'Dim StrCode() As String
        Try
            If Me.cboLeave.SelectedIndex < 0 Then
                'bersihkan objek
                Me.TxtPrimaryLeave.Text = "-1"
                'Else
                '    StrCode = Me.cboLeave.Text.Split("-")
                '    Me.TxtPrimaryLeave.Text = Me.ObjCuti.GetPrimaryByCode(StrCode(0).Trim)
            End If
            If Me.TxtPrimaryEmployee.Text <> "" Then
                Me.ObjPengajuanCuti.HitungDurasi(Me.TxtPrimaryEmployee.Text, Me.dtpStart.Value, Me.DtpEnd.Value, _
                            Me.TxtDuration, Me.ChkHalfDayStart, Me.ChkHalfDayEnd, Me.FrmExplain, Me.TblList_StatusExplain, Me)

                If Me.TxtReason.Enabled = True Then
                    Me.SetBalance()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtEmployeeCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmployeeCode.LostFocus
        If Me.TxtEmployeeCode.ReadOnly = False Then
            Me.TxtEmployeeCode.BackColor = Drawing.Color.White
        Else
            Me.TxtEmployeeCode.BackColor = Me.TxtReadOnly.BackColor
        End If
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(10)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Dim Dummy As String = ""
        Dim DummyStartDate As Date = Today
        Dim JmlAttendance As Double = 0

        Try
            If dtpStart.Value.Date > DtpEnd.Value.Date Then
                'MsgBox("Start date must be smaller than end date", MsgBoxStyle.Information, "Information")
                'Me.DtpEnd.Value = Me.BufferDate
            Else
                ' ''jika sudah ada item attendance untuk salah satu tanggal ybs, maka tidak bisa simpan
                ''Dummy = ""
                ''DummyStartDate = Me.dtpStart.Value
                ''While DummyStartDate.Date <= Me.DtpEnd.Value.Date
                ''    JmlAttendance = Me.ObjAttendanceItem.GetJml(Me.TxtPrimaryEmployee.Text, DummyStartDate)
                ''    If JmlAttendance > 0 Then
                ''        If Dummy = "" Then
                ''            Dummy = Format(DummyStartDate, "dd MMM yyyy")
                ''        Else
                ''            Dummy = Dummy + ", " + Format(DummyStartDate, "dd MMM yyyy")
                ''        End If
                ''    End If
                ''    DummyStartDate = DummyStartDate + TimeSpan.FromDays(1)
                ''End While
                ''If Dummy <> "" Then
                ''    MsgBox("There are attendance items for the following date: " + Dummy, MsgBoxStyle.Information, "Information")
                ''    Me.DtpEnd.Value = Me.BufferDate
                ''    Exit Sub
                ''End If

                Me.ObjPengajuanCuti.HitungDurasi(Me.TxtPrimaryEmployee.Text, Me.dtpStart.Value, Me.DtpEnd.Value, _
                Me.TxtDuration, Me.ChkHalfDayStart, Me.ChkHalfDayEnd, Me.FrmExplain, Me.TblList_StatusExplain, Me)
                If Me.TxtReason.Enabled = True Then
                    Me.SetBalance()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(10)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Dim Dummy As String = ""
        Dim DummyStartDate As Date = Today
        Dim JmlAttendance As Double = 0

        Try
            If dtpStart.Value.Date > DtpEnd.Value.Date Then
                'MsgBox("Start date must be smaller than end date", MsgBoxStyle.Information, "Information")
                'Me.dtpStart.Value = Me.BufferDate
            Else
                ' ''jika sudah ada item attendance untuk salah satu tanggal ybs, maka tidak bisa simpan
                ''Dummy = ""
                ''DummyStartDate = Me.dtpStart.Value
                ''While DummyStartDate.Date <= Me.DtpEnd.Value.Date
                ''    JmlAttendance = Me.ObjAttendanceItem.GetJml(Me.TxtPrimaryEmployee.Text, DummyStartDate)
                ''    If JmlAttendance > 0 Then
                ''        If Dummy = "" Then
                ''            Dummy = Format(DummyStartDate, "dd MMM yyyy")
                ''        Else
                ''            Dummy = Dummy + ", " + Format(DummyStartDate, "dd MMM yyyy")
                ''        End If
                ''    End If
                ''    DummyStartDate = DummyStartDate + TimeSpan.FromDays(1)
                ''End While
                ''If Dummy <> "" Then
                ''    MsgBox("There are attendance items for the following date: " + Dummy, MsgBoxStyle.Information, "Information")
                ''    Me.dtpStart.Value = Me.BufferDate
                ''    Exit Sub
                ''End If

                Me.ObjPengajuanCuti.HitungDurasi(Me.TxtPrimaryEmployee.Text, Me.dtpStart.Value, Me.DtpEnd.Value, _
                Me.TxtDuration, Me.ChkHalfDayStart, Me.ChkHalfDayEnd, Me.FrmExplain, Me.TblList_StatusExplain, Me)
                If Me.TxtReason.Enabled = True Then
                    Me.SetBalance()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReason_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReason.LostFocus
        Try
            If Me.btnOK.Enabled = True Then
                If Me.TxtReason.ReadOnly = False Then
                    Me.TxtReason.BackColor = Me.TxtPrimary.BackColor
                Else
                    Me.TxtReason.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbGoToApproval_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbGoToApproval.LinkClicked
        Try
            Dim ij As HR.FrmTransaksiCuti_Approval
            Dim j As Windows.Forms.Form

            For Each j In Me.ParentForm.MdiChildren
                ij = TryCast(j, HR.FrmTransaksiCuti_Approval)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmTransaksiCuti_Approval
            ' Increment the caption counter.
            i.UserCode = Me.UserCode
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me.ParentForm
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class