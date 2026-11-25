Option Explicit On
Option Strict Off
Imports System.Windows.Forms
Imports System.Drawing
Public Class FrmTrans_Attendance
    Public IsImportFromMachine As Boolean = False
    Public IsContinueFromPrevAtt As Boolean = False
    Public FingerPrintID As String = ""
    Public IsOverrideClockIn As Boolean = False
    Public DoesNotClockIn As Boolean = False
    Public IsOverrideClockOut As Boolean = False
    Public DoesNotClockOut As Boolean = False

    Dim BufferEmployee As Double = -1

    Public IsOverrideEarlyOTRule, IsCalculateEarlyOT, IsOverrideLateRule, IsCalculateLate, IsOverrideEarlyOutRule, IsCalculateEarlyOut, IsOverrideOTRule, IsCalculateOT As Boolean
    Public IsCalculateEarlyOTWS, IsCalculateLateWS, IsCalculateEarlyOutWS, IsCalculateOTWS As Boolean
    Public CalculateAfterEarlyOTM, CalculateAfterLateM, CalculateAfterEarlyOutM, CalculateAfterOTM, CalculateAfterEarlyOTWS, CalculateAfterLateWS, CalculateAfterEarlyOutWS, CalculateAfterOTWS As Double
    Public RoundingTypeEarlyOTM, RoundingTypeLateM, RoundingTypeEarlyOutM, RoundingTypeOTM, RoundingTypeEarlyOTWS, RoundingTypeLateWS, RoundingTypeEarlyOutWS, RoundingTypeOTWS As Integer
    Public RoundingValueEarlyOTM, RoundingValueLateM, RoundingValueEarlyOutM, RoundingValueOTM, RoundingValueEarlyOTWS, RoundingValueLateWS, RoundingValueEarlyOutWS, RoundingValueOTWS As Decimal
    Public MaxValueEarlyOTM, MaxValueLateM, MaxValueEarlyOutM, MaxValueOTM, MaxValueEarlyOTWS, MaxValueLateWS, MaxValueEarlyOutWS, MaxValueOTWS As Decimal

    Public CalculateEarlyOTTolerance, CalculateLateTolerance, CalculateEarlyOutTolerance, CalculateOTTolerance As Integer
    Public CalculateEarlyOTToleranceWS, CalculateLateToleranceWS, CalculateEarlyOutToleranceWS, CalculateOTToleranceWS As Integer

    Public CalculateEarlyOTOnWeekend, CalculateLateOnWeekend, CalculateEarlyOutOnWeekend, CalculateOTOnWeekend As Integer
    Public CalculateEarlyOTOnWeekendWS, CalculateLateOnWeekendWS, CalculateEarlyOutOnWeekendWS, CalculateOTOnWeekendWS As Integer
    Public CalculateEarlyOTOnHoliday, CalculateLateOnHoliday, CalculateEarlyOutOnHoliday, CalculateOTOnHoliday As Integer
    Public CalculateEarlyOTOnHolidayWS, CalculateLateOnHolidayWS, CalculateEarlyOutOnHolidayWS, CalculateOTOnHolidayWS As Integer
    Dim tblHoliday As New DseTrans_Attendance.tblmaster_workshift_itemholiday1DataTable
    Dim rowHoliday As DseTrans_Attendance.tblmaster_workshift_itemholiday1Row

    'aul tambahan
    Public JoinWithAttPrev As Boolean = False
    Public DtpStartAttPrev As Date = "1/1/1950"
    Public DtpToAttPrev As Date = "1/1/1950"
    Public ApakahSudahReload As Boolean = False

    Public PrimaryAtt_FromQuick As Double = -1
    Public ShowAtt_FromQuick As Boolean = False

    'aul tambahan overtime order
    Public TblOT_Order As New DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_1DataTable
    Dim ConfirmBox As Boolean = True

    Enum EnumCalculationType As Integer
        WORKDAY
        WEEKEND
        HOLIDAY
    End Enum

    Enum EnumCalculateRange As Integer
        IN_RANGE
        OUT_RANGE
        BOTH
    End Enum

    Public StartingWS As Windows.Forms.FormWindowState
    Dim DurationMin As Integer = 15 ''default minimal duration
    Dim WarningColor As Color = Color.FromArgb(255, 255, 228, 225)
    Dim ObjWS As New Core_Module.ClsMasterWorkShift

    Public UserCode As String = ""
    Public BackColorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll
    Public Baru As Boolean = True
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjCd As New Core_Module.ClsCekData
    Public ObjAtt As New ClsTrans_Attendance
    Dim ObjEmp As New Core_Module.ClsMasterKaryawan
    Dim ObjJob As New Core_Module.ClsMasterJobPosition
    Dim ObjLA As New ClsTrans_Cuti_Pengajuan
    Dim ObjWS_Item As New Core_Module.ClsMasterWorkShift_Item
    Dim BufferDate As Date = #1/1/1950#
    Dim Buffer As String = ""
    Dim BufferStatus As ClsTrans_Attendance.EnumStatus
    Dim BufferTransAttendanceDate As Date = #1/1/1950#
    Dim ObjSettingsUser As New Core_Module.ClsTransaksiSettingUser

    Dim vCell As System.Windows.Forms.DataGridViewCell
    Dim vCboDgv As Windows.Forms.DataGridViewComboBoxCell
    Dim vVal As String

    Dim RoundingUnit As String = ""
    Dim RoundingType As String = ""
    Dim DailyMaxOT As String = ""
    Dim PeriodMaxOt As String = ""
    Dim RoundingUnit_EarlyOt As String = ""
    Dim RoundingType_EarlyOt As String = ""
    Dim DailyMaxOT_EarlyOt As String = ""
    Dim PeriodMaxOt_EarlyOt As String = ""
    Dim RoundingTypeLateIn As String = ""
    Dim RoundingUnitLateIn As String = ""
    Dim RoundingTypeEarlyOut As String = ""
    Dim RoundingUnitEarlyOut As String = ""

    Public DariQuick As Boolean = False

    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        Me.InitSetting()
        ' Add any initialization after the InitializeComponent() call.
    End Sub

    Private Sub InitSetting()
        Try
            'Me.ColCalculateAfterEarlyOT.Visible = True
            'Me.ColCalculateAfterEarlyLate.Visible = True
            'Me.ColCalculateAfterEarlyOut.Visible = True
            'Me.ColCalculateAfterOT.Visible = True
            'Me.ColRoundingTypeEarlyOT.Visible = True
            'Me.ColRoundingTypeLate.Visible = True
            'Me.ColRoundingTypeEarlyOut.Visible = True
            'Me.ColRoundingTypeOT.Visible = True
            'Me.ColRoundingValueEarlyOT.Visible = True
            'Me.ColRoundingValueLate.Visible = True
            'Me.ColRoundingValueEarlyOut.Visible = True
            'Me.ColRoundingValueOT.Visible = True
            'Me.ColMaxValueEarlyOT.Visible = True
            'Me.ColMaxValueLate.Visible = True
            'Me.ColMaxValueEarlyOut.Visible = True
            'Me.ColMaxValueOT.Visible = True
            'ignore breakdata
            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.IgnoreBreakData) = "1" Then
                Me.CheckBox1.Checked = True
            ElseIf Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.IgnoreBreakData) = "0" Then
                Me.CheckBox1.Checked = False
            End If
            'ignore early in
            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.IgnoreEarlyIn) = "1" Then
                Me.CheckBox3.Checked = True
            ElseIf Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.IgnoreEarlyIn) = "0" Then
                Me.CheckBox3.Checked = False
            End If
            'ignore early out
            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.IgnoreEarlyOut) = "1" Then
                Me.CheckBox4.Checked = True
            ElseIf Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.IgnoreEarlyOut) = "0" Then
                Me.CheckBox4.Checked = False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    'design notes:
    '1. semua pukul dan durasi dituliskan dalam format 00.00 dan rata kanan
    '2. semua baris holiday, diberikan warna merah
    '3. update coding untuk workshift, pada bagian item+item holiday, penambahan kolom, schedule break out dan schedule break in
    '4. cek batasan leave aplication: (1) approvalnya tidak bisa dicabut kalo sudah ada cuti aktualnya (2) leave aplication tidak boleh dihapuskan 
    '   jika sudah aktual terjadi cuti
    '5. jika ada leave aplication yang statusnya belum jelas (terima/tolak) untuk karyawan yang akan diinput datanya - pada periode ybs
    '   maka tidak bisa simpan data
    '6. 
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Try
            If Me.CheckBox1.Checked = True Then
                Me.ColActualBreakIn.Visible = False
                Me.ColActualBreakOut.Visible = False
                Me.ColSchBreakIn.Visible = False
                Me.ColSchBreakOut.Visible = False
                Me.ColEarlyBreakOutDuration.Visible = False
                Me.ColLateBreakInDuration.Visible = False
            Else
                Me.ColActualBreakIn.Visible = True
                Me.ColActualBreakOut.Visible = True
                Me.ColSchBreakIn.Visible = True
                Me.ColSchBreakOut.Visible = True
                Me.ColEarlyBreakOutDuration.Visible = True
                Me.ColLateBreakInDuration.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        Try
            If Me.CheckBox2.Checked = True Then
                Me.ColLeaveCode.Visible = False
                Me.ColLeaveDesc.Visible = False
                Me.ColLeaveDoc.Visible = False
            Else
                Me.ColLeaveCode.Visible = True
                Me.ColLeaveDesc.Visible = True
                Me.ColLeaveDoc.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        Try
            If Me.CheckBox3.Checked = True Then
                Me.ColEarlyInDuration.Visible = False
                Me.ColEarlyOTDuration.Visible = False
            Else
                Me.ColEarlyInDuration.Visible = True
                Me.ColEarlyOTDuration.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        Try
            If Me.CheckBox4.Checked = True Then
                Me.ColEarlyOutDuration.Visible = False
            Else
                Me.ColEarlyOutDuration.Visible = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged
        Try
            If Me.CheckBox5.Checked = True Then
                Me.ColStatusRemark.Visible = False
            Else
                Me.ColStatusRemark.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CheckBox6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox6.CheckedChanged
        Try
            If Me.CheckBox6.Checked = True Then
                Me.ColOTDuration.Visible = False
            Else
                Me.ColOTDuration.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CheckBox7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox7.CheckedChanged
        Try
            If Me.CheckBox7.Checked = True Then
                Me.ColWorkshift.Visible = False
            Else
                Me.ColWorkshift.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Enum EnumRoundingType
        No
        Down
        Up
        Middle
    End Enum

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim RowDgv As Windows.Forms.DataGridViewRow
        Dim CboDgv As Windows.Forms.DataGridViewComboBoxCell
        Dim hasil As DialogResult
        Try
            'tidak boleh kosong
            KoleksiNamaData.Add("Employee code")
            KoleksiTextBox.Add(Me.TxtEmployeeCode)

            If Me.ObjCd.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            Dim ObjLeave As New Core_Module.ClsMasterCuti
            If ObjLeave.IsAlreadyDefaultLeave(Me.ObjCd.GetFisrtDateOfTheYear(Me.DtpDate.Value)) = False Then
                MsgBox("There is no default leave settings in this period, please define first.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'tidak boleh double
            If Me.Baru = True Then
                If Me.ObjAtt.ApakahSudahAda(Me.TxtDocNumber.Text) = True Then
                    MsgBox("Duplication in code, please change your code", MsgBoxStyle.Information, "Information")
                    Me.TxtDocNumber.Focus()
                    Return False
                End If
            Else
                Dim CodeLama As String = Me.ObjAtt.GetNoTransByPrimary(Me.TxtPrimary.Text)
                If CodeLama <> Me.TxtDocNumber.Text Then
                    If Me.ObjAtt.ApakahSudahAda(Me.TxtDocNumber.Text) = True Then
                        MsgBox("Duplication in document nomor, please change your code", MsgBoxStyle.Information, "Information")
                        Me.TxtDocNumber.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses(edit)
                Dim OldTimeStamp As Date = Me.ObjAtt.GetTimeStampByPrimary(Me.TxtPrimary.Text)
                If Me.ObjAtt.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjAtt.CariDetail(Me.TxtPrimary.Text, Me, Me.TblOT_Order, Me.JoinWithAttPrev)
                        'Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            If Me.ObjEmp.ApakahAktif(Me.TxtPrimaryEmployee.Text) = False Then
                MsgBox("Inactive employee, you cannot save this data.", MsgBoxStyle.Information, "Information")
                Me.TxtEmployeeCode.Focus()
                Return False
            End If


            'tahun harus sama
            'If Me.dtpStart.Value.Year <> Me.DtpEnd.Value.Year Then
            '    MsgBox("Start date year and end date year must be the same", MsgBoxStyle.Information, "Information")
            '    Return False
            'End If

            'minimal harus ada 1 item
            If Me.dgvItem.RowCount < 1 Then
                MsgBox("Please input attendance item", MsgBoxStyle.Information, "Information")
                Return False
            End If
            'setiap status yang setengah hari durasi nya tidak boleh lebih dari 5 jam
            Dim workday As ClsTrans_Attendance.EnumWorkday
            Dim Durasi As ClsTrans_Attendance.EnumDurasi
            Dim Status As ClsTrans_Attendance.EnumStatus

            Dim ActIn As Double
            Dim ActBreakOut As Double
            Dim ActBreakIn As Double
            Dim ActOut As Double
            'Dim BatasDurasi As Integer
            Dim JmlOt As Double = 0
            Dim JmlEarlyOt As Double = 0
            Dim PermitFrom As Double = 0
            Dim PermitTo As Double = 0
            Dim PermitDuration As Double = 0

            For Each RowDgv In Me.dgvItem.Rows

                If Me.ObjCd.Is_Time_String(RowDgv.Cells(Me.ColActualIn.Index).Value) = False OrElse Me.ObjCd.Is_Time_String(RowDgv.Cells(Me.ColActualOut.Index).Value) = False Then
                    MsgBox("There are invalid actual in or actual out.", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                CboDgv = RowDgv.Cells(Me.ColWorkdays.Index)
                workday = CboDgv.Items.IndexOf(RowDgv.Cells(Me.ColWorkdays.Index).Value)
                CboDgv = RowDgv.Cells(Me.ColDuration.Index)
                Durasi = CboDgv.Items.IndexOf(RowDgv.Cells(Me.ColDuration.Index).Value)
                CboDgv = RowDgv.Cells(Me.ColAttendanceStatus.Index)
                Status = CboDgv.Items.IndexOf(RowDgv.Cells(Me.ColAttendanceStatus.Index).Value)



                If RowDgv.Cells(Me.ColPrimaryLA.Index).Value > 0 Then
                    If Me.ObjLA.GetStatus(RowDgv.Cells(Me.ColPrimaryLA.Index).Value) <> ClsTrans_Cuti_Pengajuan.EnumStatus.App Then
                        MsgBox("The following leave application must be approved before save :" + vbNewLine + _
                        RowDgv.Cells(Me.ColDate.Index).Value + ", " + _
                        "Worksift " + RowDgv.Cells(Me.ColWorkshift.Index).Value + ", " + _
                        RowDgv.Cells(Me.ColLeaveDoc.Index).Value + ", At Row " + _
                        (RowDgv.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                        Me.dgvItem.ClearSelection()
                        Me.dgvItem.Rows(RowDgv.Index).Selected = True
                        Return False
                    End If
                End If

                'batasan durasi sama workshift
                ActIn = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(Me.ColActualIn.Index).Value)
                ActBreakOut = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(Me.ColActualBreakOut.Index).Value)
                ActBreakIn = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(Me.ColActualBreakIn.Index).Value)
                ActOut = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(Me.ColActualOut.Index).Value)

                PermitFrom = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(Me.ColPermitOut.Index).Value)
                PermitTo = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(Me.ColPermitIn.Index).Value)
                'jika semua baris 00.00
                If Me.ObjAtt.CekHalfStatus(Status, True) = True Then
                    If ActIn = 0 And _
                    ActBreakOut = 0 And _
                    ActBreakIn = 0 And _
                    ActOut = 0 Then
                        MsgBox("All actual duration can't fill 00.00 :" + vbNewLine + _
                        RowDgv.Cells(Me.ColDate.Index).Value + ", " + _
                        RowDgv.Cells(Me.ColWorkshift.Index).Value + ", " + _
                        RowDgv.Cells(Me.ColAttendanceStatus.Index).Value + ", At Row " + _
                        (RowDgv.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                        Me.dgvItem.ClearSelection()
                        Me.dgvItem.Rows(RowDgv.Index).Selected = True
                        Return False
                    End If
                End If

                ''cek duration, jika fullday batasan 8 jam, jika halfday batasan 4 jam.
                'If Me.ObjAtt.CekHalfStatus(Status) = True Then
                '    'setengah hari
                '    BatasDurasi = 240
                '    If Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(Me.ColActualWorkDuration.Index).Value) > (5 * 60) Then
                '        MsgBox("Work duration cannot more than 5 hours :" + vbNewLine + _
                '        RowDgv.Cells(Me.ColDate.Index).Value + ", " + _
                '        RowDgv.Cells(Me.ColWorkshift.Index).Value + ", " + _
                '        RowDgv.Cells(Me.ColAttendanceStatus.Index).Value + ", At Row " + _
                '        (RowDgv.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                '        Me.dgvItem.ClearSelection()
                '        Me.dgvItem.Rows(RowDgv.Index).Selected = True
                '        Return False
                '    End If
                'Else
                '    'fullday
                '    BatasDurasi = 480
                '    If Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(Me.ColActualWorkDuration.Index).Value) > (8 * 60) Then
                '        MsgBox("Work duration cannot more than 8 hours :" + vbNewLine + _
                '        RowDgv.Cells(Me.ColDate.Index).Value + ", " + _
                '        RowDgv.Cells(Me.ColWorkshift.Index).Value + ", " + _
                '        RowDgv.Cells(Me.ColAttendanceStatus.Index).Value + ", At Row " + _
                '        (RowDgv.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                '        Me.dgvItem.ClearSelection()
                '        Me.dgvItem.Rows(RowDgv.Index).Selected = True
                '        Return False
                '    End If
                'End If


                If ActBreakOut <> 0 And ActBreakIn <> 0 Then
                    'cek break duration
                    If ActOut > ActIn Then
                        'jika jam normal
                        If (ActIn > ActBreakOut Or ActBreakOut > ActOut) Or (ActIn > ActBreakIn Or ActBreakIn > ActOut) Then
                            'istirahat tidak boleh ada diluar jam kerja
                            MsgBox("Actual break must in work duration." + vbNewLine + _
                            RowDgv.Cells(Me.ColDate.Index).Value + ", " + _
                            RowDgv.Cells(Me.ColWorkshift.Index).Value + ", " + _
                            RowDgv.Cells(Me.ColAttendanceStatus.Index).Value + ", At Row " + _
                            (RowDgv.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                            Me.dgvItem.ClearSelection()
                            Me.dgvItem.Rows(RowDgv.Index).Selected = True
                            Return False
                        Else
                            If ActBreakOut > ActBreakIn Then
                                MsgBox("Actual break must in work duration." + vbNewLine + _
                                RowDgv.Cells(Me.ColDate.Index).Value + ", " + _
                                RowDgv.Cells(Me.ColWorkshift.Index).Value + ", " + _
                                RowDgv.Cells(Me.ColAttendanceStatus.Index).Value + ", At Row " + _
                                (RowDgv.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                                Me.dgvItem.ClearSelection()
                                Me.dgvItem.Rows(RowDgv.Index).Selected = True
                                Return False
                            End If
                        End If
                    Else
                        'istirahat tidak boleh ada diluar jam kerja, jam nyebrang hari
                        If ((ActBreakOut >= ActIn And ActBreakOut < (24 * 60)) Or ActBreakOut = 0) And _
                            (ActBreakIn <= ActOut) Then
                            'kondisi benar
                        ElseIf (ActBreakIn >= ActIn And ActBreakIn < (24 * 60) Or ActBreakIn = 0) And _
                                (ActBreakOut <= ActOut) Then
                            'kondisi terbalik
                            MsgBox("Actual break must in work duration." + vbNewLine + _
                            RowDgv.Cells(Me.ColDate.Index).Value + ", " + _
                            RowDgv.Cells(Me.ColWorkshift.Index).Value + ", " + _
                            RowDgv.Cells(Me.ColAttendanceStatus.Index).Value + ", At Row " + _
                            (RowDgv.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                            Me.dgvItem.ClearSelection()
                            Me.dgvItem.Rows(RowDgv.Index).Selected = True
                            Return False
                        ElseIf ((ActBreakIn >= ActIn And ActBreakIn < (24 * 60) Or ActBreakIn = 0) And _
                        (ActBreakOut >= ActIn And ActBreakOut < (24 * 60) Or ActBreakOut = 0)) Or _
                        (ActBreakOut <= ActOut And ActBreakIn <= ActOut) Then
                            'semua sebelah kiri atau sebelah kanan
                            If ActBreakOut > ActBreakIn Then
                                MsgBox("Actual break must in work duration." + vbNewLine + _
                                RowDgv.Cells(Me.ColDate.Index).Value + ", " + _
                                RowDgv.Cells(Me.ColWorkshift.Index).Value + ", " + _
                                RowDgv.Cells(Me.ColAttendanceStatus.Index).Value + ", At Row " + _
                                (RowDgv.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                                Me.dgvItem.ClearSelection()
                                Me.dgvItem.Rows(RowDgv.Index).Selected = True
                                Return False
                            End If
                        ElseIf (ActIn > ActBreakOut And ActBreakOut > ActOut) And _
                        (ActIn > ActBreakIn And ActBreakIn > ActOut) Then
                            'di luar jam  kerja
                            MsgBox("Actual break must in work duration." + vbNewLine + _
                            RowDgv.Cells(Me.ColDate.Index).Value + ", " + _
                            RowDgv.Cells(Me.ColWorkshift.Index).Value + ", " + _
                            RowDgv.Cells(Me.ColAttendanceStatus.Index).Value + ", At Row " + _
                            (RowDgv.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                            Me.dgvItem.ClearSelection()
                            Me.dgvItem.Rows(RowDgv.Index).Selected = True
                            Return False
                        End If
                    End If
                End If

                'early in
                If Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(Me.ColEarlyInDuration.Index).Value) > _
                Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(Me.ColSchWorkDuration.Index).Value) Then
                    hasil = MsgBox("Early in duration is more than schedule work duration." + vbNewLine + _
                    RowDgv.Cells(Me.ColDate.Index).Value + ", " + _
                    RowDgv.Cells(Me.ColWorkshift.Index).Value + ", " + _
                    RowDgv.Cells(Me.ColAttendanceStatus.Index).Value + ", At Row " + _
                    (RowDgv.Index + 1).ToString + vbNewLine + "Continue save ?", MsgBoxStyle.YesNo, "Information")
                    If hasil = Windows.Forms.DialogResult.No Then
                        Me.dgvItem.ClearSelection()
                        Me.dgvItem.Rows(RowDgv.Index).Selected = True
                        Return False
                    End If
                End If
                'Late in
                If Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(Me.ColLateDuration.Index).Value) > _
                Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(Me.ColSchWorkDuration.Index).Value) Then
                    hasil = MsgBox("Late in duration is more than schedule work duration." + vbNewLine + _
                    RowDgv.Cells(Me.ColDate.Index).Value + ", " + _
                    RowDgv.Cells(Me.ColWorkshift.Index).Value + ", " + _
                    RowDgv.Cells(Me.ColAttendanceStatus.Index).Value + ", At Row " + _
                    (RowDgv.Index + 1).ToString + vbNewLine + "Continue save ?", MsgBoxStyle.YesNo, "Information")
                    If hasil = Windows.Forms.DialogResult.No Then
                        Me.dgvItem.ClearSelection()
                        Me.dgvItem.Rows(RowDgv.Index).Selected = True
                        Return False
                    End If
                End If
                'Early out
                If Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(Me.ColEarlyOutDuration.Index).Value) > _
                Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(Me.ColSchWorkDuration.Index).Value) Then
                    hasil = MsgBox("Early out duration is more than schedule work duration." + vbNewLine + _
                    RowDgv.Cells(Me.ColDate.Index).Value + ", " + _
                    RowDgv.Cells(Me.ColWorkshift.Index).Value + ", " + _
                    RowDgv.Cells(Me.ColAttendanceStatus.Index).Value + ", At Row " + _
                    (RowDgv.Index + 1).ToString + vbNewLine + "Continue save ?", MsgBoxStyle.YesNo, "Information")
                    If hasil = Windows.Forms.DialogResult.No Then
                        Me.dgvItem.ClearSelection()
                        Me.dgvItem.Rows(RowDgv.Index).Selected = True
                        Return False
                    End If
                End If


                'permit durasi
                If PermitFrom <> 0 And PermitTo <> 0 Then
                    'cek break duration
                    If ActOut > ActIn Then
                        'jika jam normal
                        If (ActIn > PermitFrom And PermitFrom > ActOut) Or (ActIn > PermitTo And PermitTo > ActOut) Then
                            'permit tidak boleh ada diluar jam kerja
                            MsgBox("Permit time must in work duration." + vbNewLine + _
                            RowDgv.Cells(Me.ColDate.Index).Value + ", " + _
                            RowDgv.Cells(Me.ColWorkshift.Index).Value + ", " + _
                            RowDgv.Cells(Me.ColAttendanceStatus.Index).Value + ", At Row " + _
                            (RowDgv.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                            Me.dgvItem.ClearSelection()
                            Me.dgvItem.Rows(RowDgv.Index).Selected = True
                            Return False
                        Else
                            If PermitFrom > PermitTo Then
                                MsgBox("Permit time must in work duration." + vbNewLine + _
                                RowDgv.Cells(Me.ColDate.Index).Value + ", " + _
                                RowDgv.Cells(Me.ColWorkshift.Index).Value + ", " + _
                                RowDgv.Cells(Me.ColAttendanceStatus.Index).Value + ", At Row " + _
                                (RowDgv.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                                Me.dgvItem.ClearSelection()
                                Me.dgvItem.Rows(RowDgv.Index).Selected = True
                                Return False
                            End If
                        End If
                    Else
                        'istirahat tidak boleh ada diluar jam kerja, jam nyebrang hari
                        If ((PermitFrom >= ActIn And PermitFrom < (24 * 60)) Or PermitFrom = 0) And _
                            (PermitTo <= ActOut) Then
                            'kondisi benar
                        ElseIf (PermitTo >= ActIn And PermitTo < (24 * 60) Or PermitTo = 0) And _
                                (PermitFrom <= ActOut) Then
                            'kondisi terbalik
                            MsgBox("Permit time must in work duration." + vbNewLine + _
                            RowDgv.Cells(Me.ColDate.Index).Value + ", " + _
                            RowDgv.Cells(Me.ColWorkshift.Index).Value + ", " + _
                            RowDgv.Cells(Me.ColAttendanceStatus.Index).Value + ", At Row " + _
                            (RowDgv.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                            Me.dgvItem.ClearSelection()
                            Me.dgvItem.Rows(RowDgv.Index).Selected = True
                            Return False
                        ElseIf ((PermitTo >= ActIn And PermitTo < (24 * 60) Or PermitTo = 0) And _
                        (PermitFrom >= ActIn And PermitFrom < (24 * 60) Or PermitFrom = 0)) Or _
                        (PermitFrom <= ActOut And PermitTo <= ActOut) Then
                            'semua sebelah kiri atau sebelah kanan
                            If PermitFrom > PermitTo Then
                                MsgBox("Permit time must in work duration." + vbNewLine + _
                                RowDgv.Cells(Me.ColDate.Index).Value + ", " + _
                                RowDgv.Cells(Me.ColWorkshift.Index).Value + ", " + _
                                RowDgv.Cells(Me.ColAttendanceStatus.Index).Value + ", At Row " + _
                                (RowDgv.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                                Me.dgvItem.ClearSelection()
                                Me.dgvItem.Rows(RowDgv.Index).Selected = True
                                Return False
                            End If
                        ElseIf (ActIn > PermitFrom And PermitFrom > ActOut) And _
                        (ActIn > PermitTo And PermitTo > ActOut) Then
                            'di luar jam  kerja
                            MsgBox("Permit time must in work duration." + vbNewLine + _
                            RowDgv.Cells(Me.ColDate.Index).Value + ", " + _
                            RowDgv.Cells(Me.ColWorkshift.Index).Value + ", " + _
                            RowDgv.Cells(Me.ColAttendanceStatus.Index).Value + ", At Row " + _
                            (RowDgv.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                            Me.dgvItem.ClearSelection()
                            Me.dgvItem.Rows(RowDgv.Index).Selected = True
                            Return False
                        End If
                    End If
                End If

                JmlOt = JmlOt + Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(Me.ColOTDuration.Index).Value)
                JmlEarlyOt = JmlEarlyOt + Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(Me.ColEarlyOTDuration.Index).Value)

                ''cek konsistensi checklist
                With RowDgv
                    If .Cells(Me.ColIsCalculateEarlyOT.Index).Value = False And Me.ObjCd.Convert_ToMinutes(.Cells(Me.ColEarlyOTDuration.Index).Value) > 0 Then
                        MsgBox("Inconsistent early OT duration with calculation type check data at row " & RowDgv.Index + 1 & ".", MsgBoxStyle.Information, "Information")
                    End If
                    If .Cells(Me.ColIsCalculateLate.Index).Value = False And Me.ObjCd.Convert_ToMinutes(.Cells(Me.ColLateDuration.Index).Value) > 0 Then
                        MsgBox("Inconsistent late duration with calculation type check data at row " & RowDgv.Index + 1 & ".", MsgBoxStyle.Information, "Information")
                    End If
                    If .Cells(Me.ColIsCalculateEarlyOut.Index).Value = False And Me.ObjCd.Convert_ToMinutes(.Cells(Me.ColEarlyOutDuration.Index).Value) > 0 Then
                        MsgBox("Inconsistent early out duration with calculation type check data at row " & RowDgv.Index + 1 & ".", MsgBoxStyle.Information, "Information")
                    End If
                    If .Cells(Me.ColIsCalculateOT.Index).Value = False And Me.ObjCd.Convert_ToMinutes(.Cells(Me.ColOTDuration.Index).Value) > 0 Then
                        MsgBox("Inconsistent OT duration with calculation type check data at row " & RowDgv.Index + 1 & ".", MsgBoxStyle.Information, "Information")
                    End If
                End With

            Next

            'cek max lembur
            If String.IsNullOrEmpty(Me.PeriodMaxOt) = False Then
                If JmlOt > Me.PeriodMaxOt Then
                    hasil = MsgBox("Overtime duration in this month is more than max overtime duration." + vbNewLine + _
                    "Countinue save data ?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = Windows.Forms.DialogResult.No Then
                        Return False
                    End If
                End If
            End If
            If String.IsNullOrEmpty(Me.PeriodMaxOt_EarlyOt) = False Then
                If JmlEarlyOt > Me.PeriodMaxOt_EarlyOt Then
                    hasil = MsgBox("Early overtime duration in this month is more than max early overtime duration." + vbNewLine + _
                    "Countinue save data ?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = Windows.Forms.DialogResult.No Then
                        Return False
                    End If
                End If
            End If

            'dtp dan dgv harus sama sama waktunya
            If CDate(Me.dgvItem.Item(Me.ColDate.Index, 0).Value) <> Me.dtpStart.Value.Date Or _
            CDate(Me.dgvItem.Item(Me.ColDate.Index, Me.dgvItem.Rows.Count - 1).Value) <> Me.DtpEnd.Value.Date Then
                MsgBox("Inconsistent attendance period. Please contact your vendor if you see this message.", _
                MsgBoxStyle.Information, "Information")
                Me.BtnReload.Focus()

                Return False
            End If

            Dim ListDateDgv As New List(Of Date)
            For Each rD As DataGridViewRow In Me.dgvItem.Rows
                If ListDateDgv.Contains(CDate(rD.Cells(Me.ColDate.Index).Value)) = False Then
                    ListDateDgv.Add(CDate(rD.Cells(Me.ColDate.Index).Value))
                End If
            Next

            If Me.CekLompatHari() = False Then
                Return False
            End If

            'cek kemungkinan:
            '1. ada lembur tapi tidak ada actual in dan out
            For Each RowDgv In Me.dgvItem.Rows
                If Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(Me.ColOTDuration.Index).Value) + Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(Me.ColEarlyOTDuration.Index).Value) + _
                    Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(Me.ColEarlyOutDuration.Index).Value) + Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(Me.ColLateDuration.Index).Value) > 0 Then
                    If Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(Me.ColActualIn.Index).Value) + Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(Me.ColActualOut.Index).Value) = 0 Then
                        MsgBox("Wrong data input on : " + RowDgv.Cells(Me.ColDate.Index).Value + ", please specify actual in and out data.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If

                'pastikan actual in dengan actual out tidak boleh sama
                If Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(Me.ColActualIn.Index).Value) = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(Me.ColActualOut.Index).Value) Then
                    If Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(Me.ColActualIn.Index).Value) > 0 Then
                        MsgBox("Actual in can't be the same as actual out", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.ClearSelection()
                        Me.dgvItem.Rows(RowDgv.Index).Selected = True
                        Return False
                    End If
                End If
            Next

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub AktifkanDetail()
        Try
            Me.ObjCd.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCd.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCd.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCd.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.CancelButton = Me.btnCancel
            Me.BtnChange.Enabled = False

            'If Me.Baru = False Then
            '    'IsiTblOvertimeOrder
            '    Me.IsiTblSpl(Me.JoinWithAttPrev)
            'End If

            'cek juga untuk overtime order
            'Me.CekVisibleCol_OTOrder()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCd.KendalikanControl(Me.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCd.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCd.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCd.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.CancelButton = Me.btnExit
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongKanField()
        Try
            Me.ObjCd.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, True)
            Me.ObjCd.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, True)
            Me.dgvItem.Rows.Clear()
            Me.TxtPrimary.Text = -1
            Me.TxtPrimaryEmployee.Text = -1
            Me.TxtPrimaryPosition.Text = -1
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "

            'aul tambahan
            Me.ApakahSudahReload = False
            Me.JoinWithAttPrev = False
            Me.DtpStartAttPrev = "1/1/1950"
            Me.DtpToAttPrev = "1/1/1950"
            Me.BtnChange.Enabled = False
            If Me.btnOK.Enabled = True Then
                Me.dtpStart.Enabled = True
                Me.BtnReload.Enabled = True
            End If

            'clear tbl spl
            Me.TblOT_Order.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeObject()
        Try
            With Me.ObjAtt
                .NoTrans = Me.TxtDocNumber.Text
                .Tanggal = Me.DtpDate.Value
                .PrimaryEmp = Me.TxtPrimaryEmployee.Text
                .PrimaryPosition = Me.TxtPrimaryPosition.Text
                .StrDate = Me.dtpStart.Value
                .EndDate = Me.DtpEnd.Value
                .Remarks = Me.TxtRemarks.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeField()
        Try
            With Me.ObjAtt
                Me.TxtPrimary.Text = .Primary
                Me.TxtDocNumber.Text = .NoTrans
                Me.DtpDate.Value = .Tanggal
                Me.TxtPrimaryEmployee.Text = .PrimaryEmp
                Me.TxtEmployeeCode.Text = Me.ObjEmp.GetCodeByPrimary(.PrimaryEmp)
                Me.TxtEmployeeName.Text = Me.ObjEmp.GetFullNameByPrimary(.PrimaryEmp)
                Me.TxtPrimaryPosition.Text = Me.ObjEmp.GetJobByPrimary(.PrimaryEmp)
                Me.TxtPosition.Text = Me.ObjJob.GetDesc(Me.TxtPrimaryPosition.Text)
                Me.dtpStart.Value = .StrDate
                Me.DtpEnd.Value = .EndDate
                Me.TxtRemarks.Text = .Remarks
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With

            Me.SetKolomPermit()

            'cek juga untuk overtime order
            Me.CekVisibleCol_OTOrder()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanNonvisibleDGv(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            'break data
            If Me.CheckBox1.Checked = True Then
                Me.dgvItem.Item(Me.ColActualBreakOut.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0)
                Me.dgvItem.Item(Me.ColActualBreakIn.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0)
                Me.dgvItem.Item(Me.ColEarlyBreakOutDuration.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0)
                Me.dgvItem.Item(Me.ColLateBreakInDuration.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0)
            End If
            'early in
            If Me.CheckBox2.Checked = True Then
                Me.dgvItem.Item(Me.ColEarlyInDuration.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0)
                Me.dgvItem.Item(Me.ColEarlyOTDuration.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0)
            End If
            'early out
            If Me.CheckBox3.Checked = True Then
                Me.dgvItem.Item(Me.ColEarlyOutDuration.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Hitung(ByVal vIdx As Integer, ByVal vColIdx As Integer, ByVal vRangeCalculate As EnumCalculateRange, _
     ByVal UpdOvertimeOrder As Boolean)
        Dim CboDgv As Windows.Forms.DataGridViewComboBoxCell
        Dim SchIn As Double = 0
        Dim SchBreakOut As Double = 0
        Dim SchBreakIn As Double = 0
        Dim SchOut As Double = 0
        Dim ActIn As Double = 0
        Dim ActBreakOut As Double = 0
        Dim ActBreakIn As Double = 0
        Dim ActOut As Double = 0
        Dim RoundingEarlyOT As String = ""
        Dim workday As ClsTrans_Attendance.EnumWorkday
        Dim durasi As ClsTrans_Attendance.EnumDurasi
        Dim status As ClsTrans_Attendance.EnumStatus
        Dim CalAfterEOT As Double = 0
        Dim CalAfterLate As Double = 0
        Dim CalAfterEOut As Double = 0
        Dim CalAfterOT As Double = 0
        Dim RTypeEOT As Integer = 0
        Dim RTypeLate As Integer = 0
        Dim RTypeEOut As Integer = 0
        Dim RTypeOT As Integer = 0
        Dim RValueEOT As Decimal = 0
        Dim RValueLate As Decimal = 0
        Dim RValueEOut As Decimal = 0
        Dim RValueOT As Decimal = 0
        Dim MValueEOT As Double = 0
        Dim MValueLate As Double = 0
        Dim MValueEOut As Double = 0
        Dim MValueOT As Double = 0
        Dim max As String

        Dim PermitFrom As Double = 0
        Dim PermitTo As Double = 0

        Try
            If Me.ObjCd.Is_Time_String(Me.dgvItem.Item(Me.ColActualIn.Index, vIdx).Value.ToString) = False Or _
            Me.ObjCd.Is_Time_String(Me.dgvItem.Item(Me.ColActualOut.Index, vIdx).Value.ToString) = False Then
                Exit Sub
            End If
            SchIn = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColSchIn.Index, vIdx).Value)
            SchBreakIn = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColSchBreakIn.Index, vIdx).Value)
            SchBreakOut = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColSchBreakOut.Index, vIdx).Value)
            SchOut = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColSchOut.Index, vIdx).Value)
            ActIn = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColActualIn.Index, vIdx).Value)
            ActBreakOut = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColActualBreakOut.Index, vIdx).Value)
            ActBreakIn = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColActualBreakIn.Index, vIdx).Value)
            ActOut = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColActualOut.Index, vIdx).Value)
            PermitFrom = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColPermitOut.Index, vIdx).Value)
            PermitTo = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColPermitIn.Index, vIdx).Value)
            'workday
            CboDgv = Me.dgvItem.Item(Me.ColWorkdays.Index, vIdx)
            workday = CboDgv.Items.IndexOf(Me.dgvItem.Item(Me.ColWorkdays.Index, vIdx).Value)
            'fullday / half day
            CboDgv = Me.dgvItem.Item(Me.ColDuration.Index, vIdx)
            durasi = CboDgv.Items.IndexOf(Me.dgvItem.Item(Me.ColDuration.Index, vIdx).Value)
            'status
            CboDgv = Me.dgvItem.Item(Me.ColAttendanceStatus.Index, vIdx)
            status = CboDgv.Items.IndexOf(Me.dgvItem.Item(Me.ColAttendanceStatus.Index, vIdx).Value)

            CalAfterEOT = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColCalculateAfterEarlyOT.Index, vIdx).Value)
            CalAfterLate = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColCalculateAfterEarlyLate.Index, vIdx).Value)
            CalAfterEOut = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColCalculateAfterEarlyOut.Index, vIdx).Value)
            CalAfterOT = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColCalculateAfterOT.Index, vIdx).Value)

            CboDgv = Me.dgvItem.Item(Me.ColRoundingTypeEarlyOT.Index, vIdx)
            RTypeEOT = CboDgv.Items.IndexOf(Me.dgvItem.Item(Me.ColRoundingTypeEarlyOT.Index, vIdx).Value)
            RTypeLate = CboDgv.Items.IndexOf(Me.dgvItem.Item(Me.ColRoundingTypeLate.Index, vIdx).Value)
            RTypeEOut = CboDgv.Items.IndexOf(Me.dgvItem.Item(Me.ColRoundingTypeEarlyOut.Index, vIdx).Value)
            RTypeOT = CboDgv.Items.IndexOf(Me.dgvItem.Item(Me.ColRoundingTypeOT.Index, vIdx).Value)

            RValueEOT = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColRoundingValueEarlyOT.Index, vIdx).Value)
            RValueLate = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColRoundingValueLate.Index, vIdx).Value)
            RValueEOut = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColRoundingValueEarlyOut.Index, vIdx).Value)
            RValueOT = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColRoundingValueOT.Index, vIdx).Value)

            MValueEOT = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColMaxValueEarlyOT.Index, vIdx).Value)
            MValueLate = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColMaxValueLate.Index, vIdx).Value)
            MValueEOut = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColMaxValueEarlyOut.Index, vIdx).Value)
            MValueOT = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColMaxValueOT.Index, vIdx).Value)

            If ActIn = 0 And ActBreakOut = 0 And ActBreakIn = 0 And ActOut = 0 Then
                Me.dgvItem.Item(Me.ColSchBreakDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                Me.dgvItem.Item(Me.ColSchWorkDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                Me.dgvItem.Item(Me.ColEarlyInDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                Me.dgvItem.Item(Me.ColIsEarlyIn.Index, vIdx).Value = 0
                Me.dgvItem.Item(Me.ColEarlyOTDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                Me.dgvItem.Item(Me.ColIsEarlyOvertime.Index, vIdx).Value = 0
                Me.dgvItem.Item(Me.ColLateDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                Me.dgvItem.Item(Me.ColIsLate.Index, vIdx).Value = 0
                Me.dgvItem.Item(Me.ColEarlyOutDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                Me.dgvItem.Item(Me.ColIsEarlyOut.Index, vIdx).Value = 0
                Me.dgvItem.Item(Me.ColOTDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                Me.dgvItem.Item(Me.ColIsOvertime.Index, vIdx).Value = 0
                Me.dgvItem.Item(Me.ColEarlyBreakOutDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                Me.dgvItem.Item(Me.ColIsEarlyBreakOut.Index, vIdx).Value = 0
                Me.dgvItem.Item(Me.ColLateBreakInDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                Me.dgvItem.Item(Me.ColIsLateBreakIn.Index, vIdx).Value = 0
                Me.dgvItem.Item(Me.ColActualBreakDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                Me.dgvItem.Item(Me.ColActualWorkDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                Me.dgvItem.Item(Me.ColPermitDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                Me.dgvItem.Item(Me.ColPermitOut.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                Me.dgvItem.Item(Me.ColPermitIn.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                'Me.dgvItem.Item(Me.ColCalculateAfterEarlyOT.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                'Me.dgvItem.Item(Me.ColCalculateAfterEarlyLate.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                'Me.dgvItem.Item(Me.ColCalculateAfterEarlyOut.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                'Me.dgvItem.Item(Me.ColCalculateAfterOT.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                'Me.dgvItem.Item(Me.ColRoundingValueEarlyOT.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                'Me.dgvItem.Item(Me.ColRoundingValueLate.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                'Me.dgvItem.Item(Me.ColRoundingValueEarlyOut.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                'Me.dgvItem.Item(Me.ColRoundingValueOT.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                'Me.dgvItem.Item(Me.ColMaxValueEarlyOT.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                'Me.dgvItem.Item(Me.ColMaxValueLate.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                'Me.dgvItem.Item(Me.ColMaxValueEarlyOut.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                'Me.dgvItem.Item(Me.ColMaxValueOT.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)

                'CboDgv = Me.dgvItem.Item(Me.ColRoundingTypeEarlyOT.Index, vIdx)
                'Me.dgvItem.Item(Me.ColRoundingTypeEarlyOT.Index, vIdx).Value = CboDgv.Items(Core_Module.ClsMasterWorkShift.enumRoundingType.No_Rounding)
                'Me.dgvItem.Item(Me.ColRoundingTypeLate.Index, vIdx).Value = CboDgv.Items(Core_Module.ClsMasterWorkShift.enumRoundingType.No_Rounding)
                'Me.dgvItem.Item(Me.ColRoundingTypeEarlyOut.Index, vIdx).Value = CboDgv.Items(Core_Module.ClsMasterWorkShift.enumRoundingType.No_Rounding)
                'Me.dgvItem.Item(Me.ColRoundingTypeOT.Index, vIdx).Value = CboDgv.Items(Core_Module.ClsMasterWorkShift.enumRoundingType.No_Rounding)
            Else
                'sch break duration
                Dim SchBreak As Double
                If SchBreakOut = SchBreakIn Then
                    SchBreak = 0
                Else
                    If SchBreakOut > SchBreakIn Then
                        SchBreak = ((24 * 60) - SchBreakOut) + SchBreakIn
                    Else
                        SchBreak = SchBreakIn - SchBreakOut
                    End If
                End If
                Me.dgvItem.Item(Me.ColSchBreakDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(SchBreak)

                'sch work duration
                Dim SchWork As Double
                If SchOut >= SchIn Then
                    SchWork = SchOut - SchIn - SchBreak
                Else
                    SchWork = ((24 * 60) - SchIn) + SchOut - SchBreak
                End If
                Me.dgvItem.Item(Me.ColSchWorkDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(SchWork - SchBreak)
                ''===================================================================================================================================
                'early in
                Dim EarlyIn As Double = 0
                Dim EarlyInDummy As Double = 0
                If (status <> ClsTrans_Attendance.EnumStatus.Present Or status <> ClsTrans_Attendance.EnumStatus.Half_Holiday) And _
                    workday <> ClsTrans_Attendance.EnumWorkday.Holiday Then
                    If SchIn < SchOut Then ' jam normal, contoh 08:00 - 16:00
                        If ActIn < SchOut Then
                            EarlyInDummy = SchIn - ActIn
                            If EarlyInDummy <= CalAfterEOT OrElse EarlyInDummy <= 0 Then
                                EarlyIn = 0
                            Else
                                EarlyIn = EarlyInDummy
                            End If
                        Else
                            EarlyInDummy = ((24 * 60) - ActIn) + SchIn
                            If EarlyInDummy <= CalAfterEOT OrElse EarlyInDummy <= 0 Then
                                EarlyIn = 0
                            Else
                                EarlyIn = EarlyInDummy
                            End If
                        End If
                    Else ' jam nyebrang hari, contoh 08:00 - 08:00
                        If ActIn > SchOut Then
                            EarlyInDummy = SchIn - ActIn
                            If EarlyInDummy <= CalAfterEOT OrElse EarlyInDummy <= 0 Then
                                EarlyIn = 0
                            Else
                                EarlyIn = EarlyInDummy
                            End If
                        End If
                    End If
                End If

                'kurangi dengan tolerance
                If Me.dgvItem.Item(Me.ColCalculateEarlyOTTolerance.Index, vIdx).Value = 0 Then
                    If EarlyIn - Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColCalculateAfterEarlyOT.Index, vIdx).Value) > 0 Then
                        EarlyIn = EarlyIn - Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColCalculateAfterEarlyOT.Index, vIdx).Value)
                    End If
                End If

                'rounding early in
                If RTypeEOT = Core_Module.ClsMasterWorkShift.enumRoundingType.No_Rounding Then
                ElseIf RTypeEOT = Core_Module.ClsMasterWorkShift.enumRoundingType.Round_down Then
                    EarlyIn = Me.ObjCd.RoundDown2(CDec(EarlyIn), RValueEOT)
                ElseIf RTypeEOT = Core_Module.ClsMasterWorkShift.enumRoundingType.Round_Up Then
                    EarlyIn = Me.ObjCd.RoundUp2(CDec(EarlyIn), RValueEOT)
                ElseIf RTypeEOT = Core_Module.ClsMasterWorkShift.enumRoundingType.Middle Then
                    EarlyIn = Me.ObjCd.MiddleRound(RValueEOT, CDec(EarlyIn))
                End If

                If MValueEOT > 0 Then
                    If MValueEOT >= EarlyIn Then
                        Me.dgvItem.Item(Me.ColEarlyInDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(EarlyIn)
                    ElseIf MValueEOT < EarlyIn Then
                        Me.dgvItem.Item(Me.ColEarlyInDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(MValueEOT)
                    End If
                Else
                    Me.dgvItem.Item(Me.ColEarlyInDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(EarlyIn)
                End If

                'is early in
                If Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColEarlyInDuration.Index, vIdx).Value) > 0 Then
                    Me.dgvItem.Item(Me.ColIsEarlyIn.Index, vIdx).Value = 1
                Else
                    Me.dgvItem.Item(Me.ColIsEarlyIn.Index, vIdx).Value = 0
                End If

                Dim EarlyOt As Double = EarlyIn
                If vRangeCalculate = EnumCalculateRange.BOTH Or vRangeCalculate = EnumCalculateRange.IN_RANGE Then
                    If Me.dgvItem.Item(Me.ColIsCalculateEarlyOT.Index, vIdx).Value = Boolean.TrueString Then
                        If (status <> ClsTrans_Attendance.EnumStatus.Present Or _
                            status <> ClsTrans_Attendance.EnumStatus.Half_Holiday) And _
                           workday <> ClsTrans_Attendance.EnumWorkday.Holiday Then
                            If workday = ClsTrans_Attendance.EnumWorkday.Holiday Then
                                If ActOut >= ActIn Then
                                    EarlyOt = ActOut - ActIn
                                Else
                                    EarlyOt = ((24 * 60) - ActIn) + ActOut
                                End If
                            End If
                        Else
                            EarlyOt = 0
                        End If
                    Else
                        EarlyOt = 0
                    End If

                    If MValueEOT > 0 Then
                        If MValueEOT >= EarlyOt Then
                            Me.dgvItem.Item(Me.ColEarlyOTDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(EarlyOt)
                        ElseIf MValueEOT < EarlyOt Then
                            Me.dgvItem.Item(Me.ColEarlyOTDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(MValueEOT)
                        End If
                    Else
                        Me.dgvItem.Item(Me.ColEarlyOTDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(EarlyOt)
                    End If

                    'cek type ot
                    CboDgv = Me.dgvItem.Item(Me.ColOT.Index, vIdx)
                    If Me.dgvItem.Item(Me.ColOT.Index, vIdx).Value <> CboDgv.Items.Item(ClsTrans_Attendance.EnumTipeOt.No_Overtime) Then
                        If Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColEarlyOTDuration.Index, vIdx).Value) > 0 Then
                            Me.dgvItem.Item(Me.ColIsEarlyOvertime.Index, vIdx).Value = 1
                        Else
                            Me.dgvItem.Item(Me.ColIsEarlyOvertime.Index, vIdx).Value = 0
                        End If
                    Else
                        Me.dgvItem.Item(Me.ColIsEarlyOvertime.Index, vIdx).Value = 0
                    End If
                End If

                '=======================================================================================================
                'late
                Dim LateIn As Double = 0
                Dim LateInDummy As Double = 0
                If vRangeCalculate = EnumCalculateRange.BOTH Or vRangeCalculate = EnumCalculateRange.IN_RANGE Then
                    If Me.dgvItem.Item(Me.ColIsCalculateLate.Index, vIdx).Value = Boolean.TrueString Then
                        If status = ClsTrans_Attendance.EnumStatus.Present AndAlso _
                            workday <> ClsTrans_Attendance.EnumWorkday.Holiday Then
                            If SchIn < SchOut Then
                                If ActIn < SchOut Then
                                    'buka - repair late lewah istirahat
                                    If ActIn >= SchBreakOut And ActIn <= SchBreakIn Then
                                        LateInDummy = SchBreakOut - SchIn
                                    ElseIf ActIn > SchBreakIn And ActIn > SchBreakOut Then
                                        LateInDummy = ActIn - SchIn - (SchBreakIn - SchBreakOut)
                                    Else
                                        LateInDummy = ActIn - SchIn
                                    End If
                                    'tutup - repair late lewah istirahat

                                    If LateInDummy <= CalAfterLate OrElse LateInDummy <= 0 Then
                                        LateIn = 0
                                    Else
                                        LateIn = LateInDummy
                                    End If
                                ElseIf ActIn > SchOut Then
                                    LateIn = 0
                                End If
                            ElseIf SchIn > SchOut Then
                                If ActIn < SchOut Then
                                    LateInDummy = ((24 * 60) - SchIn) + ActIn
                                    If LateInDummy <= CalAfterLate OrElse LateInDummy <= 0 Then
                                        LateIn = 0
                                    Else
                                        LateIn = LateInDummy
                                    End If
                                Else
                                    LateInDummy = ActIn - SchIn
                                    If LateInDummy <= CalAfterLate OrElse LateInDummy <= 0 Then
                                        LateIn = 0
                                    Else
                                        LateIn = LateInDummy
                                    End If
                                End If
                            End If

                        ElseIf status <> ClsTrans_Attendance.EnumStatus.Half_Absent _
                               OrElse status <> ClsTrans_Attendance.EnumStatus.Half_day_Leave _
                               OrElse status <> ClsTrans_Attendance.EnumStatus.Half_Day_Permit _
                               OrElse status <> ClsTrans_Attendance.EnumStatus.Half_Day_Sick Then
                            If SchIn < SchOut Then
                                If ActIn < SchOut Then
                                    LateInDummy = ActIn - SchIn
                                    If LateInDummy <= CalAfterLate OrElse LateInDummy <= 0 Then
                                        LateIn = 0
                                    Else
                                        LateIn = LateInDummy
                                    End If
                                ElseIf ActIn > SchOut Then
                                    LateIn = 0
                                End If
                            ElseIf SchIn > SchOut Then
                                If ActIn < SchOut Then
                                    LateInDummy = ((24 * 60) - SchIn) + ActIn
                                    If LateInDummy <= CalAfterLate OrElse LateInDummy <= 0 Then
                                        LateIn = 0
                                    Else
                                        LateIn = LateInDummy
                                    End If
                                Else
                                    LateInDummy = ActIn - SchIn
                                    If LateInDummy <= CalAfterLate OrElse LateInDummy <= 0 Then
                                        LateIn = 0
                                    Else
                                        LateIn = LateInDummy
                                    End If
                                End If
                            End If
                            '********************************************************************************************************************
                            'KETENTUAN BARU DARI PAK IWAN BERDASARKAN ERROR LIST CI FITRI NO 49 - 26 JUNI 2016 - PERBAIKAN : 28 JULI 2016
                            'JIKA HALFDAY (LEAVE, HOLIDAY ETC) DILAKUKAN PENGECEKAN, JAM MASUK LEBIH BERAT SEBELUM ISTIRHAT ATAU SESUDAH
                            'JIKA MASUK SEBELUM ISTIRAHAT(PAGI) MAKA TIDAK BOLEH HITUNG EARLY OUT (EARLY OUT = 00.00) TAPI LATE DIHITUNG
                            'JIKA MASUK SESUDAH ISTIRAHAT(SIANG) MAKA TIDAK BOLEH HITUNG LATE (LATE = 00.00) TAPI EARLY OUT DIHITUNG
                            '********************************************************************************************************************
                            ' '' ''saya tutup, karena memaksa early out override 0::
                            ' ''Me.dgvItem.Item(Me.ColEarlyOutDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                            ' ''Me.dgvItem.Item(Me.ColIsEarlyOut.Index, vIdx).Value = 0
                            '********************************************************************************************************************

                        End If
                    Else
                        LateIn = 0
                    End If

                    'kurangi dengan tolerance
                    If Me.dgvItem.Item(Me.ColCalculateLateTolerance.Index, vIdx).Value = 0 Then
                        If LateIn - Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColCalculateAfterEarlyLate.Index, vIdx).Value) > 0 Then
                            LateIn = LateIn - Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColCalculateAfterEarlyLate.Index, vIdx).Value)
                        End If
                    End If

                    ''--
                    'rounding late in
                    If RTypeLate = Core_Module.ClsMasterWorkShift.enumRoundingType.No_Rounding Then

                    ElseIf RTypeLate = Core_Module.ClsMasterWorkShift.enumRoundingType.Round_down Then
                        LateIn = Me.ObjCd.RoundDown2(CDec(LateIn), RValueLate)
                    ElseIf RTypeLate = Core_Module.ClsMasterWorkShift.enumRoundingType.Round_Up Then
                        LateIn = Me.ObjCd.RoundUp2(CDec(LateIn), RValueLate)
                    ElseIf RTypeLate = Core_Module.ClsMasterWorkShift.enumRoundingType.Middle Then
                        LateIn = Me.ObjCd.MiddleRound(RValueLate, CDec(LateIn))
                    End If

                    If MValueLate > 0 Then
                        If MValueLate >= LateIn Then
                            Me.dgvItem.Item(Me.ColLateDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(LateIn)
                        ElseIf MValueLate < LateIn Then
                            Me.dgvItem.Item(Me.ColLateDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(MValueLate)
                        End If
                    Else
                        Me.dgvItem.Item(Me.ColLateDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(LateIn)
                    End If

                    'is late
                    If Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColLateDuration.Index, vIdx).Value) > 0 Then
                        Me.dgvItem.Item(Me.ColIsLate.Index, vIdx).Value = 1
                    Else
                        Me.dgvItem.Item(Me.ColIsLate.Index, vIdx).Value = 0
                    End If
                End If

                '=======================================================================================================
                'early out
                Dim EarlyOut As Double = 0
                Dim EarlyOutDummy As Double = 0
                If vRangeCalculate = EnumCalculateRange.BOTH Or vRangeCalculate = EnumCalculateRange.OUT_RANGE Then
                    If Me.dgvItem.Item(Me.ColIsCalculateEarlyOut.Index, vIdx).Value = Boolean.TrueString Then

                        If (status = ClsTrans_Attendance.EnumStatus.Present) And _
                            workday <> ClsTrans_Attendance.EnumWorkday.Holiday Then
                            If SchIn < SchOut Then
                                If ActOut > SchIn Then
                                    EarlyOutDummy = SchOut - ActOut
                                    If EarlyOutDummy <= CalAfterEOut OrElse EarlyOutDummy < 0 Then
                                        EarlyOut = 0
                                    Else
                                        EarlyOut = EarlyOutDummy
                                        If SchBreakOut < SchBreakIn And SchBreakIn < SchOut Then
                                            'konversi klo ada keluar pas jam istirahat
                                            If ActOut < SchBreakOut Then
                                                EarlyOut -= (SchBreakIn - SchBreakOut)
                                            ElseIf ActOut < SchBreakIn Then
                                                EarlyOut -= (SchBreakIn - ActOut)
                                            End If
                                        End If
                                    End If
                                Else
                                    EarlyOut = 0
                                End If
                            ElseIf SchIn > SchOut Then
                                If ActOut > SchIn Then
                                    EarlyOutDummy = ((24 * 60) - ActOut) + SchOut
                                    If EarlyOutDummy <= CalAfterEOut OrElse EarlyOutDummy < 0 Then
                                        EarlyOut = 0
                                    Else
                                        EarlyOut = EarlyOutDummy
                                    End If
                                ElseIf ActOut < SchIn Then
                                    EarlyOutDummy = SchOut - ActOut
                                    If EarlyOutDummy <= CalAfterEOut OrElse EarlyOutDummy < 0 Then
                                        EarlyOut = 0
                                    Else
                                        EarlyOut = EarlyOutDummy
                                    End If
                                End If

                                If EarlyOut <> 0 Then
                                    Dim SchBreakOutTemp As Double = SchBreakOut
                                    Dim SchBreakInTemp As Double = SchBreakIn
                                    Dim ActOutTemp As Double = ActOut
                                    If SchBreakOut < SchIn Then
                                        SchBreakOutTemp = SchBreakOut + (24 * 60)
                                    End If
                                    If SchBreakIn < SchIn Then
                                        SchBreakInTemp = SchBreakIn + (24 * 60)
                                    End If
                                    If ActOut < SchIn Then
                                        ActOutTemp = ActOut + (24 * 60)
                                    End If

                                    'konversi klo ada keluar pas jam istirahat
                                    If ActOutTemp < SchBreakOutTemp Then
                                        EarlyOut -= (SchBreakInTemp - SchBreakOutTemp)
                                    ElseIf ActOutTemp < SchBreakInTemp Then
                                        EarlyOut -= (SchBreakInTemp - ActOutTemp)
                                    End If
                                End If
                            End If
                        ElseIf status <> ClsTrans_Attendance.EnumStatus.Half_Absent _
                                             OrElse status <> ClsTrans_Attendance.EnumStatus.Half_day_Leave _
                                             OrElse status <> ClsTrans_Attendance.EnumStatus.Half_Day_Permit _
                                             OrElse status <> ClsTrans_Attendance.EnumStatus.Half_Day_Sick Then
                            If SchIn < SchOut Then
                                If ActOut > SchIn Then
                                    EarlyOutDummy = SchOut - ActOut
                                    If EarlyOutDummy <= CalAfterEOut OrElse EarlyOutDummy < 0 Then
                                        EarlyOut = 0
                                    Else
                                        EarlyOut = EarlyOutDummy
                                    End If
                                Else
                                    EarlyOut = 0
                                End If
                            ElseIf SchIn > SchOut Then
                                If ActOut > SchIn Then
                                    EarlyOutDummy = ((24 * 60) - ActOut) + SchOut
                                    If EarlyOutDummy <= CalAfterEOut OrElse EarlyOutDummy < 0 Then
                                        EarlyOut = 0
                                    Else
                                        EarlyOut = EarlyOutDummy
                                    End If
                                ElseIf ActOut < SchIn Then
                                    EarlyOutDummy = SchOut - ActOut
                                    If EarlyOutDummy <= CalAfterEOut OrElse EarlyOutDummy < 0 Then
                                        EarlyOut = 0
                                    Else
                                        EarlyOut = EarlyOutDummy
                                    End If
                                End If
                            End If
                            '********************************************************************************************************************
                            'KETENTUAN BARU DARI PAK IWAN BERDASARKAN ERROR LIST CI FITRI NO 49 - 26 JUNI 2016 - PERBAIKAN : 28 JULI 2016
                            'JIKA HALFDAY (LEAVE, HOLIDAY ETC) DILAKUKAN PENGECEKAN, JAM MASUK LEBIH BERAT SEBELUM ISTIRHAT ATAU SESUDAH
                            'JIKA MASUK SEBELUM ISTIRAHAT(PAGI) MAKA TIDAK BOLEH HITUNG EARLY OUT (EARLY OUT = 00.00) TAPI LATE DIHITUNG
                            'JIKA MASUK SESUDAH ISTIRAHAT(SIANG) MAKA TIDAK BOLEH HITUNG LATE (LATE = 00.00) TAPI EARLY OUT DIHITUNG
                            '********************************************************************************************************************
                            ' '' ''saya tutup, karena memaksa early out override 0::
                            ' ''Me.dgvItem.Item(Me.ColLateDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                            ' ''Me.dgvItem.Item(Me.ColIsLate.Index, vIdx).Value = 0
                            '********************************************************************************************************************

                        End If
                    Else
                        EarlyOut = 0
                    End If

                    'kurangi dengan tolerance
                    If Me.dgvItem.Item(Me.ColCalculateEarlyOutTolerance.Index, vIdx).Value = 0 Then
                        If EarlyOut - Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColCalculateAfterEarlyOut.Index, vIdx).Value) > 0 Then
                            EarlyOut = EarlyOut - Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColCalculateAfterEarlyOut.Index, vIdx).Value)
                        End If
                    End If

                    'rounding late in
                    If RTypeEOut = Core_Module.ClsMasterWorkShift.enumRoundingType.No_Rounding Then
                    ElseIf RTypeEOut = Core_Module.ClsMasterWorkShift.enumRoundingType.Round_down Then
                        EarlyOut = Me.ObjCd.RoundDown2(CDec(EarlyOut), RValueEOut)
                    ElseIf RTypeEOut = Core_Module.ClsMasterWorkShift.enumRoundingType.Round_Up Then
                        EarlyOut = Me.ObjCd.RoundUp2(CDec(EarlyOut), RValueEOut)
                    ElseIf RTypeEOut = Core_Module.ClsMasterWorkShift.enumRoundingType.Middle Then
                        EarlyOut = Me.ObjCd.MiddleRound(RValueEOut, CDec(EarlyOut))
                    End If

                    'cek maksimal early out
                    If MValueEOut > 0 Then
                        If MValueEOut >= EarlyOut Then
                            Me.dgvItem.Item(Me.ColEarlyOutDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(EarlyOut)
                        ElseIf MValueEOut < EarlyOut Then
                            Me.dgvItem.Item(Me.ColEarlyOutDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(MValueEOut)
                        End If
                    Else
                        Me.dgvItem.Item(Me.ColEarlyOutDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(EarlyOut)
                    End If

                    'is early out 
                    If Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColEarlyOutDuration.Index, vIdx).Value) > 0 Then
                        Me.dgvItem.Item(Me.ColIsEarlyOut.Index, vIdx).Value = 1
                    Else
                        Me.dgvItem.Item(Me.ColIsEarlyOut.Index, vIdx).Value = 0
                    End If
                End If

                '========================================================================================================
                'Overtime
                Dim Ot As Double
                Dim OtDummy As Double
                If vRangeCalculate = EnumCalculateRange.BOTH Or vRangeCalculate = EnumCalculateRange.OUT_RANGE Then
                    If Me.dgvItem.Item(Me.ColIsCalculateOT.Index, vIdx).Value = Boolean.TrueString Then
                        'cek dulu workday ato holiday
                        If workday = ClsTrans_Attendance.EnumWorkday.Holiday Then
                            If ActOut >= ActIn Then
                                If ActOut <= SchBreakOut Then
                                    OtDummy = ActOut - ActIn
                                ElseIf ActOut > SchBreakIn Then
                                    OtDummy = ActOut - ActIn - SchBreak
                                Else
                                    OtDummy = SchBreakOut - ActIn
                                End If
                            Else
                                OtDummy = ((24 * 60) - ActIn) + ActOut - SchBreak
                            End If

                            If OtDummy <= 0 OrElse OtDummy <= CalAfterOT Then
                                Ot = 0
                            Else
                                Ot = OtDummy
                            End If
                        Else
                            If SchIn < SchOut Then
                                If ActOut > SchIn Then
                                    If SchOut > ActOut Then
                                        Ot = 0
                                    Else
                                        OtDummy = ActOut - SchOut
                                        If OtDummy <= 0 OrElse OtDummy <= CalAfterOT Then
                                            Ot = 0
                                        Else
                                            Ot = OtDummy
                                        End If
                                    End If
                                Else
                                    OtDummy = ((24 * 60) - SchOut) + ActOut
                                    If OtDummy <= 0 OrElse OtDummy <= CalAfterOT Then
                                        Ot = 0
                                    Else
                                        Ot = OtDummy
                                    End If
                                End If
                            ElseIf SchIn > SchOut Then
                                If ActOut > SchIn Then
                                    Ot = 0
                                ElseIf ActOut < SchIn Then
                                    OtDummy = ActOut - SchOut
                                    If OtDummy <= 0 OrElse OtDummy <= CalAfterOT Then
                                        Ot = 0
                                    Else
                                        Ot = OtDummy
                                    End If
                                End If
                            End If
                        End If
                    Else
                        Ot = 0
                    End If

                    'kurangi dengan tolerance
                    If Me.dgvItem.Item(Me.ColCalculateOTTolerance.Index, vIdx).Value = 0 Then
                        If Ot - Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColCalculateAfterOT.Index, vIdx).Value) > 0 Then
                            Ot = Ot - Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColCalculateAfterOT.Index, vIdx).Value)
                        End If
                    End If

                    ''--
                    'rounding ot
                    If RTypeOT = Core_Module.ClsMasterWorkShift.enumRoundingType.No_Rounding Then

                    ElseIf RTypeOT = Core_Module.ClsMasterWorkShift.enumRoundingType.Round_down Then
                        Ot = Me.ObjCd.RoundDown2(CDec(Ot), RValueOT)
                    ElseIf RTypeOT = Core_Module.ClsMasterWorkShift.enumRoundingType.Round_Up Then
                        Ot = Me.ObjCd.RoundUp2(CDec(Ot), RValueOT)
                    ElseIf RTypeOT = Core_Module.ClsMasterWorkShift.enumRoundingType.Middle Then
                        Ot = Me.ObjCd.MiddleRound(RValueOT, Ot)
                    End If

                    If MValueOT > 0 Then
                        If MValueOT >= Ot Then
                            Me.dgvItem.Item(Me.ColOTDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(Ot)
                        ElseIf MValueOT < Ot Then
                            Me.dgvItem.Item(Me.ColOTDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(MValueOT)
                        End If
                    Else
                        Me.dgvItem.Item(Me.ColOTDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(Ot)
                    End If

                    'is overtime
                    'cek tipe ot
                    CboDgv = Me.dgvItem.Item(Me.ColOT.Index, vIdx)
                    If Me.dgvItem.Item(Me.ColOT.Index, vIdx).Value <> CboDgv.Items.Item(ClsTrans_Attendance.EnumTipeOt.No_Overtime) Then
                        If Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColOTDuration.Index, vIdx).Value) > 0 Then
                            Me.dgvItem.Item(Me.ColIsOvertime.Index, vIdx).Value = 1
                        Else
                            Me.dgvItem.Item(Me.ColIsOvertime.Index, vIdx).Value = 0
                        End If
                    Else
                        Me.dgvItem.Item(Me.ColIsOvertime.Index, vIdx).Value = 0
                    End If
                End If


                '===============================================================================================
                'early break out
                'cek nyebrang hari ato engga
                Dim EarlyBreakOut As Double = 0
                If (status <> ClsTrans_Attendance.EnumStatus.Present Or status <> ClsTrans_Attendance.EnumStatus.Half_Holiday) And _
                workday <> ClsTrans_Attendance.EnumWorkday.Holiday Then
                    If ActBreakIn <> 0 Or ActBreakOut <> 0 Then
                        If SchBreakOut < SchBreakIn Then
                            If ActBreakOut < SchBreakOut Then
                                EarlyBreakOut = SchBreakOut - ActBreakOut
                            ElseIf ActBreakOut > SchBreakOut And ActBreakOut > SchBreakIn Then
                                EarlyBreakOut = ((24 * 60) - ActBreakOut) - SchBreakIn
                            Else
                                EarlyBreakOut = 0
                            End If
                        Else
                            EarlyBreakOut = 0
                        End If
                    Else
                        EarlyBreakOut = 0
                    End If
                End If

                'cek maksimal EarlyBreakOut
                max = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.MaxValue_Early_Break_Out)
                If String.IsNullOrEmpty(max) = True Then
                    Me.dgvItem.Item(Me.ColEarlyBreakOutDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(EarlyBreakOut)
                Else
                    If max > EarlyIn Then
                        Me.dgvItem.Item(Me.ColEarlyBreakOutDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(EarlyBreakOut)
                    ElseIf max < EarlyIn Then
                        Me.dgvItem.Item(Me.ColEarlyBreakOutDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(max)
                    End If
                End If
                'is early break out
                If Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColEarlyBreakOutDuration.Index, vIdx).Value) > 0 Then
                    Me.dgvItem.Item(Me.ColIsEarlyBreakOut.Index, vIdx).Value = 1
                Else
                    Me.dgvItem.Item(Me.ColIsEarlyBreakOut.Index, vIdx).Value = 0
                End If
                ''============================================================================================================================
                'late break in
                'cek nyebrang hari ato engga
                Dim LateBreakIn As Double = 0
                If (status <> ClsTrans_Attendance.EnumStatus.Present Or status <> ClsTrans_Attendance.EnumStatus.Half_Holiday) And _
                workday <> ClsTrans_Attendance.EnumWorkday.Holiday Then
                    If ActBreakIn <> 0 And ActBreakOut <> 0 Then
                        If SchBreakOut < SchBreakIn Then
                            If ActBreakIn > SchBreakIn Then
                                LateBreakIn = ActBreakIn - SchBreakIn
                            Else
                                LateBreakIn = 0
                            End If
                        Else
                            LateBreakIn = 0
                        End If
                    Else
                        LateBreakIn = 0
                    End If
                End If
                'cek maksimal late break in
                max = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.MaxValue_Late_Break_In)
                If String.IsNullOrEmpty(max) = True Then
                    Me.dgvItem.Item(Me.ColLateBreakInDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(LateBreakIn)
                Else
                    If max > EarlyIn Then
                        Me.dgvItem.Item(Me.ColLateBreakInDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(LateBreakIn)
                    ElseIf max < EarlyIn Then
                        Me.dgvItem.Item(Me.ColLateBreakInDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(max)
                    End If
                End If
                'is late break in
                If Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColLateBreakInDuration.Index, vIdx).Value) > 0 Then
                    Me.dgvItem.Item(Me.ColIsLateBreakIn.Index, vIdx).Value = 1
                Else
                    Me.dgvItem.Item(Me.ColIsLateBreakIn.Index, vIdx).Value = 0
                End If
                ''================================================================================================
                'act break duration
                Dim ActBreak As Double = 0
                If ActBreakOut <> 0 And ActBreakIn <> 0 Then
                    If Me.dgvItem.Item(Me.ColActualBreakOut.Index, vIdx).Value <> Me.dgvItem.Item(Me.ColActualBreakIn.Index, vIdx).Value Then
                        If ActBreakOut >= ActBreakIn Then
                            ActBreak = ((24 * 60) - ActBreakOut) + ActBreakIn
                        Else
                            ActBreak = ActBreakIn - ActBreakOut
                        End If
                    End If
                End If
                Me.dgvItem.Item(Me.ColActualBreakDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(ActBreak)
                'act work duration
                Dim ActWork As Double
                If ActOut >= ActIn Then
                    ActWork = ActOut - ActIn
                Else
                    ActWork = ((24 * 60) - ActIn) + ActOut
                End If
                Me.dgvItem.Item(Me.ColActualWorkDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(ActWork - ActBreak)

                'Fungsi permit pindah ke HitungPermit

                'cek tipe ot nya, kalo no ot maka ot dan early ot d kosongin
                CboDgv = Me.dgvItem.Item(Me.ColOT.Index, vIdx)
                If Me.dgvItem.Item(Me.ColOT.Index, vIdx).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumTipeOt.No_Overtime) Then
                    Me.dgvItem.Item(Me.ColOTDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                    Me.dgvItem.Item(Me.ColEarlyOTDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(0)
                End If

                'aul tambah overtime order
                If UpdOvertimeOrder = True Then
                    Dim rowD As DataGridViewRow = Me.dgvItem.Rows(vIdx)
                    If String.IsNullOrEmpty(Me.dgvItem(Me.ColPrimaryOvertimeOrder.Index, vIdx).Value) = False Then
                        If Me.dgvItem(Me.ColPrimaryOvertimeOrder.Index, vIdx).Value <> "-1" Then
                            Me.ObjAtt.UpdateAttBySpl_Detail(Me.TxtPrimaryEmployee.Text, rowD, Me.TblOT_Order, Me, False)
                        End If
                    End If

                    'Dim ApaOverrideOTOrder_Emp As Boolean = ObjEmp.ApakahOverrideOTOrder_Emp(Me.TxtPrimaryEmployee.Text)
                    'Dim ApakahUseOT_Order As Boolean = False
                    ''cek jika pada master employee di override ot order maka use ot mengacu ke employee
                    'If ApaOverrideOTOrder_Emp = True Then
                    '    ApakahUseOT_Order = Me.ObjEmp.ApakahUseOTOrder_Emp(Me.TxtPrimaryEmployee.Text)
                    'Else
                    '    ApakahUseOT_Order = Me.ObjWS.ApakahUseOTOrder_WS(Me.dgvItem.Item(Me.ColPrimaryWorkShift.Index, vIdx).Value)
                    'End If
                    'If ApakahUseOT_Order = True Then
                    '    Me.ObjAtt.UpdateAttBySpl_Detail(Me.TxtPrimaryEmployee.Text, rowD, Me.TblOT_Order, Me, False)
                    'End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub HitungPermit(ByVal vIdx As Integer, ByVal vColIdx As Integer)
    '    Dim SchIn As Double = 0
    '    Dim SchBreakOut As Double = 0
    '    Dim SchBreakIn As Double = 0
    '    Dim SchOut As Double = 0
    '    Dim SchBreak As Double
    '    Dim PermitFrom As Double = 0
    '    Dim PermitTo As Double = 0
    '    Dim ApaLewatBatasan As Boolean = True

    '    Try
    '        SchIn = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColSchIn.Index, vIdx).Value)
    '        SchBreakIn = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColSchBreakIn.Index, vIdx).Value)
    '        SchBreakOut = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColSchBreakOut.Index, vIdx).Value)
    '        SchOut = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColSchOut.Index, vIdx).Value)
    '        PermitFrom = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColPermitOut.Index, vIdx).Value)
    '        PermitTo = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColPermitIn.Index, vIdx).Value)

    '        If SchBreakOut >= SchBreakIn Then
    '            'Untuk Jam Malam
    '            SchBreak = ((24 * 60) - SchBreakOut) + SchBreakIn
    '        Else
    '            SchBreak = SchBreakIn - SchBreakOut
    '        End If

    '        'permit durasi
    '        If SchBreakOut > SchBreakIn Then
    '            'Untuk Jam Malam
    '            If PermitFrom < SchOut And PermitTo < SchOut Then
    '                If ((PermitFrom + 1440) >= SchBreakOut And (PermitFrom + 1440) <= (SchBreakIn + 1440)) And ((PermitTo + 1440) >= SchBreakOut And (PermitTo + 1440) <= (SchBreakIn + 1440)) Then
    '                    ApaLewatBatasan = False
    '                End If
    '            ElseIf PermitFrom < SchOut Then
    '                If ((PermitFrom + 1440) >= SchBreakOut And (PermitFrom + 1440) <= (SchBreakIn + 1440)) And (PermitTo >= SchBreakOut And PermitTo <= (SchBreakIn + 1440)) Then
    '                    ApaLewatBatasan = False
    '                End If
    '            ElseIf PermitTo < SchOut Then
    '                If (PermitFrom >= SchBreakOut And PermitFrom <= (SchBreakIn + 1440)) And ((PermitTo + 1440) >= SchBreakOut And (PermitTo + 1440) <= (SchBreakIn + 1440)) Then
    '                    ApaLewatBatasan = False
    '                End If
    '            Else
    '                If (PermitFrom >= SchBreakOut And PermitFrom <= (SchBreakIn + 1440)) And (PermitTo >= SchBreakOut And PermitTo <= (SchBreakIn + 1440)) Then
    '                    ApaLewatBatasan = False
    '                End If
    '            End If
    '        Else
    '            If (PermitFrom >= SchBreakOut And PermitFrom <= SchBreakIn) And (PermitTo >= SchBreakOut And PermitTo <= SchBreakIn) Then
    '                ApaLewatBatasan = False
    '            End If
    '        End If

    '        If ApaLewatBatasan = False And PermitFrom <> 0 And PermitTo <> 0 Then
    '            'jika sama dengan jam istirhat maka set 0 dan kasih tau jangan ijin di jam istirahat
    '            MsgBox("Can't permit in break duration.", MsgBoxStyle.Information, "Information")
    '            If vColIdx = Me.ColPermitOut.Index Then
    '                Me.dgvItem.Item(Me.ColPermitOut.Index, vIdx).Value = Me.Buffer
    '            ElseIf vColIdx = Me.ColPermitIn.Index Then
    '                Me.dgvItem.Item(Me.ColPermitIn.Index, vIdx).Value = Me.Buffer
    '            End If
    '            'Me.dgvItem.Item(Me.ColPermitDuration.Index, vIdx).Value = "00.00"
    '            Exit Sub
    '        End If

    '        'aturan hitung :
    '        'jika melewati jam istirahat maka kurang kan dengan jam istirahat yang terpakai
    '        If Me.dgvItem.Item(Me.ColPermit.Index, vIdx).Value = True Then
    '            If SchBreakOut > SchBreakIn Then
    '                'Untuk Jam Malam
    '                If PermitFrom < SchOut Then
    '                    PermitFrom += (24 * 60)
    '                End If
    '                If SchBreakOut < SchOut Then
    '                    SchBreakOut += (24 * 60)
    '                End If
    '                If SchBreakIn < SchOut Then
    '                    SchBreakIn += (24 * 60)
    '                End If
    '                If PermitTo < SchOut Then
    '                    PermitTo += 24 * 60
    '                End If
    '            End If

    '            If PermitFrom <= SchBreakOut And (PermitTo <= SchBreakIn And PermitTo >= SchBreakOut) Then
    '                Me.dgvItem.Item(Me.ColPermitDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(SchBreakOut - PermitFrom)
    '                Me.dgvItem.Item(Me.ColIsEarlyBreakOut.Index, vIdx).Value = 0
    '            ElseIf (PermitFrom >= SchBreakOut And PermitFrom <= SchBreakIn) And PermitTo >= SchBreakIn Then
    '                Me.dgvItem.Item(Me.ColPermitDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(PermitTo - SchBreakIn)
    '                Me.dgvItem.Item(Me.ColIsLateBreakIn.Index, vIdx).Value = 0
    '            ElseIf PermitFrom <= SchBreakOut And PermitTo >= SchBreakIn Then
    '                Me.dgvItem.Item(Me.ColPermitDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString((PermitTo - PermitFrom) - (SchBreak))
    '                Me.dgvItem.Item(Me.ColIsEarlyBreakOut.Index, vIdx).Value = 0
    '                Me.dgvItem.Item(Me.ColIsLateBreakIn.Index, vIdx).Value = 0
    '            Else
    '                Me.dgvItem.Item(Me.ColPermitDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(PermitTo - PermitFrom)
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub HitungPermit(ByVal vIdx As Integer, ByVal vColIdx As Integer)
        Dim SchIn As Double = 0
        Dim SchBreakOut As Double = 0
        Dim SchBreakIn As Double = 0
        Dim SchOut As Double = 0
        Dim SchBreak As Double
        Dim PermitFrom As Double = 0
        Dim PermitTo As Double = 0
        Dim ApaLewatBatasan As Boolean = True

        Try
            SchIn = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColSchIn.Index, vIdx).Value)
            SchBreakIn = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColSchBreakIn.Index, vIdx).Value)
            SchBreakOut = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColSchBreakOut.Index, vIdx).Value)
            SchOut = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColSchOut.Index, vIdx).Value)
            PermitFrom = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColPermitOut.Index, vIdx).Value)
            PermitTo = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColPermitIn.Index, vIdx).Value)

            If SchBreakOut >= SchBreakIn Then
                'Untuk Jam Malam
                SchBreak = ((24 * 60) - SchBreakOut) + SchBreakIn
            Else
                SchBreak = SchBreakIn - SchBreakOut
            End If

            'tambahan

            'permit durasi
            If SchBreakOut > SchBreakIn Then
                'Untuk Jam Malam
                If PermitFrom < SchOut And PermitTo < SchOut Then
                    If ((PermitFrom + 1440) >= SchBreakOut And (PermitFrom + 1440) <= (SchBreakIn + 1440)) And ((PermitTo + 1440) >= SchBreakOut And (PermitTo + 1440) <= (SchBreakIn + 1440)) Then
                        ApaLewatBatasan = False
                    End If
                ElseIf PermitFrom < SchOut Then
                    If ((PermitFrom + 1440) >= SchBreakOut And (PermitFrom + 1440) <= (SchBreakIn + 1440)) And (PermitTo >= SchBreakOut And PermitTo <= (SchBreakIn + 1440)) Then
                        ApaLewatBatasan = False
                    End If
                ElseIf PermitTo < SchOut Then
                    If (PermitFrom >= SchBreakOut And PermitFrom <= (SchBreakIn + 1440)) And ((PermitTo + 1440) >= SchBreakOut And (PermitTo + 1440) <= (SchBreakIn + 1440)) Then
                        ApaLewatBatasan = False
                    End If
                Else
                    If (PermitFrom >= SchBreakOut And PermitFrom <= (SchBreakIn + 1440)) And (PermitTo >= SchBreakOut And PermitTo <= (SchBreakIn + 1440)) Then
                        ApaLewatBatasan = False
                    End If
                End If
            Else
                If (PermitFrom >= SchBreakOut And PermitFrom <= SchBreakIn) And (PermitTo >= SchBreakOut And PermitTo <= SchBreakIn) Then
                    ApaLewatBatasan = False
                End If
            End If

            If ApaLewatBatasan = False And PermitFrom <> 0 And PermitTo <> 0 Then
                'jika sama dengan jam istirhat maka set 0 dan kasih tau jangan ijin di jam istirahat
                MsgBox("Can't permit in break duration.", MsgBoxStyle.Information, "Information")
                If vColIdx = Me.ColPermitOut.Index Then
                    Me.dgvItem.Item(Me.ColPermitOut.Index, vIdx).Value = Me.Buffer
                ElseIf vColIdx = Me.ColPermitIn.Index Then
                    Me.dgvItem.Item(Me.ColPermitIn.Index, vIdx).Value = Me.Buffer
                End If
                'Me.dgvItem.Item(Me.ColPermitDuration.Index, vIdx).Value = "00.00"
                Exit Sub
            End If

            'aturan hitung :
            'jika melewati jam istirahat maka kurang kan dengan jam istirahat yang terpakai
            If Me.dgvItem.Item(Me.ColPermit.Index, vIdx).Value = True Then
                If SchBreakOut > SchBreakIn Or SchBreakOut < SchIn Or SchOut < SchIn Then 'tambahan
                    'Untuk Jam Malam
                    If PermitFrom < SchOut Then
                        PermitFrom += (24 * 60)
                    End If
                    If SchBreakOut < SchOut Then
                        SchBreakOut += (24 * 60)
                    End If
                    If SchBreakIn < SchOut Then
                        SchBreakIn += (24 * 60)
                    End If
                    If PermitTo <= SchOut Then
                        PermitTo += 24 * 60
                    End If
                End If

                If PermitFrom <= SchBreakOut And (PermitTo <= SchBreakIn And PermitTo >= SchBreakOut) Then
                    Me.dgvItem.Item(Me.ColPermitDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(SchBreakOut - PermitFrom)
                    Me.dgvItem.Item(Me.ColIsEarlyBreakOut.Index, vIdx).Value = 0
                ElseIf (PermitFrom >= SchBreakOut And PermitFrom <= SchBreakIn) And PermitTo >= SchBreakIn Then
                    Me.dgvItem.Item(Me.ColPermitDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(PermitTo - SchBreakIn)
                    Me.dgvItem.Item(Me.ColIsLateBreakIn.Index, vIdx).Value = 0
                ElseIf PermitFrom <= SchBreakOut And PermitTo >= SchBreakIn Then
                    Me.dgvItem.Item(Me.ColPermitDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString((PermitTo - PermitFrom) - (SchBreak))
                    Me.dgvItem.Item(Me.ColIsEarlyBreakOut.Index, vIdx).Value = 0
                    Me.dgvItem.Item(Me.ColIsLateBreakIn.Index, vIdx).Value = 0
                Else
                    Me.dgvItem.Item(Me.ColPermitDuration.Index, vIdx).Value = Me.ObjCd.Convert_ToString(PermitTo - PermitFrom)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CekVisibleCol_OTOrder(Optional ByVal VisibleStatus As Boolean = False)
        Dim rowdgv As DataGridViewRow
        Dim ShowColOTOrder As Boolean = False

        Try
            For Each rowdgv In Me.dgvItem.Rows
                If String.IsNullOrEmpty(rowdgv.Cells(Me.ColPrimaryOvertimeOrder.Index).Value) = False Then
                    If rowdgv.Cells(Me.ColPrimaryOvertimeOrder.Index).Value <> -1 Then
                        ShowColOTOrder = True
                        Exit For
                    End If
                End If
            Next

            If ShowColOTOrder = True Then
                Me.ColOvertimeOrderNo.Visible = True
                If Me.ColEarlyOTDuration.Visible = True Then
                    Me.ColEarlyOTOrderDuration.Visible = True
                Else
                    Me.ColEarlyOTOrderDuration.Visible = False
                End If
                If Me.ColOTDuration.Visible = True Then
                    Me.ColOTOrderDuration.Visible = True
                Else
                    Me.ColOTOrderDuration.Visible = False
                End If
                If VisibleStatus = True Then
                    Me.ColStatusRemark.Visible = True
                End If
            Else
                Me.ColOvertimeOrderNo.Visible = False
                Me.ColEarlyOTOrderDuration.Visible = False
                Me.ColOTOrderDuration.Visible = False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CekVisibleDGV()
        Dim rowdgv As DataGridViewRow
        Dim ShowColLeave As Boolean = False

        Try
            For Each rowdgv In Me.dgvItem.Rows
                'If rowdgv.Index > 0 Then
                If String.IsNullOrEmpty(rowdgv.Cells(Me.ColLeaveCode.Index).Value) = False Then
                    Me.ColLeaveCode.Visible = True
                    Me.ColLeaveDesc.Visible = True
                    Me.ColLeaveDoc.Visible = True
                    Me.ColStatusRemark.Visible = True
                    ShowColLeave = True
                Else
                    If ShowColLeave = False Then
                        Me.ColLeaveCode.Visible = False
                        Me.ColLeaveDesc.Visible = False
                        Me.ColLeaveDoc.Visible = False
                        Me.ColStatusRemark.Visible = False
                    End If
                End If
                'End If

                If ShowColLeave = True Then
                    Exit For
                End If
            Next

            'cek juga untuk overtime order
            Me.CekVisibleCol_OTOrder(True)

            For Each rowdgv In Me.dgvItem.Rows
                If rowdgv.Index > 1 Then
                    If rowdgv.Cells(Me.ColWorkshift.Index).Value <> Me.dgvItem.Item(Me.ColWorkshift.Index, rowdgv.Index - 1).Value Then
                        Me.ColWorkshift.Visible = True
                        Exit For
                    End If
                End If
            Next
            Me.SetKolomPermit()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanSemuaIsStatus(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            'is present
            Me.dgvItem.Item(Me.ColIsPresent.Index, e.RowIndex).Value = 0
            'is sick
            Me.dgvItem.Item(Me.ColIsSick.Index, e.RowIndex).Value = 0
            'is leave
            Me.dgvItem.Item(Me.ColIsLeave.Index, e.RowIndex).Value = 0
            'is alfa
            Me.dgvItem.Item(Me.ColIsAlfa.Index, e.RowIndex).Value = 0
            'is permit
            Me.dgvItem.Item(Me.ColIsPermit.Index, e.RowIndex).Value = 0
            'is holiday
            Me.dgvItem.Item(Me.ColIsHoliday.Index, e.RowIndex).Value = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Attendance_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtDocNumber.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObject()
                        If Baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Windows.Forms.Cursors.WaitCursor
                            If Me.ObjAtt.SimpanBaru(Me.UserCode, Me.dgvItem, Me, Me.TblOT_Order) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Windows.Forms.Cursors.WaitCursor
                            If Me.ObjAtt.SimpanPerubahan(Me.UserCode, Me.dgvItem, Me, Me.TblOT_Order) = True Then
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

    Private Sub FrmTrans_Attendance_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            'Me.ColIsPresent.Visible = True
            'Me.ColIsLeave.Visible = True
            'Me.ColIsSick.Visible = True
            'Me.ColIsPermit.Visible = True
            'Me.ColIsAlfa.Visible = True
            'Me.ColIsLate.Visible = True
            'Me.ColIsOvertime.Visible = True
            'Me.ColIsEarlyOut.Visible = True
            'Me.ColIsEarlyOvertime.Visible = True
            'ColPrimaryLA.Visible = True
            'ColPrimaryOvertimeOrder.Visible = True

            Me.btnRepairLeave.Visible = False
            Me.btnRepairLeave.TabStop = False

            Me.ColOnSchedule.Visible = False

            Me.dgvItem.AllowUserToResizeColumns = False
            Me.dgvItem.AllowUserToResizeRows = False
            Me.dgvItem.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            Me.dgvItem.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Me.SplitContainer1.TabStop = False
            Me.SplitContainer1.SplitterDistance = Me.SplitContainer1.Width / 2
            Me.BtnFindEmployeeCode.Left = Me.SplitContainer1.Panel1.Width - 5 - Me.BtnFindEmployeeCode.Width
            Me.BtnReload.Left = Me.SplitContainer1.Panel2.Width - 10 - Me.BtnReload.Width
            Me.TxtDocNumber.Width = Me.BtnFindEmployeeCode.Right - Me.TxtDocNumber.Left
            Me.DtpDate.Width = Me.BtnFindEmployeeCode.Right - Me.TxtDocNumber.Left
            Me.TxtEmployeeCode.Width = Me.BtnFindEmployeeCode.Left - 5 - Me.TxtDocNumber.Left
            Me.TxtEmployeeName.Width = Me.BtnFindEmployeeCode.Right - Me.TxtDocNumber.Left
            Me.TxtPosition.Width = Me.BtnFindEmployeeCode.Right - Me.TxtDocNumber.Left
            Me.dtpStart.Width = Me.BtnReload.Right - Me.dtpStart.Left
            Me.DtpEnd.Width = Me.BtnChange.Left - 5 - Me.dtpStart.Left
            Me.TxtRemarks.Width = Me.BtnReload.Right - Me.dtpStart.Left
            Me.BtnChange.Left = Me.BtnReload.Left - Me.BtnChange.Width - 5
            Me.DtpEnd.Width = Me.BtnChange.Left - 5 - Me.DtpEnd.Left

            Me.Text = "Attendance List"
            Me.ShowIcon = False

            If Me.DariQuick = False Then
                System.Windows.Forms.Form.CheckForIllegalCrossThreadCalls = True
                Me.NonAktifkanDetail()
                Me.KosongKanField()
                Me.BtnStop.Visible = False
                Me.BtnStop.TabStop = False
                Me.lblProgress.Visible = False
                Me.ProgressBar1.Visible = False
            Else
                If Me.Baru = True Then
                    ''Me.btnNew.PerformClick()
                    ConfirmBox = False
                    Me.BtnReload.PerformClick()
                Else
                    Me.btnEdit.PerformClick()
                    Me.BtnStop.Visible = False
                    Me.BtnStop.TabStop = False
                    Me.lblProgress.Visible = False
                    Me.ProgressBar1.Visible = False
                End If
            End If

            If Me.IsImportFromMachine = False Then
                Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            End If

            Me.RoundingUnit = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.OT_Duration_RoundingUnit)
            Me.RoundingType = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.OT_Duration_RoundingType)
            Me.DailyMaxOT = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.MaxValue_OT_Daily)
            Me.PeriodMaxOt = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.MaxValue_OT_Period)
            Me.RoundingUnit_EarlyOt = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Early_OT_Duration_RoundingUnit)
            Me.RoundingType_EarlyOt = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Early_OT_Duration_RoundingType)
            Me.DailyMaxOT_EarlyOt = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.MaxValue_Early_OT_Daily)
            Me.PeriodMaxOt_EarlyOt = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.MaxValue_Early_OT_Period)
            Me.RoundingTypeLateIn = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Late_Duration_RoundingType)
            Me.RoundingUnitLateIn = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Late_Duration_RoundingUnit)
            Me.SettingKolom()

            'optimasi edit dari quick lambat
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            For Each Col1 As DataGridViewColumn In Me.dgvItem.Columns
                Col1.SortMode = DataGridViewColumnSortMode.NotSortable
            Next
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            If Me.ShowAtt_FromQuick = True Then
                Me.Show_Attendance(Me.PrimaryAtt_FromQuick)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub Show_Attendance(ByVal vPrimaryAtt As Double)

        Try
            If Not vPrimaryAtt = -1 Then
                Me.Cursor = Cursors.WaitCursor
                If Me.ObjAtt.CariDetail(vPrimaryAtt, Me, Me.TblOT_Order, Me.JoinWithAttPrev) = True Then
                    Me.SalinKeField()
                    Me.NonAktifkanDetail()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub DtpDate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpDate.Enter
        Try
            Me.BufferTransAttendanceDate = Me.DtpDate.Value.Date
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpDate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpDate.Leave
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.BackgroundWorker1.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindEmployeeCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindEmployeeCode.Click
        Dim vFrm As New Core_Module.FrmMasterKaryawanCari
        Try
            vFrm.txtCari.Text = Me.TxtEmployeeCode.Text
            vFrm.BackColorAll = Me.BackColorAll
            vFrm.ShowDialog()
            If vFrm.PrimaryEmploye > -1 Then
                Me.TxtPrimaryEmployee.Text = vFrm.PrimaryEmploye
                Me.TxtEmployeeCode.Text = Me.ObjEmp.GetCodeByPrimary(vFrm.PrimaryEmploye)
                Me.TxtEmployeeName.Text = Me.ObjEmp.GetFullNameByPrimary(vFrm.PrimaryEmploye)
                Me.TxtPrimaryPosition.Text = Me.ObjEmp.GetJobByPrimary(vFrm.PrimaryEmploye)
                Me.TxtPosition.Text = Me.ObjJob.GetDesc(Me.TxtPrimaryPosition.Text)
            Else
                Me.TxtPrimaryEmployee.Text = -1
                Me.TxtEmployeeCode.Text = ""
                Me.TxtEmployeeName.Text = ""
                Me.TxtPrimaryPosition.Text = -1
                Me.TxtPosition.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            vFrm.Dispose()
        End Try
    End Sub

    Private Sub TxtEmployeeCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmployeeCode.GotFocus
        Me.TxtEmployeeCode.BackColor = Me.BackColorAll
        If IsNumeric(Me.TxtPrimaryEmployee.Text) Then
            Me.BufferEmployee = CDbl(Me.TxtPrimaryEmployee.Text)
        End If
    End Sub

    Private Sub TxtEmployeeCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmployeeCode.LostFocus
        Me.TxtEmployeeCode.BackColor = Drawing.Color.White
    End Sub

    Private Sub TxtEmployeeCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtEmployeeCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                Me.TxtEmployeeCode_Leave(sender, e)
                Me.dtpStart.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtEmployeeCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmployeeCode.Leave
        Try
            If Me.TxtEmployeeCode.Text <> "" Then
                If Me.ObjEmp.ApakahSudahAda(Me.TxtEmployeeCode.Text) = True Then
                    Me.TxtPrimaryEmployee.Text = Me.ObjEmp.GetPrimaryByCode(Me.TxtEmployeeCode.Text)
                    Me.TxtEmployeeCode.Text = Me.ObjEmp.GetCodeByPrimary(Me.TxtPrimaryEmployee.Text)
                    Me.TxtEmployeeName.Text = Me.ObjEmp.GetFullNameByPrimary(Me.TxtPrimaryEmployee.Text)
                    Me.TxtPrimaryPosition.Text = Me.ObjEmp.GetJobByPrimary(Me.TxtPrimaryEmployee.Text)
                    Me.TxtPosition.Text = Me.ObjJob.GetDesc(Me.TxtPrimaryPosition.Text)
                Else
                    Me.BtnFindEmployeeCode.PerformClick()
                End If
            Else
                Me.TxtPrimaryEmployee.Text = -1
                Me.TxtEmployeeCode.Text = ""
                Me.TxtEmployeeName.Text = ""
                Me.TxtPrimaryPosition.Text = -1
                Me.TxtPosition.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarks_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus
        Me.TxtRemarks.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtRemarks_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus
        Me.TxtRemarks.BackColor = Drawing.Color.White
    End Sub

    Private Sub dtpStart_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpStart.Enter
        Me.BufferDate = Me.dtpStart.Value
    End Sub

    Private Sub DtpEnd_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpEnd.Enter
        Me.BufferDate = Me.DtpEnd.Value
    End Sub

    Private Sub BtnReload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnReload.Click
        Dim Hasil As New MsgBoxResult

        Try
            If ConfirmBox = False Then
                GoTo SkipConfir
            End If

            Hasil = MsgBox("Do you want to reload transaction attendance?", MsgBoxStyle.YesNo, "Confirm")
            If Hasil = MsgBoxResult.Yes Then
SkipConfir:
                Me.ReloadOrChangeEdit()
                Me.ApakahSudahReload = True

                If Me.JoinWithAttPrev = True Then
                    Me.dtpStart.Enabled = False
                    Me.BtnReload.Enabled = False
                    Me.BtnChange.Enabled = True
                Else
                    Me.dtpStart.Enabled = True
                    Me.BtnReload.Enabled = True
                    Me.BtnChange.Enabled = False
                End If
            End If
            ConfirmBox = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnChange_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnChange.Click
        Dim Hasil As New MsgBoxResult

        Try
            Hasil = MsgBox("Do you want to change period transaction attendance?", MsgBoxStyle.YesNo, "Confirm")
            If Hasil = MsgBoxResult.Yes Then
                Me.ReloadOrChangeEdit()
            End If
            ConfirmBox = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub ReloadOrChangeEdit()
        Dim DtpStartAttNew As Date = Today

        Try
            If Me.TxtEmployeeCode.Text = "" Then
                MsgBox("Please input employee code.", MsgBoxStyle.Information, "Information")
                Me.TxtEmployeeCode.Focus()
                Exit Sub
            End If

            If Me.ObjEmp.ApakahAktif(Me.TxtPrimaryEmployee.Text) = False Then
                MsgBox("Inactive employee, please change your employee code", MsgBoxStyle.Information, "Information")
                Me.TxtEmployeeCode.Focus()
                Exit Sub
            End If

            If Me.dtpStart.Value.Date > Me.DtpEnd.Value.Date Then
                MsgBox("Start date should not be more than end date", MsgBoxStyle.Information, "Error")
                Me.dtpStart.Value = Me.BufferDate
                Me.dtpStart.Focus()
            Else
                If Me.ObjAtt.CekValidPeriodEmp(Me.TxtPrimary.Text, Me.TxtPrimaryEmployee.Text, Me.dtpStart.Value.Date, Me.DtpEnd.Value.Date, Me.Baru) = True Then
                    Me.Cursor = Windows.Forms.Cursors.WaitCursor
                    Me.RefreshEmployeeAttendaceSettings()

                    'aul tambahan
                    Me.dgvItem.Rows.Clear()
                    If Me.IsContinueFromPrevAtt = True Then
                        Me.JoinWithAttPrev = Me.GetPrevAttendance()
                    End If
                    'IsiTblOvertimeOrder
                    Me.ObjAtt.IsiTblSpl(Me.JoinWithAttPrev, Me.TxtPrimaryEmployee.Text, Me.DtpStartAttPrev, Me.dtpStart.Value.Date, Me.DtpEnd.Value.Date, Me.TblOT_Order)
                    Me.ObjAtt.ReloadData(Me, Me.JoinWithAttPrev, Me.TblOT_Order, DtpStartAttNew)

                    ''reset actual data from machine
                    If Me.DariQuick = True And Me.IsImportFromMachine = True Then
                        Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
                        Me.FillAttendanceFromMachine(False)
                        Me.ColOnSchedule.Visible = True
                    End If
                    'Me.WindowState = Me.StartingWS

                    'cek salah penempatan clock absen terhadap actual in/actual out
                    Me.SetActualInOut_ByClock(DtpStartAttNew)

                    'aul tambah overtime order
                    Me.ObjAtt.UpdateAttBySpl(Me.TxtPrimaryEmployee.Text, DtpStartAttNew.Date, Me.DtpEnd.Value.Date, Me.TblOT_Order, Me)
                    Me.CekVisibleDGV()

                    Me.Cursor = Windows.Forms.Cursors.Default

                    Me.CekLompatHari()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            If Me.JoinWithAttPrev = True And Me.dgvItem.RowCount > 0 Then
                Me.dtpStart.Value = CDate(Me.dgvItem.Item(Me.ColDate.Index, 0).Value).Date
                Me.DtpEnd.Value = CDate(Me.dgvItem.Item(Me.ColDate.Index, Me.dgvItem.RowCount - 1).Value).Date
            End If
        End Try
    End Sub

    Private Sub SetActualInOut_ByClock(ByVal DtpStartAttNew As Date)
        Dim rd, rd_next As DataGridViewRow
        Dim CboDgv As Windows.Forms.DataGridViewComboBoxCell
        Dim workday As ClsTrans_Attendance.EnumWorkday
        Dim workdayNext As ClsTrans_Attendance.EnumWorkday
        Dim Status As ClsTrans_Attendance.EnumStatus
        Dim GetDate As String = ""
        Dim GetTime, GetTimeIn, GetTimeOut As String
        Dim GetDateTime, GetDateTimeIn, GetDateTimeOut As String
        Dim KosongkanActualIn As Boolean = False
        Dim KosongkanActualOut As Boolean = False
        Dim IsHoliday As Boolean = False
        Dim IdxRowUbah As Integer = 0

        Try
            For Each rd In Me.dgvItem.Rows
                If CDate(rd.Cells(Me.ColDate.Index).Value) >= DtpStartAttNew Then
                    KosongkanActualIn = False
                    KosongkanActualOut = False
                    IsHoliday = False
                    IdxRowUbah = rd.Index
                    GetTime = ""
                    GetTimeIn = ""
                    GetTimeOut = ""
                    GetDateTime = ""
                    GetDateTimeIn = ""
                    GetDateTimeOut = ""

                    If rd.Index < Me.dgvItem.RowCount - 1 Then
                        rd_next = Me.dgvItem.Rows(rd.Index + 1)

                        'workday
                        CboDgv = rd.Cells(Me.ColWorkdays.Index)
                        workday = CboDgv.Items.IndexOf(rd.Cells(Me.ColWorkdays.Index).Value)

                        If workday = ClsTrans_Attendance.EnumWorkday.Holiday Then
                            IsHoliday = True
                            'JIKA HOLIDAY

                            GetDate = Me.dgvItem.Item(Me.ColDate.Index, rd_next.Index).Value
                            GetTimeIn = Me.dgvItem.Item(Me.ColActualIn.Index, rd.Index).Value.ToString.Replace(".", ":")
                            GetDateTimeIn = GetDate & " " & GetTimeIn
                            GetTimeOut = Me.dgvItem.Item(Me.ColActualOut.Index, rd.Index).Value.ToString.Replace(".", ":")
                            GetDateTimeOut = GetDate & " " & GetTimeOut

                            CboDgv = rd_next.Cells(Me.ColWorkdays.Index)
                            workdayNext = CboDgv.Items.IndexOf(rd_next.Cells(Me.ColWorkdays.Index).Value)
                            If workdayNext = ClsTrans_Attendance.EnumWorkday.Holiday Then
                                'jika hari selanjutnya hari libur maka hari selanjutnya yang mengalah
                                If String.IsNullOrEmpty(Me.dgvItem.Item(Me.ColLogContents.Index, rd_next.Index).Value) = False Then
                                    'cek actual in 
                                    If Me.dgvItem.Item(Me.ColLogContents.Index, rd_next.Index).Value.ToString.Contains(GetDateTimeIn) = True _
                                    Or Me.dgvItem.Item(Me.ColLogContents.Index, rd_next.Index).Value.ToString.Contains(GetDateTimeOut) = True Then
                                        KosongkanActualIn = True
                                    End If
                                End If

                                If Me.dgvItem.Item(Me.ColActualOut.Index, rd.Index).Value = Me.dgvItem.Item(Me.ColActualIn.Index, rd_next.Index).Value Then
                                    If Me.dgvItem.Item(Me.ColActualOut.Index, rd.Index).Value = "-" Or Me.dgvItem.Item(Me.ColActualOut.Index, rd.Index).Value = "00.00" Then
                                    Else
                                        KosongkanActualIn = True
                                    End If
                                End If
                                IdxRowUbah = rd_next.Index
                            Else
                                'jika hari selanjutnya hari kerja maka hari yang di looping yang mengalah
                                If String.IsNullOrEmpty(Me.dgvItem.Item(Me.ColLogContents.Index, rd_next.Index).Value) = False Then
                                    'cek actual in 
                                    If Me.dgvItem.Item(Me.ColLogContents.Index, rd_next.Index).Value.ToString.Contains(GetDateTimeIn) = True _
                                    Or Me.dgvItem.Item(Me.ColLogContents.Index, rd_next.Index).Value.ToString.Contains(GetDateTimeOut) = True Then
                                        KosongkanActualOut = True
                                    End If
                                End If

                                If Me.dgvItem.Item(Me.ColActualOut.Index, rd.Index).Value = Me.dgvItem.Item(Me.ColActualIn.Index, rd_next.Index).Value Then
                                    If Me.dgvItem.Item(Me.ColActualOut.Index, rd.Index).Value = "-" Or Me.dgvItem.Item(Me.ColActualOut.Index, rd.Index).Value = "00.00" Then
                                    Else
                                        KosongkanActualOut = True
                                    End If
                                End If
                                IdxRowUbah = rd.Index
                            End If

                            If KosongkanActualIn = True Or KosongkanActualOut = True Then
                                If String.IsNullOrEmpty(Me.dgvItem.Item(Me.ColLogContents.Index, IdxRowUbah).Value) = False Then
                                    Me.dgvItem.Item(Me.ColLogContents.Index, IdxRowUbah).Value = Me.dgvItem.Item(Me.ColLogContents.Index, IdxRowUbah).Value.ToString.Replace(GetDateTimeIn, "")
                                    Me.dgvItem.Item(Me.ColLogContents.Index, IdxRowUbah).Value = Me.dgvItem.Item(Me.ColLogContents.Index, IdxRowUbah).Value.ToString.Replace(GetDateTimeOut, "")

                                    If KosongkanActualIn = True Then
                                        Me.dgvItem.Item(Me.ColActualIn.Index, IdxRowUbah).Value = "-"
                                    ElseIf KosongkanActualOut Then
                                        Me.dgvItem.Item(Me.ColActualOut.Index, IdxRowUbah).Value = "-"
                                    End If
                                End If
                            End If
                        Else
                            'JIKA WEEKDAY
                            'cek : jika hari berikutnya hari libur maka yang mengalah hari berikutnya
                            CboDgv = rd_next.Cells(Me.ColWorkdays.Index)
                            workdayNext = CboDgv.Items.IndexOf(rd_next.Cells(Me.ColWorkdays.Index).Value)
                            If workdayNext = ClsTrans_Attendance.EnumWorkday.Holiday Then
                                IsHoliday = True
                                IdxRowUbah = rd_next.Index

                                'cek actual in 
                                GetDate = rd_next.Cells(Me.ColDate.Index).Value
                                GetTime = rd.Cells(Me.ColActualOut.Index).Value.ToString.Replace(".", ":")
                                GetDateTime = GetDate & " " & GetTime
                                If String.IsNullOrEmpty(Me.dgvItem.Item(Me.ColLogContents.Index, IdxRowUbah).Value) = False Then
                                    If Me.dgvItem.Item(Me.ColLogContents.Index, IdxRowUbah).Value.ToString.Contains(GetDateTime) = True Then
                                        Me.dgvItem.Item(Me.ColLogContents.Index, IdxRowUbah).Value = Me.dgvItem.Item(Me.ColLogContents.Index, IdxRowUbah).Value.ToString.Replace(GetDateTime, "")
                                        Me.dgvItem.Item(Me.ColActualIn.Index, IdxRowUbah).Value = Me.ObjCd.Convert_ToString(0)
                                        KosongkanActualIn = True
                                    End If
                                End If

                                If Me.dgvItem.Item(Me.ColActualOut.Index, rd.Index).Value = Me.dgvItem.Item(Me.ColActualIn.Index, rd_next.Index).Value Then
                                    If Me.dgvItem.Item(Me.ColActualOut.Index, rd.Index).Value = "-" Or Me.dgvItem.Item(Me.ColActualOut.Index, rd.Index).Value = "00.00" Then
                                    Else
                                        KosongkanActualIn = True
                                    End If
                                End If
                            Else
                                If Me.ObjCd.Is_Time_String(rd.Cells(Me.ColActualOut.Index).Value) = True And Me.ObjCd.Is_Time_String(rd_next.Cells(Me.ColActualIn.Index).Value) = True Then
                                    If rd.Cells(Me.ColActualOut.Index).Value = rd_next.Cells(Me.ColActualIn.Index).Value Then
                                        If rd.Cells(Me.ColActualOut.Index).Value = "-" Or rd.Cells(Me.ColActualOut.Index).Value = "00.00" Then
                                        Else
                                            'jika hari berikutnya hari kerja maka yang mengalah hari yang di looping nya
                                            rd.Cells(Me.ColActualOut.Index).Value = "-"

                                            If String.IsNullOrEmpty(rd.Cells(Me.ColLogContents.Index).Value) = False Then
                                                GetDate = rd_next.Cells(Me.ColDate.Index).Value
                                                GetTime = rd_next.Cells(Me.ColActualIn.Index).Value.ToString.Replace(".", ":")
                                                GetDateTime = GetDate & " " & GetTime
                                                rd.Cells(Me.ColLogContents.Index).Value = rd.Cells(Me.ColLogContents.Index).Value.ToString.Replace(GetDateTime, "")
                                            End If

                                            rd.Cells(Me.ColOnSchedule.Index).Value = "On Schedule"
                                            rd.Cells(Me.ColRemarks.Index).Value = ""
                                            rd.DefaultCellStyle.BackColor = Color.White
                                            Me.Hitung(rd.Index, 0, EnumCalculateRange.BOTH, False)

                                            Me.MarkRow(rd.Index, ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_OUT, Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM)

                                        End If
                                    End If
                                End If
                            End If
                        End If

                        CboDgv = Me.dgvItem.Item(Me.ColAttendanceStatus.Index, IdxRowUbah)
                        Status = CboDgv.Items.IndexOf(Me.dgvItem.Item(Me.ColAttendanceStatus.Index, IdxRowUbah).Value)
                        If Status = ClsTrans_Attendance.EnumStatus.Leave Or Status = ClsTrans_Attendance.EnumStatus.Half_day_Leave Then
                        Else
                            If Me.dgvItem.Item(Me.ColActualIn.Index, IdxRowUbah).Value = "-" Or Me.dgvItem.Item(Me.ColActualIn.Index, IdxRowUbah).Value = "00.00" Then
                                Me.dgvItem.Item(Me.ColActualIn.Index, IdxRowUbah).Value = "-"
                                KosongkanActualIn = True
                            End If
                            If Me.dgvItem.Item(Me.ColActualOut.Index, IdxRowUbah).Value = "-" Or Me.dgvItem.Item(Me.ColActualOut.Index, IdxRowUbah).Value = "00.00" Then
                                Me.dgvItem.Item(Me.ColActualOut.Index, IdxRowUbah).Value = "-"
                                KosongkanActualOut = True
                            End If
                        End If

                        CboDgv = Me.dgvItem.Item(Me.ColWorkdays.Index, IdxRowUbah)
                        workday = CboDgv.Items.IndexOf(Me.dgvItem.Item(Me.ColWorkdays.Index, IdxRowUbah).Value)
                        If workday = ClsTrans_Attendance.EnumWorkday.Holiday Then
                            IsHoliday = True
                        Else
                            IsHoliday = False
                        End If
                        If IsHoliday = True Then
                            If KosongkanActualIn = True Or KosongkanActualOut = True Then
                                Me.dgvItem.Item(Me.ColIsPresent.Index, IdxRowUbah).Value = 0
                                Me.dgvItem.Item(Me.ColIsOutstation.Index, IdxRowUbah).Value = Boolean.FalseString
                                Me.dgvItem.Item(Me.ColIsHoliday.Index, IdxRowUbah).Value = 1
                                Me.dgvItem.Item(Me.ColRemarks.Index, IdxRowUbah).Value = ""
                                Me.dgvItem.Rows(IdxRowUbah).DefaultCellStyle.BackColor = Color.White

                                If KosongkanActualIn = True And KosongkanActualOut = True Then
                                    Me.dgvItem.Item(Me.ColAttendanceStatus.Index, IdxRowUbah).Value = "Holiday"
                                    Me.dgvItem.Item(Me.ColOnSchedule.Index, IdxRowUbah).Value = ""
                                    Me.dgvItem.Item(Me.ColActualIn.Index, IdxRowUbah).Value = Me.ObjCd.Convert_ToString(0)
                                    Me.dgvItem.Item(Me.ColActualOut.Index, IdxRowUbah).Value = Me.ObjCd.Convert_ToString(0)
                                    Me.Hitung(IdxRowUbah, 0, EnumCalculateRange.BOTH, False)
                                Else
                                    Me.dgvItem.Rows(IdxRowUbah).DefaultCellStyle.BackColor = Color.White
                                    Me.dgvItem.Item(Me.ColOnSchedule.Index, IdxRowUbah).Value = "On Schedule"
                                    Me.Hitung(IdxRowUbah, 0, EnumCalculateRange.BOTH, False)
                                    If KosongkanActualIn = True Then
                                        Me.MarkRow(IdxRowUbah, ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_IN, Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM)
                                    End If
                                    If KosongkanActualOut = True Then
                                        Me.MarkRow(IdxRowUbah, ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_OUT, Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM)
                                    End If
                                End If
                            End If
                        ElseIf IsHoliday = False Then
                            If KosongkanActualIn = True Or KosongkanActualOut = True Then
                                'Kosongkan Semua Is Status
                                Me.dgvItem.Item(Me.ColIsPresent.Index, IdxRowUbah).Value = 1
                                Me.dgvItem.Item(Me.ColIsOutstation.Index, IdxRowUbah).Value = Boolean.FalseString
                                Me.dgvItem.Item(Me.ColOnSchedule.Index, IdxRowUbah).Value = "On Schedule"
                                Me.dgvItem.Item(Me.ColRemarks.Index, IdxRowUbah).Value = ""
                                Me.dgvItem.Item(Me.ColIsHoliday.Index, IdxRowUbah).Value = 0
                                Me.dgvItem.Rows(IdxRowUbah).DefaultCellStyle.BackColor = Color.White

                                Me.Hitung(IdxRowUbah, 0, EnumCalculateRange.BOTH, False)
                                If KosongkanActualIn = True Then
                                    Me.MarkRow(IdxRowUbah, ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_IN, Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM)
                                End If
                                If KosongkanActualOut = True Then
                                    Me.MarkRow(IdxRowUbah, ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_OUT, Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM)
                                End If
                            End If
                        End If
                    End If
                End If
            Next

            Me.SetBackColorAlternating()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function CekLompatHari() As Boolean
        Dim ListDateDgv As New List(Of Date)
        Dim ParamDate As Date = Me.dtpStart.Value.Date
        Dim InconsistDate As String = ""

        Try
            For Each rD As DataGridViewRow In Me.dgvItem.Rows
                If ListDateDgv.Contains(CDate(rD.Cells(Me.ColDate.Index).Value)) = False Then
                    ListDateDgv.Add(CDate(rD.Cells(Me.ColDate.Index).Value))
                End If
            Next

            If Me.dgvItem.RowCount = 0 Then
                Return True
            End If

            While ParamDate <= Me.DtpEnd.Value.Date
                If ListDateDgv.Contains(ParamDate) = False Then
                    If InconsistDate = "" Then
                        InconsistDate = vbCrLf & Format(ParamDate, "dd MMM yyyy")
                    Else
                        InconsistDate &= vbCrLf & Format(ParamDate, "dd MMM yyyy")
                    End If
                End If
                ParamDate = ParamDate.AddDays(1)
            End While

            If InconsistDate <> "" Then
                MsgBox("Please allocate workshift first." & InconsistDate, MsgBoxStyle.Information, "Information")
                Me.BtnReload.Focus()
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return True
        End Try
    End Function

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Dim CboDgv As Windows.Forms.DataGridViewComboBoxCell
        Dim Status As ClsTrans_Attendance.EnumStatus
        Dim Durasi As ClsTrans_Attendance.EnumDurasi
        Dim workday As ClsTrans_Attendance.EnumWorkday
        Dim ObjUser As New Core_Module.ClsUser

        Try
            If Me.DtpDate.Enabled = True Then
                If e.ColumnIndex = Me.ColEarlyOTDuration.Index Or e.ColumnIndex = Me.ColOTDuration.Index Then
                    If Me.dgvItem.Item(Me.ColPrimaryOvertimeOrder.Index, e.RowIndex).Value <> -1 Then
                        If ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_Attendance_Spl") = False Then
                            MsgBox("You can't edit this row because refering overtime order.", MsgBoxStyle.Information, "Information")
                            e.Cancel = True
                        End If
                    End If
                End If

                ''hanya bisa edit act in dan act out jika format mereka belum benar
                If Me.ObjCd.Is_Time_String(Me.dgvItem.Item(Me.ColActualIn.Index, e.RowIndex).Value) = False And _
                        Me.ObjCd.Is_Time_String(Me.dgvItem.Item(Me.ColActualOut.Index, e.RowIndex).Value) = False Then
                    If e.ColumnIndex <> Me.ColActualIn.Index And e.ColumnIndex <> Me.ColActualOut.Index And e.ColumnIndex <> Me.ColIsOutstation.Index Then
                        e.Cancel = True
                        Exit Sub
                    End If
                ElseIf Me.ObjCd.Is_Time_String(Me.dgvItem.Item(Me.ColActualIn.Index, e.RowIndex).Value) = False Then
                    If e.ColumnIndex <> Me.ColActualIn.Index And e.ColumnIndex <> Me.ColIsOutstation.Index Then
                        e.Cancel = True
                        Exit Sub
                    End If
                ElseIf Me.ObjCd.Is_Time_String(Me.dgvItem.Item(Me.ColActualOut.Index, e.RowIndex).Value) = False Then
                    If e.ColumnIndex <> Me.ColActualOut.Index And e.ColumnIndex <> Me.ColIsOutstation.Index Then
                        e.Cancel = True
                        Exit Sub
                    End If
                End If

                If (e.ColumnIndex = Me.ColEarlyOTDuration.Index And Me.dgvItem.Item(Me.ColIsCalculateEarlyOT.Index, e.RowIndex).Value = Boolean.FalseString) Or _
                        (e.ColumnIndex = Me.ColLateDuration.Index And Me.dgvItem.Item(Me.ColIsCalculateLate.Index, e.RowIndex).Value = Boolean.FalseString) Or _
                        (e.ColumnIndex = Me.ColEarlyOutDuration.Index And Me.dgvItem.Item(Me.ColIsCalculateEarlyOut.Index, e.RowIndex).Value = Boolean.FalseString) Or _
                        (e.ColumnIndex = Me.ColOTDuration.Index And Me.dgvItem.Item(Me.ColIsCalculateOT.Index, e.RowIndex).Value = Boolean.FalseString) Then
                    e.Cancel = True
                    Exit Sub
                End If

                'workday
                CboDgv = Me.dgvItem.Item(Me.ColWorkdays.Index, e.RowIndex)
                workday = CboDgv.Items.IndexOf(Me.dgvItem.Item(Me.ColWorkdays.Index, e.RowIndex).Value)
                'fullday / half day
                CboDgv = Me.dgvItem.Item(Me.ColDuration.Index, e.RowIndex)
                Durasi = CboDgv.Items.IndexOf(Me.dgvItem.Item(Me.ColDuration.Index, e.RowIndex).Value)
                'status
                CboDgv = Me.dgvItem.Item(Me.ColAttendanceStatus.Index, e.RowIndex)
                Status = CboDgv.Items.IndexOf(Me.dgvItem.Item(Me.ColAttendanceStatus.Index, e.RowIndex).Value)
                Me.BufferStatus = Status

                'aul -penambahan outstation-
                If e.ColumnIndex = Me.ColIsOutstation.Index Then
                    If Status = ClsTrans_Attendance.EnumStatus.Present Or Status = ClsTrans_Attendance.EnumStatus.Half_Absent _
                    Or Status = ClsTrans_Attendance.EnumStatus.Half_day_Leave Or Status = ClsTrans_Attendance.EnumStatus.Half_Day_Permit _
                    Or Status = ClsTrans_Attendance.EnumStatus.Half_Day_Sick Or Status = ClsTrans_Attendance.EnumStatus.Half_Day_Sick _
                    Or Status = ClsTrans_Attendance.EnumStatus.Half_Holiday Then
                        Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                End If

                If workday = ClsTrans_Attendance.EnumWorkday.Workday And Durasi = ClsTrans_Attendance.EnumDurasi.Full_Day Then
                    'workday, fullday
                    'sch in , sch out, act in , act out, act break out, act break in
                    If e.ColumnIndex = Me.ColSchIn.Index Or e.ColumnIndex = Me.ColSchOut.Index Or e.ColumnIndex = Me.ColActualIn.Index _
                    Or e.ColumnIndex = Me.ColActualBreakOut.Index Or e.ColumnIndex = Me.ColActualBreakIn.Index Or _
                    e.ColumnIndex = Me.ColActualOut.Index Then
                        If e.ColumnIndex = Me.ColSchIn.Index Or e.ColumnIndex = Me.ColSchOut.Index Then
                            'statusnya halfday boleh di edit
                            'sch in , sch out,
                            If Me.ObjAtt.CekHalfStatus(Status) = True Then
                                Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                            Else
                                e.Cancel = True
                            End If
                        ElseIf e.ColumnIndex = Me.ColActualIn.Index Or e.ColumnIndex = Me.ColActualBreakOut.Index Or _
                        e.ColumnIndex = Me.ColActualBreakIn.Index Or e.ColumnIndex = Me.ColActualOut.Index Then
                            'statusnya halfday + present boleh di edit
                            'act in , act out, act break out, act break in
                            If Me.ObjAtt.CekHalfStatus(Status, True) = True Then
                                Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                            Else
                                e.Cancel = True
                            End If
                        End If
                    ElseIf e.ColumnIndex = Me.ColStatusRemark.Index Then
                        'status remark
                        If Status = ClsTrans_Attendance.EnumStatus.Present Or Status = ClsTrans_Attendance.EnumStatus.Absent _
                        Or Status = ClsTrans_Attendance.EnumStatus.Half_Absent Or Status = ClsTrans_Attendance.EnumStatus.Holiday _
                        Or Status = ClsTrans_Attendance.EnumStatus.Half_Holiday Then
                            e.Cancel = True
                        Else
                            Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        End If
                    ElseIf e.ColumnIndex = Me.ColStatusRemark.Index Then
                        If Status = ClsTrans_Attendance.EnumStatus.Half_Absent Then
                            e.Cancel = True
                        End If
                    Else
                        Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    If e.ColumnIndex = Me.ColSchIn.Index Or e.ColumnIndex = Me.ColSchBreakOut.Index Or e.ColumnIndex = Me.ColSchBreakIn.Index _
                    Or e.ColumnIndex = Me.ColSchOut.Index Then
                        e.Cancel = True
                    ElseIf e.ColumnIndex = Me.ColActualIn.Index Or e.ColumnIndex = Me.ColActualBreakIn.Index _
                    Or e.ColumnIndex = Me.ColActualBreakOut.Index Or e.ColumnIndex = Me.ColActualOut.Index Then
                        'semua colom actual cek dulu statusnya
                        'yang boleh di edit cuma present & half holiday
                        If Status = ClsTrans_Attendance.EnumStatus.Present Or Status = ClsTrans_Attendance.EnumStatus.Half_Holiday Then
                            Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        Else
                            e.Cancel = True
                        End If
                    Else
                        Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                End If
                'status remarks
                If e.ColumnIndex = Me.ColStatusRemark.Index Then
                    If Status = ClsTrans_Attendance.EnumStatus.Permit Or Status = ClsTrans_Attendance.EnumStatus.Half_Day_Permit Or _
                    Status = ClsTrans_Attendance.EnumStatus.Sick Or Status = ClsTrans_Attendance.EnumStatus.Half_Day_Sick Then
                        Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                End If
                'tipe OT
                If e.ColumnIndex = Me.ColOT.Index Then
                    If Me.ObjAtt.CekHalfStatus(Status, True) = True Then
                        Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                End If
                'berhubungan dengan permit
                If e.ColumnIndex = Me.ColPermit.Index Then
                    If Status = ClsTrans_Attendance.EnumStatus.Present Then
                        Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                End If
                If e.ColumnIndex = Me.ColPermitOut.Index Or e.ColumnIndex = Me.ColPermitIn.Index Then
                    If Me.dgvItem.Item(Me.ColPermit.Index, e.RowIndex).Value = True Then
                        Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                End If

                If e.ColumnIndex = Me.ColIsCalculateEarlyOT.Index Or e.ColumnIndex = Me.ColIsCalculateOT.Index Then
                    CboDgv = Me.dgvItem.Item(Me.ColWorkdays.Index, e.RowIndex)
                    workday = CboDgv.Items.IndexOf(Me.dgvItem.Item(Me.ColWorkdays.Index, e.RowIndex).Value)
                    CboDgv = Me.dgvItem.Item(Me.ColAttendanceStatus.Index, e.RowIndex)
                    Status = CboDgv.Items.IndexOf(Me.dgvItem.Item(Me.ColAttendanceStatus.Index, e.RowIndex).Value)
                    If workday = ClsTrans_Attendance.EnumWorkday.Holiday Then
                        e.Cancel = True
                    Else
                        Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                End If

                If e.ColumnIndex = Me.ColOTOrderDuration.Index Or e.ColumnIndex = Me.ColEarlyOTOrderDuration.Index Or _
                e.ColumnIndex = Me.ColPermitDuration.Index Then
                    e.Cancel = True
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjUser = Nothing
        End Try
    End Sub

    Private Sub dgvItem_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
        Dim FrmLog As New FrmTrans_Attendance_ShowLog
        Try
            If e.RowIndex > -1 Then
                If e.ColumnIndex = Me.ColLeaveDoc.Index Then
                    Dim Frm As New FrmTransaksiCuti_Approval
                    Frm.UserCode = Me.UserCode
                    Frm.ShowDialog()
                    ConfirmBox = False
                    Me.BtnReload.PerformClick()
                ElseIf e.ColumnIndex = Me.ColOvertimeOrderNo.Index Or e.ColumnIndex = Me.ColOTOrderDuration.Index Or e.ColumnIndex = Me.ColEarlyOTOrderDuration.Index Then
                    If Me.dgvItem.Item(Me.ColPrimaryOvertimeOrder.Index, e.RowIndex).Value = -1 Or _
                    String.IsNullOrEmpty(Me.dgvItem.Item(Me.ColPrimaryOvertimeOrder.Index, e.RowIndex).Value) = True Then
                        Exit Sub
                    End If

                    Me.Cursor = Cursors.WaitCursor

                    Dim Frm As New FrmTrans_SPL
                    Frm.UserCode = Me.UserCode
                    Frm.FromAtt = True
                    Frm.ShowSPL(Me.dgvItem.Item(Me.ColPrimaryOvertimeOrder.Index, e.RowIndex).Value)
                    Frm.ShowDialog()

                    Me.Cursor = Cursors.Default
                ElseIf e.ColumnIndex = Me.ColShowLog.Index Then
                    If IsNothing(Me.dgvItem.Item(Me.ColLogContents.Index, e.RowIndex).Value) = False Then
                        If Me.dgvItem.Item(Me.ColLogContents.Index, e.RowIndex).Value.ToString <> "" Then
                            FrmLog.TextBox1.Text = Me.dgvItem.Item(Me.ColLogContents.Index, e.RowIndex).Value.ToString
                            FrmLog.ShowDialog()
                        Else
                            MsgBox("There is no log.", MsgBoxStyle.Information, "Information")
                        End If
                    Else
                        MsgBox("There is no log.", MsgBoxStyle.Information, "Information")
                    End If
                ElseIf e.ColumnIndex = Me.ColOnSchedule.Index Then
                    If Not String.IsNullOrEmpty(Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value) Then
                        If Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Contains("; No actual in") = True Then
                            'set actual in = sch in
                            Me.dgvItem.Item(Me.ColActualIn.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColSchIn.Index, e.RowIndex).Value
                            Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Replace("; No actual in", "")
                        End If
                        If Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Contains("No actual in") = True Then
                            'set actual in = sch in
                            Me.dgvItem.Item(Me.ColActualIn.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColSchIn.Index, e.RowIndex).Value
                            Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Replace("No actual in", "")
                        End If
                        If Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Contains("; No actual out") = True Then
                            'set actual out = sch out
                            Me.dgvItem.Item(Me.ColActualOut.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColSchOut.Index, e.RowIndex).Value
                            Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Replace("; No actual out", "")
                        End If
                        If Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Contains("No actual out") = True Then
                            'set actual out = sch out
                            Me.dgvItem.Item(Me.ColActualOut.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColSchOut.Index, e.RowIndex).Value
                            Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Replace("No actual out", "")
                        End If
                        Me.dgvItem.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.White
                        Me.dgvItem.Item(Me.ColOnSchedule.Index, e.RowIndex).Value = ""
                        Me.SetBackColorAlternating()
                        Me.Hitung(e.RowIndex, e.ColumnIndex, EnumCalculateRange.BOTH, True)
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetBackColorAlternating()
        Dim rowDGV As DataGridViewRow
        Dim ij As Integer = 0

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If IsNothing(rowDGV.Cells(Me.ColOnSchedule.Index).Value) = False Then
                    If rowDGV.Cells(Me.ColOnSchedule.Index).Value.ToString.Contains("Schedule") = True Then
                        Exit Sub
                    End If
                End If
            Next
            'kalo sampe sini, maka sudah semua
            For Each rowDGV In Me.dgvItem.Rows
                If ij = 0 Then
                    rowDGV.DefaultCellStyle.BackColor = Color.White
                    ij += 1
                Else
                    rowDGV.DefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
                    ij -= 1
                End If
            Next
            Me.ColOnSchedule.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim CboDgv As Windows.Forms.DataGridViewComboBoxCell
        Dim Status As ClsTrans_Attendance.EnumStatus
        Dim Durasi As ClsTrans_Attendance.EnumDurasi
        Dim workday As ClsTrans_Attendance.EnumWorkday
        Dim BaruBetulPertamaKali As Boolean = False

        Try
            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Buffer
                Exit Sub
            End If
            If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Buffer Then
                Exit Sub
            End If
            'workday
            CboDgv = Me.dgvItem.Item(Me.ColWorkdays.Index, e.RowIndex)
            workday = CboDgv.Items.IndexOf(Me.dgvItem.Item(Me.ColWorkdays.Index, e.RowIndex).Value)
            'fullday / half day
            CboDgv = Me.dgvItem.Item(Me.ColDuration.Index, e.RowIndex)
            Durasi = CboDgv.Items.IndexOf(Me.dgvItem.Item(Me.ColDuration.Index, e.RowIndex).Value)
            'status
            CboDgv = Me.dgvItem.Item(Me.ColAttendanceStatus.Index, e.RowIndex)
            Status = CboDgv.Items.IndexOf(Me.dgvItem.Item(Me.ColAttendanceStatus.Index, e.RowIndex).Value)
            If e.ColumnIndex = Me.ColAttendanceStatus.Index Then
                If Me.ObjAtt.CekValidStatus(workday, Durasi, Status) = True Then
                    'jika status valid
                    If Me.ObjAtt.CekHalfStatus(Status, True) = False Then
                        'set nilai actual jadi nol
                        Me.dgvItem.Item(Me.ColActualIn.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0)
                        Me.dgvItem.Item(Me.ColActualBreakIn.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0)
                        Me.dgvItem.Item(Me.ColActualBreakOut.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0)
                        Me.dgvItem.Item(Me.ColActualOut.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0)
                    Else
                        'jika nilai sebelumnya nol maka sesuaikan dengan sch, jika tidak jangan di rubah
                        If Me.dgvItem.Item(Me.ColActualIn.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0) And _
                        Me.dgvItem.Item(Me.ColActualBreakIn.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0) And _
                        Me.dgvItem.Item(Me.ColActualBreakOut.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0) And _
                        Me.dgvItem.Item(Me.ColActualOut.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0) Then
                            Me.dgvItem.Item(Me.ColActualIn.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColSchIn.Index, e.RowIndex).Value
                            Me.dgvItem.Item(Me.ColActualBreakIn.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColSchBreakIn.Index, e.RowIndex).Value
                            Me.dgvItem.Item(Me.ColActualBreakOut.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColSchBreakOut.Index, e.RowIndex).Value
                            Me.dgvItem.Item(Me.ColActualOut.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColSchOut.Index, e.RowIndex).Value
                            If Me.ObjCd.Is_Time_String(Me.dgvItem.Item(Me.ColActualIn.Index, e.RowIndex).Value.ToString) = True And Me.ObjCd.Is_Time_String(Me.dgvItem.Item(Me.ColActualOut.Index, e.RowIndex).Value.ToString) Then
                                Me.Hitung(e.RowIndex, e.ColumnIndex, EnumCalculateRange.BOTH, True)
                            End If
                        End If
                    End If
                    'Batasan Leave dan Halfday Leave
                    If Status = ClsTrans_Attendance.EnumStatus.Leave Then
                        If Me.dgvItem.Item(Me.ColPrimaryLA.Index, e.RowIndex).Value = "-1" Then
                            MsgBox("You cannot input this status, please create leave approval.", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(Me.ColAttendanceStatus.Index, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    ElseIf Status = ClsTrans_Attendance.EnumStatus.Half_day_Leave Then
                        If Me.dgvItem.Item(Me.ColPrimaryLA.Index, e.RowIndex).Value = "-1" Then
                            MsgBox("You cannot input this status, please create leave approval.", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(Me.ColAttendanceStatus.Index, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    Else
                        Me.dgvItem.Item(Me.ColPrimaryLA.Index, e.RowIndex).Value = -1
                        Me.dgvItem.Item(Me.ColPrimaryLeave.Index, e.RowIndex).Value = -1
                        Me.dgvItem.Item(Me.ColStatusRemark.Index, e.RowIndex).Value = ""
                    End If
                    '--------------------------------
                    'sesuai kan dengan is status
                    Me.KosongkanSemuaIsStatus(e)
                    If Status = ClsTrans_Attendance.EnumStatus.Present Then
                        If Durasi = ClsTrans_Attendance.EnumDurasi.Full_Day Then
                            Me.dgvItem.Item(Me.ColIsPresent.Index, e.RowIndex).Value = 1
                        ElseIf Durasi = ClsTrans_Attendance.EnumDurasi.Half_Day Then
                            Me.dgvItem.Item(Me.ColIsPresent.Index, e.RowIndex).Value = 0.5
                        End If
                    ElseIf Status = ClsTrans_Attendance.EnumStatus.Sick Then
                        Me.dgvItem.Item(Me.ColIsSick.Index, e.RowIndex).Value = 1
                        Me.dgvItem.Item(Me.ColIsOutstation.Index, e.RowIndex).Value = Boolean.FalseString
                    ElseIf Status = ClsTrans_Attendance.EnumStatus.Half_Day_Sick Then
                        Me.dgvItem.Item(Me.ColIsSick.Index, e.RowIndex).Value = 0.5
                        If Durasi = ClsTrans_Attendance.EnumDurasi.Full_Day Then
                            Me.dgvItem.Item(Me.ColIsPresent.Index, e.RowIndex).Value = 0.5
                        ElseIf Durasi = ClsTrans_Attendance.EnumDurasi.Half_Day Then
                            Me.dgvItem.Item(Me.ColIsPresent.Index, e.RowIndex).Value = 0
                        End If
                    ElseIf Status = ClsTrans_Attendance.EnumStatus.Leave Then
                        Me.dgvItem.Item(Me.ColIsLeave.Index, e.RowIndex).Value = 1
                        Me.dgvItem.Item(Me.ColIsOutstation.Index, e.RowIndex).Value = Boolean.FalseString
                    ElseIf Status = ClsTrans_Attendance.EnumStatus.Half_day_Leave Then
                        Me.dgvItem.Item(Me.ColIsLeave.Index, e.RowIndex).Value = 0.5
                        If Durasi = ClsTrans_Attendance.EnumDurasi.Full_Day Then
                            Me.dgvItem.Item(Me.ColIsPresent.Index, e.RowIndex).Value = 0.5
                        ElseIf Durasi = ClsTrans_Attendance.EnumDurasi.Half_Day Then
                            Me.dgvItem.Item(Me.ColIsPresent.Index, e.RowIndex).Value = 0
                        End If
                    ElseIf Status = ClsTrans_Attendance.EnumStatus.Absent Then
                        Me.dgvItem.Item(Me.ColIsAlfa.Index, e.RowIndex).Value = 1
                        Me.dgvItem.Item(Me.ColIsOutstation.Index, e.RowIndex).Value = Boolean.FalseString
                    ElseIf Status = ClsTrans_Attendance.EnumStatus.Half_Absent Then
                        Me.dgvItem.Item(Me.ColIsAlfa.Index, e.RowIndex).Value = 0.5
                        If Durasi = ClsTrans_Attendance.EnumDurasi.Full_Day Then
                            Me.dgvItem.Item(Me.ColIsPresent.Index, e.RowIndex).Value = 0.5
                        ElseIf Durasi = ClsTrans_Attendance.EnumDurasi.Half_Day Then
                            Me.dgvItem.Item(Me.ColIsPresent.Index, e.RowIndex).Value = 0
                        End If
                    ElseIf Status = ClsTrans_Attendance.EnumStatus.Permit Then
                        Me.dgvItem.Item(Me.ColIsPermit.Index, e.RowIndex).Value = 1
                        Me.dgvItem.Item(Me.ColIsOutstation.Index, e.RowIndex).Value = Boolean.FalseString
                    ElseIf Status = ClsTrans_Attendance.EnumStatus.Half_Day_Permit Then
                        Me.dgvItem.Item(Me.ColIsPermit.Index, e.RowIndex).Value = 0.5
                        If Durasi = ClsTrans_Attendance.EnumDurasi.Full_Day Then
                            Me.dgvItem.Item(Me.ColIsPresent.Index, e.RowIndex).Value = 0.5
                        ElseIf Durasi = ClsTrans_Attendance.EnumDurasi.Half_Day Then
                            Me.dgvItem.Item(Me.ColIsPresent.Index, e.RowIndex).Value = 0
                        End If
                    ElseIf Status = ClsTrans_Attendance.EnumStatus.Holiday Then
                        Me.dgvItem.Item(Me.ColIsHoliday.Index, e.RowIndex).Value = 1
                        Me.dgvItem.Item(Me.ColIsOutstation.Index, e.RowIndex).Value = Boolean.FalseString
                    ElseIf Status = ClsTrans_Attendance.EnumStatus.Half_Holiday Then
                        Me.dgvItem.Item(Me.ColIsHoliday.Index, e.RowIndex).Value = 0.5
                        If Durasi = ClsTrans_Attendance.EnumDurasi.Full_Day Then
                            Me.dgvItem.Item(Me.ColIsPresent.Index, e.RowIndex).Value = 0.5
                        ElseIf Durasi = ClsTrans_Attendance.EnumDurasi.Half_Day Then
                            Me.dgvItem.Item(Me.ColIsPresent.Index, e.RowIndex).Value = 0
                        End If
                    End If
                    CboDgv = Me.dgvItem.Item(Me.ColOT.Index, e.RowIndex)
                    If Status <> ClsTrans_Attendance.EnumStatus.Present Then
                        Me.dgvItem.Item(Me.ColOT.Index, e.RowIndex).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumTipeOt.No_Overtime)
                        Me.dgvItem.Item(Me.ColPermitOut.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0)
                        Me.dgvItem.Item(Me.ColPermitIn.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0)
                        Me.dgvItem.Item(Me.ColPermitDuration.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0)
                        Me.dgvItem.Item(Me.ColPermit.Index, e.RowIndex).Value = False
                    Else
                        'Request dari ci Fitri, kalo holday present, maka Overtime nya diset dengan default holiday overtime
                        If Me.dgvItem.Item(Me.ColWorkdays.Index, e.RowIndex).Value = "Holiday" Then
                            Me.dgvItem.Item(Me.ColOT.Index, e.RowIndex).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumTipeOt.Holiday_Overtime)
                        Else
                            Me.dgvItem.Item(Me.ColOT.Index, e.RowIndex).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumTipeOt.Weekday_Overtime)
                        End If
                        '----------------------------------------------------------------------------------------------------
                    End If
                Else
                    'jika status tidak valid balikin ke buffer
                    CboDgv = Me.dgvItem.Item(Me.ColAttendanceStatus.Index, e.RowIndex)
                    Me.dgvItem.Item(Me.ColAttendanceStatus.Index, e.RowIndex).Value = CboDgv.Items.Item(Me.BufferStatus)
                End If
                CboDgv = Me.dgvItem.Item(Me.ColAttendanceStatus.Index, e.RowIndex)
                If Me.dgvItem.Item(Me.ColAttendanceStatus.Index, e.RowIndex).Value <> CboDgv.Items.Item(Me.BufferStatus) Then
                    If Me.ObjCd.Is_Time_String(Me.dgvItem.Item(Me.ColActualIn.Index, e.RowIndex).Value.ToString) = True And Me.ObjCd.Is_Time_String(Me.dgvItem.Item(Me.ColActualOut.Index, e.RowIndex).Value.ToString) Then
                        Me.Hitung(e.RowIndex, e.ColumnIndex, EnumCalculateRange.BOTH, True)
                    End If

                    'Me.KosongkanSemuaIsStatus(e)
                End If
            End If

            'cek format
            If e.ColumnIndex = Me.ColActualIn.Index Or e.ColumnIndex = Me.ColActualBreakOut.Index Or _
            e.ColumnIndex = Me.ColActualBreakIn.Index Or e.ColumnIndex = Me.ColActualOut.Index Or _
            e.ColumnIndex = Me.ColEarlyInDuration.Index Or e.ColumnIndex = Me.ColLateDuration.Index Or _
            e.ColumnIndex = Me.ColEarlyBreakOutDuration.Index Or e.ColumnIndex = Me.ColLateBreakInDuration.Index Or _
            e.ColumnIndex = Me.ColEarlyOutDuration.Index Or e.ColumnIndex = Me.ColOTDuration.Index Or _
            e.ColumnIndex = Me.ColEarlyOTDuration.Index Or e.ColumnIndex = Me.ColPermitOut.Index Or e.ColumnIndex = Me.ColPermitIn.Index Then
                If String.IsNullOrEmpty(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.ObjCd.Is_Time_String(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        If Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > 1439 Then
                            MsgBox("Schedule in not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If

                        If e.ColumnIndex = Me.ColActualIn.Index Or e.ColumnIndex = Me.ColActualBreakOut.Index Or _
                        e.ColumnIndex = Me.ColActualBreakIn.Index Or e.ColumnIndex = Me.ColActualOut.Index Or _
                        e.ColumnIndex = Me.ColPermitOut.Index Or e.ColumnIndex = Me.ColPermitIn.Index Then

                            If Me.IsImportFromMachine = True Then
                                If String.IsNullOrEmpty(Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value) = False Then
                                    If e.ColumnIndex = Me.ColActualIn.Index Then
                                        If Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Contains("No actual in; ") Then
                                            Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Replace("No actual in; ", "")
                                            BaruBetulPertamaKali = True
                                        ElseIf Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Contains("No actual in") Then
                                            Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Replace("No actual in", "")
                                            BaruBetulPertamaKali = True
                                        End If
                                    End If
                                    If e.ColumnIndex = Me.ColActualOut.Index Then
                                        If Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Contains("; No actual out") Then
                                            Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Replace("; No actual out", "")
                                            BaruBetulPertamaKali = True
                                        ElseIf Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Contains("No actual out") Then
                                            Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Replace("No actual out", "")
                                            BaruBetulPertamaKali = True
                                        End If
                                    End If
                                    If Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Contains("; No actual in") = False And _
                                        Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Contains("No actual in") = False And _
                                        Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Contains("; No actual out") = False And _
                                        Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Contains("No actual out") = False Then
                                        Me.dgvItem.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.White
                                        Me.dgvItem.Item(Me.ColOnSchedule.Index, e.RowIndex).Value = ""
                                        Me.SetBackColorAlternating()
                                    End If
                                End If
                            End If

                            '------------------------------------------------------------------------
                            Status = CboDgv.Items.IndexOf(Me.dgvItem.Item(Me.ColAttendanceStatus.Index, e.RowIndex).Value)

                            If Me.ObjCd.Is_Time_String(Me.dgvItem.Item(Me.ColActualIn.Index, e.RowIndex).Value.ToString) = True And Me.ObjCd.Is_Time_String(Me.dgvItem.Item(Me.ColActualOut.Index, e.RowIndex).Value.ToString) Then
                                'If Status = ClsTrans_Attendance.EnumStatus.Half_Absent OrElse _
                                'Status = ClsTrans_Attendance.EnumStatus.Half_day_Leave OrElse _
                                'Status = ClsTrans_Attendance.EnumStatus.Half_Day_Permit OrElse _
                                'Status = ClsTrans_Attendance.EnumStatus.Half_Day_Sick OrElse _
                                'Status = ClsTrans_Attendance.EnumStatus.Half_Holiday Then

                                '    If Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColActualIn.Index, e.RowIndex).Value.ToString) > 0 Then
                                '        If Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColActualOut.Index, e.RowIndex).Value.ToString) > 0 Then
                                '            Me.Hitung(e.RowIndex, EnumCalculateRange.BOTH)
                                '        Else
                                '            Me.Hitung(e.RowIndex, EnumCalculateRange.IN_RANGE)
                                '        End If
                                '    End If

                                '    If Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColActualOut.Index, e.RowIndex).Value.ToString) > 0 Then
                                '        If Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColActualIn.Index, e.RowIndex).Value.ToString) > 0 Then
                                '            Me.Hitung(e.RowIndex, EnumCalculateRange.BOTH)
                                '        Else
                                '            Me.Hitung(e.RowIndex, EnumCalculateRange.OUT_RANGE)
                                '        End If
                                '    End If
                                'Else
                                '    If e.ColumnIndex = Me.ColActualIn.Index Then
                                '        Me.Hitung(e.RowIndex, EnumCalculateRange.IN_RANGE)
                                '    ElseIf e.ColumnIndex = Me.ColActualOut.Index Then
                                '        Me.Hitung(e.RowIndex, EnumCalculateRange.OUT_RANGE)
                                '    Else
                                '        Me.Hitung(e.RowIndex, EnumCalculateRange.BOTH)
                                '    End If
                                'End If

                                '********************************************************************************************************************
                                'KETENTUAN BARU DARI PAK IWAN BERDASARKAN ERROR LIST CI FITRI NO 49 - 26 JUNI 2016 - PERBAIKAN : 28 JULI 2016
                                'JIKA HALFDAY (LEAVE, HOLIDAY ETC) DILAKUKAN PENGECEKAN, JAM MASUK LEBIH BERAT SEBELUM ISTIRHAT ATAU SESUDAH
                                'JIKA MASUK SEBELUM ISTIRAHAT(PAGI) MAKA TIDAK BOLEH HITUNG EARLY OUT (EARLY OUT = 00.00) TAPI LATE DIHITUNG
                                'JIKA MASUK SESUDAH ISTIRAHAT(SIANG) MAKA TIDAK BOLEH HITUNG LATE (LATE = 00.00) TAPI EARLY OUT DIHITUNG
                                '********************************************************************************************************************
                                If Status = ClsTrans_Attendance.EnumStatus.Half_Absent OrElse _
                                      Status = ClsTrans_Attendance.EnumStatus.Half_day_Leave OrElse _
                                      Status = ClsTrans_Attendance.EnumStatus.Half_Day_Permit OrElse _
                                      Status = ClsTrans_Attendance.EnumStatus.Half_Day_Sick OrElse _
                                      Status = ClsTrans_Attendance.EnumStatus.Half_Holiday Then
                                    Dim ActIn As Decimal = 0
                                    If Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColActualIn.Index, e.RowIndex).Value.ToString) > 0 Then
                                        ActIn = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColActualIn.Index, e.RowIndex).Value.ToString)
                                    End If
                                    Dim ActOut As Decimal = 0
                                    If Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColActualOut.Index, e.RowIndex).Value.ToString) > 0 Then
                                        ActOut = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColActualOut.Index, e.RowIndex).Value.ToString)
                                    End If
                                    Dim BOut As Decimal = 0
                                    If Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColActualBreakOut.Index, e.RowIndex).Value.ToString) > 0 Then
                                        BOut = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColActualBreakOut.Index, e.RowIndex).Value.ToString)
                                    End If

                                    Dim BIN As Decimal = 0
                                    If Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColActualBreakIn.Index, e.RowIndex).Value.ToString) > 0 Then
                                        BIN = Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColActualBreakIn.Index, e.RowIndex).Value.ToString)
                                    End If

                                    Dim selisihpagi As Decimal = 0
                                    Dim selisihSiang As Decimal = 0

                                    selisihpagi = BOut - ActIn
                                    selisihSiang = ActOut - BIN

                                    If selisihpagi >= selisihSiang AndAlso selisihpagi > 0 Then
                                        Me.Hitung(e.RowIndex, e.ColumnIndex, EnumCalculateRange.IN_RANGE, True)
                                        'Me.dgvItem.Item(Me.ColEarlyOutDuration.Index, e.RowIndex).Value = "00.00"
                                    Else
                                        Me.Hitung(e.RowIndex, e.ColumnIndex, EnumCalculateRange.OUT_RANGE, True)
                                        'Me.dgvItem.Item(Me.ColLateDuration.Index, e.RowIndex).Value = "00.00"
                                    End If
                                    '********************************************************************************************************************
                                    '********************************************************************************************************************
                                Else
                                    CboDgv = Me.dgvItem.Item(Me.ColWorkdays.Index, e.RowIndex)
                                    If e.ColumnIndex = Me.ColActualIn.Index Then
                                        If CboDgv.Items.IndexOf(Me.dgvItem.Item(Me.ColWorkdays.Index, e.RowIndex).Value) = ClsTrans_Attendance.EnumWorkday.Holiday Then
                                            Me.Hitung(e.RowIndex, e.ColumnIndex, EnumCalculateRange.BOTH, True)
                                        Else
                                            If BaruBetulPertamaKali = True Then
                                                Me.Hitung(e.RowIndex, e.ColumnIndex, EnumCalculateRange.BOTH, True)
                                            Else
                                                Me.Hitung(e.RowIndex, e.ColumnIndex, EnumCalculateRange.IN_RANGE, True)
                                            End If
                                        End If
                                    ElseIf e.ColumnIndex = Me.ColActualOut.Index Then
                                        If CboDgv.Items.IndexOf(Me.dgvItem.Item(Me.ColWorkdays.Index, e.RowIndex).Value) = ClsTrans_Attendance.EnumWorkday.Holiday Then
                                            Me.Hitung(e.RowIndex, e.ColumnIndex, EnumCalculateRange.BOTH, True)
                                        Else
                                            If BaruBetulPertamaKali = True Then
                                                Me.Hitung(e.RowIndex, e.ColumnIndex, EnumCalculateRange.BOTH, True)
                                            Else
                                                Me.Hitung(e.RowIndex, e.ColumnIndex, EnumCalculateRange.OUT_RANGE, True)
                                            End If
                                        End If
                                    Else
                                        Me.Hitung(e.RowIndex, e.ColumnIndex, EnumCalculateRange.BOTH, True)
                                    End If
                                End If


                            End If

                            'If Me.ObjCd.Is_Time_String(Me.dgvItem.Item(Me.ColActualIn.Index, e.RowIndex).Value.ToString) = True And Me.ObjCd.Is_Time_String(Me.dgvItem.Item(Me.ColActualOut.Index, e.RowIndex).Value.ToString) Then
                            '    If e.ColumnIndex = Me.ColActualIn.Index Then
                            '        Me.Hitung(e.RowIndex, EnumCalculateRange.IN_RANGE)
                            '    ElseIf e.ColumnIndex = Me.ColActualOut.Index Then
                            '        Me.Hitung(e.RowIndex, EnumCalculateRange.OUT_RANGE)
                            '    Else
                            '        Me.Hitung(e.RowIndex, EnumCalculateRange.BOTH)
                            '    End If

                            'End If

                            'Me.KosongkanSemuaIsStatus(e)
                        Else
                            Dim max As String = ""
                            If e.ColumnIndex = Me.ColEarlyInDuration.Index Then
                                max = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
                                                             Core_Module.ClsTransaksiSetting.EnumNamaParameter.MaxValue_Early_In)
                            ElseIf e.ColumnIndex = Me.ColLateDuration.Index Then
                                max = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
                                                             Core_Module.ClsTransaksiSetting.EnumNamaParameter.MaxValue_Late)
                                'is late
                                If Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > 0 Then
                                    Me.dgvItem.Item(Me.ColIsLate.Index, e.RowIndex).Value = 1
                                Else
                                    Me.dgvItem.Item(Me.ColIsLate.Index, e.RowIndex).Value = 0
                                End If
                            ElseIf e.ColumnIndex = Me.ColEarlyBreakOutDuration.Index Then
                                max = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
                                                             Core_Module.ClsTransaksiSetting.EnumNamaParameter.MaxValue_Early_Break_Out)
                            ElseIf e.ColumnIndex = Me.ColLateBreakInDuration.Index Then
                                max = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
                                                             Core_Module.ClsTransaksiSetting.EnumNamaParameter.MaxValue_Late_Break_In)
                            ElseIf e.ColumnIndex = Me.ColEarlyOutDuration.Index Then
                                max = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
                                                             Core_Module.ClsTransaksiSetting.EnumNamaParameter.MaxValue_Early_Out)
                                'is early out
                                If Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > 0 Then
                                    Me.dgvItem.Item(Me.ColIsEarlyOut.Index, e.RowIndex).Value = 1
                                Else
                                    Me.dgvItem.Item(Me.ColIsEarlyOut.Index, e.RowIndex).Value = 0
                                End If
                            ElseIf e.ColumnIndex = Me.ColOTDuration.Index Then
                                max = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
                                                             Core_Module.ClsTransaksiSetting.EnumNamaParameter.MaxValue_OT_Daily)
                                'is ot
                                If Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > 0 Then
                                    Me.dgvItem.Item(Me.ColIsOvertime.Index, e.RowIndex).Value = 1
                                Else
                                    Me.dgvItem.Item(Me.ColIsOvertime.Index, e.RowIndex).Value = 0
                                End If
                            ElseIf e.ColumnIndex = Me.ColEarlyOTDuration.Index Then
                                max = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
                                                             Core_Module.ClsTransaksiSetting.EnumNamaParameter.MaxValue_Early_OT_Daily)
                                'is early ot
                                If Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > 0 Then
                                    Me.dgvItem.Item(Me.ColIsEarlyOvertime.Index, e.RowIndex).Value = 1
                                Else
                                    Me.dgvItem.Item(Me.ColIsEarlyOvertime.Index, e.RowIndex).Value = 0
                                End If
                            End If
                            'kalo max = null ato 0, tidak ada batasan
                            If String.IsNullOrEmpty(max) = False Then
                                If max <> "" Then
                                    If max > 0 Then
                                        If Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > max Then
                                            MsgBox("Value cannot more than " + Me.ObjCd.Convert_ToString(max), MsgBoxStyle.Information, "Information")
                                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        MsgBox("Please input hh.mm format", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                Else
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "00.00"
                End If
            End If
            'chk permit
            'If e.ColumnIndex = Me.ColPermit.Index Then
            '    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = True Then
            '        Me.dgvItem.Item(Me.ColPermitOut.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColSchIn.Index, e.RowIndex).Value))
            '        Me.dgvItem.Item(Me.ColPermitIn.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColSchOut.Index, e.RowIndex).Value))
            '        Me.dgvItem.Item(Me.ColPermitDuration.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColSchOut.Index, e.RowIndex).Value) - Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColSchIn.Index, e.RowIndex).Value))
            '        '-------------------------------------------------------------------------------------------------------
            '    Else
            '        Me.dgvItem.Item(Me.ColPermitOut.Index, e.RowIndex).Value = "00.00"
            '        Me.dgvItem.Item(Me.ColPermitIn.Index, e.RowIndex).Value = "00.00"
            '        Me.dgvItem.Item(Me.ColPermitDuration.Index, e.RowIndex).Value = "00.00"
            '    End If
            '    Me.Hitung(e.RowIndex, e.ColumnIndex, EnumCalculateRange.BOTH, True)
            '    Me.SetKolomPermit()
            'End If
            If e.ColumnIndex = Me.ColPermit.Index Then
                If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = True Then
                    Me.dgvItem.Item(Me.ColPermitOut.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColSchIn.Index, e.RowIndex).Value))
                    Me.dgvItem.Item(Me.ColPermitIn.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColSchOut.Index, e.RowIndex).Value))
                    Me.HitungPermit(e.RowIndex, e.ColumnIndex)
                    'Me.dgvItem.Item(Me.ColPermitDuration.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColSchOut.Index, e.RowIndex).Value) - Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColSchIn.Index, e.RowIndex).Value))
                    '-------------------------------------------------------------------------------------------------------
                Else
                    Me.dgvItem.Item(Me.ColPermitOut.Index, e.RowIndex).Value = "00.00"
                    Me.dgvItem.Item(Me.ColPermitIn.Index, e.RowIndex).Value = "00.00"
                    Me.dgvItem.Item(Me.ColPermitDuration.Index, e.RowIndex).Value = "00.00"
                End If
                Me.Hitung(e.RowIndex, e.ColumnIndex, EnumCalculateRange.BOTH, True)
                Me.SetKolomPermit()
            End If
            If e.ColumnIndex = Me.ColPermitOut.Index Then
                If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> Me.Buffer Then
                    If Me.BatasanPermitOutHours(Me.dgvItem.Item(Me.ColSchIn.Index, e.RowIndex).Value, _
                    Me.dgvItem.Item(Me.ColSchOut.Index, e.RowIndex).Value, _
                    Me.dgvItem.Item(Me.ColPermitOut.Index, e.RowIndex).Value) = False Then
                        MsgBox("You can't set permit duration out of schedule duration.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    If Me.BatasanPermitOutPermitIn(Me.dgvItem.Item(Me.ColSchIn.Index, e.RowIndex).Value, _
                    Me.dgvItem.Item(Me.ColSchOut.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPermitOut.Index, e.RowIndex).Value, _
                    Me.dgvItem.Item(Me.ColPermitIn.Index, e.RowIndex).Value, True) = False Then
                        MsgBox("you can't input permit out bigger than permit in.", MsgBoxStyle.Information, "Infomation")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If
            End If
            If e.ColumnIndex = Me.ColPermitIn.Index Then
                If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> Me.Buffer Then
                    If Me.BatasanPermitInHours(Me.dgvItem.Item(Me.ColSchIn.Index, e.RowIndex).Value, _
                    Me.dgvItem.Item(Me.ColSchOut.Index, e.RowIndex).Value, _
                    Me.dgvItem.Item(Me.ColPermitIn.Index, e.RowIndex).Value) = False Then
                        MsgBox("You can't set permit duration out of schedule duration.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    If Me.BatasanPermitOutPermitIn(Me.dgvItem.Item(Me.ColSchIn.Index, e.RowIndex).Value, _
                    Me.dgvItem.Item(Me.ColSchOut.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPermitOut.Index, e.RowIndex).Value, _
                    Me.dgvItem.Item(Me.ColPermitIn.Index, e.RowIndex).Value, False) = False Then
                        MsgBox("you can't input permit in less than permit out.", MsgBoxStyle.Information, "Infomation")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If
            End If
            If e.ColumnIndex = Me.ColOT.Index Then
                CboDgv = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex)
                If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumTipeOt.No_Overtime) Then
                    Me.dgvItem.Item(Me.ColIsOvertime.Index, e.RowIndex).Value = 0
                    Me.dgvItem.Item(Me.ColIsEarlyOvertime.Index, e.RowIndex).Value = 0
                    If Me.ObjCd.Is_Time_String(Me.dgvItem.Item(Me.ColActualIn.Index, e.RowIndex).Value.ToString) = True And Me.ObjCd.Is_Time_String(Me.dgvItem.Item(Me.ColActualOut.Index, e.RowIndex).Value.ToString) Then
                        Me.Hitung(e.RowIndex, e.ColumnIndex, EnumCalculateRange.BOTH, True)
                    End If
                Else
                    If Me.ObjCd.Is_Time_String(Me.dgvItem.Item(Me.ColActualIn.Index, e.RowIndex).Value.ToString) = True And Me.ObjCd.Is_Time_String(Me.dgvItem.Item(Me.ColActualOut.Index, e.RowIndex).Value.ToString) Then
                        Me.Hitung(e.RowIndex, e.ColumnIndex, EnumCalculateRange.BOTH, True)
                    End If
                End If
            End If
            'Me.Hitung(e)

            If e.ColumnIndex = Me.ColIsCalculateEarlyOT.Index Or e.ColumnIndex = Me.ColIsCalculateLate.Index Then
                If Me.ObjCd.Is_Time_String(Me.dgvItem.Item(Me.ColActualIn.Index, e.RowIndex).Value.ToString) = True And Me.ObjCd.Is_Time_String(Me.dgvItem.Item(Me.ColActualOut.Index, e.RowIndex).Value.ToString) Then
                    Me.Hitung(e.RowIndex, e.ColumnIndex, EnumCalculateRange.IN_RANGE, True)
                End If
            End If
            If e.ColumnIndex = Me.ColIsCalculateEarlyOut.Index Or e.ColumnIndex = Me.ColIsCalculateOT.Index Then
                If Me.ObjCd.Is_Time_String(Me.dgvItem.Item(Me.ColActualIn.Index, e.RowIndex).Value.ToString) = True And Me.ObjCd.Is_Time_String(Me.dgvItem.Item(Me.ColActualOut.Index, e.RowIndex).Value.ToString) Then
                    Me.Hitung(e.RowIndex, e.ColumnIndex, EnumCalculateRange.OUT_RANGE, True)
                End If
            End If

            'aul -penambahan outstation-
            If e.ColumnIndex = Me.ColIsOutstation.Index Then
                If Not String.IsNullOrEmpty(Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value) Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.TrueString Then
                        Dim IfOutstationIn, IfOutstationOut As Core_Module.ClsMasterWorkShift.EnumDefaultIfOutstation
                        Dim DefaultCustomIn, DefaultCustomOut As String
                        Dim PrimaryWorkshift_ As Double = -1

                        PrimaryWorkshift_ = CDbl(Me.dgvItem.Item(Me.ColPrimaryWorkShift.Index, e.RowIndex).Value.ToString)
                        IfOutstationIn = Me.ObjWS.GetDefaultIfOutstationInOut(PrimaryWorkshift_, Core_Module.ClsMasterWorkShift.EnumScheduleType.SCHEDULE_IN)
                        IfOutstationOut = Me.ObjWS.GetDefaultIfOutstationInOut(PrimaryWorkshift_, Core_Module.ClsMasterWorkShift.EnumScheduleType.SCHEDULE_OUT)
                        DefaultCustomIn = Me.ObjWS.GetDefaultCustom(PrimaryWorkshift_, Core_Module.ClsMasterWorkShift.EnumScheduleType.SCHEDULE_IN)
                        DefaultCustomOut = Me.ObjWS.GetDefaultCustom(PrimaryWorkshift_, Core_Module.ClsMasterWorkShift.EnumScheduleType.SCHEDULE_OUT)

                        'berkaitan actual in
                        If Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Contains("; No actual in") = True Then
                            'set actual in = sch in
                            If IfOutstationIn = Core_Module.ClsMasterWorkShift.EnumDefaultIfOutstation.SCHEDULE Then
                                Me.dgvItem.Item(Me.ColActualIn.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColSchIn.Index, e.RowIndex).Value
                                Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Replace("; No actual in", "")
                            ElseIf IfOutstationIn = Core_Module.ClsMasterWorkShift.EnumDefaultIfOutstation.CUSTOM Then
                                Me.dgvItem.Item(Me.ColActualIn.Index, e.RowIndex).Value = DefaultCustomIn
                                If Me.ObjCd.Is_Time_String(DefaultCustomIn) = True Then
                                    If Me.ObjCd.Convert_ToMinutes(DefaultCustomIn) < Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColActualIn.Index, e.RowIndex).Value) Then
                                        Exit Sub
                                    End If
                                Else
                                    Exit Sub
                                End If
                            End If
                        End If
                        If Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Contains("No actual in") = True Then
                            'set actual in = sch in
                            If IfOutstationIn = Core_Module.ClsMasterWorkShift.EnumDefaultIfOutstation.SCHEDULE Then
                                Me.dgvItem.Item(Me.ColActualIn.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColSchIn.Index, e.RowIndex).Value
                                Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Replace("No actual in", "")
                            ElseIf IfOutstationIn = Core_Module.ClsMasterWorkShift.EnumDefaultIfOutstation.CUSTOM Then
                                Me.dgvItem.Item(Me.ColActualIn.Index, e.RowIndex).Value = DefaultCustomIn
                                If Me.ObjCd.Is_Time_String(DefaultCustomIn) = True Then
                                    If Me.ObjCd.Convert_ToMinutes(DefaultCustomIn) < Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColActualIn.Index, e.RowIndex).Value) Then
                                        Exit Sub
                                    End If
                                Else
                                    Exit Sub
                                End If
                            End If
                        End If

                        'berkaitan actual out
                        If Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Contains("; No actual out") = True Then
                            'set actual out = sch out
                            If IfOutstationOut = Core_Module.ClsMasterWorkShift.EnumDefaultIfOutstation.SCHEDULE Then
                                Me.dgvItem.Item(Me.ColActualOut.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColSchOut.Index, e.RowIndex).Value
                                Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Replace("; No actual out", "")
                            ElseIf IfOutstationIn = Core_Module.ClsMasterWorkShift.EnumDefaultIfOutstation.CUSTOM Then
                                Me.dgvItem.Item(Me.ColActualOut.Index, e.RowIndex).Value = DefaultCustomOut
                                If Me.ObjCd.Is_Time_String(DefaultCustomOut) = True Then
                                    If Me.ObjCd.Convert_ToMinutes(DefaultCustomOut) < Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColActualOut.Index, e.RowIndex).Value) Then
                                        Exit Sub
                                    End If
                                Else
                                    Exit Sub
                                End If
                            End If
                        End If
                        If Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Contains("No actual out") = True Then
                            'set actual out = sch out
                            If IfOutstationOut = Core_Module.ClsMasterWorkShift.EnumDefaultIfOutstation.SCHEDULE Then
                                Me.dgvItem.Item(Me.ColActualOut.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColSchOut.Index, e.RowIndex).Value
                                Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value.ToString.Replace("No actual out", "")
                            ElseIf IfOutstationIn = Core_Module.ClsMasterWorkShift.EnumDefaultIfOutstation.CUSTOM Then
                                Me.dgvItem.Item(Me.ColActualOut.Index, e.RowIndex).Value = DefaultCustomOut
                                If Me.ObjCd.Is_Time_String(DefaultCustomOut) = True Then
                                    If Me.ObjCd.Convert_ToMinutes(DefaultCustomOut) < Me.ObjCd.Convert_ToMinutes(Me.dgvItem.Item(Me.ColActualOut.Index, e.RowIndex).Value) Then
                                        Exit Sub
                                    End If
                                Else
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If

                Me.dgvItem.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.White
                Me.dgvItem.Item(Me.ColOnSchedule.Index, e.RowIndex).Value = ""
                Me.SetBackColorAlternating()
                Me.Hitung(e.RowIndex, e.ColumnIndex, EnumCalculateRange.BOTH, True)
            End If

            If e.ColumnIndex = Me.ColPermit.Index Or e.ColumnIndex = Me.ColPermitOut.Index Or e.ColumnIndex = Me.ColPermitIn.Index Then
                Me.HitungPermit(e.RowIndex, e.ColumnIndex)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function BatasanPermitOutPermitIn(ByVal SchIn As String, ByVal SchOut As String, _
     ByVal PermitOut As String, ByVal PermitIn As String, Optional ByVal ApakahCekPermitOut As Boolean = True) As Boolean
        Dim MinPermitOut As Double = Me.ObjCd.Convert_ToMinutes(PermitOut)
        Dim MinPermitIn As Double = Me.ObjCd.Convert_ToMinutes(PermitIn)
        Dim MinSchIn As Double = Me.ObjCd.Convert_ToMinutes(SchIn)
        Dim MinSchOut As Double = Me.ObjCd.Convert_ToMinutes(SchOut)
        Try
            If MinSchIn > MinSchOut Then
                If MinPermitOut > MinPermitIn Then
                    If MinPermitOut < MinSchOut And MinPermitIn > MinSchIn Then
                        Return False
                    End If
                End If

                If MinPermitIn > MinSchOut And MinPermitOut < MinSchOut Then
                    Return False
                End If
            Else
                If ApakahCekPermitOut = True Then
                    If MinPermitOut > MinPermitIn Then
                        Return False
                    End If
                Else
                    If MinPermitIn < MinPermitOut Then
                        Return False
                    End If
                End If
            End If

            If MinPermitOut = MinPermitIn Then
                If MinPermitOut <> 0 And MinPermitIn <> 0 Then
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function BatasanPermitOutHours(ByVal SchIn As String, ByVal SchOut As String, ByVal PermitOut As String) As Boolean
        Dim MinSchIn As Double = Me.ObjCd.Convert_ToMinutes(SchIn)
        Dim MinSchOut As Double = Me.ObjCd.Convert_ToMinutes(SchOut)
        Dim MinPermitOut As Double = Me.ObjCd.Convert_ToMinutes(PermitOut)
        'Night
        Dim NHariPertama As Double
        Try
            If MinSchIn > MinSchOut Then
                'Night
                If MinPermitOut > MinSchOut And MinPermitOut < MinSchIn Then
                    Return False
                End If

                'NHariPertama = 1439 - MinSchIn
                'If MinPermitOut < NHariPertama Then
                '    Return False
                'Else
                '    '00.00 - 08.00
                '    If MinPermitOut > MinSchOut Then
                '        Return False
                '    End If
                'End If
            Else
                'Day
                If MinPermitOut < MinSchIn Then
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function BatasanPermitInHours(ByVal SchIn As String, ByVal SchOut As String, ByVal PermitIn As String) As Boolean
        Dim MinSchIn As Double = Me.ObjCd.Convert_ToMinutes(SchIn)
        Dim MinSchOut As Double = Me.ObjCd.Convert_ToMinutes(SchOut)
        Dim MinPermitIn As Double = Me.ObjCd.Convert_ToMinutes(PermitIn)
        Try
            'Day/Night
            If MinSchIn > MinSchOut Then
                If MinPermitIn > MinSchOut And MinPermitIn < MinSchIn Then
                    Return False
                End If
            Else
                If MinPermitIn > MinSchOut Then
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Dim Duration As Double = 0
        Try
            'hitungan
            'early In
            Me.dgvItem.Item(Me.ColEarlyInDuration.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0)
            'late
            Me.dgvItem.Item(Me.ColLateDuration.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0)
            'early out
            Me.dgvItem.Item(Me.ColEarlyOutDuration.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0)
            'ot duration
            Me.dgvItem.Item(Me.ColOTDuration.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0)
            'early ot
            Me.dgvItem.Item(Me.ColEarlyOTDuration.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0)
            'early break out
            Me.dgvItem.Item(Me.ColEarlyBreakOutDuration.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0)
            'late break in
            Me.dgvItem.Item(Me.ColLateBreakInDuration.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0)
            'primary leave
            Me.dgvItem.Item(Me.ColPrimaryLeave.Index, e.RowIndex).Value = -1
            'primary LA
            Me.dgvItem.Item(Me.ColPrimaryLA.Index, e.RowIndex).Value = -1
            'permit out
            Me.dgvItem.Item(Me.ColPermitOut.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0)
            'permit in
            Me.dgvItem.Item(Me.ColPermitIn.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0)
            'permit durasi
            Me.dgvItem.Item(Me.ColPermitDuration.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0)
            'primary overtime order
            Me.dgvItem.Item(Me.ColPrimaryOvertimeOrder.Index, e.RowIndex).Value = -1
            'ot duration SPL
            Me.dgvItem.Item(Me.ColOTOrderDuration.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0)
            'early ot SPL
            Me.dgvItem.Item(Me.ColEarlyOTOrderDuration.Index, e.RowIndex).Value = Me.ObjCd.Convert_ToString(0)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.Baru = True
            Me.ObjAtt.KosongkanObject()
            Me.KosongKanField()
            Me.AktifkanDetail()
            Me.TxtDocNumber.Text = Me.ObjAtt.GetNoBaru(Me.DtpDate.Value)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.DtpDate.Focus()
            Me.ColShowLog.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateIsOvertime()
        Dim rowDGV As DataGridViewRow
        Dim CboDgv As Windows.Forms.DataGridViewComboBoxCell
        Dim Status As ClsTrans_Attendance.EnumStatus
        Dim Durasi As ClsTrans_Attendance.EnumDurasi
        Try
            For Each rowDGV In Me.dgvItem.Rows
                CboDgv = Me.dgvItem.Item(Me.ColAttendanceStatus.Index, rowDGV.Index)
                Status = CboDgv.Items.IndexOf(Me.dgvItem.Item(Me.ColAttendanceStatus.Index, rowDGV.Index).Value)
                If Status = ClsTrans_Attendance.EnumStatus.Present Then
                    If Durasi = ClsTrans_Attendance.EnumDurasi.Full_Day Then
                        Me.dgvItem.Item(Me.ColIsPresent.Index, rowDGV.Index).Value = 1
                    ElseIf Durasi = ClsTrans_Attendance.EnumDurasi.Half_Day Then
                        Me.dgvItem.Item(Me.ColIsPresent.Index, rowDGV.Index).Value = 0.5
                    End If
                End If
                If Me.ObjCd.Convert_ToMinutes(rowDGV.Cells(Me.ColOTDuration.Index).Value) > 0 Or _
                Me.ObjCd.Convert_ToMinutes(rowDGV.Cells(Me.ColEarlyOTDuration.Index).Value) > 0 Then
                    rowDGV.Cells(Me.ColIsOvertime.Index).Value = 1
                Else
                    rowDGV.Cells(Me.ColIsOvertime.Index).Value = 0
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            'cek data kosong
            Me.UpdateIsOvertime()
            If Me.ApakahDataBenar() = True Then
                hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                Me.SalinKeObject()
                If Baru = True Then
                    'konfirmasi disimpan atau tidak
                    If hasil = MsgBoxResult.Yes Then
                        'konfirmasi disimpan atau tidak
                        Windows.Forms.Cursor.Current = Windows.Forms.Cursors.WaitCursor
                        If Me.ObjAtt.SimpanBaru(Me.UserCode, Me.dgvItem, Me, Me.TblOT_Order) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ObjAtt.SimpanPerubahan(Me.UserCode, Me.dgvItem, Me, Me.TblOT_Order) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                        Else
                            Exit Sub
                        End If
                    End If
                End If
            End If
            If Me.DariQuick = True Then
                If hasil = MsgBoxResult.Yes Then
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtDocNumber.Text = "" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtDocNumber.Text <> "" Then
                If Me.ObjAtt.ApakahBolehEdit(Me.TxtPrimary.Text, Me.UserCode) = True Then
                    Me.Cursor = Cursors.WaitCursor
                    Me.Baru = False
                    Me.ObjAtt.CariDetail(Me.TxtPrimary.Text, Me, Me.TblOT_Order, Me.JoinWithAttPrev)
                    Me.SalinKeField()
                    Me.AktifkanDetail()
                    Me.TxtEmployeeCode.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtDocNumber.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtDocNumber.Text <> "" Then
                If Me.ObjAtt.ApakahBolehHapus(Me.TxtPrimary.Text, Me.UserCode) = True Then
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        'Me.ObjAtt.IsiTblSpl(False, Me.TxtPrimaryEmployee.Text, Me.DtpStartAttPrev, Me.dtpStart.Value.Date, Me.DtpEnd.Value.Date, Me.TblOT_Order)
                        If Me.ObjAtt.Hapus(Me.TxtPrimary.Text, Me.UserCode, Me, Me.TblOT_Order, False) = True Then
                            Me.KosongKanField()
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

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtDocNumber.Text <> "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Baru = True Then
                        Me.NonAktifkanDetail()
                        Me.KosongKanField()
                    Else
                        Me.NonAktifkanDetail()
                        Me.ObjAtt.CariDetail(Me.TxtPrimary.Text, Me, Me.TblOT_Order, Me.JoinWithAttPrev)
                        Me.SalinKeField()
                    End If
                Else
                    If Me.DariQuick = True Then
                        Me.TxtDocNumber.Enabled = False
                    End If
                    Exit Sub
                End If
            Else
                Me.NonAktifkanDetail()
                Me.KosongKanField()
            End If
            Me.btnNew.Focus()

            If Me.DariQuick = True Then
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmTrans_Attendance_Cari
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.primary = -1 Then
                Me.Cursor = Cursors.WaitCursor
                If Me.ObjAtt.CariDetail(hj.primary, Me, Me.TblOT_Order, Me.JoinWithAttPrev) = True Then
                    Me.SalinKeField()
                    Me.NonAktifkanDetail()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Dim row As DataGridViewRow
        Try
            For Each row In Me.dgvItem.Rows
                row.Cells(Me.ColNo.Index).Value = row.Index + 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub Lbl_LinkClicked(ByVal sender As System.Object, ByVal e As Windows.Forms.LinkLabelLinkClickedEventArgs)
        Dim lnk As Windows.Forms.LinkLabel = CType(sender, Windows.Forms.LinkLabel)
        MsgBox(lnk.Text)
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim FrmCetak As New FrmTrans_Attendance_Cetak
        Dim FrmOpt As New FrmTrans_Attendance_Cetak_Opt
        Dim RowSetting As DseTrans_Attendance_Cetak_IGI.TblSettingRow

        Try
            If Me.TxtPrimary.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtPrimary.Text <> "-1" Then
                FrmCetak.mShowQuickAttendance = False
                FrmCetak.mPrimaryAttendance = Me.TxtPrimary.Text
                FrmCetak.mTblSetting.Rows.Clear()
                RowSetting = FrmCetak.mTblSetting.NewTblSettingRow
                RowSetting.primary_transaksi = Me.TxtPrimary.Text
                RowSetting.no_transaksi = Me.TxtDocNumber.Text
                RowSetting.Employee_Code = Me.TxtEmployeeCode.Text
                RowSetting.start_date = Me.dtpStart.Value.Date
                RowSetting.end_date = Me.DtpEnd.Value.Date
                RowSetting.index = 0
                FrmCetak.mTblSetting.Rows.Add(RowSetting)
                FrmOpt.ShowDialog()
                FrmCetak.FormatCetak = FrmOpt.CboFormat.SelectedIndex
                FrmCetak.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmCetak.Dispose()
            FrmOpt.Dispose()
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.Baru = True Then
                Me.TxtDocNumber.Text = Me.ObjAtt.GetNoBaru(Me.DtpDate.Value)
            Else
                If Me.DtpDate.Value.Month = ObjAtt.GetTanggalTransaksi(Me.TxtPrimary.Text).Month And _
                Me.DtpDate.Value.Year = ObjAtt.GetTanggalTransaksi(Me.TxtPrimary.Text).Year Then
                    Me.TxtDocNumber.Text = Me.ObjAtt.GetNoTransByPrimary(Me.TxtPrimary.Text)
                Else
                    If Me.DtpDate.Value.Month <> Me.ObjAtt.Tanggal.Month Or _
                    Me.DtpDate.Value.Year <> Me.ObjAtt.Tanggal.Year Then
                        Me.TxtDocNumber.Text = Me.ObjAtt.GetNoBaru(Me.DtpDate.Value)
                    Else
                        Me.TxtDocNumber.Text = Me.ObjAtt.GetNoTransByPrimary(Me.TxtPrimary.Text)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnStop_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnStop.Click
        Try
            If MsgBox("Stop creating attendance?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.Abort
                Me.NonAktifkanDetail()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SetKolomPermit()
        Dim rowDGV As DataGridViewRow
        Dim AdaPermit As Boolean = False

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColPermit.Index).Value = Boolean.TrueString Then
                    AdaPermit = True
                    Exit For
                End If
            Next
            If AdaPermit = False Then
                Me.ColPermitIn.Visible = False
                Me.ColPermitOut.Visible = False
                Me.ColPermitDuration.Visible = False
            Else
                Me.ColPermitIn.Visible = True
                Me.ColPermitOut.Visible = True
                Me.ColPermitDuration.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub RefreshEmployeeAttendaceSettings()
        Try
            Me.FingerPrintID = Me.ObjEmp.GetFingerprintID(Me.TxtPrimaryEmployee.Text)
            Me.IsOverrideClockIn = Me.ObjEmp.GetIsOverrideClockIn(Me.TxtPrimaryEmployee.Text)
            Me.DoesNotClockIn = Me.ObjEmp.GetDoesNotClockIn(Me.TxtPrimaryEmployee.Text)
            Me.IsOverrideClockOut = Me.ObjEmp.GetIsOverrideClockOut(Me.TxtPrimaryEmployee.Text)
            Me.DoesNotClockOut = Me.ObjEmp.GetDoesNotClockOut(Me.TxtPrimaryEmployee.Text)

            Me.ObjEmp.GetNewOverrideSettings(Me.TxtPrimaryEmployee.Text, _
                    Me.IsOverrideEarlyOTRule, Me.IsCalculateEarlyOT, Me.IsOverrideLateRule, Me.IsCalculateLate, _
                    Me.IsOverrideEarlyOutRule, Me.IsCalculateEarlyOut, Me.IsOverrideOTRule, Me.IsCalculateOT, _
                    Me.CalculateAfterEarlyOTM, Me.CalculateAfterLateM, Me.CalculateAfterEarlyOutM, Me.CalculateAfterOTM, _
                    Me.RoundingTypeEarlyOTM, Me.RoundingTypeLateM, Me.RoundingTypeEarlyOutM, Me.RoundingTypeOTM, _
                    Me.RoundingValueEarlyOTM, Me.RoundingValueLateM, Me.RoundingValueEarlyOutM, Me.RoundingValueOTM, _
                    Me.MaxValueEarlyOTM, Me.MaxValueLateM, Me.MaxValueEarlyOutM, Me.MaxValueOTM, _
                    Me.CalculateEarlyOTTolerance, Me.CalculateLateTolerance, Me.CalculateEarlyOutTolerance, Me.CalculateOTTolerance, _
                    Me.CalculateEarlyOTOnWeekend, Me.CalculateLateOnWeekend, Me.CalculateEarlyOutOnWeekend, Me.CalculateOTOnWeekend, _
                    Me.CalculateEarlyOTOnHoliday, Me.CalculateLateOnHoliday, Me.CalculateEarlyOutOnHoliday, Me.CalculateOTOnHoliday)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetCalculationType(ByVal vEnumWorkDay As ClsTrans_Attendance.EnumWorkday, ByVal vDate As Date, ByVal vPrimaryWorkShift As Double, _
            ByVal vTbl As DseTrans_Attendance.tblmaster_workshift_itemholiday1DataTable) As FrmTrans_Attendance.EnumCalculationType
        Try
            If vEnumWorkDay = ClsTrans_Attendance.EnumWorkday.Holiday Then
                vTbl.DefaultView.RowFilter = "[date] = #" & vDate & "# AND [primary transaksi] = " & vPrimaryWorkShift.ToString
                If vTbl.DefaultView.Count > 0 Then
                    Return EnumCalculationType.HOLIDAY
                Else
                    Return EnumCalculationType.WEEKEND
                End If
            End If
            Return EnumCalculationType.WORKDAY
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function FillAttendanceFromMachine(ByVal UpdOvertimeOrder As Boolean) As Boolean
        Dim cboDgv As Windows.Forms.DataGridViewComboBoxCell
        Dim workday As ClsTrans_Attendance.EnumWorkday
        Dim dtaHoliday As New DseTrans_AttendanceTableAdapters.tblmaster_workshift_itemholiday1TableAdapter

        Dim ObjSrv As New ClsMasterAttendanceServer
        Dim DtaMachineLog As New DseTrans_AttendanceTableAdapters.DataTable81TableAdapter
        Dim TblMachineLogDummy As New DseTrans_Attendance.DataTable81DataTable
        Dim TblMachineLog As New DseTrans_Attendance.DataTable81DataTable
        Dim RowMachineLog As DseTrans_Attendance.DataTable81Row
        Dim ListServer As New List(Of Double)
        Dim ListActiveTerminalID As New List(Of String)
        Dim ThereIsActiveTerminal As Boolean = False
        Dim IntSchIn, IntSchOut, IntLogTime, IntLogTime2 As Integer
        Dim PrimaryWorkshift_ As Double
        Dim StrLogTime, StrLogTime2 As String

        Dim DtaLogMySQL As New DseTrans_AttendanceTableAdapters.tbltrans_attendancelogTableAdapter
        Dim TblLogMySQL As New DseTrans_Attendance.tbltrans_attendancelogDataTable
        Dim RowLogMySQL As DseTrans_Attendance.tbltrans_attendancelogRow

        Dim IsGetValueForIn, IsGetValueForOut As Boolean
        Dim IntInMinimum, IntInMaximum, IntOutMinimum, IntOutMaximum As Integer
        Dim FirstOrLastIn, FirstOrLastOut As Integer
        Dim IfNoClockIn, IfNoClockOut As Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut
        Dim DefaultCustomIn, DefaultCustomOut As String

        Dim IsSkipMidnight As Boolean = False
        Dim IntNextDay As Integer = 0
        Dim IntNextDayMinute As Integer = 0
        Dim DummyInt As Integer = 0

        Try
            dtaHoliday.Connection.ConnectionString = Core_Module.GetConstring
            dtaHoliday.Fill(Me.tblHoliday, Me.dtpStart.Value.Date, Me.DtpEnd.Value.Date)

            'aul tambahan
            If Me.Baru = True Then
                Me.TxtDocNumber.Text = Me.ObjAtt.GetNoBaru(Me.DtpDate.Value)
            End If

            '''get log from mysql-----------------------------------------------------------------------------------------------------------------
            'DtaLogMySQL.Connection.ConnectionString = Core_Module.GetConstring
            'DtaLogMySQL.Fill_FingerprintIDAndLogDate(TblLogMySQL, FingerPrintID, Format(CDate(Me.dtpStart.Value), "yyyy-MM-dd 00:00:00"), Format(CDate(Me.DtpEnd.Value.AddDays(1)), "yyyy-MM-dd 23:59:59"))
            'If TblLogMySQL.Count > 0 Then
            '    For Each RowLogMySQL In TblLogMySQL
            '        RowMachineLog = TblMachineLog.NewDataTable81Row
            '        RowMachineLog.IndexKey = -1
            '        RowMachineLog.ID = RowLogMySQL.ID
            '        RowMachineLog.Name = ""
            '        RowMachineLog.TerminalID = RowLogMySQL.TerminalID
            '        RowMachineLog.TerminalName = ""
            '        RowMachineLog.TransactionTime = RowLogMySQL.TransactionTime
            '        TblMachineLog.AddDataTable81Row(RowMachineLog)
            '    Next
            'End If

            ''get log from mysql-----------------------------------------------------------------------------------------------------------------
            DtaLogMySQL.Connection.ConnectionString = Core_Module.GetConstring
            DtaLogMySQL.Fill_FingerprintIDAndLogDate(TblLogMySQL, FingerPrintID, Format(CDate(Me.dtpStart.Value), "yyyy-MM-dd 00:00:00"), Format(CDate(Me.DtpEnd.Value.AddDays(1)), "yyyy-MM-dd 23:59:59"))
            If TblLogMySQL.Count > 0 Then
                Dim TglAcuan As Date = #1/1/1950#
                For Each RowLogMySQL In TblLogMySQL
                    If Format(TglAcuan, "dd MMM yyyy HH:mm") <> Format(RowLogMySQL.TransactionTime, "dd MMM yyyy HH:mm") Then
                        TglAcuan = Format(RowLogMySQL.TransactionTime, "dd MMM yyyy HH:mm")

                        RowMachineLog = TblMachineLog.NewDataTable81Row
                        RowMachineLog.IndexKey = -1
                        RowMachineLog.ID = RowLogMySQL.ID
                        RowMachineLog.Name = ""
                        RowMachineLog.TerminalID = RowLogMySQL.TerminalID
                        RowMachineLog.TerminalName = ""
                        RowMachineLog.TransactionTime = TglAcuan
                        TblMachineLog.AddDataTable81Row(RowMachineLog)
                    End If
                Next
            End If

            ''check and reset value in datagridview-------------------------------------------------------------------------------------------------------
            'aul tambahan 

            For Each rowdgv As DataGridViewRow In Me.dgvItem.Rows
                DummyInt += 1
                With rowdgv
                    If Me.JoinWithAttPrev = False Then
                        GoTo ChangeItem
                    End If
                    If CDate(.Cells(Me.ColDate.Index).Value).Date > Me.DtpToAttPrev Then
ChangeItem:
                        PrimaryWorkshift_ = CDbl(.Cells(Me.ColPrimaryWorkShift.Index).Value.ToString)
                        IntSchIn = Me.ObjCd.Convert_ToMinutes(.Cells(Me.ColSchIn.Index).Value.ToString)
                        IntSchOut = Me.ObjCd.Convert_ToMinutes(.Cells(Me.ColSchOut.Index).Value.ToString)
                        IsGetValueForIn = Me.ObjWS.GetIsGetValue(PrimaryWorkshift_, Core_Module.ClsMasterWorkShift.EnumScheduleType.SCHEDULE_IN)
                        IsGetValueForOut = Me.ObjWS.GetIsGetValue(PrimaryWorkshift_, Core_Module.ClsMasterWorkShift.EnumScheduleType.SCHEDULE_OUT)
                        IntInMinimum = Me.ObjWS.GetLimitBefore(PrimaryWorkshift_, Core_Module.ClsMasterWorkShift.EnumScheduleType.SCHEDULE_IN)
                        IntInMaximum = Me.ObjWS.GetLimitAfter(PrimaryWorkshift_, Core_Module.ClsMasterWorkShift.EnumScheduleType.SCHEDULE_IN)
                        IntOutMinimum = Me.ObjWS.GetLimitBefore(PrimaryWorkshift_, Core_Module.ClsMasterWorkShift.EnumScheduleType.SCHEDULE_OUT)
                        IntOutMaximum = Me.ObjWS.GetLimitAfter(PrimaryWorkshift_, Core_Module.ClsMasterWorkShift.EnumScheduleType.SCHEDULE_OUT)
                        FirstOrLastIn = Me.ObjWS.GetFirstOrLast(PrimaryWorkshift_, Core_Module.ClsMasterWorkShift.EnumScheduleType.SCHEDULE_IN)
                        FirstOrLastOut = Me.ObjWS.GetFirstOrLast(PrimaryWorkshift_, Core_Module.ClsMasterWorkShift.EnumScheduleType.SCHEDULE_OUT)
                        IfNoClockIn = Me.ObjWS.GetDefaultIfNoClockInOut(PrimaryWorkshift_, Core_Module.ClsMasterWorkShift.EnumScheduleType.SCHEDULE_IN)
                        IfNoClockOut = Me.ObjWS.GetDefaultIfNoClockInOut(PrimaryWorkshift_, Core_Module.ClsMasterWorkShift.EnumScheduleType.SCHEDULE_OUT)
                        DefaultCustomIn = Me.ObjWS.GetDefaultCustom(PrimaryWorkshift_, Core_Module.ClsMasterWorkShift.EnumScheduleType.SCHEDULE_IN)
                        DefaultCustomOut = Me.ObjWS.GetDefaultCustom(PrimaryWorkshift_, Core_Module.ClsMasterWorkShift.EnumScheduleType.SCHEDULE_OUT)

                        cboDgv = Me.dgvItem.Item(Me.ColWorkdays.Index, rowdgv.Index)
                        workday = cboDgv.Items.IndexOf(Me.dgvItem.Item(Me.ColWorkdays.Index, rowdgv.Index).Value)
                        ApplyAttendanceSettings(rowdgv, PrimaryWorkshift_, Me.GetCalculationType(workday, Me.dgvItem.Item(Me.ColDate.Index, rowdgv.Index).Value, Me.dgvItem.Item(Me.ColPrimaryWorkShift.Index, rowdgv.Index).Value, Me.tblHoliday))

                        ''@rosadinnn
                        ''perubahan filter jika melewati tengah malam
                        ''mengatasi error ambil actual in/out saat batas libur
                        ''range filter in ambil dari jam masuk - limit in down
                        ''range filter out ambil sampai jam keluar + limit out up pada tanggal selanjutnya (hari besok)

                        ''check skip midnight-----------------------------------------------------------------------------------------------------------------
                        Dim filter_ As String = ""
                        If IntSchOut < IntSchIn Then
                            IsSkipMidnight = True
                            IntNextDay = 1
                            IntNextDayMinute = 1440
                            IntSchOut += IntNextDayMinute
                            filter_ = "[TransactionTime] >= #" & Format(CDate(.Cells(Me.ColDate.Index).Value), "yyyy-MM-dd " & ObjCd.Convert_ToString(IntSchIn - IntInMinimum).Replace(".", ":") & ":00") & "# AND [TransactionTime] <= #" & Format(CDate(.Cells(Me.ColDate.Index).Value).AddDays(IntNextDay), "yyyy-MM-dd " & ObjCd.Convert_ToString(IntSchOut - IntNextDayMinute + IntOutMaximum).Replace(".", ":") & ":00") & "#"
                        Else
                            IsSkipMidnight = False
                            IntNextDay = 0
                            IntNextDayMinute = 0
                            IntSchOut += IntNextDayMinute
                            'filter_ = "[TransactionTime] >= #" & Format(CDate(.Cells(Me.ColDate.Index).Value), "yyyy-MM-dd 00:00:00") & "# AND [TransactionTime] <= #" & Format(CDate(.Cells(Me.ColDate.Index).Value).AddDays(IntNextDay), "yyyy-MM-dd 23:59:59") & "#"
                            filter_ = "([TransactionTime] >= #" & Format(CDate(.Cells(Me.ColDate.Index).Value).AddMinutes(IntSchIn).AddMinutes(-1 * IntInMinimum), "yyyy-MM-dd HH:mm:ss") & "# AND [TransactionTime] <= #" & Format(CDate(.Cells(Me.ColDate.Index).Value).AddMinutes(IntSchIn).AddMinutes(IntInMaximum), "yyyy-MM-dd HH:mm:ss") & "#) OR ([TransactionTime] >= #" & Format(CDate(.Cells(Me.ColDate.Index).Value).AddMinutes(IntSchOut).AddMinutes(-1 * IntOutMinimum), "yyyy-MM-dd HH:mm:ss") & "# AND [TransactionTime] <= #" & Format(CDate(.Cells(Me.ColDate.Index).Value).AddMinutes(IntSchOut).AddMinutes(IntOutMaximum), "yyyy-MM-dd HH:mm:ss") & "#)"
                        End If

                        TblMachineLog.DefaultView.RowFilter = filter_

                        ''end of @rosadinnn

                        ''fill log column---------------------------------------------------------------------------------------------------------------------
                        .Cells(Me.ColShowLog.Index).Value = "Show"
                        TblMachineLog.DefaultView.Sort = "[TransactionTime] ASC"
                        If TblMachineLog.DefaultView.Count > 0 Then
                            For i As Integer = 0 To TblMachineLog.DefaultView.Count - 1
                                RowMachineLog = TblMachineLog.DefaultView(i).Row
                                StrLogTime = Format(CDate(RowMachineLog.TransactionTime), "dd MMM yyyy HH:mm")
                                If i < TblMachineLog.DefaultView.Count - 1 Then
                                    If IsNothing(.Cells(Me.ColLogContents.Index).Value) = False Then
                                        If .Cells(Me.ColLogContents.Index).Value.ToString.Contains(StrLogTime) = False Then
                                            .Cells(Me.ColLogContents.Index).Value += StrLogTime & vbNewLine
                                        End If
                                    Else
                                        .Cells(Me.ColLogContents.Index).Value += StrLogTime & vbNewLine
                                    End If
                                Else
                                    If IsNothing(.Cells(Me.ColLogContents.Index).Value) = False Then
                                        If .Cells(Me.ColLogContents.Index).Value.ToString.Contains(StrLogTime) = False Then
                                            .Cells(Me.ColLogContents.Index).Value += StrLogTime
                                        End If
                                    Else
                                        .Cells(Me.ColLogContents.Index).Value += StrLogTime
                                    End If
                                End If
                            Next
                        End If

                        ''leave but present, or present on holiday--------------------------------------------------------------------------------------------
                        If .Cells(Me.ColAttendanceStatus.Index).Value.ToString = ClsTrans_Attendance.EnumStatus.Leave.ToString Or _
                                (.Cells(Me.ColAttendanceStatus.Index).Value.ToString <> ClsTrans_Attendance.EnumStatus.Leave.ToString And _
                                .Cells(Me.ColWorkdays.Index).Value.ToString = ClsTrans_Attendance.EnumWorkday.Holiday.ToString) Then
                            Dim OnlyOneLogAccept As Boolean = False
                            If TblMachineLog.DefaultView.Count = 1 Then
                                RowMachineLog = TblMachineLog.DefaultView(0).Row
                                StrLogTime = Format(CDate(RowMachineLog.TransactionTime), "HH.mm")
                                IntLogTime = Me.ObjCd.Convert_ToMinutes(StrLogTime)
                                If RowMachineLog.TransactionTime > "#" & Format(CDate(.Cells(Me.ColDate.Index).Value), "yyyy-MM-dd 23:59:59") & "#" Then
                                    IntLogTime += IntNextDayMinute
                                End If
                                If IntLogTime < IntSchIn Then
                                    .Cells(Me.ColActualIn.Index).Value = StrLogTime
                                    If IfNoClockOut = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.SCHEDULE Then
                                        .Cells(Me.ColActualOut.Index).Value = .Cells(Me.ColSchOut.Index).Value
                                    ElseIf IfNoClockOut = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM Then
                                        .Cells(Me.ColActualOut.Index).Value = DefaultCustomOut
                                    End If
                                    If .Cells(Me.ColAttendanceStatus.Index).Value.ToString <> ClsTrans_Attendance.EnumStatus.Leave.ToString And .Cells(Me.ColWorkdays.Index).Value.ToString = ClsTrans_Attendance.EnumWorkday.Holiday.ToString Then
                                        Me.MarkRow(Me.dgvItem.Rows.IndexOf(rowdgv), ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_OUT, IfNoClockOut)
                                    End If

                                ElseIf IntLogTime > IntSchOut Then
                                    .Cells(Me.ColActualOut.Index).Value = StrLogTime
                                    If IfNoClockIn = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.SCHEDULE Then
                                        .Cells(Me.ColActualIn.Index).Value = .Cells(Me.ColSchIn.Index).Value
                                    ElseIf IfNoClockIn = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM Then
                                        .Cells(Me.ColActualIn.Index).Value = DefaultCustomIn
                                    End If
                                    If .Cells(Me.ColAttendanceStatus.Index).Value.ToString <> ClsTrans_Attendance.EnumStatus.Leave.ToString And .Cells(Me.ColWorkdays.Index).Value.ToString = ClsTrans_Attendance.EnumWorkday.Holiday.ToString Then
                                        Me.MarkRow(Me.dgvItem.Rows.IndexOf(rowdgv), ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_IN, IfNoClockIn)
                                    End If
                                Else
                                    If IntLogTime - IntSchIn <= IntSchOut - IntLogTime Then
                                        .Cells(Me.ColActualIn.Index).Value = StrLogTime
                                        If IfNoClockOut = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.SCHEDULE Then
                                            .Cells(Me.ColActualOut.Index).Value = .Cells(Me.ColSchOut.Index).Value
                                        ElseIf IfNoClockOut = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM Then
                                            .Cells(Me.ColActualOut.Index).Value = DefaultCustomOut
                                        End If
                                        If .Cells(Me.ColAttendanceStatus.Index).Value.ToString <> ClsTrans_Attendance.EnumStatus.Leave.ToString And .Cells(Me.ColWorkdays.Index).Value.ToString = ClsTrans_Attendance.EnumWorkday.Holiday.ToString Then
                                            Me.MarkRow(Me.dgvItem.Rows.IndexOf(rowdgv), ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_OUT, IfNoClockOut)
                                        End If
                                    Else
                                        .Cells(Me.ColActualOut.Index).Value = StrLogTime
                                        If IfNoClockIn = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.SCHEDULE Then
                                            .Cells(Me.ColActualIn.Index).Value = .Cells(Me.ColSchIn.Index).Value
                                        ElseIf IfNoClockIn = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM Then
                                            .Cells(Me.ColActualIn.Index).Value = DefaultCustomIn
                                        End If
                                        If .Cells(Me.ColAttendanceStatus.Index).Value.ToString <> ClsTrans_Attendance.EnumStatus.Leave.ToString And .Cells(Me.ColWorkdays.Index).Value.ToString = ClsTrans_Attendance.EnumWorkday.Holiday.ToString Then
                                            Me.MarkRow(Me.dgvItem.Rows.IndexOf(rowdgv), ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_IN, IfNoClockIn)
                                        End If
                                    End If
                                End If
                            ElseIf TblMachineLog.DefaultView.Count > 1 Then
                                RowMachineLog = TblMachineLog.DefaultView(0).Row
                                StrLogTime = Format(CDate(RowMachineLog.TransactionTime), "HH.mm")
                                IntLogTime = Me.ObjCd.Convert_ToMinutes(StrLogTime)
                                If RowMachineLog.TransactionTime > "#" & Format(CDate(.Cells(Me.ColDate.Index).Value), "yyyy-MM-dd 23:59:59") & "#" Then
                                    IntLogTime += IntNextDayMinute
                                End If
                                RowMachineLog = TblMachineLog.DefaultView(TblMachineLog.DefaultView.Count - 1).Row
                                StrLogTime2 = Format(CDate(RowMachineLog.TransactionTime), "HH.mm")
                                IntLogTime2 = Me.ObjCd.Convert_ToMinutes(StrLogTime2)
                                If RowMachineLog.TransactionTime > "#" & Format(CDate(.Cells(Me.ColDate.Index).Value), "yyyy-MM-dd 23:59:59") & "#" Then
                                    IntLogTime2 += IntNextDayMinute
                                End If
                                If (IntLogTime2 + IntNextDayMinute - IntLogTime) >= Me.DurationMin Then
                                    .Cells(Me.ColActualIn.Index).Value = StrLogTime
                                    .Cells(Me.ColActualOut.Index).Value = StrLogTime2
                                Else
                                    OnlyOneLogAccept = True
                                    If IntLogTime < IntSchIn Then
                                        .Cells(Me.ColActualIn.Index).Value = StrLogTime
                                        If IfNoClockOut = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.SCHEDULE Then
                                            .Cells(Me.ColActualOut.Index).Value = .Cells(Me.ColSchOut.Index).Value
                                        ElseIf IfNoClockOut = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM Then
                                            .Cells(Me.ColActualOut.Index).Value = DefaultCustomOut
                                        End If
                                        If .Cells(Me.ColAttendanceStatus.Index).Value.ToString <> ClsTrans_Attendance.EnumStatus.Leave.ToString And .Cells(Me.ColWorkdays.Index).Value.ToString = ClsTrans_Attendance.EnumWorkday.Holiday.ToString Then
                                            Me.MarkRow(Me.dgvItem.Rows.IndexOf(rowdgv), ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_OUT, IfNoClockOut)
                                        End If
                                    ElseIf IntLogTime > IntSchOut Then
                                        .Cells(Me.ColActualOut.Index).Value = StrLogTime
                                        If IfNoClockIn = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.SCHEDULE Then
                                            .Cells(Me.ColActualIn.Index).Value = .Cells(Me.ColSchIn.Index).Value
                                        ElseIf IfNoClockIn = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM Then
                                            .Cells(Me.ColActualIn.Index).Value = DefaultCustomIn
                                        End If
                                        If .Cells(Me.ColAttendanceStatus.Index).Value.ToString <> ClsTrans_Attendance.EnumStatus.Leave.ToString And .Cells(Me.ColWorkdays.Index).Value.ToString = ClsTrans_Attendance.EnumWorkday.Holiday.ToString Then
                                            Me.MarkRow(Me.dgvItem.Rows.IndexOf(rowdgv), ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_IN, IfNoClockIn)
                                        End If
                                    Else
                                        If IntLogTime - IntSchIn <= IntSchOut - IntLogTime Then
                                            .Cells(Me.ColActualIn.Index).Value = StrLogTime
                                            If IfNoClockOut = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.SCHEDULE Then
                                                .Cells(Me.ColActualOut.Index).Value = .Cells(Me.ColSchOut.Index).Value
                                            ElseIf IfNoClockOut = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM Then
                                                .Cells(Me.ColActualOut.Index).Value = DefaultCustomOut
                                            End If
                                            If .Cells(Me.ColAttendanceStatus.Index).Value.ToString <> ClsTrans_Attendance.EnumStatus.Leave.ToString And .Cells(Me.ColWorkdays.Index).Value.ToString = ClsTrans_Attendance.EnumWorkday.Holiday.ToString Then
                                                Me.MarkRow(Me.dgvItem.Rows.IndexOf(rowdgv), ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_OUT, IfNoClockOut)
                                            End If
                                        Else
                                            .Cells(Me.ColActualOut.Index).Value = StrLogTime
                                            If IfNoClockIn = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.SCHEDULE Then
                                                .Cells(Me.ColActualIn.Index).Value = .Cells(Me.ColSchIn.Index).Value
                                            ElseIf IfNoClockIn = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM Then
                                                .Cells(Me.ColActualIn.Index).Value = DefaultCustomIn
                                            End If
                                            If .Cells(Me.ColAttendanceStatus.Index).Value.ToString <> ClsTrans_Attendance.EnumStatus.Leave.ToString And .Cells(Me.ColWorkdays.Index).Value.ToString = ClsTrans_Attendance.EnumWorkday.Holiday.ToString Then
                                                Me.MarkRow(Me.dgvItem.Rows.IndexOf(rowdgv), ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_IN, IfNoClockIn)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            If .Cells(Me.ColAttendanceStatus.Index).Value.ToString = ClsTrans_Attendance.EnumStatus.Leave.ToString And TblMachineLog.DefaultView.Count > 0 Then
                                Me.MarkRow(Me.dgvItem.Rows.IndexOf(rowdgv), ClsTrans_Attendance.EnumRemarks.LEAVE_BUT_PRESENT, Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM)
                            End If
                            If (.Cells(Me.ColAttendanceStatus.Index).Value.ToString <> ClsTrans_Attendance.EnumStatus.Leave.ToString And _
                                .Cells(Me.ColWorkdays.Index).Value.ToString = ClsTrans_Attendance.EnumWorkday.Holiday.ToString) And _
                                TblMachineLog.DefaultView.Count > 0 Then
                                .Cells(Me.ColAttendanceStatus.Index).Value = "Present"
                            End If
                            'ElseIf (.Cells(Me.ColAttendanceStatus.Index).Value.ToString <> ClsTrans_Attendance.EnumStatus.Leave.ToString _
                            '       And .Cells(Me.ColAttendanceStatus.Index).Value.ToString <> "Half-day Leave" _
                            '       And .Cells(Me.ColWorkdays.Index).Value.ToString = ClsTrans_Attendance.EnumWorkday.Workday.ToString) Then
                        ElseIf (.Cells(Me.ColAttendanceStatus.Index).Value.ToString <> ClsTrans_Attendance.EnumStatus.Leave.ToString _
                           And .Cells(Me.ColAttendanceStatus.Index).Value.ToString <> "Half-day Leave" _
                           And .Cells(Me.ColWorkdays.Index).Value.ToString = ClsTrans_Attendance.EnumWorkday.Workday.ToString) Then
                            ''get actual in-------------------------------------------------------------------------------------------------------------------
                            If IntSchIn + IntInMaximum < 1440 Then
                                If IntSchIn - IntInMinimum < 0 Then
                                    ''jika batas bawah schin mundur ke hari sebelumnya
                                    TblMachineLog.DefaultView.RowFilter = "[TransactionTime] >= #" & Format(CDate(.Cells(Me.ColDate.Index).Value).AddDays(-1), "yyyy-MM-dd " & Me.ObjCd.Convert_ToString((IntSchIn + Me.ObjCd.Convert_ToMinutes("23.59") + 1) - IntInMinimum).Replace(".", ":") & ":00") & _
                                            "# AND [TransactionTime] <= #" & Format(CDate(.Cells(Me.ColDate.Index).Value), "yyyy-MM-dd " & Me.ObjCd.Convert_ToString(IntSchIn + IntInMaximum).Replace(".", ":") & ":00") & "#"
                                Else
                                    TblMachineLog.DefaultView.RowFilter = "[TransactionTime] >= #" & Format(CDate(.Cells(Me.ColDate.Index).Value), "yyyy-MM-dd " & Me.ObjCd.Convert_ToString(IntSchIn - IntInMinimum).Replace(".", ":") & ":00") & _
                                            "# AND [TransactionTime] <= #" & Format(CDate(.Cells(Me.ColDate.Index).Value), "yyyy-MM-dd " & Me.ObjCd.Convert_ToString(IntSchIn + IntInMaximum).Replace(".", ":") & ":00") & "#"
                                End If
                            Else
                                ''jika batas atas schin melewati malam hari
                                TblMachineLog.DefaultView.RowFilter = "[TransactionTime] >= #" & Format(CDate(.Cells(Me.ColDate.Index).Value), "yyyy-MM-dd " & Me.ObjCd.Convert_ToString(IntSchIn - IntInMinimum).Replace(".", ":") & ":00") & _
                                    "# AND [TransactionTime] <= #" & Format(CDate(.Cells(Me.ColDate.Index).Value).AddDays(IntNextDay), "yyyy-MM-dd " & Me.ObjCd.Convert_ToString(IntSchIn + IntInMaximum - IntNextDayMinute).Replace(".", ":") & ":00") & "#"
                            End If
                            If TblMachineLog.DefaultView.Count = 0 Then
                                If IsOverrideClockIn = True Then
                                    If DoesNotClockIn = True Then
                                        .Cells(Me.ColActualIn.Index).Value = .Cells(Me.ColSchIn.Index).Value
                                    Else
                                        If IfNoClockIn = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.SCHEDULE Then
                                            .Cells(Me.ColActualIn.Index).Value = .Cells(Me.ColSchIn.Index).Value
                                        ElseIf IfNoClockIn = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM Then
                                            .Cells(Me.ColActualIn.Index).Value = DefaultCustomIn
                                            Me.MarkRow(Me.dgvItem.Rows.IndexOf(rowdgv), ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_IN, Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM)
                                        End If
                                    End If
                                Else
                                    If IfNoClockIn = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.SCHEDULE Then
                                        .Cells(Me.ColActualIn.Index).Value = .Cells(Me.ColSchIn.Index).Value
                                    ElseIf IfNoClockIn = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM Then
                                        .Cells(Me.ColActualIn.Index).Value = DefaultCustomIn
                                        Me.MarkRow(Me.dgvItem.Rows.IndexOf(rowdgv), ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_IN, Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM)
                                    End If
                                End If
                            ElseIf TblMachineLog.DefaultView.Count = 1 Then
                                RowMachineLog = TblMachineLog.DefaultView(0).Row
                                StrLogTime = Format(CDate(RowMachineLog.TransactionTime), "HH.mm")
                                IntLogTime = Me.ObjCd.Convert_ToMinutes(StrLogTime)
                                If IsOverrideClockIn = True Then
                                    If DoesNotClockIn = True Then
                                        .Cells(Me.ColActualIn.Index).Value = .Cells(Me.ColSchIn.Index).Value
                                    Else
                                        .Cells(Me.ColActualIn.Index).Value = StrLogTime
                                    End If
                                Else
                                    If IsGetValueForIn = True Then
                                        .Cells(Me.ColActualIn.Index).Value = StrLogTime
                                    Else
                                        If IfNoClockIn = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.SCHEDULE Then
                                            .Cells(Me.ColActualIn.Index).Value = .Cells(Me.ColSchIn.Index).Value
                                        ElseIf IfNoClockIn = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM Then
                                            .Cells(Me.ColActualIn.Index).Value = DefaultCustomIn
                                            Me.MarkRow(Me.dgvItem.Rows.IndexOf(rowdgv), ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_IN, Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM)
                                        End If
                                    End If
                                End If
                            ElseIf TblMachineLog.DefaultView.Count > 1 Then
                                If FirstOrLastIn = 0 Then
                                    RowMachineLog = TblMachineLog.DefaultView(0).Row
                                Else
                                    RowMachineLog = TblMachineLog.DefaultView(TblMachineLog.DefaultView.Count - 1).Row
                                End If
                                StrLogTime = Format(CDate(RowMachineLog.TransactionTime), "HH.mm")
                                If IsOverrideClockIn = True Then
                                    If DoesNotClockIn = True Then
                                        .Cells(Me.ColActualIn.Index).Value = .Cells(Me.ColSchIn.Index).Value
                                    Else
                                        .Cells(Me.ColActualIn.Index).Value = StrLogTime
                                    End If
                                Else
                                    If IsGetValueForIn = True Then
                                        .Cells(Me.ColActualIn.Index).Value = StrLogTime
                                    Else
                                        If IfNoClockIn = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.SCHEDULE Then
                                            .Cells(Me.ColActualIn.Index).Value = .Cells(Me.ColSchIn.Index).Value
                                        ElseIf IfNoClockIn = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM Then
                                            .Cells(Me.ColActualIn.Index).Value = DefaultCustomIn
                                            Me.MarkRow(Me.dgvItem.Rows.IndexOf(rowdgv), ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_IN, Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM)
                                        End If
                                    End If
                                End If
                            End If

                            ''get actual out------------------------------------------------------------------------------------------------------------------
                            If (IntSchOut - IntNextDayMinute) - IntOutMinimum < 0 Then
                                TblMachineLog.DefaultView.RowFilter = Nothing
                                TblMachineLog.DefaultView.RowFilter = "[TransactionTime] >= #" & Format(CDate(.Cells(Me.ColDate.Index).Value), "yyyy-MM-dd " & Me.ObjCd.Convert_ToString(IntSchOut - IntOutMinimum).Replace(".", ":") & ":00") & _
                                    "# AND [TransactionTime] <= #" & Format(CDate(.Cells(Me.ColDate.Index).Value).AddDays(IntNextDay), "yyyy-MM-dd " & Me.ObjCd.Convert_ToString(IntSchOut - IntNextDayMinute + IntOutMaximum).Replace(".", ":") & ":00") & "#"
                                ''jika batas bawah schout mundur ke hari sebelumnya (hari dia masuk), ini kemungkinan karena schout pindah hari
                            Else
                                If (IntSchOut + IntOutMaximum) > (Me.ObjCd.Convert_ToMinutes("23.59") + 1) Then
                                    If IntSchOut - IntOutMinimum > 1440 Then
                                        ''jika schout, batas atasnya melewati tengah malam
                                        TblMachineLog.DefaultView.RowFilter = "[TransactionTime] >= #" & Format(CDate(.Cells(Me.ColDate.Index).Value).AddDays(IntNextDay), "yyyy-MM-dd " & Me.ObjCd.Convert_ToString(IntSchOut - IntOutMinimum - (Me.ObjCd.Convert_ToMinutes("23.59") + 1)).Replace(".", ":") & ":00") & _
                                            "# AND [TransactionTime] <= #" & Format(CDate(.Cells(Me.ColDate.Index).Value).AddDays(1), "yyyy-MM-dd " & Me.ObjCd.Convert_ToString((IntSchOut + IntOutMaximum) - (Me.ObjCd.Convert_ToMinutes("23.59") + 1)).Replace(".", ":") & ":00") & "#"
                                    Else
                                        TblMachineLog.DefaultView.RowFilter = "[TransactionTime] >= #" & Format(CDate(.Cells(Me.ColDate.Index).Value).AddDays(IntNextDay), "yyyy-MM-dd " & Me.ObjCd.Convert_ToString(IntSchOut - IntOutMinimum).Replace(".", ":") & ":00") & _
                                            "# AND [TransactionTime] <= #" & Format(CDate(.Cells(Me.ColDate.Index).Value).AddDays(1), "yyyy-MM-dd " & Me.ObjCd.Convert_ToString((IntSchOut + IntOutMaximum) - (Me.ObjCd.Convert_ToMinutes("23.59") + 1)).Replace(".", ":") & ":00") & "#"
                                    End If
                                Else
                                    TblMachineLog.DefaultView.RowFilter = "[TransactionTime] >= #" & Format(CDate(.Cells(Me.ColDate.Index).Value).AddDays(IntNextDay), "yyyy-MM-dd " & Me.ObjCd.Convert_ToString(IntSchOut - IntOutMinimum).Replace(".", ":") & ":00") & _
                                            "# AND [TransactionTime] <= #" & Format(CDate(.Cells(Me.ColDate.Index).Value).AddDays(IntNextDay), "yyyy-MM-dd " & Me.ObjCd.Convert_ToString(IntSchOut + IntOutMaximum).Replace(".", ":") & ":00") & "#"
                                End If
                            End If


                            If TblMachineLog.DefaultView.Count = 0 Then
                                If IsOverrideClockOut = True Then
                                    If DoesNotClockOut = True Then
                                        .Cells(Me.ColActualOut.Index).Value = .Cells(Me.ColSchOut.Index).Value
                                    Else
                                        If IfNoClockOut = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.SCHEDULE Then
                                            .Cells(Me.ColActualOut.Index).Value = .Cells(Me.ColSchOut.Index).Value
                                        ElseIf IfNoClockOut = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM Then
                                            .Cells(Me.ColActualOut.Index).Value = DefaultCustomOut
                                            Me.MarkRow(Me.dgvItem.Rows.IndexOf(rowdgv), ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_OUT, Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM)
                                        End If
                                    End If
                                Else
                                    If IfNoClockOut = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.SCHEDULE Then
                                        .Cells(Me.ColActualOut.Index).Value = .Cells(Me.ColSchOut.Index).Value
                                    ElseIf IfNoClockOut = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM Then
                                        .Cells(Me.ColActualOut.Index).Value = DefaultCustomOut
                                        Me.MarkRow(Me.dgvItem.Rows.IndexOf(rowdgv), ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_OUT, Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM)
                                    End If
                                End If
                            ElseIf TblMachineLog.DefaultView.Count = 1 Then
                                RowMachineLog = TblMachineLog.DefaultView(0).Row
                                StrLogTime = Format(CDate(RowMachineLog.TransactionTime), "HH.mm")
                                IntLogTime = Me.ObjCd.Convert_ToMinutes(StrLogTime)
                                If IsOverrideClockOut = True Then
                                    If DoesNotClockOut = True Then
                                        .Cells(Me.ColActualOut.Index).Value = .Cells(Me.ColSchOut.Index).Value
                                    Else
                                        .Cells(Me.ColActualOut.Index).Value = StrLogTime
                                    End If
                                Else
                                    If IsGetValueForOut = True Then
                                        .Cells(Me.ColActualOut.Index).Value = StrLogTime
                                    Else
                                        If IfNoClockOut = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.SCHEDULE Then
                                            .Cells(Me.ColActualOut.Index).Value = .Cells(Me.ColSchOut.Index).Value
                                        ElseIf IfNoClockOut = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM Then
                                            .Cells(Me.ColActualOut.Index).Value = DefaultCustomOut
                                            Me.MarkRow(Me.dgvItem.Rows.IndexOf(rowdgv), ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_OUT, Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM)
                                        End If
                                    End If
                                End If
                            ElseIf TblMachineLog.DefaultView.Count > 1 Then
                                If FirstOrLastOut = 0 Then
                                    RowMachineLog = TblMachineLog.DefaultView(0).Row
                                Else
                                    RowMachineLog = TblMachineLog.DefaultView(TblMachineLog.DefaultView.Count - 1).Row
                                End If
                                StrLogTime = Format(CDate(RowMachineLog.TransactionTime), "HH.mm")
                                If IsOverrideClockOut = True Then
                                    If DoesNotClockOut = True Then
                                        .Cells(Me.ColActualOut.Index).Value = .Cells(Me.ColSchOut.Index).Value
                                    Else
                                        .Cells(Me.ColActualOut.Index).Value = StrLogTime
                                    End If
                                Else
                                    If IsGetValueForOut = True Then
                                        .Cells(Me.ColActualOut.Index).Value = StrLogTime
                                    Else
                                        If IfNoClockOut = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.SCHEDULE Then
                                            .Cells(Me.ColActualOut.Index).Value = .Cells(Me.ColSchOut.Index).Value
                                        ElseIf IfNoClockOut = Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM Then
                                            .Cells(Me.ColActualOut.Index).Value = DefaultCustomOut
                                            Me.MarkRow(Me.dgvItem.Rows.IndexOf(rowdgv), ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_OUT, Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM)
                                        End If
                                    End If
                                End If
                            End If

                        End If

                        ''recalculate others column if act in&out is valid------------------------------------------------------------------------------------
                        If Me.ObjCd.Is_Time_String(.Cells(Me.ColActualIn.Index).Value.ToString) = True And Me.ObjCd.Is_Time_String(.Cells(Me.ColActualOut.Index).Value.ToString) = True Then
                            Me.Hitung(Me.dgvItem.Rows.IndexOf(rowdgv), -1, EnumCalculateRange.BOTH, False)
                        End If
                    End If
                End With
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message + " - " + DummyInt.ToString, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Function

    Private Sub MarkRow(ByVal vIndex As Integer, ByVal vRemarks As ClsTrans_Attendance.EnumRemarks, ByVal vIfNoClock As Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut)
        Try
            Dim DgvRow As DataGridViewRow
            DgvRow = Me.dgvItem.Rows(vIndex)

            If vIfNoClock <> Core_Module.ClsMasterWorkShift.EnumDefaultIfNoClockInOut.SCHEDULE Then
                DgvRow.DefaultCellStyle.BackColor = WarningColor
                If IsNothing(DgvRow.Cells(Me.ColRemarks.Index).Value) = False Then
                    If DgvRow.Cells(Me.ColRemarks.Index).Value.ToString <> "" Then
                        If vRemarks = ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_IN Then
                            DgvRow.Cells(Me.ColRemarks.Index).Value += "; No actual in"
                            DgvRow.Cells(Me.ColOnSchedule.Index).Value = "On Schedule"
                        ElseIf vRemarks = ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_OUT Then
                            DgvRow.Cells(Me.ColRemarks.Index).Value += "; No actual out"
                            DgvRow.Cells(Me.ColOnSchedule.Index).Value = "On Schedule"
                        ElseIf vRemarks = ClsTrans_Attendance.EnumRemarks.LEAVE_BUT_PRESENT Then
                            DgvRow.Cells(Me.ColRemarks.Index).Value += "; Leave but present"
                        End If
                    Else
                        If vRemarks = ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_IN Then
                            DgvRow.Cells(Me.ColRemarks.Index).Value = "No actual in"
                            DgvRow.Cells(Me.ColOnSchedule.Index).Value = "On Schedule"
                        ElseIf vRemarks = ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_OUT Then
                            DgvRow.Cells(Me.ColRemarks.Index).Value = "No actual out"
                            DgvRow.Cells(Me.ColOnSchedule.Index).Value = "On Schedule"
                        ElseIf vRemarks = ClsTrans_Attendance.EnumRemarks.LEAVE_BUT_PRESENT Then
                            DgvRow.Cells(Me.ColRemarks.Index).Value = "Leave but present"
                        End If
                    End If
                Else
                    If vRemarks = ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_IN Then
                        DgvRow.Cells(Me.ColRemarks.Index).Value = "No actual in"
                        DgvRow.Cells(Me.ColOnSchedule.Index).Value = "On Schedule"
                    ElseIf vRemarks = ClsTrans_Attendance.EnumRemarks.NO_ACTUAL_OUT Then
                        DgvRow.Cells(Me.ColRemarks.Index).Value = "No actual out"
                        DgvRow.Cells(Me.ColOnSchedule.Index).Value = "On Schedule"
                    ElseIf vRemarks = ClsTrans_Attendance.EnumRemarks.LEAVE_BUT_PRESENT Then
                        DgvRow.Cells(Me.ColRemarks.Index).Value = "Leave but present"
                    End If
                End If
                Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Color.White
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub SettingKolom()
        Dim ShowOT1, ShowOT2 As Boolean
        Try
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowSchedule, Me.UserCode) = "1" Then
                Me.ColSchBreakIn.Visible = True
                Me.ColSchBreakOut.Visible = True
                Me.ColSchIn.Visible = True
                Me.ColSchOut.Visible = True
            Else
                Me.ColSchBreakIn.Visible = False
                Me.ColSchBreakOut.Visible = False
                Me.ColSchIn.Visible = False
                Me.ColSchOut.Visible = False
            End If
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
           Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowActualBreak, Me.UserCode) = "1" Then
                Me.ColActualBreakIn.Visible = True
                Me.ColActualBreakOut.Visible = True
            Else
                Me.ColActualBreakIn.Visible = False
                Me.ColActualBreakOut.Visible = False
            End If
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowEartlyOT, Me.UserCode) = "1" Then
                Me.ColEarlyOTDuration.Visible = True
                ShowOT1 = True
            Else
                Me.ColEarlyOTDuration.Visible = False
                ShowOT1 = False
            End If
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
           Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowOT, Me.UserCode) = "1" Then
                Me.ColOTDuration.Visible = True
                ShowOT2 = True
            Else
                Me.ColOTDuration.Visible = False
                ShowOT2 = False
            End If
            If ShowOT1 = False And ShowOT2 = False Then
                Me.ColOT.Visible = False
            Else
                Me.ColOT.Visible = True
            End If
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowEarlyOut, Me.UserCode) = "1" Then
                Me.ColEarlyOutDuration.Visible = True
            Else
                Me.ColEarlyOutDuration.Visible = False
            End If
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowLate, Me.UserCode) = "1" Then
                Me.ColLateDuration.Visible = True
            Else
                Me.ColLateDuration.Visible = False
            End If

            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowCalculateEarlyOT, Me.UserCode) = "1" Then
                Me.ColIsCalculateEarlyOT.Visible = True
            Else
                Me.ColIsCalculateEarlyOT.Visible = False
            End If
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowCalculateLate, Me.UserCode) = "1" Then
                Me.ColIsCalculateLate.Visible = True
            Else
                Me.ColIsCalculateLate.Visible = False
            End If
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowCalculateEarlyOut, Me.UserCode) = "1" Then
                Me.ColIsCalculateEarlyOut.Visible = True
            Else
                Me.ColIsCalculateEarlyOut.Visible = False
            End If
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowCalculateOT, Me.UserCode) = "1" Then
                Me.ColIsCalculateOT.Visible = True
            Else
                Me.ColIsCalculateOT.Visible = False
            End If

            'aul -penambahan outstation-
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_ShowOutstation, Me.UserCode) = "1" Then
                Me.ColIsOutstation.Visible = True
            Else
                Me.ColIsOutstation.Visible = False
            End If

            Me.CekVisibleCol_OTOrder()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbColumnSettings_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbColumnSettings.LinkClicked
        Dim ij As New FrmTrans_Attendance_ColumnsSettings
        Try
            ij.UserName = Me.UserCode
            ij.ShowDialog()
            Me.SettingKolom()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnRepairLeave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRepairLeave.Click
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColDate.Index).Value = #7/21/2015# Then
                    Me.ObjAtt.InitLeave(Me, rowDGV, #7/21/2015#)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub ApplyAttendanceSettings(ByVal rowdgv As DataGridViewRow, ByVal PrimaryWorkshift_ As Double, ByVal vCalculationType As FrmTrans_Attendance.EnumCalculationType)
        Dim CboDgvEarlyOT As Windows.Forms.DataGridViewComboBoxCell
        Dim CboDgvLate As Windows.Forms.DataGridViewComboBoxCell
        Dim CboDgvEarlyOut As Windows.Forms.DataGridViewComboBoxCell
        Dim CboDgvOT As Windows.Forms.DataGridViewComboBoxCell

        Dim finalCalculateEarlyOT, finalCalculateLate, finalCalculateEarlyOut, finalCalculateOT As Boolean
        Dim finalCalculateEarlyOTWS, finalCalculateLateWS, finalCalculateEarlyOutWS, finalCalculateOTWS As Boolean

        Try
            With rowdgv
                Me.ObjWS.GetAttendaceSettings(PrimaryWorkshift_, IsCalculateEarlyOTWS, IsCalculateLateWS, IsCalculateEarlyOutWS, IsCalculateOTWS, _
                        Me.CalculateAfterEarlyOTWS, Me.CalculateAfterLateWS, Me.CalculateAfterEarlyOutWS, Me.CalculateAfterOTWS, _
                        Me.RoundingTypeEarlyOTWS, Me.RoundingTypeLateWS, Me.RoundingTypeEarlyOutWS, Me.RoundingTypeOTWS, _
                        Me.RoundingValueEarlyOTWS, Me.RoundingValueLateWS, Me.RoundingValueEarlyOutWS, Me.RoundingValueOTWS, _
                        Me.MaxValueEarlyOTWS, Me.MaxValueLateWS, Me.MaxValueEarlyOutWS, Me.MaxValueOTWS, _
                        Me.CalculateEarlyOTToleranceWS, Me.CalculateLateToleranceWS, Me.CalculateEarlyOutToleranceWS, Me.CalculateOTToleranceWS, _
                        Me.CalculateEarlyOTOnWeekendWS, Me.CalculateLateOnWeekendWS, Me.CalculateEarlyOutOnWeekendWS, Me.CalculateOTOnWeekendWS, _
                        Me.CalculateEarlyOTOnHolidayWS, Me.CalculateLateOnHolidayWS, Me.CalculateEarlyOutOnHolidayWS, Me.CalculateOTOnHolidayWS)

                CboDgvEarlyOT = rowdgv.Cells(Me.ColRoundingTypeEarlyOT.Index)
                CboDgvLate = rowdgv.Cells(Me.ColRoundingTypeLate.Index)
                CboDgvEarlyOut = rowdgv.Cells(Me.ColRoundingTypeEarlyOut.Index)
                CboDgvOT = rowdgv.Cells(Me.ColRoundingTypeOT.Index)

                finalCalculateEarlyOT = Me.IsCalculateEarlyOT
                finalCalculateEarlyOTWS = Me.IsCalculateEarlyOTWS
                finalCalculateLate = Me.IsCalculateLate
                finalCalculateLateWS = Me.IsCalculateLateWS
                finalCalculateEarlyOut = Me.IsCalculateEarlyOut
                finalCalculateEarlyOutWS = Me.IsCalculateEarlyOutWS
                finalCalculateOT = Me.IsCalculateOT
                finalCalculateOTWS = Me.IsCalculateOTWS
                If vCalculationType = EnumCalculationType.WEEKEND Then
                    finalCalculateEarlyOT = Me.CalculateEarlyOTOnWeekend
                    finalCalculateEarlyOTWS = Me.CalculateEarlyOTOnWeekendWS
                    finalCalculateLate = Me.CalculateLateOnWeekend
                    finalCalculateLateWS = Me.CalculateLateOnWeekendWS
                    finalCalculateEarlyOut = Me.CalculateEarlyOutOnWeekend
                    finalCalculateEarlyOutWS = Me.CalculateEarlyOutOnWeekendWS
                    finalCalculateOT = Me.CalculateOTOnWeekend
                    finalCalculateOTWS = Me.CalculateOTOnWeekendWS
                ElseIf vCalculationType = EnumCalculationType.HOLIDAY Then
                    finalCalculateEarlyOT = Me.CalculateEarlyOTOnHoliday
                    finalCalculateEarlyOTWS = Me.CalculateEarlyOTOnHolidayWS
                    finalCalculateLate = Me.CalculateLateOnHoliday
                    finalCalculateLateWS = Me.CalculateLateOnHolidayWS
                    finalCalculateEarlyOut = Me.CalculateEarlyOutOnHoliday
                    finalCalculateEarlyOutWS = Me.CalculateEarlyOutOnHolidayWS
                    finalCalculateOT = Me.CalculateOTOnHoliday
                    finalCalculateOTWS = Me.CalculateOTOnHolidayWS
                End If

                If Me.IsOverrideEarlyOTRule = True Then
                    If finalCalculateEarlyOT = True Then
                        .Cells(Me.ColIsCalculateEarlyOT.Index).Value = True
                    Else
                        .Cells(Me.ColIsCalculateEarlyOT.Index).Value = False
                    End If
                    .Cells(Me.ColCalculateAfterEarlyOT.Index).Value = Me.ObjCd.Convert_ToString(Me.CalculateAfterEarlyOTM)
                    .Cells(Me.ColRoundingTypeEarlyOT.Index).Value = CboDgvEarlyOT.Items.Item(Me.RoundingTypeEarlyOTM)
                    .Cells(Me.ColRoundingValueEarlyOT.Index).Value = Me.ObjCd.Convert_ToString(Me.RoundingValueEarlyOTM)
                    .Cells(Me.ColMaxValueEarlyOT.Index).Value = Me.ObjCd.Convert_ToString(Me.MaxValueEarlyOTM)
                    .Cells(Me.ColCalculateEarlyOTTolerance.Index).Value = Me.CalculateEarlyOTTolerance
                Else
                    If finalCalculateEarlyOTWS = True Then
                        .Cells(Me.ColIsCalculateEarlyOT.Index).Value = True
                    Else
                        .Cells(Me.ColIsCalculateEarlyOT.Index).Value = False
                    End If
                    .Cells(Me.ColCalculateAfterEarlyOT.Index).Value = Me.ObjCd.Convert_ToString(Me.CalculateAfterEarlyOTWS)
                    .Cells(Me.ColRoundingTypeEarlyOT.Index).Value = CboDgvEarlyOT.Items.Item(Me.RoundingTypeEarlyOTWS)
                    .Cells(Me.ColRoundingValueEarlyOT.Index).Value = Me.ObjCd.Convert_ToString(Me.RoundingValueEarlyOTWS)
                    .Cells(Me.ColMaxValueEarlyOT.Index).Value = Me.ObjCd.Convert_ToString(Me.MaxValueEarlyOTWS)
                    .Cells(Me.ColCalculateEarlyOTTolerance.Index).Value = Me.CalculateEarlyOTToleranceWS
                End If

                ''---
                If Me.IsOverrideLateRule = True Then
                    If finalCalculateLate = True Then
                        .Cells(Me.ColIsCalculateLate.Index).Value = True
                    Else
                        .Cells(Me.ColIsCalculateLate.Index).Value = False
                    End If
                    .Cells(Me.ColCalculateAfterEarlyLate.Index).Value = Me.ObjCd.Convert_ToString(Me.CalculateAfterLateM)
                    .Cells(Me.ColRoundingTypeLate.Index).Value = CboDgvLate.Items.Item(Me.RoundingTypeLateM)
                    .Cells(Me.ColRoundingValueLate.Index).Value = Me.ObjCd.Convert_ToString(Me.RoundingValueLateM)
                    .Cells(Me.ColMaxValueLate.Index).Value = Me.ObjCd.Convert_ToString(Me.MaxValueLateM)
                    .Cells(Me.ColCalculateLateTolerance.Index).Value = Me.CalculateLateTolerance
                Else
                    If finalCalculateLateWS = True Then
                        .Cells(Me.ColIsCalculateLate.Index).Value = True
                    Else
                        .Cells(Me.ColIsCalculateLate.Index).Value = False
                    End If
                    .Cells(Me.ColCalculateAfterEarlyLate.Index).Value = Me.ObjCd.Convert_ToString(Me.CalculateAfterLateWS)
                    .Cells(Me.ColRoundingTypeLate.Index).Value = CboDgvLate.Items.Item(Me.RoundingTypeLateWS)
                    .Cells(Me.ColRoundingValueLate.Index).Value = Me.ObjCd.Convert_ToString(Me.RoundingValueLateWS)
                    .Cells(Me.ColMaxValueLate.Index).Value = Me.ObjCd.Convert_ToString(Me.MaxValueLateWS)
                    .Cells(Me.ColCalculateLateTolerance.Index).Value = Me.CalculateLateToleranceWS
                End If

                ''--
                If Me.IsOverrideEarlyOutRule = True Then
                    If finalCalculateEarlyOut = True Then
                        .Cells(Me.ColIsCalculateEarlyOut.Index).Value = True
                    Else
                        .Cells(Me.ColIsCalculateEarlyOut.Index).Value = False
                    End If
                    .Cells(Me.ColCalculateAfterEarlyOut.Index).Value = Me.ObjCd.Convert_ToString(Me.CalculateAfterEarlyOutM)
                    .Cells(Me.ColRoundingTypeEarlyOut.Index).Value = CboDgvEarlyOut.Items.Item(Me.RoundingTypeEarlyOutM)
                    .Cells(Me.ColRoundingValueEarlyOut.Index).Value = Me.ObjCd.Convert_ToString(Me.RoundingValueEarlyOutM)
                    .Cells(Me.ColMaxValueEarlyOut.Index).Value = Me.ObjCd.Convert_ToString(Me.MaxValueEarlyOutM)
                    .Cells(Me.ColCalculateEarlyOutTolerance.Index).Value = Me.CalculateEarlyOutTolerance
                Else
                    If finalCalculateEarlyOutWS = True Then
                        .Cells(Me.ColIsCalculateEarlyOut.Index).Value = True
                    Else
                        .Cells(Me.ColIsCalculateEarlyOut.Index).Value = False
                    End If
                    .Cells(Me.ColCalculateAfterEarlyOut.Index).Value = Me.ObjCd.Convert_ToString(Me.CalculateAfterEarlyOutWS)
                    .Cells(Me.ColRoundingTypeEarlyOut.Index).Value = CboDgvEarlyOut.Items.Item(Me.RoundingTypeEarlyOutWS)
                    .Cells(Me.ColRoundingValueEarlyOut.Index).Value = Me.ObjCd.Convert_ToString(Me.RoundingValueEarlyOutWS)
                    .Cells(Me.ColMaxValueEarlyOut.Index).Value = Me.ObjCd.Convert_ToString(Me.MaxValueEarlyOutWS)
                    .Cells(Me.ColCalculateEarlyOutTolerance.Index).Value = Me.CalculateEarlyOutToleranceWS
                End If

                ''--
                If Me.IsOverrideOTRule = True Then
                    If finalCalculateOT = True Then
                        .Cells(Me.ColIsCalculateOT.Index).Value = True
                    Else
                        .Cells(Me.ColIsCalculateOT.Index).Value = False
                    End If
                    .Cells(Me.ColCalculateAfterOT.Index).Value = Me.ObjCd.Convert_ToString(Me.CalculateAfterOTM)
                    .Cells(Me.ColRoundingTypeOT.Index).Value = CboDgvOT.Items.Item(Me.RoundingTypeOTM)
                    .Cells(Me.ColRoundingValueOT.Index).Value = Me.ObjCd.Convert_ToString(Me.RoundingValueOTM)
                    .Cells(Me.ColMaxValueOT.Index).Value = Me.ObjCd.Convert_ToString(Me.MaxValueOTM)
                    .Cells(Me.ColCalculateOTTolerance.Index).Value = Me.CalculateOTTolerance
                Else
                    If finalCalculateOTWS = True Then
                        .Cells(Me.ColIsCalculateOT.Index).Value = True
                    Else
                        .Cells(Me.ColIsCalculateOT.Index).Value = False
                    End If
                    .Cells(Me.ColCalculateAfterOT.Index).Value = Me.ObjCd.Convert_ToString(Me.CalculateAfterOTWS)
                    .Cells(Me.ColRoundingTypeOT.Index).Value = CboDgvOT.Items.Item(Me.RoundingTypeOTWS)
                    .Cells(Me.ColRoundingValueOT.Index).Value = Me.ObjCd.Convert_ToString(Me.RoundingValueOTWS)
                    .Cells(Me.ColMaxValueOT.Index).Value = Me.ObjCd.Convert_ToString(Me.MaxValueOTWS)
                    .Cells(Me.ColCalculateOTTolerance.Index).Value = Me.CalculateOTToleranceWS
                End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPrimaryEmployee_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPrimaryEmployee.TextChanged
        Try
            If IsNumeric(Me.TxtPrimaryEmployee.Text) AndAlso Me.TxtRemarks.Enabled = True Then
                Me.dgvItem.Rows.Clear()
                'aul tambahan
                Me.ApakahSudahReload = False
                Me.JoinWithAttPrev = False
                Me.DtpStartAttPrev = "1/1/1950"
                Me.DtpToAttPrev = "1/1/1950"
                Me.BtnChange.Enabled = False
                If Me.btnOK.Enabled = True Then
                    Me.dtpStart.Enabled = True
                    Me.BtnReload.Enabled = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindEmployeeCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnFindEmployeeCode.GotFocus
        Try
            If IsNumeric(Me.TxtPrimaryEmployee.Text) Then
                Me.BufferEmployee = CDbl(Me.TxtPrimaryEmployee.Text)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'aul tambahan
    Public Function GetPrevAttendance() As Boolean
        Dim PrimaryPrevAtt As Double = -1
        Dim DtpStartBuff As Date = Me.dtpStart.Value.Date
        Dim DtpEndBuff As Date = Me.DtpEnd.Value.Date
        Dim DtpDumm As Date = #1/1/1950#

        Try
            If Me.JoinWithAttPrev = False Then
                DtpDumm = Me.dtpStart.Value.Date
            Else
                DtpDumm = Me.DtpToAttPrev.Date + TimeSpan.FromDays(1)
            End If

            'pengecualian jika di panggil dari btn reload
            PrimaryPrevAtt = Me.ObjAtt.GetPrimaryPrevAttendance(Me.TxtPrimaryEmployee.Text, DtpDumm.Date)
            If PrimaryPrevAtt <> -1 Then
                If Me.TxtPrimary.Text <> -1 Then
                    If Me.TxtPrimary.Text <> PrimaryPrevAtt Then
                        Return False
                    Else
                        If Me.JoinWithAttPrev = False Then
                            Return False
                        End If
                    End If
                End If

                Me.ObjAtt.CariDetail(PrimaryPrevAtt, Me, Me.TblOT_Order, Me.JoinWithAttPrev)
                If Me.ObjAtt.ApakahBolehEdit(PrimaryPrevAtt, Me.UserCode, False, False) = False Then
                    Me.ObjAtt.KosongkanObject()
                    Me.dgvItem.Rows.Clear()
                    Return False
                Else
                    Me.Baru = False
                    Me.SalinKeField()

                    Me.DtpStartAttPrev = Me.dtpStart.Value.Date
                    Me.DtpToAttPrev = Me.DtpEnd.Value.Date
                    Me.dtpStart.Value = DtpStartBuff
                    Me.DtpEnd.Value = DtpEndBuff
                    Return True
                End If
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function
End Class