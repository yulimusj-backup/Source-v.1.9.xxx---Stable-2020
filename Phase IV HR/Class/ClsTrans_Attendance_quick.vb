Option Explicit On
Option Strict Off

Imports Core_Module, System.Drawing, System.Windows.Forms, System.Threading
Public Class ClsTrans_Attendance_quick

    Dim mFrm As FrmTrans_Attendance_Quick2
    Public listEmp As New List(Of Double)

    Public Sub ApakahBolehHapus_Quick(ByVal ListPrimaryAttn As List(Of Double), ByVal TBLMsgBox As DseTrans_Attendance.TBLMessageBoxDataTable, _
    ByVal vShowMsg As Boolean, ByVal vUser As String, ByVal vFrm As FrmTrans_Attendance_DataImport)
        Dim dta As New DseTrans_AttendanceTableAdapters.tbltrans_attendanceTableAdapter
        Dim DtaSalary As New DseTrans_AttendanceTableAdapters.DataTable3TableAdapter
        Dim TblSalary As New DseTrans_Attendance.DataTable3DataTable
        Dim RowMsgBox As DseTrans_Attendance.TBLMessageBoxRow
        Dim RowDGV As DataGridViewRow
        Dim DummyUser As String = ""
        Dim ListString As New List(Of String)
        Dim ObjEmp As New Core_Module.ClsMasterKaryawan
        Dim EmpCode As String = ""
        Dim EmpFullName As String = ""
        Dim PrimaryEmp As Double = -1
        Dim Ketemu As Boolean = False
        Dim DummyNoTrans As String = ""


        Try
            'tambahan dari leave app setting
            dta.Connection.ConnectionString = Core_Module.GetConstring
            DtaSalary.Connection.ConnectionString = Core_Module.GetConstring

            For Each RowDGV In vFrm.DGVAttendanceData.Rows
                Ketemu = False
                If RowDGV.Cells(vFrm.ColChkDelete.Index).Value = True Then
                    EmpCode = RowDGV.Cells(vFrm.ColEmployeeCode.Index).Value
                    EmpFullName = RowDGV.Cells(vFrm.ColEmployeeName.Index).Value
                    PrimaryEmp = ObjEmp.GetPrimaryByCode(EmpCode)
                    '1. strict edit
                    DummyUser = dta.ScalarQuery_GetCreatedBy(RowDGV.Cells(vFrm.ColPrimaryTransaksi.Index).Value)
                    If DummyUser <> vUser Then
                        If vShowMsg = True Then
                            If ListString.Contains(RowDGV.Cells(vFrm.ColPrimaryTransaksi.Index).Value.ToString + "-" + "Only " + DummyUser + " user is allowed to delete this transaction.") = False Then
                                ListString.Add(RowDGV.Cells(vFrm.ColPrimaryTransaksi.Index).Value.ToString + "-" + "Only " + DummyUser + " user is allowed to delete this transaction.")
                                RowMsgBox = TBLMsgBox.NewRow
                                RowMsgBox.Employee_Code = EmpCode
                                RowMsgBox.Employee_Name = EmpFullName
                                RowMsgBox.Reason = "Only user : " + DummyUser + " is allowed to delete this transaction."
                                RowMsgBox.Document_No = RowDGV.Cells(vFrm.ColDocumentNo.Index).Value.ToString + "(" + RowDGV.Cells(vFrm.ColStartDate.Index).Value + " - " + RowDGV.Cells(vFrm.ColEndDate.Index).Value + ")"
                                TBLMsgBox.Rows.Add(RowMsgBox)

                                RowDGV.Cells(vFrm.ColChkDelete.Index).Value = Boolean.FalseString
                            End If
                            'MsgBox("Only " + DummyUser + " user is allowed to delete this transaction.", _
                            'MsgBoxStyle.Information, "Information")
                        End If
                        Ketemu = True
                    End If
                    If Ketemu = False Then
                        '2. cek disalary
                        DtaSalary.FillBy_PrimayEmpNDate(TblSalary, PrimaryEmp, RowDGV.Cells(vFrm.ColEndDate.Index).Value, RowDGV.Cells(vFrm.ColStartDate.Index).Value)
                        If TblSalary.Rows.Count > 0 Then
                            If vShowMsg = True Then
                                DummyNoTrans = ""
                                For Each rowSalary As DseTrans_Attendance.DataTable3Row In TblSalary.Rows
                                    If DummyNoTrans = "" Then
                                        DummyNoTrans = rowSalary.No_Transaksi
                                    Else
                                        DummyNoTrans = DummyNoTrans + ", " + rowSalary.No_Transaksi
                                    End If
                                Next
                                If ListString.Contains(RowDGV.Cells(vFrm.ColPrimaryTransaksi.Index).Value.ToString + "-" + "Refered by salary calculation: " + _
                                DummyNoTrans) = False Then
                                    ListString.Add(RowDGV.Cells(vFrm.ColPrimaryTransaksi.Index).Value.ToString + "-" + "Refered by salary calculation: " + _
                                    DummyNoTrans)

                                    RowMsgBox = TBLMsgBox.NewRow
                                    RowMsgBox.Employee_Code = EmpCode
                                    RowMsgBox.Employee_Name = EmpFullName
                                    RowMsgBox.Reason = "Refered by salary calculation: " + DummyNoTrans
                                    RowMsgBox.Document_No = RowDGV.Cells(vFrm.ColDocumentNo.Index).Value.ToString + "(" + RowDGV.Cells(vFrm.ColStartDate.Index).Value + " - " + RowDGV.Cells(vFrm.ColEndDate.Index).Value + ")"

                                    TBLMsgBox.Rows.Add(RowMsgBox)

                                    RowDGV.Cells(vFrm.ColChkDelete.Index).Value = Boolean.FalseString
                                End If
                                'MsgBox("Data attendance cannot be deleted because this data is refering to calculate salary.", _
                                'MsgBoxStyle.Information, "information")
                            End If
                            Ketemu = True
                        End If
                    End If
                    '3 Masukan ke dalam List
                    If Ketemu = False Then
                        If ListPrimaryAttn.Contains(RowDGV.Cells(vFrm.ColPrimaryTransaksi.Index).Value) = False Then
                            ListPrimaryAttn.Add(RowDGV.Cells(vFrm.ColPrimaryTransaksi.Index).Value)
                        End If
                    End If
                    '------------------
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            DtaSalary.Dispose()
            TblSalary.Dispose()
            ListString = Nothing
            ObjEmp = Nothing
        End Try
    End Sub

    Public Sub Show(ByVal vStrDate As Date, ByVal vEndDate As Date, ByVal vFrm As FrmTrans_Attendance_DataImport)
        Dim Dta As New DseTrans_AttendanceTableAdapters.DataTable7TableAdapter
        Dim tbl As New DseTrans_Attendance.DataTable7DataTable
        Dim row As DseTrans_Attendance.DataTable7Row
        Dim rowDgv As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            vFrm.Cursor = Cursors.WaitCursor
            vFrm.DGVAttendanceData.Rows.Clear()
            vFrm.DGVAttendanceData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vFrm.DGVAttendanceData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If vFrm.CboTypeDateFilter.SelectedIndex = 0 Then
                'transaction date
                Dta.FillBy_Date(tbl, vStrDate.Date, vEndDate.Date)
            Else
                'period
                Dta.FillBy_Period(tbl, vStrDate.Date, vEndDate.Date)
            End If

            vFrm.ColDate.DefaultCellStyle.Format = "dd MMM yyyy"

            Dim dv As DataRowView
            Dim ObjSort As New ClsTransaksiQuick_Sorting
            Dim ObjUser As New Core_Module.ClsUser
            Dim PrimaryUser As Double = ObjUser.GetPrimary(vFrm.UserCode)
            tbl.DefaultView.RowFilter = ""
            tbl.DefaultView.Sort = ObjSort.GetSetting(PrimaryUser, ClsTransaksiQuick_Sorting.EnumTipeTransaksi.QuickAttendance)

            For Each dv In tbl.DefaultView
                row = dv.Row
                vFrm.DGVAttendanceData.Rows.Add()
                rowDgv = vFrm.DGVAttendanceData.Rows(vFrm.DGVAttendanceData.Rows.Count - 1)
                rowDgv.Cells(vFrm.ColNo.Index).Value = rowDgv.Index + 1
                rowDgv.Cells(vFrm.ColEmployeeCode.Index).Value = row.Emp_Code
                rowDgv.Cells(vFrm.ColEmployeeName.Index).Value = row.Emp_Name
                rowDgv.Cells(vFrm.ColDept.Index).Value = row.Dept_Desc
                rowDgv.Cells(vFrm.ColJobPosition.Index).Value = row.Job_Desc
                rowDgv.Cells(vFrm.ColStartDate.Index).Value = row.start_date 'Format(row.start_date, "dd MMM yyyy")
                rowDgv.Cells(vFrm.ColEndDate.Index).Value = row.end_date 'Format(row.end_date, "dd MMM yyyy")
                rowDgv.Cells(vFrm.ColPrimaryTransaksi.Index).Value = row.primary
                rowDgv.Cells(vFrm.ColDocumentNo.Index).Value = row.no_transaksi
                rowDgv.Cells(vFrm.ColDate.Index).Value = row.tanggal

                rowDgv.Cells(vFrm.ColEmployeeNIK.Index).Value = row.Emp_NIK
                rowDgv.Cells(vFrm.ColDeptCode.Index).Value = row.Dept_Code
                rowDgv.Cells(vFrm.ColJobPositionCode.Index).Value = row.Job_Code
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            vFrm.Cursor = Cursors.Default
            vFrm.DGVAttendanceData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vFrm.DGVAttendanceData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        End Try
    End Sub

    Private Function CekBatasan_dtp_DgvHarusSamaWaktunya(ByVal vPrimaryEmp As Double, _
    ByVal DtpStart As Date, ByVal DtpEnd As Date) As Boolean
        Dim Dta As New DseTrans_AttendanceTableAdapters.DataTable2TableAdapter
        Dim Tbl As New DseTrans_Attendance.DataTable2DataTable
        Dim Row As DseTrans_Attendance.DataTable2Row
        Dim TglAwal As Date = #1/1/1950#
        Dim TglAkhir As Date = #1/1/1950#
        Dim ListDateAllocate As New List(Of Date)

        Try
            Dta.Fill(Tbl, vPrimaryEmp, DtpStart, DtpEnd)
            If Tbl.Rows.Count > 0 Then
                Tbl.DefaultView.RowFilter = ""
                Tbl.DefaultView.Sort = "[start date] ASC"
                For Each dv As DataRowView In Tbl.DefaultView
                    Row = dv.Row
                    TglAwal = Row.start_date
                    While TglAwal <= Row.end_date
                        If ListDateAllocate.Contains(TglAwal) = False Then
                            ListDateAllocate.Add(TglAwal)
                        End If
                        TglAwal = TglAwal.AddDays(1)
                    End While
                Next

                TglAwal = DtpStart
                While TglAwal <= DtpEnd
                    If ListDateAllocate.Contains(TglAwal) = False Then
                        Return False
                    End If
                    TglAwal = TglAwal.AddDays(1)
                End While

                ''Batasan Tanggal Awal
                'Row = Tbl.Rows(0)
                ''cari waktu yang akan dimasukin ke dgv nya
                'If Row.start_date < DtpStart And Row.end_date > DtpEnd Then
                '    TglAwal = DtpStart
                'ElseIf Row.start_date < DtpStart And Row.end_date <= DtpEnd Then
                '    TglAwal = DtpStart
                'ElseIf Row.start_date >= DtpStart And Row.end_date > DtpEnd Then
                '    TglAwal = Row.start_date
                'ElseIf Row.start_date >= DtpStart And Row.end_date <= DtpEnd Then
                '    TglAwal = Row.start_date
                'End If
                ''Batasan Tanggal Akhir
                'Row = Tbl.Rows(Tbl.Rows.Count - 1)
                ''cari waktu yang akan dimasukin ke dgv nya
                'If Row.start_date < DtpStart And Row.end_date > DtpEnd Then
                '    TglAkhir = DtpEnd
                'ElseIf Row.start_date < DtpStart And Row.end_date <= DtpEnd Then
                '    TglAkhir = Row.end_date
                'ElseIf Row.start_date >= DtpStart And Row.end_date > DtpEnd Then
                '    TglAkhir = DtpEnd
                'ElseIf Row.start_date >= DtpStart And Row.end_date <= DtpEnd Then
                '    TglAkhir = Row.end_date
                'End If
            Else
                Return False
            End If

            ''dtp dan dgv harus sama sama waktunya
            'If CDate(TglAwal) <> DtpStart Or _
            'CDate(TglAkhir) <> DtpEnd Then
            '    Return False
            'End If

            'Sampai sini benar
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
            Tbl.Dispose()
        End Try
    End Function

    Private Function CekBatasanButtonReload(ByVal RowEmp As DseTrans_Attendance.DataTable9Row, ByVal vFrm As FrmTrans_Attendance_Quick2, _
    ByVal TBLMsgbox As DseTrans_Attendance.TBLMessageBoxDataTable, ByVal ListString As List(Of String), _
    Optional ByVal vShowMsgBox As Boolean = False) As Boolean
        Dim RowMsgbox As DseTrans_Attendance.TBLMessageBoxRow
        Dim ObjEmp As New Core_Module.ClsMasterKaryawan
        Dim ObjAtt As New ClsTrans_Attendance
        Dim EmpCode As String = ""

        Try
            'jika employee tidak aktif, maka tidak perlu dibuatkan attendancenya
            If vFrm.ObjEmp.ApakahAktif(RowEmp.Emp_Primary) = False Then
                If vShowMsgBox = True Then
                    If ListString.Contains(RowEmp.Emp_Code + "-" + "Inactive employee") = False Then
                        ListString.Add(RowEmp.Emp_Code + "-" + "Inactive employee")
                        RowMsgbox = TBLMsgbox.NewRow
                        RowMsgbox.Employee_Code = RowEmp.Emp_Code
                        RowMsgbox.Employee_Name = RowEmp.Emp_Name
                        RowMsgbox.Reason = "Inactive employee."
                        TBLMsgbox.Rows.Add(RowMsgbox)
                    End If
                End If
                Return False
            End If

            EmpCode = RowEmp.Emp_Code
            If EmpCode = "" Then
                If vShowMsgBox = True Then
                    If ListString.Contains(RowEmp.Emp_Code + "-" + "Employee code is empty.") = False Then
                        ListString.Add(RowEmp.Emp_Code + "-" + "Employee code is empty.")
                        RowMsgbox = TBLMsgbox.NewRow
                        RowMsgbox.Employee_Code = RowEmp.Emp_Code
                        RowMsgbox.Employee_Name = RowEmp.Emp_Name
                        RowMsgbox.Reason = "Employee code is empty."
                        TBLMsgbox.Rows.Add(RowMsgbox)
                    End If
                End If
                Return False
            End If
            If vFrm.dtpPeriod.Value.Date > vFrm.DtpTo.Value.Date Then
                If vShowMsgBox = True Then
                    If ListString.Contains(RowEmp.Emp_Code + "-" + "Period from cannot be more than period to.") = False Then
                        ListString.Add(RowEmp.Emp_Code + "-" + "Period from cannot be more than period to.")
                        RowMsgbox = TBLMsgbox.NewRow
                        RowMsgbox.Employee_Code = RowEmp.Emp_Code
                        RowMsgbox.Employee_Name = RowEmp.Emp_Name
                        RowMsgbox.Reason = "Period from cannot be more than period to."
                        TBLMsgbox.Rows.Add(RowMsgbox)
                    End If
                End If
                Return False
            Else
                If ObjAtt.CekValidPeriodEmp(-1, RowEmp.Emp_Primary, vFrm.dtpPeriod.Value.Date, vFrm.DtpTo.Value.Date, True, False) = False Then
                    If vShowMsgBox = True Then
                        If ListString.Contains(RowEmp.Emp_Code + "-" + "Conflicting with existing attendance data") = False Then
                            ListString.Add(RowEmp.Emp_Code + "-" + "Conflicting with existing attendance data")
                            RowMsgbox = TBLMsgbox.NewRow
                            RowMsgbox.Employee_Code = RowEmp.Emp_Code
                            RowMsgbox.Employee_Name = RowEmp.Emp_Name
                            RowMsgbox.Reason = "Conflicting with existing attendance data."
                            TBLMsgbox.Rows.Add(RowMsgbox)
                        End If
                    End If
                    Return False
                End If
            End If
            If Me.CekBatasan_dtp_DgvHarusSamaWaktunya(RowEmp.Emp_Primary, vFrm.dtpPeriod.Value.Date, vFrm.DtpTo.Value.Date) = False Then
                If vShowMsgBox = True Then
                    If ListString.Contains(RowEmp.Emp_Code + "-" + "Please allocate workshift.") = False Then
                        ListString.Add(RowEmp.Emp_Code + "-" + "Please allocate workshift.")
                        RowMsgbox = TBLMsgbox.NewRow
                        RowMsgbox.Employee_Code = RowEmp.Emp_Code
                        RowMsgbox.Employee_Name = RowEmp.Emp_Name
                        RowMsgbox.Reason = "Please allocate workshift."
                        TBLMsgbox.Rows.Add(RowMsgbox)
                    End If
                End If
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjEmp = Nothing
            ObjAtt = Nothing
        End Try
    End Function

    Private Function BuatAttendance(ByRef KetemuSimpan As Boolean, ByVal vFrm As FrmTrans_Attendance_Quick2, _
    ByVal Index As Integer, ByVal RowEmp As DseTrans_Attendance.DataTable9Row) As Boolean
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub HapusBarisTBLEmp(ByVal vFrm As FrmTrans_Attendance_Quick2, _
    ByVal TblEmp As DseTrans_Attendance.DataTable9DataTable, _
    ByVal TBLMsgbox As DseTrans_Attendance.TBLMessageBoxDataTable)
        Dim RowEmp As DseTrans_Attendance.DataTable9Row
        Dim ij As Integer = 0
        Dim ListString As New List(Of String)

        Try
            While ij <= TblEmp.Rows.Count - 1
                RowEmp = TblEmp.Rows(ij)
                'jika all tiap row masukin ke dalam list
                If vFrm.RboAll.Checked = True Then
                    Me.listEmp.Add(RowEmp.Emp_Primary)
                End If
                If Me.listEmp.Contains(RowEmp.Emp_Primary) = True Then
                    If Me.CekBatasanButtonReload(RowEmp, vFrm, TBLMsgbox, ListString, True) = True Then
                        ij += 1
                    ElseIf Me.CekBatasanButtonReload(RowEmp, vFrm, TBLMsgbox, ListString, True) = False Then
                        TblEmp.Rows.RemoveAt(ij)
                    End If
                Else
                    ij += 1
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListString = Nothing
        End Try
    End Sub

    Public Function Simpan(ByVal vFrm As FrmTrans_Attendance_Quick2, ByVal TblMsgBox As DseTrans_Attendance.TBLMessageBoxDataTable, _
    ByVal TBL_JP As Core_Module.DseMasterWorkshift_Allocation.tblmaster_jobpositionDataTable, _
    ByVal TBL_Dept As Core_Module.DseMasterWorkshift_Allocation.tblmaster_departemenDataTable, _
    ByVal vIsImportFromMachine As Boolean, ByVal vIsContinueFromPrevAtt As Boolean) As Boolean
        Dim DtaEmp As New DseTrans_AttendanceTableAdapters.DataTable9TableAdapter
        Dim TblEmp As New DseTrans_Attendance.DataTable9DataTable
        Dim RowEmp As DseTrans_Attendance.DataTable9Row
        Dim KetemuSimpan As Boolean = False
        Dim StopProcess As Boolean = False
        Dim FrmError As New FrmTrans_Gaji_MsgBox
        Dim Number As Integer = 0
        Dim ListString As New List(Of String)
        Dim hasil As MsgBoxResult
        Dim WS As Windows.Forms.FormWindowState = FormWindowState.Normal
        Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser

        Try
            Me.mFrm = vFrm
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DtaEmp.Fill(TblEmp)
            Me.listEmp.Clear()
            If vFrm.RboSelected.Checked = True Then
                'Me.FilterSelectedData(vFrm, TblEmp)
                Me.FilterDataKaryawanNew(vFrm, TblEmp, TBL_JP, TBL_Dept)
            End If
            Me.HapusBarisTBLEmp(vFrm, TblEmp, TblMsgBox)
            If TblMsgBox.Rows.Count > 0 Then
                TblMsgBox.DefaultView.Sort = "[Employee Code], [Employee Name]"
                FrmError.dgvItem.DataSource = TblMsgBox.DefaultView
                FrmError.dgvItem.Columns(0).Visible = False
                FrmError.LblProgress.Text = "Total Rows = " + TblMsgBox.Rows.Count.ToString
                FrmError.Text = "Can't create attendance data for the following employee"
                If FrmError.dgvItem.ColumnCount > 0 Then
                    FrmError.dgvItem.Columns(1).Visible = False
                End If

                FrmError.ShowDialog()
            End If

            hasil = MsgBox("Input attendance data?", MsgBoxStyle.YesNo, "Confirm")
            If hasil = MsgBoxResult.Yes Then
                ObjSettingUser.SetSettingValue(ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_MergeAttendance, vIsContinueFromPrevAtt, vFrm.User)
                'aul tambahan - cek apakah jika ada penggabungan attendance melebihi batas attendance dalam sebulan
                If TblEmp.Count > 0 And vIsContinueFromPrevAtt = True Then
                    If Me.ApakahMelebihiPeriod(vFrm, TblEmp, TblMsgBox) = True Then
                        hasil = MsgBox("Cancel merge attendance data?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            vIsContinueFromPrevAtt = False
                        End If
                    End If
                End If

                For i As Integer = 0 To TblEmp.Rows.Count - 1
                    RowEmp = TblEmp.Rows(i)
                    'jika all tiap row masukin ke dalam list
                    If vFrm.RboAll.Checked = True Then
                        Me.listEmp.Add(RowEmp.Emp_Primary)
                    End If
                    If Me.listEmp.Contains(RowEmp.Emp_Primary) = True Then
                        If Me.CekBatasanButtonReload(RowEmp, vFrm, TblMsgBox, ListString, False) = True Then
                            KetemuSimpan = True
                            Dim FrmAtt As New FrmTrans_Attendance()
                            'Koding Progress Bar
                            vFrm.LblProgress.Text = "Data " + (i + 1).ToString + " of " + TblEmp.Rows.Count.ToString + " - " + RowEmp.Emp_Name
                            vFrm.LblProgress.Refresh()
                            vFrm.ProgressBar1.Visible = True
                            vFrm.ProgressBar1.Maximum = TblEmp.Rows.Count
                            vFrm.ProgressBar1.Value = i + 1
                            'Progress bar form attendance
                            FrmAtt.BtnStop.Visible = True
                            FrmAtt.lblProgress.Visible = True
                            FrmAtt.ProgressBar1.Visible = True
                            FrmAtt.ProgressBar1.Maximum = TblEmp.Rows.Count
                            FrmAtt.ProgressBar1.Value = i + 1
                            FrmAtt.lblProgress.Text = "Data " + (i + 1).ToString + " of " + TblEmp.Rows.Count.ToString
                            '------------------
                            FrmAtt.BackColorAll = vFrm.BackColorAll
                            FrmAtt.IsImportFromMachine = vIsImportFromMachine
                            FrmAtt.IsContinueFromPrevAtt = vIsContinueFromPrevAtt   'aul tambahan
                            FrmAtt.DariQuick = True
                            FrmAtt.Baru = True
                            FrmAtt.AktifkanDetail()
                            FrmAtt.DtpDate.Value = vFrm.DtpTanggalTransaksi.Value
                            FrmAtt.TxtDocNumber.Text = FrmAtt.ObjAtt.GetNoBaru(FrmAtt.DtpDate.Value)
                            FrmAtt.tslCreatedBy.Text = "Created By: " + vFrm.User
                            FrmAtt.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
                            FrmAtt.tslModifiedBy.Text = "Modified By: " + vFrm.User
                            FrmAtt.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
                            FrmAtt.TxtPrimaryEmployee.Text = RowEmp.Emp_Primary
                            FrmAtt.TxtEmployeeName.Text = RowEmp.Emp_Name
                            FrmAtt.TxtEmployeeCode.Text = RowEmp.Emp_Code
                            FrmAtt.TxtPrimaryPosition.Text = RowEmp.Job_Primary
                            FrmAtt.TxtPosition.Text = RowEmp.Job_Desc
                            FrmAtt.dtpStart.Value = vFrm.dtpPeriod.Value
                            FrmAtt.DtpEnd.Value = vFrm.DtpTo.Value
                            FrmAtt.UserCode = vFrm.User
                            FrmAtt.WindowState = WS
                            'FrmAtt.SetKolomPermit()
                            If FrmAtt.ShowDialog() = Windows.Forms.DialogResult.Abort Then
                                StopProcess = True
                                Exit For
                            End If
                            WS = FrmAtt.WindowState
                            FrmAtt.Dispose()
                        End If
                    End If
                Next
                If KetemuSimpan = False Then
                    MsgBox("No employee data can be processed, please cek allocate workshift & attendance data.", MsgBoxStyle.Information, "Information")
                End If
                If StopProcess = True Then
                    MsgBox("Process stopped by user.", MsgBoxStyle.Information, "Information")
                End If
            End If
            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaEmp.Dispose()
            TblEmp.Dispose()
            FrmError.Dispose()
            ListString = Nothing
            ObjSettingUser = Nothing
        End Try
    End Function

#Region "FilterDataKaryawan"
    '-----cara koding sama dengan quick salary---------
    Private Sub FilterDataKaryawanNew(ByVal vFrm As FrmTrans_Attendance_Quick2, _
    ByVal TblEmp As DseTrans_Attendance.DataTable9DataTable, _
    ByVal TBL_JP As Core_Module.DseMasterWorkshift_Allocation.tblmaster_jobpositionDataTable, _
    ByVal TBL_Dept As Core_Module.DseMasterWorkshift_Allocation.tblmaster_departemenDataTable)
        Dim PrimaryUser As Double = -1
        Dim ObjUser As New Core_Module.ClsUser
        Dim mAllData As Boolean = True
        Dim mListDataFilter As New List(Of Double)
        Try
            If vFrm.RboSelected.Checked = True Then
                If vFrm.ChkDepartment.Checked = True Then
                    mAllData = False
                ElseIf vFrm.ChkEmployee.Checked = True Then
                    mAllData = False
                ElseIf vFrm.chkSalaryFormula.Checked = True Then
                    mAllData = False
                ElseIf vFrm.chkJobPosition.Checked = True Then
                    mAllData = False
                End If
            End If
            'Fill Data Karyawan dan sesuaikan dengan kriteria Selection Data
            PrimaryUser = ObjUser.GetPrimary(vFrm.User)
            If mAllData = True Then
            Else
                mListDataFilter.Clear()
                If vFrm.ChkDepartment.Checked = True Then
                    If mListDataFilter.Contains(FrmMasterWorkShift_Allocate.EnumLaporan.Departemen) = False Then
                        mListDataFilter.Add(FrmMasterWorkShift_Allocate.EnumLaporan.Departemen)
                    End If
                End If
                If vFrm.ChkEmployee.Checked = True Then
                    If mListDataFilter.Contains(FrmMasterWorkShift_Allocate.EnumLaporan.Employee) = False Then
                        mListDataFilter.Add(FrmMasterWorkShift_Allocate.EnumLaporan.Employee)
                    End If
                End If
                If vFrm.chkSalaryFormula.Checked = True Then
                    If mListDataFilter.Contains(FrmMasterWorkShift_Allocate.EnumLaporan.SalaryFormula) = False Then
                        mListDataFilter.Add(FrmMasterWorkShift_Allocate.EnumLaporan.SalaryFormula)
                    End If
                End If
                If vFrm.chkJobPosition.Checked = True Then
                    If mListDataFilter.Contains(FrmMasterWorkShift_Allocate.EnumLaporan.JobPosition) = False Then
                        mListDataFilter.Add(FrmMasterWorkShift_Allocate.EnumLaporan.JobPosition)
                    End If
                End If
                vFrm.dgvEmployee.AllowUserToAddRows = False
                Me.KelolaDataKaryawan(PrimaryUser, TblEmp, mListDataFilter, vFrm.dgvEmployee, TBL_JP, _
                vFrm.ClbPosition, TBL_Dept, vFrm.ClbDepartement, vFrm)
                vFrm.dgvEmployee.AllowUserToAddRows = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjUser = Nothing
            mListDataFilter = Nothing
        End Try
    End Sub

    Private Sub KelolaDataKaryawan(ByVal vPrimaryUser As Double, ByVal TBLKaryawan As DseTrans_Attendance.DataTable9DataTable, _
    ByVal ListDataFilter As List(Of Double), ByVal DGVEmployee As DataGridView, ByVal TBL_JP As Core_Module.DseMasterWorkshift_Allocation.tblmaster_jobpositionDataTable, _
    ByVal Clb_JP As CheckedListBox, ByVal TBL_Dept As Core_Module.DseMasterWorkshift_Allocation.tblmaster_departemenDataTable, _
    ByVal Clb_dept As CheckedListBox, ByVal vFrm As FrmTrans_Attendance_Quick2)
        Dim RowKaryawan As DseTrans_Attendance.DataTable9Row
        Dim EnumTipeDataFilter As Double
        Dim ij As Integer = 0
        Dim Ketemu As Boolean = False
        Dim Primary As Double = -1
        Dim RV As Data.DataRowView
        Dim RowDGV As DataGridViewRow
        Dim Filter As String = ""
        Try
            If ListDataFilter.Count > 0 Then
                Me.listEmp.Clear()
                For Each EnumTipeDataFilter In ListDataFilter
                    If EnumTipeDataFilter = Core_Module.FrmMasterWorkShift_Allocate.EnumLaporan.Departemen Then
                        ij = 0
                        While ij <= TBLKaryawan.Rows.Count - 1
                            RowKaryawan = TBLKaryawan.Rows(ij)
                            Ketemu = False
                            If Clb_dept.CheckedItems.Count > 0 Then
                                For Each RV In Clb_dept.CheckedItems
                                    Primary = RV.Item("Primary")
                                    If RowKaryawan.Dept_Primary = Primary Then
                                        Ketemu = True
                                        Exit For
                                    End If
                                Next
                                If Ketemu = False Then
                                    TBLKaryawan.Rows.RemoveAt(ij)
                                Else
                                    Me.InputListEmp(Me.listEmp, RowKaryawan.Emp_Primary)
                                    ij += 1
                                End If
                            Else
                                ij += 1
                            End If
                        End While
                    End If
                    If EnumTipeDataFilter = Core_Module.FrmMasterWorkShift_Allocate.EnumLaporan.Employee Then
                        ij = 0
                        DGVEmployee.AllowUserToAddRows = False
                        DGVEmployee.AllowUserToDeleteRows = False
                        While ij <= TBLKaryawan.Rows.Count - 1
                            RowKaryawan = TBLKaryawan.Rows(ij)
                            Ketemu = False
                            If DGVEmployee.Rows.Count > 0 Then
                                For Each RowDGV In DGVEmployee.Rows
                                    Primary = RowDGV.Cells(vFrm.ColPrimaryCode.Index).Value
                                    If RowKaryawan.Emp_Primary = Primary Then
                                        Ketemu = True
                                        Exit For
                                    End If
                                Next
                            End If
                            If Ketemu = False Then
                                TBLKaryawan.Rows.RemoveAt(ij)
                            Else
                                Me.InputListEmp(Me.listEmp, RowKaryawan.Emp_Primary)
                                ij += 1
                            End If
                        End While
                        DGVEmployee.AllowUserToAddRows = True
                        DGVEmployee.AllowUserToDeleteRows = True
                    End If
                    If EnumTipeDataFilter = Core_Module.FrmMasterWorkShift_Allocate.EnumLaporan.JobPosition Then
                        ij = 0
                        While ij <= TBLKaryawan.Rows.Count - 1
                            RowKaryawan = TBLKaryawan.Rows(ij)
                            Ketemu = False
                            If Clb_JP.CheckedItems.Count > 0 Then
                                For Each RV In Clb_JP.CheckedItems
                                    Primary = RV.Item("primary")
                                    If RowKaryawan.Job_Primary = Primary Then
                                        Ketemu = True
                                        Exit For
                                    End If
                                Next
                                If Ketemu = False Then
                                    TBLKaryawan.Rows.RemoveAt(ij)
                                Else
                                    Me.InputListEmp(Me.listEmp, RowKaryawan.Emp_Primary)
                                    ij += 1
                                End If
                            Else
                                ij += 1
                            End If
                        End While
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub InputListEmp(ByVal listEmp As List(Of Double), ByVal vPrimaryEmp As Double)
        Try
            If listEmp.Contains(vPrimaryEmp) = False Then
                listEmp.Add(vPrimaryEmp)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

    '-----Cara koding yudha pertama------
    Private Sub FilterSelectedData(ByVal vFrm As FrmTrans_Attendance_Quick2, ByVal TblEmp As DseTrans_Attendance.DataTable9DataTable)
        Dim RowDgv As DataGridViewRow
        Dim RV As Data.DataRowView
        Dim ListJob As New List(Of Double)
        Dim listDept As New List(Of Double)
        Dim listSF As New List(Of Double)
        Dim Primary As Double = -1
        Dim ada As Boolean = True
        Dim RowEmp As DseTrans_Attendance.DataTable9Row
        Try
            Me.listEmp.Clear()
            ListJob.Clear()
            listDept.Clear()
            listSF.Clear()
            'cari list job , dept , SF
            'job
            For Each RV In vFrm.ClbPosition.CheckedItems
                Primary = RV.Item("primary")
                ListJob.Add(Primary)
            Next
            'dept
            For Each RV In vFrm.ClbDepartement.CheckedItems
                Primary = RV.Item("Primary")
                listDept.Add(Primary)
            Next
            'SF 
            For Each RV In vFrm.clbSalaryFormula.CheckedItems
                Primary = RV.Item("primary")
                listSF.Add(Primary)
            Next

            'sesuaikan dengan emp yang dipilih
            If vFrm.ChkEmployee.Checked = True Then
                For Each RowDgv In vFrm.dgvEmployee.Rows
                    'cek job
                    ada = True
                    If ListJob.Contains(RowDgv.Cells(vFrm.ColPrimaryJob.Index).Value) = False Then
                        ada = False
                    End If
                    'cek dept
                    If listDept.Contains(RowDgv.Cells(vFrm.ColPrimaryDept.Index).Value) = False Then
                        ada = False
                    End If
                    'cek
                    If ListJob.Contains(RowDgv.Cells(vFrm.ColPrimaryJob.Index).Value) = False Then
                        ada = False
                    End If
                    'jika ada masukin kelist
                    If ada = True Then
                        Me.listEmp.Add(RowDgv.Cells(vFrm.ColPrimaryCode.Index).Value)
                    End If
                Next
            Else
                For Each RowEmp In TblEmp.Rows
                    'cek job
                    ada = True
                    If ListJob.Contains(RowEmp.Job_Primary) = False Then
                        ada = False
                    End If
                    'cek dept
                    If listDept.Contains(RowEmp.Dept_Primary) = False Then
                        ada = False
                    End If
                    'cek
                    If listSF.Contains(RowEmp.SF_Primary) = False Then
                        If RowEmp.SF_Primary <> -1 Then
                            ada = False
                        End If
                    End If
                    'jika ada masukin kelist
                    If ada = True Then
                        Me.listEmp.Add(RowEmp.Emp_Primary)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'aul tambahan
    Public Function ApakahMelebihiPeriod(ByVal vFrm As FrmTrans_Attendance_Quick2, _
    ByVal vTblEmp As DseTrans_Attendance.DataTable9DataTable, _
    ByVal vTblMsgBox As DseTrans_Attendance.TBLMessageBoxDataTable) As Boolean
        Dim Dta As New DseTrans_AttendanceTableAdapters.tbl_lastenddatesalaryTableAdapter
        Dim Tbl As New DseTrans_Attendance.tbl_lastenddatesalaryDataTable
        Dim row As DseTrans_Attendance.tbl_lastenddatesalaryRow

        Dim Dta2 As New DseTrans_AttendanceTableAdapters.tbltrans_attendanceTableAdapter
        Dim Tbl2 As New DseTrans_Attendance.tbltrans_attendanceDataTable
        Dim row2 As DseTrans_Attendance.tbltrans_attendanceRow

        Dim RowEmp As DseTrans_Attendance.DataTable9Row
        Dim rowMsgBox As DseTrans_Attendance.TBLMessageBoxRow
        Dim FrmError As New FrmTrans_Gaji_MsgBox

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta2.Connection.ConnectionString = Core_Module.GetConstring
            Dta.Fill_LastEndDateSalary(Tbl)
            Dta2.FillBy_EndDatePeriod(Tbl2, vFrm.dtpPeriod.Value.Date - TimeSpan.FromDays(1))
            If Tbl.Count > 0 Then
                vTblMsgBox.Clear()
                For i As Integer = 0 To vTblEmp.Rows.Count - 1
                    RowEmp = vTblEmp.Rows(i)
                    'pastikan apakah attendance karyawan akan digabung 
                    Tbl2.DefaultView.RowFilter = "[primary karyawan] = " & RowEmp.Emp_Primary
                    If Tbl2.DefaultView.Count > 0 Then
                        row2 = Tbl2.DefaultView(0).Row
                        'pastikan bila akan digabung, attendace sebelumnya dapat di edit
                        If Me.ApakahBolehEdit_Quick(row2.primary_karyawan, vFrm.User, row2.created_by, row2.start_date, row2.end_date) = True Then
                            Tbl.DefaultView.RowFilter = "[Primary Employee] = " & RowEmp.Emp_Primary
                            If Tbl.DefaultView.Count > 0 Then
                                row = Tbl.DefaultView(0).Row
                                Dim DDD As Date = row.Last_End_Date_Salary.AddMonths(1)
                                'cek apabila attendance digabung periodnya malah melebihi period satu bulan
                                'jika melebihi masukan ke list msgbox
                                If vFrm.DtpTo.Value.Date > DDD.Date And row2.start_date <= DDD.Date Then
                                    rowMsgBox = vTblMsgBox.NewRow
                                    rowMsgBox.Employee_Code = RowEmp.Emp_Code
                                    rowMsgBox.Employee_Name = RowEmp.Emp_Name
                                    rowMsgBox.Reason = "Period attendance " & Format(vFrm.dtpPeriod.Value.Date, "dd MMM yyyy") & " - " & Format(vFrm.DtpTo.Value.Date, "dd MMM yyyy") & " will be merge with attendance data previous payroll period."
                                    vTblMsgBox.Rows.Add(rowMsgBox)
                                End If
                            End If
                        End If
                    End If
                Next
            End If

            If vTblMsgBox.Count = 0 Then
                Return False
            Else
                vTblMsgBox.DefaultView.Sort = "[Employee Code], [Employee Name]"
                FrmError.dgvItem.DataSource = vTblMsgBox.DefaultView
                FrmError.dgvItem.Columns(0).Visible = False
                FrmError.LblProgress.Text = "Total Rows = " + vTblMsgBox.Rows.Count.ToString
                FrmError.Text = "Join attendance data for the following employee will exceed the payroll period"
                If FrmError.dgvItem.ColumnCount > 0 Then
                    FrmError.dgvItem.Columns(1).Visible = False
                End If

                FrmError.ShowDialog()

                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            Dta.Dispose()
            Tbl.Dispose()
            FrmError.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahBolehEdit_Quick(ByVal vPrimaryEmp As Double, ByVal vUserCode As String, _
    ByVal vCreatedBy As String, ByVal vStartDate As Date, ByVal vEndDate As Date) As Boolean
        Dim dta As New DseTrans_AttendanceTableAdapters.tbltrans_attendanceTableAdapter
        Dim DtaSalary As New DseTrans_AttendanceTableAdapters.DataTable3TableAdapter
        Dim TblSalary As New DseTrans_Attendance.DataTable3DataTable
        Dim DummyUser As String

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            DtaSalary.Connection.ConnectionString = Core_Module.GetConstring

            'strict edit
            If vCreatedBy <> vUserCode Then
                Return False
            End If

            'cek disalary
            DtaSalary.FillBy_PrimayEmpNDate(TblSalary, vPrimaryEmp, vEndDate, vStartDate)
            If TblSalary.Rows.Count > 0 Then
                Return False
            End If

            'benar
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            DtaSalary.Dispose()
            TblSalary.Dispose()
        End Try
    End Function
End Class
