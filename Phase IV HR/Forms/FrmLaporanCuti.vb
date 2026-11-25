Option Explicit On
Option Strict Off
Imports System.Windows.Forms

Public Class FrmLaporanCuti
    Public UserCode As String = ""
    Public BackColorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll

    Dim Buffer As String = ""
    Dim ObjProduct As New Core_Module.ClsMasterKaryawan
    Dim EraseRowIndex As Integer = 0
    Dim TBL_JP As New DseLap_Cuti.tblmaster_jobpositionDataTable
    Dim TBL_Dept As New DseLap_Cuti.tblmaster_departemenDataTable
    Dim TBL_Leave As New DseLap_Cuti.tblmaster_cutiDataTable
    Dim Report_Rekap As New CrpLap_Cuti_Rekap
    Dim Report_Detail As New CrpLap_Cuti_Detail

    ''Crystal Report
    Dim DS As New DseLap_Cuti
    Dim TBL As DseLap_Cuti.Tbl_CetakDataTable = DS.Tbl_Cetak
    Dim Row As DseLap_Cuti.Tbl_CetakRow

    'Cetak Dummy
    Dim DTADummy As New DseLap_CutiTableAdapters.Tbl_CetakTableAdapter
    Dim TBLDummy As New DseLap_Cuti.Tbl_CetakDataTable
    Dim RowDummy As DseLap_Cuti.Tbl_CetakRow

    'Cuti
    Dim DTAAttendance As New DseLap_CutiTableAdapters.TblTrans_AttendanceTableAdapter
    Dim TblAttendance As New DseLap_Cuti.TblTrans_AttendanceDataTable
    Dim RowAttendance As DseLap_Cuti.TblTrans_AttendanceRow

    'Pengajuan
    Dim DTAPengajuanCuti As New DseLap_CutiTableAdapters.TblTrans_CutiPengajuanTableAdapter
    Dim TblPengajuanCuti As New DseLap_Cuti.TblTrans_CutiPengajuanDataTable
    Dim RowPengajuanCuti As DseLap_Cuti.TblTrans_CutiPengajuanRow

    'MAX Date
    Dim DTAMAX As New DseLap_CutiTableAdapters.TblGet_Tanggal_MAXTableAdapter
    Dim TBLMAX As New DseLap_Cuti.TblGet_Tanggal_MAXDataTable
    Dim RowMAX As DseLap_Cuti.TblGet_Tanggal_MAXRow

    Private Enum EnumLaporan As Integer
        AllData
        JobPosition
        Departemen
        Leave
        Employee
    End Enum

    Private Enum EnumType As Integer
        Recap
        Detail
    End Enum

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanCuti_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.MinimumSize = New Drawing.Point(1000, 670)
            Me.dgvEmployee.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvEmployee.AllowUserToAddRows = False
            Me.IsiListBoxJobPosition(Me.ClbJobPosition, Me.TBL_JP)
            Me.IsiListBoxDepartemen(Me.ClbDepartemen, Me.TBL_Dept)
            Me.IsiListBoxCuti(Me.clbLeave, Me.TBL_Leave)
            Me.KendalikanKontrol(EnumLaporan.AllData, False)
            Me.cboFormat.SelectedIndex = 0
            Me.CboFilterEmployee.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiListBoxJobPosition(ByVal Lbo As System.Windows.Forms.CheckedListBox, ByVal TBL As DseLap_Cuti.tblmaster_jobpositionDataTable)
        Dim DTA As New DseLap_CutiTableAdapters.tblmaster_jobpositionTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            TBL.Rows.Clear()
            DTA.Fill(TBL)
            Lbo.DataSource = Nothing
            Lbo.DataSource = TBL.DefaultView

            'Isikan
            Lbo.DisplayMember = "Position"
            For i As Integer = 0 To Lbo.Items.Count - 1
                Lbo.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub IsiListBoxDepartemen(ByVal Lbo As System.Windows.Forms.CheckedListBox, ByVal TBL As DseLap_Cuti.tblmaster_departemenDataTable)
        Dim DTA As New DseLap_CutiTableAdapters.tblmaster_departemenTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            TBL.Rows.Clear()
            DTA.Fill(TBL)
            Lbo.DataSource = Nothing
            Lbo.DataSource = TBL.DefaultView

            'Isikan
            Lbo.DisplayMember = "Departement"
            For i As Integer = 0 To Lbo.Items.Count - 1
                Lbo.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiListBoxCuti(ByVal Lbo As System.Windows.Forms.CheckedListBox, ByVal TBL As DseLap_Cuti.tblmaster_cutiDataTable)
        Dim DTA As New DseLap_CutiTableAdapters.tblmaster_cutiTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            TBL.Rows.Clear()
            DTA.Fill(TBL)
            Lbo.DataSource = Nothing
            Lbo.DataSource = TBL.DefaultView

            'Isikan
            Lbo.DisplayMember = "code"
            For i As Integer = 0 To Lbo.Items.Count - 1
                Lbo.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Lbl_Dept_Checked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Lbl_Dept_Checked.Click
        Try
            For i As Integer = 0 To Me.ClbDepartemen.Items.Count - 1
                Me.ClbDepartemen.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
            ' Me.ChkDepartemen.Checked = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub Lbl_Dept_UnChecked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Lbl_Dept_UnChecked.Click
        Try
            For i As Integer = 0 To Me.ClbDepartemen.Items.Count - 1
                Me.ClbDepartemen.SetItemCheckState(i, Windows.Forms.CheckState.Unchecked)
            Next
            ' Me.ChkDepartemen.Checked = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub Lbl_JP_Checked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Lbl_JP_Checked.Click
        Try
            For i As Integer = 0 To Me.ClbJobPosition.Items.Count - 1
                Me.ClbJobPosition.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
            '  Me.ChkJobPosition.Checked = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub Lbl_JP_UnChecked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Lbl_JP_UnChecked.Click
        Try
            For i As Integer = 0 To Me.ClbJobPosition.Items.Count - 1
                Me.ClbJobPosition.SetItemCheckState(i, Windows.Forms.CheckState.Unchecked)
            Next
            '  Me.ChkJobPosition.Checked = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub Lbl_Leave_Checked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Lbl_Leave_Checked.Click
        Try
            For i As Integer = 0 To Me.clbLeave.Items.Count - 1
                Me.clbLeave.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
            '   Me.ChkLeaveData.Checked = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub Lbl_Leave_UnChecked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Lbl_Leave_UnChecked.Click
        Try
            For i As Integer = 0 To Me.clbLeave.Items.Count - 1
                Me.clbLeave.SetItemCheckState(i, Windows.Forms.CheckState.Unchecked)
            Next
            '  Me.ChkLeaveData.Checked = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.Lbl_Dept_Checked.Enabled = True
            Me.Lbl_Dept_UnChecked.Enabled = True
            Me.Lbl_JP_Checked.Enabled = True
            Me.Lbl_JP_UnChecked.Enabled = True
            Me.Lbl_Leave_Checked.Enabled = True
            Me.Lbl_Leave_UnChecked.Enabled = True
            Me.ClbDepartemen.Enabled = True
            Me.ClbJobPosition.Enabled = True
            Me.clbLeave.Enabled = True
            'Me.dgvEmployee.AllowUserToAddRows = False
            'Me.dgvEmployee.Rows.Clear()
            'Me.dgvEmployee.AllowUserToAddRows = True
            Me.dgvEmployee.Enabled = True
            For i As Integer = 0 To Me.ClbDepartemen.Items.Count - 1
                Me.ClbDepartemen.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
            For i As Integer = 0 To Me.ClbJobPosition.Items.Count - 1
                Me.ClbJobPosition.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
            For i As Integer = 0 To Me.clbLeave.Items.Count - 1
                Me.clbLeave.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.Lbl_Dept_Checked.Enabled = False
            Me.Lbl_Dept_UnChecked.Enabled = False
            Me.Lbl_JP_Checked.Enabled = False
            Me.Lbl_JP_UnChecked.Enabled = False
            Me.Lbl_Leave_Checked.Enabled = False
            Me.Lbl_Leave_UnChecked.Enabled = False
            Me.ClbDepartemen.Enabled = False
            Me.ClbJobPosition.Enabled = False
            Me.clbLeave.Enabled = False
            'Me.dgvEmployee.AllowUserToAddRows = False
            'Me.dgvEmployee.Rows.Clear()
            'Me.dgvEmployee.AllowUserToAddRows = True
            Me.dgvEmployee.Enabled = False
            For i As Integer = 0 To Me.ClbDepartemen.Items.Count - 1
                Me.ClbDepartemen.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
            For i As Integer = 0 To Me.ClbJobPosition.Items.Count - 1
                Me.ClbJobPosition.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
            For i As Integer = 0 To Me.clbLeave.Items.Count - 1
                Me.clbLeave.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KendalikanKontrol(ByVal vFilter As EnumLaporan, ByVal vAktif As Boolean)
        Try
            If vFilter = EnumLaporan.AllData Then
                If vAktif = False Then
                    Me.NonAktifkanDetail()
                Else
                    Me.AktifkanDetail()
                End If
            End If
            If vFilter = EnumLaporan.JobPosition Then
                If vAktif = False Then
                    Me.Lbl_JP_Checked.Enabled = False
                    Me.Lbl_JP_UnChecked.Enabled = False
                    Me.ClbJobPosition.Enabled = False
                    For i As Integer = 0 To Me.ClbJobPosition.Items.Count - 1
                        Me.ClbJobPosition.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
                    Next
                Else
                    Me.Lbl_JP_Checked.Enabled = True
                    Me.Lbl_JP_UnChecked.Enabled = True
                    Me.ClbJobPosition.Enabled = True
                End If
            End If
            If vFilter = EnumLaporan.Departemen Then
                If vAktif = False Then
                    Me.Lbl_Dept_Checked.Enabled = False
                    Me.Lbl_Dept_UnChecked.Enabled = False
                    Me.ClbDepartemen.Enabled = False
                    For i As Integer = 0 To Me.ClbDepartemen.Items.Count - 1
                        Me.ClbDepartemen.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
                    Next
                Else
                    Me.Lbl_Dept_Checked.Enabled = True
                    Me.Lbl_Dept_UnChecked.Enabled = True
                    Me.ClbDepartemen.Enabled = True
                End If
            End If
            If vFilter = EnumLaporan.Leave Then
                If vAktif = False Then
                    Me.Lbl_Leave_Checked.Enabled = False
                    Me.Lbl_Leave_UnChecked.Enabled = False
                    Me.clbLeave.Enabled = False
                    For i As Integer = 0 To Me.clbLeave.Items.Count - 1
                        Me.clbLeave.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
                    Next
                Else
                    Me.Lbl_Leave_Checked.Enabled = True
                    Me.Lbl_Leave_UnChecked.Enabled = True
                    Me.clbLeave.Enabled = True
                End If
            End If
            If vFilter = EnumLaporan.Employee Then
                Me.dgvEmployee.AllowUserToAddRows = False
                If vAktif = False Then
                    Me.dgvEmployee.Enabled = False
                    Me.dgvEmployee.Rows.Clear()
                Else
                    Me.dgvEmployee.Enabled = True
                    Me.dgvEmployee.Rows.Clear()
                End If
                Me.dgvEmployee.AllowUserToAddRows = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkDepartemen_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkDepartemen.CheckedChanged
        Try
            If Me.ChkDepartemen.Checked = False Then
                Me.KendalikanKontrol(EnumLaporan.Departemen, False)
            Else
                Me.KendalikanKontrol(EnumLaporan.Departemen, True)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkEmployee_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkEmployee.CheckedChanged
        Try
            If Me.ChkEmployee.Checked = False Then
                Me.KendalikanKontrol(EnumLaporan.Employee, False)
                Me.dgvEmployee.AllowUserToAddRows = False
            Else
                Me.KendalikanKontrol(EnumLaporan.Employee, True)
                Me.dgvEmployee.AllowUserToAddRows = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkJobPosition_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkJobPosition.CheckedChanged
        Try
            If Me.ChkJobPosition.Checked = False Then
                Me.KendalikanKontrol(EnumLaporan.JobPosition, False)
            Else
                Me.KendalikanKontrol(EnumLaporan.JobPosition, True)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkLeaveData_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkLeaveData.CheckedChanged
        Try
            If Me.ChkLeaveData.Checked = False Then
                Me.KendalikanKontrol(EnumLaporan.Leave, False)
            Else
                Me.KendalikanKontrol(EnumLaporan.Leave, True)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemInput(ByVal vPrimaryEmployee As Double) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.dgvEmployee.Rows
                If rowDGV.Cells(Me.ColPrimaryCode.Index).Value = vPrimaryEmployee Then
                    Return rowDGV.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub dgvEmployee_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvEmployee.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColCode.Index Then
                Me.Buffer = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmployee_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvEmployee.CellEndEdit
        Dim Row As DataGridViewRow = Me.dgvEmployee.Rows(e.RowIndex)
        Dim Index As Integer = -1
        Try
            With Row
                If IsNothing(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If

                    ''Cell kode 
                    If e.ColumnIndex = Me.ColCode.Index Then
                        Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjProduct.ApakahSudahAda(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary aku
                            Index = Me.GetIndexItemInput(Me.ObjProduct.GetPrimaryByCode(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value))

                            If Index < 0 Then
                                Me.dgvEmployee.Item(Me.ColPrimaryCode.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimaryByCode(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.dgvEmployee.Item(Me.ColName.Index, e.RowIndex).Value = Me.ObjProduct.GetFullNameByPrimary(Me.dgvEmployee.Item(Me.ColPrimaryCode.Index, e.RowIndex).Value)
                                Me.dgvEmployee.Item(Me.ColNIK.Index, e.RowIndex).Value = Me.ObjProduct.GetNIK(Me.dgvEmployee.Item(Me.ColPrimaryCode.Index, e.RowIndex).Value)
                            Else
                                If e.RowIndex <> Index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvEmployee.Item(Me.ColCode.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCekData.SetNumber(Me.dgvItem)
                                    Exit Sub
                                End If
                            End If
                        Else
                            '2. kode akun yang diketik tidak ada --> tampilkan form cari
                            Dim ij As New Core_Module.FrmMasterKaryawanCari
                            ij.txtCari.Text = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryEmploye <> -1 Then
                                '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                                Index = Me.GetIndexItemInput(ij.PrimaryEmploye)
                                If Index < 0 Then
                                    Me.dgvEmployee.Item(Me.ColCode.Index, e.RowIndex).Value = Me.ObjProduct.GetCodeByPrimary(ij.PrimaryEmploye)
                                    Me.dgvEmployee.Item(Me.ColName.Index, e.RowIndex).Value = Me.ObjProduct.GetFullNameByPrimary(ij.PrimaryEmploye)
                                    Me.dgvEmployee.Item(Me.ColNIK.Index, e.RowIndex).Value = Me.ObjProduct.GetNIK(ij.PrimaryEmploye)
                                    Me.dgvEmployee.Item(Me.ColPrimaryCode.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                Else
                                    If e.RowIndex <> Index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCekData.SetNumber(Me.dgvItem)
                                        Exit Sub
                                    Else
                                        Me.dgvEmployee.Item(Me.ColCode.Index, e.RowIndex).Value = Me.ObjProduct.GetCodeByPrimary(ij.PrimaryEmploye)
                                        Me.dgvEmployee.Item(Me.ColName.Index, e.RowIndex).Value = Me.ObjProduct.GetFullNameByPrimary(ij.PrimaryEmploye)
                                        Me.dgvEmployee.Item(Me.ColNIK.Index, e.RowIndex).Value = Me.ObjProduct.GetNIK(ij.PrimaryEmploye)
                                        Me.dgvEmployee.Item(Me.ColPrimaryCode.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                    End If
                                End If
                            ElseIf ij.PrimaryEmploye = -1 Then
                                If Me.dgvEmployee.Rows.Count - 1 <> e.RowIndex Then
                                    Me.EraseRowIndex = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCekData.SetNumber(Me.dgvItem)
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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
            Me.dgvEmployee.Rows.RemoveAt(Me.EraseRowIndex)
            'Me.ObjCekData.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetSelectionData(ByVal vTBLDummy As DseLap_Cuti.Tbl_CetakDataTable, _
    ByVal vEnumTipeLaporan As EnumLaporan)
        Dim RowDummy As DseLap_Cuti.Tbl_CetakRow
        Dim RowView As Data.DataRowView
        Dim RowDGV As DataGridViewRow
        Dim Primary As Double = -1
        Dim Ketemu As Boolean = False
        Dim ij As Integer = 0
        Try
            If vTBLDummy.Rows.Count > 0 Then
                ij = 0
                If vEnumTipeLaporan = EnumLaporan.Departemen Then
                    While ij <= vTBLDummy.Rows.Count - 1
                        RowDummy = vTBLDummy.Rows(ij)
                        Ketemu = False
                        If ClbDepartemen.CheckedItems.Count > 0 Then
                            For Each RowView In Me.ClbDepartemen.CheckedItems
                                Primary = RowView.Item("Primary")
                                If RowDummy.primary_departemen = Primary Then
                                    Ketemu = True
                                    Exit For
                                End If
                            Next
                        End If

                        If Ketemu = False Then
                            vTBLDummy.Rows.RemoveAt(ij)
                        Else
                            ij += 1
                        End If
                    End While
                End If

                If vEnumTipeLaporan = EnumLaporan.Employee Then
                    ij = 0
                    Me.dgvEmployee.AllowUserToDeleteRows = False
                    Me.dgvEmployee.AllowUserToAddRows = False
                    If vTBLDummy.Rows.Count > 0 Then
                        While ij <= vTBLDummy.Rows.Count - 1
                            RowDummy = vTBLDummy.Rows(ij)
                            Ketemu = False
                            If Me.dgvEmployee.Rows.Count > 0 Then
                                For Each RowDGV In Me.dgvEmployee.Rows
                                    Primary = RowDGV.Cells(Me.ColPrimaryCode.Index).Value
                                    If RowDummy.primary_karyawan = Primary Then
                                        Ketemu = True
                                        Exit For
                                    End If
                                Next
                            End If
                            If Ketemu = False Then
                                vTBLDummy.Rows.RemoveAt(ij)
                            Else
                                ij += 1
                            End If
                        End While
                    End If
                    Me.dgvEmployee.AllowUserToDeleteRows = True
                    Me.dgvEmployee.AllowUserToAddRows = True
                End If

                If vEnumTipeLaporan = EnumLaporan.JobPosition Then
                    ij = 0
                    While ij <= vTBLDummy.Rows.Count - 1
                        RowDummy = vTBLDummy.Rows(ij)
                        Ketemu = False
                        If Me.ClbJobPosition.CheckedItems.Count > 0 Then
                            For Each RowView In Me.ClbJobPosition.CheckedItems
                                Primary = RowView.Item("primary")
                                If RowDummy.primary_position = Primary Then
                                    Ketemu = True
                                    Exit For
                                End If
                            Next
                        End If

                        If Ketemu = False Then
                            vTBLDummy.Rows.RemoveAt(ij)
                        Else
                            ij += 1
                        End If
                    End While
                End If

                If vEnumTipeLaporan = EnumLaporan.Leave Then
                    ij = 0
                    While ij <= vTBLDummy.Rows.Count - 1
                        RowDummy = vTBLDummy.Rows(ij)
                        Ketemu = False
                        If clbLeave.CheckedItems.Count > 0 Then
                            For Each RowView In Me.clbLeave.CheckedItems
                                Primary = RowView.Item("primary")
                                If RowDummy.primary_leave = Primary Then
                                    Ketemu = True
                                    Exit For
                                End If
                            Next
                        End If

                        If Ketemu = False Then
                            vTBLDummy.Rows.RemoveAt(ij)
                        Else
                            ij += 1
                        End If
                    End While
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BuatLaporanRekap(ByVal vCRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vPeriod As Date)
        Me.Cursor = Cursors.WaitCursor
        'Variable
        Dim ObjCD As New Core_Module.ClsCekData
        Dim vStartPeriod As Date = ObjCD.GetFisrtDateOfTheYear(vPeriod)
        Dim vEndPeriod As Date = vPeriod
        Dim vDepartement As String = ""
        Dim Filter As String = ""
        Dim RowView As Data.DataRowView
        Dim Index As Integer = 0
        Dim ListKategori As New List(Of String)
        Dim TotalRow As Integer = 0
        Dim DV As DataRowView

        Try
            DTADummy.Connection.ConnectionString = Core_Module.GetConstring
            DTAAttendance.Connection.ConnectionString = Core_Module.GetConstring
            DTAPengajuanCuti.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ChkDepartemen.Checked = True Then
                For Each RowView In Me.ClbDepartemen.CheckedItems
                    If String.IsNullOrEmpty(vDepartement) = True Then
                        vDepartement = RowView.Item("Departement").ToString
                    Else
                        vDepartement += ", " + RowView.Item("Departement").ToString
                    End If
                Next
            Else
                vDepartement = "All Data"
            End If

            TBLDummy.Rows.Clear()
            TblAttendance.Rows.Clear()
            TblPengajuanCuti.Rows.Clear()
            DTADummy.Fill(TBLDummy, vStartPeriod)
            '@ tambahan filteran active or not active for employee data
            Me.FilterAktifEmployee(TBLDummy)
            'Cek berkaitan dengan Filter Kategori data
            If Me.ChkDepartemen.Checked = True Then
                Me.GetSelectionData(TBLDummy, EnumLaporan.Departemen)
            End If
            If Me.ChkEmployee.Checked = True Then
                Me.GetSelectionData(TBLDummy, EnumLaporan.Employee)
            End If
            If Me.ChkJobPosition.Checked = True Then
                Me.GetSelectionData(TBLDummy, EnumLaporan.JobPosition)
            End If
            If Me.ChkLeaveData.Checked = True Then
                Me.GetSelectionData(TBLDummy, EnumLaporan.Leave)
            End If
            DTAAttendance.Fill(TblAttendance, vEndPeriod, vStartPeriod)
            DTAPengajuanCuti.Fill(TblPengajuanCuti, vEndPeriod, vStartPeriod)
            DTAMAX.Fill(TBLMAX, vPeriod)
            Me.DS.Tbl_Cetak.Rows.Clear()

            'Fill Primary MAX
            For Each RowAttendance In TblAttendance.Rows
                Filter = "[primary karyawan] = " + RowAttendance.primary_karyawan.ToString + ""
                TBLMAX.DefaultView.RowFilter = Filter
                If TBLMAX.DefaultView.Count > 0 Then
                    RowMAX = TBLMAX.DefaultView(0).Row
                    RowAttendance.Tanggal_MAX = RowMAX.MAX_Date
                Else
                    RowAttendance.Tanggal_MAX = #1/1/1950#
                End If
            Next
            For Each RowPengajuanCuti In TblPengajuanCuti.Rows
                Filter = "[primary karyawan] = " + RowPengajuanCuti.primary_karyawan.ToString + ""
                TBLMAX.DefaultView.RowFilter = Filter
                If TBLMAX.DefaultView.Count > 0 Then
                    RowMAX = TBLMAX.DefaultView(0).Row
                    RowPengajuanCuti.Tanggal_MAX = RowMAX.MAX_Date
                Else
                    RowPengajuanCuti.Tanggal_MAX = #1/1/1950#
                End If
            Next

            ''Cari data yang harus dilengkapi
            If TBLDummy.Rows.Count > 0 Then
                For Each RowDummy In TBLDummy.Rows
                    Dim decTotal As Decimal = 0
                    Dim JmlBaris As Integer = 0

                    Filter = "[primary karyawan] = " + RowDummy.primary_karyawan.ToString + _
                    " AND [primary cuti] = " + RowDummy.primary_leave.ToString
                    TblAttendance.DefaultView.RowFilter = Filter
                    For Each DV In TblAttendance.DefaultView
                        RowAttendance = DV.Row
                        If RowAttendance._date <= RowAttendance.Tanggal_MAX Then
                            With RowDummy
                                .Used_Qty += RowAttendance.is_cuti
                                decTotal += RowAttendance.is_cuti
                            End With
                        End If
                    Next

                    '==============================================================================================================================
                    'Filter TblPengajuanCuti and calculate
                    TblPengajuanCuti.DefaultView.RowFilter = Filter
                    For Each DV In TblPengajuanCuti.DefaultView
                        RowPengajuanCuti = DV.Row
                        If RowPengajuanCuti._date > RowPengajuanCuti.Tanggal_MAX Then
                            With RowDummy
                                .Used_Qty += RowPengajuanCuti.duration
                                decTotal += RowPengajuanCuti.duration
                            End With
                        End If
                    Next

                    'masukkan data cuti ke tabel cetak
                    Row = TBL.NewTbl_CetakRow
                    With Row
                        .Begining_Qty = RowDummy.Begining_Qty
                        .Dept_Code = RowDummy.Dept_Code
                        .Dept_Desc = RowDummy.Dept_Desc
                        .End_Qty = RowDummy.End_Qty
                        .Full_Name = RowDummy.Full_Name
                        .Leave_Code = RowDummy.Leave_Code
                        .Leave_Desc = RowDummy.Leave_Desc
                        .NIK = RowDummy.NIK
                        .No_Pengajuan_Cuti = ""
                        .Period = RowDummy.Period
                        .primary_departemen = RowDummy.primary_departemen
                        .primary_karyawan = RowDummy.primary_karyawan
                        .primary_leave = RowDummy.primary_leave
                        .primary_position = RowDummy.primary_position
                        .Tanggal = #1/1/1950#
                        .Total_Used_Qty = RowDummy.Total_Used_Qty
                        .Urutan = RowDummy.Urutan
                        .Total_Used_Qty = decTotal
                        .End_Qty = Row.Begining_Qty - Row.Total_Used_Qty
                        .Aktif = RowDummy.Aktif
                    End With
                    TBL.Rows.Add(Row)
                Next
            End If

            'pindahkan data dari tbl cetak ke tbl dummy, persiapan pengurutan data
            TBLDummy.Rows.Clear()
            For Each Row In TBL.Rows
                RowDummy = TBLDummy.NewTbl_CetakRow
                For i As Integer = 0 To TBLDummy.Columns.Count - 1
                    RowDummy.Item(i) = Row.Item(i)
                Next
                TBLDummy.Rows.Add(RowDummy)
            Next
            TBL.Rows.Clear()

            ''==============================================================================================================================
            Dim dtvDummy As New DataView(TBLDummy)
            dtvDummy.Sort = "[Dept Desc], [Full Name], [Leave Desc], Tanggal"

            'Pindahin data dummy ke TBLCetak
            TBL.Rows.Clear()
            ' DTADummy.Fill(TBLDummy, vStartPeriod)
            Dim intD As Integer = 0
            If dtvDummy.Count > 0 Then
                For intD = 0 To dtvDummy.Count - 1
                    Row = TBL.NewRow
                    'dimasuk in berdasarkan field-nya
                    Row.primary_leave = dtvDummy.Item(intD)("primary leave")
                    Row.Dept_Code = dtvDummy.Item(intD)("Dept Code")
                    Row.Dept_Desc = dtvDummy.Item(intD)("Dept Desc")
                    Row.Full_Name = dtvDummy.Item(intD)("Full Name")
                    Row.Leave_Code = dtvDummy.Item(intD)("Leave Code")
                    Row.Leave_Desc = dtvDummy.Item(intD)("Leave Desc")
                    Row.Period = dtvDummy.Item(intD)("Period")
                    Row.Begining_Qty = dtvDummy.Item(intD)("Begining Qty")
                    Row.primary_karyawan = dtvDummy.Item(intD)("primary karyawan")
                    Row.Used_Qty = dtvDummy.Item(intD)("Used Qty")
                    Row.End_Qty = dtvDummy.Item(intD)("End Qty")
                    Row.Tanggal = dtvDummy.Item(intD)("Tanggal")
                    Row.NIK = dtvDummy.Item(intD)("NIK")
                    Row.Total_Used_Qty = dtvDummy.Item(intD)("Total Used Qty")
                    Row.No_Pengajuan_Cuti = dtvDummy.Item(intD)("No Pengajuan Cuti")
                    Row.primary_departemen = dtvDummy.Item(intD)("primary departemen")
                    Row.primary_position = dtvDummy.Item(intD)("primary position")
                    Row.Aktif = dtvDummy.Item(intD)("Aktif")
                    ' Row.Hide = RowDummy.Hide
                    TBL.Rows.Add(Row)
                Next
            End If

            Dim intA As Integer = 0
            Dim intT As Integer = 0
            Dim intX As Integer = 0
            Dim intHapus As Integer = 0
            If TBL.Rows.Count > 0 Then
                For intA = 0 To TBL.Rows.Count - 1
                    If intA > 0 Then
                        '---------------------------------------------------------------------------------------------
                        'ngisi hide --> jika primary karyawan = primary karyawan baris sebelumnya maka hide true 
                        '---------------------------------------------------------------------------------------------
                        If TBL.Rows(intA)("primary karyawan") = TBL.Rows(intA - 1)("primary karyawan") Then
                            intT += 1
                            TBL.Rows(intA)("Hide") = intT
                        Else
                            TBL.Rows(intA)("Hide") = 0
                            intT = 0
                        End If
                        '---------------------------------------------------------------------------------------------
                        '---------------------------------------------------------------------------------------------
                        'ngisi Urutan --> jika hide = false (0) maka baris sebelumnya = 1 (utk garis bawah panjang)
                        '---------------------------------------------------------------------------------------------
                        If TBL.Rows(intA)("Hide") = 0 Then
                            TBL.Rows(intA - 1)("Urutan") = 1
                        End If
                        '---------------------------------------------------------------------------------------------
                    End If
                    '---------------------------------------------------------------------------------------------
                    'ngisi Garis --> jika baris sudah 25, maka garis direset kembali ke 0
                    '---------------------------------------------------------------------------------------------
                    intX += 1

                    If intA = 32 Then
                        intX = 1
                    End If

                    TBL.Rows(intA)("Garis") = intX
                    '---------------------------------------------------------------------------------------------
                    If TBL.Rows(intA)("Used Qty") = 0 Then
                        intHapus += 1
                    End If
                Next
                '---------------------------------------------------------------------------------------------
                'urutan dibaris paling terakhir selalu diisi = 1
                '---------------------------------------------------------------------------------------------
                TBL.Rows(TBL.Rows.Count - 1)("Urutan") = 1
            End If

            'munculkan data ke crp
            Report_Rekap.SetDataSource(DS)
            Report_Rekap.SetParameterValue(Report_Rekap.Parameter_Visible_Detail.ParameterFieldName, 1)
            Report_Rekap.SetParameterValue(Report_Rekap.Parameter_Departement.ParameterFieldName, vDepartement)
            'Report_Detail.SetParameterValue(Report_Detail.Parameter_Visible_Detail.ParameterFieldName, intJmlRow)
            crv1.ReportSource = Report_Rekap
            crv1.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTAAttendance.Dispose()
            TblAttendance.Dispose()
            DTAPengajuanCuti.Dispose()
            TblPengajuanCuti.Dispose()
            TBLDummy.Dispose()
            ObjCD = Nothing
            ListKategori = Nothing
            Me.Cursor = Nothing
            DTAMAX.Dispose()
            TBLMAX.Dispose()
        End Try
    End Sub

    Private Sub BuatLaporan_obs(ByVal vCRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal vTipeLaporan As Integer, ByVal vPeriod As Date)
        Me.Cursor = Cursors.WaitCursor
        Dim Report_Rekap As New CrpLap_Cuti_Rekap
        Dim Report_Detail As New CrpLap_Cuti_Detail
        ''Crystal Report
        Dim DS As New DseLap_Cuti
        Dim TBL As DseLap_Cuti.Tbl_CetakDataTable = DS.Tbl_Cetak
        Dim Row As DseLap_Cuti.Tbl_CetakRow

        'Cetak Dummy
        Dim DTADummy As New DseLap_CutiTableAdapters.Tbl_CetakTableAdapter
        Dim TBLDummy As New DseLap_Cuti.Tbl_CetakDataTable
        Dim RowDummy As DseLap_Cuti.Tbl_CetakRow

        'Cuti
        Dim DTAAttendance As New DseLap_CutiTableAdapters.TblTrans_AttendanceTableAdapter
        Dim TblAttendance As New DseLap_Cuti.TblTrans_AttendanceDataTable
        Dim RowAttendance As DseLap_Cuti.TblTrans_AttendanceRow

        'Pengajuan
        Dim DTAPengajuanCuti As New DseLap_CutiTableAdapters.TblTrans_CutiPengajuanTableAdapter
        Dim TblPengajuanCuti As New DseLap_Cuti.TblTrans_CutiPengajuanDataTable
        Dim RowPengajuanCuti As DseLap_Cuti.TblTrans_CutiPengajuanRow

        'MAX Date
        Dim DTAMAX As New DseLap_CutiTableAdapters.TblGet_Tanggal_MAXTableAdapter
        Dim TBLMAX As New DseLap_Cuti.TblGet_Tanggal_MAXDataTable
        Dim RowMAX As DseLap_Cuti.TblGet_Tanggal_MAXRow

        'Variable
        Dim ObjCD As New Core_Module.ClsCekData
        Dim vStartPeriod As Date = ObjCD.GetFisrtDateOfTheYear(vPeriod)
        Dim vEndPeriod As Date = vPeriod
        Dim vDepartement As String = ""
        Dim Filter As String = ""
        Dim RowView As Data.DataRowView
        Dim Index As Integer = 0
        Dim ListKategori As New List(Of String)
        Dim TotalRow As Integer = 0

        Try
            DTADummy.Connection.ConnectionString = Core_Module.GetConstring
            DTAAttendance.Connection.ConnectionString = Core_Module.GetConstring
            DTAPengajuanCuti.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ChkDepartemen.Checked = True Then
                For Each RowView In Me.ClbDepartemen.CheckedItems
                    If String.IsNullOrEmpty(vDepartement) = True Then
                        vDepartement = RowView.Item("Departement").ToString
                    Else
                        vDepartement += ", " + RowView.Item("Departement").ToString
                    End If
                Next
            Else
                vDepartement = "All Data"
            End If

            DTADummy.Fill(TBLDummy, vStartPeriod)
            'Cek berkaitan dengan Filter Kategori data
            If Me.ChkDepartemen.Checked = True Then
                Me.GetSelectionData(TBLDummy, EnumLaporan.Departemen)
            End If
            If Me.ChkEmployee.Checked = True Then
                Me.GetSelectionData(TBLDummy, EnumLaporan.Employee)
            End If
            If Me.ChkJobPosition.Checked = True Then
                Me.GetSelectionData(TBLDummy, EnumLaporan.JobPosition)
            End If
            If Me.ChkLeaveData.Checked = True Then
                Me.GetSelectionData(TBLDummy, EnumLaporan.Leave)
            End If
            DTAAttendance.Fill(TblAttendance, vEndPeriod, vStartPeriod)
            DTAPengajuanCuti.Fill(TblPengajuanCuti, vEndPeriod, vStartPeriod)
            DTAMAX.Fill(TBLMAX, vPeriod)

            'Fill Primary MAX
            For Each RowAttendance In TblAttendance.Rows
                Filter = "[primary karyawan] = " + RowAttendance.primary_karyawan.ToString + ""
                TBLMAX.DefaultView.RowFilter = Filter
                If TBLMAX.DefaultView.Count > 0 Then
                    RowMAX = TBLMAX.DefaultView(0).Row
                    RowAttendance.Tanggal_MAX = RowMAX.MAX_Date
                Else
                    RowAttendance.Tanggal_MAX = #1/1/1950#
                End If
            Next
            For Each RowPengajuanCuti In TblPengajuanCuti.Rows
                Filter = "[primary karyawan] = " + RowPengajuanCuti.primary_karyawan.ToString + ""
                TBLMAX.DefaultView.RowFilter = Filter
                If TBLMAX.DefaultView.Count > 0 Then
                    RowMAX = TBLMAX.DefaultView(0).Row
                    RowPengajuanCuti.Tanggal_MAX = RowMAX.MAX_Date
                Else
                    RowPengajuanCuti.Tanggal_MAX = #1/1/1950#
                End If
            Next

            ''Cari data yang harus dilengkapi
            If TBLDummy.Rows.Count > 0 Then
                For Each RowDummy In TBLDummy.Rows
                    Dim decTotal As Decimal = 0

                    For Each RowAttendance In TblAttendance.Rows
                        If RowAttendance._date <= RowAttendance.Tanggal_MAX Then
                            '==============================================================================================================================
                            'Filter TblAttendance and calculate
                            Filter = "[primary leave] = " + RowAttendance.primary_cuti.ToString + _
                                         " And [primary karyawan] = " + RowAttendance.primary_karyawan.ToString + _
                                         ""
                            TBLDummy.DefaultView.RowFilter = Filter
                            If TBLDummy.DefaultView.Count > 0 Then
                                RowDummy = TBLDummy.DefaultView(0).Row
                                RowDummy.Tanggal = RowAttendance._date
                                RowDummy.Used_Qty = RowAttendance.is_cuti
                                ' decTotal += RowAttendance.Duration
                                ' MessageBox.Show(decTotal)
                            End If
                        End If
                    Next

                    '==============================================================================================================================
                    'Filter TblPengajuanCuti and calculate
                    For Each RowPengajuanCuti In TblPengajuanCuti.Rows
                        If RowPengajuanCuti._date > RowPengajuanCuti.Tanggal_MAX Then
                            Filter = " [primary leave] = " + RowPengajuanCuti.primary_cuti.ToString + _
                                     " And [primary karyawan] = " + RowPengajuanCuti.primary_karyawan.ToString + _
                                     ""
                            TBLDummy.DefaultView.RowFilter = Filter
                            If TBLDummy.DefaultView.Count > 0 Then
                                RowDummy = TBLDummy.DefaultView(0).Row
                                RowDummy.No_Pengajuan_Cuti = RowPengajuanCuti.no_transaksi
                                RowDummy.Used_Qty = RowPengajuanCuti.duration - RowDummy.Used_Qty
                            End If
                        End If
                    Next
                Next
                For Each RowDummy In TBLDummy.Rows
                    '  MessageBox.Show(RowDummy.Used_Qty)
                    Filter = " [primary departemen] = " + RowDummy.primary_departemen.ToString + _
                             " And [primary karyawan] = " + RowDummy.primary_karyawan.ToString + _
                             " And [primary leave] = " + RowDummy.primary_leave.ToString + _
                             ""
                    TBLDummy.DefaultView.RowFilter = Filter
                    If TBLDummy.DefaultView.Count > 0 Then
                        RowDummy.Total_Used_Qty += RowDummy.Used_Qty
                    End If

                    '  MessageBox.Show(RowDummy.Total_Used_Qty)
                    RowDummy.End_Qty = RowDummy.Begining_Qty - RowDummy.Total_Used_Qty
                Next
            End If


            ''==============================================================================================================================
            Dim dtvDummy As New DataView(TBLDummy)
            dtvDummy.Sort = "[Dept Code], [Primary karyawan], [Leave Desc]"

            'Pindahin data dummy ke TBLCetak
            TBL.Rows.Clear()
            ' DTADummy.Fill(TBLDummy, vStartPeriod)
            Dim intD As Integer = 0
            If dtvDummy.Count > 0 Then
                For intD = 0 To dtvDummy.Count - 1
                    Row = TBL.NewRow
                    'dimasuk in berdasarkan field-nya
                    Row.primary_leave = dtvDummy.Item(intD)("primary leave")
                    Row.Dept_Code = dtvDummy.Item(intD)("Dept Code")
                    Row.Dept_Desc = dtvDummy.Item(intD)("Dept Desc")
                    Row.Full_Name = dtvDummy.Item(intD)("Full Name")
                    Row.Leave_Code = dtvDummy.Item(intD)("Leave Code")
                    Row.Leave_Desc = dtvDummy.Item(intD)("Leave Desc")
                    Row.Period = dtvDummy.Item(intD)("Period")
                    Row.Begining_Qty = dtvDummy.Item(intD)("Begining Qty")
                    Row.primary_karyawan = dtvDummy.Item(intD)("primary karyawan")
                    Row.Used_Qty = dtvDummy.Item(intD)("Used Qty")
                    Row.End_Qty = dtvDummy.Item(intD)("End Qty")
                    Row.Tanggal = dtvDummy.Item(intD)("Tanggal")
                    Row.NIK = dtvDummy.Item(intD)("NIK")
                    Row.Total_Used_Qty = dtvDummy.Item(intD)("Total Used Qty")
                    Row.No_Pengajuan_Cuti = dtvDummy.Item(intD)("No Pengajuan Cuti")
                    Row.primary_departemen = dtvDummy.Item(intD)("primary departemen")
                    Row.primary_position = dtvDummy.Item(intD)("primary position")
                    ' Row.Hide = RowDummy.Hide
                    TBL.Rows.Add(Row)
                Next
            End If

            ''Pindahin data dummy ke TBLCetak
            'TBL.Rows.Clear()
            '' DTADummy.Fill(TBLDummy, vStartPeriod)
            'For Each RowDummy In TBLDummy.Rows
            '    Row = TBL.NewRow
            '    'dimasuk in berdasarkan field-nya
            '    Row.primary_leave = RowDummy.primary_leave
            '    Row.Dept_Code = RowDummy.Dept_Code
            '    Row.Dept_Desc = RowDummy.Dept_Desc
            '    Row.Full_Name = RowDummy.Full_Name
            '    Row.Leave_Code = RowDummy.Leave_Code
            '    Row.Leave_Desc = RowDummy.Leave_Desc
            '    Row.Period = RowDummy.Period
            '    Row.Begining_Qty = RowDummy.Begining_Qty
            '    Row.primary_karyawan = RowDummy.primary_karyawan
            '    Row.Used_Qty = RowDummy.Used_Qty
            '    Row.End_Qty = RowDummy.End_Qty
            '    Row.Tanggal = RowDummy.Tanggal
            '    Row.NIK = RowDummy.NIK
            '    Row.Total_Used_Qty = RowDummy.Total_Used_Qty
            '    Row.No_Pengajuan_Cuti = RowDummy.No_Pengajuan_Cuti
            '    Row.primary_departemen = RowDummy.primary_departemen
            '    Row.primary_position = RowDummy.primary_position
            '    ' Row.Hide = RowDummy.Hide
            '    TBL.Rows.Add(Row)
            'Next


            ''While Index <= TBL.Rows.Count - 1
            ''    Row = TBL.Rows(Index)
            ''    If ListKategori.Contains(Row.primary_karyawan.ToString) = False Then
            ''        ListKategori.Add(Row.primary_karyawan.ToString)
            ''        If Index > 0 Then
            ''            RowDummy = TBL.Rows(Index - 1)
            ''            RowDummy.Garis = 1
            ''        End If
            ''    End If
            ''    Index += 1
            ''End While

            Dim intA As Integer = 0
            Dim intT As Integer = 0
            Dim intX As Integer = 0
            Dim intHapus As Integer = 0
            If TBL.Rows.Count > 0 Then
                For intA = 0 To TBL.Rows.Count - 1
                    If intA > 0 Then
                        '---------------------------------------------------------------------------------------------
                        'ngisi hide --> jika primary karyawan = primary karyawan baris sebelumnya maka hide true 
                        '---------------------------------------------------------------------------------------------
                        If TBL.Rows(intA)("primary karyawan") = TBL.Rows(intA - 1)("primary karyawan") Then
                            intT += 1
                            TBL.Rows(intA)("Hide") = intT
                        Else
                            TBL.Rows(intA)("Hide") = 0
                            intT = 0
                        End If
                        '---------------------------------------------------------------------------------------------
                        '---------------------------------------------------------------------------------------------
                        'ngisi Urutan --> jika hide = false (0) maka baris sebelumnya = 1 (utk garis bawah panjang)
                        '---------------------------------------------------------------------------------------------
                        If TBL.Rows(intA)("Hide") = 0 Then
                            TBL.Rows(intA - 1)("Urutan") = 1
                        End If
                        '---------------------------------------------------------------------------------------------
                    End If
                    '---------------------------------------------------------------------------------------------
                    'ngisi Garis --> jika baris sudah 25, maka garis direset kembali ke 0
                    '---------------------------------------------------------------------------------------------
                    intX += 1

                    If vTipeLaporan = 0 Then
                        If intA = 35 Then
                            intX = 1
                        End If
                    Else
                        If intA = 32 Then
                            intX = 1
                        End If
                    End If

                    TBL.Rows(intA)("Garis") = intX
                    '---------------------------------------------------------------------------------------------
                    If TBL.Rows(intA)("Used Qty") = 0 Then
                        intHapus += 1
                    End If
                Next
                '---------------------------------------------------------------------------------------------
                'urutan dibaris paling terakhir selalu diisi = 1
                '---------------------------------------------------------------------------------------------
                TBL.Rows(TBL.Rows.Count - 1)("Urutan") = 1
            End If

            'munculkan data ke crp
            If vTipeLaporan = 0 Then
                Report_Rekap.SetDataSource(DS)
                Report_Rekap.SetParameterValue(Report_Rekap.Parameter_Visible_Detail.ParameterFieldName, vTipeLaporan)
                Report_Rekap.SetParameterValue(Report_Rekap.Parameter_Departement.ParameterFieldName, vDepartement)
                'Report_Rekap.SetParameterValue(Report_Rekap.Parameter_TotalRow.ParameterFieldName, TotalRow)
                crv1.ReportSource = Report_Rekap
            Else
                Report_Detail.SetDataSource(DS)
                Report_Detail.SetParameterValue(Report_Detail.Parameter_Visible_Detail.ParameterFieldName, vTipeLaporan)
                Report_Detail.SetParameterValue(Report_Detail.Parameter_Departement.ParameterFieldName, vDepartement)
                'Report_Detail.SetParameterValue(Report_Detail.Parameter_Visible_Detail.ParameterFieldName, intJmlRow)
                crv1.ReportSource = Report_Detail
            End If
            crv1.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTAAttendance.Dispose()
            TblAttendance.Dispose()
            DTAPengajuanCuti.Dispose()
            TblPengajuanCuti.Dispose()
            TBLDummy.Dispose()
            ObjCD = Nothing
            ListKategori = Nothing
            Me.Cursor = Nothing
            DTAMAX.Dispose()
            TBLMAX.Dispose()
        End Try
    End Sub

    Private Function FilterAktifEmployee(ByRef vTbl As DataTable) As Boolean
        Dim vRow As DataRow
        Try
            If Me.CboFilterEmployee.SelectedIndex = 1 Then 'aktive
                vTbl.DefaultView.RowFilter = "[Aktif]=1"
                For i As Integer = vTbl.DefaultView.Count - 1 To 0 Step -1
                    vRow = vTbl.DefaultView(0).Row
                    vTbl.Rows.Remove(vRow)
                Next
            ElseIf Me.CboFilterEmployee.SelectedIndex = 2 Then 'not active
                vTbl.DefaultView.RowFilter = "[Aktif]=0"
                For i As Integer = vTbl.DefaultView.Count - 1 To 0 Step -1
                    vRow = vTbl.DefaultView(0).Row
                    vTbl.Rows.Remove(vRow)
                Next
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BuatLaporanDetail(ByVal vCRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vPeriod As Date)
        Me.Cursor = Cursors.WaitCursor
        'Variable
        Dim ObjCD As New Core_Module.ClsCekData
        Dim vStartPeriod As Date = ObjCD.GetFisrtDateOfTheYear(vPeriod)
        Dim vEndPeriod As Date = vPeriod
        Dim vDepartement As String = ""
        Dim Filter As String = ""
        Dim RowView As Data.DataRowView
        Dim Index As Integer = 0
        Dim ListKategori As New List(Of String)
        Dim TotalRow As Integer = 0
        Dim DV As DataRowView

        Try
            DTADummy.Connection.ConnectionString = Core_Module.GetConstring
            DTAAttendance.Connection.ConnectionString = Core_Module.GetConstring
            DTAPengajuanCuti.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ChkDepartemen.Checked = True Then
                For Each RowView In Me.ClbDepartemen.CheckedItems
                    If String.IsNullOrEmpty(vDepartement) = True Then
                        vDepartement = RowView.Item("Departement").ToString
                    Else
                        vDepartement += ", " + RowView.Item("Departement").ToString
                    End If
                Next
            Else
                vDepartement = "All Data"
            End If

            TBLDummy.Rows.Clear()
            TblAttendance.Rows.Clear()
            TblPengajuanCuti.Rows.Clear()
            DTADummy.Fill(TBLDummy, vStartPeriod)
            '@ tambahan filteran active or not active for employee data
            Me.FilterAktifEmployee(TBLDummy)
            'Cek berkaitan dengan Filter Kategori data
            If Me.ChkDepartemen.Checked = True Then
                Me.GetSelectionData(TBLDummy, EnumLaporan.Departemen)
            End If
            If Me.ChkEmployee.Checked = True Then
                Me.GetSelectionData(TBLDummy, EnumLaporan.Employee)
            End If
            If Me.ChkJobPosition.Checked = True Then
                Me.GetSelectionData(TBLDummy, EnumLaporan.JobPosition)
            End If
            If Me.ChkLeaveData.Checked = True Then
                Me.GetSelectionData(TBLDummy, EnumLaporan.Leave)
            End If
            DTAAttendance.Fill(TblAttendance, vEndPeriod, vStartPeriod)
            DTAPengajuanCuti.Fill(TblPengajuanCuti, vEndPeriod, vStartPeriod)
            DTAMAX.Fill(TBLMAX, vPeriod)
            Me.DS.Tbl_Cetak.Rows.Clear()

            'Fill Primary MAX
            For Each RowAttendance In TblAttendance.Rows
                Filter = "[primary karyawan] = " + RowAttendance.primary_karyawan.ToString + ""
                TBLMAX.DefaultView.RowFilter = Filter
                If TBLMAX.DefaultView.Count > 0 Then
                    RowMAX = TBLMAX.DefaultView(0).Row
                    RowAttendance.Tanggal_MAX = RowMAX.MAX_Date
                Else
                    RowAttendance.Tanggal_MAX = #1/1/1950#
                End If
            Next
            For Each RowPengajuanCuti In TblPengajuanCuti.Rows
                Filter = "[primary karyawan] = " + RowPengajuanCuti.primary_karyawan.ToString + ""
                TBLMAX.DefaultView.RowFilter = Filter
                If TBLMAX.DefaultView.Count > 0 Then
                    RowMAX = TBLMAX.DefaultView(0).Row
                    RowPengajuanCuti.Tanggal_MAX = RowMAX.MAX_Date
                Else
                    RowPengajuanCuti.Tanggal_MAX = #1/1/1950#
                End If
            Next

            ''Cari data yang harus dilengkapi
            If TBLDummy.Rows.Count > 0 Then
                For Each RowDummy In TBLDummy.Rows
                    Dim decTotal As Decimal = 0
                    Dim JmlBaris As Integer = 0

                    Filter = "[primary karyawan] = " + RowDummy.primary_karyawan.ToString + _
                    " AND [primary cuti] = " + RowDummy.primary_leave.ToString
                    TblAttendance.DefaultView.RowFilter = Filter
                    For Each DV In TblAttendance.DefaultView
                        RowAttendance = DV.Row
                        If RowAttendance._date <= RowAttendance.Tanggal_MAX Then
                            Row = TBL.NewTbl_CetakRow
                            With Row
                                .Begining_Qty = RowDummy.Begining_Qty
                                .Dept_Code = RowDummy.Dept_Code
                                .Dept_Desc = RowDummy.Dept_Desc
                                .End_Qty = RowDummy.End_Qty
                                .Full_Name = RowDummy.Full_Name
                                .Leave_Code = RowDummy.Leave_Code
                                .Leave_Desc = RowDummy.Leave_Desc
                                .NIK = RowDummy.NIK
                                .No_Pengajuan_Cuti = RowAttendance.no_transaksi
                                .Period = RowDummy.Period
                                .primary_departemen = RowDummy.primary_departemen
                                .primary_karyawan = RowDummy.primary_karyawan
                                .primary_leave = RowDummy.primary_leave
                                .primary_position = RowDummy.primary_position
                                .Tanggal = RowAttendance._date
                                .Total_Used_Qty = RowDummy.Total_Used_Qty
                                .Urutan = RowDummy.Urutan
                                .Used_Qty = RowAttendance.is_cuti
                                decTotal += RowAttendance.is_cuti
                                .Aktif = RowDummy.Aktif
                            End With
                            TBL.Rows.Add(Row)
                            JmlBaris += 1
                        End If
                    Next

                    '==============================================================================================================================
                    'Filter TblPengajuanCuti and calculate
                    TblPengajuanCuti.DefaultView.RowFilter = Filter
                    For Each DV In TblPengajuanCuti.DefaultView
                        RowPengajuanCuti = DV.Row
                        If RowPengajuanCuti._date > RowPengajuanCuti.Tanggal_MAX Then
                            Row = TBL.NewTbl_CetakRow
                            With Row
                                .Begining_Qty = RowDummy.Begining_Qty
                                .Dept_Code = RowDummy.Dept_Code
                                .Dept_Desc = RowDummy.Dept_Desc
                                .End_Qty = RowDummy.End_Qty
                                .Full_Name = RowDummy.Full_Name
                                .Leave_Code = RowDummy.Leave_Code
                                .Leave_Desc = RowDummy.Leave_Desc
                                .NIK = RowDummy.NIK
                                .No_Pengajuan_Cuti = RowPengajuanCuti.no_transaksi
                                .Period = RowDummy.Period
                                .primary_departemen = RowDummy.primary_departemen
                                .primary_karyawan = RowDummy.primary_karyawan
                                .primary_leave = RowDummy.primary_leave
                                .primary_position = RowDummy.primary_position
                                .Tanggal = RowPengajuanCuti._date
                                .Total_Used_Qty = RowDummy.Total_Used_Qty
                                .Urutan = RowDummy.Urutan
                                .Used_Qty = RowPengajuanCuti.duration
                                decTotal += RowPengajuanCuti.duration
                                .Aktif = RowDummy.Aktif
                            End With
                            TBL.Rows.Add(Row)
                        End If
                        JmlBaris += 1
                    Next

                    'masukkan data cuti yang belum digunakan sama sekali
                    If JmlBaris = 0 Then
                        Row = TBL.NewTbl_CetakRow
                        With Row
                            .Begining_Qty = RowDummy.Begining_Qty
                            .Dept_Code = RowDummy.Dept_Code
                            .Dept_Desc = RowDummy.Dept_Desc
                            .End_Qty = RowDummy.End_Qty
                            .Full_Name = RowDummy.Full_Name
                            .Leave_Code = RowDummy.Leave_Code
                            .Leave_Desc = RowDummy.Leave_Desc
                            .NIK = RowDummy.NIK
                            .No_Pengajuan_Cuti = ""
                            .Period = RowDummy.Period
                            .primary_departemen = RowDummy.primary_departemen
                            .primary_karyawan = RowDummy.primary_karyawan
                            .primary_leave = RowDummy.primary_leave
                            .primary_position = RowDummy.primary_position
                            .Tanggal = #1/1/1950#
                            .Total_Used_Qty = RowDummy.Total_Used_Qty
                            .Urutan = RowDummy.Urutan
                            .Used_Qty = 0
                            .Aktif = RowDummy.Aktif
                        End With
                        TBL.Rows.Add(Row)
                    End If

                    'hitung total used = total cuti yang aktual terjadi + total pengajuan yang belum terjadi
                    TBL.DefaultView.RowFilter = "[primary leave] = " + RowDummy.primary_leave.ToString + _
                    " AND [primary karyawan] = " + RowDummy.primary_karyawan.ToString
                    For Each DV In TBL.DefaultView
                        Row = DV.Row
                        Row.Total_Used_Qty = decTotal
                        Row.End_Qty = Row.Begining_Qty - Row.Total_Used_Qty
                    Next
                Next
            End If

            'pindahkan data dari tbl cetak ke tbl dummy, persiapan pengurutan data
            TBLDummy.Rows.Clear()
            For Each Row In TBL.Rows
                RowDummy = TBLDummy.NewTbl_CetakRow
                For i As Integer = 0 To TBLDummy.Columns.Count - 1
                    RowDummy.Item(i) = Row.Item(i)
                Next
                TBLDummy.Rows.Add(RowDummy)
            Next
            TBL.Rows.Clear()

            ''==============================================================================================================================
            Dim dtvDummy As New DataView(TBLDummy)
            dtvDummy.Sort = "[Dept Desc], [Full Name], [Leave Desc], Tanggal"

            'Pindahin data dummy ke TBLCetak
            TBL.Rows.Clear()
            ' DTADummy.Fill(TBLDummy, vStartPeriod)
            Dim intD As Integer = 0
            If dtvDummy.Count > 0 Then
                For intD = 0 To dtvDummy.Count - 1
                    Row = TBL.NewRow
                    'dimasuk in berdasarkan field-nya
                    Row.primary_leave = dtvDummy.Item(intD)("primary leave")
                    Row.Dept_Code = dtvDummy.Item(intD)("Dept Code")
                    Row.Dept_Desc = dtvDummy.Item(intD)("Dept Desc")
                    Row.Full_Name = dtvDummy.Item(intD)("Full Name")
                    Row.Leave_Code = dtvDummy.Item(intD)("Leave Code")
                    Row.Leave_Desc = dtvDummy.Item(intD)("Leave Desc")
                    Row.Period = dtvDummy.Item(intD)("Period")
                    Row.Begining_Qty = dtvDummy.Item(intD)("Begining Qty")
                    Row.primary_karyawan = dtvDummy.Item(intD)("primary karyawan")
                    Row.Used_Qty = dtvDummy.Item(intD)("Used Qty")
                    Row.End_Qty = dtvDummy.Item(intD)("End Qty")
                    Row.Tanggal = dtvDummy.Item(intD)("Tanggal")
                    Row.NIK = dtvDummy.Item(intD)("NIK")
                    Row.Total_Used_Qty = dtvDummy.Item(intD)("Total Used Qty")
                    Row.No_Pengajuan_Cuti = dtvDummy.Item(intD)("No Pengajuan Cuti")
                    Row.primary_departemen = dtvDummy.Item(intD)("primary departemen")
                    Row.primary_position = dtvDummy.Item(intD)("primary position")
                    Row.Aktif = dtvDummy.Item(intD)("Aktif")
                    ' Row.Hide = RowDummy.Hide
                    TBL.Rows.Add(Row)
                Next
            End If

            Dim intA As Integer = 0
            Dim intT As Integer = 0
            Dim intX As Integer = 0
            Dim intHapus As Integer = 0
            If TBL.Rows.Count > 0 Then
                For intA = 0 To TBL.Rows.Count - 1
                    If intA > 0 Then
                        '---------------------------------------------------------------------------------------------
                        'ngisi hide --> jika primary karyawan = primary karyawan baris sebelumnya maka hide true 
                        '---------------------------------------------------------------------------------------------
                        If TBL.Rows(intA)("primary karyawan") = TBL.Rows(intA - 1)("primary karyawan") Then
                            intT += 1
                            TBL.Rows(intA)("Hide") = intT
                        Else
                            TBL.Rows(intA)("Hide") = 0
                            intT = 0
                        End If
                        '---------------------------------------------------------------------------------------------
                        '---------------------------------------------------------------------------------------------
                        'ngisi Urutan --> jika hide = false (0) maka baris sebelumnya = 1 (utk garis bawah panjang)
                        '---------------------------------------------------------------------------------------------
                        If TBL.Rows(intA)("Hide") = 0 Then
                            TBL.Rows(intA - 1)("Urutan") = 1
                        End If
                        '---------------------------------------------------------------------------------------------
                    End If
                    '---------------------------------------------------------------------------------------------
                    'ngisi Garis --> jika baris sudah 25, maka garis direset kembali ke 0
                    '---------------------------------------------------------------------------------------------
                    intX += 1

                    If intA = 32 Then
                        intX = 1
                    End If

                    TBL.Rows(intA)("Garis") = intX
                    '---------------------------------------------------------------------------------------------
                    If TBL.Rows(intA)("Used Qty") = 0 Then
                        intHapus += 1
                    End If
                Next
                '---------------------------------------------------------------------------------------------
                'urutan dibaris paling terakhir selalu diisi = 1
                '---------------------------------------------------------------------------------------------
                TBL.Rows(TBL.Rows.Count - 1)("Urutan") = 1
            End If

            'munculkan data ke crp
            Report_Detail.SetDataSource(DS)
            Report_Detail.SetParameterValue(Report_Detail.Parameter_Visible_Detail.ParameterFieldName, 1)
            Report_Detail.SetParameterValue(Report_Detail.Parameter_Departement.ParameterFieldName, vDepartement)
            'Report_Detail.SetParameterValue(Report_Detail.Parameter_Visible_Detail.ParameterFieldName, intJmlRow)
            crv1.ReportSource = Report_Detail
            crv1.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTAAttendance.Dispose()
            TblAttendance.Dispose()
            DTAPengajuanCuti.Dispose()
            TblPengajuanCuti.Dispose()
            TBLDummy.Dispose()
            ObjCD = Nothing
            ListKategori = Nothing
            Me.Cursor = Nothing
            DTAMAX.Dispose()
            TBLMAX.Dispose()
        End Try
    End Sub

    Private Sub ListEmployee(ByVal vListEmployee As List(Of Double), ByVal vExecute As Boolean)
        Dim RowDGV As New DataGridViewRow
        Dim PrimaryLeaveData As Double
        Try
            vListEmployee.Clear()
            If vExecute = True Then
                Me.dgvEmployee.AllowUserToDeleteRows = False
                Me.dgvEmployee.AllowUserToAddRows = False
                For Each RowDGV In Me.dgvEmployee.Rows
                    PrimaryLeaveData = RowDGV.Cells(Me.ColPrimaryCode.Index).Value
                    If vListEmployee.Contains(PrimaryLeaveData) = False Then
                        vListEmployee.Add(PrimaryLeaveData)
                    End If
                Next
                Me.dgvEmployee.AllowUserToDeleteRows = True
                Me.dgvEmployee.AllowUserToAddRows = True
            End If


            'If vEnumTipeLaporan = EnumLaporan.Employee Then
            '    ij = 0
            '    Me.dgvEmployee.AllowUserToDeleteRows = False
            '    Me.dgvEmployee.AllowUserToAddRows = False
            '    If Me.dgvEmployee.Rows.Count > 0 Then
            '        While ij <= vTBLDummy.Rows.Count - 1
            '            RowDummy = vTBLDummy.Rows(ij)
            '            Ketemu = False
            '            For Each RowDGV In Me.dgvEmployee.Rows
            '                Primary = RowDGV.Cells(Me.ColPrimaryCode.Index).Value
            '                If RowDummy.primary_karyawan = Primary Then
            '                    Ketemu = True
            '                    Exit For
            '                End If
            '            Next
            '            If Ketemu = False Then
            '                vTBLDummy.Rows.RemoveAt(ij)
            '            Else
            '                ij += 1
            '            End If
            '        End While
            '    End If
            '    Me.dgvEmployee.AllowUserToDeleteRows = True
            '    Me.dgvEmployee.AllowUserToAddRows = True
            'End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ListLeaveData(ByVal vListLD As List(Of Double), ByVal vExecute As Boolean)
        Dim RowView As Data.DataRowView
        Dim PrimaryLeaveData As Double
        Try
            vListLD.Clear()
            If vExecute = True Then
                For Each RowView In Me.clbLeave.CheckedItems
                    PrimaryLeaveData = RowView.Row.Item("primary")
                    If vListLD.Contains(PrimaryLeaveData) = False Then
                        vListLD.Add(PrimaryLeaveData)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ListJobPosition(ByVal vListJP As List(Of Double), ByVal vExecute As Boolean)
        Dim RowView As Data.DataRowView
        Dim PrimaryJobPosition As Double
        Try
            vListJP.Clear()
            If vExecute = True Then
                For Each RowView In Me.clbJobPosition.CheckedItems
                    PrimaryJobPosition = RowView.Row.Item("primary")
                    If vListJP.Contains(PrimaryJobPosition) = False Then
                        vListJP.Add(PrimaryJobPosition)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ListDepartement(ByVal vListDept As List(Of Double), ByVal vExecute As Boolean)
        Dim RowView As Data.DataRowView
        Dim PrimaryDepartement As Double
        Try
            vListDept.Clear()
            If vExecute = True Then
                For Each RowView In Me.ClbDepartemen.CheckedItems
                    PrimaryDepartement = RowView.Row.Item("Primary")
                    If vListDept.Contains(PrimaryDepartement) = False Then
                        vListDept.Add(PrimaryDepartement)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function BatasanLaporan() As Boolean
        Dim ListJob As New List(Of Double)
        Dim ListDept As New List(Of Double)
        Dim ListLeave As New List(Of Double)
        Dim ListKaryawan As New List(Of Double)
        Try
            If Me.ChkJobPosition.Checked = True Then
                Me.ListJobPosition(ListJob, True)
                If Me.ClbJobPosition.Items.Count > 0 Then
                    If ListJob.Count <= 0 Then
                        MsgBox("Please checked job position", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If
            If Me.ChkDepartemen.Checked = True Then
                Me.ListDepartement(ListDept, True)
                If Me.ClbDepartemen.Items.Count > 0 Then
                    If ListDept.Count <= 0 Then
                        MsgBox("Please checked department", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            If Me.ChkLeaveData.Checked = True Then
                Me.ListLeaveData(ListLeave, True)
                If Me.clbLeave.Items.Count > 0 Then
                    If ListLeave.Count <= 0 Then
                        MsgBox("Please checked leave data", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            If Me.ChkEmployee.Checked = True Then
                Me.ListEmployee(ListKaryawan, True)
                If Me.dgvEmployee.Rows.Count > 0 Then
                    If ListKaryawan.Count <= 0 Then
                        MsgBox("Please input Karyawan", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            'If Me.ChkJobPosition.Checked = False And _
            '   Me.ChkDepartemen.Checked = False And _
            '   Me.ChkLeaveData.Checked = False And _
            '   Me.ChkEmployee.Checked = False Then
            '    Return False
            'End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListJob = Nothing
            ListDept = Nothing
        End Try
    End Function

    Private Sub btnShow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Dim intType As Integer = 0
        Try
            If BatasanLaporan() = True Then
                If Me.cboFormat.SelectedIndex = 0 Then
                    Me.BuatLaporanRekap(Me.crv1, Me.DtpPeriod.Value)
                Else
                    Me.BuatLaporanDetail(Me.crv1, Me.DtpPeriod.Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class