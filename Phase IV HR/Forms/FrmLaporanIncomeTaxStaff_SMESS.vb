Option Explicit On
Option Strict Off

Imports Core_Module, System.Windows.Forms, System.Drawing
Public Class FrmLaporanIncomeTaxStaff_SMESS
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Dim DefPrec As Double = ObjCurr.GetDecimal_DefaultCurrency()
    Dim TBL_JP As New DseLaporanSalaryStaff_SMESS.tblmaster_jobpositionDataTable
    Dim TBL_Dept As New DseLaporanSalaryStaff_SMESS.tblmaster_departemenDataTable
    Dim TBL_Formula As New DseLaporanSalaryStaff_SMESS.tblmaster_salary_formula1DataTable
    Dim ObjProduct As New Core_Module.ClsMasterKaryawan
    Dim ObjCD As New ClsCekData
    Dim Buffer As String = ""
    Dim AllData As Boolean = True
    Dim EraseRowIndex As Integer = 0

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Keterangan urutan pengecekan
    '1. Department
    '2. Employee
    '3. Job position
    '4. Salary Formula
    Private Sub IsiListBoxFormula(ByVal Lbo As System.Windows.Forms.CheckedListBox, ByVal TBL As DseLaporanSalaryStaff_SMESS.tblmaster_salary_formula1DataTable)
        Dim DTA As New DseLaporanSalaryStaff_SMESSTableAdapters.tblmaster_salary_formula1TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            TBL.Rows.Clear()
            DTA.Fill(TBL)
            Lbo.DataSource = Nothing
            Lbo.DataSource = TBL.DefaultView

            'Isikan
            Lbo.DisplayMember = "Formula"
            For i As Integer = 0 To Lbo.Items.Count - 1
                Lbo.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub IsiListBoxJobPosition(ByVal Lbo As System.Windows.Forms.CheckedListBox, ByVal TBL As DseLaporanSalaryStaff_SMESS.tblmaster_jobpositionDataTable)
        Dim DTA As New DseLaporanSalaryStaff_SMESSTableAdapters.tblmaster_jobpositionTableAdapter
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

    Private Sub IsiListBoxDepartemen(ByVal Lbo As System.Windows.Forms.CheckedListBox, ByVal TBL As DseLaporanSalaryStaff_SMESS.tblmaster_departemenDataTable)
        Dim DTA As New DseLaporanSalaryStaff_SMESSTableAdapters.tblmaster_departemenTableAdapter
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

    Private Sub FrmLaporanSalaryStaff_SMESS_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.Text = " Income Tax Salary"
            Me.dgvEmployee.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            '----------------------
            Me.crvTaxIncome.DisplayBackgroundEdge = True
            Me.crvTaxIncome.DisplayStatusBar = True
            Me.crvTaxIncome.DisplayToolbar = True
            Me.crvTaxIncome.ShowExportButton = True
            Me.crvTaxIncome.ShowGotoPageButton = True
            Me.crvTaxIncome.ShowPageNavigateButtons = True
            Me.crvTaxIncome.ShowPrintButton = True
            Me.crvTaxIncome.ShowTextSearchButton = True
            Me.crvTaxIncome.ShowZoomButton = True
            '----------------------
            Me.dgvEmployee.AllowUserToAddRows = False
            Me.IsiListBoxJobPosition(Me.ClbPositon, Me.TBL_JP)
            Me.IsiListBoxDepartemen(Me.ClbDepartemen, Me.TBL_Dept)
            Me.IsiListBoxFormula(Me.ClbSalaryFormula, Me.TBL_Formula)
            Me.KendalikanKontrol(EnumLaporan.AllData, False)
            Me.DtpStart.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.DtpStart.Value)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Enum Pengendalian 
    Public Enum EnumLaporan As Integer
        AllData
        JobPosition
        Departemen
        SalaryFormula
        Employee
    End Enum

    'Pengendalian Kontrol
    Private Sub AktifkanDetail()
        Try
            Me.Lbl_Departemen_Checked.Enabled = True
            Me.Lbl_Departemen_Unchecked.Enabled = True
            Me.Lbl_Position_Checked.Enabled = True
            Me.Lbl_Position_Unchecked.Enabled = True
            Me.Lbl_SF_Checked.Enabled = True
            Me.Lbl_SF_Unchecked.Enabled = True
            Me.ClbDepartemen.Enabled = True
            Me.ClbPositon.Enabled = True
            Me.ClbSalaryFormula.Enabled = True
            Me.dgvEmployee.Enabled = True
            For i As Integer = 0 To Me.ClbDepartemen.Items.Count - 1
                Me.ClbDepartemen.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
            For i As Integer = 0 To Me.ClbPositon.Items.Count - 1
                Me.ClbPositon.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
            For i As Integer = 0 To Me.ClbSalaryFormula.Items.Count - 1
                Me.ClbSalaryFormula.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.Lbl_Departemen_Checked.Enabled = False
            Me.Lbl_Departemen_Unchecked.Enabled = False
            Me.Lbl_Position_Checked.Enabled = False
            Me.Lbl_Position_Unchecked.Enabled = False
            Me.Lbl_SF_Checked.Enabled = False
            Me.Lbl_SF_Unchecked.Enabled = False
            Me.ClbDepartemen.Enabled = False
            Me.ClbPositon.Enabled = False
            Me.ClbSalaryFormula.Enabled = False
            Me.dgvEmployee.Enabled = False
            For i As Integer = 0 To Me.ClbDepartemen.Items.Count - 1
                Me.ClbDepartemen.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
            For i As Integer = 0 To Me.ClbPositon.Items.Count - 1
                Me.ClbPositon.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
            For i As Integer = 0 To Me.ClbSalaryFormula.Items.Count - 1
                Me.ClbSalaryFormula.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
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
                    Me.Lbl_Position_Checked.Enabled = False
                    Me.Lbl_Position_Unchecked.Enabled = False
                    Me.ClbPositon.Enabled = False
                    For i As Integer = 0 To Me.ClbPositon.Items.Count - 1
                        Me.ClbPositon.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
                    Next
                Else
                    Me.Lbl_Position_Checked.Enabled = True
                    Me.Lbl_Position_Unchecked.Enabled = True
                    Me.ClbPositon.Enabled = True
                End If
            End If
            If vFilter = EnumLaporan.Departemen Then
                If vAktif = False Then
                    Me.Lbl_Departemen_Checked.Enabled = False
                    Me.Lbl_Departemen_Unchecked.Enabled = False
                    Me.ClbDepartemen.Enabled = False
                    For i As Integer = 0 To Me.ClbDepartemen.Items.Count - 1
                        Me.ClbDepartemen.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
                    Next
                Else
                    Me.Lbl_Departemen_Checked.Enabled = True
                    Me.Lbl_Departemen_Unchecked.Enabled = True
                    Me.ClbDepartemen.Enabled = True
                End If
            End If
            If vFilter = EnumLaporan.SalaryFormula Then
                If vAktif = False Then
                    Me.Lbl_SF_Checked.Enabled = False
                    Me.Lbl_SF_Unchecked.Enabled = False
                    Me.ClbSalaryFormula.Enabled = False
                    For i As Integer = 0 To Me.ClbSalaryFormula.Items.Count - 1
                        Me.ClbSalaryFormula.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
                    Next
                Else
                    Me.Lbl_SF_Checked.Enabled = True
                    Me.Lbl_SF_Unchecked.Enabled = True
                    Me.ClbSalaryFormula.Enabled = True
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

    'CheckBox Event
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

    Private Sub ChkPosition_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkPosition.CheckedChanged
        Try
            If Me.ChkPosition.Checked = False Then
                Me.KendalikanKontrol(EnumLaporan.JobPosition, False)
            Else
                Me.KendalikanKontrol(EnumLaporan.JobPosition, True)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkFormula_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkFormula.CheckedChanged
        Try
            If Me.ChkFormula.Checked = False Then
                Me.KendalikanKontrol(EnumLaporan.SalaryFormula, False)
            Else
                Me.KendalikanKontrol(EnumLaporan.SalaryFormula, True)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Lbo Event
    Private Sub Lbl_Position_Checked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Lbl_Position_Checked.Click
        Try
            For i As Integer = 0 To Me.ClbPositon.Items.Count - 1
                Me.ClbPositon.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub Lbl_Position_UnChecked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Lbl_Position_Unchecked.Click
        Try
            For i As Integer = 0 To Me.ClbPositon.Items.Count - 1
                Me.ClbPositon.SetItemCheckState(i, Windows.Forms.CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub Lbl_Departemen_Checked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Lbl_Departemen_Checked.Click
        Try
            For i As Integer = 0 To Me.ClbDepartemen.Items.Count - 1
                Me.ClbDepartemen.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub Lbl_Departemen_UnChecked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Lbl_Departemen_Unchecked.Click
        Try
            For i As Integer = 0 To Me.ClbDepartemen.Items.Count - 1
                Me.ClbDepartemen.SetItemCheckState(i, Windows.Forms.CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub Lbl_SF_Checked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Lbl_SF_Checked.Click
        Try
            For i As Integer = 0 To Me.ClbSalaryFormula.Items.Count - 1
                Me.ClbSalaryFormula.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub Lbl_SF_UnChecked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Lbl_SF_Unchecked.Click
        Try
            For i As Integer = 0 To Me.ClbSalaryFormula.Items.Count - 1
                Me.ClbSalaryFormula.SetItemCheckState(i, Windows.Forms.CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    'DGV Employee
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
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
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

    'Fungsi untuk mengecek Posisi Filter data sebelah kiri
    Private Function CekFilterData(ByVal RowSalary As DseLaporanIncomeTax_SMESS.DataTable1Row) As Boolean
        Dim RowView As DataRowView
        Dim RowDGV As DataGridViewRow
        Dim Primary As Double
        Dim PrimaryFormula As Double
        Dim Ketemu As Boolean
        Try
            'Department
            If Me.ChkDepartemen.Checked = True Then
                Primary = -1
                Ketemu = False
                For Each RowView In Me.ClbDepartemen.CheckedItems
                    Primary = RowView.Row.Item("Primary")
                    If Primary = RowSalary.primary_departemen Then
                        Ketemu = True
                        Exit For
                    End If
                Next
                If Ketemu = False Then
                    Return False
                End If
            Else
                'All data
                Ketemu = True
            End If
            'Employee Data
            If Me.ChkEmployee.Checked = True Then
                Primary = -1
                Ketemu = False
                For ij As Integer = 0 To Me.dgvEmployee.Rows.Count - 2
                    RowDGV = Me.dgvEmployee.Rows(ij)
                    Primary = CDbl(RowDGV.Cells(Me.ColPrimaryCode.Index).Value)
                    If Primary = RowSalary.Primary_Employee Then
                        Ketemu = True
                        Exit For
                    End If
                Next
                If Ketemu = False Then
                    Return False
                End If
            Else
                'All data
                Ketemu = True
            End If
            'Job Position
            If Me.ChkPosition.Checked = True Then
                Primary = -1
                Ketemu = False
                For Each RowView In Me.ClbPositon.CheckedItems
                    Primary = RowView.Row.Item("primary")
                    If Primary = RowSalary.primary_position Then
                        Ketemu = True
                        Exit For
                    End If
                Next
                If Ketemu = False Then
                    Return False
                End If
            Else
                'All data
                Ketemu = True
            End If
            'Salary Formula
            If Me.ChkFormula.Checked = True Then
                Primary = -1
                Ketemu = False
                PrimaryFormula = CDbl(RowSalary.Primary_Salary_Formula)
                For Each RowView In Me.ClbSalaryFormula.CheckedItems
                    Primary = RowView.Row.Item("primary")
                    If Primary = PrimaryFormula Then
                        Ketemu = True
                        Exit For
                    End If
                Next
                If Ketemu = False Then
                    Return False
                End If
            Else
                'All data
                Ketemu = True
            End If

            'Result
            If Ketemu = True Then
                Return True
            ElseIf Ketemu = False Then
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BtnFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnFilter.Click
        Try
            Me.CetaklaporanIncomeTax()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CetaklaporanIncomeTax()
        Dim Report As New CrpLaporanIncomeTax_SMESS
        Dim DS As New DseTrans_Tax_Beginning_Balance
        Dim TBLCetak As DseTrans_Tax_Beginning_Balance.TblCetakTaxDataTable = DS.TblCetakTax
        Dim RowCetak As DseTrans_Tax_Beginning_Balance.TblCetakTaxRow
        Dim DTA As New DseLaporanIncomeTax_SMESSTableAdapters.DataTable1TableAdapter
        Dim TBL As New DseLaporanIncomeTax_SMESS.DataTable1DataTable
        Dim Row As DseLaporanIncomeTax_SMESS.DataTable1Row
        Dim StartDate As Date = Me.ObjCD.GetFirstDateofTheMonth(Me.DtpStart.Value.Date)
        Dim EndDate As Date = Me.ObjCD.GetLasDateOfTheMonth(Me.DtpStart.Value.Date)
        Dim Idx As Double = -1
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimarySalary(TBL, EndDate, StartDate)
            If TBL.Rows.Count > 0 Then
                For Each Row In TBL.Rows
                    If Me.CekFilterData(Row) = True Then
                        Idx += 1
                        RowCetak = TBLCetak.NewRow
                        RowCetak.Primary = Idx
                        RowCetak.Document_No = Row.No_Transaksi
                        RowCetak.Tanggal_Transaksi = #1/1/1950#
                        RowCetak.Period = #1/1/1950#
                        RowCetak.Primary_Karyawan = Row.Primary_Employee
                        RowCetak.Full_Name = Row.Employee_Name
                        RowCetak.Category = Row.A_Status_NPWP
                        RowCetak.No_NPWP = Row.A_NPWP
                        RowCetak.Gaji_Pensiun = Row._01Gaji
                        RowCetak.Tunjangan_PPh = Row._02TunjanganPPh
                        RowCetak.Tunjangan_Lainnya = Row._03TunjanganLainnya
                        RowCetak.Honorarium = Row._04Honorarium
                        RowCetak.Premi_Asuransi = Row._05PremiAsuransi
                        RowCetak.Penerimaan_Natura = Row._06Natura
                        RowCetak.Tantiem_Bonus = Row._07Tantiem
                        RowCetak.Jumlah_Penghasilan_Bruto = Row._08PenghasilanBruto
                        RowCetak.Biaya_Jabatan = Row._09BiayaJabatan
                        RowCetak.Iuran_Pensiun = Row._10IuranPensiun
                        RowCetak.Jumlah_Pengurangan = Row._11JumlahPengurangan
                        RowCetak.Jumlah_Penghasilan_Netto = Row._12JumlahPenghasilanNetto
                        RowCetak.Penghasilan_Netto_Masa_Pjk_Sblm = Row._13JumlahPenghasilanMasaPjkSblmnya
                        RowCetak.Perhitungan_PPh_Pasal_21_Disetahunkan = Row._14JumlahPenghasilanNettoUtkPPh
                        RowCetak.PTKP = Row._15PTKP
                        RowCetak.PKP = Row._16PKPSetahun
                        RowCetak.PPh_Ps_21_Atas_PKP_Disetahunkan = Row._17PPh21
                        RowCetak.PPh_Ps_21_Dipotong_Ms_Pjk_Sblmnya = Row._18PPh21YgTelahDipotongMasaPjkSblmnya
                        RowCetak.PPh_Ps_21_Terhutang = Row._19PPhTerhutang
                        RowCetak.PPh_Ps_21_Yg_Lunas = Row._20PPhPs21DanPPhPs26YgTelahDipotongNLunas
                        RowCetak.Remarks = ""
                        RowCetak.npwp_category = Row.A_Status_NPWP
                        RowCetak.Dept_Desc = Row.Dept_Desc
                        TBLCetak.Rows.Add(RowCetak)
                    End If
                Next
            End If

            Report.SetDataSource(DS)
            Me.crvTaxIncome.ReportSource = Report
            Me.crvTaxIncome.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub
End Class