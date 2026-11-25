Option Explicit On
Option Strict Off

Imports Core_Module, System.Windows.Forms, System.Drawing

Public Class FrmLaporanIncomeTaxStaff
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
            Me.cboFormat.SelectedIndex = 0
            Me.cboFormat.Anchor = Me.DtpStart.Anchor
            Me.Text = " Income Tax Salary"
            Me.dgvEmployee.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            '----------------------

            '----------------------
            Me.dgvEmployee.AllowUserToAddRows = False
            Me.IsiListBoxJobPosition(Me.ClbPositon, Me.TBL_JP)
            Me.IsiListBoxDepartemen(Me.ClbDepartemen, Me.TBL_Dept)
            Me.IsiListBoxFormula(Me.ClbSalaryFormula, Me.TBL_Formula)
            Me.KendalikanKontrol(EnumLaporan.AllData, False)
            Me.DtpStart.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.DtpStart.Value)
            Me.dtpEnd.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpEnd.Value)
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
            'jika ada yang filter, maka minimal harus ada yang dipilih
            If Me.ChkEmployee.Checked = True Then
                If Me.dgvEmployee.RowCount <= 1 Then
                    MsgBox("Please specify at least 1 employee.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.ChkFormula.Checked = True Then
                If Me.ClbSalaryFormula.CheckedItems.Count = 0 Then
                    MsgBox("Please specify at least 1 salary formula.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.ChkPosition.Checked = True Then
                If Me.ClbPositon.CheckedItems.Count = 0 Then
                    MsgBox("Please specify at least 1 job position.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.ChkDepartemen.Checked = True Then
                If Me.ClbDepartemen.CheckedItems.Count = 0 Then
                    MsgBox("Please specify at least 1 department.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If Me.cboFormat.SelectedIndex = 0 Then
                Me.CetakLaporanIncomeTaxRecap()
            Else
                'start period <= end period
                If Me.dtpEnd.Value.Date < Me.DtpStart.Value.Date Then
                    MsgBox("Please make sure start period is smaller than end period.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                Me.CetaklaporanIncomeTax()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CetakLaporanIncomeTaxRecap()
        Dim ij As New FrmLaporanIncomeTaxStaff_Recap
        Dim vTBLSalary As New DseMonthlyTax.DataTable1DataTable

        Try
            ij.RefPeriod = Me.DtpStart.Value
            ij.vTblSalary = vTBLSalary
            Me.IsiTabelRecap(vTBLSalary)

            If vTBLSalary.Rows.Count <> 0 Then
                ij.ShowDialog()
            Else
                MsgBox("No data found", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
            vTBLSalary.Dispose()
        End Try
    End Sub

    Private Sub IsiTabelRecap(ByVal vTBLSalary As DseMonthlyTax.DataTable1DataTable)
        Dim vDTASalary As New DseMonthlyTaxTableAdapters.DataTable1TableAdapter
        Dim vRowSalary As DseMonthlyTax.DataTable1Row
        Dim ListPrimary As New List(Of Double)
        Dim rowDGV As DataGridViewRow
        Dim rowSalaryFormula As DseLaporanSalaryStaff_SMESS.tblmaster_salary_formula1Row
        Dim rowSalaryJobPosition As DseLaporanSalaryStaff_SMESS.tblmaster_jobpositionRow
        Dim rowDepartment As DseLaporanSalaryStaff_SMESS.tblmaster_departemenRow

        Try
            vTBLSalary.Rows.Clear()
            vDTASalary.Connection.ConnectionString = Core_Module.GetConstring

            'filter data
            vDTASalary.Fill(vTBLSalary, Me.ObjCD.GetLasDateOfTheMonth(Me.DtpStart.Value.Date), Me.ObjCD.GetFirstDateofTheMonth(Me.DtpStart.Value.Date))

            'filter
            If Me.ChkDepartemen.Checked = True Or _
            Me.ChkEmployee.Checked = True Or _
            Me.ChkFormula.Checked = True Or _
            Me.ChkPosition.Checked = True Then
                'buang yang tidak cocok primary salary formula
                If Me.ChkFormula.Checked = True Then
                    ListPrimary.Clear()
                    For Each DV As DataRowView In Me.ClbSalaryFormula.CheckedItems
                        rowSalaryFormula = DV.Row
                        If ListPrimary.Contains(rowSalaryFormula.primary) = False Then
                            ListPrimary.Add(rowSalaryFormula.primary)
                        End If
                    Next
                    For i As Double = vTBLSalary.Rows.Count - 1 To 0 Step -1
                        vRowSalary = vTBLSalary.Rows(i)
                        If ListPrimary.Contains(vRowSalary.Primary_Salary_Formula) = False Then 
                            vTBLSalary.Rows.RemoveAt(i)
                        End If
                    Next
                End If

                'buang yang tidak cocok dgv employee
                If Me.ChkEmployee.Checked = True Then
                    ListPrimary.Clear()
                    For Each rowDGV In Me.dgvEmployee.Rows
                        If rowDGV.Index <> Me.dgvEmployee.RowCount - 1 Then
                            If ListPrimary.Contains(rowDGV.Cells(Me.ColPrimaryCode.Index).Value) = False Then
                                ListPrimary.Add(rowDGV.Cells(Me.ColPrimaryCode.Index).Value)
                            End If
                        End If
                    Next
                    For i As Double = vTBLSalary.Rows.Count - 1 To 0 Step -1
                        vRowSalary = vTBLSalary.Rows(i)
                        If ListPrimary.Contains(vRowSalary.Primary_Employee) = False Then
                            vTBLSalary.Rows.RemoveAt(i)
                        End If
                    Next
                End If

                'buang yang tidak cocok primary job position
                If Me.ChkPosition.Checked = True Then
                    ListPrimary.Clear()
                    For Each DV As DataRowView In Me.ClbPositon.CheckedItems
                        rowSalaryJobPosition = DV.Row
                        If ListPrimary.Contains(rowSalaryJobPosition.primary) = False Then
                            ListPrimary.Add(rowSalaryJobPosition.primary)
                        End If
                    Next
                    For i As Double = vTBLSalary.Rows.Count - 1 To 0 Step -1
                        vRowSalary = vTBLSalary.Rows(i)
                        If ListPrimary.Contains(vRowSalary.Primary_Position) = False Then
                            vTBLSalary.Rows.RemoveAt(i)
                        End If
                    Next
                End If

                'buang yang tidak cocok primary dept
                If Me.ChkDepartemen.Checked = True Then
                    ListPrimary.Clear()
                    For Each DV As DataRowView In Me.ClbDepartemen.CheckedItems
                        rowDepartment = DV.Row
                        If ListPrimary.Contains(rowDepartment.Primary) = False Then
                            ListPrimary.Add(rowDepartment.Primary)
                        End If
                    Next
                    For i As Double = vTBLSalary.Rows.Count - 1 To 0 Step -1
                        vRowSalary = vTBLSalary.Rows(i)
                        If ListPrimary.Contains(vRowSalary.primary_departemen) = False Then
                            vTBLSalary.Rows.RemoveAt(i)
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            vDTASalary.Dispose() 
            ListPrimary = Nothing
        End Try
    End Sub

    Private Sub IsiTabel(ByVal vTBLSalary As DseTransaksiSalary.tbltrans_salary3DataTable, ByVal vTblItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxDataTable)
        Dim vDTASalary As New DseTransaksiSalaryTableAdapters.tbltrans_salary3TableAdapter
        Dim vRowSalary As DseTransaksiSalary.tbltrans_salary3Row
        Dim vDTAItemTax As New DseTransaksiSalaryTableAdapters.tbltrans_salary_itemtaxTableAdapter
        Dim ListPrimary As New List(Of Double)
        Dim rowDGV As DataGridViewRow
        Dim rowSalaryFormula As DseLaporanSalaryStaff_SMESS.tblmaster_salary_formula1Row
        Dim rowSalaryJobPosition As DseLaporanSalaryStaff_SMESS.tblmaster_jobpositionRow
        Dim rowDepartment As DseLaporanSalaryStaff_SMESS.tblmaster_departemenRow

        Try
            vTBLSalary.Rows.Clear()
            vTblItemTax.Rows.Clear()
            vDTASalary.Connection.ConnectionString = Core_Module.GetConstring
            vDTAItemTax.Connection.ConnectionString = vDTASalary.Connection.ConnectionString

            'filter data
            vDTASalary.Fill(vTBLSalary, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpEnd.Value.Date), Me.ObjCD.GetFirstDateofTheMonth(Me.DtpStart.Value.Date))
            vDTAItemTax.FillBy_TglSampaiDari(vTblItemTax, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpEnd.Value.Date), Me.ObjCD.GetFirstDateofTheMonth(Me.DtpStart.Value.Date))

            'filter
            If Me.ChkDepartemen.Checked = True Or _
            Me.ChkEmployee.Checked = True Or _
            Me.ChkFormula.Checked = True Or _
            Me.ChkPosition.Checked = True Then
                'buang yang tidak cocok primary salary formula
                If Me.ChkFormula.Checked = True Then
                    ListPrimary.Clear()
                    For Each DV As DataRowView In Me.ClbSalaryFormula.CheckedItems
                        rowSalaryFormula = DV.Row
                        If ListPrimary.Contains(rowSalaryFormula.primary) = False Then
                            ListPrimary.Add(rowSalaryFormula.primary)
                        End If
                    Next
                    For i As Double = vTBLSalary.Rows.Count - 1 To 0 Step -1
                        vRowSalary = vTBLSalary.Rows(i)
                        If ListPrimary.Contains(vRowSalary.Primary_Salary_Formula) = False Then
                            Me.RemoveByPrimaryTransaksi(vRowSalary.Primary, vTblItemTax)
                            vTBLSalary.Rows.RemoveAt(i)
                        End If
                    Next
                End If

                'buang yang tidak cocok dgv employee
                If Me.ChkEmployee.Checked = True Then
                    ListPrimary.Clear()
                    For Each rowDGV In Me.dgvEmployee.Rows
                        If rowDGV.Index <> Me.dgvEmployee.RowCount - 1 Then
                            If ListPrimary.Contains(rowDGV.Cells(Me.ColPrimaryCode.Index).Value) = False Then
                                ListPrimary.Add(rowDGV.Cells(Me.ColPrimaryCode.Index).Value)
                            End If
                        End If
                    Next
                    For i As Double = vTBLSalary.Rows.Count - 1 To 0 Step -1
                        vRowSalary = vTBLSalary.Rows(i)
                        If ListPrimary.Contains(vRowSalary.Primary_Employee) = False Then
                            Me.RemoveByPrimaryTransaksi(vRowSalary.Primary, vTblItemTax)
                            vTBLSalary.Rows.RemoveAt(i)
                        End If
                    Next
                End If

                'buang yang tidak cocok primary job position
                If Me.ChkPosition.Checked = True Then
                    ListPrimary.Clear()
                    For Each DV As DataRowView In Me.ClbPositon.CheckedItems
                        rowSalaryJobPosition = DV.Row
                        If ListPrimary.Contains(rowSalaryJobPosition.primary) = False Then
                            ListPrimary.Add(rowSalaryJobPosition.primary)
                        End If
                    Next
                    For i As Double = vTBLSalary.Rows.Count - 1 To 0 Step -1
                        vRowSalary = vTBLSalary.Rows(i)
                        If ListPrimary.Contains(vRowSalary.Primary_Position) = False Then
                            Me.RemoveByPrimaryTransaksi(vRowSalary.Primary, vTblItemTax)
                            vTBLSalary.Rows.RemoveAt(i)
                        End If
                    Next
                End If

                'buang yang tidak cocok primary dept
                If Me.ChkDepartemen.Checked = True Then
                    ListPrimary.Clear()
                    For Each DV As DataRowView In Me.ClbDepartemen.CheckedItems
                        rowDepartment = DV.Row
                        If ListPrimary.Contains(rowDepartment.Primary) = False Then
                            ListPrimary.Add(rowDepartment.Primary)
                        End If
                    Next
                    For i As Double = vTBLSalary.Rows.Count - 1 To 0 Step -1
                        vRowSalary = vTBLSalary.Rows(i)
                        If ListPrimary.Contains(vRowSalary.primary_departemen) = False Then
                            Me.RemoveByPrimaryTransaksi(vRowSalary.Primary, vTblItemTax)
                            vTBLSalary.Rows.RemoveAt(i)
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            vDTASalary.Dispose()
            vDTAItemTax.Dispose()
            ListPrimary = Nothing
        End Try
    End Sub

    Private Sub CetaklaporanIncomeTax()
        Dim k As New FrmTrans_Gaji_IncomeTax2
        Dim vTBLSalary As New DseTransaksiSalary.tbltrans_salary3DataTable
        Dim vTblItemTax As New DseTransaksiSalary.tbltrans_salary_itemtaxDataTable

        Try
            k.vTblSalary = vTBLSalary
            k.vTBLItemTax = vTblItemTax
            k.vFromMonthlyReport = True

            Me.IsiTabel(vTBLSalary, vTblItemTax)

            If vTBLSalary.Rows.Count = 0 Then
                MsgBox("There is no salary calculation data in the specified period.", MsgBoxStyle.Information, "Information")
                Exit Sub
            Else
                k.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            vTBLSalary.Dispose()
            vTblItemTax.Dispose()
        End Try
    End Sub

    Private Sub RemoveByPrimaryTransaksi(ByVal vPrimaryTransaksi As Double, ByVal vTblItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxDataTable)
        Try
            vTblItemTax.DefaultView.RowFilter = "[primary transaksi] = " + vPrimaryTransaksi.ToString
            While vTblItemTax.DefaultView.Count > 0
                vTblItemTax.Rows.Remove(vTblItemTax.DefaultView(0).Row)
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboFormat_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFormat.SelectedIndexChanged
        Try
            If Me.cboFormat.SelectedIndex = 0 Then
                Me.dtpEnd.Enabled = False
            Else
                Me.dtpEnd.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class