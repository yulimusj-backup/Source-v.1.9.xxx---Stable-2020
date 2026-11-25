Option Explicit On
Option Strict Off

Imports Core_Module, System.Windows.Forms, System.Drawing
Public Class FrmLaporanSalaryStaff_SMESS
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
            Me.dgvEmployee.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            '--------------------
            Me.crvRekap.DisplayBackgroundEdge = True
            Me.crvRekap.DisplayStatusBar = True
            Me.crvRekap.DisplayToolbar = True
            Me.crvRekap.ShowExportButton = True
            Me.crvRekap.ShowGotoPageButton = True
            Me.crvRekap.ShowPageNavigateButtons = True
            Me.crvRekap.ShowPrintButton = True
            Me.crvRekap.ShowTextSearchButton = True
            Me.crvRekap.ShowZoomButton = True
            '---------------------
            Me.crvDetails.DisplayBackgroundEdge = True
            Me.crvDetails.DisplayStatusBar = True
            Me.crvDetails.DisplayToolbar = True
            Me.crvDetails.ShowExportButton = True
            Me.crvDetails.ShowGotoPageButton = True
            Me.crvDetails.ShowPageNavigateButtons = True
            Me.crvDetails.ShowPrintButton = True
            Me.crvDetails.ShowTextSearchButton = True
            Me.crvDetails.ShowZoomButton = True
            '----------------------
            Me.dgvEmployee.AllowUserToAddRows = False
            Me.IsiListBoxJobPosition(Me.ClbPositon, Me.TBL_JP)
            Me.IsiListBoxDepartemen(Me.ClbDepartemen, Me.TBL_Dept)
            Me.IsiListBoxFormula(Me.ClbSalaryFormula, Me.TBL_Formula)
            Me.KendalikanKontrol(EnumLaporan.AllData, False)
            Me.AllData = True
            Me.DtpTo.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.DtpTo.Value)
            Me.DtpFrom.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.DtpFrom.Value)
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
    Private Function CekFilterData(ByVal RowSalary As DseLaporanSalaryStaff_SMESS.DataTable1Row) As Boolean
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
            If Me.ChkFormula.Checked = True And RowSalary.Identifier_Salary_Component.Contains("_") Then
                Primary = -1
                Ketemu = False
                PrimaryFormula = CDbl(RowSalary.Identifier_Salary_Component.Substring(0, RowSalary.Identifier_Salary_Component.IndexOf("_")))
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

    'Insert Data Print
    Public Function ApakahBolehPrint() As Boolean
        Try
            If Me.DtpTo.Value.Date < Me.DtpFrom.Value.Date Then
                MsgBox("Period to cannot be less then period from.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BtnFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnFilter.Click
        Dim i As New FrmLaporanSalaryStaff_Cetak_IGI
        Try
            If Me.ApakahBolehPrint() = True Then
                Me.CetaklaporanRekapDetail()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub CetaklaporanRekapDetail()
        Dim ReportRekap As New CrpLaporanStaffSalary_Rekap_SMESS
        Dim ReportDetail As New CrpLaporanStaffSalary_Detail_SMESS
        Dim DateFirst As Date = Me.ObjCD.GetFirstDateofTheMonth(Me.DtpFrom.Value.Date)
        Dim DateLast As Date = Me.ObjCD.GetLasDateOfTheMonth(Me.DtpTo.Value.Date)
        Dim ListEmployee As New List(Of Double)
        Dim RV As DataRowView
        Dim FormatDecimal As Double = Me.DefPrec
        Dim Filter As String = ""
        Dim Filter2 As String = ""
        Dim GroupPertamaStr As String = ""
        Dim DS As New DseLaporanSalaryStaff_SMESS
        Dim TBLCetak As DseLaporanSalaryStaff_SMESS.TblCetakDataTable = DS.TblCetak
        Dim RowCetak As DseLaporanSalaryStaff_SMESS.TblCetakRow
        '--Income, deduction--
        Dim DTA_A0 As New DseLaporanSalaryStaff_SMESSTableAdapters.DataTable1TableAdapter
        Dim TBL_A0 As New DseLaporanSalaryStaff_SMESS.DataTable1DataTable
        Dim TBL_A0_ As New DseLaporanSalaryStaff_SMESS.DataTable1DataTable
        Dim Row_A0 As DseLaporanSalaryStaff_SMESS.DataTable1Row
        Dim Row_A0_ As DseLaporanSalaryStaff_SMESS.DataTable1Row
        Dim DTA_A1 As New DseLaporanSalaryStaff_SMESSTableAdapters.DataTable6TableAdapter
        Dim TBL_A1 As New DseLaporanSalaryStaff_SMESS.DataTable6DataTable
        Dim Row_A1 As DseLaporanSalaryStaff_SMESS.DataTable6Row
        Dim DTA_A2 As New DseLaporanSalaryStaff_SMESSTableAdapters.tbltrans_salary_itemTableAdapter
        Dim TBL_A2 As New DseLaporanSalaryStaff_SMESS.tbltrans_salary_itemDataTable
        Dim Row_A2 As DseLaporanSalaryStaff_SMESS.tbltrans_salary_itemRow
        '--Other Deduction--
        Dim DTA_B0 As New DseLaporanSalaryStaff_SMESSTableAdapters.tbltrans_salary_item_otherdeductionTableAdapter
        Dim TBL_B0 As New DseLaporanSalaryStaff_SMESS.tbltrans_salary_item_otherdeductionDataTable
        Dim Row_B0 As DseLaporanSalaryStaff_SMESS.tbltrans_salary_item_otherdeductionRow
        Dim DTA_B1 As New DseLaporanSalaryStaff_SMESSTableAdapters.DataTable7TableAdapter
        Dim TBL_B1 As New DseLaporanSalaryStaff_SMESS.DataTable7DataTable
        Dim Row_B1 As DseLaporanSalaryStaff_SMESS.DataTable7Row
        Dim DTA_B2 As New DseLaporanSalaryStaff_SMESSTableAdapters.tbltrans_salary_item_otherdeduction1TableAdapter
        Dim TBL_B2 As New DseLaporanSalaryStaff_SMESS.tbltrans_salary_item_otherdeduction1DataTable
        Dim Row_B2 As DseLaporanSalaryStaff_SMESS.tbltrans_salary_item_otherdeduction1Row
        '--Payment--
        Dim DTA_C0 As New DseLaporanSalaryStaff_SMESSTableAdapters.DataTable2TableAdapter
        Dim TBL_C0 As New DseLaporanSalaryStaff_SMESS.DataTable2DataTable
        Dim Row_C0 As DseLaporanSalaryStaff_SMESS.DataTable2Row
        Dim DTA_C1 As New DseLaporanSalaryStaff_SMESSTableAdapters.DataTable8TableAdapter
        Dim TBL_C1 As New DseLaporanSalaryStaff_SMESS.DataTable8DataTable
        Dim Row_C1 As DseLaporanSalaryStaff_SMESS.DataTable8Row
        Dim DTA_C2 As New DseLaporanSalaryStaff_SMESSTableAdapters.tbltrans_salary_employee_d_igiTableAdapter
        Dim TBL_C2 As New DseLaporanSalaryStaff_SMESS.tbltrans_salary_employee_d_igiDataTable
        Dim Row_C2 As DseLaporanSalaryStaff_SMESS.tbltrans_salary_employee_d_igiRow
        'loan payment, compan expense, salary expense
        Dim DTA1 As New DseLaporanSalaryStaff_SMESSTableAdapters.tbltrans_salaryTableAdapter
        Dim TBL1 As New DseLaporanSalaryStaff_SMESS.tbltrans_salaryDataTable
        Dim row1 As DseLaporanSalaryStaff_SMESS.tbltrans_salaryRow

        Try
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            '-------------
            DTA_A0.Connection.ConnectionString = Core_Module.GetConstring
            DTA_A1.Connection.ConnectionString = Core_Module.GetConstring
            DTA_A2.Connection.ConnectionString = Core_Module.GetConstring
            DTA_B0.Connection.ConnectionString = Core_Module.GetConstring
            DTA_B1.Connection.ConnectionString = Core_Module.GetConstring
            DTA_B2.Connection.ConnectionString = Core_Module.GetConstring
            DTA_C0.Connection.ConnectionString = Core_Module.GetConstring
            DTA_C1.Connection.ConnectionString = Core_Module.GetConstring
            DTA_C2.Connection.ConnectionString = Core_Module.GetConstring
            '-------------
            DTA_A0.FillBy_TanggalDariSampai(TBL_A0, DateLast, DateFirst)
            TBL_A0_ = TBL_A0.Copy
            DTA_A1.FillBy_TanggalDariSampai(TBL_A1, DateLast, DateFirst)
            DTA_A2.FillBy_TanggalDariSampai(TBL_A2, DateLast, DateFirst)
            DTA_B0.FillBy_TanggalDariSampai(TBL_B0, DateLast, DateFirst)
            DTA_B1.FillBy_TanggalDariSampai(TBL_B1, DateLast, DateFirst)
            DTA_B2.FillBy_TanggalDariSampai(TBL_B2, DateLast, DateFirst)
            DTA_C0.FillBy_TanggalDariSampai(TBL_C0, DateLast, DateFirst)
            DTA_C1.FillBy_TanggalDariSampai(TBL_C1, DateLast, DateFirst)
            DTA_C2.FillBy_TanggalDariSampai(TBL_C2, DateLast, DateFirst)
            '--------
            Dim ApakahPakaiFilterKiri As Boolean = False
            If Me.ChkDepartemen.Checked = True Then
                ApakahPakaiFilterKiri = True
            Else
                If Me.ChkEmployee.Checked = True Then
                    ApakahPakaiFilterKiri = True
                Else
                    If Me.ChkPosition.Checked = True Then
                        ApakahPakaiFilterKiri = True
                    Else
                        If Me.ChkFormula.Checked = True Then
                            ApakahPakaiFilterKiri = True
                        End If
                    End If
                End If
            End If
            If ApakahPakaiFilterKiri = True Then
                TBL_A1.Rows.Clear()
                TBL_A2.Rows.Clear()
                TBL_B1.Rows.Clear()
                TBL_B2.Rows.Clear()
                TBL_C1.Rows.Clear()
                TBL_C2.Rows.Clear()
                TBL_A0.DefaultView.Sort = "[primary departemen]"
                TBL_B0.DefaultView.Sort = "[primary departemen]"
                TBL_C0.DefaultView.Sort = "[primary departemen]"
                For Each RV In TBL_A0.DefaultView
                    Row_A0 = RV.Row

                    If Row_A0.Identifier_Salary_Component.Length = 0 Then
                        TBL_A0_.DefaultView.RowFilter = "[Primary Employee] = " & Row_A0.Primary_Employee.ToString & " AND [Identifier Salary Component] <> ''"
                        If TBL_A0_.DefaultView.Count > 0 Then
                            Row_A0_ = TBL_A0_.DefaultView(0).Row
                            Row_A0.Identifier_Salary_Component = Row_A0_.Identifier_Salary_Component
                        End If
                    End If

                    If Me.CekFilterData(Row_A0) = True Then
                        'Income deduction
                        Filter = "[Identifier Salary Component] = '" + Row_A0.Identifier_Salary_Component.ToString + _
                        "' AND [primary departemen] = " + Row_A0.primary_departemen.ToString + ""
                        TBL_A1.DefaultView.RowFilter = Filter
                        If TBL_A1.DefaultView.Count > 0 Then
                            Row_A1 = TBL_A1.DefaultView(0).Row
                            Row_A1.TotalAmount_PerDept += Row_A0.Amount
                        Else
                            Row_A1 = TBL_A1.NewDataTable6Row
                            Row_A1.Identifier_Salary_Component = Row_A0.Identifier_Salary_Component
                            Row_A1.primary_departemen = Row_A0.primary_departemen
                            Row_A1.TotalAmount_PerDept = Row_A0.Amount
                            TBL_A1.Rows.Add(Row_A1)
                        End If
                        Filter = "[Identifier Salary Component] = '" + Row_A0.Identifier_Salary_Component.ToString + "'"
                        TBL_A2.DefaultView.RowFilter = Filter
                        If TBL_A2.DefaultView.Count > 0 Then
                            Row_A2 = TBL_A2.DefaultView(0).Row
                            Row_A2.TotalAmount_All += Row_A0.Amount
                        Else
                            Row_A2 = TBL_A2.Newtbltrans_salary_itemRow
                            Row_A2.Identifier_Salary_Component = Row_A0.Identifier_Salary_Component
                            Row_A2.TotalAmount_All = Row_A0.Amount
                            TBL_A2.Rows.Add(Row_A2)
                        End If
                        If ListEmployee.Contains(Row_A0.Primary_Employee) = False Then
                            ListEmployee.Add(Row_A0.Primary_Employee)
                        End If
                    End If
                Next
                For Each RV In TBL_B0.DefaultView
                    Row_B0 = RV.Row
                    If ListEmployee.Contains(Row_B0.Primary_Employee) = True Then
                        'Other Deduction
                        Filter = "[Primary Other Deduction] = " + Row_B0.Primary_Other_Deduction.ToString + _
                        " AND [primary departemen] = " + Row_B0.primary_departemen.ToString + ""
                        TBL_B1.DefaultView.RowFilter = Filter
                        If TBL_B1.DefaultView.Count > 0 Then
                            Row_B1 = TBL_B1.DefaultView(0).Row
                            Row_B1.TotalAmount_PerDept += Row_B0.Amount
                        Else
                            Row_B1 = TBL_B1.NewDataTable7Row
                            Row_B1.Primary_Other_Deduction = Row_B0.Primary_Other_Deduction
                            Row_B1.primary_departemen = Row_B0.primary_departemen
                            Row_B1.TotalAmount_PerDept = Row_B0.Amount
                            TBL_B1.Rows.Add(Row_B1)
                        End If
                        Filter = "[Primary Other Deduction] = " + Row_B0.Primary_Other_Deduction.ToString + ""
                        TBL_B2.DefaultView.RowFilter = Filter
                        If TBL_B2.DefaultView.Count > 0 Then
                            Row_B2 = TBL_B2.DefaultView(0).Row
                            Row_B2.TotalAmount_PerDept += Row_B0.Amount
                        Else
                            Row_B2 = TBL_B2.Newtbltrans_salary_item_otherdeduction1Row
                            Row_B2.Primary_Other_Deduction = Row_B0.Primary_Other_Deduction
                            Row_B2.TotalAmount_PerDept = Row_B0.Amount
                            TBL_B2.Rows.Add(Row_B2)
                        End If
                    End If
                Next
                For Each RV In TBL_C0.DefaultView
                    Row_C0 = RV.Row
                    'Payment
                    If ListEmployee.Contains(Row_C0.Primary_Employee) = True Then
                        Filter = "[primary akun payment] = " + Row_C0.primary_akun_payment.ToString + _
                        " AND [primary departemen] = " + Row_C0.primary_departemen.ToString + ""
                        TBL_C1.DefaultView.RowFilter = Filter
                        If TBL_C1.DefaultView.Count > 0 Then
                            Row_C1 = TBL_C1.DefaultView(0).Row
                            Row_C1.TotalAmount_PerDept += Row_C0.Amount
                        Else
                            Row_C1 = TBL_C1.NewDataTable8Row
                            Row_C1.primary_akun_payment = Row_C0.primary_akun_payment
                            Row_C1.primary_departemen = Row_C0.primary_departemen
                            Row_C1.TotalAmount_PerDept = Row_C0.Amount
                            TBL_C1.Rows.Add(Row_C1)
                        End If
                        Filter = "[primary akun payment] = " + Row_C0.primary_akun_payment.ToString + ""
                        TBL_C2.DefaultView.RowFilter = Filter
                        If TBL_C2.DefaultView.Count > 0 Then
                            Row_C2 = TBL_C2.DefaultView(0).Row
                            Row_C2.TotalAmount_PerDept += Row_C0.Amount
                        Else
                            Row_C2 = TBL_C2.Newtbltrans_salary_employee_d_igiRow
                            Row_C2.primary_akun_payment = Row_C0.primary_akun_payment
                            Row_C2.TotalAmount_PerDept = Row_C0.Amount
                            TBL_C2.Rows.Add(Row_C2)
                        End If
                    End If
                Next
                TBL_A0.DefaultView.Sort = ""
                TBL_B0.DefaultView.Sort = ""
                TBL_C0.DefaultView.Sort = ""
            End If
            'Insert ke TBLCetak
            If TBL_A0.Rows.Count > 0 Then
                ListEmployee.Clear()
                For Each Row_A0 In TBL_A0.Rows
                    If Me.CekFilterData(Row_A0) = True Then
                        If ListEmployee.Contains(Row_A0.Primary_Employee) = False Then
                            ListEmployee.Add(Row_A0.Primary_Employee)
                            If ListEmployee.Count = 1 Then
                                GroupPertamaStr = Row_A0.Dept_Desc
                            End If
                            RowCetak = TBLCetak.NewRow
                            RowCetak.Department = Row_A0.Dept_Desc
                            RowCetak.Employee_Name = Row_A0.Nama
                            RowCetak.Primary_Dept = Row_A0.primary_departemen
                            RowCetak.Primary_Employee = Row_A0.Primary_Employee
                            RowCetak.Income = Row_A0.Income
                            RowCetak.Deduction = Row_A0.Deduction
                            RowCetak.Gross_Salary = Row_A0.Gross_Salary
                            RowCetak.Income_Tax = Row_A0.Income_Tax
                            RowCetak.Other_Deduction = Row_A0.Other_Deduction
                            RowCetak.Take_Home_Pay = Row_A0.Take_Home_Pay
                            '------Item-Item------
                            If Row_A0.Type = 0 Then
                                RowCetak.Detail_Income = Row_A0.Salary_Desc
                                RowCetak.Detail_Income_Amount = Me.ObjCD.DigitGrouping(Row_A0.Amount, Me.DefPrec)
                                Filter2 = "[Identifier Salary Component] = '" + Row_A0.Identifier_Salary_Component.ToString + _
                                "' AND [primary departemen] = " + Row_A0.primary_departemen.ToString + ""
                                TBL_A1.DefaultView.RowFilter = Filter2
                                If TBL_A1.DefaultView.Count > 0 Then
                                    Row_A1 = TBL_A1.DefaultView(0).Row
                                    RowCetak.TD_Income = Row_A0.Salary_Desc
                                    RowCetak.TD_Income_Amount = Me.ObjCD.DigitGrouping(Row_A1.TotalAmount_PerDept, Me.DefPrec)
                                End If
                                Filter2 = "[Identifier Salary Component] = '" + Row_A0.Identifier_Salary_Component.ToString + "'"
                                TBL_A2.DefaultView.RowFilter = Filter2
                                If TBL_A2.DefaultView.Count > 0 Then
                                    Row_A2 = TBL_A2.DefaultView(0).Row
                                    RowCetak.Grand_Income = Row_A0.Salary_Desc
                                    RowCetak.Grand_Income_Amount = Me.ObjCD.DigitGrouping(Row_A2.TotalAmount_All, Me.DefPrec)
                                End If
                            Else
                                RowCetak.Detail_Deduction = Row_A0.Salary_Desc
                                RowCetak.Detail_Deduction_Amount = Me.ObjCD.DigitGrouping(Row_A0.Amount, Me.DefPrec)
                                Filter2 = "[Identifier Salary Component] = '" + Row_A0.Identifier_Salary_Component.ToString + _
                                "' AND [primary departemen] = " + Row_A0.primary_departemen.ToString + ""
                                TBL_A1.DefaultView.RowFilter = Filter2
                                If TBL_A1.DefaultView.Count > 0 Then
                                    Row_A1 = TBL_A1.DefaultView(0).Row
                                    RowCetak.TD_Deduction = Row_A0.Salary_Desc
                                    RowCetak.TD_Deduction_Amount = Me.ObjCD.DigitGrouping(Row_A1.TotalAmount_PerDept, Me.DefPrec)
                                End If
                                Filter2 = "[Identifier Salary Component] = '" + Row_A0.Identifier_Salary_Component.ToString + "'"
                                TBL_A2.DefaultView.RowFilter = Filter2
                                If TBL_A2.DefaultView.Count > 0 Then
                                    Row_A2 = TBL_A2.DefaultView(0).Row
                                    RowCetak.Grand_Deduction = Row_A0.Salary_Desc
                                    RowCetak.Grand_Deduction_Amount = Me.ObjCD.DigitGrouping(Row_A2.TotalAmount_All, Me.DefPrec)
                                End If
                            End If
                            Filter = "[primary departemen] = " + Row_A0.primary_departemen.ToString _
                            + " AND [Primary Employee] = " + Row_A0.Primary_Employee.ToString + ""
                            TBL_B0.DefaultView.RowFilter = Filter
                            If TBL_B0.DefaultView.Count > 0 Then
                                For Each RV In TBL_B0.DefaultView
                                    Row_B0 = RV.Row
                                    If String.IsNullOrEmpty(RowCetak.Detail_Other) = True Then
                                        RowCetak.Detail_Other = Row_B0.Salary_Desc
                                        RowCetak.Detail_Other_Amount = Me.ObjCD.DigitGrouping(Row_B0.Amount, Me.DefPrec)
                                    Else
                                        RowCetak.Detail_Other += vbNewLine + Row_B0.Salary_Desc
                                        RowCetak.Detail_Other_Amount += vbNewLine + Me.ObjCD.DigitGrouping(Row_B0.Amount, Me.DefPrec)
                                    End If
                                    Filter2 = "[primary departemen] = " + Row_B0.primary_departemen.ToString + _
                                    " AND [Primary Other Deduction] = " + Row_B0.Primary_Other_Deduction.ToString + ""
                                    TBL_B1.DefaultView.RowFilter = Filter2
                                    If TBL_B1.DefaultView.Count > 0 Then
                                        Row_B1 = TBL_B1.DefaultView(0).Row
                                        If String.IsNullOrEmpty(RowCetak.TD_Other) = True Then
                                            RowCetak.TD_Other = Row_B0.Salary_Desc
                                            RowCetak.TD_Other_Amount = Me.ObjCD.DigitGrouping(Row_B1.TotalAmount_PerDept, Me.DefPrec)
                                        Else
                                            RowCetak.TD_Other += vbNewLine + Row_B0.Salary_Desc
                                            RowCetak.TD_Other_Amount += vbNewLine + Me.ObjCD.DigitGrouping(CDbl(Row_B1.TotalAmount_PerDept), Me.DefPrec)
                                        End If
                                    End If
                                    Filter2 = "[Primary Other Deduction] = " + Row_B0.Primary_Other_Deduction.ToString + ""
                                    TBL_B2.DefaultView.RowFilter = Filter2
                                    If TBL_B2.DefaultView.Count > 0 Then
                                        Row_B2 = TBL_B2.DefaultView(0).Row
                                        If String.IsNullOrEmpty(RowCetak.Grand_Other) = True Then
                                            RowCetak.Grand_Other = Row_B0.Salary_Desc
                                            RowCetak.Grand_Other_Amount = Me.ObjCD.DigitGrouping(Row_B2.TotalAmount_PerDept, Me.DefPrec)
                                        Else
                                            RowCetak.Grand_Other += vbNewLine + Row_B0.Salary_Desc
                                            RowCetak.Grand_Other_Amount += vbNewLine + Me.ObjCD.DigitGrouping(CDbl(Row_B2.TotalAmount_PerDept), Me.DefPrec)
                                        End If
                                    End If
                                Next
                            End If
                            TBL_C0.DefaultView.RowFilter = Filter
                            If TBL_C0.DefaultView.Count > 0 Then
                                For Each RV In TBL_C0.DefaultView
                                    Row_C0 = RV.Row
                                    If String.IsNullOrEmpty(RowCetak.Detail_Payment) = True Then
                                        RowCetak.Detail_Payment = Row_C0.Description
                                        RowCetak.Detail_Payment_Amount = Me.ObjCD.DigitGrouping(Row_C0.Amount, Me.DefPrec)
                                    Else
                                        RowCetak.Detail_Payment += vbNewLine + Row_C0.Description
                                        RowCetak.Detail_Payment_Amount += vbNewLine + Me.ObjCD.DigitGrouping(CDbl(Row_C0.Amount), Me.DefPrec)
                                    End If
                                    Filter2 = "[primary departemen] = " + Row_C0.primary_departemen.ToString + _
                                    " AND [primary akun payment] = " + Row_C0.primary_akun_payment.ToString + ""
                                    TBL_C1.DefaultView.RowFilter = Filter2
                                    If TBL_C1.DefaultView.Count > 0 Then
                                        Row_C1 = TBL_C1.DefaultView(0).Row
                                        If String.IsNullOrEmpty(RowCetak.TD_Payment) = True Then
                                            RowCetak.TD_Payment = Row_C0.Description
                                            RowCetak.TD_Payment_Amount = Me.ObjCD.DigitGrouping(Row_C1.TotalAmount_PerDept, Me.DefPrec)
                                        Else
                                            RowCetak.TD_Payment += vbNewLine + Row_C0.Description
                                            RowCetak.TD_Payment_Amount += vbNewLine + Me.ObjCD.DigitGrouping(CDbl(Row_C1.TotalAmount_PerDept), Me.DefPrec)
                                        End If
                                    End If
                                    Filter2 = "[Primary akun payment] = " + Row_C0.primary_akun_payment.ToString + ""
                                    TBL_C2.DefaultView.RowFilter = Filter2
                                    If TBL_C2.DefaultView.Count > 0 Then
                                        Row_C2 = TBL_C2.DefaultView(0).Row
                                        If String.IsNullOrEmpty(RowCetak.Grand_Payment) = True Then
                                            RowCetak.Grand_Payment = Row_C0.Description
                                            RowCetak.Grand_Payment_Amount = Me.ObjCD.DigitGrouping(Row_C2.TotalAmount_PerDept, Me.DefPrec)
                                        Else
                                            RowCetak.Grand_Payment += vbNewLine + Row_C0.Description
                                            RowCetak.Grand_Payment_Amount = vbNewLine + Me.ObjCD.DigitGrouping(CDbl(Row_C2.TotalAmount_PerDept), Me.DefPrec)
                                        End If
                                    End If
                                Next
                            End If
                            TBLCetak.Rows.Add(RowCetak)
                        Else
                            '------Item-Item------
                            RowCetak = TBLCetak.Rows(TBLCetak.Rows.Count - 1)
                            If Row_A0.Type = 0 Then
                                RowCetak.Detail_Income += vbNewLine + Row_A0.Salary_Desc
                                RowCetak.Detail_Income_Amount += vbNewLine + Me.ObjCD.DigitGrouping(Row_A0.Amount, Me.DefPrec)
                                Filter2 = "[Identifier Salary Component] = '" + Row_A0.Identifier_Salary_Component.ToString + _
                                "' AND [primary departemen] = " + Row_A0.primary_departemen.ToString + ""
                                TBL_A1.DefaultView.RowFilter = Filter2
                                If TBL_A1.DefaultView.Count > 0 Then
                                    Row_A1 = TBL_A1.DefaultView(0).Row
                                    RowCetak.TD_Income += vbNewLine + Row_A0.Salary_Desc
                                    RowCetak.TD_Income_Amount += vbNewLine + Me.ObjCD.DigitGrouping(Row_A1.TotalAmount_PerDept, Me.DefPrec)
                                End If
                                Filter2 = "[Identifier Salary Component] = '" + Row_A0.Identifier_Salary_Component.ToString + "'"
                                TBL_A2.DefaultView.RowFilter = Filter2
                                If TBL_A2.DefaultView.Count > 0 Then
                                    Row_A2 = TBL_A2.DefaultView(0).Row
                                    RowCetak.Grand_Income += vbNewLine + Row_A0.Salary_Desc
                                    RowCetak.Grand_Income_Amount += vbNewLine + Me.ObjCD.DigitGrouping(Row_A2.TotalAmount_All, Me.DefPrec)
                                End If
                            Else
                                RowCetak.Detail_Deduction += vbNewLine + Row_A0.Salary_Desc
                                RowCetak.Detail_Deduction_Amount += vbNewLine + Me.ObjCD.DigitGrouping(Row_A0.Amount, Me.DefPrec)
                                Filter2 = "[Identifier Salary Component] = '" + Row_A0.Identifier_Salary_Component.ToString + _
                                "' AND [primary departemen] = " + Row_A0.primary_departemen.ToString + ""
                                TBL_A1.DefaultView.RowFilter = Filter2
                                If TBL_A1.DefaultView.Count > 0 Then
                                    Row_A1 = TBL_A1.DefaultView(0).Row
                                    RowCetak.TD_Deduction += vbNewLine + Row_A0.Salary_Desc
                                    RowCetak.TD_Deduction_Amount += vbNewLine + Me.ObjCD.DigitGrouping(Row_A1.TotalAmount_PerDept, Me.DefPrec)
                                End If
                                Filter2 = "[Identifier Salary Component] = '" + Row_A0.Identifier_Salary_Component.ToString + "'"
                                TBL_A2.DefaultView.RowFilter = Filter2
                                If TBL_A2.DefaultView.Count > 0 Then
                                    Row_A2 = TBL_A2.DefaultView(0).Row
                                    RowCetak.Grand_Deduction += vbNewLine + Row_A0.Salary_Desc
                                    RowCetak.Grand_Deduction_Amount += vbNewLine + Me.ObjCD.DigitGrouping(Row_A2.TotalAmount_All, Me.DefPrec)
                                End If
                            End If
                        End If
                        '-------------
                    End If
                Next

                'isikan data load payment, company expense, dan employee expense
                DTA1.Fill(TBL1, Me.ObjCD.GetLasDateOfTheMonth(Me.DtpTo.Value.Date), Me.ObjCD.GetFirstDateofTheMonth(Me.DtpFrom.Value.Date))
                For Each row1 In TBL1.Rows
                    DS.TblCetak.DefaultView.RowFilter = "[Primary Employee] = " + row1.Primary_Employee.ToString
                    For Each dv As DataRowView In DS.TblCetak.DefaultView
                        RowCetak = dv.Row
                        RowCetak.Loan_Payment = row1.Loan_Installment
                        RowCetak.Company_Expense = row1.Company_Expense
                        RowCetak.Employee_Expense = row1.Employee_Expense
                    Next
                Next

                ReportDetail.SetDataSource(DS)
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Period_From.ParameterFieldName, Format(DateFirst, "MMM yyyy"))
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Period_To.ParameterFieldName, Format(DateLast, "MMM yyyy"))
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Format_Decimal.ParameterFieldName, FormatDecimal)
                ReportDetail.SetParameterValue(ReportDetail.Parameter_ParamStr.ParameterFieldName, "Detail Report")
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Standar_hari_kerja.ParameterFieldName, "0")
                ReportDetail.SetParameterValue(ReportDetail.Parameter_GroupPertama.ParameterFieldName, GroupPertamaStr)
                ReportRekap.SetDataSource(DS)
                ReportRekap.SetParameterValue(ReportRekap.Parameter_Period_From.ParameterFieldName, Format(DateFirst, "MMM yyyy"))
                ReportRekap.SetParameterValue(ReportRekap.Parameter_Period_To.ParameterFieldName, Format(DateLast, "MMM yyyy"))
                ReportRekap.SetParameterValue(ReportRekap.Parameter_Format_Decimal.ParameterFieldName, FormatDecimal)
                ReportRekap.SetParameterValue(ReportRekap.Parameter_ParamStr.ParameterFieldName, "Recapitulation Report")
                ReportRekap.SetParameterValue(ReportRekap.Parameter_Standar_hari_kerja.ParameterFieldName, "0")
                ReportRekap.SetParameterValue(ReportRekap.Parameter_GroupPertama.ParameterFieldName, GroupPertamaStr)
                Me.crvDetails.ReportSource = ReportDetail
                Me.crvRekap.ReportSource = ReportRekap
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListEmployee = Nothing
            DS.Dispose()
            DTA_A0.Dispose()
            DTA_A1.Dispose()
            DTA_A2.Dispose()
            DTA_B0.Dispose()
            DTA_B1.Dispose()
            DTA_B2.Dispose()
            DTA_C0.Dispose()
            DTA_C1.Dispose()
            DTA_C2.Dispose()
            TBL_A0.Dispose()
            TBL_A1.Dispose()
            TBL_A2.Dispose()
            TBL_B0.Dispose()
            TBL_B1.Dispose()
            TBL_B2.Dispose()
            TBL_C0.Dispose()
            TBL_C1.Dispose()
            TBL_C2.Dispose()
        End Try
    End Sub
End Class