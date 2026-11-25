Option Explicit On
Option Strict Off

Public Class FrmMasterWorkShift_Allocate
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim Buffer As String = ""
    Dim BufferFrom As String = ""
    Dim BufferTo As String = ""
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjProduct As New Core_Module.ClsMasterKaryawan
    Dim EraseRowIndex As Integer = 0
    Dim EraseRowIndex_WS As Integer = 0
    Dim EraseRowIndex_Del As Integer = 0
    Dim TBL_JP As New Core_Module.DseMasterWorkshift_Allocation.tblmaster_jobpositionDataTable
    Dim TBL_Dept As New Core_Module.DseMasterWorkshift_Allocation.tblmaster_departemenDataTable
    Dim TBL_Formula As New Core_Module.DseMasterWorkshift_Allocation.tblmaster_salary_formula1DataTable
    Dim AllData As Boolean = True
    Dim ListDataFilter As New List(Of Double)
    Dim ObjWork As New Core_Module.ClsMasterWorkShift
    Dim TBLKary As New Core_Module.DseMasterWorkshift_Allocation.tblmaster_karyawanDataTable
    Dim TBLAllocate_Item As New Core_Module.DseMasterWorkshift_Allocation.tblmaster_workshift_itemallocateDataTable
    Dim TBlOutputDGV As New Core_Module.DseMasterWorkshift_Allocation.Output_DGVDataTable
    Dim ObjAnt As New Core_Module.ClsAntrian
    Dim UserCodeLain As String = ""
    Dim NamaKomputerLain As String = ""
    Dim mClose As Boolean = False

    Public Enum EnumLaporan As Integer
        AllData
        JobPosition
        Departemen
        SalaryFormula
        Employee
    End Enum

    Private Sub IsiListBoxFormula(ByVal Lbo As System.Windows.Forms.CheckedListBox, ByVal TBL As DseMasterWorkshift_Allocation.tblmaster_salary_formula1DataTable)
        Dim DTA As New DseMasterWorkshift_AllocationTableAdapters.tblmaster_salary_formula1TableAdapter
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

    Private Sub IsiListBoxJobPosition(ByVal Lbo As System.Windows.Forms.CheckedListBox, ByVal TBL As DseMasterWorkshift_Allocation.tblmaster_jobpositionDataTable)
        Dim DTA As New DseMasterWorkshift_AllocationTableAdapters.tblmaster_jobpositionTableAdapter
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

    Private Sub IsiListBoxDepartemen(ByVal Lbo As System.Windows.Forms.CheckedListBox, ByVal TBL As DseMasterWorkshift_Allocation.tblmaster_departemenDataTable)
        Dim DTA As New DseMasterWorkshift_AllocationTableAdapters.tblmaster_departemenTableAdapter
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
            'Me.dgvEmployee.AllowUserToAddRows = False
            'Me.dgvEmployee.Rows.Clear()
            'Me.dgvEmployee.AllowUserToAddRows = True
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
            'Me.dgvEmployee.AllowUserToAddRows = False
            'Me.dgvEmployee.Rows.Clear()
            'Me.dgvEmployee.AllowUserToAddRows = True
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

    Private Function GetIndexItemInputWorkshift(ByVal vPrimaryWS As Double, ByVal vPrimaryKaryawan As Double, _
    ByVal vStartDate As Date, ByVal vEndDate As Date) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.DgvWorkshift.Rows
                If rowDGV.Cells(Me.ColPrimaryWorkShift.Index).Value = vPrimaryWS And _
                rowDGV.Cells(Me.ColPrimaryEntity.Index).Value = vPrimaryKaryawan And _
                rowDGV.Cells(Me.ColStartDate.Index).Value = vStartDate And _
                rowDGV.Cells(Me.ColEndDate.Index).Value = vEndDate Then
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

    Private Sub FrmMasterWorkShift_Allocate_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            If Me.mClose = True Then
                MsgBox("Can't access allocate workshift transaction because user code : " + _
                Me.UserCodeLain.ToString + " at computer name-ip address : " + Me.NamaKomputerLain.ToString + _
                " is using this transaction.", MsgBoxStyle.Information, "Information")
            End If
            Me.ObjAnt.EndAntrianTunggal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterWorkShift_Allocate_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.CancelButton = Me.BtnClose
            Me.dgvEmployee.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvWorkshift.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvEmployee.AllowUserToAddRows = False
            Me.IsiListBoxJobPosition(Me.ClbPositon, Me.TBL_JP)
            Me.IsiListBoxDepartemen(Me.ClbDepartemen, Me.TBL_Dept)
            Me.IsiListBoxFormula(Me.ClbSalaryFormula, Me.TBL_Formula)
            Me.KendalikanKontrol(EnumLaporan.AllData, False)
            Me.AllData = True
            'Antrian Tunggal
            If Me.ObjAnt.StartAntrianTunggal(ClsAntrian.EnumTipeAntrian.AntrianTunggalWorkshiftAllocate, Me.UserCode, Me.UserCodeLain, Me.NamaKomputerLain) = False Then
                Me.mClose = True
                Me.BackgroundWorker_CloseForm.RunWorkerAsync()
            End If
            Me.DtpStart.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFilter.Click
        Dim TotalBarisOutputDGV As Integer = -1
        Try
            If Me.DtpStart.Value.Date > Me.DtpEnd.Value.Date Then
                MsgBox("Period from should not be greater than period to.", MsgBoxStyle.Information, "Information")
                Me.DgvWorkshift.Rows.Clear()
                Me.TBlOutputDGV.Rows.Clear()
                Exit Sub
            End If
            If Me.ChkDepartemen.Checked = True Then
                Me.AllData = False
            ElseIf Me.ChkEmployee.Checked = True Then
                Me.AllData = False
            ElseIf Me.ChkFormula.Checked = True Then
                Me.AllData = False
            ElseIf Me.ChkPosition.Checked = True Then
                Me.AllData = False
            End If
            'Fill Data Karyawan dan sesuaikan dengan kriteria Selection Data
            If Me.AllData = True Then
                Me.ObjWork.KelolaDataKaryawan(Me.TBLKary)
            Else
                Me.ListDataFilter.Clear()
                If Me.ChkDepartemen.Checked = True Then
                    If Me.ListDataFilter.Contains(FrmMasterWorkShift_Allocate.EnumLaporan.Departemen) = False Then
                        Me.ListDataFilter.Add(FrmMasterWorkShift_Allocate.EnumLaporan.Departemen)
                    End If
                End If
                If Me.ChkEmployee.Checked = True Then
                    If Me.ListDataFilter.Contains(FrmMasterWorkShift_Allocate.EnumLaporan.Employee) = False Then
                        Me.ListDataFilter.Add(FrmMasterWorkShift_Allocate.EnumLaporan.Employee)
                    End If
                End If
                If Me.ChkFormula.Checked = True Then
                    If Me.ListDataFilter.Contains(FrmMasterWorkShift_Allocate.EnumLaporan.SalaryFormula) = False Then
                        Me.ListDataFilter.Add(FrmMasterWorkShift_Allocate.EnumLaporan.SalaryFormula)
                    End If
                End If
                If Me.ChkPosition.Checked = True Then
                    If Me.ListDataFilter.Contains(FrmMasterWorkShift_Allocate.EnumLaporan.JobPosition) = False Then
                        Me.ListDataFilter.Add(FrmMasterWorkShift_Allocate.EnumLaporan.JobPosition)
                    End If
                End If
                Me.ObjWork.KelolaDataKaryawan(Me.TBLKary, Me.ListDataFilter, Me.dgvEmployee, Me.TBL_JP, _
                Me.ClbPositon, Me.TBL_Dept, Me.ClbDepartemen, Me.TBL_Formula, Me.ClbSalaryFormula, Me)
            End If
            'Penyesuaian data dengan analogi allocate workshift
            Me.ObjWork.FilterDataAllocation(Me.TBLKary, Me.TBLAllocate_Item, Me.DtpEnd.Value, Me.DtpStart.Value, Me.TBlOutputDGV)
            'key in DGV-nya
            Me.ObjWork.KeyInDGV_Allocation(Me.TBlOutputDGV, Me.DgvWorkshift, Me)
            'Total Baris
            TotalBarisOutputDGV = Me.TBlOutputDGV.Rows.Count
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnAllocate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAllocate.Click
        Dim ij As New FrmMasterWorkShift_Cari(ClsMasterWorkShift.EnumTipeWorkshift.HR)
        Dim PrimaryWS As Double = -1
        Dim Hasil As MsgBoxResult
        '-----Parameter Set Ulang-----
        'Dim RowOutput As Core_Module.DseMasterWorkshift_Allocation.Output_DGVRow
        'Dim TBLSetUlang As New Core_Module.DseMasterWorkshift_Allocation.Output_DGVDataTable
        'Dim RowSetUlang As Core_Module.DseMasterWorkshift_Allocation.Output_DGVRow
        'Dim RV As Data.DataRowView
        'Dim TglDariAllocate As Date = #1/1/1950#
        'Dim TglSampaiAllocate As Date = #1/1/1950#
        'Dim Filter As String = ""
        '------------------------------
        Try
            'If Me.ObjWork.ApakahDataBolehEditAllocate(Me, TBLSetUlang) = True Then
            If Me.ObjWork.ApakahDataBolehHapusAllocate(Me.TBlOutputDGV, Me) = True Then
                ij.StartPosition = FormStartPosition.CenterScreen
                ij.BackColorAll = Me.BackColorAll
                ij.ShowDialog()
                If ij.PrimaryWorkshift <> -1 Then
                    PrimaryWS = ij.PrimaryWorkshift
                    Hasil = MsgBox("Do you want to allocate this data ?", MsgBoxStyle.YesNo, "Confirm")
                    If Hasil = MsgBoxResult.Yes Then
                        '1a. Hapus data allocate yang tidak ada Attendance.
                        Me.BtnFilter.PerformClick()
                        '----------------------------------------------------
                        'Periksa ketersesuaian dengan Allocation, kalo ada allocation dihapus sementara
                        Me.ObjWork.ApakahDataBolehHapusAllocate(Me.TBlOutputDGV, Me, True)
                        'Hapus data
                        Me.ObjWork.DeleteDataAllocation(Me.TBlOutputDGV, Me.DtpEnd.Value.Date, Me.DtpStart.Value.Date)
                        '1b. Insert data allocate yang tidak ada Attendance
                        Me.ObjWork.FilterDataAllocation(Me.TBLKary, Me.TBLAllocate_Item, Me.DtpEnd.Value, Me.DtpStart.Value, Me.TBlOutputDGV)
                        Me.ObjWork.KeyInDGV_Allocation(Me.TBlOutputDGV, Me.DgvWorkshift, Me)
                        'Periksa ketersesuaian dengan Allocation, kalo ada allocation dihapus sementara
                        Me.ObjWork.ApakahDataBolehHapusAllocate(Me.TBlOutputDGV, Me, True)
                        'insert data 
                        Me.ObjWork.InsertDataAllocation(Me.TBlOutputDGV, PrimaryWS)
                        MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                        Me.DgvWorkshift.Rows.Clear()
                        Me.BtnFilter.PerformClick()
                        ''----Coding set ulang kembali data yang tidak ingin di hapus----
                        'If TBLSetUlang.Rows.Count > 0 Then
                        '    'Delete Allocate
                        '    For Each RowSetUlang In TBLSetUlang.Rows
                        '        Filter = "[Employee Code] = '" + RowSetUlang.Employee_Code.ToString + "'"
                        '        Me.TBlOutputDGV.DefaultView.RowFilter = Filter
                        '        If Me.TBlOutputDGV.DefaultView.Count > 0 Then
                        '            For Each RV In Me.TBlOutputDGV.DefaultView
                        '                RowOutput = RV.Row
                        '                TglDariAllocate = #1/1/1950#
                        '                TglSampaiAllocate = #1/1/1950#
                        '                If RowOutput.Primary_WS <> -1 Then
                        '                    If RowSetUlang.Start_Date <= RowOutput.Start_Date And RowSetUlang.End_Date >= RowOutput.End_Date Then
                        '                        TglDariAllocate = #1/1/1950#
                        '                        TglSampaiAllocate = #1/1/1950#
                        '                        PrimaryWS = -1
                        '                        Me.ObjWork.HapusItemAllocate(RowOutput.Primary)
                        '                    ElseIf RowSetUlang.Start_Date >= RowOutput.Start_Date And RowSetUlang.End_Date >= RowOutput.End_Date Then
                        '                        TglDariAllocate = RowOutput.Start_Date
                        '                        TglSampaiAllocate = #1/1/1950#
                        '                        PrimaryWS = RowOutput.Primary_WS
                        '                        Me.ObjWork.HapusItemAllocate(RowOutput.Primary)
                        '                        Me.ObjWork.SimpanItemAllocate(RowOutput.Primary_Entity, RowOutput.Tipe_Entity, _
                        '                        RowSetUlang.Start_Date, TglDariAllocate, PrimaryWS, RowOutput.Remarks)
                        '                    ElseIf RowSetUlang.Start_Date <= RowOutput.Start_Date And RowSetUlang.End_Date <= RowOutput.End_Date Then
                        '                        TglDariAllocate = #1/1/1950#
                        '                        TglSampaiAllocate = RowOutput.End_Date
                        '                        PrimaryWS = RowOutput.Primary_WS
                        '                        Me.ObjWork.HapusItemAllocate(RowOutput.Primary)
                        '                        Me.ObjWork.SimpanItemAllocate(RowOutput.Primary_Entity, RowOutput.Tipe_Entity, _
                        '                         TglSampaiAllocate, RowSetUlang.End_Date, PrimaryWS, RowOutput.Remarks)
                        '                    ElseIf RowSetUlang.Start_Date >= RowOutput.Start_Date And RowSetUlang.End_Date <= RowOutput.End_Date Then
                        '                        TglDariAllocate = RowOutput.Start_Date
                        '                        TglSampaiAllocate = RowOutput.End_Date
                        '                        PrimaryWS = RowOutput.Primary_WS
                        '                        Me.ObjWork.HapusItemAllocate(RowOutput.Primary)
                        '                        Me.ObjWork.SimpanItemAllocate(RowOutput.Primary_Entity, RowOutput.Tipe_Entity, _
                        '                        RowSetUlang.Start_Date, TglDariAllocate, PrimaryWS, RowOutput.Remarks)
                        '                        Me.ObjWork.SimpanItemAllocate(RowOutput.Primary_Entity, RowOutput.Tipe_Entity, _
                        '                        TglSampaiAllocate, RowSetUlang.End_Date, PrimaryWS, RowOutput.Remarks)
                        '                    End If
                        '                End If
                        '            Next
                        '        End If
                        '    Next
                        '    Me.TBlOutputDGV.DefaultView.RowFilter = ""
                        '    'Insert allocate
                        '    For Each RowSetUlang In TBLSetUlang.Rows
                        '        Me.ObjWork.SimpanItemAllocate(RowSetUlang.Primary_Entity, RowSetUlang.Tipe_Entity, RowSetUlang.End_Date, _
                        '        RowSetUlang.Start_Date, RowSetUlang.Primary_WS, RowSetUlang.Remarks)
                        '    Next
                        'End If
                        ''---------------------------------------------------------------
                        'Me.DgvWorkshift.Rows.Clear()
                        'Me.BtnFilter.PerformClick()
                        Me.BtnFilter.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
            'TBLSetUlang.Dispose()
        End Try
    End Sub

    Private Function ApakahBolehHapusAllocate() As Boolean
        Dim RowDGV As DataGridViewRow
        Dim Ketemu As Boolean = False
        Try
            If Me.DgvWorkshift.Rows.Count > 0 Then
                For Each RowDGV In Me.DgvWorkshift.Rows
                    If RowDGV.Cells(Me.ColPrimary.Index).Value <> -1 Then
                        Ketemu = True
                        Exit For
                    End If
                Next
            End If
            If Ketemu = True Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Function

    Private Sub BtnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDelete.Click
        Dim Hasil As MsgBoxResult
        Try
            If Me.ObjWork.ApakahDataBolehHapusAllocate(Me.TBlOutputDGV, Me) = True Then
                If Me.ApakahBolehHapusAllocate() = True Then
                    Hasil = MsgBox("Do you want to delete this data ?", MsgBoxStyle.YesNo, "Confirm")
                    If Hasil = MsgBoxResult.Yes Then
                        Me.ObjWork.DeleteDataAllocation(Me.TBlOutputDGV, Me.DtpEnd.Value.Date, Me.DtpStart.Value.Date)
                        MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        Me.DgvWorkshift.Rows.Clear()
                        Me.BtnFilter.PerformClick()
                        Me.BtnFilter.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpStart_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpStart.Enter
        Try
            Me.BufferFrom = Me.DtpStart.Value.Date
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpStart_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpStart.Leave
        Try
            'Me.BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            If CDate(Me.BufferFrom) <> Me.DtpStart.Value.Date Then
                If Me.DtpStart.Value.Date > Me.DtpEnd.Value.Date Then
                    MsgBox("Period from should not be greater than period to.", MsgBoxStyle.Information, "Information")
                    Me.DtpStart.Value = Me.BufferFrom
                    Exit Sub
                End If
                Me.BtnFilter.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpEnd_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpEnd.Enter
        Try
            Me.BufferTo = Me.DtpEnd.Value.Date
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpEnd_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpEnd.Leave
        Try
            'Me.BackgroundWorker3.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Try
            If CDate(Me.BufferTo) <> Me.DtpEnd.Value.Date Then
                If Me.DtpEnd.Value.Date < Me.DtpStart.Value.Date Then
                    MsgBox("Period to should not be less than period from.", MsgBoxStyle.Information, "Information")
                    Me.DtpEnd.Value = Me.BufferTo
                    Exit Sub
                End If
                Me.BtnFilter.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvWorkshift_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvWorkshift.CellBeginEdit
        Dim Row As DataGridViewRow = Me.DgvWorkshift.Rows(e.RowIndex)
        Dim TBL1 As New DseMasterWorkshift_Allocation.tbltrans_attendance_headerDataTable
        Dim Row1 As DseMasterWorkshift_Allocation.tbltrans_attendance_headerRow
        Dim TBL2 As New DseMasterWorkshift_Allocation.tbl_overtime_order_headerDataTable
        Dim Row2 As DseMasterWorkshift_Allocation.tbl_overtime_order_headerRow
        Dim Nomor As Integer = 1
        Dim Retval As String = ""
        Try
            If e.ColumnIndex = Me.ColStartDate.Index Or _
            e.ColumnIndex = Me.ColEndDate.Index Or _
            e.ColumnIndex = Me.ColWorkshiftCode.Index Or _
            e.ColumnIndex = Me.ColRemarks.Index Then
                If e.ColumnIndex = Me.ColStartDate.Index Or e.ColumnIndex = Me.ColEndDate.Index Then
                    If Row.Cells(Me.ColPrimaryWorkShift.Index).Value = "-1" Then
                        e.Cancel = True
                    Else
                        Me.Buffer = Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                ElseIf e.ColumnIndex = Me.ColWorkshiftCode.Index Then
                    Me.ObjWork.FungsiFillDataAttendanceItem(TBL1, CDate(Row.Cells(Me.ColStartDate.Index).Value).Date, _
                    CDate(Row.Cells(Me.ColEndDate.Index).Value).Date, Row.Cells(Me.ColPrimaryEntity.Index).Value)
                    If TBL1.Rows.Count > 0 Then
                        For Each Row1 In TBL1.Rows
                            If Retval = "" Then
                                Retval = Nomor.ToString + ". " + Row1.no_transaksi.ToString
                            Else
                                Retval += vbCrLf + Nomor.ToString + ". " + Row1.no_transaksi.ToString
                            End If
                            Nomor += 1
                        Next
                        MsgBox("Workshift code cannot edit because in period " + _
                        Format(CDate(Row.Cells(Me.ColStartDate.Index).Value), "dd MMM yyyy") + _
                        " to " + Format(CDate(Row.Cells(Me.ColEndDate.Index).Value), "dd MMM yyyy") + _
                        " is refering to attendance. " + vbCrLf + Retval.ToString, MsgBoxStyle.Information, "Information")
                        e.Cancel = True
                        Exit Sub
                    End If

                    'aul tambahan cek juga ke overtime order
                    Me.ObjWork.FungsiFillDataOvertimeOrderItem(TBL2, CDate(Row.Cells(Me.ColStartDate.Index).Value).Date, _
                    CDate(Row.Cells(Me.ColEndDate.Index).Value).Date, Row.Cells(Me.ColPrimaryEntity.Index).Value)
                    If TBL2.Rows.Count > 0 Then
                        For Each Row2 In TBL2.Rows
                            If Retval = "" Then
                                Retval = Nomor.ToString + ". " + Row2.No_Transaksi.ToString
                            Else
                                Retval += vbCrLf + Nomor.ToString + ". " + Row2.No_Transaksi.ToString
                            End If
                            Nomor += 1
                        Next
                        MsgBox("Workshift code cannot edit because in period " + _
                        Format(CDate(Row.Cells(Me.ColStartDate.Index).Value), "dd MMM yyyy") + _
                        " to " + Format(CDate(Row.Cells(Me.ColEndDate.Index).Value), "dd MMM yyyy") + _
                        " is refering to overtime order. " + vbCrLf + Retval.ToString, MsgBoxStyle.Information, "Information")
                        e.Cancel = True
                        Exit Sub
                    End If

                    Me.Buffer = Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    Me.Buffer = Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBL1.Dispose()
        End Try
    End Sub

    Private Function ApakahBolehHapusRowAllocate(ByVal Row As DataGridViewRow) As Boolean
        Dim TBL1 As New DseMasterWorkshift_Allocation.tbltrans_attendance_headerDataTable
        Dim Row1 As DseMasterWorkshift_Allocation.tbltrans_attendance_headerRow
        Dim TBL2 As New DseMasterWorkshift_Allocation.tbl_overtime_order_headerDataTable
        Dim Row2 As DseMasterWorkshift_Allocation.tbl_overtime_order_headerRow
        Dim Nomor As Integer = 1
        Dim Retval As String = ""
        Try
            Me.ObjWork.FungsiFillDataAttendanceItem(TBL1, CDate(Row.Cells(Me.ColStartDate.Index).Value).Date, _
            CDate(Row.Cells(Me.ColEndDate.Index).Value).Date, Row.Cells(Me.ColPrimaryEntity.Index).Value)
            If TBL1.Rows.Count > 0 Then
                For Each Row1 In TBL1.Rows
                    If Retval = "" Then
                        Retval = Nomor.ToString + ". " + Row1.no_transaksi.ToString
                    Else
                        Retval += vbCrLf + Nomor.ToString + ". " + Row1.no_transaksi.ToString
                    End If
                    Nomor += 1
                Next
                MsgBox("Workshift code cannot delete because in period " + _
                Format(CDate(Row.Cells(Me.ColStartDate.Index).Value), "dd MMM yyyy") + _
                " to " + Format(CDate(Row.Cells(Me.ColEndDate.Index).Value), "dd MMM yyyy") + _
                " is refering to attendance. " + vbCrLf + Retval.ToString, MsgBoxStyle.Information, "Information")
                Return False
            End If

            'aul tambahan cek juga ke overtime order
            Me.ObjWork.FungsiFillDataOvertimeOrderItem(TBL2, CDate(Row.Cells(Me.ColStartDate.Index).Value).Date, _
            CDate(Row.Cells(Me.ColEndDate.Index).Value).Date, Row.Cells(Me.ColPrimaryEntity.Index).Value)
            If TBL2.Rows.Count > 0 Then
                For Each Row2 In TBL2.Rows
                    If Retval = "" Then
                        Retval = Nomor.ToString + ". " + Row2.No_Transaksi.ToString
                    Else
                        Retval += vbCrLf + Nomor.ToString + ". " + Row2.No_Transaksi.ToString
                    End If
                    Nomor += 1
                Next
                MsgBox("Workshift code cannot delete because in period " + _
                Format(CDate(Row.Cells(Me.ColStartDate.Index).Value), "dd MMM yyyy") + _
                " to " + Format(CDate(Row.Cells(Me.ColEndDate.Index).Value), "dd MMM yyyy") + _
                " is refering to overtime order. " + vbCrLf + Retval.ToString, MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBL1.Dispose()
        End Try
    End Function

    Private Sub SimpanRowWorkshift(ByVal Row As DataGridViewRow)
        Dim Hasil As MsgBoxResult
        Try
            If Row.Cells(Me.ColSave.Index).Value = "save" Then
                Hasil = MsgBoxResult.Yes
                If Hasil = MsgBoxResult.Yes Then
                    Dim Output_Index As Integer = Row.Index
                    Dim RowOutput_Index As Core_Module.DseMasterWorkshift_Allocation.Output_DGVRow
                    RowOutput_Index = Me.TBlOutputDGV.Rows(Output_Index)
                    If Row.Cells(Me.ColPrimary.Index).Value = "-1" Then
                        'simpan baris baru, update data ke TBLOutputDGV, dan get primary auto
                        Dim Auto_Primary As Double = -1
                        Me.ObjWork.SimpanItemAllocate(Row.Cells(Me.ColPrimaryEntity.Index).Value, Row.Cells(Me.ColEntityTipe.Index).Value, _
                        Row.Cells(Me.ColEndDate.Index).Value, Row.Cells(Me.ColStartDate.Index).Value, Row.Cells(Me.ColPrimaryWorkShift.Index).Value, _
                        Row.Cells(Me.ColRemarks.Index).Value)
                        'Get Primary auto number
                        Auto_Primary = Me.ObjWork.GetPrimaryAutoAllocate(Row.Cells(Me.ColPrimaryEntity.Index).Value, Row.Cells(Me.ColEntityTipe.Index).Value, _
                        Row.Cells(Me.ColStartDate.Index).Value, Row.Cells(Me.ColEndDate.Index).Value, Row.Cells(Me.ColPrimaryWorkShift.Index).Value)
                        Row.Cells(Me.ColPrimary.Index).Value = Auto_Primary
                        'Update data TBLOutput
                        With RowOutput_Index
                            .Employee_Code = Row.Cells(Me.ColEmployeeCode.Index).Value
                            .Employee_Name = Row.Cells(Me.ColEmployeeName.Index).Value
                            .Start_Date = CDate(Row.Cells(Me.ColStartDate.Index).Value)
                            .End_Date = CDate(Row.Cells(Me.ColEndDate.Index).Value)
                            .Workshift_Code = Row.Cells(Me.ColWorkshiftCode.Index).Value
                            .Workshift_Name = Row.Cells(Me.ColWorkShiftName.Index).Value
                            .Remarks = Row.Cells(Me.ColRemarks.Index).Value
                            .Primary_Entity = Row.Cells(Me.ColPrimaryEntity.Index).Value
                            .Primary_WS = Row.Cells(Me.ColPrimaryWorkShift.Index).Value
                            .Tipe_Entity = Row.Cells(Me.ColEntityTipe.Index).Value
                            .Primary = Row.Cells(Me.ColPrimary.Index).Value
                        End With
                        'MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                        Row.Cells(Me.ColSave.Index).Value = ""
                    Else
                        'Update Database, dan Update TBLOutputDGV
                        Me.ObjWork.UpdateDataPerItemAllocate(Row, Row.Cells(Me.ColPrimary.Index).Value, Me)
                        'Update data TBLOutput
                        With RowOutput_Index
                            .Employee_Code = Row.Cells(Me.ColEmployeeCode.Index).Value
                            .Employee_Name = Row.Cells(Me.ColEmployeeName.Index).Value
                            My.Application.ChangeCulture("id-ID")
                            My.Application.ChangeUICulture("id-ID")
                            .Start_Date = CDate(Row.Cells(Me.ColStartDate.Index).Value)
                            .End_Date = CDate(Row.Cells(Me.ColEndDate.Index).Value)
                            My.Application.ChangeCulture("en-US")
                            My.Application.ChangeUICulture("en-US")
                            .Workshift_Code = Row.Cells(Me.ColWorkshiftCode.Index).Value
                            .Workshift_Name = Row.Cells(Me.ColWorkShiftName.Index).Value
                            .Remarks = Row.Cells(Me.ColRemarks.Index).Value
                            .Primary_Entity = Row.Cells(Me.ColPrimaryEntity.Index).Value
                            .Primary_WS = Row.Cells(Me.ColPrimaryWorkShift.Index).Value
                            .Tipe_Entity = Row.Cells(Me.ColEntityTipe.Index).Value
                            .Primary = Row.Cells(Me.ColPrimary.Index).Value
                        End With
                        'MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                        Row.Cells(Me.ColSave.Index).Value = ""
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvWorkshift_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvWorkshift.CellContentClick
        Dim Hasil As MsgBoxResult
        Dim Row As DataGridViewRow
        Try
            If e.ColumnIndex = -1 Then
                Exit Sub
            End If
            If e.RowIndex = -1 Then
                Exit Sub
            End If
            Row = Me.DgvWorkshift.Rows(e.RowIndex)
            If e.ColumnIndex = Me.ColSave.Index Then
                If Row.Cells(Me.ColSave.Index).Value = "save" Then
                    'Hasil = MsgBox("Saved row?", MsgBoxStyle.YesNo, "Confirm")
                    'Me.SimpanRowWorkshift(Row)
                End If
            End If
            If e.ColumnIndex = Me.ColDelete.Index Then
                'Apakah Boleh Hapus Baris DGV Allocate
                If Me.ApakahBolehHapusRowAllocate(Row) = True Then
                    Hasil = MsgBox("Delete row?", MsgBoxStyle.YesNo, "Confirm")
                    If Hasil = MsgBoxResult.Yes Then
                        'Hapus
                        If Row.Cells(Me.ColPrimary.Index).Value <> -1 Then
                            Me.ObjWork.HapusItemAllocate(Row.Cells(Me.ColPrimary.Index).Value)
                            Me.EraseRowIndex_Del = e.RowIndex
                            Me.BackgroundWorker5.RunWorkerAsync()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        Else
                            Me.EraseRowIndex_Del = e.RowIndex
                            Me.BackgroundWorker5.RunWorkerAsync()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                        'Load data ulang
                        Me.BtnFilter.PerformClick()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvWorkshift_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvWorkshift.CellEndEdit
        Dim Row As DataGridViewRow = Me.DgvWorkshift.Rows(e.RowIndex)
        Dim Index As Integer
        Dim TBLStart As New DseMasterWorkshift_Allocation.tbltrans_attendance_headerDataTable
        Dim TBLEnd As New DseMasterWorkshift_Allocation.tbltrans_attendance_headerDataTable
        Dim RowStart As DseMasterWorkshift_Allocation.tbltrans_attendance_headerRow
        Dim RowEnd As DseMasterWorkshift_Allocation.tbltrans_attendance_headerRow
        'Dim TS As TimeSpan
        Try
            With Row
                If IsNothing(Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        If Me.ObjCD.CekDataKosong(.Cells(Me.ColRemarks.Index).Value, False) = False Then
                            MsgBox("Data shouldn't include the character: ' , % , or * .", MsgBoxStyle.Information, "Information")
                            Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        If Me.ObjWork.ApakahDataBenar_ItemAllocate_Save(Row, Me) = True Then
                            Row.Cells(Me.ColSave.Index).Value = "save"
                            Me.SimpanRowWorkshift(Row)
                        End If
                    End If
                    If e.ColumnIndex = Me.ColStartDate.Index Then
                        My.Application.ChangeCulture("id-ID")
                        My.Application.ChangeUICulture("id-ID")
                        If IsDate(Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            If CDate(Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value).Date = CDate(Me.Buffer).Date Then
                                My.Application.ChangeCulture("en-US")
                                My.Application.ChangeUICulture("en-US")
                                Exit Sub
                            Else
                                'sebelum cek pastikan start date, end date pada posisi yang benar
                                If CDate(Row.Cells(Me.ColStartDate.Index).Value).Date > CDate(Row.Cells(Me.ColEndDate.Index).Value).Date Then
                                    MsgBox("Period start date should not be greater than period end date.", MsgBoxStyle.Information, "Information")
                                    Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    My.Application.ChangeCulture("en-US")
                                    My.Application.ChangeUICulture("en-US")
                                    Exit Sub
                                End If
                                'pengecekan setelah edit
                                '1. Boleh Geser
                                If Me.DgvWorkshift.Item(Me.ColEndDate.Index, e.RowIndex).Value < CDate(Me.Buffer) Then
                                    If Me.ObjWork.GetJmlAttendance(Row.Cells(Me.ColPrimaryWorkShift.Index).Value, _
                                    CDate(Me.Buffer).Date, Row.Cells(Me.ColPrimaryEntity.Index).Value) > 0 Then
                                        'MsgBox("Data cannot edit because this period used in transaction attendance.", MsgBoxStyle.Information, "Information")
                                        MsgBox("Data cannot edit because this period: " + Format(CDate(Me.Buffer).Date, "dd MMM yyyy").ToString + " used in transaction attendance no. " + Me.ObjWork.GetNomerAttendance(Row.Cells(Me.ColPrimaryWorkShift.Index).Value, CDate(Me.Buffer).Date, Row.Cells(Me.ColPrimaryEntity.Index).Value).ToString + ".", MsgBoxStyle.Information, "Information")
                                        Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        My.Application.ChangeCulture("en-US")
                                        My.Application.ChangeUICulture("en-US")
                                        Exit Sub
                                    End If

                                    'cek juga ke overtime order
                                    If Me.ObjWork.GetJmlOvertimeOrder(Row.Cells(Me.ColPrimaryWorkShift.Index).Value, _
                                    CDate(Me.Buffer).Date, Row.Cells(Me.ColPrimaryEntity.Index).Value) > 0 Then
                                        'MsgBox("Data cannot edit because this period used in transaction attendance.", MsgBoxStyle.Information, "Information")
                                        MsgBox("Data cannot edit because this period: " + Format(CDate(Me.Buffer).Date, "dd MMM yyyy").ToString + " used in transaction overtime order no. " + Me.ObjWork.GetNomerOvertimeOrder(Row.Cells(Me.ColPrimaryWorkShift.Index).Value, CDate(Me.Buffer).Date, Row.Cells(Me.ColPrimaryEntity.Index).Value).ToString + ".", MsgBoxStyle.Information, "Information")
                                        Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        My.Application.ChangeCulture("en-US")
                                        My.Application.ChangeUICulture("en-US")
                                        Exit Sub
                                    End If
                                End If

                                '2. Geser ke Luar
                                If CDate(Row.Cells(Me.ColStartDate.Index).Value).Date < CDate(Me.Buffer).Date Then
                                    'My.Application.ChangeCulture("en-US")
                                    'My.Application.ChangeUICulture("en-US")
                                    If Me.ObjWork.GetJmlBatasanGeserLuarStartDate(CDate(Row.Cells(Me.ColStartDate.Index).Value).Date, _
                                    CDate(Me.Buffer).Date, CDbl(Row.Cells(Me.ColPrimary.Index).Value), Row.Cells(Me.ColPrimaryEntity.Index).Value) > 0 Then
                                        'My.Application.ChangeCulture("id-ID")
                                        'My.Application.ChangeUICulture("id-ID")
                                        MsgBox("Data cannot edit because this period used in other allocation.", MsgBoxStyle.Information, "Information")
                                        Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        My.Application.ChangeCulture("en-US")
                                        My.Application.ChangeUICulture("en-US")
                                        Exit Sub
                                    End If
                                    'My.Application.ChangeCulture("id-ID")
                                    'My.Application.ChangeUICulture("id-ID")
                                End If
                                '3. Geser ke dalam
                                If CDate(Row.Cells(Me.ColStartDate.Index).Value).Date > CDate(Me.Buffer).Date Then
                                    If Me.ObjWork.GetJmlAttendance(Row.Cells(Me.ColPrimaryWorkShift.Index).Value, _
                                    CDate(Row.Cells(Me.ColStartDate.Index).Value).Date - TimeSpan.FromDays(1), Row.Cells(Me.ColPrimaryEntity.Index).Value) > 0 Then 'aul perubahan - CDate(Row.Cells(Me.ColStartDate.Index).Value).Date - TimeSpan.FromDays(1)
                                        'MsgBox("Data cannot edit because this period used in transaction attendance.", MsgBoxStyle.Information, "Information")
                                        MsgBox("Data cannot edit because this period: " + Format(CDate(Row.Cells(Me.ColStartDate.Index).Value).Date, "dd MMM yyyy").ToString + " used in transaction attendance no. " + Me.ObjWork.GetNomerAttendance(Row.Cells(Me.ColPrimaryWorkShift.Index).Value, CDate(Row.Cells(Me.ColStartDate.Index).Value).Date, Row.Cells(Me.ColPrimaryEntity.Index).Value).ToString + ".", MsgBoxStyle.Information, "Information")
                                        Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        My.Application.ChangeCulture("en-US")
                                        My.Application.ChangeUICulture("en-US")
                                        Exit Sub
                                    End If

                                    'cek juga ke overtime order
                                    If Me.ObjWork.GetJmlOvertimeOrder(Row.Cells(Me.ColPrimaryWorkShift.Index).Value, _
                                    CDate(Row.Cells(Me.ColStartDate.Index).Value).Date - TimeSpan.FromDays(1), Row.Cells(Me.ColPrimaryEntity.Index).Value) > 0 Then 'aul perubahan - CDate(Row.Cells(Me.ColStartDate.Index).Value).Date - TimeSpan.FromDays(1)
                                        'MsgBox("Data cannot edit because this period used in transaction attendance.", MsgBoxStyle.Information, "Information")
                                        MsgBox("Data cannot edit because this period: " + Format(CDate(Row.Cells(Me.ColStartDate.Index).Value).Date, "dd MMM yyyy").ToString + " used in transaction overtime order no. " + Me.ObjWork.GetNomerOvertimeOrder(Row.Cells(Me.ColPrimaryWorkShift.Index).Value, CDate(Row.Cells(Me.ColStartDate.Index).Value).Date, Row.Cells(Me.ColPrimaryEntity.Index).Value).ToString + ".", MsgBoxStyle.Information, "Information")
                                        Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        My.Application.ChangeCulture("en-US")
                                        My.Application.ChangeUICulture("en-US")
                                        Exit Sub
                                    End If

                                    If CDate(Row.Cells(Me.ColStartDate.Index).Value).Date - TimeSpan.FromDays(1) <> CDate(Me.Buffer).Date + TimeSpan.FromDays(1) Then
                                        'My.Application.ChangeCulture("en-US")
                                        'My.Application.ChangeUICulture("en-US")
                                        Me.ObjWork.FungsiFillDataAttendanceItem(TBLStart, CDate(Me.Buffer).Date + TimeSpan.FromDays(1), CDate(Row.Cells(Me.ColStartDate.Index).Value).Date - TimeSpan.FromDays(1), Row.Cells(Me.ColPrimaryEntity.Index).Value)
                                        If TBLStart.Rows.Count > 0 Then
                                            Dim ik_Start_No As Integer = 0
                                            Dim ik_Start_Idx As Integer = 0
                                            Dim ik_Start_Retval As String = ""
                                            While ik_Start_Idx <= TBLStart.Count - 1
                                                RowStart = TBLStart.Rows(ik_Start_Idx)
                                                ik_Start_No += 1
                                                If ik_Start_Retval = "" Then
                                                    ik_Start_Retval = ik_Start_No.ToString + ". " + RowStart.no_transaksi.ToString
                                                Else
                                                    ik_Start_Retval += vbCrLf + ik_Start_No.ToString + ". " + RowStart.no_transaksi.ToString
                                                End If
                                                ik_Start_Idx += 1
                                            End While
                                            'My.Application.ChangeCulture("id-ID")
                                            'My.Application.ChangeUICulture("id-ID")
                                            MsgBox("Period cannot edit because in period " + Format(CDate(Me.Buffer).Date + TimeSpan.FromDays(1), "dd MMM yyyy") + " to " + _
                                            Format(CDate(Row.Cells(Me.ColStartDate.Index).Value).Date - TimeSpan.FromDays(1), "dd MMM yyyy") + _
                                            " is refering to attendance. " + vbCrLf + ik_Start_Retval.ToString, MsgBoxStyle.Information, "Information")
                                            Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                            My.Application.ChangeCulture("en-US")
                                            My.Application.ChangeUICulture("en-US")
                                            Exit Sub
                                        End If
                                        My.Application.ChangeCulture("id-ID")
                                        My.Application.ChangeUICulture("id-ID")
                                    End If
                                End If
                                'data benar
                                Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")
                                If Me.ObjWork.ApakahDataBenar_ItemAllocate_Save(Row, Me) = True Then
                                    Row.Cells(Me.ColSave.Index).Value = "save"
                                    Me.SimpanRowWorkshift(Row)
                                    'Me.BtnFilter.PerformClick()
                                    Me.BackgroundWorker6.RunWorkerAsync()
                                End If
                            End If
                        Else
                            'data salah
                            Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        My.Application.ChangeCulture("en-US")
                        My.Application.ChangeUICulture("en-US")
                    End If
                    If e.ColumnIndex = Me.ColEndDate.Index Then
                        My.Application.ChangeCulture("id-ID")
                        My.Application.ChangeUICulture("id-ID")
                        If IsDate(Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            If CDate(Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value).Date = CDate(Me.Buffer).Date Then
                                My.Application.ChangeCulture("en-US")
                                My.Application.ChangeUICulture("en-US")
                                Exit Sub
                            Else
                                'sebelum cek pastikan start date, end date pada posisi yang benar
                                If CDate(Row.Cells(Me.ColStartDate.Index).Value).Date > CDate(Row.Cells(Me.ColEndDate.Index).Value).Date Then
                                    MsgBox("Period start date should not be greater than period end date.", MsgBoxStyle.Information, "Information")
                                    Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    My.Application.ChangeCulture("en-US")
                                    My.Application.ChangeUICulture("en-US")
                                    Exit Sub
                                End If
                                'Pengecekan setelah edit 
                                '1. Boleh Geser
                                If Me.DgvWorkshift.Item(Me.ColEndDate.Index, e.RowIndex).Value < CDate(Me.Buffer) Then
                                    If Me.ObjWork.GetJmlAttendance(Row.Cells(Me.ColPrimaryWorkShift.Index).Value, _
                                    CDate(Me.Buffer).Date, Row.Cells(Me.ColPrimaryEntity.Index).Value) > 0 Then
                                        'MsgBox("Data cannot edit because this period used in transaction attendance.", MsgBoxStyle.Information, "Information")
                                        MsgBox("Data cannot edit because this period: " + Format(CDate(Me.Buffer).Date, "dd MMM yyyy").ToString + " used in transaction attendance no. " + Me.ObjWork.GetNomerAttendance(Row.Cells(Me.ColPrimaryWorkShift.Index).Value, CDate(Me.Buffer).Date, Row.Cells(Me.ColPrimaryEntity.Index).Value).ToString + ".", MsgBoxStyle.Information, "Information")
                                        Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        My.Application.ChangeCulture("en-US")
                                        My.Application.ChangeUICulture("en-US")
                                        Exit Sub
                                    End If

                                    'cek juga ke overtime order
                                    If Me.ObjWork.GetJmlOvertimeOrder(Row.Cells(Me.ColPrimaryWorkShift.Index).Value, _
                                    CDate(Me.Buffer).Date, Row.Cells(Me.ColPrimaryEntity.Index).Value) > 0 Then
                                        'MsgBox("Data cannot edit because this period used in transaction attendance.", MsgBoxStyle.Information, "Information")
                                        MsgBox("Data cannot edit because this period: " + Format(CDate(Me.Buffer).Date, "dd MMM yyyy").ToString + " used in transaction overtime order no. " + Me.ObjWork.GetNomerOvertimeOrder(Row.Cells(Me.ColPrimaryWorkShift.Index).Value, CDate(Me.Buffer).Date, Row.Cells(Me.ColPrimaryEntity.Index).Value).ToString + ".", MsgBoxStyle.Information, "Information")
                                        Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        My.Application.ChangeCulture("en-US")
                                        My.Application.ChangeUICulture("en-US")
                                        Exit Sub
                                    End If
                                End If

                                '2. Geser ke Luar
                                If CDate(Row.Cells(Me.ColEndDate.Index).Value).Date > CDate(Me.Buffer).Date Then
                                    'My.Application.ChangeCulture("en-US")
                                    'My.Application.ChangeUICulture("en-US")
                                    If Me.ObjWork.GetJmlBatasanGeserLuarEndDate(CDate(Row.Cells(Me.ColEndDate.Index).Value).Date, _
                                    CDate(Me.Buffer).Date, CDbl(Row.Cells(Me.ColPrimary.Index).Value), Row.Cells(Me.ColPrimaryEntity.Index).Value) > 0 Then
                                        'My.Application.ChangeCulture("id-ID")
                                        'My.Application.ChangeUICulture("id-ID")
                                        MsgBox("Data cannot edit because this period used in other allocation.", MsgBoxStyle.Information, "Information")
                                        Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        My.Application.ChangeCulture("en-US")
                                        My.Application.ChangeUICulture("en-US")
                                        Exit Sub
                                    End If
                                    'My.Application.ChangeCulture("id-ID")
                                    'My.Application.ChangeUICulture("id-ID")
                                End If
                                '3. Geser ke dalam
                                If CDate(Row.Cells(Me.ColEndDate.Index).Value).Date < CDate(Me.Buffer).Date Then
                                    If Me.ObjWork.GetJmlAttendance(Row.Cells(Me.ColPrimaryWorkShift.Index).Value, _
                                    CDate(Row.Cells(Me.ColEndDate.Index).Value).Date + TimeSpan.FromDays(1), Row.Cells(Me.ColPrimaryEntity.Index).Value) > 0 Then 'aul perubahan - CDate(Row.Cells(Me.ColEndDate.Index).Value).Date + TimeSpan.FromDays(1)
                                        'MsgBox("Data cannot edit because this period used in transaction attendance.", MsgBoxStyle.Information, "Information")
                                        MsgBox("Data cannot edit because this period: " + Format(CDate(Row.Cells(Me.ColEndDate.Index).Value).Date, "dd MMM yyyy").ToString + " used in transaction attendance no. " + Me.ObjWork.GetNomerAttendance(Row.Cells(Me.ColPrimaryWorkShift.Index).Value, CDate(Row.Cells(Me.ColEndDate.Index).Value).Date, Row.Cells(Me.ColPrimaryEntity.Index).Value).ToString + ".", MsgBoxStyle.Information, "Information")
                                        Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        My.Application.ChangeCulture("en-US")
                                        My.Application.ChangeUICulture("en-US")
                                        Exit Sub
                                    End If

                                    'cek juga ke overtime order
                                    If Me.ObjWork.GetJmlOvertimeOrder(Row.Cells(Me.ColPrimaryWorkShift.Index).Value, _
                                    CDate(Row.Cells(Me.ColEndDate.Index).Value).Date + TimeSpan.FromDays(1), Row.Cells(Me.ColPrimaryEntity.Index).Value) > 0 Then 'aul perubahan - CDate(Row.Cells(Me.ColEndDate.Index).Value).Date + TimeSpan.FromDays(1)
                                        'MsgBox("Data cannot edit because this period used in transaction attendance.", MsgBoxStyle.Information, "Information")
                                        MsgBox("Data cannot edit because this period: " + Format(CDate(Row.Cells(Me.ColEndDate.Index).Value).Date, "dd MMM yyyy").ToString + " used in transaction overtime order no. " + Me.ObjWork.GetNomerOvertimeOrder(Row.Cells(Me.ColPrimaryWorkShift.Index).Value, CDate(Row.Cells(Me.ColEndDate.Index).Value).Date, Row.Cells(Me.ColPrimaryEntity.Index).Value).ToString + ".", MsgBoxStyle.Information, "Information")
                                        Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        My.Application.ChangeCulture("en-US")
                                        My.Application.ChangeUICulture("en-US")
                                        Exit Sub
                                    End If

                                    If CDate(Row.Cells(Me.ColEndDate.Index).Value).Date + TimeSpan.FromDays(1) <> CDate(Me.Buffer).Date - TimeSpan.FromDays(1) Then
                                        'My.Application.ChangeCulture("en-US")
                                        'My.Application.ChangeUICulture("en-US")
                                        'Me.ObjWork.FungsiFillDataAttendanceItem(TBLEnd, CDate(Row.Cells(Me.ColStartDate.Index).Value).Date + TimeSpan.FromDays(1), CDate(Me.Buffer).Date - TimeSpan.FromDays(1), Row.Cells(Me.ColPrimaryEntity.Index).Value)
                                        'aul perubahan - CDate(Row.Cells(Me.ColEndDate.Index).Value).Date
                                        Me.ObjWork.FungsiFillDataAttendanceItem(TBLEnd, CDate(Row.Cells(Me.ColEndDate.Index).Value).Date + TimeSpan.FromDays(1), CDate(Me.Buffer).Date - TimeSpan.FromDays(1), Row.Cells(Me.ColPrimaryEntity.Index).Value)
                                        If TBLEnd.Rows.Count > 0 Then
                                            Dim ik_End_No As Integer = 0
                                            Dim ik_End_Idx As Integer = 0
                                            Dim ik_End_Retval As String = ""
                                            While ik_End_Idx <= TBLEnd.Count - 1
                                                RowEnd = TBLEnd.Rows(ik_End_Idx)
                                                ik_End_No += 1
                                                If ik_End_Retval = "" Then
                                                    ik_End_Retval = ik_End_No.ToString + ". " + RowEnd.no_transaksi.ToString
                                                Else
                                                    ik_End_Retval += vbCrLf + ik_End_No.ToString + ". " + RowEnd.no_transaksi.ToString
                                                End If
                                                ik_End_Idx += 1
                                            End While
                                            'My.Application.ChangeCulture("id-ID")
                                            'My.Application.ChangeUICulture("id-ID")
                                            MsgBox("Period cannot edit because in period " + Format(CDate(Row.Cells(Me.ColEndDate.Index).Value).Date + TimeSpan.FromDays(1), "dd MMM yyyy") + " to " + _
                                            Format(CDate(Me.Buffer).Date - TimeSpan.FromDays(1), "dd MMM yyyy") + _
                                            " is refering to attendance. " + vbCrLf + ik_End_Retval.ToString, MsgBoxStyle.Information, "Information")
                                            Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                            My.Application.ChangeCulture("en-US")
                                            My.Application.ChangeUICulture("en-US")
                                            Exit Sub
                                        End If
                                        My.Application.ChangeCulture("id-ID")
                                        My.Application.ChangeUICulture("id-ID")
                                    End If
                                End If
                                'data benar
                                Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")
                                If Me.ObjWork.ApakahDataBenar_ItemAllocate_Save(Row, Me) = True Then
                                    Row.Cells(Me.ColSave.Index).Value = "save"
                                    Me.SimpanRowWorkshift(Row)
                                    'Me.BtnFilter.PerformClick()
                                    Me.BackgroundWorker6.RunWorkerAsync()
                                End If
                            End If
                        Else
                            'data salah
                            Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        My.Application.ChangeCulture("en-US")
                        My.Application.ChangeUICulture("en-US")
                    End If
                    If e.ColumnIndex = Me.ColWorkshiftCode.Index Then
                        Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjWork.ApakahSudahAda(Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value, ClsMasterWorkShift.EnumTipeWorkshift.HR) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            Index = Me.GetIndexItemInputWorkshift(Me.ObjWork.GetPrimary(Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value, ClsMasterWorkShift.EnumTipeWorkshift.HR), Row.Cells(Me.ColPrimaryEntity.Index).Value, _
                            Row.Cells(Me.ColStartDate.Index).Value, Row.Cells(Me.ColEndDate.Index).Value)
                            If Index < 0 Then
                                Me.DgvWorkshift.Item(Me.ColPrimaryWorkShift.Index, e.RowIndex).Value = Me.ObjWork.GetPrimary(Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value, ClsMasterWorkShift.EnumTipeWorkshift.HR)
                                Me.DgvWorkshift.Item(Me.ColWorkshiftCode.Index, e.RowIndex).Value = Me.ObjWork.GetCodeByPrimary(Me.DgvWorkshift.Item(Me.ColPrimaryWorkShift.Index, e.RowIndex).Value)
                                Me.DgvWorkshift.Item(Me.ColWorkShiftName.Index, e.RowIndex).Value = Me.ObjWork.GetDescription(Me.DgvWorkshift.Item(Me.ColPrimaryWorkShift.Index, e.RowIndex).Value)
                            Else
                                If e.RowIndex <> Index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvWorkshift.Item(Me.ColWorkshiftCode.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex_WS = e.RowIndex
                                        Me.BackgroundWorker4.RunWorkerAsync()
                                    Else
                                        Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCekData.SetNumber(Me.dgvItem)
                                    Exit Sub
                                End If
                            End If
                        Else
                            '2. kode akun yang diketik tidak ada --> tampilkan form cari
                            Dim ij As New Core_Module.FrmMasterWorkShift_Cari(ClsMasterWorkShift.EnumTipeWorkshift.HR)
                            ij.txtCari.Text = Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryWorkshift <> -1 Then
                                '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                                Index = Me.GetIndexItemInputWorkshift(ij.PrimaryWorkshift, Row.Cells(Me.ColPrimaryEntity.Index).Value, _
                                Row.Cells(Me.ColStartDate.Index).Value, Row.Cells(Me.ColEndDate.Index).Value)
                                If Index < 0 Then
                                    Me.DgvWorkshift.Item(Me.ColWorkshiftCode.Index, e.RowIndex).Value = Me.ObjWork.GetCodeByPrimary(ij.PrimaryWorkshift)
                                    Me.DgvWorkshift.Item(Me.ColWorkShiftName.Index, e.RowIndex).Value = Me.ObjWork.GetDescription(ij.PrimaryWorkshift)
                                    Me.DgvWorkshift.Item(Me.ColPrimaryWorkShift.Index, e.RowIndex).Value = ij.PrimaryWorkshift
                                Else
                                    If e.RowIndex <> Index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex_WS = e.RowIndex
                                            Me.BackgroundWorker4.RunWorkerAsync()
                                        Else
                                            Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCekData.SetNumber(Me.dgvItem)
                                        Exit Sub
                                    Else
                                        Me.DgvWorkshift.Item(Me.ColWorkshiftCode.Index, e.RowIndex).Value = Me.ObjWork.GetCodeByPrimary(ij.PrimaryWorkshift)
                                        Me.DgvWorkshift.Item(Me.ColWorkShiftName.Index, e.RowIndex).Value = Me.ObjWork.GetDescription(ij.PrimaryWorkshift)
                                        Me.DgvWorkshift.Item(Me.ColPrimaryWorkShift.Index, e.RowIndex).Value = ij.PrimaryWorkshift
                                    End If
                                End If
                            ElseIf ij.PrimaryWorkshift = -1 Then
                                If Me.DgvWorkshift.Rows.Count - 1 <> e.RowIndex Then
                                    Me.EraseRowIndex_WS = e.RowIndex
                                    Row.Cells(Me.ColWorkshiftCode.Index).Value = Me.ObjWork.GetCodeByPrimary(Row.Cells(Me.ColPrimaryWorkShift.Index).Value)
                                    'Me.BackgroundWorker4.RunWorkerAsync()
                                Else
                                    Row.Cells(Me.ColWorkshiftCode.Index).Value = Me.ObjWork.GetCodeByPrimary(Row.Cells(Me.ColPrimaryWorkShift.Index).Value)
                                End If
                            End If
                            ij = Nothing
                        End If
                        If Me.ObjWork.ApakahDataBenar_ItemAllocate_Save(Row, Me) = True Then
                            Row.Cells(Me.ColSave.Index).Value = "save"
                            Me.SimpanRowWorkshift(Row)
                        End If
                    End If
                    ''
                Else
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.DgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLStart.Dispose()
            TBLEnd.Dispose()
        End Try
    End Sub

    Private Sub DgvWorkshift_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles DgvWorkshift.UserDeletingRow
        Dim Row As Core_Module.DseMasterWorkshift_Allocation.Output_DGVRow
        Dim Filter As String = ""
        Dim idx As Integer = 0
        Dim TBL1 As New DseMasterWorkshift_Allocation.tbltrans_attendance_headerDataTable
        Dim Row1 As DseMasterWorkshift_Allocation.tbltrans_attendance_headerRow
        Dim Nomor As Integer = 1
        Dim Retval As String = ""
        Try
            If Me.DgvWorkshift.Rows.Count > 0 Then
                '1. Tidak boleh hapus kalau sudah ada transaksi attendance 
                Me.ObjWork.FungsiFillDataAttendanceItem(TBL1, CDate(e.Row.Cells(Me.ColStartDate.Index).Value), _
                CDate(e.Row.Cells(Me.ColEndDate.Index).Value), e.Row.Cells(Me.ColPrimaryEntity.Index).Value)
                If TBL1.Rows.Count > 0 Then
                    For Each Row1 In TBL1.Rows
                        If Retval = "" Then
                            Retval = Nomor + ". " + Row1.no_transaksi.ToString
                        Else
                            Retval += vbCrLf + Nomor + ". " + Row1.no_transaksi.ToString
                        End If
                        Nomor += 1
                    Next
                    MsgBox("Workshift code cannot deleted because in period " + _
                    Format(CDate(e.Row.Cells(Me.ColStartDate.Index).Value), "dd MMM yyyy") + _
                    " to " + Format(CDate(e.Row.Cells(Me.ColEndDate.Index).Value), "dd MMM yyyy") + _
                    " is refering to attendance. " + vbCrLf + Retval.ToString, MsgBoxStyle.Information, "Information")
                    e.Cancel = True
                End If
                '2. ketika DGVWorkshift dihapus, maka data di TBLOutputDGV ikut terhapus
                Filter = "[Primary] = " + e.Row.Cells(Me.ColPrimary.Index).Value.ToString + _
                "AND [Primary Entity] = " + e.Row.Cells(Me.ColPrimaryEntity.Index).Value.ToString + ""
                Me.TBlOutputDGV.DefaultView.RowFilter = Filter
                If TBlOutputDGV.DefaultView.Count > 0 Then
                    Row = Me.TBlOutputDGV.DefaultView(0).Row
                    Me.TBlOutputDGV.Rows.Remove(Row)
                End If
                Me.TBlOutputDGV.DefaultView.RowFilter = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBL1.Dispose()
        End Try
    End Sub

    Private Sub BackgroundWorker4_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker4.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker4_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker4.RunWorkerCompleted
        Try
            Me.DgvWorkshift.Rows.RemoveAt(Me.EraseRowIndex)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker5_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker5.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker5_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker5.RunWorkerCompleted
        Dim Row As Core_Module.DseMasterWorkshift_Allocation.Output_DGVRow
        Dim Filter As String = ""
        Try
            '1. ketika DGVWorkshift dihapus, maka data di TBLOutputDGV ikut terhapus
            Filter = "[Primary] = " + Me.DgvWorkshift.Item(Me.ColPrimary.Index, Me.EraseRowIndex_Del).Value.ToString + _
            "AND [Primary Entity] = " + Me.DgvWorkshift.Item(Me.ColPrimaryEntity.Index, Me.EraseRowIndex_Del).Value.ToString + ""
            Me.TBlOutputDGV.DefaultView.RowFilter = Filter
            If TBlOutputDGV.DefaultView.Count > 0 Then
                Row = Me.TBlOutputDGV.DefaultView(0).Row
                Me.TBlOutputDGV.Rows.Remove(Row)
            End If
            Me.TBlOutputDGV.DefaultView.RowFilter = ""
            '2. Hapus Baris DGV Allocate
            Me.DgvWorkshift.Rows.RemoveAt(Me.EraseRowIndex_Del)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker_CloseForm_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_CloseForm.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker_CloseForm_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker_CloseForm.RunWorkerCompleted
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnData.Click
        Dim ij As New FrmMasterWorkShift_Allocate_Data01
        Try
            ij.TBL = Me.TBlOutputDGV
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub BackgroundWorker6_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker6.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker6_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker6.RunWorkerCompleted
        Try
            Me.BtnFilter.PerformClick()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class