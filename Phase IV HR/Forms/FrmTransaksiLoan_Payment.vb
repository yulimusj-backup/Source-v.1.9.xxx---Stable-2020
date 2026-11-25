Option Explicit On
Option Strict Off

Imports Core_Module, System.Drawing, System.Windows.Forms
Public Class FrmTransaksiLoan_Payment
    Public PrimaryCashIn As String = ""
    Public PrimaryGain As String = ""
    Public PrimaryLoss As String = ""
    Public UserCode As String = ""
    Public EnumDipanggilDari As ClsTransaksiLoanPayment.EnumFrmTransLoanPaymentDari = ClsTransaksiLoanPayment.EnumFrmTransLoanPaymentDari.FrmTransLoanPayment
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public TBLLoan As New DseTransaksiLoanPayment.DataTable2DataTable
    Public TBLCBOItem As New DseTransaksiLoanPayment.DataTable2DataTable
    Public vFromSalaryCalc As Boolean = False

    Public ShowFromSalary As Boolean = False
    Public PrimaryLoanPaymentFromSalary As Double = -1
    Public ShowFromLinkLabel As Boolean = False
    Public PengaturanFocusDataBenar As Double = -1
    Public Baru As Boolean = False
    Public Obj As New ClsTransaksiLoanPayment
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Public Def_Prec As Double = Me.ObjCurr.GetDecimal_DefaultCurrency
    Public TotalLoanInstallment As Decimal = 0
    Public TBLLoanStatus As New DseTransaksiLoanPayment.TBLUpdateStatus_LoanDataTable
    Public ReadOnlyFromSalary As Boolean = False

    Dim ObjClosing As New AccModule.ClsTransaksiClosing
    Dim ObjItem As New ClsTransaksiLoanPayment_Item
    Dim ObjLoan As New ClsTransaksiLoan
    Dim ObjSalary As New ClsTransaksiSalary
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjAkun As New AccModule.ClsMasterAkun
    Dim PrimaryCurr_Default As Double = 1
    Dim ObjKrywn As New Core_Module.ClsMasterKaryawan

    Dim Buffer As String = ""
    Dim ApakahSimpanData As Boolean = False
    Dim TBLCash As New DseTransaksiLoanPayment.tblmaster_akunDataTable
    Dim ObjSetting As New ClsTransaksiSetting

    Dim objJournal As New AccModule.ClsTransaksiJournal

    Public Enum EnumListFocus As Integer
        GaAda
        EmployeeCode
        GainAcc
        LossAcc
        Remarks
        DGVItem
    End Enum

    Public Sub Fungsi_Load()
        Dim RowDGV As DataGridViewRow
        Try
            If Me.ShowFromLinkLabel = True Then
                If Me.dgvItem.Rows.Count > 0 Then
                    Me.TotalLoanInstallment = 0
                    For Each RowDGV In Me.dgvItem.Rows
                        TotalLoanInstallment += CDec(RowDGV.Cells(Me.ColPayment.Index).Value)
                    Next
                End If
                Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
                If Me.ReadOnlyFromSalary = True Then
                    Me.gboDetail.Enabled = False
                Else
                    Me.gboDetail.Enabled = True
                End If
                Exit Sub
            End If
            If Me.ShowFromSalary = True Then
                If Me.ReadOnlyFromSalary = True Then
                    Me.gboDetail.Enabled = False
                Else
                    Me.gboDetail.Enabled = True
                End If
                If Me.dgvItem.RowCount <= 0 Then
                    Me.Obj = New ClsTransaksiLoanPayment
                    'Me.Baru = True
                    Me.KosongkanField()
                    Me.NonAktifkanDetail()
                    Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
                    If Me.Baru = False Then
                        Me.FungsiBtnFind()
                        Me.FungsiBtnEdit()
                    Else
                        Me.FungsiBtnNew()
                    End If
                End If
                If Me.PengaturanFocusDataBenar = -1 Then
                    Me.TxtPayment.Focus()
                ElseIf Me.PengaturanFocusDataBenar = EnumListFocus.EmployeeCode Then
                    Me.TxtEmployeeCode1.Select()
                    Me.TxtEmployeeCode1.Focus()
                ElseIf Me.PengaturanFocusDataBenar = EnumListFocus.GainAcc Then
                    Me.TxtGainAccount.Select()
                    Me.TxtGainAccount.Focus()
                ElseIf Me.PengaturanFocusDataBenar = EnumListFocus.LossAcc Then
                    Me.TxtLossAccount.Select()
                    Me.TxtLossAccount.Focus()
                ElseIf Me.PengaturanFocusDataBenar = EnumListFocus.Remarks Then
                    Me.TxtRemark.Select()
                    Me.TxtRemark.Focus()
                ElseIf Me.PengaturanFocusDataBenar = EnumListFocus.DGVItem Then
                    Me.dgvItem.Focus()
                End If
                Me.txtJournalNo.BackColor = Me.TxtReadOnly.BackColor
                Me.txtJournalNo.ReadOnly = True
                Me.TxtJournalDesc.BackColor = Me.TxtReadOnly.BackColor
                Me.TxtJournalDesc.ReadOnly = True
            Else
                Me.Obj = New ClsTransaksiLoanPayment
                Me.Baru = True
                Me.KosongkanField()
                Me.NonAktifkanDetail()
                Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiTicketing_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.ShowIcon = False
            Me.Fungsi_Load()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FungsiBtnOK_Salary() As Boolean
        Try
            If Me.ApakahDataBenar(Me.ShowFromSalary) = True Then
                If Me.ApakahSimpanData = True Then
                    Me.SalinKeObjek()
                    If Baru = True Then
                        'konfirmasi disimpan atau tidak
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.Obj.SimpanBaru1(Me.UserCode, Me.ObjCD, Me.TxtPrimaryLoanPayment.Text, Me.dgvItem, Me) = True Then
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            'Me.btnFind.Focus()
                        Else
                            Return False
                        End If
                    Else
                        'konfirmasi disimpan atau tidak
                        If Me.Obj.SimpanPerubahan1(Me.UserCode, Me.TxtPrimaryLoanPayment.Text, Me.ObjCD, Me.dgvItem, Me) = True Then
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            'Me.btnEdit.Focus()
                        Else
                            Return False
                        End If
                        Me.FungsiBtnEdit()
                    End If
                Else
                    'Jika posisi edit, dan semua DGV Item Loan Payment, kosong maka transaksi loan dihapus
                    If Me.Baru = False Then
                        If Me.dgvItem.Rows.Count <= 0 Then
                            If Me.TxtPrimaryLoanPayment.Text <> "-1" Then
                                If Me.Obj.SimpanPerubahan1(Me.UserCode, Me.TxtPrimaryLoanPayment.Text, Me.ObjCD, Me.dgvItem, Me) = True Then
                                    Me.SalinKeField()
                                    Me.NonAktifkanDetail()
                                    Me.Obj.Hapus1(Me.TxtPrimaryLoanPayment.Text, Me.UserCode, True, True, Me.dgvItem, Me)
                                    'Update data primary LoanPayment di Transaksi Salary
                                    Dim PrimaryDefault As Double = -1
                                    PrimaryDefault = -1
                                    Me.ObjSalary.UpdatePrimaryLoanPayment(PrimaryDefault, Me.TxtPrimarySalary.Text)
                                    Me.KosongkanField()
                                Else
                                    Return False
                                End If
                                Me.FungsiBtnEdit()
                            End If
                        End If
                    End If
                End If
            Else
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub FungsiBtnOK()
        Dim hasil As MsgBoxResult
        Try
            If Me.ShowFromSalary = True Then
                Me.Hide()
                If Me.ShowFromLinkLabel = True Then
                    Me.ShowFromLinkLabel = False
                End If
                If Me.PengaturanFocusDataBenar <> -1 Then
                    Me.PengaturanFocusDataBenar = -1
                End If
                Exit Sub
            End If
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If Baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.Obj.SimpanBaru(Me.UserCode, Me.ObjCD, Me.TxtPrimaryLoanPayment.Text, Me.dgvItem, Me) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.SimpanPerubahan(Me.UserCode, Me.ObjCD, Me.TxtPrimaryLoanPayment.Text, Me.dgvItem, Me) = True Then
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

    'Kontrol Button
    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            Me.FungsiBtnOK()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FungsiBtnCancel()
        Dim hasil As MsgBoxResult
        Try
            If Not Me.TxtDocumentNo.Text = "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Baru = True Then
                        Me.KosongkanField()
                        Me.TxtDocumentNo.Clear()
                    Else
                        Me.Obj.CariDetail(Me.TxtPrimaryLoanPayment.Text, Me.ObjCD, _
                        Me.dgvItem, Me.Def_Prec, Me)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.KosongkanField()
            End If
            Me.NonAktifkanDetail()
            Me.btnNew.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.FungsiBtnCancel()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FungsiBtnNew()
        Dim AdaFlightService As Boolean = False
        Try
            Me.Baru = True
            Obj.ResetObject(Me.dgvItem)
            Me.AktifkanDetail()
            Me.KosongkanField()
            Me.TxtDocumentNo.Text = Me.Obj.GetNoBaru(Me.dtpDate.Value, Me.ObjCD)
            Me.Obj.IsiTBLCashCode(Me.CboCashAccount, Me.PrimaryCurr_Default, Me.TBLCash, Me)
            Me.CboCashAccount.SelectedIndex = 0
            If Me.PrimaryCashIn = -1 Then
                Me.PrimaryCashIn = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, ClsTransaksiSetting.EnumNamaParameter.Default_LoanPayment_PrimaryCashInAccount)
            End If
            If Me.PrimaryGain = -1 Then
                Me.PrimaryGain = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, ClsTransaksiSetting.EnumNamaParameter.Default_LoanPayment_PrimaryGainAccount)
            End If
            If Me.PrimaryLoss = -1 Then
                Me.PrimaryLoss = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, ClsTransaksiSetting.EnumNamaParameter.Default_LoanPayment_PrimaryLossAccount)
            End If
            If String.IsNullOrEmpty(Me.PrimaryGain) = True Then
                Me.TxtPrimaryGainAcc.Text = -1
                Me.TxtGainAccount.Clear()
                Me.TxtDescGain.Clear()
            Else
                Me.TxtPrimaryGainAcc.Text = PrimaryGain
                Me.TxtGainAccount.Text = Me.ObjAkun.GetAccountCode(PrimaryGain)
                Me.TxtDescGain.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(PrimaryGain)
            End If
            If String.IsNullOrEmpty(PrimaryLoss) = True Then
                Me.TxtPrimaryLoss.Text = -1
                Me.TxtLossAccount.Clear()
                Me.TxtDescLoss.Clear()
            Else
                Me.TxtPrimaryLoss.Text = PrimaryLoss
                Me.TxtLossAccount.Text = Me.ObjAkun.GetAccountCode(PrimaryLoss)
                Me.TxtDescLoss.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(PrimaryLoss)
            End If
            If String.IsNullOrEmpty(PrimaryCashIn) = False Then
                Me.CboCashAccount.Text = Me.ObjAkun.GetAccountCode(PrimaryCashIn)
            End If
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.dtpDate.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.FungsiBtnNew()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FungsiBtnFind()
        Dim ijk_N As New FrmTransaksiLoan_Payment_Cari
        Dim ijk As New FrmTransaksiLoan_Payment_Cari12
        Try
            If Me.ShowFromSalary = False Then
                ijk.BackColorAll = Me.BackColorAll
                ijk.cboCriteria.Items.Clear()
                ijk.txtCari.Text = ""
                ijk.dgvHasil.DataSource = ""
                ijk.ShowDialog()

                If Not ijk.mPrimary = -1 Then
                    If Me.Obj.CariDetail(ijk.mPrimary, Me.ObjCD, Me.dgvItem, Me.Def_Prec, Me) = True Then
                        Me.SalinKeField()
                        Me.btnEdit.Focus()
                    Else
                        Me.btnFind.Focus()
                    End If
                    Me.btnFind.Focus()
                End If
            Else
                ijk_N.mPrimary = Me.PrimaryLoanPaymentFromSalary
                If Not ijk_N.mPrimary = -1 Then
                    If Me.Obj.CariDetail(ijk_N.mPrimary, Me.ObjCD, Me.dgvItem, Me.Def_Prec, Me) = True Then
                        Me.SalinKeField()
                        Me.btnEdit.Focus()
                    Else
                        Me.btnFind.Focus()
                    End If
                    Me.btnFind.Focus()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ijk.Dispose()
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Try
            Me.FungsiBtnFind()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FungsiBtnEdit()
        Try
            If Me.TxtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtDocumentNo.Text <> "" Then
                If Me.Obj.ApakahBolehEdit(Me.TxtPrimaryLoanPayment.Text, Me.UserCode, True, True, Me.vFromSalaryCalc, Me.EnumDipanggilDari) = True Then
                    Me.Baru = False
                    Me.AktifkanDetail()
                    If Me.TxtRemark.Enabled = True Then
                        If Me.TxtPrimaryEmployee.Text <> "-1" Then
                            Me.Obj.IsiTBLLoanPayment_OnGoing(Me.TBLLoan, Me.TxtPrimaryEmployee.Text, Me.Baru, _
                            Me.TxtPrimaryLoanPayment.Text, Me.dgvItem, Me)
                            Me.Obj.IsiCBOLoanPayment(Me.TBLLoan, Me.ObjCD, Me.Def_Prec, Me, Me.TBLCBOItem)
                        End If
                        Me.dtpDate.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            Me.FungsiBtnEdit()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FungsiBtnHapus_Salary() As Boolean
        Try
            If Me.TxtPrimaryLoanPayment.Text <> "-1" Then
                If Me.Obj.ApakahBolehHapus(Me.TxtPrimaryLoanPayment.Text, Me.UserCode, True, True, Me.dgvItem, Me, Me.vFromSalaryCalc, Me.EnumDipanggilDari) = False Then
                    'me.btnFind.Focus 
                    Return False
                Else
                    If Me.Obj.Hapus1(Me.TxtPrimaryLoanPayment.Text, Me.UserCode, True, True, Me.dgvItem, _
                     Me) = True Then
                        'Me.KosongkanField()
                        'Me.btnNew.Focus()
                    Else
                        Return False
                    End If
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub FungsiBtnHapus()
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtDocumentNo.Text <> "" Then
                If Me.Obj.ApakahBolehHapus(Me.TxtPrimaryLoanPayment.Text, Me.UserCode, True, True, Me.dgvItem, Me, Me.vFromSalaryCalc, Me.EnumDipanggilDari) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.Hapus(Me.UserCode, Me.ObjCD, Me.TxtPrimaryLoanPayment.Text, Me.dgvItem, _
                         Me, True) = True Then
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

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            Me.FungsiBtnHapus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.ShowFromSalary = False Then
                If Me.TxtDocumentNo.Enabled = True Then
                    Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ApakahDataBenar() = True Then
                            Me.SalinKeObjek()
                            If Baru = True Then
                                'simpan baru
                                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                                If Me.Obj.SimpanBaru(Me.UserCode, Me.ObjCD, Me.TxtPrimaryLoanPayment.Text, _
                                Me.dgvItem, Me) = True Then
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
                                If Me.Obj.SimpanPerubahan(Me.UserCode, Me.ObjCD, Me.TxtPrimaryLoanPayment.Text, Me.dgvItem, Me) = True Then
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
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'GetNoBaru
    Private Sub dtpDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.Buffer = Me.TxtDocumentNo.Text
                If Me.Baru = True Then
                    Me.TxtDocumentNo.Text = Me.Obj.GetNoBaru(Me.dtpDate.Value, Me.ObjCD)
                Else
                    If String.IsNullOrEmpty(Me.TxtPrimaryLoanPayment.Text) = True Then
                        Me.TxtPrimaryLoanPayment.Text = "-1"
                    End If
                    If Me.dtpDate.Value.Month = Me.Obj.GetTanggalTransaksi(Me.TxtPrimaryLoanPayment.Text).Month And _
                    Me.dtpDate.Value.Year = Obj.GetTanggalTransaksi(Me.TxtPrimaryLoanPayment.Text).Year Then
                        Me.TxtDocumentNo.Text = Me.Obj.GetNoTransaksi(Me.TxtPrimaryLoanPayment.Text)
                    Else
                        If Me.dtpDate.Value.Month <> Me.Obj.TanggalTransaksi.Month Or _
                        Me.dtpDate.Value.Year <> Me.Obj.TanggalTransaksi.Year Then
                            Me.TxtDocumentNo.Text = Me.Obj.GetNoBaru(Me.dtpDate.Value, Me.ObjCD)
                        Else
                            Me.TxtDocumentNo.Text = Me.Obj.GetNoTransaksi(Me.TxtPrimaryLoanPayment.Text)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''Fungsi Standar
    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.GboItem.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer4.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer4.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.Controls, True, True, True, True, True, True, True, True)
            Me.CancelButton = Me.btnExit
            Me.BackColorDefault()

            Me.dgvItem.ReadOnly = True
            Me.dgvItem.AllowUserToDeleteRows = False

            Me.KosongkanFieldItem()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.GboItem.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer4.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer4.Panel2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False)

            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.BackColorReset()

            'set dgv
            Me.dgvItem.ReadOnly = False
            Me.dgvItem.AllowUserToDeleteRows = True

            If Me.dgvItem.SelectedRows.Count = 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If

            'Posisi Edit
            If Me.Baru = False Then
                Me.TxtEmployeeCode1.Enabled = False
                Me.BtnFindEmployee.Enabled = False
            End If

            If Me.ShowFromSalary = True Then
                Me.btnCancel.Enabled = False
                Me.dtpDate.Enabled = False
                Me.txtJournalNo.BackColor = Me.TxtReadOnly.BackColor
                Me.txtJournalNo.ReadOnly = True
                Me.TxtJournalDesc.BackColor = Me.TxtReadOnly.BackColor
                Me.TxtJournalDesc.ReadOnly = True
            End If

            If Me.ShowFromLinkLabel = True Then
                Me.dtpDate.Enabled = False
            End If

            'Untuk TBLLoanStatus, untuk settingan update status
            Me.TBLLoanStatus.Rows.Clear()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorReset()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.GboItem.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel2.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Color.White, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorDefault()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.GboItem.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel2.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtReadOnly.BackColor, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub KosongkanField()
        Try
            Me.ObjCD.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            Me.ObjCD.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, True, False)
            Me.ObjCD.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, True, False)
            Me.ObjCD.KosongkanControl(Me.SplitContainer4.Panel1.Controls, True, True, False)
            Me.ObjCD.KosongkanControl(Me.SplitContainer4.Panel2.Controls, True, True, False)
            Me.TxtPrimaryLoanPayment.Text = "-1"
            Me.TxtPrimarySalary.Text = "-1"
            Me.TxtPrimaryJurnal.Text = "-1"
            Me.TxtPrimaryEmployee.Text = "-1"
            Me.TxtPrimaryCashAcc.Text = "-1"
            Me.TxtPrimaryGainAcc.Text = "-1"
            Me.TxtPrimaryLoss.Text = "-1"
            Me.CboCashAccount.DataSource = Nothing
            Me.CboCashAccount.Items.Clear()
            If String.IsNullOrEmpty(Me.PrimaryCashIn) = True Then
                Me.PrimaryCashIn = -1
            End If
            If String.IsNullOrEmpty(Me.PrimaryGain) = True Then
                Me.PrimaryGain = -1
            End If
            If String.IsNullOrEmpty(Me.PrimaryLoss) = True Then
                Me.PrimaryLoss = -1
            End If
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.KosongkanFieldItem()
            Me.dgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanFieldItem()
        Try
            Me.ObjCD.KosongkanControl(Me.GboItem.Controls, True, True, False)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, False)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, False)
            Me.TxtPrimaryLoan.Text = "-1"
            Me.TxtIdentifier.Clear()
            Me.CboLoanNumber.DataSource = Nothing
            Me.CboLoanNumber.Items.Clear()
            Me.TxtInstallmentAt.Text = "0"
            Me.TxtBase.Text = "0"
            Me.TxtInterest.Text = "0"
            Me.TxtTotalInstallment.Text = "0"
            Me.TxtPayment.Text = "0"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function HapusSpasi(ByVal Data As String) As String
        Try
            Do While Data.EndsWith(" ") And Data.Length > 0
                Data = Data.Remove(Data.Length - 1, 1)
            Loop
            Return Data
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function ApakahDataBenar(Optional ByVal vShowFromSalary As Boolean = False) As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim i As Int16
        Dim data As String
        Dim AdaDataKosong As Boolean = False
        Dim PayrollConventionDate As Date = Me.ObjClosing.GetPayrollConvertionDate()
        Dim DocNoBuffer As String = ""
        Try
            If Me.ShowFromSalary = True Then
                Me.ApakahSimpanData = False
                If Me.dgvItem.Rows.Count <= 0 Then
                    Me.ApakahSimpanData = False
                    Return True
                Else
                    Me.ApakahSimpanData = True
                End If
            End If
            '1. cek data yang harus diisi
            KoleksiNamaData.Add("Employee Code")
            KoleksiTextBox.Add(Me.TxtEmployeeCode1)
            KoleksiNamaData.Add("Gain Account")
            KoleksiTextBox.Add(Me.TxtGainAccount)
            KoleksiNamaData.Add("Loss Account")
            KoleksiTextBox.Add(Me.TxtLossAccount)
            If Me.ShowFromSalary = False Then
                KoleksiNamaData.Add("Journal No")
                KoleksiTextBox.Add(Me.txtJournalNo)
                KoleksiNamaData.Add("Journal Description")
                KoleksiTextBox.Add(Me.TxtJournalDesc)
            End If


            If Me.ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                If vShowFromSalary = True Then
                    i = 0
                    AdaDataKosong = False
                    Do While i < KoleksiTextBox.Count
                        'salin data
                        data = KoleksiTextBox.Item(i + 1).text
                        data = Me.HapusSpasi(data)

                        'cek data
                        If data.Length = 0 Then
                            KoleksiTextBox.Item(i + 1).focus()
                            AdaDataKosong = True
                            Exit Do
                        End If
                        'cek berkaitan dengan karakter yang tidak diingikan --> ', *, %
                        If data.Contains("'") = True Or data.Contains("*") = True Or data.Contains("%") = True Then
                            AdaDataKosong = True
                            Exit Do
                        End If
                        i += 1
                    Loop
                    If AdaDataKosong = True Then
                        Me.PengaturanFocusDataBenar = i + 1
                    End If
                    Me.ShowDialog()
                End If
                Return False
            End If

            '2. Fungsi Alt+O Btn Find
            If Me.TxtPrimaryEmployee.Text = -1 Then
                If Me.TxtEmployeeCode1.Text <> "" Then
                    If ObjKrywn.ApakahSudahAda(Me.TxtEmployeeCode1.Text) = False Then
                        Me.BtnFindEmployee.PerformClick()
                    End If
                End If
            End If
            If Me.TxtPrimaryGainAcc.Text = -1 Then
                If Me.TxtGainAccount.Text <> "" Then
                    If ObjAkun.ApakahSudahAda(Me.TxtGainAccount.Text) = False Then
                        Me.BtnFindGain.PerformClick()
                    End If
                End If
            End If
            If Me.TxtPrimaryLoss.Text = -1 Then
                If Me.TxtLossAccount.Text <> "" Then
                    If ObjKrywn.ApakahSudahAda(Me.TxtLossAccount.Text) = False Then
                        Me.BtnFindLoss.PerformClick()
                    End If
                End If
            End If

            '3. harus ada data item
            If Me.dgvItem.Rows.Count <= 0 Then
                MsgBox("Please input data item.", MsgBoxStyle.Information, "Information")
                If Me.ShowFromSalary = True Then
                    Me.PengaturanFocusDataBenar = EnumListFocus.DGVItem
                    Me.ShowDialog()
                End If
                Me.dgvItem.Focus()
                Return False
            End If

            'Harus input Convention date HR
            If PayrollConventionDate = #1/1/1950# Then
                MsgBox("Please input payroll convention date in global setting data.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'Harus isi
            If Me.TxtPrimaryCashAcc.Text = "-1" Then
                MsgBox("Please choose cash account.", MsgBoxStyle.Information, "Information")
                Me.CboCashAccount.Focus()
                Return False
            End If
            If Me.TxtPrimaryEmployee.Text = "-1" Then
                MsgBox("Please input employee code.", MsgBoxStyle.Information, "Information")
                Me.TxtEmployeeCode1.Focus()
                If Me.ShowFromSalary = True Then
                    Me.PengaturanFocusDataBenar = EnumListFocus.EmployeeCode
                    Me.ShowDialog()
                End If
                Return False
            End If
            If Me.TxtPrimaryGainAcc.Text = "-1" Then
                MsgBox("Please input gain account.", MsgBoxStyle.Information, "Information")
                Me.TxtGainAccount.Focus()
                If Me.ShowFromSalary = True Then
                    Me.PengaturanFocusDataBenar = EnumListFocus.GainAcc
                    Me.ShowDialog()
                End If
                Return False
            End If
            If Me.TxtPrimaryLoss.Text = "-1" Then
                MsgBox("Please input loss account.", MsgBoxStyle.Information, "Information")
                Me.TxtLossAccount.Focus()
                If Me.ShowFromSalary = True Then
                    Me.PengaturanFocusDataBenar = EnumListFocus.LossAcc
                    Me.ShowDialog()
                End If
                Return False
            End If
            If Me.ObjCD.CekDataKosong(Me.TxtRemark.Text, False) = False Then
                MsgBox("Data shouldn't include the character: ' , % , or * .", MsgBoxStyle.Information, "Information")
                If Me.ShowFromSalary = True Then
                    Me.PengaturanFocusDataBenar = EnumListFocus.Remarks
                    Me.ShowDialog()
                End If
                Return False
            End If

            '4. cek perubahan nomor transaksi
            If Me.Baru = False Then
                If Me.Obj.TanggalTransaksi.Month = Me.dtpDate.Value.Month And _
                Me.Obj.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                    Me.TxtDocumentNo.Text = Me.Obj.DocumentNo
                Else
                    Dim NoTransaksiLama As String = Me.Obj.GetNoTransaksi(Me.TxtPrimaryLoanPayment.Text)
                    If NoTransaksiLama <> Me.TxtDocumentNo.Text Then
                        If Obj.ApakahSudahAda(Me.TxtDocumentNo.Text) = False Then
                            If MsgBox("You just changed transaction date, this will change your transaction " _
                                      + "number and there's a possibility your transaction no cannot be change back." _
                                      + "Continue and accept this change?", _
                                      MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                                Me.dtpDate.Value = Me.Obj.TanggalTransaksi
                                Me.TxtDocumentNo.Text = Me.Obj.DocumentNo
                                Return False
                            End If
                        End If
                    End If
                End If
            End If

            '5. cek data yang salah, double dalam kode
            If Me.Baru = True Then
                If Me.Obj.ApakahSudahAda(Me.TxtDocumentNo.Text) = True Then
                    DocNoBuffer = Me.TxtDocumentNo.Text
                    Me.TxtDocumentNo.Text = Obj.GetNoBaru(Me.dtpDate.Value, Me.ObjCD)
                End If
            Else
                Dim TipeCodeLama As String = Me.Obj.GetNoTransaksi(Me.TxtPrimaryLoanPayment.Text)
                If TipeCodeLama <> Me.TxtDocumentNo.Text Then
                    If Me.Obj.ApakahSudahAda(Me.TxtDocumentNo.Text) = True Then
                        DocNoBuffer = Me.TxtDocumentNo.Text
                        Me.TxtDocumentNo.Text = Obj.GetNoBaru(Me.dtpDate.Value, Me.ObjCD)
                    End If
                End If
                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.Obj.GetTimestamp(Me.TxtPrimaryLoanPayment.Text)
                If Me.Obj.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have modified this data, reload latest data?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.Obj.CariDetail(Me.Obj.Primary, ObjCD, Me.dgvItem, Me.Def_Prec, Me)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            '6. Cek Account 
            Dim mApakahAkunIncome As Boolean = False
            If CDbl(Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryGainAcc.Text)) <> CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Income) And _
            CDbl(Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryGainAcc.Text)) <> CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Other_Income) And _
            CDbl(Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryGainAcc.Text)) <> CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Asset) And _
             CDbl(Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryGainAcc.Text)) <> CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Other_Expense) And _
             CDbl(Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryGainAcc.Text)) <> CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Cost_of_Goods_Sales) And _
             CDbl(Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryGainAcc.Text)) <> CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Expense) Then
                MsgBox("Please input account code with type income account or other income account.", MsgBoxStyle.Information, "Information")
                Me.TxtGainAccount.Focus()
                If Me.ShowFromSalary = True Then
                    Me.PengaturanFocusDataBenar = EnumListFocus.GainAcc
                    Me.ShowDialog()
                End If
                Return False
            Else
                mApakahAkunIncome = True
            End If
            If mApakahAkunIncome = False Then
                If CDbl(Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryGainAcc.Text)) <> CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Other_Income) Then
                    MsgBox("Please input account code with type income account or other income account.", MsgBoxStyle.Information, "Information")
                    Me.TxtGainAccount.Focus()
                    If Me.ShowFromSalary = True Then
                        Me.PengaturanFocusDataBenar = EnumListFocus.GainAcc
                        Me.ShowDialog()
                    End If
                    Return False
                End If
            End If
            Dim mApakahAkunExpence As Boolean = False
            If CDbl(Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryLoss.Text)) <> CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Expense) And _
            CDbl(Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryLoss.Text)) <> CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Cost_of_Goods_Sales) And _
            CDbl(Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryLoss.Text)) <> CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Asset) And _
            CDbl(Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryLoss.Text)) <> CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Other_Income) And _
             CDbl(Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryLoss.Text)) <> CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Other_Expense) Then
                MsgBox("Please input account code with type expense account, COGS account, or other expense account.", MsgBoxStyle.Information, "Information")
                Me.TxtLossAccount.Focus()
                If Me.ShowFromSalary = True Then
                    Me.PengaturanFocusDataBenar = EnumListFocus.LossAcc
                    Me.ShowDialog()
                End If
                Return False
            Else
                mApakahAkunExpence = True
            End If
            If mApakahAkunExpence = False Then
                If CDbl(Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryLoss.Text)) <> CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Other_Expense) Then
                    MsgBox("Please input account code with type expense account, COGS account, or other expense account.", MsgBoxStyle.Information, "Information")
                    Me.TxtLossAccount.Focus()
                    If Me.ShowFromSalary = True Then
                        Me.PengaturanFocusDataBenar = EnumListFocus.LossAcc
                        Me.ShowDialog()
                    End If
                    Return False
                End If
            End If
            If Me.ObjAkun.GetPrimaryCurrency(Me.TxtPrimaryGainAcc.Text) <> 1 Then
                MsgBox("Please input gain account code with IDR currency.", MsgBoxStyle.Information, "Information")
                Me.TxtGainAccount.Focus()
                If Me.ShowFromSalary = True Then
                    Me.PengaturanFocusDataBenar = EnumListFocus.GainAcc
                    Me.ShowDialog()
                End If
                Return False
            End If
            If Me.ObjAkun.GetPrimaryCurrency(Me.TxtPrimaryLoss.Text) <> 1 Then
                MsgBox("Please input loss account code with IDR currency.", MsgBoxStyle.Information, "Information")
                Me.TxtLossAccount.Focus()
                If Me.ShowFromSalary = True Then
                    Me.PengaturanFocusDataBenar = EnumListFocus.LossAcc
                    Me.ShowDialog()
                End If
                Return False
            End If

            If Me.EnumDipanggilDari = ClsTransaksiLoanPayment.EnumFrmTransLoanPaymentDari.FrmTransLoanPayment Then
                'Harus dibuat setelah konvention date HR
                If Me.dtpDate.Value < PayrollConventionDate Then
                    MsgBox("Please make sure date in transaction loan payment, loan payment's date must be bigger than payroll convention date.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            ElseIf Me.EnumDipanggilDari = ClsTransaksiLoanPayment.EnumFrmTransLoanPaymentDari.FrmTransLoanPayment_Beginning_Balance Then
                'Harus dibuat setelah konvention date HR
                If Me.dtpDate.Value >= PayrollConventionDate Then
                    MsgBox("Please make sure date in transaction beginning balance loan payment, loan payment's date must be bigger than payroll convention date.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            '7. kalo sudah melewati periode closing maka tidak bisa input data
            'Dim ObjClosing As New ClsTransaksiClosing
            'If ObjClosing.GetClosingDate.Date >= Me.dtpDate.Value.Date Then
            'MsgBox("Periode <= " + Format(ObjClosing.GetClosingDate.Date, "dd MMM yyyy") _
            '+ " sudah di kunci. Hubungi admin Anda.", MsgBoxStyle.Information, "Informasi")
            'Return False
            'End If
            'ObjClosing = Nothing

            If DocNoBuffer <> "" Then
                If Me.txtJournalNo.Text = DocNoBuffer Then
                    Me.txtJournalNo.Text = Me.TxtDocumentNo.Text
                End If
                If Me.TxtJournalDesc.Text.Contains(DocNoBuffer) = True Then
                    Me.TxtJournalDesc.Text = Me.TxtJournalDesc.Text.Replace(DocNoBuffer, Me.TxtDocumentNo.Text)
                End If
            Else
                If Me.Baru And Me.ObjLoan.GetJmlByJournalNo(Me.txtJournalNo.Text) > 0 Then
                    MsgBox("Journal No: " & Me.txtJournalNo.Text & " is already exist.", MsgBoxStyle.Information, "Information")
                    Me.txtJournalNo.Focus()
                    Return False
                ElseIf Me.Baru = False And Me.txtJournalNo.Text <> Me.objJournal.GetRefCode(Me.Obj.GetPrimaryJournal(Me.TxtPrimaryLoanPayment.Text)) And _
                        Me.ObjLoan.GetJmlByJournalNo(Me.txtJournalNo.Text) > 0 Then
                    MsgBox("Journal No: " & Me.txtJournalNo.Text & " is already exist.", MsgBoxStyle.Information, "Information")
                    Me.txtJournalNo.Focus()
                    Return False
                End If
            End If

            'sampe sini
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            With Me.Obj
                .DocumentNo = Me.TxtDocumentNo.Text
                .TanggalTransaksi = Me.dtpDate.Value
                .PrimarySalary = Me.TxtPrimarySalary.Text
                .PrimaryJurnal = Me.TxtPrimaryJurnal.Text
                .PrimaryKaryawan = Me.TxtPrimaryEmployee.Text
                .PrimaryCash = Me.TxtPrimaryCashAcc.Text
                .PrimaryGain = Me.TxtPrimaryGainAcc.Text
                .PrimaryLoss = Me.TxtPrimaryLoss.Text
                .Remarks = Me.TxtRemark.Text
                .JournalNo = Me.txtJournalNo.Text
                .JournalDesc = Me.TxtJournalDesc.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeField()
        Dim RowCash As DseTransaksiLoanPayment.tblmaster_akunRow
        Dim RV As Data.DataRowView
        Dim IndexCash As Double = -1
        Try
            With Me.Obj
                Me.TxtPrimaryLoanPayment.Text = .Primary
                Me.TxtDocumentNo.Text = .DocumentNo
                Me.TxtPrimaryJurnal.Text = .PrimaryJurnal
                Me.txtJournalNo.Text = .JournalNo  'Me.objJournal.GetRefCode(Me.TxtPrimaryJurnal.Text)
                Me.TxtJournalDesc.Text = .JournalDesc 'Me.objJournal.GetDescription(Me.TxtPrimaryJurnal.Text)
                Me.dtpDate.Value = .TanggalTransaksi
                Me.TxtPrimarySalary.Text = .PrimarySalary
                Me.TxtPrimaryEmployee.Text = .PrimaryKaryawan
                Me.TxtEmployeeCode1.Text = Me.ObjKrywn.GetCodeByPrimary(Me.TxtPrimaryEmployee.Text)
                Me.TxtEmployeeName.Text = Me.ObjKrywn.GetFullNameByPrimary(Me.TxtPrimaryEmployee.Text)
                Me.Obj.IsiTBLCashCode(Me.CboCashAccount, Me.PrimaryCurr_Default, Me.TBLCash, Me)
                Me.TxtPrimaryCashAcc.Text = .PrimaryCash
                For Each RV In Me.TBLCash.DefaultView
                    RowCash = RV.Row
                    IndexCash += 1
                    If RowCash.Primary = Me.TxtPrimaryCashAcc.Text Then
                        Me.CboCashAccount.SelectedIndex = IndexCash
                        Me.TxtDescCash.Text = RowCash.Description
                        Exit For
                    End If
                Next
                Me.TxtPrimaryGainAcc.Text = .PrimaryGain
                Me.TxtGainAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryGainAcc.Text)
                Me.TxtDescGain.Text = Me.ObjAkun.GetDescription(Me.TxtGainAccount.Text)
                Me.TxtPrimaryLoss.Text = .PrimaryLoss
                Me.TxtLossAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryLoss.Text)
                Me.TxtDescLoss.Text = Me.ObjAkun.GetDescription(Me.TxtLossAccount.Text)
                Me.TxtRemark.Text = .Remarks
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPrint.Click
        Dim FrmJournal As New FrmTransaksiLoan_CetakJournal
        Try
            If Me.TxtPrimaryLoanPayment.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtPrimaryLoanPayment.Text <> "-1" Then
                FrmJournal.PrimaryJournal = Me.TxtPrimaryJurnal.Text
                FrmJournal.ObjLoanPayment = Me.Obj
                FrmJournal.cetak = FrmTransaksiLoan_CetakJournal.EnumCetak.LoanPayment
                FrmJournal.vDgv = Me.dgvItem
                FrmJournal.vFrm = Me
                FrmJournal.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Employee Code
    Private Sub BtnFindEmployee_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindEmployee.Click
        Dim i As New Core_Module.FrmMasterKaryawanCari
        i.BackColorAll = Me.BackColorAll
        Dim Primary_Lama As Double = -1
        Try
            Primary_Lama = Me.TxtPrimaryEmployee.Text
            If Me.TxtEmployeeCode1.Text <> "" Then
                i.txtCari.Text = Me.TxtEmployeeCode1.Text
            End If

            i.ShowDialog()

            Me.Buffer = Me.TxtPrimaryEmployee.Text
            If i.PrimaryEmploye > 0 Then
                Me.TxtPrimaryEmployee.Text = i.PrimaryEmploye
                Me.TxtEmployeeCode1.Text = Me.ObjKrywn.GetCodeByPrimary(Me.TxtPrimaryEmployee.Text)
                Me.TxtEmployeeName.Text = Me.ObjKrywn.GetFullNameByPrimary(Me.TxtPrimaryEmployee.Text)
                Me.Obj.IsiTBLLoanPayment_OnGoing(Me.TBLLoan, Me.TxtPrimaryEmployee.Text, Me.Baru, _
                Me.TxtPrimaryLoanPayment.Text, Me.dgvItem, Me)
                If Primary_Lama <> Me.TxtPrimaryEmployee.Text Then
                    Me.Obj.IsiCBOLoanPayment(Me.TBLLoan, Me.ObjCD, Me.Def_Prec, Me, Me.TBLCBOItem)
                End If
                Me.CboCashAccount.Focus()
            Else
                Me.TxtPrimaryEmployee.Text = "-1"
                Me.TxtEmployeeCode1.Focus()
                Me.TxtEmployeeName.Clear()
                Me.KosongkanFieldItem()
            End If
            Me.SetTxtJournal(False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtEmployeeCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtEmployeeCode1.KeyDown
        Dim objCode As New ClsMasterKaryawan
        Dim i As New FrmMasterKaryawanCari
        Dim Primary_Lama As Double = -1
        Try
            If e.KeyValue = 13 Then
                If Me.TxtEmployeeCode1.Text <> "" Then
                    Primary_Lama = Me.TxtPrimaryEmployee.Text
                    If Me.ObjKrywn.ApakahSudahAda(Me.TxtEmployeeCode1.Text) = True Then
                        Me.TxtPrimaryEmployee.Text = Me.ObjKrywn.GetPrimaryByCode(Me.TxtEmployeeCode1.Text)
                        Me.TxtEmployeeCode1.Text = Me.ObjKrywn.GetCodeByPrimary(Me.TxtPrimaryEmployee.Text)
                        Me.TxtEmployeeName.Text = Me.ObjKrywn.GetFullNameByPrimary(Me.TxtPrimaryEmployee.Text)
                        Me.Obj.IsiTBLLoanPayment_OnGoing(Me.TBLLoan, Me.TxtPrimaryEmployee.Text, Me.Baru, _
                        Me.TxtPrimaryLoanPayment.Text, Me.dgvItem, Me)
                        If Primary_Lama <> Me.TxtPrimaryEmployee.Text Then
                            Me.Obj.IsiCBOLoanPayment(Me.TBLLoan, Me.ObjCD, Me.Def_Prec, Me, Me.TBLCBOItem)
                        End If
                        Me.CboCashAccount.Focus()
                    Else
                        Me.BtnFindEmployee.PerformClick()
                    End If
                Else
                    Me.BtnFindEmployee.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtEmployeeCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmployeeCode1.Leave
        Dim i As New FrmMasterKaryawanCari
        Dim objCode As New ClsMasterKaryawan
        i.BackColorAll = Me.BackColorAll
        Dim Primary_Lama As Double = -1
        Try
            If Me.TxtEmployeeCode1.Text <> "" Then
                If Me.ObjKrywn.ApakahSudahAda(Me.TxtEmployeeCode1.Text) = True Then
                    Primary_Lama = Me.TxtPrimaryEmployee.Text
                    Me.TxtPrimaryEmployee.Text = Me.ObjKrywn.GetPrimaryByCode(Me.TxtEmployeeCode1.Text)
                    Me.TxtEmployeeCode1.Text = Me.ObjKrywn.GetCodeByPrimary(Me.TxtPrimaryEmployee.Text)
                    Me.TxtEmployeeName.Text = Me.ObjKrywn.GetFullNameByPrimary(Me.TxtPrimaryEmployee.Text)
                    Me.Obj.IsiTBLLoanPayment_OnGoing(Me.TBLLoan, Me.TxtPrimaryEmployee.Text, Me.Baru, _
                    Me.TxtPrimaryLoanPayment.Text, Me.dgvItem, Me)
                    If Primary_Lama <> Me.TxtPrimaryEmployee.Text Then
                        Me.Obj.IsiCBOLoanPayment(Me.TBLLoan, Me.ObjCD, Me.Def_Prec, Me, Me.TBLCBOItem)
                    End If
                Else
                    Primary_Lama = Me.TxtPrimaryEmployee.Text
                    If Me.TxtEmployeeCode1.Text <> "" Then
                        i.txtCari.Text = Me.TxtEmployeeCode1.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryEmploye > 0 Then
                        Me.TxtPrimaryEmployee.Text = i.PrimaryEmploye
                        Me.TxtEmployeeCode1.Text = Me.ObjKrywn.GetCodeByPrimary(Me.TxtPrimaryEmployee.Text)
                        Me.TxtEmployeeName.Text = Me.ObjKrywn.GetFullNameByPrimary(Me.TxtPrimaryEmployee.Text)
                        Me.Obj.IsiTBLLoanPayment_OnGoing(Me.TBLLoan, Me.TxtPrimaryEmployee.Text, Me.Baru, _
                        Me.TxtPrimaryLoanPayment.Text, Me.dgvItem, Me)
                        If Primary_Lama <> Me.TxtPrimaryEmployee.Text Then
                            Me.Obj.IsiCBOLoanPayment(Me.TBLLoan, Me.ObjCD, Me.Def_Prec, Me, Me.TBLCBOItem)
                        End If
                    Else
                        Me.TxtPrimaryEmployee.Text = "-1"
                        Me.TxtEmployeeCode1.Clear()
                        Me.TxtEmployeeName.Clear()
                        Me.KosongkanFieldItem()
                    End If
                End If
            Else
                Me.TxtEmployeeCode1.Clear()
                Me.TxtPrimaryEmployee.Text = "-1"
                Me.TxtEmployeeName.Clear()
                Me.KosongkanFieldItem()
            End If

            Me.SetTxtJournal(False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    'Cash Account
    Private Sub CboCashOutAccount_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboCashAccount.SelectedIndexChanged
        Dim row As DseTransaksiLoanPayment.tblmaster_akunRow
        Try
            If Me.CboCashAccount.SelectedIndex = -1 Then
                Me.TxtPrimaryCashAcc.Text = "-1"
            Else
                row = Me.TBLCash.DefaultView(Me.CboCashAccount.SelectedIndex).Row
                Me.TxtPrimaryCashAcc.Text = row.Primary
                Me.TxtDescCash.Text = row.Description
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Gain Account
    Private Sub BtnFindGain_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindGain.Click
        Dim i As New AccModule.FrmMasterAkun_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtGainAccount.Text <> "" Then
                i.txtCari.Text = Me.TxtGainAccount.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.TxtPrimaryGainAcc.Text = i.mPrimaryAkun
                Me.TxtGainAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryGainAcc.Text)
                Me.TxtDescGain.Text = Me.ObjAkun.GetDescription(Me.TxtGainAccount.Text)
                Me.TxtLossAccount.Focus()
            Else
                Me.TxtGainAccount.Clear()
                Me.TxtPrimaryGainAcc.Text = "-1"
                Me.TxtGainAccount.Focus()
                Me.TxtDescGain.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtGainAccount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtGainAccount.KeyDown
        Dim objCode As New AccModule.ClsMasterAkun
        Dim i As New AccModule.FrmMasterAkun_Cari
        Try
            If e.KeyValue = 13 Then
                If Me.TxtGainAccount.Text <> "" Then
                    If objCode.ApakahSudahAda_Aktif_NonHeader(Me.TxtGainAccount.Text) = True Then
                        Me.TxtPrimaryGainAcc.Text = objCode.GetPrimary(Me.TxtGainAccount.Text)
                        Me.TxtGainAccount.Text = objCode.GetAccountCode(Me.TxtPrimaryGainAcc.Text)
                        Me.TxtDescGain.Text = objCode.GetDescription(Me.TxtGainAccount.Text)
                        Me.TxtLossAccount.Focus()
                    Else
                        Me.BtnFindGain.PerformClick()
                    End If
                Else
                    Me.BtnFindGain.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtGainAccount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtGainAccount.Leave
        Dim i As New AccModule.FrmMasterAkun_Cari
        Dim objCode As New AccModule.ClsMasterAkun
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtGainAccount.Text <> "" Then
                If objCode.ApakahSudahAda_Aktif_NonHeader(Me.TxtGainAccount.Text) = True Then
                    Me.TxtPrimaryGainAcc.Text = objCode.GetPrimary(Me.TxtGainAccount.Text)
                    Me.TxtGainAccount.Text = objCode.GetAccountCode(Me.TxtPrimaryGainAcc.Text)
                    Me.TxtDescGain.Text = objCode.GetDescription(Me.TxtGainAccount.Text)
                Else
                    If Me.TxtGainAccount.Text <> "" Then
                        i.txtCari.Text = Me.TxtGainAccount.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryGainAcc.Text = i.mPrimaryAkun
                        Me.TxtGainAccount.Text = objCode.GetAccountCode(Me.TxtPrimaryGainAcc.Text)
                        Me.TxtDescGain.Text = objCode.GetDescription(Me.TxtGainAccount.Text)
                    Else
                        Me.TxtGainAccount.Clear()
                        Me.TxtPrimaryGainAcc.Text = "-1"
                        Me.TxtDescGain.Clear()
                    End If
                End If
            Else
                Me.TxtGainAccount.Clear()
                Me.TxtPrimaryGainAcc.Text = "-1"
                Me.TxtDescGain.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    'Loss Account
    Private Sub BtnFindLoss_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindLoss.Click
        Dim i As New AccModule.FrmMasterAkun_Cari
        i.BackColorAll = Me.BackColorAll
        Dim ObjCode As New AccModule.ClsMasterAkun
        Try
            If Me.TxtLossAccount.Text <> "" Then
                i.txtCari.Text = Me.TxtLossAccount.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.TxtPrimaryLoss.Text = i.mPrimaryAkun
                Me.TxtLossAccount.Text = ObjCode.GetAccountCode(Me.TxtPrimaryLoss.Text)
                Me.TxtDescLoss.Text = ObjCode.GetDescription_ByPrimaryAcc(Me.TxtPrimaryLoss.Text)
                Me.TxtRemark.Focus()
            Else
                Me.TxtLossAccount.Clear()
                Me.TxtPrimaryLoss.Text = "-1"
                Me.TxtDescLoss.Text = " "
                Me.TxtLossAccount.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            ObjCode = Nothing
        End Try
    End Sub

    Private Sub TxtLossAccount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtLossAccount.KeyDown
        Dim objCode As New AccModule.ClsMasterAkun
        Dim i As New AccModule.FrmMasterAkun_Cari
        Try
            If e.KeyValue = 13 Then
                If Me.TxtLossAccount.Text <> "" Then
                    If objCode.ApakahSudahAda_Aktif_NonHeader(Me.TxtLossAccount.Text) = True Then
                        Me.TxtPrimaryLoss.Text = objCode.GetPrimary(Me.TxtLossAccount.Text)
                        Me.TxtLossAccount.Text = objCode.GetAccountCode(Me.TxtPrimaryLoss.Text)
                        Me.TxtDescLoss.Text = objCode.GetDescription_ByPrimaryAcc(Me.TxtPrimaryLoss.Text)
                        Me.TxtRemark.Focus()
                    Else
                        Me.BtnFindLoss.PerformClick()
                    End If
                Else
                    Me.BtnFindLoss.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtLossAccount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtLossAccount.Leave
        Dim i As New AccModule.FrmMasterAkun_Cari
        Dim objCode As New AccModule.ClsMasterAkun
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtLossAccount.Text <> "" Then
                If objCode.ApakahSudahAda_Aktif_NonHeader(Me.TxtLossAccount.Text) = True Then
                    Me.TxtPrimaryLoss.Text = objCode.GetPrimary(Me.TxtLossAccount.Text)
                    Me.TxtLossAccount.Text = objCode.GetAccountCode(Me.TxtPrimaryLoss.Text)
                    Me.TxtDescLoss.Text = objCode.GetDescription_ByPrimaryAcc(Me.TxtPrimaryLoss.Text)
                Else
                    If Me.TxtLossAccount.Text <> "" Then
                        i.txtCari.Text = Me.TxtLossAccount.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryLoss.Text = i.mPrimaryAkun
                        Me.TxtLossAccount.Text = objCode.GetAccountCode(Me.TxtPrimaryLoss.Text)
                        Me.TxtDescLoss.Text = objCode.GetDescription_ByPrimaryAcc(Me.TxtPrimaryLoss.Text)
                    Else
                        Me.TxtLossAccount.Clear()
                        Me.TxtPrimaryLoss.Text = "-1"
                        Me.TxtDescLoss.Text = " "
                    End If
                End If
            Else
                Me.TxtLossAccount.Clear()
                Me.TxtPrimaryLoss.Text = "-1"
                Me.TxtDescLoss.Text = " "
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    'Up Down
    Private Sub DGVItemSorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsAdded(ByVal sender As Object, ByVal e As _
    System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Try
            Me.dgvItem.TabStop = True
            If Me.TxtRemarksItem.Enabled = True Then
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsRemoved(ByVal sender As Object, ByVal e As _
    System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItem.RowsRemoved
        Try
            If Me.dgvItem.Rows.Count = 0 Then
                Me.dgvItem.TabStop = False
                If Me.TxtRemarksItem.Enabled = True Then
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
                Me.TxtEmployeeCode1.Enabled = True
                Me.BtnFindEmployee.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e _
    As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvItem)
            'Posisi Employee Code 
            If Me.dgvItem.Rows.Count <= 0 Then
                If Me.Baru = True Then
                    Me.TxtPrimaryEmployee.Text = "-1"
                    Me.TxtEmployeeCode1.Clear()
                    Me.TxtEmployeeName.Clear()
                    Me.TBLLoan.Rows.Clear()
                    Me.TBLCBOItem.Rows.Clear()
                    Me.TxtPrimaryLoan.Text = "-1"
                    Me.TxtIdentifier.Text = ""
                    Me.CboLoanNumber.DataSource = Nothing
                    Me.CboLoanNumber.Items.Clear()
                    Me.TxtEmployeeCode1.Enabled = True
                    Me.BtnFindEmployee.Enabled = True
                    Me.KosongkanFieldItem()
                    Me.TxtEmployeeCode1.Focus()
                Else
                    Me.CboLoanNumber.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles dgvItem.SelectionChanged
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.dgvItem.SelectedRows.Count > 0 Then
                    Me.btnUP.Enabled = True
                    Me.btnDown.Enabled = True
                Else
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUP.Click
        Try
            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItem.SelectedRows(0)

                If row.Index <> 0 Then
                    Dim indexbaru = row.Index - 1
                    Me.dgvItem.Rows.Remove(row)
                    Me.dgvItem.Rows.InsertRange(indexbaru, row)
                    Me.dgvItem.Focus()
                    Me.dgvItem.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItem)
                    If row.Displayed = False Then
                        Me.ShowRowProcess(row, True)
                    End If
                End If

                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False

                Dim row As DataGridViewRow = Me.dgvItem.SelectedRows(0)
                If row.Index <> Me.dgvItem.Rows.Count - 1 Then
                    Dim indexbaru = row.Index + 1
                    Me.dgvItem.Rows.Remove(row)
                    Me.dgvItem.Rows.InsertRange(indexbaru, row)
                    Me.dgvItem.Focus()
                    Me.dgvItem.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItem)
                    Me.ShowRowProcess(row, False)
                End If

                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' tampilkan row yang diberikan pada parameter
    ''' </summary>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Private Sub ShowRowProcess(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < Me.dgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.dgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While Me.dgvItem.Rows(row.Index).Displayed = False
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False) Then
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Payment Amount
    Private Sub TxtPayment_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPayment.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtPayment, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPayment_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPayment.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtPayment, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Back Color
    Private Sub TxtEmployeeCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmployeeCode1.GotFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                Me.Buffer = Me.TxtPrimaryEmployee.Text
                If Me.TxtEmployeeCode1.Enabled = True Then
                    If Me.TxtEmployeeCode1.ReadOnly = False Then
                        Me.TxtEmployeeCode1.BackColor = Me.BackColorAll
                    Else
                        Me.TxtEmployeeCode1.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtEmployeeCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmployeeCode1.LostFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtEmployeeCode1.Enabled = True Then
                    If Me.TxtEmployeeCode1.ReadOnly = False Then
                        Me.TxtEmployeeCode1.BackColor = Me.TxtPrimaryLoanPayment.BackColor
                    Else
                        Me.TxtEmployeeCode1.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub GotFocusTxtGainAccount()
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtGainAccount.Enabled = True Then
                    If Me.TxtGainAccount.ReadOnly = False Then
                        TxtGainAccount.BackColor = Me.BackColorAll
                    Else
                        TxtGainAccount.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtGainAccount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtGainAccount.GotFocus
        Try
            Me.GotFocusTxtGainAccount()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub GotLostFocusTxtGainAccount()
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtGainAccount.Enabled = True Then
                    If Me.TxtGainAccount.ReadOnly = False Then
                        TxtGainAccount.BackColor = Me.TxtPrimaryLoanPayment.BackColor
                    Else
                        TxtGainAccount.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtGainAccount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtGainAccount.LostFocus
        Try
            Me.GotLostFocusTxtGainAccount()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLossAccount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLossAccount.GotFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtLossAccount.Enabled = True Then
                    If Me.TxtLossAccount.ReadOnly = False Then
                        TxtLossAccount.BackColor = Me.BackColorAll
                    Else
                        TxtLossAccount.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLossAccount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLossAccount.LostFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtLossAccount.Enabled = True Then
                    If Me.TxtLossAccount.ReadOnly = False Then
                        TxtLossAccount.BackColor = Me.TxtPrimaryLoanPayment.BackColor
                    Else
                        TxtLossAccount.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPayment_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPayment.GotFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtPayment.Enabled = True Then
                    If Me.TxtPayment.ReadOnly = False Then
                        TxtPayment.BackColor = Me.BackColorAll
                    Else
                        TxtPayment.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPayment_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPayment.LostFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtPayment.Enabled = True Then
                    If Me.TxtPayment.ReadOnly = False Then
                        TxtPayment.BackColor = Me.TxtPrimaryLoanPayment.BackColor
                    Else
                        TxtPayment.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarksItem_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarksItem.GotFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtRemarksItem.Enabled = True Then
                    If Me.TxtRemarksItem.ReadOnly = False Then
                        TxtRemarksItem.BackColor = Me.BackColorAll
                    Else
                        TxtRemarksItem.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarksItem_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarksItem.LostFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtRemarksItem.Enabled = True Then
                    If Me.TxtRemarksItem.ReadOnly = False Then
                        TxtRemarksItem.BackColor = Me.TxtPrimaryLoanPayment.BackColor
                    Else
                        TxtRemarksItem.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.GotFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtRemark.Enabled = True Then
                    If Me.TxtRemark.ReadOnly = False Then
                        TxtRemark.BackColor = Me.BackColorAll
                    Else
                        TxtRemark.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.LostFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtRemark.Enabled = True Then
                    If Me.TxtRemark.ReadOnly = False Then
                        TxtRemark.BackColor = Me.TxtPrimaryLoanPayment.BackColor
                    Else
                        TxtRemark.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub TxtPayment_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPayment.Enter
        Try
            If CDec(Me.TxtPayment.Text) = 0 And CDec(Me.TxtTotalInstallment.Text) <> 0 Then
                Me.TxtPayment.Text = Me.TxtTotalInstallment.Text
            End If
            Me.ObjCD.Number_Enter(Me.TxtPayment, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPayment_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPayment.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtPayment, sender, e, Me.Def_Prec, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Data Item Payment 
    Private Function ApakahDataItemBenar() As Boolean
        Try
            If Me.TxtPrimaryLoan.Text = "-1" Then
                MsgBox("Please choose loan number.", MsgBoxStyle.Information, "Information")
                Me.CboLoanNumber.Focus()
                Return False
            End If
            If Me.TxtPrimaryEmployee.Text = "-1" Then
                MsgBox("Please input employee code.", MsgBoxStyle.Information, "Information")
                If Me.TxtEmployeeCode1.Enabled = True Then
                    Me.TxtEmployeeCode1.Focus()
                End If
                Return False
            End If
            If Me.ObjCD.CekBilanganPositif(Me.TxtPayment.Text) = False Then
                MsgBox("Please input positive number.", MsgBoxStyle.Information, "Infomation")
                Me.TxtPayment.Focus()
                Return False
            End If
            If Me.ObjCD.CekDataKosong(Me.TxtRemarksItem.Text, False) = False Then
                MsgBox("Data shouldn't include the character: ' , % , or * .", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetIndexItem(ByVal vIdentifierLoan As String) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColIdentifier.Index).Value = vIdentifierLoan Then
                    Return rowDGV.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub BtnInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnInput.Click
        Dim index As Double = -1
        Dim rowDGV As DataGridViewRow
        Dim RowLoan As DseTransaksiLoanPayment.DataTable2Row
        Dim RowStatusLoan As DseTransaksiLoanPayment.TBLUpdateStatus_LoanRow
        Dim Filter As String = ""
        Try
            If Me.ApakahDataItemBenar() = True Then
                Me.TxtEmployeeCode1.Enabled = False
                Me.BtnFindEmployee.Enabled = False
                'cek apakah sudah ada item jenis pekerjaan sama yang masuk
                index = Me.GetIndexItem(Me.TxtIdentifier.Text)
                If index < 0 Then
                    Me.dgvItem.Rows.Add()
                    rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    With rowDGV
                        .Cells(Me.ColNo.Index).Value = ""
                        .Cells(Me.ColPrimaryLoan.Index).Value = Me.TxtPrimaryLoan.Text
                        .Cells(Me.ColIdentifier.Index).Value = Me.TxtIdentifier.Text
                        .Cells(Me.ColLoanNo.Index).Value = Me.Obj.GetLoanNumber(.Cells(Me.ColPrimaryLoan.Index).Value)
                        .Cells(Me.ColInstallementAt.Index).Value = Me.TxtInstallmentAt.Text
                        'Berkaitan dengan status loan
                        If Me.ObjLoan.GetInstallmentTo(.Cells(Me.ColLoanNo.Index).Value.ToString) = CDbl(Me.TxtInstallmentAt.Text) Then
                            Me.TBLLoanStatus.DefaultView.RowFilter = "[Loan Number] = '" + .Cells(Me.ColLoanNo.Index).Value.ToString + "'"
                            If TBLLoanStatus.DefaultView.Count > 0 Then
                                RowStatusLoan = TBLLoanStatus.DefaultView(0).Row
                                RowStatusLoan.Posisi = ClsTransaksiLoan.EnumStatusLoan.Finished
                            Else
                                RowStatusLoan = Me.TBLLoanStatus.NewRow
                                RowStatusLoan.Loan_Number = .Cells(Me.ColLoanNo.Index).Value.ToString
                                RowStatusLoan.Installment_To = .Cells(Me.ColInstallementAt.Index).Value.ToString
                                RowStatusLoan.Identifier_Loan = .Cells(Me.ColIdentifier.Index).Value.ToString
                                RowStatusLoan.Posisi = ClsTransaksiLoan.EnumStatusLoan.Finished
                                Me.TBLLoanStatus.Rows.Add(RowStatusLoan)
                            End If
                            Me.TBLLoanStatus.DefaultView.RowFilter = ""
                        End If
                        .Cells(Me.ColBase.Index).Value = Me.TxtBase.Text
                        .Cells(Me.ColInterest.Index).Value = Me.TxtInterest.Text
                        .Cells(Me.ColTotalInstallment.Index).Value = Me.TxtTotalInstallment.Text
                        .Cells(Me.ColPayment.Index).Value = Me.TxtPayment.Text
                        .Cells(Me.ColRemarks.Index).Value = Me.TxtRemarksItem.Text
                        .Cells(Me.ColPrimaryRev.Index).Value = Me.TxtPrimaryRev.Text
                        If Me.ShowFromLinkLabel = True Then
                            Me.TotalLoanInstallment += CDec(Me.TxtPayment.Text)
                        End If
                    End With
                    'set number
                    Me.ObjCD.SetNumber(Me.dgvItem)
                    'ubah data index untuk fokus
                    index = Me.dgvItem.Rows.Count - 1
                    'Berkaitan dengan Loan Payment
                    If Me.TBLLoan.Rows.Count > 0 Then
                        Filter = "[Identifier] = '" + Me.TxtIdentifier.Text.ToString + "' AND [Installment Period] = " + Me.TxtInstallmentAt.Text.ToString + ""
                        Me.TBLLoan.DefaultView.RowFilter = Filter
                        If Me.TBLLoan.DefaultView.Count > 0 Then
                            RowLoan = Me.TBLLoan.DefaultView(0).Row
                            RowLoan.Payment_Amount = CDec(Me.TxtPayment.Text)
                            RowLoan.Update = 1
                            Me.Obj.IsiCBOLoanPayment(Me.TBLLoan, Me.ObjCD, Me.Def_Prec, Me, Me.TBLCBOItem)
                        End If
                        Me.TBLLoan.DefaultView.RowFilter = ""
                    End If
                    'Posisi Edit
                    If Me.Baru = False Then
                        If Me.dgvItem.Rows.Count > 0 Then
                            Me.TxtEmployeeCode1.Enabled = False
                            Me.BtnFindEmployee.Enabled = False
                        End If
                    End If
                Else
                    MsgBox("Duplication data item.", MsgBoxStyle.Information, "Information")
                    Me.CboLoanNumber.Focus()
                    Exit Sub
                End If
                'End If
                Me.TxtPayment.Text = "0"
                Me.TxtRemarksItem.Text = ""
                Me.CboLoanNumber.Focus()

                'fokuskan ke baris ybs
                If Me.dgvItem.FirstDisplayedScrollingRowIndex < index Then
                    While Me.dgvItem.Rows(index).Displayed = False _
                    Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                        Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                    End While
                Else
                    While Me.dgvItem.Rows(index).Displayed = False
                        Me.dgvItem.FirstDisplayedScrollingRowIndex = _
                        Me.dgvItem.FirstDisplayedScrollingRowIndex - 1
                    End While
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Dim RowDGV As DataGridViewRow = Me.dgvItem.Rows(e.Row.Index)
        Dim Row As DataGridViewRow
        Dim RowLoan As DseTransaksiLoanPayment.DataTable2Row
        Dim RowStatusLoan As DseTransaksiLoanPayment.TBLUpdateStatus_LoanRow
        Dim Filter As String = ""
        Dim BolehHapus As Boolean = True
        Dim MAXCicilan As String = -1
        Try
            '1. Cek sebelum dihapus, harus cicilan yang terakhir dihapus lebih dahulu, berdasarkan DGVItem
            For Each Row In Me.dgvItem.Rows
                If Row.Index <> RowDGV.Index Then
                    If Row.Cells(Me.ColPrimaryLoan.Index).Value = RowDGV.Cells(Me.ColPrimaryLoan.Index).Value Then
                        If CDbl(RowDGV.Cells(Me.ColInstallementAt.Index).Value) < CDbl(Row.Cells(Me.ColInstallementAt.Index).Value) Then
                            BolehHapus = False
                            MAXCicilan = Row.Cells(Me.ColInstallementAt.Index).Value.ToString
                            Exit For
                        End If
                    End If
                End If
            Next
            If BolehHapus = False Then
                MsgBox("Payment cannot deleted, last installment in payment code :" + Me.TxtDocumentNo.Text.ToString + ", loan code : " + RowDGV.Cells(Me.ColLoanNo.Index).Value.ToString + _
                ", installment at: " + MAXCicilan + ".", MsgBoxStyle.Information, "Information")
                e.Cancel = True
                Exit Sub
            End If
            '2. Cek sebelum dihapus, harus cicilan yang terakhir dihapus lebih dahulu, berdasarkan database
            Dim NoDoc_Payment As String = ""
            Dim NoLoan As String = ""
            MAXCicilan = -1
            If Me.Baru = False Then
                If Me.Obj.BatasanHapusLoanPayment_Row_Database(RowDGV.Cells(Me.ColPrimaryLoan.Index).Value, _
                Me.TxtPrimaryLoanPayment.Text, RowDGV.Cells(Me.ColInstallementAt.Index).Value, _
                NoDoc_Payment, NoLoan, MAXCicilan) = False Then
                    MsgBox("Payment cannot deleted, last installment in payment code :" + NoDoc_Payment + ", loan code : " + NoLoan + _
                    ", installment at: " + MAXCicilan + ".", MsgBoxStyle.Information, "Information")
                    e.Cancel = True
                    Exit Sub
                End If
            End If
            '3. Update nilai amount paid, berdasarkan identifier dan cicilan ke berapa, kemudian masukan ulang CBOLoan.
            If Me.TBLLoan.Rows.Count > 0 Then
                Filter = "[Identifier] = '" + RowDGV.Cells(Me.ColIdentifier.Index).Value.ToString + "' AND [Installment Period] = " + RowDGV.Cells(Me.ColInstallementAt.Index).Value.ToString + ""
                Me.TBLLoan.DefaultView.RowFilter = Filter
                If Me.TBLLoan.DefaultView.Count > 0 Then
                    RowLoan = Me.TBLLoan.DefaultView(0).Row
                    RowLoan.Payment_Amount = 0
                    RowLoan.Update = 1
                    Me.Obj.IsiCBOLoanPayment(Me.TBLLoan, Me.ObjCD, Me.Def_Prec, Me, Me.TBLCBOItem)
                End If
                Me.TBLLoan.DefaultView.RowFilter = ""
            End If
            '4 kurangkan Total Installment
            If Me.ShowFromSalary = True Then
                Me.TotalLoanInstallment = Me.TotalLoanInstallment - CDec(e.Row.Cells(Me.ColPayment.Index).Value)
            End If
            '5. Berkaitan dengan status loan, jika yang dihapus cicilan terakhir maka harus update status = On Going
            If Me.ObjLoan.GetInstallmentTo(RowDGV.Cells(Me.ColLoanNo.Index).Value.ToString) = CDbl(Me.TxtInstallmentAt.Text) Then
                Me.TBLLoanStatus.DefaultView.RowFilter = "[Loan Number] = '" + RowDGV.Cells(Me.ColLoanNo.Index).Value.ToString + "'"
                If TBLLoanStatus.DefaultView.Count > 0 Then
                    RowStatusLoan = TBLLoanStatus.DefaultView(0).Row
                    RowStatusLoan.Posisi = ClsTransaksiLoan.EnumStatusLoan.OnGoing
                Else
                    RowStatusLoan = Me.TBLLoanStatus.NewRow
                    RowStatusLoan.Loan_Number = RowDGV.Cells(Me.ColLoanNo.Index).Value.ToString
                    RowStatusLoan.Installment_To = RowDGV.Cells(Me.ColInstallementAt.Index).Value.ToString
                    RowStatusLoan.Identifier_Loan = RowDGV.Cells(Me.ColIdentifier.Index).Value.ToString
                    RowStatusLoan.Posisi = ClsTransaksiLoan.EnumStatusLoan.OnGoing
                    Me.TBLLoanStatus.Rows.Add(RowStatusLoan)
                End If
                Me.TBLLoanStatus.DefaultView.RowFilter = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColPayment.Index Or _
            e.ColumnIndex = Me.ColRemarks.Index Then
                Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim Row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
        Dim RowLoan As DseTransaksiLoanPayment.DataTable2Row
        Dim Filter As String = ""
        Try
            'pastikan ada isinya
            With Row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColPayment.Index Then
                        If IsNumeric(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        If Me.ObjCD.CekBilanganPositif(Row.Cells(Me.ColPayment.Index).Value) = False Then
                            MsgBox("Please input positive number.", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        If Me.TBLLoan.Rows.Count > 0 Then
                            Filter = "[Identifier] = '" + Row.Cells(Me.ColIdentifier.Index).Value.ToString + "' AND [Installment Period] = " + Row.Cells(Me.ColInstallementAt.Index).Value.ToString + ""
                            Me.TBLLoan.DefaultView.RowFilter = Filter
                            If Me.TBLLoan.DefaultView.Count > 0 Then
                                RowLoan = Me.TBLLoan.DefaultView(0).Row
                                RowLoan.Payment_Amount = Row.Cells(Me.ColPayment.Index).Value.ToString
                                RowLoan.Update = 1
                                Me.Obj.IsiCBOLoanPayment(Me.TBLLoan, Me.ObjCD, Me.Def_Prec, Me, Me.TBLCBOItem)
                            End If
                            Me.TBLLoan.DefaultView.RowFilter = ""
                        End If
                        If Me.ShowFromSalary = True Then
                            Me.TotalLoanInstallment = Me.TotalLoanInstallment - Me.Buffer
                            Me.TotalLoanInstallment = Me.TotalLoanInstallment + CDec(Row.Cells(Me.ColPayment.Index).Value)
                        End If
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, Me.Def_Prec)
                    End If
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        If Me.ObjCD.CekDataKosong(Row.Cells(Me.ColRemarks.Index).Value) = False Then
                            MsgBox("Data shouldn't include the character: ' , % , or * .", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboLoanNumber_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboLoanNumber.SelectedIndexChanged
        Dim RowCBOItem As DseTransaksiLoanPayment.DataTable2Row
        Try

            If Me.CboLoanNumber.Text <> "" Then
                If Me.TxtPrimaryLoan.Text = "-1" Then
                    Me.TxtIdentifier.Text = ""
                    Me.TxtPrimaryLoan.Text = "-1"
                    Me.TxtInstallmentAt.Text = "0"
                    Me.TxtBase.Text = "0"
                    Me.TxtInterest.Text = "0"
                    Me.TxtTotalInstallment.Text = "0"
                    Me.TxtPayment.Text = "0"
                    Me.CboLoanNumber.DataSource = Nothing
                Else
                    RowCBOItem = Me.TBLCBOItem(Me.CboLoanNumber.SelectedIndex)
                    Me.TxtIdentifier.Text = RowCBOItem.Identifier
                    Me.TxtPrimaryLoan.Text = RowCBOItem.Primary
                    Me.TxtInstallmentAt.Text = RowCBOItem.Installment_Period
                    Me.TxtBase.Text = Me.ObjCD.DigitGrouping(RowCBOItem.Base, Me.Def_Prec)
                    Me.TxtInterest.Text = Me.ObjCD.DigitGrouping(RowCBOItem.Interest, Me.Def_Prec)
                    Me.TxtTotalInstallment.Text = Me.ObjCD.DigitGrouping(CDec(RowCBOItem.Base + RowCBOItem.Interest), Me.Def_Prec)
                    Me.TxtPayment.Text = "0"
                    Me.TxtPrimaryRev.Text = RowCBOItem.Primary_Receivable_Account
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtJournalNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtJournalNo.GotFocus
        Try
            If Me.txtJournalNo.ReadOnly = False And Me.txtJournalNo.Enabled = True Then
                Me.txtJournalNo.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtJournalNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtJournalNo.LostFocus

        Try
            If Me.txtJournalNo.Enabled = True Then
                If Me.txtJournalNo.ReadOnly = True Then
                    Me.txtJournalNo.BackColor = Me.BackColorAll
                Else
                    Me.txtJournalNo.BackColor = Color.White
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtJournalDesc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJournalDesc.GotFocus
        Try
            If Me.TxtJournalDesc.Enabled = True Then
                If Me.TxtJournalDesc.ReadOnly = True Then
                    Me.TxtJournalDesc.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtJournalDesc.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtJournalDesc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJournalDesc.LostFocus

        Try
            If Me.TxtJournalDesc.Enabled = True Then
                If Me.TxtJournalDesc.ReadOnly = True Then
                    Me.TxtJournalDesc.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtJournalDesc.BackColor = Color.White
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDocumentNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDocumentNo.TextChanged
        Try
            Me.SetTxtJournal(True)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetTxtJournal(ByVal vChangeDoc As Boolean)
        Dim JournalDesc As String = ""

        Try
            'set journal no
            If Me.txtJournalNo.Text = Me.Buffer Then
                Me.txtJournalNo.Text = Me.TxtDocumentNo.Text
            ElseIf Me.txtJournalNo.Text = "" Then
                Me.txtJournalNo.Text = Me.TxtDocumentNo.Text
            End If

            'set journal description
            If Me.TxtJournalDesc.Text = "" Then
                Me.TxtJournalDesc.Text = Me.TxtDocumentNo.Text
            End If
            If vChangeDoc = True Then
                If Me.Buffer <> "" Then
                    If Me.TxtJournalDesc.Text.Contains(Me.Buffer) = True Then
                        Me.TxtJournalDesc.Text = Me.TxtJournalDesc.Text.Replace(Me.Buffer, Me.TxtDocumentNo.Text)
                    End If
                End If
            Else
                If Me.Buffer <> "-1" Then
                    JournalDesc = Me.TxtDocumentNo.Text & "-" & Me.ObjKrywn.GetCodeByPrimary(Me.Buffer) + "-" & Me.ObjKrywn.GetFullNameByPrimary(Me.Buffer)
                Else
                    JournalDesc = Me.TxtDocumentNo.Text
                End If

                If Me.TxtJournalDesc.Text = JournalDesc Then
                    If Me.TxtPrimaryEmployee.Text <> "-1" Then
                        Me.TxtJournalDesc.Text = Me.TxtDocumentNo.Text & "-" & Me.TxtEmployeeCode1.Text + "-" & Me.TxtEmployeeName.Text
                    Else
                        Me.TxtJournalDesc.Text = Me.TxtDocumentNo.Text
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted

    End Sub
End Class