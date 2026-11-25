Option Explicit On
Option Strict Off


Partial Public Class FrmSettingAP_AR
    Dim SettingPajakPPH21Berubah As Boolean
    Dim ObjPTKP As New HR.ClsMasterSettingPajak
    Dim DefPrec As Double = Me.ObjCurrency.GetDecimal_DefaultCurrency
    Dim DefPrec_Persen As Double = 2
    Dim MAXLength_Tarif As Decimal = 100
    Dim MAXLength_Angka As Decimal = 999999999999
    Dim MAXLength_Angka_Str As String = Me.ObjCD.DigitGrouping(Me.MAXLength_Angka, 0)
    Dim EraseRowIndex As Double = 0
    Dim FrmHR As New HR.FrmMasterSettingPajak

    Private Sub SimpanDataSettingPPH21()
        Try
            Me.ObjPTKP.Simpan(Me.dgvPTKP, Me.dgvRate, Me.FrmHR)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvPTKP_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvPTKP.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColPTKP_PTKP_Tahunan.Index Or _
            e.ColumnIndex = Me.ColPTKP_Remarks.Index Then
                Me.Buffer = Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvPTKP_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvPTKP.CellEndEdit
        Dim Row As DataGridViewRow = Me.dgvPTKP.Rows(e.RowIndex)
        Try
            'pastikan ada isinya
            With Row
                If IsNothing(Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColPTKP_PTKP_Tahunan.Index Then
                        If IsNumeric(Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            If IsNumeric(Me.Buffer) = True Then
                                Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer)
                            Else
                                Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value = 0
                            End If
                            Exit Sub
                        End If
                        If Me.ObjCD.CekBilanganPositif(.Cells(Me.ColPTKP_PTKP_Tahunan.Index).Value) = False Then
                            MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                            Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec)
                            Exit Sub
                        Else
                            Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping( _
                            Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec)
                        End If
                        If CDec(.Cells(Me.ColPTKP_PTKP_Tahunan.Index).Value) > Me.MAXLength_Angka Then
                            MsgBox("Max Length number 12-digits.", MsgBoxStyle.Information, "Information")
                            Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec)
                            Exit Sub
                        Else
                            Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping( _
                            Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec)
                        End If
                        Me.SettingPajakPPH21Berubah = True
                    End If
                    If e.ColumnIndex = Me.ColPTKP_Remarks.Index Then
                        If Me.ObjCD.CekDataKosong(.Cells(Me.ColPTKP_Remarks.Index).Value, False) = False Then
                            MsgBox("Data shouldn't include the character: ' , % , or * .", MsgBoxStyle.Information, "Information")
                            Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        Me.SettingPajakPPH21Berubah = True
                    End If
                Else
                    If e.ColumnIndex = Me.ColPTKP_Remarks.Index Then
                        Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        Me.SettingPajakPPH21Berubah = True
                    Else
                        Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvRate_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvRate.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColRate_Maximum.Index Or _
            e.ColumnIndex = Me.ColRate_Remarks.Index Or _
            e.ColumnIndex = Me.ColRate_Tarif.Index Then
                Me.Buffer = Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahCellRateKosong(ByVal vRow As DataGridViewRow, Optional ByVal CekRemarks As Boolean = True) As Boolean
        Try
            With vRow
                If .Cells(Me.ColRate_Maximum.Index).Value = "" Or CDec(.Cells(Me.ColRate_Maximum.Index).Value) = 0 Then
                    If .Cells(Me.ColRate_Tarif.Index).Value = "" Or CDec(.Cells(Me.ColRate_Tarif.Index).Value) = 0 Then
                        If CekRemarks = True Then
                            If .Cells(Me.ColRate_Remarks.Index).Value = "" Then
                                Return True
                            End If
                        Else
                            Return True
                        End If
                    End If
                End If
            End With
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvRate_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvRate.CellEndEdit
        Dim Row As DataGridViewRow = Me.dgvRate.Rows(e.RowIndex)
        Dim RowIndex_Atas As Integer = e.RowIndex - 1
        Dim RowIndex_Bawah As Integer = e.RowIndex + 1
        Try
            'pastikan ada isinya
            With Row
                If IsNothing(Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvRate.AllowUserToAddRows = False
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Me.dgvRate.AllowUserToAddRows = True
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColRate_Maximum.Index Then
                        If Me.ObjCD.CekBilanganPositif(.Cells(Me.ColRate_Maximum.Index).Value) = False Then
                            MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                            If Me.Buffer <> "" Then
                                Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec)
                            Else
                                Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            If Me.ApakahCellRateKosong(Row) = True Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            End If
                            Me.dgvRate.AllowUserToAddRows = True
                            Exit Sub
                        Else
                            Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping( _
                            Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec)
                        End If
                        If CDec(.Cells(Me.ColRate_Maximum.Index).Value) > Me.MAXLength_Angka Then
                            MsgBox("Max Length number 12-digits.", MsgBoxStyle.Information, "Information")
                            If Me.Buffer <> "" Then
                                Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec)
                            Else
                                Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            If Me.ApakahCellRateKosong(Row) = True Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            End If
                            Me.dgvRate.AllowUserToAddRows = True
                            Exit Sub
                        Else
                            Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping( _
                            Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec)
                        End If
                        If Me.dgvRate.Rows.Count > 1 Then
                            'batas atas
                            If e.RowIndex > 0 Then
                                Dim MAXRate As Decimal = Me.dgvRate.Item(e.ColumnIndex, RowIndex_Atas).Value
                                If CDec(Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value) < MAXRate Then
                                    MsgBox("You cannot input maximum income lower than " + ObjCD.DigitGrouping(MAXRate, Me.DefPrec) + " .", MsgBoxStyle.Information, "Information")
                                    If Me.Buffer <> "" Then
                                        Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec)
                                    Else
                                        Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    If Me.ApakahCellRateKosong(Row) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    End If
                                    Me.dgvRate.AllowUserToAddRows = True
                                    Exit Sub
                                End If
                            End If
                            'batas bawah
                            If e.RowIndex < Me.dgvRate.Rows.Count - 1 Then
                                Dim MAXRate As Decimal = Me.dgvRate.Item(e.ColumnIndex, RowIndex_Bawah).Value
                                If CDec(Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value) > MAXRate Then
                                    MsgBox("You cannot input maximum income larger than " + ObjCD.DigitGrouping(MAXRate, Me.DefPrec) + " .", MsgBoxStyle.Information, "Information")
                                    If Me.Buffer <> "" Then
                                        Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec)
                                    Else
                                        Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    If Me.ApakahCellRateKosong(Row) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    End If
                                    Me.dgvRate.AllowUserToAddRows = True
                                    Exit Sub
                                End If
                            End If
                            Me.SettingPajakPPH21Berubah = True
                        End If
                        If Me.dgvRate.Item(Me.ColRate_Tarif.Index, e.RowIndex).Value = "" Or CDec(Me.dgvRate.Item(Me.ColRate_Tarif.Index, e.RowIndex).Value) = 0 Then
                            Me.dgvRate.Item(Me.ColRate_Tarif.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(1, Me.DefPrec_Persen)
                            Me.SettingPajakPPH21Berubah = True
                        End If
                    End If
                    If e.ColumnIndex = Me.ColRate_Tarif.Index Then
                        If Me.ObjCD.CekBilanganPositif(.Cells(Me.ColRate_Tarif.Index).Value) = False Then
                            MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                            If Me.Buffer <> "" Then
                                Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec_Persen)
                            Else
                                Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            If Me.ApakahCellRateKosong(Row) = True Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            End If
                            Me.dgvRate.AllowUserToAddRows = True
                            Exit Sub
                        Else
                            Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping( _
                            Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec_Persen)
                        End If
                        If CDec(.Cells(Me.ColRate_Tarif.Index).Value) > Me.MAXLength_Tarif Then
                            MsgBox("Maximum value tarif = 100.", MsgBoxStyle.Information, "Information")
                            If Me.Buffer <> "" Then
                                Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec_Persen)
                            Else
                                Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            If Me.ApakahCellRateKosong(Row) = True Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            End If
                            Me.dgvRate.AllowUserToAddRows = True
                            Exit Sub
                        Else
                            Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping( _
                            Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec_Persen)
                        End If
                        If Me.dgvRate.Item(Me.ColRate_Maximum.Index, e.RowIndex).Value = "0" Or _
                        Me.dgvRate.Item(Me.ColRate_Maximum.Index, e.RowIndex).Value = "" Then
                            MsgBox("Please input maximum income.", MsgBoxStyle.Information, "Information")
                            If Me.Buffer <> "" Then
                                Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec_Persen)
                            Else
                                Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            If Me.ApakahCellRateKosong(Row) = True Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            End If
                            Me.dgvRate.AllowUserToAddRows = True
                            Exit Sub
                        End If
                        Me.SettingPajakPPH21Berubah = True
                    End If
                    If e.ColumnIndex = Me.ColRate_Remarks.Index Then
                        If Me.ObjCD.CekDataKosong(.Cells(Me.ColRate_Remarks.Index).Value, False) = False Then
                            MsgBox("Data shouldn't include the character: ' , % , or * .", MsgBoxStyle.Information, "Information")
                            Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If Me.ApakahCellRateKosong(Row) = True Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            End If
                            Me.dgvRate.AllowUserToAddRows = True
                            Exit Sub
                        End If
                        Me.SettingPajakPPH21Berubah = True
                    End If
                    If Me.ApakahCellRateKosong(Row, False) = True Then
                        Me.EraseRowIndex = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                    End If
                    Me.dgvRate.AllowUserToAddRows = True
                Else
                    If e.ColumnIndex = Me.ColRate_Remarks.Index Then
                        Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        Me.SettingPajakPPH21Berubah = True
                    Else
                        Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeFieldSettingLoanPayroll()
        Dim Dummy As String = ""
        Try
            'Gain Akun
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_LoanPayment_PrimaryGainAccount)
            If Dummy = "" Then
                Me.TxtPrimaryGainAkun.Text = "-1"
                Me.TxtGainAkunCode.Text = ""
            Else
                Me.TxtPrimaryGainAkun.Text = Dummy
                Me.TxtGainAkunCode.Text = Me.ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryGainAkun.Text))
            End If
            Me.TxtGainAkunDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryGainAkun.Text)
            'Loss Akun
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_LoanPayment_PrimaryLossAccount)
            If Dummy = "" Then
                Me.TxtPrimaryLossAkun.Text = "-1"
                Me.TxtLossAkunCode.Text = ""
            Else
                Me.TxtPrimaryLossAkun.Text = Dummy
                Me.TxtLossAkunCode.Text = Me.ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryLossAkun.Text))
            End If
            Me.TxtLossAkunDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryLossAkun.Text)
            'Cash In Akun
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_LoanPayment_PrimaryCashInAccount)
            If Dummy = "" Then
                Me.TxtPrimaryCashInAkun.Text = "-1"
                Me.TxtCashInAkunCode.Text = ""
            Else
                Me.TxtPrimaryCashInAkun.Text = Dummy
                Me.TxtCashInAkunCode.Text = Me.ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryCashInAkun.Text))
            End If
            Me.TxtCashInAkunDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryCashInAkun.Text)
            'Cash Akun
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_Salary_PrimaryCashAccount)
            If Dummy = "" Then
                Me.TxtPrimaryCashAccount.Text = "-1"
                Me.TxtCashAkunCode.Text = ""
            Else
                Me.TxtPrimaryCashAccount.Text = Dummy
                Me.TxtCashAkunCode.Text = Me.ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryCashAccount.Text))
            End If
            Me.TxtCashAkunDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryCashAccount.Text)
            'Tax Akun
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_Salary_PrimaryTaxExpenseAccount)
            If Dummy = "" Then
                Me.TxtPrimaryTaxAkun.Text = "-1"
                Me.TxtTaxAkunCode.Text = ""
            Else
                Me.TxtPrimaryTaxAkun.Text = Dummy
                Me.TxtTaxAkunCode.Text = Me.ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryTaxAkun.Text))
            End If
            Me.TxtTaxAkunDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryTaxAkun.Text)
            'Custom Print for Salary Slip
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SalaryCalculation, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic)
            If Dummy = "" Then
                Me.chkUseCustomPrint_SalaryCalc.Checked = False
            Else
                If Dummy = 0 Then
                    Me.chkUseCustomPrint_SalaryCalc.Checked = False
                Else
                    Me.chkUseCustomPrint_SalaryCalc.Checked = True
                End If
            End If
            'Tax Payable Akun
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_Salary_PrimaryTaxPayableAccount1)
            If Dummy = "" Then
                Me.TxtPrimaryTaxPayableAccount.Text = "-1"
                Me.TxtTaxPayableAccount.Text = ""
            Else
                Me.TxtPrimaryTaxPayableAccount.Text = Dummy
                Me.TxtTaxPayableAccount.Text = Me.ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryTaxPayableAccount.Text))
            End If
            Me.TxtTaxPayableAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryTaxPayableAccount.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanDataSettingLoanPayroll()
        Try
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_LoanPayment_PrimaryGainAccount, Me.TxtPrimaryGainAkun.Text)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_LoanPayment_PrimaryLossAccount, Me.TxtPrimaryLossAkun.Text)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_LoanPayment_PrimaryCashInAccount, Me.TxtPrimaryCashInAkun.Text)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_Salary_PrimaryCashAccount, Me.TxtPrimaryCashAccount.Text)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_Salary_PrimaryTaxExpenseAccount, Me.TxtPrimaryTaxAkun.Text)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_Salary_PrimaryTaxPayableAccount1, Me.TxtPrimaryTaxPayableAccount.Text)
            If Me.chkUseCustomPrint_SalaryCalc.Checked = True Then
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SalaryCalculation, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic, 1)
            Else
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SalaryCalculation, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic, 0)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeFieldSettingPph21()
        Try
            Me.dgvPTKP.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvRate.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            'Cek Apakah Sudah Ada Data Yang Disimpan Ke database
            If Me.ObjPTKP.ApakahSudahAda_PTKP() = True Then
                'load data
                Me.ObjPTKP.IsiDGVItem(Me.dgvPTKP, Me.dgvRate, Me.FrmHR)
            Else
                'isi data baru
                Me.ObjPTKP.IsiDGVItem_Default(Me.dgvPTKP, Me.dgvRate, Me.FrmHR)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvRate_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvRate.UserDeletingRow
        Try
            Me.dgvRate.AllowUserToAddRows = False
            If Me.dgvRate.Rows.Count <= 1 Then
                e.Cancel = True
            End If
            Me.dgvRate.AllowUserToAddRows = True
            Me.SettingPajakPPH21Berubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtGainAkunCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtGainAkunCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtGainAkunCode.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtGainAkunCode.Text) = True Then
                        Me.TxtPrimaryGainAkun.Text = CStr(ObjAkun.GetPrimary(Me.TxtGainAkunCode.Text))
                        Me.TxtGainAkunCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryGainAkun.Text))
                        Me.TxtLossAkunCode.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.BtnFindGainAkun.PerformClick()
                    End If
                Else
                    Me.BtnFindGainAkun.PerformClick()
                End If
            End If
            Me.TxtGainAkunDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryGainAkun.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtGainAkunCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtGainAkunCode.Leave
        Dim ij As New AccModule.FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtGainAkunCode.Text <> "" Then
                If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtGainAkunCode.Text) = True Then
                    Me.TxtPrimaryGainAkun.Text = CStr(ObjAkun.GetPrimary(Me.TxtGainAkunCode.Text))
                    Me.TxtGainAkunCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryGainAkun.Text))
                Else
                    If Me.TxtGainAkunCode.Text <> "" Then
                        ij.txtCari.Text = Me.TxtGainAkunCode.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryGainAkun.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtGainAkunCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryGainAkun.Text))
                    Else
                        Me.TxtGainAkunCode.Clear()
                        Me.TxtPrimaryGainAkun.Text = "-1"
                        Me.TxtLossAkunCode.Focus()
                    End If
                End If
            Else
                Me.TxtGainAkunCode.Clear()
                Me.TxtPrimaryGainAkun.Text = "-1"
            End If
            Me.TxtGainAkunDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryGainAkun.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub BtnFindGainAkun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindGainAkun.Click
        Dim ij As New AccModule.FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtGainAkunCode.Text <> "" Then
                ij.txtCari.Text = Me.TxtGainAkunCode.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtPrimaryGainAkun.Text = CStr(ij.mPrimaryAkun)
                Me.TxtGainAkunCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryGainAkun.Text))
                Me.TxtGainAkunDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryGainAkun.Text)
                Me.TxtLossAkunCode.Focus()
            Else
                Me.TxtGainAkunCode.Clear()
                Me.TxtPrimaryGainAkun.Text = "-1"
                Me.TxtGainAkunDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryGainAkun.Text)
                Me.TxtGainAkunCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtGainAkunCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtGainAkunCode.GotFocus
        Try
            If Me.TxtGainAkunCode.Enabled = True Then
                If Me.TxtGainAkunCode.ReadOnly = True Then
                    Me.TxtGainAkunCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtGainAkunCode.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtGainAkunCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtGainAkunCode.LostFocus
        Try
            If Me.TxtGainAkunCode.Enabled = True Then
                If Me.TxtGainAkunCode.ReadOnly = True Then
                    Me.TxtGainAkunCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtGainAkunCode.BackColor = Me.TxtPrimaryDutystamp.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Account Loss Account
    Private Sub TxtLossAkunCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles TxtLossAkunCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtLossAkunCode.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtLossAkunCode.Text) = True Then
                        Me.TxtPrimaryLossAkun.Text = CStr(ObjAkun.GetPrimary(Me.TxtLossAkunCode.Text))
                        Me.TxtLossAkunCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryLossAkun.Text))
                        Me.TxtCashInAkunCode.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.BtnFindLossAkun.PerformClick()
                    End If
                Else
                    Me.BtnFindLossAkun.PerformClick()
                End If
            End If
            Me.TxtLossAkunDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryLossAkun.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLossAkunCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtLossAkunCode.Leave
        Dim ij As New AccModule.FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtLossAkunCode.Text <> "" Then
                If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtLossAkunCode.Text) = True Then
                    Me.TxtPrimaryLossAkun.Text = CStr(ObjAkun.GetPrimary(Me.TxtLossAkunCode.Text))
                    Me.TxtLossAkunCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryLossAkun.Text))
                Else
                    If Me.TxtLossAkunCode.Text <> "" Then
                        ij.txtCari.Text = Me.TxtLossAkunCode.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryLossAkun.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtLossAkunCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryLossAkun.Text))
                    Else
                        Me.TxtLossAkunCode.Clear()
                        Me.TxtPrimaryLossAkun.Text = "-1"
                        Me.TxtCashInAkunCode.Focus()
                    End If
                End If
            Else
                Me.TxtLossAkunCode.Clear()
                Me.TxtPrimaryLossAkun.Text = "-1"
            End If
            Me.TxtLossAkunDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryLossAkun.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub BtnFindLossAkun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindLossAkun.Click
        Dim ij As New AccModule.FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtLossAkunCode.Text <> "" Then
                ij.txtCari.Text = Me.TxtLossAkunCode.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtPrimaryLossAkun.Text = CStr(ij.mPrimaryAkun)
                Me.TxtLossAkunCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryLossAkun.Text))
                Me.TxtLossAkunDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryLossAkun.Text)
                Me.TxtCashInAkunCode.Focus()
            Else
                Me.TxtLossAkunCode.Clear()
                Me.TxtPrimaryLossAkun.Text = "-1"
                Me.TxtLossAkunDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryLossAkun.Text)
                Me.TxtLossAkunCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtLossAkunCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLossAkunCode.GotFocus
        Try
            If Me.TxtLossAkunCode.Enabled = True Then
                If Me.TxtLossAkunCode.ReadOnly = True Then
                    Me.TxtLossAkunCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtLossAkunCode.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLossAkunCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLossAkunCode.LostFocus
        Try
            If Me.TxtLossAkunCode.Enabled = True Then
                If Me.TxtLossAkunCode.ReadOnly = True Then
                    Me.TxtLossAkunCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtLossAkunCode.BackColor = Me.TxtPrimaryDutystamp.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Account Cash Account
    Private Sub TxtCashAkunCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles TxtCashAkunCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtCashAkunCode.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtCashAkunCode.Text) = True Then
                        Me.TxtPrimaryCashAccount.Text = CStr(ObjAkun.GetPrimary(Me.TxtCashAkunCode.Text))
                        Me.TxtCashAkunCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryCashAccount.Text))
                        Me.TxtTaxAkunCode.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.BtnFindCashAkun.PerformClick()
                    End If
                Else
                    Me.BtnFindCashAkun.PerformClick()
                End If
            End If
            Me.TxtCashAkunDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryCashAccount.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCashAkunCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtCashAkunCode.Leave
        Dim ij As New AccModule.FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCashAkunCode.Text <> "" Then
                If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtCashAkunCode.Text) = True Then
                    Me.TxtPrimaryCashAccount.Text = CStr(ObjAkun.GetPrimary(Me.TxtCashAkunCode.Text))
                    Me.TxtCashAkunCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryCashAccount.Text))
                Else
                    If Me.TxtCashAkunCode.Text <> "" Then
                        ij.txtCari.Text = Me.TxtCashAkunCode.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryCashAccount.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtCashAkunCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryCashAccount.Text))
                    Else
                        Me.TxtCashAkunCode.Clear()
                        Me.TxtPrimaryCashAccount.Text = "-1"
                        Me.TxtTaxAkunCode.Focus()
                    End If
                End If
            Else
                Me.TxtCashAkunCode.Clear()
                Me.TxtPrimaryCashAccount.Text = "-1"
            End If
            Me.TxtCashAkunDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryCashAccount.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub BtnFindCashAkun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCashAkun.Click
        Dim ij As New AccModule.FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCashAkunCode.Text <> "" Then
                ij.txtCari.Text = Me.TxtCashAkunCode.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtPrimaryCashAccount.Text = CStr(ij.mPrimaryAkun)
                Me.TxtCashAkunCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryCashAccount.Text))
                Me.TxtCashAkunDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryCashAccount.Text)
                Me.TxtTaxAkunCode.Focus()
            Else
                Me.TxtCashAkunCode.Clear()
                Me.TxtPrimaryCashAccount.Text = "-1"
                Me.TxtCashAkunDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryCashAccount.Text)
                Me.TxtCashAkunCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtCashAkunCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCashAkunCode.GotFocus
        Try
            If Me.TxtCashAkunCode.Enabled = True Then
                If Me.TxtCashAkunCode.ReadOnly = True Then
                    Me.TxtCashAkunCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtCashAkunCode.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCashAkunCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCashAkunCode.LostFocus
        Try
            If Me.TxtCashAkunCode.Enabled = True Then
                If Me.TxtCashAkunCode.ReadOnly = True Then
                    Me.TxtCashAkunCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtCashAkunCode.BackColor = Me.TxtPrimaryDutystamp.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Account Cash Account
    Private Sub TxtTaxAkunCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles TxtTaxAkunCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtTaxAkunCode.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtTaxAkunCode.Text) = True Then
                        Me.TxtPrimaryTaxAkun.Text = CStr(ObjAkun.GetPrimary(Me.TxtTaxAkunCode.Text))
                        Me.TxtTaxAkunCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryTaxAkun.Text))
                        Me.TxtTaxPayableAccount.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.BtnFindTaxAkun.PerformClick()
                    End If
                Else
                    Me.BtnFindTaxAkun.PerformClick()
                End If
            End If
            Me.TxtTaxAkunDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryTaxAkun.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxAkunCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtTaxAkunCode.Leave
        Dim ij As New AccModule.FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtTaxAkunCode.Text <> "" Then
                If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtTaxAkunCode.Text) = True Then
                    Me.TxtPrimaryTaxAkun.Text = CStr(ObjAkun.GetPrimary(Me.TxtTaxAkunCode.Text))
                    Me.TxtTaxAkunCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryTaxAkun.Text))
                Else
                    If Me.TxtTaxAkunCode.Text <> "" Then
                        ij.txtCari.Text = Me.TxtTaxAkunCode.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryTaxAkun.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtTaxAkunCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryTaxAkun.Text))
                    Else
                        Me.TxtTaxAkunCode.Clear()
                        Me.TxtPrimaryTaxAkun.Text = "-1"
                        Me.TxtTaxPayableAccount.Focus()
                    End If
                End If
            Else
                Me.TxtTaxAkunCode.Clear()
                Me.TxtPrimaryTaxAkun.Text = "-1"
            End If
            Me.TxtTaxAkunDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryTaxAkun.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub BtnFindTaxAkun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindTaxAkun.Click
        Dim ij As New AccModule.FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtTaxAkunCode.Text <> "" Then
                ij.txtCari.Text = Me.TxtTaxAkunCode.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtPrimaryTaxAkun.Text = CStr(ij.mPrimaryAkun)
                Me.TxtTaxAkunCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryTaxAkun.Text))
                Me.TxtTaxAkunDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryTaxAkun.Text)
                Me.TxtTaxPayableAccount.Focus()
            Else
                Me.TxtTaxAkunCode.Clear()
                Me.TxtPrimaryTaxAkun.Text = "-1"
                Me.TxtTaxAkunDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryTaxAkun.Text)
                Me.TxtTaxAkunCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtTaxAkunCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxAkunCode.GotFocus
        Try
            If Me.TxtTaxAkunCode.Enabled = True Then
                If Me.TxtTaxAkunCode.ReadOnly = True Then
                    Me.TxtTaxAkunCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtTaxAkunCode.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxAkunCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxAkunCode.LostFocus
        Try
            If Me.TxtTaxAkunCode.Enabled = True Then
                If Me.TxtTaxAkunCode.ReadOnly = True Then
                    Me.TxtTaxAkunCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtTaxAkunCode.BackColor = Me.TxtPrimaryDutystamp.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCashInAkunCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles TxtCashInAkunCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtCashInAkunCode.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtCashInAkunCode.Text) = True Then
                        Me.TxtPrimaryCashInAkun.Text = CStr(ObjAkun.GetPrimary(Me.TxtCashInAkunCode.Text))
                        Me.TxtCashInAkunCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryCashInAkun.Text))
                        Me.TxtCashAkunCode.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.BtnFindCashInAkun.PerformClick()
                    End If
                Else
                    Me.BtnFindCashInAkun.PerformClick()
                End If
            End If
            Me.TxtCashInAkunDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryCashInAkun.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCashInAkunCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtCashInAkunCode.Leave
        Dim ij As New AccModule.FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCashInAkunCode.Text <> "" Then
                If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtCashInAkunCode.Text) = True Then
                    Me.TxtPrimaryCashInAkun.Text = CStr(ObjAkun.GetPrimary(Me.TxtCashInAkunCode.Text))
                    Me.TxtCashInAkunCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryCashInAkun.Text))
                Else
                    If Me.TxtCashInAkunCode.Text <> "" Then
                        ij.txtCari.Text = Me.TxtCashInAkunCode.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryCashInAkun.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtCashInAkunCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryCashInAkun.Text))
                    Else
                        Me.TxtCashInAkunCode.Clear()
                        Me.TxtPrimaryCashInAkun.Text = "-1"
                        Me.TxtCashInAkunCode.Focus()
                    End If
                End If
            Else
                Me.TxtCashInAkunCode.Clear()
                Me.TxtPrimaryCashInAkun.Text = "-1"
            End If
            Me.TxtCashInAkunDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryCashInAkun.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub BtnFindCashInAkun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCashInAkun.Click
        Dim ij As New AccModule.FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCashInAkunCode.Text <> "" Then
                ij.txtCari.Text = Me.TxtCashInAkunCode.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtPrimaryCashInAkun.Text = CStr(ij.mPrimaryAkun)
                Me.TxtCashInAkunCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryCashInAkun.Text))
                Me.TxtCashInAkunDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryCashInAkun.Text)
                Me.TxtCashAkunCode.Focus()
            Else
                Me.TxtCashInAkunCode.Clear()
                Me.TxtPrimaryCashInAkun.Text = "-1"
                Me.TxtCashInAkunDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryCashInAkun.Text)
                Me.TxtCashInAkunCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtCashInAkunCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCashInAkunCode.GotFocus
        Try
            If Me.TxtCashInAkunCode.Enabled = True Then
                If Me.TxtCashInAkunCode.ReadOnly = True Then
                    Me.TxtCashInAkunCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtCashInAkunCode.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCashInAkunCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCashInAkunCode.LostFocus
        Try
            If Me.TxtCashInAkunCode.Enabled = True Then
                If Me.TxtCashInAkunCode.ReadOnly = True Then
                    Me.TxtCashInAkunCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtCashInAkunCode.BackColor = Me.TxtPrimaryDutystamp.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkUseCustomPrint_SalaryCalc_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkUseCustomPrint_SalaryCalc.CheckedChanged
        Try
            Me.SettingHRBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TaxPayableAccount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles TxtTaxPayableAccount.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtTaxPayableAccount.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtTaxPayableAccount.Text) = True Then
                        Me.TxtPrimaryTaxPayableAccount.Text = CStr(ObjAkun.GetPrimary(Me.TxtTaxPayableAccount.Text))
                        Me.TxtTaxPayableAccount.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryTaxPayableAccount.Text))
                        Me.dgvPTKP.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.btnFindTaxPayableAccount.PerformClick()
                    End If
                Else
                    Me.btnFindTaxPayableAccount.PerformClick()
                End If
            End If
            Me.TxtTaxPayableAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryTaxPayableAccount.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TaxPayableAccount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtTaxPayableAccount.Leave
        Dim ij As New AccModule.FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtTaxPayableAccount.Text <> "" Then
                If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtTaxPayableAccount.Text) = True Then
                    Me.TxtPrimaryTaxPayableAccount.Text = CStr(ObjAkun.GetPrimary(Me.TxtTaxPayableAccount.Text))
                    Me.TxtTaxPayableAccount.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryTaxPayableAccount.Text))
                Else
                    If Me.TxtTaxPayableAccount.Text <> "" Then
                        ij.txtCari.Text = Me.TxtTaxPayableAccount.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryTaxPayableAccount.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtTaxPayableAccount.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryTaxPayableAccount.Text))
                    Else
                        Me.TxtTaxPayableAccount.Clear()
                        Me.TxtPrimaryTaxPayableAccount.Text = "-1"
                        Me.dgvPTKP.Focus()
                    End If
                End If
            Else
                Me.TxtTaxPayableAccount.Clear()
                Me.TxtPrimaryTaxPayableAccount.Text = "-1"
            End If
            Me.TxtTaxPayableAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryTaxPayableAccount.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnFindTaxPayableAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindTaxPayableAccount.Click
        Dim ij As New AccModule.FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtTaxPayableAccount.Text <> "" Then
                ij.txtCari.Text = Me.TxtTaxPayableAccount.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtPrimaryTaxPayableAccount.Text = CStr(ij.mPrimaryAkun)
                Me.TxtTaxPayableAccount.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryTaxPayableAccount.Text))
                Me.TxtTaxPayableAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryTaxPayableAccount.Text)
                Me.dgvPTKP.Focus()
            Else
                Me.TxtTaxPayableAccount.Clear()
                Me.TxtPrimaryTaxPayableAccount.Text = "-1"
                Me.TxtTaxPayableAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryTaxPayableAccount.Text)
                Me.TxtTaxPayableAccount.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TaxPayableAccount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxPayableAccount.GotFocus
        Try
            If Me.TxtTaxPayableAccount.Enabled = True Then
                If Me.TxtTaxPayableAccount.ReadOnly = True Then
                    Me.TxtTaxPayableAccount.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtTaxPayableAccount.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TaxPayableAccount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxPayableAccount.LostFocus
        Try
            If Me.TxtTaxPayableAccount.Enabled = True Then
                If Me.TxtTaxPayableAccount.ReadOnly = True Then
                    Me.TxtTaxPayableAccount.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtTaxPayableAccount.BackColor = Me.TxtPrimaryDutystamp.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
