Option Explicit On
Option Strict Off
Imports Core_Module, System.Drawing, System.Windows.Forms
Public Class FrmTrans_Gaji_Quick2
    Public UserCode As String = ""
    Public PeriodFrom As Date = #1/1/1950#
    Public PeriodTo As Date = #1/1/1950#
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim ObjAkun As New AccModule.ClsMasterAkun
    Dim ObjUser As New ClsUser
    Dim AllData As Boolean = True
    Dim EraseRowIndex As Integer = 0
    Dim ObjProduct As New Core_Module.ClsMasterKaryawan
    Dim Buffer As String = ""
    Dim ListDataFilter As New List(Of Double)
    Dim TBLKary As New DseTransaksiSalary_Quick.tblmaster_karyawanDataTable
    Dim FrmSalary As New FrmTrans_Gaji
    Dim PrimaryCurr_Default As Double = 1
    Dim ObjPay As New ClsTransaksiLoanPayment
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjSalary As New ClsTransaksiSalary
    Dim ObjKary As New Core_Module.ClsMasterKaryawan
    Dim ObjCD As New Core_Module.ClsCekData
    Dim PrimaryCashIn As String = "-1"
    Dim PrimaryGain As String = "-1"
    Dim PrimaryLoss As String = "-1"
    Dim mShowMsg As Boolean = False
    Dim RefText As String = ""

    Dim StopProses As Boolean = False
    Dim OnProgress As Boolean = False
    Dim RowKary As DseTransaksiSalary_Quick.tblmaster_karyawanRow
    Dim i As Integer = 0

    Dim TBL_JP As New DseTransaksiSalary_Quick.tblmaster_jobpositionDataTable
    Dim TBL_Dept As New DseTransaksiSalary_Quick.tblmaster_departemenDataTable
    Dim TBL_Formula As New DseTransaksiSalary_Quick.tblmaster_salary_formula1DataTable
    Dim TBLCash As New DseTransaksiLoanPayment.tblmaster_akunDataTable


    'Dim TBLSL_Kary_Ref As New DseTransaksiSalary.tblmaster_salary_formula_component_detailemployeeDataTable
    'Dim mPrimarySalaryFormulaRef As Double = -1

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If MsgBox("Cancel creating calculate salary?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.StopProses = True
                'Me.Close()
            End If
            'Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RemoveDataEmployeeYgTdkAdaAttendance(ByVal TBLMsgBox As DseTransaksiSalary.TBLMessageBoxDataTable, _
    ByVal ListString As List(Of String), ByVal vShowMsgBox As Boolean)
        Dim RowKary As DseTransaksiSalary_Quick.tblmaster_karyawanRow
        Dim RowMsgBox As DseTransaksiSalary.TBLMessageBoxRow
        Dim Counter As Integer = 0
        Dim Ketemu As Boolean = False
        Dim ObjClosing As New AccModule.ClsTransaksiClosing
        Dim ConventionDate As Date = ObjClosing.GetPayrollConvertionDate
        Dim KaryAktif As Boolean = True
        Dim JmlCekAtt As Integer = 0
        Dim vFirstTaxPeriod As Date = #1/1/1950#
        Dim vJmlAcuan As Integer = 0

        Try
            If TBLKary.Rows.Count > 0 Then
                While Counter <= Me.TBLKary.Rows.Count - 1
                    RowKary = Me.TBLKary.Rows(Counter)
                    Ketemu = False
                    KaryAktif = True
                    JmlCekAtt = 0

                    'Cek Berkaitan dengan attendance 
                    If RowKary.Aktif = CInt(ClsMasterKaryawan.EnumAktif.Active) Then
                        KaryAktif = True
                    Else
                        KaryAktif = False
                    End If
                    'KaryAktif = Me.ObjKary.ApakahAktif(RowKary.Primary)
                    If KaryAktif = True Then
cekAtt:
                        JmlCekAtt += 1
                        If Me.ObjSalary.CekValidPeriodEmp_HubunganAttendance(RowKary.Primary, Me.dtpPeriod.Value, Me.DtpTo.Value, Me.mShowMsg) = False Then
                            If vShowMsgBox = True Then
                                If ListString.Contains(RowKary.Primary.ToString + "-" + "Calculate salary period isn't suitable with period attendance.") = False Then
                                    ListString.Add(RowKary.Primary.ToString + "-" + "Incomplete attendance data.")
                                    RowMsgBox = TBLMsgBox.NewRow
                                    RowMsgBox.Employee_Code = RowKary.Code
                                    RowMsgBox.Employee_Name = RowKary.Full_Name
                                    RowMsgBox.Reason = "Incomplete attendance data."
                                    TBLMsgBox.Rows.Add(RowMsgBox)
                                End If
                            End If
                            Ketemu = True
                        End If
                    End If
                    If Ketemu = False Then
                        If KaryAktif = False Then
                            If RowKary.Resign_Date.Date < Me.dtpPeriod.Value.Date Then
                                If vShowMsgBox = True Then
                                    If ListString.Contains(RowKary.Primary.ToString + "-" + "Inactive employee.") = False Then
                                        ListString.Add(RowKary.Primary.ToString + "-" + "Inactive employee.")
                                        RowMsgBox = TBLMsgBox.NewRow
                                        RowMsgBox.Employee_Code = RowKary.Code
                                        RowMsgBox.Employee_Name = RowKary.Full_Name
                                        RowMsgBox.Reason = "Inactive employee."
                                        TBLMsgBox.Rows.Add(RowMsgBox)
                                    End If
                                End If
                                Ketemu = True
                            Else
                                If JmlCekAtt <= 1 Then
                                    GoTo cekAtt
                                End If
                            End If
                        End If
                    End If

                    If Ketemu = False Then
                        'Cek apakah ada yang tumpang tindih
                        If Me.ObjSalary.FillTBLStartEnd(RowKary.Primary, Me.DtpTo.Value, Me.dtpPeriod.Value, True, -1, Me.mShowMsg) = False Then
                            If vShowMsgBox = True Then
                                If ListString.Contains(RowKary.Primary.ToString + "-" + "Conflicting with existing salary calculation.") = False Then
                                    ListString.Add(RowKary.Primary.ToString + "-" + "Conflicting with existing salary calculation.")
                                    RowMsgBox = TBLMsgBox.NewRow
                                    RowMsgBox.Employee_Code = RowKary.Code
                                    RowMsgBox.Employee_Name = RowKary.Full_Name
                                    RowMsgBox.Reason = "Conflicting with existing salary calculation."
                                    TBLMsgBox.Rows.Add(RowMsgBox)
                                End If
                            End If
                            Ketemu = True
                        End If
                    End If
                    If Ketemu = False Then
                        '1 bulan 1 transaksi salary
                        If Me.cboTaxCalculation.SelectedIndex <> 3 Then
                            If Me.ObjSalary.GetJmlByTanggalDariSampai(Me.ObjCD.GetLasDateOfTheMonth(Me.DtpDate.Value), _
                                Me.ObjCD.GetFirstDateofTheMonth(Me.DtpDate.Value), RowKary.Primary) > 0 Then
                                If vShowMsgBox = True Then
                                    If ListString.Contains(RowKary.Primary.ToString + "-" + "There should only be one of salary for each month.") = False Then
                                        ListString.Add(RowKary.Primary.ToString + "-" + "There should only be one of salary for each month.")
                                        RowMsgBox = TBLMsgBox.NewRow
                                        RowMsgBox.Employee_Code = RowKary.Code
                                        RowMsgBox.Employee_Name = RowKary.Full_Name
                                        RowMsgBox.Reason = "There should only be one of salary for each month."
                                        TBLMsgBox.Rows.Add(RowMsgBox)
                                    End If
                                End If
                                'MsgBox("Data can't save because in this month should be create 1 transaction.", MsgBoxStyle.Information, "Information")
                                'Me.DtpDate.Focus()
                                Ketemu = True
                            End If
                        End If
                    End If
                    'If Ketemu = False Then
                    '    If Me.cboTaxCalculation.SelectedIndex <> 3 Then
                    '        If Me.ObjSalary.GetJmlByTanggalDariSampai(Me.ObjCD.GetLasDateOfTheMonth(Me.DtpDate.Value), _
                    '                                Me.ObjCD.GetFirstDateofTheMonth(Me.DtpDate.Value), RowKary.Primary) > 0 Then
                    '            If vShowMsgBox = True Then
                    '                If ListString.Contains(RowKary.Primary.ToString + "-" + "Data can't save because in this month should be create 1 transaction.") = False Then
                    '                    ListString.Add(RowKary.Primary.ToString + "-" + "Data can't save because in this month should be create 1 transaction.")
                    '                    RowMsgBox = TBLMsgBox.NewRow
                    '                    RowMsgBox.Employee_Code = RowKary.Code
                    '                    RowMsgBox.Employee_Name = RowKary.Full_Name
                    '                    RowMsgBox.Reason = "Data can't save because in this month should be create 1 transaction."
                    '                    TBLMsgBox.Rows.Add(RowMsgBox)
                    '                End If
                    '            End If
                    '            'MsgBox("Data can't save because in this month should be create 1 transaction.", MsgBoxStyle.Information, "Information")
                    '            'Me.DtpDate.Focus()
                    '            Ketemu = True
                    '        End If
                    '    End If
                    'End If
                    If Ketemu = False Then
                        If Me.DtpDate.Value.Date.Month = 12 And Me.DtpDate.Value.Year = ConventionDate.Year Then
                            If Me.cboTaxCalculation.SelectedIndex <> 3 Then
                                Dim MonthConvertion As Integer = ConventionDate.Month - 1
                                Dim DateConvention As Date = ConventionDate
                                'Dim Hasil_1 As MsgBoxResult
                                If Me.ObjSalary.GetJmlByPeriodTaxBeginning(ConventionDate, Me.ObjCD, RowKary.Primary) <> MonthConvertion Then
                                    If vShowMsgBox = True Then
                                        If ListString.Contains(RowKary.Primary.ToString + "-" + "This employee doesn't have tax beginning balance calculation for the following month : Jan - " + Format(DateConvention, "MMM").ToString) = False Then
                                            ListString.Add(RowKary.Primary.ToString + "-" + "This employee doesn't have tax beginning balance calculation for the following month : Jan - " + Format(DateConvention, "MMM").ToString)
                                            RowMsgBox = TBLMsgBox.NewRow
                                            RowMsgBox.Employee_Code = RowKary.Code
                                            RowMsgBox.Employee_Name = RowKary.Full_Name
                                            RowMsgBox.Reason = "This employee doesn't have tax beginning balance calculation for the following month : Jan - " + Format(DateConvention, "MMM").ToString
                                            TBLMsgBox.Rows.Add(RowMsgBox)
                                        End If
                                    End If
                                    'Hasil_1 = MsgBox("This employee doesn't have tax beginning balance calculation for the following month : Jan - " _
                                    ' + Format(DateConvention, "MMM").ToString + " Continue ?", MsgBoxStyle.YesNo, "Confirm")
                                    'If Hasil_1 = MsgBoxResult.No Then
                                    Ketemu = True
                                    'End If
                                End If
                            End If

                        End If
                    End If

                    vFirstTaxPeriod = ObjKary.GetFirstTaxPeriod(RowKary.Primary)
                    If Ketemu = False Then
                        '7. Batasan dengan Tax Beginning Balance
                        If Me.DtpDate.Value.Date.Month = 12 Then
                            If Me.cboTaxCalculation.SelectedIndex <> 3 Then
                                Dim YearConvertion As Integer = ConventionDate.Year
                                If YearConvertion <> Me.DtpDate.Value.Year Then
                                    'Dim Hasil_2 As MsgBoxResult
                                    Dim mFirstDateYear As Date = Me.ObjCD.GetFisrtDateOfTheYear(Me.DtpDate.Value.Date)
                                    Dim mLastDateYear As Date = "#11/30/" + Me.DtpDate.Value.Year.ToString + "#"
                                    If vFirstTaxPeriod.Year < Me.DtpDate.Value.Year Then
                                        If Me.ObjSalary.GetJmlByTanggalDariSampai_BuatAntaraTaxDenganSalary(mLastDateYear, mFirstDateYear, RowKary.Primary) <> 11 Then
                                            If vShowMsgBox = True Then
                                                If ListString.Contains(RowKary.Primary.ToString + "-" + "This employee doesn't have salary calculation for the following month : Jan - " + Format(CDate("#11/1/1950#"), "MMM").ToString) = False Then
                                                    ListString.Add(RowKary.Primary.ToString + "-" + "This employee doesn't have salary calculation for the following month : Jan - " + Format(CDate("#11/1/1950#"), "MMM").ToString)
                                                    RowMsgBox = TBLMsgBox.NewRow
                                                    RowMsgBox.Employee_Code = RowKary.Code
                                                    RowMsgBox.Employee_Name = RowKary.Full_Name
                                                    RowMsgBox.Reason = "This employee doesn't have salary calculation for the following month : Jan - " + Format(CDate("#11/1/1950#"), "MMM").ToString
                                                    TBLMsgBox.Rows.Add(RowMsgBox)
                                                End If
                                            End If
                                            Ketemu = True
                                        End If
                                    Else
                                        If vFirstTaxPeriod.Month <> 12 Then
                                            vJmlAcuan = 12 - vFirstTaxPeriod.Month
                                            If Me.ObjSalary.GetJmlByTanggalDariSampai_BuatAntaraTaxDenganSalary(mLastDateYear, mFirstDateYear, RowKary.Primary) <> vJmlAcuan Then
                                                If vShowMsgBox = True Then
                                                    If ListString.Contains(RowKary.Primary.ToString + "-" + "This employee doesn't have salary calculation for the following month : " + Format(CDate("#" + vFirstTaxPeriod.Month.ToString + "/1/1950#"), "MMM").ToString + " - " + Format(CDate("#11/1/1950#"), "MMM").ToString) = False Then
                                                        ListString.Add(RowKary.Primary.ToString + "-" + "This employee doesn't have salary calculation for the following month : " + Format(CDate("#" + vFirstTaxPeriod.Month.ToString + "/1/1950#"), "MMM").ToString + " - " + Format(CDate("#11/1/1950#"), "MMM").ToString)
                                                        RowMsgBox = TBLMsgBox.NewRow
                                                        RowMsgBox.Employee_Code = RowKary.Code
                                                        RowMsgBox.Employee_Name = RowKary.Full_Name
                                                        RowMsgBox.Reason = "This employee doesn't have salary calculation for the following month : " + Format(CDate("#" + vFirstTaxPeriod.Month.ToString + "/1/1950#"), "MMM").ToString + " - " + Format(CDate("#11/1/1950#"), "MMM").ToString
                                                        TBLMsgBox.Rows.Add(RowMsgBox)
                                                    End If
                                                End If
                                                Ketemu = True
                                            End If
                                        End If
                                    End If
                                Else
                                    Dim MonthConvertion As Integer = ConventionDate.Month
                                    'Dim Hasil_2 As MsgBoxResult
                                    Dim mFirstDateYear As Date = Me.ObjCD.GetFisrtDateOfTheYear(Me.DtpDate.Value.Date)
                                    Dim mLastDateYear As Date = "#11/30/" + Me.DtpDate.Value.Year.ToString + "#"
                                    If 11 - Me.ObjSalary.GetJmlByPeriodTaxBeginning(ConventionDate, Me.ObjCD, RowKary.Primary) <> Me.ObjSalary.GetJmlByTanggalDariSampai_BuatAntaraTaxDenganSalary(mLastDateYear, mFirstDateYear, RowKary.Primary) Then
                                        If vShowMsgBox = True Then
                                            If ListString.Contains(RowKary.Primary.ToString + "-" + "This employee doesn't have salary calculation for the following month : " + Format(CDate("#" + MonthConvertion.ToString + "/1/1950#"), "MMM") + " - " + Format(CDate("#11/1/1950#"), "MMM").ToString) = False Then
                                                ListString.Add(RowKary.Primary.ToString + "-" + "This employee doesn't have salary calculation for the following month : " + Format(CDate("#" + MonthConvertion.ToString + "/1/1950#"), "MMM") + " - " + Format(CDate("#11/1/1950#"), "MMM").ToString)
                                                RowMsgBox = TBLMsgBox.NewRow
                                                RowMsgBox.Employee_Code = RowKary.Code
                                                RowMsgBox.Employee_Name = RowKary.Full_Name
                                                RowMsgBox.Reason = "This employee doesn't have salary calculation for the following month : " + Format(CDate("#" + MonthConvertion.ToString + "/1/1950#"), "MMM") + " - " + Format(CDate("#11/1/1950#"), "MMM").ToString
                                                TBLMsgBox.Rows.Add(RowMsgBox)
                                            End If
                                        End If
                                        'Hasil_2 = MsgBox("This employee doesn't have salary calculation for the following month : " + Format(CDate("#" + MonthConvertion.ToString + "/1/1950#"), "MMM") + " - " _
                                        '+ Format(CDate("#11/1/1950#"), "MMM").ToString + " Continue ?", MsgBoxStyle.YesNo, "Confirm")
                                        'If Hasil_2 = MsgBoxResult.No Then
                                        Ketemu = True
                                        'End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If Ketemu = False Then
                        '8. batasan bulan 12, cek bulan 1-11 metode tax calculation harus sama
                        If Me.ObjSalary.ApakahSelama1TahunTaxCalculationSama(Me.DtpDate.Value.Date, Me.cboTaxCalculation.SelectedIndex, RowKary.Primary, -1) = False Then
                            If vShowMsgBox = True Then
                                If ListString.Contains(RowKary.Primary.ToString + "-" + "Please cek tax calculation because you must choose the same tax calculation for around one year.") = False Then
                                    ListString.Add(RowKary.Primary.ToString + "-" + "Please cek tax calculation because you must choose the same tax calculation for around one year.")
                                    RowMsgBox = TBLMsgBox.NewRow
                                    RowMsgBox.Employee_Code = RowKary.Code
                                    RowMsgBox.Employee_Name = RowKary.Full_Name
                                    RowMsgBox.Reason = "Please cek tax calculation because you must choose the same tax calculation for around one year."
                                    TBLMsgBox.Rows.Add(RowMsgBox)
                                End If
                            End If
                            Ketemu = True
                        End If
                    End If
                    If Ketemu = False Then
                        '9. Batasan salary dengan tax beginning balance, jika tax calulation = netto/gross maka tidak boleh menginput transaksi di bulan 12
                        If Me.DtpDate.Value.Date.Month = 12 Then
                            If Me.cboTaxCalculation.SelectedIndex = ClsTransaksiSalary.EnumTax.Gross Or Me.cboTaxCalculation.SelectedIndex = ClsTransaksiSalary.EnumTax.Netto Then
                                If Me.ObjSalary.ApakahTunjanganPPhLebihBesarDariNol(Me.DtpDate.Value.Date, RowKary.Primary) = False Then
                                    If vShowMsgBox = True Then
                                        If ListString.Contains(RowKary.Primary.ToString + "-" + "Tunjangan PPh in tax beginning balance have to zero value because tax calculation method in salary transaction is netto or gross.") = False Then
                                            ListString.Add(RowKary.Primary.ToString + "-" + "Tunjangan PPh in tax beginning balance have to zero value because tax calculation method in salary transaction is netto or gross.")
                                            RowMsgBox = TBLMsgBox.NewRow
                                            RowMsgBox.Employee_Code = RowKary.Code
                                            RowMsgBox.Employee_Name = RowKary.Full_Name
                                            RowMsgBox.Reason = "Tunjangan PPh in tax beginning balance have to zero value because tax calculation method in salary transaction is netto or gross."
                                            TBLMsgBox.Rows.Add(RowMsgBox)
                                        End If
                                    End If
                                    Ketemu = True
                                End If
                            End If
                        End If
                    End If
                    If Ketemu = True Then
                        'Hapus baris employee
                        Me.TBLKary.Rows.RemoveAt(Counter)
                    Else
                        Counter += 1
                    End If
                End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjClosing = Nothing
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim TipeAkunDummy As ClsMasterAkun1.EnumTipeAkun = ClsMasterAkun1.EnumTipeAkun.Asset

        Try
            'Engga boleh kosong
            'If Me.TxtPrimaryTaxExpAcc.Text = "-1" Then
            '    MsgBox("Please choose tax expense account.", MsgBoxStyle.Information, "Information")
            '    Me.TxtExpAcc.Focus()
            '    Return False
            'End If

            'Exp Tax Account
            'Dim Counter As Double = 0
            'Dim TotalCounter As Double = 5
            'Dim Ketemu As Boolean = False
            'While Counter <= TotalCounter - 1
            '    If Counter = 0 Then
            '        If CDbl(Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryTaxExpAcc.Text)) = CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Income) Then
            '            Ketemu = True
            '            Exit While
            '        End If
            '    End If
            '    If Counter = 1 Then
            '        If CDbl(Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryTaxExpAcc.Text)) = CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Other_Income) Then
            '            Ketemu = True
            '            Exit While
            '        End If
            '    End If
            '    If Counter = 2 Then
            '        If CDbl(Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryTaxExpAcc.Text)) = CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Expense) Then
            '            Ketemu = True
            '            Exit While
            '        End If
            '    End If
            '    If Counter = 3 Then
            '        If CDbl(Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryTaxExpAcc.Text)) = CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Other_Expense) Then
            '            Ketemu = True
            '            Exit While
            '        End If
            '    End If
            '    If Counter = 4 Then
            '        If CDbl(Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryTaxExpAcc.Text)) = CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Cost_of_Goods_Sales) Then
            '            Ketemu = True
            '            Exit While
            '        End If
            '    End If
            '    Counter += 1
            'End While
            'If Ketemu = False Then
            '    MsgBox("Please input account code with type income account or other income account or expense or other expense or COGS.", MsgBoxStyle.Information, "Information")
            '    Me.TxtExpAcc.Focus()
            '    Return False
            'End If

            If Me.DtpTo.Value.Date < Me.dtpPeriod.Value.Date Then
                MsgBox("Period from can't be grather than period to.", MsgBoxStyle.Information, "Information")
                Me.dtpPeriod.Focus()
                Return False
            End If

            'Tanggal Transaksi Tidak boleh < dari DtpFrom
            If Me.DtpDate.Value.Date < Me.dtpPeriod.Value.Date Then
                MsgBox("Transaction date can't be less than period date.", MsgBoxStyle.Information, "Information")
                Me.DtpDate.Focus()
                Return False
            End If

            If Me.ChkEmployee.Checked = True Then
                Me.dgvEmployee.AllowUserToAddRows = False
                If Me.dgvEmployee.Rows.Count <= 0 Then
                    MsgBox("Please input data employee.", MsgBoxStyle.Information, "Information")
                    Me.dgvEmployee.Focus()
                    Return False
                End If
                Me.dgvEmployee.AllowUserToAddRows = True
            End If

            If Me.chkSalaryFormula.Checked = True Then
                If Me.clbSalaryFormula.CheckedItems.Count <= 0 Then
                    MsgBox("Please selected checklistbox salary formula.", MsgBoxStyle.Information, "Information")
                    Me.clbSalaryFormula.Focus()
                    Return False
                End If
            End If

            If Me.ChkDepartment.Checked = True Then
                If Me.ClbDepartment.CheckedItems.Count <= 0 Then
                    MsgBox("Please selected checklistbox departement.", MsgBoxStyle.Information, "Information")
                    Me.ClbDepartment.Focus()
                    Return False
                End If
            End If

            If Me.chkJobPosition.Checked = True Then
                If Me.ClbPosition.CheckedItems.Count <= 0 Then
                    MsgBox("Please selected checklistbox position.", MsgBoxStyle.Information, "Information")
                    Me.ClbPosition.Focus()
                    Return False
                End If
            End If
            If Me.cboTaxCalculation.SelectedIndex <> 3 Then
                If Format(Me.DtpDate.Value, "MMM yyyy") <> Format(Me.DtpTo.Value.Date, "MMM yyyy") Then
                    MsgBox("Salary calculation period must be in the same month as salary end period.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'Harus isi
            If Me.TxtPrimaryCashAcc.Text = "-1" Then
                MsgBox("Please choose cash account.", MsgBoxStyle.Information, "Information")
                Me.CboCashAccount.Focus()
                Return False
            End If
            If Me.TxtPrimaryGainAcc.Text = "-1" Then
                MsgBox("Please input gain account.", MsgBoxStyle.Information, "Information")
                Me.TxtGainAccount.Focus()
                Return False
            End If
            If Me.TxtPrimaryLoss.Text = "-1" Then
                MsgBox("Please input loss account.", MsgBoxStyle.Information, "Information")
                Me.TxtLossAccount.Focus()
                Return False
            End If
            '6. Cek Account 
            Dim mApakahAkunIncome As Boolean = False
            TipeAkunDummy = Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryGainAcc.Text)
            If TipeAkunDummy <> CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Income) And _
            TipeAkunDummy <> CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Other_Income) And _
            TipeAkunDummy <> CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Asset) And _
            TipeAkunDummy <> CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Other_Expense) And _
            TipeAkunDummy <> CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Cost_of_Goods_Sales) And _
            TipeAkunDummy <> CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Expense) Then
                MsgBox("Please input account code with type income account or other income account.", MsgBoxStyle.Information, "Information")
                Me.TxtGainAccount.Focus()
                Return False
            Else
                mApakahAkunIncome = True
            End If
            If mApakahAkunIncome = False Then
                If CDbl(Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryGainAcc.Text)) <> CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Other_Income) Then
                    MsgBox("Please input account code with type income account or other income account.", MsgBoxStyle.Information, "Information")
                    Me.TxtGainAccount.Focus()
                    Return False
                End If
            End If
            Dim mApakahAkunExpence As Boolean = False
            TipeAkunDummy = CDbl(Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryLoss.Text))
            If TipeAkunDummy <> CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Expense) And _
            TipeAkunDummy <> CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Cost_of_Goods_Sales) And _
            TipeAkunDummy <> CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Asset) And _
            TipeAkunDummy <> CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Other_Income) And _
            TipeAkunDummy <> CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Other_Expense) Then
                MsgBox("Please input account code with type expense account, COGS account, or other expense account.", MsgBoxStyle.Information, "Information")
                Me.TxtLossAccount.Focus()
                Return False
            Else
                mApakahAkunExpence = True
            End If
            If mApakahAkunExpence = False Then
                If CDbl(Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryLoss.Text)) <> CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Other_Expense) Then
                    MsgBox("Please input account code with type expense account, COGS account, or other expense account.", MsgBoxStyle.Information, "Information")
                    Me.TxtLossAccount.Focus()
                    Return False
                End If
            End If
            If Me.ObjAkun.GetPrimaryCurrency(Me.TxtPrimaryGainAcc.Text) <> 1 Then
                MsgBox("Please input gain account code with IDR currency.", MsgBoxStyle.Information, "Information")
                Me.TxtGainAccount.Focus()
                Return False
            End If
            If Me.ObjAkun.GetPrimaryCurrency(Me.TxtPrimaryLoss.Text) <> 1 Then
                MsgBox("Please input loss account code with IDR currency.", MsgBoxStyle.Information, "Information")
                Me.TxtLossAccount.Focus()
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Dim RowKary As DseTransaksiSalary_Quick.tblmaster_karyawanRow
        Dim StartTime As Date
        Dim Dur As TimeSpan

        Try
            StartTime = Now
            ObjCD.UpdateThreading()

            For i As Integer = 0 To Me.TBLKary.Rows.Count - 1
                RowKary = Me.TBLKary.Rows(i)
                If Me.StopProses = True Then
                    Me.BackgroundWorker2.ReportProgress(1)
                    Exit Sub
                End If
                Me.i = i
                Me.RowKary = RowKary

                Me.BackgroundWorker2.ReportProgress(0)
                Threading.Thread.Sleep(3000)
                While Me.OnProgress = True
                    Threading.Thread.Sleep(3000)
                End While

                Dur = (Now - StartTime)
                If Dur.TotalSeconds < 60 Then
                    Me.Text = RefText + " (elapsed time: " + Me.ObjCD.DigitGrouping(Dur.TotalSeconds, 0) + " seconds)"
                ElseIf Dur.TotalMinutes < 60 Then
                    Me.Text = RefText + " (elapsed time: " + Me.ObjCD.DigitGrouping(Dur.TotalMinutes, 1) + " minutes)"
                Else
                    Me.Text = RefText + " (elapsed time: " + Me.ObjCD.DigitGrouping(Dur.TotalHours, 2) + " hours)"
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker2.ProgressChanged
        Try
            If e.ProgressPercentage = 0 Then
                If Me.BuatPayment = False Then
                    'Penambahan yang ketinggalan
                    Me.StopProses = True
                    Exit Sub
                End If
            Else
                MsgBox("Process stopped by user.", MsgBoxStyle.Information, "Information")
                Me.btnCancel.Visible = False
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.btnCancel.Visible = False
            Me.btnCancel.Enabled = False
            Me.btnCancel.TabStop = False

            Me.BtnPause.Visible = False
            Me.BtnPause.Enabled = False
            Me.BtnPause.TabStop = False

            If Me.StopProses = False Then
                MsgBox("Calculate salary creation finish.", MsgBoxStyle.Information, "Information")
                Me.Text = Me.RefText
            End If

            Me.btnOK.Enabled = True
            Me.btnCancel.Enabled = True
            Me.BtnPause.Enabled = True
            Me.ControlBox = True

            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim Hasil As MsgBoxResult
        Dim TBLMsgBox As New DseTransaksiSalary.TBLMessageBoxDataTable
        Dim ListString As New List(Of String)
        Dim FrmError As New FrmTrans_Gaji_MsgBox
        Try
            If Me.ApakahDataBenar() = True Then
                Me.FilterData(TBLMsgBox, ListString, True)
                If TBLMsgBox.Rows.Count > 0 Then
                    TBLMsgBox.DefaultView.Sort = "[Employee Code], [Employee Name]"
                    FrmError.dgvItem.DataSource = TBLMsgBox.DefaultView
                    FrmError.dgvItem.Columns(0).Visible = False
                    FrmError.LblProgress.Text = "Total Rows = " + TBLMsgBox.Rows.Count.ToString
                    FrmError.Text = "The following employees will not be processed."
                    FrmError.ShowDialog()
                End If
                If Me.TBLKary.Rows.Count <= 0 Then
                    MsgBox("No employee data can be processed, please cek attendance data & salary data.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                Hasil = MsgBox("Do you want to process data ?", MsgBoxStyle.YesNo, "Confirm")
                If Hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        If FrmSalary.NewQuick_Data = False Then
                            Me.FrmSalary.Baru = True
                            Me.FrmSalary.UserCode = Me.UserCode
                            Me.FrmSalary.ShowQuick_Edit = True
                            Me.FrmSalary.NewQuick_Data = True
                            Me.FrmSalary.BackColorAll = Me.BackColorAll
                            Me.FrmSalary.Show()
                            Me.FrmSalary.FungsiLoad_Event()
                            Me.FrmSalary.Top = (Me.FrmSalary.Height * -1) - 1
                            Me.FrmSalary.ControlBox = False
                            Me.FrmSalary.WaitingForm = False
                            Me.FrmSalary.Hide()
                        End If

                        Me.btnCancel.Visible = True
                        Me.btnCancel.Enabled = True
                        Me.btnCancel.TabStop = True

                        Me.BtnPause.Visible = True
                        Me.BtnPause.Enabled = True
                        Me.BtnPause.TabStop = True

                        Me.btnOK.Enabled = False
                        Me.ControlBox = False
                        Me.DtpDate.Enabled = False
                        Me.chkDirectlySave.Enabled = False

                        Me.DtpDate.Enabled = False
                        Me.cboTaxCalculation.Enabled = False
                        Me.dtpPeriod.Enabled = False
                        Me.DtpTo.Enabled = False
                        Me.CboCashAccount.Enabled = False
                        Me.TxtGainAccount.Enabled = False
                        Me.TxtLossAccount.Enabled = False
                        Me.ChkDepartment.Enabled = False
                        Me.ChkEmployee.Enabled = False
                        Me.chkJobPosition.Enabled = False
                        Me.chkSalaryFormula.Enabled = False
                        Me.Lbl_Dept_Checked.Enabled = False
                        Me.Lbl_Dept_Unchecked.Enabled = False
                        Me.Lbl_JP_checked.Enabled = False
                        Me.Lbl_JP_Unchecked.Enabled = False
                        Me.Lbl_SF_Checked.Enabled = False
                        Me.Lbl_SF_Unchecked.Enabled = False
                        Me.ClbPosition.Enabled = False
                        Me.ClbDepartment.Enabled = False
                        Me.clbSalaryFormula.Enabled = False
                        Me.dgvEmployee.Enabled = False

                        Me.UpdateThreading()
                        Me.BackgroundWorker2.WorkerReportsProgress = True
                        Me.BackgroundWorker2.RunWorkerAsync()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLMsgBox.Dispose()
            ListString = Nothing
            FrmError.Dispose()
        End Try
    End Sub

    Private Sub UpdateThreading()
        Dim ij As New System.Globalization.DateTimeFormatInfo
        Dim jk As New System.Globalization.NumberFormatInfo
        Try
            My.Application.ChangeCulture("en-US")
            My.Application.ChangeUICulture("en-US")
            ij.LongDatePattern = "dddd,MMM dd yyyy"
            ij.ShortDatePattern = "MMM dd yyyy"
            jk.CurrencyDecimalSeparator = "."
            jk.CurrencyGroupSeparator = ","
            My.Application.Culture.NumberFormat = jk
            My.Application.Culture.DateTimeFormat = ij
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function BuatPayment() As Boolean
        Dim RetVal As Boolean = True
        Try
            Me.OnProgress = True
            'Me.RowKary = Me.TBLKary.Rows(Me.i)
            Me.FrmSalary.BtnStop.Visible = True
            Me.FrmSalary.lblProgress.Visible = True
            Me.FrmSalary.ProgressBar1.Visible = True
            Me.FrmSalary.ProgressBar1.Maximum = Me.TBLKary.Rows.Count
            Me.FrmSalary.ProgressBar1.Value = Me.i + 1
            Me.FrmSalary.lblProgress.Text = "Data " + (Me.i + 1).ToString + " of " + Me.TBLKary.Rows.Count.ToString
            Me.LblProgress.Text = "Data " + (Me.i + 1).ToString + " of " + Me.TBLKary.Rows.Count.ToString + " - " + Me.RowKary.Full_Name
            Me.LblProgress.Refresh()
            Me.ProgressBar1.Visible = True
            Me.ProgressBar1.Maximum = Me.TBLKary.Rows.Count
            Me.ProgressBar1.Value = Me.i + 1
            '----------------------
            Me.FrmSalary.Baru = True
            Me.FrmSalary.UserCode = Me.UserCode
            Me.FrmSalary.ShowQuick_Edit = True
            Me.FrmSalary.NewQuick_Data = True
            'Masukan data untuk keperluan Loan Payment
            Me.FrmSalary.mPrimaryCashIn = CDbl(Me.TxtPrimaryCashAcc.Text)
            Me.FrmSalary.mPrimaryGainAcc = CDbl(Me.TxtPrimaryGainAcc.Text)
            Me.FrmSalary.mPrimaryLossAcc = CDbl(Me.TxtPrimaryLoss.Text)
            Me.FrmSalary.BackColorAll = Me.BackColorAll
            Me.FrmSalary.FungsiBtnNew()
            'masukan data yang dibutuhkan 
            'If Me.mPrimarySalaryFormulaRef <> -1 Then
            '    Me.FrmSalary.mPrimarySalaryFormulaRef = Me.mPrimarySalaryFormulaRef
            '    Me.FrmSalary.CopyData(Me.TBLSL_Kary_Ref, Me.FrmSalary.TBLSL_Kary_Ref)
            'End If
            Me.FrmSalary.DtpDate.Value = Me.DtpDate.Value
            Me.FrmSalary.TxtDocumentNo.Text = Me.FrmSalary.Obj.GetNoBaru(Me.FrmSalary.DtpDate.Value, Me.FrmSalary.ObjCD)
            Me.FrmSalary.Txt_P_NoForm.Text = Me.FrmSalary.Obj.GetNoBaru_BuatNPWP(Me.FrmSalary.DtpDate.Value, Me.FrmSalary.ObjCD)
            Me.FrmSalary.TxtPrimaryEmployee.Text = Me.RowKary.Primary
            Me.FrmSalary.dtpPeriod.Value = Me.dtpPeriod.Value
            Me.FrmSalary.DtpTo.Value = Me.DtpTo.Value
            Me.FrmSalary.FungsiBtnFindEmployee()
            Me.FrmSalary.dtpPeriod.Value = Me.dtpPeriod.Value
            Me.FrmSalary.DtpTo.Value = Me.DtpTo.Value
            Me.FrmSalary.cboTaxCalculation.SelectedIndex = Me.cboTaxCalculation.SelectedIndex
            Me.FrmSalary.TxtPrimaryTaxExpAcc.Text = Me.TxtPrimaryTaxExpAcc.Text
            Me.FrmSalary.FungsiBtnFindTaxAcc()
            Me.FrmSalary.FungsiBtnHitung()
            'bedain directly save ato bukan
            If Me.chkDirectlySave.Checked = False Then
                Me.FrmSalary.DariBtnPause = True
                Me.FrmSalary.TemporaryDataJgnDiHapus = True
                Me.FrmSalary.ShowDialog()
            Else
                'Execution
                Me.FrmSalary.BtnSimpan_PerformClick()
            End If

            'Penambahan yang ketinggalan
            If Me.FrmSalary.DariBtnStop = True Then
                RetVal = False
            End If

            Me.OnProgress = False
            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.OnProgress = False
        End Try
    End Function

    Private Sub FilterData(ByVal TBLMsgBox As DseTransaksiSalary.TBLMessageBoxDataTable, _
    ByVal ListString As List(Of String), ByVal vShowMsg As Boolean)
        Dim PrimaryUser As Double = -1
        Try
            If Me.RboSelected.Checked = True Then
                If Me.ChkDepartment.Checked = True Then
                    Me.AllData = False
                ElseIf Me.ChkEmployee.Checked = True Then
                    Me.AllData = False
                ElseIf Me.chkSalaryFormula.Checked = True Then
                    Me.AllData = False
                ElseIf Me.chkJobPosition.Checked = True Then
                    Me.AllData = False
                End If
            End If
            'Fill Data Karyawan dan sesuaikan dengan kriteria Selection Data
            PrimaryUser = Me.ObjUser.GetPrimary(Me.UserCode)
            If Me.AllData = True Then
                Me.KelolaDataKaryawan(Me.TBLKary, PrimaryUser)
            Else
                Me.ListDataFilter.Clear()
                If Me.ChkDepartment.Checked = True Then
                    If Me.ListDataFilter.Contains(FrmMasterWorkShift_Allocate.EnumLaporan.Departemen) = False Then
                        Me.ListDataFilter.Add(FrmMasterWorkShift_Allocate.EnumLaporan.Departemen)
                    End If
                End If
                If Me.ChkEmployee.Checked = True Then
                    If Me.ListDataFilter.Contains(FrmMasterWorkShift_Allocate.EnumLaporan.Employee) = False Then
                        Me.ListDataFilter.Add(FrmMasterWorkShift_Allocate.EnumLaporan.Employee)
                    End If
                End If
                If Me.chkSalaryFormula.Checked = True Then
                    If Me.ListDataFilter.Contains(FrmMasterWorkShift_Allocate.EnumLaporan.SalaryFormula) = False Then
                        Me.ListDataFilter.Add(FrmMasterWorkShift_Allocate.EnumLaporan.SalaryFormula)
                    End If
                End If
                If Me.chkJobPosition.Checked = True Then
                    If Me.ListDataFilter.Contains(FrmMasterWorkShift_Allocate.EnumLaporan.JobPosition) = False Then
                        Me.ListDataFilter.Add(FrmMasterWorkShift_Allocate.EnumLaporan.JobPosition)
                    End If
                End If
                Me.KelolaDataKaryawan(PrimaryUser, Me.TBLKary, Me.ListDataFilter, Me.dgvEmployee, Me.TBL_JP, _
                Me.ClbPosition, Me.TBL_Dept, Me.ClbDepartment, Me.TBL_Formula, Me.clbSalaryFormula)
            End If
            'Periksa apakah data salary yang di mau input punya Attendance, jika tidak hapus.
            Me.RemoveDataEmployeeYgTdkAdaAttendance(TBLMsgBox, ListString, vShowMsg)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KelolaDataKaryawan(ByVal TBLKaryawan As DseTransaksiSalary_Quick.tblmaster_karyawanDataTable, ByVal vPrimaryUser As Double)
        Dim DTAKaryawan As New DseTransaksiSalary_QuickTableAdapters.tblmaster_karyawanTableAdapter
        Try
            DTAKaryawan.Connection.ConnectionString = Core_Module.GetConstring
            TBLKaryawan.Rows.Clear()
            DTAKaryawan.Fill(TBLKaryawan, vPrimaryUser)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAKaryawan.Dispose()
        End Try
    End Sub

    Private Sub KelolaDataKaryawan(ByVal vPrimaryUser As Double, ByVal TBLKaryawan As DseTransaksiSalary_Quick.tblmaster_karyawanDataTable, _
    ByVal ListDataFilter As List(Of Double), ByVal DGVEmployee As DataGridView, ByVal TBL_JP As DseTransaksiSalary_Quick.tblmaster_jobpositionDataTable, _
    ByVal Clb_JP As CheckedListBox, ByVal TBL_Dept As DseTransaksiSalary_Quick.tblmaster_departemenDataTable, _
    ByVal Clb_dept As CheckedListBox, ByVal TBL_Formula As DseTransaksiSalary_Quick.tblmaster_salary_formula1DataTable, _
    ByVal Clb_Formula As CheckedListBox)
        Dim DTAKaryawan As New DseTransaksiSalary_QuickTableAdapters.tblmaster_karyawanTableAdapter
        Dim RowKaryawan As DseTransaksiSalary_Quick.tblmaster_karyawanRow
        Dim DTAMasterSalary As New DseTransaksiSalary_QuickTableAdapters.tblmaster_salary_formulaTableAdapter
        Dim TBLMasterSalary As New DseTransaksiSalary_Quick.tblmaster_salary_formulaDataTable
        Dim RowMasterSalary As DseTransaksiSalary_Quick.tblmaster_salary_formulaRow
        Dim EnumTipeDataFilter As Double
        Dim ij As Integer = 0
        Dim Ketemu As Boolean = False
        Dim Primary As Double = -1
        Dim RV As Data.DataRowView
        Dim RowDGV As DataGridViewRow
        Dim Filter As String = ""
        Try
            DTAKaryawan.Connection.ConnectionString = Core_Module.GetConstring
            TBLKaryawan.Rows.Clear()
            DTAKaryawan.Fill(TBLKaryawan, vPrimaryUser)
            If ListDataFilter.Count > 0 Then
                For Each EnumTipeDataFilter In ListDataFilter
                    If EnumTipeDataFilter = Core_Module.FrmMasterWorkShift_Allocate.EnumLaporan.Departemen Then
                        ij = 0
                        While ij <= TBLKaryawan.Rows.Count - 1
                            RowKaryawan = TBLKaryawan.Rows(ij)
                            Ketemu = False
                            If Clb_dept.CheckedItems.Count > 0 Then
                                For Each RV In Clb_dept.CheckedItems
                                    Primary = RV.Item("Primary")
                                    If RowKaryawan.primary_departemen = Primary Then
                                        Ketemu = True
                                        Exit For
                                    End If
                                Next
                                If Ketemu = False Then
                                    TBLKaryawan.Rows.RemoveAt(ij)
                                Else
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
                                    Primary = RowDGV.Cells(Me.ColPrimaryCode.Index).Value
                                    If RowKaryawan.Primary = Primary Then
                                        Ketemu = True
                                        Exit For
                                    End If
                                Next
                            End If
                            If Ketemu = False Then
                                TBLKaryawan.Rows.RemoveAt(ij)
                            Else
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
                                    If RowKaryawan.primary_position = Primary Then
                                        Ketemu = True
                                        Exit For
                                    End If
                                Next
                                If Ketemu = False Then
                                    TBLKaryawan.Rows.RemoveAt(ij)
                                Else
                                    ij += 1
                                End If
                            Else
                                ij += 1
                            End If
                        End While
                    End If
                    If EnumTipeDataFilter = Core_Module.FrmMasterWorkShift_Allocate.EnumLaporan.SalaryFormula Then
                        'MsgBox("Salary Formula masih dipending", MsgBoxStyle.Information, "Informasi")
                        DTAMasterSalary.Fill(TBLMasterSalary)
                        ij = 0
                        While ij <= TBLKaryawan.Rows.Count - 1
                            RowKaryawan = TBLKaryawan.Rows(ij)
                            Ketemu = False
                            Filter = "[Primary Karyawan] = " + RowKaryawan.Primary.ToString + ""
                            TBLMasterSalary.DefaultView.RowFilter = Filter
                            If TBLMasterSalary.DefaultView.Count > 0 Then
                                RowMasterSalary = TBLMasterSalary.DefaultView(0).Row
                                If Clb_Formula.CheckedItems.Count > 0 Then
                                    For Each RV In Clb_Formula.CheckedItems
                                        Primary = RV.Item("primary")
                                        If RowMasterSalary.primary = Primary Then
                                            Ketemu = True
                                            Exit For
                                        End If
                                    Next
                                    If Ketemu = False Then
                                        TBLKaryawan.Rows.RemoveAt(ij)
                                    Else
                                        ij += 1
                                    End If
                                Else
                                    ij += 1
                                End If
                            Else
                                TBLKaryawan.Rows.RemoveAt(ij)
                            End If
                        End While
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAKaryawan.Dispose()
            DTAMasterSalary.Dispose()
            TBLMasterSalary.Dispose()
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.DtpDate.Enabled = True
            Me.TxtExpAcc.Enabled = True
            Me.BtnFindTaxAcc.Enabled = True
            Me.TxtExpAccDesc.Enabled = True
            Me.cboTaxCalculation.Enabled = True
            Me.dtpPeriod.Enabled = True
            Me.DtpTo.Enabled = True

            If Me.AllData = True Then
                Me.ClbDepartment.Enabled = True
                Me.ClbPosition.Enabled = True
                Me.clbSalaryFormula.Enabled = True
                Me.dgvEmployee.Enabled = True
            Else
                Me.ClbDepartment.Enabled = False
                Me.ClbPosition.Enabled = False
                Me.clbSalaryFormula.Enabled = False
                Me.dgvEmployee.Enabled = False
            End If
            Me.ChkDepartment.Enabled = True
            Me.ChkEmployee.Enabled = True
            Me.chkJobPosition.Enabled = True
            Me.chkSalaryFormula.Enabled = True

            If Me.ChkDepartment.Checked = False Then
                Me.Lbl_Dept_Checked.Enabled = False
                Me.Lbl_Dept_Unchecked.Enabled = False
            Else
                Me.Lbl_Dept_Checked.Enabled = True
                Me.Lbl_Dept_Unchecked.Enabled = True
            End If
            If Me.ChkEmployee.Checked = False Then
                Me.Lbl_JP_checked.Enabled = False
                Me.Lbl_JP_Unchecked.Enabled = False
            Else
                Me.Lbl_JP_checked.Enabled = True
                Me.Lbl_JP_Unchecked.Enabled = True
            End If
            If Me.chkJobPosition.Checked = False Then
                Me.Lbl_SF_Checked.Enabled = False
                Me.Lbl_SF_Unchecked.Enabled = False
            Else
                Me.Lbl_SF_Checked.Enabled = True
                Me.Lbl_SF_Unchecked.Enabled = True
            End If
            Me.gboItem.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.DtpDate.Enabled = True
            Me.TxtExpAcc.Enabled = True
            Me.BtnFindTaxAcc.Enabled = True
            Me.TxtExpAccDesc.Enabled = True
            Me.cboTaxCalculation.Enabled = True
            Me.dtpPeriod.Enabled = True
            Me.DtpTo.Enabled = True

            Me.ChkDepartment.Enabled = False
            Me.ChkEmployee.Enabled = False
            Me.chkJobPosition.Enabled = False
            Me.chkSalaryFormula.Enabled = False
            Me.ClbDepartment.Enabled = False
            Me.ClbPosition.Enabled = False
            Me.clbSalaryFormula.Enabled = False
            Me.dgvEmployee.Enabled = False

            Me.Lbl_Dept_Checked.Enabled = False
            Me.Lbl_Dept_Unchecked.Enabled = False
            Me.Lbl_JP_checked.Enabled = False
            Me.Lbl_JP_Unchecked.Enabled = False
            Me.Lbl_SF_Checked.Enabled = False
            Me.Lbl_SF_Unchecked.Enabled = False
            Me.gboItem.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindTaxAcc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindTaxAcc.Click
        Dim i As New AccModule.FrmMasterAkun_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtExpAcc.Text <> "" Then
                i.txtCari.Text = Me.TxtExpAcc.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.TxtPrimaryTaxExpAcc.Text = i.mPrimaryAkun
                Me.TxtExpAcc.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryTaxExpAcc.Text)
                Me.TxtExpAccDesc.Text = Me.ObjAkun.GetDescription(Me.TxtExpAcc.Text)
                Me.cboTaxCalculation.Focus()
            Else
                Me.TxtExpAcc.Clear()
                Me.TxtPrimaryTaxExpAcc.Text = "-1"
                Me.TxtExpAcc.Focus()
                Me.TxtExpAccDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtExpAcc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtExpAcc.KeyDown
        Dim objCode As New AccModule.ClsMasterAkun
        Dim i As New AccModule.FrmMasterAkun_Cari
        Try
            If e.KeyValue = 13 Then
                If Me.TxtExpAcc.Text <> "" Then
                    If objCode.ApakahSudahAda_Aktif_NonHeader(Me.TxtExpAcc.Text) = True Then
                        Me.TxtPrimaryTaxExpAcc.Text = objCode.GetPrimary(Me.TxtExpAcc.Text)
                        Me.TxtExpAcc.Text = objCode.GetAccountCode(Me.TxtPrimaryTaxExpAcc.Text)
                        Me.TxtExpAccDesc.Text = objCode.GetDescription(Me.TxtExpAcc.Text)
                        Me.cboTaxCalculation.Focus()
                    Else
                        Me.BtnFindTaxAcc.PerformClick()
                    End If
                Else
                    Me.BtnFindTaxAcc.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtExpAcc_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtExpAcc.Leave
        Dim i As New AccModule.FrmMasterAkun_Cari
        Dim objCode As New AccModule.ClsMasterAkun
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtExpAcc.Text <> "" Then
                If objCode.ApakahSudahAda_Aktif_NonHeader(Me.TxtExpAcc.Text) = True Then
                    Me.TxtPrimaryTaxExpAcc.Text = objCode.GetPrimary(Me.TxtExpAcc.Text)
                    Me.TxtExpAcc.Text = objCode.GetAccountCode(Me.TxtPrimaryTaxExpAcc.Text)
                    Me.TxtExpAccDesc.Text = objCode.GetDescription(Me.TxtExpAcc.Text)
                Else
                    If Me.TxtExpAcc.Text <> "" Then
                        i.txtCari.Text = Me.TxtExpAcc.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryTaxExpAcc.Text = i.mPrimaryAkun
                        Me.TxtExpAcc.Text = objCode.GetAccountCode(Me.TxtPrimaryTaxExpAcc.Text)
                        Me.TxtExpAccDesc.Text = objCode.GetDescription(Me.TxtExpAcc.Text)
                    Else
                        Me.TxtExpAcc.Clear()
                        Me.TxtPrimaryTaxExpAcc.Text = "-1"
                        Me.TxtExpAccDesc.Clear()
                    End If
                End If
            Else
                Me.TxtExpAcc.Clear()
                Me.TxtPrimaryTaxExpAcc.Text = "-1"
                Me.TxtExpAccDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub FrmTrans_Gaji_Quick2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim FrmPay As New FrmTransaksiLoan_Payment
        Try
            Control.CheckForIllegalCrossThreadCalls = False

            Me.ShowIcon = False
            Me.StopProses = False
            Me.OnProgress = False
            Me.chkDirectlySave.Checked = True
            Me.i = 0
            Me.TBLKary.Rows.Clear()
            Me.LblProgress.Text = "......................................................"
            Me.LblProgress.Refresh()
            Me.ProgressBar1.Visible = True
            Me.ProgressBar1.Maximum = Me.TBLKary.Rows.Count
            Me.ProgressBar1.Value = Me.i
            Me.btnCancel.Visible = False
            Me.btnCancel.Enabled = False
            Me.btnCancel.TabStop = False
            Me.BtnPause.Visible = False
            Me.BtnPause.Enabled = False
            Me.BtnPause.TabStop = False
            ''
            Me.dtpPeriod.Value = Me.PeriodFrom
            Me.DtpTo.Value = Me.PeriodTo
            Me.cboTaxCalculation.SelectedIndex = 0
            Me.dgvEmployee.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            ''
            Me.dgvEmployee.AllowUserToAddRows = False
            Me.IsiListBoxJobPosition(Me.ClbPosition, Me.TBL_JP)
            Me.IsiListBoxDepartemen(Me.ClbDepartment, Me.TBL_Dept)
            Me.IsiListBoxFormula(Me.clbSalaryFormula, Me.TBL_Formula)
            Me.KendalikanKontrol(EnumLaporan.AllData, False)
            Me.AllData = True
            Me.ObjPay.IsiTBLCashCode(Me.CboCashAccount, Me.PrimaryCurr_Default, Me.TBLCash, FrmPay)
            Me.CboCashAccount.SelectedIndex = 0
            Me.TxtAccDesc.Text = Me.TBLCash(Me.CboCashAccount.SelectedIndex).Description
            '------Setting Global Setting------- 
            If Me.PrimaryCashIn = "-1" Then
                Me.PrimaryCashIn = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, ClsTransaksiSetting.EnumNamaParameter.Default_LoanPayment_PrimaryCashInAccount)
            End If
            If Me.PrimaryGain = "-1" Then
                Me.PrimaryGain = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, ClsTransaksiSetting.EnumNamaParameter.Default_LoanPayment_PrimaryGainAccount)
            End If
            If Me.PrimaryLoss = "-1" Then
                Me.PrimaryLoss = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, ClsTransaksiSetting.EnumNamaParameter.Default_LoanPayment_PrimaryLossAccount)
            End If
            If String.IsNullOrEmpty(Me.PrimaryGain) = True Then
                Me.TxtPrimaryGainAcc.Text = -1
                Me.TxtGainAccount.Clear()
                'Me.TxtDescGain.Clear()
            Else
                Me.TxtPrimaryGainAcc.Text = PrimaryGain
                Me.TxtGainAccount.Text = Me.ObjAkun.GetAccountCode(PrimaryGain)
                'Me.TxtDescGain.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(PrimaryGain)
            End If
            If String.IsNullOrEmpty(PrimaryLoss) = True Then
                Me.TxtPrimaryLoss.Text = -1
                Me.TxtLossAccount.Clear()
                'Me.TxtDescLoss.Clear()
            Else
                Me.TxtPrimaryLoss.Text = PrimaryLoss
                Me.TxtLossAccount.Text = Me.ObjAkun.GetAccountCode(PrimaryLoss)
                'Me.TxtDescLoss.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(PrimaryLoss)
            End If
            If String.IsNullOrEmpty(PrimaryCashIn) = False Then
                Me.CboCashAccount.Text = Me.ObjAkun.GetAccountCode(PrimaryCashIn)
            End If
            '----------------------------------------
            Me.RboSelected.Checked = True
            'Me.SplitContainer4.Panel1Collapsed = True

            Me.Height = Me.LblProgress.Top + Me.LblProgress.Height + 5

            Me.RefText = Me.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmPay.Dispose()
        End Try
    End Sub

    Private Sub RboAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RboAll.CheckedChanged
        Try
            If Me.RboAll.Checked = True Then
                Me.AllData = True
                Me.NonAktifkanDetail()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RboSelected_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RboSelected.CheckedChanged
        Try
            If Me.RboSelected.Checked = True Then
                Me.AllData = False
                Me.AktifkanDetail()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Navigasi untuk CLB, dan dgv Employee
    Private Sub IsiListBoxFormula(ByVal Lbo As System.Windows.Forms.CheckedListBox, ByVal TBL As DseTransaksiSalary_Quick.tblmaster_salary_formula1DataTable)
        Dim DTA As New DseTransaksiSalary_QuickTableAdapters.tblmaster_salary_formula1TableAdapter
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

    Private Sub IsiListBoxJobPosition(ByVal Lbo As System.Windows.Forms.CheckedListBox, ByVal TBL As DseTransaksiSalary_Quick.tblmaster_jobpositionDataTable)
        Dim DTA As New DseTransaksiSalary_QuickTableAdapters.tblmaster_jobpositionTableAdapter
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

    Private Sub IsiListBoxDepartemen(ByVal Lbo As System.Windows.Forms.CheckedListBox, ByVal TBL As DseTransaksiSalary_Quick.tblmaster_departemenDataTable)
        Dim DTA As New DseTransaksiSalary_QuickTableAdapters.tblmaster_departemenTableAdapter
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

    Private Sub Lbl_Position_Checked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Lbl_JP_checked.Click
        Try
            For i As Integer = 0 To Me.ClbPosition.Items.Count - 1
                Me.ClbPosition.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub Lbl_Position_UnChecked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Lbl_JP_Unchecked.Click
        Try
            For i As Integer = 0 To Me.ClbPosition.Items.Count - 1
                Me.ClbPosition.SetItemCheckState(i, Windows.Forms.CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub Lbl_Departemen_Checked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Lbl_Dept_Checked.Click
        Try
            For i As Integer = 0 To Me.ClbDepartment.Items.Count - 1
                Me.ClbDepartment.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub Lbl_Departemen_UnChecked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Lbl_Dept_Unchecked.Click
        Try
            For i As Integer = 0 To Me.ClbDepartment.Items.Count - 1
                Me.ClbDepartment.SetItemCheckState(i, Windows.Forms.CheckState.Unchecked)
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

    Public Enum EnumLaporan As Integer
        AllData
        JobPosition
        Departemen
        SalaryFormula
        Employee
    End Enum

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
                    Me.Lbl_JP_checked.Enabled = False
                    Me.Lbl_JP_Unchecked.Enabled = False
                    Me.ClbPosition.Enabled = False
                    For i As Integer = 0 To Me.ClbPosition.Items.Count - 1
                        Me.ClbPosition.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
                    Next
                Else
                    Me.Lbl_JP_checked.Enabled = True
                    Me.Lbl_JP_Unchecked.Enabled = True
                    Me.ClbPosition.Enabled = True
                End If
            End If
            If vFilter = EnumLaporan.Departemen Then
                If vAktif = False Then
                    Me.Lbl_Dept_Checked.Enabled = False
                    Me.Lbl_Dept_Unchecked.Enabled = False
                    Me.ClbDepartment.Enabled = False
                    For i As Integer = 0 To Me.ClbDepartment.Items.Count - 1
                        Me.ClbDepartment.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
                    Next
                Else
                    Me.Lbl_Dept_Checked.Enabled = True
                    Me.Lbl_Dept_Unchecked.Enabled = True
                    Me.ClbDepartment.Enabled = True
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

    Private Sub ChkDepartemen_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkDepartment.CheckedChanged
        Try
            If Me.ChkDepartment.Checked = False Then
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

    Private Sub ChkPosition_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkJobPosition.CheckedChanged
        Try
            If Me.chkJobPosition.Checked = False Then
                Me.KendalikanKontrol(EnumLaporan.JobPosition, False)
            Else
                Me.KendalikanKontrol(EnumLaporan.JobPosition, True)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkFormula_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkSalaryFormula.CheckedChanged
        Try
            If Me.chkSalaryFormula.Checked = False Then
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
                        If Me.ObjProduct.ApakahSudahAda_ShowSalary(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value, Me.UserCode) = True Then
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
                            ij.ShowSalary = True
                            ij.USERCode = Me.UserCode
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
                Else
                    If IsNothing(Me.dgvEmployee.Item(Me.ColPrimaryCode.Index, e.RowIndex).Value) = False Then
                        If e.ColumnIndex = Me.ColCode.Index Then
                            Me.EraseRowIndex = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If
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
                'Me.TxtDescGain.Text = Me.ObjAkun.GetDescription(Me.TxtGainAccount.Text)
                Me.TxtLossAccount.Focus()
            Else
                Me.TxtGainAccount.Clear()
                Me.TxtPrimaryGainAcc.Text = "-1"
                Me.TxtGainAccount.Focus()
                'Me.TxtDescGain.Clear()
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
                        'Me.TxtDescGain.Text = objCode.GetDescription(Me.TxtGainAccount.Text)
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
                    'Me.TxtDescGain.Text = objCode.GetDescription(Me.TxtGainAccount.Text)
                Else
                    If Me.TxtGainAccount.Text <> "" Then
                        i.txtCari.Text = Me.TxtGainAccount.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryGainAcc.Text = i.mPrimaryAkun
                        Me.TxtGainAccount.Text = objCode.GetAccountCode(Me.TxtPrimaryGainAcc.Text)
                        'Me.TxtDescGain.Text = objCode.GetDescription(Me.TxtGainAccount.Text)
                    Else
                        Me.TxtGainAccount.Clear()
                        Me.TxtPrimaryGainAcc.Text = "-1"
                        'Me.TxtDescGain.Clear()
                    End If
                End If
            Else
                Me.TxtGainAccount.Clear()
                Me.TxtPrimaryGainAcc.Text = "-1"
                'Me.TxtDescGain.Clear()
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
                'Me.TxtDescLoss.Text = ObjCode.GetDescription_ByPrimaryAcc(Me.TxtPrimaryLoss.Text)
                If Me.chkJobPosition.Enabled = False Then
                    Me.btnOK.Focus()
                Else
                    Me.chkJobPosition.Focus()
                End If
            Else
                Me.TxtLossAccount.Clear()
                Me.TxtPrimaryLoss.Text = "-1"
                'Me.TxtDescLoss.Text = " "
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
                        'Me.TxtDescLoss.Text = objCode.GetDescription_ByPrimaryAcc(Me.TxtPrimaryLoss.Text)
                        If Me.chkJobPosition.Enabled = False Then
                            Me.btnOK.Focus()
                        Else
                            Me.chkJobPosition.Focus()
                        End If
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
                    'Me.TxtDescLoss.Text = objCode.GetDescription_ByPrimaryAcc(Me.TxtPrimaryLoss.Text)
                Else
                    If Me.TxtLossAccount.Text <> "" Then
                        i.txtCari.Text = Me.TxtLossAccount.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryLoss.Text = i.mPrimaryAkun
                        Me.TxtLossAccount.Text = objCode.GetAccountCode(Me.TxtPrimaryLoss.Text)
                        'Me.TxtDescLoss.Text = objCode.GetDescription_ByPrimaryAcc(Me.TxtPrimaryLoss.Text)
                    Else
                        Me.TxtLossAccount.Clear()
                        Me.TxtPrimaryLoss.Text = "-1"
                        'Me.TxtDescLoss.Text = " "
                    End If
                End If
            Else
                Me.TxtLossAccount.Clear()
                Me.TxtPrimaryLoss.Text = "-1"
                'Me.TxtDescLoss.Text = " "
            End If
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
                Me.TxtAccDesc.Text = row.Description
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExpAcc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExpAcc.GotFocus
        Try
            'If Me.TxtRemarks.Enabled = True Then
            If Me.TxtExpAcc.Enabled = True Then
                If Me.TxtExpAcc.ReadOnly = False Then
                    Me.TxtExpAcc.BackColor = Me.BackColorAll
                Else
                    Me.TxtExpAcc.BackColor = Me.TxtExpAccDesc.BackColor
                End If
            End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExpAcc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExpAcc.LostFocus
        Try
            'If Me.TxtRemarks.Enabled = True Then
            If Me.TxtExpAcc.Enabled = True Then
                If Me.TxtExpAcc.ReadOnly = False Then
                    Me.TxtExpAcc.BackColor = Me.TxtPrimaryTaxExpAcc.BackColor
                Else
                    Me.TxtExpAcc.BackColor = Me.TxtExpAccDesc.BackColor
                End If
            End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtGainAccount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtGainAccount.GotFocus
        Try
            'If Me.TxtRemarksItem.Enabled = True Then
            If Me.TxtGainAccount.Enabled = True Then
                If Me.TxtGainAccount.ReadOnly = False Then
                    TxtGainAccount.BackColor = Me.BackColorAll
                Else
                    TxtGainAccount.BackColor = Me.TxtExpAccDesc.BackColor
                End If
            End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub TxtGainAccount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtGainAccount.LostFocus
        Try
            'If Me.TxtRemarksItem.Enabled = True Then
            If Me.TxtGainAccount.Enabled = True Then
                If Me.TxtGainAccount.ReadOnly = False Then
                    TxtGainAccount.BackColor = Me.TxtPrimaryTaxExpAcc.BackColor
                Else
                    TxtGainAccount.BackColor = Me.TxtExpAccDesc.BackColor
                End If
            End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLossAccount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLossAccount.GotFocus
        Try
            'If Me.TxtRemarksItem.Enabled = True Then
            If Me.TxtLossAccount.Enabled = True Then
                If Me.TxtLossAccount.ReadOnly = False Then
                    TxtLossAccount.BackColor = Me.BackColorAll
                Else
                    TxtLossAccount.BackColor = Me.TxtExpAccDesc.BackColor
                End If
            End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLossAccount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLossAccount.LostFocus
        Try
            'If Me.TxtRemarksItem.Enabled = True Then
            If Me.TxtLossAccount.Enabled = True Then
                If Me.TxtLossAccount.ReadOnly = False Then
                    TxtLossAccount.BackColor = Me.TxtPrimaryTaxExpAcc.BackColor
                Else
                    TxtLossAccount.BackColor = Me.TxtExpAccDesc.BackColor
                End If
            End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnPause_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnPause.Click
        Try
            If MsgBox("Pause creating calculate salary?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.FrmSalary.DariBtnPause = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub
End Class