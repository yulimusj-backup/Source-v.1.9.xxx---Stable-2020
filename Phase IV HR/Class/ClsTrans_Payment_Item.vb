Option Explicit On
Option Strict Off

Imports System.Windows.Forms, System.Drawing
Public Class ClsTrans_Payment_Item

    Public Function IsiDGVPayment(ByVal vPrimary As Double, ByVal vDGVPayment As DataGridView, _
    ByVal vDGVSalary As DataGridView, ByVal vDef_Prec As Double, ByVal IndexSalary As Integer, ByVal i As FrmTransaksiPayment) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItem As New DseTransPaymentTableAdapters.tbltrans_payment_item_employee_d_IsiDGVTableAdapter
        Dim TBLItem As New DseTransPayment.tbltrans_payment_item_employee_d_IsiDGVDataTable
        Dim RowItem As DseTransPayment.tbltrans_payment_item_employee_d_IsiDGVRow
        Dim RowTBL As DseTransPayment.tbltrans_payment_item_employee_dRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim Filter As String = ""
        Dim RowView As Data.DataRowView
        Try
            'update connection string
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring

            'autosize dgv
            vDGVPayment.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVPayment.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVPayment.AllowUserToAddRows = False
            vDGVSalary.AllowUserToAddRows = False

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan data item overhead cost
            DTAItem.FillBy_PrimaryTransaksi(TBLItem, vPrimary)
            'Salin data ke dalam mTBLTrans_payment_Employee_d
            i.mTBLItem_Employee_d.Rows.Clear()
            If TBLItem.Rows.Count > 0 Then
                For Each RowItem In TBLItem.Rows
                    RowTBL = i.mTBLItem_Employee_d.Newtbltrans_payment_item_employee_dRow
                    RowTBL.primary = RowItem.primary
                    RowTBL.primary_transaksi = RowItem.primary_transaksi
                    RowTBL.primary_akun_payment = RowItem.primary_akun_payment
                    RowTBL.identifier = RowItem.identifier
                    RowTBL.Index_Row_Salary = RowItem.Index_Row_Salary
                    RowTBL.amount = RowItem.amount
                    RowTBL.remarks = RowItem.remarks
                    RowTBL.akun_payment_code = RowItem.Account_Code
                    RowTBL.akun_payment_desc = RowItem.Description
                    i.mTBLItem_Employee_d.Rows.Add(RowTBL)
                Next
            End If
            Filter = "[identifier] = '" + vDGVSalary.Item(i.ColIdentifier.Index, IndexSalary).Value.ToString + "'"
            TBLItem.DefaultView.RowFilter = Filter
            vDGVPayment.Rows.Clear()
            If TBLItem.DefaultView.Count > 0 Then
                For Each RowView In TBLItem.DefaultView
                    RowItem = RowView.Row
                    vDGVPayment.Rows.Add()
                    rowDGV = vDGVPayment.Rows(vDGVPayment.Rows.Count - 1)
                    With rowDGV
                        .Cells(i.ColDetailPayment_AccountCode.Index).Value = RowItem.Account_Code
                        .Cells(i.ColDetailPayment_AccountDesc.Index).Value = RowItem.Description
                        .Cells(i.ColDetailPayment_Amount.Index).Value = ObjCD.DigitGrouping(RowItem.amount, vDef_Prec)
                        .Cells(i.ColDetailPayment_Identifier.Index).Value = RowItem.identifier
                        .Cells(i.ColDetailPayment_Remarks.Index).Value = RowItem.remarks
                        .Cells(i.ColDetailPayment_PrimaryAkun.Index).Value = RowItem.primary_akun_payment
                        .Cells(i.ColDetailPayment_IndexRowSalary.Index).Value = RowItem.Index_Row_Salary
                    End With
                Next
            End If

            'Masukin Nilai Textbox
            rowDGV = i.dgvItem.Rows(IndexSalary)
            i.TxtIncome.Text = rowDGV.Cells(i.ColEmp_Income.Index).Value
            i.TxtDeduction.Text = rowDGV.Cells(i.ColEmp_Deduction.Index).Value
            i.TxtIncome.Text = rowDGV.Cells(i.ColEmp_Income.Index).Value
            i.TxtGrossSalary.Text = rowDGV.Cells(i.ColEmp_GrossSalary.Index).Value
            i.TxtIncomeTax.Text = rowDGV.Cells(i.ColEmp_IncomeTax.Index).Value
            i.TxtLoanInstallment.Text = rowDGV.Cells(i.ColEmp_LoanInstallment.Index).Value
            i.TxtOtherDedution.Text = rowDGV.Cells(i.ColEmp_OtherDeduction.Index).Value
            i.TxtSalaryReceived.Text = rowDGV.Cells(i.ColEmp_TakeHomePay.Index).Value

            'autosize dgv
            vDGVPayment.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVPayment.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            vDGVPayment.AllowUserToAddRows = True
            vDGVPayment.AllowUserToAddRows = True

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function HapusItemPayment(ByVal vPrimary As Double) As Boolean
        Dim DTAItem As New DseTransPaymentTableAdapters.tbltrans_payment_item_employee_dTableAdapter
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Delete(vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
        End Try
    End Function

    Public Function SimpanItemPayment(ByVal vPrimary As Double, _
    ByVal mTBLPayment_Employee_d As DseTransPayment.tbltrans_payment_item_employee_dDataTable, ByVal i As FrmTransaksiPayment) As Boolean
        Dim DTAItem As New DseTransPaymentTableAdapters.tbltrans_payment_item_employee_dTableAdapter
        Dim RowItem As DseTransPayment.tbltrans_payment_item_employee_dRow
        Dim ObjCD As New Core_Module.ClsCekData
        'Parameter Identifier
        Dim IdentifierLama As String = ""
        Dim Identifier As String = ""
        Dim counter As Integer = 0
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If i.mTBLItem_Employee_d.Rows.Count > 0 Then
                For Each RowItem In mTBLPayment_Employee_d.Rows
                    'cek apakah data identifier perlu diubah
                    IdentifierLama = RowItem.identifier.ToString
                    If IdentifierLama.StartsWith(vPrimary.ToString) = False Then
                        'nilai start index = 1 agar tanda minus tidak diperhitungkan
                        counter = IdentifierLama.Substring(IdentifierLama.IndexOf("-", 1) + 1)
                        Identifier = vPrimary.ToString + "-" + counter.ToString
                        'ganti data identifier lama dengan yang baru, untuk item
                        RowItem.identifier = Identifier
                    Else
                        Identifier = IdentifierLama
                    End If
                    DTAItem.Insert(vPrimary, RowItem.identifier, RowItem.primary_akun_payment, RowItem.amount, _
                    RowItem.remarks, RowItem.Index_Row_Salary)
                Next
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function IsiDGVItemSalary(ByVal vPrimary As Double, ByVal vDGVSalary As DataGridView, _
    ByVal vDef_Prec As Double, ByVal i As FrmTransaksiPayment) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItem As New DseTransPaymentTableAdapters.tbltrans_payment_item_employee_IsiDGVTableAdapter
        Dim TBLItem As New DseTransPayment.tbltrans_payment_item_employee_IsiDGVDataTable
        Dim RowItem As DseTransPayment.tbltrans_payment_item_employee_IsiDGVRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim ObjSalary As New ClsTransaksiSalary
        Dim cboDGV As DataGridViewComboBoxCell
        Dim TipeEnumTax As ClsTransaksiSalary.EnumTax
        Try
            'update connection string
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring

            'autosize dgv
            vDGVSalary.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVSalary.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVSalary.AllowUserToAddRows = False

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan data item overhead cost
            DTAItem.FillBy_PrimaryTransaksi(TBLItem, vPrimary)
            vDGVSalary.Rows.Clear()
            For Each RowItem In TBLItem.Rows
                vDGVSalary.Rows.Add()
                rowDGV = vDGVSalary.Rows(vDGVSalary.Rows.Count - 1)
                With rowDGV
                    .Cells(i.ColEmp_No.Index).Value = ""
                    .Cells(i.ColEmp_SalaryDoc.Index).Value = RowItem.Salary_No
                    .Cells(i.ColEmp_SalaryDate.Index).Value = Format(RowItem.Salary_Date, "dd MMM yyyy")
                    .Cells(i.ColEmp_EmployeeCode.Index).Value = RowItem.Code
                    .Cells(i.ColEmp_Name.Index).Value = RowItem.Full_Name
                    .Cells(i.ColEmp_Income.Index).Value = ObjCD.DigitGrouping(RowItem.income, vDef_Prec)
                    .Cells(i.ColEmp_Deduction.Index).Value = ObjCD.DigitGrouping(RowItem.deduction, vDef_Prec)
                    .Cells(i.ColEmp_GrossSalary.Index).Value = ObjCD.DigitGrouping(RowItem.gross_salary, vDef_Prec)
                    .Cells(i.ColEmp_IncomeTax.Index).Value = ObjCD.DigitGrouping(RowItem.income_tax, vDef_Prec)
                    .Cells(i.ColEmp_LoanInstallment.Index).Value = ObjCD.DigitGrouping(RowItem.loan_installment, vDef_Prec)
                    .Cells(i.ColEmp_OtherDeduction.Index).Value = ObjCD.DigitGrouping(RowItem.other_deduction, vDef_Prec)
                    .Cells(i.ColEmp_TakeHomePay.Index).Value = ObjCD.DigitGrouping(RowItem.take_home_pay, vDef_Prec)
                    .Cells(i.ColEmp_CompanyExpense.Index).Value = ObjCD.DigitGrouping(RowItem.company_expense, vDef_Prec)
                    .Cells(i.ColEmp_EmployeeExpense.Index).Value = ObjCD.DigitGrouping(RowItem.employee_expense, vDef_Prec)

                    .Cells(i.ColEmp_PrimarySalary.Index).Value = RowItem.primary_transaksi_salary
                    .Cells(i.ColEmp_PrimaryEmployee.Index).Value = RowItem.primary_karyawan
                    .Cells(i.ColEmp_Remarks.Index).Value = RowItem.remarks
                    .Cells(i.ColIdentifier.Index).Value = RowItem.identifier
                    .Cells(i.ColEmp_SalaryDate.Index).Value = Format(RowItem.Salary_Date, "dd MMM yyyy")
                    cboDGV = .Cells(i.ColEmp_TipeTaxIncome.Index)
                    TipeEnumTax = ObjSalary.GetTaxTypeCalculation(RowItem.primary_transaksi_salary)
                    .Cells(i.ColEmp_TipeTaxIncome.Index).Value = cboDGV.Items.Item(TipeEnumTax)
                    .Cells(i.ColEmp_TipeTaxIncome_Enum.Index).Value = CInt(TipeEnumTax)
                End With
            Next

            'set number
            ObjCD.SetNumber(vDGVSalary)

            'autosize dgv
            vDGVSalary.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVSalary.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            vDGVSalary.AllowUserToAddRows = True

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
            ObjCD = Nothing
            ObjSetting = Nothing
            ObjSalary = Nothing
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimary As Double) As Boolean
        Dim DTAItem As New DseTransPaymentTableAdapters.tbltrans_payment_item_employeeTableAdapter
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Delete(vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
        End Try
    End Function

    Public Function SimpanItem(ByVal vPrimary As Double, _
    ByVal vDGVItem As DataGridView, ByVal i As FrmTransaksiPayment) As Boolean
        Dim DTAItem As New DseTransPaymentTableAdapters.tbltrans_payment_item_employeeTableAdapter
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        'Parameter Identifier
        Dim IdentifierLama As String = ""
        Dim Identifier As String = ""
        Dim counter As Integer = 0
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            vDGVItem.AllowUserToAddRows = False
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            For Each rowDGV In vDGVItem.Rows
                With rowDGV
                    'cek apakah data identifier perlu diubah
                    IdentifierLama = rowDGV.Cells(i.ColIdentifier.Index).Value.ToString
                    If IdentifierLama.StartsWith(vPrimary.ToString) = False Then
                        'nilai start index = 1 agar tanda minus tidak diperhitungkan
                        counter = IdentifierLama.Substring(IdentifierLama.IndexOf("-", 1) + 1)
                        Identifier = vPrimary.ToString + "-" + counter.ToString
                        'ganti data identifier lama dengan yang baru, untuk item
                        .Cells(i.ColIdentifier.Index).Value = Identifier
                    Else
                        Identifier = IdentifierLama
                    End If
                    DTAItem.Insert(.Cells(i.ColIdentifier.Index).Value, vPrimary, .Cells(i.ColEmp_PrimarySalary.Index).Value, _
                    .Cells(i.ColEmp_PrimaryEmployee.Index).Value, .Cells(i.ColEmp_Income.Index).Value, _
                    .Cells(i.ColEmp_Deduction.Index).Value, .Cells(i.ColEmp_GrossSalary.Index).Value, _
                    .Cells(i.ColEmp_IncomeTax.Index).Value, .Cells(i.ColEmp_LoanInstallment.Index).Value, _
                    .Cells(i.ColEmp_OtherDeduction.Index).Value, .Cells(i.ColEmp_TakeHomePay.Index).Value, _
                    .Cells(i.ColEmp_Remarks.Index).Value, .Cells(i.ColEmp_CompanyExpense.Index).Value, _
                    .Cells(i.ColEmp_EmployeeExpense.Index).Value)
                End With
            Next
            vDGVItem.AllowUserToAddRows = True

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function IsiDGVItemJurnal(ByVal vPrimary As Double, ByVal vDGVJurnal As DataGridView, _
    ByVal vDef_Prec As Double, ByVal i As FrmTransaksiPayment, ByVal vObjPayment As ClsTrans_Payment) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItem As New DseTransPaymentTableAdapters.tbltrans_payment_item_jurnal_IsiDGVTableAdapter
        Dim TBLItem As New DseTransPayment.tbltrans_payment_item_jurnal_IsiDGVDataTable
        Dim RowItem As DseTransPayment.tbltrans_payment_item_jurnal_IsiDGVRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjCurr As New Core_Module.clsMasterCurrency
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Try
            'update connection string
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring

            'autosize dgv
            vDGVJurnal.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVJurnal.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan data item overhead cost
            DTAItem.FillBy_PrimaryTransaksi(TBLItem, vPrimary)
            vDGVJurnal.Rows.Clear()
            For Each RowItem In TBLItem.Rows
                vDGVJurnal.Rows.Add()
                rowDGV = vDGVJurnal.Rows(vDGVJurnal.Rows.Count - 1)
                With rowDGV
                    .Cells(i.ColJurnal_TipeJurnal_str.Index).Value = vObjPayment.GetStringTipeJurnal(RowItem.tipe_item_jurnal)
                    .Cells(i.ColJurnal_TipeJurnal_dbl.Index).Value = RowItem.tipe_item_jurnal
                    .Cells(i.ColJurnal_Remarks.Index).Value = RowItem.Remarks
                    .Cells(i.ColJurnal_Account.Index).Value = RowItem.Account_Code
                    .Cells(i.ColJurnal_DescAccount.Index).Value = RowItem.Account_Description
                    .Cells(i.ColJurnal_Debet.Index).Value = ObjCD.DigitGrouping(RowItem.debet, vDef_Prec)
                    .Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(RowItem.debet_original, vDef_Prec)
                    .Cells(i.ColJurnal_Kredit.Index).Value = ObjCD.DigitGrouping(RowItem.credit, vDef_Prec)
                    .Cells(i.ColJurnal_Credit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(RowItem.credit_original, vDef_Prec)
                    .Cells(i.ColJurnal_ExchangeRate.Index).Value = RowItem.ratio
                    .Cells(i.ColJurnal_OriginalCurrency.Index).Value = ObjCurr.GetDecimal_DefaultCurrency()
                    .Cells(i.ColJurnal_Dept.Index).Value = RowItem.Dept_Desc
                    .Cells(i.ColJurnal_Job.Index).Value = RowItem.Job_Desc
                    .Cells(i.ColJurnal_PrimaryDept.Index).Value = RowItem.primary_dept
                    .Cells(i.ColJurnal_PrimaryJob.Index).Value = RowItem.primary_job
                    .Cells(i.ColJurnal_PrimaryAkun.Index).Value = RowItem.primary_akun
                    .Cells(i.ColJurnal_PrimaryCurrency.Index).Value = ObjCurr.GetPrimary(.Cells(i.ColJurnal_Credit_OriginalCurrency.Index).Value)
                End With
            Next

            'autosize dgv
            vDGVJurnal.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVJurnal.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
            ObjCD = Nothing
            ObjCurr = Nothing
            ObjSetting = Nothing
        End Try
    End Function

    Public Function HapusItemJurnal(ByVal vPrimary As Double) As Boolean
        Dim DTAItem As New DseTransPaymentTableAdapters.tbltrans_payment_item_jurnalTableAdapter
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Delete(vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
        End Try
    End Function

    Public Function SimpanItemJurnal(ByVal vPrimary As Double, _
    ByVal vDGVItem As DataGridView, ByVal i As FrmTransaksiPayment) As Boolean
        Dim DTAItem As New DseTransPaymentTableAdapters.tbltrans_payment_item_jurnalTableAdapter
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For Each rowDGV In vDGVItem.Rows
                With rowDGV
                    DTAItem.Insert(vPrimary, .Cells(i.ColJurnal_PrimaryAkun.Index).Value, .Cells(i.ColJurnal_Debet.Index).Value, _
                    .Cells(i.ColJurnal_Kredit.Index).Value, .Cells(i.ColJurnal_Remarks.Index).Value, .Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value, _
                    .Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value, .Cells(i.ColJurnal_ExchangeRate.Index).Value, _
                    .Cells(i.ColJurnal_PrimaryDept.Index).Value, .Cells(i.ColJurnal_PrimaryJob.Index).Value, _
                    .Cells(i.ColJurnal_TipeJurnal_dbl.Index).Value)
                End With
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function GetIdentifierUrutan(ByVal vDGVItem As DataGridView, _
    ByVal vColIdentifier As Integer, ByVal baru As Boolean, ByVal vPrimaryTransaksi As Double) As String
        Dim Row As DataGridViewRow
        Try
            'Dim i As String = "2-32"
            'Dim k As String = "-1-2"
            'Dim j As String = i.Substring(i.IndexOf("-") + 1)kalo edit
            'Dim l As String = i.Substring(i.IndexOf("-", 1) + 1)kalo new
            Dim Identifier As String = ""
            Dim UrutanMax As Double = 0
            If baru = True Then
                If vDGVItem.Rows.Count > 0 Then
                    For i As Integer = 0 To vDGVItem.Rows.Count - 1
                        Row = vDGVItem.Rows(i)
                        With Row
                            If i = 0 Then
                                UrutanMax = CDbl(.Cells(vColIdentifier).Value.ToString.Substring(CStr(.Cells(vColIdentifier).Value).IndexOf("-", 1) + 1))
                            Else
                                If CDbl(.Cells(vColIdentifier).Value.ToString.Substring(CStr(.Cells(vColIdentifier).Value).IndexOf("-", 1) + 1)) > UrutanMax Then
                                    UrutanMax = CDbl(.Cells(vColIdentifier).Value.ToString.Substring(CStr(.Cells(vColIdentifier).Value).IndexOf("-", 1) + 1))
                                End If
                            End If
                        End With
                    Next
                    Identifier = CStr(vPrimaryTransaksi.ToString + "-" + CStr(UrutanMax + 1))
                Else
                    Identifier = CStr(vPrimaryTransaksi.ToString + "-" + "1")
                End If
            Else
                If vDGVItem.Rows.Count > 0 Then
                    For i As Integer = 0 To vDGVItem.Rows.Count - 1
                        Row = vDGVItem.Rows(i)
                        With Row
                            If i = 0 Then
                                UrutanMax = CDbl(.Cells(vColIdentifier).Value.ToString.Substring(CStr(.Cells(vColIdentifier).Value).IndexOf("-") + 1))
                            Else
                                If CDbl(.Cells(vColIdentifier).Value.ToString.Substring(CStr(.Cells(vColIdentifier).Value).IndexOf("-") + 1)) > UrutanMax Then
                                    UrutanMax = CDbl(.Cells(vColIdentifier).Value.ToString.Substring(CStr(.Cells(vColIdentifier).Value).IndexOf("-") + 1))
                                End If
                            End If
                        End With
                    Next
                    Identifier = CStr(vPrimaryTransaksi.ToString + "-" + CStr(UrutanMax + 1))
                Else
                    Identifier = CStr(vPrimaryTransaksi.ToString + "-" + "1")
                End If
            End If

            Return Identifier
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetIdentifierUrutan_BuatDGVCellEndEdit(ByVal vDGVItem As DataGridView, _
    ByVal vColIdentifier As Integer, ByVal baru As Boolean, ByVal vPrimaryTransaksi As Double, _
    ByVal vIndexDGVEdit As Integer) As String
        Dim Row As DataGridViewRow
        Try
            'Dim i As String = "2-32"
            'Dim k As String = "-1-2"
            'Dim j As String = i.Substring(i.IndexOf("-") + 1)kalo edit
            'Dim l As String = i.Substring(i.IndexOf("-", 1) + 1)kalo new
            Dim Identifier As String = ""
            Dim UrutanMax As Double = 0
            If baru = True Then
                If vDGVItem.Rows.Count > 0 Then
                    vDGVItem.AllowUserToAddRows = False
                    For i As Integer = 0 To vDGVItem.Rows.Count - 1
                        Row = vDGVItem.Rows(i)
                        If Row.Index <> vIndexDGVEdit Then
                            With Row
                                If i = 0 Then
                                    UrutanMax = CDbl(.Cells(vColIdentifier).Value.ToString.Substring(CStr(.Cells(vColIdentifier).Value).IndexOf("-", 1) + 1))
                                Else
                                    If CDbl(.Cells(vColIdentifier).Value.ToString.Substring(CStr(.Cells(vColIdentifier).Value).IndexOf("-", 1) + 1)) > UrutanMax Then
                                        UrutanMax = CDbl(.Cells(vColIdentifier).Value.ToString.Substring(CStr(.Cells(vColIdentifier).Value).IndexOf("-", 1) + 1))
                                    End If
                                End If
                            End With
                        End If
                    Next
                    vDGVItem.AllowUserToAddRows = True
                    Identifier = CStr(vPrimaryTransaksi.ToString + "-" + CStr(UrutanMax + 1))
                Else
                    Identifier = CStr(vPrimaryTransaksi.ToString + "-" + "1")
                End If
            Else
                If vDGVItem.Rows.Count > 0 Then
                    vDGVItem.AllowUserToAddRows = False
                    For i As Integer = 0 To vDGVItem.Rows.Count - 1
                        Row = vDGVItem.Rows(i)
                        If Row.Index <> vIndexDGVEdit Then
                            With Row
                                If i = 0 Then
                                    UrutanMax = CDbl(.Cells(vColIdentifier).Value.ToString.Substring(CStr(.Cells(vColIdentifier).Value).IndexOf("-") + 1))
                                Else
                                    If CDbl(.Cells(vColIdentifier).Value.ToString.Substring(CStr(.Cells(vColIdentifier).Value).IndexOf("-") + 1)) > UrutanMax Then
                                        UrutanMax = CDbl(.Cells(vColIdentifier).Value.ToString.Substring(CStr(.Cells(vColIdentifier).Value).IndexOf("-") + 1))
                                    End If
                                End If
                            End With
                        End If
                    Next
                    vDGVItem.AllowUserToAddRows = True
                    Identifier = CStr(vPrimaryTransaksi.ToString + "-" + CStr(UrutanMax + 1))
                Else
                    Identifier = CStr(vPrimaryTransaksi.ToString + "-" + "1")
                End If
            End If

            Return Identifier
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function
End Class
