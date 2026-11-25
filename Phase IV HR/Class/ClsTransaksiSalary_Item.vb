Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Public Class ClsTransaksiSalary_Item

    Dim ObjCD As New Core_Module.ClsCekData
    Public Enum EnumType As Integer
        Income
        Deduction
    End Enum

    Public Enum EnumTaxElement As Integer
        No1_Gaji_Pensiunan_THT_JHT
        No2_TunjanganPPH
        No3_TunjLain_lembur_dsb
        No4_Honorarium_ImbalanLain_Sejenisnya
        No5_PremiAsuransi_DibayarkanPemberiKerja
        No6_Natura_KenikmatanLain_KenaPPH
        No7_Tantiem_Bonus_Gratifikasi_JasaProduksi
        No8_Jumlah_Penghasilan_Bruto
        No9_BiayaJabatan_BiayaPensiun
        No10_IuranPensiun_THT_JHT
        No11_JumlahPengurang
        No12_JumlahPenghasilanUntukPerhitunganPPh
        No13_PenghasilanNettoMasaSebelum '#sinkron2018_pphcls
        No14_PenghasilanNettoKenaPajak
        No15_PTKP
        No16_PKP
        No17_PPHSetahun
        No18_PPH21DipotongMasaPajakSebelum
        No19_PPHTerutang
        No20_PPHLunas_Terpotong
        NoTax
        No21_Pesangon
        No21_TunjPPHPesangon
        No21_PPHFinalPesangon
    End Enum

    Public Sub IsiDgvAttendance(ByVal vPrimaryKrywn As Double, ByVal vEndDate As Date, ByVal vStartDate As Date, ByVal vFrm As FrmTrans_Gaji_ShowAttendance)
        Dim Dta As New DseTransaksiSalaryTableAdapters.DataTable5TableAdapter
        Dim tbl As New DseTransaksiSalary.DataTable5DataTable
        Dim row As DseTransaksiSalary.DataTable5Row
        Dim RowDgv As System.Windows.Forms.DataGridViewRow
        Dim CboDgv As System.Windows.Forms.DataGridViewComboBoxCell
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta.Fill(tbl, vEndDate, vStartDate, vPrimaryKrywn)
            vFrm.dgvItem.Rows.Clear()
            For Each row In tbl.Rows
                vFrm.dgvItem.Rows.Add()
                RowDgv = vFrm.dgvItem.Rows(vFrm.dgvItem.Rows.Count - 1)
                RowDgv.Cells(vFrm.ColNo.Index).Value = vFrm.dgvItem.Rows.Count
                My.Application.ChangeCulture("id-ID")
                My.Application.ChangeUICulture("id-ID")
                RowDgv.Cells(vFrm.ColDate.Index).Value = Format(row._date, "dd MMM yyyy")
                My.Application.ChangeCulture("en-US")
                My.Application.ChangeUICulture("en-US")
                RowDgv.Cells(vFrm.ColWorkshift.Index).Value = row.Code_WS
                CboDgv = RowDgv.Cells(vFrm.ColWorkdays.Index)
                RowDgv.Cells(vFrm.ColWorkdays.Index).Value = CboDgv.Items.Item(row.Tipe_Workdays)
                CboDgv = RowDgv.Cells(vFrm.ColDuration.Index)
                RowDgv.Cells(vFrm.ColDuration.Index).Value = CboDgv.Items.Item(row.duration)
                CboDgv = RowDgv.Cells(vFrm.ColAttendanceStatus.Index)
                RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(row.presence_status)
                RowDgv.Cells(vFrm.ColSchIn.Index).Value = Me.ObjCD.Convert_ToString(row.sch_in)
                RowDgv.Cells(vFrm.ColSchBreakOut.Index).Value = Me.ObjCD.Convert_ToString(row.sch_break_out)
                RowDgv.Cells(vFrm.ColSchBreakIn.Index).Value = Me.ObjCD.Convert_ToString(row.sch_break_in)
                RowDgv.Cells(vFrm.ColSchOut.Index).Value = Me.ObjCD.Convert_ToString(row.sch_out)
                RowDgv.Cells(vFrm.ColActualIn.Index).Value = Me.ObjCD.Convert_ToString(row.actual_in)
                RowDgv.Cells(vFrm.ColActualBreakOut.Index).Value = Me.ObjCD.Convert_ToString(row.actual_break_out)
                RowDgv.Cells(vFrm.ColActualBreakIn.Index).Value = Me.ObjCD.Convert_ToString(row.actual_break_in)
                RowDgv.Cells(vFrm.ColActualOut.Index).Value = Me.ObjCD.Convert_ToString(row.actual_out)
                RowDgv.Cells(vFrm.ColStatusRemark.Index).Value = row.presence_status_remarks
                RowDgv.Cells(vFrm.ColLeaveCode.Index).Value = row.Code_Cuti
                RowDgv.Cells(vFrm.ColLeaveDesc.Index).Value = row.Desc_cuti
                RowDgv.Cells(vFrm.ColLeaveDoc.Index).Value = row.Cuti_app
                RowDgv.Cells(vFrm.ColEarlyInDuration.Index).Value = Me.ObjCD.Convert_ToString(row.early_in_duration)
                RowDgv.Cells(vFrm.ColLateDuration.Index).Value = Me.ObjCD.Convert_ToString(row.late_duration)
                RowDgv.Cells(vFrm.ColEarlyOutDuration.Index).Value = Me.ObjCD.Convert_ToString(row.early_out_duration)
                RowDgv.Cells(vFrm.ColEarlyInDuration.Index).Value = Me.ObjCD.Convert_ToString(row.early_in_duration)
                CboDgv = RowDgv.Cells(vFrm.ColOT.Index)
                RowDgv.Cells(vFrm.ColOT.Index).Value = CboDgv.Items.Item(row.tipe_overtime)
                RowDgv.Cells(vFrm.ColOTDuration.Index).Value = Me.ObjCD.Convert_ToString(row.ot_duration)
                RowDgv.Cells(vFrm.ColEarlyOTDuration.Index).Value = Me.ObjCD.Convert_ToString(row.early_ot_duration)
                RowDgv.Cells(vFrm.ColEarlyBreakOutDuration.Index).Value = Me.ObjCD.Convert_ToString(row.early_break_out_duration)
                RowDgv.Cells(vFrm.ColLateBreakInDuration.Index).Value = Me.ObjCD.Convert_ToString(row.late_break_in_duration)
                RowDgv.Cells(vFrm.ColRemarks.Index).Value = row.remarks
                RowDgv.Cells(vFrm.ColIsLate.Index).Value = row.is_late
                RowDgv.Cells(vFrm.ColIsEarlyIn.Index).Value = row.is_early_in
                RowDgv.Cells(vFrm.ColIsEarlyOut.Index).Value = row.is_early_out
                RowDgv.Cells(vFrm.ColIsEarlyBreakOut.Index).Value = row.is_early_break_out
                RowDgv.Cells(vFrm.ColIsLateBreakIn.Index).Value = row.is_late_break_in
                RowDgv.Cells(vFrm.ColSchWorkDuration.Index).Value = Me.ObjCD.Convert_ToString(row.sch_work_duration)
                RowDgv.Cells(vFrm.ColSchBreakDuration.Index).Value = Me.ObjCD.Convert_ToString(row.sch_break_duration)
                RowDgv.Cells(vFrm.ColActualWorkDuration.Index).Value = Me.ObjCD.Convert_ToString(row.actual_work_duration)
                RowDgv.Cells(vFrm.ColActualBreakDuration.Index).Value = Me.ObjCD.Convert_ToString(row.actual_break_duration)
                RowDgv.Cells(vFrm.ColIsPresent.Index).Value = row.is_present
                RowDgv.Cells(vFrm.ColIsSick.Index).Value = row.is_sick
                RowDgv.Cells(vFrm.ColIsLeave.Index).Value = row.is_leave
                RowDgv.Cells(vFrm.ColIsAlfa.Index).Value = row.Is_alfa
                RowDgv.Cells(vFrm.ColIsPermit.Index).Value = row.is_permit
                RowDgv.Cells(vFrm.ColIsOvertime.Index).Value = row.is_overtime
                RowDgv.Cells(vFrm.ColIsEarlyOvertime.Index).Value = row.is_early_overtime
                RowDgv.Cells(vFrm.ColIsHoliday.Index).Value = row.is_holiday
                RowDgv.Cells(vFrm.ColPrimaryWorkShift.Index).Value = row.primary_workshift
                RowDgv.Cells(vFrm.ColPrimaryLeave.Index).Value = row.primary_cuti
                RowDgv.Cells(vFrm.ColPermitDuration.Index).Value = Me.ObjCD.Convert_ToString(row.Permit_duration)
                'RowDgv.Cells(vFrm.ColPrimaryLA.Index).Value = row.Primary_Leave_App
            Next

            vFrm.dgvItem.Columns(vFrm.ColPermitDuration.Index).Visible = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
            tbl.Dispose()
        End Try
    End Sub

    Public Function IsiDGVOtherDeduction(ByVal vPrimary As Double, ByVal vDGVItem As DataGridView, _
    ByVal vDef_Prec As Double, ByVal i As FrmTrans_Gaji_OtherDeduction, _
    ByVal FrmSalary As FrmTrans_Gaji) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItem As New DseTransaksiSalaryTableAdapters.DataTable4TableAdapter
        Dim TBLItem As New DseTransaksiSalary.DataTable4DataTable
        Dim RowItem As DseTransaksiSalary.DataTable4Row
        Dim ObjCD As New Core_Module.ClsCekData
        Dim CboDGV As DataGridViewComboBoxCell
        Try
            'update connection string
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan data item overhead cost
            DTAItem.FillBy_PrimaryTransaksi(TBLItem, vPrimary)
            vDGVItem.Rows.Clear()
            For Each RowItem In TBLItem.Rows
                vDGVItem.Rows.Add()
                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                With rowDGV
                    .Cells(i.ColNo.Index).Value = ""
                    .Cells(i.ColSlipDesc.Index).Value = RowItem.SlipDescDescription
                    .Cells(i.ColComponentName.Index).Value = RowItem.Description
                    .Cells(i.ColTipeKomponen.Index).Value = "Other Deduction"
                    CboDGV = .Cells(i.ColTaxElement.Index)
                    .Cells(i.ColTaxElement.Index).Value = CboDGV.Items.Item(RowItem.Tax_Component)
                    .Cells(i.ColExpenseAccount.Index).Value = RowItem.Account_Code
                    .Cells(i.ColAmount.Index).Value = ObjCD.DigitGrouping(RowItem.Amount, vDef_Prec)
                    If RowItem.Override = 1 Then
                        .Cells(i.ColOverride.Index).Value = True
                    Else
                        .Cells(i.ColOverride.Index).Value = False
                    End If
                    .Cells(i.ColOverrideReason.Index).Value = RowItem.Override_Reason
                    .Cells(i.ColOriginalValue.Index).Value = ObjCD.DigitGrouping(RowItem.Original_Value, vDef_Prec)
                    .Cells(i.ColRemarks.Index).Value = RowItem.Remarks
                    .Cells(i.ColPrimarySlipDesc.Index).Value = RowItem.Primary_Slip_Desc
                    .Cells(i.ColIdentifierKomponen.Index).Value = RowItem.Primary
                    .Cells(i.ColPrimaryAkunExpense.Index).Value = RowItem.Primary_Expense_Account
                    .Cells(i.ColExplaination.Index).Value = RowItem.Explaination
                End With
            Next

            'set number
            If FrmSalary.ChkShowZeroAmount.Checked = True Then
                If vDGVItem.Rows.Count > 0 Then
                    For Each rowDGV In vDGVItem.Rows
                        vDGVItem.Rows(rowDGV.Index).Visible = True
                    Next
                End If
                FrmSalary.SetNumber(vDGVItem, True, True)
            Else
                If vDGVItem.Rows.Count > 0 Then
                    For Each rowDGV In vDGVItem.Rows
                        If CDec(rowDGV.Cells(i.ColAmount.Index).Value) = 0 Then
                            vDGVItem.Rows(rowDGV.Index).Visible = False
                        End If
                    Next
                End If
                FrmSalary.SetNumber(vDGVItem, True, False)
            End If

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function HapusItemOtherDeduction(ByVal vPrimary As Double) As Boolean
        Dim DTAItem As New DseTransaksiSalaryTableAdapters.tbltrans_salary_item_otherdeductionTableAdapter
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

    Public Function SimpanItemOtherDeduction(ByVal vPrimary As Double, _
    ByVal vDGVItem As DataGridView, ByVal i As FrmTrans_Gaji_OtherDeduction) As Boolean
        Dim DTAItem As New DseTransaksiSalaryTableAdapters.tbltrans_salary_item_otherdeductionTableAdapter
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim TipeKomponen As EnumType = EnumType.Deduction
        Dim Override As Integer = 0
        Dim CboDgv As DataGridViewComboBoxCell
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For Each rowDGV In vDGVItem.Rows
                If rowDGV.Cells(i.ColTipeKomponen.Index).Value = "Income" Then
                    TipeKomponen = EnumType.Income
                Else
                    TipeKomponen = EnumType.Deduction
                End If
                CboDgv = rowDGV.Cells(i.ColTaxElement.Index)
                If rowDGV.Cells(i.ColOverride.Index).Value = False Then
                    Override = 0
                Else
                    Override = 1
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColComponentName.Index).Value) = True Then
                    rowDGV.Cells(i.ColComponentName.Index).Value = ""
                End If
                With rowDGV
                    DTAItem.Insert(vPrimary, .Cells(i.ColPrimarySlipDesc.Index).Value, _
                    .Cells(i.ColComponentName.Index).Value, TipeKomponen, CboDgv.Items.IndexOf(.Cells(i.ColTaxElement.Index).Value), _
                    .Cells(i.ColPrimaryAkunExpense.Index).Value, .Cells(i.ColAmount.Index).Value, _
                    .Cells(i.ColRemarks.Index).Value, .Cells(i.ColIdentifierKomponen.Index).Value, _
                    .Cells(i.ColExplaination.Index).Value, Override, _
                    .Cells(i.ColOverrideReason.Index).Value, .Cells(i.ColOriginalValue.Index).Value)
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

    Public Function IsiDGV(ByVal vPrimary As Double, ByVal vDGVItem As DataGridView, _
    ByVal vDef_Prec As Double, ByVal i As FrmTrans_Gaji, ByVal vTBLItemIterasi As DseTransaksiSalary.tbltrans_salary_item_iterationDataTable, _
    ByVal vTBLItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxDataTable, _
    ByVal vTBLOtherPayment As DseTransaksiSalary.tbltrans_salary_item_other1DataTable, _
    ByVal vTBLPrevIncome As DseTransaksiSalary.tbltrans_salary_item_prev_income1DataTable) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItem As New DseTransaksiSalaryTableAdapters.DataTable2TableAdapter
        Dim TBLItem As New DseTransaksiSalary.DataTable2DataTable
        Dim RowItem As DseTransaksiSalary.DataTable2Row
        Dim ObjCD As New Core_Module.ClsCekData
        Dim CboDGV As DataGridViewComboBoxCell
        Dim DTAItemItetasi As New DseTransaksiSalaryTableAdapters.tbltrans_salary_item_iterationTableAdapter
        Dim DTAItemTax As New DseTransaksiSalaryTableAdapters.tbltrans_salary_itemtaxTableAdapter
        Dim DTAOtherPayment As New DseTransaksiSalaryTableAdapters.tbltrans_salary_item_other1TableAdapter
        Dim DTAPrevIncome As New DseTransaksiSalaryTableAdapters.tbltrans_salary_item_prev_income1TableAdapter

        Try
            If vDGVItem.ColumnCount = 0 Then
                Return False
            End If

            'update connection string
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemItetasi.Connection.ConnectionString = DTAItem.Connection.ConnectionString
            DTAItemTax.Connection.ConnectionString = DTAItem.Connection.ConnectionString
            DTAOtherPayment.Connection.ConnectionString = Core_Module.GetConstring
            DTAPrevIncome.Connection.ConnectionString = Core_Module.GetConstring

            DTAPrevIncome.Fill(vTBLPrevIncome, vPrimary)

            'isikan item pada data tabel
            DTAItemItetasi.FillBy_PrimaryTransaksi(vTBLItemIterasi, vPrimary)
            DTAItemTax.FillBy_PrimaryTransaksi(vTBLItemTax, vPrimary)
            vTBLOtherPayment.Rows.Clear()
            DTAOtherPayment.Fill(vTBLOtherPayment, vPrimary)

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan data item overhead cost
            DTAItem.FillBy_PrimaryTransaksi(TBLItem, vPrimary)
            vDGVItem.Rows.Clear()
            For Each RowItem In TBLItem.Rows
                vDGVItem.Rows.Add()
                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                CboDGV = rowDGV.Cells(i.ColTipeKomponen.Index)
                With rowDGV
                    .Cells(i.ColNo.Index).Value = ""
                    .Cells(i.ColSlipDesc.Index).Value = RowItem.SlipDescDescription
                    .Cells(i.ColComponentName.Index).Value = RowItem.Description
                    .Cells(i.ColTipeKomponen.Index).Value = CboDGV.Items.Item(RowItem.Type)

                    'If RowItem.Type = EnumType.Income Then
                    '    .Cells(i.ColTipeKomponen.Index).Value = "Income"
                    'Else
                    '    .Cells(i.ColTipeKomponen.Index).Value = "Deduction"
                    'End If
                    CboDGV = .Cells(i.ColTaxElement.Index)
                    .Cells(i.ColTaxElement.Index).Value = CboDGV.Items.Item(RowItem.Tax_Component)
                    .Cells(i.ColExpenseAccount.Index).Value = RowItem.Account_Code
                    .Cells(i.ColAmount.Index).Value = ObjCD.DigitGrouping(RowItem.Amount, vDef_Prec)
                    If RowItem.Override = 1 Then
                        .Cells(i.ColOverride.Index).Value = True
                    Else
                        .Cells(i.ColOverride.Index).Value = False
                    End If
                    .Cells(i.ColOverrideReason.Index).Value = RowItem.Override_Reason
                    .Cells(i.ColOriginalValue.Index).Value = ObjCD.DigitGrouping(RowItem.Original_Value, vDef_Prec)
                    .Cells(i.ColRemarks.Index).Value = RowItem.Remarks
                    .Cells(i.ColPrimarySlipDesc.Index).Value = RowItem.Primary_Slip_Desc
                    .Cells(i.ColIdentifierKomponen.Index).Value = RowItem.Identifier_Salary_Component
                    .Cells(i.ColPrimaryAkunExpense.Index).Value = RowItem.Primary_Expense_Account
                    .Cells(i.ColExplaination.Index).Value = RowItem.Explaination
                    .Cells(i.ColPrimaryAkunPayment.Index).Value = RowItem.Primary_Payment_Account
                    .Cells(i.ColIsRutin.Index).Value = RowItem.Is_Rutin
                End With
            Next

            'set number
            If i.ChkShowZeroAmount.Checked = True Then
                If vDGVItem.Rows.Count > 0 Then
                    For Each rowDGV In vDGVItem.Rows
                        vDGVItem.Rows(rowDGV.Index).Visible = True
                    Next
                End If
                i.SetNumber(vDGVItem, True, True)
            Else
                If vDGVItem.Rows.Count > 0 Then
                    For Each rowDGV In vDGVItem.Rows
                        If CDec(rowDGV.Cells(i.ColAmount.Index).Value) = 0 Then
                            vDGVItem.Rows(rowDGV.Index).Visible = False
                        End If
                    Next
                End If
                i.SetNumber(vDGVItem, True, False)
            End If

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
            ObjCD = Nothing
            DTAItemItetasi.Dispose()
            DTAItemTax.Dispose()
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimary As Double) As Boolean
        Dim DTAItem As New DseTransaksiSalaryTableAdapters.tbltrans_salary_itemTableAdapter
        Dim DTAItemTax As New DseTransaksiSalaryTableAdapters.tbltrans_salary_itemtaxTableAdapter
        Dim DTAItemIteration As New DseTransaksiSalaryTableAdapters.tbltrans_salary_item_iterationTableAdapter
        Dim DTAItemOtherPayment As New DseTransaksiSalaryTableAdapters.tbltrans_salary_item_otherTableAdapter

        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemTax.Connection.ConnectionString = DTAItem.Connection.ConnectionString
            DTAItemIteration.Connection.ConnectionString = DTAItem.Connection.ConnectionString
            DTAItemOtherPayment.Connection.ConnectionString = Core_Module.GetConstring

            DTAItem.Delete(vPrimary)
            DTAItemTax.Delete(vPrimary)
            DTAItemIteration.Delete(vPrimary)
            DTAItemOtherPayment.Delete(vPrimary)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            DTAItemIteration.Dispose()
            DTAItemTax.Dispose()
            DTAItemOtherPayment.Dispose()
        End Try
    End Function

    Public Function SimpanItem(ByVal vPrimary As Double, _
    ByVal vDGVItem As DataGridView, ByVal i As FrmTrans_Gaji, ByVal vTBLItemIterasi As DseTransaksiSalary.tbltrans_salary_item_iterationDataTable, _
    ByVal vTBLItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxDataTable, ByVal vTBLOtherPayment As DseTransaksiSalary.tbltrans_salary_item_other1DataTable, _
    ByVal vTBLPrevIncome As DseTransaksiSalary.tbltrans_salary_item_prev_income1DataTable) As Boolean
        'Dim DTAItem As New DseTransaksiSalaryTableAdapters.tbltrans_salary_itemTableAdapter
        'Dim rowDGV As DataGridViewRow
        'Dim ObjCD As New Core_Module.ClsCekData
        'Dim TipeKomponen As EnumType = EnumType.Deduction
        'Dim Override As Integer = 0
        'Dim CboDgv As DataGridViewComboBoxCell
        'Dim rowItemIterasi As DseTransaksiSalary.tbltrans_salary_item_iterationRow
        'Dim rowItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxRow
        'Dim DTAItemItetasi As New DseTransaksiSalaryTableAdapters.tbltrans_salary_item_iterationTableAdapter
        'Dim DTAItemTax As New DseTransaksiSalaryTableAdapters.tbltrans_salary_itemtaxTableAdapter
        'Dim DTAItemOtherPayment As New DseTransaksiSalaryTableAdapters.tbltrans_salary_item_otherTableAdapter
        'Dim rowItemOther As DseTransaksiSalary.tbltrans_salary_item_other1Row
        Dim DTAPrevIncome As New DseTransaksiSalaryTableAdapters.tbltrans_salary_item_prev_incomeTableAdapter
        Dim rowPrevIncome As DseTransaksiSalary.tbltrans_salary_item_prev_income1Row

        Try
            'DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            'DTAItemItetasi.Connection.ConnectionString = DTAItem.Connection.ConnectionString
            'DTAItemTax.Connection.ConnectionString = DTAItem.Connection.ConnectionString
            'DTAItemOtherPayment.Connection.ConnectionString = Core_Module.GetConstring
            DTAPrevIncome.Connection.ConnectionString = Core_Module.GetConstring

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTAPrevIncome.Delete(vPrimary)
            For Each rowPrevIncome In vTBLPrevIncome.Rows
                DTAPrevIncome.Insert(vPrimary, rowPrevIncome.primary)
            Next

            Me.SimpanItem_UsingScript(vPrimary, vDGVItem, i)
            'For Each rowDGV In vDGVItem.Rows
            '    CboDgv = rowDGV.Cells(i.ColTipeKomponen.Index)
            '    TipeKomponen = CboDgv.Items.IndexOf(rowDGV.Cells(i.ColTipeKomponen.Index).Value)

            '    CboDgv = rowDGV.Cells(i.ColTaxElement.Index)
            '    If rowDGV.Cells(i.ColOverride.Index).Value = False Then
            '        Override = 0
            '    Else
            '        Override = 1
            '    End If
            '    If String.IsNullOrEmpty(rowDGV.Cells(i.ColComponentName.Index).Value) = True Then
            '        rowDGV.Cells(i.ColComponentName.Index).Value = ""
            '    End If
            'With rowDGV
            '    DTAItem.Insert(vPrimary, .Cells(i.ColPrimarySlipDesc.Index).Value, _
            '    .Cells(i.ColComponentName.Index).Value, TipeKomponen, CboDgv.Items.IndexOf(.Cells(i.ColTaxElement.Index).Value), _
            '    .Cells(i.ColPrimaryAkunExpense.Index).Value, .Cells(i.ColAmount.Index).Value, _
            '    .Cells(i.ColRemarks.Index).Value, .Cells(i.ColIdentifierKomponen.Index).Value, _
            '    .Cells(i.ColExplaination.Index).Value, Override, _
            '    .Cells(i.ColOverrideReason.Index).Value, .Cells(i.ColOriginalValue.Index).Value, _
            '    .Cells(i.ColPrimaryAkunPayment.Index).Value, .Cells(i.ColIsRutin.Index).Value)
            'End With
            'Next

            ''simpan data iterasi
            'For Each rowItemIterasi In vTBLItemIterasi.Rows
            '    With rowItemIterasi
            '        DTAItemItetasi.Insert(vPrimary, .PPH_Terutang_Non_Rutin, ._Tunjangan_PPH_Non_Rutin, .PPH_Terutang_Rutin, .Tunjangan_PPH_Rutin)
            '    End With
            'Next

            'simpan data tax
            Me.SimpanItemTax_UsingScript(vPrimary, vTBLItemTax)
            'For Each rowItemTax In vTBLItemTax.Rows
            '    With rowItemTax
            '        DTAItemTax.Insert(vPrimary, .tipe_tax, .identifier_salary_component, .is_rutin, .amount_month, _
            '        .amount_gabungan_disetahunkan, .remarks_gabungan, .amount_rutin_disetahunkan, .remarks_rutin)
            '    End With
            'Next

            'simpan item payment other
            Me.SimpanItemOther_UsingScript(vPrimary, vTBLOtherPayment)
            'For Each rowItemOther In vTBLOtherPayment.Rows
            '    With rowItemOther
            '        DTAItemOtherPayment.Insert(vPrimary, .Description, .Type, .Tax_Component, .Amount, .Remarks, .Identifier_Salary_Component, .Is_Rutin, .Primary_Ref_Doc)
            '    End With
            'Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'DTAItem.Dispose()
            'ObjCD = Nothing
            'DTAItemItetasi.Dispose()
            'DTAItemTax.Dispose()
            'DTAItemOtherPayment.Dispose()
        End Try
    End Function

    Private Function SimpanItemOther_UsingScript(ByVal vPrimaryTransaksi As Double, _
    ByVal vTBLOtherPayment As DseTransaksiSalary.tbltrans_salary_item_other1DataTable) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 50
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBLOtherPayment.Rows.Count
                sql = Me.GetInsertString_ItemOther(StartRowIndex, BatchQty, vPrimaryTransaksi, vTBLOtherPayment)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemOther(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vPrimaryTransaksi As Double, _
    ByVal vTBLOtherPayment As DseTransaksiSalary.tbltrans_salary_item_other1DataTable) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim row As DseTransaksiSalary.tbltrans_salary_item_other1Row

        Try
            retval = "INSERT INTO `tbltrans_salary_item_other` (`Primary Transaksi`,`Description`,`Type`,`Tax Component`,`Amount`," + _
            "`Remarks`,`Identifier Salary Component`,`Is Rutin`,`Primary Ref Doc`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBLOtherPayment.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBLOtherPayment.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vTBLOtherPayment.Rows(vStartRowIndex + i)

                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + ",'" + row.Description + "'," _
                    + row.Type.ToString + "," _
                    + row.Tax_Component.ToString + "," _
                    + row.Amount.ToString + "," _
                    + "'" + row.Remarks + "'," _
                    + "'" + row.Identifier_Salary_Component + "'," _
                    + row.Is_Rutin.ToString + "," _
                    + CDec(row.Primary_Ref_Doc).ToString + "" _
                    + ")" + Delimiter
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function SimpanItem_UsingScript(ByVal vPrimaryTransaksi As Double, _
    ByVal vDGVItem As DataGridView, ByVal i As FrmTrans_Gaji) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 50
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vDGVItem.Rows.Count
                sql = Me.GetInsertStringItem(StartRowIndex, BatchQty, vDGVItem, vPrimaryTransaksi, i)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertStringItem(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vdgvItem As DataGridView, _
    ByVal vPrimaryTransaksi As Double, ByVal ijk As FrmTrans_Gaji) As String
        Dim retval As String = "" 
        Dim Delimiter As String = ","
        Dim TipeKomponen As EnumType = EnumType.Deduction
        Dim Override As Integer = 0
        Dim CboDgv As DataGridViewComboBoxCell
        Dim rowDGV As DataGridViewRow

        Try
            retval = "INSERT INTO `tbltrans_salary_item` (`Primary Transaksi`,`Primary Slip Desc`,`Description`,`Type`,`Tax Component`,`Primary Expense Account`," + _
            "`Amount`,`Remarks`,`Identifier Salary Component`,`Explaination`,`Override`,`Override Reason`,`Original Value`,`Primary Payment Account`,`Is Rutin`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vdgvItem.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vdgvItem.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    rowDGV = vdgvItem.Rows(vStartRowIndex + i)

                    CboDgv = rowDGV.Cells(ijk.ColTipeKomponen.Index)
                    TipeKomponen = CboDgv.Items.IndexOf(rowDGV.Cells(ijk.ColTipeKomponen.Index).Value)

                    CboDgv = rowDGV.Cells(ijk.ColTaxElement.Index)
                    If rowDGV.Cells(ijk.ColOverride.Index).Value = False Then
                        Override = 0
                    Else
                        Override = 1
                    End If
                    If String.IsNullOrEmpty(rowDGV.Cells(ijk.ColComponentName.Index).Value) = True Then
                        rowDGV.Cells(ijk.ColComponentName.Index).Value = ""
                    End If 

                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + "," + CDbl(rowDGV.Cells(ijk.ColPrimarySlipDesc.Index).Value).ToString + "," _
                    + "'" + rowDGV.Cells(ijk.ColComponentName.Index).Value.ToString + "'," _
                    + CInt(TipeKomponen).ToString + "," _
                    + CInt(CboDgv.Items.IndexOf(rowDGV.Cells(ijk.ColTaxElement.Index).Value)).ToString + "," _
                     + CDbl(rowDGV.Cells(ijk.ColPrimaryAkunExpense.Index).Value).ToString + "," _
                    + CDec(rowDGV.Cells(ijk.ColAmount.Index).Value).ToString + "," _
                    + "'" + rowDGV.Cells(ijk.ColRemarks.Index).Value.ToString + "'," _
                    + "'" + rowDGV.Cells(ijk.ColIdentifierKomponen.Index).Value.ToString + "'," _
                    + "'" + rowDGV.Cells(ijk.ColExplaination.Index).Value.ToString + "'," _
                    + CInt(Override).ToString + "," _
                    + "'" + rowDGV.Cells(ijk.ColOverrideReason.Index).Value.ToString + "'," _
                    + CDec(rowDGV.Cells(ijk.ColOriginalValue.Index).Value).ToString + "," _
                    + CDbl(rowDGV.Cells(ijk.ColPrimaryAkunPayment.Index).Value).ToString + "," _
                    + CInt(rowDGV.Cells(ijk.ColIsRutin.Index).Value).ToString + "" _
                    + ")" + Delimiter
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function SimpanItemTax_UsingScript(ByVal vPrimaryTransaksi As Double, _
    ByVal vTBLItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxDataTable) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 50
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBLItemTax.Rows.Count
                sql = Me.GetInsertString_ItemTax(StartRowIndex, BatchQty, vPrimaryTransaksi, vTBLItemTax)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemTax(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vPrimaryTransaksi As Double, _
    ByVal vTBLItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxDataTable) As String
        Dim retval As String = ""
        Dim Delimiter As String = "," 
        Dim row As DseTransaksiSalary.tbltrans_salary_itemtaxRow

        Try
            retval = "INSERT INTO `tbltrans_salary_itemtax` (`primary transaksi`,`tipe tax`,`identifier salary component`,`is rutin`," + _
            "`amount month`,`amount gabungan disetahunkan`,`remarks gabungan`,`amount rutin disetahunkan`,`remarks rutin`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBLItemTax.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBLItemTax.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vTBLItemTax.Rows(vStartRowIndex + i)

                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + "," + CInt(row.tipe_tax).ToString + "," _
                    + "'" + row.identifier_salary_component + "'," _
                    + CInt(row.is_rutin).ToString + "," _
                    + CDec(row.amount_month).ToString + "," _
                    + CDec(row.amount_gabungan_disetahunkan).ToString + "," _
                    + "'" + row.remarks_gabungan + "'," _
                    + "" + CDec(row.amount_rutin_disetahunkan).ToString + "," _
                    + "'" + row.remarks_rutin + "'" _
                    + ")" + Delimiter
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function SimpanItemPayment(ByVal vPrimary As Double, ByVal vDGVPayment As DataGridView, ByVal i As FrmTrans_Gaji_Payment) As Boolean
        Dim DTAItem As New DseTransaksiSalaryTableAdapters.tbltrans_salary_employee_d_igiTableAdapter
        Dim RowDGV As DataGridViewRow
        'Parameter Identifier
        Dim IdentifierLama As String = ""
        Dim Identifier As String = ""
        Dim counter As Integer = 0
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For Each RowDGV In vDGVPayment.Rows
                'cek apakah data identifier perlu diubah
                IdentifierLama = RowDGV.Cells(i.ColDetailPayment_Identifier.Index).Value.ToString
                If IdentifierLama.StartsWith(vPrimary.ToString) = False Then
                    'nilai start index = 1 agar tanda minus tidak diperhitungkan
                    counter = IdentifierLama.Substring(IdentifierLama.IndexOf("-", 1) + 1)
                    Identifier = vPrimary.ToString + "-" + counter.ToString
                    'ganti data identifier lama dengan yang baru, untuk item
                    RowDGV.Cells(i.ColDetailPayment_Identifier.Index).Value = Identifier
                Else
                    Identifier = IdentifierLama
                End If
                With RowDGV
                    If String.IsNullOrEmpty(RowDGV.Cells(i.ColDetailPayment_Remarks.Index).Value) = True Then
                        RowDGV.Cells(i.ColDetailPayment_Remarks.Index).Value = ""
                    End If
                    DTAItem.Insert(vPrimary, .Cells(i.ColDetailPayment_Identifier.Index).Value, _
                    .Cells(i.ColDetailPayment_PrimaryAkun.Index).Value, .Cells(i.ColDetailPayment_Amount.Index).Value, _
                    .Cells(i.ColDetailPayment_Remarks.Index).Value)
                End With
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
        End Try
    End Function

    Public Function HapusItemPayment(ByVal vPrimary As Double) As Boolean
        Dim DTAItem As New DseTransaksiSalaryTableAdapters.tbltrans_salary_employee_d_igiTableAdapter
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

    Public Function IsiDGVItemPayment(ByVal vPrimary As Double, ByVal vDGVSalary As DataGridView, _
    ByVal vDef_Prec As Double, ByVal i As FrmTrans_Gaji_Payment) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItem As New DseTransaksiSalaryTableAdapters.DataTable9TableAdapter
        Dim TBLItem As New DseTransaksiSalary.DataTable9DataTable
        Dim RowItem As DseTransaksiSalary.DataTable9Row
        Dim ObjCD As New Core_Module.ClsCekData
        Try
            'update connection string
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring

            'autosize dgv
            vDGVSalary.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVSalary.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan data item overhead cost
            DTAItem.FillBy_PrimarySalary(TBLItem, vPrimary)
            vDGVSalary.Rows.Clear()
            For Each RowItem In TBLItem.Rows
                vDGVSalary.Rows.Add()
                rowDGV = vDGVSalary.Rows(vDGVSalary.Rows.Count - 1)
                With rowDGV
                    If String.IsNullOrEmpty(RowItem.remarks) = True Then
                        .Cells(i.ColDetailPayment_Remarks.Index).Value = ""
                    Else
                        .Cells(i.ColDetailPayment_Remarks.Index).Value = RowItem.remarks
                    End If
                    .Cells(i.ColDetailPayment_AccountCode.Index).Value = RowItem.Account_Code
                    .Cells(i.ColDetailPayment_AccountDesc.Index).Value = RowItem.Description
                    .Cells(i.ColDetailPayment_Amount.Index).Value = ObjCD.DigitGrouping(RowItem.amount, vDef_Prec)
                    .Cells(i.ColDetailPayment_Identifier.Index).Value = RowItem.identifier
                    .Cells(i.ColDetailPayment_PrimaryAkun.Index).Value = RowItem.primary_akun_payment
                End With
            Next

            'set number
            ObjCD.SetNumber(vDGVSalary)

            'autosize dgv
            vDGVSalary.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVSalary.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function GetJmlItem(ByVal vIdentifierSalaryComponent As String, ByVal vSalaryDate As Date) As Double
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salary_itemTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            RetVal = DTA.ScalarQuery_GetJml(vIdentifierSalaryComponent, vSalaryDate)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
