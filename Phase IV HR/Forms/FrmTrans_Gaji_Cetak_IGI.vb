Option Explicit On
Option Strict Off

Imports Core_Module, System.Windows.Forms
Public Class FrmTrans_Gaji_Cetak_IGI
    Public mQuickSalary_Form As Boolean = False
    '--Single Setting--
    Public mPrimarySalary As Double = -1
    '--Quick Setting--
    Public mListPrimarySalary As New List(Of Double)
    Public mStartDate_Min As Date = Date.MaxValue
    Public mEndDate_Max As Date = Date.MinValue
    Public mTanggalTrans_Dari As Date = #1/1/1950#
    Public mTanggalTrans_Sampai As Date = #1/1/1950#
    Public mDgvsalaryQuick As DataGridView
    Public mIndexColoumn_StartDate As Integer = 0
    Public mIndexColoumn_EndDate As Integer = 0
    Public mIndexColoumn_PrimarySalary As Integer = 0
    Public mFormatCetakan As FrmTrans_Gaji_Cetak_IGI.EnumFormatCetakan = EnumFormatCetakan.A5

    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjParameter As New ClsMasterParameterGaji
    Dim ObjCurr As New clsMasterCurrency
    Dim ObjSalary As New ClsTransaksiSalary
    Dim mDefPrec As Double = Me.ObjCurr.GetDecimal_DefaultCurrency
    Dim ObjKary As New Core_Module.ClsMasterKaryawan

    '--Untuk Mendapatkan Nilai Jumlah Overtime---
    Dim TBLParam1 As New DseMasterParameter.DataTable1DataTable
    Dim TBLParam2 As New DseMasterParameter.tbltrans_attendanceDataTable
    Dim TBLWeekDay As New DseMasterParameter.tbltrans_attendance1DataTable
    Dim TBLHoliday As New DseMasterParameter.tbltrans_attendance1DataTable
    Dim TblKaryawan As New DseMasterParameter.tblmaster_karyawanDataTable

    Public Enum EnumFormatCetakan As Integer
        A5
        Continuous
    End Enum

    Public Enum EnumTipeOT As Integer
        NoOvertime
        WeekdayOvertime
        WeekendOvertime
    End Enum

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub FrmTransaksiJournal_Cetak1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Me.mQuickSalary_Form = False Then
                Me.CetakTransaksiSalary_IGI(Me.crvReport, Me.mPrimarySalary)
            Else
                'Quick Setting 
                Me.CetakTransaksiSalaryQuick_IGI_(Me.crvReport)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub CreatedPDF()
        Try
            Me.CetakTransaksiSalary_IGI(Me.crvReport, Me.mPrimarySalary, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetRounding(ByVal vTBL As DseLaporanSalaryStaff_IGI.Tbl_Overtime_Round_IGIDataTable, _
    ByVal vPrimarySalary As Double, ByVal vPrimarySlipDesc As Double, ByRef xRoundValue As Double, ByVal xRoundType As Double)
        Dim Row As DseLaporanSalaryStaff_IGI.Tbl_Overtime_Round_IGIRow
        Dim Filter As String = ""
        Try
            If vTBL.Rows.Count > 0 Then
                Filter = "[Primary Salary] = " + vPrimarySalary.ToString + " AND [Primary Slip Desc] = " + vPrimarySlipDesc.ToString + ""
                vTBL.DefaultView.RowFilter = Filter
                If vTBL.DefaultView.Count > 0 Then
                    Row = vTBL.DefaultView(0).Row
                    xRoundValue = Row.Rounding_Unit
                    xRoundType = Row.Rounding_Type
                End If
            End If
            vTBL.DefaultView.RowFilter = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub CetakTransaksiSalary_IGI(ByVal vCRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vPrimarySalary As Double, _
    Optional ByVal vCreateFile As Boolean = False)
        'Laporan
        Dim Report As New CrpCetakTransSalary_IGI
        Dim Report_Stngh As New CrpCetakTransSalary_IGI_Detail

        Dim DS As New DseTransaksiSalary_Cetakan_IGI
        Dim TBLCetak As DseTransaksiSalary_Cetakan_IGI.TblCetakanDataTable = DS.TblCetakan
        Dim rowcetak As DseTransaksiSalary_Cetakan_IGI.TblCetakanRow

        Dim DTA_Header As New DseTransaksiSalary_Cetakan_IGITableAdapters.DataHeaderTableAdapter
        Dim TBL_Header As New DseTransaksiSalary_Cetakan_IGI.DataHeaderDataTable
        Dim row_Header As DseTransaksiSalary_Cetakan_IGI.DataHeaderRow

        Dim DTA_item As New DseTransaksiSalary_Cetakan_IGITableAdapters.tbltrans_salary_item1TableAdapter
        Dim TBL_item As New DseTransaksiSalary_Cetakan_IGI.tbltrans_salary_item1DataTable
        Dim Row_item As DseTransaksiSalary_Cetakan_IGI.tbltrans_salary_item1Row

        Dim DTA_Transport As New DseTransaksiSalary_Cetakan_IGITableAdapters.TBLHitung_hari_TransportTableAdapter
        Dim TBL_Transport As New DseTransaksiSalary_Cetakan_IGI.TBLHitung_hari_TransportDataTable
        'Dim Row_Transport As DseTransaksiSalary_Cetakan_IGI.TBLHitung_hari_TransportRow

        Dim DTA_Overtime As New DseTransaksiSalary_Cetakan_IGITableAdapters.Tbl_OvertimeTableAdapter
        Dim TBL_Overtime As New DseTransaksiSalary_Cetakan_IGI.Tbl_OvertimeDataTable
        Dim Row_Overtime As DseTransaksiSalary_Cetakan_IGI.Tbl_OvertimeRow

        Dim DTA_Other As New DseTransaksiSalary_Cetakan_IGITableAdapters.tbltrans_salary_item_otherdeductionTableAdapter
        Dim TBL_Other As New DseTransaksiSalary_Cetakan_IGI.tbltrans_salary_item_otherdeductionDataTable
        Dim Row_Other As DseTransaksiSalary_Cetakan_IGI.tbltrans_salary_item_otherdeductionRow

        Dim DTA_Pay As New DseTransaksiSalary_Cetakan_IGITableAdapters.tbltrans_salary_employee_d_igiTableAdapter
        Dim TBL_Pay As New DseTransaksiSalary_Cetakan_IGI.tbltrans_salary_employee_d_igiDataTable
        Dim Row_Pay As DseTransaksiSalary_Cetakan_IGI.tbltrans_salary_employee_d_igiRow

        Dim DTARound As New DseLaporanSalaryStaff_IGITableAdapters.Tbl_Overtime_Round_IGITableAdapter
        Dim TBLRound As New DseLaporanSalaryStaff_IGI.Tbl_Overtime_Round_IGIDataTable

        Dim Jpk_Persen As Decimal = 0
        Dim Filter As String = ""
        Dim PrimaryKaryawan As Double = -1
        Dim Filter_Transport As String = ""

        Dim DV As Data.DataRowView
        Dim TotalOT_Weekday1 As Double = 0
        Dim TotalOT_Weekday2 As Double = 0
        Dim TotalOT_Weekend1 As Double = 0
        Dim TotalOT_Weekend2 As Double = 0
        Dim TotalOT_Weekend3 As Double = 0

        'Primary Cash, BCA, dan Mandiri
        Dim mPrimarySettingAkunBCA As Double = 156 '17 
        Dim mPrimarySettingAkunMandiri As Double = 12 '15

        'Parameter Rounding Value dan Type 
        Dim xRoundValue As Double = 1
        Dim xRoundType As Double = 1
        Try
            'Connection string
            DTA_Header.Connection.ConnectionString = Core_Module.GetConstring
            DTA_item.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Transport.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Overtime.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Other.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Pay.Connection.ConnectionString = Core_Module.GetConstring
            DTARound.Connection.ConnectionString = Core_Module.GetConstring

            'Fill Data Header dan Item berdasarkan Primary Salary
            DTA_Header.FillBy_PrimaryTransaksi(TBL_Header, vPrimarySalary)
            DTA_item.FillBy_PrimaryTransaksi(TBL_item, vPrimarySalary)
            DTA_Transport.Fill(TBL_Transport, Me.mTanggalTrans_Sampai, Me.mTanggalTrans_Dari)
            If TBL_Header.Rows.Count > 0 Then
                row_Header = TBL_Header.Rows(0)
                PrimaryKaryawan = row_Header.Primary_Employee
                Filter_Transport = "[primary karyawan] = " + PrimaryKaryawan.ToString + ""
                TBL_Transport.DefaultView.RowFilter = Filter_Transport
            End If
            DTA_Overtime.Fill(TBL_Overtime, Me.mTanggalTrans_Sampai, Me.mTanggalTrans_Dari)
            DTA_Other.FillBy_PrimaryTransaksiSalary(TBL_Other, vPrimarySalary)
            DTA_Pay.FillBy_PrimarySalary(TBL_Pay, vPrimarySalary)
            DTARound.FillBy_PrimarySalary(TBLRound, vPrimarySalary)

            'Buat baris di tabel cetak
            If TBL_Header.Rows.Count > 0 Then
                row_Header = TBL_Header.Rows(0)
                'mengambil data overtime (dari data header)
                Dim Filter_Ot As String = ""
                Filter_Ot = "[Tanggal Attendance] >= #" + Format(row_Header.Start_Date, "MM/dd/yyyy") + "# And [Tanggal Attendance] <= #" + Format(row_Header.End_Date, "MM/dd/yyyy") + "# And [primary karyawan] = " + row_Header.Primary_Employee.ToString + ""
                TBL_Overtime.DefaultView.RowFilter = Filter_Ot
                If TBL_Overtime.DefaultView.Count > 0 Then
                    For Each DV In TBL_Overtime.DefaultView
                        Row_Overtime = DV.Row
                        If Row_Overtime.tipe_overtime = EnumTipeOT.WeekdayOvertime Then
                            If Row_Overtime.OT_Duration >= 1.5 Then
                                'Konversi dalam menit
                                Row_Overtime.OT_Weekday_1 = 60
                                Row_Overtime.OT_Weekday_2 = Row_Overtime.OT_Duration * 60 - Row_Overtime.OT_Weekday_1 - 30
                                'Round slip desc = 20
                                Me.GetRounding(TBLRound, row_Header.Primary, 20, xRoundValue, xRoundType)
                                Row_Overtime.OT_Weekday_2 = Me.ObjCD.GetRoundingValue(Row_Overtime.OT_Weekday_2, 0, Core_Module.ClsCekData.EnumTipeRounding.Down, xRoundValue, Me.ObjCD) / 60
                                'Konversi dalam jam buat tampilan
                                Row_Overtime.OT_Weekday_1 = Row_Overtime.OT_Weekday_1 / 60
                                '--------------------
                            Else
                                If Row_Overtime.OT_Duration <> 0 Then
                                    'Round slip desc = 20
                                    Me.GetRounding(TBLRound, row_Header.Primary, 20, xRoundValue, xRoundType)
                                    Row_Overtime.OT_Duration = Row_Overtime.OT_Duration * 60
                                    Row_Overtime.OT_Weekday_1 = Me.ObjCD.GetRoundingValue(Row_Overtime.OT_Duration, 0, Core_Module.ClsCekData.EnumTipeRounding.Down, xRoundValue, Me.ObjCD) / 60
                                End If
                            End If
                        End If
                        If Row_Overtime.tipe_overtime = EnumTipeOT.WeekendOvertime Then
                            If Row_Overtime.Actual_Work_Duration > 9 Then
                                Row_Overtime.OT_Weekend_1 = 8
                                Row_Overtime.OT_Weekend_2 = 1
                                'Round slip desc = 23
                                Me.GetRounding(TBLRound, row_Header.Primary, 23, xRoundValue, xRoundType)
                                Row_Overtime.Actual_Work_Duration = Row_Overtime.Actual_Work_Duration * 60
                                Row_Overtime.Actual_Work_Duration = Me.ObjCD.GetRoundingValue(Row_Overtime.Actual_Work_Duration, 0, Core_Module.ClsCekData.EnumTipeRounding.Down, xRoundValue, Me.ObjCD)
                                '--------------------
                                Row_Overtime.OT_Weekend_3 = (Row_Overtime.Actual_Work_Duration / 60) - 9
                            ElseIf Row_Overtime.Actual_Work_Duration = 9 Then
                                Row_Overtime.OT_Weekend_1 = 8
                                Row_Overtime.OT_Weekend_2 = 1
                            ElseIf Row_Overtime.Actual_Work_Duration > 8 And Row_Overtime.Actual_Work_Duration < 9 Then
                                Row_Overtime.OT_Weekend_1 = 8
                                'Round slip desc = 23
                                Me.GetRounding(TBLRound, row_Header.Primary, 23, xRoundValue, xRoundType)
                                Row_Overtime.Actual_Work_Duration = Row_Overtime.Actual_Work_Duration * 60
                                Row_Overtime.Actual_Work_Duration = Me.ObjCD.GetRoundingValue(Row_Overtime.Actual_Work_Duration, 0, Core_Module.ClsCekData.EnumTipeRounding.Down, xRoundValue, Me.ObjCD)
                                '--------------------
                                Row_Overtime.OT_Weekend_2 = (Row_Overtime.Actual_Work_Duration / 60) - 8
                            ElseIf Row_Overtime.Actual_Work_Duration <= 8 Then
                                'Round slip desc = 21
                                Me.GetRounding(TBLRound, row_Header.Primary, 21, xRoundValue, xRoundType)
                                Row_Overtime.Actual_Work_Duration = Row_Overtime.Actual_Work_Duration * 60
                                Row_Overtime.Actual_Work_Duration = Me.ObjCD.GetRoundingValue(Row_Overtime.Actual_Work_Duration, 0, Core_Module.ClsCekData.EnumTipeRounding.Down, xRoundValue, Me.ObjCD)
                                '--------------------
                                Row_Overtime.OT_Weekend_1 = Row_Overtime.Actual_Work_Duration / 60
                            End If
                        End If
                        'If Row_Overtime.tipe_overtime = EnumTipeOT.WeekdayOvertime Then
                        '    If Row_Overtime.OT_Duration > 1 Then
                        '        Row_Overtime.OT_Weekday_1 = 1
                        '        Row_Overtime.OT_Weekday_2 = Row_Overtime.OT_Duration - Row_Overtime.OT_Weekday_1
                        '    Else
                        '        If Row_Overtime.OT_Duration <> 0 Then
                        '            Row_Overtime.OT_Weekday_1 = 1
                        '        End If
                        '    End If
                        'End If
                        'If Row_Overtime.tipe_overtime = EnumTipeOT.WeekendOvertime Then
                        '    If Row_Overtime.Actual_Work_Duration > 9 Then
                        '        Row_Overtime.OT_Weekend_1 = 8
                        '        Row_Overtime.OT_Weekend_2 = 1
                        '        Row_Overtime.OT_Weekend_3 = Row_Overtime.Actual_Work_Duration - 9
                        '    ElseIf Row_Overtime.Actual_Work_Duration = 9 Then
                        '        Row_Overtime.OT_Weekend_1 = 8
                        '        Row_Overtime.OT_Weekend_2 = 1
                        '    ElseIf Row_Overtime.Actual_Work_Duration <= 8 Then
                        '        Row_Overtime.OT_Weekend_1 = Row_Overtime.Actual_Work_Duration
                        '    End If
                        'End If
                        Row_Overtime.Primary_Salary = row_Header.Primary
                    Next
                End If
                '--------------
                rowcetak = TBLCetak.NewRow
                'Header Data
                If row_Header.Tax = ClsTransaksiSalary.EnumTax.Gross Then
                    rowcetak.Tax_Metode = "Gross"
                ElseIf row_Header.Tax = ClsTransaksiSalary.EnumTax.Gross_Up Then
                    rowcetak.Tax_Metode = "Gross Up"
                ElseIf row_Header.Tax = ClsTransaksiSalary.EnumTax.Netto Then
                    rowcetak.Tax_Metode = "Netto"
                End If
                rowcetak.Primary_Transaksi = row_Header.Primary
                rowcetak.Tanggal_Transaksi = row_Header.Tanggal_Transaksi
                rowcetak.Start_Date = row_Header.Start_Date
                rowcetak.End_Date = row_Header.End_Date
                rowcetak.Primary_Employee = row_Header.Primary_Employee
                rowcetak.Name = row_Header.Code
                rowcetak.Loan_Amount = row_Header.Loan_Installment
                'Item data
                Filter = "[Primary Slip Desc]=" + rowcetak.Basic_Primary_Slip.ToString + ""
                TBL_item.DefaultView.RowFilter = Filter
                If TBL_item.DefaultView.Count > 0 Then
                    Row_item = TBL_item.DefaultView(0).Row
                    rowcetak.Basic_Amount = Row_item.Amount
                End If

                Filter = "[Primary Slip Desc]=" + rowcetak.PlusMin_Primary_Slip.ToString + ""
                TBL_item.DefaultView.RowFilter = Filter
                If TBL_item.DefaultView.Count > 0 Then
                    Row_item = TBL_item.DefaultView(0).Row
                    rowcetak.PlusMin_Amount = Row_item.Amount * -1
                End If

                Filter = "[Primary Slip Desc]=" + rowcetak.Jamsostek_Primary_Slip.ToString + ""
                TBL_Other.DefaultView.RowFilter = Filter
                If TBL_Other.DefaultView.Count > 0 Then
                    Row_Other = TBL_Other.DefaultView(0).Row
                    rowcetak.Jamsostek_Amount = Row_Other.Amount * -1
                End If

                Filter = "[Primary Slip Desc]=" + rowcetak.Transport_Primary_Slip.ToString + ""
                TBL_item.DefaultView.RowFilter = Filter
                If TBL_item.DefaultView.Count > 0 Then
                    Row_item = TBL_item.DefaultView(0).Row
                    rowcetak.Transport_Amount = Row_item.Amount
                End If

                Filter = "[Primary Slip Desc]=" + rowcetak.OT_Weekday_1_Primary_Slip.ToString + ""
                TBL_item.DefaultView.RowFilter = Filter
                If TBL_item.DefaultView.Count > 0 Then
                    Row_item = TBL_item.DefaultView(0).Row
                    rowcetak.OT_Weekday_1_Amount = Row_item.Amount
                End If

                Filter = "[Primary Slip Desc]=" + rowcetak.OT_Weekday_2_Primary_Slip.ToString + ""
                TBL_item.DefaultView.RowFilter = Filter
                If TBL_item.DefaultView.Count > 0 Then
                    Row_item = TBL_item.DefaultView(0).Row
                    rowcetak.OT_Weekday_2_Amount = Row_item.Amount
                End If

                Filter = "[Primary Slip Desc]=" + rowcetak.OT_Weekend_1_Primary_Slip.ToString + ""
                TBL_item.DefaultView.RowFilter = Filter
                If TBL_item.DefaultView.Count > 0 Then
                    Row_item = TBL_item.DefaultView(0).Row
                    rowcetak.OT_Weekend_1_Amount = Row_item.Amount
                End If

                Filter = "[Primary Slip Desc]=" + rowcetak.OT_Weekend_2_Primary_Slip.ToString + ""
                TBL_item.DefaultView.RowFilter = Filter
                If TBL_item.DefaultView.Count > 0 Then
                    Row_item = TBL_item.DefaultView(0).Row
                    rowcetak.OT_Weekend_2_Amount = Row_item.Amount
                End If

                Filter = "[Primary Slip Desc]=" + rowcetak.OT_Weekend_3_Primary_Slip.ToString + ""
                TBL_item.DefaultView.RowFilter = Filter
                If TBL_item.DefaultView.Count > 0 Then
                    Row_item = TBL_item.DefaultView(0).Row
                    rowcetak.OT_Weekend_3_Amount = Row_item.Amount
                End If

                Filter = "[Primary Slip Desc]=" + rowcetak.Insentive_Primary_Slip.ToString + ""
                TBL_item.DefaultView.RowFilter = Filter
                If TBL_item.DefaultView.Count > 0 Then
                    Row_item = TBL_item.DefaultView(0).Row
                    rowcetak.Insentive_Amount = Row_item.Amount
                End If

                Filter = "[Primary Slip Desc]=" + rowcetak.Tunjangan_Primary_Slip.ToString + ""
                TBL_item.DefaultView.RowFilter = Filter
                If TBL_item.DefaultView.Count > 0 Then
                    Row_item = TBL_item.DefaultView(0).Row
                    rowcetak.Tunjangan_Amount = Row_item.Amount
                    If Me.ObjKary.GetMartialStatus(rowcetak.Primary_Employee) = ClsMasterKaryawan.EnumMartialStatus.Married Then
                        rowcetak.Tunjangan_Jml_Anak = Me.ObjParameter.GetJmlAnak(rowcetak.Primary_Employee)
                        If rowcetak.Tunjangan_Jml_Anak > 3 Then
                            rowcetak.Tunjangan_Jml_Anak = 3
                        End If
                    End If
                End If

                Filter = "[Primary Slip Desc]=" + rowcetak.Tunjangan_Lain_Primary_Slip.ToString + ""
                TBL_item.DefaultView.RowFilter = Filter
                If TBL_item.DefaultView.Count > 0 Then
                    Row_item = TBL_item.DefaultView(0).Row
                    rowcetak.Tunjangan_Lain_Amount = Row_item.Amount
                End If

                Filter = "[Primary Slip Desc]=" + rowcetak.Bonus_Primary_Slip.ToString + ""
                TBL_item.DefaultView.RowFilter = Filter
                If TBL_item.DefaultView.Count > 0 Then
                    Row_item = TBL_item.DefaultView(0).Row
                    rowcetak.Bonus_Amount = Row_item.Amount
                End If

                Filter = "[Primary Slip Desc]=" + rowcetak.Others_Primary_Slip.ToString + ""
                TBL_item.DefaultView.RowFilter = Filter
                If TBL_item.DefaultView.Count > 0 Then
                    Row_item = TBL_item.DefaultView(0).Row
                    rowcetak.Others_Amount = Row_item.Amount
                End If

                Filter = "[Primary Slip Desc]=" + rowcetak.JPK_Primary_Slip.ToString + ""
                TBL_item.DefaultView.RowFilter = Filter
                If TBL_item.DefaultView.Count > 0 Then
                    Row_item = TBL_item.DefaultView(0).Row
                    rowcetak.JPK_Amount = Row_item.Amount
                End If

                Filter = "[Primary Slip Desc]=" + rowcetak.Claim_Primary_Slip.ToString + ""
                TBL_item.DefaultView.RowFilter = Filter
                If TBL_item.DefaultView.Count > 0 Then
                    Row_item = TBL_item.DefaultView(0).Row
                    rowcetak.Claim_Amount = Row_item.Amount
                End If

                Filter = "[Primary Slip Desc]=" + rowcetak.Paid_Middle_Primary_Slip.ToString + ""
                TBL_Other.DefaultView.RowFilter = Filter
                If TBL_Other.DefaultView.Count > 0 Then
                    Row_Other = TBL_Other.DefaultView(0).Row
                    rowcetak.Paid_Middle_Amount = Row_Other.Amount
                End If

                Filter = "[Primary Slip Desc]=" + rowcetak.Koperasi_Primary_Slip.ToString + ""
                TBL_Other.DefaultView.RowFilter = Filter
                If TBL_Other.DefaultView.Count > 0 Then
                    Row_Other = TBL_Other.DefaultView(0).Row
                    rowcetak.Koperasi_Amount = Row_Other.Amount
                End If

                Filter = "[Primary Slip Desc]=" + rowcetak.Others_Deduction_Primary_Slip.ToString + ""
                TBL_Other.DefaultView.RowFilter = Filter
                If TBL_Other.DefaultView.Count > 0 Then
                    Row_Other = TBL_Other.DefaultView(0).Row
                    rowcetak.Others_Deduction_Amount = Row_Other.Amount
                End If

                rowcetak.Total_Salary = rowcetak.Basic_Amount + rowcetak.PlusMin_Amount + rowcetak.Jamsostek_Amount + rowcetak.Transport_Amount _
                                        + rowcetak.OT_Weekday_1_Amount + rowcetak.OT_Weekday_2_Amount + rowcetak.OT_Weekend_1_Amount + rowcetak.OT_Weekend_2_Amount + rowcetak.OT_Weekend_3_Amount

                rowcetak.Total_Income = rowcetak.Total_Salary + rowcetak.Insentive_Amount + rowcetak.Tunjangan_Amount + rowcetak.Tunjangan_Lain_Amount + rowcetak.Bonus_Amount + _
                                        rowcetak.Others_Amount + rowcetak.JPK_Amount + rowcetak.Claim_Amount

                rowcetak.Total_Nett_Income = rowcetak.Total_Income - rowcetak.Loan_Amount - rowcetak.Paid_Middle_Amount - rowcetak.Koperasi_Amount - rowcetak.Others_Deduction_Amount

                'Fill Table Hitung hari Transport
                If PrimaryKaryawan <> -1 Then
                    rowcetak.Ket_Transport_Faktor_Kali = TBL_Transport.DefaultView.Count
                Else
                    rowcetak.Ket_Transport_Faktor_Kali = 0
                End If
                If rowcetak.Ket_Transport_Faktor_Kali <> 0 Then
                    rowcetak.Ket_Transport_Value = rowcetak.Transport_Amount / rowcetak.Ket_Transport_Faktor_Kali
                Else
                    rowcetak.Ket_Transport_Value = 0
                End If

                If rowcetak.Basic_Amount = 0 Then
                    Jpk_Persen = 0
                Else
                    If rowcetak.Basic_Amount = 0 Then
                        Jpk_Persen = 0
                    Else
                        Jpk_Persen = rowcetak.JPK_Amount / rowcetak.Basic_Amount * 100
                    End If
                End If
                rowcetak.JPK_Persen = Math.Round(Jpk_Persen, 2, MidpointRounding.AwayFromZero).ToString

                'If Jpk_Persen >= 6 Then
                '    rowcetak.JPK_Persen = "6%"
                'Else
                '    rowcetak.JPK_Persen = "3%"
                'End If

                'Input Data Overtime
                If TBL_Overtime.DefaultView.Count > 0 Then
                    For Each DV In TBL_Overtime.DefaultView
                        Row_Overtime = DV.Row
                        TotalOT_Weekday1 += Row_Overtime.OT_Weekday_1
                        TotalOT_Weekday2 += Row_Overtime.OT_Weekday_2
                        TotalOT_Weekend1 += Row_Overtime.OT_Weekend_1
                        TotalOT_Weekend2 += Row_Overtime.OT_Weekend_2
                        TotalOT_Weekend3 += Row_Overtime.OT_Weekend_3
                    Next
                    rowcetak.OT_Weekday_1_Duration = TotalOT_Weekday1
                    rowcetak.OT_Weekday_2_Duration = TotalOT_Weekday2
                    rowcetak.OT_Weekend_1_Duration = TotalOT_Weekend1
                    rowcetak.OT_Weekend_2_Duration = TotalOT_Weekend2
                    rowcetak.OT_Weekend_3_Duration = TotalOT_Weekend3
                    'Hitung Value Overtime
                    If rowcetak.OT_Weekday_1_Duration <> 0 Then
                        rowcetak.OT_Weekday_1_Value = rowcetak.OT_Weekday_1_Amount / rowcetak.OT_Weekday_1_Duration
                    End If
                    If rowcetak.OT_Weekday_2_Duration <> 0 Then
                        rowcetak.OT_Weekday_2_Value = rowcetak.OT_Weekday_2_Amount / rowcetak.OT_Weekday_2_Duration
                    End If
                    If rowcetak.OT_Weekend_1_Duration <> 0 Then
                        rowcetak.OT_Weekend_Value = rowcetak.OT_Weekend_1_Amount / rowcetak.OT_Weekend_1_Duration
                    End If
                    If rowcetak.OT_Weekend_2_Duration <> 0 Then
                        rowcetak.OT_Weekend_2_Value = rowcetak.OT_Weekend_2_Amount / rowcetak.OT_Weekend_2_Duration
                    End If
                    If rowcetak.OT_Weekend_3_Duration <> 0 Then
                        rowcetak.OT_Weekend_3_Value = rowcetak.OT_Weekend_3_Amount / rowcetak.OT_Weekend_3_Duration
                    End If
                End If
                'Koding untuk Payment
                If TBL_Pay.Rows.Count > 0 Then
                    Dim HasilPaymentBCAMandiri As Decimal = 0
                    Dim HasilPaymentCash As Decimal = 0
                    Filter = "[primary akun payment] = " + mPrimarySettingAkunBCA.ToString + ""
                    TBL_Pay.DefaultView.RowFilter = Filter
                    If TBL_Pay.DefaultView.Count > 0 Then
                        Row_Pay = TBL_Pay.DefaultView(0).Row
                        rowcetak.Amount_Transfer_1 = Row_Pay.amount
                        rowcetak.Desc_Transfer_1 = "BCA"
                        HasilPaymentBCAMandiri += Row_Pay.amount
                    End If
                    Filter = "[primary akun payment] = " + mPrimarySettingAkunMandiri.ToString + ""
                    TBL_Pay.DefaultView.RowFilter = Filter
                    If TBL_Pay.DefaultView.Count > 0 Then
                        Row_Pay = TBL_Pay.DefaultView(0).Row
                        rowcetak.Amount_Transfer_2 = Row_Pay.amount
                        rowcetak.Desc_Transfer_2 = "Mandiri"
                        HasilPaymentBCAMandiri += Row_Pay.amount
                    End If
                    Filter = "[primary akun payment] <> " + mPrimarySettingAkunBCA.ToString + " AND [primary akun payment] <> " + mPrimarySettingAkunMandiri.ToString + ""
                    TBL_Pay.DefaultView.RowFilter = Filter
                    If TBL_Pay.DefaultView.Count > 0 Then
                        For Each DV In TBL_Pay.DefaultView
                            Row_Pay = DV.Row
                            HasilPaymentCash += Row_Pay.amount
                        Next
                        rowcetak.Amount_Cash = HasilPaymentCash
                    Else
                        rowcetak.Amount_Cash = 0
                    End If
                End If
                '-------------------------
                TBLCetak.Rows.Add(rowcetak)
            End If
            'Laporan
            If Me.mFormatCetakan = EnumFormatCetakan.A5 Then
                Report.SetDataSource(DS)
                Report.SetParameterValue(Report.Parameter_Format_Decimal.ParameterFieldName, Me.mDefPrec)
                vCRV.ReportSource = Report
            ElseIf Me.mFormatCetakan = EnumFormatCetakan.Continuous Then
                Report_Stngh.SetDataSource(DS)
                Report_Stngh.SetParameterValue(Report_Stngh.Parameter_Format_Decimal.ParameterFieldName, Me.mDefPrec)
                vCRV.ReportSource = Report_Stngh
            End If
            vCRV.Show()

            If vCreateFile = True Then
                Report.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, _
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\Salary Slip - " + _
                Me.ObjKary.GetFullNameByPrimary(Me.ObjSalary.GetPrimaryEmployeeByPrimary(vPrimarySalary)) + ".pdf")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Header.Dispose()
            DTA_item.Dispose()
            DTA_Overtime.Dispose()
            DTA_Transport.Dispose()
            DTA_Other.Dispose()
            TBL_Header.Dispose()
            TBL_item.Dispose()
            TBL_Overtime.Dispose()
            TBL_Transport.Dispose()
            TBL_Other.Dispose()
            DTA_Pay.Dispose()
            TBL_Pay.Dispose()
            DTARound.Dispose()
            TBLRound.Dispose()
        End Try
    End Sub

    Public Sub CetakTransaksiSalaryQuick_IGI_(ByVal vCRV As CrystalDecisions.Windows.Forms.CrystalReportViewer)
        Dim Report As New CrpCetakTransSalary_IGI
        Dim Report_Stngh As New CrpCetakTransSalary_IGI_Detail
        Dim DS As New DseTransaksiSalary_Cetakan_IGI
        Dim TBLCetak As DseTransaksiSalary_Cetakan_IGI.TblCetakanDataTable = DS.TblCetakan
        Dim rowcetak As DseTransaksiSalary_Cetakan_IGI.TblCetakanRow

        Dim DTA_Header As New DseTransaksiSalary_Cetakan_IGITableAdapters.DataHeaderTableAdapter
        Dim TBL_Header As New DseTransaksiSalary_Cetakan_IGI.DataHeaderDataTable
        Dim row_Header As DseTransaksiSalary_Cetakan_IGI.DataHeaderRow

        Dim DTA_item As New DseTransaksiSalary_Cetakan_IGITableAdapters.tbltrans_salary_item_QuickTableAdapter
        Dim TBL_item As New DseTransaksiSalary_Cetakan_IGI.tbltrans_salary_item_QuickDataTable
        Dim Row_item As DseTransaksiSalary_Cetakan_IGI.tbltrans_salary_item_QuickRow

        Dim DTA_Transport As New DseTransaksiSalary_Cetakan_IGITableAdapters.TBLHitung_hari_TransportTableAdapter
        Dim TBL_Transport As New DseTransaksiSalary_Cetakan_IGI.TBLHitung_hari_TransportDataTable
        'Dim Row_Transport As DseTransaksiSalary_Cetakan_IGI.TBLHitung_hari_TransportRow

        Dim DTA_Overtime As New DseTransaksiSalary_Cetakan_IGITableAdapters.Tbl_OvertimeTableAdapter
        Dim TBL_Overtime As New DseTransaksiSalary_Cetakan_IGI.Tbl_OvertimeDataTable
        Dim Row_Overtime As DseTransaksiSalary_Cetakan_IGI.Tbl_OvertimeRow

        Dim DTA_Other As New DseTransaksiSalary_Cetakan_IGITableAdapters.tbltrans_salary_item_otherdeduction_QuickTableAdapter
        Dim TBL_Other As New DseTransaksiSalary_Cetakan_IGI.tbltrans_salary_item_otherdeduction_QuickDataTable
        Dim Row_Other As DseTransaksiSalary_Cetakan_IGI.tbltrans_salary_item_otherdeduction_QuickRow

        Dim DTA_Pay As New DseTransaksiSalary_Cetakan_IGITableAdapters.tbltrans_salary_employee_d_igiTableAdapter
        Dim TBL_Pay As New DseTransaksiSalary_Cetakan_IGI.tbltrans_salary_employee_d_igiDataTable
        Dim Row_Pay As DseTransaksiSalary_Cetakan_IGI.tbltrans_salary_employee_d_igiRow

        Dim DTARound As New DseLaporanSalaryStaff_IGITableAdapters.Tbl_Overtime_Round_IGITableAdapter
        Dim TBLRound As New DseLaporanSalaryStaff_IGI.Tbl_Overtime_Round_IGIDataTable

        Dim Jpk_Persen As Decimal = 0
        Dim Filter As String = ""
        Dim PrimaryMinSalary As Double = Double.MaxValue
        'Dim RowDgv As DataGridViewRow
        Dim Filter_Transport As String = ""
        Dim PrimaryTransaksi As Double = -1

        Dim DV As Data.DataRowView
        Dim TotalOT_Weekday1 As Double = 0
        Dim TotalOT_Weekday2 As Double = 0
        Dim TotalOT_Weekend1 As Double = 0
        Dim TotalOT_Weekend2 As Double = 0
        Dim TotalOT_Weekend3 As Double = 0

        'Primary Cash, BCA, dan Mandiri
        Dim mPrimarySettingAkunBCA As Double = 156 '17
        Dim mPrimarySettingAkunMandiri As Double = 12 '15

        'Parameter Rounding Value dan Type 
        Dim xRoundValue As Double = 1
        Dim xRoundType As Double = 1
        Dim xGroupCetak As Double = 0
        Try
            'Connection string
            DTA_Header.Connection.ConnectionString = Core_Module.GetConstring
            DTA_item.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Transport.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Overtime.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Other.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Pay.Connection.ConnectionString = Core_Module.GetConstring
            DTARound.Connection.ConnectionString = Core_Module.GetConstring

            'Fill Data Header dan Item berdasarkan Primary Salary
            DTA_Header.FillBy_BuatQuick(TBL_Header, Me.mTanggalTrans_Sampai, Me.mTanggalTrans_Dari)
            'Untuk Pengecekan CHK Print
            If TBL_Header.Rows.Count > 0 Then
                Dim Counter As Integer = 0
                Dim Ketemu As Boolean = False
                While Counter <= TBL_Header.Rows.Count - 1
                    Ketemu = False
                    row_Header = TBL_Header.Rows(Counter)
                    For Each PrimaryTransaksi In Me.mListPrimarySalary
                        If PrimaryTransaksi = row_Header.Primary Then
                            Ketemu = True
                            Exit For
                        End If
                    Next
                    If Ketemu = True Then
                        Counter += 1
                    Else
                        TBL_Header.Rows.RemoveAt(Counter)
                    End If
                End While
                'Untuk mendapatkan Primary Salary Minimum
                For Each row_Header In TBL_Header.Rows
                    If PrimaryMinSalary > row_Header.Primary Then
                        PrimaryMinSalary = row_Header.Primary
                    End If
                Next
            End If
            DTA_item.FillBy_BuatQuick(TBL_item, PrimaryMinSalary)
            DTA_Other.FillBy_BuatQuick(TBL_Other, PrimaryMinSalary)
            DTA_Transport.Fill(TBL_Transport, Me.mEndDate_Max, Me.mStartDate_Min)
            DTA_Overtime.Fill(TBL_Overtime, Me.mEndDate_Max, Me.mStartDate_Min)
            DTA_Pay.FillBy_PrimarySalaryBuatQuick(TBL_Pay, PrimaryMinSalary)
            DTARound.FillBy_BuatQuickSalary(TBLRound, PrimaryMinSalary)

            'Buat baris di tabel cetak
            If TBL_Header.Rows.Count > 0 Then
                For ij As Integer = 0 To TBL_Header.Rows.Count - 1
                    row_Header = TBL_Header.Rows(ij)
                    'mengambil data overtime (dari data header)
                    Dim Filter_Ot As String = ""
                    Filter_Ot = "[Tanggal Attendance] >= #" + Format(row_Header.Start_Date, "MM/dd/yyyy") + "# And [Tanggal Attendance] <= #" + Format(row_Header.End_Date, "MM/dd/yyyy") + "# And [primary karyawan] = " + row_Header.Primary_Employee.ToString + ""
                    TBL_Overtime.DefaultView.RowFilter = Filter_Ot
                    If TBL_Overtime.DefaultView.Count > 0 Then
                        For Each DV In TBL_Overtime.DefaultView
                            Row_Overtime = DV.Row
                            If Row_Overtime.tipe_overtime = EnumTipeOT.WeekdayOvertime Then
                                If Row_Overtime.OT_Duration >= 1.5 Then
                                    'Konversi dalam menit
                                    Row_Overtime.OT_Weekday_1 = 60
                                    Row_Overtime.OT_Weekday_2 = Row_Overtime.OT_Duration * 60 - Row_Overtime.OT_Weekday_1 - 30
                                    'Round slip desc = 20
                                    Me.GetRounding(TBLRound, row_Header.Primary, 20, xRoundValue, xRoundType)
                                    Row_Overtime.OT_Weekday_2 = Me.ObjCD.GetRoundingValue(Row_Overtime.OT_Weekday_2, 0, Core_Module.ClsCekData.EnumTipeRounding.Down, xRoundValue, Me.ObjCD) / 60
                                    'Konversi dalam jam buat tampilan
                                    Row_Overtime.OT_Weekday_1 = Row_Overtime.OT_Weekday_1 / 60
                                    '--------------------
                                Else
                                    If Row_Overtime.OT_Duration <> 0 Then
                                        'Round slip desc = 20
                                        Me.GetRounding(TBLRound, row_Header.Primary, 20, xRoundValue, xRoundType)
                                        Row_Overtime.OT_Duration = Row_Overtime.OT_Duration * 60
                                        Row_Overtime.OT_Weekday_1 = Me.ObjCD.GetRoundingValue(Row_Overtime.OT_Duration, 0, Core_Module.ClsCekData.EnumTipeRounding.Down, xRoundValue, Me.ObjCD) / 60
                                    End If
                                End If
                            End If
                            If Row_Overtime.tipe_overtime = EnumTipeOT.WeekendOvertime Then
                                If Row_Overtime.Actual_Work_Duration > 9 Then
                                    Row_Overtime.OT_Weekend_1 = 8
                                    Row_Overtime.OT_Weekend_2 = 1
                                    'Round slip desc = 23
                                    Me.GetRounding(TBLRound, row_Header.Primary, 23, xRoundValue, xRoundType)
                                    Row_Overtime.Actual_Work_Duration = Row_Overtime.Actual_Work_Duration * 60
                                    Row_Overtime.Actual_Work_Duration = Me.ObjCD.GetRoundingValue(Row_Overtime.Actual_Work_Duration, 0, Core_Module.ClsCekData.EnumTipeRounding.Down, xRoundValue, Me.ObjCD)
                                    '--------------------
                                    Row_Overtime.OT_Weekend_3 = (Row_Overtime.Actual_Work_Duration / 60) - 9
                                ElseIf Row_Overtime.Actual_Work_Duration = 9 Then
                                    Row_Overtime.OT_Weekend_1 = 8
                                    Row_Overtime.OT_Weekend_2 = 1
                                ElseIf Row_Overtime.Actual_Work_Duration > 8 And Row_Overtime.Actual_Work_Duration < 9 Then
                                    Row_Overtime.OT_Weekend_1 = 8
                                    'Round slip desc = 23
                                    Me.GetRounding(TBLRound, row_Header.Primary, 23, xRoundValue, xRoundType)
                                    Row_Overtime.Actual_Work_Duration = Row_Overtime.Actual_Work_Duration * 60
                                    Row_Overtime.Actual_Work_Duration = Me.ObjCD.GetRoundingValue(Row_Overtime.Actual_Work_Duration, 0, Core_Module.ClsCekData.EnumTipeRounding.Down, xRoundValue, Me.ObjCD)
                                    '--------------------
                                    Row_Overtime.OT_Weekend_2 = (Row_Overtime.Actual_Work_Duration / 60) - 8
                                ElseIf Row_Overtime.Actual_Work_Duration <= 8 Then
                                    'Round slip desc = 21
                                    Me.GetRounding(TBLRound, row_Header.Primary, 21, xRoundValue, xRoundType)
                                    Row_Overtime.Actual_Work_Duration = Row_Overtime.Actual_Work_Duration * 60
                                    Row_Overtime.Actual_Work_Duration = Me.ObjCD.GetRoundingValue(Row_Overtime.Actual_Work_Duration, 0, Core_Module.ClsCekData.EnumTipeRounding.Down, xRoundValue, Me.ObjCD)
                                    '--------------------
                                    Row_Overtime.OT_Weekend_1 = Row_Overtime.Actual_Work_Duration / 60
                                End If
                            End If
                            'If Row_Overtime.tipe_overtime = EnumTipeOT.WeekdayOvertime Then
                            '    If Row_Overtime.OT_Duration > 1 Then
                            '        Row_Overtime.OT_Weekday_1 = 1
                            '        Row_Overtime.OT_Weekday_2 = Row_Overtime.OT_Duration - Row_Overtime.OT_Weekday_1
                            '    Else
                            '        If Row_Overtime.OT_Duration <> 0 Then
                            '            Row_Overtime.OT_Weekday_1 = Row_Overtime.OT_Duration
                            '        End If
                            '    End If
                            'End If
                            'If Row_Overtime.tipe_overtime = EnumTipeOT.WeekendOvertime Then
                            '    If Row_Overtime.Actual_Work_Duration > 9 Then
                            '        Row_Overtime.OT_Weekend_1 = 8
                            '        Row_Overtime.OT_Weekend_2 = 1
                            '        Row_Overtime.OT_Weekend_3 = Row_Overtime.Actual_Work_Duration - 9
                            '    ElseIf Row_Overtime.Actual_Work_Duration = 9 Then
                            '        Row_Overtime.OT_Weekend_1 = 8
                            '        Row_Overtime.OT_Weekend_2 = 1
                            '    ElseIf Row_Overtime.Actual_Work_Duration <= 8 Then
                            '        Row_Overtime.OT_Weekend_1 = Row_Overtime.Actual_Work_Duration
                            '    End If
                            'End If
                            Row_Overtime.Primary_Salary = row_Header.Primary
                        Next
                    End If
                    '--------------
                    rowcetak = TBLCetak.NewRow
                    'Header Data
                    If row_Header.Tax = ClsTransaksiSalary.EnumTax.Gross Then
                        rowcetak.Tax_Metode = "Gross"
                    ElseIf row_Header.Tax = ClsTransaksiSalary.EnumTax.Gross_Up Then
                        rowcetak.Tax_Metode = "Gross Up"
                    ElseIf row_Header.Tax = ClsTransaksiSalary.EnumTax.Netto Then
                        rowcetak.Tax_Metode = "Netto"
                    End If
                    rowcetak.Primary_Transaksi = row_Header.Primary
                    rowcetak.Tanggal_Transaksi = row_Header.Tanggal_Transaksi
                    rowcetak.Start_Date = row_Header.Start_Date
                    rowcetak.End_Date = row_Header.End_Date
                    rowcetak.Primary_Employee = row_Header.Primary_Employee
                    rowcetak.Name = row_Header.Code
                    rowcetak.Loan_Amount = row_Header.Loan_Installment
                    'Item data
                    Filter = "[Primary Slip Desc]=" + rowcetak.Basic_Primary_Slip.ToString + " And [Primary Transaksi] =" + row_Header.Primary.ToString + ""
                    TBL_item.DefaultView.RowFilter = Filter
                    If TBL_item.DefaultView.Count > 0 Then
                        Row_item = TBL_item.DefaultView(0).Row
                        rowcetak.Basic_Amount = Row_item.Amount
                    End If

                    Filter = "[Primary Slip Desc]=" + rowcetak.PlusMin_Primary_Slip.ToString + " And [Primary Transaksi] =" + row_Header.Primary.ToString + ""
                    TBL_item.DefaultView.RowFilter = Filter
                    If TBL_item.DefaultView.Count > 0 Then
                        Row_item = TBL_item.DefaultView(0).Row
                        rowcetak.PlusMin_Amount = Row_item.Amount * -1
                    End If

                    Filter = "[Primary Slip Desc]=" + rowcetak.Jamsostek_Primary_Slip.ToString + " And [Primary Transaksi] =" + row_Header.Primary.ToString + ""
                    TBL_Other.DefaultView.RowFilter = Filter
                    If TBL_Other.DefaultView.Count > 0 Then
                        Row_Other = TBL_Other.DefaultView(0).Row
                        rowcetak.Jamsostek_Amount = Row_Other.Amount * -1
                    End If

                    Filter = "[Primary Slip Desc]=" + rowcetak.Transport_Primary_Slip.ToString + " And [Primary Transaksi] =" + row_Header.Primary.ToString + ""
                    TBL_item.DefaultView.RowFilter = Filter
                    If TBL_item.DefaultView.Count > 0 Then
                        Row_item = TBL_item.DefaultView(0).Row
                        rowcetak.Transport_Amount = Row_item.Amount
                    End If

                    Filter = "[Primary Slip Desc]=" + rowcetak.OT_Weekday_1_Primary_Slip.ToString + " And [Primary Transaksi] =" + row_Header.Primary.ToString + ""
                    TBL_item.DefaultView.RowFilter = Filter
                    If TBL_item.DefaultView.Count > 0 Then
                        Row_item = TBL_item.DefaultView(0).Row
                        rowcetak.OT_Weekday_1_Amount = Row_item.Amount
                    End If

                    Filter = "[Primary Slip Desc]=" + rowcetak.OT_Weekday_2_Primary_Slip.ToString + " And [Primary Transaksi] =" + row_Header.Primary.ToString + ""
                    TBL_item.DefaultView.RowFilter = Filter
                    If TBL_item.DefaultView.Count > 0 Then
                        Row_item = TBL_item.DefaultView(0).Row
                        rowcetak.OT_Weekday_2_Amount = Row_item.Amount
                    End If

                    Filter = "[Primary Slip Desc]=" + rowcetak.OT_Weekend_1_Primary_Slip.ToString + " And [Primary Transaksi] =" + row_Header.Primary.ToString + ""
                    TBL_item.DefaultView.RowFilter = Filter
                    If TBL_item.DefaultView.Count > 0 Then
                        Row_item = TBL_item.DefaultView(0).Row
                        rowcetak.OT_Weekend_1_Amount = Row_item.Amount
                    End If

                    Filter = "[Primary Slip Desc]=" + rowcetak.OT_Weekend_2_Primary_Slip.ToString + " And [Primary Transaksi] =" + row_Header.Primary.ToString + ""
                    TBL_item.DefaultView.RowFilter = Filter
                    If TBL_item.DefaultView.Count > 0 Then
                        Row_item = TBL_item.DefaultView(0).Row
                        rowcetak.OT_Weekend_2_Amount = Row_item.Amount
                    End If

                    Filter = "[Primary Slip Desc]=" + rowcetak.OT_Weekend_3_Primary_Slip.ToString + " And [Primary Transaksi] =" + row_Header.Primary.ToString + ""
                    TBL_item.DefaultView.RowFilter = Filter
                    If TBL_item.DefaultView.Count > 0 Then
                        Row_item = TBL_item.DefaultView(0).Row
                        rowcetak.OT_Weekend_3_Amount = Row_item.Amount
                    End If

                    Filter = "[Primary Slip Desc]=" + rowcetak.Insentive_Primary_Slip.ToString + " And [Primary Transaksi] =" + row_Header.Primary.ToString + ""
                    TBL_item.DefaultView.RowFilter = Filter
                    If TBL_item.DefaultView.Count > 0 Then
                        Row_item = TBL_item.DefaultView(0).Row
                        rowcetak.Insentive_Amount = Row_item.Amount
                    End If

                    Filter = "[Primary Slip Desc]=" + rowcetak.Tunjangan_Primary_Slip.ToString + " And [Primary Transaksi] =" + row_Header.Primary.ToString + ""
                    TBL_item.DefaultView.RowFilter = Filter
                    If TBL_item.DefaultView.Count > 0 Then
                        Row_item = TBL_item.DefaultView(0).Row
                        rowcetak.Tunjangan_Amount = Row_item.Amount
                        If Me.ObjKary.GetMartialStatus(rowcetak.Primary_Employee) = ClsMasterKaryawan.EnumMartialStatus.Married Then
                            rowcetak.Tunjangan_Jml_Anak = Me.ObjParameter.GetJmlAnak(rowcetak.Primary_Employee)
                            If rowcetak.Tunjangan_Jml_Anak > 3 Then
                                rowcetak.Tunjangan_Jml_Anak = 3
                            End If
                        End If
                    End If

                    Filter = "[Primary Slip Desc]=" + rowcetak.Tunjangan_Lain_Primary_Slip.ToString + " And [Primary Transaksi] =" + row_Header.Primary.ToString + ""
                    TBL_item.DefaultView.RowFilter = Filter
                    If TBL_item.DefaultView.Count > 0 Then
                        Row_item = TBL_item.DefaultView(0).Row
                        rowcetak.Tunjangan_Lain_Amount = Row_item.Amount
                    End If

                    Filter = "[Primary Slip Desc]=" + rowcetak.Bonus_Primary_Slip.ToString + " And [Primary Transaksi] =" + row_Header.Primary.ToString + ""
                    TBL_item.DefaultView.RowFilter = Filter
                    If TBL_item.DefaultView.Count > 0 Then
                        Row_item = TBL_item.DefaultView(0).Row
                        rowcetak.Bonus_Amount = Row_item.Amount
                    End If

                    Filter = "[Primary Slip Desc]=" + rowcetak.Others_Primary_Slip.ToString + " And [Primary Transaksi] =" + row_Header.Primary.ToString + ""
                    TBL_item.DefaultView.RowFilter = Filter
                    If TBL_item.DefaultView.Count > 0 Then
                        Row_item = TBL_item.DefaultView(0).Row
                        rowcetak.Others_Amount = Row_item.Amount
                    End If

                    Filter = "[Primary Slip Desc]=" + rowcetak.JPK_Primary_Slip.ToString + " And [Primary Transaksi] =" + row_Header.Primary.ToString + ""
                    TBL_item.DefaultView.RowFilter = Filter
                    If TBL_item.DefaultView.Count > 0 Then
                        Row_item = TBL_item.DefaultView(0).Row
                        rowcetak.JPK_Amount = Row_item.Amount
                    End If

                    Filter = "[Primary Slip Desc]=" + rowcetak.Claim_Primary_Slip.ToString + " And [Primary Transaksi] =" + row_Header.Primary.ToString + ""
                    TBL_item.DefaultView.RowFilter = Filter
                    If TBL_item.DefaultView.Count > 0 Then
                        Row_item = TBL_item.DefaultView(0).Row
                        rowcetak.Claim_Amount = Row_item.Amount
                    End If

                    Filter = "[Primary Slip Desc]=" + rowcetak.Paid_Middle_Primary_Slip.ToString + " And [Primary Transaksi] =" + row_Header.Primary.ToString + ""
                    TBL_Other.DefaultView.RowFilter = Filter
                    If TBL_Other.DefaultView.Count > 0 Then
                        Row_Other = TBL_Other.DefaultView(0).Row
                        rowcetak.Paid_Middle_Amount = Row_Other.Amount
                    End If

                    Filter = "[Primary Slip Desc]=" + rowcetak.Koperasi_Primary_Slip.ToString + " And [Primary Transaksi] =" + row_Header.Primary.ToString + ""
                    TBL_Other.DefaultView.RowFilter = Filter
                    If TBL_Other.DefaultView.Count > 0 Then
                        Row_Other = TBL_Other.DefaultView(0).Row
                        rowcetak.Koperasi_Amount = Row_Other.Amount
                    End If

                    Filter = "[Primary Slip Desc]=" + rowcetak.Others_Deduction_Primary_Slip.ToString + " And [Primary Transaksi] =" + row_Header.Primary.ToString + ""
                    TBL_Other.DefaultView.RowFilter = Filter
                    If TBL_Other.DefaultView.Count > 0 Then
                        Row_Other = TBL_Other.DefaultView(0).Row
                        rowcetak.Others_Deduction_Amount = Row_Other.Amount
                    End If

                    rowcetak.Total_Salary = rowcetak.Basic_Amount + rowcetak.PlusMin_Amount + rowcetak.Jamsostek_Amount + rowcetak.Transport_Amount _
                                            + rowcetak.OT_Weekday_1_Amount + rowcetak.OT_Weekday_2_Amount + rowcetak.OT_Weekend_1_Amount + rowcetak.OT_Weekend_2_Amount + rowcetak.OT_Weekend_3_Amount

                    rowcetak.Total_Income = rowcetak.Total_Salary + rowcetak.Insentive_Amount + rowcetak.Tunjangan_Amount + rowcetak.Tunjangan_Lain_Amount + rowcetak.Bonus_Amount + _
                                            rowcetak.Others_Amount + rowcetak.JPK_Amount + rowcetak.Claim_Amount

                    rowcetak.Total_Nett_Income = rowcetak.Total_Income - rowcetak.Loan_Amount - rowcetak.Paid_Middle_Amount - rowcetak.Koperasi_Amount - rowcetak.Others_Deduction_Amount

                    'Fill Table Hitung hari Transport
                    Filter_Transport = "[date] >=#" + Format(CDate(row_Header.Start_Date), "MM/dd/yyyy") + "# And [date] <=#" + Format(CDate(row_Header.End_Date), "MM/dd/yyyy") + "#" _
                    + " AND [primary karyawan] = " + row_Header.Primary_Employee.ToString + ""
                    TBL_Transport.DefaultView.RowFilter = Filter_Transport
                    If TBL_Transport.DefaultView.Count > 0 Then
                        rowcetak.Ket_Transport_Faktor_Kali = TBL_Transport.DefaultView.Count
                    End If
                    rowcetak.Ket_Transport_Value = rowcetak.Transport_Amount / rowcetak.Ket_Transport_Faktor_Kali

                    If rowcetak.Basic_Amount = 0 Then
                        Jpk_Persen = 0
                    Else
                        Jpk_Persen = rowcetak.JPK_Amount / rowcetak.Basic_Amount * 100
                    End If
                    rowcetak.JPK_Persen = Math.Round(Jpk_Persen, 2, MidpointRounding.AwayFromZero).ToString
                    'If Jpk_Persen >= 6 Then
                    '    rowcetak.JPK_Persen = "6%"
                    'Else
                    '    rowcetak.JPK_Persen = "3%"
                    'End If
                    'Input Data Overtime
                    If TBL_Overtime.DefaultView.Count > 0 Then
                        TotalOT_Weekday1 = 0
                        TotalOT_Weekday2 = 0
                        TotalOT_Weekend1 = 0
                        TotalOT_Weekend2 = 0
                        TotalOT_Weekend3 = 0
                        For Each DV In TBL_Overtime.DefaultView
                            Row_Overtime = DV.Row
                            TotalOT_Weekday1 += Row_Overtime.OT_Weekday_1
                            TotalOT_Weekday2 += Row_Overtime.OT_Weekday_2
                            TotalOT_Weekend1 += Row_Overtime.OT_Weekend_1
                            TotalOT_Weekend2 += Row_Overtime.OT_Weekend_2
                            TotalOT_Weekend3 += Row_Overtime.OT_Weekend_3
                        Next
                        rowcetak.OT_Weekday_1_Duration = TotalOT_Weekday1
                        rowcetak.OT_Weekday_2_Duration = TotalOT_Weekday2
                        rowcetak.OT_Weekend_1_Duration = TotalOT_Weekend1
                        rowcetak.OT_Weekend_2_Duration = TotalOT_Weekend2
                        rowcetak.OT_Weekend_3_Duration = TotalOT_Weekend3
                        'Hitung Value Overtime
                        If rowcetak.OT_Weekday_1_Duration <> 0 Then
                            rowcetak.OT_Weekday_1_Value = rowcetak.OT_Weekday_1_Amount / rowcetak.OT_Weekday_1_Duration
                        End If
                        If rowcetak.OT_Weekday_2_Duration <> 0 Then
                            rowcetak.OT_Weekday_2_Value = rowcetak.OT_Weekday_2_Amount / rowcetak.OT_Weekday_2_Duration
                        End If
                        If rowcetak.OT_Weekend_1_Duration <> 0 Then
                            rowcetak.OT_Weekend_Value = rowcetak.OT_Weekend_1_Amount / rowcetak.OT_Weekend_1_Duration
                        End If
                        If rowcetak.OT_Weekend_2_Duration <> 0 Then
                            rowcetak.OT_Weekend_2_Value = rowcetak.OT_Weekend_2_Amount / rowcetak.OT_Weekend_2_Duration
                        End If
                        If rowcetak.OT_Weekend_3_Duration <> 0 Then
                            rowcetak.OT_Weekend_3_Value = rowcetak.OT_Weekend_3_Amount / rowcetak.OT_Weekend_3_Duration
                        End If
                    End If
                    'Koding untuk Payment
                    If TBL_Pay.Rows.Count > 0 Then
                        Dim HasilPaymentBCAMandiri As Decimal = 0
                        Dim HasilPaymentCash As Decimal = 0
                        Filter = "[primary akun payment] = " + mPrimarySettingAkunBCA.ToString + " And [primary transaksi] =" + row_Header.Primary.ToString + ""
                        TBL_Pay.DefaultView.RowFilter = Filter
                        If TBL_Pay.DefaultView.Count > 0 Then
                            Row_Pay = TBL_Pay.DefaultView(0).Row
                            rowcetak.Amount_Transfer_1 = Row_Pay.amount
                            rowcetak.Desc_Transfer_1 = "BCA"
                            HasilPaymentBCAMandiri += Row_Pay.amount
                        End If
                        Filter = "[primary akun payment] = " + mPrimarySettingAkunMandiri.ToString + " And [primary transaksi] =" + row_Header.Primary.ToString + ""
                        TBL_Pay.DefaultView.RowFilter = Filter
                        If TBL_Pay.DefaultView.Count > 0 Then
                            Row_Pay = TBL_Pay.DefaultView(0).Row
                            rowcetak.Amount_Transfer_2 = Row_Pay.amount
                            rowcetak.Desc_Transfer_2 = "Mandiri"
                            HasilPaymentBCAMandiri += Row_Pay.amount
                        End If
                        Filter = "[primary akun payment] <> " + mPrimarySettingAkunBCA.ToString + " AND [primary akun payment] <> " + mPrimarySettingAkunMandiri.ToString + " And [primary transaksi] =" + row_Header.Primary.ToString + ""
                        TBL_Pay.DefaultView.RowFilter = Filter
                        If TBL_Pay.DefaultView.Count > 0 Then
                            For Each DV In TBL_Pay.DefaultView
                                Row_Pay = DV.Row
                                HasilPaymentCash += Row_Pay.amount
                            Next
                            rowcetak.Amount_Cash = HasilPaymentCash
                        Else
                            rowcetak.Amount_Cash = 0
                        End If
                    End If
                    'Pergantian Group Untuk Cetak Setengah Letter
                    'If Me.ObjCD.CekGroupingCetakanSalaryStnghLetter(ij) = True Then
                    xGroupCetak += 1
                    'End If
                    rowcetak.Group = xGroupCetak
                    '-------------------------
                    TBLCetak.Rows.Add(rowcetak)
                Next
            End If
            'Laporan
            If Me.mFormatCetakan = EnumFormatCetakan.A5 Then
                Report.SetDataSource(DS)
                Report.SetParameterValue(Report.Parameter_Format_Decimal.ParameterFieldName, Me.mDefPrec)
                vCRV.ReportSource = Report
            ElseIf Me.mFormatCetakan = EnumFormatCetakan.Continuous Then
                Report_Stngh.SetDataSource(DS)
                Report_Stngh.SetParameterValue(Report_Stngh.Parameter_Format_Decimal.ParameterFieldName, Me.mDefPrec)
                vCRV.ReportSource = Report_Stngh
            End If
            vCRV.Show()


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Header.Dispose()
            DTA_item.Dispose()
            DTA_Overtime.Dispose()
            DTA_Transport.Dispose()
            DTA_Other.Dispose()
            TBL_Header.Dispose()
            TBL_item.Dispose()
            TBL_Overtime.Dispose()
            TBL_Transport.Dispose()
            TBL_Other.Dispose()
            DTA_Pay.Dispose()
            TBL_Pay.Dispose()
            DTARound.Dispose()
            TBLRound.Dispose()
        End Try
    End Sub
End Class