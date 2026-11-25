Option Explicit On
Option Strict Off

Imports Ciloci.Flee
Imports Ciloci.Flee.CalcEngine

Public Class ClsMasterSalaryFormula_item

    Dim ObjSlip As New ClsMasterSalaryFormula_SlipDesc
    Dim ObjAcc As New ClsMasterAkun1
    Dim ObjKaryawan As New ClsMasterKaryawan
    Dim ObjCd As New ClsCekData

    Public Enum EnumCalculateOn As Integer
        Jan
        Feb
        Mar
        Apr
        May
        June
        Jul
        Agt
        Sept
        Oct
        Nov
        Dec
    End Enum

    Public Enum EnumTax2
        gaji_pensiun
        tunjangan_lainnya
        honorarium
        premi_asuransi
        penerimaan_dalam_bentuk_natura
        Tanteiem_bonus_gratifikasi_jasa_produksi_thr
        iuran_pensiun
        no_tax
        pesangon
    End Enum

    Public Enum EnumTipeConstraitVariable As Integer
        Hardcode
        Kompetensi
        CustomValue
        ServiceRecord
    End Enum

    Public Function ConversionMonth(ByVal CalculateOn As EnumCalculateOn) As Integer
        Try
            If CalculateOn = EnumCalculateOn.Jan Then
                Return 1
            End If
            If CalculateOn = EnumCalculateOn.Feb Then
                Return 2
            End If
            If CalculateOn = EnumCalculateOn.Mar Then
                Return 3
            End If
            If CalculateOn = EnumCalculateOn.Apr Then
                Return 4
            End If
            If CalculateOn = EnumCalculateOn.May Then
                Return 5
            End If
            If CalculateOn = EnumCalculateOn.June Then
                Return 6
            End If
            If CalculateOn = EnumCalculateOn.Jul Then
                Return 7
            End If
            If CalculateOn = EnumCalculateOn.Agt Then
                Return 8
            End If
            If CalculateOn = EnumCalculateOn.Sept Then
                Return 9
            End If
            If CalculateOn = EnumCalculateOn.Oct Then
                Return 10
            End If
            If CalculateOn = EnumCalculateOn.Nov Then
                Return 11
            End If
            If CalculateOn = EnumCalculateOn.Dec Then
                Return 12
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Enum EnumRoundingType As Integer
        No_Rounding
        Round_Down
        Round_Middlepoint
        Round_Up
    End Enum

    Public Function GetTipeTaxComponent_Str(ByVal vValue As EnumTax2) As String
        Try
            If vValue = EnumTax2.pesangon Then
                Return "PESANGON"
            ElseIf vValue = EnumTax2.no_tax Then
                Return "BUKAN OBJEK PAJAK"
            ElseIf vValue = EnumTax2.iuran_pensiun Then
                Return "IURAN PENSIUN ATAU IURAN THT/JHT"
            ElseIf vValue = EnumTax2.Tanteiem_bonus_gratifikasi_jasa_produksi_thr Then
                Return "TANTIEM, BONUS, GRATIFIKASI, JASA PRODUKSI DAN THR"
            ElseIf vValue = EnumTax2.penerimaan_dalam_bentuk_natura Then
                Return "PENERIMAAN DALAM BENTUK NATURA DAN KENIKMATAN LAINNYA YANG DIKENAKAN PEMOTONGAN PPh PASAL 21"
            ElseIf vValue = EnumTax2.premi_asuransi Then
                Return "PREMI ASURANSI YANG DIBAYARKAN PEMBERI KERJA"
            ElseIf vValue = EnumTax2.honorarium Then
                Return "HONORARIUM DAN IMBALAN LAIN SEJENISNYA"
            ElseIf vValue = EnumTax2.tunjangan_lainnya Then
                Return "TUNJANGAN LAINNYA, UANG LEMBUR DAN SEBAGAINYA"
            ElseIf vValue = EnumTax2.gaji_pensiun Then
                Return "GAJI/PENSIUN ATAU THT/JHT"
            Else
                Return "GAJI/PENSIUN ATAU THT/JHT"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "GAJI/PENSIUN ATAU THT/JHT"
        End Try
    End Function

    Public Function GetTipeTaxComponent(ByVal vValue As String) As EnumTax2
        Try
            If vValue = "PESANGON" Then
                Return EnumTax2.pesangon
            ElseIf vValue = "BUKAN OBJEK PAJAK" Then
                Return EnumTax2.no_tax
            ElseIf vValue = "IURAN PENSIUN ATAU IURAN THT/JHT" Then
                Return EnumTax2.iuran_pensiun
            ElseIf vValue = "TANTIEM, BONUS, GRATIFIKASI, JASA PRODUKSI DAN THR" Then
                Return EnumTax2.Tanteiem_bonus_gratifikasi_jasa_produksi_thr
            ElseIf vValue = "PENERIMAAN DALAM BENTUK NATURA DAN KENIKMATAN LAINNYA YANG DIKENAKAN PEMOTONGAN PPh PASAL 21" Then
                Return EnumTax2.penerimaan_dalam_bentuk_natura
            ElseIf vValue = "PREMI ASURANSI YANG DIBAYARKAN PEMBERI KERJA" Then
                Return EnumTax2.premi_asuransi
            ElseIf vValue = "HONORARIUM DAN IMBALAN LAIN SEJENISNYA" Then
                Return EnumTax2.honorarium
            ElseIf vValue = "TUNJANGAN LAINNYA, UANG LEMBUR DAN SEBAGAINYA" Then
                Return EnumTax2.tunjangan_lainnya
            ElseIf vValue = "GAJI/PENSIUN ATAU THT/JHT" Then
                Return EnumTax2.gaji_pensiun
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub Simpan(ByVal vFrm As FrmMasterSalaryFormula_Component, ByVal vTblVar As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_variable1DataTable, _
    ByVal vTblConst As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_constraint1DataTable, ByVal vtblLog As DseMasterSalaryFormula.tbl_log_userDataTable)
        Dim row As DataGridViewRow
        Dim CboDgv As DataGridViewComboBoxCell
        Dim DtaSalaryComponent As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_salarycomponentTableAdapter
        Dim DtaSalaryComponent1 As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_salarycomponent1TableAdapter
        Dim DtaEmp As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_component_detailemployeeTableAdapter
        Dim DtaVar As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_salarycomponent_variableTableAdapter
        Dim DtaConts As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_salarycomponent_constraintTableAdapter
        Dim DtaLog As New DseMasterSalaryFormulaTableAdapters.tbl_log_userTableAdapter
        Dim RowVar As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_variable1Row
        Dim RowConst As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_constraint1Row
        Dim RowLog As DseMasterSalaryFormula.tbl_log_userRow
        Dim Rutin_Int As Integer = 0

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DtaConts.Connection.ConnectionString = Core_Module.GetConstring
            DtaEmp.Connection.ConnectionString = DtaConts.Connection.ConnectionString
            DtaSalaryComponent.Connection.ConnectionString = DtaConts.Connection.ConnectionString
            DtaVar.Connection.ConnectionString = DtaConts.Connection.ConnectionString
            DtaLog.Connection.ConnectionString = DtaConts.Connection.ConnectionString


            'hapus salary component
            DtaSalaryComponent.Delete(vFrm.mPrimarySalary)

            'simpan salary component
            Dim type As Integer = -1
            Dim tax2 As EnumTax2
            Dim CalPeriod As Integer = -1
            Dim CalOn As Integer = -1
            Dim str As String = ""
            For Each row In vFrm.DgvComponent.Rows
                If row.Index <> vFrm.DgvComponent.RowCount - 1 Then
                    'type
                    CboDgv = row.Cells(vFrm.ColIncome_ComponentTipe.Index)
                    type = CboDgv.Items.IndexOf(row.Cells(vFrm.ColIncome_ComponentTipe.Index).Value)
                    'tax
                    CboDgv = row.Cells(vFrm.ColIncome_TaxComponent.Index)
                    tax2 = Me.GetTipeTaxComponent(row.Cells(vFrm.ColIncome_TaxComponent.Index).Value)
                    'cal period
                    CboDgv = row.Cells(vFrm.ColIncome_CalculatePeriod.Index)
                    CalPeriod = CboDgv.Items.IndexOf(row.Cells(vFrm.ColIncome_CalculatePeriod.Index).Value)
                    'cal on
                    CboDgv = row.Cells(vFrm.ColIncome_CalculateOn.Index)
                    CalOn = CboDgv.Items.IndexOf(row.Cells(vFrm.ColIncome_CalculateOn.Index).Value)
                    If row.Cells(vFrm.ColIncome_Routine.Index).Value = Boolean.TrueString Then
                        Rutin_Int = 1
                    Else
                        Rutin_Int = 0
                    End If


                    Dim CreatedDate As Date = Now
                    Dim ModifyDate As Date = Now
                    If row.Cells(vFrm.Colincome_Primary.Index).Value <> -1 Then
                        CreatedDate = CDate(row.Cells(vFrm.ColIncome_CreatedDate.Index).Value)
                        ModifyDate = CDate(row.Cells(vFrm.ColIncome_ModifiedDate.Index).Value)
                        If row.Cells(vFrm.ColIsModify.Index).Value = Boolean.TrueString Then
                            ModifyDate = Now
                        End If
                    End If


                    'MsgBox(row.Cells(vFrm.ColIncome_CreatedDate.Index).Value)

                    'simpan salary component
                    DtaSalaryComponent1.InsertQuery(vFrm.mPrimarySalary, type, row.Cells(vFrm.ColIncome_Desc.Index).Value, row.Cells(vFrm.ColIncome_PrimarySlipDesc.Index).Value, _
                    tax2, row.Cells(vFrm.ColIncome_PrimaryExpenseAccount.Index).Value, row.Cells(vFrm.ColFormula.Index).Value, CalPeriod, CalOn, _
                    row.Cells(vFrm.ColYearlyMaxValue.Index).Value, row.Cells(vFrm.ColMonthlyMaxValue.Index).Value, row.Cells(vFrm.ColIncome_Identifier.Index).Value, _
                    row.Cells(vFrm.ColRoundingUnit.Index).Value, row.Cells(vFrm.ColRoundingType.Index).Value, row.Cells(vFrm.ColYearlyMinValue.Index).Value, _
                    row.Cells(vFrm.ColMonthlyMinValue.Index).Value, row.Cells(vFrm.ColYearlyDate.Index).Value, row.Cells(vFrm.ColIncome_Urutan.Index).Value, _
                    Rutin_Int, row.Cells(vFrm.ColIncome_PrimaryPaymentAccount.Index).Value, _
                    row.Cells(vFrm.ColIncome_CreatedBy.Index).Value, _
                    CreatedDate, _
                    row.Cells(vFrm.ColIncome_ModifiedBy.Index).Value, _
                    ModifyDate)

                    'If DtaSalaryComponent.Insert(vFrm.mPrimarySalary, type, row.Cells(vFrm.ColIncome_Desc.Index).Value, row.Cells(vFrm.ColIncome_PrimarySlipDesc.Index).Value, _
                    'tax, row.Cells(vFrm.ColIncome_PrimaryExpenseAccount.Index).Value, row.Cells(vFrm.ColFormula.Index).Value, CalPeriod, CalOn, _
                    'row.Cells(vFrm.ColYearlyMaxValue.Index).Value, row.Cells(vFrm.ColMonthlyMaxValue.Index).Value, row.Cells(vFrm.ColIncome_Identifier.Index).Value, _
                    'row.Cells(vFrm.ColRoundingUnit.Index).Value, row.Cells(vFrm.ColRoundingType.Index).Value, row.Cells(vFrm.ColYearlyMinValue.Index).Value, _
                    'row.Cells(vFrm.ColMonthlyMinValue.Index).Value, row.Cells(vFrm.ColYearlyDate.Index).Value, row.Cells(vFrm.ColIncome_Urutan.Index).Value, _
                    'Rutin_Int, row.Cells(vFrm.ColIncome_PrimaryPaymentAccount.Index).Value) = True Then
                    'Else

                    '    str &= vbNewLine & vFrm.mPrimarySalary & "," & type & "," & row.Cells(vFrm.ColIncome_Desc.Index).Value & "," & _
                    '          row.Cells(vFrm.ColIncome_PrimarySlipDesc.Index).Value & "," & _
                    '                            tax & "," & row.Cells(vFrm.ColIncome_PrimaryExpenseAccount.Index).Value & "," & _
                    '                            row.Cells(vFrm.ColFormula.Index).Value & "," & CalPeriod & "," & CalOn & "," & _
                    '                            row.Cells(vFrm.ColYearlyMaxValue.Index).Value & "," & _
                    '                            row.Cells(vFrm.ColMonthlyMaxValue.Index).Value & "," & _
                    '                            row.Cells(vFrm.ColIncome_Identifier.Index).Value & "," & _
                    '                            row.Cells(vFrm.ColRoundingUnit.Index).Value & "," & _
                    '                            row.Cells(vFrm.ColRoundingType.Index).Value & "," & _
                    '                            row.Cells(vFrm.ColYearlyMinValue.Index).Value & "," & _
                    '                            row.Cells(vFrm.ColMonthlyMinValue.Index).Value & "," & _
                    '                            row.Cells(vFrm.ColYearlyDate.Index).Value & "," & row.Cells(vFrm.ColIncome_Urutan.Index).Value & "," & _
                    '                            Rutin_Int & "," & row.Cells(vFrm.ColIncome_PrimaryPaymentAccount.Index).Value
                    'End If
                End If
            Next

            'MsgBox(str)

            'hapus employee detail
            DtaEmp.Delete(vFrm.mPrimarySalary)
            'simpan employee
            Dim TaxCal As Integer = -1
            Dim TaxComp As Integer = -1
            Dim remarks As String = ""
            For Each row In vFrm.dgvEmployee.Rows
                If row.Index <> vFrm.dgvEmployee.RowCount - 1 Then
                    'tax cal
                    CboDgv = row.Cells(vFrm.ColEmp_TaxCalculation.Index)
                    TaxCal = CboDgv.Items.IndexOf(row.Cells(vFrm.ColEmp_TaxCalculation.Index).Value)
                    'tax component
                    CboDgv = row.Cells(vFrm.ColEmp_TaxComponent.Index)
                    TaxComp = CboDgv.Items.IndexOf(row.Cells(vFrm.ColEmp_TaxComponent.Index).Value)
                    'remarks 
                    If String.IsNullOrEmpty(row.Cells(vFrm.ColEmp_Remarks.Index).Value) = True Then
                        remarks = ""
                    Else
                        remarks = row.Cells(vFrm.ColEmp_Remarks.Index).Value
                    End If

                    Dim CreatedDate As Date = Now
                    Dim ModifyDate As Date = Now
                    If row.Cells(vFrm.ColEmp_Primary.Index).Value <> -1 Then
                        CreatedDate = CDate(row.Cells(vFrm.ColEmp_CreatedDate.Index).Value)
                        ModifyDate = CDate(row.Cells(vFrm.ColEmp_ModifiedDate.Index).Value)
                        If row.Cells(vFrm.ColEmp_IsModify.Index).Value = Boolean.TrueString Then
                            ModifyDate = Now
                        End If
                    End If

                    DtaEmp.Insert(vFrm.mPrimarySalary, row.Cells(vFrm.ColEmp_PrimaryEmployee.Index).Value, row.Cells(vFrm.ColEmp_Basic.Index).Value, _
                    remarks, TaxCal, row.Cells(vFrm.ColEmp_PrimarySlipDesc.Index).Value, _
                    TaxComp, row.Cells(vFrm.ColEmp_PrimaryExpenseAcc.Index).Value, _
                    row.Cells(vFrm.ColEmp_PrimaryPayment.Index).Value, _
                    row.Cells(vFrm.ColEmp_CreatedBy.Index).Value, _
                    CreatedDate, _
                    row.Cells(vFrm.ColEmp_ModifedBy.Index).Value, _
                     ModifyDate)
                End If
            Next

            'hapus item variable
            DtaVar.Delete(vFrm.mPrimarySalary)
            'simpan item variable
            For Each RowVar In vTblVar.Rows
                With RowVar
                    DtaVar.Insert(.Primary_Transaksi, .Primary_Parameter, .Type, .Variable, .Identifier_Salary_Component, .Rounding_Unit, .Rounding_Type, .Type2)
                End With
            Next

            'hapus item constrain
            DtaConts.Delete(vFrm.mPrimarySalary)
            'simpan item constrain
            For Each RowConst In vTblConst.Rows
                With RowConst
                    DtaConts.Insert(.Primary_Transaksi, .Primary_Parameter, .Type, .Compare, .Ref_Value, .Identifier_Salary_Component, .Type2)
                End With
            Next

            If vtblLog.Rows.Count > 0 Then
                'Dim I As New Form2
                'I.DataGridView1.DataSource = Nothing
                'I.DataGridView1.DataSource = vtblLog
                'I.ShowDialog()
                For Each RowLog In vtblLog.Rows
                    With RowLog
                        DtaLog.Insert(.Activity_Type, .User, .Activity_Time, _
                        .Remarks, .Primary_Salary_Formula, vFrm.ObjSalaryFormula.GetCodeByPrimary(.Primary_Salary_Formula))
                    End With
                Next
            End If

            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaSalaryComponent.Dispose()
            DtaEmp.Dispose()
            DtaVar.Dispose()
            DtaConts.Dispose()
        End Try
    End Sub

    'Private Function Simpan_Component(ByVal i As FrmMasterSalaryFormula_Component, _
    '                                           ByVal vTBLST As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponentDataTable, ByVal VPrimaryTransaksi As Double) As Boolean

    '    Dim conStr As String = My.Settings.Core_DbConnectionString
    '    Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
    '    Dim sql As String = ""
    '    Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
    '    Dim BatchQty As Double = 500
    '    Dim StartRowIndex As Double = 0
    '    Dim DTAST As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_salarycomponentTableAdapter

    '    Try
    '        'open connection
    '        Conn.Open()

    '        While StartRowIndex < i.DgvComponent.RowCount
    '            sql = Me.GetInsert_StockTarget_Script(VPrimaryTransaksi, StartRowIndex, BatchQty, vTBLST)
    '            script.Query = sql
    '            'execute script
    '            script.Execute()
    '            StartRowIndex += BatchQty
    '        End While

    '        Conn.Close()

    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        Conn = Nothing
    '        script = Nothing
    '    End Try
    'End Function

    'Private Function GetInsert_StockTarget_Script(ByVal vPrimaryTransaksi As Double, _
    '                                             ByVal vStartRowIndex As Double, _
    '                                             ByVal BatchQty As Double, _
    '                                             ByVal frm As FrmMasterSalaryFormula_Component) As String
    '    Dim retval As String = ""
    '    Dim row As New DataGridViewRow
    '    Dim Delimiter As String = ","
    '    Dim Filter As String = ""
    '    Dim PrimaryCustomer As Double = -1
    '    Dim PrimaryProduct As Double = -1

    '    Try
    '        retval = "INSERT INTO `tblmaster_salary_formula_SalaryComponent` " & _
    '                              "(`Primary Transaksi`, " & _
    '                              "`Type`, " & _
    '                              "`Description`, " & _
    '                              "`Primary Slip Desc`, " & _
    '                              "`Tax Component`, " & _
    '                              "`Primary Expense Account`, " & _
    '                              "`Formula`, " & _
    '                              "`Calculation Period`, " & _
    '                              "`Calculation On`, " & _
    '                              "`Yearly Max Value`, " & _
    '                              "`MOnthly Max Value`, " & _
    '                              "`Identifier`, " & _
    '                              "`Rounding Unit`, " & _
    '                              "`Rounding type`, " & _
    '                              "`Yearly Min Value`, " & _
    '                              "`Monthly Min Value`, " & _
    '                              "`Yearly Date`, " & _
    '                              "`Urutan`, " & _
    '                              "`rutin_disetahunkan`, " & _
    '                              "`Primary payment account`) VALUES "

    '        For i As Double = 0 To BatchQty - 1 Step 1
    '            If frm.DgvComponent.RowCount > vStartRowIndex + i Then
    '                row = frm.DgvComponent.Rows(i + vStartRowIndex)
    '                'jika baris terakhir dgv atau baris terakhir untuk batch tsb
    '                If frm.DgvComponent.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
    '                    Delimiter = ";"
    '                End If

    '                With row
    '                    .Cells (frm.Col = Me.ObjCd.DigitGrouping(.Jumlah_Period_Using, defPrec)
    '                    .Jumlah_Period_Total = Me.ObjCd.DigitGrouping(.Jumlah_Period_Total, defPrec)
    '                    .Jumlah_Period = Me.ObjCd.DigitGrouping(.Jumlah_Period, defPrec)
    '                    .Jumlah_Pindah = Me.ObjCd.DigitGrouping(.Jumlah_Pindah, defPrec)
    '                    .Total_Historical_Using = Me.ObjCd.DigitGrouping(.Total_Historical_Using, defPrec)
    '                    .Total_Sales = Me.ObjCd.DigitGrouping(.Total_Sales, defPrec)
    '                    .Target_Bulan = Me.ObjCd.DigitGrouping(.Target_Bulan, defPrec)
    '                    .Qty_Override = Me.ObjCd.DigitGrouping(.Qty_Override, defPrec)
    '                    .Stock_Target = Me.ObjCd.DigitGrouping(.Stock_Target, defPrec)
    '                    .Bln_Override = Me.ObjCd.DigitGrouping(.Bln_Override, defPrec)


    '                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString & ", '" & _
    '                                                                          .Sales_Code.ToString & "', '" & _
    '                                                                           Format(CDate(.Start_Period), "yyyy-MM-dd") & "', '" & _
    '                                                                           Format(CDate(.End_Period), "yyyy-MM-dd") & "', " & _
    '                                                                           CDec(.Jumlah_Period_Using).ToString & ", " & _
    '                                                                           CDec(.Jumlah_Period_Total).ToString & ", " & _
    '                                                                           CDec(.Jumlah_Period).ToString & ", " & _
    '                                                                           CDec(.Jumlah_Pindah).ToString & ", " & _
    '                                                                           CDec(.Total_Historical_Using).ToString & ", " & _
    '                                                                           CDec(.Total_Sales).ToString & ", " & _
    '                                                                           CDec(.Rata2Using).ToString & ", " & _
    '                                                                           CDec(.Target_Bulan).ToString & ", " & _
    '                                                                           CDec(.Qty_Override).ToString & ", '" & _
    '                                                                           .No_Transaksi & "', '" & _
    '                                                                           .txt & "', " & _
    '                                                                           CDec(.Stock_Target).ToString & ", " & _
    '                                                                           CDec(.Bln_Override).ToString & ", " & _
    '                                                                           CDec(.Jenis).ToString & ")" + Delimiter
    '                End With
    '            Else
    '                Exit For

    '            End If

    '        Next
    '        retval = retval + vbCrLf + vbCrLf + "COMMIT;"
    '        Return retval
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return ""
    '    Finally
    '        '  ObjCD = Nothing
    '    End Try
    'End Function

    Public Sub IsiDgvComponent(ByVal vFrm As FrmMasterSalaryFormula_Component, ByVal vDefPrec As Integer)
        Dim dta As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_salarycomponent1TableAdapter
        Dim Tbl As New DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent1DataTable
        Dim row As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent1Row

        Dim DtaEmp As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_component_detailemployee1TableAdapter
        Dim TblEmp As New DseMasterSalaryFormula.tblmaster_salary_formula_component_detailemployee1DataTable
        Dim rowEmp As DseMasterSalaryFormula.tblmaster_salary_formula_component_detailemployee1Row

        Dim rowDgv As DataGridViewRow
        Dim CboDgv As DataGridViewComboBoxCell
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            DtaEmp.Connection.ConnectionString = dta.Connection.ConnectionString

            dta.Fill(Tbl, vFrm.mPrimarySalary)
            dta.Fill(vFrm.tblSC, vFrm.mPrimarySalary)
            'Dim Ix As New Core_Module.Form2
            'Ix.DataGridView1.DataSource = Nothing
            'Ix.DataGridView1.DataSource = Tbl
            'Ix.ShowDialog()

            vFrm.DgvComponent.Rows.Clear()
            vFrm.SedangIsiDGV = True
            For Each row In Tbl.Rows
                vFrm.DgvComponent.Rows.Add()
                rowDgv = vFrm.DgvComponent.Rows(vFrm.DgvComponent.RowCount - 2)
                CboDgv = rowDgv.Cells(vFrm.ColIncome_ComponentTipe.Index)
                rowDgv.Cells(vFrm.ColIncome_ComponentTipe.Index).Value = CboDgv.Items.Item(row.Type)
                rowDgv.Cells(vFrm.ColIncome_Desc.Index).Value = row.Description
                rowDgv.Cells(vFrm.ColIncome_PrimarySlipDesc.Index).Value = row.Primary_Slip_Desc
                rowDgv.Cells(vFrm.ColIncome_SlipDescription.Index).Value = row.Slip_Desc_Code
                CboDgv = rowDgv.Cells(vFrm.ColIncome_TaxComponent.Index)
                rowDgv.Cells(vFrm.ColIncome_TaxComponent.Index).Value = CboDgv.Items.Item(row.Tax_Component)
                rowDgv.Cells(vFrm.ColFormula.Index).Value = row.Formula
                CboDgv = rowDgv.Cells(vFrm.ColIncome_CalculatePeriod.Index)
                rowDgv.Cells(vFrm.ColIncome_CalculatePeriod.Index).Value = CboDgv.Items.Item(row.Calculation_Period)
                CboDgv = rowDgv.Cells(vFrm.ColIncome_CalculateOn.Index)
                rowDgv.Cells(vFrm.ColIncome_CalculateOn.Index).Value = CboDgv.Items.Item(row.Calculate_On)
                If row.Primary_Expense_Account = -1 Then
                    rowDgv.Cells(vFrm.ColIncome_PrimaryExpenseAccount.Index).Value = -1
                    rowDgv.Cells(vFrm.ColIncome_ExpenseAccount.Index).Value = ""
                Else
                    rowDgv.Cells(vFrm.ColIncome_PrimaryExpenseAccount.Index).Value = row.Primary_Expense_Account
                    rowDgv.Cells(vFrm.ColIncome_ExpenseAccount.Index).Value = row.Expense_Account_Code
                End If
                rowDgv.Cells(vFrm.ColRoundingUnit.Index).Value = row.Rounding_Unit
                rowDgv.Cells(vFrm.ColRoundingType.Index).Value = row.Rounding_Type
                rowDgv.Cells(vFrm.ColYearlyMaxValue.Index).Value = row.Yearly_Max_Value
                rowDgv.Cells(vFrm.ColYearlyMinValue.Index).Value = row.Yearly_Min_Value
                rowDgv.Cells(vFrm.ColYearlyDate.Index).Value = row.Yearly_Date
                rowDgv.Cells(vFrm.ColMonthlyMaxValue.Index).Value = row.Monthly_Max_Value
                rowDgv.Cells(vFrm.ColMonthlyMinValue.Index).Value = row.Monthly_Min_Value

                'If vCopy = False Then
                'rowDgv.Cells(vFrm.ColIncome_Urutan.Index).Value = row.Urutan
                rowDgv.Cells(vFrm.ColIncome_Urutan.Index).Value = row.Identifier.Substring(row.Identifier.IndexOf("_") + 1)
                rowDgv.Cells(vFrm.ColIncome_Identifier.Index).Value = row.Identifier
                If row.rutin_disetahunkan = 1 Then
                    rowDgv.Cells(vFrm.ColIncome_Routine.Index).Value = Boolean.TrueString
                Else
                    rowDgv.Cells(vFrm.ColIncome_Routine.Index).Value = Boolean.FalseString
                End If
                rowDgv.Cells(vFrm.ColIncome_PrimaryPaymentAccount.Index).Value = row.primary_payment_account
                rowDgv.Cells(vFrm.ColIncome_PaymentAccount.Index).Value = row.Payment_Account_Code
                rowDgv.Cells(vFrm.ColIncome_CreatedBy.Index).Value = row.Created_By
                rowDgv.Cells(vFrm.ColIncome_ModifiedBy.Index).Value = row.Modified_By
                rowDgv.Cells(vFrm.ColIncome_CreatedDate.Index).Value = Format(row.Created_Date, "dd-MMM-yyyy") + " " + CDate(row.Created_Date).ToShortTimeString
                rowDgv.Cells(vFrm.ColIncome_ModifiedDate.Index).Value = Format(row.Modified_Date, "dd-MMM-yyyy") + " " + CDate(row.Modified_Date).ToShortTimeString
                'Else
                '    rowDgv.Cells(vFrm.ColIncome_Urutan.Index).Value = vFrm.DgvComponent.Rows.Count
                '    rowDgv.Cells(vFrm.ColIncome_Identifier.Index).Value = vPrimary.ToString + "_" + vFrm.DgvComponent.Rows.Count.ToString
                '    vFrm.MAXUrutan = vFrm.DgvComponent.Rows.Count
                'End If


                CboDgv = rowDgv.Cells(vFrm.ColIncome_ComponentTipe.Index)
                If CboDgv.Items.IndexOf(rowDgv.Cells(vFrm.ColIncome_ComponentTipe.Index).Value) = FrmMasterSalaryFormula_Component.EnumType.Income Then
                    CboDgv = rowDgv.Cells(vFrm.ColIncome_TaxComponent.Index)
                    If CboDgv.Items.Contains("IURAN PENSIUN ATAU IURAN THT/JHT") = True Then
                        CboDgv.Items.Remove("IURAN PENSIUN ATAU IURAN THT/JHT")
                    End If
                ElseIf CboDgv.Items.IndexOf(rowDgv.Cells(vFrm.ColIncome_ComponentTipe.Index).Value) = FrmMasterSalaryFormula_Component.EnumType.Deduction Or _
                CboDgv.Items.IndexOf(rowDgv.Cells(vFrm.ColIncome_ComponentTipe.Index).Value) = FrmMasterSalaryFormula_Component.EnumType.CompanyExpense Then
                    CboDgv = rowDgv.Cells(vFrm.ColIncome_TaxComponent.Index)
                    If CboDgv.Items.Contains("IURAN PENSIUN ATAU IURAN THT/JHT") = False Then
                        CboDgv.Items.Add("IURAN PENSIUN ATAU IURAN THT/JHT")
                    End If
                End If
                rowDgv.Cells(vFrm.Colincome_Primary.Index).Value = row.Primary
            Next
            vFrm.SedangIsiDGV = False
            'betulkan nilai max urutan dan identifier untuk item terakhir
            Dim Max As Double = 0
            Max = vFrm.GetMAXUrutan() + 1
            vFrm.DgvComponent.Item(vFrm.ColIncome_Identifier.Index, vFrm.DgvComponent.Rows.Count - 1).Value = CInt(vFrm.mPrimarySalary).ToString + "_" + Max.ToString
            vFrm.DgvComponent.Item(vFrm.ColIncome_Urutan.Index, vFrm.DgvComponent.Rows.Count - 1).Value = Max.ToString
            vFrm.MAXUrutan = CInt(vFrm.DgvComponent.Item(vFrm.ColIncome_Urutan.Index, vFrm.DgvComponent.Rows.Count - 1).Value)

            'If vCopy = False Then
            DtaEmp.Fill(TblEmp, vFrm.mPrimarySalary)
            DtaEmp.Fill(vFrm.tblEMP, vFrm.mPrimarySalary)
            vFrm.dgvEmployee.Rows.Clear()
            For Each rowEmp In TblEmp.Rows
                vFrm.dgvEmployee.Rows.Add()
                rowDgv = vFrm.dgvEmployee.Rows(vFrm.dgvEmployee.RowCount - 2)
                rowDgv.Cells(vFrm.ColEmp_EmployeeCode.Index).Value = rowEmp.Employee_Code
                rowDgv.Cells(vFrm.ColEmp_EmployeeName.Index).Value = rowEmp.Employee_Name
                rowDgv.Cells(vFrm.ColEmp_Basic.Index).Value = Me.ObjCd.DigitGrouping(rowEmp.Basic_Salary, vDefPrec)
                CboDgv = rowDgv.Cells(vFrm.ColEmp_TaxCalculation.Index)
                rowDgv.Cells(vFrm.ColEmp_TaxCalculation.Index).Value = CboDgv.Items.Item(rowEmp.Tax_Calculation)
                rowDgv.Cells(vFrm.ColEmp_SalarySlipDesc.Index).Value = rowEmp.Slip_Desc_Code
                CboDgv = rowDgv.Cells(vFrm.ColEmp_TaxComponent.Index)
                rowDgv.Cells(vFrm.ColEmp_TaxComponent.Index).Value = CboDgv.Items.Item(0)
                rowDgv.Cells(vFrm.ColEmp_ExpenseAcc.Index).Value = rowEmp.Expense_Account_Code
                rowDgv.Cells(vFrm.ColEmp_ExpenseAccDesc.Index).Value = rowEmp.Expense_Account_Desc
                rowDgv.Cells(vFrm.ColEmp_PaymentAcc.Index).Value = rowEmp.Payment_Account_Code
                rowDgv.Cells(vFrm.ColEmp_PaymentAccDesc.Index).Value = rowEmp.Payment_Account_Desc
                rowDgv.Cells(vFrm.ColEmp_Remarks.Index).Value = rowEmp.Remarks
                rowDgv.Cells(vFrm.ColEmp_PrimaryEmployee.Index).Value = rowEmp.Primary_Karyawan
                rowDgv.Cells(vFrm.ColEmp_PrimarySlipDesc.Index).Value = rowEmp.Primary_Slip_Desc
                rowDgv.Cells(vFrm.ColEmp_PrimaryExpenseAcc.Index).Value = rowEmp.Primary_Expense_Account
                rowDgv.Cells(vFrm.ColEmp_PrimaryPayment.Index).Value = rowEmp.Primary_Default_Payment_Account
                rowDgv.Cells(vFrm.ColEmp_Primary.Index).Value = rowEmp.Primary
                rowDgv.Cells(vFrm.ColEmp_CreatedBy.Index).Value = rowEmp.Created_by
                rowDgv.Cells(vFrm.ColEmp_ModifedBy.Index).Value = rowEmp.Modified_by
                rowDgv.Cells(vFrm.ColEmp_CreatedDate.Index).Value = Format(rowEmp.Created_Date, "dd-MMM-yyyy") + " " + CDate(rowEmp.Created_Date).ToShortTimeString
                rowDgv.Cells(vFrm.ColEmp_ModifiedDate.Index).Value = Format(rowEmp.Modified_Date, "dd-MMM-yyyy") + " " + CDate(rowEmp.Modified_Date).ToShortTimeString

            Next
            'End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            Tbl.Dispose()
            DtaEmp.Dispose()
            TblEmp.Dispose()
        End Try
    End Sub

    Public Sub IsiTabelVarConts(ByVal vFrm As FrmMasterSalaryFormula_Component)

        Dim DtaVar As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_salarycomponent_variableTableAdapter
        Dim TblVar As New DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_variableDataTable
        Dim rowVar As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_variableRow
        Dim DtaConst As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_salarycomponent_constraintTableAdapter
        Dim TblConst As New DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_constraintDataTable
        Dim rowConst As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_constraintRow

        Dim RowVarDummy As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_variable1Row
        Dim RowConstDummy As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_constraint1Row

        Dim RowVarDummyS As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_variable1Row
        Dim RowConstDummyS As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_constraint1Row

        Try
            DtaConst.Connection.ConnectionString = Core_Module.GetConstring
            DtaVar.Connection.ConnectionString = DtaConst.Connection.ConnectionString

            DtaVar.FillBy_PrimarySalary(TblVar, vFrm.mPrimarySalary)
            DtaConst.FillBy_PrimarySalary(TblConst, vFrm.mPrimarySalary)

            'hapus semua isi datatable
            vFrm.TblVar.Rows.Clear()
            vFrm.TblVarOld.Rows.Clear()
            For Each rowVar In TblVar.Rows
                RowVarDummy = vFrm.TblVar.NewRow
                With RowVarDummy
                    .Primary = rowVar.Primary
                    .Primary_Transaksi = rowVar.Primary_Transaksi
                    .Primary_Parameter = rowVar.Primary_Parameter
                    .Type = rowVar.Type
                    .Variable = rowVar.Variable
                    .Identifier_Salary_Component = rowVar.Identifier_Salary_Component
                    .Rounding_Unit = rowVar.Rounding_Unit
                    .Rounding_Type = rowVar.Rounding_Type
                    .Type2 = rowVar.Type2
                End With
                vFrm.TblVar.Rows.Add(RowVarDummy)

                RowVarDummyS = vFrm.TblVarOld.NewRow
                With RowVarDummyS
                    .Primary = rowVar.Primary
                    .Primary_Transaksi = rowVar.Primary_Transaksi
                    .Primary_Parameter = rowVar.Primary_Parameter
                    .Type = rowVar.Type
                    .Variable = rowVar.Variable
                    .Identifier_Salary_Component = rowVar.Identifier_Salary_Component
                    .Rounding_Unit = rowVar.Rounding_Unit
                    .Rounding_Type = rowVar.Rounding_Type
                    .Type2 = rowVar.Type2
                End With
                vFrm.TblVarOld.Rows.Add(RowVarDummyS)
            Next

            'hapus semua isi datatable
            vFrm.TblCons.Rows.Clear()
            vFrm.TblConsOld.Rows.Clear()
            For Each rowConst In TblConst.Rows
                RowConstDummy = vFrm.TblCons.NewRow
                With RowConstDummy
                    .Primary = rowConst.Primary
                    .Primary_Transaksi = rowConst.Primary_Transaksi
                    .Primary_Parameter = rowConst.Primary_Parameter
                    .Type = rowConst.Type
                    .Compare = rowConst.Compare
                    .Ref_Value = rowConst.Ref_Value
                    .Identifier_Salary_Component = rowConst.Identifier_Salary_Component
                    .Type2 = rowConst.Type2
                End With
                vFrm.TblCons.Rows.Add(RowConstDummy)

                RowConstDummyS = vFrm.TblConsOld.NewRow
                With RowConstDummyS
                    .Primary = rowConst.Primary
                    .Primary_Transaksi = rowConst.Primary_Transaksi
                    .Primary_Parameter = rowConst.Primary_Parameter
                    .Type = rowConst.Type
                    .Compare = rowConst.Compare
                    .Ref_Value = rowConst.Ref_Value
                    .Identifier_Salary_Component = rowConst.Identifier_Salary_Component
                    .Type2 = rowConst.Type2
                End With
                vFrm.TblConsOld.Rows.Add(RowConstDummyS)
            Next

          

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaConst.Dispose()
            DtaVar.Dispose()
        End Try
    End Sub

    Public Function GetMAXUrutan(ByVal vPrimaryTrans As Double) As Double
        Dim DTA As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_salarycomponentTableAdapter
        Dim Dummy As System.Nullable(Of Integer)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetMAXUrutanItem(vPrimaryTrans)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPayrollAcces(ByVal vPrimaryTransaksi As Double, ByVal vPrimaryUser As Double) As Boolean
        Dim dta As New DseMasterSalaryFormulaTableAdapters.DataTable2TableAdapter
        Dim jml As New Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            jml = dta.ScalarQuery_GetPayrollAccess(vPrimaryTransaksi, vPrimaryUser)
            If jml.HasValue Then
                If jml.Value > 0 Then
                    Return True
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Sub CopySalaryComponent(ByVal vFrm As FrmMasterSalaryFormula_Component, ByVal vDefPrec As Integer, ByVal vTblVar As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_variable1DataTable, _
    ByVal vTblCons As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_constraint1DataTable, ByVal vPrimaryTrans As Double, ByVal vUserCode As String)
        Dim dta As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_salarycomponent1TableAdapter
        Dim Tbl As New DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent1DataTable
        Dim row As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent1Row
        Dim DtaVar As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_salarycomponent_variableTableAdapter
        Dim TblVar As New DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_variableDataTable
        Dim RowVar As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_variableRow
        Dim DtaConst As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_salarycomponent_constraintTableAdapter
        Dim TblConst As New DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_constraintDataTable
        Dim RowConst As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_constraintRow
        Dim rowDgv As DataGridViewRow
        Dim CboDgv As DataGridViewComboBoxCell
        Dim filter As String = ""

        Dim RowVarDummy As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_variable1Row
        Dim RowConstDummy As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_constraint1Row

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            DtaConst.Connection.ConnectionString = dta.Connection.ConnectionString
            DtaVar.Connection.ConnectionString = dta.Connection.ConnectionString

            vTblVar.Rows.Clear()
            vTblCons.Rows.Clear()
            dta.Fill(Tbl, vFrm.mPrimarySalary)
            DtaVar.FillBy_PrimarySalary(TblVar, vFrm.mPrimarySalary)
            DtaConst.FillBy_PrimarySalary(TblConst, vFrm.mPrimarySalary)
            vFrm.DgvComponent.Rows.Clear()
            For Each row In Tbl.Rows
                vFrm.DgvComponent.Rows.Add()
                rowDgv = vFrm.DgvComponent.Rows(vFrm.DgvComponent.RowCount - 1)
                CboDgv = rowDgv.Cells(vFrm.ColIncome_ComponentTipe.Index)
                rowDgv.Cells(vFrm.ColIncome_ComponentTipe.Index).Value = CboDgv.Items.Item(row.Type)
                rowDgv.Cells(vFrm.ColIncome_Desc.Index).Value = row.Description
                rowDgv.Cells(vFrm.ColIncome_SlipDescription.Index).Value = row.Slip_Desc_Code
                CboDgv = rowDgv.Cells(vFrm.ColIncome_TaxComponent.Index)
                rowDgv.Cells(vFrm.ColIncome_TaxComponent.Index).Value = CboDgv.Items.Item(row.Tax_Component)
                rowDgv.Cells(vFrm.ColIncome_ExpenseAccount.Index).Value = row.Expense_Account_Code
                rowDgv.Cells(vFrm.ColFormula.Index).Value = row.Formula
                CboDgv = rowDgv.Cells(vFrm.ColIncome_CalculatePeriod.Index)
                rowDgv.Cells(vFrm.ColIncome_CalculatePeriod.Index).Value = CboDgv.Items.Item(row.Calculation_Period)
                CboDgv = rowDgv.Cells(vFrm.ColIncome_CalculateOn.Index)
                rowDgv.Cells(vFrm.ColIncome_CalculateOn.Index).Value = CboDgv.Items.Item(row.Calculate_On)
                rowDgv.Cells(vFrm.ColIncome_PrimaryExpenseAccount.Index).Value = row.Primary_Expense_Account
                rowDgv.Cells(vFrm.ColIncome_PrimarySlipDesc.Index).Value = row.Primary_Slip_Desc
                rowDgv.Cells(vFrm.ColRoundingUnit.Index).Value = row.Rounding_Unit
                rowDgv.Cells(vFrm.ColRoundingType.Index).Value = row.Rounding_Type
                rowDgv.Cells(vFrm.ColYearlyMaxValue.Index).Value = row.Yearly_Max_Value
                rowDgv.Cells(vFrm.ColYearlyMinValue.Index).Value = row.Yearly_Min_Value
                rowDgv.Cells(vFrm.ColYearlyDate.Index).Value = row.Yearly_Date
                rowDgv.Cells(vFrm.ColMonthlyMaxValue.Index).Value = row.Monthly_Max_Value
                rowDgv.Cells(vFrm.ColMonthlyMinValue.Index).Value = row.Monthly_Min_Value
                rowDgv.Cells(vFrm.Colincome_Primary.Index).Value = -1
                rowDgv.Cells(vFrm.ColIncome_CreatedBy.Index).Value = vUserCode
                rowDgv.Cells(vFrm.ColIncome_ModifiedBy.Index).Value = vUserCode
                rowDgv.Cells(vFrm.ColIncome_CreatedDate.Index).Value = Format(Now, "dd-MMM-yyyy HH:mm")
                rowDgv.Cells(vFrm.ColIncome_ModifiedDate.Index).Value = Format(Now, "dd-MMM-yyyy HH:mm")
                If row.rutin_disetahunkan = 1 Then
                    rowDgv.Cells(vFrm.ColIncome_Routine.Index).Value = Boolean.TrueString
                Else
                    rowDgv.Cells(vFrm.ColIncome_Routine.Index).Value = Boolean.FalseString
                End If
                rowDgv.Cells(vFrm.ColIncome_PaymentAccount.Index).Value = row.Payment_Account_Code
                rowDgv.Cells(vFrm.ColIncome_PrimaryPaymentAccount.Index).Value = row.primary_payment_account

                'If vCopy = False Then
                '    rowDgv.Cells(vFrm.ColIncome_Urutan.Index).Value = row.Urutan
                '    rowDgv.Cells(vFrm.ColIncome_Identifier.Index).Value = row.Identifier
                'Else
                rowDgv.Cells(vFrm.ColIncome_Urutan.Index).Value = vFrm.DgvComponent.Rows.Count
                rowDgv.Cells(vFrm.ColIncome_Identifier.Index).Value = vPrimaryTrans.ToString + "_" + vFrm.DgvComponent.Rows.Count.ToString
                vFrm.MAXUrutan = vFrm.DgvComponent.Rows.Count
                'End If
                filter = "[Identifier Salary Component] ='" + row.Identifier.ToString + "'"
                TblVar.DefaultView.RowFilter = filter
                If TblVar.DefaultView.Count > 0 Then
                    Dim i As Integer = 1
                    For i = 1 To TblVar.DefaultView.Count
                        RowVar = TblVar.DefaultView(i - 1).Row
                        RowVarDummy = vTblVar.NewRow
                        RowVarDummy.Primary_Transaksi = vPrimaryTrans
                        RowVarDummy.Primary_Parameter = RowVar.Primary_Parameter
                        RowVarDummy.Type = RowVar.Type
                        RowVarDummy.Variable = RowVar.Variable
                        RowVarDummy.Identifier_Salary_Component = vPrimaryTrans.ToString + "_" + vFrm.DgvComponent.Rows.Count.ToString
                        RowVarDummy.Rounding_Unit = RowVar.Rounding_Unit
                        RowVarDummy.Rounding_Type = RowVar.Rounding_Type
                        RowVarDummy.Type2 = RowVar.Type2
                        RowVarDummy.Variable = RowVar.Variable
                        RowVarDummy.Primary = -1
                        vTblVar.Addtblmaster_salary_formula_salarycomponent_variable1Row(RowVarDummy)
                    Next
                End If
                TblConst.DefaultView.RowFilter = filter
                If TblConst.DefaultView.Count > 0 Then
                    Dim i As Integer = 1
                    For i = 1 To TblConst.DefaultView.Count
                        RowConst = TblConst.DefaultView(i - 1).Row
                        RowConstDummy = vTblCons.NewRow
                        RowConstDummy.Primary_Transaksi = vPrimaryTrans
                        RowConstDummy.Primary_Parameter = RowConst.Primary_Parameter
                        RowConstDummy.Type = RowConst.Type
                        RowConstDummy.Compare = RowConst.Compare
                        RowConstDummy.Identifier_Salary_Component = vPrimaryTrans.ToString + "_" + vFrm.DgvComponent.Rows.Count.ToString
                        RowConstDummy.Ref_Value = RowConst.Ref_Value
                        RowConstDummy.Type2 = RowConst.Type2
                        RowConstDummy.Primary = -1
                        vTblCons.Addtblmaster_salary_formula_salarycomponent_constraint1Row(RowConstDummy)
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            Tbl.Dispose()
            DtaConst.Dispose()
            TblConst.Dispose()
            DtaVar.Dispose()
            TblVar.Dispose()
        End Try
    End Sub

    Public Function GetDescComponent(ByVal vIdentifier As String) As String
        Dim DTA As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_salarycomponentTableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetDescComponent(vIdentifier)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetRefValueConstraint(ByVal vIdentifier_Formula As String, ByVal vPrimaryParameter As Double) As String
        Dim DTA As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_salarycomponent_constraintTableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetRefValue(vPrimaryParameter, vIdentifier_Formula)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "-1"
        Finally
            DTA.Dispose()
        End Try
    End Function

    'Rounding unit, dan rounding type yang diambil dari berdasarkan Salary Component Item 
    Public Function Get_SCI_RoundingType(ByVal vIdentifier As String) As Integer
        Dim DTA As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_salarycomponentTableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetRoundingType(vIdentifier)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Get_SCIRoundingUnit(ByVal vIdentifier As String) As Decimal
        Dim DTA As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_salarycomponentTableAdapter
        Dim Dummy As System.Nullable(Of Decimal)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetRoundingUnit(vIdentifier)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTipeConstrainVariable(ByVal vName As String) As Core_Module.ClsMasterSalaryFormula_item.EnumTipeConstraitVariable
        Try
            If vName = "Constraint List" Or vName = "Variable List" Then
                Return ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode
            ElseIf vName = "Competence" Then
                Return ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi
            ElseIf vName = "Custom Value" Then
                Return ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.CustomValue
            ElseIf vName = "Service Record" Then
                Return ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.ServiceRecord
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetTipeConstraint_Desc(ByVal vName As EnumTipeConstraitVariable) As String
        Try
            If vName = EnumTipeConstraitVariable.Hardcode Then
                Return "Constraint List"
            ElseIf vName = EnumTipeConstraitVariable.Kompetensi Then
                Return "Competence"
            ElseIf vName = EnumTipeConstraitVariable.CustomValue Then
                Return "Custom Value"
            ElseIf vName = EnumTipeConstraitVariable.ServiceRecord Then
                Return "Service Record"
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetTipeVariable_Desc(ByVal vName As EnumTipeConstraitVariable) As String
        Try
            If vName = EnumTipeConstraitVariable.Hardcode Then
                Return "Variable List"
            ElseIf vName = EnumTipeConstraitVariable.Kompetensi Then
                Return "Competence"
            ElseIf vName = EnumTipeConstraitVariable.CustomValue Then
                Return "Custom Value"
            ElseIf vName = EnumTipeConstraitVariable.ServiceRecord Then
                Return "Service Record"
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function TestFormula(ByVal Rumus As String, ByVal vDGVVariable As DataGridView, _
ByVal vFrm As Core_Module.FrmMasterSalaryFormula_Component2) As Boolean
        Dim Formula As New ExpressionContext
        Dim var As VariableCollection = Formula.Variables
        Dim Hasil As Decimal = 0
        Dim vDef_Prec As Decimal = 0
        Dim Dummy As Boolean = False
        Dim RowDGV As DataGridViewRow

        Try
            Rumus = Rumus.Replace(",", "")
            Formula.ParserOptions.DecimalSeparator = "."
            Formula.Options.ParseCulture.NumberFormat.NumberDecimalSeparator = "."
            Formula.Options.ParseCulture.NumberFormat.NumberGroupSeparator = ","
            Formula.ParserOptions.FunctionArgumentSeparator = ";"
            Formula.ParserOptions.RecreateParser()

            Dummy = vDGVVariable.AllowUserToAddRows
            vDGVVariable.AllowUserToAddRows = False

            For Each RowDGV In vDGVVariable.Rows
                var.Add(RowDGV.Cells(vFrm.ColVariable_Variable.Index).Value, 0)
            Next
            Hasil = CDec(Formula.CompileDynamic(Rumus).Evaluate)
            vDGVVariable.AllowUserToAddRows = Dummy

            Return True
        Catch ex As Exception
            MsgBox("Formula is incorrect, you cannot get result formula value.", MsgBoxStyle.Information, "Information")
            vDGVVariable.AllowUserToAddRows = Dummy
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Public Function TestFormula(ByVal Rumus As String, ByVal vDGVVariable As DataGridView, _
  ByVal vFrm As Core_Module.FrmMasterSalaryFormula_Component1) As Boolean
        Dim Formula As New ExpressionContext
        Dim var As VariableCollection = Formula.Variables
        Dim Hasil As Decimal = 0
        Dim vDef_Prec As Decimal = 0
        Dim Dummy As Boolean = False
        Dim RowDGV As DataGridViewRow

        Try
            Rumus = Rumus.Replace(",", "")
            Formula.ParserOptions.DecimalSeparator = "."
            Formula.Options.ParseCulture.NumberFormat.NumberDecimalSeparator = "."
            Formula.Options.ParseCulture.NumberFormat.NumberGroupSeparator = ","
            Formula.ParserOptions.FunctionArgumentSeparator = ";"
            Formula.ParserOptions.RecreateParser()

            Dummy = vDGVVariable.AllowUserToAddRows
            vDGVVariable.AllowUserToAddRows = False

            For Each RowDGV In vDGVVariable.Rows
                var.Add(RowDGV.Cells(vFrm.ColVariable_Variable.Index).Value, 0)
            Next
            Hasil = CDec(Formula.CompileDynamic(Rumus).Evaluate)
            vDGVVariable.AllowUserToAddRows = Dummy

            Return True
        Catch ex As Exception
            MsgBox("Formula is incorrect, you cannot get result formula value.", MsgBoxStyle.Information, "Information")
            vDGVVariable.AllowUserToAddRows = Dummy
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Public Function GetIdentifierFormula(ByVal vComponentName As String) As String
        Dim DTA As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_salarycomponentTableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetIdentifier(vComponentName)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return "-1-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "-1-1"
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryFormula_Identifier(ByVal vIdentifier As String) As Double
        Dim DTA As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_salarycomponentTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimaryTransaksiByIdentifier(vIdentifier)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryFormula(ByVal vComponentName As String) As Double
        Dim DTA As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_salarycomponentTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimaryTransaksi(vComponentName)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
