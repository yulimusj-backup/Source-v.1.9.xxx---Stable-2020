Option Explicit On
Option Strict Off

Imports System.Windows.Forms, System.Drawing
Public Class FrmMasterSalaryFormula_Print
    Dim ObjCuti As New ClsMasterCuti
    Dim ObjJP As New ClsMasterJobPosition
    Dim ObjKompetensei As New ClsMasterKompetensi

    Public Enum EnumReportType As Integer
        EMPLOYEE_LIST
        SALARY_FORMULA
        SALARY_PARAMS
    End Enum

    Public mReportType As FrmMasterSalaryFormula_Print.EnumReportType = EnumReportType.EMPLOYEE_LIST
    Public mPrimary As Double = -1
    Public mCode As String = ""
    Dim ObjCd As New Core_Module.ClsCekData
    Dim ObjCurrency As New Core_Module.clsMasterCurrency
    Dim DefPecIDR As Integer = 0
    Dim RptEmployeeList As New CrpCetakSalaryFormula_EmployeeList
    Dim RptSalaryFormula As New CrpCetakSalaryFormula_Formula
    Dim RptSalaryParams As New CrpCetakSalaryFormula_Parameter

    Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterSalaryFormula_Print_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.DefPecIDR = Me.ObjCurrency.GetDecimal_DefaultCurrency
            If Me.mReportType = EnumReportType.EMPLOYEE_LIST Then
                Me.Text = "Print Employee List"
                PrintEmployeeList(mPrimary, mCode)
            ElseIf Me.mReportType = EnumReportType.SALARY_FORMULA Then
                Me.Text = "Print Salary Formula"
                PrintSalaryFormula(Me.mPrimary)
            ElseIf Me.mReportType = EnumReportType.SALARY_PARAMS Then
                Me.Text = "Print Salary Parameters"
                PrintSalaryParams()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Function PrintEmployeeList(ByVal vPrimaryMasterSalary As Double, ByVal vCode As String) As Boolean
        Dim DsCetak As New DseMasterSalaryFormula_Print
        Dim Dta As New DseMasterSalaryFormula_PrintTableAdapters.Tbl_Cetak_EmployeeListTableAdapter
        Dim Tbl As DseMasterSalaryFormula_Print.Tbl_Cetak_EmployeeListDataTable = DsCetak.Tbl_Cetak_EmployeeList
        Try
            Me.Cursor = Cursors.WaitCursor
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            If Dta.FillBy_PrimarySalary(Tbl, vPrimaryMasterSalary) > 0 Then
                For Each Row As DseMasterSalaryFormula_Print.Tbl_Cetak_EmployeeListRow In Tbl
                    Row.GroupingPayDanExpAccbyCode = Row.Pay_AccCode & "-" & Row.Exp_AccCode
                Next
                Me.RptEmployeeList.SetDataSource(DsCetak)
                Me.RptEmployeeList.SetParameterValue(RptEmployeeList.Parameter_P_DefPrecIDR.ParameterFieldName, Me.DefPecIDR)
                Me.RptEmployeeList.SetParameterValue(RptEmployeeList.Parameter_P_SalaryFormulaCode.ParameterFieldName, Me.mCode)
                Me.crvReport.ReportSource = Me.RptEmployeeList
                Me.crvReport.Zoom(90)
                Me.crvReport.Refresh()
                Me.crvReport.Show()
            Else
                MsgBox("There is no data.", MsgBoxStyle.Information, "Information")
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Function

    Private Function PrintSalaryFormula(ByVal vPrimarySalaryFormula As Double) As Boolean
        Dim DsCetak As New DseMasterSalaryFormula_Print
        Dim DtaHeader As New DseMasterSalaryFormula_PrintTableAdapters.tblmaster_salary_formula_salarycomponentTableAdapter
        Dim TblHeader As New DseMasterSalaryFormula_Print.tblmaster_salary_formula_salarycomponentDataTable
        Dim RowHeader As DseMasterSalaryFormula_Print.tblmaster_salary_formula_salarycomponentRow
        Dim DtaVar As New DseMasterSalaryFormula_PrintTableAdapters.tblmaster_salary_formula_salarycomponent_variableTableAdapter
        Dim TblVar As New DseMasterSalaryFormula_Print.tblmaster_salary_formula_salarycomponent_variableDataTable
        Dim RowVar As DseMasterSalaryFormula_Print.tblmaster_salary_formula_salarycomponent_variableRow
        Dim DtaCon As New DseMasterSalaryFormula_PrintTableAdapters.tblmaster_salary_formula_salarycomponent_constraintTableAdapter
        Dim TblCon As New DseMasterSalaryFormula_Print.tblmaster_salary_formula_salarycomponent_constraintDataTable
        Dim RowCon As DseMasterSalaryFormula_Print.tblmaster_salary_formula_salarycomponent_constraintRow
        Dim TblCetak As DseMasterSalaryFormula_Print.TblCetakDataTable = DsCetak.TblCetak
        Dim RowCetak As DseMasterSalaryFormula_Print.TblCetakRow
        Dim Param_Header0 As String
        Dim Param_HeaderType_CalcPeriod As String
        Dim Param_HeaderRounding As String
        Dim Param_HeaderMonthlyMin As String
        Dim Param_HeaderMonthlyMax As String
        Dim Param_HeaderStartOn_CalcOn As String
        Dim Param_HeaderYearlyMin As String
        Dim Param_HeaderYearlyMax As String
        Dim Formula As String
        Try
            Me.Cursor = Cursors.WaitCursor
            DtaHeader.Connection.ConnectionString = Core_Module.GetConstring
            DtaVar.Connection.ConnectionString = DtaHeader.Connection.ConnectionString
            DtaCon.Connection.ConnectionString = DtaHeader.Connection.ConnectionString
            If DtaHeader.FillBy_PrimarySalary(TblHeader, vPrimarySalaryFormula) > 0 Then
                DtaVar.FillBy_PrimarySalary(TblVar, vPrimarySalaryFormula)
                DtaCon.FillBy_PrimarySalary(TblCon, vPrimarySalaryFormula)
                For Each RowHeader In TblHeader
                    Param_Header0 = RowHeader.Description & ", formula = " & RowHeader.Formula
                    Param_HeaderType_CalcPeriod = Me.GetPeriod(RowHeader.Calculation_Period)
                    Param_HeaderRounding = Me.GetRoundingType(RowHeader.Rounding_Type)
                    If RowHeader.Rounding_Type <> ClsMasterSalaryFormula_item.EnumRoundingType.No_Rounding Then
                        Param_HeaderRounding += ", unit: " & Me.ObjCd.BuangNolDibelakangDecimal(RowHeader.Rounding_Unit)
                    End If
                    Param_HeaderMonthlyMin = RowHeader.Monthly_Min_Value
                    If IsNumeric(Param_HeaderMonthlyMin) = True Then
                        If CDec(Param_HeaderMonthlyMin) = 0 Then
                            Param_HeaderMonthlyMin = "-"
                        End If
                    End If
                    Param_HeaderMonthlyMax = RowHeader.Monthly_Max_Value
                    If IsNumeric(Param_HeaderMonthlyMax) = True Then
                        If CDec(Param_HeaderMonthlyMax) = 0 Then
                            Param_HeaderMonthlyMax = "-"
                        End If
                    End If
                    Param_HeaderStartOn_CalcOn = Me.GetMonthStr(RowHeader.Calculate_On)
                    Param_HeaderYearlyMin = RowHeader.Yearly_Min_Value
                    If IsNumeric(Param_HeaderYearlyMin) = True Then
                        If CDec(Param_HeaderYearlyMin) = 0 Then
                            Param_HeaderYearlyMin = "-"
                        End If
                    End If
                    Param_HeaderYearlyMax = RowHeader.Yearly_Max_Value
                    If IsNumeric(Param_HeaderYearlyMax) = True Then
                        If CDec(Param_HeaderYearlyMax) = 0 Then
                            Param_HeaderYearlyMax = "-"
                        End If
                    End If
                    Formula = RowHeader.Formula
                    TblVar.DefaultView.RowFilter = "[Identifier Salary Component] = '" & RowHeader.Identifier & "'"
                    Dim noItem As Boolean = True
                    For Each DRV As DataRowView In TblVar.DefaultView
                        noItem = False
                        RowVar = DRV.Row
                        RowCetak = TblCetak.NewTblCetakRow
                        RowCetak.Primary_Auto_Component = RowHeader.Primary
                        RowCetak.Description = RowHeader.Description
                        RowCetak.Var_Const = "A"
                        RowCetak.Variable = RowVar.Variable
                        If RowVar.Type2 = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode Then
                            RowCetak.Parameter = RowVar.Param_Code
                            If RowVar.Primary_Parameter = 131 Then
                                'Khusus Parameter Job Position 
                                RowCetak.RUnit_RefValue = Me.ObjJP.GetCode(RowVar.Rounding_Unit)
                            ElseIf RowVar.Primary_Parameter = 133 Then
                                'Khusus Parameter Leave Code 
                                RowCetak.RUnit_RefValue = Me.ObjCuti.GetCodeByPrimary(RowVar.Rounding_Unit)
                            Else
                                If IsNumeric(RowVar.Rounding_Unit) = True Then
                                    RowCetak.RUnit_RefValue = Me.GetRoundingType(RowVar.Rounding_Type) & ", unit: " & Me.ObjCd.BuangNolDibelakangDecimal(RowVar.Rounding_Unit)
                                Else
                                    RowCetak.RUnit_RefValue = Me.GetRoundingType(RowVar.Rounding_Type) & ", unit: " & RowVar.Rounding_Unit
                                End If
                                If RowVar.Rounding_Type = ClsMasterSalaryFormula_item.EnumRoundingType.No_Rounding Then
                                    RowCetak.RUnit_RefValue = "-"
                                End If

                            End If
                        Else
                            RowCetak.Parameter = Me.ObjKompetensei.GetDesc(RowVar.Primary_Parameter)
                            RowCetak.RUnit_RefValue = Me.ObjKompetensei.SetFormatDataCustom(RowVar.Rounding_Unit, RowVar.Primary_Parameter, Me.ObjCd)
                            If IsNumeric(RowVar.Rounding_Unit) = True Then
                                RowCetak.RUnit_RefValue = Me.GetRoundingType(RowVar.Rounding_Type) & ", unit: " & Me.ObjCd.BuangNolDibelakangDecimal(RowVar.Rounding_Unit)
                            Else
                                RowCetak.RUnit_RefValue = Me.GetRoundingType(RowVar.Rounding_Type) & ", unit: " & RowVar.Rounding_Unit
                            End If
                            If RowVar.Rounding_Type = ClsMasterSalaryFormula_item.EnumRoundingType.No_Rounding Then
                                RowCetak.RUnit_RefValue = "-"
                            End If
                        End If

                        RowCetak.Param_Header0 = Param_Header0
                        RowCetak.Param_HeaderType_CalcPeriod = Param_HeaderType_CalcPeriod
                        RowCetak.Param_HeaderRounding = Param_HeaderRounding
                        RowCetak.Param_HeaderMonthlyMin = Param_HeaderMonthlyMin
                        RowCetak.Param_HeaderMonthlyMax = Param_HeaderMonthlyMax
                        RowCetak.Param_HeaderStartOn_CalcOn = Param_HeaderStartOn_CalcOn
                        RowCetak.Param_HeaderYearlyMin = Param_HeaderYearlyMin
                        RowCetak.Param_HeaderYearlyMax = Param_HeaderYearlyMax
                        RowCetak.Formula = Formula
                        TblCetak.Rows.Add(RowCetak)
                    Next
                    TblCon.DefaultView.RowFilter = "[Identifier Salary Component] = '" & RowHeader.Identifier & "'"
                    For Each DRV As DataRowView In TblCon.DefaultView
                        noItem = False
                        RowCon = DRV.Row
                        RowCetak = TblCetak.NewTblCetakRow
                        RowCetak.Primary_Auto_Component = RowHeader.Primary
                        RowCetak.Description = RowHeader.Description
                        RowCetak.Var_Const = "B"
                        RowCetak.Parameter = RowCon.Param_Code
                        RowCetak.Variable = Me.GetComparisonSymbol(RowCon.Compare)
                        If RowCon.Type2 = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode Then
                            RowCetak.Parameter = RowCon.Param_Code
                            If RowCon.Primary_Parameter = 131 Then
                                'Khusus Parameter Job Position 
                                RowCetak.RUnit_RefValue = Me.ObjJP.GetCode(RowCon.Ref_Value) & " - " & Me.ObjJP.GetDesc(Me.ObjJP.GetPrimary(Me.ObjJP.GetCode(RowCon.Ref_Value)))
                            ElseIf RowCon.Primary_Parameter = 133 Then
                                'Khusus Parameter Leave Code 
                                RowCetak.RUnit_RefValue = Me.ObjCuti.GetCodeByPrimary(RowCon.Ref_Value)
                            Else
                                If IsNumeric(RowCon.Ref_Value) = True Then
                                    RowCetak.RUnit_RefValue = Me.ObjCd.BuangNolDibelakangDecimal(RowCon.Ref_Value)
                                Else
                                    RowCetak.RUnit_RefValue = RowCon.Ref_Value
                                End If

                                ''@rosadinnn
                                If Me.GetUnit(RowCon.Primary_Parameter).Trim.ToLower = "days" Then
                                    RowCetak.RUnit_RefValue = Me.ObjCd.BuangNolDibelakangDecimal(RowCon.Ref_Value) & " days."
                                ElseIf Me.GetUnit(RowCon.Primary_Parameter).Trim.ToLower = "duration" Or Me.GetUnit(RowCon.Primary_Parameter).Trim.ToLower = "time" Then
                                    RowCetak.RUnit_RefValue = RowCon.Ref_Value
                                ElseIf Me.GetUnit(RowCon.Primary_Parameter).Trim.ToLower = "numeric" Then
                                    RowCetak.RUnit_RefValue = Me.ObjCd.BuangNolDibelakangDecimal(RowCon.Ref_Value)
                                End If

                            End If
                        Else
                            RowCetak.Parameter = Me.ObjKompetensei.GetDesc(RowCon.Primary_Parameter)
                            RowCetak.RUnit_RefValue = Me.ObjKompetensei.SetFormatDataCustom(RowCon.Ref_Value, RowCon.Primary_Parameter, Me.ObjCd)
                        End If

                        RowCetak.Param_Header0 = Param_Header0
                        RowCetak.Param_HeaderType_CalcPeriod = Param_HeaderType_CalcPeriod
                        RowCetak.Param_HeaderRounding = Param_HeaderRounding
                        RowCetak.Param_HeaderMonthlyMin = Param_HeaderMonthlyMin
                        RowCetak.Param_HeaderMonthlyMax = Param_HeaderMonthlyMax
                        RowCetak.Param_HeaderStartOn_CalcOn = Param_HeaderStartOn_CalcOn
                        RowCetak.Param_HeaderYearlyMin = Param_HeaderYearlyMin
                        RowCetak.Param_HeaderYearlyMax = Param_HeaderYearlyMax
                        RowCetak.Formula = Formula
                        TblCetak.Rows.Add(RowCetak)
                    Next
                    If noItem = True Then
                        RowCetak = TblCetak.NewTblCetakRow
                        RowCetak.Primary_Auto_Component = RowHeader.Primary
                        RowCetak.Description = RowHeader.Description
                        RowCetak.Var_Const = "$$$DO NOT CHANGE THIS CHARS$$$"
                        'RowCetak.Parameter = RowCon.Param_Code
                        'RowCetak.Variable = Me.GetComparisonSymbol(RowCon.Compare)
                        'RowCetak.Parameter = Me.ObjKompetensei.GetDesc(RowCon.Primary_Parameter)
                        'RowCetak.RUnit_RefValue = Me.ObjKompetensei.SetFormatDataCustom(RowCon.Ref_Value, RowCon.Primary_Parameter, Me.ObjCd)
                        RowCetak.Param_Header0 = Param_Header0
                        RowCetak.Param_HeaderType_CalcPeriod = Param_HeaderType_CalcPeriod
                        RowCetak.Param_HeaderRounding = Param_HeaderRounding
                        RowCetak.Param_HeaderMonthlyMin = Param_HeaderMonthlyMin
                        RowCetak.Param_HeaderMonthlyMax = Param_HeaderMonthlyMax
                        RowCetak.Param_HeaderStartOn_CalcOn = Param_HeaderStartOn_CalcOn
                        RowCetak.Param_HeaderYearlyMin = Param_HeaderYearlyMin
                        RowCetak.Param_HeaderYearlyMax = Param_HeaderYearlyMax
                        RowCetak.Formula = Formula
                        TblCetak.Rows.Add(RowCetak)
                    End If
                Next
                Me.RptSalaryFormula.SetDataSource(DsCetak)
                Me.RptSalaryFormula.SetParameterValue(Me.RptSalaryFormula.Parameter_P_SalaryCode.ParameterFieldName, Me.mCode)
                Me.crvReport.ReportSource = Me.RptSalaryFormula
                Me.crvReport.Zoom(90)
                Me.crvReport.Refresh()
                Me.crvReport.Show()
            Else
                MsgBox("There is no data.", MsgBoxStyle.Information, "Information")
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Function

    Private Function GetUnit(ByVal vPrimaryParameter As Double) As String
        Dim dta As New DseMasterSalaryFormula_PrintTableAdapters.tblmaster_salary_formula_parameterTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Return dta.ScalarQuery_GetTipeParameterByPrimary(vPrimaryParameter)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function


    Private Function PrintSalaryParams() As Boolean
        ''@aul
        ''@201506250900
        Dim DsCetak As New DseMasterSalaryFormula_Print
        Dim Dta As New DseMasterSalaryFormula_PrintTableAdapters.tblmaster_salary_formula_parameterTableAdapter
        Dim Tbl As DseMasterSalaryFormula_Print.tblmaster_salary_formula_parameterDataTable = DsCetak.tblmaster_salary_formula_parameter
        Try
            Me.Cursor = Cursors.WaitCursor
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            If Dta.Fill(Tbl) > 0 Then
                For Each Row As DseMasterSalaryFormula_Print.tblmaster_salary_formula_parameterRow In Tbl
                    If Row.Is_Variable = 1 Then
                        Row.VarConst = "V"
                    End If
                    If Row.Is_Constraint = 1 Then
                        If String.IsNullOrEmpty(Row.VarConst) = True Then
                            Row.VarConst = "C "
                        Else
                            Row.VarConst += "/C "
                        End If
                    End If
                Next
                Me.RptSalaryParams.SetDataSource(DsCetak)
                Me.crvReport.ReportSource = Me.RptSalaryParams
                Me.crvReport.Zoom(90)
                Me.crvReport.Refresh()
                Me.crvReport.Show()
            Else
                MsgBox("There is no data.", MsgBoxStyle.Information, "Information")
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Function

    Private Function GetPeriod(ByVal vIdx As Integer) As String
        Try
            Select Case vIdx
                Case 0
                    Return "daily"
                Case 1
                    Return "monthly"
                Case 2
                    Return "yearly"
            End Select
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally

        End Try
    End Function

    Private Function GetComparisonSymbol(ByVal vIdx As Integer) As String
        Try
            Select Case vIdx
                Case ClsMasterParameterGaji.EnumComparer.SamaDengan
                    Return "="
                Case ClsMasterParameterGaji.EnumComparer.LebihBesar
                    Return ">"
                Case ClsMasterParameterGaji.EnumComparer.LebihKecil
                    Return "<"
                Case ClsMasterParameterGaji.EnumComparer.LebihBesar_SamaDengan
                    Return ">="
                Case ClsMasterParameterGaji.EnumComparer.LebihKecil_SamaDengan
                    Return "<="
                Case ClsMasterParameterGaji.EnumComparer.TidakSamaDengan
                    Return "<>"
            End Select
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally

        End Try
    End Function

    Private Function GetMonthStr(ByVal vIdx As Integer) As String
        Try
            Select Case vIdx
                Case 0
                    Return "January"
                Case 1
                    Return "February"
                Case 2
                    Return "March"
                Case 3
                    Return "April"
                Case 4
                    Return "May"
                Case 5
                    Return "June"
                Case 6
                    Return "July"
                Case 7
                    Return "August"
                Case 8
                    Return "September"
                Case 9
                    Return "October"
                Case 10
                    Return "November"
                Case 11
                    Return "December"
            End Select
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally

        End Try
    End Function

    Private Function GetRoundingType(ByVal vIdx As Integer) As String
        Try
            Select Case vIdx
                Case ClsMasterSalaryFormula_item.EnumRoundingType.No_Rounding
                    Return "No Rounding"
                Case ClsMasterSalaryFormula_item.EnumRoundingType.Round_Down
                    Return "Round Down"
                Case ClsMasterSalaryFormula_item.EnumRoundingType.Round_Middlepoint
                    Return "Middlepoint"
                Case ClsMasterSalaryFormula_item.EnumRoundingType.Round_Up
                    Return "Round Up"
            End Select
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally

        End Try
    End Function
End Class