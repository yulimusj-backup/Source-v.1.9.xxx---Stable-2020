Option Explicit On
Option Strict Off

Imports Core_Module.ClsCekData

Public Class ClsTrans_OtherDeduction_item
    Dim ObjParameter As New ClsMasterParameterGaji

    Public Enum EnumRoundingType As Integer
        No_Rounding
        Round_Down
        Round_Middlepoint
        Round_Up
    End Enum

    Public Sub IsiDgvVar(ByVal vFrm As FrmMasterSalaryFormula_Component2, ByVal vPrimaryTransaksi As Double, _
    ByVal vTbl As DseTrans_OtherDeduction.tblmaster_salary_formula_otherdeduction_variableDataTable)
        Dim row As DseTrans_OtherDeduction.tblmaster_salary_formula_otherdeduction_variableRow
        Dim RowDgv As DataGridViewRow
        Dim CboDgv As DataGridViewComboBoxCell
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            vFrm.DGVVariable.Rows.Clear()
            For Each row In vTbl.Rows
                vFrm.DGVVariable.Rows.Add()
                RowDgv = vFrm.DGVVariable.Rows(vFrm.DGVVariable.Rows.Count - 2)
                RowDgv.Cells(vFrm.ColVariable_No.Index).Value = vFrm.DGVVariable.Rows.Count - 1
                RowDgv.Cells(vFrm.ColVariable_PrimaryParameter.Index).Value = row.Primary_Parameter
                RowDgv.Cells(vFrm.ColVariable_Parameter_Code.Index).Value = Me.ObjParameter.GetCodeByPrimary(row.Primary_Parameter)
                RowDgv.Cells(vFrm.ColVariable_Parameter_Desc.Index).Value = Me.ObjParameter.GetDescByPrimary(row.Primary_Parameter)
                CboDgv = RowDgv.Cells(vFrm.ColVariable_Type.Index)
                RowDgv.Cells(vFrm.ColVariable_Type.Index).Value = CboDgv.Items.Item(row.Type)
                RowDgv.Cells(vFrm.ColVariable_Unit.Index).Value = Me.ObjParameter.GetUnitByPrimary(row.Primary_Parameter)
                RowDgv.Cells(vFrm.ColVariable_Variable.Index).Value = row.Variable
                CboDgv = RowDgv.Cells(vFrm.ColRoundingTipe_Variable.Index)
                RowDgv.Cells(vFrm.ColRoundingTipe_Variable.Index).Value = CboDgv.Items.Item(row.Rounding_Type)
                RowDgv.Cells(vFrm.ColRoundingUnit_Variable.Index).Value = ObjCD.BuangNolDibelakangDecimal(ObjCD.DigitGrouping(row.Rounding_Unit))
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
        End Try
    End Sub

    Public Sub IsiDgvConstraint(ByVal vFrm As FrmMasterSalaryFormula_Component2, ByVal vPrimaryTransaksi As Double, _
    ByVal vTbl As DseTrans_OtherDeduction.tblmaster_salary_formula_otherdeduction_constraintDataTable)
        Dim row As DseTrans_OtherDeduction.tblmaster_salary_formula_otherdeduction_constraintRow
        Dim RowDgv As DataGridViewRow
        Dim CboDgv As DataGridViewComboBoxCell
        Dim ObjJP As New Core_Module.ClsMasterJobPosition
        Dim ObjCuti As New Core_Module.ClsMasterCuti
        Try
            vFrm.DGVConstaint.Rows.Clear()
            For Each row In vTbl.Rows
                vFrm.DGVConstaint.Rows.Add()
                RowDgv = vFrm.DGVConstaint.Rows(vFrm.DGVConstaint.Rows.Count - 2)
                RowDgv.Cells(vFrm.ColConts_No.Index).Value = vFrm.DGVConstaint.Rows.Count - 1
                RowDgv.Cells(vFrm.ColConst_PrimaryParameter.Index).Value = row.Primary_Parameter
                RowDgv.Cells(vFrm.ColConst_ParameterCode.Index).Value = Me.ObjParameter.GetCodeByPrimary(row.Primary_Parameter)
                RowDgv.Cells(vFrm.ColConst_ParameterDesc.Index).Value = Me.ObjParameter.GetDescByPrimary(row.Primary_Parameter)
                CboDgv = RowDgv.Cells(vFrm.ColConst_Type.Index)
                RowDgv.Cells(vFrm.ColConst_Type.Index).Value = CboDgv.Items.Item(row.Type)
                CboDgv = RowDgv.Cells(vFrm.ColConst_Compare.Index)
                RowDgv.Cells(vFrm.ColConst_Compare.Index).Value = CboDgv.Items.Item(row.Compare)
                'Ref Value
                If row.Primary_Parameter = 131 Then
                    'Khusus Parameter Job Position 
                    RowDgv.Cells(vFrm.ColPrimaryRefValue.Index).Value = row.Ref_Value
                    RowDgv.Cells(vFrm.ColConst_RefValue.Index).Value = ObjJP.GetCode(row.Ref_Value).ToString
                ElseIf row.Primary_Parameter = 133 Then
                    'Khusus Paramenter Leave Code
                    RowDgv.Cells(vFrm.ColPrimaryRefValue.Index).Value = row.Ref_Value
                    RowDgv.Cells(vFrm.ColConst_RefValue.Index).Value = ObjCuti.GetCodeByPrimary(row.Ref_Value).ToString
                Else
                    RowDgv.Cells(vFrm.ColConst_RefValue.Index).Value = row.Ref_Value
                End If
                RowDgv.Cells(vFrm.ColConst_Unit.Index).Value = Me.ObjParameter.GetUnitByPrimary(row.Primary_Parameter)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjJP = Nothing
            ObjCuti = Nothing
        End Try
    End Sub

    Public Sub IsiTableVar(ByVal vTbl As DseTrans_OtherDeduction.tblmaster_salary_formula_otherdeduction_variableDataTable, _
    ByVal vPrimaryTrans As Double)
        Dim Dta As New DseTrans_OtherDeductionTableAdapters.tblmaster_salary_formula_otherdeduction_variableTableAdapter
        Try
            Dta.FillBy_PrimaryTransaksi(vTbl, vPrimaryTrans)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Sub

    Public Sub IsiTableConts(ByVal vTbl As DseTrans_OtherDeduction.tblmaster_salary_formula_otherdeduction_constraintDataTable, _
    ByVal vPrimaryTrans As Double)
        Dim Dta As New DseTrans_OtherDeductionTableAdapters.tblmaster_salary_formula_otherdeduction_constraintTableAdapter
        Try
            Dta.FillBy_PrimaryTransaksi(vTbl, vPrimaryTrans)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Sub

    Public Sub IsiDGVItem( ByVal Frm As FrmMasterSalaryFormula_OtherDeduction, ByVal vPrimaryTransaksi As Double)
        Dim DTA As New DseTrans_OtherDeductionTableAdapters.tblmaster_salary_formula_otherdeduction_employee1TableAdapter
        Dim TBL As New DseTrans_OtherDeduction.tblmaster_salary_formula_otherdeduction_employee1DataTable
        Dim row As DseTrans_OtherDeduction.tblmaster_salary_formula_otherdeduction_employee1Row
        Dim rowDGV As DataGridViewRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'kosongkan dgv
            Frm.DgvItem.AllowUserToAddRows = True
            Frm.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Frm.DgvItem.Rows.Clear()

            ''isikan(tabel)
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        Frm.DgvItem.Rows.Add("")
                        rowDGV = Frm.DgvItem.Rows(Frm.DgvItem.Rows.Count - 2)
                        rowDGV.Cells(Frm.ColNo.Index).Value = Frm.DgvItem.Rows.Count - 1
                        rowDGV.Cells(Frm.ColPrimaryEmployee.Index).Value = .primary_employee
                        rowDGV.Cells(Frm.ColEmployeeCode.Index).Value = .Employee_Code
                        rowDGV.Cells(Frm.ColEmployeeName.Index).Value = .Employee_Name
                        rowDGV.Cells(Frm.ColNIK.Index).Value = .NIK
                        rowDGV.Cells(Frm.ColDept.Index).Value = .Dept_Desc
                        rowDGV.Cells(Frm.ColPosition.Index).Value = .Position_Desc
                        rowDGV.Cells(Frm.ColRemarks.Index).Value = .remarks
                    End With
                Next
            End If

            'set nomer
            Frm.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Frm.DgvItem.AllowUserToAddRows = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
        End Try
    End Sub

    Public Sub simpan(ByVal vFrm As FrmMasterSalaryFormula_OtherDeduction, ByVal vPrimaryTransaksi As Double)
        Dim RowDgv As DataGridViewRow
        Dim remark As String
        Dim DtaEmp As New DseTrans_OtherDeductionTableAdapters.tblmaster_salary_formula_otherdeduction_employeeTableAdapter
        Dim DtaVar As New DseTrans_OtherDeductionTableAdapters.tblmaster_salary_formula_otherdeduction_variableTableAdapter
        Dim DtaCon As New DseTrans_OtherDeductionTableAdapters.tblmaster_salary_formula_otherdeduction_constraintTableAdapter
        Dim RowVar As DseTrans_OtherDeduction.tblmaster_salary_formula_otherdeduction_variableRow
        Dim RowCon As DseTrans_OtherDeduction.tblmaster_salary_formula_otherdeduction_constraintRow
        Try
            'simpan item emp
            For Each RowDgv In vFrm.DgvItem.Rows
                If RowDgv.Index <> vFrm.DgvItem.RowCount - 1 Then
                    If String.IsNullOrEmpty(RowDgv.Cells(vFrm.ColRemarks.Index).Value) = True Then
                        remark = ""
                    Else
                        remark = RowDgv.Cells(vFrm.ColRemarks.Index).Value
                    End If
                    DtaEmp.Insert(vPrimaryTransaksi, RowDgv.Cells(vFrm.ColPrimaryEmployee.Index).Value, remark)
                End If
            Next
            'simpan item var
            For Each RowVar In vFrm.TblVar
                With RowVar
                    DtaVar.Insert(vPrimaryTransaksi, .Primary_Parameter, .Type, .Variable, .Rounding_Unit, .Rounding_Type)
                End With
            Next
            'simpan item con
            For Each RowCon In vFrm.TblCons
                With RowCon
                    DtaCon.Insert(vPrimaryTransaksi, .Primary_Parameter, .Type, .Compare, .Ref_Value)
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaEmp.Dispose()
            DtaVar.Dispose()
            DtaCon.Dispose()
        End Try
    End Sub

    Public Sub hapus(ByVal vPrimaryTransaksi As Double)
        Dim DtaEmp As New DseTrans_OtherDeductionTableAdapters.tblmaster_salary_formula_otherdeduction_employeeTableAdapter
        Dim DtaVar As New DseTrans_OtherDeductionTableAdapters.tblmaster_salary_formula_otherdeduction_variableTableAdapter
        Dim DtaCon As New DseTrans_OtherDeductionTableAdapters.tblmaster_salary_formula_otherdeduction_constraintTableAdapter
        Try
            DtaEmp.Delete(vPrimaryTransaksi)
            DtaVar.Delete(vPrimaryTransaksi)
            DtaCon.Delete(vPrimaryTransaksi)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaEmp.Dispose()
            DtaVar.Dispose()
            DtaCon.Dispose()
        End Try
    End Sub

    Public Function GetRefValueConstraint_Others(ByVal vPrimaryTransaksi As Double, ByVal vPrimaryParameter As Double) As String
        Dim DTA As New DseTrans_OtherDeductionTableAdapters.tblmaster_salary_formula_otherdeduction_constraintTableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetRefValue_Others(vPrimaryTransaksi, vPrimaryParameter)
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
    Public Function Get_SCI_RoundingType(ByVal vPrimaryOther As Double) As Integer
        Dim DTA As New DseTrans_OtherDeductionTableAdapters.tblmaster_salary_formula_otherdeductionTableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetRoundingType(vPrimaryOther)
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

    Public Function Get_SCIRoundingUnit(ByVal vPrimaryOther As String) As Decimal
        Dim DTA As New DseTrans_OtherDeductionTableAdapters.tblmaster_salary_formula_otherdeductionTableAdapter
        Dim Dummy As System.Nullable(Of Decimal)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetRoundingUnit(vPrimaryOther)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class
