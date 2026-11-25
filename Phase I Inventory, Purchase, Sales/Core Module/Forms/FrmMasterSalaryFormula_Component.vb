Option Explicit On
Option Strict Off

Imports System.Threading
Public Class FrmMasterSalaryFormula_Component
    Public BackColorAll As Color = Color.AliceBlue
    Public mPrimarySalary As Double = -1
    Public mUserCode As String = ""
    Public LocationFileSumber As String = ""
    Dim mClose As Boolean = False
    Dim Buffer As String = ""
    Public MAXUrutan As Integer = 0
    Dim ObjCurr As New clsMasterCurrency
    Dim ObjCd As New ClsCekData
    Dim ObjAcc As New ClsMasterAkun1
    Public ObjSalaryFormula As New ClsMasterSalaryFormula

    Dim ObjSlipSalary As New ClsMasterSalaryFormula_SlipDesc
    Dim ObjEmp As New ClsMasterKaryawan
    Dim ObjParameter As New ClsMasterParameterGaji
    Dim ObjCompetence As New Core_Module.ClsMasterKompetensi
    Dim ObjItem As New ClsMasterSalaryFormula_item
    Dim ObjJP As New ClsMasterJobPosition
    Dim ObjCuti As New ClsMasterCuti
    Dim EraseRowIndex As Integer = 0
    Dim mDGV As DataGridView
    Dim DefCurr As Integer = Me.ObjCurr.GetDecimal_DefaultCurrency
    Dim UserLain, NamaKomputerLain As String
    Public dtDGV As New DataTable
    Public SedangIsiDGV As Boolean = False

    'Dim ListPrimaryDeleted As New List(Of Double)
    'Public ListPrimaryDeletedVar As New List(Of Double)
    'Public ListPrimaryDeletedConts As New List(Of Double)


    Dim ObjAntrian As New ClsAntrian

    Public TblVar As New DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_variable1DataTable
    Public TblCons As New DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_constraint1DataTable
    Dim RowVar As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_variable1Row
    Dim RowCons As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_constraint1Row

    Public TblVarOld As New DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_variable1DataTable
    Public TblConsOld As New DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_constraint1DataTable
    Dim RowVarOld As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_variable1Row
    Dim RowConsOld As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_constraint1Row

    Public tblSC As New DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent1DataTable
    Public tblEMP As New DseMasterSalaryFormula.tblmaster_salary_formula_component_detailemployee1DataTable

    Public tblLogUser As New DseMasterSalaryFormula.tbl_log_userDataTable
    Dim rowLogUser As DseMasterSalaryFormula.tbl_log_userRow

    Dim _ok As Boolean = False
    Dim e As System.Windows.Forms.DataGridViewCellEventArgs

    Public Sub New(ByVal vPrimaryTransaksi As Double)
        Me.mPrimarySalary = vPrimaryTransaksi
        Me.InitializeComponent()
        'pindahin ke dgv
        Me.ObjItem.IsiDgvComponent(Me, Me.DefCurr)
        'Me.ColIsModify.Visible = True
        'Me.ColIncome_Identifier.Visible = True

        Me.ColIncome_Identifier.Visible = False

    End Sub

    Enum EnumCalculatePeriod
        Daily
        monthly
        Yearly
    End Enum

    Enum EnumType
        Income
        Deduction
        CompanyExpense
        EmployeeExpense
    End Enum

    Private Function GetJmlConst(ByVal vIdentifier As String) As Double
        Try
            Me.TblCons.DefaultView.RowFilter = "[Identifier Salary Component] = '" + vIdentifier + "'"
            Return Me.TblCons.DefaultView.Count
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlVar(ByVal vIdentifier As String) As Double
        Try
            Me.TblVar.DefaultView.RowFilter = "[Identifier Salary Component] = '" + vIdentifier + "'"
            Return Me.TblVar.DefaultView.Count
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahDataBenar() As Boolean
        Dim row As DataGridViewRow
        Dim CboDgv As DataGridViewComboBoxCell

        Try
            'dgv component tidak boleh kosong
            If Me.DgvComponent.Rows.Count <= 1 Then
                MsgBox("Salary component can not be empty", MsgBoxStyle.Information, "Information")
                Me.TabControl1.SelectedTab = Me.TabPage1
                Return False
            End If

            For Each row In Me.DgvComponent.Rows
                If row.Index <> Me.DgvComponent.Rows.Count - 1 Then
                    With row
                        If String.IsNullOrEmpty(.Cells(Me.ColIncome_Desc.Index).Value) = True Then
                            MsgBox("Description in Salary component cannot be an empty string", MsgBoxStyle.Information, "Information")
                            Me.TabControl1.SelectedTab = Me.TabPage1
                            Me.DgvComponent.Select()
                            .Cells(Me.ColIncome_Desc.Index).Selected = True
                            Return False
                        End If
                        If String.IsNullOrEmpty(.Cells(Me.ColIncome_SlipDescription.Index).Value) = True Then
                            CboDgv = Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, row.Index)
                            If CboDgv.Items.IndexOf(Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, row.Index).Value) <> EnumType.CompanyExpense Then
                                MsgBox("Salary slip description in salary component cannot be an empty string", MsgBoxStyle.Information, "Information")
                                Me.TabControl1.SelectedTab = Me.TabPage1
                                Me.DgvComponent.Select()
                                .Cells(Me.ColIncome_SlipDescription.Index).Selected = True
                                Return False
                            End If
                        End If

                        If String.IsNullOrEmpty(.Cells(Me.ColIncome_ExpenseAccount.Index).Value) = True Then
                            CboDgv = Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, row.Index)
                            If CboDgv.Items.IndexOf(Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, row.Index).Value) <> EnumType.EmployeeExpense Then
                                MsgBox("Expense account in Salary component cannot be an empty string", MsgBoxStyle.Information, "Information")
                                Me.TabControl1.SelectedTab = Me.TabPage1
                                .Cells(Me.ColIncome_ExpenseAccount.Index).Selected = True
                                Return False
                            End If
                        End If
                        If String.IsNullOrEmpty(.Cells(Me.ColFormula.Index).Value) = True Then
                            MsgBox("Formula in Salary component cannot be an empty string", MsgBoxStyle.Information, "Information")
                            Me.TabControl1.SelectedTab = Me.TabPage1
                            .Cells(Me.ColFormula.Index).Selected = True
                            Return False
                        End If
                    End With
                End If
            Next

            'dgv emp tidak boleh kosong
            If Me.dgvEmployee.Rows.Count <= 1 Then
                'MsgBox("Detail Employee can not be empty", MsgBoxStyle.Information, "Information")
                'Me.TabControl1.SelectedTab = Me.TabPage2
                'Return False
            End If
            For Each row In Me.dgvEmployee.Rows
                If row.Index <> Me.dgvEmployee.Rows.Count - 1 Then
                    'sementara
                    'With row
                    '    If String.IsNullOrEmpty(.Cells(Me.ColEmp_EmployeeCode.Index).Value) = True Then
                    '        MsgBox("Employee code in Detail employee cannot be an empty string", MsgBoxStyle.Information, "Information")
                    '        Me.TabControl1.SelectedTab = Me.TabPage2
                    '        Me.dgvEmployee.Rows(.Index).Selected = True
                    '        Return False
                    '    End If
                    '    If String.IsNullOrEmpty(.Cells(Me.ColEmp_Basic.Index).Value) = True Then
                    '        MsgBox("Basic salary in Detail employee cannot be an empty string", MsgBoxStyle.Information, "Information")
                    '        Me.TabControl1.SelectedTab = Me.TabPage2
                    '        Me.dgvEmployee.Rows(.Index).Selected = True
                    '        Return False
                    '    End If
                    'End With
                    If String.IsNullOrEmpty(row.Cells(Me.ColEmp_Basic.Index).Value) = True Then
                        MsgBox("Basic salary in employee detail cannot be an empty string", MsgBoxStyle.Information, "Information")
                        Me.TabControl1.SelectedTab = Me.TabPage2
                        Me.dgvEmployee.Select()
                        row.Cells(Me.ColEmp_Basic.Index).Selected = True
                        Return False
                    Else
                        If Me.ObjCd.CekBilanganPositif(row.Cells(Me.ColEmp_Basic.Index).Value) = False Then
                            MsgBox("Please input postive number in basic salary", MsgBoxStyle.Information, "Information")
                            Me.TabControl1.SelectedTab = Me.TabPage2
                            Me.dgvEmployee.Select()
                            row.Cells(Me.ColEmp_Basic.Index).Selected = True
                            Return False
                        End If
                    End If
                    If String.IsNullOrEmpty(row.Cells(Me.ColEmp_SalarySlipDesc.Index).Value) = True Then
                        MsgBox("Salary slip description in employee detail cannot be an empty string", MsgBoxStyle.Information, "Information")
                        Me.TabControl1.SelectedTab = Me.TabPage2
                        Me.dgvEmployee.Select()
                        row.Cells(Me.ColEmp_SalarySlipDesc.Index).Selected = True
                        Return False
                    End If
                    If String.IsNullOrEmpty(row.Cells(Me.ColEmp_ExpenseAcc.Index).Value) = True Then
                        MsgBox("Expense account in employee detail cannot be an empty string", MsgBoxStyle.Information, "Information")
                        Me.TabControl1.SelectedTab = Me.TabPage2
                        row.Cells(Me.ColEmp_ExpenseAcc.Index).Selected = True
                        Return False
                    End If
                    If String.IsNullOrEmpty(row.Cells(Me.ColEmp_PaymentAcc.Index).Value) = True Then
                        MsgBox("Payment account in employee detail cannot be an empty string", MsgBoxStyle.Information, "Information")
                        Me.TabControl1.SelectedTab = Me.TabPage2
                        row.Cells(Me.ColEmp_PaymentAcc.Index).Selected = True
                        Return False
                    End If
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub WriteModified()
        Dim RowDGV As New DataGridViewRow
        Dim rowSC As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent1Row
        Dim rowEMP As DseMasterSalaryFormula.tblmaster_salary_formula_component_detailemployee1Row
        Dim CboDgvType As New DataGridViewComboBoxCell
        Dim CboDgvTax As New DataGridViewComboBoxCell
        Dim CboDgvPeriodOn As New DataGridViewComboBoxCell
        Dim CboDgvCalPeriod As New DataGridViewComboBoxCell
        Dim RowDGVE As New DataGridViewRow
        Dim CboTaxCal As New DataGridViewComboBoxCell

        Try
            'Cek DGV
            Me.tblLogUser.Rows.Clear()

            Me.TblCons.DefaultView.RowFilter = Nothing
            Me.TblConsOld.DefaultView.RowFilter = Nothing
            Me.TblVar.DefaultView.RowFilter = Nothing
            Me.TblVarOld.DefaultView.RowFilter = Nothing
            '=======================================================================================================================================================================
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            'TAB SALARY MULAI()
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            '=======================================================================================================================================================================
            '=======================================================================================================================================================================
            'DELETE >>>> Pengecekan apakah ada component, variable , constraint yang hilang >> maka data dihapus (mulai)
            '=======================================================================================================================================================================
            Dim ListNewPrimary As New List(Of Double)
            If Me.DgvComponent.Rows.Count > 0 Then
                For Each RowDGV In Me.DgvComponent.Rows
                    With RowDGV
                        If RowDGV.Index <> Me.DgvComponent.Rows.Count - 1 Then
                            '-------------------------------------------------------------------------------------------------------
                            'simpan primary component baru pada list
                            If ListNewPrimary.Contains(RowDGV.Cells(Me.Colincome_Primary.Index).Value) = False Then
                                If RowDGV.Cells(Me.Colincome_Primary.Index).Value <> -1 Then
                                    ListNewPrimary.Add(RowDGV.Cells(Me.Colincome_Primary.Index).Value)
                                End If
                            End If
                            '-------------------------------------------------------------------------------------------------------
                            '=====================================================================================================================================
                            'VARIABLE
                            '-------------------------------------------------------------------------------------------------------
                            'simpan primary variable baru pada list
                            Dim strDeleteVar As String = ""
                            Dim ListNewPrimaryVar As New List(Of Double)
                            Me.TblVar.DefaultView.RowFilter = Nothing
                            Me.TblVar.DefaultView.RowFilter = " [Identifier Salary Component] = '" & .Cells(Me.ColIncome_Identifier.Index).Value & "'"
                            If Me.TblVar.DefaultView.Count > 0 Then
                                For Each dvV As Data.DataRowView In Me.TblVar.DefaultView
                                    Me.RowVar = dvV.Row
                                    If Me.RowVar.Primary <> -1 Then
                                        If ListNewPrimaryVar.Contains(Me.RowVar.Primary) = False Then
                                            ListNewPrimaryVar.Add(Me.RowVar.Primary)
                                        End If
                                    End If
                                Next
                            End If
                            '-------------------------------------------------------------------------------------------------------
                            '-------------------------------------------------------------------------------------------------------
                            'catat variable hilang pada log user
                            Me.TblVarOld.DefaultView.RowFilter = Nothing
                            Me.TblVarOld.DefaultView.RowFilter = " [Identifier Salary Component] = '" & .Cells(Me.ColIncome_Identifier.Index).Value & "'"
                            If Me.TblVarOld.DefaultView.Count > 0 Then
                                For Each dvV As Data.DataRowView In Me.TblVarOld.DefaultView
                                    Me.RowVarOld = dvV.Row
                                    If Me.RowVarOld.Primary <> -1 Then
                                        If ListNewPrimaryVar.Contains(Me.RowVarOld.Primary) = False Then
                                            RowDGV.Cells(Me.ColIsModify.Index).Value = True
                                            If RowVarOld.Type2 = Core_Module.ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode Then
                                                strDeleteVar &= vbNewLine & "Deleting variable with parameter variable list " & _
                                                Me.ObjParameter.GetCodeByPrimary(RowVarOld.Primary_Parameter)
                                            ElseIf RowVarOld.Type2 = Core_Module.ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi Then
                                                strDeleteVar &= vbNewLine & "Deleting variable with parameter competence " & _
                                                Me.ObjCompetence.GetDataCode(RowVarOld.Primary_Parameter)
                                            End If

                                            'Me.rowLogUser = Me.tblLogUser.NewRow
                                            'Me.rowLogUser.Primary = -1
                                            'Me.rowLogUser.Activity_Type = ClsLogUser.EnumActivityType.Modify
                                            'Me.rowLogUser.User = Me.mUserCode
                                            'Me.rowLogUser.Activity_Time = Now
                                            'Me.rowLogUser.Primary_Salary_Formula = Me.mPrimarySalary

                                            'If RowVarOld.Type2 = Core_Module.ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode Then
                                            '    Me.rowLogUser.Remarks = "Deleting variable with parameter variable list " & _
                                            '    Me.ObjParameter.GetCodeByPrimary(RowVarOld.Primary_Parameter) & _
                                            '    " On salary component " & .Cells(Me.ColIncome_Desc.Index).Value
                                            'ElseIf RowVarOld.Type2 = Core_Module.ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi Then
                                            '    Me.rowLogUser.Remarks = vbNewLine & "Deleting variable with parameter competence " & _
                                            '    Me.ObjCompetence.GetDataCode(RowVarOld.Primary_Parameter) & " On salary component " _
                                            '    & .Cells(Me.ColIncome_Desc.Index).Value
                                            'End If
                                            'Me.tblLogUser.Rows.Add(Me.rowLogUser)
                                        End If
                                    End If
                                Next
                            End If
                            ''-------------------------------------------------------------------------------------------------------
                            ''=====================================================================================================================================
                            '=====================================================================================================================================
                            'CONSTRAINT
                            '-------------------------------------------------------------------------------------------------------
                            'simpan primary Constraint baru pada list
                            Dim ListNewPrimaryConts As New List(Of Double)
                            Dim strDeleteConts As String = ""
                            Me.TblCons.DefaultView.RowFilter = Nothing
                            Me.TblCons.DefaultView.RowFilter = " [Identifier Salary Component] = '" & .Cells(Me.ColIncome_Identifier.Index).Value & "'"
                            If Me.TblCons.DefaultView.Count > 0 Then
                                For Each dvC As Data.DataRowView In Me.TblCons.DefaultView
                                    Me.RowCons = dvC.Row
                                    If Me.RowCons.Primary <> -1 Then
                                        If ListNewPrimaryConts.Contains(Me.RowCons.Primary) = False Then
                                            ListNewPrimaryConts.Add(Me.RowCons.Primary)
                                        End If
                                    End If
                                Next
                            End If
                            '-------------------------------------------------------------------------------------------------------
                            '-------------------------------------------------------------------------------------------------------
                            'catat Constraint hilang pada log user
                            Me.TblConsOld.DefaultView.RowFilter = Nothing
                            Me.TblConsOld.DefaultView.RowFilter = " [Identifier Salary Component] = '" & .Cells(Me.ColIncome_Identifier.Index).Value & "'"
                            If Me.TblConsOld.DefaultView.Count > 0 Then
                                For Each dvC As Data.DataRowView In Me.TblConsOld.DefaultView
                                    Me.RowConsOld = dvC.Row
                                    If Me.RowConsOld.Primary <> -1 Then
                                        If ListNewPrimaryConts.Contains(Me.RowConsOld.Primary) = False Then
                                            RowDGV.Cells(Me.ColIsModify.Index).Value = True
                                            If RowConsOld.Type2 = Core_Module.ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode Then
                                                strDeleteConts &= vbNewLine & "Deleting constraint with parameter variable list " & _
                                                Me.ObjParameter.GetCodeByPrimary(RowConsOld.Primary_Parameter)
                                            ElseIf RowConsOld.Type2 = Core_Module.ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi Then
                                                strDeleteConts &= vbNewLine & "Deleting constraint with parameter competence " & _
                                                Me.ObjCompetence.GetDataCode(RowConsOld.Primary_Parameter)
                                            End If
                                            'Me.rowLogUser = Me.tblLogUser.Newtbl_log_userRow
                                            'Me.rowLogUser.Primary = -1
                                            'Me.rowLogUser.Activity_Type = ClsLogUser.EnumActivityType.Modify
                                            'Me.rowLogUser.User = Me.mUserCode
                                            'Me.rowLogUser.Activity_Time = Now
                                            'Me.rowLogUser.Primary_Salary_Formula = Me.mPrimarySalary
                                            'If RowConsOld.Type2 = Core_Module.ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode Then
                                            '    Me.rowLogUser.Remarks = vbNewLine & "Deleting constraint with parameter constraint list " & _
                                            '    Me.ObjParameter.GetCodeByPrimary(RowConsOld.Primary_Parameter) & _
                                            '    " On salary component " & .Cells(Me.ColIncome_Desc.Index).Value
                                            'ElseIf RowConsOld.Type2 = Core_Module.ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi Then
                                            '    Me.rowLogUser.Remarks = vbNewLine & "Deleting constraint with parameter competence " & _
                                            '    Me.ObjCompetence.GetDataCode(RowConsOld.Primary_Parameter) & " On salary component " _
                                            '    & .Cells(Me.ColIncome_Desc.Index).Value
                                            'End If
                                            'Me.tblLogUser.Rows.Add(Me.rowLogUser)
                                        End If
                                    End If
                                Next
                            End If
                            '-------------------------------------------------------------------------------------------------------
                            '=====================================================================================================================================
                            If strDeleteConts <> "" OrElse strDeleteVar <> "" Then
                                RowDGV.Cells(Me.ColIsModify.Index).Value = True
                                Me.rowLogUser = Me.tblLogUser.NewRow
                                Me.rowLogUser.Primary = -1
                                Me.rowLogUser.Activity_Type = ClsLogUser.EnumActivityType.Modify
                                Me.rowLogUser.User = Me.mUserCode
                                Me.rowLogUser.Activity_Time = Now
                                Me.rowLogUser.Primary_Salary_Formula = Me.mPrimarySalary
                                Me.rowLogUser.Remarks = "Modify on salary component " & .Cells(Me.ColIncome_Desc.Index).Value & " at : "
                                Me.rowLogUser.Salary_Formula_Code = Me.ObjSalaryFormula.GetCodeByPrimary(Me.mPrimarySalary)
                                If strDeleteConts <> "" Then
                                    Me.rowLogUser.Remarks &= strDeleteConts
                                End If
                                If strDeleteVar <> "" Then
                                    Me.rowLogUser.Remarks &= strDeleteVar
                                End If
                                Me.tblLogUser.Rows.Add(Me.rowLogUser)
                            End If

                        End If
                    End With
                Next
            End If
            '-------------------------------------------------------------------------------------------------------
            'catat component hilang pada log user
            If Me.tblSC.Rows.Count > 0 Then
                For Each rowSC In Me.tblSC.Rows
                    With rowSC
                        If .Primary <> -1 Then
                            If ListNewPrimary.Contains(rowSC.Primary) = False Then
                                Me.rowLogUser = Me.tblLogUser.Newtbl_log_userRow
                                Me.rowLogUser.Primary = -1
                                Me.rowLogUser.Activity_Type = ClsLogUser.EnumActivityType.Delete
                                Me.rowLogUser.User = Me.mUserCode
                                Me.rowLogUser.Activity_Time = Now
                                Me.rowLogUser.Remarks = "Deleting salary component " & rowSC.Description
                                Me.rowLogUser.Primary_Salary_Formula = Me.mPrimarySalary
                                Me.rowLogUser.Salary_Formula_Code = Me.ObjSalaryFormula.GetCodeByPrimary(Me.mPrimarySalary)
                                Me.tblLogUser.Rows.Add(Me.rowLogUser)
                            End If
                        End If
                    End With
                Next
            End If
            '=======================================================================================================================================================================
            'Pengecekan apakah ada data yang hilang >> maka data dihapus (selesai)
            '=======================================================================================================================================================================
            '=======================================================================================================================================================================




            '=======================================================================================================================================================================
            '=======================================================================================================================================================================
            'MODIFY, CREATE
            '=======================================================================================================================================================================
            For Each RowDGV In Me.DgvComponent.Rows
                With RowDGV
                    If RowDGV.Index <> Me.DgvComponent.Rows.Count - 1 Then
                        Dim strPerubahan As String = ""
                        Dim Routindisetahunkan As Integer = 0
                        Dim Type As Integer = 0
                        If .Cells(Me.ColIncome_Routine.Index).Value = Boolean.TrueString Then
                            Routindisetahunkan = 1
                        Else
                            Routindisetahunkan = 0
                        End If

                        CboDgvType = Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, RowDGV.Index)
                        If CboDgvType.Items.IndexOf(Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, RowDGV.Index).Value) = EnumType.CompanyExpense Then
                            Type = EnumType.CompanyExpense
                        ElseIf CboDgvType.Items.IndexOf(Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, RowDGV.Index).Value) = EnumType.Deduction Then
                            Type = EnumType.Deduction
                        ElseIf CboDgvType.Items.IndexOf(Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, RowDGV.Index).Value) = EnumType.EmployeeExpense Then
                            Type = EnumType.EmployeeExpense
                        ElseIf CboDgvType.Items.IndexOf(Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, RowDGV.Index).Value) = EnumType.Income Then
                            Type = EnumType.Income
                        End If
                        CboDgvTax = Me.DgvComponent.Item(Me.ColIncome_TaxComponent.Index, RowDGV.Index)
                        CboDgvPeriodOn = Me.DgvComponent.Item(Me.ColIncome_CalculateOn.Index, RowDGV.Index)
                        CboDgvCalPeriod = Me.DgvComponent.Item(Me.ColIncome_CalculatePeriod.Index, RowDGV.Index)


                        Me.tblSC.DefaultView.RowFilter = Nothing
                        Me.tblSC.DefaultView.RowFilter = "[Primary] = " & .Cells(Me.Colincome_Primary.Index).Value
                        If tblSC.DefaultView.Count = 0 Then
                            '=====================================================================================================================================
                            'Jika tidak terdapat primary yang sama didatabase, maka CREATE new salary component
                            '=====================================================================================================================================
                            Me.rowLogUser = Me.tblLogUser.NewRow
                            Me.rowLogUser.Primary = -1
                            Me.rowLogUser.Activity_Type = ClsLogUser.EnumActivityType.Create
                            Me.rowLogUser.User = Me.mUserCode
                            Me.rowLogUser.Activity_Time = Now
                            Me.rowLogUser.Primary_Salary_Formula = Me.mPrimarySalary
                            Me.rowLogUser.Salary_Formula_Code = Me.ObjSalaryFormula.GetCodeByPrimary(Me.mPrimarySalary)
                            Me.rowLogUser.Remarks = "Creating new salary component " & .Cells(Me.ColIncome_Desc.Index).Value
                            Me.tblLogUser.Rows.Add(Me.rowLogUser)
                            '=====================================================================================================================================
                        ElseIf tblSC.DefaultView.Count > 0 AndAlso .Cells(Me.Colincome_Primary.Index).Value <> -1 Then
                            '=====================================================================================================================================
                            'Jika terdapat primary yang sama didatabase, maka modified SALARY COMPONENT
                            '=====================================================================================================================================
                            rowSC = Me.tblSC.DefaultView(0).Row
                            If rowSC.Type <> Type Then
                                Dim strType1 As String = ""
                                Dim strType2 As String = ""
                                '=================================================================================================================================
                                'Pencatatan perubahan Salary component (mulai)
                                '---------------------------------------------------------------------------------------------------------------------------------
                                'change on type salary component
                                If rowSC.Type = EnumType.CompanyExpense Then
                                    strType1 = "Company Expense"
                                ElseIf rowSC.Type = EnumType.Deduction Then
                                    strType1 = "Deduction"
                                ElseIf rowSC.Type = EnumType.EmployeeExpense Then
                                    strType1 = "Employee Expense"
                                ElseIf rowSC.Type = EnumType.Income Then
                                    strType1 = "Income"
                                End If

                                If Type = EnumType.CompanyExpense Then
                                    strType2 = "Company Expense"
                                ElseIf Type = EnumType.Deduction Then
                                    strType2 = "Deduction"
                                ElseIf Type = EnumType.EmployeeExpense Then
                                    strType2 = "Employee Expense"
                                ElseIf Type = EnumType.Income Then
                                    strType2 = "Income"
                                End If

                                If strPerubahan = "" Then
                                    strPerubahan &= vbNewLine & " type salary component from " & strType1 & " to " & strType2
                                End If
                            End If
                            '---------------------------------------------------------------------------------------------------------------------------
                            'change on component name
                            If rowSC.Description <> .Cells(Me.ColIncome_Desc.Index).Value Then
                                strPerubahan &= vbNewLine & " component name from " & rowSC.Description & " to " & .Cells(Me.ColIncome_Desc.Index).Value
                            End If
                            '---------------------------------------------------------------------------------------------------------------------------
                            'change on rutin/annualized
                            If rowSC.rutin_disetahunkan <> Routindisetahunkan Then
                                If Routindisetahunkan = 1 Then
                                    strPerubahan &= vbNewLine & " rutin disetahunkan from false to true"
                                Else
                                    strPerubahan &= vbNewLine & " rutin disetahunkan from true to false"
                                End If
                            End If
                            '---------------------------------------------------------------------------------------------------------------------------
                            'change on Salary slip desc
                            If rowSC.Primary_Slip_Desc <> .Cells(Me.ColIncome_PrimarySlipDesc.Index).Value Then
                                strPerubahan &= vbNewLine & " slip desc from " & rowSC.Slip_Desc_Code & " to " & .Cells(Me.ColIncome_SlipDescription.Index).Value
                            End If
                            '---------------------------------------------------------------------------------------------------------------------------
                            'change on tax component
                            If rowSC.Tax_Component <> Me.ObjItem.GetTipeTaxComponent(Me.DgvComponent.Item(Me.ColIncome_TaxComponent.Index, RowDGV.Index).Value) Then
                                strPerubahan &= vbNewLine & " tax component from " & Me.ObjItem.GetTipeTaxComponent_Str(rowSC.Tax_Component) & " to " & .Cells(Me.ColIncome_TaxComponent.Index).Value
                            End If
                            '---------------------------------------------------------------------------------------------------------------------------
                            'change on expense account
                            If rowSC.Primary_Expense_Account <> .Cells(Me.ColIncome_PrimaryExpenseAccount.Index).Value Then
                                strPerubahan &= vbNewLine & " expense acc from " & rowSC.Expense_Account_Code & " to " & .Cells(Me.ColIncome_ExpenseAccount.Index).Value
                            End If
                            '---------------------------------------------------------------------------------------------------------------------------
                            'change on payment account
                            If rowSC.primary_payment_account <> .Cells(Me.ColIncome_PrimaryPaymentAccount.Index).Value Then
                                strPerubahan &= vbNewLine & " urutan from " & rowSC.Payment_Account_Code & " to " & .Cells(Me.ColIncome_PaymentAccount.Index).Value
                            End If
                            '---------------------------------------------------------------------------------------------------------------------------
                            'change on formula
                            If IsNumeric(rowSC.Formula) AndAlso IsNumeric(.Cells(Me.ColFormula.Index).Value) Then
                                If CDec(rowSC.Formula) <> CDec(.Cells(Me.ColFormula.Index).Value) Then
                                    strPerubahan &= vbNewLine & " formula from " & rowSC.Formula & " to " & .Cells(Me.ColFormula.Index).Value
                                End If
                            Else
                                If rowSC.Formula <> .Cells(Me.ColFormula.Index).Value Then
                                    strPerubahan &= vbNewLine & " formula from " & rowSC.Formula & " to " & .Cells(Me.ColFormula.Index).Value
                                End If
                            End If
                            '---------------------------------------------------------------------------------------------------------------------------
                            'change on Calculation_Period
                            If rowSC.Calculation_Period <> CboDgvCalPeriod.Items.IndexOf(Me.DgvComponent.Item(Me.ColIncome_CalculatePeriod.Index, RowDGV.Index).Value) Then
                                strPerubahan &= vbNewLine & " calculate period from " & CboDgvCalPeriod.Items(rowSC.Calculation_Period) & " to " & .Cells(Me.ColIncome_CalculatePeriod.Index).Value
                            End If
                            '---------------------------------------------------------------------------------------------------------------------------
                            'change on Calculate_On
                            If rowSC.Calculate_On <> CboDgvPeriodOn.Items.IndexOf(Me.DgvComponent.Item(Me.ColIncome_CalculateOn.Index, RowDGV.Index).Value) Then
                                strPerubahan &= vbNewLine & " calculate on from " & CboDgvPeriodOn.Items(rowSC.Calculate_On) & " to " & .Cells(Me.ColIncome_CalculateOn.Index).Value
                            End If
                            '---------------------------------------------------------------------------------------------------------------------------
                            'change on Rounding_Type
                            If rowSC.Rounding_Type <> .Cells(Me.ColRoundingType.Index).Value Then
                                strPerubahan &= vbNewLine & " rounding type from " & rowSC.Rounding_Type & " to " & .Cells(Me.ColRoundingType.Index).Value
                            End If
                            '---------------------------------------------------------------------------------------------------------------------------
                            'change on Rounding_Unit
                            If rowSC.Rounding_Unit <> .Cells(Me.ColRoundingUnit.Index).Value Then
                                strPerubahan &= vbNewLine & " rounding unit from " & rowSC.Rounding_Unit & " to " & .Cells(Me.ColRoundingUnit.Index).Value
                            End If
                            '---------------------------------------------------------------------------------------------------------------------------
                            'change on Yearly_Min_Value
                            If IsNumeric(rowSC.Yearly_Min_Value) AndAlso IsNumeric(.Cells(Me.ColYearlyMinValue.Index).Value) Then
                                If CDec(rowSC.Yearly_Min_Value) <> CDec(.Cells(Me.ColYearlyMinValue.Index).Value) Then
                                    strPerubahan &= vbNewLine & " yearly min from " & rowSC.Yearly_Min_Value & " to " & .Cells(Me.ColYearlyMinValue.Index).Value
                                End If
                            Else
                                If rowSC.Yearly_Min_Value <> .Cells(Me.ColYearlyMinValue.Index).Value Then
                                    strPerubahan &= vbNewLine & " yearly min from " & rowSC.Yearly_Min_Value & " to " & .Cells(Me.ColYearlyMinValue.Index).Value
                                End If
                            End If
                            '---------------------------------------------------------------------------------------------------------------------------
                            'change on Yearly_Max_Value
                            If IsNumeric(rowSC.Yearly_Max_Value) AndAlso IsNumeric(.Cells(Me.ColYearlyMaxValue.Index).Value) Then
                                If CDec(rowSC.Yearly_Max_Value) <> CDec(.Cells(Me.ColYearlyMaxValue.Index).Value) Then
                                    strPerubahan &= vbNewLine & " yearly max from " & rowSC.Yearly_Max_Value & " to " & .Cells(Me.ColYearlyMaxValue.Index).Value
                                End If
                            Else
                                If rowSC.Yearly_Max_Value <> .Cells(Me.ColYearlyMaxValue.Index).Value Then
                                    strPerubahan &= vbNewLine & " yearly max from " & rowSC.Yearly_Max_Value & " to " & .Cells(Me.ColYearlyMaxValue.Index).Value
                                End If
                            End If
                            '---------------------------------------------------------------------------------------------------------------------------
                            'change on Yearly_Date
                            If rowSC.Yearly_Date <> .Cells(Me.ColYearlyDate.Index).Value Then
                                strPerubahan &= vbNewLine & " yearly date from " & rowSC.Yearly_Date & " to " & .Cells(Me.ColYearlyDate.Index).Value
                            End If
                            '---------------------------------------------------------------------------------------------------------------------------
                            'change on Monthly_Min_Value
                            If IsNumeric(rowSC.Monthly_Min_Value) AndAlso IsNumeric(.Cells(Me.ColMonthlyMinValue.Index).Value) Then
                                If CDec(rowSC.Monthly_Min_Value) <> CDec(.Cells(Me.ColMonthlyMinValue.Index).Value) Then
                                    strPerubahan &= vbNewLine & " yearly min from " & rowSC.Monthly_Min_Value & " to " & .Cells(Me.ColMonthlyMinValue.Index).Value
                                End If
                            Else
                                If rowSC.Monthly_Min_Value <> .Cells(Me.ColMonthlyMinValue.Index).Value Then
                                    strPerubahan &= vbNewLine & " yearly min from " & rowSC.Monthly_Min_Value & " to " & .Cells(Me.ColMonthlyMinValue.Index).Value
                                End If
                            End If

                            '---------------------------------------------------------------------------------------------------------------------------
                            'change on Monthly_Max_Value
                            If IsNumeric(rowSC.Monthly_Max_Value) AndAlso IsNumeric(.Cells(Me.ColMonthlyMaxValue.Index).Value) Then
                                If CDec(rowSC.Monthly_Max_Value) <> CDec(.Cells(Me.ColMonthlyMaxValue.Index).Value) Then
                                    strPerubahan &= vbNewLine & " monthly max from " & rowSC.Monthly_Max_Value & " to " & .Cells(Me.ColMonthlyMaxValue.Index).Value
                                End If
                            Else
                                If rowSC.Monthly_Max_Value <> .Cells(Me.ColMonthlyMaxValue.Index).Value Then
                                    strPerubahan &= vbNewLine & " monthly max from " & rowSC.Monthly_Max_Value & " to " & .Cells(Me.ColMonthlyMaxValue.Index).Value
                                End If
                            End If
                            '---------------------------------------------------------------------------------------------------------------------------
                            'change on ColIncome_Urutan
                            If rowSC.Urutan <> .Cells(Me.ColIncome_Urutan.Index).Value Then
                                strPerubahan &= vbNewLine & " urutan from " & rowSC.Urutan & " to " & .Cells(Me.ColIncome_Urutan.Index).Value
                            End If
                            '---------------------------------------------------------------------------------------------------------------------------
                            'change on ColIncome_Identifier
                            If rowSC.Identifier <> .Cells(Me.ColIncome_Identifier.Index).Value Then
                                strPerubahan &= vbNewLine & " identifier from " & rowSC.Identifier & " to " & .Cells(Me.ColIncome_Identifier.Index).Value
                            End If
                            '---------------------------------------------------------------------------------------------------------------------------
                            '---------------------------------------------------------------------------------------------------------------------------
                            'catat perubahan pada tbl log user
                            If strPerubahan <> "" Then
                                RowDGV.Cells(Me.ColIsModify.Index).Value = True
                                Me.rowLogUser = Me.tblLogUser.NewRow
                                Me.rowLogUser.Primary = -1
                                Me.rowLogUser.Activity_Type = ClsLogUser.EnumActivityType.Modify
                                Me.rowLogUser.User = Me.mUserCode
                                Me.rowLogUser.Activity_Time = Now
                                Me.rowLogUser.Primary_Salary_Formula = Me.mPrimarySalary
                                Me.rowLogUser.Salary_Formula_Code = Me.ObjSalaryFormula.GetCodeByPrimary(Me.mPrimarySalary)
                                Me.rowLogUser.Remarks = "Modify on salary component " & rowSC.Description & " at : " & strPerubahan
                                Me.tblLogUser.Rows.Add(Me.rowLogUser)
                            End If
                            '---------------------------------------------------------------------------------------------------------------------------------
                            '---------------------------------------------------------------------------------------------------------------------------------
                            'Pencatatan perubahan Salary component (selesai)
                            '=================================================================================================================================

                            '=================================================================================================================================
                            'Pencatatan perubahan Salary component - variable  (mulai)
                            '---------------------------------------------------------------------------------------------------------------------------------
                            Dim strPerubahanVar As String = ""
                            Me.TblVar.DefaultView.RowFilter = Nothing
                            Me.TblVar.DefaultView.RowFilter = "[Identifier Salary Component] = '" & .Cells(Me.ColIncome_Identifier.Index).Value & "'"
                            If Me.TblVar.DefaultView.Count > 0 Then
                                Dim dv As Data.DataRowView
                                For Each dv In Me.TblVar.DefaultView
                                    Me.RowVar = dv.Row
                                    Me.TblVarOld.DefaultView.RowFilter = Nothing
                                    Me.TblVarOld.DefaultView.RowFilter = "[primary] = " & Me.RowVar.Primary & _
                                                                       " And [Identifier Salary Component] = '" & .Cells(Me.ColIncome_Identifier.Index).Value & "'"
                                    If Me.TblVarOld.DefaultView.Count = 0 Then
                                        '----------------------------------------------------------------------------------------------------------
                                        'jika tidak ada baris, maka CREATE variable
                                        '----------------------------------------------------------------------------------------------------------
                                        If RowVar.Type2 = Core_Module.ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode Then
                                            strPerubahanVar &= vbNewLine & "creating new variable with parameter variable list " & _
                                                                            Me.ObjParameter.GetCodeByPrimary(RowVar.Primary_Parameter)
                                        ElseIf RowVar.Type2 = Core_Module.ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi Then
                                            strPerubahanVar &= vbNewLine & "creating new variable with parameter competence " & _
                                                                            Me.ObjCompetence.GetDataCode(RowVar.Primary_Parameter)
                                        End If
                                    ElseIf Me.TblVarOld.DefaultView.Count > 0 Then
                                        '----------------------------------------------------------------------------------------------------------
                                        'jika ada baris maka MODIFY variable, dicek perubahannya
                                        '----------------------------------------------------------------------------------------------------------
                                        Me.RowVarOld = Me.TblVarOld.DefaultView(0).Row
                                        strPerubahanVar = ""
                                        If RowVarOld.Type <> Me.RowVar.Type Then
                                            strPerubahanVar &= vbNewLine & "type from " & RowVarOld.Type & " to " & Me.RowVar.Type
                                        End If

                                        If RowVarOld.Primary_Parameter <> Me.RowVar.Primary_Parameter Then
                                            Dim strParameterOld As String = ""
                                            Dim strParameterNew As String = ""
                                            If RowVarOld.Type2 = Core_Module.ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode Then
                                                strParameterOld = Me.ObjParameter.GetCodeByPrimary(RowVarOld.Primary_Parameter)
                                            ElseIf RowVarOld.Type2 = Core_Module.ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi Then
                                                strParameterOld = Me.ObjCompetence.GetDataCode(RowVarOld.Primary_Parameter)
                                            End If

                                            If RowVar.Type2 = Core_Module.ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode Then
                                                strParameterNew = Me.ObjParameter.GetCodeByPrimary(RowVar.Primary_Parameter)
                                            ElseIf RowVar.Type2 = Core_Module.ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi Then
                                                strParameterNew = Me.ObjCompetence.GetDataCode(RowVar.Primary_Parameter)
                                            End If

                                            strPerubahanVar &= vbNewLine & "parameter from " & strParameterOld & " to " & strParameterNew
                                        End If


                                        If RowVarOld.Variable <> Me.RowVar.Variable Then
                                            strPerubahanVar &= vbNewLine & "variable from " & RowVarOld.Variable & " to " & Me.RowVar.Variable
                                        End If

                                        If RowVarOld.Rounding_Unit <> Me.RowVar.Rounding_Unit Then
                                            strPerubahanVar &= vbNewLine & "rounding unit from " & RowVarOld.Rounding_Unit & " to " & Me.RowVar.Rounding_Unit
                                        End If

                                        If RowVarOld.Rounding_Type <> Me.RowVar.Rounding_Type Then
                                            strPerubahanVar &= vbNewLine & "rounding type from " & RowVarOld.Rounding_Type & " to " & Me.RowVar.Rounding_Type
                                        End If

                                        If RowVarOld.Type2 <> Me.RowVar.Type2 Then
                                            strPerubahanVar &= vbNewLine & "type2 from " & RowVarOld.Type2 & " to " & Me.RowVar.Type2
                                        End If

                                        If RowVarOld.Identifier_Salary_Component <> Me.RowVar.Identifier_Salary_Component Then
                                            strPerubahanVar &= vbNewLine & "identifier salary component from " & RowVarOld.Identifier_Salary_Component & " to " & _
                                            Me.RowVar.Identifier_Salary_Component
                                        End If
                                    End If
                                Next

                                'If strPerubahanVar <> "" Then
                                '    Me.rowLogUser = Me.tblLogUser.NewRow
                                '    Me.rowLogUser.Primary = -1
                                '    Me.rowLogUser.Activity_Type = ClsLogUser.EnumActivityType.Modify
                                '    Me.rowLogUser.User = Me.mUserCode
                                '    Me.rowLogUser.Activity_Time = Now
                                '    Me.rowLogUser.Remarks = "Modify salary variable on salary component " & rowSC.Description & ":" & strPerubahanVar
                                '    Me.tblLogUser.Rows.Add(Me.rowLogUser)
                                'End If
                            End If
                            '---------------------------------------------------------------------------------------------------------------------------------
                            'Pencatatan perubahan Salary component - variable  (selesai)
                            '=================================================================================================================================



                            '=================================================================================================================================
                            'Pencatatan perubahan Salary component - constraint  (mulai)
                            '---------------------------------------------------------------------------------------------------------------------------------
                            Dim strPerubahancons As String = ""
                            Dim strCreateCons As String = ""
                            Me.TblCons.DefaultView.RowFilter = Nothing
                            Me.TblCons.DefaultView.RowFilter = "[Identifier Salary Component] = '" & .Cells(Me.ColIncome_Identifier.Index).Value & "'"
                            If Me.TblCons.DefaultView.Count > 0 Then
                                Dim dv As Data.DataRowView
                                For Each dv In Me.TblCons.DefaultView
                                    Me.RowCons = dv.Row
                                    Me.TblConsOld.DefaultView.RowFilter = Nothing
                                    Me.TblConsOld.DefaultView.RowFilter = "[primary] = " & Me.RowCons.Primary & " And [Identifier Salary Component] = '" & .Cells(Me.ColIncome_Identifier.Index).Value & "'"
                                    If Me.TblConsOld.DefaultView.Count = 0 Then
                                        '----------------------------------------------------------------------------------------------------------
                                        'jika tidak ada baris, maka terjadi penambahan data consiable
                                        '----------------------------------------------------------------------------------------------------------
                                        If RowCons.Type2 = Core_Module.ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode Then
                                            strPerubahancons &= vbNewLine & "creating new constraint with parameter variable list " & Me.ObjParameter.GetCodeByPrimary(RowCons.Primary_Parameter)
                                        ElseIf RowCons.Type2 = Core_Module.ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi Then
                                            strPerubahancons &= vbNewLine & "creating new constraint with parameter competence " & Me.ObjCompetence.GetDataCode(RowCons.Primary_Parameter)
                                        End If
                                    ElseIf Me.TblConsOld.DefaultView.Count > 0 Then
                                        '----------------------------------------------------------------------------------------------------------
                                        'jika ada baris, dicek perubahannya
                                        '----------------------------------------------------------------------------------------------------------
                                        Me.RowConsOld = Me.TblConsOld.DefaultView(0).Row
                                        strPerubahancons = ""
                                        If RowConsOld.Type <> Me.RowCons.Type Then
                                            strPerubahancons &= vbNewLine & "type from " & RowConsOld.Type & " to " & Me.RowCons.Type
                                        End If

                                        If RowConsOld.Primary_Parameter <> Me.RowCons.Primary_Parameter Then
                                            Dim strParameterOld As String = ""
                                            Dim strParameterNew As String = ""

                                            If RowConsOld.Type2 = Core_Module.ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode Then
                                                strParameterOld = Me.ObjParameter.GetCodeByPrimary(RowConsOld.Primary_Parameter)
                                            ElseIf RowConsOld.Type2 = Core_Module.ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi Then
                                                strParameterOld = Me.ObjCompetence.GetDataCode(RowConsOld.Primary_Parameter)
                                            End If


                                            If RowCons.Type2 = Core_Module.ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode Then
                                                strParameterNew = Me.ObjParameter.GetCodeByPrimary(RowCons.Primary_Parameter)
                                            ElseIf RowCons.Type2 = Core_Module.ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi Then
                                                strParameterNew = Me.ObjCompetence.GetDataCode(RowCons.Primary_Parameter)
                                            End If
                                            strPerubahancons &= vbNewLine & "parameter from " & strParameterOld & " to " & strParameterNew
                                        End If


                                        If RowConsOld.Compare <> Me.RowCons.Compare Then
                                            strPerubahancons &= vbNewLine & "Compare from " & RowConsOld.Compare & " to " & Me.RowCons.Compare
                                        End If

                                        If RowConsOld.Ref_Value <> Me.RowCons.Ref_Value Then
                                            strPerubahancons &= vbNewLine & "Ref Value from " & RowConsOld.Ref_Value & " to " & Me.RowCons.Ref_Value
                                        End If

                                        If RowConsOld.Identifier_Salary_Component <> Me.RowCons.Identifier_Salary_Component Then
                                            strPerubahancons &= vbNewLine & "identifier salary component from " & RowConsOld.Identifier_Salary_Component & _
                                            " to " & Me.RowCons.Identifier_Salary_Component
                                        End If
                                        If RowConsOld.Type2 <> Me.RowCons.Type2 Then
                                            strPerubahancons &= vbNewLine & "type2 from " & RowConsOld.Type2 & " to " & Me.RowCons.Type2
                                        End If
                                    End If
                                Next
                                'If strPerubahancons <> "" Then
                                '    Me.rowLogUser = Me.tblLogUser.NewRow
                                '    Me.rowLogUser.Primary = -1
                                '    Me.rowLogUser.Activity_Type = ClsLogUser.EnumActivityType.Modify
                                '    Me.rowLogUser.User = Me.mUserCode
                                '    Me.rowLogUser.Activity_Time = Now
                                '    Me.rowLogUser.Remarks = "Modify salary constraint on salary component " & rowSC.Description & ":" & strPerubahancons
                                '    Me.tblLogUser.Rows.Add(Me.rowLogUser)
                                'End If
                            End If
                            '---------------------------------------------------------------------------------------------------------------------------------
                            'Pencatatan perubahan Salary component - constraint  (selesai)
                            '=================================================================================================================================
                            If strPerubahancons <> "" OrElse strPerubahanVar <> "" Then
                                RowDGV.Cells(Me.ColIsModify.Index).Value = True
                                Me.rowLogUser = Me.tblLogUser.NewRow
                                Me.rowLogUser.Primary = -1
                                Me.rowLogUser.Activity_Type = ClsLogUser.EnumActivityType.Modify
                                Me.rowLogUser.User = Me.mUserCode
                                Me.rowLogUser.Activity_Time = Now
                                Me.rowLogUser.Primary_Salary_Formula = Me.mPrimarySalary
                                Me.rowLogUser.Salary_Formula_Code = Me.ObjSalaryFormula.GetCodeByPrimary(Me.mPrimarySalary)
                                Me.rowLogUser.Remarks = "Modify on salary component " & rowSC.Description & " at : "
                                If strPerubahancons <> "" Then
                                    Me.rowLogUser.Remarks &= strPerubahancons
                                End If
                                If strPerubahanVar <> "" Then
                                    Me.rowLogUser.Remarks &= strPerubahanVar
                                End If
                                Me.tblLogUser.Rows.Add(Me.rowLogUser)
                            End If
                        End If
                    End If
                End With
            Next
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            'TAB SALARY SELESAI()
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            '=======================================================================================================================================================================
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            'TAB EMPLOYEE MULAI()
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            Dim ListNewPrimaryEmployee As New List(Of Double)
            If Me.dgvEmployee.RowCount > 0 Then
                For Each RowDGVE In Me.dgvEmployee.Rows
                    With RowDGVE
                        If RowDGVE.Index <> Me.dgvEmployee.Rows.Count - 1 Then
                            CboTaxCal = Me.dgvEmployee.Item(Me.ColEmp_TaxCalculation.Index, RowDGVE.Index)
                            If .Cells(Me.ColEmp_Primary.Index).Value <> -1 Then
                                If ListNewPrimaryEmployee.Contains(.Cells(Me.ColEmp_Primary.Index).Value) = False Then
                                    ListNewPrimaryEmployee.Add(.Cells(Me.ColEmp_Primary.Index).Value)
                                End If
                            End If
                        End If
                    End With
                Next
            End If


            Me.tblEMP.DefaultView.RowFilter = Nothing
            If Me.tblEMP.Rows.Count > 0 Then
                For Each rowEMP In Me.tblEMP.Rows
                    With rowEMP
                        If .Primary <> -1 Then
                            If ListNewPrimaryEmployee.Contains(.Primary) = False Then
                                Me.rowLogUser = Me.tblLogUser.Newtbl_log_userRow
                                Me.rowLogUser.Primary = -1
                                Me.rowLogUser.Activity_Type = ClsLogUser.EnumActivityType.Delete
                                Me.rowLogUser.User = Me.mUserCode
                                Me.rowLogUser.Activity_Time = Now
                                Me.rowLogUser.Remarks = "Deleting data for employee " & Me.ObjEmp.GetFullNameByPrimary(rowEMP.Primary_Karyawan)
                                Me.rowLogUser.Primary_Salary_Formula = Me.mPrimarySalary
                                Me.rowLogUser.Salary_Formula_Code = Me.ObjSalaryFormula.GetCodeByPrimary(Me.mPrimarySalary)
                                Me.tblLogUser.Rows.Add(Me.rowLogUser)
                            End If
                        End If
                    End With
                Next
            End If

            For Each RowDGVE In Me.dgvEmployee.Rows
                With RowDGVE
                    If RowDGVE.Index <> Me.dgvEmployee.Rows.Count - 1 Then

                        Dim strPerubahanEMP As String = ""
                        Me.tblEMP.DefaultView.RowFilter = Nothing
                        Me.tblEMP.DefaultView.RowFilter = "[Primary] = " & .Cells(Me.ColEmp_Primary.Index).Value
                        If Me.tblEMP.DefaultView.Count = 0 Then
                            'creating new data
                            Me.rowLogUser = Me.tblLogUser.Newtbl_log_userRow
                            Me.rowLogUser.Primary = -1
                            Me.rowLogUser.Activity_Type = ClsLogUser.EnumActivityType.Create
                            Me.rowLogUser.User = Me.mUserCode
                            Me.rowLogUser.Activity_Time = Now
                            Me.rowLogUser.Remarks = "creating new data for employee " & .Cells(Me.ColEmp_EmployeeName.Index).Value
                            Me.rowLogUser.Primary_Salary_Formula = Me.mPrimarySalary
                            Me.rowLogUser.Salary_Formula_Code = Me.ObjSalaryFormula.GetCodeByPrimary(Me.mPrimarySalary)
                            Me.tblLogUser.Rows.Add(Me.rowLogUser)
                        ElseIf Me.tblEMP.DefaultView.Count > 0 Then
                            'checking for modification
                            For Each dvEmp As Data.DataRowView In Me.tblEMP.DefaultView
                                rowEMP = dvEmp.Row
                                If rowEMP.Primary_Karyawan <> .Cells(Me.ColEmp_PrimaryEmployee.Index).Value Then
                                    strPerubahanEMP &= vbNewLine & "Employee data from " & _
                                                                rowEMP.Employee_Name & " to " & .Cells(Me.ColEmp_EmployeeName.Index).Value
                                End If


                                If IsNumeric(rowEMP.Basic_Salary) AndAlso IsNumeric(.Cells(Me.ColEmp_Basic.Index).Value) Then
                                    If CDec(rowEMP.Basic_Salary) <> CDec(.Cells(Me.ColEmp_Basic.Index).Value) Then
                                        strPerubahanEMP &= vbNewLine & "Basic salary from " & _
                                                                    Me.ObjCd.DigitGrouping(rowEMP.Basic_Salary, Me.DefCurr) & " to " & .Cells(Me.ColEmp_Basic.Index).Value
                                    End If
                                Else
                                    If rowEMP.Basic_Salary <> .Cells(Me.ColEmp_Basic.Index).Value Then
                                        strPerubahanEMP &= vbNewLine & "Basic salary from " & _
                                                                    rowEMP.Basic_Salary & " to " & .Cells(Me.ColEmp_Basic.Index).Value
                                    End If
                                End If

                                If rowEMP.Tax_Calculation <> CboTaxCal.Items.IndexOf(Me.dgvEmployee.Item(Me.ColEmp_TaxCalculation.Index, RowDGVE.Index).Value) Then
                                    strPerubahanEMP &= vbNewLine & "Tax calculation from " & _
                                                      CboTaxCal.Items(rowEMP.Tax_Calculation) & " to " & _
                                                      Me.dgvEmployee.Item(Me.ColEmp_TaxCalculation.Index, RowDGVE.Index).Value

                                End If

                                If rowEMP.Primary_Slip_Desc <> .Cells(Me.ColEmp_PrimarySlipDesc.Index).Value Then
                                    strPerubahanEMP &= vbNewLine & "Slip desc from " & _
                                                       Me.ObjSlipSalary.GetCodeByPrimary(rowEMP.Primary_Slip_Desc) & " to " & _
                                                       Me.ObjSlipSalary.GetCodeByPrimary(.Cells(Me.ColEmp_PrimarySlipDesc.Index).Value)
                                End If

                                If rowEMP.Primary_Expense_Account <> .Cells(Me.ColEmp_PrimaryExpenseAcc.Index).Value Then
                                    strPerubahanEMP &= vbNewLine & "Expense account from " & _
                                                       Me.ObjAcc.GetAccountCode(rowEMP.Primary_Expense_Account) & " to " & _
                                                       .Cells(Me.ColEmp_ExpenseAcc.Index).Value
                                End If

                                If rowEMP.Primary_Default_Payment_Account <> .Cells(Me.ColEmp_PrimaryPayment.Index).Value Then
                                    strPerubahanEMP &= vbNewLine & "Payment account from " & _
                                                       Me.ObjAcc.GetAccountCode(rowEMP.Primary_Default_Payment_Account) & " to " & _
                                                       .Cells(Me.ColEmp_PaymentAcc.Index).Value
                                End If
                            Next

                            If strPerubahanEMP <> "" Then
                                Dim strEmployee As String = ""
                                'Me.tblEMP.DefaultView.RowFilter = Nothing
                                'Me.tblEMP.DefaultView.RowFilter = "[Primary] = " & .Cells(Me.ColEmp_Primary.Index).Value
                                'If Me.tblEMP.DefaultView.Count > 0 Then
                                '    rowEMP = Me.tblEMP.DefaultView(0).Row
                                '    strEmployee = rowEMP.Employee_Name
                                'End If
                                .Cells(Me.ColEmp_IsModify.Index).Value = True
                                Me.rowLogUser = Me.tblLogUser.Newtbl_log_userRow
                                Me.rowLogUser.Primary = -1
                                Me.rowLogUser.Activity_Type = ClsLogUser.EnumActivityType.Modify
                                Me.rowLogUser.User = Me.mUserCode
                                Me.rowLogUser.Activity_Time = Now
                                Me.rowLogUser.Remarks = "Modify data for employee " & strEmployee & " at : " & strPerubahanEMP
                                Me.rowLogUser.Primary_Salary_Formula = Me.mPrimarySalary
                                Me.rowLogUser.Salary_Formula_Code = Me.ObjSalaryFormula.GetCodeByPrimary(Me.mPrimarySalary)
                                Me.tblLogUser.Rows.Add(Me.rowLogUser)
                            End If
                        End If
                    End If
                End With
            Next

            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            'TAB EMPLOYEE Selesai()
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub SalinKeComponent1(ByVal vFrm As FrmMasterSalaryFormula_Component1, _
    ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim CboDgv As DataGridViewComboBoxCell
        Dim RowDgv As DataGridViewRow
        Dim Filter As String = ""
        Dim TipeKompetensi As Core_Module.ClsMasterKompetensi.EnumDataType = ClsMasterKompetensi.EnumDataType.Alfabetic
        Dim dtCreatedDate As Date = Now
        Dim dtModifiedDate As Date = Now

        Try
            'header
            'primary transaksi 
            vFrm.mPrimaryTransaksi = Me.mPrimarySalary
            'identifier
            vFrm.mIdentifiyComp = Me.DgvComponent.Item(Me.ColIncome_Identifier.Index, e.RowIndex).Value
            'tbl variable
            vFrm.mTblVar = Me.TblVar
            'tbl contraint
            vFrm.mTblConst = Me.TblCons
            'clscekdata
            vFrm.Objcd = Me.ObjCd
            'UserCode
            vFrm.mUserCode = Me.mUserCode


            'vFrm.ColVariable_Primary.Visible = True
            'vFrm.ColConst_Primary.Visible = True




            If IsDate(Me.DgvComponent.Item(Me.ColIncome_CreatedDate.Index, e.RowIndex).Value) Then
                dtCreatedDate = CDate(Me.DgvComponent.Item(Me.ColIncome_CreatedDate.Index, e.RowIndex).Value)
            End If

            If IsDate(Me.DgvComponent.Item(Me.ColIncome_ModifiedDate.Index, e.RowIndex).Value) Then
                dtModifiedDate = CDate(Me.DgvComponent.Item(Me.ColIncome_ModifiedDate.Index, e.RowIndex).Value)
            End If

            vFrm.tslCreatedBy.Text = "Created By: " & Me.DgvComponent.Item(Me.ColIncome_CreatedBy.Index, e.RowIndex).Value
            vFrm.tslModifiedBy.Text = "Modified By: " & Me.DgvComponent.Item(Me.ColIncome_ModifiedBy.Index, e.RowIndex).Value
            vFrm.tslCreatedDate.Text = "Created Date: " & Format(dtCreatedDate, "dd-MMM-yyyy HH:mm")
            vFrm.tslModifiedDate.Text = "Modified Date: " & Format(dtModifiedDate, "dd-MMM-yyyy HH:mm")

            'formula
            vFrm.TxtFormula.Text = Me.DgvComponent.Item(Me.ColFormula.Index, e.RowIndex).Value

            'Calculate On
            CboDgv = Me.DgvComponent.Item(Me.ColIncome_CalculateOn.Index, e.RowIndex)
            vFrm.DtpYear.Value = (CboDgv.Items.IndexOf(Me.DgvComponent.Item(Me.ColIncome_CalculateOn.Index, e.RowIndex).Value) + 1).ToString + "/1/1950"

            'Calculate Period
            CboDgv = Me.DgvComponent.Item(Me.ColIncome_CalculatePeriod.Index, e.RowIndex)
            vFrm.mType = CboDgv.Items.IndexOf(Me.DgvComponent.Item(Me.ColIncome_CalculatePeriod.Index, e.RowIndex).Value)
            vFrm.Buffer_CalculatePeriod = Me.DgvComponent.Item(Me.ColIncome_CalculatePeriod.Index, e.RowIndex).Value.ToString

            'rounding unit
            If Me.DgvComponent.Item(Me.ColRoundingUnit.Index, e.RowIndex).Value.ToString.IndexOf(".") <> -1 _
            And Me.DgvComponent.Item(Me.ColRoundingUnit.Index, e.RowIndex).Value.ToString.IndexOf(".") <> 1 Then
                vFrm.CboRoundingUnit.Text = Me.DgvComponent.Item(Me.ColRoundingUnit.Index, e.RowIndex).Value.ToString.Remove(Me.DgvComponent.Item(Me.ColRoundingUnit.Index, e.RowIndex).Value.ToString.IndexOf("."))
            Else
                vFrm.CboRoundingUnit.Text = Me.DgvComponent.Item(Me.ColRoundingUnit.Index, e.RowIndex).Value
            End If
            If vFrm.CboRoundingUnit.Text = "" Then
                vFrm.CboRoundingUnit.Text = 1
            End If
            'rounding type
            vFrm.CboRoundingType.SelectedIndex = Me.DgvComponent.Item(Me.ColRoundingType.Index, e.RowIndex).Value
            'year max val
            vFrm.TxtYearlyMaxValue.Text = Me.DgvComponent.Item(Me.ColYearlyMaxValue.Index, e.RowIndex).Value
            'year min val
            vFrm.TxtYearlyMinValue.Text = Me.DgvComponent.Item(Me.ColYearlyMinValue.Index, e.RowIndex).Value
            'month max val
            vFrm.TxtMonthlyMaxValue.Text = Me.DgvComponent.Item(Me.ColMonthlyMaxValue.Index, e.RowIndex).Value
            'month min val
            vFrm.TxtMonthlyMinValue.Text = Me.DgvComponent.Item(Me.ColMonthlyMinValue.Index, e.RowIndex).Value

            'item
            'pindahin ke dgv variable
            vFrm.DGVVariable.Rows.Clear()
            Dim IdxVar As Integer = 0
            Filter = "[Primary Transaksi] = " + Me.mPrimarySalary.ToString + _
            "AND [Identifier Salary Component] = '" + Me.DgvComponent.Item(Me.ColIncome_Identifier.Index, e.RowIndex).Value.ToString + "'"
            Me.TblVar.DefaultView.RowFilter = Filter
            If Me.TblVar.DefaultView.Count > 0 Then
                For Each rowView As DataRowView In Me.TblVar.DefaultView
                    RowVar = Me.TblVar.DefaultView(IdxVar).Row
                    vFrm.DGVVariable.Rows.Add()
                    RowDgv = vFrm.DGVVariable.Rows(vFrm.DGVVariable.RowCount - 2)
                    If RowVar.Type2 = Core_Module.ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode Then
                        RowDgv.Cells(vFrm.ColConst_Type2.Index).Value = "Varible List"
                        'parameter code
                        RowDgv.Cells(vFrm.ColVariable_Parameter_Code.Index).Value = Me.ObjParameter.GetCodeByPrimary(RowVar.Primary_Parameter)
                        'parameter desc
                        RowDgv.Cells(vFrm.ColVariable_Parameter_Desc.Index).Value = Me.ObjParameter.GetDescByPrimary(RowVar.Primary_Parameter)
                        'type
                        CboDgv = RowDgv.Cells(vFrm.ColVariable_Type.Index)
                        RowDgv.Cells(vFrm.ColVariable_Type.Index).Value = CboDgv.Items.Item(RowVar.Type)
                        'unit
                        RowDgv.Cells(vFrm.ColVariable_Unit.Index).Value = Me.ObjParameter.GetUnitByPrimary(RowVar.Primary_Parameter)
                        'variable
                        RowDgv.Cells(vFrm.ColVariable_Variable.Index).Value = RowVar.Variable
                        'rounding unit
                        RowDgv.Cells(vFrm.ColRoundingUnit_Variable.Index).Value = Me.ObjCd.BuangNolDibelakangDecimal(Me.ObjCd.DigitGrouping(RowVar.Rounding_Unit))
                        'rounding type
                        CboDgv = RowDgv.Cells(vFrm.ColRoundingTipe_Variable.Index)
                        RowDgv.Cells(vFrm.ColRoundingTipe_Variable.Index).Value = CboDgv.Items.Item(RowVar.Rounding_Type)
                        'primary parameter
                        RowDgv.Cells(vFrm.ColVariable_PrimaryParameter.Index).Value = RowVar.Primary_Parameter
                        'primary 
                        RowDgv.Cells(vFrm.ColVariable_Primary.Index).Value = RowVar.Primary
                        IdxVar += 1
                    ElseIf RowVar.Type2 = Core_Module.ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi Then
                        RowDgv.Cells(vFrm.ColConst_Type2.Index).Value = "Competence"
                        'parameter code
                        RowDgv.Cells(vFrm.ColVariable_Parameter_Code.Index).Value = Me.ObjCompetence.GetDataCode(RowVar.Primary_Parameter)
                        'parameter desc
                        RowDgv.Cells(vFrm.ColVariable_Parameter_Desc.Index).Value = Me.ObjCompetence.GetDesc(RowVar.Primary_Parameter)
                        'type
                        CboDgv = RowDgv.Cells(vFrm.ColVariable_Type.Index)
                        RowDgv.Cells(vFrm.ColVariable_Type.Index).Value = CboDgv.Items.Item(RowVar.Type)
                        'unit
                        RowDgv.Cells(vFrm.ColVariable_Unit.Index).Value = ""
                        'variable
                        RowDgv.Cells(vFrm.ColVariable_Variable.Index).Value = RowVar.Variable
                        'rounding unit
                        RowDgv.Cells(vFrm.ColRoundingUnit_Variable.Index).Value = Me.ObjCd.BuangNolDibelakangDecimal(Me.ObjCd.DigitGrouping(RowVar.Rounding_Unit))
                        'rounding type
                        CboDgv = RowDgv.Cells(vFrm.ColRoundingTipe_Variable.Index)
                        RowDgv.Cells(vFrm.ColRoundingTipe_Variable.Index).Value = CboDgv.Items.Item(RowVar.Rounding_Type)
                        'primary parameter
                        RowDgv.Cells(vFrm.ColVariable_PrimaryParameter.Index).Value = RowVar.Primary_Parameter
                        'primary 
                        RowDgv.Cells(vFrm.ColVariable_Primary.Index).Value = RowVar.Primary
                        IdxVar += 1
                    End If
                Next
            End If


            'pindahin ke dgv constraint
            Dim IdxCons As Integer = 0
            Filter = "[Primary Transaksi] = " + Me.mPrimarySalary.ToString + _
            "AND [Identifier Salary Component] = '" + Me.DgvComponent.Item(Me.ColIncome_Identifier.Index, e.RowIndex).Value.ToString + "'"
            Me.TblCons.DefaultView.RowFilter = Filter
            If Me.TblCons.DefaultView.Count > 0 Then
                For Each rowView As DataRowView In Me.TblCons.DefaultView
                    RowCons = Me.TblCons.DefaultView(IdxCons).Row
                    vFrm.DGVConstaint.Rows.Add()
                    RowDgv = vFrm.DGVConstaint.Rows(vFrm.DGVConstaint.RowCount - 2)
                    If RowCons.Type2 = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode Then
                        RowDgv.Cells(vFrm.ColConst_Type2.Index).Value = "Constraint List"
                        'parameter code
                        RowDgv.Cells(vFrm.ColConst_ParameterCode.Index).Value = Me.ObjParameter.GetCodeByPrimary(RowCons.Primary_Parameter)
                        'parameter desc
                        RowDgv.Cells(vFrm.ColConst_ParameterDesc.Index).Value = Me.ObjParameter.GetDescByPrimary(RowCons.Primary_Parameter)
                        'type
                        CboDgv = RowDgv.Cells(vFrm.ColConst_Type.Index)
                        RowDgv.Cells(vFrm.ColConst_Type.Index).Value = CboDgv.Items.Item(RowCons.Type)
                        'compare
                        CboDgv = RowDgv.Cells(vFrm.ColConst_Compare.Index)
                        RowDgv.Cells(vFrm.ColConst_Compare.Index).Value = CboDgv.Items.Item(RowCons.Compare)
                        'ref value
                        If RowCons.Primary_Parameter = 131 Then
                            'Khusus Parameter Job Position 
                            RowDgv.Cells(vFrm.ColPrimaryRefValue.Index).Value = RowCons.Ref_Value
                            RowDgv.Cells(vFrm.ColConst_RefValue.Index).Value = Me.ObjJP.GetCode(RowCons.Ref_Value).ToString
                        ElseIf RowCons.Primary_Parameter = 133 Then
                            'Khusus Parameter Leave Code 
                            RowDgv.Cells(vFrm.ColPrimaryRefValue.Index).Value = RowCons.Ref_Value
                            RowDgv.Cells(vFrm.ColConst_RefValue.Index).Value = Me.ObjCuti.GetCodeByPrimary(RowCons.Ref_Value).ToString
                        Else
                            RowDgv.Cells(vFrm.ColConst_RefValue.Index).Value = RowCons.Ref_Value
                        End If
                        'unit
                        RowDgv.Cells(vFrm.ColConst_Unit.Index).Value = Me.ObjParameter.GetUnitByPrimary(RowCons.Primary_Parameter)
                        'primary parameter
                        RowDgv.Cells(vFrm.ColConst_PrimaryParameter.Index).Value = RowCons.Primary_Parameter
                        'primary 
                        RowDgv.Cells(vFrm.ColConst_Primary.Index).Value = RowCons.Primary
                    ElseIf RowCons.Type2 = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi Then
                        RowDgv.Cells(vFrm.ColConst_Type2.Index).Value = "Competence"

                        'parameter code
                        RowDgv.Cells(vFrm.ColConst_ParameterCode.Index).Value = Me.ObjCompetence.GetDataCode(RowCons.Primary_Parameter)
                        'parameter desc
                        RowDgv.Cells(vFrm.ColConst_ParameterDesc.Index).Value = Me.ObjCompetence.GetDesc(RowCons.Primary_Parameter)
                        'type
                        CboDgv = RowDgv.Cells(vFrm.ColConst_Type.Index)
                        RowDgv.Cells(vFrm.ColConst_Type.Index).Value = CboDgv.Items.Item(RowCons.Type)
                        'compare
                        CboDgv = RowDgv.Cells(vFrm.ColConst_Compare.Index)
                        RowDgv.Cells(vFrm.ColConst_Compare.Index).Value = CboDgv.Items.Item(RowCons.Compare)
                        'ref value
                        TipeKompetensi = Me.ObjCompetence.GetDataType(RowCons.Primary_Parameter)
                        If TipeKompetensi = ClsMasterKompetensi.EnumDataType.Alfabetic Then
                            Dim DataAplhabetic As Core_Module.ClsMasterKompetensi.Aphabetic = ClsMasterKompetensi.Aphabetic.A
                            DataAplhabetic = RowCons.Ref_Value
                            If DataAplhabetic = ClsMasterKompetensi.Aphabetic.A Then
                                RowDgv.Cells(vFrm.ColConst_RefValue.Index).Value = "A"
                            ElseIf DataAplhabetic = ClsMasterKompetensi.Aphabetic.B Then
                                RowDgv.Cells(vFrm.ColConst_RefValue.Index).Value = "B"
                            ElseIf DataAplhabetic = ClsMasterKompetensi.Aphabetic.C Then
                                RowDgv.Cells(vFrm.ColConst_RefValue.Index).Value = "C"
                            ElseIf DataAplhabetic = ClsMasterKompetensi.Aphabetic.D Then
                                RowDgv.Cells(vFrm.ColConst_RefValue.Index).Value = "D"
                            ElseIf DataAplhabetic = ClsMasterKompetensi.Aphabetic.E Then
                                RowDgv.Cells(vFrm.ColConst_RefValue.Index).Value = "E"
                            ElseIf DataAplhabetic = ClsMasterKompetensi.Aphabetic.F Then
                                RowDgv.Cells(vFrm.ColConst_RefValue.Index).Value = "F"
                            End If
                            RowDgv.Cells(vFrm.ColPrimaryRefValue.Index).Value = CInt(DataAplhabetic)
                        ElseIf TipeKompetensi = ClsMasterKompetensi.EnumDataType.Number Then
                            RowDgv.Cells(vFrm.ColConst_RefValue.Index).Value = Me.ObjCd.DigitGrouping(RowCons.Ref_Value, _
                            Me.ObjCompetence.GetDefPrec(RowCons.Primary_Parameter))
                        ElseIf TipeKompetensi = ClsMasterKompetensi.EnumDataType.Tanggal Then
                            RowDgv.Cells(vFrm.ColConst_RefValue.Index).Value = Format(CDate(RowCons.Ref_Value), "dd MMM yyyy")
                        ElseIf TipeKompetensi = ClsMasterKompetensi.EnumDataType.Text Then
                            RowDgv.Cells(vFrm.ColConst_RefValue.Index).Value = RowCons.Ref_Value
                        ElseIf TipeKompetensi = ClsMasterKompetensi.EnumDataType.YN Then
                            If RowCons.Ref_Value = 0 Then
                                RowDgv.Cells(vFrm.ColConst_RefValue.Index).Value = "No"
                            Else
                                RowDgv.Cells(vFrm.ColConst_RefValue.Index).Value = "Yes"
                            End If
                        End If
                        'unit
                        RowDgv.Cells(vFrm.ColConst_Unit.Index).Value = ""
                        'primary parameter
                        RowDgv.Cells(vFrm.ColConst_PrimaryParameter.Index).Value = RowCons.Primary_Parameter
                        'primary 
                        RowDgv.Cells(vFrm.ColConst_Primary.Index).Value = RowCons.Primary
                    ElseIf RowCons.Type2 = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.CustomValue Then
                        RowDgv.Cells(vFrm.ColConst_Type2.Index).Value = "Custom Value"
                    ElseIf RowCons.Type2 = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.ServiceRecord Then
                        RowDgv.Cells(vFrm.ColConst_Type2.Index).Value = "Service Record"
                    End If

                    IdxCons += 1
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeComponent(ByVal vFrm As FrmMasterSalaryFormula_Component1, _
    ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim CboDgv As DataGridViewComboBoxCell
        Try
            'formula
            Me.DgvComponent.Item(Me.ColFormula.Index, e.RowIndex).Value = vFrm.TxtFormula.Text
            'calculate period
            CboDgv = Me.DgvComponent.Item(Me.ColIncome_CalculatePeriod.Index, e.RowIndex)
            Me.DgvComponent.Item(Me.ColIncome_CalculatePeriod.Index, e.RowIndex).Value = CboDgv.Items.Item(vFrm.mType)
            'calculate on
            CboDgv = Me.DgvComponent.Item(Me.ColIncome_CalculateOn.Index, e.RowIndex)
            Me.DgvComponent.Item(Me.ColIncome_CalculateOn.Index, e.RowIndex).Value = _
            CboDgv.Items.Item(vFrm.DtpYear.Value.Month - 1)
            'rounding unit
            Me.DgvComponent.Item(Me.ColRoundingUnit.Index, e.RowIndex).Value = vFrm.CboRoundingUnit.Text
            'rounding type
            Me.DgvComponent.Item(Me.ColRoundingType.Index, e.RowIndex).Value = vFrm.CboRoundingType.SelectedIndex
            'yearly max val
            Me.DgvComponent.Item(Me.ColYearlyMaxValue.Index, e.RowIndex).Value = vFrm.TxtYearlyMaxValue.Text
            'yearly min val
            Me.DgvComponent.Item(Me.ColYearlyMinValue.Index, e.RowIndex).Value = vFrm.TxtYearlyMinValue.Text
            'mounthly max val
            Me.DgvComponent.Item(Me.ColMonthlyMaxValue.Index, e.RowIndex).Value = vFrm.TxtMonthlyMaxValue.Text
            'mounthly min val
            Me.DgvComponent.Item(Me.ColMonthlyMinValue.Index, e.RowIndex).Value = vFrm.TxtMonthlyMinValue.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemComp(ByVal vData As String, ByVal vIndexColom As Integer, ByVal vIndexRow As Integer) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.DgvComponent.Rows
                If String.IsNullOrEmpty(rowDGV.Cells(vIndexColom).Value) = False Then
                    If rowDGV.Cells(vIndexColom).Value.ToString.ToUpper = vData.ToUpper Then
                        If rowDGV.Index <> vIndexRow Then
                            Return rowDGV.Index
                        End If
                    End If
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Function GetIndex(ByVal vComponentName As String) As Integer
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.DgvComponent.Rows
                If rowDGV.Index <> Me.DgvComponent.Rows.Count - 1 Then
                    If rowDGV.Cells(Me.ColIncome_Desc.Index).Value = vComponentName Then
                        Return rowDGV.Index
                    End If
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetIndexItemEmp(ByVal vData As String, ByVal vIndexColom As Integer, ByVal vIndexRow As Integer) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.dgvEmployee.Rows
                If String.IsNullOrEmpty(rowDGV.Cells(vIndexColom).Value) = False Then
                    If rowDGV.Cells(vIndexColom).Value.ToString.ToUpper = vData.ToUpper Then
                        If rowDGV.Index <> vIndexRow Then
                            Return rowDGV.Index
                        End If
                    End If
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub DgvComponent_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvComponent.RowsAdded
        'Dim CboDgv As DataGridViewComboBoxCell
        Dim RowIndexMax As Integer = 0
        Dim MAX As Integer = 0

        Try
            If Me.BackgroundWorker2.IsBusy = False Then
                Me.BackgroundWorker2.RunWorkerAsync()
                Me.BackgroundWorker2.Dispose()
            End If

            If Me.SedangIsiDGV = False Then
                MAX = Me.GetMaxUrutan() + 1
                Me.DgvComponent.Item(Me.ColIncome_Identifier.Index, e.RowIndex).Value = CInt(Me.mPrimarySalary).ToString + "_" + MAX.ToString
                Me.DgvComponent.Item(Me.ColIncome_Urutan.Index, e.RowIndex).Value = MAX.ToString
                Me.MAXUrutan = CInt(Me.DgvComponent.Item(Me.ColIncome_Urutan.Index, e.RowIndex).Value)
            End If

            If Me.DgvComponent.Rows.Count <= 1 Then
                Me.DgvComponent.Item(Me.ColIncome_CreatedBy.Index, e.RowIndex).Value = Me.mUserCode
                Me.DgvComponent.Item(Me.ColIncome_ModifiedBy.Index, e.RowIndex).Value = Me.mUserCode
                Me.DgvComponent.Item(Me.ColIncome_CreatedDate.Index, e.RowIndex).Value = Format(Now, "dd-MMM-yyyy") + " " + Now.ToShortTimeString
                Me.DgvComponent.Item(Me.ColIncome_ModifiedDate.Index, e.RowIndex).Value = Format(Now, "dd-MMM-yyyy") + " " + Now.ToShortTimeString
            Else
                'If e.RowIndex <> 0 Then
                '    Me.DgvComponent.Item(Me.ColIncome_Identifier.Index, e.RowIndex).Value = _
                '    CInt(Me.mPrimarySalary).ToString + "_" + (Me.DgvComponent.Item(Me.ColIncome_Urutan.Index, e.RowIndex - 1).Value + 1).ToString
                '    Me.DgvComponent.Item(Me.ColIncome_Urutan.Index, e.RowIndex).Value = _
                '    (Me.DgvComponent.Item(Me.ColIncome_Urutan.Index, e.RowIndex - 1).Value + 1).ToString
                'End If
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetMaxUrutan() As Double
        Dim rowDGV As DataGridViewRow
        Dim RetVal As Double = 0

        Try
            For Each rowDGV In Me.DgvComponent.Rows
                If CDbl(rowDGV.Cells(Me.ColIncome_Urutan.Index).Value) > RetVal Then
                    RetVal = rowDGV.Cells(Me.ColIncome_Urutan.Index).Value
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DgvComponent_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvComponent.RowsRemoved
        Dim RowIndexMax As Integer = 0
        Dim MAX As Integer = 0
        Try
            If Me.BackgroundWorker2.IsBusy = False Then
                Me.BackgroundWorker2.Dispose()
                Me.BackgroundWorker2.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvComponent_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvComponent.CellBeginEdit
        Dim CboDgv As DataGridViewComboBoxCell
        Try
            If Me.HakAksesUser_SalaryFormula(Me.mUserCode) = False Then
                e.Cancel = True
                Exit Sub
            End If

            If e.ColumnIndex = Me.ColIncome_ComponentTipe.Index Or e.ColumnIndex = Me.ColIncome_SlipDescription.Index _
            Or e.ColumnIndex = Me.ColIncome_TaxComponent.Index Or e.ColumnIndex = Me.ColIncome_ExpenseAccount.Index _
            Or e.ColumnIndex = Me.ColFormula.Index Or e.ColumnIndex = Me.ColEditFormula.Index Or e.ColumnIndex = Me.ColIncome_ComponentTipe.Index _
            Or e.ColumnIndex = Me.ColIncome_CalculateOn.Index Or e.ColumnIndex = Me.ColIncome_CalculatePeriod.Index _
            Or e.ColumnIndex = Me.ColIncome_PaymentAccount.Index Or e.ColumnIndex = Me.ColIncome_Routine.Index Or _
            e.ColumnIndex = Me.ColIncome_Desc.Index Then
                If e.RowIndex <> Me.DgvComponent.Rows.Count - 1 Then
                    If String.IsNullOrEmpty(Me.DgvComponent.Item(Me.ColIncome_Desc.Index, e.RowIndex).Value) = True Then
                        e.Cancel = True
                    Else
                        If e.ColumnIndex = Me.ColIncome_CalculatePeriod.Index Then
                            'batasan tambahan jika calculate period: jika sudah ada variable dan constraint tidak bisa diedit, karena
                            'nanti ketuka antara variable dan constrait yang monthly, daily, dan yearly
                            If Me.GetJmlConst(Me.DgvComponent.Item(Me.ColIncome_Identifier.Index, e.RowIndex).Value) > 0 Or _
                            Me.GetJmlVar(Me.DgvComponent.Item(Me.ColIncome_Identifier.Index, e.RowIndex).Value) > 0 Then
                                MsgBox("Can't change this data because the formula used is using variable or constraint.", MsgBoxStyle.Information, "Information")
                                e.Cancel = True
                            Else
                                Me.Buffer = Me.DgvComponent(e.ColumnIndex, e.RowIndex).Value
                            End If
                        ElseIf e.ColumnIndex = Me.ColIncome_TaxComponent.Index Then
                            'CboDgv = Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex)
                            'If CboDgv.Items.IndexOf(Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex).Value) = EnumType.CompanyExpense Or _
                            'CboDgv.Items.IndexOf(Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex).Value) = EnumType.EmployeeExpense Then
                            '    MsgBox("Can't change tax data for the following salary component type: company expense/employee expense.", MsgBoxStyle.Information, "Information")
                            '    e.Cancel = True
                            'End If
                            Me.Buffer = Me.DgvComponent(e.ColumnIndex, e.RowIndex).Value
                        ElseIf e.ColumnIndex = Me.ColIncome_Routine.Index Then
                            'yang boleh di edit hanya yang tipe akun nya income 
                            'dedution harus rutin karena jika deduction non-rutin, maka ada kemungkinan tunj. PPH non rutin minus.
                            'karena nilai pendapatan rutin + non rutin < pendapatan rutin
                            'sudah di coba ketika simulasi CLS Des 2018
                            CboDgv = Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex)
                            If Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex).Value = CboDgv.Items.Item(EnumType.Income) Then
                                Me.Buffer = Me.DgvComponent(e.ColumnIndex, e.RowIndex).Value
                            Else
                                e.Cancel = True
                            End If
                        ElseIf e.ColumnIndex = Me.ColIncome_ExpenseAccount.Index Then
                            CboDgv = Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex)
                            If CboDgv.Items.IndexOf(Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex).Value) = EnumType.EmployeeExpense Then
                                MsgBox("Can't change expense account for the following salary component type: employee expense.", MsgBoxStyle.Information, "Information")
                                e.Cancel = True
                            Else
                                Me.Buffer = Me.DgvComponent(e.ColumnIndex, e.RowIndex).Value
                            End If
                        Else
                            Me.Buffer = Me.DgvComponent(e.ColumnIndex, e.RowIndex).Value
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColIncome_Desc.Index = True Then
                        Me.Buffer = Me.DgvComponent(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                End If
            End If

            If e.ColumnIndex = Me.ColIncome_CalculateOn.Index Then
                CboDgv = Me.DgvComponent.Item(Me.ColIncome_CalculatePeriod.Index, e.RowIndex)
                If Me.DgvComponent.Item(Me.ColIncome_CalculatePeriod.Index, e.RowIndex).Value = CboDgv.Items.Item(EnumCalculatePeriod.Daily) Or _
                Me.DgvComponent.Item(Me.ColIncome_CalculatePeriod.Index, e.RowIndex).Value = CboDgv.Items.Item(EnumCalculatePeriod.monthly) Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Information")
        End Try
    End Sub

    Private Sub DgvComponent_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvComponent.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvComponent.Rows(e.RowIndex)
        Dim CboDgv As DataGridViewComboBoxCell
        Dim PrimaryAcc As Double = -1
        Dim PrimarySlip As Double = -1
        Dim TipeAkun As Integer = -1
        Dim hasil As DialogResult
        Dim index As Integer = -1
        Try
            'kolom type
            If e.ColumnIndex = Me.ColIncome_ComponentTipe.Index Then
                CboDgv = Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex)
                If CboDgv.Items.IndexOf(Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value) = EnumType.Income Then
                    CboDgv = Me.DgvComponent.Item(Me.ColIncome_TaxComponent.Index, e.RowIndex)
                    If CboDgv.Items.Contains("IURAN PENSIUN ATAU IURAN THT/JHT") = True Then
                        Me.DgvComponent.Item(Me.ColIncome_TaxComponent.Index, e.RowIndex).Value = CboDgv.Items.Item(0)
                    End If
                    CboDgv.Items.Remove("IURAN PENSIUN ATAU IURAN THT/JHT")

                    If CboDgv.Items.Contains("PESANGON") = False Then
                        CboDgv.Items.Insert(CboDgv.Items.Count, "PESANGON")
                    End If
                ElseIf CboDgv.Items.IndexOf(Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value) = EnumType.Deduction Or _
                CboDgv.Items.IndexOf(Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value) = EnumType.CompanyExpense Or _
                CboDgv.Items.IndexOf(Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value) = EnumType.EmployeeExpense Then
                    CboDgv = Me.DgvComponent.Item(Me.ColIncome_TaxComponent.Index, e.RowIndex)
                    If CboDgv.Items.Contains("IURAN PENSIUN ATAU IURAN THT/JHT") = False Then
                        CboDgv.Items.Insert(6, "IURAN PENSIUN ATAU IURAN THT/JHT")
                    End If
                    If CboDgv.Items.Contains("PESANGON") = True Then
                        Me.DgvComponent.Item(Me.ColIncome_TaxComponent.Index, e.RowIndex).Value = CboDgv.Items.Item(0)
                        CboDgv.Items.Remove("PESANGON")
                    End If
                End If

                CboDgv = Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex)
                If CboDgv.Items.IndexOf(Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value) = EnumType.CompanyExpense Then
                    'salary slip desc boleh kosong
                    Me.DgvComponent.Item(Me.ColIncome_SlipDescription.Index, e.RowIndex).Value = ""
                    Me.DgvComponent.Item(Me.ColIncome_PrimarySlipDesc.Index, e.RowIndex).Value = "-1"
                    CboDgv = Me.DgvComponent.Item(Me.ColIncome_TaxComponent.Index, e.RowIndex)
                    Me.DgvComponent.Item(Me.ColIncome_TaxComponent.Index, e.RowIndex).Value = "PREMI ASURANSI YANG DIBAYARKAN PEMBERI KERJA"
                    'tipe data-nya harus rutin
                    Me.DgvComponent.Item(Me.ColIncome_Routine.Index, e.RowIndex).Value = Boolean.TrueString
                ElseIf CboDgv.Items.IndexOf(Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value) = EnumType.EmployeeExpense Then
                    Me.DgvComponent.Item(Me.ColIncome_TaxComponent.Index, e.RowIndex).Value = "IURAN PENSIUN ATAU IURAN THT/JHT"
                    Me.DgvComponent.Item(Me.ColIncome_ExpenseAccount.Index, e.RowIndex).Value = ""
                    Me.DgvComponent.Item(Me.ColIncome_PrimaryExpenseAccount.Index, e.RowIndex).Value = "-1"
                    'tipe data-nya harus rutin
                    Me.DgvComponent.Item(Me.ColIncome_Routine.Index, e.RowIndex).Value = Boolean.TrueString
                Else
                    CboDgv = Me.DgvComponent.Item(Me.ColIncome_TaxComponent.Index, e.RowIndex)
                    Me.DgvComponent.Item(Me.ColIncome_TaxComponent.Index, e.RowIndex).Value = CboDgv.Items.Item(0)

                    'khusus untuk tipe data expense, harus rutin/annualized
                    CboDgv = Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex)
                    If CboDgv.Items.IndexOf(Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value) = EnumType.Deduction Then
                        'tipe data-nya harus rutin
                        Me.DgvComponent.Item(Me.ColIncome_Routine.Index, e.RowIndex).Value = Boolean.TrueString
                    End If
                End If
            End If

            'kolom deskripsi
            If e.ColumnIndex = Me.ColIncome_Desc.Index Then
                If String.IsNullOrEmpty(Trim(Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value)) = False Then
                    'harus unique, tidak boleh sama dengan baris yang lain
                    If Me.GetIndexItemComp(Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value, e.ColumnIndex, e.RowIndex) < 0 Then
                        Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = _
                        Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'cari identifier
                        If String.IsNullOrEmpty(Me.DgvComponent.Item(Me.ColIncome_Identifier.Index, e.RowIndex).Value) = True Then
                            If e.RowIndex <> 0 Then
                                Me.DgvComponent.Item(Me.ColIncome_Identifier.Index, e.RowIndex).Value = _
                                CInt(Me.mPrimarySalary).ToString + "_" + (Me.MAXUrutan + 1).ToString
                                Me.DgvComponent.Item(Me.ColIncome_Urutan.Index, e.RowIndex).Value = _
                                (Me.MAXUrutan + 1).ToString
                                Me.MAXUrutan = CInt(Me.DgvComponent.Item(Me.ColIncome_Urutan.Index, e.RowIndex).Value)
                            End If
                        End If
                    Else
                        MsgBox("Description has been used, please change the description.", MsgBoxStyle.Information, "Information")
                        Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        If Me.Buffer = "" Then
                            'Me.mDGV = Me.DgvComponent
                            'Me.EraseRowIndex = row.Index
                            'Me.BackgroundWorker1.RunWorkerAsync()
                            'Me.BackgroundWorker1.Dispose()
                        End If
                    End If
                ElseIf String.IsNullOrEmpty(Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value) = True OrElse Trim(Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value) = "" Then
                    If Me.Buffer = "" Then
                        '    MsgBox("Description cannot be null or empty string, please change the description.", MsgBoxStyle.Information, "Information")
                        Me.mDGV = Me.DgvComponent
                        Me.EraseRowIndex = row.Index
                        Me.BackgroundWorker1.RunWorkerAsync()
                        Me.BackgroundWorker1.Dispose()
                    Else
                        Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            End If

            'kolom salary slip desc
            If e.ColumnIndex = Me.ColIncome_SlipDescription.Index Then
                CboDgv = Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex)
                If String.IsNullOrEmpty(Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.ObjSlipSalary.ApakahSudahAda(Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        PrimarySlip = Me.ObjSlipSalary.GetPrimaryByCode(Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value)
                        If PrimarySlip <> -1 Then
                            Me.DgvComponent.Item(Me.ColIncome_PrimarySlipDesc.Index, e.RowIndex).Value = PrimarySlip
                            Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjSlipSalary.GetCodeByPrimary(PrimarySlip)
                            If CboDgv.Items.IndexOf(Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex).Value) = EnumType.Deduction Or _
                            CboDgv.Items.IndexOf(Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex).Value) = EnumType.Income Then
                                Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex).Value = CboDgv.Items.Item(Me.ObjSlipSalary.GetTypeByPrimary(PrimarySlip))
                            End If
                        Else
                            If CboDgv.Items.IndexOf(Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex).Value) = EnumType.Deduction Or _
                            CboDgv.Items.IndexOf(Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex).Value) = EnumType.Income Then
                                MsgBox("Salary slip description cannot be an empty string", MsgBoxStyle.Information, "Information")
                                Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                        End If
                    Else
                        If CboDgv.Items.IndexOf(Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value) = EnumType.CompanyExpense Then
                            'kalo tipenya company expense, maka boleh tidak muncul dilsary slip desc
                            Me.DgvComponent.Item(Me.ColIncome_SlipDescription.Index, e.RowIndex).Value = ""
                            Me.DgvComponent.Item(Me.ColIncome_PrimarySlipDesc.Index, e.RowIndex).Value = "-1"
                        Else
                            'diketik tidak ada lalu tampilkan form cari
                            Dim ij As New FrmMasterSalaryFormula_SlipDesc_Cari
                            ij.txtCari.Text = Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.Primary <> -1 Then
                                Me.DgvComponent.Item(Me.ColIncome_PrimarySlipDesc.Index, e.RowIndex).Value = ij.Primary
                                Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjSlipSalary.GetCodeByPrimary(ij.Primary)

                                If CboDgv.Items.IndexOf(Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex).Value) = EnumType.Deduction Or _
                                CboDgv.Items.IndexOf(Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex).Value) = EnumType.Income Then
                                    Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex).Value = CboDgv.Items.Item(Me.ObjSlipSalary.GetTypeByPrimary(ij.Primary))
                                End If
                            Else
                                Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                'Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Selected = True
                            End If
                            ij = Nothing
                        End If
                    End If
                Else
                    If CboDgv.Items.IndexOf(Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex).Value) = EnumType.CompanyExpense Then
                        Me.DgvComponent.Item(Me.ColIncome_SlipDescription.Index, e.RowIndex).Value = ""
                        Me.DgvComponent.Item(Me.ColIncome_PrimarySlipDesc.Index, e.RowIndex).Value = "-1"
                    Else
                        MsgBox("Salary slip description cannot be an empty string", MsgBoxStyle.Information, "Information")
                        Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            End If

            'kolom expense acc
            If e.ColumnIndex = Me.ColIncome_ExpenseAccount.Index Then
                If String.IsNullOrEmpty(Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.ObjAcc.ApakahSudahAda_Aktif_NonHeader(Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        PrimaryAcc = Me.ObjAcc.GetPrimary(Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value)
                        'cek akun
                        If PrimaryAcc <> -1 Then
                            TipeAkun = Me.ObjAcc.GetTipeAkun(PrimaryAcc)
                            'selain expense, other expense dan COGS kasih peringantan yes / no
                            If TipeAkun = ClsMasterAkun1.EnumTipeAkun.Expense Or TipeAkun = ClsMasterAkun1.EnumTipeAkun.Other_Expense Or _
                            TipeAkun = ClsMasterAkun1.EnumTipeAkun.Cost_of_Goods_Sales Then
                                Me.DgvComponent.Item(Me.ColIncome_PrimaryExpenseAccount.Index, e.RowIndex).Value = PrimaryAcc
                                Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAcc.GetAccountCode(PrimaryAcc)
                            Else
                                hasil = MsgBox("Account type is not expense, other expense or COGS." + vbNewLine + _
                                "Do you want to continue ?", MsgBoxStyle.YesNo, "Information")
                                If hasil = Windows.Forms.DialogResult.Yes Then
                                    Me.DgvComponent.Item(Me.ColIncome_PrimaryExpenseAccount.Index, e.RowIndex).Value = PrimaryAcc
                                    Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAcc.GetAccountCode(PrimaryAcc)
                                Else
                                    Me.DgvComponent.Item(Me.ColIncome_PrimaryExpenseAccount.Index, e.RowIndex).Value = -1
                                    Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                End If
                            End If
                        Else
                            CboDgv = Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex)
                            If CboDgv.Items.IndexOf(Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex).Value) = EnumType.EmployeeExpense Then
                                Me.DgvComponent.Item(Me.ColIncome_ExpenseAccount.Index, e.RowIndex).Value = ""
                                Me.DgvComponent.Item(Me.ColIncome_PrimaryExpenseAccount.Index, e.RowIndex).Value = "-1"
                            Else
                                MsgBox("Expense account cannot be an empty string", MsgBoxStyle.Information, "Information")
                                Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                        End If
                    Else
                        ''diketik tidak ada lalu tampilkan form cari
                        Dim ij As New FrmMasterAkun_Cari1
                        ij.txtCari.Text = Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.mPrimaryAkun <> -1 Then
                            'cek akun
                            TipeAkun = Me.ObjAcc.GetTipeAkun(ij.mPrimaryAkun)
                            'selain expense, other expense dan COGS kasih peringantan yes / no
                            If TipeAkun = ClsMasterAkun1.EnumTipeAkun.Expense Or TipeAkun = ClsMasterAkun1.EnumTipeAkun.Other_Expense Or _
                            TipeAkun = ClsMasterAkun1.EnumTipeAkun.Cost_of_Goods_Sales Then
                                Me.DgvComponent.Item(Me.ColIncome_PrimaryExpenseAccount.Index, e.RowIndex).Value = ij.mPrimaryAkun
                                Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAcc.GetAccountCode(ij.mPrimaryAkun)
                            Else
                                hasil = MsgBox("Account type is not expense, other expense or COGS." + vbNewLine + _
                                "Do you want to continue ?", MsgBoxStyle.YesNo, "Information")
                                If hasil = Windows.Forms.DialogResult.Yes Then
                                    Me.DgvComponent.Item(Me.ColIncome_PrimaryExpenseAccount.Index, e.RowIndex).Value = ij.mPrimaryAkun
                                    Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAcc.GetAccountCode(ij.mPrimaryAkun)
                                Else
                                    Me.DgvComponent.Item(Me.ColIncome_PrimaryExpenseAccount.Index, e.RowIndex).Value = -1
                                    Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                End If
                            End If
                        Else
                            CboDgv = Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex)
                            If CboDgv.Items.IndexOf(Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex).Value) = EnumType.EmployeeExpense Then
                                Me.DgvComponent.Item(Me.ColIncome_ExpenseAccount.Index, e.RowIndex).Value = ""
                                Me.DgvComponent.Item(Me.ColIncome_PrimaryExpenseAccount.Index, e.RowIndex).Value = "-1"
                            Else
                                MsgBox("Expense account cannot be an empty string", MsgBoxStyle.Information, "Information")
                                Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                        End If
                        ij = Nothing
                    End If
                Else
                    CboDgv = Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex)
                    If CboDgv.Items.IndexOf(Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex).Value) = EnumType.EmployeeExpense Then
                        Me.DgvComponent.Item(Me.ColIncome_ExpenseAccount.Index, e.RowIndex).Value = ""
                        Me.DgvComponent.Item(Me.ColIncome_PrimaryExpenseAccount.Index, e.RowIndex).Value = "-1"
                    Else
                        MsgBox("Expense account cannot be an empty string", MsgBoxStyle.Information, "Information")
                        Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            End If

            'kolom payment acc
            If e.ColumnIndex = Me.ColIncome_PaymentAccount.Index Then
                If String.IsNullOrEmpty(Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.ObjAcc.ApakahSudahAda_Aktif_NonHeader(Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        PrimaryAcc = Me.ObjAcc.GetPrimary(Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value)
                        'cek akun
                        If PrimaryAcc <> -1 Then
                            TipeAkun = Me.ObjAcc.GetTipeAkun(PrimaryAcc)
                            'selain payment, other payment dan COGS kasih peringantan yes / no
                            If TipeAkun = ClsMasterAkun1.EnumTipeAkun.Asset Or TipeAkun = ClsMasterAkun1.EnumTipeAkun.Liabilities Then
                                Me.DgvComponent.Item(Me.ColIncome_PrimaryPaymentAccount.Index, e.RowIndex).Value = PrimaryAcc
                                Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAcc.GetAccountCode(PrimaryAcc)
                            Else
                                hasil = MsgBox("Account type is not asset or liabilities." + vbNewLine + _
                                "Do you want to continue ?", MsgBoxStyle.YesNo, "Information")
                                If hasil = Windows.Forms.DialogResult.Yes Then
                                    Me.DgvComponent.Item(Me.ColIncome_PrimaryPaymentAccount.Index, e.RowIndex).Value = PrimaryAcc
                                    Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAcc.GetAccountCode(PrimaryAcc)
                                Else
                                    Me.DgvComponent.Item(Me.ColIncome_PrimaryPaymentAccount.Index, e.RowIndex).Value = -1
                                    Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                End If
                            End If
                        Else
                            Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                    Else
                        ''diketik tidak ada lalu tampilkan form cari
                        Dim ij As New FrmMasterAkun_Cari1
                        ij.txtCari.Text = Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.mPrimaryAkun <> -1 Then
                            'cek akun
                            TipeAkun = Me.ObjAcc.GetTipeAkun(ij.mPrimaryAkun)
                            'selain payment, other payment dan COGS kasih peringantan yes / no
                            If TipeAkun = ClsMasterAkun1.EnumTipeAkun.Asset Or TipeAkun = ClsMasterAkun1.EnumTipeAkun.Liabilities Then
                                Me.DgvComponent.Item(Me.ColIncome_PrimaryPaymentAccount.Index, e.RowIndex).Value = ij.mPrimaryAkun
                                Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAcc.GetAccountCode(ij.mPrimaryAkun)
                            Else
                                hasil = MsgBox("Account type is not asset or liabilies." + vbNewLine + _
                                "Do you want to continue ?", MsgBoxStyle.YesNo, "Information")
                                If hasil = Windows.Forms.DialogResult.Yes Then
                                    Me.DgvComponent.Item(Me.ColIncome_PrimaryPaymentAccount.Index, e.RowIndex).Value = ij.mPrimaryAkun
                                    Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAcc.GetAccountCode(ij.mPrimaryAkun)
                                Else
                                    Me.DgvComponent.Item(Me.ColIncome_PrimaryPaymentAccount.Index, e.RowIndex).Value = -1
                                    Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                End If
                            End If
                        Else
                            Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        ij = Nothing
                    End If
                Else
                    Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Me.DgvComponent.Item(Me.ColIncome_PrimaryPaymentAccount.Index, e.RowIndex).Value = -1
                End If
            End If

            'kolom formula
            If e.ColumnIndex = Me.ColFormula.Index Then
                If Me.ObjCd.CekNumerik(Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    If Me.ObjCd.CekNonNegatif(Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please input positive number.", MsgBoxStyle.Information, "Information")
                        Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
                Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
            End If

            If e.ColumnIndex = Me.ColIncome_CalculatePeriod.Index Then
                CboDgv = Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex)
                If Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = CboDgv.Items.Item(EnumCalculatePeriod.Daily) Or _
                Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = CboDgv.Items.Item(EnumCalculatePeriod.monthly) Then
                    CboDgv = Me.DgvComponent.Item(Me.ColIncome_CalculateOn.Index, e.RowIndex)
                    Me.DgvComponent.Item(Me.ColIncome_CalculateOn.Index, e.RowIndex).Value = CboDgv.Items.Item(0)
                End If
            End If

            If e.ColumnIndex = Me.ColIncome_TaxComponent.Index Then
                CboDgv = Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex)
                If CboDgv.Items.IndexOf(Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex).Value) = EnumType.CompanyExpense Then
                    If Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToLower <> ("Bukan Objek Pajak").ToLower And _
                    Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToLower <> ("PREMI ASURANSI YANG DIBAYARKAN PEMBERI KERJA").ToLower Then
                        MsgBox("Please specify a valid tax component. User either: 'PREMI ASURANSI YANG DIBAYARKAN PEMBERI KERJA' or 'BUKAN OBJEK PAJAK'", MsgBoxStyle.Information, "Information")
                        Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                ElseIf CboDgv.Items.IndexOf(Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, e.RowIndex).Value) = EnumType.EmployeeExpense Then
                    If Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToLower <> ("Bukan Objek Pajak").ToLower And _
                    Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToLower <> ("IURAN PENSIUN ATAU IURAN THT/JHT").ToLower Then
                        MsgBox("Please specify a valid tax component. User either: 'IURAN PENSIUN ATAU IURAN THT/JHT' or 'BUKAN OBJEK PAJAK'", MsgBoxStyle.Information, "Information")
                        Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Information")
        End Try
    End Sub

    Private Sub DgvComponent_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvComponent.CellContentClick
        Dim NewId As String = ""
        Dim rowDGVSumber, rowDGVTarget As DataGridViewRow
        Dim RowVarTarget, RowVarSumber As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_variable1Row
        Dim RowContTarget, RowContSumber As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_constraint1Row
        Dim DummyJml As Integer = 0
        Dim DummyDesc As String = ""

        Try
            If e.RowIndex = -1 Then
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColEditFormula.Index Then
                If String.IsNullOrEmpty(Me.DgvComponent(Me.ColIncome_Desc.Index, e.RowIndex).Value) = False Then
                    Dim vFrm As New FrmMasterSalaryFormula_Component1
                    If Me.HakAksesUser_SalaryFormula(Me.mUserCode, False) = False Then
                        vFrm.PunyaHakAksesFull = False
                    End If
                    'sesuaikan dari kolom ke header
                    Me.SalinKeComponent1(vFrm, e)
                    vFrm.BackColorAll = Me.BackColorAll
                    vFrm.mUserCode = Me.mUserCode
                    vFrm.ShowDialog()

                    'sesuaikan dari header ke koloms
                    If vFrm.salin = True Then
                        Me.SalinKeComponent(vFrm, e)
                    End If
                End If
            ElseIf e.ColumnIndex = Me.ColIncome_Duplicate.Index Then
                If String.IsNullOrEmpty(Me.DgvComponent(Me.ColIncome_Desc.Index, e.RowIndex).Value) = False Then
                    If Me.HakAksesUser_SalaryFormula(Me.mUserCode) = False Then
                        Exit Sub
                    End If
                    If MsgBox("Duplicate this salary component: " + Me.DgvComponent.Item(Me.ColIncome_Desc.Index, e.RowIndex).Value.ToString, _
                                               MsgBoxStyle.YesNo, "Information") = MsgBoxResult.Yes Then
                        'tambahkan baris baru
                        Me.SedangIsiDGV = True
                        rowDGVSumber = Me.DgvComponent.Rows(e.RowIndex)
                        Me.DgvComponent.Rows.Insert(rowDGVSumber.Index + 1, 1)
                        rowDGVTarget = Me.DgvComponent.Rows(e.RowIndex + 1)
                        NewId = rowDGVTarget.Cells(Me.ColIncome_Identifier.Index).Value
                        'copy data header
                        For ij As Integer = 0 To Me.DgvComponent.ColumnCount - 1
                            rowDGVTarget.Cells(ij).Value = rowDGVSumber.Cells(ij).Value
                        Next
                        'modifikasi data header
                        DummyDesc = "COPY OF " + rowDGVTarget.Cells(Me.ColIncome_Desc.Index).Value
                        While Me.GetIndex(DummyDesc) <> -1
                            DummyJml += 1
                            DummyDesc = "COPY OF " + rowDGVTarget.Cells(Me.ColIncome_Desc.Index).Value + "(" + DummyJml.ToString + ")"
                        End While
                        rowDGVTarget.Cells(Me.ColIncome_Desc.Index).Value = DummyDesc
                        rowDGVTarget.Cells(Me.ColIncome_Identifier.Index).Value = _
                        CInt(Me.mPrimarySalary).ToString + "_" + (Me.MAXUrutan + 1).ToString
                        rowDGVTarget.Cells(Me.ColIncome_Urutan.Index).Value = _
                        (Me.MAXUrutan + 1).ToString
                        Me.MAXUrutan = CInt(rowDGVTarget.Cells(Me.ColIncome_Urutan.Index).Value)
                        NewId = rowDGVTarget.Cells(Me.ColIncome_Identifier.Index).Value
                        rowDGVTarget.Cells(Me.Colincome_Primary.Index).Value = -1
                        rowDGVTarget.Cells(Me.ColIncome_CreatedBy.Index).Value = Me.mUserCode
                        rowDGVTarget.Cells(Me.ColIncome_ModifiedBy.Index).Value = Me.mUserCode
                        rowDGVTarget.Cells(Me.ColIncome_CreatedDate.Index).Value = Format(Now, "dd-MMM-yyyy HH:mm")
                        rowDGVTarget.Cells(Me.ColIncome_ModifiedDate.Index).Value = Format(Now, "dd-MMM-yyyy HH:mm")

                        'copy data variable
                        Me.TblVar.DefaultView.RowFilter = "[Primary Transaksi] = " + Me.mPrimarySalary.ToString + _
                        "AND [Identifier Salary Component] = '" + rowDGVSumber.Cells(Me.ColIncome_Identifier.Index).Value.ToString + "'"
                        If Me.TblVar.DefaultView.Count > 0 Then
                            For Each rowView As DataRowView In Me.TblVar.DefaultView
                                RowVarSumber = rowView.Row
                                RowVarTarget = Me.TblVar.Newtblmaster_salary_formula_salarycomponent_variable1Row
                                For ij As Integer = 0 To Me.TblVar.Columns.Count - 1
                                    RowVarTarget.Item(ij) = RowVarSumber.Item(ij)
                                Next
                                RowVarTarget.Identifier_Salary_Component = NewId
                                Me.TblVar.Rows.Add(RowVarTarget)
                            Next
                        End If

                        'pindahin ke dgv constraint
                        Dim IdxCons As Integer = 0
                        Me.TblCons.DefaultView.RowFilter = "[Primary Transaksi] = " + Me.mPrimarySalary.ToString + _
                        "AND [Identifier Salary Component] = '" + rowDGVSumber.Cells(Me.ColIncome_Identifier.Index).Value.ToString + "'"
                        If Me.TblCons.DefaultView.Count > 0 Then
                            For Each rowView As DataRowView In Me.TblCons.DefaultView
                                RowContSumber = rowView.Row
                                RowContTarget = Me.TblCons.Newtblmaster_salary_formula_salarycomponent_constraint1Row
                                For ij As Integer = 0 To Me.TblCons.Columns.Count - 1
                                    RowContTarget.Item(ij) = RowContSumber.Item(ij)
                                Next
                                RowContTarget.Identifier_Salary_Component = NewId
                                Me.TblCons.Rows.Add(RowContTarget)
                            Next
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.SedangIsiDGV = False
        End Try
    End Sub

    Private Sub dgvEmployee_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvEmployee.CellContentClick
        Dim frm As New FrmMasterSalaryFormula_OtherDeduction_ItemEmp
        Try
            If e.ColumnIndex = Me.ColEmp_OtherDeduction.Index Then
                frm.PrimaryEmp = Me.dgvEmployee.Item(Me.ColEmp_PrimaryEmployee.Index, e.RowIndex).Value
                frm.mUserCode = Me.mUserCode
                frm.TxtEmployeeCode.Text = Me.dgvEmployee.Item(Me.ColEmp_EmployeeCode.Index, e.RowIndex).Value
                frm.TxtEmployeeName.Text = Me.dgvEmployee.Item(Me.ColEmp_EmployeeName.Index, e.RowIndex).Value
                frm.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmployee_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvEmployee.RowsAdded
        Try
            If Me.BackgroundWorker3.IsBusy = False Then
                Me.BackgroundWorker3.Dispose()
                Me.BackgroundWorker3.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmployee_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvEmployee.RowsRemoved
        Try
            If Me.BackgroundWorker3.IsBusy = False Then
                Me.BackgroundWorker3.Dispose()
                Me.BackgroundWorker3.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmployee_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvEmployee.CellBeginEdit
        Try
            If Me.HakAksesUser_SalaryFormula(Me.mUserCode) = False Then
                e.Cancel = True
                Exit Sub
            End If

            If String.IsNullOrEmpty(Me.dgvEmployee.Item(Me.ColEmp_EmployeeCode.Index, e.RowIndex).Value) = True Then
                If e.ColumnIndex = Me.ColEmp_Basic.Index Or e.ColumnIndex = Me.ColEmp_EmployeeName.Index _
                Or e.ColumnIndex = Me.ColEmp_TaxCalculation.Index Or e.ColumnIndex = Me.ColEmp_Remarks.Index _
                Or e.ColumnIndex = Me.ColEmp_SalarySlipDesc.Index Or e.ColumnIndex = Me.ColEmp_TaxComponent.Index _
                Or e.ColumnIndex = Me.ColEmp_ExpenseAcc.Index Or e.ColumnIndex = Me.ColEmp_PaymentAcc.Index Then
                    e.Cancel = True
                Else
                    If e.ColumnIndex = Me.ColEmp_EmployeeName.Index Then
                        e.Cancel = True
                    Else
                        Me.Buffer = Me.dgvEmployee(e.ColumnIndex, e.RowIndex).Value
                    End If
                End If
            Else
                If e.ColumnIndex = Me.ColEmp_EmployeeName.Index Then
                    e.Cancel = True
                Else
                    Me.Buffer = Me.dgvEmployee(e.ColumnIndex, e.RowIndex).Value
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Information")
        End Try
    End Sub

    Private Sub dgvEmployee_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvEmployee.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvEmployee.Rows(e.RowIndex)
        Dim PrimaryEmp As Double = -1
        Try
            'kolom kode
            If e.ColumnIndex = Me.ColEmp_EmployeeCode.Index Then
                If String.IsNullOrEmpty(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.ObjEmp.ApakahSudahAda(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        PrimaryEmp = Me.ObjEmp.GetPrimaryByCode(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value)
                        If PrimaryEmp <> -1 Then
                            'tidak boleh double
                            If Me.GetIndexItemEmp(PrimaryEmp, Me.ColEmp_PrimaryEmployee.Index, e.RowIndex) < 0 Then
                                'cek apakah sudah ada pada formula yang lain
                                If Me.ObjSalaryFormula.ApakahSudahAdaEmployee(PrimaryEmp, Me.mPrimarySalary) = False Then
                                    Me.dgvEmployee.Item(Me.ColEmp_PrimaryEmployee.Index, e.RowIndex).Value = PrimaryEmp
                                    Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjEmp.GetCodeByPrimary(PrimaryEmp)
                                    Me.dgvEmployee.Item(Me.ColEmp_EmployeeName.Index, e.RowIndex).Value = Me.ObjEmp.GetFullNameByPrimary(PrimaryEmp)
                                    If e.RowIndex <> 0 Then
                                        Me.CopyFromRowBefore(e.RowIndex, e.RowIndex - 1)
                                    End If
                                Else
                                    MsgBox("Employee code has been used in other salary formula, please change employee code.", MsgBoxStyle.Information, "Information")
                                    If Me.Buffer = "" Then
                                        Me.mDGV = Me.dgvEmployee
                                        If row.Index <> Me.dgvEmployee.Rows.Count - 1 Then
                                            Me.EraseRowIndex = row.Index
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        End If
                                    Else
                                        Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                End If
                            Else
                                MsgBox("Employee code has been used, please change employee code.", MsgBoxStyle.Information, "Information")
                                If Me.Buffer = "" Then
                                    Me.mDGV = Me.dgvEmployee
                                    If row.Index <> Me.dgvEmployee.Rows.Count - 1 Then
                                        Me.EraseRowIndex = row.Index
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    End If
                                Else
                                    Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                End If
                            End If
                        Else
                            Me.dgvEmployee.Item(Me.ColEmp_PrimaryEmployee.Index, e.RowIndex).Value = -1
                            Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = ""
                            Me.dgvEmployee.Item(Me.ColEmp_EmployeeName.Index, e.RowIndex).Value = ""
                        End If
                    Else
                        ''diketik tidak ada lalu tampilkan form cari
                        Dim ij As New FrmMasterKaryawanCari
                        ij.txtCari.Text = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.BackColorAll = Me.BackColorAll
                        ij.ShowDialog()
                        If ij.PrimaryEmploye <> -1 Then
                            'tidak boleh double
                            If Me.GetIndexItemEmp(ij.PrimaryEmploye, Me.ColEmp_PrimaryEmployee.Index, e.RowIndex) < 0 Then
                                'cek apakah sudah ada pada formula yang lain
                                If Me.ObjSalaryFormula.ApakahSudahAdaEmployee(ij.PrimaryEmploye, Me.mPrimarySalary) = False Then
                                    Me.dgvEmployee.Item(Me.ColEmp_PrimaryEmployee.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                    Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjEmp.GetCodeByPrimary(ij.PrimaryEmploye)
                                    Me.dgvEmployee.Item(Me.ColEmp_EmployeeName.Index, e.RowIndex).Value = Me.ObjEmp.GetFullNameByPrimary(ij.PrimaryEmploye)
                                    If e.RowIndex <> 0 Then
                                        Me.CopyFromRowBefore(e.RowIndex, e.RowIndex - 1)
                                    End If
                                Else
                                    MsgBox("Employee code has been used in other salary formula, please change employee code.", MsgBoxStyle.Information, "Information")
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        If dgvEmployee.Rows.Count > 1 Then
                                            Me.mDGV = Me.dgvEmployee
                                            If row.Index <> Me.dgvEmployee.Rows.Count - 1 Then
                                                Me.EraseRowIndex = row.Index
                                                Me.BackgroundWorker1.RunWorkerAsync()
                                            End If
                                        End If
                                    Else
                                        Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                End If
                            Else
                                MsgBox("Employee code has been used, please change employee code.", MsgBoxStyle.Information, "Information")
                                If String.IsNullOrEmpty(Me.Buffer) = True Then
                                    If dgvEmployee.Rows.Count > 1 Then
                                        Me.mDGV = Me.dgvEmployee
                                        If row.Index <> Me.dgvEmployee.Rows.Count - 1 Then
                                            Me.EraseRowIndex = row.Index
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        End If
                                    End If
                                Else
                                    Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                End If
                            End If
                        Else
                            Me.dgvEmployee.Item(Me.ColEmp_PrimaryEmployee.Index, e.RowIndex).Value = -1
                            Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = ""
                            Me.dgvEmployee.Item(Me.ColEmp_EmployeeName.Index, e.RowIndex).Value = ""
                        End If
                        ij = Nothing
                    End If
                Else
                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                        If dgvEmployee.Rows.Count > 1 Then
                            Me.mDGV = Me.dgvEmployee
                            If row.Index <> Me.dgvEmployee.Rows.Count - 1 Then
                                Me.EraseRowIndex = row.Index
                                Me.BackgroundWorker1.RunWorkerAsync()
                            End If
                        End If
                    Else
                        Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            End If

            'kolom Basic salary
            If e.ColumnIndex = Me.ColEmp_Basic.Index Then
                If String.IsNullOrEmpty(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'cek numerik
                    If Me.ObjCd.CekNumerik(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please input number.", MsgBoxStyle.Information, "Information")
                        Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    'cek negatife
                    If Me.ObjCd.CekNonNegatif(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Cannot input negative number.", MsgBoxStyle.Information, "Information")
                        Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    'cek panjang
                    If Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value > 999999999999999 Then
                        MsgBox("Max value cannot than 999,999,999,999,999.", MsgBoxStyle.Information, "Information")
                        Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = _
                    Me.ObjCd.DigitGrouping(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefCurr)
                Else
                    Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCd.DigitGrouping(0, Me.DefCurr)
                End If
            End If

            'kolom expense acc
            If e.ColumnIndex = Me.ColEmp_ExpenseAcc.Index Then
                Dim PrimaryAcc As Double = -1
                Dim TipeAkun As Integer = -1
                Dim hasil As DialogResult
                If String.IsNullOrEmpty(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.ObjAcc.ApakahSudahAda_Aktif_NonHeader(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        PrimaryAcc = Me.ObjAcc.GetPrimary(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value)
                        'cek akun
                        If PrimaryAcc <> -1 Then
                            TipeAkun = Me.ObjAcc.GetTipeAkun(PrimaryAcc)
                            'selain expense, other expense dan COGS kasih peringantan yes / no
                            If TipeAkun = ClsMasterAkun1.EnumTipeAkun.Expense Or TipeAkun = ClsMasterAkun1.EnumTipeAkun.Other_Expense Or _
                            TipeAkun = ClsMasterAkun1.EnumTipeAkun.Cost_of_Goods_Sales Then
                                Me.dgvEmployee.Item(Me.ColEmp_PrimaryExpenseAcc.Index, e.RowIndex).Value = PrimaryAcc
                                Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAcc.GetAccountCode(PrimaryAcc)
                                Me.dgvEmployee.Item(Me.ColEmp_ExpenseAccDesc.Index, e.RowIndex).Value = Me.ObjAcc.GetDescription_ByPrimaryAcc(PrimaryAcc)
                            Else
                                hasil = MsgBox("Account type is not expense, other expense or COGS." + vbNewLine + _
                                "Do you want to continue ?", MsgBoxStyle.YesNo, "Information")
                                If hasil = Windows.Forms.DialogResult.Yes Then
                                    Me.dgvEmployee.Item(Me.ColEmp_PrimaryExpenseAcc.Index, e.RowIndex).Value = PrimaryAcc
                                    Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAcc.GetAccountCode(PrimaryAcc)
                                    Me.dgvEmployee.Item(Me.ColEmp_ExpenseAccDesc.Index, e.RowIndex).Value = Me.ObjAcc.GetDescription_ByPrimaryAcc(PrimaryAcc)
                                Else
                                    Me.dgvEmployee.Item(Me.ColEmp_PrimaryExpenseAcc.Index, e.RowIndex).Value = -1
                                    Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                    Me.dgvEmployee.Item(Me.ColEmp_ExpenseAccDesc.Index, e.RowIndex).Value = ""
                                End If
                            End If
                        Else
                            Me.dgvEmployee.Item(Me.ColEmp_PrimaryExpenseAcc.Index, e.RowIndex).Value = -1
                            Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = ""
                            Me.dgvEmployee.Item(Me.ColEmp_ExpenseAccDesc.Index, e.RowIndex).Value = ""
                        End If
                    Else
                        ''diketik tidak ada lalu tampilkan form cari
                        Dim ij As New FrmMasterAkun_Cari1
                        ij.txtCari.Text = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.mPrimaryAkun <> -1 Then
                            'cek akun
                            TipeAkun = Me.ObjAcc.GetTipeAkun(ij.mPrimaryAkun)
                            'selain expense, other expense dan COGS kasih peringantan yes / no
                            If TipeAkun = ClsMasterAkun1.EnumTipeAkun.Expense Or TipeAkun = ClsMasterAkun1.EnumTipeAkun.Other_Expense Or _
                            TipeAkun = ClsMasterAkun1.EnumTipeAkun.Cost_of_Goods_Sales Then
                                Me.dgvEmployee.Item(Me.ColEmp_PrimaryExpenseAcc.Index, e.RowIndex).Value = ij.mPrimaryAkun
                                Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAcc.GetAccountCode(ij.mPrimaryAkun)
                                Me.dgvEmployee.Item(Me.ColEmp_ExpenseAccDesc.Index, e.RowIndex).Value = Me.ObjAcc.GetDescription_ByPrimaryAcc(ij.mPrimaryAkun)
                            Else
                                hasil = MsgBox("Account type is not expense, other expense or COGS." + vbNewLine + _
                                "Do you want to continue ?", MsgBoxStyle.YesNo, "Information")
                                If hasil = Windows.Forms.DialogResult.Yes Then
                                    Me.dgvEmployee.Item(Me.ColEmp_PrimaryExpenseAcc.Index, e.RowIndex).Value = ij.mPrimaryAkun
                                    Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAcc.GetAccountCode(ij.mPrimaryAkun)
                                    Me.dgvEmployee.Item(Me.ColEmp_ExpenseAccDesc.Index, e.RowIndex).Value = Me.ObjAcc.GetDescription_ByPrimaryAcc(ij.mPrimaryAkun)
                                Else
                                    Me.dgvEmployee.Item(Me.ColEmp_PrimaryExpenseAcc.Index, e.RowIndex).Value = -1
                                    Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                    Me.dgvEmployee.Item(Me.ColEmp_ExpenseAccDesc.Index, e.RowIndex).Value = ""
                                End If
                            End If
                        Else
                            Me.dgvEmployee.Item(Me.ColEmp_PrimaryExpenseAcc.Index, e.RowIndex).Value = -1
                            Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = ""
                            Me.dgvEmployee.Item(Me.ColEmp_ExpenseAccDesc.Index, e.RowIndex).Value = ""
                            ij = Nothing
                        End If
                    End If
                Else
                    Me.dgvEmployee.Item(Me.ColEmp_PrimaryExpenseAcc.Index, e.RowIndex).Value = -1
                    Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Me.dgvEmployee.Item(Me.ColEmp_ExpenseAccDesc.Index, e.RowIndex).Value = ""
                End If
            End If

            'payment acc
            If e.ColumnIndex = Me.ColEmp_PaymentAcc.Index Then
                Dim PrimaryAcc As Double = -1
                Dim TipeAkun As Integer = -1
                Dim TipeAkun2 As Integer = -1
                Dim hasil As DialogResult
                If String.IsNullOrEmpty(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.ObjAcc.ApakahSudahAda_Aktif_NonHeader(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        PrimaryAcc = Me.ObjAcc.GetPrimary(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value)
                        'cek akun
                        If PrimaryAcc <> -1 Then
                            TipeAkun = Me.ObjAcc.GetTipeAkun(PrimaryAcc)
                            TipeAkun2 = Me.ObjAcc.GetTipeAkun2(PrimaryAcc)
                            'selain expense, other expense dan COGS kasih peringantan yes / no
                            If TipeAkun = ClsMasterAkun1.EnumTipeAkun.Asset And TipeAkun2 = ClsMasterAkun1.EnumTipeAkun2.Cash_And_Bank Then
                                Me.dgvEmployee.Item(Me.ColEmp_PrimaryPayment.Index, e.RowIndex).Value = PrimaryAcc
                                Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAcc.GetAccountCode(PrimaryAcc)
                                Me.dgvEmployee.Item(Me.ColEmp_PaymentAccDesc.Index, e.RowIndex).Value = Me.ObjAcc.GetDescription_ByPrimaryAcc(PrimaryAcc)
                            Else
                                hasil = MsgBox("Account type is not asset cash and bank." + vbNewLine + _
                                "Do you want to continue ?", MsgBoxStyle.YesNo, "Information")
                                If hasil = Windows.Forms.DialogResult.Yes Then
                                    Me.dgvEmployee.Item(Me.ColEmp_PrimaryPayment.Index, e.RowIndex).Value = PrimaryAcc
                                    Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAcc.GetAccountCode(PrimaryAcc)
                                    Me.dgvEmployee.Item(Me.ColEmp_PaymentAccDesc.Index, e.RowIndex).Value = Me.ObjAcc.GetDescription_ByPrimaryAcc(PrimaryAcc)
                                Else
                                    Me.dgvEmployee.Item(Me.ColEmp_PrimaryPayment.Index, e.RowIndex).Value = -1
                                    Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                    Me.dgvEmployee.Item(Me.ColEmp_PaymentAccDesc.Index, e.RowIndex).Value = ""
                                End If
                            End If
                        Else
                            Me.dgvEmployee.Item(Me.ColEmp_PrimaryPayment.Index, e.RowIndex).Value = -1
                            Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = ""
                            Me.dgvEmployee.Item(Me.ColEmp_PaymentAccDesc.Index, e.RowIndex).Value = ""
                        End If
                    Else
                        ''diketik tidak ada lalu tampilkan form cari
                        Dim ij As New FrmMasterAkun_Cari1
                        ij.txtCari.Text = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.mPrimaryAkun <> -1 Then
                            'cek akun
                            TipeAkun = Me.ObjAcc.GetTipeAkun(ij.mPrimaryAkun)
                            TipeAkun2 = Me.ObjAcc.GetTipeAkun(ij.mPrimaryAkun)
                            'selain expense, other expense dan COGS kasih peringantan yes / no
                            If TipeAkun = ClsMasterAkun1.EnumTipeAkun.Asset And TipeAkun2 = ClsMasterAkun1.EnumTipeAkun2.Cash_And_Bank Then
                                Me.dgvEmployee.Item(Me.ColEmp_PrimaryPayment.Index, e.RowIndex).Value = ij.mPrimaryAkun
                                Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAcc.GetAccountCode(ij.mPrimaryAkun)
                                Me.dgvEmployee.Item(Me.ColEmp_PaymentAccDesc.Index, e.RowIndex).Value = Me.ObjAcc.GetDescription_ByPrimaryAcc(ij.mPrimaryAkun)
                            Else
                                hasil = MsgBox("Account type is not asset cash and bank." + vbNewLine + _
                                "Do you want to continue ?", MsgBoxStyle.YesNo, "Information")
                                If hasil = Windows.Forms.DialogResult.Yes Then
                                    Me.dgvEmployee.Item(Me.ColEmp_PrimaryPayment.Index, e.RowIndex).Value = ij.mPrimaryAkun
                                    Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAcc.GetAccountCode(ij.mPrimaryAkun)
                                    Me.dgvEmployee.Item(Me.ColEmp_PaymentAccDesc.Index, e.RowIndex).Value = Me.ObjAcc.GetDescription_ByPrimaryAcc(ij.mPrimaryAkun)
                                Else
                                    Me.dgvEmployee.Item(Me.ColEmp_PrimaryPayment.Index, e.RowIndex).Value = -1
                                    Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                    Me.dgvEmployee.Item(Me.ColEmp_PaymentAccDesc.Index, e.RowIndex).Value = ""
                                End If
                            End If
                        Else
                            Me.dgvEmployee.Item(Me.ColEmp_PrimaryPayment.Index, e.RowIndex).Value = -1
                            Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = ""
                            Me.dgvEmployee.Item(Me.ColEmp_PaymentAccDesc.Index, e.RowIndex).Value = ""
                            ij = Nothing
                        End If
                    End If
                Else
                    Me.dgvEmployee.Item(Me.ColEmp_PrimaryPayment.Index, e.RowIndex).Value = -1
                    Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Me.dgvEmployee.Item(Me.ColEmp_PaymentAccDesc.Index, e.RowIndex).Value = ""
                End If
            End If

            'kolom salary slip desc
            If e.ColumnIndex = Me.ColEmp_SalarySlipDesc.Index Then
                Dim PrimarySlip As Double = -1
                If String.IsNullOrEmpty(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.ObjSlipSalary.ApakahSudahAda(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        PrimarySlip = Me.ObjSlipSalary.GetPrimaryByCode(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value)
                        If PrimarySlip <> -1 Then
                            Me.dgvEmployee.Item(Me.ColEmp_PrimarySlipDesc.Index, e.RowIndex).Value = PrimarySlip
                            Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjSlipSalary.GetCodeByPrimary(PrimarySlip)
                        Else
                            Me.dgvEmployee.Item(Me.ColEmp_PrimarySlipDesc.Index, e.RowIndex).Value = -1
                            Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        ''diketik tidak ada lalu tampilkan form cari
                        Dim ij As New FrmMasterSalaryFormula_SlipDesc_Cari
                        ij.txtCari.Text = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.BackColorAll = Me.BackColorAll
                        ij.ShowDialog()
                        If ij.Primary <> -1 Then
                            Me.dgvEmployee.Item(Me.ColEmp_PrimarySlipDesc.Index, e.RowIndex).Value = ij.Primary
                            Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjSlipSalary.GetCodeByPrimary(ij.Primary)
                        Else
                            Me.dgvEmployee.Item(Me.ColEmp_PrimarySlipDesc.Index, e.RowIndex).Value = -1
                            Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                        ij = Nothing
                    End If
                Else
                    MsgBox("Salary slip description cannot be an empty string", MsgBoxStyle.Information, "Information")
                    Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    'Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Selected = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CopyFromRowBefore(ByVal vIdxTarget As Integer, ByVal vIdxSource As Integer)
        Dim rowTarget, rowSource As DataGridViewRow

        Try
            rowTarget = Me.dgvEmployee.Rows(vIdxTarget)
            rowSource = Me.dgvEmployee.Rows(vIdxSource)

            'copy data
            '1. salary slip desc + primary
            rowTarget.Cells(Me.ColEmp_PrimarySlipDesc.Index).Value = rowSource.Cells(Me.ColEmp_PrimarySlipDesc.Index).Value
            rowTarget.Cells(Me.ColEmp_SalarySlipDesc.Index).Value = rowSource.Cells(Me.ColEmp_SalarySlipDesc.Index).Value
            '2. expense account + primary
            rowTarget.Cells(Me.ColEmp_PrimaryExpenseAcc.Index).Value = rowSource.Cells(Me.ColEmp_PrimaryExpenseAcc.Index).Value
            rowTarget.Cells(Me.ColEmp_ExpenseAcc.Index).Value = rowSource.Cells(Me.ColEmp_ExpenseAcc.Index).Value
            rowTarget.Cells(Me.ColEmp_ExpenseAccDesc.Index).Value = rowSource.Cells(Me.ColEmp_ExpenseAccDesc.Index).Value
            '3. payment account + primary
            rowTarget.Cells(Me.ColEmp_PrimaryPayment.Index).Value = rowSource.Cells(Me.ColEmp_PrimaryPayment.Index).Value
            rowTarget.Cells(Me.ColEmp_PaymentAcc.Index).Value = rowSource.Cells(Me.ColEmp_PaymentAcc.Index).Value
            rowTarget.Cells(Me.ColEmp_PaymentAccDesc.Index).Value = rowSource.Cells(Me.ColEmp_PaymentAccDesc.Index).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowRowProcess(ByVal vDgv As DataGridView, ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < vDgv.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    vDgv.FirstDisplayedScrollingRowIndex = vDgv.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > vDgv.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While vDgv.Rows(row.Index).Displayed = False
                    vDgv.FirstDisplayedScrollingRowIndex = vDgv.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If vDgv.DisplayedRowCount(True) <> vDgv.DisplayedRowCount(False) Then
                    vDgv.FirstDisplayedScrollingRowIndex = vDgv.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUP.Click
        Dim vDgv As DataGridView
        Try
            Me.SedangIsiDGV = True

            If Me.TabPage1.Focus = True Then
                vDgv = Me.DgvComponent
            ElseIf Me.TabPage2.Focus = True Then
                vDgv = Me.dgvEmployee
            End If
            vDgv.AllowUserToAddRows = False

            If vDgv.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = vDgv.SelectedRows(0)

                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    vDgv.Rows.Remove(row)
                    vDgv.Rows.InsertRange(indexbaru, row)
                    vDgv.Focus()
                    vDgv.ClearSelection()
                    row.Selected = True
                    ''Me.ObjCD.SetNumber(Me.dgvItem)
                    If row.Displayed = False Then
                        Me.ShowRowProcess(vDgv, row, True)
                    End If
                End If

                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
            Me.SedangIsiDGV = False
            vDgv.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.SedangIsiDGV = False
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Dim vDgv As DataGridView
        Try
            Me.SedangIsiDGV = True

            If Me.TabPage1.Focus = True Then
                vDgv = Me.DgvComponent
            ElseIf Me.TabPage2.Focus = True Then
                vDgv = Me.dgvEmployee
            End If
            vDgv.AllowUserToAddRows = False
            If vDgv.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False

                Dim row As DataGridViewRow = vDgv.SelectedRows(0)
                If row.Index <> vDgv.Rows.Count - 1 Then
                    Dim indexbaru As Integer = row.Index + 1
                    vDgv.Rows.Remove(row)
                    vDgv.Rows.InsertRange(indexbaru, row)
                    vDgv.Focus()
                    vDgv.ClearSelection()
                    row.Selected = True
                    ''Me.ObjCD.SetNumber(Me.dgvItem)
                    Me.ShowRowProcess(vDgv, row, False)
                End If

                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If

            Me.SedangIsiDGV = False
            vDgv.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.SedangIsiDGV = False
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
            If Me.EraseRowIndex <> Me.DgvComponent.Rows.Count - 1 Then
                Me.mDGV.Rows.RemoveAt(Me.EraseRowIndex)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Dim CboDgv As DataGridViewComboBoxCell
        Dim i As Integer = 0
        Dim MAX As Integer = 0
        Try
            CboDgv = Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, Me.DgvComponent.Rows.Count - 1)
            Me.DgvComponent.Item(Me.ColIncome_ComponentTipe.Index, Me.DgvComponent.Rows.Count - 1).Value = CboDgv.Items.Item(i)
            CboDgv = Me.DgvComponent.Item(Me.ColIncome_TaxComponent.Index, Me.DgvComponent.Rows.Count - 1)
            Me.DgvComponent.Item(Me.ColIncome_TaxComponent.Index, Me.DgvComponent.Rows.Count - 1).Value = CboDgv.Items.Item(i)
            CboDgv = Me.DgvComponent.Item(Me.ColIncome_CalculateOn.Index, Me.DgvComponent.Rows.Count - 1)
            Me.DgvComponent.Item(Me.ColIncome_CalculateOn.Index, Me.DgvComponent.Rows.Count - 1).Value = CboDgv.Items.Item(i)
            CboDgv = Me.DgvComponent.Item(Me.ColIncome_CalculatePeriod.Index, Me.DgvComponent.Rows.Count - 1)
            Me.DgvComponent.Item(Me.ColIncome_CalculatePeriod.Index, Me.DgvComponent.Rows.Count - 1).Value = CboDgv.Items.Item(i)

            Me.DgvComponent.Item(Me.ColFormula.Index, Me.DgvComponent.Rows.Count - 1).Value = 0
            Me.DgvComponent.Item(Me.ColRoundingUnit.Index, Me.DgvComponent.Rows.Count - 1).Value = 1
            Me.DgvComponent.Item(Me.ColRoundingType.Index, Me.DgvComponent.Rows.Count - 1).Value = 0
            Me.DgvComponent.Item(Me.ColYearlyMaxValue.Index, Me.DgvComponent.Rows.Count - 1).Value = 0
            Me.DgvComponent.Item(Me.ColYearlyMinValue.Index, Me.DgvComponent.Rows.Count - 1).Value = 0
            Me.DgvComponent.Item(Me.ColMonthlyMaxValue.Index, Me.DgvComponent.Rows.Count - 1).Value = 0
            Me.DgvComponent.Item(Me.ColMonthlyMinValue.Index, Me.DgvComponent.Rows.Count - 1).Value = 0
            Me.DgvComponent.Item(Me.Colincome_Primary.Index, Me.DgvComponent.Rows.Count - 1).Value = -1
            Me.DgvComponent.Item(Me.ColIncome_CreatedBy.Index, Me.DgvComponent.Rows.Count - 1).Value = Me.mUserCode
            Me.DgvComponent.Item(Me.ColIncome_CreatedDate.Index, Me.DgvComponent.Rows.Count - 1).Value = Format(Now, "dd-MMM-yyyy") + " " + CDate(Now).ToShortTimeString
            Me.DgvComponent.Item(Me.ColIncome_ModifiedBy.Index, Me.DgvComponent.Rows.Count - 1).Value = Me.mUserCode
            Me.DgvComponent.Item(Me.ColIncome_ModifiedDate.Index, Me.DgvComponent.Rows.Count - 1).Value = Format(Now, "dd-MMM-yyyy") + " " + CDate(Now).ToShortTimeString
            Me.DgvComponent.Item(Me.ColIncome_Routine.Index, Me.DgvComponent.Rows.Count - 1).Value = Boolean.TrueString
            Me.DgvComponent.Item(Me.ColIncome_PrimaryPaymentAccount.Index, Me.DgvComponent.RowCount - 1).Value = -1
            Me.DgvComponent.Item(Me.ColIncome_PaymentAccount.Index, Me.DgvComponent.RowCount - 1).Value = ""

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub HapusVarNConst_ByIdentifierComponent(ByVal vIdentifier As String)
        Dim rowVar As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_variable1Row
        Dim rowCons As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_constraint1Row

        Try
            'Hapus Variable
            Me.TblVar.DefaultView.RowFilter = "[Identifier Salary Component] = '" & vIdentifier.ToString & "'"
            While Me.TblVar.DefaultView.Count > 0
                rowVar = Me.TblVar.DefaultView(0).Row
                Me.TblVar.Rows.Remove(rowVar)
            End While

            'Hapus Constraint
            Me.TblCons.DefaultView.RowFilter = "[Identifier Salary Component] = '" & vIdentifier.ToString & "'"
            While Me.TblCons.DefaultView.Count > 0
                rowCons = Me.TblCons.DefaultView(0).Row
                Me.TblCons.Rows.Remove(rowCons)
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub BackgroundWorker3_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Dim CboDgv As DataGridViewComboBoxCell
        Dim i As Integer = 0
        Try
            CboDgv = Me.dgvEmployee.Item(Me.ColEmp_TaxCalculation.Index, Me.dgvEmployee.Rows.Count - 1)
            Me.dgvEmployee.Item(Me.ColEmp_TaxCalculation.Index, Me.dgvEmployee.Rows.Count - 1).Value = _
            CboDgv.Items.Item(i)
            Me.dgvEmployee.Item(Me.ColEmp_Basic.Index, Me.dgvEmployee.Rows.Count - 1).Value = Me.ObjCd.DigitGrouping(0, Me.DefCurr)
            CboDgv = Me.dgvEmployee.Item(Me.ColEmp_TaxComponent.Index, Me.dgvEmployee.Rows.Count - 1)
            Me.dgvEmployee.Item(Me.ColEmp_TaxComponent.Index, Me.dgvEmployee.Rows.Count - 1).Value = _
            CboDgv.Items.Item(ClsMasterSalaryFormula_item.EnumTax2.gaji_pensiun)
            Me.dgvEmployee.Item(Me.ColEmp_Primary.Index, Me.dgvEmployee.Rows.Count - 1).Value = -1
            Me.dgvEmployee.Item(Me.ColEmp_CreatedBy.Index, Me.dgvEmployee.Rows.Count - 1).Value = Me.mUserCode
            Me.dgvEmployee.Item(Me.ColEmp_ModifedBy.Index, Me.dgvEmployee.Rows.Count - 1).Value = Me.mUserCode
            Me.dgvEmployee.Item(Me.ColEmp_ModifiedDate.Index, Me.dgvEmployee.Rows.Count - 1).Value = Format(Now, "dd-MMM-yyyy") + " " + CDate(Now).ToShortTimeString
            Me.dgvEmployee.Item(Me.ColEmp_CreatedDate.Index, Me.dgvEmployee.Rows.Count - 1).Value = Format(Now, "dd-MMM-yyyy") + " " + CDate(Now).ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker_CloseForm_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_CloseForm.DoWork
        Try
            Thread.Sleep(1)
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

    Private Sub FrmMasterSalaryFormula_Component_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            If Me.mClose = True Then
                MsgBox("Can't access salary formula transaction because user code : " + Me.UserLain.ToString + _
                " at computer name-ip address : " + Me.NamaKomputerLain.ToString + " is using this transaction.", _
                MsgBoxStyle.Information, "Information")
            End If
            Me.ObjAntrian.EndAntrianTunggal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterSalaryFormula_Component_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim hasil As DialogResult
        Try
            If Me.mClose = False Then
                If _ok = False Then
                    If Me.DgvComponent.Rows.Count > 1 Then
                        If Me.HakAksesUser_SalaryFormula(Me.mUserCode, False) = False Then
                            Me.BackgroundWorker_CloseForm.RunWorkerAsync()
                        Else
                            hasil = MsgBox("Save data before closing?", MsgBoxStyle.YesNoCancel, "Confirm")
                        End If
                        If hasil = Windows.Forms.DialogResult.Yes Then
                            If Me.ApakahDataBenar = True Then
                                WriteModified()
                                Me.ObjItem.Simpan(Me, Me.TblVar, Me.TblCons, Me.tblLogUser)
                                MsgBox("Data saved.", MsgBoxStyle.Information, "Information")
                                Me._ok = True
                                Me.Close()
                            Else
                                e.Cancel = True
                            End If
                        ElseIf hasil = Windows.Forms.DialogResult.No Then
                            Me.BackgroundWorker_CloseForm.RunWorkerAsync()
                        ElseIf hasil = Windows.Forms.DialogResult.Cancel Then
                            e.Cancel = True
                        End If
                    End If
                End If
            Else
                Me.BackgroundWorker_CloseForm.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterSalaryFormula_Component_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.ShowIcon = False
            System.Windows.Forms.Form.CheckForIllegalCrossThreadCalls = False
            Me.DgvComponent.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvEmployee.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            UserLain = ""
            NamaKomputerLain = ""
            If Me.ObjAntrian.StartAntrianTunggal(ClsAntrian.EnumTipeAntrian.AntrianTunggalSalaryFormula, Me.mUserCode, UserLain, NamaKomputerLain) = False Then
                Me.mClose = True
                Me.BackgroundWorker_CloseForm.RunWorkerAsync()
            End If
            'cari data component nya dulu 
            ''pindahin ke dgv
            'Me.ObjItem.IsiDgvComponent(Me, Me.DefCurr)
            'isi tabel variabel dan constrain
            Me.ObjItem.IsiTabelVarConts(Me)
            Me.CancelButton = Me.btnClose
            'ColIncome_Identifier.Visible = True

            '@12
            Me.llbImport.Visible = False
            Me.llbExport.Visible = False

            If Me.HakAksesUser_SalaryFormula(Me.mUserCode, False) = False Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                'Me.llbExport.Enabled = False
                Me.llbImport.Enabled = False
                Me.LinkLabel1.Enabled = False
                Me.LinkLabel2.Enabled = False
                Me.btnOK.Enabled = False
                Me.btnSave.Enabled = False
                'Me.DgvComponent.AllowUserToAddRows = False
                'Me.DgvComponent.AllowUserToDeleteRows = False
                'Me.dgvEmployee.AllowUserToAddRows = False
                ''Me.dgvEmployee.AllowUserToDeleteRows = False
            End If

            Me.ColEmp_OtherDeduction.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim frm As New FrmMasterSalaryFormula_SlipDesc
        Try
            frm.BackColorAll = Me.BackColorAll
            frm.ShowDialog()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            frm.Dispose()
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult

        Try
            If Me.ApakahDataBenar = True Then
                hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    WriteModified()
                    Me.ObjItem.Simpan(Me, Me.TblVar, Me.TblCons, Me.tblLogUser)
                    MsgBox("Data saved.", MsgBoxStyle.Information, "Information")
                    Me._ok = True
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LinkLabel2.Click
        Dim hj As New FrmMasterSalaryFormula_Cari
        Dim TempPrimary As Double = -1
        Try

            TempPrimary = mPrimarySalary
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.Primary = -1 Then
                Me.DgvComponent.Rows.Clear()
                mPrimarySalary = hj.Primary
                Me.DgvComponent.AllowUserToAddRows = False
                Me.ObjItem.CopySalaryComponent(Me, Me.DefCurr, Me.TblVar, Me.TblCons, TempPrimary, Me.mUserCode)
                Me.DgvComponent.AllowUserToAddRows = True
            End If
            mPrimarySalary = TempPrimary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub dgvEmployee_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvEmployee.UserDeletingRow
        'Dim dta As New DseMasterSalaryFormulaTableAdapters.tbltrans_salaryTableAdapter
        'Dim tbl As New DseMasterSalaryFormula.tbltrans_salaryDataTable
        'Dim row As DseMasterSalaryFormula.tbltrans_salaryRow
        'Dim list As String = ""
        Try
            If Me.HakAksesUser_SalaryFormula(Me.mUserCode) = False Then
                e.Cancel = True
                Exit Sub
            End If
            '    dta.Fill(tbl, Me.dgvEmployee.Item(Me.ColEmp_PrimaryEmployee.Index, e.Row.Index).Value)
            '    If tbl.Rows.Count > 0 Then
            '        Dim i As Integer = 1
            '        For Each row In tbl.Rows
            '            list = i.ToString + ". " + row.No_Transaksi.ToString + " Period : " + Format(row.Start_Date, "dd MMM yyyy") + " - " + Format(row.End_Date, "dd MMM yyyy")
            '            i = i + 1
            '        Next
            '        MsgBox("Employee detail can't delete because refering to calculate salary : " + vbNewLine + _
            '        list, MsgBoxStyle.Information, "Information")
            '        e.Cancel = True
            '    End If
            '    'cek apakah sudah di calculate salary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker4_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker4.DoWork
        Try
            Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker4_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker4.RunWorkerCompleted
        Try
            SetTaxComp(Me.e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetTaxComp(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim CboDgv As DataGridViewComboBoxCell
        Try
            CboDgv = Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex)
            If CboDgv.Items.IndexOf(Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value) = EnumType.Income Then
                CboDgv = Me.DgvComponent.Item(Me.ColIncome_TaxComponent.Index, e.RowIndex)
                If CboDgv.Items.Contains("IURAN PENSIUN ATAU IURAN THT/JHT") = True Then
                    Me.DgvComponent.Item(Me.ColIncome_TaxComponent.Index, e.RowIndex).Value = CboDgv.Items.Item(0)
                End If
                CboDgv.Items.Remove("IURAN PENSIUN ATAU IURAN THT/JHT")
            ElseIf CboDgv.Items.IndexOf(Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value) = EnumType.Deduction Or _
            CboDgv.Items.IndexOf(Me.DgvComponent.Item(e.ColumnIndex, e.RowIndex).Value) = EnumType.CompanyExpense Then
                CboDgv = Me.DgvComponent.Item(Me.ColIncome_TaxComponent.Index, e.RowIndex)
                CboDgv.Items.Insert(6, "IURAN PENSIUN ATAU IURAN THT/JHT")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvComponent_DataError(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvComponent.DataError
        Try
            'do nothing, hanya untuk handle masalah error saja
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim hasil As MsgBoxResult

        Try
            If Me.ApakahDataBenar = True Then
                hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    WriteModified()
                    Me.ObjItem.Simpan(Me, Me.TblVar, Me.TblCons, Me.tblLogUser)
                    MsgBox("Data saved.", MsgBoxStyle.Information, "Information")
                    'pindahin ke dgv
                    Me.ObjItem.IsiDgvComponent(Me, Me.DefCurr)
                    'If Me.ObjAntrian.StartAntrianTunggal(ClsAntrian.EnumTipeAntrian.AntrianTunggalSalaryFormula, Me.mUserCode, UserLain, NamaKomputerLain) = False Then
                    '    Me.mClose = True
                    '    Me.BackgroundWorker_CloseForm.RunWorkerAsync()
                    'End If
                    Me.ObjItem.IsiTabelVarConts(Me)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvComponent_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles DgvComponent.UserDeletingRow
        Try
            If Me.HakAksesUser_SalaryFormula(Me.mUserCode) = False Then
                e.Cancel = True
                Exit Sub
            End If

            'jika sudah digunakan oleh salary item cacl, maka tidak boleh didelete
            If Me.ObjSalaryFormula.GetJml_ItemSalaryCalc_ByIdentifierSalaryFormula(Me.DgvComponent.Item(Me.ColIncome_Identifier.Index, e.Row.Index).Value) > 0 Then
                MsgBox("This component is used is salary calculation, you can't delete this data.", MsgBoxStyle.Information, "Information")
                e.Cancel = True
            Else
                Me.HapusVarNConst_ByIdentifierComponent(Me.DgvComponent.Item(Me.ColIncome_Identifier.Index, e.Row.Index).Value)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbExport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles llbExport.Click
        Dim DTAEmp As New DseMasterSalaryFormulaTableAdapters.DataTable3TableAdapter
        Dim TBLEmp As New DseMasterSalaryFormula.DataTable3DataTable
        Dim RowEmp As DseMasterSalaryFormula.DataTable3Row
        Dim listEmployee As New List(Of String)
        Dim ix As Integer = 0
        Dim FileExtStr As String = ""
        Dim hj As New FrmMassageCSV
        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.dgvEmployee.RowCount < 1 Then
                Exit Sub
            End If

            If Me.GetFrmDia(FileExtStr) = False Then
                Exit Sub
            End If

            DTAEmp.Connection.ConnectionString = Core_Module.GetConstring
            DTAEmp.Fill(TBLEmp, mPrimarySalary)

            For Each rowd As DataGridViewRow In Me.dgvEmployee.Rows
                If rowd.Cells(Me.ColEmp_EmployeeCode.Index).Value <> "" Then
                    TBLEmp.DefaultView.RowFilter = "[Primary] = " & rowd.Cells(Me.ColEmp_Primary.Index).Value
                    If TBLEmp.DefaultView.Count > 0 Then
                        RowEmp = TBLEmp.DefaultView(0).Row
                        RowEmp.Amount = CDec(rowd.Cells(Me.ColEmp_Basic.Index).Value)
                        RowEmp.Urutan = rowd.Index + 1
                    End If
                End If
            Next

            While ix <= TBLEmp.Rows.Count - 1
                RowEmp = TBLEmp.Rows(ix)
                If RowEmp.Amount = 0 Then
                    TBLEmp.Rows.Remove(RowEmp)
                Else
                    ix += 1
                End If
            End While

            If FileExtStr.StartsWith(".csv") Then
                If hj.ShowDialog = Windows.Forms.DialogResult.Yes Then
                    If hj.Rbo = 0 Then
                        Me.ExportFileCsv(TBLEmp)
                    Else
                        Me.ExportFileCsv2(TBLEmp)
                    End If
                End If

                'ElseIf FileExtStr.StartsWith(".xlsx") Then
                '    Me.ExportFileExcel(TBLEmp)
                'ElseIf FileExtStr.StartsWith(".xls") Then
                '    Me.ExportFileExcel(TBLEmp)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Function Openfile(ByRef vFileExtStr As String) As Boolean
        Dim FileExtStr As String = ""

        Try
            '1. cari file csv format infinity
            Me.OpenFileDialog1.Title = "Please specify file"
            Me.OpenFileDialog1.Filter = "csv files (*.csv)|*.csv"
            'Me.OpenFileDialog1.Filter = "Excel Worksheets (*.xlsx)|*.xlsx|excel files (*.xls)|*.xls|csv files (*.csv)|*.csv"
            Me.OpenFileDialog1.FilterIndex = 1
            Me.OpenFileDialog1.DefaultExt = "csv"
            Me.OpenFileDialog1.CheckFileExists = True
            Me.OpenFileDialog1.Multiselect = False
            Me.OpenFileDialog1.FileName = ""
            If Me.OpenFileDialog1.ShowDialog <> Windows.Forms.DialogResult.OK Then
                Return False
            End If
            If Me.OpenFileDialog1.FileName = "" Then
                Return False
            End If

            Me.LocationFileSumber = Me.OpenFileDialog1.FileName
            vFileExtStr = System.IO.Path.GetExtension(Me.OpenFileDialog1.FileName)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub llbImport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles llbImport.Click
        Dim FileExtStr As String = ""

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If Me.Openfile(FileExtStr) = False Then
                Exit Sub
            End If


            If FileExtStr.StartsWith(".csv") Then
                Me.ImportFile()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            If TabControl1.SelectedIndex = 0 Then
                Me.llbImport.Visible = False
                Me.llbExport.Visible = False
            ElseIf TabControl1.SelectedIndex = 1 Then
                Me.llbExport.Visible = True
                Me.llbImport.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetFrmDia(ByRef vFileExtStr As String) As Boolean
        Try
            Me.SaveDia.FileName = ""
            Me.SaveDia.Title = "Please specify file"
            Me.SaveDia.Filter = "csv files (*.csv)|*.csv"
            'Me.SaveDia.Filter = "Excel Worksheets (*.xlsx)|*.xlsx|excel files (*.xls)|*.xls|csv files (*.csv)|*.csv"
            Me.SaveDia.FilterIndex = 1
            'Me.SaveDia.DefaultExt = "xls"
            Me.SaveDia.DefaultExt = "csv"
            If Me.SaveDia.ShowDialog <> Windows.Forms.DialogResult.OK And Me.SaveDia.FileName.Trim.Length = 0 Then
                Return False
            End If

            vFileExtStr = System.IO.Path.GetExtension(Me.SaveDia.FileName)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub ExportFileCsv(ByVal TBLEmp As DseMasterSalaryFormula.DataTable3DataTable)
        Dim RowEmp As DseMasterSalaryFormula.DataTable3Row
        Dim csvVal As String = ""

        Try
            If Me.dgvEmployee.Rows.Count - 1 <= 0 Then
                MsgBox("There is no data will be exported.", MsgBoxStyle.Information, "Information")
            Else
                csvVal = "NIK;Employee Code;Employee Name;Department Code;Department Description;Job Position Code;Job Position Description;Basic Salary"

                TBLEmp.DefaultView.RowFilter = Nothing
                TBLEmp.DefaultView.Sort = "[Urutan] ASC"
                For i As Integer = 0 To TBLEmp.DefaultView.Count - 1
                    RowEmp = TBLEmp.DefaultView(i).Row
                    With RowEmp
                        csvVal += vbNewLine & .NIK.ToString.Replace(",", "") & ";"
                        csvVal += .Employee_Code.ToString.Replace(",", "") & ";"
                        csvVal += .Employee_Name.ToString.Replace(",", "") & ";"
                        csvVal += .Dept_Code.ToString.Replace(",", "") & ";"
                        csvVal += .Dept_Desc.ToString.Replace(",", "") & ";"
                        csvVal += .Code_Post.ToString.Replace(",", "") & ";"
                        csvVal += .Desc_Post.ToString.Replace(",", "") & ";"
                        csvVal += .Amount.ToString.Replace(",", "")
                    End With
                Next

                My.Computer.FileSystem.WriteAllText(SaveDia.FileName, csvVal, False)
                MsgBox("File Export Success", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '
    Private Sub ExportFileCsv2(ByVal TBLEmp As DseMasterSalaryFormula.DataTable3DataTable)
        Dim RowEmp As DseMasterSalaryFormula.DataTable3Row
        Dim csvVal As String = ""

        Try
            If Me.dgvEmployee.Rows.Count - 1 <= 0 Then
                MsgBox("There is no data will be exported.", MsgBoxStyle.Information, "Information")
            Else
                csvVal = "NIK, Employee Code, Employee Name, Departement Code, Departement Description, Job Position Code, Job Position Description, Basic Salary"

                TBLEmp.DefaultView.RowFilter = Nothing
                TBLEmp.DefaultView.Sort = "[Urutan] ASC"
                For i As Integer = 0 To TBLEmp.DefaultView.Count - 1
                    RowEmp = TBLEmp.DefaultView(i).Row
                    With RowEmp
                        csvVal += vbNewLine & .NIK.ToString.Replace(";", "") & ","
                        csvVal += .Employee_Code.ToString.Replace(";", "") & ","
                        csvVal += .Employee_Name.ToString.Replace(";", "") & ","
                        csvVal += .Dept_Code.ToString.Replace(";", "") & ","
                        csvVal += .Dept_Desc.ToString.Replace(";", "") & ","
                        csvVal += .Code_Post.ToString.Replace(";", "") & ","
                        csvVal += .Desc_Post.ToString.Replace(";", "") & ","
                        csvVal += .Amount.ToString.Replace(";", "")
                    End With
                Next

                My.Computer.FileSystem.WriteAllText(Me.SaveDia.FileName, csvVal, False)
                MsgBox("File Export Success", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    '
    Private Sub ImportFile()
        Dim SR As System.IO.StreamReader = New System.IO.StreamReader(Me.LocationFileSumber)
        Dim line As String = SR.ReadLine()
        'Dim strArray As String() = line.Split(";"c)
        Dim Tbl As DataTable = New DataTable
        Dim row As DataRow

        Dim NamaFilter As String = ""
        Dim JmlKaryawan As Integer = 0
        Dim ListNama_InDGV As New List(Of String)
        Dim ListNama_1 As New List(Of String) 'list karyawan di dgv yang ga ada di csv
        Dim ListNama_2 As New List(Of String) 'list karyawan di csv yang ga ada di dgv
        Dim StrMsgBox_1 As String = ""
        Dim StrMsgBox_2 As String = ""
        Dim StrMsgBox_3 As String = ""
        Dim ApaExitSub As Boolean = False
        Dim iCOunt As String() = line.Split(","c)
        Dim Icount2 As String() = line.Split(";"c)
        Dim strArray As String()
        Try
            If iCOunt.Length > Icount2.Length Then
                strArray = line.Split(","c)
            Else
                strArray = line.Split(";"c)
            End If
            'pindahkan data csv ke dalam data table
            For Each s As String In strArray
                Tbl.Columns.Add(New DataColumn())
            Next
            Tbl.Columns.Add(New DataColumn())

            If Tbl.Columns.Count <> 3 Then
                MsgBox("Make sure csv file consist of name and amount.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Do
                line = SR.ReadLine
                If Not line = String.Empty Then
                    If iCOunt.Length > Icount2.Length Then
                        row = Tbl.NewRow
                        row.ItemArray = line.Split(","c)
                        row.Item(2) = row.Item(0).ToString.Replace(" ", "").Replace(".", "").Replace(",", "").ToUpper
                        Tbl.Rows.Add(row)
                    Else
                        row = Tbl.NewRow
                        row.ItemArray = line.Split(";"c)
                        row.Item(2) = row.Item(0).ToString.Replace(" ", "").Replace(".", "").Replace(",", "").ToUpper
                        Tbl.Rows.Add(row)
                    End If
                Else
                    Exit Do
                End If
            Loop

            'pindahkan list karyawan yang ada pada dgv ke tabel csv
            For Each rowd As DataGridViewRow In Me.dgvEmployee.Rows
                If String.IsNullOrEmpty(rowd.Cells(Me.ColEmp_EmployeeName.Index).Value) = False Then
                    NamaFilter = rowd.Cells(Me.ColEmp_EmployeeName.Index).Value.ToString.Replace(" ", "").Replace(".", "").Replace(",", "").ToUpper
                    If ListNama_InDGV.Contains(NamaFilter) = False Then
                        ListNama_InDGV.Add(NamaFilter)
                        'pastikan hanya pindahkan apabila karyawan tsb ada di list karyawan pada file csv
                        Tbl.DefaultView.RowFilter = "[Column3] = '" & NamaFilter.ToString & "'"
                        If Tbl.DefaultView.Count > 0 Then
                            row = Tbl.DefaultView(0).Row
                            rowd.Cells(Me.ColEmp_Basic.Index).Value = Me.ObjCd.DigitGrouping(row.Item(1), Me.DefCurr)
                        Else
                            If ListNama_1.Contains(rowd.Cells(Me.ColEmp_EmployeeName.Index).Value) = False Then
                                ListNama_1.Add(rowd.Cells(Me.ColEmp_EmployeeName.Index).Value)
                            End If
                        End If
                    End If
                End If
            Next

            For Each row In Tbl.Rows
                If ListNama_InDGV.Contains(row.Item(2).ToString.Replace(" ", "").Replace(".", "").Replace(",", "").ToUpper) = False Then
                    If ListNama_2.Contains(row.Item(0)) = False Then
                        ListNama_2.Add(row.Item(0))
                    End If
                End If
            Next
            If ListNama_1.Count > 0 Then
                ListNama_1.Sort()
                For Each NamaKary As String In ListNama_1
                    If StrMsgBox_1 = "" Then
                        StrMsgBox_1 = "Employee not found in csv file." & vbCrLf & NamaKary
                    Else
                        StrMsgBox_1 &= ", " & NamaKary
                    End If
                Next
            End If
            If ListNama_2.Count > 0 Then
                ListNama_2.Sort()
                For Each NamaKary As String In ListNama_2
                    If StrMsgBox_2 = "" Then
                        StrMsgBox_2 = "Employee not found in item" & vbCrLf & NamaKary
                    Else
                        StrMsgBox_2 &= ", " & NamaKary
                    End If
                Next
                StrMsgBox_2 &= vbCrLf
            End If

            If StrMsgBox_1 <> "" Or StrMsgBox_2 <> "" Then
                If StrMsgBox_1 <> "" Then
                    StrMsgBox_1 &= vbCrLf & vbCrLf
                End If
                MsgBox(StrMsgBox_1 & StrMsgBox_2, MsgBoxStyle.Information, "Information")
            End If

            MsgBox("File Import Success", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbImport_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbImport.LinkClicked

    End Sub

    Private Sub llbExport_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbExport.LinkClicked

    End Sub

    Dim ObjUser As New Core_Module.ClsUser
    Public Function HakAksesUser_SalaryFormula(ByVal vUserCode As String, Optional ByVal vShowMsg As Boolean = True) As Boolean
        Try
            If Me.ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmMasterSalaryFormula_Detail") = False Then
                If vShowMsg = True Then
                    MsgBox("You don't have authorization form this information.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class
