Option Explicit On
Option Strict Off

Public Class FrmMasterSalaryFormula_Component1
    Public BackColorAll As Color = Color.AliceBlue
    Public Buffer_CalculatePeriod As String = ""
    Public Objcd As ClsCekData
    Public mType As ClsMasterParameterGaji.EnumType = ClsMasterParameterGaji.EnumType.Daily
    Public mIdentifiyComp As String = ""
    Public mPrimaryTransaksi As Double = -1
    Public mTblVar As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_variable1DataTable
    Public mTblConst As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_constraint1DataTable
    Public PunyaHakAksesFull As Boolean = True
    Dim RowVar As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_variable1Row
    Dim RowConst As DseMasterSalaryFormula.tblmaster_salary_formula_salarycomponent_constraint1Row
    Dim ObjCurr As New clsMasterCurrency
    Dim ObjKompetensi As New Core_Module.ClsMasterKompetensi
    Dim ObjSalaryFormula As New Core_Module.ClsMasterSalaryFormula_item
    Dim DefPrec As Integer = Me.ObjCurr.GetDecimal_DefaultCurrency()
    Dim Buffer As String = ""
    Dim mDGV As DataGridView
    Dim EraseRowIndex As Integer = 0
    Dim ObjParameter As New ClsMasterParameterGaji
    Dim ObjJP As New Core_Module.ClsMasterJobPosition
    Dim ObjCuti As New Core_Module.ClsMasterCuti
    'Dim i As New Ciloci.Flee.CalcEngine.CalculationEngine
    Dim FromOtherButton As Boolean = False
    Dim CancelClose As Boolean = False
    Public salin As Boolean = False
    Public mUserCode As String = ""
    Dim BerhasilSave As Boolean = True


    Enum EnumCompare
        a '=
        b '>
        c '<
        d '>=
        e '<=
        f '<>
    End Enum

    Private Function GetIndexItemVar(ByVal vData As String, ByVal vDgv As DataGridView, _
    ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In vDgv.Rows
                If String.IsNullOrEmpty(rowDGV.Cells(e.ColumnIndex).Value) = False Then
                    If rowDGV.Cells(e.ColumnIndex).Value.ToString.ToUpper = vData.ToUpper Then
                        If rowDGV.Index <> e.RowIndex Then
                            If rowDGV.Cells(Me.ColVariable_Type2.Index).Value = Me.DGVVariable.Item(Me.ColVariable_Type2.Index, e.RowIndex).Value Then
                                Return rowDGV.Index
                            End If
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

    Private Function GetIndexItemConst(ByVal vData As String, ByVal vDgv As DataGridView, _
    ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In vDgv.Rows
                If String.IsNullOrEmpty(rowDGV.Cells(e.ColumnIndex).Value) = False Then
                    If rowDGV.Cells(e.ColumnIndex).Value.ToString.ToUpper = vData.ToUpper Then
                        If rowDGV.Index <> e.RowIndex Then
                            If rowDGV.Cells(Me.ColConst_Type2.Index).Value = Me.DGVConstaint.Item(Me.ColConst_Type2.Index, e.RowIndex).Value Then
                                Return rowDGV.Index
                            End If
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

    Private Sub SetNoDgv(ByVal vDGV As DataGridView)
        Dim row As DataGridViewRow
        Try
            For Each row In vDGV.Rows
                row.Cells(Me.ColVariable_No.Index).Value = row.Index + 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function CekValueUnit(ByVal vUnit As String, ByVal vValue As String) As Boolean
        Try
            If vUnit.ToLower = "time" Then
                If Me.Objcd.Is_Time_String(vValue) = True Then
                    If Me.Objcd.Convert_ToMinutes(vValue) > 1439 Then
                        MsgBox("Value not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    Return True
                Else
                    MsgBox("Please input hh.mm format", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            ElseIf vUnit.ToLower = "duration" Then
                If Me.Objcd.Is_Time_String(vValue) = True Then
                    Return True
                Else
                    MsgBox("Please input hh.mm format", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            ElseIf vUnit.ToLower = "yes/no" Then
                If vValue.ToLower = "yes" Or vValue.ToLower = "no" Then
                    Return True
                Else
                    MsgBox("Please input yes or no", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            ElseIf vUnit.ToLower = "day of week" Then
                If vValue = "Sunday" Or vValue = "Monday" Or vValue = "Tuesday" _
                Or vValue = "Wednesday" Or vValue = "Thursday" Or vValue = "Friday" Or vValue = "Saturday" Then
                    Return True
                Else
                    MsgBox("Please input name of the days.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            ElseIf vUnit.ToString = "0/1/0.5".ToString Then
                If vValue = "0" Or vValue = "1" Or vValue = "0.5" Then
                    Return True
                Else
                    MsgBox("Please input 0 / 1 / 0.5.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            ElseIf vUnit.ToLower = "days" Then
                If Me.Objcd.CekNonNegatif(vValue) = False Then
                    MsgBox("Please input non negative number", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            ElseIf vUnit.ToLower = "male/female" Then
                If vValue.ToLower = "male" Or vValue.ToLower = "female" Then
                    Return True
                Else
                    MsgBox("Please input male or female.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            ElseIf vUnit.ToLower = "numeric" Then
                If Me.Objcd.CekNonNegatif(vValue) = False Then
                    MsgBox("Please input non negative number", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            ElseIf vUnit.ToLower = "single/married" Then
                If vValue.ToLower = "single" Or vValue.ToLower = "married" Then
                    Return True
                Else
                    MsgBox("Please input single or married.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            ElseIf vUnit.ToLower = "probation/contract/jobholder" Then
                If vValue.ToLower = "probation" Or vValue.ToLower = "contract" Or vValue.ToLower = "jobholder" Then
                    Return True
                Else
                    MsgBox("Please input probation/contract/jobholder.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SetDefValueUnit(ByVal vUnit As String, _
    ByVal vDefValue As Windows.Forms.DataGridViewCell)
        Try
            If vUnit.ToLower = "time" Then
                vDefValue.Value = "00.00"
            ElseIf vUnit.ToLower = "duration" Then
                vDefValue.Value = "00.00"
            ElseIf vUnit.ToLower = "yes/no" Then
                vDefValue.Value = "yes"
            ElseIf vUnit.ToLower = "day of week" Then
                vDefValue.Value = "Sunday"
            ElseIf vUnit.ToLower = "0/1/0.5" Then
                vDefValue.Value = "0"
            ElseIf vUnit.ToLower = "days" Then
                vDefValue.Value = "0"
            ElseIf vUnit.ToLower = "male/female" Then
                vDefValue.Value = "male"
            ElseIf vUnit.ToLower = "numeric" Then
                vDefValue.Value = "0"
            ElseIf vUnit.ToLower = "single/married" Then
                vDefValue.Value = "Single"
            ElseIf vUnit.ToLower = "probation/contract/jobholder" Then
                vDefValue.Value = "probation"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Dim CboDgv As DataGridViewComboBoxCell
        Dim i As Integer = 0
        Try
            'kolom no 
            Me.SetNoDgv(Me.DGVVariable)
            If Me.DGVVariable.RowCount > 0 Then
                'kolom type
                CboDgv = Me.DGVVariable.Item(Me.ColVariable_Type.Index, Me.DGVVariable.Rows.Count - 1)
                Me.DGVVariable.Item(Me.ColVariable_Type.Index, Me.DGVVariable.Rows.Count - 1).Value = _
                CboDgv.Items.Item(mType)
                'rounding unit
                Me.DGVVariable.Item(Me.ColRoundingUnit_Variable.Index, Me.DGVVariable.Rows.Count - 1).Value = 1
                'kolom rounding type
                CboDgv = Me.DGVVariable.Item(Me.ColRoundingTipe_Variable.Index, Me.DGVVariable.Rows.Count - 1)
                Me.DGVVariable.Item(Me.ColRoundingTipe_Variable.Index, Me.DGVVariable.Rows.Count - 1).Value = _
                CboDgv.Items.Item(i)
                Me.DGVVariable.Item(Me.ColVariable_Primary.Index, Me.DGVVariable.Rows.Count - 1).Value = -1
            End If


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
            'kolom no 
            Me.SetNoDgv(Me.DGVConstaint)
            If Me.DGVConstaint.RowCount > 0 Then
                'kolom type
                CboDgv = Me.DGVConstaint.Item(Me.ColConst_Type.Index, Me.DGVConstaint.Rows.Count - 1)
                Me.DGVConstaint.Item(Me.ColConst_Type.Index, Me.DGVConstaint.Rows.Count - 1).Value = _
                CboDgv.Items.Item(mType)
                'kolom compare
                CboDgv = Me.DGVConstaint.Item(Me.ColConst_Compare.Index, Me.DGVConstaint.Rows.Count - 1)
                Me.DGVConstaint.Item(Me.ColConst_Compare.Index, Me.DGVConstaint.Rows.Count - 1).Value = _
                CboDgv.Items.Item(i)
                Me.DGVConstaint.Item(Me.ColConst_Primary.Index, Me.DGVConstaint.Rows.Count - 1).Value = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            If Me.EraseRowIndex <> Me.mDGV.Rows.Count - 1 Then
                Me.mDGV.Rows.RemoveAt(Me.EraseRowIndex)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        'Dim KoleksiNamaData As New Collection
        'Dim KoleksiTextBox As New Collection
        Dim ListStr As String = ""
        Dim RowDGV As DataGridViewRow
        Try
            'tidak boleh kosong
            'KoleksiNamaData.Add("Yearly Max")
            'KoleksiTextBox.Add(Me.TxtYearlyMaxValue)
            'KoleksiNamaData.Add("Yearly Min")
            'KoleksiTextBox.Add(Me.TxtYearlyMinValue)
            'KoleksiNamaData.Add("Monthly Max")
            'KoleksiTextBox.Add(Me.TxtMonthlyMaxValue)
            'KoleksiNamaData.Add("Monthly Min")
            'KoleksiTextBox.Add(Me.TxtMonthlyMinValue)

            'If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
            '    Return False
            'End If
            'KoleksiNamaData.Clear()
            'KoleksiTextBox.Clear()
            If Me.Objcd.CekNumerik(Me.TxtYearlyMaxValue.Text) = True And Me.Objcd.CekNumerik(Me.TxtYearlyMinValue.Text) = True Then
                If CDec(Me.TxtYearlyMinValue.Text) > CDec(Me.TxtYearlyMaxValue.Text) Then
                    MsgBox("Yearly min value can't more than yearly max value", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                If Me.TxtYearlyMaxValue.Text = "" Then
                    MsgBox("Yearly max value can't fill empty string", MsgBoxStyle.Information, "Information")
                    Me.TxtYearlyMaxValue.Focus()
                    Return False
                ElseIf Me.TxtYearlyMinValue.Text = "" Then
                    MsgBox("Yearly min value can't fill empty string", MsgBoxStyle.Information, "Information")
                    Me.TxtYearlyMinValue.Focus()
                    Return False
                End If
            End If

            If Me.Objcd.CekNumerik(Me.TxtMonthlyMaxValue.Text) = True And Me.Objcd.CekNumerik(Me.TxtMonthlyMinValue.Text) = True Then
                If CDec(Me.TxtMonthlyMinValue.Text) > CDec(Me.TxtMonthlyMaxValue.Text) Then
                    MsgBox("Monthly min value can't more than Monthly max value", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                If Me.TxtMonthlyMaxValue.Text = "" Then
                    MsgBox("Monthly max value can't fill empty string", MsgBoxStyle.Information, "Information")
                    Me.TxtMonthlyMaxValue.Focus()
                    Return False
                ElseIf Me.TxtMonthlyMinValue.Text = "" Then
                    MsgBox("Monthly min value can't fill empty string", MsgBoxStyle.Information, "Information")
                    Me.TxtMonthlyMinValue.Focus()
                    Return False
                End If
            End If


            ListStr = Me.GetListVariable(Me.TxtFormula.Text)
            If ListStr <> "" Then
                MsgBox("The following variables in formula : " & ListStr & " is not defined", MsgBoxStyle.Information, "Error")
                Return False
            End If
            ListStr = Me.GetListVariable(Me.TxtYearlyMaxValue.Text)
            If ListStr <> "" Then
                MsgBox("The following variables in yearly max value : " & ListStr & " is not defined", MsgBoxStyle.Information, "Error")
                Return False
            End If
            ListStr = Me.GetListVariable(Me.TxtYearlyMinValue.Text)
            If ListStr <> "" Then
                MsgBox("The following variables in yearly minimum value : " & ListStr & " is not defined", MsgBoxStyle.Information, "Error")
                Return False
            End If
            ListStr = Me.GetListVariable(Me.TxtMonthlyMaxValue.Text)
            If ListStr <> "" Then
                MsgBox("The following variables in monthly max value : " & ListStr & " is not defined", MsgBoxStyle.Information, "Error")
                Return False
            End If
            ListStr = Me.GetListVariable(Me.TxtMonthlyMinValue.Text)
            If ListStr <> "" Then
                MsgBox("The following variables in monthly minimum value : " & ListStr & " is not defined", MsgBoxStyle.Information, "Error")
                Return False
            End If

            If Me.CboRoundingUnit.Text = "" Then
                MsgBox("Please select rounding unit.", MsgBoxStyle.Information, "Error")
                Return False
            End If

            'Batasan Constraint Ref Value Harus disi
            If Me.DGVConstaint.Rows.Count > 1 Then
                For i As Integer = 0 To Me.DGVConstaint.Rows.Count - 2
                    RowDGV = Me.DGVConstaint.Rows(i)
                    If String.IsNullOrEmpty(RowDGV.Cells(Me.ColConst_RefValue.Index).Value) = True Then
                        MsgBox("Please input ref value.", MsgBoxStyle.Information, "Information")
                        Me.DGVConstaint.ClearSelection()
                        RowDGV.Selected = True
                        Return False
                    End If
                Next
            End If

            'Batasan Jika Ada Total Leave Balance, maka di constraint di harus ada leave 
            Dim KetemuTotalLeaveBalance As Boolean = False
            Dim KetemuLeaveCodeConstraint As Boolean = False
            If Me.DGVVariable.Rows.Count > 1 Then
                For i As Integer = 0 To Me.DGVVariable.Rows.Count - 2
                    RowDGV = Me.DGVVariable.Rows(i)
                    If RowDGV.Cells(Me.ColVariable_PrimaryParameter.Index).Value = 132 Then
                        KetemuTotalLeaveBalance = True
                        Exit For
                    End If
                Next
                If KetemuTotalLeaveBalance = True Then
                    For i As Integer = 0 To Me.DGVConstaint.Rows.Count - 2
                        RowDGV = Me.DGVConstaint.Rows(i)
                        If RowDGV.Cells(Me.ColConst_PrimaryParameter.Index).Value = 133 Then
                            KetemuLeaveCodeConstraint = True
                            Exit For
                        End If
                    Next
                    If KetemuLeaveCodeConstraint = False Then
                        MsgBox("Please input data constraint with parameter leave code.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            'Tidak boleh input beda calculate period
            Dim ErrorMsg As String = ""
            For i As Integer = 0 To Me.DGVVariable.Rows.Count - 2
                RowDGV = Me.DGVVariable.Rows(i)
                If RowDGV.Cells(Me.ColVariable_Type.Index).Value <> Me.Buffer_CalculatePeriod Then
                    If RowDGV.Cells(Me.ColVariable_Type.Index).Value <> "Person Base" Then
                        ErrorMsg = "you cannot input variable data type " + RowDGV.Cells(Me.ColVariable_Type.Index).Value.ToString.ToLower + ", please input data type " + Me.Buffer_CalculatePeriod.ToString.ToLower + " or person base."
                        Exit For
                    End If
                End If
            Next
            If String.IsNullOrEmpty(ErrorMsg) = True Then
                For i As Integer = 0 To Me.DGVConstaint.Rows.Count - 2
                    RowDGV = Me.DGVConstaint.Rows(i)
                    If RowDGV.Cells(Me.ColConst_Type.Index).Value <> Me.Buffer_CalculatePeriod Then
                        If RowDGV.Cells(Me.ColConst_Type.Index).Value <> "Person Base" Then
                            If ErrorMsg = "" Then
                                ErrorMsg = "Inconsistent data calculation period for the following constraint: "
                            End If
                            ErrorMsg = ErrorMsg + vbCrLf + RowDGV.Cells(Me.ColConst_ParameterCode.Index).Value.ToString + " at row: " + (RowDGV.Index + 1).ToString
                            Me.DGVConstaint.ClearSelection()
                            RowDGV.Selected = True
                            Exit For
                        End If
                    End If
                Next
                If String.IsNullOrEmpty(ErrorMsg) = False Then
                    MsgBox(ErrorMsg, MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                MsgBox(ErrorMsg, MsgBoxStyle.Information, "Information")
                Return False
            End If

            'testing formula
            If Me.TxtFormula.Text <> "" And IsNumeric(Me.TxtFormula.Text) = False Then
                If Me.ObjSalaryFormula.TestFormula(Me.TxtFormula.Text, Me.DGVVariable, Me) = False Then
                    Return False
                End If
            End If
            If Me.TxtYearlyMaxValue.Text <> "" And IsNumeric(Me.TxtYearlyMaxValue.Text) = False Then
                If Me.ObjSalaryFormula.TestFormula(Me.TxtYearlyMaxValue.Text, Me.DGVVariable, Me) = False Then
                    Return False
                End If
            End If
            If Me.TxtYearlyMinValue.Text <> "" And IsNumeric(Me.TxtYearlyMinValue.Text) = False Then
                If Me.ObjSalaryFormula.TestFormula(Me.TxtYearlyMinValue.Text, Me.DGVVariable, Me) = False Then
                    Return False
                End If
            End If
            If Me.TxtMonthlyMaxValue.Text <> "" And IsNumeric(Me.TxtMonthlyMaxValue.Text) = False Then
                If Me.ObjSalaryFormula.TestFormula(Me.TxtMonthlyMaxValue.Text, Me.DGVVariable, Me) = False Then
                    Return False
                End If
            End If
            If Me.TxtMonthlyMinValue.Text <> "" And IsNumeric(Me.TxtMonthlyMinValue.Text) = False Then
                If Me.ObjSalaryFormula.TestFormula(Me.TxtMonthlyMinValue.Text, Me.DGVVariable, Me) = False Then
                    Return False
                End If
            End If

            'sampai sini
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub Hitung()
        'i.Add ("A",
    End Sub

    Private Sub TxtRoundingUnit_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRoundingUnit.Enter
        Try
            Me.Buffer = Me.TxtRoundingUnit.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRoundingUnit_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRoundingUnit.GotFocus
        Me.TxtRoundingUnit.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtRoundingUnit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtRoundingUnit.KeyDown
        Try
            Me.Objcd.Number_KeyDown(Me.TxtRoundingUnit, sender, e, 999999999999999)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRoundingUnit_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtRoundingUnit.KeyUp
        Try
            Me.Objcd.Number_KeyUp(Me.TxtRoundingUnit, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRoundingUnit_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRoundingUnit.Leave
        Try
            If Me.Objcd.CekBilanganPositif(Me.TxtRoundingUnit.Text) = False Then
                MsgBox("Please input positive number", MsgBoxStyle.Information, "Information")
                Me.TxtRoundingUnit.Text = Me.Buffer
                Me.TxtRoundingUnit.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtFormula_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFormula.Enter
        Try
            Me.Buffer = Me.TxtFormula.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtFormula_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFormula.GotFocus
        Me.TxtFormula.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtFormula_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFormula.Leave
        Try
            If Me.Objcd.CekNumerik(Me.TxtFormula.Text) = True Then
                If Me.Objcd.CekNonNegatif(Me.TxtFormula.Text) = False Then
                    MsgBox("Please input positif number", MsgBoxStyle.Information, "Information")
                    Me.TxtFormula.Text = Me.Buffer
                    Me.TxtFormula.Focus()
                Else
                    If Me.TxtFormula.Text > 999999999999999 Then
                        MsgBox("Max value is 999,999,999,999,999", MsgBoxStyle.Information, "Information")
                        Me.TxtFormula.Text = Me.Buffer
                        Me.TxtFormula.Focus()
                    Else
                        Me.TxtFormula.Text = Me.Objcd.DigitGrouping(Me.TxtFormula.Text, Me.DefPrec)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtYearlyMaxValue_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtYearlyMaxValue.Enter
        Try
            Me.Buffer = Me.TxtYearlyMaxValue.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtYearlyMaxValue_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtYearlyMaxValue.GotFocus
        Me.TxtYearlyMaxValue.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtYearlyMaxValue_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtYearlyMaxValue.Leave
        Try
            If Me.Objcd.CekNumerik(Me.TxtYearlyMaxValue.Text) = True Then
                If Me.Objcd.CekNonNegatif(Me.TxtYearlyMaxValue.Text) = False Then
                    MsgBox("Please input positif number", MsgBoxStyle.Information, "Information")
                    Me.TxtYearlyMaxValue.Text = Me.Buffer
                    Me.TxtYearlyMaxValue.Focus()
                Else
                    If Me.TxtYearlyMaxValue.Text > 999999999999999 Then
                        MsgBox("Yearly max value must not be larger than 999,999,999,999,999", MsgBoxStyle.Information, "Information")
                        Me.TxtYearlyMaxValue.Text = Me.Buffer
                        Me.TxtYearlyMaxValue.Focus()
                    Else
                        Me.TxtYearlyMaxValue.Text = Me.Objcd.DigitGrouping(Me.TxtYearlyMaxValue.Text, Me.DefPrec)
                    End If
                End If
            Else
                If Me.TxtYearlyMaxValue.Text = "" Then
                    Me.DtpYear.Enabled = False
                Else
                    Me.DtpYear.Enabled = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtYearlyMinValue_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtYearlyMinValue.Enter
        Try
            Me.Buffer = Me.TxtYearlyMinValue.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtYearlyMinValue_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtYearlyMinValue.GotFocus
        Me.TxtYearlyMinValue.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtYearlyMinValue_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtYearlyMinValue.Leave
        Try
            If Me.Objcd.CekNumerik(Me.TxtYearlyMinValue.Text) = True Then
                If Me.Objcd.CekNonNegatif(Me.TxtYearlyMinValue.Text) = False Then
                    MsgBox("Please input positif number", MsgBoxStyle.Information, "Information")
                    Me.TxtYearlyMinValue.Text = Me.Buffer
                    Me.TxtYearlyMinValue.Focus()
                Else
                    If CDec(Me.TxtYearlyMinValue.Text) > 999999999999999 Then
                        MsgBox("Yearly min value can't more than 999,999,999,999,999", MsgBoxStyle.Information, "Information")
                        Me.TxtYearlyMinValue.Text = Me.Buffer
                        Me.TxtYearlyMinValue.Focus()
                    Else
                        Me.TxtYearlyMinValue.Text = Me.Objcd.DigitGrouping(Me.TxtYearlyMinValue.Text, Me.DefPrec)
                    End If
                End If
            Else
                If Me.TxtYearlyMinValue.Text = "" Then
                    Me.DtpYear.Enabled = False
                Else
                    Me.DtpYear.Enabled = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMonthlyMaxValue_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMonthlyMaxValue.Enter
        Try
            Me.Buffer = Me.TxtMonthlyMaxValue.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMonthlyMaxValue_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMonthlyMaxValue.GotFocus
        Me.TxtMonthlyMaxValue.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtMonthlyMaxValue_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMonthlyMaxValue.Leave
        Try
            If String.IsNullOrEmpty(Me.TxtMonthlyMaxValue.Text) = False Then


                If Me.Objcd.CekNumerik(Me.TxtMonthlyMaxValue.Text) = True Then
                    If Me.Objcd.CekNonNegatif(Me.TxtMonthlyMaxValue.Text) = False Then
                        MsgBox("Please input positif number", MsgBoxStyle.Information, "Information")
                        Me.TxtMonthlyMaxValue.Text = Me.Buffer
                        Me.TxtMonthlyMaxValue.Focus()
                    Else
                        If Me.TxtMonthlyMaxValue.Text > 999999999999999 Then
                            MsgBox("Monthly max value must not be larger than 999,999,999,999,999", MsgBoxStyle.Information, "Information")
                            Me.TxtMonthlyMaxValue.Text = Me.Buffer
                            Me.TxtMonthlyMaxValue.Focus()
                        Else
                            Me.TxtMonthlyMaxValue.Text = Me.Objcd.DigitGrouping(Me.TxtMonthlyMaxValue.Text, Me.DefPrec)
                        End If
                    End If
                End If
            Else
                Me.TxtMonthlyMaxValue.Text = Me.Objcd.DigitGrouping(0, Me.DefPrec)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMonthlyMinValue_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMonthlyMinValue.Enter
        Try
            Me.Buffer = Me.TxtMonthlyMinValue.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMonthlyMinValue_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMonthlyMinValue.GotFocus
        Me.TxtMonthlyMinValue.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtMonthlyMinValue_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMonthlyMinValue.Leave
        Try
            If String.IsNullOrEmpty(Me.TxtMonthlyMinValue.Text) = False Then
                If Me.Objcd.CekNumerik(Me.TxtMonthlyMinValue.Text) = True Then
                    If Me.Objcd.CekNonNegatif(Me.TxtMonthlyMinValue.Text) = False Then
                        MsgBox("Please input positif number", MsgBoxStyle.Information, "Information")
                        Me.TxtMonthlyMinValue.Text = Me.Buffer
                        Me.TxtMonthlyMinValue.Focus()
                    Else
                        If CDec(Me.TxtMonthlyMinValue.Text) > 999999999999999 Then
                            MsgBox("Monthly min value can't more than 999,999,999,999,999", MsgBoxStyle.Information, "Information")
                            Me.TxtMonthlyMinValue.Text = Me.Buffer
                            Me.TxtMonthlyMinValue.Focus()
                        Else
                            Me.TxtMonthlyMinValue.Text = Me.Objcd.DigitGrouping(Me.TxtMonthlyMinValue.Text, Me.DefPrec)
                        End If
                    End If
                End If
            Else
                Me.TxtMonthlyMinValue.Text = Me.Objcd.DigitGrouping(0, Me.DefPrec)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVVariable_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DGVVariable.CellBeginEdit
        Dim TipeVariable As ClsMasterSalaryFormula_item.EnumTipeConstraitVariable = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode
        Dim TipeKompetensi As Core_Module.ClsMasterKompetensi.EnumDataType = ClsMasterKompetensi.EnumDataType.Alfabetic

        Try
            If String.IsNullOrEmpty(Me.DGVVariable.Item(Me.ColVariable_Parameter_Code.Index, e.RowIndex).Value) = True Then
                If e.ColumnIndex = Me.ColVariable_No.Index Or e.ColumnIndex = Me.ColVariable_Parameter_Desc.Index _
                Or e.ColumnIndex = Me.ColVariable_Type.Index Or e.ColumnIndex = Me.ColVariable_Variable.Index _
                Or e.ColumnIndex = Me.ColRoundingTipe_Variable.Index Or e.ColumnIndex = Me.ColRoundingUnit_Variable.Index Or _
                e.ColumnIndex = Me.ColVariable_Unit.Index Then
                    e.Cancel = True
                Else
                    If e.ColumnIndex = Me.ColConst_ParameterCode.Index Then
                        If String.IsNullOrEmpty(Me.DGVVariable.Item(Me.ColVariable_Type2.Index, e.RowIndex).Value) = True Then
                            e.Cancel = True
                        Else
                            Me.Buffer = Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value
                        End If
                    ElseIf e.ColumnIndex = Me.ColVariable_Type2.Index Then
                        If e.RowIndex = 0 Then
                            Me.Buffer = Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value
                        Else
                            If IsNothing(Me.DGVVariable.Item(Me.ColVariable_Parameter_Code.Index, e.RowIndex - 1).Value) = True Then
                                e.Cancel = True
                            Else
                                Me.Buffer = Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value
                            End If
                        End If
                    Else
                        e.Cancel = True
                    End If

                    Me.Buffer = Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            Else
                TipeVariable = Me.ObjSalaryFormula.GetTipeConstrainVariable(Me.DGVVariable.Item(Me.ColVariable_Type2.Index, e.RowIndex).Value)
                If TipeVariable = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode Then
                    If Me.DGVVariable.Item(Me.ColVariable_PrimaryParameter.Index, e.RowIndex).Value <> 115 And _
                    Me.DGVVariable.Item(Me.ColVariable_PrimaryParameter.Index, e.RowIndex).Value <> 75 Then
                        If e.ColumnIndex = Me.ColVariable_No.Index Or e.ColumnIndex = Me.ColVariable_Parameter_Desc.Index _
                          Or e.ColumnIndex = Me.ColVariable_Type.Index Or e.ColumnIndex = Me.ColVariable_Unit.Index Then
                            e.Cancel = True
                        Else
                            Me.Buffer = Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value
                        End If
                    Else
                        If e.ColumnIndex = Me.ColRoundingTipe_Variable.Index Or e.ColumnIndex = Me.ColRoundingUnit_Variable.Index Then
                            e.Cancel = True
                        Else
                            Me.Buffer = Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value
                        End If
                    End If
                ElseIf TipeVariable = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi Then
                    TipeKompetensi = Me.ObjKompetensi.GetDataType(Me.DGVVariable.Item(Me.ColVariable_PrimaryParameter.Index, e.RowIndex).Value)
                    If TipeKompetensi = ClsMasterKompetensi.EnumDataType.Number Then
                        'bisa set rounding
                        If e.ColumnIndex = Me.ColVariable_No.Index Or e.ColumnIndex = Me.ColVariable_Parameter_Desc.Index _
                        Or e.ColumnIndex = Me.ColVariable_Type.Index Or e.ColumnIndex = Me.ColVariable_Unit.Index Or _
                        e.ColumnIndex = Me.ColVariable_Type2.Index Then
                            e.Cancel = True
                        Else
                            Me.Buffer = Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value
                        End If
                    Else
                        'tidak bisa set rounding
                        If e.ColumnIndex = Me.ColRoundingTipe_Variable.Index Or e.ColumnIndex = Me.ColRoundingUnit_Variable.Index Then
                            e.Cancel = True
                        Else
                            If e.ColumnIndex = Me.ColVariable_No.Index Or e.ColumnIndex = Me.ColVariable_Parameter_Desc.Index _
                            Or e.ColumnIndex = Me.ColVariable_Type.Index Or e.ColumnIndex = Me.ColVariable_Unit.Index Or _
                            e.ColumnIndex = Me.ColVariable_Type2.Index Then
                                e.Cancel = True
                            Else
                                Me.Buffer = Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CellEndEdit_FindCompetence_Variable(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim CboDgv As DataGridViewComboBoxCell
        Dim PrimaryParameter As Double = -1
        Dim rowDGV As DataGridViewRow

        Try
            CboDgv = Me.DGVVariable.Item(Me.ColVariable_Type.Index, e.RowIndex)
            rowDGV = Me.DGVVariable.Rows(e.RowIndex)

            If String.IsNullOrEmpty(Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                If e.RowIndex <> Me.DGVVariable.RowCount - 1 Then
                    MsgBox("Parameter code cannot input an empty string.", MsgBoxStyle.Information, "Information")
                End If
                Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                If Me.Buffer = "" Then
                    Me.mDGV = Me.DGVVariable
                    Me.EraseRowIndex = e.RowIndex
                    Me.BackgroundWorker2.RunWorkerAsync()
                    Me.BackgroundWorker2.Dispose()
                End If
            Else
                If Me.ObjKompetensi.ApakahSudahAda(Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    PrimaryParameter = Me.ObjKompetensi.GetPrimary(Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value)
                    If PrimaryParameter <> -1 Then
                        If Me.GetIndexItemVar(Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value, Me.DGVVariable, e) < 0 Then
                            Me.DGVVariable.Item(Me.ColVariable_PrimaryParameter.Index, e.RowIndex).Value = PrimaryParameter
                            Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjKompetensi.GetDataCode(PrimaryParameter)
                            Me.DGVVariable.Item(Me.ColVariable_Parameter_Desc.Index, e.RowIndex).Value = Me.ObjKompetensi.GetDesc(PrimaryParameter)
                            CboDgv = rowDGV.Cells(Me.ColVariable_Type.Index)
                            rowDGV.Cells(Me.ColVariable_Type.Index).Value = CboDgv.Items.Item(Me.mType)
                            Me.DGVVariable.Item(Me.ColVariable_Unit.Index, e.RowIndex).Value = ""
                        Else
                            MsgBox("Parameter code has been used, please change the parameter code.", MsgBoxStyle.Information, "Information")
                            Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If Me.Buffer = "" Then
                                Me.mDGV = Me.DGVVariable
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker2.RunWorkerAsync()
                                Me.BackgroundWorker2.Dispose()
                            End If
                        End If
                    End If
                Else
                    ''diketik tidak ada lalu tampilkan form cari
                    Dim ij As New FrmMasterKompetensi_Cari
                    ij.NumberOnly = True
                    ij.txtCari.Text = Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value
                    ij.BackColorAll = Me.BackColorAll
                    ij.ShowDialog()
                    If ij.mPrimary <> -1 Then
                        If Me.GetIndexItemVar(Me.ObjKompetensi.GetDataCode(ij.mPrimary), Me.DGVVariable, e) < 0 Then
                            Me.DGVVariable.Item(Me.ColVariable_PrimaryParameter.Index, e.RowIndex).Value = ij.mPrimary
                            Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjKompetensi.GetDataCode(ij.mPrimary)
                            Me.DGVVariable.Item(Me.ColVariable_Parameter_Desc.Index, e.RowIndex).Value = Me.ObjKompetensi.GetDesc(ij.mPrimary)
                            CboDgv = rowDGV.Cells(Me.ColVariable_Type.Index)
                            rowDGV.Cells(Me.ColVariable_Type.Index).Value = CboDgv.Items.Item(Me.mType)
                            Me.DGVVariable.Item(Me.ColVariable_Unit.Index, e.RowIndex).Value = ""
                        Else
                            MsgBox("Parameter code has been used, please change the parameter code.", MsgBoxStyle.Information, "Information")
                            Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If Me.Buffer = "" Then
                                Me.mDGV = Me.DGVVariable
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker2.RunWorkerAsync()
                                Me.BackgroundWorker2.Dispose()
                            End If
                        End If
                    Else
                        Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        If Me.Buffer = "" Then
                            Me.mDGV = Me.DGVVariable
                            Me.EraseRowIndex = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                            Me.BackgroundWorker2.Dispose()
                        End If
                    End If
                    ij = Nothing
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVVariable_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVVariable.CellEndEdit
        Dim PrimaryParameter As Double = -1
        Dim CboDgv As DataGridViewComboBoxCell
        Dim TipeVariable As ClsMasterSalaryFormula_item.EnumTipeConstraitVariable = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode

        Try
            'kolom parameter code
            If e.ColumnIndex = Me.ColVariable_Parameter_Code.Index Then
                TipeVariable = Me.ObjSalaryFormula.GetTipeConstrainVariable(Me.DGVVariable.Item(Me.ColVariable_Type2.Index, e.RowIndex).Value)
                If TipeVariable = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode Then
                    CboDgv = Me.DGVVariable.Item(Me.ColVariable_Type.Index, e.RowIndex)
                    If String.IsNullOrEmpty(Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        If e.RowIndex <> Me.DGVVariable.RowCount - 1 Then
                            MsgBox("Parameter code cannot input an empty string.", MsgBoxStyle.Information, "Information")
                        End If
                        Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        If Me.Buffer = "" Then
                            Me.mDGV = Me.DGVVariable
                            Me.EraseRowIndex = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                            Me.BackgroundWorker2.Dispose()
                        End If
                    Else
                        If Me.ObjParameter.ApakahSudahAda(Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value, Me.mType, _
                        ClsMasterParameterGaji.EnumIsVariable.true_, ClsMasterParameterGaji.EnumIsConstrain.False_) = True Then
                            PrimaryParameter = Me.ObjParameter.GetPrimaryByCode(Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value)
                            If PrimaryParameter <> -1 Then
                                If Me.GetIndexItemVar(Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value, Me.DGVVariable, e) < 0 Then
                                    Me.DGVVariable.Item(Me.ColVariable_PrimaryParameter.Index, e.RowIndex).Value = PrimaryParameter
                                    Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjParameter.GetCodeByPrimary(PrimaryParameter)
                                    Me.DGVVariable.Item(Me.ColVariable_Parameter_Desc.Index, e.RowIndex).Value = Me.ObjParameter.GetDescByPrimary(PrimaryParameter)
                                    Me.DGVVariable.Item(Me.ColVariable_Type.Index, e.RowIndex).Value = _
                                    CboDgv.Items.Item(Me.ObjParameter.GetTypeByPrimary(PrimaryParameter))
                                    Me.DGVVariable.Item(Me.ColVariable_Unit.Index, e.RowIndex).Value = Me.ObjParameter.GetUnitByPrimary(PrimaryParameter)
                                Else
                                    MsgBox("Parameter code has been used, please change the parameter code.", MsgBoxStyle.Information, "Information")
                                    Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    If Me.Buffer = "" Then
                                        Me.mDGV = Me.DGVVariable
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                        Me.BackgroundWorker2.Dispose()
                                    End If
                                End If
                            End If
                        Else
                            ''diketik tidak ada lalu tampilkan form cari
                            Dim ij As New FrmMasterSalaryFormulaParameter1_Cari
                            ij.txtCari.Text = Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.mVariable = ClsMasterParameterGaji.EnumIsVariable.true_
                            ij.mConstrain = ClsMasterParameterGaji.EnumIsConstrain.False_
                            ij.mType = Me.mType
                            ij.ShowDialog()
                            If ij.Primary <> -1 Then
                                If Me.GetIndexItemVar(Me.ObjParameter.GetCodeByPrimary(ij.Primary), Me.DGVVariable, e) < 0 Then
                                    Me.DGVVariable.Item(Me.ColVariable_PrimaryParameter.Index, e.RowIndex).Value = ij.Primary
                                    Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjParameter.GetCodeByPrimary(ij.Primary)
                                    Me.DGVVariable.Item(Me.ColVariable_Parameter_Desc.Index, e.RowIndex).Value = Me.ObjParameter.GetDescByPrimary(ij.Primary)
                                    Me.DGVVariable.Item(Me.ColVariable_Type.Index, e.RowIndex).Value = _
                                    CboDgv.Items.Item(Me.ObjParameter.GetTypeByPrimary(ij.Primary))
                                    Me.DGVVariable.Item(Me.ColVariable_Unit.Index, e.RowIndex).Value = Me.ObjParameter.GetUnitByPrimary(ij.Primary)
                                Else
                                    MsgBox("Parameter code has been used, please change the parameter code.", MsgBoxStyle.Information, "Information")
                                    Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    If Me.Buffer = "" Then
                                        Me.mDGV = Me.DGVVariable
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                        Me.BackgroundWorker2.Dispose()
                                    End If
                                End If
                            Else
                                Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                If Me.Buffer = "" Then
                                    Me.mDGV = Me.DGVVariable
                                    Me.EraseRowIndex = e.RowIndex
                                    Me.BackgroundWorker2.RunWorkerAsync()
                                    Me.BackgroundWorker2.Dispose()
                                End If
                            End If
                            ij = Nothing
                        End If
                    End If
                ElseIf TipeVariable = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi Then
                    Me.CellEndEdit_FindCompetence_Variable(e)
                End If
            End If

            'variable: tidak ada masalah untuk hardcode, competence, custom value atau service record
            If e.ColumnIndex = Me.ColVariable_Variable.Index Then
                If String.IsNullOrEmpty(Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'tidak input selain huruf
                    If Me.Objcd.CekHuruf(Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper) = True Then
                        'tidak boleh ada yang sama
                        If Me.GetIndexItemVar(Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper, Me.DGVVariable, e) < 0 Then
                            'cek apakah berubah dari nilai sebelumnya
                            If Me.Buffer <> "" Then
                                If Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper <> Me.Buffer Then
                                    'jika rubah, rubah variable yang ada d formula
                                    'formula
                                    Me.TxtFormula.Text = Me.Objcd.GetFormulaBaru _
                                    (Me.TxtFormula.Text, Me.Buffer, Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper)
                                    'yearly max value
                                    If Me.Objcd.CekNumerik(Me.TxtYearlyMaxValue.Text) = False Then
                                        Me.TxtYearlyMaxValue.Text = Me.Objcd.GetFormulaBaru _
                                        (Me.TxtYearlyMaxValue.Text, Me.Buffer, Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper)
                                    End If
                                    'yearly min value
                                    If Me.Objcd.CekNumerik(Me.TxtYearlyMaxValue.Text) = False Then
                                        Me.TxtYearlyMinValue.Text = Me.Objcd.GetFormulaBaru _
                                        (Me.TxtYearlyMinValue.Text, Me.Buffer, Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper)
                                    End If
                                    'mounthly max value
                                    If Me.Objcd.CekNumerik(Me.TxtMonthlyMaxValue.Text) = False Then
                                        Me.TxtMonthlyMaxValue.Text = Me.Objcd.GetFormulaBaru _
                                        (Me.TxtMonthlyMaxValue.Text, Me.Buffer, Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper)
                                    End If
                                    'mounthly min value
                                    If Me.Objcd.CekNumerik(Me.TxtMonthlyMinValue.Text) = False Then
                                        Me.TxtMonthlyMinValue.Text = Me.Objcd.GetFormulaBaru _
                                        (Me.TxtMonthlyMinValue.Text, Me.Buffer, Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper)
                                    End If
                                End If
                            End If
                            Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value = Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        Else
                            MsgBox("Variable has been used, please change the variable", MsgBoxStyle.Information, "Information")
                            Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                    Else
                        MsgBox("please input alphabetic only", MsgBoxStyle.Information, "information")
                        Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            End If

            'rounding unit
            If e.ColumnIndex = Me.ColRoundingUnit_Variable.Index Then
                If Me.Objcd.CekBilanganPositif(Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please input positif number.", MsgBoxStyle.Information, "Information")
                    Me.DGVVariable.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVVariable_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DGVVariable.DataError
        'untuk handle ketika baru buka form dan mau dinomorin
    End Sub

    Private Sub DGVVariable_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DGVVariable.RowsAdded
        Try
            If Me.BackgroundWorker1.IsBusy = False Then
                Me.BackgroundWorker1.Dispose()
                Me.BackgroundWorker1.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVVariable_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DGVVariable.RowsRemoved
        Try
            If Me.BackgroundWorker1.IsBusy = False Then
                Me.BackgroundWorker1.Dispose()
                Me.BackgroundWorker1.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVConstaint_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DGVConstaint.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.DGVConstaint.Item(Me.ColConst_ParameterCode.Index, e.RowIndex).Value) = True Then
                If e.ColumnIndex = Me.ColConst_ParameterCode.Index Then
                    If String.IsNullOrEmpty(Me.DGVConstaint.Item(Me.ColConst_Type2.Index, e.RowIndex).Value) = True Then
                        e.Cancel = True
                    Else
                        Me.Buffer = Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                ElseIf e.ColumnIndex = Me.ColConst_Type2.Index Then
                    If e.RowIndex = 0 Then
                        Me.Buffer = Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        If IsNothing(Me.DGVConstaint.Item(Me.ColConst_ParameterCode.Index, e.RowIndex - 1).Value) = True Then
                            e.Cancel = True
                        Else
                            Me.Buffer = Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value
                        End If
                    End If
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColConst_ParameterCode.Index Or _
                e.ColumnIndex = Me.ColConst_Compare.Index Or e.ColumnIndex = Me.ColConst_RefValue.Index Then
                    Me.Buffer = Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CellEndEdit_RefValue_Contraint(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim TipeKompetensi As Core_Module.ClsMasterKompetensi.EnumDataType = ClsMasterKompetensi.EnumDataType.Alfabetic
        Dim rowDGV As DataGridViewRow

        Try
            If Me.CekValueUnit(Me.DGVConstaint.Item(Me.ColConst_Unit.Index, e.RowIndex).Value, Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            Else
                rowDGV = Me.DGVConstaint.Rows(e.RowIndex)
                TipeKompetensi = Me.ObjKompetensi.GetDataType(rowDGV.Cells(Me.ColConst_PrimaryParameter.Index).Value)
                If String.IsNullOrEmpty(Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    DefPrec = Me.ObjKompetensi.GetDefPrec(Me.DGVConstaint.Item(Me.ColConst_PrimaryParameter.Index, e.RowIndex).Value)
                    TipeKompetensi = Me.ObjKompetensi.GetDataType(Me.DGVConstaint.Item(Me.ColConst_PrimaryParameter.Index, e.RowIndex).Value)
                    If TipeKompetensi = ClsMasterKompetensi.EnumDataType.Alfabetic Then
                        If Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "A" Or _
                        Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "B" Or _
                        Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "C" Or _
                        Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "D" Or _
                        Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "E" Or _
                        Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "F" Then
                            Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        Else
                            MsgBox("Must input A , B , C ,D , E ,F", MsgBoxStyle.Information, "Information")
                            Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                    ElseIf TipeKompetensi = ClsMasterKompetensi.EnumDataType.Number Then
                        If Me.Objcd.CekNumerik(Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Must input number format", MsgBoxStyle.Information, "Information")
                            Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Else
                            Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.ObjKompetensi.SetFormatDataCustom(Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value, TipeKompetensi, DefPrec, Me.Objcd)
                        End If
                    ElseIf TipeKompetensi = ClsMasterKompetensi.EnumDataType.Tanggal Then
                        'set region ke indonesia
                        If IsDate(Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Must input date format.", MsgBoxStyle.Information, "Information")
                            Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Else
                            Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.ObjKompetensi.SetFormatDataCustom(Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value, TipeKompetensi, DefPrec, Me.Objcd)
                        End If
                    ElseIf TipeKompetensi = ClsMasterKompetensi.EnumDataType.Text Then
                        Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = _
                        Me.ObjKompetensi.SetFormatDataCustom(Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value, TipeKompetensi, DefPrec, Me.Objcd)
                    ElseIf TipeKompetensi = ClsMasterKompetensi.EnumDataType.YN Then
                        If Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "YES" Then
                            Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = "Yes"
                        ElseIf Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "NO" Then
                            Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = "No"
                        ElseIf Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "Y" Then
                            Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = "Yes"
                        ElseIf Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "N" Then
                            Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = "No"
                        Else
                            MsgBox("Must input Yes / No .", MsgBoxStyle.Information, "Information")
                            Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                    End If
                Else
                    Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CellEndEdit_Compare_Constraint(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim CboDgv As DataGridViewComboBoxCell
        Dim TipeKompetensi As Core_Module.ClsMasterKompetensi.EnumDataType = ClsMasterKompetensi.EnumDataType.Alfabetic
        Dim rowDGV As DataGridViewRow

        Try
            Dim BuffCbo As Integer = 0
            rowDGV = Me.DGVConstaint.Rows(e.RowIndex)
            TipeKompetensi = Me.ObjKompetensi.GetDataType(rowDGV.Cells(Me.ColConst_PrimaryParameter.Index).Value)
            If TipeKompetensi = ClsMasterKompetensi.EnumDataType.YN Or _
            TipeKompetensi = ClsMasterKompetensi.EnumDataType.Text Then
                CboDgv = Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex)
                If Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = CboDgv.Items.Item(EnumCompare.a) Or _
                Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = CboDgv.Items.Item(EnumCompare.f) Then
                Else
                    MsgBox("comparation can only choose = or <>", MsgBoxStyle.Information, "information")
                    Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = CboDgv.Items.Item(EnumCompare.a)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CellEndEdit_FindCompetence_Constraint(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim CboDgv As DataGridViewComboBoxCell
        Dim rowDGV As DataGridViewRow
        Dim PrimaryParameter As Double = -1

        Try
            CboDgv = Me.DGVConstaint.Item(Me.ColConst_Type.Index, e.RowIndex)
            If String.IsNullOrEmpty(Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                MsgBox("Parameter code cannot input an empty string.", MsgBoxStyle.Information, "Information")
                Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                If Me.Buffer = "" Then
                    Me.mDGV = Me.DGVConstaint
                    Me.EraseRowIndex = e.RowIndex
                    Me.BackgroundWorker2.RunWorkerAsync()
                    Me.BackgroundWorker2.Dispose()
                End If
            Else
                Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                rowDGV = Me.DGVConstaint.Rows(e.RowIndex)

                If Me.ObjKompetensi.ApakahSudahAda(Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    PrimaryParameter = Me.ObjKompetensi.GetPrimary(Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value)
                    If PrimaryParameter <> -1 Then
                        'If Me.GetIndexItem(Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value, Me.DGVConstaint, e) < 0 Then ''
                        Me.DGVConstaint.Item(Me.ColConst_PrimaryParameter.Index, e.RowIndex).Value = PrimaryParameter
                        Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjKompetensi.GetDataCode(PrimaryParameter)
                        Me.DGVConstaint.Item(Me.ColConst_ParameterDesc.Index, e.RowIndex).Value = Me.ObjKompetensi.GetDesc(PrimaryParameter)
                        CboDgv = rowDGV.Cells(Me.ColConst_Type.Index)
                        rowDGV.Cells(Me.ColConst_Type.Index).Value = CboDgv.Items.Item(Me.mType)
                        Me.DGVConstaint.Item(Me.ColConst_Unit.Index, e.RowIndex).Value = ""
                        Me.DGVConstaint.Item(Me.ColConst_RefValue.Index, e.RowIndex).Value = Me.ObjKompetensi.GetValue_CustomFormatted(PrimaryParameter, Me.Objcd, _
                        Me.ObjKompetensi.GetDataType(PrimaryParameter))
                    Else
                        Me.mDGV = Me.DGVConstaint
                        Me.EraseRowIndex = e.RowIndex
                        Me.BackgroundWorker2.RunWorkerAsync()
                        Me.BackgroundWorker2.Dispose()
                        'End If ''
                    End If
                Else
                    ''diketik tidak ada lalu tampilkan form cari
                    Dim ij As New Core_Module.FrmMasterKompetensi_Cari
                    ij.txtCari.Text = Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value
                    ij.BackColorAll = Me.BackColorAll
                    ij.ShowDialog()
                    If ij.mPrimary <> -1 Then
                        Me.DGVConstaint.Item(Me.ColConst_PrimaryParameter.Index, e.RowIndex).Value = ij.mPrimary
                        Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjKompetensi.GetDataCode(ij.mPrimary)
                        Me.DGVConstaint.Item(Me.ColConst_ParameterDesc.Index, e.RowIndex).Value = Me.ObjKompetensi.GetDesc(ij.mPrimary)
                        CboDgv = rowDGV.Cells(Me.ColConst_Type.Index)
                        rowDGV.Cells(Me.ColConst_Type.Index).Value = CboDgv.Items.Item(Me.mType)
                        Me.DGVConstaint.Item(Me.ColConst_Unit.Index, e.RowIndex).Value = ""
                        Me.DGVConstaint.Item(Me.ColConst_RefValue.Index, e.RowIndex).Value = Me.ObjKompetensi.GetValue_CustomFormatted(ij.mPrimary, Me.Objcd, _
                        Me.ObjKompetensi.GetDataType(ij.mPrimary))
                    Else
                        'Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer ''
                        'If Me.Buffer = "" Then ''
                        Me.mDGV = Me.DGVConstaint
                        Me.EraseRowIndex = e.RowIndex
                        Me.BackgroundWorker2.RunWorkerAsync()
                        Me.BackgroundWorker2.Dispose()
                        'End If ''
                    End If
                    ij = Nothing
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVConstaint_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVConstaint.CellEndEdit
        Dim PrimaryParameter As Double = -1
        Dim CboDgv As DataGridViewComboBoxCell
        Try
            'kolom parameter code
            If IsNothing(Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColConst_ParameterCode.Index Then
                    If Me.ObjSalaryFormula.GetTipeConstrainVariable(Me.DGVConstaint.Item(Me.ColConst_Type2.Index, e.RowIndex).Value) = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode Then
                        CboDgv = Me.DGVConstaint.Item(Me.ColConst_Type.Index, e.RowIndex)
                        If String.IsNullOrEmpty(Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            MsgBox("Parameter code cannot input an empty string.", MsgBoxStyle.Information, "Information")
                            Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If Me.Buffer = "" Then
                                Me.mDGV = Me.DGVConstaint
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker2.RunWorkerAsync()
                                Me.BackgroundWorker2.Dispose()
                            End If
                        Else
                            If Me.ObjParameter.ApakahSudahAda(Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value, Me.mType, _
                            ClsMasterParameterGaji.EnumIsVariable.False_, ClsMasterParameterGaji.EnumIsConstrain.true_) = True Then
                                PrimaryParameter = Me.ObjParameter.GetPrimaryByCode(Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value)
                                If PrimaryParameter <> -1 Then
                                    'If Me.GetIndexItem(Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value, Me.DGVConstaint, e) < 0 Then ''
                                    Me.DGVConstaint.Item(Me.ColConst_PrimaryParameter.Index, e.RowIndex).Value = PrimaryParameter
                                    Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjParameter.GetCodeByPrimary(PrimaryParameter)
                                    Me.DGVConstaint.Item(Me.ColConst_ParameterDesc.Index, e.RowIndex).Value = Me.ObjParameter.GetDescByPrimary(PrimaryParameter)
                                    Me.DGVConstaint.Item(Me.ColConst_Type.Index, e.RowIndex).Value = _
                                    CboDgv.Items.Item(Me.ObjParameter.GetTypeByPrimary(PrimaryParameter))
                                    Me.DGVConstaint.Item(Me.ColConst_Unit.Index, e.RowIndex).Value = Me.ObjParameter.GetUnitByPrimary(PrimaryParameter)
                                    Me.SetDefValueUnit(Me.DGVConstaint.Item(Me.ColConst_Unit.Index, e.RowIndex).Value, Me.DGVConstaint.Item(Me.ColConst_RefValue.Index, e.RowIndex))
                                Else
                                    Me.mDGV = Me.DGVConstaint
                                    Me.EraseRowIndex = e.RowIndex
                                    Me.BackgroundWorker2.RunWorkerAsync()
                                    Me.BackgroundWorker2.Dispose()
                                    'End If ''
                                End If
                            Else
                                ''diketik tidak ada lalu tampilkan form cari
                                Dim ij As New FrmMasterSalaryFormulaParameter1_Cari
                                ij.txtCari.Text = Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value
                                ij.BackColorAll = Me.BackColorAll
                                ij.mVariable = ClsMasterParameterGaji.EnumIsVariable.False_
                                ij.mConstrain = ClsMasterParameterGaji.EnumIsConstrain.true_
                                ij.mType = Me.mType
                                ij.ShowDialog()
                                If ij.Primary <> -1 Then
                                    'If Me.GetIndexItem(Me.ObjParameter.GetCodeByPrimary(ij.Primary), Me.DGVConstaint, e) < 0 Then ''
                                    Me.DGVConstaint.Item(Me.ColConst_PrimaryParameter.Index, e.RowIndex).Value = ij.Primary
                                    Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjParameter.GetCodeByPrimary(ij.Primary)
                                    Me.DGVConstaint.Item(Me.ColConst_ParameterDesc.Index, e.RowIndex).Value = Me.ObjParameter.GetDescByPrimary(ij.Primary)
                                    Me.DGVConstaint.Item(Me.ColConst_Type.Index, e.RowIndex).Value = _
                                    CboDgv.Items.Item(Me.ObjParameter.GetTypeByPrimary(ij.Primary))
                                    Me.DGVConstaint.Item(Me.ColConst_Unit.Index, e.RowIndex).Value = Me.ObjParameter.GetUnitByPrimary(ij.Primary)
                                    Me.SetDefValueUnit(Me.DGVConstaint.Item(Me.ColConst_Unit.Index, e.RowIndex).Value, Me.DGVConstaint.Item(Me.ColConst_RefValue.Index, e.RowIndex))
                                    'Else ''
                                    '    MsgBox("Parameter code has been used, please change the parameter code.", MsgBoxStyle.Information, "Information") ''
                                    '    Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer ''
                                    '    If Me.Buffer = "" Then ''
                                    '        Me.mDGV = Me.DGVConstaint ''
                                    '        Me.EraseRowIndex = e.RowIndex ''
                                    '        Me.BackgroundWorker2.RunWorkerAsync() ''
                                    '        Me.BackgroundWorker2.Dispose() ''
                                    '    End If ''
                                    'End If ''
                                Else
                                    'Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer ''
                                    'If Me.Buffer = "" Then ''
                                    Me.mDGV = Me.DGVConstaint
                                    Me.EraseRowIndex = e.RowIndex
                                    Me.BackgroundWorker2.RunWorkerAsync()
                                    Me.BackgroundWorker2.Dispose()
                                    'End If ''
                                End If
                                ij = Nothing
                            End If
                        End If
                    ElseIf Me.ObjSalaryFormula.GetTipeConstrainVariable(Me.DGVConstaint.Item(Me.ColConst_Type2.Index, e.RowIndex).Value) = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi Then
                        Me.CellEndEdit_FindCompetence_Constraint(e)
                    End If
                End If

                'compare
                If e.ColumnIndex = Me.ColConst_Compare.Index Then
                    If Me.ObjSalaryFormula.GetTipeConstrainVariable(Me.DGVConstaint.Item(Me.ColConst_Type2.Index, e.RowIndex).Value) = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode Then
                        Dim BuffCbo As Integer = 0
                        If Me.DGVConstaint.Item(Me.ColConst_Unit.Index, e.RowIndex).Value.ToString.ToLower = "yes/no" Or _
                        Me.DGVConstaint.Item(Me.ColConst_Unit.Index, e.RowIndex).Value.ToString = "Single/Married" Or _
                        Me.DGVConstaint.Item(Me.ColConst_Unit.Index, e.RowIndex).Value.ToString = "probation/contract/jobholder" Or _
                        Me.DGVConstaint.Item(Me.ColConst_Unit.Index, e.RowIndex).Value.ToString = "From master job position" Or _
                        Me.DGVConstaint.Item(Me.ColConst_Unit.Index, e.RowIndex).Value.ToString = "From master leave" Then
                            CboDgv = Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex)
                            If Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = CboDgv.Items.Item(EnumCompare.a) Or _
                            Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = CboDgv.Items.Item(EnumCompare.f) Then
                            Else
                                MsgBox("comparation can only choose = or <>", MsgBoxStyle.Information, "information")
                                Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = CboDgv.Items.Item(EnumCompare.a)
                            End If
                        End If
                    ElseIf Me.ObjSalaryFormula.GetTipeConstrainVariable(Me.DGVConstaint.Item(Me.ColConst_Type2.Index, e.RowIndex).Value) = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi Then
                        Me.CellEndEdit_Compare_Constraint(e)
                    End If
                End If

                'kolom ref value
                If e.ColumnIndex = Me.ColConst_RefValue.Index Then
                    If Me.ObjSalaryFormula.GetTipeConstrainVariable(Me.DGVConstaint.Item(Me.ColConst_Type2.Index, e.RowIndex).Value) = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode Then
                        If Me.CekValueUnit(Me.DGVConstaint.Item(Me.ColConst_Unit.Index, e.RowIndex).Value, Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Else
                            If Trim(Me.DGVConstaint.Item(Me.ColConst_Unit.Index, e.RowIndex).Value.ToString.ToLower) <> "from master job position" And _
                            Trim(Me.DGVConstaint.Item(Me.ColConst_Unit.Index, e.RowIndex).Value.ToString.ToLower) <> "from master leave" Then
                                If Me.Objcd.CekNumerik(Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                    If Me.DGVConstaint.Item(Me.ColConst_Unit.Index, e.RowIndex).Value.ToString.ToLower = "days" Then
                                        If Me.Objcd.CekRefValueDenganFormatDays(Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                            MsgBox("Please input correct format days." + vbNewLine + "For examples : 0, 0.5, 1, 1.5, 2, 2.5, 3,..... etc.", MsgBoxStyle.Information, "Information")
                                            Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                            Exit Sub
                                        Else
                                            Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = _
                                            Me.Objcd.DigitGrouping(Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value, 1, ClsCekData.EnumTipeRounding.Middlepoint)
                                        End If
                                    ElseIf Me.DGVConstaint.Item(Me.ColConst_Unit.Index, e.RowIndex).Value = "numeric" Then
                                        Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = _
                                        Me.Objcd.DigitGrouping(Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec)
                                    End If
                                End If
                            Else
                                '----Fungsi Pencarian Job Position----
                                If Trim(Me.DGVConstaint.Item(Me.ColConst_Unit.Index, e.RowIndex).Value.ToString.ToLower) = "from master job position" Then
                                    If String.IsNullOrEmpty(Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                        Me.DGVConstaint.Item(Me.ColPrimaryRefValue.Index, e.RowIndex).Value = ""
                                        Me.DGVConstaint.Item(Me.ColConst_RefValue.Index, e.RowIndex).Value = ""
                                        Exit Sub
                                    End If
                                    Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                                    If Me.ObjJP.ApakahSudahAda(Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                        Me.DGVConstaint.Item(Me.ColPrimaryRefValue.Index, e.RowIndex).Value = Me.ObjJP.GetPrimary(Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value)
                                    Else
                                        Dim ij As New Core_Module.FrmMasterJobPosition_Cari
                                        ij.txtCari.Text = Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value
                                        ij.BackColorAll = Me.BackColorAll
                                        ij.ShowDialog()
                                        If ij.mPrimary <> -1 Then
                                            Me.DGVConstaint.Item(Me.ColPrimaryRefValue.Index, e.RowIndex).Value = ij.mPrimary
                                            Me.DGVConstaint.Item(Me.ColConst_RefValue.Index, e.RowIndex).Value = Me.ObjJP.GetCode(ij.mPrimary)
                                        ElseIf ij.mPrimary = -1 Then
                                            Me.DGVConstaint.Item(Me.ColPrimaryRefValue.Index, e.RowIndex).Value = ""
                                            Me.DGVConstaint.Item(Me.ColConst_RefValue.Index, e.RowIndex).Value = ""
                                        End If
                                        ij = Nothing
                                    End If
                                End If
                                '----Fungsi Pencarian Leave Data----
                                If Trim(Me.DGVConstaint.Item(Me.ColConst_Unit.Index, e.RowIndex).Value.ToString.ToLower) = "from master leave" Then
                                    If String.IsNullOrEmpty(Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                        Me.DGVConstaint.Item(Me.ColPrimaryRefValue.Index, e.RowIndex).Value = ""
                                        Me.DGVConstaint.Item(Me.ColConst_RefValue.Index, e.RowIndex).Value = ""
                                        Exit Sub
                                    End If
                                    Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                                    If Me.ObjJP.ApakahSudahAda(Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                        Me.DGVConstaint.Item(Me.ColPrimaryRefValue.Index, e.RowIndex).Value = Me.ObjCuti.GetPrimaryByCode(Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value)
                                    Else
                                        Dim ij As New Core_Module.FrmMasterCuti_Cari
                                        ij.txtCari.Text = Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value
                                        ij.BackColorAll = Me.BackColorAll
                                        ij.ShowDialog()
                                        If ij.PrimaryCuti <> -1 Then
                                            Me.DGVConstaint.Item(Me.ColPrimaryRefValue.Index, e.RowIndex).Value = ij.PrimaryCuti
                                            Me.DGVConstaint.Item(Me.ColConst_RefValue.Index, e.RowIndex).Value = Me.ObjCuti.GetCodeByPrimary(ij.PrimaryCuti)
                                        ElseIf ij.PrimaryCuti = -1 Then
                                            Me.DGVConstaint.Item(Me.ColPrimaryRefValue.Index, e.RowIndex).Value = ""
                                            Me.DGVConstaint.Item(Me.ColConst_RefValue.Index, e.RowIndex).Value = ""
                                        End If
                                        ij = Nothing
                                    End If
                                End If
                                '-------------------------------------
                            End If
                        End If
                    ElseIf Me.ObjSalaryFormula.GetTipeConstrainVariable(Me.DGVConstaint.Item(Me.ColConst_Type2.Index, e.RowIndex).Value) = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi Then
                        Me.CellEndEdit_RefValue_Contraint(e)
                    End If
                End If
            Else
                Me.DGVConstaint.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVConstaint_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DGVConstaint.DataError
        'untuk handle ketika baru buka form dan mau dinomorin
    End Sub

    Private Sub DGVConstaint_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DGVConstaint.RowsAdded
        Try
            If Me.BackgroundWorker3.IsBusy = False Then
                Me.BackgroundWorker3.Dispose()
                Me.BackgroundWorker3.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVConstaint_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DGVConstaint.RowsRemoved
        Try
            If Me.BackgroundWorker3.IsBusy = False Then
                Me.BackgroundWorker3.Dispose()
                Me.BackgroundWorker3.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanData()
        Dim row As DataGridViewRow
        Dim CboDgv As DataGridViewComboBoxCell
        Dim TipeConstraint As Core_Module.ClsMasterSalaryFormula_item.EnumTipeConstraitVariable = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode
        Dim TipeKompetensi As Core_Module.ClsMasterKompetensi.EnumDataType = ClsMasterKompetensi.EnumDataType.Alfabetic

        Try
            'hapus dulu semua item yang bersangkutan
            If Me.mTblVar.Count > 0 Then
                Dim IdxVar As Integer = 0
                Dim count As Integer = 0
                Dim Max As Integer = Me.mTblVar.Count
                While count <> Max
                    If Me.mTblVar.Item(count).Identifier_Salary_Component = Me.mIdentifiyComp Then
                        Me.mTblVar.Rows.RemoveAt(count)
                        count = count - 1
                    End If
                    count += 1
                    If count = Me.mTblVar.Count Then
                        Exit While
                    End If
                End While
            End If

            'simpan item variable baru
            For Each row In Me.DGVVariable.Rows
                If row.Index <> Me.DGVVariable.RowCount - 1 Then
                    If String.IsNullOrEmpty(row.Cells(Me.ColVariable_Variable.Index).Value) = True Then
                        MsgBox("Please input variable", MsgBoxStyle.Information, "Information")
                        Me.DGVVariable.ClearSelection()
                        row.Cells(Me.ColVariable_Variable.Index).Selected = True
                        BerhasilSave = False
                        Exit Sub
                    Else
                        Me.RowVar = Me.mTblVar.NewRow
                        With Me.RowVar
                            .Type2 = Me.ObjSalaryFormula.GetTipeConstrainVariable(row.Cells(Me.ColVariable_Type2.Index).Value)
                            .Primary_Transaksi = Me.mPrimaryTransaksi
                            .Primary_Parameter = row.Cells(Me.ColVariable_PrimaryParameter.Index).Value
                            CboDgv = row.Cells(Me.ColVariable_Type.Index)
                            .Type = CboDgv.Items.IndexOf(row.Cells(Me.ColVariable_Type.Index).Value)
                            .Variable = row.Cells(Me.ColVariable_Variable.Index).Value
                            .Identifier_Salary_Component = Me.mIdentifiyComp
                            .Rounding_Unit = row.Cells(Me.ColRoundingUnit_Variable.Index).Value
                            CboDgv = row.Cells(Me.ColRoundingTipe_Variable.Index)
                            .Rounding_Type = CboDgv.Items.IndexOf(row.Cells(Me.ColRoundingTipe_Variable.Index).Value)
                            .Primary = row.Cells(Me.ColVariable_Primary.Index).Value
                        End With
                        Me.mTblVar.Rows.Add(Me.RowVar)
                    End If
                End If
            Next

            'hapus dulu semua item yang bersangkutan
            If Me.mTblConst.Count > 0 Then
                Dim IdxConts As Integer = 0
                Dim count As Integer = 0
                Dim Max As Integer = Me.mTblConst.Count
                While count <> Max
                    If Me.mTblConst.Item(count).Identifier_Salary_Component = Me.mIdentifiyComp Then
                        Me.mTblConst.Rows.RemoveAt(count)
                        count = count - 1
                    End If
                    count += 1
                    If count = Me.mTblConst.Count Then
                        Exit While
                    End If
                End While
            End If
            'simpan item constraint
            For Each row In Me.DGVConstaint.Rows
                If row.Index <> Me.DGVConstaint.RowCount - 1 Then
                    Me.RowConst = Me.mTblConst.NewRow
                    With Me.RowConst
                        .Primary_Transaksi = Me.mPrimaryTransaksi
                        .Type2 = Me.ObjSalaryFormula.GetTipeConstrainVariable(row.Cells(Me.ColConst_Type2.Index).Value)
                        .Primary_Parameter = row.Cells(Me.ColConst_PrimaryParameter.Index).Value
                        CboDgv = row.Cells(Me.ColConst_Type.Index)
                        .Type = CboDgv.Items.IndexOf(row.Cells(Me.ColConst_Type.Index).Value)
                        CboDgv = row.Cells(Me.ColConst_Compare.Index)
                        .Compare = CboDgv.Items.IndexOf(row.Cells(Me.ColConst_Compare.Index).Value)
                        .Identifier_Salary_Component = Me.mIdentifiyComp
                        .Primary = row.Cells(Me.ColVariable_Primary.Index).Value

                        If .Type2 = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode Then
                            If .Primary_Parameter = 131 Then
                                'jabatan
                                'yang di simpan primary-nya
                                .Ref_Value = row.Cells(Me.ColPrimaryRefValue.Index).Value
                            ElseIf .Primary_Parameter = 133 Then
                                'jenis cuti
                                'yang di simpan primary-nya
                                .Ref_Value = row.Cells(Me.ColPrimaryRefValue.Index).Value
                            Else
                                .Ref_Value = row.Cells(Me.ColConst_RefValue.Index).Value
                            End If
                        ElseIf .Type2 = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi Then
                            TipeKompetensi = Me.ObjKompetensi.GetDataType(row.Cells(Me.ColConst_PrimaryParameter.Index).Value)
                            'default penulisan value dalam bentk string:
                            '1. number      : menggunakan fungsi obj digit group dan menggunakan default prec
                            '2. tanggal     : format MM/dd/yyyy
                            '3. text        : sesuai dengan textnya
                            '4. YN          : false -> 0, true -> 1
                            '5. Alfabetic   : A -> 0, B -> 1, C -> 2, dst

                            If TipeKompetensi = ClsMasterKompetensi.EnumDataType.Alfabetic Then
                                If row.Cells(Me.ColConst_RefValue.Index).Value = "A" Then
                                    .Ref_Value = CInt(ClsMasterKompetensi.Aphabetic.A)
                                ElseIf row.Cells(Me.ColConst_RefValue.Index).Value = "B" Then
                                    .Ref_Value = CInt(ClsMasterKompetensi.Aphabetic.B)
                                ElseIf row.Cells(Me.ColConst_RefValue.Index).Value = "C" Then
                                    .Ref_Value = CInt(ClsMasterKompetensi.Aphabetic.C)
                                ElseIf row.Cells(Me.ColConst_RefValue.Index).Value = "D" Then
                                    .Ref_Value = CInt(ClsMasterKompetensi.Aphabetic.D)
                                ElseIf row.Cells(Me.ColConst_RefValue.Index).Value = "E" Then
                                    .Ref_Value = CInt(ClsMasterKompetensi.Aphabetic.E)
                                ElseIf row.Cells(Me.ColConst_RefValue.Index).Value = "F" Then
                                    .Ref_Value = CInt(ClsMasterKompetensi.Aphabetic.F)
                                End If
                            ElseIf TipeKompetensi = ClsMasterKompetensi.EnumDataType.Number Then
                                .Ref_Value = row.Cells(Me.ColConst_RefValue.Index).Value
                            ElseIf TipeKompetensi = ClsMasterKompetensi.EnumDataType.Tanggal Then
                                .Ref_Value = Format(CDate(row.Cells(Me.ColConst_RefValue.Index).Value), "MM/dd/yyyy")
                            ElseIf TipeKompetensi = ClsMasterKompetensi.EnumDataType.Text Then
                                .Ref_Value = row.Cells(Me.ColConst_RefValue.Index).Value
                            ElseIf TipeKompetensi = ClsMasterKompetensi.EnumDataType.YN Then
                                If row.Cells(Me.ColConst_RefValue.Index).Value = "Yes" Then
                                    .Ref_Value = 1
                                Else
                                    .Ref_Value = 0
                                End If
                            End If
                        End If
                    End With
                    Me.mTblConst.Rows.Add(Me.RowConst)
                End If
            Next

            'Me.ColConst_Primary.Visible = True
            'Me.ColVariable_Primary.Visible = True

            'Dim i As New Form2
            'i.DataGridView1.DataSource = Nothing
            'i.DataGridView1.DataSource = Me.mTblConst
            'i.ShowDialog()


            'Dim ia As New Form2
            'ia.DataGridView1.DataSource = Nothing
            'ia.DataGridView1.DataSource = Me.mTblVar
            'iA.ShowDialog()



            Me.salin = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterSalaryFormula_Component1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Hasil As MsgBoxResult

        Try
            If Me.FromOtherButton = False Then
                Hasil = MsgBox("Accept data changes before closing form?", MsgBoxStyle.YesNoCancel, "Confirm")
                If Hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar = True Then
                        Me.SimpanData()
                        If Me.BerhasilSave = False Then
                            e.Cancel = True
                        End If
                    End If
                ElseIf Hasil = MsgBoxResult.No Then
                    Me.salin = False
                    Me.FromOtherButton = True
                    Me.Close()
                ElseIf Hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            Else
                If Me.CancelClose = True Then
                    Me.CancelClose = False
                    e.Cancel = True
                End If
                Me.FromOtherButton = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterSalaryFormula_Component1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.DGVVariable.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DGVConstaint.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.CancelButton = Me.BtnClose
            Me.ShowIcon = False

            If Me.PunyaHakAksesFull = False Then
                Me.TxtFormula.Enabled = False
                Me.CboRoundingType.Enabled = False
                Me.CboRoundingUnit.Enabled = False
                Me.TxtYearlyMinValue.Enabled = False
                Me.TxtMonthlyMinValue.Enabled = False
                Me.TxtYearlyMaxValue.Enabled = False
                Me.TxtMonthlyMaxValue.Enabled = False
                Me.DtpYear.Enabled = False
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False

                Me.DGVVariable.AllowUserToAddRows = False
                Me.DGVVariable.AllowUserToDeleteRows = False
                Me.DGVConstaint.AllowUserToAddRows = False
                Me.DGVConstaint.AllowUserToDeleteRows = False
                Me.DGVVariable.ReadOnly = True
                Me.DGVConstaint.ReadOnly = True
                Me.btnOK.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Try
            Me.FromOtherButton = False
            If MsgBox("Close form and discard data changes?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                Me.salin = False
                Me.FromOtherButton = True
                Me.Close()
            Else
                Me.FromOtherButton = True
                Me.CancelClose = True
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        'Dim hasil As DialogResult
        'Dim CboDgv As DataGridViewComboBoxCell
        'Dim row As DataGridViewRow
        Dim TipeConstraint As Core_Module.ClsMasterSalaryFormula_item.EnumTipeConstraitVariable = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode
        Dim TipeKompetensi As Core_Module.ClsMasterKompetensi.EnumDataType = ClsMasterKompetensi.EnumDataType.Alfabetic

        Try
            Me.FromOtherButton = False
            If Me.ApakahDataBenar = True Then
                If MsgBox("Close form and accept data changes?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Exit Sub
                End If
                Me.SimpanData()
                Me.FromOtherButton = True
                If Me.BerhasilSave = True Then
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVVariable_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DGVVariable.Sorted
        Dim row As DataGridViewRow
        Try
            For Each row In DGVVariable.Rows
                row.Cells(Me.ColVariable_No.Index).Value = row.Index + 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVVariable_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles DGVVariable.UserDeletingRow
        Try
            'cek
            If String.IsNullOrEmpty(Me.DGVVariable.Item(Me.ColVariable_Variable.Index, e.Row.Index).Value) = False Then
                If Me.TxtFormula.Text.Contains(Me.DGVVariable.Item(Me.ColVariable_Variable.Index, e.Row.Index).Value.ToString) Then
                    e.Cancel = True
                    MsgBox("Data variable cannot deleted, it used to formula salary,", MsgBoxStyle.Information, "Information")
                End If
                If Me.TxtYearlyMaxValue.Text.Contains(Me.DGVVariable.Item(Me.ColVariable_Variable.Index, e.Row.Index).Value.ToString) Then
                    e.Cancel = True
                    MsgBox("Data variable cannot deleted, it used to formula salary,", MsgBoxStyle.Information, "Information")
                End If
                If Me.TxtYearlyMinValue.Text.Contains(Me.DGVVariable.Item(Me.ColVariable_Variable.Index, e.Row.Index).Value.ToString) Then
                    e.Cancel = True
                    MsgBox("Data variable cannot deleted, it used to formula salary,", MsgBoxStyle.Information, "Information")
                End If
                If Me.TxtMonthlyMaxValue.Text.Contains(Me.DGVVariable.Item(Me.ColVariable_Variable.Index, e.Row.Index).Value.ToString) Then
                    e.Cancel = True
                    MsgBox("Data variable cannot deleted, it used to formula salary,", MsgBoxStyle.Information, "Information")
                End If
                If Me.TxtMonthlyMinValue.Text.Contains(Me.DGVVariable.Item(Me.ColVariable_Variable.Index, e.Row.Index).Value.ToString) Then
                    e.Cancel = True
                    MsgBox("Data variable cannot deleted, it used to formula salary,", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtFormula_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFormula.LostFocus
        Me.TxtFormula.BackColor = Color.White
    End Sub

    Private Sub TxtMonthlyMaxValue_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMonthlyMaxValue.LostFocus
        Me.TxtMonthlyMaxValue.BackColor = Color.White
    End Sub

    Private Sub TxtMonthlyMinValue_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMonthlyMinValue.LostFocus
        Me.TxtMonthlyMinValue.BackColor = Color.White
    End Sub

    Private Sub TxtRoundingUnit_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRoundingUnit.LostFocus
        Me.TxtRoundingUnit.BackColor = Color.White
    End Sub

    Private Sub TxtYearlyMaxValue_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtYearlyMaxValue.LostFocus
        Me.TxtYearlyMaxValue.BackColor = Color.White
    End Sub

    Private Sub TxtYearlyMinValue_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtYearlyMinValue.LostFocus
        Me.TxtYearlyMinValue.BackColor = Color.White
    End Sub

    Private Sub DGVConstaint_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DGVConstaint.Sorted
        Dim row As DataGridViewRow
        Try
            For Each row In DGVConstaint.Rows
                row.Cells(Me.ColConts_No.Index).Value = row.Index + 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboRoundingType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboRoundingType.SelectedIndexChanged
        Try
            If Me.CboRoundingType.SelectedIndex = 0 Then
                Me.CboRoundingUnit.Text = 1
                Me.CboRoundingUnit.Enabled = False
            Else
                Me.CboRoundingUnit.Enabled = True
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Function GetListVariable(ByVal vFormula As String) As String
        Dim row As DataGridViewRow
        Dim cek As Boolean = True
        Dim var As String = ""
        Dim ketemu As Boolean
        Dim ListStr As String = ""
        Try
            For i As Integer = 0 To vFormula.Length - 1
                cek = False
                'cek dulu variablenya , BUKAN KARAKTERNYA
                If Me.ObjCd.CekHuruf(vFormula.Chars(i)) = True Then
                    var += vFormula.Chars(i)
                End If
                'pastiin karakter berikutnya masih ada char atau tidak
                'kalo ada tambahin dlu , kalo tidak langsung cek
                'kalo karakter terakhir kosong langsung cek
                If i <> vFormula.Length - 1 Then
                    If Me.ObjCd.CekHuruf(vFormula.Chars(i + 1)) = False Then
                        cek = True
                    End If
                Else
                    cek = True
                End If
                If cek = True Then
                    If Me.ObjCd.CekHuruf(var) = True And var <> "" Then
                        ketemu = False
                        For Each row In DGVVariable.Rows
                            If row.Cells(Me.ColVariable_Variable.Index).Value = var Then
                                ketemu = True
                                Exit For
                            End If
                        Next
                        If ketemu = False And var.ToLower <> "if" Then
                            If ListStr = "" Then
                                ListStr = var
                            Else
                                ListStr += ", " & var
                            End If

                        End If
                        var = ""
                    End If
                End If
            Next
            Return ListStr
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ListStr
        End Try
    End Function
End Class