Option Explicit On
Option Strict Off

Imports System.Windows.Forms, System.Drawing, Core_Module
Public Class FrmTrans_Gaji_OtherDeduction

    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public ijGaji As FrmTrans_Gaji
    Public ijPay As FrmTrans_Gaji_Payment
    Public DipanggilKetika_PosisiNonAktif As Boolean = False
    Public NonAktif As Boolean = False
    Dim Buffer As String = ""
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjKrywn As New Core_Module.ClsMasterKaryawan
    Dim Obj As New ClsTransaksiSalary

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub LoadData()
        Try
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            If Me.NonAktif = False Then
                Me.dgvItem.ReadOnly = True
            Else
                Me.dgvItem.ReadOnly = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Gaji_OtherDeduction_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.ShowIcon = False
            Me.LoadData()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If Me.DipanggilKetika_PosisiNonAktif = True Then
                If e.ColumnIndex = Me.ColExplain.Index Then
                Else
                    e.Cancel = True
                End If
            Else
                If Me.dgvItem.Item(Me.ColOverride.Index, e.RowIndex).Value = True Then
                    If Me.dgvItem.Item(Me.ColOriginalValue.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColAmount.Index, e.RowIndex).Value Then
                        If e.ColumnIndex = Me.ColAmount.Index Or e.ColumnIndex = Me.ColRemarks.Index Then
                            Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        Else
                            e.Cancel = True
                        End If
                    ElseIf Me.dgvItem.Item(Me.ColOriginalValue.Index, e.RowIndex).Value <> Me.dgvItem.Item(Me.ColAmount.Index, e.RowIndex).Value Then
                        If e.ColumnIndex = Me.ColAmount.Index Or _
                        e.ColumnIndex = Me.ColOverrideReason.Index Or _
                        e.ColumnIndex = Me.ColRemarks.Index Then
                            Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        Else
                            e.Cancel = True
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColOverride.Index Or e.ColumnIndex = Me.ColRemarks.Index Then
                        Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellClick
        Dim FrmExplain As New FrmTrans_Gaji_ShowExplaination
        Dim RowDGV As DataGridViewRow
        Try
            If ijGaji.TxtEmployeeCode1.Text = "-1" Then
                Exit Sub
            End If
            If Me.dgvItem.Rows.Count <= 0 Then
                Exit Sub
            End If
            If e.RowIndex < 0 Then
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColExplain.Index Then
                FrmExplain.TxtFormula.Text = Me.dgvItem.Item(Me.ColExplaination.Index, e.RowIndex).Value
                If Me.dgvItem.Item(Me.ColOverride.Index, e.RowIndex).Value = True Then
                    FrmExplain.TxtFormula.Text += vbCrLf + "Override Value = " + Me.dgvItem.Item(Me.ColAmount.Index, e.RowIndex).Value.ToString
                    FrmExplain.TxtFormula.Text += vbCrLf + "Reason = " + Me.dgvItem.Item(Me.ColOverrideReason.Index, e.RowIndex).Value.ToString
                End If
                FrmExplain.ShowDialog()
            End If
            If Me.NonAktif = True Then
                If e.ColumnIndex = Me.ColOverride.Index Then
                    Me.Buffer = Me.dgvItem.Item(Me.ColOverride.Index, e.RowIndex).Value
                    Me.dgvItem.Item(Me.ColAmount.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColOriginalValue.Index, e.RowIndex).Value
                    Me.dgvItem.Item(Me.ColOverrideReason.Index, e.RowIndex).Value = ""
                    Me.dgvItem.Item(Me.ColOverride.Index, e.RowIndex).Value = False
                End If
                'Checklist kalo false 
                If e.ColumnIndex = Me.ColOverride.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = False Then
                        '---------------------
                        'Hitung Total Diheader
                        ijGaji.TxtOtherDedution.Text = "0"
                        ijGaji.TxtSalaryReceived.Text = "0"
                        If Me.dgvItem.Rows.Count > 0 Then
                            For Each RowDGV In Me.dgvItem.Rows
                                ijGaji.TxtOtherDedution.Text += CDec(RowDGV.Cells(Me.ColAmount.Index).Value)
                            Next
                        End If
                        ijGaji.TxtOtherDedution.Text = Me.ObjCD.DigitGrouping(ijGaji.TxtOtherDedution.Text, ijGaji.Def_Prec)
                        'Salary Received
                        ijGaji.TxtSalaryReceived.Text = Me.ObjCD.DigitGrouping(ijGaji.HitungTakeHomePay(CDec(ijGaji.TxtIncome.Text), CDec(ijGaji.TxtDeduction.Text), _
                        CDec(ijGaji.TxtIncomeTax.Text), CDec(ijGaji.TxtLoanInstallment.Text), CDec(ijGaji.TxtOtherDedution.Text), CDec(ijGaji.TxtKurangBayarTax.Text), _
                        CDec(ijGaji.TxtEmployeeExpense.Text)), ijGaji.Def_Prec)
                        '-------------
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmExplain.Dispose()
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim Row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
        Dim RowDGV As DataGridViewRow
        Try
            'pastikan ada isinya
            With Row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColAmount.Index Then
                        If Me.ObjCD.CekNonNegatif(Row.Cells(Me.ColAmount.Index).Value) = False Then
                            MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        Else
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, ijGaji.Def_Prec)
                            ijGaji.TxtOtherDedution.Text -= CDec(Me.Buffer)
                            ijGaji.TxtOtherDedution.Text += CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                            ijGaji.TxtOtherDedution.Text = Me.ObjCD.DigitGrouping(ijGaji.TxtOtherDedution.Text, ijGaji.Def_Prec)
                            'Salary Received
                            ijGaji.TxtSalaryReceived.Text = Me.ObjCD.DigitGrouping(ijGaji.HitungTakeHomePay(CDec(ijGaji.TxtIncome.Text), CDec(ijGaji.TxtDeduction.Text), _
                            CDec(ijGaji.TxtIncomeTax.Text), CDec(ijGaji.TxtLoanInstallment.Text), CDec(ijGaji.TxtOtherDedution.Text), CDec(ijGaji.TxtKurangBayarTax.Text), _
                            CDec(ijGaji.TxtEmployeeExpense.Text)), ijGaji.Def_Prec)
                            If CDec(Row.Cells(Me.ColAmount.Index).Value) = CDec(Row.Cells(Me.ColOriginalValue.Index).Value) Then
                                Row.Cells(Me.ColOverride.Index).Value = False
                                Row.Cells(Me.ColOverrideReason.Index).Value = ""
                            End If
                            'Update Payment
                            If Me.ijPay.DGVPayment.Rows.Count = 1 Then
                                RowDGV = Me.ijPay.DGVPayment.Rows(0)
                                RowDGV.Cells(Me.ijPay.ColDetailPayment_Amount.Index).Value = Me.ObjCD.DigitGrouping(Me.ijGaji.TxtSalaryReceived.Text, Me.ijGaji.Def_Prec)
                                Me.ijGaji.TxtPayment.Text = Me.ObjCD.DigitGrouping(Me.ijGaji.TxtSalaryReceived.Text, Me.ijGaji.Def_Prec)
                            End If
                            'TxtPayment.Text
                            Me.ijGaji.TxtPayment.Text = Me.ijGaji.TxtSalaryReceived.Text
                            '----------
                        End If
                    End If
                    If e.ColumnIndex = Me.ColOverrideReason.Index Then
                        If Me.ObjCD.CekDataKosong(Row.Cells(Me.ColOverrideReason.Index).Value, False) = False Then
                            MsgBox("Data shouldn't include the character: ' , % , or * .", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        If Me.ObjCD.CekDataKosong(Row.Cells(Me.ColRemarks.Index).Value, False) = False Then
                            MsgBox("Data shouldn't include the character: ' , % , or * .", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

End Class