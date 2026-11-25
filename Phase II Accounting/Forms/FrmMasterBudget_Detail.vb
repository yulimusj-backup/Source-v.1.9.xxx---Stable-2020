Option Explicit On
Option Strict Off

Public Class FrmMasterBudget_Detail
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public Ref_Row As DataGridViewRow
    Public Ref_Form As FrmMasterBudget

    Dim ObjCD As New Core_Module.ClsCekData
    Dim mbuffer As String = ""
    Dim HandleFormClosing As Boolean = True

    Dim Def_Prec As Integer = 0
    Dim ObjCur As New Core_Module.clsMasterCurrency

    Private Sub FrmMasterBudget_Detail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Def_Prec = Me.ObjCur.GetDecimal_DefaultCurrency

            With Ref_Row
                Me.dgvItem.Rows.Add("January", "-", "-", .Cells(Me.Ref_Form.ColJan.Index).Value)
                Me.dgvItem.Rows.Add("February", 0, 0, .Cells(Me.Ref_Form.ColFeb.Index).Value)
                Me.dgvItem.Rows.Add("March", 0, 0, .Cells(Me.Ref_Form.ColMar.Index).Value)
                Me.dgvItem.Rows.Add("April", 0, 0, .Cells(Me.Ref_Form.ColApr.Index).Value)
                Me.dgvItem.Rows.Add("May", 0, 0, .Cells(Me.Ref_Form.ColMay.Index).Value)
                Me.dgvItem.Rows.Add("June", 0, 0, .Cells(Me.Ref_Form.ColJune.Index).Value)
                Me.dgvItem.Rows.Add("July", 0, 0, .Cells(Me.Ref_Form.ColJul.Index).Value)
                Me.dgvItem.Rows.Add("August", 0, 0, .Cells(Me.Ref_Form.ColAgt.Index).Value)
                Me.dgvItem.Rows.Add("September", 0, 0, .Cells(Me.Ref_Form.ColSept.Index).Value)
                Me.dgvItem.Rows.Add("October", 0, 0, .Cells(Me.Ref_Form.ColOct.Index).Value)
                Me.dgvItem.Rows.Add("November", 0, 0, .Cells(Me.Ref_Form.ColNov.Index).Value)
                Me.dgvItem.Rows.Add("December", 0, 0, .Cells(Me.Ref_Form.ColDec.Index).Value)
            End With
            Me.Hitung_Difference()
            Me.cboMonth.SelectedIndex = 0
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.TxtBeginingBalance.Text = Me.ObjCD.DigitGrouping(Me.TxtBeginingBalance.Text, Me.Def_Prec)
            Me.TxtPencentage.Text = Me.ObjCD.DigitGrouping(Me.TxtPencentage.Text, 2)
            Me.TxtNominal.Text = Me.ObjCD.DigitGrouping(Me.TxtNominal.Text, Me.Def_Prec)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' Hitung difference untuk setiap barisnnya
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Hitung_Difference()
        Dim Cur_Row, Base_Row As DataGridViewRow
        Dim Diff As Double = 0
        Try
            For Each Cur_Row In Me.dgvItem.Rows
                If Cur_Row.Index <> 0 Then
                    Base_Row = Me.dgvItem.Rows(Cur_Row.Index - 1)
                    Diff = CDec(Cur_Row.Cells(Me.ColSaldo.Index).Value) - CDec(Base_Row.Cells(Me.ColSaldo.Index).Value)
                    Cur_Row.Cells(Me.ColDifference.Index).Value = Me.ObjCD.DigitGrouping(Diff)
                    If Diff = 0 Then
                        Cur_Row.Cells(Me.ColDifference_Percentage.Index).Value = "-"
                        Cur_Row.Cells(Me.ColDifference.Index).Value = "-"
                    Else
                        If CDec(Base_Row.Cells(Me.ColSaldo.Index).Value) = 0 Then
                            Cur_Row.Cells(Me.ColDifference_Percentage.Index).Value = "-"
                        Else
                            Cur_Row.Cells(Me.ColDifference_Percentage.Index).Value = _
                                                   Me.ObjCD.DigitGrouping(Math.Round((Diff * 100) / CDec(Base_Row.Cells(Me.ColSaldo.Index).Value), 2, _
                                                   MidpointRounding.AwayFromZero), 2)
                        End If

                        Cur_Row.Cells(Me.ColDifference.Index).Value = Me.ObjCD.DigitGrouping(Diff, Me.Def_Prec)
                    End If
                Else
                    Cur_Row.Cells(Me.ColDifference_Percentage.Index).Value = "-"
                    Cur_Row.Cells(Me.ColDifference.Index).Value = "-"
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'yang bisa diedit adalah kolom saldo saja
            If e.ColumnIndex <> Me.ColSaldo.Index Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            'pastikan ada isinya
            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                'cek panjang maksimum untuk string dan nilai maksimum untuk angka
                'untuk kolom budget cek data angka
                If e.ColumnIndex = Me.ColSaldo.Index Then
                    'cek bilangan non negatif
                    If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Data must be a non negative number.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.mbuffer
                        Exit Sub
                    Else
                        If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value > 999999999999999 Then
                            MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(999999999999999), _
                            MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.mbuffer
                            Exit Sub
                        End If
                    End If
                End If

                'sampe sini data angka ok, format
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, _
                e.RowIndex).Value, Me.Def_Prec)
                'update difference
                Me.Hitung_Difference()
            Else
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.mbuffer
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterBudget_Detail_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Hasil As MsgBoxResult

        Try
            If Me.HandleFormClosing = True Then
                Hasil = MsgBox("Accept data change before closing?", MsgBoxStyle.YesNoCancel, "Confirm")
                If Hasil = MsgBoxResult.Yes Then
                    'terima perubahan data
                    With Me.Ref_Row
                        .Cells(Me.Ref_Form.ColJan.Index).Value = Me.dgvItem.Rows(0).Cells(Me.ColSaldo.Index).Value
                        .Cells(Me.Ref_Form.ColFeb.Index).Value = Me.dgvItem.Rows(1).Cells(Me.ColSaldo.Index).Value
                        .Cells(Me.Ref_Form.ColMar.Index).Value = Me.dgvItem.Rows(2).Cells(Me.ColSaldo.Index).Value
                        .Cells(Me.Ref_Form.ColApr.Index).Value = Me.dgvItem.Rows(3).Cells(Me.ColSaldo.Index).Value
                        .Cells(Me.Ref_Form.ColMay.Index).Value = Me.dgvItem.Rows(4).Cells(Me.ColSaldo.Index).Value
                        .Cells(Me.Ref_Form.ColJune.Index).Value = Me.dgvItem.Rows(5).Cells(Me.ColSaldo.Index).Value
                        .Cells(Me.Ref_Form.ColJul.Index).Value = Me.dgvItem.Rows(6).Cells(Me.ColSaldo.Index).Value
                        .Cells(Me.Ref_Form.ColAgt.Index).Value = Me.dgvItem.Rows(7).Cells(Me.ColSaldo.Index).Value
                        .Cells(Me.Ref_Form.ColSept.Index).Value = Me.dgvItem.Rows(8).Cells(Me.ColSaldo.Index).Value
                        .Cells(Me.Ref_Form.ColOct.Index).Value = Me.dgvItem.Rows(9).Cells(Me.ColSaldo.Index).Value
                        .Cells(Me.Ref_Form.ColNov.Index).Value = Me.dgvItem.Rows(10).Cells(Me.ColSaldo.Index).Value
                        .Cells(Me.Ref_Form.ColDec.Index).Value = Me.dgvItem.Rows(11).Cells(Me.ColSaldo.Index).Value
                    End With
                    Me.Ref_Form.UpdateSaldoHeader()
                ElseIf Hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If MsgBox("Accept data changes?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                'terima perubahan data
                With Me.Ref_Row
                    .Cells(Me.Ref_Form.ColJan.Index).Value = Me.dgvItem.Rows(0).Cells(Me.ColSaldo.Index).Value
                    .Cells(Me.Ref_Form.ColFeb.Index).Value = Me.dgvItem.Rows(1).Cells(Me.ColSaldo.Index).Value
                    .Cells(Me.Ref_Form.ColMar.Index).Value = Me.dgvItem.Rows(2).Cells(Me.ColSaldo.Index).Value
                    .Cells(Me.Ref_Form.ColApr.Index).Value = Me.dgvItem.Rows(3).Cells(Me.ColSaldo.Index).Value
                    .Cells(Me.Ref_Form.ColMay.Index).Value = Me.dgvItem.Rows(4).Cells(Me.ColSaldo.Index).Value
                    .Cells(Me.Ref_Form.ColJune.Index).Value = Me.dgvItem.Rows(5).Cells(Me.ColSaldo.Index).Value
                    .Cells(Me.Ref_Form.ColJul.Index).Value = Me.dgvItem.Rows(6).Cells(Me.ColSaldo.Index).Value
                    .Cells(Me.Ref_Form.ColAgt.Index).Value = Me.dgvItem.Rows(7).Cells(Me.ColSaldo.Index).Value
                    .Cells(Me.Ref_Form.ColSept.Index).Value = Me.dgvItem.Rows(8).Cells(Me.ColSaldo.Index).Value
                    .Cells(Me.Ref_Form.ColOct.Index).Value = Me.dgvItem.Rows(9).Cells(Me.ColSaldo.Index).Value
                    .Cells(Me.Ref_Form.ColNov.Index).Value = Me.dgvItem.Rows(10).Cells(Me.ColSaldo.Index).Value
                    .Cells(Me.Ref_Form.ColDec.Index).Value = Me.dgvItem.Rows(11).Cells(Me.ColSaldo.Index).Value
                End With
                Me.Ref_Form.UpdateSaldoHeader()
                Me.HandleFormClosing = False
                'tutup form
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If MsgBox("Close form and discard data changes?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.HandleFormClosing = False
                'tutup form
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rbo_Percentage_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbo_Percentage.CheckedChanged
        Try
            If Me.rbo_Percentage.Checked = True Then
                Me.TxtNominal.Enabled = False
                Me.TxtPencentage.Enabled = True
                Me.TxtNominal.Text = Me.ObjCD.DigitGrouping("0", Me.Def_Prec)
            Else
                Me.TxtNominal.Enabled = True
                Me.TxtPencentage.Enabled = False
                Me.TxtPencentage.Text = "0.00"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnModify_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnModify.Click
        Try
            If Me.rbo_Percentage.Checked = True Then
                Me.Modify_Percentage(Me.cboMonth.SelectedIndex)
            Else
                Me.Modify_Nominal(Me.cboMonth.SelectedIndex)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Modify_Nominal(ByVal vStartingRowIndex As Integer)
        Dim rowdgv As DataGridViewRow
        Dim Base_Row As DataGridViewRow
        Dim diff As Decimal = 0

        Try
            For i As Integer = vStartingRowIndex To Me.dgvItem.Rows.Count - 1
                rowdgv = Me.dgvItem.Rows(i)
                If rowdgv.Index = vStartingRowIndex Then
                    rowdgv.Cells(Me.ColSaldo.Index).Value = Me.TxtBeginingBalance.Text
                Else
                    Base_Row = Me.dgvItem.Rows(i - 1)
                    diff = CDec(Me.TxtNominal.Text)
                    rowdgv.Cells(Me.ColSaldo.Index).Value = Me.ObjCD.DigitGrouping(CDec(Base_Row.Cells(Me.ColSaldo.Index).Value) + diff, Me.Def_Prec)
                End If
            Next

            Me.Hitung_Difference()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Modify_Percentage(ByVal vStartingRowIndex As Integer)
        Dim rowdgv As DataGridViewRow
        Dim Base_Row As DataGridViewRow
        Dim diff As Decimal = 0

        Try
            For i As Integer = vStartingRowIndex To Me.dgvItem.Rows.Count - 1
                rowdgv = Me.dgvItem.Rows(i)
                If rowdgv.Index = vStartingRowIndex Then
                    rowdgv.Cells(Me.ColSaldo.Index).Value = Me.TxtBeginingBalance.Text
                Else
                    Base_Row = Me.dgvItem.Rows(i - 1)
                    diff = Math.Round((CDec(Base_Row.Cells(Me.ColSaldo.Index).Value) * _
                    CDec(Me.TxtPencentage.Text)) / 100, Me.Def_Prec, MidpointRounding.AwayFromZero)
                    rowdgv.Cells(Me.ColSaldo.Index).Value = Me.ObjCD.DigitGrouping(CDec(Base_Row.Cells(Me.ColSaldo.Index).Value) + diff, Me.Def_Prec)
                End If
            Next

            Me.Hitung_Difference()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBeginingBalance_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBeginingBalance.Click
        Try
            Me.ObjCD.Number_Click(Me.TxtBeginingBalance)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBeginingBalance_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBeginingBalance.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtBeginingBalance, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBeginingBalance_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtBeginingBalance.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtBeginingBalance, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBeginingBalance_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtBeginingBalance.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtBeginingBalance, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBeginingBalance_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBeginingBalance.GotFocus
        Try
            Me.TxtBeginingBalance.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBeginingBalance_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBeginingBalance.LostFocus
        Try
            Me.TxtBeginingBalance.BackColor = Me.TxtPrimary.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPencentage_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPencentage.Click
        Try
            Me.ObjCD.Number_Click(Me.TxtPencentage)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPencentage_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPencentage.Enter
        Try
            If Me.TxtPencentage.Enabled = True Then
                Me.ObjCD.Number_Enter(Me.TxtPencentage, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPencentage_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPencentage.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtPencentage, sender, e, 100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPencentage_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPencentage.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtPencentage, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPencentage_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPencentage.GotFocus
        Try
            Me.TxtPencentage.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPencentage_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPencentage.LostFocus
        Try
            Me.TxtPencentage.BackColor = Me.TxtPrimary.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNominal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNominal.Click
        Try
            Me.ObjCD.Number_Click(Me.TxtNominal)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNominal_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNominal.Enter
        Try
            If Me.TxtNominal.Enabled = True Then
                Me.ObjCD.Number_Enter(Me.TxtNominal, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNominal_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtNominal.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtNominal, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNominal_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtNominal.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtNominal, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNominal_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNominal.GotFocus
        Try
            Me.TxtNominal.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNominal_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNominal.LostFocus
        Try
            Me.TxtNominal.BackColor = Me.TxtPrimary.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBeginingBalance_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBeginingBalance.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtBeginingBalance, sender, e, Me.Def_Prec, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPencentage_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPencentage.Leave
        Try
            If Me.TxtPencentage.Enabled = True Then
                Me.ObjCD.Number_Leave(Me.TxtPencentage, sender, e, 2, Me.ObjCD)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNominal_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNominal.Leave
        Try
            If Me.TxtNominal.Enabled = True Then
                Me.ObjCD.Number_Leave(Me.TxtNominal, sender, e, Me.Def_Prec, Me.ObjCD)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class