Option Explicit On
Option Strict Off

Public Class FrmTransaksiAOP_AdjustDPP2
    Dim ObjCurrency As New Core_Module.clsMasterCurrency
    Dim ObjCD As New Core_Module.ClsCekData
    Dim LangsungClose As Boolean = False
    Dim Buffer As String = ""

    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = ""
    Public IsReadOnly As Boolean = True
    Public vFromQuick As Boolean = True

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If MsgBox("Accept data changes?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.LangsungClose = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If Me.IsReadOnly = False Then
                If MsgBox("Cancel data changes?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                    Me.DialogResult = Windows.Forms.DialogResult.Cancel
                    Me.LangsungClose = True
                    Me.Close()
                End If
            Else
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiAOP_AdjustDPP_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim ij As MsgBoxResult

        Try
            If Me.LangsungClose = False Then
                ij = MsgBox("Accept data changes before closing?", MsgBoxStyle.YesNoCancel, "Confirmation")
                If ij = MsgBoxResult.Yes Then
                    Me.DialogResult = Windows.Forms.DialogResult.OK
                    Me.LangsungClose = True
                    Me.Close()
                ElseIf ij = MsgBoxResult.No Then
                    Me.DialogResult = Windows.Forms.DialogResult.Cancel
                    Me.LangsungClose = True
                    Me.Close()
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'yang boleh di edit hanya kolom sub total dan reason
            If Me.IsReadOnly = True Then
                e.Cancel = True
            Else
                If e.ColumnIndex = Me.ColItemAdjustment.Index Or e.ColumnIndex = Me.ColItemAdjustmentReason.Index Then
                    Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim DefaultPrecisionRow As Integer = 0
        Dim row As DataGridViewRow
        Dim SubTotalAwal As Decimal = 0
        Dim UnitPrice As Decimal = 0
        Dim NettPrice As Decimal = 0
        Dim DiscAmount As Decimal = 0
        Dim DicsPerct As Decimal = 0
        Dim Amount As Decimal = 0

        Try
            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                If e.ColumnIndex = Me.ColItemAdjustmentReason.Index Then
                    'jika remarks, masukkan string kosong
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                ElseIf e.ColumnIndex = Me.ColItemAdjustment.Index Then
                    'jika adjustment, maka masukkan buffernya
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                End If
            Else
                If e.ColumnIndex = Me.ColItemAdjustmentReason.Index Then
                    'jika remarks, pastikan lengthnya tidak lebih dr 250
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If
                ElseIf e.ColumnIndex = Me.ColItemAdjustment.Index Then
                    row = Me.dgvItem.Rows(e.RowIndex)
                    DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)

                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 15 Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    'pastikan nilainya angka
                    If IsNumeric(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    Else
                        'pastikan nilai angka tidak lebih dari 9 M
                        If Math.Abs(CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)) > 9999999999 Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                    'kalo ga ada perubahan ya tidak usah di ubah apa2
                    If CDec(Me.Buffer) = CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    'digit grouping
                    row.Cells(Me.ColItemAdjustment.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColItemAdjustment.Index).Value, DefaultPrecisionRow)

                    'pastikan nilai adjustmentnya tidak lebih dari nilai sub total asli
                    SubTotalAwal = CDec(row.Cells(Me.ColAmount.Index).Value) - CDec(Me.Buffer)
                    If CDec(row.Cells(Me.ColItemAdjustment.Index).Value) + SubTotalAwal <= 0 Then
                        MsgBox("Adjustment value must be smaller than " + Me.ObjCD.DigitGrouping(SubTotalAwal, DefaultPrecisionRow), MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    'nett price tidak boleh nol
                    Amount = SubTotalAwal + CDec(row.Cells(Me.ColItemAdjustment.Index).Value)
                    'NettPrice = Math.Round(Amount / CDec(row.Cells(Me.ColQtyMulti_Item.Index).Value), DefaultPrecisionRow, MidpointRounding.AwayFromZero)
                    'If NettPrice <= 0 Then
                    '    MsgBox("Invalid nett price.", MsgBoxStyle.Information, "Information")
                    '    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    '    Exit Sub
                    'End If
                    ''hitung unit price + disc amount + disc %
                    'If row.Cells(Me.ColDiscFromPerct.Index).Value = "1" Then
                    '    UnitPrice = Me.ObjCD.DigitGrouping((NettPrice * 100) / (100 - CDec(row.Cells(Me.ColDiscount.Index).Value)), DefaultPrecisionRow)
                    '    DiscAmount = UnitPrice - NettPrice
                    '    DicsPerct = row.Cells(Me.ColDiscount.Index).Value
                    'Else
                    '    DiscAmount = row.Cells(Me.ColDiscAmount.Index).Value
                    '    UnitPrice = NettPrice - DiscAmount
                    '    DicsPerct = Math.Round(DiscAmount / UnitPrice, 2, MidpointRounding.AwayFromZero)
                    'End If
                    ''unit price tidak boleh nol
                    'If UnitPrice <= 0 Then
                    '    MsgBox("Invalid unit price.", MsgBoxStyle.Information, "Information")
                    '    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    '    Exit Sub
                    'End If

                    ''update ke tampilan dgv
                    'row.Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(UnitPrice, DefaultPrecisionRow)
                    'row.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(NettPrice, DefaultPrecisionRow)
                    row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Amount, DefaultPrecisionRow)
                    'If row.Cells(Me.ColDiscFromPerct.Index).Value = "1" Then
                    '    'unit price, disc amount, nett price, amount
                    '    row.Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(DiscAmount, DefaultPrecisionRow)
                    'Else
                    '    'unit price, disc perct, nett price, amount
                    '    row.Cells(Me.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping(DicsPerct, 2)
                    'End If

                    row.Cells(Me.ColItemAdjustmentUser.Index).Value = Me.UserCode
                    row.Cells(Me.ColItemAdjustmentTimestamp.Index).Value = Format(Now, "dd MMM yyyy HH:mm")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiAOP_AdjustDPP_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            If Me.IsReadOnly = True Then
                Me.btnOK.Visible = False
                Me.btnOK.TabStop = True
                Me.btnCancel.Text = "Close"
                Me.LangsungClose = True
                Me.dgvItem.ReadOnly = True
            Else
                Me.btnOK.Visible = True
                Me.btnOK.TabStop = True
                Me.btnCancel.Text = "Cancel"
                Me.dgvItem.ReadOnly = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class