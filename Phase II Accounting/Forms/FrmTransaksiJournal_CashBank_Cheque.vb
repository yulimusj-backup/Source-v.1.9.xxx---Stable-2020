Option Explicit On
Option Strict Off

Public Class FrmTransaksiJournal_CashBank_Cheque
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjCd As New Core_Module.ClsCekData
    Dim ObjJurnal As New ClsTransaksiJournal
    Dim ObjCheque As New ClsTransaksiJournal_Item_Cheque
    Dim Buffer As String = ""

    Public mPrimaryJurnal As Double = -1
    Public mBaru As Boolean = False
    Public mPrimaryCurr As Double = -1
    Public mTypeInOut As ClsTransaksiJournal.EnumTipeInOut

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click

        Try
            'yang harus diisikan:
            '1. yang harus diisikan: cheque no, bank, cheque date
            '2. cheque amount harus positive
            If Me.ApakahDatabenar = True Then
                Me.Close()
                'Me.ObjCheque.HapusItem(Me.mPrimaryJurnal)
                'Me.ObjCheque.SimpanCheque(Me.mPrimaryJurnal, Me.dgvCheque, Me)
                'MsgBox("Data Saved", MsgBoxStyle.Information)
            End If
            Me.dgvCheque.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCheque_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvCheque.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.dgvCheque(Me.ColChequeNo.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColBank.Index Or e.ColumnIndex = Me.ColChequeNo.Index Or _
                e.ColumnIndex = Me.ColDueDate.Index Or e.ColumnIndex = Me.ColTerms.Index Or _
                e.ColumnIndex = Me.ColAmount.Index Or e.ColumnIndex = Me.ColRemarks.Index Or _
                e.ColumnIndex = Me.ColSPaid.Index Or e.ColumnIndex = Me.ColDate.Index Then
                    Me.Buffer = Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColChequeNo.Index Then
                    Me.Buffer = Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCheque_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCheque.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvCheque.Rows(e.RowIndex)
        Try
            With row
                If e.ColumnIndex = Me.ColChequeNo.Index Then
                    If Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        MsgBox("Cheque no can not be empty", MsgBoxStyle.Information, "Information")
                        Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        If Me.Buffer = "" Then
                            Me.dgvCheque.Rows.RemoveAt(e.RowIndex)
                        End If
                        Exit Sub
                    End If
                    If row.Cells(Me.ColIdentifier.Index).Value = "" Then
                        Dim Identifier As String = ""
                        Identifier = ObjCheque.GetIdentifierUrutan _
                        (Me.dgvCheque, Me.ColIdentifier.Index, Me.mBaru, Me.mPrimaryJurnal, row.Index)
                        row.Cells(Me.ColIdentifier.Index).Value = Identifier
                    End If
                End If
                If e.ColumnIndex = Me.ColBank.Index Then
                    If Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        MsgBox("Bank can not be empty", MsgBoxStyle.Information, "Information")
                        Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    Else
                        Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpperInvariant
                    End If
                End If
                If e.ColumnIndex = Me.ColDate.Index Then
                    If Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value = "" Or _
                    IsDate(Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please enter a date", MsgBoxStyle.Information, "Information")
                        Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    ElseIf CDate(Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value) > CDate(row.Cells(Me.ColDueDate.Index).Value) Then
                        MsgBox("Cheque date must not be greater than due date", MsgBoxStyle.Information, "Information")
                        Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    Else
                        Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")
                        row.Cells(Me.ColTerms.Index).Value = (CDate(row.Cells(Me.ColDueDate.Index).Value) - CDate(Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value)).Days
                    End If
                End If
                If e.ColumnIndex = Me.ColDueDate.Index Then
                    If Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value = "" Or _
                    IsDate(Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please enter a date", MsgBoxStyle.Information, "Information")
                        Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    ElseIf CDate(row.Cells(Me.ColDate.Index).Value) > CDate(Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value) Then
                        MsgBox("Due date should not be smaller than cheque date", MsgBoxStyle.Information, "Information")
                        Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    Else
                        Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")
                        row.Cells(Me.ColTerms.Index).Value = (CDate(Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value) - CDate(row.Cells(Me.ColDate.Index).Value)).Days
                    End If
                End If
                If e.ColumnIndex = Me.ColTerms.Index Then
                    If Me.ObjCd.CekNonNegatif(Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = False Then
                        MsgBox("Please enter a positive number", MsgBoxStyle.Information, "Information")
                        Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    Else
                        row.Cells(Me.ColDueDate.Index).Value = CDate(row.Cells(Me.ColDate.Index).Value).AddDays(Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value)
                        row.Cells(Me.ColDueDate.Index).Value = Format(row.Cells(Me.ColDueDate.Index).Value, "dd MMM yyyy")
                    End If
                End If
                If e.ColumnIndex = Me.ColAmount.Index Then
                    If Me.ObjCd.CekNonNegatif(Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value) = False Or _
                   Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value = "0" Then
                        MsgBox("Please enter a positive number and not 0 for amount ", MsgBoxStyle.Information, "Information")
                        Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    Else
                        Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCd.DigitGrouping(Me.dgvCheque.Item(e.ColumnIndex, e.RowIndex).Value.ToString, _
                        Me.ObjCur.GetDefaultPrecission(Me.mPrimaryCurr))
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCheque_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvCheque.RowsAdded
        'default:
        '1. date = due date = today
        '2. amount =0
        '3. cheque no, bank, remarks
        '4. paid = false
        Dim row As DataGridViewRow
        Try
            row = Me.dgvCheque.Rows(Me.dgvCheque.RowCount - 1)
            row.Cells(Me.ColCurrency.Index).Value = Me.ObjCur.GetCurrencyCode(Me.mPrimaryCurr)
            row.Cells(Me.ColPrimaryCurrency.Index).Value = Me.mPrimaryCurr
            row.Cells(Me.ColDate.Index).Value = Format(Today, "dd MMM yyyy")
            row.Cells(Me.ColDueDate.Index).Value = Format(Today, "dd MMM yyyy")
            row.Cells(Me.ColTerms.Index).Value = 0
            row.Cells(Me.ColAmount.Index).Value = Me.ObjCd.DigitGrouping("0", Me.ObjCur.GetDefaultPrecission(Me.mPrimaryCurr))
            If Me.mTypeInOut = ClsTransaksiJournal.EnumTipeInOut._In Then
                row.Cells(Me.ColTipeInOut.Index).Value = Me.mTypeInOut
            Else : Me.mTypeInOut = ClsTransaksiJournal.EnumTipeInOut.out
                row.Cells(Me.ColTipeInOut.Index).Value = Me.mTypeInOut
            End If
            row.Cells(Me.ColSPaid.Index).Value = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Function ApakahDatabenar() As Boolean
        Dim row As DataGridViewRow
        Try
            Me.dgvCheque.AllowUserToAddRows = False
            For Each row In Me.dgvCheque.Rows
                If String.IsNullOrEmpty(row.Cells(Me.ColChequeNo.Index).Value) = True Then
                    MsgBox("Cheque no can not be empty", MsgBoxStyle.Information, "Information")
                    row.Cells(Me.ColChequeNo.Index).Selected = True
                    Return False
                    Exit Function
                End If
                If String.IsNullOrEmpty(row.Cells(Me.ColBank.Index).Value) = True Then
                    MsgBox("Bank can not be empty", MsgBoxStyle.Information, "Information")
                    row.Cells(Me.ColBank.Index).Selected = True
                    Return False
                    Exit Function
                End If
                If IsDate(row.Cells(Me.ColDate.Index).Value) = False Or _
                row.Cells(Me.ColDate.Index).Value.ToString = "" Then
                    MsgBox("Date can not be empty", MsgBoxStyle.Information, "Information")
                    row.Cells(Me.ColDate.Index).Selected = True
                    Return False
                    Exit Function
                End If
                If row.Cells(Me.ColAmount.Index).Value <= 0 Then
                    MsgBox("Please enter a positive number for amount ", MsgBoxStyle.Information, "Information")
                    row.Cells(Me.ColAmount.Index).Selected = True
                    Return False
                    Exit Function
                End If
            Next

            Me.dgvCheque.AllowUserToAddRows = True
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub FrmTransaksiJournal_CashBank_Cheque_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim row As DataGridViewRow
        Try
            Me.dgvCheque.AllowUserToAddRows = False
            If Me.mBaru = False And Me.dgvCheque.Rows.Count <= 0 Then
                ObjCheque.IsiDGV(Me.dgvCheque, mPrimaryJurnal, Me)
            End If
            Me.dgvCheque.AllowUserToAddRows = True

            For Each row In Me.dgvCheque.Rows
                row.Cells(Me.ColPrimaryCurrency.Index).Value = Me.mPrimaryCurr
                row.Cells(Me.ColTipeInOut.Index).Value = Me.mTypeInOut
                row.Cells(Me.ColCurrency.Index).Value = Me.ObjCur.GetCurrencyCode(Me.mPrimaryCurr)
                row.Cells(Me.ColAmount.Index).Value = Me.ObjCd.DigitGrouping(row.Cells(Me.ColAmount.Index).Value, _
                Me.ObjCur.GetDefaultPrecission(Me.mPrimaryCurr))
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class