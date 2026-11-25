Option Explicit On
Option Strict Off

Public Class FrmTransaksiMonthlyClosing_CyclicBundlingExplain2
    Public mTBLItemDetail As New DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable
    Public mPrimaryRefDoc As Double = -1
    Public mBackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public mRefPeriod As Date = Today
    Public mDefPrecCur As Integer
    Public mTipRefDoc As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust
    Public mDocNo As String = ""

    Dim rowItemDetail As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim mTipRefDoc1 As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling

    Private Sub FrmTransaksiMonthlyClosing_CyclicBundlingExplain2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim DGV As DataGridView

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            Me.dgv1.AlternatingRowsDefaultCellStyle.BackColor = Me.mBackColorAll
            Me.dgv2.AlternatingRowsDefaultCellStyle.BackColor = Me.mBackColorAll
            Me.dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgv2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgv2.Rows.Clear()
            Me.dgv1.Rows.Clear()

            If Me.mTipRefDoc <> ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling And _
            Me.mTipRefDoc <> ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
                Me.SplitContainer1.Panel2Collapsed = True
            Else
                Me.gbo1.Text = "Detail Item Before"
                Me.gbo2.Text = "Detail Item After"
            End If

            'isikan dgv 1
            Me.mTBLItemDetail.DefaultView.RowFilter = "[tipe transaksi asal] = " + CInt(Me.mTipRefDoc).ToString + " AND [primary transaksi asal] = " + Me.mPrimaryRefDoc.ToString
            If Me.mTipRefDoc = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Then
                DGV = Me.dgv2
            Else
                DGV = Me.dgv1
            End If
            Me.AddRow(DGV)

            'isikan dgv 2
            If Me.mTipRefDoc = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
                Me.mTBLItemDetail.DefaultView.RowFilter = "[tipe transaksi asal] = " + CInt(Me.mTipRefDoc1).ToString + _
                " AND [primary transaksi asal] = " + Me.mPrimaryRefDoc.ToString
                Me.AddRow(Me.dgv2)
            ElseIf Me.mTipRefDoc = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Then
                Me.mTipRefDoc1 = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling
                Me.mTBLItemDetail.DefaultView.RowFilter = "[tipe transaksi asal] = " + CInt(Me.mTipRefDoc1).ToString + _
                " AND [primary transaksi asal] = " + Me.mPrimaryRefDoc.ToString
                Me.AddRow(Me.dgv1)
            End If

            Me.Col1_Amount.DefaultCellStyle.Format = "N" + Me.mDefPrecCur.ToString
            Me.dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgv2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Me.Text = Me.Text + ": " + Me.mDocNo
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddRow(ByVal DGV As DataGridView)
        Dim DV As DataRowView
        Dim rowDGV As DataGridViewRow

        Try
            For Each DV In Me.mTBLItemDetail.DefaultView
                Me.rowItemDetail = DV.Row
                DGV.Rows.Add()
                rowDGV = DGV.Rows(DGV.Rows.Count - 1)

                rowDGV.Cells(Me.Col1_Amount.Index).Value = Me.rowItemDetail.amount
                rowDGV.Cells(Me.Col1_CodeNo.Index).Value = Me.ObjProduct.GetProductCode(Me.rowItemDetail.primary_code_no)
                rowDGV.Cells(Me.Col1_Desc.Index).Value = Me.ObjProduct.GetDescription(Me.rowItemDetail.primary_code_no)
                rowDGV.Cells(Me.Col1_No.Index).Value = DGV.Rows.Count
                rowDGV.Cells(Me.Col1_Qty.Index).Value = Me.ObjCD.DigitGrouping(Me.rowItemDetail.qty, Me.ObjProduct.GetAngkaDecimal(Me.rowItemDetail.primary_code_no))
                rowDGV.Cells(Me.Col1_Unit.Index).Value = Me.ObjProduct.GetUnitCode(Me.rowItemDetail.primary_code_no)
                rowDGV.Cells(Me.Col1_PrimaryCodeNo.Index).Value = Me.rowItemDetail.primary_code_no
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgv1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv1.CellClick
        Try
            If e.RowIndex <> -1 Then
                If e.ColumnIndex = Me.Col1_CodeNo.Index Then
                    Me.dgv1.ClearSelection()
                    Me.ShowDetailCodeNo(Me.dgv1.Item(Me.Col1_PrimaryCodeNo.Index, e.RowIndex).Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgv2_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv2.CellClick
        Try
            If e.RowIndex <> -1 Then
                If e.ColumnIndex = Me.Col1_CodeNo.Index Then
                    Me.dgv2.ClearSelection()
                    Me.ShowDetailCodeNo(Me.dgv1.Item(Me.Col1_PrimaryCodeNo.Index, e.RowIndex).Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowDetailCodeNo(ByVal vPrimaryCodeNo As Double)
        Try
            Dim ij As New FrmTransaksiMonthlyClosing_CyclicBundlingExplain1

            ij.mBackColorAll = Me.mBackColorAll
            ij.mDefPrecCur = Me.mDefPrecCur
            ij.mDefPrecQty = Me.ObjProduct.GetAngkaDecimal(vPrimaryCodeNo)
            ij.mPrimaryCodeNo = vPrimaryCodeNo
            ij.mRefPeriod = Me.mRefPeriod
            ij.mTBLItemDetail = Me.mTBLItemDetail

            ij.ShowDialog()
            ij.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class