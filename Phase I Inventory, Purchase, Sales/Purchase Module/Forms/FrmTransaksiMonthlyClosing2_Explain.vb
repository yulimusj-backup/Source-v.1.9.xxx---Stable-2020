Option Explicit On
Option Strict Off

Public Class FrmTransaksiMonthlyClosing2_Explain
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Public vCostGroup As String = ""
    Public vBundlingNo As String = ""
    Public vBundlingDate As Date = Today
    Public vProcessCode As String = ""
    Public vTotalOverhead As Decimal = 0
    Public vTotalQty As Decimal = 0
    Public vTotalQtyBundling As Decimal = 0
    Public vUnit As String = ""
    Public vOverheadAllocated As Decimal = 0
    Public vIdentifier As String = ""
    Public vPrimaryBundling As Double = -1
    Public vTBLCostGroup_CodeNo As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_noDataTable

    Dim RefPosition As Integer = 0
    Dim vRowCodeNo As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_noRow
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim DefPrec As Integer = 0

    Private Sub FrmTransaksiMonthlyClosing2_Explain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim rowDGV As DataGridViewRow

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            DefPrec = ObjCur.GetDecimal_DefaultCurrency
            Me.dgvItemCost_DetailTransaction.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.RefPosition = Me.LblTotalQty.Left + Me.LblTotalQty.Width

            'tampilkan lbl header
            Me.LblTotalQty.Text = Me.ObjCD.BuangNolDibelakangDecimal(Me.ObjCD.DigitGrouping(Me.vTotalQty))
            Me.LblCostGroup.Text = Me.vCostGroup
            Me.LblTotalOverheadAmount.Text = Me.ObjCD.DigitGrouping(Me.vTotalOverhead)
            Me.dtpDate.Value = Me.vBundlingDate
            Me.LblCostGroup.Text = Me.vCostGroup
            Me.LblProcess.Text = Me.vProcessCode
            Me.LblDocumentNo.Text = Me.vBundlingNo

            If Me.vUnit = "" Then
                Me.LblUnits.Text = "Units"
            Else
                Me.LblUnits.Text = Me.vUnit
            End If
            Me.LblUnits_Bundling.Text = Me.LblUnits.Text

            'tampilkan dgv item
            Me.vTBLCostGroup_CodeNo.DefaultView.RowFilter = "identifier = '" + Me.vIdentifier + "' AND [primary bundling] = " + Me.vPrimaryBundling.ToString
            Me.vTBLCostGroup_CodeNo.DefaultView.Sort = "[Tanggal Transaksi], [Bundling No], [primary code no]"
            Me.dgvItemCost_DetailTransaction.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.vOverheadAllocated = 0
            Me.vTotalQtyBundling = 0
            For Each DV As DataRowView In Me.vTBLCostGroup_CodeNo.DefaultView
                Me.vRowCodeNo = DV.Row
                Me.dgvItemCost_DetailTransaction.Rows.Add()
                rowDGV = Me.dgvItemCost_DetailTransaction.Rows(Me.dgvItemCost_DetailTransaction.Rows.Count - 1)
                With Me.vRowCodeNo
                    rowDGV.Cells(Me.ColCostCodeNoD_AllocatedAmount.Index).Value = Me.ObjCD.DigitGrouping(.allocated_amount, DefPrec)
                    Me.vOverheadAllocated += .allocated_amount
                    rowDGV.Cells(Me.ColCostCodeNoD_CodeNo.Index).Value = .Code_No
                    rowDGV.Cells(Me.ColCostCodeNoD_Desc.Index).Value = .Desc
                    rowDGV.Cells(Me.ColCostCodeNoD_Identifier.Index).Value = .identifier
                    rowDGV.Cells(Me.ColCostCodeNoD_PrimaryCodeNo.Index).Value = .primary_code_no
                    rowDGV.Cells(Me.ColCostCodeNoD_PrimaryProcess.Index).Value = .primary_process
                    rowDGV.Cells(Me.ColCostCodeNoD_PrimaryUnit.Index).Value = .primary_unit
                    rowDGV.Cells(Me.ColCostCodeNoD_CostFactor.Index).Value = Me.ObjCD.DigitGrouping(.cost_factor, 2)
                    rowDGV.Cells(Me.ColCostCodeNoD_Qty.Index).Value = Me.ObjCD.DigitGrouping(.qty_ref, .decimal_qty)
                    Me.vTotalQtyBundling += .qty_ref
                    rowDGV.Cells(Me.ColCostCodeNoD_Unit.Index).Value = .Unit_Code
                End With
            Next
            Me.dgvItemCost_DetailTransaction.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.llbAlocatedOverhead.Text = Me.ObjCD.DigitGrouping(Me.vOverheadAllocated, DefPrec)
            Me.LblTotalQty_Bundling.Text = Me.ObjCD.BuangNolDibelakangDecimal(Me.ObjCD.DigitGrouping(Me.vTotalQtyBundling))

            Me.dgvItemCost_DetailTransaction.ClearSelection()
            Me.TxtExplain.Text = ""

            'set ulang posisi
            Me.LblTotalQty.Left = Me.RefPosition - Me.LblTotalQty.Width
            Me.LblTotalQty_Bundling.Left = Me.RefPosition - Me.LblTotalQty_Bundling.Width
            Me.LblTotalOverheadAmount.Left = Me.RefPosition - Me.LblTotalOverheadAmount.Width
            Me.llbAlocatedOverhead.Left = Me.RefPosition - Me.llbAlocatedOverhead.Width
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemCost_DetailTransaction_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItemCost_DetailTransaction.SelectionChanged
        Dim Dummy As Decimal = 0
        Dim DummyStr As String = ""
        Dim rowDGV As DataGridViewRow

        Try
            If Me.dgvItemCost_DetailTransaction.SelectedCells.Count > 0 Then
                If Me.dgvItemCost_DetailTransaction.SelectedCells(0).RowIndex = Me.dgvItemCost_DetailTransaction.RowCount - 1 Then
                    rowDGV = Me.dgvItemCost_DetailTransaction.Rows(Me.dgvItemCost_DetailTransaction.SelectedCells(0).RowIndex)
                    If IsNothing(rowDGV.Cells(Me.ColCostCodeNoD_AllocatedAmount.Index).Value) = False Then
                        'hitung berdasarkan sisanya
                        Dummy = 0
                        For Each rowDGV In Me.dgvItemCost_DetailTransaction.Rows
                            If rowDGV.Index <> Me.dgvItemCost_DetailTransaction.SelectedCells(0).RowIndex Then
                                Dummy += CDec(rowDGV.Cells(Me.ColCostCodeNoD_AllocatedAmount.Index).Value)
                                If DummyStr = "" Then
                                    DummyStr = rowDGV.Cells(Me.ColCostCodeNoD_AllocatedAmount.Index).Value.ToString
                                Else
                                    DummyStr = DummyStr + " + " + rowDGV.Cells(Me.ColCostCodeNoD_AllocatedAmount.Index).Value.ToString
                                End If
                            End If
                        Next
                        If DummyStr = "" Then
                            DummyStr = Me.ObjCD.DigitGrouping(0, DefPrec)
                        End If
                        Me.TxtExplain.Text = ""
                        Me.TxtExplain.Text = "= Overhead Allocated - Other Allocated Amount"
                        Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= " + Me.llbAlocatedOverhead.Text + " - (" + DummyStr + ")"
                        rowDGV = Me.dgvItemCost_DetailTransaction.Rows(Me.dgvItemCost_DetailTransaction.SelectedCells(0).RowIndex)
                        Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= " + rowDGV.Cells(Me.ColCostCodeNoD_AllocatedAmount.Index).Value.ToString
                    Else
                        Me.TxtExplain.Text = ""
                    End If
                Else
                    'hitung berdasarkan cost factor
                    Me.TxtExplain.Text = ""
                    Me.TxtExplain.Text = "= Overhead Allocated * Cost Factor"
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= " + Me.llbAlocatedOverhead.Text + " * " + _
                    Me.dgvItemCost_DetailTransaction.Item(Me.ColCostCodeNoD_CostFactor.Index, _
                    Me.dgvItemCost_DetailTransaction.SelectedCells(0).RowIndex).Value.ToString + " %"
                    rowDGV = Me.dgvItemCost_DetailTransaction.Rows(Me.dgvItemCost_DetailTransaction.SelectedCells(0).RowIndex)
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= " + rowDGV.Cells(Me.ColCostCodeNoD_AllocatedAmount.Index).Value.ToString
                End If
            Else
                Me.TxtExplain.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAlocatedOverhead_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAlocatedOverhead.LinkClicked
        Try
            'menjelaskan angka allocated amount
            Me.dgvItemCost_DetailTransaction.ClearSelection()
            Me.TxtExplain.Text = ""
            Me.TxtExplain.Text = "= (Total Qty In Bundling * Total Overhead) / Total Qty"
            Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= (" + Me.LblTotalQty_Bundling.Text + " * " + Me.LblTotalOverheadAmount.Text + ") / " + Me.LblTotalQty.Text
            Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= " + Me.llbAlocatedOverhead.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class