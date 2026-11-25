''rosadin-r16
Option Explicit On
Option Strict Off
Public Class FrmTrans_ProductionSchedule_SplitMerge_Rate
    Public UserCode As String = String.Empty
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public TblSchBar As New DseTrans_Schedule.tbl_sch_barDataTable
    Public PrimaryWU_Destination As Double = -1

    Dim rowSchBar As DseTrans_Schedule.tbl_sch_barRow
    Dim objCD As New Core_Module.ClsCekData
    Dim rowDgv As DataGridViewRow
    Dim buffer As String = ""
    Dim tblWU_Product As New DseTrans_ProductionSchedule_ChangeWU.tblmaster_product_item_workunitDataTable
    Dim rowWU_Product As DseTrans_ProductionSchedule_ChangeWU.tblmaster_product_item_workunitRow

    Private Sub FrmTrans_ProductionSchedule_ChangeWU_Rate_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim dtaWU_Product As New DseTrans_ProductionSchedule_ChangeWUTableAdapters.tblmaster_product_item_workunitTableAdapter
        Try
            If TblSchBar.Columns.Contains("Qty Sisa") = False Then
                TblSchBar.Columns.Add("Qty Sisa", System.Type.GetType("System.Decimal"))
                TblSchBar.Columns("Qty Sisa").DefaultValue = 0
            End If

            dtaWU_Product.Connection.ConnectionString = Core_Module.GetConstring()
            dtaWU_Product.ClearBeforeFill = False
            Me.tblWU_Product.Clear()
            For Each Me.rowSchBar In Me.TblSchBar
                dtaWU_Product.Fill(Me.tblWU_Product, Me.rowSchBar.primary_product, Me.PrimaryWU_Destination)
            Next

            Me.dgvItem.SelectionMode = DataGridViewSelectionMode.CellSelect
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.Rows.Clear()
            For Each Me.rowSchBar In Me.TblSchBar
                Me.rowSchBar.primary_work_unit = Me.PrimaryWU_Destination
                Me.rowDgv = Me.dgvItem.Rows(Me.dgvItem.Rows.Add)
                With Me.rowDgv
                    .Cells(Me.ColMoNo.Index).Value = rowSchBar.mo_no
                    .Cells(Me.ColJONo.Index).Value = rowSchBar.jo_no
                    .Cells(Me.ColProcessCode.Index).Value = rowSchBar.process_code
                    .Cells(Me.ColWONo.Index).Value = rowSchBar.wo_no
                    .Cells(Me.ColCodeNo.Index).Value = rowSchBar.product_code
                    .Cells(Me.ColDesc.Index).Value = rowSchBar.product_desc
                    .Cells(Me.colQtyOri.Index).Value = rowSchBar.qty_produce
                    .Cells(Me.ColQty.Index).Value = Me.objCD.DigitGrouping(rowSchBar.qty_produce, rowSchBar.decimal_point)
                    .Cells(Me.ColUnit.Index).Value = rowSchBar.unit_code
                    .Cells(Me.ColRate_Ori.Index).Value = Me.objCD.DigitGrouping(rowSchBar.rate, rowSchBar.decimal_point)
                    .Cells(Me.ColWH_Ori.Index).Value = Me.objCD.Convert_ToString(rowSchBar.total_work_hour)

                    .Cells(Me.ColRate_New.Index).Value = Me.objCD.DigitGrouping(Me.GetNewRate(rowSchBar.primary_product), rowSchBar.decimal_point)
                    .Cells(Me.ColRateEfficiency_New.Index).Value = Me.objCD.DigitGrouping(Me.GetNewEfficiency(rowSchBar.primary_product), 2)
                    .Cells(Me.ColRateEffective_New.Index).Value = Me.objCD.DigitGrouping(.Cells(Me.ColRate_New.Index).Value * (.Cells(Me.ColRateEfficiency_New.Index).Value / 100), rowSchBar.decimal_point)

                    .Cells(Me.ColWH_New.Index).Value = Me.objCD.Convert_ToString(Me.objCD.DigitGrouping(.Cells(Me.ColQty.Index).Value / .Cells(Me.ColRateEffective_New.Index).Value * 60, 0, Core_Module.ClsCekData.EnumTipeRounding.Up))

                    .Cells(Me.ColPrimaryMO.Index).Value = rowSchBar.primary_mo
                    .Cells(Me.ColIdentifierJO.Index).Value = rowSchBar.identifier_jo
                    .Cells(Me.ColPrimaryProcess.Index).Value = rowSchBar.primary_process
                    .Cells(Me.ColIdentifierWO.Index).Value = rowSchBar.identifier_wo
                    .Cells(Me.ColPrimaryProduct.Index).Value = rowSchBar.primary_product
                    .Cells(Me.ColPrimaryUnit.Index).Value = rowSchBar.primary_unit

                    .Cells(Me.ColHide_Rate_New.Index).Value = .Cells(Me.ColRate_New.Index).Value
                    .Cells(Me.ColHide_Rate_Efficiency_New.Index).Value = .Cells(Me.ColRateEfficiency_New.Index).Value
                    .Cells(Me.ColHide_Rate_Effective_New.Index).Value = .Cells(Me.ColRateEffective_New.Index).Value
                    .Cells(Me.ColHide_WH_New.Index).Value = .Cells(Me.ColWH_New.Index).Value
                    .Cells(Me.ColDecimalPoint.Index).Value = rowSchBar.decimal_point
                End With
            Next
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.RadioButton2.Checked = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function GetNewRate(ByVal vPrimaryProduct As Double) As Decimal
        Try
            Me.tblWU_Product.DefaultView.RowFilter = "[primary product] = " & vPrimaryProduct.ToString
            If Me.tblWU_Product.DefaultView.Count > 0 Then
                Me.rowWU_Product = Me.tblWU_Product.DefaultView(0).Row
                Return Me.rowWU_Product.rate_per_hour
            Else
                MsgBox("Logic Error.", MsgBoxStyle.Critical, "Error")
                Return 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 1
        End Try
    End Function

    Function GetNewEfficiency(ByVal vPrimaryProduct As Double) As Decimal
        Try
            Me.tblWU_Product.DefaultView.RowFilter = "[primary product] = " & vPrimaryProduct.ToString
            If Me.tblWU_Product.DefaultView.Count > 0 Then
                Me.rowWU_Product = Me.tblWU_Product.DefaultView(0).Row
                Return Me.rowWU_Product.efficiency
            Else
                MsgBox("Logic Error.", MsgBoxStyle.Critical, "Error")
                Return 100
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 100
        End Try
    End Function

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColQty.Index Or (Me.RadioButton3.Checked And (e.ColumnIndex = Me.ColRate_New.Index Or e.ColumnIndex = Me.ColRateEfficiency_New.Index Or e.ColumnIndex = Me.ColRateEffective_New.Index)) Then
                Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        Try
            If Me.RadioButton1.Checked Then
                Me.ColRate_New.Visible = False
                Me.ColRateEfficiency_New.Visible = False
                Me.ColRateEffective_New.Visible = False
                Me.ColWH_New.Visible = False
            Else
                Me.ColRate_New.Visible = True
                Me.ColRateEfficiency_New.Visible = True
                Me.ColRateEffective_New.Visible = True
                Me.ColWH_New.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        Try
            For Each Me.rowDgv In Me.dgvItem.Rows
                With Me.rowDgv
                    .Cells(Me.ColRate_New.Index).Value = .Cells(Me.ColHide_Rate_New.Index).Value
                    .Cells(Me.ColRateEfficiency_New.Index).Value = .Cells(Me.ColHide_Rate_Efficiency_New.Index).Value
                    .Cells(Me.ColRateEffective_New.Index).Value = .Cells(Me.ColHide_Rate_Effective_New.Index).Value
                    .Cells(Me.ColWH_New.Index).Value = .Cells(Me.ColHide_WH_New.Index).Value
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If String.IsNullOrEmpty(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                        Return
                    End If
                    If Me.objCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Return
                    End If
                    If e.ColumnIndex = Me.ColRate_New.Index Or e.ColumnIndex = Me.ColQty.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, Me.dgvItem.Item(Me.ColDecimalPoint.Index, e.RowIndex).Value)
                        If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) <= 0 Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Return
                        End If
                    End If
                    If e.ColumnIndex = Me.ColRateEfficiency_New.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                        If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) <= 0 Or CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > 100 Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Return
                        End If
                    End If
                    Me.dgvItem.Item(Me.ColRateEffective_New.Index, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.dgvItem.Item(Me.ColRate_New.Index, e.RowIndex).Value * _
                    (Me.dgvItem.Item(Me.ColRateEfficiency_New.Index, e.RowIndex).Value / 100), Me.dgvItem.Item(Me.ColDecimalPoint.Index, e.RowIndex).Value)
                    Me.dgvItem.Item(Me.ColWH_New.Index, e.RowIndex).Value = Me.objCD.Convert_ToString(Me.objCD.DigitGrouping(Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value / _
                    Me.dgvItem.Item(Me.ColRateEffective_New.Index, e.RowIndex).Value * 60, 0, Core_Module.ClsCekData.EnumTipeRounding.Up))
                Else
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
            Else
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If Me.RadioButton2.Checked Or Me.RadioButton3.Checked Then
                For Each Me.rowDgv In Me.dgvItem.Rows
                    Me.TblSchBar.DefaultView.RowFilter = "[primary mo]=" & Me.rowDgv.Cells(Me.ColPrimaryMO.Index).Value & " and [identifier jo]='" & Me.rowDgv.Cells(Me.ColIdentifierJO.Index).Value & "' and [identifier wo]='" & Me.rowDgv.Cells(Me.ColIdentifierWO.Index).Value & "' and [primary work unit]=" & Me.PrimaryWU_Destination.ToString
                    For Each DV As Data.DataRowView In Me.TblSchBar.DefaultView
                        Me.rowSchBar = DV.Row
                        Me.rowSchBar.rate = Me.rowDgv.Cells(Me.ColRateEffective_New.Index).Value
                        Me.rowSchBar.total_work_hour = Me.objCD.Convert_ToMinutes(Me.rowDgv.Cells(Me.ColWH_New.Index).Value)
                        Me.rowSchBar.qty_produce = CDec(Me.rowDgv.Cells(Me.ColQty.Index).Value)
                        Me.rowSchBar.Item("Qty Sisa") = Me.objCD.DigitGrouping(CDec(Me.rowDgv.Cells(Me.colQtyOri.Index).Value) - CDec(Me.rowDgv.Cells(Me.ColQty.Index).Value), Me.rowDgv.Cells(Me.ColDecimalPoint.Index).Value)
                    Next
                Next
            End If
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class