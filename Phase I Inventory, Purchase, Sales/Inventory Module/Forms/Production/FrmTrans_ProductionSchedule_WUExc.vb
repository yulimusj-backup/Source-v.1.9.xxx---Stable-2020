Public Class FrmTrans_ProductionSchedule_WUExc
    Public TBLDummy As DseTrans_ManufacturingOrder.tbltrans_mo_item_wo3DataTable

    Dim ObjCD As New Core_Module.ClsCekData
    Dim rowDummy As DseTrans_ManufacturingOrder.tbltrans_mo_item_wo3Row

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ProductionSchedule_WUExc_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDummy In Me.TBLDummy.Rows
                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                rowDGV.Cells(Me.ColCodeNo.Index).Value = rowDummy.Product_Code
                rowDGV.Cells(Me.ColDesc.Index).Value = rowDummy.Description
                rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvItem.Rows.Count
                rowDGV.Cells(Me.ColJONo.Index).Value = rowDummy.JO_Number
                rowDGV.Cells(Me.ColMONo.Index).Value = rowDummy.MO_Number
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class