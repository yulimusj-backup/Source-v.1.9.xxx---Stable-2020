Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_ExplainPendingDO
    Public vPrimaryProduct As Double = -1
    Public vDT_ItemDO As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdoDataTable
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public vUnitCode As String = ""

    Dim DefPrec As Integer = 0
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjCust As New Core_Module.clsMasterCustomer

    Private Sub FrmTrans_MaterialPlanningDashboard_ExplainPendingDO_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim DV As Data.DataRowView
        Dim rowDGV As DataGridViewRow
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdoRow

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.ColDate.DefaultCellStyle.Format = "dd MMM yyyy"
            Me.DefPrec = Me.ObjProduct.GetAngkaDecimal(Me.vPrimaryProduct)
            Me.vDT_ItemDO.DefaultView.RowFilter = "[Primary Product] = " + Me.vPrimaryProduct.ToString
            Me.vDT_ItemDO.DefaultView.Sort = "[DO Date], [DO Number]"
            For Each DV In Me.vDT_ItemDO.DefaultView
                row = DV.Row
                Me.dgvData.Rows.Add()
                rowDGV = Me.dgvData.Rows(Me.dgvData.Rows.Count - 1)
                rowDGV.Cells(Me.ColDate.Index).Value = row.DO_Date
                rowDGV.Cells(Me.ColDOQty.Index).Value = ObjCD.DigitGrouping(row.qty, DefPrec)
                rowDGV.Cells(Me.ColDONumber.Index).Value = row.DO_Number
                rowDGV.Cells(Me.ColCustomer.Index).Value = Me.ObjCust.GetCustName(row.primary_customer_rec)
                rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvData.RowCount
                rowDGV.Cells(Me.ColUnit.Index).Value = Me.vUnitCode
            Next

            Me.dgvData.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class