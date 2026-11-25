Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_ExplainGIT
    Public vPrimaryProduct As Double = -1
    Public vDT_ItemGIT As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemgitDataTable
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public vUnitCode As String = ""

    Dim DefPrec As Integer = 0
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjCust As New Core_Module.clsMasterSupplier
    Dim ObjGIT As New ClsTransaksiPackingList

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_ExplainSOB_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim DV As Data.DataRowView
        Dim rowDGV As DataGridViewRow
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemgitRow

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.ColDate.DefaultCellStyle.Format = "dd MMM yyyy"
            Me.DefPrec = Me.ObjProduct.GetAngkaDecimal(Me.vPrimaryProduct)
            Me.vDT_ItemGIT.DefaultView.RowFilter = "[Primary Product] = " + Me.vPrimaryProduct.ToString
            Me.vDT_ItemGIT.DefaultView.Sort = "[Tgl Packing List], [No Packing List]"

            For Each DV In Me.vDT_ItemGIT.DefaultView
                row = DV.Row
                Me.dgvData.Rows.Add()
                rowDGV = Me.dgvData.Rows(Me.dgvData.Rows.Count - 1)
                rowDGV.Cells(Me.ColDate.Index).Value = row.Tgl_Packing_List
                rowDGV.Cells(Me.ColOrderQty.Index).Value = ObjCD.DigitGrouping(row.qty, DefPrec)
                rowDGV.Cells(Me.ColGITNO.Index).Value = row.No_Packing_List
                rowDGV.Cells(Me.ColSupplier.Index).Value = Me.ObjCust.GetSupplierName(row.primary_suppier_rec)
                rowDGV.Cells(Me.ColSupplierPackingListNo.Index).Value = Me.ObjGIT.GetSupplierPackingListNo(row.primary_packing_list_rec)
                rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvData.RowCount
                rowDGV.Cells(Me.ColUnit.Index).Value = Me.vUnitCode
            Next

            Me.dgvData.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class