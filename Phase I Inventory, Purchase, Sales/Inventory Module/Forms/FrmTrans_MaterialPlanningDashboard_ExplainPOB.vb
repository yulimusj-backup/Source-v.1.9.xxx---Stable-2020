Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_ExplainPOB
    Public vPrimaryProduct As Double = -1
    Public vDT_ItemPOB As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempobDataTable
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public vUnitCode As String = ""

    Dim DefPrec As Integer = 0
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjCust As New Core_Module.clsMasterSupplier

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_ExplainSOB_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim DV As Data.DataRowView
        Dim rowDGV As DataGridViewRow
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempobRow

        Try
            Me.ShowIcon = False
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.ColDate.DefaultCellStyle.Format = "dd MMM yyyy"
            Me.DefPrec = Me.ObjProduct.GetAngkaDecimal(Me.vPrimaryProduct)
            Me.vDT_ItemPOB.DefaultView.RowFilter = "[Primary Product] = " + Me.vPrimaryProduct.ToString
            Me.vDT_ItemPOB.DefaultView.Sort = "[Tanggal Transaksi], [No Transaksi]"
            For Each DV In Me.vDT_ItemPOB.DefaultView
                row = DV.Row
                Me.dgvData.Rows.Add()
                rowDGV = Me.dgvData.Rows(Me.dgvData.Rows.Count - 1)
                rowDGV.Cells(Me.ColDate.Index).Value = row.Tanggal_Transaksi
                rowDGV.Cells(Me.ColOrderQty.Index).Value = ObjCD.DigitGrouping(row.qty, DefPrec)
                rowDGV.Cells(Me.ColPONumber.Index).Value = row.No_Transaksi
                rowDGV.Cells(Me.ColSupplier.Index).Value = Me.ObjCust.GetSupplierName(row.primary_supplier_rec)
                rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvData.RowCount
                rowDGV.Cells(Me.ColUnit.Index).Value = Me.vUnitCode
            Next

            Me.dgvData.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class