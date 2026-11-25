Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_ExplainMOB
    Public vPrimaryProduct As Double = -1
    Public vDT_ItemMOB As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemmobDataTable
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim DefPrec As Integer = 0
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjCD As New Core_Module.ClsCekData
    Dim TBLProses As New Core_Module.DseMasterProses.tblmaster_proses1DataTable
    Dim ObjProses As New Core_Module.ClsMasterProses

    Private Function GetProses(ByVal vPrimaryProses As Double) As String
        Dim rowProses As Core_Module.DseMasterProses.tblmaster_proses1Row
        Try
            For Each rowProses In Me.TBLProses.Rows
                If rowProses.Primary = vPrimaryProses Then
                    Return rowProses.Code
                End If
            Next

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub FrmTrans_MaterialPlanningDashboard_ExplainMOB_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim DV As Data.DataRowView
        Dim rowDGV As DataGridViewRow
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemmobRow
        Dim UnitCode As String = ""
        Dim PrimaryMOMin, PrimaryMOMax As Double
        Dim DTA1 As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item4TableAdapter
        Dim TBL1 As New DseTrans_ManufacturingOrder.tbltrans_mo_item4DataTable
        Dim row1 As DseTrans_ManufacturingOrder.tbltrans_mo_item4Row

        Try
            Me.Text = Me.Text + " - " + Me.ObjProduct.GetProductCode(Me.vPrimaryProduct)
            Me.ObjProses.FillTabelProses(Me.TBLProses)
            DTA1.Connection.ConnectionString = Core_Module.GetConstring

            PrimaryMOMin = Double.MaxValue
            PrimaryMOMax = Double.MinValue

            Me.dgvDataMOB.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvDataMOB.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            UnitCode = Me.ObjProduct.GetUnitCode(Me.vPrimaryProduct)

            Me.ColMODate.DefaultCellStyle.Format = "dd MMM yyyy"
            Me.ColTargetFinishDate.DefaultCellStyle.Format = "dd MMM yyyy"
            Me.DefPrec = Me.ObjProduct.GetAngkaDecimal(Me.vPrimaryProduct)
            Me.ColBalanceQty.DefaultCellStyle.Format = "N" + Me.DefPrec.ToString

            Me.vDT_ItemMOB.DefaultView.RowFilter = "[Primary Product] = " + Me.vPrimaryProduct.ToString
            Me.vDT_ItemMOB.DefaultView.Sort = "[mo date rec], [mo no rec], [jo no rec]"

            For Each DV In Me.vDT_ItemMOB.DefaultView
                row = DV.Row
                Me.dgvDataMOB.Rows.Add()

                rowDGV = Me.dgvDataMOB.Rows(Me.dgvDataMOB.Rows.Count - 1)
                rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvDataMOB.RowCount
                rowDGV.Cells(Me.ColBalanceQty.Index).Value = row.qty
                rowDGV.Cells(Me.ColJONo.Index).Value = row.jo_no_rec
                rowDGV.Cells(Me.ColMODate.Index).Value = row.mo_date_rec
                rowDGV.Cells(Me.ColMONo.Index).Value = row.mo_no_rec
                rowDGV.Cells(Me.ColProgress.Index).Value = "Show"
                rowDGV.Cells(Me.ColTargetFinishDate.Index).Value = #1/1/1950#
                rowDGV.Cells(Me.ColUnitCode.Index).Value = UnitCode
                rowDGV.Cells(Me.ColIdentifierJO.Index).Value = row.identifier_jo_rec
                rowDGV.Cells(Me.ColIdentifierWO.Index).Value = row.identifier_wo_rec
                rowDGV.Cells(Me.ColWONo.Index).Value = row.wo_no_rec
                rowDGV.Cells(Me.ColProses.Index).Value = Me.GetProses(row.primary_proses)

                If PrimaryMOMax < row.primary_mo_rec Then
                    PrimaryMOMax = row.primary_mo_rec
                End If
                If PrimaryMOMin > row.primary_mo_rec Then
                    PrimaryMOMin = row.primary_mo_rec
                End If
            Next

            Me.dgvDataMOB.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class