Option Explicit On
Option Strict Off

Public Class FrmTrans_ManufacturingOrder_Exp_QtyUsed
    Public IdentifierWO As String = ""
    Public IdentifierJO As String = ""
    Public PrimaryProduct As Double = -1
    Public PrimaryUnit As Double = -1
    Public CodeNo As String = ""
    Public Desc As String = ""
    Public WONO As String = ""
    Public ProcessCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UnitCode As String = ""

    Dim ObjProduct As New Core_Module.clsMasterProduct

    Private Sub FrmTrans_ManufacturingOrder_Exp_QtyUsed_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltransaksi_stockbundling1TableAdapter
        Dim TBL As New DseTrans_ManufacturingOrder.tbltransaksi_stockbundling1DataTable
        Dim row As DseTrans_ManufacturingOrder.tbltransaksi_stockbundling1Row
        Dim rowDGV As DataGridViewRow
        Dim DefPrec As Integer = 0

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            DefPrec = Me.ObjProduct.GetAngkaDecimal(Me.PrimaryProduct)
            Me.ColQty.DefaultCellStyle.Format = "N" + DefPrec.ToString
            Me.ColStartDate.DefaultCellStyle.Format = "dd MMM yyy HH.mm"
            Me.ColEndDate.DefaultCellStyle.Format = "dd MMM yyy HH.mm"

            Me.DgvItemBundling.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'isikan data header
            Me.TxtCodeNo.Text = Me.CodeNo
            Me.TxtDesc.Text = Me.Desc
            Me.TxtWONumber.Text = Me.WONO
            Me.TxtProcessCode.Text = Me.ProcessCode

            'load data item bundling
            DTA.Fill(TBL, Me.IdentifierWO, Me.PrimaryProduct, Me.IdentifierJO)

            'isikan ke dgv
            Me.DgvItemBundling.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvItemBundling.Rows.Clear()
            For Each row In TBL.Rows
                Me.DgvItemBundling.Rows.Add()
                rowDGV = Me.DgvItemBundling.Rows(Me.DgvItemBundling.Rows.Count - 1)
                rowDGV.Cells(Me.ColDocNo.Index).Value = row.No_Transaksi
                rowDGV.Cells(Me.ColEndDate.Index).Value = row.end_time_actual
                rowDGV.Cells(Me.ColNo.Index).Value = rowDGV.Index + 1
                rowDGV.Cells(Me.ColPrimaryBundling.Index).Value = row.Primary
                rowDGV.Cells(Me.ColQty.Index).Value = row.Qty
                rowDGV.Cells(Me.ColStartDate.Index).Value = row.start_time_actual
                rowDGV.Cells(Me.ColUnit.Index).Value = Me.UnitCode
            Next
            Me.DgvItemBundling.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class