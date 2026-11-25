Option Explicit On
Option Strict Off

Public Class ClsTransaksiPOItem
    Public Sub IsiDGVPOB(ByVal vPrimaryProduct As Double, ByVal vDGV As DataGridView, ByVal vFrm As Inventory_Module.FrmMasterKondisiStock, _
    ByVal vTBL As DseTransaksiPOItem.tbltrans_po_itemDataTable)
        Dim vDTA As New DseTransaksiPOItemTableAdapters.tbltrans_po_itemTableAdapter
        Dim row As DseTransaksiPOItem.tbltrans_po_itemRow
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGV.DataSource = Nothing

            vDTA.Connection.ConnectionString = Core_Module.GetConstring
            If vPrimaryProduct = -1 Then
                vDTA.Fill(vTBL)
            Else
                vDTA.FillBy_PrimaryProduct(vTBL, vPrimaryProduct)
            End If
            For Each row In vTBL.Rows
                row.Order_Balance = ObjCD.DigitGrouping(row.POB, row._Decimal)
                If row.Delivery_Date = #1/1/1950# Then
                    row.Estimate_Delivery = "-"
                Else
                    row.Estimate_Delivery = Format(row.Delivery_Date, "dd MMM yyyy")
                End If
            Next
            vDGV.DataSource = vTBL.DefaultView
            vDGV.Columns(0).HeaderText = "PO Number"
            vDGV.Columns(1).HeaderText = "Date"
            vDGV.Columns(1).DefaultCellStyle.Format = "dd MMM yyyy"
            vDGV.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            vDGV.Columns(6).Visible = False
            vDGV.Columns(7).Visible = False
            vDGV.Columns(8).Visible = False
            vDGV.Columns(9).Visible = False
            vDGV.Columns(10).Visible = False
            If vPrimaryProduct <> -1 Then
                vDGV.Columns(4).Visible = False
                vDGV.Columns(5).Visible = False
            End If
            vDGV.Columns(11).SortMode = DataGridViewColumnSortMode.Programmatic
            vDGV.Columns(12).SortMode = DataGridViewColumnSortMode.Programmatic

            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            vDTA.Dispose()
            ObjCD = Nothing
        End Try
    End Sub

    Public Function GetPOB(ByVal vPrimaryProduct As Double) As Decimal
        Dim DTA As New DseTransaksiPOItemTableAdapters.tbltrans_po_item1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPOB(vPrimaryProduct).HasValue = True Then
                Return DTA.ScalarQuery_GetPOB(vPrimaryProduct).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
