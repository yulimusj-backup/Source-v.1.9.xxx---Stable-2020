Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_ExplainQtyOnHand
    Public vDT_ItemQty As New DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_qtyDataTable
    Public vDT_ItemSettingLocation As New DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_location_exclusionDataTable
    Public vPrimaryProduct As Double = -1
    Public vBackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UnitCode As String = ""
    Public DefPrec As Integer = 0

    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim vrowItemQty As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_qtyRow
    Dim ObjLocation As New Core_Module.clsMasterStorageArea
    Dim ObjWarehouse As New Core_Module.clsMasterWarehouse

    Private Sub FrmTrans_MaterialPlanningDashboard_ExplainQtyOnHand_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim rowDGV As DataGridViewRow
        Dim AdaBatch As Boolean = False
        Dim AdaExpiry As Boolean = False
        Dim AdaLocation As Boolean = False

        Try
            Me.DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Me.DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Me.vBackColorAll
            Me.Text = "Qty Data Detail: " + Me.ObjProduct.GetProductCode(vPrimaryProduct)

            Me.vDT_ItemQty.DefaultView.RowFilter = "[Primary Product] = " + vPrimaryProduct.ToString
            For Each DV As Data.DataRowView In Me.vDT_ItemQty.DefaultView
                Me.vrowItemQty = DV.Row
                Me.DataGridView1.Rows.Add()
                rowDGV = Me.DataGridView1.Rows(Me.DataGridView1.Rows.Count - 1)
                With rowDGV
                    .Cells(Me.ColBatchNo.Index).Value = Me.vrowItemQty.Lot_No
                    .Cells(Me.ColExpiry.Index).Value = Me.vrowItemQty.Expiry_Date
                    .Cells(Me.ColLocation.Index).Value = Me.ObjLocation.GetStorageAreaCode(Me.vrowItemQty.Primary_Storage_Area)
                    .Cells(Me.ColQty.Index).Value = Me.vrowItemQty.Qty
                    .Cells(Me.ColUnit.Index).Value = UnitCode
                    .Cells(Me.ColWarehouse.Index).Value = Me.ObjLocation.GetWarehouseCode(Me.vrowItemQty.Primary_Storage_Area)

                    If Me.vrowItemQty.Lot_No <> "" Then
                        AdaBatch = True
                    End If
                    If Me.vrowItemQty.Primary_Storage_Area <> -1 Then
                        AdaLocation = True
                    End If
                    If Me.vrowItemQty.Expiry_Date <> #1/1/1950# Then
                        AdaExpiry = True
                    End If
                End With
            Next

            If AdaExpiry = False Then
                Me.ColExpiry.Visible = False
            Else
                Me.ColExpiry.Visible = True
            End If
            If AdaBatch = False Then
                Me.ColBatchNo.Visible = False
            Else
                Me.ColExpiry.Visible = True
            End If
            If AdaExpiry = False Then
                Me.ColExpiry.Visible = False
            Else
                Me.ColExpiry.Visible = True
            End If

            Me.ColQty.DefaultCellStyle.Format = "N" + DefPrec.ToString
            Me.ColExpiry.DefaultCellStyle.Format = "dd MMM yyyy"

            Me.DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.Close()
    End Sub

    Private Sub llbMP_LocationSetting_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbMP_LocationSetting.LinkClicked
        Dim i As New FrmTrans_MaterialPlanningDashboard_ExplainWarehouseExc

        Try
            If Me.vDT_ItemSettingLocation.Rows.Count = 0 Then
                MsgBox("No location setting for this transaction.", MsgBoxStyle.Information, "Information")
            Else
                i.TBL = Me.vDT_ItemSettingLocation
                i.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i = Nothing
        End Try
    End Sub
End Class