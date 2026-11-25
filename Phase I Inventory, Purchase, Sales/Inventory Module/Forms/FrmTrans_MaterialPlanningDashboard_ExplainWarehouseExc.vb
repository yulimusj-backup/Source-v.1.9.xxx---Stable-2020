Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_ExplainWarehouseExc
    Public TBL As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_location_exclusionDataTable

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.Close()
    End Sub

    Private Sub LoadDGVWarehouse()
        Dim Dummy As String = ""
        Dim row As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_location_exclusionRow
        Dim rowDGV As DataGridViewRow
        Dim DTA_W As New DseTransaksiMaterialPlanning2TableAdapters.tblmaster_warehouseTableAdapter
        Dim TBL_W As New DseTransaksiMaterialPlanning2.tblmaster_warehouseDataTable
        Dim row_W As DseTransaksiMaterialPlanning2.tblmaster_warehouseRow
        Dim DTA_L As New DseTransaksiMaterialPlanning2TableAdapters.tblmaster_storage_areaTableAdapter
        Dim TBL_L As New DseTransaksiMaterialPlanning2.tblmaster_storage_areaDataTable
        Dim row_L As DseTransaksiMaterialPlanning2.tblmaster_storage_areaRow
        Dim LocationLevel1 As FrmTrans_MaterialPlanningDashboard_Adv_Exclusion.EnumLocationLevel = FrmTrans_MaterialPlanningDashboard_Adv_Exclusion.EnumLocationLevel.Location
        Dim Inc_Exc As FrmTrans_MaterialPlanningDashboard_Adv_Exclusion.EnumExcludeInclude = FrmTrans_MaterialPlanningDashboard_Adv_Exclusion.EnumExcludeInclude.Include

        Try
            DTA_W.Connection.ConnectionString = Core_Module.GetConstring
            DTA_L.Connection.ConnectionString = Core_Module.GetConstring

            Me.dgvWarehouse.AllowUserToAddRows = False
            Me.dgvWarehouse.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvWarehouse.Rows.Clear()

            For Each row In TBL.Rows
                Me.dgvWarehouse.Rows.Add()
                rowDGV = Me.dgvWarehouse.Rows(Me.dgvWarehouse.Rows.Count - 1)

                rowDGV.Cells(Me.ColWarehouse_No.Index).Value = rowDGV.Index + 1
                rowDGV.Cells(Me.ColWarehouse_PrimaryAcuan.Index).Value = row.primary_acuan
                If row.level = LocationLevel1 Then
                    If TBL_L.Rows.Count = 0 Then
                        DTA_L.Fill(TBL_L)
                    End If
                    TBL_L.DefaultView.RowFilter = "Primary = " + row.primary_acuan.ToString
                    If TBL_L.DefaultView.Count = 0 Then
                        MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    row_L = TBL_L.DefaultView(0).Row
                    'location
                    rowDGV.Cells(Me.ColWarehouse_Code.Index).Value = row_L.Storage_Area_Code
                    rowDGV.Cells(Me.ColWarehouse_Desc.Index).Value = row_L.Description
                    rowDGV.Cells(Me.ColWarehouse_Level.Index).Value = "Location"
                Else
                    'warehouse
                    If TBL_W.Rows.Count = 0 Then
                        DTA_W.Fill(TBL_W)
                    End If
                    TBL_W.DefaultView.RowFilter = "Primary = " + row.primary_acuan.ToString
                    If TBL_W.DefaultView.Count = 0 Then
                        MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    row_W = TBL_W.DefaultView(0).Row
                    'location
                    rowDGV.Cells(Me.ColWarehouse_Code.Index).Value = row_W.Warehouse_Code
                    rowDGV.Cells(Me.ColWarehouse_Desc.Index).Value = row_W.Description
                    rowDGV.Cells(Me.ColWarehouse_Level.Index).Value = "Warehouse"
                End If
                If row.include_exclude = Inc_Exc Then
                    rowDGV.Cells(Me.ColWarehouse_ExcludeInclude.Index).Value = "Include"
                Else
                    rowDGV.Cells(Me.ColWarehouse_ExcludeInclude.Index).Value = "Exclude"
                End If
                rowDGV.Cells(Me.ColWarehouse_Remarks.Index).Value = row.remarks
            Next

            Me.dgvWarehouse.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBL.Dispose()
            DTA_W.Dispose()
            TBL_W.Dispose()
            DTA_L.Dispose()
            TBL_L.Dispose()
        End Try
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_ExplainWarehouseExc_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.LoadDGVWarehouse()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class