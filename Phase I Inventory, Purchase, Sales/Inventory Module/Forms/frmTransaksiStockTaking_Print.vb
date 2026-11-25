Option Explicit On
Option Strict Off
Imports Core_Module

Public Class frmTransaksiStockTaking_Print
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim TBLBrand As New Core_Module.dseMasterProductGroup.tblmaster_product_group2DataTable
    Dim rowBrand As Core_Module.dseMasterProductGroup.tblmaster_product_group2Row
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjST As New clsTransaksiStockTaking
    Dim ObjLocation As New clsMasterStorageArea
    Dim TBLLocation As New dseMasterStorageArea.tblmaster_storage_area1DataTable
    Dim rowLocation As dseMasterStorageArea.tblmaster_storage_area1Row
    Dim Buffer As String = ""
    Dim TBLRaw As New dseLaporanStockTaking.tblmaster_productqtyDataTable
    Dim rowRaw As dseLaporanStockTaking.tblmaster_productqtyRow
    Dim TBLBrandRef As New DseLaporanStockTaking_PrintForm2.TBLBrandDataTable
    Dim DTABrandRef As New DseLaporanStockTaking_PrintForm2TableAdapters.TBLBrandTableAdapter
    Dim rowBrandRef As DseLaporanStockTaking_PrintForm2.TBLBrandRow
    Dim DS As New DseLaporanStockTaking_PrintForm2
    Dim TBLCetak As DseLaporanStockTaking_PrintForm2.ReportDataTable = DS.Report
    Dim TBLCetak_Dummy As New DseLaporanStockTaking_PrintForm2.ReportDataTable
    Dim rowCetak, rowCetak1 As DseLaporanStockTaking_PrintForm2.ReportRow
    Dim report As New CrpFormStockTaking

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub frmTransaksiStockTaking_Print_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim OBjBrand As New Core_Module.ClsMasterProductGroup
        Try
            'isikan stock taking
            If ObjST.GetJmlActiveStockTaking > 0 Then
                Me.TxtPrimaryStockTaking.Text = Me.ObjST.GetPrimaryMax
                Me.TxtStockTakingNumber.Text = Me.ObjST.GetNoTransaksi(Me.TxtPrimaryStockTaking.Text)
            Else
                MsgBox("There is no active stock taking, please create stock taking first", MsgBoxStyle.Information, "Information")
                'Me.Close()
                Exit Sub
            End If

            'isikan brand
            OBjBrand.IsiListBoxBrand(Me.clbBrand, Me.TBLBrand)
            Me.TxtPrimaryLocation.Text = "-1"
            Me.TxtPrimaryWarehouse.Text = "-1"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            OBjBrand = Nothing
        End Try
    End Sub

    Private Sub Cetak()
        Dim ReferenceDate As Date = Today
        Dim ProductCodeDummy As String = ""

        Try
            If Me.TxtStockTakingNumber.Text <> "-1" And Me.TxtPrimaryStockTaking.Text <> "" Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor

                'get reference date
                ReferenceDate = Me.ObjST.GetStartDate(Me.TxtPrimaryStockTaking.Text)

                'isikan posisi asal
                Me.ObjST.IsiItemStartForFormPrinting(Me.TBLRaw)

                'isikan brand
                Me.TBLBrandRef.Rows.Clear()
                Me.TBLCetak.Rows.Clear()
                Me.DTABrandRef.ClearBeforeFill = False
                For Each dv As Data.DataRowView In Me.clbBrand.CheckedItems
                    Me.rowBrand = dv.Row
                    Me.DTABrandRef.Fill(Me.TBLBrandRef, Me.rowBrand.Primary)
                Next

                'pindahkan sesuai brand
                For Each Me.rowRaw In Me.TBLRaw.Rows
                    Me.TBLBrandRef.DefaultView.RowFilter = "[Primary Product] = " + Me.rowRaw.Primary_Product.ToString
                    If Me.TBLBrandRef.DefaultView.Count > 0 Then
                        Me.rowBrandRef = Me.TBLBrandRef.DefaultView(0).Row
                        Me.rowCetak = Me.TBLCetak.NewReportRow
                        With rowCetak
                            If Me.rowRaw.Expiry_Date = #1/1/1950# Then
                                .Batch_Number = Me.rowRaw.Lot_No
                            Else
                                .Batch_Number = Me.rowRaw.Lot_No + " - " + Format(Me.rowRaw.Expiry_Date, "dd MMM yyyy")
                            End If
                            .Brand = Me.rowBrandRef.Brand_Code + " - " + Me.rowBrandRef.Brand_Desc
                            .Category = Me.rowBrandRef.Category_Code + " - " + Me.rowBrandRef.Category_Desc
                            .Code_No = Me.rowBrandRef.Product_Code
                            If Me.chkUseSpecification.Checked = False Then
                                .Desc = Me.rowBrandRef.Product_Desc
                            Else
                                If Me.rowBrandRef.Specification <> "" Then
                                    .Desc = Me.rowBrandRef.Specification
                                Else
                                    .Desc = Me.rowBrandRef.Product_Desc
                                End If
                            End If
                            .Expiry = rowRaw.Expiry_Date
                            .Qty_Book = Me.rowRaw.Qty
                            .Unit_Code = Me.rowBrandRef.Unit_Code
                            ._Decimal = Me.rowBrandRef._Decimal
                            .Visible = 1
                        End With
                        Me.TBLCetak.Rows.Add(Me.rowCetak)
                    End If
                Next

                'urutkan berdasarkan brand, category, code no, expiry
                Me.TBLCetak_Dummy.Rows.Clear()
                For Each Me.rowCetak In Me.TBLCetak.Rows
                    Me.rowCetak1 = Me.TBLCetak_Dummy.NewReportRow
                    For ij As Integer = 0 To Me.TBLCetak.Columns.Count - 1
                        Me.rowCetak1.Item(ij) = Me.rowCetak.Item(ij)
                    Next
                    Me.TBLCetak_Dummy.Rows.Add(Me.rowCetak1)
                Next
                Me.TBLCetak.Rows.Clear()
                Me.TBLCetak_Dummy.DefaultView.Sort = "Brand, Category, [Code No], [Batch Number]"
                For Each dv As Data.DataRowView In Me.TBLCetak_Dummy.DefaultView
                    Me.rowCetak1 = dv.Row
                    Me.rowCetak = Me.TBLCetak.NewReportRow
                    For ij As Integer = 0 To Me.TBLCetak.Columns.Count - 1
                        Me.rowCetak.Item(ij) = Me.rowCetak1.Item(ij)
                    Next
                    Me.TBLCetak.Rows.Add(Me.rowCetak)
                    If rowCetak.Code_No <> ProductCodeDummy Then
                        ProductCodeDummy = rowCetak.Code_No
                        rowCetak.Visible = 1
                    Else
                        rowCetak.Visible = 0
                    End If
                Next

                'tampilkan form print
                report.SetDataSource(DS)
                report.SetParameterValue(report.Parameter_Location.ParameterFieldName, Me.cboLocation.Text)
                report.SetParameterValue(report.Parameter_Stock_Taking_Date.ParameterFieldName, Me.ObjST.GetStartDateTransaksiYangMasihAktif)
                report.SetParameterValue(report.Parameter_Stock_Taking_No.ParameterFieldName, Me.TxtStockTakingNumber.Text)
                report.SetParameterValue(report.Parameter_Warehouse.ParameterFieldName, Me.TxtWarehouseName.Text)

                crv.ReportSource = report
                crv.ShowRefreshButton = False
                crv.Show()
            Else
                MsgBox("Please specify stock taking.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub TxtWarehouseCode_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtWarehouseCode.Enter

        Try
            Me.buffer = Me.TxtPrimaryWarehouse.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtWarehouseCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtWarehouseCode.KeyDown
        Dim i As New frmCariMasterWarehouse
        Dim Obj As New clsMasterWarehouse

        Try
            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Obj.ApakahSudahAda(Me.TxtWarehouseCode.Text) = True Then
                    Me.TxtPrimaryWarehouse.Text = Obj.GetPrimary(Me.TxtWarehouseCode.Text)
                    Me.TxtWarehouseName.Text = Obj.GetDescription(Me.TxtPrimaryWarehouse.Text)
                    Me.cboLocation.Focus()
                Else
                    If Me.TxtWarehouseCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtWarehouseCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryWarehouse > 0 Then
                        Me.TxtPrimaryWarehouse.Text = i.PrimaryWarehouse
                        Me.TxtWarehouseCode.Text = Obj.GetWarehouseCode(Me.TxtPrimaryWarehouse.Text)
                        Me.TxtWarehouseName.Text = Obj.GetDescription(Me.TxtPrimaryWarehouse.Text)
                        Me.cboLocation.Focus()
                    Else
                        Me.TxtWarehouseCode.Clear()
                        Me.TxtPrimaryWarehouse.Text = "-1"
                        Me.TxtWarehouseName.Clear()
                        Me.TxtWarehouseCode.Focus()
                    End If
                End If
                Me.IsikanCboGudang()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            Obj = Nothing
        End Try
    End Sub

    Private Sub txtWarehouseCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtWarehouseCode.Leave
        Dim i As New frmCariMasterWarehouse
        Dim obj As New clsMasterWarehouse

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.TxtWarehouseCode.Text <> "" Then
                If obj.ApakahSudahAda(Me.TxtWarehouseCode.Text) = True Then
                    Me.TxtPrimaryWarehouse.Text = obj.GetPrimary(Me.TxtWarehouseCode.Text)
                    Me.TxtWarehouseName.Text = obj.GetDescription(Me.TxtPrimaryWarehouse.Text)
                Else
                    If Me.TxtWarehouseCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtWarehouseCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryWarehouse > 0 Then
                        Me.TxtPrimaryWarehouse.Text = i.PrimaryWarehouse
                        Me.TxtWarehouseCode.Text = obj.GetWarehouseCode(Me.TxtPrimaryWarehouse.Text)
                        Me.TxtWarehouseName.Text = obj.GetDescription(Me.TxtPrimaryWarehouse.Text)
                    Else
                        Me.TxtWarehouseCode.Clear()
                        Me.TxtWarehouseName.Clear()
                        Me.TxtPrimaryWarehouse.Text = "-1"
                        Me.TxtWarehouseCode.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryWarehouse.Text = "-1"
            End If
            If Me.TxtPrimaryWarehouse.Text <> Me.Buffer Then
                Me.IsikanCboGudang()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFindWarehouse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindWarehouse.Click
        Dim i As New frmCariMasterWarehouse
        Dim obj As New clsMasterWarehouse

        Try
            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BackColorAll
            If Me.TxtWarehouseCode.Text <> "" Then
                i.txtCari.Text = Me.TxtWarehouseCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryWarehouse > 0 Then
                Me.TxtPrimaryWarehouse.Text = i.PrimaryWarehouse
                Me.TxtWarehouseCode.Text = obj.GetWarehouseCode(Me.TxtPrimaryWarehouse.Text)
                Me.TxtWarehouseName.Text = obj.GetDescription(Me.TxtPrimaryWarehouse.Text)
                Me.cboLocation.Focus()
            Else
                Me.TxtWarehouseCode.Clear()
                Me.TxtWarehouseName.Clear()
                Me.TxtWarehouseCode.Focus()
            End If
            Me.IsikanCboGudang()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub IsikanCboGudang()
        Try
            If Me.TxtPrimaryWarehouse.Text = "" Or Me.TxtPrimaryWarehouse.Text = "-1" Then
                Me.cboLocation.DataSource = Nothing
                Me.cboLocation.Items.Clear()
            Else
                Me.ObjLocation.IsiTBL(Me.TBLLocation)
                Me.rowLocation = Me.TBLLocation.NewRow
                Me.rowLocation.Primary = -1
                Me.rowLocation.Storage_Area = ""
                Me.TBLLocation.Rows.InsertAt(rowLocation, 0)
                Me.cboLocation.DataSource = Me.TBLLocation.DefaultView
                Me.cboLocation.DisplayMember = "Storage Area"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboLocation_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboLocation.SelectedValueChanged
        Try
            If Me.cboLocation.SelectedIndex = -1 Then
                Me.TxtPrimaryLocation.Text = "-1"
            Else
                Me.rowLocation = Me.TBLLocation.DefaultView(Me.cboLocation.SelectedIndex).Row
                Me.TxtPrimaryLocation.Text = Me.rowLocation.Primary
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        If Me.clbBrand.CheckedItems.Count = 0 Then
            MsgBox("Please specify at least 1 brand.", MsgBoxStyle.Information, "Information")
            Exit Sub
        End If
        Me.Cetak()
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            For i As Integer = 0 To Me.clbBrand.Items.Count - 1
                Me.clbBrand.SetItemCheckState(i, CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Try
            For i As Integer = 0 To Me.clbBrand.Items.Count - 1
                Me.clbBrand.SetItemCheckState(i, CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class