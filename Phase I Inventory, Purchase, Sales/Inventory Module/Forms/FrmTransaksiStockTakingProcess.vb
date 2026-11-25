Option Explicit On
Option Strict Off

Public Class FrmTransaksiStockTakingProcess
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim PrimaryStockTaking As Double = -1
    Dim TBL As New Core_Module.dseMasterProductGroup.tblmaster_product_group2DataTable

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub CetakStockTakingReportDetail1()
        Dim DS As New dseLaporanStockTaking
        Dim report As New crpLaporanStockTakingDetail1
        Dim ObjStockTaking As New clsTransaksiStockTaking
        Dim StockTakingDate As Date = ObjStockTaking.GetStartDateTransaksiYangMasihAktif()
        ObjStockTaking = Nothing
        Dim row As dseLaporanStockTaking.laporan_stock_taking_detail_viewRow
        Dim i As Integer = 0
        Dim BrandCode As String = ""
        Dim PrimaryBrand As Double = -1
        Dim DR As Data.DataRowView
        Dim DTA As New dseLaporanStockTakingTableAdapters.laporan_stock_taking_detail_viewTableAdapter
        Dim TipeItemGoods As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods

        Try
            For Each DR In Me.CheckedListBox1.CheckedItems
                PrimaryBrand = DR.Item("Primary")
                If BrandCode = "" Then
                    BrandCode = DR.Item(1)
                Else
                    BrandCode += ", " + DR.Item(1)
                End If
                DTA.ClearBeforeFill = False
                DTA.FillBy_PrimaryBrand(DS.laporan_stock_taking_detail_view, PrimaryBrand, CInt(TipeItemGoods))
            Next
            While i < DS.laporan_stock_taking_detail_view.Rows.Count
                row = DS.laporan_stock_taking_detail_view.Rows(i)
                If row.Total_On_Hand = 0 And row.Total_Taking = 0 Then
                    DS.laporan_stock_taking_detail_view.Rows.Remove(row)
                Else
                    i += 1
                End If
            End While
            DTA.Dispose()

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_Brand.ParameterFieldName, BrandCode)
            report.SetParameterValue(report.Parameter_FormatDays.ParameterFieldName, Me.txtFormatDay.Text)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, Me.txtFormatMonth.Text)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, Me.txtFormatYears.Text)
            report.SetParameterValue(report.Parameter_StockTakingDate.ParameterFieldName, StockTakingDate)
            Me.crvStockTakingReportDetail.ReportSource = report
            Me.crvStockTakingReportDetail.DisplayStatusBar = True
            Me.crvStockTakingReportDetail.DisplayToolbar = True
            Me.crvStockTakingReportDetail.ShowGroupTreeButton = True
            Me.crvStockTakingReportDetail.ShowRefreshButton = False
            Me.crvStockTakingReportDetail.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
        End Try
    End Sub

    Private Sub CetakStockTakingReport1()
        Dim DS As New dseLaporanStockTaking
        Dim report As New crpLaporanStockTaking1
        Dim ObjStockTaking As New clsTransaksiStockTaking
        Dim StockTakingDate As Date = ObjStockTaking.GetStartDateTransaksiYangMasihAktif()
        ObjStockTaking = Nothing
        Dim row As dseLaporanStockTaking.laporan_stock_taking_viewRow
        Dim i As Integer = 0
        Dim BrandCode As String = ""
        Dim PrimaryBrand As Double = -1
        Dim DR As Data.DataRowView
        Dim DTA As New dseLaporanStockTakingTableAdapters.laporan_stock_taking_viewTableAdapter
        Dim TipeItemGoods As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods

        Try
            For Each DR In Me.CheckedListBox1.CheckedItems
                PrimaryBrand = DR.Item("Primary")
                If BrandCode = "" Then
                    BrandCode = DR.Item(1)
                Else
                    BrandCode += ", " + DR.Item(1)
                End If
                DTA.ClearBeforeFill = False
                DTA.FillBy_PrimaryBrand(DS.laporan_stock_taking_view, PrimaryBrand, CInt(TipeItemGoods))
            Next
            While i < DS.laporan_stock_taking_view.Rows.Count
                row = DS.laporan_stock_taking_view.Rows(i)
                If row.Total_Qty_On_Hand = 0 And row.Total_Qty_Taking = 0 Then
                    DS.laporan_stock_taking_view.Rows.Remove(row)
                Else
                    i += 1
                End If
            End While
            DTA.Dispose()

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_Brand.ParameterFieldName, BrandCode)
            report.SetParameterValue(report.Parameter_FormatDays.ParameterFieldName, Me.txtFormatDay.Text)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, Me.txtFormatMonth.Text)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, Me.txtFormatYears.Text)
            report.SetParameterValue(report.Parameter_StockTakingDate.ParameterFieldName, StockTakingDate)
            Me.crvStockTakingReport.ReportSource = report
            Me.crvStockTakingReport.DisplayStatusBar = True
            Me.crvStockTakingReport.DisplayToolbar = True
            Me.crvStockTakingReport.ShowGroupTreeButton = True
            Me.crvStockTakingReport.ShowRefreshButton = False
            Me.crvStockTakingReport.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
        End Try
    End Sub

    Private Sub CetakExpiryDateDataCheck()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim DTA As New DseLaporanStockTaking_ExpiryDataCheckTableAdapters.laporan_stock_taking_view_expirydate3_viewSumberTableAdapter
        Dim DTATbl3 As New DseLaporanStockTaking_ExpiryDataCheckTableAdapters.DataTable3TableAdapter
        Dim DS As New DseLaporanStockTaking_ExpiryDataCheck
        Dim report As New crpLaporanStockTaking_ED_DataCheck
        Dim ListTag As New List(Of String)
        Dim TipeItem As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods

        Try
            DTA.Fill(DS.laporan_stock_taking_view_expirydate3_viewSumber, Me.PrimaryStockTaking, TipeItem)
            DTA.Dispose()
            DTATbl3.Fill(DS.DataTable3, Me.PrimaryStockTaking)
            DTATbl3.Dispose()

            Dim rowLaporanSumber As DseLaporanStockTaking_ExpiryDataCheck.laporan_stock_taking_view_expirydate3_viewSumberRow
            Dim rowTbl3 As DseLaporanStockTaking_ExpiryDataCheck.DataTable3Row

            'Delete yang sudah benar
            For i As Integer = 0 To DS.laporan_stock_taking_view_expirydate3_viewSumber.Rows.Count - 1
                If i > DS.laporan_stock_taking_view_expirydate3_viewSumber.Rows.Count - 1 Then
                    Exit For
                End If
                If DS.laporan_stock_taking_view_expirydate3_viewSumber.Rows(i).Item("Expiry Date Taking") = _
                DS.laporan_stock_taking_view_expirydate3_viewSumber.Rows(i).Item("Expiry Date") Then
                    DS.laporan_stock_taking_view_expirydate3_viewSumber.Rows.RemoveAt(i)
                    i -= 1
                End If
            Next

            'tandai tag number asal
            For Each rowLaporanSumber In DS.laporan_stock_taking_view_expirydate3_viewSumber.Rows
                For Each rowTbl3 In DS.DataTable3.Rows
                    If rowLaporanSumber.Primary_Product = rowTbl3.Primary_Product And _
                    rowLaporanSumber.Lot_No = rowTbl3.Lot_No And rowLaporanSumber.Expiry_Date_Taking = _
                    rowTbl3.Expiry_Date Then
                        If rowLaporanSumber.IsTag_NoNull = True Then
                            rowLaporanSumber.Tag_No = rowTbl3.Tag_No
                        Else
                            If rowLaporanSumber.Tag_No = "" Then
                                rowLaporanSumber.Tag_No = rowTbl3.Tag_No
                            Else
                                rowLaporanSumber.Tag_No = rowLaporanSumber.Tag_No + ", " + rowTbl3.Tag_No
                            End If
                        End If
                    End If
                Next

                If rowLaporanSumber.Tag_No = "" Then
                    MsgBox("No Tag-> CodeNo: " + rowLaporanSumber.Product_Code + "; BatchNo: " + _
                    rowLaporanSumber.Lot_No)
                End If
            Next

            Dim List As New List(Of clsLaporanStockTakingCheckList)
            Dim obj As clsLaporanStockTakingCheckList
            Dim Ketemu As Boolean = False

            'Bikin Group
            For Each rowLaporanSumber In DS.laporan_stock_taking_view_expirydate3_viewSumber
                Ketemu = False
                If List.Count = 0 Then
                    obj = New clsLaporanStockTakingCheckList
                    obj.PrimaryProduct = rowLaporanSumber.Primary_Product
                    obj.Description = rowLaporanSumber.Description
                    obj.ExpiryDate = rowLaporanSumber.Expiry_Date
                    obj.LotNO = rowLaporanSumber.Lot_No
                    obj.QtyONHand = rowLaporanSumber.Qty_On_Hand
                    obj.QtyTaking = rowLaporanSumber.Qty_Taking
                    obj.UnitCode = rowLaporanSumber.Unit_Code
                    rowLaporanSumber.Group_No = rowLaporanSumber.Product_Code + " - " + rowLaporanSumber.Lot_No
                    obj.Group = rowLaporanSumber.Product_Code + " - " + rowLaporanSumber.Lot_No
                    List.Add(obj)
                Else
                    For Each obj In List
                        If obj.PrimaryProduct = rowLaporanSumber.Primary_Product And obj.Description = rowLaporanSumber.Description And _
                        obj.ExpiryDate = rowLaporanSumber.Expiry_Date And obj.LotNO = rowLaporanSumber.Lot_No And _
                        obj.QtyONHand = rowLaporanSumber.Qty_On_Hand And obj.QtyTaking = rowLaporanSumber.Qty_Taking And _
                        obj.UnitCode = rowLaporanSumber.Unit_Code Then
                            rowLaporanSumber.Group_No = obj.Group
                            Ketemu = True
                            Exit For
                        End If
                    Next
                    If Ketemu = False Then
                        obj = New clsLaporanStockTakingCheckList
                        obj.PrimaryProduct = rowLaporanSumber.Primary_Product
                        obj.Description = rowLaporanSumber.Description
                        obj.ExpiryDate = rowLaporanSumber.Expiry_Date
                        obj.LotNO = rowLaporanSumber.Lot_No
                        obj.QtyONHand = rowLaporanSumber.Qty_On_Hand
                        obj.QtyTaking = rowLaporanSumber.Qty_Taking
                        obj.UnitCode = rowLaporanSumber.Unit_Code
                        rowLaporanSumber.Group_No = rowLaporanSumber.Product_Code + " - " + rowLaporanSumber.Lot_No
                        obj.Group = rowLaporanSumber.Product_Code + " - " + rowLaporanSumber.Lot_No
                        List.Add(obj)
                    End If
                End If
            Next
            List = Nothing

            Dim rowLaporan As DseLaporanStockTaking_ExpiryDataCheck.laporan_stock_taking_view_expirydate3_viewRow
            For Each rowLaporanSumber In DS.laporan_stock_taking_view_expirydate3_viewSumber
                If DS.laporan_stock_taking_view_expirydate3_view.Rows.Count = 0 Then
                    Dim TagNo As String = ""
                    If rowLaporanSumber.IsTag_NoNull = False Then
                        TagNo = rowLaporanSumber.Tag_No
                    End If
                    DS.laporan_stock_taking_view_expirydate3_view.Addlaporan_stock_taking_view_expirydate3_viewRow(rowLaporanSumber.Primary_Product, _
                    rowLaporanSumber.Product_Code, rowLaporanSumber.Description, rowLaporanSumber.Unit_Code, rowLaporanSumber.Lot_No, _
                    rowLaporanSumber.Qty_Taking, rowLaporanSumber.Qty_On_Hand, rowLaporanSumber.Expiry_Date_Taking, rowLaporanSumber.Expiry_Date, _
                    TagNo, rowLaporanSumber.Group_No, rowLaporanSumber._Decimal)
                Else
                    Ketemu = False
                    For Each rowLaporan In DS.laporan_stock_taking_view_expirydate3_view
                        If rowLaporanSumber.Primary_Product = rowLaporan.Primary_Product And _
                        rowLaporanSumber.Description = rowLaporan.Description And _
                        rowLaporanSumber.Unit_Code = rowLaporan.Unit_Code And _
                        rowLaporanSumber.Qty_On_Hand = rowLaporan.Qty_Taking And _
                        rowLaporanSumber.Lot_No = rowLaporan.Lot_No And _
                        rowLaporanSumber.Expiry_Date = rowLaporan.Expiry_Date And _
                        rowLaporanSumber.Expiry_Date_Taking = rowLaporan.Expiry_Date_Taking Then
                            If rowLaporan.Tag_No <> rowLaporanSumber.Tag_No Then
                                rowLaporan.Tag_No += ", " + rowLaporanSumber.Tag_No
                            End If
                            Ketemu = True
                            Exit For
                        End If
                    Next
                    If Ketemu = False Then
                        DS.laporan_stock_taking_view_expirydate3_view.Addlaporan_stock_taking_view_expirydate3_viewRow(rowLaporanSumber.Primary_Product, _
                        rowLaporanSumber.Product_Code, rowLaporanSumber.Description, rowLaporanSumber.Unit_Code, rowLaporanSumber.Lot_No, _
                        rowLaporanSumber.Qty_Taking, rowLaporanSumber.Qty_On_Hand, rowLaporanSumber.Expiry_Date_Taking, rowLaporanSumber.Expiry_Date, _
                        rowLaporanSumber.Tag_No, rowLaporanSumber.Group_No, rowLaporanSumber._Decimal)
                    End If
                End If
            Next

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDays.ParameterFieldName, 1)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, 3)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, 1)
            Me.crvExpiryDataCheck.ReportSource = report
            Me.crvExpiryDataCheck.DisplayStatusBar = True
            Me.crvExpiryDataCheck.DisplayToolbar = True
            Me.crvExpiryDataCheck.ShowGroupTreeButton = True
            Me.crvExpiryDataCheck.ShowRefreshButton = False
            Me.crvExpiryDataCheck.Show()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
        End Try
    End Sub

    Private Sub CetakQtyDataCheck()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim DS As New DseLaporanStockTaking_LocationCheck
        Dim DTA As New DseLaporanStockTaking_LocationCheckTableAdapters.laporan_stock_taking_view_location1_viewTableAdapter
        Dim DTAView2 As New DseLaporanStockTaking_LocationCheckTableAdapters.laporan_stock_taking_view_location2_viewTableAdapter
        Dim report As New crpLaporanStockTaking_Qty_DataCheck
        Dim filter As String = ""
        Dim DataView As Data.DataRowView
        Dim TipeItem As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods

        Try
            DTA.Fill(DS.laporan_stock_taking_view_location1_view, TipeItem)
            DTA.Dispose()
            DTAView2.Fill(DS.laporan_stock_taking_view_location2_view, Me.PrimaryStockTaking, TipeItem)
            DTAView2.Dispose()

            Dim rowLaporanSumber As DseLaporanStockTaking_LocationCheck.laporan_stock_taking_view_location1_viewRow
            Dim rowView2 As DseLaporanStockTaking_LocationCheck.laporan_stock_taking_view_location2_viewRow

            For Each rowLaporanSumber In DS.laporan_stock_taking_view_location1_view.Rows
                For Each rowView2 In DS.laporan_stock_taking_view_location2_view.Rows
                    If rowLaporanSumber.Primary_Storage_Area = -1 Then
                        If rowLaporanSumber.Primary_Product = rowView2.Primary_Product And _
                        rowLaporanSumber.Lot_No = rowView2.Lot_No Then
                            If rowLaporanSumber.IsTag_NoNull = True Then
                                rowLaporanSumber.Tag_No = rowView2.Tag_No
                            Else
                                If rowLaporanSumber.Tag_No = "" Then
                                    rowLaporanSumber.Tag_No = rowView2.Tag_No
                                Else
                                    rowLaporanSumber.Tag_No = rowLaporanSumber.Tag_No + ", " + rowView2.Tag_No
                                End If
                            End If
                        End If
                    Else
                        If rowLaporanSumber.Primary_Product = rowView2.Primary_Product And _
                        rowLaporanSumber.Primary_Storage_Area = rowView2.Primary_Storage_Area And _
                        rowLaporanSumber.Lot_No = rowView2.Lot_No Then
                            If rowLaporanSumber.IsTag_NoNull = True Then
                                rowLaporanSumber.Tag_No = rowView2.Tag_No
                            Else
                                If rowLaporanSumber.Tag_No = "" Then
                                    rowLaporanSumber.Tag_No = rowView2.Tag_No
                                Else
                                    rowLaporanSumber.Tag_No = rowLaporanSumber.Tag_No + ", " + rowView2.Tag_No
                                End If
                            End If
                        End If
                    End If

                    If rowLaporanSumber.IsTag_NoNull = False Then
                        If rowLaporanSumber.Tag_No = "" Then
                            MsgBox("")
                        End If
                    End If
                Next
            Next

            Dim List As New List(Of clsLaporanStockTakingCheckList)
            Dim obj As clsLaporanStockTakingCheckList
            Dim Ketemu As Boolean = False

            'Bikin Group
            For Each rowLaporanSumber In DS.laporan_stock_taking_view_location1_view
                Ketemu = False
                If List.Count = 0 Then
                    obj = New clsLaporanStockTakingCheckList
                    obj.PrimaryProduct = rowLaporanSumber.Primary_Product
                    obj.Description = rowLaporanSumber.Description
                    obj.LotNO = rowLaporanSumber.Lot_No
                    obj.QtyONHand = rowLaporanSumber.Total_Qty_On_Hand
                    obj.QtyTaking = rowLaporanSumber.Total_Qty_Taking
                    obj.UnitCode = rowLaporanSumber.Unit_Code
                    obj.Difference = rowLaporanSumber.TotalDifference
                    If rowLaporanSumber.Lot_No <> "" Then
                        rowLaporanSumber.Group_No = rowLaporanSumber.Product_Code + " - " + rowLaporanSumber.Lot_No
                        obj.Group = rowLaporanSumber.Product_Code + " - " + rowLaporanSumber.Lot_No
                    Else
                        rowLaporanSumber.Group_No = rowLaporanSumber.Product_Code
                        obj.Group = rowLaporanSumber.Product_Code
                    End If
                    List.Add(obj)
                Else
                    For Each obj In List
                        If obj.PrimaryProduct = rowLaporanSumber.Primary_Product And obj.Description = rowLaporanSumber.Description And _
                        obj.LotNO = rowLaporanSumber.Lot_No And obj.Difference = rowLaporanSumber.TotalDifference And _
                        obj.QtyONHand = rowLaporanSumber.Total_Qty_On_Hand And obj.QtyTaking = rowLaporanSumber.Total_Qty_Taking And _
                        obj.UnitCode = rowLaporanSumber.Unit_Code Then
                            rowLaporanSumber.Group_No = obj.Group
                            Ketemu = True
                            Exit For
                        End If
                    Next
                    If Ketemu = False Then
                        obj = New clsLaporanStockTakingCheckList
                        obj.PrimaryProduct = rowLaporanSumber.Primary_Product
                        obj.Description = rowLaporanSumber.Description
                        obj.LotNO = rowLaporanSumber.Lot_No
                        obj.QtyONHand = rowLaporanSumber.Total_Qty_On_Hand
                        obj.QtyTaking = rowLaporanSumber.Total_Qty_Taking
                        obj.UnitCode = rowLaporanSumber.Unit_Code
                        obj.Difference = rowLaporanSumber.TotalDifference
                        If rowLaporanSumber.Lot_No = "" Then
                            rowLaporanSumber.Group_No = rowLaporanSumber.Product_Code
                            obj.Group = rowLaporanSumber.Product_Code
                        Else
                            rowLaporanSumber.Group_No = rowLaporanSumber.Product_Code + " - " + rowLaporanSumber.Lot_No
                            obj.Group = rowLaporanSumber.Product_Code + " - " + rowLaporanSumber.Lot_No
                        End If
                        List.Add(obj)
                    End If
                End If
            Next

            'untuk setiap group, filter kembali yang ditampilkan hanya yang salah satu barisnya
            'ada selisih antara qty taking dan qty on hand
            For Each obj In List
                filter = "[Group No] = '" + obj.Group + "'"
                DS.laporan_stock_taking_view_location1_view.DefaultView.RowFilter = filter
                Ketemu = False

                For Each DataView In DS.laporan_stock_taking_view_location1_view.DefaultView
                    rowLaporanSumber = DataView.Row
                    If rowLaporanSumber.Qty_on_Hand_Per_Lokasi <> rowLaporanSumber.Qty_Taking_Per_Lokasi Then
                        'tandai bahwa group in boleh dipertahankan
                        Ketemu = True
                        Exit For
                    End If
                Next

                If Ketemu = False Then
                    'jika ternyata setiap baris pada laporan qty on hand = qty taking
                    'hapus semuanya
                    While DS.laporan_stock_taking_view_location1_view.DefaultView.Count > 0
                        DS.laporan_stock_taking_view_location1_view.Rows.Remove(DS.laporan_stock_taking_view_location1_view.DefaultView(0).Row)
                    End While
                End If
            Next
            List = Nothing

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDays.ParameterFieldName, 1)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, 3)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, 1)
            Me.crvQtyDataCheck.ReportSource = report
            Me.crvQtyDataCheck.DisplayStatusBar = True
            Me.crvQtyDataCheck.DisplayToolbar = True
            Me.crvQtyDataCheck.ShowGroupTreeButton = True
            Me.crvQtyDataCheck.ShowRefreshButton = False
            Me.crvQtyDataCheck.Show()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
        End Try
    End Sub

    Private Sub FrmTransaksiStockTakingProcess_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim obj As New Core_Module.ClsMasterProductGroup
            obj.IsiListBoxBrand(Me.CheckedListBox1, TBL)
            obj = Nothing
            Dim objTaking As New clsTransaksiStockTaking
            Me.PrimaryStockTaking = objTaking.GetPrimaryMax()
            objTaking = Nothing
            For i As Integer = 0 To Me.CheckedListBox1.Items.Count - 1
                Me.CheckedListBox1.SetItemCheckState(i, CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadData()
        Dim ObjStockTaking As New clsTransaksiStockTaking

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            ObjStockTaking.IsiItemBuffer()
            Me.CetakStockTakingReportDetail1()
            Me.CetakStockTakingReport1()
            Me.CetakExpiryDateDataCheck()
            Me.CetakQtyDataCheck()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjStockTaking = Nothing
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            For i As Integer = 0 To Me.CheckedListBox1.Items.Count - 1
                Me.CheckedListBox1.SetItemCheckState(i, CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Try
            For i As Integer = 0 To Me.CheckedListBox1.Items.Count - 1
                Me.CheckedListBox1.SetItemCheckState(i, CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub llbProcessingData_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs)

    End Sub

    Private Sub IsikanTabelAcuan(ByVal vTBLAcuan As dseTransaksiStockTaking.tbltransaksi_stockinout_itemDataTable)
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim TBL1 As New dseTransaksiStockTaking.tbltransaksi_stockinout_itemDataTable
        Dim row, row1 As dseTransaksiStockTaking.tbltransaksi_stockinout_itemRow
        Dim DTA2 As New dseTransaksiStockTakingTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
        Dim TBL2 As New dseTransaksiStockTaking.tbltransaksi_stockbundling_item_afterDataTable
        Dim row2 As dseTransaksiStockTaking.tbltransaksi_stockbundling_item_afterRow
        Dim TipeStockIn As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
        Dim Filter As String = ""

        Try
            'isikan dengan stock in purchase
            DTA.Fill(vTBLAcuan, TipeStockIn, #1/1/1950#)
            'isikan dengan stock in adjust
            TipeStockIn = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
            DTA.Fill(TBL1, TipeStockIn, #1/1/1950#)
            For Each row In TBL1.Rows
                Filter = "[Primary Product] = " + row.Primary_Product.ToString + " AND [Batch No] = '" + row.Batch_No + "'"
                vTBLAcuan.DefaultView.RowFilter = Filter
                If vTBLAcuan.DefaultView.Count = 0 Then
                    'tambahkan baris baru
                    row1 = vTBLAcuan.Newtbltransaksi_stockinout_itemRow
                    For i As Integer = 0 To vTBLAcuan.Columns.Count - 1
                        row1.Item(i) = row.Item(i)
                    Next
                    vTBLAcuan.Rows.Add(row1)
                End If
            Next

            'isikan dengan bundling after
            DTA2.Fill(TBL2, #1/1/1950#)
            For Each row2 In TBL2.Rows
                Filter = "[Primary Product] = " + row2.Primary_Product.ToString + " AND [Batch No] = '" + row2.Batch_No + "'"
                vTBLAcuan.DefaultView.RowFilter = Filter
                If vTBLAcuan.DefaultView.Count = 0 Then
                    'tambahkan baris baru
                    row1 = vTBLAcuan.Newtbltransaksi_stockinout_itemRow
                    For i As Integer = 0 To vTBLAcuan.Columns.Count - 1
                        row1.Item(i) = row2.Item(i)
                    Next
                    vTBLAcuan.Rows.Add(row1)
                End If
            Next

            'jika ada transaksi lain yang mengakibatkan penambahkan batch number baru dan expiry baru ke dalam database product bisa ditambahkan disini
            'msgbox("belum")

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL1.Dispose()
            DTA2.Dispose()
            TBL2.Dispose()
        End Try
    End Sub

    Private Sub llbRefreshReport_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs)

    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnStartProses_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStartProses.Click
        Dim Obj As New clsTransaksiStockTaking
        Dim PrimaryStockTaking As Double = Obj.GetPrimaryMax
        Dim StockTakingDate As Date = Obj.GetStartDate(PrimaryStockTaking)
        Dim StockTakingNumber As String = Obj.GetNoTransaksi(PrimaryStockTaking)

        Dim DS As New dseLaporanStockTaking
        Dim DTACheck1 As New dseLaporanStockTakingTableAdapters.tbltrans_stock_taking_tag_itemTableAdapter
        Dim TBLCheck1 As dseLaporanStockTaking.tbltrans_stock_taking_tag_itemDataTable = _
        DS.tbltrans_stock_taking_tag_item
        Dim Filter As String = ""
        Dim rowChk As dseLaporanStockTaking.tbltrans_stock_taking_tag_itemRow
        Dim rowChk2 As dseLaporanStockTaking.laporan_stock_taking_view1Row
        Dim RowView As Data.DataRowView

        Dim TBLBuffer As dseLaporanStockTaking.tbltrans_stock_taking_itembufferDataTable = DS.tbltrans_stock_taking_itembuffer
        Dim DTABuffer As New dseLaporanStockTakingTableAdapters.tbltrans_stock_taking_itembufferTableAdapter
        Dim rowBuffer As dseLaporanStockTaking.tbltrans_stock_taking_itembufferRow
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim DTAProductQty As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        Dim ObjSC As New ClsTransaksiStockCard
        Dim DTASC As New dseTransaksiStockCardTableAdapters.tbltransaksi_stock_cardTableAdapter
        Dim StockAwal, StockAkhir, QtyIn, QtyOut As Double

        Dim DTAPenyesuaian As New dseLaporanStockTakingTableAdapters.tbltrans_stock_taking_itempenyesuaianTableAdapter

        Dim DTAPenyesuaian1 As New dseLaporanStockTakingTableAdapters.tbltrans_stock_taking_itempenyesuaian1TableAdapter
        Dim TBLPenyesuaian1 As dseLaporanStockTaking.tbltrans_stock_taking_itempenyesuaian1DataTable = _
        DS.tbltrans_stock_taking_itempenyesuaian1
        Dim rowPenyesuaian1 As dseLaporanStockTaking.tbltrans_stock_taking_itempenyesuaian1Row

        Dim DTAFinalCheck As New dseLaporanStockTakingTableAdapters.Final_CheckTableAdapter
        Dim TBL_FinalCheck As New dseLaporanStockTaking.Final_CheckDataTable
        Dim row_FinalCheck As dseLaporanStockTaking.Final_CheckRow

        Dim TBLDataCheck As dseLaporanStockTaking.Data_CheckDataTable = DS.Data_Check
        Dim rowCheck As dseLaporanStockTaking.Data_CheckRow

        Dim DTA_ED As New dseTransaksiStockTakingTableAdapters.DataTable2TableAdapter
        Dim TBL_ED As New dseTransaksiStockTaking.DataTable2DataTable
        Dim row_ED As dseTransaksiStockTaking.DataTable2Row
        Dim TBLAcuan_ED As New dseTransaksiStockTaking.tbltransaksi_stockinout_itemDataTable
        Dim Lanjut As Boolean = False
        Dim rowAcuan_ED As dseTransaksiStockTaking.tbltransaksi_stockinout_itemRow
        Dim ObjClosing As New ClsTransaksiMonthlyClosingItem
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            'cek apakah proses stock taking boleh dilakukan
            'stock taking tidak boleh dilakukan bila:
            '0. jika stock takingnya sudah diproses ya tidak bisa lagi dilakukan
            If Obj.GetStatus(PrimaryStockTaking) = clsTransaksiStockTaking.EnumStatusStockTaking.Processed Then
                MsgBox("This stock taking have been processed.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            '1. Pada data stock tag ada 2 atau lebih expiry date yang berbeda untuk batch number yang sama
            DTACheck1.Fill(TBLCheck1, PrimaryStockTaking)
            If TBLCheck1.Rows.Count > 0 Then
                Dim frm As New FrmLaporanStockTaking_LastEDCheck
                frm.DS = DS

                Dim DTACheck2 As New dseLaporanStockTakingTableAdapters.laporan_stock_taking_view1TableAdapter
                Dim TBLCheck2 As dseLaporanStockTaking.laporan_stock_taking_view1DataTable = _
                DS.laporan_stock_taking_view1
                Dim TBLCont As dseLaporanStockTaking.ED_Cont_TableDataTable = _
                DS.ED_Cont_Table
                Dim rowCont As dseLaporanStockTaking.ED_Cont_TableRow

                DTACheck2.Fill(TBLCheck2, PrimaryStockTaking)
                For Each rowChk In TBLCheck1.Rows
                    'filter
                    Filter = "[Primary Product] = " + rowChk.Primary_Product.ToString + _
                    " AND [Batch No] = '" + rowChk.Batch_No + "'"

                    TBLCheck2.DefaultView.RowFilter = Filter

                    For Each RowView In TBLCheck2.DefaultView
                        rowChk2 = RowView.Row
                        'tambahkan baris baru
                        rowCont = TBLCont.NewED_Cont_TableRow
                        With rowChk2
                            rowCont.Tag_Number = .Tag_No
                            rowCont.Product_Code = .Product_Code
                            rowCont.Expiry_Date = Format(.Expiry_Date, "dd MMM yyyy")
                            rowCont.Batch_Number = .Batch_No
                            rowCont.Group = .Product_Code + " - " + .Batch_No
                            rowCont.Description = .Description
                            rowCont.Units = .Unit_Code
                        End With

                        TBLCont.Rows.Add(rowCont)
                    Next
                Next

                DTACheck2.Dispose()

                MsgBox("The following products have batch number with several expiry date." + vbCrLf _
                + "One batch number must refer to one expiry date. Please fix this, before proceeding", _
                MsgBoxStyle.Information, "Information")

                frm.ShowDialog()
                frm = Nothing
                Exit Sub
            End If

            '2. Jika perubahan/koreksi data akan mengakibatkan stock menjadi minus
            Obj.IsiItemBuffer()
            DTAFinalCheck.Fill(TBL_FinalCheck)
            TBLDataCheck.Rows.Clear()
            For Each row_FinalCheck In TBL_FinalCheck.Rows
                With row_FinalCheck
                    StockAwal = ObjProduct.GetJml2(.Primary_Product, DTAProductQty, _
                    .Qty_Secondary, .Primary_Storage_Area, .Lot_No)
                    StockAkhir = StockAwal + .Difference
                    If StockAkhir < 0 Then
                        rowCheck = TBLDataCheck.NewData_CheckRow

                        rowCheck.Differrence = .Difference
                        rowCheck.Location = .Location
                        rowCheck.Lot_Number = .Lot_No
                        rowCheck.Product_Code = .Product_Code
                        rowCheck.Qty_On_Hand = .Qty_On_Hand
                        rowCheck.Qty_On_Hand_Today = StockAwal
                        rowCheck.Qty_Taking = .Qty_Taking

                        TBLDataCheck.Rows.Add(rowCheck)
                    End If
                End With
            Next
            If TBLDataCheck.Rows.Count <> 0 Then
                MsgBox("The following item will be minus if the stock taking process continue, please fix following data.", _
                MsgBoxStyle.Information, "Information")

                Dim Frm_Chk As New FrmLaporanStockTaking_FinalCheck
                Frm_Chk.DS = DS
                Frm_Chk.ShowDialog()
                Frm_Chk = Nothing

                Exit Sub
            End If

            '3. Jika ada product yang taking qty-nya ada, kemudian batch numbernya sama dengan 
            'batch number yang ada, tapi expiry datenya berbeda-load dulu datanya, diperbaiki kemudian
            DTA_ED.Fill(TBL_ED, PrimaryStockTaking, #1/1/1950#)

            '4. jika bulan stock taking dilakukan masih ada transaksi closing yang belum dibuka
            If ObjClosing.GetLastClosingPeriod >= ObjCD.GetFirstDateofTheMonth(StockTakingDate) Then
                MsgBox("There is inventory closing for: " + Format(StockTakingDate, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            '4. msgbox("belum")--> cari tahu apa lagi yang mengakibatkan tidak boleh
            '   melaksanakan proses stock taking

            'konfirmasi ke user untuk mulai proses stock taking
            If MsgBox("Start processing stock taking data?" + vbCrLf + _
            "We recomend you to back up your database first and please make sure nobody else is using this software.", _
            MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                'kunci semua transaksi
                'msgbox("belum")
                'pastikan tidak ada transaksi yang sedang mengantri untuk simpan data
                'msgbox("belum")
                'back up database secara otomatis
                'start proses stock taking
                '0. hitung ulang item buffer

                '1. Load item buffer, yang di loading hanya yang differencenya <> 0
                DTABuffer.Fill_ForProcessing(TBLBuffer)

                '2. Untuk setiap item buffer
                For Each rowBuffer In TBLBuffer.Rows
                    With rowBuffer
                        '0. update expiry sesuai dengan stock taking
                        DTAProductQty.UpdateQuery_SetED(.Expiry_Date, .Primary_Product, .Lot_No)
                        'a. Simpan item perubahan
                        DTAPenyesuaian.Insert(PrimaryStockTaking, .Primary_Product, .Lot_No, _
                        .Expiry_Date, .Primary_Storage_Area, .Difference, .Qty_Secondary)
                        'b. Koreksi stock + Simpan ke dalam stock card
                        'koreksi data stock
                        StockAwal = ObjProduct.GetJml2(.Primary_Product, DTAProductQty, _
                        .Qty_Secondary, .Primary_Storage_Area, .Lot_No)

                        StockAkhir = StockAwal + .Difference
                        If StockAwal <> 0 Then
                            If StockAkhir = 0 Then
                                'hapus data
                                ObjProduct.Hapus(.Primary_Product, DTAProductQty, .Qty_Secondary, .Primary_Storage_Area, _
                                .Lot_No)
                            Else
                                'set jumlah
                                ObjProduct.SetJml(StockAkhir, .Primary_Product, DTAProductQty, .Qty_Secondary, _
                                .Primary_Storage_Area, .Lot_No)
                            End If
                        ElseIf StockAkhir <> 0 Then
                            'simpan data baru
                            ObjProduct.SimpanQtyBaru(StockAkhir, .Primary_Product, DTAProductQty, .Qty_Secondary, _
                            .Primary_Storage_Area, .Lot_No, .Expiry_Date)
                        End If
                    End With
                Next

                'simpan data ke stock card
                DTAPenyesuaian1.FillBy_PrimaryStockTaking(TBLPenyesuaian1, PrimaryStockTaking)
                For Each rowPenyesuaian1 In TBLPenyesuaian1.Rows
                    With rowPenyesuaian1
                        If .Qty <> 0 Then
                            If .Qty < 0 Then
                                QtyIn = 0
                                QtyOut = Math.Abs(.Qty)
                            Else
                                QtyIn = .Qty
                                QtyOut = 0
                            End If
                            ObjSC.Simpan(PrimaryStockTaking, ClsTransaksiStockCard.EnumJenisTransaksiAsal.Stock_Taking, _
                            StockTakingNumber, .Primary_Product, QtyIn, QtyOut, "Stock Taking Adjustment", _
                            ClsTransaksiStockCard.EnumJenisStockCard.Stock_Taking, StockTakingDate, DTASC, .Batch_No, .Expiry_Date, .Primary_Location, "")
                        End If
                    End With
                Next

                'updata data batch number untuk yang hanya berubah expiry datenya saja:
                If TBL_ED.Rows.Count > 0 Then
                    Dim DTA_InOut As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_itemTableAdapter
                    Dim DTA_StockTransfer As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransfer_itemTableAdapter
                    Dim DTA_ItemBefore As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_beforeTableAdapter
                    Dim DTA_ItemAfter As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
                    Dim DTA_Booking As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_invitemTableAdapter
                    Dim DTA_DO As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter
                    Dim DTA_GRN As New DseTransaksiGRNTableAdapters.tbltrans_grn_itemTableAdapter
                    Dim DTA_SR As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_itemTableAdapter
                    Dim DTA_Invoice As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemTableAdapter
                    Dim DTA_Inv_Obs As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoiceobselete_itemTableAdapter
                    Dim DTA_PR As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_itemTableAdapter

                    For Each row_ED In TBL_ED.Rows
                        With row_ED
                            Lanjut = True

                            Filter = "[Primary Product] = " + .Primary_Product.ToString + " AND [Batch No] = '" + .Lot_No + "'"
                            TBLAcuan_ED.DefaultView.RowFilter = Filter
                            If TBLAcuan_ED.DefaultView.Count > 0 Then
                                rowAcuan_ED = TBLAcuan_ED.DefaultView(0).Row
                                If .Expiry_Date_Taking.Date <> rowAcuan_ED.Expiry_Date.Date Then
                                    Lanjut = True
                                Else
                                    Lanjut = False
                                End If
                            Else
                                Lanjut = True
                            End If

                            If Lanjut = True Then
                                '1. Product qty
                                DTAProductQty.UpdateQuery_SetED(row_ED.Expiry_Date_Taking, .Primary_Product, .Lot_No)
                                '2. Stock In Out
                                DTA_InOut.UpdateQuery_SetED(.Expiry_Date_Taking, .Primary_Product, .Lot_No)
                                '3. Stock Transfer
                                DTA_StockTransfer.UpdateQuery_SetED(.Expiry_Date_Taking, .Primary_Product, .Lot_No)
                                '4. Stock Bundling Item B4 + After
                                DTA_ItemBefore.UpdateQuery_SetED(.Expiry_Date_Taking, .Primary_Product, .Lot_No)
                                DTA_ItemAfter.UpdateQuery_SetED(.Expiry_Date_Taking, .Primary_Product, .Lot_No)
                                '5. Booking inv
                                DTA_Booking.UpdateQuery_SetED(.Expiry_Date_Taking, .Primary_Product, .Lot_No)
                                '6. DO
                                DTA_DO.UpdateQuery_SetED(.Expiry_Date_Taking, .Primary_Product, .Lot_No)
                                '7. GRN
                                DTA_GRN.UpdateQuery_SetED(.Expiry_Date_Taking, .Primary_Product, .Lot_No)
                                '8. Sales Return Item
                                DTA_SR.UpdateQuery_SetED(.Expiry_Date_Taking, .Primary_Product, .Lot_No)
                                'jika ada modifikasi penambahan batch number pada sales invoice, aop, dan stock card, jangan lupa juga update 
                                'data2 disana --> update ke smess
                                '9. Purchase return item
                                DTA_PR.UpdateQuery_SetExpiry(.Expiry_Date_Taking, .Primary_Product, .Lot_No)
                                '10. sales invoice item
                                DTA_Invoice.Update(.Expiry_Date_Taking, .Primary_Product, .Lot_No)
                                '11. obselete invoice
                                DTA_Inv_Obs.Update(.Expiry_Date_Taking, .Primary_Product, .Lot_No)
                            End If
                        End With
                    Next

                    DTA_InOut.Dispose()
                    DTA_StockTransfer.Dispose()
                    DTA_ItemBefore.Dispose()
                    DTA_ItemAfter.Dispose()
                    DTA_Booking.Dispose()
                    DTA_DO.Dispose()
                    DTA_GRN.Dispose()
                    DTA_SR.Dispose()
                    DTA_Invoice.Dispose()
                    DTA_Inv_Obs.Dispose()
                    DTA_PR.Dispose()
                End If

                'set status stock taking
                Obj.SetStatus(PrimaryStockTaking, clsTransaksiStockTaking.EnumStatusStockTaking.Processed)

                'simpan data stock taking laporan cost
                Me.Simpan_ItemCost_UsingScript(PrimaryStockTaking)

                'konfirmasi ke user
                MsgBox("Finished", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTACheck1.Dispose()
            DTABuffer.Dispose()
            DTAProductQty.Dispose()
            DTASC.Dispose()
            DTAPenyesuaian.Dispose()

            DS.Dispose()
            Obj = Nothing
            ObjProduct = Nothing
            ObjSC = Nothing

            DTAFinalCheck.Dispose()
            TBL_FinalCheck.Dispose()

            DTA_ED.Dispose()
            TBL_ED.Dispose()
            TBLAcuan_ED.Dispose()
            ObjClosing = Nothing
            ObjCD = Nothing
        End Try
    End Sub

    Public Function Simpan_ItemCost_UsingScript(ByVal vPrimaryStockTaking As Double) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0
        Dim DTA As New DseLaporanStockTakingBaruTableAdapters.DataTable2_1TableAdapter
        Dim vTBL As New DseLaporanStockTakingBaru.DataTable2_1DataTable

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'hapuskan data lama
            DTA.DeleteQuery(vPrimaryStockTaking)

            'isikan data baru
            DTA.Fill(vTBL)

            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemCost(StartRowIndex, BatchQty, vTBL, vPrimaryStockTaking)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
            DTA.Dispose()
            vTBL.Dispose()
        End Try
    End Function

    Private Function GetInsertString_ItemCost(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DseLaporanStockTakingBaru.DataTable2_1DataTable, ByVal vPrimaryStockTaking As Double) As String
        Dim retval As String = ""
        Dim row As DseLaporanStockTakingBaru.DataTable2_1Row
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltrans_stock_taking_laporancost` (`primary closing`, `primary stock taking`, `primary product`, `qty on hand`, `qty taking`, `difference`, `primary unit`, `sub total`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        retval = retval + vbCrLf + "(-1, " + vPrimaryStockTaking.ToString + ", " + _
                        CDbl(.Primary_Product).ToString + "," + CDec(.Qty_On_Hand).ToString + "," _
                        + CDec(.Qty_Taking).ToString + "," _
                        + CDec(.Difference).ToString + "," _
                        + CDbl(.Primary_Unit).ToString + "," _
                        + CDec(0).ToString + ")" + Delimiter
                    End With

                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Private Sub btnRefreshReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefreshReport.Click
        Try
            Me.LoadData()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class