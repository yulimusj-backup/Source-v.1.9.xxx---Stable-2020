Option Explicit On
Option Strict Off

Public Class frmLaporanStockTakingDataCheck
    Public BackColorAll As Color = BackColorAll
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub CetakExpiryDateDataCheck()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim DTA As New DseLaporanStockTaking_ExpiryDataCheckTableAdapters.laporan_stock_taking_view_expirydate3_viewSumberTableAdapter
        Dim DTATbl3 As New DseLaporanStockTaking_ExpiryDataCheckTableAdapters.DataTable3TableAdapter
        Dim DS As New DseLaporanStockTaking_ExpiryDataCheck
        Dim report As New crpLaporanStockTaking_ED_DataCheck
        Dim TipeItem As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods

        Try
            DTA.Fill(DS.laporan_stock_taking_view_expirydate3_viewSumber, Me.txtPrimary.Text, TipeItem)
            DTA.Dispose()
            DTATbl3.Fill(DS.DataTable3, Me.txtPrimary.Text)
            DTATbl3.Dispose()
            Dim rowLaporanSumber As DseLaporanStockTaking_ExpiryDataCheck.laporan_stock_taking_view_expirydate3_viewSumberRow
            Dim rowTbl3 As DseLaporanStockTaking_ExpiryDataCheck.DataTable3Row
            For Each rowLaporanSumber In DS.laporan_stock_taking_view_expirydate3_viewSumber.Rows
                For Each rowTbl3 In DS.DataTable3.Rows
                    If rowLaporanSumber.Primary_Product = rowTbl3.Primary_Product And rowLaporanSumber.Lot_No = rowTbl3.Lot_No And rowLaporanSumber.Expiry_Date = rowTbl3.Expiry_Date Then
                        rowLaporanSumber.Tag_No = rowTbl3.Tag_No
                        Exit For
                    End If
                Next
            Next

            Dim List As New List(Of clsLaporanStockTakingCheckList)
            Dim obj As clsLaporanStockTakingCheckList
            Dim group As Double = 1
            Dim Ketemu As Boolean = False
            'Delete yang sudah benar
            For i As Integer = 0 To DS.laporan_stock_taking_view_expirydate3_viewSumber.Rows.Count - 1
                If i > DS.laporan_stock_taking_view_expirydate3_viewSumber.Rows.Count - 1 Then
                    Exit For
                End If
                If DS.laporan_stock_taking_view_expirydate3_viewSumber.Rows(i).Item("Expiry Date Taking") = DS.laporan_stock_taking_view_expirydate3_viewSumber.Rows(i).Item("Expiry Date") Then
                    DS.laporan_stock_taking_view_expirydate3_viewSumber.Rows.RemoveAt(i)
                    i -= 1
                End If
            Next


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
                    rowLaporanSumber.Group_No = group
                    obj.Group = CStr(group)
                    group += 1
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
                        rowLaporanSumber.Group_No = group
                        obj.Group = CStr(group)
                        group += 1
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
                            rowLaporan.Tag_No += ", " + rowLaporanSumber.Tag_No
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
            Me.CrvLaporan.ReportSource = report
            Me.CrvLaporan.Show()
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
        Dim TipeItemGood As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods

        Try
            DTA.Fill(DS.laporan_stock_taking_view_location1_view, TipeItemGood)
            DTA.Dispose()
            DTAView2.Fill(DS.laporan_stock_taking_view_location2_view, Me.txtPrimary.Text, TipeItemGood)
            DTAView2.Dispose()
            Dim rowLaporanSumber As DseLaporanStockTaking_LocationCheck.laporan_stock_taking_view_location1_viewRow
            Dim rowView2 As DseLaporanStockTaking_LocationCheck.laporan_stock_taking_view_location2_viewRow
            For Each rowLaporanSumber In DS.laporan_stock_taking_view_location1_view.Rows
                For Each rowView2 In DS.laporan_stock_taking_view_location2_view.Rows
                    If rowLaporanSumber.Primary_Product = rowView2.Primary_Product And _
                    rowLaporanSumber.Primary_Storage_Area = rowView2.Primary_Storage_Area And _
                    rowLaporanSumber.Lot_No = rowView2.Lot_No Then
                        rowLaporanSumber.Tag_No = rowView2.Tag_No
                        Exit For
                    End If
                Next
            Next

            Dim List As New List(Of clsLaporanStockTakingCheckList)
            Dim obj As clsLaporanStockTakingCheckList
            Dim group As Double = 1
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
                    rowLaporanSumber.Group_No = group
                    obj.Group = CStr(group)
                    group += 1
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
                        rowLaporanSumber.Group_No = group
                        obj.Group = CStr(group)
                        group += 1
                        List.Add(obj)
                    End If
                End If
            Next
            List = Nothing

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDays.ParameterFieldName, 1)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, 3)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, 1)
            Me.CrvLaporan.ReportSource = report
            Me.CrvLaporan.Show()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
        End Try
    End Sub
    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Dim hj As New FrmCariTransaksiStockTaking
        Try
            '   FrmCariTransaksiStockTaking.JenisTransaks = Me.TipeTransaksi
            hj.ShowDialog()
            If Not hj.PrimaryTransaksiTaking = -1 Then
                Me.txtPrimary.Text = hj.PrimaryTransaksiTaking
                Dim obj As New clsTransaksiStockTaking
                Me.txtTakingNo.Text = obj.GetNoTransaksi(Me.txtPrimary.Text)
                obj = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub txtTakingNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTakingNo.GotFocus
        txtTakingNo.BackColor = Me.BackColorAll
    End Sub

    Private Sub txtTakingNo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTakingNo.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                Dim obj As New clsTransaksiStockTaking
                Dim primary As String = obj.GetPrimary(Me.txtTakingNo.Text)
                obj = Nothing
                If primary <> "" Then
                    Me.txtPrimary.Text = primary
                Else
                    Me.btnSearch.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTakingNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTakingNo.Leave
        Try
            If Me.txtTakingNo.Text <> "" Then
                Dim obj As New clsTransaksiStockTaking
                Dim primary As String = obj.GetPrimary(Me.txtTakingNo.Text)
                obj = Nothing
                If primary <> "" Then
                    Me.txtPrimary.Text = primary
                Else
                    Me.btnSearch.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Try
            If Me.cboPeriod.SelectedIndex > -1 Then
                If Me.cboPeriod.SelectedItem = "Qty Data Check" Then
                    Me.CetakQtyDataCheck()
                Else
                    Me.CetakExpiryDateDataCheck()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTakingNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTakingNo.LostFocus
        txtTakingNo.BackColor = Me.txtPrimary.BackColor
    End Sub
End Class