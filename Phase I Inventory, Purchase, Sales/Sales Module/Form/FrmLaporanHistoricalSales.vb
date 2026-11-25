Option Explicit On
Option Strict Off

Public Class FrmLaporanHistoricalSales
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim TBlBrand As New Core_Module.dseMasterProductGroup.tblmaster_product_group2DataTable
    Dim TBLCust As New Core_Module.dseCustomer.tblmaster_customer1DataTable
    Dim ObjLicense As New Core_Module.ClsLicense

    Private Sub chkBrand_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBrand.CheckedChanged
        Try
            If Me.chkBrand.Checked = True Then
                Me.clbBrand.Enabled = False
                For i As Integer = 0 To Me.clbBrand.Items.Count - 1
                    Me.clbBrand.SetItemCheckState(i, CheckState.Checked)
                Next
            Else
                For i As Integer = 0 To Me.clbBrand.Items.Count - 1
                    Me.clbBrand.SetItemCheckState(i, CheckState.Unchecked)
                Next
                Me.clbBrand.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanHistoricalSales_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.WindowState = FormWindowState.Normal
            Dim ObjBrand As New Core_Module.ClsMasterProductGroup
            ObjBrand.IsiListBoxBrand(Me.clbBrand, Me.TBlBrand)
            ObjBrand = Nothing
            Dim objCust As New Core_Module.clsMasterCustomer
            objCust.IsiListBoxCustomer(Me.clbCustomer, TBLCust)
            objCust = Nothing
            Me.txtPrimaryCodeNo.Text = "-1"
            Me.CboKonversi.SelectedIndex = 0
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub chkCustomer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCustomer.CheckedChanged
        Try
            If Me.chkCustomer.Checked = True Then
                Me.clbCustomer.Enabled = False
                For i As Integer = 0 To Me.clbCustomer.Items.Count - 1
                    Me.clbCustomer.SetItemCheckState(i, CheckState.Checked)
                Next
            Else
                For i As Integer = 0 To Me.clbCustomer.Items.Count - 1
                    Me.clbCustomer.SetItemCheckState(i, CheckState.Unchecked)
                Next
                Me.clbCustomer.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnShowReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShowReport.Click
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.dtpFrom.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpFrom.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpTo.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpTo.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If Me.CekApakahBolehCetak() = True Then
                Me.Cetak()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function CekApakahBolehCetak() As Boolean
        Try
            ''If Me.txtCodeNo.Text = "" Then
            ''    MsgBox("Please insert product code.", MsgBoxStyle.Information, "Information")
            ''    Me.txtCodeNo.Focus()
            ''    Return False
            ''End If

            If Me.clbBrand.CheckedItems.Count <= 0 Then
                MsgBox("Please select brand.", MsgBoxStyle.Information, "Information")
                Me.chkBrand.Focus()
                Return False
            End If

            If Me.clbCustomer.CheckedItems.Count <= 0 Then
                MsgBox("Please select customer.", MsgBoxStyle.Information, "Information")
                Me.chkCustomer.Focus()
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub txtCodeNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodeNo.GotFocus
        Me.txtCodeNo.BackColor = Me.BackColorAll
    End Sub
    Private Sub txtCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodeNo.Leave
        Dim hj As New Core_Module.frmCariMasterProduct
        Dim objProduct As New Core_Module.clsMasterProduct
        hj.BackColorAll = Me.BackColorAll
        Try
            If Me.txtCodeNo.Text <> "" Then
                If objProduct.ApakahSudahAda(Me.txtCodeNo.Text) = True Then
                    Me.txtPrimaryCodeNo.Text = objProduct.GetPrimary(Me.txtCodeNo.Text)
                    Me.txtCodeNo.Text = objProduct.GetProductCode(Me.txtPrimaryCodeNo.Text)
                    Me.txtDescription.Text = objProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                Else
                    If Me.txtCodeNo.Text <> "" Then
                        hj.txtCari.Text = Me.txtCodeNo.Text
                    End If
                    hj.ShowDialog()

                    If hj.PrimaryProduct > 0 Then
                        Me.txtPrimaryCodeNo.Text = hj.PrimaryProduct
                        Me.txtCodeNo.Text = objProduct.GetProductCode(Me.txtPrimaryCodeNo.Text)
                        Me.txtDescription.Text = objProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                    Else
                        Me.txtPrimaryCodeNo.Text = "-1"
                        Me.txtDescription.Clear()
                        Me.txtCodeNo.Clear()
                        Me.txtCodeNo.Focus()
                    End If
                End If
            Else
                Me.txtPrimaryCodeNo.Text = "-1"
                Me.txtDescription.Clear()
                Me.txtCodeNo.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub
    Private Sub txtCodeNo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCodeNo.KeyDown
        Dim objProduct As New Core_Module.clsMasterProduct
        Try
            If e.KeyValue = 13 Then
                If Me.txtCodeNo.Text <> "" Then
                    If objProduct.ApakahSudahAda(Me.txtCodeNo.Text) = True Then
                        Me.txtPrimaryCodeNo.Text = objProduct.GetPrimary(Me.txtCodeNo.Text)
                        Me.txtCodeNo.Text = objProduct.GetProductCode(Me.txtPrimaryCodeNo.Text)
                        Me.txtDescription.Text = objProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                    Else
                        Me.btnFindCodeNo.PerformClick()
                    End If
                Else
                    Me.btnFindCodeNo.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub btnFindCodeNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCodeNo.Click
        Dim objProduct As New Core_Module.clsMasterProduct
        Dim hj As New Core_Module.frmCariMasterProduct
        Try
            hj.BackColorAll = Me.BackColorAll
            If txtCodeNo.Text <> "" Then
                hj.txtCari.Text = txtCodeNo.Text
            End If

            hj.ShowDialog()
            If hj.PrimaryProduct > 0 Then
                Me.txtPrimaryCodeNo.Text = hj.PrimaryProduct
                txtCodeNo.Text = objProduct.GetProductCode(Me.txtPrimaryCodeNo.Text)
                Me.txtDescription.Text = objProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
            Else
                Me.txtPrimaryCodeNo.Text = "-1"
                Me.txtDescription.Clear()
                txtCodeNo.Clear()
                txtCodeNo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub
    Private Sub txtCodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodeNo.LostFocus
        Me.txtCodeNo.BackColor = Me.txtPrimaryCodeNo.BackColor
    End Sub

    Private Sub Cetak()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim report As New CrpLaporanHistoricalSales
        Dim DS As New DseLaporanTransaksiHistoricalSales

        Dim DTA As New DseLaporanTransaksiHistoricalSalesTableAdapters.laporan_historicalsales_viewTableAdapter
        Dim tbl As DseLaporanTransaksiHistoricalSales.laporan_historicalsales_viewDataTable = DS.laporan_historicalsales_view
        Dim tblDummy As New DseLaporanTransaksiHistoricalSales.laporan_historicalsales_viewDataTable
        Dim row, rowDummy As DseLaporanTransaksiHistoricalSales.laporan_historicalsales_viewRow
        Dim tbl2 As New DseLaporanTransaksiHistoricalSales.DataTable3DataTable
        Dim row2 As DseLaporanTransaksiHistoricalSales.DataTable3Row

        Dim Brand, Customer As String
        Dim DR As DataRowView
        Dim Ketemu As Boolean = False
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim dv As DataRowView
        Dim dv2 As DataRowView

        Try
            If Me.txtPrimaryCodeNo.Text <> "-1" Then
                DTA.FillBy_PrimaryProductNTanggalDariSampai(tblDummy, Me.dtpFrom.Value, Me.dtpTo.Value, Me.txtPrimaryCodeNo.Text)
            Else
                DTA.FillBy_TanggalDariSampai(tblDummy, Me.dtpFrom.Value, Me.dtpTo.Value)
            End If

            DTA.Dispose()
            If Me.chkBrand.Checked = False Then
                For i As Integer = 0 To tblDummy.Rows.Count - 1
                    Ketemu = False
                    If i = tblDummy.Rows.Count Then
                        Exit For
                    End If
                    For Each DR In Me.clbBrand.CheckedItems
                        Brand = DR.Item(1)
                        If Brand = tblDummy.Rows(i).Item(tblDummy.Brand_CodeColumn.ColumnName) Then
                            Ketemu = True
                            Exit For
                        End If
                    Next
                    If Ketemu = False Then
                        tblDummy.Rows.RemoveAt(i)
                        i -= 1
                    End If
                Next
            End If

            If Me.chkCustomer.Checked = False Then
                For i As Integer = 0 To tblDummy.Rows.Count - 1
                    Ketemu = False
                    If i = tblDummy.Rows.Count Then
                        Exit For
                    End If
                    For Each DR In Me.clbCustomer.CheckedItems
                        Customer = DR.Item(1)
                        If Customer = tblDummy.Rows(i).Item(tblDummy.Customer_CodeColumn.ColumnName) Then
                            Ketemu = True
                            Exit For
                        End If
                    Next
                    If Ketemu = False Then
                        tblDummy.Rows.RemoveAt(i)
                        i -= 1
                    End If
                Next
            End If

            Brand = ""
            If Me.chkBrand.Checked = True Then
                Brand = "All"
            Else
                For Each DR In Me.clbBrand.CheckedItems
                    If Brand = "" Then
                        Brand = DR.Item(1)
                    Else
                        Brand &= ", " + DR.Item(1)
                    End If
                Next
            End If

            If tblDummy.Rows.Count > 0 Then
                For i As Integer = 0 To tblDummy.Rows.Count - 1
                    rowDummy = tblDummy.DefaultView(i).Row
                    rowDummy.DefPrec = ObjProduct.GetAngkaDecimal(rowDummy.Primary_Product, rowDummy.Primary_Units_Multi)
                    'aul tambahan
                    If Me.CboKonversi.SelectedIndex = 1 Then
                        rowDummy.Unit_Price = rowDummy.Unit_Price * rowDummy.Exchange_Rate
                        If rowDummy.Discount <> 0 Then
                            rowDummy.Discount = rowDummy.Discount
                        End If
                        rowDummy.Nett_Price = rowDummy.Unit_Price - ((rowDummy.Unit_Price * rowDummy.Discount) / 100)
                    End If
                    tbl2.DefaultView.RowFilter = Nothing
                    tbl2.DefaultView.RowFilter = "[Primary Product] = " & rowDummy.Primary_Product.ToString
                    If tbl2.DefaultView.Count > 0 Then
                        row2 = tbl2.DefaultView(0).Row
                        row2.Total_Qty += rowDummy.Total_Qty
                    Else
                        row2 = tbl2.NewDataTable3Row
                        row2.Primary_Product = rowDummy.Primary_Product
                        row2.Total_Qty = rowDummy.Total_Qty
                        tbl2.Rows.Add(row2)
                    End If
                    '---
                Next
            End If

            'aul
            tbl2.DefaultView.RowFilter = Nothing
            tbl2.DefaultView.Sort = "[Total Qty] DESC"
            If tbl2.DefaultView.Count > 0 Then
                For Each dv In tbl2.DefaultView
                    row2 = dv.Row
                    tblDummy.DefaultView.RowFilter = Nothing
                    tblDummy.DefaultView.RowFilter = "[Primary Product] = " & row2.Primary_Product
                    If tblDummy.DefaultView.Count > 0 Then
                        tblDummy.DefaultView.Sort = "[DO Date]"
                        For Each dv2 In tblDummy.DefaultView
                            rowDummy = dv2.Row
                            row = tbl.Newlaporan_historicalsales_viewRow
                            For ijk As Integer = 0 To tbl.Columns.Count - 1
                                row.Item(ijk) = rowDummy.Item(ijk)
                            Next
                            tbl.Rows.Add(row)
                        Next
                    End If
                Next
            End If
            '---
            Me.SetDecimal(DS.laporan_historicalsales_view)
            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_Brand.ParameterFieldName, Brand)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, 1)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, 2)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, 2)
            report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, 1)
            report.SetParameterValue(report.Parameter_ItemCode.ParameterFieldName, Me.txtCodeNo.Text)
            report.SetParameterValue(report.Parameter_TglDari.ParameterFieldName, Me.dtpFrom.Value)
            report.SetParameterValue(report.Parameter_TglSampai.ParameterFieldName, Me.dtpTo.Value)
            report.SetParameterValue(report.Parameter_Konversi.ParameterFieldName, Me.CboKonversi.SelectedIndex)
            Me.crv1.ReportSource = report
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
        End Try
    End Sub

    Private Function SetDecimal(ByRef tbl As DseLaporanTransaksiHistoricalSales.laporan_historicalsales_viewDataTable) As Boolean
        Dim Row As DseLaporanTransaksiHistoricalSales.laporan_historicalsales_viewRow
        Dim dt As New DseLaporanTransaksiHistoricalSalesTableAdapters.DataTable2TableAdapter
        Dim t As New DseLaporanTransaksiHistoricalSales.DataTable2DataTable
        Dim r As DseLaporanTransaksiHistoricalSales.DataTable2Row
        Try
            dt.Connection.ConnectionString = Core_Module.GetConstring
            dt.Fill(t)
            For Each Row In tbl.Rows
                t.DefaultView.RowFilter = "[Primary Product]=" & Row.Primary_Product & " AND [Primary Unit]=" & Row.Primary_Units_Multi
                If t.DefaultView.Count > 0 Then
                    r = t.DefaultView(0).Row
                    Row.DefPrec = r._Decimal
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class