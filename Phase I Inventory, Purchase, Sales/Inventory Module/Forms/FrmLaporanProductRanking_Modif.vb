Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmLaporanProductRanking_Modif
    Dim ObjCategory As New ClsMasterCategory
    Dim ObjProduct As New clsMasterProduct
    Dim DTAClosing As New dsLaporanAdjustmentListTableAdapters.DataTable1TableAdapter
    Dim TBLClosing As New dsLaporanAdjustmentList.DataTable1DataTable
    Dim TanggalClosing As Date = #1/1/1950#
    Dim ObjCD As New Core_Module.ClsCekData
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjLicense As New Core_Module.ClsLicense

    Public Enum EnumKategory As Integer
        _6_Bulan
        _12_Bulan
        _18_Bulan
        _24_bulan
        _36_Bulan
    End Enum

    Private Sub txtCategory_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCategory.KeyDown
        Dim i As New frmCariMasterCategory
        Dim Jml As Long = 0
        Try
            If e.KeyValue = 13 Then
                Jml = ObjCategory.GetJml(Me.txtPrimaryCategory.Text)
                If Jml = 1 Then
                    Me.txtPrimaryCategory.Text = Me.ObjCategory.GetCategoryCode(Me.txtPrimaryCategory.Text)
                    Me.txtCategoryDesc.Text = Me.ObjCategory.GetDesc(Me.txtPrimaryCategory.Text)
                    Me.caribrand()
                    Me.txtParentGroupCode.ReadOnly = True
                    Me.btnFindParentGroup.Enabled = False
                    Me.btnFilter.Focus()
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.txtCategory.Text <> "" Then
                        i.txtCari.Text = Me.txtCategory.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCategory > 0 Then
                        Me.txtPrimaryCategory.Text = i.PrimaryCategory
                        Me.txtCategory.Text = Me.ObjCategory.GetCategoryCode(Me.txtPrimaryCategory.Text)
                        Me.txtCategoryDesc.Text = Me.ObjCategory.GetDesc(Me.txtPrimaryCategory.Text)
                        Me.caribrand()
                        Me.txtParentGroupCode.ReadOnly = True
                        Me.btnFindParentGroup.Enabled = False
                        Me.btnFilter.Focus()
                    Else
                        Me.txtCategory.Clear()
                        Me.txtPrimaryCategory.Text = "-1"
                        Me.txtPrimaryParentGroup.Text = "-1"
                        Me.txtParentGroupCode.Clear()
                        Me.txtBrandDesc.Clear()
                        Me.txtParentGroupCode.ReadOnly = False
                        Me.btnFindParentGroup.Enabled = True
                        Me.txtParentGroupCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtCategory_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCategory.Leave
        Dim i As New frmCariMasterCategory
        Try
            If Me.txtCategory.Text <> "" Then
                If Me.ObjCategory.ApakahSudahAda(Me.txtCategory.Text) = True Then
                    Me.txtPrimaryCategory.Text = Me.ObjCategory.GetPrimary(Me.txtCategory.Text)
                    Me.txtCategoryDesc.Text = Me.ObjCategory.GetDesc(Me.txtPrimaryCategory.Text)
                    Me.caribrand()
                    Me.txtParentGroupCode.ReadOnly = True
                    Me.btnFindParentGroup.Enabled = False
                    Me.btnFilter.Focus()
                Else
                    If Me.txtCategory.Text <> "" Then
                        i.txtCari.Text = Me.txtCategory.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCategory > 0 Then
                        Me.txtPrimaryCategory.Text = i.PrimaryCategory
                        Me.txtCategory.Text = Me.ObjCategory.GetCategoryCode(Me.txtPrimaryCategory.Text)
                        Me.txtCategoryDesc.Text = Me.ObjCategory.GetDesc(Me.txtPrimaryCategory.Text)
                        Me.caribrand()
                        Me.txtParentGroupCode.ReadOnly = True
                        Me.btnFindParentGroup.Enabled = False
                        Me.btnFilter.Focus()
                    Else
                        Me.txtCategory.Clear()
                        Me.txtPrimaryCategory.Text = "-1"
                        Me.txtPrimaryParentGroup.Text = "-1"
                        Me.txtParentGroupCode.Clear()
                        Me.txtBrandDesc.Clear()
                        Me.txtParentGroupCode.ReadOnly = False
                        Me.btnFindParentGroup.Enabled = True
                        Me.txtParentGroupCode.Focus()
                    End If
                End If
            Else
                Me.txtPrimaryCategory.Text = "-1"
                Me.txtCategoryDesc.Clear()
                Me.txtParentGroupCode.ReadOnly = False
                Me.btnFindParentGroup.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindCategory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCategory.Click
        Dim i As New frmCariMasterCategory
        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.txtCategory.Text <> "" Then
                i.txtCari.Text = Me.txtCategory.Text
            End If

            i.ShowDialog()
            If i.PrimaryCategory > 0 Then
                Me.txtPrimaryCategory.Text = i.PrimaryCategory
                Me.txtCategory.Text = Me.ObjCategory.GetCategoryCode(Me.txtPrimaryCategory.Text)
                Me.txtCategoryDesc.Text = Me.ObjCategory.GetDesc(Me.txtPrimaryCategory.Text)
                Me.caribrand()
                Me.txtParentGroupCode.ReadOnly = True
                Me.btnFindParentGroup.Enabled = False
                Me.btnFilter.Focus()
            Else
                Me.txtCategory.Clear()
                Me.txtPrimaryCategory.Text = "-1"
                Me.txtPrimaryParentGroup.Text = "-1"
                Me.txtParentGroupCode.Clear()
                Me.txtBrandDesc.Clear()
                Me.txtParentGroupCode.ReadOnly = False
                Me.btnFindParentGroup.Enabled = True
                Me.txtParentGroupCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub caribrand()
        Try
            If Me.txtPrimaryCategory.Text <> "" Then
                Me.txtPrimaryParentGroup.Text = CStr(ObjCategory.GetPrimaryBrand(Me.txtPrimaryCategory.Text))
                Me.txtParentGroupCode.Text = ObjCategory.GetBrandCode(Me.txtPrimaryCategory.Text)
                Me.txtBrandDesc.Text = ObjCategory.GetBrandDescription(Me.txtPrimaryCategory.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtParentGroupCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtParentGroupCode.GotFocus

    End Sub

    Private Sub txtParentGroupCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtParentGroupCode.KeyDown
        Dim i As New FrmCariMasterProductGroup
        Dim Obj As New ClsMasterProductGroup

        Try
            '' i.BackColorAll = Me.BackcolorAll
            If Me.txtParentGroupCode.ReadOnly = False Then
                If e.KeyValue = 13 Then
                    If Obj.ApakahSudahAda(Me.txtParentGroupCode.Text, 2) = True Then
                        Me.txtPrimaryParentGroup.Text = Obj.GetPrimary(Me.txtParentGroupCode.Text, 2)
                        Me.txtBrandDesc.Text = Obj.GetGroupDesc(Me.txtPrimaryParentGroup.Text)
                        Me.txtCategory.Focus()
                    Else
                        If Me.txtParentGroupCode.Text <> "" Then
                            i.txtCari.Text = Me.txtParentGroupCode.Text
                        End If

                        i.ProductGroupLevel = 2
                        i.ShowDialog()

                        If i.PrimaryProductGroup > 0 Then
                            Me.txtPrimaryParentGroup.Text = i.PrimaryProductGroup
                            Me.txtParentGroupCode.Text = Obj.GetGroupCode(Me.txtPrimaryParentGroup.Text)
                            Me.txtBrandDesc.Text = Obj.GetGroupDesc(Me.txtPrimaryParentGroup.Text)
                            Me.txtCategory.Focus()
                        Else
                            Me.txtParentGroupCode.Clear()
                            Me.txtPrimaryParentGroup.Text = "-1"
                            Me.txtParentGroupCode.Focus()
                        End If
                    End If
                Else
                    Me.txtPrimaryParentGroup.Text = "-1"
                    Me.txtBrandDesc.Clear()
                    Me.txtParentGroupCode.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            Obj = Nothing
        End Try
    End Sub

    Private Sub txtParentGroupCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtParentGroupCode.Leave
        Dim i As New FrmCariMasterProductGroup
        Dim obj As New ClsMasterProductGroup

        Try
            ''i.BackColorAll = Me.BackcolorAll
            If Me.txtParentGroupCode.ReadOnly = False Then
                If Me.txtParentGroupCode.Text <> "" Then
                    If obj.ApakahSudahAda(Me.txtParentGroupCode.Text, 2) = True Then
                        Me.txtPrimaryParentGroup.Text = obj.GetPrimary(Me.txtParentGroupCode.Text, 2)
                        Me.txtBrandDesc.Text = obj.GetGroupDesc(Me.txtPrimaryParentGroup.Text)
                        Me.txtCategory.Focus()
                    Else
                        If Me.txtParentGroupCode.Text <> "" Then
                            i.txtCari.Text = Me.txtParentGroupCode.Text
                        End If

                        i.ProductGroupLevel = 2
                        i.ShowDialog()
                        If i.PrimaryProductGroup > 0 Then
                            Me.txtPrimaryParentGroup.Text = i.PrimaryProductGroup
                            Me.txtParentGroupCode.Text = obj.GetGroupCode(Me.txtPrimaryParentGroup.Text)
                            Me.txtBrandDesc.Text = obj.GetGroupDesc(Me.txtPrimaryParentGroup.Text)
                            Me.txtCategory.Focus()
                        Else
                            Me.txtParentGroupCode.Clear()
                            Me.txtPrimaryParentGroup.Text = "-1"
                            Me.txtParentGroupCode.Focus()
                        End If
                    End If
                Else
                    Me.txtPrimaryParentGroup.Text = "-1"
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFindParentGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindParentGroup.Click
        Dim i As New FrmCariMasterProductGroup
        Dim obj As New ClsMasterProductGroup
        Try
            '' i.BackColorAll = Me.BackcolorAll
            '' i.cboCriteria.SelectedIndex = 0
            i.ProductGroupLevel = 2
            If Me.txtParentGroupCode.Text <> "" Then
                i.txtCari.Text = Me.txtParentGroupCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryProductGroup > 0 Then
                Me.txtPrimaryParentGroup.Text = i.PrimaryProductGroup
                Me.txtParentGroupCode.Text = obj.GetGroupCode(Me.txtPrimaryParentGroup.Text)
                Me.txtBrandDesc.Text = obj.GetGroupDesc(Me.txtPrimaryParentGroup.Text)
                Me.txtCategory.Focus()
            Else
                Me.txtParentGroupCode.Clear()
                Me.txtPrimaryParentGroup.Text = "-1"
                Me.txtParentGroupCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Dim EndLicenseDate As Date = #1/1/1950#
        Try

            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.dtpStartDate.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpStartDate.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpEndDate.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpEndDate.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If Me.chkAbsoluteData.Checked = True Then
                If RBNoGroup.Checked = True Then
                    Me.Cetak1_Absolute()
                End If
                If RBGroup.Checked = True Then
                    If Me.chkCost.Checked = True Then
                        Me.Cetak2_Absolute_Modified()
                    Else
                        Me.Cetak2_Absolute()
                    End If
                End If
            Else
                If RBNoGroup.Checked = True Then
                    Me.Cetak1()
                End If
                If RBGroup.Checked = True Then
                    Me.Cetak2()
                End If
            End If

            Windows.Forms.Cursor.Current = Cursors.Arrow
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' grouping by brand
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Cetak2()
        Dim report As New CrpProductRanking2
        Dim DS As New DseLaporanProductRanking_Modif
        Dim objCD As New Core_Module.ClsCekData
        Dim TglDari As Date = objCD.GetFirstDateofTheMonth(Me.dtpStartDate.Value)
        Dim TglSampai As Date = objCD.GetLasDateOfTheMonth(Me.dtpEndDate.Value)
        objCD = Nothing
        Dim brand As String = ""
        Dim category As String = ""
        Dim SortBy As String = ""
        Dim groupby As String = ""
        Dim Period As String = Format(Me.dtpStartDate.Value, "dd MMM yyyy") + " - " + Format(Me.dtpEndDate.Value, "dd MMM yyyy")
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim formatday As Integer = 1
        Dim formatmonth As Integer = 2
        Dim formatyear As Integer = 1
        Dim DTA As New DseLaporanProductRanking_ModifTableAdapters.DataTable1TableAdapter
        Try
            If Me.cboSort.SelectedIndex = 0 Then
                '1. Ascending
                If Me.txtPrimaryParentGroup.Text <> "-1" And Me.txtPrimaryCategory.Text = "-1" Then
                    '   a. Filter By Brand
                    DTA.FillBy_PrimaryBrandASCENDING(DS.DataTable1, dtpEndDate.Value, dtpStartDate.Value, _
                                                Me.txtParentGroupCode.Text)
                    brand = Me.txtParentGroupCode.Text
                    category = "All"
                    SortBy = "Ascending"
                    groupby = "Brand"

                ElseIf Me.txtPrimaryParentGroup.Text <> "-1" And Me.txtPrimaryCategory.Text <> "-1" Then
                    '   b. Filter By Category
                    DTA.FillBy_CategoryASCENDING(DS.DataTable1, Me.dtpEndDate.Value, Me.dtpStartDate.Value, _
                    Me.txtCategory.Text)
                    brand = Me.txtParentGroupCode.Text
                    category = Me.txtCategory.Text
                    SortBy = "Ascending"
                    groupby = "Brand"

                ElseIf Me.txtPrimaryParentGroup.Text = "-1" And Me.txtPrimaryCategory.Text = "-1" Then
                    '   c. No FIlter()
                    DTA.FillBy_ascending(DS.DataTable1, Me.dtpEndDate.Value, _
                    Me.dtpStartDate.Value)
                    brand = "All"
                    category = "All"
                    SortBy = "Ascending"
                    groupby = "Brand"
                End If

            ElseIf Me.cboSort.SelectedIndex = 1 Then
                '2. Descending
                If Me.txtPrimaryParentGroup.Text <> "-1" And Me.txtPrimaryCategory.Text = "-1" Then
                    '   a. Filter By Brand
                    DTA.FillBy_PrimaryBrandDESCENDING(DS.DataTable1, Me.dtpEndDate.Value, Me.dtpStartDate.Value, _
                                                     Me.txtParentGroupCode.Text)
                    brand = Me.txtParentGroupCode.Text
                    category = "All"
                    SortBy = "Descending"
                    groupby = "Brand"

                ElseIf Me.txtPrimaryParentGroup.Text <> "-1" And Me.txtPrimaryCategory.Text <> "-1" Then
                    '   b. Filter By Category
                    DTA.FillBy_CategoryDESCENDING(DS.DataTable1, dtpEndDate.Value, dtpStartDate.Value, _
                                                Me.txtCategory.Text)
                    brand = Me.txtParentGroupCode.Text
                    category = Me.txtCategory.Text
                    SortBy = "Descending"
                    groupby = "Brand"

                ElseIf Me.txtPrimaryParentGroup.Text = "-1" And Me.txtPrimaryCategory.Text = "-1" Then
                    '   c. No FIlter
                    DTA.FillBy_descending(DS.DataTable1, Me.dtpEndDate.Value, Me.dtpStartDate.Value)
                    brand = "All"
                    category = "All"
                    SortBy = "Descending"
                    groupby = "Brand"
                End If
            Else
                MsgBox("Please Choose Sort by", MsgBoxStyle.Information, "Information")
            End If
            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FilterByBrand.ParameterFieldName, brand)
            report.SetParameterValue(report.Parameter_FilterByCategory.ParameterFieldName, category)
            report.SetParameterValue(report.Parameter_AscendingDescending.ParameterFieldName, SortBy)
            report.SetParameterValue(report.Parameter_Quantity.ParameterFieldName, 0)
            report.SetParameterValue(report.Parameter_Tanggal.ParameterFieldName, Period)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, formatday)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, formatmonth)
            report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, formatyear)
            report.SetParameterValue(report.Parameter_GroupBy.ParameterFieldName, groupby)
            report.SetParameterValue(report.Parameter_Judul_Laporan.ParameterFieldName, "PRODUCT RANGKING")
            report.SetParameterValue(report.Parameter_Judul_Qty.ParameterFieldName, "Moving Out")

            Me.crv.ReportSource = report
            Me.crv.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
        End Try
    End Sub

    Private Sub Cetak2_Absolute()
        Dim report As New CrpProductRanking2
        Dim DS As New DseLaporanProductRanking_Modif
        Dim objCD As New Core_Module.ClsCekData
        Dim TglDari As Date = objCD.GetFirstDateofTheMonth(Me.dtpStartDate.Value)
        Dim TglSampai As Date = objCD.GetLasDateOfTheMonth(Me.dtpEndDate.Value)
        objCD = Nothing
        Dim brand As String = ""
        Dim category As String = ""
        Dim SortBy As String = ""
        Dim groupby As String = ""
        Dim Period As String = Format(Me.dtpStartDate.Value, "dd MMM yyyy") + " - " + Format(Me.dtpEndDate.Value, "dd MMM yyyy")
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim formatday As Integer = 1
        Dim formatmonth As Integer = 2
        Dim formatyear As Integer = 1
        Dim DTA As New DseLaporanProductRanking_ModifTableAdapters.DataTable1TableAdapter

        Dim DTAMaster As New DseLaporanProductRanking_ModifTableAdapters.Tbl_AbsoluteTableAdapter
        Dim TBLMaster As New DseLaporanProductRanking_Modif.Tbl_AbsoluteDataTable
        Dim rowMaster As DseLaporanProductRanking_Modif.Tbl_AbsoluteRow
        Dim ij As Double = 0
        Dim filter As String = ""
        Dim row1 As DseLaporanProductRanking_Modif.DataTable1Row

        Dim DTAPurchase As New DseLaporanProductRanking_ModifTableAdapters.DataTable5TableAdapter
        Dim TBLPurchase As DseLaporanProductRanking_Modif.DataTable5DataTable = DS.DataTable5
        Dim TipeTrans As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase

        Try
            If Me.cboSort.SelectedIndex = 0 Then
                '1. Ascending
                If Me.txtPrimaryParentGroup.Text <> "-1" And Me.txtPrimaryCategory.Text = "-1" Then
                    '   a. Filter By Brand
                    DTA.FillBy_PrimaryBrandASCENDING(DS.DataTable1, dtpEndDate.Value, dtpStartDate.Value, _
                                                Me.txtParentGroupCode.Text)
                    brand = Me.txtParentGroupCode.Text
                    category = "All"
                    SortBy = "Ascending"
                    groupby = "Brand"
                    DTAMaster.FillBy_GroupByBrand_PrimaryBrandASC(TBLMaster, Me.txtPrimaryParentGroup.Text)
                ElseIf Me.txtPrimaryParentGroup.Text <> "-1" And Me.txtPrimaryCategory.Text <> "-1" Then
                    '   b. Filter By Category
                    DTA.FillBy_CategoryASCENDING(DS.DataTable1, Me.dtpEndDate.Value, Me.dtpStartDate.Value, _
                    Me.txtCategory.Text)
                    brand = Me.txtParentGroupCode.Text
                    category = Me.txtCategory.Text
                    SortBy = "Ascending"
                    groupby = "Brand"
                    DTAMaster.FillBy_GroupByBrand_PrimaryCategoryASC(TBLMaster, Me.txtPrimaryCategory.Text)
                ElseIf Me.txtPrimaryParentGroup.Text = "-1" And Me.txtPrimaryCategory.Text = "-1" Then
                    '   c. No FIlter()
                    DTA.FillBy_ascending(DS.DataTable1, Me.dtpEndDate.Value, _
                    Me.dtpStartDate.Value)
                    brand = "All"
                    category = "All"
                    SortBy = "Ascending"
                    groupby = "Brand"
                    DTAMaster.FillBy_GroupByBrandQtyASC(TBLMaster)
                End If
            ElseIf Me.cboSort.SelectedIndex = 1 Then
                '2. Descending
                If Me.txtPrimaryParentGroup.Text <> "-1" And Me.txtPrimaryCategory.Text = "-1" Then
                    '   a. Filter By Brand
                    DTA.FillBy_PrimaryBrandDESCENDING(DS.DataTable1, Me.dtpEndDate.Value, Me.dtpStartDate.Value, _
                                                     Me.txtParentGroupCode.Text)
                    brand = Me.txtParentGroupCode.Text
                    category = "All"
                    SortBy = "Descending"
                    groupby = "Brand"
                    DTAMaster.FillBy_GroupByBrand_PrimaryBrandDESC(TBLMaster, Me.txtPrimaryParentGroup.Text)
                ElseIf Me.txtPrimaryParentGroup.Text <> "-1" And Me.txtPrimaryCategory.Text <> "-1" Then
                    '   b. Filter By Category
                    DTA.FillBy_CategoryDESCENDING(DS.DataTable1, dtpEndDate.Value, dtpStartDate.Value, _
                                                Me.txtCategory.Text)
                    brand = Me.txtParentGroupCode.Text
                    category = Me.txtCategory.Text
                    SortBy = "Descending"
                    groupby = "Brand"
                    DTAMaster.FillBy_GroupByBrand_PrimaryCategoryDESC(TBLMaster, Me.txtPrimaryCategory.Text)
                ElseIf Me.txtPrimaryParentGroup.Text = "-1" And Me.txtPrimaryCategory.Text = "-1" Then
                    '   c. No FIlter
                    DTA.FillBy_descending(DS.DataTable1, Me.dtpEndDate.Value, Me.dtpStartDate.Value)
                    brand = "All"
                    category = "All"
                    SortBy = "Descending"
                    groupby = "Brand"
                    DTAMaster.FillBy_GroupByBrandQtyDESC(TBLMaster)
                End If
            Else
                MsgBox("Please Choose Sort by", MsgBoxStyle.Information, "Information")
            End If

            'buang baris, pada tblmaster yang ada di dalam data table 1, artinya ada transaksi keluar selama periode tsb
            While ij < TBLMaster.Count - 1
                rowMaster = TBLMaster.Rows(ij)
                filter = "[Primary Product] = " + rowMaster.Primary_Product.ToString
                DS.DataTable1.DefaultView.RowFilter = filter

                'If rowMaster.Primary_Product = 381 Then
                'MsgBox("")
                'End If
                If DS.DataTable1.DefaultView.Count = 0 Then
                    ij += 1
                Else
                    TBLMaster.Rows.Remove(rowMaster)
                End If
            End While

            'buang baris pada row master, jika, sudah ada pembelian sebelum tanggal mulai
            DTAPurchase.Fill(DS.DataTable5, TipeTrans, Me.dtpEndDate.Value)
            ij = 0
            While ij <= TBLMaster.Count - 1
                rowMaster = TBLMaster.Rows(ij)
                filter = "[Primary Product] =" + rowMaster.Primary_Product.ToString
                DS.DataTable5.DefaultView.RowFilter = filter
                If DS.DataTable5.DefaultView.Count = 0 Then
                    ij += 1
                Else
                    TBLMaster.Rows.Remove(rowMaster)
                End If
            End While

            'pindahkan row master pada data tabel 1
            DS.DataTable1.Rows.Clear()
            For Each rowMaster In TBLMaster.Rows
                row1 = DS.DataTable1.NewDataTable1Row

                With row1
                    .Brand = rowMaster.Brand
                    .Brand_Desc = rowMaster.Brand_Desc
                    .Category_Code = rowMaster.Category_Code
                    .Category_Desc = rowMaster.Category_Desc
                    .Description = rowMaster.Description
                    .Primary_Product = rowMaster.Primary_Product
                    .Product_Code = rowMaster.Product_Code
                    .Qty = rowMaster.Qty
                    .Unit_Code = rowMaster.Unit_Code
                End With

                DS.DataTable1.Rows.Add(row1)
            Next

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FilterByBrand.ParameterFieldName, brand)
            report.SetParameterValue(report.Parameter_FilterByCategory.ParameterFieldName, category)
            report.SetParameterValue(report.Parameter_AscendingDescending.ParameterFieldName, SortBy)
            report.SetParameterValue(report.Parameter_Quantity.ParameterFieldName, 0)
            report.SetParameterValue(report.Parameter_Tanggal.ParameterFieldName, Period)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, formatday)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, formatmonth)
            report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, formatyear)
            report.SetParameterValue(report.Parameter_GroupBy.ParameterFieldName, groupby)
            report.SetParameterValue(report.Parameter_Judul_Laporan.ParameterFieldName, "NON MOVING STOCK")
            report.SetParameterValue(report.Parameter_Judul_Qty.ParameterFieldName, "Qty On Hand")

            Me.crv.ReportSource = report
            Me.crv.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
            DTAPurchase.Dispose()
        End Try
    End Sub

    Private Sub Cetak2_Absolute_Modified()
        Dim report As New CrpProductRanking3
        Dim DS As New DseLaporanProductRanking_Modif

        Dim Period As String = Format(Me.dtpStartDate.Value, "dd MMM yyyy") + " - " + Format(Me.dtpEndDate.Value, "dd MMM yyyy")
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim formatday As Integer = 1
        Dim formatmonth As Integer = 2
        Dim formatyear As Integer = 1
        Dim TBLMaster As New DseLaporanProductRanking_Modif.Tbl_AbsoluteDataTable
        'Dim row1 As DseLaporanProductRanking_Modif.DataTable1Row

        Dim brand As String = ""
        Dim category As String = ""
        Dim SortBy As String = ""
        Dim groupby As String = ""
        Dim TanggalDari As Date = Today
        Dim TBLDummy As New DseLaporanProductRanking_Modif.DataTable1DataTable
        Dim GroupByPeriode As Boolean = False

        Try
            DS.DataTable1.Rows.Clear()

            If Me.cboSort.SelectedIndex = 0 Then
                '1. Ascending
                If Me.txtPrimaryParentGroup.Text <> "-1" And Me.txtPrimaryCategory.Text = "-1" Then
                    '   a. Filter By Brand
                    brand = Me.txtParentGroupCode.Text
                    category = "All"
                    SortBy = "Ascending"
                    groupby = "Brand"
                ElseIf Me.txtPrimaryParentGroup.Text <> "-1" And Me.txtPrimaryCategory.Text <> "-1" Then
                    '   b. Filter By Category
                    brand = Me.txtParentGroupCode.Text
                    category = Me.txtCategory.Text
                    SortBy = "Ascending"
                    groupby = "Brand"
                ElseIf Me.txtPrimaryParentGroup.Text = "-1" And Me.txtPrimaryCategory.Text = "-1" Then
                    '   c. No FIlter()
                    brand = "All"
                    category = "All"
                    SortBy = "Ascending"
                    groupby = "Brand"
                End If
            ElseIf Me.cboSort.SelectedIndex = 1 Then
                '2. Descending
                If Me.txtPrimaryParentGroup.Text <> "-1" And Me.txtPrimaryCategory.Text = "-1" Then
                    '   a. Filter By Brand
                    brand = Me.txtParentGroupCode.Text
                    category = "All"
                    SortBy = "Descending"
                    groupby = "Brand"
                ElseIf Me.txtPrimaryParentGroup.Text <> "-1" And Me.txtPrimaryCategory.Text <> "-1" Then
                    '   b. Filter By Category
                    brand = Me.txtParentGroupCode.Text
                    category = Me.txtCategory.Text
                    SortBy = "Descending"
                    groupby = "Brand"
                ElseIf Me.txtPrimaryParentGroup.Text = "-1" And Me.txtPrimaryCategory.Text = "-1" Then
                    '   c. No FIlter
                    brand = "All"
                    category = "All"
                    SortBy = "Descending"
                    groupby = "Brand"
                End If
            Else
                MsgBox("Please Choose Sort by", MsgBoxStyle.Information, "Information")
            End If

            'isikan data berdasarkan 5 kategori
            For i As Integer = 0 To 4
                If i = 0 Then
                    TanggalDari = Me.GetTanggal(EnumKategory._36_Bulan)
                ElseIf i = 1 Then
                    TanggalDari = Me.GetTanggal(EnumKategory._24_bulan)
                ElseIf i = 2 Then
                    TanggalDari = Me.GetTanggal(EnumKategory._18_Bulan)
                ElseIf i = 3 Then
                    TanggalDari = Me.GetTanggal(EnumKategory._12_Bulan)
                ElseIf i = 4 Then
                    TanggalDari = Me.GetTanggal(EnumKategory._6_Bulan)
                End If

                TBLMaster.Rows.Clear()
                TBLDummy.Rows.Clear()

                Me.IsikanData(TanggalDari, TBLDummy, TBLMaster, DS.DataTable5)

                'pindahkan row master pada data tabel 1
                If i = 0 Then
                    Me.PindahkanBaris(EnumKategory._36_Bulan, TBLMaster, DS)
                ElseIf i = 1 Then
                    Me.PindahkanBaris(EnumKategory._24_bulan, TBLMaster, DS)
                ElseIf i = 2 Then
                    Me.PindahkanBaris(EnumKategory._18_Bulan, TBLMaster, DS)
                ElseIf i = 3 Then
                    Me.PindahkanBaris(EnumKategory._12_Bulan, TBLMaster, DS)
                ElseIf i = 4 Then
                    Me.PindahkanBaris(EnumKategory._6_Bulan, TBLMaster, DS)
                End If
            Next

            'isikan data harga
            Me.IsikanHarga1(DS.DataTable1)

            'buang baris jika ada filter periode
            If Me.cboPeriod.SelectedIndex <> 0 Then
                Me.FilterPeriode(DS.DataTable1)
            End If

            'balikan group by period
            If Me.ChkGroupByPeriod.Checked = True Then
                Me.Group_By_Periode(DS.DataTable1)
                GroupByPeriode = True
            Else
                GroupByPeriode = False
            End If

            'urutkan ulang
            Me.UrutkanUlang(DS.DataTable1, GroupByPeriode)

            report.SetDataSource(DS)
            If GroupByPeriode = True Then
                report.SetParameterValue(report.Parameter_Judul_Period.ParameterFieldName, "Brand")
            Else
                report.SetParameterValue(report.Parameter_Judul_Period.ParameterFieldName, "Period")
            End If
            report.SetParameterValue(report.Parameter_FilterByBrand.ParameterFieldName, brand)
            report.SetParameterValue(report.Parameter_FilterByCategory.ParameterFieldName, category)
            report.SetParameterValue(report.Parameter_AscendingDescending.ParameterFieldName, SortBy)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, formatday)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, formatmonth)
            report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, formatyear)
            report.SetParameterValue(report.Parameter_Judul_Laporan.ParameterFieldName, "NON MOVING STOCK")
            report.SetParameterValue(report.Parameter_Judul_Qty.ParameterFieldName, "Qty On Hand")
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.ObjCur.GetDecimal_DefaultCurrency)

            Me.crv.ReportSource = report
            Me.crv.Show()

            'Dim ij As New System.Drawing.Printing.PaperSource
            'ij.SourceName = "Stgh Letter"

            'report.PrintOptions.CustomPaperSource = ij
            'report.PrintOptions.PaperSource = CrystalDecisions.Shared.PaperSource.Tractor
            'report.PrintOptions.PrinterName = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
        End Try
    End Sub

    Private Sub Group_By_Periode(ByVal TBL As DseLaporanProductRanking_Modif.DataTable1DataTable)
        Dim row As DseLaporanProductRanking_Modif.DataTable1Row
        Dim i As Double = 0
        Dim Dummy As String = ""

        Try
            While i <= TBL.Rows.Count - 1
                row = TBL.Rows(i)

                Dummy = row.Bulan_Str
                row.Bulan_Str = row.Brand1
                row.Brand1 = Dummy
                If row.Bulan = 6 Then
                    row.Brand = "1. " + Dummy
                ElseIf row.Bulan = 12 Then
                    row.Brand = "2. " + Dummy
                ElseIf row.Bulan = 18 Then
                    row.Brand = "3. " + Dummy
                ElseIf row.Bulan = 24 Then
                    row.Brand = "4. " + Dummy
                ElseIf row.Bulan = 36 Then
                    row.Brand = "5. " + Dummy
                End If

                i += 1
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FilterPeriode(ByVal TBL As DseLaporanProductRanking_Modif.DataTable1DataTable)
        Dim row As DseLaporanProductRanking_Modif.DataTable1Row
        Dim i As Double = 0

        Try
            While i <= TBL.Rows.Count - 1
                row = TBL.Rows(i)
                If Me.cboPeriod.SelectedIndex = 1 Then
                    '6 bulan
                    If row.Bulan <> 6 Then
                        TBL.Rows.Remove(row)
                    Else
                        i += 1
                    End If
                ElseIf Me.cboPeriod.SelectedIndex = 2 Then
                    '12 bulan 
                    If row.Bulan <> 12 Then
                        TBL.Rows.Remove(row)
                    Else
                        i += 1
                    End If
                ElseIf Me.cboPeriod.SelectedIndex = 3 Then
                    '18 bulan
                    If row.Bulan <> 18 Then
                        TBL.Rows.Remove(row)
                    Else
                        i += 1
                    End If
                ElseIf Me.cboPeriod.SelectedIndex = 4 Then
                    '24 bulan
                    If row.Bulan <> 24 Then
                        TBL.Rows.Remove(row)
                    Else
                        i += 1
                    End If
                Else
                    '36 bulan
                    If row.Bulan <> 36 Then
                        TBL.Rows.Remove(row)
                    Else
                        i += 1
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UrutkanUlang(ByVal TBL As DseLaporanProductRanking_Modif.DataTable1DataTable, Optional ByVal vGroupByPeriod As Boolean = False)
        Dim TBLDummy As New DseLaporanProductRanking_Modif.DataTable1DataTable
        Dim DV As Data.DataRowView
        Dim rowDummy, row As DseLaporanProductRanking_Modif.DataTable1Row

        Try
            TBL.DefaultView.RowFilter = ""
            If vGroupByPeriod = False Then
                If Me.cboSort.SelectedIndex = 0 Then
                    TBL.DefaultView.Sort = "Brand, Bulan, Amount"
                Else
                    TBL.DefaultView.Sort = "Brand, Bulan, Amount DESC"
                End If
            Else
                If Me.cboSort.SelectedIndex = 0 Then
                    TBL.DefaultView.Sort = "Brand, [Bulan Str], Amount"
                Else
                    TBL.DefaultView.Sort = "Brand, [Bulan Str], Amount DESC"
                End If
            End If

            For Each DV In TBL.DefaultView
                rowDummy = TBLDummy.NewDataTable1Row
                row = DV.Row

                For i As Integer = 0 To TBLDummy.Columns.Count - 1
                    rowDummy.Item(i) = row.Item(i)
                Next

                TBLDummy.Rows.Add(rowDummy)
            Next

            TBL.Rows.Clear()
            For Each rowDummy In TBLDummy.Rows
                row = TBL.NewDataTable1Row
                For i As Integer = 0 To TBLDummy.Columns.Count - 1
                    row.Item(i) = rowDummy.Item(i)
                Next
                TBL.Rows.Add(row)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanHarga1(ByVal TBL As DseLaporanProductRanking_Modif.DataTable1DataTable)
        Dim filter As String = ""
        Dim row1 As DseLaporanProductRanking_Modif.DataTable1Row
        Dim DTAClosingDetail As New DseLaporanProductRanking_ModifTableAdapters.tbltrans_purchase_closing_item_detail1TableAdapter
        Dim TBLClosingDetail As New DseLaporanProductRanking_Modif.tbltrans_purchase_closing_item_detail1DataTable
        Dim rowClosingDetail As DseLaporanProductRanking_Modif.tbltrans_purchase_closing_item_detail1Row
        Dim DefPrec As Integer = 0

        Try
            'sesuaikan tanggal closing, isikan data closing sesuai dengan tanggal closing
            DTAClosingDetail.Connection.ConnectionString = Core_Module.GetConstring
            DTAClosingDetail.Fill(TBLClosingDetail, CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem))
            DefPrec = Me.ObjCur.GetDecimal_DefaultCurrency

            For Each row1 In TBL.Rows
                filter = "[primary code no] =" + row1.Primary_Product.ToString
                TBLClosingDetail.DefaultView.RowFilter = filter
                If TBLClosingDetail.DefaultView.Count <> 0 Then
                    'rowClosingDetail = TBLClosingDetail.DefaultView(0).Row
                    rowClosingDetail = TBLClosingDetail.DefaultView(TBLClosingDetail.DefaultView.Count - 1).Row
                    row1.Amount = Math.Round(rowClosingDetail.amount, DefPrec, MidpointRounding.AwayFromZero)
                Else
                    row1.Amount = Math.Round(0, DefPrec, MidpointRounding.AwayFromZero)
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub IsikanHarga(ByVal TBL As DseLaporanProductRanking_Modif.DataTable1DataTable)
    '    Dim filter As String = ""
    '    Dim row1 As DseLaporanProductRanking_Modif.DataTable1Row
    '    Dim row As dsLaporanAdjustmentList.DataTable1Row

    '    Try
    '        'sesuaikan tanggal closing, isikan data closing sesuai dengan tanggal closing
    '        Me.DTAClosing.Fill(Me.TBLClosing, Me.ObjCD.GetFirstDateofTheMonth(Me.ObjCD.GetFirstDateofTheMonth(Today) - TimeSpan.FromDays(1)))

    '        For Each row1 In TBL.Rows
    '            filter = "[Primary Product] =" + row1.Primary_Product.ToString
    '            Me.TBLClosing.DefaultView.RowFilter = filter
    '            If Me.TBLClosing.DefaultView.Count <> 0 Then
    '                row = Me.TBLClosing.DefaultView(0).Row
    '                If row.Ending_Qty = 0 Or row.Ending_Amount = 0 Then
    '                    If row.Out_Qty = 0 Then
    '                        row1.Amount = 0
    '                    Else
    '                        row1.Amount = Math.Round(row.Out_Amount * row1.Qty / row.Out_Qty, 0, MidpointRounding.AwayFromZero)
    '                    End If
    '                Else
    '                    If row.Ending_Qty = 0 Then
    '                        row1.Amount = 0
    '                    Else
    '                        row1.Amount = Math.Round(row.Ending_Amount * row1.Qty / row.Ending_Qty, 0, MidpointRounding.AwayFromZero)
    '                    End If
    '                End If
    '            Else
    '                row1.Amount = 0
    '            End If
    '        Next

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub PindahkanBaris(ByVal vKategory As EnumKategory, _
    ByVal TBLMaster As DseLaporanProductRanking_Modif.Tbl_AbsoluteDataTable, ByVal DS As DseLaporanProductRanking_Modif)
        Dim rowMaster As DseLaporanProductRanking_Modif.Tbl_AbsoluteRow
        Dim row1 As DseLaporanProductRanking_Modif.DataTable1Row
        Dim Filter As String = ""
        Dim FilterLama As String = ""

        Try
            'catat filter lama
            FilterLama = TBLMaster.DefaultView.RowFilter

            'isikan data baris
            For Each rowMaster In TBLMaster.Rows
                'isikan data baris
                Filter = "[Primary Product] = " + rowMaster.Primary_Product.ToString
                DS.DataTable1.DefaultView.RowFilter = Filter

                If DS.DataTable1.DefaultView.Count = 0 Then
                    row1 = DS.DataTable1.NewDataTable1Row

                    With row1
                        .Brand = rowMaster.Brand
                        .Brand_Desc = rowMaster.Brand_Desc
                        .Category_Code = rowMaster.Category_Code
                        .Category_Desc = rowMaster.Category_Desc
                        .Description = rowMaster.Description
                        .Primary_Product = rowMaster.Primary_Product
                        .Product_Code = rowMaster.Product_Code
                        .Qty = rowMaster.Qty
                        .Unit_Code = rowMaster.Unit_Code
                        ._Decimal = rowMaster._Decimal
                        .Brand1 = .Brand + " - " + .Brand_Desc
                    End With

                    DS.DataTable1.Rows.Add(row1)

                    'isikan data periode non moving
                    If vKategory = EnumKategory._6_Bulan Then
                        row1.Bulan = 6
                        row1.Bulan_Str = "6 Months"
                    ElseIf vKategory = EnumKategory._12_Bulan Then
                        row1.Bulan = 12
                        row1.Bulan_Str = "1 Year"
                    ElseIf vKategory = EnumKategory._18_Bulan Then
                        row1.Bulan = 18
                        row1.Bulan_Str = "1.5 Years"
                    ElseIf vKategory = EnumKategory._24_bulan Then
                        row1.Bulan = 24
                        row1.Bulan_Str = "2 Years"
                    ElseIf vKategory = EnumKategory._36_Bulan Then
                        row1.Bulan = 36
                        row1.Bulan_Str = "3 Years"
                    End If
                End If
            Next

            'kembalikan filter lama
            'MsgBox(i.ToString)
            FilterLama = TBLMaster.DefaultView.RowFilter
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetTanggal(ByVal vKategory As EnumKategory) As Date
        Try
            If vKategory = EnumKategory._6_Bulan Then
                Return Today - TimeSpan.FromDays(6 * 30)
            ElseIf vKategory = EnumKategory._12_Bulan Then
                Return Today - TimeSpan.FromDays(12 * 30)
            ElseIf vKategory = EnumKategory._18_Bulan Then
                Return Today - TimeSpan.FromDays(18 * 30)
            ElseIf vKategory = EnumKategory._24_bulan Then
                Return Today - TimeSpan.FromDays(24 * 30)
            Else
                Return Today - TimeSpan.FromDays(36 * 30)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub IsikanData(ByVal vTanggalDari As Date, ByVal TblDummy As DseLaporanProductRanking_Modif.DataTable1DataTable, _
    ByVal TBLMaster As DseLaporanProductRanking_Modif.Tbl_AbsoluteDataTable, ByVal TblPurchase As DseLaporanProductRanking_Modif.DataTable5DataTable)
        Dim DTAMaster As New DseLaporanProductRanking_ModifTableAdapters.Tbl_AbsoluteTableAdapter
        Dim rowMaster As DseLaporanProductRanking_Modif.Tbl_AbsoluteRow
        Dim DTA As New DseLaporanProductRanking_ModifTableAdapters.DataTable1TableAdapter

        Dim ij As Double = 0
        Dim filter As String = ""

        Dim DTAPurchase As New DseLaporanProductRanking_ModifTableAdapters.DataTable5TableAdapter
        Dim TipeTrans As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase


        Try
            If Me.cboSort.SelectedIndex = 0 Then
                '1. Ascending
                If Me.txtPrimaryParentGroup.Text <> "-1" And Me.txtPrimaryCategory.Text = "-1" Then
                    '   a. Filter By Brand
                    DTA.FillBy_PrimaryBrandASCENDING(TblDummy, dtpEndDate.Value, vTanggalDari, _
                                                Me.txtParentGroupCode.Text)
                    DTAMaster.FillBy_GroupByBrand_PrimaryBrandASC(TBLMaster, Me.txtPrimaryParentGroup.Text)
                ElseIf Me.txtPrimaryParentGroup.Text <> "-1" And Me.txtPrimaryCategory.Text <> "-1" Then
                    '   b. Filter By Category
                    DTA.FillBy_CategoryASCENDING(TblDummy, Me.dtpEndDate.Value, vTanggalDari, _
                    Me.txtCategory.Text)
                    DTAMaster.FillBy_GroupByBrand_PrimaryCategoryASC(TBLMaster, Me.txtPrimaryCategory.Text)
                ElseIf Me.txtPrimaryParentGroup.Text = "-1" And Me.txtPrimaryCategory.Text = "-1" Then
                    '   c. No FIlter()
                    DTA.FillBy_ascending(TblDummy, Me.dtpEndDate.Value, _
                    vTanggalDari)
                    DTAMaster.FillBy_GroupByBrandQtyASC(TBLMaster)
                End If
            ElseIf Me.cboSort.SelectedIndex = 1 Then
                '2. Descending
                If Me.txtPrimaryParentGroup.Text <> "-1" And Me.txtPrimaryCategory.Text = "-1" Then
                    '   a. Filter By Brand
                    DTA.FillBy_PrimaryBrandDESCENDING(TblDummy, Me.dtpEndDate.Value, vTanggalDari, _
                                                     Me.txtParentGroupCode.Text)
                    DTAMaster.FillBy_GroupByBrand_PrimaryBrandDESC(TBLMaster, Me.txtPrimaryParentGroup.Text)
                ElseIf Me.txtPrimaryParentGroup.Text <> "-1" And Me.txtPrimaryCategory.Text <> "-1" Then
                    '   b. Filter By Category
                    DTA.FillBy_CategoryDESCENDING(TblDummy, dtpEndDate.Value, vTanggalDari, _
                                                Me.txtCategory.Text)
                    DTAMaster.FillBy_GroupByBrand_PrimaryCategoryDESC(TBLMaster, Me.txtPrimaryCategory.Text)
                ElseIf Me.txtPrimaryParentGroup.Text = "-1" And Me.txtPrimaryCategory.Text = "-1" Then
                    '   c. No FIlter
                    DTA.FillBy_descending(TblDummy, Me.dtpEndDate.Value, vTanggalDari)
                    DTAMaster.FillBy_GroupByBrandQtyDESC(TBLMaster)
                End If
            Else
                MsgBox("Please Choose Sort by", MsgBoxStyle.Information, "Information")
            End If

            'buang baris, pada tblmaster yang ada di dalam data table 1, artinya ada transaksi keluar selama periode tsb
            While ij < TBLMaster.Count - 1
                rowMaster = TBLMaster.Rows(ij)
                filter = "[Primary Product] = " + rowMaster.Primary_Product.ToString

                TblDummy.DefaultView.RowFilter = filter
                If TblDummy.DefaultView.Count = 0 Then
                    ij += 1
                Else
                    TBLMaster.Rows.Remove(rowMaster)
                End If
            End While

            'buang baris pada row master, jika, sudah ada pembelian sebelum tanggal mulai
            DTAPurchase.Fill(TblPurchase, TipeTrans, Me.dtpEndDate.Value)
            ij = 0
            While ij <= TBLMaster.Count - 1
                rowMaster = TBLMaster.Rows(ij)
                filter = "[Primary Product] =" + rowMaster.Primary_Product.ToString
                TblPurchase.DefaultView.RowFilter = filter
                If TblPurchase.DefaultView.Count = 0 Then
                    ij += 1
                Else
                    TBLMaster.Rows.Remove(rowMaster)
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAPurchase.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' no grouping
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Cetak1()
        Dim report As New crpLaporanProductRanking_Modif
        Dim DS As New DseLaporanProductRanking_Modif
        Dim objCD As New Core_Module.ClsCekData
        Dim TglDari As Date = objCD.GetFirstDateofTheMonth(Me.dtpStartDate.Value)
        Dim TglSampai As Date = objCD.GetLasDateOfTheMonth(Me.dtpEndDate.Value)
        objCD = Nothing
        Dim brand As String = ""
        Dim category As String = ""
        Dim SortBy As String = ""
        Dim groupby As String = ""
        Dim formatday As Integer = 1
        Dim formatmonth As Integer = 2
        Dim formatyear As Integer = 1
        Dim Period As String = Format(Me.dtpStartDate.Value, "dd MMM yyyy") + " - " + Format(Me.dtpEndDate.Value, "dd MMM yyyy")
        Dim DTA As New DseLaporanProductRanking_ModifTableAdapters.TblCetak1TableAdapter

        Try
            'tentukan menggunakan query apa
            'kombinasi yang mungkin:
            If Me.cboSort.SelectedIndex = 0 Then
                '1. Ascending
                If Me.txtPrimaryParentGroup.Text <> "-1" And Me.txtPrimaryCategory.Text = "-1" Then
                    '   a. Filter By Brand
                    DTA.FillBy_PrimaryBrandASCEN(DS.TblCetak1, Me.dtpEndDate.Value, _
                    Me.dtpStartDate.Value, Me.txtPrimaryParentGroup.Text)
                    brand = Me.txtParentGroupCode.Text
                    category = "All"
                    SortBy = "Ascending"
                    groupby = "No Grouping"

                ElseIf Me.txtPrimaryParentGroup.Text <> "-1" And Me.txtPrimaryCategory.Text <> "-1" Then
                    '   b. Filter By Category
                    DTA.FillBy_CategoryASCEN(DS.TblCetak1, Me.dtpEndDate.Value, Me.dtpStartDate.Value, _
                    Me.txtPrimaryCategory.Text)
                    brand = Me.txtParentGroupCode.Text
                    category = Me.txtCategory.Text
                    SortBy = "Ascending"
                    groupby = "No Grouping"

                ElseIf Me.txtPrimaryParentGroup.Text = "-1" And Me.txtPrimaryCategory.Text = "-1" Then
                    '   c. No FIlter()
                    DTA.FillBy_Ascending(DS.TblCetak1, Me.dtpEndDate.Value, _
                    Me.dtpStartDate.Value)
                    brand = "All"
                    category = "All"
                    SortBy = "Ascending"
                    groupby = "No Grouping"
                End If
            ElseIf Me.cboSort.SelectedIndex = 1 Then
                '2. Descending
                If Me.txtPrimaryParentGroup.Text <> "-1" And Me.txtPrimaryCategory.Text = "-1" Then
                    '   a. Filter By Brand
                    DTA.FillBy_PrimaryBrandDESC(DS.TblCetak1, Me.dtpEndDate.Value, Me.dtpStartDate.Value, _
                   Me.txtPrimaryParentGroup.Text)
                    brand = Me.txtParentGroupCode.Text
                    category = "All"
                    SortBy = "Descending"
                    groupby = "No Grouping"

                ElseIf Me.txtPrimaryParentGroup.Text <> "-1" And Me.txtPrimaryCategory.Text <> "-1" Then
                    '   b. Filter By Category
                    DTA.FillBy_CategoryDESC(DS.TblCetak1, Me.dtpEndDate.Value, Me.dtpStartDate.Value, _
                     Me.txtPrimaryCategory.Text)
                    brand = Me.txtParentGroupCode.Text
                    category = Me.txtCategory.Text
                    SortBy = "Descending"
                    groupby = "No Grouping"

                ElseIf Me.txtPrimaryParentGroup.Text = "-1" And Me.txtPrimaryCategory.Text = "-1" Then
                    '   c. No FIlter
                    DTA.FillBy_Descending(DS.TblCetak1, Me.dtpEndDate.Value, _
                    Me.dtpStartDate.Value)
                    brand = "All"
                    category = "All"
                    SortBy = "Descending"
                    groupby = "No Grouping"
                End If
            Else
                MsgBox("Please Choose Sort by", MsgBoxStyle.Information, "Information")
            End If

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FilterByBrand.ParameterFieldName, brand)
            report.SetParameterValue(report.Parameter_FilterByCategory.ParameterFieldName, category)
            report.SetParameterValue(report.Parameter_AscendingDescending.ParameterFieldName, SortBy)
            report.SetParameterValue(report.Parameter_Quantity.ParameterFieldName, 0)
            report.SetParameterValue(report.Parameter_Tanggal.ParameterFieldName, Period)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, formatday)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, formatmonth)
            report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, formatyear)
            report.SetParameterValue(report.Parameter_GroupBy.ParameterFieldName, groupby)
            report.SetParameterValue(report.Parameter_Judul_Laporan.ParameterFieldName, "PRODUCT RANGKING")
            report.SetParameterValue(report.Parameter_Judul_Qty.ParameterFieldName, "Moving Out")
            Me.crv.ReportSource = report
            Me.crv.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
        End Try
    End Sub

    Private Sub Cetak1_Absolute()
        Dim report As New crpLaporanProductRanking_Modif
        Dim DS As New DseLaporanProductRanking_Modif
        Dim objCD As New Core_Module.ClsCekData
        Dim TglDari As Date = objCD.GetFirstDateofTheMonth(Me.dtpStartDate.Value)
        Dim TglSampai As Date = objCD.GetLasDateOfTheMonth(Me.dtpEndDate.Value)
        objCD = Nothing
        Dim brand As String = ""
        Dim category As String = ""
        Dim SortBy As String = ""
        Dim groupby As String = ""
        Dim formatday As Integer = 1
        Dim formatmonth As Integer = 2
        Dim formatyear As Integer = 1
        Dim Period As String = Format(Me.dtpStartDate.Value, "dd MMM yyyy") + " - " + Format(Me.dtpEndDate.Value, "dd MMM yyyy")
        Dim DTA As New DseLaporanProductRanking_ModifTableAdapters.TblCetak1TableAdapter

        Dim DTAMaster As New DseLaporanProductRanking_ModifTableAdapters.Tbl_AbsoluteTableAdapter
        Dim TBLMaster As New DseLaporanProductRanking_Modif.Tbl_AbsoluteDataTable
        Dim rowMaster As DseLaporanProductRanking_Modif.Tbl_AbsoluteRow
        Dim ij As Double = 0
        Dim filter As String = ""
        Dim row1 As DseLaporanProductRanking_Modif.TblCetak1Row

        Dim DTAPurchase As New DseLaporanProductRanking_ModifTableAdapters.DataTable5TableAdapter
        Dim TBLPurchase As DseLaporanProductRanking_Modif.DataTable5DataTable = DS.DataTable5
        Dim TipeTrans As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase

        Try
            'tentukan menggunakan query apa
            'kombinasi yang mungkin:
            If Me.cboSort.SelectedIndex = 0 Then
                '1. Ascending
                If Me.txtPrimaryParentGroup.Text <> "-1" And Me.txtPrimaryCategory.Text = "-1" Then
                    '   a. Filter By Brand
                    DTA.FillBy_PrimaryBrandASCEN(DS.TblCetak1, Me.dtpEndDate.Value, _
                    Me.dtpStartDate.Value, Me.txtPrimaryParentGroup.Text)
                    brand = Me.txtParentGroupCode.Text
                    category = "All"
                    SortBy = "Ascending"
                    groupby = "No Grouping"
                    DTAMaster.FillBy_NoGrouping_PrimaryBrandASC(TBLMaster, Me.txtPrimaryParentGroup.Text)
                ElseIf Me.txtPrimaryParentGroup.Text <> "-1" And Me.txtPrimaryCategory.Text <> "-1" Then
                    '   b. Filter By Category
                    DTA.FillBy_CategoryASCEN(DS.TblCetak1, Me.dtpEndDate.Value, Me.dtpStartDate.Value, _
                    Me.txtPrimaryCategory.Text)
                    brand = Me.txtParentGroupCode.Text
                    category = Me.txtCategory.Text
                    SortBy = "Ascending"
                    groupby = "No Grouping"
                    DTAMaster.FillBy_NoGroupingPrimaryCategoryASC(TBLMaster, Me.txtPrimaryCategory.Text)
                ElseIf Me.txtPrimaryParentGroup.Text = "-1" And Me.txtPrimaryCategory.Text = "-1" Then
                    '   c. No FIlter()
                    DTA.FillBy_Ascending(DS.TblCetak1, Me.dtpEndDate.Value, _
                    Me.dtpStartDate.Value)
                    brand = "All"
                    category = "All"
                    SortBy = "Ascending"
                    groupby = "No Grouping"
                    DTAMaster.FillBy_NoGroupingQtyASC(TBLMaster)
                End If
            ElseIf Me.cboSort.SelectedIndex = 1 Then
                '2. Descending
                If Me.txtPrimaryParentGroup.Text <> "-1" And Me.txtPrimaryCategory.Text = "-1" Then
                    '   a. Filter By Brand
                    DTA.FillBy_PrimaryBrandDESC(DS.TblCetak1, Me.dtpEndDate.Value, Me.dtpStartDate.Value, _
                   Me.txtPrimaryParentGroup.Text)
                    brand = Me.txtParentGroupCode.Text
                    category = "All"
                    SortBy = "Descending"
                    groupby = "No Grouping"
                    DTAMaster.FillBy_NoGrouping_PrimaryBrandDESC(TBLMaster, Me.txtPrimaryParentGroup.Text)
                ElseIf Me.txtPrimaryParentGroup.Text <> "-1" And Me.txtPrimaryCategory.Text <> "-1" Then
                    '   b. Filter By Category
                    DTA.FillBy_CategoryDESC(DS.TblCetak1, Me.dtpEndDate.Value, Me.dtpStartDate.Value, _
                     Me.txtPrimaryCategory.Text)
                    brand = Me.txtParentGroupCode.Text
                    category = Me.txtCategory.Text
                    SortBy = "Descending"
                    groupby = "No Grouping"
                    DTAMaster.FillBy_NoGrouping_PrimaryCategoryDESC(TBLMaster, Me.txtPrimaryCategory.Text)
                ElseIf Me.txtPrimaryParentGroup.Text = "-1" And Me.txtPrimaryCategory.Text = "-1" Then
                    '   c. No FIlter
                    DTA.FillBy_Descending(DS.TblCetak1, Me.dtpEndDate.Value, _
                    Me.dtpStartDate.Value)
                    brand = "All"
                    category = "All"
                    SortBy = "Descending"
                    groupby = "No Grouping"
                    DTAMaster.FillBy_NoGroupingQtyDESC(TBLMaster)
                End If
            Else
                MsgBox("Please Choose Sort by", MsgBoxStyle.Information, "Information")
            End If

            'buang baris, pada tblmaster yang ada di dalam data table 1
            While ij <= TBLMaster.Count - 1
                rowMaster = TBLMaster.Rows(ij)
                filter = "[Primary Product] = " + rowMaster.Primary_Product.ToString
                DS.TblCetak1.DefaultView.RowFilter = filter
                If DS.TblCetak1.DefaultView.Count = 0 Then
                    ij += 1
                Else
                    TBLMaster.Rows.Remove(rowMaster)
                End If
            End While

            'buang baris pada row master, jika, sudah ada pembelian sebelum tanggal mulai
            DTAPurchase.Fill(DS.DataTable5, TipeTrans, Me.dtpEndDate.Value)
            ij = 0
            While ij <= TBLMaster.Count - 1
                rowMaster = TBLMaster.Rows(ij)
                filter = "[Primary Product] =" + rowMaster.Primary_Product.ToString
                DS.DataTable5.DefaultView.RowFilter = filter
                If DS.DataTable5.DefaultView.Count = 0 Then
                    ij += 1
                Else
                    TBLMaster.Rows.Remove(rowMaster)
                End If
            End While

            'pindahkan row master pada data tabel 1
            DS.TblCetak1.Rows.Clear()
            For Each rowMaster In TBLMaster.Rows
                row1 = DS.TblCetak1.NewTblCetak1Row

                With row1
                    .BrandCode = rowMaster.Brand
                    .Category_Code = rowMaster.Category_Code
                    .Description = rowMaster.Description
                    .Primary_Product = rowMaster.Primary_Product
                    .Product_Code = rowMaster.Product_Code
                    .Qty = rowMaster.Qty
                    .Unit_Code = rowMaster.Unit_Code
                    .Brand_Desc = rowMaster.Brand_Desc

                End With

                DS.TblCetak1.Rows.Add(row1)
            Next

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FilterByBrand.ParameterFieldName, brand)
            report.SetParameterValue(report.Parameter_FilterByCategory.ParameterFieldName, category)
            report.SetParameterValue(report.Parameter_AscendingDescending.ParameterFieldName, SortBy)
            report.SetParameterValue(report.Parameter_Quantity.ParameterFieldName, 0)
            report.SetParameterValue(report.Parameter_Tanggal.ParameterFieldName, Period)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, formatday)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, formatmonth)
            report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, formatyear)
            report.SetParameterValue(report.Parameter_GroupBy.ParameterFieldName, groupby)
            report.SetParameterValue(report.Parameter_Judul_Laporan.ParameterFieldName, "NON MOVING STOCK")
            report.SetParameterValue(report.Parameter_Judul_Qty.ParameterFieldName, "Qty On Hand")

            Me.crv.ReportSource = report
            Me.crv.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
            DTAMaster.Dispose()
            TBLMaster.Dispose()
            DTAPurchase.Dispose()
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmLaporanProductRanking_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.cboSort.SelectedIndex = 0
            Me.cboPeriod.SelectedIndex = 0

            Dim DTA1 As New DseLaporanProductRanking_ModifTableAdapters.DataTable6TableAdapter
            Dim TBL1 As New DseLaporanProductRanking_Modif.DataTable6DataTable

            DTA1.Fill(TBL1, Me.dtpEndDate.Value, Me.dtpStartDate.Value)
            Me.chkAbsoluteData.Checked = True
            'Me.chkAbsoluteData.Checked = False
            Me.cboPeriod.Enabled = True
            Me.chkCost.Checked = True
            Me.MinimumSize = New Point(Me.Size.Width, Me.Size.Height)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkAbsoluteData_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAbsoluteData.CheckedChanged
        Try
            Me.SetCostChk()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetCostChk()
        Try
            If Me.RBGroup.Checked = True Then
                If Me.chkAbsoluteData.Checked = True Then
                    Me.chkCost.Enabled = True
                    Me.dtpEndDate.Enabled = False
                    Me.dtpStartDate.Enabled = False
                    Me.dtpStartDate.Value = Today
                    Me.dtpEndDate.Value = Today
                Else
                    Me.chkCost.Enabled = False
                    Me.chkCost.Checked = False
                    Me.dtpEndDate.Enabled = True
                    Me.dtpStartDate.Enabled = True
                End If
            Else
                Me.chkCost.Enabled = False
                Me.chkCost.Checked = False
                Me.dtpEndDate.Enabled = True
                Me.dtpStartDate.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RBGroup_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBGroup.CheckedChanged
        Try
            Me.SetCostChk()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkCost_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCost.CheckedChanged
        Try
            If Me.chkCost.Checked = True Then
                Me.cboPeriod.Enabled = True
            Else
                Me.cboPeriod.Enabled = False
                Me.cboPeriod.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCategory_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCategory.GotFocus
        Try
            If Me.txtCategory.Enabled = True Then
                Me.txtCategory.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCategory_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCategory.LostFocus
        Try
            If Me.txtCategory.Enabled = True Then
                Me.txtCategory.BackColor = Me.txtPrimaryParentGroup.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtParentGroupCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtParentGroupCode.GotFocus
        Try
            If Me.txtParentGroupCode.Enabled = True Then
                Me.txtParentGroupCode.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtParentGroupCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtParentGroupCode.LostFocus
        Try
            If Me.txtParentGroupCode.Enabled = True Then
                Me.txtParentGroupCode.BackColor = Me.txtPrimaryParentGroup.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class