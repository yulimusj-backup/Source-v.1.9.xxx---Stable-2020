Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmLaporanProductRanking
    Dim ObjCategory As New ClsMasterCategory
    Dim ObjProduct As New clsMasterProduct
    Dim ObjLicense As New Core_Module.ClsLicense

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

    Private Sub txtParentGroupCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtParentGroupCode.KeyDown
        Dim i As New FrmCariMasterProductGroup
        Dim Obj As New ClsMasterProductGroup

        Try
            '' i.BackColorAll = Me.BackcolorAll
            If Me.txtParentGroupCode.ReadOnly = False Then
                If e.KeyValue = 13 Then
                    If Obj.ApakahSudahAda(Me.txtParentGroupCode.Text, 2) = True Then
                        Me.txtPrimaryParentGroup.Text = Obj.GetPrimary(Me.txtParentGroupCode.Text, 2)
                        Me.txtBrandDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryParentGroup.Text)
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
                            Me.txtBrandDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryParentGroup.Text)
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
                        Me.txtBrandDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryParentGroup.Text)
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
                            Me.txtBrandDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryParentGroup.Text)
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
                Me.txtBrandDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryParentGroup.Text)
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

            If RBNoGroup.Checked = True Then
                Me.Cetak1()
            End If
            If RBGroup.Checked = True Then
                Me.Cetak2()
            End If
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
        Dim DS As New DseLaporanProductRanking
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
        Dim DTA As New DseLaporanProductRankingTableAdapters.DataTable1TableAdapter

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
            report.SetParameterValue(report.Parameter_Judul_Laporan.ParameterFieldName, "Product Ranking Report")
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

    ''' <summary>
    ''' no grouping
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Cetak1()
        Dim report As New crpLaporanProductRanking
        Dim DS As New DseLaporanProductRanking
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
        Dim DTA As New DseLaporanProductRankingTableAdapters.TblCetak1TableAdapter
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
                    DTA.FillBy_descending(DS.TblCetak1, Me.dtpEndDate.Value, _
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
            Me.crv.ReportSource = report
            Me.crv.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmLaporanProductRanking_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.cboSort.SelectedIndex = 0
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class