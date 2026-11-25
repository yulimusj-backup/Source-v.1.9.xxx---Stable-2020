Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmLaporanPurchaseOrderBalance
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim ObjSupplier As New clsMasterSupplier
    Dim ObjProduct As New clsMasterProduct
    Dim ObjPO As New ClsTransaksiPO
    Dim ObjUnit As New ClsMasterUnit
    Dim ObjCd As New ClsCekData

    Dim buffer As String = ""
    Dim DummyIndex As Integer = -1
    Dim ProductGroup_Level As Integer = 3
    Dim Brand_Level As Integer = 2
    Dim SubBrand_Level As Integer = 1

    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Dim ObjSettingProd As New Core_Module.ClsTransaksiSetting_ProdGroup
    Dim objItem As New Inventory_Module.ClsTransaksiMaterialPlanningGroupItem
    Dim ObjCategory As New Core_Module.ClsMasterCategory
    Dim ObjProductGroup As New Core_Module.ClsMasterProductGroup
    Dim EraseRowIndex As Integer = 0
    Dim TipeProductGroup As Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup
    Dim ObjLicense As New Core_Module.ClsLicense

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub txtSupplierCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSupplierCode.KeyDown
        If Me.txtSupplierCode.Enabled = False Or Me.txtSupplierCode.ReadOnly = True Then
            Exit Sub
        End If

        Dim i As New FrmCariMasterSupplier
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If Me.ObjSupplier.ApakahSudahAda(Me.txtSupplierCode.Text) = True Then
                    Me.txtPrimarySupplier.Text = Me.ObjSupplier.GetPrimary(Me.txtSupplierCode.Text)
                    Me.txtSupplierName.Text = Me.ObjSupplier.GetSupplierName(Me.txtPrimarySupplier.Text)
                    Me.btnFindSupplier.Focus()
                Else
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.txtSupplierCode.Text <> "" Then
                        i.txtCari.Text = Me.txtSupplierCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimarySupplier > 0 Then
                        Me.txtPrimarySupplier.Text = i.PrimarySupplier
                        Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                        Me.txtSupplierName.Text = Me.ObjSupplier.GetSupplierName(Me.txtPrimarySupplier.Text)
                        Me.btnFindSupplier.Focus()
                    Else
                        Me.txtSupplierCode.Clear()
                        Me.txtSupplierName.Clear()
                        Me.txtPrimarySupplier.Text = "-1"
                        Me.txtSupplierCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtSupplierCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierCode.Leave
        If Me.txtSupplierCode.Enabled = False Or Me.txtSupplierCode.ReadOnly = True Then
            Exit Sub
        End If

        Dim i As New FrmCariMasterSupplier
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.txtSupplierCode.Text <> "" Then
                If Me.ObjSupplier.ApakahSudahAda(Me.txtSupplierCode.Text) Then
                    Dim PrimaryBaru As Double = Me.ObjSupplier.GetPrimary(Me.txtSupplierCode.Text)
                    If PrimaryBaru <> Me.txtPrimarySupplier.Text Then
                        Me.txtPrimarySupplier.Text = Me.ObjSupplier.GetPrimary(Me.txtSupplierCode.Text)
                        Me.txtSupplierName.Text = Me.ObjSupplier.GetSupplierName(Me.txtPrimarySupplier.Text)
                        Me.btnFindSupplier.Focus()
                    End If
                Else
                    If Me.txtSupplierCode.Text <> "" Then
                        i.txtCari.Text = Me.txtSupplierCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimarySupplier > 0 Then
                        Me.txtPrimarySupplier.Text = i.PrimarySupplier
                        Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                        Me.txtSupplierName.Text = Me.ObjSupplier.GetSupplierName(Me.txtPrimarySupplier.Text)
                        Me.btnFindSupplier.Focus()
                    Else
                        Me.txtSupplierCode.Clear()
                        Me.txtSupplierName.Clear()
                        Me.txtPrimarySupplier.Text = "-1"
                        Me.txtSupplierCode.Focus()
                    End If
                End If
            Else
                Me.txtSupplierCode.Clear()
                Me.txtSupplierName.Clear()
                Me.txtPrimarySupplier.Text = "-1"
                Me.btnFindSupplier.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindSupplier_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindSupplier.Click
        Dim i As New FrmCariMasterSupplier
        i.BackColorAll = Me.BackColorAll
        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.txtSupplierCode.Text <> "" Then
                i.txtCari.Text = Me.txtSupplierCode.Text
            End If

            i.ShowDialog()
            If i.PrimarySupplier > 0 Then
                Me.txtPrimarySupplier.Text = i.PrimarySupplier
                Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                Me.txtSupplierName.Text = Me.ObjSupplier.GetSupplierName(Me.txtPrimarySupplier.Text)
                Me.btnFindSupplier.Focus()
            Else
                Me.txtSupplierCode.Clear()
                Me.txtSupplierName.Clear()
                Me.txtPrimarySupplier.Text = "-1"
                Me.txtSupplierCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtProductCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtProductCode.KeyDown
        If Me.TxtProductCode.Enabled = False Or Me.TxtProductCode.ReadOnly = True Then
            Exit Sub
        End If

        Dim i As New frmCariMasterProduct
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If Me.ObjProduct.ApakahSudahAda(Me.TxtProductCode.Text) = True Then
                    Me.TxtPrimaryProduct.Text = Me.ObjProduct.GetPrimary(Me.TxtProductCode.Text)
                    Me.TxtProductDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryProduct.Text)
                    Me.btnFindProductCode1.Focus()
                Else
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.TxtProductCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtProductCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryProduct.Text = i.PrimaryProduct
                        Me.TxtProductCode.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryProduct.Text)
                        Me.TxtProductDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryProduct.Text)
                        Me.btnFindProductCode1.Focus()
                    Else
                        Me.TxtProductCode.Clear()
                        Me.TxtProductDesc.Clear()
                        Me.TxtPrimaryProduct.Text = "-1"
                        Me.TxtProductCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtProductCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtProductCode.Leave
        If Me.TxtProductCode.Enabled = False Or Me.TxtProductCode.ReadOnly = True Then
            Exit Sub
        End If

        Dim i As New frmCariMasterProduct
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtProductCode.Text <> "" Then
                If Me.ObjProduct.ApakahSudahAda(Me.TxtProductCode.Text) Then
                    Dim PrimaryBaru As Double = Me.ObjProduct.GetPrimary(Me.TxtProductCode.Text)
                    If PrimaryBaru <> Me.TxtPrimaryProduct.Text Then
                        Me.TxtPrimaryProduct.Text = Me.ObjProduct.GetPrimary(Me.TxtProductCode.Text)
                        Me.TxtProductDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryProduct.Text)
                        Me.btnFindProductCode1.Focus()
                    End If
                Else
                    If Me.TxtProductCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtProductCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryProduct.Text = i.PrimaryProduct
                        Me.TxtProductCode.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryProduct.Text)
                        Me.TxtProductDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryProduct.Text)
                        Me.btnFindProductCode1.Focus()
                    Else
                        Me.TxtProductCode.Clear()
                        Me.TxtProductDesc.Clear()
                        Me.TxtPrimaryProduct.Text = "-1"
                        Me.TxtProductCode.Focus()
                    End If
                End If
            Else
                Me.TxtProductCode.Clear()
                Me.TxtProductDesc.Clear()
                Me.TxtPrimaryProduct.Text = "-1"
                Me.btnFindProductCode1.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindProductCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindProductCode1.Click
        Dim i As New frmCariMasterProduct
        i.BackColorAll = Me.BackColorAll
        Try
            'i.cboCriteria.SelectedIndex = 0
            If Me.txtSupplierCode.Text <> "" Then
                i.txtCari.Text = Me.TxtProductCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.TxtPrimaryProduct.Text = i.PrimaryProduct
                Me.TxtProductCode.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryProduct.Text)
                Me.TxtProductDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryProduct.Text)
                Me.btnFindProductCode1.Focus()
            Else
                Me.TxtProductCode.Clear()
                Me.TxtProductDesc.Clear()
                Me.TxtPrimaryProduct.Text = "-1"
                Me.TxtProductCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtSupplierCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierCode.GotFocus
        txtSupplierCode.BackColor = Me.BackColorAll
    End Sub

    Private Sub txtSupplierCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierCode.LostFocus
        txtSupplierCode.BackColor = Me.txtPrimarySupplier.BackColor
    End Sub

    Private Sub TxtProductCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtProductCode.GotFocus
        TxtProductCode.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtProductCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtProductCode.LostFocus
        TxtProductCode.BackColor = Me.TxtPrimaryProduct.BackColor
    End Sub


    Private Sub TxtPONumber_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPONumber.GotFocus
        txtPONumber.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtPONumber_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPONumber.LostFocus
        txtPONumber.BackColor = Me.txtPrimaryPO.BackColor
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
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

            If Me.dtpFrom.Value.Date > Me.dtpTo.Value.Date Then
                MsgBox("Wrong date filter", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If


            If Me.rboSupplier.Checked = True Then
                Me.Cetak2(Me.dtpTo.Value, Me.dtpFrom.Value, CDbl(Me.TxtPrimaryProduct.Text), CDbl(Me.txtPrimarySupplier.Text), CDbl(Me.txtPrimaryPO.Text))
            ElseIf Me.rboProduct.Checked = True Then
                Me.Cetak3(Me.dtpTo.Value, Me.dtpFrom.Value, CDbl(Me.TxtPrimaryProduct.Text), CDbl(Me.txtPrimarySupplier.Text), CDbl(Me.txtPrimaryPO.Text))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub Cetak_Obselete()
        'fungsi ini sudah obselete
        'dulu setiap kali generate laporan ini, maka software akan menghitung ulang semuanya, yang baru nilai
        'pob disimpan pada item

        Dim DS As New DseLaporanPurchaseOrderBalance
        Dim report As New CrpLaporanOrderBalanceReportBySupplier
        Try
            'Fillby Tgl Untuk orderbalance_report_po_item_view
            Dim DTA As New DseLaporanPurchaseOrderBalanceTableAdapters.orderbalance_report_po_item_viewTableAdapter
            DTA.FillBy_All2(DS.orderbalance_report_po_item_view)
            DTA.Dispose()

            'Lakukan Pengurangan dengan POC
            If DS.orderbalance_report_po_item_view.Rows.Count > 0 Then
                'Lakukan Penambahan dengan P.Return
                Me.LakukanPenambahanDenganPurchaseReturn1(DS)
                'Lakukan Pengurangan dengan POC
                Me.LakukanPenguranganDenganPOC1(DS)
                'Lakukan Pengurangan dengan GIT
                Me.LakukanPenguranganDenganGIT1(DS)
                'Lakukan Pengurangan dengan GRH
                Me.LakukanPenguranganDenganGRN1(DS)
                'Lakukan Perhitungan Kembali Amount
                Me.HitungKembaliNilaiTotalAmount(DS)
                'Qty Yang Nol Nga usah di munculin 
                Me.EliminasiNol(DS)
            End If

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, Me.txtFormatDay.Text)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.txtFormatDecimal.Text)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, Me.txtFormatMonth.Text)
            report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, Me.txtFormatYears.Text)
            report.SetParameterValue(report.Parameter_SupplierCode.ParameterFieldName, Me.txtSupplierCode.Text)
            report.SetParameterValue(report.Parameter_SupplierName.ParameterFieldName, Me.txtSupplierName.Text)
            Me.crvReport.ReportSource = report
            Me.crvReport.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
        End Try
    End Sub

    Private Sub Cetak2(ByVal EndDate As Date, ByVal StartDate As Date, ByVal PrimaryProduct As Double, ByVal PrimarySupplier As Double, ByVal PrimaryPO As Double)
        Dim DS As New DseLaporanPurchaseOrderBalance
        Dim report As New CrpLaporanOrderBalanceReportBySupplier
        Dim DTA As New DseLaporanPurchaseOrderBalanceTableAdapters.orderbalance_report_po_item_viewTableAdapter
        Dim TBL As DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewDataTable = DS.orderbalance_report_po_item_view
        Dim TBL1 As New DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewDataTable
        Dim TBLFilter As New DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewDataTable
        Dim Row, Row1, RowFilter As DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewRow
        Dim DV As DataRowView

        Dim paramStartDate As Date = Now
        Dim paramEndDate As Date = Now
        Dim paramChkPeriod As Boolean
        Dim filter As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'kosongkan tabel
            TBL.Rows.Clear()
            TBL1.Rows.Clear()
            TBLFilter.Rows.Clear()

            paramStartDate = Me.dtpFrom.Value
            paramEndDate = Me.dtpTo.Value
            paramChkPeriod = Me.chkPeriod.Checked

            '''aul
            'If Me.ChkIgnoreLowTolerance.Checked = False Then
            '    If paramChkPeriod = True Then
            '        If Me.ChkProduct.Checked = True Then
            '            If PrimaryProduct <> -1 And PrimarySupplier <> -1 Then
            '                DTA.FillBy_AllProductPO4(TBL1, EndDate, StartDate, PrimaryProduct, PrimarySupplier)
            '            Else
            '                DTA.FillBy_AllProductPO2(TBL1, EndDate, StartDate, PrimaryProduct)
            '            End If
            '        Else
            '            DTA.FillBy_All(TBL1, EndDate, StartDate)
            '        End If
            '    Else
            '        If Me.ChkProduct.Checked = True Then
            '            If PrimaryProduct <> -1 And PrimarySupplier <> -1 Then
            '                DTA.FillBy_AllProductPO3(TBL1, PrimaryProduct, PrimarySupplier)
            '            ElseIf PrimaryProduct <> -1 Then
            '                DTA.FillBy_AllProductPO1(TBL1, PrimaryProduct)
            '            End If
            '        Else
            '            DTA.FillBy_All2(TBL1)
            '        End If
            '    End If
            'Else
            '    If paramChkPeriod = True Then
            '        If Me.ChkProduct.Checked = True Then
            '            If PrimaryProduct <> -1 And PrimarySupplier <> -1 Then
            '                DTA.FillBy_AllProductPO4_UpTolerance(TBL1, EndDate, StartDate, PrimaryProduct, PrimarySupplier)
            '            Else
            '                DTA.FillBy_AllProductPO2_UpTolerance(TBL1, EndDate, StartDate, PrimaryProduct)
            '            End If
            '        Else
            '            DTA.FillBy_All_UpTolerance(TBL1, EndDate, StartDate)
            '        End If
            '    Else
            '        If Me.ChkProduct.Checked = True Then
            '            If PrimaryProduct <> -1 And PrimarySupplier <> -1 Then
            '                DTA.FillBy_AllProductPO3_UpTolerance(TBL1, PrimaryProduct, PrimarySupplier)
            '            ElseIf PrimaryProduct <> -1 Then
            '                DTA.FillBy_AllProductPO1_UpTolerance(TBL1, PrimaryProduct)
            '            End If
            '        Else
            '            DTA.FillBy_All2_UpTolerance(TBL1)
            '        End If
            '    End If
            'End If
            ''---

            If Me.ChkIgnoreLowTolerance.Checked = False Then
                DTA.FillBy_All2(TBL1)
                DTA.FillBy_All2(TBLFilter)
            Else
                DTA.FillBy_All2_UpTolerance(TBL1)
                DTA.FillBy_All2_UpTolerance(TBLFilter)
            End If



            filter = "[Primary PO] = [Primary PO]"


            If paramChkPeriod = True Then
                filter &= " AND [PO Date] >= '" & Format(CDate(StartDate), "yyyy-MM-dd") & "' And [PO Date] <= '" & Format(CDate(EndDate), "yyyy-MM-dd") & "'"
            End If


            If PrimarySupplier <> -1 Then
                filter &= " AND [Primary Supplier] = " & PrimarySupplier
            End If

            If PrimaryPO <> -1 Then
                filter &= " AND [Primary PO] = " & PrimaryPO
            End If


            Dim FilterUser As String = ""
            'Created By
            If Me.DgvUser.Rows.Count > 0 Then
                For Each rowDGV1 As DataGridViewRow In Me.DgvUser.Rows
                    If Not IsDBNull(rowDGV1.Cells(Me.ColCreatedBy.Index).Value) AndAlso rowDGV1.Cells(Me.ColCreatedBy.Index).Value <> "" AndAlso Not String.IsNullOrEmpty(rowDGV1.Cells(Me.ColCreatedBy.Index).Value) Then
                        If FilterUser = "" Then
                            FilterUser &= " [Created By] = '" & rowDGV1.Cells(Me.ColCreatedBy.Index).Value.ToString & "'"
                        Else
                            FilterUser &= " Or [Created By] = '" & rowDGV1.Cells(Me.ColCreatedBy.Index).Value.ToString & "'"
                        End If
                    End If
                Next
            End If

            If FilterUser <> "" Then
                filter &= " AND (" & FilterUser & ") "
            End If




            If Me.ChkProduct.Checked = True Then
                '======================================================================================================================================
                'Show All Product --> hanya filter primary PO yang punya Product ybs
                '======================================================================================================================================
                Dim ListPrimaryPO As New List(Of Double)
                Dim FilterProduct As String = ""
                FilterProduct = "[Primary PO] = [Primary PO]"
                If PrimaryProduct <> -1 Then
                    FilterProduct &= " And [Primary Product] = " & PrimaryProduct
                End If

                'product group, dkk
                Dim FilterGroup As String = ""
                For Each rowDGV As DataGridViewRow In Me.dgvItemProductGroup.Rows
                    If rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value <> -1 Then
                        TipeProductGroup = Me.ObjSettingProd.GetLevelInt(rowDGV.Cells(Me.ColProductGroup_Level.Index).Value)
                        If FilterGroup = "" Then
                            If TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup Then
                                FilterGroup &= " [Primary Group] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Brand Then
                                FilterGroup &= " [Primary Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.SubBrand Then
                                FilterGroup &= " [Primary Sub Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Category Then
                                FilterGroup &= " [Primary Category] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            Else
                                FilterGroup &= " [primary Product] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            End If
                        Else
                            If TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup Then
                                FilterGroup &= " OR [Primary Group] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Brand Then
                                FilterGroup &= " OR [Primary Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.SubBrand Then
                                FilterGroup &= " OR [Primary Sub Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Category Then
                                FilterGroup &= " OR [Primary Category] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            Else
                                FilterGroup &= " OR [primary Product] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            End If
                        End If
                    End If
                Next

                If FilterGroup <> "" Then
                    FilterProduct &= " AND (" & FilterGroup & ") "
                End If


                TBLFilter.DefaultView.RowFilter = Nothing
                TBLFilter.DefaultView.RowFilter = FilterProduct
                If TBLFilter.DefaultView.Count > 0 Then
                    For Each dvFilter As Data.DataRowView In TBLFilter.DefaultView
                        RowFilter = dvFilter.Row
                        With RowFilter
                            If ListPrimaryPO.Contains(.Primary_PO) = False Then
                                ListPrimaryPO.Add(.Primary_PO)
                            End If
                        End With
                    Next

                    Dim FilterPrimaryPO As String = ""
                    If ListPrimaryPO.Count > 0 Then
                        For intA As Integer = 0 To ListPrimaryPO.Count - 1
                            If FilterPrimaryPO = "" Then
                                FilterPrimaryPO &= " [Primary PO] = " & ListPrimaryPO.Item(intA)
                            Else
                                FilterPrimaryPO &= " OR [Primary PO] = " & ListPrimaryPO.Item(intA)
                            End If
                        Next
                    End If
                    filter &= " AND (" & FilterPrimaryPO & ") "
                Else
                    'jika filter product tidak sesuai maka kosongkan data
                    filter &= " AND [Primary PO] = 0 "
                End If
            Else
                '======================================================================================================================================
                'Show All filter primary Product yang sesuai
                '======================================================================================================================================
                If PrimaryProduct <> -1 Then
                    filter &= " And [Primary Product] = " & PrimaryProduct
                End If

                'product group, dkk
                Dim FilterGroup As String = ""
                For Each rowDGV As DataGridViewRow In Me.dgvItemProductGroup.Rows
                    If rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value <> -1 Then
                        TipeProductGroup = Me.ObjSettingProd.GetLevelInt(rowDGV.Cells(Me.ColProductGroup_Level.Index).Value)
                        If FilterGroup = "" Then
                            If TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup Then
                                FilterGroup &= " [Primary Group] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Brand Then
                                FilterGroup &= " [Primary Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.SubBrand Then
                                FilterGroup &= " [Primary Sub Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Category Then
                                FilterGroup &= " [Primary Category] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            Else
                                FilterGroup &= " [primary Product] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            End If
                        Else
                            If TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup Then
                                FilterGroup &= " OR [Primary Group] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Brand Then
                                FilterGroup &= " OR [Primary Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.SubBrand Then
                                FilterGroup &= " OR [Primary Sub Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Category Then
                                FilterGroup &= " OR [Primary Category] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            Else
                                FilterGroup &= " OR [primary Product] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            End If
                        End If
                    End If
                Next

                If FilterGroup <> "" Then
                    filter &= " AND (" & FilterGroup & ") "
                End If
            End If



            TBL1.DefaultView.RowFilter = Nothing
            TBL1.DefaultView.RowFilter = filter
            If TBL1.DefaultView.Count > 0 Then
                For Each DV In TBL1.DefaultView
                    Row1 = DV.Row
                    Row1.Received = Row1.Total_Qty - Row1.POB_Multi
                    Row1.Persen = (Row1.Received / Row1.Total_Qty) * 100
                    'isi tblCetak dengan tblCetakDummy hasil filteran disini --------------------------------------------------------------------------
                    Row = TBL.Neworderbalance_report_po_item_viewRow
                    Row.Primary_PO = Row1.Primary_PO
                    Row.Primary_Product = Row1.Primary_Product
                    Row.Primary_Supplier = Row1.Primary_Supplier
                    Row.Primary_IPO = Row1.Primary_IPO
                    Row.Unit_Price = Row1.Unit_Price
                    Row.Discount = Row1.Discount
                    Row.Disc_Amount = Row1.Disc_Amount
                    Row.Nett_Price = Row1.Nett_Price
                    Row.Use_Secondary = Row1.Use_Secondary
                    Row.PO_No = Row1.PO_No
                    Row.Po_Date = Row1.Po_Date
                    Row.Currency_Code = Row1.Currency_Code
                    Row.Currency_Description = Row1.Currency_Description
                    Row.Product_Code = Row1.Product_Code
                    Row.Supplier_Code = Row1.Supplier_Code
                    Row.Supplier_Name = Row1.Supplier_Name
                    Row.Payment_Term = Row1.Payment_Term
                    Row.Unit_Code = Row1.Unit_Code
                    Row.Unit_Description = Row1.Unit_Description
                    Row.Total_Qty = Row1.Total_Qty
                    Row.Received = Row1.Received
                    Row.Persen = Row1.Persen
                    Row.Default_Precision = Row1.Default_Precision
                    Row.Decimal_Product = Row1.Decimal_Product

                    'aul
                    Row.Tolerance_Down = Row1.Tolerance_Down
                    Row.Tolerance_Up = Row1.Tolerance_Up
                    Row.Primary_Unit_Multi = Row1.Primary_Unit_Multi
                    Row.Primary_Unit_Default = Row1.Primary_Unit_Default
                    Row.TotalQty_Default = Row1.TotalQty_Default
                    Row.Ratio_Default = Row1.Ratio_Default
                    Row.Ratio_Reference = Row1.Ratio_Reference
                    Row.POB_Multi = Row1.POB_Multi

                    Row.Decimal_Default = ObjProduct.GetAngkaDecimal(Row1.Primary_Product, Row1.Primary_Unit_Default)
                    Row.Unit_Code_Default = ObjUnit.GetUnitCode(Row1.Primary_Unit_Default)
                    Row.Unit_Desc_Default = ""
                    Row.Created_By = Row1.Created_By

                    Row.Primary_Brand = Row1.Primary_Brand
                    Row.Primary_Category = Row1.Primary_Category
                    Row.Primary_Group = Row1.Primary_Group
                    Row.Primary_Sub_Brand = Row1.Primary_Sub_Brand

                    If Me.ChkShowRatio.Checked = True And Row.Primary_Unit_Multi <> Row.Primary_Unit_Default Then
                        Row.Decimal_Multi = ObjProduct.GetAngkaDecimal(Row.Primary_Product, Row.Primary_Unit_Multi)
                        Row.Ratio_Desc = "(" & ObjCd.DigitGrouping(Row.Ratio_Reference, Row.Decimal_Multi) & " " & Row.Unit_Code & "=" & _
                        ObjCd.DigitGrouping(Row.Ratio_Default, Row.Decimal_Default) & " " & Row.Unit_Code_Default & ")"
                    Else
                        Row.Ratio_Desc = ""
                    End If
                    If Row.Primary_Unit_Default <> Row.Primary_Unit_Multi Then
                        If Me.ChkShowRatio.Checked = False Then
                            Row.Decimal_Multi = ObjProduct.GetAngkaDecimal(Row.Primary_Product, Row.Primary_Unit_Multi)
                        End If
                        Row.Decimal_Multi = Row.Decimal_Multi
                    Else
                        Row.Decimal_Multi = Row.Decimal_Default
                    End If

                    If Row.Tolerance_Down <> 0 And Row.Tolerance_Up <> 0 Then
                        Row.Product_Description = Row1.Product_Description & " (-" & ObjCd.DigitGrouping(Row1.Tolerance_Down, Row.Decimal_Multi) & ", +" & ObjCd.DigitGrouping(Row1.Tolerance_Up, Row.Decimal_Multi) & ")"
                    ElseIf Row.Tolerance_Down <> 0 And Row.Tolerance_Up = 0 Then
                        Row.Product_Description = Row1.Product_Description & " (-" & ObjCd.DigitGrouping(Row1.Tolerance_Down, Row.Decimal_Multi) & ")"
                    ElseIf Row.Tolerance_Up <> 0 And Row.Tolerance_Down = 0 Then
                        Row.Product_Description = Row1.Product_Description & " (+" & ObjCd.DigitGrouping(Row1.Tolerance_Up, Row.Decimal_Multi) & ")"
                    Else
                        Row.Product_Description = Row1.Product_Description
                    End If

                    'If Me.ChkIgnoreLowTolerance.Checked = False Then
                    '    Row.Amount = Row1.POB_Multi * Row1.Nett_Price
                    'Else
                    '    Row.Amount = (Row1.POB_Multi + Row1.Tolerance_Up) * Row1.Nett_Price
                    'End If
                    Dim POB_Temp As Decimal = Row1.POB_Multi
                    If Row1.Primary_Unit_Default <> Row1.Primary_Unit_Multi And Row1.price_by_multi_unit = 0 Then
                        POB_Temp = Row1.POB
                    End If
                    Row.Amount = POB_Temp * Row1.Nett_Price
                    If Row.Amount < 0 Then
                        Row.Amount = 0
                    End If
                    '---
                    TBL.Rows.Add(Row)
                Next
            End If

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, Me.txtFormatDay.Text)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.txtFormatDecimal.Text)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, Me.txtFormatMonth.Text)
            report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, Me.txtFormatYears.Text)
            report.SetParameterValue(report.Parameter_SupplierCode.ParameterFieldName, Me.txtSupplierCode.Text)
            report.SetParameterValue(report.Parameter_SupplierName.ParameterFieldName, Me.txtSupplierName.Text)
            report.SetParameterValue(report.Parameter_StartDate.ParameterFieldName, paramStartDate)
            report.SetParameterValue(report.Parameter_EndDate.ParameterFieldName, paramEndDate)
            report.SetParameterValue(report.Parameter_Period.ParameterFieldName, paramChkPeriod)
            Me.crvReport.ReportSource = report
            Me.crvReport.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
        End Try
    End Sub

    Private Sub Cetak3(ByVal EndDate As Date, ByVal StartDate As Date, ByVal PrimaryProduct As Double, ByVal PrimarySupplier As Double, ByVal PrimaryPO As Double)
        Dim DS As New DseLaporanPurchaseOrderBalance
        Dim report As New CrpLaporanOrderBalanceReportByProduct
        Dim DTA As New DseLaporanPurchaseOrderBalanceTableAdapters.orderbalance_report_po_item_viewTableAdapter
        Dim TBL As DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewDataTable = DS.orderbalance_report_po_item_view
        Dim TBL1 As New DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewDataTable
        Dim TBLFilter As New DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewDataTable
        Dim TBLTotal As DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_view_totalcurrDataTable = DS.orderbalance_report_po_item_view_totalcurr
        Dim Row, Row1, RowFilter As DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewRow
        Dim RowTotal As DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_view_totalcurrRow
        Dim DV As DataRowView
        Dim paramStartDate As Date = Now
        Dim paramEndDate As Date = Now
        Dim paramChkPeriod As Boolean
        Dim filter As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'kosongkan tabel
            TBL.Rows.Clear()
            TBL1.Rows.Clear()
            TBLFilter.Rows.Clear()
            paramStartDate = Me.dtpFrom.Value
            paramEndDate = Me.dtpTo.Value
            paramChkPeriod = Me.chkPeriod.Checked

            ''aul
            'If Me.ChkIgnoreLowTolerance.Checked = False Then
            '    If paramChkPeriod = True Then
            '        DTA.FillBy_All(TBL1, EndDate, StartDate)
            '    Else
            '        DTA.FillBy_All2(TBL1)
            '    End If
            'Else
            '    If paramChkPeriod = True Then
            '        DTA.FillBy_All_UpTolerance(TBL1, EndDate, StartDate)
            '    Else
            '        DTA.FillBy_All2_UpTolerance(TBL1)
            '    End If
            'End If
            ''---

            'If Me.ChkIgnoreLowTolerance.Checked = False Then
            '    DTA.FillBy_All2(TBL1)
            'Else
            '    DTA.FillBy_All2_UpTolerance(TBL1)
            'End If

            'filter = "[Primary PO] = [Primary PO]"

            'If PrimarySupplier <> -1 Then
            '    filter &= " AND [Primary Supplier] = " & PrimarySupplier & ""
            'End If

            'If Me.chkPeriod.Checked = True Then
            '    filter &= " AND [Po Date] >= '" & Format(CDate(paramStartDate), "yyyy-MM-dd") & "'"
            '    filter &= " AND [Po Date] <= '" & Format(CDate(paramStartDate), "yyyy-MM-dd") & "'"
            'End If

            'If Me.ChkProduct.Checked = True Then
            '    If PrimaryProduct <> -1 Then
            '        filter &= " AND [Primary Product] = " & PrimaryProduct & ""
            '    End If
            'End If

            ''product group, dkk
            'For Each rowDGV As DataGridViewRow In Me.dgvItemProductGroup.Rows
            '    If rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value <> -1 Then
            '        TipeProductGroup = Me.ObjSettingProd.GetLevelInt(rowDGV.Cells(Me.ColProductGroup_Level.Index).Value)

            '        If TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup Then
            '            filter &= " AND [Primary Group] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
            '        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Brand Then
            '            filter &= " AND [Primary Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
            '        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.SubBrand Then
            '            filter &= " AND [Primary Sub Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
            '        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Category Then
            '            filter &= " AND [Primary Category] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
            '        Else
            '            filter &= " AND [primary Product] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
            '        End If
            '    End If
            'Next

            'Dim FilterUser As String = ""
            ''Created By
            'If Me.DgvUser.Rows.Count > 0 Then
            '    For Each rowDGV1 As DataGridViewRow In Me.DgvUser.Rows
            '        If Not IsDBNull(rowDGV1.Cells(Me.ColCreatedBy.Index).Value) AndAlso rowDGV1.Cells(Me.ColCreatedBy.Index).Value <> "" AndAlso Not String.IsNullOrEmpty(rowDGV1.Cells(Me.ColCreatedBy.Index).Value) Then
            '            If FilterUser = "" Then
            '                FilterUser &= " [Created By] = '" & rowDGV1.Cells(Me.ColCreatedBy.Index).Value.ToString & "'"
            '            Else
            '                FilterUser &= " Or [Created By] = '" & rowDGV1.Cells(Me.ColCreatedBy.Index).Value.ToString & "'"
            '            End If
            '        End If
            '    Next
            'End If


            'If FilterUser <> "" Then
            '    filter &= " AND (" & FilterUser & ") "
            'End If



            If Me.ChkIgnoreLowTolerance.Checked = False Then
                DTA.FillBy_All2(TBL1)
                DTA.FillBy_All2(TBLFilter)
            Else
                DTA.FillBy_All2_UpTolerance(TBL1)
                DTA.FillBy_All2_UpTolerance(TBLFilter)
            End If



            filter = "[Primary PO] = [Primary PO]"


            If paramChkPeriod = True Then
                filter &= " AND [PO Date] >= '" & Format(CDate(StartDate), "yyyy-MM-dd") & "' And [PO Date] <= '" & Format(CDate(EndDate), "yyyy-MM-dd") & "'"
            End If


            If PrimarySupplier <> -1 Then
                filter &= " AND [Primary Supplier] = " & PrimarySupplier
            End If


            If PrimaryPO <> -1 Then
                filter &= " AND [Primary PO] = " & PrimaryPO
            End If

            Dim FilterUser As String = ""
            'Created By
            If Me.DgvUser.Rows.Count > 0 Then
                For Each rowDGV1 As DataGridViewRow In Me.DgvUser.Rows
                    If Not IsDBNull(rowDGV1.Cells(Me.ColCreatedBy.Index).Value) AndAlso rowDGV1.Cells(Me.ColCreatedBy.Index).Value <> "" AndAlso Not String.IsNullOrEmpty(rowDGV1.Cells(Me.ColCreatedBy.Index).Value) Then
                        If FilterUser = "" Then
                            FilterUser &= " [Created By] = '" & rowDGV1.Cells(Me.ColCreatedBy.Index).Value.ToString & "'"
                        Else
                            FilterUser &= " Or [Created By] = '" & rowDGV1.Cells(Me.ColCreatedBy.Index).Value.ToString & "'"
                        End If
                    End If
                Next
            End If

            If FilterUser <> "" Then
                filter &= " AND (" & FilterUser & ") "
            End If







            If Me.ChkProduct.Checked = True Then
                '======================================================================================================================================
                'Show All Product --> hanya filter primary PO yang punya Product ybs
                '======================================================================================================================================
                Dim ListPrimaryPO As New List(Of Double)
                Dim FilterProduct As String = ""
                FilterProduct = "[Primary PO] = [Primary PO]"
                If PrimaryProduct <> -1 Then
                    FilterProduct &= " And [Primary Product] = " & PrimaryProduct
                End If

                'product group, dkk
                Dim FilterGroup As String = ""
                For Each rowDGV As DataGridViewRow In Me.dgvItemProductGroup.Rows
                    If rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value <> -1 Then
                        TipeProductGroup = Me.ObjSettingProd.GetLevelInt(rowDGV.Cells(Me.ColProductGroup_Level.Index).Value)
                        If FilterGroup = "" Then
                            If TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup Then
                                FilterGroup &= " [Primary Group] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Brand Then
                                FilterGroup &= " [Primary Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.SubBrand Then
                                FilterGroup &= " [Primary Sub Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Category Then
                                FilterGroup &= " [Primary Category] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            Else
                                FilterGroup &= " [primary Product] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            End If
                        Else
                            If TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup Then
                                FilterGroup &= " OR [Primary Group] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Brand Then
                                FilterGroup &= " OR [Primary Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.SubBrand Then
                                FilterGroup &= " OR [Primary Sub Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Category Then
                                FilterGroup &= " OR [Primary Category] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            Else
                                FilterGroup &= " OR [primary Product] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            End If
                        End If
                    End If
                Next

                If FilterGroup <> "" Then
                    FilterProduct &= " AND (" & FilterGroup & ") "
                End If


                TBLFilter.DefaultView.RowFilter = Nothing
                TBLFilter.DefaultView.RowFilter = FilterProduct
                If TBLFilter.DefaultView.Count > 0 Then
                    For Each dvFilter As Data.DataRowView In TBLFilter.DefaultView
                        RowFilter = dvFilter.Row
                        With RowFilter
                            If ListPrimaryPO.Contains(.Primary_PO) = False Then
                                ListPrimaryPO.Add(.Primary_PO)
                            End If
                        End With
                    Next

                    Dim FilterPrimaryPO As String = ""
                    If ListPrimaryPO.Count > 0 Then
                        For intA As Integer = 0 To ListPrimaryPO.Count - 1
                            If FilterPrimaryPO = "" Then
                                FilterPrimaryPO &= " [Primary PO] = " & ListPrimaryPO.Item(intA)
                            Else
                                FilterPrimaryPO &= " OR [Primary PO] = " & ListPrimaryPO.Item(intA)
                            End If
                        Next
                    End If
                    filter &= " AND (" & FilterPrimaryPO & ") "
                Else
                    'jika filter product tidak sesuai maka kosongkan data
                    filter &= " AND [Primary PO] = 0 "
                End If
            Else
                '======================================================================================================================================
                'Show All filter primary Product yang sesuai
                '======================================================================================================================================
                If PrimaryProduct <> -1 Then
                    filter &= " And [Primary Product] = " & PrimaryProduct
                End If

                'product group, dkk
                Dim FilterGroup As String = ""
                For Each rowDGV As DataGridViewRow In Me.dgvItemProductGroup.Rows
                    If rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value <> -1 Then
                        TipeProductGroup = Me.ObjSettingProd.GetLevelInt(rowDGV.Cells(Me.ColProductGroup_Level.Index).Value)
                        If FilterGroup = "" Then
                            If TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup Then
                                FilterGroup &= " [Primary Group] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Brand Then
                                FilterGroup &= " [Primary Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.SubBrand Then
                                FilterGroup &= " [Primary Sub Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Category Then
                                FilterGroup &= " [Primary Category] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            Else
                                FilterGroup &= " [primary Product] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            End If
                        Else
                            If TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup Then
                                FilterGroup &= " OR [Primary Group] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Brand Then
                                FilterGroup &= " OR [Primary Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.SubBrand Then
                                FilterGroup &= " OR [Primary Sub Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Category Then
                                FilterGroup &= " OR [Primary Category] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            Else
                                FilterGroup &= " OR [primary Product] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            End If
                        End If
                    End If
                Next

                If FilterGroup <> "" Then
                    filter &= " AND (" & FilterGroup & ") "
                End If
            End If

            TBL1.DefaultView.RowFilter = Nothing
            TBL1.DefaultView.RowFilter = filter
            If TBL1.DefaultView.Count > 0 Then
                For Each DV In TBL1.DefaultView
                    Row1 = DV.Row
                    Row1.Received = Row1.Total_Qty - Row1.POB_Multi
                    Row1.Persen = (Row1.Received / Row1.Total_Qty) * 100

                    'isi tblCetak dengan tblCetakDummy hasil filteran disini --------------------------------------------------------------------------
                    Row = TBL.Neworderbalance_report_po_item_viewRow
                    Row.Primary_PO = Row1.Primary_PO
                    Row.Primary_Product = Row1.Primary_Product
                    Row.Primary_Supplier = Row1.Primary_Supplier
                    Row.Primary_IPO = Row1.Primary_IPO
                    Row.Unit_Price = Row1.Unit_Price
                    Row.Discount = Row1.Discount
                    Row.Disc_Amount = Row1.Disc_Amount
                    Row.Nett_Price = Row1.Nett_Price
                    Row.Use_Secondary = Row1.Use_Secondary
                    Row.PO_No = Row1.PO_No
                    Row.Po_Date = Row1.Po_Date
                    Row.Currency_Code = Row1.Currency_Code
                    Row.Currency_Description = Row1.Currency_Description
                    Row.Product_Code = Row1.Product_Code
                    Row.Product_Description = Row1.Product_Description
                    Row.Supplier_Code = Row1.Supplier_Code
                    Row.Supplier_Name = Row1.Supplier_Name
                    If IsNumeric(Row1.Payment_Term) Then
                        Row.Payment_Term = Row1.Payment_Term & " DAYS"
                    Else
                        Row.Payment_Term = Row1.Payment_Term
                    End If
                    Row.Unit_Code = Row1.Unit_Code
                    Row.Unit_Description = Row1.Unit_Description
                    Row.Total_Qty = Row1.Total_Qty
                    Row.Received = Row1.Received
                    Row.Persen = Row1.Persen
                    Row.Default_Precision = Row1.Default_Precision
                    Row.Decimal_Product = Row1.Decimal_Product

                    'aul
                    Row.Tolerance_Down = Row1.Tolerance_Down
                    Row.Tolerance_Up = Row1.Tolerance_Up
                    Row.Primary_Unit_Multi = Row1.Primary_Unit_Multi
                    Row.Primary_Unit_Default = Row1.Primary_Unit_Default
                    Row.TotalQty_Default = Row1.TotalQty_Default
                    Row.Ratio_Default = Row1.Ratio_Default
                    Row.Ratio_Reference = Row1.Ratio_Reference
                    Row.Decimal_Default = ObjProduct.GetAngkaDecimal(Row1.Primary_Product, Row1.Primary_Unit_Default)
                    Row.Unit_Code_Default = ObjUnit.GetUnitCode(Row1.Primary_Unit_Default)
                    Row.Unit_Desc_Default = ""
                    Row.POB_Multi = Row1.POB_Multi
                    Row.Created_By = Row1.Created_By


                    Row.Primary_Brand = Row1.Primary_Brand
                    Row.Primary_Category = Row1.Primary_Category
                    Row.Primary_Group = Row1.Primary_Group
                    Row.Primary_Sub_Brand = Row1.Primary_Sub_Brand


                    If Me.ChkShowRatio.Checked = True And Row.Primary_Unit_Multi <> Row.Primary_Unit_Default Then
                        Row.Decimal_Multi = ObjProduct.GetAngkaDecimal(Row.Primary_Product, Row.Primary_Unit_Multi)
                        Row.Ratio_Desc = "(" & ObjCd.DigitGrouping(Row.Ratio_Reference, Row.Decimal_Multi) & " " & Row.Unit_Code & "=" & _
                        ObjCd.DigitGrouping(Row.Ratio_Default, Row.Decimal_Default) & " " & Row.Unit_Code_Default & ")"
                    Else
                        Row.Ratio_Desc = ""
                    End If
                    If Row.Primary_Unit_Default <> Row.Primary_Unit_Multi Then
                        If Me.ChkShowRatio.Checked = False Then
                            Row.Decimal_Multi = ObjProduct.GetAngkaDecimal(Row.Primary_Product, Row.Primary_Unit_Multi)
                        End If
                        Row.Decimal_Multi = Row.Decimal_Multi
                    Else
                        Row.Decimal_Multi = Row.Decimal_Default
                    End If

                    If Row.Tolerance_Down <> 0 And Row.Tolerance_Up <> 0 Then
                        Row.POB_Multi_ = "(-" & ObjCd.DigitGrouping(Row1.Tolerance_Down, Row.Decimal_Multi) & ", +" & ObjCd.DigitGrouping(Row1.Tolerance_Up, Row.Decimal_Multi) & ")"
                    ElseIf Row.Tolerance_Down <> 0 And Row.Tolerance_Up = 0 Then
                        Row.POB_Multi_ = "(-" & ObjCd.DigitGrouping(Row1.Tolerance_Down, Row.Decimal_Multi) & ")"
                    ElseIf Row.Tolerance_Up <> 0 And Row.Tolerance_Down = 0 Then
                        Row.POB_Multi_ = "(+" & ObjCd.DigitGrouping(Row1.Tolerance_Up, Row.Decimal_Multi) & ")"
                    Else
                        Row.POB_Multi_ = ""
                    End If

                    'If Me.ChkIgnoreLowTolerance.Checked = False Then
                    '    Row.Amount = Row1.POB_Multi * Row1.Nett_Price
                    'Else
                    '    Row.Amount = (Row1.POB_Multi + Row1.Tolerance_Up) * Row1.Nett_Price
                    'End If
                    Dim POB_Temp As Decimal = Row1.POB_Multi
                    If Row1.Primary_Unit_Default <> Row1.Primary_Unit_Multi And Row1.price_by_multi_unit = 0 Then
                        POB_Temp = Row1.POB
                    End If
                    Row.Amount = POB_Temp * Row1.Nett_Price
                    'Row.Amount = Row1.POB_Multi * Row1.Nett_Price
                    If Row.Amount < 0 Then
                        Row.Amount = 0
                    End If
                    '---
                    TBL.Rows.Add(Row)
                Next
            End If

            For Each Row In TBL.Rows
                TBLTotal.DefaultView.RowFilter = Nothing
                TBLTotal.DefaultView.RowFilter = "[Currency Code] = '" & Row.Currency_Code & "'"
                If TBLTotal.DefaultView.Count > 0 Then
                    For Each DV In TBLTotal.DefaultView
                        RowTotal = DV.Row
                        If IsNumeric(Row.Amount) Then
                            RowTotal.Total += CDec(Row.Amount)
                        End If
                    Next
                Else
                    RowTotal = TBLTotal.Neworderbalance_report_po_item_view_totalcurrRow
                    RowTotal.Currency_Code = Row.Currency_Code
                    RowTotal.Default_Precision = Row.Default_Precision
                    RowTotal.Total = CDec(Row.Amount)
                    TBLTotal.Rows.Add(RowTotal)
                End If
            Next

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_StartDate.ParameterFieldName, paramStartDate)
            report.SetParameterValue(report.Parameter_EndDate.ParameterFieldName, paramEndDate)
            report.SetParameterValue(report.Parameter_Period.ParameterFieldName, paramChkPeriod)
            Me.crvReport.ReportSource = report
            Me.crvReport.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
            TBL.Dispose()
            TBLTotal.Dispose()

        End Try
    End Sub

    Private Sub EliminasiNol(ByRef DS As DseLaporanPurchaseOrderBalance)
        Try
            For i As Integer = 0 To DS.orderbalance_report_po_item_view.Rows.Count - 1
                If i > DS.orderbalance_report_po_item_view.Rows.Count - 1 Then
                    Exit For
                End If
                If DS.orderbalance_report_po_item_view.Rows(i).Item("Total Qty") = 0 Then
                    DS.orderbalance_report_po_item_view.Rows.RemoveAt(i)
                    i -= 1
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LakukanPenguranganDenganPOC1(ByRef DS As DseLaporanPurchaseOrderBalance)
        Dim rowPO As DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewRow
        Dim rowPOC As DseLaporanPurchaseOrderBalance.orderbalance_report_item_poc_viewRow
        Try
            Dim DTAPOC As New DseLaporanPurchaseOrderBalanceTableAdapters.orderbalance_report_item_poc_viewTableAdapter
            DTAPOC.FillBy_TglDariSampai(DS.orderbalance_report_item_poc_view, Me.dtpFrom.Value, Me.dtpTo.Value)
            DTAPOC.Dispose()
            If DS.orderbalance_report_item_poc_view.Rows.Count > 0 Then
                For Each rowPOC In DS.orderbalance_report_item_poc_view.Rows
                    Dim TBL As DataTable = DS.orderbalance_report_po_item_view
                    Dim View As DataView = TBL.DefaultView
                    Dim cmd As String = "[Primary PO] =" + CStr(rowPOC.Primary_PO) + _
                    " and [Discount] =" + CStr(rowPOC.Discount) + _
                    " and [Primary Product] =" + CStr(rowPOC.Primary_Product) + _
                    " and [Unit Price] =" + CStr(rowPOC.Unit_Price) + _
                    " and [Nett Price] =" + CStr(rowPOC.Nett_Price) + _
                    " and [Primary IPO] > -1"
                    Dim rowview As DataRowView

                    View.RowFilter = cmd
                    View.Sort = "[Primary IPO] ASC"

                    For Each rowview In View
                        rowPO = rowview.Row
                        If rowPOC.Total_Qty = 0 Then
                            Exit For
                        End If
                        If rowPO.Total_Qty >= rowPOC.Total_Qty Then
                            rowPO.Total_Qty -= rowPOC.Total_Qty
                            rowPOC.Total_Qty -= rowPOC.Total_Qty
                        Else
                            'Jika Qty PO Lebih Kecil maka Kurangi dengan diri sendiri agar menjadi Nol Kemudian biarkan 
                            'Agar Item Selanjutnya yang di kurangkan
                            rowPOC.Total_Qty -= rowPO.Total_Qty
                            rowPO.Total_Qty -= rowPO.Total_Qty
                        End If
                    Next

                    'View Kedua
                    cmd = "[Primary PO] =" + CStr(rowPOC.Primary_PO) + _
                    " and Discount =" + CStr(rowPOC.Discount) + _
                    " and [Primary Product] =" + CStr(rowPOC.Primary_Product) + _
                    " and [Unit Price] =" + CStr(rowPOC.Unit_Price) + _
                    " and [Nett Price] =" + CStr(rowPOC.Nett_Price) + _
                    " and [Primary IPO] =-1"
                    View.RowFilter = cmd
                    For Each rowview In View
                        rowPO = rowview.Row
                        If rowPOC.Total_Qty = 0 Then
                            Exit For
                        End If
                        If rowPO.Total_Qty >= rowPOC.Total_Qty Then
                            rowPO.Total_Qty -= rowPOC.Total_Qty
                            rowPOC.Total_Qty -= rowPOC.Total_Qty
                        Else
                            'Jika Qty PO Lebih Kecil maka Kurangi dengan diri sendiri agar menjadi Nol Kemudian biarkan 
                            'Agar Item Selanjutnya yang di kurangkan
                            rowPOC.Total_Qty -= rowPO.Total_Qty
                            rowPO.Total_Qty -= rowPO.Total_Qty
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LakukanPenguranganDenganPOC(ByRef DS As DseLaporanPurchaseOrderBalance)
        Dim rowPO As DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewRow
        Dim rowPOC As DseLaporanPurchaseOrderBalance.orderbalance_report_item_poc_viewRow
        Try
            Dim DTAPOC As New DseLaporanPurchaseOrderBalanceTableAdapters.orderbalance_report_item_poc_viewTableAdapter
            DTAPOC.FillBy_TglDariSampai(DS.orderbalance_report_item_poc_view, Me.dtpFrom.Value, Me.dtpTo.Value)
            DTAPOC.Dispose()
            If DS.orderbalance_report_item_poc_view.Rows.Count > 0 Then
                For Each rowPO In DS.orderbalance_report_po_item_view.Rows
                    For Each rowPOC In DS.orderbalance_report_item_poc_view.Rows
                        If rowPO.Primary_PO = rowPOC.Primary_PO And rowPO.Discount = rowPOC.Discount And _
                        rowPO.Primary_Product = rowPOC.Primary_Product And rowPO.Unit_Price = rowPOC.Unit_Price And _
                        rowPO.Nett_Price = rowPOC.Nett_Price Then


                            If rowPO.Total_Qty >= rowPOC.Total_Qty Then
                                rowPO.Total_Qty -= rowPOC.Total_Qty
                                rowPOC.Total_Qty -= rowPOC.Total_Qty
                            Else
                                'Jika Qty PO Lebih Kecil maka Kurangi dengan diri sendiri agar menjadi Nol Kemudian biarkan 
                                'Agar Item Selanjutnya yang di kurangkan
                                rowPO.Total_Qty -= rowPO.Total_Qty
                                rowPOC.Total_Qty -= rowPO.Total_Qty
                            End If
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LakukanPenguranganDenganGIT1(ByRef DS As DseLaporanPurchaseOrderBalance)
        Dim rowPO As DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewRow
        Dim rowGIT As DseLaporanPurchaseOrderBalance.orderbalance_report_goods_in_transit_viewRow
        Try
            Dim DTAGIT As New DseLaporanPurchaseOrderBalanceTableAdapters.orderbalance_report_goods_in_transit_viewTableAdapter
            DTAGIT.FillBy_TglDariSampai(DS.orderbalance_report_goods_in_transit_view, Me.dtpFrom.Value, Me.dtpTo.Value)
            DTAGIT.Dispose()
            If DS.orderbalance_report_goods_in_transit_view.Rows.Count > 0 Then
                For Each rowGIT In DS.orderbalance_report_goods_in_transit_view.Rows
                    Dim TBL As DataTable = DS.orderbalance_report_po_item_view
                    Dim View As DataView = TBL.DefaultView
                    Dim cmd As String = "[Primary PO] =" + CStr(rowGIT.Primary_PO) + _
                    " and [Discount] =" + CStr(rowGIT.Discount) + _
                    " and [Primary Product] =" + CStr(rowGIT.Primary_Product) + _
                    " and [Unit Price] =" + CStr(rowGIT.Unit_Price) + _
                    " and [Nett Price] =" + CStr(rowGIT.Nett_Price) + _
                    " and [Primary IPO] > -1"
                    Dim rowview As DataRowView

                    View.RowFilter = cmd
                    View.Sort = "[Primary IPO] ASC"

                    For Each rowview In View
                        rowPO = rowview.Row
                        If rowGIT.Total_Qty = 0 Then
                            Exit For
                        End If
                        If rowPO.Total_Qty >= rowGIT.Total_Qty Then
                            rowPO.Total_Qty -= rowGIT.Total_Qty
                            rowGIT.Total_Qty -= rowGIT.Total_Qty
                        Else
                            'Jika Qty PO Lebih Kecil maka Kurangi dengan diri sendiri agar menjadi Nol Kemudian biarkan 
                            'Agar Item Selanjutnya yang di kurangkan
                            rowGIT.Total_Qty -= rowPO.Total_Qty
                            rowPO.Total_Qty -= rowPO.Total_Qty
                        End If
                    Next

                    'View Kedua
                    cmd = "[Primary PO] =" + CStr(rowGIT.Primary_PO) + _
                    " and Discount =" + CStr(rowGIT.Discount) + _
                    " and [Primary Product] =" + CStr(rowGIT.Primary_Product) + _
                    " and [Unit Price] =" + CStr(rowGIT.Unit_Price) + _
                    " and [Nett Price] =" + CStr(rowGIT.Nett_Price) + _
                    " and [Primary IPO] =-1"
                    View.RowFilter = cmd
                    For Each rowview In View
                        rowPO = rowview.Row
                        If rowGIT.Total_Qty = 0 Then
                            Exit For
                        End If
                        If rowPO.Total_Qty >= rowGIT.Total_Qty Then
                            rowPO.Total_Qty -= rowGIT.Total_Qty
                            rowGIT.Total_Qty -= rowGIT.Total_Qty
                        Else
                            'Jika Qty PO Lebih Kecil maka Kurangi dengan diri sendiri agar menjadi Nol Kemudian biarkan 
                            'Agar Item Selanjutnya yang di kurangkan
                            rowGIT.Total_Qty -= rowPO.Total_Qty
                            rowPO.Total_Qty -= rowPO.Total_Qty
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LakukanPenguranganDenganGIT(ByRef DS As DseLaporanPurchaseOrderBalance)
        Dim rowPO As DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewRow
        Dim rowGIT As DseLaporanPurchaseOrderBalance.orderbalance_report_goods_in_transit_viewRow
        Try
            Dim DTAGIT As New DseLaporanPurchaseOrderBalanceTableAdapters.orderbalance_report_goods_in_transit_viewTableAdapter
            DTAGIT.FillBy_TglDariSampai(DS.orderbalance_report_goods_in_transit_view, Me.dtpFrom.Value, Me.dtpTo.Value)
            DTAGIT.Dispose()
            If DS.orderbalance_report_goods_in_transit_view.Rows.Count > 0 Then
                For Each rowPO In DS.orderbalance_report_po_item_view.Rows
                    For Each rowGIT In DS.orderbalance_report_goods_in_transit_view.Rows
                        If rowPO.Primary_PO = rowGIT.Primary_PO And rowPO.Discount = rowGIT.Discount And _
                        rowPO.Primary_Product = rowGIT.Primary_Product And rowPO.Unit_Price = rowGIT.Unit_Price And _
                        rowPO.Nett_Price = rowGIT.Nett_Price Then
                            If rowPO.Total_Qty >= rowGIT.Total_Qty Then
                                rowPO.Total_Qty -= rowGIT.Total_Qty
                                rowGIT.Total_Qty -= rowGIT.Total_Qty
                            Else
                                'Jika Qty PO Lebih Kecil maka Kurangi dengan diri sendiri agar menjadi Nol Kemudian biarkan 
                                'Agar Item Selanjutnya yang di kurangkan
                                rowPO.Total_Qty -= rowPO.Total_Qty
                                rowGIT.Total_Qty -= rowPO.Total_Qty
                            End If
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LakukanPenguranganDenganGRN1(ByRef DS As DseLaporanPurchaseOrderBalance)
        Dim rowPO As DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewRow
        Dim rowGRN As DseLaporanPurchaseOrderBalance.orderbalance_report_grnitem_viewRow
        Try
            Dim DTAGRN As New DseLaporanPurchaseOrderBalanceTableAdapters.orderbalance_report_grnitem_viewTableAdapter
            DTAGRN.FillBy_TanggalSampaiDari(DS.orderbalance_report_grnitem_view, Me.dtpTo.Value, Me.dtpFrom.Value)
            DTAGRN.Dispose()
            If DS.orderbalance_report_grnitem_view.Rows.Count > 0 Then
                For Each rowGRN In DS.orderbalance_report_grnitem_view.Rows
                    Dim TBL As DataTable = DS.orderbalance_report_po_item_view
                    Dim View As DataView = TBL.DefaultView
                    Dim cmd As String = "[Primary PO] =" + CStr(rowGRN.Primary_PO) + _
                    " and [Discount] =" + CStr(rowGRN.Discount) + _
                    " and [Primary Product] =" + CStr(rowGRN.Primary_Product) + _
                    " and [Unit Price] =" + CStr(rowGRN.Unit_Price) + _
                    " and [Nett Price] =" + CStr(rowGRN.Nett_Price) + _
                    " and [Primary IPO] > -1"
                    Dim rowview As DataRowView

                    View.RowFilter = cmd
                    View.Sort = "[Primary IPO] ASC"

                    For Each rowview In View
                        rowPO = rowview.Row
                        If rowGRN.Total_Qty = 0 Then
                            Exit For
                        End If
                        If rowPO.Total_Qty >= rowGRN.Total_Qty Then
                            rowPO.Total_Qty -= rowGRN.Total_Qty
                            rowGRN.Total_Qty -= rowGRN.Total_Qty
                        Else
                            'Jika Qty PO Lebih Kecil maka Kurangi dengan diri sendiri agar menjadi Nol Kemudian biarkan 
                            'Agar Item Selanjutnya yang di kurangkan
                            rowGRN.Total_Qty -= rowPO.Total_Qty
                            rowPO.Total_Qty -= rowPO.Total_Qty
                        End If
                    Next

                    'View Kedua
                    cmd = "[Primary PO] =" + CStr(rowGRN.Primary_PO) + _
                    " and Discount =" + CStr(rowGRN.Discount) + _
                    " and [Primary Product] =" + CStr(rowGRN.Primary_Product) + _
                    " and [Unit Price] =" + CStr(rowGRN.Unit_Price) + _
                    " and [Nett Price] =" + CStr(rowGRN.Nett_Price) + _
                    " and [Primary IPO] =-1"
                    View.RowFilter = cmd
                    For Each rowview In View
                        rowPO = rowview.Row
                        If rowGRN.Total_Qty = 0 Then
                            Exit For
                        End If
                        If rowPO.Total_Qty >= rowGRN.Total_Qty Then
                            rowPO.Total_Qty -= rowGRN.Total_Qty
                            rowGRN.Total_Qty -= rowGRN.Total_Qty
                        Else
                            'Jika Qty PO Lebih Kecil maka Kurangi dengan diri sendiri agar menjadi Nol Kemudian biarkan 
                            'Agar Item Selanjutnya yang di kurangkan
                            rowGRN.Total_Qty -= rowPO.Total_Qty
                            rowPO.Total_Qty -= rowPO.Total_Qty
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LakukanPenguranganDenganGRN(ByRef DS As DseLaporanPurchaseOrderBalance)
        Dim rowPO As DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewRow
        Dim rowGRN As DseLaporanPurchaseOrderBalance.orderbalance_report_grnitem_viewRow
        Try
            Dim DTAGRN As New DseLaporanPurchaseOrderBalanceTableAdapters.orderbalance_report_grnitem_viewTableAdapter
            DTAGRN.FillBy_TanggalSampaiDari(DS.orderbalance_report_grnitem_view, Me.dtpTo.Value, Me.dtpFrom.Value)
            DTAGRN.Dispose()
            If DS.orderbalance_report_grnitem_view.Rows.Count > 0 Then
                For Each rowPO In DS.orderbalance_report_po_item_view.Rows
                    For Each rowGRN In DS.orderbalance_report_grnitem_view.Rows
                        If rowPO.Primary_PO = rowGRN.Primary_PO And rowPO.Discount = rowGRN.Discount And _
                        rowPO.Primary_Product = rowGRN.Primary_Product And rowPO.Unit_Price = rowGRN.Unit_Price And _
                        rowPO.Nett_Price = rowGRN.Nett_Price Then
                            If rowPO.Total_Qty >= rowGRN.Total_Qty Then
                                rowPO.Total_Qty -= rowGRN.Total_Qty
                                rowGRN.Total_Qty -= rowGRN.Total_Qty
                            Else
                                'Jika Qty PO Lebih Kecil maka Kurangi dengan diri sendiri agar menjadi Nol Kemudian biarkan 
                                'Agar Item Selanjutnya yang di kurangkan
                                rowPO.Total_Qty -= rowPO.Total_Qty
                                rowGRN.Total_Qty -= rowPO.Total_Qty
                            End If
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LakukanPenambahanDenganPurchaseReturn1(ByRef DS As DseLaporanPurchaseOrderBalance)
        Dim rowPO As DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewRow
        Dim rowPurchaseReturn As DseLaporanPurchaseOrderBalance.orderbalance_report_return_item_viewRow
        Try
            Dim DTAPurchaseReturn As New DseLaporanPurchaseOrderBalanceTableAdapters.orderbalance_report_return_item_viewTableAdapter
            DTAPurchaseReturn.FillBy_TglDariSampai(DS.orderbalance_report_return_item_view, Me.dtpFrom.Value, Me.dtpTo.Value)
            DTAPurchaseReturn.Dispose()
            If DS.orderbalance_report_return_item_view.Rows.Count > 0 Then
                For Each rowPurchaseReturn In DS.orderbalance_report_return_item_view.Rows
                    Dim TBL As DataTable = DS.orderbalance_report_po_item_view
                    Dim View As DataView = TBL.DefaultView
                    Dim cmd As String = ""
                    cmd = "[Primary PO] =" + CStr(rowPurchaseReturn.Primary_PO) + _
                    " and Discount =" + CStr(rowPurchaseReturn.Discount) + _
                    " and [Primary Product] =" + CStr(rowPurchaseReturn.Primary_Product) + _
                    " and [Unit Price] =" + CStr(rowPurchaseReturn.Unit_Price) + _
                    " and [Nett Price] =" + CStr(rowPurchaseReturn.Nett_Price) + _
                    " and [Primary IPO] =-1"
                    Dim rowview As DataRowView
                    View.RowFilter = cmd

                    For Each rowview In View
                        rowPO = rowview.Row
                        If rowPurchaseReturn.Total_Qty = 0 Then
                            Exit For
                        End If
                        rowPO.Total_Qty += rowPurchaseReturn.Total_Qty
                        rowPurchaseReturn.Total_Qty -= rowPurchaseReturn.Total_Qty
                    Next

                    'View Kedua
                    cmd = "[Primary PO] =" + CStr(rowPurchaseReturn.Primary_PO) + _
                    " and [Discount] =" + CStr(rowPurchaseReturn.Discount) + _
                    " and [Primary Product] =" + CStr(rowPurchaseReturn.Primary_Product) + _
                    " and [Unit Price] =" + CStr(rowPurchaseReturn.Unit_Price) + _
                    " and [Nett Price] =" + CStr(rowPurchaseReturn.Nett_Price) + _
                    " and [Primary IPO] > -1"
                    View.RowFilter = cmd
                    View.Sort = "[Primary IPO] DESC"
                    For Each rowview In View
                        rowPO = rowview.Row
                        If rowPurchaseReturn.Total_Qty = 0 Then
                            Exit For
                        End If
                        rowPO.Total_Qty += rowPurchaseReturn.Total_Qty
                        rowPurchaseReturn.Total_Qty -= rowPurchaseReturn.Total_Qty
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LakukanPenambahanDenganPurchaseReturn(ByRef DS As DseLaporanPurchaseOrderBalance)
        Dim rowPO As DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewRow
        Dim rowPurchaseReturn As DseLaporanPurchaseOrderBalance.orderbalance_report_return_item_viewRow
        Try
            Dim DTAPurchaseReturn As New DseLaporanPurchaseOrderBalanceTableAdapters.orderbalance_report_return_item_viewTableAdapter
            DTAPurchaseReturn.FillBy_TglDariSampai(DS.orderbalance_report_return_item_view, Me.dtpFrom.Value, Me.dtpTo.Value)
            DTAPurchaseReturn.Dispose()
            If DS.orderbalance_report_return_item_view.Rows.Count > 0 Then
                For Each rowPO In DS.orderbalance_report_po_item_view.Rows
                    For Each rowPurchaseReturn In DS.orderbalance_report_return_item_view.Rows
                        If rowPO.Primary_PO = rowPurchaseReturn.Primary_PO And rowPO.Discount = rowPurchaseReturn.Discount And _
                        rowPO.Primary_Product = rowPurchaseReturn.Primary_Product And rowPO.Unit_Price = rowPurchaseReturn.Unit_Price And _
                        rowPO.Nett_Price = rowPurchaseReturn.Nett_Price Then
                            rowPO.Total_Qty += rowPurchaseReturn.Total_Qty
                            rowPurchaseReturn.Total_Qty -= rowPurchaseReturn.Total_Qty
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungKembaliNilaiTotalAmount(ByRef DS As DseLaporanPurchaseOrderBalance)
        Dim rowPO As DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewRow
        Try
            For Each rowPO In DS.orderbalance_report_po_item_view.Rows
                rowPO.Amount = rowPO.Total_Qty * rowPO.Nett_Price
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkPeriod_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPeriod.CheckedChanged
        Try
            If Me.chkPeriod.Checked = True Then
                Me.dtpFrom.Enabled = True
                Me.dtpTo.Enabled = True
            Else
                Me.dtpFrom.Enabled = False
                Me.dtpTo.Enabled = False
                Me.dtpFrom.Value = Today
                Me.dtpTo.Value = Today
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPrimaryProduct_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtPrimaryProduct.TextChanged
        Try
            If Me.TxtPrimaryProduct.Text = -1 Then
                Me.ChkProduct.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboSupplier.CheckedChanged
        Try
            If Me.rboSupplier.Checked = False Then
                Me.ChkProduct.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkProduct_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkProduct.Click
        Try
            If Me.rboSupplier.Checked = True Then
                If Me.TxtPrimaryProduct.Text = -1 AndAlso (Me.dgvItemProductGroup.RowCount > 0 AndAlso _
                                                          CDec(Me.dgvItemProductGroup.Rows(0).Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value) <= 0) Then
                    Me.ChkProduct.Checked = False
                    MsgBox("Product Or " & Me.dgvItemProductGroup.Rows(0).Cells(Me.ColProductGroup_Level.Index).Value & " cannot be an empty string ", MsgBoxStyle.Information, "Information")
                    Me.TxtProductCode.Focus()
                End If
            Else
                Me.ChkProduct.Checked = False
                MsgBox("Only used on the group by supplier", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub dgvItemProductGroup_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemProductGroup.CellBeginEdit
        Dim rowDGV As DataGridViewRow

        Try
            'yang bisa diedit:
            '1. jika baris baru
            If e.RowIndex = Me.dgvItemProductGroup.Rows.Count - 1 Then
                If e.ColumnIndex = Me.ColProductGroup_Level.Index Then
                    'pastikan baris diatasnya sudah beres dahulu
                    If e.RowIndex > 0 Then
                        rowDGV = Me.dgvItemProductGroup.Rows(e.RowIndex - 1)
                        If rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value.ToString.Trim = "" Then
                            MsgBox("Please specify " + rowDGV.Cells(Me.ColProductGroup_Level.Index).Value.ToString + _
                            " code on the previous row.", MsgBoxStyle.Information, "Information")
                            Me.dgvItemProductGroup.ClearSelection()
                            rowDGV.Selected = True
                            e.Cancel = True
                            Exit Sub
                        Else
                            Me.buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                        End If
                    Else
                        Me.buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    e.Cancel = True
                End If
            Else
                If Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = -1 Then
                    '2. jika bukan baris baru dan primary =-1
                    'yang bisa diedit hanya code product group dan level
                    If e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Or e.ColumnIndex = Me.ColProductGroup_Level.Index Then
                        Me.buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                Else
                    '3. jika bukan baris baru dan primary <> -1
                    'yang bisa diedit: code, remarks
                    If e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Or _
                    e.ColumnIndex = Me.ColProductGroup_Remarks.Index Then
                        Me.buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub dgvItemProductGroup_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemProductGroup.CellEndEdit
        Dim Level As Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand

        Try
            If IsNothing(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True _
                       Or Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColProductGroup_Remarks.Index Then
                    'Remarks
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = ""
                ElseIf e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Then
                    'kalo kode product group dihapus, maka hapus satu bari semuanya
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                    Exit Sub
                End If
            Else
                'level
                If e.ColumnIndex = Me.ColProductGroup_Level.Index Then
                    'do nothing

                ElseIf e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Then
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If

                    'cek ke master
                    Level = Me.objItem.GetProductLevel(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
                    If Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                        Me.EndEditProductGroup(e, Me.ProductGroup_Level)
                    ElseIf Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                        Me.EndEditProductGroup(e, Me.Brand_Level)
                    ElseIf Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                        Me.EndEditProductGroup(e, Me.SubBrand_Level)
                    ElseIf Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                        Me.EndEditCategory(e)
                    Else
                        Me.EndEditProduct(e)
                    End If
                ElseIf e.ColumnIndex = Me.ColProductGroup_Remarks.Index Then
                    'remarks
                    'pastikan ditrim dan  max 250 karakter
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If
                    'tidak boleh ada karakter aneh2
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("'") Or _
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("%") Or _
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("*") Or _
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("#") Then
                        MsgBox("Please avoid using the following character: ', *, %, #.", MsgBoxStyle.Information, "Information")
                        Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Function GetIndex_ItemProductGroup(ByVal vCode As String, ByVal vRefIndex As Integer, ByVal vProductGroupLevelStr As String) As Integer
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItemProductGroup.Rows
                If rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value = vCode And _
                rowDGV.Index <> vRefIndex And _
                rowDGV.Cells(Me.ColProductGroup_Level.Index).Value = vProductGroupLevelStr Then
                    Return rowDGV.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub EndEditProductGroup(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, ByVal vProductGroup_Level As Integer)
        Try
            If Me.ObjProductGroup.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, vProductGroup_Level) = True Then
                '1. kode product group ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProductGroup.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, vProductGroup_Level)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProductGroup.GetGroupDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode product group yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.FrmCariMasterProductGroup
                ij.ProductGroupLevel = vProductGroup_Level
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryProductGroup <> -1 Then
                    '2.a. product group yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryProductGroup
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjProductGroup.GetGroupCode(ij.PrimaryProductGroup)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProductGroup.GetGroupDesc(ij.PrimaryProductGroup)
                ElseIf ij.PrimaryProductGroup = -1 Then
                    '2.b. kode akun yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    If Trim(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If
            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If Buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProductGroup.GetPrimary(Me.buffer, vProductGroup_Level)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProductGroup.GetGroupDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditCategory(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If Me.ObjCategory.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                '1. kode category ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjCategory.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode category yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.frmCariMasterCategory
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryCategory <> -1 Then
                    '2.a. category yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryCategory
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCategory.GetCategoryCode(ij.PrimaryCategory)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(ij.PrimaryCategory)
                ElseIf ij.PrimaryCategory = -1 Then
                    '2.b. kode category yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If

            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If Buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjCategory.GetPrimary(Me.buffer)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub EndEditProduct(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If Me.ObjProduct.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                '1. kode product ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode product yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.frmCariMasterProduct
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryProduct <> -1 Then
                    '2.a. product yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryProduct
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjProduct.GetProductCode(ij.PrimaryProduct)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(ij.PrimaryProduct)
                ElseIf ij.PrimaryProduct = -1 Then
                    '2.b. kode product yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If

            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If Buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimary(Me.buffer)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemProductGroup.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgvItemProductGroup.Rows(Me.dgvItemProductGroup.Rows.Count - 1)
            rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value = ""
            rowDGV.Cells(Me.ColProductGroup_Desc.Index).Value = ""
            rowDGV.Cells(Me.ColProductGroup_Level.Index).Value = Me.objItem.GetProductLevelStr(Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup)
            rowDGV.Cells(Me.ColProductGroup_No.Index).Value = rowDGV.Index + 1
            rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value = -1
            rowDGV.Cells(Me.ColProductGroup_Remarks.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItemProductGroup.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub IsiDGV_ItemGroup()
        Dim ObjSettingPG As New Core_Module.ClsTransaksiSetting_ProdGroup

        Try
            ObjSettingPG.IsikanDGV(Me.UserCode, Core_Module.ClsTransaksiSetting_ProdGroup.EnumTipeTransaksi.LaporanPurchaseRequest, Me.dgvItemProductGroup, _
            Me.ColProductGroup_No.Index, Me.ColProductGroup_Level.Index, Me.ColProductGroup_CodeNo.Index, Me.ColProductGroup_Desc.Index, _
            Me.ColProductGroup_Remarks.Index, Me.ColProductGroup_PrimaryAcuan.Index)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjSettingPG = Nothing
        End Try
    End Sub

    Private Sub dgvItemProductGroup_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItemProductGroup.DataError
        'do nothing -> dibuat untuk handle data error event, jangan dihapuskan
    End Sub


    Private Sub DgvUser_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvUser.CellBeginEdit
        Try
            If Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value <> -1 Then
                If e.ColumnIndex = Me.ColCreatedBy.Index Then
                    Me.buffer = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColCreatedBy.Index Then
                    Me.buffer = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem(ByVal vKodeUser As String, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvUser.Rows.Count - 2
                RowDGV = Me.DgvUser.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColCreatedBy.Index).Value = vKodeUser Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DgvUser_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvUser.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvUser.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim KodeUser As String = ""
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColCreatedBy.Index Then
                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjUser.ApakahSudahAda(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            KodeUser = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                            index = Me.GetIndexItem(KodeUser, e.RowIndex)
                            If index < 0 Then
                                Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = Me.ObjUser.GetPrimary(Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value)
                                Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value
                                Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.ObjUser.GetNama(Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BgwUser.RunWorkerAsync()
                                    Else
                                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterUser
                            ij.TxtCari.Text = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.backColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.Kodeuser <> "" Then
                                index = Me.GetIndexItem(ij.Kodeuser, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = ij.Kodeuser
                                    Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.ObjUser.GetNama(ij.Kodeuser)
                                    Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = Me.ObjUser.GetPrimary(ij.Kodeuser)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BgwUser.RunWorkerAsync()
                                        Else
                                            Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Exit Sub
                                    Else
                                        Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = ij.Kodeuser
                                        Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.ObjUser.GetNama(ij.Kodeuser)
                                        Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = Me.ObjUser.GetPrimary(ij.Kodeuser)
                                    End If
                                End If
                            ElseIf ij.Kodeuser = "" Then
                                If e.ColumnIndex = Me.ColCreatedBy.Index Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        If Me.DgvUser.Rows.Count - 1 <> e.RowIndex Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BgwUser.RunWorkerAsync()
                                        End If
                                    Else
                                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                Else
                                    Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.DgvUser)
                    End If
                Else
                    Me.EraseRowIndex = e.RowIndex
                    Me.BgwUser.RunWorkerAsync()
                End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvUser_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvUser.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            ' If Me.TxtNoPO.Enabled = True Then
            rowDGV = Me.DgvUser.Rows(Me.DgvUser.RowCount - 1)
            With rowDGV
                ' .Cells(Me.ColN.Index).Value = rowDGV.Index + 1
                .Cells(Me.ColCreatedBy.Index).Value = ""
                .Cells(Me.ColCreatedBy_Name.Index).Value = ""
                .Cells(Me.ColCreatedBy_Primary.Index).Value = "-1"
            End With
            ' End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwUser_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwUser.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwUser_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwUser.RunWorkerCompleted
        Try
            Me.DgvUser.Rows.RemoveAt(Me.EraseRowIndex)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub BackgroundWorker3_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Try
            If Me.DummyIndex <> -1 Then
                Me.dgvItemProductGroup.Rows.RemoveAt(Me.DummyIndex)
                DummyIndex = -1
                Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanPurchaseOrderBalance_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Me.TxtPrimaryProduct.Visible = True
        'Me.txtPrimarySupplier.Visible = True
        Me.dgvItemProductGroup.Rows.Clear()
        Me.DgvUser.Rows.Clear()
        Me.ColProductGroup_Remarks.Visible = False
    End Sub

    Private Sub txtPONumber_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPONumber.KeyDown
        If Me.txtPONumber.Enabled = False Or Me.txtPONumber.ReadOnly = True Then
            Exit Sub
        End If

        Dim i As New FrmCariTransaksiPO
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If Me.ObjPO.ApakahSudahAda(Me.txtPONumber.Text) = True Then
                    Me.txtPrimaryPO.Text = Me.ObjPO.GetPrimary(Me.txtPONumber.Text)
                    Me.btnFindPO.Focus()
                Else
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.txtPONumber.Text <> "" Then
                        i.txtCari.Text = Me.txtPONumber.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryPO > 0 Then
                        Me.txtPrimaryPO.Text = i.PrimaryPO
                        Me.txtPONumber.Text = Me.ObjPO.GetNoTransaksiPO(Me.txtPrimaryPO.Text)
                        Me.btnFindPO.Focus()
                    Else
                        Me.txtPONumber.Clear()
                        Me.txtPrimaryPO.Text = "-1"
                        Me.txtPONumber.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtPONumber_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPONumber.Leave
        If Me.txtPONumber.Enabled = False Or Me.txtPONumber.ReadOnly = True Then
            Exit Sub
        End If

        Dim i As New FrmCariTransaksiPO
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.txtPONumber.Text <> "" Then
                If Me.ObjPO.ApakahSudahAda(Me.txtPONumber.Text) Then
                    Dim PrimaryBaru As Double = Me.ObjPO.GetPrimary(Me.txtPONumber.Text)
                    If PrimaryBaru <> Me.txtPrimaryPO.Text Then
                        Me.txtPrimaryPO.Text = Me.ObjPO.GetPrimary(Me.txtPONumber.Text)
                        Me.btnFindPO.Focus()
                    End If
                Else
                    If Me.txtPONumber.Text <> "" Then
                        i.txtCari.Text = Me.txtPONumber.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryPO > 0 Then
                        Me.txtPrimaryPO.Text = i.PrimaryPO
                        Me.txtPONumber.Text = Me.ObjPO.GetNoTransaksiPO(Me.txtPrimaryPO.Text)
                        Me.btnFindPO.Focus()
                    Else
                        Me.txtPONumber.Clear()
                        Me.txtPrimaryPO.Text = "-1"
                        Me.txtPONumber.Focus()
                    End If
                End If
            Else
                Me.txtPONumber.Clear()
                Me.txtPrimaryPO.Text = "-1"
                Me.btnFindPO.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindPO_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindPO.Click
        Dim i As New FrmCariTransaksiPO
        i.BackColorAll = Me.BackColorAll
        Try
            'i.cboCriteria.SelectedIndex = 0
            If Me.txtPONumber.Text <> "" Then
                i.txtCari.Text = Me.txtPONumber.Text
            End If

            i.ShowDialog()
            If i.PrimaryPO > 0 Then
                Me.txtPrimaryPO.Text = i.PrimaryPO
                Me.txtPONumber.Text = Me.ObjPO.GetNoTransaksiPO(Me.txtPrimaryPO.Text)
                Me.btnFindPO.Focus()
            Else
                Me.txtPONumber.Clear()
                Me.txtPrimaryPO.Text = "-1"
                Me.txtPONumber.Focus()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub
End Class