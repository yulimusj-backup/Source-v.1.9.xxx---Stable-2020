Option Explicit On
Option Strict Off

Public Class FrmLaporanPurchaseOrderReportByDateDetail
    Public BAckColorAll As Color = Color.FromArgb(170, 255, 190)
    Public UserCode As String = ""

    Dim ObjCd As New Core_Module.ClsCekData

    Dim DS As New DsePurchaseOrderReportByDateDetail
    Dim report As New CrpLaporanPurchaseOrderReportDetail
    Dim buffer As String = ""
    Dim DummyIndex As Integer = -1
    Dim ProductGroup_Level As Integer = 3
    Dim Brand_Level As Integer = 2
    Dim SubBrand_Level As Integer = 1

    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjSupplier As New Core_Module.clsMasterSupplier
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Dim ObjSettingProd As New Core_Module.ClsTransaksiSetting_ProdGroup
    Dim objItem As New Inventory_Module.ClsTransaksiMaterialPlanningGroupItem
    Dim ObjCategory As New Core_Module.ClsMasterCategory
    Dim ObjProductGroup As New Core_Module.ClsMasterProductGroup
    Dim EraseRowIndex As Integer = 0
    Dim ObjLicense As New Core_Module.ClsLicense

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
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
                MsgBox("Starting period must be smaller than ending period", MsgBoxStyle.Information, "Information")
                Me.dtpFrom.Focus()
                Exit Sub
            End If

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.CetakDetail(Me.dtpTo.Value, Me.dtpFrom.Value, CDbl(Me.txtPrimaryProduct.Text), CDbl(Me.txtPrimarySupplier.Text), CDbl(Me.txtPrimaryCurrency.Text))

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CetakDetail(ByVal EndDate As Date, ByVal StartDate As Date, ByVal PrimaryProduct As Double, ByVal PrimarySupplier As Double, ByVal PrimaryCurrency As Double)
        Dim Report As New CrpLaporanPurchaseOrderReportDetail
        Dim Report2 As New CrpLaporanPurchaseOrderReportDetailGroup
        Dim DS As New DsePurchaseOrderReportByDateDetail
        Dim Decimal_Multi As Integer = 0

        Dim DTACetakDummy As New DsePurchaseOrderReportByDateDetailTableAdapters.tblCetakTableAdapter
        Dim DTADelivery As New DsePurchaseOrderReportByDateDetailTableAdapters.tbltrans_po_itemdeliveryTableAdapter
        Dim TBLCetak As DsePurchaseOrderReportByDateDetail.tblCetakDataTable = DS.tblCetak
        Dim TBLCetakDummy As New DsePurchaseOrderReportByDateDetail.tblCetakDataTable
        Dim TBLFilter As New DsePurchaseOrderReportByDateDetail.tblCetakDataTable
        Dim TBLTotalCurr As DsePurchaseOrderReportByDateDetail.tblTotalCurrDataTable = DS.tblTotalCurr
        Dim TBLDelivery As New DsePurchaseOrderReportByDateDetail.tbltrans_po_itemdeliveryDataTable
        Dim rowDelivery As DsePurchaseOrderReportByDateDetail.tbltrans_po_itemdeliveryRow
        Dim RowCetak As DsePurchaseOrderReportByDateDetail.tblCetakRow
        Dim RowFilter As DsePurchaseOrderReportByDateDetail.tblCetakRow
        Dim RowCetakDummy As DsePurchaseOrderReportByDateDetail.tblCetakRow
        Dim RowTotalCurr As DsePurchaseOrderReportByDateDetail.tblTotalCurrRow
        Dim DV As DataRowView
        Dim DV1 As DataRowView
        Dim DV2 As DataRowView

        Dim TipeProductGroup As Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup
        Dim paramStartDate As Date = Now
        Dim paramEndDate As Date = Now
        Dim filter As String = ""
        Try
            DTACetakDummy.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'kosongkan tabel
            TBLCetak.Rows.Clear()
            TBLCetakDummy.Rows.Clear()

            paramStartDate = Me.dtpFrom.Value
            paramEndDate = Me.dtpTo.Value

            'Isi tblCetakDummy ----------------------------------------------------------------------------------------------------------------
            'If Me.ChkAllProductPO.Checked = True Then
            '    DTACetakDummy.FillBy_AllProductPO(TBLCetakDummy, Me.dtpFrom.Value, Me.dtpTo.Value, PrimaryProduct)
            'Else
            DTACetakDummy.Fill(TBLCetakDummy, Me.dtpFrom.Value, Me.dtpTo.Value)
            DTACetakDummy.Fill(TBLFilter, Me.dtpFrom.Value, Me.dtpTo.Value)
            DTADelivery.Fill(TBLDelivery, Me.dtpFrom.Value, Me.dtpTo.Value)
            'End If

            'Filter tblCetakDummy disini ------------------------------------------------------------------------------------------------------
            If Me.ChkOS.Checked = True Then
                filter = "[Primary PO] = [Primary PO] AND [OS] > 0"
            Else
                filter = "[Primary PO] = [Primary PO]"
            End If

            'If PrimaryProduct <> -1 Then
            '    If Me.ChkAllProductPO.Checked = False Then
            '        filter &= " AND [Primary Product] = " & PrimaryProduct & ""
            '    Else
            '        filter = filter
            '    End If
            'Else
            '    filter = filter
            'End If

            If String.IsNullOrEmpty(Me.txtRemarks.Text) Then


            End If

            If PrimarySupplier <> -1 Then
                filter &= " AND [Primary Supplier] = " & PrimarySupplier & ""
            End If
            If PrimaryCurrency <> -1 Then
                filter &= " AND [Primary Currency] = " & PrimaryCurrency & ""
            End If

            ''product group, dkk
            'Dim FilterGroup As String = ""
            'For Each rowDGV As DataGridViewRow In Me.dgvItemProductGroup.Rows
            '    If rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value <> -1 Then
            '        TipeProductGroup = Me.ObjSettingProd.GetLevelInt(rowDGV.Cells(Me.ColProductGroup_Level.Index).Value)
            '        If FilterGroup = "" Then
            '            If TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup Then
            '                FilterGroup &= " [Primary Group] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
            '            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Brand Then
            '                FilterGroup &= " [Primary Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
            '            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.SubBrand Then
            '                FilterGroup &= " [Primary Sub Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
            '            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Category Then
            '                FilterGroup &= " [Primary Category] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
            '            Else
            '                FilterGroup &= " [primary Product] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
            '            End If
            '        Else
            '            If TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup Then
            '                FilterGroup &= " Or [Primary Group] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
            '            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Brand Then
            '                FilterGroup &= " Or [Primary Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
            '            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.SubBrand Then
            '                FilterGroup &= " Or [Primary Sub Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
            '            ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Category Then
            '                FilterGroup &= " Or [Primary Category] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
            '            Else
            '                FilterGroup &= " Or [primary Product] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
            '            End If
            '        End If
            '    End If
            'Next

            'If FilterGroup <> "" Then
            '    filter &= " AND (" & FilterGroup & ") "
            'End If




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


            'All Purchase Order
            'Authorized By Purchase Staff
            'Authorized By Purchase Manager
            'Authorized By Director
            'Unauthorized(Only)

            If Me.CboStatusPO.SelectedIndex = ClsTransaksiPO.EnumStatusPO.AuthorizeLvl_1 Then
                filter &= " AND [Authorization Level] = 1 "
            ElseIf Me.CboStatusPO.SelectedIndex = ClsTransaksiPO.EnumStatusPO.AuthorizeLvl_2 Then
                filter &= " AND [Authorization Level] = 2 "
            ElseIf Me.CboStatusPO.SelectedIndex = ClsTransaksiPO.EnumStatusPO.AuthorizeLvl_3 Then
                filter &= " AND [Authorization Level] = 3 "
            ElseIf Me.CboStatusPO.SelectedIndex = ClsTransaksiPO.EnumStatusPO.UnauthorizedOnly Then
                filter &= " AND [Authorization Level] = 0 "
            End If



            If Me.ChkAllProductPO.Checked = True Then
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



            '@hdar'#tambah filter remarks header dan remarks item
            Dim rH As String = Me.txtRemarks.Text.ToUpper.ToString.Trim.Replace(" ", "").Replace(vbCrLf, "")
            If String.IsNullOrEmpty(rH) = False Then
                For hj As Integer = TBLCetakDummy.Rows.Count - 1 To 0 Step -1
                    RowCetakDummy = TBLCetakDummy.Rows(hj)
                    Dim rH2 As String = RowCetakDummy.Remarks.ToUpper.ToString.Trim.Replace(" ", "").Replace(vbCrLf, "")
                    If String.IsNullOrEmpty(rH2) = False Then
                        If rH.Contains(rH2) = True Or rH2.Contains(rH) = True Then
                        Else
                            TBLCetakDummy.Rows.Remove(RowCetakDummy)
                        End If
                    Else
                        TBLCetakDummy.Rows.Remove(RowCetakDummy)
                    End If
                Next
            End If

            Dim rI As String = Me.txtItemRemarks.Text.ToUpper.ToString.Trim.Replace(" ", "").Replace(vbCrLf, "")
            If String.IsNullOrEmpty(Me.txtItemRemarks.Text.ToString) = False Then
                For hj As Integer = TBLCetakDummy.Rows.Count - 1 To 0 Step -1
                    RowCetakDummy = TBLCetakDummy.Rows(hj)
                    Dim rI2 As String = RowCetakDummy.Remarks_Item.ToString.ToUpper.Trim.Replace(" ", "").Replace(vbCrLf, "")
                    If String.IsNullOrEmpty(rI2) = False Then
                        If rI2.Contains(rI) = True Or rI.Contains(rI2) = True Then
                        Else
                            TBLCetakDummy.Rows.Remove(RowCetakDummy)
                        End If
                    Else
                        TBLCetakDummy.Rows.Remove(RowCetakDummy)
                    End If
                Next
            End If
            '--

            'MsgBox(filter)
            TBLCetakDummy.DefaultView.RowFilter = Nothing
            TBLCetakDummy.DefaultView.RowFilter = filter
            If Me.rboPONumber.Checked = False Then
                TBLCetakDummy.DefaultView.Sort = "[Product Code] "
            End If
            If TBLCetakDummy.DefaultView.Count > 0 Then
                For Each DV In TBLCetakDummy.DefaultView
                    RowCetakDummy = DV.Row
                    RowCetakDummy.Received = RowCetakDummy.Ordered - RowCetakDummy.OS
                    RowCetakDummy.Persen = (RowCetakDummy.Received / RowCetakDummy.Ordered) * 100

                    'isi tblCetak dengan tblCetakDummy hasil filteran disini --------------------------------------------------------------------------
                    RowCetak = TBLCetak.NewtblCetakRow
                    RowCetak.Primary_PO = RowCetakDummy.Primary_PO
                    RowCetak.Primary_Product = RowCetakDummy.Primary_Product
                    RowCetak.Primary_Supplier = RowCetakDummy.Primary_Supplier
                    RowCetak.Primary_Units = RowCetakDummy.Primary_Units
                    RowCetak.PO_No = RowCetakDummy.PO_No
                    RowCetak.PO_Date = RowCetakDummy.PO_Date
                    RowCetak.Product_Code = RowCetakDummy.Product_Code
                    RowCetak.Product_Name = RowCetakDummy.Product_Name
                    RowCetak.Supplier_Code = RowCetakDummy.Supplier_Code
                    RowCetak.Supplier_Name = RowCetakDummy.Supplier_Name
                    If IsNumeric(RowCetakDummy.Payment_Term) Then
                        RowCetak.Payment_Term = RowCetakDummy.Payment_Term & " DAYS"
                    Else
                        RowCetak.Payment_Term = RowCetakDummy.Payment_Term
                    End If
                    RowCetak.Unit_Code = RowCetakDummy.Unit_Code
                    RowCetak.Ordered = RowCetakDummy.Ordered
                    RowCetak.Received = RowCetakDummy.Received
                    RowCetak.OS = RowCetakDummy.OS
                    RowCetak.Persen = RowCetakDummy.Persen
                    RowCetak.Primary_Currency = RowCetakDummy.Primary_Currency
                    RowCetak.Currency_Code = RowCetakDummy.Currency_Code
                    RowCetak.Nett_Price = RowCetakDummy.Nett_Price
                    RowCetak.Amount = RowCetakDummy.Amount
                    RowCetak.Tax = RowCetakDummy.Tax
                    RowCetak.SubTotal = 0
                    RowCetak.VAT = 0
                    RowCetak.Total = 0
                    RowCetak.Default_Precision = RowCetakDummy.Default_Precision

                    ''@rosadinnn
                    ''@MUaT
                    RowCetak.Ratio_Default = RowCetakDummy.Ratio_Default
                    RowCetak._Decimal = RowCetakDummy._Decimal
                    RowCetak.Unit_Code_Default = RowCetakDummy.Unit_Code_Default
                    RowCetak.Ratio_Reference = RowCetakDummy.Ratio_Reference
                    RowCetak.Primary_Units_Default = RowCetakDummy.Primary_Units_Default

                    RowCetak.Created_By = RowCetakDummy.Created_By
                    RowCetak.Primary_Brand = RowCetakDummy.Primary_Brand
                    RowCetak.Primary_Category = RowCetakDummy.Primary_Category
                    RowCetak.Primary_Sub_Brand = RowCetakDummy.Primary_Sub_Brand
                    RowCetak.Primary_Group = RowCetakDummy.Primary_Group
                    RowCetak.Primary_Director_Authorization = RowCetakDummy.Primary_Director_Authorization
                    RowCetak.Primary_Purchase_Manager_Authorization = RowCetakDummy.Primary_Purchase_Manager_Authorization
                    RowCetak.Primary_Purchase_Staff_Authorization = RowCetakDummy.Primary_Purchase_Staff_Authorization
                    RowCetak.ApprovedByName = RowCetakDummy.ApprovedByName
                    RowCetak.identifier = RowCetakDummy.identifier
                    RowCetak.Authorization_Level = RowCetakDummy.Authorization_Level

                    TBLDelivery.DefaultView.RowFilter = Nothing
                    TBLDelivery.DefaultView.RowFilter = "[Identifier] = '" & RowCetakDummy.identifier & "'"
                    TBLDelivery.DefaultView.Sort = "[primary]"
                    If TBLDelivery.DefaultView.Count > 0 Then
                        rowDelivery = TBLDelivery.DefaultView(0).Row
                        RowCetak.Req_Date = Format(CDate(rowDelivery.delivery_date), "dd MMM yyyy")
                    End If

                    'add remarks po & po item
                    RowCetak.Remarks = RowCetakDummy.Remarks.Replace(vbCrLf, " ")
                    RowCetak.Remarks_Item = RowCetakDummy.Remarks_Item.Replace(vbCrLf, " ")

                    TBLCetak.Rows.Add(RowCetak)
                Next
            End If

            For Each RowCetak In TBLCetak.Rows
                'Dim DecSubTotal As Decimal = 0
                TBLCetakDummy.DefaultView.RowFilter = Nothing
                If PrimaryProduct <> -1 Then
                    If Me.ChkAllProductPO.Checked = False Then
                        If Me.ChkOS.Checked = True Then
                            TBLCetakDummy.DefaultView.RowFilter = filter & " And [Primary Po] = " & RowCetak.Primary_PO & _
                                                                  " AND [Primary Product] = " & PrimaryProduct & _
                                                                  " AND [OS] > 0"
                        Else
                            TBLCetakDummy.DefaultView.RowFilter = filter & " And [Primary Po] = " & RowCetak.Primary_PO & _
                                                                  " AND [Primary Product] = " & PrimaryProduct & ""
                        End If
                    Else
                        If Me.ChkOS.Checked = True Then
                            TBLCetakDummy.DefaultView.RowFilter = filter & " And [Primary Po] = " & RowCetak.Primary_PO & " AND [OS] > 0"
                        Else
                            TBLCetakDummy.DefaultView.RowFilter = filter & " And [Primary Po] = " & RowCetak.Primary_PO
                        End If
                    End If
                Else
                    If Me.ChkOS.Checked = True Then
                        TBLCetakDummy.DefaultView.RowFilter = filter & " And [Primary Po] = " & RowCetak.Primary_PO & " AND [OS] > 0"
                    Else
                        TBLCetakDummy.DefaultView.RowFilter = filter & " And [Primary Po] = " & RowCetak.Primary_PO
                    End If
                End If


                For Each DV1 In TBLCetakDummy.DefaultView
                    RowCetakDummy = DV1.Row
                    If IsNumeric(RowCetakDummy.Amount) Then
                        RowCetak.SubTotal += CDec(RowCetakDummy.Amount)
                        RowCetak.VAT += RowCetakDummy.Amount * (CDec(RowCetakDummy.Tax) / 100)
                    End If
                Next
                'RowCetak.SubTotal = DecSubTotal
                RowCetak.Total = RowCetak.SubTotal + RowCetak.VAT
            Next

            For Each RowCetak In TBLCetak.Rows
                TBLTotalCurr.DefaultView.RowFilter = Nothing
                TBLTotalCurr.DefaultView.RowFilter = "[Primary Currency ] = " & RowCetak.Primary_Currency
                If TBLTotalCurr.DefaultView.Count > 0 Then
                    For Each DV2 In TBLTotalCurr.DefaultView
                        RowTotalCurr = DV2.Row
                        If IsNumeric(RowCetak.Amount) Then
                            RowTotalCurr.Subtotal += CDec(RowCetak.Amount)
                            RowTotalCurr.VAT += RowCetak.Amount * (CDec(RowCetak.Tax) / 100)
                            RowTotalCurr.Total = RowTotalCurr.Subtotal + RowTotalCurr.VAT
                        End If
                    Next
                Else
                    RowTotalCurr = TBLTotalCurr.NewtblTotalCurrRow
                    RowTotalCurr.Primary_Currency_ = RowCetak.Primary_Currency
                    RowTotalCurr.Currency_Code = RowCetak.Currency_Code
                    RowTotalCurr.Default_Precision = RowCetak.Default_Precision
                    RowTotalCurr.Subtotal += CDec(RowCetak.Amount)
                    RowTotalCurr.VAT = RowCetak.Amount * (CDec(RowCetak.Tax) / 100)
                    RowTotalCurr.Total = RowTotalCurr.Subtotal + RowTotalCurr.VAT
                    TBLTotalCurr.Rows.Add(RowTotalCurr)
                End If

                ''@rosadinnn
                ''@MUaT
                With RowCetak
                    If Me.ChkShowRatio.Checked = True And RowCetak.Primary_Units_Default <> RowCetak.Primary_Units Then
                        Decimal_Multi = ObjProduct.GetAngkaDecimal(.Primary_Product, .Primary_Units)
                        .Ratio_Desc = "(" & ObjCd.DigitGrouping(.Ratio_Reference, Decimal_Multi) & " " & .Unit_Code & "=" & _
                        ObjCd.DigitGrouping(.Ratio_Default, ._Decimal) & " " & .Unit_Code_Default & ")"
                    Else
                        RowCetak.Ratio_Desc = ""
                    End If
                    If RowCetak.Primary_Units_Default <> RowCetak.Primary_Units Then
                        If Me.ChkShowRatio.Checked = False Then
                            Decimal_Multi = ObjProduct.GetAngkaDecimal(.Primary_Product, .Primary_Units)
                        End If
                        RowCetak.Decimal_Multi = Decimal_Multi
                    Else
                        RowCetak.Decimal_Multi = RowCetak._Decimal
                    End If
                End With
            Next

            'add remarks po & po item
            Dim ShowRemarksItem As Boolean = False

            If Me.ChkShowRemarksPO_Item.Checked = True Then
                ShowRemarksItem = True
            End If
            If Me.rboPONumber.Checked = True Then
                'generate report
                Report.SetDataSource(DS)
                Report.SetParameterValue(Report.Parameter_StartDate.ParameterFieldName, paramStartDate)
                Report.SetParameterValue(Report.Parameter_EndDate.ParameterFieldName, paramEndDate)
                Report.SetParameterValue(Report.Parameter_ShowRemarksItem.ParameterFieldName, ShowRemarksItem)
                crv.ReportSource = Report
                crv.Show()
            Else
                'generate report
                Report2.SetDataSource(DS)
                Report2.SetParameterValue(Report2.Parameter_StartDate.ParameterFieldName, paramStartDate)
                Report2.SetParameterValue(Report2.Parameter_EndDate.ParameterFieldName, paramEndDate)
                Report2.SetParameterValue(Report2.Parameter_ShowRemarksItem.ParameterFieldName, ShowRemarksItem)
                crv.ReportSource = Report2
                crv.Show()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTACetakDummy.Dispose()
            TBLCetak.Dispose()
            TBLTotalCurr.Dispose()
            TBLCetakDummy.Dispose()
            DS.Dispose()
        End Try
    End Sub

    Private Sub txtProductCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtProductCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.txtProductCode.Text <> "" Then
                    If Me.ObjProduct.ApakahSudahAda(Me.txtProductCode.Text) = True Then
                        Me.txtPrimaryProduct.Text = CStr(Me.ObjProduct.GetPrimary(Me.txtProductCode.Text))
                        Me.txtProductCode.Text = Me.ObjProduct.GetProductCode(CDbl(Me.txtPrimaryProduct.Text))
                        Me.txtProductName.Text = Me.ObjProduct.GetDescription(CDbl(Me.txtPrimaryProduct.Text))
                        Me.txtSupplierCode.Focus()
                    Else
                        Me.btnFindProduct.PerformClick()
                    End If
                Else
                    Me.btnFindProduct.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtProductCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductCode.Leave
        Dim hj As New Core_Module.frmCariMasterProduct
        'hj.BackColorAll = Me.BAckColorAll
        Try
            If Me.txtProductCode.Text <> "" Then
                If Me.ObjProduct.ApakahSudahAda(Me.txtProductCode.Text) = True Then
                    Me.txtPrimaryProduct.Text = CStr(Me.ObjProduct.GetPrimary(Me.txtProductCode.Text))
                    Me.txtProductCode.Text = Me.ObjProduct.GetProductCode(CDbl(Me.txtPrimaryProduct.Text))
                    Me.txtProductName.Text = Me.ObjProduct.GetDescription(CDbl(Me.txtPrimaryProduct.Text))
                Else
                    If Me.txtProductCode.Text <> "" Then
                        hj.txtCari.Text = Me.txtProductCode.Text
                    End If
                    hj.ShowDialog()

                    If hj.PrimaryProduct > 0 Then
                        Me.txtPrimaryProduct.Text = CStr(hj.PrimaryProduct)
                        Me.txtProductCode.Text = Me.ObjProduct.GetProductCode(CDbl(Me.txtPrimaryProduct.Text))
                        Me.txtProductName.Text = Me.ObjProduct.GetDescription(CDbl(Me.txtPrimaryProduct.Text))
                        Me.txtSupplierCode.Focus()
                    Else
                        Me.txtPrimaryProduct.Text = "-1"
                        Me.txtProductName.Clear()
                        Me.txtProductCode.Clear()
                        Me.txtProductCode.Focus()
                    End If
                End If
            Else
                Me.txtPrimaryProduct.Text = "-1"
                Me.txtProductName.Clear()
                Me.txtProductCode.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnFindProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindProduct.Click
        Dim hj As New Core_Module.frmCariMasterProduct
        Try
            'hj.BackColorAll = Me.BAckColorAll
            If Me.txtProductCode.Text <> "" Then
                hj.txtCari.Text = Me.txtProductCode.Text
            End If

            hj.ShowDialog()
            If hj.PrimaryProduct > 0 Then
                Me.txtPrimaryProduct.Text = CStr(hj.PrimaryProduct)
                Me.txtProductCode.Text = Me.ObjProduct.GetProductCode(CDbl(Me.txtPrimaryProduct.Text))
                Me.txtProductName.Text = Me.ObjProduct.GetDescription(CDbl(Me.txtPrimaryProduct.Text))
                Me.txtSupplierCode.Focus()
            Else
                Me.txtPrimaryProduct.Text = "-1"
                Me.txtProductName.Clear()
                Me.txtProductCode.Clear()
                Me.txtProductCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub txtSupplierCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSupplierCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.txtSupplierCode.Text <> "" Then
                    If Me.ObjSupplier.ApakahSudahAda(Me.txtSupplierCode.Text) = True Then
                        Me.txtPrimarySupplier.Text = CStr(Me.ObjSupplier.GetPrimary(Me.txtSupplierCode.Text))
                        Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(CDbl(Me.txtPrimarySupplier.Text))
                        Me.txtSupplierName.Text = Me.ObjSupplier.GetSupplierName(CDbl(Me.txtPrimarySupplier.Text))
                        Me.txtCurrencyCode.Focus()
                    Else
                        Me.btnFindSupplier.PerformClick()
                    End If
                Else
                    Me.btnFindSupplier.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSupplierCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierCode.Leave
        Dim hj As New Core_Module.FrmCariMasterSupplier
        'hj.BackColorAll = Me.BAckColorAll
        Try
            If Me.txtSupplierCode.Text <> "" Then
                If Me.ObjSupplier.ApakahSudahAda(Me.txtSupplierCode.Text) = True Then
                    Me.txtPrimarySupplier.Text = CStr(Me.ObjSupplier.GetPrimary(Me.txtSupplierCode.Text))
                    Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(CDbl(Me.txtPrimarySupplier.Text))
                    Me.txtSupplierName.Text = Me.ObjSupplier.GetSupplierName(CDbl(Me.txtPrimarySupplier.Text))
                Else
                    If Me.txtSupplierCode.Text <> "" Then
                        hj.txtCari.Text = Me.txtSupplierCode.Text
                    End If
                    hj.ShowDialog()

                    If hj.PrimarySupplier > 0 Then
                        Me.txtPrimarySupplier.Text = CStr(hj.PrimarySupplier)
                        Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(CDbl(Me.txtPrimarySupplier.Text))
                        Me.txtSupplierName.Text = Me.ObjSupplier.GetSupplierName(CDbl(Me.txtPrimarySupplier.Text))
                        Me.txtCurrencyCode.Focus()
                    Else
                        Me.txtPrimarySupplier.Text = "-1"
                        Me.txtSupplierName.Clear()
                        Me.txtSupplierCode.Clear()
                        Me.txtSupplierCode.Focus()
                    End If
                End If
            Else
                Me.txtPrimarySupplier.Text = "-1"
                Me.txtSupplierName.Clear()
                Me.txtSupplierCode.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnFindSupplier_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindSupplier.Click
        Dim hj As New Core_Module.FrmCariMasterSupplier
        Try
            'hj.BackColorAll = Me.BAckColorAll
            If Me.txtSupplierCode.Text <> "" Then
                hj.txtCari.Text = Me.txtSupplierCode.Text
            End If

            hj.ShowDialog()
            If hj.PrimarySupplier > 0 Then
                Me.txtPrimarySupplier.Text = CStr(hj.PrimarySupplier)
                Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(CDbl(Me.txtPrimarySupplier.Text))
                Me.txtSupplierName.Text = Me.ObjSupplier.GetSupplierName(CDbl(Me.txtPrimarySupplier.Text))
                Me.txtCurrencyCode.Focus()
            Else
                Me.txtPrimarySupplier.Text = "-1"
                Me.txtSupplierName.Clear()
                Me.txtSupplierCode.Clear()
                Me.txtSupplierCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub txtCurrencyCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCurrencyCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.txtCurrencyCode.Text <> "" Then
                    If Me.ObjCurr.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                        Me.txtPrimaryCurrency.Text = CStr(Me.ObjCurr.GetPrimary(Me.txtCurrencyCode.Text))
                        Me.txtCurrencyCode.Text = Me.ObjCurr.GetCurrencyCode(CDbl(Me.txtPrimaryCurrency.Text))
                        Me.btnFilter.Focus()
                    Else
                        Me.btnFindCurrency.PerformClick()
                    End If
                Else
                    Me.btnFindCurrency.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCurrencyCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.Leave
        Dim hj As New Core_Module.FrmCariMasterCurrency
        'hj.BackColorAll = Me.BAckColorAll
        Try
            If Me.txtCurrencyCode.Text <> "" Then
                If Me.ObjCurr.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                    Me.txtPrimaryCurrency.Text = CStr(Me.ObjCurr.GetPrimary(Me.txtCurrencyCode.Text))
                    Me.txtCurrencyCode.Text = Me.ObjCurr.GetCurrencyCode(CDbl(Me.txtPrimaryCurrency.Text))
                Else
                    If Me.txtCurrencyCode.Text <> "" Then
                        hj.txtCari.Text = Me.txtCurrencyCode.Text
                    End If
                    hj.ShowDialog()

                    If hj.PrimaryCurrency > 0 Then
                        Me.txtPrimaryCurrency.Text = CStr(hj.PrimaryCurrency)
                        Me.txtCurrencyCode.Text = Me.ObjCurr.GetCurrencyCode(CDbl(Me.txtPrimaryCurrency.Text))
                        Me.btnFilter.Focus()
                    Else
                        Me.txtPrimaryCurrency.Text = "-1"
                        Me.txtCurrencyCode.Clear()
                        Me.txtCurrencyCode.Focus()
                    End If
                End If
            Else
                Me.txtPrimaryCurrency.Text = "-1"
                Me.txtCurrencyCode.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnFindCurrency_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCurrency.Click
        Dim hj As New Core_Module.FrmCariMasterCurrency
        Try
            'hj.BackColorAll = Me.BAckColorAll
            If Me.txtCurrencyCode.Text <> "" Then
                hj.txtCari.Text = Me.txtCurrencyCode.Text
            End If

            hj.ShowDialog()
            If hj.PrimaryCurrency > 0 Then
                Me.txtPrimaryCurrency.Text = CStr(hj.PrimaryCurrency)
                Me.txtCurrencyCode.Text = Me.ObjCurr.GetCurrencyCode(CDbl(Me.txtPrimaryCurrency.Text))
                Me.btnFilter.Focus()
            Else
                Me.txtPrimaryCurrency.Text = "-1"
                Me.txtCurrencyCode.Clear()
                Me.txtCurrencyCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub


    Private Sub txtProductCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductCode.GotFocus, txtCurrencyCode.GotFocus, txtSupplierCode.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnFilter.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.txtProductName.BackColor
                        Else
                            Txt.BackColor = Me.BAckColorAll
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtProductCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductCode.LostFocus, txtSupplierCode.LostFocus, txtCurrencyCode.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnFilter.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.txtProductName.BackColor
                        Else
                            Txt.BackColor = Me.txtPrimaryProduct.BackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPrimaryProduct_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPrimaryProduct.TextChanged
        Try
            If Me.txtPrimaryProduct.Text = -1 Then
                Me.ChkAllProductPO.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkAllProductPO_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkAllProductPO.Click
        Try
            If Me.txtPrimaryProduct.Text = -1 AndAlso (Me.dgvItemProductGroup.RowCount > 0 AndAlso _
                                                      CDec(Me.dgvItemProductGroup.Rows(0).Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value) <= 0) Then
                Me.ChkAllProductPO.Checked = False
                MsgBox("Product Or " & Me.dgvItemProductGroup.Rows(0).Cells(Me.ColProductGroup_Level.Index).Value & " cannot be an empty string ", MsgBoxStyle.Information, "Information")
                Me.txtProductCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanPurchaseOrderReportByDateDetail_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.ChkOS.Text = "Only show PO with balance"
            Me.dgvItemProductGroup.Rows.Clear()
            Me.DgvUser.Rows.Clear()
            Me.CboStatusPO.SelectedIndex = 0
            Me.txtItemRemarks.Enabled = False
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

    Private Sub rboProduct_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rboProduct.CheckedChanged
        Try
            If Me.rboProduct.Checked = True Then
                Me.ChkAllProductPO.Enabled = False
                Me.txtRemarks.Enabled = False
            ElseIf Me.rboPONumber.Checked = True Then
                Me.ChkAllProductPO.Enabled = True
                Me.txtRemarks.Enabled = True
            End If
            Me.txtRemarks.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtRemarks_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemarks.GotFocus
        Me.txtRemarks.BackColor = Me.BAckColorAll
    End Sub

    Private Sub txtRemarks_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemarks.Leave
        Me.txtRemarks.Text = Me.txtRemarks.Text.Trim
    End Sub

    Private Sub txtRemarks_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemarks.LostFocus
        Me.txtRemarks.BackColor = Me.txtItemRemarks.BackColor
    End Sub

    Private Sub txtItemRemarks_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtItemRemarks.GotFocus
        Me.txtItemRemarks.BackColor = Me.BAckColorAll
    End Sub

    Private Sub txtItemRemarks_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtItemRemarks.Leave
        Me.txtItemRemarks.Text = Me.txtItemRemarks.Text.Trim
    End Sub

    Private Sub txtItemRemarks_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtItemRemarks.LostFocus
        Me.txtItemRemarks.BackColor = Me.txtRemarks.BackColor
    End Sub

    Private Sub ChkShowRemarksPO_Item_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkShowRemarksPO_Item.CheckedChanged
        Try
            If Me.ChkShowRemarksPO_Item.Checked = True Then
                Me.txtItemRemarks.Enabled = True
            Else
                Me.txtItemRemarks.Enabled = False
            End If
            Me.txtItemRemarks.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class