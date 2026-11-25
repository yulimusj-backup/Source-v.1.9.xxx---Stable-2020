Option Explicit On
Option Strict Off
Imports Core_Module

Public Class FrmTransaksiSalesInvoice_AddData
    Public PrimaryInvoice As Double = -1
    Public PrimaryCustomer As Double = -1
    Public PrimaryCurrency As Double = -1
    Public OtherCharge As Decimal = 0
    Public OtherChargeDummy As Decimal = 0
    Public TBLHeaderCost As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_add_dataDataTable
    Public rowHeaderCost As DseTransaksiSalesInvoice.tbltrans_salesinvoice_add_dataRow
    Public TBLitemCost As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_item_costDataTable
    Public rowitemCost As DseTransaksiSalesInvoice.tbltrans_salesinvoice_item_costRow
    Public Aktif As Boolean = True
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Public PrimaryCustomerShipTo As Double = -1
    Public IdentifierAddressCustomerShipTo As String = ""

    Dim ObjInv As New ClsTransaksiSalesInvoice
    Dim ObjUnit As New Core_Module.ClsMasterUnit
    Dim ObjCustomer As New Core_Module.clsMasterCustomer
    Dim Objakun As New Core_Module.ClsMasterAkun1
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCurrency As New Core_Module.clsMasterCurrency
    Dim AdaData As Boolean = False
    Dim Buffer As String = ""
    Dim TipeAkun As New ClsMasterAkun1.EnumTipeAkun2

    Private Sub FrmTransaksiSalesInvoice_AddData_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.dgvCost.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvCost.AllowUserToDeleteRows = False
            Me.dgvCost.AllowUserToAddRows = False
            If Me.TBLHeaderCost.Rows.Count > 0 Then
                SalinDariDataTable()
            Else
                SalinHistoricalData()
            End If
            Me.dgvCost.AllowUserToDeleteRows = True
            Me.dgvCost.AllowUserToAddRows = True

            If Me.Aktif = False Then
                NonAktifkanDetail()
                Me.btnOK.Visible = False
                Me.btnOK.TabStop = False
            Else
                Me.btnOK.Visible = True
                Me.btnOK.TabStop = True
            End If
            Me.txtShippingMarks.Width = Me.SplitContainer1.Panel2.Width - 5 - txtShippingMarks.Left
            Me.txtInvoiceType.Width = Me.txtShippingMarks.Width
            Me.txtVar5_Value.Width = Me.SplitContainer1.Panel1.Width - 5 - Me.txtVar5_Value.Left
            Me.txtVar4_Value.Width = Me.txtVar5_Value.Width
            Me.txtVar3_Value.Width = Me.txtVar5_Value.Width
            Me.txtVar2_Value.Width = Me.txtVar5_Value.Width
            Me.txtVar1_Value.Width = Me.txtVar5_Value.Width
            Me.txtTransportationMode.Width = Me.txtVar5_Value.Width
            Me.CancelButton = Me.btnCancel
            Me.dtpDestinationDate.Left = Me.txtVar5_Value.Width + Me.txtVar5_Value.Left - Me.dtpDestinationDate.Width
            Me.dtpLoadingDate.Left = Me.dtpDestinationDate.Left
            Me.txtDestinationPort.Width = Me.dtpDestinationDate.Left - 5 - Me.txtDestinationPort.Left
            Me.txtLoadingPort.Width = Me.txtDestinationPort.Width
            Me.btnFindUnitVolume.Left = Me.txtVar5_Value.Width + Me.txtVar5_Value.Left - Me.btnFindUnitVolume.Width
            Me.btnFindUnitWeight.Left = Me.btnFindUnitVolume.Left
            txtUnitWeight.Width = Me.btnFindUnitWeight.Left - 5 - Me.txtUnitWeight.Left
            Me.txtUnitVolume.Width = txtUnitWeight.Width

            Me.TxtPrimaryCustomerShipTo.Text = Me.PrimaryCustomerShipTo
            Me.TxtCustomerCodeShipTo.Text = Me.ObjCustomer.GetCustomerCode(Me.PrimaryCustomerShipTo)
            Me.TxtCustomerDescShipTo.Text = Me.ObjCustomer.GetCustName(Me.PrimaryCustomerShipTo)
            Me.ObjCustomer.IsiCboAddress(Me.CboAddressCustomerShipTo, Me.PrimaryCustomerShipTo)
            If Me.CboAddressCustomerShipTo.Items.Count > 0 Then
                Me.CboAddressCustomerShipTo.Text = Me.ObjCustomer.GetKodeAddress(Me.IdentifierAddressCustomerShipTo)
            End If
            Me.TxtIdentifierAddressCustomerShipTo.Text = Me.IdentifierAddressCustomerShipTo
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False, False, False)
            Me.ObjCD.KendalikanControl(Me.gboAll.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.gboCost.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.SetBackColor(Me.Controls, Me.TxtCustomerDescShipTo.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboCost.Controls, Me.TxtCustomerDescShipTo.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboAll.Controls, Me.TxtCustomerDescShipTo.BackColor, False)
            Me.dgvCost.AllowUserToDeleteRows = False
            Me.dgvCost.AllowUserToAddRows = False
            Me.btnOK.Enabled = True
            Me.btnCancel.Enabled = True
            Me.dgvCost.ReadOnly = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeObject()
        Dim rowDGV As New DataGridViewRow
        Try
            Me.TBLHeaderCost.Rows.Clear()
            Me.TBLitemCost.Rows.Clear()
            Me.OtherCharge = 0
            Me.rowHeaderCost = Me.TBLHeaderCost.Newtbltrans_salesinvoice_add_dataRow
            Me.rowHeaderCost.primary = -1
            Me.rowHeaderCost.primary_inv = Me.PrimaryInvoice
            Me.rowHeaderCost.primary_weight_unit = Me.txtPrimaryUnitWeight.Text
            Me.rowHeaderCost.primary_volume_unit = Me.txtPrimaryUnitVolume.Text
            Me.rowHeaderCost.Total_Gross_Volume = Me.txtGrossVolume.Text
            Me.rowHeaderCost.Total_Gross_Weight = Me.txtGrossWeight.Text
            Me.rowHeaderCost.Total_Nett_Volume = Me.txtNettVolume.Text
            Me.rowHeaderCost.Total_Nett_Weight = Me.txtNettWeight.Text
            Me.rowHeaderCost.ETA_Port = Me.txtLoadingPort.Text
            Me.rowHeaderCost.ETA_Date = Me.dtpLoadingDate.Value
            Me.rowHeaderCost.ETD_Port = Me.txtDestinationPort.Text
            Me.rowHeaderCost.ETD_Date = Me.dtpDestinationDate.Value
            Me.rowHeaderCost.transportation_mode = Me.txtTransportationMode.Text
            Me.rowHeaderCost.var1_name = Me.txtVar1_Name.Text
            Me.rowHeaderCost.var2_name = Me.txtVar2_Name.Text
            Me.rowHeaderCost.var3_name = Me.txtVar3_Name.Text
            Me.rowHeaderCost.var4_name = Me.txtVar4_Name.Text
            Me.rowHeaderCost.var5_name = Me.txtVar5_Name.Text
            Me.rowHeaderCost.var1_value = Me.txtVar1_Value.Text
            Me.rowHeaderCost.var2_value = Me.txtVar2_Value.Text
            Me.rowHeaderCost.var3_value = Me.txtVar3_Value.Text
            Me.rowHeaderCost.var4_value = Me.txtVar4_Value.Text
            Me.rowHeaderCost.var5_value = Me.txtVar5_Value.Text
            Me.rowHeaderCost.invoice_type = Me.txtInvoiceType.Text
            Me.rowHeaderCost.shipping_mark = Me.txtShippingMarks.Text
            Me.rowHeaderCost.primary_cust_ship_to = Me.PrimaryCustomerShipTo
            Me.rowHeaderCost.identifier_cust_ship_to = Me.IdentifierAddressCustomerShipTo
            Me.TBLHeaderCost.Rows.Add(Me.rowHeaderCost)
            If Me.dgvCost.RowCount > 0 Then
                For Each rowDGV In Me.dgvCost.Rows
                    With rowDGV
                        If rowDGV.Index <> Me.dgvCost.RowCount - 1 Then
                            Me.rowitemCost = Me.TBLitemCost.Newtbltrans_salesinvoice_item_costRow
                            Me.rowitemCost.Primary = -1
                            Me.rowitemCost.Primary_Inv = Me.PrimaryInvoice
                            Me.rowitemCost.Primary_Akun = .Cells(Me.ColPrimaryAccount.Index).Value
                            Me.rowitemCost.Cost_Description = .Cells(Me.ColDescCost.Index).Value
                            If IsNumeric(.Cells(Me.ColCostAmount.Index).Value) Then
                                Me.rowitemCost.Cost = .Cells(Me.ColCostAmount.Index).Value
                                Me.OtherCharge += CDec(.Cells(Me.ColCostAmount.Index).Value)
                            End If
                            If String.IsNullOrEmpty(.Cells(Me.ColRemarksCost.Index).Value) Then
                                Me.rowitemCost.Remarks = ""
                            Else
                                Me.rowitemCost.Remarks = .Cells(Me.ColRemarksCost.Index).Value
                            End If
                            Me.TBLitemCost.Rows.Add(Me.rowitemCost)
                        End If
                    End With
                Next
            End If
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungOtherCharge()
        Dim rowDgv As New DataGridViewRow
        Try
            If Me.dgvCost.RowCount > 0 Then
                Me.OtherCharge = 0
                For Each rowDgv In Me.dgvCost.Rows
                    With rowDgv
                        If IsNumeric(.Cells(Me.ColCostAmount.Index).Value) Then
                            Me.OtherCharge += CDec(.Cells(Me.ColCostAmount.Index).Value)
                        End If
                    End With
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinHistoricalData()
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_add_dataTableAdapter
        Dim tbl As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_add_dataDataTable
        Dim row As DseTransaksiSalesInvoice.tbltrans_salesinvoice_add_dataRow

        Dim DTACost As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_item_costTableAdapter
        Dim tblCost As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_item_costDataTable
        Dim rowCost As DseTransaksiSalesInvoice.tbltrans_salesinvoice_item_costRow

        Dim PrimaryInvoiceMax As Double = -1
        Dim rowDGV As New DataGridViewRow
        Dim PrimaryCurrencyInvB4 As Double = -1
        Dim roundB4 As Decimal = 2
        Try
            PrimaryInvoiceMax = Me.ObjInv.GetPrimaryMax(Me.PrimaryCustomer)
            DTA.FillBy(tbl, PrimaryInvoiceMax)
            DTACost.Fill(tblCost, PrimaryInvoiceMax)

            Me.OtherCharge = 0
            If tbl.Rows.Count > 0 Then
                For Each row In tbl.Rows
                    With row

                        roundB4 = Me.ObjCD.GetJmlPrecision(.Total_Nett_Weight)
                        If roundB4 >= 2 Then
                            Me.txtNettWeight.Text = ObjCD.DigitGrouping(.Total_Nett_Weight, roundB4)
                        Else
                            Me.txtNettWeight.Text = ObjCD.DigitGrouping(.Total_Nett_Weight, 2)
                        End If

                        roundB4 = Me.ObjCD.GetJmlPrecision(.Total_Gross_Weight)
                        If roundB4 >= 2 Then
                            Me.txtGrossWeight.Text = ObjCD.DigitGrouping(.Total_Gross_Weight, roundB4)
                        Else
                            Me.txtGrossWeight.Text = ObjCD.DigitGrouping(.Total_Gross_Weight, 2)
                        End If

                        Me.txtPrimaryUnitWeight.Text = .primary_weight_unit
                        Me.txtUnitWeight.Text = ObjUnit.GetUnitCode(.primary_weight_unit)

                        roundB4 = Me.ObjCD.GetJmlPrecision(.Total_Nett_Volume)
                        If roundB4 >= 2 Then
                            Me.txtNettVolume.Text = ObjCD.DigitGrouping(.Total_Nett_Volume, roundB4)
                        Else
                            Me.txtNettVolume.Text = ObjCD.DigitGrouping(.Total_Nett_Volume, 2)
                        End If

                        roundB4 = Me.ObjCD.GetJmlPrecision(.Total_Gross_Volume)
                        If roundB4 >= 2 Then
                            Me.txtGrossVolume.Text = ObjCD.DigitGrouping(.Total_Gross_Volume, roundB4)
                        Else
                            Me.txtGrossVolume.Text = ObjCD.DigitGrouping(.Total_Gross_Volume, 2)
                        End If

                        Me.txtPrimaryUnitVolume.Text = .primary_volume_unit
                        Me.txtUnitVolume.Text = ObjUnit.GetUnitCode(.primary_volume_unit)
                        Me.txtLoadingPort.Text = .ETA_Port
                        Me.dtpLoadingDate.Value = .ETA_Date
                        Me.txtDestinationPort.Text = .ETD_Port
                        Me.dtpDestinationDate.Value = .ETD_Date
                        Me.txtTransportationMode.Text = .transportation_mode
                        Me.txtVar1_Value.Text = .var1_value
                        Me.txtVar1_Name.Text = .var1_name
                        Me.txtVar2_Value.Text = .var2_value
                        Me.txtVar2_Name.Text = .var2_name
                        Me.txtVar3_Value.Text = .var3_value
                        Me.txtVar3_Name.Text = .var3_name
                        Me.txtVar4_Value.Text = .var4_value
                        Me.txtVar4_Name.Text = .var4_name
                        Me.txtVar5_Value.Text = .var5_value
                        Me.txtVar5_Name.Text = .var5_name
                        Me.txtInvoiceType.Text = .invoice_type
                        Me.txtShippingMarks.Text = .shipping_mark
                        PrimaryCurrencyInvB4 = Me.ObjInv.Get_Primary_Currency(.primary_inv)
                    End With
                Next

                If tblCost.Rows.Count > 0 Then
                    For Each rowCost In tblCost.Rows
                        With rowCost
                            Me.dgvCost.Rows.Add()
                            rowDGV = Me.dgvCost.Rows(Me.dgvCost.Rows.Count - 1)
                            rowDGV.Cells(Me.ColPrimaryCurrencyCost.Index).Value = 1
                            rowDGV.Cells(Me.ColExcRateCost.Index).Value = 1
                            rowDGV.Cells(Me.ColCurrencyCost.Index).Value = Me.ObjCurrency.GetCurrencyCode(Me.PrimaryCurrency)

                            If PrimaryCurrencyInvB4 = Me.PrimaryCurrency Then
                                If IsNumeric(.Cost) Then
                                    rowDGV.Cells(Me.ColCostNominal.Index).Value = ObjCD.DigitGrouping(.Cost, Me.ObjCurrency.GetDefaultPrecission(Me.PrimaryCurrency))
                                End If
                            Else
                                rowDGV.Cells(Me.ColCostNominal.Index).Value = ObjCD.DigitGrouping(0, Me.ObjCurrency.GetDefaultPrecission(Me.PrimaryCurrency))
                            End If

                            rowDGV = Me.dgvCost.Rows(Me.dgvCost.Rows.Count - 1)
                            rowDGV.Cells(Me.ColDescCost.Index).Value = .Cost_Description

                            If PrimaryCurrencyInvB4 = Me.PrimaryCurrency Then
                                rowDGV.Cells(Me.ColCostAmount.Index).Value = ObjCD.DigitGrouping(.Cost, Me.ObjCurrency.GetDefaultPrecission(Me.PrimaryCurrency))
                            Else
                                rowDGV.Cells(Me.ColCostAmount.Index).Value = ObjCD.DigitGrouping(0, Me.ObjCurrency.GetDefaultPrecission(Me.PrimaryCurrency))
                            End If

                            If String.IsNullOrEmpty(.Remarks) Then
                                rowDGV.Cells(Me.ColRemarksCost.Index).Value = ""
                            Else
                                rowDGV.Cells(Me.ColRemarksCost.Index).Value = .Remarks
                            End If

                            rowDGV.Cells(Me.ColPrimaryAccount.Index).Value = .Primary_Akun
                            rowDGV.Cells(Me.ColAccounrCode.Index).Value = Objakun.GetAccountCode(.Primary_Akun)
                            rowDGV.Cells(Me.ColAccountDesc.Index).Value = Objakun.GetDescription_ByPrimaryAcc(.Primary_Akun)
                            rowDGV.Cells(Me.ColDescCost.Index).Value = .Cost_Description
                        End With
                    Next
                    Me.ObjCD.SetNumber(Me.dgvCost)
                End If
                AdaData = True
            Else
                AdaData = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            tbl.Dispose()
            DTACost.Dispose()
            tblCost.Dispose()
        End Try
    End Sub

    Private Sub SalinDariDataTable()
        Dim rowDGV As New DataGridViewRow
        Dim RoundB4 As Decimal = 2
        Try
            Me.OtherCharge = 0
            If Me.TBLHeaderCost.Rows.Count > 0 Then
                For Each Me.rowHeaderCost In Me.TBLHeaderCost.Rows
                    With Me.rowHeaderCost
                        roundB4 = Me.ObjCD.GetJmlPrecision(.Total_Nett_Weight)
                        If roundB4 >= 2 Then
                            Me.txtNettWeight.Text = ObjCD.DigitGrouping(.Total_Nett_Weight, roundB4)
                        Else
                            Me.txtNettWeight.Text = ObjCD.DigitGrouping(.Total_Nett_Weight, 2)
                        End If

                        roundB4 = Me.ObjCD.GetJmlPrecision(.Total_Gross_Weight)
                        If roundB4 >= 2 Then
                            Me.txtGrossWeight.Text = ObjCD.DigitGrouping(.Total_Gross_Weight, roundB4)
                        Else
                            Me.txtGrossWeight.Text = ObjCD.DigitGrouping(.Total_Gross_Weight, 2)
                        End If

                        Me.txtPrimaryUnitWeight.Text = .primary_weight_unit
                        Me.txtUnitWeight.Text = ObjUnit.GetUnitCode(.primary_weight_unit)

                        roundB4 = Me.ObjCD.GetJmlPrecision(.Total_Nett_Volume)
                        If roundB4 >= 2 Then
                            Me.txtNettVolume.Text = ObjCD.DigitGrouping(.Total_Nett_Volume, roundB4)
                        Else
                            Me.txtNettVolume.Text = ObjCD.DigitGrouping(.Total_Nett_Volume, 2)
                        End If

                        roundB4 = Me.ObjCD.GetJmlPrecision(.Total_Gross_Volume)
                        If roundB4 >= 2 Then
                            Me.txtGrossVolume.Text = ObjCD.DigitGrouping(.Total_Gross_Volume, roundB4)
                        Else
                            Me.txtGrossVolume.Text = ObjCD.DigitGrouping(.Total_Gross_Volume, 2)
                        End If

                        Me.txtPrimaryUnitVolume.Text = .primary_volume_unit
                        Me.txtUnitVolume.Text = ObjUnit.GetUnitCode(.primary_volume_unit)

                        Me.txtPrimaryUnitVolume.Text = .primary_volume_unit
                        Me.txtLoadingPort.Text = .ETA_Port
                        Me.dtpLoadingDate.Value = .ETA_Date
                        Me.txtDestinationPort.Text = .ETD_Port
                        Me.dtpDestinationDate.Value = .ETD_Date
                        Me.txtTransportationMode.Text = .transportation_mode
                        Me.txtVar1_Value.Text = .var1_value
                        Me.txtVar1_Name.Text = .var1_name
                        Me.txtVar2_Value.Text = .var2_value
                        Me.txtVar2_Name.Text = .var2_name
                        Me.txtVar3_Value.Text = .var3_value
                        Me.txtVar3_Name.Text = .var3_name
                        Me.txtVar4_Value.Text = .var4_value
                        Me.txtVar4_Name.Text = .var4_name
                        Me.txtVar5_Value.Text = .var5_value
                        Me.txtVar5_Name.Text = .var5_name
                        Me.txtInvoiceType.Text = .invoice_type
                        Me.txtShippingMarks.Text = .shipping_mark
                    End With
                Next

                If Me.TBLitemCost.Rows.Count > 0 Then
                    Me.dgvCost.Rows.Clear()
                    For Each Me.rowitemCost In Me.TBLitemCost.Rows
                        With Me.rowitemCost
                            Me.dgvCost.Rows.Add()
                            rowDGV = Me.dgvCost.Rows(Me.dgvCost.Rows.Count - 1)
                            rowDGV.Cells(Me.ColPrimaryCurrencyCost.Index).Value = 1
                            rowDGV.Cells(Me.ColCurrencyCost.Index).Value = Me.ObjCurrency.GetCurrencyCode(Me.PrimaryCurrency)
                            rowDGV.Cells(Me.ColExcRateCost.Index).Value = 1
                            If IsNumeric(.Cost) Then
                                rowDGV.Cells(Me.ColCostNominal.Index).Value = ObjCD.DigitGrouping(.Cost, Me.ObjCurrency.GetDefaultPrecission(Me.PrimaryCurrency))
                            End If
                            rowDGV = Me.dgvCost.Rows(Me.dgvCost.Rows.Count - 1)
                            rowDGV.Cells(Me.ColDescCost.Index).Value = .Cost_Description
                            rowDGV.Cells(Me.ColCostAmount.Index).Value = ObjCD.DigitGrouping(.Cost, Me.ObjCurrency.GetDefaultPrecission(Me.PrimaryCurrency))
                            If String.IsNullOrEmpty(.Remarks) Then
                                rowDGV.Cells(Me.ColRemarksCost.Index).Value = ""
                            Else
                                rowDGV.Cells(Me.ColRemarksCost.Index).Value = .Remarks
                            End If
                            rowDGV.Cells(Me.ColPrimaryAccount.Index).Value = .Primary_Akun
                            rowDGV.Cells(Me.ColAccounrCode.Index).Value = Objakun.GetAccountCode(.Primary_Akun)
                            rowDGV.Cells(Me.ColAccountDesc.Index).Value = Objakun.GetDescription_ByPrimaryAcc(.Primary_Akun)
                            rowDGV.Cells(Me.ColDescCost.Index).Value = .Cost_Description
                        End With
                    Next
                    Me.ObjCD.SetNumber(Me.dgvCost)
                End If
                AdaData = True
            Else
                AdaData = False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If Me.Aktif = True Then
                If Me.ApakahDataBenar = True Then
                    Me.SalinKeObject()
                End If
            Else
                HitungOtherCharge()
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Try
            'cek data yang harus diisi
            If CDec(Me.txtGrossWeight.Text) <> 0 Or CDec(Me.txtNettWeight.Text) <> 0 Then
                KoleksiNamaData.Add("Unit Weight")
                KoleksiTextBox.Add(Me.txtUnitWeight)
            End If
            If CDec(Me.txtGrossVolume.Text) <> 0 Or CDec(Me.txtNettVolume.Text) <> 0 Then
                KoleksiNamaData.Add("Unit Volume")
                KoleksiTextBox.Add(Me.txtUnitVolume)
            End If

            Me.txtLoadingPort.Text = Me.txtLoadingPort.Text.Trim
            Me.txtDestinationPort.Text = Me.txtDestinationPort.Text.Trim
            Me.txtTransportationMode.Text = Me.txtTransportationMode.Text.Trim
            Me.txtInvoiceType.Text = Me.txtInvoiceType.Text.Trim

            'KoleksiNamaData.Add("Loading port")
            'KoleksiTextBox.Add(Me.txtLoadingPort)
            'KoleksiNamaData.Add("Destination port")
            'KoleksiTextBox.Add(Me.txtDestinationPort)
            'KoleksiNamaData.Add("Transportation Mode")
            'KoleksiTextBox.Add(Me.txtTransportationMode)
            'KoleksiNamaData.Add("Invoice Type")
            'KoleksiTextBox.Add(Me.txtInvoiceType)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'minimal harus punya 1 data item
            If Me.dgvCost.RowCount < 1 OrElse IsDBNull(Me.dgvCost.Rows(0).Cells(Me.ColAccounrCode.Index).Value) Then
                MsgBox("Please enter at least 1 item.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Dim rowDGV As New DataGridViewRow
            'untuk setiap item cost, account code harus diisikan dan tipenya harusnya adalah asset atau payable
            For Each rowDGV In Me.dgvCost.Rows
                If rowDGV.Index <> Me.dgvCost.Rows.Count - 1 Then
                    If String.IsNullOrEmpty(rowDGV.Cells(Me.ColDescCost.Index).Value) AndAlso Not String.IsNullOrEmpty(rowDGV.Cells(Me.ColAccounrCode.Index).Value) Then
                        MsgBox("Cost description cannot be an empty string.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If

                    'If ApakahDescriptionBerulang(rowDGV.Cells(Me.ColDescCost.Index).Value, rowDGV.Index) = True Then
                    '    MsgBox("Repetition of cost description.", MsgBoxStyle.Information, "Information")
                    '    Return False
                    'End If

                    'pastikan akunnya diisikan
                    If IsNumeric(rowDGV.Cells(Me.ColPrimaryAccount.Index).Value) = False Then
                        MsgBox("Please specify account code for cost data on row " + (rowDGV.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                        Me.dgvCost.ClearSelection()
                        Me.dgvCost.Rows(rowDGV.Index).Selected = True
                        Return False
                    End If

                    'pastikan akunnya ada dan valid
                    If Me.Objakun.ApakahSudahAda_Aktif_NonHeader(rowDGV.Cells(Me.ColAccounrCode.Index).Value) = False Then
                        MsgBox("Please specify valid account code for cost data on row " + (rowDGV.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                        Me.dgvCost.ClearSelection()
                        Me.dgvCost.Rows(rowDGV.Index).Selected = True
                        Return False
                    Else
                        'pastikan tipe akunnya adalah Income atau Payable
                        TipeAkun = Me.Objakun.GetTipeAkun(rowDGV.Cells(Me.ColPrimaryAccount.Index).Value)
                        If TipeAkun = Core_Module.ClsMasterAkun1.EnumTipeAkun.Income OrElse _
                        TipeAkun = Core_Module.ClsMasterAkun1.EnumTipeAkun.Expense OrElse _
                        TipeAkun = Core_Module.ClsMasterAkun1.EnumTipeAkun.Other_Income OrElse _
                        TipeAkun = Core_Module.ClsMasterAkun1.EnumTipeAkun.Other_Expense OrElse _
                        TipeAkun = Core_Module.ClsMasterAkun1.EnumTipeAkun.Cost_of_Goods_Sales Then
                        Else
                            'kalo ga maka stop
                            MsgBox("Please specify account code with 'Income', 'Other Income', 'Expense', 'Other Expernse' or 'COGS' account type for cost data on row " + (rowDGV.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                            Me.dgvCost.ClearSelection()
                            Me.dgvCost.Rows(rowDGV.Index).Selected = True
                            Return False
                        End If

                        'currency of Account harus idr, tetapi nilai nominalnya ngikut currency induk
                        If Me.Objakun.GetPrimaryCurrency(rowDGV.Cells(Me.ColPrimaryAccount.Index).Value) <> Core_Module.PrimaryBaseCurrency Then
                            MsgBox("Please specify account code with " + Me.ObjCurrency.GetCurrencyCode(Core_Module.PrimaryBaseCurrency) + _
                            " currency for cost data on row " + (rowDGV.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                            Me.dgvCost.ClearSelection()
                            Me.dgvCost.Rows(rowDGV.Index).Selected = True
                            Return False
                        End If

                    End If
                End If
            Next

            'data ongkos harus memiliki keterangan yang jelas, dan tidak boleh bernilai nol
            For Each rowDGV In Me.dgvCost.Rows
                If rowDGV.Cells(Me.ColCostAmount.Index).Value = 0 And rowDGV.Index <> Me.dgvCost.RowCount - 1 Then
                    MsgBox("Please specify cost data on row " + (rowDGV.Index + 1).ToString + ". Cost amount must be a positive number", _
                    MsgBoxStyle.Information, "Information")
                    Me.dgvCost.ClearSelection()
                    rowDGV.Selected = True
                    Me.dgvCost.Focus()
                    Return False
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahDescriptionBerulang(ByVal vDesc As String, ByVal index As Integer) As Boolean
        Dim RowDGV As New DataGridViewRow
        Try
            For Each RowDGV In Me.dgvCost.Rows
                With RowDGV
                    If Me.dgvCost.Rows.Count - 1 <> RowDGV.Index Then
                        If Not String.IsNullOrEmpty(.Cells(Me.ColDescCost.Index).Value) Then
                            If RowDGV.Index <> index AndAlso .Cells(Me.ColDescCost.Index).Value.ToString.Trim.ToUpper = vDesc.Trim.ToUpper Then
                                Return True
                            End If
                        End If
                    End If
                End With
            Next

            Return False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub txtVar1_Name_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVar1_Name.GotFocus

        Try
            If Me.txtShippingMarks.Enabled = True Then
                If Me.txtVar1_Name.ReadOnly = True Then
                    Me.txtVar1_Name.BackColor = Me.TxtCustomerDescShipTo.BackColor
                Else
                    Me.txtVar1_Name.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtVar1_Name_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVar1_Name.LostFocus

        Try
            If Me.txtShippingMarks.Enabled = True Then
                If Me.txtVar1_Name.ReadOnly = True Then
                    Me.txtVar1_Name.BackColor = Me.TxtCustomerDescShipTo.BackColor
                Else
                    Me.txtVar1_Name.BackColor = Me.TxtPrimaryCustomerShipTo.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtVar2_Name_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVar2_Name.GotFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtVar2_Name.ReadOnly = True Then
                Me.txtVar2_Name.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtVar2_Name.BackColor = Me.BackColorAll
            End If
        End If
    End Sub

    Private Sub txtVar2_Name_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVar2_Name.LostFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtVar2_Name.ReadOnly = True Then
                Me.txtVar2_Name.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtVar2_Name.BackColor = Me.TxtPrimaryCustomerShipTo.BackColor
            End If
        End If
    End Sub

    Private Sub txtVar3_Name_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVar3_Name.GotFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtVar3_Name.ReadOnly = True Then
                Me.txtVar3_Name.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtVar3_Name.BackColor = Me.BackColorAll
            End If
        End If
    End Sub

    Private Sub txtVar3_Name_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVar3_Name.LostFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtVar3_Name.ReadOnly = True Then
                Me.txtVar3_Name.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtVar3_Name.BackColor = Me.TxtPrimaryCustomerShipTo.BackColor
            End If
        End If
    End Sub

    Private Sub txtVar4_Name_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVar4_Name.GotFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtVar4_Name.ReadOnly = True Then
                Me.txtVar4_Name.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtVar4_Name.BackColor = Me.BackColorAll
            End If
        End If
    End Sub

    Private Sub txtVar4_Name_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVar4_Name.LostFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtVar4_Name.ReadOnly = True Then
                Me.txtVar4_Name.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtVar4_Name.BackColor = Me.TxtPrimaryCustomerShipTo.BackColor
            End If
        End If
    End Sub

    Private Sub txtVar5_Name_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVar5_Name.GotFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtVar5_Name.ReadOnly = True Then
                Me.txtVar5_Name.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtVar5_Name.BackColor = Me.BackColorAll
            End If
        End If
    End Sub

    Private Sub txtVar5_Name_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVar5_Name.LostFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtVar5_Name.ReadOnly = True Then
                Me.txtVar5_Name.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtVar5_Name.BackColor = Me.TxtPrimaryCustomerShipTo.BackColor
            End If
        End If
    End Sub

    Private Sub txtTransportationMode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTransportationMode.GotFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtTransportationMode.ReadOnly = True Then
                Me.txtTransportationMode.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtTransportationMode.BackColor = Me.BackColorAll
            End If
        End If
    End Sub

    Private Sub txtTransportationMode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTransportationMode.LostFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtTransportationMode.ReadOnly = True Then
                Me.txtTransportationMode.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtTransportationMode.BackColor = Me.TxtPrimaryCustomerShipTo.BackColor
            End If
        End If
    End Sub

    Private Sub txtShippingMarks_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtShippingMarks.GotFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtShippingMarks.ReadOnly = True Then
                Me.txtShippingMarks.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtShippingMarks.BackColor = Me.BackColorAll
            End If
        End If
    End Sub

    Private Sub txtShippingMarks_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtShippingMarks.LostFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtShippingMarks.ReadOnly = True Then
                Me.txtShippingMarks.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtShippingMarks.BackColor = Me.TxtPrimaryCustomerShipTo.BackColor
            End If
        End If
    End Sub

    Private Sub txtLoadingPort_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLoadingPort.GotFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtLoadingPort.ReadOnly = True Then
                Me.txtLoadingPort.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtLoadingPort.BackColor = Me.BackColorAll
            End If
        End If
    End Sub

    Private Sub txtLoadingPort_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLoadingPort.LostFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtLoadingPort.ReadOnly = True Then
                Me.txtLoadingPort.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtLoadingPort.BackColor = Me.TxtPrimaryCustomerShipTo.BackColor
            End If
        End If
    End Sub

    Private Sub txtInvoiceType_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInvoiceType.GotFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtInvoiceType.ReadOnly = True Then
                Me.txtInvoiceType.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtInvoiceType.BackColor = Me.BackColorAll
            End If
        End If
    End Sub

    Private Sub txtInvoiceType_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInvoiceType.LostFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtInvoiceType.ReadOnly = True Then
                Me.txtInvoiceType.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtInvoiceType.BackColor = Me.TxtPrimaryCustomerShipTo.BackColor
            End If
        End If
    End Sub

    Private Sub txtDestinationPort_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDestinationPort.GotFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtDestinationPort.ReadOnly = True Then
                Me.txtDestinationPort.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtDestinationPort.BackColor = Me.BackColorAll
            End If
        End If
    End Sub

    Private Sub txtDestinationPort_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDestinationPort.LostFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtDestinationPort.ReadOnly = True Then
                Me.txtDestinationPort.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtDestinationPort.BackColor = Me.TxtPrimaryCustomerShipTo.BackColor
            End If
        End If
    End Sub

    Private Sub txtUnitWeight_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUnitWeight.GotFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtUnitWeight.ReadOnly = True Then
                Me.txtUnitWeight.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtUnitWeight.BackColor = Me.BackColorAll
            End If
        End If
    End Sub

    Private Sub txtUnitWeight_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUnitWeight.KeyDown
        Dim i As New FrmCariMasterUnit
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyCode = Keys.Enter Then
                If ObjUnit.ApakahSudahAda(Me.txtUnitWeight.Text) = True Then
                    Me.txtPrimaryUnitWeight.Text = ObjUnit.GetPrimary(Me.txtUnitWeight.Text)
                    Me.txtUnitWeight.Text = ObjUnit.GetUnitCode(Me.txtPrimaryUnitWeight.Text)
                Else
                    If Me.txtPrimaryUnitWeight.Text <> "" Then
                        i.txtCari.Text = Me.txtUnitWeight.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryUnit > 0 Then
                        Me.txtPrimaryUnitWeight.Text = i.PrimaryUnit
                        Me.txtUnitWeight.Text = ObjUnit.GetUnitCode(i.PrimaryUnit)
                        Me.SelectNextControl(Me.txtUnitWeight, True, True, True, True)
                    Else
                        Me.txtPrimaryUnitWeight.Clear()
                        Me.txtPrimaryUnitWeight.Text = "-1"
                        Me.txtUnitWeight.Focus()

                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtUnitWeight_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUnitWeight.Leave
        Dim i As New FrmCariMasterUnit
        Try
            If Me.txtUnitWeight.Text <> "" Then
                If ObjUnit.ApakahSudahAda(Me.txtUnitWeight.Text) = True Then
                    Me.txtPrimaryUnitWeight.Text = ObjUnit.GetPrimary(Me.txtUnitWeight.Text)
                    Me.txtUnitWeight.Text = ObjUnit.GetUnitCode(Me.txtPrimaryUnitWeight.Text)
                Else
                    If Me.txtUnitWeight.Text <> "" Then
                        i.txtCari.Text = Me.txtUnitWeight.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryUnit > 0 Then
                        Me.txtPrimaryUnitWeight.Text = i.PrimaryUnit
                        Me.txtUnitWeight.Text = ObjUnit.GetUnitCode(Me.txtPrimaryUnitWeight.Text)
                    Else
                        Me.txtUnitWeight.Clear()
                        Me.txtPrimaryUnitWeight.Text = "-1"
                    End If
                End If
            Else
                Me.txtPrimaryUnitWeight.Text = "-1"
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtUnitWeight_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUnitWeight.LostFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtUnitWeight.ReadOnly = True Then
                Me.txtUnitWeight.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtUnitWeight.BackColor = Me.TxtPrimaryCustomerShipTo.BackColor
            End If
        End If
    End Sub

    Private Sub txtUnitVolume_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUnitVolume.GotFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtUnitVolume.ReadOnly = True Then
                Me.txtUnitVolume.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtUnitVolume.BackColor = Me.BackColorAll
            End If
        End If
    End Sub

    Private Sub txtUnitVolume_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUnitVolume.KeyDown
        Dim i As New FrmCariMasterUnit
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyCode = Keys.Enter Then
                If ObjUnit.ApakahSudahAda(Me.txtUnitVolume.Text) = True Then
                    Me.txtPrimaryUnitVolume.Text = ObjUnit.GetPrimary(Me.txtUnitVolume.Text)
                    Me.txtUnitVolume.Text = ObjUnit.GetUnitCode(Me.txtPrimaryUnitVolume.Text)
                Else
                    If Me.txtPrimaryUnitVolume.Text <> "" Then
                        i.txtCari.Text = Me.txtUnitVolume.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryUnit > 0 Then
                        Me.txtPrimaryUnitVolume.Text = i.PrimaryUnit
                        Me.txtUnitVolume.Text = ObjUnit.GetUnitCode(i.PrimaryUnit)
                        Me.SelectNextControl(Me.txtUnitVolume, True, True, True, True)
                    Else
                        Me.txtPrimaryUnitVolume.Clear()
                        Me.txtPrimaryUnitVolume.Text = "-1"
                        Me.txtUnitVolume.Focus()

                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtUnitVolume_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUnitVolume.Leave
        Dim i As New FrmCariMasterUnit
        Try
            If Me.txtUnitVolume.Text <> "" Then
                If ObjUnit.ApakahSudahAda(Me.txtUnitVolume.Text) = True Then
                    Me.txtPrimaryUnitVolume.Text = ObjUnit.GetPrimary(Me.txtUnitVolume.Text)
                    Me.txtUnitVolume.Text = ObjUnit.GetUnitCode(Me.txtPrimaryUnitVolume.Text)
                Else
                    If Me.txtUnitVolume.Text <> "" Then
                        i.txtCari.Text = Me.txtUnitVolume.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryUnit > 0 Then
                        Me.txtPrimaryUnitVolume.Text = i.PrimaryUnit
                        Me.txtUnitVolume.Text = ObjUnit.GetUnitCode(i.PrimaryUnit)
                    Else
                        Me.txtUnitVolume.Clear()
                        Me.txtPrimaryUnitVolume.Text = "-1"
                    End If
                End If
            Else
                Me.txtPrimaryUnitVolume.Text = "-1"
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtUnitVolume_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUnitVolume.LostFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtUnitVolume.ReadOnly = True Then
                Me.txtUnitVolume.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtUnitVolume.BackColor = Me.TxtPrimaryCustomerShipTo.BackColor
            End If
        End If
    End Sub

    Private Sub txtGrossVolume_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGrossVolume.GotFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtGrossVolume.ReadOnly = True Then
                Me.txtGrossVolume.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtGrossVolume.BackColor = Me.BackColorAll
            End If
        End If
    End Sub

    Private Sub txtGrossVolume_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtGrossVolume.KeyDown
        Me.ObjCD.Number_KeyDown(Me.txtGrossVolume, sender, e)
    End Sub

    Private Sub txtGrossVolume_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtGrossVolume.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtGrossVolume, sender, e)
    End Sub

    Private Sub txtGrossVolume_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGrossVolume.LostFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtGrossVolume.ReadOnly = True Then
                Me.txtGrossVolume.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtGrossVolume.BackColor = Me.TxtPrimaryCustomerShipTo.BackColor
            End If
        End If
    End Sub

    Private Sub txtGrossWeight_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGrossWeight.GotFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtGrossWeight.ReadOnly = True Then
                Me.txtGrossWeight.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtGrossWeight.BackColor = Me.BackColorAll
            End If
        End If
    End Sub

    Private Sub txtGrossWeight_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtGrossWeight.KeyDown
        Me.ObjCD.Number_KeyDown(Me.txtGrossWeight, sender, e)

    End Sub

    Private Sub txtGrossWeight_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtGrossWeight.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtGrossWeight, sender, e)
    End Sub

    Private Sub txtGrossWeight_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGrossWeight.LostFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtGrossWeight.ReadOnly = True Then
                Me.txtGrossWeight.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtGrossWeight.BackColor = Me.TxtPrimaryCustomerShipTo.BackColor
            End If
        End If
    End Sub

    Private Sub txtNettVolume_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNettVolume.GotFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtNettVolume.ReadOnly = True Then
                Me.txtNettVolume.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtNettVolume.BackColor = Me.BackColorAll
            End If
        End If
    End Sub

    Private Sub txtNettVolume_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtNettVolume.KeyDown
        Me.ObjCD.Number_KeyDown(Me.txtNettVolume, sender, e)
    End Sub

    Private Sub txtNettVolume_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtNettVolume.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtNettVolume, sender, e)
    End Sub

    Private Sub txtNettVolume_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNettVolume.LostFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtNettVolume.ReadOnly = True Then
                Me.txtNettVolume.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtNettVolume.BackColor = Me.TxtPrimaryCustomerShipTo.BackColor
            End If
        End If
    End Sub

    Private Sub txtNettWeight_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNettWeight.GotFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtNettWeight.ReadOnly = True Then
                Me.txtNettWeight.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtNettWeight.BackColor = Me.BackColorAll
            End If
        End If
    End Sub

    Private Sub txtNettWeight_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtNettWeight.KeyDown
        Me.ObjCD.Number_KeyDown(Me.txtNettWeight, sender, e)
    End Sub

    Private Sub txtNettWeight_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtNettWeight.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtNettWeight, sender, e)
    End Sub

    Private Sub txtNettWeight_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNettWeight.LostFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtNettWeight.ReadOnly = True Then
                Me.txtNettWeight.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtNettWeight.BackColor = Me.TxtPrimaryCustomerShipTo.BackColor
            End If
        End If
    End Sub

    Private Sub txtVar1_Value_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVar1_Value.GotFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtVar1_Value.ReadOnly = True Then
                Me.txtVar1_Value.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtVar1_Value.BackColor = Me.BackColorAll
            End If
        End If
    End Sub

    Private Sub txtVar1_Value_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVar1_Value.LostFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtVar1_Value.ReadOnly = True Then
                Me.txtVar1_Value.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtVar1_Value.BackColor = Me.TxtPrimaryCustomerShipTo.BackColor
            End If
        End If
    End Sub

    Private Sub txtVar2_Value_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVar2_Value.GotFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtVar2_Value.ReadOnly = True Then
                Me.txtVar2_Value.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtVar2_Value.BackColor = Me.BackColorAll
            End If
        End If
    End Sub

    Private Sub txtVar2_Value_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVar2_Value.LostFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtVar2_Value.ReadOnly = True Then
                Me.txtVar2_Value.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtVar2_Value.BackColor = Me.TxtPrimaryCustomerShipTo.BackColor
            End If
        End If
    End Sub

    Private Sub txtVar3_Value_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVar3_Value.GotFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtVar3_Value.ReadOnly = True Then
                Me.txtVar3_Value.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtVar3_Value.BackColor = Me.BackColorAll
            End If
        End If
    End Sub

    Private Sub txtVar3_Value_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVar3_Value.LostFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtVar3_Value.ReadOnly = True Then
                Me.txtVar3_Value.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtVar3_Value.BackColor = Me.TxtPrimaryCustomerShipTo.BackColor
            End If
        End If
    End Sub

    Private Sub txtVar4_Value_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVar4_Value.GotFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtVar4_Value.ReadOnly = True Then
                Me.txtVar4_Value.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtVar4_Value.BackColor = Me.BackColorAll
            End If
        End If
    End Sub

    Private Sub txtVar4_Value_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVar4_Value.LostFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtVar4_Value.ReadOnly = True Then
                Me.txtVar4_Value.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtVar4_Value.BackColor = Me.TxtPrimaryCustomerShipTo.BackColor
            End If
        End If
    End Sub

    Private Sub txtVar5_Value_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVar5_Value.GotFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtVar5_Value.ReadOnly = True Then
                Me.txtVar5_Value.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtVar5_Value.BackColor = Me.BackColorAll
            End If
        End If
    End Sub

    Private Sub txtVar5_Value_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVar5_Value.LostFocus
        If Me.txtShippingMarks.Enabled = True Then
            If Me.txtVar5_Value.ReadOnly = True Then
                Me.txtVar5_Value.BackColor = Me.TxtCustomerDescShipTo.BackColor
            Else
                Me.txtVar5_Value.BackColor = Me.TxtPrimaryCustomerShipTo.BackColor
            End If
        End If
    End Sub

    Private Sub btnFindUnitVolume_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindUnitVolume.Click
        Dim i As New FrmCariMasterUnit
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.ObjUnit.ApakahSudahAda(Me.txtUnitVolume.Text) = True Then
                Me.txtPrimaryUnitVolume.Text = ObjUnit.GetPrimary(Me.txtUnitVolume.Text)
            Else
                i.cboCriteria.SelectedIndex = 0
                If Me.txtUnitVolume.Text <> "" Then
                    i.txtCari.Text = Me.txtUnitVolume.Text
                End If

                i.ShowDialog()
                If i.PrimaryUnit > 0 Then
                    Me.txtPrimaryUnitVolume.Text = i.PrimaryUnit
                    Me.txtUnitVolume.Text = ObjUnit.GetUnitCode(i.PrimaryUnit)
                Else
                    Me.txtUnitVolume.Clear()
                    Me.txtUnitVolume.Focus()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindUnitWeight_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindUnitWeight.Click
        Dim i As New FrmCariMasterUnit
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.ObjUnit.ApakahSudahAda(Me.txtUnitWeight.Text) = True Then
                Me.txtPrimaryUnitWeight.Text = ObjUnit.GetPrimary(Me.txtUnitWeight.Text)
            Else
                i.cboCriteria.SelectedIndex = 0
                If Me.txtUnitWeight.Text <> "" Then
                    i.txtCari.Text = Me.txtUnitWeight.Text
                End If

                i.ShowDialog()
                If i.PrimaryUnit > 0 Then
                    Me.txtPrimaryUnitWeight.Text = i.PrimaryUnit
                    Me.txtUnitWeight.Text = ObjUnit.GetUnitCode(i.PrimaryUnit)
                Else
                    Me.txtUnitWeight.Clear()
                    Me.txtUnitWeight.Focus()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.Aktif = True Then
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Me.TBLHeaderCost = Me.TBLHeaderCost
                    Me.TBLitemCost = Me.TBLitemCost
                    Me.OtherCharge = Me.OtherChargeDummy
                Else
                    Exit Sub
                End If
            Else
                HitungOtherCharge()
            End If
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCost_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvCost.CellBeginEdit
        Try
            'yang bisa diedit adalah:
            '1. currency
            '2. cost nominal
            '3. exchange rate
            '4. remarks
            '5. description
            '6. Account --> hanya jika sudah lewat convertion date
            If e.ColumnIndex = Me.ColRemarksCost.Index Or _
            e.ColumnIndex = Me.ColDescCost.Index Or _
            e.ColumnIndex = Me.ColCostNominal.Index Then
                Me.Buffer = Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value
            ElseIf e.ColumnIndex = Me.ColExcRateCost.Index Then
                If Me.dgvCost.Item(Me.ColPrimaryCurrencyCost.Index, e.RowIndex).Value = 1 Then
                    'jika sudah IDR, maka exchange rate untuk cost tidak bisa diubah lagi
                    e.Cancel = True
                End If
                If e.ColumnIndex = Me.ColCostNominal.Index Or e.ColumnIndex = Me.ColExcRateCost.Index Then
                    'jika kolom yg diedit adalah exchange rate dan cost nominal maka buang decimal
                    Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = _
                    Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value)
                End If
            ElseIf e.ColumnIndex = Me.ColAccounrCode.Index Then
                Me.Buffer = Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCost_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCost.CellEndEdit
        Dim DefPrec As Integer = 0
        Dim PrimaryCurrencyAkun As Double = -1

        Try
            If IsNothing(Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                'jika tidak ada perubahan maka tidak perlu lakukan pengecekan
                If Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                    Exit Sub
                End If

                If e.ColumnIndex = Me.ColCurrencyCost.Index Then
                    'end edit currency
                    If Me.ObjCurrency.ApakahSudahAda(Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Dim i As New FrmCariMasterCurrency
                        If Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value <> "" Then
                            i.txtCari.Text = Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value
                        End If

                        i.ShowDialog()
                        If i.PrimaryCurrency <> -1 Then
                            Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.ObjCurrency.GetCurrencyCode(i.PrimaryCurrency)).ToUpper
                            Me.dgvCost.Item(Me.ColPrimaryCurrencyCost.Index, e.RowIndex).Value = i.PrimaryCurrency
                            Me.dgvCost.Item(Me.ColExcRateCost.Index, e.RowIndex).Value = 1
                            'format ulang cost
                            DefPrec = Me.ObjCurrency.GetDefaultPrecission(Me.dgvCost.Item(Me.ColPrimaryCurrencyCost.Index, e.RowIndex).Value)
                            Me.dgvCost.Item(Me.ColCostNominal.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvCost.Item(Me.ColCostNominal.Index, e.RowIndex).Value, _
                            DefPrec)
                        Else
                            Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        i = Nothing
                    Else
                        Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                        Me.dgvCost.Item(Me.ColPrimaryCurrencyCost.Index, e.RowIndex).Value = Me.ObjCurrency.GetPrimary(Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvCost.Item(Me.ColExcRateCost.Index, e.RowIndex).Value = 1
                        'format ulang cost
                        DefPrec = Me.ObjCurrency.GetDefaultPrecission(Me.dgvCost.Item(Me.ColPrimaryCurrencyCost.Index, e.RowIndex).Value)
                        Me.dgvCost.Item(Me.ColCostNominal.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvCost.Item(Me.ColCostNominal.Index, e.RowIndex).Value, _
                        DefPrec)
                    End If
                ElseIf e.ColumnIndex = Me.ColDescCost.Index Then
                    'deskripsi
                    'tidak boleh kosong
                    If IsNothing(Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = True Then
                        Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    If Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value.ToString = "" Then
                        Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                ElseIf e.ColumnIndex = Me.ColExcRateCost.Index Then
                    'perhitungan cost amount
                    'harus positif nilainnya
                    If Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 15 Then
                        Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    If Me.ObjCD.CekBilanganPositif(Me.dgvCost.Item(e.ColumnIndex, _
                    e.RowIndex).Value) = False Then
                        Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    Else
                        'format tampilan
                        Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = _
                        Me.ObjCD.DigitGrouping(Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value, Core_Module.DefPrecExchangeRate)
                    End If
                ElseIf e.ColumnIndex = Me.ColCostNominal.Index Then
                    If Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 15 Then
                        Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Me.dgvCost.Item(Me.ColExcRateCost.Index, e.RowIndex).Value = 1
                        Me.dgvCost.Item(Me.ColCostAmount.Index, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    If Me.ObjCD.CekBilanganPositif(Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Me.dgvCost.Item(Me.ColExcRateCost.Index, e.RowIndex).Value = 1
                        Me.dgvCost.Item(Me.ColCostAmount.Index, e.RowIndex).Value = Me.Buffer
                    Else
                        'format tampilan
                        Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = _
                        Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value, _
                        Me.ObjCurrency.GetDefaultPrecission(Me.PrimaryCurrency)))
                        Me.dgvCost.Item(Me.ColExcRateCost.Index, e.RowIndex).Value = 1
                        Me.dgvCost.Item(Me.ColCostAmount.Index, e.RowIndex).Value = _
                        Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value, _
                        Me.ObjCurrency.GetDefaultPrecission(Me.PrimaryCurrency)))
                    End If

                ElseIf e.ColumnIndex = Me.ColAccounrCode.Index Then
                    Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    'kemungkinan
                    If Me.ApakahPengulanganDetail(Me.Objakun.GetPrimary(Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value), e.RowIndex) = False Then
                        If Me.Objakun.ApakahSudahAda_Aktif_NonHeader(Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            Me.dgvCost.Item(Me.ColPrimaryAccount.Index, e.RowIndex).Value = Me.Objakun.GetPrimary(Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value)
                            Me.dgvCost.Item(Me.ColAccountDesc.Index, e.RowIndex).Value = Me.Objakun.GetDescription_ByPrimaryAcc(Me.dgvCost.Item(Me.ColPrimaryAccount.Index, e.RowIndex).Value)
                        Else
                            '2. kode akun yang diketik tidak ada --> tampilkan form cari
                            Dim ij As New Core_Module.FrmMasterAkun_Cari1
                            ij.BackColorAll = Me.BackColorAll
                            ij.txtCari.Text = Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.ShowDialog()
                            If ij.mPrimaryAkun <> -1 Then
                                '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                                Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = Me.Objakun.GetAccountCode(ij.mPrimaryAkun)
                                Me.dgvCost.Item(Me.ColAccountDesc.Index, e.RowIndex).Value = Me.Objakun.GetDescription_ByPrimaryAcc(ij.mPrimaryAkun)
                                Me.dgvCost.Item(Me.ColPrimaryAccount.Index, e.RowIndex).Value = ij.mPrimaryAkun
                            ElseIf ij.mPrimaryAkun = -1 Then
                                '2.b. kode akun yang dicari tidak ada --> kembalikan ke kondisi buffer
                                If Me.Buffer <> "" AndAlso CDec(Me.dgvCost.Item(Me.ColPrimaryAccount.Index, e.RowIndex).Value) > 0 Then
                                    Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Else
                                    Me.dgvCost.Rows.RemoveAt(e.RowIndex)
                                End If
                            End If
                            ij = Nothing
                        End If
                    Else
                        If Me.Buffer <> "" Then
                            Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Else
                            Me.dgvCost.Rows.RemoveAt(e.RowIndex)
                        End If
                    End If
                End If
            Else
                If e.ColumnIndex = Me.ColAccounrCode.Index Then
                    Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Me.dgvCost.Item(Me.ColAccountDesc.Index, e.RowIndex).Value = ""
                    Me.dgvCost.Item(Me.ColPrimaryAccount.Index, e.RowIndex).Value = -1
                Else
                    Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahPengulanganDetail(ByVal vPrimaryAccount As Double, ByVal Index As Integer) As Boolean
        Dim rowDGV As New DataGridViewRow
        Try
            'For Each rowDGV In Me.dgvCost.Rows
            '    With rowDGV
            '        If rowDGV.Index <> Index AndAlso vPrimaryAccount <> -1 Then
            '            If vPrimaryAccount = .Cells(Me.ColPrimaryAccount.Index).Value Then
            '                Dim strAccount As String = Me.Objakun.GetDescription_ByPrimaryAcc(vPrimaryAccount)
            '                MsgBox("Account code duplication, please choose another account. ", MsgBoxStyle.Information, "Information")
            '                Me.dgvCost.ClearSelection()
            '                Me.dgvCost.Rows(rowDGV.Index).Selected = True
            '                Return True
            '            End If
            '        End If
            '    End With
            'Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvCost_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvCost.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvCost)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCost_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvCost.RowsAdded
        Dim DefPrecBase As Integer = 0
        Try
            DefPrecBase = Me.ObjCurrency.GetDecimal_DefaultCurrency

            If Me.dgvCost.AllowUserToAddRows = True Then
                With Me.dgvCost.Rows(Me.dgvCost.Rows.Count - 1)
                    .Cells(Me.ColDescCost.Index).Value = "-"
                    .Cells(Me.ColCurrencyCost.Index).Value = Me.ObjCurrency.GetCurrencyCode(Me.PrimaryCurrency)
                    .Cells(Me.ColPrimaryCurrencyCost.Index).Value = Core_Module.PrimaryBaseCurrency
                    .Cells(Me.ColCostNominal.Index).Value = Me.ObjCD.DigitGrouping("0", DefPrecBase)
                    .Cells(Me.ColRemarksCost.Index).Value = ""
                    .Cells(Me.ColExcRateCost.Index).Value = Me.ObjCD.DigitGrouping("1", Core_Module.DefPrecExchangeRate)
                    .Cells(Me.ColCostAmount.Index).Value = Me.ObjCD.DigitGrouping("0", DefPrecBase)
                    .Cells(Me.ColAccounrCode.Index).Value = ""
                    .Cells(Me.ColAccountDesc.Index).Value = ""
                    .Cells(Me.ColPrimaryAccount.Index).Value = -1
                End With
            End If
            ObjCD.SetNumber(Me.dgvCost)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCost_UserAddedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvCost.UserAddedRow
        Dim DefPrecBase As Integer = 0
        Try
            DefPrecBase = Me.ObjCurrency.GetDefaultPrecission(Me.PrimaryCurrency)
            With Me.dgvCost.Rows(Me.dgvCost.Rows.Count - 1)
                .Cells(Me.ColDescCost.Index).Value = "-"
                .Cells(Me.ColCurrencyCost.Index).Value = Me.ObjCurrency.GetCurrencyCode(Me.PrimaryCurrency)
                .Cells(Me.ColPrimaryCurrencyCost.Index).Value = Core_Module.PrimaryBaseCurrency
                .Cells(Me.ColCostNominal.Index).Value = Me.ObjCD.DigitGrouping("0", DefPrecBase)
                .Cells(Me.ColRemarksCost.Index).Value = ""
                .Cells(Me.ColExcRateCost.Index).Value = Me.ObjCD.DigitGrouping("1", Core_Module.DefPrecExchangeRate)
                .Cells(Me.ColCostAmount.Index).Value = Me.ObjCD.DigitGrouping("0", DefPrecBase)
                .Cells(Me.ColAccounrCode.Index).Value = ""
                .Cells(Me.ColAccountDesc.Index).Value = ""
                .Cells(Me.ColPrimaryAccount.Index).Value = -1
            End With
            ObjCD.SetNumber(Me.dgvCost)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbViewAddressCustomerShipTo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LlbViewAddressCustomerShipTo.Click
        Try
            If IsNumeric(Me.TxtPrimaryCustomerShipTo.Text) AndAlso CDec(Me.TxtPrimaryCustomerShipTo.Text) > 0 Then
                If Me.TxtPrimaryCustomerShipTo.Text = "" Or Me.TxtPrimaryCustomerShipTo.Text = "-1" Then
                    MsgBox("Please specify customer first.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                Dim ij As New Core_Module.FrmDetailAddress
                ij.vAddIdentifier = Me.TxtIdentifierAddressCustomerShipTo.Text
                ij.PrimaryCust = Me.TxtPrimaryCustomerShipTo.Text

                ij.cboCustAddCode.Enabled = False

                ij.AcceptButton = ij.btnOK
                ij.btnExit.TabStop = False
                ij.ShowDialog()
                ij.Dispose()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class