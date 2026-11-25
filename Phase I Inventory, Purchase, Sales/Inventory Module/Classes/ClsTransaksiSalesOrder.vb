Option Explicit On
Option Strict Off

Public Class ClsTransaksiSalesOrder
    Dim ObjCD As New Core_Module.ClsCekData

    Public Enum EnumJenisFakturPajak As Integer
        FakturPajakSederhana
        FakturPajakStandar
    End Enum

    Public Function GetDeliveryDateMinimum(ByVal vPrimarySO As Double) As Date
        Dim DTA As New DseTransaksiSO_OrderBalanceTableAdapters.tbltrans_so_itemdeliveryTableAdapter
        Dim TBL As New DseTransaksiSO_OrderBalance.tbltrans_so_itemdeliveryDataTable
        Dim row As DseTransaksiSO_OrderBalance.tbltrans_so_itemdeliveryRow
        Dim RetVal As Date = Date.MaxValue

        Try
            'mengembalikan delivery date minimum untuk sales order yang diberikan primarynya
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimarySO(TBL, vPrimarySO)

            If TBL.Rows.Count = 0 Then
                Return #1/1/1950#
            Else
                For Each row In TBL.Rows
                    If RetVal.Date > row.delivery_date And row.delivery_date <> #1/1/1950# Then
                        RetVal = row.delivery_date
                    End If
                Next
                If RetVal <> Date.MaxValue Then
                    Return RetVal
                Else
                    Return #1/1/1950#
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        End Try
    End Function

    Public Function GetTanggalTransaksi(ByVal vIdentifierItemSO As String) As Date
        Dim DTA As New dseTransaksiSOTableAdapters.tbltrans_soTableAdapter
        Dim RetVal As Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetDateByIdentifier(vIdentifierItemSO)

            If RetVal.HasValue = True Then
                Return RetVal
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahAdaAROverdue(ByVal vPrimaryCustomer As Double, ByVal vDueDate As Date) As Boolean
        Dim DTA As New dseTransaksiSOTableAdapters.tbltrans_balance_salesTableAdapter
        Dim TBL As New dseTransaksiSO.tbltrans_balance_salesDataTable
        Dim TipeBalance_Inv As ClsTransaksiSalesInvoiceItem.EnumTipeSalesBalance = ClsTransaksiSalesInvoiceItem.EnumTipeSalesBalance.Automatic
        Dim TipeBalance_SR As ClsTransaksiSalesInvoiceItem.EnumTipeSalesBalance = ClsTransaksiSalesInvoiceItem.EnumTipeSalesBalance.Automatic_SR

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryCustomer, vDueDate)

            If TBL.Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetAR(ByVal vPrimaryCustomer As Double) As Decimal
        Dim DTA As New dseTransaksiSOTableAdapters.tbltrans_balance_salesTableAdapter
        Dim TBL As New dseTransaksiSO.tbltrans_balance_salesDataTable
        Dim row As dseTransaksiSO.tbltrans_balance_salesRow
        Dim TipeBalance_Inv As ClsTransaksiSalesInvoiceItem.EnumTipeSalesBalance = ClsTransaksiSalesInvoiceItem.EnumTipeSalesBalance.Automatic
        Dim TipeBalance_SR As ClsTransaksiSalesInvoiceItem.EnumTipeSalesBalance = ClsTransaksiSalesInvoiceItem.EnumTipeSalesBalance.Automatic_SR
        Dim RetVal As Decimal = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_All(TBL, vPrimaryCustomer)
            For Each row In TBL.Rows
                RetVal += row.Balance
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Sub IsiDGVSalesOrderBalance(ByVal vPrimaryProduct As Double, ByVal vTBL As DseTransaksiSO_OrderBalance.tbltrans_so_item1DataTable, _
   ByVal vDGV As DataGridView)
        Dim DTA As New DseTransaksiSO_OrderBalanceTableAdapters.tbltrans_so_item1TableAdapter
        Dim row As DseTransaksiSO_OrderBalance.tbltrans_so_item1Row
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DTA2 As New DseTransaksiSO_OrderBalanceTableAdapters.tbltrans_so_itemdeliveryTableAdapter
        Dim Tbl2 As New DseTransaksiSO_OrderBalance.tbltrans_so_itemdeliveryDataTable
        Dim row2 As DseTransaksiSO_OrderBalance.tbltrans_so_item1Row
        Dim PrimaryMin As Double = -1
        Dim PrimaryMax As Double = -1
        Dim rowDGV As DataGridViewRow
        Dim vFrm As New FrmMasterKondisiStock

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGV.DataSource = Nothing

            If vPrimaryProduct = -1 Then
                DTA.Fill(vTBL)
            Else
                DTA.FillBy_PrimaryProduct(vTBL, vPrimaryProduct)
            End If

            vTBL.DefaultView.RowFilter = Nothing
            vTBL.DefaultView.Sort = "[Primary] ASC"
            If vTBL.DefaultView.Count > 0 Then
                row2 = vTBL.DefaultView(0).Row
                PrimaryMin = row2.Primary
                row2 = vTBL.DefaultView(vTBL.DefaultView.Count - 1).Row
                PrimaryMax = row2.Primary
                If PrimaryMin <> -1 Or PrimaryMax <> -1 Then
                    DTA2.FillBy_PrimaryMinMax(Tbl2, PrimaryMin, PrimaryMax)
                End If
            End If

            For Each row In vTBL.Rows
                If row.Delivery_Date = #1/1/1950# Then
                    If Tbl2.Count > 0 Then
                        Tbl2.DefaultView.RowFilter = Nothing
                        Tbl2.DefaultView.RowFilter = "[Primary Transaksi] = " & row.Primary
                        If Tbl2.DefaultView.Count > 0 Then
                            row.Promise_To_Deliver = "Multiple"
                        Else
                            row.Promise_To_Deliver = "-"
                        End If
                    Else
                        row.Promise_To_Deliver = "-"
                    End If
                Else
                    row.Promise_To_Deliver = Format(row.Delivery_Date, "dd MMM yyyy")
                End If
                row.Order_Balance = ObjCD.DigitGrouping(row.SOB, row._Decimal)

                vDGV.Rows.Add()
                rowDGV = vDGV.Rows(vDGV.Rows.Count - 1)
                With rowDGV
                    .Cells(vFrm.ColSOB_SONo.Index).Value = row.SO_Number
                    .Cells(vFrm.ColSOB_CustPO.Index).Value = row.Customer_PO_Number
                    .Cells(vFrm.ColSOB_Date.Index).Value = Format(row._Date, "dd MMM yyyy")
                    .Cells(vFrm.ColSOB_CustomerCode.Index).Value = row.Customer_Code
                    .Cells(vFrm.ColSOB_CustomerName.Index).Value = row.Customer_Name
                    .Cells(vFrm.ColSOB_OrderBalance.Index).Value = row.Order_Balance
                    .Cells(vFrm.ColSOB_PromiseToDeliver.Index).Value = row.Promise_To_Deliver
                    .Cells(vFrm.ColSOB_PrimarySO.Index).Value = row.Primary
                    .Cells(vFrm.ColSOB_IdentifierSOItem.Index).Value = row.identifier
                    .Cells(vFrm.ColSOB_PrimaryUnitMulti.Index).Value = row.Primary_Units_Multi
                End With
            Next

            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Sub

    Public Sub Set_Order_Balance(ByVal vListPrimarySO As List(Of Double))
        Dim PrimarySO As Double
        Try
            For Each PrimarySO In vListPrimarySO
                Me.Set_Order_Balance(PrimarySO)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub Set_Order_Balance(ByVal vPrimarySO As Double)
        Dim TBLSO As New DseTransaksiSO_OrderBalance.tbltrans_so_itemDataTable
        Dim TBLDO As New DseTransaksiSO_OrderBalance.tbltrans_do_itemDataTable
        Dim TBLSOC As New DseTransaksiSO_OrderBalance.DataTable1DataTable
        Dim TBLSR As New DseTransaksiSO_OrderBalance.tbltrans_salesreturn_itemDataTable
        Dim TBLDOLoss As New DseTransaksiSO_OrderBalance.tbltrans_do_loss_itemDataTable
        Dim DTASO As New DseTransaksiSO_OrderBalanceTableAdapters.tbltrans_so_itemTableAdapter
        Dim DTADO As New DseTransaksiSO_OrderBalanceTableAdapters.tbltrans_do_itemTableAdapter
        Dim DTASOC As New DseTransaksiSO_OrderBalanceTableAdapters.DataTable1TableAdapter
        Dim DTASR As New DseTransaksiSO_OrderBalanceTableAdapters.tbltrans_salesreturn_itemTableAdapter
        Dim DTADOLoss As New DseTransaksiSO_OrderBalanceTableAdapters.tbltrans_do_loss_itemTableAdapter
        Dim rowSO As DseTransaksiSO_OrderBalance.tbltrans_so_itemRow
        Dim JmlDO_Multi, JmlSOC_Multi, JmlSR_Multi, JmlDOLoss_Multi, SOB_Multi, SOB As Decimal
        'Dim JmlDO, JmlSOC, JmlSR, JmlDOLoss, SOB As Decimal

        Try
            'fill database
            DTASO.FillBy_PrimarySO(TBLSO, vPrimarySO)
            DTADO.FillBy_PrimarySO(TBLDO, vPrimarySO)
            DTASOC.FillBy_PrimarySO(TBLSOC, vPrimarySO)
            DTASR.FillBy(TBLSR, vPrimarySO)
            DTADOLoss.FillBy_PrimarySO(TBLDOLoss, vPrimarySO)

            For Each rowSO In TBLSO.Rows
                With rowSO
                    JmlDOLoss_Multi = Me.GetJmlDOLoss(.Primary_Transaksi, .Primary_Product, TBLDOLoss)
                    JmlDO_Multi = Me.GetJmlDO(.Primary_Transaksi, .Primary_Product, TBLDO)
                    JmlSR_Multi = Me.GetJmlSR(.Primary_Product, .Primary_Transaksi, TBLSR)
                    JmlSOC_Multi = Me.GetJmlSOC(.Primary_Transaksi, .Primary_Product, TBLSOC)
                    SOB_Multi = .Qty_Multi - JmlDO_Multi - JmlSOC_Multi + JmlSR_Multi - JmlDOLoss_Multi
                    If .Ratio_Default = 0 Then
                        .Ratio_Default = 1
                    End If
                    If .Ratio_Reference = 0 Then
                        .Ratio_Reference = 1
                    End If
                    SOB = Math.Round(CDec(SOB_Multi * .Ratio_Default / .Ratio_Reference), CInt(._Decimal), MidpointRounding.AwayFromZero)
                    DTASO.UpdateQuery_SetSOB(SOB, SOB_Multi, .Primary)
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLSO.Dispose()
            TBLDO.Dispose()
            TBLSOC.Dispose()
            TBLSR.Dispose()
            TBLDOLoss.Dispose()
            DTASO.Dispose()
            DTADO.Dispose()
            DTASOC.Dispose()
            DTASR.Dispose()
            DTADOLoss.Dispose()
        End Try
    End Sub

    Public Function GetRatioDesc(ByVal vPrimaryProduct As Double, ByVal vPrimaryPrimaryUnit As Double) As String
        Dim DTA As New DseTransaksiSO_OrderBalanceTableAdapters.tbltrans_so_item1TableAdapter
        Dim Retval As String = ""

        Try
            Retval = DTA.ScalarQuery_GetRatioDesc(vPrimaryProduct, vPrimaryPrimaryUnit)
            If String.IsNullOrEmpty(Retval) = False Then
                Return Retval
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function GetJmlDOLoss(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double, _
      ByVal TBLDOLoss As DseTransaksiSO_OrderBalance.tbltrans_do_loss_itemDataTable) As Decimal
        Dim rowView As Data.DataRowView
        Dim rowDOLoss As DseTransaksiSO_OrderBalance.tbltrans_do_loss_itemRow
        Dim Filter As String = "[Primary SO] =" + vPrimarySO.ToString _
        + " AND [Primary Product] =" + vPrimaryProduct.ToString
        Dim Dummy As Decimal = 0

        Try
            TBLDOLoss.DefaultView.RowFilter = Filter
            If TBLDOLoss.DefaultView.Count > 0 Then
                For Each rowView In TBLDOLoss.DefaultView
                    rowDOLoss = rowView.Row
                    Dummy += rowDOLoss.Total_Qty_Multi
                Next

                Return Dummy
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDO(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double, _
    ByVal TBLDO As DseTransaksiSO_OrderBalance.tbltrans_do_itemDataTable) As Decimal
        Dim rowView As Data.DataRowView
        Dim rowDO As DseTransaksiSO_OrderBalance.tbltrans_do_itemRow
        Dim Filter As String = "[Primary SO] =" + vPrimarySO.ToString _
        + " AND [Primary Product] =" + vPrimaryProduct.ToString
        Dim Dummy As Decimal = 0

        Try
            TBLDO.DefaultView.RowFilter = Filter
            If TBLDO.DefaultView.Count > 0 Then
                For Each rowView In TBLDO.DefaultView
                    rowDO = rowView.Row
                    Dummy += rowDO.Total_Qty_Multi
                Next

                Return Dummy
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlSOC(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double, _
    ByVal TBLSOC As DseTransaksiSO_OrderBalance.DataTable1DataTable) As Decimal
        Dim rowView As Data.DataRowView
        Dim rowSOC As DseTransaksiSO_OrderBalance.DataTable1Row
        Dim Filter As String = "[Primary SO] =" + vPrimarySO.ToString _
        + " AND [Primary Product] =" + vPrimaryProduct.ToString
        Dim Dummy As Decimal = 0

        Try
            TBLSOC.DefaultView.RowFilter = Filter
            If TBLSOC.DefaultView.Count > 0 Then
                For Each rowView In TBLSOC.DefaultView
                    rowSOC = rowView.Row
                    Dummy += rowSOC.Total_Qty_Multi
                Next

                Return Dummy
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlSR(ByVal vPrimaryProduct As Double, ByVal vPrimarySO As Double, _
  ByVal TBLSR As DseTransaksiSO_OrderBalance.tbltrans_salesreturn_itemDataTable) As Decimal
        Dim rowView As Data.DataRowView
        Dim rowSR As DseTransaksiSO_OrderBalance.tbltrans_salesreturn_itemRow
        Dim Filter As String = "[Primary SO] =" + vPrimarySO.ToString _
        + " AND [Primary Product] =" + vPrimaryProduct.ToString
        Dim Dummy As Decimal = 0

        Try
            TBLSR.DefaultView.RowFilter = Filter
            If TBLSR.DefaultView.Count > 0 Then
                For Each rowView In TBLSR.DefaultView
                    rowSR = rowView.Row
                    Dummy += rowSR.Total_Qty_Multi
                Next

                Return Dummy
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub GetPajakNJenisFakturNPrimaryCurrency(ByVal vPrimarySO As Double, _
 ByRef vPrimaryCurrency As Double, ByRef vJenisFaktur As EnumJenisFakturPajak, ByRef vPersenPajak _
 As Double)
        Dim DTA As New dseTransaksiSOTableAdapters.tbltrans_soTableAdapter
        Dim TBL As New dseTransaksiSO.tbltrans_soDataTable
        Dim row As dseTransaksiSO.tbltrans_soRow

        Try
            If DTA.FillBy_Primary(TBL, vPrimarySO) > 0 Then
                row = TBL.Rows(0)
                vPrimaryCurrency = row.Primary_Currency
                vPersenPajak = row.Tax
                vJenisFaktur = row.Jenis_Faktur_Pajak
            Else
                vPrimaryCurrency = -1
                vPersenPajak = 0
                vJenisFaktur = EnumJenisFakturPajak.FakturPajakSederhana
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Function GetRemarksItem(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double) As String
        Dim DTA As New dseTransaksiSOTableAdapters.tbltrans_so_itemTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetRemarks(vPrimarySO, vPrimaryProduct)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetRemarks(vPrimarySO, vPrimaryProduct)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSalesOrderBalance(ByVal vprimarySO As Double, ByVal vPrimaryProduct As Double) _
   As Decimal
        Dim DTA As New dseTransaksiSOTableAdapters.laporan_sales_orderbalance_view9d_lamaTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetOrderBalance_a(vPrimaryProduct, vprimarySO).HasValue = True Then
                'MsgBox(DTA.ScalarQuery_GetOrderBalance(vprimarySO, vPrimaryProduct).Value.ToString)
                Return DTA.ScalarQuery_GetOrderBalance_a(vPrimaryProduct, vprimarySO).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSalesOrderBalance(ByVal vPrimaryProduct As Double) _
  As Decimal
        Dim DTA As New dseTransaksiSOTableAdapters.laporan_sales_orderbalance_view9d_lamaTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetOrderBalance_b(vPrimaryProduct).HasValue = True Then
                'MsgBox(DTA.ScalarQuery_GetOrderBalance(vprimarySO, vPrimaryProduct).Value.ToString)
                Return DTA.ScalarQuery_GetOrderBalance_b(vPrimaryProduct).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub IsikanTBLProduct2(ByVal vPrimarySalesOrder As Double, ByVal TBL As dseTransaksiSO.laporan_sales_orderbalance_view9dDataTable)
        Dim DTA As New dseTransaksiSOTableAdapters.laporan_sales_orderbalance_view9dTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimarySO(TBL, vPrimarySalesOrder)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLSO3(ByVal vPrimaryCustomer As Double, ByVal TBL As dseTransaksiSO.laporan_sales_orderbalance_view9dDataTable, _
  ByVal cbo As ComboBox)
        Dim DTA As New dseTransaksiSOTableAdapters.laporan_sales_orderbalance_view9dTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimaryCust(TBL, vPrimaryCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function Get_CreatedBy(ByVal vPrimarySO As Double) As String
        Dim DTA As New dseTransaksiSOTableAdapters.tbltrans_soTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCreatedBy(vPrimarySO)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetCreatedBy(vPrimarySO)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryCustomer(ByVal vPrimarySO As Double) As Double
        Dim DTA As New dseTransaksiSOTableAdapters.tbltrans_soTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryCust(vPrimarySO).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryCust(vPrimarySO).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New dseTransaksiSOTableAdapters.tbltrans_soTableAdapter

        Try
            If DTA.ScalarQuery_GetJml(vNoTransaksi).HasValue = True Then
                If DTA.ScalarQuery_GetJml(vNoTransaksi).Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vSoNo As String) As Double
        Dim DTA As New dseTransaksiSOTableAdapters.tbltrans_soTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimary(vSoNo).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimary(vSoNo).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetCustPO(ByVal vPrimarySO As Double) As String
        Dim DTA As New dseTransaksiSOTableAdapters.tbltrans_soTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCustPONo(vPrimarySO)) = False Then
                Return DTA.ScalarQuery_GetCustPONo(vPrimarySO)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoSO(ByVal vPrimarySO As Double) As String
        Dim DTA As New dseTransaksiSOTableAdapters.tbltrans_soTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetSOno(vPrimarySO)) = False Then
                Return DTA.ScalarQuery_GetSOno(vPrimarySO)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
      ByVal DGV As DataGridView, ByVal TBL As dseTransaksiSO.DataTable2DataTable, _
      ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
      Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
      Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New dseTransaksiSOTableAdapters.DataTable2TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_TanggalSampaiDari(TBL, vTanggalSampai, vTanggalDari)
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerPO_NO Then
                QueryResult = DTA.FillBy_CustPONOLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                QueryResult = DTA.FillBy_CustCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                QueryResult = DTA.FillBy_ProductCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                QueryResult = DTA.FillBy_ProductDescLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                QueryResult = DTA.FillBy_ProductSpecLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
            End If

            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.TotalSeconds > 1 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalSeconds, 2) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalMilliseconds, 2, MidpointRounding.AwayFromZero) + " miliseconds"
                End If
            Else
                tsl.Text = "No result found"
            End If
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function GetJml(ByVal Criteria As String, _
    ByVal SearchCriteria As EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New dseTransaksiSOTableAdapters.tbltrans_soTableAdapter

        Try
            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.ScalarQuery_GetJmlByNoTransaksiLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                Return DTA.ScalarQuery_GetJmlByTanggalSampaiDari(vTanggalSampai, vTanggalDari).Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                Return DTA.ScalarQuery_GetJmlByProductCodeLIKE(Criteria + "%").Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                Return DTA.ScalarQuery_GetJmlByProductDescLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                Return DTA.ScalarQuery_GetJmlByProductSpecLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                Return DTA.ScalarQuery_GetJmlByCustCodeLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                Return DTA.ScalarQuery_GetJmlByCustNameLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerPO_NO Then
                Return DTA.ScalarQuery_GetJmlByCustPONoLIKE(Criteria + "%")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlProductByIdentifier2(ByVal vIdentifier As String) As Decimal
        Dim DTA As New dseTransaksiSOTableAdapters.tbltrans_so_item_oldTableAdapter
        Dim RetVal As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJml(vIdentifier)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlProductByIdentifier3(ByVal vIdentifier As String) As Decimal
        Dim DTA As New dseTransaksiSOTableAdapters.tbltrans_so_item_oldTableAdapter
        Dim RetVal As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJml2(vIdentifier)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlProductByIdentifier(ByVal vIdentifier As String) As Double
        Dim DTA As New dseTransaksiSOTableAdapters.tbltrans_so_item_oldTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJmlByIdentifier(vIdentifier)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryProductByIdentifier(ByVal vIdentifier As String) As Double
        Dim DTA As New dseTransaksiSOTableAdapters.tbltrans_so_item_oldTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetPrimaryProduct(vIdentifier)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahDeliveryDateControl(ByVal vIdentifier As String) As Boolean
        Dim DTA As New dseTransaksiSOTableAdapters.tbltrans_so_item_oldTableAdapter
        Dim RetVal As Nullable(Of Boolean)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetDeliveryDateCont(vIdentifier)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub SetDeliveryDateControl(ByVal vDelDate As Date, ByVal vIdentifier As String)
        Dim DTA As New dseTransaksiSOTableAdapters.tbltrans_so_item_oldTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_SetDelDate(1, vDelDate, vIdentifier)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub SetDeliveryDateControl(ByVal vDelDateCont As Boolean, ByVal vDeliveryDate As Date, ByVal vIdentifier As String)
        Dim DTA As New dseTransaksiSOTableAdapters.tbltrans_so_item_oldTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vDelDateCont = True Then
                DTA.UpdateQuery_SetDelDate(1, vDeliveryDate, vIdentifier)
            Else
                DTA.UpdateQuery_SetDelDate(0, #1/1/1950#, vIdentifier)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub HapusItemDetailDelivery(ByVal vIdentifierSOItem As String)
        Dim DTA As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_so_itemdeliveryTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.DeleteQuery(vIdentifierSOItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub SimpanItemDetailDelivery(ByVal vTBL As DseTransaksiMaterialPlanning2.tbltrans_so_itemdeliveryDataTable)
        Dim DTA As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_so_itemdeliveryTableAdapter
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_so_itemdeliveryRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            For Each row In vTBL.Rows
                With row
                    DTA.InsertQuery(.primary_transaksi, .identifier, .delivery_qty, .delivery_date, .remarks, .delivery_priority, .delivery_qty_multi)
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetOrderBalanceByIdentifier(ByVal vIdentifier As String) As Decimal
        Dim DTA As New dseTransaksiSOTableAdapters.tbltrans_so_item_oldTableAdapter
        Dim RetVal As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetSOB(vIdentifier)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetOrderMultiByIdentifier(ByVal vIdentifier As String) As Decimal
        Dim DTA As New dseTransaksiSOTableAdapters.tbltrans_so_item_oldTableAdapter
        Dim RetVal As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJml2(vIdentifier)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSalesOrderBalance_Multi(ByVal vprimarySO As Double, ByVal vPrimaryProduct As Double) As Decimal
        Dim DTA As New DseTransaksiSOTableAdapters.laporan_sales_orderbalance_view9aTableAdapter
        Try
            If DTA.ScalarQuery_GetOrderBalance_Multi(vprimarySO, vPrimaryProduct).HasValue = True Then
                'MsgBox(DTA.ScalarQuery_GetOrderBalance(vprimarySO, vPrimaryProduct).Value.ToString)
                Return DTA.ScalarQuery_GetOrderBalance_Multi(vprimarySO, vPrimaryProduct).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    'tambahan aul sampai bawah
    Public Function GetToleranceUp(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double) As Double
        Dim Dta As New DseTransaksiSOTableAdapters.tbltrans_so_itemTableAdapter
        Dim NoD As Nullable(Of Decimal)
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring()
            NoD = Dta.ScalarQuery_GetToleranceUp(vPrimarySO, vPrimaryProduct)
            If NoD.HasValue Then
                Return NoD.Value
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetCustPONo(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter
        Try
            Return DTA.ScalarQuery_GetCustPONo(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub IsikanTBLSO3_Tolerance(ByVal vPrimaryCustomer As Double, ByVal TBL As dseTransaksiSO.laporan_sales_orderbalance_view9dDataTable, _
        ByVal cbo As ComboBox)
        Dim DTA As New dseTransaksiSOTableAdapters.laporan_sales_orderbalance_view9dTableAdapter

        Try
            DTA.FillBy_PrimaryCust_Tolerance(TBL, vPrimaryCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLSO3_Default(ByVal vPrimaryCustomer As Double, ByVal TBL As dseTransaksiSO.laporan_sales_orderbalance_view9dDataTable, _
        ByVal cbo As ComboBox)
        Dim DTA As New DseTransaksiSOTableAdapters.laporan_sales_orderbalance_view9dTableAdapter

        Try
            DTA.FillBy_PrimaryCust(TBL, vPrimaryCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLProduct2_Default(ByVal vPrimarySalesOrder As Double, ByVal TBL As dseTransaksiSO.laporan_sales_orderbalance_view9dDataTable)
        Dim DTA As New DseTransaksiSOTableAdapters.laporan_sales_orderbalance_view9dTableAdapter

        Try
            DTA.FillBy_PrimarySO(TBL, vPrimarySalesOrder)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLProduct2_Tolerance(ByVal vPrimarySalesOrder As Double, ByVal TBL As dseTransaksiSO.laporan_sales_orderbalance_view9dDataTable)
        Dim DTA As New DseTransaksiSOTableAdapters.laporan_sales_orderbalance_view9dTableAdapter

        Try
            DTA.FillBy_PrimarySO_Tolerance(TBL, vPrimarySalesOrder)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub UpdateQtySpare(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double, ByVal vQtySpare As Decimal)
        Dim DTA As New dseTransaksiSOTableAdapters.tbltrans_so_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_SetSpareQty(vQtySpare, vPrimarySO, vPrimaryProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetQtyMulti(ByVal vIdentifierSO As String) As Decimal
        Dim DTA As New dseTransaksiSOTableAdapters.tbltrans_so_item_oldTableAdapter
        Dim retVal As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            retVal = DTA.ScalarQuery_GetJml2(vIdentifierSO)
            If retVal.HasValue = True Then
                Return retVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetQty(ByVal vIdentifierSO As String) As Decimal
        Dim DTA As New dseTransaksiSOTableAdapters.tbltrans_so_item_oldTableAdapter
        Dim retVal As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            retVal = DTA.ScalarQuery_GetJml(vIdentifierSO)
            If retVal.HasValue = True Then
                Return retVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class
