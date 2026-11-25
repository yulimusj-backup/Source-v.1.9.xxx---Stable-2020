Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTrans_MaterialPlanningDashboard_SF_AddDO
    Dim FromButton As Boolean = False
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjCustomer As New clsMasterCustomer
    Dim rowDGV As DataGridViewRow
    Dim objSO As New ClsTransaksiSalesOrder
    Dim ObjInvoiceBalance As New ClsTransaksiSalesBalance
    Dim TBLSalesBalance As New DseTransaksiSalesBalance.balance_invoice_viewDataTable
    Dim TBLSRBalance As New DseTransaksiSalesBalance.balance_sr_viewDataTable
    Dim FrmSales As New FrmTransaksiJournal_Sales
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim TBLItemSOB As New DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable
    Dim rowItemSOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
    Dim TBLDO As New DseTransaksiShippingForecast.DODataTable
    Dim rowDO, rowDummyDO As DseTransaksiShippingForecast.DORow
    Dim ObjDO As New ClsTransaksiDO
    Dim objNumber As New Core_Module.ClsMasterTransactionNumbering
    Dim TblDummyDO As New DseTransaksiShippingForecast.DODataTable
    Dim Buffer As String = ""
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjStorage As New Core_Module.clsMasterStorageArea
    Dim StopProses As Boolean = False
    Dim TBLSOBDummy1 As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable
    Dim OnProgressDO As Boolean = False
    Dim PrinterAdd As String
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjOtorisasi As New Core_Module.ClsMasterTipeOtorisasi

    Public vFrmRef As FrmTrans_MaterialPlanningDashboard
    Public vdgvRef As DataGridView
    Public vUserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public vDObySO As Boolean = False

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        Try
            Me.FromButton = True
            Me.TabControl1.SelectedIndex -= 1
            If Me.TabControl1.SelectedIndex = 0 Then
                Me.btnBack.Enabled = False
            End If
            Me.btnNext.Enabled = True
            Me.FromButton = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            If Me.TabControl1.SelectedIndex = 0 Then
                Me.btnBack.Enabled = False
                Me.btnOK.Enabled = False
            ElseIf Me.TabControl1.SelectedIndex = Me.TabControl1.TabCount - 1 Then
                Me.btnNext.Enabled = False
                Me.btnOK.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_Selecting(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TabControlCancelEventArgs) Handles TabControl1.Selecting
        If FromButton = False Then
            e.Cancel = True
        End If
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Function ApakahAdaOverrideCustomer() As Boolean
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvException.Rows
                If rowDGV.Cells(Me.ColExc_CreateDO.Index).Value = Boolean.TrueString Then
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub ResetObjOtorisasi()
        Try
            With Me.ObjOtorisasi
                .NoDocumentHist = ""
                .Primary = -1
                .PrimaryApprovedBy = -1
                .PrimaryDokumenAcuan = -1
                .PrimaryRequestBy = -1
                .PrimaryTipeApproval = ClsMasterTipeOtorisasi.EnumTipeOtorisasi.SF_Override_DO
                .Reason = ""
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        Dim FrmO As New Core_Module.FrmMasterUser_Otorisasi

        Try
            Me.ResetObjOtorisasi()
            Me.FromButton = True
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.vUserCode, "FrmTrans_MaterialPlanningDashboard_SF_OverrideDO") = False And _
            Me.ApakahAdaOverrideCustomer = True Then
                If Me.ObjUser.GetJmlByNamaHakAkses("FrmTrans_MaterialPlanningDashboard_SF_OverrideDO") > 0 Then
                    FrmO.AccessName = "FrmTrans_MaterialPlanningDashboard_SF_OverrideDO"
                    If FrmO.ShowDialog <> Windows.Forms.DialogResult.OK Then
                        MsgBox("You don't have the access to create delivery order for customers with overdue invoice or over credit limit account receivable.", _
                        MsgBoxStyle.Information, "Information")
                        Exit Sub
                    Else
                        With Me.ObjOtorisasi
                            .NoDocumentHist = ""
                            .Primary = 1
                            .PrimaryApprovedBy = Me.ObjUser.GetPrimary(FrmO.cboUserCode.Text)
                            .PrimaryDokumenAcuan = -1
                            .PrimaryRequestBy = Me.ObjUser.GetPrimary(Me.vUserCode)
                            .PrimaryTipeApproval = ClsMasterTipeOtorisasi.EnumTipeOtorisasi.SF_Override_DO
                            .Reason = FrmO.TxtReason.Text
                        End With
                    End If
                Else
                    MsgBox("Please specify user that can create delivery order for customers with overdue invoice or over credit limit account receivable.")
                    Exit Sub
                End If
            End If

            Me.TabControl1.SelectedIndex += 1
            If Me.TabControl1.SelectedIndex = Me.TabControl1.TabCount - 1 Then
                Me.btnNext.Enabled = False
            End If
            Me.btnBack.Enabled = True
            Me.FromButton = False
            If Me.TabControl1.SelectedIndex = 1 Then
                Me.IsikanDataPersiapan()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmO = Nothing
        End Try
    End Sub

    Private Sub IsikanDataPersiapan()
        Dim row1 As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim vTBLRef As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable
        Dim ListExc As New List(Of Double)
        Dim TBLSO As New DseTransaksiShippingForecast.tbltrans_soDataTable
        Dim DTASO As New DseTransaksiShippingForecastTableAdapters.tbltrans_soTableAdapter
        Dim rowSO As DseTransaksiShippingForecast.tbltrans_soRow
        Dim PrimarySOMax, PrimarySOMin As Double
        Dim ListPrimarySO As New List(Of Double)
        Dim PrimarySO As Double
        Dim DummyUrutan As Double = 0
        Dim AcuanPrimaryCurr, AcuanTaxPercentage As Double

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTASO.Connection.ConnectionString = Core_Module.GetConstring

            'isikan list untuk customer yang tidak dibuatkan do
            ListExc.Clear()
            'Me.dgvException.RowCount
            For Each rowDGV In Me.dgvException.Rows
                If rowDGV.Cells(Me.ColExc_CreateDO.Index).Value = Boolean.FalseString Then
                    If ListExc.Contains(rowDGV.Cells(Me.ColExc_PrimaryCust.Index).Value) = False Then
                        ListExc.Add(rowDGV.Cells(Me.ColExc_PrimaryCust.Index).Value)
                    End If
                End If
            Next

            'isikan data item delivery yang akan dibuatkan do
            PrimarySOMax = 0
            PrimarySOMin = Double.MaxValue
            vTBLRef = Me.vFrmRef.vTBLItemDelivery
            Me.TBLItemSOB.Rows.Clear()
            For Each Me.rowDGV In Me.vdgvRef.Rows
                If rowDGV.Cells(Me.vFrmRef.ColSF_CreateDO.Index).Value = Boolean.TrueString And _
                rowDGV.Cells(Me.vFrmRef.ColSF_TypeInt.Index).Value = ClsTransaksiShippingForecast_Item.EnumTipeItemSF.SOB Then
                    vTBLRef.DefaultView.RowFilter = "primary = " + rowDGV.Cells(vFrmRef.ColSF_Primary.Index).Value.ToString
                    If vTBLRef.DefaultView.Count = 0 Then
                        MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    Else
                        'tambahkan item shipment untuk primary cust yang tidak masuk exception
                        row1 = vTBLRef.DefaultView(0).Row
                        If ListExc.Contains(row1.primary_cust) = False Then
                            Me.rowItemSOB = Me.TBLItemSOB.Newtbltrans_shipping_forecast_itemdeliveryRow
                            For ij As Integer = 0 To Me.TBLItemSOB.Columns.Count - 1
                                Me.rowItemSOB.Item(ij) = row1.Item(ij)
                            Next
                            Me.TBLItemSOB.Rows.Add(Me.rowItemSOB)
                            If PrimarySOMax < Me.rowItemSOB.primary_so Then
                                PrimarySOMax = Me.rowItemSOB.primary_so
                            End If
                            If PrimarySOMin > Me.rowItemSOB.primary_so Then
                                PrimarySOMin = Me.rowItemSOB.primary_so
                            End If
                        End If
                    End If
                End If
            Next

            'isikan primary customer, primary currency, use tax dan tax percentage untuk item shipping forecast
            If PrimarySOMin = Double.MaxValue Then
                Exit Sub
            End If
            DTASO.Fill(TBLSO, PrimarySOMax, PrimarySOMin)
            For Each Me.rowItemSOB In Me.TBLItemSOB.Rows
                If ListPrimarySO.Contains(Me.rowItemSOB.primary_so) = False Then
                    ListPrimarySO.Add(Me.rowItemSOB.primary_so)
                    PrimarySO = Me.rowItemSOB.primary_so
                    TBLSO.DefaultView.RowFilter = "primary = " + Me.rowItemSOB.primary_so.ToString
                    If TBLSO.DefaultView.Count = 0 Then
                        MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    Else
                        rowSO = TBLSO.DefaultView(0).Row
                        For Each row1 In Me.TBLItemSOB.Rows
                            If row1.primary_so = PrimarySO Then
                                row1.create_do_primary_currency = rowSO.Primary_Currency
                                row1.create_do_tax_percentage = rowSO.Tax
                                If rowSO.Use_Tax = True Then
                                    row1.create_do_use_tax = 1
                                Else
                                    row1.create_do_use_tax = 0
                                End If
                                'row1.create_do_use_tax = rowSO.Use_Tax
                            End If
                        Next
                    End If
                End If
            Next

            'kelompokkan item delivery yang dicentang berdasarkan informasi: primary cust, primary currency, jenis faktur pajak yang digunakan, % pajak yang digunakan
            Me.TBLDO.Rows.Clear()
            Me.dgvDONumbering.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvDONumbering.Rows.Clear()
            For Each Me.rowItemSOB In Me.TBLItemSOB.Rows
                If Me.vDObySO = True Then
                    Me.TBLDO.DefaultView.RowFilter = "[Primary SO] = '" & Me.rowItemSOB.primary_so & "'"
                    If Me.TBLDO.DefaultView.Count = 0 Then
                        Me.rowDO = Me.TBLDO.NewDORow
                        Me.rowDO.Customer = Me.rowItemSOB.customer_name
                        Me.rowDO.DO_Number = ""
                        Me.rowDO.Primary_Curr = Me.rowItemSOB.create_do_primary_currency
                        Me.rowDO.Primary_Cust = Me.rowItemSOB.primary_cust
                        Me.rowDO.Use_Tax = Me.rowItemSOB.create_do_use_tax
                        Me.rowDO.Tax_Percentage = Me.rowItemSOB.create_do_tax_percentage
                        Me.rowDO.Remarks = ""
                        Me.rowDO.Primary_SO = Me.rowItemSOB.primary_so
                        Me.rowDO.SO_Number = Me.rowItemSOB.so_number
                        Me.TBLDO.Rows.Add(Me.rowDO)

                        Me.rowDO = Me.TblDummyDO.NewDORow
                        Me.rowDO.Customer = Me.rowItemSOB.customer_name
                        Me.rowDO.DO_Number = ""
                        Me.rowDO.Primary_Curr = Me.rowItemSOB.create_do_primary_currency
                        Me.rowDO.Primary_Cust = Me.rowItemSOB.primary_cust
                        Me.rowDO.Use_Tax = Me.rowItemSOB.create_do_use_tax
                        Me.rowDO.Tax_Percentage = Me.rowItemSOB.create_do_tax_percentage
                        Me.rowDO.Remarks = ""
                        Me.rowDO.Primary_SO = Me.rowItemSOB.primary_so
                        Me.rowDO.SO_Number = Me.rowItemSOB.so_number
                        Me.TblDummyDO.Rows.Add(Me.rowDO)
                    End If
                Else
                    Me.TBLDO.DefaultView.RowFilter = "[Primary Cust] = " + Me.rowItemSOB.primary_cust.ToString + " AND [Primary Curr] = " + Me.rowItemSOB.create_do_primary_currency.ToString + _
                    " AND [Use Tax] = " + Me.rowItemSOB.create_do_use_tax.ToString + " AND [Tax Percentage] = " + Me.rowItemSOB.create_do_tax_percentage.ToString
                    If Me.TBLDO.DefaultView.Count = 0 Then
                        Me.rowDO = Me.TBLDO.NewDORow
                        Me.rowDO.Customer = Me.rowItemSOB.customer_name
                        Me.rowDO.DO_Number = ""
                        Me.rowDO.Primary_Curr = Me.rowItemSOB.create_do_primary_currency
                        Me.rowDO.Primary_Cust = Me.rowItemSOB.primary_cust
                        Me.rowDO.Use_Tax = Me.rowItemSOB.create_do_use_tax
                        Me.rowDO.Tax_Percentage = Me.rowItemSOB.create_do_tax_percentage
                        Me.rowDO.Remarks = ""
                        Me.rowDO.Primary_SO = Me.rowItemSOB.primary_so
                        Me.rowDO.SO_Number = Me.rowItemSOB.so_number
                        Me.TBLDO.Rows.Add(Me.rowDO)
                    End If

                    'update data ditabel dummy
                    Me.TblDummyDO.DefaultView.RowFilter = "[Primary Cust] = " + Me.rowItemSOB.primary_cust.ToString + " AND [Primary Curr] = " + Me.rowItemSOB.create_do_primary_currency.ToString + _
                    " AND [Use Tax] = " + Me.rowItemSOB.create_do_use_tax.ToString + " AND [Tax Percentage] = " + Me.rowItemSOB.create_do_tax_percentage.ToString
                    If Me.TblDummyDO.DefaultView.Count = 0 Then
                        Me.rowDO = Me.TblDummyDO.NewDORow
                        Me.rowDO.Customer = Me.rowItemSOB.customer_name
                        Me.rowDO.DO_Number = ""
                        Me.rowDO.Primary_Curr = Me.rowItemSOB.create_do_primary_currency
                        Me.rowDO.Primary_Cust = Me.rowItemSOB.primary_cust
                        Me.rowDO.Use_Tax = Me.rowItemSOB.create_do_use_tax
                        Me.rowDO.Tax_Percentage = Me.rowItemSOB.create_do_tax_percentage
                        Me.rowDO.Remarks = ""
                        Me.rowDO.Primary_SO = Me.rowItemSOB.primary_so
                        Me.rowDO.SO_Number = Me.rowItemSOB.so_number
                        Me.TblDummyDO.Rows.Add(Me.rowDO)
                    End If
                End If
            Next

            'salin ke dalam dgv
            Me.TBLDO.DefaultView.Sort = "Customer ASC"
            Me.TBLDO.DefaultView.RowFilter = ""
            For Each DV As Data.DataRowView In Me.TBLDO.DefaultView
                rowDO = DV.Row

                If Me.vDObySO = False Then
                    'update remarks
                    Me.TblDummyDO.DefaultView.RowFilter = "[Primary Cust] = " + rowDO.Primary_Cust.ToString + " AND [Primary Curr] = " + rowDO.Primary_Curr.ToString + _
                    " AND [Use Tax] = " + rowDO.Use_Tax.ToString + " AND [Tax Percentage] = " + rowDO.Tax_Percentage.ToString
                    If Me.TblDummyDO.DefaultView.Count > 0 Then
                        rowDummyDO = Me.TblDummyDO.DefaultView(0).Row
                        If rowDummyDO.Remarks <> "" Then
                            rowDO.Remarks = rowDummyDO.Remarks
                        End If
                    End If
                End If

                'tambahkan rowdgv
                Me.dgvDONumbering.Rows.Add()
                rowDGV = Me.dgvDONumbering.Rows(Me.dgvDONumbering.Rows.Count - 1)
                With rowDGV
                    .Cells(Me.ColDONo_Customer.Index).Value = Me.rowDO.Customer
                    .Cells(Me.ColDONo_DoNumber.Index).Value = ""
                    .Cells(Me.ColDONo_No.Index).Value = Me.dgvDONumbering.Rows.Count
                    .Cells(Me.ColDONo_Remarks.Index).Value = rowDO.Remarks
                    .Cells(Me.ColDONo_PrimaryCur.Index).Value = rowDO.Primary_Curr
                    .Cells(Me.ColDONo_CurrencyCode.Index).Value = Me.ObjCur.GetCurrencyCode(rowDO.Primary_Curr)
                    .Cells(Me.ColDONo_TarPercentage.Index).Value = Me.ObjCD.DigitGrouping(rowDO.Tax_Percentage, 2)
                    If rowDO.Use_Tax = 1 Then
                        .Cells(Me.ColDONo_UseTax.Index).Value = Boolean.TrueString
                    Else
                        .Cells(Me.ColDONo_UseTax.Index).Value = Boolean.FalseString
                    End If
                    .Cells(Me.ColDONo_PrimaryCust.Index).Value = rowDO.Primary_Cust
                    .Cells(Me.ColDONo_SoNumber.Index).Value = rowDO.SO_Number
                    .Cells(Me.ColDONo_PrimarySO.Index).Value = rowDO.Primary_SO
                End With
            Next

            'numbering
            If Me.objNumber.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO) = True Then
                'ambil auto number
                DummyUrutan = Me.objNumber.GetUrutan(Me.ObjDO.GetNoBaru1(Me.dtpDate.Value), ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO, Me.dtpDate.Value)
                For Each rowDGV In Me.dgvDONumbering.Rows
                    rowDGV.Cells(Me.ColDONo_DoNumber.Index).Value = Me.objNumber.GetNomorTransaksiBaru1(Me.dtpDate.Value, DummyUrutan, ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO)
                    DummyUrutan += 1
                Next
            Else
                'no do berdasarkan tabel dummy
                For Each rowDGV In Me.dgvDONumbering.Rows
                    If rowDGV.Cells(Me.ColDONo_UseTax.Index).Value = Boolean.TrueString Then
                        DummyUrutan = 1
                    Else
                        DummyUrutan = 0
                    End If

                    If Me.vDObySO = False Then
                        Me.TblDummyDO.DefaultView.RowFilter = "[Primary Cust] = " + rowDGV.Cells(Me.ColDONo_PrimaryCust.Index).Value.ToString + _
                        " AND [Primary Curr] = " + rowDGV.Cells(Me.ColDONo_PrimaryCur.Index).Value.ToString + _
                        " AND [Use Tax] = " + DummyUrutan.ToString + " AND [Tax Percentage] = " + CDec(rowDGV.Cells(Me.ColDONo_TarPercentage.Index).Value).ToString
                    Else
                        Me.TblDummyDO.DefaultView.RowFilter = "[Primary SO] = " & rowDGV.Cells(Me.ColDONo_PrimarySO.Index).Value 'ganti ke primary so
                    End If

                    If Me.TblDummyDO.DefaultView.Count > 0 Then
                        rowDummyDO = Me.TblDummyDO.DefaultView(0).Row
                        If rowDummyDO.DO_Number <> "" Then
                            rowDGV.Cells(Me.ColDONo_DoNumber.Index).Value = rowDummyDO.DO_Number
                        End If
                    End If
                Next
            End If

            'tentukan visibility untuk kolom currency code, use tax, dan tax percentage
            If Me.dgvDONumbering.Rows.Count > 1 Then
                AcuanPrimaryCurr = Me.dgvDONumbering.Item(Me.ColDONo_PrimaryCur.Index, 0).Value
                For Each rowDGV In Me.dgvDONumbering.Rows
                    If rowDGV.Index <> 0 Then
                        If rowDGV.Cells(Me.ColDONo_PrimaryCur.Index).Value <> AcuanPrimaryCurr Then
                            AcuanPrimaryCurr = -2
                            Exit For
                        End If
                    End If
                Next
                If AcuanPrimaryCurr = -2 Then
                    Me.ColDONo_CurrencyCode.Visible = True
                Else
                    Me.ColDONo_CurrencyCode.Visible = False
                End If

                AcuanTaxPercentage = Me.dgvDONumbering.Item(Me.ColDONo_TarPercentage.Index, 0).Value
                For Each rowDGV In Me.dgvDONumbering.Rows
                    If rowDGV.Index <> 0 Then
                        If rowDGV.Cells(Me.ColDONo_TarPercentage.Index).Value <> AcuanTaxPercentage Then
                            AcuanTaxPercentage = -2
                            Exit For
                        End If
                    End If
                Next
                If AcuanTaxPercentage = -2 Then
                    Me.ColDONo_TarPercentage.Visible = True
                    Me.ColDONo_UseTax.Visible = True
                Else
                    Me.ColDONo_TarPercentage.Visible = False
                    Me.ColDONo_UseTax.Visible = False
                End If
            Else
                Me.ColDONo_UseTax.Visible = False
                Me.ColDONo_CurrencyCode.Visible = False
                Me.ColDONo_TarPercentage.Visible = False
            End If
            Me.dgvDONumbering.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListExc = Nothing
            TBLSO.Dispose()
            DTASO.Dispose()
            ListPrimarySO = Nothing
        End Try
    End Sub

    Private Sub ApakahAdaCustomerBermasalah()
        Dim ListPrimary As New List(Of Double)
        Dim CreditLimit As Decimal = 0
        Dim PrimaryCust As Double
        Dim OverDue As Boolean = False
        Dim OverLimit As Boolean = False
        Dim Tolerance As Decimal = 0
        Dim Dummy As String = ""

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.dgvException.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvException.Rows.Clear()

            For Each rowDGV In Me.vdgvRef.Rows
                If rowDGV.Cells(Me.vFrmRef.ColSF_CreateDO.Index).Value = Boolean.TrueString And _
                rowDGV.Cells(Me.vFrmRef.ColSF_TypeInt.Index).Value = ClsTransaksiShippingForecast_Item.EnumTipeItemSF.SOB Then
                    If ListPrimary.Contains(rowDGV.Cells(Me.vFrmRef.ColSF_PrimaryCust.Index).Value) = False Then
                        ListPrimary.Add(rowDGV.Cells(Me.vFrmRef.ColSF_PrimaryCust.Index).Value)
                    End If
                End If
            Next

            'ada 2 kelompok customer yang bermasalah:
            For Each PrimaryCust In ListPrimary
                OverDue = False
                OverLimit = False

                '1. yang punya ar diatas credit limit
                CreditLimit = Me.ObjCustomer.GetCreditLimit(PrimaryCust)
                Dummy = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, ClsTransaksiSetting.EnumNamaParameter.Reminder_AccountReceivable_ExceedsCreditLimit_Tolerance)
                If Dummy = "" Then
                    Tolerance = 0
                Else
                    Tolerance = Dummy
                End If
                If CreditLimit <> 0 Then
                    CreditLimit += Tolerance
                    If CreditLimit < Me.objSO.GetAR(PrimaryCust) Then
                        OverLimit = True
                    End If
                End If

                '2. yang punya ar overdue
                If Me.objSO.ApakahAdaAROverdue(PrimaryCust, Me.dtpOverdue.Value.Date) = True Then
                    OverDue = True
                End If

                If OverDue = True Or OverLimit = True Then
                    Me.dgvException.Rows.Add()
                    rowDGV = Me.dgvException.Rows(Me.dgvException.Rows.Count - 1)
                    With rowDGV
                        .Cells(Me.ColExc_AboveCreditLimit.Index).Value = OverLimit
                        .Cells(Me.ColExc_CreateDO.Index).Value = Boolean.FalseString
                        .Cells(Me.ColExc_Customer.Index).Value = Me.ObjCustomer.GetCustName(PrimaryCust)
                        .Cells(Me.ColExc_No.Index).Value = Me.dgvException.Rows.Count
                        .Cells(Me.ColExc_OverdueInvoice.Index).Value = OverDue
                        .Cells(Me.ColExc_PrimaryCust.Index).Value = PrimaryCust
                    End With
                End If
            Next

            Me.dgvException.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimary = Nothing
        End Try
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_SF_AddDO_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Dummy As String = ""
        Dim Tolerance As Decimal = 0

        Try
            Dummy = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, ClsTransaksiSetting.EnumNamaParameter.Reminder_AccountReceivable_Overdue_Tolerance)
            If Dummy <> "" Then
                Tolerance = Dummy
            Else
                Tolerance = 0
            End If
            Me.dtpOverdue.Value = Today - TimeSpan.FromDays(Tolerance)

            Me.dtpDate.Value = Today
            Me.dgvException.AllowUserToAddRows = False
            Me.ApakahAdaCustomerBermasalah()
            If Me.dgvException.Rows.Count = 0 Then
                Me.btnNext.PerformClick()
            Else
                Me.dgvException.ClearSelection()
                Me.dgvException.Rows(0).Selected = True
            End If

            If Me.ObjUser.ApakahMemilikiHakAkses(Me.vUserCode, "FrmTrans_MaterialPlanningDashboard_SF_OverrideDO") = False Then
                Me.dtpOverdue.Enabled = False
            Else
                Me.dtpOverdue.Enabled = True
            End If

            If Me.vDObySO = False Then
                Me.ColDONo_SoNumber.Visible = False
            Else
                Me.ColDONo_SoNumber.Visible = True
            End If

            Me.dgvDONumbering.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvException.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvSalesBalance.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvException_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvException.SelectionChanged
        Dim DefPrec As Integer = 0
        Dim TotalAR, TotalAROverDue As Decimal
        Dim CurCode As String = ""

        Try
            If Me.dgvException.SelectedCells.Count = 0 Then
                Me.dgvSalesBalance.Rows.Clear()
                Me.LblCreditLimit.Text = ": -"
                Me.LblTotalARAbove.Text = ": -"
                Me.LblTotalOverDue.Text = ": -"
            Else
                Me.LoadDetailCustomer(Me.dgvException.Item(Me.ColExc_PrimaryCust.Index, Me.dgvException.SelectedCells(0).RowIndex).Value)
                'isikan data credit limit
                DefPrec = Me.ObjCur.GetDecimal_DefaultCurrency
                CurCode = Me.ObjCur.GetCurrencyCode(1)
                Me.LblCreditLimit.Text = ": " + CurCode + " " + _
                ObjCD.DigitGrouping(Me.ObjCustomer.GetCreditLimit(Me.dgvException.Item(Me.ColExc_PrimaryCust.Index, Me.dgvException.SelectedCells(0).RowIndex).Value), DefPrec)
                TotalAR = 0
                TotalAROverDue = 0
                For Each rowDGV In Me.dgvSalesBalance.Rows
                    TotalAR += CDec(rowDGV.Cells(Me.ColBalanceSales.Index).Value)
                    If rowDGV.DefaultCellStyle.ForeColor = Color.Red Then
                        TotalAROverDue += CDec(rowDGV.Cells(Me.ColBalanceSales.Index).Value)
                    End If
                Next
                Me.LblTotalARAbove.Text = ": " + CurCode + " " + Me.ObjCD.DigitGrouping(TotalAR, DefPrec)
                Me.LblTotalOverDue.Text = ": " + CurCode + " " + Me.ObjCD.DigitGrouping(TotalAROverDue, DefPrec)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvException_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvException.CellBeginEdit
        Try
            If Me.ColExc_CreateDO.Index <> e.ColumnIndex Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvException_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvException.CellEndEdit
        Try
            'tidak banyak coding karena memang tidak ada batasan untuk edit kolom created do
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbCek_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCek.LinkClicked
        Try
            For Each Me.rowDGV In Me.dgvException.Rows
                rowDGV.Cells(Me.ColExc_CreateDO.Index).Value = Boolean.TrueString
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUncek_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUncek.LinkClicked
        Try
            For Each Me.rowDGV In Me.dgvException.Rows
                rowDGV.Cells(Me.ColExc_CreateDO.Index).Value = Boolean.FalseString
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDetailCustomer(ByVal vPrimaryCustomer As Double)
        Dim i As Double = 0
        Dim AdaForeignCurrency As Boolean = False

        Try
            Me.ObjInvoiceBalance.IsiDGVBalanceTanpaJurnal(Me.dgvSalesBalance, Me.TBLSalesBalance, Me.FrmSales, Me.ObjCD, _
            vPrimaryCustomer, Me.ObjCur.GetDecimal_DefaultCurrency, Me.TBLSRBalance)

            'buang yg belum sampai due date
            Me.dgvSalesBalance.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvSalesBalance.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            While i <= Me.dgvSalesBalance.Rows.Count - 1
                rowDGV = Me.dgvSalesBalance.Rows(i)
                If CDate(rowDGV.Cells(Me.ColDueDateInvoice.Index).Value).Date < Me.dtpOverdue.Value.Date Then
                    rowDGV.DefaultCellStyle.ForeColor = Color.Red
                End If
                i += 1
            End While

            For Each rowDGV In Me.dgvSalesBalance.Rows
                If rowDGV.Cells(Me.ColPrimaryCurrencyBalanceSales.Index).Value <> Core_Module.PrimaryBaseCurrency Then
                    AdaForeignCurrency = True
                    Exit For
                End If
            Next

            If AdaForeignCurrency = True Then
                Me.ColExhangeRateBalance.Visible = True
                Me.ColBalanceSales.Visible = True
            Else
                Me.ColExhangeRateBalance.Visible = False
                Me.ColBalanceSales.Visible = False
            End If
            Me.dgvSalesBalance.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvException_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvException.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvException, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpDate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.Enter
        Try
            Me.Buffer = Format(Me.dtpDate.Value, "MMM yyyy")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpDate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.Leave
        Try
            If Me.Buffer <> Format(Me.dtpDate.Value, "MMM yyyy") Then
                If Me.objNumber.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO) = True Then
                    Me.BackgroundWorker1.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Dim dummyurutan As Double

        Try
            'ambil auto number
            dummyurutan = Me.objNumber.GetUrutan(Me.ObjDO.GetNoBaru1(Me.dtpDate.Value), ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO, Me.dtpDate.Value)
            Me.dgvDONumbering.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rowDGV In Me.dgvDONumbering.Rows
                rowDGV.Cells(Me.ColDONo_DoNumber.Index).Value = Me.objNumber.GetNomorTransaksiBaru1(Me.dtpDate.Value, dummyurutan, ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO)
                DummyUrutan += 1
            Next
            Me.dgvDONumbering.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvDONumbering_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvDONumbering.CellBeginEdit
        Try
            'yang boleh diedit adalah:
            '1. remarks
            '2. do number bila tidak autonumber
            If e.ColumnIndex = Me.ColDONo_Remarks.Index Then
                Me.Buffer = Me.dgvDONumbering.Item(e.ColumnIndex, e.RowIndex).Value
            ElseIf e.ColumnIndex = Me.ColDONo_DoNumber.Index And Me.objNumber.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO) = False Then
                Me.Buffer = Me.dgvDONumbering.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvDONumbering_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDONumbering.CellEndEdit
        Dim DummyInt As Integer = 0

        Try
            If IsNothing(Me.dgvDONumbering.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                If e.ColumnIndex = Me.ColDONo_Remarks.Index Then
                    Me.dgvDONumbering.Item(e.ColumnIndex, e.RowIndex).Value = ""
                End If
            End If

            If e.ColumnIndex = Me.ColDONo_Remarks.Index Then
                'trim kanan kiri
                Me.dgvDONumbering.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvDONumbering.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim()
                'maksimal 250 karakter
                If Me.dgvDONumbering.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                    Me.dgvDONumbering.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvDONumbering.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                End If
                'update ke tabel dummy
                rowDGV = Me.dgvDONumbering.Rows(e.RowIndex)
                If rowDGV.Cells(Me.ColDONo_UseTax.Index).Value = Boolean.TrueString Then
                    DummyInt = 1
                Else
                    DummyInt = 0
                End If
                Me.TblDummyDO.DefaultView.RowFilter = "[Primary Cust] = " + rowDGV.Cells(Me.ColDONo_PrimaryCust.Index).Value.ToString + _
                " AND [Primary Curr] = " + rowDGV.Cells(Me.ColDONo_PrimaryCur.Index).Value.ToString + _
                " AND [Use Tax] = " + DummyInt.ToString + " AND [Tax Percentage] = " + CDec(rowDGV.Cells(Me.ColDONo_TarPercentage.Index).Value).ToString
                If Me.TblDummyDO.DefaultView.Count > 0 Then
                    rowDummyDO = Me.TblDummyDO.DefaultView(0).Row
                    rowDummyDO.Remarks = rowDGV.Cells(Me.ColDONo_Remarks.Index).Value
                Else
                    MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    Me.dgvDONumbering.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If
            ElseIf e.ColumnIndex = Me.ColDONo_DoNumber.Index Then
                'trim kanan kiri
                Me.dgvDONumbering.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvDONumbering.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim()
                'maksimal 250 karakter
                If Me.dgvDONumbering.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                    Me.dgvDONumbering.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvDONumbering.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                End If
                'upper string
                Me.dgvDONumbering.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvDONumbering.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                'tidak boleh sama dengan nomor lain
                If Me.ApakahSama_DONumber(Me.dgvDONumbering.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex) = True Then
                    MsgBox("DO number duplication, please change do number.", MsgBoxStyle.Information, "Information")
                    Me.dgvDONumbering.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If
                'tidak boleh sudah digunakan dalam database
                If Me.ObjDO.ApakahSudahAda(Me.dgvDONumbering.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    MsgBox("DO number duplication, please change do number.", MsgBoxStyle.Information, "Information")
                    Me.dgvDONumbering.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If
                'update ke tabel dummy
                'update ke tabel dummy
                rowDGV = Me.dgvDONumbering.Rows(e.RowIndex)
                If rowDGV.Cells(Me.ColDONo_UseTax.Index).Value = Boolean.TrueString Then
                    DummyInt = 1
                Else
                    DummyInt = 0
                End If
                Me.TblDummyDO.DefaultView.RowFilter = "[Primary Cust] = " + rowDGV.Cells(Me.ColDONo_PrimaryCust.Index).Value.ToString + _
                " AND [Primary Curr] = " + rowDGV.Cells(Me.ColDONo_PrimaryCur.Index).Value.ToString + _
                " AND [Use Tax] = " + DummyInt.ToString + " AND [Tax Percentage] = " + CDec(rowDGV.Cells(Me.ColDONo_TarPercentage.Index).Value).ToString
                If Me.TblDummyDO.DefaultView.Count > 0 Then
                    rowDummyDO = Me.TblDummyDO.DefaultView(0).Row
                    rowDummyDO.DO_Number = rowDGV.Cells(Me.ColDONo_DoNumber.Index).Value
                Else
                    MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    Me.dgvDONumbering.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahSama_DONumber(ByVal vDONumber As String, ByVal vIndexAcuan As Double) As Boolean
        Try
            For Each rowDGV In Me.dgvDONumbering.Rows
                If rowDGV.Index <> vIndexAcuan Then
                    If rowDGV.Cells(Me.ColDONo_DoNumber.Index).Value = vDONumber Then
                        Return True
                    End If
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetPrimaryMax(ByVal vTBLSOBDummy As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable) As Double
        Dim rowSOBDummy As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim RetVal As Double = 0

        Try
            For Each rowSOBDummy In vTBLSOBDummy.Rows
                If RetVal < rowSOBDummy.primary Then
                    RetVal += 1
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim DTA_ProductQty As New DseTransaksiShippingForecastTableAdapters.tblmaster_productqty1TableAdapter
        Dim TBL_ProductQty As New DseTransaksiShippingForecast.tblmaster_productqty1DataTable
        'Dim TBL_ProductQty_Booked As New DseTransaksiShippingForecast.tblmaster_productqty1DataTable
        Dim row_ProductQty As DseTransaksiShippingForecast.tblmaster_productqty1Row
        Dim DTAPendingDO As New DseTransaksiShippingForecastTableAdapters.tbltrans_do_item1TableAdapter
        Dim TBLPendingDO As New DseTransaksiShippingForecast.tbltrans_do_item1DataTable
        Dim rowPendingDO As DseTransaksiShippingForecast.tbltrans_do_item1Row
        Dim DTABooking As New DseTransaksiShippingForecastTableAdapters.tbltrans_booking_invitemTableAdapter
        Dim TBLBooking As New DseTransaksiShippingForecast.tbltrans_booking_invitemDataTable
        Dim IndexDummy As Integer = 0
        Dim DV, DV1 As Data.DataRowView
        Dim ExpiredDateStr As String = ""
        Dim rowBooking As DseTransaksiShippingForecast.tbltrans_booking_invitemRow
        Dim TBLSOBDummy As New DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable
        Dim JmlBookingDummy As Decimal = 0

        Try
            'Cara kerjanya:
            '1. kumpulkan item stock on hand
            '2. kurangkan item stock dengan pending do
            '3. kumpulkan item sob + alokasikan sesuai stock yg ada, untuk setiap item sob: (a) ambil datanya dari tabel qty booked bila ada, sejumlah qty booked, (b) sisa qty available - qty booked, 
            'diambil dari tabel item stock biasa
            '4. berdasarkan row dgv do dan item sob yang sudah dialokasikan, buat do-nya

            DTA_ProductQty.Connection.ConnectionString = Core_Module.GetConstring
            DTAPendingDO.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
            DTABooking.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString

            'cek kebenaran data:
            '1. minimal 1 baris do
            If Me.dgvDONumbering.Rows.Count = 0 Then
                MsgBox("There are no item available to create do. Process cancelled.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            '2. jika tidak auto number, maka tiap baris harus ada isinya untuk kolom do number
            If Me.objNumber.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO) = False Then
                For Each Me.rowDGV In Me.dgvDONumbering.Rows
                    If rowDGV.Cells(Me.ColDONo_DoNumber.Index).Value.ToString.Trim = "" Then
                        MsgBox("Please make sure to enter DO number on each row.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                Next
            End If

            If MsgBox("Create DO?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                Exit Sub
            End If

            'pilih printer klo memang misalnya jadi ngeprint
            PrinterAdd = ""
            If Me.chkDirectlyPrint.Checked = True Then
                If Me.PrintDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    PrinterAdd = Me.PrintDialog1.PrinterSettings.PrinterName
                Else
                    Exit Sub
                End If
            End If


            'siapkan data stock
            'load data stock
            DTA_ProductQty.Fill(TBL_ProductQty)
            'kurangkan data stock dengan pending do
            DTAPendingDO.Fill(TBLPendingDO)
            For Each rowPendingDO In TBLPendingDO.Rows
                TBL_ProductQty.DefaultView.RowFilter = "[Primary Product] = " + rowPendingDO.Primary_Product.ToString + _
                " AND [Lot No] = '" + rowPendingDO.Batch_No + "' AND [Expiry Date] = #" + Format(rowPendingDO.Expiry_Date, "MM/dd/yyyy") + "#"
                TBL_ProductQty.DefaultView.Sort = "[Storage Area] ASC"

                IndexDummy = 0
                While TBL_ProductQty.DefaultView.Count > IndexDummy And rowPendingDO.Total_Qty > 0
                    row_ProductQty = TBL_ProductQty.DefaultView(IndexDummy).Row
                    If row_ProductQty.Qty >= rowPendingDO.Total_Qty Then
                        row_ProductQty.Qty -= rowPendingDO.Total_Qty
                        rowPendingDO.Total_Qty = 0
                    Else
                        rowPendingDO.Total_Qty -= row_ProductQty.Qty
                        row_ProductQty.Qty = 0
                        TBL_ProductQty.Rows.Remove(row_ProductQty)
                    End If
                End While
            Next

            '3. alokasikan item sob, berdasarkan urutan yang sama yang digunakan dalam perhitungan
            'do yang tidak jadi dikirimkan, tetap dimasukkan dalam perhitungan ini, karena begitulah urutan perhitungannya dilakukan pada 
            'rumusan perhitungan pada form ini atau form shipping forecast. Ini juga menjamin item sob yang sudah melakukan booking sebelumnya,
            'tetap mendapatkan jatah bookingnya
            Me.TBLItemSOB.DefaultView.RowFilter = ""
            Me.TBLItemSOB.DefaultView.Sort = "[booked] DESC, [promise to deliver] ASC, [urutan priority level] ASC, [so date rec] ASC, [primary so] ASC"
            DTABooking.Fill(TBLBooking)
            TBLSOBDummy.Rows.Clear()
            For Each DV In Me.TBLItemSOB.DefaultView
                Me.rowItemSOB = DV.Row
                Me.rowItemSOB.create_do2_expiry_date = #1/1/1950#
                Me.rowItemSOB.create_do2_primary_storage = -1
                Me.rowItemSOB.create_do2_qty_delivery = 0
                Me.rowItemSOB.create_do2_storage_code = ""
                Me.rowItemSOB.create_do2_batch_no = ""
                Me.rowItemSOB.print_grouping = ""
                Me.rowItemSOB.print_detail = ""

                If Me.rowItemSOB.tipe_product = Core_Module.clsMasterProduct.EnumTipeProduct.Services Then
                    'kalo service, selalu ada stock
                    Me.TambahkanItem(TBLSOBDummy, "", #1/1/1950#, -1, _
                    "", Me.rowItemSOB.qty_sob, Me.rowItemSOB)
                Else
                    If Me.rowItemSOB.no_booking.Trim <> "" Then
                        'untuk booked qty, ambil datanya dari tabel booking
                        TBLBooking.DefaultView.RowFilter = "[Primary SO] = " + Me.rowItemSOB.primary_so.ToString + _
                        " AND [Primary Product] = " + Me.rowItemSOB.primary_code_no.ToString + _
                        " AND Qty > 0"

                        For Each DV1 In TBLBooking.DefaultView
                            rowBooking = DV1.Row
                            'untuk kasus partial sending atas qty booking, qty bookingnya disimpan dl, baru nanti ditambahkan kembali
                            JmlBookingDummy = 0
                            If rowBooking.Qty > Me.rowItemSOB.qty_ready_stock Then
                                JmlBookingDummy = rowBooking.Qty - Me.rowItemSOB.qty_ready_stock
                                rowBooking.Qty = Me.rowItemSOB.qty_ready_stock
                            End If
                            TBL_ProductQty.DefaultView.RowFilter = "[Primary Product] = " + Me.rowItemSOB.primary_code_no.ToString + _
                            " AND [Lot No] = '" + rowBooking.Batch_No + "' AND [Expiry Date] = #" + Format(rowBooking.Expiry_Date, "MM/dd/yyyy") + "#" + _
                            " AND Qty > 0"
                            TBL_ProductQty.DefaultView.Sort = "[Storage Area] ASC"
                            IndexDummy = 0
                            While TBL_ProductQty.DefaultView.Count > IndexDummy And rowBooking.Qty > 0
                                row_ProductQty = TBL_ProductQty.DefaultView(IndexDummy).Row
                                If row_ProductQty.Qty >= rowBooking.Qty Then
                                    'tambahkan ke dalam tabel sob dummy
                                    Me.TambahkanItem(TBLSOBDummy, row_ProductQty.Lot_No, row_ProductQty.Expiry_Date, row_ProductQty.Primary_Storage_Area, _
                                    row_ProductQty.Storage_Area, rowBooking.Qty, Me.rowItemSOB)
                                    'kurangkan product qty
                                    row_ProductQty.Qty -= rowBooking.Qty
                                    'kurang qty available sob
                                    Me.rowItemSOB.qty_ready_stock -= rowBooking.Qty
                                    rowBooking.Qty = 0
                                    If row_ProductQty.Qty = 0 Then
                                        TBL_ProductQty.Rows.Remove(row_ProductQty)
                                    End If
                                Else
                                    'tambahkan ke dalam dgv
                                    Me.TambahkanItem(TBLSOBDummy, row_ProductQty.Lot_No, row_ProductQty.Expiry_Date, row_ProductQty.Primary_Storage_Area, _
                                    row_ProductQty.Storage_Area, row_ProductQty.Qty, Me.rowItemSOB)
                                    'kurangkan product qty
                                    rowBooking.Qty -= row_ProductQty.Qty
                                    'kurangkan qty available sob
                                    Me.rowItemSOB.qty_ready_stock -= row_ProductQty.Qty
                                    row_ProductQty.Qty = 0
                                    TBL_ProductQty.Rows.Remove(row_ProductQty)
                                End If
                            End While

                            rowBooking.Qty += JmlBookingDummy
                        Next
                    End If

                    'sisanya ambil dari tabel normal
                    If Me.rowItemSOB.qty_ready_stock > 0 Then
                        TBL_ProductQty.DefaultView.RowFilter = "[Primary Product] = " + Me.rowItemSOB.primary_code_no.ToString + _
                        " AND Qty > 0"
                        TBL_ProductQty.DefaultView.Sort = "[Expiry Date] ASC,[Storage Area] ASC"
                        IndexDummy = 0

                        While TBL_ProductQty.DefaultView.Count > IndexDummy And Me.rowItemSOB.qty_ready_stock > 0
                            row_ProductQty = TBL_ProductQty.DefaultView(IndexDummy).Row
                            If row_ProductQty.Qty >= Me.rowItemSOB.qty_ready_stock Then
                                Me.TambahkanItem(TBLSOBDummy, row_ProductQty.Lot_No, row_ProductQty.Expiry_Date, row_ProductQty.Primary_Storage_Area, _
                                row_ProductQty.Storage_Area, Me.rowItemSOB.qty_ready_stock, Me.rowItemSOB)
                                'kurangkan product qty
                                row_ProductQty.Qty -= Me.rowItemSOB.qty_ready_stock
                                'kurang qty available sob
                                Me.rowItemSOB.qty_ready_stock = 0
                                If row_ProductQty.Qty = 0 Then
                                    TBL_ProductQty.Rows.Remove(row_ProductQty)
                                End If
                            Else
                                'tambahkan ke dalam dgv
                                Me.TambahkanItem(TBLSOBDummy, row_ProductQty.Lot_No, row_ProductQty.Expiry_Date, row_ProductQty.Primary_Storage_Area, _
                                row_ProductQty.Storage_Area, row_ProductQty.Qty, Me.rowItemSOB)
                                'kurangkan qty available sob
                                Me.rowItemSOB.qty_ready_stock -= row_ProductQty.Qty
                                row_ProductQty.Qty = 0
                                TBL_ProductQty.Rows.Remove(row_ProductQty)
                            End If
                        End While
                    End If
                End If
            Next

            '4. start import do berdasarkan row dgv
            If Me.chkDirectlySave.Checked = True Then
                Me.btnCancelProses.Visible = True
                Me.btnCancelProses.Enabled = True
                Me.btnCancelProses.TabStop = True

                Me.btnOK.Enabled = False
                Me.btnCancel.Enabled = False
                Me.ControlBox = False
                Me.btnBack.Enabled = False
                Me.dtpDate.Enabled = False
                Me.chkDirectlySave.Enabled = False
                Me.chkDirectlyPrint.Enabled = False
                Me.chkDirectlyCreateStockOut.Enabled = False
                Me.chkDirectlyCreateDO.Enabled = False
                Me.dgvDONumbering.ReadOnly = True

                Me.TBLSOBDummy1 = TBLSOBDummy
                Me.BackgroundWorker2.RunWorkerAsync()
            Else
                For Each rowDGV In Me.dgvDONumbering.Rows
                    If Me.StopProses = True Then
                        MsgBox("Process stopped by user.", MsgBoxStyle.Information, "Information")
                        Me.btnCancelProses.Visible = False
                        Exit Sub
                    End If
                    If Me.BuatDO(rowDGV, TBLSOBDummy) = False Then
                        Exit For
                    End If
                Next

                'kalo sampe sini, maka tutup form
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_ProductQty.Dispose()
            TBL_ProductQty.Dispose()
            DTAPendingDO.Dispose()
            TBLPendingDO.Dispose()
            DTABooking.Dispose()
            TBLBooking.Dispose()
            TBLSOBDummy.Dispose()
        End Try
    End Sub

    Private Sub TambahkanItem(ByVal TBLSOBDummy As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable, ByVal vBatchNo As String, _
    ByVal vExpiry As Date, ByVal vPrimaryLocation As Double, ByVal vLocation As String, ByVal vQty As Decimal, _
    ByVal vRowAcuan As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow)
        Dim rowSOBDummy As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow

        Try
            TBLSOBDummy.DefaultView.RowFilter = "[create do2 expiry date] = #" + Format(vExpiry, "MM/dd/yyyy") + "#" + _
            " AND [create do2 batch no] = '" + vBatchNo + "'" + _
            " AND [primary so] = " + vRowAcuan.primary_so.ToString + _
            " AND [primary code no] = " + vRowAcuan.primary_code_no.ToString

            If TBLSOBDummy.DefaultView.Count = 0 Then
                rowSOBDummy = TBLSOBDummy.Newtbltrans_shipping_forecast_itemdeliveryRow
                rowSOBDummy.primary = Me.GetPrimaryMax(TBLSOBDummy) + 1
                For ab As Integer = 1 To TBLSOBDummy.Columns.Count - 1
                    rowSOBDummy.Item(ab) = vRowAcuan.Item(ab)
                Next
                TBLSOBDummy.Rows.Add(rowSOBDummy)
                rowSOBDummy.create_do2_expiry_date = vExpiry
                rowSOBDummy.create_do2_primary_storage = vPrimaryLocation
                rowSOBDummy.create_do2_qty_delivery = vQty
                rowSOBDummy.create_do2_storage_code = vLocation
                rowSOBDummy.create_do2_batch_no = vBatchNo
                rowSOBDummy.print_grouping = ""
                rowSOBDummy.print_detail = ""
            Else
                rowSOBDummy = TBLSOBDummy.DefaultView(0).Row
                rowSOBDummy.create_do2_qty_delivery += vQty
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahPerluOtorisasi(ByVal vPrimaryCust As Double) As Boolean
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvException.Rows
                If rowDGV.Cells(Me.ColExc_PrimaryCust.Index).Value = vPrimaryCust Then
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function BuatDO(ByVal rowDGV1 As DataGridViewRow, ByVal TBLSOBDummy As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable) As Boolean
        Dim ij As New FrmTransaksiDO3
        Dim Dummy_UseTax As Integer = 0
        Dim DV As Data.DataRowView
        Dim IndexDummy As Integer = 0
        Dim ExpiredDateStr As String = ""
        Dim dtaSOMulti As New DseTransaksiShippingForecastTableAdapters.tbltans_so_item1TableAdapter
        Dim tblSOMulti As New DseTransaksiShippingForecast.tbltans_so_item1DataTable
        Dim rowSOMulti As DseTransaksiShippingForecast.tbltans_so_item1Row
        Dim rowSOBDummy As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim RetVal As Boolean = True
        Dim Result As Windows.Forms.DialogResult
        Dim vFrm As New FrmTransaksiDO3
        Dim rowDgv As DataGridViewRow
        Dim DefPrecMulti As Integer = 0
        Dim PrimarySOMin As Double = -1
        Dim PrimairySOMax As Double = -1
        Dim QtyMulti As Decimal = 0

        Try
            Me.OnProgressDO = True

            ij = New FrmTransaksiDO3
            'cari data qty available: qty on hand - qty pending do
            'qty booking sudah diperhitungkan dalam perhitungan shipping forecast, jadi dianggap masih sama
            'qty on hand - qty pending do, barusan diperhitungkan dalam shipping forecast, jadi dianggap masih sama
            'jika sampai ada masalah pada perhitungan qty available, diharapkan sudah di blok pada fungsi apakah data benar di do

            'isikan data header
            ij.Show()
            ij.Location = New Point((ij.Width * -1) - 1, (ij.Height * -1) - 1)
            ij.UserCode = Me.vUserCode

            ij.DO_Baru()
            ij.txtPrimaryCustomer.Text = rowDGV1.Cells(Me.ColDONo_PrimaryCust.Index).Value
            ij.SetDataCustomer()
            ij.TxtRemark.Text = rowDGV1.Cells(Me.ColDONo_Remarks.Index).Value

            'isikan data item
            If rowDGV1.Cells(Me.ColDONo_UseTax.Index).Value = Boolean.TrueString Then
                Dummy_UseTax = 1
            Else
                Dummy_UseTax = 0
            End If

            If Me.vDObySO = True Then
                TBLSOBDummy.DefaultView.RowFilter = "[primary so] = " & rowDGV1.Cells(Me.ColDONo_PrimarySO.Index).Value
            Else
                TBLSOBDummy.DefaultView.RowFilter = "[primary cust] = " + rowDGV1.Cells(Me.ColDONo_PrimaryCust.Index).Value.ToString + _
                " AND [create do primary currency] = " + rowDGV1.Cells(Me.ColDONo_PrimaryCur.Index).Value.ToString + _
                " AND [create do use tax] = " + Dummy_UseTax.ToString + " AND [create do tax percentage] = " + CDec(rowDGV1.Cells(Me.ColDONo_TarPercentage.Index).Value).ToString
            End If
            
            If TBLSOBDummy.DefaultView.Count > 0 Then
                TBLSOBDummy.DefaultView.Sort = "[primary so] ASC"
                rowSOBDummy = TBLSOBDummy.DefaultView(0).Row
                PrimarySOMin = rowSOBDummy.primary_so
                rowSOBDummy = TBLSOBDummy.DefaultView(TBLSOBDummy.DefaultView.Count - 1).Row
                PrimairySOMax = rowSOBDummy.primary_so
            End If

            dtaSOMulti.Connection.ConnectionString = Core_Module.GetConstring
            dtaSOMulti.FillBy_PrimarySOMaxMin(tblSOMulti, PrimairySOMax, PrimarySOMin)

            TBLSOBDummy.DefaultView.Sort = "[booked] DESC, [promise to deliver] ASC, [urutan priority level] ASC, [so date rec] ASC, [primary so] ASC"
            For Each DV In TBLSOBDummy.DefaultView
                rowSOBDummy = DV.Row

                tblSOMulti.DefaultView.RowFilter = "[Primary SO] = " & rowSOBDummy.primary_so.ToString & " AND [Primary Product] = " & rowSOBDummy.primary_code_no
                If tblSOMulti.DefaultView.Count > 0 Then
                    rowSOMulti = tblSOMulti.DefaultView(0).Row
                Else
                    'tidak mungkin tidak ada baris so yang di filter seharusnya selalu memiliki satu baris item
                    MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                If rowSOBDummy.create_do2_expiry_date = #1/1/1950# Then
                    ExpiredDateStr = "-"
                Else
                    ExpiredDateStr = Format(rowSOBDummy.create_do2_expiry_date, "dd MMM yyyy")
                End If
                
                With rowSOBDummy
                    DefPrecMulti = ObjProduct.GetAngkaDecimal(.primary_code_no, rowSOMulti.Primary_Units_Multi)
                    ij.dgvItem.Rows.Add()
                    rowDgv = ij.dgvItem.Rows(ij.dgvItem.Rows.Count - 1)
                    rowDgv.Cells(vFrm.ColNo.Index).Value = ij.dgvItem.RowCount
                    rowDgv.Cells(vFrm.ColCodeNo.Index).Value = .Product_Code
                    rowDgv.Cells(vFrm.ColDescription.Index).Value = .Description
                    rowDgv.Cells(vFrm.ColCustCodeNo.Index).Value = ""
                    rowDgv.Cells(vFrm.ColCustDesc.Index).Value = ""
                    rowDgv.Cells(vFrm.ColSONO.Index).Value = .so_number
                    rowDgv.Cells(vFrm.ColBatchNo.Index).Value = .create_do2_batch_no
                    rowDgv.Cells(vFrm.ColExpiryDate.Index).Value = ExpiredDateStr
                    rowDgv.Cells(vFrm.ColQtyMaster.Index).Value = 0
                    rowDgv.Cells(vFrm.ColQtyPerMaster.Index).Value = 0
                    rowDgv.Cells(vFrm.ColMasterUnits.Index).Value = ""
                    rowDgv.Cells(vFrm.ColQty.Index).Value = 0
                    rowDgv.Cells(vFrm.ColTotalQty.Index).Value = ObjCD.DigitGrouping(.create_do2_qty_delivery, ._Decimal)
                    rowDgv.Cells(vFrm.ColUnit.Index).Value = .Unit_Code
                    rowDgv.Cells(vFrm.ColSecQty.Index).Value = -1
                    rowDgv.Cells(vFrm.ColSecUnit.Index).Value = ""
                    rowDgv.Cells(vFrm.ColRemarks.Index).Value = ""
                    rowDgv.Cells(vFrm.ColPrimaryCodeNo.Index).Value = .primary_code_no
                    rowDgv.Cells(vFrm.ColPrimarySO.Index).Value = .primary_so
                    rowDgv.Cells(vFrm.ColPrimaryMasterUnits.Index).Value = -1
                    rowDgv.Cells(vFrm.ColPrimaryUnit.Index).Value = .primary_unit
                    rowDgv.Cells(vFrm.ColPrimarySecUnit.Index).Value = -1
                    'qtyMulti : qty delivery DO yang di buat multi nya(Qty tersebut menentukan qty yang siap kirim)
                    QtyMulti = Me.ObjCD.DigitGrouping((.create_do2_qty_delivery * rowSOMulti.Ratio_Reference) / rowSOMulti.Ratio_Default, DefPrecMulti)

                    If rowSOMulti.Qty_Multi > QtyMulti Then
                        rowSOMulti.Qty_Multi -= QtyMulti
                        rowSOMulti.SOB_Multi -= QtyMulti
                        rowSOMulti.Qty -= .create_do2_qty_delivery
                        If rowSOMulti.Qty = 0 Then
                            QtyMulti += rowSOMulti.Qty_Multi
                            rowSOMulti.Qty_Multi = 0
                            rowSOMulti.SOB_Multi = 0
                        End If
                    Else
                        QtyMulti = rowSOMulti.Qty_Multi
                    End If

                    rowDgv.Cells(vFrm.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(QtyMulti, DefPrecMulti)
                    rowDgv.Cells(vFrm.ColUnitsCodeMulti.Index).Value = rowSOMulti.Unit_Multi
                    rowDgv.Cells(vFrm.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(rowSOMulti.Ratio_Default, ._Decimal)
                    rowDgv.Cells(vFrm.ColRatioReference.Index).Value = ObjCD.DigitGrouping(rowSOMulti.Ratio_Reference, DefPrecMulti)
                    rowDgv.Cells(vFrm.ColPrimaryUnitsMulti.Index).Value = rowSOMulti.Primary_Units_Multi
                    If rowSOMulti.Primary_Units_Multi <> .primary_unit Then
                        rowDgv.Cells(vFrm.ColRatioDesc.Index).Value = ObjCD.DigitGrouping(rowSOMulti.Ratio_Reference, DefPrecMulti) _
                        & " " & rowSOMulti.Unit_Multi & " = " & ObjCD.DigitGrouping(rowSOMulti.Ratio_Default, ._Decimal) & " " & .Unit_Code
                    Else
                        rowDgv.Cells(vFrm.ColRatioDesc.Index).Value = "-"
                    End If
                End With

                'ij.dgvItem.Rows.Add(ij.dgvItem.RowCount + 1, rowSOBDummy.Product_Code, rowSOBDummy.Description, "", _
                '"", rowSOBDummy.so_number, rowSOBDummy.create_do2_batch_no, ExpiredDateStr, _
                '0, 0, "", 0, ObjCD.DigitGrouping(rowSOBDummy.create_do2_qty_delivery, rowSOBDummy._Decimal), rowSOBDummy.Unit_Code, -1, "", _
                '"", rowSOBDummy.primary_code_no, rowSOBDummy.primary_so, -1, rowSOBDummy.primary_unit, _
                '-1)
            Next

            'tampilkan form do
            ij.BackColorAll = Me.BackColorAll
            ij.SetGboItem()
            ij.FormatDGV()
            ij.Hide()
            ij.btnStopProcess.Visible = True
            ij.lblProgress.Visible = True
            ij.prgAll.Visible = True
            ij.prgAll.Maximum = Me.dgvDONumbering.RowCount
            ij.prgAll.Value = rowDGV1.Index + 1
            ij.lblProgress.Text = "Data " + (rowDGV1.Index + 1).ToString + " of " + Me.dgvDONumbering.RowCount.ToString
            ij.LangsungCetakPackingList = Me.chkDirectlyPrint.Checked
            ij.PrintAdd = Me.PrinterAdd
            If Me.objNumber.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO) = False Then
                ij.TxtDONo.Text = rowDGV1.Cells(Me.ColDONo_DoNumber.Index).Value
            End If
            Me.LblProgress.Text = "Data " + (rowDGV1.Index + 1).ToString + " of " + Me.dgvDONumbering.RowCount.ToString + " - " + rowDGV1.Cells(Me.ColDONo_Customer.Index).Value
            Me.LblProgress.Refresh()

            Me.prgAll.Visible = True
            Me.prgAll.Maximum = Me.dgvDONumbering.RowCount
            Me.prgAll.Value = rowDGV1.Index + 1

            If ij.dgvItem.RowCount > 0 Then
                If Me.chkDirectlySave.Checked = False Then
                    Me.btnCancelProses.Visible = False
                    Result = ij.ShowDialog
                    If Result = Windows.Forms.DialogResult.Abort Then
                        RetVal = False
                    ElseIf Result = Windows.Forms.DialogResult.OK Then
                        If Me.ObjOtorisasi.Primary <> -1 Then
                            If Me.ApakahPerluOtorisasi(rowDGV1.Cells(Me.ColDONo_PrimaryCust.Index).Value) = True Then
                                With Me.ObjOtorisasi
                                    .PrimaryDokumenAcuan = ij.TxtPrimaryDO.Text
                                    .NoDocumentHist = ij.TxtDONo.Text
                                End With
                                Me.ObjOtorisasi.SimpanData()
                            End If
                        End If
                    End If
                Else
                    Me.btnCancelProses.Visible = True
                    If ij.SimpanDOBaruLangsung = False Then
                        If ij.ShowDialog() = Windows.Forms.DialogResult.Abort Then
                            RetVal = False
                        End If
                    Else
                        If Me.ObjOtorisasi.Primary <> -1 Then
                            If Me.ApakahPerluOtorisasi(rowDGV1.Cells(Me.ColDONo_PrimaryCust.Index).Value) = True Then
                                With Me.ObjOtorisasi
                                    .PrimaryDokumenAcuan = ij.TxtPrimaryDO.Text
                                    .NoDocumentHist = ij.TxtDONo.Text
                                End With
                                Me.ObjOtorisasi.SimpanData()
                            End If
                        End If
                    End If
                End If
            Else
                MsgBox("No Stock is available, other user may have use the first.", MsgBoxStyle.Information, "Information")
            End If

            ij = Nothing

            'tandai supaya back ground worker bisa jalan buat do utnuk baris selanjutnya
            Me.OnProgressDO = False

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij = Nothing
            Me.OnProgressDO = False
            dtaSOMulti.Dispose()
            tblSOMulti.Dispose()
        End Try
    End Function

    Private Sub dgvDONumbering_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvDONumbering.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvDONumbering, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancelProses_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancelProses.Click
        Try
            If MsgBox("Cancel creating DO?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.StopProses = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            For Each rowDGV In Me.dgvDONumbering.Rows
                If Me.StopProses = True Then
                    Me.BackgroundWorker2.ReportProgress(1)
                    Exit Sub
                End If
                Me.rowDGV = rowDGV

                Me.BackgroundWorker2.ReportProgress(0)
                Threading.Thread.Sleep(1000)
                While Me.OnProgressDO = True
                    Threading.Thread.Sleep(1000)
                End While
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker2.ProgressChanged
        Try
            If e.ProgressPercentage = 0 Then
                If Me.BuatDO(Me.rowDGV, Me.TBLSOBDummy1) = False Then
                    Exit Sub
                End If
            Else
                MsgBox("Process stopped by user.", MsgBoxStyle.Information, "Information")
                Me.btnCancelProses.Visible = False
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.btnCancelProses.Visible = False
            Me.btnCancelProses.Enabled = False
            Me.btnCancelProses.TabStop = False

            If Me.StopProses = False Then
                MsgBox("DO creation finish.", MsgBoxStyle.Information, "Information")
            End If

            Me.btnOK.Enabled = True
            Me.btnCancel.Enabled = True
            Me.ControlBox = True
            Me.btnBack.Enabled = True

            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class