'''@rosadinnn

Option Explicit On
Option Strict Off

Public Class FrmCetakAR

    Dim objCd As New Core_Module.ClsCekData
    Dim objCust As New Core_Module.clsMasterCustomer
    Dim objCustGroup As New Core_Module.clsMasterCustomerGroup
    Dim objCurr As New Core_Module.clsMasterCurrency
    Dim objAcc As New Core_Module.ClsMasterAkun1
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim ObjLicense As New Core_Module.ClsLicense

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
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

            Me.Cursor = Cursors.WaitCursor
            If Me.rboDetail.Checked = True Then
                Me.LoadDataDetail()
            Else
                Me.LoadData()
            End If
        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.GotFocus, txtCustomerGroupCode.GotFocus, txtPaymentAccountCode.GotFocus, txtArAccountCode.GotFocus, txtCurrencyCode.GotFocus
        Dim txt As TextBox
        txt = TryCast(sender, TextBox)
        Try
            txt.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub txtArAccountCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.LostFocus, txtCustomerGroupCode.LostFocus, txtPaymentAccountCode.LostFocus, txtArAccountCode.LostFocus, txtCurrencyCode.LostFocus
        Dim txt As TextBox
        txt = TryCast(sender, TextBox)
        Try
            txt.BackColor = Color.White
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub LoadData()
        Dim dse As New DseLaporanARPayment
        Dim dta As New DseLaporanARPaymentTableAdapters.DataTable1TableAdapter
        Dim tbl As DseLaporanARPayment.DataTable1DataTable = dse.DataTable1
        Dim row, rowdummy As DseLaporanARPayment.DataTable1Row
        Dim rpt As New CrpLaporanARPayment
        Dim tbldummy As New DseLaporanARPayment.DataTable1DataTable
        Dim Filter As String = ""

        'Dim dta1 As New DseLaporanARPaymentTableAdapters.tbltrans_salesinvoiceTableAdapter
        'Dim tbl1 As New DseLaporanARPayment.tbltrans_salesinvoiceDataTable
        'Dim row1 As DseLaporanARPayment.tbltrans_salesinvoiceRow

        'Dim dta2 As New DseLaporanARPaymentTableAdapters.tbltrans_historicalsalesTableAdapter
        'Dim tbl2 As New DseLaporanARPayment.tbltrans_historicalsalesDataTable
        'Dim row2 As DseLaporanARPayment.tbltrans_historicalsalesRow

        'Dim dta3 As New DseLaporanARPaymentTableAdapters.tbltrans_salesreturnTableAdapter
        'Dim tbl3 As New DseLaporanARPayment.tbltrans_salesreturnDataTable
        'Dim row3 As DseLaporanARPayment.tbltrans_salesreturnRow

        'Dim dta4 As New DseLaporanARPaymentTableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter
        'Dim tbl4 As New DseLaporanARPayment.tbltrans_historicalsalesreturn_manualDataTable
        'Dim row4 As DseLaporanARPayment.tbltrans_historicalsalesreturn_manualRow

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            tbl.Rows.Clear()
            tbldummy.Rows.Clear()

            ''filter dg tanggal dan tipe jurnal
            dta.FillBy2_Iwan(tbldummy, Me.dtpFrom.Value, Me.dtpTo.Value, ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal)
            Me.SetTanggalNoPO(tbldummy)
            Filter = ""
            If Me.txtPrimaryCustomerGroup.Text <> "-1" Then
                If Filter = "" Then
                    Filter = "[Primary Group Pihak Acuan] = " + Me.txtPrimaryCustomerGroup.Text
                Else
                    Filter = Filter + " AND " + "[Primary Group Pihak Acuan] = " + Me.txtPrimaryCustomerGroup.Text
                End If
            End If
            If Me.txtPrimaryCustomerCode.Text <> "-1" Then
                If Filter = "" Then
                    Filter = "[Primary Pihak Acuan] = " + Me.txtPrimaryCustomerCode.Text
                Else
                    Filter = Filter + " AND " + "[Primary Pihak Acuan] = " + Me.txtPrimaryCustomerCode.Text
                End If
            End If
            If Me.txtPrimaryCurrency.Text <> "-1" Then
                If Filter = "" Then
                    Filter = "[primary currency] = " + Me.txtPrimaryCurrency.Text
                Else
                    Filter = Filter + " AND " + "[primary currency] = " + Me.txtPrimaryCurrency.Text
                End If
            End If
            If Me.txtPrimaryArAccount.Text <> "-1" Then
                If Filter = "" Then
                    Filter = "[primary akun ap ar] = " + Me.txtPrimaryArAccount.Text
                Else
                    Filter = Filter + " AND " + "[primary akun ap ar] = " + Me.txtPrimaryArAccount.Text
                End If
            End If
            If Me.txtPrimaryPaymentAccount.Text <> "-1" Then
                If Filter = "" Then
                    Filter = "[primary akun pay to] = " + Me.txtPrimaryPaymentAccount.Text
                Else
                    Filter = Filter + " AND " + "[primary akun pay to] = " + Me.txtPrimaryPaymentAccount.Text
                End If
            End If
            tbldummy.DefaultView.RowFilter = Filter
            For Each DV As DataRowView In tbldummy.DefaultView
                rowdummy = DV.Row
                row = tbl.NewDataTable1Row
                For ij As Integer = 0 To tbl.Columns.Count - 1
                    row.Item(ij) = rowdummy.Item(ij)
                Next
                tbl.Rows.Add(row)
            Next

            If tbl.Rows.Count > 0 Then
                Me.SetGroup(tbl)
                Me.UrutkanUlang(tbl)
                rpt.SetDataSource(dse)
                If rboCustCode.Checked = True Then
                    rpt.SetParameterValue(rpt.Parameter_GroupBy.ParameterFieldName, "Customer")
                ElseIf rboCustGroup.Checked = True Then
                    rpt.SetParameterValue(rpt.Parameter_GroupBy.ParameterFieldName, "Customer Group")
                ElseIf rboPaymentAccount.Checked = True Then
                    rpt.SetParameterValue(rpt.Parameter_GroupBy.ParameterFieldName, "Payment Account")
                ElseIf rboArAccount.Checked = True Then
                    rpt.SetParameterValue(rpt.Parameter_GroupBy.ParameterFieldName, "AR Account")
                ElseIf rboCurrency.Checked = True Then
                    rpt.SetParameterValue(rpt.Parameter_GroupBy.ParameterFieldName, "Currency")
                End If
                rpt.SetParameterValue(rpt.Parameter_Tittle.ParameterFieldName, "Account Receivable Payment Report")
                rpt.SetParameterValue(rpt.Parameter_PayTo.ParameterFieldName, "Pay To")
                Me.crvAR.ReportSource = rpt
                Me.crvAR.Zoom(85)
                Me.crvAR.Refresh()
                Me.crvAR.Show()
            Else
                MsgBox("There isn't data in this period", MsgBoxStyle.Information, "Information")
                Me.crvAR.ReportSource = Nothing
                Me.crvAR.Refresh()
                Me.crvAR.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDataDetail()
        Dim dse As New DseLaporanARPayment
        Dim dta As New DseLaporanARPaymentTableAdapters.DataTable1TableAdapter
        Dim tbl As DseLaporanARPayment.DataTable1DataTable = dse.DataTable1
        Dim row, rowdummy As DseLaporanARPayment.DataTable1Row
        Dim rpt As New CrpLaporanARPaymentDetail
        Dim tbldummy As New DseLaporanARPayment.DataTable1DataTable
        Dim Filter As String = ""

        'Dim dta1 As New DseLaporanARPaymentTableAdapters.tbltrans_salesinvoiceTableAdapter
        'Dim tbl1 As New DseLaporanARPayment.tbltrans_salesinvoiceDataTable
        'Dim row1 As DseLaporanARPayment.tbltrans_salesinvoiceRow

        'Dim dta2 As New DseLaporanARPaymentTableAdapters.tbltrans_historicalsalesTableAdapter
        'Dim tbl2 As New DseLaporanARPayment.tbltrans_historicalsalesDataTable
        'Dim row2 As DseLaporanARPayment.tbltrans_historicalsalesRow

        'Dim dta3 As New DseLaporanARPaymentTableAdapters.tbltrans_salesreturnTableAdapter
        'Dim tbl3 As New DseLaporanARPayment.tbltrans_salesreturnDataTable
        'Dim row3 As DseLaporanARPayment.tbltrans_salesreturnRow

        'Dim dta4 As New DseLaporanARPaymentTableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter
        'Dim tbl4 As New DseLaporanARPayment.tbltrans_historicalsalesreturn_manualDataTable
        'Dim row4 As DseLaporanARPayment.tbltrans_historicalsalesreturn_manualRow

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            tbl.Rows.Clear()
            tbldummy.Rows.Clear()

            ''filter dg tanggal dan tipe jurnal
            dta.FillBy2_Iwan(tbldummy, Me.dtpFrom.Value, Me.dtpTo.Value, ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal)
            Me.SetTanggalNoPO(tbldummy)
            Filter = ""
            If Me.txtPrimaryCustomerGroup.Text <> "-1" Then
                If Filter = "" Then
                    Filter = "[Primary Group Pihak Acuan] = " + Me.txtPrimaryCustomerGroup.Text
                Else
                    Filter = Filter + " AND " + "[Primary Group Pihak Acuan] = " + Me.txtPrimaryCustomerGroup.Text
                End If
            End If
            If Me.txtPrimaryCustomerCode.Text <> "-1" Then
                If Filter = "" Then
                    Filter = "[Primary Pihak Acuan] = " + Me.txtPrimaryCustomerCode.Text
                Else
                    Filter = Filter + " AND " + "[Primary Pihak Acuan] = " + Me.txtPrimaryCustomerCode.Text
                End If
            End If
            If Me.txtPrimaryCurrency.Text <> "-1" Then
                If Filter = "" Then
                    Filter = "[primary currency] = " + Me.txtPrimaryCurrency.Text
                Else
                    Filter = Filter + " AND " + "[primary currency] = " + Me.txtPrimaryCurrency.Text
                End If
            End If
            If Me.txtPrimaryArAccount.Text <> "-1" Then
                If Filter = "" Then
                    Filter = "[primary akun ap ar] = " + Me.txtPrimaryArAccount.Text
                Else
                    Filter = Filter + " AND " + "[primary akun ap ar] = " + Me.txtPrimaryArAccount.Text
                End If
            End If
            If Me.txtPrimaryPaymentAccount.Text <> "-1" Then
                If Filter = "" Then
                    Filter = "[primary akun pay to] = " + Me.txtPrimaryPaymentAccount.Text
                Else
                    Filter = Filter + " AND " + "[primary akun pay to] = " + Me.txtPrimaryPaymentAccount.Text
                End If
            End If

            tbldummy.DefaultView.RowFilter = Filter
            Me.TambahkanDetailOthers(tbldummy, Me.dtpFrom.Value, Me.dtpTo.Value)
            For Each DV As DataRowView In tbldummy.DefaultView
                rowdummy = DV.Row
                row = tbl.NewDataTable1Row
                For ij As Integer = 0 To tbl.Columns.Count - 1
                    row.Item(ij) = rowdummy.Item(ij)
                Next
                tbl.Rows.Add(row)
            Next


            Dim ListD As New List(Of Double)
            For Each row In tbl.Rows
                'If ListD.Contains(row.Primary_Journal) = False Then
                '    ListD.Add(row.Primary_Journal)
                'End If
                Dim decPaymentHeader As Decimal = 0
                tbldummy.DefaultView.RowFilter = Nothing
                tbldummy.DefaultView.RowFilter = "[Primary Journal] = " & row.Primary_Journal
                If tbldummy.DefaultView.Count > 0 Then
                    For Each dv2 As DataRowView In tbldummy.DefaultView
                        rowdummy = dv2.Row
                        If IsNumeric(rowdummy.Payment) Then
                            decPaymentHeader += CDec(rowdummy.Payment)
                        End If
                    Next
                End If
                row.Payment_Header = decPaymentHeader
            Next
            If tbl.Rows.Count > 0 Then
                Me.SetGroup(tbl)
                Me.UrutkanUlang(tbl)

                'Dim i As New Form5
                'i.DataGridView1.DataSource = Nothing
                'i.DataGridView1.DataSource = tbl
                'i.ShowDialog()


                rpt.SetDataSource(dse)
                rpt.SetParameterValue(rpt.Parameter_GroupBy.ParameterFieldName, "Customer")
                rpt.SetParameterValue(rpt.Parameter_Tittle.ParameterFieldName, "Customer Receivable Report Detail")
                rpt.SetParameterValue(rpt.Parameter_PayTo.ParameterFieldName, "Pay to")

                Me.crvAR.ReportSource = rpt
                Me.crvAR.Zoom(85)
                Me.crvAR.Refresh()
                Me.crvAR.Show()
            Else
                MsgBox("There isn't data in this period", MsgBoxStyle.Information, "Information")
                Me.crvAR.ReportSource = Nothing
                Me.crvAR.Refresh()
                Me.crvAR.Show()
            End If

            'If tbl.Rows.Count > 0 Then
            '    Me.SetGroup(tbl)
            '    Me.UrutkanUlang(tbl)
            '    rpt.SetDataSource(dse)
            '    If rboCustCode.Checked = True Then
            '        rpt.SetParameterValue(rpt.Parameter_GroupBy.ParameterFieldName, "Customer")
            '    ElseIf rboCustGroup.Checked = True Then
            '        rpt.SetParameterValue(rpt.Parameter_GroupBy.ParameterFieldName, "Customer Group")
            '    ElseIf rboPaymentAccount.Checked = True Then
            '        rpt.SetParameterValue(rpt.Parameter_GroupBy.ParameterFieldName, "Payment Account")
            '    ElseIf rboArAccount.Checked = True Then
            '        rpt.SetParameterValue(rpt.Parameter_GroupBy.ParameterFieldName, "AR Account")
            '    ElseIf rboCurrency.Checked = True Then
            '        rpt.SetParameterValue(rpt.Parameter_GroupBy.ParameterFieldName, "Currency")
            '    End If
            '    rpt.SetParameterValue(rpt.Parameter_Tittle.ParameterFieldName, "Account Receivable Payment Report")
            '    rpt.SetParameterValue(rpt.Parameter_PayTo.ParameterFieldName, "Pay To")
            '    Me.crvAR.ReportSource = rpt
            '    Me.crvAR.Zoom(85)
            '    Me.crvAR.Refresh()
            '    Me.crvAR.Show()
            'Else
            '    MsgBox("There isn't data in this period", MsgBoxStyle.Information, "Information")
            '    Me.crvAR.ReportSource = Nothing
            '    Me.crvAR.Refresh()
            '    Me.crvAR.Show()
            'End If
        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetTanggalNoPO(ByRef vtbl As DseLaporanARPayment.DataTable1DataTable)
        Dim vrow As DseLaporanARPayment.DataTable1Row

        Dim DTAInvoice As New DseLaporanARPaymentTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim DTAInvoiceDP As New DseLaporanARPaymentTableAdapters.tbltrans_salesinvoiceDPTableAdapter
        Dim DTAReturn As New DseLaporanARPaymentTableAdapters.tbltrans_salesreturnTableAdapter
        Dim DTAHSales As New DseLaporanARPaymentTableAdapters.tbltrans_historicalsalesTableAdapter
        Dim DTAHSReturn As New DseLaporanARPaymentTableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter
        Dim DTACDNote As New DseLaporanARPaymentTableAdapters.tbltrans_credit_note_InvoiceTableAdapter

        Dim tblInvoice As New DseLaporanARPayment.tbltrans_salesinvoiceDataTable
        Dim tblInvoiceDP As New DseLaporanARPayment.tbltrans_salesinvoiceDPDataTable
        Dim tblReturn As New DseLaporanARPayment.tbltrans_salesreturnDataTable
        Dim tblHSales As New DseLaporanARPayment.tbltrans_historicalsalesDataTable
        Dim tblHReturn As New DseLaporanARPayment.tbltrans_historicalsalesreturn_manualDataTable
        Dim tblCDNote As New DseLaporanARPayment.tbltrans_credit_note_InvoiceDataTable

        Dim rowInvoice As DseLaporanARPayment.tbltrans_salesinvoiceRow
        Dim rowInvoiceDP As DseLaporanARPayment.tbltrans_salesinvoiceDPRow
        Dim rowReturn As DseLaporanARPayment.tbltrans_salesreturnRow
        Dim rowHSales As DseLaporanARPayment.tbltrans_historicalsalesRow
        Dim rowHReturn As DseLaporanARPayment.tbltrans_historicalsalesreturn_manualRow
        Dim rowCDNote As DseLaporanARPayment.tbltrans_credit_note_InvoiceRow


        Try
            If vtbl.Rows.Count > 0 Then
                DTAInvoice.Fill(tblInvoice)
                DTAInvoiceDP.Fill(tblInvoiceDP)
                DTAReturn.Fill(tblReturn)
                DTAHSales.Fill(tblHSales)
                DTAHSReturn.Fill(tblHReturn)
                DTACDNote.Fill(tblCDNote)

                Dim ListX As New List(Of String)

                For Each vrow In vtbl.Rows
                    With vrow
                        Dim strPONo As String = ""
                        If .Tipe_Document_Acuan = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic Then
                            tblInvoice.DefaultView.RowFilter = Nothing
                            tblInvoice.DefaultView.RowFilter = "[Primary] = " & .Primary_Document_Acuan
                            If tblInvoice.DefaultView.Count > 0 Then
                                Dim dv As Data.DataRowView
                                For Each dv In tblInvoice.DefaultView
                                    rowInvoice = dv.Row
                                    If ListX.Contains(rowInvoice.No_Transaksi & "-" & rowInvoice.No_SO) = False Then
                                        If strPONo = "" Then
                                            strPONo &= rowInvoice.No_SO
                                        Else
                                            If Not String.IsNullOrEmpty(rowInvoice.No_SO) Then
                                                strPONo &= ", " & rowInvoice.No_SO
                                            Else
                                                strPONo &= ""
                                            End If
                                        End If
                                        ListX.Add(rowInvoice.No_Transaksi & "-" & rowInvoice.No_SO)
                                    End If
                                    .Tipe_Document_Acuan_Desc = "Invoice"
                                    .Invoice_Date = rowInvoice.Tanggal_Transaksi
                                Next
                            End If

                            tblInvoiceDP.DefaultView.RowFilter = Nothing
                            tblInvoiceDP.DefaultView.RowFilter = "[Primary] = " & .Primary_Document_Acuan
                            If tblInvoiceDP.DefaultView.Count > 0 Then
                                Dim dv As Data.DataRowView
                                For Each dv In tblInvoiceDP.DefaultView
                                    rowInvoiceDP = dv.Row
                                    If ListX.Contains(rowInvoiceDP.No_Transaksi & "-" & rowInvoiceDP.No_SO) = False Then
                                        If strPONo = "" Then
                                            strPONo &= rowInvoiceDP.No_SO
                                        Else
                                            If Not String.IsNullOrEmpty(rowInvoiceDP.No_SO) Then
                                                strPONo &= ", " & rowInvoiceDP.No_SO
                                            Else
                                                strPONo &= ""
                                            End If
                                        End If
                                        ListX.Add(rowInvoiceDP.No_Transaksi & "-" & rowInvoiceDP.No_SO)
                                    End If
                                    .Tipe_Document_Acuan_Desc = "Invoice DP"
                                    .Invoice_Date = rowInvoiceDP.Tanggal_Transaksi
                                Next
                            End If

                        ElseIf .Tipe_Document_Acuan = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual Then
                            tblHSales.DefaultView.RowFilter = Nothing
                            tblHSales.DefaultView.RowFilter = "[Primary] = " & .Primary_Document_Acuan
                            If tblHSales.DefaultView.Count > 0 Then
                                Dim dv As Data.DataRowView
                                For Each dv In tblHSales.DefaultView
                                    rowHSales = dv.Row
                                    If ListX.Contains("-") = False Then
                                        strPONo = "-"
                                        ListX.Add("-")
                                    End If
                                    .Tipe_Document_Acuan_Desc = "Hist Sales"
                                    .Invoice_Date = rowHSales.Invoice_Date
                                Next
                            End If
                        ElseIf .Tipe_Document_Acuan = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR Then
                            tblReturn.DefaultView.RowFilter = Nothing
                            tblReturn.DefaultView.RowFilter = "[Primary] = " & .Primary_Document_Acuan
                            If tblReturn.DefaultView.Count > 0 Then
                                Dim dv As Data.DataRowView
                                For Each dv In tblReturn.DefaultView
                                    rowReturn = dv.Row
                                    If ListX.Contains(rowReturn.No_Transaksi & "-" & rowReturn.No_SO) = False Then
                                        If strPONo = "" Then
                                            strPONo &= rowReturn.No_SO
                                        Else
                                            If Not String.IsNullOrEmpty(rowReturn.No_SO) Then
                                                strPONo &= ", " & rowReturn.No_SO
                                            Else
                                                strPONo &= ""
                                            End If
                                        End If
                                        ListX.Add(rowReturn.No_Transaksi & "-" & rowReturn.No_SO)
                                    End If
                                    .Invoice_Date = rowReturn.Tanggal_Transaksi
                                    .Tipe_Document_Acuan_Desc = "Sales Return"
                                Next
                            End If
                        ElseIf .Tipe_Document_Acuan = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR Then
                            tblHReturn.DefaultView.RowFilter = Nothing
                            tblHReturn.DefaultView.RowFilter = "[Primary] = " & .Primary_Document_Acuan
                            If tblHReturn.DefaultView.Count > 0 Then
                                Dim dv As Data.DataRowView
                                For Each dv In tblHReturn.DefaultView
                                    rowHReturn = dv.Row
                                    If ListX.Contains("-") = False Then
                                        strPONo = "-"
                                        ListX.Add("-")
                                    End If
                                    .Invoice_Date = rowHReturn.AOSR_Date
                                    .Tipe_Document_Acuan_Desc = "Hist Sales Return"
                                Next
                            End If
                        ElseIf .Tipe_Document_Acuan = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note OrElse _
                       .Tipe_Document_Acuan = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note Then
                            tblCDNote.DefaultView.RowFilter = Nothing
                            tblCDNote.DefaultView.RowFilter = "[Primary] = " & .Primary_Document_Acuan
                            If tblCDNote.DefaultView.Count > 0 Then
                                Dim dv As Data.DataRowView
                                For Each dv In tblCDNote.DefaultView
                                    rowCDNote = dv.Row
                                    If Not String.IsNullOrEmpty(rowCDNote.No_PO) AndAlso ListX.Contains(rowCDNote.no_transksi & "-" & rowCDNote.No_PO) = False Then
                                        If strPONo = "" Then
                                            strPONo &= rowCDNote.No_PO
                                        Else
                                            If Not String.IsNullOrEmpty(rowCDNote.No_PO) Then
                                                strPONo &= ", " & rowCDNote.No_PO
                                            Else
                                                strPONo &= ""
                                            End If
                                        End If
                                        ListX.Add(rowCDNote.no_transksi & "-" & rowCDNote.No_PO)
                                    End If


                                    If .Tipe_Document_Acuan = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note Then
                                        .Tipe_Document_Acuan_Desc = "Credit Note"
                                    Else
                                        .Tipe_Document_Acuan_Desc = "Debet Note"
                                    End If
                                    .Invoice_Date = rowCDNote.tgl
                                    .document2 = rowCDNote.customer_debet_note
                                Next
                            End If
                        End If


                        If strPONo <> "" Then
                            .PO_Number = strPONo
                        End If
                    End With
                Next
            End If


            'If vtbl.Rows.Count > 0 Then
            '    DTAInvoice.Fill(tblInvoice)
            '    DTAInvoiceDP.Fill(tblInvoiceDP)
            '    DTAReturn.Fill(tblReturn)
            '    DTAHSales.Fill(tblHSales)
            '    DTAHSReturn.Fill(tblHReturn)
            '    DTACDNote.Fill(tblCDNote)

            '    For Each vrow In vtbl.Rows
            '        With vrow
            '            Dim strPONo As String = ""
            '            tblInvoice.DefaultView.RowFilter = Nothing
            '            tblInvoice.DefaultView.RowFilter = "[Primary] = " & .Primary_Document_Acuan
            '            If tblInvoice.DefaultView.Count > 0 Then
            '                Dim dv As Data.DataRowView
            '                For Each dv In tblInvoice.DefaultView
            '                    rowInvoice = dv.Row
            '                    If strPONo = "" Then
            '                        strPONo &= rowInvoice.No_SO
            '                    Else
            '                        If Not String.IsNullOrEmpty(rowInvoice.No_SO) Then
            '                            strPONo &= ", " & rowInvoice.No_SO
            '                        Else
            '                            strPONo &= ""
            '                        End If
            '                    End If
            '                    .Invoice_Date = rowInvoice.Tanggal_Transaksi
            '                Next
            '            End If

            '            tblReturn.DefaultView.RowFilter = Nothing
            '            tblReturn.DefaultView.RowFilter = "[Primary] = " & .Primary_Document_Acuan
            '            If tblReturn.DefaultView.Count > 0 Then
            '                Dim dv As Data.DataRowView
            '                For Each dv In tblReturn.DefaultView
            '                    rowReturn = dv.Row
            '                    If strPONo = "" Then
            '                        strPONo &= rowReturn.No_SO
            '                    Else
            '                        If Not String.IsNullOrEmpty(rowReturn.No_SO) Then
            '                            strPONo &= ", " & rowReturn.No_SO
            '                        Else
            '                            strPONo &= ""
            '                        End If
            '                    End If
            '                    .Invoice_Date = rowReturn.Tanggal_Transaksi
            '                Next
            '            End If

            '            tblHSales.DefaultView.RowFilter = Nothing
            '            tblHSales.DefaultView.RowFilter = "[Primary] = " & .Primary_Document_Acuan
            '            If tblHSales.DefaultView.Count > 0 Then
            '                Dim dv As Data.DataRowView
            '                For Each dv In tblHSales.DefaultView
            '                    rowHSales = dv.Row
            '                    strPONo = "-"
            '                    .Invoice_Date = rowHSales.Invoice_Date
            '                Next
            '            End If

            '            tblHReturn.DefaultView.RowFilter = Nothing
            '            tblHReturn.DefaultView.RowFilter = "[Primary] = " & .Primary_Document_Acuan
            '            If tblHReturn.DefaultView.Count > 0 Then
            '                Dim dv As Data.DataRowView
            '                For Each dv In tblHReturn.DefaultView
            '                    rowHReturn = dv.Row
            '                    strPONo = "-"
            '                    .Invoice_Date = rowHReturn.AOSR_Date
            '                Next
            '            End If

            '            tblCDNote.DefaultView.RowFilter = Nothing
            '            tblCDNote.DefaultView.RowFilter = "[Primary] = " & .Primary_Document_Acuan
            '            If tblCDNote.DefaultView.Count > 0 Then
            '                Dim dv As Data.DataRowView
            '                For Each dv In tblCDNote.DefaultView
            '                    rowCDNote = dv.Row
            '                    If strPONo = "" Then
            '                        strPONo &= rowCDNote.No_PO
            '                    Else
            '                        If Not String.IsNullOrEmpty(rowCDNote.No_PO) Then
            '                            strPONo &= ", " & rowCDNote.No_PO
            '                        Else
            '                            strPONo &= ""
            '                        End If
            '                    End If
            '                    .Invoice_Date = rowCDNote.tgl
            '                Next
            '            End If

            '            If strPONo <> "" Then
            '                .PO_Number = strPONo
            '            End If
            '        End With
            '    Next
            'End If

        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Critical, "Error")
        Finally
            DTAInvoice.Dispose()
            DTAReturn.Dispose()
            DTAHSales.Dispose()
            DTAHSReturn.Dispose()
            DTACDNote.Dispose()
            tblInvoice.Dispose()
            tblReturn.Dispose()
            tblHSales.Dispose()
            tblHReturn.Dispose()
            tblCDNote.Dispose()
        End Try
    End Sub

    Private Sub TambahkanDetailOthers(ByVal vtbl As DseLaporanARPayment.DataTable1DataTable, ByVal vDari As Date, ByVal vSampai As Date)
        Dim DTAO As New DseLaporanARPaymentTableAdapters.tbltrans_jurnal_item_otherTableAdapter
        Dim TblO As New DseLaporanARPayment.tbltrans_jurnal_item_otherDataTable
        Dim RowO As DseLaporanARPayment.tbltrans_jurnal_item_otherRow
        Dim vRow As DseLaporanARPayment.DataTable1Row
        Dim vtblD As New DseLaporanARPayment.DataTable1DataTable
        Dim vrowD As DseLaporanARPayment.DataTable1Row
        Dim ListC As New List(Of Double)
        Try

            For Each DV1 As DataRowView In vtbl.DefaultView
                vRow = DV1.Row
                ''tentukan tipe document
                'If vRow.Tipe_Document_Acuan = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic Then
                '    vRow.Tipe_Document_Acuan_Desc = "Invoice"
                'ElseIf vRow.Tipe_Document_Acuan = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual Then
                '    vRow.Tipe_Document_Acuan_Desc = "Hist Sales"
                'ElseIf vRow.Tipe_Document_Acuan = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR Then
                '    vRow.Tipe_Document_Acuan_Desc = "Sales Return"
                'ElseIf vRow.Tipe_Document_Acuan = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR Then
                '    vRow.Tipe_Document_Acuan_Desc = "Hist Sales Return"
                'ElseIf vRow.Tipe_Document_Acuan = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note Then
                '    vRow.Tipe_Document_Acuan_Desc = "Credit Note"
                'ElseIf vRow.Tipe_Document_Acuan = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note Then
                '    vRow.Tipe_Document_Acuan_Desc = "Debet Note"
                'ElseIf vRow.Tipe_Document_Acuan = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_Transfer_Invoice Then
                '    vRow.Tipe_Document_Acuan_Desc = "Invoice"
                'End If

                vrowD = vtblD.NewRow
                For ij As Integer = 0 To vtblD.Columns.Count - 1
                    vrowD.Item(ij) = vRow.Item(ij)
                Next
                vtblD.Rows.Add(vrowD)
            Next




            Dim ListX As New List(Of String)
            Dim ListA As New List(Of String)
            Dim ListB As New List(Of String)
            If vtblD.Rows.Count > 0 Then
                DTAO.Fill(TblO, vDari, vSampai)
                For Each vrowD In vtblD.Rows
                    With vrowD
                        'tambahkan baris untuk other journal item
                        TblO.DefaultView.RowFilter = Nothing
                        TblO.DefaultView.RowFilter = "[Primary Transaksi] = " & .Primary_Journal
                        Dim dv As Data.DataRowView
                        If TblO.DefaultView.Count > 0 Then
                            For Each dv In TblO.DefaultView
                                RowO = dv.Row
                                If ListX.Contains(.Primary_Journal & "-" & RowO.primary) = False Then
                                    vRow = vtbl.NewRow
                                    vRow.Primary = .Primary
                                    vRow.Primary_Journal = .Primary_Journal
                                    vRow.Payment = RowO.value
                                    vRow.TipePayment = .TipePayment
                                    vRow.Primary_Document_Acuan = .Primary_Document_Acuan
                                    vRow.Tipe_Document_Acuan = 97
                                    vRow.Tipe_Document_Acuan_Desc = ""
                                    vRow.primary_currency = .primary_currency
                                    vRow.exchange_rate = RowO.exchange_rate
                                    vRow.payment_original = RowO.value_original
                                    vRow.primary_akun_ap_ar = RowO.primary_akun
                                    vRow.document1 = .document1
                                    vRow.document2 = .document2
                                    vRow.Ref_Number = .Ref_Number
                                    vRow._Date = ._Date
                                    vRow.Tipe_Journal = .Tipe_Journal
                                    vRow.Primary_Pihak_Acuan = .Primary_Pihak_Acuan
                                    vRow.Desc = .Desc
                                    vRow.primary_akun_pay_to = .primary_akun_pay_to
                                    vRow.Currency_Code = .Currency_Code
                                    vRow.Default_Precision = .Default_Precision
                                    vRow.Customer_Code = .Customer_Code
                                    vRow.Customer_Name = .Customer_Name
                                    vRow.Description = .Description
                                    vRow.Invoice_No_ = .Invoice_No_
                                    vRow.customer_group_code = .customer_group_code
                                    vRow.customer_group_desc = .customer_group_desc
                                    vRow.master_group_code = .master_group_code
                                    vRow.master_group_desc = .master_group_desc
                                    vRow.ar_code = Me.objAcc.GetAccountCode(RowO.primary_akun)
                                    vRow.ar_desc = Me.objAcc.GetDescription_ByPrimaryAcc(RowO.primary_akun)
                                    vRow.payment_code = .payment_code
                                    vRow.Primary_Group_Pihak_Acuan = .Primary_Group_Pihak_Acuan
                                    vRow.Currency_Code_payment = .Currency_Code_payment
                                    vRow.Primary_Currency_Payment = .Primary_Currency_Payment
                                    vRow.difference = .difference
                                    vRow.primary_akun_difference = .primary_akun_difference
                                    vRow.exchange_rate_gain_loss = .exchange_rate_gain_loss
                                    vRow.primary_exhange_gain_loss = .primary_exhange_gain_loss
                                    vRow.Primary_Realized_Exchange_Gain_Loss = .Primary_Realized_Exchange_Gain_Loss
                                    vRow.Primary_Unrealized_Exchange_Gain_Loss_Account = .Primary_Unrealized_Exchange_Gain_Loss_Account
                                    vRow.payment_method = .payment_method
                                    vRow.Invoice_Date = .Invoice_Date
                                    vRow.Hide = .Hide
                                    vRow.DeffPrecIDR = Me.objCurr.GetDecimal_DefaultCurrency
                                    vtbl.Rows.Add(vRow)
                                    ListX.Add(.Primary_Journal & "-" & RowO.primary)
                                End If
                            Next
                        End If

                        'tambahkan baris untuk akun different
                        If .difference <> 0 Then
                            If ListA.Contains(.Primary_Journal) = False Then
                                vRow = vtbl.NewRow
                                vRow.Primary = .Primary
                                vRow.Primary_Journal = .Primary_Journal
                                vRow.Payment = .difference
                                vRow.TipePayment = .TipePayment
                                vRow.Primary_Document_Acuan = .Primary_Document_Acuan
                                vRow.Tipe_Document_Acuan = 98
                                vRow.Tipe_Document_Acuan_Desc = ""
                                vRow.primary_currency = Core_Module.PrimaryBaseCurrency
                                vRow.exchange_rate = 1
                                vRow.payment_original = .difference
                                vRow.primary_akun_ap_ar = .primary_akun_difference
                                vRow.document1 = .document1
                                vRow.document2 = .document2
                                vRow.Ref_Number = .Ref_Number
                                vRow._Date = ._Date
                                vRow.Tipe_Journal = .Tipe_Journal
                                vRow.Primary_Pihak_Acuan = .Primary_Pihak_Acuan
                                vRow.Desc = .Desc
                                vRow.primary_akun_pay_to = .primary_akun_pay_to
                                vRow.Currency_Code = Me.objCurr.GetCurrencyCode(Core_Module.PrimaryBaseCurrency)
                                vRow.Default_Precision = .Default_Precision
                                vRow.Customer_Code = .Customer_Code
                                vRow.Customer_Name = .Customer_Name
                                vRow.Description = .Description
                                vRow.Invoice_No_ = .Invoice_No_
                                vRow.customer_group_code = .customer_group_code
                                vRow.customer_group_desc = .customer_group_desc
                                vRow.master_group_code = .master_group_code
                                vRow.master_group_desc = .master_group_desc
                                vRow.ar_code = Me.objAcc.GetAccountCode(.primary_akun_difference)
                                vRow.ar_desc = Me.objAcc.GetDescription_ByPrimaryAcc(.primary_akun_difference)
                                vRow.payment_code = .payment_code
                                vRow.Primary_Group_Pihak_Acuan = .Primary_Group_Pihak_Acuan
                                vRow.Currency_Code_payment = .Currency_Code_payment
                                vRow.Primary_Currency_Payment = .Primary_Currency_Payment
                                vRow.difference = .difference
                                vRow.primary_akun_difference = .primary_akun_difference
                                vRow.exchange_rate_gain_loss = .exchange_rate_gain_loss
                                vRow.primary_exhange_gain_loss = .primary_exhange_gain_loss
                                vRow.Primary_Realized_Exchange_Gain_Loss = .Primary_Realized_Exchange_Gain_Loss
                                vRow.Primary_Unrealized_Exchange_Gain_Loss_Account = .Primary_Unrealized_Exchange_Gain_Loss_Account
                                vRow.payment_method = .payment_method
                                vRow.Invoice_Date = .Invoice_Date
                                vRow.Hide = .Hide
                                vRow.DeffPrecIDR = Me.objCurr.GetDecimal_DefaultCurrency
                                vtbl.Rows.Add(vRow)
                                ListA.Add(.Primary_Journal)
                            End If
                        End If


                        'tambahkan baris untuk akun different
                        If .exchange_rate_gain_loss <> 0 Then
                            If ListB.Contains(.Primary_Journal) = False Then
                                vRow = vtbl.NewRow
                                vRow.Primary = .Primary
                                vRow.Primary_Journal = .Primary_Journal
                                vRow.TipePayment = .TipePayment
                                vRow.Primary_Document_Acuan = .Primary_Document_Acuan
                                vRow.Tipe_Document_Acuan = 99
                                vRow.Tipe_Document_Acuan_Desc = ""
                                vRow.primary_currency = Core_Module.PrimaryBaseCurrency
                                vRow.primary_akun_ap_ar = .primary_akun_difference
                                vRow.document1 = .document1
                                vRow.document2 = .document2
                                vRow.Ref_Number = .Ref_Number
                                vRow._Date = ._Date
                                vRow.Tipe_Journal = .Tipe_Journal
                                vRow.Primary_Pihak_Acuan = .Primary_Pihak_Acuan
                                vRow.Desc = .Desc
                                vRow.primary_akun_pay_to = .primary_akun_pay_to
                                vRow.Currency_Code = Me.objCurr.GetCurrencyCode(Core_Module.PrimaryBaseCurrency)
                                vRow.Default_Precision = .Default_Precision
                                vRow.Customer_Code = .Customer_Code
                                vRow.Customer_Name = .Customer_Name
                                vRow.Description = .Description
                                vRow.Invoice_No_ = .Invoice_No_
                                vRow.customer_group_code = .customer_group_code
                                vRow.customer_group_desc = .customer_group_desc
                                vRow.master_group_code = .master_group_code
                                vRow.master_group_desc = .master_group_desc
                                vRow.Invoice_Date = .Invoice_Date
                                vRow.exchange_rate = 1
                                If .primary_exhange_gain_loss > -1 Then
                                    vRow.ar_code = Me.objAcc.GetAccountCode(.primary_exhange_gain_loss)
                                    vRow.ar_desc = Me.objAcc.GetDescription_ByPrimaryAcc(.primary_exhange_gain_loss)
                                    vRow.Payment = -(.exchange_rate_gain_loss)
                                    vRow.payment_original = -(.exchange_rate_gain_loss)
                                Else
                                    vRow.ar_code = Me.objAcc.GetAccountCode(.Primary_Unrealized_Exchange_Gain_Loss_Account)
                                    vRow.ar_desc = Me.objAcc.GetDescription_ByPrimaryAcc(.Primary_Unrealized_Exchange_Gain_Loss_Account)
                                    vRow.Payment = -(.exchange_rate_gain_loss)
                                    vRow.payment_original = -(.exchange_rate_gain_loss)
                                    'If .exchange_rate_gain_loss < 0 Then
                                    '    vRow.ar_code = Me.objAcc.GetAccountCode(.Primary_Realized_Exchange_Gain_Loss)
                                    '    vRow.ar_desc = Me.objAcc.GetDescription_ByPrimaryAcc(.Primary_Realized_Exchange_Gain_Loss)
                                    '    vRow.Payment = -(.exchange_rate_gain_loss)
                                    '    vRow.payment_original = -(.exchange_rate_gain_loss)
                                    'Else
                                    '    vRow.ar_code = Me.objAcc.GetAccountCode(.Primary_Unrealized_Exchange_Gain_Loss_Account)
                                    '    vRow.ar_desc = Me.objAcc.GetDescription_ByPrimaryAcc(.Primary_Unrealized_Exchange_Gain_Loss_Account)
                                    '    vRow.Payment = .exchange_rate_gain_loss
                                    '    vRow.payment_original = -(.exchange_rate_gain_loss)
                                    'End If
                                End If
                                vRow.payment_code = .payment_code
                                vRow.Primary_Group_Pihak_Acuan = .Primary_Group_Pihak_Acuan
                                vRow.Currency_Code_payment = .Currency_Code_payment
                                vRow.Primary_Currency_Payment = .Primary_Currency_Payment
                                vRow.difference = .difference
                                vRow.primary_akun_difference = .primary_akun_difference
                                vRow.exchange_rate_gain_loss = .exchange_rate_gain_loss
                                vRow.primary_exhange_gain_loss = .primary_exhange_gain_loss
                                vRow.Primary_Realized_Exchange_Gain_Loss = .Primary_Realized_Exchange_Gain_Loss
                                vRow.Primary_Unrealized_Exchange_Gain_Loss_Account = .Primary_Unrealized_Exchange_Gain_Loss_Account
                                vRow.payment_method = .payment_method
                                vRow.Invoice_Date = .Invoice_Date
                                vRow.Hide = .Hide
                                vRow.DeffPrecIDR = Me.objCurr.GetDecimal_DefaultCurrency
                                vtbl.Rows.Add(vRow)
                                ListB.Add(.Primary_Journal)
                            End If
                        End If
                    End With
                Next
            End If


            vtbl.DefaultView.Sort = " [Customer Code], Date, [Ref Number], [Tipe Document Acuan], [Invoice Date], [Document1] "
            Dim decPayment As Decimal = 0
            For Each DV2 As DataRowView In vtbl.DefaultView
                vRow = DV2.Row
                If ListC.Contains(vRow.Primary_Pihak_Acuan) = False Then
                    vRow.Hide = 1
                    ListC.Add(vRow.Primary_Pihak_Acuan)
                End If
            Next


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAO.Dispose()
            TblO.Dispose()
            vtblD.Dispose()
        End Try
    End Sub

    Private Sub UrutkanUlang(ByRef vtbl As DseLaporanARPayment.DataTable1DataTable)
        Dim vtbl_dummy As New DseLaporanARPayment.DataTable1DataTable
        Dim row, rowDummy As DseLaporanARPayment.DataTable1Row

        Try
            vtbl_dummy.Rows.Clear()
            'diurutkan berdasarkan group code, payment date, journal no, invoice number 
            If Me.rboDetail.Checked = False Then
                vtbl.DefaultView.Sort = "[master_group_code], Date, [Ref Number], [Document1]"
                vtbl.DefaultView.RowFilter = ""
            Else
                vtbl.DefaultView.Sort = "[Customer Code], Date, [Ref Number], [Tipe Document Acuan], [Invoice Date], [Document1] "
                vtbl.DefaultView.RowFilter = ""
            End If

            For Each DV As DataRowView In vtbl.DefaultView
                row = DV.Row
                rowDummy = vtbl_dummy.NewDataTable1Row
                For ij As Integer = 0 To vtbl.Columns.Count - 1
                    rowDummy.Item(ij) = row.Item(ij)
                Next
                vtbl_dummy.Rows.Add(rowDummy)
            Next

            vtbl.Rows.Clear()
            For Each rowDummy In vtbl_dummy.Rows
                row = vtbl.NewDataTable1Row
                For ij As Integer = 0 To vtbl.Columns.Count - 1
                    row.Item(ij) = rowDummy.Item(ij)
                Next
                vtbl.Rows.Add(row)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetGroup(ByRef vtbl As DseLaporanARPayment.DataTable1DataTable)
        Dim vrow As DseLaporanARPayment.DataTable1Row
        Try
            For Each vrow In vtbl
                With vrow
                    If rboCustCode.Checked = True Then
                        .master_group_code = .Customer_Code
                        .master_group_desc = " - " & .Customer_Name
                    ElseIf rboCustGroup.Checked = True Then
                        .master_group_code = .customer_group_code
                        .master_group_desc = " - " & .customer_group_desc
                    ElseIf rboPaymentAccount.Checked = True Then
                        .master_group_code = .payment_code
                        .master_group_desc = " - " & .Description
                    ElseIf rboArAccount.Checked = True Then
                        .master_group_code = .ar_code
                        .master_group_desc = " - " & .ar_desc
                    ElseIf rboCurrency.Checked = True Then
                        .master_group_code = .Currency_Code
                        .master_group_desc = ""
                    End If
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    'Private Sub SetInvoiceDate(ByRef vtbl As DseLaporanARPayment.DataTable1DataTable)
    '    Dim vrow As DseLaporanARPayment.DataTable1Row

    '    Dim DTAInvoice As New DseLaporanARPaymentTableAdapters.tbltrans_salesinvoiceTableAdapter
    '    Dim DTAReturn As New DseLaporanARPaymentTableAdapters.tbltrans_salesreturnTableAdapter
    '    Dim DTAHSales As New DseLaporanARPaymentTableAdapters.tbltrans_historicalsalesTableAdapter
    '    Dim DTAHSReturn As New DseLaporanARPaymentTableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter
    '    Dim DTACDNote As New DseLaporanARPaymentTableAdapters.tbltrans_credit_noteTableAdapter

    '    Dim tblInvoice As New DseLaporanARPayment.tbltrans_salesinvoiceDataTable
    '    Dim tblReturn As New DseLaporanARPayment.tbltrans_salesreturnDataTable
    '    Dim tblHSales As New DseLaporanARPayment.tbltrans_historicalsalesDataTable
    '    Dim tblHReturn As New DseLaporanARPayment.tbltrans_historicalsalesreturn_manualDataTable
    '    Dim tblCDNote As New DseLaporanARPayment.tbltrans_credit_noteDataTable

    '    Dim rowInvoice As DseLaporanARPayment.tbltrans_salesinvoiceRow
    '    Dim rowReturn As DseLaporanARPayment.tbltrans_salesreturnRow
    '    Dim rowHSales As DseLaporanARPayment.tbltrans_historicalsalesRow
    '    Dim rowHReturn As DseLaporanARPayment.tbltrans_historicalsalesreturn_manualRow
    '    Dim rowCDNote As DseLaporanARPayment.tbltrans_credit_noteRow


    '    Try
    '        If vtbl.Rows.Count > 0 Then
    '            DTAInvoice.Fill(tblInvoice)
    '            DTAReturn.Fill(tblReturn)
    '            DTAHSales.Fill(tblHSales)
    '            DTAHSReturn.Fill(tblHReturn)
    '            DTACDNote.Fill(tblCDNote)

    '            For Each vrow In vtbl.Rows
    '                With vrow
    '                    tblInvoice.DefaultView.RowFilter = Nothing
    '                    tblInvoice.DefaultView.RowFilter = "[Primary] = " & .Primary_Document_Acuan
    '                    If tblInvoice.DefaultView.Count > 0 Then
    '                        rowInvoice = tblInvoice.DefaultView(0).Row
    '                        .Invoice_Date = rowInvoice.Tanggal_Transaksi
    '                    End If

    '                    tblReturn.DefaultView.RowFilter = Nothing
    '                    tblReturn.DefaultView.RowFilter = "[Primary] = " & .Primary_Document_Acuan
    '                    If tblReturn.DefaultView.Count > 0 Then
    '                        rowReturn = tblReturn.DefaultView(0).Row
    '                        .Invoice_Date = rowReturn.Tanggal_Transaksi
    '                    End If

    '                    tblHSales.DefaultView.RowFilter = Nothing
    '                    tblHSales.DefaultView.RowFilter = "[Primary] = " & .Primary_Document_Acuan
    '                    If tblHSales.DefaultView.Count > 0 Then
    '                        rowHSales = tblHSales.DefaultView(0).Row
    '                        .Invoice_Date = rowHSales.Invoice_Date
    '                    End If

    '                    tblHReturn.DefaultView.RowFilter = Nothing
    '                    tblHReturn.DefaultView.RowFilter = "[Primary] = " & .Primary_Document_Acuan
    '                    If tblHReturn.DefaultView.Count > 0 Then
    '                        rowHReturn = tblHReturn.DefaultView(0).Row
    '                        .Invoice_Date = rowHReturn.AOSR_Date
    '                    End If

    '                    tblCDNote.DefaultView.RowFilter = Nothing
    '                    tblCDNote.DefaultView.RowFilter = "[Primary] = " & .Primary_Document_Acuan
    '                    If tblCDNote.DefaultView.Count > 0 Then
    '                        rowCDNote = tblCDNote.DefaultView(0).Row
    '                        .Invoice_Date = rowCDNote.tgl
    '                    End If
    '                End With
    '            Next
    '        End If

    '    Catch ex As Exception
    '        MsgBox(ex.Message.ToString, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTAInvoice.Dispose()
    '        DTAReturn.Dispose()
    '        DTAHSales.Dispose()
    '        DTAHSReturn.Dispose()
    '        DTACDNote.Dispose()
    '        tblInvoice.Dispose()
    '        tblReturn.Dispose()
    '        tblHSales.Dispose()
    '        tblHReturn.Dispose()
    '        tblCDNote.Dispose()
    '    End Try
    'End Sub

    Private Sub FrmCetakAR_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.dtpFrom.Value = Now
            Me.dtpTo.Value = Now
            Me.crvAR.ReportSource = Nothing
            Me.crvAR.Zoom(85)
            Me.crvAR.Refresh()
            Me.crvAR.Show()
            Me.Text = "AR Payment Report"
            Me.btnFindArAccount.TabStop = False
            Me.btnFindCurrency.TabStop = False
            Me.btnFindCustomerCode.TabStop = False
            Me.btnFindCustomerGroup.TabStop = False
            Me.btnFindPaymentAccount.TabStop = False
            Me.btnShow.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindCustomerCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCustomerCode.Click
        Dim frm As New Core_Module.FrmCariMasterCustomer
        Try
            If Me.txtCustomerCode.Text <> "" Then
                frm.txtCari.Text = Me.txtCustomerCode.Text
            End If
            frm.ShowDialog()
            If frm.PrimaryCustomer > 0 Then
                Me.txtPrimaryCustomerCode.Text = frm.PrimaryCustomer
                Me.txtCustomerCode.Text = Me.objCust.GetCustomerCode(frm.PrimaryCustomer)
                Me.txtCustomerDesc.Text = Me.objCust.GetCustName(frm.PrimaryCustomer)
            Else
                Me.txtPrimaryCustomerCode.Text = "-1"
                Me.txtCustomerCode.Clear()
                Me.txtCustomerDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            frm = Nothing
        End Try
    End Sub

    Private Sub txtCustomerCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown
        Try
            If e.KeyCode = 13 Then
                If Me.txtCustomerCode.Text <> "" Then
                    If Me.objCust.ApakahSudahAda(Me.txtCustomerCode.Text) = True Then
                        Me.txtPrimaryCustomerCode.Text = Me.objCust.GetPrimary(Me.txtCustomerCode.Text)
                        Me.txtCustomerCode.Text = Me.objCust.GetCustomerCode(Me.txtPrimaryCustomerCode.Text)
                        Me.txtCustomerDesc.Text = Me.objCust.GetCustName(Me.txtPrimaryCustomerCode.Text)
                        Me.btnFindCustomerCode.Focus()
                    Else
                        Me.btnFindCustomerCode.PerformClick()
                    End If
                Else
                    Me.txtPrimaryCustomerCode.Text = "-1"
                    Me.txtCustomerCode.Clear()
                    Me.txtCustomerDesc.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub txtCustomerCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.Leave
        Try
            If Me.txtCustomerCode.Text <> "" Then
                If Me.objCust.ApakahSudahAda(Me.txtCustomerCode.Text) = True Then
                    Me.txtPrimaryCustomerCode.Text = Me.objCust.GetPrimary(Me.txtCustomerCode.Text)
                    Me.txtCustomerCode.Text = Me.objCust.GetCustomerCode(Me.txtPrimaryCustomerCode.Text)
                    Me.txtCustomerDesc.Text = Me.objCust.GetCustName(Me.txtPrimaryCustomerCode.Text)
                    Me.btnFindCustomerCode.Focus()
                Else
                    Me.btnFindCustomerCode.PerformClick()
                End If
            Else
                Me.txtPrimaryCustomerCode.Text = "-1"
                Me.txtCustomerCode.Clear()
                Me.txtCustomerDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub btnFindCustomerGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCustomerGroup.Click
        Dim frm As New Core_Module.FrmCariMasterCustomerGroup
        Try
            If Me.txtCustomerGroupCode.Text <> "" Then
                frm.txtCari.Text = Me.txtCustomerGroupCode.Text
            End If
            frm.ShowDialog()
            If frm.PrimaryCustomerGroup > 0 Then
                Me.txtPrimaryCustomerGroup.Text = frm.PrimaryCustomerGroup
                Me.txtCustomerGroupCode.Text = Me.objCustGroup.GetGroupCode(frm.PrimaryCustomerGroup)
                Me.objCustGroup.CariDetail(frm.PrimaryCustomerGroup)
                Me.txtCustomerGroupDesc.Text = Me.objCustGroup.Description
            Else
                Me.txtPrimaryCustomerGroup.Text = "-1"
                Me.txtCustomerGroupCode.Clear()
                Me.txtCustomerGroupDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            frm = Nothing
        End Try
    End Sub

    Private Sub txtCustomerGroupCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomerGroupCode.KeyDown
        Try
            If e.KeyCode = 13 Then
                If Me.txtCustomerGroupCode.Text <> "" Then
                    If Me.objCustGroup.ApakahSudahAda(Me.txtCustomerGroupCode.Text) = True Then
                        Me.txtPrimaryCustomerGroup.Text = Me.objCustGroup.GetPrimary(Me.txtCustomerGroupCode.Text)
                        Me.txtCustomerGroupCode.Text = Me.objCustGroup.GetGroupCode(Me.txtPrimaryCustomerGroup.Text)
                        Me.objCustGroup.CariDetail(Me.txtPrimaryCustomerGroup.Text)
                        Me.txtCustomerGroupDesc.Text = Me.objCustGroup.Description
                        Me.btnFindCustomerGroup.Focus()
                    Else
                        Me.btnFindCustomerGroup.PerformClick()
                    End If
                Else
                    Me.txtPrimaryCustomerGroup.Text = "-1"
                    Me.txtCustomerGroupCode.Clear()
                    Me.txtCustomerGroupDesc.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub txtCustomerGroupCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerGroupCode.Leave
        Try
            If Me.txtCustomerGroupCode.Text <> "" Then
                If Me.objCustGroup.ApakahSudahAda(Me.txtCustomerGroupCode.Text) = True Then
                    Me.txtPrimaryCustomerGroup.Text = Me.objCustGroup.GetPrimary(Me.txtCustomerGroupCode.Text)
                    Me.txtCustomerGroupCode.Text = Me.objCustGroup.GetGroupCode(Me.txtPrimaryCustomerGroup.Text)
                    Me.objCustGroup.CariDetail(Me.txtPrimaryCustomerGroup.Text)
                    Me.txtCustomerGroupDesc.Text = Me.objCustGroup.Description
                    Me.btnFindCustomerGroup.Focus()
                Else
                    Me.btnFindCustomerGroup.PerformClick()
                End If
            Else
                Me.txtPrimaryCustomerGroup.Text = "-1"
                Me.txtCustomerGroupCode.Clear()
                Me.txtCustomerGroupDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub btnFindPaymentAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindPaymentAccount.Click
        Dim frm As New Core_Module.FrmMasterAkun_Cari1
        Try
            If Me.txtPaymentAccountCode.Text <> "" Then
                frm.txtCari.Text = Me.txtPaymentAccountCode.Text
            End If
            frm.ShowDialog()
            If frm.mPrimaryAkun > 0 Then
                Me.txtPrimaryPaymentAccount.Text = frm.mPrimaryAkun
                Me.txtPaymentAccountCode.Text = Me.objAcc.GetAccountCode(frm.mPrimaryAkun)
                Me.txtPaymentAccountDesc.Text = Me.objAcc.GetDescription_ByPrimaryAcc(frm.mPrimaryAkun)
            Else
                Me.txtPrimaryPaymentAccount.Text = "-1"
                Me.txtPaymentAccountCode.Clear()
                Me.txtPaymentAccountDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            frm = Nothing
        End Try
    End Sub

    Private Sub txtPaymentAccountCode_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPaymentAccountCode.KeyDown
        Try
            If e.KeyCode = 13 Then
                If (Me.txtPaymentAccountCode.Text <> "") Then
                    If Me.objAcc.ApakahSudahAda_Aktif_NonHeader(Me.txtPaymentAccountCode.Text) = True Then
                        Me.txtPrimaryPaymentAccount.Text = Me.objAcc.GetPrimary(Me.txtPaymentAccountCode.Text).ToString
                        Me.txtPaymentAccountCode.Text = Me.objAcc.GetAccountCode(Me.txtPrimaryPaymentAccount.Text)
                        Me.txtPaymentAccountDesc.Text = Me.objAcc.GetDescription_ByPrimaryAcc(Me.txtPrimaryPaymentAccount.Text)
                        Me.btnFindPaymentAccount.Focus()
                    Else
                        Me.btnFindPaymentAccount.PerformClick()
                    End If
                Else
                    Me.txtPrimaryPaymentAccount.Text = "-1"
                    Me.txtPaymentAccountCode.Clear()
                    Me.txtPaymentAccountDesc.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub txtPaymentAccountCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPaymentAccountCode.Leave
        Try
            If (Me.txtPaymentAccountCode.Text <> "") Then
                If Me.objAcc.ApakahSudahAda_Aktif_NonHeader(Me.txtPaymentAccountCode.Text) = True Then
                    Me.txtPrimaryPaymentAccount.Text = Me.objAcc.GetPrimary(Me.txtPaymentAccountCode.Text).ToString
                    Me.txtPaymentAccountCode.Text = Me.objAcc.GetAccountCode(Me.txtPrimaryPaymentAccount.Text)
                    Me.txtPaymentAccountDesc.Text = Me.objAcc.GetDescription_ByPrimaryAcc(Me.txtPrimaryPaymentAccount.Text)
                    Me.btnFindPaymentAccount.Focus()
                Else
                    Me.btnFindPaymentAccount.PerformClick()
                End If
            Else
                Me.txtPrimaryPaymentAccount.Text = "-1"
                Me.txtPaymentAccountCode.Clear()
                Me.txtPaymentAccountDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub btnFindArAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindArAccount.Click
        Dim frm As New Core_Module.FrmMasterAkun_Cari1
        Try
            If Me.txtArAccountCode.Text <> "" Then
                frm.txtCari.Text = Me.txtArAccountCode.Text
            End If
            frm.ShowDialog()
            If frm.mPrimaryAkun > 0 Then
                Me.txtPrimaryArAccount.Text = frm.mPrimaryAkun
                Me.txtArAccountCode.Text = Me.objAcc.GetAccountCode(frm.mPrimaryAkun)
                Me.txtArAccountDesc.Text = Me.objAcc.GetDescription_ByPrimaryAcc(frm.mPrimaryAkun)
            Else
                Me.txtPrimaryArAccount.Text = "-1"
                Me.txtArAccountCode.Clear()
                Me.txtArAccountDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            frm = Nothing
        End Try
    End Sub

    Private Sub txtArAccountCode_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtArAccountCode.KeyDown
        Try
            If e.KeyCode = 13 Then
                If (Me.txtArAccountCode.Text <> "") Then
                    If Me.objAcc.ApakahSudahAda_Aktif_NonHeader(Me.txtArAccountCode.Text) = True Then
                        Me.txtPrimaryArAccount.Text = Me.objAcc.GetPrimary(Me.txtArAccountCode.Text).ToString
                        Me.txtArAccountCode.Text = Me.objAcc.GetAccountCode(Me.txtPrimaryArAccount.Text)
                        Me.txtArAccountDesc.Text = Me.objAcc.GetDescription_ByPrimaryAcc(Me.txtPrimaryArAccount.Text)
                        Me.btnFindArAccount.Focus()
                    Else
                        Me.btnFindArAccount.PerformClick()
                    End If
                Else
                    Me.txtPrimaryArAccount.Text = "-1"
                    Me.txtArAccountCode.Clear()
                    Me.txtArAccountDesc.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub txtArAccountCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtArAccountCode.Leave
        Try
            If (Me.txtArAccountCode.Text <> "") Then
                If Me.objAcc.ApakahSudahAda_Aktif_NonHeader(Me.txtArAccountCode.Text) = True Then
                    Me.txtPrimaryArAccount.Text = Me.objAcc.GetPrimary(Me.txtArAccountCode.Text).ToString
                    Me.txtArAccountCode.Text = Me.objAcc.GetAccountCode(Me.txtPrimaryArAccount.Text)
                    Me.txtArAccountDesc.Text = Me.objAcc.GetDescription_ByPrimaryAcc(Me.txtPrimaryArAccount.Text)
                    Me.btnFindArAccount.Focus()
                Else
                    Me.btnFindArAccount.PerformClick()
                End If
            Else
                Me.txtPrimaryArAccount.Text = "-1"
                Me.txtArAccountCode.Clear()
                Me.txtArAccountDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub btnFindCurrency_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCurrency.Click
        Dim frm As New Core_Module.FrmCariMasterCurrency
        Try
            If Me.txtCurrencyCode.Text <> "" Then
                frm.txtCari.Text = Me.txtCurrencyCode.Text
            End If
            frm.ShowDialog()
            If frm.PrimaryCurrency > 0 Then
                Me.txtPrimaryCurrency.Text = frm.PrimaryCurrency
                Me.txtCurrencyCode.Text = Me.objCurr.GetCurrencyCode(frm.PrimaryCurrency)
                Me.objCurr.CariDetail(frm.PrimaryCurrency)
                Me.txtCurrencyDesc.Text = Me.objCurr.Description
            Else
                Me.txtPrimaryCurrency.Text = "-1"
                Me.txtCurrencyCode.Clear()
                Me.txtCurrencyDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            frm = Nothing
        End Try
    End Sub

    Private Sub txtCurrencyCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCurrencyCode.KeyDown
        Try
            If e.KeyCode = 13 Then
                If Me.txtCurrencyCode.Text <> "" Then
                    If Me.objCurr.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                        Me.txtPrimaryCurrency.Text = Me.objCurr.GetPrimary(Me.txtCurrencyCode.Text)
                        Me.txtCurrencyCode.Text = Me.objCurr.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                        Me.objCurr.CariDetail(Me.txtPrimaryCurrency.Text)
                        Me.txtCurrencyDesc.Text = Me.objCurr.Description
                        Me.btnFindCurrency.Focus()
                    Else
                        Me.btnFindCurrency.PerformClick()
                    End If
                Else
                    Me.txtPrimaryCurrency.Text = "-1"
                    Me.txtCurrencyCode.Clear()
                    Me.txtCurrencyDesc.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub txtCurrencyCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.Leave
        Try
            If Me.txtCurrencyCode.Text <> "" Then
                If Me.objCurr.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                    Me.txtPrimaryCurrency.Text = Me.objCurr.GetPrimary(Me.txtCurrencyCode.Text)
                    Me.txtCurrencyCode.Text = Me.objCurr.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                    Me.objCurr.CariDetail(Me.txtPrimaryCurrency.Text)
                    Me.txtCurrencyDesc.Text = Me.objCurr.Description
                    Me.btnFindCurrency.Focus()
                Else
                    Me.btnFindCurrency.PerformClick()
                End If
            Else
                Me.txtPrimaryCurrency.Text = "-1"
                Me.txtCurrencyCode.Clear()
                Me.txtCurrencyDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub btnClose_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

End Class