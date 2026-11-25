Option Explicit On
Option Strict Off
Public Class FrmLaporanSalesDP

    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim objCD As New Core_Module.ClsCekData
    Dim ObjSalesP As New Core_Module.ClsMasterSalesPerson
    Dim buffer As String = ""

    Private Sub TxtCustomerCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerCode.GotFocus
        TxtCustomerCode.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtCustomerCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerCode.LostFocus
        Me.TxtCustomerCode.BackColor = Me.TxtPrimaryCustomer.BackColor
    End Sub

    Private Sub TxtCustomerCode_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCustomerCode.KeyDown
        Dim ObjCust As New Core_Module.clsMasterCustomer
        Try
            If e.KeyCode = Keys.Enter Then
                Dim primary As Double = ObjCust.GetPrimary(Me.TxtCustomerCode.Text)
                If primary <> -1 Then
                    Me.TxtPrimaryCustomer.Text = primary
                    Me.TxtCustomerCode.Text = ObjCust.GetCustomerCode(primary)
                    Me.TxtCustomerName.Text = ObjCust.GetCustName(primary)
                Else
                    Me.BtnFindCustomer.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCust = Nothing
        End Try
    End Sub

    Private Sub TxtCustomerCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerCode.Leave
        Dim ObjCust As New Core_Module.clsMasterCustomer
        Try
            If Me.TxtCustomerCode.Text <> "" Then
                Dim primary As Double = ObjCust.GetPrimary(Me.TxtCustomerCode.Text)
                If primary <> -1 Then
                    Me.TxtPrimaryCustomer.Text = primary
                    Me.TxtCustomerCode.Text = ObjCust.GetCustomerCode(primary)
                    Me.TxtCustomerName.Text = ObjCust.GetCustName(primary)
                Else
                    Me.BtnFindCustomer.PerformClick()
                End If
            Else
                Me.TxtPrimaryCustomer.Text = "-1"
                Me.TxtCustomerName.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCust = Nothing
        End Try
    End Sub

    Private Sub BtnFindCustomer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCustomer.Click
        Dim i As New Core_Module.FrmCariMasterCustomer
        Try
            i.BackColorAll = Me.BackColorAll
            i.txtCari.Text = Me.TxtCustomerCode.Text
            i.ShowDialog()
            If i.PrimaryCustomer <> -1 Then
                Dim objCust As New Core_Module.clsMasterCustomer
                Me.TxtPrimaryCustomer.Text = i.PrimaryCustomer
                Me.TxtCustomerCode.Text = objCust.GetCustomerCode(i.PrimaryCustomer)
                Me.TxtCustomerName.Text = objCust.GetCustName(i.PrimaryCustomer)
                objCust = Nothing
            Else
                Me.TxtPrimaryCustomer.Text = "-1"
                Me.TxtCustomerCode.Text = ""
                Me.TxtCustomerName.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtSalesPersonCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesPersonCode.GotFocus
        TxtSalesPersonCode.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtSalesPersonCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesPersonCode.LostFocus
        Me.TxtSalesPersonCode.BackColor = Me.TxtPrimaryCustomer.BackColor
    End Sub

    Private Sub TxtSalesPersonCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtSalesPersonCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.ObjSalesP.ApakahSudahAda(Me.TxtSalesPersonCode.Text) = True Then
                    Me.TxtPrimarySalesPerson.Text = Me.ObjSalesP.GetPrimary(Me.TxtSalesPersonCode.Text)
                    Me.TxtSalesPersonDesc.Text = Me.ObjSalesP.GetName(Me.TxtPrimarySalesPerson.Text)
                Else
                    Me.BtnFindSalesPerson.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesPersonCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesPersonCode.Leave
        Try
            If Me.TxtSalesPersonCode.Text <> "" Then
                If Me.ObjSalesP.ApakahSudahAda(Me.TxtSalesPersonCode.Text) = True Then
                    Me.TxtPrimarySalesPerson.Text = Me.ObjSalesP.GetPrimary(Me.TxtSalesPersonCode.Text)
                    Me.TxtSalesPersonDesc.Text = Me.ObjSalesP.GetName(Me.TxtPrimarySalesPerson.Text)
                Else
                    Me.BtnFindSalesPerson.PerformClick()
                End If
            Else
                Me.TxtPrimarySalesPerson.Text = -1
                Me.TxtSalesPersonCode.Text = ""
                Me.TxtSalesPersonDesc.Text = ""
            End If 
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindSalesPerson_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnFindSalesPerson.Click
        Dim FrmFind As New Core_Module.FrmMasterSalesPerson_Cari
        Try
            FrmFind.BackColorAll = Me.BackColorAll
            FrmFind.txtCari.Text = Me.TxtSalesPersonCode.Text
            FrmFind.ShowDialog()
            If CDbl(FrmFind.PrimarySalesPerson) > 0 Then
                Me.TxtPrimarySalesPerson.Text = FrmFind.PrimarySalesPerson
                Me.TxtSalesPersonCode.Text = Me.ObjSalesP.GetCode(FrmFind.PrimarySalesPerson)
                Me.TxtSalesPersonDesc.Text = Me.ObjSalesP.GetName(FrmFind.PrimarySalesPerson)
            Else
                Me.TxtPrimarySalesPerson.Text = -1
                Me.TxtSalesPersonCode.Text = ""
                Me.TxtSalesPersonDesc.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error") 
        End Try
    End Sub

    Private Sub FrmLaporanSalesDP_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Try
            Me.CboDPStatus.SelectedIndex = 0
            Me.CboGroup.SelectedIndex = 0
            Me.CboPaymentStatus.SelectedIndex = 0
            Me.CboRealizationStatus.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RboSO_Date_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RboSO_Date.CheckedChanged
        Try
            Me.RefreshDtp()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RboDP_Date_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RboDP_Date.CheckedChanged
        Try
            Me.RefreshDtp()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RboInvoiceDate_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RboInvoiceDate.CheckedChanged
        Try
            Me.RefreshDtp()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub RefreshDtp()
        Try
            If Me.RboSO_Date.Checked Then
                Me.DtpSO_DateFrom.Enabled = True
                Me.DtpSO_DateTo.Enabled = True
                Me.DtpDP_DateFrom.Enabled = False
                Me.DtpDP_DateTo.Enabled = False
                Me.DtpInvoice_DateFrom.Enabled = False
                Me.DtpInvoice_DateTo.Enabled = False
            ElseIf Me.RboDP_Date.Checked Then
                Me.DtpSO_DateFrom.Enabled = False
                Me.DtpSO_DateTo.Enabled = False
                Me.DtpDP_DateFrom.Enabled = True
                Me.DtpDP_DateTo.Enabled = True
                Me.DtpInvoice_DateFrom.Enabled = False
                Me.DtpInvoice_DateTo.Enabled = False
            ElseIf Me.RboInvoiceDate.Checked Then
                Me.DtpSO_DateFrom.Enabled = False
                Me.DtpSO_DateTo.Enabled = False
                Me.DtpDP_DateFrom.Enabled = False
                Me.DtpDP_DateTo.Enabled = False
                Me.DtpInvoice_DateFrom.Enabled = True
                Me.DtpInvoice_DateTo.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFilter.Click
        Try
            If Me.DtpSO_DateTo.Value.Date < Me.DtpSO_DateFrom.Value.Date And Me.RboSO_Date.Checked Then
                MsgBox("Invalid SO Date filter.", MsgBoxStyle.Information, "Information")
                Return
            End If
            If Me.DtpDP_DateTo.Value.Date < Me.DtpDP_DateFrom.Value.Date And Me.RboDP_Date.Checked Then
                MsgBox("Invalid DP Date filter.", MsgBoxStyle.Information, "Information")
                Return
            End If
            If Me.DtpInvoice_DateTo.Value.Date < Me.DtpInvoice_DateFrom.Value.Date And Me.RboInvoiceDate.Checked Then
                MsgBox("Invalid Invoice Date filter.", MsgBoxStyle.Information, "Information")
                Return
            End If
            Me.LoadReport()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub LoadReport()
        Dim DS As New DseLaporanSalesDP
        Dim TblPrint As DseLaporanSalesDP.PrintDataDataTable = DS.PrintData
        Dim RowPrint As DseLaporanSalesDP.PrintDataRow

        Dim DtaDP As New DseLaporanSalesDPTableAdapters.DownPaymentDataTableAdapter
        Dim TblDP As New DseLaporanSalesDP.DownPaymentDataDataTable
        Dim RowDP As DseLaporanSalesDP.DownPaymentDataRow
        Dim DtaSO As New DseLaporanSalesDPTableAdapters.SalesOrderDataTableAdapter
        Dim TblSO As New DseLaporanSalesDP.SalesOrderDataDataTable
        Dim RowSO As DseLaporanSalesDP.SalesOrderDataRow
        Dim DtaJournal As New DseLaporanSalesDPTableAdapters.JournalDataTableAdapter
        Dim TblJournal As New DseLaporanSalesDP.JournalDataDataTable
        Dim RowJournal As DseLaporanSalesDP.JournalDataRow
        Dim DtaInvoice As New DseLaporanSalesDPTableAdapters.InvoiceDataTableAdapter
        Dim TblInvoice As New DseLaporanSalesDP.InvoiceDataDataTable
        Dim RowInvoice As DseLaporanSalesDP.InvoiceDataRow

        Dim NullablePrimaryCustomer As New Nullable(Of Decimal)
        Dim NullablePrimarySalesPerson As New Nullable(Of Decimal)
        Dim SO_DateFrom As Date = #1/1/1950#
        Dim SO_DateTo As Date = #1/1/2200#
        Dim DP_DateFrom As Date = #1/1/1950#
        Dim DP_DateTo As Date = #1/1/2200#
        Dim Invoice_DateFrom As Date = #1/1/1950#
        Dim Invoice_DateTo As Date = #1/1/2200#

        Dim Report As New CrpLaporanSalesDP
        Dim FilterDateName As String = ""
        Dim FilterDateStr As String = ""
        Dim FilterCustomer As String = "All"
        Dim FilterSalesPerson As String = "All"
        Dim GT_CurrencyCode As String = ""
        Dim GT_DP_Amount As String = ""
        Dim GT_JournalAmount As String = ""
        Dim GT_InvoiceAmount As String = ""
        Try
            If Me.TxtPrimaryCustomer.Text <> -1 Then
                NullablePrimaryCustomer = Me.TxtPrimaryCustomer.Text
            End If
            If Me.TxtPrimarySalesPerson.Text <> -1 Then
                NullablePrimarySalesPerson = Me.TxtPrimarySalesPerson.Text
            End If
            If Me.RboSO_Date.Checked Then
                SO_DateFrom = Me.DtpSO_DateFrom.Value.Date
                SO_DateTo = Me.DtpSO_DateTo.Value.Date
            End If
            If Me.RboDP_Date.Checked Then
                DP_DateFrom = Me.DtpDP_DateFrom.Value.Date
                DP_DateTo = Me.DtpDP_DateTo.Value.Date
            End If
            If Me.RboInvoiceDate.Checked Then
                Invoice_DateFrom = Me.DtpInvoice_DateFrom.Value.Date
                Invoice_DateTo = Me.DtpInvoice_DateTo.Value.Date
            End If

            Me.Cursor = Cursors.WaitCursor
            DtaDP.Connection.ConnectionString = Core_Module.GetConstring
            DtaSO.Connection.ConnectionString = Core_Module.GetConstring
            DtaJournal.Connection.ConnectionString = Core_Module.GetConstring
            DtaInvoice.Connection.ConnectionString = Core_Module.GetConstring

            DtaDP.Fill(TblDP, SO_DateFrom, SO_DateTo, DP_DateFrom, DP_DateTo, NullablePrimarySalesPerson, NullablePrimaryCustomer)
            If Me.CboDPStatus.SelectedIndex = 0 Or Me.CboDPStatus.SelectedIndex = 1 Then ''tidak perlu load jurnal dan invoice, jika filter status not created only
                DtaJournal.Fill(TblJournal, NullablePrimaryCustomer, NullablePrimarySalesPerson, SO_DateFrom, DP_DateFrom, SO_DateTo, DP_DateTo)
                DtaInvoice.Fill(TblInvoice, NullablePrimaryCustomer, NullablePrimarySalesPerson, SO_DateFrom, DP_DateFrom, SO_DateTo, DP_DateTo, Invoice_DateFrom, Invoice_DateTo)
            End If

            If Me.RboSO_Date.Checked Or Me.RboDP_Date.Checked Then
                For Each RowDP In TblDP
                    RowPrint = TblPrint.NewPrintDataRow
                    RowPrint.Sales_Person_Code = RowDP.Sales_Person_Code
                    RowPrint.Sales_Person_Desc = RowDP.Sales_Person_Desc
                    RowPrint.Customer_Code = RowDP.Customer_Code
                    RowPrint.Customer_Desc = RowDP.Customer_Desc
                    RowPrint.SO_No = RowDP.SO_No
                    RowPrint.SO_Date = RowDP.SO_Date
                    RowPrint.SO_Use_DP = RowDP._SO__Use__DP
                    RowPrint.SO_DP_Amount = RowDP.SO_DP_Amount
                    RowPrint.DP_No = RowDP.DP_No
                    RowPrint.DP_Date = RowDP.DP_Date
                    RowPrint.DP_Amount = RowDP.DP_Amount
                    RowPrint.Currency_Code = RowDP.Currency_Code
                    RowPrint.Primary_Sales_Person = RowDP.Primary_Sales_Person
                    RowPrint.Primary_Customer = RowDP.Primary_Customer
                    RowPrint.Primary_SO = RowDP.Primary_SO
                    RowPrint.Primary_DP = RowDP.Primary_DP
                    RowPrint.Decimal_Currency = RowDP.__Decimal__Currency
                    TblPrint.Rows.Add(RowPrint)
                Next
                If Me.RboSO_Date.Checked And (Me.CboDPStatus.SelectedIndex = 0 Or Me.CboDPStatus.SelectedIndex = 2) Then ''hanya load uncreated DP jika filter by SO date dan bukan filter status created only
                    DtaSO.Fill(TblSO, SO_DateFrom, NullablePrimarySalesPerson, NullablePrimaryCustomer, SO_DateTo)
                    For Each RowSO In TblSO
                        TblPrint.DefaultView.RowFilter = "[Primary SO] = " & RowSO.Primary_SO.ToString
                        If TblPrint.DefaultView.Count = 0 Then
                            RowPrint = TblPrint.NewPrintDataRow
                            RowPrint.Sales_Person_Code = RowSO.Sales_Person_Code
                            RowPrint.Sales_Person_Desc = RowSO.Sales_Person_Desc
                            RowPrint.Customer_Code = RowSO.Customer_Code
                            RowPrint.Customer_Desc = RowSO.Customer_Desc
                            RowPrint.SO_No = RowSO.SO_No
                            RowPrint.SO_Date = RowSO.SO_Date
                            RowPrint.SO_Use_DP = RowSO._SO__Use__DP
                            RowPrint.SO_DP_Amount = RowSO.SO_DP_Amount
                            RowPrint.Currency_Code = RowSO.Currency_Code
                            RowPrint.Primary_Sales_Person = RowSO.Primary_Sales_Person
                            RowPrint.Primary_Customer = RowSO.Primary_Customer
                            RowPrint.Primary_SO = RowSO.Primary_SO
                            RowPrint.Decimal_Currency = RowSO.__Decimal__Currency
                            TblPrint.Rows.Add(RowPrint)
                        End If
                    Next
                End If
                Dim ListOfCheckedDP As New List(Of Double)
                For Each RowPrint In TblPrint
                    If ListOfCheckedDP.Contains(RowPrint.Primary_DP) = False Then
                        ListOfCheckedDP.Add(RowPrint.Primary_DP)
                        TblInvoice.DefaultView.RowFilter = "[Primary DP] = " & RowPrint.Primary_DP.ToString
                        If TblInvoice.DefaultView.Count > 0 Then
                            RowInvoice = TblInvoice.DefaultView(0).Row
                            RowPrint.Primary_Invoice = RowInvoice.Primary_Invoice
                            RowPrint.Invoice_No = RowInvoice.Invoice_No
                            RowPrint.Invoice_Date = RowInvoice.Invoice_Date
                            RowPrint.Invoice_Amount = RowInvoice.Invoice_DP_Amount
                        End If
                    End If
                Next
                ListOfCheckedDP.Clear()
                For ii As Integer = 0 To TblPrint.Count - 1
                    RowPrint = TblPrint(ii)
                    If ListOfCheckedDP.Contains(RowPrint.Primary_DP) = False Then
                        ListOfCheckedDP.Add(RowPrint.Primary_DP)
                        TblJournal.DefaultView.RowFilter = "[Primary DP] = " & RowPrint.Primary_DP.ToString
                        TblJournal.DefaultView.Sort = "[Journal Date], [Journal No]"
                        If TblJournal.DefaultView.Count > 0 Then
                            Dim counter As Integer = 0
                            For Each DV As DataRowView In TblJournal.DefaultView
                                RowJournal = DV.Row
                                counter += 1
                                If counter = 1 Then
                                    RowPrint.Primary_Journal = RowJournal.Primary_Journal
                                    RowPrint.Journal_No = RowJournal.Journal_No
                                    RowPrint.Journal_Date = RowJournal.Journal_Date
                                    RowPrint.Journal_Amount = RowJournal.Journal_Amount
                                Else
                                    Dim RowPrint_ As DseLaporanSalesDP.PrintDataRow
                                    RowPrint_ = TblPrint.NewPrintDataRow
                                    For i As Integer = 0 To TblPrint.Columns.Count - 1
                                        RowPrint_.Item(i) = RowPrint.Item(i)
                                    Next
                                    RowPrint_.Is_Additional_Row_For_Journal = 1
                                    RowPrint_.Primary_Journal = RowJournal.Primary_Journal
                                    RowPrint_.Journal_No = RowJournal.Journal_No
                                    RowPrint_.Journal_Date = RowJournal.Journal_Date
                                    RowPrint_.Journal_Amount = RowJournal.Journal_Amount
                                    TblPrint.Rows.Add(RowPrint_)
                                End If
                            Next
                        End If
                    End If
                Next
                Console.WriteLine("")
            ElseIf Me.RboInvoiceDate.Checked Then
                For Each RowInvoice In TblInvoice
                    RowPrint = TblPrint.NewPrintDataRow
                    RowPrint.Primary_Invoice = RowInvoice.Primary_Invoice
                    RowPrint.Invoice_No = RowInvoice.Invoice_No
                    RowPrint.Invoice_Date = RowInvoice.Invoice_Date
                    RowPrint.Invoice_Amount = RowInvoice.Invoice_DP_Amount
                    RowPrint.Primary_DP = RowInvoice.Primary_DP
                    TblPrint.Rows.Add(RowPrint)
                Next
                Dim ListOfPrimaryDP As New List(Of Double)
                For Each RowPrint In TblPrint
                    If ListOfPrimaryDP.Contains(RowPrint.Primary_DP.ToString) = False Then
                        ListOfPrimaryDP.Add(RowPrint.Primary_DP.ToString)
                    End If
                    TblDP.DefaultView.RowFilter = "[Primary DP] = " & RowPrint.Primary_DP.ToString
                    If TblDP.DefaultView.Count > 0 Then
                        RowDP = TblDP.DefaultView(0).Row
                        RowPrint.Primary_Sales_Person = RowDP.Primary_Sales_Person
                        RowPrint.Sales_Person_Code = RowDP.Sales_Person_Code
                        RowPrint.Sales_Person_Desc = RowDP.Sales_Person_Desc
                        RowPrint.Primary_Customer = RowDP.Primary_Customer
                        RowPrint.Customer_Code = RowDP.Customer_Code
                        RowPrint.Customer_Desc = RowDP.Customer_Desc
                        RowPrint.Primary_SO = RowDP.Primary_SO
                        RowPrint.SO_No = RowDP.SO_No
                        RowPrint.SO_Date = RowDP.SO_Date
                        RowPrint.SO_Use_DP = RowDP._SO__Use__DP
                        RowPrint.SO_DP_Amount = RowDP.SO_DP_Amount
                        RowPrint.Primary_DP = RowDP.Primary_DP
                        RowPrint.DP_No = RowDP.DP_No
                        RowPrint.DP_Date = RowDP.DP_Date
                        RowPrint.DP_Amount = RowDP.DP_Amount
                        RowPrint.Currency_Code = RowDP.Currency_Code
                        RowPrint.Decimal_Currency = RowDP.__Decimal__Currency
                    End If
                Next
                For Each PrimaryDP As Double In ListOfPrimaryDP
                    TblJournal.DefaultView.RowFilter = "[Primary DP] = " & PrimaryDP.ToString
                    TblJournal.DefaultView.Sort = "[Journal Date], [Journal No]"
                    TblPrint.DefaultView.RowFilter = "[Primary DP] = " & PrimaryDP.ToString
                    If TblJournal.DefaultView.Count > 0 And TblPrint.DefaultView.Count > 0 Then
                        RowPrint = TblPrint.DefaultView(0).Row
                        RowJournal = TblJournal.DefaultView(0).Row
                        RowPrint.Primary_Journal = RowJournal.Primary_Journal
                        RowPrint.Journal_No = RowJournal.Journal_No
                        RowPrint.Journal_Date = RowJournal.Journal_Date
                        RowPrint.Journal_Amount = RowJournal.Journal_Amount
                        If TblJournal.DefaultView.Count > TblPrint.DefaultView.Count Then
                            For i As Integer = 1 To TblJournal.DefaultView.Count - TblPrint.DefaultView.Count
                                RowJournal = TblJournal.DefaultView(i).Row
                                Dim RowPrint_ As DseLaporanSalesDP.PrintDataRow
                                RowPrint_ = TblPrint.NewPrintDataRow
                                For ii As Integer = 0 To TblPrint.Columns.Count - 1
                                    RowPrint_.Item(ii) = RowPrint.Item(ii)
                                Next
                                RowPrint_.Is_Additional_Row_For_Journal = 1
                                RowPrint_.Primary_Journal = RowJournal.Primary_Journal
                                RowPrint_.Journal_No = RowJournal.Journal_No
                                RowPrint_.Journal_Date = RowJournal.Journal_Date
                                RowPrint_.Journal_Amount = RowJournal.Journal_Amount
                                TblPrint.Rows.Add(RowPrint_)
                            Next
                        End If
                    End If
                Next
                Console.WriteLine("")
            End If

            If Me.CboDPStatus.SelectedIndex <> 0 Then
                Me.FilterDP_Status(TblPrint)
            End If

            If Me.CboPaymentStatus.SelectedIndex <> 0 Then
                Me.FilterDP_PaymentStatus(TblPrint)
            End If

            If Me.CboRealizationStatus.SelectedIndex <> 0 Then
                Me.FilterDP_Realization(TblPrint)
            End If

            Me.SetDifferentDP(TblPrint)

            Me.OrderDataByGroup(TblPrint)

            Me.CalculateTotalAmount(TblPrint, GT_CurrencyCode, GT_DP_Amount, GT_JournalAmount, GT_InvoiceAmount)

            If Me.RboSO_Date.Checked Then
                FilterDateName = "SO"
                FilterDateStr = Format(Me.DtpSO_DateFrom.Value.Date, "dd MMM yyyy") & " to " & Format(Me.DtpSO_DateTo.Value.Date, "dd MMM yyyy")
            ElseIf Me.RboDP_Date.Checked Then
                FilterDateName = "DP"
                FilterDateStr = Format(Me.DtpDP_DateFrom.Value.Date, "dd MMM yyyy") & " to " & Format(Me.DtpDP_DateTo.Value.Date, "dd MMM yyyy")
            ElseIf Me.RboInvoiceDate.Checked Then
                FilterDateName = "Invoice"
                FilterDateStr = Format(Me.DtpInvoice_DateFrom.Value.Date, "dd MMM yyyy") & " to " & Format(Me.DtpInvoice_DateTo.Value.Date, "dd MMM yyyy")
            End If
            If Me.TxtCustomerName.Text <> "" Then
                FilterCustomer = Me.TxtCustomerName.Text
            End If
            If Me.TxtSalesPersonDesc.Text <> "" Then
                FilterSalesPerson = Me.TxtSalesPersonDesc.Text
            End If
            If TblPrint.Count > 0 Then
                Report.SetDataSource(DS)
                Report.SetParameterValue(Report.Parameter_FilterDateName.ParameterFieldName, FilterDateName)
                Report.SetParameterValue(Report.Parameter_FilterDateStr.ParameterFieldName, FilterDateStr)
                Report.SetParameterValue(Report.Parameter_FilterCustomer.ParameterFieldName, FilterCustomer)
                Report.SetParameterValue(Report.Parameter_FilterSalesPerson.ParameterFieldName, FilterSalesPerson)
                Report.SetParameterValue(Report.Parameter_FilterDPStatus.ParameterFieldName, Me.CboDPStatus.Text)
                Report.SetParameterValue(Report.Parameter_FilterPayment.ParameterFieldName, Me.CboPaymentStatus.Text)
                Report.SetParameterValue(Report.Parameter_FilterRealization.ParameterFieldName, Me.CboRealizationStatus.Text)
                Report.SetParameterValue(Report.Parameter_GroupNameStr.ParameterFieldName, Me.CboGroup.Text.Replace("By ", ""))
                Report.SetParameterValue(Report.Parameter_P_GT_CurrencyCode.ParameterFieldName, GT_CurrencyCode)
                Report.SetParameterValue(Report.Parameter_P_GT_DP_Amount.ParameterFieldName, GT_DP_Amount)
                Report.SetParameterValue(Report.Parameter_P_GT_JournalAmount.ParameterFieldName, GT_JournalAmount)
                Report.SetParameterValue(Report.Parameter_P_GT_InvoiceAmount.ParameterFieldName, GT_InvoiceAmount)
                Me.crv.ReportSource = Report
            Else
                MsgBox("There is no data.", MsgBoxStyle.Information, "Information")
                Me.crv.ReportSource = Nothing
            End If
            'Me.crv.Zoom(90)
            Me.crv.Refresh()
            Me.crv.Show()

            Console.WriteLine("")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Sub FilterDP_Status(ByVal vTblPrint As DseLaporanSalesDP.PrintDataDataTable)
        Dim vRowPrint As DseLaporanSalesDP.PrintDataRow
        Try
            For i As Integer = vTblPrint.Rows.Count - 1 To 0 Step -1
                vRowPrint = vTblPrint(i)
                If Me.CboDPStatus.SelectedIndex = 1 Then ''hapus SO yang tidak ada DP-nya, jika filter status: created DP only
                    If vRowPrint.Primary_DP = -1 Then
                        vTblPrint.Rows.Remove(vRowPrint)
                    End If
                ElseIf Me.CboDPStatus.SelectedIndex = 2 Then ''hapus DP yang sudah dibuat, jika filter status: not created DP only
                    If vRowPrint.Primary_DP <> -1 Then
                        vTblPrint.Rows.Remove(vRowPrint)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub FilterDP_PaymentStatus(ByVal vTblPrint As DseLaporanSalesDP.PrintDataDataTable)
        Dim vRowPrint As DseLaporanSalesDP.PrintDataRow
        Dim TblPrintDummy As New DseLaporanSalesDP.PrintDataDataTable
        Dim RowPrintDummy As DseLaporanSalesDP.PrintDataRow
        Try
            TblPrintDummy = vTblPrint.Copy
            For i As Integer = vTblPrint.Rows.Count - 1 To 0 Step -1
                vRowPrint = vTblPrint(i)
                TblPrintDummy.DefaultView.RowFilter = "[Primary DP] = " & vRowPrint.Primary_DP.ToString
                Dim TotalPayment As Decimal = 0
                For Each DV As DataRowView In TblPrintDummy.DefaultView
                    RowPrintDummy = DV.Row
                    TotalPayment += RowPrintDummy.Journal_Amount
                Next
                If Me.CboPaymentStatus.SelectedIndex = 1 Then ''jika filter payment status: paid only, hapus yang total payment <> dp amount
                    If vRowPrint.DP_Amount <> TotalPayment Or (vRowPrint.DP_Amount = TotalPayment And TotalPayment = 0) Then
                        vTblPrint.Rows.Remove(vRowPrint)
                    End If
                ElseIf Me.CboPaymentStatus.SelectedIndex = 2 Then ''jika filter payment status: partial/not payment only, hapus yang total payment = dp amount
                    If vRowPrint.DP_Amount = TotalPayment And TotalPayment > 0 Then
                        vTblPrint.Rows.Remove(vRowPrint)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub FilterDP_Realization(ByVal vTblPrint As DseLaporanSalesDP.PrintDataDataTable)
        Dim vRowPrint As DseLaporanSalesDP.PrintDataRow
        Dim TblPrintDummy As New DseLaporanSalesDP.PrintDataDataTable
        Dim RowPrintDummy As DseLaporanSalesDP.PrintDataRow
        Try
            TblPrintDummy = vTblPrint.Copy
            For i As Integer = vTblPrint.Rows.Count - 1 To 0 Step -1
                vRowPrint = vTblPrint(i)
                TblPrintDummy.DefaultView.RowFilter = "[Primary DP] = " & vRowPrint.Primary_DP.ToString & " AND [Is Additional Row For Journal]=0"
                Dim TotalRealization As Decimal = 0
                For Each DV As DataRowView In TblPrintDummy.DefaultView
                    RowPrintDummy = DV.Row
                    TotalRealization += RowPrintDummy.Invoice_Amount
                Next
                If Me.CboRealizationStatus.SelectedIndex = 1 Then ''full realization
                    If vRowPrint.DP_Amount <> TotalRealization Or (vRowPrint.DP_Amount = TotalRealization And TotalRealization = 0) Then
                        vTblPrint.Rows.Remove(vRowPrint)
                    End If
                ElseIf Me.CboRealizationStatus.SelectedIndex = 2 Then ''not realization or partially realization
                    If vRowPrint.DP_Amount = TotalRealization And TotalRealization > 0 Then
                        vTblPrint.Rows.Remove(vRowPrint)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub SetDifferentDP(ByVal vTblPrint As DseLaporanSalesDP.PrintDataDataTable)
        Dim vRowPrint As DseLaporanSalesDP.PrintDataRow
        Dim TblPrintDummy As New DseLaporanSalesDP.PrintDataDataTable
        Dim RowPrintDummy As DseLaporanSalesDP.PrintDataRow
        Try
            TblPrintDummy = vTblPrint.Copy
            For i As Integer = vTblPrint.Rows.Count - 1 To 0 Step -1
                vRowPrint = vTblPrint(i)
                If vRowPrint.SO_Use_DP = 1 Then
                    TblPrintDummy.DefaultView.RowFilter = "[Primary SO] = " & vRowPrint.Primary_SO.ToString & " AND [Is Additional Row For Journal]=0"
                    Dim TotalCreatedDP As Decimal = 0
                    For Each DV As DataRowView In TblPrintDummy.DefaultView
                        RowPrintDummy = DV.Row
                        TotalCreatedDP += RowPrintDummy.DP_Amount
                    Next
                    If vRowPrint.SO_DP_Amount <> TotalCreatedDP And vRowPrint.Primary_DP <> -1 Then
                        vRowPrint.Different_DP_Amount = 1
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub OrderDataByGroup(ByVal vTblPrint As DseLaporanSalesDP.PrintDataDataTable)
        'Dim vRowPrint As DseLaporanSalesDP.PrintDataRow
        Dim TblPrintDummy As New DseLaporanSalesDP.PrintDataDataTable
        Dim RowPrintDummy As DseLaporanSalesDP.PrintDataRow
        Try
            TblPrintDummy = vTblPrint.Copy
            vTblPrint.Rows.Clear()
            If Me.CboGroup.SelectedIndex = 0 Then
                TblPrintDummy.DefaultView.RowFilter = ""
                TblPrintDummy.DefaultView.Sort = "[Customer Desc], [SO No], [DP Date], [DP No], [Journal Date], [Journal No], [Is Additional Row For Journal]"
                For Each DV As DataRowView In TblPrintDummy.DefaultView
                    RowPrintDummy = DV.Row
                    If RowPrintDummy.SO_Use_DP = 1 And RowPrintDummy.Primary_DP = -1 Then
                        RowPrintDummy.DP_Amount = RowPrintDummy.SO_DP_Amount
                    End If
                    RowPrintDummy.Group_Str = RowPrintDummy.Customer_Desc
                    vTblPrint.ImportRow(RowPrintDummy)
                Next
            Else
                TblPrintDummy.DefaultView.RowFilter = ""
                TblPrintDummy.DefaultView.Sort = "[Sales Person Desc], [SO No], [DP Date], [DP No], [Journal Date], [Journal No], [Is Additional Row For Journal]"
                For Each DV As DataRowView In TblPrintDummy.DefaultView
                    RowPrintDummy = DV.Row
                    If RowPrintDummy.SO_Use_DP = 1 And RowPrintDummy.Primary_DP = -1 Then
                        RowPrintDummy.DP_Amount = RowPrintDummy.SO_DP_Amount
                    End If
                    RowPrintDummy.Group_Str = RowPrintDummy.Sales_Person_Desc
                    vTblPrint.ImportRow(RowPrintDummy)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub CalculateTotalAmount(ByVal vTblPrint As DseLaporanSalesDP.PrintDataDataTable, ByRef vGT_CurrencyCode As String, ByRef vGT_DPamount As String, ByRef vGT_JournalAmount As String, ByRef vGT_InvoiceAmount As String)
        Dim vRowPrint As DseLaporanSalesDP.PrintDataRow
        Dim TblSubTotalPerRefParty As New DseLaporanSalesDP.SubTotalDataDataTable
        Dim RowSubTotalPerRefParty As DseLaporanSalesDP.SubTotalDataRow
        Dim TblGrandTotalPerRefParty As New DseLaporanSalesDP.SubTotalDataDataTable
        Dim RowGrandTotalPerRefParty As DseLaporanSalesDP.SubTotalDataRow
        Dim PrimaryRefParty As Double = -1
        Dim PrevPrimaryDP As Double = -1
        Dim PrevPrimaryInvoice As Double = -1
        Dim PrevPrimarySO As Double = -1
        Try
            ''sub total
            For Each vRowPrint In vTblPrint
                If Me.CboGroup.SelectedIndex = 0 Then
                    TblSubTotalPerRefParty.DefaultView.RowFilter = "[Primary Ref Party] = " & vRowPrint.Primary_Customer.ToString & " AND [Currency Code] = '" & vRowPrint.Currency_Code & "'"
                    PrimaryRefParty = vRowPrint.Primary_Customer
                ElseIf Me.CboGroup.SelectedIndex = 1 Then
                    TblSubTotalPerRefParty.DefaultView.RowFilter = "[Primary Ref Party] = " & vRowPrint.Primary_Sales_Person.ToString & " AND [Currency Code] = '" & vRowPrint.Currency_Code & "'"
                    PrimaryRefParty = vRowPrint.Primary_Sales_Person
                End If
                If TblSubTotalPerRefParty.DefaultView.Count = 0 Then
                    RowSubTotalPerRefParty = TblSubTotalPerRefParty.NewSubTotalDataRow
                    RowSubTotalPerRefParty.Primary_Ref_Party = PrimaryRefParty
                    RowSubTotalPerRefParty.Currency_Code = vRowPrint.Currency_Code
                    RowSubTotalPerRefParty.DP_Amount = vRowPrint.DP_Amount
                    RowSubTotalPerRefParty.Journal_Amount = vRowPrint.Journal_Amount
                    RowSubTotalPerRefParty.Invoice_Amount = vRowPrint.Invoice_Amount
                    RowSubTotalPerRefParty.Precision = vRowPrint.Decimal_Currency
                    TblSubTotalPerRefParty.Rows.Add(RowSubTotalPerRefParty)
                Else
                    RowSubTotalPerRefParty = TblSubTotalPerRefParty.DefaultView(0).Row
                    If vRowPrint.Primary_DP <> PrevPrimaryDP Then ''jangan jumlahkan jika DP No sama dengan sebelumnya, ini dikarenakan item ybs adalah item journal
                        RowSubTotalPerRefParty.DP_Amount += vRowPrint.DP_Amount
                    End If
                    RowSubTotalPerRefParty.Journal_Amount += vRowPrint.Journal_Amount
                    RowSubTotalPerRefParty.Invoice_Amount += vRowPrint.Invoice_Amount
                End If
                PrevPrimaryDP = vRowPrint.Primary_DP
                If vRowPrint.Primary_SO = PrevPrimarySO Then
                    vRowPrint.Same_Sales_Order = 1
                End If
                PrevPrimaryInvoice = vRowPrint.Primary_Invoice
                PrevPrimarySO = vRowPrint.Primary_SO
            Next
            For Each vRowPrint In vTblPrint
                If Me.CboGroup.SelectedIndex = 0 Then
                    PrimaryRefParty = vRowPrint.Primary_Customer
                ElseIf Me.CboGroup.SelectedIndex = 1 Then
                    PrimaryRefParty = vRowPrint.Primary_Sales_Person
                End If
                TblSubTotalPerRefParty.DefaultView.RowFilter = "[Primary Ref Party] = " & PrimaryRefParty.ToString
                TblSubTotalPerRefParty.DefaultView.Sort = "[Currency Code]"
                Dim counter As Integer = 0
                For Each DV As DataRowView In TblSubTotalPerRefParty.DefaultView
                    RowSubTotalPerRefParty = DV.Row
                    counter += 1
                    If counter > 1 Then
                        vRowPrint.Sub_Total_Currency_Str = vRowPrint.Sub_Total_Currency_Str & vbNewLine
                        vRowPrint.Sub_Total_DP_Amount_Str = vRowPrint.Sub_Total_DP_Amount_Str & vbNewLine
                        vRowPrint.Sub_Total_Journal_Amount_Str = vRowPrint.Sub_Total_Journal_Amount_Str & vbNewLine
                        vRowPrint.Sub_Total_Invoice_Amount_Str = vRowPrint.Sub_Total_Invoice_Amount_Str & vbNewLine
                    End If
                    vRowPrint.Sub_Total_Currency_Str = vRowPrint.Sub_Total_Currency_Str & RowSubTotalPerRefParty.Currency_Code
                    If RowSubTotalPerRefParty.DP_Amount > 0 Then
                        vRowPrint.Sub_Total_DP_Amount_Str = vRowPrint.Sub_Total_DP_Amount_Str & Me.objCD.DigitGrouping(RowSubTotalPerRefParty.DP_Amount, RowSubTotalPerRefParty.Precision)
                    Else
                        vRowPrint.Sub_Total_DP_Amount_Str = vRowPrint.Sub_Total_DP_Amount_Str & ""
                    End If
                    If RowSubTotalPerRefParty.Journal_Amount > 0 Then
                        vRowPrint.Sub_Total_Journal_Amount_Str = vRowPrint.Sub_Total_Journal_Amount_Str & Me.objCD.DigitGrouping(RowSubTotalPerRefParty.Journal_Amount, RowSubTotalPerRefParty.Precision)
                    Else
                        vRowPrint.Sub_Total_Journal_Amount_Str = vRowPrint.Sub_Total_Journal_Amount_Str & ""
                    End If
                    If RowSubTotalPerRefParty.Invoice_Amount > 0 Then
                        vRowPrint.Sub_Total_Invoice_Amount_Str = vRowPrint.Sub_Total_Invoice_Amount_Str & Me.objCD.DigitGrouping(RowSubTotalPerRefParty.Invoice_Amount, RowSubTotalPerRefParty.Precision)
                    Else
                        vRowPrint.Sub_Total_Invoice_Amount_Str = vRowPrint.Sub_Total_Invoice_Amount_Str & ""
                    End If
                Next
            Next

            ''grand total
            TblSubTotalPerRefParty.DefaultView.RowFilter = ""
            TblSubTotalPerRefParty.DefaultView.Sort = "[Currency Code]"
            For Each DV As DataRowView In TblSubTotalPerRefParty.DefaultView
                RowSubTotalPerRefParty = DV.Row
                TblGrandTotalPerRefParty.DefaultView.RowFilter = "[Currency Code] = '" & RowSubTotalPerRefParty.Currency_Code & "'"
                If TblGrandTotalPerRefParty.DefaultView.Count = 0 Then
                    RowGrandTotalPerRefParty = TblGrandTotalPerRefParty.NewSubTotalDataRow
                    RowGrandTotalPerRefParty.Currency_Code = RowSubTotalPerRefParty.Currency_Code
                    RowGrandTotalPerRefParty.DP_Amount = RowSubTotalPerRefParty.DP_Amount
                    RowGrandTotalPerRefParty.Journal_Amount = RowSubTotalPerRefParty.Journal_Amount
                    RowGrandTotalPerRefParty.Invoice_Amount = RowSubTotalPerRefParty.Invoice_Amount
                    RowGrandTotalPerRefParty.Precision = RowSubTotalPerRefParty.Precision
                    TblGrandTotalPerRefParty.Rows.Add(RowGrandTotalPerRefParty)
                Else
                    RowGrandTotalPerRefParty = TblGrandTotalPerRefParty.DefaultView(0).Row
                    RowGrandTotalPerRefParty.DP_Amount += RowSubTotalPerRefParty.DP_Amount
                    RowGrandTotalPerRefParty.Journal_Amount += RowSubTotalPerRefParty.Journal_Amount
                    RowGrandTotalPerRefParty.Invoice_Amount += RowSubTotalPerRefParty.Invoice_Amount
                End If
            Next
            Dim AddNewLine As Boolean = False
            For Each RowGrandTotalPerRefParty In TblGrandTotalPerRefParty
                If AddNewLine Then
                    vGT_CurrencyCode = vGT_CurrencyCode & vbNewLine
                    vGT_DPamount = vGT_DPamount & vbNewLine
                    vGT_JournalAmount = vGT_JournalAmount & vbNewLine
                    vGT_InvoiceAmount = vGT_InvoiceAmount & vbNewLine
                End If
                AddNewLine = True
                vGT_CurrencyCode = vGT_CurrencyCode & RowGrandTotalPerRefParty.Currency_Code
                If RowGrandTotalPerRefParty.DP_Amount > 0 Then
                    vGT_DPamount = vGT_DPamount & Me.objCD.DigitGrouping(RowGrandTotalPerRefParty.DP_Amount, RowGrandTotalPerRefParty.Precision)
                Else
                    vGT_DPamount = vGT_DPamount & ""
                End If
                If RowGrandTotalPerRefParty.Journal_Amount > 0 Then
                    vGT_JournalAmount = vGT_JournalAmount & Me.objCD.DigitGrouping(RowGrandTotalPerRefParty.Journal_Amount, RowGrandTotalPerRefParty.Precision)
                Else
                    vGT_JournalAmount = vGT_JournalAmount & ""
                End If
                If RowGrandTotalPerRefParty.Invoice_Amount > 0 Then
                    vGT_InvoiceAmount = vGT_InvoiceAmount & Me.objCD.DigitGrouping(RowGrandTotalPerRefParty.Invoice_Amount, RowGrandTotalPerRefParty.Precision)
                Else
                    vGT_InvoiceAmount = vGT_InvoiceAmount & ""
                End If
            Next
            Console.WriteLine("")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class