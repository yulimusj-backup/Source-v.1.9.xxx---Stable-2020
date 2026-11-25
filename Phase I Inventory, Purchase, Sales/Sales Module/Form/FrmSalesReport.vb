Option Explicit On
Option Strict Off
Public Class FrmSalesReport
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.FromArgb(170, 255, 190)
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjCustomer As New Core_Module.clsMasterCustomer
    Dim ObjSales As New Core_Module.ClsMasterSalesPerson
    Dim ObjAccount As New Core_Module.ClsMasterAkun1
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Dim ObjCD As New Core_Module.ClsCekData
    Dim objItem As New Inventory_Module.ClsTransaksiMaterialPlanningGroupItem
    Dim ObjCategory As New Core_Module.ClsMasterCategory
    Dim ObjProductGroup As New Core_Module.ClsMasterProductGroup
    Dim EraseRowIndex As Integer = 0
    Dim DummyIndex As Integer = -1
    Dim buffer As String = ""
    Dim ProductGroup_Level As Integer = 3
    Dim Brand_Level As Integer = 2
    Dim SubBrand_Level As Integer = 1

    'Find Product Code
    Private Sub txtCodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCodeNo.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.txtCodeNo.Text <> "" Then
                    If Me.ObjProduct.ApakahSudahAda(Me.txtCodeNo.Text) = True Then
                        Me.txtPrimaryCodeNo.Text = CStr(Me.ObjProduct.GetPrimary(Me.txtCodeNo.Text))
                        Me.txtCodeNo.Text = Me.ObjProduct.GetProductCode(CDbl(Me.txtPrimaryCodeNo.Text))
                        Me.TxtDescProduct.Text = ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                    Else
                        Me.BtnFindCodeNo.PerformClick()
                    End If
                Else
                    Me.BtnFindCodeNo.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodeNo.Leave
        Dim hj As New Core_Module.frmCariMasterProduct
        Try
            If Me.txtCodeNo.Text <> "" Then
                If Me.ObjProduct.ApakahSudahAda(Me.txtCodeNo.Text) = True Then
                    Me.txtPrimaryCodeNo.Text = CStr(Me.ObjProduct.GetPrimary(Me.txtCodeNo.Text))
                    Me.txtCodeNo.Text = Me.ObjProduct.GetProductCode(CDbl(Me.txtPrimaryCodeNo.Text))
                    Me.TxtDescProduct.Text = ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                Else
                    If Me.txtCodeNo.Text <> "" Then
                        hj.txtCari.Text = Me.txtCodeNo.Text
                    End If
                    hj.ShowDialog()
                    If hj.PrimaryProduct > 0 Then
                        Me.txtPrimaryCodeNo.Text = CStr(hj.PrimaryProduct)
                        Me.txtCodeNo.Text = Me.ObjProduct.GetProductCode(CDbl(Me.txtPrimaryCodeNo.Text))
                        Me.TxtDescProduct.Text = ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                    Else
                        Me.txtPrimaryCodeNo.Text = "-1"
                        Me.txtCodeNo.Clear()
                        Me.TxtDescProduct.Clear()
                        Me.txtCodeNo.Focus()
                    End If
                End If
            Else
                Me.txtPrimaryCodeNo.Text = "-1"
                Me.TxtDescProduct.Clear()
                Me.txtCodeNo.Clear()
                Me.ChkAllProduct.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub
    Private Sub BtnFindCodeNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCodeNo.Click
        Dim hj As New Core_Module.frmCariMasterProduct
        Try
            If Me.txtCodeNo.Text <> "" Then
                hj.txtCari.Text = Me.txtCodeNo.Text
            End If
            hj.ShowDialog()
            If hj.PrimaryProduct > 0 Then
                Me.txtPrimaryCodeNo.Text = CStr(hj.PrimaryProduct)
                Me.txtCodeNo.Text = Me.ObjProduct.GetProductCode(CDbl(Me.txtPrimaryCodeNo.Text))
                Me.TxtDescProduct.Text = ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
            Else
                Me.txtPrimaryCodeNo.Text = "-1"
                Me.txtCodeNo.Clear()
                Me.TxtDescProduct.Clear()
                Me.txtCodeNo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub
    'Find Customer Code
    Private Sub TxtCustomerCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCustomerCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtCustomerCode.Text <> "" Then
                    If Me.ObjCustomer.ApakahSudahAda(Me.TxtCustomerCode.Text) = True Then
                        Me.TxtPrimaryCustomer.Text = CStr(Me.ObjCustomer.GetPrimary(Me.TxtCustomerCode.Text))
                        Me.TxtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(CDbl(Me.TxtPrimaryCustomer.Text))
                        Me.TxtDescCustomer.Text = Me.ObjCustomer.GetCustName(CDbl(Me.TxtPrimaryCustomer.Text))
                    Else
                        Me.BtnFindCustomerCode.PerformClick()
                    End If
                Else
                    Me.BtnFindCustomerCode.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub TxtCustomerCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerCode.Leave
        Dim hj As New Core_Module.FrmCariMasterCustomer
        Try
            If Me.TxtCustomerCode.Text <> "" Then
                If Me.ObjCustomer.ApakahSudahAda(Me.TxtCustomerCode.Text) = True Then
                    Me.TxtPrimaryCustomer.Text = CStr(Me.ObjCustomer.GetPrimary(Me.TxtCustomerCode.Text))
                    Me.TxtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(CDbl(Me.TxtPrimaryCustomer.Text))
                    Me.TxtDescCustomer.Text = Me.ObjCustomer.GetCustName(CDbl(Me.TxtPrimaryCustomer.Text))
                Else
                    If Me.TxtCustomerCode.Text <> "" Then
                        hj.txtCari.Text = Me.TxtCustomerCode.Text
                    End If
                    hj.ShowDialog()

                    If hj.PrimaryCustomer > 0 Then
                        Me.TxtPrimaryCustomer.Text = CStr(hj.PrimaryCustomer)
                        Me.TxtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(CDbl(Me.TxtPrimaryCustomer.Text))
                        Me.TxtDescCustomer.Text = Me.ObjCustomer.GetCustName(CDbl(Me.TxtPrimaryCustomer.Text))
                    Else
                        Me.TxtPrimaryCustomer.Text = "-1"
                        Me.TxtCustomerCode.Clear()
                        Me.TxtDescCustomer.Clear()
                        Me.TxtCustomerCode.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryCustomer.Text = "-1"
                Me.TxtCustomerCode.Clear()
                Me.TxtDescCustomer.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub
    Private Sub BtnFindCustomerCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCustomerCode.Click
        Dim hj As New Core_Module.FrmCariMasterCustomer
        Try
            If Me.TxtCustomerCode.Text <> "" Then
                hj.txtCari.Text = Me.TxtCustomerCode.Text
            End If

            hj.ShowDialog()
            If hj.PrimaryCustomer > 0 Then
                Me.TxtPrimaryCustomer.Text = CStr(hj.PrimaryCustomer)
                Me.TxtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(CDbl(Me.TxtPrimaryCustomer.Text))
                Me.TxtDescCustomer.Text = Me.ObjCustomer.GetCustName(CDbl(Me.TxtPrimaryCustomer.Text))
            Else
                Me.TxtPrimaryCustomer.Text = "-1"
                Me.TxtCustomerCode.Clear()
                Me.TxtDescCustomer.Clear()
                Me.TxtCustomerCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub
    'Find Sales Code
    Private Sub TxtKodeSales_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtKodeSales.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtKodeSales.Text <> "" Then
                    If Me.ObjSales.ApakahSudahAda(Me.TxtKodeSales.Text) = True Then
                        Me.TxtPrimarySales.Text = CStr(Me.ObjSales.GetPrimary(Me.TxtKodeSales.Text))
                        Me.TxtKodeSales.Text = Me.ObjSales.GetCode(CDbl(Me.TxtPrimarySales.Text))
                        Me.TxtDescSalesPerson.Text = Me.ObjSales.GetName(CDbl(Me.TxtPrimarySales.Text))

                    Else
                        Me.BtnFindSales.PerformClick()
                    End If
                Else
                    Me.BtnFindSales.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub TxtKodeSales_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtKodeSales.Leave
        Dim hj As New Core_Module.FrmMasterSalesPerson_Cari
        Try
            If Me.TxtKodeSales.Text <> "" Then
                If Me.ObjSales.ApakahSudahAda(Me.TxtKodeSales.Text) = True Then
                    Me.TxtPrimarySales.Text = CStr(Me.ObjSales.GetPrimary(Me.TxtKodeSales.Text))
                    Me.TxtKodeSales.Text = Me.ObjSales.GetCode(CDbl(Me.TxtPrimarySales.Text))
                    Me.TxtDescSalesPerson.Text = Me.ObjSales.GetName(CDbl(Me.TxtPrimarySales.Text))
                Else
                    If Me.TxtKodeSales.Text <> "" Then
                        hj.txtCari.Text = Me.TxtKodeSales.Text
                    End If
                    hj.ShowDialog()
                    If hj.PrimarySalesPerson > 0 Then
                        Me.TxtPrimarySales.Text = CStr(hj.PrimarySalesPerson)
                        Me.TxtKodeSales.Text = Me.ObjSales.GetCode(CDbl(Me.TxtPrimarySales.Text))
                        Me.TxtDescSalesPerson.Text = Me.ObjSales.GetName(CDbl(Me.TxtPrimarySales.Text))
                    Else
                        Me.TxtPrimarySales.Text = "-1"
                        Me.TxtKodeSales.Clear()
                        Me.TxtDescSalesPerson.Clear()
                        Me.TxtKodeSales.Focus()
                    End If
                End If
            Else
                Me.TxtPrimarySales.Text = "-1"
                Me.TxtKodeSales.Clear()
                Me.TxtDescSalesPerson.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub BtnFindSales_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindSales.Click
        Dim hj As New Core_Module.FrmMasterSalesPerson_Cari
        Try
            If Me.TxtKodeSales.Text <> "" Then
                hj.txtCari.Text = Me.TxtKodeSales.Text
            End If
            hj.ShowDialog()
            If hj.PrimarySalesPerson > 0 Then
                Me.TxtPrimarySales.Text = CStr(hj.PrimarySalesPerson)
                Me.TxtKodeSales.Text = Me.ObjSales.GetCode(CDbl(Me.TxtPrimarySales.Text))
                Me.TxtDescSalesPerson.Text = Me.ObjSales.GetName(CDbl(Me.TxtPrimarySales.Text))
            Else
                Me.TxtPrimarySales.Text = "-1"
                Me.TxtKodeSales.Clear()
                Me.TxtDescSalesPerson.Clear()
                Me.TxtKodeSales.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub
    'Find Receivable Account
    Private Sub TxtAccCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtAccCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtAccCode.Text <> "" Then
                    If Me.ObjAccount.ApakahSudahAda(Me.TxtAccCode.Text) = True Then
                        Me.TxtPrimaryAcc.Text = CStr(Me.ObjAccount.GetPrimary(Me.TxtAccCode.Text))
                        Me.TxtAccCode.Text = Me.ObjAccount.GetAccountCode(CDbl(Me.TxtPrimaryAcc.Text))
                        Me.TxtDescAccount.Text = Me.ObjAccount.GetDescription_ByPrimaryAcc(CDbl(Me.TxtPrimaryAcc.Text))
                    Else
                        Me.BtnFindAcc.PerformClick()
                    End If
                Else
                    Me.BtnFindAcc.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub TxtAccCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccCode.Leave
        Dim hj As New Core_Module.FrmMasterAkun_Cari1
        Try
            If Me.TxtAccCode.Text <> "" Then
                If Me.ObjAccount.ApakahSudahAda(Me.TxtAccCode.Text) = True Then
                    Me.TxtPrimaryAcc.Text = CStr(Me.ObjAccount.GetPrimary(Me.TxtAccCode.Text))
                    Me.TxtAccCode.Text = Me.ObjAccount.GetAccountCode(CDbl(Me.TxtPrimaryAcc.Text))
                    Me.TxtDescAccount.Text = Me.ObjAccount.GetDescription_ByPrimaryAcc(CDbl(Me.TxtPrimaryAcc.Text))
                Else
                    If Me.TxtAccCode.Text <> "" Then
                        hj.txtCari.Text = Me.TxtAccCode.Text
                    End If
                    hj.ShowDialog()
                    If hj.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryAcc.Text = CStr(hj.mPrimaryAkun)
                        Me.TxtAccCode.Text = Me.ObjAccount.GetAccountCode(CDbl(Me.TxtPrimaryAcc.Text))
                        Me.TxtDescAccount.Text = Me.ObjAccount.GetDescription_ByPrimaryAcc(CDbl(Me.TxtPrimaryAcc.Text))
                    Else
                        Me.TxtPrimaryAcc.Text = "-1"
                        Me.TxtAccCode.Clear()
                        Me.TxtDescAccount.Clear()
                        Me.TxtAccCode.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryAcc.Text = "-1"
                Me.TxtAccCode.Clear()
                Me.TxtDescAccount.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub
    Private Sub BtnFindAcc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindAcc.Click
        Dim hj As New Core_Module.FrmMasterAkun_Cari1
        Try
            If Me.TxtAccCode.Text <> "" Then
                hj.txtCari.Text = Me.TxtAccCode.Text
            End If
            hj.ShowDialog()
            If hj.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAcc.Text = CStr(hj.mPrimaryAkun)
                Me.TxtAccCode.Text = Me.ObjAccount.GetAccountCode(CDbl(Me.TxtPrimaryAcc.Text))
                Me.TxtDescAccount.Text = Me.ObjAccount.GetDescription_ByPrimaryAcc(CDbl(Me.TxtPrimaryAcc.Text))
            Else
                Me.TxtPrimaryAcc.Text = "-1"
                Me.TxtAccCode.Clear()
                Me.TxtDescAccount.Clear()
                Me.TxtAccCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub
    Private Sub txtCodeNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodeNo.GotFocus, TxtCustomerCode.GotFocus, TxtKodeSales.GotFocus, TxtAccCode.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.BtnPrint.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.txtCodeNo.BackColor
                        Else
                            Txt.BackColor = Me.BackColorAll
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub txtCodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodeNo.LostFocus, TxtCustomerCode.LostFocus, TxtKodeSales.LostFocus, TxtAccCode.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.BtnPrint.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.txtCodeNo.BackColor
                        Else
                            Txt.BackColor = Me.txtPrimaryCodeNo.BackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


#Region "ALL-CETAKAN"
    Public Sub CetakanDetail(ByVal StartDate As Date, ByVal EndDate As Date, ByVal PrimarySalesPerson As Double, ByVal PrimaryProduct As Double, ByVal PrimaryCustomer As Double, ByVal PrimaryAccount As Double)
        Dim ReportDetail As New CrpSalesReportByInvoiceDetail
        Dim reportTotal As New CrpSalesReportByInvoiceDetail_SubReport
        Dim DS As New DseLaporanSalesReportByInvoice
        Dim DTAHeader As New DseLaporanSalesReportByInvoiceTableAdapters.TblDetail_HeaderTableAdapter
        Dim DTAItem As New DseLaporanSalesReportByInvoiceTableAdapters.TblDetail_ItemTableAdapter
        Dim TBLHeader As DseLaporanSalesReportByInvoice.TblDetail_HeaderDataTable = DS.TblDetail_Header
        Dim TBLItem As DseLaporanSalesReportByInvoice.TblDetail_ItemDataTable = DS.TblDetail_Item
        Dim TBLItemTemp As New DseLaporanSalesReportByInvoice.TblDetail_ItemDataTable
        Dim TBLCetak As DseLaporanSalesReportByInvoice.TblDetailDataTable = DS.TblDetail
        Dim TBLTotal As DseLaporanSalesReportByInvoice.TBLDetail_TotalDataTable = DS.TBLDetail_Total
        Dim TBLDetail_SubTotal As DseLaporanSalesReportByInvoice.TBLDetail_SubtotalDataTable = DS.TBLDetail_Subtotal
        Dim TBLDetail_SubTotalTemp As New DseLaporanSalesReportByInvoice.TBLDetail_SubtotalDataTable
        Dim rowHeader As DseLaporanSalesReportByInvoice.TblDetail_HeaderRow
        Dim rowItem, rowItemTemp As DseLaporanSalesReportByInvoice.TblDetail_ItemRow
        Dim rowCetak As DseLaporanSalesReportByInvoice.TblDetailRow
        Dim rowTotal As DseLaporanSalesReportByInvoice.TBLDetail_TotalRow
        Dim rowDetail_SubTotal As DseLaporanSalesReportByInvoice.TBLDetail_SubtotalRow
        Dim DefPrec As Double = ObjCurr.GetDecimal_DefaultCurrency
        Dim paramStartDate As Date
        Dim paramEndDate As Date
        Dim PrimaryMax As Double = -1
        Dim PrimaryMin As Double = -1
        Dim PrimaryTransaksi As Double = -1
        Dim ketemu As Boolean = False
        Dim index As Integer = 0
        Dim Filter1 As String = ""
        Dim Filter2 As String = ""
        Dim Filter3 As String = ""
        Dim Filter4 As String = ""
        Dim Filter5 As String = ""
        Dim Filter6 As String = ""
        Dim ListPrimary As New List(Of Double)
        Dim ListPrimCurr As New List(Of Double)
        Dim DV1 As DataRowView
        Dim DV2 As DataRowView
        Dim DTASR As New DseLaporanSalesReportByInvoiceTableAdapters.TblDetail_Item_SRTableAdapter
        Dim TBLSR As New DseLaporanSalesReportByInvoice.TblDetail_Item_SRDataTable

        Try
            '>>>Connection String
            DTAHeader.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTASR.Connection.ConnectionString = Core_Module.GetConstring

            paramStartDate = Me.DtpFrom.Value
            paramEndDate = Me.DtpTo.Value
           
            DTAHeader.Fill(TBLHeader, paramStartDate, paramEndDate)
            DTAItem.Fill(TBLItem, paramStartDate, paramEndDate)

            '>>>Filter
            If PrimaryCustomer <> -1 Then
                Filter1 &= "[Primary Customer] = " & PrimaryCustomer & ""
            End If
            If PrimaryAccount <> -1 Then
                If Filter1 <> "" Then
                    Filter1 &= " AND "
                End If
                Filter1 &= "[primary akun ar] = " & PrimaryAccount & ""
            End If

            If Me.cboPrintStatus.SelectedIndex = 1 Then
                'print only
                If Filter1 <> "" Then
                    Filter1 &= " AND "
                End If
                Filter1 &= "[Printed By] <> ''"
            ElseIf Me.cboPrintStatus.SelectedIndex = 2 Then
                'not print only
                If Filter1 <> "" Then
                    Filter1 &= " AND "
                End If
                Filter1 &= "[Printed By] = ''"
            End If

            '>>>Sort By Primary (Max-MIn)
            TBLHeader.DefaultView.RowFilter = Filter1
            TBLHeader.DefaultView.Sort = "[Primary]"
            If TBLHeader.DefaultView.Count > 0 Then
                rowHeader = TBLHeader.DefaultView(0).Row
                PrimaryMin = rowHeader.Primary
                rowHeader = TBLHeader.DefaultView(TBLHeader.DefaultView.Count - 1).Row
                PrimaryMax = rowHeader.Primary
            End If
            'DTAItem.FillBy_PrimaryMaxnMin(TBLItem, PrimaryMin, PrimaryMax)
            DTASR.Fill(TBLSR, PrimaryMax, PrimaryMin)
            Me.KurangkanItemReturn(TBLItem, TBLSR)

            'filter tblitem by primary customer, primary akun ar, primary sales person
            For i As Integer = TBLItem.Rows.Count - 1 To 0 Step -1
                rowItem = TBLItem.Rows(i)
                If rowItem.Primary_Customer <> PrimaryCustomer And PrimaryCustomer <> -1 Then
                    TBLItem.Rows.Remove(rowItem)
                Else
                    If rowItem.primary_akun_ar <> PrimaryAccount And PrimaryAccount <> -1 Then
                        TBLItem.Rows.Remove(rowItem)
                    Else
                        If rowItem.Primary_Sales_Person <> PrimarySalesPerson And PrimarySalesPerson <> -1 Then
                            TBLItem.Rows.Remove(rowItem)
                        End If
                    End If
                End If
            Next

            'filter tblitem by primary product > group
            Dim FilterProduct As String = Me.FilterProduct()
            If FilterProduct <> "" Then
                TBLItem.DefaultView.RowFilter = FilterProduct
                TBLItemTemp.Merge(TBLItem.DefaultView.ToTable.Copy)
                TBLItem.Rows.Clear()
                TBLItem = TBLItemTemp.Copy
                TBLItemTemp.Rows.Clear()
            End If
            TBLItemTemp = TBLItem.Copy

            TBLHeader.DefaultView.RowFilter = Filter1
            TBLHeader.DefaultView.Sort = "[Primary]"
            TBLItem.DefaultView.RowFilter = Nothing
            TBLItem.DefaultView.Sort = "[Primary Transaksi]"
            If TBLHeader.DefaultView.Count > 0 Then
                For Each DV1 In TBLHeader.DefaultView
                    rowHeader = DV1.Row
                    ListPrimary.Clear()
                    ketemu = False
                    While TBLItem.DefaultView.Count > 0
                        rowItem = TBLItem.DefaultView(0).Row
                        If rowHeader.Primary < rowItem.Primary_Transaksi Then
                            Exit While
                        ElseIf rowHeader.Primary = rowItem.Primary_Transaksi Then
                            If ListPrimary.Contains(rowItem.Primary_Transaksi) = False Then
                                rowDetail_SubTotal = TBLDetail_SubTotal.NewTBLDetail_SubtotalRow
                                TBLDetail_SubTotal.Rows.Add(rowDetail_SubTotal)

                                rowDetail_SubTotal = TBLDetail_SubTotal.Rows(TBLDetail_SubTotal.Rows.Count - 1)
                                rowDetail_SubTotal.Primary_Transaksi = rowItem.Primary_Transaksi
                                rowDetail_SubTotal.Primary_Currency = rowHeader.Primary_Currency
                                rowDetail_SubTotal.DP = rowHeader.Down_Payment * -1
                                rowDetail_SubTotal.Shipping = rowHeader.Shipping
                                ListPrimary.Add(rowItem.Primary_Transaksi)
                            End If

                            rowDetail_SubTotal = TBLDetail_SubTotal.Rows(TBLDetail_SubTotal.Rows.Count - 1)
                            rowDetail_SubTotal.Subtotal += rowItem.Amount
                            If rowHeader.Tax_Exclude = False Then
                                rowDetail_SubTotal.VAT = 0
                            ElseIf rowHeader.Tax_Exclude = True Then
                                If rowHeader.Shipment_With_VAT = 0 Then 'false
                                    rowDetail_SubTotal.VAT = (rowDetail_SubTotal.Subtotal - rowHeader.Down_Payment) * rowHeader.Tax / 100
                                Else
                                    rowDetail_SubTotal.VAT = (rowDetail_SubTotal.Subtotal - rowHeader.Down_Payment + rowHeader.Shipping) * rowHeader.Tax / 100
                                End If
                            End If
                            rowDetail_SubTotal.Total = rowDetail_SubTotal.Subtotal - rowHeader.Down_Payment + rowHeader.Shipping + rowDetail_SubTotal.VAT

                            TBLItem.Rows.Remove(rowItem)
                            ketemu = True
                        Else
                            TBLItem.Rows.Remove(rowItem)
                        End If
                    End While

                    If ketemu = True Then
                        TBLTotal.DefaultView.RowFilter = "[Primary Currency] = " & rowHeader.Primary_Currency.ToString
                        If TBLTotal.DefaultView.Count > 0 Then
                            rowTotal = TBLTotal.DefaultView(0).Row
                            If TBLDetail_SubTotal.Rows.Count > 0 Then
                                rowDetail_SubTotal = TBLDetail_SubTotal.Rows(TBLDetail_SubTotal.Rows.Count - 1)
                                rowTotal.Total_SubTotal += rowDetail_SubTotal.Subtotal
                                rowTotal.Total_VAT += rowDetail_SubTotal.VAT
                                rowTotal.Total_Amount += rowDetail_SubTotal.Total
                            End If
                            rowTotal.Total_Down_Payment += rowHeader.Down_Payment
                            rowTotal.Total_Shipping_Cost += rowHeader.Shipping
                            rowTotal.Total_Stamp_Duty += rowHeader.Duty_Stamp
                        Else
                            rowTotal = TBLTotal.NewTBLDetail_TotalRow
                            rowTotal.Total_Description = "Total " & rowHeader.Currency_Code.ToString
                            rowTotal.Primary_Currency = rowHeader.Primary_Currency
                            rowTotal.Primary_Transaksi = rowHeader.Primary
                            rowTotal.Total_Down_Payment = rowHeader.Down_Payment
                            rowTotal.Total_Shipping_Cost = rowHeader.Shipping
                            rowTotal.Total_Stamp_Duty = rowHeader.Duty_Stamp
                            If TBLDetail_SubTotal.Rows.Count > 0 Then
                                rowDetail_SubTotal = TBLDetail_SubTotal.Rows(TBLDetail_SubTotal.Rows.Count - 1)
                                rowTotal.Total_SubTotal = rowDetail_SubTotal.Subtotal
                                rowTotal.Total_VAT = rowDetail_SubTotal.VAT
                                rowTotal.Total_Amount = rowDetail_SubTotal.Total
                            End If
                            rowTotal.Default_Precision = rowHeader.Default_Precision
                            TBLTotal.Rows.Add(rowTotal)
                        End If
                    End If
                Next
            End If
            '--
            TBLItem.Rows.Clear()
            TBLItem = TBLItemTemp.Copy
            TBLDetail_SubTotalTemp = TBLDetail_SubTotal.Copy
            TBLHeader.DefaultView.RowFilter = Filter1
            TBLHeader.DefaultView.Sort = "[Primary]"
            TBLItem.DefaultView.RowFilter = Nothing
            TBLItem.DefaultView.Sort = "[Primary Transaksi]"
            TBLDetail_SubTotalTemp.DefaultView.RowFilter = ""
            TBLDetail_SubTotalTemp.DefaultView.Sort = "[Primary Transaksi]"
            If TBLHeader.DefaultView.Count > 0 Then
                For Each DV1 In TBLHeader.DefaultView
                    rowHeader = DV1.Row
                    While TBLItem.DefaultView.Count > 0
                        rowItem = TBLItem.DefaultView(0).Row
                        If rowHeader.Primary < rowItem.Primary_Transaksi Then
                            Exit While
                        ElseIf rowHeader.Primary = rowItem.Primary_Transaksi Then
                            rowCetak = TBLCetak.NewTblDetailRow
                            rowCetak.Sales_Person = rowItem.SalesPerson_Desc
                            If rowCetak.Sales_Person = "" Then
                                rowCetak.Sales_Person = "-"
                            End If
                            rowCetak.Curr_Code = rowHeader.Currency_Code
                            rowCetak.Curr_Description = rowHeader.Currency_Desc
                            rowCetak._Date = rowHeader.Tanggal_Transaksi
                            rowCetak.Invoice_Number = rowHeader.Invoice_No
                            rowCetak.Customer_Name = rowHeader.Customer_Name
                            rowCetak.Default_Precision = rowHeader.Default_Precision
                            rowCetak._Decimal = rowItem.Decimal_Multi
                            If IsNumeric(rowHeader.Payment_Term) Then
                                rowCetak.Payment_Term = rowHeader.Payment_Term & " DAYS"
                            Else
                                rowCetak.Payment_Term = rowHeader.Payment_Term
                            End If
                            rowCetak.Default_Precision = rowHeader.Default_Precision
                            rowCetak.Do_Number = rowItem.DO_No
                            rowCetak.SO_Number = rowItem.SO_No
                            Dim RatioStr As String = ""
                            If Me.ChkShowRatio.Checked = True Then
                                RatioStr &= " (" & Me.ObjCD.DigitGrouping(rowItem.Ratio_Default, rowItem._Decimal) & " " & rowItem.Unit_Code_Def & " = " & _
                                Me.ObjCD.DigitGrouping(rowItem.Ratio_Reference, rowItem.Decimal_Multi) & " " & rowItem.Unit_Code & ")"
                            End If
                            If rowItem.Primary_Units <> rowItem.Primary_Units_Multi And rowItem.price_by_multi_unit = 0 Then
                                rowCetak.Qty = rowItem.Total_Qty_Def
                                rowCetak.Unit = rowItem.Unit_Code_Def
                                'If Me.ChkShowRatio.Checked = True Then
                                RatioStr = " (" & Me.ObjCD.DigitGrouping(rowItem.Total_Qty, rowItem.Decimal_Multi) & " " & rowItem.Unit_Code & ")"
                                'End If
                                rowCetak._Decimal = rowItem._Decimal
                            End If
                            rowCetak.Product_Description = rowItem.Product_Desc & RatioStr
                            rowCetak.Unit_Price = rowItem.Nett_Price
                            rowCetak.Total_Price = rowItem.Amount
                            'TBLDetail_SubTotal.DefaultView.RowFilter = "[Primary Transaksi] =" & rowItem.Primary_Transaksi.ToString
                            If TBLDetail_SubTotalTemp.DefaultView.Count > 0 Then
                                While TBLDetail_SubTotalTemp.DefaultView.Count > 0
                                    rowDetail_SubTotal = TBLDetail_SubTotalTemp.DefaultView(0).Row
                                    If rowDetail_SubTotal.Primary_Transaksi > rowItem.Primary_Transaksi Then
                                        Exit While
                                    ElseIf rowDetail_SubTotal.Primary_Transaksi = rowItem.Primary_Transaksi Then
                                        rowCetak.SubTotal = rowDetail_SubTotal.Subtotal
                                        rowCetak.DP = rowDetail_SubTotal.DP
                                        rowCetak.Shipping = rowDetail_SubTotal.Shipping
                                        rowCetak.VAT = rowDetail_SubTotal.VAT
                                        rowCetak.Total = rowDetail_SubTotal.Total
                                        Exit While
                                    Else
                                        TBLDetail_SubTotalTemp.Rows.Remove(rowDetail_SubTotal)
                                    End If
                                End While
                            End If
                            TBLCetak.Rows.Add(rowCetak)

                            TBLItem.Rows.Remove(rowItem)
                        Else
                            TBLItem.Rows.Remove(rowItem)
                        End If
                    End While
                Next
            End If
            '>>>jika no data
            If Me.DtpFrom.Value.Date > Me.DtpTo.Value.Date Then
                MsgBox("There is no data in this period", MsgBoxStyle.Information, "Information")
                Me.Crv.ReportSource = Nothing
                Me.DtpFrom.Focus()
                Exit Sub
            End If
            If TBLCetak.DefaultView.Count <= 0 Then
                MsgBox("There is no data in this period", MsgBoxStyle.Information, "Information")
                Me.Crv.ReportSource = Nothing
                Exit Sub
            End If
            ReportDetail.SetDataSource(DS)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_StartDate.ParameterFieldName, paramStartDate)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_EndDate.ParameterFieldName, paramEndDate)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_FormatDecimal.ParameterFieldName, DefPrec)
            Me.Crv.ReportSource = ReportDetail
            Me.Crv.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTAHeader.Dispose()
            DTAItem.Dispose()
            TBLHeader.Dispose()
            TBLItem.Dispose()
            TBLCetak.Dispose()
            ListPrimary = Nothing
            ListPrimCurr = Nothing
            DTASR.Dispose()
            TBLSR.Dispose()
        End Try
    End Sub

    'Private Function GetJmlSR(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryDO As Double, _
    'ByVal vPrimarySO As Double, ByVal vBatchNumber As String, ByVal vED As Date, _
    ' ByVal vTBLSRItem As DseLaporanSalesReportByInvoice.TblDetail_Item_SRDataTable) As Decimal
    '    Dim rowSR As DseLaporanSalesReportByInvoice.TblDetail_Item_SRRow

    '    Try
    '        Dim row As DataGridViewRow
    '        Dim Jml As Decimal = 0

    '        For Each rowSR In vTBLSRItem.Rows
    '            If vED = #1/1/1950# Then
    '                If IsDate(rowSR.Expiry_Date) = False Then
    '                    If rowSR.Primary_Product = vPrimaryCodeNo And _
    '                         rowSR.Primary_DO = vPrimaryDO And _
    '                         rowSR.Primary_SO = vPrimarySO And _
    '                         rowSR.Batch_No = vBatchNumber Then
    '                        Jml = Jml + CDec(rowSR.Total_Qty_Multi)
    '                    End If
    '                End If
    '            Else
    '                If IsDate(rowSR.Expiry_Date) = True Then
    '                    If rowSR.Primary_Product = vPrimaryCodeNo And _
    '                       rowSR.Primary_DO = vPrimaryDO And _
    '                       rowSR.Primary_SO = vPrimarySO And _
    '                       rowSR.Batch_No = vBatchNumber And _
    '                       CDate(rowSR.Expiry_Date) = vED Then
    '                        Jml = Jml + CDec(rowSR.Total_Qty_Multi)
    '                    End If
    '                End If
    '            End If
    '        Next

    '        Return Jml
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    Private Sub KurangkanItemReturn(ByVal vTBLItem As DseLaporanSalesReportByInvoice.TblDetail_ItemDataTable, ByVal vTBLSR As DseLaporanSalesReportByInvoice.TblDetail_Item_SRDataTable)
        Dim rowSR As DseLaporanSalesReportByInvoice.TblDetail_Item_SRRow
        Dim rowItem As DseLaporanSalesReportByInvoice.TblDetail_ItemRow
        Dim DVSR, DVItem As Data.DataRowView

        Try
            vTBLSR.DefaultView.Sort = "[Primary Sales Invoice]"
            vTBLItem.DefaultView.Sort = "[Primary Transaksi]"

            For Each DVSR In vTBLSR.DefaultView
                rowSR = DVSR.Row

                For Each DVItem In vTBLItem.DefaultView
                    rowItem = DVItem.Row
                    If rowItem.Primary_Transaksi > rowSR.Primary_Sales_Invoice Then
                        Exit For
                    Else
                        If rowSR.Primary_Product = rowItem.Primary_Product And _
                           rowSR.Primary_DO = rowItem.Primary_DO And _
                           rowSR.Primary_SO = rowItem.Primary_SO And _
                           rowSR.Batch_No = rowItem.Batch_Number And _
                           rowSR.Primary_Sales_Invoice = rowItem.Primary_Transaksi Then
                            rowItem.Total_Qty -= rowSR.Total_Qty_Multi
                            If rowItem.Total_Qty = 0 Then
                                vTBLItem.Rows.Remove(rowItem)
                            Else
                                rowItem.Amount = Me.ObjCD.SetDecimal(rowItem.Total_Qty * rowItem.Nett_Price, rowItem.Default_Precision)
                            End If
                            Exit For
                        End If
                    End If
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Enum EnumGroupLevel As Integer
        ProductGroup
        Brand
        SubBrand
        Category
        Product
    End Enum

    Public Function GetLevel_Str(ByVal vIdx As EnumGroupLevel) As String
        Try
            If vIdx = EnumGroupLevel.ProductGroup Then
                Return "Product Group"
            ElseIf vIdx = EnumGroupLevel.Brand Then
                Return "Brand"
            ElseIf vIdx = EnumGroupLevel.SubBrand Then
                Return "Sub Brand"
            ElseIf vIdx = EnumGroupLevel.Category Then
                Return "Category"
            ElseIf vIdx = EnumGroupLevel.Product Then
                Return "Product"
            End If

            Return "Product"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "Product"
        End Try
    End Function

    Public Function GetLevelInt(ByVal vLevelStr As String) As EnumGroupLevel
        Try
            If vLevelStr = "Product Group" Then
                Return EnumGroupLevel.ProductGroup
            ElseIf vLevelStr = "Brand" Then
                Return EnumGroupLevel.Brand
            ElseIf vLevelStr = "Sub Brand" Then
                Return EnumGroupLevel.SubBrand
            ElseIf vLevelStr = "Category" Then
                Return EnumGroupLevel.Category
            ElseIf vLevelStr = "Product" Then
                Return EnumGroupLevel.Product
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 4
        End Try
    End Function

    Dim ObjSettingProd As New Core_Module.ClsTransaksiSetting_ProdGroup
    Dim TipeProductGroup As Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup
    Private Function FilterProduct() As String
        Dim FilterGroup As String = ""

        Try
            For Each rowDGV As DataGridViewRow In Me.dgvItemProductGroup.Rows
                If rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value <> -1 Then
                    TipeProductGroup = Me.ObjSettingProd.GetLevelInt(rowDGV.Cells(Me.ColProductGroup_Level.Index).Value)
                    If FilterGroup = "" Then
                        If TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup Then
                            FilterGroup &= " [Primary Product Group] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
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
                            FilterGroup &= " OR [Primary Product Group] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
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

            'If FilterGroup <> "" Then
            '    FilterProduct &= " AND (" & FilterGroup & ") "
            'End If
            Return FilterGroup
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub CetakanRekap(ByVal StartDate As Date, ByVal EndDate As Date, ByVal PrimarySalesPerson As Double, ByVal PrimaryProduct As Double, ByVal PrimaryCustomer As Double, _
    ByVal PrimaryAccount As Double)
        Dim ReportDetail As New CrpSalesReportByInvoiceRekap
        Dim DS As New DseLaporanSalesReportByInvoice
        Dim DTAHeader As New DseLaporanSalesReportByInvoiceTableAdapters.TblDetail_HeaderTableAdapter
        Dim DTAItem As New DseLaporanSalesReportByInvoiceTableAdapters.TblDetail_ItemTableAdapter
        Dim TBLHeader As DseLaporanSalesReportByInvoice.TblDetail_HeaderDataTable = DS.TblDetail_Header
        Dim TBLItem As DseLaporanSalesReportByInvoice.TblDetail_ItemDataTable = DS.TblDetail_Item
        Dim TBLItemTemp As New DseLaporanSalesReportByInvoice.TblDetail_ItemDataTable
        Dim TBLCetak As DseLaporanSalesReportByInvoice.TblRekapDataTable = DS.TblRekap
        Dim TBLTotal As DseLaporanSalesReportByInvoice.TBLDetail_TotalDataTable = DS.TBLDetail_Total
        Dim TBLDetail_SubTotal As DseLaporanSalesReportByInvoice.TBLDetail_SubtotalDataTable = DS.TBLDetail_Subtotal
        Dim rowHeader As DseLaporanSalesReportByInvoice.TblDetail_HeaderRow
        Dim rowItem As DseLaporanSalesReportByInvoice.TblDetail_ItemRow
        Dim rowCetak As DseLaporanSalesReportByInvoice.TblRekapRow
        Dim rowTotal As DseLaporanSalesReportByInvoice.TBLDetail_TotalRow
        Dim rowDetail_SubTotal As DseLaporanSalesReportByInvoice.TBLDetail_SubtotalRow
        Dim DefPrec As Double = ObjCurr.GetDecimal_DefaultCurrency
        Dim paramStartDate As Date
        Dim paramEndDate As Date
        Dim PrimaryMax As Double = -1
        Dim PrimaryMin As Double = -1
        Dim PrimaryTransaksi As Double = -1
        Dim ketemu As Boolean = False
        Dim index As Integer = 0
        Dim Filter1 As String = ""
        Dim Filter2 As String = ""
        Dim Filter3 As String = ""
        Dim Filter4 As String = ""
        Dim Filter5 As String = ""
        Dim Filter6 As String = ""
        Dim Filter7 As String = ""
        Dim ListPrimary As New List(Of Double)
        Dim ListPrimCurr As New List(Of Double)
        Dim ListNoinvoice As New List(Of Double)
        Dim LIstForInvoice As New List(Of Double)
        Dim DV1 As DataRowView
        Dim DV2 As DataRowView
        Dim DV3 As DataRowView
        Dim DTASR As New DseLaporanSalesReportByInvoiceTableAdapters.TblDetail_Item_SRTableAdapter
        Dim TBLSR As New DseLaporanSalesReportByInvoice.TblDetail_Item_SRDataTable

        Try
            '>>>Connection String
            DTAHeader.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTASR.Connection.ConnectionString = Core_Module.GetConstring

            paramStartDate = Me.DtpFrom.Value
            paramEndDate = Me.DtpTo.Value

            DTAHeader.Fill(TBLHeader, paramStartDate, paramEndDate)
            DTAItem.Fill(TBLItem, paramStartDate, paramEndDate)

            '>>>Filter
            If PrimaryCustomer <> -1 Then
                Filter1 &= "[Primary Customer] = " & PrimaryCustomer & ""
            End If
            If PrimaryAccount <> -1 Then
                If Filter1 <> "" Then
                    Filter1 &= " AND "
                End If
                Filter1 &= "[primary akun ar] = " & PrimaryAccount & ""
            End If

            If Me.cboPrintStatus.SelectedIndex = 1 Then
                'print only
                If Filter1 <> "" Then
                    Filter1 &= " AND "
                End If
                Filter1 &= "[Printed By] <> ''"
            ElseIf Me.cboPrintStatus.SelectedIndex = 2 Then
                'not print only
                If Filter1 <> "" Then
                    Filter1 &= " AND "
                End If
                Filter1 &= "[Printed By] = ''"
            End If
            '--
            '>>>Urutkan berdasarkan Primary & Fill Max-Min
            TBLHeader.DefaultView.RowFilter = Filter1
            TBLHeader.DefaultView.Sort = "[Primary]"
            If TBLHeader.DefaultView.Count > 0 Then
                rowHeader = TBLHeader.DefaultView(0).Row
                PrimaryMin = rowHeader.Primary
                rowHeader = TBLHeader.DefaultView(TBLHeader.DefaultView.Count - 1).Row
                PrimaryMax = rowHeader.Primary
            End If
            DTASR.Fill(TBLSR, PrimaryMax, PrimaryMin)
            Me.KurangkanItemReturn(TBLItem, TBLSR)

            'filter tblitem by primary customer, primary akun ar, primary sales person
            For i As Integer = TBLItem.Rows.Count - 1 To 0 Step -1
                rowItem = TBLItem.Rows(i)
                If rowItem.Primary_Customer <> PrimaryCustomer And PrimaryCustomer <> -1 Then
                    TBLItem.Rows.Remove(rowItem)
                Else
                    If rowItem.primary_akun_ar <> PrimaryAccount And PrimaryAccount <> -1 Then
                        TBLItem.Rows.Remove(rowItem)
                    Else
                        If rowItem.Primary_Sales_Person <> PrimarySalesPerson And PrimarySalesPerson <> -1 Then
                            TBLItem.Rows.Remove(rowItem)
                        End If
                    End If
                End If
            Next

            'filter tblitem by primary product > group
            Dim FilterProduct As String = Me.FilterProduct()
            If FilterProduct <> "" Then
                TBLItem.DefaultView.RowFilter = FilterProduct
                TBLItemTemp.Merge(TBLItem.DefaultView.ToTable.Copy)
                TBLItem.Rows.Clear()
                TBLItem = TBLItemTemp.Copy
                TBLItemTemp.Rows.Clear()
            End If
            TBLItemTemp = TBLItem.Copy
            Dim CekTotal As Decimal = 0
            TBLHeader.DefaultView.RowFilter = Filter1
            TBLHeader.DefaultView.Sort = "[Primary]"
            TBLItem.DefaultView.RowFilter = Nothing
            TBLItem.DefaultView.Sort = "[Primary Transaksi]"
            If TBLHeader.DefaultView.Count > 0 Then
                For Each DV1 In TBLHeader.DefaultView
                    rowHeader = DV1.Row
                    ListPrimary.Clear()

                    ketemu = False
                    If TBLItem.DefaultView.Count > 0 Then
                        While TBLItem.DefaultView.Count > 0
                            rowItem = TBLItem.DefaultView(0).Row
                            If rowHeader.Primary < rowItem.Primary_Transaksi Then
                                Exit While
                            ElseIf rowHeader.Primary = rowItem.Primary_Transaksi Then
                                If ListPrimary.Contains(rowItem.Primary_Transaksi) = False Then
                                    rowDetail_SubTotal = TBLDetail_SubTotal.NewTBLDetail_SubtotalRow
                                    TBLDetail_SubTotal.Rows.Add(rowDetail_SubTotal)

                                    rowDetail_SubTotal = TBLDetail_SubTotal.Rows(TBLDetail_SubTotal.Rows.Count - 1)
                                    rowDetail_SubTotal.Primary_Transaksi = rowItem.Primary_Transaksi
                                    rowDetail_SubTotal.Primary_Currency = rowHeader.Primary_Currency
                                    rowDetail_SubTotal.DP = rowHeader.Down_Payment * -1
                                    rowDetail_SubTotal.Shipping = rowHeader.Shipping
                                    ListPrimary.Add(rowItem.Primary_Transaksi)
                                End If

                                rowDetail_SubTotal = TBLDetail_SubTotal.Rows(TBLDetail_SubTotal.Rows.Count - 1)
                                rowDetail_SubTotal.Subtotal += rowItem.Amount
                                If rowHeader.Tax_Exclude = False Then
                                    rowDetail_SubTotal.VAT = 0
                                ElseIf rowHeader.Tax_Exclude = True Then
                                    If rowHeader.Shipment_With_VAT = 0 Then 'false
                                        rowDetail_SubTotal.VAT = (rowDetail_SubTotal.Subtotal - rowHeader.Down_Payment) * rowHeader.Tax / 100
                                    Else
                                        rowDetail_SubTotal.VAT = (rowDetail_SubTotal.Subtotal - rowHeader.Down_Payment + rowHeader.Shipping) * rowHeader.Tax / 100
                                    End If
                                End If
                                rowDetail_SubTotal.Total = rowDetail_SubTotal.Subtotal - rowHeader.Down_Payment + rowHeader.Shipping + rowDetail_SubTotal.VAT

                                TBLItem.Rows.Remove(rowItem)
                                Ketemu = True
                            Else
                                TBLItem.Rows.Remove(rowItem)
                            End If
                        End While
                    End If

                    '>>>Total Per Currency
                    If ketemu = True Then
                        TBLTotal.DefaultView.RowFilter = "[Primary Currency] = " & rowHeader.Primary_Currency.ToString
                        If TBLTotal.DefaultView.Count > 0 Then
                            rowTotal = TBLTotal.DefaultView(0).Row
                            If TBLDetail_SubTotal.Rows.Count > 0 Then
                                rowDetail_SubTotal = TBLDetail_SubTotal.Rows(TBLDetail_SubTotal.Rows.Count - 1)
                                rowTotal.Total_SubTotal += rowDetail_SubTotal.Subtotal
                                rowTotal.Total_VAT += rowDetail_SubTotal.VAT
                                rowTotal.Total_Amount += rowDetail_SubTotal.Total
                            End If
                            rowTotal.Total_Down_Payment += rowHeader.Down_Payment
                            rowTotal.Total_Shipping_Cost += rowHeader.Shipping
                            rowTotal.Total_Stamp_Duty += rowHeader.Duty_Stamp
                        Else
                            rowTotal = TBLTotal.NewTBLDetail_TotalRow
                            rowTotal.Total_Description = "Total " & rowHeader.Currency_Code.ToString
                            rowTotal.Primary_Currency = rowHeader.Primary_Currency
                            rowTotal.Primary_Transaksi = rowHeader.Primary
                            rowTotal.Total_Down_Payment = rowHeader.Down_Payment
                            rowTotal.Total_Shipping_Cost = rowHeader.Shipping
                            rowTotal.Total_Stamp_Duty = rowHeader.Duty_Stamp
                            If TBLDetail_SubTotal.Rows.Count > 0 Then
                                rowDetail_SubTotal = TBLDetail_SubTotal.Rows(TBLDetail_SubTotal.Rows.Count - 1)
                                rowTotal.Total_SubTotal = rowDetail_SubTotal.Subtotal
                                rowTotal.Total_VAT = rowDetail_SubTotal.VAT
                                rowTotal.Total_Amount = rowDetail_SubTotal.Total
                            End If
                            rowTotal.Default_Precision = rowHeader.Default_Precision
                            TBLTotal.Rows.Add(rowTotal)
                        End If
                    End If
                Next
            End If
            '--
            TBLItem.Rows.Clear()
            TBLItem = TBLItemTemp.Copy
            TBLHeader.DefaultView.RowFilter = Filter1
            TBLHeader.DefaultView.Sort = "[Primary]"
            TBLItem.DefaultView.RowFilter = Nothing
            TBLItem.DefaultView.Sort = "[Primary Transaksi]"
            TBLDetail_SubTotal.DefaultView.RowFilter = ""
            TBLDetail_SubTotal.DefaultView.Sort = "[Primary Transaksi]"
            If TBLHeader.DefaultView.Count > 0 Then
                For Each DV1 In TBLHeader.DefaultView
                    rowHeader = DV1.Row
                    While TBLItem.DefaultView.Count > 0
                        rowItem = TBLItem.DefaultView(0).Row
                        If rowHeader.Primary < rowItem.Primary_Transaksi Then
                            Exit While
                        ElseIf rowHeader.Primary = rowItem.Primary_Transaksi Then
                            rowCetak = TBLCetak.NewTblRekapRow
                            rowCetak.Primary = rowHeader.Primary
                            rowCetak.Curr_Code = rowHeader.Currency_Code
                            rowCetak.Curr_Description = rowHeader.Currency_Desc
                            rowCetak._Date = rowHeader.Tanggal_Transaksi
                            rowCetak.Invoice_Number = rowHeader.Invoice_No
                            rowCetak.Default_Precision = rowHeader.Default_Precision
                            rowCetak.Customer = rowHeader.Customer_Name
                            If IsNumeric(rowHeader.Payment_Term) Then
                                rowCetak.Payment_Term = rowHeader.Payment_Term & " DAYS"
                            Else
                                rowCetak.Payment_Term = rowHeader.Payment_Term
                            End If
                            rowCetak.Stamp_Duty = rowHeader.Duty_Stamp
                            '--
                            'TBLDetail_SubTotal.DefaultView.RowFilter = "[Primary Transaksi] = " & rowItem.Primary_Transaksi
                            If TBLDetail_SubTotal.DefaultView.Count > 0 Then
                                While TBLDetail_SubTotal.DefaultView.Count > 0
                                    rowDetail_SubTotal = TBLDetail_SubTotal.DefaultView(0).Row
                                    If rowDetail_SubTotal.Primary_Transaksi > rowItem.Primary_Transaksi Then
                                        Exit While
                                    ElseIf rowDetail_SubTotal.Primary_Transaksi = rowItem.Primary_Transaksi Then
                                        rowCetak.Sub_Total = rowDetail_SubTotal.Subtotal
                                        rowCetak.DP = rowDetail_SubTotal.DP
                                        rowCetak.Shipping_Cost = rowDetail_SubTotal.Shipping
                                        rowCetak.VAT = rowDetail_SubTotal.VAT
                                        rowCetak.Total = rowDetail_SubTotal.Total
                                        Exit While
                                    Else
                                        TBLDetail_SubTotal.Rows.Remove(rowDetail_SubTotal)
                                    End If
                                End While
                            End If

                            rowCetak.Sales_Person = rowItem.SalesPerson_Desc
                            'TBLItem.DefaultView.RowFilter = "[Primary Sales Person] = " & PrimarySalesPerson.ToString
                            'If TBLItem.DefaultView.Count > 0 Then
                            '    For Each DV3 In TBLItem.DefaultView
                            '        rowItem = DV3.Row
                            '        '>>>Dari item
                            '        rowCetak.Sales_Person = rowItem.SalesPerson_Desc
                            '    Next
                            'End If
                            rowCetak.Down_Payment = rowHeader.Down_Payment
                            TBLCetak.Rows.Add(rowCetak)

                            TBLItem.Rows.Remove(rowItem)
                        Else
                            TBLItem.Rows.Remove(rowItem)
                        End If
                    End While
                Next
            End If
            '>>>jika no data
            If Me.DtpFrom.Value.Date > Me.DtpTo.Value.Date Then
                MsgBox("There is no data in this period", MsgBoxStyle.Information, "Information")
                Me.Crv.ReportSource = Nothing
                Me.DtpFrom.Focus()
                Exit Sub
            End If
            If TBLCetak.DefaultView.Count <= 0 Then
                MsgBox("There is no data in this period", MsgBoxStyle.Information, "Information")
                Me.Crv.ReportSource = Nothing
                Exit Sub
            End If
            ReportDetail.SetDataSource(DS)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_StartDate.ParameterFieldName, paramStartDate)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_EndDate.ParameterFieldName, paramEndDate)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_FormatDecimal.ParameterFieldName, DefPrec)
            Me.Crv.ReportSource = ReportDetail
            Me.Crv.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTAHeader.Dispose()
            DTAItem.Dispose()
            TBLHeader.Dispose()
            TBLItem.Dispose()
            TBLCetak.Dispose()
            ListPrimary = Nothing
            ListPrimCurr = Nothing
            DTASR.Dispose()
            TBLSR.Dispose()
        End Try
    End Sub
#End Region

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPrint.Click
        Try
            If Me.CboFormat.SelectedIndex = 0 Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                CetakanRekap(Me.DtpFrom.Value, Me.DtpTo.Value, CDbl(Me.TxtPrimarySales.Text), CDbl(Me.txtPrimaryCodeNo.Text), CDbl(Me.TxtPrimaryCustomer.Text), CDbl(Me.TxtPrimaryAcc.Text))
                ChkAllProduct.Enabled = True
            Else
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                CetakanDetail(Me.DtpFrom.Value, Me.DtpTo.Value, CDbl(Me.TxtPrimarySales.Text), CDbl(Me.txtPrimaryCodeNo.Text), CDbl(Me.TxtPrimaryCustomer.Text), CDbl(Me.TxtPrimaryAcc.Text))
                ChkAllProduct.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmSalesReport_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.CboFormat.SelectedIndex = 0
            Me.DtpFrom.Value = Now
            Me.DtpTo.Value = Now
            Me.CboFormat.Focus()
            Me.cboPrintStatus.SelectedIndex = 0
            Me.MinimizeBox = True
            Me.dgvItemProductGroup.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.ColProductGroup_Remarks.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkAllProduct_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkAllProduct.CheckedChanged
        Try
            Crv.ReportSource = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkAllProduct_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkAllProduct.Click
        Try
            If txtPrimaryCodeNo.Text = -1 Then
                Me.ChkAllProduct.Checked = False
                MsgBox("Product cannot be an empty string", MsgBoxStyle.Information, "Information")
                Me.txtCodeNo.Focus()
            End If
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

    Private Sub dgvItemProductGroup_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItemProductGroup.DataError
        'do nothing -> dibuat untuk handle data error event, jangan dihapuskan
    End Sub
End Class