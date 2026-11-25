Option Explicit On
Option Strict Off

Public Class FrmLaporanSalesOrderBalance
    Public UserCode As String = ""
    Public BackColorAll As Color = BackColor
    Public FormNo As Integer = 0
    Dim FormatDay As Integer = 1
    Dim FormatMonth As Integer = 3
    Dim FormatYear As Integer = 1
    Dim FormatQty As Integer = 0
    Dim TBLPendingDO1 As New DseTransaksiDO.pendingdo_view1DataTable
    Dim TBLPendingDO2 As New DseTransaksiDO.pendingdo_view2DataTable
    Dim TBLPendingDO3 As New DseTransaksiDO.pendingdo_view3DataTable
    Dim TBLPendingDO4 As New DseTransaksiDO.pendingdo_view4DataTable
    Dim ObjLicense As New Core_Module.ClsLicense

    Private Sub SetPendingDO(ByVal TBLCetak As dseLaporanSalesOrderBalance2.TblCetakDataTable)
        Dim rowCetak As dseLaporanSalesOrderBalance2.TblCetakRow

        Try
            For Each rowCetak In TBLCetak.Rows

            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetTotalSOB(ByVal vPrimaryProduct As Double, ByVal vTBL As _
    DseLaporanSalesOrderBalance.DataTable13DataTable) As Double
        Dim cmd As String = ""
        Dim row As DseLaporanSalesOrderBalance.DataTable13Row
        Dim rowView As DataRowView
        Dim Retval As Double = 0

        Try
            cmd = "[Primary Product] =" + vPrimaryProduct.ToString
            vTBL.DefaultView.RowFilter = cmd
            If vTBL.DefaultView.Count > 0 Then
                For Each rowView In vTBL.DefaultView
                    row = rowView.Row
                    Retval += row.SOB
                Next
            End If

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub Cetak4()
        Dim cmd As String = ""
        Dim report As New CrpLaporanSalesOrderBalance_3

        Dim ds As New DseLaporanSalesOrderBalance
        Dim dta As New DseLaporanSalesOrderBalanceTableAdapters.DataTable13TableAdapter
        Dim TBL As DseLaporanSalesOrderBalance.DataTable13DataTable = ds.DataTable13
        Dim row, rowDummy As DseLaporanSalesOrderBalance.DataTable13Row

        'tambahan aul
        Dim dtaPOB As New DseLaporanSalesOrderBalanceTableAdapters.tbl_POBTableAdapter
        Dim TBLPOB As DseLaporanSalesOrderBalance.tbl_POBDataTable = ds.tbl_POB
        Dim rowPOB As DseLaporanSalesOrderBalance.tbl_POBRow

        Dim DTA_BookingGIT As New DseLaporanSalesOrderBalanceTableAdapters.tbltrans_booking_gititem3TableAdapter
        Dim TBLBookingGIT As DseLaporanSalesOrderBalance.tbltrans_booking_gititem3DataTable = ds.tbltrans_booking_gititem3
        Dim DTA_BookingInv As New DseLaporanSalesOrderBalanceTableAdapters.tbltrans_booking_invitem1TableAdapter
        Dim TBL_BookingInv As DseLaporanSalesOrderBalance.tbltrans_booking_invitem1DataTable = ds.tbltrans_booking_invitem1
        Dim row_BookingInv As DseLaporanSalesOrderBalance.tbltrans_booking_invitem1Row
        Dim row_BookingGIT As DseLaporanSalesOrderBalance.tbltrans_booking_gititem3Row

        Dim DTA_Stock As New DseLaporanSalesOrderBalanceTableAdapters.DataTable10TableAdapter
        Dim TBL_Stock As DseLaporanSalesOrderBalance.DataTable10DataTable = ds.DataTable10
        Dim row_Stock As DseLaporanSalesOrderBalance.DataTable10Row
        Dim JmlBookingInv, JmlBookingGIT, JmlTotalSOB As Double
        Dim i As Integer = 0
        Dim PrimaryProductMin As Double = -1
        Dim PrimaryProductMax As Double = -1

        Windows.Forms.Cursor.Current = Cursors.WaitCursor

        Try
            'isikan data order balance
            If Me.chkFilter.Checked = True Then
                If Me.txtPrimaryCustomer.Text = "-1" Or Me.txtPrimaryCustomer.Text = "" Then
                    dta.FillBy_TanggalSampaiDari(TBL, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
                Else
                    dta.FillBy_CustCode_TanggalSampaiDari(TBL, Me.txtCusomerCode.Text, _
                    Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
                End If
            Else
                If Me.txtPrimaryCustomer.Text = "-1" Or Me.txtPrimaryCustomer.Text = "" Then
                    dta.Fill(TBL)
                Else
                    dta.FillBy_CustCode(TBL, Me.txtCusomerCode.Text)
                End If
            End If

            'dapatkan primary product min
            TBL.DefaultView.Sort = Nothing
            TBL.DefaultView.Sort = "[Primary Product]"
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                PrimaryProductMin = row.Primary_Product
            End If

            'dapatkan primary product max
            TBL.DefaultView.Sort = Nothing
            TBL.DefaultView.Sort = "[Primary Product] DESC"
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                PrimaryProductMax = row.Primary_Product
            End If

            'isikan data POB
            dtaPOB.FillBy_PrimaryProductMinMax(TBLPOB, PrimaryProductMin, PrimaryProductMax)

            'isikan data available to deliver
            DTA_Stock.Fill(TBL_Stock)

            'isikan data booking GIT
            DTA_BookingGIT.Fill(TBLBookingGIT)

            'isikan data Booking Inv
            DTA_BookingInv.Fill(TBL_BookingInv)

            'hitung Total Qty Sales Order Balance
            For Each row In TBL.Rows
                If row.IsTotal_Sales_Order_BalanceNull = True Then

                    'tambahan aul
                    TBLPOB.DefaultView.RowFilter = Nothing
                    TBLPOB.DefaultView.RowFilter = "[Primary Product] = " & row.Primary_Product
                    If TBLPOB.DefaultView.Count > 0 Then
                        rowPOB = TBLPOB.DefaultView(0).Row
                        row.Total_POB = rowPOB.POB
                    Else
                        row.Total_POB = 0
                    End If
                    '-----

                    JmlTotalSOB = Me.GetTotalSOB(row.Primary_Product, TBL)
                    row.Total_Sales_Order_Balance = JmlTotalSOB

                    For j As Integer = i + 1 To TBL.Rows.Count - 1
                        rowDummy = TryCast(TBL.Rows(j), DseLaporanSalesOrderBalance.DataTable13Row)
                        If rowDummy.Primary_Product = row.Primary_Product Then
                            row.Total_Sales_Order_Balance = JmlTotalSOB
                        End If
                    Next
                End If

                i += 1
            Next

            'hitung total qty booking, qty pending do, qty on hand
            For Each row In ds.DataTable13.Rows
                'isikan data available to deliver
                cmd = "[Product Code]='" + row.Product_Code + "'"
                TBL_Stock.DefaultView.RowFilter = cmd
                If TBL_Stock.DefaultView.Count > 0 Then
                    row_Stock = TryCast(TBL_Stock.DefaultView(0).Row, DseLaporanSalesOrderBalance.DataTable10Row)
                    row.Qty_Available = row_Stock.TotalOnHand - row_Stock.QtyBookingInv - row_Stock.QtyPendingDO
                    row.Qty_On_Hand = row_Stock.TotalOnHand
                Else
                    row.Qty_Available = 0
                    row.Qty_On_Hand = 0
                End If
                'isikan data booking
                cmd = "[Primary Product] = " + row.Primary_Product.ToString
                TBL_BookingInv.DefaultView.RowFilter = cmd
                If TBL_BookingInv.DefaultView.Count > 0 Then
                    row_BookingInv = TryCast(TBL_BookingInv.DefaultView(0).Row, DseLaporanSalesOrderBalance.tbltrans_booking_invitem1Row)
                    JmlBookingInv = row_BookingInv.Qty_Booking_Inv
                Else
                    JmlBookingInv = 0
                End If
                TBLBookingGIT.DefaultView.RowFilter = cmd
                If TBLBookingGIT.DefaultView.Count > 0 Then
                    row_BookingGIT = TryCast(TBLBookingGIT.DefaultView(0).Row, DseLaporanSalesOrderBalance.tbltrans_booking_gititem3Row)
                    JmlBookingGIT = row_BookingGIT.Qty_Booking_GIT
                Else
                    JmlBookingGIT = 0
                End If
                row.Qty_Booking = JmlBookingGIT + JmlBookingInv
            Next

            report.SetDataSource(ds)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
            report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, Me.FormatYear)
            report.SetParameterValue(report.Parameter_FormatQty.ParameterFieldName, Me.FormatQty)
            report.SetParameterValue(report.Parameter_TglDari.ParameterFieldName, Me.dtpFrom.Value)
            report.SetParameterValue(report.Parameter_TglSampai.ParameterFieldName, Me.dtpTo.Value)
            Me.crv1.ReportSource = report
            Me.crv1.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ds.Dispose()
            dta.Dispose()
            DTA_BookingGIT.Dispose()
            DTA_BookingInv.Dispose()
            DTA_Stock.Dispose()
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnFindCustomer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCustomer.Click
        Dim i As New Core_Module.FrmCariMasterCustomer
        Try
            i.BackColorAll = Me.BackColorAll
            i.txtCari.Text = Me.txtCusomerCode.Text
            i.ShowDialog()
            If i.PrimaryCustomer <> -1 Then
                Dim objCust As New Core_Module.clsMasterCustomer
                Me.txtPrimaryCustomer.Text = i.PrimaryCustomer
                Me.txtCusomerCode.Text = objCust.GetCustomerCode(i.PrimaryCustomer)
                Me.txtCustomerName.Text = objCust.GetCustName(i.PrimaryCustomer)
                objCust = Nothing
            Else
                Me.txtCusomerCode.Text = ""
                Me.txtCustomerName.Clear()
                Me.txtPrimaryCustomer.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtCusomerCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCusomerCode.GotFocus
        txtCusomerCode.BackColor = Me.BackColorAll
    End Sub

    Private Sub txtCusomerCode_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCusomerCode.KeyDown
        Dim objCust As New Core_Module.clsMasterCustomer
        Try
            If e.KeyCode = Keys.Enter Then
                Dim primary As Double = objCust.GetPrimary(Me.txtCusomerCode.Text)
                If primary <> -1 Then
                    Me.txtPrimaryCustomer.Text = primary
                    Me.txtCusomerCode.Text = objCust.GetCustomerCode(primary)
                    Me.txtCustomerName.Text = objCust.GetCustName(primary)
                Else
                    Me.btnFindCustomer.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCust = Nothing
        End Try
    End Sub

    Private Sub txtCusomerCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCusomerCode.Leave
        Dim objCust As New Core_Module.clsMasterCustomer
        Try
            If Me.txtCusomerCode.Text <> "" Then
                Dim primary As Double = objCust.GetPrimary(Me.txtCusomerCode.Text)
                If primary <> -1 Then
                    Me.txtPrimaryCustomer.Text = primary
                    Me.txtCusomerCode.Text = objCust.GetCustomerCode(primary)
                    Me.txtCustomerName.Text = objCust.GetCustName(primary)
                Else
                    Me.btnFindCustomer.PerformClick()
                End If
            Else
                Me.txtCustomerName.Clear()
                Me.txtPrimaryCustomer.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCust = Nothing
        End Try
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

            If Me.rboGroupBySalesOrder.Checked = True Then
                Me.Cetak_3()
            Else
                Me.Cetak4()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub txtCusomerCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCusomerCode.LostFocus
        Me.txtCusomerCode.BackColor = Me.txtPrimaryCustomer.BackColor
    End Sub

    Private Sub Cetak_3()
        Dim report As New CrpLaporanSalesOrderBalance
        Dim DS As New DseLaporanSalesOrderBalance
        Dim TBLIpo As New DseLaporanSalesOrderBalance.DataTable2DataTable
        Dim rowIpo As DseLaporanSalesOrderBalance.DataTable2Row
        Dim DTA As New DseLaporanSalesOrderBalanceTableAdapters.DataTable9TableAdapter
        Dim row As DseLaporanSalesOrderBalance.DataTable9Row

        Dim DtaPOB As New DseLaporanSalesOrderBalanceTableAdapters.tbl_POBTableAdapter
        Dim TblPOB As DseLaporanSalesOrderBalance.tbl_POBDataTable = DS.tbl_POB
        Dim rowPOB As DseLaporanSalesOrderBalance.tbl_POBRow

        Dim row_View As DataRowView
        Dim DTA_BookingGIT As New DseLaporanSalesOrderBalanceTableAdapters.tbltrans_booking_gititem2TableAdapter
        Dim TBLBookingGIT As DseLaporanSalesOrderBalance.tbltrans_booking_gititem2DataTable = DS.tbltrans_booking_gititem2
        Dim DTA_BookingInv As New DseLaporanSalesOrderBalanceTableAdapters.DataTable11TableAdapter
        Dim TBL_BookingInv As DseLaporanSalesOrderBalance.DataTable11DataTable = DS.DataTable11
        Dim row_BookingInv As DseLaporanSalesOrderBalance.DataTable11Row
        Dim row_BookingGIT As DseLaporanSalesOrderBalance.tbltrans_booking_gititem2Row
        Dim DTA_Stock As New DseLaporanSalesOrderBalanceTableAdapters.DataTable10TableAdapter
        Dim TBL_Stock As DseLaporanSalesOrderBalance.DataTable10DataTable = DS.DataTable10
        Dim row_Stock As DseLaporanSalesOrderBalance.DataTable10Row
        Dim JmlBookingInv, JmlBookingGIT As Double

        Dim PrimaryProductMin As Double = -1
        Dim PrimaryProductMax As Double = -1

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'isikan data IPO
            Dim DTAIPO As New DseLaporanSalesOrderBalanceTableAdapters.DataTable2TableAdapter
            If Me.txtCusomerCode.Text = "" Then
                DTAIPO.FillBy_TanggalDariSampai(TBLIpo, Me.dtpFrom.Value, Me.dtpTo.Value)
            Else
                DTAIPO.FillBy_TglDariSampaiNCustomerCode(TBLIpo, Me.dtpFrom.Value, Me.dtpTo.Value, Me.txtCusomerCode.Text)
            End If
            DTAIPO.Dispose()

            'isikan data available to deliver
            DTA_Stock.Fill(TBL_Stock)

            'isikan data booking GIT
            DTA_BookingGIT.Fill(TBLBookingGIT)

            'isikan data Booking Inv
            DTA_BookingInv.Fill(TBL_BookingInv)

            'isikan data order balance
            If Me.chkFilter.Checked = False Then
                If Me.txtCusomerCode.Text <> "" Then
                    DTA.FillBy_CustCode(DS.DataTable9, Me.txtCusomerCode.Text)
                Else
                    DTA.Fill(DS.DataTable9)
                End If
            Else
                If Me.txtCusomerCode.Text <> "" Then
                    DTA.FillBy_CustCode_TanggalSampaiDari(DS.DataTable9, Me.txtCusomerCode.Text, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
                Else
                    DTA.FillBy_TanggalSampaiDari(DS.DataTable9, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
                End If
            End If

            'dapatkan primary product min
            DS.DataTable9.DefaultView.Sort = Nothing
            DS.DataTable9.DefaultView.Sort = "[Primary Product]"
            If DS.DataTable9.DefaultView.Count > 0 Then
                row = DS.DataTable9.DefaultView(0).Row
                PrimaryProductMin = row.Primary_Product
            End If

            'dapatkan primary product max
            DS.DataTable9.DefaultView.Sort = Nothing
            DS.DataTable9.DefaultView.Sort = "[Primary Product] DESC"
            If DS.DataTable9.DefaultView.Count > 0 Then
                row = DS.DataTable9.DefaultView(0).Row
                PrimaryProductMax = row.Primary_Product
            End If

            'isikan data POB
            DtaPOB.FillBy_PrimaryProductMinMax(TblPOB, PrimaryProductMin, PrimaryProductMax)

            Dim cmd As String = ""
            For Each row In DS.DataTable9.Rows
                'isikan data IPO
                cmd = "[No SO] = '" + row.SO_Number + "' and [Product Code]='" + row.Product_Code + "'"
                TBLIpo.DefaultView.RowFilter = cmd

                For Each row_View In TBLIpo.DefaultView
                    rowIpo = row_View.Row
                    If row.IsIPO_NumberNull = True Then
                        row.IPO_Number = rowIpo.IPO_Number
                    Else
                        If row.IPO_Number = "" Then
                            row.IPO_Number = rowIpo.IPO_Number
                        Else
                            row.IPO_Number = row.IPO_Number + ", " + rowIpo.IPO_Number
                        End If
                    End If
                Next

                'tambahan aul
                TblPOB.DefaultView.RowFilter = Nothing
                TblPOB.DefaultView.RowFilter = "[Primary Product] = " & row.Primary_Product
                If TblPOB.DefaultView.Count > 0 Then
                    rowPOB = TblPOB.DefaultView(0).Row
                    row.Total_POB = rowPOB.POB
                Else
                    row.Total_POB = 0
                End If
                '-----

                'isikan data available to deliver
                cmd = "[Product Code]='" + row.Product_Code + "'"
                TBL_Stock.DefaultView.RowFilter = cmd
                If TBL_Stock.DefaultView.Count > 0 Then
                    row_Stock = TBL_Stock.DefaultView(0).Row
                    row.Qty_Available = row_Stock.TotalOnHand - row_Stock.QtyBookingInv - row_Stock.QtyPendingDO
                Else
                    row.Qty_Available = 0
                End If
                'isikan data booking
                cmd = "[Primary SO] =" + row.PrimarySO.ToString + " AND [Primary Product] = " + row.Primary_Product.ToString
                TBL_BookingInv.DefaultView.RowFilter = cmd
                If TBL_BookingInv.DefaultView.Count > 0 Then
                    row_BookingInv = TBL_BookingInv.DefaultView(0).Row
                    JmlBookingInv = row_BookingInv.Qty_Booking_Inv
                Else
                    JmlBookingInv = 0
                End If
                cmd = "PrimarySO =" + row.PrimarySO.ToString + " AND [Primary Product] = " + row.Primary_Product.ToString
                TBLBookingGIT.DefaultView.RowFilter = cmd
                If TBLBookingGIT.DefaultView.Count > 0 Then
                    row_BookingGIT = TBLBookingGIT.DefaultView(0).Row
                    JmlBookingGIT = row_BookingGIT.Qty_Booking_GIT
                Else
                    JmlBookingGIT = 0
                End If
                row.Qty_Booking = JmlBookingGIT + JmlBookingInv
            Next

            'pindahkan ke table cetakan
            DS.laporan_sales_orderbalance_view9.Rows.Clear()


            report.SetDataSource(DS)
            If Me.txtCusomerCode.Text <> "" Then
                report.SetParameterValue(report.Parameter_CustomerCode.ParameterFieldName, Me.txtCusomerCode.Text)
                report.SetParameterValue(report.Parameter_CustomerName.ParameterFieldName, Me.txtCustomerName.Text)
            Else
                report.SetParameterValue(report.Parameter_CustomerCode.ParameterFieldName, "All")
                report.SetParameterValue(report.Parameter_CustomerName.ParameterFieldName, "All")
            End If
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, 1)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, 3)
            report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, 1)
            report.SetParameterValue(report.Parameter_TglDari.ParameterFieldName, Me.dtpFrom.Value)
            report.SetParameterValue(report.Parameter_TglSampai.ParameterFieldName, Me.dtpTo.Value)
            Me.crv1.ReportSource = report
            Me.crv1.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            TBLIpo.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
            DTA.Dispose()
            DTA_BookingGIT.Dispose()
            DTA_BookingInv.Dispose()
        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFilter.CheckedChanged
        Try
            If Me.chkFilter.Checked = True Then
                Me.dtpFrom.Enabled = True
                Me.dtpTo.Enabled = True
            Else
                Me.dtpFrom.Enabled = False
                Me.dtpTo.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanSalesOrderBalance_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class