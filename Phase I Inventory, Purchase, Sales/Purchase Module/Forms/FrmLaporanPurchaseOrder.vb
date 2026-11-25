Option Explicit On
Option Strict Off

Public Class FrmLaporanPurchaseOrder
    Public UserCode As String = ""
    Dim FormatDay As Integer = 1
    Dim FormatMonth As Integer = 2
    Dim FormatYear As Integer = 1
    Dim FormatDecimal As Integer = 2
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
            Me.Cetak1()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Cetak1()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim report As New CrpLaporanPurchaseOrder2
        Dim DS As New DseLaporanPurchaseOrder
        Dim counter As Double = 0
        Dim DTA As New DseLaporanPurchaseOrderTableAdapters.DataTable3TableAdapter
        Dim row As DseLaporanPurchaseOrder.DataTable3Row
        Dim filter As String = ""
        Dim DTAPOC As New DseLaporanPurchaseOrderTableAdapters.tbltrans_pocTableAdapter
        Dim TBLPOC As New DseLaporanPurchaseOrder.tbltrans_pocDataTable
        Dim rowPOC As DseLaporanPurchaseOrder.tbltrans_pocRow
        Dim DTAAOP As New DseLaporanPurchaseOrderTableAdapters.DataTable4TableAdapter
        Dim TBLAOP As New DseLaporanPurchaseOrder.DataTable4DataTable
        Dim rowAOP As DseLaporanPurchaseOrder.DataTable4Row
        Dim DTAPR As New DseLaporanPurchaseOrderTableAdapters.DataTable5TableAdapter
        Dim TBLPR As New DseLaporanPurchaseOrder.DataTable5DataTable
        Dim rowPR As DseLaporanPurchaseOrder.DataTable5Row

        Dim Dummy As String = ""
        Dim RowView As Data.DataRowView

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAPOC.Connection.ConnectionString = Core_Module.GetConstring
            DTAAOP.Connection.ConnectionString = Core_Module.GetConstring
            DTAPR.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(DS.DataTable3, Me.dtpTo.Value, Me.dtpFrom.Value)
            DTAPOC.Fill(TBLPOC)
            DTAAOP.Fill(TBLAOP)
            DTAPR.Fill(TBLPR)

            For Each row In DS.DataTable3.Rows
                filter = "[Primary PO] = " + row.Primary.ToString
                Dummy = ""
                TBLPOC.DefaultView.RowFilter = filter
                If TBLPOC.DefaultView.Count > 0 Then
                    For Each RowView In TBLPOC.DefaultView
                        rowPOC = RowView.Row
                        If Dummy = "" Then
                            Dummy = rowPOC.POC
                        Else
                            Dummy = Dummy + ", " + rowPOC.POC
                        End If
                    Next
                End If

                If Dummy <> "" Then
                    row.POC = Dummy
                Else
                    row.POC = ""
                End If

                Dummy = ""
                TBLAOP.DefaultView.RowFilter = filter
                If TBLAOP.DefaultView.Count > 0 Then
                    For Each RowView In TBLAOP.DefaultView
                        rowAOP = RowView.Row
                        If Dummy = "" Then
                            Dummy = rowAOP.AOPNumber
                        Else
                            Dummy = Dummy + ", " + rowAOP.AOPNumber
                        End If
                    Next
                End If

                If Dummy <> "" Then
                    row.AOP = Dummy
                Else
                    row.AOP = ""
                End If

                Dummy = ""
                TBLPR.DefaultView.RowFilter = filter
                If TBLPR.DefaultView.Count > 0 Then
                    For Each RowView In TBLPR.DefaultView
                        rowPR = RowView.Row
                        If Dummy = "" Then
                            Dummy = rowPR.AOPR
                        Else
                            Dummy = Dummy + ", " + rowPR.AOPR
                        End If
                    Next
                End If

                If Dummy <> "" Then
                    row.AOPR = Dummy
                Else
                    row.AOPR = ""
                End If
            Next


            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
            report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, Me.FormatYear)
            report.SetParameterValue(report.Parameter_TglDari.ParameterFieldName, Me.dtpFrom.Value)
            report.SetParameterValue(report.Parameter_TglSampai.ParameterFieldName, Me.dtpTo.Value)
            Me.crvReport.ReportSource = report
            Me.crvReport.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
            DTAPOC.Dispose()
            TBLPOC.Dispose()
        End Try
    End Sub

    Private Sub Cetak()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim report As New CrpLaporanPurchaseOrder
        Dim DS As New DseLaporanPurchaseOrder
        Dim counter As Double = 0

        Try
            Dim DTA As New DseLaporanPurchaseOrderTableAdapters.laporan_po_recapitulation_view2TableAdapter
            DTA.FillBy_tglDariSampai(DS.laporan_po_recapitulation_view2, Me.dtpFrom.Value, Me.dtpTo.Value)
            DTA.Dispose()
            Dim row As DseLaporanPurchaseOrder.laporan_po_recapitulation_view2Row
            Dim rowTotal As DseLaporanPurchaseOrder.TblTotALRow
            For Each row In DS.laporan_po_recapitulation_view2.Rows
                If DS.TblTotAL.Rows.Count = 0 Then
                    Dim ActualPurchase As Double = row.AOP_Amount - row.AOPR_Amount
                    DS.TblTotAL.Rows.Add(row.Currency_Code, row.AOP_Amount, row.AOPR_Amount, row.POC_Amount, row.Total_PO, ActualPurchase)
                Else
                    counter = 0
                    While counter < DS.TblTotAL.Rows.Count
                        rowTotal = DS.TblTotAL.Rows(counter)
                        If rowTotal.Currency_Code = row.Currency_Code Then
                            Dim ActualPurchase As Double = row.AOP_Amount - row.AOPR_Amount
                            rowTotal.Total_AOP += row.AOPR_Amount
                            rowTotal.Total_PR += row.AOPR_Amount
                            rowTotal.TotalPOC += row.POC_Amount
                            rowTotal.Total_Amount += row.Total_PO
                            rowTotal.Actual_Purchase += ActualPurchase
                        Else
                            Dim ActualPurchase As Double = row.AOP_Amount - row.AOPR_Amount
                            DS.TblTotAL.Rows.Add(row.Currency_Code, row.AOP_Amount, row.AOPR_Amount, row.POC_Amount, row.Total_PO, ActualPurchase)
                        End If
                        counter += 1
                    End While
                End If
            Next

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
            report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, Me.FormatYear)
            report.SetParameterValue(report.Parameter_TglDari.ParameterFieldName, Me.dtpFrom.Value)
            report.SetParameterValue(report.Parameter_TglSampai.ParameterFieldName, Me.dtpTo.Value)
            Me.crvReport.ReportSource = report
            Me.crvReport.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
        End Try
    End Sub
End Class