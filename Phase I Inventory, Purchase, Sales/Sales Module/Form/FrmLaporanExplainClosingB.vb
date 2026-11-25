Option Explicit On
Option Strict Off

Public Class FrmLaporanExplainClosingB
    Dim TBLBrand As New Core_Module.dseMasterProductGroup.tblmaster_product_group2DataTable
    Dim TanggalClosing As Date = #1/1/1950#
    Dim objCD As New Core_Module.ClsCekData

    Dim DTAClosing As New DseClosing_Explain_OutTableAdapters.DataTable2TableAdapter
    Dim TBLClosing As New DseClosing_Explain_Out.DataTable2DataTable

    Private Sub FrmLaporanExplainClosing_1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim OBjBrand As New Core_Module.ClsMasterProductGroup
        Try
            OBjBrand.IsiListBoxBrand(Me.clbBrand, Me.TBLBrand)
            Me.btnShowReport.PerformClick()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            OBjBrand = Nothing
        End Try
    End Sub


    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            For i As Integer = 0 To Me.clbBrand.Items.Count - 1
                Me.clbBrand.SetItemCheckState(i, CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Try
            For i As Integer = 0 To Me.clbBrand.Items.Count - 1
                Me.clbBrand.SetItemCheckState(i, CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Function ApakahBolehCetak() As Boolean
        Try
            If Me.clbBrand.CheckedItems.Count <= 0 Then
                MsgBox("Please select at least 1 Brand.", MsgBoxStyle.Information, "Information")
                Me.clbBrand.Focus()
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Function GetCostOfSales(ByVal vPrimaryProduct As Double, ByVal vClosingPeriod As Date) As Double
        Dim filter As String = ""
        Dim row As DseClosing_Explain_Out.DataTable2Row

        Try
            'sesuaikan tanggal closing, isikan data closing sesuai dengan tanggal closing
            If vClosingPeriod <> Me.TanggalClosing Then
                Me.DTAClosing.Fill(Me.TBLClosing, vClosingPeriod)
                Me.TanggalClosing = vClosingPeriod
            End If

            filter = "[Primary Product] =" + vPrimaryProduct.ToString
            Me.TBLClosing.DefaultView.RowFilter = filter
            If Me.TBLClosing.DefaultView.Count <> 0 Then
                row = Me.TBLClosing.DefaultView(0).Row
                If row.Ending_Qty = 0 Or row.Ending_Amount = 0 Then
                    If row.Out_Qty = 0 Then
                        Return 0
                    Else
                        Return Math.Round(row.Out_Amount / row.Out_Qty, 0, MidpointRounding.AwayFromZero)
                    End If
                Else
                    If row.Ending_Qty = 0 Then
                        Return 0
                    Else
                        Return Math.Round(row.Ending_Amount / row.Ending_Qty, 0, MidpointRounding.AwayFromZero)
                    End If
                End If
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnShowReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShowReport.Click
        Dim DS As New DseClosing_Explain_Out
        Dim DTA1 As New DseClosing_Explain_OutTableAdapters.DataTable7TableAdapter
        Dim report As New CrpLaporanClosingB
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim row As DseClosing_Explain_Out.DataTable4Row
        Dim row1 As DseClosing_Explain_Out.DataTable7Row
        Dim TglDari As Date = objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value)
        Dim TglSampai As Date = objCD.GetLasDateOfTheMonth(Me.dtpFrom.Value)
        Dim DV As Data.DataRowView
        Dim primaryBrand As Double = 0

        Try

            'stock out adjust
            DTA1.Fill(DS.DataTable7, _
            TglSampai, TglDari)

            For Each row1 In DS.DataTable7.Rows
                row = DS.DataTable4.NewDataTable4Row
                With row1
                    row.Brand_Code = .Brand
                    row.Cost = .Cost
                    row.Invoice_Number = .Invoice_Number
                    row.Primary_Brand = .Primary_Brand
                    row.Primary_Product = .Primary_Product
                    row.Product_Code = .Product_Code
                    row.Qty = .Total_Qty
                    row.Stock_Out_Number = .Stock_Out_Number
                    row.Tanggal_Invoice = .Tanggal_Invoice
                    row.Tanggal_Stock_Out = .Tanggal_Stock_Out
                    row.Unit_Code = .Unit_Code
                End With
                DS.DataTable4.Rows.Add(row)
            Next

            'hitung cost inventory
            'For Each row In DS.DataTable4.Rows
            'row.Cost = row.Qty * Me.GetCostOfSales(row.Primary_Product, _
            'Me.objCD.GetFirstDateofTheMonth(row.Tanggal_Stock_Out))
            'Next

            Dim Brand As String = ""
            If Me.clbBrand.CheckedItems.Count = Me.clbBrand.Items.Count Then
                Brand = "All"
            Else
                For Each DV In Me.clbBrand.CheckedItems
                    If Brand = "" Then
                        Brand = DV.Item(1)
                    Else
                        Brand += ", " + DV.Item(1)
                    End If
                Next
            End If

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_Brand.ParameterFieldName, Brand)
            report.SetParameterValue(report.Parameter_Periode.ParameterFieldName, Format(Me.dtpFrom.Value, "MMMMM yyyy"))

            Me.CrystalReportViewer1.ReportSource = report
            Me.CrystalReportViewer1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA1.Dispose()
        End Try
    End Sub
End Class