Option Explicit On
Option Strict Off

Public Class FrmLaporanPurchaseReturnBySupplier
    Public UserCode As String = ""
    Dim FormatDay As Integer = 1
    Dim FormatMonth As Integer = 2
    Dim FormatYears As Integer = 1
    Dim FormatDecimal As Integer = 2
    Dim TBL As New Core_Module.dseMasterSupplier.tblmaster_supplier2DataTable

    Private Sub FrmLaporanInternalPurchaseOrderBySupplier_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Obj As New Core_Module.clsMasterSupplier

        Try
            Obj.IsiCheckListBox(Me.clbSupplier, Me.TBL)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj = Nothing
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            For i As Integer = 0 To Me.clbSupplier.Items.Count - 1
                Me.clbSupplier.SetItemCheckState(i, CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Try
            For i As Integer = 0 To Me.clbSupplier.Items.Count - 1
                Me.clbSupplier.SetItemCheckState(i, CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Cetak()
        Dim DS As New DseLaporanPurchaseReturnBySupplier
        Dim report As New CrpLaporanPurchaseReturnBySupplier
        Try
            Dim DTA As New DseLaporanPurchaseReturnBySupplierTableAdapters.DataTable1TableAdapter
            DTA.FillBy_TglDariSampai(DS.DataTable1, Me.dtpFrom.Value, Me.dtpTo.Value)
            DTA.Dispose()
            Dim DV As DataRowView
            Dim KodeSupplier As String = ""
            Dim Ketemu As Boolean = True
            For i As Integer = 0 To DS.DataTable1.Rows.Count - 1
                Ketemu = False
                If i = DS.DataTable1.Rows.Count Then
                    Exit For
                End If
                For Each DV In Me.clbSupplier.CheckedItems
                    If DV.Item(1).ToString = DS.DataTable1.Rows(i).Item("Supplier Code") Then
                        Ketemu = True
                        Exit For
                    End If
                Next
                If Ketemu = False Then
                    DS.DataTable1.Rows.RemoveAt(i)
                    i -= 1
                End If
            Next

            KodeSupplier = ""
            If Me.clbSupplier.CheckedItems.Count = Me.clbSupplier.Items.Count Then
                KodeSupplier = "All"
            Else
                For Each DV In Me.clbSupplier.CheckedItems
                    If KodeSupplier = "" Then
                        KodeSupplier = DV.Item(1)
                    Else
                        KodeSupplier += ", " + DV.Item(1)
                    End If
                Next
            End If

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
            report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
            report.SetParameterValue(report.Parameter_SupplierCode.ParameterFieldName, KodeSupplier)
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

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Try
            Me.Cetak()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class