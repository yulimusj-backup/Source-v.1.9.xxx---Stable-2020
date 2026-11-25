Option Explicit On
Option Strict Off

Public Class frmLaporanInventoryCheckTag
    Public primaryCheckTag As String
    Public TagNo As String
    Public TeamNo As String
    Public DateCheckTag As Date
    Public mLocation As String = ""

    Private Sub Cetak()
        Dim AdaBatchNo As Boolean = False
        Dim adaED As Boolean = False
        Dim AdaSecondariQty As Boolean = False
        Dim PrimaryProductDummy As Double = 0
        Dim TotalQty As Double = 0

        Try
            Dim ds As New dseLaporanInventoryCheckTagComplete
            Dim DTA As New dseLaporanInventoryCheckTagCompleteTableAdapters.tblLaporanCheckTagTableAdapter
            DTA.FillBy_PrimaryTag(ds.tblLaporanCheckTag, primaryCheckTag)
            DTA.Dispose()
            Dim row As dseLaporanInventoryCheckTagComplete.tblLaporanCheckTagRow
            For Each row In ds.tblLaporanCheckTag.Rows
                If String.IsNullOrEmpty(row.Batch_No) = False Then
                    AdaBatchNo = True
                    Exit For
                End If
            Next
            For Each row In ds.tblLaporanCheckTag.Rows
                If row.Expiry_Date <> #1/1/1950# Then
                    adaED = True
                    Exit For
                End If
            Next
            For Each row In ds.tblLaporanCheckTag.Rows
                If row.Secondary_Qty > 0 Then
                    AdaSecondariQty = True
                    Exit For
                End If
            Next

            'hitung ulang nilai total qty
            For Each row In ds.tblLaporanCheckTag.Rows
                If row.Primary_Product <> PrimaryProductDummy Then
                    PrimaryProductDummy = row.Primary_Product
                    TotalQty = Me.GetTotal(PrimaryProductDummy, ds.tblLaporanCheckTag)
                End If
                row.TotalQty = TotalQty
            Next
            ds.tblLaporanCheckTag.DefaultView.RowFilter = ""

            If AdaBatchNo = True And adaED = True Then
                Me.CetakComplete(ds)
            ElseIf AdaBatchNo = True And adaED = False Then
                Me.CetakTanpaED(ds)
            ElseIf AdaBatchNo = False And AdaSecondariQty = True Then
                Me.CetakTanpaBatchNo(ds)
            ElseIf AdaBatchNo = False And AdaSecondariQty = False Then
                Me.CetakTanpaEDdanBatchNo(ds)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetTotal(ByVal vPrimaryProduct As Double, ByVal vTBL As dseLaporanInventoryCheckTagComplete.tblLaporanCheckTagDataTable) As Double
        Dim Filter As String = ""
        Dim rowView As Data.DataRowView
        Dim row As dseLaporanInventoryCheckTagComplete.tblLaporanCheckTagRow
        Dim RetVal As Double = 0

        Try
            Filter = "[Primary Product] = " + vPrimaryProduct.ToString
            vTBL.DefaultView.RowFilter = Filter
            For Each rowView In vTBL.DefaultView
                row = rowView.Row
                RetVal += row.Qty
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub CetakTanpaEDdanBatchNo(ByVal ds As dseLaporanInventoryCheckTagComplete)
        Dim report As New crpLaporanInventoryCheckTagTanpaEDdanBatchNo
        Try
            '       DTA.FillBy_PrimaryTag(ds.tblLaporanCheckTag, primaryCheckTag)
            report.SetDataSource(ds)
            report.SetParameterValue(report.Parameter_Date.ParameterFieldName, DateCheckTag)
            report.SetParameterValue(report.Parameter_FormatDays.ParameterFieldName, 1)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, 3)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, 1)
            report.SetParameterValue(report.Parameter_TagNo.ParameterFieldName, TagNo)
            report.SetParameterValue(report.Parameter_TeamNo.ParameterFieldName, TeamNo)
            report.SetParameterValue(report.Parameter_Location.ParameterFieldName, Me.mLocation)
            Me.CrvLaporan.ReportSource = report
            Me.CrvLaporan.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ds.Dispose()
        End Try
    End Sub

    Private Sub CetakTanpaBatchNo(ByVal ds As dseLaporanInventoryCheckTagComplete)
        Dim report As New crpLaporanInventoryCheckTagTanpaBatchNo
        Try
            '       DTA.FillBy_PrimaryTag(ds.tblLaporanCheckTag, primaryCheckTag)
            report.SetDataSource(ds)
            report.SetParameterValue(report.Parameter_Date.ParameterFieldName, DateCheckTag)
            report.SetParameterValue(report.Parameter_FormatDays.ParameterFieldName, 1)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, 3)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, 1)
            report.SetParameterValue(report.Parameter_TagNo.ParameterFieldName, TagNo)
            report.SetParameterValue(report.Parameter_TeamNo.ParameterFieldName, TeamNo)
            report.SetParameterValue(report.Parameter_Location.ParameterFieldName, Me.mLocation)
            Me.CrvLaporan.ReportSource = report
            Me.CrvLaporan.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ds.Dispose()
        End Try
    End Sub

    Private Sub CetakTanpaED(ByVal ds As dseLaporanInventoryCheckTagComplete)
        Dim report As New crpLaporanInventoryCheckTagTanpaED
        Try
            '       DTA.FillBy_PrimaryTag(ds.tblLaporanCheckTag, primaryCheckTag)
            report.SetDataSource(ds)
            report.SetParameterValue(report.Parameter_Date.ParameterFieldName, DateCheckTag)
            report.SetParameterValue(report.Parameter_FormatDays.ParameterFieldName, 1)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, 3)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, 1)
            report.SetParameterValue(report.Parameter_TagNo.ParameterFieldName, TagNo)
            report.SetParameterValue(report.Parameter_TeamNo.ParameterFieldName, TeamNo)
            report.SetParameterValue(report.Parameter_Location.ParameterFieldName, Me.mLocation)
            Me.CrvLaporan.ReportSource = report
            Me.CrvLaporan.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ds.Dispose()
        End Try
    End Sub

    Private Sub CetakComplete(ByVal ds As dseLaporanInventoryCheckTagComplete)
        Dim report As New crpLaporanInventoryCheckTagComplete
        Try
            '       DTA.FillBy_PrimaryTag(ds.tblLaporanCheckTag, primaryCheckTag)
            report.SetDataSource(ds)
            report.SetParameterValue(report.Parameter_Date.ParameterFieldName, DateCheckTag)
            report.SetParameterValue(report.Parameter_FormatDays.ParameterFieldName, 1)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, 3)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, 1)
            report.SetParameterValue(report.Parameter_TagNo.ParameterFieldName, TagNo)
            report.SetParameterValue(report.Parameter_TeamNo.ParameterFieldName, TeamNo)
            report.SetParameterValue(report.Parameter_Location.ParameterFieldName, Me.mLocation)
            Me.CrvLaporan.ReportSource = report
            Me.CrvLaporan.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ds.Dispose()
        End Try
    End Sub

    Private Sub frmLaporanInventoryCheckTag_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Cetak()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class