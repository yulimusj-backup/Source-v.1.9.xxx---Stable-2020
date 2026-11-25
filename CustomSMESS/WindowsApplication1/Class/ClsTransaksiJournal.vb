Option Explicit On
Option Strict Off

Public Class ClsTransaksiJournal
    Public Function ApakahBolehEditCustom(ByVal vPrimary_Jurnal_SMESS As Double) As Boolean
        Return True
    End Function

    Public Sub CetakJurnal_Cash_n_Bank1(ByVal vPrimary As Double, _
ByVal Crv As CrystalDecisions.Windows.Forms.CrystalReportViewer)
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function CetakJournal_DetailInvoice_Purchase(ByVal vPrimary As Double, _
    ByVal Crv As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vDGV As DataGridView, _
    ByVal vIndexPaymentOriginal As Integer, ByVal vIndexCurrencyBalance As Integer, _
    ByVal vIndexSupplierInvoiceNumber As Integer, ByVal vIndexAOPNumber As Integer, _
    ByVal vIndexAOPDate As Integer) As Boolean
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function CetakJournal_DetailInvoice_Sales( _
    ByVal vPrimary As Double, _
    ByVal Crv As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal vDGV As DataGridView, ByVal vIndexReceivementOriginal As Integer, ByVal vIndexCurrencyBalance As Integer, _
    ByVal vIndexInvoiceNo As Integer, ByVal vIndexInvoiceDate As Integer) As Boolean
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function CetakJurnal(ByVal vPrimary As Double, _
    ByVal Crv As CrystalDecisions.Windows.Forms.CrystalReportViewer) As Boolean
        Dim DTA As New DseTransaksiJurnalTableAdapters.tblCetakJurnalTableAdapter
        Dim DS As New DseTransaksiJurnal
        Dim TBL As DseTransaksiJurnal.tblCetakJurnalDataTable = DS.tblCetakJurnal
        Dim FormatDecimal As Integer = 0
        Dim Negatif As Integer = 3
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim Report As New CrpCetakJurnal

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimary)

            Report.SetDataSource(DS)
            Report.SetParameterValue(Report.Parameter_FormatDecimal.ParameterFieldName, ObjCur.GetDecimal_DefaultCurrency)
            Report.SetParameterValue(Report.Parameter_Negatif.ParameterFieldName, Negatif)

            Crv.ReportSource = Report
            Crv.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCur = Nothing
            ObjSetting = Nothing
        End Try
    End Function

    Public Function ApakahBolehHapusCustom(ByVal vPrimary_Jurnal_SMESS As Double) As Boolean
        Return True
    End Function
End Class
