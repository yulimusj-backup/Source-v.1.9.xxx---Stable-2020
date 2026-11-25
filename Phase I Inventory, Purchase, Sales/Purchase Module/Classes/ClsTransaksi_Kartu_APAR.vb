Option Explicit On
Option Strict Off

Public Class ClsTransaksi_Kartu_APAR
    Public Enum EnumJenisTransaksiAsal As Integer
        Invoice_Otomatis
        Invoice_Manual
        Invoice_Manual_Transfer
        AOP_Otomatis
        AOP_Manual
        AOP_Manual_Transfer
        SR_Otomatis
        SR_Manual
        SR_Manual_Transfer
        PR_Otomatis
        PR_Manual
        PR_Manual_Transfer
        Sales_Receivement
        Sales_Return_Payment
        Purchase_Payment
        Purchase_Return_Receivement
    End Enum

    Public Enum EnumJenisKartu As Integer
        AR
        AP
    End Enum

    Public Function Simpan(ByVal vPrimaryTransaksiAsal As Double, ByVal vJenisTransaksi As EnumJenisTransaksiAsal, ByVal vDocumentNo As String, _
ByVal vPrimaryPihakAcuan As Double, ByVal vDebet As Double, ByVal vKredit As Double, ByVal vDescription As String, _
ByVal vInOut As EnumJenisKartu, ByVal vDate As Date, ByVal DTA As DseTrans_Kartu_APARTableAdapters.tbltransaksi_kartu_ar_apTableAdapter, _
ByVal vPrimaryTransaksiAcuan As Double, ByVal vTanggalTransaksiAcuan As Date, ByVal vPaymentTermTransaksiAcuan As Double, _
ByVal vJenisTransaksiAcuan As EnumJenisTransaksiAsal, ByVal PrimaryCurrency As Double, ByVal vDebetOriginal As Decimal, ByVal vCreditOriginal As Decimal, _
  ByVal vPrimaryAkun As Double, ByVal vExchangeRate As Decimal) As Boolean

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Insert(vPrimaryTransaksiAsal, vJenisTransaksi, vPrimaryPihakAcuan, _
            vDocumentNo, vDebet, vKredit, vDescription, vInOut, vDate, vPrimaryTransaksiAcuan, vTanggalTransaksiAcuan, _
            vPaymentTermTransaksiAcuan, vJenisTransaksiAcuan, PrimaryCurrency, vExchangeRate, vDebetOriginal, vCreditOriginal, vPrimaryAkun) > 0 Then
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Simpan(ByVal vPrimaryTransaksiAsal As Double, ByVal vJenisTransaksi As EnumJenisTransaksiAsal, ByVal vDocumentNo As String, _
    ByVal vPrimaryPihakAcuan As Double, ByVal vDebet As Double, ByVal vKredit As Double, ByVal vDescription As String, _
    ByVal vInOut As EnumJenisKartu, ByVal vDate As Date, _
    ByVal vPrimaryTransaksiAcuan As Double, ByVal vTanggalTransaksiAcuan As Date, ByVal vPaymentTermTransaksiAcuan As Double, _
    ByVal vJenisTransaksiAcuan As EnumJenisTransaksiAsal, ByVal PrimaryCurrency As Double, ByVal vDebetOriginal As Decimal, ByVal vCreditOriginal As Decimal, _
    ByVal vPrimaryAkun As Double, ByVal vExchangeRate As Decimal) As Boolean
        Dim DTA As New DseTrans_Kartu_APARTableAdapters.tbltransaksi_kartu_ar_apTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Insert(vPrimaryTransaksiAsal, vJenisTransaksi, vPrimaryPihakAcuan, _
            vDocumentNo, vDebet, vKredit, vDescription, vInOut, vDate, vPrimaryTransaksiAcuan, vTanggalTransaksiAcuan, _
            vPaymentTermTransaksiAcuan, vJenisTransaksiAcuan, PrimaryCurrency, vExchangeRate, vDebetOriginal, vCreditOriginal, vPrimaryAkun) > 0 Then
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksiAsal As Double, ByVal vJenisTransaksiAsal As EnumJenisTransaksiAsal) As Boolean
        Dim DTA As New DseTrans_Kartu_APARTableAdapters.tbltransaksi_kartu_ar_apTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Delete(vPrimaryTransaksiAsal, vJenisTransaksiAsal) > 0 Then
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
