Option Explicit On
Option Strict Off

Public Class ClsTransaksiSetting

    Public Enum EnumTipeTransaksi As Long
        SalesJurnal
        PurchaseJurnal
        Invoice
        AOP
        DeliveryOrder
        PO
        PurchaseReturn
    End Enum

    Public Enum EnumNamaParameter As Long
        Akun_AR
        Akun_AP
        UseTax
        TaxPercentage
        PrintUseCustomDO_Bool
        CustomDO_Str
        PrintUseCustomInvoice_Bool
        CustomInvoiceAddress_Str
        PrintUseCustomTaxInvoice_Bool
        CustomTaxInvoice_Str
        Use_Custom_Print_Logic
        SignedByName
        SignedByJabatan
    End Enum

    Public Function GetSettingValue(ByVal vTipeTransaksi As EnumTipeTransaksi, ByVal vNamaParameter As EnumNamaParameter) As String
        Dim DTA As New DseTransaksiSettingTableAdapters.tbltrans_settingTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNilaiParameter(vTipeTransaksi, vNamaParameter)) = False Then
                Return DTA.ScalarQuery_GetNilaiParameter(vTipeTransaksi, vNamaParameter)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub SetSettingValue(ByVal vTipeTransaksi As EnumTipeTransaksi, ByVal vNamaParameter As EnumNamaParameter, ByVal vNilaiParameter As String)
        Dim DTA As New DseTransaksiSettingTableAdapters.tbltrans_settingTableAdapter
        Dim Insert As Boolean = False

        Try
            If DTA.ScalarQuery(vTipeTransaksi, vNamaParameter).HasValue = True Then
                If DTA.ScalarQuery(vTipeTransaksi, vNamaParameter).Value = 0 Then
                    Insert = True
                Else
                    Insert = False
                End If
            Else
                Insert = True
            End If

            If Insert = True Then
                DTA.Insert(vTipeTransaksi, vNamaParameter, vNilaiParameter)
            Else
                DTA.Update(vNilaiParameter, vTipeTransaksi, vNamaParameter)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub
End Class
