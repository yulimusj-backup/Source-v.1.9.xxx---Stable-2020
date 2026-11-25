Option Explicit On
Option Strict Off

Public Class ClsMultiAccount
    Public Sub IsiTBLAPARAcount(ByVal vPerkiraan1 As Core_Module.ClsMasterAkun1.EnumTipeAkun, ByVal vPerkiraan2 As Core_Module.ClsMasterAkun1.EnumTipeAkun2, _
     ByVal TBL As DseMultiAccount.tblisi_cbo_ar_akunDataTable, ByVal vPrimaryCurrency As Double, Optional ByVal vClearBeforeFill As Boolean = True)
        Dim DTA As New DseMultiAccountTableAdapters.tblisi_cbo_ar_akunTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.ClearBeforeFill = vClearBeforeFill
            DTA.Fill_ByTipePerkiraan(TBL, vPerkiraan1, vPerkiraan2, vPrimaryCurrency)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Sub IsiTBLAOPAcount(ByVal vPrimary As Double, _
    ByVal TBL As DseMultiAccount.tblmaster_akunDataTable)
        Dim DTA As New DseMultiAccountTableAdapters.tblmaster_akunTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetPrimaryAkunAPBase(ByVal vPrimaryCurrency As Double) As Double
        '- isicbo customer
        Dim tbl As New DseMultiAccount.tblisi_cbo_ar_akunDataTable
        Dim row As DseMultiAccount.tblisi_cbo_ar_akunRow

        Dim tp1 As New Core_Module.ClsMasterAkun1.EnumTipeAkun
        Dim tp2 As New Core_Module.ClsMasterAkun1.EnumTipeAkun2
        tp1 = Core_Module.ClsMasterAkun1.EnumTipeAkun.Liabilities
        tp2 = Core_Module.ClsMasterAkun1.EnumTipeAkun2.Account_Payable
        Dim PrimaryAP As Double = -1

        Try
            Me.IsiTBLAPARAcount(tp1, tp2, tbl, vPrimaryCurrency)
            If tbl.Rows.Count > 0 Then
                row = tbl.Rows(0)
                PrimaryAP = row.Primary
            End If
            Return PrimaryAP

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function


    Public Function GetPrimaryAkunARBase(ByVal vPrimaryCurrency As Double) As Double
        '- isicbo customer
        Dim tbl As New DseMultiAccount.tblisi_cbo_ar_akunDataTable
        Dim row As DseMultiAccount.tblisi_cbo_ar_akunRow

        Dim tp1 As New Core_Module.ClsMasterAkun1.EnumTipeAkun
        Dim tp2 As New Core_Module.ClsMasterAkun1.EnumTipeAkun2
        tp1 = Core_Module.ClsMasterAkun1.EnumTipeAkun.Asset
        tp2 = Core_Module.ClsMasterAkun1.EnumTipeAkun2.Account_Receivable
        Dim PrimaryAR As Double = -1

        Try
            Me.IsiTBLAPARAcount(tp1, tp2, tbl, vPrimaryCurrency)
            If tbl.Rows.Count > 0 Then
                row = tbl.Rows(0)
                PrimaryAR = row.Primary
            End If
            Return PrimaryAR

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub IsiTBLCurrTax(ByVal vPrimary As Double, _
    ByVal TBL As DseMultiAccount.tblCurrencyTaxDataTable)
        Dim DTA As New DseMultiAccountTableAdapters.tblCurrencyTaxTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub
End Class
