Option Explicit On
Option Strict Off

Public Class FrmLaporanFormatStockInOut_Bundling
    Public primary As Double = -1
    Public TipeLaporan As EnumTipeLaporan = EnumTipeLaporan.StockInPurchase
    Public FormatLaporan As EnumFormatLaporan = EnumFormatLaporan.A4

    Dim ObjSetting As New Core_Module.ClsTransaksiSetting

    Public Enum EnumTipeLaporan As Integer
        StockInPurchase
        StockOutPurchaseReturn
        StockInSalesReturn
        StockOutSales
        StockInAdjust
        StockOutAdjust
        StockBundling
        StockTransfer
    End Enum

    Public Enum EnumFormatLaporan As Integer
        A4
        A5_Landscape
    End Enum

    Private Sub FrmLaporanFormatStockInOut_Bundling_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim FrmPrint As New FrmCetakTransaksiStockTransfer


        Try

            ''jika custom, langsung panggil fungsi cetak custom
            If Me.TipeLaporan = EnumTipeLaporan.StockOutPurchaseReturn Then
                If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Out_Purchase_Return, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
                    ObjSetting = Nothing
                    Dim ObjCustom As New CustomSMESS.ClsTransaksiStockInOut
                    ObjCustom.CetakStockOutPurchaseReturn(FrmPrint.crv1, Me.primary)
                    FrmPrint.TipeLaporan = EnumTipeLaporan.StockOutPurchaseReturn
                    FrmPrint.CustomReport = True
                    FrmPrint.ShowDialog()
                    Me.Close()
                End If
            End If
            If Me.TipeLaporan = EnumTipeLaporan.StockInPurchase Then
                If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_In_Purchase, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
                    ObjSetting = Nothing
                    Dim ObjCustom As New CustomSMESS.ClsTransaksiStockInOut
                    ObjCustom.CetakStockInPurchase(FrmPrint.crv1, Me.primary)
                    FrmPrint.TipeLaporan = EnumTipeLaporan.StockInPurchase
                    FrmPrint.CustomReport = True
                    FrmPrint.ShowDialog()
                    Me.Close()
                End If
            End If
            If Me.TipeLaporan = EnumTipeLaporan.StockTransfer Then
                If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Transfer, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
                    ObjSetting = Nothing
                    Dim ObjCustom As New CustomSMESS.clsTransaksiStockTransfer
                    ObjCustom.CetakStockTransfer(FrmPrint.crv1, Me.primary)

                    FrmPrint.TipeLaporan = EnumTipeLaporan.StockTransfer
                    FrmPrint.CustomReport = True
                    FrmPrint.ShowDialog()
                    Me.Close()
                End If
            End If
            If Me.TipeLaporan = EnumTipeLaporan.StockInAdjust Then
                If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_In_Adjust, _
                                Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
                    ObjSetting = Nothing
                    Dim ObjCustom As New CustomSMESS.clsTransaksiStockInOut
                    ObjCustom.CetakStockInAdjust(FrmPrint.crv1, Me.primary)
                    FrmPrint.TipeLaporan = EnumTipeLaporan.StockInAdjust
                    FrmPrint.CustomReport = True
                    FrmPrint.ShowDialog()
                    Me.Close()
                End If
            End If
            If Me.TipeLaporan = EnumTipeLaporan.StockOutAdjust Then
                If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Out_Adjust, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
                    ObjSetting = Nothing
                    Dim ObjCustom As New CustomSMESS.clsTransaksiStockInOut
                    ObjCustom.CetakStockOutAdjust(FrmPrint.crv1, Me.primary)
                    FrmPrint.TipeLaporan = EnumTipeLaporan.StockOutAdjust
                    FrmPrint.CustomReport = True
                    FrmPrint.ShowDialog()
                    Me.Close()
                End If
            End If
            If Me.TipeLaporan = EnumTipeLaporan.StockInSalesReturn Then
                If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_In_SalesReturn, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
                    ObjSetting = Nothing
                    Dim ObjCustom As New CustomSMESS.clsTransaksiStockInOut
                    ObjCustom.CetakStockInSalesReturn(FrmPrint.crv1, Me.primary)
                    FrmPrint.TipeLaporan = EnumTipeLaporan.StockInSalesReturn
                    FrmPrint.CustomReport = True
                    FrmPrint.ShowDialog()
                    Me.Close()
                End If
            End If
            If Me.TipeLaporan = EnumTipeLaporan.StockOutSales Then
                If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Out_Sales, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
                    ObjSetting = Nothing
                    Dim ObjCustom As New CustomSMESS.clsTransaksiStockInOut
                    If ObjCustom.CetakStockOutSales(FrmPrint.crv1, Me.primary) = True Then
                        FrmPrint.TipeLaporan = EnumTipeLaporan.StockOutSales
                        FrmPrint.CustomReport = True
                        FrmPrint.ShowDialog()
                    End If
                    Me.Close()
                End If
            End If
            Me.CboFormat.SelectedIndex = 0

            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmPrint = Nothing
        End Try
    End Sub

    Private Sub BtnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnShow.Click
        Dim frm As New FrmCetakTransaksiStockTransfer

        Try
            frm.primary = primary
            If Me.CboFormat.SelectedIndex = 0 Then
                frm.FormatLaporan = FrmCetakTransaksiStockTransfer.EnumFormatLaporan.A4
            Else
                frm.FormatLaporan = FrmCetakTransaksiStockTransfer.EnumFormatLaporan.A5_Landscape
            End If
            frm.TipeLaporan = Me.TipeLaporan
            frm.ShowDialog()
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            frm = Nothing
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class