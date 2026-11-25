Option Explicit On
Option Strict Off

Public Class FrmCetakMasterProduct_PriceList
    Public mShowCustomerGroup As Boolean = False
    Public PrimaryCurrency As Double
    Public PrimaryCustomer As Double
    Public ProductCode As String
    Public TypeRbo As EnumGeneral

    Dim report As New CrystalDecisions.CrystalReports.Engine.ReportDocument
    'Dim reportGeneralAll As New CrpLapMasterProduct_PriceList_General
    Dim reportGeneral As New CrpLapMasterProduct_PriceList_General
    Dim reportGeneralAda As New CrpLapMasterProduct_PriceList_GeneralAdaPrototype
    Dim reportCust As New CrpLapMasterProduct_PriceList_Cust
    Dim reportGroup As New CrpLapMasterProduct_PriceList_Group

    Enum EnumGeneral
        general
        CustomerGroup
        CustomerGroupSatuan
        Customer
        CustomerSatuan
        GeneralAll
    End Enum

    '---------------------------- Ke TBL General Deafault
    'Public Sub CetakGeneral(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vPrimary As Double, ByVal vProdCode As String)
    '    Dim DS As New DseMasterProduct_EditPriceList

    '    Dim DTAGen As New DseMasterProduct_EditPriceListTableAdapters.tbl_Cetak_GeneralTableAdapter
    '    Dim TBLGen As DseMasterProduct_EditPriceList.tbl_Cetak_GeneralDataTable = DS.tbl_Cetak_General

    '    Try
    '        If Me.TypeRbo = EnumGeneral.general Then
    '            DTAGen.FillBy_ProdCodeAndCurrency(TBLGen, vProdCode, vPrimary)
    '            Me.reportGeneral.SetDataSource(DS)
    '            CRV.ReportSource = reportGeneral
    '            CRV.Show()
    '        End If

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
    '    End Try
    'End Sub
    '---------------------------- Ke TBL General2
    Public Sub CetakGeneral(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vPrimary As Double, ByVal vProdCode As String)
        Dim DS As New DseMasterProduct_EditPriceList

        Dim DTAGen As New DseMasterProduct_EditPriceListTableAdapters.TblCetak_General2TableAdapter
        Dim TBLGen As DseMasterProduct_EditPriceList.TblCetak_General2DataTable = DS.TblCetak_General2

        Try
            If Me.TypeRbo = EnumGeneral.general Then
                DTAGen.FillByProdCodeAndCurrency(TBLGen, vPrimary, vProdCode)
                Me.reportGeneralAda.SetDataSource(DS)
                CRV.ReportSource = reportGeneralAda
                CRV.Show()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub
    '-----------------------------------------------------------------------
    Public Sub CetakGeneralAll(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vPrimary As Double)
        Dim DS As New DseMasterProduct_EditPriceList

        Dim DTAGen As New DseMasterProduct_EditPriceListTableAdapters.tbl_Cetak_GeneralTableAdapter
        Dim TBLGen As DseMasterProduct_EditPriceList.tbl_Cetak_GeneralDataTable = DS.tbl_Cetak_General

        Try
            If Me.TypeRbo = EnumGeneral.GeneralAll Then
                DTAGen.FillBy_PrimCurrency_Cetak(TBLGen, vPrimary)
                Me.reportGeneral.SetDataSource(DS)
                CRV.ReportSource = reportGeneral
                CRV.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub


    'Public Sub CetakCustomerGroup(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vPrimary As Double)
    '    Dim DS As New DseMasterProduct_EditPriceList

    '    Dim DTAGroup As New DseMasterProduct_EditPriceListTableAdapters.tblCetak_GroupTableAdapter
    '    Dim TBLGroup As DseMasterProduct_EditPriceList.TblCetakGroupDataTable = DS.TblCetakGroup

    '    Try
    '        If Me.TypeRbo = EnumGeneral.CustomerGroup Then
    '            DTAGroup.FillBy_PrimaryCurrency_Cetak(TBLGroup, vPrimary)
    '            Me.reportGroup.SetDataSource(DS)
    '            CRV.ReportSource = reportGroup
    '            CRV.Show()
    '        End If



    '    Catch ex As Exception

    '    End Try
    'End Sub


    Public Sub CetakCustomer(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vPrimary As Double)
        Dim DS As New DseMasterProduct_EditPriceList


        Dim DTACust As New DseMasterProduct_EditPriceListTableAdapters.TblCetak_CustomerTableAdapter
        Dim TBLCust As DseMasterProduct_EditPriceList.TblCetak_CustomerDataTable = DS.TblCetak_Customer

        Try
            If Me.TypeRbo = EnumGeneral.Customer Then
                DTACust.FillBy_primaryCurrency(TBLCust, vPrimary)
                Me.reportCust.SetDataSource(DS)
                CRV.ReportSource = reportCust
                CRV.Show()
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Public Sub CetakCustomerSatuan(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vPrimary As Double, ByVal vPrimaryCustomer As Double)
        Dim DS As New DseMasterProduct_EditPriceList


        Dim DTACust As New DseMasterProduct_EditPriceListTableAdapters.TblCetak_CustomerTableAdapter
        Dim TBLCust As DseMasterProduct_EditPriceList.TblCetak_CustomerDataTable = DS.TblCetak_Customer

        Try
            If Me.TypeRbo = EnumGeneral.CustomerSatuan Then
                DTACust.FillBy_PrimaryCurrencyAndCustomer(TBLCust, vPrimary, vPrimaryCustomer)
                Me.reportCust.SetDataSource(DS)
                CRV.ReportSource = reportCust
                CRV.Show()
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub FrmCetakMasterProduct_PriceList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If Me.mShowCustomerGroup = False Then
                If Me.TypeRbo = EnumGeneral.GeneralAll Then
                    Me.CetakGeneralAll(Me.CrvLaporanMasterPriceList, Me.PrimaryCurrency)
                ElseIf Me.TypeRbo = EnumGeneral.general Then
                    Me.CetakGeneral(Me.CrvLaporanMasterPriceList, Me.PrimaryCurrency, Me.ProductCode)
                ElseIf Me.TypeRbo = EnumGeneral.Customer Then
                    Me.CetakCustomer(Me.CrvLaporanMasterPriceList, Me.PrimaryCurrency)
                ElseIf Me.TypeRbo = EnumGeneral.CustomerSatuan Then
                    Me.CetakCustomerSatuan(Me.CrvLaporanMasterPriceList, Me.PrimaryCurrency, Me.PrimaryCustomer)

                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub
End Class