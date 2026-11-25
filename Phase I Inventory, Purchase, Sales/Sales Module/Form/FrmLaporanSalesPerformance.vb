Option Explicit On
Option Strict Off

Public Class FrmLaporanSalesPerformance
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjSalesP As New Core_Module.ClsMasterSalesPerson
    Dim ObjSalesPGroup As New Core_Module.ClsMasterSalesPersonGroup
    Dim ObjClosing As New ClsTransaksiMonthlyClosing
    Dim vTBLSP As New Core_Module.DseMasterSalesPerson.tblmaster_salesperson2DataTable
    Dim vRowSP As Core_Module.DseMasterSalesPerson.tblmaster_salesperson2Row
    Dim vTBLSG As New Core_Module.DseMasterSalesPersonGroup.tblmaster_salesperson_group2DataTable
    Dim vRowSG As Core_Module.DseMasterSalesPersonGroup.tblmaster_salesperson_group2Row
    Dim ReferenceData As EnumTipeReference = EnumTipeReference.NoReference
    Dim TipeLaporan As EnumTipeLaporan = EnumTipeLaporan.NoReport
    Dim TipeDetail As EnumTipeDetail = EnumTipeDetail.NoDetail
    Dim TipeGroup As EnumTipeDetail = EnumTipeDetail.NoDetail
    Dim DS As New DseLaporanSalesPerformance
    Dim rpt_Recap1 As New CrpLaporanSalesPerformanceRecap
    Dim rpt_Recap2 As New CrpLaporanSalesPerformanceRecap2
    Dim rpt_Recap3 As New CrpLaporanSalesPerformanceRecap4
    Dim rpt_Recap4 As New CrpLaporanSalesPerformanceRecap3
    Dim rpt_Recap5 As New CrpLaporanSalesPerformanceRecap5
    Dim rpt_Trend As New CrpLaporanSalesPerformance_Trend
    Dim rpt_Trend2 As New CrpLaporanSalesPerformance_Trend2
    Dim rpt_Composition As New CrpLaporanSalesPerformance_Composition
    Dim rpt_Composition2 As New CrpLaporanSalesPerformance_Composition2
    Dim ObjCurrency As New Core_Module.clsMasterCurrency
    Dim ObjLicense As New Core_Module.ClsLicense

    Private Enum EnumTipeReference As Integer
        NoReference
        SO
        Invoice
        GrossMargin
    End Enum

    Private Enum EnumTipeLaporan As Integer
        NoReport
        Sales_Person_Recap
        Sales_Person_Group_Recap
        Sales_Person_Group_Detail
        Sales_Person_Detail
        Sales_Person_Group_Trend
        Sales_Person_Group_Trend_Perct
        Sales_Person_Trend
        Sales_Person_Trend_Perct
        Sales_Person_Group_Composition
        Sales_Person_Composition
    End Enum

    Private Enum EnumTipeDetail As Integer
        NoDetail
        Detail_SalesP
        Detail_SalesPGroup
        Detail_Cust_Segmentation
        Detail_Sales_Area
        Detail_Cust_Type
        Detail_Cust_Group
        Detail_Product_Group
        Detail_Product_Brand
        Detail_Product_Sub_Brand
        Detail_Product_Category
        Detail_Product
    End Enum

    Private Sub FrmLaporanSalesPerformance_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.dtpDari.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpDari.Value)
            Me.dtpSampai.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpSampai.Value)
            Me.dtpDari2.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpDari2.Value)
            Me.dtpSampai2.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpSampai2.Value)

            Me.cboReference.SelectedIndex = 0
            Me.cboFormat.SelectedIndex = 0
            Me.cboFormat1.SelectedIndex = 0
            Me.cboFormat2.SelectedIndex = 0

            Me.ObjSalesP.IsikanChecklistbox(Me.vTBLSP, Me.clbSalesPerson)
            Me.ObjSalesPGroup.IsikanChecklistbox(Me.vTBLSG, Me.clbSalesPersonGroup)
            Me.chkSalesPerson.Checked = True
            Me.chkSalesGroup.Checked = True

            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkSalesGroup_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSalesGroup.CheckedChanged
        Try
            If Me.chkSalesGroup.Checked = True Then
                Me.clbSalesPersonGroup.Enabled = False
                Me.llbSG_CheckAll.Enabled = False
                Me.llbSG_UncheckAll.Enabled = False
                Me.CheckAll_SG()
            Else
                Me.clbSalesPersonGroup.Enabled = True
                Me.llbSG_CheckAll.Enabled = True
                Me.llbSG_UncheckAll.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkSalesPerson_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSalesPerson.CheckedChanged
        Try
            If Me.chkSalesPerson.Checked = True Then
                Me.clbSalesPerson.Enabled = False
                Me.llbSP_CheckAll.Enabled = False
                Me.llbSP_UncheckAll.Enabled = False
                Me.CheckAllSP()
            Else
                Me.clbSalesPerson.Enabled = True
                Me.llbSP_CheckAll.Enabled = True
                Me.llbSP_UncheckAll.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CheckAll_SG()
        Try
            For i As Integer = 0 To Me.clbSalesPersonGroup.Items.Count - 1
                Me.clbSalesPersonGroup.SetItemCheckState(i, CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbSG_CheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbSG_CheckAll.LinkClicked
        Me.CheckAll_SG()
    End Sub

    Private Sub UncheckAllSG()
        Try
            For i As Integer = 0 To Me.clbSalesPersonGroup.Items.Count - 1
                Me.clbSalesPersonGroup.SetItemCheckState(i, CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbSG_UncheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbSG_UncheckAll.LinkClicked
        Me.UncheckAllSG()
    End Sub

    Private Sub CheckAllSP()
        Try
            For i As Integer = 0 To Me.clbSalesPerson.Items.Count - 1
                Me.clbSalesPerson.SetItemCheckState(i, CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbSP_CheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbSP_CheckAll.LinkClicked
        Me.CheckAllSP()
    End Sub

    Private Sub UncheckAllSP()
        Try
            For i As Integer = 0 To Me.clbSalesPerson.Items.Count - 1
                Me.clbSalesPerson.SetItemCheckState(i, CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbSP_UncheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbSP_UncheckAll.LinkClicked
        Me.UncheckAllSP()
    End Sub

    Private Sub cboReference_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboReference.SelectedIndexChanged
        Try
            If Me.cboReference.SelectedIndex = -1 Then
                Me.ReferenceData = EnumTipeReference.NoReference
            ElseIf Me.cboReference.SelectedIndex = 0 Then
                Me.ReferenceData = EnumTipeReference.SO
            ElseIf Me.cboReference.SelectedIndex = 1 Then
                Me.ReferenceData = EnumTipeReference.Invoice
            ElseIf Me.cboReference.SelectedIndex = 2 Then
                Me.ReferenceData = EnumTipeReference.GrossMargin
            Else
                Me.ReferenceData = EnumTipeReference.NoReference
            End If

            If Me.ReferenceData = EnumTipeReference.SO Then
                Me.chkReturn.Checked = False
                Me.chkReturn.Enabled = False
            Else
                Me.chkReturn.Checked = True
                Me.chkReturn.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboFormat_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFormat.SelectedIndexChanged
        Try
            If Me.cboFormat.Text = "Sales Person Recap" Then
                Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Recap
            ElseIf Me.cboFormat.Text = "Sales Person Group Recap" Then
                Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Group_Recap
            ElseIf Me.cboFormat.Text = "Sales Person Group Detail" Then
                Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Group_Detail
            ElseIf Me.cboFormat.Text = "Sales Person Detail" Then
                Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Detail
            ElseIf Me.cboFormat.Text = "Sales Person Group Trend" Then
                Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Group_Trend
            ElseIf Me.cboFormat.Text = "Sales Person Group Trend %" Then
                Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Group_Trend_Perct
            ElseIf Me.cboFormat.Text = "Sales Person Trend" Then
                Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Trend
            ElseIf Me.cboFormat.Text = "Sales Person Trend %" Then
                Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Trend_Perct
            ElseIf Me.cboFormat.Text = "Sales Person Group Composition" Then
                Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Group_Composition
            ElseIf Me.cboFormat.Text = "Sales Person Composition" Then
                Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Composition
            End If

            'yang tidak bisa with budget di atur:
            'If Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Recap Or _
            'Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Group_Recap Or _
            'Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Detail Or _
            'Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Group_Detail Or _
            'Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Composition Or _
            'Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Group_Composition Then
            '    'bisa 2 period
            '    'Me.cboFormat1.SelectedIndex = 0
            '    Me.cboFormat1.Enabled = True
            'Else
            '    'tidak bisa 2 period
            '    Me.cboFormat1.SelectedIndex = 0
            '    Me.cboFormat1.Enabled = False
            'End If
            Me.AturGrouping()
            Me.AturDetail()
            Me.Atur2Period()

            'atur detail
           

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboFormat1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFormat1.SelectedIndexChanged
        Try
            If Me.cboFormat1.Items.Count = 3 Then
                If Me.cboFormat1.SelectedIndex = 0 Or Me.cboFormat1.SelectedIndex = 1 Then
                    Me.dtpDari2.Enabled = False
                    Me.dtpSampai2.Enabled = False
                    Me.lblPeriod2.Enabled = False
                    Me.LblTo2.Enabled = False
                Else
                    Me.dtpDari2.Enabled = True
                    Me.dtpSampai2.Enabled = True
                    Me.lblPeriod2.Enabled = True
                    Me.LblTo2.Enabled = True
                End If
            Else
                If Me.cboFormat1.SelectedIndex = 0 Then
                    Me.dtpDari2.Enabled = False
                    Me.dtpSampai2.Enabled = False
                    Me.lblPeriod2.Enabled = False
                    Me.LblTo2.Enabled = False
                Else
                    Me.dtpDari2.Enabled = True
                    Me.dtpSampai2.Enabled = True
                    Me.lblPeriod2.Enabled = True
                    Me.LblTo2.Enabled = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboFormat2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFormat2.SelectedIndexChanged
        Try
            If Me.cboFormat2.Text = "-" Then
                Me.TipeDetail = EnumTipeDetail.NoDetail
            ElseIf Me.cboFormat2.Text = "Cust.Segmentation" Then
                Me.TipeDetail = EnumTipeDetail.Detail_Cust_Segmentation
            ElseIf Me.cboFormat2.Text = "Sales Area" Then
                Me.TipeDetail = EnumTipeDetail.Detail_Sales_Area
            ElseIf Me.cboFormat2.Text = "Cust.Type" Then
                Me.TipeDetail = EnumTipeDetail.Detail_Cust_Type
            ElseIf Me.cboFormat2.Text = "Cust.Group" Then
                Me.TipeDetail = EnumTipeDetail.Detail_Cust_Group
            ElseIf Me.cboFormat2.Text = "Product Group" Then
                Me.TipeDetail = EnumTipeDetail.Detail_Product_Group
            ElseIf Me.cboFormat2.Text = "Product Brand" Then
                Me.TipeDetail = EnumTipeDetail.Detail_Product_Brand
            ElseIf Me.cboFormat2.Text = "Product Sub Brand" Then
                Me.TipeDetail = EnumTipeDetail.Detail_Product_Sub_Brand
            ElseIf Me.cboFormat2.Text = "Product Category" Then
                Me.TipeDetail = EnumTipeDetail.Detail_Product_Category
            ElseIf Me.cboFormat2.Text = "Product" Then
                Me.TipeDetail = EnumTipeDetail.Detail_Product
            ElseIf Me.cboFormat2.Text = "Sales Person" Then
                Me.TipeDetail = EnumTipeDetail.Detail_SalesP
            ElseIf Me.cboFormat2.Text = "Sales Person Group" Then
                Me.TipeDetail = EnumTipeDetail.Detail_SalesPGroup
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnShowReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShowReport.Click
        Dim StartPeriod As Date = Today
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.dtpDari.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpDari.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpSampai.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpSampai.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpDari2.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpDari2.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpSampai2.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpSampai2.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If Me.chkGoods.Checked = False And Me.chkService.Checked = False Then
                MsgBox("Please select at least one: goods or service. You can choose both.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If Me.dtpDari.Value.Date > Me.dtpSampai.Value.Date Then
                MsgBox("Invalid filter period.", MsgBoxStyle.Information, "Information")
                Me.dtpDari.Focus()
                Exit Sub
            End If
            If Me.dtpDari2.Enabled = True Then
                If Me.dtpDari2.Value.Date > Me.dtpSampai2.Value.Date Then
                    MsgBox("Invalid filter period.", MsgBoxStyle.Information, "Information")
                    Me.dtpDari2.Focus()
                    Exit Sub
                End If
            End If
            If Me.ReferenceData = EnumTipeReference.GrossMargin Then
                'pastikan bulan ybs sudah closing
                StartPeriod = Me.dtpDari.Value.Date
                While StartPeriod <= Me.dtpSampai.Value
                    If Me.ObjClosing.GetPrimaryClosing(StartPeriod) = -1 Then
                        MsgBox("Please do inventory closing first for " + Format(StartPeriod, "MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    StartPeriod = Me.ObjCD.GetFirstDateofTheMonth(StartPeriod + TimeSpan.FromDays(32))
                End While
            End If

            If Me.ReferenceData = EnumTipeReference.SO Then
                If Me.ApakahLengkapKodeSalesPersonnya_SO(#1/1/1950#, #1/1/2050#) = False Then
                    If Me.dtpDari2.Enabled = False Then
                        'cek period 1 saja
                        If Me.ApakahLengkapKodeSalesPersonnya_SO(Me.dtpDari.Value, Me.dtpSampai.Value) = False Then
                            Dim ij As New FrmTransaksiRepairInvoice
                            ij.dtpInvPeriod.Value = Me.dtpDari.Value
                            ij.ShowDialog()
                            ij.Dispose()
                            If Me.ApakahLengkapKodeSalesPersonnya_SO(Me.dtpDari.Value, Me.dtpSampai.Value) = False Then
                                MsgBox("There's still incomplete sales person data in one or several sales order.", MsgBoxStyle.Information, "Information")
                            End If
                            Me.BuatLaporanSO()
                        Else
                            Me.BuatLaporanSO()
                        End If
                    Else
                        'cek period 1 dan 2
                        If Me.ApakahLengkapKodeSalesPersonnya_SO(Me.dtpDari.Value, Me.dtpSampai.Value) = False Or _
                        Me.ApakahLengkapKodeSalesPersonnya_SO(Me.dtpDari2.Value, Me.dtpSampai2.Value) = False Then
                            Dim ij As New FrmTransaksiRepairInvoice
                            ij.dtpInvPeriod.Value = Me.dtpDari.Value
                            ij.ShowDialog()
                            ij.Dispose()
                            If Me.ApakahLengkapKodeSalesPersonnya_SO(Me.dtpDari.Value, Me.dtpSampai.Value) = False Or _
                            Me.ApakahLengkapKodeSalesPersonnya_SO(Me.dtpDari2.Value, Me.dtpSampai2.Value) = False Then
                                MsgBox("There's still incomplete sales person data in one or several sales order.", MsgBoxStyle.Information, "Information")
                            End If
                            Me.BuatLaporanSO()
                        Else
                            Me.BuatLaporanSO()
                        End If
                    End If
                Else
                    Me.BuatLaporanSO()
                End If
            ElseIf Me.ReferenceData = EnumTipeReference.Invoice Or _
            Me.ReferenceData = EnumTipeReference.GrossMargin Then
                If Me.ApakahLengkapKodeSalesPersonnya_SO(#1/1/1950#, #1/1/2050#) = False Then
                    If Me.dtpDari2.Enabled = False Then
                        'cek period 1 saja
                        If Me.ApakahLengkapKodeSalesPersonnya_Invoice(Me.dtpDari.Value, Me.dtpSampai.Value) = False Then
                            Dim ij As New FrmTransaksiRepairInvoice
                            ij.dtpInvPeriod.Value = Me.dtpDari.Value
                            ij.ShowDialog()
                            ij.Dispose()
                            If Me.ApakahLengkapKodeSalesPersonnya_Invoice(Me.dtpDari.Value, Me.dtpSampai.Value) = False Then
                                MsgBox("There's still incomplete sales person data in one or several sales order.", MsgBoxStyle.Information, "Information")
                            End If
                            Me.BuatLaporanInvoice()
                        Else
                            Me.BuatLaporanInvoice()
                        End If
                    Else
                        'cek period 1 dan 2
                        If Me.ApakahLengkapKodeSalesPersonnya_Invoice(Me.dtpDari.Value, Me.dtpSampai.Value) = False Or _
                        Me.ApakahLengkapKodeSalesPersonnya_Invoice(Me.dtpDari2.Value, Me.dtpSampai2.Value) = False Then
                            Dim ij As New FrmTransaksiRepairInvoice
                            ij.dtpInvPeriod.Value = Me.dtpDari.Value
                            ij.ShowDialog()
                            ij.Dispose()
                            If Me.ApakahLengkapKodeSalesPersonnya_Invoice(Me.dtpDari.Value, Me.dtpSampai.Value) = False Or _
                            Me.ApakahLengkapKodeSalesPersonnya_Invoice(Me.dtpDari2.Value, Me.dtpSampai2.Value) = False Then
                                MsgBox("There's still incomplete sales person data in one or several sales order.", MsgBoxStyle.Information, "Information")
                            End If
                            Me.BuatLaporanInvoice()
                        Else
                            Me.BuatLaporanInvoice()
                        End If
                    End If
                Else
                    Me.BuatLaporanInvoice()
                End If
            Else
                MsgBox("Please specify reference data.", MsgBoxStyle.Information, "Information")
                Me.cboReference.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BuatLaporanInvoice()
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If Me.TipeLaporan = EnumTipeLaporan.NoReport Then
                MsgBox("Please specify report format.", MsgBoxStyle.Information, "Information")
                Me.cboFormat1.Focus()
            ElseIf Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Group_Recap Then
                If Me.cboFormat1.SelectedIndex = 0 Then
                    Me.BuatLaporanInvoice_SalesPersonGroupRecap()
                Else
                    Me.BuatLaporanInvoice_SalesPersonGroupRecap_2Period()
                End If
            ElseIf Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Recap Then
                If Me.cboFormat1.SelectedIndex = 0 Then
                    Me.BuatLaporanInvoice_SalesPersonRecap()
                Else
                    Me.BuatLaporanInvoice_SalesPersonRecap_2Period()
                End If
            ElseIf Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Detail Then
                If Me.cboFormat1.SelectedIndex = 0 Then
                    Me.BuatLaporanInvoice_SalesPersonDetail()
                Else
                    Me.BuatLaporanInvoice_SalesPersonDetail_2Period()
                End If
            ElseIf Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Group_Trend Then
                Me.BuatLaporanInvoice_TrendAmount(False, True)
            ElseIf Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Group_Trend_Perct Then
                Me.BuatLaporanInvoice_TrendAmount(True, True)
            ElseIf Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Trend Then
                Me.BuatLaporanInvoice_TrendAmount(False, False)
            ElseIf Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Trend_Perct Then
                Me.BuatLaporanInvoice_TrendAmount(True, False)
            ElseIf Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Composition Then
                Me.BuatLaporanInvoice_SalesComposition()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BuatLaporanSO()
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If Me.TipeLaporan = EnumTipeLaporan.NoReport Then
                MsgBox("Please specify report format.", MsgBoxStyle.Information, "Information")
                Me.cboFormat1.Focus()
            ElseIf Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Group_Recap Then
                If Me.cboFormat1.SelectedIndex = 0 Then
                    Me.BuatLaporanSO_SalesPersonGroupRecap()
                Else
                    Me.BuatLaporanSO_SalesPersonGroupRecap_2Period()
                End If
            ElseIf Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Recap Then
                If Me.cboFormat1.SelectedIndex = 0 Then
                    Me.BuatLaporanSO_SalesPersonRecap()
                Else
                    Me.BuatLaporanSO_SalesPersonRecap_2Period()
                End If
            ElseIf Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Detail Then
                If Me.cboFormat1.SelectedIndex = 0 Then
                    Me.BuatLaporanSO_SalesPersonDetail()
                Else
                    Me.BuatLaporanSO_SalesPersonDetail_2Period()
                End If
            ElseIf Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Group_Trend Then
                Me.BuatLaporanSO_TrendAmount(False, True)
            ElseIf Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Group_Trend_Perct Then
                Me.BuatLaporanSO_TrendAmount(True, True)
            ElseIf Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Trend Then
                Me.BuatLaporanSO_TrendAmount(False, False)
            ElseIf Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Trend_Perct Then
                Me.BuatLaporanSO_TrendAmount(True, False)
            ElseIf Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Composition Then
                Me.BuatLaporanSO_SalesComposition()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BuatLaporanSO_SalesComposition()
        Try
            If Me.TipeGroup = EnumTipeDetail.NoDetail Then
                Me.BuatLaporanSO_SalesComposition_NoGrouping()
            Else
                Me.BuatLaporanSO_SalesComposition_Grouping()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BuatLaporanSO_SalesComposition_Grouping()
        Dim Filter As String = ""
        Dim DTASO As New DseLaporanSalesPerformanceTableAdapters.tblsoTableAdapter
        Dim TBLSO As New DseLaporanSalesPerformance.tblsoDataTable
        Dim rowSO As DseLaporanSalesPerformance.tblsoRow
        Dim DV As DataRowView
        Dim TBLCetak As DseLaporanSalesPerformance.Sales_CompositionDataTable = DS.Sales_Composition
        Dim rowCetak As DseLaporanSalesPerformance.Sales_CompositionRow
        Dim DefPrec As Integer = 0
        Dim Total As Decimal = 0
        Dim ListGroup As New List(Of String)
        Dim Group As String = ""
        Dim TotalPerc As Decimal = 0
        Dim TBLDummy As New DseLaporanSalesPerformance.Sales_CompositionDataTable
        Dim rowDummy As DseLaporanSalesPerformance.Sales_CompositionRow

        Try
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
            DTASO.Connection.ConnectionString = Core_Module.GetConstring
            TBLCetak.Rows.Clear()
            DTASO.Fill(TBLSO, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value), Me.dtpDari.Value)
            Filter = Me.BuatFilterSO
            TBLSO.DefaultView.RowFilter = Filter

            'isikan trend tabel
            For Each DV In TBLSO.DefaultView
                rowSO = DV.Row

                'perhatikan nilai excludenya, kalo include, yang masuk ke laporan hanya nilai nett salesnya saja, tidak termasuk pajak
                If rowSO.Price_Exclude = False Then
                    rowSO.Amount = Math.Round((rowSO.Amount * (100 / (rowSO.Tax + 100))), _
                    rowSO.Default_Precision, MidpointRounding.AwayFromZero)
                End If
                'perhatikan exchange rate yang digunakan
                If rowSO.Primary_Currency <> 1 Then
                    rowSO.Amount = Math.Round(rowSO.Amount * rowSO.Exchange_Rate, DefPrec, MidpointRounding.AwayFromZero)
                End If

                If Me.TipeDetail = EnumTipeDetail.Detail_Cust_Group Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail 2] = '" + rowSO.Cust_Group_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Cust_Segmentation Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail 2] = '" + rowSO.Cust_Segmentation_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Cust_Type Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail 2] = '" + rowSO.Cust_Type_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail 2] = '" + rowSO.Product_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Brand Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail 2] = '" + rowSO.Brand_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Category Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail 2] = '" + rowSO.Category_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Group Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail 2] = '" + rowSO.Product_Group_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Sub_Brand Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail 2] = '" + rowSO.Sub_Brand_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Sales_Area Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail 2] = '" + rowSO.Sales_Area_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_SalesP Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail 2] = '" + rowSO.Sales_Person_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_SalesPGroup Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail 2] = '" + rowSO.Sales_Person_Group_Desc + "'"
                End If
                If Me.TipeGroup = EnumTipeDetail.Detail_Cust_Group Then
                    TBLCetak.DefaultView.RowFilter = TBLCetak.DefaultView.RowFilter + " AND [Sales Person Detail] = '" + rowSO.Cust_Group_Desc + "'"
                ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Cust_Segmentation Then
                    TBLCetak.DefaultView.RowFilter = TBLCetak.DefaultView.RowFilter + " AND [Sales Person Detail] = '" + rowSO.Cust_Segmentation_Desc + "'"
                ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Cust_Type Then
                    TBLCetak.DefaultView.RowFilter = TBLCetak.DefaultView.RowFilter + " AND [Sales Person Detail] = '" + rowSO.Cust_Type_Desc + "'"
                ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Product Then
                    TBLCetak.DefaultView.RowFilter = TBLCetak.DefaultView.RowFilter + " AND [Sales Person Detail] = '" + rowSO.Product_Desc + "'"
                ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Product_Brand Then
                    TBLCetak.DefaultView.RowFilter = TBLCetak.DefaultView.RowFilter + " AND [Sales Person Detail] = '" + rowSO.Brand_Desc + "'"
                ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Product_Category Then
                    TBLCetak.DefaultView.RowFilter = TBLCetak.DefaultView.RowFilter + " AND [Sales Person Detail] = '" + rowSO.Category_Desc + "'"
                ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Product_Group Then
                    TBLCetak.DefaultView.RowFilter = TBLCetak.DefaultView.RowFilter + " AND [Sales Person Detail] = '" + rowSO.Product_Group_Desc + "'"
                ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Product_Sub_Brand Then
                    TBLCetak.DefaultView.RowFilter = TBLCetak.DefaultView.RowFilter + " AND [Sales Person Detail] = '" + rowSO.Sub_Brand_Desc + "'"
                ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Sales_Area Then
                    TBLCetak.DefaultView.RowFilter = TBLCetak.DefaultView.RowFilter + " AND [Sales Person Detail] = '" + rowSO.Sales_Area_Desc + "'"
                ElseIf Me.TipeGroup = EnumTipeDetail.Detail_SalesP Then
                    TBLCetak.DefaultView.RowFilter = TBLCetak.DefaultView.RowFilter + " AND [Sales Person Detail] = '" + rowSO.Sales_Person_Desc + "'"
                ElseIf Me.TipeGroup = EnumTipeDetail.Detail_SalesPGroup Then
                    TBLCetak.DefaultView.RowFilter = TBLCetak.DefaultView.RowFilter + " AND [Sales Person Detail] = '" + rowSO.Sales_Person_Group_Desc + "'"
                End If
                If TBLCetak.DefaultView.Count = 0 Then
                    'tambahkan baris baru
                    rowCetak = TBLCetak.NewSales_CompositionRow
                    rowCetak.Amount = rowSO.Amount
                    rowCetak.Percentage = 0
                    rowCetak.Sub_Total_Amount = 0
                    rowCetak.Sub_Total_Percetage = 0
                    rowCetak.Sales_Person_Detail_2 = ""
                    If Me.TipeDetail = EnumTipeDetail.Detail_Cust_Group Then
                        rowCetak.Sales_Person_Detail_2 = rowSO.Cust_Group_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Cust_Segmentation Then
                        rowCetak.Sales_Person_Detail_2 = rowSO.Cust_Segmentation_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Cust_Type Then
                        rowCetak.Sales_Person_Detail_2 = rowSO.Cust_Type_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product Then
                        rowCetak.Sales_Person_Detail_2 = rowSO.Product_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Brand Then
                        rowCetak.Sales_Person_Detail_2 = rowSO.Brand_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Category Then
                        rowCetak.Sales_Person_Detail_2 = rowSO.Category_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Group Then
                        rowCetak.Sales_Person_Detail_2 = rowSO.Product_Group_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Sub_Brand Then
                        rowCetak.Sales_Person_Detail_2 = rowSO.Sub_Brand_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Sales_Area Then
                        rowCetak.Sales_Person_Detail_2 = rowSO.Sales_Area_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_SalesP Then
                        rowCetak.Sales_Person_Detail_2 = rowSO.Sales_Person_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_SalesPGroup Then
                        rowCetak.Sales_Person_Detail_2 = rowSO.Sales_Person_Group_Desc
                    End If
                    If Me.TipeGroup = EnumTipeDetail.Detail_Cust_Group Then
                        rowCetak.Sales_Person_Detail = rowSO.Cust_Group_Desc
                    ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Cust_Segmentation Then
                        rowCetak.Sales_Person_Detail = rowSO.Cust_Segmentation_Desc
                    ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Cust_Type Then
                        rowCetak.Sales_Person_Detail = rowSO.Cust_Type_Desc
                    ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Product Then
                        rowCetak.Sales_Person_Detail = rowSO.Product_Desc
                    ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Product_Brand Then
                        rowCetak.Sales_Person_Detail = rowSO.Brand_Desc
                    ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Product_Category Then
                        rowCetak.Sales_Person_Detail = rowSO.Category_Desc
                    ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Product_Group Then
                        rowCetak.Sales_Person_Detail = rowSO.Product_Group_Desc
                    ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Product_Sub_Brand Then
                        rowCetak.Sales_Person_Detail = rowSO.Sub_Brand_Desc
                    ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Sales_Area Then
                        rowCetak.Sales_Person_Detail = rowSO.Sales_Area_Desc
                    ElseIf Me.TipeGroup = EnumTipeDetail.Detail_SalesP Then
                        rowCetak.Sales_Person_Detail = rowSO.Sales_Person_Desc
                    ElseIf Me.TipeGroup = EnumTipeDetail.Detail_SalesPGroup Then
                        rowCetak.Sales_Person_Detail = rowSO.Sales_Person_Group_Desc
                    End If

                    TBLCetak.Rows.Add(rowCetak)
                Else
                    'masukkan amount saja
                    rowCetak = TBLCetak.DefaultView(0).Row
                    rowCetak.Amount += rowSO.Amount
                End If
            Next
            'hitung persetase
            ListGroup.Clear()
            For Each rowCetak In TBLCetak.Rows
                Total += rowCetak.Amount
                If ListGroup.Contains(rowCetak.Sales_Person_Detail) = False Then
                    ListGroup.Add(rowCetak.Sales_Person_Detail)
                End If
            Next
            For Each rowCetak In TBLCetak.Rows
                rowCetak.Percentage = Math.Round(rowCetak.Amount * 100 / Total, 2, MidpointRounding.AwayFromZero)
            Next
            'sub total persentase dan amount
            For Each Group In ListGroup
                TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + Group + "'"
                Total = 0
                TotalPerc = 0
                For Each DV In TBLCetak.DefaultView
                    rowCetak = DV.Row
                    Total += rowCetak.Amount
                    TotalPerc += rowCetak.Percentage
                Next
                For Each DV In TBLCetak.DefaultView
                    rowCetak = DV.Row
                    rowCetak.Sub_Total_Amount = Total
                    rowCetak.Sub_Total_Percetage = TotalPerc
                Next
            Next
            'urutkan
            For Each rowCetak In TBLCetak.Rows
                rowDummy = TBLDummy.NewSales_CompositionRow
                For ij As Integer = 0 To TBLDummy.Columns.Count - 1
                    rowDummy.Item(ij) = rowCetak.Item(ij)
                Next
                TBLDummy.Rows.Add(rowDummy)
            Next
            TBLDummy.DefaultView.RowFilter = ""
            TBLDummy.DefaultView.Sort = "[Sub Total Amount] DESC, [Sales Person Detail] ASC, Amount DESC, [Sales Person Detail 2] ASC"
            TBLCetak.Rows.Clear()
            For Each DV In TBLDummy.DefaultView
                rowDummy = DV.Row
                rowCetak = TBLCetak.NewSales_CompositionRow
                For ij As Integer = 0 To TBLDummy.Columns.Count - 1
                    rowCetak.Item(ij) = rowDummy.Item(ij)
                Next
                TBLCetak.Rows.Add(rowCetak)
            Next

            'tampilkan laporan 
            Me.rpt_Composition2.SetDataSource(DS)
            rpt_Composition2.SetParameterValue(rpt_Composition2.Parameter_Date1.ParameterFieldName, Me.dtpDari.Value)
            rpt_Composition2.SetParameterValue(rpt_Composition2.Parameter_Date2.ParameterFieldName, Me.dtpSampai.Value)
            rpt_Composition2.SetParameterValue(rpt_Composition2.Parameter_CrpLaporanSalesPerformance_CompositionDetail_DefPrec.ParameterFieldName, DefPrec)
            Me.crv1.ReportSource = rpt_Composition2
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASO.Dispose()
            TBLSO.Dispose()
            ListGroup = Nothing
            TBLDummy = Nothing
        End Try
    End Sub

    Private Sub BuatLaporanSO_SalesComposition_NoGrouping()
        Dim Filter As String = ""
        Dim DTASO As New DseLaporanSalesPerformanceTableAdapters.tblsoTableAdapter
        Dim TBLSO As New DseLaporanSalesPerformance.tblsoDataTable
        Dim rowSO As DseLaporanSalesPerformance.tblsoRow
        Dim DV As DataRowView
        Dim TBLCetak As DseLaporanSalesPerformance.Sales_CompositionDataTable = DS.Sales_Composition
        Dim rowCetak As DseLaporanSalesPerformance.Sales_CompositionRow
        Dim DefPrec As Integer = 0
        Dim Total As Decimal = 0
        Dim TBLDummy As New DseLaporanSalesPerformance.Sales_CompositionDataTable
        Dim rowDummy As DseLaporanSalesPerformance.Sales_CompositionRow

        Try
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
            DTASO.Connection.ConnectionString = Core_Module.GetConstring
            TBLCetak.Rows.Clear()
            DTASO.Fill(TBLSO, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value), Me.dtpDari.Value)
            Filter = Me.BuatFilterSO
            TBLSO.DefaultView.RowFilter = Filter

            'isikan trend tabel
            For Each DV In TBLSO.DefaultView
                rowSO = DV.Row

                'perhatikan nilai excludenya, kalo include, yang masuk ke laporan hanya nilai nett salesnya saja, tidak termasuk pajak
                If rowSO.Price_Exclude = False Then
                    rowSO.Amount = Math.Round((rowSO.Amount * (100 / (rowSO.Tax + 100))), _
                    rowSO.Default_Precision, MidpointRounding.AwayFromZero)
                End If
                'perhatikan exchange rate yang digunakan
                If rowSO.Primary_Currency <> 1 Then
                    rowSO.Amount = Math.Round(rowSO.Amount * rowSO.Exchange_Rate, DefPrec, MidpointRounding.AwayFromZero)
                End If

                If Me.TipeDetail = EnumTipeDetail.Detail_Cust_Group Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + rowSO.Cust_Group_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Cust_Segmentation Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + rowSO.Cust_Segmentation_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Cust_Type Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + rowSO.Cust_Type_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + rowSO.Product_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Brand Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + rowSO.Brand_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Category Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + rowSO.Category_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Group Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + rowSO.Product_Group_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Sub_Brand Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + rowSO.Sub_Brand_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Sales_Area Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + rowSO.Sales_Area_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_SalesP Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + rowSO.Sales_Person_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_SalesPGroup Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + rowSO.Sales_Person_Group_Desc + "'"
                End If
                If TBLCetak.DefaultView.Count = 0 Then
                    'tambahkan baris baru
                    rowCetak = TBLCetak.NewSales_CompositionRow
                    rowCetak.Amount = rowSO.Amount
                    rowCetak.Percentage = 0
                    rowCetak.Sub_Total_Amount = 0
                    rowCetak.Sub_Total_Percetage = 0
                    rowCetak.Sales_Person_Detail_2 = ""
                    If Me.TipeDetail = EnumTipeDetail.Detail_Cust_Group Then
                        rowCetak.Sales_Person_Detail = rowSO.Cust_Group_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Cust_Segmentation Then
                        rowCetak.Sales_Person_Detail = rowSO.Cust_Segmentation_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Cust_Type Then
                        rowCetak.Sales_Person_Detail = rowSO.Cust_Type_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product Then
                        rowCetak.Sales_Person_Detail = rowSO.Product_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Brand Then
                        rowCetak.Sales_Person_Detail = rowSO.Brand_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Category Then
                        rowCetak.Sales_Person_Detail = rowSO.Category_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Group Then
                        rowCetak.Sales_Person_Detail = rowSO.Product_Group_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Sub_Brand Then
                        rowCetak.Sales_Person_Detail = rowSO.Sub_Brand_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Sales_Area Then
                        rowCetak.Sales_Person_Detail = rowSO.Sales_Area_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_SalesP Then
                        rowCetak.Sales_Person_Detail = rowSO.Sales_Person_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_SalesPGroup Then
                        rowCetak.Sales_Person_Detail = rowSO.Sales_Person_Group_Desc
                    End If

                    TBLCetak.Rows.Add(rowCetak)
                Else
                    'masukkan amount saja
                    rowCetak = TBLCetak.DefaultView(0).Row
                    rowCetak.Amount += rowSO.Amount
                End If
            Next
            'hitung persetase
            For Each rowCetak In TBLCetak.Rows
                Total += rowCetak.Amount
            Next
            For Each rowCetak In TBLCetak.Rows
                rowCetak.Percentage = Math.Round(rowCetak.Amount * 100 / Total, 2, MidpointRounding.AwayFromZero)
            Next
            'urutkan
            For Each rowCetak In TBLCetak.Rows
                rowDummy = TBLDummy.NewSales_CompositionRow
                For ij As Integer = 0 To TBLDummy.Columns.Count - 1
                    rowDummy.Item(ij) = rowCetak.Item(ij)
                Next
                TBLDummy.Rows.Add(rowDummy)
            Next
            TBLDummy.DefaultView.RowFilter = ""
            TBLDummy.DefaultView.Sort = "Amount DESC, [Sales Person Detail] ASC"
            TBLCetak.Rows.Clear()
            For Each DV In TBLDummy.DefaultView
                rowDummy = DV.Row
                rowCetak = TBLCetak.NewSales_CompositionRow()
                For ij As Integer = 0 To TBLCetak.Columns.Count - 1
                    rowCetak.Item(ij) = rowDummy.Item(ij)
                Next
                TBLCetak.Rows.Add(rowCetak)
            Next

            'tampilkan laporan 
            Me.rpt_Composition.SetDataSource(DS)
            rpt_Composition.SetParameterValue(rpt_Composition.Parameter_Date1.ParameterFieldName, Me.dtpDari.Value)
            rpt_Composition.SetParameterValue(rpt_Composition.Parameter_Date2.ParameterFieldName, Me.dtpSampai.Value)
            rpt_Composition.SetParameterValue(rpt_Composition.Parameter_CrpLaporanSalesPerformance_CompositionDetail_DefPrec.ParameterFieldName, DefPrec)
            Me.crv1.ReportSource = rpt_Composition
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASO.Dispose()
            TBLSO.Dispose()
            TBLDummy.Dispose()
        End Try
    End Sub

    Private Sub BuatLaporanSO_TrendAmount(ByVal vUsePercent As Boolean, ByVal vSalesGroup As Boolean)
        Dim Filter As String = ""
        Dim DTASO As New DseLaporanSalesPerformanceTableAdapters.tblsoTableAdapter
        Dim TBLSO As New DseLaporanSalesPerformance.tblsoDataTable
        Dim rowSO As DseLaporanSalesPerformance.tblsoRow
        Dim DV As DataRowView
        Dim TBLCetak As DseLaporanSalesPerformance.Trend_AmountDataTable = DS.Trend_Amount
        Dim rowCetak As DseLaporanSalesPerformance.Trend_AmountRow
        Dim DefPrec As Integer = 0
        Dim TBLRefMonth As New DseLaporanSalesPerformance.Ref_PeriodDataTable
        Dim rowRefMonth As DseLaporanSalesPerformance.Ref_PeriodRow
        Dim StartRefPeriod As Date
        Dim Dummy As Integer = 0
        Dim DummyStr As String = ""

        Try
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
            DTASO.Connection.ConnectionString = Core_Module.GetConstring
            TBLCetak.Rows.Clear()
            DTASO.Fill(TBLSO, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value), Me.dtpDari.Value)
            Filter = Me.BuatFilterSO
            TBLSO.DefaultView.RowFilter = Filter

            'isikan ref period
            StartRefPeriod = Me.dtpDari.Value.Date
            TBLRefMonth.Rows.Clear()
            While StartRefPeriod <= Me.dtpSampai.Value.Date
                Dummy += 1
                rowRefMonth = TBLRefMonth.NewRef_PeriodRow
                With rowRefMonth
                    .Ref_Date = StartRefPeriod
                    DummyStr = Dummy.ToString
                    If DummyStr.Length = 1 Then
                        DummyStr = "0" + DummyStr
                    End If
                    .Month = DummyStr + ". " + Format(.Ref_Date, "MMM yyyy")
                End With
                TBLRefMonth.Rows.Add(rowRefMonth)

                StartRefPeriod = Me.ObjCD.GetFirstDateofTheMonth(StartRefPeriod + TimeSpan.FromDays(32))
            End While

            'isikan trend tabel
            For Each DV In TBLSO.DefaultView
                rowSO = DV.Row

                'perhatikan nilai excludenya, kalo include, yang masuk ke laporan hanya nilai nett salesnya saja, tidak termasuk pajak
                If rowSO.Price_Exclude = False Then
                    rowSO.Amount = Math.Round((rowSO.Amount * (100 / (rowSO.Tax + 100))), _
                    rowSO.Default_Precision, MidpointRounding.AwayFromZero)
                End If
                'perhatikan exchange rate yang digunakan
                If rowSO.Primary_Currency <> 1 Then
                    rowSO.Amount = Math.Round(rowSO.Amount * rowSO.Exchange_Rate, DefPrec, MidpointRounding.AwayFromZero)
                End If

                TBLRefMonth.DefaultView.RowFilter = "[Ref Date] = #" + Format(Me.ObjCD.GetFirstDateofTheMonth(rowSO.Tanggal_Transaksi), "MM/dd/yyyy") + "#"
                rowRefMonth = TBLRefMonth.DefaultView(0).Row
                If vSalesGroup = False Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + rowSO.Sales_Person_Desc + "' AND Month = '" + rowRefMonth.Month + "'"
                Else
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + rowSO.Sales_Person_Group_Desc + "' AND Month = '" + rowRefMonth.Month + "'"
                End If
                If TBLCetak.DefaultView.Count = 0 Then
                    'tambahkan baris baru
                    rowCetak = TBLCetak.NewTrend_AmountRow
                    rowCetak.Amount = rowSO.Amount
                    rowCetak.Month = rowRefMonth.Month
                    If vSalesGroup = False Then
                        rowCetak.Sales_Person_Detail = rowSO.Sales_Person_Desc
                    Else
                        rowCetak.Sales_Person_Detail = rowSO.Sales_Person_Group_Desc
                    End If

                    TBLCetak.Rows.Add(rowCetak)
                Else
                    'masukkan amount saja
                    rowCetak = TBLCetak.DefaultView(0).Row
                    rowCetak.Amount += rowSO.Amount
                End If
            Next

            If vUsePercent = False Then
                Me.rpt_Trend.SetDataSource(DS)
                rpt_Trend.SetParameterValue(rpt_Trend.Parameter_Date1.ParameterFieldName, Me.dtpDari.Value)
                rpt_Trend.SetParameterValue(rpt_Trend.Parameter_Date2.ParameterFieldName, Me.dtpSampai.Value)
                rpt_Trend.SetParameterValue(rpt_Trend.Parameter_DefPrec.ParameterFieldName, DefPrec)

                Me.crv1.ReportSource = rpt_Trend
            Else
                Me.rpt_Trend2.SetDataSource(DS)
                rpt_Trend2.SetParameterValue(rpt_Trend2.Parameter_Date1.ParameterFieldName, Me.dtpDari.Value)
                rpt_Trend2.SetParameterValue(rpt_Trend2.Parameter_Date2.ParameterFieldName, Me.dtpSampai.Value)
                rpt_Trend2.SetParameterValue(rpt_Trend2.Parameter_DefPrec.ParameterFieldName, DefPrec)

                Me.crv1.ReportSource = rpt_Trend2
            End If
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASO.Dispose()
            TBLSO.Dispose()
            TBLRefMonth.Dispose()
        End Try
    End Sub

    Private Function BuatFilterSO() As String
        Dim Filter As String = ""
        Dim DV As DataRowView
        Dim Filter1 As String = ""

        Try
            'sales person
            If Me.chkSalesPerson.Checked <> True Then
                If Me.clbSalesPerson.CheckedItems.Count <> Me.clbSalesPerson.Items.Count Then
                    For Each DV In Me.clbSalesPerson.CheckedItems
                        Me.vRowSP = DV.Row
                        If Filter1 = "" Then
                            Filter1 = "[Sales Person Desc] = '" + Me.vRowSP.Description + "'"
                        Else
                            Filter1 = Filter1 + " OR " + "[Sales Person Desc] = '" + Me.vRowSP.Description + "'"
                        End If
                    Next
                    If Filter1 <> "" Then
                        If Filter <> "" Then
                            Filter = Filter + " AND " + Filter1
                        Else
                            Filter = Filter1
                        End If
                    End If
                End If
            End If
            'sales person group
            If Me.chkSalesGroup.Checked <> True Then
                If Me.clbSalesPersonGroup.CheckedItems.Count <> Me.clbSalesPersonGroup.Items.Count Then
                    For Each DV In Me.clbSalesPersonGroup.CheckedItems
                        Me.vRowSG = DV.Row
                        If Filter = "" Then
                            Filter = "[Sales Person Group Desc] = '" + Me.vRowSG.Description + "'"
                        Else
                            Filter = Filter + " AND " + "[Sales Person Group Desc] = '" + Me.vRowSG.Description + "'"
                        End If
                    Next
                End If
            End If
            If Me.chkGoods.Checked = False Then
                'service saja
                If Filter = "" Then
                    Filter = "Type = " + CInt(Core_Module.clsMasterProduct.EnumTipeProduct.Services).ToString
                Else
                    Filter = Filter + " AND " + "Type = " + CInt(Core_Module.clsMasterProduct.EnumTipeProduct.Services).ToString
                End If
            ElseIf Me.chkService.Checked = False Then
                'goods saja
                If Filter = "" Then
                    Filter = "Type = " + CInt(Core_Module.clsMasterProduct.EnumTipeProduct.Goods).ToString
                Else
                    Filter = Filter + " AND " + "Type = " + CInt(Core_Module.clsMasterProduct.EnumTipeProduct.Goods).ToString
                End If
            End If

            Return Filter
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub BuatLaporanSO_SalesPersonRecap_2Period()
        Dim Filter As String = ""
        Dim DTASO As New DseLaporanSalesPerformanceTableAdapters.tblsoTableAdapter
        Dim TBLSO As New DseLaporanSalesPerformance.tblsoDataTable
        Dim rowSO As DseLaporanSalesPerformance.tblsoRow
        Dim DV As DataRowView
        Dim TBLCetak As DseLaporanSalesPerformance.Sales_Group_RecapDataTable = Me.DS.Sales_Group_Recap
        Dim TblDummy As New DseLaporanSalesPerformance.Sales_Group_RecapDataTable
        Dim rowCetak, rowDummy As DseLaporanSalesPerformance.Sales_Group_RecapRow
        Dim DefPrec As Integer = 0
        Dim DTABudget As New DseLaporanSalesPerformanceTableAdapters.tblbudgetTableAdapter
        Dim TBLBudget As New DseLaporanSalesPerformance.tblbudgetDataTable
        Dim TipeBudget As ClsTransaksiSalesBudget.EnumTipeBudget = ClsTransaksiSalesBudget.EnumTipeBudget.Invoice
        Dim ListSalesPersonGroup As New List(Of String)
        Dim SalesPersonG As String = ""
        Dim SubTotal As Decimal = 0

        Try
            TBLCetak.Rows.Clear()
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
            DTASO.Connection.ConnectionString = Core_Module.GetConstring
            DTABudget.Connection.ConnectionString = DTASO.Connection.ConnectionString

            DTASO.Fill(TBLSO, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value), Me.dtpDari.Value)
            Filter = Me.BuatFilterSO
            TBLSO.DefaultView.RowFilter = Filter
            For Each DV In TBLSO.DefaultView
                rowSO = DV.Row

                If ListSalesPersonGroup.Contains(rowSO.Sales_Person_Group_Code) = False Then
                    ListSalesPersonGroup.Add(rowSO.Sales_Person_Group_Code)
                End If

                'perhatikan nilai excludenya, kalo include, yang masuk ke laporan hanya nilai nett salesnya saja, tidak termasuk pajak
                If rowSO.Price_Exclude = False Then
                    rowSO.Amount = Math.Round((rowSO.Amount * (100 / (rowSO.Tax + 100))), _
                    rowSO.Default_Precision, MidpointRounding.AwayFromZero)
                End If
                'perhatikan exchange rate yang digunakan
                If rowSO.Primary_Currency <> 1 Then
                    rowSO.Amount = Math.Round(rowSO.Amount * rowSO.Exchange_Rate, DefPrec, MidpointRounding.AwayFromZero)
                End If

                If Me.cboFormat1.SelectedIndex = 1 Then
                    'budget
                    TblDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowSO.Sales_Person_Code + "'"
                    If TblDummy.DefaultView.Count = 0 Then
                        'buat baris baru
                        rowCetak = TblDummy.NewSales_Group_RecapRow
                        With rowCetak
                            .Amount1 = 0
                            .Amount2 = rowSO.Amount
                            .SalesP_Group_Code = rowSO.Sales_Person_Group_Code
                            .SalesP_Group_Name = rowSO.Sales_Person_Group_Desc
                            .Variance = 0
                            .Sales_P_Code = rowSO.Sales_Person_Code
                            .Sales_P_Name = rowSO.Sales_Person_Desc
                            .Sub_Total_1 = 0
                            .Sub_Total_2 = 0
                        End With
                        TblDummy.Rows.Add(rowCetak)
                    Else
                        'tambahkan amount saja
                        rowCetak = TblDummy.DefaultView(0).Row
                        rowCetak.Amount2 += rowSO.Amount
                    End If
                Else
                    '2 period
                    TblDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowSO.Sales_Person_Code + "'"
                    If TblDummy.DefaultView.Count = 0 Then
                        'buat baris baru
                        rowCetak = TblDummy.NewSales_Group_RecapRow
                        With rowCetak
                            .Amount1 = rowSO.Amount
                            .Amount2 = 0
                            .SalesP_Group_Code = rowSO.Sales_Person_Group_Code
                            .SalesP_Group_Name = rowSO.Sales_Person_Group_Desc
                            .Variance = 0
                            .Sales_P_Code = rowSO.Sales_Person_Code
                            .Sales_P_Name = rowSO.Sales_Person_Desc
                            .Sub_Total_1 = 0
                            .Sub_Total_2 = 0
                        End With
                        TblDummy.Rows.Add(rowCetak)
                    Else
                        'tambahkan amount saja
                        rowCetak = TblDummy.DefaultView(0).Row
                        rowCetak.Amount1 += rowSO.Amount
                    End If
                End If
            Next

            'tambahkan data periode ke-2
            If Me.cboFormat1.SelectedIndex = 1 Then
                'budget
                TipeBudget = ClsTransaksiSalesBudget.EnumTipeBudget.SO

                DTABudget.Fill(TBLBudget, TipeBudget, Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpSampai.Value), Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpDari.Value))
                For Each rowDummy In TblDummy.Rows
                    rowDummy.Amount1 = Me.GetBudget(rowDummy.Sales_P_Code, Me.dtpDari.Value, _
                    Me.ObjCD.GetFirstDateofTheMonth(Me.dtpSampai.Value), TBLBudget)
                Next
            Else
                'periode 2
                Me.IsikanAmount2_RecapSalesPerson(TBLSO, DTASO, Filter, TblDummy, DefPrec)
            End If

            'hitung variance
            For Each rowDummy In TblDummy.Rows
                If rowDummy.Amount1 = 0 Then
                    rowDummy.Variance = 0
                Else
                    rowDummy.Variance = Math.Round(((rowDummy.Amount2 - rowDummy.Amount1) * 100) / rowDummy.Amount1, 2, MidpointRounding.AwayFromZero)
                End If
            Next

            'sorting ulang berdasarkan penjualan paling tinggi, per group per sales person
            '1.hitung sub total per group
            For Each SalesPersonG In ListSalesPersonGroup
                TblDummy.DefaultView.RowFilter = "[SalesP Group Code] = '" + SalesPersonG + "'"
                SubTotal = 0
                For Each DV In TblDummy.DefaultView
                    rowDummy = DV.Row
                    SubTotal += rowDummy.Amount1
                Next
                For Each DV In TblDummy.DefaultView
                    rowDummy = DV.Row
                    rowDummy.Sub_Total_1 = SubTotal
                Next
            Next
            '2. sorting
            TblDummy.DefaultView.RowFilter = ""
            TblDummy.DefaultView.Sort = "[Sub Total 1] DESC, Amount1 DESC, [Sales P Code] ASC"
            For Each DV In TblDummy.DefaultView
                rowDummy = DV.Row
                rowCetak = TBLCetak.NewSales_Group_RecapRow
                For ij As Integer = 0 To TBLCetak.Columns.Count - 1
                    rowCetak.Item(ij) = rowDummy.Item(ij)
                Next
                TBLCetak.Rows.Add(rowCetak)
            Next

            Me.rpt_Recap4.SetDataSource(DS)
            rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Judul1.ParameterFieldName, "Sales Person Performance Report")
            If Me.cboFormat1.SelectedIndex = 1 Then
                'with budget
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Judul2.ParameterFieldName, "Sales Person Recap From Sales Order -  With Budget")
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Amount2.ParameterFieldName, "Amount")
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Amount1.ParameterFieldName, "Budget")
            Else
                '2 period comparison
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Judul2.ParameterFieldName, "Sales Person Recap From Sales Order -  2 Period")
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Amount1.ParameterFieldName, "Period I")
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Amount2.ParameterFieldName, "Period II")
            End If
            rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Date1.ParameterFieldName, Me.dtpDari.Value)
            rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Date2.ParameterFieldName, Me.dtpSampai.Value)
            rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Date3.ParameterFieldName, Me.dtpDari2.Value)
            rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Date4.ParameterFieldName, Me.dtpSampai2.Value)
            rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_DefPrec.ParameterFieldName, DefPrec)

            Me.crv1.ReportSource = rpt_Recap4
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASO.Dispose()
            TBLSO.Dispose()
            TblDummy.Dispose()
            DTABudget.Dispose()
            TBLBudget.Dispose()
        End Try
    End Sub

    Private Sub BuatLaporanSO_SalesPersonRecap()
        Dim Filter As String = ""
        Dim DTASO As New DseLaporanSalesPerformanceTableAdapters.tblsoTableAdapter
        Dim TBLSO As New DseLaporanSalesPerformance.tblsoDataTable
        Dim rowSO As DseLaporanSalesPerformance.tblsoRow
        Dim DV As DataRowView
        Dim TBLCetak As DseLaporanSalesPerformance.Sales_Group_RecapDataTable = Me.DS.Sales_Group_Recap
        Dim rowCetak, rowDummy As DseLaporanSalesPerformance.Sales_Group_RecapRow
        Dim DefPrec As Integer = 0
        Dim TBLDummy As New DseLaporanSalesPerformance.Sales_Group_RecapDataTable
        Dim ListSalesPersonGroup As New List(Of String)
        Dim SalesPersonG As String = ""
        Dim SubTotal As Decimal = 0

        Try
            TBLCetak.Rows.Clear()
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
            DTASO.Connection.ConnectionString = Core_Module.GetConstring
            DTASO.Fill(TBLSO, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value), Me.dtpDari.Value)

            Filter = Me.BuatFilterSO
            TBLSO.DefaultView.RowFilter = Filter
            For Each DV In TBLSO.DefaultView
                rowSO = DV.Row

                If ListSalesPersonGroup.Contains(rowSO.Sales_Person_Group_Code) = False Then
                    ListSalesPersonGroup.Add(rowSO.Sales_Person_Group_Code)
                End If

                'perhatikan nilai excludenya, kalo include, yang masuk ke laporan hanya nilai nett salesnya saja, tidak termasuk pajak
                If rowSO.Price_Exclude = False Then
                    rowSO.Amount = Math.Round((rowSO.Amount * (100 / (rowSO.Tax + 100))), _
                    rowSO.Default_Precision, MidpointRounding.AwayFromZero)
                End If
                'perhatikan exchange rate yang digunakan
                If rowSO.Primary_Currency <> 1 Then
                    rowSO.Amount = Math.Round(rowSO.Amount * rowSO.Exchange_Rate, DefPrec, MidpointRounding.AwayFromZero)
                End If

                TBLDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowSO.Sales_Person_Code + "'"
                If TBLDummy.DefaultView.Count = 0 Then
                    'buat baris baru
                    rowCetak = TBLDummy.NewSales_Group_RecapRow
                    With rowCetak
                        .Amount1 = rowSO.Amount
                        .Amount2 = 0
                        .SalesP_Group_Code = rowSO.Sales_Person_Group_Code
                        .SalesP_Group_Name = rowSO.Sales_Person_Group_Desc
                        .Sales_P_Code = rowSO.Sales_Person_Code
                        .Sales_P_Name = rowSO.Sales_Person_Desc
                        .Variance = 0
                        .Sub_Total_1 = 0
                        .Sub_Total_2 = 0
                    End With
                    TBLDummy.Rows.Add(rowCetak)
                Else
                    'tambahkan amount saja
                    rowCetak = TBLDummy.DefaultView(0).Row
                    rowCetak.Amount1 += rowSO.Amount
                End If
            Next

            'sorting ulang berdasarkan penjualan paling tinggi, per group per sales person
            '1.hitung sub total per group
            For Each SalesPersonG In ListSalesPersonGroup
                TBLDummy.DefaultView.RowFilter = "[SalesP Group Code] = '" + SalesPersonG + "'"
                SubTotal = 0
                For Each DV In TBLDummy.DefaultView
                    rowDummy = DV.Row
                    SubTotal += rowDummy.Amount1
                Next
                For Each DV In TBLDummy.DefaultView
                    rowDummy = DV.Row
                    rowDummy.Sub_Total_1 = SubTotal
                Next
            Next
            '2. sorting
            TBLDummy.DefaultView.RowFilter = ""
            TBLDummy.DefaultView.Sort = "[Sub Total 1] DESC, Amount1 DESC, [Sales P Code] ASC"
            For Each DV In TBLDummy.DefaultView
                rowDummy = DV.Row
                rowCetak = TBLCetak.NewSales_Group_RecapRow
                For ij As Integer = 0 To TBLCetak.Columns.Count - 1
                    rowCetak.Item(ij) = rowDummy.Item(ij)
                Next
                TBLCetak.Rows.Add(rowCetak)
            Next

            Me.rpt_Recap1.SetDataSource(DS)
            rpt_Recap1.SetParameterValue(rpt_Recap1.Parameter_Judul1.ParameterFieldName, "Sales Person Performance Report")
            rpt_Recap1.SetParameterValue(rpt_Recap1.Parameter_Judul2.ParameterFieldName, "Sales Person Recap From Sales Order")
            rpt_Recap1.SetParameterValue(rpt_Recap1.Parameter_Date1.ParameterFieldName, Me.dtpDari.Value)
            rpt_Recap1.SetParameterValue(rpt_Recap1.Parameter_Date2.ParameterFieldName, Me.dtpSampai.Value)
            rpt_Recap1.SetParameterValue(rpt_Recap1.Parameter_DefPrec.ParameterFieldName, DefPrec)

            Me.crv1.ReportSource = rpt_Recap1
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASO.Dispose()
            TBLSO.Dispose()
            TBLDummy.Dispose()
            ListSalesPersonGroup = Nothing
        End Try
    End Sub

    Private Sub BuatLaporanSO_SalesPersonGroupRecap_2Period()
        Dim Filter As String = ""
        Dim DTASO As New DseLaporanSalesPerformanceTableAdapters.tblsoTableAdapter
        Dim TBLSO As New DseLaporanSalesPerformance.tblsoDataTable
        Dim rowSO As DseLaporanSalesPerformance.tblsoRow
        Dim DV As DataRowView
        Dim TBLCetak As DseLaporanSalesPerformance.Sales_Group_RecapDataTable = Me.DS.Sales_Group_Recap
        Dim TblDummy As New DseLaporanSalesPerformance.Sales_Group_RecapDataTable
        Dim rowCetak, rowDummy As DseLaporanSalesPerformance.Sales_Group_RecapRow
        Dim DefPrec As Integer = 0
        Dim DTABudget As New DseLaporanSalesPerformanceTableAdapters.tblbudgetTableAdapter
        Dim TBLBudget As New DseLaporanSalesPerformance.tblbudgetDataTable
        Dim TipeBudget As ClsTransaksiSalesBudget.EnumTipeBudget = ClsTransaksiSalesBudget.EnumTipeBudget.Invoice

        Try
            TBLCetak.Rows.Clear()
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
            DTASO.Connection.ConnectionString = Core_Module.GetConstring
            DTABudget.Connection.ConnectionString = DTASO.Connection.ConnectionString

            DTASO.Fill(TBLSO, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value), Me.dtpDari.Value)
            Filter = Me.BuatFilterSO
            TBLSO.DefaultView.RowFilter = Filter
            For Each DV In TBLSO.DefaultView
                rowSO = DV.Row

                'perhatikan nilai excludenya, kalo include, yang masuk ke laporan hanya nilai nett salesnya saja, tidak termasuk pajak
                If rowSO.Price_Exclude = False Then
                    rowSO.Amount = Math.Round((rowSO.Amount * (100 / (rowSO.Tax + 100))), _
                    rowSO.Default_Precision, MidpointRounding.AwayFromZero)
                End If
                'perhatikan exchange rate yang digunakan
                If rowSO.Primary_Currency <> 1 Then
                    rowSO.Amount = Math.Round(rowSO.Amount * rowSO.Exchange_Rate, DefPrec, MidpointRounding.AwayFromZero)
                End If

                If Me.cboFormat1.SelectedIndex = 1 Then
                    'budget
                    TblDummy.DefaultView.RowFilter = "[SalesP Group Code] = '" + rowSO.Sales_Person_Group_Code + "'"
                    If TblDummy.DefaultView.Count = 0 Then
                        'buat baris baru
                        rowCetak = TblDummy.NewSales_Group_RecapRow
                        With rowCetak
                            .Amount1 = 0
                            .Amount2 = rowSO.Amount
                            .SalesP_Group_Code = rowSO.Sales_Person_Group_Code
                            .SalesP_Group_Name = rowSO.Sales_Person_Group_Desc
                            .Variance = 0
                            .Sales_P_Code = ""
                            .Sales_P_Name = ""
                            .Sub_Total_1 = 0
                            .Sub_Total_2 = 0
                        End With
                        TblDummy.Rows.Add(rowCetak)
                    Else
                        'tambahkan amount saja
                        rowCetak = TblDummy.DefaultView(0).Row
                        rowCetak.Amount2 += rowSO.Amount
                    End If
                Else
                    '2 period
                    TblDummy.DefaultView.RowFilter = "[SalesP Group Code] = '" + rowSO.Sales_Person_Group_Code + "'"
                    If TblDummy.DefaultView.Count = 0 Then
                        'buat baris baru
                        rowCetak = TblDummy.NewSales_Group_RecapRow
                        With rowCetak
                            .Amount1 = rowSO.Amount
                            .Amount2 = 0
                            .SalesP_Group_Code = rowSO.Sales_Person_Group_Code
                            .SalesP_Group_Name = rowSO.Sales_Person_Group_Desc
                            .Variance = 0
                            .Sales_P_Code = ""
                            .Sales_P_Name = ""
                            .Sub_Total_1 = 0
                            .Sub_Total_2 = 0
                        End With
                        TblDummy.Rows.Add(rowCetak)
                    Else
                        'tambahkan amount saja
                        rowCetak = TblDummy.DefaultView(0).Row
                        rowCetak.Amount1 += rowSO.Amount
                    End If
                End If
            Next

            'tambahkan data periode ke-2
            If Me.cboFormat1.SelectedIndex = 1 Then
                'budget
                TipeBudget = ClsTransaksiSalesBudget.EnumTipeBudget.SO

                DTABudget.Fill(TBLBudget, TipeBudget, Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpSampai.Value), Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpDari.Value))
                For Each rowDummy In TblDummy.Rows
                    rowDummy.Amount1 = Me.GetBudget_PerGroup(rowDummy.SalesP_Group_Code, Me.dtpDari.Value, _
                    Me.ObjCD.GetFirstDateofTheMonth(Me.dtpSampai.Value), TBLBudget)
                Next
            Else
                'periode 2
                Me.IsikanAmount2_RecapSalesPersonGroup(TBLSO, DTASO, Filter, TblDummy, DefPrec)
            End If

            'hitung variance
            For Each rowDummy In TblDummy.Rows
                If rowDummy.Amount1 = 0 Then
                    rowDummy.Variance = 0
                Else
                                          rowDummy.Variance = Math.Round(((rowDummy.Amount2 - rowDummy.Amount1) * 100) / rowDummy.Amount1, 2, MidpointRounding.AwayFromZero)
                End If
            Next

            'sorting ulang berdasarkan penjualan paling tinggi
            TblDummy.DefaultView.Sort = "Amount1 DESC, [SalesP Group Code] ASC"
            TblDummy.DefaultView.RowFilter = ""
            For Each DV In TblDummy.DefaultView
                rowDummy = DV.Row
                rowCetak = TBLCetak.NewSales_Group_RecapRow
                For ij As Integer = 0 To TBLCetak.Columns.Count - 1
                    rowCetak.Item(ij) = rowDummy.Item(ij)
                Next
                TBLCetak.Rows.Add(rowCetak)
            Next

            Me.rpt_Recap3.SetDataSource(DS)
            rpt_Recap3.SetParameterValue(rpt_Recap3.Parameter_Judul1.ParameterFieldName, "Sales Person Performance Report")
            If Me.cboFormat1.SelectedIndex = 1 Then
                'with budget
                rpt_Recap3.SetParameterValue(rpt_Recap3.Parameter_Judul2.ParameterFieldName, "Sales Person Group Recap From Sales Order -  With Budget")
                rpt_Recap3.SetParameterValue(rpt_Recap3.Parameter_Amount2.ParameterFieldName, "Amount")
                rpt_Recap3.SetParameterValue(rpt_Recap3.Parameter_Amount1.ParameterFieldName, "Budget")
            Else
                '2 period comparison
                rpt_Recap3.SetParameterValue(rpt_Recap3.Parameter_Judul2.ParameterFieldName, "Sales Person Group Recap From Sales Order -  2 Period")
                rpt_Recap3.SetParameterValue(rpt_Recap3.Parameter_Amount1.ParameterFieldName, "Period I")
                rpt_Recap3.SetParameterValue(rpt_Recap3.Parameter_Amount2.ParameterFieldName, "Period II")
            End If
            rpt_Recap3.SetParameterValue(rpt_Recap3.Parameter_Date1.ParameterFieldName, Me.dtpDari.Value)
            rpt_Recap3.SetParameterValue(rpt_Recap3.Parameter_Date2.ParameterFieldName, Me.dtpSampai.Value)
            rpt_Recap3.SetParameterValue(rpt_Recap3.Parameter_Date3.ParameterFieldName, Me.dtpDari2.Value)
            rpt_Recap3.SetParameterValue(rpt_Recap3.Parameter_Date4.ParameterFieldName, Me.dtpSampai2.Value)
            rpt_Recap3.SetParameterValue(rpt_Recap3.Parameter_DefPrec.ParameterFieldName, DefPrec)

            Me.crv1.ReportSource = rpt_Recap3
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASO.Dispose()
            TBLSO.Dispose()
            TblDummy.Dispose()
            DTABudget.Dispose()
            TBLBudget.Dispose()
        End Try
    End Sub

    Private Sub IsikanAmount2_RecapSalesPersonGroup(ByVal TBLSO As DseLaporanSalesPerformance.tblsoDataTable, _
    ByVal DTASO As DseLaporanSalesPerformanceTableAdapters.tblsoTableAdapter, ByVal Filter As String, _
    ByVal TblDummy As DseLaporanSalesPerformance.Sales_Group_RecapDataTable, ByVal DefPrec As Integer)
        Dim rowSO As DseLaporanSalesPerformance.tblsoRow
        Dim DV As DataRowView
        Dim rowCetak As DseLaporanSalesPerformance.Sales_Group_RecapRow

        Try
            DTASO.Fill(TBLSO, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai2.Value), Me.dtpDari2.Value)
            TBLSO.DefaultView.RowFilter = Filter

            For Each DV In TBLSO.DefaultView
                rowSO = DV.Row

                'perhatikan nilai excludenya, kalo include, yang masuk ke laporan hanya nilai nett salesnya saja, tidak termasuk pajak
                If rowSO.Price_Exclude = False Then
                    rowSO.Amount = Math.Round((rowSO.Amount * (100 / (rowSO.Tax + 100))), _
                    rowSO.Default_Precision, MidpointRounding.AwayFromZero)
                End If
                'perhatikan exchange rate yang digunakan
                If rowSO.Primary_Currency <> 1 Then
                    rowSO.Amount = Math.Round(rowSO.Amount * rowSO.Exchange_Rate, DefPrec, MidpointRounding.AwayFromZero)
                End If

                TblDummy.DefaultView.RowFilter = "[SalesP Group Code] = '" + rowSO.Sales_Person_Group_Code + "'"
                If TblDummy.DefaultView.Count = 0 Then
                    'buat baris baru
                    rowCetak = TblDummy.NewSales_Group_RecapRow
                    With rowCetak
                        .Amount1 = 0
                        .Amount2 = rowSO.Amount
                        .SalesP_Group_Code = rowSO.Sales_Person_Group_Code
                        .SalesP_Group_Name = rowSO.Sales_Person_Group_Desc
                        .Variance = 0
                        .Sales_P_Code = ""
                        .Sales_P_Name = ""
                        .Sub_Total_1 = 0
                        .Sub_Total_2 = 0
                    End With
                    TblDummy.Rows.Add(rowCetak)
                Else
                    'tambahkan amount saja
                    rowCetak = TblDummy.DefaultView(0).Row
                    rowCetak.Amount2 += rowSO.Amount
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanAmount2_RecapSalesPerson(ByVal TBLSO As DseLaporanSalesPerformance.tblsoDataTable, _
ByVal DTASO As DseLaporanSalesPerformanceTableAdapters.tblsoTableAdapter, ByVal Filter As String, _
ByVal TblDummy As DseLaporanSalesPerformance.Sales_Group_RecapDataTable, ByVal DefPrec As Integer)
        Dim rowSO As DseLaporanSalesPerformance.tblsoRow
        Dim DV As DataRowView
        Dim rowCetak As DseLaporanSalesPerformance.Sales_Group_RecapRow

        Try
            DTASO.Fill(TBLSO, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai2.Value), Me.dtpDari2.Value)
            TBLSO.DefaultView.RowFilter = Filter

            For Each DV In TBLSO.DefaultView
                rowSO = DV.Row

                'perhatikan nilai excludenya, kalo include, yang masuk ke laporan hanya nilai nett salesnya saja, tidak termasuk pajak
                If rowSO.Price_Exclude = False Then
                    rowSO.Amount = Math.Round((rowSO.Amount * (100 / (rowSO.Tax + 100))), _
                    rowSO.Default_Precision, MidpointRounding.AwayFromZero)
                End If
                'perhatikan exchange rate yang digunakan
                If rowSO.Primary_Currency <> 1 Then
                    rowSO.Amount = Math.Round(rowSO.Amount * rowSO.Exchange_Rate, DefPrec, MidpointRounding.AwayFromZero)
                End If

                TblDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowSO.Sales_Person_Code + "'"
                If TblDummy.DefaultView.Count = 0 Then
                    'buat baris baru
                    rowCetak = TblDummy.NewSales_Group_RecapRow
                    With rowCetak
                        .Amount1 = 0
                        .Amount2 = rowSO.Amount
                        .SalesP_Group_Code = rowSO.Sales_Person_Group_Code
                        .SalesP_Group_Name = rowSO.Sales_Person_Group_Desc
                        .Variance = 0
                        .Sales_P_Code = rowSO.Sales_Person_Code
                        .Sales_P_Name = rowSO.Sales_Person_Desc
                        .Sub_Total_1 = 0
                        .Sub_Total_2 = 0
                    End With
                    TblDummy.Rows.Add(rowCetak)
                Else
                    'tambahkan amount saja
                    rowCetak = TblDummy.DefaultView(0).Row
                    rowCetak.Amount2 += rowSO.Amount
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BuatLaporanSO_SalesPersonGroupRecap()
        Dim Filter As String = ""
        Dim DTASO As New DseLaporanSalesPerformanceTableAdapters.tblsoTableAdapter
        Dim TBLSO As New DseLaporanSalesPerformance.tblsoDataTable
        Dim rowSO As DseLaporanSalesPerformance.tblsoRow
        Dim DV As DataRowView
        Dim TBLCetak As DseLaporanSalesPerformance.Sales_Group_RecapDataTable = Me.DS.Sales_Group_Recap
        Dim TblDummy As New DseLaporanSalesPerformance.Sales_Group_RecapDataTable
        Dim rowCetak, rowDummy As DseLaporanSalesPerformance.Sales_Group_RecapRow
        Dim DefPrec As Integer = 0

        Try
            TBLCetak.Rows.Clear()
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
            DTASO.Connection.ConnectionString = Core_Module.GetConstring
            DTASO.Fill(TBLSO, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value), Me.dtpDari.Value)

            Filter = Me.BuatFilterSO
            TBLSO.DefaultView.RowFilter = Filter
            For Each DV In TBLSO.DefaultView
                rowSO = DV.Row

                'perhatikan nilai excludenya, kalo include, yang masuk ke laporan hanya nilai nett salesnya saja, tidak termasuk pajak
                If rowSO.Price_Exclude = False Then
                    rowSO.Amount = Math.Round((rowSO.Amount * (100 / (rowSO.Tax + 100))), _
                    rowSO.Default_Precision, MidpointRounding.AwayFromZero)
                End If
                'perhatikan exchange rate yang digunakan
                If rowSO.Primary_Currency <> 1 Then
                    rowSO.Amount = Math.Round(rowSO.Amount * rowSO.Exchange_Rate, DefPrec, MidpointRounding.AwayFromZero)
                End If

                TblDummy.DefaultView.RowFilter = "[SalesP Group Code] = '" + rowSO.Sales_Person_Group_Code + "'"
                If TblDummy.DefaultView.Count = 0 Then
                    'buat baris baru
                    rowCetak = TblDummy.NewSales_Group_RecapRow
                    With rowCetak
                        .Amount1 = rowSO.Amount
                        .Amount2 = 0
                        .SalesP_Group_Code = rowSO.Sales_Person_Group_Code
                        .SalesP_Group_Name = rowSO.Sales_Person_Group_Desc
                        .Variance = 0
                        .Sub_Total_1 = 0
                        .Sub_Total_2 = 0
                        .Sales_P_Code = ""
                        .Sales_P_Name = ""
                    End With
                    TblDummy.Rows.Add(rowCetak)
                Else
                    'tambahkan amount saja
                    rowCetak = TblDummy.DefaultView(0).Row
                    rowCetak.Amount1 += rowSO.Amount
                End If
            Next

            'sorting ulang berdasarkan penjualan paling tinggi
            TblDummy.DefaultView.Sort = "Amount1 DESC, [SalesP Group Code] ASC"
            TblDummy.DefaultView.RowFilter = ""
            For Each DV In TblDummy.DefaultView
                rowDummy = DV.Row
                rowCetak = TBLCetak.NewSales_Group_RecapRow
                For ij As Integer = 0 To TBLCetak.Columns.Count - 1
                    rowCetak.Item(ij) = rowDummy.Item(ij)
                Next
                TBLCetak.Rows.Add(rowCetak)
            Next

            Me.rpt_Recap2.SetDataSource(DS)
            rpt_Recap2.SetParameterValue(rpt_Recap2.Parameter_Judul1.ParameterFieldName, "Sales Person Performance Report")
            rpt_Recap2.SetParameterValue(rpt_Recap2.Parameter_Judul2.ParameterFieldName, "Sales Person Group Recap From Sales Order")
            rpt_Recap2.SetParameterValue(rpt_Recap2.Parameter_Judul3.ParameterFieldName, "Sales Person Group")
            rpt_Recap2.SetParameterValue(rpt_Recap2.Parameter_StartPeriod.ParameterFieldName, Me.dtpDari.Value)
            rpt_Recap2.SetParameterValue(rpt_Recap2.Parameter_EndPeriod.ParameterFieldName, Me.dtpSampai.Value)
            rpt_Recap2.SetParameterValue(rpt_Recap2.Parameter_DefPrec.ParameterFieldName, DefPrec)

            Me.crv1.ReportSource = rpt_Recap2
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASO.Dispose()
            TBLSO.Dispose()
            TblDummy.Dispose()
        End Try
    End Sub

    Private Function GetBudget(ByVal vSalesPersonCode As String, ByVal vStartDate As Date, ByVal vEndDate As Date, _
    ByVal vTBL As DseLaporanSalesPerformance.tblbudgetDataTable) As Decimal
        Dim Retval As Decimal = 0
        Dim row As DseLaporanSalesPerformance.tblbudgetRow
        Dim DV As DataRowView

        Try
            While vStartDate <= vEndDate
                vTBL.DefaultView.RowFilter = "[SalesP Code] = '" + vSalesPersonCode + "' AND period = #" + _
                Format(Me.ObjCD.GetFisrtDateOfTheYear(vStartDate), "MM/dd/yyyy") + "#"
                For Each DV In vTBL.DefaultView
                    row = DV.Row
                    row = vTBL.DefaultView(0).Row
                    If vStartDate.Month = 1 Then
                        Retval += row.jan
                    ElseIf vStartDate.Month = 2 Then
                        Retval += row.feb
                    ElseIf vStartDate.Month = 3 Then
                        Retval += row.mar
                    ElseIf vStartDate.Month = 4 Then
                        Retval += row.apr
                    ElseIf vStartDate.Month = 5 Then
                        Retval += row.mei
                    ElseIf vStartDate.Month = 6 Then
                        Retval += row.jun
                    ElseIf vStartDate.Month = 7 Then
                        Retval += row.jul
                    ElseIf vStartDate.Month = 8 Then
                        Retval += row.agt
                    ElseIf vStartDate.Month = 9 Then
                        Retval += row.sep
                    ElseIf vStartDate.Month = 10 Then
                        Retval += row.okt
                    ElseIf vStartDate.Month = 11 Then
                        Retval += row.nov
                    ElseIf vStartDate.Month = 12 Then
                        Retval += row.des
                    End If
                Next
                vStartDate = Me.ObjCD.GetFirstDateofTheMonth(vStartDate + TimeSpan.FromDays(32))
            End While

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetBudget_PerGroup(ByVal vSalesPersonGroupCode As String, ByVal vStartDate As Date, ByVal vEndDate As Date, _
    ByVal vTBL As DseLaporanSalesPerformance.tblbudgetDataTable) As Decimal
        Dim Retval As Decimal = 0
        Dim row As DseLaporanSalesPerformance.tblbudgetRow
        Dim DV As DataRowView

        Try
            While vStartDate <= vEndDate
                vTBL.DefaultView.RowFilter = "[Sales Group Code] = '" + vSalesPersonGroupCode + "' AND period = #" + Format(Me.ObjCD.GetFisrtDateOfTheYear(vStartDate), "MM/dd/yyyy") + "#"
                For Each DV In vTBL.DefaultView
                    row = DV.Row
                    row = vTBL.DefaultView(0).Row
                    If vStartDate.Month = 1 Then
                        Retval += row.jan
                    ElseIf vStartDate.Month = 2 Then
                        Retval += row.feb
                    ElseIf vStartDate.Month = 3 Then
                        Retval += row.mar
                    ElseIf vStartDate.Month = 4 Then
                        Retval += row.apr
                    ElseIf vStartDate.Month = 5 Then
                        Retval += row.mei
                    ElseIf vStartDate.Month = 6 Then
                        Retval += row.jun
                    ElseIf vStartDate.Month = 7 Then
                        Retval += row.jul
                    ElseIf vStartDate.Month = 8 Then
                        Retval += row.agt
                    ElseIf vStartDate.Month = 9 Then
                        Retval += row.sep
                    ElseIf vStartDate.Month = 10 Then
                        Retval += row.okt
                    ElseIf vStartDate.Month = 11 Then
                        Retval += row.nov
                    ElseIf vStartDate.Month = 12 Then
                        Retval += row.des
                    End If
                Next
                vStartDate = Me.ObjCD.GetFirstDateofTheMonth(vStartDate + TimeSpan.FromDays(32))
            End While

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BuatLaporanSO_SalesPersonDetail_2Period()
        Dim Filter As String = ""
        Dim DTASO As New DseLaporanSalesPerformanceTableAdapters.tblsoTableAdapter
        Dim TBLSO As New DseLaporanSalesPerformance.tblsoDataTable
        Dim DV As DataRowView
        Dim TBLCetak As DseLaporanSalesPerformance.Sales_Group_RecapDataTable = Me.DS.Sales_Group_Recap
        Dim rowCetak, rowDummy As DseLaporanSalesPerformance.Sales_Group_RecapRow
        Dim DefPrec As Integer = 0
        Dim TBLDummy As New DseLaporanSalesPerformance.Sales_Group_RecapDataTable
        Dim ListSalesPersonGroup As New List(Of String)
        Dim SalesPersonG As String = ""
        Dim SubTotal As Decimal = 0
        Dim DTABudget As New DseLaporanSalesPerformanceTableAdapters.tblbudgetTableAdapter
        Dim TBLBudget As New DseLaporanSalesPerformance.tblbudgetDataTable
        Dim TipeBudget As ClsTransaksiSalesBudget.EnumTipeBudget = ClsTransaksiSalesBudget.EnumTipeBudget.Invoice

        Try
            TBLCetak.Rows.Clear()
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
            DTASO.Connection.ConnectionString = Core_Module.GetConstring
            DTABudget.Connection.ConnectionString = DTASO.Connection.ConnectionString

            DTASO.Fill(TBLSO, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value), Me.dtpDari.Value)


            'tambahkan data periode ke-2
            If Me.TipeGroup = EnumTipeDetail.Detail_SalesP And Me.cboFormat1.SelectedIndex = 1 Then
                'isikan amount 2
                Me.IsikanTabelSO_Amount1(TBLSO, TBLDummy, DefPrec, ListSalesPersonGroup, True)

                'budget
                TipeBudget = ClsTransaksiSalesBudget.EnumTipeBudget.SO

                DTABudget.Fill(TBLBudget, TipeBudget, Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpSampai.Value), Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpDari.Value))
                For Each rowDummy In TBLDummy.Rows
                    rowDummy.Amount1 = Me.GetBudget(rowDummy.SalesP_Group_Code, Me.dtpDari.Value, _
                    Me.ObjCD.GetFirstDateofTheMonth(Me.dtpSampai.Value), TBLBudget)
                Next
            ElseIf Me.TipeGroup = EnumTipeDetail.Detail_SalesPGroup And Me.cboFormat1.SelectedIndex = 1 Then
                'isikan amount 2
                Me.IsikanTabelSO_Amount1(TBLSO, TBLDummy, DefPrec, ListSalesPersonGroup, True)

                'budget
                TipeBudget = ClsTransaksiSalesBudget.EnumTipeBudget.SO

                DTABudget.Fill(TBLBudget, TipeBudget, Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpSampai.Value), Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpDari.Value))
                For Each rowDummy In TBLDummy.Rows
                    rowDummy.Amount1 = Me.GetBudget_PerGroup(rowDummy.SalesP_Group_Code, Me.dtpDari.Value, _
                    Me.ObjCD.GetFirstDateofTheMonth(Me.dtpSampai.Value), TBLBudget)
                Next
            Else
                'isikan amount 1
                Me.IsikanTabelSO_Amount1(TBLSO, TBLDummy, DefPrec, ListSalesPersonGroup, False)
                'isikan amount 2
                TBLSO.Rows.Clear()
                DTASO.Fill(TBLSO, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai2.Value), Me.dtpDari2.Value)
                Me.IsikanTabelSO_Amount1(TBLSO, TBLDummy, DefPrec, ListSalesPersonGroup, True)
            End If

            'hitung variance
            For Each rowDummy In TBLDummy.Rows
                If rowDummy.Amount1 = 0 Then
                    rowDummy.Variance = 0
                Else
                                            rowDummy.Variance = Math.Round(((rowDummy.Amount2 - rowDummy.Amount1) * 100) / rowDummy.Amount1, 2, MidpointRounding.AwayFromZero)
                End If
            Next

            'sorting ulang berdasarkan penjualan paling tinggi, per group per sales person
            '1.hitung sub total per group
            For Each SalesPersonG In ListSalesPersonGroup
                TBLDummy.DefaultView.RowFilter = "[SalesP Group Code] = '" + SalesPersonG + "'"
                SubTotal = 0
                For Each DV In TBLDummy.DefaultView
                    rowDummy = DV.Row
                    SubTotal += rowDummy.Amount1
                Next
                For Each DV In TBLDummy.DefaultView
                    rowDummy = DV.Row
                    rowDummy.Sub_Total_1 = SubTotal
                Next
            Next
            For Each SalesPersonG In ListSalesPersonGroup
                TBLDummy.DefaultView.RowFilter = "[SalesP Group Code] = '" + SalesPersonG + "'"
                SubTotal = 0
                For Each DV In TBLDummy.DefaultView
                    rowDummy = DV.Row
                    SubTotal += rowDummy.Amount2
                Next
                For Each DV In TBLDummy.DefaultView
                    rowDummy = DV.Row
                    rowDummy.Sub_Total_2 = SubTotal
                Next
            Next
            '2. sorting
            TBLDummy.DefaultView.RowFilter = ""
            TBLDummy.DefaultView.Sort = "[Sub Total 1] DESC, [SalesP Group Code] ASC, Amount1 DESC, [Sales P Code] ASC"
            For Each DV In TBLDummy.DefaultView
                rowDummy = DV.Row
                rowCetak = TBLCetak.NewSales_Group_RecapRow
                For ij As Integer = 0 To TBLCetak.Columns.Count - 1
                    rowCetak.Item(ij) = rowDummy.Item(ij)
                Next
                TBLCetak.Rows.Add(rowCetak)
            Next

            If (Me.TipeGroup = EnumTipeDetail.Detail_SalesP And Me.cboFormat1.SelectedIndex = 1) Or _
            (Me.TipeGroup = EnumTipeDetail.Detail_SalesPGroup And Me.cboFormat1.SelectedIndex = 1) Then
                Me.rpt_Recap5.SetDataSource(DS)
                rpt_Recap5.SetParameterValue(rpt_Recap5.Parameter_Judul1.ParameterFieldName, "Sales Person Performance Report")
                rpt_Recap5.SetParameterValue(rpt_Recap5.Parameter_Judul2.ParameterFieldName, "Sales Person Detail From Sales Order -  With Budget")
                rpt_Recap5.SetParameterValue(rpt_Recap5.Parameter_Amount1.ParameterFieldName, "Budget")
                rpt_Recap5.SetParameterValue(rpt_Recap5.Parameter_Amount2.ParameterFieldName, "Amount")
                rpt_Recap5.SetParameterValue(rpt_Recap5.Parameter_Date1.ParameterFieldName, Me.dtpDari.Value)
                rpt_Recap5.SetParameterValue(rpt_Recap5.Parameter_Date2.ParameterFieldName, Me.dtpSampai.Value)
                rpt_Recap5.SetParameterValue(rpt_Recap5.Parameter_Date3.ParameterFieldName, Me.dtpDari2.Value)
                rpt_Recap5.SetParameterValue(rpt_Recap5.Parameter_Date4.ParameterFieldName, Me.dtpSampai2.Value)
                rpt_Recap5.SetParameterValue(rpt_Recap5.Parameter_DefPrec.ParameterFieldName, DefPrec)

                Me.crv1.ReportSource = rpt_Recap5
                Me.crv1.Refresh()
                Me.crv1.Show()
            Else
                Me.rpt_Recap4.SetDataSource(DS)
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Judul1.ParameterFieldName, "Sales Person Performance Report")
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Judul2.ParameterFieldName, "Sales Person Detail From Sales Order -  2 Period")
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Amount1.ParameterFieldName, "Period I")
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Amount2.ParameterFieldName, "Period II")
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Date1.ParameterFieldName, Me.dtpDari.Value)
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Date2.ParameterFieldName, Me.dtpSampai.Value)
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Date3.ParameterFieldName, Me.dtpDari2.Value)
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Date4.ParameterFieldName, Me.dtpSampai2.Value)
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_DefPrec.ParameterFieldName, DefPrec)

                Me.crv1.ReportSource = rpt_Recap4
                Me.crv1.Refresh()
                Me.crv1.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASO.Dispose()
            TBLSO.Dispose()
            TBLDummy.Dispose()
            ListSalesPersonGroup = Nothing
            DTABudget.Dispose()
            TBLBudget.Dispose()
        End Try
    End Sub

    Private Sub IsikanTabelSO_Amount1(ByVal TBLSO As DseLaporanSalesPerformance.tblsoDataTable, ByVal TBLDummy As DseLaporanSalesPerformance.Sales_Group_RecapDataTable, _
    ByVal DefPrec As Integer, ByVal ListSalesPersonGroup As List(Of String), ByVal vAmount2 As Boolean)
        Dim Filter As String = ""
        Dim DV As DataRowView
        Dim rowSO As DseLaporanSalesPerformance.tblsoRow
        Dim rowCetak As DseLaporanSalesPerformance.Sales_Group_RecapRow

        Try
            Filter = Me.BuatFilterSO
            TBLSO.DefaultView.RowFilter = Filter
            For Each DV In TBLSO.DefaultView
                rowSO = DV.Row

                'perhatikan nilai excludenya, kalo include, yang masuk ke laporan hanya nilai nett salesnya saja, tidak termasuk pajak
                If rowSO.Price_Exclude = False Then
                    rowSO.Amount = Math.Round((rowSO.Amount * (100 / (rowSO.Tax + 100))), _
                    rowSO.Default_Precision, MidpointRounding.AwayFromZero)
                End If
                'perhatikan exchange rate yang digunakan
                If rowSO.Primary_Currency <> 1 Then
                    rowSO.Amount = Math.Round(rowSO.Amount * rowSO.Exchange_Rate, DefPrec, MidpointRounding.AwayFromZero)
                End If

                If TipeDetail = EnumTipeDetail.Detail_Cust_Group Then
                    TBLDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowSO.Cust_Group_Code + "'"
                ElseIf TipeDetail = EnumTipeDetail.Detail_Cust_Segmentation Then
                    TBLDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowSO.Cust_Segmentation_Code + "'"
                ElseIf TipeDetail = EnumTipeDetail.Detail_Cust_Type Then
                    TBLDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowSO.Cust_Type_Code + "'"
                ElseIf TipeDetail = EnumTipeDetail.Detail_Product Then
                    TBLDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowSO.Product_Code + "'"
                ElseIf TipeDetail = EnumTipeDetail.Detail_Product_Brand Then
                    TBLDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowSO.Brand_Code + "'"
                ElseIf TipeDetail = EnumTipeDetail.Detail_Product_Category Then
                    TBLDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowSO.Category_Code + "'"
                ElseIf TipeDetail = EnumTipeDetail.Detail_Product_Group Then
                    TBLDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowSO.Product_Group_Code + "'"
                ElseIf TipeDetail = EnumTipeDetail.Detail_Product_Sub_Brand Then
                    TBLDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowSO.Sub_Brand_Code + "'"
                ElseIf TipeDetail = EnumTipeDetail.Detail_Sales_Area Then
                    TBLDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowSO.Sales_Area_Code + "'"
                ElseIf TipeDetail = EnumTipeDetail.Detail_SalesP Then
                    TBLDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowSO.Sales_Person_Code + "'"
                ElseIf TipeDetail = EnumTipeDetail.Detail_SalesPGroup Then
                    TBLDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowSO.Sales_Person_Group_Code + "'"
                Else
                    Exit Sub
                End If
                If TipeGroup = EnumTipeDetail.Detail_Cust_Group Then
                    If ListSalesPersonGroup.Contains(rowSO.Cust_Group_Code) = False Then
                        ListSalesPersonGroup.Add(rowSO.Cust_Group_Code)
                    End If
                    TBLDummy.DefaultView.RowFilter = TBLDummy.DefaultView.RowFilter + " AND " + "[SalesP Group Code] = '" + rowSO.Cust_Group_Code + "'"
                ElseIf TipeGroup = EnumTipeDetail.Detail_Cust_Segmentation Then
                    If ListSalesPersonGroup.Contains(rowSO.Cust_Segmentation_Code) = False Then
                        ListSalesPersonGroup.Add(rowSO.Cust_Segmentation_Code)
                    End If
                    TBLDummy.DefaultView.RowFilter = TBLDummy.DefaultView.RowFilter + " AND " + "[SalesP Group Code] = '" + rowSO.Cust_Segmentation_Code + "'"
                ElseIf TipeGroup = EnumTipeDetail.Detail_Cust_Type Then
                    If ListSalesPersonGroup.Contains(rowSO.Cust_Type_Code) = False Then
                        ListSalesPersonGroup.Add(rowSO.Cust_Type_Code)
                    End If
                    TBLDummy.DefaultView.RowFilter = TBLDummy.DefaultView.RowFilter + " AND " + "[SalesP Group Code] = '" + rowSO.Cust_Type_Code + "'"
                ElseIf TipeGroup = EnumTipeDetail.Detail_Product Then
                    If ListSalesPersonGroup.Contains(rowSO.Product_Code) = False Then
                        ListSalesPersonGroup.Add(rowSO.Product_Code)
                    End If
                    TBLDummy.DefaultView.RowFilter = TBLDummy.DefaultView.RowFilter + " AND " + "[SalesP Group Code] = '" + rowSO.Product_Code + "'"
                ElseIf TipeGroup = EnumTipeDetail.Detail_Product_Brand Then
                    If ListSalesPersonGroup.Contains(rowSO.Brand_Code) = False Then
                        ListSalesPersonGroup.Add(rowSO.Brand_Code)
                    End If
                    TBLDummy.DefaultView.RowFilter = TBLDummy.DefaultView.RowFilter + " AND " + "[SalesP Group Code] = '" + rowSO.Brand_Code + "'"
                ElseIf TipeGroup = EnumTipeDetail.Detail_Product_Category Then
                    If ListSalesPersonGroup.Contains(rowSO.Category_Code) = False Then
                        ListSalesPersonGroup.Add(rowSO.Category_Code)
                    End If
                    TBLDummy.DefaultView.RowFilter = TBLDummy.DefaultView.RowFilter + " AND " + "[SalesP Group Code] = '" + rowSO.Category_Code + "'"
                ElseIf TipeGroup = EnumTipeDetail.Detail_Product_Group Then
                    If ListSalesPersonGroup.Contains(rowSO.Product_Group_Code) = False Then
                        ListSalesPersonGroup.Add(rowSO.Product_Group_Code)
                    End If
                    TBLDummy.DefaultView.RowFilter = TBLDummy.DefaultView.RowFilter + " AND " + "[SalesP Group Code] = '" + rowSO.Product_Group_Code + "'"
                ElseIf TipeGroup = EnumTipeDetail.Detail_Product_Sub_Brand Then
                    If ListSalesPersonGroup.Contains(rowSO.Sub_Brand_Code) = False Then
                        ListSalesPersonGroup.Add(rowSO.Sub_Brand_Code)
                    End If
                    TBLDummy.DefaultView.RowFilter = TBLDummy.DefaultView.RowFilter + " AND " + "[SalesP Group Code] = '" + rowSO.Sub_Brand_Code + "'"
                ElseIf TipeGroup = EnumTipeDetail.Detail_Sales_Area Then
                    If ListSalesPersonGroup.Contains(rowSO.Sales_Area_Code) = False Then
                        ListSalesPersonGroup.Add(rowSO.Sales_Area_Code)
                    End If
                    TBLDummy.DefaultView.RowFilter = TBLDummy.DefaultView.RowFilter + " AND " + "[SalesP Group Code] = '" + rowSO.Sales_Area_Code + "'"
                ElseIf TipeGroup = EnumTipeDetail.Detail_SalesP Then
                    If ListSalesPersonGroup.Contains(rowSO.Sales_Person_Code) = False Then
                        ListSalesPersonGroup.Add(rowSO.Sales_Person_Code)
                    End If
                    TBLDummy.DefaultView.RowFilter = TBLDummy.DefaultView.RowFilter + " AND " + "[SalesP Group Code] = '" + rowSO.Sales_Person_Code + "'"
                ElseIf TipeGroup = EnumTipeDetail.Detail_SalesPGroup Then
                    If ListSalesPersonGroup.Contains(rowSO.Sales_Person_Group_Code) = False Then
                        ListSalesPersonGroup.Add(rowSO.Sales_Person_Group_Code)
                    End If
                    TBLDummy.DefaultView.RowFilter = TBLDummy.DefaultView.RowFilter + " AND " + "[SalesP Group Code] = '" + rowSO.Sales_Person_Group_Code + "'"
                Else
                    Exit Sub
                End If

                If TBLDummy.DefaultView.Count = 0 Then
                    'buat baris baru
                    rowCetak = TBLDummy.NewSales_Group_RecapRow
                    With rowCetak
                        If vAmount2 = False Then
                            .Amount1 = rowSO.Amount
                            .Amount2 = 0
                        Else
                            .Amount1 = 0
                            .Amount2 = rowSO.Amount
                        End If
                        If TipeDetail = EnumTipeDetail.Detail_Cust_Group Then
                            .Sales_P_Code = rowSO.Cust_Group_Code
                            .Sales_P_Name = rowSO.Cust_Group_Desc
                        ElseIf TipeDetail = EnumTipeDetail.Detail_Cust_Segmentation Then
                            .Sales_P_Code = rowSO.Cust_Segmentation_Code
                            .Sales_P_Name = rowSO.Cust_Segmentation_Desc
                        ElseIf TipeDetail = EnumTipeDetail.Detail_Cust_Type Then
                            .Sales_P_Code = rowSO.Cust_Type_Code
                            .Sales_P_Name = rowSO.Cust_Type_Desc
                        ElseIf TipeDetail = EnumTipeDetail.Detail_Product Then
                            .Sales_P_Code = rowSO.Product_Code
                            .Sales_P_Name = rowSO.Product_Desc
                        ElseIf TipeDetail = EnumTipeDetail.Detail_Product_Brand Then
                            .Sales_P_Code = rowSO.Brand_Code
                            .Sales_P_Name = rowSO.Brand_Desc
                        ElseIf TipeDetail = EnumTipeDetail.Detail_Product_Category Then
                            .Sales_P_Code = rowSO.Category_Code
                            .Sales_P_Name = rowSO.Category_Desc
                        ElseIf TipeDetail = EnumTipeDetail.Detail_Product_Group Then
                            .Sales_P_Code = rowSO.Product_Group_Code
                            .Sales_P_Name = rowSO.Product_Group_Desc
                        ElseIf TipeDetail = EnumTipeDetail.Detail_Product_Sub_Brand Then
                            .Sales_P_Code = rowSO.Sub_Brand_Code
                            .Sales_P_Name = rowSO.Sub_Brand_Desc
                        ElseIf TipeDetail = EnumTipeDetail.Detail_Sales_Area Then
                            .Sales_P_Code = rowSO.Sales_Area_Code
                            .Sales_P_Name = rowSO.Sales_Area_Desc
                        ElseIf TipeDetail = EnumTipeDetail.Detail_SalesP Then
                            .Sales_P_Code = rowSO.Sales_Person_Code
                            .Sales_P_Name = rowSO.Sales_Person_Desc
                        ElseIf TipeDetail = EnumTipeDetail.Detail_SalesPGroup Then
                            .Sales_P_Code = rowSO.Sales_Person_Group_Code
                            .Sales_P_Name = rowSO.Sales_Person_Group_Desc
                        Else
                            Exit Sub
                        End If
                        If TipeGroup = EnumTipeDetail.Detail_Cust_Group Then
                            .SalesP_Group_Code = rowSO.Cust_Group_Code
                            .SalesP_Group_Name = rowSO.Cust_Group_Desc
                        ElseIf TipeGroup = EnumTipeDetail.Detail_Cust_Segmentation Then
                            .SalesP_Group_Code = rowSO.Cust_Segmentation_Code
                            .SalesP_Group_Name = rowSO.Cust_Segmentation_Desc
                        ElseIf TipeGroup = EnumTipeDetail.Detail_Cust_Type Then
                            .SalesP_Group_Code = rowSO.Cust_Type_Code
                            .SalesP_Group_Name = rowSO.Cust_Type_Desc
                        ElseIf TipeGroup = EnumTipeDetail.Detail_Product Then
                            .SalesP_Group_Code = rowSO.Product_Code
                            .SalesP_Group_Name = rowSO.Product_Desc
                        ElseIf TipeGroup = EnumTipeDetail.Detail_Product_Brand Then
                            .SalesP_Group_Code = rowSO.Brand_Code
                            .SalesP_Group_Name = rowSO.Brand_Desc
                        ElseIf TipeGroup = EnumTipeDetail.Detail_Product_Category Then
                            .SalesP_Group_Code = rowSO.Category_Code
                            .SalesP_Group_Name = rowSO.Category_Desc
                        ElseIf TipeGroup = EnumTipeDetail.Detail_Product_Group Then
                            .SalesP_Group_Code = rowSO.Product_Group_Code
                            .SalesP_Group_Name = rowSO.Product_Group_Desc
                        ElseIf TipeGroup = EnumTipeDetail.Detail_Product_Sub_Brand Then
                            .SalesP_Group_Code = rowSO.Sub_Brand_Code
                            .SalesP_Group_Name = rowSO.Sub_Brand_Desc
                        ElseIf TipeGroup = EnumTipeDetail.Detail_Sales_Area Then
                            .SalesP_Group_Code = rowSO.Sales_Area_Code
                            .SalesP_Group_Name = rowSO.Sales_Area_Desc
                        ElseIf TipeGroup = EnumTipeDetail.Detail_SalesP Then
                            .SalesP_Group_Code = rowSO.Sales_Person_Code
                            .SalesP_Group_Name = rowSO.Sales_Person_Desc
                        ElseIf TipeGroup = EnumTipeDetail.Detail_SalesPGroup Then
                            .SalesP_Group_Code = rowSO.Sales_Person_Group_Code
                            .SalesP_Group_Name = rowSO.Sales_Person_Group_Desc
                        Else
                            Exit Sub
                        End If
                        .Variance = 0
                        .Sub_Total_1 = 0
                        .Sub_Total_2 = 0
                    End With
                    TBLDummy.Rows.Add(rowCetak)
                Else
                    'tambahkan amount saja
                    rowCetak = TBLDummy.DefaultView(0).Row
                    If vAmount2 = False Then
                        rowCetak.Amount1 += rowSO.Amount
                    Else
                        rowCetak.Amount2 += rowSO.Amount
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BuatLaporanSO_SalesPersonDetail()
        Dim Filter As String = ""
        Dim DTASO As New DseLaporanSalesPerformanceTableAdapters.tblsoTableAdapter
        Dim TBLSO As New DseLaporanSalesPerformance.tblsoDataTable
        Dim DV As DataRowView
        Dim TBLCetak As DseLaporanSalesPerformance.Sales_Group_RecapDataTable = Me.DS.Sales_Group_Recap
        Dim rowCetak, rowDummy As DseLaporanSalesPerformance.Sales_Group_RecapRow
        Dim DefPrec As Integer = 0
        Dim TBLDummy As New DseLaporanSalesPerformance.Sales_Group_RecapDataTable
        Dim ListSalesPersonGroup As New List(Of String)
        Dim SalesPersonG As String = ""
        Dim SubTotal As Decimal = 0

        Try
            TBLCetak.Rows.Clear()
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
            DTASO.Connection.ConnectionString = Core_Module.GetConstring
            DTASO.Fill(TBLSO, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value), Me.dtpDari.Value)

            'isikan amount 1
            Me.IsikanTabelSO_Amount1(TBLSO, TBLDummy, DefPrec, ListSalesPersonGroup, False)

            'sorting ulang berdasarkan penjualan paling tinggi, per group per sales person
            '1.hitung sub total per group
            For Each SalesPersonG In ListSalesPersonGroup
                TBLDummy.DefaultView.RowFilter = "[SalesP Group Code] = '" + SalesPersonG + "'"
                SubTotal = 0
                For Each DV In TBLDummy.DefaultView
                    rowDummy = DV.Row
                    SubTotal += rowDummy.Amount1
                Next
                For Each DV In TBLDummy.DefaultView
                    rowDummy = DV.Row
                    rowDummy.Sub_Total_1 = SubTotal
                Next
            Next
            '2. sorting
            TBLDummy.DefaultView.RowFilter = ""
            TBLDummy.DefaultView.Sort = "[Sub Total 1] DESC, [SalesP Group Code] ASC, Amount1 DESC, [Sales P Code] ASC"
            For Each DV In TBLDummy.DefaultView
                rowDummy = DV.Row
                rowCetak = TBLCetak.NewSales_Group_RecapRow
                For ij As Integer = 0 To TBLCetak.Columns.Count - 1
                    rowCetak.Item(ij) = rowDummy.Item(ij)
                Next
                TBLCetak.Rows.Add(rowCetak)
            Next

            Me.rpt_Recap1.SetDataSource(DS)
            rpt_Recap1.SetParameterValue(rpt_Recap1.Parameter_Judul1.ParameterFieldName, "Sales Person Performance Report")
            rpt_Recap1.SetParameterValue(rpt_Recap1.Parameter_Judul2.ParameterFieldName, "Sales Person Detail From Sales Order")
            rpt_Recap1.SetParameterValue(rpt_Recap1.Parameter_Date1.ParameterFieldName, Me.dtpDari.Value)
            rpt_Recap1.SetParameterValue(rpt_Recap1.Parameter_Date2.ParameterFieldName, Me.dtpSampai.Value)
            rpt_Recap1.SetParameterValue(rpt_Recap1.Parameter_DefPrec.ParameterFieldName, DefPrec)

            Me.crv1.ReportSource = rpt_Recap1
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASO.Dispose()
            TBLSO.Dispose()
            TBLDummy.Dispose()
            ListSalesPersonGroup = Nothing
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub AturGrouping()
        Try
            If Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Composition Then
                Me.cboGroup.Enabled = True
                Me.cboGroup.Items.Clear()
                Me.cboGroup.Items.Add("-")
                Me.cboGroup.Items.Add("Sales Person")
                Me.cboGroup.Items.Add("Sales Person Group")
                Me.cboGroup.Items.Add("Cust. Segmentation")
                Me.cboGroup.Items.Add("Sales Area")
                Me.cboGroup.Items.Add("Cust. Type")
                Me.cboGroup.Items.Add("Cust. Group")
                Me.cboGroup.Items.Add("Product Group")
                Me.cboGroup.Items.Add("Product Brand")
                Me.cboGroup.Items.Add("Product Sub Brand")
                Me.cboGroup.Items.Add("Product Category")
                Me.cboGroup.Items.Add("Product")
                Me.cboGroup.SelectedIndex = 0
            ElseIf Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Recap Or _
            Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Group_Recap Or _
            Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Group_Trend Or _
            Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Group_Trend_Perct Or _
            Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Trend Or _
            Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Trend_Perct Then
                Me.cboGroup.Enabled = False
                Me.cboGroup.Items.Clear()
                Me.cboGroup.Items.Add("-")
                Me.cboGroup.SelectedIndex = 0
            Else
                Me.cboGroup.Enabled = True
                Me.cboGroup.Items.Clear()
                Me.cboGroup.Items.Add("Sales Person")
                Me.cboGroup.Items.Add("Sales Person Group")
                Me.cboGroup.Items.Add("Cust. Segmentation")
                Me.cboGroup.Items.Add("Sales Area")
                Me.cboGroup.Items.Add("Cust. Type")
                Me.cboGroup.Items.Add("Cust. Group")
                Me.cboGroup.Items.Add("Product Group")
                Me.cboGroup.Items.Add("Product Brand")
                Me.cboGroup.Items.Add("Product Sub Brand")
                Me.cboGroup.Items.Add("Product Category")
                Me.cboGroup.Items.Add("Product")
                Me.cboGroup.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AturDetail()
        Try
            If Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Recap Or _
           Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Group_Recap Or _
           Me.TipeLaporan = EnumTipeLaporan.NoReport Or _
           Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Group_Trend Or _
           Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Group_Trend_Perct Or _
           Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Trend Or _
           Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Trend_Perct Then
                Me.cboFormat2.SelectedIndex = 0
                Me.cboFormat2.Enabled = False
                'Me.cboGroup.Text = "-"
                'Me.cboGroup.Enabled = False
            Else
                Me.cboFormat2.SelectedIndex = 0
                Me.cboFormat2.Enabled = True
                'Me.cboGroup.SelectedIndex = 1
                'Me.cboGroup.Enabled = True
            End If

            If Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Detail Or _
            Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Composition Then
                If Me.TipeGroup = EnumTipeDetail.NoDetail Then
                    If Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Composition Then
                        Me.cboFormat2.Enabled = True
                        Me.cboFormat2.Items.Clear()
                        Me.cboFormat2.Items.Add("Sales Person")
                        Me.cboFormat2.Items.Add("Sales Person Group")
                        Me.cboFormat2.Items.Add("Cust.Segmentation")
                        Me.cboFormat2.Items.Add("Sales Area")
                        Me.cboFormat2.Items.Add("Cust.Type")
                        Me.cboFormat2.Items.Add("Cust.Group")
                        Me.cboFormat2.Items.Add("Product Group")
                        Me.cboFormat2.Items.Add("Product Brand")
                        Me.cboFormat2.Items.Add("Product Sub Brand")
                        Me.cboFormat2.Items.Add("Product Category")
                        Me.cboFormat2.Items.Add("Product")
                        Me.cboFormat2.SelectedIndex = 0
                    Else
                        Me.cboFormat2.Enabled = False
                        Me.cboFormat2.Items.Clear()
                        Me.cboFormat2.Items.Add("-")
                        Me.cboFormat2.SelectedIndex = 0
                    End If
                ElseIf Me.TipeGroup = EnumTipeDetail.Detail_SalesP Then
                    Me.cboFormat2.Enabled = True
                    Me.cboFormat2.Items.Clear()
                    Me.cboFormat2.Items.Add("Cust.Segmentation")
                    Me.cboFormat2.Items.Add("Sales Area")
                    Me.cboFormat2.Items.Add("Cust.Type")
                    Me.cboFormat2.Items.Add("Cust.Group")
                    Me.cboFormat2.Items.Add("Product Group")
                    Me.cboFormat2.Items.Add("Product Brand")
                    Me.cboFormat2.Items.Add("Product Sub Brand")
                    Me.cboFormat2.Items.Add("Product Category")
                    Me.cboFormat2.Items.Add("Product")
                    Me.cboFormat2.SelectedIndex = 0
                ElseIf Me.TipeGroup = EnumTipeDetail.Detail_SalesPGroup Then
                    Me.cboFormat2.Enabled = True
                    Me.cboFormat2.Items.Clear()
                    If Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Composition Then
                        Me.cboFormat2.Items.Add("Sales Person")
                    End If
                    Me.cboFormat2.Items.Add("Cust.Segmentation")
                    Me.cboFormat2.Items.Add("Sales Area")
                    Me.cboFormat2.Items.Add("Cust.Type")
                    Me.cboFormat2.Items.Add("Cust.Group")
                    Me.cboFormat2.Items.Add("Product Group")
                    Me.cboFormat2.Items.Add("Product Brand")
                    Me.cboFormat2.Items.Add("Product Sub Brand")
                    Me.cboFormat2.Items.Add("Product Category")
                    Me.cboFormat2.Items.Add("Product")
                    Me.cboFormat2.SelectedIndex = 0
                Else
                    Me.cboFormat2.Enabled = True
                    Me.cboFormat2.Items.Clear()
                    Me.cboFormat2.Items.Add("Sales Person")
                    Me.cboFormat2.Items.Add("Sales Person Group")
                    Me.cboFormat2.SelectedIndex = 0
                End If
            ElseIf Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Group_Trend Or _
            Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Group_Trend_Perct Or _
            Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Trend Or _
            Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Trend_Perct Then
                Me.cboFormat2.Enabled = False
                Me.cboFormat2.Items.Clear()
                Me.cboFormat2.Items.Add("-")
                Me.cboFormat2.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Atur2Period()
        Try
            If Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Detail Then
                If Me.TipeGroup = EnumTipeDetail.Detail_SalesP Or Me.TipeGroup = EnumTipeDetail.Detail_SalesPGroup Then
                    If Me.cboFormat1.Items.Count = 2 Then
                        Me.cboFormat1.Items.Clear()
                        Me.cboFormat1.Items.Add("-")
                        Me.cboFormat1.Items.Add("With Budget")
                        Me.cboFormat1.Items.Add("2 Period Comparison")
                        Me.cboFormat1.SelectedIndex = 0
                    End If
                Else
                    If Me.cboFormat1.Items.Count = 3 Then
                        Me.cboFormat1.Items.Clear()
                        Me.cboFormat1.Items.Add("-")
                        Me.cboFormat1.Items.Add("2 Period Comparison")
                        Me.cboFormat1.SelectedIndex = 0
                    End If
                End If
                cboFormat1.Enabled = True
            ElseIf Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Group_Trend Or _
            Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Group_Trend_Perct Or _
            Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Trend Or _
            Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Trend_Perct Or _
            Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Composition Or _
            Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Group_Composition Then
                Me.cboFormat1.Enabled = False
                Me.cboFormat1.SelectedIndex = 0
            ElseIf Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Group_Recap Or _
            Me.TipeLaporan = EnumTipeLaporan.Sales_Person_Recap Then
                If Me.cboFormat1.Items.Count = 2 Then
                    Me.cboFormat1.Items.Clear()
                    Me.cboFormat1.Items.Add("-")
                    Me.cboFormat1.Items.Add("With Budget")
                    Me.cboFormat1.Items.Add("2 Period Comparison")
                    Me.cboFormat1.SelectedIndex = 0
                End If
                Me.cboFormat1.Enabled = True
            Else
                If Me.cboFormat1.Items.Count = 2 Then
                    Me.cboFormat1.Items.Clear()
                    Me.cboFormat1.Items.Add("-")
                    Me.cboFormat1.Items.Add("With Budget")
                    Me.cboFormat1.Items.Add("2 Period Comparison")
                    Me.cboFormat1.SelectedIndex = 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboGroup_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboGroup.SelectedIndexChanged
        Try
            If Me.cboGroup.Text = "-" Then
                Me.TipeGroup = EnumTipeDetail.NoDetail
            ElseIf Me.cboGroup.Text = "Cust. Segmentation" Then
                Me.TipeGroup = EnumTipeDetail.Detail_Cust_Segmentation
            ElseIf Me.cboGroup.Text = "Sales Area" Then
                Me.TipeGroup = EnumTipeDetail.Detail_Sales_Area
            ElseIf Me.cboGroup.Text = "Cust. Type" Then
                Me.TipeGroup = EnumTipeDetail.Detail_Cust_Type
            ElseIf Me.cboGroup.Text = "Cust. Group" Then
                Me.TipeGroup = EnumTipeDetail.Detail_Cust_Group
            ElseIf Me.cboGroup.Text = "Product Group" Then
                Me.TipeGroup = EnumTipeDetail.Detail_Product_Group
            ElseIf Me.cboGroup.Text = "Product Brand" Then
                Me.TipeGroup = EnumTipeDetail.Detail_Product_Brand
            ElseIf Me.cboGroup.Text = "Product Sub Brand" Then
                Me.TipeGroup = EnumTipeDetail.Detail_Product_Sub_Brand
            ElseIf Me.cboGroup.Text = "Product Category" Then
                Me.TipeGroup = EnumTipeDetail.Detail_Product_Category
            ElseIf Me.cboGroup.Text = "Product" Then
                Me.TipeGroup = EnumTipeDetail.Detail_Product
            ElseIf Me.cboGroup.Text = "Sales Person" Then
                Me.TipeGroup = EnumTipeDetail.Detail_SalesP
            ElseIf Me.cboGroup.Text = "Sales Person Group" Then
                Me.TipeGroup = EnumTipeDetail.Detail_SalesPGroup
            End If
            Me.AturDetail()
            Me.Atur2Period()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BuatLaporanInvoice_SalesPersonGroupRecap()
        Dim Filter As String = ""
        Dim DTAInv As New DseLaporanSalesPerformanceTableAdapters.tbltrans_salesinvoice_item1TableAdapter
        Dim TBLInv As New DseLaporanSalesPerformance.tbltrans_salesinvoice_item1DataTable
        Dim rowInv As DseLaporanSalesPerformance.tbltrans_salesinvoice_item1Row
        Dim DV As DataRowView
        Dim TBLCetak As DseLaporanSalesPerformance.Sales_Group_RecapDataTable = Me.DS.Sales_Group_Recap
        Dim TblDummy As New DseLaporanSalesPerformance.Sales_Group_RecapDataTable
        Dim rowCetak, rowDummy As DseLaporanSalesPerformance.Sales_Group_RecapRow
        Dim DefPrec As Integer = 0

        Try
            TBLCetak.Rows.Clear()
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
            DTAInv.Connection.ConnectionString = Core_Module.GetConstring
            DTAInv.Fill(TBLInv, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value), Me.dtpDari.Value)
            If Me.chkReturn.Checked = True Then
                Me.KurangkanItemReturn(TBLInv, Me.dtpDari.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value))
            End If

            Filter = Me.BuatFilterSO
            TBLInv.DefaultView.RowFilter = Filter
            For Each DV In TBLInv.DefaultView
                rowInv = DV.Row
                TblDummy.DefaultView.RowFilter = "[SalesP Group Code] = '" + rowInv.Sales_Person_Group_Code + "'"
                If TblDummy.DefaultView.Count = 0 Then
                    'buat baris baru
                    rowCetak = TblDummy.NewSales_Group_RecapRow
                    With rowCetak
                        If Me.ReferenceData = EnumTipeReference.Invoice Then
                            .Amount1 = rowInv.Sales_Amount
                        Else
                            .Amount1 = rowInv.Sales_Amount - rowInv.Cost
                        End If
                        .Amount2 = 0
                        .SalesP_Group_Code = rowInv.Sales_Person_Group_Code
                        .SalesP_Group_Name = rowInv.Sales_Person_Group_Desc
                        .Variance = 0
                        .Sub_Total_1 = 0
                        .Sub_Total_2 = 0
                        .Sales_P_Code = ""
                        .Sales_P_Name = ""
                    End With
                    TblDummy.Rows.Add(rowCetak)
                Else
                    'tambahkan amount saja
                    rowCetak = TblDummy.DefaultView(0).Row
                    If Me.ReferenceData = EnumTipeReference.Invoice Then
                        rowCetak.Amount1 += rowInv.Sales_Amount
                    Else
                        rowCetak.Amount1 += (rowInv.Sales_Amount - rowInv.Cost)
                    End If
                End If
            Next

            'sorting ulang berdasarkan penjualan paling tinggi
            TblDummy.DefaultView.Sort = "Amount1 DESC, [SalesP Group Code] ASC"
            TblDummy.DefaultView.RowFilter = ""
            For Each DV In TblDummy.DefaultView
                rowDummy = DV.Row
                rowCetak = TBLCetak.NewSales_Group_RecapRow
                For ij As Integer = 0 To TBLCetak.Columns.Count - 1
                    rowCetak.Item(ij) = rowDummy.Item(ij)
                Next
                TBLCetak.Rows.Add(rowCetak)
            Next

            Me.rpt_Recap2.SetDataSource(DS)
            rpt_Recap2.SetParameterValue(rpt_Recap2.Parameter_Judul1.ParameterFieldName, "Sales Person Performance Report")
            rpt_Recap2.SetParameterValue(rpt_Recap2.Parameter_Judul2.ParameterFieldName, "Sales Person Group Recap From Sales Invoice")
            rpt_Recap2.SetParameterValue(rpt_Recap2.Parameter_Judul3.ParameterFieldName, "Sales Person Group")
            rpt_Recap2.SetParameterValue(rpt_Recap2.Parameter_StartPeriod.ParameterFieldName, Me.dtpDari.Value)
            rpt_Recap2.SetParameterValue(rpt_Recap2.Parameter_EndPeriod.ParameterFieldName, Me.dtpSampai.Value)
            rpt_Recap2.SetParameterValue(rpt_Recap2.Parameter_DefPrec.ParameterFieldName, DefPrec)

            Me.crv1.ReportSource = rpt_Recap2
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAInv.Dispose()
            TBLInv.Dispose()
            TblDummy.Dispose()
        End Try
    End Sub

    Private Sub BuatLaporanInvoice_SalesPersonGroupRecap_2Period()
        Dim Filter As String = ""
        Dim DTAInv As New DseLaporanSalesPerformanceTableAdapters.tbltrans_salesinvoice_item1TableAdapter
        Dim TBLInv As New DseLaporanSalesPerformance.tbltrans_salesinvoice_item1DataTable
        Dim rowInv As DseLaporanSalesPerformance.tbltrans_salesinvoice_item1Row
        Dim DV As DataRowView
        Dim TBLCetak As DseLaporanSalesPerformance.Sales_Group_RecapDataTable = Me.DS.Sales_Group_Recap
        Dim TblDummy As New DseLaporanSalesPerformance.Sales_Group_RecapDataTable
        Dim rowCetak, rowDummy As DseLaporanSalesPerformance.Sales_Group_RecapRow
        Dim DefPrec As Integer = 0
        Dim DTABudget As New DseLaporanSalesPerformanceTableAdapters.tblbudgetTableAdapter
        Dim TBLBudget As New DseLaporanSalesPerformance.tblbudgetDataTable
        Dim TipeBudget As ClsTransaksiSalesBudget.EnumTipeBudget = ClsTransaksiSalesBudget.EnumTipeBudget.Invoice

        Try
            TBLCetak.Rows.Clear()
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
            DTAInv.Connection.ConnectionString = Core_Module.GetConstring
            DTABudget.Connection.ConnectionString = DTAInv.Connection.ConnectionString

            DTAInv.Fill(TBLInv, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value), Me.dtpDari.Value)
            If Me.chkReturn.Checked = True Then
                Me.KurangkanItemReturn(TBLInv, Me.dtpDari.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value))
            End If

            Filter = Me.BuatFilterSO
            TBLInv.DefaultView.RowFilter = Filter
            For Each DV In TBLInv.DefaultView
                rowInv = DV.Row

                If Me.cboFormat1.SelectedIndex = 1 Then
                    'budget
                    TblDummy.DefaultView.RowFilter = "[SalesP Group Code] = '" + rowInv.Sales_Person_Group_Code + "'"
                    If TblDummy.DefaultView.Count = 0 Then
                        'buat baris baru
                        rowCetak = TblDummy.NewSales_Group_RecapRow
                        With rowCetak
                            .Amount1 = 0
                            If Me.ReferenceData = EnumTipeReference.Invoice Then
                                .Amount2 = rowInv.Sales_Amount
                            Else
                                .Amount2 = rowInv.Sales_Amount - rowInv.Cost
                            End If
                            .SalesP_Group_Code = rowInv.Sales_Person_Group_Code
                            .SalesP_Group_Name = rowInv.Sales_Person_Group_Desc
                            .Variance = 0
                            .Sales_P_Code = ""
                            .Sales_P_Name = ""
                            .Sub_Total_1 = 0
                            .Sub_Total_2 = 0
                        End With
                        TblDummy.Rows.Add(rowCetak)
                    Else
                        'tambahkan amount saja
                        rowCetak = TblDummy.DefaultView(0).Row
                        If Me.ReferenceData = EnumTipeReference.Invoice Then
                            rowCetak.Amount2 += rowInv.Sales_Amount
                        Else
                            rowCetak.Amount2 += (rowInv.Sales_Amount - rowInv.Cost)
                        End If
                    End If
                Else
                    '2 period
                    TblDummy.DefaultView.RowFilter = "[SalesP Group Code] = '" + rowInv.Sales_Person_Group_Code + "'"
                    If TblDummy.DefaultView.Count = 0 Then
                        'buat baris baru
                        rowCetak = TblDummy.NewSales_Group_RecapRow
                        With rowCetak
                            If Me.ReferenceData = EnumTipeReference.Invoice Then
                                rowCetak.Amount1 = rowInv.Sales_Amount
                            Else
                                rowCetak.Amount1 = (rowInv.Sales_Amount - rowInv.Cost)
                            End If
                            .Amount2 = 0
                            .SalesP_Group_Code = rowInv.Sales_Person_Group_Code
                            .SalesP_Group_Name = rowInv.Sales_Person_Group_Desc
                            .Variance = 0
                            .Sales_P_Code = ""
                            .Sales_P_Name = ""
                            .Sub_Total_1 = 0
                            .Sub_Total_2 = 0
                        End With
                        TblDummy.Rows.Add(rowCetak)
                    Else
                        'tambahkan amount saja
                        rowCetak = TblDummy.DefaultView(0).Row
                        If Me.ReferenceData = EnumTipeReference.Invoice Then
                            rowCetak.Amount1 += rowInv.Sales_Amount
                        Else
                            rowCetak.Amount1 += (rowInv.Sales_Amount - rowInv.Cost)
                        End If
                    End If
                End If
            Next

            'tambahkan data periode ke-2
            If Me.cboFormat1.SelectedIndex = 1 Then
                'budget
                If Me.ReferenceData = EnumTipeReference.Invoice Then
                    TipeBudget = ClsTransaksiSalesBudget.EnumTipeBudget.Invoice
                ElseIf Me.ReferenceData = EnumTipeReference.GrossMargin Then
                    TipeBudget = ClsTransaksiSalesBudget.EnumTipeBudget.Invoice_GrossMargin
                End If

                DTABudget.Fill(TBLBudget, TipeBudget, Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpSampai.Value), Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpDari.Value))
                For Each rowDummy In TblDummy.Rows
                    rowDummy.Amount1 = Me.GetBudget_PerGroup(rowDummy.SalesP_Group_Code, Me.dtpDari.Value, _
                    Me.ObjCD.GetFirstDateofTheMonth(Me.dtpSampai.Value), TBLBudget)
                Next
            Else
                'periode 2
                Me.IsikanAmount2_Invoice_RecapSalesPersonGroup(TBLInv, DTAInv, Filter, TblDummy, DefPrec)
            End If

            'hitung variance
            For Each rowDummy In TblDummy.Rows
                If rowDummy.Amount1 = 0 Then
                    rowDummy.Variance = 0
                Else
                        rowDummy.Variance = Math.Round(((rowDummy.Amount2 - rowDummy.Amount1) * 100) / rowDummy.Amount1, 2, MidpointRounding.AwayFromZero)
                End If
            Next

            'sorting ulang berdasarkan penjualan paling tinggi
            TblDummy.DefaultView.Sort = "Amount1 DESC, [SalesP Group Code] ASC"
            TblDummy.DefaultView.RowFilter = ""
            For Each DV In TblDummy.DefaultView
                rowDummy = DV.Row
                rowCetak = TBLCetak.NewSales_Group_RecapRow
                For ij As Integer = 0 To TBLCetak.Columns.Count - 1
                    rowCetak.Item(ij) = rowDummy.Item(ij)
                Next
                TBLCetak.Rows.Add(rowCetak)
            Next

            Me.rpt_Recap3.SetDataSource(DS)
            rpt_Recap3.SetParameterValue(rpt_Recap3.Parameter_Judul1.ParameterFieldName, "Sales Person Performance Report")
            If Me.cboFormat1.SelectedIndex = 1 Then
                'with budget
                If Me.ReferenceData = EnumTipeReference.Invoice Then
                    rpt_Recap3.SetParameterValue(rpt_Recap3.Parameter_Judul2.ParameterFieldName, "Sales Person Group Recap From Sales Invoice -  With Budget")
                ElseIf Me.ReferenceData = EnumTipeReference.GrossMargin Then
                    rpt_Recap3.SetParameterValue(rpt_Recap3.Parameter_Judul2.ParameterFieldName, "Sales Person Group Recap From Gross Margin -  With Budget")
                End If
                rpt_Recap3.SetParameterValue(rpt_Recap3.Parameter_Amount2.ParameterFieldName, "Amount")
                rpt_Recap3.SetParameterValue(rpt_Recap3.Parameter_Amount1.ParameterFieldName, "Budget")
            Else
                '2 period comparison
                If Me.ReferenceData = EnumTipeReference.Invoice Then
                    rpt_Recap3.SetParameterValue(rpt_Recap3.Parameter_Judul2.ParameterFieldName, "Sales Person Group Recap From Sales Invoice -  2 Period")
                ElseIf Me.ReferenceData = EnumTipeReference.GrossMargin Then
                    rpt_Recap3.SetParameterValue(rpt_Recap3.Parameter_Judul2.ParameterFieldName, "Sales Person Group Recap From Gross Margin -  2 Period")
                End If
                rpt_Recap3.SetParameterValue(rpt_Recap3.Parameter_Amount1.ParameterFieldName, "Period I")
                rpt_Recap3.SetParameterValue(rpt_Recap3.Parameter_Amount2.ParameterFieldName, "Period II")
            End If
            rpt_Recap3.SetParameterValue(rpt_Recap3.Parameter_Date1.ParameterFieldName, Me.dtpDari.Value)
            rpt_Recap3.SetParameterValue(rpt_Recap3.Parameter_Date2.ParameterFieldName, Me.dtpSampai.Value)
            rpt_Recap3.SetParameterValue(rpt_Recap3.Parameter_Date3.ParameterFieldName, Me.dtpDari2.Value)
            rpt_Recap3.SetParameterValue(rpt_Recap3.Parameter_Date4.ParameterFieldName, Me.dtpSampai2.Value)
            rpt_Recap3.SetParameterValue(rpt_Recap3.Parameter_DefPrec.ParameterFieldName, DefPrec)

            Me.crv1.ReportSource = rpt_Recap3
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAInv.Dispose()
            TBLInv.Dispose()
            TblDummy.Dispose()
            DTABudget.Dispose()
            TBLBudget.Dispose()
        End Try
    End Sub

    Private Sub IsikanAmount2_Invoice_RecapSalesPersonGroup(ByVal TBLInv As DseLaporanSalesPerformance.tbltrans_salesinvoice_item1DataTable, _
ByVal DTAInv As DseLaporanSalesPerformanceTableAdapters.tbltrans_salesinvoice_item1TableAdapter, ByVal Filter As String, _
ByVal TblDummy As DseLaporanSalesPerformance.Sales_Group_RecapDataTable, ByVal DefPrec As Integer)
        Dim rowInv As DseLaporanSalesPerformance.tbltrans_salesinvoice_item1Row
        Dim DV As DataRowView
        Dim rowCetak As DseLaporanSalesPerformance.Sales_Group_RecapRow

        Try
            DTAInv.Fill(TBLInv, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai2.Value), Me.dtpDari2.Value)
            If Me.chkReturn.Checked = True Then
                Me.KurangkanItemReturn(TBLInv, Me.dtpDari2.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai2.Value))
            End If

            TBLInv.DefaultView.RowFilter = Filter

            For Each DV In TBLInv.DefaultView
                rowInv = DV.Row

                TblDummy.DefaultView.RowFilter = "[SalesP Group Code] = '" + rowInv.Sales_Person_Group_Code + "'"
                If TblDummy.DefaultView.Count = 0 Then
                    'buat baris baru
                    rowCetak = TblDummy.NewSales_Group_RecapRow
                    With rowCetak
                        .Amount1 = 0
                        If Me.ReferenceData = EnumTipeReference.Invoice Then
                            .Amount2 = rowInv.Sales_Amount
                        Else
                            .Amount2 = rowInv.Sales_Amount - rowInv.Cost
                        End If
                        .SalesP_Group_Code = rowInv.Sales_Person_Group_Code
                        .SalesP_Group_Name = rowInv.Sales_Person_Group_Desc
                        .Variance = 0
                        .Sales_P_Code = ""
                        .Sales_P_Name = ""
                        .Sub_Total_1 = 0
                        .Sub_Total_2 = 0
                    End With
                    TblDummy.Rows.Add(rowCetak)
                Else
                    'tambahkan amount saja
                    rowCetak = TblDummy.DefaultView(0).Row
                    If Me.ReferenceData = EnumTipeReference.Invoice Then
                        rowCetak.Amount2 += rowInv.Sales_Amount
                    Else
                        rowCetak.Amount2 += (rowInv.Sales_Amount - rowInv.Cost)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BuatLaporanInvoice_SalesPersonRecap()
        Dim Filter As String = ""
        Dim DTAInv As New DseLaporanSalesPerformanceTableAdapters.tbltrans_salesinvoice_item1TableAdapter
        Dim TBLInv As New DseLaporanSalesPerformance.tbltrans_salesinvoice_item1DataTable
        Dim rowInv As DseLaporanSalesPerformance.tbltrans_salesinvoice_item1Row
        Dim DV As DataRowView
        Dim TBLCetak As DseLaporanSalesPerformance.Sales_Group_RecapDataTable = Me.DS.Sales_Group_Recap
        Dim rowCetak, rowDummy As DseLaporanSalesPerformance.Sales_Group_RecapRow
        Dim DefPrec As Integer = 0
        Dim TBLDummy As New DseLaporanSalesPerformance.Sales_Group_RecapDataTable
        Dim ListSalesPersonGroup As New List(Of String)
        Dim SalesPersonG As String = ""
        Dim SubTotal As Decimal = 0

        Try
            TBLCetak.Rows.Clear()
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
            DTAInv.Connection.ConnectionString = Core_Module.GetConstring
            DTAInv.Fill(TBLInv, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value), Me.dtpDari.Value)
            If Me.chkReturn.Checked = True Then
                Me.KurangkanItemReturn(TBLInv, Me.dtpDari.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value))
            End If

            Filter = Me.BuatFilterSO
            TBLInv.DefaultView.RowFilter = Filter
            For Each DV In TBLInv.DefaultView
                rowInv = DV.Row

                If ListSalesPersonGroup.Contains(rowInv.Sales_Person_Group_Code) = False Then
                    ListSalesPersonGroup.Add(rowInv.Sales_Person_Group_Code)
                End If

                TBLDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowInv.Sales_Person_Code + "'"
                If TBLDummy.DefaultView.Count = 0 Then
                    'buat baris baru
                    rowCetak = TBLDummy.NewSales_Group_RecapRow
                    With rowCetak
                        If Me.ReferenceData = EnumTipeReference.GrossMargin Then
                            .Amount1 = rowInv.Sales_Amount - rowInv.Cost
                        Else
                            .Amount1 = rowInv.Sales_Amount
                        End If
                        .Amount2 = 0
                        .SalesP_Group_Code = rowInv.Sales_Person_Group_Code
                        .SalesP_Group_Name = rowInv.Sales_Person_Group_Desc
                        .Sales_P_Code = rowInv.Sales_Person_Code
                        .Sales_P_Name = rowInv.Sales_Person_Desc
                        .Variance = 0
                        .Sub_Total_1 = 0
                        .Sub_Total_2 = 0
                    End With
                    TBLDummy.Rows.Add(rowCetak)
                Else
                    'tambahkan amount saja
                    rowCetak = TBLDummy.DefaultView(0).Row
                    If Me.ReferenceData = EnumTipeReference.GrossMargin Then
                        rowCetak.Amount1 += (rowInv.Sales_Amount - rowInv.Cost)
                    Else
                        rowCetak.Amount1 += rowInv.Sales_Amount
                    End If
                End If
            Next

            'sorting ulang berdasarkan penjualan paling tinggi, per group per sales person
            '1.hitung sub total per group
            For Each SalesPersonG In ListSalesPersonGroup
                TBLDummy.DefaultView.RowFilter = "[SalesP Group Code] = '" + SalesPersonG + "'"
                SubTotal = 0
                For Each DV In TBLDummy.DefaultView
                    rowDummy = DV.Row
                    SubTotal += rowDummy.Amount1
                Next
                For Each DV In TBLDummy.DefaultView
                    rowDummy = DV.Row
                    rowDummy.Sub_Total_1 = SubTotal
                Next
            Next
            '2. sorting
            TBLDummy.DefaultView.RowFilter = ""
            TBLDummy.DefaultView.Sort = "[Sub Total 1] DESC, Amount1 DESC, [Sales P Code] ASC"
            For Each DV In TBLDummy.DefaultView
                rowDummy = DV.Row
                rowCetak = TBLCetak.NewSales_Group_RecapRow
                For ij As Integer = 0 To TBLCetak.Columns.Count - 1
                    rowCetak.Item(ij) = rowDummy.Item(ij)
                Next
                TBLCetak.Rows.Add(rowCetak)
            Next

            Me.rpt_Recap1.SetDataSource(DS)
            rpt_Recap1.SetParameterValue(rpt_Recap1.Parameter_Judul1.ParameterFieldName, "Sales Person Performance Report")
            If Me.ReferenceData = EnumTipeReference.GrossMargin Then
                rpt_Recap1.SetParameterValue(rpt_Recap1.Parameter_Judul2.ParameterFieldName, "Sales Person Recap From Gross Margin")
            Else
                rpt_Recap1.SetParameterValue(rpt_Recap1.Parameter_Judul2.ParameterFieldName, "Sales Person Recap From Sales Invoice")
            End If
            rpt_Recap1.SetParameterValue(rpt_Recap1.Parameter_Date1.ParameterFieldName, Me.dtpDari.Value)
            rpt_Recap1.SetParameterValue(rpt_Recap1.Parameter_Date2.ParameterFieldName, Me.dtpSampai.Value)
            rpt_Recap1.SetParameterValue(rpt_Recap1.Parameter_DefPrec.ParameterFieldName, DefPrec)

            Me.crv1.ReportSource = rpt_Recap1
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAInv.Dispose()
            TBLInv.Dispose()
            TBLDummy.Dispose()
            ListSalesPersonGroup = Nothing
        End Try
    End Sub

    Private Sub BuatLaporanInvoice_SalesPersonRecap_2Period()
        Dim Filter As String = ""
        Dim DTAInv As New DseLaporanSalesPerformanceTableAdapters.tbltrans_salesinvoice_item1TableAdapter
        Dim TBLInv As New DseLaporanSalesPerformance.tbltrans_salesinvoice_item1DataTable
        Dim rowInv As DseLaporanSalesPerformance.tbltrans_salesinvoice_item1Row
        Dim DV As DataRowView
        Dim TBLCetak As DseLaporanSalesPerformance.Sales_Group_RecapDataTable = Me.DS.Sales_Group_Recap
        Dim TblDummy As New DseLaporanSalesPerformance.Sales_Group_RecapDataTable
        Dim rowCetak, rowDummy As DseLaporanSalesPerformance.Sales_Group_RecapRow
        Dim DefPrec As Integer = 0
        Dim DTABudget As New DseLaporanSalesPerformanceTableAdapters.tblbudgetTableAdapter
        Dim TBLBudget As New DseLaporanSalesPerformance.tblbudgetDataTable
        Dim TipeBudget As ClsTransaksiSalesBudget.EnumTipeBudget = ClsTransaksiSalesBudget.EnumTipeBudget.Invoice
        Dim ListSalesPersonGroup As New List(Of String)
        Dim SalesPersonG As String = ""
        Dim SubTotal As Decimal = 0

        Try
            TBLCetak.Rows.Clear()
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
            DTAInv.Connection.ConnectionString = Core_Module.GetConstring
            DTABudget.Connection.ConnectionString = DTAInv.Connection.ConnectionString

            DTAInv.Fill(TBLInv, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value), Me.dtpDari.Value)
            If Me.chkReturn.Checked = True Then
                Me.KurangkanItemReturn(TBLInv, Me.dtpDari.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value))
            End If
            Filter = Me.BuatFilterSO
            TBLInv.DefaultView.RowFilter = Filter
            For Each DV In TBLInv.DefaultView
                rowInv = DV.Row

                If ListSalesPersonGroup.Contains(rowInv.Sales_Person_Group_Code) = False Then
                    ListSalesPersonGroup.Add(rowInv.Sales_Person_Group_Code)
                End If

                If Me.cboFormat1.SelectedIndex = 1 Then
                    'budget
                    TblDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowInv.Sales_Person_Code + "'"
                    If TblDummy.DefaultView.Count = 0 Then
                        'buat baris baru
                        rowCetak = TblDummy.NewSales_Group_RecapRow
                        With rowCetak
                            .Amount1 = 0
                            If Me.ReferenceData = EnumTipeReference.Invoice Then
                                .Amount2 = rowInv.Sales_Amount
                            Else
                                .Amount2 = rowInv.Sales_Amount - rowInv.Cost
                            End If
                            .SalesP_Group_Code = rowInv.Sales_Person_Group_Code
                            .SalesP_Group_Name = rowInv.Sales_Person_Group_Desc
                            .Variance = 0
                            .Sales_P_Code = rowInv.Sales_Person_Code
                            .Sales_P_Name = rowInv.Sales_Person_Desc
                            .Sub_Total_1 = 0
                            .Sub_Total_2 = 0
                        End With
                        TblDummy.Rows.Add(rowCetak)
                    Else
                        'tambahkan amount saja
                        rowCetak = TblDummy.DefaultView(0).Row
                        If Me.ReferenceData = EnumTipeReference.Invoice Then
                            rowCetak.Amount2 += rowInv.Sales_Amount
                        Else
                            rowCetak.Amount2 += (rowInv.Sales_Amount - rowInv.Cost)
                        End If
                    End If
                Else
                    '2 period
                    TblDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowInv.Sales_Person_Code + "'"
                    If TblDummy.DefaultView.Count = 0 Then
                        'buat baris baru
                        rowCetak = TblDummy.NewSales_Group_RecapRow
                        With rowCetak
                            If Me.ReferenceData = EnumTipeReference.Invoice Then
                                .Amount1 = rowInv.Sales_Amount
                            Else
                                .Amount1 = rowInv.Cost
                            End If
                            .Amount2 = 0
                            .SalesP_Group_Code = rowInv.Sales_Person_Group_Code
                            .SalesP_Group_Name = rowInv.Sales_Person_Group_Desc
                            .Variance = 0
                            .Sales_P_Code = rowInv.Sales_Person_Code
                            .Sales_P_Name = rowInv.Sales_Person_Desc
                            .Sub_Total_1 = 0
                            .Sub_Total_2 = 0
                        End With
                        TblDummy.Rows.Add(rowCetak)
                    Else
                        'tambahkan amount saja
                        rowCetak = TblDummy.DefaultView(0).Row
                        If Me.ReferenceData = EnumTipeReference.Invoice Then
                            rowCetak.Amount1 += rowInv.Sales_Amount
                        Else
                            rowCetak.Amount1 += (rowInv.Sales_Amount - rowInv.Cost)
                        End If
                    End If
                End If
            Next

            'tambahkan data periode ke-2
            If Me.cboFormat1.SelectedIndex = 1 Then
                'budget
                If Me.ReferenceData = EnumTipeReference.Invoice Then
                    TipeBudget = ClsTransaksiSalesBudget.EnumTipeBudget.Invoice
                Else
                    TipeBudget = ClsTransaksiSalesBudget.EnumTipeBudget.Invoice_GrossMargin
                End If

                DTABudget.Fill(TBLBudget, TipeBudget, Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpSampai.Value), Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpDari.Value))
                For Each rowDummy In TblDummy.Rows
                    rowDummy.Amount1 = Me.GetBudget(rowDummy.Sales_P_Code, Me.dtpDari.Value, _
                    Me.ObjCD.GetFirstDateofTheMonth(Me.dtpSampai.Value), TBLBudget)
                Next
            Else
                'periode 2
                Me.IsikanAmount2_Invoice_RecapSalesPerson(TBLInv, DTAInv, Filter, TblDummy, DefPrec)
            End If

            'hitung variance
            For Each rowDummy In TblDummy.Rows
                If rowDummy.Amount1 = 0 Then
                    rowDummy.Variance = 0
                Else
                    rowDummy.Variance = Math.Round(((rowDummy.Amount2 - rowDummy.Amount1) * 100) / rowDummy.Amount1, 2, MidpointRounding.AwayFromZero)
                End If
            Next

            'sorting ulang berdasarkan penjualan paling tinggi, per group per sales person
            '1.hitung sub total per group
            For Each SalesPersonG In ListSalesPersonGroup
                TblDummy.DefaultView.RowFilter = "[SalesP Group Code] = '" + SalesPersonG + "'"
                SubTotal = 0
                For Each DV In TblDummy.DefaultView
                    rowDummy = DV.Row
                    SubTotal += rowDummy.Amount1
                Next
                For Each DV In TblDummy.DefaultView
                    rowDummy = DV.Row
                    rowDummy.Sub_Total_1 = SubTotal
                Next
            Next
            '2. sorting
            TblDummy.DefaultView.RowFilter = ""
            TblDummy.DefaultView.Sort = "[Sub Total 1] DESC, Amount1 DESC, [Sales P Code] ASC"
            For Each DV In TblDummy.DefaultView
                rowDummy = DV.Row
                rowCetak = TBLCetak.NewSales_Group_RecapRow
                For ij As Integer = 0 To TBLCetak.Columns.Count - 1
                    rowCetak.Item(ij) = rowDummy.Item(ij)
                Next
                TBLCetak.Rows.Add(rowCetak)
            Next

            Me.rpt_Recap4.SetDataSource(DS)
            rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Judul1.ParameterFieldName, "Sales Person Performance Report")
            If Me.cboFormat1.SelectedIndex = 1 Then
                'with budget
                If Me.ReferenceData = EnumTipeReference.Invoice Then
                    rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Judul2.ParameterFieldName, "Sales Person Recap From Sales Invoice -  With Budget")
                Else
                    rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Judul2.ParameterFieldName, "Sales Person Recap From Gross Margin -  With Budget")
                End If
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Amount2.ParameterFieldName, "Amount")
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Amount1.ParameterFieldName, "Budget")
            Else
                '2 period comparison
                If Me.ReferenceData = EnumTipeReference.Invoice Then
                    rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Judul2.ParameterFieldName, "Sales Person Recap From Sales Invoice -  2 Period")
                Else
                    rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Judul2.ParameterFieldName, "Sales Person Recap From Gross Margin -  2 Period")
                End If
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Amount1.ParameterFieldName, "Period I")
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Amount2.ParameterFieldName, "Period II")
            End If
            rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Date1.ParameterFieldName, Me.dtpDari.Value)
            rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Date2.ParameterFieldName, Me.dtpSampai.Value)
            rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Date3.ParameterFieldName, Me.dtpDari2.Value)
            rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Date4.ParameterFieldName, Me.dtpSampai2.Value)
            rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_DefPrec.ParameterFieldName, DefPrec)

            Me.crv1.ReportSource = rpt_Recap4
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAInv.Dispose()
            TBLInv.Dispose()
            TblDummy.Dispose()
            DTABudget.Dispose()
            TBLBudget.Dispose()
        End Try
    End Sub

    Private Sub IsikanAmount2_Invoice_RecapSalesPerson(ByVal TBLInv As DseLaporanSalesPerformance.tbltrans_salesinvoice_item1DataTable, _
ByVal DTAInv As DseLaporanSalesPerformanceTableAdapters.tbltrans_salesinvoice_item1TableAdapter, ByVal Filter As String, _
ByVal TblDummy As DseLaporanSalesPerformance.Sales_Group_RecapDataTable, ByVal DefPrec As Integer)
        Dim rowInv As DseLaporanSalesPerformance.tbltrans_salesinvoice_item1Row
        Dim DV As DataRowView
        Dim rowCetak As DseLaporanSalesPerformance.Sales_Group_RecapRow

        Try
            DTAInv.Fill(TBLInv, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai2.Value), Me.dtpDari2.Value)
            If Me.chkReturn.Checked = True Then
                Me.KurangkanItemReturn(TBLInv, Me.dtpDari2.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai2.Value))
            End If
            TBLInv.DefaultView.RowFilter = Filter

            For Each DV In TBLInv.DefaultView
                rowInv = DV.Row

                TblDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowInv.Sales_Person_Code + "'"
                If TblDummy.DefaultView.Count = 0 Then
                    'buat baris baru
                    rowCetak = TblDummy.NewSales_Group_RecapRow
                    With rowCetak
                        .Amount1 = 0
                        If Me.ReferenceData = EnumTipeReference.GrossMargin Then
                            .Amount2 = (rowInv.Sales_Amount - rowInv.Cost)
                        Else
                            .Amount2 = rowInv.Sales_Amount
                        End If
                        .SalesP_Group_Code = rowInv.Sales_Person_Group_Code
                        .SalesP_Group_Name = rowInv.Sales_Person_Group_Desc
                        .Variance = 0
                        .Sales_P_Code = rowInv.Sales_Person_Code
                        .Sales_P_Name = rowInv.Sales_Person_Desc
                        .Sub_Total_1 = 0
                        .Sub_Total_2 = 0
                    End With
                    TblDummy.Rows.Add(rowCetak)
                Else
                    'tambahkan amount saja
                    rowCetak = TblDummy.DefaultView(0).Row
                    If Me.ReferenceData = EnumTipeReference.GrossMargin Then
                        rowCetak.Amount2 += (rowInv.Sales_Amount - rowInv.Cost)
                    Else
                        rowCetak.Amount2 += rowInv.Sales_Amount
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BuatLaporanInvoice_SalesPersonDetail()
        Dim Filter As String = ""
        Dim DTAInv As New DseLaporanSalesPerformanceTableAdapters.tbltrans_salesinvoice_item1TableAdapter
        Dim TBLInv As New DseLaporanSalesPerformance.tbltrans_salesinvoice_item1DataTable
        Dim DV As DataRowView
        Dim TBLCetak As DseLaporanSalesPerformance.Sales_Group_RecapDataTable = Me.DS.Sales_Group_Recap
        Dim rowCetak, rowDummy As DseLaporanSalesPerformance.Sales_Group_RecapRow
        Dim DefPrec As Integer = 0
        Dim TBLDummy As New DseLaporanSalesPerformance.Sales_Group_RecapDataTable
        Dim ListSalesPersonGroup As New List(Of String)
        Dim SalesPersonG As String = ""
        Dim SubTotal As Decimal = 0

        Try
            TBLCetak.Rows.Clear()
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
            DTAInv.Connection.ConnectionString = Core_Module.GetConstring
            DTAInv.Fill(TBLInv, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value), Me.dtpDari.Value)
            If Me.chkReturn.Checked = True Then
                Me.KurangkanItemReturn(TBLInv, Me.dtpDari.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value))
            End If

            'isikan amount 1
            Me.IsikanTabelInvoice_Amount1(TBLInv, TBLDummy, DefPrec, ListSalesPersonGroup, False)
            'sorting ulang berdasarkan penjualan paling tinggi, per group per sales person
            '1.hitung sub total per group
            For Each SalesPersonG In ListSalesPersonGroup
                TBLDummy.DefaultView.RowFilter = "[SalesP Group Code] = '" + SalesPersonG + "'"
                SubTotal = 0
                For Each DV In TBLDummy.DefaultView
                    rowDummy = DV.Row
                    SubTotal += rowDummy.Amount1
                Next
                For Each DV In TBLDummy.DefaultView
                    rowDummy = DV.Row
                    rowDummy.Sub_Total_1 = SubTotal
                Next
            Next
            '2. sorting
            TBLDummy.DefaultView.RowFilter = ""
            TBLDummy.DefaultView.Sort = "[Sub Total 1] DESC, [SalesP Group Code] ASC, Amount1 DESC, [Sales P Code] ASC"
            For Each DV In TBLDummy.DefaultView
                rowDummy = DV.Row
                rowCetak = TBLCetak.NewSales_Group_RecapRow
                For ij As Integer = 0 To TBLCetak.Columns.Count - 1
                    rowCetak.Item(ij) = rowDummy.Item(ij)
                Next
                TBLCetak.Rows.Add(rowCetak)
            Next

            Me.rpt_Recap1.SetDataSource(DS)
            rpt_Recap1.SetParameterValue(rpt_Recap1.Parameter_Judul1.ParameterFieldName, "Sales Person Performance Report")
            If Me.ReferenceData = EnumTipeReference.Invoice Then
                rpt_Recap1.SetParameterValue(rpt_Recap1.Parameter_Judul2.ParameterFieldName, "Sales Person Detail From Sales Invoice")
            Else
                rpt_Recap1.SetParameterValue(rpt_Recap1.Parameter_Judul2.ParameterFieldName, "Sales Person Detail From Gross Margin")
            End If
            rpt_Recap1.SetParameterValue(rpt_Recap1.Parameter_Date1.ParameterFieldName, Me.dtpDari.Value)
            rpt_Recap1.SetParameterValue(rpt_Recap1.Parameter_Date2.ParameterFieldName, Me.dtpSampai.Value)
            rpt_Recap1.SetParameterValue(rpt_Recap1.Parameter_DefPrec.ParameterFieldName, DefPrec)

            Me.crv1.ReportSource = rpt_Recap1
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAInv.Dispose()
            TBLInv.Dispose()
            TBLDummy.Dispose()
            ListSalesPersonGroup = Nothing
        End Try
    End Sub

    Private Sub IsikanTabelInvoice_Amount1(ByVal TBLInv As DseLaporanSalesPerformance.tbltrans_salesinvoice_item1DataTable, ByVal TBLDummy As DseLaporanSalesPerformance.Sales_Group_RecapDataTable, _
ByVal DefPrec As Integer, ByVal ListSalesPersonGroup As List(Of String), ByVal vAmount2 As Boolean)
        Dim Filter As String = ""
        Dim DV As DataRowView
        Dim rowInv As DseLaporanSalesPerformance.tbltrans_salesinvoice_item1Row
        Dim rowCetak As DseLaporanSalesPerformance.Sales_Group_RecapRow

        Try
            Filter = Me.BuatFilterSO
            TBLInv.DefaultView.RowFilter = Filter
            For Each DV In TBLInv.DefaultView
                rowInv = DV.Row

                If TipeDetail = EnumTipeDetail.Detail_Cust_Group Then
                    TBLDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowInv.Cust_Group_Code + "'"
                ElseIf TipeDetail = EnumTipeDetail.Detail_Cust_Segmentation Then
                    TBLDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowInv.Segmentation_Code + "'"
                ElseIf TipeDetail = EnumTipeDetail.Detail_Cust_Type Then
                    TBLDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowInv.Cust_Type_Code + "'"
                ElseIf TipeDetail = EnumTipeDetail.Detail_Product Then
                    TBLDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowInv.Product_Code + "'"
                ElseIf TipeDetail = EnumTipeDetail.Detail_Product_Brand Then
                    TBLDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowInv.Brand_Code + "'"
                ElseIf TipeDetail = EnumTipeDetail.Detail_Product_Category Then
                    TBLDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowInv.Category_Code + "'"
                ElseIf TipeDetail = EnumTipeDetail.Detail_Product_Group Then
                    TBLDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowInv.Product_Group_Code + "'"
                ElseIf TipeDetail = EnumTipeDetail.Detail_Product_Sub_Brand Then
                    TBLDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowInv.Sub_Brand_Code + "'"
                ElseIf TipeDetail = EnumTipeDetail.Detail_Sales_Area Then
                    TBLDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowInv.Sales_Area_Code + "'"
                ElseIf TipeDetail = EnumTipeDetail.Detail_SalesP Then
                    TBLDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowInv.Sales_Person_Code + "'"
                ElseIf TipeDetail = EnumTipeDetail.Detail_SalesPGroup Then
                    TBLDummy.DefaultView.RowFilter = "[Sales P Code] = '" + rowInv.Sales_Person_Group_Code + "'"
                Else
                    Exit Sub
                End If
                If TipeGroup = EnumTipeDetail.Detail_Cust_Group Then
                    If ListSalesPersonGroup.Contains(rowInv.Cust_Group_Code) = False Then
                        ListSalesPersonGroup.Add(rowInv.Cust_Group_Code)
                    End If
                    TBLDummy.DefaultView.RowFilter = TBLDummy.DefaultView.RowFilter + " AND " + "[SalesP Group Code] = '" + rowInv.Cust_Group_Code + "'"
                ElseIf TipeGroup = EnumTipeDetail.Detail_Cust_Segmentation Then
                    If ListSalesPersonGroup.Contains(rowInv.Segmentation_Code) = False Then
                        ListSalesPersonGroup.Add(rowInv.Segmentation_Code)
                    End If
                    TBLDummy.DefaultView.RowFilter = TBLDummy.DefaultView.RowFilter + " AND " + "[SalesP Group Code] = '" + rowInv.Segmentation_Code + "'"
                ElseIf TipeGroup = EnumTipeDetail.Detail_Cust_Type Then
                    If ListSalesPersonGroup.Contains(rowInv.Cust_Type_Code) = False Then
                        ListSalesPersonGroup.Add(rowInv.Cust_Type_Code)
                    End If
                    TBLDummy.DefaultView.RowFilter = TBLDummy.DefaultView.RowFilter + " AND " + "[SalesP Group Code] = '" + rowInv.Cust_Type_Code + "'"
                ElseIf TipeGroup = EnumTipeDetail.Detail_Product Then
                    If ListSalesPersonGroup.Contains(rowInv.Product_Code) = False Then
                        ListSalesPersonGroup.Add(rowInv.Product_Code)
                    End If
                    TBLDummy.DefaultView.RowFilter = TBLDummy.DefaultView.RowFilter + " AND " + "[SalesP Group Code] = '" + rowInv.Product_Code + "'"
                ElseIf TipeGroup = EnumTipeDetail.Detail_Product_Brand Then
                    If ListSalesPersonGroup.Contains(rowInv.Brand_Code) = False Then
                        ListSalesPersonGroup.Add(rowInv.Brand_Code)
                    End If
                    TBLDummy.DefaultView.RowFilter = TBLDummy.DefaultView.RowFilter + " AND " + "[SalesP Group Code] = '" + rowInv.Brand_Code + "'"
                ElseIf TipeGroup = EnumTipeDetail.Detail_Product_Category Then
                    If ListSalesPersonGroup.Contains(rowInv.Category_Code) = False Then
                        ListSalesPersonGroup.Add(rowInv.Category_Code)
                    End If
                    TBLDummy.DefaultView.RowFilter = TBLDummy.DefaultView.RowFilter + " AND " + "[SalesP Group Code] = '" + rowInv.Category_Code + "'"
                ElseIf TipeGroup = EnumTipeDetail.Detail_Product_Group Then
                    If ListSalesPersonGroup.Contains(rowInv.Product_Group_Code) = False Then
                        ListSalesPersonGroup.Add(rowInv.Product_Group_Code)
                    End If
                    TBLDummy.DefaultView.RowFilter = TBLDummy.DefaultView.RowFilter + " AND " + "[SalesP Group Code] = '" + rowInv.Product_Group_Code + "'"
                ElseIf TipeGroup = EnumTipeDetail.Detail_Product_Sub_Brand Then
                    If ListSalesPersonGroup.Contains(rowInv.Sub_Brand_Code) = False Then
                        ListSalesPersonGroup.Add(rowInv.Sub_Brand_Code)
                    End If
                    TBLDummy.DefaultView.RowFilter = TBLDummy.DefaultView.RowFilter + " AND " + "[SalesP Group Code] = '" + rowInv.Sub_Brand_Code + "'"
                ElseIf TipeGroup = EnumTipeDetail.Detail_Sales_Area Then
                    If ListSalesPersonGroup.Contains(rowInv.Sales_Area_Code) = False Then
                        ListSalesPersonGroup.Add(rowInv.Sales_Area_Code)
                    End If
                    TBLDummy.DefaultView.RowFilter = TBLDummy.DefaultView.RowFilter + " AND " + "[SalesP Group Code] = '" + rowInv.Sales_Area_Code + "'"
                ElseIf TipeGroup = EnumTipeDetail.Detail_SalesP Then
                    If ListSalesPersonGroup.Contains(rowInv.Sales_Person_Code) = False Then
                        ListSalesPersonGroup.Add(rowInv.Sales_Person_Code)
                    End If
                    TBLDummy.DefaultView.RowFilter = TBLDummy.DefaultView.RowFilter + " AND " + "[SalesP Group Code] = '" + rowInv.Sales_Person_Code + "'"
                ElseIf TipeGroup = EnumTipeDetail.Detail_SalesPGroup Then
                    If ListSalesPersonGroup.Contains(rowInv.Sales_Person_Group_Code) = False Then
                        ListSalesPersonGroup.Add(rowInv.Sales_Person_Group_Code)
                    End If
                    TBLDummy.DefaultView.RowFilter = TBLDummy.DefaultView.RowFilter + " AND " + "[SalesP Group Code] = '" + rowInv.Sales_Person_Group_Code + "'"
                Else
                    Exit Sub
                End If

                If TBLDummy.DefaultView.Count = 0 Then
                    'buat baris baru
                    rowCetak = TBLDummy.NewSales_Group_RecapRow
                    With rowCetak
                        If vAmount2 = False Then
                            If Me.ReferenceData = EnumTipeReference.Invoice Then
                                .Amount1 = rowInv.Sales_Amount
                            Else
                                .Amount1 = (rowInv.Sales_Amount - rowInv.Cost)
                            End If
                            .Amount2 = 0
                        Else
                            .Amount1 = 0
                            If Me.ReferenceData = EnumTipeReference.Invoice Then
                                .Amount2 = rowInv.Sales_Amount
                            Else
                                .Amount2 = rowInv.Cost
                            End If
                        End If
                        If TipeDetail = EnumTipeDetail.Detail_Cust_Group Then
                            .Sales_P_Code = rowInv.Cust_Group_Code
                            .Sales_P_Name = rowInv.Cust_Group_Desc
                        ElseIf TipeDetail = EnumTipeDetail.Detail_Cust_Segmentation Then
                            .Sales_P_Code = rowInv.Segmentation_Code
                            .Sales_P_Name = rowInv.Segmentation_Desc
                        ElseIf TipeDetail = EnumTipeDetail.Detail_Cust_Type Then
                            .Sales_P_Code = rowInv.Cust_Type_Code
                            .Sales_P_Name = rowInv.Cust_Type_Name
                        ElseIf TipeDetail = EnumTipeDetail.Detail_Product Then
                            .Sales_P_Code = rowInv.Product_Code
                            .Sales_P_Name = rowInv.Product_Desc
                        ElseIf TipeDetail = EnumTipeDetail.Detail_Product_Brand Then
                            .Sales_P_Code = rowInv.Brand_Code
                            .Sales_P_Name = rowInv.Brand_Desc
                        ElseIf TipeDetail = EnumTipeDetail.Detail_Product_Category Then
                            .Sales_P_Code = rowInv.Category_Code
                            .Sales_P_Name = rowInv.Category_Desc
                        ElseIf TipeDetail = EnumTipeDetail.Detail_Product_Group Then
                            .Sales_P_Code = rowInv.Product_Group_Code
                            .Sales_P_Name = rowInv.Product_Group_Desc
                        ElseIf TipeDetail = EnumTipeDetail.Detail_Product_Sub_Brand Then
                            .Sales_P_Code = rowInv.Sub_Brand_Code
                            .Sales_P_Name = rowInv.Sub_Brand_Desc
                        ElseIf TipeDetail = EnumTipeDetail.Detail_Sales_Area Then
                            .Sales_P_Code = rowInv.Sales_Area_Code
                            .Sales_P_Name = rowInv.Sales_Area_Desc
                        ElseIf TipeDetail = EnumTipeDetail.Detail_SalesP Then
                            .Sales_P_Code = rowInv.Sales_Person_Code
                            .Sales_P_Name = rowInv.Sales_Person_Desc
                        ElseIf TipeDetail = EnumTipeDetail.Detail_SalesPGroup Then
                            .Sales_P_Code = rowInv.Sales_Person_Group_Code
                            .Sales_P_Name = rowInv.Sales_Person_Group_Desc
                        Else
                            Exit Sub
                        End If
                        If TipeGroup = EnumTipeDetail.Detail_Cust_Group Then
                            .SalesP_Group_Code = rowInv.Cust_Group_Code
                            .SalesP_Group_Name = rowInv.Cust_Group_Desc
                        ElseIf TipeGroup = EnumTipeDetail.Detail_Cust_Segmentation Then
                            .SalesP_Group_Code = rowInv.Segmentation_Code
                            .SalesP_Group_Name = rowInv.Segmentation_Desc
                        ElseIf TipeGroup = EnumTipeDetail.Detail_Cust_Type Then
                            .SalesP_Group_Code = rowInv.Cust_Type_Code
                            .SalesP_Group_Name = rowInv.Cust_Type_Name
                        ElseIf TipeGroup = EnumTipeDetail.Detail_Product Then
                            .SalesP_Group_Code = rowInv.Product_Code
                            .SalesP_Group_Name = rowInv.Product_Desc
                        ElseIf TipeGroup = EnumTipeDetail.Detail_Product_Brand Then
                            .SalesP_Group_Code = rowInv.Brand_Code
                            .SalesP_Group_Name = rowInv.Brand_Desc
                        ElseIf TipeGroup = EnumTipeDetail.Detail_Product_Category Then
                            .SalesP_Group_Code = rowInv.Category_Code
                            .SalesP_Group_Name = rowInv.Category_Desc
                        ElseIf TipeGroup = EnumTipeDetail.Detail_Product_Group Then
                            .SalesP_Group_Code = rowInv.Product_Group_Code
                            .SalesP_Group_Name = rowInv.Product_Group_Desc
                        ElseIf TipeGroup = EnumTipeDetail.Detail_Product_Sub_Brand Then
                            .SalesP_Group_Code = rowInv.Sub_Brand_Code
                            .SalesP_Group_Name = rowInv.Sub_Brand_Desc
                        ElseIf TipeGroup = EnumTipeDetail.Detail_Sales_Area Then
                            .SalesP_Group_Code = rowInv.Sales_Area_Code
                            .SalesP_Group_Name = rowInv.Sales_Area_Desc
                        ElseIf TipeGroup = EnumTipeDetail.Detail_SalesP Then
                            .SalesP_Group_Code = rowInv.Sales_Person_Code
                            .SalesP_Group_Name = rowInv.Sales_Person_Desc
                        ElseIf TipeGroup = EnumTipeDetail.Detail_SalesPGroup Then
                            .SalesP_Group_Code = rowInv.Sales_Person_Group_Code
                            .SalesP_Group_Name = rowInv.Sales_Person_Group_Desc
                        Else
                            Exit Sub
                        End If
                        .Variance = 0
                        .Sub_Total_1 = 0
                        .Sub_Total_2 = 0
                    End With
                    TBLDummy.Rows.Add(rowCetak)
                Else
                    'tambahkan amount saja
                    rowCetak = TBLDummy.DefaultView(0).Row
                    If vAmount2 = False Then
                        If Me.ReferenceData = EnumTipeReference.Invoice Then
                            rowCetak.Amount1 += rowInv.Sales_Amount
                        Else
                            rowCetak.Amount1 += rowInv.Sales_Amount - rowInv.Cost
                        End If
                    Else
                        If Me.ReferenceData = EnumTipeReference.Invoice Then
                            rowCetak.Amount2 += rowInv.Sales_Amount
                        Else
                            rowCetak.Amount2 += (rowInv.Sales_Amount - rowInv.Cost)
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BuatLaporanInvoice_SalesPersonDetail_2Period()
        Dim Filter As String = ""
        Dim DTAInv As New DseLaporanSalesPerformanceTableAdapters.tbltrans_salesinvoice_item1TableAdapter
        Dim TBLInv As New DseLaporanSalesPerformance.tbltrans_salesinvoice_item1DataTable
        Dim DV As DataRowView
        Dim TBLCetak As DseLaporanSalesPerformance.Sales_Group_RecapDataTable = Me.DS.Sales_Group_Recap
        Dim rowCetak, rowDummy As DseLaporanSalesPerformance.Sales_Group_RecapRow
        Dim DefPrec As Integer = 0
        Dim TBLDummy As New DseLaporanSalesPerformance.Sales_Group_RecapDataTable
        Dim ListSalesPersonGroup As New List(Of String)
        Dim SalesPersonG As String = ""
        Dim SubTotal As Decimal = 0
        Dim DTABudget As New DseLaporanSalesPerformanceTableAdapters.tblbudgetTableAdapter
        Dim TBLBudget As New DseLaporanSalesPerformance.tblbudgetDataTable
        Dim TipeBudget As ClsTransaksiSalesBudget.EnumTipeBudget = ClsTransaksiSalesBudget.EnumTipeBudget.Invoice

        Try
            TBLCetak.Rows.Clear()
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
            DTAInv.Connection.ConnectionString = Core_Module.GetConstring
            DTABudget.Connection.ConnectionString = DTAInv.Connection.ConnectionString

            DTAInv.Fill(TBLInv, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value), Me.dtpDari.Value)
            If Me.chkReturn.Checked = True Then
                Me.KurangkanItemReturn(TBLInv, Me.dtpDari.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value))
            End If

            'tambahkan data periode ke-2
            If Me.TipeGroup = EnumTipeDetail.Detail_SalesP And Me.cboFormat1.SelectedIndex = 1 Then
                'isikan amount 2
                Me.IsikanTabelInvoice_Amount1(TBLInv, TBLDummy, DefPrec, ListSalesPersonGroup, True)

                'budget
                If Me.ReferenceData = EnumTipeReference.Invoice Then
                    TipeBudget = ClsTransaksiSalesBudget.EnumTipeBudget.Invoice
                Else
                    TipeBudget = ClsTransaksiSalesBudget.EnumTipeBudget.Invoice_GrossMargin
                End If

                DTABudget.Fill(TBLBudget, TipeBudget, Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpSampai.Value), Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpDari.Value))
                For Each rowDummy In TBLDummy.Rows
                    rowDummy.Amount1 = Me.GetBudget(rowDummy.SalesP_Group_Code, Me.dtpDari.Value, _
                    Me.ObjCD.GetFirstDateofTheMonth(Me.dtpSampai.Value), TBLBudget)
                Next
            ElseIf Me.TipeGroup = EnumTipeDetail.Detail_SalesPGroup And Me.cboFormat1.SelectedIndex = 1 Then
                'isikan amount 2
                Me.IsikanTabelInvoice_Amount1(TBLInv, TBLDummy, DefPrec, ListSalesPersonGroup, True)

                'budget
                If Me.ReferenceData = EnumTipeReference.Invoice Then
                    TipeBudget = ClsTransaksiSalesBudget.EnumTipeBudget.Invoice
                Else
                    TipeBudget = ClsTransaksiSalesBudget.EnumTipeBudget.Invoice_GrossMargin
                End If

                DTABudget.Fill(TBLBudget, TipeBudget, Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpSampai.Value), Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpDari.Value))
                For Each rowDummy In TBLDummy.Rows
                    rowDummy.Amount1 = Me.GetBudget_PerGroup(rowDummy.SalesP_Group_Code, Me.dtpDari.Value, _
                    Me.ObjCD.GetFirstDateofTheMonth(Me.dtpSampai.Value), TBLBudget)
                Next
            Else
                'isikan amount 1
                Me.IsikanTabelInvoice_Amount1(TBLInv, TBLDummy, DefPrec, ListSalesPersonGroup, False)
                'isikan amount 2
                TBLInv.Rows.Clear()
                DTAInv.Fill(TBLInv, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai2.Value), Me.dtpDari2.Value)
                If Me.chkReturn.Checked = True Then
                    Me.KurangkanItemReturn(TBLInv, Me.dtpDari2.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai2.Value))
                End If
                Me.IsikanTabelInvoice_Amount1(TBLInv, TBLDummy, DefPrec, ListSalesPersonGroup, True)
            End If

            'hitung variance
            For Each rowDummy In TBLDummy.Rows
                If rowDummy.Amount1 = 0 Then
                    rowDummy.Variance = 0
                Else
                                            rowDummy.Variance = Math.Round(((rowDummy.Amount2 - rowDummy.Amount1) * 100) / rowDummy.Amount1, 2, MidpointRounding.AwayFromZero)
                End If
            Next

            'sorting ulang berdasarkan penjualan paling tinggi, per group per sales person
            '1.hitung sub total per group
            For Each SalesPersonG In ListSalesPersonGroup
                TBLDummy.DefaultView.RowFilter = "[SalesP Group Code] = '" + SalesPersonG + "'"
                SubTotal = 0
                For Each DV In TBLDummy.DefaultView
                    rowDummy = DV.Row
                    SubTotal += rowDummy.Amount1
                Next
                For Each DV In TBLDummy.DefaultView
                    rowDummy = DV.Row
                    rowDummy.Sub_Total_1 = SubTotal
                Next
            Next
            For Each SalesPersonG In ListSalesPersonGroup
                TBLDummy.DefaultView.RowFilter = "[SalesP Group Code] = '" + SalesPersonG + "'"
                SubTotal = 0
                For Each DV In TBLDummy.DefaultView
                    rowDummy = DV.Row
                    SubTotal += rowDummy.Amount2
                Next
                For Each DV In TBLDummy.DefaultView
                    rowDummy = DV.Row
                    rowDummy.Sub_Total_2 = SubTotal
                Next
            Next
            '2. sorting
            TBLDummy.DefaultView.RowFilter = ""
            TBLDummy.DefaultView.Sort = "[Sub Total 1] DESC, [SalesP Group Code] ASC, Amount1 DESC, [Sales P Code] ASC"
            For Each DV In TBLDummy.DefaultView
                rowDummy = DV.Row
                rowCetak = TBLCetak.NewSales_Group_RecapRow
                For ij As Integer = 0 To TBLCetak.Columns.Count - 1
                    rowCetak.Item(ij) = rowDummy.Item(ij)
                Next
                TBLCetak.Rows.Add(rowCetak)
            Next

            If (Me.TipeGroup = EnumTipeDetail.Detail_SalesP And Me.cboFormat1.SelectedIndex = 1) Or _
            (Me.TipeGroup = EnumTipeDetail.Detail_SalesPGroup And Me.cboFormat1.SelectedIndex = 1) Then
                Me.rpt_Recap5.SetDataSource(DS)
                rpt_Recap5.SetParameterValue(rpt_Recap5.Parameter_Judul1.ParameterFieldName, "Sales Person Performance Report")
                If Me.ReferenceData = EnumTipeReference.Invoice Then
                    rpt_Recap5.SetParameterValue(rpt_Recap5.Parameter_Judul2.ParameterFieldName, "Sales Person Detail From Sales Invoice -  With Budget")
                Else
                    rpt_Recap5.SetParameterValue(rpt_Recap5.Parameter_Judul2.ParameterFieldName, "Sales Person Detail From Gross Margin -  With Budget")
                End If
                rpt_Recap5.SetParameterValue(rpt_Recap5.Parameter_Amount1.ParameterFieldName, "Budget")
                rpt_Recap5.SetParameterValue(rpt_Recap5.Parameter_Amount2.ParameterFieldName, "Amount")
                rpt_Recap5.SetParameterValue(rpt_Recap5.Parameter_Date1.ParameterFieldName, Me.dtpDari.Value)
                rpt_Recap5.SetParameterValue(rpt_Recap5.Parameter_Date2.ParameterFieldName, Me.dtpSampai.Value)
                rpt_Recap5.SetParameterValue(rpt_Recap5.Parameter_Date3.ParameterFieldName, Me.dtpDari2.Value)
                rpt_Recap5.SetParameterValue(rpt_Recap5.Parameter_Date4.ParameterFieldName, Me.dtpSampai2.Value)
                rpt_Recap5.SetParameterValue(rpt_Recap5.Parameter_DefPrec.ParameterFieldName, DefPrec)

                Me.crv1.ReportSource = rpt_Recap5
                Me.crv1.Refresh()
                Me.crv1.Show()
            Else
                Me.rpt_Recap4.SetDataSource(DS)
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Judul1.ParameterFieldName, "Sales Person Performance Report")
                If Me.ReferenceData = EnumTipeReference.Invoice Then
                    rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Judul2.ParameterFieldName, "Sales Person Detail From Sales Invoice -  2 Period")
                Else
                    rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Judul2.ParameterFieldName, "Sales Person Detail From Gross Margin -  2 Period")
                End If
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Amount1.ParameterFieldName, "Period I")
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Amount2.ParameterFieldName, "Period II")
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Date1.ParameterFieldName, Me.dtpDari.Value)
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Date2.ParameterFieldName, Me.dtpSampai.Value)
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Date3.ParameterFieldName, Me.dtpDari2.Value)
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_Date4.ParameterFieldName, Me.dtpSampai2.Value)
                rpt_Recap4.SetParameterValue(rpt_Recap4.Parameter_DefPrec.ParameterFieldName, DefPrec)

                Me.crv1.ReportSource = rpt_Recap4
                Me.crv1.Refresh()
                Me.crv1.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAInv.Dispose()
            TBLInv.Dispose()
            TBLDummy.Dispose()
            ListSalesPersonGroup = Nothing
            DTABudget.Dispose()
            TBLBudget.Dispose()
        End Try
    End Sub

    Private Sub BuatLaporanInvoice_TrendAmount(ByVal vUsePercent As Boolean, ByVal vSalesGroup As Boolean)
        Dim Filter As String = ""
        Dim DTAInv As New DseLaporanSalesPerformanceTableAdapters.tbltrans_salesinvoice_item1TableAdapter
        Dim TBLInv As New DseLaporanSalesPerformance.tbltrans_salesinvoice_item1DataTable
        Dim rowInv As DseLaporanSalesPerformance.tbltrans_salesinvoice_item1Row
        Dim DV As DataRowView
        Dim TBLCetak As DseLaporanSalesPerformance.Trend_AmountDataTable = DS.Trend_Amount
        Dim rowCetak As DseLaporanSalesPerformance.Trend_AmountRow
        Dim DefPrec As Integer = 0
        Dim TBLRefMonth As New DseLaporanSalesPerformance.Ref_PeriodDataTable
        Dim rowRefMonth As DseLaporanSalesPerformance.Ref_PeriodRow
        Dim StartRefPeriod As Date
        Dim Dummy As Integer = 0
        Dim DummyStr As String = ""

        Try
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
            DTAInv.Connection.ConnectionString = Core_Module.GetConstring
            TBLCetak.Rows.Clear()
            DTAInv.Fill(TBLInv, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value), Me.dtpDari.Value)
            If Me.chkReturn.Checked = True Then
                Me.KurangkanItemReturn(TBLInv, Me.dtpDari.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value))
            End If
            Filter = Me.BuatFilterSO
            TBLInv.DefaultView.RowFilter = Filter

            'isikan ref period
            StartRefPeriod = Me.dtpDari.Value.Date
            TBLRefMonth.Rows.Clear()
            While StartRefPeriod <= Me.dtpSampai.Value.Date
                Dummy += 1
                rowRefMonth = TBLRefMonth.NewRef_PeriodRow
                With rowRefMonth
                    .Ref_Date = StartRefPeriod
                    DummyStr = Dummy.ToString
                    If DummyStr.Length = 1 Then
                        DummyStr = "0" + DummyStr
                    End If
                    .Month = DummyStr + ". " + Format(.Ref_Date, "MMM yyyy")
                End With
                TBLRefMonth.Rows.Add(rowRefMonth)

                StartRefPeriod = Me.ObjCD.GetFirstDateofTheMonth(StartRefPeriod + TimeSpan.FromDays(32))
            End While

            'isikan trend tabel
            For Each DV In TBLInv.DefaultView
                rowInv = DV.Row

                TBLRefMonth.DefaultView.RowFilter = "[Ref Date] = #" + Format(Me.ObjCD.GetFirstDateofTheMonth(rowInv.Tanggal_Transaksi), "MM/dd/yyyy") + "#"
                rowRefMonth = TBLRefMonth.DefaultView(0).Row
                If vSalesGroup = False Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + rowInv.Sales_Person_Desc + "' AND Month = '" + rowRefMonth.Month + "'"
                Else
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + rowInv.Sales_Person_Group_Desc + "' AND Month = '" + rowRefMonth.Month + "'"
                End If
                If TBLCetak.DefaultView.Count = 0 Then
                    'tambahkan baris baru
                    rowCetak = TBLCetak.NewTrend_AmountRow
                    If Me.ReferenceData = EnumTipeReference.Invoice Then
                        rowCetak.Amount = rowInv.Sales_Amount
                    Else
                        rowCetak.Amount = (rowInv.Sales_Amount - rowInv.Cost)
                    End If
                    rowCetak.Month = rowRefMonth.Month
                    If vSalesGroup = False Then
                        rowCetak.Sales_Person_Detail = rowInv.Sales_Person_Desc
                    Else
                        rowCetak.Sales_Person_Detail = rowInv.Sales_Person_Group_Desc
                    End If

                    TBLCetak.Rows.Add(rowCetak)
                Else
                    'masukkan amount saja
                    rowCetak = TBLCetak.DefaultView(0).Row
                    If Me.ReferenceData = EnumTipeReference.Invoice Then
                        rowCetak.Amount += rowInv.Sales_Amount
                    Else
                        rowCetak.Amount += (rowInv.Sales_Amount - rowInv.Cost)
                    End If
                End If
            Next

            If vUsePercent = False Then
                Me.rpt_Trend.SetDataSource(DS)
                rpt_Trend.SetParameterValue(rpt_Trend.Parameter_Date1.ParameterFieldName, Me.dtpDari.Value)
                rpt_Trend.SetParameterValue(rpt_Trend.Parameter_Date2.ParameterFieldName, Me.dtpSampai.Value)
                rpt_Trend.SetParameterValue(rpt_Trend.Parameter_DefPrec.ParameterFieldName, DefPrec)

                Me.crv1.ReportSource = rpt_Trend
            Else
                Me.rpt_Trend2.SetDataSource(DS)
                rpt_Trend2.SetParameterValue(rpt_Trend2.Parameter_Date1.ParameterFieldName, Me.dtpDari.Value)
                rpt_Trend2.SetParameterValue(rpt_Trend2.Parameter_Date2.ParameterFieldName, Me.dtpSampai.Value)
                rpt_Trend2.SetParameterValue(rpt_Trend2.Parameter_DefPrec.ParameterFieldName, DefPrec)

                Me.crv1.ReportSource = rpt_Trend2
            End If
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAInv.Dispose()
            TBLInv.Dispose()
            TBLRefMonth.Dispose()
        End Try
    End Sub

    Private Sub BuatLaporanInvoice_SalesComposition()
        Try
            If Me.TipeGroup = EnumTipeDetail.NoDetail Then
                Me.BuatLaporanInvoice_SalesComposition_NoGrouping()
            Else
                Me.BuatLaporanInvoice_SalesComposition_Grouping()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BuatLaporanInvoice_SalesComposition_Grouping()
        Dim Filter As String = ""
        Dim DTAInv As New DseLaporanSalesPerformanceTableAdapters.tbltrans_salesinvoice_item1TableAdapter
        Dim TBLInv As New DseLaporanSalesPerformance.tbltrans_salesinvoice_item1DataTable
        Dim rowInv As DseLaporanSalesPerformance.tbltrans_salesinvoice_item1Row
        Dim DV As DataRowView
        Dim TBLCetak As DseLaporanSalesPerformance.Sales_CompositionDataTable = DS.Sales_Composition
        Dim rowCetak As DseLaporanSalesPerformance.Sales_CompositionRow
        Dim DefPrec As Integer = 0
        Dim Total As Decimal = 0
        Dim ListGroup As New List(Of String)
        Dim Group As String = ""
        Dim TotalPerc As Decimal = 0
        Dim TBLDummy As New DseLaporanSalesPerformance.Sales_CompositionDataTable
        Dim rowDummy As DseLaporanSalesPerformance.Sales_CompositionRow

        Try
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
            DTAInv.Connection.ConnectionString = Core_Module.GetConstring
            TBLCetak.Rows.Clear()
            DTAInv.Fill(TBLInv, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value), Me.dtpDari.Value)
            If Me.chkReturn.Checked = True Then
                Me.KurangkanItemReturn(TBLInv, Me.dtpDari.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value))
            End If
            Filter = Me.BuatFilterSO
            TBLInv.DefaultView.RowFilter = Filter

            'isikan trend tabel
            For Each DV In TBLInv.DefaultView
                rowInv = DV.Row

                If Me.TipeDetail = EnumTipeDetail.Detail_Cust_Group Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail 2] = '" + rowInv.Cust_Group_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Cust_Segmentation Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail 2] = '" + rowInv.Segmentation_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Cust_Type Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail 2] = '" + rowInv.Cust_Type_Name + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail 2] = '" + rowInv.Product_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Brand Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail 2] = '" + rowInv.Brand_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Category Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail 2] = '" + rowInv.Category_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Group Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail 2] = '" + rowInv.Product_Group_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Sub_Brand Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail 2] = '" + rowInv.Sub_Brand_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Sales_Area Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail 2] = '" + rowInv.Sales_Area_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_SalesP Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail 2] = '" + rowInv.Sales_Person_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_SalesPGroup Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail 2] = '" + rowInv.Sales_Person_Group_Desc + "'"
                End If
                If Me.TipeGroup = EnumTipeDetail.Detail_Cust_Group Then
                    TBLCetak.DefaultView.RowFilter = TBLCetak.DefaultView.RowFilter + " AND [Sales Person Detail] = '" + rowInv.Cust_Group_Desc + "'"
                ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Cust_Segmentation Then
                    TBLCetak.DefaultView.RowFilter = TBLCetak.DefaultView.RowFilter + " AND [Sales Person Detail] = '" + rowInv.Segmentation_Desc + "'"
                ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Cust_Type Then
                    TBLCetak.DefaultView.RowFilter = TBLCetak.DefaultView.RowFilter + " AND [Sales Person Detail] = '" + rowInv.Cust_Type_Name + "'"
                ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Product Then
                    TBLCetak.DefaultView.RowFilter = TBLCetak.DefaultView.RowFilter + " AND [Sales Person Detail] = '" + rowInv.Product_Desc + "'"
                ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Product_Brand Then
                    TBLCetak.DefaultView.RowFilter = TBLCetak.DefaultView.RowFilter + " AND [Sales Person Detail] = '" + rowInv.Brand_Desc + "'"
                ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Product_Category Then
                    TBLCetak.DefaultView.RowFilter = TBLCetak.DefaultView.RowFilter + " AND [Sales Person Detail] = '" + rowInv.Category_Desc + "'"
                ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Product_Group Then
                    TBLCetak.DefaultView.RowFilter = TBLCetak.DefaultView.RowFilter + " AND [Sales Person Detail] = '" + rowInv.Product_Group_Desc + "'"
                ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Product_Sub_Brand Then
                    TBLCetak.DefaultView.RowFilter = TBLCetak.DefaultView.RowFilter + " AND [Sales Person Detail] = '" + rowInv.Sub_Brand_Desc + "'"
                ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Sales_Area Then
                    TBLCetak.DefaultView.RowFilter = TBLCetak.DefaultView.RowFilter + " AND [Sales Person Detail] = '" + rowInv.Sales_Area_Desc + "'"
                ElseIf Me.TipeGroup = EnumTipeDetail.Detail_SalesP Then
                    TBLCetak.DefaultView.RowFilter = TBLCetak.DefaultView.RowFilter + " AND [Sales Person Detail] = '" + rowInv.Sales_Person_Desc + "'"
                ElseIf Me.TipeGroup = EnumTipeDetail.Detail_SalesPGroup Then
                    TBLCetak.DefaultView.RowFilter = TBLCetak.DefaultView.RowFilter + " AND [Sales Person Detail] = '" + rowInv.Sales_Person_Group_Desc + "'"
                End If
                If TBLCetak.DefaultView.Count = 0 Then
                    'tambahkan baris baru
                    rowCetak = TBLCetak.NewSales_CompositionRow
                    If Me.ReferenceData = EnumTipeReference.Invoice Then
                        rowCetak.Amount = rowInv.Sales_Amount
                    Else
                        rowCetak.Amount = rowInv.Sales_Amount - rowInv.Cost
                    End If
                    rowCetak.Percentage = 0
                    rowCetak.Sub_Total_Amount = 0
                    rowCetak.Sub_Total_Percetage = 0
                    rowCetak.Sales_Person_Detail_2 = ""
                    If Me.TipeDetail = EnumTipeDetail.Detail_Cust_Group Then
                        rowCetak.Sales_Person_Detail_2 = rowInv.Cust_Group_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Cust_Segmentation Then
                        rowCetak.Sales_Person_Detail_2 = rowInv.Segmentation_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Cust_Type Then
                        rowCetak.Sales_Person_Detail_2 = rowInv.Cust_Type_Name
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product Then
                        rowCetak.Sales_Person_Detail_2 = rowInv.Product_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Brand Then
                        rowCetak.Sales_Person_Detail_2 = rowInv.Brand_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Category Then
                        rowCetak.Sales_Person_Detail_2 = rowInv.Category_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Group Then
                        rowCetak.Sales_Person_Detail_2 = rowInv.Product_Group_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Sub_Brand Then
                        rowCetak.Sales_Person_Detail_2 = rowInv.Sub_Brand_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Sales_Area Then
                        rowCetak.Sales_Person_Detail_2 = rowInv.Sales_Area_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_SalesP Then
                        rowCetak.Sales_Person_Detail_2 = rowInv.Sales_Person_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_SalesPGroup Then
                        rowCetak.Sales_Person_Detail_2 = rowInv.Sales_Person_Group_Desc
                    End If
                    If Me.TipeGroup = EnumTipeDetail.Detail_Cust_Group Then
                        rowCetak.Sales_Person_Detail = rowInv.Cust_Group_Desc
                    ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Cust_Segmentation Then
                        rowCetak.Sales_Person_Detail = rowInv.Segmentation_Desc
                    ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Cust_Type Then
                        rowCetak.Sales_Person_Detail = rowInv.Cust_Type_Name
                    ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Product Then
                        rowCetak.Sales_Person_Detail = rowInv.Product_Desc
                    ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Product_Brand Then
                        rowCetak.Sales_Person_Detail = rowInv.Brand_Desc
                    ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Product_Category Then
                        rowCetak.Sales_Person_Detail = rowInv.Category_Desc
                    ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Product_Group Then
                        rowCetak.Sales_Person_Detail = rowInv.Product_Group_Desc
                    ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Product_Sub_Brand Then
                        rowCetak.Sales_Person_Detail = rowInv.Sub_Brand_Desc
                    ElseIf Me.TipeGroup = EnumTipeDetail.Detail_Sales_Area Then
                        rowCetak.Sales_Person_Detail = rowInv.Sales_Area_Desc
                    ElseIf Me.TipeGroup = EnumTipeDetail.Detail_SalesP Then
                        rowCetak.Sales_Person_Detail = rowInv.Sales_Person_Desc
                    ElseIf Me.TipeGroup = EnumTipeDetail.Detail_SalesPGroup Then
                        rowCetak.Sales_Person_Detail = rowInv.Sales_Person_Group_Desc
                    End If

                    TBLCetak.Rows.Add(rowCetak)
                Else
                    'masukkan amount saja
                    rowCetak = TBLCetak.DefaultView(0).Row
                    If Me.ReferenceData = EnumTipeReference.Invoice Then
                        rowCetak.Amount += rowInv.Sales_Amount
                    Else
                        rowCetak.Amount += (rowInv.Sales_Amount - rowInv.Cost)
                    End If
                End If
            Next

            'hitung persetase
            ListGroup.Clear()
            For Each rowCetak In TBLCetak.Rows
                Total += rowCetak.Amount
                If ListGroup.Contains(rowCetak.Sales_Person_Detail) = False Then
                    ListGroup.Add(rowCetak.Sales_Person_Detail)
                End If
            Next
            For Each rowCetak In TBLCetak.Rows
                rowCetak.Percentage = Math.Round(rowCetak.Amount * 100 / Total, 2, MidpointRounding.AwayFromZero)
            Next
            'sub total persentase dan amount
            For Each Group In ListGroup
                TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + Group + "'"
                Total = 0
                TotalPerc = 0
                For Each DV In TBLCetak.DefaultView
                    rowCetak = DV.Row
                    Total += rowCetak.Amount
                    TotalPerc += rowCetak.Percentage
                Next
                For Each DV In TBLCetak.DefaultView
                    rowCetak = DV.Row
                    rowCetak.Sub_Total_Amount = Total
                    rowCetak.Sub_Total_Percetage = TotalPerc
                Next
            Next
            'urutkan
            For Each rowCetak In TBLCetak.Rows
                rowDummy = TBLDummy.NewSales_CompositionRow
                For ij As Integer = 0 To TBLDummy.Columns.Count - 1
                    rowDummy.Item(ij) = rowCetak.Item(ij)
                Next
                TBLDummy.Rows.Add(rowDummy)
            Next
            TBLDummy.DefaultView.RowFilter = ""
            TBLDummy.DefaultView.Sort = "[Sub Total Amount] DESC, [Sales Person Detail] ASC, Amount DESC, [Sales Person Detail 2] ASC"
            TBLCetak.Rows.Clear()
            For Each DV In TBLDummy.DefaultView
                rowDummy = DV.Row
                rowCetak = TBLCetak.NewSales_CompositionRow
                For ij As Integer = 0 To TBLDummy.Columns.Count - 1
                    rowCetak.Item(ij) = rowDummy.Item(ij)
                Next
                TBLCetak.Rows.Add(rowCetak)
            Next

            'tampilkan laporan 
            Me.rpt_Composition2.SetDataSource(DS)
            rpt_Composition2.SetParameterValue(rpt_Composition2.Parameter_Date1.ParameterFieldName, Me.dtpDari.Value)
            rpt_Composition2.SetParameterValue(rpt_Composition2.Parameter_Date2.ParameterFieldName, Me.dtpSampai.Value)
            rpt_Composition2.SetParameterValue(rpt_Composition2.Parameter_CrpLaporanSalesPerformance_CompositionDetail_DefPrec.ParameterFieldName, DefPrec)
            Me.crv1.ReportSource = rpt_Composition2
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAInv.Dispose()
            TBLInv.Dispose()
            ListGroup = Nothing
            TBLDummy = Nothing
        End Try
    End Sub

    Private Sub BuatLaporanInvoice_SalesComposition_NoGrouping()
        Dim Filter As String = ""
        Dim DTAInv As New DseLaporanSalesPerformanceTableAdapters.tbltrans_salesinvoice_item1TableAdapter
        Dim TBLInv As New DseLaporanSalesPerformance.tbltrans_salesinvoice_item1DataTable
        Dim rowInv As DseLaporanSalesPerformance.tbltrans_salesinvoice_item1Row
        Dim DV As DataRowView
        Dim TBLCetak As DseLaporanSalesPerformance.Sales_CompositionDataTable = DS.Sales_Composition
        Dim rowCetak As DseLaporanSalesPerformance.Sales_CompositionRow
        Dim DefPrec As Integer = 0
        Dim Total As Decimal = 0
        Dim TBLDummy As New DseLaporanSalesPerformance.Sales_CompositionDataTable
        Dim rowDummy As DseLaporanSalesPerformance.Sales_CompositionRow

        Try
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
            DTAInv.Connection.ConnectionString = Core_Module.GetConstring
            TBLCetak.Rows.Clear()
            DTAInv.Fill(TBLInv, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value), Me.dtpDari.Value)
            If Me.chkReturn.Checked = True Then
                Me.KurangkanItemReturn(TBLInv, Me.dtpDari.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSampai.Value))
            End If
            Filter = Me.BuatFilterSO
            TBLInv.DefaultView.RowFilter = Filter

            'isikan trend tabel
            For Each DV In TBLInv.DefaultView
                rowInv = DV.Row


                If Me.TipeDetail = EnumTipeDetail.Detail_Cust_Group Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + rowInv.Cust_Group_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Cust_Segmentation Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + rowInv.Segmentation_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Cust_Type Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + rowInv.Cust_Type_Name + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + rowInv.Product_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Brand Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + rowInv.Brand_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Category Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + rowInv.Category_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Group Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + rowInv.Product_Group_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Sub_Brand Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + rowInv.Sub_Brand_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Sales_Area Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + rowInv.Sales_Area_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_SalesP Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + rowInv.Sales_Person_Desc + "'"
                ElseIf Me.TipeDetail = EnumTipeDetail.Detail_SalesPGroup Then
                    TBLCetak.DefaultView.RowFilter = "[Sales Person Detail] = '" + rowInv.Sales_Person_Group_Desc + "'"
                End If
                If TBLCetak.DefaultView.Count = 0 Then
                    'tambahkan baris baru
                    rowCetak = TBLCetak.NewSales_CompositionRow
                    If Me.ReferenceData = EnumTipeReference.Invoice Then
                        rowCetak.Amount = rowInv.Sales_Amount
                    Else
                        rowCetak.Amount = rowInv.Sales_Amount - rowInv.Cost
                    End If
                    rowCetak.Percentage = 0
                    rowCetak.Sub_Total_Amount = 0
                    rowCetak.Sub_Total_Percetage = 0
                    rowCetak.Sales_Person_Detail_2 = ""
                    If Me.TipeDetail = EnumTipeDetail.Detail_Cust_Group Then
                        rowCetak.Sales_Person_Detail = rowInv.Cust_Group_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Cust_Segmentation Then
                        rowCetak.Sales_Person_Detail = rowInv.Segmentation_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Cust_Type Then
                        rowCetak.Sales_Person_Detail = rowInv.Cust_Type_Name
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product Then
                        rowCetak.Sales_Person_Detail = rowInv.Product_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Brand Then
                        rowCetak.Sales_Person_Detail = rowInv.Brand_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Category Then
                        rowCetak.Sales_Person_Detail = rowInv.Category_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Group Then
                        rowCetak.Sales_Person_Detail = rowInv.Product_Group_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Product_Sub_Brand Then
                        rowCetak.Sales_Person_Detail = rowInv.Sub_Brand_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_Sales_Area Then
                        rowCetak.Sales_Person_Detail = rowInv.Sales_Area_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_SalesP Then
                        rowCetak.Sales_Person_Detail = rowInv.Sales_Person_Desc
                    ElseIf Me.TipeDetail = EnumTipeDetail.Detail_SalesPGroup Then
                        rowCetak.Sales_Person_Detail = rowInv.Sales_Person_Group_Desc
                    End If

                    TBLCetak.Rows.Add(rowCetak)
                Else
                    'masukkan amount saja
                    rowCetak = TBLCetak.DefaultView(0).Row
                    If Me.ReferenceData = EnumTipeReference.Invoice Then
                        rowCetak.Amount += rowInv.Sales_Amount
                    Else
                        rowCetak.Amount += (rowInv.Sales_Amount - rowInv.Cost)
                    End If
                End If
            Next
            'hitung persetase
            For Each rowCetak In TBLCetak.Rows
                Total += rowCetak.Amount
            Next
            For Each rowCetak In TBLCetak.Rows
                rowCetak.Percentage = Math.Round(rowCetak.Amount * 100 / Total, 2, MidpointRounding.AwayFromZero)
            Next
            'urutkan
            For Each rowCetak In TBLCetak.Rows
                rowDummy = TBLDummy.NewSales_CompositionRow
                For ij As Integer = 0 To TBLDummy.Columns.Count - 1
                    rowDummy.Item(ij) = rowCetak.Item(ij)
                Next
                TBLDummy.Rows.Add(rowDummy)
            Next
            TBLDummy.DefaultView.RowFilter = ""
            TBLDummy.DefaultView.Sort = "Amount DESC, [Sales Person Detail] ASC"
            TBLCetak.Rows.Clear()
            For Each DV In TBLDummy.DefaultView
                rowDummy = DV.Row
                rowCetak = TBLCetak.NewSales_CompositionRow()
                For ij As Integer = 0 To TBLCetak.Columns.Count - 1
                    rowCetak.Item(ij) = rowDummy.Item(ij)
                Next
                TBLCetak.Rows.Add(rowCetak)
            Next

            'tampilkan laporan 
            Me.rpt_Composition.SetDataSource(DS)
            rpt_Composition.SetParameterValue(rpt_Composition.Parameter_Date1.ParameterFieldName, Me.dtpDari.Value)
            rpt_Composition.SetParameterValue(rpt_Composition.Parameter_Date2.ParameterFieldName, Me.dtpSampai.Value)
            rpt_Composition.SetParameterValue(rpt_Composition.Parameter_CrpLaporanSalesPerformance_CompositionDetail_DefPrec.ParameterFieldName, DefPrec)
            Me.crv1.ReportSource = rpt_Composition
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAInv.Dispose()
            TBLInv.Dispose()
            TBLDummy.Dispose()
        End Try
    End Sub

    Private Sub KurangkanItemReturn(ByVal vTBLInv As DseLaporanSalesPerformance.tbltrans_salesinvoice_item1DataTable, ByVal vStartPeriod As Date, ByVal vEndPeriod As Date)
        Dim DTASR_Raw As New DseLaporanSalesPerformanceTableAdapters.tbltrans_sr_itemTableAdapter
        Dim TBLSR_Raw_Quick As New DseLaporanSalesPerformance.tbltrans_sr_itemDataTable
        Dim TBLSR_Raw As New DseLaporanSalesPerformance.tbltrans_sr_itemDataTable
        Dim rowSR_Raw, rowSR_Raw1 As DseLaporanSalesPerformance.tbltrans_sr_itemRow
        Dim TipeStockInSR As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
        Dim TipeSR As ClsTransaksiSalesReturn.EnumTipeSR = ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak
        Dim cmd As String = ""
        Dim DV As DataRowView
        Dim Filter As String = ""
        Dim rowInv As DseLaporanSalesPerformance.tbltrans_salesinvoice_item1Row
        Dim DefPrec As Integer = 0

        Try
            DTASR_Raw.Connection.ConnectionString = Core_Module.GetConstring
            'isikan data sr biasa yang nota return pajak saja
            DTASR_Raw.FillBy_Normal(TBLSR_Raw, TipeSR, Me.ObjCD.GetLasDateOfTheMonth(vEndPeriod), vStartPeriod)
            'isikan data SR quick
            DTASR_Raw.Fill(TBLSR_Raw_Quick, TipeStockInSR, TipeSR, Me.ObjCD.GetLasDateOfTheMonth(vEndPeriod), vStartPeriod, 1)
            'gabungkan keduannya
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
            For Each rowSR_Raw In TBLSR_Raw.Rows
                rowSR_Raw1 = TBLSR_Raw_Quick.Newtbltrans_sr_itemRow
                For ij As Integer = 0 To TBLSR_Raw.Columns.Count - 1
                    rowSR_Raw1.Item(ij) = rowSR_Raw.Item(ij)
                Next
                TBLSR_Raw_Quick.Rows.Add(rowSR_Raw1)
            Next
            Me.IsikanReturnCost_ReturnAmoun(TBLSR_Raw_Quick, vStartPeriod, vEndPeriod)
            Filter = Me.BuatFilterSO

            TBLSR_Raw_Quick.DefaultView.RowFilter = Filter
            For Each DV In TBLSR_Raw_Quick.DefaultView
                rowSR_Raw = DV.Row
                'tambahkan baris baru ke dalam tabel sales invoice item
                rowInv = vTBLInv.Newtbltrans_salesinvoice_item1Row
                With rowInv
                    .Brand_Code = rowSR_Raw.Brand_Code
                    .Brand_Desc = rowSR_Raw.Brand_Desc
                    .Category_Code = rowSR_Raw.Category_Code
                    .Category_Desc = rowSR_Raw.Category_Desc
                    .Cost = rowSR_Raw.Cost * -1
                    .Cust_Group_Code = rowSR_Raw.Cust_Group_Code
                    .Cust_Group_Desc = rowSR_Raw.Cust_Group_Desc
                    .Cust_Type_Code = rowSR_Raw.Cust_Type_Code
                    .Cust_Type_Name = rowSR_Raw.Cust_Type_Desc
                    .Customer_Code = rowSR_Raw.Customer_Code
                    .Customer_Name = rowSR_Raw.Customer_Name
                    .Primary = -1
                    .Primary_Customer = rowSR_Raw.Primary_Customer
                    .Primary_DO = -1
                    .Primary_Product = rowSR_Raw.Primary_Product
                    .Primary_SO = -1
                    .Primary_Transaksi = -1
                    .Product_Code = rowSR_Raw.Product_Code
                    .Product_Desc = rowSR_Raw.Product_Desc
                    .Product_Group_Code = rowSR_Raw.Product_Group_Code
                    .Product_Group_Desc = rowSR_Raw.Product_Group_Desc
                    .Sales_Amount = rowSR_Raw.Amount * -1
                    .Sales_Area_Code = rowSR_Raw.Sales_Area_Code
                    .Sales_Area_Desc = rowSR_Raw.Sales_Area_Desc
                    .Sales_Person_Code = rowSR_Raw.Sales_Person_Code
                    .Sales_Person_Desc = rowSR_Raw.Sales_Person_Desc
                    .Sales_Person_Group_Code = rowSR_Raw.Sales_Person_Group_Code
                    .Sales_Person_Group_Desc = rowSR_Raw.Sales_Person_Group_Desc
                    .Segmentation_Code = rowSR_Raw.Segmentation_Code
                    .Segmentation_Desc = rowSR_Raw.Segementasi_Desc
                    .Sub_Brand_Code = rowSR_Raw.Sub_Brand_Code
                    .Sub_Brand_Desc = rowSR_Raw.Sub_Brand_Desc
                    .Tanggal_Transaksi = rowSR_Raw.Tanggal_Transaksi
                    .Type = rowSR_Raw.Type
                End With
                vTBLInv.Rows.Add(rowInv)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASR_Raw.Dispose()
            TBLSR_Raw_Quick.Dispose()
        End Try
    End Sub

    'menghitung return cost dan menghitung ulang return amount untuk sales return non-IDR
    Private Sub IsikanReturnCost_ReturnAmoun(ByVal vTBLSR_Raw As DseLaporanSalesPerformance.tbltrans_sr_itemDataTable, ByVal vStartDate As Date, ByVal vEndDate As Date)
        Dim rowSR_Raw As DseLaporanSalesPerformance.tbltrans_sr_itemRow
        Dim DTASR As New DseLaporanSalesPerformanceTableAdapters.tbltrans_sr_item_costTableAdapter
        Dim TBLSR As New DseLaporanSalesPerformance.tbltrans_sr_item_costDataTable
        Dim rowSR As DseLaporanSalesPerformance.tbltrans_sr_item_costRow
        Dim Cost As Decimal = 0
        Dim DV As DataRowView
        Dim RefDate As Date = vStartDate
        Dim DTASR1 As New DseLaporanSalesPerformanceTableAdapters.tbltrans_salesreturnTableAdapter
        Dim TBLSR1 As New DseLaporanSalesPerformance.tbltrans_salesreturnDataTable
        Dim rowSR1 As DseLaporanSalesPerformance.tbltrans_salesreturnRow
        Dim SubTotal As Decimal = 0
        Dim DefPrec As Integer = Me.ObjCurrency.GetDecimal_DefaultCurrency
        Dim TipeSR As ClsTransaksiSalesReturn.EnumTipeSR = ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak

        Try
            DTASR.Connection.ConnectionString = Core_Module.GetConstring
            DTASR1.Connection.ConnectionString = DTASR.Connection.ConnectionString

            While RefDate <= vEndDate
                DTASR.Fill(TBLSR, Me.ObjClosing.GetPrimaryClosing(Me.ObjCD.GetFirstDateofTheMonth(RefDate)), _
                ClsTransaksiMonthlyClosing.EnumTipeTranskasiAsal.InSR)
                For Each rowSR_Raw In vTBLSR_Raw.Rows
                    If Me.ObjCD.GetFirstDateofTheMonth(rowSR_Raw.Tanggal_Transaksi) = RefDate Then
                        Cost = 0
                        TBLSR.DefaultView.RowFilter = "[Primary Stock In Sales Return] = " + rowSR_Raw.Primary_Stock_In_Sales_Return.ToString + " AND " + _
                        "[Primary Product] = " + rowSR_Raw.Primary_Product.ToString
                        If TBLSR.DefaultView.Count > 0 Then
                            For Each DV In TBLSR.DefaultView
                                rowSR = DV.Row
                                Cost += rowSR.Cost
                            Next
                        End If
                        rowSR_Raw.Cost = Cost
                    End If
                Next

                RefDate = Me.ObjCD.GetFirstDateofTheMonth(RefDate + TimeSpan.FromDays(32))
            End While

            'khusus untuk sales return non-IDR, hitung sales return amountnya berdasarkan exchange rate yang digunakan
            DTASR1.Fill(TBLSR1, vEndDate, vStartDate, Core_Module.PrimaryBaseCurrency, CInt(TipeSR))
            For Each rowSR1 In TBLSR1.Rows
                rowSR1.Sub_Total_Base = Math.Round(rowSR1.Sub_Total * rowSR1.Tax_Exchange_Rate, DefPrec, MidpointRounding.AwayFromZero)

                vTBLSR_Raw.DefaultView.RowFilter = "[Primary SR] = " + rowSR1.Primary.ToString
                SubTotal = 0
                If vTBLSR_Raw.DefaultView.Count > 0 Then
                    For ij As Integer = 0 To vTBLSR_Raw.DefaultView.Count - 1
                        rowSR_Raw = vTBLSR_Raw.DefaultView(ij).Row
                        If ij = vTBLSR_Raw.DefaultView.Count - 1 Then
                            'If Math.Round(rowSR_Raw.Amount * rowSR_Raw.Exchange_Rate, DefPrec, MidpointRounding.AwayFromZero) <> rowSR1.Sub_Total_Base - SubTotal Then
                            '    MsgBox("Found it")
                            'End If
                            'baris terakhir
                            rowSR_Raw.Amount = rowSR1.Sub_Total_Base - SubTotal
                        Else
                            'bukan baris terakhir, hitung biasa, catatkan sub total
                            If rowSR_Raw.Primary_Currency <> Core_Module.PrimaryBaseCurrency Then
                                rowSR_Raw.Amount = Math.Round(rowSR_Raw.Amount * rowSR_Raw.Exchange_Rate, DefPrec, _
                                    MidpointRounding.AwayFromZero)
                                SubTotal += rowSR_Raw.Amount
                            End If
                        End If
                    Next
                Else
                    MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASR.Dispose()
            TBLSR.Dispose()
            DTASR1.Dispose()
            TBLSR1.Dispose()
        End Try
    End Sub

    Private Function ApakahLengkapKodeSalesPersonnya_SO(ByVal vStartDate As Date, ByVal vEndDate As Date) As Boolean
        Dim DTASO As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTASO.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTASO.ScalarQuery_GetJmlSOTanpaSalesPerson(Me.ObjCD.GetLasDateOfTheMonth(vEndDate), Me.ObjCD.GetFirstDateofTheMonth(vStartDate))
            If RetVal.HasValue = True Then
                If RetVal.Value > 0 Then
                    Return False
                Else
                    Return True
                End If
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASO.Dispose()
        End Try
    End Function

    Private Function ApakahLengkapKodeSalesPersonnya_Invoice(ByVal vStartDate As Date, ByVal vEndDate As Date) As Boolean
        Dim DTA As New DseTransaksiRepairSalesPersonTableAdapters.DataTable1TableAdapter
        Dim TBL As New DseTransaksiRepairSalesPerson.DataTable1DataTable
        Dim row As DseTransaksiRepairSalesPerson.DataTable1Row
        Dim DTA1 As New DseTransaksiRepairSalesPersonTableAdapters.DataTable2TableAdapter
        Dim TBL1 As New DseTransaksiRepairSalesPerson.DataTable2DataTable
        Dim row1 As DseTransaksiRepairSalesPerson.DataTable2Row

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString

            DTA.Fill(TBL, Me.ObjCD.GetLasDateOfTheMonth(vEndDate), Me.ObjCD.GetFirstDateofTheMonth(vStartDate))
            DTA1.Fill(TBL1, Me.ObjCD.GetLasDateOfTheMonth(vEndDate), Me.ObjCD.GetFirstDateofTheMonth(vStartDate))

            'remove row yang sudah ada sales personnya jika ditentukan demikian
            For ij As Integer = TBL.Rows.Count - 1 To 0 Step -1
                row = TBL.Rows(ij)
                If row.Sales_Person_Code <> "" Then
                    'remove row
                    TBL.Rows.RemoveAt(ij)
                End If
            Next
            For ij As Integer = TBL1.Rows.Count - 1 To 0 Step -1
                row1 = TBL1.Rows(ij)
                If row1.Sales_Person_Code <> "" Then
                    'remove row
                    TBL1.Rows.RemoveAt(ij)
                End If
            Next

            'pindahkan baris sr ke dalam baris invoice
            For Each row1 In TBL1.Rows
                TBL.DefaultView.RowFilter = "[Primary SO] = " + row1.Primary_SO.ToString
                If TBL.DefaultView.Count = 0 Then
                    row = TBL.NewDataTable1Row
                    With row
                        .Customer_Code = row1.Customer_Code
                        .Customer_Name = row1.Customer_Name
                        .Primary_Customer = row1.Primary_Customer
                        .primary_sales_person = row1.primary_sales_person
                        .Primary_SO = row1.Primary_SO
                        .Sales_Person_Code = row1.Sales_Person_Code
                        .Sales_Person_Name = row1.Sales_Person_Name
                        .SO_Date = row1.SO_Date
                        .So_Number = row1.SO_Number
                    End With
                    TBL.Rows.Add(row)
                End If
            Next

            'return value
            If TBL.Rows.Count > 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            DTA.Dispose()
            DTA1.Dispose()
            TBL.Dispose()
            TBL1.Dispose()
        End Try
    End Function
End Class