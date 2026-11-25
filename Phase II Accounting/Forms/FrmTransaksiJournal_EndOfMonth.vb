Option Explicit On
Option Strict Off

Imports Core_Module
Public Class FrmTransaksiJournal_EndOfMonth
    Public UserCOde As String = ""
    Public BackcolorAll As Color = Color.LightSeaGreen
    Public TipeJournal As ClsTransaksiJournal.EnumTipeJournal = ClsTransaksiJournal.EnumTipeJournal.AdjustmenJournal
    Public TipeJurnal2 As ClsTransaksiJournal.EnumTipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation

    Dim Buffer As String = ""
    Dim DefaultPrecision As Integer = 0
    Dim Baru As Boolean
    Dim ObjAkun As New ClsMasterAkun
    Dim ObjJurnal As New ClsTransaksiJournal
    Dim ObjCD As New ClsCekData
    Dim ObjQuickJurnal As New ClsTransaksiJournalQuickAdd
    Dim ObjLicense As New ClsLicense
    Dim ObjClosing As New ClsTransaksiClosing
    Dim ObjInvClosing As New ClsTransaksiMonthlyClosing
    Dim Obj_EndOFMonth As New ClsTransaksiJournal_EndOfMonth
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim PrimaryBaseCurrency As Double = 1
    Dim ObjTaxRate As New ClsTransaksiSalesInvoiceTaxRate
    Dim ObjDefaultValue As New Core_Module.ClsTransaksiSetting
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting

    Private Sub FrmTransaksiJournal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ''Me.cboJournalTipe.SelectedIndex = Me.TipeJournal
        Dim NullValue As String = ""

        Try
            Me.dgvItem.AllowUserToAddRows = False
            Me.ObjJurnal = New ClsTransaksiJournal
            Me.Baru = True
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackcolorAll
            Me.dgvExchangeRate.AlternatingRowsDefaultCellStyle.BackColor = Me.BackcolorAll
            Me.dgvItem.Rows.Clear()
            Me.dgvExchangeRate.Rows.Clear()
            Me.KosongkanField()
            Me.NonAktifkanDetail()

            If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.AdjustmenJournal Then
                Me.Text = "Adjustment Journal"
            ElseIf Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.ClosingJournal Then
                Me.Text = "Closing Journal"
                'persiapkan untuk closing 
                Me.dtpDate.CustomFormat = "yyyy"
                'jika sudah ada jurnal closing sebelumnya, maka tanggal closing harus setahun setelahnya
                If Me.ObjJurnal.GetJml(Me.TipeJournal) > 0 Then
                    Me.dtpDate.Enabled = False
                End If

                Me.cboQuickJournal.Visible = False
                Me.cboQuickJournal.TabStop = False
                Me.llbEditQuickJournal.Visible = False
                Me.LblQuickJurnal.Visible = False
            ElseIf Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.GeneralJournal Then
                If Me.TipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.Jurnal_NoGrouping Then
                    Me.Text = "General Journal"
                ElseIf Me.TipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation Then
                    Me.Text = "End Of Month Journal"
                End If
            End If
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackcolorAll

            'set null value untuk dgv
            Me.DefaultPrecision = Me.ObjCur.GetDecimal_DefaultCurrency
            NullValue = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            Me.ColDebet.DefaultCellStyle.NullValue = NullValue
            Me.ColKredit.DefaultCellStyle.NullValue = NullValue
            Me.TxtUnbalance.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)

            Me.TabControl2.TabPages.Remove(Me.ReccuringTab)
            Me.TabControl2.TabPages.Remove(Me.AssetDepreciationTab)
            'MsgBox("Belum coding ketentuan revaluasi yg ada di papan tulis")
            Me.ShowIcon = False

            Me.dtpDate.Left = Me.SplitContainer1.Panel1.Width - Me.dtpDate.Width
            Me.TxtRefNo.Width = Me.dtpDate.Left - 5 - Me.TxtRefNo.Left
            Me.TxtRemarks.Width = Me.SplitContainer1.Panel1.Width - 5 - Me.TxtRemarks.Left
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnViewJurnal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnViewJurnal.Click
        Dim FrmParent As Windows.Forms.Form = Me.MdiParent

        Try
            Dim ij As FrmTransaksiJournal_View
            Dim j As Windows.Forms.Form

            For Each j In FrmParent.MdiChildren
                ij = TryCast(j, FrmTransaksiJournal_View)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New FrmTransaksiJournal_View
            i.UserCode = Me.UserCOde
            ' Set the caption to be unique.
            i.BackCOlorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = FrmParent
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellClick
        Dim TipeExplain As String = ""
        Dim PrimaryCurrency As Double = -1
        Dim IndexAcuan As Integer = 0
        Dim rowDGV As DataGridViewRow

        Try
            If e.ColumnIndex = Me.ColShowDetail.Index Then
                IndexAcuan = Me.GetIndexAcuan(Me.GetIndexAcuan(e.RowIndex))
                rowDGV = Me.dgvItem.Rows(IndexAcuan)
                TipeExplain = rowDGV.Cells(Me.ColRemarks.Index).Value
                PrimaryCurrency = rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value

                If TipeExplain.StartsWith("revaluate AR:") = True Then
                    Me.ExplainARDetail(PrimaryCurrency, rowDGV.Cells(Me.ColPrimaryAkun.Index).Value)
                ElseIf TipeExplain.StartsWith("revaluate AP:") = True Then
                    Me.ExplainAPDetail(PrimaryCurrency, rowDGV.Cells(Me.ColPrimaryAkun.Index).Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ExplainARDetail(ByVal vPrimaryCurrency As Double, ByVal vPrimaryAkunAR As Double)
        Dim i As New FrmTransaksiJournal_EndOfMonth_DetailAPAR
        Dim rowDGV, rowDGV1 As DataGridViewRow
        Dim CurrencyCode As String = ""
        Dim DefPrecOri As Integer = 0
        Dim DefPrecBase As Integer = 0
        Dim DTAInv As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim TBLInv As New DseTransaksiJournal_EndOfMonth.tbltrans_salesinvoiceDataTable
        Dim DTASR As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltrans_salesreturnTableAdapter
        Dim TBLSR As New DseTransaksiJournal_EndOfMonth.tbltrans_salesreturnDataTable
        Dim DTAInvManual As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltrans_historicalsalesTableAdapter
        Dim TBLInvManual As New DseTransaksiJournal_EndOfMonth.tbltrans_historicalsalesDataTable
        Dim DTASRMAnual As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter
        Dim TBLSRManual As New DseTransaksiJournal_EndOfMonth.tbltrans_historicalsalesreturn_manualDataTable
        Dim DTA_CNDN As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltrans_credit_noteTableAdapter
        Dim TBL_CNDN As New DseTransaksiJournal_EndOfMonth.tbltrans_credit_noteDataTable
        Dim MinDate, MaxDate As Date

        Try
            If Me.dgvItemAP.RowCount = 0 Then
                Exit Sub
            End If
            MinDate = Date.MaxValue
            MaxDate = Date.MinValue
            For Each rowDGV In Me.dgvItemAR.Rows
                If CDate(rowDGV.Cells(Me.ColItemAR_Tanggal.Index).Value).Date < MinDate.Date Then
                    MinDate = CDate(rowDGV.Cells(Me.ColItemAR_Tanggal.Index).Value).Date
                End If
                If CDate(rowDGV.Cells(Me.ColItemAR_Tanggal.Index).Value).Date > MaxDate.Date Then
                    MaxDate = CDate(rowDGV.Cells(Me.ColItemAR_Tanggal.Index).Value).Date
                End If
            Next

            DTAInv.Connection.ConnectionString = Core_Module.GetConstring
            DTASR.Connection.ConnectionString = DTAInv.Connection.ConnectionString
            DTAInvManual.Connection.ConnectionString = DTAInv.Connection.ConnectionString
            DTASRMAnual.Connection.ConnectionString = DTAInv.Connection.ConnectionString
            DTA_CNDN.Connection.ConnectionString = Core_Module.GetConstring

            DTAInv.Fill(TBLInv, Core_Module.PrimaryBaseCurrency, MaxDate, MinDate)
            DTASR.Fill(TBLSR, Core_Module.PrimaryBaseCurrency, MaxDate, MinDate)
            DTAInvManual.Fill(TBLInvManual)
            DTASRMAnual.Fill(TBLSRManual)
            DTA_CNDN.FillBy_TglSampaiDari(TBL_CNDN, MaxDate, MinDate, Core_Module.PrimaryBaseCurrency)

            i.Show()
            i.Hide()
            CurrencyCode = Me.ObjCur.GetCurrencyCode(vPrimaryCurrency)
            DefPrecOri = Me.ObjCur.GetDefaultPrecission(vPrimaryCurrency)
            DefPrecBase = Me.ObjCur.GetDefaultPrecission(Core_Module.PrimaryBaseCurrency)

            i.Text = "Detail AR Revaluation - " + CurrencyCode

            For Each rowDGV In Me.dgvItemAR.Rows
                If rowDGV.Cells(Me.ColItemAR_PrimaryCurrency.Index).Value = vPrimaryCurrency And _
                rowDGV.Cells(Me.ColItemAR_PrimaryAkunAR.Index).Value = vPrimaryAkunAR Then
                    i.dgvDetail.Rows.Add()
                    rowDGV1 = i.dgvDetail.Rows(i.dgvDetail.Rows.Count - 1)

                    rowDGV1.Cells(i.ColBalance.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColItemAR_SaldoOriginal.Index).Value, DefPrecOri)
                    rowDGV1.Cells(i.ColDifference.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColItemAR_Selisih.Index).Value, DefPrecBase)
                    rowDGV1.Cells(i.ColExchangeRateBaru.Index).Value = Me.ObjCD.DigitGrouping(Me.GetExchangeRate(vPrimaryCurrency), Core_Module.DefPrecExchangeRate)
                    rowDGV1.Cells(i.ColExchangeRateLama.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColItemAR_ExchangeRateLama.Index).Value, Core_Module.DefPrecExchangeRate)
                    rowDGV1.Cells(i.ColNoTransaksi.Index).Value = Me.GetNumberTransaksiAR(TBLInv, TBLSR, TBLInvManual, TBLSRManual, _
                    rowDGV.Cells(Me.ColItemAR_TipeTransaksiAsal.Index).Value, rowDGV.Cells(Me.ColItemAR_PrimaryTransaksiAsal.Index).Value, TBL_CNDN)
                End If
            Next

            i.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            DTAInv.Dispose()
            TBLInv.Dispose()
            DTASR.Dispose()
            TBLSR.Dispose()
            DTAInvManual.Dispose()
            TBLInvManual.Dispose()
            DTASRMAnual.Dispose()
            TBLSRManual.Dispose()
            DTA_CNDN.Dispose()
            TBL_CNDN.Dispose()
        End Try
    End Sub

    Private Sub ExplainAPDetail(ByVal vPrimaryCurrency As Double, ByVal vPrimaryAkunAP As Double)
        Dim i As New FrmTransaksiJournal_EndOfMonth_DetailAPAR
        Dim rowDGV, rowDGV1 As DataGridViewRow
        Dim CurrencyCode As String = ""
        Dim DefPrecOri As Integer = 0
        Dim DefPrecBase As Integer = 0
        Dim DTAAOP As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltrans_aopTableAdapter
        Dim TBLAOP As New DseTransaksiJournal_EndOfMonth.tbltrans_aopDataTable
        Dim DTAPR As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltrans_purchasereturnTableAdapter
        Dim TBLPR As New DseTransaksiJournal_EndOfMonth.tbltrans_purchasereturnDataTable
        Dim DTAAOPManual As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltrans_historicalpurchasebalancemanualTableAdapter
        Dim TBLAOPManual As New DseTransaksiJournal_EndOfMonth.tbltrans_historicalpurchasebalancemanualDataTable
        Dim DTAPRMAnual As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltrans_historicalpurchasereturnbalancemanualTableAdapter
        Dim TBLPRManual As New DseTransaksiJournal_EndOfMonth.tbltrans_historicalpurchasereturnbalancemanualDataTable
        Dim MinDate, MaxDate As Date
        Dim DTA_CNDN As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltrans_credit_noteTableAdapter
        Dim TBL_CNDN As New DseTransaksiJournal_EndOfMonth.tbltrans_credit_noteDataTable

        Try
            DTA_CNDN.Connection.ConnectionString = Core_Module.GetConstring
            DTAAOPManual.Connection.ConnectionString = DTA_CNDN.Connection.ConnectionString
            DTAPRMAnual.Connection.ConnectionString = DTA_CNDN.Connection.ConnectionString
            DTAAOP.Connection.ConnectionString = DTA_CNDN.Connection.ConnectionString
            DTAPR.Connection.ConnectionString = DTA_CNDN.Connection.ConnectionString

            If Me.dgvItemAP.RowCount = 0 Then
                Exit Sub
            End If
            MinDate = Date.MaxValue
            MaxDate = Date.MinValue
            For Each rowDGV In Me.dgvItemAP.Rows
                If CDate(rowDGV.Cells(Me.ColItemAP_Tanggal.Index).Value).Date < MinDate.Date Then
                    MinDate = CDate(rowDGV.Cells(Me.ColItemAP_Tanggal.Index).Value).Date
                End If
                If CDate(rowDGV.Cells(Me.ColItemAP_Tanggal.Index).Value).Date > MaxDate.Date Then
                    MaxDate = CDate(rowDGV.Cells(Me.ColItemAP_Tanggal.Index).Value).Date
                End If
            Next

            DTAAOP.Fill(TBLAOP, Core_Module.PrimaryBaseCurrency, MaxDate, MinDate)
            DTAPR.Fill(TBLPR, Core_Module.PrimaryBaseCurrency, MaxDate, MinDate)
            DTAAOPManual.Fill(TBLAOPManual)
            DTAPRMAnual.Fill(TBLPRManual)
            DTA_CNDN.FillBy_TglSampaiDari(TBL_CNDN, MaxDate, MinDate, Core_Module.PrimaryBaseCurrency)

            i.Show()
            i.Hide()
            CurrencyCode = Me.ObjCur.GetCurrencyCode(vPrimaryCurrency)
            DefPrecOri = Me.ObjCur.GetDefaultPrecission(vPrimaryCurrency)
            DefPrecBase = Me.ObjCur.GetDefaultPrecission(Core_Module.PrimaryBaseCurrency)

            i.Text = "Detail AP Revaluation - " + CurrencyCode

            For Each rowDGV In Me.dgvItemAP.Rows
                If rowDGV.Cells(Me.ColItemAP_PrimaryCurrency.Index).Value = vPrimaryCurrency And _
                rowDGV.Cells(Me.ColItemAP_PrimaryAkunAP.Index).Value = vPrimaryAkunAP Then
                    i.dgvDetail.Rows.Add()
                    rowDGV1 = i.dgvDetail.Rows(i.dgvDetail.Rows.Count - 1)

                    rowDGV1.Cells(i.ColBalance.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColItemAP_SaldoOriginal.Index).Value, DefPrecOri)
                    rowDGV1.Cells(i.ColDifference.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColItemAP_Selisih.Index).Value, DefPrecBase)
                    rowDGV1.Cells(i.ColExchangeRateBaru.Index).Value = Me.ObjCD.DigitGrouping(Me.GetExchangeRate(vPrimaryCurrency), Core_Module.DefPrecExchangeRate)
                    rowDGV1.Cells(i.ColExchangeRateLama.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColItemAP_ExchangeRateLama.Index).Value, Core_Module.DefPrecExchangeRate)
                    rowDGV1.Cells(i.ColNoTransaksi.Index).Value = Me.GetNumberTransaksiAP(TBLAOP, TBLPR, TBLAOPManual, TBLPRManual, _
                    rowDGV.Cells(Me.ColItemAP_TipeTransaksiAsal.Index).Value, rowDGV.Cells(Me.ColItemAP_PrimaryTransaksiAsal.Index).Value, TBL_CNDN)
                End If
            Next

            i.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            DTAAOP.Dispose()
            TBLAOP.Dispose()
            DTAPR.Dispose()
            TBLPR.Dispose()
            DTAAOPManual.Dispose()
            TBLAOPManual.Dispose()
            DTAPRMAnual.Dispose()
            TBLPRManual.Dispose()
            DTA_CNDN.Dispose()
            TBL_CNDN.Dispose()
        End Try
    End Sub

    Private Function GetNumberTransaksiAR(ByVal vTBLInv As DseTransaksiJournal_EndOfMonth.tbltrans_salesinvoiceDataTable, ByVal vTBLSR As DseTransaksiJournal_EndOfMonth.tbltrans_salesreturnDataTable, _
    ByVal vTBLInvManual As DseTransaksiJournal_EndOfMonth.tbltrans_historicalsalesDataTable, ByVal vTBLSRManual As DseTransaksiJournal_EndOfMonth.tbltrans_historicalsalesreturn_manualDataTable, _
    ByVal vTipeTransaksiAsal As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal, ByVal vPrimaryTransaksiAsal As Double, _
    ByVal vTBL_CNDN As DseTransaksiJournal_EndOfMonth.tbltrans_credit_noteDataTable) As String
        Dim Filter As String = ""
        Dim rowInv As DseTransaksiJournal_EndOfMonth.tbltrans_salesinvoiceRow
        Dim rowInvManual As DseTransaksiJournal_EndOfMonth.tbltrans_historicalsalesRow
        Dim rowSRManual As DseTransaksiJournal_EndOfMonth.tbltrans_historicalsalesreturn_manualRow
        Dim rowSR As DseTransaksiJournal_EndOfMonth.tbltrans_salesreturnRow
        Dim row_CNDN As DseTransaksiJournal_EndOfMonth.tbltrans_credit_noteRow

        Try
            Filter = "Primary = " + vPrimaryTransaksiAsal.ToString
            If vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual Then
                vTBLInvManual.DefaultView.RowFilter = Filter
                If vTBLInvManual.DefaultView.Count = 0 Then
                    Return ""
                Else
                    rowInvManual = vTBLInvManual.DefaultView(0).Row
                    Return rowInvManual.Invoice_Number
                End If
            ElseIf vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis Then
                vTBLInv.DefaultView.RowFilter = Filter
                If vTBLInv.DefaultView.Count = 0 Then
                    Return ""
                Else
                    rowInv = vTBLInv.DefaultView(0).Row
                    Return rowInv.No_Transaksi
                End If
            ElseIf vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual Then
                vTBLSRManual.DefaultView.RowFilter = Filter
                If vTBLSRManual.DefaultView.Count = 0 Then
                    Return ""
                Else
                    rowSRManual = vTBLSRManual.DefaultView(0).Row
                    Return rowSRManual.AOSR_Number
                End If
            ElseIf vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis Then
                vTBLSR.DefaultView.RowFilter = Filter
                If vTBLSR.DefaultView.Count = 0 Then
                    Return ""
                Else
                    rowSR = vTBLSR.DefaultView(0).Row
                    Return rowSR.No_Transaksi
                End If
            ElseIf vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note Or _
            vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note Then
                vTBL_CNDN.DefaultView.RowFilter = Filter
                If vTBL_CNDN.DefaultView.Count = 0 Then
                    Return ""
                Else
                    row_CNDN = vTBL_CNDN.DefaultView(0).Row
                    Return row_CNDN.no_transksi
                End If
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetNumberTransaksiAP(ByVal vTBLAOP As DseTransaksiJournal_EndOfMonth.tbltrans_aopDataTable, _
    ByVal vTBLPR As DseTransaksiJournal_EndOfMonth.tbltrans_purchasereturnDataTable, _
   ByVal vTBLAOPManual As DseTransaksiJournal_EndOfMonth.tbltrans_historicalpurchasebalancemanualDataTable, _
   ByVal vTBLPRManual As DseTransaksiJournal_EndOfMonth.tbltrans_historicalpurchasereturnbalancemanualDataTable, _
   ByVal vTipeTransaksiAsal As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal, ByVal vPrimaryTransaksiAsal As Double, _
    ByVal vTBL_CNDN As DseTransaksiJournal_EndOfMonth.tbltrans_credit_noteDataTable) As String
        Dim Filter As String = ""
        Dim rowAOP As DseTransaksiJournal_EndOfMonth.tbltrans_aopRow
        Dim rowAOPManual As DseTransaksiJournal_EndOfMonth.tbltrans_historicalpurchasebalancemanualRow
        Dim rowPRManual As DseTransaksiJournal_EndOfMonth.tbltrans_historicalpurchasereturnbalancemanualRow
        Dim rowPR As DseTransaksiJournal_EndOfMonth.tbltrans_purchasereturnRow
        Dim row_CNDN As DseTransaksiJournal_EndOfMonth.tbltrans_credit_noteRow

        Try
            Filter = "Primary = " + vPrimaryTransaksiAsal.ToString
            If vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual Then
                vTBLAOPManual.DefaultView.RowFilter = Filter
                If vTBLAOPManual.DefaultView.Count = 0 Then
                    Return ""
                Else
                    rowAOPManual = vTBLAOPManual.DefaultView(0).Row
                    Return rowAOPManual.AOP_Number
                End If
            ElseIf vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis Then
                vTBLAOP.DefaultView.RowFilter = Filter
                If vTBLAOP.DefaultView.Count = 0 Then
                    Return ""
                Else
                    rowAOP = vTBLAOP.DefaultView(0).Row
                    Return rowAOP.No_Transaksi
                End If
            ElseIf vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Manual Then
                vTBLPRManual.DefaultView.RowFilter = Filter
                If vTBLPRManual.DefaultView.Count = 0 Then
                    Return ""
                Else
                    rowPRManual = vTBLPRManual.DefaultView(0).Row
                    Return rowPRManual.AOPR_Number
                End If
            ElseIf vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Otomatis Then
                vTBLPR.DefaultView.RowFilter = Filter
                If vTBLPR.DefaultView.Count = 0 Then
                    Return ""
                Else
                    rowPR = vTBLPR.DefaultView(0).Row
                    Return rowPR.No_Transaksi
                End If
            ElseIf vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note Or _
            vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note Then
                vTBL_CNDN.DefaultView.RowFilter = Filter
                If vTBL_CNDN.DefaultView.Count = 0 Then
                    Return ""
                Else
                    row_CNDN = vTBL_CNDN.DefaultView(0).Row
                    Return row_CNDN.no_transksi
                End If
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetIndexAcuan(ByVal vrowIndex As Integer) As Integer
        Dim rowDGV As DataGridViewRow

        Try
            For i As Integer = vrowIndex To 0 Step -1
                rowDGV = Me.dgvItem.Rows(i)
                If rowDGV.Cells(Me.ColRemarks.Index).Value <> "" Then
                    Return rowDGV.Index
                End If
            Next

            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvItemProcess_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Dim rowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0

        Try
            Me.dgvItem.TabStop = True
            If Me.TxtRemarks.Enabled = True Then
                'set nilai default baris baru
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value = Me.PrimaryBaseCurrency
                DefPrecBase = Me.ObjCur.GetDefaultPrecission(Me.PrimaryBaseCurrency)
                rowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrecBase)
                rowDGV.Cells(Me.ColCredit_OriginalCurrency.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrecBase)
                rowDGV.Cells(Me.ColDebet.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrecBase)
                rowDGV.Cells(Me.ColKredit.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrecBase)
                rowDGV.Cells(Me.ColExchangeRate.Index).Value = Me.ObjCD.DigitGrouping(1, DefPrecExchangeRate)
                rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value = Me.PrimaryBaseCurrency
                rowDGV.Cells(Me.ColOriginalCurrency.Index).Value = Me.ObjCur.GetCurrencyCode(Me.PrimaryBaseCurrency)
                rowDGV.Cells(Me.ColPrimaryAkun.Index).Value = -1
                rowDGV.Cells(Me.ColAccount.Index).Value = ""
                rowDGV.Cells(Me.ColRemarks.Index).Value = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProcess_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            ''hitung debet dan kredit
            Dim row As DataGridViewRow
            Dim TotalDebet As Decimal = 0
            Dim TotalKredit As Decimal = 0
            For Each row In dgvItem.Rows
                TotalDebet += CDec(Me.dgvItem.Item(Me.ColDebet.Index, row.Index).Value)
                TotalKredit += CDec(Me.dgvItem.Item(Me.ColKredit.Index, row.Index).Value)
            Next
            Me.TxtUnbalance.Text = CStr(ObjCD.DigitGrouping(Me.HitungUnbalance(TotalDebet, TotalKredit), Me.DefaultPrecision))
            Me.SetMulticurrency()
            'If Me.TxtRemarks.Enabled = True Then
            'If Me.dgvItem.RowCount = 0 Then
            ''
            'End If
            'End If
            ''Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' tampilkan row yang diberikan pada parameter
    ''' </summary>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Private Sub ShowRowProcess(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < Me.dgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.dgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While Me.dgvItem.Rows(row.Index).Displayed = False
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False) Then
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmTransaksiJournal_Cetak
        ij.TipeJournal = Me.TipeJournal
        ij.TipeJurnal2 = Me.TipeJurnal2

        Try
            ij.BackColorAll = Me.BackcolorAll
            ij.ShowDialog()
            Me.dgvItem.AllowUserToAddRows = False
            If Not ij.PrimaryPP = -1 Then
                If Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, ClsTransaksiSetting.EnumNamaParameter.Jurnal_IsLedgerAccessActive) = "1" Then
                    If Me.ObjJurnal.ApakahMemilikiLedgerAccess(Me.UserCOde, Me.ObjJurnal.GetCreatedBy(ij.PrimaryPP)) = False Then
                        MsgBox("You are not authorized to view this data.", MsgBoxStyle.Information, "Information")
                        Me.KosongkanField()
                        Exit Sub
                    End If
                End If

                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                If Me.ObjJurnal.CariDetail_EndOfMonth(ij.PrimaryPP, Me.dgvItem, Me, ij.TipeJournal, Me.dgvExchangeRate, Me.dgvItemAR, Me.dgvItemAP) = True Then
                    Me.SalinkeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    ''fungsi standar 
    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryJournal.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimaryJournal.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimaryJournal.BackColor, False)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.TabPage2.Controls, True, False, True, False, True, False, True, True)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None

            Me.dgvItem.AllowUserToDeleteRows = True
            Me.dgvItem.ReadOnly = False

            ''Me.txtNoTransaksi.ReadOnly = True
            ''Me.txtNoTransaksi.BackColor = Me.txtReadOnly.BackColor
            Me.cboQuickJournal.SelectedIndex = -1
            Me.llbEditQuickJournal.Enabled = True
            Me.ObjQuickJurnal.IsiCboQuickAdd(Me.TipeJournal, Me.cboQuickJournal)

            Me.dtpDate.Enabled = False
            Me.dgvItem.AllowUserToAddRows = False
            Me.dgvItem.ReadOnly = True

            Me.dgvExchangeRate.ReadOnly = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.TabPage2.Controls, True, True, True, True, True, True, True, True)
            Me.CancelButton = Me.btnExit

            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True

            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.BackColor, True)

            Me.cboQuickJournal.SelectedIndex = -1
            Me.cboQuickJournal.DataSource = Nothing
            Me.cboQuickJournal.Items.Clear()
            Me.llbEditQuickJournal.Enabled = False

            Me.dgvExchangeRate.ReadOnly = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, False)
            Me.TxtPrimaryJournal.Text = "-1"
            Me.TxtUnbalance.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "

            Me.dgvItem.Rows.Clear()
            Me.dgvExchangeRate.Rows.Clear()
            Me.dgvItemAR.Rows.Clear()
            ''Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinkeObjek()
        Try
            With ObjJurnal
                .RefNumber = Me.TxtRefNo.Text
                .Description = Me.TxtRemarks.Text
                .Tanggal = Me.dtpDate.Value
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinkeField()
        Try
            With ObjJurnal
                Me.TxtPrimaryJournal.Text = .Primary
                Me.TxtRefNo.Text = .RefNumber
                Me.TxtRemarks.Text = .Description
                Me.dtpDate.Value = .Tanggal
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString

                Me.SetMulticurrency()
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim row As DataGridViewRow

        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("Reference number")
            KoleksiTextBox.Add(Me.TxtRefNo)
            KoleksiNamaData.Add("Description")
            KoleksiTextBox.Add(Me.TxtRemarks)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data default precision harus non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'Cek DGVItem
            'If Me.dgvItem.Rows.Count < 2 Then
            'MsgBox("Minimum item = 2, please input item again", MsgBoxStyle.Information, "Information")
            'Me.dgvItem.Focus()
            'Return False
            'End If

            ' Cek debet kredit harus balance
            Me.HitungUnbalance()
            If CDec(Me.TxtUnbalance.Text) <> 0 Then
                MsgBox("This is an unbalance account. Please add numbers in debet or credit column so that a balanced calculation", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'tidak boleh bikin transaksi di belakang convertion date
            If Me.ObjClosing.GetConvertionDate.Date > Me.dtpDate.Value.Date Then
                MsgBox("Minimum transaction date is " + Format(Me.ObjClosing.GetConvertionDate.Date, "dd MMM yyyy") + ".", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek apakah ada account number yang debet dan kreditnya bernilai 0
            '2019-09-18: batasan ini dihapuskan karena ada kemungkinan nilai balance debit kredit nol karena saling cancel out
            'For Each row In Me.dgvItem.Rows
            '    If Me.dgvItem.Item(Me.ColDebet.Index, row.Index).Value = 0 And _
            '    Me.dgvItem.Item(Me.ColKredit.Index, row.Index).Value = 0 Then
            '        MsgBox("Please add positive number in account no. " + Me.dgvItem.Item(Me.ColAccount.Index, row.Index).Value, _
            '        MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If
            'Next

            'cek data yang salah, double dalam kode
            If Me.Baru = True Then
                If Me.ObjJurnal.ApakahSudahAda(Me.TxtRefNo.Text) = True Then
                    MsgBox("Duplication in no journal =" + Me.TxtRefNo.Text + ", please change your number", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtRefNo.Focus()
                    Return False
                End If
            Else
                Dim TipeCodeLama As String = Me.ObjJurnal.GetRefCode(Me.TxtPrimaryJournal.Text)
                If TipeCodeLama <> Me.TxtRefNo.Text Then
                    If Me.ObjJurnal.ApakahSudahAda(Me.TxtRefNo.Text) = True Then
                        MsgBox("Duplication in no journal =" + Me.TxtRefNo.Text + ", please change your number", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtRefNo.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.ObjJurnal.GetTimestamp(Me.TxtPrimaryJournal.Text)
                If Me.ObjJurnal.Timestamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjJurnal.CariDetail_EndOfMonth(Me.ObjJurnal.Primary, Me.dgvItem, Me, Me.TipeJournal, Me.dgvExchangeRate, Me.dgvItemAR, _
                        Me.dgvItemAP)
                        Me.SalinkeField()
                    End If
                    Return False
                End If
            End If

            'pastikan semua kode akun ada di dalam master akun
            For Each row In Me.dgvItem.Rows
                If IsNothing(row.Cells(Me.ColPrimaryAkun.Index).Value) = True Then
                    MsgBox("Invalid account in row " + row.Index.ToString, MsgBoxStyle.Information, "Information")
                    Me.dgvItem.ClearSelection()
                    row.Selected = True
                    Return False
                Else
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader2(row.Cells(Me.ColPrimaryAkun.Index).Value) = False Then
                        MsgBox("Invalid account in row " + row.Index.ToString, MsgBoxStyle.Information, "Information")
                        Me.dgvItem.ClearSelection()
                        row.Selected = True
                        Return False
                    End If
                End If
            Next

            'jika sudah ada closing jurnal untuk tahun ybs maka tidak bisa input data baru
            If Me.ObjJurnal.GetJml(ClsTransaksiJournal.EnumTipeJournal.ClosingJournal, Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpDate.Value), _
            Me.ObjCD.GetLastDateOfTheYear(Me.dtpDate.Value.Date)) > 0 Then
                MsgBox("There are closing jurnal for this year you cannot save this data", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Function

    Private Sub txtRemarks_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.Enter
        Try
            Me.Buffer = Me.TxtRemarks.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''back color dan tampilan 
    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.ReadOnly = False Then
                    Me.TxtRemarks.BackColor = Me.BackcolorAll
                Else
                    TxtRemarks.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.ReadOnly = False Then
                    Me.TxtRemarks.BackColor = Me.TxtPrimaryJournal.BackColor
                Else
                    TxtRemarks.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtScheduleNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRefNo.GotFocus
        Try
            If Me.TxtRefNo.Enabled = True Then
                If Me.TxtRefNo.ReadOnly = False Then
                    Me.TxtRefNo.BackColor = Me.BackcolorAll
                Else
                    TxtRefNo.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtScheduleNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRefNo.LostFocus
        Try
            If Me.TxtRefNo.Enabled = True Then
                If Me.TxtRefNo.ReadOnly = False Then
                    Me.TxtRefNo.BackColor = Me.TxtPrimaryJournal.BackColor
                Else
                    TxtRefNo.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Dim PrimaryDummy As Double = -1
        Dim PrimaryAkunRE As String = ""
        Dim TglRevaluasiSelanjutnya As Date = #1/1/1950#
        Dim DTA_Unbalance_Journal As New AccModule.DseTransaksiJournalTableAdapters.UnBalanceJurnalTableAdapter
        Dim TBL_Unbalance_Journal As New AccModule.DseTransaksiJournal.UnBalanceJurnalDataTable
        Dim row_UnBalance_Journal As AccModule.DseTransaksiJournal.UnBalanceJurnalRow
        Dim Counter As Double = 0
        Dim Dummy As String = ""

        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'cek apakah ada unbalance jurnal 
            DTA_Unbalance_Journal.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Unbalance_Journal.Fill(TBL_Unbalance_Journal)
            Counter = 1
            If TBL_Unbalance_Journal.Rows.Count > 0 Then
                For Each row_UnBalance_Journal In TBL_Unbalance_Journal.Rows
                    If Dummy = "" Then
                        Dummy = Counter.ToString + ". " + row_UnBalance_Journal.Ref_Number
                    Else
                        Dummy = Dummy + vbCrLf + Counter.ToString + ". " + row_UnBalance_Journal.Ref_Number
                    End If
                    Counter += 1
                Next

                MsgBox("The following journal is unbalance, please contact your vendor to solve this issue." + vbCrLf + Dummy, _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.ObjJurnal.GetJml(Me.TipeJournal, Me.TipeJurnal2) = 0 Then
                'jika belum ada jurnal revaluasi sebelumnya, maka pastikan minimal sudah ada 1 inventory closing
                'jika tidak ada maka tidak bisa melakukan revaluasi
                If Me.ObjInvClosing.GetJml = 0 Then
                    MsgBox("Please do inventory closing first.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                'tanggal revaluasi pertama adalah akhir bulan untuk inventory closing pertama
                Me.dtpDate.Value = Me.ObjCD.GetLasDateOfTheMonth(Me.ObjInvClosing.GetPeriodeClosingPertama + TimeSpan.FromDays(33))
                TglRevaluasiSelanjutnya = Me.dtpDate.Value
            Else
                TglRevaluasiSelanjutnya = Me.ObjJurnal.GetTglRevaluasiSelanjutnya
                If TglRevaluasiSelanjutnya = #1/1/1950# Then
                    MsgBox("Please do inventory closing first.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                'jika sudah ada jurnal revaluasi sebelumnya, maka pastikan sudah ada inventory closing untuk revaluasi selanjutnya
                If Me.ObjInvClosing.GetJml(Me.ObjCD.GetFirstDateofTheMonth(TglRevaluasiSelanjutnya)) = 0 Then
                    MsgBox("Please do inventory closing first.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                Me.dtpDate.Value = TglRevaluasiSelanjutnya
                TglRevaluasiSelanjutnya = Me.dtpDate.Value
            End If

            'Me.dgvItem.AllowUserToAddRows = True
            Me.Baru = True
            Me.KosongkanField()
            If Me.ObjJurnal.GetJml(Me.TipeJournal, Me.TipeJurnal2) = 0 Then
                'TglRevaluasiSelanjutnya = Me.ObjCD.GetLasDateOfTheMonth(OBJSE)
                'TglRevaluasiSelanjutnya = #5/31/2013#
            End If
            Me.dtpDate.Value = TglRevaluasiSelanjutnya
            Me.AktifkanDetail()
            Me.TxtRefNo.Focus()
            Me.ObjJurnal.ResetObject(Me.dgvItem)
            Me.ObjJurnal.TipeJurnal2 = Me.TipeJurnal2
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCOde
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCOde
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

            'isikan exhange rate berdasarkan default exchanget rate
            Me.IsikanDefaultExchangeRate()
            Me.IsikanRevaluasi()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Unbalance_Journal.Dispose()
            TBL_Unbalance_Journal.Dispose()
        End Try
    End Sub

    Private Sub IsikanDefaultExchangeRate()
        Try
            Me.dgvExchangeRate.Rows.Clear()
            Me.Obj_EndOFMonth.IsikanDGVCurrencyExchangeRate(Me.dgvExchangeRate, Me, Me.dtpDate.Value)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanRevaluasi()
        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.Rows.Clear()

            'isikan revaluasi kas bank
            Me.IsikanRevaluasiKasBank()

            'isikan revaluasi ar
            Me.IsikanRevaluasiAR()

            'isikan revaluasi ap
            Me.IsikanRevaluasiAP()

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'hitung unbalance
            Me.HitungUnbalance()
            Me.SetMulticurrency()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetExchangeRate(ByVal vPrimaryTransaksiAsal As Double, ByVal vTipeTransaksiAsal As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal, _
    ByVal vTBLExch As DseTransaksiJournal_EndOfMonth.tbltrans_balance_salesDataTable, ByVal vPrimaryCurrency As Double) As Decimal
        Dim row As DseTransaksiJournal_EndOfMonth.tbltrans_balance_salesRow
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance
        Dim ObjBalance As New ClsTransaksiSalesBalance
        Dim ObjKartu As New ClsTransaksi_Kartu_APAR
        Dim Filter As String = ""

        Try
            TipeBalance = ObjKartu.GetTipeSalesBalance(vTipeTransaksiAsal)
            Filter = "Tipe = " + CInt(TipeBalance).ToString + " AND [Primary Invoice] = " + _
            vPrimaryTransaksiAsal.ToString + " AND [primary currency] = " + vPrimaryCurrency.ToString

            vTBLExch.DefaultView.RowFilter = Filter
            If vTBLExch.DefaultView.Count = 0 Then
                Return 1
            Else
                row = vTBLExch.DefaultView(0).Row
                Return row.exchange_rate
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 1
        End Try
    End Function

    Private Function GetExchangeRate_ByItemJurnal(ByVal vPrimaryTransaksiAsal As Double, ByVal vTipeTransaksiAsal As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal, _
  ByVal vTBLExch As DseTransaksiJournal.tbltrans_jurnal_item_revaluasi_arDataTable, ByVal vPrimaryCurrency As Double) As Decimal
        Dim row As DseTransaksiJournal.tbltrans_jurnal_item_revaluasi_arRow
        Dim ObjKartu As New ClsTransaksi_Kartu_APAR
        Dim Filter As String = ""

        Try
            Filter = "[tipe transaksi asal] = " + CInt(vTipeTransaksiAsal).ToString + " AND [primary transaksi asal] = " + _
            vPrimaryTransaksiAsal.ToString + " AND [primary currency] = " + vPrimaryCurrency.ToString

            vTBLExch.DefaultView.RowFilter = Filter
            If vTBLExch.DefaultView.Count = 0 Then
                Return 1
            Else
                row = vTBLExch.DefaultView(0).Row
                Return row.exchange_rate_lama
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 1
        End Try
    End Function

    Private Function GetExchangeRate(ByVal vPrimaryTransaksiAsal As Double, ByVal vTipeTransaksiAsal As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal, _
  ByVal vTBLExch As DseTransaksiJournal_EndOfMonth.tbltrans_balance_purchaseDataTable, ByVal vPrimaryCurrency As Double) As Decimal
        Dim row As DseTransaksiJournal_EndOfMonth.tbltrans_balance_purchaseRow
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance
        Dim ObjBalance As New ClsTransaksiSalesBalance
        Dim ObjKartu As New ClsTransaksi_Kartu_APAR
        Dim Filter As String = ""

        Try
            TipeBalance = ObjKartu.GetTipeSalesBalance(vTipeTransaksiAsal)
            Filter = "Tipe = " + CInt(TipeBalance).ToString + " AND [Primary AOP] = " + _
            vPrimaryTransaksiAsal.ToString + " AND [primary currency] = " + vPrimaryCurrency.ToString

            vTBLExch.DefaultView.RowFilter = Filter
            If vTBLExch.DefaultView.Count = 0 Then
                Return 1
            Else
                row = vTBLExch.DefaultView(0).Row
                Return row.exchange_rate
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 1
        End Try
    End Function

    Private Function GetExchangeRate_ByItemJurnal(ByVal vPrimaryTransaksiAsal As Double, ByVal vTipeTransaksiAsal As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal, _
  ByVal vTBLExch As DseTransaksiJournal.tbltrans_jurnal_item_revaluasi_apDataTable, ByVal vPrimaryCurrency As Double) As Decimal
        Dim row As DseTransaksiJournal.tbltrans_jurnal_item_revaluasi_apRow
        Dim ObjKartu As New ClsTransaksi_Kartu_APAR
        Dim Filter As String = ""

        Try
            Filter = "[tipe transaksi asal] = " + CInt(vTipeTransaksiAsal).ToString + " AND [primary transaksi asal] = " + _
            vPrimaryTransaksiAsal.ToString + " AND [primary currency] = " + vPrimaryCurrency.ToString

            vTBLExch.DefaultView.RowFilter = Filter
            If vTBLExch.DefaultView.Count = 0 Then
                Return 1
            Else
                row = vTBLExch.DefaultView(0).Row
                Return row.exchange_rate_lama
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 1
        End Try
    End Function

    Private Sub IsikanRevaluasiAR()
        Dim DTA As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltransaksi_kartu_ar_apTableAdapter
        Dim TBL As New DseTransaksiJournal_EndOfMonth.tbltransaksi_kartu_ar_apDataTable
        Dim row As DseTransaksiJournal_EndOfMonth.tbltransaksi_kartu_ar_apRow
        Dim TipeKartu As ClsTransaksi_Kartu_APAR.EnumJenisKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR
        Dim rowDGV, rowDGV1 As DataGridViewRow
        Dim TBLExch As New DseTransaksiJournal_EndOfMonth.tbltrans_balance_salesDataTable
        Dim DTAExch As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltrans_balance_salesTableAdapter
        Dim Selisih As Decimal = 0
        Dim ExchangeRateLama As Decimal = 0
        Dim ExchangeRateBaru As Decimal = 0
        Dim ListKombinasi As New List(Of String)
        Dim SaldoOriginal, Saldo As Decimal
        Dim Debet, Kredit, DebetOriginal, KreditOriginal, ExchangeRate As Decimal
        Dim DefPrecBase As Integer = 0

        Dim DTAItem As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_revaluasi_arTableAdapter
        Dim TBLItem As New DseTransaksiJournal.tbltrans_jurnal_item_revaluasi_arDataTable
        Dim rowItem As DseTransaksiJournal.tbltrans_jurnal_item_revaluasi_arRow
        Dim Filter As String = ""

        Try
            DefPrecBase = Me.ObjCur.GetDefaultPrecission(Core_Module.PrimaryBaseCurrency)
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAExch.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAItem.Connection.ConnectionString = DTA.Connection.ConnectionString

            Me.dgvItemAR.Rows.Clear()

            'kumpulkan balance ar yang masih <> 0 pada saat revalasi dan yang mata uangnya non idr
            DTA.Fill(TBL, TipeKartu, Me.dtpDate.Value, Core_Module.PrimaryBaseCurrency)
            DTAExch.Fill(TBLExch)

            'modifikan, kembalikan ke nilai semua untuk item balance ar
            If Me.Baru = False Then
                'jurnal yang ada tambahkan selisih dari item ar jurnal
                DTAItem.FillBy_PrimaryJurnal(TBLItem, Me.TxtPrimaryJournal.Text)
                For Each rowItem In TBLItem.Rows
                    If rowItem.selisih <> 0 Then
                        'tambahkan dengan selisih item
                        Filter = "[Primary Transaksi Asal] = " + rowItem.primary_transaksi_asal.ToString + _
                        " AND [Type Transaksi Asal] = " + rowItem.tipe_transaksi_asal.ToString

                        TBL.DefaultView.RowFilter = Filter

                        If TBL.DefaultView.Count > 0 Then
                            row = TBL.DefaultView(0).Row
                            row.Saldo += rowItem.selisih
                        End If
                    End If
                Next
            End If
            TBL.DefaultView.RowFilter = ""

            'catatkan pada data item, tipe transaksi asal, primary transaksi asal, exchange rate yang lama, exchange rate yang baru, nilai balance, 
            For Each row In TBL.Rows
                Me.dgvItemAR.Rows.Add()
                rowDGV = Me.dgvItemAR.Rows(Me.dgvItemAR.Rows.Count - 1)

                With rowDGV
                    If Me.Baru = True Then
                        ExchangeRateLama = Me.GetExchangeRate(row.Primary_Transaksi_Asal, row.Type_Transaksi_Asal, _
                        TBLExch, row.primary_currency)
                    Else
                        ExchangeRateLama = Me.GetExchangeRate_ByItemJurnal(row.Primary_Transaksi_Asal, row.Type_Transaksi_Asal, _
                        TBLItem, row.primary_currency)
                    End If
                    ExchangeRateBaru = Me.GetExchangeRate(row.primary_currency)
                    .Cells(Me.ColItemAR_ExchangeRateLama.Index).Value = ExchangeRateLama
                    .Cells(Me.ColItemAR_PrimaryAkunAR.Index).Value = row.primary_akun
                    .Cells(Me.ColItemAR_PrimaryCurrency.Index).Value = row.primary_currency
                    .Cells(Me.ColItemAR_PrimaryTransaksiAsal.Index).Value = row.Primary_Transaksi_Asal
                    .Cells(Me.ColItemAR_Saldo.Index).Value = row.Saldo
                    .Cells(Me.ColItemAR_SaldoOriginal.Index).Value = row.Saldo_Original
                    Selisih = Math.Round(CDec(.Cells(Me.ColItemAR_SaldoOriginal.Index).Value) * ExchangeRateLama, DefPrecBase) _
                    - Math.Round(CDec(.Cells(Me.ColItemAR_SaldoOriginal.Index).Value) * ExchangeRateBaru, DefPrecBase)
                    .Cells(Me.ColItemAR_Selisih.Index).Value = Selisih
                    .Cells(Me.ColItemAR_TipeTransaksiAsal.Index).Value = row.Type_Transaksi_Asal
                    .Cells(Me.ColItemAR_PaymentTermAcuan.Index).Value = row.Payment_Terms_Transaksi_Acuan
                    .Cells(Me.ColItemAR_PrimaryPihakAcuan.Index).Value = row.Primary_Customer
                    .Cells(Me.ColItemAR_Tanggal.Index).Value = row._Date
                End With
            Next

            'untuk setiap mata uang
            'buat item jurnal untuk membalikan balance menjadi nol
            'Exit Sub
            For Each rowDGV In Me.dgvItemAR.Rows
                If ListKombinasi.Contains(rowDGV.Cells(Me.ColItemAR_PrimaryCurrency.Index).Value.ToString + "-" + rowDGV.Cells(Me.ColItemAR_PrimaryAkunAR.Index).Value.ToString) = False Then
                    ListKombinasi.Add(rowDGV.Cells(Me.ColItemAR_PrimaryCurrency.Index).Value.ToString + "-" + rowDGV.Cells(Me.ColItemAR_PrimaryAkunAR.Index).Value.ToString)

                    SaldoOriginal = Me.GetSaldoOriginal(rowDGV.Cells(Me.ColItemAR_PrimaryCurrency.Index).Value, rowDGV.Cells(Me.ColItemAR_PrimaryAkunAR.Index).Value)
                    Saldo = Me.GetSaldo(rowDGV.Cells(Me.ColItemAR_PrimaryCurrency.Index).Value, rowDGV.Cells(Me.ColItemAR_PrimaryAkunAR.Index).Value)
                    Selisih = Me.GetSelisihAR(rowDGV.Cells(Me.ColItemAR_PrimaryCurrency.Index).Value, rowDGV.Cells(Me.ColItemAR_PrimaryAkunAR.Index).Value)

                    If Saldo > 0 Then
                        Debet = 0
                        Kredit = Saldo
                        DebetOriginal = 0
                        KreditOriginal = SaldoOriginal
                        ExchangeRate = Math.Round(Saldo / SaldoOriginal, Core_Module.DefPrecExchangeRate)
                    Else
                        Debet = Math.Abs(Saldo)
                        Kredit = 0
                        DebetOriginal = Math.Abs(SaldoOriginal)
                        KreditOriginal = 0
                        ExchangeRate = Math.Abs(Math.Round(Saldo / SaldoOriginal, Core_Module.DefPrecExchangeRate))
                    End If

                    Me.AddRowItem(rowDGV.Cells(Me.ColItemAR_PrimaryAkunAR.Index).Value, ExchangeRate, _
                    rowDGV.Cells(Me.ColItemAR_PrimaryCurrency.Index).Value, Debet, Kredit, DebetOriginal, KreditOriginal)
                    rowDGV1 = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    rowDGV1.Cells(Me.ColRemarks.Index).Value = "revaluate AR: " + _
                    Me.ObjCur.GetCurrencyCode(rowDGV.Cells(Me.ColItemAR_PrimaryCurrency.Index).Value)

                    'buat item jurnal untuk mengembalikan nilai balance ke nilai original semula
                    If Saldo > 0 Then
                        KreditOriginal = 0
                        DebetOriginal = SaldoOriginal
                        ExchangeRate = Me.GetExchangeRate(rowDGV.Cells(Me.ColItemAR_PrimaryCurrency.Index).Value)
                        Kredit = 0
                        Debet = Saldo - Selisih
                    Else
                        Debet = 0
                        Kredit = Math.Abs(Saldo - Selisih)
                        DebetOriginal = 0
                        KreditOriginal = Math.Abs(SaldoOriginal)
                        ExchangeRate = Me.GetExchangeRate(rowDGV.Cells(Me.ColItemAR_PrimaryCurrency.Index).Value)
                    End If
                    Me.AddRowItem(rowDGV.Cells(Me.ColItemAR_PrimaryAkunAR.Index).Value, ExchangeRate, _
                    rowDGV.Cells(Me.ColItemAR_PrimaryCurrency.Index).Value, Debet, Kredit, DebetOriginal, KreditOriginal)

                    'buat jurnal item selisih exchange rate
                    If Selisih <> 0 Then
                        If Selisih > 0 Then
                            Debet = Selisih
                            Kredit = 0
                            DebetOriginal = Debet
                            KreditOriginal = Kredit
                        Else
                            Debet = 0
                            Kredit = Math.Abs(Selisih)
                            DebetOriginal = Debet
                            KreditOriginal = Kredit
                        End If
                        ExchangeRate = 1
                        Me.AddRowItem(Me.ObjCur.GetPrimary_Akun_Realize_Exch_Rate(rowDGV.Cells(Me.ColItemAR_PrimaryCurrency.Index).Value), ExchangeRate, _
                        Core_Module.PrimaryBaseCurrency, Debet, Kredit, DebetOriginal, KreditOriginal)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ListKombinasi = Nothing
            DTAItem.Dispose()
            TBLItem.Dispose()
        End Try
    End Sub

    Private Sub IsikanRevaluasiAP()
        Dim DTA As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltransaksi_kartu_ar_ap1TableAdapter
        Dim TBL As New DseTransaksiJournal_EndOfMonth.tbltransaksi_kartu_ar_ap1DataTable
        Dim row As DseTransaksiJournal_EndOfMonth.tbltransaksi_kartu_ar_ap1Row
        Dim TipeKartu As ClsTransaksi_Kartu_APAR.EnumJenisKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP
        Dim rowDGV, rowDGV1 As DataGridViewRow
        Dim TBLExch As New DseTransaksiJournal_EndOfMonth.tbltrans_balance_purchaseDataTable
        Dim DTAExch As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltrans_balance_purchaseTableAdapter
        Dim Selisih As Decimal = 0
        Dim ExchangeRateLama As Decimal = 0
        Dim ExchangeRateBaru As Decimal = 0
        Dim ListKombinasi As New List(Of String)
        Dim SaldoOriginal, Saldo As Decimal
        Dim Debet, Kredit, DebetOriginal, KreditOriginal, ExchangeRate As Decimal
        Dim DefPrecBase As Integer = 0

        Dim DTAItem As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_revaluasi_apTableAdapter
        Dim TBLItem As New DseTransaksiJournal.tbltrans_jurnal_item_revaluasi_apDataTable
        Dim rowItem As DseTransaksiJournal.tbltrans_jurnal_item_revaluasi_apRow
        Dim Filter As String = ""

        Try
            DefPrecBase = Me.ObjCur.GetDefaultPrecission(Core_Module.PrimaryBaseCurrency)
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAExch.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAItem.Connection.ConnectionString = DTA.Connection.ConnectionString

            Me.dgvItemAP.Rows.Clear()

            'kumpulkan balance ar yang masih <> 0 pada saat revalasi dan yang mata uangnya non idr
            DTA.Fill(TBL, TipeKartu, Me.dtpDate.Value, Core_Module.PrimaryBaseCurrency)
            DTAExch.Fill(TBLExch)

            'modifikan, kembalikan ke nilai semua untuk item balance ar
            If Me.Baru = False Then
                'jurnal yang ada tambahkan selisih dari item ar jurnal
                DTAItem.FillBy_PrimaryTransaksi(TBLItem, Me.TxtPrimaryJournal.Text)
                For Each rowItem In TBLItem.Rows
                    If rowItem.selisih <> 0 Then
                        'tambahkan dengan selisih item
                        Filter = "[Primary Transaksi Asal] = " + rowItem.primary_transaksi_asal.ToString + _
                        " AND [Type Transaksi Asal] = " + rowItem.tipe_transaksi_asal.ToString

                        TBL.DefaultView.RowFilter = Filter

                        If TBL.DefaultView.Count > 0 Then
                            row = TBL.DefaultView(0).Row
                            row.Saldo += rowItem.selisih
                        End If
                    End If
                Next
            End If
            TBL.DefaultView.RowFilter = ""

            'catatkan pada data item, tipe transaksi asal, primary transaksi asal, exchange rate yang lama, exchange rate yang baru, nilai balance, 
            For Each row In TBL.Rows
                If row.Saldo <> 0 Then
                    Me.dgvItemAP.Rows.Add()
                    rowDGV = Me.dgvItemAP.Rows(Me.dgvItemAP.Rows.Count - 1)

                    With rowDGV
                        If Me.Baru = True Then
                            ExchangeRateLama = Me.GetExchangeRate(row.Primary_Transaksi_Asal, row.Type_Transaksi_Asal, _
                            TBLExch, row.primary_currency)
                        Else
                            ExchangeRateLama = Me.GetExchangeRate_ByItemJurnal(row.Primary_Transaksi_Asal, row.Type_Transaksi_Asal, _
                            TBLItem, row.primary_currency)
                        End If
                        ExchangeRateBaru = Me.GetExchangeRate(row.primary_currency)
                        .Cells(Me.ColItemAP_ExchangeRateLama.Index).Value = ExchangeRateLama
                        .Cells(Me.ColItemAP_PrimaryAkunAP.Index).Value = row.primary_akun
                        .Cells(Me.ColItemAP_PrimaryCurrency.Index).Value = row.primary_currency
                        .Cells(Me.ColItemAP_PrimaryTransaksiAsal.Index).Value = row.Primary_Transaksi_Asal
                        .Cells(Me.ColItemAP_Saldo.Index).Value = row.Saldo
                        .Cells(Me.ColItemAP_SaldoOriginal.Index).Value = row.Saldo_Original
                        Selisih = Math.Round(CDec(.Cells(Me.ColItemAP_SaldoOriginal.Index).Value) * ExchangeRateBaru, DefPrecBase) - _
                        Math.Round(CDec(.Cells(Me.ColItemAP_SaldoOriginal.Index).Value) * ExchangeRateLama, DefPrecBase)
                        .Cells(Me.ColItemAP_Selisih.Index).Value = Selisih
                        .Cells(Me.ColItemAP_TipeTransaksiAsal.Index).Value = row.Type_Transaksi_Asal
                        .Cells(Me.ColItemAP_PaymentTermAcuan.Index).Value = row.Payment_Terms_Transaksi_Acuan
                        .Cells(Me.ColItemAP_PrimaryPihakAcuan.Index).Value = row.Primary_Supplier
                        .Cells(Me.ColItemAP_Tanggal.Index).Value = row._Date
                    End With
                End If

            Next

            'untuk setiap mata uang
            'buat item jurnal untuk membalikan balance menjadi nol
            'Exit Sub
            For Each rowDGV In Me.dgvItemAP.Rows
                If ListKombinasi.Contains(rowDGV.Cells(Me.ColItemAP_PrimaryCurrency.Index).Value.ToString + "-" + rowDGV.Cells(Me.ColItemAP_PrimaryAkunAP.Index).Value.ToString) = False Then
                    ListKombinasi.Add(rowDGV.Cells(Me.ColItemAP_PrimaryCurrency.Index).Value.ToString + "-" + rowDGV.Cells(Me.ColItemAP_PrimaryAkunAP.Index).Value.ToString)
                    SaldoOriginal = Me.GetSaldoOriginal_AP(rowDGV.Cells(Me.ColItemAP_PrimaryCurrency.Index).Value, rowDGV.Cells(Me.ColItemAP_PrimaryAkunAP.Index).Value.ToString)
                    Saldo = Me.GetSaldo_AP(rowDGV.Cells(Me.ColItemAP_PrimaryCurrency.Index).Value, rowDGV.Cells(Me.ColItemAP_PrimaryAkunAP.Index).Value.ToString)
                    Selisih = Me.GetSelisih_AP(rowDGV.Cells(Me.ColItemAP_PrimaryCurrency.Index).Value, rowDGV.Cells(Me.ColItemAP_PrimaryAkunAP.Index).Value.ToString)

                    If Saldo > 0 Then
                        Kredit = 0
                        Debet = Saldo
                        KreditOriginal = 0
                        DebetOriginal = SaldoOriginal
                        If SaldoOriginal = 0 Then
                            ExchangeRate = Math.Round(1, Core_Module.DefPrecExchangeRate)
                        Else
                            ExchangeRate = Math.Round(Saldo / SaldoOriginal, Core_Module.DefPrecExchangeRate)
                        End If
                    Else
                        Kredit = Math.Abs(Saldo)
                        Debet = 0
                        KreditOriginal = Math.Abs(SaldoOriginal)
                        DebetOriginal = 0
                        If SaldoOriginal = 0 Then
                            ExchangeRate = Math.Abs(Math.Round(1, Core_Module.DefPrecExchangeRate))
                        Else
                            ExchangeRate = Math.Abs(Math.Round(Saldo / SaldoOriginal, Core_Module.DefPrecExchangeRate))
                        End If
                    End If

                    Me.AddRowItem(rowDGV.Cells(Me.ColItemAP_PrimaryAkunAP.Index).Value, ExchangeRate, _
                    rowDGV.Cells(Me.ColItemAP_PrimaryCurrency.Index).Value, Debet, Kredit, DebetOriginal, KreditOriginal)
                    rowDGV1 = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    rowDGV1.Cells(Me.ColRemarks.Index).Value = "revaluate AP: " + _
                    Me.ObjCur.GetCurrencyCode(rowDGV.Cells(Me.ColItemAP_PrimaryCurrency.Index).Value)

                    'buat item jurnal untuk mengembalikan nilai balance ke nilai original semula
                    If Saldo > 0 Then
                        DebetOriginal = 0
                        KreditOriginal = SaldoOriginal
                        ExchangeRate = Me.GetExchangeRate(rowDGV.Cells(Me.ColItemAP_PrimaryCurrency.Index).Value)
                        Debet = 0
                        Kredit = Saldo + Selisih
                    Else
                        Kredit = 0
                        Debet = Math.Abs(Saldo + Selisih)
                        KreditOriginal = 0
                        DebetOriginal = Math.Abs(SaldoOriginal)
                        ExchangeRate = Me.GetExchangeRate(rowDGV.Cells(Me.ColItemAP_PrimaryCurrency.Index).Value)
                    End If
                    Me.AddRowItem(rowDGV.Cells(Me.ColItemAP_PrimaryAkunAP.Index).Value, ExchangeRate, _
                    rowDGV.Cells(Me.ColItemAP_PrimaryCurrency.Index).Value, Debet, Kredit, DebetOriginal, KreditOriginal)

                    'buat jurnal item selisih exchange rate
                    If Selisih <> 0 Then
                        If Selisih > 0 Then
                            Debet = Selisih
                            Kredit = 0
                            DebetOriginal = Debet
                            KreditOriginal = Kredit
                        Else
                            Debet = 0
                            Kredit = Math.Abs(Selisih)
                            DebetOriginal = Debet
                            KreditOriginal = Kredit
                        End If
                        ExchangeRate = 1
                        Me.AddRowItem(Me.ObjCur.GetPrimary_Akun_Realize_Exch_Rate(rowDGV.Cells(Me.ColItemAP_PrimaryCurrency.Index).Value), ExchangeRate, _
                        Core_Module.PrimaryBaseCurrency, Debet, Kredit, DebetOriginal, KreditOriginal)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ListKombinasi = Nothing
            DTAItem.Dispose()
            TBLItem.Dispose()
        End Try
    End Sub

    Private Sub AddRowItem(ByVal vPrimaryAkun As Double, ByVal vExchangeRate As Decimal, ByVal vPrimaryCurrency As Double, ByVal vDebet As Decimal, _
  ByVal vCredit As Decimal, ByVal vDebetOriginal As Decimal, ByVal vCreditOriginal As Decimal)
        Dim rowDGV As DataGridViewRow
        Dim DefPrec As Integer = 0

        Try
            DefPrec = Me.ObjCur.GetDefaultPrecission(vPrimaryCurrency)
            Me.dgvItem.Rows.Add()
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
            With rowDGV
                .Cells(Me.ColPrimaryAkun.Index).Value = vPrimaryAkun
                .Cells(Me.ColAccount.Index).Value = Me.ObjAkun.GetAccountCode(vPrimaryAkun)
                .Cells(Me.ColDescAccount.Index).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(vPrimaryAkun)
                .Cells(Me.ColExchangeRate.Index).Value = Me.ObjCD.DigitGrouping(vExchangeRate, Core_Module.DefPrecExchangeRate)
                .Cells(Me.ColRemarks.Index).Value = ""
                .Cells(Me.ColPrimaryCurrency.Index).Value = vPrimaryCurrency
                .Cells(Me.ColOriginalCurrency.Index).Value = Me.ObjCur.GetCurrencyCode(vPrimaryCurrency)

                .Cells(Me.ColDebet.Index).Value = Me.ObjCD.DigitGrouping(vDebet, Me.ObjCur.GetDecimal_DefaultCurrency)
                .Cells(Me.ColKredit.Index).Value = Me.ObjCD.DigitGrouping(vCredit, Me.ObjCur.GetDecimal_DefaultCurrency)
                .Cells(Me.ColDebet_OriginalCurr.Index).Value = Me.ObjCD.DigitGrouping(vDebetOriginal, DefPrec)
                .Cells(Me.ColCredit_OriginalCurrency.Index).Value = Me.ObjCD.DigitGrouping(vCreditOriginal, DefPrec)
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetSelisihAR(ByVal vPrimaryCurrency As Double, ByVal vPrimaryAkunAR As Double) As Decimal
        Dim rowDGV As DataGridViewRow
        Dim RetVal As Decimal = 0

        Try
            For Each rowDGV In Me.dgvItemAR.Rows
                If rowDGV.Cells(Me.ColItemAR_PrimaryCurrency.Index).Value = vPrimaryCurrency And _
                rowDGV.Cells(Me.ColItemAR_PrimaryAkunAR.Index).Value = vPrimaryAkunAR Then
                    RetVal += CDec(rowDGV.Cells(Me.ColItemAR_Selisih.Index).Value)
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetSelisih_AP(ByVal vPrimaryCurrency As Double, ByVal vPrimaryAkunAP As Double) As Decimal
        Dim rowDGV As DataGridViewRow
        Dim RetVal As Decimal = 0

        Try
            For Each rowDGV In Me.dgvItemAP.Rows
                If rowDGV.Cells(Me.ColItemAP_PrimaryCurrency.Index).Value = vPrimaryCurrency And _
                rowDGV.Cells(Me.ColItemAP_PrimaryAkunAP.Index).Value = vPrimaryAkunAP Then
                    RetVal += CDec(rowDGV.Cells(Me.ColItemAP_Selisih.Index).Value)
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetSaldo(ByVal vPrimaryCurrency As Double, ByVal vPrimaryAkunAR As Double) As Decimal
        Dim rowDGV As DataGridViewRow
        Dim RetVal As Decimal = 0

        Try
            For Each rowDGV In Me.dgvItemAR.Rows
                If rowDGV.Cells(Me.ColItemAR_PrimaryCurrency.Index).Value = vPrimaryCurrency And _
                rowDGV.Cells(Me.ColItemAR_PrimaryAkunAR.Index).Value = vPrimaryAkunAR Then
                    RetVal += CDec(rowDGV.Cells(Me.ColItemAR_Saldo.Index).Value)
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetSaldo_AP(ByVal vPrimaryCurrency As Double, ByVal vPrimaryAkunAP As Double) As Decimal
        Dim rowDGV As DataGridViewRow
        Dim RetVal As Decimal = 0

        Try
            For Each rowDGV In Me.dgvItemAP.Rows
                If rowDGV.Cells(Me.ColItemAP_PrimaryCurrency.Index).Value = vPrimaryCurrency And _
                rowDGV.Cells(Me.ColItemAP_PrimaryAkunAP.Index).Value = vPrimaryAkunAP Then
                    RetVal += CDec(rowDGV.Cells(Me.ColItemAP_Saldo.Index).Value)
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetSaldoOriginal(ByVal vPrimaryCurrency As Double, ByVal vPrimaryAkunAR As Double) As Decimal
        Dim rowDGV As DataGridViewRow
        Dim RetVal As Decimal = 0

        Try
            For Each rowDGV In Me.dgvItemAR.Rows
                If rowDGV.Cells(Me.ColItemAR_PrimaryCurrency.Index).Value = vPrimaryCurrency And _
                rowDGV.Cells(Me.ColItemAR_PrimaryAkunAR.Index).Value = vPrimaryAkunAR Then
                    RetVal += CDec(rowDGV.Cells(Me.ColItemAR_SaldoOriginal.Index).Value)
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetSaldoOriginal_AP(ByVal vPrimaryCurrency As Double, ByVal vPrimaryAkunAP As Double) As Decimal
        Dim rowDGV As DataGridViewRow
        Dim RetVal As Decimal = 0

        Try
            For Each rowDGV In Me.dgvItemAP.Rows
                If rowDGV.Cells(Me.ColItemAP_PrimaryCurrency.Index).Value = vPrimaryCurrency And _
                rowDGV.Cells(Me.ColItemAP_PrimaryAkunAP.Index).Value = vPrimaryAkunAP Then
                    RetVal += CDec(rowDGV.Cells(Me.ColItemAP_SaldoOriginal.Index).Value)
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub IsikanRevaluasiKasBank()
        Dim DTA As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltrans_jurnal_item1TableAdapter
        Dim TBL As New DseTransaksiJournal_EndOfMonth.tbltrans_jurnal_item1DataTable
        Dim row As DseTransaksiJournal_EndOfMonth.tbltrans_jurnal_item1Row
        Dim TipeAkun As ClsMasterAkun.EnumTipeAkun2 = ClsMasterAkun.EnumTipeAkun2.Cash_And_Bank
        Dim SaldoNormal As ClsMasterAkun.EnumSaldoNormal = ClsMasterAkun.EnumSaldoNormal.Debet
        Dim Total As Decimal = 0
        Dim DTA2 As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltrans_jurnal_item2TableAdapter
        Dim TBL2 As New DseTransaksiJournal_EndOfMonth.tbltrans_jurnal_item2DataTable
        Dim row2 As DseTransaksiJournal_EndOfMonth.tbltrans_jurnal_item2Row
        Dim rowDGV As DataGridViewRow
        Dim Debet, Credit, ExchangetRate, Selisih As Decimal
        Dim DefPrecBase As Integer = 0
        Dim PrimaryAkunSelisih As Double = -1
        Dim Ketemu As Boolean = True

        Try
            DefPrecBase = Me.ObjCur.GetDecimal_DefaultCurrency
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString

            'kumpulkan semua akun kas dan bank, lengkap dengan saldo akhirnya
            DTA.Fill(TBL, TipeAkun, Core_Module.PrimaryBaseCurrency, _
            Me.dtpDate.Value, #1/1/1950#)

            'hitung saldo selama bulan berjalan
            For Each row In TBL.Rows
                SaldoNormal = row.Saldo_Normal

                If SaldoNormal = ClsMasterAkun.EnumSaldoNormal.Debet Then
                    Total = row.Debet - row.Kredit + row.Beginning_Balance
                    If Total > 0 Then
                        row.Debet = Total
                        row.Kredit = 0
                    Else
                        row.Debet = 0
                        row.Kredit = Math.Abs(Total)
                    End If

                    Total = row.Debet_Original - row.Credit_Original + row.Beggining_Balance_Original_Cur
                    If Total > 0 Then
                        row.Debet_Original = Total
                        row.Credit_Original = 0
                    Else
                        row.Debet_Original = 0
                        row.Credit_Original = Math.Abs(Total)
                    End If
                Else
                    Total = row.Kredit - row.Debet + row.Beginning_Balance
                    If Total > 0 Then
                        row.Debet = 0
                        row.Kredit = Total
                    Else
                        row.Debet = Math.Abs(Total)
                        row.Kredit = 0
                    End If

                    Total = row.Credit_Original - row.Debet_Original + row.Beggining_Balance_Original_Cur
                    If Total > 0 Then
                        row.Debet_Original = 0
                        row.Credit_Original = Total
                    Else
                        row.Debet_Original = Math.Abs(Total)
                        row.Credit_Original = 0
                    End If
                End If
            Next

            'buang saldo karena jurnal revaluasi untuk bulan ini, khusus untuk edit
            If Me.Baru = False Then
                DTA2.Fill(TBL2, Me.TxtPrimaryJournal.Text)

                For Each row2 In TBL2.Rows
                    For Each row In TBL.Rows
                        If row.Primary_Akun = row2.Primary_Akun Then
                            row.Debet += row2.Kredit
                            row.Debet_Original += row2.Credit_Original
                            row.Kredit += row2.Debet
                            row.Credit_Original += row2.Debet_Original

                            If row.Debet <> 0 And row.Kredit <> 0 Then
                                Total = row.Debet - row.Kredit
                                If Total > 0 Then
                                    row.Debet = Total
                                    row.Kredit = 0
                                ElseIf Total < 0 Then
                                    row.Kredit = Math.Abs(Total)
                                    row.Debet = 0
                                Else
                                    row.Debet = 0
                                    row.Kredit = 0
                                End If

                                Total = row.Debet_Original - row.Credit_Original
                                If Total > 0 Then
                                    row.Debet_Original = Total
                                    row.Credit_Original = 0
                                ElseIf Total < 0 Then
                                    row.Credit_Original = Math.Abs(Total)
                                    row.Debet_Original = 0
                                Else
                                    row.Credit_Original = 0
                                    row.Debet_Original = 0
                                End If
                            End If

                            Exit For
                        End If
                    Next
                Next
            End If

            'masukkan item saldo akhir, hitung exhange rate
            For Each row In TBL.Rows
                If row.Debet_Original > 0 Or row.Credit_Original > 0 Then
                    'buat item berdasarkan saldo akhir, posisi debet kredit dibalik
                    Me.dgvItem.Rows.Add()
                    rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)

                    With rowDGV
                        .Cells(Me.ColAccount.Index).Value = row.Account_Code
                        Debet = row.Credit_Original
                        Credit = row.Debet_Original
                        .Cells(Me.ColCredit_OriginalCurrency.Index).Value = Me.ObjCD.DigitGrouping(Credit, row.Default_Precision)
                        .Cells(Me.ColDebet_OriginalCurr.Index).Value = Me.ObjCD.DigitGrouping(Debet, row.Default_Precision)
                        Debet = row.Kredit
                        Credit = row.Debet
                        .Cells(Me.ColKredit.Index).Value = Me.ObjCD.DigitGrouping(Credit, DefPrecBase)
                        .Cells(Me.ColDebet.Index).Value = Me.ObjCD.DigitGrouping(Debet, DefPrecBase)
                        .Cells(Me.ColDescAccount.Index).Value = row.Description
                        If row.Debet_Original > 0 Then
                            ExchangetRate = Math.Round(row.Debet / row.Debet_Original, Core_Module.DefPrecExchangeRate)
                        ElseIf row.Credit_Original > 0 Then
                            ExchangetRate = Math.Round(row.Kredit / row.Credit_Original, Core_Module.DefPrecExchangeRate)
                        Else
                            ExchangetRate = 1
                        End If

                        .Cells(Me.ColExchangeRate.Index).Value = Me.ObjCD.DigitGrouping(ExchangetRate, Core_Module.DefPrecExchangeRate)
                        .Cells(Me.ColOriginalCurrency.Index).Value = row.Currency_Code
                        .Cells(Me.ColPrimaryAkun.Index).Value = row.Primary_Akun
                        .Cells(Me.ColPrimaryCurrency.Index).Value = row.primary_currency
                        .Cells(Me.ColRemarks.Index).Value = "revaluate " + row.Currency_Code
                        .Cells(Me.ColShowDetail.Index).Value = ""
                    End With

                    'buat item jurnal balik dengan saldo original yang sama, dengan exchange rate yang terakhir
                    Me.dgvItem.Rows.Add()
                    rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    With rowDGV
                        .Cells(Me.ColAccount.Index).Value = row.Account_Code
                        .Cells(Me.ColCredit_OriginalCurrency.Index).Value = Me.ObjCD.DigitGrouping(row.Credit_Original, row.Default_Precision)
                        .Cells(Me.ColDebet_OriginalCurr.Index).Value = Me.ObjCD.DigitGrouping(row.Debet_Original, row.Default_Precision)
                        .Cells(Me.ColDescAccount.Index).Value = row.Description
                        ExchangetRate = Me.GetExchangeRate(row.primary_currency)
                        .Cells(Me.ColExchangeRate.Index).Value = Me.ObjCD.DigitGrouping(ExchangetRate, Core_Module.DefPrecExchangeRate)

                        Credit = Math.Round(ExchangetRate * row.Credit_Original, DefPrecBase)
                        Debet = Math.Round(ExchangetRate * row.Debet_Original, DefPrecBase)
                        .Cells(Me.ColKredit.Index).Value = Me.ObjCD.DigitGrouping(Credit, DefPrecBase)
                        .Cells(Me.ColDebet.Index).Value = Me.ObjCD.DigitGrouping(Debet, DefPrecBase)

                        .Cells(Me.ColOriginalCurrency.Index).Value = row.Currency_Code
                        .Cells(Me.ColPrimaryAkun.Index).Value = row.Primary_Akun
                        .Cells(Me.ColPrimaryCurrency.Index).Value = row.primary_currency
                        .Cells(Me.ColRemarks.Index).Value = ""
                        .Cells(Me.ColShowDetail.Index).Value = ""
                    End With

                    'buat realized exchange gain/loss
                    Selisih = Me.GetSelisih1(row.primary_currency, row.Primary_Akun)
                    If Selisih <> 0 Then
                        Me.dgvItem.Rows.Add()
                        rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                        With rowDGV
                            PrimaryAkunSelisih = Me.ObjCur.GetPrimary_Akun_Realize_Exch_Rate(row.primary_currency)
                            If PrimaryAkunSelisih = -1 Then
                                MsgBox("Logic error please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                                Exit Sub
                            End If
                            .Cells(Me.ColAccount.Index).Value = Me.ObjAkun.GetAccountCode(PrimaryAkunSelisih)
                            .Cells(Me.ColDescAccount.Index).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(PrimaryAkunSelisih)

                            .Cells(Me.ColExchangeRate.Index).Value = Me.ObjCD.DigitGrouping(1, Core_Module.DefPrecExchangeRate)
                            If Selisih > 0 Then
                                Credit = Selisih
                                Debet = 0
                            Else
                                Credit = 0
                                Debet = Math.Abs(Selisih)
                            End If
                            .Cells(Me.ColKredit.Index).Value = Me.ObjCD.DigitGrouping(Credit, DefPrecBase)
                            .Cells(Me.ColDebet.Index).Value = Me.ObjCD.DigitGrouping(Debet, DefPrecBase)
                            .Cells(Me.ColCredit_OriginalCurrency.Index).Value = .Cells(Me.ColKredit.Index).Value
                            .Cells(Me.ColDebet_OriginalCurr.Index).Value = .Cells(Me.ColDebet.Index).Value

                            .Cells(Me.ColOriginalCurrency.Index).Value = row.Currency_Code
                            .Cells(Me.ColPrimaryAkun.Index).Value = PrimaryAkunSelisih
                            .Cells(Me.ColPrimaryCurrency.Index).Value = Core_Module.PrimaryBaseCurrency
                            .Cells(Me.ColRemarks.Index).Value = ""
                            .Cells(Me.ColShowDetail.Index).Value = ""
                        End With
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTA2.Dispose()
            TBL2.Dispose()
        End Try
    End Sub

    Private Function GetSelisih1(ByVal vPrimaryCurrency As Double, ByVal vPrimaryAkun As Double) As Decimal
        Dim rowDGV As DataGridViewRow
        Dim Debet, Credit As Decimal

        Try
            Debet = 0
            Credit = 0
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value = vPrimaryCurrency And _
                rowDGV.Cells(Me.ColPrimaryAkun.Index).Value = vPrimaryAkun Then
                    Debet += CDec(rowDGV.Cells(Me.ColDebet.Index).Value)
                    Credit += CDec(rowDGV.Cells(Me.ColKredit.Index).Value)
                End If
            Next

            Return Debet - Credit
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetExchangeRate(ByVal vPrimaryCurrency As Double) As Decimal
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvExchangeRate.Rows
                If rowDGV.Cells(Me.ColItemExc_PrimaryCurrency.Index).Value = vPrimaryCurrency Then
                    Return rowDGV.Cells(Me.ColItemExc_ExchangeRate.Index).Value
                End If
            Next

            Return 1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            Me.dgvItem.AllowUserToAddRows = False
            If Me.TxtRefNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtRefNo.Text <> "" Then
                If Me.ObjJurnal.ApakahBolehEdit(Me.ObjJurnal.Primary, UserCOde, True, True) = True Then
                    Me.Baru = False
                    Me.ObjJurnal.CariDetail_EndOfMonth(Me.ObjJurnal.Primary, Me.dgvItem, Me, TipeJournal, Me.dgvExchangeRate, Me.dgvItemAR, Me.dgvItemAP)
                    Me.SalinkeField()
                    Me.AktifkanDetail()
                    'Me.dgvItem.AllowUserToAddRows = True
                    Me.TxtRefNo.Focus()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            Me.dgvItem.AllowUserToAddRows = False
            If Me.TxtRefNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtRefNo.Text <> "" Then
                If Me.ObjJurnal.ApakahBolehHapus(Me.TxtPrimaryJournal.Text, True, UserCOde, False, True) = False Then
                    Me.btnFind.Focus()
                ElseIf Me.ObjJurnal.ApakahBolehHapus_EndOFMonthJournal(Me.TxtPrimaryJournal.Text, True, UserCOde, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjJurnal.Hapus_EndOFMonthJournal(Me.TxtPrimaryJournal.Text, True, UserCOde, Me, Me.dgvExchangeRate, Me.dgvItemAR) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            Me.dgvItem.AllowUserToAddRows = False
            Me.SalinkeObjek()
            If Baru = True Then
                If ObjJurnal.ApakahDataBenar1(Baru, Me.dgvItem, Me.ColPrimaryCurrency.Index, Me.dtpDate.Value) = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If Me.ApakahDataBenar() = True Then
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjJurnal.SimpanBaru_EndOfMonth(Me.dgvItem, Me.UserCOde, Me, TipeJournal, Me.dgvExchangeRate, Me.dgvItemAR, Me.dgvItemAP) Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.NonAktifkanDetail()
                                Me.SalinkeField()
                                Me.btnFind.Focus()
                            End If
                            Me.dgvItem.AllowUserToAddRows = False
                        ElseIf hasil = MsgBoxResult.No Then
                            'Me.dgvItem.AllowUserToAddRows = True
                        End If
                    Else
                        'Me.dgvItem.AllowUserToAddRows = True
                    End If
                Else
                    'Me.dgvItem.AllowUserToAddRows = True
                End If
            Else
                'konfirmasi disimpan atau tidak
                If Me.ApakahDataBenar = True Then
                    If ObjJurnal.ApakahBolehEdit(Me.TxtPrimaryJournal.Text, UserCOde, True, True) = True Then
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjJurnal.SimpanPerubahan_EndOfMonth(Me.dgvItem, Me.UserCOde, Me, TipeJournal, Me.dgvExchangeRate, Me.dgvItemAR, Me.dgvItemAP) Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.NonAktifkanDetail()
                                Me.SalinkeField()
                                Me.btnEdit.Focus()
                            End If
                            Me.dgvItem.AllowUserToAddRows = False
                        ElseIf hasil = MsgBoxResult.No Then
                            'Me.dgvItem.AllowUserToAddRows = True
                        End If
                    End If
                Else
                    'Me.dgvItem.AllowUserToAddRows = True
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult

        Try
            Me.dgvItem.AllowUserToAddRows = False
            If Me.TxtRemarks.Text = "" And Me.dgvItem.RowCount = 0 And _
            Me.TxtRefNo.Text = "" Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    ''Me.KosongkanItem()
                    If Baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.ObjJurnal.CariDetail_EndOfMonth(Me.TxtPrimaryJournal.Text, Me.dgvItem, Me, TipeJournal, Me.dgvExchangeRate, Me.dgvItemAR, Me.dgvItemAP)
                        Me.SalinkeField()
                    End If
                Else
                    'Me.dgvItem.AllowUserToAddRows = True
                    Exit Sub
                End If
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPO_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Me.dgvItem.AllowUserToAddRows = False
            If Me.TxtRemarks.Enabled = True Then
                Me.BringToFront()
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinkeObjek()
                        If Baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjJurnal.SimpanBaru_EndOfMonth(Me.dgvItem, Me.UserCOde, Me, TipeJournal, Me.dgvExchangeRate, Me.dgvItemAR, Me.dgvItemAP) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinkeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjJurnal.SimpanPerubahan_EndOfMonth(Me.dgvItem, Me.UserCOde, Me, TipeJournal, Me.dgvExchangeRate, Me.dgvItemAR, Me.dgvItemAP) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinkeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            Else
                                e.Cancel = True
                            End If
                        End If
                    Else
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
            'Me.dgvItem.AllowUserToAddRows = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''Fungsi DGV
    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'tidak bisa isi dgv
            e.Cancel = True
            Exit Sub

            If e.ColumnIndex <> Me.ColAccount.Index Then
                If e.ColumnIndex <> Me.ColRemarks.Index And e.ColumnIndex <> Me.ColOriginalCurrency.Index And e.ColumnIndex <> Me.ColDescAccount.Index Then
                    If e.ColumnIndex = Me.ColDebet.Index Or e.ColumnIndex = Me.ColKredit.Index Or _
                    e.ColumnIndex = Me.ColDebet_OriginalCurr.Index Or e.ColumnIndex = Me.ColCredit_OriginalCurrency.Index Or _
                    e.ColumnIndex = Me.ColExchangeRate.Index Then
                        'untuk kolom exchange rate, debet original, creadit original tidak bisa diedit untuk akun dengan mata uang idr
                        If Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value = Me.PrimaryBaseCurrency Then
                            If e.ColumnIndex = Me.ColExchangeRate.Index Or _
                            e.ColumnIndex = Me.ColDebet_OriginalCurr.Index Or _
                            e.ColumnIndex = Me.ColCredit_OriginalCurrency.Index Then
                                e.Cancel = True
                                Exit Sub
                            End If
                        End If

                        If IsNothing(Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value) = True Then
                            'bukan baris teakhir
                            e.Cancel = True
                            Exit Sub
                        ElseIf Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader2(Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value) = False Then
                            'ada akunnya, aktif, dan bukan header
                            e.Cancel = True
                            Exit Sub
                        Else
                            Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                            'buang nol dibelakang decimal
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString)
                        End If
                    End If
                Else
                    If Me.dgvItem.Rows.Count - 1 = e.RowIndex Then
                        e.Cancel = True
                        Exit Sub
                    Else
                        If e.ColumnIndex = Me.ColRemarks.Index Then
                            Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        Else
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If
                End If
            Else
                'col account
                If e.ColumnIndex = Me.ColAccount.Index Then
                    Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim Rounding As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim ExRate As Decimal = 1
        Dim DefPrecOriginalCurr As Integer = 0

        Try
            Rounding = Me.DefaultPrecision
            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True _
            Or Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColRemarks.Index Then
                    'Remarks
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                ElseIf e.ColumnIndex = Me.ColAccount.Index Then
                    'kalo kode akun dihapus, maka hapus satu bari semuanya
                    If e.RowIndex <> Me.dgvItem.RowCount - 1 Then
                        Me.dgvItem.Rows.RemoveAt(e.RowIndex)
                    End If
                    Exit Sub
                Else
                    'untuk kolom debet kredit, langsung di balikin ke nilai semula
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                End If
            Else
                ''cek debet
                If e.ColumnIndex = Me.ColDebet.Index Then
                    If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, _
                            e.RowIndex).Value) = False Then
                        MsgBox("Please enter positive number", MsgBoxStyle.Information, _
                        "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, Rounding)
                    ''Cek kredit
                    If Me.dgvItem.Item(Me.ColKredit.Index, e.RowIndex).Value <> 0 Then
                        Me.dgvItem.Item(Me.ColKredit.Index, e.RowIndex).Value = ObjCD.DigitGrouping(0, Rounding)
                    End If
                    Me.BalanceDebetKredit(e.RowIndex, True, Rounding)
                    Me.HitungOriginalValues(e.RowIndex)
                End If
                ''Cek kredit 
                If e.ColumnIndex = Me.ColKredit.Index Then
                    If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, _
                            e.RowIndex).Value) = False Then
                        MsgBox("Please enter positive number", MsgBoxStyle.Information, _
                        "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, Rounding)
                    ''Cek debet
                    If Me.dgvItem.Item(Me.ColDebet.Index, e.RowIndex).Value <> 0 Then
                        Me.dgvItem.Item(Me.ColDebet.Index, e.RowIndex).Value = ObjCD.DigitGrouping(0, Rounding)
                    End If
                    Me.BalanceDebetKredit(e.RowIndex, False, Rounding)
                    Me.HitungOriginalValues(e.RowIndex)
                End If

                ''data account number
                If e.ColumnIndex = Me.ColAccount.Index Then
                    If e.ColumnIndex <> Me.ColDebet.Index Then
                        If e.ColumnIndex <> Me.ColKredit.Index Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                            'kemungkinan
                            If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                                Me.dgvItem.Item(Me.ColDescAccount.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value = Me.ObjAkun.GetPrimary(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value = Me.ObjAkun.GetPrimaryCurrency(Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value)
                                Me.dgvItem.Item(Me.ColOriginalCurrency.Index, e.RowIndex).Value = Me.ObjCur.GetCurrencyCode(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                                ExRate = Me.ObjTaxRate.GetTaxRate(Me.dtpDate.Value, Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                                If ExRate = 0 Then
                                    ExRate = 1
                                End If
                                Me.dgvItem.Item(Me.ColExchangeRate.Index, e.RowIndex).Value = ObjCD.DigitGrouping(ExRate, DefPrecExchangeRate)
                                'DefPrecOriginalCurr = Me.ObjCur.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                                Me.HitungOriginalValues(e.RowIndex)
                            Else
                                '2. kode akun yang diketik tidak ada --> tampilkan form cari
                                Dim ij As New FrmMasterAkun_Cari
                                ij.txtCari.Text = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                                ij.ShowDialog()
                                If ij.mPrimaryAkun <> -1 Then
                                    '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAkun.GetAccountCode(ij.mPrimaryAkun)
                                    Me.dgvItem.Item(Me.ColDescAccount.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                                    Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value = ij.mPrimaryAkun
                                    Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value = Me.ObjAkun.GetPrimaryCurrency(Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value)
                                    Me.dgvItem.Item(Me.ColOriginalCurrency.Index, e.RowIndex).Value = Me.ObjCur.GetCurrencyCode(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                                    ExRate = Me.ObjTaxRate.GetTaxRate(Me.dtpDate.Value, Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                                    If ExRate = 0 Then
                                        ExRate = 1
                                    End If
                                    Me.dgvItem.Item(Me.ColExchangeRate.Index, e.RowIndex).Value = ObjCD.DigitGrouping(ExRate, DefPrecExchangeRate)
                                    'DefPrecOriginalCurr = Me.ObjCur.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                                    Me.HitungOriginalValues(e.RowIndex)
                                ElseIf ij.mPrimaryAkun = -1 Then
                                    '2.b. kode akun yang dicari tidak ada --> kembalikan ke kondisi buffer
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                        Me.dgvItem.Rows.RemoveAt(e.RowIndex)
                                        Exit Sub
                                    End If
                                End If
                                ij = Nothing
                            End If
                        End If
                    End If
                    Me.SetMulticurrency()
                End If

                '3. exchange rate, hanya jika non base currency
                If e.ColumnIndex = Me.ColExchangeRate.Index Then
                    If Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value <> Me.PrimaryBaseCurrency Then
                        If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, _
                                                   e.RowIndex).Value) = False Then
                            MsgBox("Please enter positive number", MsgBoxStyle.Information, _
                            "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        'digit groping
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                        ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, DefPrecExchangeRate)
                        'hitung original value
                        Me.HitungOriginalValues(e.RowIndex)
                    End If
                End If

                '3. debet original currency
                If e.ColumnIndex = Me.ColDebet_OriginalCurr.Index Then
                    If Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value <> Me.PrimaryBaseCurrency Then
                        If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, _
                                                   e.RowIndex).Value) = False Then
                            MsgBox("Please enter positive number", MsgBoxStyle.Information, _
                            "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        'digit groping
                        DefPrecOriginalCurr = Me.ObjCur.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                        ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, _
                        DefPrecOriginalCurr)
                        'kosongkan credit
                        Me.dgvItem.Item(Me.ColCredit_OriginalCurrency.Index, e.RowIndex).Value = ObjCD.DigitGrouping(0, DefPrecOriginalCurr)

                        Me.HitungBaseValues(e.RowIndex)

                        If Me.dgvItem.Item(Me.ColDebet.Index, e.RowIndex).Value > 0 Then
                            Me.BalanceDebetKredit(e.RowIndex, True, Rounding, True)
                        Else
                            Me.BalanceDebetKredit(e.RowIndex, False, Rounding, True)
                        End If
                    End If
                End If

                '4. credit original currency
                If e.ColumnIndex = Me.ColCredit_OriginalCurrency.Index Then
                    If Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value <> Me.PrimaryBaseCurrency Then
                        If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, _
                                                   e.RowIndex).Value) = False Then
                            MsgBox("Please enter positive number", MsgBoxStyle.Information, _
                            "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        'digit groping
                        DefPrecOriginalCurr = Me.ObjCur.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                        ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, _
                        DefPrecOriginalCurr)
                        'kosongkan credit
                        Me.dgvItem.Item(Me.ColDebet_OriginalCurr.Index, e.RowIndex).Value = ObjCD.DigitGrouping(0, DefPrecOriginalCurr)

                        Me.HitungBaseValues(e.RowIndex)

                        If Me.dgvItem.Item(Me.ColDebet.Index, e.RowIndex).Value > 0 Then
                            Me.BalanceDebetKredit(e.RowIndex, True, Rounding, True)
                        Else
                            Me.BalanceDebetKredit(e.RowIndex, False, Rounding, True)
                        End If
                    End If
                End If

                '3 cek apakah ada kode akun yang sama pada referensi kode yang dibuat 
                'If Me.dgvItem.Rows.Count > 2 Then
                'Me.dgvItem.AllowUserToAddRows = False
                'Dim item As String = ""
                'For i As Integer = 0 To dgvItem.Rows.Count - 1
                'If e.RowIndex <> i Then
                'item = CStr(Me.dgvItem.Rows(i).Cells(Me.ColAccount.Index).Value)
                'If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = item Then
                'MsgBox("Duplication account number at data item, please change your number", _
                'MsgBoxStyle.Information, "Information")
                'Me.dgvItem.Rows.RemoveAt(e.RowIndex)
                'Me.dgvItem.AllowUserToAddRows = True
                'Exit Sub
                'End If
                'End If
                'Next
                'Me.dgvItem.AllowUserToAddRows = True
                'End If
            End If

            ''hitung debet dan kredit
            Me.HitungUnbalance()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCur = Nothing
        End Try
    End Sub

    Private Function HitungUnbalance(ByVal vdebet As Decimal, ByVal vkredit As Decimal) As Decimal
        Dim Unbalance As Decimal = 0
        Try
            Unbalance = Math.Round(vdebet - vkredit, Me.DefaultPrecision, MidpointRounding.AwayFromZero)
            Return Unbalance
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub llbEditQuickJournal_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbEditQuickJournal.LinkClicked
        Dim ij As New FrmTransaksiJournal_quickAdd

        Try
            ij.UserCOde = Me.UserCOde
            ij.BackcolorAll = Me.BackcolorAll
            ij.TipeJournal = Me.TipeJournal
            ij.ShowDialog()
            Me.ObjQuickJurnal.IsiCboQuickAdd(Me.TipeJournal, Me.cboQuickJournal)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij = Nothing
        End Try
    End Sub

    Private Sub HitungUnbalance()
        Dim row As DataGridViewRow
        Try
            ''hitung debet dan kredit
            Dim TotalDebet As Decimal = 0
            Dim TotalKredit As Decimal = 0
            For Each row In dgvItem.Rows
                TotalDebet += CDec(Me.dgvItem.Item(Me.ColDebet.Index, row.Index).Value)
                TotalKredit += CDec(Me.dgvItem.Item(Me.ColKredit.Index, row.Index).Value)
            Next
            Me.TxtUnbalance.Text = CStr(ObjCD.DigitGrouping(Me.HitungUnbalance(TotalDebet, TotalKredit), Me.DefaultPrecision))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BalanceDebetKredit(ByVal vIndexAcuan As Integer, ByVal vAcuanDebet As Boolean, ByVal rounding As Integer, _
   Optional ByVal vHitungOriginalCurrency As Boolean = True)

        Try
            Dim rowAcuan As DataGridViewRow = Me.dgvItem.Rows(vIndexAcuan)
            Dim rowTarget As DataGridViewRow

            If Me.TxtRemarks.Enabled = True And Me.dgvItem.Rows.Count = 3 Then
                'dibalance jika hanya ada 2 jurnal, maka debet = kredit
                If rowAcuan.Index = 0 Then
                    rowTarget = Me.dgvItem.Rows(1)
                Else
                    rowTarget = Me.dgvItem.Rows(0)
                End If

                If vAcuanDebet = True Then
                    rowTarget.Cells(Me.ColKredit.Index).Value = rowAcuan.Cells(Me.ColDebet.Index).Value
                    rowTarget.Cells(Me.ColDebet.Index).Value = Me.ObjCD.DigitGrouping(0, rounding)
                Else
                    rowTarget.Cells(Me.ColDebet.Index).Value = rowAcuan.Cells(Me.ColKredit.Index).Value
                    rowTarget.Cells(Me.ColKredit.Index).Value = Me.ObjCD.DigitGrouping(0, rounding)
                End If

                If vHitungOriginalCurrency = True Then
                    Me.HitungOriginalValues(rowTarget.Index)
                End If
            ElseIf Me.TxtRemarks.Enabled = True And Me.dgvItem.Rows.Count > 3 Then
                'dibalance jika ada 3 jurnal, jika 2 dikredit atau 2 didebet
                If Me.GetJmlNol = 0 Then
                    If vAcuanDebet = True Then
                        If Me.GetJmlKredit = 1 Then
                            rowTarget = Me.dgvItem.Rows(Me.GetIndexJmlKredit)
                            rowTarget.Cells(Me.ColKredit.Index).Value = Me.ObjCD.DigitGrouping(Me.GetTotalDebet, rounding)
                            rowTarget.Cells(Me.ColDebet.Index).Value = Me.ObjCD.DigitGrouping(0, rounding)
                        End If
                    Else
                        If Me.GetJmlDebet = 1 Then
                            rowTarget = Me.dgvItem.Rows(Me.GetIndexJmlDebet)
                            rowTarget.Cells(Me.ColDebet.Index).Value = Me.ObjCD.DigitGrouping(Me.GetTotalKredit, rounding)
                            rowTarget.Cells(Me.ColKredit.Index).Value = Me.ObjCD.DigitGrouping(0, rounding)
                        End If
                    End If

                    If vHitungOriginalCurrency = True Then
                        If IsNothing(rowTarget) = False Then
                            Me.HitungOriginalValues(rowTarget.Index)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetTotalDebet() As Decimal
        Dim row As DataGridViewRow
        Dim retval As Decimal = 0

        Try
            If Me.dgvItem.Rows.Count > 1 Then
                For i As Integer = 0 To Me.dgvItem.Rows.Count - 2
                    row = Me.dgvItem.Rows(i)
                    If row.Cells(Me.ColDebet.Index).Value <> 0 Then
                        retval += CDec(row.Cells(Me.ColDebet.Index).Value)
                    End If
                Next
            End If

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDebet() As Integer
        Dim row As DataGridViewRow
        Dim retval As Integer = 0

        Try
            If Me.dgvItem.Rows.Count > 1 Then
                For i As Integer = 0 To Me.dgvItem.Rows.Count - 2
                    row = Me.dgvItem.Rows(i)
                    If row.Cells(Me.ColDebet.Index).Value <> 0 Then
                        retval += 1
                    End If
                Next
            End If

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetIndexJmlDebet() As Integer
        Dim row As DataGridViewRow

        Try
            If Me.dgvItem.Rows.Count > 1 Then
                For i As Integer = 0 To Me.dgvItem.Rows.Count - 2
                    row = Me.dgvItem.Rows(i)
                    If row.Cells(Me.ColDebet.Index).Value <> 0 Then
                        Return row.Index
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetTotalKredit() As Decimal
        Dim row As DataGridViewRow
        Dim retval As Decimal = 0

        Try
            If Me.dgvItem.Rows.Count > 1 Then
                For i As Integer = 0 To Me.dgvItem.Rows.Count - 2
                    row = Me.dgvItem.Rows(i)
                    If row.Cells(Me.ColKredit.Index).Value <> 0 Then
                        retval += CDec(row.Cells(Me.ColKredit.Index).Value)
                    End If
                Next
            End If

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlKredit() As Integer
        Dim row As DataGridViewRow
        Dim retval As Integer = 0

        Try
            If Me.dgvItem.Rows.Count > 1 Then
                For i As Integer = 0 To Me.dgvItem.Rows.Count - 2
                    row = Me.dgvItem.Rows(i)
                    If row.Cells(Me.ColKredit.Index).Value <> 0 Then
                        retval += 1
                    End If
                Next
            End If

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetIndexJmlKredit() As Integer
        Dim row As DataGridViewRow

        Try
            If Me.dgvItem.Rows.Count > 1 Then
                For i As Integer = 0 To Me.dgvItem.Rows.Count - 2
                    row = Me.dgvItem.Rows(i)
                    If row.Cells(Me.ColKredit.Index).Value <> 0 Then
                        Return row.Index
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''' <summary>
    ''' fungsi ini mengembalikan jumlah baris yang nilai debet dan kreditnya =0
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetJmlNol() As Integer
        Dim row As DataGridViewRow
        Dim retval As Integer = 0

        Try
            If Me.dgvItem.Rows.Count > 1 Then
                For i As Integer = 0 To Me.dgvItem.Rows.Count - 2
                    row = Me.dgvItem.Rows(i)
                    If row.Cells(Me.ColDebet.Index).Value = 0 And row.Cells(Me.ColKredit.Index).Value = 0 Then
                        retval += 1
                    End If
                Next
            End If

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SetMulticurrency()
        Dim rowDGV As DataGridViewRow
        Dim Multicurrency As Boolean = False

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Index = Me.dgvItem.RowCount - 1 Then
                    If Me.dgvItem.AllowUserToAddRows = True Then
                        'tidak dicek
                    Else
                        'dicek
                        If rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value <> Me.PrimaryBaseCurrency Then
                            Multicurrency = True
                            Exit For
                        End If
                    End If
                Else
                    If rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value <> Me.PrimaryBaseCurrency Then
                        Multicurrency = True
                        Exit For
                    End If
                End If
            Next

            If Multicurrency = True Then
                Me.ColExchangeRate.Visible = True
                Me.ColDebet_OriginalCurr.Visible = True
                Me.ColCredit_OriginalCurrency.Visible = True
                Me.ColOriginalCurrency.Visible = True
            Else
                Me.ColOriginalCurrency.Visible = False
                Me.ColExchangeRate.Visible = False
                Me.ColDebet_OriginalCurr.Visible = False
                Me.ColCredit_OriginalCurrency.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Try
            If Me.TxtRefNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtRefNo.Text <> "" Then
                Dim i As New FrmTransaksiJournal_Cetak1
                i.PrimaryJurnal = Me.TxtPrimaryJournal.Text
                i.ShowDialog()
                i.Dispose()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungExchangeRate(ByVal vRowIndex As Integer)
        Dim vRowDGV As DataGridViewRow

        Try
            vRowDGV = Me.dgvItem.Rows(vRowIndex)
            vRowDGV.Cells(Me.ColExchangeRate.Index).Value = ObjCD.DigitGrouping(Math.Round(CDec(vRowDGV.Cells(Me.ColDebet.Index).Value) / _
            CDec(vRowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value), DefPrecExchangeRate, MidpointRounding.AwayFromZero))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungBaseValues(ByVal vRowIndex As Integer)
        Dim vRowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0
        Dim DefPrecOriginal As Integer = 0

        Try
            vRowDGV = Me.dgvItem.Rows(vRowIndex)
            DefPrecBase = Me.ObjCur.GetDefaultPrecission(Me.PrimaryBaseCurrency)
            DefPrecOriginal = Me.ObjCur.GetDefaultPrecission(vRowDGV.Cells(Me.ColPrimaryCurrency.Index).Value)

            If CDec(vRowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value) > 0 Then
                vRowDGV.Cells(Me.ColKredit.Index).Value = ObjCD.DigitGrouping(0, DefPrecBase)
                vRowDGV.Cells(Me.ColDebet.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value) * _
                CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value), DefPrecBase)
            Else
                vRowDGV.Cells(Me.ColKredit.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColCredit_OriginalCurrency.Index).Value) * _
                CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value), DefPrecBase)
                vRowDGV.Cells(Me.ColDebet.Index).Value = ObjCD.DigitGrouping(0, DefPrecBase)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungOriginalValues(ByVal vRowIndex As Integer)
        Dim vRowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0
        Dim DefPrecOriginal As Integer = 0

        Try
            vRowDGV = Me.dgvItem.Rows(vRowIndex)
            DefPrecBase = Me.ObjCur.GetDefaultPrecission(Me.PrimaryBaseCurrency)
            DefPrecOriginal = Me.ObjCur.GetDefaultPrecission(vRowDGV.Cells(Me.ColPrimaryCurrency.Index).Value)

            If CDec(vRowDGV.Cells(Me.ColDebet.Index).Value) > 0 Then
                vRowDGV.Cells(Me.ColCredit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(0, DefPrecOriginal)
                vRowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColDebet.Index).Value) / _
                CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value), DefPrecOriginal)
            Else
                vRowDGV.Cells(Me.ColCredit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColKredit.Index).Value) / _
                CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value), DefPrecOriginal)
                vRowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(0, DefPrecOriginal)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvExchangeRate_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvExchangeRate.CellBeginEdit
        Try
            'yang bisa diedit hanya exchange rate saja
            If e.ColumnIndex = Me.ColItemExc_ExchangeRate.Index Or e.ColumnIndex = Me.ColItemExc_Remarks.Index Then
                Me.Buffer = Me.dgvExchangeRate.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvExchangeRate_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvExchangeRate.CellEndEdit
        Try
            'pastikan ada isinya
            If IsNothing(Me.dgvExchangeRate.Item(e.ColumnIndex, e.RowIndex).Value) = True _
            Or Me.dgvExchangeRate.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColItemExc_Remarks.Index Then
                    Me.dgvExchangeRate.Item(e.ColumnIndex, e.RowIndex).Value = ""
                Else
                    Me.dgvExchangeRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                End If
                Exit Sub
            Else
                If e.ColumnIndex = Me.ColItemExc_ExchangeRate.Index Then
                    'pastikan positif
                    If Me.ObjCD.CekBilanganPositif(Me.dgvExchangeRate.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.dgvExchangeRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    'digit grouping
                    Me.dgvExchangeRate.Item(e.ColumnIndex, e.RowIndex).Value = _
                    Me.ObjCD.DigitGrouping(CDec(Me.dgvExchangeRate.Item(e.ColumnIndex, e.RowIndex).Value), Core_Module.DefPrecExchangeRate)

                    If CDec(Me.dgvExchangeRate.Item(e.ColumnIndex, e.RowIndex).Value) <> CDec(Me.Buffer) Then
                        Me.IsikanRevaluasi()
                    End If
                ElseIf e.ColumnIndex = Me.ColItemExc_Remarks.Index Then
                    'maximal lengthnya 250
                    If CStr(Me.dgvExchangeRate.Item(e.ColumnIndex, e.RowIndex).Value.ToString).Length > 250 Then
                        Me.dgvExchangeRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvExchangeRate.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Try
            'tidak bisa dihapuskan secara manual
            e.Cancel = True
            Exit Sub
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class