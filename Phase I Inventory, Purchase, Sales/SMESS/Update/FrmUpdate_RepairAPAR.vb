Option Explicit On
Option Strict Off

Public Class FrmUpdate_RepairAPAR
    Dim DTASupplier As New DseUpdate_RepairKartuAPARTableAdapters.tblmaster_supplierTableAdapter
    Dim TBLSupplier As New DseUpdate_RepairKartuAPAR.tblmaster_supplierDataTable
    Dim DTACust As New DseUpdate_RepairKartuAPARTableAdapters.tblmaster_customerTableAdapter
    Dim TBLCust As New DseUpdate_RepairKartuAPAR.tblmaster_customerDataTable
    Dim DTABalancePurchase As New DseUpdate_RepairKartuAPARTableAdapters.tbltrans_balance_purchaseTableAdapter
    Dim TBLBalancePurchase As New DseUpdate_RepairKartuAPAR.tbltrans_balance_purchaseDataTable
    Dim DTABalanceSales As New DseUpdate_RepairKartuAPARTableAdapters.tbltrans_balance_salesTableAdapter
    Dim TBLBalanceSales As New DseUpdate_RepairKartuAPAR.tbltrans_balance_salesDataTable
    Dim DTAKartuAPAR As New DseUpdate_RepairKartuAPARTableAdapters.tbltransaksi_kartu_ar_apTableAdapter
    Dim TBLKartuAPAR As New DseUpdate_RepairKartuAPAR.tbltransaksi_kartu_ar_apDataTable
    Dim TBLRef As New DseUpdate_RepairKartuAPAR.TblRef_DataTable
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim DefPrec As Integer = 0
    Dim ObjCD As New Core_Module.ClsCekData
    Dim TipeKartu As EnumJenisKartu = EnumJenisKartu.AP

    Public Enum EnumJenisKartu As Integer
        AR
        AP
    End Enum

    Public Enum EnumTipeJournal As Integer
        GeneralJournal
        AdjustmenJournal
        ClosingJournal
        SalesJournal
        SalesReturnJournal
        SalesDownPayment
        PurchaseJournal
        PurchaseReturnJournal
        PurchaseDownPaymentJournal
        SalesReceivementJournal
        SalesReturnPaymentJournal
        PurchasePaymentJournal
        PurchaseReturnReceivement
    End Enum

    Public Enum EnumTipeJurnal2 As Integer
        Jurnal_NoGrouping
        GeneralJurnal_MonthlyRevaluation
        Cash_n_Bank
        AssetAcquistion_Other
        AssetRevaluation_Other
        AssetDepreciation
        AssetDisposal
        AssetSetFinish
        HR_Payroll_Payment
        HR_Payroll_Loan
        HR_Payroll_Loan_Payment
        GeneralJurnal_Recurring
        Credit_Note
        Debit_Note
    End Enum

    Private Function ApakahAdaMonthlyRevJurnal() As Boolean
        Dim DTA As New DseUpdate_RepairKartuAPARTableAdapters.tbltrans_jurnalTableAdapter
        Dim TBL As New DseUpdate_RepairKartuAPAR.tbltrans_jurnalDataTable
        Dim row As DseUpdate_RepairKartuAPAR.tbltrans_jurnalRow
        Dim TipeJurnal As EnumTipeJournal = EnumTipeJournal.GeneralJournal
        Dim TipeJurnal2 As EnumTipeJurnal2 = EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation

        Try
            DTA.Fill(TBL, TipeJurnal, TipeJurnal2)

            If TBL.Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Private Sub btnRepair_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRepair.Click
        Try
            If Me.ApakahAdaMonthlyRevJurnal = True Then
                MsgBox("Please delete monthly revaluation journal first. The function in this form can't work perfectly if there's monthly rev journal.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.dgvRepair.RowCount = 0 Then
                MsgBox("Please specify at least 1 data to repair.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'coding disini
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCek_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCek.Click
        

        Try
            TBLRef.Rows.Clear()

            If Me.ApakahAdaMonthlyRevJurnal = True Then
                MsgBox("Please delete monthly revaluation journal first. The function in this form can't work perfectly if there's monthly rev journal.", _
                MsgBoxStyle.Information, "Information")

                Exit Sub
            End If

            Control.CheckForIllegalCrossThreadCalls = False
            Me.bgw1.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Cek_AR(ByVal TipeKartu As EnumJenisKartu, ByVal TBLRef As DseUpdate_RepairKartuAPAR.TblRef_DataTable)
        Dim rowRef As DseUpdate_RepairKartuAPAR.TblRef_Row
        Dim counter As Double = 0
        Dim rowKartu As DseUpdate_RepairKartuAPAR.tbltransaksi_kartu_ar_apRow
        Dim rowBalanceSales As DseUpdate_RepairKartuAPAR.tbltrans_balance_salesRow
        Dim JmlKartu, JmlBalance As Decimal

        Try
            'isikan data balance
            Me.DTABalanceSales.Fill(Me.TBLBalanceSales)
            'isikan data kartu
            Me.DTAKartuAPAR.Fill(TBLKartuAPAR, TipeKartu)

            'compare data kartu vs data balance, jika tidak sama maka masukkan ke dgv
            counter = 0
            Me.prg1.Value = 0
            Me.prg1.Maximum = Me.TBLRef.Rows.Count

            While counter <= TBLRef.Rows.Count - 1
                JmlKartu = 0
                JmlBalance = 0

                rowRef = TBLRef.Rows(counter)
                Me.prg1.Value += 1
                Me.LblProgress.Text = "Checking data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString + " - " + rowRef.Name

                If TBLBalanceSales.Rows.Count > 0 Then
                    rowBalanceSales = TBLBalanceSales.Rows(0)
                    While rowBalanceSales.Primary_Customer < rowRef.Primary
                        TBLBalanceSales.Rows.Remove(rowBalanceSales)
                        If TBLBalanceSales.Rows.Count = 0 Then
                            Exit While
                        Else
                            rowBalanceSales = TBLBalanceSales.Rows(0)
                        End If
                    End While

                    If TBLBalanceSales.Rows.Count = 0 Then
                        JmlBalance = 0
                    Else
                        rowBalanceSales = TBLBalanceSales.Rows(0)
                        If rowBalanceSales.Primary_Customer = rowRef.Primary Then
                            JmlBalance = rowBalanceSales.Balance
                        Else
                            JmlBalance = 0
                        End If
                    End If
                Else
                    JmlBalance = 0
                End If
                If TBLKartuAPAR.Rows.Count > 0 Then
                    rowKartu = TBLKartuAPAR.Rows(0)
                    While rowKartu.Primary_Pihak_Acuan < rowRef.Primary
                        TBLKartuAPAR.Rows.Remove(rowKartu)
                        If TBLKartuAPAR.Rows.Count = 0 Then
                            Exit While
                        Else
                            rowKartu = TBLKartuAPAR.Rows(0)
                        End If
                    End While

                    If TBLKartuAPAR.Rows.Count > 0 Then
                        rowKartu = TBLKartuAPAR.Rows(0)
                        If rowKartu.Primary_Pihak_Acuan = rowRef.Primary Then
                            JmlKartu = rowKartu.Debet - rowKartu.Kredit
                        Else
                            JmlKartu = 0
                        End If
                    Else
                        JmlKartu = 0
                    End If
                Else
                    JmlKartu = 0
                End If

                If JmlKartu = JmlBalance Then
                    TBLRef.Rows.Remove(rowRef)
                Else
                    counter += 1
                    rowRef.Jml_Balance = JmlBalance
                    rowRef.Jml_Kartu = JmlKartu
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Cek_AP(ByVal TipeKartu As EnumJenisKartu, ByVal TBLRef As DseUpdate_RepairKartuAPAR.TblRef_DataTable)
        Dim rowRef As DseUpdate_RepairKartuAPAR.TblRef_Row
        Dim counter As Double = 0
        Dim rowKartu As DseUpdate_RepairKartuAPAR.tbltransaksi_kartu_ar_apRow
        Dim rowBalancePurchase As DseUpdate_RepairKartuAPAR.tbltrans_balance_purchaseRow
        Dim rowBalanceSales As DseUpdate_RepairKartuAPAR.tbltrans_balance_salesRow
        Dim JmlKartu, JmlBalance As Decimal

        Try
            'isikan data balance
            Me.DTABalancePurchase.Fill(TBLBalancePurchase)
            'isikan data kartu
            Me.DTAKartuAPAR.Fill(TBLKartuAPAR, TipeKartu)

            'compare data kartu vs data balance, jika tidak sama maka masukkan ke dgv
            counter = 0
            Me.prg1.Value = 0
            Me.prg1.Maximum = Me.TBLRef.Rows.Count

            While counter <= TBLRef.Rows.Count - 1
                JmlKartu = 0
                JmlBalance = 0

                rowRef = TBLRef.Rows(counter)
                Me.prg1.Value += 1
                Me.LblProgress.Text = "Checking data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString + " - " + rowRef.Name

                If TBLBalancePurchase.Rows.Count > 0 Then
                    rowBalancePurchase = TBLBalancePurchase.Rows(0)
                    While rowBalancePurchase.Primary_Supplier < rowRef.Primary
                        TBLBalancePurchase.Rows.Remove(rowBalancePurchase)
                        If TBLBalancePurchase.Rows.Count = 0 Then
                            Exit While
                        Else
                            rowBalancePurchase = TBLBalancePurchase.Rows(0)
                        End If
                    End While

                    If TBLBalancePurchase.Rows.Count = 0 Then
                        JmlBalance = 0
                    Else
                        rowBalancePurchase = TBLBalancePurchase.Rows(0)
                        If rowBalancePurchase.Primary_Supplier = rowRef.Primary Then
                            JmlBalance = rowBalancePurchase.Balance
                        Else
                            JmlBalance = 0
                        End If
                    End If
                Else
                    JmlBalance = 0
                End If
                If TBLKartuAPAR.Rows.Count > 0 Then
                    rowKartu = TBLKartuAPAR.Rows(0)
                    While rowKartu.Primary_Pihak_Acuan < rowRef.Primary
                        TBLKartuAPAR.Rows.Remove(rowKartu)
                        If TBLKartuAPAR.Rows.Count = 0 Then
                            Exit While
                        Else
                            rowKartu = TBLKartuAPAR.Rows(0)
                        End If
                    End While

                    If TBLKartuAPAR.Rows.Count > 0 Then
                        rowKartu = TBLKartuAPAR.Rows(0)
                        If rowKartu.Primary_Pihak_Acuan = rowRef.Primary Then
                            JmlKartu = rowKartu.Kredit - rowKartu.Debet
                        Else
                            JmlKartu = 0
                        End If
                    Else
                        JmlKartu = 0
                    End If
                Else
                    JmlKartu = 0
                End If

                If JmlKartu = JmlBalance Then
                    TBLRef.Rows.Remove(rowRef)
                Else
                    counter += 1
                    rowRef.Jml_Balance = JmlBalance
                    rowRef.Jml_Kartu = JmlKartu
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiTBLRef(ByVal vTBLRef As DseUpdate_RepairKartuAPAR.TblRef_DataTable, ByVal vTipeKartu As EnumJenisKartu)
        Dim rowRef As DseUpdate_RepairKartuAPAR.TblRef_Row
        Dim rowSupp As DseUpdate_RepairKartuAPAR.tblmaster_supplierRow
        Dim rowCust As DseUpdate_RepairKartuAPAR.tblmaster_customerRow

        Try
            Me.TBLCust.Rows.Clear()
            Me.TBLSupplier.Rows.Clear()
            vTBLRef.Rows.Clear()

            If vTipeKartu = EnumJenisKartu.AP Then
                Me.DTASupplier.Fill(TBLSupplier)

                For Each rowSupp In TBLSupplier.Rows
                    rowRef = vTBLRef.NewTblRef_Row
                    rowRef.Primary = rowSupp.Primary
                    rowRef.Code = rowSupp.Supplier_Code
                    rowRef.Name = rowSupp.Supplier_Name
                    vTBLRef.Rows.Add(rowRef)
                Next
            Else
                Me.DTACust.Fill(TBLCust)

                For Each rowCust In TBLCust.Rows
                    rowRef = vTBLRef.NewTblRef_Row
                    rowRef.Primary = rowCust.Primary
                    rowRef.Code = rowCust.Customer_Code
                    rowRef.Name = rowCust.Customer_Name
                    vTBLRef.Rows.Add(rowRef)
                Next
            End If

            Me.TBLCust.Rows.Clear()
            Me.TBLSupplier.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmUpdate_RepairAPAR_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.DTACust.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTASupplier.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTABalancePurchase.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTABalanceSales.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTAKartuAPAR.Connection.ConnectionString = Core_Module.GetConstring
            DefPrec = Me.ObjCur.GetDecimal_DefaultCurrency
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw1.DoWork
        Try
            Me.ObjCD.UpdateThreading()

            If Me.rboAP.Checked = True Then
                TipeKartu = EnumJenisKartu.AP
            Else
                TipeKartu = EnumJenisKartu.AR
            End If
            Me.IsiTBLRef(TBLRef, TipeKartu)

            Me.TBLBalancePurchase.Rows.Clear()
            Me.TBLBalanceSales.Rows.Clear()
            Me.TBLKartuAPAR.Rows.Clear()
            If TipeKartu = EnumJenisKartu.AP Then
                Cek_AP(TipeKartu, TBLRef)
            Else
                Cek_AR(TipeKartu, TBLRef)
            End If
            Me.TBLBalancePurchase.Rows.Clear()
            Me.TBLBalanceSales.Rows.Clear()
            Me.TBLKartuAPAR.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw1.RunWorkerCompleted
        Dim rowRef As DseUpdate_RepairKartuAPAR.TblRef_Row
        Dim rowDGV As DataGridViewRow

        Try
            'pindahakan ke dgv repair jika ada sisa di tbl ref
            Me.dgvRepair.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            If TBLRef.Rows.Count > 0 Then
                TBLRef.DefaultView.Sort = "Code"
                For Each DV As DataRowView In TBLRef.DefaultView
                    rowRef = DV.Row
                    Me.dgvRepair.Rows.Add()
                    rowDGV = Me.dgvRepair.Rows(Me.dgvRepair.Rows.Count - 1)
                    rowDGV.Cells(Me.ColPrimaryRefParty.Index).Value = rowRef.Primary
                    rowDGV.Cells(Me.ColRefParty.Index).Value = rowRef.Code + " - " + rowRef.Name
                    rowDGV.Cells(Me.ColTotalBalance.Index).Value = Me.ObjCD.DigitGrouping(rowRef.Jml_Balance, Me.DefPrec)
                    rowDGV.Cells(Me.ColTotalKartu.Index).Value = Me.ObjCD.DigitGrouping(rowRef.Jml_Kartu, Me.DefPrec)
                    rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvRepair.Rows.Count
                    rowDGV.Cells(Me.ColSelisih.Index).Value = Me.ObjCD.DigitGrouping(rowRef.Jml_Balance - rowRef.Jml_Kartu, Me.DefPrec)
                Next
            End If
            Me.dgvRepair.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Dim ij As New Core_Module.FrmCekKartuAPAR
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class