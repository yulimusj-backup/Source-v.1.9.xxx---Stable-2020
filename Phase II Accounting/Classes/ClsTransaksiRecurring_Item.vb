Option Explicit On
Option Strict Off

Public Class ClsTransaksiRecurring_Item
    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrin As New Core_Module.ClsAntrian

    Public Function IsiDGVItem(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
                               ByVal frm As FrmTransaksiRecuring) As Boolean
        Dim DTA As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journal_account_entriesTableAdapter
        Dim TBL As New DseTransaksiRecurring.tbltrans_recurring_journal_account_entriesDataTable
        Dim row As DseTransaksiRecurring.tbltrans_recurring_journal_account_entriesRow

        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim ObjAcc As New ClsMasterAkun
        Dim objDept As New Core_Module.ClsMasterDepartement
        Dim objJob As New Core_Module.ClsMasterJob

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.Rows.Clear()

            'isikan tabel
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVItem.Rows.Add()
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(frm.ColPrimaryAkun.Index).Value = .Primary_Akun
                        rowDGV.Cells(frm.ColAccount.Index).Value = ObjAcc.GetAccountCode(.Primary_Akun)
                        rowDGV.Cells(frm.ColDescAccount.Index).Value = ObjAcc.GetDescription(rowDGV.Cells(frm.ColAccount.Index).Value.ToString)
                        rowDGV.Cells(frm.ColDebet.Index).Value = ObjCD.DigitGrouping(.Debet.ToString, DefPrecBase)
                        rowDGV.Cells(frm.ColKredit.Index).Value = ObjCD.DigitGrouping(.Kredit.ToString, DefPrecBase)
                        rowDGV.Cells(frm.ColRemarks.Index).Value = .Remarks
                        rowDGV.Cells(frm.ColPrimaryCurrency.Index).Value = ObjAcc.GetPrimaryCurrency(.Primary_Akun)
                        rowDGV.Cells(frm.ColOriginalCurrency.Index).Value = ObjCur.GetCurrencyCode(rowDGV.Cells(frm.ColPrimaryCurrency.Index).Value)
                        Dim DeffPrec As Decimal = 0
                        DeffPrec = ObjCur.GetDefaultPrecission(rowDGV.Cells(frm.ColPrimaryCurrency.Index).Value)
                        rowDGV.Cells(frm.ColDebet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(.Debet_Original, DeffPrec)
                        rowDGV.Cells(frm.ColCredit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(.Credit_Original, DeffPrec)
                        rowDGV.Cells(frm.ColExchangeRate.Index).Value = ObjCD.DigitGrouping(.Ratio.ToString, Core_Module.DefPrecExchangeRate)
                        rowDGV.Cells(frm.ColDept.Index).Value = objDept.GetDeptCode(.primary_dept)
                        rowDGV.Cells(frm.ColPrimaryDept.Index).Value = .primary_dept
                        rowDGV.Cells(frm.ColPrimaryJob.Index).Value = .primary_job
                        rowDGV.Cells(frm.ColJob.Index).Value = objJob.GetJobCode(.primary_job)
                        If .Increment_Style = ClsTransaksiRecurring.EnumIncrementStyle.Fixed_Increment Then
                            rowDGV.Cells(frm.ColAcc_IncrementStyle.Index).Value = frm.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.Fixed_Increment)
                            rowDGV.Cells(frm.ColAcc_IncrementValue.Index).Value = ObjCD.DigitGrouping(.Increment_Value, DefPrecBase)
                        ElseIf .Increment_Style = ClsTransaksiRecurring.EnumIncrementStyle.No_Increment Then
                            rowDGV.Cells(frm.ColAcc_IncrementStyle.Index).Value = frm.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.No_Increment)
                            rowDGV.Cells(frm.ColAcc_IncrementValue.Index).Value = ObjCD.DigitGrouping(.Increment_Value, DefPrecBase)
                        ElseIf .Increment_Style = ClsTransaksiRecurring.EnumIncrementStyle.Percentage_Increment Then
                            rowDGV.Cells(frm.ColAcc_IncrementStyle.Index).Value = frm.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.Percentage_Increment)
                            rowDGV.Cells(frm.ColAcc_IncrementValue.Index).Value = ObjCD.DigitGrouping(.Increment_Value, 2)
                        End If

                    End With
                Next
            End If



            'set nomer
            ''ObjCD.SetNumber(vDGVItem)
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
            ObjCur = Nothing
            ObjAcc = Nothing
            objDept = Nothing
            objJob = Nothing
        End Try
    End Function

    Public Function IsiDGVEntries(ByVal vDGVEntries As DataGridView, ByVal vPrimaryTransaksi As Double, _
                                  ByVal frm As FrmTransaksiRecuring) As Boolean
        Dim DTA As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journal_transaction_entriesTableAdapter
        Dim TBL As New DseTransaksiRecurring.tbltrans_recurring_journal_transaction_entriesDataTable
        Dim row As DseTransaksiRecurring.tbltrans_recurring_journal_transaction_entriesRow

        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim ObjAcc As New ClsMasterAkun
        Dim objDept As New Core_Module.ClsMasterDepartement
        Dim objJob As New Core_Module.ClsMasterJob

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
            'kosongkan dgv
            vDGVEntries.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVEntries.Rows.Clear()

            'isikan tabel
            If DTA.FillBy(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVEntries.Rows.Add()
                        rowDGV = vDGVEntries.Rows(vDGVEntries.Rows.Count - 1)
                        rowDGV.Cells(frm.ColTrans_PrimaryJournal.Index).Value = .Primary_Journal
                        rowDGV.Cells(frm.ColTrans_JournalNumber.Index).Value = .Journal_No
                        rowDGV.Cells(frm.ColTrans_JournalDate.Index).Value = Format(CDate(.Journal_Date), "dd MMM yyyy")
                        rowDGV.Cells(frm.ColTrans_Description.Index).Value = .Description
                        rowDGV.Cells(frm.ColTrans_Remarks.Index).Value = .Remarks
                        rowDGV.Cells(frm.ColTrans_Status.Index).Value = frm.cboStatus.Items.Item(.Status)
                        rowDGV.Cells(frm.ColTrans_NoUrut.Index).Value = .No_Urut
                    End With
                Next
            End If

            'set nomer
            ObjCD.SetNumber(vDGVEntries)
            vDGVEntries.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
            ObjCur = Nothing
            ObjAcc = Nothing
            objDept = Nothing
            objJob = Nothing
        End Try
    End Function

    Public Function IsiDGVReminder(ByVal vDGVReminder As DataGridView, ByVal vPrimaryTransaksi As Double, _
                                 ByVal frm As FrmTransaksiRecuring) As Boolean
        Dim DTA As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journal_reminderTableAdapter
        Dim TBL As New DseTransaksiRecurring.tbltrans_recurring_journal_reminderDataTable
        Dim row As DseTransaksiRecurring.tbltrans_recurring_journal_reminderRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0
        Dim ObjUser As New Core_Module.ClsUser

        Try
            'kosongkan dgv
            vDGVReminder.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVReminder.Rows.Clear()
            DTA.Connection.ConnectionString = Core_Module.GetConstring


            'isikan tabel
            If DTA.FillBy(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVReminder.Rows.Add()
                        rowDGV = vDGVReminder.Rows(vDGVReminder.Rows.Count - 1)
                        rowDGV.Cells(frm.ColReminder_PrimaryUser.Index).Value = .Primary_User
                        rowDGV.Cells(frm.ColReminder_UserCode.Index).Value = ObjUser.GetKodeUser(.Primary_User)
                        rowDGV.Cells(frm.ColReminder_Remarks.Index).Value = .Remarks
                        rowDGV.Cells(frm.ColReminder_Interval.Index).Value = .Start_Alarm_days_before
                    End With
                Next
            End If

            'set nomer
            ObjCD.SetNumber(vDGVReminder)
            vDGVReminder.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
            ObjUser = Nothing
        End Try
    End Function

    Public Function SimpanDGVItem(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal Frm As FrmTransaksiRecuring) As Boolean
        Dim row As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DTA As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journal_account_entriesTableAdapter

        Dim DefaultPrecision As Integer = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'simpan data Produk Code No
            For Each row In vDGVItem.Rows
                If String.IsNullOrEmpty(row.Cells(Frm.ColRemarks.Index).Value) = True Then
                    row.Cells(Frm.ColRemarks.Index).Value = ""
                End If

                Dim IntStyle As Integer = 0
                If row.Cells(Frm.ColAcc_IncrementStyle.Index).Value = Frm.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.Fixed_Increment) Then
                    IntStyle = ClsTransaksiRecurring.EnumIncrementStyle.Fixed_Increment
                ElseIf row.Cells(Frm.ColAcc_IncrementStyle.Index).Value = Frm.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.No_Increment) Then
                    IntStyle = ClsTransaksiRecurring.EnumIncrementStyle.No_Increment
                ElseIf row.Cells(Frm.ColAcc_IncrementStyle.Index).Value = Frm.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.Percentage_Increment) Then
                    IntStyle = ClsTransaksiRecurring.EnumIncrementStyle.Percentage_Increment
                End If
                If row.Cells(Frm.ColAccount.Index).Value <> "" Then
                    'simpan ke basis data item 
                    DTA.Insert(vPrimaryTransaksi, CDbl(row.Cells(Frm.ColPrimaryAkun.Index).Value), _
                     CDec(row.Cells(Frm.ColDebet.Index).Value), CDec(row.Cells(Frm.ColKredit.Index).Value), _
                    row.Cells(Frm.ColRemarks.Index).Value.ToString, CDec(row.Cells(Frm.ColDebet_OriginalCurr.Index).Value), _
                    CDec(row.Cells(Frm.ColCredit_OriginalCurrency.Index).Value), CDec(row.Cells(Frm.ColExchangeRate.Index).Value), _
                    CDbl(row.Cells(Frm.ColPrimaryDept.Index).Value), CDbl(row.Cells(Frm.ColPrimaryJob.Index).Value), _
                    IntStyle, CDec(row.Cells(Frm.ColAcc_IncrementValue.Index).Value))
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function


    Public Function SimpanDGVEntries(ByVal vDGVEntries As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal Frm As FrmTransaksiRecuring) As Boolean
        Dim row As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DTA As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journal_transaction_entriesTableAdapter

        Dim DefaultPrecision As Integer = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'simpan data Produk Code No
            For Each row In vDGVEntries.Rows
                If String.IsNullOrEmpty(row.Cells(Frm.ColTrans_Remarks.Index).Value) = True Then
                    row.Cells(Frm.ColTrans_Remarks.Index).Value = ""
                End If

                Dim intStatus As Integer = 0
                If row.Cells(Frm.ColTrans_Status.Index).Value = Frm.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed) Then
                    intStatus = ClsTransaksiRecurring.EnumStatus.Planed
                ElseIf row.Cells(Frm.ColTrans_Status.Index).Value = Frm.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Created) Then
                    intStatus = ClsTransaksiRecurring.EnumStatus.Created
                ElseIf row.Cells(Frm.ColTrans_Status.Index).Value = Frm.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Skiped) Then
                    intStatus = ClsTransaksiRecurring.EnumStatus.Skiped
                ElseIf row.Cells(Frm.ColTrans_Status.Index).Value = Frm.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Locked) Then
                    intStatus = ClsTransaksiRecurring.EnumStatus.Locked
                Else


                End If
                If row.Cells(Frm.ColTrans_JournalNumber.Index).Value <> "" Then
                    'simpan ke basis data item 
                    DTA.Insert(vPrimaryTransaksi, CDbl(row.Cells(Frm.ColTrans_PrimaryJournal.Index).Value), _
                    row.Cells(Frm.ColTrans_JournalNumber.Index).Value, row.Cells(Frm.ColTrans_JournalDate.Index).Value, _
                    row.Cells(Frm.ColTrans_Description.Index).Value, row.Cells(Frm.ColTrans_Remarks.Index).Value, _
                    intStatus, CDbl(row.Cells(Frm.ColTrans_NoUrut.Index).Value))
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function


    Public Function SimpanDGVReminders(ByVal vDGVReminders As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal Frm As FrmTransaksiRecuring) As Boolean
        Dim row As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DTA As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journal_reminderTableAdapter

        Dim DefaultPrecision As Integer = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'simpan data Produk Code No
            For Each row In vDGVReminders.Rows
                If String.IsNullOrEmpty(row.Cells(Frm.ColReminder_Remarks.Index).Value) = True Then
                    row.Cells(Frm.ColReminder_Remarks.Index).Value = ""
                End If
                If row.Cells(Frm.ColReminder_UserCode.Index).Value <> "" Then
                    'simpan ke basis data item 
                    DTA.Insert(vPrimaryTransaksi, _
                    CDbl(row.Cells(Frm.ColReminder_PrimaryUser.Index).Value), _
                     row.Cells(Frm.ColReminder_Remarks.Index).Value.ToString, _
                     CDbl(row.Cells(Frm.ColReminder_Interval.Index).Value))
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Overridable Function UpdateStatus(ByVal PrimaryDet As Double, _
                                            ByVal NoJournal As String, _
                                            ByVal dtDate As Date, _
                                            ByVal Remarks As String, _
                                            ByVal status As Integer, _
                                            ByVal NoUrut As Integer, ByVal vUserCode As String) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice, _
            Me.vPrimary_Antrian, PrimaryDet, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiJurnal_General_SimpanPerubahan, vUserCode)


            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.UpdateStatus1(PrimaryDet, NoJournal, dtDate, Remarks, status, NoUrut)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                            Threading.Thread.Sleep(1000)
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjAntri = Nothing
            SW = Nothing
            Me.DTAInOUT.Dispose()
        End Try
    End Function

    Public Overridable Function UpdateStatus1(ByVal PrimaryDet As Double, _
                                              ByVal NoJournal As String, _
                                              ByVal dtDate As Date, _
                                              ByVal Remarks As String, _
                                              ByVal status As Integer, _
                                              ByVal NoUrut As Integer) As Boolean
        Dim DTA As New DseTransaksiRecurringTableAdapters.Khusus_Update_Trans_EntriesTableAdapter
        Dim ObjItem As New ClsTransaksiRecurring_Item

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.UpdateQuery(PrimaryDet, NoJournal, dtDate, Remarks, status, NoUrut, NoJournal) > 0 Then
                'nilai balikan
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + NoJournal + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                ''Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.UpdateStatus1(NoJournal, NoJournal, dtDate, Remarks, status, NoUrut)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    '  Public Function SimpanItem(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ' ByVal Frm As FrmTransaksiJournal_CashBank, ByVal vDGV_Pay_Detail As DataGridView) As Boolean
    '      Dim row As DataGridViewRow
    '      Dim ObjCD As New Core_Module.ClsCekData
    '      Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_itemTableAdapter
    '      Dim DTA_Payment_Detail As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_payment_detailTableAdapter
    '      Dim DefaultPrecision As Integer = 0

    '      Try
    '          'simpan data Produk Code No
    '          For Each row In vDGVItem.Rows
    '              If String.IsNullOrEmpty(row.Cells(Frm.ColRemarks.Index).Value) = True Then
    '                  row.Cells(Frm.ColRemarks.Index).Value = ""
    '              End If
    '              If row.Cells(Frm.ColAccount.Index).Value <> "" Then
    '                  'simpan ke basis data item
    '                  DTA.Insert(vPrimaryTransaksi, CDbl(row.Cells(Frm.ColPrimaryAkun.Index).Value), _
    '                   CDec(row.Cells(Frm.ColDebet.Index).Value), CDec(row.Cells(Frm.ColKredit.Index).Value), _
    '                  row.Cells(Frm.ColRemarks.Index).Value, CDec(row.Cells(Frm.ColDebet_OriginalCurr.Index).Value), _
    '                  CDec(row.Cells(Frm.ColCredit_OriginalCurrency.Index).Value), CDec(row.Cells(Frm.ColExchangeRate.Index).Value), _
    '                  row.Cells(Frm.ColPrimaryDept.Index).Value, row.Cells(Frm.ColPrimaryJob.Index).Value)
    '              End If
    '          Next

    '          For Each row In vDGV_Pay_Detail.Rows
    '              If String.IsNullOrEmpty(row.Cells(Frm.COlItemPay_Remarks.Index).Value) = True Then
    '                  row.Cells(Frm.COlItemPay_Remarks.Index).Value = ""
    '              End If
    '              If row.Cells(Frm.COlItemPay_Account.Index).Value <> "" Then
    '                  DTA_Payment_Detail.Insert(vPrimaryTransaksi, CDec(row.Cells(Frm.COlItemPay_AmountOriginal.Index).Value), row.Cells(Frm.COlItemPay_Remarks.Index).Value, _
    '                  CDec(row.Cells(Frm.ColItemPay_Amount.Index).Value), CDbl(row.Cells(Frm.COlItemPay_PrimaryDept.Index).Value), CDbl(row.Cells(Frm.COlItemPay_PrimaryJob.Index).Value), _
    '                  CDbl(row.Cells(Frm.COlItemPay_PrimaryAkun.Index).Value), CDec(row.Cells(Frm.ColItemPay_ExRate.Index).Value))
    '              End If
    '          Next
    '          Return True
    '      Catch ex As Exception
    '          MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '      Finally
    '          DTA.Dispose()
    '          ObjCD = Nothing
    '          DTA_Payment_Detail.Dispose()
    '      End Try
    '  End Function

    '  '' untuk sales journal
    '  ''

    Public Function GetJml(ByVal vPrimaryAkun As Double) As Double
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_itemTableAdapter
        Try
            If DTA.ScalarQuery_GetJmlJournalByPrimaryAkun(vPrimaryAkun).HasValue Then
                Return DTA.ScalarQuery_GetJmlJournalByPrimaryAkun(vPrimaryAkun).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTAItem As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journal_account_entriesTableAdapter
        Dim DTAEntries As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journal_transaction_entriesTableAdapter
        Dim DTAReminder As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journal_reminderTableAdapter

        Try
            'hapus data item
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAEntries.Connection.ConnectionString = Core_Module.GetConstring
            DTAReminder.Connection.ConnectionString = Core_Module.GetConstring

            DTAItem.Delete(vPrimaryTransaksi)
            DTAEntries.Delete(vPrimaryTransaksi)
            DTAReminder.Delete(vPrimaryTransaksi)


            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            DTAEntries.Dispose()
            DTAReminder.Dispose()
        End Try
    End Function

    '    ''Untuk Sales Journal
    '    Public Function SimpanItemSalesJournal(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
    '       ByVal Frm As FrmTransaksiJournal_Sales, ByVal vDGVItemPayment As DataGridView, _
    '       ByVal vTipeJurnal As ClsTransaksiJournal.EnumTipeJournal, ByVal vJurnalNumber As String, _
    '       ByVal vPrimaryCust As Double, ByVal vTanggalJurnal As Date, ByVal vDGVItemAkunOther As DataGridView) As Boolean
    '        Dim row As DataGridViewRow
    '        Dim ObjCD As New Core_Module.ClsCekData
    '        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_itemTableAdapter
    '        Dim DefaultPrecision As Integer = 0
    '        Dim DTA_Payment As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_paymentTableAdapter
    '        Dim TipePayment As ClsTransaksiSalesBalance.EnumTipePayment = ClsTransaksiSalesBalance.EnumTipePayment.SalesInvoice_SalesInvoiceDP
    '        Dim DTABalance As New DseTransaksiSalesBalanceTableAdapters.tbltrans_balance_salesTableAdapter
    '        Dim DTABalance_SR As New DseTransaksiSalesReturnBalanceTableAdapters.tbltrans_balance_salesreturnTableAdapter
    '        Dim ObjKartu As New ClsTransaksi_Kartu_APAR
    '        Dim DTAOther As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_otherTableAdapter

    '        Try
    '            'simpan data item
    '            For Each row In vDGVItem.Rows
    '                If String.IsNullOrEmpty(row.Cells(Frm.ColRemarks.Index).Value) = True Then
    '                    row.Cells(Frm.ColRemarks.Index).Value = ""
    '                End If
    '                If row.Cells(Frm.ColAccount.Index).Value <> "" Then
    '                    'simpan ke basis data item
    '                    DTA.Insert(vPrimaryTransaksi, CDbl(row.Cells(Frm.ColPrimaryAkun.Index).Value), _
    '                    CDec(row.Cells(Frm.ColDebet.Index).Value), CDec(row.Cells(Frm.ColKredit.Index).Value), _
    '                    row.Cells(Frm.ColRemarks.Index).Value, CDec(row.Cells(Frm.ColDebetOriginal.Index).Value), CDec(row.Cells(Frm.ColCreditOriginal.Index).Value), _
    '                    CDec(row.Cells(Frm.ColExchangeRate.Index).Value), -1, -1)
    '                End If
    '            Next

    '            If vTipeJurnal = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
    '                TipePayment = ClsTransaksiSalesBalance.EnumTipePayment.SalesInvoice_SalesInvoiceDP
    '            ElseIf vTipeJurnal = ClsTransaksiJournal.EnumTipeJournal.SalesReturnPaymentJournal Then
    '                TipePayment = ClsTransaksiSalesBalance.EnumTipePayment.SalesReturn
    '            End If

    '            For Each row In vDGVItemPayment.Rows
    '                If CDec(row.Cells(Frm.ColReceivement.Index).Value) <> 0 Then
    '                    If vTipeJurnal = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
    '                        'simpan data receivement
    '                        DTA_Payment.Insert(vPrimaryTransaksi, CDec(row.Cells(Frm.ColReceivement.Index).Value), TipePayment, _
    '                        row.Cells(Frm.ColPrimaryDocumenAcuan_Sales.Index).Value, row.Cells(Frm.ColTipeDocumentAcuan_Sales.Index).Value, _
    '                        row.Cells(Frm.ColPrimaryCurrencyBalance.Index).Value, row.Cells(Frm.ColExhangeRateBalance.Index).Value, _
    '                        row.Cells(Frm.ColReceivementOriginal.Index).Value, row.Cells(Frm.ColPrimaryAkunBalance.Index).Value, _
    '                        row.Cells(Frm.ColInvoiceNo.Index).Value, "")

    '                        'update data balance dan received
    '                        DTABalance.UpdateQuery(row.Cells(Frm.ColBalanceSales.Index).Value, _
    '                        CDec(row.Cells(Frm.ColReceivement.Index).Value) + CDec(row.Cells(Frm.ColReceived.Index).Value), _
    '                        CDec(row.Cells(Frm.ColReceivementOriginal.Index).Value) + CDec(row.Cells(Frm.ColReceivedOriginal.Index).Value), _
    '                        row.Cells(Frm.ColBalanceOriginal.Index).Value, _
    '                        row.Cells(Frm.ColPrimarySalesBalance.Index).Value, CDbl(row.Cells(Frm.ColPrimaryCurrencyBalance.Index).Value), _
    '                        row.Cells(Frm.ColTipeDocumentAcuan_Sales.Index).Value)

    '                        'simpan data kartu
    '                        If CDec(row.Cells(Frm.ColReceivement.Index).Value) > 0 Then
    '                            'untuk pembayaran sales invoice
    '                            ObjKartu.Simpan(vPrimaryTransaksi, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Sales_Receivement, _
    '                            vJurnalNumber, vPrimaryCust, 0, CDec(row.Cells(Frm.ColReceivement.Index).Value), _
    '                            row.Cells(Frm.ColInvoiceNo.Index).Value, ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR, vTanggalJurnal, _
    '                            CDbl(row.Cells(Frm.ColPrimaryDocumenAcuan_Sales.Index).Value), CDate(row.Cells(Frm.ColInvoiceDate.Index).Value), _
    '                            Math.Abs((CDate(row.Cells(Frm.ColInvoiceDate.Index).Value) - CDate(row.Cells(Frm.ColDueDateInvoice.Index).Value)).Days), _
    '                            Me.GetTipeTransaksiAsalSalesReceivementJurnal(row.Cells(Frm.ColTipeDocumentAcuan_Sales.Index).Value), row.Cells(Frm.ColPrimaryCurrencyBalance.Index).Value, _
    '                            0, CDec(row.Cells(Frm.ColReceivementOriginal.Index).Value), CDec(row.Cells(Frm.ColExhangeRateBalance.Index).Value), _
    '                            CDec(row.Cells(Frm.ColPrimaryAkunBalance.Index).Value))
    '                        Else
    '                            'untuk pembayaran sales return
    '                            ObjKartu.Simpan(vPrimaryTransaksi, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Sales_Receivement, _
    '                            vJurnalNumber, vPrimaryCust, Math.Abs(CDec(row.Cells(Frm.ColReceivement.Index).Value)), 0, _
    '                            row.Cells(Frm.ColInvoiceNo.Index).Value, ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR, vTanggalJurnal, _
    '                            CDbl(row.Cells(Frm.ColPrimaryDocumenAcuan_Sales.Index).Value), CDate(row.Cells(Frm.ColInvoiceDate.Index).Value), _
    '                            Math.Abs((CDate(row.Cells(Frm.ColInvoiceDate.Index).Value) - CDate(row.Cells(Frm.ColDueDateInvoice.Index).Value)).Days), _
    '                            Me.GetTipeTransaksiAsalSalesReceivementJurnal(row.Cells(Frm.ColTipeDocumentAcuan_Sales.Index).Value), row.Cells(Frm.ColPrimaryCurrencyBalance.Index).Value, _
    '                            Math.Abs(CDec(row.Cells(Frm.ColReceivementOriginal.Index).Value)), 0, CDec(row.Cells(Frm.ColExhangeRateBalance.Index).Value), _
    '                             CDec(row.Cells(Frm.ColPrimaryAkunBalance.Index).Value))
    '                        End If
    '                    ElseIf vTipeJurnal = ClsTransaksiJournal.EnumTipeJournal.SalesReturnPaymentJournal Then
    '                        MsgBox("Logic error, please contact your vendor if you see this message", MsgBoxStyle.Information, "Information")
    '                    End If
    '                End If
    '            Next

    '            'simpan akun other
    '            For Each row In vDGVItemAkunOther.Rows
    '                With row
    '                    DTAOther.Insert(vPrimaryTransaksi, .Cells(Frm.ColAccountDiff_PrimaryAkun.Index).Value, .Cells(Frm.ColAccountDiff_Value.Index).Value, _
    '                    .Cells(Frm.ColAccountDiff_ValueOriginalCurrency.Index).Value, .Cells(Frm.ColAccountDiff_Remarks.Index).Value, _
    '                    .Cells(Frm.ColAccountDiff_PrimaryCurrency.Index).Value, .Cells(Frm.ColAccountDiff_ExchangeRate.Index).Value)
    '                End With
    '            Next

    '            Return True
    '        Catch ex As Exception
    '            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Finally
    '            DTA.Dispose()
    '            DTA_Payment.Dispose()
    '            DTABalance.Dispose()
    '            DTABalance_SR.Dispose()
    '            ObjCD = Nothing
    '            ObjKartu = Nothing
    '            DTAOther.Dispose()
    '        End Try
    '    End Function

    '    Public Function HapusItem_SalesJurnal(ByVal vPrimaryTransaksi As Double, ByVal vTipeJurnal As ClsTransaksiJournal.EnumTipeJournal) As Boolean
    '        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_itemTableAdapter
    '        Dim DTA_Payment As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_paymentTableAdapter
    '        Dim TBL_Payment As New DseTransaksiJournal.tbltrans_jurnal_item_paymentDataTable
    '        Dim row_Payment As DseTransaksiJournal.tbltrans_jurnal_item_paymentRow
    '        Dim DTABalance As New DseTransaksiSalesBalanceTableAdapters.tbltrans_balance_salesTableAdapter
    '        'Dim DTABalance_SR As New DseTransaksiSalesReturnBalanceTableAdapters.tbltrans_balance_salesreturnTableAdapter
    '        Dim Balance_Lama, Received_Lama, BalanceOriginal_Lama, ReceivedOriginal_Lama As Decimal
    '        Dim ObjKartu As New ClsTransaksi_Kartu_APAR
    '        Dim DTAOther As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_otherTableAdapter

    '        Try
    '            'update data balance sales
    '            DTA_Payment.FillBy_PrimaryJurnal(TBL_Payment, vPrimaryTransaksi)

    '            For Each row_Payment In TBL_Payment.Rows
    '                If vTipeJurnal = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
    '                    If DTABalance.ScalarQuery_GetBalance(row_Payment.Primary_Document_Acuan, _
    '                    row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).HasValue = True Then
    '                        Balance_Lama = DTABalance.ScalarQuery_GetBalance(row_Payment.Primary_Document_Acuan, _
    '                        row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).Value
    '                    Else
    '                        Balance_Lama = 0
    '                    End If

    '                    If DTABalance.ScalarQuery_GetReceived(row_Payment.Primary_Document_Acuan, _
    '                    row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).HasValue = True Then
    '                        Received_Lama = DTABalance.ScalarQuery_GetReceived(row_Payment.Primary_Document_Acuan, _
    '                        row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).Value
    '                    Else
    '                        Received_Lama = 0
    '                    End If

    '                    If DTABalance.ScalarQuery_GetBalanceOriginalLama(row_Payment.Primary_Document_Acuan, _
    '                    row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).HasValue = True Then
    '                        BalanceOriginal_Lama = DTABalance.ScalarQuery_GetBalanceOriginalLama(row_Payment.Primary_Document_Acuan, _
    '                        row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).Value
    '                    Else
    '                        BalanceOriginal_Lama = 0
    '                    End If

    '                    If DTABalance.ScalarQuery_GetReceivedOriginalLama(row_Payment.Primary_Document_Acuan, _
    '                   row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).HasValue = True Then
    '                        ReceivedOriginal_Lama = DTABalance.ScalarQuery_GetReceivedOriginalLama(row_Payment.Primary_Document_Acuan, _
    '                   row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).Value
    '                    Else
    '                        ReceivedOriginal_Lama = 0
    '                    End If

    '                    DTABalance.UpdateQuery1(Balance_Lama + row_Payment.Payment, Received_Lama - _
    '                    row_Payment.Payment, BalanceOriginal_Lama + row_Payment.payment_original, ReceivedOriginal_Lama - row_Payment.payment_original, _
    '                    row_Payment.Primary_Document_Acuan, row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency)

    '                    'hapus data kartu
    '                    ObjKartu.Hapus(vPrimaryTransaksi, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Sales_Receivement)
    '                ElseIf vTipeJurnal = ClsTransaksiJournal.EnumTipeJournal.SalesReturnPaymentJournal Then
    '                    MsgBox("Logic error, contact your vendor if you see this message", MsgBoxStyle.Information, "Information")
    '                End If
    '            Next

    '            'hapus data item payment + jurnal + other
    '            DTA_Payment.Delete(vPrimaryTransaksi)
    '            DTA.Delete(vPrimaryTransaksi)
    '            DTAOther.Delete(vPrimaryTransaksi)

    '            'nilai balikan
    '            Return True
    '        Catch ex As Exception
    '            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Finally
    '            DTA.Dispose()
    '            DTA_Payment.Dispose()
    '            TBL_Payment.Dispose()
    '            ObjKartu = Nothing
    '            DTAOther.Dispose()
    '        End Try
    '    End Function

    '    ''Untuk Purchase Journal
    '    ''
    '    Public Function SimpanItemPurchaseJournal(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
    '    ByVal Frm As FrmTransaksiJournal_Purchase, ByVal vDGVItemPayment As DataGridView, _
    '    ByVal vTipeJurnal As ClsTransaksiJournal.EnumTipeJournal, ByVal vJurnalNumber As String, _
    '    ByVal vPrimarySupplier As Double, ByVal vTanggalJurnal As Date, ByVal vDGVItemAkunOther As DataGridView) As Boolean
    '        Dim row As DataGridViewRow
    '        Dim ObjCD As New Core_Module.ClsCekData
    '        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_itemTableAdapter
    '        Dim DefaultPrecision As Integer = 0
    '        Dim DTA_Payment As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_paymentTableAdapter
    '        Dim TipePayment As ClsTransaksiSalesBalance.EnumTipePayment = ClsTransaksiSalesBalance.EnumTipePayment.AOP_AOPDP
    '        Dim DTABalance As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_balance_purchaseTableAdapter
    '        Dim DTABalance_SR As New DseTransaskiPurchaseReturnBalanceTableAdapters.tbltrans_balance_purchasereturnTableAdapter
    '        Dim ObjKartu As New ClsTransaksi_Kartu_APAR
    '        Dim DTAOther As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_otherTableAdapter

    '        Try
    '            'simpan data Produk Code No
    '            For Each row In vDGVItem.Rows
    '                If String.IsNullOrEmpty(row.Cells(Frm.ColRemarks.Index).Value) = True Then
    '                    row.Cells(Frm.ColRemarks.Index).Value = ""
    '                End If
    '                If row.Cells(Frm.ColAccount.Index).Value <> "" Then
    '                    'simpan ke basis data item
    '                    DTA.Insert(vPrimaryTransaksi, CDbl(row.Cells(Frm.ColPrimaryAkun.Index).Value), _
    '                    CDec(row.Cells(Frm.ColDebet.Index).Value), CDec(row.Cells(Frm.ColKredit.Index).Value), _
    '                    row.Cells(Frm.ColRemarks.Index).Value, CDec(row.Cells(Frm.ColDebetOriginal.Index).Value), CDec(row.Cells(Frm.ColCreditOriginal.Index).Value), _
    '                    CDec(row.Cells(Frm.ColExchangeRate.Index).Value), -1, -1)
    '                End If
    '            Next

    '            If vTipeJurnal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
    '                TipePayment = ClsTransaksiSalesBalance.EnumTipePayment.AOP_AOPDP
    '            ElseIf vTipeJurnal = ClsTransaksiJournal.EnumTipeJournal.PurchaseReturnReceivement Then
    '                TipePayment = ClsTransaksiSalesBalance.EnumTipePayment.Purchase_Return
    '            End If

    '            For Each row In vDGVItemPayment.Rows
    '                If CDec(row.Cells(Frm.ColPayment.Index).Value) <> 0 Then
    '                    If vTipeJurnal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
    '                        'simpan data payment
    '                        DTA_Payment.Insert(vPrimaryTransaksi, CDec(row.Cells(Frm.ColPayment.Index).Value), TipePayment, _
    '                        row.Cells(Frm.ColPrimaryDocumentAcuan_Purchase.Index).Value, row.Cells(Frm.ColTipeDocumentAcuan_Purchase.Index).Value, _
    '                        row.Cells(Frm.ColPrimaryCurrencyBalance.Index).Value, row.Cells(Frm.ColExchangeRateBalance.Index).Value, _
    '                        row.Cells(Frm.ColPaymentOriginal.Index).Value, row.Cells(Frm.ColPrimaryAkunBalance.Index).Value, _
    '                        row.Cells(Frm.ColAOPNumber.Index).Value, row.Cells(Frm.ColSupplierInvoiceNumber.Index).Value)

    '                        'update data balance dan received
    '                        DTABalance.UpdateQuery(row.Cells(Frm.ColBalancePurchase.Index).Value, _
    '                        CDec(row.Cells(Frm.ColPayment.Index).Value) + CDec(row.Cells(Frm.ColPaid.Index).Value), _
    '                        CDec(row.Cells(Frm.ColBalanceOriginal.Index).Value), _
    '                        CDec(row.Cells(Frm.ColPaymentOriginal.Index).Value + CDec(row.Cells(Frm.ColPaidOriginal.Index).Value)), _
    '                        row.Cells(Frm.ColPrimaryPurchaseBalance.Index).Value, CDbl(row.Cells(Frm.ColPrimaryCurrencyBalance.Index).Value), _
    '                        row.Cells(Frm.ColTipeDocumentAcuan_Purchase.Index).Value)

    '                        'simpan data kartu
    '                        If CDec(row.Cells(Frm.ColPaymentOriginal.Index).Value) > 0 Then
    '                            'pembayaran aop, didebet
    '                            ObjKartu.Simpan(vPrimaryTransaksi, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Purchase_Payment, _
    '                            vJurnalNumber, vPrimarySupplier, CDec(row.Cells(Frm.ColPayment.Index).Value), 0, _
    '                            row.Cells(Frm.ColAOPNumber.Index).Value, ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
    '                            vTanggalJurnal, CDbl(row.Cells(Frm.ColPrimaryDocumentAcuan_Purchase.Index).Value), CDate(row.Cells(Frm.ColAOPDate.Index).Value), _
    '                            Math.Abs((CDate(row.Cells(Frm.ColDueDatePurchase.Index).Value) - CDate(row.Cells(Frm.ColAOPDate.Index).Value)).Days), _
    '                            Me.GetTipeTransaksiAsalPurchasePaymentJurnal(row.Cells(Frm.ColTipeDocumentAcuan_Purchase.Index).Value), row.Cells(Frm.ColPrimaryCurrencyBalance.Index).Value, _
    '                            Math.Abs(CDec(row.Cells(Frm.ColPaymentOriginal.Index).Value)), 0, CDec(row.Cells(Frm.ColExchangeRateBalance.Index).Value), _
    '                            CDec(row.Cells(Frm.ColPrimaryAkunBalance.Index).Value))
    '                        Else
    '                            'pembayaran pd dikredit
    '                            ObjKartu.Simpan(vPrimaryTransaksi, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Purchase_Payment, _
    '                            vJurnalNumber, vPrimarySupplier, 0, Math.Abs(CDec(row.Cells(Frm.ColPayment.Index).Value)), _
    '                            row.Cells(Frm.ColAOPNumber.Index).Value, ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
    '                            vTanggalJurnal, CDbl(row.Cells(Frm.ColPrimaryDocumentAcuan_Purchase.Index).Value), CDate(row.Cells(Frm.ColAOPDate.Index).Value), _
    '                            Math.Abs((CDate(row.Cells(Frm.ColDueDatePurchase.Index).Value) - CDate(row.Cells(Frm.ColAOPDate.Index).Value)).Days), _
    '                            Me.GetTipeTransaksiAsalPurchasePaymentJurnal(row.Cells(Frm.ColTipeDocumentAcuan_Purchase.Index).Value), row.Cells(Frm.ColPrimaryCurrencyBalance.Index).Value, _
    '                            0, Math.Abs(CDec(row.Cells(Frm.ColPaymentOriginal.Index).Value)), CDec(row.Cells(Frm.ColExchangeRateBalance.Index).Value), _
    '                            CDec(row.Cells(Frm.ColPrimaryAkunBalance.Index).Value))
    '                        End If
    '                    ElseIf vTipeJurnal = ClsTransaksiJournal.EnumTipeJournal.PurchaseReturnReceivement Then
    '                        MsgBox("logic error please contact your vendor if you see this message", MsgBoxStyle.Information, "Information")
    '                    End If
    '                End If
    '            Next

    '            'simpan akun other
    '            For Each row In vDGVItemAkunOther.Rows
    '                With row
    '                    DTAOther.Insert(vPrimaryTransaksi, .Cells(Frm.ColAccountDiff_PrimaryAkun.Index).Value, .Cells(Frm.ColAccountDiff_Value.Index).Value, _
    '                    .Cells(Frm.ColAccountDiff_ValueOriginalCurrency.Index).Value, .Cells(Frm.ColAccountDiff_Remarks.Index).Value, _
    '                    .Cells(Frm.ColAccountDiff_PrimaryCurrency.Index).Value, .Cells(Frm.ColAccountDiff_ExchangeRate.Index).Value)
    '                End With
    '            Next

    '            Return True
    '        Catch ex As Exception
    '            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Finally
    '            DTA.Dispose()
    '            DTA_Payment.Dispose()
    '            DTABalance.Dispose()
    '            DTABalance_SR.Dispose()
    '            ObjCD = Nothing
    '            ObjKartu = Nothing
    '            DTAOther.Dispose()
    '        End Try
    '    End Function

    '    Private Function GetTipeTransaksiAsalPurchasePaymentJurnal(ByVal vTipeBalance As _
    'ClsTransaksiSalesBalance.EnumTipeSalesBalance) As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal
    '        Try
    '            If vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic Then
    '                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis
    '            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual Then
    '                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual
    '            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_Transfer_Invoice Then
    '                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual_Transfer
    '            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR Then
    '                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Otomatis
    '            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR Then
    '                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Manual
    '            End If
    '        Catch ex As Exception
    '            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        End Try
    '    End Function


    '    Private Function GetTipeTransaksiAsalSalesReceivementJurnal(ByVal vTipeBalance As _
    '    ClsTransaksiSalesBalance.EnumTipeSalesBalance) As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal
    '        Try
    '            If vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic Then
    '                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis
    '            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual Then
    '                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual
    '            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_Transfer_Invoice Then
    '                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual_Transfer
    '            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR Then
    '                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis
    '            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR Then
    '                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual
    '            End If
    '        Catch ex As Exception
    '            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        End Try
    '    End Function

    '    Public Function HapusItem_PurchaseJurnal(ByVal vPrimaryTransaksi As Double, ByVal vTipeJurnal As ClsTransaksiJournal.EnumTipeJournal) As Boolean
    '        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_itemTableAdapter
    '        Dim DTA_Payment As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_paymentTableAdapter
    '        Dim TBL_Payment As New DseTransaksiJournal.tbltrans_jurnal_item_paymentDataTable
    '        Dim row_Payment As DseTransaksiJournal.tbltrans_jurnal_item_paymentRow
    '        Dim DTABalance_Purchase As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_balance_purchaseTableAdapter
    '        'Dim DTABalance_PR As New DseTransaskiPurchaseReturnBalanceTableAdapters.tbltrans_balance_purchasereturnTableAdapter
    '        Dim Balance_Lama, Received_Lama, Balance_Lama_Original, Received_Lama_Original As Decimal
    '        Dim ObjKartu As New ClsTransaksi_Kartu_APAR
    '        Dim DTAPurchase As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_otherTableAdapter
    '        Dim DTAOther As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_otherTableAdapter

    '        Try
    '            'update data balance sales
    '            DTA_Payment.FillBy_PrimaryJurnal(TBL_Payment, vPrimaryTransaksi)
    '            For Each row_Payment In TBL_Payment.Rows
    '                If vTipeJurnal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
    '                    If DTABalance_Purchase.ScalarQuery_GetBalance(row_Payment.Primary_Document_Acuan, _
    '                    row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).HasValue = True Then
    '                        Balance_Lama = DTABalance_Purchase.ScalarQuery_GetBalance(row_Payment.Primary_Document_Acuan, _
    '                        row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).Value
    '                    Else
    '                        Balance_Lama = 0
    '                    End If
    '                    If DTABalance_Purchase.ScalarQuery_GetPaid(row_Payment.Primary_Document_Acuan, _
    '                    row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).HasValue = True Then
    '                        Received_Lama = DTABalance_Purchase.ScalarQuery_GetPaid(row_Payment.Primary_Document_Acuan, _
    '                        row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).Value
    '                    Else
    '                        Received_Lama = 0
    '                    End If

    '                    If DTABalance_Purchase.ScalarQuery_GetBalanceOriginal2(row_Payment.Primary_Document_Acuan, _
    '                    row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).HasValue = True Then
    '                        Balance_Lama_Original = DTABalance_Purchase.ScalarQuery_GetBalanceOriginal2(row_Payment.Primary_Document_Acuan, _
    '                        row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).Value
    '                    Else
    '                        Balance_Lama_Original = 0
    '                    End If
    '                    If DTABalance_Purchase.ScalarQuery_GetPaidOriginal(row_Payment.Primary_Document_Acuan, _
    '                    row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).HasValue = True Then
    '                        Received_Lama_Original = DTABalance_Purchase.ScalarQuery_GetPaidOriginal(row_Payment.Primary_Document_Acuan, _
    '                        row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).Value
    '                    Else
    '                        Received_Lama_Original = 0
    '                    End If

    '                    DTABalance_Purchase.UpdateQuery1(Balance_Lama + row_Payment.Payment, Received_Lama - _
    '                    row_Payment.Payment, Balance_Lama_Original + row_Payment.payment_original, _
    '                    Received_Lama_Original - row_Payment.payment_original, row_Payment.Primary_Document_Acuan, _
    '                    row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency)

    '                    ObjKartu.Hapus(vPrimaryTransaksi, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Purchase_Payment)
    '                ElseIf vTipeJurnal = ClsTransaksiJournal.EnumTipeJournal.PurchaseReturnReceivement Then
    '                    MsgBox("Logic error, contact your vendor if you see this message", MsgBoxStyle.Information, "Information")
    '                End If
    '            Next

    '            'hapus data item payment + jurnal
    '            DTA_Payment.Delete(vPrimaryTransaksi)
    '            DTA.Delete(vPrimaryTransaksi)
    '            DTAPurchase.Delete(vPrimaryTransaksi)
    '            DTAOther.Delete(vPrimaryTransaksi)

    '            'nilai balikan
    '            Return True
    '        Catch ex As Exception
    '            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Finally
    '            DTA.Dispose()
    '            DTA_Payment.Dispose()
    '            TBL_Payment.Dispose()
    '            DTAPurchase.Dispose()
    '            ObjKartu = Nothing
    '            DTAOther.Dispose()
    '        End Try
    '    End Function

End Class
