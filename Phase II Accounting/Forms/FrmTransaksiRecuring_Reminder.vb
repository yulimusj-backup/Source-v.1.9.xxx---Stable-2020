Option Explicit On
Option Strict Off

Public Class FrmTransaksiRecuring_Reminder
    Dim tblHeader As New DseTransaksiRecurring.tblHeaderDataTable
    Dim tblItem As New DseTransaksiRecurring.tblItemDataTable
    Public UserCode As String = ""
    Dim rowHeader As DseTransaksiRecurring.tblHeaderRow
    Dim rowItem As DseTransaksiRecurring.tblItemRow
    Public cboStatus As New DataGridViewComboBoxCell
    Dim ObjCD As New Core_Module.ClsCekData
    Dim PrimaryBaseCurrency As Double = 1
    Dim rowIndex As Integer = 0
    Dim IRowIndex As Integer = -1
    Dim Splitter1_SplitterDistance As Integer = 0
    Public BackcolorAll As Color = Color.LightSeaGreen
    Dim BlnUnSelect As Boolean = False


    Private Sub FrmTransaksiRecuring_Reminder_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

        Try
            FrmTransaksiRecuring_Reminder_Load(Nothing, Nothing)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiRecuring_Reminder_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Leave

    End Sub

    Private Sub FrmTransaksiRecuring_Reminder_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Click

        Try
            FrmTransaksiRecuring_Reminder_Load(Nothing, Nothing)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiRecuring_Reminder_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus

        Try
            FrmTransaksiRecuring_Reminder_Load(Nothing, Nothing)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FrmTransaksiRecuring_Reminder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim ObjUser As New Core_Module.ClsUser
        Dim PrimaryUser As Double = -1
        Try
            IRowIndex = -1
            rowIndex = 0
            Me.tblHeader.Rows.Clear()
            Me.tblItem.Rows.Clear()
            Me.dgvHeader.Rows.Clear()
            Me.dgvItem.Rows.Clear()

            cboStatus.Items.Clear()
            Me.cboStatus.Items.Add("Planed")
            Me.cboStatus.Items.Add("Created")
            Me.cboStatus.Items.Add("Skiped")
            Me.cboStatus.Items.Add("Lock")

            PrimaryUser = ObjUser.GetPrimary(Me.UserCode)
            Me.FillHeaderItem(PrimaryUser, Date.Now)
            Me.isiKeDGVHeader()

            Dim PrimaryTransaksi As Double = -1
            Dim NoJournal As String = ""
            If Me.dgvHeader.Rows.Count > 0 Then
                PrimaryTransaksi = Me.dgvHeader.Item(Me.ColPrimaryRecurring.Index, Me.rowIndex).Value
                NoJournal = Me.dgvHeader.Item(Me.ColRecurringTrans_JournalNumber.Index, Me.rowIndex).Value
                Me.isiKeDGVItem(PrimaryTransaksi, NoJournal)
                'Else
                'Me.Close()
            End If

            Me.dgvHeader.AlternatingRowsDefaultCellStyle.BackColor = Me.BackcolorAll
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackcolorAll
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjUser = Nothing
        End Try

    End Sub

    Public Sub isiKeDGVHeader()
        Try
            Me.dgvHeader.Rows.Clear()
            Me.dgvItem.Rows.Clear()
            If Me.tblHeader.Rows.Count > 0 Then
                Dim rowDGV As New DataGridViewRow
                ' Me.dgvHeader.Rows.Clear()
                For Each Me.rowHeader In Me.tblHeader.Rows
                    With rowHeader
                        Me.dgvHeader.Rows.Add()
                        rowDGV = Me.dgvHeader.Rows(Me.dgvHeader.Rows.Count - 1)
                        rowDGV.Cells(Me.ColPrimaryRecurring.Index).Value = .Primary_Recurring
                        rowDGV.Cells(Me.ColRecurringTrans_RecurringNo.Index).Value = .Recurring_No
                        rowDGV.Cells(Me.ColRecurringTrans_ReccuringDesc.Index).Value = .Description
                        rowDGV.Cells(Me.ColRecurringTrans_PrimaryJournal.Index).Value = .Primary_Journal
                        rowDGV.Cells(Me.ColRecurringTrans_JournalNumber.Index).Value = .Journal_No
                        rowDGV.Cells(Me.ColRecurringTrans_JournalDate.Index).Value = Format(.Journal_Date, "dd MMM yyyy")
                        rowDGV.Cells(Me.ColRecurringTrans_Description.Index).Value = .Journal_Description
                        rowDGV.Cells(Me.ColRecurringTrans_Remarks.Index).Value = .Journal_Remarks
                        rowDGV.Cells(Me.ColRecurringTrans_Status.Index).Value = Me.cboStatus.Items.Item(.Journal_status)
                    End With
                Next
            End If
            '  MsgBox(Me.tblHeader.Rows.Count)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub isiKeDGVItem(ByVal VPrimaryTransaksi As Double, ByVal vNoJournal As String)
        Dim ObjDept As New Core_Module.ClsMasterDepartement
        Dim ObjJob As New Core_Module.ClsMasterJob
        Dim Rounding As Integer = 0
        Dim Objcurr As New Core_Module.clsMasterCurrency
        Try
            Me.tblItem.DefaultView.RowFilter = Nothing
            Me.tblItem.DefaultView.RowFilter = "[Journal No] = '" & vNoJournal & _
                                                 "' And [Primary Transaksi] = " & VPrimaryTransaksi
            Me.dgvItem.Rows.Clear()
            Rounding = Objcurr.GetDecimal_DefaultCurrency
            If Me.tblItem.DefaultView.Count > 0 Then
                Dim dv As Data.DataRowView
                Dim rowDGV As New DataGridViewRow
                For Each dv In Me.tblItem.DefaultView
                    rowItem = dv.Row
                    With rowItem
                        Me.dgvItem.Rows.Add()
                        rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                        rowDGV.Cells(Me.ColPrimaryRecurringDet.Index).Value = .Primary_Transaksi
                        rowDGV.Cells(Me.ColPrimaryJournal.Index).Value = .Primary_Journal
                        rowDGV.Cells(Me.ColPrimaryAkun.Index).Value = .Primary_Akun
                        Dim defPrec As Decimal = 0
                        defPrec = Objcurr.GetDefaultPrecission(.Primary_Currency)
                        rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value = .Primary_Currency
                        rowDGV.Cells(Me.ColOriginalCurrency.Index).Value = Objcurr.GetCurrencyCode(.Primary_Currency)
                        rowDGV.Cells(Me.ColAccount.Index).Value = .Account_No
                        rowDGV.Cells(Me.ColDescAccount.Index).Value = .Account_Description
                        rowDGV.Cells(Me.ColDebet.Index).Value = ObjCD.DigitGrouping(.Debet, Rounding)
                        rowDGV.Cells(Me.ColKredit.Index).Value = ObjCD.DigitGrouping(.Kredit, Rounding)
                        rowDGV.Cells(Me.ColExchangeRate.Index).Value = ObjCD.DigitGrouping(.Ratio, Rounding)
                        rowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(.Debet_Original, defPrec)
                        rowDGV.Cells(Me.ColCredit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(.Credit_Original, defPrec)
                        rowDGV.Cells(Me.ColRemarks.Index).Value = .Remarks
                        rowDGV.Cells(Me.ColPrimaryDept.Index).Value = .Primary_Dept
                        rowDGV.Cells(Me.ColDept.Index).Value = ObjDept.GetDeptCode(.Primary_Dept)
                        rowDGV.Cells(Me.ColPrimaryJob.Index).Value = .Primary_Job
                        rowDGV.Cells(Me.ColJob.Index).Value = ObjDept.GetDeptCode(.Primary_Job)
                    End With
                Next
            End If
            SetMulticurrency()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjDept = Nothing
            ObjJob = Nothing
            Objcurr = Nothing
        End Try
    End Sub


    Public Sub FillHeaderItem(ByVal VPrimaryUser As Double, ByVal dtDate As Date)
        Dim DTARecc As New DseTransaksiRecurringTableAdapters.RecurringActiveTableAdapter
        Dim DTAJ As New DseTransaksiRecurringTableAdapters.Trans_Journal_ActiveTableAdapter

        Dim TblRecc As New DseTransaksiRecurring.RecurringActiveDataTable
        Dim tblJ As New DseTransaksiRecurring.Trans_Journal_ActiveDataTable

        Dim rowRecc As DseTransaksiRecurring.RecurringActiveRow
        Dim rowJ As DseTransaksiRecurring.Trans_Journal_ActiveRow
        Dim ObjAcc As New ClsMasterAkun
        Dim ObjRec As New ClsTransaksiRecurring
        Dim ObjCur As New Core_Module.clsMasterCurrency


        Try
            DTARecc.Connection.ConnectionString = Core_Module.GetConstring
            DTAJ.Connection.ConnectionString = DTARecc.Connection.ConnectionString
            DTARecc.Fill(TblRecc, VPrimaryUser, ClsTransaksiRecurring.EnumStatus.Planed)
            DTAJ.Fill(tblJ)

            Me.tblHeader.Rows.Clear()
            Me.tblItem.Rows.Clear()
            For Each rowRecc In TblRecc.Rows
                With rowRecc
                    Dim IntStart As Integer = 0
                    If IsNumeric(.Start_Alarm_days_before) Then
                        IntStart = CInt(.Start_Alarm_days_before)
                    End If
                    If CDate(dtDate.AddDays(IntStart)).Date >= .Journal_Date.Date Then
                        Me.rowHeader = Me.tblHeader.NewtblHeaderRow
                        Me.rowHeader.Primary_Recurring = .Primary_Transaksi
                        Me.rowHeader.Recurring_No = ObjRec.GetNoTransasi(.Primary_Transaksi)
                        Me.rowHeader.Description = .Description_Header
                        Me.rowHeader.Primary_Journal = .Primary_Journal
                        Me.rowHeader.Journal_No = .Journal_No
                        Me.rowHeader.Journal_Date = .Journal_Date
                        Me.rowHeader.Journal_Description = .Description
                        Me.rowHeader.Journal_Remarks = .Remarks
                        Me.rowHeader.Journal_status = .Status
                        Me.rowHeader.Urutan = .No_Urut
                        Me.rowHeader.Urutan_Journal = Me.GetUrutanJournal(.Primary_Transaksi, .Primary_Item, .Ignore_Increment_on_skip)
                        Me.tblHeader.Rows.Add(Me.rowHeader)
                    End If
                End With
            Next

            If Me.tblHeader.Rows.Count > 0 Then
                For Each Me.rowHeader In Me.tblHeader.Rows
                    With Me.rowHeader
                        tblJ.DefaultView.RowFilter = Nothing
                        tblJ.DefaultView.RowFilter = "[Primary Transaksi] = " & .Primary_Recurring
                        If tblJ.DefaultView.Count > 0 Then
                            Dim dvJ As Data.DataRowView
                            For Each dvJ In tblJ.DefaultView
                                rowJ = dvJ.Row
                                Me.rowItem = Me.tblItem.NewtblItemRow
                                Me.rowItem.Primary_Transaksi = .Primary_Recurring
                                Me.rowItem.Primary_Journal = .Primary_Journal
                                Me.rowItem.Journal_No = .Journal_No
                                Me.rowItem.Primary_Akun = rowJ.Primary_Akun
                                Me.rowItem.Primary_Currency = ObjAcc.GetPrimaryCurrency(rowJ.Primary_Akun)
                                Me.rowItem.Account_No = ObjAcc.GetAccountCode(rowJ.Primary_Akun)
                                Me.rowItem.Account_Description = ObjAcc.GetDescription_ByPrimaryAcc(rowJ.Primary_Akun)
                                If rowJ.Increment_Style = ClsTransaksiRecurring.EnumIncrementStyle.No_Increment Then
                                    Me.rowItem.Debet = rowJ.Debet
                                    Me.rowItem.Kredit = rowJ.Kredit
                                    Me.rowItem.Debet_Original = rowJ.Debet_Original
                                    Me.rowItem.Credit_Original = rowJ.Credit_Original
                                ElseIf rowJ.Increment_Style = ClsTransaksiRecurring.EnumIncrementStyle.Fixed_Increment Then
                                    Dim Increment As Decimal = 0
                                    Increment = .Urutan_Journal * CDec(rowJ.Increment_Value)
                                    If rowJ.Debet <> 0 Then
                                        Me.rowItem.Debet = rowJ.Debet + Increment
                                        If rowJ.Ratio <> 0 Then
                                            Me.rowItem.Debet_Original = Me.rowItem.Debet / rowJ.Ratio
                                        Else
                                            Me.rowItem.Debet_Original = Me.rowItem.Debet
                                        End If
                                        Me.rowItem.Kredit = 0
                                        Me.rowItem.Credit_Original = 0
                                    Else
                                        Me.rowItem.Kredit = rowJ.Kredit + Increment
                                        If rowJ.Ratio <> 0 Then
                                            Me.rowItem.Credit_Original = Me.rowItem.Kredit / rowJ.Ratio
                                        Else
                                            Me.rowItem.Credit_Original = Me.rowItem.Kredit
                                        End If
                                        Me.rowItem.Debet = 0
                                        Me.rowItem.Debet_Original = 0
                                    End If
                                ElseIf rowJ.Increment_Style = ClsTransaksiRecurring.EnumIncrementStyle.Percentage_Increment Then
                                    Dim Increment As Decimal = 0
                                    Increment = (CDec(100 + rowJ.Increment_Value) / 100) ^ .Urutan_Journal
                                    If rowJ.Debet <> 0 Then
                                        Me.rowItem.Debet = rowJ.Debet * Increment
                                        If rowJ.Ratio <> 0 Then
                                            Me.rowItem.Debet_Original = Me.rowItem.Debet / rowJ.Ratio
                                        Else
                                            Me.rowItem.Debet_Original = Me.rowItem.Debet
                                        End If
                                        Me.rowItem.Kredit = 0
                                        Me.rowItem.Credit_Original = 0
                                    Else
                                        Me.rowItem.Kredit = rowJ.Kredit * Increment
                                        If rowJ.Ratio <> 0 Then
                                            Me.rowItem.Credit_Original = Me.rowItem.Kredit / rowJ.Ratio
                                        Else
                                            Me.rowItem.Credit_Original = Me.rowItem.Kredit
                                        End If
                                        Me.rowItem.Debet = 0
                                        Me.rowItem.Debet_Original = 0
                                    End If
                                End If
                                Me.rowItem.Remarks = rowJ.Remarks
                                Me.rowItem.Ratio = rowJ.Ratio
                                Me.rowItem.Primary_Dept = rowJ.primary_dept
                                Me.rowItem.Primary_Job = rowJ.primary_job
                                Me.tblItem.Rows.Add(Me.rowItem)
                            Next
                        End If
                    End With
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTARecc.Dispose()
            DTAJ.Dispose()
            TblRecc.Dispose()
            tblJ.Dispose()
            ObjCur = Nothing
            ObjAcc = Nothing
            ObjRec = Nothing
        End Try
    End Sub

    Private Function GetUrutanJournal(ByVal vPrimaryTransaksi As Double, ByVal PrimaryItem As Double, ByVal IgnoreIncrementOnSkip As Integer) As Integer
        Dim rowDgv As New DataGridViewRow
        Dim DTA As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journal_transaction_entriesTableAdapter
        Dim TBL As New DseTransaksiRecurring.tbltrans_recurring_journal_transaction_entriesDataTable
        Dim row As DseTransaksiRecurring.tbltrans_recurring_journal_transaction_entriesRow
        Try

            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.FillBy(TBL, vPrimaryTransaksi)
            Dim Urutan As Integer = 0
            If TBL.Rows.Count > 0 Then
                For intA As Integer = 0 To TBL.Rows.Count - 1
                    row = TBL.Rows(intA)
                    With row
                        If IgnoreIncrementOnSkip = 0 Then
                            Urutan = intA
                        Else
                            If .Status <> ClsTransaksiRecurring.EnumStatus.Skiped Then
                                Urutan += 1
                            End If
                        End If
                        If .Primary = PrimaryItem Then
                            Exit For
                        End If
                    End With
                Next
            End If

            'If Me.dgvHeader.Rows.Count > 0 Then
            '    For intA As Integer = 0 To Me.dgvHeader.Rows.Count - 1
            '        Dim intB As Integer = 0
            '        rowDgv = Me.dgvHeader.Rows(intA)
            '        With rowDgv
            '            If Me.chkIgnoreIncrementOnSkip.Checked = False Then
            '                Urutan = intA
            '            Else
            '                If .Cells(Me.ColTrans_Status.Index).Value <> Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Skiped) Then
            '                    Urutan += 1
            '                End If
            '            End If
            '            If intA = VRowIndex Then
            '                Exit For
            '            End If
            '        End With
            '    Next
            'End If
            Return Urutan

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
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

    Private Sub SplitContainer1_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles SplitContainer1.SizeChanged
        Try
            If Me.Splitter1_SplitterDistance <> 0 Then
                Me.SplitContainer1.SplitterDistance = Me.Splitter1_SplitterDistance
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnRecurring_Show_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRecurring_Show.Click

        Try

            If Me.SplitContainer1.Panel2Collapsed = True Then
                Me.SplitContainer1.Panel2Collapsed = False
                Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, True)

            Else
                Me.SplitContainer1.Panel2Collapsed = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHeader_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvHeader.CellBeginEdit

        Try
            If e.ColumnIndex <> Me.colModify.Index Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHeader_CellMouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvHeader.CellMouseDown

        Try
            If Me.dgvHeader.Rows.Count > 0 AndAlso e.RowIndex > -1 Then
                Me.rowIndex = e.RowIndex
                If e.RowIndex <> 0 Then
                    Me.tms_Create.Enabled = False
                    Me.tms_CreateOpen.Enabled = False
                    Me.tms_Skip.Enabled = False
                    Me.GoToRecurringJournalToolStripMenuItem.Enabled = True
                Else
                    Me.tms_Create.Enabled = True
                    Me.tms_CreateOpen.Enabled = True
                    Me.tms_Skip.Enabled = True
                    Me.GoToRecurringJournalToolStripMenuItem.Enabled = True

                    'If Me.dgvHeader.Item(Me.ColRecurringTrans_Status.Index, e.RowIndex).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Created) Then
                    '    Me.tms_Create.Enabled = False
                    '    Me.tms_CreateOpen.Enabled = False
                    '    Me.tms_Skip.Enabled = False
                    'End If

                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tms_Create_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tms_Create.Click
        Dim I As New FrmTransaksiRecuring
        Dim ObjRec As New ClsTransaksiRecurring
        Dim PrimaryTransaksi As Double = -1
        Dim ObjUser As New Core_Module.ClsUser
        Dim PrimaryUser As Double = -1

        Try
            PrimaryTransaksi = Me.dgvHeader.Item(Me.ColPrimaryRecurring.Index, Me.rowIndex).Value
            IRowIndex = Me.Fill_Recurring(I, PrimaryTransaksi, Me.rowIndex)
            If IRowIndex > -1 Then
                If create(I, False, IRowIndex) = True Then
                    Dim hasil As MsgBoxResult
                    hasil = MsgBox("Data saved to journal transaction, do you want to refresh your data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Me.dgvHeader.Rows.Clear()
                        Me.dgvItem.Rows.Clear()
                        Me.tblHeader.Rows.Clear()
                        Me.tblItem.Rows.Clear()
                        PrimaryUser = ObjUser.GetPrimary(Me.UserCode)
                        Me.FillHeaderItem(PrimaryUser, Date.Now)
                        Me.isiKeDGVHeader()
                        Me.rowIndex = 0
                        If Me.dgvHeader.Rows.Count > 0 Then
                            Dim NoJournal As String = ""
                            PrimaryTransaksi = Me.dgvHeader.Item(Me.ColPrimaryRecurring.Index, Me.rowIndex).Value
                            NoJournal = Me.dgvHeader.Item(Me.ColRecurringTrans_JournalNumber.Index, Me.rowIndex).Value
                            Me.isiKeDGVItem(PrimaryTransaksi, NoJournal)
                        End If
                    Else
                        MsgBox("the reminder recurring will be close. ", MsgBoxStyle.Information, "Information")
                        Me.Close()
                    End If
                    ' Me.dgvHeader.Item(Me.ColRecurringTrans_Status.Index, Me.rowIndex).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Created)
                End If
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            I.Dispose()
            ObjRec = Nothing
            ObjUser = Nothing
        End Try
    End Sub

    Private Sub tms_CreateOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tms_CreateOpen.Click
        Dim I As New FrmTransaksiRecuring
        Dim ObjRec As New ClsTransaksiRecurring
        Dim PrimaryTransaksi As Double = -1
        Dim ObjUser As New Core_Module.ClsUser
        Dim PrimaryUser As Double = -1

        Try
            PrimaryTransaksi = Me.dgvHeader.Item(Me.ColPrimaryRecurring.Index, Me.rowIndex).Value
            IRowIndex = Me.Fill_Recurring(I, PrimaryTransaksi, Me.rowIndex)
            If IRowIndex > -1 Then
                If create(I, True, IRowIndex) = True Then
                    Dim hasil As MsgBoxResult
                    hasil = MsgBox("Data saved to journal transaction, do you want to refresh your data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Me.dgvHeader.Rows.Clear()
                        Me.dgvItem.Rows.Clear()
                        Me.tblHeader.Rows.Clear()
                        Me.tblItem.Rows.Clear()
                        PrimaryUser = ObjUser.GetPrimary(Me.UserCode)
                        Me.FillHeaderItem(PrimaryUser, Date.Now)
                        Me.isiKeDGVHeader()
                        Me.rowIndex = 0
                        If Me.dgvHeader.Rows.Count > 0 Then
                            Dim NoJournal As String = ""
                            PrimaryTransaksi = Me.dgvHeader.Item(Me.ColPrimaryRecurring.Index, Me.rowIndex).Value
                            NoJournal = Me.dgvHeader.Item(Me.ColRecurringTrans_JournalNumber.Index, Me.rowIndex).Value
                            Me.isiKeDGVItem(PrimaryTransaksi, NoJournal)
                        End If
                        'Me.dgvHeader.Item(Me.ColRecurringTrans_Status.Index, Me.rowIndex).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Created)
                    Else
                        MsgBox("the reminder recurring will be close. ", MsgBoxStyle.Information, "Information")
                        Me.Close()
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            I.Dispose()
            I.Dispose()
            ObjRec = Nothing
            ObjUser = Nothing
        End Try
    End Sub

    Private Sub GoToRecurringJournalToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GoToRecurringJournalToolStripMenuItem.Click
        Dim I As New FrmTransaksiRecuring
        Dim ObjRec As New ClsTransaksiRecurring
        Dim PrimaryTransaksi As Double = -1
        Dim ObjUser As New Core_Module.ClsUser
        Dim PrimaryUser As Double = -1
        Dim ObjRecurring As New ClsTransaksiRecurring

        Try
            PrimaryTransaksi = Me.dgvHeader.Item(Me.ColPrimaryRecurring.Index, Me.rowIndex).Value
            If ObjRecurring.ApakahBolehEdit(PrimaryTransaksi, Me.UserCode, True, True) = True Then
                IRowIndex = Me.Fill_Recurring(I, PrimaryTransaksi, Me.rowIndex)
                If IRowIndex > -1 Then
                    I.btnEdit_Click(Nothing, Nothing)
                    I.FromReminder = True
                    I.UserCOde = Me.UserCode
                    I.ShowDialog()
                    'Me.dgvHeader.Rows.Clear()
                    'Me.dgvItem.Rows.Clear()
                    'Me.tblHeader.Rows.Clear()
                    'Me.tblItem.Rows.Clear()
                    'PrimaryUser = ObjUser.GetPrimary(Me.UserCode)
                    'Me.FillHeaderItem(PrimaryUser, Date.Now)
                    'Me.isiKeDGVHeader()
                    'Me.rowIndex = 0
                    'If Me.dgvHeader.Rows.Count > 0 Then
                    '    Dim NoJournal As String = ""
                    '    PrimaryTransaksi = Me.dgvHeader.Item(Me.ColPrimaryRecurring.Index, Me.rowIndex).Value
                    '    NoJournal = Me.dgvHeader.Item(Me.ColRecurringTrans_JournalNumber.Index, Me.rowIndex).Value
                    '    If PrimaryTransaksi > -1 Then
                    '        Me.isiKeDGVItem(PrimaryTransaksi, NoJournal)
                    '    End If
                    'End If
                End If
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            I.Dispose()
            ObjRec = Nothing
            ObjUser = Nothing
            ObjRecurring = Nothing
        End Try

    End Sub

    Private Sub tms_Skip_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tms_Skip.Click
        Dim I As New FrmTransaksiRecuring
        Dim ObjRec As New ClsTransaksiRecurring
        Dim PrimaryTransaksi As Double = -1
        Dim ObjUser As New Core_Module.ClsUser
        Dim PrimaryUser As Double = -1
        Try

            PrimaryTransaksi = Me.dgvHeader.Item(Me.ColPrimaryRecurring.Index, Me.rowIndex).Value
            IRowIndex = Me.Fill_Recurring(I, PrimaryTransaksi, Me.rowIndex)
            If IRowIndex > -1 Then
                If Skip(I, IRowIndex) = True Then
                    Dim hasil As MsgBoxResult
                    hasil = MsgBox("Data has skiped, do you want to refresh your data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Me.dgvHeader.Rows.Clear()
                        Me.dgvItem.Rows.Clear()
                        Me.tblHeader.Rows.Clear()
                        Me.tblItem.Rows.Clear()
                        PrimaryUser = ObjUser.GetPrimary(Me.UserCode)
                        Me.FillHeaderItem(PrimaryUser, Date.Now)
                        Me.isiKeDGVHeader()
                        Me.rowIndex = 0
                        If Me.dgvHeader.Rows.Count > 0 Then
                            Dim NoJournal As String = ""
                            PrimaryTransaksi = Me.dgvHeader.Item(Me.ColPrimaryRecurring.Index, Me.rowIndex).Value
                            NoJournal = Me.dgvHeader.Item(Me.ColRecurringTrans_JournalNumber.Index, Me.rowIndex).Value
                            Me.isiKeDGVItem(PrimaryTransaksi, NoJournal)
                        End If
                        'Me.dgvHeader.Item(Me.ColRecurringTrans_Status.Index, Me.rowIndex).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Skiped)
                    Else
                        MsgBox("the reminder recurring will be close. ", MsgBoxStyle.Information, "Information")
                        Me.Close()
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            I.Dispose()
            ObjRec = Nothing
            ObjUser = Nothing
        End Try
    End Sub

    Private Function Fill_Recurring(ByVal i As FrmTransaksiRecuring, ByVal PrimaryTransaksi As Double, ByVal vrowIndex As Integer) As Integer
        Dim IRowIndexX As Integer = 0
        Try
            'I.KosongkanField()
            i.NonAktifkanDetail()
            i.FromReminder = True
            i.UserCOde = Me.UserCode
            i.cboDGV.Items.Clear()
            i.cboDGV.Items.Add("No Increment")
            i.cboDGV.Items.Add("Fixed Increment")
            i.cboDGV.Items.Add("Percentage Increment")

            i.cboStatus.Items.Clear()
            i.cboStatus.Items.Add("Planed")
            i.cboStatus.Items.Add("Created")
            i.cboStatus.Items.Add("Skiped")
            i.cboStatus.Items.Add("Lock")

            i.ij.PrimaryPP = PrimaryTransaksi
            i.btnFind_Click(Nothing, Nothing)
            Dim JournalNo As String = ""
            Dim rowDgv As New DataGridViewRow
            JournalNo = Me.dgvHeader.Item(Me.ColRecurringTrans_JournalNumber.Index, vrowIndex).Value.ToString
            For Each rowDgv In i.dgvEntries.Rows
                With rowDgv
                    If JournalNo = rowDgv.Cells(i.ColTrans_JournalNumber.Index).Value.ToString Then
                        IRowIndexX = rowDgv.Index
                        Exit For
                    End If
                End With
            Next

            Return IRowIndexX
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function create(ByVal I As FrmTransaksiRecuring, ByVal show As Boolean, ByVal rowIndex As Integer) As Boolean
        Dim DTAUpdate As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journal_transaction_entriesTableAdapter
        Try
            DTAUpdate.Connection.ConnectionString = Core_Module.GetConstring
            I.rowIndex = rowIndex
            If Not String.IsNullOrEmpty(I.dgvEntries.Item(I.ColTrans_JournalNumber.Index, I.rowIndex).Value) Then
                If I.ApakahDataBenarJournal(I.TxtUnbalance.Text, I.dgvItem) = True Then
                    Dim Urutan As Integer = 0
                    Urutan = I.GetUrutanJournal(I.rowIndex)
                    Dim PrimaryJournal As Double = -1
                    PrimaryJournal = I.dgvEntries.Item(I.ColTrans_PrimaryJournal.Index, I.rowIndex).Value
                    If I.IsiKeJournal(I.rowIndex, Urutan, PrimaryJournal, show, ClsTransaksiRecurring.EnumAction.Create) = True Then
                        'I.dgvEntries.Item(I.ColTrans_Status.Index, I.rowIndex).Value = I.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Created)
                        If I.Limit > (I.dgvEntries.Rows.Count) AndAlso I.rowIndex = I.dgvEntries.Rows.Count - 1 Then
                            I.intBatasBaris = I.dgvEntries.Rows.Count + 1
                            I.TambahkanBarisEntries(I.Limit)
                            Dim IntRowX As Integer = 0
                            IntRowX = CInt(I.dgvEntries.Rows.Count - 1)
                            If String.IsNullOrEmpty(I.dgvEntries.Item(I.ColTrans_Remarks.Index, IntRowX).Value) Then
                                I.dgvEntries.Item(I.ColTrans_Remarks.Index, IntRowX).Value = ""
                            End If
                            DTAUpdate.Insert(I.TxtPrimaryJournal.Text, _
                                             I.dgvEntries.Item(I.ColTrans_PrimaryJournal.Index, IntRowX).Value, _
                                             I.dgvEntries.Item(I.ColTrans_JournalNumber.Index, IntRowX).Value, _
                                             I.dgvEntries.Item(I.ColTrans_JournalDate.Index, IntRowX).Value, _
                                             I.dgvEntries.Item(I.ColTrans_Description.Index, IntRowX).Value, _
                                             I.dgvEntries.Item(I.ColTrans_Remarks.Index, IntRowX).Value, _
                                             ClsTransaksiRecurring.EnumStatus.Planed, _
                                             I.dgvEntries.Item(I.ColTrans_NoUrut.Index, IntRowX).Value)
                            'I.ObjRecurring.CariDetail(I.TxtPrimaryJournal.Text, I.dgvItem, I)
                            'I.SalinkeField()
                        End If
                        Return True
                    Else
                        MsgBox("Failed to save data.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAUpdate.Dispose()
        End Try
    End Function

    Private Function Skip(ByVal I As FrmTransaksiRecuring, ByVal rowIndex As Integer) As Boolean
        Dim DTAU As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journal_transaction_entriesTableAdapter
        Dim DTAUpdate As New DseTransaksiRecurringTableAdapters.Khusus_Update_Trans_EntriesTableAdapter
        Try
            DTAU.Connection.ConnectionString = Core_Module.GetConstring
            DTAUpdate.Connection.ConnectionString = DTAU.Connection.ConnectionString

            If Not String.IsNullOrEmpty(I.dgvEntries.Item(I.ColTrans_JournalNumber.Index, rowIndex).Value) Then
                DTAUpdate.UpdateQuery(I.dgvEntries.Item(I.ColTrans_PrimaryJournal.Index, rowIndex).Value, _
                                      I.dgvEntries.Item(I.ColTrans_JournalNumber.Index, rowIndex).Value, _
                                      I.dgvEntries.Item(I.ColTrans_JournalDate.Index, rowIndex).Value, _
                                      I.dgvEntries.Item(I.ColTrans_Description.Index, rowIndex).Value, _
                                      ClsTransaksiRecurring.EnumStatus.Skiped, _
                                      I.dgvEntries.Item(I.ColTrans_NoUrut.Index, rowIndex).Value, _
                                      I.dgvEntries.Item(I.ColTrans_JournalNumber.Index, rowIndex).Value)

                If I.Limit > (I.dgvEntries.Rows.Count) AndAlso rowIndex = I.dgvEntries.Rows.Count - 1 Then
                    I.intBatasBaris = I.dgvEntries.Rows.Count + 1
                    I.TambahkanBarisEntries(I.Limit)
                    Dim IntRowX As Integer = 0
                    IntRowX = CInt(I.dgvEntries.Rows.Count - 1)
                    If String.IsNullOrEmpty(I.dgvEntries.Item(I.ColTrans_Remarks.Index, IntRowX).Value) Then
                        I.dgvEntries.Item(I.ColTrans_Remarks.Index, IntRowX).Value = ""
                    End If
                    DTAU.Insert(I.TxtPrimaryJournal.Text, _
                                          I.dgvEntries.Item(I.ColTrans_PrimaryJournal.Index, IntRowX).Value, _
                                          I.dgvEntries.Item(I.ColTrans_JournalNumber.Index, IntRowX).Value, _
                                          I.dgvEntries.Item(I.ColTrans_JournalDate.Index, IntRowX).Value, _
                                          I.dgvEntries.Item(I.ColTrans_Description.Index, IntRowX).Value, _
                                          I.dgvEntries.Item(I.ColTrans_Remarks.Index, IntRowX).Value, _
                                          ClsTransaksiRecurring.EnumStatus.Planed, _
                                          I.dgvEntries.Item(I.ColTrans_NoUrut.Index, IntRowX).Value)
                End If
                'I.ObjRecurring.CariDetail(I.TxtPrimaryJournal.Text, Me.dgvItem, Me)
                'I.SalinkeField()

                'Me.dgvEntries.Item(Me.ColTrans_Status.Index, Me.rowIndex).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Skiped)
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAU.Dispose()
            DTAUpdate.Dispose()
        End Try
    End Function


    Private Sub btnCreate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCreate.Click
        Dim I As New FrmTransaksiRecuring
        Dim ObjRec As New ClsTransaksiRecurring
        Dim PrimaryTransaksi As Double = -1
        Dim ObjUser As New Core_Module.ClsUser
        Dim PrimaryUser As Double = -1
        Try
            If Me.dgvHeader.Rows.Count > 0 Then
                Dim ada As Integer = 0
                Dim rowDGV As New DataGridViewRow
                For Each rowDGV In Me.dgvHeader.Rows
                    With rowDGV
                        If .Cells(Me.colModify.Index).Value = True Then
                            PrimaryTransaksi = .Cells(Me.ColPrimaryRecurring.Index).Value
                            IRowIndex = Me.Fill_Recurring(I, PrimaryTransaksi, rowDGV.Index)
                            If IRowIndex > -1 Then
                                If Me.create(I, False, IRowIndex) = True Then
                                    ada += 1
                                    ' Me.dgvHeader.Item(Me.ColRecurringTrans_Status.Index, Me.rowIndex).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Skiped)
                                End If
                            End If
                        End If
                    End With
                Next

                If ada = 0 Then
                    MsgBox("Please specify at least 1 item to create journal.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If


                If ada > 0 Then
                    Dim hasil As MsgBoxResult
                    hasil = MsgBox("Data saved to journal transaction, do you want to refresh your data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        PrimaryUser = ObjUser.GetPrimary(Me.UserCode)
                        Me.FillHeaderItem(PrimaryUser, Date.Now)
                        BlnUnSelect = True
                        Me.isiKeDGVHeader()
                        BlnUnSelect = False
                        Me.rowIndex = 0
                        If Me.dgvHeader.Rows.Count > 0 Then
                            Dim NoJournal As String = ""
                            PrimaryTransaksi = Me.dgvHeader.Item(Me.ColPrimaryRecurring.Index, Me.rowIndex).Value
                            NoJournal = Me.dgvHeader.Item(Me.ColRecurringTrans_JournalNumber.Index, Me.rowIndex).Value
                            Me.isiKeDGVItem(PrimaryTransaksi, NoJournal)
                        End If
                    Else
                        MsgBox("the reminder recurring will be close. ", MsgBoxStyle.Information, "Information")
                        Me.Close()
                    End If
                End If
              
                'Dim PrimaryTransaksi As Double = -1
                'Dim NoJournal As String = ""
                'PrimaryTransaksi = Me.dgvHeader.Item(Me.ColPrimaryRecurring.Index, Me.rowIndex).Value
                'NoJournal = Me.dgvHeader.Item(Me.ColRecurringTrans_JournalNumber.Index, Me.rowIndex).Value
                'Me.isiKeDGVItem(PrimaryTransaksi, NoJournal)

                'PrimaryTransaksi = Me.dgvHeader.Item(Me.ColPrimaryRecurring.Index, Me.rowIndex).Value
                'IRowIndex = Me.Fill_Recurring(I, PrimaryTransaksi)
                'If IRowIndex > 0 Then
                '    If Skip(I, IRowIndex) = True Then
                '        Me.dgvHeader.Item(Me.ColRecurringTrans_Status.Index, Me.rowIndex).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Skiped)
                '    End If
                'End If
                'Else
                '    MsgBox("Please checked minimal 1 coloumn modify to create journal. ", MsgBoxStyle.Information, "Information")
                '    Exit Sub
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            I.Dispose()
            ObjRec = Nothing
            ObjUser = Nothing
        End Try
    End Sub

    Private Sub btnSkip_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSkip.Click
        Dim I As New FrmTransaksiRecuring
        Dim ObjRec As New ClsTransaksiRecurring
        Dim PrimaryTransaksi As Double = -1
        Dim ObjUser As New Core_Module.ClsUser
        Dim PrimaryUser As Double = -1
        Try
            If Me.dgvHeader.Rows.Count > 0 Then
                Dim ada As Integer = 0
                Dim rowDGV As New DataGridViewRow
                For Each rowDGV In Me.dgvHeader.Rows
                    With rowDGV
                        If .Cells(Me.colModify.Index).Value = True Then
                            PrimaryTransaksi = .Cells(Me.ColPrimaryRecurring.Index).Value
                            IRowIndex = Me.Fill_Recurring(I, PrimaryTransaksi, rowDGV.Index)
                            If IRowIndex > -1 Then
                                If Skip(I, IRowIndex) = True Then
                                    ada += 1
                                    ' Me.dgvHeader.Item(Me.ColRecurringTrans_Status.Index, Me.rowIndex).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Skiped)
                                End If
                            End If
                        End If
                    End With
                Next

                If ada = 0 Then
                    MsgBox("Please specify at least 1 item to skip journal.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                If ada > 0 Then
                    Dim hasil As MsgBoxResult
                    hasil = MsgBox("Data has skiped, do you want to refresh your data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Me.dgvHeader.Rows.Clear()
                        Me.dgvItem.Rows.Clear()
                        Me.tblHeader.Rows.Clear()
                        Me.tblItem.Rows.Clear()
                        PrimaryUser = ObjUser.GetPrimary(Me.UserCode)
                        Me.FillHeaderItem(PrimaryUser, Date.Now)
                        Me.BlnUnSelect = True
                        Me.isiKeDGVHeader()
                        Me.BlnUnSelect = False
                        Me.rowIndex = 0
                        If Me.dgvHeader.Rows.Count > 0 Then
                            Dim NoJournal As String = ""
                            PrimaryTransaksi = Me.dgvHeader.Item(Me.ColPrimaryRecurring.Index, Me.rowIndex).Value
                            NoJournal = Me.dgvHeader.Item(Me.ColRecurringTrans_JournalNumber.Index, Me.rowIndex).Value
                            Me.isiKeDGVItem(PrimaryTransaksi, NoJournal)
                        End If
                    Else
                        MsgBox("the reminder recurring will be close. ", MsgBoxStyle.Information, "Information")
                        Me.Close()
                    End If
                End If
              
                'PrimaryTransaksi = Me.dgvHeader.Item(Me.ColPrimaryRecurring.Index, Me.rowIndex).Value
                'IRowIndex = Me.Fill_Recurring(I, PrimaryTransaksi)
                'If IRowIndex > 0 Then
                '    If Skip(I, IRowIndex) = True Then
                '        Me.dgvHeader.Item(Me.ColRecurringTrans_Status.Index, Me.rowIndex).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Skiped)
                '    End If
                'End If
                'MsgBox("Please checked minimal 1 coloumn modify to skip journal. ", MsgBoxStyle.Information, "Information")
                'Exit Sub
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            I.Dispose()
            ObjRec = Nothing
            ObjUser = Nothing
        End Try
    End Sub

    Private Sub btnCreateAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCreateAll.Click
        Dim I As New FrmTransaksiRecuring
        Dim ObjRec As New ClsTransaksiRecurring
        Dim PrimaryTransaksi As Double = -1
        Dim ObjUser As New Core_Module.ClsUser
        Dim PrimaryUser As Double = -1
        Try
            If Me.dgvHeader.Rows.Count > 0 Then
                Dim rowDGV As New DataGridViewRow
                For Each rowDGV In Me.dgvHeader.Rows
                    With rowDGV
                        ' If .Cells(Me.colModify.Index).Value = True Then
                        PrimaryTransaksi = .Cells(Me.ColPrimaryRecurring.Index).Value
                        IRowIndex = Me.Fill_Recurring(I, PrimaryTransaksi, rowDGV.Index)
                        If IRowIndex > -1 Then
                            If Me.create(I, False, IRowIndex) = True Then
                                ' Me.dgvHeader.Item(Me.ColRecurringTrans_Status.Index, Me.rowIndex).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Skiped)
                            End If
                        End If
                        ' End If
                    End With
                Next

                Me.dgvHeader.Rows.Clear()
                Me.dgvItem.Rows.Clear()
                Me.tblHeader.Rows.Clear()
                Me.tblItem.Rows.Clear()
                PrimaryUser = ObjUser.GetPrimary(Me.UserCode)
                Me.FillHeaderItem(PrimaryUser, Date.Now)
                Me.isiKeDGVHeader()
                Me.rowIndex = 0
                If Me.dgvHeader.Rows.Count > 0 Then
                    Dim NoJournal As String = ""
                    PrimaryTransaksi = Me.dgvHeader.Item(Me.ColPrimaryRecurring.Index, Me.rowIndex).Value
                    NoJournal = Me.dgvHeader.Item(Me.ColRecurringTrans_JournalNumber.Index, Me.rowIndex).Value
                    Me.isiKeDGVItem(PrimaryTransaksi, NoJournal)
                End If

                'PrimaryTransaksi = Me.dgvHeader.Item(Me.ColPrimaryRecurring.Index, Me.rowIndex).Value
                'IRowIndex = Me.Fill_Recurring(I, PrimaryTransaksi)
                'If IRowIndex > 0 Then
                '    If Skip(I, IRowIndex) = True Then
                '        Me.dgvHeader.Item(Me.ColRecurringTrans_Status.Index, Me.rowIndex).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Skiped)
                '    End If
                'End If
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            I.Dispose()
            ObjRec = Nothing
            ObjUser = Nothing
        End Try
    End Sub

    Private Sub dgvHeader_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvHeader.CellValueChanged
        Try
            If e.RowIndex > -1 Then
                If e.ColumnIndex = Me.colModify.Index Then
                    If e.RowIndex > 0 AndAlso Me.dgvHeader.Item(e.ColumnIndex, e.RowIndex).Value = True AndAlso Me.dgvHeader.Item(e.ColumnIndex, e.RowIndex - 1).Value = False Then
                        MsgBox("You can not skip or create " & Me.dgvHeader.Item(Me.ColRecurringTrans_JournalNumber.Index, e.RowIndex).Value & _
                               " before " & Me.dgvHeader.Item(Me.ColRecurringTrans_JournalNumber.Index, e.RowIndex - 1).Value & ".", MsgBoxStyle.Information, "Information")
                        Me.dgvHeader.Item(e.ColumnIndex, e.RowIndex).Value = False
                        Exit Sub
                    End If

                    If e.RowIndex < Me.dgvHeader.Rows.Count - 1 AndAlso Me.dgvHeader.Item(e.ColumnIndex, e.RowIndex).Value = False AndAlso Me.dgvHeader.Item(e.ColumnIndex, e.RowIndex + 1).Value = True Then
                        MsgBox("You can not skip or create " & Me.dgvHeader.Item(Me.ColRecurringTrans_JournalNumber.Index, e.RowIndex).Value & _
                               " before " & Me.dgvHeader.Item(Me.ColRecurringTrans_JournalNumber.Index, e.RowIndex + 1).Value & ".", MsgBoxStyle.Information, "Information")
                        Me.dgvHeader.Item(e.ColumnIndex, e.RowIndex).Value = True
                        Exit Sub
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHeader_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvHeader.SelectionChanged

        Try
            If Me.BlnUnSelect = False AndAlso Me.dgvHeader.SelectedRows.Count > 0 AndAlso Me.dgvHeader.Rows.Count > 0 AndAlso Me.rowIndex >= 0 Then
                Dim PrimaryTransaksi As Double = -1
                Dim NoJournal As String = ""
                If IsNumeric(Me.dgvHeader.Rows(Me.rowIndex).Cells(Me.ColPrimaryRecurring.Index).Value) AndAlso _
                             CDec(Me.dgvHeader.Rows(Me.rowIndex).Cells(Me.ColPrimaryRecurring.Index).Value) >= 0 Then
                    PrimaryTransaksi = Me.dgvHeader.Rows(Me.rowIndex).Cells(Me.ColPrimaryRecurring.Index).Value
                    NoJournal = Me.dgvHeader.Rows(Me.rowIndex).Cells(Me.ColRecurringTrans_JournalNumber.Index).Value
                    Me.isiKeDGVItem(PrimaryTransaksi, NoJournal)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub llbCheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCheckAll.LinkClicked
        Dim Row As New DataGridViewRow

        Try
            If Me.dgvHeader.Rows.Count > 0 Then
                For Each Row In Me.dgvHeader.Rows
                    With Row
                        If .Cells(Me.colModify.Index).Value = False Then
                            .Cells(Me.colModify.Index).Value = True
                        End If
                    End With
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUncheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUncheckAll.LinkClicked
        Dim Row As New DataGridViewRow

        Try
            If Me.dgvHeader.Rows.Count > 0 Then
                Dim IntA As Integer = 0
                IntA = Me.dgvHeader.Rows.Count - 1
                While IntA >= 0
                    Row = Me.dgvHeader.Rows(IntA)
                    With Row
                        .Cells(Me.colModify.Index).Value = False
                        IntA -= 1
                    End With
                End While
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click

        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class