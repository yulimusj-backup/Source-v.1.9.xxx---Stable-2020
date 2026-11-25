Option Explicit On
Option Strict Off

Imports Core_Module, System.Windows.Forms

Public Class FrmTransaksiRecuring
    Public UserCOde As String = ""
    Public BackcolorAll As Color = Core_Module.GetSMESSBackColorAll
    Public TipeJournal As ClsTransaksiJournal.EnumTipeJournal = ClsTransaksiJournal.EnumTipeJournal.AdjustmenJournal
    Public TipeJournal2 As ClsTransaksiJournal.EnumTipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.Jurnal_NoGrouping
    Dim Buffer As String = ""
    Dim DefaultPrecision As Integer = 0
    Dim Baru As Boolean
    Dim PrimaryBaseCurrency As Double = 1
    Dim ObjCD As New ClsCekData
    Dim ObjAkun As New ClsMasterAkun
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjTaxRate As New ClsTransaksiSalesInvoiceTaxRate
    Dim ObjDept As New ClsMasterDepartement
    Dim ObjJob As New Core_Module.ClsMasterJob
    Dim ObjDefaultValue As New Core_Module.ClsTransaksiSetting
    Dim ObjLicense As New ClsLicense
    Dim ObjRCR As New ClsTransaksiRecurring
    Dim ObjRecurring As New ClsTransaksiRecurring
    Dim ObjClosing As New ClsTransaksiClosing
    Dim ObjJournal As New ClsTransaksiJournal
    Public FromReminder As Boolean = False
    Public ij As New FrmTransaksiRecurring_Cari
    Public Dgv As New DataGridView

    Dim BufferTrans As String = ""
    Dim BufferReminder As String = ""

    'Dim BufferLimit As Integer = 0
    Dim BufferTglAwal As Date = Now.Date
    Dim BufferTglAkhir As Date = Now.Date
    Public cboDGV As New DataGridViewComboBoxCell
    Public cboStatus As New DataGridViewComboBoxCell

    Dim BerhasilDelete As Boolean = False

    Public Frequency As ClsTransaksiRecurring.EnumFrequency
    Public EveryMonth As ClsTransaksiRecurring.EnumEvery_Month
    Public EveryWeek As ClsTransaksiRecurring.EnumEvery_Week
    Public Status As ClsTransaksiRecurring.EnumStatus
    Public DeleteAllBerhasil As Boolean = False
    Public AdaJournalBaru As Boolean = False
    Public rowIndex As Integer = -1
    Dim rowIndexReminder As Integer = -1
    Dim rowIndexItem As Integer = -1
    'Dim UrutanJournal As Integer = 0
    Public Limit As Integer = 10
    Public intBatasBaris As Integer = 10
    Dim dgvItemHapus As Boolean = False
    Dim dgvReminderHapus As Boolean = False

    Private Enum enumEntryDari As Integer
        NilaiOriginal
        NilaiRupiah
        Kurs
    End Enum

    Private Sub FrmTransaksiRecuring_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

        Try
            If Me.TxtRemarks.Enabled = False AndAlso IsNumeric(Me.TxtPrimaryJournal.Text) AndAlso CDbl(Me.TxtPrimaryJournal.Text) > -1 Then
                Dim PrimaryTransaksi As Double = -1
                PrimaryTransaksi = Me.TxtPrimaryJournal.Text
                Me.KosongkanField()
                Me.ObjRecurring.CariDetail(PrimaryTransaksi, Me.dgvItem, Me)
                Me.SalinkeField()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub FrmTransaksiRecuring_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Try
            Me.dgvItem.AllowUserToAddRows = False
            Me.dgvReminder.AllowUserToAddRows = False
            If Me.TxtRemarks.Enabled = True Then
                Me.BringToFront()
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Me.dgvItem.AllowUserToAddRows = False
                    Me.dgvReminder.AllowUserToAddRows = False
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinkeObjek()
                        If Baru = True Then
                            If ObjRecurring.ApakahDataBenar1(Baru) = True Then
                                'simpan baru
                                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                                If Me.ObjRecurring.SimpanBaru(Me.dgvItem, Me.UserCOde, Me) = True Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                    Me.SalinkeField()
                                    Me.NonAktifkanDetail()
                                    Me.btnFind.Focus()
                                Else
                                    Me.dgvItem.AllowUserToAddRows = True
                                    Me.dgvReminder.AllowUserToAddRows = True
                                    e.Cancel = True
                                End If
                            Else
                                Me.dgvItem.AllowUserToAddRows = True
                                Me.dgvReminder.AllowUserToAddRows = True
                                e.Cancel = True
                            End If
                        Else
                            If ObjRecurring.ApakahBolehEdit(Me.TxtPrimaryJournal.Text, UserCOde, True, True) = True Then
                                'simpan perubahan
                                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                                If Me.ObjRecurring.SimpanPerubahan(Me.dgvItem, Me.UserCOde, Me) = True Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                    Me.SalinkeField()
                                    Me.NonAktifkanDetail()
                                    Me.btnEdit.Focus()
                                Else
                                    Me.dgvItem.AllowUserToAddRows = True
                                    Me.dgvReminder.AllowUserToAddRows = True
                                    e.Cancel = True
                                End If
                            Else
                                Me.dgvItem.AllowUserToAddRows = True
                                Me.dgvReminder.AllowUserToAddRows = True
                                e.Cancel = True
                            End If
                        End If
                    Else
                        Me.dgvItem.AllowUserToAddRows = True
                        Me.dgvReminder.AllowUserToAddRows = True
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
            Me.dgvItem.AllowUserToAddRows = True
            Me.dgvReminder.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub FrmTransaksiRecuring_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus

        Try
            If Me.TxtRemarks.Enabled = False AndAlso IsNumeric(Me.TxtPrimaryJournal.Text) AndAlso CDbl(Me.TxtPrimaryJournal.Text) > -1 Then
                Dim PrimaryTransaksi As Double = -1
                PrimaryTransaksi = Me.TxtPrimaryJournal.Text
                Me.KosongkanField()
                Me.ObjRecurring.CariDetail(PrimaryTransaksi, Me.dgvItem, Me)
                Me.SalinkeField()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    'status item:
    '1. planned/pending
    '2. created
    '3. locked
    '4. skipped

    'fungsi menu:

    '1. kalo item transaction yang dihapuskan bukan item created terakhir, maka dia jadinya skip, bukan planned. Klo item tersebut adalah item terakhir, maka berubah jadi planned
    '2. item skip boleh create asal periode tersebut tidak ada revaluastion, end of year journal, dan transaksi locking akunting
    '3. semua perubahan jadwal, misal dari harian, ke mingguan -> hanya dilakukan pada item planned
    '4. kalo ada item jurnal yang diganti tanggalnya, transaksi selanjutnya tetap merujuk berdasarkan jadwal, bukan tanggal aktual jurnal sebelumnya
    '5. setiap item recurring pasti memliki 10 atau kurang item transaksi yang planned, jika menambahkan 1 item jurnal yang dimasukkan
    '   beneran sebagai jurnal, pastikan tambahkan 1 item planned, kecuali kalo memang limitnya tidak mengijikan ada 10 item
    '

    Private Sub FrmTransaksiRecurring_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ''Me.cboJournalTipe.SelectedIndex = Me.TipeJournal
        Dim NullValue As String = ""

        Try
            Me.ShowIcon = False
            If FromReminder = False Then
                Me.dgvItem.AllowUserToAddRows = False
                Me.dgvReminder.AllowUserToAddRows = False
                Me.Baru = True
                Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackcolorAll
                Me.dgvReminder.AlternatingRowsDefaultCellStyle.BackColor = Me.BackcolorAll
                Me.dgvEntries.AlternatingRowsDefaultCellStyle.BackColor = Me.BackcolorAll

                cboDGV.Items.Clear()
                Me.cboDGV.Items.Add("No Increment")
                Me.cboDGV.Items.Add("Fixed Increment")
                Me.cboDGV.Items.Add("Percentage Increment")

                cboStatus.Items.Clear()
                Me.cboStatus.Items.Add("Planed")
                Me.cboStatus.Items.Add("Created")
                Me.cboStatus.Items.Add("Skiped")
                Me.cboStatus.Items.Add("Lock")

                Me.dgvItem.Rows.Clear()
                Me.dgvReminder.Rows.Clear()
                Me.dgvEntries.Rows.Clear()
                Me.KosongkanField()
                Me.NonAktifkanDetail()



                'Me.dgvEntries.SortOrder = False
                'set null value untuk dgv
                Me.DefaultPrecision = Me.ObjCur.GetDecimal_DefaultCurrency
                NullValue = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                Me.ColDebet.DefaultCellStyle.NullValue = NullValue
                Me.ColKredit.DefaultCellStyle.NullValue = NullValue
                Me.TxtUnbalance.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                Me.txtUnbalanceIncrement.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)

                'jika tidak ada hak akses atas dept, maka kolom dept ditutup
                'If Me.ObjHakAkses.ApakahSudahAda("FrmMasterDepartement") = False Then
                Me.ColDept.Visible = False
                'End If

                'jika tidak ada hak akses atas job, maka kolom job ditutup
                'If Me.ObjHakAkses.ApakahSudahAda("FrmMasterJob") = False Then
                Me.ColJob.Visible = False
                'End Iff
                Me.TxtRefNo.Text = ""
            Else
                Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackcolorAll
                Me.dgvReminder.AlternatingRowsDefaultCellStyle.BackColor = Me.BackcolorAll
                Me.dgvEntries.AlternatingRowsDefaultCellStyle.BackColor = Me.BackcolorAll
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryJournal.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimaryJournal.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimaryJournal.BackColor, False)
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, False, True, True, True, True, True, True)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None

            Me.dgvItem.AllowUserToDeleteRows = True
            Me.dgvItem.ReadOnly = False
            Me.dgvReminder.AllowUserToDeleteRows = True
            Me.dgvReminder.ReadOnly = False
            Me.ObjCD.SetNumber(Me.dgvReminder)

            If Me.dgvItem.Rows.Count > 0 AndAlso Me.rowIndexItem > 0 Then
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            Else
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If

            If Baru = True Then
                Me.chkActive.Checked = True
                Me.chkLimit.Checked = True
                Me.cboFrequency.SelectedIndex = 0
                Me.cboEveryMonth.SelectedIndex = 0
                Me.cboEveryWeek.SelectedIndex = 0
            End If

            Me.chkLimit_CheckedChanged(Nothing, Nothing)
            Me.cboFrequency_SelectedIndexChanged(Nothing, Nothing)

            Me.TabControl1.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, True, True, True, True, False, True, False)
            Me.CancelButton = Me.btnExit

            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True
            Me.dgvReminder.AllowUserToDeleteRows = False
            Me.dgvReminder.ReadOnly = True
            Me.tms_Create.Enabled = False
            Me.tms_Delete.Enabled = False
            Me.tms_DeleteAll.Enabled = False
            Me.tms_Edit.Enabled = False
            Me.tms_Lock.Enabled = False
            Me.tms_Skip.Enabled = False
            Me.tms_Update.Enabled = False
            Me.tms_UpdateAll.Enabled = False
            Me.cboEveryWeek.Enabled = False
            Me.cboEveryMonth.Enabled = False
            Me.cboFrequency.Enabled = False


            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.BackColor, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub KosongkanField()
        Try
            Me.dgvItem.Rows.Clear()
            Me.dgvEntries.Rows.Clear()
            Me.dgvReminder.Rows.Clear()
            Me.ObjCD.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, False)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, False)
            Me.TxtPrimaryJournal.Text = "-1"
            Me.TxtUnbalance.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            Me.txtUnbalanceIncrement.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            Me.chkActive.Checked = True
            Me.chkIgnoreIncrementOnSkip.Checked = False
            Me.chkLimit.Checked = False
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            'Me.cboFrequency.SelectedIndex = 0
            Me.cboEveryWeek.Enabled = False
            Me.cboEveryMonth.Enabled = False
            'Me.cboEveryMonth.SelectedIndex = 0
            'Me.cboEveryWeek.SelectedIndex = 0
            Me.chkActive.Checked = True
            Me.chkLimit.Checked = True
            Me.AdaJournalBaru = False
            Me.DeleteAllBerhasil = False
            Me.BerhasilDelete = False
            Me.rowIndexItem = -1
            Me.rowIndex = -1
            Me.rowIndexReminder = -1
            Me.dgvItemHapus = False
            Me.dgvReminderHapus = False
            Me.ObjCD.SetNumber(Me.dgvReminder)
            ''Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''Fungsi DGV
    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try


            Me.Buffer = ""
            If e.ColumnIndex <> Me.ColAccount.Index Then
                If e.ColumnIndex <> Me.ColRemarks.Index And e.ColumnIndex <> Me.ColOriginalCurrency.Index And e.ColumnIndex <> Me.ColDescAccount.Index Then

                    If e.ColumnIndex = Me.ColDebet.Index Or e.ColumnIndex = Me.ColKredit.Index Or _
                    e.ColumnIndex = Me.ColDebet_OriginalCurr.Index Or e.ColumnIndex = Me.ColCredit_OriginalCurrency.Index Or _
                    e.ColumnIndex = Me.ColExchangeRate.Index Or e.ColumnIndex = Me.ColJob.Index Or _
                    e.ColumnIndex = Me.ColDept.Index Or e.ColumnIndex = Me.ColAcc_IncrementValue.Index Then
                        'jika belum diisikan account codenya, maka tidak bisa edit dept dan job
                        If e.ColumnIndex = Me.ColDept.Index Or e.ColumnIndex = Me.ColJob.Index Then
                            If Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value = -1 Then
                                e.Cancel = True
                                Exit Sub
                            End If
                        End If

                        If e.ColumnIndex = Me.ColAcc_IncrementValue.Index Then
                            If Not IsNumeric(Me.dgvItem.Item(Me.ColAcc_IncrementValue.Index, e.RowIndex).Value) Then
                                e.Cancel = True
                                Exit Sub
                            End If
                            If Me.dgvItem.Item(Me.ColAcc_IncrementStyle.Index, e.RowIndex).Value = Me.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.No_Increment) Then
                                e.Cancel = True
                                Exit Sub
                            End If
                        End If

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
                            If e.ColumnIndex = Me.ColDebet.Index Or _
                            e.ColumnIndex = Me.ColDebet_OriginalCurr.Index Or _
                            e.ColumnIndex = Me.ColCredit_OriginalCurrency.Index Or _
                            e.ColumnIndex = Me.ColKredit.Index Or _
                            e.ColumnIndex = Me.ColExchangeRate.Index Or _
                            e.ColumnIndex = Me.ColAcc_IncrementValue.Index Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                                Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString)
                            End If
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
                        Me.rowIndexItem = e.RowIndex
                        Me.dgvItemHapus = True
                        Me.BackgroundWorker1.RunWorkerAsync()
                        'Me.dgvItem.Rows.RemoveAt(e.RowIndex)
                    End If
                    Exit Sub
                ElseIf e.ColumnIndex = Me.ColJob.Index Then
                    'hapuskan kode job
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Me.dgvItem.Item(Me.ColPrimaryJob.Index, e.RowIndex).Value = -1
                ElseIf e.ColumnIndex = Me.ColDept.Index Then
                    'hapuskan kode dept
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Me.dgvItem.Item(Me.ColPrimaryDept.Index, e.RowIndex).Value = -1
                Else
                    'untuk kolom debet kredit, langsung di balikin ke nilai semula
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                End If
            Else

                ''cek ColRemarks
                If e.ColumnIndex = Me.ColRemarks.Index Then
                    If Not String.IsNullOrEmpty(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) Then
                        If Me.ObjCD.CekDataKosong(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString, False) = False Then
                            If Not String.IsNullOrEmpty(Me.Buffer) Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Else
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                            End If
                        End If
                    End If
                End If


                ''cek debet
                If e.ColumnIndex = Me.ColDebet.Index Then
                    If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
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
                    Me.HitungOriginalValues(e.RowIndex, enumEntryDari.NilaiRupiah)
                End If


                ''Cek kredit 
                If e.ColumnIndex = Me.ColKredit.Index Then
                    If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, Rounding)
                    ''Cek debet
                    If Me.dgvItem.Item(Me.ColDebet.Index, e.RowIndex).Value <> 0 Then
                        Me.dgvItem.Item(Me.ColDebet.Index, e.RowIndex).Value = ObjCD.DigitGrouping(0, Rounding)
                    End If
                    Me.BalanceDebetKredit(e.RowIndex, False, Rounding)
                    Me.HitungOriginalValues(e.RowIndex, enumEntryDari.NilaiRupiah)
                End If


                'Cek Increment 
                If e.ColumnIndex = Me.ColAcc_IncrementStyle.Index Then
                    If String.IsNullOrEmpty(Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value) Then
                        MsgBox("Please enter account no before choosing increment style ", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        If e.RowIndex <> Me.dgvItem.Rows.Count - 1 Then
                            Me.rowIndexItem = e.RowIndex
                            Me.dgvItemHapus = True
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If
                        Exit Sub
                    End If

                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.No_Increment) Then
                        Me.dgvItem.Item(Me.ColAcc_IncrementValue.Index, e.RowIndex).Value = 0
                    ElseIf Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.Fixed_Increment) Then
                        Me.dgvItem.Item(Me.ColAcc_IncrementValue.Index, e.RowIndex).Value = 0
                    ElseIf Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.Percentage_Increment) Then
                        Me.dgvItem.Item(Me.ColAcc_IncrementValue.Index, e.RowIndex).Value = 0
                    End If
                    SamakanValue(Rounding, e.RowIndex)
                End If

                If e.ColumnIndex = Me.ColAcc_IncrementValue.Index Then
                    If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    If Me.dgvItem.Item(Me.ColAcc_IncrementStyle.Index, e.RowIndex).Value = _
                       Me.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.Percentage_Increment) Then
                        If Not IsNumeric(Me.dgvItem.Item(Me.ColAcc_IncrementValue.Index, e.RowIndex).Value) Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If

                        If CDec(Me.dgvItem.Item(Me.ColAcc_IncrementValue.Index, e.RowIndex).Value) > 100 Then
                            MsgBox("Maximum value of percentage number is 100 percent", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    ElseIf Me.dgvItem.Item(Me.ColAcc_IncrementStyle.Index, e.RowIndex).Value = _
                        Me.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.No_Increment) Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = 0
                    ElseIf Me.dgvItem.Item(Me.ColAcc_IncrementStyle.Index, e.RowIndex).Value = _
                       Me.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.Fixed_Increment) Then
                        If Not IsNumeric(Me.dgvItem.Item(Me.ColAcc_IncrementValue.Index, e.RowIndex).Value) Then
                            If IsNumeric(Me.Buffer) Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Rounding)
                            Else
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(0, Rounding)
                            End If
                        Else
                            Me.dgvItem.Item(Me.ColAcc_IncrementValue.Index, e.RowIndex).Value = ObjCD.DigitGrouping(Me.dgvItem.Item(Me.ColAcc_IncrementValue.Index, e.RowIndex).Value, Rounding)
                        End If
                    End If
                    SamakanValue(Rounding, e.RowIndex)
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
                                If e.RowIndex = 0 Then
                                    Me.dgvItem.Item(Me.ColAcc_IncrementStyle.Index, e.RowIndex).Value = Me.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.No_Increment)
                                    Me.dgvItem.Item(Me.ColAcc_IncrementValue.Index, e.RowIndex).Value = ObjCD.DigitGrouping(0, Rounding)
                                Else
                                    SamakanValue(Rounding, 0)
                                End If
                                ExRate = Me.ObjTaxRate.GetTaxRate(Me.dtpDate.Value, Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                                If ExRate = 0 Then
                                    ExRate = 1
                                End If
                                Me.dgvItem.Item(Me.ColExchangeRate.Index, e.RowIndex).Value = ObjCD.DigitGrouping(ExRate, DefPrecExchangeRate)
                                'DefPrecOriginalCurr = Me.ObjCur.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                                Me.HitungOriginalValues(e.RowIndex, enumEntryDari.NilaiOriginal)
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
                                    If e.RowIndex = 0 Then
                                        Me.dgvItem.Item(Me.ColAcc_IncrementStyle.Index, e.RowIndex).Value = Me.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.No_Increment)
                                        Me.dgvItem.Item(Me.ColAcc_IncrementValue.Index, e.RowIndex).Value = 0
                                    Else
                                        SamakanValue(Rounding, 0)
                                    End If
                                    ExRate = Me.ObjTaxRate.GetTaxRate(Me.dtpDate.Value, Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                                    If ExRate = 0 Then
                                        ExRate = 1
                                    End If
                                    Me.dgvItem.Item(Me.ColExchangeRate.Index, e.RowIndex).Value = ObjCD.DigitGrouping(ExRate, DefPrecExchangeRate)
                                    'DefPrecOriginalCurr = Me.ObjCur.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                                    Me.HitungOriginalValues(e.RowIndex, enumEntryDari.NilaiOriginal)
                                ElseIf ij.mPrimaryAkun = -1 Then
                                    '2.b. kode akun yang dicari tidak ada --> kembalikan ke kondisi buffer
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                        If e.RowIndex <> Me.dgvItem.Rows.Count - 1 Then
                                            Me.rowIndexItem = e.RowIndex
                                            Me.dgvItemHapus = True
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        End If
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
                        If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Please enter positive number", MsgBoxStyle.Information, _
                            "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        'digit groping
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                        ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, DefPrecExchangeRate)
                        'hitung original value
                        Me.HitungOriginalValues(e.RowIndex, enumEntryDari.Kurs)
                    End If
                End If


                '3. debet original currency
                If e.ColumnIndex = Me.ColDebet_OriginalCurr.Index Then
                    If Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value <> Me.PrimaryBaseCurrency Then
                        If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
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
                        If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
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

                '5. Dept
                If e.ColumnIndex = Me.ColDept.Index Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    If Me.ObjDept.ApakahSudahAda(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvItem.Item(Me.ColPrimaryDept.Index, e.RowIndex).Value = Me.ObjDept.GetPrimary(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                    Else
                        Dim ij As New Core_Module.FrmMasterDepartementCari
                        ij.txtCari.Text = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.PrimaryDepartement <> -1 Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjDept.GetDeptCode(ij.PrimaryDepartement)
                            Me.dgvItem.Item(Me.ColPrimaryDept.Index, e.RowIndex).Value = ij.PrimaryDepartement
                        ElseIf ij.PrimaryDepartement = -1 Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                Me.dgvItem.Item(Me.ColPrimaryDept.Index, e.RowIndex).Value = -1
                                Exit Sub
                            End If
                        End If
                        ij = Nothing
                    End If
                End If

                '6. Job
                If e.ColumnIndex = Me.ColJob.Index Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    If Me.ObjJob.ApakahSudahAda(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvItem.Item(Me.ColPrimaryJob.Index, e.RowIndex).Value = Me.ObjJob.GetPrimary(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                    Else
                        Dim ij As New Core_Module.FrmMasterJobCari
                        ij.txtCari.Text = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.PrimaryJob <> -1 Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjJob.GetJobCode(ij.PrimaryJob)
                            Me.dgvItem.Item(Me.ColPrimaryJob.Index, e.RowIndex).Value = ij.PrimaryJob
                        ElseIf ij.PrimaryJob = -1 Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                Me.dgvItem.Item(Me.ColPrimaryJob.Index, e.RowIndex).Value = -1
                                Exit Sub
                            End If
                        End If
                        ij = Nothing
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
            Me.HitungUnbalanceIncrement()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCur = Nothing
        End Try
    End Sub

    Private Sub SamakanValue(ByVal rounding As Integer, ByVal Index As Integer)
        Dim rowDGV As DataGridViewRow
        Try
            Dim Value As Decimal = 0
            Dim style As Integer = 0
            Value = Me.dgvItem.Rows(Index).Cells(Me.ColAcc_IncrementValue.Index).Value
            If Me.dgvItem.Rows(Index).Cells(Me.ColAcc_IncrementStyle.Index).Value = Me.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.No_Increment) Then
                style = ClsTransaksiRecurring.EnumIncrementStyle.No_Increment
            ElseIf Me.dgvItem.Rows(Index).Cells(Me.ColAcc_IncrementStyle.Index).Value = Me.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.Fixed_Increment) Then
                style = ClsTransaksiRecurring.EnumIncrementStyle.Fixed_Increment
            ElseIf Me.dgvItem.Rows(Index).Cells(Me.ColAcc_IncrementStyle.Index).Value = Me.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.Percentage_Increment) Then
                style = ClsTransaksiRecurring.EnumIncrementStyle.Percentage_Increment
            End If

            Dim CountDebet As Integer = 0
            Dim CountCredit As Integer = 0
            For Each rowDGV In Me.dgvItem.Rows
                With rowDGV
                    If IsNumeric(.Cells(Me.ColDebet.Index).Value) AndAlso CDec(.Cells(Me.ColDebet.Index).Value) > 0 Then
                        CountDebet += 1
                    End If

                    If IsNumeric(.Cells(Me.ColKredit.Index).Value) AndAlso CDec(.Cells(Me.ColKredit.Index).Value) > 0 Then
                        CountCredit += 1
                    End If
                End With
            Next

            For Each rowDGV In Me.dgvItem.Rows
                With rowDGV
                    If IsNumeric(.Cells(Me.ColPrimaryAkun.Index).Value) AndAlso CDec(.Cells(Me.ColPrimaryAkun.Index).Value) > -1 Then
                        .Cells(Me.ColAcc_IncrementStyle.Index).Value = Me.cboDGV.Items.Item(style)
                        If style = ClsTransaksiRecurring.EnumIncrementStyle.Fixed_Increment Then
                            If CountDebet = CountCredit Then
                                .Cells(Me.ColAcc_IncrementValue.Index).Value = Me.ObjCD.DigitGrouping(Value, rounding)
                            End If
                        Else
                            .Cells(Me.ColAcc_IncrementValue.Index).Value = Me.ObjCD.DigitGrouping(Value, rounding)
                        End If
                    End If
                End With
            Next
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
                    Me.HitungOriginalValues(rowTarget.Index, enumEntryDari.NilaiRupiah)
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
                            Me.HitungOriginalValues(rowTarget.Index, enumEntryDari.NilaiRupiah)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BalanceIncrement(ByVal vIndexAcuan As Integer, ByVal vAcuanDebet As Boolean, ByVal rounding As Integer, ByVal Increment As Decimal, _
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
                    Me.HitungOriginalValues(rowTarget.Index, enumEntryDari.NilaiOriginal)
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
                            Me.HitungOriginalValues(rowTarget.Index, enumEntryDari.NilaiRupiah)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
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

    Private Function HitungUnbalance(ByVal vdebet As Decimal, ByVal vkredit As Decimal) As Decimal
        Dim Unbalance As Decimal = 0
        Try
            Unbalance = Math.Round(vdebet - vkredit, Me.DefaultPrecision, MidpointRounding.AwayFromZero)
            Return Unbalance
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function


    Private Function HitungUnbalanceIncrement() As Decimal
        Dim row As DataGridViewRow
        Try
            Dim Unbalance As Decimal = 0
            ''hitung debet dan kredit
            Dim TotalDebet As Decimal = 0
            Dim TotalKredit As Decimal = 0
            For Each row In dgvItem.Rows
                If Me.dgvItem.Item(Me.ColAcc_IncrementStyle.Index, row.Index).Value = Me.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.Fixed_Increment) Then
                    If CDec(Me.dgvItem.Item(Me.ColDebet.Index, row.Index).Value) > 0 Then
                        TotalDebet += CDec(Me.dgvItem.Item(Me.ColAcc_IncrementValue.Index, row.Index).Value)
                    End If
                    If CDec(Me.dgvItem.Item(Me.ColKredit.Index, row.Index).Value) > 0 Then
                        TotalKredit += CDec(Me.dgvItem.Item(Me.ColAcc_IncrementValue.Index, row.Index).Value)
                    End If
                End If
            Next
            Me.txtUnbalanceIncrement.Text = CStr(ObjCD.DigitGrouping(Me.HitungUnbalance(TotalDebet, TotalKredit), Me.DefaultPrecision))
            Unbalance = CDec(ObjCD.DigitGrouping(Me.HitungUnbalance(TotalDebet, TotalKredit), Me.DefaultPrecision))
            Return Unbalance

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub HitungOriginalValues(ByVal vRowIndex As Integer, ByVal venumEntryDari As enumEntryDari)
        Dim vRowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0
        Dim DefPrecOriginal As Integer = 0

        Try
            vRowDGV = Me.dgvItem.Rows(vRowIndex)
            DefPrecBase = Me.ObjCur.GetDefaultPrecission(Me.PrimaryBaseCurrency)
            DefPrecOriginal = Me.ObjCur.GetDefaultPrecission(vRowDGV.Cells(Me.ColPrimaryCurrency.Index).Value)
            If venumEntryDari = enumEntryDari.Kurs Then
                If CDec(vRowDGV.Cells(Me.ColDebet.Index).Value) > 0 Then
                    vRowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColDebet.Index).Value) / _
                                                                          CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value), DefPrecOriginal)
                    vRowDGV.Cells(Me.ColCredit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(0, DefPrecOriginal)
                ElseIf CDec(vRowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value) > 0 Then
                    vRowDGV.Cells(Me.ColDebet.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value) * _
                                                                          CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value), DefPrecBase)
                    vRowDGV.Cells(Me.ColKredit.Index).Value = ObjCD.DigitGrouping(0, DefPrecBase)
                ElseIf CDec(vRowDGV.Cells(Me.ColKredit.Index).Value) > 0 Then
                    vRowDGV.Cells(Me.ColCredit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColKredit.Index).Value) / _
                                                                         CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value), DefPrecOriginal)
                    vRowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(0, DefPrecOriginal)
                ElseIf CDec(vRowDGV.Cells(Me.ColCredit_OriginalCurrency.Index).Value) > 0 Then
                    vRowDGV.Cells(Me.ColKredit.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColCredit_OriginalCurrency.Index).Value) * _
                                                                          CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value), DefPrecBase)
                    vRowDGV.Cells(Me.ColDebet.Index).Value = ObjCD.DigitGrouping(0, DefPrecBase)
                End If

            ElseIf venumEntryDari = enumEntryDari.NilaiOriginal Then
                If CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value) > 0 Then
                    vRowDGV.Cells(Me.ColDebet.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value) * _
                                                                          CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value), DefPrecBase)
                    vRowDGV.Cells(Me.ColKredit.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColCredit_OriginalCurrency.Index).Value) * _
                                                                          CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value), DefPrecBase)
                ElseIf CDec(vRowDGV.Cells(Me.ColDebet.Index).Value) > 0 Then
                    vRowDGV.Cells(Me.ColExchangeRate.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value) / _
                                                                    CDec(vRowDGV.Cells(Me.ColDebet.Index).Value), DefPrecBase)

                ElseIf CDec(vRowDGV.Cells(Me.ColKredit.Index).Value) > 0 Then
                    vRowDGV.Cells(Me.ColExchangeRate.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColCredit_OriginalCurrency.Index).Value) / _
                                                                    CDec(vRowDGV.Cells(Me.ColKredit.Index).Value), DefPrecBase)
                End If
            ElseIf venumEntryDari = enumEntryDari.NilaiRupiah Then

                If CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value) > 0 Then
                    vRowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColDebet.Index).Value) / _
                                                                          CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value), DefPrecOriginal)
                    vRowDGV.Cells(Me.ColCredit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColKredit.Index).Value) / _
                                                                          CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value), DefPrecOriginal)
                ElseIf CDec(vRowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value) > 0 Then
                    vRowDGV.Cells(Me.ColExchangeRate.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColDebet.Index).Value) / _
                                                                    CDec(vRowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value), DefPrecBase)

                ElseIf CDec(vRowDGV.Cells(Me.ColCredit_OriginalCurrency.Index).Value) > 0 Then
                    vRowDGV.Cells(Me.ColExchangeRate.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColKredit.Index).Value) / _
                                                                    CDec(vRowDGV.Cells(Me.ColCredit_OriginalCurrency.Index).Value), DefPrecBase)
                End If
            End If

            'If CDec(vRowDGV.Cells(Me.ColDebet.Index).Value) > 0 Then
            '    vRowDGV.Cells(Me.ColCredit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(0, DefPrecOriginal)
            '    vRowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColDebet.Index).Value) / _
            '    CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value), DefPrecOriginal)
            'Else
            '    vRowDGV.Cells(Me.ColCredit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColKredit.Index).Value) / _
            '    CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value), DefPrecOriginal)
            '    vRowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(0, DefPrecOriginal)
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Dim PrimaryDummy As Double = -1
        Dim PrimaryAkunRE As String = ""

        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.ClosingJournal Then
                'jika akun retained earning tidak ada maka tidak bisa lanjut
                PrimaryAkunRE = Me.ObjDefaultValue.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, ClsTransaksiSetting.EnumNamaParameter.Akun_Retained_Earning)
                If PrimaryAkunRE = "" Or PrimaryAkunRE = "-1" Then
                    MsgBox("Please specify retained earning account in global setting data.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            Me.Baru = True
            Me.KosongkanField()
            Me.TxtRefNo.Text = ObjRCR.GetNoBaru(Me.dtpDate.Value)
            Me.AktifkanDetail()
            Me.dgvItem.AllowUserToAddRows = True
            Me.dgvReminder.AllowUserToAddRows = True
            Me.dtpDate.Focus()
            'Me.ObjJurnal.ResetObject(Me.dgvItem)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCOde
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCOde
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult

        Try
            'If Me.TxtPrimaryJournal.Text = -1 AndAlso Me.dgvEntries.Rows.Count > 0 Then
            '    Dim strNoJournal As String = ""
            '    Dim rowDgv As New DataGridViewRow
            '    For Each rowDgv In Me.dgvEntries.Rows
            '        With rowDgv
            '            If .Cells(Me.ColTrans_Status.Index).Value = _
            '                Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Created) OrElse _
            '                .Cells(Me.ColTrans_Status.Index).Value = _
            '                Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Locked) Then
            '                If strNoJournal = "" Then
            '                    strNoJournal &= vbNewLine & .Cells(Me.ColTrans_JournalNumber.Index).Value
            '                Else
            '                    strNoJournal &= vbNewLine & ", " & .Cells(Me.ColTrans_JournalNumber.Index).Value
            '                End If
            '            End If
            '        End With
            '    Next
            '    If strNoJournal <> "" Then
            '        MsgBox("Can not cancel transaction because " & strNoJournal & " already exist in journal transaction ", MsgBoxStyle.Information, "Information")
            '        Exit Sub
            '    End If
            'End If

            'If Me.AdaJournalBaru = True Then
            '    MsgBox("Can not cancel transaction because the data already exist in journal transaction ", MsgBoxStyle.Information, "Information")
            '    Exit Sub
            'End If

            Me.dgvItem.AllowUserToAddRows = False
            Me.dgvReminder.AllowUserToAddRows = False
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
                        Me.ObjRecurring.CariDetail(Me.TxtPrimaryJournal.Text, Me.dgvItem, Me)
                        Me.SalinkeField()
                    End If
                Else
                    Me.dgvItem.AllowUserToAddRows = True
                    Me.dgvReminder.AllowUserToAddRows = True
                    Exit Sub
                End If
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
            End If

            If Me.FromReminder = True Then
                Me.Close()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub cboFrequency_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboFrequency.SelectedIndexChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.cboFrequency.SelectedIndex = ClsTransaksiRecurring.EnumFrequency.Daily OrElse _
                              Me.cboFrequency.SelectedIndex = ClsTransaksiRecurring.EnumFrequency.Yearly Then
                    Me.cboEveryMonth.Enabled = False
                    Me.cboEveryWeek.Enabled = False
                ElseIf Me.cboFrequency.SelectedIndex = ClsTransaksiRecurring.EnumFrequency.Bi_Weekly OrElse _
                       Me.cboFrequency.SelectedIndex = ClsTransaksiRecurring.EnumFrequency.Weekly Then
                    Me.cboEveryWeek.Enabled = True
                    Me.cboEveryMonth.Enabled = False
                    Me.cboEveryMonth.Visible = False
                    Me.cboEveryWeek.Visible = True
                ElseIf Me.cboFrequency.SelectedIndex = ClsTransaksiRecurring.EnumFrequency.Monthly OrElse _
                        Me.cboFrequency.SelectedIndex = ClsTransaksiRecurring.EnumFrequency.Quarter_Yearly OrElse _
                        Me.cboFrequency.SelectedIndex = ClsTransaksiRecurring.EnumFrequency.Semester Then
                    Me.cboEveryWeek.Enabled = False
                    Me.cboEveryMonth.Enabled = True
                    Me.cboEveryMonth.Visible = True
                    Me.cboEveryWeek.Visible = False
                End If
                Me.dtpStartDate.Value = Me.SettingTanggalAwal(Me.cboFrequency.SelectedIndex, Me.cboEveryWeek.SelectedIndex, Me.cboEveryMonth.SelectedIndex, Me.dtpStartDate.Value)
                If Me.chkLimit.Checked = True Then
                    If IsNumeric(Me.txtLimit.Text) Then
                        Me.Limit = CLng(Me.txtLimit.Text)
                    Else
                        Me.Limit = 0
                    End If
                    Me.dtpEndDate.Value = Me.SettingTanggalAkhir(Me.cboFrequency.SelectedIndex, Me.dtpStartDate.Value, Me.Limit)
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkLimit_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkLimit.CheckedChanged
        Dim dtAwal As Date = Date.Now

        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.chkLimit.Checked = True Then
                    'If Me.Baru = True Then
                    'End If
                    Me.txtLimit.Text = Me.GetLimit(Me.dtpStartDate.Value, Me.dtpEndDate.Value)
                    Me.txtLimit.ReadOnly = False
                    Me.txtLimit.TabStop = True
                    Me.txtLimit.BackColor = Me.TxtRemarks.BackColor
                    Me.dtpEndDate.Enabled = False
                    Me.Limit = CDec(Me.txtLimit.Text)
                    If Me.Baru = True Then
                        Me.dtpEndDate.Value = Me.SettingTanggalAkhir(Me.cboFrequency.SelectedIndex, Me.dtpStartDate.Value, Me.Limit - 1)
                    End If
                    Me.TambahkanBarisEntries(Me.Limit)
                Else
                    Me.txtLimit.Text = 0
                    Me.txtLimit.ReadOnly = True
                    Me.txtLimit.TabStop = False
                    Me.txtLimit.BackColor = Me.txtReadOnly.BackColor
                    Me.dtpEndDate.Enabled = True
                    dtAwal = Me.SettingTanggalAwal(Me.cboFrequency.SelectedIndex, Me.cboEveryWeek.SelectedIndex, Me.cboEveryMonth.SelectedIndex, Me.dtpStartDate.Value)
                    Me.dtpStartDate.Value = Format(dtAwal, "dd MMM yyyy")
                    If Me.Baru = True Then
                        Me.dtpEndDate.Value = Me.SettingTanggalAkhir(Me.cboFrequency.SelectedIndex, Me.dtpStartDate.Value, 9)
                        Me.Limit = Me.GetLimit(Me.dtpStartDate.Value, Me.dtpEndDate.Value)
                    End If
                    Me.TambahkanBarisEntries(Me.Limit)
                End If
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function SettingTanggalAwal(ByVal vIndexFrequency As Integer, ByVal vIndexWeek As Integer, ByVal vIndexMonth As Integer, ByVal dtAwal As Date) As Date
        Dim objcd As New Core_Module.ClsCekData
        Try
            dtAwal = Me.setTanggal(dtAwal, vIndexFrequency, vIndexWeek, vIndexMonth, False)
            Return dtAwal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function setTanggal(ByVal vdtTanggal As Date, ByVal vEnumFrekuensi As ClsTransaksiRecurring.EnumFrequency, _
    ByVal vEnumWeek As ClsTransaksiRecurring.EnumEvery_Week, ByVal vEnumMonth As ClsTransaksiRecurring.EnumEvery_Month, _
    ByVal ShowMsg As Boolean) As Date
        Dim dtDate As Date = vdtTanggal
        Try
            If vEnumFrekuensi = ClsTransaksiRecurring.EnumFrequency.Weekly OrElse _
               vEnumFrekuensi = ClsTransaksiRecurring.EnumFrequency.Bi_Weekly Then
                '=========================================================================================================================
                'khusus mingguan & duamingguan
                '=========================================================================================================================
                If vEnumWeek = ClsTransaksiRecurring.EnumEvery_Week.Sunday Then
                    'Sunday
                    dtDate = dtDate.AddDays(DayOfWeek.Sunday - dtDate.DayOfWeek)
                ElseIf vEnumWeek = ClsTransaksiRecurring.EnumEvery_Week.Monday Then
                    'Monday 
                    dtDate = dtDate.AddDays(DayOfWeek.Monday - dtDate.DayOfWeek)
                ElseIf vEnumWeek = ClsTransaksiRecurring.EnumEvery_Week.Tuesday Then
                    'Tuesday 
                    dtDate = dtDate.AddDays(DayOfWeek.Tuesday - dtDate.DayOfWeek)
                ElseIf vEnumWeek = ClsTransaksiRecurring.EnumEvery_Week.Thursday Then
                    'Thursday 
                    dtDate = dtDate.AddDays(DayOfWeek.Thursday - dtDate.DayOfWeek)
                ElseIf vEnumWeek = ClsTransaksiRecurring.EnumEvery_Week.Wednesday Then
                    'Wednesday 
                    dtDate = dtDate.AddDays(DayOfWeek.Wednesday - dtDate.DayOfWeek)
                ElseIf vEnumWeek = ClsTransaksiRecurring.EnumEvery_Week.Friday Then
                    'Friday 
                    dtDate = dtDate.AddDays(DayOfWeek.Friday - dtDate.DayOfWeek)
                ElseIf vEnumWeek = ClsTransaksiRecurring.EnumEvery_Week.Saturday Then
                    'Saturday 
                    dtDate = dtDate.AddDays(DayOfWeek.Saturday - dtDate.DayOfWeek)
                End If
            ElseIf vEnumFrekuensi = ClsTransaksiRecurring.EnumFrequency.Monthly OrElse _
               vEnumFrekuensi = ClsTransaksiRecurring.EnumFrequency.Quarter_Yearly OrElse _
               vEnumFrekuensi = ClsTransaksiRecurring.EnumFrequency.Semester Then
                '=========================================================================================================================
                'khusus Bulanan
                '=========================================================================================================================
                If vEnumMonth = ClsTransaksiRecurring.EnumEvery_Month.First_Day_Of_The_Month Then
                    'First_Day_Of_The_Month
                    dtDate = ObjCD.GetFirstDateofTheMonth(dtDate)
                ElseIf vEnumMonth = ClsTransaksiRecurring.EnumEvery_Month.Last_Day_Of_The_Month Then
                    'Last_Day_Of_The_Month
                    dtDate = ObjCD.GetLasDateOfTheMonth(dtDate)
                ElseIf vEnumMonth = ClsTransaksiRecurring.EnumEvery_Month.First_Friday Then
                    'First_Friday
                    dtDate = ObjCD.GetFirstDateofTheMonth(dtDate)
                    If dtDate.AddDays(DayOfWeek.Friday - dtDate.DayOfWeek) >= dtDate Then
                        dtDate = dtDate.AddDays(DayOfWeek.Friday - dtDate.DayOfWeek)
                    Else
                        dtDate = dtDate.AddDays((DayOfWeek.Friday - dtDate.DayOfWeek) + 7)
                    End If
                ElseIf vEnumMonth = ClsTransaksiRecurring.EnumEvery_Month.First_Monday Then
                    'First_Monday
                    dtDate = ObjCD.GetFirstDateofTheMonth(dtDate)
                    If dtDate.AddDays(DayOfWeek.Monday - dtDate.DayOfWeek) >= dtDate Then
                        dtDate = dtDate.AddDays(DayOfWeek.Monday - dtDate.DayOfWeek)
                    Else
                        dtDate = dtDate.AddDays((DayOfWeek.Monday - dtDate.DayOfWeek) + 7)
                    End If
                ElseIf vEnumMonth = ClsTransaksiRecurring.EnumEvery_Month.First_Saturday Then
                    'First_Saturday
                    dtDate = ObjCD.GetFirstDateofTheMonth(dtDate)
                    If dtDate.AddDays(DayOfWeek.Saturday - dtDate.DayOfWeek) >= dtDate Then
                        dtDate = dtDate.AddDays(DayOfWeek.Saturday - dtDate.DayOfWeek)
                    Else
                        dtDate = dtDate.AddDays((DayOfWeek.Saturday - dtDate.DayOfWeek) + 7)
                    End If
                ElseIf vEnumMonth = ClsTransaksiRecurring.EnumEvery_Month.First_Sunday Then
                    'First_Sunday
                    dtDate = ObjCD.GetFirstDateofTheMonth(dtDate)
                    If dtDate.AddDays(DayOfWeek.Sunday - dtDate.DayOfWeek) >= dtDate Then
                        dtDate = dtDate.AddDays(DayOfWeek.Sunday - dtDate.DayOfWeek)
                    Else
                        dtDate = dtDate.AddDays((DayOfWeek.Sunday - dtDate.DayOfWeek) + 7)
                    End If
                ElseIf vEnumMonth = ClsTransaksiRecurring.EnumEvery_Month.First_Thursday Then
                    'First_Thursday
                    dtDate = ObjCD.GetFirstDateofTheMonth(dtDate)
                    If dtDate.AddDays(DayOfWeek.Thursday - dtDate.DayOfWeek) >= dtDate Then
                        dtDate = dtDate.AddDays(DayOfWeek.Thursday - dtDate.DayOfWeek)
                    Else
                        dtDate = dtDate.AddDays((DayOfWeek.Thursday - dtDate.DayOfWeek) + 7)
                    End If
                ElseIf vEnumMonth = ClsTransaksiRecurring.EnumEvery_Month.First_Tuesday Then
                    'First_Tuesday
                    dtDate = ObjCD.GetFirstDateofTheMonth(dtDate)
                    If dtDate.AddDays(DayOfWeek.Tuesday - dtDate.DayOfWeek) >= dtDate Then
                        dtDate = dtDate.AddDays(DayOfWeek.Tuesday - dtDate.DayOfWeek)
                    Else
                        dtDate = dtDate.AddDays((DayOfWeek.Tuesday - dtDate.DayOfWeek) + 7)
                    End If
                ElseIf vEnumMonth = ClsTransaksiRecurring.EnumEvery_Month.First_Wednesday Then
                    'First_Wednesday
                    dtDate = ObjCD.GetFirstDateofTheMonth(dtDate)
                    If dtDate.AddDays(DayOfWeek.Wednesday - dtDate.DayOfWeek) >= dtDate Then
                        dtDate = dtDate.AddDays(DayOfWeek.Wednesday - dtDate.DayOfWeek)
                    Else
                        dtDate = dtDate.AddDays((DayOfWeek.Wednesday - dtDate.DayOfWeek) + 7)
                    End If
                ElseIf vEnumMonth = ClsTransaksiRecurring.EnumEvery_Month.Last_Friday Then
                    'Last_Friday
                    dtDate = ObjCD.GetLasDateOfTheMonth(dtDate)
                    If dtDate.AddDays(DayOfWeek.Friday - dtDate.DayOfWeek) <= dtDate Then
                        dtDate = dtDate.AddDays(DayOfWeek.Friday - dtDate.DayOfWeek)
                    Else
                        dtDate = dtDate.AddDays((DayOfWeek.Friday - dtDate.DayOfWeek) - 7)
                    End If
                ElseIf vEnumMonth = ClsTransaksiRecurring.EnumEvery_Month.Last_Monday Then
                    'Last_Monday
                    dtDate = ObjCD.GetLasDateOfTheMonth(dtDate)
                    If dtDate.AddDays(DayOfWeek.Monday - dtDate.DayOfWeek) <= dtDate Then
                        dtDate = dtDate.AddDays(DayOfWeek.Monday - dtDate.DayOfWeek)
                    Else
                        dtDate = dtDate.AddDays((DayOfWeek.Monday - dtDate.DayOfWeek) - 7)
                    End If
                 ElseIf vEnumMonth = ClsTransaksiRecurring.EnumEvery_Month.Last_Saturday Then
                    'Last_Saturday
                    dtDate = ObjCD.GetLasDateOfTheMonth(dtDate)
                    If dtDate.AddDays(DayOfWeek.Saturday - dtDate.DayOfWeek) <= dtDate Then
                        dtDate = dtDate.AddDays(DayOfWeek.Saturday - dtDate.DayOfWeek)
                    Else
                        dtDate = dtDate.AddDays((DayOfWeek.Saturday - dtDate.DayOfWeek) - 7)
                    End If
                ElseIf vEnumMonth = ClsTransaksiRecurring.EnumEvery_Month.Last_Sunday Then
                    'Last_Sunday
                    dtDate = ObjCD.GetLasDateOfTheMonth(dtDate)
                    If dtDate.AddDays(DayOfWeek.Sunday - dtDate.DayOfWeek) <= dtDate Then
                        dtDate = dtDate.AddDays(DayOfWeek.Sunday - dtDate.DayOfWeek)
                    Else
                        dtDate = dtDate.AddDays((DayOfWeek.Sunday - dtDate.DayOfWeek) - 7)
                    End If
                ElseIf vEnumMonth = ClsTransaksiRecurring.EnumEvery_Month.Last_Thusday Then
                    'Last_Thusday
                    dtDate = ObjCD.GetLasDateOfTheMonth(dtDate)
                    If dtDate.AddDays(DayOfWeek.Thursday - dtDate.DayOfWeek) <= dtDate Then
                        dtDate = dtDate.AddDays(DayOfWeek.Thursday - dtDate.DayOfWeek)
                    Else
                        dtDate = dtDate.AddDays((DayOfWeek.Thursday - dtDate.DayOfWeek) - 7)
                    End If
                ElseIf vEnumMonth = ClsTransaksiRecurring.EnumEvery_Month.Last_Tuesday Then
                    'Last_Tuesday
                    dtDate = ObjCD.GetLasDateOfTheMonth(dtDate)
                    If dtDate.AddDays(DayOfWeek.Tuesday - dtDate.DayOfWeek) <= dtDate Then
                        dtDate = dtDate.AddDays(DayOfWeek.Tuesday - dtDate.DayOfWeek)
                    Else
                        dtDate = dtDate.AddDays((DayOfWeek.Tuesday - dtDate.DayOfWeek) - 7)
                    End If
                ElseIf vEnumMonth = ClsTransaksiRecurring.EnumEvery_Month.Last_Wednesday Then
                    'Last_Wednesday
                    dtDate = ObjCD.GetLasDateOfTheMonth(dtDate)
                    If dtDate.AddDays(DayOfWeek.Wednesday - dtDate.DayOfWeek) <= dtDate Then
                        dtDate = dtDate.AddDays(DayOfWeek.Wednesday - dtDate.DayOfWeek)
                    Else
                        dtDate = dtDate.AddDays((DayOfWeek.Wednesday - dtDate.DayOfWeek) - 7)
                    End If
                Else
                    If IsNumeric(vEnumMonth) Then
                        Dim intD As Integer = 0
                        intD = Date.DaysInMonth(dtDate.Year, dtDate.Month)
                        If vEnumMonth >= ClsTransaksiRecurring.EnumEvery_Month._29 Then
                            If intD < 29 Then
                                If ShowMsg = True Then
                                    MsgBox("Maximum date of " & Format(dtDate, "MMMM") & _
                                           " is " & Date.DaysInMonth(dtDate.Year, dtDate.Month) & _
                                           "th .", MsgBoxStyle.Information, "Information")
                                    'Jika MaxTanggal < Tanggal pilihan, maka ambil MaxTanggal
                                    dtDate = ObjCD.GetLasDateOfTheMonth(dtDate)
                                End If
                            Else
                                Dim hari As Integer = 0
                                hari = CInt(Me.cboEveryMonth.Text) - 1
                                dtDate = ObjCD.GetFirstDateofTheMonth(dtDate).AddDays(hari)
                            End If
                        Else
                            Dim hari As Integer = 0
                            hari = CInt(Me.cboEveryMonth.Text) - 1
                            dtDate = ObjCD.GetFirstDateofTheMonth(dtDate).AddDays(hari)
                        End If
                    End If
                End If
            End If

            Return dtDate
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function


    Public Function GetLimit(ByVal vdtAwal As Date, ByVal vdtAkhir As Date) As Integer
        Dim intLimitX As Integer = 0
        'Dim DummyDate As Date = vdtAwal

        Try
            If Me.cboFrequency.SelectedIndex = ClsTransaksiRecurring.EnumFrequency.Daily Then
                intLimitX = DateDiff(DateInterval.Day, vdtAwal.Date, vdtAkhir.Date) + 1
            ElseIf Me.cboFrequency.SelectedIndex = ClsTransaksiRecurring.EnumFrequency.Yearly Then
                intLimitX = DateDiff(DateInterval.Year, vdtAwal.Date, vdtAkhir.Date) + 1
            ElseIf Me.cboFrequency.SelectedIndex = ClsTransaksiRecurring.EnumFrequency.Monthly Then
                intLimitX = DateDiff(DateInterval.Month, vdtAwal.Date, vdtAkhir.Date) + 1
            ElseIf Me.cboFrequency.SelectedIndex = ClsTransaksiRecurring.EnumFrequency.Quarter_Yearly Then
                intLimitX = DateDiff(DateInterval.Quarter, vdtAwal.Date, vdtAkhir.Date) + 1
            ElseIf Me.cboFrequency.SelectedIndex = ClsTransaksiRecurring.EnumFrequency.Semester Then
                Dim day As Integer = 0
                day = DateDiff(DateInterval.Month, vdtAwal.Date, vdtAkhir.Date) + 1
                Dim intA As Integer = 0
                While intA < day
                    intA += 6
                    intLimitX += 1
                End While
            ElseIf Me.cboFrequency.SelectedIndex = ClsTransaksiRecurring.EnumFrequency.Weekly Then
                intLimitX = DateDiff(DateInterval.Weekday, vdtAwal.Date, vdtAkhir.Date) + 1
            ElseIf Me.cboFrequency.SelectedIndex = ClsTransaksiRecurring.EnumFrequency.Bi_Weekly Then
                Dim day As Integer = 0
                day = DateDiff(DateInterval.Day, vdtAwal.Date, vdtAkhir.Date) + 1
                Dim intA As Integer = 0
                While intA < day
                    intA += 14
                    intLimitX += 1
                End While

            End If

            Return intLimitX
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub GetTanggalAkhir()
        Dim rowDGV As DataGridViewRow

        Try
            If Me.dgvEntries.RowCount = 0 Then
                Me.dtpEndDate.Value = Me.dtpStartDate.Value
            Else
                rowDGV = Me.dgvEntries.Rows(Me.dgvEntries.RowCount - 1)
                Me.dtpEndDate.Value = rowDGV.Cells(Me.ColTrans_JournalDate.Index).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function SettingTanggalAkhir(ByVal vIndexFrequency As Integer, ByVal dtAwal As Date, ByVal Kali As Integer) As Date
        Dim dtAkhir As Date = Date.Now
        Dim objcd As New Core_Module.ClsCekData
        ' Dim Kali As Integer = 0

        Try
            'If Kali > 0 Then
            '    Kali = Kali - 1
            'End If
            If Me.cboFrequency.SelectedIndex = ClsTransaksiRecurring.EnumFrequency.Daily Then
                'ditambah 9 hari
                Kali = Kali
                dtAkhir = dtAwal.AddDays(Kali)
             ElseIf Me.cboFrequency.SelectedIndex = ClsTransaksiRecurring.EnumFrequency.Weekly Then
                'ditambah 9 minggu
                Kali = Kali * 7
                dtAkhir = dtAwal.AddDays(Kali)
            ElseIf Me.cboFrequency.SelectedIndex = ClsTransaksiRecurring.EnumFrequency.Bi_Weekly Then
                'ditambah 9 (2mingguan)
                Kali = Kali * 14
                dtAkhir = dtAwal.AddDays(Kali)
            ElseIf Me.cboFrequency.SelectedIndex = ClsTransaksiRecurring.EnumFrequency.Monthly Then
                'ditambah 9 bulan
                Kali = Kali
                dtAkhir = dtAwal.AddMonths(Kali)
                If Me.cboEveryMonth.SelectedIndex = ClsTransaksiRecurring.EnumEvery_Month.First_Day_Of_The_Month Then
                    dtAkhir = Me.ObjCD.GetFirstDateofTheMonth(dtAkhir)
                ElseIf Me.cboEveryMonth.SelectedIndex = ClsTransaksiRecurring.EnumEvery_Month.Last_Day_Of_The_Month Then
                    dtAkhir = Me.ObjCD.GetLasDateOfTheMonth(dtAkhir)
                End If
            ElseIf Me.cboFrequency.SelectedIndex = ClsTransaksiRecurring.EnumFrequency.Quarter_Yearly Then
                'ditambah 9 bulan
                Kali = Kali * 3
                dtAkhir = dtAwal.AddMonths(Kali)
            ElseIf Me.cboFrequency.SelectedIndex = ClsTransaksiRecurring.EnumFrequency.Semester Then
                Kali = Kali * 6
                dtAkhir = dtAwal.AddMonths(Kali)
            ElseIf Me.cboFrequency.SelectedIndex = ClsTransaksiRecurring.EnumFrequency.Yearly Then
                Kali = Kali
                dtAkhir = dtAwal.AddYears(Kali)
            End If

            Return dtAkhir
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub txtLimit_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLimit.GotFocus
        Try
            If Me.txtLimit.Enabled = True Then
                If Me.txtLimit.ReadOnly = False Then
                    Me.txtLimit.BackColor = Me.BackcolorAll
                Else
                    txtLimit.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtLimit_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLimit.Leave
        Try
            If Me.TxtRemarks.Enabled = True Then

                Dim intLimit As Long = 0
                If Me.chkLimit.Checked = True Then
                    If IsNumeric(Me.txtLimit.Text) Then
                        intLimit = CLng(Me.txtLimit.Text)
                    End If
                    If Me.Baru = True Then
                        Me.dtpEndDate.Value = Me.SettingTanggalAkhir(Me.cboFrequency.SelectedIndex, Me.dtpStartDate.Value, Me.Limit - 1)
                    End If
                Else
                    intLimit = Me.GetLimit(Me.dtpStartDate.Value, Me.dtpEndDate.Value)
                End If

                Me.Limit = intLimit

                If Me.chkLimit.Checked = True Then
                    Me.txtLimit.Text = Me.Limit
                End If


                If Me.dgvEntries.Rows.Count > 0 Then
                    If Me.Limit > Me.intBatasBaris AndAlso _
                    Me.dgvEntries.Rows(Me.dgvEntries.Rows.Count - 1).Cells(Me.ColTrans_Status.Index).Value <> _
                    Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed) Then
                        Me.intBatasBaris = Me.dgvEntries.RowCount + 1
                    Else
                        Me.intBatasBaris = 10
                    End If
                End If

                'If Me.dgvEntries.Rows.Count > 0 AndAlso _
                '   Me.dgvEntries.Rows(Me.dgvEntries.Rows.Count - 1).Cells(Me.ColTrans_Status.Index).Value <> _
                '   Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed) AndAlso _
                '   Me.dgvEntries.Rows.Count <= Me.Limit AndAlso _
                ' Me.Limit > Me.intBatasBaris Then
                '    Me.intBatasBaris = Me.dgvEntries.RowCount + 1
                'Else
                '    Me.intBatasBaris = 10
                'End If

                TambahkanBarisEntries(Me.Limit)
            End If

            If Me.chkLimit.Checked = True Then
                If Me.txtLimit.Enabled = True Then
                    If IsNumeric(Me.txtLimit.Text) Then
                        Me.Limit = CLng(Me.txtLimit.Text)
                        If Me.Baru = True Then
                            Me.dtpEndDate.Value = Me.SettingTanggalAkhir(Me.cboFrequency.SelectedIndex, Me.dtpStartDate.Value, Me.Limit - 1)
                        End If
                    Else
                        Me.Limit = 0
                    End If
                    Me.dtpEndDate.Value = Me.SettingTanggalAkhir(Me.cboFrequency.SelectedIndex, Me.dtpStartDate.Value, Me.Limit - 1)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function TambahkanBarisEntries(ByVal vLimit As Integer) As Boolean
        Dim ObjCD As New Core_Module.ClsCekData
        Try
            Dim RowDGV As New DataGridViewRow


            If Me.dgvEntries.RowCount <= 0 Then
                '=====================================================================================================================================
                '=====================================================================================================================================
                'Jika DgvEntries masih kosong, belum pernah terisi sama sekali
                '=====================================================================================================================================
                If vLimit <= intBatasBaris Then
                    '---------------------------------------------------------------------------------------------------------------------------------
                    '---------------------------------------------------------------------------------------------------------------------------------
                    'Jika Limit <= BatasBaris (10), maka tambahkan baris sesuai Jml Limit
                    '---------------------------------------------------------------------------------------------------------------------------------
                    For IntA As Integer = 0 To vLimit - 1
                        Me.dgvEntries.Rows.Add()
                        RowDGV = dgvEntries.Rows(IntA)
                        With RowDGV
                            .Cells(Me.ColTrans_JournalNumber.Index).Value = Me.TxtRefNo.Text & "-" & Format(IntA + 1, "00#")
                            .Cells(Me.ColTrans_PrimaryJournal.Index).Value = -1
                            Dim dtDate As Date = Me.dtpDate.Value
                            dtDate = Me.SettingTanggalAkhir(Me.cboFrequency.SelectedIndex, dtDate, IntA)
                            dtDate = Me.setTanggal(dtDate, Me.cboFrequency.SelectedIndex, Me.cboEveryWeek.SelectedIndex, Me.cboEveryMonth.SelectedIndex, True)
                            .Cells(Me.ColTrans_JournalDate.Index).Value = Format(dtDate, "dd MMM yyyy")
                            .Cells(Me.ColTrans_Description.Index).Value = Me.TxtRemarks.Text
                            .Cells(Me.ColTrans_Status.Index).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed)
                            .Cells(Me.ColTrans_NoUrut.Index).Value = IntA + 1
                            Me.tms_Delete.Enabled = False
                            Me.tms_DeleteAll.Enabled = False
                            Me.tms_UpdateAll.Enabled = False
                            Me.tms_Update.Enabled = False
                            Me.tms_Edit.Enabled = False
                        End With
                    Next
                    ObjCD.SetNumber(Me.dgvEntries)
                    '---------------------------------------------------------------------------------------------------------------------------------
                Else
                    '---------------------------------------------------------------------------------------------------------------------------------
                    '---------------------------------------------------------------------------------------------------------------------------------
                    'Jika Limit > BatasBaris (10), maka tambahkan baris sesuai Batas Baris
                    '---------------------------------------------------------------------------------------------------------------------------------
                    For IntA As Integer = 0 To intBatasBaris - 1
                        Me.dgvEntries.Rows.Add()
                        RowDGV = dgvEntries.Rows(IntA)
                        With RowDGV
                            .Cells(Me.ColTrans_JournalNumber.Index).Value = Me.TxtRefNo.Text & "-" & Format(IntA + 1, "00#")
                            .Cells(Me.ColTrans_PrimaryJournal.Index).Value = -1
                            Dim dtDate As Date = Me.dtpDate.Value
                            dtDate = Me.SettingTanggalAkhir(Me.cboFrequency.SelectedIndex, dtDate, IntA)
                            dtDate = Me.setTanggal(dtDate, Me.cboFrequency.SelectedIndex, Me.cboEveryWeek.SelectedIndex, Me.cboEveryMonth.SelectedIndex, True)
                            .Cells(Me.ColTrans_JournalDate.Index).Value = Format(dtDate, "dd MMM yyyy")
                            .Cells(Me.ColTrans_Description.Index).Value = Me.TxtRemarks.Text
                            .Cells(Me.ColTrans_Status.Index).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed)
                            .Cells(Me.ColTrans_NoUrut.Index).Value = IntA + 1
                            Me.tms_Delete.Enabled = False
                            Me.tms_DeleteAll.Enabled = False
                            Me.tms_UpdateAll.Enabled = False
                            Me.tms_Update.Enabled = False
                            Me.tms_Edit.Enabled = False
                        End With
                    Next
                    ObjCD.SetNumber(Me.dgvEntries)
                    '---------------------------------------------------------------------------------------------------------------------------------
                End If
                '=====================================================================================================================================
            Else
                '=====================================================================================================================================
                '=====================================================================================================================================
                'Jika DgvEntries sudah ada isinya
                '=====================================================================================================================================
                If Me.dgvEntries.RowCount < vLimit Then
                    Dim intSisa As Integer = 0
                    intSisa = vLimit - CInt(Me.dgvEntries.RowCount)
                    If (CInt(Me.dgvEntries.RowCount) + intSisa) <= intBatasBaris Then
                        For intA As Integer = 0 To intSisa - 1
                            Me.dgvEntries.Rows.Add()
                            RowDGV = dgvEntries.Rows(dgvEntries.Rows.Count - 1)
                            With RowDGV
                                .Cells(Me.ColTrans_JournalNumber.Index).Value = Me.TxtRefNo.Text & "-" & Format(dgvEntries.Rows.Count, "00#")
                                .Cells(Me.ColTrans_PrimaryJournal.Index).Value = -1
                                Dim dtDate As Date = Me.dtpDate.Value
                                dtDate = Me.SettingTanggalAkhir(Me.cboFrequency.SelectedIndex, dtDate, intA)
                                dtDate = Me.setTanggal(dtDate, Me.cboFrequency.SelectedIndex, Me.cboEveryWeek.SelectedIndex, Me.cboEveryMonth.SelectedIndex, True)
                                .Cells(Me.ColTrans_JournalDate.Index).Value = Format(dtDate, "dd MMM yyyy")
                                .Cells(Me.ColTrans_Description.Index).Value = Me.TxtRemarks.Text
                                .Cells(Me.ColTrans_Status.Index).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed)
                                '  .Cells(Me.ColTrans_NoUrut.Index).Value = intA + 1
                                Me.tms_Delete.Enabled = False
                                Me.tms_DeleteAll.Enabled = False
                                Me.tms_UpdateAll.Enabled = False
                                Me.tms_Update.Enabled = False
                                Me.tms_Edit.Enabled = False
                            End With
                        Next
                        ObjCD.SetNumber(Me.dgvEntries)
                    Else
                        For inta As Integer = 0 To CInt(intBatasBaris - CInt(Me.dgvEntries.RowCount)) - 1
                            Me.dgvEntries.Rows.Add()
                            RowDGV = dgvEntries.Rows(dgvEntries.Rows.Count - 1)
                            With RowDGV
                                .Cells(Me.ColTrans_JournalNumber.Index).Value = Me.TxtRefNo.Text & "-" & Format(dgvEntries.Rows.Count, "00#")
                                .Cells(Me.ColTrans_PrimaryJournal.Index).Value = -1
                                Dim dtDate As Date = Me.dtpDate.Value
                                dtDate = Me.SettingTanggalAkhir(Me.cboFrequency.SelectedIndex, dtDate, inta)
                                dtDate = Me.setTanggal(dtDate, Me.cboFrequency.SelectedIndex, Me.cboEveryWeek.SelectedIndex, Me.cboEveryMonth.SelectedIndex, True)
                                .Cells(Me.ColTrans_JournalDate.Index).Value = Format(dtDate, "dd MMM yyyy")
                                .Cells(Me.ColTrans_Description.Index).Value = Me.TxtRemarks.Text
                                .Cells(Me.ColTrans_Status.Index).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed)
                                '.Cells(Me.ColTrans_NoUrut.Index).Value = inta + 1
                                Me.tms_Delete.Enabled = False
                                Me.tms_DeleteAll.Enabled = False
                                Me.tms_UpdateAll.Enabled = False
                                Me.tms_Update.Enabled = False
                                Me.tms_Edit.Enabled = False
                            End With
                        Next
                        ObjCD.SetNumber(Me.dgvEntries)
                    End If
                Else
                    Dim intSelisih As Integer = 0
                    intSelisih = Me.dgvEntries.Rows.Count - vLimit
                    For Inta As Integer = 0 To intSelisih - 1
                        RowDGV = dgvEntries.Rows(dgvEntries.Rows.Count - 1)
                        With RowDGV
                            If .Cells(Me.ColTrans_Status.Index).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed) Then
                                'RowDGV.Dispose()
                                Me.dgvEntries.Rows.RemoveAt(dgvEntries.Rows.Count - 1)
                            Else
                                MsgBox("The last transaction entries status is not 'planned', so you can't input limit data < " & Me.dgvEntries.RowCount, MsgBoxStyle.Information, "Information")
                                Me.Limit = Me.dgvEntries.RowCount
                                If Me.chkLimit.Checked = True Then
                                    Me.txtLimit.Text = Me.dgvEntries.RowCount
                                End If
                                Exit For
                            End If
                        End With
                    Next
                    'MsgBox("cannot input limit < " & Me.dgvEntries.RowCount, MsgBoxStyle.Information, "Information")
                    'Return False
                    'Exit Function
                End If
                '=====================================================================================================================================
            End If

            If Me.dgvEntries.Rows.Count > 0 Then
                For intb As Integer = 0 To Me.dgvEntries.Rows.Count - 1
                    RowDGV = dgvEntries.Rows(intb)
                    With RowDGV
                        .Cells(Me.ColTrans_NoUrut.Index).Value = intb + 1
                        If .Cells(Me.ColTrans_Status.Index).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed) OrElse _
                           .Cells(Me.ColTrans_Status.Index).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Skiped) Then
                            .Cells(Me.ColTrans_Description.Index).Value = Me.TxtRemarks.Text
                            Dim dtDate As Date = Me.dtpStartDate.Value
                            dtDate = Me.SettingTanggalAkhir(Me.cboFrequency.SelectedIndex, dtDate, intb)
                            dtDate = Me.setTanggal(dtDate, Me.cboFrequency.SelectedIndex, Me.cboEveryWeek.SelectedIndex, Me.cboEveryMonth.SelectedIndex, True)
                            .Cells(Me.ColTrans_JournalDate.Index).Value = Format(dtDate, "dd MMM yyyy")
                        End If
                    End With
                Next
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub TxtRemarks_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus
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


    Private Sub TxtRemarks_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.Leave

        Try
            If Me.TxtRemarks.Enabled = True Then
                TambahkanBarisEntries(CInt(Me.Limit))
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEntries_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvEntries.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColTrans_Remarks.Index Then
                e.Cancel = True
            Else
                If Not String.IsNullOrEmpty(Me.dgvEntries.Item(Me.ColRemarks.Index, e.RowIndex).Value) Then
                    Me.BufferTrans = Me.dgvEntries.Item(Me.ColRemarks.Index, e.RowIndex).Value
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub dgvEntries_CellMouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvEntries.CellMouseDown
        Try
            If Me.btnOK.Enabled = True Then
                Me.tms_Create.Enabled = False
                Me.tms_Skip.Enabled = False
                Me.tms_Delete.Enabled = False
                Me.tms_DeleteAll.Enabled = False
                Me.tms_UpdateAll.Enabled = False
                Me.tms_Update.Enabled = False
                Me.tms_Edit.Enabled = False
                Me.tms_Lock.Enabled = False
            Else
                If e.Button = Windows.Forms.MouseButtons.Right Then
                    If Me.dgvEntries.Rows.Count > 0 Then
                        Me.rowIndex = e.RowIndex
                        If Me.rowIndex > -1 Then
                            If Me.dgvEntries.Item(Me.ColTrans_Status.Index, Me.rowIndex).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed) Then
                                If Me.rowIndex = 0 Then
                                    Me.tms_Create.Enabled = True
                                    Me.tms_Skip.Enabled = True
                                    Me.tms_Delete.Enabled = False
                                    Me.tms_DeleteAll.Enabled = False
                                    Me.tms_UpdateAll.Enabled = False
                                    Me.tms_Update.Enabled = False
                                    Me.tms_Edit.Enabled = False
                                    Me.tms_Lock.Enabled = False
                                Else
                                    If Me.dgvEntries.Item(Me.ColTrans_Status.Index, (Me.rowIndex - 1)).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed) Then
                                        Me.tms_Create.Enabled = False
                                        Me.tms_Skip.Enabled = False
                                        Me.tms_Delete.Enabled = False
                                        Me.tms_DeleteAll.Enabled = False
                                        Me.tms_UpdateAll.Enabled = False
                                        Me.tms_Update.Enabled = False
                                        Me.tms_Edit.Enabled = False
                                        Me.tms_Lock.Enabled = False
                                    Else
                                        Me.tms_Create.Enabled = True
                                        Me.tms_Skip.Enabled = True
                                        Me.tms_Delete.Enabled = False
                                        Me.tms_DeleteAll.Enabled = False
                                        Me.tms_UpdateAll.Enabled = False
                                        Me.tms_Update.Enabled = False
                                        Me.tms_Edit.Enabled = False
                                        Me.tms_Lock.Enabled = False
                                    End If
                                End If
                            ElseIf Me.dgvEntries.Item(Me.ColTrans_Status.Index, Me.rowIndex).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Skiped) Then
                                If e.RowIndex = Me.dgvEntries.Rows.Count - 1 Then
                                    Me.tms_Create.Enabled = True
                                    Me.tms_Skip.Enabled = False
                                    Me.tms_Delete.Enabled = True
                                    Me.tms_DeleteAll.Enabled = True
                                    Me.tms_UpdateAll.Enabled = False
                                    Me.tms_Update.Enabled = False
                                    Me.tms_Edit.Enabled = False
                                    Me.tms_Lock.Enabled = False
                                Else
                                    If Me.dgvEntries.Item(Me.ColTrans_Status.Index, Me.rowIndex + 1).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed) Then
                                        Me.tms_Create.Enabled = True
                                        Me.tms_Skip.Enabled = False
                                        Me.tms_Delete.Enabled = True
                                        Me.tms_DeleteAll.Enabled = True
                                        Me.tms_UpdateAll.Enabled = False
                                        Me.tms_Update.Enabled = False
                                        Me.tms_Edit.Enabled = False
                                        Me.tms_Lock.Enabled = False
                                    Else
                                        Me.tms_Create.Enabled = False
                                        Me.tms_Skip.Enabled = False
                                        Me.tms_Delete.Enabled = False
                                        Me.tms_DeleteAll.Enabled = False
                                        Me.tms_UpdateAll.Enabled = False
                                        Me.tms_Update.Enabled = False
                                        Me.tms_Edit.Enabled = False
                                        Me.tms_Lock.Enabled = False
                                    End If
                                End If
                            ElseIf Me.dgvEntries.Item(Me.ColTrans_Status.Index, Me.rowIndex).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Created) Then
                                If e.RowIndex = Me.dgvEntries.Rows.Count - 1 Then
                                    Me.tms_Create.Enabled = False
                                    Me.tms_Skip.Enabled = False
                                    Me.tms_Delete.Enabled = True
                                    Me.tms_DeleteAll.Enabled = True
                                    Me.tms_UpdateAll.Enabled = True
                                    Me.tms_Update.Enabled = True
                                    Me.tms_Edit.Enabled = True
                                    Me.tms_Lock.Enabled = True
                                Else
                                    If Me.dgvEntries.Item(Me.ColTrans_Status.Index, Me.rowIndex + 1).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed) Then
                                        Me.tms_Create.Enabled = False
                                        Me.tms_Skip.Enabled = False
                                        Me.tms_Delete.Enabled = True
                                        Me.tms_DeleteAll.Enabled = True
                                        Me.tms_UpdateAll.Enabled = True
                                        Me.tms_Update.Enabled = True
                                        Me.tms_Edit.Enabled = True
                                        Me.tms_Lock.Enabled = True
                                    Else
                                        Me.tms_Create.Enabled = False
                                        Me.tms_Skip.Enabled = False
                                        Me.tms_Delete.Enabled = False
                                        Me.tms_DeleteAll.Enabled = False
                                        Me.tms_UpdateAll.Enabled = False
                                        Me.tms_Update.Enabled = False
                                        Me.tms_Edit.Enabled = False
                                        Me.tms_Lock.Enabled = False
                                    End If
                                End If
                            ElseIf Me.dgvEntries.Item(Me.ColTrans_Status.Index, Me.rowIndex).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Locked) Then
                                Me.tms_Create.Enabled = False
                                Me.tms_Skip.Enabled = False
                                Me.tms_Delete.Enabled = False
                                Me.tms_DeleteAll.Enabled = False
                                Me.tms_UpdateAll.Enabled = False
                                Me.tms_Update.Enabled = False
                                Me.tms_Edit.Enabled = False
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboEveryWeek_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboEveryWeek.SelectedIndexChanged

        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.dtpStartDate.Value = Me.SettingTanggalAwal(Me.cboFrequency.SelectedIndex, Me.cboEveryWeek.SelectedIndex, Me.cboEveryMonth.SelectedIndex, Me.dtpStartDate.Value)
                If Me.chkLimit.Checked = True Then
                    If IsNumeric(Me.txtLimit.Text) Then
                        Me.Limit = CLng(Me.txtLimit.Text)
                    Else
                        Me.Limit = 0
                    End If
                    Me.dtpEndDate.Value = Me.SettingTanggalAkhir(Me.cboFrequency.SelectedIndex, Me.dtpStartDate.Value, Me.Limit)
                    Me.TambahkanBarisEntries(Me.Limit)
                Else
                    Me.Limit = Me.GetLimit(Me.dtpStartDate.Value, Me.dtpEndDate.Value)
                    Me.TambahkanBarisEntries(Me.Limit)
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboEveryMonth_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboEveryMonth.SelectedIndexChanged

        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.dtpStartDate.Value = Me.SettingTanggalAwal(Me.cboFrequency.SelectedIndex, Me.cboEveryWeek.SelectedIndex, Me.cboEveryMonth.SelectedIndex, Me.dtpStartDate.Value)
                If Me.chkLimit.Checked = True Then
                    If IsNumeric(Me.txtLimit.Text) Then
                        Me.Limit = CLng(Me.txtLimit.Text)
                    Else
                        Me.Limit = 0
                    End If
                    Me.dtpEndDate.Value = Me.SettingTanggalAkhir(Me.cboFrequency.SelectedIndex, Me.dtpStartDate.Value, Me.Limit)
                    Me.TambahkanBarisEntries(Me.Limit)
                Else
                    Me.Limit = Me.GetLimit(Me.dtpStartDate.Value, Me.dtpEndDate.Value)
                    Me.TambahkanBarisEntries(Me.Limit)
                End If
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpStartDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpStartDate.GotFocus

        Try
            Me.BufferTglAwal = Me.dtpStartDate.Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub dtpEndDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpEndDate.GotFocus

        Try
            Me.BufferTglAkhir = Me.dtpEndDate.Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub dtpStartDate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpStartDate.Leave
        Try
            If Me.chkLimit.Checked = False Then
                If Me.dtpStartDate.Value.Date > Me.dtpEndDate.Value.Date Then
                    MsgBox("Please make sure start date > end date.", MsgBoxStyle.Information, "Information")
                    Me.dtpStartDate.Value = Me.BufferTglAwal
                    Me.dtpStartDate.Focus()
                    Exit Sub
                End If
            End If

            If Me.TxtRemarks.Enabled = True Then
                If IsDate(Me.dtpStartDate.Value) Then
                    If Me.chkLimit.Checked = False Then
                        Me.Limit = Me.GetLimit(Me.dtpStartDate.Value, Me.dtpEndDate.Value)
                    Else
                        If IsNumeric(Me.txtLimit.Text) Then
                            Me.Limit = CLng(Me.txtLimit.Text)
                        Else
                            Me.Limit = 0
                        End If
                    End If


                    If Me.dgvEntries.Rows.Count > 0 Then
                        If Me.Limit > Me.intBatasBaris AndAlso _
                        Me.dgvEntries.Rows(Me.dgvEntries.Rows.Count - 1).Cells(Me.ColTrans_Status.Index).Value <> _
                        Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed) Then
                            Me.intBatasBaris = Me.dgvEntries.RowCount + 1
                        Else
                            Me.intBatasBaris = 10
                        End If
                    End If

                    'If Me.dgvEntries.Rows.Count > 0 AndAlso _
                    ' Me.dgvEntries.Rows.Count <= Me.Limit AndAlso _
                    ' Me.dgvEntries.Rows(Me.dgvEntries.Rows.Count - 1).Cells(Me.ColTrans_Status.Index).Value <> _
                    ' Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed) AndAlso _
                    ' Me.Limit > Me.intBatasBaris Then
                    '    Me.intBatasBaris = Me.dgvEntries.Rows.Count + 1
                    'End If

                    Me.TambahkanBarisEntries(Me.Limit)
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpStartDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpStartDate.ValueChanged
        Try
            If Me.chkLimit.Checked = False Then
                If Me.dtpStartDate.Value.Date > Me.dtpEndDate.Value.Date Then
                    MsgBox("Please make suke start date <= end date.", MsgBoxStyle.Information, "Information")
                    Me.dtpStartDate.Value = Me.BufferTglAwal
                    Me.dtpStartDate.Focus()
                    Exit Sub
                End If
            End If

            If Me.TxtRemarks.Enabled = True Then
                If IsDate(Me.dtpStartDate.Value) Then
                    If Me.chkLimit.Checked = False Then
                        Me.Limit = Me.GetLimit(Me.dtpStartDate.Value, Me.dtpEndDate.Value)
                    Else
                        If IsNumeric(Me.txtLimit.Text) Then
                            Me.Limit = CLng(Me.txtLimit.Text)
                        Else
                            Me.Limit = 0
                        End If
                    End If

                    'If Me.dgvEntries.Rows.Count > 0 AndAlso _
                    ' Me.dgvEntries.Rows.Count <= Me.Limit AndAlso _
                    ' Me.dgvEntries.Rows(Me.dgvEntries.Rows.Count - 1).Cells(Me.ColTrans_Status.Index).Value <> _
                    ' Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed) AndAlso _
                    ' Me.Limit > Me.intBatasBaris Then
                    '    Me.intBatasBaris = Me.dgvEntries.Rows.Count + 1
                    'End If


                    If Me.dgvEntries.Rows.Count > 0 Then
                        If Me.Limit > Me.intBatasBaris AndAlso _
                        Me.dgvEntries.Rows(Me.dgvEntries.Rows.Count - 1).Cells(Me.ColTrans_Status.Index).Value <> _
                        Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed) Then
                            Me.intBatasBaris = Me.dgvEntries.RowCount + 1
                        Else
                            Me.intBatasBaris = 10
                        End If
                    End If


                    Me.TambahkanBarisEntries(Me.Limit)
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub dtpEndDate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpEndDate.Leave

    '    Try
    '        If Me.chkLimit.Checked = False Then
    '            If Me.dtpEndDate.Value.DayOfYear < Me.dtpStartDate.Value.DayOfYear Then
    '                MsgBox("cannot input start date > end date, it could be error.", MsgBoxStyle.Information, "Information")
    '                Me.dtpEndDate.Value = Me.BufferTglAkhir
    '                Me.dtpEndDate.Focus()
    '                Exit Sub
    '            End If
    '        End If

    '        If Me.TxtRemarks.Enabled = True Then
    '            If IsDate(Me.dtpEndDate.Value) Then
    '                If Me.chkLimit.Checked = False Then
    '                    Me.Limit = Me.GetLimit(Me.dtpStartDate.Value, Me.dtpEndDate.Value)
    '                Else
    '                    If IsNumeric(Me.txtLimit.Text) Then
    '                        Me.Limit = Clng(Me.txtLimit.Text)
    '                    Else
    '                        Me.Limit = 0
    '                    End If
    '                End If


    '                If Me.dgvEntries.Rows.Count > 0 Then
    '                    If Me.Limit > Me.intBatasBaris AndAlso _
    '                    Me.dgvEntries.Rows(Me.dgvEntries.Rows.Count - 1).Cells(Me.ColTrans_Status.Index).Value <> _
    '                    Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed) Then
    '                        Me.intBatasBaris = Me.dgvEntries.RowCount + 1
    '                    Else
    '                        Me.intBatasBaris = 10
    '                    End If
    '                End If

    '                'If Me.dgvEntries.Rows.Count > 0 AndAlso _
    '                '   Me.dgvEntries.Rows.Count <= Me.Limit AndAlso _
    '                '   Me.dgvEntries.Rows(Me.dgvEntries.Rows.Count - 1).Cells(Me.ColTrans_Status.Index).Value <> _
    '                '   Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed) AndAlso _
    '                '   Me.Limit > Me.intBatasBaris Then
    '                '    Me.intBatasBaris = Me.dgvEntries.Rows.Count + 1

    '                'End If

    '                Me.TambahkanBarisEntries(Me.Limit)
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try

    'End Sub

    Private Sub dtpEndDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpEndDate.ValueChanged
        Try
            If Me.chkLimit.Checked = False Then
                If Me.TxtRemarks.Enabled = True Then
                    If Me.dtpStartDate.Value.Date > Me.dtpEndDate.Value.Date Then
                        MsgBox("Please make sure start date <= end date.", MsgBoxStyle.Information, "Information")
                        Me.dtpEndDate.Value = Me.BufferTglAkhir
                        Me.dtpEndDate.Focus()
                        Exit Sub
                    End If
                End If
            End If


            If Me.TxtRemarks.Enabled = True Then
                If IsDate(Me.dtpEndDate.Value) Then
                    If Me.chkLimit.Checked = False Then
                        Me.Limit = Me.GetLimit(Me.dtpStartDate.Value, Me.dtpEndDate.Value)
                    Else
                        If IsNumeric(Me.txtLimit.Text) Then
                            Me.Limit = CLng(Me.txtLimit.Text)
                        Else
                            Me.Limit = 0
                        End If
                    End If


                    If Me.dgvEntries.Rows.Count > 0 Then
                        If Me.Limit > Me.intBatasBaris AndAlso _
                        Me.dgvEntries.Rows(Me.dgvEntries.Rows.Count - 1).Cells(Me.ColTrans_Status.Index).Value <> _
                        Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed) Then
                            Me.intBatasBaris = Me.dgvEntries.RowCount + 1
                        Else
                            Me.intBatasBaris = 10
                        End If
                    End If

                    'If Me.dgvEntries.Rows.Count > 0 AndAlso _
                    '   Me.dgvEntries.Rows.Count <= Me.Limit AndAlso _
                    '   Me.dgvEntries.Rows(Me.dgvEntries.Rows.Count - 1).Cells(Me.ColTrans_Status.Index).Value <> _
                    '   Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed) AndAlso _
                    '   Me.Limit > Me.intBatasBaris Then
                    '    Me.intBatasBaris = Me.dgvEntries.Rows.Count + 1

                    'End If

                    Me.TambahkanBarisEntries(Me.Limit)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUP_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUP.Click
        Try
            Me.dgvItem.AllowUserToAddRows = False
            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItem.SelectedRows(0)

                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.dgvItem.Rows.Remove(row)
                    Me.dgvItem.Rows.InsertRange(indexbaru, row)
                    Me.dgvItem.Focus()
                    Me.dgvItem.ClearSelection()
                    row.Selected = True
                    ''Me.ObjCD.SetNumber(Me.dgvItem)
                    If row.Displayed = False Then
                        Me.ShowRowProcess(row, True)
                    End If
                End If

                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
            Me.dgvItem.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            Me.dgvItem.AllowUserToAddRows = False
            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False

                Dim row As DataGridViewRow = Me.dgvItem.SelectedRows(0)
                If row.Index <> Me.dgvItem.Rows.Count - 1 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.dgvItem.Rows.Remove(row)
                    Me.dgvItem.Rows.InsertRange(indexbaru, row)
                    Me.dgvItem.Focus()
                    Me.dgvItem.ClearSelection()
                    row.Selected = True
                    ''Me.ObjCD.SetNumber(Me.dgvItem)
                    Me.ShowRowProcess(row, False)
                End If

                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
            Me.dgvItem.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

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

    Private Sub TabControl1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.Click

        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TabPage1.Focus = True Then
                    If Me.dgvItem.Rows.Count > 0 AndAlso Me.rowIndexItem > 0 Then
                        Me.btnUP.Enabled = True
                        Me.btnDown.Enabled = True
                    Else
                        Me.btnUP.Enabled = False
                        Me.btnDown.Enabled = False
                    End If
                ElseIf Me.TabPage2.Focus = True Then
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                ElseIf Me.TabPage3.Focus = True Then
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvReminder_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvReminder.CellBeginEdit
        Dim RowDGV As DataGridViewRow
        Try
            RowDGV = Me.dgvReminder.Rows(e.RowIndex)
            Me.BufferReminder = ""

            If e.ColumnIndex <> Me.ColReminder_No.Index Then
                If e.ColumnIndex = Me.ColReminder_Remarks.Index OrElse e.ColumnIndex = Me.ColReminder_Interval.Index Then
                    If String.IsNullOrEmpty(Me.dgvReminder.Item(Me.ColReminder_UserCode.Index, e.RowIndex).Value) Then
                        MsgBox("Please input user code first.", MsgBoxStyle.Information, "Information")
                        e.Cancel = True
                        Exit Sub
                    End If
                End If
                If e.ColumnIndex = Me.ColReminder_UserCode.Index OrElse e.ColumnIndex = Me.ColReminder_Remarks.Index Then
                    If Not String.IsNullOrEmpty(RowDGV.Cells(e.ColumnIndex).Value) Then
                        If ObjCD.CekDataKosong(RowDGV.Cells(e.ColumnIndex).Value, True) = True Then
                            Me.BufferReminder = RowDGV.Cells(e.ColumnIndex).Value
                        End If
                    Else
                        Me.BufferReminder = RowDGV.Cells(e.ColumnIndex).Value
                    End If
                   

                    'If Not String.IsNullOrEmpty(RowDGV.Cells(e.ColumnIndex).Value) Then
                    '    Me.BufferReminder = RowDGV.Cells(e.ColumnIndex).Value
                    'End If
                End If

                If e.ColumnIndex = Me.ColReminder_Interval.Index Then
                    If IsNumeric(RowDGV.Cells(e.ColumnIndex).Value) Then
                        Me.BufferReminder = CInt(RowDGV.Cells(e.ColumnIndex).Value)
                    End If
                End If
            Else
                e.Cancel = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvReminder_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvReminder.CellEndEdit
        Dim objUser As New Core_Module.ClsUser
        Dim RowDGV As DataGridViewRow
        Dim ij As New Core_Module.FrmCariMasterUser
        Try
            RowDGV = Me.dgvReminder.Rows(e.RowIndex)

            If e.ColumnIndex = Me.ColReminder_UserCode.Index Then
                If Not String.IsNullOrEmpty(Me.dgvReminder.Item(e.ColumnIndex, e.RowIndex).Value) AndAlso _
                   Me.ObjCD.CekDataKosong(Me.dgvReminder.Item(e.ColumnIndex, e.RowIndex).Value, True) Then
                    Dim Primary As Double = -1
                    Primary = objUser.GetPrimary(Me.dgvReminder.Item(e.ColumnIndex, e.RowIndex).Value)
                    If Primary > -1 Then
                        Me.dgvReminder.Item(e.ColumnIndex, e.RowIndex).Value = objUser.GetUserCode(Primary).ToUpper
                        Me.dgvReminder.Item(Me.ColReminder_PrimaryUser.Index, e.RowIndex).Value = Primary
                        Me.dgvReminder.Item(Me.ColReminder_Interval.Index, e.RowIndex).Value = 0
                    Else
                        ij.TxtCari.Text = Me.dgvReminder.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.Kodeuser <> "" Then
                            Me.dgvReminder.Item(Me.ColReminder_PrimaryUser.Index, e.RowIndex).Value = objUser.GetPrimary(ij.Kodeuser)
                            Me.dgvReminder.Item(e.ColumnIndex, e.RowIndex).Value = ij.Kodeuser.ToUpper
                            Me.dgvReminder.Item(Me.ColReminder_Interval.Index, e.RowIndex).Value = 0
                        Else
                            If Not String.IsNullOrEmpty(Me.BufferReminder) Then
                                Me.dgvReminder.Item(e.ColumnIndex, e.RowIndex).Value = Me.BufferReminder
                                Me.dgvReminder.Item(Me.ColReminder_PrimaryUser.Index, e.RowIndex).Value = objUser.GetPrimary(Me.BufferReminder)
                                Me.dgvReminder.Item(Me.ColReminder_Interval.Index, e.RowIndex).Value = 0
                            Else
                                If e.RowIndex <> Me.dgvReminder.Rows.Count - 1 Then
                                    Me.dgvReminderHapus = True
                                    Me.rowIndexReminder = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                Else
                    If Not String.IsNullOrEmpty(Me.BufferReminder) Then
                        Me.dgvReminder.Item(e.ColumnIndex, e.RowIndex).Value = Me.BufferReminder
                        Me.dgvReminder.Item(Me.ColReminder_PrimaryUser.Index, e.RowIndex).Value = objUser.GetPrimary(Me.BufferReminder)
                        Me.dgvReminder.Item(Me.ColReminder_Interval.Index, e.RowIndex).Value = 0
                    Else
                        If e.RowIndex <> Me.dgvReminder.Rows.Count - 1 Then
                            Me.dgvReminderHapus = True
                            Me.rowIndexReminder = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                            Exit Sub
                        End If
                    End If
                End If
            End If

            If e.ColumnIndex = Me.ColReminder_Interval.Index Then
                If IsNumeric(Me.dgvReminder.Item(e.ColumnIndex, e.RowIndex).Value) Then
                    If CInt(Me.dgvReminder.Item(e.ColumnIndex, e.RowIndex).Value) > 30 Then
                        MsgBox("Maximum value is 30 days", MsgBoxStyle.Information, "Information")
                        If IsNumeric(Me.BufferReminder) Then
                            Me.dgvReminder.Item(Me.ColReminder_Interval.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.BufferReminder, 0)
                        Else
                            Me.dgvReminder.Item(Me.ColReminder_Interval.Index, e.RowIndex).Value = 0
                        End If
                        Exit Sub
                    End If
                    Me.dgvReminder.Item(Me.ColReminder_Interval.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvReminder.Item(e.ColumnIndex, e.RowIndex).Value, 0)
                Else
                    If IsNumeric(Me.BufferReminder) Then
                        Me.dgvReminder.Item(Me.ColReminder_Interval.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.BufferReminder, 0)
                    Else
                        Me.dgvReminder.Item(Me.ColReminder_Interval.Index, e.RowIndex).Value = 0
                    End If
                End If
            End If

            ''cek ColRemarks
            If e.ColumnIndex = Me.ColReminder_Remarks.Index Then
                If Not String.IsNullOrEmpty(Me.dgvReminder.Item(e.ColumnIndex, e.RowIndex).Value) Then
                    If Me.ObjCD.CekDataKosong(Me.dgvReminder.Item(e.ColumnIndex, e.RowIndex).Value.ToString, False) = False Then
                        If Not String.IsNullOrEmpty(Me.BufferReminder) Then
                            Me.dgvReminder.Item(e.ColumnIndex, e.RowIndex).Value = Me.BufferReminder
                        Else
                            Me.dgvReminder.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If

            End If

            If e.ColumnIndex = Me.ColReminder_UserCode.Index Then
                If RowDGV.Cells(Me.ColReminder_UserCode.Index).Value <> "" Then
                    If Me.CekUserSama(RowDGV.Cells(Me.ColReminder_UserCode.Index).Value, e.RowIndex) = True Then
                        MsgBox("User code duplication, please choose another user", MsgBoxStyle.Information, "Information")
                        'Me.Rowindex = Rows.Index
                        If Me.Buffer = "" Then
                            If e.RowIndex <> Me.dgvReminder.Rows.Count - 1 Then
                                Me.dgvReminderHapus = True
                                Me.rowIndexReminder = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                                Exit Sub
                            End If
                            'Me.dgvItem.Rows.RemoveAt(e.RowIndex)
                        Else
                            ' Rows.Cells(Me.ColProductCode.Index).Value = Me.buffer
                            RowDGV.Cells(Me.ColReminder_UserCode.Index).Value = Me.BufferReminder
                            Me.dgvItem.AllowUserToAddRows = True
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvItem)
                        'Me.DgvItem2.Rows.Clear()
                        'Me.Sama = False
                    End If
                End If
            End If

            Me.ObjCD.SetNumber(Me.dgvReminder)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvReminder_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvReminder.UserDeletedRow

        Try
            If Me.dgvReminder.Rows.Count > 0 Then
                Me.ObjCD.SetNumber(Me.dgvReminder)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvReminder_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvReminder.UserDeletingRow

        Try
            If Me.dgvReminder.Rows.Count > 0 Then
                Me.ObjCD.SetNumber(Me.dgvReminder)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted

        Try
            'Me.baru = False
            If Me.dgvReminderHapus = True Then
                Me.dgvReminder.Rows.RemoveAt(Me.rowIndexReminder)
                Me.ObjCD.SetNumber(Me.dgvReminder)
                Me.dgvReminderHapus = False
            End If

            If Me.dgvItemHapus = True Then
                Me.dgvItem.Rows.RemoveAt(Me.rowIndexItem)
                'Me.ObjCD.SetNumber(Me.dgvItem)
                Me.dgvItemHapus = False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function ApakahSudahClosing(ByVal tglDari As Date, ByVal tglSampai As Date, ByVal strJenisAction As String) As Boolean

        Try
            'jika sudah ada closing jurnal untuk tahun ybs maka tidak bisa input data baru
            If Me.ObjJournal.GetJml(ClsTransaksiJournal.EnumTipeJournal.ClosingJournal, tglDari, tglSampai) > 0 Then
                MsgBox("There are closing jurnal for this year you can't " & strJenisAction & " this data", MsgBoxStyle.Information, "Information")
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Function


    Public Sub tms_Create_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tms_Create.Click
        Dim DTAUpdate As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journal_transaction_entriesTableAdapter
        Try
            If Not String.IsNullOrEmpty(Me.dgvEntries.Item(Me.ColTrans_JournalNumber.Index, Me.rowIndex).Value) Then
                If ApakahDataBenarJournal(Me.TxtUnbalance.Text, Me.dgvItem) = True Then
                    Dim Urutan As Integer = 0
                    Urutan = Me.GetUrutanJournal(Me.rowIndex)
                    Dim PrimaryJournal As Double = -1
                    PrimaryJournal = Me.dgvEntries.Item(Me.ColTrans_PrimaryJournal.Index, Me.rowIndex).Value
                    Dim tglDari As Date = Date.Now
                    Dim tglSampai As Date = Date.Now
                    Dim tgljurnal As Date = Date.Now

                    If IsDate(Me.dgvEntries.Item(Me.ColTrans_JournalDate.Index, Me.rowIndex).Value) Then
                        tgljurnal = CDate(Me.dgvEntries.Item(Me.ColTrans_JournalDate.Index, Me.rowIndex).Value)
                    End If

                    tglDari = ObjCD.GetFisrtDateOfTheYear(tgljurnal)
                    tglSampai = ObjCD.GetLastDateOfTheYear(tgljurnal)
                    'tglDari = ObjCD.GetFirstDateofTheMonth(tgljurnal)
                    'tglSampai = tglDari.AddDays(Date.DaysInMonth(tgljurnal.Year, tgljurnal.Month) - 1)

                    If Me.ApakahSudahClosing(tglDari, tglSampai, "create") = False Then
                        If Me.IsiKeJournal(Me.rowIndex, Urutan, PrimaryJournal, True, ClsTransaksiRecurring.EnumAction.Create) = True Then
                            'MsgBox("Data saved.", MsgBoxStyle.Information, "Information")
                            'Me.dgvEntries.Item(Me.ColTrans_Status.Index, Me.rowIndex).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Created)
                            If Me.Limit > (Me.dgvEntries.Rows.Count) AndAlso Me.rowIndex = Me.dgvEntries.Rows.Count - 1 Then
                                Me.intBatasBaris = Me.dgvEntries.Rows.Count + 1
                                Me.TambahkanBarisEntries(Me.Limit)
                                Dim IntRowX As Integer = 0
                                IntRowX = CInt(Me.dgvEntries.Rows.Count - 1)
                                If String.IsNullOrEmpty(Me.dgvEntries.Item(Me.ColTrans_Remarks.Index, IntRowX).Value) Then
                                    Me.dgvEntries.Item(Me.ColTrans_Remarks.Index, IntRowX).Value = ""
                                End If
                                DTAUpdate.Insert(Me.TxtPrimaryJournal.Text, _
                                                      Me.dgvEntries.Item(Me.ColTrans_PrimaryJournal.Index, IntRowX).Value, _
                                                      Me.dgvEntries.Item(Me.ColTrans_JournalNumber.Index, IntRowX).Value, _
                                                      Me.dgvEntries.Item(Me.ColTrans_JournalDate.Index, IntRowX).Value, _
                                                      Me.dgvEntries.Item(Me.ColTrans_Description.Index, IntRowX).Value, _
                                                      Me.dgvEntries.Item(Me.ColTrans_Remarks.Index, IntRowX).Value, _
                                                      ClsTransaksiRecurring.EnumStatus.Planed, _
                                                      Me.dgvEntries.Item(Me.ColTrans_NoUrut.Index, IntRowX).Value)
                                Me.ObjRecurring.CariDetail(Me.TxtPrimaryJournal.Text, Me.dgvItem, Me)
                                Me.SalinkeField()
                            End If
                        Else
                            'MsgBox("Failed to save data.", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAUpdate.Dispose()
        End Try
    End Sub

    Private Sub tms_Delete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tms_Delete.Click
        Dim DTAUpdate As New DseTransaksiRecurringTableAdapters.Khusus_Update_Trans_EntriesTableAdapter

        Try
            If Not String.IsNullOrEmpty(Me.dgvEntries.Item(Me.ColTrans_JournalNumber.Index, Me.rowIndex).Value) Then
                Dim Urutan As Integer = 0
                Urutan = Me.GetUrutanJournal(Me.rowIndex)
                Dim PrimaryJournal As Double = -1
                PrimaryJournal = Me.dgvEntries.Item(Me.ColTrans_PrimaryJournal.Index, Me.rowIndex).Value
                If Me.dgvEntries.Item(Me.ColTrans_Status.Index, Me.rowIndex).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Created) Then
                    Dim tglDari As Date = Date.Now
                    Dim tglSampai As Date = Date.Now
                    Dim tgljurnal As Date = Date.Now

                    If IsDate(Me.dgvEntries.Item(Me.ColTrans_JournalDate.Index, Me.rowIndex).Value) Then
                        tgljurnal = CDate(Me.dgvEntries.Item(Me.ColTrans_JournalDate.Index, Me.rowIndex).Value)
                    End If

                    'tglDari = ObjCD.GetFirstDateofTheMonth(tgljurnal)
                    'tglSampai = ObjCD.GetLasDateOfTheMonth(tgljurnal)

                    tglDari = ObjCD.GetFisrtDateOfTheYear(tgljurnal)
                    tglSampai = ObjCD.GetLastDateOfTheYear(tgljurnal)


                    If Me.ApakahSudahClosing(tglDari, tglSampai, "delete") = False AndAlso Me.ObjJournal.ApakahBolehHapus(PrimaryJournal, True, Me.UserCOde, True) = True Then
                        If Me.IsiKeJournal(Me.rowIndex, Urutan, PrimaryJournal, False, ClsTransaksiRecurring.EnumAction.Delete) = True Then
                            MsgBox("Journal deleted.", MsgBoxStyle.Information, "Information")
                        Else
                            MsgBox("Failed to delete journal.", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If
                    End If
                ElseIf Me.dgvEntries.Item(Me.ColTrans_Status.Index, Me.rowIndex).Value = _
                       Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Skiped) Then

                    DTAUpdate.UpdateQuery(PrimaryJournal, _
                    Me.dgvEntries.Item(Me.ColTrans_JournalNumber.Index, Me.rowIndex).Value, _
                    Me.dgvEntries.Item(Me.ColTrans_JournalDate.Index, Me.rowIndex).Value, _
                    Me.dgvEntries.Item(Me.ColTrans_Description.Index, Me.rowIndex).Value, _
                    ClsTransaksiRecurring.EnumStatus.Planed, _
                    Me.dgvEntries.Item(Me.ColTrans_NoUrut.Index, Me.rowIndex).Value, _
                    Me.dgvEntries.Item(Me.ColTrans_JournalNumber.Index, Me.rowIndex).Value)
                    Me.ObjRecurring.CariDetail(Me.TxtPrimaryJournal.Text, Me.dgvItem, Me)
                    Me.SalinkeField()
                End If
                End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAUpdate.Dispose()
        End Try
    End Sub

    Private Sub tms_Edit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tms_Edit.Click
        Try
            If Not String.IsNullOrEmpty(Me.dgvEntries.Item(Me.ColTrans_JournalNumber.Index, Me.rowIndex).Value) Then
                'Me.dgvEntries.Item(Me.ColTrans_Status.Index, Me.rowIndex).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed)
                'Me.TambahkanBarisEntries(Me.Limit)
                Dim Urutan As Integer = 0
                Urutan = Me.GetUrutanJournal(Me.rowIndex)
                Dim PrimaryJournal As Double = -1
                PrimaryJournal = Me.dgvEntries.Item(Me.ColTrans_PrimaryJournal.Index, Me.rowIndex).Value
                Dim tglDari As Date = Date.Now
                Dim tglSampai As Date = Date.Now
                Dim tgljurnal As Date = Date.Now

                If IsDate(Me.dgvEntries.Item(Me.ColTrans_JournalDate.Index, Me.rowIndex).Value) Then
                    tgljurnal = CDate(Me.dgvEntries.Item(Me.ColTrans_JournalDate.Index, Me.rowIndex).Value)
                End If

                'tglDari = ObjCD.GetFirstDateofTheMonth(tgljurnal)
                'tglSampai = ObjCD.GetLastDateOfTheMonths(tgljurnal)
                tglDari = ObjCD.GetFisrtDateOfTheYear(tgljurnal)
                tglSampai = ObjCD.GetLastDateOfTheYear(tgljurnal)

                If Me.ApakahSudahClosing(tglDari, tglSampai, "edit") = False AndAlso Me.ObjJournal.ApakahBolehEdit(PrimaryJournal, Me.UserCOde, True, True) = True Then
                    If Me.IsiKeJournal(Me.rowIndex, Urutan, PrimaryJournal, True, ClsTransaksiRecurring.EnumAction.Edit) = True Then
                        'MsgBox("Journal edited", MsgBoxStyle.Information, "Information")
                        ' Me.dgvEntries.Item(Me.ColTrans_Status.Index, Me.rowIndex).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Created)
                    Else
                        'MsgBox("Failed to edit journal.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tms_Skip_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tms_Skip.Click
        Dim DTAU As New DseTransaksiRecurringTableAdapters.tbltrans_recurring_journal_transaction_entriesTableAdapter
        Dim DTAUpdate As New DseTransaksiRecurringTableAdapters.Khusus_Update_Trans_EntriesTableAdapter
        Try

            If Not String.IsNullOrEmpty(Me.dgvEntries.Item(Me.ColTrans_JournalNumber.Index, Me.rowIndex).Value) Then
                DTAUpdate.UpdateQuery(Me.dgvEntries.Item(Me.ColTrans_PrimaryJournal.Index, Me.rowIndex).Value, _
                                      Me.dgvEntries.Item(Me.ColTrans_JournalNumber.Index, Me.rowIndex).Value, _
                                      Me.dgvEntries.Item(Me.ColTrans_JournalDate.Index, Me.rowIndex).Value, _
                                      Me.dgvEntries.Item(Me.ColTrans_Description.Index, Me.rowIndex).Value, _
                                      ClsTransaksiRecurring.EnumStatus.Skiped, _
                                      Me.dgvEntries.Item(Me.ColTrans_NoUrut.Index, Me.rowIndex).Value, _
                                      Me.dgvEntries.Item(Me.ColTrans_JournalNumber.Index, Me.rowIndex).Value)

                If Me.Limit > (Me.dgvEntries.Rows.Count) AndAlso Me.rowIndex = Me.dgvEntries.Rows.Count - 1 Then
                    Me.intBatasBaris = Me.dgvEntries.Rows.Count + 1
                    Me.TambahkanBarisEntries(Me.Limit)
                    Dim IntRowX As Integer = 0
                    IntRowX = CInt(Me.dgvEntries.Rows.Count - 1)
                    If String.IsNullOrEmpty(Me.dgvEntries.Item(Me.ColTrans_Remarks.Index, IntRowX).Value) Then
                        Me.dgvEntries.Item(Me.ColTrans_Remarks.Index, IntRowX).Value = ""
                    End If
                    DTAU.Insert(Me.TxtPrimaryJournal.Text, _
                                          Me.dgvEntries.Item(Me.ColTrans_PrimaryJournal.Index, IntRowX).Value, _
                                          Me.dgvEntries.Item(Me.ColTrans_JournalNumber.Index, IntRowX).Value, _
                                          Me.dgvEntries.Item(Me.ColTrans_JournalDate.Index, IntRowX).Value, _
                                          Me.dgvEntries.Item(Me.ColTrans_Description.Index, IntRowX).Value, _
                                          Me.dgvEntries.Item(Me.ColTrans_Remarks.Index, IntRowX).Value, _
                                          ClsTransaksiRecurring.EnumStatus.Planed, _
                                          Me.dgvEntries.Item(Me.ColTrans_NoUrut.Index, IntRowX).Value)
                End If
                Me.ObjRecurring.CariDetail(Me.TxtPrimaryJournal.Text, Me.dgvItem, Me)
                Me.SalinkeField()

                'Me.dgvEntries.Item(Me.ColTrans_Status.Index, Me.rowIndex).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Skiped)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAU.Dispose()
            DTAUpdate.Dispose()
        End Try
    End Sub

    Private Sub tms_Lock_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tms_Lock.Click
        Dim DTAUpdate As New DseTransaksiRecurringTableAdapters.Khusus_Update_Trans_EntriesTableAdapter

        Try
            If Not String.IsNullOrEmpty(Me.dgvEntries.Item(Me.ColTrans_JournalNumber.Index, Me.rowIndex).Value) Then
                DTAUpdate.UpdateQuery(Me.dgvEntries.Item(Me.ColTrans_PrimaryJournal.Index, Me.rowIndex).Value, _
                                                     Me.dgvEntries.Item(Me.ColTrans_JournalNumber.Index, Me.rowIndex).Value, _
                                                     Me.dgvEntries.Item(Me.ColTrans_JournalDate.Index, Me.rowIndex).Value, _
                                                     Me.dgvEntries.Item(Me.ColTrans_Description.Index, Me.rowIndex).Value, _
                                                     ClsTransaksiRecurring.EnumStatus.Locked, _
                                                     Me.dgvEntries.Item(Me.ColTrans_NoUrut.Index, Me.rowIndex).Value, _
                                                     Me.dgvEntries.Item(Me.ColTrans_JournalNumber.Index, Me.rowIndex).Value)
                Me.ObjRecurring.CariDetail(Me.TxtPrimaryJournal.Text, Me.dgvItem, Me)
                Me.SalinkeField()
                'Me.dgvEntries.Item(Me.ColTrans_Status.Index, Me.rowIndex).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Locked)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAUpdate.Dispose()
        End Try

    End Sub

    Private Sub tms_Update_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tms_Update.Click

        Try
            If Not String.IsNullOrEmpty(Me.dgvEntries.Item(Me.ColTrans_JournalNumber.Index, Me.rowIndex).Value) Then
                Me.dgvEntries.Item(Me.ColTrans_Status.Index, Me.rowIndex).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed)
                Me.TambahkanBarisEntries(Me.Limit)
                Dim Urutan As Integer = 0
                Urutan = Me.GetUrutanJournal(Me.rowIndex)
                Dim PrimaryJournal As Double = -1
                PrimaryJournal = Me.dgvEntries.Item(Me.ColTrans_PrimaryJournal.Index, Me.rowIndex).Value
                ' MessageBox.Show(Me.dgvEntries.Item(Me.ColTrans_Description.Index, Me.rowIndex).Value)
                If Me.IsiKeJournal(Me.rowIndex, Urutan, PrimaryJournal, True, ClsTransaksiRecurring.EnumAction.Update) = True Then
                    'MsgBox("Journal updated", MsgBoxStyle.Information, "Information")
                    ' Me.dgvEntries.Item(Me.ColTrans_Status.Index, Me.rowIndex).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Created)
                Else
                    'MsgBox("Failed to update journal.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                    'Me.dgvEntries.Item(Me.ColTrans_Status.Index, Me.rowIndex).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Created)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tms_UpdateAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tms_UpdateAll.Click
        Dim row As New DataGridViewRow
        Try
            If Me.dgvEntries.Rows.Count > 0 Then
                Dim Msg As Boolean = True
                For Each row In Me.dgvEntries.Rows
                    With row
                        If .Cells(Me.ColTrans_Status.Index).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Created) Then
                            .Cells(Me.ColTrans_Status.Index).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed)
                            Me.TambahkanBarisEntries(Me.Limit)
                            Dim Urutan As Integer = 0
                            Urutan = Me.GetUrutanJournal(row.Index)
                            Dim PrimaryJournal As Double = -1
                            PrimaryJournal = .Cells(Me.ColTrans_PrimaryJournal.Index).Value
                            '  .Cells(Me.ColTrans_Remarks.Index).Value = Urutan
                            If Me.IsiKeJournal(row.Index, Urutan, PrimaryJournal, True, ClsTransaksiRecurring.EnumAction.Update_All) = True Then
                                Msg = True
                            Else
                                Msg = False
                                MsgBox("Failed to update journal: " & .Cells(Me.ColTrans_JournalNumber.Index).Value & _
                                       ", update process will be stoped.", MsgBoxStyle.Information, "Information")
                                Exit Sub
                                '  .Cells(Me.ColTrans_Status.Index).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Created)
                            End If
                        End If
                    End With
                Next
                If Msg = True Then
                    MsgBox("Journal updated.", MsgBoxStyle.Information, "Information")
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub tms_DeleteAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tms_DeleteAll.Click

        Dim row As New DataGridViewRow
        Try
            If Me.dgvEntries.Rows.Count > 0 Then
                Dim msg As Boolean = True

                Dim intD As Integer = 0
                intD = Me.dgvEntries.Rows.Count - 1
                While intD >= 0
                    row = Me.dgvEntries.Rows(intD)
                    With row
                        If Not String.IsNullOrEmpty(.Cells(Me.ColTrans_JournalNumber.Index).Value) Then
                            If .Cells(Me.ColTrans_Status.Index).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Created) Then
                                Dim Urutan As Integer = 0
                                Urutan = Me.GetUrutanJournal(intD)
                                Dim PrimaryJournal As Double = -1
                                PrimaryJournal = .Cells(Me.ColTrans_PrimaryJournal.Index).Value
                                If Me.IsiKeJournal(intD, Urutan, PrimaryJournal, False, ClsTransaksiRecurring.EnumAction.Delete_All) = True Then
                                    msg = True
                                Else
                                    msg = False
                                    MsgBox("Failed to delete data journal: " & .Cells(Me.ColTrans_JournalNumber.Index).Value & _
                                       " update process will be stoped.", MsgBoxStyle.Information, "Information")
                                    Me.BerhasilDelete = False
                                    Exit Sub
                                End If
                            End If
                        End If
                    End With
                    intD -= 1
                End While

                'For intX As Integer = Me.dgvEntries.Rows.Count - 1 To 0
                '    row = Me.dgvEntries.Rows(intX)
                '    'For Each row In Me.dgvEntries.Rows
                '    With row
                '        If Not String.IsNullOrEmpty(.Cells(Me.ColTrans_JournalNumber.Index).Value) Then
                '            If .Cells(Me.ColTrans_Status.Index).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Created) Then
                '                Dim Urutan As Integer = 0
                '                Urutan = Me.GetUrutanJournal(row.Index)
                '                Dim PrimaryJournal As Double = -1
                '                PrimaryJournal = .Cells(Me.ColTrans_PrimaryJournal.Index).Value
                '                If Me.IsiKeJournal(intX, Urutan, PrimaryJournal, False, ClsTransaksiRecurring.EnumAction.Delete_All) = True Then
                '                    msg = True
                '                Else
                '                    msg = False
                '                    MsgBox("Failed to delete data journal transaction no " & .Cells(Me.ColTrans_JournalNumber.Index).Value & _
                '                       " process update will be stoped.", MsgBoxStyle.Information, "Information")
                '                    Me.BerhasilDelete = False
                '                    Exit Sub
                '                End If
                '            End If
                '        End If
                '    End With
                'Next

                '  Next
                If msg = True Then
                    MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                    Me.BerhasilDelete = True
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function IsiKeJournal(ByVal VrowIndex As Integer, ByVal vUrutanJournal As Integer, _
                                 ByVal vPrimaryJournal As Double, ByVal Show As Boolean, ByVal EnumAction As ClsTransaksiRecurring.EnumAction) As Boolean

        Dim I As New FrmTransaksiJournal
        'Dim ObjJurnal As New ClsTransaksiJournal
        'Dim ObjJurnalItem As New ClsTransaksiJournal_Item
        Dim rowDGV As New DataGridViewRow
        Dim IrowDGV As New DataGridViewRow
        'Dim UrutanJournal As Integer = 0
        Dim DTAUpdate As New DseTransaksiRecurringTableAdapters.Khusus_Update_Trans_EntriesTableAdapter
        Dim objDet As New ClsTransaksiRecurring_Item


        Try
            I.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
            I.UserCOde = Me.UserCOde
            If EnumAction = ClsTransaksiRecurring.EnumAction.Create Then
                I.Baru = True
                I.btnNew_Click(Nothing, Nothing)
            ElseIf EnumAction = ClsTransaksiRecurring.EnumAction.Update OrElse _
                   EnumAction = ClsTransaksiRecurring.EnumAction.Update_All OrElse _
                   EnumAction = ClsTransaksiRecurring.EnumAction.Delete OrElse _
                   EnumAction = ClsTransaksiRecurring.EnumAction.Delete_All OrElse _
                   EnumAction = ClsTransaksiRecurring.EnumAction.Edit Then
                I.Baru = False
                If I.ObjJurnal.CariDetail(vPrimaryJournal, I.dgvItem, I, I.TipeJournal) = True Then
                    I.SalinkeField()
                Else
                    Me.ObjRecurring.CariDetail(Me.TxtPrimaryJournal.Text, Me.dgvItem, Me)
                    Me.SalinkeField()
                    Return False
                End If
            End If

            ' MessageBox.Show(Me.dgvEntries.Item(Me.ColTrans_Description.Index, Me.rowIndex).Value)



            If EnumAction <> ClsTransaksiRecurring.EnumAction.Edit Then
                Dim Rounding As Integer = 0
                Rounding = Me.DefaultPrecision
                I.TxtPrimaryJournal.Text = vPrimaryJournal
                I.TxtRefNo.Text = Me.dgvEntries.Item(Me.ColTrans_JournalNumber.Index, VrowIndex).Value
                I.dtpDate.Value = Format(CDate(Me.dgvEntries.Item(Me.ColTrans_JournalDate.Index, VrowIndex).Value), "dd MMM yyyy")
                I.TxtRemarks.Text = Me.dgvEntries.Item(Me.ColTrans_Description.Index, VrowIndex).Value
                I.TxtUnbalance.Text = CDec(Me.TxtUnbalance.Text)
                '  I.txtUnbalanceIncrement.Text = CDec(Me.txtUnbalanceIncrement.Text)
                'UrutanJournal = CInt(VrowIndex + 1)
                I.dgvItem.Rows.Clear()
                I.dgvItem.AllowUserToAddRows = True
                For Each rowDGV In Me.dgvItem.Rows
                    With rowDGV
                        If Not String.IsNullOrEmpty(.Cells(Me.ColAccount.Index).Value) Then
                            I.dgvItem.Rows.Add()
                            IrowDGV = I.dgvItem.Rows(I.dgvItem.Rows.Count - 2)
                            IrowDGV.Cells(I.ColPrimaryAkun.Index).Value = .Cells(Me.ColPrimaryAkun.Index).Value
                            IrowDGV.Cells(I.ColAccount.Index).Value = .Cells(Me.ColAccount.Index).Value
                            IrowDGV.Cells(I.ColDescAccount.Index).Value = .Cells(Me.ColDescAccount.Index).Value
                            IrowDGV.Cells(I.ColExchangeRate.Index).Value = .Cells(Me.ColExchangeRate.Index).Value
                            IrowDGV.Cells(I.ColRemarks.Index).Value = .Cells(Me.ColRemarks.Index).Value
                            IrowDGV.Cells(I.ColPrimaryDept.Index).Value = .Cells(Me.ColPrimaryDept.Index).Value
                            IrowDGV.Cells(I.ColDept.Index).Value = .Cells(Me.ColDept.Index).Value
                            IrowDGV.Cells(I.ColPrimaryJob.Index).Value = .Cells(Me.ColPrimaryJob.Index).Value
                            IrowDGV.Cells(I.ColJob.Index).Value = .Cells(Me.ColJob.Index).Value
                            IrowDGV.Cells(I.ColPrimaryCurrency.Index).Value = .Cells(Me.ColPrimaryCurrency.Index).Value
                            IrowDGV.Cells(I.ColOriginalCurrency.Index).Value = .Cells(Me.ColOriginalCurrency.Index).Value
                            If .Cells(Me.ColAcc_IncrementStyle.Index).Value = Me.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.No_Increment) Then
                                If IsNumeric(.Cells(Me.ColDebet.Index).Value) AndAlso CDec(.Cells(Me.ColDebet.Index).Value) > 0 Then
                                    IrowDGV.Cells(I.ColDebet.Index).Value = ObjCD.DigitGrouping(.Cells(Me.ColDebet.Index).Value, Rounding)
                                Else
                                    IrowDGV.Cells(I.ColDebet.Index).Value = 0
                                End If
                                If IsNumeric(.Cells(Me.ColKredit.Index).Value) AndAlso CDec(.Cells(Me.ColKredit.Index).Value) > 0 Then
                                    IrowDGV.Cells(I.ColKredit.Index).Value = ObjCD.DigitGrouping(.Cells(Me.ColKredit.Index).Value, Rounding)
                                Else
                                    IrowDGV.Cells(I.ColKredit.Index).Value = 0
                                End If
                            ElseIf .Cells(Me.ColAcc_IncrementStyle.Index).Value = Me.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.Fixed_Increment) Then
                                If IsNumeric(.Cells(Me.ColDebet.Index).Value) AndAlso CDec(.Cells(Me.ColDebet.Index).Value) > 0 Then
                                    IrowDGV.Cells(I.ColDebet.Index).Value = ObjCD.DigitGrouping(CDec(.Cells(Me.ColDebet.Index).Value) + _
                                                                            (vUrutanJournal * CDec(.Cells(Me.ColAcc_IncrementValue.Index).Value)), Rounding)
                                Else
                                    IrowDGV.Cells(I.ColDebet.Index).Value = 0
                                End If
                                If IsNumeric(.Cells(Me.ColKredit.Index).Value) AndAlso CDec(.Cells(Me.ColKredit.Index).Value) > 0 Then
                                    IrowDGV.Cells(I.ColKredit.Index).Value = ObjCD.DigitGrouping(CDec(.Cells(Me.ColKredit.Index).Value) + _
                                                                            (vUrutanJournal * CDec(.Cells(Me.ColAcc_IncrementValue.Index).Value)), Rounding)
                                Else
                                    IrowDGV.Cells(I.ColKredit.Index).Value = 0
                                End If
                            ElseIf .Cells(Me.ColAcc_IncrementStyle.Index).Value = Me.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.Percentage_Increment) Then
                                Dim Persen As Decimal = 0
                                Persen = ((CDec(100 + CDec(.Cells(Me.ColAcc_IncrementValue.Index).Value)) / 100) ^ vUrutanJournal)

                                If IsNumeric(.Cells(Me.ColDebet.Index).Value) AndAlso CDec(.Cells(Me.ColDebet.Index).Value) > 0 Then
                                    IrowDGV.Cells(I.ColDebet.Index).Value = ObjCD.DigitGrouping(CDec(.Cells(Me.ColDebet.Index).Value) * Persen, Rounding)
                                Else
                                    IrowDGV.Cells(I.ColDebet.Index).Value = 0
                                End If
                                If IsNumeric(.Cells(Me.ColKredit.Index).Value) AndAlso CDec(.Cells(Me.ColKredit.Index).Value) > 0 Then
                                    IrowDGV.Cells(I.ColKredit.Index).Value = ObjCD.DigitGrouping(CDec(.Cells(Me.ColKredit.Index).Value) * Persen, Rounding)
                                Else
                                    IrowDGV.Cells(I.ColKredit.Index).Value = 0
                                End If
                            End If

                            I.HitungOriginalValues(IrowDGV.Index)

                        End If
                    End With
                Next

                I.SetMulticurrency()
            End If

            I.Recurring = True
            Dim Status As Integer = 0
            If EnumAction = ClsTransaksiRecurring.EnumAction.Create Then
                Status = ClsTransaksiRecurring.EnumStatus.Created
                I.Recurring = True
                If Show = False Then
                    I.HideMsg = True
                    I.btnOK_Click(Nothing, Nothing)
                End If
            ElseIf EnumAction = ClsTransaksiRecurring.EnumAction.Update Then
                I.Recurring = True
                I.btnEdit_Click(Nothing, Nothing)
                Status = ClsTransaksiRecurring.EnumStatus.Created
            ElseIf EnumAction = ClsTransaksiRecurring.EnumAction.Delete Then
                I.btnDelete_Click(Nothing, Nothing)
                If VrowIndex = Me.dgvEntries.RowCount - 1 Then
                    Status = ClsTransaksiRecurring.EnumStatus.Planed
                Else
                    If (VrowIndex + 1) <= Me.dgvEntries.RowCount - 1 AndAlso _
                       Me.dgvEntries.Item(Me.ColTrans_Status.Index, VrowIndex + 1).Value = _
                       Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed) Then
                        Status = ClsTransaksiRecurring.EnumStatus.Planed
                    Else
                        Status = ClsTransaksiRecurring.EnumStatus.Skiped
                    End If
                End If

            ElseIf EnumAction = ClsTransaksiRecurring.EnumAction.Edit Then
                Status = ClsTransaksiRecurring.EnumStatus.Created
                I.Recurring = False
                I.btnEdit_Click(Nothing, Nothing)

            ElseIf EnumAction = ClsTransaksiRecurring.EnumAction.Delete_All Then
                I.btnDelete_Click(Nothing, Nothing)
                If VrowIndex = Me.dgvEntries.RowCount - 1 Then
                    Status = ClsTransaksiRecurring.EnumStatus.Planed
                Else
                    If (VrowIndex + 1) <= Me.dgvEntries.RowCount - 1 AndAlso _
                       Me.dgvEntries.Item(Me.ColTrans_Status.Index, VrowIndex + 1).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed) Then
                        Status = ClsTransaksiRecurring.EnumStatus.Planed
                    Else
                        Status = ClsTransaksiRecurring.EnumStatus.Skiped
                    End If
                End If
            ElseIf EnumAction = ClsTransaksiRecurring.EnumAction.Update_All Then
                I.btnEdit_Click(Nothing, Nothing)
                I.btnOK_Click(Nothing, Nothing)
                Status = ClsTransaksiRecurring.EnumStatus.Created
            End If


            If Show = True Then
                I.Recurring = True
                I.TxtRefNo.ReadOnly = True
                I.UserCOde = Me.UserCOde
                I.BackColor = Me.BackColor
                I.BackcolorAll = Me.BackcolorAll
                I.ShowDialog()
            End If


            If I.Berhasil = True Then
                Me.AdaJournalBaru = True
                'Update dulu primary journal
                Dim Primary As Double = -1
                Dim NoUrut As Double = 0
                Dim NoJournal As String = ""
                Dim dtDate As Date = Date.Now
                Dim Remarks As String = ""

                If EnumAction = ClsTransaksiRecurring.EnumAction.Create OrElse EnumAction = ClsTransaksiRecurring.EnumAction.Edit Then
                    Primary = I.ObjJurnal.GetPrimary(Me.dgvEntries.Item(Me.ColTrans_JournalNumber.Index, VrowIndex).Value)
                    NoJournal = I.TxtRefNo.Text
                    dtDate = I.dtpDate.Value
                    Remarks = I.TxtRemarks.Text
                    NoUrut = Me.dgvEntries.Item(Me.ColTrans_NoUrut.Index, VrowIndex).Value
                Else
                    Primary = I.ObjJurnal.GetPrimary(Me.dgvEntries.Item(Me.ColTrans_JournalNumber.Index, VrowIndex).Value)
                    NoJournal = Me.dgvEntries.Item(Me.ColTrans_JournalNumber.Index, VrowIndex).Value
                    dtDate = Me.dgvEntries.Item(Me.ColTrans_JournalDate.Index, VrowIndex).Value
                    Remarks = Me.dgvEntries.Item(Me.ColTrans_Description.Index, VrowIndex).Value
                    NoUrut = Me.dgvEntries.Item(Me.ColTrans_NoUrut.Index, VrowIndex).Value
                End If

                If objDet.UpdateStatus(Primary, NoJournal, dtDate, Remarks, Status, NoUrut, Me.UserCOde) = True Then
                    Me.ObjRecurring.CariDetail(Me.TxtPrimaryJournal.Text, Me.dgvItem, Me)
                    Me.SalinkeField()
                End If
                'If DTAUpdate.UpdateQuery(Primary, NoJournal, dtDate, Remarks, Status, NoUrut, NoJournal) > 0 Then
                '    Me.ObjRecurring.CariDetail(Me.TxtPrimaryJournal.Text, Me.dgvItem, Me)
                '    Me.SalinkeField()
                'End If
                Return True
            Else
                Me.ObjRecurring.CariDetail(Me.TxtPrimaryJournal.Text, Me.dgvItem, Me)
                Me.SalinkeField()
                Return False
            End If


            'I.ShowDialog()
            ' I.btnOK.PerformClick()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            I.Dispose()
            'ObjJurnal = Nothing
            'ObjJurnalItem = Nothing
            DTAUpdate.Dispose()
        End Try
    End Function
  

    'Public Function Update_tms(ByVal VrowIndex As Integer, ByVal vUrutanJournal As Integer) As Boolean
    '    Dim I As New FrmTransaksiJournal
    '    Dim ObjJurnal As New ClsTransaksiJournal
    '    Dim ObjJurnalItem As New ClsTransaksiJournal_Item
    '    Dim rowDGV As New DataGridViewRow
    '    Dim IrowDGV As New DataGridViewRow
    '    Dim UrutanJournal As Integer = 0

    '    Try

    '        I.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
    '        I.btnNew_Click(Nothing, Nothing)
    '        I.TxtPrimaryJournal.Text = Me.dgvEntries.Item(Me.ColTrans_PrimaryJournal.Index, VrowIndex).Value
    '        I.TxtRefNo.Text = Me.dgvEntries.Item(Me.ColTrans_JournalNumber.Index, VrowIndex).Value
    '        I.dtpDate.Value = Format(CDate(Me.dgvEntries.Item(Me.ColTrans_JournalDate.Index, VrowIndex).Value), "dd MMM yyyy")
    '        I.TxtRemarks.Text = Me.dgvEntries.Item(Me.ColTrans_Description.Index, VrowIndex).Value
    '        I.TxtUnbalance.Text = CDec(Me.TxtUnbalance.Text)
    '        UrutanJournal = CInt(VrowIndex + 1)

    '        I.dgvItem.Rows.Clear()
    '        I.dgvItem.AllowUserToAddRows = True
    '        For Each rowDGV In Me.dgvItem.Rows
    '            With rowDGV
    '                If Not String.IsNullOrEmpty(.Cells(Me.ColAccount.Index).Value) Then
    '                    I.dgvItem.Rows.Add()
    '                    IrowDGV = I.dgvItem.Rows(I.dgvItem.Rows.Count - 2)
    '                    IrowDGV.Cells(I.ColPrimaryAkun.Index).Value = .Cells(Me.ColPrimaryAkun.Index).Value
    '                    IrowDGV.Cells(I.ColAccount.Index).Value = .Cells(Me.ColAccount.Index).Value
    '                    IrowDGV.Cells(I.ColDescAccount.Index).Value = .Cells(Me.ColDescAccount.Index).Value
    '                    If .Cells(Me.ColAcc_IncrementStyle.Index).Value = Me.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.No_Increment) Then
    '                        If IsNumeric(.Cells(Me.ColDebet.Index).Value) AndAlso CDec(.Cells(Me.ColDebet.Index).Value) > 0 Then
    '                            IrowDGV.Cells(I.ColDebet.Index).Value = .Cells(Me.ColDebet.Index).Value
    '                        Else
    '                            IrowDGV.Cells(I.ColDebet.Index).Value = 0
    '                        End If
    '                        If IsNumeric(.Cells(Me.ColKredit.Index).Value) AndAlso CDec(.Cells(Me.ColKredit.Index).Value) > 0 Then
    '                            IrowDGV.Cells(I.ColKredit.Index).Value = .Cells(Me.ColKredit.Index).Value
    '                        Else
    '                            IrowDGV.Cells(I.ColKredit.Index).Value = 0
    '                        End If
    '                    ElseIf .Cells(Me.ColAcc_IncrementStyle.Index).Value = Me.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.Fixed_Increment) Then
    '                        If IsNumeric(.Cells(Me.ColDebet.Index).Value) AndAlso CDec(.Cells(Me.ColDebet.Index).Value) > 0 Then
    '                            IrowDGV.Cells(I.ColDebet.Index).Value = CDec(.Cells(Me.ColDebet.Index).Value) + (vUrutanJournal * CDec(.Cells(Me.ColAcc_IncrementValue.Index).Value))
    '                        Else
    '                            IrowDGV.Cells(I.ColDebet.Index).Value = 0
    '                        End If
    '                        If IsNumeric(.Cells(Me.ColKredit.Index).Value) AndAlso CDec(.Cells(Me.ColKredit.Index).Value) > 0 Then
    '                            IrowDGV.Cells(I.ColKredit.Index).Value = CDec(.Cells(Me.ColKredit.Index).Value) + (vUrutanJournal * CDec(.Cells(Me.ColAcc_IncrementValue.Index).Value))
    '                        Else
    '                            IrowDGV.Cells(I.ColKredit.Index).Value = 0
    '                        End If
    '                    ElseIf .Cells(Me.ColAcc_IncrementStyle.Index).Value = Me.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.Percentage_Increment) Then
    '                        Dim Persen As Decimal = 0
    '                        Persen = CDec(100 + CDec(.Cells(Me.ColAcc_IncrementValue.Index).Value)) / 100

    '                        If IsNumeric(.Cells(Me.ColDebet.Index).Value) AndAlso CDec(.Cells(Me.ColDebet.Index).Value) > 0 Then
    '                            IrowDGV.Cells(I.ColDebet.Index).Value = CDec(.Cells(Me.ColDebet.Index).Value) * (Persen ^ vUrutanJournal)
    '                        Else
    '                            IrowDGV.Cells(I.ColDebet.Index).Value = 0
    '                        End If
    '                        If IsNumeric(.Cells(Me.ColKredit.Index).Value) AndAlso CDec(.Cells(Me.ColKredit.Index).Value) > 0 Then
    '                            IrowDGV.Cells(I.ColKredit.Index).Value = CDec(.Cells(Me.ColKredit.Index).Value) * (Persen ^ vUrutanJournal)
    '                        Else
    '                            IrowDGV.Cells(I.ColKredit.Index).Value = 0
    '                        End If
    '                    End If
    '                    IrowDGV.Cells(I.ColExchangeRate.Index).Value = .Cells(Me.ColExchangeRate.Index).Value
    '                    IrowDGV.Cells(I.ColDebet_OriginalCurr.Index).Value = .Cells(Me.ColDebet_OriginalCurr.Index).Value
    '                    IrowDGV.Cells(I.ColCredit_OriginalCurrency.Index).Value = .Cells(Me.ColCredit_OriginalCurrency.Index).Value
    '                    IrowDGV.Cells(I.ColRemarks.Index).Value = .Cells(Me.ColRemarks.Index).Value
    '                    IrowDGV.Cells(I.ColPrimaryDept.Index).Value = .Cells(Me.ColPrimaryDept.Index).Value
    '                    IrowDGV.Cells(I.ColDept.Index).Value = .Cells(Me.ColDept.Index).Value
    '                    IrowDGV.Cells(I.ColPrimaryJob.Index).Value = .Cells(Me.ColPrimaryJob.Index).Value
    '                    IrowDGV.Cells(I.ColJob.Index).Value = .Cells(Me.ColJob.Index).Value
    '                    IrowDGV.Cells(I.ColPrimaryCurrency.Index).Value = .Cells(Me.ColPrimaryCurrency.Index).Value
    '                    IrowDGV.Cells(I.ColOriginalCurrency.Index).Value = .Cells(Me.ColOriginalCurrency.Index).Value
    '                End If
    '            End With
    '        Next
    '        I.SetMulticurrency()
    '        I.Recurring = True
    '        I.btnOK_Click(Nothing, Nothing)
    '        If I.Berhasil Then
    '            Return True
    '        Else
    '            Return False
    '        End If
    '        'I.ShowDialog()
    '        ' I.btnOK.PerformClick()
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        I.Dispose()
    '        ObjJurnal = Nothing
    '        ObjJurnalItem = Nothing
    '    End Try
    'End Function

    Public Function ApakahDataBenarJournal(ByVal VBalance As Decimal, ByVal DGV As DataGridView) As Boolean
        Try
            If DGV.Rows.Count > 0 AndAlso String.IsNullOrEmpty(DGV.Item(Me.ColAccount.Index, 0).Value) Then
                MsgBox("Can't create journal transaction, you must enter at least 1 balance journal in 'Account Entries'.", MsgBoxStyle.Information, "Information")
                Return False
                Exit Function
            End If

            If DGV.Rows.Count > 0 AndAlso String.IsNullOrEmpty(Me.TxtRemarks.Text) Then
                MsgBox("Can't create journal, you must entry description. ", MsgBoxStyle.Information, "Information")
                Return False
                Exit Function
            End If

            If VBalance <> 0 Then
                MsgBox("Can't create journal, make sure your recurring journal in 'Account Entries' is balanced. ", MsgBoxStyle.Information, "Information")
                Return False
                Exit Function
            End If

            'tidak boleh bikin transaksi di belakang convertion date
            Dim dtDate As Date = Date.Now
            If IsDate(Me.dgvEntries.Item(Me.ColTrans_JournalDate.Index, Me.rowIndex).Value) Then
                dtDate = Me.dgvEntries.Item(Me.ColTrans_JournalDate.Index, Me.rowIndex).Value
            End If

            If Me.ObjClosing.GetConvertionDate.Date > dtDate.Date Then
                MsgBox("Minimum transaction date is " & Format(Me.ObjClosing.GetConvertionDate.Date, "dd MMM yyyy") & ".", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If


            'cek data yang salah, double dalam kode
            Dim strNo As String = ""
            Dim PrimaryJournal As Double = -1
            If Not String.IsNullOrEmpty(Me.dgvEntries.Item(Me.ColTrans_JournalNumber.Index, Me.rowIndex).Value) Then
                strNo = Me.dgvEntries.Item(Me.ColTrans_JournalNumber.Index, Me.rowIndex).Value
                PrimaryJournal = Me.dgvEntries.Item(Me.ColTrans_PrimaryJournal.Index, Me.rowIndex).Value
            End If

            If PrimaryJournal = -1 Then
                If Me.ObjJournal.ApakahSudahAda(strNo) = True Then
                    MsgBox("Journal number duplication on: " & strNo & ", please change your journal number", _
                    MsgBoxStyle.Information, "Information")
                    Me.dgvEntries.Focus()
                    Return False
                End If
            Else
                Dim TipeCodeLama As String = Me.ObjJournal.GetRefCode(PrimaryJournal)
                If TipeCodeLama <> strNo Then
                    If Me.ObjJournal.ApakahSudahAda(strNo) = True Then
                        MsgBox("Journal number duplication on: " & strNo & ", please change your journal number", _
                        MsgBoxStyle.Information, "Information")
                        Me.dgvEntries.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.ObjJournal.GetTimestamp(PrimaryJournal)
                If Me.ObjJournal.Timestamp <> OldTimeStamp Then
                    'MsgBox(Me.ObjJournal.Timestamp & "," & OldTimeStamp)
                    'If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    'MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    'Me.ObjJournal.CariDetail(Me.ObjJournal.Primary, Me.dgvItem, Me, Me.TipeJournal)
                    'Me.SalinkeField()
                    'End If
                    MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm")
                    Return False
                End If
            End If


            'jika sudah ada closing jurnal untuk tahun ybs maka tidak bisa input data baru
            If Me.ObjJournal.GetJml(ClsTransaksiJournal.EnumTipeJournal.ClosingJournal, Me.ObjCD.GetFisrtDateOfTheYear(dtDate), _
            Me.ObjCD.GetLastDateOfTheYear(dtDate)) > 0 Then
                MsgBox("There are closing jurnal for this year you can't save this data", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Private Function ApakahBolehUpdate(ByVal VBalance As Decimal, ByVal DGV As DataGridView) As Boolean
    '    Try
    '        If DGV.Rows.Count > 0 AndAlso String.IsNullOrEmpty(DGV.Item(Me.ColAccount.Index, 0).Value) Then
    '            MsgBox("Can not create journal, you must entry minimum 1 balance journal in 'Account Entries'. ", MsgBoxStyle.Information, "Information")
    '            Return False
    '            Exit Function
    '        End If

    '        If DGV.Rows.Count > 0 AndAlso String.IsNullOrEmpty(Me.TxtRemarks.Text) Then
    '            MsgBox("Can not create journal, you must entry description. ", MsgBoxStyle.Information, "Information")
    '            Return False
    '            Exit Function
    '        End If

    '        If VBalance <> 0 Then
    '            MsgBox("Can not create journal, make sure your recurring journal in 'Account Entries' is balance. ", MsgBoxStyle.Information, "Information")
    '            Return False
    '            Exit Function
    '        End If

    '        Return True

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    Public Function GetUrutanJournal(ByVal VRowIndex As Integer) As Integer
        Dim rowDgv As New DataGridViewRow
        Try
            Dim Urutan As Integer = -1
            If Me.dgvEntries.Rows.Count > 0 Then
                For intA As Integer = 0 To Me.dgvEntries.Rows.Count - 1
                    Dim intB As Integer = 0
                    rowDgv = Me.dgvEntries.Rows(intA)
                    With rowDgv
                        If Me.chkIgnoreIncrementOnSkip.Checked = False Then
                            Urutan = intA
                        Else
                            If .Cells(Me.ColTrans_Status.Index).Value <> Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Skiped) Then
                                Urutan += 1
                            End If
                        End If
                        If intA = VRowIndex Then
                            Exit For
                        End If
                    End With
                Next
            End If
            Return Urutan

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Dim BerhasilSave As Boolean = False
        Try
            Me.dgvItem.AllowUserToAddRows = False
            Me.dgvReminder.AllowUserToAddRows = False
            Me.SalinkeObjek()
            If Baru = True Then
                If ObjRecurring.ApakahDataBenar1(Baru) = True Then
                    'konfirmasi disimpan atau tidak
                    If Me.ApakahDataBenar() = True Then
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")

                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjRecurring.SimpanBaru(Me.dgvItem, Me.UserCOde, Me) Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                BerhasilSave = True
                                Me.SalinkeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            End If
                            Me.dgvItem.AllowUserToAddRows = False
                            Me.dgvReminder.AllowUserToAddRows = False
                        ElseIf hasil = MsgBoxResult.No Then
                            BerhasilSave = False
                            Me.dgvItem.AllowUserToAddRows = True
                            Me.dgvReminder.AllowUserToAddRows = True
                        End If
                    Else
                        BerhasilSave = False
                        Me.dgvItem.AllowUserToAddRows = True
                        Me.dgvReminder.AllowUserToAddRows = True
                    End If
                Else
                    BerhasilSave = False
                    Me.dgvItem.AllowUserToAddRows = True
                    Me.dgvReminder.AllowUserToAddRows = True
                End If
            Else
                'konfirmasi disimpan atau tidak
                If Me.ApakahDataBenar = True Then
                    If ObjRecurring.ApakahBolehEdit(Me.TxtPrimaryJournal.Text, UserCOde, True, True) = True Then
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjRecurring.SimpanPerubahan(Me.dgvItem, Me.UserCOde, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                BerhasilSave = True
                                Me.SalinkeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            End If
                            Me.dgvItem.AllowUserToAddRows = False
                            Me.dgvReminder.AllowUserToAddRows = False
                        ElseIf hasil = MsgBoxResult.No Then
                            Me.dgvItem.AllowUserToAddRows = True
                            Me.dgvReminder.AllowUserToAddRows = True
                            BerhasilSave = False
                        End If
                    End If
                Else
                    BerhasilSave = False
                    Me.dgvItem.AllowUserToAddRows = True
                    Me.dgvReminder.AllowUserToAddRows = True
                End If
            End If

            If Me.FromReminder = True AndAlso BerhasilSave = True Then
                Me.Close()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub SalinkeObjek()
        Try
            With ObjRecurring
                .RefNumber = Me.TxtRefNo.Text
                .Description = Me.TxtRemarks.Text
                .Tanggal = Me.dtpDate.Value
                If Me.chkIgnoreIncrementOnSkip.Checked = True Then
                    .IgnoreIncrementOnSkip = 1
                Else
                    .IgnoreIncrementOnSkip = 0
                End If
                .Frequency = Me.cboFrequency.SelectedIndex
                If .Frequency = ClsTransaksiRecurring.EnumFrequency.Daily OrElse _
                   .Frequency = ClsTransaksiRecurring.EnumFrequency.Yearly Then
                    .Every = 0
                ElseIf .Frequency = ClsTransaksiRecurring.EnumFrequency.Weekly OrElse _
                       .Frequency = ClsTransaksiRecurring.EnumFrequency.Bi_Weekly Then
                    .Every = Me.cboEveryWeek.SelectedIndex
                ElseIf .Frequency = ClsTransaksiRecurring.EnumFrequency.Monthly OrElse _
                       .Frequency = ClsTransaksiRecurring.EnumFrequency.Quarter_Yearly OrElse _
                       .Frequency = ClsTransaksiRecurring.EnumFrequency.Semester Then
                    .Every = Me.cboEveryMonth.SelectedIndex
                End If

                If Me.chkActive.Checked = True Then
                    .Active = 1
                Else
                    .Active = 0
                End If
                If Me.chkLimit.Checked = True Then
                    .Limit = 1
                Else
                    .Limit = 0
                End If
                .LimitValue = Me.Limit
                'If IsNumeric(Me.txtLimit.Text) Then
                '    .LimitValue = Me.Limit
                'Else
                '    .LimitValue = Me.Limit
                'End If
                .StartDate = Me.dtpStartDate.Value
                .EndDate = Me.dtpEndDate.Value
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinkeField()
        Try
            With ObjRecurring
                cboDGV.Items.Clear()
                Me.cboDGV.Items.Add("No Increment")
                Me.cboDGV.Items.Add("Fixed Increment")
                Me.cboDGV.Items.Add("Percentage Increment")

                cboStatus.Items.Clear()
                Me.cboStatus.Items.Add("Planed")
                Me.cboStatus.Items.Add("Created")
                Me.cboStatus.Items.Add("Skiped")
                Me.cboStatus.Items.Add("Lock")

                Me.AdaJournalBaru = False
                Me.TxtPrimaryJournal.Text = .Primary
                Me.TxtRefNo.Text = .RefNumber
                Me.TxtRemarks.Text = .Description
                Me.dtpDate.Value = .Tanggal
                If .IgnoreIncrementOnSkip = 1 Then
                    Me.chkIgnoreIncrementOnSkip.Checked = True
                Else
                    Me.chkIgnoreIncrementOnSkip.Checked = False
                End If

                If .Active = 1 Then
                    Me.chkActive.Checked = True
                Else
                    Me.chkActive.Checked = False
                End If
                Me.cboFrequency.SelectedIndex = .Frequency
                'Me.cboFrequency_SelectedIndexChanged(Nothing, Nothing)

              
                If .Limit = 1 Then
                    Me.chkLimit.Checked = True
                    Me.txtLimit.Text = .LimitValue
                    Me.Limit = .LimitValue
                    Me.dtpStartDate.Value = .StartDate
                    Me.dtpEndDate.Value = .EndDate
                    ' Me.dtpEndDate.Enabled = False
                Else
                    Me.dtpStartDate.Value = .StartDate
                    Me.dtpEndDate.Value = .EndDate
                    Me.chkLimit.Checked = False
                    Me.txtLimit.Text = 0
                    Me.Limit = Me.GetLimit(Me.dtpStartDate.Value, Me.dtpEndDate.Value)
                    ' Me.dtpEndDate.Enabled = True
                End If


                'Me.chkLimit_CheckedChanged(Nothing, Nothing)
                If Me.dgvEntries.Rows.Count > intBatasBaris Then
                    Me.intBatasBaris = CInt(Me.dgvEntries.Rows.Count)
                Else
                    Me.intBatasBaris = 10
                End If


                If .Frequency = ClsTransaksiRecurring.EnumFrequency.Daily OrElse _
                  .Frequency = ClsTransaksiRecurring.EnumFrequency.Yearly Then
                    Me.cboEveryWeek.SelectedIndex = 0
                    Me.cboEveryMonth.SelectedIndex = 0
                    'Me.cboEveryWeek.Enabled = False
                    'Me.cboEveryMonth.Enabled = False
                ElseIf .Frequency = ClsTransaksiRecurring.EnumFrequency.Weekly OrElse _
                        .Frequency = ClsTransaksiRecurring.EnumFrequency.Bi_Weekly Then
                    Me.cboEveryWeek.SelectedIndex = .Every
                    Me.cboEveryMonth.SelectedIndex = 0
                    'Me.cboEveryWeek.Enabled = True
                    'Me.cboEveryMonth.Enabled = False
                    'Me.cboEveryWeek_SelectedIndexChanged(Nothing, Nothing)
                ElseIf .Frequency = ClsTransaksiRecurring.EnumFrequency.Monthly OrElse _
                       .Frequency = ClsTransaksiRecurring.EnumFrequency.Quarter_Yearly OrElse _
                       .Frequency = ClsTransaksiRecurring.EnumFrequency.Semester Then
                    Me.cboEveryWeek.SelectedIndex = 0
                    Me.cboEveryMonth.SelectedIndex = .Every
                    'Me.cboEveryWeek.Enabled = False
                    'Me.cboEveryMonth.Enabled = True
                    'Me.cboEveryMonth_SelectedIndexChanged(Nothing, Nothing)
                End If

                Me.BerhasilDelete = False
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

            If Me.ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data default precision harus non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            If Me.chkLimit.Checked = True Then
                If Not IsNumeric(Me.txtLimit.Text) OrElse Me.txtLimit.Text < 0 Then
                    MsgBox("Please use positive number in limit data.", MsgBoxStyle.Information, "Information")
                    Me.txtLimit.Focus()
                    Return False
                End If
            End If

            'Cek DGVItem
            If Me.dgvItem.Rows.Count < 2 Then
                MsgBox("Please specify at least 2 different accounts in account entries data.", MsgBoxStyle.Information, "Information")
                Me.TabControl1.SelectedIndex = 0
                Me.dgvItem.Focus()
                Return False
            End If

            'Cek DgvEtries
            If Me.dgvEntries.Rows.Count < 1 Then
                MsgBox("Please specify at least 1 item in transaction entries data.", MsgBoxStyle.Information, "Information")
                Me.dgvEntries.Focus()
                Me.TabControl1.SelectedIndex = 1
                Return False
            End If

            'Cek DgvReminder
            If Me.dgvReminder.Rows.Count < 1 Then
                MsgBox("Please specify at least 1 item in reminder entries data.", MsgBoxStyle.Information, "Information")
                Me.TabControl1.SelectedIndex = 2
                Me.dgvReminder.Focus()
                Return False
            End If



            ' Cek debet kredit harus balance
            Me.HitungUnbalance()
            If CDec(Me.TxtUnbalance.Text) <> 0 Then
                MsgBox("This is an unbalance journal, please make sure debit and credit are always balanced.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            Dim Unbalance As Decimal = 0
            Unbalance = Me.HitungUnbalanceIncrement
            If Unbalance <> 0 Then
                MsgBox("This journal will not be balanced in the next following journal entries, please check your data.", _
                         MsgBoxStyle.Information, "Information")
                Return False
            End If

            'tidak boleh bikin transaksi di belakang convertion date
            If Me.ObjClosing.GetConvertionDate.Date > Me.dtpStartDate.Value.Date Then
                MsgBox("Minimum transaction start date is " + Format(Me.ObjClosing.GetConvertionDate.Date, "dd MMM yyyy") + ".", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek apakah ada account number yang debet dan kreditnya bernilai 0
            For Each row In Me.dgvItem.Rows
                If Me.dgvItem.Item(Me.ColDebet.Index, row.Index).Value = 0 And _
                Me.dgvItem.Item(Me.ColKredit.Index, row.Index).Value = 0 Then
                    MsgBox("Please add positive number in account: " + Me.dgvItem.Item(Me.ColAccount.Index, row.Index).Value, _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If

                If Me.dgvItem.Item(Me.ColAcc_IncrementStyle.Index, row.Index).Value <> _
                   Me.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.No_Increment) AndAlso _
                   Me.dgvItem.Item(Me.ColAcc_IncrementValue.Index, row.Index).Value = 0 Then
                    MsgBox("Please add positive number in increment value if you choose increment style: " & _
                    Me.dgvItem.Item(Me.ColAcc_IncrementStyle.Index, row.Index).Value, _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Next

            If Me.dgvReminder.Rows.Count > 0 Then
                For Each row In Me.dgvReminder.Rows
                    If row.Index <> Me.dgvReminder.Rows.Count Then
                        If String.IsNullOrEmpty(Me.dgvReminder.Item(Me.ColReminder_UserCode.Index, row.Index).Value) Then
                            MsgBox("User code is null please input item reminder entries again", MsgBoxStyle.Information, "Information")
                            Me.dgvReminder.Focus()
                            Return False
                            Exit For
                        End If

                        'If String.IsNullOrEmpty(Me.dgvReminder.Item(Me.ColReminder_Remarks.Index, row.Index).Value) Then
                        '    MsgBox("remarks of reminder is null please input item reminder entries again", MsgBoxStyle.Information, "Information")
                        '    Me.dgvReminder.Focus()
                        '    Return False
                        '    Exit For
                        'End If

                    End If
                Next
            End If


            ''cek data yang salah, double dalam kode
            '' If Me.Baru = True Then
            'Dim strJournal As String = ""
            'For Each row In Me.dgvEntries.Rows
            '    With row
            '        If .Cells(Me.ColTrans_Status.Index).Value = Me.cboStatus.Items.Item(ClsTransaksiRecurring.EnumStatus.Planed) Then
            '            If .Cells(Me.ColTrans_PrimaryJournal.Index).Value = -1 Then
            '                If Me.ObjJournal.ApakahSudahAda(.Cells(Me.ColTrans_JournalNumber.Index).Value) = True Then
            '                    If strJournal = "" Then
            '                        strJournal &= .Cells(Me.ColTrans_JournalNumber.Index).Value
            '                    Else
            '                        strJournal &= ", " & .Cells(Me.ColTrans_JournalNumber.Index).Value
            '                    End If
            '                End If
            '            Else
            '                Dim CodeLama As String = Me.ObjJournal.GetRefCode(.Cells(Me.ColTrans_PrimaryJournal.Index).Value)
            '                If CodeLama <> .Cells(Me.ColTrans_JournalNumber.Index).Value Then
            '                    If Me.ObjJournal.ApakahSudahAda(.Cells(Me.ColTrans_JournalNumber.Index).Value) = True Then
            '                        If strJournal = "" Then
            '                            strJournal &= .Cells(Me.ColTrans_JournalNumber.Index).Value
            '                        Else
            '                            strJournal &= ", " & .Cells(Me.ColTrans_JournalNumber.Index).Value
            '                        End If
            '                    End If
            '                End If
            '            End If
            '        End If
            '    End With
            'Next
            'If strJournal = "" Then
            '    MsgBox("Duplication in no journal = " & strJournal & ", please change your number ", _
            '    MsgBoxStyle.Information, "Information")
            '    Me.dgvEntries.Focus()
            '    Return False
            'End If

            'cek data yang salah, double dalam kode
            If Me.Baru = True Then
                If Me.ObjRecurring.ApakahSudahAda(Me.TxtRefNo.Text) = True Then
                    Me.TxtRefNo.Text = Me.ObjRecurring.GetNoBaru(Me.dtpDate.Value.Date)
                    'MsgBox("Duplication in Recurring No: " + Me.TxtRefNo.Text + ", please change your number", _
                    'MsgBoxStyle.Information, "Information")
                    'Me.TxtRefNo.Focus()
                    'Return False
                End If
            Else
                Dim TipeCodeLama As String = Me.ObjRecurring.GetNoTransasi(Me.TxtPrimaryJournal.Text)
                If TipeCodeLama <> Me.TxtRefNo.Text Then
                    If Me.ObjRecurring.ApakahSudahAda(Me.TxtRefNo.Text) = True Then
                        Me.TxtRefNo.Text = Me.ObjRecurring.GetNoBaru(Me.dtpDate.Value.Date)
                        'MsgBox("Duplication in Recurring No =" + Me.TxtRefNo.Text + ", please change your number", _
                        'MsgBoxStyle.Information, "Information")
                        'Me.TxtRefNo.Focus()
                        'Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.ObjRecurring.GetTimestamp(Me.TxtPrimaryJournal.Text)
                If Me.ObjRecurring.Timestamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjRecurring.CariDetail(Me.ObjRecurring.Primary, Me.dgvItem, Me)
                        Me.SalinkeField()
                    End If
                    Return False
                End If
            End If

            If Me.dtpEndDate.Value.Date < Me.dtpStartDate.Value.Date Then
                MsgBox("Please make sure start date > end date.", MsgBoxStyle.Information, "Information")
                Me.dtpStartDate.Value = Me.BufferTglAwal
                Me.dtpStartDate.Focus()
                Return False
            End If

            'pastikan semua kode akun ada di dalam master akun
            For Each row In Me.dgvItem.Rows
                If IsNothing(row.Cells(Me.ColPrimaryAkun.Index).Value) = True Then
                    row.Selected = True
                    MsgBox("Invalid account in row " + row.Index.ToString, MsgBoxStyle.Information, "Information")
                    Me.dgvItem.ClearSelection()
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

            ''jika sudah ada closing jurnal untuk tahun ybs maka tidak bisa input data baru
            'If Me.ObjJurnal.GetJML(ClsTransaksiJournal.EnumTipeJournal.ClosingJournal, Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpDate.Value), _
            'Me.ObjCD.GetLastDateOfTheYear(Me.dtpDate.Value.Date)) > 0 Then
            '    MsgBox("There are closing jurnal for this year you cannot save this data", MsgBoxStyle.Information, "Information")
            '    Return False
            'End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Function

    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            Me.dgvItem.AllowUserToAddRows = False
            Me.dgvReminder.AllowUserToAddRows = False
            If Me.TxtPrimaryJournal.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtPrimaryJournal.Text <> "-1" Then
                If Me.ObjRecurring.ApakahBolehHapus(Me.TxtPrimaryJournal.Text, True, UserCOde, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        'Me.tms_DeleteAll_Click(Nothing, Nothing)
                        'If Me.BerhasilDelete = True Then
                        If Me.ObjRecurring.Hapus(Me.TxtPrimaryJournal.Text, True, UserCOde, Me) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                        ' End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Public Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        ' ij.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
        Try
            ij.BackColorAll = Me.BackcolorAll
            If Me.FromReminder = False Then
                ij.ShowDialog()
            End If
            Me.dgvItem.AllowUserToAddRows = False
            Me.dgvReminder.AllowUserToAddRows = False
            If Not ij.PrimaryPP = -1 Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                If Me.ObjRecurring.CariDetail(ij.PrimaryPP, Me.dgvItem, Me) = True Then
                    Me.SalinkeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            Me.dgvItem.AllowUserToAddRows = False
            Me.dgvReminder.AllowUserToAddRows = False
            If Me.TxtPrimaryJournal.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtPrimaryJournal.Text <> "-1" Then
                If Me.ObjRecurring.ApakahBolehEdit(Me.ObjRecurring.Primary, UserCOde, True, True) = True Then
                    Me.Baru = False
                    Me.ObjRecurring.CariDetail(Me.ObjRecurring.Primary, Me.dgvItem, Me)
                    Me.SalinkeField()
                    Me.AktifkanDetail()
                    Me.dgvItem.AllowUserToAddRows = True
                    Me.dgvReminder.AllowUserToAddRows = True
                    Me.TxtRefNo.Focus()
                End If
            End If


            'Me.dgvItem.AllowUserToAddRows = False
            'Me.dgvReminder.AllowUserToAddRows = False
            'If Me.TxtRefNo.Text = "" Then
            '    Me.btnFind.PerformClick()
            '    If Me.ObjRecurring.ApakahBolehEdit(Me.ObjRecurring.Primary, UserCOde, True, True) = True Then
            '        Me.Baru = False
            '        Me.ObjRecurring.CariDetail(Me.ObjRecurring.Primary, Me.dgvItem, Me)
            '        Me.SalinkeField()
            '        Me.AktifkanDetail()
            '        Me.dgvItem.AllowUserToAddRows = True
            '        Me.dgvReminder.AllowUserToAddRows = True
            '        Me.TxtRefNo.Focus()
            '    End If
            'Else
            '    Me.Baru = False
            '    Me.AktifkanDetail()
            '    Me.dgvItem.AllowUserToAddRows = True
            '    Me.dgvReminder.AllowUserToAddRows = True
            '    ' Me.ObjCD.SetNumber(Me.dgvItem)
            '    Me.ObjCD.SetNumber(Me.dgvReminder)
            'End If

            'If Me.TxtRefNo.Text = "" Then
            '    If Me.ObjRecurring.ApakahBolehEdit(Me.ObjRecurring.Primary, UserCOde, True, True) = True Then
            '        Me.Baru = False
            '        Me.ObjRecurring.CariDetail(Me.ObjRecurring.Primary, Me.dgvItem, Me)
            '        Me.SalinkeField()
            '        Me.AktifkanDetail()
            '        Me.dgvItem.AllowUserToAddRows = True
            '        Me.dgvReminder.AllowUserToAddRows = True
            '        Me.TxtRefNo.Focus()
            '    End If
            'Else
            '    Me.AktifkanDetail()
            'End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnViewJurnal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnViewJurnal.Click
        Dim FrmParent As Windows.Forms.Form = Me.MdiParent

        Try
            Dim iE As FrmTransaksiJournal_View
            Dim j As Windows.Forms.Form

            For Each j In FrmParent.MdiChildren
                iE = TryCast(j, FrmTransaksiJournal_View)
                If IsNothing(iE) = False Then
                    iE.BringToFront()
                    iE.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New FrmTransaksiJournal_View
            If Me.dgvEntries.Rows.Count > 0 Then
                i.dtpFrom.Value = CDate(Me.dgvEntries.Item(Me.ColTrans_JournalDate.Index, 0).Value)
                i.dtpTo.Value = CDate(Me.dgvEntries.Item(Me.ColTrans_JournalDate.Index, Me.dgvEntries.Rows.Count - 1).Value)
            Else
                i.dtpFrom.Value = Date.Now
                i.dtpTo.Value = Date.Now
            End If

            i.UserCode = Me.UserCOde
            ' Set the caption to be unique.
            i.BackCOlorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = FrmParent
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

        'Dim IJ2 As New FrmTransaksiJournal_View

        'Try
        '    'IJ2.UserCode = Me.UserCOde
        '    'If Me.dgvEntries.Rows.Count > 0 Then
        '    '    IJ2.dtpFrom.Value = CDate(Me.dgvEntries.Item(Me.ColTrans_JournalDate.Index, 0).Value)
        '    '    IJ2.dtpTo.Value = CDate(Me.dgvEntries.Item(Me.ColTrans_JournalDate.Index, Me.dgvEntries.Rows.Count - 1).Value)
        '    'Else
        '    '    IJ2.dtpFrom.Value = Date.Now
        '    '    IJ2.dtpTo.Value = Date.Now
        '    'End If
        '    'IJ2.ShowDialog()
        '    'IJ2.BringToFront()
        '    'IJ2.Visible = True


        '    Dim j As Windows.Forms.Form
        '    For Each j In Me.MdiChildren
        '        IJ2 = TryCast(j, FrmTransaksiJournal_View)
        '        If IsNothing(IJ2) = False Then
        '            IJ2.BringToFront()
        '            IJ2.Focus()
        '            Exit Sub
        '        End If
        '    Next
        '    'Dim i As New FrmTransaksiJournal_View
        '    ' Increment the caption counter.
        '    ''i.UserCode = FrmLogin.UsernameTextBox.Text
        '    ' Set the caption to be unique.
        '    ''            i.BackColorAll = Core_Module.GetSMESSBackColorAll
        '    IJ2.MdiParent = Me
        '    IJ2.Show()

        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'Finally
        '    IJ2 = Nothing
        '    'IJ2.Dispose()
        'End Try

    End Sub

    Private Function CekUserSama(ByVal vUserCode As String, ByVal idx As Integer) As Boolean
        Dim RowDgv As DataGridViewRow
        Try
            For Each RowDgv In Me.dgvReminder.Rows
                If RowDgv.Index <> idx Then
                    If RowDgv.Cells(Me.ColReminder_UserCode.Index).Value = vUserCode.ToUpper Then
                        Return True
                    End If
                End If
            Next

            Return False
        Catch ex As Exception
        End Try
    End Function

    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Try
            If Me.TxtPrimaryJournal.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtPrimaryJournal.Text <> "-1" Then
                Dim i As New FrmTransaksiRecuring_Print
                i.Primary = Me.TxtPrimaryJournal.Text
                i.ShowDialog()
                i.Dispose()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRefNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRefNo.GotFocus
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

    Private Sub TxtRefNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRefNo.LostFocus
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

    Private Sub TxtRemarks_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus
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

    Private Sub txtLimit_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLimit.LostFocus
        Try
            If Me.txtLimit.Enabled = True Then
                If Me.txtLimit.ReadOnly = False Then
                    Me.txtLimit.BackColor = Me.TxtPrimaryJournal.BackColor
                Else
                    txtLimit.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged

        Try
            If Me.TxtRemarks.Enabled = True AndAlso IsDate(Me.dtpDate.Value) Then
                If Me.Baru = False Then
                    If Me.ObjRecurring.Tanggal.Month = Me.dtpDate.Value.Month And Me.ObjRecurring.Tanggal.Year = Me.dtpDate.Value.Year Then
                        'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                        Me.TxtRefNo.Text = Me.ObjRecurring.RefNumber
                    Else
                        'jika pindah bulan maka akan ganti no
                        Me.TxtRefNo.Text = Me.ObjRecurring.GetNoBaru(Me.dtpDate.Value)
                    End If
                Else
                    Me.TxtRefNo.Text = Me.ObjRecurring.GetNoBaru(Me.dtpDate.Value)
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Dim rowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0

        Try
            Me.dgvItem.TabStop = True
            If Me.TxtRemarks.Enabled = True And Me.dgvItem.AllowUserToAddRows = True Then
                If Me.dgvItem.RowCount > 0 Then
                    If Me.rowIndexItem > 0 Then
                        Me.btnUP.Enabled = True
                        Me.btnDown.Enabled = True
                    End If
                   
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
                    rowDGV.Cells(Me.ColDept.Index).Value = ""
                    rowDGV.Cells(Me.ColJob.Index).Value = ""
                    rowDGV.Cells(Me.ColPrimaryDept.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryJob.Index).Value = -1
                    'rowDGV.Cells(Me.ColAcc_IncrementStyle.Index).Value = "No Increment"
                    'rowDGV.Cells(Me.ColAcc_IncrementValue.Index).Value = 0

                    'Me.dgvItem.Item(Me.ColAcc_IncrementStyle.Index, e.RowIndex).Value = Me.cboDGV.Items.Item(ClsTransaksiRecurring.EnumIncrementStyle.No_Increment)
                    Me.dgvItem.Item(Me.ColAcc_IncrementValue.Index, e.RowIndex).Value = ObjCD.DigitGrouping(0, DefPrecBase)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.SelectionChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.dgvItem.SelectedRows.Count > 0 Then
                    Me.btnUP.Enabled = True
                    Me.btnDown.Enabled = True
                Else
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class