Option Explicit On
Option Strict Off

Imports Core_Module
Public Class FrmTransaksiJournal
    Public UserCOde As String = ""
    Public BackcolorAll As Color = Color.LightSeaGreen
    Public TipeJournal As ClsTransaksiJournal.EnumTipeJournal = ClsTransaksiJournal.EnumTipeJournal.AdjustmenJournal
    Public TipeJournal2 As ClsTransaksiJournal.EnumTipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.Jurnal_NoGrouping
    'Kode Erna =================================================
    Public Berhasil As Boolean = False
    Public Recurring As Boolean = False
    Public HideMsg As Boolean = False
    Public Baru As Boolean
    Public ObjJurnal As New ClsTransaksiJournal
    Public DefaultPrecision As Integer = 0
     'Kode Erna =================================================
    Dim Buffer As String = ""
    Dim ObjAkun As New ClsMasterAkun
    Dim ObjCD As New ClsCekData
    Dim ObjQuickJurnal As New ClsTransaksiJournalQuickAdd
    Dim ObjLicense As New ClsLicense
    Dim ObjClosing As New ClsTransaksiClosing
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim PrimaryBaseCurrency As Double = 1
    Dim ObjTaxRate As New ClsTransaksiSalesInvoiceTaxRate
    Dim ObjDefaultValue As New Core_Module.ClsTransaksiSetting
    Dim ObjDept As New ClsMasterDepartement
    Dim ObjJob As New Core_Module.ClsMasterJob
    Dim ObjHakAkses As New Core_Module.ClsHakAkses
    Dim rowIndex As Integer = -1
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim strBufferCbo As String = ""


    Private Sub FrmTransaksiJournal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ''Me.cboJournalTipe.SelectedIndex = Me.TipeJournal
        Dim NullValue As String = ""

        Try
            If Me.Recurring = False Then
                Me.dgvItem.AllowUserToAddRows = False
                Me.ObjJurnal = New ClsTransaksiJournal
                Me.Baru = True
                Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackcolorAll
                Me.dgvItem.Rows.Clear()
                Me.KosongkanField()
                Me.NonAktifkanDetail()
                'End If
            End If

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
                Me.Text = "General Journal"
            End If
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackcolorAll

            'set null value untuk dgv
            Me.DefaultPrecision = Me.ObjCur.GetDecimal_DefaultCurrency
            NullValue = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            Me.ColDebet.DefaultCellStyle.NullValue = NullValue
            Me.ColKredit.DefaultCellStyle.NullValue = NullValue
            Me.TxtUnbalance.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)

            'jika tidak ada hak akses atas dept, maka kolom dept ditutup
            'If Me.ObjHakAkses.ApakahSudahAda("FrmMasterDepartement") = False Then
            Me.ColDept.Visible = False
            'End If

            'jika tidak ada hak akses atas job, maka kolom job ditutup
            'If Me.ObjHakAkses.ApakahSudahAda("FrmMasterJob") = False Then
            Me.ColJob.Visible = False
            Me.ShowIcon = False

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

    ''Pengaturan DGV Item
    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            ''Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProcess_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Dim rowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0

        Try
            Me.dgvItem.TabStop = True
            If Me.TxtRemarks.Enabled = True And Me.dgvItem.AllowUserToAddRows = True Then
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True

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
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProcess_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItem.RowsRemoved
        Try
            If Me.dgvItem.Rows.Count = 0 Then
                Me.dgvItem.TabStop = False
                If Me.TxtRemarks.Enabled = True Then
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
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

    Private Sub btnUP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUP.Click
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

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
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

    Private Sub dgvItemProcess_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.SelectionChanged
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

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmTransaksiJournal_Cetak
        ij.TipeJournal = Me.TipeJournal
        Try
            ij.BackColorAll = Me.BackcolorAll
            ij.ShowDialog()
            Me.dgvItem.AllowUserToAddRows = False
            If Not ij.PrimaryPP = -1 Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                If Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, ClsTransaksiSetting.EnumNamaParameter.Jurnal_IsLedgerAccessActive) = "1" Then
                    If Me.ObjJurnal.ApakahMemilikiLedgerAccess(Me.UserCOde, Me.ObjJurnal.GetCreatedBy(ij.PrimaryPP)) = False Then
                        MsgBox("You are not authorized to view this data.", MsgBoxStyle.Information, "Information")
                        Me.KosongkanField()
                        Exit Sub
                    End If
                End If

                If Me.ObjJurnal.CariDetail(ij.PrimaryPP, Me.dgvItem, Me, ij.TipeJournal) = True Then
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
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, False, False, False, True, True)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None

            Me.dgvItem.AllowUserToDeleteRows = True
            Me.dgvItem.ReadOnly = False

            If Me.dgvItem.Rows.Count = 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
            ''Me.txtNoTransaksi.ReadOnly = True
            ''Me.txtNoTransaksi.BackColor = Me.txtReadOnly.BackColor
            Me.cboQuickJournal.SelectedIndex = -1
            Me.llbEditQuickJournal.Enabled = True
            Me.ObjQuickJurnal.IsiCboQuickAdd(Me.TipeJournal, Me.cboQuickJournal)

            If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.ClosingJournal Then
                Me.dtpDate.Enabled = False
                Me.dgvItem.AllowUserToAddRows = False
                Me.dgvItem.ReadOnly = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, True, True, False, True, False)
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.ObjCD.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, False)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, False)
            Me.TxtPrimaryJournal.Text = "-1"
            Me.TxtUnbalance.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            rowIndex = -1
            Me.dgvItem.Rows.Clear()
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

    Public Sub SalinkeField()
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

            If Me.ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data default precision harus non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'Cek DGVItem
            If Me.dgvItem.Rows.Count < 2 Then
                MsgBox("Minimum item = 2, please input item again", MsgBoxStyle.Information, "Information")
                Me.dgvItem.Focus()
                Return False
            End If

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
            For Each row In Me.dgvItem.Rows
                If Me.dgvItem.Item(Me.ColDebet.Index, row.Index).Value = 0 And _
                Me.dgvItem.Item(Me.ColKredit.Index, row.Index).Value = 0 Then
                    MsgBox("Please add positive number in account no. " + Me.dgvItem.Item(Me.ColAccount.Index, row.Index).Value, _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Next

            'cek perubahan nomor transaksi
            ''If Me.baru = False Then
            ''If Me.ObjPP.TanggalTransaksi.Month = Me.dtpTanggalTransaksi.Value.Month And _
            ''Me.ObjPP.TanggalTransaksi.Year = Me.dtpTanggalTransaksi.Value.Year Then
            ''Me.txtNoTransaksi.Text = Me.ObjPP.NoTransaksi
            ''Else
            ''If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
            ''MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
            ''Me.dtpTanggalTransaksi.Value = Me.ObjPP.TanggalTransaksi
            ''Me.txtNoTransaksi.Text = Me.ObjPP.NoTransaksi
            ''Return False
            ''End If
            ''End If
            ''End If


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
                    MsgBox(Me.ObjJurnal.Timestamp & "," & OldTimeStamp)
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjJurnal.CariDetail(Me.ObjJurnal.Primary, Me.dgvItem, Me, Me.TipeJournal)
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
                    If Me.TipeJournal <> ClsTransaksiJournal.EnumTipeJournal.ClosingJournal Then '-> ditambahkan 2017-11-12: karena end of year mungkin saja menggunakan jurnal yang tidak aktif
                        If ObjAkun.ApakahSudahAda_Aktif_NonHeader2(row.Cells(Me.ColPrimaryAkun.Index).Value) = False Then
                            MsgBox("Invalid account in row " + row.Index.ToString, MsgBoxStyle.Information, "Information")
                            Me.dgvItem.ClearSelection()
                            row.Selected = True
                            Return False
                        End If
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

    Public Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
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

            Me.dgvItem.AllowUserToAddRows = True
            Me.Baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.TxtRefNo.Focus()
            Me.ObjJurnal.ResetObject(Me.dgvItem)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCOde
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCOde
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

            If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.ClosingJournal Then

                'set tanggal closing
                PrimaryDummy = Me.ObjJurnal.GetPrimaryMax(Me.TipeJournal)
                If PrimaryDummy = -1 Then
                    Me.dtpDate.Value = Me.ObjCD.GetLastDateOfTheYear(Me.ObjClosing.GetConvertionDate)
                Else
                    Me.dtpDate.Value = Me.ObjJurnal.GetTanggalJournal(PrimaryDummy)
                    Me.dtpDate.Value = Me.ObjCD.GetLastDateOfTheYear(Me.dtpDate.Value + TimeSpan.FromDays(10))
                End If

                'langsung quick closing
                Me.QuickClosing()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            Me.dgvItem.AllowUserToAddRows = False
            If Me.TxtRefNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtRefNo.Text <> "" Then
                If Me.ObjJurnal.ApakahBolehEdit(Me.ObjJurnal.Primary, UserCOde, True, True) = True Then
                    Me.Baru = False
                    If Me.Recurring = False Then
                        Me.ObjJurnal.CariDetail(Me.ObjJurnal.Primary, Me.dgvItem, Me, TipeJournal)
                        Me.SalinkeField()
                    End If
                    Me.AktifkanDetail()
                    Me.dgvItem.AllowUserToAddRows = True
                    Me.TxtRefNo.Focus()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            Me.dgvItem.AllowUserToAddRows = False
            If Me.TxtRefNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtRefNo.Text <> "" Then
                If Me.ObjJurnal.ApakahBolehHapus(Me.TxtPrimaryJournal.Text, True, UserCOde, Me.Recurring, True) = False Then
                    Me.btnFind.Focus()
                    'Kode Erna =================================================
                    Berhasil = False
                    'Kode Erna =================================================
                Else
                    If Me.Recurring = False Then
                        hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    Else
                        hasil = MsgBoxResult.Yes
                    End If

                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjJurnal.Hapus(Me.TxtPrimaryJournal.Text, True, UserCOde, Me.Recurring) = True Then
                            'Kode Erna =================================================
                            Berhasil = True
                            'Kode Erna =================================================
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            If Me.Recurring = False Then
                                MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                            End If
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

    Public Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            Me.dgvItem.AllowUserToAddRows = False
            Me.SalinkeObjek()
            If Me.dgvItem.Rows.Count > 0 Then
                If String.IsNullOrEmpty(Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Cells(Me.ColAccount.Index).Value) OrElse _
                      IsDBNull(Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Cells(Me.ColAccount.Index).Value) OrElse _
                Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Cells(Me.ColAccount.Index).Value = "" Then
                    Me.dgvItem.Rows.RemoveAt(Me.dgvItem.Rows.Count - 1)
                End If
            End If
            If Baru = True Then
                If ObjJurnal.ApakahDataBenar1(Baru, Me.dgvItem, Me.ColPrimaryCurrency.Index, Me.dtpDate.Value) = True Then
                    'konfirmasi disimpan atau tidak
                    If HideMsg = False Then
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    Else
                        hasil = MsgBoxResult.Yes
                    End If

                    If Me.ApakahDataBenar() = True Then
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjJurnal.SimpanBaru(Me.dgvItem, Me.UserCOde, Me, TipeJournal) Then
                                If HideMsg = False Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                End If
                                'Kode Erna =================================================
                                Berhasil = True
                                'Kode Erna =================================================
                                Me.SalinkeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            End If
                            Me.dgvItem.AllowUserToAddRows = False
                        ElseIf hasil = MsgBoxResult.No Then
                            'Kode Erna =================================================
                            Berhasil = False
                            'Kode Erna =================================================
                            Me.dgvItem.AllowUserToAddRows = True
                        End If
                    Else
                        'Kode Erna =================================================
                        Berhasil = False
                        'Kode Erna =================================================
                        Me.dgvItem.AllowUserToAddRows = True
                    End If
                Else
                    'Kode Erna =================================================
                    Berhasil = False
                    'Kode Erna =================================================
                    Me.dgvItem.AllowUserToAddRows = True
                End If
            Else
                'konfirmasi disimpan atau tidak
                If Me.ApakahDataBenar = True Then
                    If ObjJurnal.ApakahBolehEdit(Me.TxtPrimaryJournal.Text, UserCOde, True, True) = True Then
                        If HideMsg = False Then
                            hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        Else
                            hasil = MsgBoxResult.Yes
                        End If

                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjJurnal.SimpanPerubahan(Me.dgvItem, Me.UserCOde, Me, TipeJournal) Then
                                If HideMsg = False Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                End If
                                'Kode Erna =================================================
                                Berhasil = True
                                'Kode Erna =================================================
                                Me.SalinkeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            End If
                            Me.dgvItem.AllowUserToAddRows = False
                        ElseIf hasil = MsgBoxResult.No Then
                            'Kode Erna =================================================
                            Berhasil = False
                            'Kode Erna =================================================
                            Me.dgvItem.AllowUserToAddRows = True
                        End If
                    End If
                Else
                    'Kode Erna =================================================
                    Berhasil = False
                    'Kode Erna =================================================
                    Me.dgvItem.AllowUserToAddRows = True
                End If
            End If

            If Me.Recurring = True Then
                If Berhasil = True Then
                    Me.Close()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            'If Me.TxtRemarks.Enabled = True Then
            '    Me.dgvItem.AllowUserToAddRows = True
            'End If
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult

        Try
            Me.dgvItem.AllowUserToAddRows = False
            If Me.dgvItem.Rows.Count > 0 Then
                If String.IsNullOrEmpty(Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Cells(Me.ColAccount.Index).Value) OrElse _
                      IsDBNull(Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Cells(Me.ColAccount.Index).Value) OrElse _
                Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Cells(Me.ColAccount.Index).Value = "" Then
                    Me.dgvItem.Rows.RemoveAt(Me.dgvItem.Rows.Count - 1)
                End If
            End If

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
                        Me.ObjJurnal.CariDetail(Me.TxtPrimaryJournal.Text, Me.dgvItem, Me, TipeJournal)
                        Me.SalinkeField()
                    End If
                Else
                    Me.dgvItem.AllowUserToAddRows = True
                    Exit Sub
                End If
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                If Me.Recurring = True Then
                    Me.Close()
                End If
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
                    Me.dgvItem.AllowUserToAddRows = False
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinkeObjek()
                        If Baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjJurnal.SimpanBaru(Me.dgvItem, Me.UserCOde, Me, TipeJournal) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                'Kode Erna =================================================
                                Berhasil = True
                                'Kode Erna =================================================
                                Me.SalinkeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                'Kode Erna =================================================
                                Berhasil = False
                                'Kode Erna =================================================
                                Me.dgvItem.AllowUserToAddRows = True
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjJurnal.SimpanPerubahan(Me.dgvItem, Me.UserCOde, Me, TipeJournal) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                'Kode Erna =================================================
                                Berhasil = True
                                'Kode Erna =================================================
                                Me.SalinkeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            Else
                                Me.dgvItem.AllowUserToAddRows = True
                                'Kode Erna =================================================
                                Berhasil = False
                                'Kode Erna =================================================
                                e.Cancel = True
                            End If
                        End If
                    Else
                        'Kode Erna =================================================
                        Berhasil = False
                        'Kode Erna =================================================
                        Me.dgvItem.AllowUserToAddRows = True
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    'Kode Erna =================================================
                    Berhasil = False
                    'Kode Erna =================================================
                    e.Cancel = True
                End If
            End If
            Me.dgvItem.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''Fungsi DGV
    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColAccount.Index Then
                If e.ColumnIndex <> Me.ColRemarks.Index And e.ColumnIndex <> Me.ColOriginalCurrency.Index And e.ColumnIndex <> Me.ColDescAccount.Index Then
                    If e.ColumnIndex = Me.ColDebet.Index Or e.ColumnIndex = Me.ColKredit.Index Or _
                    e.ColumnIndex = Me.ColDebet_OriginalCurr.Index Or e.ColumnIndex = Me.ColCredit_OriginalCurrency.Index Or _
                    e.ColumnIndex = Me.ColExchangeRate.Index Or e.ColumnIndex = Me.ColJob.Index Or _
                    e.ColumnIndex = Me.ColDept.Index Then
                        'jika belum diisikan account codenya, maka tidak bisa edit dept dan job
                        If e.ColumnIndex = Me.ColDept.Index Or e.ColumnIndex = Me.ColJob.Index Then
                            If Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value = -1 Then
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

    Public Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
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
                        Me.rowIndex = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                        '  Me.dgvItem.Rows.RemoveAt(e.RowIndex)
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
                                        Me.rowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                        ' Me.dgvItem.Rows.RemoveAt(e.RowIndex)
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
                        'digit gropin
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

    Private Sub QuickClosing()
        Dim DTA As New DseTransaksiJournalTableAdapters.tblmaster_akunTableAdapter
        Dim TBL As New DseTransaksiJournal.tblmaster_akunDataTable
        Dim row As DseTransaksiJournal.tblmaster_akunRow
        Dim ObjSetting As New ClsTransaksiClosing
        Dim ConvertionDate As Date
        Dim Debet, Kredit As Decimal
        Dim TotalDebet As Decimal = 0
        Dim TotalKredit As Decimal = 0
        Dim StartPeriod, EndPeriod As Date
        Dim rowdgv As DataGridViewRow
        Dim PrimaryAkunRetainedEarning As String = -1

        Try
            ConvertionDate = ObjSetting.GetConvertionDate

            Me.dgvItem.AllowUserToAddRows = False
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            StartPeriod = Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpDate.Value.Date)
            EndPeriod = Me.ObjCD.GetLastDateOfTheYear(Me.dtpDate.Value.Date)

            'hapuskan baris yang ada
            Me.dgvItem.Rows.Clear()

            'masukkan jurnal closing
            ConvertionDate = Me.ObjCD.GetFisrtDateOfTheYear(ConvertionDate)
            DTA.Fill(TBL, EndPeriod, StartPeriod)

            'jika yang diclosing adalah period pertama, maka harus masukkan data beginning balance
            If StartPeriod = Me.ObjCD.GetFirstDateofTheMonth(ConvertionDate) Then
                For Each row In TBL.Rows
                    If row.Saldo_Normal = 0 Then
                        row.Debet += row.Beginning_Balance
                    Else
                        row.Kredit += row.Beginning_Balance
                    End If
                Next
            End If

            Me.dgvItem.AllowUserToAddRows = False
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            For Each row In TBL.Rows
                With row
                    If row.Debet <> 0 Or row.Kredit <> 0 Then
                        If row.Debet > row.Kredit Then
                            Kredit = row.Debet - row.Kredit
                            Debet = 0
                        Else
                            Debet = row.Kredit - row.Debet
                            Kredit = 0
                        End If
                        If Kredit <> 0 Or Debet <> 0 Then
                            Me.dgvItem.Rows.Add()
                            rowdgv = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                            With rowdgv
                                .Cells(Me.ColAccount.Index).Value = row.Account_Code
                                .Cells(Me.ColDebet.Index).Value = Me.ObjCD.DigitGrouping(Debet, Me.DefaultPrecision)
                                .Cells(Me.ColKredit.Index).Value = Me.ObjCD.DigitGrouping(Kredit, Me.DefaultPrecision)
                                .Cells(Me.ColDescAccount.Index).Value = row.Description
                                .Cells(Me.ColPrimaryAkun.Index).Value = row.Primary
                                .Cells(Me.ColRemarks.Index).Value = ""
                                .Cells(Me.ColPrimaryCurrency.Index).Value = 1
                                .Cells(Me.ColExchangeRate.Index).Value = 1
                                .Cells(Me.ColDebet_OriginalCurr.Index).Value = .Cells(Me.ColDebet.Index).Value
                                .Cells(Me.ColCredit_OriginalCurrency.Index).Value = .Cells(Me.ColKredit.Index).Value
                            End With
                        End If
                    End If
                End With
            Next

            Me.dgvItem.AllowUserToAddRows = False
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'hitung unbalance
            Me.HitungUnbalance()

            'masukkan data retained earning
            PrimaryAkunRetainedEarning = ObjDefaultValue.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, ClsTransaksiSetting.EnumNamaParameter.Akun_Retained_Earning)
            Me.dgvItem.Rows.Add()
            rowdgv = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            With rowdgv
                .Cells(Me.ColAccount.Index).Value = Me.ObjAkun.GetAccountCode(PrimaryAkunRetainedEarning)
                .Cells(Me.ColDescAccount.Index).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(PrimaryAkunRetainedEarning)
                .Cells(Me.ColPrimaryAkun.Index).Value = PrimaryAkunRetainedEarning
                .Cells(Me.ColRemarks.Index).Value = ""
                .Cells(Me.ColPrimaryCurrency.Index).Value = Me.ObjAkun.GetPrimaryCurrency(PrimaryAkunRetainedEarning)
            End With

            If CDec(Me.TxtUnbalance.Text) > 0 Then
                'masukkan retained earning ke dalam credit
                With rowdgv
                    .Cells(Me.ColDebet.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                    .Cells(Me.ColKredit.Index).Value = Me.ObjCD.DigitGrouping(Math.Abs(CDec(Me.TxtUnbalance.Text)), Me.DefaultPrecision)
                End With
            Else
                'masukkan retained earning ke dalam debet
                With rowdgv
                    .Cells(Me.ColDebet.Index).Value = Me.ObjCD.DigitGrouping(Math.Abs(CDec(Me.TxtUnbalance.Text)), Me.DefaultPrecision)
                    .Cells(Me.ColKredit.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                End With
            End If

            With rowdgv
                If .Cells(Me.ColPrimaryCurrency.Index).Value = Core_Module.PrimaryBaseCurrency Then
                    .Cells(Me.ColExchangeRate.Index).Value = 1
                    .Cells(Me.ColDebet_OriginalCurr.Index).Value = .Cells(Me.ColDebet.Index).Value
                    .Cells(Me.ColCredit_OriginalCurrency.Index).Value = .Cells(Me.ColKredit.Index).Value
                Else
                    MsgBox("Logic error, please contact your vendor if you see this message", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End With

            Me.HitungUnbalance()
            Me.SetMulticurrency()

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

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

    Private Sub InsertQuickJurnal(ByVal vConfirm As Boolean)
        Dim rowDGV As DataGridViewRow
        Dim ExcRate As Decimal = 1
        Dim DefPrec As Integer = 0
        Dim AdaNonIDR As Boolean = False
        Dim IndexAcuan As Integer = 0
        Dim DebetKah As Boolean = True
        ' Dim BlnNo As Boolean = False
        Dim tbl As New DataTable

        Try
            tbl.Rows.Clear()
            tbl.Columns.Clear()
            For intX As Integer = 0 To Me.dgvItem.Columns.Count - 1
                tbl.Columns.Add()
            Next

            For intY As Integer = 0 To Me.dgvItem.Rows.Count - 1
                tbl.Rows.Add()
                For intX As Integer = 0 To Me.dgvItem.Columns.Count - 1
                    tbl.Rows(intY)(intX) = Me.dgvItem.Rows(intY).Cells(intX).Value
                Next
            Next

          
            If Me.ObjQuickJurnal.IsiDGVQuickAdd(Me.TipeJournal, Me.dgvItem, 1, _
                Me.cboQuickJournal.Text, Me.ColAccount.Index, Me.ColDescAccount.Index, Me.ColPrimaryAkun.Index, _
                Me.ColDebet.Index, Me.ColKredit.Index, Me.ColRemarks.Index, Me.ObjCD, Me.DefaultPrecision, vConfirm, _
                Me.ColPrimaryDept.Index, Me.ColPrimaryJob.Index, Me.cboQuickJournal.Text) = False Then
                Me.cboQuickJournal.Text = strBufferCbo
                Me.dgvItem.Rows.Clear()
                For intA As Integer = 0 To tbl.Rows.Count - 1
                    Me.dgvItem.Rows.Add()
                    For intB As Integer = 0 To tbl.Columns.Count - 1
                        Me.dgvItem.Rows(intA).Cells(intB).Value = tbl.Rows(intA)(intB)
                    Next
                Next
            Else
                'If BlnNo = True Then
                '    Me.cboQuickJournal.SelectedText = ""
                'End If

                Me.TxtRemarks.Text = Me.ObjQuickJurnal.GetDesc_ByCode(Me.cboQuickJournal.Text)

                For Each rowDGV In Me.dgvItem.Rows
                    ExcRate = 1
                    If rowDGV.Index = Me.dgvItem.RowCount - 1 Then
                        If Me.dgvItem.AllowUserToAddRows = True Then
                            Exit For
                        End If
                    End If

                    'isikan primary currency
                    rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value = _
                    Me.ObjAkun.GetPrimaryCurrency(rowDGV.Cells(Me.ColPrimaryAkun.Index).Value)
                    rowDGV.Cells(Me.ColOriginalCurrency.Index).Value = Me.ObjCur.GetCurrencyCode(rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value)
                    'isikan exchange rate
                    If rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value <> Core_Module.PrimaryBaseCurrency Then
                        ExcRate = Me.ObjTaxRate.GetTaxRate(Me.dtpDate.Value, rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value)
                        If ExcRate = 0 Then
                            ExcRate = Me.ObjCur.GetDefaultExhangeRate(rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value)
                            If ExcRate = 0 Then
                                ExcRate = 1
                            End If
                        End If
                        DefPrec = Me.ObjCur.GetDefaultPrecission(rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value)
                    Else
                        ExcRate = 1
                        DefPrec = Me.ObjCur.GetDecimal_DefaultCurrency
                    End If
                    rowDGV.Cells(Me.ColExchangeRate.Index).Value = Me.ObjCD.DigitGrouping(ExcRate, DefPrecExchangeRate)

                    If rowDGV.Cells(Me.ColDebet.Index).Value <> 0 Then
                        rowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value = Me.ObjCD.DigitGrouping(1, DefPrec)
                        rowDGV.Cells(Me.ColCredit_OriginalCurrency.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
                    Else
                        rowDGV.Cells(Me.ColCredit_OriginalCurrency.Index).Value = Me.ObjCD.DigitGrouping(1, DefPrec)
                        rowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
                    End If

                    'hitung original values
                    Me.HitungBaseValues(rowDGV.Index)
                Next

                'kalo dia ada 2 baris dan salah satunya non IDR, maka non IDR-nya di edit lalu dibalance
                If Me.dgvItem.RowCount = 3 Then
                    For Each rowDGV In Me.dgvItem.Rows
                        If rowDGV.Index <> Me.dgvItem.Rows.Count - 1 Then
                            If rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value <> Core_Module.PrimaryBaseCurrency Then
                                AdaNonIDR = True
                                If rowDGV.Cells(Me.ColDebet.Index).Value > 0 Then
                                    DebetKah = True
                                Else
                                    DebetKah = False
                                End If
                                IndexAcuan = rowDGV.Index
                                Exit For
                            End If
                        End If
                    Next

                    If AdaNonIDR = True Then
                        Me.BalanceDebetKredit(IndexAcuan, DebetKah, Me.ObjCur.GetDecimal_DefaultCurrency)
                    End If
                End If
            End If

           
            'cek unbalance
            Me.HitungUnbalance()
            'set multicurrency
            Me.SetMulticurrency()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboQuickJournal_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboQuickJournal.GotFocus

        Try
            If Not String.IsNullOrEmpty(Me.cboQuickJournal.Text) Then
                strBufferCbo = Me.cboQuickJournal.Text
            Else
                strBufferCbo = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try


    End Sub

    Private Sub cboQuickJournal_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboQuickJournal.Leave
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.cboQuickJournal.SelectedIndex <> -1 Then
                    If Me.cboQuickJournal.Text <> "" Then
                        Me.InsertQuickJurnal(True)
                    End If
                End If
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

    Public Sub BalanceDebetKredit(ByVal vIndexAcuan As Integer, ByVal vAcuanDebet As Boolean, ByVal rounding As Integer, _
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
                rowTarget = Me.dgvItem.Rows(vIndexAcuan)
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
                End If

                If vHitungOriginalCurrency = True Then
                    If IsNothing(rowTarget.Index) = False Then
                        Me.HitungOriginalValues(rowTarget.Index)
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

    Public Sub SetMulticurrency()
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

    Public Sub HitungOriginalValues(ByVal vRowIndex As Integer)
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

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted

        Try
             Me.dgvItem.Rows.RemoveAt(Me.rowIndex)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

End Class