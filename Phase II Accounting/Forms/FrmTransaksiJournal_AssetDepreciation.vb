Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiJournal_AssetDepreciation
    Public UserCode As String = ""
    Public BackcolorAll As Color = Color.LightSeaGreen
    Public TipeJournal_Asset As ClsTransaksiJournal.EnumTipeJournal = ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
    Public TipeJournal2_Asset As ClsTransaksiJournal.EnumTipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.AssetDepreciation

    Dim Buffer As String = ""
    Dim DefaultPrecision_Asset As Integer = 0
    Dim Baru As Boolean
    Dim ObjAkun As New ClsMasterAkun
    Dim ObjJurnal_Asset As New ClsTransaksiJournal
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

    Private Sub FrmTransaksiJournal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ''Me.cboJournalTipe.SelectedIndex = Me.TipeJournal_Asset 
        Dim NullValue As String = ""

        Try
            Me.dgvAsset_Item.AllowUserToAddRows = False
            Me.ObjJurnal_Asset = New ClsTransaksiJournal
            Me.Baru = True
            Me.dgvAsset_Item.AlternatingRowsDefaultCellStyle.BackColor = Me.BackcolorAll
            Me.dgvAsset_Item.Rows.Clear()
            Me.KosongkanField_Asset()
            Me.NonAktifkanDetail_Asset()
            If Me.TipeJournal_Asset = ClsTransaksiJournal.EnumTipeJournal.AdjustmenJournal Then
                Me.Text = "Adjustment Journal"
            ElseIf Me.TipeJournal_Asset = ClsTransaksiJournal.EnumTipeJournal.ClosingJournal Then
                Me.Text = "Closing Journal"
                'persiapkan untuk closing 
                Me.dtpAsset_Date.CustomFormat = "yyyy"
                'jika sudah ada jurnal closing sebelumnya, maka tanggal closing harus setahun setelahnya
                If Me.ObjJurnal_Asset.GetJml(Me.TipeJournal_Asset) > 0 Then
                    Me.dtpAsset_Date.Enabled = False
                End If
            ElseIf Me.TipeJournal_Asset = ClsTransaksiJournal.EnumTipeJournal.GeneralJournal Then
                Me.Text = "General Journal"
            End If
            Me.dgvAsset_Item.AlternatingRowsDefaultCellStyle.BackColor = Me.BackcolorAll

            'set null value untuk dgv
            Me.DefaultPrecision_Asset = Me.ObjCur.GetDecimal_DefaultCurrency
            NullValue = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision_Asset)
            Me.ColAsset_Debet.DefaultCellStyle.NullValue = NullValue
            Me.ColAsset_Kredit.DefaultCellStyle.NullValue = NullValue
            Me.TxtAsset_Unbalance.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision_Asset)

            'jika tidak ada hak akses atas dept, maka kolom dept ditutup
            'If Me.ObjHakAkses.ApakahSudahAda("FrmMasterDepartement") = False Then
            Me.ColAsset_Dept.Visible = False
            'End If

            'jika tidak ada hak akses atas job, maka kolom job ditutup
            'If Me.ObjHakAkses.ApakahSudahAda("FrmMasterJob") = False Then
            Me.ColAsset_Job.Visible = False
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnViewJurnal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAsset_ViewJurnal.Click
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
            i.UserCode = Me.UserCode
            ' Set the caption to be unique.
            i.BackCOlorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = FrmParent
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''Pengaturan DGV Item
    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvAsset_Item.Sorted
        Try
            ''Me.ObjCD.SetNumber(Me.dgvAsset_Item)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProcess_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvAsset_Item.RowsAdded
        Dim rowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0

        Try
            Me.dgvAsset_Item.TabStop = True
            If Me.TxtAsset_Remarks.Enabled = True And Me.dgvAsset_Item.AllowUserToAddRows = True Then

                'set nilai default baris baru
                rowDGV = Me.dgvAsset_Item.Rows(Me.dgvAsset_Item.RowCount - 1)
                rowDGV.Cells(Me.ColAsset_PrimaryCurrency.Index).Value = Me.PrimaryBaseCurrency
                DefPrecBase = Me.ObjCur.GetDefaultPrecission(Me.PrimaryBaseCurrency)
                rowDGV.Cells(Me.ColAsset_Debet_OriginalCurr.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrecBase)
                rowDGV.Cells(Me.ColAsset_Credit_OriginalCurrency.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrecBase)
                rowDGV.Cells(Me.ColAsset_Debet.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrecBase)
                rowDGV.Cells(Me.ColAsset_Kredit.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrecBase)
                rowDGV.Cells(Me.ColAsset_ExchangeRate.Index).Value = Me.ObjCD.DigitGrouping(1, DefPrecExchangeRate)
                rowDGV.Cells(Me.ColAsset_PrimaryCurrency.Index).Value = Me.PrimaryBaseCurrency
                rowDGV.Cells(Me.ColAsset_OriginalCurrency.Index).Value = Me.ObjCur.GetCurrencyCode(Me.PrimaryBaseCurrency)
                rowDGV.Cells(Me.ColAsset_PrimaryAkun.Index).Value = -1
                rowDGV.Cells(Me.ColAsset_Account.Index).Value = ""
                rowDGV.Cells(Me.ColAsset_Remarks.Index).Value = ""
                rowDGV.Cells(Me.ColAsset_Dept.Index).Value = ""
                rowDGV.Cells(Me.ColAsset_Job.Index).Value = ""
                rowDGV.Cells(Me.ColAsset_PrimaryDept.Index).Value = -1
                rowDGV.Cells(Me.ColAsset_PrimaryJob.Index).Value = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProcess_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvAsset_Item.RowsRemoved
        Try
            If Me.dgvAsset_Item.Rows.Count = 0 Then
                Me.dgvAsset_Item.TabStop = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub dgvItemProcess_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvAsset_Item.UserDeletedRow
        Try
            ''hitung debet dan kredit
            Dim row As DataGridViewRow
            Dim TotalDebet As Decimal = 0
            Dim TotalKredit As Decimal = 0
            For Each row In dgvAsset_Item.Rows
                TotalDebet += CDec(Me.dgvAsset_Item.Item(Me.ColAsset_Debet.Index, row.Index).Value)
                TotalKredit += CDec(Me.dgvAsset_Item.Item(Me.ColAsset_Kredit.Index, row.Index).Value)
            Next
            Me.TxtAsset_Unbalance.Text = CStr(ObjCD.DigitGrouping(Me.HitungUnbalance_Asset(TotalDebet, TotalKredit), Me.DefaultPrecision_Asset))
            Me.SetMulticurrency_Asset()
            'If Me.TxtRemarks.Enabled = True Then
            'If Me.dgvAsset_Item.RowCount = 0 Then
            ''
            'End If
            'End If
            ''Me.ObjCD.SetNumber(Me.dgvAsset_Item)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' tampilkan row yang diberikan pada parameter
    ''' </summary>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Private Sub ShowRowProcess_Asset(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < Me.dgvAsset_Item.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    Me.dgvAsset_Item.FirstDisplayedScrollingRowIndex = Me.dgvAsset_Item.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.dgvAsset_Item.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While Me.dgvAsset_Item.Rows(row.Index).Displayed = False
                    Me.dgvAsset_Item.FirstDisplayedScrollingRowIndex = Me.dgvAsset_Item.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If Me.dgvAsset_Item.DisplayedRowCount(True) <> Me.dgvAsset_Item.DisplayedRowCount(False) Then
                    Me.dgvAsset_Item.FirstDisplayedScrollingRowIndex = Me.dgvAsset_Item.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAsset_Find.Click
        Dim ij As New FrmTransaksiJournal_Cetak
        ij.TipeJournal = Me.TipeJournal_Asset
        Try
            ij.BackColorAll = Me.BackcolorAll
            ij.ShowDialog()
            Me.dgvAsset_Item.AllowUserToAddRows = False

            If Not ij.PrimaryPP = -1 Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                If Me.ObjJurnal_Asset.CariDetail_DepresiasiAsset(ij.PrimaryPP, Me.dgvAsset_Item, Me, ij.TipeJournal, Me.dgvAsset_ItemAsset) = True Then
                    Me.SalinkeField_Asset()
                    Me.btnAsset_Edit.Focus()
                Else
                    Me.btnAsset_Find.Focus()
                End If
                Me.btnAsset_Find.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    ''fungsi standar 
    Private Sub AktifkanDetail_Asset()
        Try
            Me.ObjCD.SetBackColor(Me.gboAsset_Detail.Controls, Me.TxtAsset_PrimaryJournal.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1_Asset.Panel1.Controls, Me.TxtAsset_PrimaryJournal.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1_Asset.Panel2.Controls, Me.TxtAsset_PrimaryJournal.BackColor, False)
            ModCekData.KendalikanControl(Me.gboAsset_Detail.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1_Asset.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1_Asset.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, True, True)
            Me.CancelButton = Me.btnAsset_Cancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None

            Me.dgvAsset_Item.AllowUserToDeleteRows = True
            Me.dgvAsset_Item.ReadOnly = False

            If Me.TipeJournal_Asset = ClsTransaksiJournal.EnumTipeJournal.ClosingJournal Then
                Me.dtpAsset_Date.Enabled = False
                Me.dgvAsset_Item.AllowUserToAddRows = False
                Me.dgvAsset_Item.ReadOnly = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail_Asset()
        Try
            ModCekData.KendalikanControl(Me.gboAsset_Detail.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1_Asset.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1_Asset.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, True, False, True, False)
            Me.CancelButton = Me.btnAsset_Exit

            Me.dgvAsset_Item.AllowUserToDeleteRows = False
            Me.dgvAsset_Item.ReadOnly = True

            Me.ObjCD.SetBackColor(Me.gboAsset_Detail.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1_Asset.Panel1.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1_Asset.Panel2.Controls, Me.BackColor, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField_Asset()
        Try
            ModCekData.KosongkanControl(Me.gboAsset_Detail.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer1_Asset.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer1_Asset.Panel2.Controls, True, True, False)
            Me.TxtAsset_PrimaryJournal.Text = "-1"
            Me.TxtAsset_Unbalance.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision_Asset)
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "

            Me.dgvAsset_Item.Rows.Clear()
            ''Me.ObjCD.SetNumber(Me.dgvAsset_Item)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinkeObjek_Asset()
        Try
            With ObjJurnal_Asset
                .RefNumber = Me.TxtRefNo.Text
                .Description = Me.TxtAsset_Remarks.Text
                .Tanggal = Me.dtpAsset_Date.Value
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinkeField_Asset()
        Try
            With ObjJurnal_Asset
                Me.TxtAsset_PrimaryJournal.Text = .Primary
                Me.TxtRefNo.Text = .RefNumber
                Me.TxtAsset_Remarks.Text = .Description
                Me.dtpAsset_Date.Value = .Tanggal
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString

                Me.SetMulticurrency_Asset()
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar_Asset() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim row As DataGridViewRow

        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("Reference number")
            KoleksiTextBox.Add(Me.TxtRefNo)
            KoleksiNamaData.Add("Description")
            KoleksiTextBox.Add(Me.TxtAsset_Remarks)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data default precision harus non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'Cek dgvAsset_Item
            If Me.dgvAsset_Item.Rows.Count < 2 Then
                MsgBox("Minimum item = 2, please input item again", MsgBoxStyle.Information, "Information")
                Me.dgvAsset_Item.Focus()
                Return False
            End If

            ' Cek debet kredit harus balance
            If CDec(Me.TxtAsset_Unbalance.Text) <> 0 Then
                Me.HitungUnbalance_Asset()
                MsgBox("This is an unbalance account. Please add numbers in debet or credit column so that a balanced calculation", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'tidak boleh bikin transaksi di belakang convertion date
            If Me.ObjClosing.GetConvertionDate.Date > Me.dtpAsset_Date.Value.Date Then
                MsgBox("Minimum transaction date is " + Format(Me.ObjClosing.GetConvertionDate.Date, "dd MMM yyyy") + ".", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek apakah ada account number yang debet dan kreditnya bernilai 0
            For Each row In Me.dgvAsset_Item.Rows
                If Me.dgvAsset_Item.Item(Me.ColAsset_Debet.Index, row.Index).Value = 0 And _
                Me.dgvAsset_Item.Item(Me.ColAsset_Kredit.Index, row.Index).Value = 0 Then
                    MsgBox("Please add positive number in account no. " + Me.dgvAsset_Item.Item(Me.ColAsset_Account.Index, row.Index).Value, _
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
                If Me.ObjJurnal_Asset.ApakahSudahAda(Me.TxtRefNo.Text) = True Then
                    MsgBox("Duplication in no journal =" + Me.TxtRefNo.Text + ", please change your number", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtRefNo.Focus()
                    Return False
                End If
            Else
                Dim TipeCodeLama As String = Me.ObjJurnal_Asset.GetRefCode(Me.TxtAsset_PrimaryJournal.Text)
                If TipeCodeLama <> Me.TxtRefNo.Text Then
                    If Me.ObjJurnal_Asset.ApakahSudahAda(Me.TxtRefNo.Text) = True Then
                        MsgBox("Duplication in no journal =" + Me.TxtRefNo.Text + ", please change your number", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtRefNo.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.ObjJurnal_Asset.GetTimestamp(Me.TxtAsset_PrimaryJournal.Text)
                If Me.ObjJurnal_Asset.Timestamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjJurnal_Asset.CariDetail_DepresiasiAsset(Me.ObjJurnal_Asset.Primary, Me.dgvAsset_Item, Me, Me.TipeJournal_Asset, Me.dgvAsset_ItemAsset)
                        Me.SalinkeField_Asset()
                    End If
                    Return False
                End If
            End If

            'pastikan semua kode akun ada di dalam master akun
            For Each row In Me.dgvAsset_Item.Rows
                If IsNothing(row.Cells(Me.ColAsset_PrimaryAkun.Index).Value) = True Then
                    MsgBox("Invalid account in row " + row.Index.ToString, MsgBoxStyle.Information, "Information")
                    Me.dgvAsset_Item.ClearSelection()
                    row.Selected = True
                    Return False
                Else
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader2(row.Cells(Me.ColAsset_PrimaryAkun.Index).Value) = False Then
                        MsgBox("Invalid account in row " + row.Index.ToString, MsgBoxStyle.Information, "Information")
                        Me.dgvAsset_Item.ClearSelection()
                        row.Selected = True
                        Return False
                    End If
                End If
            Next

            'jika sudah ada closing jurnal untuk tahun ybs maka tidak bisa input data baru
            If Me.ObjJurnal_Asset.GetJml(ClsTransaksiJournal.EnumTipeJournal.ClosingJournal, Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpAsset_Date.Value), _
            Me.ObjCD.GetLastDateOfTheYear(Me.dtpAsset_Date.Value.Date)) > 0 Then
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

    Private Sub txtRemarks_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAsset_Remarks.Enter
        Try
            Me.Buffer = Me.TxtAsset_Remarks.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''back color dan tampilan 
    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAsset_Remarks.GotFocus
        Try
            If Me.TxtAsset_Remarks.Enabled = True Then
                If Me.TxtAsset_Remarks.ReadOnly = False Then
                    Me.TxtAsset_Remarks.BackColor = Me.BackcolorAll
                Else
                    TxtAsset_Remarks.BackColor = Me.txtAsset_ReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAsset_Remarks.LostFocus
        Try
            If Me.TxtAsset_Remarks.Enabled = True Then
                If Me.TxtAsset_Remarks.ReadOnly = False Then
                    Me.TxtAsset_Remarks.BackColor = Me.TxtAsset_PrimaryJournal.BackColor
                Else
                    TxtAsset_Remarks.BackColor = Me.txtAsset_ReadOnly.BackColor
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
                    TxtRefNo.BackColor = Me.txtAsset_ReadOnly.BackColor
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
                    Me.TxtRefNo.BackColor = Me.TxtAsset_PrimaryJournal.BackColor
                Else
                    TxtRefNo.BackColor = Me.txtAsset_ReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAsset_New.Click
        Dim PrimaryDummy As Double = -1
        Dim PrimaryAkunRE As String = ""

        Try
            TipeJournal_Asset = ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
            TipeJournal2_Asset = ClsTransaksiJournal.EnumTipeJurnal2.AssetDepreciation
            Me.ObjJurnal_Asset.TipeJournal = Me.TipeJournal_Asset
            Me.ObjJurnal_Asset.TipeJurnal2 = Me.TipeJournal2_Asset

            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.TipeJournal_Asset = ClsTransaksiJournal.EnumTipeJournal.ClosingJournal Then
                'jika akun retained earning tidak ada maka tidak bisa lanjut
                PrimaryAkunRE = Me.ObjDefaultValue.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, ClsTransaksiSetting.EnumNamaParameter.Akun_CurrenctYearEarning) = ""
                If PrimaryAkunRE = "" Or PrimaryAkunRE = "-1" Then
                    MsgBox("Please specify retained earning account in global setting data.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            Me.dgvAsset_Item.AllowUserToAddRows = True
            Me.Baru = True
            Me.KosongkanField_Asset()
            Me.AktifkanDetail_Asset()
            Me.TxtRefNo.Focus()
            Me.ObjJurnal_Asset.ResetObject(Me.dgvAsset_Item)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

            If Me.TipeJournal_Asset = ClsTransaksiJournal.EnumTipeJournal.ClosingJournal Then
                'set tanggal closing
                PrimaryDummy = Me.ObjJurnal_Asset.GetPrimaryMax(Me.TipeJournal_Asset)
                If PrimaryDummy = -1 Then
                    Me.dtpAsset_Date.Value = Me.ObjCD.GetLastDateOfTheYear(Me.ObjClosing.GetConvertionDate)
                Else
                    Me.dtpAsset_Date.Value = Me.ObjJurnal_Asset.GetTanggalJournal(PrimaryDummy)
                    Me.dtpAsset_Date.Value = Me.ObjCD.GetLastDateOfTheYear(Me.dtpAsset_Date.Value + TimeSpan.FromDays(10))
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAsset_Edit.Click
        Try
            Me.dgvAsset_Item.AllowUserToAddRows = False
            If Me.TxtRefNo.Text = "" Then
                Me.btnAsset_Find.PerformClick()
            End If

            If Me.TxtRefNo.Text <> "" Then
                If Me.ObjJurnal_Asset.ApakahBolehEdit(Me.ObjJurnal_Asset.Primary, UserCode, True, True) = True Then
                    Me.Baru = False
                    Me.ObjJurnal_Asset.CariDetail_DepresiasiAsset(Me.ObjJurnal_Asset.Primary, Me.dgvAsset_Item, Me, TipeJournal_Asset, Me.dgvAsset_ItemAsset)
                    Me.SalinkeField_Asset()
                    Me.AktifkanDetail_Asset()
                    Me.dgvAsset_Item.AllowUserToAddRows = True
                    Me.TxtRefNo.Focus()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAsset_Delete.Click
        Dim hasil As MsgBoxResult
        Try
            Me.dgvAsset_Item.AllowUserToAddRows = False
            If Me.TxtRefNo.Text = "" Then
                Me.btnAsset_Find.PerformClick()
            End If

            If Me.TxtRefNo.Text <> "" Then
                If Me.ObjJurnal_Asset.ApakahBolehHapus(Me.TxtAsset_PrimaryJournal.Text, True, UserCode, False, True) = False Then
                    Me.btnAsset_Find.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjJurnal_Asset.Hapus(Me.TxtAsset_PrimaryJournal.Text, True, UserCode, False) = True Then
                            Me.KosongkanField_Asset()
                            Me.btnAsset_New.Focus()
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

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAsset_Exit.Click
        Me.Close()
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAsset_OK.Click
        Dim hasil As MsgBoxResult
        Try
            Me.dgvAsset_Item.AllowUserToAddRows = False
            Me.SalinkeObjek_Asset()
            If Baru = True Then
                If ObjJurnal_Asset.ApakahDataBenar1(Baru, Me.dgvAsset_Item, Me.ColAsset_PrimaryCurrency.Index, Me.dtpAsset_Date.Value) = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If Me.ApakahDataBenar_Asset() = True Then
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjJurnal_Asset.SimpanBaru(Me.dgvAsset_Item, Me.UserCode, Me, Me.dgvAsset_ItemAsset) Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinkeField_Asset()
                                Me.NonAktifkanDetail_Asset()
                                Me.btnAsset_Find.Focus()
                            End If
                            Me.dgvAsset_Item.AllowUserToAddRows = False
                        ElseIf hasil = MsgBoxResult.No Then
                            Me.dgvAsset_Item.AllowUserToAddRows = True
                        End If
                    Else
                        Me.dgvAsset_Item.AllowUserToAddRows = True
                    End If
                Else
                    Me.dgvAsset_Item.AllowUserToAddRows = True
                End If
            Else
                'konfirmasi disimpan atau tidak
                If Me.ApakahDataBenar_Asset = True Then
                    If ObjJurnal_Asset.ApakahBolehEdit(Me.TxtAsset_PrimaryJournal.Text, UserCode, True, True) = True Then
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjJurnal_Asset.SimpanPerubahan(Me.dgvAsset_Item, Me.UserCode, Me, Me.dgvAsset_ItemAsset, TipeJournal_Asset) Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinkeField_Asset()
                                Me.NonAktifkanDetail_Asset()
                                Me.btnAsset_Edit.Focus()
                            End If
                            Me.dgvAsset_Item.AllowUserToAddRows = False
                        ElseIf hasil = MsgBoxResult.No Then
                            Me.dgvAsset_Item.AllowUserToAddRows = True
                        End If
                    End If
                Else
                    Me.dgvAsset_Item.AllowUserToAddRows = True
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAsset_Cancel.Click
        Dim hasil As MsgBoxResult

        Try
            Me.dgvAsset_Item.AllowUserToAddRows = False
            If Me.TxtAsset_Remarks.Text = "" And Me.dgvAsset_Item.RowCount = 0 And _
            Me.TxtRefNo.Text = "" Then
                Me.NonAktifkanDetail_Asset()
                Me.btnAsset_New.Focus()
                Me.KosongkanField_Asset()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    ''Me.KosongkanItem()
                    If Baru = True Then
                        Me.KosongkanField_Asset()
                    Else
                        Me.ObjJurnal_Asset.CariDetail_DepresiasiAsset(Me.TxtAsset_PrimaryJournal.Text, Me.dgvAsset_Item, _
                        Me, TipeJournal_Asset, Me.dgvAsset_ItemAsset)
                        Me.SalinkeField_Asset()
                    End If
                Else
                    Me.dgvAsset_Item.AllowUserToAddRows = True
                    Exit Sub
                End If
                Me.NonAktifkanDetail_Asset()
                Me.btnAsset_New.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPO_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Me.dgvAsset_Item.AllowUserToAddRows = False
            If Me.TxtAsset_Remarks.Enabled = True Then
                Me.BringToFront()
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar_Asset() = True Then
                        Me.SalinkeObjek_Asset()
                        If Baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjJurnal_Asset.SimpanBaru(Me.dgvAsset_Item, Me.UserCode, Me, Me.dgvAsset_ItemAsset) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinkeField_Asset()
                                Me.NonAktifkanDetail_Asset()
                                Me.btnAsset_Find.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjJurnal_Asset.SimpanPerubahan(Me.dgvAsset_Item, Me.UserCode, Me, Me.dgvAsset_ItemAsset, TipeJournal_Asset) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinkeField_Asset()
                                Me.NonAktifkanDetail_Asset()
                                Me.btnAsset_Edit.Focus()
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
            Me.dgvAsset_Item.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''Fungsi DGV
    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvAsset_Item.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColAsset_Account.Index Then
                If e.ColumnIndex <> Me.ColAsset_Remarks.Index And e.ColumnIndex <> Me.ColAsset_OriginalCurrency.Index And e.ColumnIndex <> Me.ColAsset_DescAccount.Index Then
                    If e.ColumnIndex = Me.ColAsset_Debet.Index Or e.ColumnIndex = Me.ColAsset_Kredit.Index Or _
                    e.ColumnIndex = Me.ColAsset_Debet_OriginalCurr.Index Or e.ColumnIndex = Me.ColAsset_Credit_OriginalCurrency.Index Or _
                    e.ColumnIndex = Me.ColAsset_ExchangeRate.Index Or e.ColumnIndex = Me.ColAsset_Job.Index Or _
                    e.ColumnIndex = Me.ColAsset_Dept.Index Then
                        'jika belum diisikan account codenya, maka tidak bisa edit dept dan job
                        If e.ColumnIndex = Me.ColAsset_Dept.Index Or e.ColumnIndex = Me.ColAsset_Job.Index Then
                            If Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryAkun.Index, e.RowIndex).Value = -1 Then
                                e.Cancel = True
                                Exit Sub
                            End If
                        End If

                        'untuk kolom exchange rate, debet original, creadit original tidak bisa diedit untuk akun dengan mata uang idr
                        If Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryCurrency.Index, e.RowIndex).Value = Me.PrimaryBaseCurrency Then
                            If e.ColumnIndex = Me.ColAsset_ExchangeRate.Index Or _
                            e.ColumnIndex = Me.ColAsset_Debet_OriginalCurr.Index Or _
                            e.ColumnIndex = Me.ColAsset_Credit_OriginalCurrency.Index Then
                                e.Cancel = True
                                Exit Sub
                            End If
                        End If

                        If IsNothing(Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryAkun.Index, e.RowIndex).Value) = True Then
                            'bukan baris teakhir
                            e.Cancel = True
                            Exit Sub
                        ElseIf Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader2(Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryAkun.Index, e.RowIndex).Value) = False Then
                            'ada akunnya, aktif, dan bukan header
                            e.Cancel = True
                            Exit Sub
                        Else
                            Me.Buffer = Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value
                            'buang nol dibelakang decimal
                            Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value.ToString)
                        End If
                    End If
                Else
                    If Me.dgvAsset_Item.Rows.Count - 1 = e.RowIndex Then
                        e.Cancel = True
                        Exit Sub
                    Else
                        If e.ColumnIndex = Me.ColAsset_Remarks.Index Then
                            Me.Buffer = Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value
                        Else
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If
                End If
            Else
                'col account
                If e.ColumnIndex = Me.ColAsset_Account.Index Then
                    Me.Buffer = Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvAsset_Item.CellEndEdit
        Dim Rounding As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim ExRate As Decimal = 1
        Dim DefPrecOriginalCurr As Integer = 0

        Try
            Rounding = Me.DefaultPrecision_Asset
            If IsNothing(Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value) = True _
            Or Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColAsset_Remarks.Index Then
                    'Remarks
                    Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = ""
                ElseIf e.ColumnIndex = Me.ColAsset_Account.Index Then
                    'kalo kode akun dihapus, maka hapus satu bari semuanya
                    If e.RowIndex <> Me.dgvAsset_Item.RowCount - 1 Then
                        Me.dgvAsset_Item.Rows.RemoveAt(e.RowIndex)
                    End If
                    Exit Sub
                ElseIf e.ColumnIndex = Me.ColAsset_Job.Index Then
                    'hapuskan kode job
                    Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryJob.Index, e.RowIndex).Value = -1
                ElseIf e.ColumnIndex = Me.ColAsset_Dept.Index Then
                    'hapuskan kode dept
                    Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryDept.Index, e.RowIndex).Value = -1
                Else
                    'untuk kolom debet kredit, langsung di balikin ke nilai semula
                    Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                End If
            Else
                ''cek debet
                If e.ColumnIndex = Me.ColAsset_Debet.Index Then
                    If Me.ObjCD.CekNonNegatif(Me.dgvAsset_Item.Item(e.ColumnIndex, _
                            e.RowIndex).Value) = False Then
                        MsgBox("Please enter positive number", MsgBoxStyle.Information, _
                        "Information")
                        Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value, Rounding)
                    ''Cek kredit
                    If Me.dgvAsset_Item.Item(Me.ColAsset_Kredit.Index, e.RowIndex).Value <> 0 Then
                        Me.dgvAsset_Item.Item(Me.ColAsset_Kredit.Index, e.RowIndex).Value = ObjCD.DigitGrouping(0, Rounding)
                    End If
                    Me.BalanceDebetKredit_Asset(e.RowIndex, True, Rounding)
                    Me.HitungOriginalValues_Asset(e.RowIndex)
                End If
                ''Cek kredit 
                If e.ColumnIndex = Me.ColAsset_Kredit.Index Then
                    If Me.ObjCD.CekNonNegatif(Me.dgvAsset_Item.Item(e.ColumnIndex, _
                            e.RowIndex).Value) = False Then
                        MsgBox("Please enter positive number", MsgBoxStyle.Information, _
                        "Information")
                        Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value, Rounding)
                    ''Cek debet
                    If Me.dgvAsset_Item.Item(Me.ColAsset_Debet.Index, e.RowIndex).Value <> 0 Then
                        Me.dgvAsset_Item.Item(Me.ColAsset_Debet.Index, e.RowIndex).Value = ObjCD.DigitGrouping(0, Rounding)
                    End If
                    Me.BalanceDebetKredit_Asset(e.RowIndex, False, Rounding)
                    Me.HitungOriginalValues_Asset(e.RowIndex)
                End If

                ''data account number
                If e.ColumnIndex = Me.ColAsset_Account.Index Then
                    If e.ColumnIndex <> Me.ColAsset_Debet.Index Then
                        If e.ColumnIndex <> Me.ColAsset_Kredit.Index Then
                            Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                            'kemungkinan
                            If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                                Me.dgvAsset_Item.Item(Me.ColAsset_DescAccount.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription(Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryAkun.Index, e.RowIndex).Value = Me.ObjAkun.GetPrimary(Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryCurrency.Index, e.RowIndex).Value = Me.ObjAkun.GetPrimaryCurrency(Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryAkun.Index, e.RowIndex).Value)
                                Me.dgvAsset_Item.Item(Me.ColAsset_OriginalCurrency.Index, e.RowIndex).Value = Me.ObjCur.GetCurrencyCode(Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryCurrency.Index, e.RowIndex).Value)
                                ExRate = Me.ObjTaxRate.GetTaxRate(Me.dtpAsset_Date.Value, Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryCurrency.Index, e.RowIndex).Value)
                                If ExRate = 0 Then
                                    ExRate = 1
                                End If
                                Me.dgvAsset_Item.Item(Me.ColAsset_ExchangeRate.Index, e.RowIndex).Value = ObjCD.DigitGrouping(ExRate, DefPrecExchangeRate)
                                'DefPrecOriginalCurr = Me.ObjCur.GetDefaultPrecission(Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryCurrency.Index, e.RowIndex).Value)
                                Me.HitungOriginalValues_Asset(e.RowIndex)
                            Else
                                '2. kode akun yang diketik tidak ada --> tampilkan form cari
                                Dim ij As New FrmMasterAkun_Cari
                                ij.txtCari.Text = Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value
                                ij.ShowDialog()
                                If ij.mPrimaryAkun <> -1 Then
                                    '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                                    Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAkun.GetAccountCode(ij.mPrimaryAkun)
                                    Me.dgvAsset_Item.Item(Me.ColAsset_DescAccount.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription(Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value)
                                    Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryAkun.Index, e.RowIndex).Value = ij.mPrimaryAkun
                                    Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryCurrency.Index, e.RowIndex).Value = Me.ObjAkun.GetPrimaryCurrency(Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryAkun.Index, e.RowIndex).Value)
                                    Me.dgvAsset_Item.Item(Me.ColAsset_OriginalCurrency.Index, e.RowIndex).Value = Me.ObjCur.GetCurrencyCode(Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryCurrency.Index, e.RowIndex).Value)
                                    ExRate = Me.ObjTaxRate.GetTaxRate(Me.dtpAsset_Date.Value, Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryCurrency.Index, e.RowIndex).Value)
                                    If ExRate = 0 Then
                                        ExRate = 1
                                    End If
                                    Me.dgvAsset_Item.Item(Me.ColAsset_ExchangeRate.Index, e.RowIndex).Value = ObjCD.DigitGrouping(ExRate, DefPrecExchangeRate)
                                    'DefPrecOriginalCurr = Me.ObjCur.GetDefaultPrecission(Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryCurrency.Index, e.RowIndex).Value)
                                    Me.HitungOriginalValues_Asset(e.RowIndex)
                                ElseIf ij.mPrimaryAkun = -1 Then
                                    '2.b. kode akun yang dicari tidak ada --> kembalikan ke kondisi buffer
                                    Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    If Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                        Me.dgvAsset_Item.Rows.RemoveAt(e.RowIndex)
                                        Exit Sub
                                    End If
                                End If
                                ij = Nothing
                            End If
                        End If
                    End If
                    Me.SetMulticurrency_Asset()
                End If

                '3. exchange rate, hanya jika non base currency
                If e.ColumnIndex = Me.ColAsset_ExchangeRate.Index Then
                    If Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryCurrency.Index, e.RowIndex).Value <> Me.PrimaryBaseCurrency Then
                        If Me.ObjCD.CekBilanganPositif(Me.dgvAsset_Item.Item(e.ColumnIndex, _
                                                   e.RowIndex).Value) = False Then
                            MsgBox("Please enter positive number", MsgBoxStyle.Information, _
                            "Information")
                            Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        'digit groping
                        Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = _
                        ObjCD.DigitGrouping(Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value, DefPrecExchangeRate)
                        'hitung original value
                        Me.HitungOriginalValues_Asset(e.RowIndex)
                    End If
                End If

                '3. debet original currency
                If e.ColumnIndex = Me.ColAsset_Debet_OriginalCurr.Index Then
                    If Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryCurrency.Index, e.RowIndex).Value <> Me.PrimaryBaseCurrency Then
                        If Me.ObjCD.CekNonNegatif(Me.dgvAsset_Item.Item(e.ColumnIndex, _
                                                   e.RowIndex).Value) = False Then
                            MsgBox("Please enter positive number", MsgBoxStyle.Information, _
                            "Information")
                            Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        'digit groping
                        DefPrecOriginalCurr = Me.ObjCur.GetDefaultPrecission(Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryCurrency.Index, e.RowIndex).Value)
                        Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = _
                        ObjCD.DigitGrouping(Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value, _
                        DefPrecOriginalCurr)
                        'kosongkan credit
                        Me.dgvAsset_Item.Item(Me.ColAsset_Credit_OriginalCurrency.Index, e.RowIndex).Value = ObjCD.DigitGrouping(0, DefPrecOriginalCurr)

                        Me.HitungBaseValues_Asset(e.RowIndex)

                        If Me.dgvAsset_Item.Item(Me.ColAsset_Debet.Index, e.RowIndex).Value > 0 Then
                            Me.BalanceDebetKredit_Asset(e.RowIndex, True, Rounding, True)
                        Else
                            Me.BalanceDebetKredit_Asset(e.RowIndex, False, Rounding, True)
                        End If
                    End If
                End If

                '4. credit original currency
                If e.ColumnIndex = Me.ColAsset_Credit_OriginalCurrency.Index Then
                    If Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryCurrency.Index, e.RowIndex).Value <> Me.PrimaryBaseCurrency Then
                        If Me.ObjCD.CekNonNegatif(Me.dgvAsset_Item.Item(e.ColumnIndex, _
                                                   e.RowIndex).Value) = False Then
                            MsgBox("Please enter positive number", MsgBoxStyle.Information, _
                            "Information")
                            Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        'digit groping
                        DefPrecOriginalCurr = Me.ObjCur.GetDefaultPrecission(Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryCurrency.Index, e.RowIndex).Value)
                        Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = _
                        ObjCD.DigitGrouping(Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value, _
                        DefPrecOriginalCurr)
                        'kosongkan credit
                        Me.dgvAsset_Item.Item(Me.ColAsset_Debet_OriginalCurr.Index, e.RowIndex).Value = ObjCD.DigitGrouping(0, DefPrecOriginalCurr)

                        Me.HitungBaseValues_Asset(e.RowIndex)

                        If Me.dgvAsset_Item.Item(Me.ColAsset_Debet.Index, e.RowIndex).Value > 0 Then
                            Me.BalanceDebetKredit_Asset(e.RowIndex, True, Rounding, True)
                        Else
                            Me.BalanceDebetKredit_Asset(e.RowIndex, False, Rounding, True)
                        End If
                    End If
                End If

                '5. Dept
                If e.ColumnIndex = Me.ColAsset_Dept.Index Then
                    Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    If Me.ObjDept.ApakahSudahAda(Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryDept.Index, e.RowIndex).Value = Me.ObjDept.GetPrimary(Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value)
                    Else
                        Dim ij As New Core_Module.FrmMasterDepartementCari
                        ij.txtCari.Text = Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.PrimaryDepartement <> -1 Then
                            Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjDept.GetDeptCode(ij.PrimaryDepartement)
                            Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryDept.Index, e.RowIndex).Value = ij.PrimaryDepartement
                        ElseIf ij.PrimaryDepartement = -1 Then
                            Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryDept.Index, e.RowIndex).Value = -1
                                Exit Sub
                            End If
                        End If
                        ij = Nothing
                    End If
                End If

                '6. Job
                If e.ColumnIndex = Me.ColAsset_Job.Index Then
                    Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    If Me.ObjJob.ApakahSudahAda(Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryJob.Index, e.RowIndex).Value = Me.ObjJob.GetPrimary(Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value)
                    Else
                        Dim ij As New Core_Module.FrmMasterJobCari
                        ij.txtCari.Text = Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.PrimaryJob <> -1 Then
                            Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjJob.GetJobCode(ij.PrimaryJob)
                            Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryJob.Index, e.RowIndex).Value = ij.PrimaryJob
                        ElseIf ij.PrimaryJob = -1 Then
                            Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                Me.dgvAsset_Item.Item(Me.ColAsset_PrimaryJob.Index, e.RowIndex).Value = -1
                                Exit Sub
                            End If
                        End If
                        ij = Nothing
                    End If
                End If

                '3 cek apakah ada kode akun yang sama pada referensi kode yang dibuat 
                'If Me.dgvAsset_Item.Rows.Count > 2 Then
                'Me.dgvAsset_Item.AllowUserToAddRows = False
                'Dim item As String = ""
                'For i As Integer = 0 To dgvAsset_Item.Rows.Count - 1
                'If e.RowIndex <> i Then
                'item = CStr(Me.dgvAsset_Item.Rows(i).Cells(Me.ColAsset_Account.Index).Value)
                'If Me.dgvAsset_Item.Item(e.ColumnIndex, e.RowIndex).Value = item Then
                'MsgBox("Duplication account number at data item, please change your number", _
                'MsgBoxStyle.Information, "Information")
                'Me.dgvAsset_Item.Rows.RemoveAt(e.RowIndex)
                'Me.dgvAsset_Item.AllowUserToAddRows = True
                'Exit Sub
                'End If
                'End If
                'Next
                'Me.dgvAsset_Item.AllowUserToAddRows = True
                'End If
            End If

            ''hitung debet dan kredit
            Me.HitungUnbalance_Asset()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCur = Nothing
        End Try
    End Sub

    Private Function HitungUnbalance_Asset(ByVal vdebet As Decimal, ByVal vkredit As Decimal) As Decimal
        Dim Unbalance As Decimal = 0
        Try
            Unbalance = Math.Round(vdebet - vkredit, Me.DefaultPrecision_Asset, MidpointRounding.AwayFromZero)
            Return Unbalance
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub HitungUnbalance_Asset()
        Dim row As DataGridViewRow
        Try
            ''hitung debet dan kredit
            Dim TotalDebet As Decimal = 0
            Dim TotalKredit As Decimal = 0
            For Each row In dgvAsset_Item.Rows
                TotalDebet += CDec(Me.dgvAsset_Item.Item(Me.ColAsset_Debet.Index, row.Index).Value)
                TotalKredit += CDec(Me.dgvAsset_Item.Item(Me.ColAsset_Kredit.Index, row.Index).Value)
            Next
            Me.TxtAsset_Unbalance.Text = CStr(ObjCD.DigitGrouping(Me.HitungUnbalance_Asset(TotalDebet, TotalKredit), Me.DefaultPrecision_Asset))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BalanceDebetKredit_Asset(ByVal vIndexAcuan As Integer, ByVal vAcuanDebet As Boolean, ByVal rounding As Integer, _
   Optional ByVal vHitungOriginalCurrency As Boolean = True)

        Try
            Dim rowAcuan As DataGridViewRow = Me.dgvAsset_Item.Rows(vIndexAcuan)
            Dim rowTarget As DataGridViewRow

            If Me.TxtAsset_Remarks.Enabled = True And Me.dgvAsset_Item.Rows.Count = 3 Then
                'dibalance jika hanya ada 2 jurnal, maka debet = kredit
                If rowAcuan.Index = 0 Then
                    rowTarget = Me.dgvAsset_Item.Rows(1)
                Else
                    rowTarget = Me.dgvAsset_Item.Rows(0)
                End If

                If vAcuanDebet = True Then
                    rowTarget.Cells(Me.ColAsset_Kredit.Index).Value = rowAcuan.Cells(Me.ColAsset_Debet.Index).Value
                    rowTarget.Cells(Me.ColAsset_Debet.Index).Value = Me.ObjCD.DigitGrouping(0, rounding)
                Else
                    rowTarget.Cells(Me.ColAsset_Debet.Index).Value = rowAcuan.Cells(Me.ColAsset_Kredit.Index).Value
                    rowTarget.Cells(Me.ColAsset_Kredit.Index).Value = Me.ObjCD.DigitGrouping(0, rounding)
                End If

                If vHitungOriginalCurrency = True Then
                    Me.HitungOriginalValues_Asset(rowTarget.Index)
                End If
            ElseIf Me.TxtAsset_Remarks.Enabled = True And Me.dgvAsset_Item.Rows.Count > 3 Then
                'dibalance jika ada 3 jurnal, jika 2 dikredit atau 2 didebet
                If Me.GetJmlNol_Asset = 0 Then
                    If vAcuanDebet = True Then
                        If Me.GetJmlKredit_Asset = 1 Then
                            rowTarget = Me.dgvAsset_Item.Rows(Me.GetIndexJmlKredit_Asset)
                            rowTarget.Cells(Me.ColAsset_Kredit.Index).Value = Me.ObjCD.DigitGrouping(Me.GetTotalDebet_Asset, rounding)
                            rowTarget.Cells(Me.ColAsset_Debet.Index).Value = Me.ObjCD.DigitGrouping(0, rounding)
                        End If
                    Else
                        If Me.GetJmlDebet_Asset = 1 Then
                            rowTarget = Me.dgvAsset_Item.Rows(Me.GetIndexJmlDebet_Asset)
                            rowTarget.Cells(Me.ColAsset_Debet.Index).Value = Me.ObjCD.DigitGrouping(Me.GetTotalKredit_Asset, rounding)
                            rowTarget.Cells(Me.ColAsset_Kredit.Index).Value = Me.ObjCD.DigitGrouping(0, rounding)
                        End If
                    End If

                    If vHitungOriginalCurrency = True Then
                        If IsNothing(rowTarget) = False Then
                            Me.HitungOriginalValues_Asset(rowTarget.Index)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetTotalDebet_Asset() As Decimal
        Dim row As DataGridViewRow
        Dim retval As Decimal = 0

        Try
            If Me.dgvAsset_Item.Rows.Count > 1 Then
                For i As Integer = 0 To Me.dgvAsset_Item.Rows.Count - 2
                    row = Me.dgvAsset_Item.Rows(i)
                    If row.Cells(Me.ColAsset_Debet.Index).Value <> 0 Then
                        retval += CDec(row.Cells(Me.ColAsset_Debet.Index).Value)
                    End If
                Next
            End If

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDebet_Asset() As Integer
        Dim row As DataGridViewRow
        Dim retval As Integer = 0

        Try
            If Me.dgvAsset_Item.Rows.Count > 1 Then
                For i As Integer = 0 To Me.dgvAsset_Item.Rows.Count - 2
                    row = Me.dgvAsset_Item.Rows(i)
                    If row.Cells(Me.ColAsset_Debet.Index).Value <> 0 Then
                        retval += 1
                    End If
                Next
            End If

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetIndexJmlDebet_Asset() As Integer
        Dim row As DataGridViewRow

        Try
            If Me.dgvAsset_Item.Rows.Count > 1 Then
                For i As Integer = 0 To Me.dgvAsset_Item.Rows.Count - 2
                    row = Me.dgvAsset_Item.Rows(i)
                    If row.Cells(Me.ColAsset_Debet.Index).Value <> 0 Then
                        Return row.Index
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetTotalKredit_Asset() As Decimal
        Dim row As DataGridViewRow
        Dim retval As Decimal = 0

        Try
            If Me.dgvAsset_Item.Rows.Count > 1 Then
                For i As Integer = 0 To Me.dgvAsset_Item.Rows.Count - 2
                    row = Me.dgvAsset_Item.Rows(i)
                    If row.Cells(Me.ColAsset_Kredit.Index).Value <> 0 Then
                        retval += CDec(row.Cells(Me.ColAsset_Kredit.Index).Value)
                    End If
                Next
            End If

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlKredit_Asset() As Integer
        Dim row As DataGridViewRow
        Dim retval As Integer = 0

        Try
            If Me.dgvAsset_Item.Rows.Count > 1 Then
                For i As Integer = 0 To Me.dgvAsset_Item.Rows.Count - 2
                    row = Me.dgvAsset_Item.Rows(i)
                    If row.Cells(Me.ColAsset_Kredit.Index).Value <> 0 Then
                        retval += 1
                    End If
                Next
            End If

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetIndexJmlKredit_Asset() As Integer
        Dim row As DataGridViewRow

        Try
            If Me.dgvAsset_Item.Rows.Count > 1 Then
                For i As Integer = 0 To Me.dgvAsset_Item.Rows.Count - 2
                    row = Me.dgvAsset_Item.Rows(i)
                    If row.Cells(Me.ColAsset_Kredit.Index).Value <> 0 Then
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
    Private Function GetJmlNol_Asset() As Integer
        Dim row As DataGridViewRow
        Dim retval As Integer = 0

        Try
            If Me.dgvAsset_Item.Rows.Count > 1 Then
                For i As Integer = 0 To Me.dgvAsset_Item.Rows.Count - 2
                    row = Me.dgvAsset_Item.Rows(i)
                    If row.Cells(Me.ColAsset_Debet.Index).Value = 0 And row.Cells(Me.ColAsset_Kredit.Index).Value = 0 Then
                        retval += 1
                    End If
                Next
            End If

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SetMulticurrency_Asset()
        Dim rowDGV As DataGridViewRow
        Dim Multicurrency As Boolean = False

        Try
            For Each rowDGV In Me.dgvAsset_Item.Rows
                If rowDGV.Index = Me.dgvAsset_Item.RowCount - 1 Then
                    If Me.dgvAsset_Item.AllowUserToAddRows = True Then
                        'tidak dicek
                    Else
                        'dicek
                        If rowDGV.Cells(Me.ColAsset_PrimaryCurrency.Index).Value <> Me.PrimaryBaseCurrency Then
                            Multicurrency = True
                            Exit For
                        End If
                    End If
                Else
                    If rowDGV.Cells(Me.ColAsset_PrimaryCurrency.Index).Value <> Me.PrimaryBaseCurrency Then
                        Multicurrency = True
                        Exit For
                    End If
                End If
            Next

            If Multicurrency = True Then
                Me.ColAsset_ExchangeRate.Visible = True
                Me.ColAsset_Debet_OriginalCurr.Visible = True
                Me.ColAsset_Credit_OriginalCurrency.Visible = True
                Me.ColAsset_OriginalCurrency.Visible = True
            Else
                Me.ColAsset_OriginalCurrency.Visible = False
                Me.ColAsset_ExchangeRate.Visible = False
                Me.ColAsset_Debet_OriginalCurr.Visible = False
                Me.ColAsset_Credit_OriginalCurrency.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAsset_Print.Click
        Try
            If Me.TxtRefNo.Text = "" Then
                Me.btnAsset_Find.PerformClick()
            End If

            If Me.TxtRefNo.Text <> "" Then
                Dim i As New FrmTransaksiJournal_Cetak1
                i.PrimaryJurnal = Me.TxtAsset_PrimaryJournal.Text
                i.ShowDialog()
                i.Dispose()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungExchangeRate_Asset(ByVal vRowIndex As Integer)
        Dim vRowDGV As DataGridViewRow

        Try
            vRowDGV = Me.dgvAsset_Item.Rows(vRowIndex)
            vRowDGV.Cells(Me.ColAsset_ExchangeRate.Index).Value = ObjCD.DigitGrouping(Math.Round(CDec(vRowDGV.Cells(Me.ColAsset_Debet.Index).Value) / _
            CDec(vRowDGV.Cells(Me.ColAsset_Debet_OriginalCurr.Index).Value), DefPrecExchangeRate, MidpointRounding.AwayFromZero))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungBaseValues_Asset(ByVal vRowIndex As Integer)
        Dim vRowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0
        Dim DefPrecOriginal As Integer = 0

        Try
            vRowDGV = Me.dgvAsset_Item.Rows(vRowIndex)
            DefPrecBase = Me.ObjCur.GetDefaultPrecission(Me.PrimaryBaseCurrency)
            DefPrecOriginal = Me.ObjCur.GetDefaultPrecission(vRowDGV.Cells(Me.ColAsset_PrimaryCurrency.Index).Value)

            If CDec(vRowDGV.Cells(Me.ColAsset_Debet_OriginalCurr.Index).Value) > 0 Then
                vRowDGV.Cells(Me.ColAsset_Kredit.Index).Value = ObjCD.DigitGrouping(0, DefPrecBase)
                vRowDGV.Cells(Me.ColAsset_Debet.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColAsset_Debet_OriginalCurr.Index).Value) * _
                CDec(vRowDGV.Cells(Me.ColAsset_ExchangeRate.Index).Value), DefPrecBase)
            Else
                vRowDGV.Cells(Me.ColAsset_Kredit.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColAsset_Credit_OriginalCurrency.Index).Value) * _
                CDec(vRowDGV.Cells(Me.ColAsset_ExchangeRate.Index).Value), DefPrecBase)
                vRowDGV.Cells(Me.ColAsset_Debet.Index).Value = ObjCD.DigitGrouping(0, DefPrecBase)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungOriginalValues_Asset(ByVal vRowIndex As Integer)
        Dim vRowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0
        Dim DefPrecOriginal As Integer = 0

        Try
            vRowDGV = Me.dgvAsset_Item.Rows(vRowIndex)
            DefPrecBase = Me.ObjCur.GetDefaultPrecission(Me.PrimaryBaseCurrency)
            DefPrecOriginal = Me.ObjCur.GetDefaultPrecission(vRowDGV.Cells(Me.ColAsset_PrimaryCurrency.Index).Value)

            If CDec(vRowDGV.Cells(Me.ColAsset_Debet.Index).Value) > 0 Then
                vRowDGV.Cells(Me.ColAsset_Credit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(0, DefPrecOriginal)
                vRowDGV.Cells(Me.ColAsset_Debet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColAsset_Debet.Index).Value) / _
                CDec(vRowDGV.Cells(Me.ColAsset_ExchangeRate.Index).Value), DefPrecOriginal)
            Else
                vRowDGV.Cells(Me.ColAsset_Credit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColAsset_Kredit.Index).Value) / _
                CDec(vRowDGV.Cells(Me.ColAsset_ExchangeRate.Index).Value), DefPrecOriginal)
                vRowDGV.Cells(Me.ColAsset_Debet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(0, DefPrecOriginal)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class