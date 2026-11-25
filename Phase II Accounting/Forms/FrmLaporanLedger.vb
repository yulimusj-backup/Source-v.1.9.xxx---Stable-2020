Option Explicit On
Option Strict Off

Public Class FrmLaporanLedger
    Public BackColorAll As Color = Color.LightSeaGreen
    Public UserCode As String = ""

    Dim ObjAcc As New ClsMasterAkun
    Dim Obj_Closing As New ClsTransaksiClosing
    Dim DS As New DseLaporanLedger
    Dim DTA As New DseLaporanLedgerTableAdapters.Ledger_ViewTableAdapter
    Dim DTA_With_Reverse As New DseLaporanLedgerTableAdapters.Ledger_View_With_ReverseTableAdapter
    Dim report As New CrpLaporanLedger
    Dim report2 As New CrpLaporanLedger_Multi
    Dim report2_1 As New CrpLaporanLedger_MultiNP
    Dim report3 As New CrpLaporanLedger_Multi2
    Dim report3_1 As New CrpLaporanLedger_Multi2NP
    Dim report1 As New CrpLaporanLedger_NonIDR
    Dim report_reverse As New CrpLaporanLedgerWithReverse
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim Negatif As Integer = 0
    Dim ObjHakAkses As New Core_Module.ClsHakAkses
    Dim ObjJournal As New ClsTransaksiJournal
    Dim TBLAkun As New DseMasterAkun.tblmaster_akun7DataTable
    Dim ObjLicense As New Core_Module.ClsLicense

    Dim EraseRowIndex As Integer = 0
    Dim EraseRowIndexAcc As Integer = 0
    Dim ObjUser As New Core_Module.ClsUser

    Dim buffer As String = ""

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Dim MinDate As Date = Today
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.dtpFrom.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpFrom.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpTo.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpTo.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            '@@@@@ahdar
            If Me.dgvAccount.Rows.Count <= 1 And Me.clbTipeJournal.CheckedItems.Count = 0 Then
                MsgBox("Please select at least one item for account or journal type.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            ''''''

            ''rosadin-ac1
            If Not Me.chkShowItemRemarks.Checked And Not Me.ChkShowHeaderDesc.Checked Then
                MsgBox("Please select one, header description, item remarks, or both that will be displayed as description.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            MinDate = Me.Obj_Closing.GetConvertionDate
            If Me.dtpFrom.Value.Date < MinDate.Date Then
                MsgBox("Minimum data is : " + _
                Format(MinDate, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If Me.dtpFrom.Value.Date > Me.dtpTo.Value.Date Then
                MsgBox("Starting period must be smaller than ending period", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'If Me.chkAccountFilter.Checked = False And Me.TxtPrimaryAcc.Text = "-1" Then
            '    If MsgBox("Show ledger for all account? This may a take a while to perform.", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
            '        Exit Sub
            '    End If
            'ElseIf Me.chkAccountFilter.Checked = True And Me.clbAccount.CheckedItems.Count = 0 Then
            '    MsgBox("Please choose at least 1 account to show ledger.", MsgBoxStyle.Information, "Information")
            '    Exit Sub
            'End If
            If Me.dgvAccount.Rows.Count <= 1 And Me.clbTipeJournal.CheckedItems.Count = Me.clbTipeJournal.Items.Count Then
                If Me.dgvAccount.Rows.Count - 1 = 0 Then
                    If MsgBox("Show ledger for all account? This may a take a while to perform.", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        Exit Sub
                    End If
                End If
            End If

            If Me.chkShowReverseAccount.Checked = True Then
                Me.CreatedLedger_WithReverse()
            Else
                Me.CreatedLedger()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindCodeNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindAcc.Click
        Dim i As New FrmMasterAkun_Cari

        Try
            'cek apakah sudah isi kode currency
            i.BackColorAll = Me.BackColorAll

            ''i.cboCriteria.SelectedIndex = 0
            If Me.TxtCodeAcc.Text <> "" Then
                i.txtCari.Text = Me.TxtCodeAcc.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAcc.Text = i.mPrimaryAkun
                Me.TxtCodeAcc.Text = Me.ObjAcc.GetAccountCode(Me.TxtPrimaryAcc.Text)
                Me.TxtDesc.Text = Me.ObjAcc.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAcc.Text)
                Me.dtpFrom.Focus()
            Else
                Me.TxtCodeAcc.Clear()
                Me.TxtPrimaryAcc.Text = "-1"
                Me.TxtDesc.Clear()
                Me.TxtCodeAcc.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeAcc_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeAcc.KeyDown
        Dim i As New FrmMasterAkun_Cari
        Dim Jml As Long = 0
        i.BackColorAll = Me.BackColorAll
        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjAcc.GetJml_NonHeader_Aktif(Me.TxtCodeAcc.Text)
                If Jml = 1 Then
                    Me.TxtPrimaryAcc.Text = Me.ObjAcc.GetPrimary(Me.TxtCodeAcc.Text)
                    Me.TxtDesc.Text = Me.ObjAcc.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAcc.Text)
                    Me.dtpFrom.Focus()
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.TxtCodeAcc.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeAcc.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryAcc.Text = i.mPrimaryAkun
                        Me.TxtCodeAcc.Text = Me.ObjAcc.GetAccountCode(Me.TxtPrimaryAcc.Text)
                        Me.TxtDesc.Text = Me.ObjAcc.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAcc.Text)
                        Me.dtpFrom.Focus()
                    Else
                        Me.TxtCodeAcc.Clear()
                        Me.TxtPrimaryAcc.Text = "-1"
                        Me.TxtDesc.Clear()
                        Me.TxtCodeAcc.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeAcc_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeAcc.Leave
        Dim i As New FrmMasterAkun_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCodeAcc.Text <> "" Then
                If Me.ObjAcc.GetJml_NonHeader_Aktif(Me.TxtCodeAcc.Text) = 1 Then
                    Me.TxtPrimaryAcc.Text = Me.ObjAcc.GetPrimary(Me.TxtCodeAcc.Text)
                    Me.TxtDesc.Text = Me.ObjAcc.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAcc.Text)
                    Me.dtpFrom.Focus()
                Else
                    If Me.TxtCodeAcc.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeAcc.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryAcc.Text = i.mPrimaryAkun
                        Me.TxtCodeAcc.Text = Me.ObjAcc.GetAccountCode(Me.TxtPrimaryAcc.Text)
                        Me.TxtDesc.Text = Me.ObjAcc.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAcc.Text)
                        Me.dtpFrom.Focus()
                    Else
                        Me.TxtCodeAcc.Clear()
                        Me.TxtPrimaryAcc.Text = "-1"
                        Me.TxtDesc.Clear()
                        Me.TxtCodeAcc.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryAcc.Text = "-1"
                Me.TxtDesc.Clear()
            End If

            If Me.TxtPrimaryAcc.Text <> "" And Me.TxtPrimaryAcc.Text <> "-1" Then
                Me.chkAccountFilter.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeAcc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeAcc.LostFocus
        Try
            Me.TxtCodeAcc.BackColor = Me.TxtPrimaryAcc.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeAcc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeAcc.GotFocus
        Try
            Me.TxtCodeAcc.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CreatedLedger()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim BB As Decimal = 0
        Dim BB1 As Decimal = 0
        Dim row, rowDummy, row_BB, row_BB1, row_BB_Before As DseLaporanLedger.Ledger_ViewRow
        Dim Saldo_Normal As Integer = Me.ObjAcc.GetSaldoNormal(Me.TxtPrimaryAcc.Text)
        Dim Ref_date As Date = Me.Obj_Closing.GetConvertionDate
        Dim TBLDummy As New DseLaporanLedger.Ledger_ViewDataTable
        Dim DV As Data.DataRowView
        Dim DummyAccount As String = ""
        Dim AccountList As New List(Of String)
        Dim i, StartIndex As Integer
        Dim ListPrimaryAkun As New List(Of String)
        Dim ListCreatedBy As New List(Of String)
        Dim rowItem As DseMasterAkun.tblmaster_akun7Row
        Dim ListUser As New List(Of String)
        Dim TipeAccess As ClsTransaksiJournal.EnumJournalAccesType = ClsTransaksiJournal.EnumJournalAccesType.FullAccess

        Try
            'isikan data
            If Me.TxtPrimaryAcc.Text <> "" And Me.TxtPrimaryAcc.Text <> "-1" Then
                'untuk akun non idr, maka gunakan ledger non idr
                'If Me.ObjAcc.GetPrimaryCurrency(Me.TxtPrimaryAcc.Text) <> Core_Module.PrimaryBaseCurrency Then
                'Me.ComboBox1.SelectedIndex = 1
                'End If

                If Me.TxtPrimaryDept.Text = "-1" And Me.TxtPrimaryJob.Text = "-1" Then
                    Me.DTA.FillBy_TanggalSampaiDari(Me.DS.Ledger_View, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date, Me.TxtPrimaryAcc.Text)
                ElseIf Me.TxtPrimaryDept.Text <> "-1" And Me.TxtPrimaryJob.Text = "-1" Then
                    Me.DTA.FillBy_Dept(Me.DS.Ledger_View, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date, Me.TxtPrimaryAcc.Text, Me.TxtPrimaryDept.Text)
                ElseIf Me.TxtPrimaryDept.Text = "-1" And Me.TxtPrimaryJob.Text <> "-1" Then
                    Me.DTA.FillBy_Job(Me.DS.Ledger_View, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date, Me.TxtPrimaryAcc.Text, Me.TxtPrimaryJob.Text)
                Else
                    Me.DTA.FillBy_Dept_Job(Me.DS.Ledger_View, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date, Me.TxtPrimaryAcc.Text, Me.TxtPrimaryDept.Text, _
                    Me.TxtPrimaryJob.Text)
                End If
            Else
                If Me.TxtPrimaryDept.Text = "-1" And Me.TxtPrimaryJob.Text = "-1" Then
                    Me.DTA.FillBy_All_TanggalSampaiDari(Me.DS.Ledger_View, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
                ElseIf Me.TxtPrimaryDept.Text <> "-1" And Me.TxtPrimaryJob.Text = "-1" Then
                    Me.DTA.FillBy_Dept(Me.DS.Ledger_View, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date, Me.TxtPrimaryAcc.Text, Me.TxtPrimaryDept.Text)
                ElseIf Me.TxtPrimaryDept.Text = "-1" And Me.TxtPrimaryJob.Text <> "-1" Then
                    Me.DTA.FillBy_Job(Me.DS.Ledger_View, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date, Me.TxtPrimaryAcc.Text, Me.TxtPrimaryJob.Text)
                Else
                    Me.DTA.FillBy_Dept_Job(Me.DS.Ledger_View, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date, Me.TxtPrimaryAcc.Text, Me.TxtPrimaryDept.Text, _
                    Me.TxtPrimaryJob.Text)
                End If
            End If

            '

            '@ahdar ' optimasi
            Me.DTA_A.Fill(Me.TBL_A, Me.dtpFrom.Value.Date - TimeSpan.FromDays(1))
            Me.DTA_B.Fill(Me.TBL_B, Me.dtpFrom.Value.Date - TimeSpan.FromDays(1))
            '
            ''rosadin-ac1
            ''atur description sesuai dengan pilihan user
            'For Each row_BB In DS.Ledger_View
            '    If Me.ChkShowHeaderDesc.Checked = False And row_BB.Remarks <> "" Then
            '        row_BB.Desc = ""
            '    End If
            '    If Me.chkShowItemRemarks.Checked = True And row_BB.Remarks <> "" Then
            '        If row_BB.Desc = "" Then
            '            row_BB.Desc = row_BB.Remarks
            '        Else
            '            row_BB.Desc = row_BB.Desc + " - " + row_BB.Remarks
            '        End If
            '    End If
            'Next

            For Each row_BB In DS.Ledger_View
                Dim DescBuffer As String = row_BB.Desc
                row_BB.Desc = String.Empty
                If Me.ChkShowHeaderDesc.Checked Then
                    row_BB.Desc = DescBuffer
                End If
                If Me.chkShowItemRemarks.Checked Then
                    If Not String.IsNullOrEmpty(row_BB.Desc) Then
                        If Not String.IsNullOrEmpty(row_BB.Remarks) Then
                            row_BB.Desc = row_BB.Desc & " - " & row_BB.Remarks
                        End If
                    Else
                        row_BB.Desc = row_BB.Remarks
                    End If
                End If
            Next

            'urutkan sesuai dengan saldo normal
            For Each row_BB In DS.Ledger_View
                If row_BB.Saldo_Normal = 0 Then
                    If row_BB.Debet > 0 Then
                        row_BB.InOut = 0
                    Else
                        row_BB.InOut = 1
                    End If
                Else
                    If row_BB.Debet > 0 Then
                        row_BB.InOut = 1
                    Else
                        row_BB.InOut = 0
                    End If
                End If
            Next


            ''''TR//Filter akun diganti jadi DGV 
            ''''buang akun yang tidak dilist
            ''''If Me.dgvAccount.Rows.Count - 1 <> 0 Then
            ''''    'kumpulkan list primary akun
            ''''    For x As Integer = 0 To Me.dgvAccount.Rows.Count - 2
            ''''        If ListPrimaryAkun.Contains(Me.dgvAccount.Item(Me.ColPrimaryAcc.Index, x).Value) = False Then
            ''''            ListPrimaryAkun.Add(Me.dgvAccount.Item(Me.ColPrimaryAcc.Index, x).Value)
            ''''        End If
            ''''    Next

            ''''    'buang baris yang tidak ada dalam primary akun
            ''''    i = 0
            ''''    While i <= Me.DS.Ledger_View.Rows.Count - 1
            ''''        row_BB = DS.Ledger_View.Rows(i)
            ''''        If ListPrimaryAkun.Contains(row_BB.Primary_Akun) = False Then
            ''''            DS.Ledger_View.Rows.Remove(row_BB)
            ''''        Else
            ''''            i += 1
            ''''        End If
            ''''    End While
            ''''End If

            'buang user yang tidak dilist
            If Me.DgvUser.Rows.Count - 1 <> 0 Then
                'kumpulkan list primary user
                For x As Integer = 0 To Me.DgvUser.Rows.Count - 2
                    If ListUser.Contains(Me.DgvUser.Item(Me.ColCreatedBy.Index, x).Value) = False Then
                        ListUser.Add(Me.DgvUser.Item(Me.ColCreatedBy.Index, x).Value)
                    End If
                Next

                'buang baris yang tidak ada dalam created by
                i = 0
                While i <= Me.DS.Ledger_View.Rows.Count - 1
                    row_BB = DS.Ledger_View.Rows(i)
                    If ListUser.Contains(row_BB.Created_By) = False Then
                        DS.Ledger_View.Rows.Remove(row_BB)
                    Else
                        i += 1
                    End If
                End While
            End If

            'aul new
            Dim dtaAkun As New DseLaporanLedgerTableAdapters.tblmaster_akunTableAdapter
            Dim tblAkun As New DseLaporanLedger.tblmaster_akunDataTable
            Dim rowAkun As DseLaporanLedger.tblmaster_akunRow

            dtaAkun.Connection.ConnectionString = Core_Module.GetConstring
            dtaAkun.Fill(tblAkun, 1, 0)

            Dim ListPrimayAkun As New List(Of Double)
            For Each row In Me.DS.Ledger_View
                If ListPrimayAkun.Contains(row.Primary_Akun) = False Then
                    ListPrimayAkun.Add(row.Primary_Akun)
                End If
            Next
            For ijk As Integer = tblAkun.Rows.Count - 1 To 0 Step -1
                rowAkun = tblAkun.Rows(ijk)

                If ListPrimayAkun.Contains(rowAkun.Primary) = False Then
                    row = Me.DS.Ledger_View.NewLedger_ViewRow
                    row.Primary_Akun = rowAkun.Primary
                    row.Account_Code = rowAkun.Account_Code
                    row.Account_Desc = rowAkun.Description
                    row.Journal_No = "Beginning Balance Only"
                    Me.DS.Ledger_View.Rows.Add(row)
                    'tblAkun.Rows.Remove(rowAkun)
                End If
            Next
            '-----------------------------------------------------------

            'buang yang tidak termasuk kedalam list tipe journal yang dipilih
            '***@hdar
            Dim TblC As New DseLaporanLedger.Ledger_ViewDataTable
            Dim RowC As DseLaporanLedger.Ledger_ViewRow
            Dim RowD As DseLaporanLedger.Ledger_ViewRow
            Dim PrimaryAcc As String = ""
            TblC = DS.Ledger_View.Copy
            DS.Ledger_View.Clear()
            For ij As Integer = 0 To Me.dgvAccount.Rows.Count - 1
                If String.IsNullOrEmpty(Me.dgvAccount.Item(Me.ColNameAcc.Index, ij).Value) = False Then
                    If PrimaryAcc = "" Then
                        PrimaryAcc = "[Primary Akun]=" & Me.dgvAccount.Item(Me.ColPrimaryAcc.Index, ij).Value
                    Else
                        PrimaryAcc &= " or " & "[Primary Akun]=" & Me.dgvAccount.Item(Me.ColPrimaryAcc.Index, ij).Value
                    End If
                End If
            Next
            Dim Filter As String = ""
            If PrimaryAcc <> "" And FilterByTipeJournal() <> "" Then
                Filter = "(" & PrimaryAcc & ") and (" & FilterByTipeJournal() & ")"
            ElseIf PrimaryAcc <> "" And FilterByTipeJournal() = "" Then
                Filter = PrimaryAcc
            ElseIf PrimaryAcc = "" And FilterByTipeJournal() <> "" Then
                Filter = FilterByTipeJournal()
            End If

            TblC.DefaultView.RowFilter = Filter
            For ik As Integer = 0 To TblC.DefaultView.Count - 1
                RowC = TblC.DefaultView(ik).Row
                RowD = DS.Ledger_View.NewLedger_ViewRow
                For ik2 As Integer = 0 To TblC.Columns.Count - 1
                    RowD.Item(ik2) = RowC.Item(ik2)
                Next
                DS.Ledger_View.AddLedger_ViewRow(RowD)
            Next
            TblC.Dispose()
            '***

            TBLDummy.Rows.Clear()
            For Each row In Me.DS.Ledger_View.Rows
                rowDummy = TBLDummy.NewLedger_ViewRow
                For jk As Integer = 0 To TBLDummy.Columns.Count - 1
                    rowDummy.Item(jk) = row.Item(jk)
                Next
                TBLDummy.Rows.Add(rowDummy)
            Next
            TBLDummy.DefaultView.Sort = "[Account Code], Date, [Journal No] ASC"

            If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.Jurnal_IsLedgerAccessActive) = "1" Then
                TipeAccess = Me.ObjJournal.GetLedgerAccess(Me.UserCode)
            Else
                TipeAccess = ClsTransaksiJournal.EnumJournalAccesType.FullAccess
            End If

            If TipeAccess = ClsTransaksiJournal.EnumJournalAccesType.Respected_User_Only Or _
            TipeAccess = ClsTransaksiJournal.EnumJournalAccesType.Selected_Users_Only Then
                Me.ObjJournal.GetLedgerAccess(Me.UserCode, ListCreatedBy)
            End If

            Me.DS.Ledger_View.Rows.Clear()
            For Each DV In TBLDummy.DefaultView
                rowDummy = DV.Row

                If rowDummy.Account_Code <> DummyAccount Then
                    DummyAccount = rowDummy.Account_Code
                    'tambahkan data saldo, sampai dengan hari sebelumnya
                    If Ref_date.Date <= Me.dtpFrom.Value.Date Then
                        'baru-----setelah optimasi
                        BB = Me.GetBalance_AkunNonHeader(Me.ObjAcc.GetPrimary(rowDummy.Account_Code), _
                        Me.dtpFrom.Value.Date - TimeSpan.FromDays(1))
                        BB1 = Me.GetBalanceOriginal_AkunNonHeader(Me.ObjAcc.GetPrimary(rowDummy.Account_Code), _
                                               Me.dtpFrom.Value.Date - TimeSpan.FromDays(1))

                        'lama'----sebelum optimasi
                        'BB = Me.ObjAcc.GetBalance_AkunNonHeader(Me.ObjAcc.GetPrimary(rowDummy.Account_Code), _
                        'Me.dtpFrom.Value.Date - TimeSpan.FromDays(1))
                        'BB1 = Me.ObjAcc.GetBalanceOriginal_AkunNonHeader(Me.ObjAcc.GetPrimary(rowDummy.Account_Code), _
                        'Me.dtpFrom.Value.Date - TimeSpan.FromDays(1))

                    ElseIf Ref_date.Date > Me.dtpFrom.Value.Date And Ref_date <= Me.dtpTo.Value.Date Then
                        'baru----setelah optimasi
                        BB = Me.GetBalance_AkunNonHeader(Me.ObjAcc.GetPrimary(rowDummy.Account_Code), _
                                              Me.dtpFrom.Value.Date - TimeSpan.FromDays(1))
                        BB1 = Me.GetBalanceOriginal_AkunNonHeader(Me.ObjAcc.GetPrimary(rowDummy.Account_Code), _
                                              Ref_date)

                        'lama'----sebelum optimasi
                        'BB = Me.ObjAcc.GetBalance_AkunNonHeader(Me.ObjAcc.GetPrimary(rowDummy.Account_Code), _
                        'Ref_date)
                        'BB1 = Me.ObjAcc.GetBalanceOriginal_AkunNonHeader(Me.ObjAcc.GetPrimary(rowDummy.Account_Code), _
                        'Ref_date)
                    Else
                        BB = 0
                        BB1 = 0
                    End If

                    'aul new
                    'Dim addBeginningBalance As Boolean = True
                    'If rowDummy.Journal_No = "Beginning Balance Only" And BB = 0 And BB1 = 0 Then
                    '    addBeginningBalance = False
                    'End If
                    '-----------------------------------------

                    'If addBeginningBalance = True Then 'aul new
                    row_BB = Me.DS.Ledger_View.NewLedger_ViewRow
                    With row_BB
                        .Debet = 0
                        .Desc = "Beginning Balance"
                        .Journal_No = ""
                        .Kredit = 0
                        .Primary = -1
                        .Primary_Akun = rowDummy.Primary_Akun
                        .Remarks = ""
                        .Saldo = BB
                        .Ratio = 1
                        ._Date = Me.dtpFrom.Value.Date
                        .Saldo_Normal = rowDummy.Saldo_Normal
                        .Debet_Original = 0
                        .Credit_Original = 0
                        .Saldo_Original = BB1
                        .Account_Code = rowDummy.Account_Code
                        .Account_Desc = rowDummy.Account_Desc
                        .Default_Precision = Me.ObjCur.GetDefaultPrecission(Me.ObjAcc.GetPrimaryCurrency(rowDummy.Primary_Akun))
                        .Currency_Code = Me.ObjCur.GetCurrencyCode(Me.ObjAcc.GetPrimaryCurrency(rowDummy.Primary_Akun))
                        .Created_By = ""
                        .Tipe_Journal = rowDummy.Tipe_Journal
                        .Tipe_Jurnal_2 = rowDummy.Tipe_Jurnal_2
                    End With
                    Me.DS.Ledger_View.Rows.Add(row_BB)
                    'End If
                End If

                If rowDummy.Journal_No <> "Beginning Balance Only" Then 'aul new
                    row = Me.DS.Ledger_View.NewLedger_ViewRow
                    For jk As Integer = 0 To TBLDummy.Columns.Count - 1
                        row.Item(jk) = rowDummy.Item(jk)
                    Next
                    If row.Journal_No <> "Beginning Balance" Then
                        If TipeAccess = ClsTransaksiJournal.EnumJournalAccesType.Respected_User_Only Or _
                                       TipeAccess = ClsTransaksiJournal.EnumJournalAccesType.Selected_Users_Only Then
                            If ListCreatedBy.Contains(row.Created_By) = False Then
                                row.Journal_No = "***"
                                row.Desc = "***(Created by:" + row.Created_By + ")***"
                            End If
                        End If
                    End If
                    Me.DS.Ledger_View.Rows.Add(row)
                End If
            Next

            If Me.DS.Ledger_View.Rows.Count > 1 Then
                'hitung saldo mpe bawah
                DummyAccount = ""
                StartIndex = 1
                For Each row_BB In DS.Ledger_View.Rows
                    If DummyAccount <> row_BB.Account_Code Then
                        DummyAccount = row_BB.Account_Code
                        For i = StartIndex To Me.DS.Ledger_View.Rows.Count - 1
                            row_BB1 = DS.Ledger_View.Rows(i)
                            If row_BB1.Account_Code <> DummyAccount Then
                                Exit For
                            Else
                                row_BB_Before = DS.Ledger_View.Rows(i - 1)
                                If row_BB1.Saldo_Normal = 0 Then
                                    row_BB1.Saldo = row_BB_Before.Saldo + row_BB1.Debet - row_BB1.Kredit
                                    row_BB1.Saldo_Original = row_BB_Before.Saldo_Original + row_BB1.Debet_Original - row_BB1.Credit_Original
                                Else
                                    row_BB1.Saldo = row_BB_Before.Saldo - row_BB1.Debet + row_BB1.Kredit
                                    row_BB1.Saldo_Original = row_BB_Before.Saldo_Original - row_BB1.Debet_Original + row_BB1.Credit_Original
                                End If
                            End If
                        Next
                    End If
                    StartIndex += 1
                Next
            Else
                If Me.TxtPrimaryAcc.Text <> -1 Then
                    'tambahkan data saldo, sampai dengan hari sebelumnya
                    If Ref_date.Date <= Me.dtpFrom.Value.Date Then
                        BB = Me.ObjAcc.GetBalance_AkunNonHeader(Me.TxtPrimaryAcc.Text, _
                        Me.dtpFrom.Value.Date - TimeSpan.FromDays(1))
                        BB1 = Me.ObjAcc.GetBalanceOriginal_AkunNonHeader(Me.TxtPrimaryAcc.Text, _
                        Me.dtpFrom.Value.Date - TimeSpan.FromDays(1))
                    ElseIf Ref_date.Date > Me.dtpFrom.Value.Date And Ref_date <= Me.dtpTo.Value.Date Then
                        BB = Me.ObjAcc.GetBalance_AkunNonHeader(Me.TxtPrimaryAcc.Text, _
                        Ref_date)
                        BB1 = Me.ObjAcc.GetBalanceOriginal_AkunNonHeader(Me.TxtPrimaryAcc.Text, _
                        Ref_date)
                    Else
                        BB = 0
                        BB1 = 0
                    End If

                    row_BB = Me.DS.Ledger_View.NewLedger_ViewRow
                    With row_BB
                        .Debet = 0
                        .Desc = ""
                        .Journal_No = "Beginning Balance"
                        .Kredit = 0
                        .Primary = -1
                        .Primary_Akun = Me.TxtPrimaryAcc.Text
                        .Remarks = ""
                        .Saldo = BB
                        .Ratio = 1
                        ._Date = Me.dtpFrom.Value.Date
                        .Saldo_Normal = Me.ObjAcc.GetSaldoNormal(Me.TxtPrimaryAcc.Text)
                        .Debet_Original = 0
                        .Credit_Original = 0
                        .Saldo_Original = BB1
                        .Account_Code = Me.ObjAcc.GetAccountCode(Me.TxtPrimaryAcc.Text)
                        .Account_Desc = Me.ObjAcc.GetDescription(.Account_Code)
                        .Default_Precision = Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryAcc.Text)
                        .Currency_Code = Me.ObjCur.GetCurrencyCode(Me.ObjAcc.GetPrimaryCurrency(Me.TxtPrimaryAcc.Text))
                        .Created_By = ""
                        .Tipe_Journal = rowDummy.Tipe_Journal
                        .Tipe_Jurnal_2 = rowDummy.Tipe_Jurnal_2
                    End With
                    Me.DS.Ledger_View.Rows.Add(row_BB)
                Else

                End If
            End If

            ''buang akun yang tidak di cek list
            'If Me.chkAccountFilter.Checked = True Then
            '    'kumpulkan list primary akun
            '    For Each DV In Me.clbAccount.CheckedItems
            '        rowItem = DV.Row
            '        If ListPrimaryAkun.Contains(rowItem.Primary) = False Then
            '            ListPrimaryAkun.Add(rowItem.Primary)
            '        End If
            '    Next

            '    'buang baris yang tidak ada dalam primary akun
            '    i = 0
            '    While i <= Me.DS.Ledger_View.Rows.Count - 1
            '        row_BB = DS.Ledger_View.Rows(i)
            '        If ListPrimaryAkun.Contains(row_BB.Primary_Akun) = False Then
            '            DS.Ledger_View.Rows.Remove(row_BB)
            '        Else
            '            i += 1
            '        End If
            '    End While
            'End If



            'pindahkan nilai original, bila potrait dan diminta tampilan original
            If Me.ComboBox1.SelectedIndex = 0 And Me.chkOriginal.Checked = True Then
                For Each row_BB In DS.Ledger_View.Rows
                    row_BB.Saldo = row_BB.Saldo_Original
                    row_BB.Debet = row_BB.Debet_Original
                    row_BB.Kredit = row_BB.Credit_Original
                Next
            End If

            'tampilkan report
            If Me.ComboBox1.SelectedIndex = 0 Then
                If Me.chkPageAfter.Checked = True Then
                    Me.report2_1.SetDataSource(DS)
                    Me.report2_1.SetParameterValue(report2_1.Parameter_TanggalDari_1.ParameterFieldName, Me.dtpFrom.Value.Date)
                    Me.report2_1.SetParameterValue(report2_1.Parameter_TanggalSampai_1.ParameterFieldName, Me.dtpTo.Value.Date)
                    Me.report2_1.SetParameterValue(report2_1.Parameter_AccountNumber.ParameterFieldName, Me.TxtCodeAcc.Text)
                    Me.report2_1.SetParameterValue(report2_1.Parameter_AccountDesc.ParameterFieldName, Me.TxtDesc.Text)
                    Me.report2_1.SetParameterValue(report2_1.Parameter_FormatDecimal.ParameterFieldName, Me.ObjCur.GetDecimal_DefaultCurrency)
                    If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                        Me.Negatif = 1
                    Else
                        Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
                    End If
                    Me.report2_1.SetParameterValue(report2.Parameter_Negatif.ParameterFieldName, Me.Negatif)

                    Me.crv1.ReportSource = Me.report2_1
                    Me.crv1.Refresh()
                    Me.crv1.Show()
                Else
                    Me.report2.SetDataSource(DS)
                    Me.report2.SetParameterValue(report2.Parameter_TanggalDari_1.ParameterFieldName, Me.dtpFrom.Value.Date)
                    Me.report2.SetParameterValue(report2.Parameter_TanggalSampai_1.ParameterFieldName, Me.dtpTo.Value.Date)
                    Me.report2.SetParameterValue(report2.Parameter_AccountNumber.ParameterFieldName, Me.TxtCodeAcc.Text)
                    Me.report2.SetParameterValue(report2.Parameter_AccountDesc.ParameterFieldName, Me.TxtDesc.Text)
                    Me.report2.SetParameterValue(report2.Parameter_FormatDecimal.ParameterFieldName, Me.ObjCur.GetDecimal_DefaultCurrency)
                    If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                        Me.Negatif = 1
                    Else
                        Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
                    End If
                    Me.report2.SetParameterValue(report2.Parameter_Negatif.ParameterFieldName, Me.Negatif)

                    Me.crv1.ReportSource = Me.report2
                    Me.crv1.Refresh()
                    Me.crv1.Show()
                End If

            Else
                If Me.chkPageAfter.Checked = True Then
                    Me.report3_1.SetDataSource(DS)
                    Me.report3_1.SetParameterValue(report3.Parameter_TanggalDari_1.ParameterFieldName, Me.dtpFrom.Value.Date)
                    Me.report3_1.SetParameterValue(report3.Parameter_TanggalSampai_1.ParameterFieldName, Me.dtpTo.Value.Date)
                    Me.report3_1.SetParameterValue(report3.Parameter_AccountNumber.ParameterFieldName, Me.TxtCodeAcc.Text)
                    Me.report3_1.SetParameterValue(report3.Parameter_AccountDesc.ParameterFieldName, Me.TxtDesc.Text)
                    Me.report3_1.SetParameterValue(report3.Parameter_FormatDecimal.ParameterFieldName, Me.ObjCur.GetDecimal_DefaultCurrency)
                    If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                        Me.Negatif = 1
                    Else
                        Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
                    End If
                    Me.report3_1.SetParameterValue(report3_1.Parameter_Negatif.ParameterFieldName, Me.Negatif)

                    Me.crv1.ReportSource = Me.report3_1
                    Me.crv1.Refresh()
                    Me.crv1.Show()
                Else
                    Me.report3.SetDataSource(DS)
                    Me.report3.SetParameterValue(report3.Parameter_TanggalDari_1.ParameterFieldName, Me.dtpFrom.Value.Date)
                    Me.report3.SetParameterValue(report3.Parameter_TanggalSampai_1.ParameterFieldName, Me.dtpTo.Value.Date)
                    Me.report3.SetParameterValue(report3.Parameter_AccountNumber.ParameterFieldName, Me.TxtCodeAcc.Text)
                    Me.report3.SetParameterValue(report3.Parameter_AccountDesc.ParameterFieldName, Me.TxtDesc.Text)
                    Me.report3.SetParameterValue(report3.Parameter_FormatDecimal.ParameterFieldName, Me.ObjCur.GetDecimal_DefaultCurrency)
                    If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                        Me.Negatif = 1
                    Else
                        Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
                    End If
                    Me.report3.SetParameterValue(report3.Parameter_Negatif.ParameterFieldName, Me.Negatif)

                    Me.crv1.ReportSource = Me.report3
                    Me.crv1.Refresh()
                    Me.crv1.Show()
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            AccountList = Nothing
            ListPrimaryAkun = Nothing
            ListCreatedBy = Nothing
        End Try
    End Sub

    Private Sub CreatedLedger_WithReverse()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim row As DseLaporanLedger.Ledger_View_With_ReverseRow
        Dim i As Integer = 0
        Dim Dummy As Double = 0
        Dim BB As Double = 0
        Dim row_BB, row_BB_Before As DseLaporanLedger.Ledger_View_With_ReverseRow
        Dim Saldo_Normal As Integer = Me.ObjAcc.GetSaldoNormal(Me.TxtPrimaryAcc.Text)
        Dim Ref_date As Date = Me.Obj_Closing.GetConvertionDate
        Dim TBLDummy As New DseLaporanLedger.Ledger_View_With_ReverseDataTable
        Dim rowDummy As DseLaporanLedger.Ledger_View_With_ReverseRow
        Dim DV As DataRowView

        Try
            If Me.TxtPrimaryAcc.Text <> "" And Me.TxtPrimaryAcc.Text <> "-1" Then
                '1. isikan data
                Me.DTA_With_Reverse.Fill(Me.DS.Ledger_View_With_Reverse, _
                Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date, Me.TxtPrimaryAcc.Text)
                '2. Hilangkan semua yang akunnya sesuai dengan yang diminta
                While i < Me.DS.Ledger_View_With_Reverse.Rows.Count
                    row = Me.DS.Ledger_View_With_Reverse.Rows(i)
                    If row.Account_Code = Me.TxtCodeAcc.Text Then
                        Me.DS.Ledger_View_With_Reverse.Rows.Remove(row)
                    Else
                        i += 1
                    End If
                End While
                '3. Balik posisi debet dan credit
                For Each row In Me.DS.Ledger_View_With_Reverse.Rows
                    Dummy = row.Debet
                    row.Debet = row.Kredit
                    row.Kredit = Dummy

                    If Saldo_Normal = 0 Then
                        If row.Debet > 0 Then
                            row.InOut = 0
                        Else
                            row.InOut = 1
                        End If
                    Else
                        If row.Debet > 0 Then
                            row.InOut = 1
                        Else
                            row.InOut = 0
                        End If
                    End If
                Next
                '4. urutan berdasarkan tanggal, in out, dan no transaksi
                TBLDummy.Rows.Clear()
                For Each row In Me.DS.Ledger_View_With_Reverse.Rows
                    rowDummy = TBLDummy.NewLedger_View_With_ReverseRow
                    For jk As Integer = 0 To TBLDummy.Columns.Count - 1
                        rowDummy.Item(jk) = row.Item(jk)
                    Next
                    TBLDummy.Rows.Add(rowDummy)
                Next
                TBLDummy.DefaultView.Sort = "Date, InOut, [Journal Number]"
                Me.DS.Ledger_View_With_Reverse.Rows.Clear()
                For Each DV In TBLDummy.DefaultView
                    rowDummy = DV.Row
                    row = Me.DS.Ledger_View_With_Reverse.NewLedger_View_With_ReverseRow
                    For jk As Integer = 0 To TBLDummy.Columns.Count - 1
                        row.Item(jk) = rowDummy.Item(jk)
                    Next
                    Me.DS.Ledger_View_With_Reverse.Rows.Add(row)
                Next

                'tambahkan data saldo, sampai denga hari sebelumnya
                If Ref_date.Date <= Me.dtpFrom.Value.Date Then
                    BB = Me.ObjAcc.GetBalance_AkunNonHeader(Me.TxtPrimaryAcc.Text, _
                    Me.dtpFrom.Value.Date - TimeSpan.FromDays(1))
                ElseIf Ref_date.Date > Me.dtpFrom.Value.Date And Ref_date <= Me.dtpTo.Value.Date Then
                    BB = Me.ObjAcc.GetBalance_AkunNonHeader(Me.TxtPrimaryAcc.Text, _
                    Ref_date)
                Else
                    BB = 0
                End If

                row_BB = Me.DS.Ledger_View_With_Reverse.NewLedger_View_With_ReverseRow
                With row_BB
                    .Debet = 0
                    .Description = "Beginning Balance"
                    .Journal_Number = ""
                    .Kredit = 0
                    .Remarks = ""
                    .Saldo = BB
                    ._Date = Me.dtpFrom.Value.Date
                    .Account_Code = ""
                    .Desc = ""
                End With
                Me.DS.Ledger_View_With_Reverse.Rows.InsertAt(row_BB, 0)
                If Me.DS.Ledger_View_With_Reverse.Rows.Count > 1 Then
                    'hitung saldo mpe bawah
                    For i = 1 To Me.DS.Ledger_View_With_Reverse.Rows.Count - 1
                        row_BB = DS.Ledger_View_With_Reverse.Rows(i)
                        row_BB_Before = DS.Ledger_View_With_Reverse.Rows(i - 1)
                        If Saldo_Normal = 0 Then
                            row_BB.Saldo = row_BB_Before.Saldo + row_BB.Debet - row_BB.Kredit
                        Else
                            row_BB.Saldo = row_BB_Before.Saldo - row_BB.Debet + row_BB.Kredit
                        End If
                    Next
                End If

                '4. Tampilkan laporan
                Me.report_reverse.SetDataSource(DS)
                Me.report_reverse.SetParameterValue(report_reverse.Parameter_TanggalDari_1.ParameterFieldName, Me.dtpFrom.Value.Date)
                Me.report_reverse.SetParameterValue(report_reverse.Parameter_TanggalSampai_1.ParameterFieldName, Me.dtpTo.Value.Date)
                Me.report_reverse.SetParameterValue(report_reverse.Parameter_AccountNumber.ParameterFieldName, Me.TxtCodeAcc.Text)
                Me.report_reverse.SetParameterValue(report_reverse.Parameter_AccountDesc.ParameterFieldName, Me.TxtDesc.Text)
                Me.report_reverse.SetParameterValue(report_reverse.Parameter_FormatDecimal.ParameterFieldName, Me.ObjCur.GetDecimal_DefaultCurrency)
                If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                    Me.Negatif = 1
                Else
                    Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
                End If
                Me.report_reverse.SetParameterValue(report_reverse.Parameter_Negatif.ParameterFieldName, Me.Negatif)

                Me.crv1.ReportSource = Me.report_reverse
                Me.crv1.Refresh()
                Me.crv1.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub TxtDeptCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDeptCode.KeyDown
        Dim i As New Core_Module.FrmMasterDepartementCari
        Dim Obj As New Core_Module.ClsMasterDepartement
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If Obj.ApakahSudahAda(Me.TxtDeptCode.Text) = True Then
                    Me.TxtPrimaryDept.Text = Obj.GetPrimary(Me.TxtDeptCode.Text)
                    If Me.TxtJobCode.Visible = True Then
                        Me.TxtJobCode.Focus()
                    Else
                        Me.btnShow.Focus()
                    End If
                Else
                    If Me.TxtDeptCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtDeptCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryDepartement > 0 Then
                        Me.TxtPrimaryDept.Text = i.PrimaryDepartement
                        Me.TxtDeptCode.Text = Obj.GetDeptCode(Me.TxtPrimaryDept.Text)
                        If Me.TxtJobCode.Visible = True Then
                            Me.TxtJobCode.Focus()
                        Else
                            Me.btnShow.Focus()
                        End If
                    Else
                        Me.TxtDeptCode.Clear()
                        Me.TxtPrimaryDept.Text = "-1"
                        Me.TxtDeptCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            Obj = Nothing
        End Try
    End Sub

    Private Sub TxtDeptCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDeptCode.Leave
        Dim i As New Core_Module.FrmMasterDepartementCari
        Dim obj As New Core_Module.ClsMasterDepartement
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtDeptCode.Text <> "" Then
                If obj.ApakahSudahAda(Me.TxtDeptCode.Text) = True Then
                    Me.TxtPrimaryDept.Text = obj.GetPrimary(Me.TxtDeptCode.Text)
                Else
                    If Me.TxtDeptCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtDeptCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryDepartement > 0 Then
                        Me.TxtPrimaryDept.Text = i.PrimaryDepartement
                        Me.TxtDeptCode.Text = obj.GetDeptCode(Me.TxtPrimaryDept.Text)
                        If Me.TxtJobCode.Visible = True Then
                            Me.TxtJobCode.Focus()
                        Else
                            Me.btnShow.Focus()
                        End If
                    Else
                        Me.TxtDeptCode.Clear()
                        Me.TxtPrimaryDept.Text = "-1"
                        Me.TxtDeptCode.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryDept.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFindDept_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindDept.Click
        Dim i As New Core_Module.FrmMasterDepartementCari
        Dim obj As New Core_Module.ClsMasterDepartement

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.TxtDeptCode.Text <> "" Then
                i.txtCari.Text = Me.TxtDeptCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryDepartement > 0 Then
                Me.TxtPrimaryDept.Text = i.PrimaryDepartement
                Me.TxtDeptCode.Text = obj.GetDeptCode(Me.TxtPrimaryDept.Text)
                If Me.TxtJobCode.Visible = True Then
                    Me.TxtJobCode.Focus()
                Else
                    Me.btnShow.Focus()
                End If
            Else
                Me.TxtDeptCode.Clear()
                Me.TxtDeptCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub TxtJobCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtJobCode.KeyDown
        Dim i As New Core_Module.FrmMasterJobCari
        Dim Obj As New Core_Module.ClsMasterJob
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If Obj.ApakahSudahAda(Me.TxtJobCode.Text) = True Then
                    Me.TxtPrimaryJob.Text = Obj.GetPrimary(Me.TxtJobCode.Text)
                    If Me.TxtJobCode.Visible = True Then
                        Me.TxtJobCode.Focus()
                    Else
                        Me.btnShow.Focus()
                    End If
                Else
                    If Me.TxtJobCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtJobCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryJob > 0 Then
                        Me.TxtPrimaryJob.Text = i.PrimaryJob
                        Me.TxtJobCode.Text = Obj.GetJobCode(Me.TxtPrimaryJob.Text)
                        If Me.TxtJobCode.Visible = True Then
                            Me.TxtJobCode.Focus()
                        Else
                            Me.btnShow.Focus()
                        End If
                    Else
                        Me.TxtJobCode.Clear()
                        Me.TxtPrimaryJob.Text = "-1"
                        Me.TxtJobCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            Obj = Nothing
        End Try
    End Sub

    Private Sub TxtJobCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJobCode.Leave
        Dim i As New Core_Module.FrmMasterJobCari
        Dim obj As New Core_Module.ClsMasterJob
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtJobCode.Text <> "" Then
                If obj.ApakahSudahAda(Me.TxtJobCode.Text) = True Then
                    Me.TxtPrimaryJob.Text = obj.GetPrimary(Me.TxtJobCode.Text)
                Else
                    If Me.TxtJobCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtJobCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryJob > 0 Then
                        Me.TxtPrimaryJob.Text = i.PrimaryJob
                        Me.TxtJobCode.Text = obj.GetJobCode(Me.TxtPrimaryJob.Text)
                        If Me.TxtJobCode.Visible = True Then
                            Me.TxtJobCode.Focus()
                        Else
                            Me.btnShow.Focus()
                        End If
                    Else
                        Me.TxtJobCode.Clear()
                        Me.TxtPrimaryJob.Text = "-1"
                        Me.TxtJobCode.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryJob.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFindJob_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindJob.Click
        Dim i As New Core_Module.FrmMasterJobCari
        Dim obj As New Core_Module.ClsMasterJob

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.TxtJobCode.Text <> "" Then
                i.txtCari.Text = Me.TxtJobCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryJob > 0 Then
                Me.TxtPrimaryJob.Text = i.PrimaryJob
                Me.TxtJobCode.Text = obj.GetJobCode(Me.TxtPrimaryJob.Text)
                If Me.TxtJobCode.Visible = True Then
                    Me.TxtJobCode.Focus()
                Else
                    Me.btnShow.Focus()
                End If
            Else
                Me.TxtJobCode.Clear()
                Me.TxtJobCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub TxtJobCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJobCode.GotFocus
        Try
            Me.TxtJobCode.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtJobCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJobCode.LostFocus
        Try
            Me.TxtJobCode.BackColor = Me.TxtPrimaryDept.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDeptCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDeptCode.GotFocus
        Try
            Me.TxtDeptCode.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDeptCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDeptCode.LostFocus
        Try
            Me.TxtDeptCode.BackColor = Me.TxtPrimaryDept.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanLedger_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            Me.chkShowItemRemarks.Visible = True ''rosadin-ac1

            Me.DgvUser.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvAccount.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            'jika tidak ada hak akses atas job, maka kolom job ditutup
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterJob") = False Then
                Me.LblJob.Enabled = False
                Me.TxtJobCode.Enabled = False
                Me.TxtJobCode.TabStop = False
                Me.btnFindJob.Enabled = False
            End If

            'jika tidak ada hak akses atas dept, maka kolom dept ditutup
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterDepartement") = False Then
                Me.LblDept.Enabled = False
                Me.TxtDeptCode.Enabled = False
                Me.TxtDeptCode.TabStop = False
                Me.btnFindDept.Enabled = False
            End If

            Me.ComboBox1.SelectedIndex = 0
            Me.ObjAcc.IsikanCLBAkun(Me.TBLAkun, Me.clbAccount)
            Me.IsikanJournalType()
            For i As Integer = 0 To Me.clbTipeJournal.Items.Count - 1
                Me.clbTipeJournal.SetItemCheckState(i, CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkAccountFilter_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAccountFilter.CheckedChanged
        Try
            If Me.chkAccountFilter.Checked = True Then
                Me.clbAccount.Enabled = True
                Me.llbCheckAll.Enabled = True
                Me.llbUncheckAll.Enabled = True
                Me.TxtPrimaryAcc.Text = "-1"
                Me.TxtCodeAcc.Clear()
                Me.TxtDesc.Clear()
            Else
                'Me.TxtPrimaryAcc.Text = "-1"
                'Me.TxtCodeAcc.Clear()
                'Me.TxtDesc.Clear()
                Me.clbAccount.Enabled = False
                Me.llbCheckAll.Enabled = False
                Me.llbUncheckAll.Enabled = False
                Me.UncheckAll()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCheckAll.LinkClicked
        Try
            For i As Integer = 0 To Me.clbAccount.Items.Count - 1
                Me.clbAccount.SetItemCheckState(i, CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUncheckAll.LinkClicked
        Me.UncheckAll()
    End Sub

    Private Sub UncheckAll()
        Try
            For i As Integer = 0 To Me.clbAccount.Items.Count - 1
                Me.clbAccount.SetItemCheckState(i, CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            If Me.ComboBox1.SelectedIndex = 0 Then
                'potrait
                Me.chkOriginal.Enabled = True
                Me.chkOriginal.TabStop = True
            Else
                'landscape
                Me.chkOriginal.Enabled = False
                Me.chkOriginal.TabStop = False
                Me.chkOriginal.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvUser_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvUser.CellBeginEdit
        Try
            If Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value <> -1 Then
                If e.ColumnIndex = Me.ColCreatedBy.Index Then
                    Me.buffer = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColCreatedBy.Index Then
                    Me.buffer = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvUser_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvUser.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvUser.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim KodeUser As String = ""
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColCreatedBy.Index Then
                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjUser.ApakahSudahAda(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            KodeUser = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                            index = Me.GetIndexItem(KodeUser, e.RowIndex)
                            If index < 0 Then
                                Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = Me.ObjUser.GetPrimary(Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value)
                                Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value
                                Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.ObjUser.GetNama(Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BgwUser.RunWorkerAsync()
                                    Else
                                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterUser
                            ij.TxtCari.Text = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.backColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.Kodeuser <> "" Then
                                index = Me.GetIndexItem(ij.Kodeuser, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = ij.Kodeuser
                                    Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.ObjUser.GetNama(ij.Kodeuser)
                                    Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = Me.ObjUser.GetPrimary(ij.Kodeuser)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BgwUser.RunWorkerAsync()
                                        Else
                                            Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Exit Sub
                                    Else
                                        Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = ij.Kodeuser
                                        Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.ObjUser.GetNama(ij.Kodeuser)
                                        Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = Me.ObjUser.GetPrimary(ij.Kodeuser)
                                    End If
                                End If
                            ElseIf ij.Kodeuser = "" Then
                                If e.ColumnIndex = Me.ColCreatedBy.Index Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        If Me.DgvUser.Rows.Count - 1 <> e.RowIndex Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BgwUser.RunWorkerAsync()
                                        End If
                                    Else
                                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                Else
                                    Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.DgvUser)
                    End If
                Else
                    Me.EraseRowIndex = e.RowIndex
                    Me.BgwUser.RunWorkerAsync()
                End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvUser_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvUser.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try

            rowDGV = Me.DgvUser.Rows(Me.DgvUser.RowCount - 1)
            With rowDGV
                '.Cells(Me.ColNo.Index).Value = rowDGV.Index + 1
                .Cells(Me.ColCreatedBy.Index).Value = ""
                .Cells(Me.ColCreatedBy_Name.Index).Value = ""
                .Cells(Me.ColCreatedBy_Primary.Index).Value = "-1"
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Function GetIndexItem(ByVal vKodeUser As String, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvUser.Rows.Count - 2
                RowDGV = Me.DgvUser.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColCreatedBy.Index).Value = vKodeUser Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    Private Function GetIndexItemAcc(ByVal vNoAccount As String, _
        ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.dgvAccount.Rows.Count - 2
                RowDGV = Me.dgvAccount.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColNoAcc.Index).Value = vNoAccount Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
 
    Private Sub BgwUser_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwUser.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwUser_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwUser.RunWorkerCompleted
        Try
            Me.DgvUser.Rows.RemoveAt(Me.EraseRowIndex)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAccount_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvAccount.CellBeginEdit
        Try
            If Me.dgvAccount.Item(Me.ColPrimaryAcc.Index, e.RowIndex).Value <> -1 Then
                If e.ColumnIndex = Me.ColNoAcc.Index Then
                    Me.buffer = Me.dgvAccount.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColNameAcc.Index Then
                    Me.buffer = Me.dgvAccount.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAccount_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvAccount.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvAccount.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim NoAccount As String = ""
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.dgvAccount.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColNoAcc.Index Then
                        Me.dgvAccount.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvAccount.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjAcc.ApakahSudahAda_Aktif_NonHeader(Me.dgvAccount.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            NoAccount = Me.dgvAccount.Item(e.ColumnIndex, e.RowIndex).Value
                            index = Me.GetIndexItemAcc(NoAccount, e.RowIndex)
                            If index < 0 Then
                                Me.dgvAccount.Item(Me.ColPrimaryAcc.Index, e.RowIndex).Value = Me.ObjAcc.GetPrimary(Me.dgvAccount.Item(Me.ColNoAcc.Index, e.RowIndex).Value)
                                Me.dgvAccount.Item(Me.ColNoAcc.Index, e.RowIndex).Value = Me.dgvAccount.Item(Me.ColNoAcc.Index, e.RowIndex).Value
                                Me.dgvAccount.Item(Me.ColNameAcc.Index, e.RowIndex).Value = Me.ObjAcc.GetDescription(Me.dgvAccount.Item(Me.ColNoAcc.Index, e.RowIndex).Value)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvAccount.Item(Me.ColNoAcc.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndexAcc = e.RowIndex
                                        Me.BgwAccount.RunWorkerAsync()
                                    Else
                                        Me.dgvAccount.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New FrmMasterAkun_Cari
                            ij.txtCari.Text = Me.dgvAccount.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.mPrimaryAkun <> "-1" Then
                                NoAccount = ObjAcc.GetAccountCode(ij.mPrimaryAkun)
                                index = Me.GetIndexItemAcc(NoAccount, e.RowIndex)
                                If index < 0 Then
                                    Me.dgvAccount.Item(Me.ColNoAcc.Index, e.RowIndex).Value = NoAccount
                                    Me.dgvAccount.Item(Me.ColNameAcc.Index, e.RowIndex).Value = Me.ObjAcc.GetDescription(NoAccount)
                                    Me.dgvAccount.Item(Me.ColPrimaryAcc.Index, e.RowIndex).Value = ij.mPrimaryAkun
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseRowIndexAcc = e.RowIndex
                                            Me.BgwAccount.RunWorkerAsync()
                                        Else
                                            Me.dgvAccount.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Exit Sub
                                    Else
                                        Me.dgvAccount.Item(Me.ColNoAcc.Index, e.RowIndex).Value = NoAccount
                                        Me.dgvAccount.Item(Me.ColNameAcc.Index, e.RowIndex).Value = Me.ObjAcc.GetDescription(NoAccount)
                                        Me.dgvAccount.Item(Me.ColPrimaryAcc.Index, e.RowIndex).Value = ij.mPrimaryAkun
                                    End If
                                End If
                            ElseIf ij.mPrimaryAkun = "-1" Then
                                If e.ColumnIndex = Me.ColNoAcc.Index Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        If Me.dgvAccount.Rows.Count - 1 <> e.RowIndex Then
                                            Me.EraseRowIndexAcc = e.RowIndex
                                            Me.BgwAccount.RunWorkerAsync()
                                        End If
                                    Else
                                        Me.dgvAccount.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                Else
                                    Me.dgvAccount.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.DgvUser)
                    End If
                Else
                    If e.RowIndex <> Me.dgvAccount.Rows.Count - 1 Then
                        Me.EraseRowIndexAcc = e.RowIndex
                        Me.BgwAccount.RunWorkerAsync()
                    End If
                End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwAccount_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwAccount.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwAccount_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwAccount.RunWorkerCompleted
        Try
            Me.dgvAccount.Rows.RemoveAt(Me.EraseRowIndexAcc)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanJournalType() '@hdar
        Try
            Me.clbTipeJournal.Items.Clear()
            Me.clbTipeJournal.Items.Add("General Journal")
            Me.clbTipeJournal.Items.Add("Cash And Bank")
            Me.clbTipeJournal.Items.Add("End of Month Journal")
            Me.clbTipeJournal.Items.Add("Sales Journal")
            Me.clbTipeJournal.Items.Add("Sales Return Journal")
            Me.clbTipeJournal.Items.Add("Sales Down Payment")
            Me.clbTipeJournal.Items.Add("Purchase Journal")
            Me.clbTipeJournal.Items.Add("Purchase DownPayment Journal")
            Me.clbTipeJournal.Items.Add("Purchase Return Journal")
            Me.clbTipeJournal.Items.Add("Sales Receivement Journal")
            Me.clbTipeJournal.Items.Add("Purchase Payment Journal")
            Me.clbTipeJournal.Items.Add("Credit Note")
            Me.clbTipeJournal.Items.Add("Debit Note")
            If Me.ObjJournal.GetJml(ClsTransaksiJournal.EnumTipeJournal.GeneralJournal, ClsTransaksiJournal.EnumTipeJurnal2.HR_Payroll_Loan) > 0 Or _
                      Me.ObjJournal.GetJml(ClsTransaksiJournal.EnumTipeJournal.GeneralJournal, ClsTransaksiJournal.EnumTipeJurnal2.HR_Payroll_Loan_Payment) > 0 Or _
                      Me.ObjJournal.GetJml(ClsTransaksiJournal.EnumTipeJournal.GeneralJournal, ClsTransaksiJournal.EnumTipeJurnal2.HR_Payroll_Payment) > 0 Then
                Me.clbTipeJournal.Items.Add("Loan Journal")
                Me.clbTipeJournal.Items.Add("Loan Payment Journal")
                Me.clbTipeJournal.Items.Add("Salary Payment Journal")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbchekalltipe_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbchekalltipe.LinkClicked '@hdar
        Try
            For i As Integer = 0 To Me.clbTipeJournal.Items.Count - 1
                Me.clbTipeJournal.SetItemCheckState(i, CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llblunchekall_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llblunchekall.LinkClicked '@hdar
        Try
            For i As Integer = 0 To Me.clbTipeJournal.Items.Count - 1
                Me.clbTipeJournal.SetItemCheckState(i, CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function FilterByTipeJournal() As String
        Dim TipeJurnal1 As ClsTransaksiJournal.EnumTipeJournal = ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
        Dim TipeJurnal2 As ClsTransaksiJournal.EnumTipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.Cash_n_Bank
        Dim Filter As String = ""
        Dim CheckedItem As String = ""
        Try
            For i As Integer = 0 To clbTipeJournal.CheckedItems.Count - 1
                CheckedItem = clbTipeJournal.CheckedItems.Item(i)
                If CheckedItem = "General Journal" Then
                    'general journal
                    TipeJurnal1 = ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
                    TipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.Jurnal_NoGrouping
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString
                    Else
                        Filter += " OR " + "([Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString + ")"
                    End If
                ElseIf CheckedItem = "Cash And Bank" Then
                    TipeJurnal1 = ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
                    TipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.Cash_n_Bank
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString
                    Else
                        Filter += " OR " + "([Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString + ")"
                    End If
                ElseIf CheckedItem = "End of Month Journal" Then
                    TipeJurnal1 = ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
                    TipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString
                    Else
                        Filter += " OR " + "([Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString + ")"
                    End If
                ElseIf CheckedItem = "Loan Journal" Then
                    TipeJurnal1 = ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
                    TipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.HR_Payroll_Loan
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString
                    Else
                        Filter += " OR " + "([Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString + ")"
                    End If
                ElseIf CheckedItem = "Loan Payment Journal" Then
                    TipeJurnal1 = ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
                    TipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.HR_Payroll_Loan_Payment
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString
                    Else
                        Filter += " OR " + "([Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString + ")"
                    End If
                ElseIf CheckedItem = "Salary Payment Journal" Then
                    TipeJurnal1 = ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
                    TipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.HR_Payroll_Payment
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString
                    Else
                        Filter += " OR " + "([Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString + ")"
                    End If
                ElseIf CheckedItem = "Sales Journal" Then
                    TipeJurnal1 = ClsTransaksiJournal.EnumTipeJournal.SalesJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Sales Return Journal" Then
                    TipeJurnal1 = ClsTransaksiJournal.EnumTipeJournal.SalesReturnJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Sales Down Payment" Then
                    TipeJurnal1 = ClsTransaksiJournal.EnumTipeJournal.SalesDownPayment
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Purchase Journal" Then
                    TipeJurnal1 = ClsTransaksiJournal.EnumTipeJournal.PurchaseJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Purchase Return Journal" Then
                    TipeJurnal1 = ClsTransaksiJournal.EnumTipeJournal.PurchaseReturnJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Purchase DownPayment Journal" Then
                    TipeJurnal1 = ClsTransaksiJournal.EnumTipeJournal.PurchaseDownPaymentJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Sales Receivement Journal" Then
                    TipeJurnal1 = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Purchase Payment Journal" Then
                    TipeJurnal1 = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    Else
                        Filter += " OR " + "[Tipe Journal] =" + CInt(TipeJurnal1).ToString
                    End If
                ElseIf CheckedItem = "Credit Note" Then
                    TipeJurnal1 = ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
                    TipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.Credit_Note
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString
                    Else
                        Filter += " OR " + "([Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString + ")"
                    End If
                ElseIf CheckedItem = "Debit Note" Then
                    TipeJurnal1 = ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
                    TipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.Debit_Note
                    If Filter = "" Then
                        Filter += "[Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString
                    Else
                        Filter += " OR " + "([Tipe Journal] =" + CInt(TipeJurnal1).ToString + " AND [Tipe Jurnal 2] = " + CInt(TipeJurnal2).ToString + ")"
                    End If
                End If
            Next
            Return Filter
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    '@ahdar
    Dim DTA_A As New DseMasterAkunTableAdapters.tblmaster_akun8TableAdapter
    Dim TBL_A As New DseMasterAkun.tblmaster_akun8DataTable
    Dim row_A As DseMasterAkun.tblmaster_akun8Row
    Dim DTA_B As New DseMasterAkunTableAdapters.tblmaster_akun9TableAdapter
    Dim TBL_B As New DseMasterAkun.tblmaster_akun9DataTable
    Dim row_B As DseMasterAkun.tblmaster_akun9Row
    Dim ClosingDate As Date = Obj_Closing.GetConvertionDate.Date

    Public Function GetBalance_AkunNonHeader(ByVal vPrimaryAkun As Double, ByVal vTanggalSampai As Date) As Decimal
        Dim ObjClosing As New ClsTransaksiClosing

        Try
            If ClosingDate > vTanggalSampai.Date + TimeSpan.FromDays(1) Then
                Return 0
            Else
                TBL_A.DefaultView.RowFilter = "[Primary]=" & vPrimaryAkun
                If Me.TBL_A.DefaultView.Count > 0 Then
                    row_A = TBL_A.DefaultView(0).Row
                    If row_A.Header = 1 Then
                        'fungsi ini hanya terbatas untuk akun2 non header
                        Return 0
                    Else
                        If row_A.Saldo_Normal = 0 Then
                            'debet
                            Return row_A.Beginning_Balance + row_A.Debet - row_A.Kredit
                        Else
                            'credit
                            Return row_A.Beginning_Balance - row_A.Debet + row_A.Kredit
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetBalanceOriginal_AkunNonHeader(ByVal vPrimaryAkun As Double, ByVal vTanggalSampai As Date) As Decimal

        Dim ObjClosing As New ClsTransaksiClosing
        Try
            If ClosingDate > vTanggalSampai.Date + TimeSpan.FromDays(1) Then
                Return 0
            Else
                TBL_B.DefaultView.RowFilter = "[Primary]=" & vPrimaryAkun
                If Me.TBL_B.DefaultView.Count > 0 Then
                    row_B = TBL_B.DefaultView(0).Row
                    If row_B.Header = 1 Then
                        'fungsi ini hanya terbatas untuk akun2 non header
                        Return 0
                    Else
                        If row_B.Saldo_Normal = 0 Then
                            'debet
                            Return row_B.Beggining_Balance_Original_Cur + row_B.Debet_Original - row_B.Credit_Original
                        Else
                            'credit
                            Return row_B.Beggining_Balance_Original_Cur - row_B.Debet_Original + row_B.Credit_Original
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjClosing = Nothing
        End Try
    End Function
    '--
End Class