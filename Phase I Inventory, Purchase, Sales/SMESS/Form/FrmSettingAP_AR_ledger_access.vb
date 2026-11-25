Option Explicit On
Option Strict Off

Partial Public Class FrmSettingAP_AR
    Dim TBLDetailSetting As New DseSettingAP_AR_LedgerAcces.tbltrans_setting_ledger_access_item1DataTable
    Dim rowDetailSetting As DseSettingAP_AR_LedgerAcces.tbltrans_setting_ledger_access_item1Row
    Dim Objuser As New Core_Module.ClsUser
    Dim rowIndexLA_H As Integer = -1
    Dim rowIndexLA_I As Integer = -1
    Dim BufferPrimaryUser As String = ""
    Dim TipeAccess As AccModule.ClsTransaksiJournal.EnumJournalAccesType = AccModule.ClsTransaksiJournal.EnumJournalAccesType.FullAccess

    Private Sub TabControl5_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl5.SelectedIndexChanged
        Try
            If Me.TabControl5.SelectedTab.Name = Me.TabPage10.Name Then
                'cek ledger access
                If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.Jurnal_IsLedgerAccessActive) = "1" Then
                    If Me.Objuser.GetLedgerAccess(Me.UserCode) <> Core_Module.ClsUser.EnumJournalAccesType.FullAccess Then
                        MsgBox("Only user with 'full view' ledger access authorization can make changes on this tab.", MsgBoxStyle.Information, "Information")
                        Me.TabControl5.SelectedTab = Me.TabPage9
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvLA_Header_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvLA_Header.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColLA_Header_ColUserCode.Index Then
                If Me.dgvLA_Header.Rows.Count - 1 = e.RowIndex Then
                    e.Cancel = True
                    Exit Sub
                Else
                    If e.ColumnIndex = Me.ColLA_Header_AccessType.Index Then
                        Me.Buffer = Me.dgvLA_Header.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                        Exit Sub
                    End If
                End If
            Else
                'col user code
                If e.ColumnIndex = Me.ColLA_Header_ColUserCode.Index Then
                    Me.Buffer = Me.dgvLA_Header.Item(e.ColumnIndex, e.RowIndex).Value
                    Me.BufferPrimaryUser = Me.dgvLA_Header.Item(Me.ColLA_Header_PrimaryUser.Index, e.RowIndex).Value
                Else
                    e.Cancel = True
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexHeader(ByVal vPrimaryUser As Double, ByVal vRefIndex As Integer) As Integer
        Dim rowDGV As DataGridViewRow
        Dim RetVal As Integer = -1

        Try
            For ij As Integer = 0 To Me.dgvLA_Header.RowCount - 2
                If ij <> vRefIndex Then
                    rowDGV = Me.dgvLA_Header.Rows(ij)
                    If rowDGV.Cells(Me.ColLA_Header_PrimaryUser.Index).Value = vPrimaryUser Then
                        RetVal = rowDGV.Index
                        Exit For
                    End If
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub GantiPrimaryUserReference(ByVal vPrimaryLama As Double, ByVal vPrimaryBaru As Double)
        Try
            For Each Me.rowDetailSetting In Me.TBLDetailSetting.Rows
                If Me.rowDetailSetting.primary_user_ref = vPrimaryLama Then
                    Me.rowDetailSetting.primary_user_ref = vPrimaryBaru
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvLA_Header_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvLA_Header.CellEndEdit
        Try
            Control.CheckForIllegalCrossThreadCalls = False

            If IsNothing(Me.dgvLA_Header.Item(e.ColumnIndex, e.RowIndex).Value) = True _
            Or Me.dgvLA_Header.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                Me.dgvLA_Header.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            Else
                'cek access type
                If e.ColumnIndex = Me.ColLA_Header_AccessType.Index Then
                    If Me.dgvLA_Header.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        'do nothing
                    Else
                        Me.TBLDetailSetting.DefaultView.RowFilter = "[primary user ref] = " + Me.dgvLA_Header.Item(Me.ColLA_Header_PrimaryUser.Index, e.RowIndex).Value.ToString
                        If Me.dgvLA_Header.Item(e.ColumnIndex, e.RowIndex).Value = "Full View" Then
                            'buang semua item detail yang merefer ke primary user ybs                           
                            While TBLDetailSetting.DefaultView.Count > 0
                                TBLDetailSetting.Rows.Remove(TBLDetailSetting.DefaultView(0).Row)
                            End While
                            Me.dgvLA_Item.AllowUserToAddRows = False
                        Else
                            Me.dgvLA_Item.AllowUserToAddRows = True
                        End If
                        Me.SettinglocationAccessBerubah = True
                    End If
                ElseIf e.ColumnIndex = Me.ColLA_Header_ColUserCode.Index Then
                    Me.dgvLA_Header.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvLA_Header.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    'kemungkinan
                    If Me.Objuser.ApakahSudahAda(Me.dgvLA_Header.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        '1. kode user yang diketik ada --> load data deskripsi + primary user
                        Me.dgvLA_Header.Item(Me.ColLA_Header_UserName.Index, e.RowIndex).Value = Me.Objuser.GetNama(Me.dgvLA_Header.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvLA_Header.Item(Me.ColLA_Header_PrimaryUser.Index, e.RowIndex).Value = Me.Objuser.GetPrimary(Me.dgvLA_Header.Item(e.ColumnIndex, e.RowIndex).Value)
                        'cek duplication
                        If Me.GetIndexHeader(Me.dgvLA_Header.Item(Me.ColLA_Header_PrimaryUser.Index, e.RowIndex).Value, e.RowIndex) <> -1 Then
                            MsgBox("User code duplication, please select another user.", MsgBoxStyle.Information, "Information")
                            If Me.Buffer <> "" Then
                                Me.dgvLA_Header.Item(e.ColumnIndex, e.RowIndex).Value = Buffer
                            Else
                                Me.rowIndexLA_H = e.RowIndex
                                Me.BgwHapusRow_LA.RunWorkerAsync()
                            End If
                            Exit Sub
                        End If
                        'update nilai primary user ref ke user yang baru
                        If CDec(Me.BufferPrimaryUser) <> Me.dgvLA_Header.Item(Me.ColLA_Header_PrimaryUser.Index, e.RowIndex).Value Then
                            Me.GantiPrimaryUserReference(Me.BufferPrimaryUser, Me.dgvLA_Header.Item(Me.ColLA_Header_PrimaryUser.Index, e.RowIndex).Value)
                        End If
                        Me.ReloadItem()
                        Me.SettinglocationAccessBerubah = True
                    Else
                        '2. kode user yang diketik tidak ada --> tampilkan form cari
                        Dim ij As New Core_Module.FrmCariMasterUser
                        ij.TxtCari.Text = Me.dgvLA_Header.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.Kodeuser <> "" Then
                            '2.a. kode user yang dicari ada --> load data deskripsi + primary user + kode user
                            Me.dgvLA_Header.Item(e.ColumnIndex, e.RowIndex).Value = ij.Kodeuser
                            Me.dgvLA_Header.Item(Me.ColLA_Header_UserName.Index, e.RowIndex).Value = Me.Objuser.GetNama(Me.dgvLA_Header.Item(e.ColumnIndex, e.RowIndex).Value)
                            Me.dgvLA_Header.Item(Me.ColLA_Header_PrimaryUser.Index, e.RowIndex).Value = Me.Objuser.GetPrimary(Me.dgvLA_Header.Item(e.ColumnIndex, e.RowIndex).Value)

                            'cek duplication
                            If Me.GetIndexHeader(Me.dgvLA_Header.Item(Me.ColLA_Header_PrimaryUser.Index, e.RowIndex).Value, e.RowIndex) <> -1 Then
                                MsgBox("User code duplication, please select another user.", MsgBoxStyle.Information, "Information")
                                If Me.Buffer <> "" Then
                                    Me.dgvLA_Header.Item(e.ColumnIndex, e.RowIndex).Value = Buffer
                                Else
                                    Me.rowIndexLA_H = e.RowIndex
                                    Me.BgwHapusRow_LA.RunWorkerAsync()
                                End If
                                Exit Sub
                            End If
                            'update nilai primary user ref ke user yang baru
                            If CDec(Me.BufferPrimaryUser) <> Me.dgvLA_Header.Item(Me.ColLA_Header_PrimaryUser.Index, e.RowIndex).Value Then
                                Me.GantiPrimaryUserReference(Me.BufferPrimaryUser, Me.dgvLA_Header.Item(Me.ColLA_Header_PrimaryUser.Index, e.RowIndex).Value)
                            End If
                            Me.SettinglocationAccessBerubah = True
                        Else
                            '2.b. kode user yang dicari tidak ada --> kembalikan ke kondisi buffer
                            Me.dgvLA_Header.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If Me.dgvLA_Header.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                Me.rowIndexLA_H = e.RowIndex
                                Me.BgwHapusRow_LA.RunWorkerAsync()
                                ' Me.dgvLA_Header.Rows.RemoveAt(e.RowIndex)
                                Exit Sub
                            End If
                        End If
                        ij = Nothing
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvLA_Header_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvLA_Header.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgvLA_Header.Rows(Me.dgvLA_Header.Rows.Count - 1)
            rowDGV.Cells(Me.ColLA_Header_AccessType.Index).Value = "Full View"
            rowDGV.Cells(Me.ColLA_Header_PrimaryUser.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwHapusRow_LA_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwHapusRow_LA.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub BgwHapusRow_LA_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwHapusRow_LA.RunWorkerCompleted
        Try
            Me.dgvLA_Header.Rows.RemoveAt(Me.rowIndexLA_H)
            Me.rowIndexLA_H = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanDGVItem_LedgerAccess(ByVal vPrimaryUserRef As Double)
        Dim rowDGV As DataGridViewRow

        Try
            'asumsi : allow user to add row -> false
            Me.TBLDetailSetting.DefaultView.RowFilter = "[primary user ref] = " + vPrimaryUserRef.ToString
            Me.TBLDetailSetting.DefaultView.Sort = "[Kode User] ASC"
            Me.dgvLA_Item.Rows.Clear()

            For Each DV As DataRowView In Me.TBLDetailSetting.DefaultView
                Me.rowDetailSetting = DV.Row
                Me.dgvLA_Item.Rows.Add()
                rowDGV = Me.dgvLA_Item.Rows(Me.dgvLA_Item.Rows.Count - 1)
                rowDGV.Cells(Me.ColLA_Item_UserCode.Index).Value = Me.rowDetailSetting.Kode_User
                rowDGV.Cells(Me.ColLA_Item_UserName.Index).Value = Me.rowDetailSetting.Nama_User
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ReloadItem()
        Dim rowDGV As DataGridViewRow

        Try
            If Me.dgvLA_Header.SelectedCells.Count = 0 Then
                Me.dgvLA_Item.AllowUserToAddRows = False
                Me.dgvLA_Item.Rows.Clear()
            Else
                rowDGV = Me.dgvLA_Header.Rows(Me.dgvLA_Header.SelectedCells(0).RowIndex)
                If rowDGV.Index = Me.dgvLA_Header.Rows.Count - 1 Then
                    'kalo baris terakhir, kosongkan dgv item dan allow user to add row -> false
                    Me.dgvLA_Item.AllowUserToAddRows = False
                    Me.dgvLA_Item.Rows.Clear()
                Else
                    Me.dgvLA_Item.AllowUserToAddRows = False
                    Me.IsikanDGVItem_LedgerAccess(rowDGV.Cells(Me.ColLA_Header_PrimaryUser.Index).Value)
                    Me.dgvLA_Item.AllowUserToAddRows = True

                    If Me.dgvLA_Header.Item(Me.ColLA_Header_AccessType.Index, rowDGV.Index).Value = "Full View" Then
                        'buang semua item detail yang merefer ke primary user ybs                           
                        Me.dgvLA_Item.AllowUserToAddRows = False
                    Else
                        Me.dgvLA_Item.AllowUserToAddRows = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvLA_Header_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvLA_Header.SelectionChanged
        Try
            Me.ReloadItem()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanDataSetting_LedgerAccess()
        Dim DTA As New DseSettingAP_AR_LedgerAccesTableAdapters.tbltrans_setting_ledger_accesTableAdapter
        Dim DTA1 As New DseSettingAP_AR_LedgerAccesTableAdapters.tbltrans_setting_ledger_access_itemTableAdapter
        Dim TipeAccess As AccModule.ClsTransaksiJournal.EnumJournalAccesType = AccModule.ClsTransaksiJournal.EnumJournalAccesType.FullAccess

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = Core_Module.GetConstring

            'simpan data header
            DTA.Delete()
            For Each rowDGV As DataGridViewRow In Me.dgvLA_Header.Rows
                If rowDGV.Cells(Me.ColLA_Header_PrimaryUser.Index).Value <> -1 Then
                    If rowDGV.Cells(Me.ColLA_Header_AccessType.Index).Value = "Full View" Then
                        TipeAccess = AccModule.ClsTransaksiJournal.EnumJournalAccesType.FullAccess
                    Else
                        TipeAccess = AccModule.ClsTransaksiJournal.EnumJournalAccesType.Selected_Users_Only
                    End If
                    DTA.Insert(rowDGV.Cells(Me.ColLA_Header_PrimaryUser.Index).Value, CInt(TipeAccess))
                End If
            Next

            'simpan data item
            DTA1.Delete()
            For Each Me.rowDetailSetting In Me.TBLDetailSetting.Rows
                DTA1.Insert(Me.rowDetailSetting.primary_user_ref, Me.rowDetailSetting.primary_user)
            Next

            'simpan data setting
            If Me.chkActivateLedgerAccess.Checked = True Then
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.Jurnal_IsLedgerAccessActive, 1)
            Else
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.Jurnal_IsLedgerAccessActive, 0)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
        End Try
    End Sub

    Private Sub SalinKeField_LedgerAccess()
        Dim DTA As New DseSettingAP_AR_LedgerAccesTableAdapters.tbltrans_setting_ledger_acces1TableAdapter
        Dim TBL As New DseSettingAP_AR_LedgerAcces.tbltrans_setting_ledger_acces1DataTable
        Dim row As DseSettingAP_AR_LedgerAcces.tbltrans_setting_ledger_acces1Row
        Dim DTA1 As New DseSettingAP_AR_LedgerAccesTableAdapters.tbltrans_setting_ledger_access_item1TableAdapter
        Dim rowDGV As DataGridViewRow
        Dim Dummy As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = Core_Module.GetConstring

            'item
            DTA1.Fill(Me.TBLDetailSetting)

            'header
            DTA.Fill(TBL)
            Me.dgvLA_Header.AllowUserToAddRows = False
            Me.dgvLA_Header.Rows.Clear()
            Me.dgvLA_Header.ClearSelection()
            For Each row In TBL.Rows
                Me.dgvLA_Header.Rows.Add()
                rowDGV = Me.dgvLA_Header.Rows(Me.dgvLA_Header.Rows.Count - 1)
                If row.tipe_access = CInt(AccModule.ClsTransaksiJournal.EnumJournalAccesType.FullAccess) Then
                    rowDGV.Cells(Me.ColLA_Header_AccessType.Index).Value = "Full View"
                ElseIf row.tipe_access = CInt(AccModule.ClsTransaksiJournal.EnumJournalAccesType.Selected_Users_Only) Then
                    rowDGV.Cells(Me.ColLA_Header_AccessType.Index).Value = "Selected Users Only"
                End If
                rowDGV.Cells(Me.ColLA_Header_ColUserCode.Index).Value = row.Kode_User
                rowDGV.Cells(Me.ColLA_Header_PrimaryUser.Index).Value = row.primary_user
                rowDGV.Cells(Me.ColLA_Header_UserName.Index).Value = row.Nama_User
            Next
            Me.dgvLA_Header.AllowUserToAddRows = True

            If Me.dgvLA_Header.Rows.Count > 0 Then
                Me.dgvLA_Header.Rows(0).Selected = True
            End If

            'setting
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Jurnal_IsLedgerAccessActive)
            If Dummy = "" Or Dummy = "0" Then
                Me.chkActivateLedgerAccess.Checked = False
            Else
                Me.chkActivateLedgerAccess.Checked = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTA1.Dispose()
        End Try
    End Sub

    Private Sub dgvLA_Header_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvLA_Header.UserDeletedRow
        Try
            Me.SettinglocationAccessBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvLA_Header_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvLA_Header.UserDeletingRow
        Dim PrimaryUserRef As Double = -1

        Try
            'hapuskan semua detail item yang merujuk pada primary user ybs
            PrimaryUserRef = Me.dgvLA_Header.Item(Me.ColLA_Header_PrimaryUser.Index, e.Row.Index).Value
            For ij As Integer = Me.TBLDetailSetting.Rows.Count - 1 To 0 Step -1
                Me.rowDetailSetting = TBLDetailSetting.Rows(ij)
                If Me.rowDetailSetting.primary_user_ref = PrimaryUserRef Then
                    Me.TBLDetailSetting.Rows.Remove(Me.rowDetailSetting)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvLA_Item_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvLA_Item.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColLA_Item_UserCode.Index Then
                e.Cancel = True
                Exit Sub
            Else
                'col user code
                If e.ColumnIndex = Me.ColLA_Item_UserCode.Index Then
                    Me.Buffer = Me.dgvLA_Item.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem(ByVal vUserCode As String, ByVal vIndexRef As Integer) As Integer
        Dim rowDGV As DataGridViewRow
        Dim RetVal As Integer = -1

        Try
            For ij As Integer = 0 To Me.dgvLA_Item.RowCount - 2
                If ij <> vIndexRef Then
                    rowDGV = Me.dgvLA_Item.Rows(ij)
                    If rowDGV.Cells(Me.ColLA_Item_UserCode.Index).Value = vUserCode Then
                        RetVal = rowDGV.Index
                        Exit For
                    End If
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetPrimaryMax() As Double
        Dim rowDetail As DseSettingAP_AR_LedgerAcces.tbltrans_setting_ledger_access_item1Row
        Dim RetVal As Double = -1

        Try
            For Each rowDetail In Me.TBLDetailSetting.Rows
                If rowDetail.primary > RetVal Then
                    RetVal = rowDetail.primary
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub UpdateTabelDetail(ByVal vUserCodeBaru As String, ByVal vNamaUserBaru As String, ByVal vUserCodeLama As String, ByVal vPrimaryUserRef As Double)
        Try
            If vUserCodeLama = "" Then
                'tambahkan baris baru
                Me.rowDetailSetting = Me.TBLDetailSetting.Newtbltrans_setting_ledger_access_item1Row
                With Me.rowDetailSetting
                    .Kode_User = vUserCodeBaru
                    .Nama_User = vNamaUserBaru
                    .primary = Me.GetPrimaryMax + 1
                    .primary_user = Me.Objuser.GetPrimary(vUserCodeBaru)
                    .primary_user_ref = vPrimaryUserRef
                End With
                Me.TBLDetailSetting.Rows.Add(Me.rowDetailSetting)
            Else
                'ganti baris lama, dengan kode user baru
                For Each Me.rowDetailSetting In Me.TBLDetailSetting.Rows
                    If Me.rowDetailSetting.Kode_User = vUserCodeLama And Me.rowDetailSetting.primary_user_ref = vPrimaryUserRef Then
                        With Me.rowDetailSetting
                            .Kode_User = vUserCodeBaru
                            .Nama_User = vNamaUserBaru
                            .primary_user = Me.Objuser.GetPrimary(vUserCodeBaru)
                        End With
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvLA_Item_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvLA_Item.CellEndEdit
        Dim rowHeader As DataGridViewRow

        Try
            Control.CheckForIllegalCrossThreadCalls = False

            If IsNothing(Me.dgvLA_Item.Item(e.ColumnIndex, e.RowIndex).Value) = True _
            Or Me.dgvLA_Item.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                Me.dgvLA_Item.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            Else
                'user code
                If e.ColumnIndex = Me.ColLA_Item_UserCode.Index Then
                    Me.dgvLA_Item.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvLA_Item.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    'kemungkinan
                    If Me.Objuser.ApakahSudahAda(Me.dgvLA_Item.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        '1. kode user yang diketik ada --> load data deskripsi + primary user
                        Me.dgvLA_Item.Item(Me.ColLA_Item_UserName.Index, e.RowIndex).Value = Me.Objuser.GetNama(Me.dgvLA_Item.Item(e.ColumnIndex, e.RowIndex).Value)
                        'cek duplication
                        If Me.GetIndexItem(Me.dgvLA_Item.Item(Me.ColLA_Item_UserCode.Index, e.RowIndex).Value, e.RowIndex) <> -1 Then
                            MsgBox("User code duplication, please select another user.", MsgBoxStyle.Information, "Information")
                            If Me.Buffer <> "" Then
                                Me.dgvLA_Item.Item(e.ColumnIndex, e.RowIndex).Value = Buffer
                            Else
                                Me.rowIndexLA_I = e.RowIndex
                                Me.BgwHapusRow_LA_Item.RunWorkerAsync()
                            End If
                            Exit Sub
                        End If
                        'update tabel setting
                        rowHeader = Me.dgvLA_Header.Rows(Me.dgvLA_Header.SelectedCells(0).RowIndex)
                        Me.UpdateTabelDetail(Me.dgvLA_Item.Item(Me.ColLA_Item_UserCode.Index, e.RowIndex).Value, _
                         Me.dgvLA_Item.Item(Me.ColLA_Item_UserName.Index, e.RowIndex).Value, Me.Buffer, _
                         rowHeader.Cells(Me.ColLA_Header_PrimaryUser.Index).Value)

                        Me.SettinglocationAccessBerubah = True
                    Else
                        '2. kode user yang diketik tidak ada --> tampilkan form cari
                        Dim ij As New Core_Module.FrmCariMasterUser
                        ij.TxtCari.Text = Me.dgvLA_Item.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.Kodeuser <> "" Then
                            '2.a. kode user yang dicari ada --> load data deskripsi + primary user + kode user
                            Me.dgvLA_Item.Item(e.ColumnIndex, e.RowIndex).Value = ij.Kodeuser
                            Me.dgvLA_Item.Item(Me.ColLA_Item_UserName.Index, e.RowIndex).Value = Me.Objuser.GetNama(Me.dgvLA_Item.Item(e.ColumnIndex, e.RowIndex).Value)

                            'cek duplication
                            If Me.GetIndexItem(Me.dgvLA_Item.Item(Me.ColLA_Item_UserCode.Index, e.RowIndex).Value, e.RowIndex) <> -1 Then
                                MsgBox("User code duplication, please select another user.", MsgBoxStyle.Information, "Information")
                                If Me.Buffer <> "" Then
                                    Me.dgvLA_Item.Item(e.ColumnIndex, e.RowIndex).Value = Buffer
                                Else
                                    Me.rowIndexLA_I = e.RowIndex
                                    Me.BgwHapusRow_LA_Item.RunWorkerAsync()
                                End If
                                Exit Sub
                            End If
                            'update nilai primary user ref ke user yang baru
                            rowHeader = Me.dgvLA_Header.Rows(Me.dgvLA_Header.SelectedCells(0).RowIndex)
                            Me.UpdateTabelDetail(Me.dgvLA_Item.Item(Me.ColLA_Item_UserCode.Index, e.RowIndex).Value, _
                             Me.dgvLA_Item.Item(Me.ColLA_Item_UserName.Index, e.RowIndex).Value, Me.Buffer, _
                             rowHeader.Cells(Me.ColLA_Header_PrimaryUser.Index).Value)

                            Me.SettinglocationAccessBerubah = True
                        Else
                            '2.b. kode user yang dicari tidak ada --> kembalikan ke kondisi buffer
                            Me.dgvLA_Item.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If Me.dgvLA_Item.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                Me.rowIndexLA_I = e.RowIndex
                                Me.BgwHapusRow_LA_Item.RunWorkerAsync()
                                ' Me.dgvLA_Item.Rows.RemoveAt(e.RowIndex)
                                Exit Sub
                            End If
                        End If
                        ij = Nothing
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvLA_Item_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvLA_Item.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgvLA_Item.Rows(Me.dgvLA_Item.Rows.Count - 1)
            rowDGV.Cells(Me.ColLA_Item_UserCode.Index).Value = ""
            rowDGV.Cells(Me.ColLA_Item_UserName.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvLA_Item_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvLA_Item.UserDeletedRow
        Try
            Me.SettinglocationAccessBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvLA_Item_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvLA_Item.UserDeletingRow
        Dim rowHeader As DataGridViewRow
        Dim PrimaryUserRef As Double = -1

        Try
            rowHeader = Me.dgvLA_Header.Rows(Me.dgvLA_Header.SelectedCells(0).RowIndex)
            PrimaryUserRef = rowHeader.Cells(Me.ColLA_Header_PrimaryUser.Index).Value

            For Each Me.rowDetailSetting In Me.TBLDetailSetting.Rows
                If Me.rowDetailSetting.primary_user_ref = PrimaryUserRef And _
                Me.rowDetailSetting.Kode_User = Me.dgvLA_Item.Item(Me.ColLA_Item_UserCode.Index, e.Row.Index).Value Then
                    Me.TBLDetailSetting.Rows.Remove(Me.rowDetailSetting)
                    Exit Sub
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwHapusRow_LA_Item_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwHapusRow_LA_Item.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub BgwHapusRow_LA_Item_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwHapusRow_LA_Item.RunWorkerCompleted
        Try
            Me.dgvLA_Item.Rows.RemoveAt(Me.rowIndexLA_I)
            Me.rowIndexLA_I = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
