Option Explicit On
Option Strict Off

Imports AccModule

Partial Public Class FrmSettingAP_AR
    'catatan:
    '1. gain account/loss account: income, other income, expense, other expense, cogs --> jika bukan maka berikan peringatan --> cek pada cell end edit
    '2. simpan langsung per baris, update perubahan per baris langsung pada cell end edit, jika itemnya benar
    '3. jika selesai memasukkan data akun, langsung masukkan akun expensen-nya secara default ambil akun expense pertama yang aktif dan non-header
    '4. ketika hapus, tidak boleh ada stock in out yang merujuk pada tipe adjustment tersebut
    '5. kolom yg boleh diedit: code, desc,remarks, gain account, loss account
    '6. kecuali remarks, semua tidak boleh kosong
    '7. ketika load form langsung load semua data yang ada di dalam database
    Dim idx As Integer = -1

    Public Sub IsiAccExpensePertama(ByVal idx As Integer)
        Dim dta As New DseSettingAdjTableAdapters.tblmaster_akunTableAdapter
        Dim tbl As New DseSettingAdj.tblmaster_akunDataTable
        Dim row As DseSettingAdj.tblmaster_akunRow

        Try
            dta.Fill(tbl, ClsMasterAkun.EnumTipeAkun.Expense)
            If tbl.Rows.Count > 0 Then
                row = tbl.Rows(0)
                Me.dgvAdjustmentType.Item(Me.ColPrimaryGainAccount.Index, idx).Value = row.Primary
                Me.dgvAdjustmentType.Item(Me.ColGainAccount.Index, idx).Value = row.Account_Code
                Me.dgvAdjustmentType.Item(Me.ColDescGain.Index, idx).Value = row.Description
                Me.dgvAdjustmentType.Item(Me.ColPrimaryLossAccount.Index, idx).Value = row.Primary
                Me.dgvAdjustmentType.Item(Me.ColLossAccount.Index, idx).Value = row.Account_Code
                Me.dgvAdjustmentType.Item(Me.ColDescLoss.Index, idx).Value = row.Description
            Else
                Me.dgvAdjustmentType.Item(Me.ColPrimaryGainAccount.Index, idx).Value = -1
                Me.dgvAdjustmentType.Item(Me.ColGainAccount.Index, idx).Value = ""
                Me.dgvAdjustmentType.Item(Me.ColDescGain.Index, idx).Value = ""
                Me.dgvAdjustmentType.Item(Me.ColPrimaryLossAccount.Index, idx).Value = -1
                Me.dgvAdjustmentType.Item(Me.ColLossAccount.Index, idx).Value = ""
                Me.dgvAdjustmentType.Item(Me.ColDescLoss.Index, idx).Value = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            tbl.Dispose()
        End Try
    End Sub

    Private Function ApakahBolehHapus(ByVal vPrimary As Double) As Boolean
        Dim dta As New DseSettingAdjTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim jml As Nullable(Of Long)

        Try
            jml = dta.ScalarQuery_GetJmlByPrimaryAdj(vPrimary)
            If jml.HasValue = True Then
                If jml.Value > 0 Then
                    MsgBox("Cannot delete this data, it's refered by stock in adjusment or stock out adjusment transaction.", _
                    MsgBoxStyle.Information, "Error")
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Private Function CekData(ByVal RowIdx As Integer) As Boolean
        Try
            If Me.dgvAdjustmentType.Item(Me.ColCode.Index, RowIdx).Value <> "" Then
                'desc tidak boleh kosong
                If Me.dgvAdjustmentType.Item(Me.ColDesc.Index, RowIdx).Value = "" Then
                    MsgBox("Please input adjustment type description.", MsgBoxStyle.Information, "Information")
                    Me.dgvAdjustmentType.Item(Me.ColDesc.Index, RowIdx).Selected = True
                    Return False
                End If
                'gain acc
                If Me.dgvAdjustmentType.Item(Me.ColPrimaryGainAccount.Index, RowIdx).Value < 1 Then
                    MsgBox("Please input gain account.", MsgBoxStyle.Information, "Information")
                    Me.dgvAdjustmentType.Item(Me.ColGainAccount.Index, RowIdx).Selected = True
                    Return False
                End If
                'loss acc
                If Me.dgvAdjustmentType.Item(Me.ColPrimaryLossAccount.Index, RowIdx).Value < 1 Then
                    MsgBox("Please input loss account.", MsgBoxStyle.Information, "Information")
                    Me.dgvAdjustmentType.Item(Me.ColLossAccount.Index, RowIdx).Selected = True
                    Return False
                End If
                Return True
            Else
                If Me.dgvAdjustmentType.Item(Me.ColPrimary.Index, RowIdx).Value > 0 Then
                    MsgBox("Please input adjustment type code.", MsgBoxStyle.Information, "Information")
                    Me.dgvAdjustmentType.Item(Me.ColCode.Index, RowIdx).Selected = True
                    Return False
                End If
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function CekCodeSama(ByVal vCode As String, ByVal vPrimary As Double) As Boolean
        Dim dta As New DseSettingAdjTableAdapters.tbl_setting_adjustmentTableAdapter
        Dim CodeLama As String = ""
        Dim jml As Nullable(Of Long)
        Try
            CodeLama = dta.ScalarQuery_GetCodeByPrimary(vPrimary)
            If String.IsNullOrEmpty(vCode) = True Then
                vCode = ""
            End If
            If CodeLama <> vCode Then
                jml = dta.ScalarQuery_GetJmlByCode(vCode)
                If jml.HasValue = True Then
                    If jml.Value > 0 Then
                        MsgBox("Duplication in code, please change code", MsgBoxStyle.Information, "Information")
                        Return True
                    Else
                        Return False
                    End If
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Private Sub HapusAdj(ByVal vPrimary As Double)
        Dim Dta As New DseSettingAdjTableAdapters.tbl_setting_adjustmentTableAdapter
        Try
            Dta.Delete(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Sub

    Private Sub Simpan(ByVal RowIndex As Integer)
        Dim Dta As New DseSettingAdjTableAdapters.tbl_setting_adjustmentTableAdapter
        Dim remark As String = ""
        Try
            If Me.dgvAdjustmentType.Item(Me.ColCode.Index, RowIndex).Value <> "" Then
                If String.IsNullOrEmpty(Me.dgvAdjustmentType.Item(Me.ColRemarks.Index, RowIndex).Value) = False Then
                    remark = Me.dgvAdjustmentType.Item(Me.ColRemarks.Index, RowIndex).Value
                End If
                Dta.Insert(Me.dgvAdjustmentType.Item(Me.ColCode.Index, RowIndex).Value, _
                Me.dgvAdjustmentType.Item(Me.ColDesc.Index, RowIndex).Value, _
                Me.dgvAdjustmentType.Item(Me.ColPrimaryGainAccount.Index, RowIndex).Value, _
                Me.dgvAdjustmentType.Item(Me.ColPrimaryLossAccount.Index, RowIndex).Value, _
                remark)

                Me.dgvAdjustmentType.Item(Me.ColPrimary.Index, RowIndex).Value = _
                Dta.ScalarQuery_GetPrimaryByCode(Me.dgvAdjustmentType.Item(Me.ColCode.Index, RowIndex).Value)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Sub

    Private Sub Edit(ByVal RowIndex As Integer)
        Dim Dta As New DseSettingAdjTableAdapters.tbl_setting_adjustmentTableAdapter
        Dim remark As String = ""
        Try
            If String.IsNullOrEmpty(Me.dgvAdjustmentType.Item(Me.ColRemarks.Index, RowIndex).Value) = False Then
                remark = Me.dgvAdjustmentType.Item(Me.ColRemarks.Index, RowIndex).Value
            End If
            Dta.Update(Me.dgvAdjustmentType.Item(Me.ColCode.Index, RowIndex).Value, _
            Me.dgvAdjustmentType.Item(Me.ColDesc.Index, RowIndex).Value, _
            Me.dgvAdjustmentType.Item(Me.ColPrimaryGainAccount.Index, RowIndex).Value, _
            Me.dgvAdjustmentType.Item(Me.ColPrimaryLossAccount.Index, RowIndex).Value, _
            remark, Me.dgvAdjustmentType.Item(Me.ColPrimary.Index, RowIndex).Value)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Sub

    Private Sub FrmTrans_SettingAdjustment_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim dta As New DseSettingAdjTableAdapters.tbl_setting_adjustmentTableAdapter
        Dim tbl As New DseSettingAdj.tbl_setting_adjustmentDataTable
        Dim row As DseSettingAdj.tbl_setting_adjustmentRow
        Dim rowdgv As DataGridViewRow
        Try
            'sementara
            Me.TabControl4.Visible = False


            Me.dgvAdjustmentType.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvRateFinal.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            dta.Fill(tbl)
            For Each row In tbl.Rows
                Me.dgvAdjustmentType.Rows.Add()
                rowdgv = Me.dgvAdjustmentType.Rows(Me.dgvAdjustmentType.Rows.Count - 2)
                rowdgv.Cells(Me.ColCode.Index).Value = row.Code
                rowdgv.Cells(Me.ColDesc.Index).Value = row.Desc
                rowdgv.Cells(Me.ColGainAccount.Index).Value = Me.ObjAkun.GetAccountCode(row.PrimaryGain)
                rowdgv.Cells(Me.ColDescGain.Index).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(row.PrimaryGain)
                rowdgv.Cells(Me.ColLossAccount.Index).Value = Me.ObjAkun.GetAccountCode(row.PrimaryLoss)
                rowdgv.Cells(Me.ColDescLoss.Index).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(row.PrimaryLoss)
                rowdgv.Cells(Me.ColRemarks.Index).Value = row.Remarks
                rowdgv.Cells(Me.ColPrimary.Index).Value = row.Primary
                rowdgv.Cells(Me.ColPrimaryGainAccount.Index).Value = row.PrimaryGain
                rowdgv.Cells(Me.ColPrimaryLossAccount.Index).Value = row.PrimaryLoss
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAdjustmentType_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvAdjustmentType.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColCode.Index Then
                If Me.dgvAdjustmentType(Me.ColCode.Index, e.RowIndex).Value = "" Then
                    e.Cancel = True
                Else
                    If String.IsNullOrEmpty(Me.dgvAdjustmentType(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.Buffer = ""
                    Else
                        Me.Buffer = Me.dgvAdjustmentType(e.ColumnIndex, e.RowIndex).Value
                    End If
                End If
            Else
                Me.Buffer = Me.dgvAdjustmentType(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAdjustmentType_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvAdjustmentType.CellEndEdit
        Dim hasil As DialogResult
        Try
            'col code
            If e.ColumnIndex = Me.ColCode.Index Then
                'jika kosong kasih peringatan hapus data atau jangan
                If Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                    If Me.dgvAdjustmentType.Item(Me.ColDesc.Index, e.RowIndex).Value <> "" Or _
                    Me.dgvAdjustmentType.Item(Me.ColPrimaryGainAccount.Index, e.RowIndex).Value > 0 Or _
                    Me.dgvAdjustmentType.Item(Me.ColPrimaryLossAccount.Index, e.RowIndex).Value > 0 Then
                        hasil = MsgBox("Delete data ?", MsgBoxStyle.YesNo, "Delete")
                        If hasil = Windows.Forms.DialogResult.Yes Then
                            'hapus
                            If Me.ApakahBolehHapus(Me.dgvAdjustmentType.Item(Me.ColPrimary.Index, e.RowIndex).Value) = True Then
                                Me.HapusAdj(Me.dgvAdjustmentType.Item(Me.ColPrimary.Index, e.RowIndex).Value)
                                Me.idx = e.RowIndex
                                Me.BgwHapusRow.RunWorkerAsync()
                                Me.BgwHapusRow.Dispose()
                            End If
                        Else
                            If Me.CekCodeSama(Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value, _
                            Me.dgvAdjustmentType.Item(Me.ColPrimary.Index, e.RowIndex).Value) = True Then
                                Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                If Me.Buffer = "" Then
                                    Me.idx = e.RowIndex
                                    Me.BgwHapusRow.RunWorkerAsync()
                                    Me.BgwHapusRow.Dispose()
                                End If
                            End If
                        End If
                    Else
                        Me.idx = e.RowIndex
                        Me.BgwHapusRow.RunWorkerAsync()
                        Me.BgwHapusRow.Dispose()
                    End If
                Else
                    If Me.CekCodeSama(Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value, _
                    Me.dgvAdjustmentType.Item(Me.ColPrimary.Index, e.RowIndex).Value) = False Then
                        Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value = _
                        Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper()
                        If String.IsNullOrEmpty(Me.dgvAdjustmentType.Item(Me.ColPrimary.Index, e.RowIndex).Value) = True And _
                        String.IsNullOrEmpty(Me.dgvAdjustmentType.Item(Me.ColPrimaryGainAccount.Index, e.RowIndex).Value) = True And _
                        String.IsNullOrEmpty(Me.dgvAdjustmentType.Item(Me.ColPrimaryLossAccount.Index, e.RowIndex).Value) = True Then
                            Me.dgvAdjustmentType.Item(Me.ColDesc.Index, e.RowIndex).Value = "Account type : " + Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value
                            Me.IsiAccExpensePertama(e.RowIndex)
                        End If
                    Else
                        If Me.dgvAdjustmentType.Item(Me.ColPrimary.Index, e.RowIndex).Value > 0 Then
                            Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Else
                            Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value = ""
                            Me.idx = e.RowIndex
                            Me.BgwHapusRow.RunWorkerAsync()
                            Me.BgwHapusRow.Dispose()
                        End If
                    End If
                End If
            End If
            'gain acc
            If e.ColumnIndex = Me.ColGainAccount.Index Then
                If Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value <> "" Then
                    If Me.ObjAkun.ApakahSudahAda(Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value) Then
                        If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value) Then
                            'cek type akun
                            '1. gain account/loss account: income, other income, expense, other expense, cogs --> jika bukan maka berikan peringatan --> cek pada cell end edit
                            Dim TypeAkun As ClsMasterAkun.EnumTipeAkun
                            Dim PrimaryAcc As Double = -1
                            PrimaryAcc = Me.ObjAkun.GetPrimary(Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value)
                            TypeAkun = Me.ObjAkun.GetTipeAkun(PrimaryAcc)
                            If TypeAkun = ClsMasterAkun.EnumTipeAkun.Income Or TypeAkun = ClsMasterAkun.EnumTipeAkun.Other_Income _
                            Or TypeAkun = ClsMasterAkun.EnumTipeAkun.Expense Or TypeAkun = ClsMasterAkun.EnumTipeAkun.Other_Expense _
                            Or TypeAkun = ClsMasterAkun.EnumTipeAkun.Cost_of_Goods_Sales Then
                                
                            Else
                                MsgBox("Please use one the following account type: income, other income, expense, other expense or COGS account.", MsgBoxStyle.Information, "Information")
                            End If
                            'primary
                            Me.dgvAdjustmentType.Item(Me.ColPrimaryGainAccount.Index, e.RowIndex).Value = PrimaryAcc
                            'kode
                            Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAkun.GetAccountCode(PrimaryAcc)
                            'desc
                            Me.dgvAdjustmentType.Item(Me.ColDescGain.Index, e.RowIndex).Value = _
                            Me.ObjAkun.GetDescription_ByPrimaryAcc(PrimaryAcc)
                        Else
                            MsgBox("Please input active and non header acccount.", MsgBoxStyle.Information, "Information")
                            'kode
                            Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            'primary
                            Me.dgvAdjustmentType.Item(Me.ColPrimaryGainAccount.Index, e.RowIndex).Value = _
                            Me.ObjAkun.GetPrimary(Me.Buffer)
                            'desc
                            Me.dgvAdjustmentType.Item(Me.ColDescGain.Index, e.RowIndex).Value = _
                            Me.ObjAkun.GetDescription(Me.Buffer)
                        End If
                    Else
                        Dim frm As New FrmMasterAkun_Cari
                        Dim CodeAcc As String = ""
                        frm.txtCari.Text = Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value
                        frm.ShowDialog()
                        CodeAcc = Me.ObjAkun.GetAccountCode(frm.mPrimaryAkun)
                        If frm.mPrimaryAkun <> -1 Then
                            If ObjAkun.ApakahSudahAda_Aktif_NonHeader(CodeAcc) Then
                                Dim TypeAkun As ClsMasterAkun.EnumTipeAkun
                                TypeAkun = Me.ObjAkun.GetTipeAkun(frm.mPrimaryAkun)
                                If TypeAkun = ClsMasterAkun.EnumTipeAkun.Income Or TypeAkun = ClsMasterAkun.EnumTipeAkun.Other_Income _
                                Or TypeAkun = ClsMasterAkun.EnumTipeAkun.Expense Or TypeAkun = ClsMasterAkun.EnumTipeAkun.Other_Expense _
                                Or TypeAkun = ClsMasterAkun.EnumTipeAkun.Cost_of_Goods_Sales Then
                                
                                Else
                                    MsgBox("Please use one the following account type: income, other income, expense, other expense or COGS account.", MsgBoxStyle.Information, "Information")
                                    
                                End If
                                'primary
                                Me.dgvAdjustmentType.Item(Me.ColPrimaryGainAccount.Index, e.RowIndex).Value = frm.mPrimaryAkun
                                'kode
                                Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value = CodeAcc
                                'desc
                                Me.dgvAdjustmentType.Item(Me.ColDescGain.Index, e.RowIndex).Value = _
                                Me.ObjAkun.GetDescription_ByPrimaryAcc(frm.mPrimaryAkun)
                            Else
                                MsgBox("Please input active and non header acccount.", MsgBoxStyle.Information, "Information")
                                'kode
                                Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                'primary
                                Me.dgvAdjustmentType.Item(Me.ColPrimaryGainAccount.Index, e.RowIndex).Value = _
                                Me.ObjAkun.GetPrimary(Me.Buffer)
                                'desc
                                Me.dgvAdjustmentType.Item(Me.ColDescGain.Index, e.RowIndex).Value = _
                                Me.ObjAkun.GetDescription(Me.Buffer)
                            End If
                        Else
                            'kode
                            Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            'primary
                            Me.dgvAdjustmentType.Item(Me.ColPrimaryGainAccount.Index, e.RowIndex).Value = _
                            Me.ObjAkun.GetPrimary(Me.Buffer)
                            'desc
                            Me.dgvAdjustmentType.Item(Me.ColDescGain.Index, e.RowIndex).Value = _
                            Me.ObjAkun.GetDescription(Me.Buffer)
                        End If
                    End If
                Else
                    'kode
                    Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    'primary
                    Me.dgvAdjustmentType.Item(Me.ColPrimaryGainAccount.Index, e.RowIndex).Value = _
                    Me.ObjAkun.GetPrimary(Me.Buffer)
                    'desc
                    Me.dgvAdjustmentType.Item(Me.ColDescGain.Index, e.RowIndex).Value = _
                    Me.ObjAkun.GetDescription(Me.Buffer)
                End If
            End If
            'loss acc
            If e.ColumnIndex = Me.ColLossAccount.Index Then
                If Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value <> "" Then
                    If Me.ObjAkun.ApakahSudahAda(Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value) Then
                        If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value) Then
                            'cek type akun
                            '1. gain account/loss account: income, other income, expense, other expense, cogs --> jika bukan maka berikan peringatan --> cek pada cell end edit
                            Dim TypeAkun As ClsMasterAkun.EnumTipeAkun
                            Dim PrimaryAcc As Double = -1
                            PrimaryAcc = Me.ObjAkun.GetPrimary(Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value)
                            TypeAkun = Me.ObjAkun.GetTipeAkun(PrimaryAcc)
                            If TypeAkun = ClsMasterAkun.EnumTipeAkun.Income Or TypeAkun = ClsMasterAkun.EnumTipeAkun.Other_Income _
                            Or TypeAkun = ClsMasterAkun.EnumTipeAkun.Expense Or TypeAkun = ClsMasterAkun.EnumTipeAkun.Other_Expense _
                            Or TypeAkun = ClsMasterAkun.EnumTipeAkun.Cost_of_Goods_Sales Then
                            Else
                                MsgBox("Please use one the following account type: income, other income, expense, other expense or COGS account.", MsgBoxStyle.Information, "Information")
                            End If
                            'primary
                            Me.dgvAdjustmentType.Item(Me.ColPrimaryLossAccount.Index, e.RowIndex).Value = PrimaryAcc
                            'kode
                            Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAkun.GetAccountCode(PrimaryAcc)
                            'desc
                            Me.dgvAdjustmentType.Item(Me.ColDescLoss.Index, e.RowIndex).Value = _
                            Me.ObjAkun.GetDescription_ByPrimaryAcc(PrimaryAcc)
                        Else
                            MsgBox("Please input active and non header acccount.", MsgBoxStyle.Information, "Information")
                            'kode
                            Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            'primary
                            Me.dgvAdjustmentType.Item(Me.ColPrimaryLossAccount.Index, e.RowIndex).Value = _
                            Me.ObjAkun.GetPrimary(Me.Buffer)
                            'desc
                            Me.dgvAdjustmentType.Item(Me.ColDescLoss.Index, e.RowIndex).Value = _
                            Me.ObjAkun.GetDescription(Me.Buffer)
                        End If
                    Else
                        Dim frm As New FrmMasterAkun_Cari
                        Dim CodeAcc As String = ""
                        frm.txtCari.Text = Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value
                        frm.ShowDialog()
                        CodeAcc = Me.ObjAkun.GetAccountCode(frm.mPrimaryAkun)
                        If frm.mPrimaryAkun <> -1 Then
                            If ObjAkun.ApakahSudahAda_Aktif_NonHeader(CodeAcc) Then
                                Dim TypeAkun As ClsMasterAkun.EnumTipeAkun
                                TypeAkun = Me.ObjAkun.GetTipeAkun(frm.mPrimaryAkun)
                                If TypeAkun = ClsMasterAkun.EnumTipeAkun.Income Or TypeAkun = ClsMasterAkun.EnumTipeAkun.Other_Income _
                                Or TypeAkun = ClsMasterAkun.EnumTipeAkun.Expense Or TypeAkun = ClsMasterAkun.EnumTipeAkun.Other_Expense _
                                Or TypeAkun = ClsMasterAkun.EnumTipeAkun.Cost_of_Goods_Sales Then
                                Else
                                    MsgBox("Please use one the following account type: income, other income, expense, other expense or COGS account.", MsgBoxStyle.Information, "Information")
                                End If
                                'primary
                                Me.dgvAdjustmentType.Item(Me.ColPrimaryLossAccount.Index, e.RowIndex).Value = frm.mPrimaryAkun
                                'kode
                                Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value = CodeAcc
                                'desc
                                Me.dgvAdjustmentType.Item(Me.ColDescLoss.Index, e.RowIndex).Value = _
                                Me.ObjAkun.GetDescription_ByPrimaryAcc(frm.mPrimaryAkun)
                            Else
                                MsgBox("Please input active and non header acccount.", MsgBoxStyle.Information, "Information")
                                'kode
                                Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                'primary
                                Me.dgvAdjustmentType.Item(Me.ColPrimaryGainAccount.Index, e.RowIndex).Value = _
                                Me.ObjAkun.GetPrimary(Me.Buffer)
                                'desc
                                Me.dgvAdjustmentType.Item(Me.ColDescLoss.Index, e.RowIndex).Value = _
                                Me.ObjAkun.GetDescription(Me.Buffer)
                            End If
                        Else
                            'kode
                            Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            'primary
                            Me.dgvAdjustmentType.Item(Me.ColPrimaryGainAccount.Index, e.RowIndex).Value = _
                            Me.ObjAkun.GetPrimary(Me.Buffer)
                            'desc
                            Me.dgvAdjustmentType.Item(Me.ColDescLoss.Index, e.RowIndex).Value = _
                            Me.ObjAkun.GetDescription(Me.Buffer)
                        End If
                    End If
                Else
                    'kode
                    Me.dgvAdjustmentType.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    'primary
                    Me.dgvAdjustmentType.Item(Me.ColPrimaryGainAccount.Index, e.RowIndex).Value = _
                    Me.ObjAkun.GetPrimary(Me.Buffer)
                    'desc
                    Me.dgvAdjustmentType.Item(Me.ColDescLoss.Index, e.RowIndex).Value = _
                    Me.ObjAkun.GetDescription(Me.Buffer)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwHapusRow_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwHapusRow.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwHapusRow_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwHapusRow.RunWorkerCompleted
        Try
            Me.dgvAdjustmentType.Rows.RemoveAt(Me.idx)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAdjustmentType_RowLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvAdjustmentType.RowLeave
        'Try
        '    If e.RowIndex <> Me.dgvAdjustmentType.Rows.Count - 1 Then
        '        Me.dgvAdjustmentType.ClearSelection()
        '        If Me.CekData(e.RowIndex) = True Then
        '            'simpan data baris
        '            If Me.dgvAdjustmentType.Item(Me.ColPrimary.Index, e.RowIndex).Value > 0 Then
        '                'update
        '                Me.Edit(e.RowIndex)
        '            Else
        '                'simpan baru
        '                Me.Simpan(e.RowIndex)
        '            End If
        '        End If
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub dgvAdjustmentType_RowValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvAdjustmentType.RowValidating
        Try
            If e.RowIndex <> Me.dgvAdjustmentType.Rows.Count - 1 Then
                Me.dgvAdjustmentType.ClearSelection()
                If Me.CekData(e.RowIndex) = False Then
                    e.Cancel = True
                Else
                    'simpan data baris
                    If Me.dgvAdjustmentType.Item(Me.ColPrimary.Index, e.RowIndex).Value > 0 Then
                        'update
                        Me.Edit(e.RowIndex)
                    Else
                        'simpan baru
                        Me.Simpan(e.RowIndex)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAdjustmentType_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvAdjustmentType.UserDeletingRow
        Dim hasil As DialogResult
        Try
            hasil = MsgBox("Delete data ?", MsgBoxStyle.YesNo)
            If hasil = Windows.Forms.DialogResult.Yes Then
                If Me.ApakahBolehHapus(Me.dgvAdjustmentType.Item(Me.ColPrimary.Index, e.Row.Index).Value) = True Then
                    Me.HapusAdj(Me.dgvAdjustmentType.Item(Me.ColPrimary.Index, e.Row.Index).Value)
                Else
                    e.Cancel = True
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
