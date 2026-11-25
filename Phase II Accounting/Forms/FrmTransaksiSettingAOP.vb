Option Explicit On
Option Strict Off

Public Class FrmTransaksiSettingAOP
    Public BackColorAll As Color = Color.LightSeaGreen
    Public UserCode As String = ""

    Dim Buffer As String = ""
    Dim ObjAkun As New ClsMasterAkun
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjSetting As New ClsTransaksiSettingTransf_AOP

    Dim AdaPerubahan As Boolean = False

    Private Sub FrmTransaksiSettingAOP_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvAkun.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvCost.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.cboCurrency.SelectedIndex = 0

            'tambahkan data cost
            Me.dgvCost.Rows.Add()
            rowDGV = Me.dgvCost.Rows(Me.dgvCost.Rows.Count - 1)
            With rowDGV
                .Cells(Me.ColNoCost.Index).Value = Me.dgvCost.Rows.Count
                .Cells(Me.ColDescCost.Index).Value = "Cleareance Fee"
                .Cells(Me.ColCurrencyCost.Index).Value = "EUR"
                .Cells(Me.ColCostNominal.Index).Value = "250.00"
                .Cells(Me.ColExcRateCost.Index).Value = "14,025.50"
                .Cells(Me.ColCostAmount.Index).Value = "3,506,375"
                .Cells(Me.ColCostAccountNo.Index).Value = "1.180.005"
                .Cells(Me.ColCostAccountDesc.Index).Value = "PREPAYMENT CLEAREANCE"
                .Cells(Me.ColRemarksCost.Index).Value = ""
            End With
            Me.dgvCost.Rows.Add()
            rowDGV = Me.dgvCost.Rows(Me.dgvCost.Rows.Count - 1)
            With rowDGV
                .Cells(Me.ColNoCost.Index).Value = Me.dgvCost.Rows.Count
                .Cells(Me.ColDescCost.Index).Value = "Surat Rekomendasi Depkes"
                .Cells(Me.ColCurrencyCost.Index).Value = "USD"
                .Cells(Me.ColCostNominal.Index).Value = "1000.00"
                .Cells(Me.ColExcRateCost.Index).Value = "9,015.25"
                .Cells(Me.ColCostAmount.Index).Value = "9,015,250"
                .Cells(Me.ColCostAccountNo.Index).Value = "1.180.006"
                .Cells(Me.ColCostAccountDesc.Index).Value = "PREPAYMENT CLEAREANCE (REKOM DEPKES)"
                .Cells(Me.ColRemarksCost.Index).Value = ""
            End With
            Me.dgvCost.Rows.Add()
            rowDGV = Me.dgvCost.Rows(Me.dgvCost.Rows.Count - 1)
            With rowDGV
                .Cells(Me.ColNoCost.Index).Value = Me.dgvCost.Rows.Count
                .Cells(Me.ColDescCost.Index).Value = "Asuransi"
                .Cells(Me.ColCurrencyCost.Index).Value = "IDR"
                .Cells(Me.ColCostNominal.Index).Value = "1,500,000"
                .Cells(Me.ColExcRateCost.Index).Value = "1"
                .Cells(Me.ColCostAmount.Index).Value = "1,500,000"
                .Cells(Me.ColCostAccountNo.Index).Value = "2.110.001"
                .Cells(Me.ColCostAccountDesc.Index).Value = "AP. SOMPO JAPAN INSURANCE"
                .Cells(Me.ColRemarksCost.Index).Value = ""
            End With

            Me.ObjSetting.LoadData(Me, Me.dgvItem, Me.tslCreatedBy, Me.tslCreatedDate)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCurrency_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCurrency.SelectedIndexChanged
        Try
            If Me.cboCurrency.SelectedIndex = 0 Then
                'Sales IDR
                Me.TxtTaxRate.Text = "1"
                Me.TxtPurchaseIDR.Width = Me.TxtTaxRate.Right - Me.TxtPurchaseUSD.Left
                Me.TxtPurchaseIDR.Left = Me.TxtPurchaseUSD.Left
                Me.LblUSD.Text = "IDR"
                Me.LblIDR.Visible = False
                Me.TxtPurchaseUSD.Visible = False
                Me.TxtPurchaseIDR.Text = "10,000,000"
                'VAT
                Me.TxtVatIDR.Width = Me.TxtTaxRate.Right - Me.TxtVATUSD.Left
                Me.TxtVatIDR.Left = Me.TxtVATUSD.Left
                Me.LblVATUSD.Text = "IDR"
                Me.LblVATIdr.Visible = False
                Me.TxtVATUSD.Visible = False
                Me.TxtVatIDR.Text = "1,000,000"
                Me.TxtPurchaseDP.Text = "1,500,000"
                'total
                Me.TxtTotal.Text = "11,006,000"
                'berkaitan dengan data cost
                Me.SplitContainer3.Panel2Collapsed = True
                Me.TxtTotalCost.Text = "0"
                Me.LblDP.Visible = True
                Me.LblDPCurrency.Visible = True
                Me.TxtPurchaseDP.Visible = True
            Else
                'Sales Non-IDR
                Me.TxtTaxRate.Text = "9,025.65"
                Me.TxtPurchaseIDR.Width = Me.TxtTaxRate.Width
                Me.TxtPurchaseIDR.Left = Me.TxtTaxRate.Left
                Me.LblUSD.Text = "USD"
                Me.LblIDR.Visible = True
                Me.TxtPurchaseUSD.Visible = True
                Me.TxtPurchaseIDR.Text = "10,159,723"
                Me.LblDP.Visible = False
                Me.LblDPCurrency.Visible = False
                Me.TxtPurchaseDP.Visible = False
                Me.TxtPurchaseDP.Text = "0"

                'VAT
                Me.TxtVatIDR.Width = Me.TxtTaxRate.Width
                Me.TxtVatIDR.Left = Me.TxtTaxRate.Left
                Me.LblVATUSD.Text = "USD"
                Me.LblVATIdr.Visible = True
                Me.TxtVATUSD.Visible = True
                Me.TxtVatIDR.Text = "1,015,972"
                'total
                Me.TxtTotal.Text = "11,181,695"
                Me.SplitContainer3.Panel2Collapsed = False
                Me.TxtTotalCost.Text = "14,021,625"
                'Me.LblTotalCost1.Visible = True
                'Me.LblTotalCost2.Visible = True
                'Me.TxtTotalCost.Visible = True
                'Me.dgvItem.Columns(Me.ColTotalCost.Index).Visible = True
            End If

            Me.Tampilkan_DGV_Akun()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.RowIndex = Me.dgvItem.RowCount - 1 Then
                If e.ColumnIndex <> Me.ColAccount.Index Then
                    e.Cancel = True
                Else
                    Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            Else
                If e.ColumnIndex <> Me.ColAccount_Desc.Index Then
                    Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex)) = True _
            Or Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColRemarks.Index Then
                    'Remarks
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                ElseIf e.ColumnIndex = Me.ColAccount.Index Then
                    'kalo kode akun dihapus, maka hapus satu bari semuanya
                    If e.RowIndex <> Me.dgvItem.RowCount - 1 Then
                        Me.dgvItem.Rows.RemoveAt(e.RowIndex)
                    End If
                    Me.Tampilkan_DGV_Akun()
                    Me.AdaPerubahan = True
                    Exit Sub
                Else
                    'untuk kolom lain, jika nothing, langsung di balikin ke nilai semula
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                End If
            Else
                ''data account number
                If e.ColumnIndex <> Me.ColRemarks.Index Then
                    If e.ColumnIndex = Me.ColAccount.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            Me.dgvItem.Item(Me.ColAccount_Desc.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                            Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value = Me.ObjAkun.GetPrimary(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                            Me.dgvItem.Item(Me.ColDutyStamp.Index, e.RowIndex).Value = "-"
                            Me.dgvItem.Item(Me.ColPurchase.Index, e.RowIndex).Value = "-"
                            Me.dgvItem.Item(Me.ColVAT.Index, e.RowIndex).Value = "-"
                            Me.dgvItem.Item(Me.ColTotalCost.Index, e.RowIndex).Value = "-"
                            Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value = ""
                            Me.Tampilkan_DGV_Akun()
                            Me.AdaPerubahan = True
                        Else
                            '2. kode akun yang diketik tidak ada --> tampilkan form cari
                            Dim ij As New FrmMasterAkun_Cari
                            ij.txtCari.Text = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.ShowDialog()
                            If ij.mPrimaryAkun <> -1 Then
                                '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAkun.GetAccountCode(ij.mPrimaryAkun)
                                Me.dgvItem.Item(Me.ColAccount_Desc.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value = ij.mPrimaryAkun
                                Me.dgvItem.Item(Me.ColDutyStamp.Index, e.RowIndex).Value = "-"
                                Me.dgvItem.Item(Me.ColPurchase.Index, e.RowIndex).Value = "-"
                                Me.dgvItem.Item(Me.ColVAT.Index, e.RowIndex).Value = "-"
                                Me.dgvItem.Item(Me.ColTotalCost.Index, e.RowIndex).Value = "-"
                                Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value = ""
                                Me.Tampilkan_DGV_Akun()
                                Me.AdaPerubahan = True
                            ElseIf ij.mPrimaryAkun = -1 Then
                                '2.b. kode akun yang dicari tidak ada --> kembalikan ke kondisi buffer
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                    Me.dgvItem.Rows.RemoveAt(e.RowIndex)
                                    Me.Tampilkan_DGV_Akun()
                                    Me.AdaPerubahan = True
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                    Else
                        Me.Tampilkan_DGV_Akun()
                        Me.AdaPerubahan = True
                    End If
                Else
                    Me.AdaPerubahan = True
                End If
                '3 cek apakah ada kode akun yang sama pada referensi kode yang dibuat 
                If Me.dgvItem.Rows.Count > 2 Then
                    Me.dgvItem.AllowUserToAddRows = False
                    Dim item As String = ""
                    For i As Integer = 0 To dgvItem.Rows.Count - 1
                        If e.RowIndex <> i Then
                            item = CStr(Me.dgvItem.Rows(i).Cells(Me.ColAccount.Index).Value)
                            If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = item Then
                                MsgBox("Duplication account number at data item, please change your number", _
                                       MsgBoxStyle.Information, "Information")
                                Me.dgvItem.Rows.RemoveAt(e.RowIndex)
                                Me.dgvItem.AllowUserToAddRows = True
                                Me.Tampilkan_DGV_Akun()
                                Me.AdaPerubahan = True
                                Exit Sub
                            End If
                        End If
                    Next
                    Me.dgvItem.AllowUserToAddRows = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub Tampilkan_DGV_Akun()
        Dim row_dgv As DataGridViewRow
        Dim row_Akun As DataGridViewRow
        Dim Debet, Kredit As Double

        Try
            Me.dgvAkun.Rows.Clear()

            For Each row_dgv In Me.dgvItem.Rows
                Debet = 0
                Kredit = 0
                If row_dgv.Index <> Me.dgvItem.Rows.Count - 1 Then
                    Me.dgvAkun.Rows.Add()
                    row_Akun = Me.dgvAkun.Rows(Me.dgvAkun.Rows.Count - 1)

                    With row_Akun
                        'salin data
                        .Cells(Me.Col_Account_Show.Index).Value = row_dgv.Cells(Me.ColAccount.Index).Value
                        .Cells(Me.ColAccount_Desc.Index).Value = row_dgv.Cells(Me.ColAccount_Desc.Index).Value
                        If row_dgv.Cells(Me.ColPurchase.Index).Value.ToString = "Debet" Then
                            Debet += CDbl(Me.TxtPurchaseIDR.Text)
                        ElseIf row_dgv.Cells(Me.ColPurchase.Index).Value.ToString = "Credit" Then
                            Kredit += CDbl(Me.TxtPurchaseIDR.Text)
                        End If
                        If row_dgv.Cells(Me.ColVAT.Index).Value.ToString = "Debet" Then
                            Debet += CDbl(Me.TxtVatIDR.Text)
                        ElseIf row_dgv.Cells(Me.ColVAT.Index).Value.ToString = "Credit" Then
                            Kredit += CDbl(Me.TxtVatIDR.Text)
                        End If
                        If row_dgv.Cells(Me.ColDutyStamp.Index).Value.ToString = "Debet" Then
                            Debet += CDbl(Me.TxtDutyStamp.Text)
                        ElseIf row_dgv.Cells(Me.ColDutyStamp.Index).Value.ToString = "Credit" Then
                            Kredit += CDbl(Me.TxtDutyStamp.Text)
                        End If
                        If row_dgv.Cells(Me.ColDownpayment.Index).Value.ToString = "Debet" Then
                            Debet += CDbl(Me.TxtPurchaseDP.Text)
                        ElseIf row_dgv.Cells(Me.ColDownpayment.Index).Value.ToString = "Credit" Then
                            Kredit += CDbl(Me.TxtPurchaseDP.Text)
                        End If
                        If Me.cboCurrency.SelectedIndex = 1 Then
                            If row_dgv.Cells(Me.ColTotalCost.Index).Value.ToString = "Debet" Then
                                Debet += CDbl(Me.TxtTotalCost.Text)
                            ElseIf row_dgv.Cells(Me.ColTotalCost.Index).Value.ToString = "Credit" Then
                                Kredit += CDbl(Me.TxtTotalCost.Text)
                            End If
                        End If

                        .Cells(Me.ColDebet.Index).Value = Me.ObjCD.DigitGrouping(Debet)
                        .Cells(Me.ColKredit.Index).Value = Me.ObjCD.DigitGrouping(Kredit)
                    End With
                End If
            Next

            'isikan data cost
            If Me.cboCurrency.SelectedIndex = 1 Then
                For Each row_dgv In Me.dgvCost.Rows
                    Me.dgvAkun.Rows.Add()
                    row_Akun = Me.dgvAkun.Rows(Me.dgvAkun.Rows.Count - 1)
                    With row_Akun
                        'salin data
                        .Cells(Me.Col_Account_Show.Index).Value = row_dgv.Cells(Me.ColCostAccountNo.Index).Value
                        .Cells(Me.ColAccount_Desc.Index).Value = row_dgv.Cells(Me.ColCostAccountDesc.Index).Value

                        .Cells(Me.ColDebet.Index).Value = "0"
                        .Cells(Me.ColKredit.Index).Value = Me.ObjCD.DigitGrouping(row_dgv.Cells(Me.ColCostAmount.Index).Value)
                    End With
                Next
            End If

            Me.Hitung_Unbalance()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Hitung_Unbalance()
        Dim row_dgv As DataGridViewRow
        Dim Debet, Kredit As Double

        Try
            For Each row_dgv In Me.dgvAkun.Rows
                Debet += CDbl(row_dgv.Cells(Me.ColDebet.Index).Value)
                Kredit += CDbl(row_dgv.Cells(Me.ColKredit.Index).Value)
            Next

            Me.TxtUnbalance.Text = Me.ObjCD.DigitGrouping((Debet - Kredit).ToString)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            Me.Tampilkan_DGV_Akun()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim row_dgv As DataGridViewRow
        Dim Purchase_Int, VAT_Int, Dutystamp_Int, TotalCost_Int, DP_Int, Acuan_Int As ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting

        Try
            'data salah jika:
            '2. Jika tidak ada item sama sekali maka berikan peringatan
            If Me.dgvItem.Rows.Count = 0 Then
                If MsgBox("There are no setting for invoice automatic data transfer. Continue?", _
                MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Return False
                End If
            End If
            '3. Pastikan primary akun diisikan
            For Each row_dgv In Me.dgvItem.Rows
                If row_dgv.Index <> Me.dgvItem.RowCount - 1 Then
                    If row_dgv.Cells(Me.ColPrimaryAkun.Index).Value = -1 Then
                        MsgBox("Please specify account number for row " + (row_dgv.Index + 1).ToString, _
                        MsgBoxStyle.Information, "Information")
                        Me.dgvItem.ClearSelection()
                        row_dgv.Selected = True
                        Me.dgvItem.Focus()
                        Return False
                    End If
                End If
            Next

            For Each row_dgv In Me.dgvItem.Rows
                If row_dgv.Index <> Me.dgvItem.RowCount - 1 Then
                    If row_dgv.Cells(Me.ColPurchase.Index).Value.ToString = "-" Then
                        Purchase_Int = ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.None
                    ElseIf row_dgv.Cells(Me.ColPurchase.Index).Value.ToString = "Credit" Then
                        Purchase_Int = ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.Kredit
                    ElseIf row_dgv.Cells(Me.ColPurchase.Index).Value.ToString = "Debet" Then
                        Purchase_Int = ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.Debet
                    End If
                    If row_dgv.Cells(Me.ColVAT.Index).Value.ToString = "-" Then
                        VAT_Int = ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.None
                    ElseIf row_dgv.Cells(Me.ColVAT.Index).Value.ToString = "Credit" Then
                        VAT_Int = ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.Kredit
                    ElseIf row_dgv.Cells(Me.ColVAT.Index).Value.ToString = "Debet" Then
                        VAT_Int = ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.Debet
                    End If
                    If row_dgv.Cells(Me.ColDutyStamp.Index).Value.ToString = "-" Then
                        Dutystamp_Int = ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.None
                    ElseIf row_dgv.Cells(Me.ColDutyStamp.Index).Value.ToString = "Credit" Then
                        Dutystamp_Int = ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.Kredit
                    ElseIf row_dgv.Cells(Me.ColDutyStamp.Index).Value.ToString = "Debet" Then
                        Dutystamp_Int = ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.Debet
                    End If
                    If row_dgv.Cells(Me.ColTotalCost.Index).Value.ToString = "-" Then
                        TotalCost_Int = ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.None
                    ElseIf row_dgv.Cells(Me.ColTotalCost.Index).Value.ToString = "Credit" Then
                        TotalCost_Int = ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.Kredit
                    ElseIf row_dgv.Cells(Me.ColTotalCost.Index).Value.ToString = "Debet" Then
                        TotalCost_Int = ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.Debet
                    End If
                    If row_dgv.Cells(Me.ColDownpayment.Index).Value.ToString = "-" Then
                        DP_Int = ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.None
                    ElseIf row_dgv.Cells(Me.ColDownpayment.Index).Value.ToString = "Credit" Then
                        DP_Int = ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.Kredit
                    ElseIf row_dgv.Cells(Me.ColDownpayment.Index).Value.ToString = "Debet" Then
                        DP_Int = ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.Debet
                    End If

                    '4. dalam 1 bari tidak boleh semuannya none
                    If Purchase_Int = ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.None And _
                    VAT_Int = ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.None And _
                    Dutystamp_Int = ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.None And _
                    TotalCost_Int = ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.None And _
                    DP_Int = ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.None Then
                        MsgBox("Please specify value on row " + (row_dgv.Index + 1).ToString, _
                        MsgBoxStyle.Information, "Information")
                        Me.dgvItem.ClearSelection()
                        row_dgv.Selected = True
                        Me.dgvItem.Focus()
                        Return False
                    Else
                        '5. dalam 1 bari tidak boleh ada debet dan credit 
                        If Purchase_Int <> ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.None Then
                            Acuan_Int = Purchase_Int
                        ElseIf VAT_Int <> ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.None Then
                            Acuan_Int = VAT_Int
                        ElseIf Dutystamp_Int <> ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.None Then
                            Acuan_Int = Dutystamp_Int
                        ElseIf TotalCost_Int <> ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.None Then
                            Acuan_Int = TotalCost_Int
                        Else
                            Acuan_Int = DP_Int
                        End If
                        'cek satu2, kecuali yang none
                        If Purchase_Int <> ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.None Then
                            If Acuan_Int <> Purchase_Int Then
                                MsgBox("The value must consistent debet or credit on one row.", _
                                MsgBoxStyle.Information, "Information")
                                Me.dgvItem.ClearSelection()
                                row_dgv.Selected = True
                                Me.dgvItem.Focus()
                                Return False
                            End If
                        End If
                        If VAT_Int <> ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.None Then
                            If Acuan_Int <> VAT_Int Then
                                MsgBox("The value must consistent debet or credit on one row.", _
                                MsgBoxStyle.Information, "Information")
                                Me.dgvItem.ClearSelection()
                                row_dgv.Selected = True
                                Me.dgvItem.Focus()
                                Return False
                            End If
                        End If
                        If Dutystamp_Int <> ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.None Then
                            If Acuan_Int <> Dutystamp_Int Then
                                MsgBox("The value must consistent debet or credit on one row.", _
                                MsgBoxStyle.Information, "Information")
                                Me.dgvItem.ClearSelection()
                                row_dgv.Selected = True
                                Me.dgvItem.Focus()
                                Return False
                            End If
                        End If
                        If TotalCost_Int <> ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.None Then
                            If Acuan_Int <> TotalCost_Int Then
                                MsgBox("The value must consistent debet or credit on one row.", _
                                MsgBoxStyle.Information, "Information")
                                Me.dgvItem.ClearSelection()
                                row_dgv.Selected = True
                                Me.dgvItem.Focus()
                                Return False
                            End If
                        End If
                        If DP_Int <> ClsTransaksiSettingTransf_Inv.Enum_Tipe_Setting.None Then
                            If Acuan_Int <> DP_Int Then
                                MsgBox("The value must consistent debet or credit on one row.", _
                                MsgBoxStyle.Information, "Information")
                                Me.dgvItem.ClearSelection()
                                row_dgv.Selected = True
                                Me.dgvItem.Focus()
                                Return False
                            End If
                        End If
                    End If
                End If
            Next

            '1. Unbalance <> 0
            If CDbl(Me.TxtUnbalance.Text) <> 0 Then
                MsgBox("Please make sure balance amount on credit and debet.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If Me.ApakahDataBenar = True Then
                'sampai sini konfirmasi simpan
                If MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    If Me.ObjSetting.SimpanData(Me, Me.dgvItem, Me.UserCode, Me.tslCreatedBy, Me.tslCreatedDate) Then
                        MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                        Me.AdaPerubahan = False
                        Me.Close()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If Me.AdaPerubahan = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel)

                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar = True Then
                        If Me.ObjSetting.SimpanData(Me, Me.dgvItem, Me.UserCode, Me.tslCreatedBy, Me.tslCreatedDate) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.AdaPerubahan = False
                            Me.Close()
                        End If
                    End If
                ElseIf hasil = MsgBoxResult.No Then
                    Me.AdaPerubahan = False
                    Me.Close()
                End If
            Else
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSettingInvoice_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Try
            If Me.AdaPerubahan = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel)
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar = True Then
                        If Me.ObjSetting.SimpanData(Me, Me.dgvItem, Me.UserCode, Me.tslCreatedBy, Me.tslCreatedDate) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
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
            Me.Tampilkan_DGV_Akun()
            Me.AdaPerubahan = True
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
            Me.Tampilkan_DGV_Akun()
            Me.AdaPerubahan = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProcess_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.SelectionChanged
        Try
            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            Else
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
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

    Private Sub dgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Dim row_dgv As DataGridViewRow

        Try
            If Me.dgvItem.AllowUserToAddRows = True Then
                row_dgv = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                With row_dgv
                    .Cells(Me.ColPrimaryAkun.Index).Value = "-1"
                    .Cells(Me.ColRemarks.Index).Value = ""
                    .Cells(Me.ColPurchase.Index).Value = "-"
                    .Cells(Me.ColVAT.Index).Value = "-"
                    .Cells(Me.ColDutyStamp.Index).Value = "-"
                    .Cells(Me.ColAccount_Desc.Index).Value = ""
                    .Cells(Me.ColTotalCost.Index).Value = "-"
                    .Cells(Me.ColAccount.Index).Value = ""
                    .Cells(Me.ColDownpayment.Index).Value = "-"
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick

    End Sub
End Class