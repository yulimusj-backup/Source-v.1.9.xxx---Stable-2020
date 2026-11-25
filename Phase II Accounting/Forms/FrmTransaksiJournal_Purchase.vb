Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiJournal_Purchase
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightSeaGreen
    Public TipeJournal As ClsTransaksiJournal.EnumTipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal
    Public ObjJurnal As New ClsTransaksiJournal

    Dim ObjCD As New Core_Module.ClsCekData
    Dim Obj_AOPBalance As New ClsTransaksiPurchaseBalance
    Dim ObjCDNote As New ClsTransaksiCreditNote
    Dim TBL_AOP_Balance As New DseTransaksiPurchaseBalance.balance_purchase_viewDataTable
    Dim Obj_AOPR_Balance As New ClsTransaksiPurchaseReturnBalance
    Dim TBL_AOPR_Balance As New DseTransaskiPurchaseReturnBalance.balance_purchase_return_viewDataTable
    Dim ObjSetting As New ClsTransaksiSetting

    Dim Buffer As String = ""
    Dim DefaultPrecision As Integer = 0
    Dim Baru As Boolean
    Dim ObjAkun As New ClsMasterAkun
    Dim ObjSupp As New clsMasterSupplier
    Dim ObjQuickJurnal As New ClsTransaksiJournalQuickAdd
    Dim ObjLicense As New ClsLicense
    Dim ObjClosing As New ClsTransaksiClosing
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjAOP As New ClsTransaksiAOP
    Dim ObjPR As New ClsTransaksiPurchaseReturn
    Dim ObjBalanceDP As New ClsTransaksiBalanceManual_DP 'add aul
    Dim PrimaryBaseCurrency As Double = 1
    Dim ObjTaxRate As New ClsTransaksiSalesInvoiceTaxRate
    Dim TBLAkun As New DseMasterAkun.tblmaster_akun4DataTable

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

    Private Sub llbEditQuickJournal_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbEditQuickJournal.LinkClicked
        Dim ij As New FrmTransaksiJournal_quickAdd

        Try
            ij.UserCOde = Me.UserCode
            ij.BackcolorAll = Me.BackColorAll
            ij.TipeJournal = Me.TipeJournal
            ij.ShowDialog()
            Me.ObjQuickJurnal.IsiCboQuickAdd(Me.TipeJournal, Me.cboQuickJournal)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij = Nothing
        End Try
    End Sub

    Private Sub FrmTransaksiJournal_Purchase_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim NullValue As String = ""

        Try
            Me.AutoSizeMode = Windows.Forms.AutoSizeMode.GrowAndShrink
            Me.Size = New System.Drawing.Size(1350, 678)
            Me.MinimumSize = New System.Drawing.Size(1350, 678)

            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvPurchaseBalance.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvPurchaseReturnBalance.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvDifference.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                Me.SplitContainer3.Panel2Collapsed = True
                Me.Text = "Purchase Payment Journal"
            Else
                Me.SplitContainer3.Panel1Collapsed = True
                Me.Text = "Purchase Return Receivement Journal"
            End If
            'Me.ColAccount.Visible = True
            'Me.ColDescription.Visible = True


            ''Me.AktifkanDetail()
            Me.NonAktifkanDetail()

            Me.dgvItem.AllowUserToAddRows = False
            Me.DgvDifference.AllowUserToAddRows = False
            Me.ObjJurnal = New ClsTransaksiJournal
            Me.ObjJurnal.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal
            Me.Baru = True
            Me.dgvItem.Rows.Clear()
            Me.DgvDifference.Rows.Clear()
            Me.KosongkanField()

            'set null value untuk dgv
            Me.DefaultPrecision = Me.ObjCur.GetDecimal_DefaultCurrency
            NullValue = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            Me.ColDebet.DefaultCellStyle.NullValue = NullValue
            Me.ColKredit.DefaultCellStyle.NullValue = NullValue
            Me.TxtUnbalance.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            Me.TxtPurchasePayment.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            Me.TxtPurchaseReturnReceivement.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)

            Me.ColAOPDate.Visible = True
            Me.ColAOPDate.HeaderText = "Date"
            Me.ShowIcon = False
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
        Dim DefPrec As Integer = 0

        Try
            Me.dgvItem.TabStop = True
            If Me.TxtRemarks.Enabled = True Then
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True

                'set nilai default baris baru
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value = Me.PrimaryBaseCurrency
                DefPrec = Me.ObjCur.GetDefaultPrecission(rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value)
                rowDGV.Cells(Me.ColDebetOriginal.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
                rowDGV.Cells(Me.ColCreditOriginal.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
                rowDGV.Cells(Me.ColDebet.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                rowDGV.Cells(Me.ColKredit.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                rowDGV.Cells(Me.ColExchangeRate.Index).Value = Me.ObjCD.DigitGrouping(1, DefPrecExchangeRate)
                rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value = Me.PrimaryBaseCurrency
                rowDGV.Cells(Me.ColCurrency.Index).Value = Me.ObjCur.GetCurrencyCode(Me.PrimaryBaseCurrency)
                rowDGV.Cells(Me.ColPrimaryAkun.Index).Value = -1
                rowDGV.Cells(Me.ColAccount.Index).Value = ""
                rowDGV.Cells(Me.ColRemarks.Index).Value = ""
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
            Me.dgvItem.AllowUserToAddRows = False
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
            Me.dgvItem.AllowUserToAddRows = False
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
        Dim ij As New FrmTransaksiJournal_CariPurchase
        ij.TipeJournal = Me.TipeJournal

        Try
            ij.BackColorAll = Me.BackColorAll
            ij.ShowDialog()
            Me.dgvItem.AllowUserToAddRows = False
            If Not ij.PrimaryPP = -1 Then
                If Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, ClsTransaksiSetting.EnumNamaParameter.Jurnal_IsLedgerAccessActive) = "1" Then
                    If Me.ObjJurnal.ApakahMemilikiLedgerAccess(Me.UserCode, Me.ObjJurnal.GetCreatedBy(ij.PrimaryPP)) = False Then
                        MsgBox("You are not authorized to view this data.", MsgBoxStyle.Information, "Information")
                        Me.KosongkanField()
                        Exit Sub
                    End If
                End If

                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                If Me.ObjJurnal.CariDetailPurchaseJournal(ij.PrimaryPP, Me.dgvItem, Me, ij.TipeJournal, Me.DgvDifference) = True Then
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

    Private Function HitungUnbalance(ByVal vdebet As Decimal, ByVal vkredit As Decimal) As Decimal
        Dim Unbalance As Decimal = 0
        Try
            Unbalance = ObjCD.DigitGrouping(vdebet - vkredit, Me.DefaultPrecision)

            Return Unbalance
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub AktifkanDetail()
        Dim row As DseMasterAkun.tblmaster_akun4Row

        Try
            'isikan data balance- dipindahkan setelah user mengisikan data customer code
            'If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
            'Me.ObjInvoiceBalance.IsiDGVBalance(Me.dgvPurchaseBalance, Me.TBLSalesBalance, Me, Me.ObjCD)
            'ElseIf Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.SalesReturnPaymentJournal Then
            'Me.ObjSRBalance.IsiDGVBalance(Me.dgvPurchaseReturnBalance, Me.TBLSRBalance, Me, Me.ObjCD)
            'End If

            ''standar
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryJournal.BackColor, False)
            Me.ObjCD.SetBackColor(Me.GroupBox1.Controls, Me.TxtPrimaryJournal.BackColor, False)
            Me.ObjCD.SetBackColor(Me.GroupBox2.Controls, Me.TxtPrimaryJournal.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimaryJournal.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimaryJournal.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.TxtPrimaryJournal.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.TxtPrimaryJournal.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Me.TxtPrimaryJournal.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Me.TxtPrimaryJournal.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel1.Controls, Me.TxtPrimaryJournal.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel2.Controls, Me.TxtPrimaryJournal.BackColor, False)

            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.GroupBox1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.GroupBox2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer4.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer4.Panel2.Controls, True, True, True, True, True, True, True, True)

            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, True, True)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None

            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = False
            Me.dgvItem.AllowUserToAddRows = False

            Me.DgvDifference.AllowUserToDeleteRows = True
            Me.DgvDifference.ReadOnly = False

            Me.dgvPurchaseBalance.AllowUserToDeleteRows = False
            Me.dgvPurchaseBalance.ReadOnly = False
            Me.dgvPurchaseReturnBalance.AllowUserToDeleteRows = False
            Me.dgvPurchaseReturnBalance.ReadOnly = False

            If Me.dgvItem.Rows.Count = 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If

            Me.cboQuickJournal.SelectedIndex = -1
            Me.llbEditQuickJournal.Enabled = True
            Me.ObjQuickJurnal.IsiCboQuickAdd(Me.TipeJournal, Me.cboQuickJournal)
            ''Me.txtNoTransaksi.ReadOnly = True
            ''Me.txtNoTransaksi.BackColor = Me.txtReadOnly.BackColor

            If Me.cboPayToAcc.SelectedIndex <> -1 Then
                row = Me.TBLAkun.DefaultView(Me.cboPayToAcc.SelectedIndex).Row

                If row.primary_currency = Core_Module.PrimaryBaseCurrency Then
                    Me.TxtExchangeRate.ReadOnly = True
                    Me.TxtExchangeRate.TabStop = False
                    Me.TxtExchangeRate.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtExchangeRate.ReadOnly = False
                    Me.TxtExchangeRate.TabStop = True
                    Me.TxtExchangeRate.BackColor = Me.TxtPrimaryAccountPayTo.BackColor
                End If
            Else
                Me.TxtExchangeRate.ReadOnly = True
                Me.TxtExchangeRate.TabStop = False
                Me.TxtExchangeRate.BackColor = Me.txtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ''standar
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.GroupBox1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.GroupBox2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer4.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer4.Panel2.Controls, True, False, True, False, True, False, True, False)

            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, True, False, True, False)
            Me.CancelButton = Me.btnExit

            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True
            Me.dgvItem.AllowUserToAddRows = False
            Me.DgvDifference.AllowUserToDeleteRows = False
            Me.DgvDifference.ReadOnly = True

            Me.dgvPurchaseBalance.AllowUserToDeleteRows = False
            Me.dgvPurchaseBalance.ReadOnly = True
            Me.dgvPurchaseReturnBalance.AllowUserToDeleteRows = False
            Me.dgvPurchaseReturnBalance.ReadOnly = True

            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.GroupBox1.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.GroupBox2.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel1.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel2.Controls, Me.BackColor, True)

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
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer3.Panel1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer3.Panel2.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer4.Panel1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer4.Panel2.Controls, True, True, True)

            Me.TxtPrimaryJournal.Text = "-1"
            Me.TxtPrimarySupp.Text = "-1"
            Me.TxtPrimaryAccountDifference.Text = "-1"
            Me.TxtPrimaryAccountPayTo.Text = "-1"
            Me.txtPrimaryExchangeRateGainLoss.Text = "-1"
            Me.TxtUnbalance.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            Me.TxtPurchasePayment.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            Me.TxtPurchaseReturnReceivement.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping("1", Core_Module.DefPrecExchangeRate)
            Me.TxtPaid.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision)

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "

            Me.dgvItem.Rows.Clear()
            Me.dgvPurchaseBalance.Rows.Clear()
            Me.dgvPurchaseReturnBalance.Rows.Clear()
            Me.DgvDifference.Rows.Clear()

            Me.cboPayToAcc.DataSource = Nothing
            Me.cboPayToAcc.Items.Clear()
            ''Me.ObjCD.SetNumber(Me.dgvItem)
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
            i.UserCode = Me.UserCode
            ' Set the caption to be unique.
            i.BackCOlorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = FrmParent
            i.Show()
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
                .PrimaryPihakAcuan = Me.TxtPrimarySupp.Text
                .PrimaryExChangeRateGainLoss = Me.txtPrimaryExchangeRateGainLoss.Text
                .PaymentMethod = Me.txtPaymentMethod.Text
                .PrimaryAkunDifference = Me.TxtPrimaryAccountDifference.Text
                .PrimaryAkunPayTo = Me.TxtPrimaryAccountPayTo.Text
                .ExchangeRate = Me.TxtExchangeRate.Text
                .Paid = Me.TxtPaid.Text
                .ARPaid_IDR = Me.TxtARPaid_IDR.Text
                .Difference = Me.TxtDifference.Text
                .Paid_IDR = Me.TxtPaid_InIDR.Text
                .ExchageRateGainLoss = Me.TxtSelisihExcRate.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub ShowJurnal(ByVal vPrimaryJurnal As Double)
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.btnCancel.PerformClick()
            End If
            If Me.TxtRemarks.Enabled = False Then
                If Me.ObjJurnal.CariDetailPurchaseJournal(vPrimaryJurnal, Me.dgvItem, Me, Me.TipeJournal, Me.DgvDifference) = True Then
                    Me.SalinkeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinkeField()
        Dim row As DseMasterAkun.tblmaster_akun4Row
        Dim m As Integer = 0
        Dim DefPrec As Integer = 0
        Dim DefPrecBase As Integer = 0
        Dim ApakahSudahAdaAkun As Boolean = False

        Try
            With ObjJurnal
                Me.TxtPrimaryJournal.Text = .Primary
                Me.TxtRefNo.Text = .RefNumber
                Me.TxtRemarks.Text = .Description
                Me.dtpDate.Value = .Tanggal
                Me.TxtPrimarySupp.Text = .PrimaryPihakAcuan
                Me.TxtSuppCode.Text = ObjSupp.GetSupplierCode(Me.TxtPrimarySupp.Text)
                Me.TxtSuppName.Text = ObjSupp.GetSupplierName(Me.TxtPrimarySupp.Text)
                Me.txtPaymentMethod.Text = .PaymentMethod
                Me.txtPrimaryExchangeRateGainLoss.Text = .PrimaryExChangeRateGainLoss
                Me.cboQuickJournal.SelectedIndex = -1
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString

                'salin data payment
                If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                    Me.Obj_AOPBalance.IsiDGVBalance(Me.dgvPurchaseBalance, _
                    Me.TBL_AOP_Balance, Me, Me.ObjCD, Me.TxtPrimarySupp.Text, Me.TxtPrimaryJournal.Text, Me.DefaultPrecision)
                    Me.TxtPurchasePayment.Text = Me.ObjCD.DigitGrouping(Me.GetTotalPaymentPurchase)
                ElseIf Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchaseReturnReceivement Then
                    Me.Obj_AOPR_Balance.IsiDGVBalance(Me.dgvPurchaseReturnBalance, _
                    Me.TBL_AOPR_Balance, Me, Me.ObjCD, Me.TxtPrimarySupp.Text, Me.TxtPrimaryJournal.Text, Me.DefaultPrecision)
                    Me.TxtPurchaseReturnReceivement.Text = Me.ObjCD.DigitGrouping(Me.GetTotalReceivementReturn)
                End If

                Me.TxtPrimaryAccountDifference.Text = Me.ObjJurnal.PrimaryAkunDifference
                Me.TxtDifferenceAccountCode.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryAccountDifference.Text)
                Me.cboPayToAcc.DataSource = Nothing
                Me.cboPayToAcc.Items.Clear()
                Me.SetCBOPayment()
                'jika ada account yang dipilih, maka cek apakah sudah ada di dalam tabel, 
                'jika belum ada maka tambahkan datanya ke dalam tabel
                For Each dv As DataRowView In Me.TBLAkun.DefaultView
                    row = dv.Row
                    If row.Primary = Me.ObjJurnal.PrimaryAkunPayTo Then
                        ApakahSudahAdaAkun = True
                        Exit For
                    End If
                Next
                If ApakahSudahAdaAkun = False Then
                    Me.cboPayToAcc.SelectedIndex = -1

                    'tambahkan baris baru pada tabel
                    row = Me.TBLAkun.Newtblmaster_akun4Row
                    row.Primary = Me.ObjJurnal.PrimaryAkunPayTo
                    row.primary_currency = Me.ObjAkun.GetPrimaryCurrency(Me.ObjJurnal.PrimaryAkunPayTo)
                    row.Description = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.ObjJurnal.PrimaryAkunPayTo)
                    row.Account_Code = Me.ObjAkun.GetAccountCode(Me.ObjJurnal.PrimaryAkunPayTo)
                    Me.TBLAkun.Rows.Add(row)
                End If

                m = 0
                For Each dv As DataRowView In Me.TBLAkun.DefaultView
                    row = dv.Row
                    If row.Primary = Me.ObjJurnal.PrimaryAkunPayTo Then
                        Me.cboPayToAcc.SelectedIndex = m
                        Exit For
                    End If
                    m += 1
                Next
                Me.TxtPrimaryAccountPayTo.Text = Me.ObjJurnal.PrimaryAkunPayTo

                Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping(.ExchangeRate, Core_Module.DefPrecExchangeRate)

                DefPrec = Me.ObjCur.GetDefaultPrecission(Me.ObjAkun.GetPrimaryCurrency(Me.TxtPrimaryAccountPayTo.Text))
                DefPrecBase = Me.ObjCur.GetDecimal_DefaultCurrency

                Me.TxtPaid.Text = Me.ObjCD.DigitGrouping(.Paid, DefPrec)
                Me.TxtARPaid_IDR.Text = Me.ObjCD.DigitGrouping(.ARPaid_IDR, DefPrecBase)
                Me.TxtDifference.Text = Me.ObjCD.DigitGrouping(.Difference, DefPrecBase)
                Me.TxtPaid_InIDR.Text = Me.ObjCD.DigitGrouping(.Paid_IDR, DefPrecBase)
                Me.TxtSelisihExcRate.Text = Me.ObjCD.DigitGrouping(.ExchageRateGainLoss, DefPrecBase)
            End With

            Me.SetGBOEquivalent()
            Me.SetMulticurrency()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahAdaPayment() As Boolean
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvPurchaseBalance.Rows
                If CDec(rowDGV.Cells(Me.ColPaymentOriginal.Index).Value) <> 0 Then
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahAdaCurrency(ByVal vPrimaryCurr As Double) As Boolean
        Dim row As DseMasterAkun.tblmaster_akun4Row

        Try
            'klo tabelnya tidak terhubung dengan cbo, maka return false
            If IsNothing(Me.cboPayToAcc.DataSource) = True Then
                Return False
            Else
                If Me.cboPayToAcc.DataSource.Equals(Me.TBLAkun.DefaultView) = False Then
                    Return False
                End If
            End If

            For Each row In Me.TBLAkun.Rows
                If row.primary_currency = vPrimaryCurr Then
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function


    Private Function ApakahAdadiferentCurrency(ByVal vPrimaryCurr As Double) As Boolean
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvPurchaseBalance.Rows
                If CDec(row.Cells(Me.ColPaymentOriginal.Index).Value) <> 0 Then
                    If row.Cells(Me.ColPrimaryCurrencyBalance.Index).Value <> vPrimaryCurr Then
                        Return True
                    End If
                End If

            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Function GetPrimaryCurrencyDokumenAcuan(ByVal vTipeDokumenAcuan As ClsTransaksiSalesBalance.EnumTipeSalesBalance, _
  ByVal vPrimaryDokumenAcuan As Double) As Double
        Try
            If vTipeDokumenAcuan = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic Then
                Return Me.ObjAOP.GetPrimaryCurrency(vPrimaryDokumenAcuan)
            ElseIf vTipeDokumenAcuan = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR Then
                Return Me.ObjPR.GetPrimaryCurrency(vPrimaryDokumenAcuan)
            ElseIf vTipeDokumenAcuan = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual Then
                Return Me.Obj_AOPBalance.GetPrimaryCurrency(vPrimaryDokumenAcuan)
            ElseIf vTipeDokumenAcuan = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR Then
                Return Me.Obj_AOPR_Balance.GetPrimaryCurrency(vPrimaryDokumenAcuan)
            ElseIf vTipeDokumenAcuan = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note Then
                Return Me.ObjCDNote.GetPrimaryCurrencyAPAR(vPrimaryDokumenAcuan)
            ElseIf vTipeDokumenAcuan = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note Then
                Return Me.ObjCDNote.GetPrimaryCurrencyAPAR(vPrimaryDokumenAcuan)
            ElseIf vTipeDokumenAcuan = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_DP Then 'add aul
                Return Me.ObjBalanceDP.GetPrimaryCurrency(vPrimaryDokumenAcuan)
            Else
                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub SetCBOPayment()
        Dim DV As DataRowView
        Dim row As DseMasterAkun.tblmaster_akun4Row
        Dim m As Integer = 0
        Dim PrimaryCurPayment As Double = -1
        Dim IsikanData As Boolean = False
        Dim rowDGV As DataGridViewRow
        Dim TipeTransaksiAsal As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
        Dim PrimaryCurrency As Double = -1

        Try
            If Me.ApakahAdaPayment = True Then
                'jika ada payment, dan belum ada isi cbo, maka cbo diisikan
                If Me.cboPayToAcc.Items.Count = 0 Then
                    IsikanData = True
                Else
                    'jika ada primary currency payment yg non idr, tetapi tidak ada primary currency tersebut dalam
                    'tbl akun
                    For Each rowDGV In Me.dgvPurchaseBalance.Rows
                        If CDec(rowDGV.Cells(Me.ColPaymentOriginal.Index).Value) <> 0 Then
                            If rowDGV.Cells(Me.ColPrimaryCurrencyBalance.Index).Value <> Core_Module.PrimaryBaseCurrency Then
                                'If Me.ApakahAdaCurrency(rowDGV.Cells(Me.ColPrimaryCurrencyBalance.Index).Value) = False Then
                                IsikanData = True
                                Exit For
                                'End If
                            ElseIf rowDGV.Cells(Me.ColPurchaseValue.Index).Value = 0 Then
                                TipeTransaksiAsal = rowDGV.Cells(Me.ColTipeDocumentAcuan_Purchase.Index).Value
                                PrimaryCurrency = Me.GetPrimaryCurrencyDokumenAcuan(TipeTransaksiAsal, _
                                rowDGV.Cells(Me.ColPrimaryDocumentAcuan_Purchase.Index).Value)
                                'khusus untuk item balance sales dari usd, maka yg dihitung adalah currency invoicenya
                                If Me.ApakahAdaCurrency(PrimaryCurrency) = False Then
                                    IsikanData = True
                                    Exit For
                                End If
                            End If
                        End If
                    Next

                    'jika semua data sales balance idr, sedangkan ada data primary akun pembayaran yg non idr, maka isikan ulang data
                    If IsikanData = False Then
                        If Me.ApakahAdaCurrencyNonIDR = True Then
                            'cek apakah ada 1 yg non idr
                            For Each rowDGV In Me.dgvPurchaseBalance.Rows
                                If rowDGV.Cells(Me.ColPrimaryCurrencyBalance.Index).Value <> 1 Then
                                    IsikanData = True
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If

                If IsikanData = True Then
                    PrimaryCurPayment = Me.GetPrimaryCurrencyPayment
                    Me.ObjAkun.IsiCboPayment(PrimaryCurPayment, Me.cboPayToAcc, Me.TBLAkun)

                    m = 0
                    For Each DV In Me.TBLAkun.DefaultView
                        row = DV.Row
                        If row.primary_currency = PrimaryCurPayment Then
                            Me.cboPayToAcc.SelectedIndex = m
                            Exit Sub
                        End If
                        m += 1
                    Next

                    Me.cboPayToAcc.SelectedIndex = -1

                    If Me.cboPayToAcc.SelectedIndex = -1 Then
                        If Me.cboPayToAcc.Items.Count > 0 Then
                            Me.cboPayToAcc.SelectedIndex = 0
                        End If
                    End If
                End If
            Else
                'jika tidak ada payment, dan sudah ada isi cbo, maka buang cbo
                If Me.cboPayToAcc.Items.Count > 0 Then
                    Me.cboPayToAcc.DataSource = Nothing
                    Me.cboPayToAcc.Items.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'mengembalikan nilai true jika ada currency non idr pada akun payment
    Private Function ApakahAdaCurrencyNonIDR() As Boolean
        Dim row As DseMasterAkun.tblmaster_akun4Row

        Try
            For Each row In Me.TBLAkun.Rows
                If row.primary_currency <> Core_Module.PrimaryBaseCurrency Then
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Function GetPrimaryCurrencyPayment() As Double
        Dim rowDGV As DataGridViewRow
        Dim PrimaryCurDokumenAcuan As Double = -1

        Try
            For Each rowDGV In Me.dgvPurchaseBalance.Rows
                If CDec(rowDGV.Cells(Me.ColPaymentOriginal.Index).Value) <> 0 Then
                    If CDec(rowDGV.Cells(Me.ColPurchaseValue.Index).Value) <> 0 Then
                        'ambil primary currency dari transaksi asal
                        PrimaryCurDokumenAcuan = Me.GetPrimaryCurrencyDokumenAcuan(rowDGV.Cells(Me.ColTipeDocumentAcuan_Purchase.Index).Value, _
                        rowDGV.Cells(Me.ColPrimaryDocumentAcuan_Purchase.Index).Value)
                        If PrimaryCurDokumenAcuan <> Core_Module.PrimaryBaseCurrency Then
                            Return PrimaryCurDokumenAcuan
                        End If
                    Else
                        'untuk tax dan duty stamp, yg dihitung adalah currency dokumen acuannya
                        If CDec(rowDGV.Cells(Me.ColPrimaryCurrencyBalance.Index).Value) <> Core_Module.PrimaryBaseCurrency Then
                            Return rowDGV.Cells(Me.ColPrimaryCurrencyBalance.Index).Value
                        End If
                    End If
                End If
            Next

            Return Core_Module.PrimaryBaseCurrency
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SetGBOEquivalent()
        Dim row As DseMasterAkun.tblmaster_akun4Row
        Dim rowDGV As DataGridViewRow
        Dim Beda As Boolean = False

        Try
            If Me.cboPayToAcc.SelectedIndex = -1 Then
                Me.gboEquivalent.Visible = False
            Else
                row = Me.TBLAkun.DefaultView(Me.cboPayToAcc.SelectedIndex).Row
                For Each rowDGV In Me.dgvPurchaseBalance.Rows
                    If row.primary_currency <> rowDGV.Cells(Me.ColPrimaryCurrencyBalance.Index).Value Then
                        Beda = True
                        Exit For
                    End If
                Next

                If Beda = False Then
                    Me.gboEquivalent.Visible = False
                Else
                    Me.gboEquivalent.Visible = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim PrimaryAkunHutang As Double = -1
        Dim NilaiPayment As Decimal = 0
        Dim NilaiPayment_AkunHutang As Decimal = 0
        Dim TotalReceivement As Decimal = 0
        Dim rowDGV As DataGridViewRow
        Dim TipeAkun As ClsMasterAkun.EnumTipeAkun
        Dim MaxDate As Date = Date.MinValue
        Dim DummyDoc As String = ""

        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("Reference number")
            KoleksiTextBox.Add(Me.TxtRefNo)
            KoleksiNamaData.Add("Description")
            KoleksiTextBox.Add(Me.TxtRemarks)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
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

            'jika diperlukan payment, maka harus diisikan nilai paymentnya
            'diperlukan bila, total receivement original <> 0, bila = 0, maka ada kemungkinkan sales return di bayarkan dengan invoice dan pekpok
            TotalReceivement = 0
            For Each rowDGV In Me.dgvPurchaseBalance.Rows
                TotalReceivement += CDec(rowDGV.Cells(Me.ColPaymentOriginal.Index).Value)
            Next
            If TotalReceivement <> 0 And CDec(Me.TxtPaid.Text) = 0 Then
                If MsgBox("No payment amount specified, continue save?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Me.TxtPaid.Focus()
                    Return False
                End If
            End If
            If Me.ApakahAdaPayment = False Then
                MsgBox("Please select at least 1 item to pay.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            ' Cek debet kredit harus balance
            Me.HitungUnbalance()
            If CDec(Me.TxtUnbalance.Text) <> 0 Then
                'Me.HitungUnbalance()
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
            Dim row As DataGridViewRow
            For Each row In Me.dgvItem.Rows
                If Me.dgvItem.Item(Me.ColDebet.Index, row.Index).Value = 0 And _
                Me.dgvItem.Item(Me.ColKredit.Index, row.Index).Value = 0 Then
                    MsgBox("Please add positive number in account no. " + Me.dgvItem.Item(Me.ColAccount.Index, row.Index).Value, _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Next

            'cek data yang salah, double dalam kode
            If Me.Baru = True Then
                If Me.ObjJurnal.ApakahSudahAda(Me.TxtRefNo.Text) = True Then
                    MsgBox("Duplication in referance number =" + Me.TxtRefNo.Text + ", please change your number", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtRefNo.Focus()
                    Return False
                End If
            Else
                Dim TipeCodeLama As String = Me.ObjJurnal.GetRefCode(Me.TxtPrimaryJournal.Text)
                If TipeCodeLama <> Me.TxtRefNo.Text Then
                    If Me.ObjJurnal.ApakahSudahAda(Me.TxtRefNo.Text) = True Then
                        MsgBox("Duplication in referance number =" + Me.TxtRefNo.Text + ", please change your number", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtRefNo.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.ObjJurnal.GetTimestamp(Me.TxtPrimaryJournal.Text)
                If Me.ObjJurnal.Timestamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjJurnal.CariDetailPurchaseJournal(Me.ObjJurnal.Primary, Me.dgvItem, Me, Me.TipeJournal, Me.DgvDifference)
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
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader2(row.Cells(Me.ColPrimaryAkun.Index).Value) = False Then
                        MsgBox("Invalid account in row " + row.Index.ToString, MsgBoxStyle.Information, "Information")
                        Me.dgvItem.ClearSelection()
                        row.Selected = True
                        Return False
                    End If
                End If
            Next

            'kalo total receveiment minus, maka payment harus minus, begitu juga sebaliknya
            If CDec(Me.TxtPaid.Text) > 0 And Me.GetTotalPaymentPurchase < 0 Then
                MsgBox("Please specify a negatif value payment.", MsgBoxStyle.Information, "Information")
                Me.TxtPaid.Focus()
                Return False
            ElseIf CDec(Me.TxtPaid.Text) < 0 And Me.GetTotalPaymentPurchase > 0 Then
                MsgBox("Please specify a positive value payment.", MsgBoxStyle.Information, "Information")
                Me.TxtPaid.Focus()
                Return False
            End If

            'pastikan akun difference adalah expense, other expense atau cogs
            If CDec(Me.TxtDifference.Text) <> 0 Then
                TipeAkun = Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryAccountDifference.Text)
                If TipeAkun <> ClsMasterAkun.EnumTipeAkun.Cost_of_Goods_Sales And TipeAkun <> ClsMasterAkun.EnumTipeAkun.Expense And _
                TipeAkun <> ClsMasterAkun.EnumTipeAkun.Other_Expense Then
                    If MsgBox("The difference account is not expense account type. Continue?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        Me.TxtDifferenceAccountCode.Focus()
                        Return False
                    End If
                End If
            End If

            'pembayaran paling cepat adalah sesuai tanggal invoice atau sr
            MaxDate = Date.MinValue
            For Each rowDGV In Me.dgvPurchaseBalance.Rows
                If CDec(rowDGV.Cells(Me.ColPaymentOriginal.Index).Value) <> 0 Then
                    If CDate(rowDGV.Cells(Me.ColAOPDate.Index).Value) > MaxDate Then
                        MaxDate = CDate(rowDGV.Cells(Me.ColAOPDate.Index).Value)
                    End If
                End If
            Next
            If MaxDate <> Date.MinValue Then
                If MaxDate > Me.dtpDate.Value.Date And Format(MaxDate, "MMM yyyy") <> Format(Me.dtpDate.Value.Date, "MMM yyyy") Then
                    For Each rowDGV In Me.dgvPurchaseBalance.Rows
                        If MaxDate = CDate(rowDGV.Cells(Me.ColAOPDate.Index).Value) And CDec(rowDGV.Cells(Me.ColPaymentOriginal.Index).Value) <> 0 Then
                            If DummyDoc = "" Then
                                DummyDoc = rowDGV.Cells(Me.ColAOPNumber.Index).Value
                            Else
                                DummyDoc = DummyDoc + "," + rowDGV.Cells(Me.ColAOPNumber.Index).Value
                            End If
                        End If
                    Next

                    MsgBox("Minimum payment date is " + Format(MaxDate, "dd MMM yyyy") + " due to document: " + _
                    DummyDoc + ".", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'cek apakah ada yg valuenya bernilai 0
            For Each row In Me.DgvDifference.Rows
                If Me.DgvDifference.Item(Me.ColAccountDiff_Value.Index, row.Index).Value = 0 Or _
                Me.DgvDifference.Item(Me.ColAccountDiff_ValueOriginalCurrency.Index, row.Index).Value = 0 Then
                    MsgBox("Please add non zero number in account no. " + Me.dgvItem.Item(Me.ColAccountDiff_AccountCode.Index, row.Index).Value, _
                    MsgBoxStyle.Information, "Information")
                    Me.DgvDifference.Focus()
                    Return False
                End If
            Next

            'pastikan semua kode akun ada di dalam master akun
            For Each row In Me.DgvDifference.Rows
                If IsNothing(row.Cells(Me.ColAccountDiff_PrimaryAkun.Index).Value) = True Then
                    MsgBox("Invalid account in row " + row.Index.ToString, MsgBoxStyle.Information, "Information")
                    Me.DgvDifference.ClearSelection()
                    row.Selected = True
                    Return False
                Else
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader2(row.Cells(Me.ColAccountDiff_PrimaryAkun.Index).Value) = False Then
                        MsgBox("Invalid account in row " + row.Index.ToString, MsgBoxStyle.Information, "Information")
                        Me.DgvDifference.ClearSelection()
                        row.Selected = True
                        Return False
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
                    Me.TxtRemarks.BackColor = Me.BackColorAll
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
                    Me.TxtRefNo.BackColor = Me.BackColorAll
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

    Private Sub TxtSuppCode_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSuppCode.Enter
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.Buffer = Me.TxtPrimarySupp.Text
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPihakAcuan_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSuppCode.GotFocus
        Try
            If Me.TxtSuppCode.Enabled = True Then
                If Me.TxtSuppCode.ReadOnly = False Then
                    Me.TxtSuppCode.BackColor = Me.BackColorAll
                Else
                    TxtSuppCode.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPihakAcuan_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSuppCode.LostFocus
        Try
            If Me.TxtSuppCode.Enabled = True Then
                If Me.TxtSuppCode.ReadOnly = False Then
                    Me.TxtSuppCode.BackColor = Me.TxtPrimaryJournal.BackColor
                Else
                    TxtSuppCode.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''Find Code No
    Private Sub TxtNoJadwalProduksi_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles TxtSuppCode.KeyDown

        Try
            If e.KeyValue = 13 Then
                If Me.TxtSuppCode.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjSupp.ApakahSudahAda(Me.TxtSuppCode.Text) = True Then
                        Me.TxtPrimarySupp.Text = ObjSupp.GetPrimary(Me.TxtSuppCode.Text)
                        Me.TxtSuppCode.Text = ObjSupp.GetSupplierCode(Me.TxtPrimarySupp.Text)
                        Me.TxtSuppName.Text = ObjSupp.GetSupplierName(Me.TxtPrimarySupp.Text)
                        Me.cboQuickJournal.Focus()

                        Me.IsikanBalance()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.BtnFindSupp.PerformClick()
                    End If
                Else
                    Me.BtnFindSupp.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNoJadwalProduksi_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtSuppCode.Leave
        Dim i As New FrmCariMasterSupplier
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtSuppCode.Text <> "" Then
                If ObjSupp.ApakahSudahAda(Me.TxtSuppCode.Text) = True Then
                    Me.TxtPrimarySupp.Text = ObjSupp.GetPrimary(Me.TxtSuppCode.Text)
                    Me.TxtSuppCode.Text = ObjSupp.GetSupplierCode(Me.TxtPrimarySupp.Text)
                    Me.TxtSuppName.Text = ObjSupp.GetSupplierName(Me.TxtPrimarySupp.Text)
                Else
                    If Me.TxtSuppCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtSuppCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimarySupplier > 0 Then
                        Me.TxtPrimarySupp.Text = i.PrimarySupplier
                        Me.TxtSuppCode.Text = ObjSupp.GetSupplierCode(Me.TxtPrimarySupp.Text)
                        Me.TxtSuppName.Text = ObjSupp.GetSupplierName(Me.TxtPrimarySupp.Text)
                    Else
                        Me.TxtSuppCode.Clear()
                        Me.TxtSuppName.Clear()
                        Me.TxtPrimarySupp.Text = "-1"
                        Me.TxtSuppCode.Focus()
                    End If
                End If
            Else
                Me.TxtSuppCode.Clear()
                Me.TxtSuppName.Clear()
                Me.TxtPrimarySupp.Text = "-1"
            End If

            'isikan data balance, jika customer berubah atau jika masih kosong semua
            If Me.TxtPrimarySupp.Text <> Me.Buffer Or Me.GetTotalPaymentPurchase_Original = 0 Then
                Me.IsikanBalance()
            End If
            If Me.TxtRemarks.Text.Trim = "" And Me.TxtPrimarySupp.Text <> "-1" Then
                Me.TxtRemarks.Text = "AP Payment: " + Me.TxtSuppCode.Text + " - " + Me.TxtSuppName.Text
            ElseIf Me.TxtRemarks.Text.Trim.Contains("AP Payment: " + Me.TxtSuppCode.Text + " - " + Me.TxtSuppName.Text) = False Then
                Me.TxtRemarks.Text = "AP Payment: " + Me.TxtSuppCode.Text + " - " + Me.TxtSuppName.Text
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub BtnFindSchedule_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles BtnFindSupp.Click
        Dim i As New FrmCariMasterSupplier
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtSuppCode.Text <> "" Then
                i.txtCari.Text = Me.TxtSuppCode.Text
            End If

            i.ShowDialog()
            If i.PrimarySupplier > 0 Then
                Me.TxtPrimarySupp.Text = i.PrimarySupplier
                Me.TxtSuppCode.Text = ObjSupp.GetSupplierCode(Me.TxtPrimarySupp.Text)
                Me.TxtSuppName.Text = ObjSupp.GetSupplierName(Me.TxtPrimarySupp.Text)
                Me.cboQuickJournal.Focus()
            Else
                Me.TxtSuppCode.Clear()
                Me.TxtSuppName.Clear()
                Me.TxtPrimarySupp.Text = "-1"
                Me.TxtSuppCode.Focus()
            End If

            Me.IsikanBalance()
            If Me.TxtRemarks.Text.Trim = "" And Me.TxtPrimarySupp.Text <> "-1" Then
                Me.TxtRemarks.Text = "AP Payment: " + Me.TxtSuppCode.Text + " - " + Me.TxtSuppName.Text
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub IsikanBalance()
        Try
            If Me.TxtPrimarySupp.Text = "" Or Me.TxtPrimarySupp.Text = "-1" Then
                Me.dgvPurchaseBalance.Rows.Clear()
            Else
                If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                    Me.Obj_AOPBalance.IsiDGVBalance(Me.dgvPurchaseBalance, Me.TBL_AOP_Balance, Me, Me.ObjCD, Me.TxtPrimarySupp.Text, _
                    Me.TxtPrimaryJournal.Text, Me.DefaultPrecision)
                ElseIf Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchaseReturnReceivement Then
                    Me.Obj_AOPR_Balance.IsiDGVBalance(Me.dgvPurchaseReturnBalance, Me.TBL_AOPR_Balance, Me, Me.ObjCD, Me.TxtPrimarySupp.Text, _
                    Me.TxtPrimaryJournal.Text, Me.DefaultPrecision)
                End If
            End If

            Me.SetCBOPayment()
            Me.UpdateJurnal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Dim PrimaryAkunDiff As String = ""

        Try
            'If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
            'MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
            'MsgBoxStyle.Information, "Information")
            'Exit Sub
            'End If

            Me.dgvItem.AllowUserToAddRows = False
            Me.Baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.TxtRefNo.Focus()
            Me.ObjJurnal.ResetObject(Me.dgvItem)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

            'difference account number harus disesuaikan dengan setting
            PrimaryAkunDiff = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            ClsTransaksiSetting.EnumNamaParameter.Akun_PaymentDifference)

            If PrimaryAkunDiff <> "" And PrimaryAkunDiff <> "-1" Then
                Me.TxtPrimaryAccountDifference.Text = PrimaryAkunDiff
                Me.TxtDifferenceAccountCode.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryAccountDifference.Text)
            End If
            Me.DgvDifference.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            Me.dgvItem.AllowUserToAddRows = False
            Me.DgvDifference.AllowUserToAddRows = False
            If Me.TxtRefNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtRefNo.Text <> "" Then
                If Me.ObjJurnal.ApakahBolehEdit(Me.ObjJurnal.Primary, UserCode, True, True) = True Then
                    Me.Baru = False
                    Me.ObjJurnal.CariDetailPurchaseJournal(Me.ObjJurnal.Primary, Me.dgvItem, Me, TipeJournal, Me.DgvDifference)
                    Me.SalinkeField()
                    Me.AktifkanDetail()
                    Me.dgvItem.AllowUserToAddRows = False
                    Me.DgvDifference.AllowUserToAddRows = True
                    Me.TxtRefNo.Focus()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            Me.dgvItem.AllowUserToAddRows = False
            Me.DgvDifference.AllowUserToAddRows = False
            If Me.TxtRefNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtRefNo.Text <> "" Then
                If Me.ObjJurnal.ApakahBolehHapus(Me.TxtPrimaryJournal.Text, True, UserCode, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjJurnal.Hapus_PurchaseJurnal(Me.TxtPrimaryJournal.Text, True, UserCode, Me.TipeJournal) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
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

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Dim hasil_save As Boolean = False

        Try
            Me.dgvItem.AllowUserToAddRows = False
            Me.DgvDifference.AllowUserToAddRows = False
            Me.SalinkeObjek()
            If Baru = True Then
                If ObjJurnal.ApakahDataBenar1(Baru, Me.dgvItem, Me.ColPrimaryCurrency.Index, Me.dtpDate.Value, Me.dgvPurchaseBalance, Me.ColPrimaryCurrencyBalance.Index, _
                Me.ColPaymentOriginal.Index, Me.ColTipeDocumentAcuan_Purchase.Index, Me.ColPrimaryDocumentAcuan_Purchase.Index, Me.ColAOPNumber.Index, _
             Me.TxtPrimarySupp.Text, Me.ColExchangeRateBalance.Index, Me.ColAOPDate.Index) = True Then
                    If Me.ApakahDataBenar() = True Then
                        'konfirmasi disimpan atau tidak
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")

                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                                hasil_save = Me.ObjJurnal.SimpanBaruPurchaseJournal(Me.dgvItem, Me.UserCode, Me, TipeJournal, Me.dgvPurchaseBalance, Me.DgvDifference)
                            Else
                                hasil_save = Me.ObjJurnal.SimpanBaruPurchaseJournal(Me.dgvItem, Me.UserCode, Me, TipeJournal, Me.dgvPurchaseReturnBalance, Me.DgvDifference)
                            End If
                            If hasil_save = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.NonAktifkanDetail()
                                Me.SalinkeField()
                                Me.btnFind.Focus()
                                Me.AddReminder()
                            End If
                            Me.dgvItem.AllowUserToAddRows = False
                            Me.DgvDifference.AllowUserToAddRows = False
                        ElseIf hasil = MsgBoxResult.No Then
                            Me.dgvItem.AllowUserToAddRows = False
                            Me.DgvDifference.AllowUserToAddRows = True
                        End If
                    Else
                        Me.dgvItem.AllowUserToAddRows = False
                        Me.DgvDifference.AllowUserToAddRows = True
                    End If
                Else
                    Me.dgvItem.AllowUserToAddRows = False
                    Me.DgvDifference.AllowUserToAddRows = True
                End If
            Else
                'konfirmasi disimpan atau tidak
                If ObjJurnal.ApakahBolehEdit(Me.TxtPrimaryJournal.Text, UserCode, True, True) = True Then
                    If Me.ApakahDataBenar = True Then
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                                hasil_save = Me.ObjJurnal.SimpanPerubahanPurchaseJournal(Me.dgvItem, Me.UserCode, Me, TipeJournal, Me.dgvPurchaseBalance, Me.DgvDifference)
                            Else
                                hasil_save = Me.ObjJurnal.SimpanPerubahanPurchaseJournal(Me.dgvItem, Me.UserCode, Me, TipeJournal, Me.dgvPurchaseReturnBalance, Me.DgvDifference)
                            End If
                            If hasil_save = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.NonAktifkanDetail()
                                Me.SalinkeField()
                                Me.btnEdit.Focus()
                                Me.AddReminder()
                            End If
                            Me.dgvItem.AllowUserToAddRows = False
                            Me.DgvDifference.AllowUserToAddRows = False
                        ElseIf hasil = MsgBoxResult.No Then
                            Me.dgvItem.AllowUserToAddRows = False
                            Me.DgvDifference.AllowUserToAddRows = True
                        End If
                    Else
                        Me.DgvDifference.AllowUserToAddRows = True
                    End If
                Else
                    Me.DgvDifference.AllowUserToAddRows = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.CancelInput()
    End Sub

    Private Sub CancelInput()
        Dim hasil As MsgBoxResult
        Try
            Me.dgvItem.AllowUserToAddRows = False
            Me.DgvDifference.AllowUserToAddRows = False

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
                    Me.NonAktifkanDetail()
                    If Baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.NonAktifkanDetail()
                        Me.ObjJurnal.CariDetailPurchaseJournal(Me.TxtPrimaryJournal.Text, Me.dgvItem, Me, TipeJournal, Me.DgvDifference)
                        Me.SalinkeField()
                    End If
                Else
                    Me.dgvItem.AllowUserToAddRows = False
                    Me.DgvDifference.AllowUserToAddRows = True
                    Exit Sub
                End If
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPO_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim hasil_res As Boolean = False

        Try
            Me.dgvItem.AllowUserToAddRows = False
            Me.DgvDifference.AllowUserToAddRows = False

            If Me.TxtRemarks.Enabled = True Then
                Me.BringToFront()
                If Me.dgvItem.RowCount > 0 Then
                    If Me.ApakahDataBenar = True Then
                        Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            If Me.ApakahDataBenar() = True Then
                                Me.SalinkeObjek()
                                If Baru = True Then
                                    'simpan baru
                                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                                    If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                                        hasil_res = Me.ObjJurnal.SimpanBaruPurchaseJournal(Me.dgvItem, Me.UserCode, Me, Me.TipeJournal, Me.dgvPurchaseBalance, Me.DgvDifference)
                                    Else
                                        hasil_res = Me.ObjJurnal.SimpanBaruPurchaseJournal(Me.dgvItem, Me.UserCode, Me, Me.TipeJournal, Me.dgvPurchaseReturnBalance, Me.DgvDifference)
                                    End If
                                    If hasil = True Then
                                        MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                        Me.SalinkeField()
                                        Me.NonAktifkanDetail()
                                        Me.btnFind.Focus()
                                    Else
                                        e.Cancel = True
                                        Me.DgvDifference.AllowUserToAddRows = True
                                    End If
                                Else
                                    'simpan perubahan
                                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                                    If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                                        hasil_res = Me.ObjJurnal.SimpanPerubahanPurchaseJournal(Me.dgvItem, Me.UserCode, Me, Me.TipeJournal, Me.dgvPurchaseBalance, Me.DgvDifference)
                                    Else
                                        hasil_res = Me.ObjJurnal.SimpanPerubahanPurchaseJournal(Me.dgvItem, Me.UserCode, Me, Me.TipeJournal, Me.dgvPurchaseReturnBalance, Me.DgvDifference)
                                    End If
                                    If hasil = True Then
                                        MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                        Me.SalinkeField()
                                        Me.NonAktifkanDetail()
                                        Me.btnEdit.Focus()
                                    Else
                                        e.Cancel = True
                                        Me.DgvDifference.AllowUserToAddRows = True
                                    End If
                                End If
                            Else
                                e.Cancel = True
                                Me.DgvDifference.AllowUserToAddRows = True
                            End If
                        ElseIf hasil = MsgBoxResult.Cancel Then
                            e.Cancel = True
                            Me.DgvDifference.AllowUserToAddRows = True
                        End If
                    End If
                End If
            End If
            Me.dgvItem.AllowUserToAddRows = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''Fungsi DGV
    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColRemarks.Index Then
                e.Cancel = True
                Exit Sub
            Else
                Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            End If

            If e.ColumnIndex <> Me.ColAccount.Index Then
                If e.ColumnIndex <> Me.ColRemarks.Index And e.ColumnIndex <> Me.ColCurrency.Index And e.ColumnIndex <> Me.ColDescription.Index Then
                    If e.ColumnIndex = Me.ColDebet.Index Or e.ColumnIndex = Me.ColKredit.Index Or _
                    e.ColumnIndex = Me.ColDebetOriginal.Index Or e.ColumnIndex = Me.ColCreditOriginal.Index Or _
                    e.ColumnIndex = Me.ColExchangeRate.Index Then
                        'untuk kolom exchange rate, debet original, creadit original tidak bisa diedit untuk akun dengan mata uang idr
                        If Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value = Me.PrimaryBaseCurrency Then
                            If e.ColumnIndex = Me.ColExchangeRate.Index Or _
                            e.ColumnIndex = Me.ColDebetOriginal.Index Or _
                            e.ColumnIndex = Me.ColCreditOriginal.Index Then
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
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                        Exit Sub
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
                        Me.dgvItem.Rows.RemoveAt(e.RowIndex)
                    End If
                    Exit Sub
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
                                Me.dgvItem.Item(Me.ColDescription.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value = Me.ObjAkun.GetPrimary(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value = Me.ObjAkun.GetPrimaryCurrency(Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value)
                                Me.dgvItem.Item(Me.ColCurrency.Index, e.RowIndex).Value = Me.ObjCur.GetCurrencyCode(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
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
                                    Me.dgvItem.Item(Me.ColDescription.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                                    Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value = ij.mPrimaryAkun
                                    Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value = Me.ObjAkun.GetPrimaryCurrency(Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value)
                                    Me.dgvItem.Item(Me.ColCurrency.Index, e.RowIndex).Value = Me.ObjCur.GetCurrencyCode(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
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
                                        Me.dgvItem.Rows.RemoveAt(e.RowIndex)
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
                If e.ColumnIndex = Me.ColDebetOriginal.Index Then
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
                        Me.dgvItem.Item(Me.ColCreditOriginal.Index, e.RowIndex).Value = ObjCD.DigitGrouping(0, DefPrecOriginalCurr)

                        Me.HitungBaseValues(e.RowIndex)

                        If Me.dgvItem.Item(Me.ColDebet.Index, e.RowIndex).Value > 0 Then
                            Me.BalanceDebetKredit(e.RowIndex, True, Rounding, True)
                        Else
                            Me.BalanceDebetKredit(e.RowIndex, False, Rounding, True)
                        End If
                    End If
                End If

                '4. credit original currency
                If e.ColumnIndex = Me.ColCreditOriginal.Index Then
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
                        Me.dgvItem.Item(Me.ColDebetOriginal.Index, e.RowIndex).Value = ObjCD.DigitGrouping(0, DefPrecOriginalCurr)

                        Me.HitungBaseValues(e.RowIndex)

                        If Me.dgvItem.Item(Me.ColDebet.Index, e.RowIndex).Value > 0 Then
                            Me.BalanceDebetKredit(e.RowIndex, True, Rounding, True)
                        Else
                            Me.BalanceDebetKredit(e.RowIndex, False, Rounding, True)
                        End If
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
                'Me.dgvItem.AllowUserToAddRows = False
                'Exit Sub
                'End If
                'End If
                'Next
                'Me.dgvItem.AllowUserToAddRows = False
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

    Private Sub TxtPurchasePayment_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPurchasePayment.Enter
        Try
            If Me.TxtRemarks.Enabled = True Then
                If CDec(Me.TxtPurchasePayment.Text) = 0 Then
                    Me.TxtPurchasePayment.Text = "0"
                    Me.TxtPurchasePayment.SelectionStart = 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPurchasePayment_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPurchasePayment.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtPurchasePayment.ReadOnly = True Then
                    Me.TxtPurchasePayment.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtPurchasePayment.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPurchasePayment_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPurchasePayment.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtPurchasePayment, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPurchasePayment_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPurchasePayment.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtPurchasePayment, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPurchasePayment_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPurchasePayment.Leave
        Try
            If Me.TxtPurchasePayment.Enabled = True Then
                Me.TxtPurchasePayment.Text = Me.ObjCD.DigitGrouping(Me.TxtPurchasePayment.Text, Me.DefaultPrecision)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPurchasePayment_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPurchasePayment.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtPurchasePayment.ReadOnly = True Then
                    Me.TxtPurchasePayment.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtPurchasePayment.BackColor = Me.TxtPrimaryJournal.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPurchaseReturnReceivement_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPurchaseReturnReceivement.Enter
        Try
            If Me.TxtRemarks.Enabled = True Then
                If CDec(Me.TxtPurchaseReturnReceivement.Text) = 0 Then
                    Me.TxtPurchaseReturnReceivement.Text = 0
                    Me.TxtPurchaseReturnReceivement.SelectionStart = 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPurchaseReturnReceivement_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPurchaseReturnReceivement.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtPurchaseReturnReceivement.ReadOnly = True Then
                    Me.TxtPurchaseReturnReceivement.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtPurchaseReturnReceivement.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPurchaseReturnReceivement_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPurchaseReturnReceivement.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtPurchaseReturnReceivement, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPurchaseReturnReceivement_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPurchaseReturnReceivement.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtPurchaseReturnReceivement, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPurchaseReturnReceivement_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPurchaseReturnReceivement.Leave
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.TxtPurchaseReturnReceivement.Text = Me.ObjCD.DigitGrouping(Me.TxtPurchaseReturnReceivement.Text, Me.DefaultPrecision)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPurchaseReturnReceivement_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPurchaseReturnReceivement.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtPurchaseReturnReceivement.ReadOnly = True Then
                    Me.TxtPurchaseReturnReceivement.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtPurchaseReturnReceivement.BackColor = Me.TxtPrimaryJournal.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDistribusiAOP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDistribusiAOP.Click
        Dim Angka As Decimal = 0
        Dim rowDGV As DataGridViewRow

        Try
            Angka = CDec(Me.TxtPurchasePayment.Text)
            For i As Double = 0 To Me.dgvPurchaseBalance.RowCount - 1
                rowDGV = Me.dgvPurchaseBalance.Rows(i)
                rowDGV.Cells(Me.ColPayment.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)

                rowDGV.Cells(Me.ColBalancePurchase.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDGV.Cells(Me.ColPurchaseValue.Index).Value) + _
                CDec(rowDGV.Cells(Me.ColVatPurchaseValue.Index).Value) + CDec(rowDGV.Cells(Me.ColDutyStamp.Index).Value) - _
                CDec(rowDGV.Cells(Me.ColPaid.Index).Value))

                If Angka <> 0 Then
                    If CDec(rowDGV.Cells(Me.ColBalancePurchase.Index).Value) < Angka Then
                        'yang dipindahkan adalah sejumlah balance sales
                        rowDGV.Cells(Me.ColPayment.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColBalancePurchase.Index).Value, Me.DefaultPrecision)
                        Angka -= CDec(rowDGV.Cells(Me.ColBalancePurchase.Index).Value)
                        rowDGV.Cells(Me.ColBalancePurchase.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                    Else
                        rowDGV.Cells(Me.ColPayment.Index).Value = Me.ObjCD.DigitGrouping(Angka, Me.DefaultPrecision)
                        rowDGV.Cells(Me.ColBalancePurchase.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDGV.Cells(Me.ColPurchaseValue.Index).Value) + _
                        CDec(rowDGV.Cells(Me.ColVatPurchaseValue.Index).Value) + CDec(rowDGV.Cells(Me.ColDutyStamp.Index).Value) - _
                        CDec(rowDGV.Cells(Me.ColPaid.Index).Value) - Angka, Me.DefaultPrecision)
                        Angka = 0
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvPurchaseBalance_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvPurchaseBalance.CellBeginEdit
        Try
            'yang boleh di edit adalah receivement original
            If e.ColumnIndex = Me.ColPaymentOriginal.Index Then
                If Me.ApakahBolehMasukkanPayment(e.RowIndex) = True Then
                    Me.Buffer = Me.dgvPurchaseBalance.Item(e.ColumnIndex, e.RowIndex).Value
                    Me.dgvPurchaseBalance.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvPurchaseBalance.Item(e.ColumnIndex, e.RowIndex).Value.ToString)
                Else
                    e.Cancel = True
                    MsgBox("You can only input payment with " + Me.ObjCur.GetCurrencyCode(Core_Module.PrimaryBaseCurrency) + " and " + _
                       Me.ObjCur.GetCurrencyCode(Me.GetPrimaryCurrencyPayment) + " currency", MsgBoxStyle.Information, "Information")
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'payment bisa diedit jika currency-nya sama atau currencynya akan sama dengan base currency
    Private Function ApakahBolehMasukkanPayment(ByVal vRowIndex As Integer) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim PrimaryCurrencyAcuan As Double = -1

        Try
            rowDGV = Me.dgvPurchaseBalance.Rows(vRowIndex)


            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            'ditutup karena instruksi pak iwan >> eur bisa pay with usd, jadi bisa pay currency manapun
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

            ''jika sama dengan primary base currency maka bisa diedit
            'If rowDGV.Cells(Me.ColPrimaryCurrencyBalance.Index).Value = Core_Module.PrimaryBaseCurrency Then
            '    Return True
            'End If

            ''jika tidak ada currency lain (non-idr) yang beda dengan yang akan diedit
            'PrimaryCurrencyAcuan = rowDGV.Cells(Me.ColPrimaryCurrencyBalance.Index).Value
            'For Each rowDGV In Me.dgvPurchaseBalance.Rows
            '    If rowDGV.Index <> vRowIndex Then
            '        If CDec(rowDGV.Cells(Me.ColPaymentOriginal.Index).Value) <> 0 Then
            '            If rowDGV.Cells(Me.ColPrimaryCurrencyBalance.Index).Value <> PrimaryCurrencyAcuan And _
            '            rowDGV.Cells(Me.ColPrimaryCurrencyBalance.Index).Value <> Core_Module.PrimaryBaseCurrency Then
            '                Return False
            '            End If
            '        End If
            '    End If
            'Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvPurchaseBalance_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvPurchaseBalance.CellContentClick
        Try
            If Me.TxtRemarks.Enabled = True Then
                If e.ColumnIndex = Me.Col_Purchase_Pay.Index Then
                    If Me.ApakahBolehMasukkanPayment(e.RowIndex) = True Then
                        Me.SetFullPayment(e.RowIndex)
                        If Me.ApakahAdaCurrency(Me.dgvPurchaseBalance.Item(Me.ColPrimaryCurrencyBalance.Index, e.RowIndex).Value) = False Then
                            Me.SetCBOPayment()
                        End If
                        Me.UpdateJurnal()
                    Else
                        MsgBox("You can only input payment with " + Me.ObjCur.GetCurrencyCode(Core_Module.PrimaryBaseCurrency) + " and " + _
                         Me.ObjCur.GetCurrencyCode(Me.GetPrimaryCurrencyPayment) + " currency", MsgBoxStyle.Information, "Information")
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub Pay(ByVal vPrimarySupplier As Double, ByVal vTipeDokumenBalancePurchase As ClsTransaksiSalesBalance.EnumTipeSalesBalance, _
    ByVal vPrimaryDokumentAcuan As Double)
        Dim rowDGV As DataGridViewRow

        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.CancelInput()
            End If

            If Me.TxtRemarks.Enabled = False Then
                Me.btnNew.PerformClick()
                Me.TxtPrimarySupp.Text = vPrimarySupplier

                Me.TxtSuppCode.Text = ObjSupp.GetSupplierCode(Me.TxtPrimarySupp.Text)
                Me.TxtSuppName.Text = ObjSupp.GetSupplierName(Me.TxtPrimarySupp.Text)

                Me.IsikanBalance()

                'lakukan pembayaran
                If Me.cboQuickJournal.Items.Count >= 2 Then
                    Me.cboQuickJournal.SelectedIndex = 1
                    'Me.InsertQuickJurnal(False)
                End If

                For Each rowDGV In Me.dgvPurchaseBalance.Rows
                    If rowDGV.Cells(Me.ColTipeDocumentAcuan_Purchase.Index).Value = vTipeDokumenBalancePurchase And _
                    vPrimaryDokumentAcuan = rowDGV.Cells(Me.ColPrimaryDocumentAcuan_Purchase.Index).Value Then
                        If Me.ApakahBolehMasukkanPayment(rowDGV.Index) = True Then
                            Me.SetFullPayment(rowDGV.Index)
                            If Me.ApakahAdaCurrency(Me.dgvPurchaseBalance.Item(Me.ColPrimaryCurrencyBalance.Index, rowDGV.Index).Value) = False Then
                                Me.SetCBOPayment()
                            End If
                            Me.UpdateJurnal()
                        End If
                    End If
                Next

                If Me.Size.Height < Me.MinimumSize.Height Or Me.Size.Width < Me.MinimumSize.Width Then
                    Me.Size = Me.MinimumSize
                End If
                Me.BringToFront()

                Me.TxtRemarks.Text = "AP Payment: " + Me.TxtSuppCode.Text + " - " + Me.TxtSuppName.Text
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetFullPayment(ByVal vRowIndex As Integer)
        Dim row As DataGridViewRow
        Dim DefPrec As Integer = 0

        Try
            DefPrec = Me.ObjCur.GetDecimal_DefaultCurrency

            'hitung ulang pembyaran sampai lunas
            row = Me.dgvPurchaseBalance.Rows(vRowIndex)
            row.Cells(Me.ColBalancePurchase.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
            'hitung pada bagian original
            row.Cells(Me.ColPaymentOriginal.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColBalanceOriginal1.Index).Value) - _
            CDec(row.Cells(Me.ColPaidOriginal.Index).Value), Me.ObjCur.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrencyBalance.Index).Value))
            'row.Cells(Me.ColPayment.Index).Value = _
            'Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColTotalPurchase_VAT_Value.Index).Value) - _
            'CDec(row.Cells(Me.ColPaid.Index).Value), DefPrec)
            row.Cells(Me.ColPayment.Index).Value = _
            Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColPaymentOriginal.Index).Value) * _
            CDec(row.Cells(Me.ColExchangeRateBalance.Index).Value), DefPrec)
            'update data purchase payment
            Me.TxtPurchasePayment.Text = Me.ObjCD.DigitGrouping(Me.GetTotalPaymentPurchase(), Me.DefaultPrecision)

            row.Cells(Me.ColBalanceOriginal.Index).Value = Me.ObjCD.DigitGrouping(0, row.Cells(Me.ColCurrencyDefPrec.Index).Value)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvPurchaseBalance_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvPurchaseBalance.CellEndEdit


        Try
            'yang boleh di edit adalah
            If e.ColumnIndex = Me.ColPaymentOriginal.Index Then
                If IsNothing(Me.dgvPurchaseBalance.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    'langsung di balikin ke nilai semula
                    Me.dgvPurchaseBalance.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Else
                    If e.ColumnIndex = Me.ColPaymentOriginal.Index Then
                        If CDec(Me.dgvPurchaseBalance.Item(Me.ColTotalPurchase_VAT_Value.Index, _
                        e.RowIndex).Value) > 0 Then
                            If Me.ObjCD.CekNonNegatif(Me.dgvPurchaseBalance.Item(e.ColumnIndex, _
                                e.RowIndex).Value) = False Then
                                MsgBox("Please enter positive number", MsgBoxStyle.Information, _
                                "Information")
                                Me.dgvPurchaseBalance.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                        Else
                            If Me.ObjCD.CekNonPositif(Me.dgvPurchaseBalance.Item(e.ColumnIndex, _
                                e.RowIndex).Value) = False Then
                                MsgBox("Please enter negative number", MsgBoxStyle.Information, _
                                "Information")
                                Me.dgvPurchaseBalance.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                        End If

                        Me.UpdatePerhitunganPayment(e.RowIndex, True)
                        Me.UpdateJurnal()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'vSetCBOPayment = false ketika dipanggil dari revaluasi
    'vSetCBOPayment = false ketika dipanggil dari cell end edit -> penggunaan normal
    Private Sub UpdatePerhitunganPayment(ByVal vRowIndex As Integer, ByVal vSetCBOPayment As Boolean)
        Dim ColIndex As Integer = 0
        Dim Balance As Decimal = 0
        Dim DefPrec As Integer = 0
        Dim DefPrecPrimaryCur As Integer = 0

        Try
            ColIndex = Me.ColPaymentOriginal.Index
            DefPrec = Me.dgvPurchaseBalance.Item(Me.ColCurrencyDefPrec.Index, vRowIndex).Value
            DefPrecPrimaryCur = Me.ObjCur.GetDecimal_DefaultCurrency

            'rounding 0
            Me.dgvPurchaseBalance.Item(ColIndex, vRowIndex).Value = _
            Me.ObjCD.DigitGrouping(Me.dgvPurchaseBalance.Item(ColIndex, vRowIndex).Value, _
            DefPrec)

            'pastikan nilai balance tidak menjadi negatif
            Balance = CDec(Me.dgvPurchaseBalance.Item(Me.ColBalanceOriginal1.Index, vRowIndex).Value) - _
            CDec(Me.dgvPurchaseBalance.Item(Me.ColPaidOriginal.Index, vRowIndex).Value) - _
            CDec(Me.dgvPurchaseBalance.Item(Me.ColPaymentOriginal.Index, vRowIndex).Value)

            If CDec(Me.dgvPurchaseBalance.Item(Me.ColTotalPurchase_VAT_Value.Index, _
            vRowIndex).Value) > 0 Then
                If Balance < 0 Then
                    MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping((CDec(Me.dgvPurchaseBalance.Item(Me.ColPaymentOriginal.Index, vRowIndex).Value) + Balance), DefPrec) + ".", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvPurchaseBalance.Item(ColIndex, vRowIndex).Value = Me.Buffer
                    Exit Sub
                End If
            Else
                If Balance > 0 Then
                    MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping((CDec(Me.dgvPurchaseBalance.Item(Me.ColPaymentOriginal.Index, vRowIndex).Value) + Balance), DefPrec) + ".", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvPurchaseBalance.Item(ColIndex, vRowIndex).Value = Me.Buffer
                    Exit Sub
                End If
            End If

            'hitung ulang nilai balance original
            Me.dgvPurchaseBalance.Item(Me.ColBalanceOriginal.Index, vRowIndex).Value = Me.ObjCD.DigitGrouping(Balance, DefPrec)
            'hitung receivement
            Me.dgvPurchaseBalance.Item(Me.ColPayment.Index, vRowIndex).Value = _
            ObjCD.DigitGrouping(CDec(Me.dgvPurchaseBalance.Item(Me.ColExchangeRateBalance.Index, vRowIndex).Value) * _
            Me.dgvPurchaseBalance.Item(Me.ColPaymentOriginal.Index, vRowIndex).Value, DefPrecPrimaryCur)
            'hitung balance
            Balance = CDec(Me.dgvPurchaseBalance.Item(Me.ColBalance1.Index, vRowIndex).Value) - _
            CDec(Me.dgvPurchaseBalance.Item(Me.ColPayment.Index, vRowIndex).Value) - CDec(Me.dgvPurchaseBalance.Item(Me.ColPaid.Index, vRowIndex).Value)
            Me.dgvPurchaseBalance.Item(Me.ColBalancePurchase.Index, vRowIndex).Value = Me.ObjCD.DigitGrouping(Balance, DefPrecPrimaryCur)

            'digit grouping
            Me.dgvPurchaseBalance.Item(ColIndex, vRowIndex).Value = _
            ObjCD.DigitGrouping(Me.dgvPurchaseBalance.Item(ColIndex, vRowIndex).Value, DefPrec)
            'update data sales receivement
            Me.TxtPurchasePayment.Text = Me.ObjCD.DigitGrouping(Me.GetTotalPaymentPurchase, DefPrec)

            If Me.ApakahAdaCurrency(Me.dgvPurchaseBalance.Item(Me.ColPrimaryCurrencyBalance.Index, vRowIndex).Value) = False Then
                If vSetCBOPayment = True Then
                    Me.SetCBOPayment()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDistribusiPR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDistribusiPR.Click
        Dim Angka As Decimal = 0
        Dim rowDGV As DataGridViewRow

        Try
            Angka = Me.TxtPurchaseReturnReceivement.Text
            For i As Double = 0 To Me.dgvPurchaseReturnBalance.RowCount - 1
                rowDGV = Me.dgvPurchaseReturnBalance.Rows(i)
                rowDGV.Cells(Me.ColReturnReceivement.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)

                rowDGV.Cells(Me.ColBalanceReturn.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDGV.Cells(Me.ColPurchaseReturnValue.Index).Value) + _
                CDec(rowDGV.Cells(Me.ColVATReturnValue.Index).Value) + CDec(rowDGV.Cells(Me.ColDutystampPR.Index).Value) - _
                CDec(rowDGV.Cells(Me.ColReceivedReturn.Index).Value))

                If Angka <> 0 Then
                    If CDec(rowDGV.Cells(Me.ColBalanceReturn.Index).Value) < Angka Then
                        'yang dipindahkan adalah sejumlah balance sales
                        rowDGV.Cells(Me.ColReturnReceivement.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColBalanceReturn.Index).Value, Me.DefaultPrecision)
                        Angka -= CDec(rowDGV.Cells(Me.ColBalanceReturn.Index).Value)
                        rowDGV.Cells(Me.ColBalanceReturn.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                    Else
                        rowDGV.Cells(Me.ColReturnReceivement.Index).Value = Me.ObjCD.DigitGrouping(Angka, Me.DefaultPrecision)
                        rowDGV.Cells(Me.ColBalanceReturn.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDGV.Cells(Me.ColPurchaseReturnValue.Index).Value) + _
                        CDec(rowDGV.Cells(Me.ColVATReturnValue.Index).Value) + CDec(rowDGV.Cells(Me.ColDutystampPR.Index).Value) - _
                        CDec(rowDGV.Cells(Me.ColReceivedReturn.Index).Value) - Angka, Me.DefaultPrecision)
                        Angka = 0
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvPurchaseReturnBalance_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvPurchaseReturnBalance.CellBeginEdit
        Try
            'yang boleh di edit adalah data payment saja
            If e.ColumnIndex = Me.ColReturnReceivement.Index Then
                Me.Buffer = Me.dgvPurchaseReturnBalance.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvPurchaseReturnBalance_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvPurchaseReturnBalance.CellEndEdit
        Dim Balance As Decimal = 0

        Try
            'yang boleh di edit adalah
            If e.ColumnIndex = Me.ColReturnReceivement.Index Then
                If IsNothing(Me.dgvPurchaseReturnBalance.Item(e.ColumnIndex, e.RowIndex)) = True Then
                    'langsung di balikin ke nilai semula
                    Me.dgvPurchaseReturnBalance.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Else
                    ''cek debet
                    If e.ColumnIndex = Me.ColReturnReceivement.Index Then
                        If Me.ObjCD.CekNonNegatif(Me.dgvPurchaseReturnBalance.Item(e.ColumnIndex, _
                            e.RowIndex).Value) = False Then
                            MsgBox("Please enter positive number", MsgBoxStyle.Information, _
                            "Information")
                            Me.dgvPurchaseReturnBalance.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        'rounding 0
                        Me.dgvPurchaseReturnBalance.Item(e.ColumnIndex, e.RowIndex).Value = _
                        Me.ObjCD.DigitGrouping(Me.dgvPurchaseReturnBalance.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefaultPrecision)
                        'pastikan nilai balance tidak menjadi negatif
                        Balance = CDec(Me.dgvPurchaseReturnBalance.Item(Me.ColPurchaseReturnValue.Index, e.RowIndex).Value) + _
                        CDec(Me.dgvPurchaseReturnBalance.Item(Me.ColVATReturnValue.Index, e.RowIndex).Value) + _
                        CDec(Me.dgvPurchaseReturnBalance.Item(Me.ColDutystampPR.Index, e.RowIndex).Value) - _
                        CDec(Me.dgvPurchaseReturnBalance.Item(Me.ColReceivedReturn.Index, e.RowIndex).Value)
                        If Balance - CDec(Me.dgvPurchaseReturnBalance.Item(e.ColumnIndex, e.RowIndex).Value) < 0 Then
                            MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(Balance, Me.DefaultPrecision) + ".", MsgBoxStyle.Information, _
                            "Information")
                            Me.dgvPurchaseReturnBalance.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        'hitung ulang nilai balance
                        Me.dgvPurchaseReturnBalance.Item(Me.ColBalanceReturn.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Balance - _
                        CDec(Me.dgvPurchaseReturnBalance.Item(e.ColumnIndex, e.RowIndex).Value), Me.DefaultPrecision)
                        'digit grouping
                        Me.dgvPurchaseReturnBalance.Item(e.ColumnIndex, e.RowIndex).Value = _
                        ObjCD.DigitGrouping(Me.dgvPurchaseReturnBalance.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefaultPrecision)
                        'update data sales return payment
                        Me.TxtPurchaseReturnReceivement.Text = Me.ObjCD.DigitGrouping(Me.GetTotalReceivementReturn, Me.DefaultPrecision)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetTotalPaymentPurchase_Original() As Decimal
        Dim rowdgv As DataGridViewRow
        Dim retval As Decimal = 0

        Try
            For Each rowdgv In Me.dgvPurchaseBalance.Rows
                retval += CDec(rowdgv.Cells(Me.ColPaymentOriginal.Index).Value)
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetTotalPaymentPurchase() As Decimal
        Dim rowdgv As DataGridViewRow
        Dim retval As Decimal = 0

        Try
            For Each rowdgv In Me.dgvPurchaseBalance.Rows
                retval += CDec(rowdgv.Cells(Me.ColPayment.Index).Value)
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetTotalReceivementReturn() As Decimal
        Dim rowdgv As DataGridViewRow
        Dim retval As Decimal = 0

        Try
            For Each rowdgv In Me.dgvPurchaseReturnBalance.Rows
                retval += CDec(rowdgv.Cells(Me.ColReturnReceivement.Index).Value)
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetTotalDebet() As Decimal
        Dim rowdgv As DataGridViewRow
        Dim retval As Decimal = 0

        Try
            For Each rowdgv In Me.dgvItem.Rows
                retval += CDec(rowdgv.Cells(Me.ColDebet.Index).Value)
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub AddReminder()
        Dim ij As New FrmTransaksiReminder_New

        Try
            If MsgBox("Add reminder?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                ij.UserCode = Me.UserCode
                ij.BackColorAll = Me.BackColorAll
                ij.Baru = True
                ij.dtpDate.Value = Me.dtpDate.Value.Date

                ij.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij = Nothing
        End Try
    End Sub

    Private Sub btnAddReminder_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddReminder.Click
        Try
            Me.AddReminder()
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
                    Me.HitungOriginalValues(rowTarget.Index)
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
                            Me.HitungOriginalValues(rowTarget.Index)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetTotalDebet2() As Decimal
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
                Me.ColDebetOriginal.Visible = True
                Me.ColCreditOriginal.Visible = True
                Me.ColCurrency.Visible = True
            Else
                Me.ColCurrency.Visible = False
                Me.ColExchangeRate.Visible = False
                Me.ColDebetOriginal.Visible = False
                Me.ColCreditOriginal.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungExchangeRate(ByVal vRowIndex As Integer)
        Dim vRowDGV As DataGridViewRow

        Try
            vRowDGV = Me.dgvItem.Rows(vRowIndex)
            vRowDGV.Cells(Me.ColExchangeRate.Index).Value = ObjCD.DigitGrouping(ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColDebet.Index).Value) / _
            CDec(vRowDGV.Cells(Me.ColDebetOriginal.Index).Value), DefPrecExchangeRate))
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

            If CDec(vRowDGV.Cells(Me.ColDebetOriginal.Index).Value) > 0 Then
                vRowDGV.Cells(Me.ColKredit.Index).Value = ObjCD.DigitGrouping(0, DefPrecBase)
                vRowDGV.Cells(Me.ColDebet.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColDebetOriginal.Index).Value) * _
                CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value), DefPrecBase)
            Else
                vRowDGV.Cells(Me.ColKredit.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColCreditOriginal.Index).Value) * _
                CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value), DefPrecBase)
                vRowDGV.Cells(Me.ColDebet.Index).Value = ObjCD.DigitGrouping(0, DefPrecBase)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungOriginalValues(ByVal vRowIndex As Integer)
        Dim vRowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0
        Dim DefPrecOriginal As Integer = 0

        Try
            vRowDGV = Me.dgvItem.Rows(vRowIndex)
            DefPrecBase = Me.ObjCur.GetDefaultPrecission(Me.PrimaryBaseCurrency)
            DefPrecOriginal = Me.ObjCur.GetDefaultPrecission(vRowDGV.Cells(Me.ColPrimaryCurrency.Index).Value)

            If CDec(vRowDGV.Cells(Me.ColDebet.Index).Value) > 0 Then
                vRowDGV.Cells(Me.ColCreditOriginal.Index).Value = ObjCD.DigitGrouping(0, DefPrecOriginal)
                vRowDGV.Cells(Me.ColDebetOriginal.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColDebet.Index).Value) / _
                CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value), DefPrecOriginal)
            Else
                vRowDGV.Cells(Me.ColCreditOriginal.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColKredit.Index).Value) / _
                CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value), DefPrecOriginal)
                vRowDGV.Cells(Me.ColDebetOriginal.Index).Value = ObjCD.DigitGrouping(0, DefPrecOriginal)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub UpdateJurnal(Optional ByVal vUpdatePaymentPerBaris As Boolean = False)
        Dim rowDGV As DataGridViewRow
        Dim TotalAPPaid As Decimal = 0
        Dim PrimaryAkunDifference As String = ""
        Dim rowPayment As DseMasterAkun.tblmaster_akun4Row
        Dim IndexAkunAP As Integer = 0
        Dim DefPrec As Integer = 0
        Dim DefPrecBase As Integer = 0
        Dim TblSelisihExchRate As New DseTransaksiJournal.TblSelisihExchRateDataTable
        Dim rowSelisihExchRate As DseTransaksiJournal.TblSelisihExchRateRow
        Dim Payment As Decimal = 0
        Dim DummyBool As Boolean = False
        Dim TotalSelisihExchangeRate As Decimal = 0

        Try
            If vUpdatePaymentPerBaris = True Then
                'hanya dipanggil ketika user melakukan revaluasi
                For Each rowDGV In Me.dgvPurchaseBalance.Rows
                    If CDec(rowDGV.Cells(Me.ColPaymentOriginal.Index).Value) <> 0 Then
                        Me.UpdatePerhitunganPayment(rowDGV.Index, False)
                    End If
                Next
            End If

            DefPrecBase = Me.ObjCur.GetDecimal_DefaultCurrency

            If Me.TxtRemarks.Enabled = False Then
                Exit Sub
            End If
            'hitung ulang difference
            Me.HitungDifference()

            'kosongkan baris
            Me.dgvItem.Rows.Clear()
            Me.dgvItem.AllowUserToAddRows = False

            'dilanjutkan bila
            '1. tidak ada ar yg dibayarkan
            For Each rowDGV In Me.dgvPurchaseBalance.Rows
                If CDec(rowDGV.Cells(Me.ColPayment.Index).Value) <> 0 Then
                    TotalAPPaid += 1
                    Exit For
                End If
            Next
            If TotalAPPaid = 0 Then
                Exit Sub
            End If
            '2. akun payment sudah ditentukan
            If Me.TxtPrimaryAccountPayTo.Text = "" Or Me.TxtPrimaryAccountPayTo.Text = "-1" Then
                Exit Sub
            End If

            'tentukan dahulu akun difference bila difference ada
            If CDec(Me.TxtDifference.Text) <> 0 Then
                If Me.TxtPrimaryAccountDifference.Text = "-1" Or Me.TxtPrimaryAccountDifference.Text = "" Then
                    PrimaryAkunDifference = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
                    ClsTransaksiSetting.EnumNamaParameter.Akun_PaymentDifference)
                    If PrimaryAkunDifference = "" Or PrimaryAkunDifference = "-1" Then
                        MsgBox("Please specify difference account", MsgBoxStyle.Information, "Information")
                        Me.TxtDifferenceAccountCode.Focus()
                        Exit Sub
                    Else
                        Me.TxtPrimaryAccountDifference.Text = PrimaryAkunDifference
                        Me.TxtDifferenceAccountCode.Text = Me.ObjAkun.GetAccountCode(PrimaryAkunDifference)
                    End If
                End If
            End If

            'akun ap, diisikan yg debet dulu baru yg credit
            For Each rowDGV In Me.dgvPurchaseBalance.Rows
                With rowDGV
                    Dim deffPrec As Integer = 0
                    deffPrec = Me.ObjCur.GetDefaultPrecission(.Cells(Me.ColPrimaryCurrencyBalance.Index).Value)
                    If CDec(rowDGV.Cells(Me.ColPaymentOriginal.Index).Value) > 0 Then
                        IndexAkunAP = Me.GetIndexAkunAP(.Cells(Me.ColPrimaryAkunBalance.Index).Value, True, .Cells(Me.ColExchangeRateBalance.Index).Value)
                        If IndexAkunAP < 0 Then
                            'tambahkan baris baru di debet
                            Me.AddRowItem(.Cells(Me.ColPrimaryAkunBalance.Index).Value, _
                            .Cells(Me.ColExchangeRateBalance.Index).Value, .Cells(Me.ColPrimaryCurrencyBalance.Index).Value, _
                            Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColPayment.Index).Value), Me.DefaultPrecision), 0, _
                            Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColPaymentOriginal.Index).Value), deffPrec), 0)
                        Else
                            'tambahkan debet/kredit saja
                            Me.AddValue(Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColPayment.Index).Value), Me.DefaultPrecision), _
                            Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColPaymentOriginal.Index).Value), deffPrec), _
                            IndexAkunAP, True)
                        End If

                    ElseIf CDec(rowDGV.Cells(Me.ColPaymentOriginal.Index).Value) < 0 Then
                        IndexAkunAP = Me.GetIndexAkunAP(.Cells(Me.ColPrimaryAkunBalance.Index).Value, False, .Cells(Me.ColExchangeRateBalance.Index).Value)
                        If IndexAkunAP < 0 Then
                            'tambahkan baris baru do kredit
                            Me.AddRowItem(.Cells(Me.ColPrimaryAkunBalance.Index).Value, _
                            .Cells(Me.ColExchangeRateBalance.Index).Value, .Cells(Me.ColPrimaryCurrencyBalance.Index).Value, _
                             0, Me.ObjCD.DigitGrouping(-CDec(.Cells(Me.ColPayment.Index).Value), Me.DefaultPrecision), 0, _
                             Me.ObjCD.DigitGrouping(-CDec(.Cells(Me.ColPaymentOriginal.Index).Value), deffPrec))
                        Else
                            'tambahkan debet/kredit saja
                            Me.AddValue(Me.ObjCD.DigitGrouping(-CDec(.Cells(Me.ColPayment.Index).Value), Me.DefaultPrecision), _
                            Me.ObjCD.DigitGrouping(-CDec(.Cells(Me.ColPaymentOriginal.Index).Value), deffPrec), _
                            IndexAkunAP, False)
                        End If
                    End If
                End With
            Next

            'akun payment
            If CDec(Me.TxtPaid.Text) <> 0 Then
                rowPayment = Me.TBLAkun.DefaultView(Me.cboPayToAcc.SelectedIndex).Row
                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                With rowDGV
                    .Cells(Me.ColPrimaryAkun.Index).Value = Me.TxtPrimaryAccountPayTo.Text
                    .Cells(Me.ColAccount.Index).Value = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryAccountPayTo.Text)
                    .Cells(Me.ColDescription.Index).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccountPayTo.Text)
                    .Cells(Me.ColExchangeRate.Index).Value = Me.TxtExchangeRate.Text
                    .Cells(Me.ColRemarks.Index).Value = ""
                    .Cells(Me.ColPrimaryCurrency.Index).Value = rowPayment.primary_currency
                    .Cells(Me.ColCurrency.Index).Value = Me.ObjCur.GetCurrencyCode(.Cells(Me.ColPrimaryCurrency.Index).Value)

                    DefPrec = Me.ObjCur.GetDefaultPrecission(rowPayment.primary_currency)
                    If rowPayment.primary_currency <> Core_Module.PrimaryBaseCurrency Then
                        Payment = ObjCD.DigitGrouping(CDec(Me.TxtPaid.Text) * CDec(Me.TxtExchangeRate.Text), Me.DefaultPrecision)
                    Else
                        Payment = ObjCD.DigitGrouping(CDec(Me.TxtPaid.Text), Me.DefaultPrecision)
                    End If

                    If CDec(Me.TxtPaid.Text) > 0 Then
                        'kredit ada isinya
                        .Cells(Me.ColKredit.Index).Value = Me.ObjCD.DigitGrouping(Payment, Me.DefaultPrecision)
                        .Cells(Me.ColDebet.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                        .Cells(Me.ColCreditOriginal.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.TxtPaid.Text)), DefPrec)
                        .Cells(Me.ColDebetOriginal.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
                    Else
                        'debet ada isinya
                        .Cells(Me.ColDebet.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(-Payment), Me.DefaultPrecision)
                        .Cells(Me.ColKredit.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                        .Cells(Me.ColDebetOriginal.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(-CDec(Me.TxtPaid.Text)), DefPrec)
                        .Cells(Me.ColCreditOriginal.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
                    End If
                End With
            End If

            'akun selisih pembayaran
            rowPayment = Me.TBLAkun.DefaultView(Me.cboPayToAcc.SelectedIndex).Row
            If rowPayment.primary_currency <> Core_Module.PrimaryBaseCurrency Then
                TblSelisihExchRate.Rows.Clear()
                TotalSelisihExchangeRate = 0
                Me.GetSelisihExchRate(TblSelisihExchRate, DefPrecBase)
                For Each rowSelisihExchRate In TblSelisihExchRate.Rows
                    TotalSelisihExchangeRate += rowSelisihExchRate.SelisihExchRate
                Next
                If TotalSelisihExchangeRate <> 0 Then
                    For Each rowSelisihExchRate In TblSelisihExchRate.Rows
                        If rowSelisihExchRate.SelisihExchRate <> 0 Then
                            With rowSelisihExchRate
                                'tambahkan bari baru dalam dgv
                                If .SelisihExchRate > 0 Then
                                    'masukkan ke dalam kredit
                                    Me.AddRowItem(.Primary_Akun_SelisihExchRate, 1, Core_Module.PrimaryBaseCurrency, _
                                    0, Me.ObjCD.DigitGrouping(.SelisihExchRate, Me.DefaultPrecision), 0, Me.ObjCD.DigitGrouping(.SelisihExchRate, Me.DefaultPrecision))
                                Else
                                    'masukkan ke dalam debet
                                    Me.AddRowItem(.Primary_Akun_SelisihExchRate, 1, Core_Module.PrimaryBaseCurrency, _
                                    Me.ObjCD.DigitGrouping(-CDec(.SelisihExchRate), Me.DefaultPrecision), 0, Me.ObjCD.DigitGrouping(-CDec(.SelisihExchRate), Me.DefaultPrecision), 0)
                                End If
                                Me.txtPrimaryExchangeRateGainLoss.Text = .Primary_Akun_SelisihExchRate
                            End With
                        End If
                    Next
                End If
            End If


            'akun difference
            If CDec(Me.TxtDifference.Text) <> 0 Then
                If CDec(Me.TxtDifference.Text) > 0 Then
                    'masuk di debet
                    Me.AddRowItem(Me.TxtPrimaryAccountDifference.Text, 1, Core_Module.PrimaryBaseCurrency, _
                    Me.ObjCD.DigitGrouping(CDec(Me.TxtDifference.Text), Me.DefaultPrecision), 0, Me.ObjCD.DigitGrouping(CDec(Me.TxtDifference.Text), Me.DefaultPrecision), 0)
                Else
                    'masuk di kredit
                    Me.AddRowItem(Me.TxtPrimaryAccountDifference.Text, 1, Core_Module.PrimaryBaseCurrency, _
                    0, Me.ObjCD.DigitGrouping(-CDec(Me.TxtDifference.Text), Me.DefaultPrecision), 0, Me.ObjCD.DigitGrouping(-CDec(Me.TxtDifference.Text), Me.DefaultPrecision))
                End If
            End If

            'tambahkan akun difference yang berlebih
            DummyBool = Me.DgvDifference.AllowUserToAddRows
            Me.DgvDifference.AllowUserToAddRows = False
            For Each rowDGV1 As DataGridViewRow In Me.DgvDifference.Rows
                With rowDGV1
                    Dim deffprec As Integer = 0
                    deffprec = Me.ObjCur.GetDefaultPrecission(.Cells(Me.ColAccountDiff_PrimaryCurrency.Index).Value)
                    If CDec(rowDGV1.Cells(Me.ColAccountDiff_Value.Index).Value) > 0 Then
                        'masukkan di debet
                        Me.AddRowItem(.Cells(Me.ColAccountDiff_PrimaryAkun.Index).Value, .Cells(Me.ColAccountDiff_ExchangeRate.Index).Value, _
                        .Cells(Me.ColAccountDiff_PrimaryCurrency.Index).Value, _
                        Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColAccountDiff_Value.Index).Value), Me.DefaultPrecision), 0, _
                        Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColAccountDiff_ValueOriginalCurrency.Index).Value), deffprec), 0, .Cells(Me.ColAccountDiff_Remarks.Index).Value)
                    ElseIf CDec(rowDGV1.Cells(Me.ColAccountDiff_Value.Index).Value) < 0 Then
                        'masukkan di kredit
                        Me.AddRowItem(.Cells(Me.ColAccountDiff_PrimaryAkun.Index).Value, .Cells(Me.ColAccountDiff_ExchangeRate.Index).Value, _
                        .Cells(Me.ColAccountDiff_PrimaryCurrency.Index).Value, _
                        0, Me.ObjCD.DigitGrouping(-CDec(.Cells(Me.ColAccountDiff_Value.Index).Value), Me.DefaultPrecision), 0, _
                        Me.ObjCD.DigitGrouping(-CDec(.Cells(Me.ColAccountDiff_ValueOriginalCurrency.Index).Value), deffprec), .Cells(Me.ColAccountDiff_Remarks.Index).Value)
                    End If
                End With
            Next
            Me.DgvDifference.AllowUserToAddRows = DummyBool

            Me.HitungUnbalance()
            Me.SetMulticurrency()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TblSelisihExchRate = Nothing
        End Try
    End Sub

    Private Sub GetSelisihExchRate(ByVal TblSelisihExchRate As DseTransaksiJournal.TblSelisihExchRateDataTable, ByVal DefPrecBase As Integer)
        Dim rowSelisihExchRate As DseTransaksiJournal.TblSelisihExchRateRow
        Dim SelisihExchRate As Decimal = 0
        Dim Filter As String = ""
        Dim rowDGV As DataGridViewRow
        Dim PrimaryAkunSelisihExchRate As Double = -1

        Try
            For Each rowDGV In Me.dgvPurchaseBalance.Rows
                SelisihExchRate = 0
                If CDbl(rowDGV.Cells(Me.ColPrimaryCurrencyBalance.Index).Value) <> Core_Module.PrimaryBaseCurrency And _
                CDec(rowDGV.Cells(Me.ColPaymentOriginal.Index).Value) <> 0 Then
                    If Me.ObjAkun.GetPrimaryCurrency(TxtPrimaryAccountPayTo.Text) = CDbl(rowDGV.Cells(Me.ColPrimaryCurrencyBalance.Index).Value) Then
                        SelisihExchRate = Me.ObjCD.DigitGrouping(CDec(rowDGV.Cells(Me.ColPaymentOriginal.Index).Value) * _
                                      (CDec(rowDGV.Cells(Me.ColExchangeRateBalance.Index).Value) - CDec(Me.TxtExchangeRate.Text)), DefPrecBase)
                    End If

                    If SelisihExchRate <> 0 Then
                        PrimaryAkunSelisihExchRate = Me.ObjCur.GetPrimary_Akun_Realize_Exch_Rate(rowDGV.Cells(Me.ColPrimaryCurrencyBalance.Index).Value)
                        Filter = "[Primary Akun SelisihExchRate] = " + PrimaryAkunSelisihExchRate.ToString
                        TblSelisihExchRate.DefaultView.RowFilter = Filter

                        If TblSelisihExchRate.DefaultView.Count = 0 Then
                            'tambahkan baris baru
                            rowSelisihExchRate = TblSelisihExchRate.NewTblSelisihExchRateRow

                            With rowSelisihExchRate
                                .Primary_Akun_SelisihExchRate = PrimaryAkunSelisihExchRate
                                .SelisihExchRate = SelisihExchRate
                            End With

                            TblSelisihExchRate.Rows.Add(rowSelisihExchRate)
                        Else
                            'update jumlah saja
                            rowSelisihExchRate = TblSelisihExchRate.DefaultView(0).Row
                            rowSelisihExchRate.SelisihExchRate += SelisihExchRate
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungDifference()
        Dim TotalAPPaid, TotalPayment As Decimal
        Dim rowDGV As DataGridViewRow
        Dim DefPrec As Integer = 0
        Dim TblSelisihExcRate As New DseTransaksiJournal.TblSelisihExchRateDataTable
        Dim rowSelisihExcRate As DseTransaksiJournal.TblSelisihExchRateRow
        Dim SelisihExcRate As Decimal = 0
        Dim rowAkunPayment As DseMasterAkun.tblmaster_akun4Row
        Dim TotalLain2 As Decimal = 0
        Dim rowPayment As DseMasterAkun.tblmaster_akun4Row

        Try
            For Each rowDGV In Me.dgvPurchaseBalance.Rows
                TotalAPPaid += CDec(rowDGV.Cells(Me.ColPayment.Index).Value)
            Next
            Me.TxtARPaid_IDR.Text = Me.ObjCD.DigitGrouping(TotalAPPaid, Me.DefaultPrecision)


            If Me.cboPayToAcc.SelectedIndex <> -1 Then
                rowAkunPayment = Me.TBLAkun.DefaultView(Me.cboPayToAcc.SelectedIndex).Row
                If rowAkunPayment.primary_currency = Core_Module.PrimaryBaseCurrency Then
                    TotalPayment = ObjCD.DigitGrouping(CDec(Me.TxtPaid.Text), Me.DefaultPrecision)
                Else
                    TotalPayment = ObjCD.DigitGrouping(CDec(Me.TxtPaid.Text) * CDec(Me.TxtExchangeRate.Text), Me.DefaultPrecision)
                End If
                Me.TxtPaid_InIDR.Text = Me.ObjCD.DigitGrouping(TotalPayment, Me.DefaultPrecision)
            Else
                TotalPayment = ObjCD.DigitGrouping(CDec(Me.TxtPaid.Text), Me.DefaultPrecision)
            End If



            If Me.cboPayToAcc.SelectedIndex <> -1 Then
                rowPayment = Me.TBLAkun.DefaultView(Me.cboPayToAcc.SelectedIndex).Row
                If rowPayment.primary_currency <> Core_Module.PrimaryBaseCurrency Then
                    Me.GetSelisihExchRate(TblSelisihExcRate, Me.DefaultPrecision)
                    For Each rowSelisihExcRate In TblSelisihExcRate.Rows
                        SelisihExcRate += rowSelisihExcRate.SelisihExchRate
                    Next
                    Me.TxtSelisihExcRate.Text = Me.ObjCD.DigitGrouping(SelisihExcRate, Me.DefaultPrecision)
                Else
                    Me.TxtSelisihExcRate.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                End If
            Else
                Me.TxtSelisihExcRate.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            End If

            For Each rowDGV In Me.DgvDifference.Rows
                TotalLain2 += CDec(rowDGV.Cells(Me.ColAccountDiff_Value.Index).Value)
            Next

            Me.TxtDifference.Text = Me.ObjCD.DigitGrouping(TotalPayment - TotalAPPaid + SelisihExcRate - TotalLain2, Me.DefaultPrecision)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TblSelisihExcRate = Nothing
        End Try
    End Sub

    Private Sub AddValue(ByVal vValue As Decimal, ByVal vOriginalValue As Decimal, ByVal vRowIndex As Integer, ByVal vDebet As Boolean)
        Dim DefPrec As Integer = 0
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgvItem.Rows(vRowIndex)
            DefPrec = Me.ObjCur.GetDefaultPrecission(rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value)

            If vDebet = True Then
                rowDGV.Cells(Me.ColDebet.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDGV.Cells(Me.ColDebet.Index).Value) + vValue, DefPrec)
                rowDGV.Cells(Me.ColDebetOriginal.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDGV.Cells(Me.ColDebetOriginal.Index).Value) + vOriginalValue, DefPrec)
            Else
                rowDGV.Cells(Me.ColKredit.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDGV.Cells(Me.ColKredit.Index).Value) + vValue, DefPrec)
                rowDGV.Cells(Me.ColCreditOriginal.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDGV.Cells(Me.ColCreditOriginal.Index).Value) + vOriginalValue, DefPrec)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexAkunAP(ByVal vPrimaryAkunAR As Double, ByVal Debet As Boolean, ByVal vExhangeRate As Decimal) As Double
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If Debet = True Then
                    If rowDGV.Cells(Me.ColPrimaryAkun.Index).Value = vPrimaryAkunAR And _
                    CDec(rowDGV.Cells(Me.ColDebet.Index).Value) > 0 And CDec(rowDGV.Cells(Me.ColExchangeRate.Index).Value) = vExhangeRate Then
                        Return rowDGV.Index
                    End If
                Else
                    If rowDGV.Cells(Me.ColPrimaryAkun.Index).Value = vPrimaryAkunAR And _
                    CDec(rowDGV.Cells(Me.ColKredit.Index).Value) > 0 And CDec(rowDGV.Cells(Me.ColExchangeRate.Index).Value) = vExhangeRate Then
                        Return rowDGV.Index
                    End If
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub AddRowItem(ByVal vPrimaryAkun As Double, ByVal vExchangeRate As Decimal, ByVal vPrimaryCurrency As Double, ByVal vDebet As Decimal, _
    ByVal vCredit As Decimal, ByVal vDebetOriginal As Decimal, ByVal vCreditOriginal As Decimal, Optional ByVal vRemarks As String = "")
        Dim rowDGV As DataGridViewRow
        Dim DefPrec As Integer = 0

        Try
            DefPrec = Me.ObjCur.GetDefaultPrecission(vPrimaryCurrency)
            Me.dgvItem.Rows.Add()
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
            With rowDGV
                .Cells(Me.ColPrimaryAkun.Index).Value = vPrimaryAkun
                .Cells(Me.ColAccount.Index).Value = Me.ObjAkun.GetAccountCode(vPrimaryAkun)
                .Cells(Me.ColDescription.Index).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(vPrimaryAkun)
                .Cells(Me.ColExchangeRate.Index).Value = Me.ObjCD.DigitGrouping(vExchangeRate, Core_Module.DefPrecExchangeRate)
                .Cells(Me.ColRemarks.Index).Value = vRemarks
                .Cells(Me.ColPrimaryCurrency.Index).Value = vPrimaryCurrency
                .Cells(Me.ColCurrency.Index).Value = Me.ObjCur.GetCurrencyCode(vPrimaryCurrency)

                .Cells(Me.ColDebet.Index).Value = Me.ObjCD.DigitGrouping(vDebet, Me.ObjCur.GetDecimal_DefaultCurrency)
                .Cells(Me.ColKredit.Index).Value = Me.ObjCD.DigitGrouping(vCredit, Me.ObjCur.GetDecimal_DefaultCurrency)
                .Cells(Me.ColDebetOriginal.Index).Value = Me.ObjCD.DigitGrouping(vDebetOriginal, DefPrec)
                .Cells(Me.ColCreditOriginal.Index).Value = Me.ObjCD.DigitGrouping(vCreditOriginal, DefPrec)
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetCurrencyNonIDR_DGVSalesBalance_Paid() As Double
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvPurchaseBalance.Rows
                If CDec(row.Cells(Me.ColPaymentOriginal.Index).Value) <> 0 Then
                    If CDbl(row.Cells(Me.ColPrimaryCurrencyBalance.Index).Value) <> Core_Module.PrimaryBaseCurrency Then
                        Return row.Cells(Me.ColPrimaryCurrencyBalance.Index).Value
                    End If
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub SetExchangeRate()
        Dim row As DseMasterAkun.tblmaster_akun4Row
        Dim PrimaryCur As Double = -1

        Try
            If Me.cboPayToAcc.SelectedIndex <> -1 Then
                row = Me.TBLAkun.DefaultView(Me.cboPayToAcc.SelectedIndex).Row

                If row.primary_currency <> Core_Module.PrimaryBaseCurrency Then
                    'untuk akun non idr
                    Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping(Me.ObjCur.GetDefaultExhangeRate(row.primary_currency), Core_Module.DefPrecExchangeRate)

                    Me.TxtExchangeRate.ReadOnly = False
                    Me.TxtExchangeRate.TabStop = True
                    Me.TxtExchangeRate.BackColor = Me.TxtPrimaryAccountPayTo.BackColor
                Else
                    'untuk akun idr
                    PrimaryCur = Me.GetCurrencyNonIDR_DGVSalesBalance_Paid
                    If PrimaryCur = -1 Then
                        'di dgv sales balance semua idr, exchange rate =1
                        Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping("1", Core_Module.DefPrecExchangeRate)

                        Me.TxtExchangeRate.ReadOnly = True
                        Me.TxtExchangeRate.TabStop = False
                        Me.TxtExchangeRate.BackColor = Me.txtReadOnly.BackColor
                    Else
                        'ada non idr, exchange rate ikut currency tersebut
                        Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping(Me.ObjCur.GetDefaultExhangeRate(PrimaryCur), Core_Module.DefPrecExchangeRate)

                        Me.TxtExchangeRate.ReadOnly = False
                        Me.TxtExchangeRate.TabStop = True
                        Me.TxtExchangeRate.BackColor = Me.TxtPrimaryAccountPayTo.BackColor
                    End If
                End If
            Else
                Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping("1", Core_Module.DefPrecExchangeRate)
                Me.TxtExchangeRate.ReadOnly = True
                Me.TxtExchangeRate.TabStop = False
                Me.TxtExchangeRate.BackColor = Me.txtReadOnly.BackColor
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboPayToAcc_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPayToAcc.SelectedValueChanged
        Dim row As DseMasterAkun.tblmaster_akun4Row
        Dim DV As DataRowView

        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.cboPayToAcc.SelectedIndex <> -1 Then
                    DV = Me.TBLAkun.DefaultView(Me.cboPayToAcc.SelectedIndex)
                    row = DV.Row
                    Me.TxtPrimaryAccountPayTo.Text = row.Primary
                    Me.SetExchangeRate()
                    Me.TxtPaid.Text = Me.ObjCD.DigitGrouping(Me.TxtPaid.Text, Me.ObjCur.GetDefaultPrecission(row.primary_currency))
                Else
                    Me.TxtPrimaryAccountPayTo.Text = -1
                    Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping("1", Core_Module.DefPrecExchangeRate)
                    'Me.TxtPaid.Text = Me.ObjCD.DigitGrouping("0", Me.ObjCur.GetDecimal_DefaultCurrency) --> bikin masalah ketika klik full payment ketika edit 
                    Me.TxtExchangeRate.ReadOnly = True
                    Me.TxtExchangeRate.TabStop = False
                    Me.TxtExchangeRate.BackColor = Me.txtReadOnly.BackColor
                End If

                Me.SetGBOEquivalent()
                Me.HitungDifference()
                Me.UpdateJurnal()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahAdaCurrency_DGVSalesBalance_Paid(ByVal vPrimaryCurr As Double) As Boolean
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvPurchaseBalance.Rows
                If CDec(row.Cells(Me.ColPaymentOriginal.Index).Value) <> 0 Then
                    If row.Cells(Me.ColPrimaryCurrencyBalance.Index).Value = vPrimaryCurr Then
                        Return True
                    End If
                End If

            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub btnCalc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCalc.Click
        Dim row As DseMasterAkun.tblmaster_akun4Row

        Try
            If Me.TxtPrimaryAccountPayTo.Text <> "" And Me.TxtPrimaryAccountPayTo.Text <> "-1" Then
                row = Me.TBLAkun.DefaultView(Me.cboPayToAcc.SelectedIndex).Row
                Me.TxtPaid.Text = "0"
                Me.HitungDifference()

                If row.primary_currency = Core_Module.PrimaryBaseCurrency Then
                    'jika sama dibayarkan ke akun kas bank dengan primary default currency
                    Me.TxtPaid.Text = Me.TxtARPaid_IDR.Text
                Else

                    If Me.ApakahAdadiferentCurrency(row.primary_currency) = True Then
                        Dim decPaid As Decimal = 0
                        If IsNumeric(Me.TxtExchangeRate.Text) AndAlso CDec(Me.TxtExchangeRate.Text) <> 0 Then
                            decPaid = CDec(((Me.GetTotalPaymentPurchase - CDec(Me.TxtSelisihExcRate.Text)) / CDec(Me.TxtExchangeRate.Text)))
                        End If
                        Me.TxtPaid.Text = Me.ObjCD.DigitGrouping(decPaid, Me.ObjCur.GetDefaultPrecission(row.primary_currency))
                    Else
                        Me.TxtPaid.Text = Me.ObjCD.DigitGrouping(Me.GetTotalPaymentPurchase_Original, _
                           Me.ObjCur.GetDefaultPrecission(row.primary_currency))
                    End If

                    'If Me.ApakahAdaCurrency_DGVSalesBalance_Paid(Core_Module.PrimaryBaseCurrency) = False Then
                    '    '1. berasal dari item dgv sales balance pure non idr
                    '    'jika dibayarkan pada akun non idr, berasal dari item sales balance yg non idr
                    '    Me.TxtPaid.Text = Me.ObjCD.DigitGrouping(Me.GetTotalPaymentPurchase_Original, _
                    '    Me.ObjCur.GetDefaultPrecission(row.primary_currency))
                    'Else
                    '    '2. berasal dari item dgv sales balance campuran
                    '    'jika dibayarkan pada akun non idr, berasal dari item sales balance yg idr + non idr
                    '    Me.TxtPaid.Text = Me.ObjCD.DigitGrouping((CDec(Me.TxtARPaid_IDR.Text) - CDec(Me.TxtSelisihExcRate.Text)) / CDec(Me.TxtExchangeRate.Text), Me.ObjCur.GetDefaultPrecission(row.primary_currency))
                    'End If
                End If

                Me.HitungDifference()
                Me.UpdateJurnal()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDifferenceAccountCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDifferenceAccountCode.KeyDown
        Dim i As New FrmMasterAkun_Cari1

        Try
            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtDifferenceAccountCode.Text) = True Then
                    Me.TxtPrimaryAccountDifference.Text = Me.ObjAkun.GetPrimary(Me.TxtDifferenceAccountCode.Text)
                    Me.dgvItem.Focus()
                Else
                    If Me.TxtDifferenceAccountCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtDifferenceAccountCode.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryAccountDifference.Text = i.mPrimaryAkun
                        Me.TxtDifferenceAccountCode.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryAccountDifference.Text)
                        Me.dgvItem.Focus()
                    Else
                        Me.TxtDifferenceAccountCode.Clear()
                        Me.TxtPrimaryAccountDifference.Text = "-1"
                        Me.TxtDifferenceAccountCode.Focus()
                    End If
                End If
            End If

            Me.UpdateJurnal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtDifferenceAccountCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDifferenceAccountCode.Leave
        Dim i As New FrmMasterAkun_Cari1

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.TxtDifferenceAccountCode.Text <> "" Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtDifferenceAccountCode.Text) = True Then
                    Me.TxtPrimaryAccountDifference.Text = Me.ObjAkun.GetPrimary(Me.TxtDifferenceAccountCode.Text)
                Else
                    If Me.TxtDifferenceAccountCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtDifferenceAccountCode.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryAccountDifference.Text = i.mPrimaryAkun
                        Me.TxtDifferenceAccountCode.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryAccountDifference.Text)
                        Me.dgvItem.Focus()
                    Else
                        Me.TxtDifferenceAccountCode.Clear()
                        Me.TxtPrimaryAccountDifference.Text = "-1"
                        Me.TxtDifferenceAccountCode.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryAccountDifference.Text = "-1"
            End If

            Me.UpdateJurnal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindPayableAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindPayableAccount.Click
        Dim i As New FrmMasterAkun_Cari1
        Dim obj As New ClsMasterAkun1

        Try
            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BackColorAll
            If Me.TxtDifferenceAccountCode.Text <> "" Then
                i.txtCari.Text = Me.TxtDifferenceAccountCode.Text
            End If

            i.ShowDialog()

            If i.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAccountDifference.Text = i.mPrimaryAkun
                Me.TxtDifferenceAccountCode.Text = obj.GetAccountCode(Me.TxtPrimaryAccountDifference.Text)
            Else
                Me.TxtDifferenceAccountCode.Clear()
                Me.TxtDifferenceAccountCode.Focus()
            End If

            Me.UpdateJurnal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub


    Private Sub TxtExchangeRate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExchangeRate.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtExchangeRate, sender, e)
            Me.Buffer = Me.TxtExchangeRate.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExchangeRate.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtExchangeRate.Enabled = True Then
                    If Me.TxtExchangeRate.ReadOnly = False Then
                        Me.TxtExchangeRate.BackColor = Me.BackColorAll
                    Else
                        Me.TxtExchangeRate.BackColor = Me.txtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtExchangeRate.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtExchangeRate, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtExchangeRate.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtExchangeRate, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExchangeRate.Leave
        Dim row As DseMasterAkun.tblmaster_akun4Row

        Try
            If Me.TxtExchangeRate.ReadOnly = False And Me.TxtExchangeRate.Enabled = True Then
                Me.ObjCD.Number_Leave(Me.TxtExchangeRate, sender, e, Core_Module.DefPrecExchangeRate, Me.ObjCD)
                If CDec(Me.TxtExchangeRate.Text) <= 0 Then
                    If Me.cboPayToAcc.SelectedIndex = -1 Then
                        Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping(1, Core_Module.DefPrecExchangeRate)
                    Else
                        row = Me.TBLAkun.DefaultView(Me.cboPayToAcc.SelectedIndex).Row
                        Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping(Me.ObjCur.GetDefaultExhangeRate(row.primary_currency), _
                        Core_Module.DefPrecExchangeRate)
                    End If
                End If

                Me.HitungDifference()
                Me.UpdateJurnal()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExchangeRate.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtExchangeRate.Enabled = True Then
                    If Me.TxtExchangeRate.ReadOnly = False Then
                        Me.TxtExchangeRate.BackColor = Me.TxtPrimaryAccountDifference.BackColor
                    Else
                        Me.TxtExchangeRate.BackColor = Me.txtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaid_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaid.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtPaid, sender, e)
            Me.Buffer = Me.TxtPaid.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaid_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaid.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtPaid.Enabled = True Then
                    If Me.TxtPaid.ReadOnly = False Then
                        Me.TxtPaid.BackColor = Me.BackColorAll
                    Else
                        Me.TxtPaid.BackColor = Me.txtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaid_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPaid.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtPaid, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaid_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPaid.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtPaid, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaid_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaid.Leave
        Dim row As DseMasterAkun.tblmaster_akun4Row

        Try
            If Me.cboPayToAcc.SelectedIndex <> -1 Then
                row = Me.TBLAkun.DefaultView(Me.cboPayToAcc.SelectedIndex).Row
                Me.ObjCD.Number_Leave(Me.TxtPaid, sender, e, Me.ObjCur.GetDefaultPrecission(row.primary_currency), Me.ObjCD)
                If Me.TxtPaid.ReadOnly = False And Me.TxtPaid.Enabled = True Then
                    Me.HitungDifference()
                    Me.UpdateJurnal()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaid_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaid.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtPaid.Enabled = True Then
                    If Me.TxtPaid.ReadOnly = False Then
                        Me.TxtPaid.BackColor = Me.TxtPrimaryAccountDifference.BackColor
                    Else
                        Me.TxtPaid.BackColor = Me.txtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDifferenceAccountCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDifferenceAccountCode.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtDifferenceAccountCode.Enabled = True Then
                    If Me.TxtDifferenceAccountCode.ReadOnly = False Then
                        Me.TxtDifferenceAccountCode.BackColor = Me.BackColorAll
                    Else
                        Me.TxtDifferenceAccountCode.BackColor = Me.txtReadOnly.BackColor
                    End If
                End If
            End If

            If Me.TxtDifferenceAccountCode.Text <> "" Then
                Me.UpdateJurnal()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDifferenceAccountCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDifferenceAccountCode.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtDifferenceAccountCode.Enabled = True Then
                    If Me.TxtDifferenceAccountCode.ReadOnly = False Then
                        Me.TxtDifferenceAccountCode.BackColor = Me.TxtPrimaryAccountDifference.BackColor
                    Else
                        Me.TxtDifferenceAccountCode.BackColor = Me.txtReadOnly.BackColor
                    End If
                End If
            End If

            If Me.TxtDifferenceAccountCode.Text <> "" Then
                Me.UpdateJurnal()
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
                Dim i As New FrmTransaksiJournal_Cetak_Purchase
                i.PrimaryJurnal = Me.TxtPrimaryJournal.Text
                i.vDGV = Me.dgvPurchaseBalance
                i.vFrm = Me

                i.ShowDialog()
                i.Dispose()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindJurnal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindOtherAccount.Click
        Dim PrimaryCurPayment As Double = -1
        Dim i As New FrmMasterAkun_Cari
        Dim PrimaryCurrencyAkun As Double = -1
        Dim ApakahSudahAdaAkun As Boolean = False
        Dim row As DseMasterAkun.tblmaster_akun4Row
        Dim m As Double = 0

        Try
            If Me.ApakahAdaPayment = True And Me.TxtRemarks.Enabled = True Then
                'keluarkan form cari
                i.BackColorAll = Me.BackColorAll
                i.ShowDialog()

                If i.mPrimaryAkun > 0 Then
                    PrimaryCurPayment = Me.GetPrimaryCurrencyPayment
                    'cek currency yang digunakan adalah sama dengan currency pada payment atau
                    'IDR
                    PrimaryCurrencyAkun = Me.ObjAkun.GetPrimaryCurrency(i.mPrimaryAkun)
                    'If PrimaryCurrencyAkun <> Core_Module.PrimaryBaseCurrency And PrimaryCurrencyAkun <> PrimaryCurPayment Then
                    '    MsgBox("Please select payment account with " + Me.ObjCur.GetCurrencyCode(Core_Module.PrimaryBaseCurrency) + _
                    '    " or " + Me.ObjCur.GetCurrencyCode(PrimaryCurPayment) + " currency")
                    '    Exit Sub
                    'End If

                    'jika ada account yang dipilih, maka cek apakah sudah ada di dalam tabel, 
                    'jika belum ada maka tambahkan datanya ke dalam tabel
                    For Each dv As DataRowView In Me.TBLAkun.DefaultView
                        row = dv.Row
                        If row.Primary = i.mPrimaryAkun Then
                            ApakahSudahAdaAkun = True
                            Exit For
                        End If
                    Next

                    If ApakahSudahAdaAkun = False Then
                        'Me.cboPayToAcc.SelectedIndex = -1

                        'tambahkan baris baru pada tabel
                        row = Me.TBLAkun.Newtblmaster_akun4Row
                        row.Primary = i.mPrimaryAkun
                        row.primary_currency = Me.ObjAkun.GetPrimaryCurrency(i.mPrimaryAkun)
                        row.Description = Me.ObjAkun.GetDescription_ByPrimaryAcc(i.mPrimaryAkun)
                        row.Account_Code = Me.ObjAkun.GetAccountCode(i.mPrimaryAkun)
                        Me.TBLAkun.Rows.Add(row)

                        'pilih account yang dipilih pada form cari

                        TBLAkun.DefaultView.Sort = "[Primary]"
                        Me.cboPayToAcc.DataSource = Nothing
                        Me.cboPayToAcc.DataSource = Me.TBLAkun.DefaultView
                        Me.cboPayToAcc.DisplayMember = "Description"

                        For Each dv As Data.DataRowView In Me.TBLAkun.DefaultView
                            row = dv.Row
                            If row.Primary = i.mPrimaryAkun Then
                                Me.cboPayToAcc.SelectedIndex = m
                                Exit For
                            End If
                            m += 1
                        Next

                        'For Each row In Me.TBLAkun.Rows
                        '    If row.Primary = i.mPrimaryAkun Then
                        '        Me.cboPayToAcc.SelectedIndex = m
                        '        Exit For
                        '    End If
                        '    m += 1
                        'Next
                        'row = Me.TBLAkun.Rows(Me.TBLAkun.Rows.Count - 1)

                        'MsgBox(row.Description)
                        ' Me.cboPayToAcc.SelectedIndex = Me.TBLAkun.Rows.Count - 1                      
                    End If


                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub DgvDifference_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvDifference.RowsAdded
        Dim rowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0

        Try
            Me.DgvDifference.TabStop = True
            If Me.TxtRemarks.Enabled = True And Me.DgvDifference.AllowUserToAddRows = True Then
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True

                'set nilai default baris baru
                rowDGV = Me.DgvDifference.Rows(Me.DgvDifference.RowCount - 1)
                rowDGV.Cells(Me.ColAccountDiff_PrimaryCurrency.Index).Value = Me.PrimaryBaseCurrency
                DefPrecBase = Me.ObjCur.GetDefaultPrecission(Me.PrimaryBaseCurrency)
                rowDGV.Cells(Me.ColAccountDiff_ValueOriginalCurrency.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrecBase)
                rowDGV.Cells(Me.ColAccountDiff_Value.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrecBase)
                rowDGV.Cells(Me.ColAccountDiff_ExchangeRate.Index).Value = Me.ObjCD.DigitGrouping(1, DefPrecExchangeRate)
                rowDGV.Cells(Me.ColAccountDiff_PrimaryCurrency.Index).Value = Me.PrimaryBaseCurrency
                rowDGV.Cells(Me.ColAccountDiff_OriginalCurrency.Index).Value = Me.ObjCur.GetCurrencyCode(Me.PrimaryBaseCurrency)
                rowDGV.Cells(Me.ColAccountDiff_PrimaryAkun.Index).Value = -1
                rowDGV.Cells(Me.ColAccountDiff_AccountCode.Index).Value = ""
                rowDGV.Cells(Me.ColAccountDiff_Remarks.Index).Value = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''Fungsi DGV
    Private Sub DgvDifference_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvDifference.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColAccountDiff_AccountCode.Index Then
                If e.ColumnIndex <> Me.ColAccountDiff_Remarks.Index And e.ColumnIndex <> Me.ColAccountDiff_OriginalCurrency.Index And _
                e.ColumnIndex <> Me.ColAccountDiff_Desc.Index Then
                    If e.ColumnIndex = Me.ColAccountDiff_Value.Index Or _
                    e.ColumnIndex = Me.ColAccountDiff_ValueOriginalCurrency.Index Or _
                    e.ColumnIndex = Me.ColAccountDiff_ExchangeRate.Index Then
                        'untuk kolom exchange rate, value original tidak bisa diedit untuk akun dengan mata uang idr
                        If Me.DgvDifference.Item(Me.ColAccountDiff_PrimaryCurrency.Index, e.RowIndex).Value = Me.PrimaryBaseCurrency Then
                            If e.ColumnIndex = Me.ColAccountDiff_ExchangeRate.Index Or _
                            e.ColumnIndex = Me.ColAccountDiff_ValueOriginalCurrency.Index Then
                                e.Cancel = True
                                Exit Sub
                            End If
                        End If

                        If IsNothing(Me.DgvDifference.Item(Me.ColAccountDiff_PrimaryAkun.Index, e.RowIndex).Value) = True Then
                            'bukan baris teakhir
                            e.Cancel = True
                            Exit Sub
                        ElseIf Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader2(Me.DgvDifference.Item(Me.ColAccountDiff_PrimaryAkun.Index, e.RowIndex).Value) = False Then
                            'ada akunnya, aktif, dan bukan header
                            e.Cancel = True
                            Exit Sub
                        Else
                            'value
                            Me.Buffer = Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value
                            'buang nol dibelakang decimal
                            Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.ObjCD.BuangNolDibelakangDecimal(Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value.ToString)
                        End If
                    End If
                Else
                    If Me.DgvDifference.Rows.Count - 1 = e.RowIndex Then
                        e.Cancel = True
                        Exit Sub
                    Else
                        If e.ColumnIndex = Me.ColAccountDiff_Remarks.Index Then
                            Me.Buffer = Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value
                        Else
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If
                End If
            Else
                'col account
                If e.ColumnIndex = Me.ColAccountDiff_AccountCode.Index Then
                    Me.Buffer = Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub DgvDifference_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvDifference.CellEndEdit
        Dim Rounding As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim ExRate As Decimal = 1
        Dim DefPrecOriginalCurr As Integer = 0

        Try
            Rounding = Me.DefaultPrecision
            If IsNothing(Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value) = True _
            Or Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColAccountDiff_Remarks.Index Then
                    'Remarks
                    Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value = ""
                ElseIf e.ColumnIndex = Me.ColAccountDiff_AccountCode.Index Then
                    'kalo kode akun dihapus, maka hapus satu bari semuanya
                    If e.RowIndex <> Me.DgvDifference.RowCount - 1 Then
                        Me.DgvDifference.Rows.RemoveAt(e.RowIndex)
                    End If
                    Exit Sub
                Else
                    'untuk kolom value atau value original, langsung di balikin ke nilai semula
                    Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                End If
            Else
                ''cek value
                If e.ColumnIndex = Me.ColAccountDiff_Value.Index Then
                    'pastikan numeric
                    If IsNumeric(Me.DgvDifference.Item(e.ColumnIndex, _
                    e.RowIndex).Value) = False Then
                        MsgBox("Please enter a number", MsgBoxStyle.Information, _
                        "Information")
                        Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value, Rounding)
                    ''Cek jurnal yang terjadi
                    Me.HitungOriginalValues_Diff(e.RowIndex)
                    Me.UpdateJurnal()
                End If

                ''data account number
                If e.ColumnIndex = Me.ColAccountDiff_AccountCode.Index Then
                    Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    'kemungkinan
                    If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                        Me.DgvDifference.Item(Me.ColAccountDiff_Desc.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription(Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.DgvDifference.Item(Me.ColAccountDiff_PrimaryAkun.Index, e.RowIndex).Value = Me.ObjAkun.GetPrimary(Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.DgvDifference.Item(Me.ColAccountDiff_PrimaryCurrency.Index, e.RowIndex).Value = Me.ObjAkun.GetPrimaryCurrency(Me.DgvDifference.Item(Me.ColAccountDiff_PrimaryAkun.Index, e.RowIndex).Value)
                        Me.DgvDifference.Item(Me.ColAccountDiff_OriginalCurrency.Index, e.RowIndex).Value = Me.ObjCur.GetCurrencyCode(Me.DgvDifference.Item(Me.ColAccountDiff_PrimaryCurrency.Index, e.RowIndex).Value)
                        ExRate = Me.ObjTaxRate.GetTaxRate(Me.dtpDate.Value, Me.DgvDifference.Item(Me.ColAccountDiff_PrimaryCurrency.Index, e.RowIndex).Value)
                        If ExRate = 0 Then
                            ExRate = 1
                        End If
                        Me.DgvDifference.Item(Me.ColAccountDiff_ExchangeRate.Index, e.RowIndex).Value = ObjCD.DigitGrouping(ExRate, DefPrecExchangeRate)
                        'DefPrecOriginalCurr = Me.ObjCur.GetDefaultPrecission(Me.DgvDifference.Item(Me.ColAccountDiff_PrimaryCurrency.Index, e.RowIndex).Value)
                        Me.HitungOriginalValues_Diff(e.RowIndex)
                    Else
                        '2. kode akun yang diketik tidak ada --> tampilkan form cari
                        Dim ij As New FrmMasterAkun_Cari
                        ij.txtCari.Text = Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.mPrimaryAkun <> -1 Then
                            '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                            Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAkun.GetAccountCode(ij.mPrimaryAkun)
                            Me.DgvDifference.Item(Me.ColAccountDiff_Desc.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription(Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value)
                            Me.DgvDifference.Item(Me.ColAccountDiff_PrimaryAkun.Index, e.RowIndex).Value = ij.mPrimaryAkun
                            Me.DgvDifference.Item(Me.ColAccountDiff_PrimaryCurrency.Index, e.RowIndex).Value = Me.ObjAkun.GetPrimaryCurrency(Me.DgvDifference.Item(Me.ColAccountDiff_PrimaryAkun.Index, e.RowIndex).Value)
                            Me.DgvDifference.Item(Me.ColAccountDiff_OriginalCurrency.Index, e.RowIndex).Value = Me.ObjCur.GetCurrencyCode(Me.DgvDifference.Item(Me.ColAccountDiff_PrimaryCurrency.Index, e.RowIndex).Value)
                            ExRate = Me.ObjTaxRate.GetTaxRate(Me.dtpDate.Value, Me.DgvDifference.Item(Me.ColAccountDiff_PrimaryCurrency.Index, e.RowIndex).Value)
                            If ExRate = 0 Then
                                ExRate = 1
                            End If
                            Me.DgvDifference.Item(Me.ColAccountDiff_ExchangeRate.Index, e.RowIndex).Value = ObjCD.DigitGrouping(ExRate, DefPrecExchangeRate)
                            'DefPrecOriginalCurr = Me.ObjCur.GetDefaultPrecission(Me.DgvDifference.Item(Me.ColAccountDiff_PrimaryCurrency.Index, e.RowIndex).Value)
                            Me.HitungOriginalValues_Diff(e.RowIndex)
                        ElseIf ij.mPrimaryAkun = -1 Then
                            '2.b. kode akun yang dicari tidak ada --> kembalikan ke kondisi buffer
                            Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                Me.DgvDifference.Rows.RemoveAt(e.RowIndex)
                                Exit Sub
                            End If
                        End If
                        ij = Nothing
                    End If
                    Me.SetMulticurrency_Diff()
                End If

                '3. exchange rate, hanya jika non base currency
                If e.ColumnIndex = Me.ColAccountDiff_ExchangeRate.Index Then
                    If Me.DgvDifference.Item(Me.ColAccountDiff_PrimaryCurrency.Index, e.RowIndex).Value <> Me.PrimaryBaseCurrency Then
                        If Me.ObjCD.CekBilanganPositif(Me.DgvDifference.Item(e.ColumnIndex, _
                        e.RowIndex).Value) = False Then
                            MsgBox("Please enter positive number", MsgBoxStyle.Information, _
                            "Information")
                            Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        'digit groping
                        Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value = _
                        ObjCD.DigitGrouping(Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value, DefPrecExchangeRate)
                        'hitung original value
                        Me.HitungOriginalValues_Diff(e.RowIndex)
                    End If
                End If

                '3. value original currency
                If e.ColumnIndex = Me.ColAccountDiff_ValueOriginalCurrency.Index Then
                    If Me.DgvDifference.Item(Me.ColAccountDiff_PrimaryCurrency.Index, e.RowIndex).Value <> Me.PrimaryBaseCurrency Then
                        'pastikan numeric
                        If IsNumeric(Me.DgvDifference.Item(e.ColumnIndex, _
                        e.RowIndex).Value) = False Then
                            MsgBox("Please enter a number", MsgBoxStyle.Information, _
                            "Information")
                            Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        'digit groping
                        DefPrecOriginalCurr = Me.ObjCur.GetDefaultPrecission(Me.DgvDifference.Item(Me.ColAccountDiff_PrimaryCurrency.Index, e.RowIndex).Value)
                        Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value = _
                        ObjCD.DigitGrouping(Me.DgvDifference.Item(e.ColumnIndex, e.RowIndex).Value, _
                        DefPrecOriginalCurr)
                        'hitung base value
                        Me.HitungBaseValues_Diff(e.RowIndex)
                    End If
                End If

                Me.UpdateJurnal()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCur = Nothing
        End Try
    End Sub

    Private Sub HitungOriginalValues_Diff(ByVal vRowIndex As Integer)
        Dim vRowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0
        Dim DefPrecOriginal As Integer = 0

        Try
            vRowDGV = Me.DgvDifference.Rows(vRowIndex)
            DefPrecBase = Me.ObjCur.GetDefaultPrecission(Me.PrimaryBaseCurrency)
            DefPrecOriginal = Me.ObjCur.GetDefaultPrecission(vRowDGV.Cells(Me.ColAccountDiff_PrimaryCurrency.Index).Value)

            If CDec(vRowDGV.Cells(Me.ColAccountDiff_Value.Index).Value) <> 0 Then
                vRowDGV.Cells(Me.ColAccountDiff_ValueOriginalCurrency.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColAccountDiff_Value.Index).Value) / _
                CDec(vRowDGV.Cells(Me.ColAccountDiff_ExchangeRate.Index).Value), DefPrecOriginal)
            Else
                vRowDGV.Cells(Me.ColAccountDiff_ValueOriginalCurrency.Index).Value = ObjCD.DigitGrouping(0, DefPrecOriginal)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungExchangeRate_Diff(ByVal vRowIndex As Integer)
        Dim vRowDGV As DataGridViewRow

        Try
            vRowDGV = Me.DgvDifference.Rows(vRowIndex)
            vRowDGV.Cells(Me.ColAccountDiff_ExchangeRate.Index).Value = ObjCD.DigitGrouping(ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColAccountDiff_Value.Index).Value) / _
            CDec(vRowDGV.Cells(Me.ColAccountDiff_ValueOriginalCurrency.Index).Value), DefPrecExchangeRate))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungBaseValues_Diff(ByVal vRowIndex As Integer)
        Dim vRowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0
        Dim DefPrecOriginal As Integer = 0

        Try
            vRowDGV = Me.DgvDifference.Rows(vRowIndex)
            DefPrecBase = Me.ObjCur.GetDefaultPrecission(Me.PrimaryBaseCurrency)
            DefPrecOriginal = Me.ObjCur.GetDefaultPrecission(vRowDGV.Cells(Me.ColAccountDiff_PrimaryCurrency.Index).Value)

            If CDec(vRowDGV.Cells(Me.ColAccountDiff_ValueOriginalCurrency.Index).Value) <> 0 Then
                vRowDGV.Cells(Me.ColAccountDiff_Value.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColAccountDiff_ValueOriginalCurrency.Index).Value) * _
                CDec(vRowDGV.Cells(Me.ColAccountDiff_ExchangeRate.Index).Value), DefPrecBase)
            Else
                vRowDGV.Cells(Me.ColAccountDiff_Value.Index).Value = ObjCD.DigitGrouping(0, DefPrecBase)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetMulticurrency_Diff()
        Dim rowDGV As DataGridViewRow
        Dim Multicurrency As Boolean = False

        Try
            For Each rowDGV In Me.DgvDifference.Rows
                If rowDGV.Index = Me.DgvDifference.RowCount - 1 Then
                    If Me.DgvDifference.AllowUserToAddRows = True Then
                        'tidak dicek
                    Else
                        'dicek
                        If rowDGV.Cells(Me.ColAccountDiff_PrimaryCurrency.Index).Value <> Me.PrimaryBaseCurrency Then
                            Multicurrency = True
                            Exit For
                        End If
                    End If
                Else
                    If rowDGV.Cells(Me.ColAccountDiff_PrimaryCurrency.Index).Value <> Me.PrimaryBaseCurrency Then
                        Multicurrency = True
                        Exit For
                    End If
                End If
            Next

            If Multicurrency = True Then
                Me.ColAccountDiff_ExchangeRate.Visible = True
                Me.ColAccountDiff_ValueOriginalCurrency.Visible = True
                Me.ColAccountDiff_OriginalCurrency.Visible = True
            Else
                Me.ColAccountDiff_OriginalCurrency.Visible = False
                Me.ColAccountDiff_ExchangeRate.Visible = False
                Me.ColAccountDiff_ValueOriginalCurrency.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvDifference_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DgvDifference.UserDeletedRow
        Try
            Me.UpdateJurnal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class