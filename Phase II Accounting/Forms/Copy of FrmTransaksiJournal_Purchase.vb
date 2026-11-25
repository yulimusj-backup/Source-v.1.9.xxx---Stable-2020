Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiJournal_Purchase
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightSeaGreen
    Public TipeJournal As ClsTransaksiJournal.EnumTipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal
    Public ObjQA As New ClsTransaksiJournal

    Dim ObjCD As New Core_Module.ClsCekData
    Dim Obj_AOPBalance As New ClsTransaksiPurchaseBalance
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
    Dim PrimaryBaseCurrency As Double = 1
    Dim ObjTaxRate As New ClsTransaksiSalesInvoiceTaxRate

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
            Me.WindowState = FormWindowState.Maximized
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvPurchaseBalance.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvPurchaseReturnBalance.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                Me.SplitContainer3.Panel2Collapsed = True
                Me.Text = "Purchase Payment Journal"
            Else
                Me.SplitContainer3.Panel1Collapsed = True
                Me.Text = "Purchase Return Receivement Journal"
            End If
            ''Me.AktifkanDetail()
            Me.NonAktifkanDetail()

            Me.dgvItem.AllowUserToAddRows = False
            Me.ObjQA = New ClsTransaksiJournal
            Me.Baru = True
            Me.dgvItem.Rows.Clear()
            Me.KosongkanField()

            'set null value untuk dgv
            Me.DefaultPrecision = Me.ObjCur.GetDecimal_DefaultCurrency
            NullValue = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            Me.ColDebet.DefaultCellStyle.NullValue = NullValue
            Me.ColKredit.DefaultCellStyle.NullValue = NullValue
            Me.TxtUnbalance.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            Me.TxtPurchasePayment.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            Me.TxtPurchaseReturnReceivement.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
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
            If Me.TxtRemarks.Enabled = True Then
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True

                'set nilai default baris baru
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value = Me.PrimaryBaseCurrency
                DefPrecBase = Me.ObjCur.GetDefaultPrecission(Me.PrimaryBaseCurrency)
                rowDGV.Cells(Me.ColDebetOriginal.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrecBase)
                rowDGV.Cells(Me.ColCreditOriginal.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrecBase)
                rowDGV.Cells(Me.ColDebet.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrecBase)
                rowDGV.Cells(Me.ColKredit.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrecBase)
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
                If Me.ObjQA.CariDetailPurchaseJournal(ij.PrimaryPP, Me.dgvItem, Me, ij.TipeJournal) = True Then
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
            Unbalance = Math.Round(vdebet - vkredit, 0, MidpointRounding.AwayFromZero)
            Return Unbalance
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub AktifkanDetail()
        Try
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

            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.GroupBox1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.GroupBox2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, True, True)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None

            Me.dgvItem.AllowUserToDeleteRows = True
            Me.dgvItem.ReadOnly = False
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
            i.BackCOlorAll = Color.FromArgb(170, 255, 190)
            i.MdiParent = FrmParent
            i.Show()
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
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, True, False, True, False)
            Me.CancelButton = Me.btnExit

            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True
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

            Me.cboQuickJournal.SelectedIndex = -1
            Me.llbEditQuickJournal.Enabled = False
            ''Me.cboCodeNo.DataSource = Nothing
            ''Me.cboCodeNo.Items.Clear()
            ''Me.cboProses.DataSource = Nothing
            ''Me.cboProses.Items.Clear()
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

            Me.TxtPrimaryJournal.Text = "-1"
            Me.TxtPrimarySupp.Text = "-1"
            Me.TxtUnbalance.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            Me.TxtPurchasePayment.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            Me.TxtPurchaseReturnReceivement.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "

            Me.dgvItem.Rows.Clear()
            Me.dgvPurchaseBalance.Rows.Clear()
            Me.dgvPurchaseReturnBalance.Rows.Clear()
            ''Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinkeObjek()
        Try
            With ObjQA
                .RefNumber = Me.TxtRefNo.Text
                .Description = Me.TxtRemarks.Text
                .Tanggal = Me.dtpDate.Value
                .PrimaryPihakAcuan = Me.TxtPrimarySupp.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinkeField()
        Try
            With ObjQA
                Me.TxtPrimaryJournal.Text = .Primary
                Me.TxtRefNo.Text = .RefNumber
                Me.TxtRemarks.Text = .Description
                Me.dtpDate.Value = .Tanggal
                Me.TxtPrimarySupp.Text = .PrimaryPihakAcuan
                Me.TxtSuppCode.Text = ObjSupp.GetSupplierCode(Me.TxtPrimarySupp.Text)
                Me.TxtSuppName.Text = ObjSupp.GetSupplierName(Me.TxtPrimarySupp.Text)
                Me.cboQuickJournal.SelectedIndex = -1
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With

            'salin data payment
            If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                Me.Obj_AOPBalance.IsiDGVBalance(Me.dgvPurchaseBalance, Me.TBL_AOP_Balance, Me, Me.ObjCD, _
                Me.TxtPrimarySupp.Text, Me.TxtPrimaryJournal.Text, Me.DefaultPrecision)
                Me.TxtPurchasePayment.Text = Me.ObjCD.DigitGrouping(Me.GetTotalPaymentPurchase)
            ElseIf Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchaseReturnReceivement Then
                Me.Obj_AOPR_Balance.IsiDGVBalance(Me.dgvPurchaseReturnBalance, Me.TBL_AOPR_Balance, Me, _
                Me.ObjCD, Me.TxtPrimarySupp.Text, Me.TxtPrimaryJournal.Text, Me.DefaultPrecision)
                Me.TxtPurchaseReturnReceivement.Text = Me.ObjCD.DigitGrouping(Me.GetTotalReceivementReturn)
            End If
            Me.SetMulticurrency()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim Total1, Total2 As Double
        Dim PrimaryAkunHutang As Double = -1
        Dim NilaiPayment As Decimal = 0
        Dim NilaiPayment_AkunHutang As Decimal = 0

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

            ' Cek debet kredit harus balance
            If CDec(Me.TxtUnbalance.Text) <> 0 Then
                Me.HitungUnbalance()
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
                If Me.ObjQA.ApakahSudahAda(Me.TxtRefNo.Text) = True Then
                    MsgBox("Duplication in referance number =" + Me.TxtRefNo.Text + ", please change your number", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtRefNo.Focus()
                    Return False
                End If
            Else
                Dim TipeCodeLama As String = Me.ObjQA.GetRefCode(Me.TxtPrimaryJournal.Text)
                If TipeCodeLama <> Me.TxtRefNo.Text Then
                    If Me.ObjQA.ApakahSudahAda(Me.TxtRefNo.Text) = True Then
                        MsgBox("Duplication in referance number =" + Me.TxtRefNo.Text + ", please change your number", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtRefNo.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.ObjQA.GetTimestamp(Me.TxtPrimaryJournal.Text)
                If Me.ObjQA.Timestamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjQA.CariDetailPurchaseJournal(Me.ObjQA.Primary, Me.dgvItem, Me, Me.TipeJournal)
                        Me.SalinkeField()
                    End If
                    Return False
                End If
            End If

            'receivement/payment tidak boleh nol
            If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchaseReturnJournal Then
                If Me.GetTotalReceivementReturn = 0 Then
                    MsgBox("Please specify receivement data", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            ElseIf Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                If Me.GetTotalPaymentPurchase = 0 Then
                    MsgBox("Please specify payment data", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            'biasanya total debet dan total kredit harus sama dengan total receivement
            If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                Total1 = Me.GetTotalDebet
                Total2 = Me.GetTotalPaymentPurchase
                If Total1 < Total2 Then
                    MsgBox("Total receivement is larger than total debet.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
                If Total1 <> Total2 Then
                    If MsgBox("Total receivement is not equal to total debet. Continue?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        Return False
                    End If
                End If
            ElseIf Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchaseReturnReceivement Then
                Total1 = Me.GetTotalDebet
                Total2 = Me.GetTotalReceivementReturn
                If Total1 < Total2 Then
                    MsgBox("Total payment is larger than total debet.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
                If Total1 <> Total2 Then
                    If MsgBox("Total payment is not equal to total debet. Continue?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        Return False
                    End If
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

            'hitung payment
            Me.TxtPurchasePayment.Text = Me.ObjCD.DigitGrouping(Me.GetTotalPaymentPurchase)
            'nilai payment harus sama dengan nilai debet pada akun hutang
            If Me.dgvItem.RowCount > 1 Then
                If IsNumeric(Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.PurchaseJurnal, ClsTransaksiSetting.EnumNamaParameter.Akun_AP)) = True Then
                    PrimaryAkunHutang = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.PurchaseJurnal, ClsTransaksiSetting.EnumNamaParameter.Akun_AP)

                    If PrimaryAkunHutang <> -1 Then
                        If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                            NilaiPayment = CDec(Me.TxtPurchasePayment.Text)
                        Else
                            NilaiPayment = CDec(Me.TxtPurchaseReturnReceivement.Text)
                        End If

                        For Each row In Me.dgvItem.Rows
                            If IsNothing(row.Cells(Me.ColPrimaryAkun.Index).Value) = False Then
                                If CDbl(row.Cells(Me.ColPrimaryAkun.Index).Value) = PrimaryAkunHutang Then
                                    If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                                        NilaiPayment_AkunHutang = CDec(row.Cells(Me.ColDebet.Index).Value) - CDec(row.Cells(Me.ColKredit.Index).Value)
                                    Else
                                        NilaiPayment_AkunHutang = CDec(row.Cells(Me.ColKredit.Index).Value) - CDec(row.Cells(Me.ColDebet.Index).Value)
                                    End If
                                End If
                            End If
                        Next
                        If NilaiPayment <> NilaiPayment_AkunHutang Then
                            If NilaiPayment > 0 Then
                                If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                                    MsgBox("The debet value for '" + Me.ObjAkun.GetDescription_ByPrimaryAcc(PrimaryAkunHutang) + "' account must be " + Me.ObjCD.DigitGrouping(NilaiPayment), _
                                    MsgBoxStyle.Information, "Information")
                                Else
                                    MsgBox("The credit value for '" + Me.ObjAkun.GetDescription_ByPrimaryAcc(PrimaryAkunHutang) + "' account must be " + Me.ObjCD.DigitGrouping(NilaiPayment), _
                                    MsgBoxStyle.Information, "Information")
                                End If
                            End If
                            Return False
                        End If
                    End If
                End If
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

            Me.IsikanBalance()
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub IsikanBalance()
        Try
            If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                Me.Obj_AOPBalance.IsiDGVBalance(Me.dgvPurchaseBalance, Me.TBL_AOP_Balance, Me, Me.ObjCD, Me.TxtPrimarySupp.Text, Me.TxtPrimaryJournal.Text, Me.DefaultPrecision)
            ElseIf Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchaseReturnReceivement Then
                Me.Obj_AOPR_Balance.IsiDGVBalance(Me.dgvPurchaseReturnBalance, Me.TBL_AOPR_Balance, Me, Me.ObjCD, Me.TxtPrimarySupp.Text, _
                Me.TxtPrimaryJournal.Text, Me.DefaultPrecision)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.dgvItem.AllowUserToAddRows = True
            Me.Baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.TxtRefNo.Focus()
            Me.ObjQA.ResetObject(Me.dgvItem)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            Me.dgvItem.AllowUserToAddRows = False
            If Me.TxtRefNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtRefNo.Text <> "" Then
                If Me.ObjQA.ApakahBolehEdit(Me.ObjQA.Primary, UserCOde, True, True) = True Then
                    Me.Baru = False
                    Me.ObjQA.CariDetailPurchaseJournal(Me.ObjQA.Primary, Me.dgvItem, Me, TipeJournal)
                    Me.SalinkeField()
                    Me.AktifkanDetail()
                    Me.dgvItem.AllowUserToAddRows = True
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
            If Me.TxtRefNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtRefNo.Text <> "" Then
                If Me.ObjQA.ApakahBolehHapus(Me.TxtPrimaryJournal.Text, True, UserCOde, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjQA.Hapus_PurchaseJurnal(Me.TxtPrimaryJournal.Text, True, UserCode, Me.TipeJournal) = True Then
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
            Me.SalinkeObjek()
            If Baru = True Then
                If ObjQA.ApakahDataBenar(Baru) = True Then
                    If Me.ApakahDataBenar() = True Then
                        'konfirmasi disimpan atau tidak
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")

                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                                hasil_save = Me.ObjQA.SimpanBaruPurchaseJournal(Me.dgvItem, Me.UserCode, Me, TipeJournal, Me.dgvPurchaseBalance)
                            Else
                                hasil_save = Me.ObjQA.SimpanBaruPurchaseJournal(Me.dgvItem, Me.UserCode, Me, TipeJournal, Me.dgvPurchaseReturnBalance)
                            End If
                            If hasil_save = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinkeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                                Me.AddReminder()
                            End If
                            Me.dgvItem.AllowUserToAddRows = False
                        ElseIf hasil = MsgBoxResult.No Then
                            Me.dgvItem.AllowUserToAddRows = True
                        End If
                    Else
                        Me.dgvItem.AllowUserToAddRows = True
                    End If
                Else
                    Me.dgvItem.AllowUserToAddRows = True
                End If
            Else
                'konfirmasi disimpan atau tidak
                If ObjQA.ApakahBolehEdit(Me.TxtPrimaryJournal.Text, UserCode, True, True) = True Then
                    If Me.ApakahDataBenar = True Then
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                                hasil_save = Me.ObjQA.SimpanPerubahanPurchaseJournal(Me.dgvItem, Me.UserCode, Me, TipeJournal, Me.dgvPurchaseBalance)
                            Else
                                hasil_save = Me.ObjQA.SimpanPerubahanPurchaseJournal(Me.dgvItem, Me.UserCode, Me, TipeJournal, Me.dgvPurchaseReturnBalance)
                            End If
                            If hasil_save = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinkeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                                Me.AddReminder()
                            End If
                            Me.dgvItem.AllowUserToAddRows = False
                        ElseIf hasil = MsgBoxResult.No Then
                            Me.dgvItem.AllowUserToAddRows = True
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

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.CancelInput()
    End Sub

    Private Sub CancelInput()
        Dim hasil As MsgBoxResult
        Try
            Me.dgvItem.AllowUserToAddRows = False
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
                        Me.ObjQA.CariDetailPurchaseJournal(Me.TxtPrimaryJournal.Text, Me.dgvItem, Me, TipeJournal)
                        Me.SalinkeField()
                    End If
                Else
                    Me.dgvItem.AllowUserToAddRows = True
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
            If Me.TxtRemarks.Enabled = True Then
                Me.BringToFront()
                If Me.ApakahDataBenar = True Then
                    Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ApakahDataBenar() = True Then
                            Me.SalinkeObjek()
                            If Baru = True Then
                                'simpan baru
                                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                                If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                                    hasil_res = Me.ObjQA.SimpanBaruPurchaseJournal(Me.dgvItem, Me.UserCode, Me, Me.TipeJournal, Me.dgvPurchaseBalance)
                                Else
                                    hasil_res = Me.ObjQA.SimpanBaruPurchaseJournal(Me.dgvItem, Me.UserCode, Me, Me.TipeJournal, Me.dgvPurchaseReturnBalance)
                                End If
                                If hasil = True Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                    Me.SalinkeField()
                                    Me.NonAktifkanDetail()
                                    Me.btnFind.Focus()
                                Else
                                    e.Cancel = True
                                End If
                            Else
                                'simpan perubahan
                                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                                If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                                    hasil_res = Me.ObjQA.SimpanPerubahanPurchaseJournal(Me.dgvItem, Me.UserCode, Me, Me.TipeJournal, Me.dgvPurchaseBalance)
                                Else
                                    hasil_res = Me.ObjQA.SimpanPerubahanPurchaseJournal(Me.dgvItem, Me.UserCode, Me, Me.TipeJournal, Me.dgvPurchaseReturnBalance)
                                End If
                                If hasil = True Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                    Me.SalinkeField()
                                    Me.NonAktifkanDetail()
                                    Me.btnEdit.Focus()
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
            'yang boleh di edit adalah
            If e.ColumnIndex = Me.ColPayment.Index Then
                Me.Buffer = Me.dgvPurchaseBalance.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvPurchaseBalance_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvPurchaseBalance.CellContentClick
        Try
            If Me.TxtRemarks.Enabled = True Then
                If e.ColumnIndex = Me.Col_Purchase_Pay.Index Then
                    Me.SetFullPayment(e.RowIndex)
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
                    Me.InsertQuickJurnal(False)
                End If

                For Each rowDGV In Me.dgvPurchaseBalance.Rows
                    If rowDGV.Cells(Me.ColTipeDocumentAcuan_Purchase.Index).Value = vTipeDokumenBalancePurchase And _
                    vPrimaryDokumentAcuan = rowDGV.Cells(Me.Col_PrimaryDocumentAcuanPurchase.Index).Value Then
                        Me.SetFullPayment(rowDGV.Index)
                    End If
                Next

                If Me.Size.Height < Me.MinimumSize.Height Or Me.Size.Width < Me.MinimumSize.Width Then
                    Me.Size = Me.MinimumSize
                End If
                Me.BringToFront()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetFullPayment(ByVal vRowIndex As Integer)
        Dim row As DataGridViewRow

        Try
            'hitung ulang pembyaran sampai lunas
            row = Me.dgvPurchaseBalance.Rows(vRowIndex)
            row.Cells(Me.ColPayment.Index).Value = _
            Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColTotalPurchase_VAT_Value.Index).Value) - _
            CDec(row.Cells(Me.ColPaid.Index).Value), Me.DefaultPrecision)
            row.Cells(Me.ColBalancePurchase.Index).Value = 0
            'update data sales receivement
            Me.TxtPurchasePayment.Text = Me.ObjCD.DigitGrouping(Me.GetTotalPaymentPurchase(), Me.DefaultPrecision)
            Me.UpdateJurnal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateJurnal()
        Dim row As DataGridViewRow

        Try
            If Me.dgvItem.Rows.Count = 3 Then
                row = Me.dgvItem.Rows(0)
                If CDec(row.Cells(Me.ColDebet.Index).Value) <> 0 Then
                    row.Cells(Me.ColDebet.Index).Value = Me.TxtPurchasePayment.Text
                ElseIf CDec(row.Cells(Me.ColKredit.Index).Value) <> 0 Then
                    row.Cells(Me.ColKredit.Index).Value = Me.TxtPurchasePayment.Text
                End If
                row = Me.dgvItem.Rows(1)
                If CDec(row.Cells(Me.ColDebet.Index).Value) <> 0 Then
                    row.Cells(Me.ColDebet.Index).Value = Me.TxtPurchasePayment.Text
                ElseIf CDec(row.Cells(Me.ColKredit.Index).Value) <> 0 Then
                    row.Cells(Me.ColKredit.Index).Value = Me.TxtPurchasePayment.Text
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvPurchaseBalance_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvPurchaseBalance.CellEndEdit
        Dim Balance As Decimal = 0

        Try
            'yang boleh di edit adalah
            If e.ColumnIndex = Me.ColPayment.Index Then
                If IsNothing(Me.dgvPurchaseBalance.Item(e.ColumnIndex, e.RowIndex)) = True Then
                    'langsung di balikin ke nilai semula
                    Me.dgvPurchaseBalance.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Else
                    ''cek debet
                    If e.ColumnIndex = Me.ColPayment.Index Then
                        If Me.ObjCD.CekNonNegatif(Me.dgvPurchaseBalance.Item(e.ColumnIndex, _
                            e.RowIndex).Value) = False Then
                            MsgBox("Please enter positive number", MsgBoxStyle.Information, _
                            "Information")
                            Me.dgvPurchaseBalance.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If

                        'rounding 0
                        Me.dgvPurchaseBalance.Item(e.ColumnIndex, e.RowIndex).Value = _
                        Me.ObjCD.DigitGrouping(CDbl(Me.dgvPurchaseBalance.Item(e.ColumnIndex, e.RowIndex).Value), Me.DefaultPrecision)
                        'pastikan nilai balance tidak menjadi negatif
                        Balance = CDec(Me.dgvPurchaseBalance.Item(Me.ColPurchaseValue.Index, e.RowIndex).Value) + _
                        CDec(Me.dgvPurchaseBalance.Item(Me.ColVatPurchaseValue.Index, e.RowIndex).Value) + _
                        CDec(Me.dgvPurchaseBalance.Item(Me.ColDutyStamp.Index, e.RowIndex).Value) - _
                        CDec(Me.dgvPurchaseBalance.Item(Me.ColPaid.Index, e.RowIndex).Value)
                        If Balance - CDec(Me.dgvPurchaseBalance.Item(e.ColumnIndex, e.RowIndex).Value) < 0 Then
                            MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(Balance, Me.DefaultPrecision) + ".", MsgBoxStyle.Information, _
                            "Information")
                            Me.dgvPurchaseBalance.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        'hitung ulang nilai balance
                        Me.dgvPurchaseBalance.Item(Me.ColBalancePurchase.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Balance - _
                        CDec(Me.dgvPurchaseBalance.Item(e.ColumnIndex, e.RowIndex).Value), Me.DefaultPrecision)
                        'digit grouping
                        Me.dgvPurchaseBalance.Item(e.ColumnIndex, e.RowIndex).Value = _
                        ObjCD.DigitGrouping(Me.dgvPurchaseBalance.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefaultPrecision)
                        'update data sales receivement
                        Me.TxtPurchasePayment.Text = Me.ObjCD.DigitGrouping(Me.GetTotalPaymentPurchase(), Me.DefaultPrecision)
                        Me.UpdateJurnal()
                    End If
                End If
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

    Private Sub InsertQuickJurnal(ByVal vConfirm As Boolean)
        Dim rowDGV As DataGridViewRow
        Dim ExcRate As Decimal = 1

        Try
            If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                Me.ObjQuickJurnal.IsiDGVQuickAdd(Me.TipeJournal, Me.dgvItem, Me.GetTotalPaymentPurchase, _
                Me.cboQuickJournal.Text, Me.ColAccount.Index, Me.ColDescription.Index, Me.ColPrimaryAkun.Index, _
                Me.ColDebet.Index, Me.ColKredit.Index, Me.ColRemarks.Index, Me.ObjCD, Me.DefaultPrecision, vConfirm)
            ElseIf Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchaseReturnReceivement Then
                Me.ObjQuickJurnal.IsiDGVQuickAdd(Me.TipeJournal, Me.dgvItem, Me.GetTotalReceivementReturn, _
                Me.cboQuickJournal.Text, Me.ColAccount.Index, Me.ColDescription.Index, Me.ColPrimaryAkun.Index, _
                Me.ColDebet.Index, Me.ColKredit.Index, Me.ColRemarks.Index, Me.ObjCD, Me.DefaultPrecision, vConfirm)
            End If
            Me.HitungUnbalance()
            Me.TxtRemarks.Text = Me.ObjQuickJurnal.GetDesc_ByCode(Me.cboQuickJournal.Text) + " " + _
            Me.TxtSuppCode.Text + "-" + Me.TxtSuppName.Text

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
                rowDGV.Cells(Me.ColCurrency.Index).Value = Me.ObjCur.GetCurrencyCode(rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value)
                'isikan exchange rate
                If rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value <> 1 Then
                    ExcRate = Me.ObjTaxRate.GetTaxRate(Me.dtpDate.Value, rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value)
                    If ExcRate = 0 Then
                        ExcRate = 1
                    End If
                End If
                rowDGV.Cells(Me.ColExchangeRate.Index).Value = Me.ObjCD.DigitGrouping(ExcRate, DefPrecExchangeRate)
                'hitung original values
                Me.HitungOriginalValues(rowDGV.Index)
            Next
            'set multicurrency
            Me.SetMulticurrency()
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
            vRowDGV.Cells(Me.ColExchangeRate.Index).Value = ObjCD.DigitGrouping(Math.Round(CDec(vRowDGV.Cells(Me.ColDebet.Index).Value) / _
            CDec(vRowDGV.Cells(Me.ColDebetOriginal.Index).Value), DefPrecExchangeRate, MidpointRounding.AwayFromZero))
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
End Class