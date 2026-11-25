Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiJournal_Sales2
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightSeaGreen
    Public TipeJournal As ClsTransaksiJournal.EnumTipeJournal = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal
    Public ObjJurnal As New ClsTransaksiJournal

    '' Dim mBaru As Boolean = False
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjInvoiceBalance As New ClsTransaksiSalesBalance
    Dim TBLSalesBalance As New DseTransaksiSalesBalance.balance_invoice_viewDataTable
    Dim ObjSRBalance As New ClsTransaksiSalesReturnBalance
    Dim TBLSRBalance As New DseTransaksiSalesReturnBalance.balance_sr_viewDataTable
    Dim ObjLicense As New ClsLicense
    Dim ObjClosing As New ClsTransaksiClosing
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim PrimaryBaseCurrency As Double = 1
    Dim TBLBeban As New DseMasterAkun.tblmaster_akun3DataTable
    Dim TBLAkun As New DseMasterAkun.tblmaster_akun4DataTable

    Dim Buffer As String = ""
    Dim DefaultPrecision As Integer = 0
    Dim Baru As Boolean
    Dim ObjAkun As New ClsMasterAkun
    Dim ObjCust As New clsMasterCustomer
    Dim ObjJurnalItem As New ClsTransaksiJournal_Item
    Dim ObjQuickJurnal As New ClsTransaksiJournalQuickAdd
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjTaxRate As New ClsTransaksiSalesInvoiceTaxRate
    Dim ObjInv As New ClsTransaksiSalesInvoice
    Dim ObjSR As New ClsTransaksiSalesReturn

    Public Sub Pay(ByVal vPrimaryCust As Double, ByVal vTipeDokumentAcuan As AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance, _
    ByVal vPrimaryDokumentAcuan As Double)
        Dim rowDGV As DataGridViewRow

        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.CancelInput()
            End If

            If Me.TxtRemarks.Enabled = False Then
                Me.btnNew.PerformClick()
                Me.TxtPrimaryCust.Text = vPrimaryCust

                Me.TxtCustCode.Text = Me.ObjCust.GetCustomerCode(Me.TxtPrimaryCust.Text)
                Me.TxtCustName.Text = Me.ObjCust.GetCustName(Me.TxtPrimaryCust.Text)

                Me.IsikanBalance()

                'lakukan pembayaran
                If Me.cboQuickJournal.Items.Count >= 2 Then
                    Me.cboQuickJournal.SelectedIndex = 1
                    'Me.InsertQuickJurnal(False)
                End If

                For Each rowDGV In Me.dgvSalesBalance.Rows
                    If rowDGV.Cells(Me.ColTipeDocumentAcuan_Sales.Index).Value = vTipeDokumentAcuan And _
                    vPrimaryDokumentAcuan = rowDGV.Cells(Me.ColPrimaryDocumenAcuan_Sales.Index).Value Then
                        If Me.ApakahBolehMasukkanPayment(rowDGV.Index) = True Then
                            Me.SetFullPayment(rowDGV.Index)
                            If Me.ApakahAdaCurrency(Me.dgvSalesBalance.Item(Me.ColPrimaryCurrencyBalance.Index, rowDGV.Index).Value) = False Then
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

                Me.TxtRemarks.Text = "AR Payment: " + Me.TxtCustCode.Text + " - " + Me.TxtCustName.Text
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSalesCycle_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim NullValue As String = ""

        Try
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvSalesBalance.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvSalesReturnBalance.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
                Me.SplitContainer3.Panel2Collapsed = True
                Me.Text = "Sales Receivement Journal"
            Else
                Me.SplitContainer3.Panel1Collapsed = True
                Me.Text = "Sales Return Payment Journal"
            End If

            ''Me.AktifkanDetail()
            Me.NonAktifkanDetail()

            Me.dgvItem.AllowUserToAddRows = False
            Me.ObjJurnal = New ClsTransaksiJournal
            Me.Baru = True
            Me.dgvItem.Rows.Clear()
            Me.KosongkanField()

            'set null value untuk dgv
            Me.DefaultPrecision = Me.ObjCur.GetDecimal_DefaultCurrency
            NullValue = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            Me.ColDebet.DefaultCellStyle.NullValue = NullValue
            Me.ColKredit.DefaultCellStyle.NullValue = NullValue
            Me.TxtUnbalance.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            Me.TxtSalesReceivement.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            Me.TxtSalesReturnPayment.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
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
        Dim ij As New FrmTransaksiJournal_Cetak
        ij.TipeJournal = Me.TipeJournal
        Try
            ij.BackColorAll = Me.BackColorAll
            ij.ShowDialog()
            Me.dgvItem.AllowUserToAddRows = False
            If Not ij.PrimaryPP = -1 Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                If Me.ObjJurnal.CariDetailSalesJournal(ij.PrimaryPP, Me.dgvItem, Me, ij.TipeJournal) = True Then
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
            Unbalance = Math.Round(vdebet - vkredit, Me.DefaultPrecision, MidpointRounding.AwayFromZero)
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
            'Me.ObjInvoiceBalance.IsiDGVBalance(Me.dgvSalesBalance, Me.TBLSalesBalance, Me, Me.ObjCD)
            'ElseIf Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.SalesReturnPaymentJournal Then
            'Me.ObjSRBalance.IsiDGVBalance(Me.dgvSalesReturnBalance, Me.TBLSRBalance, Me, Me.ObjCD)
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

            Me.dgvSalesBalance.AllowUserToDeleteRows = False
            Me.dgvSalesBalance.ReadOnly = False
            Me.dgvSalesReturnBalance.AllowUserToDeleteRows = False
            Me.dgvSalesReturnBalance.ReadOnly = False

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

            Me.dgvSalesBalance.AllowUserToDeleteRows = False
            Me.dgvSalesBalance.ReadOnly = True
            Me.dgvSalesReturnBalance.AllowUserToDeleteRows = False
            Me.dgvSalesReturnBalance.ReadOnly = True

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
            Me.TxtPrimaryCust.Text = "-1"
            Me.TxtPrimaryAccountDifference.Text = "-1"
            Me.TxtPrimaryAccountPayTo.Text = "-1"
            Me.TxtUnbalance.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            Me.TxtSalesReceivement.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            Me.TxtSalesReturnPayment.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping("1", Core_Module.DefPrecExchangeRate)
            Me.TxtPaid.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision)

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "

            Me.dgvItem.Rows.Clear()
            Me.dgvSalesBalance.Rows.Clear()
            Me.dgvSalesReturnBalance.Rows.Clear()

            Me.cboPayToAcc.DataSource = Nothing
            Me.cboPayToAcc.Items.Clear()
            ''Me.ObjCD.SetNumber(Me.dgvItem)
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
                If Me.ObjJurnal.CariDetailSalesJournal(vPrimaryJurnal, Me.dgvItem, Me, Me.TipeJournal) = True Then
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

    Private Sub SalinkeObjek()
        Try
            With ObjJurnal
                .RefNumber = Me.TxtRefNo.Text
                .Description = Me.TxtRemarks.Text
                .Tanggal = Me.dtpDate.Value
                .PrimaryPihakAcuan = Me.TxtPrimaryCust.Text

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
                Me.TxtPrimaryCust.Text = .PrimaryPihakAcuan
                Me.TxtCustCode.Text = ObjCust.GetCustomerCode(Me.TxtPrimaryCust.Text)
                Me.TxtCustName.Text = ObjCust.GetCustName(Me.TxtPrimaryCust.Text)
                Me.cboQuickJournal.SelectedIndex = -1
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString

                'salin data payment
                If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
                    Me.ObjInvoiceBalance.IsiDGVBalance(Me.dgvSalesBalance, _
                    Me.TBLSalesBalance, Me, Me.ObjCD, Me.TxtPrimaryCust.Text, Me.TxtPrimaryJournal.Text, Me.DefaultPrecision, Me.TBLSRBalance)
                    Me.TxtSalesReceivement.Text = Me.ObjCD.DigitGrouping(Me.GetTotalReceivementSales)
                ElseIf Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.SalesReturnPaymentJournal Then
                    Me.ObjSRBalance.IsiDGVBalance(Me.dgvSalesReturnBalance, _
                    Me.TBLSRBalance, Me, Me.ObjCD, Me.TxtPrimaryCust.Text, Me.TxtPrimaryJournal.Text, Me.DefaultPrecision)
                    Me.TxtSalesReturnPayment.Text = Me.ObjCD.DigitGrouping(Me.GetTotalPaymentSalesReturn)
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

                For Each row In Me.TBLAkun.Rows
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

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim PrimaryAkunPiutang As Double = -1
        Dim NilaiPayment As Decimal = 0
        Dim NilaiPayment_AkunPiutang As Decimal = 0
        Dim rowDGV As DataGridViewRow
        Dim TotalReceivement As Decimal = 0
        Dim TipeAkun As ClsMasterAkun.EnumTipeAkun
        Dim MaxDate As Date = Today
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
            For Each rowDGV In Me.dgvSalesBalance.Rows
                TotalReceivement += CDec(rowDGV.Cells(Me.ColReceivementOriginal.Index).Value)
            Next
            If TotalReceivement <> 0 And CDec(Me.TxtPaid.Text) = 0 Then
                If MsgBox("No payment amount specified, continue save?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Me.TxtPaid.Focus()
                    Return False
                End If
            End If
            'receivement/payment tidak boleh nol
            If Me.ApakahAdaPayment = False Then
                MsgBox("Please select at least 1 item to pay.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            ' Cek debet kredit harus balance
            Me.HitungUnbalance()
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
                        Me.ObjJurnal.CariDetailSalesJournal(Me.ObjJurnal.Primary, Me.dgvItem, Me, Me.TipeJournal)
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
            If CDec(Me.TxtPaid.Text) > 0 And Me.GetTotalReceivementSales < 0 Then
                MsgBox("Please specify a negatif value payment.", MsgBoxStyle.Information, "Information")
                Me.TxtPaid.Focus()
                Return False
            ElseIf CDec(Me.TxtPaid.Text) < 0 And Me.GetTotalReceivementSales > 0 Then
                MsgBox("Please specify a positive value payment.", MsgBoxStyle.Information, "Information")
                Me.TxtPaid.Focus()
                Return False
            End If

            'pastikan akun difference adalah expense, other expense atau cogs
            If CDec(Me.TxtDifference.Text) <> 0 Then
                TipeAkun = Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryAccountDifference.Text)
                If TipeAkun <> ClsMasterAkun.EnumTipeAkun.Cost_of_Goods_Sales And TipeAkun <> ClsMasterAkun.EnumTipeAkun.Expense And _
                TipeAkun <> ClsMasterAkun.EnumTipeAkun.Other_Expense Then
                    MsgBox("Please specify expense type account for difference account.", MsgBoxStyle.Information, "Information")
                    Me.TxtDifferenceAccountCode.Focus()
                    Return False
                End If
            End If

            'pembayaran paling cepat adalah sesuai tanggal invoice atau sr
            MaxDate = Date.MinValue
            For Each rowDGV In Me.dgvSalesBalance.Rows
                If CDec(rowDGV.Cells(Me.ColReceivementOriginal.Index).Value) <> 0 Then
                    If CDate(rowDGV.Cells(Me.ColInvoiceDate.Index).Value) > MaxDate Then
                        MaxDate = CDate(rowDGV.Cells(Me.ColInvoiceDate.Index).Value)
                    End If
                End If
            Next
            If MaxDate <> Date.MinValue Then
                If MaxDate > Me.dtpDate.Value.Date And Format(MaxDate, "MMM yyyy") <> Format(Me.dtpDate.Value.Date, "MMM yyyy") Then
                    For Each rowDGV In Me.dgvSalesBalance.Rows
                        If MaxDate = CDate(rowDGV.Cells(Me.ColInvoiceDate.Index).Value) And _
                        CDec(rowDGV.Cells(Me.ColReceivementOriginal.Index).Value) <> 0 Then

                            If DummyDoc = "" Then
                                DummyDoc = rowDGV.Cells(Me.ColInvoiceNo.Index).Value
                            Else
                                DummyDoc = DummyDoc + "," + rowDGV.Cells(Me.ColInvoiceNo.Index).Value
                            End If
                        End If
                    Next

                    MsgBox("Minimum payment date is " + Format(MaxDate, "dd MMM yyyy") + " due to document: " + _
                    DummyDoc + ".", MsgBoxStyle.Information, "Information")
                    Return False
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

    Private Function GetTotalReceivementOriginalSales() As Decimal
        Dim rowdgv As DataGridViewRow
        Dim retval As Decimal = 0

        Try
            For Each rowdgv In Me.dgvSalesBalance.Rows
                retval += CDec(rowdgv.Cells(Me.ColReceivementOriginal.Index).Value)
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetTotalReceivementSales() As Decimal
        Dim rowdgv As DataGridViewRow
        Dim retval As Decimal = 0

        Try
            For Each rowdgv In Me.dgvSalesBalance.Rows
                retval += CDec(rowdgv.Cells(Me.ColReceivement.Index).Value)
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetTotalPaymentSalesReturn() As Decimal
        Dim rowdgv As DataGridViewRow
        Dim retval As Decimal = 0

        Try
            For Each rowdgv In Me.dgvSalesReturnBalance.Rows
                retval += CDec(rowdgv.Cells(Me.ColReturnPayment.Index).Value)
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

    Private Sub TxtCustCode_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustCode.Enter
        Try
            Me.Buffer = Me.TxtPrimaryCust.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPihakAcuan_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustCode.GotFocus
        Try
            If Me.TxtCustCode.Enabled = True Then
                If Me.TxtCustCode.ReadOnly = False Then
                    Me.TxtCustCode.BackColor = Me.BackColorAll
                Else
                    TxtCustCode.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPihakAcuan_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustCode.LostFocus
        Try
            If Me.TxtCustCode.Enabled = True Then
                If Me.TxtCustCode.ReadOnly = False Then
                    Me.TxtCustCode.BackColor = Me.TxtPrimaryJournal.BackColor
                Else
                    TxtCustCode.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''Find Code No
    Private Sub TxtNoJadwalProduksi_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles TxtCustCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtCustCode.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjCust.ApakahSudahAda(Me.TxtCustCode.Text) = True Then
                        Me.TxtPrimaryCust.Text = ObjCust.GetPrimary(Me.TxtCustCode.Text)
                        Me.TxtCustCode.Text = ObjCust.GetCustomerCode(Me.TxtPrimaryCust.Text)
                        Me.TxtCustName.Text = ObjCust.GetCustName(Me.TxtPrimaryCust.Text)
                        Me.TxtRemarks.Focus()

                        'isikan data balance
                        Me.IsikanBalance()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.BtnFindCust.PerformClick()
                    End If
                Else
                    Me.BtnFindCust.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNoJadwalProduksi_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtCustCode.Leave
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.TxtCustCode.Text <> "" Then
                If ObjCust.ApakahSudahAda(Me.TxtCustCode.Text) = True Then
                    Me.TxtPrimaryCust.Text = ObjCust.GetPrimary(Me.TxtCustCode.Text)
                    Me.TxtCustCode.Text = ObjCust.GetCustomerCode(Me.TxtPrimaryCust.Text)
                    Me.TxtCustName.Text = ObjCust.GetCustName(Me.TxtPrimaryCust.Text)
                Else
                    If Me.TxtCustCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtCustCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.TxtPrimaryCust.Text = i.PrimaryCustomer
                        Me.TxtCustCode.Text = ObjCust.GetCustomerCode(Me.TxtPrimaryCust.Text)
                        Me.TxtCustName.Text = ObjCust.GetCustName(Me.TxtPrimaryCust.Text)
                    Else
                        Me.TxtCustCode.Clear()
                        Me.TxtCustName.Clear()
                        Me.TxtPrimaryCust.Text = "-1"
                        Me.TxtCustCode.Focus()
                    End If
                End If
            Else
                Me.TxtCustCode.Clear()
                Me.TxtCustName.Clear()
                Me.TxtPrimaryCust.Text = "-1"
            End If

            'isikan data balance, jika customer berubah atau jika masih kosong semua
            If Me.TxtPrimaryCust.Text <> Me.Buffer Or Me.GetTotalReceivementOriginalSales = 0 Then
                Me.IsikanBalance()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub BtnFindSchedule_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles BtnFindCust.Click
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCustCode.Text <> "" Then
                i.txtCari.Text = Me.TxtCustCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryCustomer > 0 Then
                Me.TxtPrimaryCust.Text = i.PrimaryCustomer
                Me.TxtCustCode.Text = ObjCust.GetCustomerCode(Me.TxtPrimaryCust.Text)
                Me.TxtCustName.Text = ObjCust.GetCustName(Me.TxtPrimaryCust.Text)
                Me.TxtRemarks.Focus()
            Else
                Me.TxtCustCode.Clear()
                Me.TxtCustName.Clear()
                Me.TxtPrimaryCust.Text = "-1"
                Me.TxtCustCode.Focus()
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
            'isikan data balance
            If Me.TxtPrimaryCust.Text = "-1" Then
                Me.dgvSalesBalance.Rows.Clear()
            Else
                If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
                    Me.ObjInvoiceBalance.IsiDGVBalance(Me.dgvSalesBalance, Me.TBLSalesBalance, Me, Me.ObjCD, Me.TxtPrimaryCust.Text, Me.TxtPrimaryJournal.Text, Me.DefaultPrecision, Me.TBLSRBalance)
                ElseIf Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.SalesReturnPaymentJournal Then
                    Me.ObjSRBalance.IsiDGVBalance(Me.dgvSalesReturnBalance, Me.TBLSRBalance, Me, Me.ObjCD, Me.TxtPrimaryCust.Text, Me.TxtPrimaryJournal.Text, Me.DefaultPrecision)
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
                If Me.ObjJurnal.ApakahBolehEdit(Me.ObjJurnal.Primary, UserCode, True, True) = True Then
                    Me.Baru = False
                    Me.ObjJurnal.CariDetailSalesJournal(Me.ObjJurnal.Primary, Me.dgvItem, Me, TipeJournal)
                    Me.SalinkeField()
                    Me.AktifkanDetail()
                    Me.dgvItem.AllowUserToAddRows = False
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
                If Me.ObjJurnal.ApakahBolehHapus(Me.TxtPrimaryJournal.Text, True, UserCode, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjJurnal.Hapus_SalesJurnal(Me.TxtPrimaryJournal.Text, True, UserCode, Me.TipeJournal) = True Then
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

    Private Sub AddReminder()
        Dim ij As New FrmTransaksiReminder_New

        Try
            Exit Sub
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

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Dim hasil_save As Boolean = False

        Try
            Me.dgvItem.AllowUserToAddRows = False
            Me.SalinkeObjek()
            If Baru = True Then
                If ObjJurnal.ApakahDataBenar1(Baru, Me.dgvItem, Me.ColPrimaryCurrency.Index, Me.dgvSalesBalance, Me.ColPrimaryCurrencyBalance.Index, _
                Me.ColReceivementOriginal.Index) = True Then
                    If Me.ApakahDataBenar() = True Then
                        'konfirmasi disimpan atau tidak
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")

                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
                                hasil_save = Me.ObjJurnal.SimpanBaruSalesJournal(Me.dgvItem, Me.UserCode, Me, TipeJournal, Me.dgvSalesBalance)
                            Else
                                hasil_save = Me.ObjJurnal.SimpanBaruSalesJournal(Me.dgvItem, Me.UserCode, Me, TipeJournal, Me.dgvSalesReturnBalance)
                            End If
                            If hasil_save = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.NonAktifkanDetail()
                                Me.SalinkeField()
                                Me.btnFind.Focus()
                                Me.AddReminder()
                            End If
                            Me.dgvItem.AllowUserToAddRows = False
                        ElseIf hasil = MsgBoxResult.No Then
                            Me.dgvItem.AllowUserToAddRows = False
                        End If
                    Else
                        Me.dgvItem.AllowUserToAddRows = False
                    End If
                Else
                    Me.dgvItem.AllowUserToAddRows = False
                End If
            Else
                'konfirmasi disimpan atau tidak
                If ObjJurnal.ApakahBolehEdit(Me.TxtPrimaryJournal.Text, UserCode, True, True) = True Then
                    If Me.ApakahDataBenar = True Then
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
                                hasil_save = Me.ObjJurnal.SimpanPerubahanSalesJournal(Me.dgvItem, Me.UserCode, Me, TipeJournal, Me.dgvSalesBalance)
                            Else
                                hasil_save = Me.ObjJurnal.SimpanPerubahanSalesJournal(Me.dgvItem, Me.UserCode, Me, TipeJournal, Me.dgvSalesReturnBalance)
                            End If
                            If hasil_save = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.NonAktifkanDetail()
                                Me.SalinkeField()
                                Me.btnEdit.Focus()
                                Me.AddReminder()
                            End If
                            Me.dgvItem.AllowUserToAddRows = False
                        ElseIf hasil = MsgBoxResult.No Then
                            Me.dgvItem.AllowUserToAddRows = False
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
        Try
            Me.CancelInput()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
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
                    Me.NonAktifkanDetail()
                    ''Me.KosongkanItem()
                    If Baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.NonAktifkanDetail()
                        Me.ObjJurnal.CariDetailSalesJournal(Me.TxtPrimaryJournal.Text, Me.dgvItem, Me, TipeJournal)
                        Me.SalinkeField()
                    End If
                Else
                    Me.dgvItem.AllowUserToAddRows = False
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
                                    hasil_res = Me.ObjJurnal.SimpanBaruSalesJournal(Me.dgvItem, Me.UserCode, Me, TipeJournal, Me.dgvSalesBalance)
                                ElseIf Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchaseReturnReceivement Then
                                    hasil_res = Me.ObjJurnal.SimpanBaruSalesJournal(Me.dgvItem, Me.UserCode, Me, TipeJournal, Me.dgvSalesReturnBalance)
                                End If
                                If hasil_res = True Then
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
                                    hasil_res = Me.ObjJurnal.SimpanPerubahanSalesJournal(Me.dgvItem, Me.UserCode, Me, TipeJournal, Me.dgvSalesBalance)
                                ElseIf Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchaseReturnReceivement Then
                                    hasil_res = Me.ObjJurnal.SimpanPerubahanSalesJournal(Me.dgvItem, Me.UserCode, Me, TipeJournal, Me.dgvSalesReturnBalance)
                                End If
                                If hasil_res = True Then
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

    Private Sub TxtSalesReceivement_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesReceivement.Enter
        Try
            If Me.TxtSalesReceivement.Enabled = True And Me.TxtRemarks.Enabled = True Then
                If CDec(Me.TxtSalesReceivement.Text) = 0 Then
                    Me.TxtSalesReceivement.Text = 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesReceivement_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesReceivement.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtSalesReceivement.ReadOnly = True Then
                    Me.TxtSalesReceivement.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtSalesReceivement.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesReceivement_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtSalesReceivement.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtSalesReceivement, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesReceivement_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtSalesReceivement.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtSalesReceivement, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesReceivement_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesReceivement.Leave
        Try
            If Me.TxtSalesReceivement.Enabled = True And Me.TxtRemarks.Enabled = True Then
                Me.TxtSalesReceivement.Text = Me.ObjCD.DigitGrouping(Me.TxtSalesReceivement.Text, Me.DefaultPrecision)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesReceivement_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesReceivement.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtSalesReceivement.ReadOnly = True Then
                    Me.TxtSalesReceivement.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtSalesReceivement.BackColor = Me.TxtPrimaryJournal.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesReturnPayment_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesReturnPayment.Enter
        Try
            If Me.TxtSalesReturnPayment.Enabled = True And Me.TxtRemarks.Enabled = True Then
                If CDec(Me.TxtSalesReturnPayment.Text) = 0 Then
                    Me.TxtSalesReturnPayment.Text = 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesReturnPayment_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesReturnPayment.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtSalesReturnPayment.ReadOnly = True Then
                    Me.TxtSalesReturnPayment.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtSalesReturnPayment.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesReturnPayment_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtSalesReturnPayment.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtSalesReturnPayment, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesReturnPayment_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtSalesReturnPayment.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtSalesReturnPayment, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesReturnPayment_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesReturnPayment.Leave
        Try
            Me.TxtSalesReturnPayment.Text = Me.ObjCD.DigitGrouping(Me.TxtSalesReturnPayment.Text, Me.DefaultPrecision)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesReturnPayment_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesReturnPayment.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtSalesReturnPayment.ReadOnly = True Then
                    Me.TxtSalesReturnPayment.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtSalesReturnPayment.BackColor = Me.TxtPrimaryJournal.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDistribusiReceivement_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDistribusiReceivement.Click
        Dim Angka As Decimal = 0
        Dim rowDGV As DataGridViewRow

        Try
            Angka = Me.TxtSalesReceivement.Text
            If Angka > 0 Then
                For i As Double = 0 To Me.dgvSalesBalance.RowCount - 1
                    rowDGV = Me.dgvSalesBalance.Rows(i)
                    rowDGV.Cells(Me.ColReceivement.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)

                    rowDGV.Cells(Me.ColBalanceSales.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDGV.Cells(Me.ColSalesValue.Index).Value) + _
                    CDec(rowDGV.Cells(Me.ColVatSalesValue.Index).Value) + CDec(rowDGV.Cells(Me.ColDutystamp.Index).Value) - _
                    CDec(rowDGV.Cells(Me.ColReceived.Index).Value))

                    If Angka <> 0 Then
                        If CDec(rowDGV.Cells(Me.ColBalanceSales.Index).Value) < Angka Then
                            'yang dipindahkan adalah sejumlah balance sales
                            rowDGV.Cells(Me.ColReceivement.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColBalanceSales.Index).Value, Me.DefaultPrecision)
                            Angka -= CDec(rowDGV.Cells(Me.ColBalanceSales.Index).Value)
                            rowDGV.Cells(Me.ColBalanceSales.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                        Else
                            rowDGV.Cells(Me.ColReceivement.Index).Value = Me.ObjCD.DigitGrouping(Angka, Me.DefaultPrecision)
                            rowDGV.Cells(Me.ColBalanceSales.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDGV.Cells(Me.ColSalesValue.Index).Value) + _
                            CDec(rowDGV.Cells(Me.ColVatSalesValue.Index).Value) + CDec(rowDGV.Cells(Me.ColDutystamp.Index).Value) - _
                            CDec(rowDGV.Cells(Me.ColReceived.Index).Value) - Angka, Me.DefaultPrecision)
                            Angka = 0
                        End If
                    End If
                Next
            Else
                'jika angka receivement yang dimasukkan minus, maka hanya ubah untuk data sales return saja
                For i As Double = 0 To Me.dgvSalesBalance.RowCount - 1
                    rowDGV = Me.dgvSalesBalance.Rows(i)
                    rowDGV.Cells(Me.ColReceivement.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)

                    rowDGV.Cells(Me.ColBalanceSales.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDGV.Cells(Me.ColSalesValue.Index).Value) + _
                    CDec(rowDGV.Cells(Me.ColVatSalesValue.Index).Value) + CDec(rowDGV.Cells(Me.ColDutystamp.Index).Value) - _
                    CDec(rowDGV.Cells(Me.ColReceived.Index).Value))

                    If Angka <> 0 Then
                        'If CDec(rowDGV.Cells(Me.ColTotalSales_VAT_Value.Index).Value) < 0 Then
                        If CDec(rowDGV.Cells(Me.ColBalanceSales.Index).Value) > Angka Then
                            'yang dipindahkan adalah sejumlah balance sales
                            rowDGV.Cells(Me.ColReceivement.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColBalanceSales.Index).Value, Me.DefaultPrecision)
                            Angka -= CDec(rowDGV.Cells(Me.ColBalanceSales.Index).Value)
                            rowDGV.Cells(Me.ColBalanceSales.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                        Else
                            rowDGV.Cells(Me.ColReceivement.Index).Value = Me.ObjCD.DigitGrouping(Angka, Me.DefaultPrecision)
                            rowDGV.Cells(Me.ColBalanceSales.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDGV.Cells(Me.ColSalesValue.Index).Value) + _
                            CDec(rowDGV.Cells(Me.ColVatSalesValue.Index).Value) + CDec(rowDGV.Cells(Me.ColDutystamp.Index).Value) - _
                            CDec(rowDGV.Cells(Me.ColReceived.Index).Value) - Angka, Me.DefaultPrecision)
                            Angka = 0
                        End If
                        'End If
                    End If
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSalesBalance_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvSalesBalance.CellBeginEdit
        Try
            'yang boleh di edit adalah receivement original
            If e.ColumnIndex = Me.ColReceivementOriginal.Index Then
                If Me.ApakahBolehMasukkanPayment(e.RowIndex) = True Then
                    Me.Buffer = Me.dgvSalesBalance.Item(e.ColumnIndex, e.RowIndex).Value
                    Me.dgvSalesBalance.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvSalesBalance.Item(e.ColumnIndex, e.RowIndex).Value.ToString)
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

    Private Sub dgvSalesBalance_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSalesBalance.CellContentClick
        Try
            If Me.TxtRemarks.Enabled = True Then
                If e.ColumnIndex = Me.ColPayInv.Index Then
                    If Me.ApakahBolehMasukkanPayment(e.RowIndex) = True Then
                        Me.SetFullPayment(e.RowIndex)
                        If Me.ApakahAdaCurrency(Me.dgvSalesBalance.Item(Me.ColPrimaryCurrencyBalance.Index, e.RowIndex).Value) = False Then
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

    Private Sub SetFullPayment(ByVal vRowIndex As Integer)
        Dim row As DataGridViewRow
        Dim DefPrec As Integer = 0

        Try
            DefPrec = Me.ObjCur.GetDecimal_DefaultCurrency

            'hitung ulang pembyaran sampai lunas
            row = Me.dgvSalesBalance.Rows(vRowIndex)
            'hitung pada bagian original
            row.Cells(Me.ColReceivementOriginal.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColBalanceOriginal01.Index).Value) - _
            CDec(row.Cells(Me.ColReceivedOriginal.Index).Value), Me.ObjCur.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrencyBalance.Index).Value))
            'receivement IDR
            row.Cells(Me.ColReceivement.Index).Value = _
            Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColReceivementOriginal.Index).Value) * _
            CDec(row.Cells(Me.ColExhangeRateBalance.Index).Value), DefPrec)
            row.Cells(Me.ColBalanceSales.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
            'update data sales receivement
            Me.TxtSalesReceivement.Text = Me.ObjCD.DigitGrouping(Me.GetTotalReceivementSales(), Me.DefaultPrecision)
            row.Cells(Me.ColBalanceOriginal.Index).Value = Me.ObjCD.DigitGrouping(0, row.Cells(Me.ColCurrencyDefPrec.Index).Value)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSalesBalance_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSalesBalance.CellEndEdit
        Dim Balance As Decimal = 0
        Dim DefPrec As Integer = 0
        Dim DefPrecPrimaryCur As Integer = 0

        Try
            'yang boleh di edit adalah
            If e.ColumnIndex = Me.ColReceivementOriginal.Index Then
                If IsNothing(Me.dgvSalesBalance.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    'langsung di balikin ke nilai semula
                    Me.dgvSalesBalance.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Else
                    If e.ColumnIndex = Me.ColReceivementOriginal.Index Then
                        DefPrec = Me.dgvSalesBalance.Item(Me.ColCurrencyDefPrec.Index, e.RowIndex).Value
                        DefPrecPrimaryCur = Me.ObjCur.GetDecimal_DefaultCurrency
                        If CDec(Me.dgvSalesBalance.Item(Me.ColTotalSales_VAT_Value.Index, _
                        e.RowIndex).Value) > 0 Then
                            If Me.ObjCD.CekNonNegatif(Me.dgvSalesBalance.Item(e.ColumnIndex, _
                                e.RowIndex).Value) = False Then
                                MsgBox("Please enter positive number", MsgBoxStyle.Information, _
                                "Information")
                                Me.dgvSalesBalance.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                        Else
                            If Me.ObjCD.CekNonPositif(Me.dgvSalesBalance.Item(e.ColumnIndex, _
                                e.RowIndex).Value) = False Then
                                MsgBox("Please enter negative number", MsgBoxStyle.Information, _
                                "Information")
                                Me.dgvSalesBalance.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                        End If

                        'rounding 0
                        Me.dgvSalesBalance.Item(e.ColumnIndex, e.RowIndex).Value = _
                        Me.ObjCD.DigitGrouping(Me.dgvSalesBalance.Item(e.ColumnIndex, e.RowIndex).Value, _
                        DefPrec)

                        'pastikan nilai balance tidak menjadi negatif
                        Balance = CDec(Me.dgvSalesBalance.Item(Me.ColBalanceOriginal01.Index, e.RowIndex).Value) - _
                        CDec(Me.dgvSalesBalance.Item(Me.ColReceivedOriginal.Index, e.RowIndex).Value) - _
                        CDec(Me.dgvSalesBalance.Item(Me.ColReceivementOriginal.Index, e.RowIndex).Value)

                        If CDec(Me.dgvSalesBalance.Item(Me.ColTotalSales_VAT_Value.Index, _
                        e.RowIndex).Value) > 0 Then
                            If Balance < 0 Then
                                MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping((CDec(Me.dgvSalesBalance.Item(Me.ColReceivementOriginal.Index, e.RowIndex).Value) + Balance), DefPrec) + ".", MsgBoxStyle.Information, _
                                "Information")
                                Me.dgvSalesBalance.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                        Else
                            If Balance > 0 Then
                                MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping((CDec(Me.dgvSalesBalance.Item(Me.ColReceivementOriginal.Index, e.RowIndex).Value) + Balance), DefPrec) + ".", MsgBoxStyle.Information, _
                                "Information")
                                Me.dgvSalesBalance.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                        End If

                        'hitung ulang nilai balance original
                        Me.dgvSalesBalance.Item(Me.ColBalanceOriginal.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Balance, DefPrec)
                        'hitung receivement
                        Me.dgvSalesBalance.Item(Me.ColReceivement.Index, e.RowIndex).Value = _
                        ObjCD.DigitGrouping(CDec(Me.dgvSalesBalance.Item(Me.ColExhangeRateBalance.Index, e.RowIndex).Value) * _
                        Me.dgvSalesBalance.Item(Me.ColReceivementOriginal.Index, e.RowIndex).Value, DefPrecPrimaryCur)
                        'hitung balance
                        Balance = CDec(Me.dgvSalesBalance.Item(Me.ColBalance1.Index, e.RowIndex).Value) - CDec(Me.dgvSalesBalance.Item(Me.ColReceivement.Index, e.RowIndex).Value) - CDec(Me.dgvSalesBalance.Item(Me.ColReceived.Index, e.RowIndex).Value)
                        Me.dgvSalesBalance.Item(Me.ColBalanceSales.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Balance, DefPrecPrimaryCur)

                        'digit grouping
                        Me.dgvSalesBalance.Item(e.ColumnIndex, e.RowIndex).Value = _
                        ObjCD.DigitGrouping(Me.dgvSalesBalance.Item(e.ColumnIndex, e.RowIndex).Value, DefPrec)
                        'update data sales receivement
                        Me.TxtSalesReceivement.Text = Me.ObjCD.DigitGrouping(Me.GetTotalReceivementSales(), DefPrec)
                        If Me.ApakahAdaCurrency(Me.dgvSalesBalance.Item(Me.ColPrimaryCurrencyBalance.Index, e.RowIndex).Value) = False Then
                            Me.SetCBOPayment()
                        End If
                        Me.UpdateJurnal()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDistribusi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDistribusi.Click
        Dim Angka As Decimal = Me.TxtSalesReturnPayment.Text
        Dim rowDGV As DataGridViewRow

        Try
            Angka = Me.TxtSalesReturnPayment.Text
            For i As Double = 0 To Me.dgvSalesReturnBalance.RowCount - 1
                rowDGV = Me.dgvSalesReturnBalance.Rows(i)
                rowDGV.Cells(Me.ColReturnPayment.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)

                rowDGV.Cells(Me.ColBalanceReturn.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDGV.Cells(Me.ColSalesReturnValue.Index).Value) + _
                CDec(rowDGV.Cells(Me.ColVATReturnValue.Index).Value) + CDec(rowDGV.Cells(Me.ColDutystampReturn.Index).Value) - _
                CDec(rowDGV.Cells(Me.ColPaidReturn.Index).Value))

                If Angka <> 0 Then
                    If CDec(rowDGV.Cells(Me.ColBalanceReturn.Index).Value) < Angka Then
                        'yang dipindahkan adalah sejumlah balance sales
                        rowDGV.Cells(Me.ColReturnPayment.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColBalanceReturn.Index).Value, Me.DefaultPrecision)
                        Angka -= CDec(rowDGV.Cells(Me.ColBalanceReturn.Index).Value)
                        rowDGV.Cells(Me.ColBalanceReturn.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                    Else
                        rowDGV.Cells(Me.ColReturnPayment.Index).Value = Me.ObjCD.DigitGrouping(Angka, Me.DefaultPrecision)
                        rowDGV.Cells(Me.ColBalanceReturn.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDGV.Cells(Me.ColSalesReturnValue.Index).Value) + _
                        CDec(rowDGV.Cells(Me.ColVATReturnValue.Index).Value) + CDec(rowDGV.Cells(Me.ColDutystampReturn.Index).Value) - _
                        CDec(rowDGV.Cells(Me.ColPaidReturn.Index).Value) - Angka)
                        Angka = 0
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSalesReturnBalance_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvSalesReturnBalance.CellBeginEdit
        Try
            'yang boleh di edit adalah data payment saja
            If e.ColumnIndex = Me.ColReturnPayment.Index Then
                Me.Buffer = Me.dgvSalesReturnBalance.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSalesReturnBalance_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSalesReturnBalance.CellEndEdit
        Dim Balance As Decimal = 0

        Try
            'yang boleh di edit adalah
            If e.ColumnIndex = Me.ColReturnPayment.Index Then
                If IsNothing(Me.dgvSalesReturnBalance.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    'langsung di balikin ke nilai semula
                    Me.dgvSalesReturnBalance.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Else
                    ''cek debet
                    If e.ColumnIndex = Me.ColReturnPayment.Index Then
                        If Me.ObjCD.CekNonNegatif(Me.dgvSalesReturnBalance.Item(e.ColumnIndex, _
                            e.RowIndex).Value) = False Then
                            MsgBox("Please enter positive number", MsgBoxStyle.Information, _
                            "Information")
                            Me.dgvSalesReturnBalance.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        'rounding 0
                        Me.dgvSalesReturnBalance.Item(e.ColumnIndex, e.RowIndex).Value = _
                        Me.ObjCD.DigitGrouping(Me.dgvSalesReturnBalance.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefaultPrecision)
                        'pastikan nilai balance tidak menjadi negatif
                        Balance = CDec(Me.dgvSalesReturnBalance.Item(Me.ColSalesReturnValue.Index, e.RowIndex).Value) + _
                        CDec(Me.dgvSalesReturnBalance.Item(Me.ColVATReturnValue.Index, e.RowIndex).Value) + _
                        CDec(Me.dgvSalesReturnBalance.Item(Me.ColDutystampReturn.Index, e.RowIndex).Value) - _
                        CDec(Me.dgvSalesReturnBalance.Item(Me.ColPaidReturn.Index, e.RowIndex).Value)
                        If Balance - CDec(Me.dgvSalesReturnBalance.Item(e.ColumnIndex, e.RowIndex).Value) < 0 Then
                            MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(Balance, Me.DefaultPrecision) + ".", MsgBoxStyle.Information, _
                            "Information")
                            Me.dgvSalesReturnBalance.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        'hitung ulang nilai balance
                        Me.dgvSalesReturnBalance.Item(Me.ColBalanceReturn.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Balance - _
                        CDec(Me.dgvSalesReturnBalance.Item(e.ColumnIndex, e.RowIndex).Value), Me.DefaultPrecision)
                        'digit grouping
                        Me.dgvSalesReturnBalance.Item(e.ColumnIndex, e.RowIndex).Value = _
                        ObjCD.DigitGrouping(Me.dgvSalesReturnBalance.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefaultPrecision)
                        'update data sales return payment
                        Me.TxtSalesReturnPayment.Text = Me.ObjCD.DigitGrouping(Me.GetTotalPaymentSalesReturn, Me.DefaultPrecision)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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
                    If CDec(row.Cells(Me.ColDebet.Index).Value) <> 0 Then
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
                    If CDec(row.Cells(Me.ColKredit.Index).Value) <> 0 Then
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
                    If Me.dgvItem.AllowUserToAddRows = False Then
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

    Private Sub GetSelisihExchRate(ByVal TblSelisihExchRate As DseTransaksiJournal.TblSelisihExchRateDataTable, ByVal DefPrecBase As Integer)
        Dim rowSelisihExchRate As DseTransaksiJournal.TblSelisihExchRateRow
        Dim SelisihExchRate As Decimal = 0
        Dim Filter As String = ""
        Dim rowDGV As DataGridViewRow
        Dim PrimaryAkunSelisihExchRate As Double = -1

        Try
            For Each rowDGV In Me.dgvSalesBalance.Rows
                If CDbl(rowDGV.Cells(Me.ColPrimaryCurrencyBalance.Index).Value) <> Core_Module.PrimaryBaseCurrency And _
                CDec(rowDGV.Cells(Me.ColReceivementOriginal.Index).Value) <> 0 Then
                    SelisihExchRate = Me.ObjCD.DigitGrouping(CDec(rowDGV.Cells(Me.ColReceivementOriginal.Index).Value) * _
                    (CDec(rowDGV.Cells(Me.ColExhangeRateBalance.Index).Value) - CDec(Me.TxtExchangeRate.Text)), DefPrecBase)
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

    Public Sub UpdateJurnal()
        Dim rowDGV As DataGridViewRow
        Dim TotalARPaid As Decimal = 0
        Dim PrimaryAkunDifference As String = ""
        Dim rowPayment As DseMasterAkun.tblmaster_akun4Row
        Dim IndexAkunAR As Integer = 0
        Dim DefPrec As Integer = 0
        Dim DefPrecBase As Integer = 0
        Dim TblSelisihExchRate As New DseTransaksiJournal.TblSelisihExchRateDataTable
        Dim rowSelisihExchRate As DseTransaksiJournal.TblSelisihExchRateRow
        Dim Payment As Decimal = 0

        Try
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
            For Each rowDGV In Me.dgvSalesBalance.Rows
                If CDec(rowDGV.Cells(Me.ColReceivement.Index).Value) <> 0 Then
                    TotalARPaid += 1
                    Exit For
                End If
            Next
            If TotalARPaid = 0 Then
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
                        Payment = Math.Round(CDec(Me.TxtPaid.Text) * CDec(Me.TxtExchangeRate.Text), DefPrecBase)
                    Else
                        Payment = Math.Round(CDec(Me.TxtPaid.Text), DefPrecBase)
                    End If

                    If CDec(Me.TxtPaid.Text) > 0 Then
                        'debet ada isinya
                        .Cells(Me.ColDebet.Index).Value = Me.ObjCD.DigitGrouping(Payment, DefPrecBase)
                        .Cells(Me.ColKredit.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrecBase)
                        .Cells(Me.ColDebetOriginal.Index).Value = Me.TxtPaid.Text
                        .Cells(Me.ColCreditOriginal.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
                    Else
                        'credit ada isinya
                        .Cells(Me.ColKredit.Index).Value = Me.ObjCD.DigitGrouping(Math.Abs(Payment), DefPrecBase)
                        .Cells(Me.ColDebet.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrecBase)
                        .Cells(Me.ColCreditOriginal.Index).Value = Me.ObjCD.DigitGrouping(Math.Abs(CDec(Me.TxtPaid.Text)), DefPrec)
                        .Cells(Me.ColDebetOriginal.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
                    End If
                End With
            End If

            'akun ar, diisikan yg debet dulu baru yg credit
            For Each rowDGV In Me.dgvSalesBalance.Rows
                With rowDGV
                    If CDec(rowDGV.Cells(Me.ColReceivementOriginal.Index).Value) > 0 Then
                        IndexAkunAR = Me.GetIndexAkunAR(.Cells(Me.ColPrimaryAkunBalance.Index).Value, False, .Cells(Me.ColExhangeRateBalance.Index).Value)
                        If IndexAkunAR < 0 Then
                            'tambahkan baris baru
                            Me.AddRowItem(.Cells(Me.ColPrimaryAkunBalance.Index).Value, _
                            .Cells(Me.ColExhangeRateBalance.Index).Value, .Cells(Me.ColPrimaryCurrencyBalance.Index).Value, _
                            0, Math.Abs(CDec(.Cells(Me.ColReceivement.Index).Value)), 0, Math.Abs(CDec(.Cells(Me.ColReceivementOriginal.Index).Value)))
                        Else
                            'tambahkan debet/kredit saja
                            Me.AddValue(Math.Abs(CDec(.Cells(Me.ColReceivement.Index).Value)), Math.Abs(CDec(.Cells(Me.ColReceivementOriginal.Index).Value)), _
                            IndexAkunAR, False)
                        End If
                    ElseIf CDec(rowDGV.Cells(Me.ColReceivementOriginal.Index).Value) < 0 Then
                        IndexAkunAR = Me.GetIndexAkunAR(.Cells(Me.ColPrimaryAkunBalance.Index).Value, True, .Cells(Me.ColExhangeRateBalance.Index).Value)

                        If IndexAkunAR < 0 Then
                            'tambahkan baris baru
                            Me.AddRowItem(.Cells(Me.ColPrimaryAkunBalance.Index).Value, _
                            .Cells(Me.ColExhangeRateBalance.Index).Value, .Cells(Me.ColPrimaryCurrencyBalance.Index).Value, _
                             Math.Abs(CDec(.Cells(Me.ColReceivement.Index).Value)), 0, Math.Abs(CDec(.Cells(Me.ColReceivementOriginal.Index).Value)), 0)
                        Else
                            'tambahkan debet/kredit saja
                            Me.AddValue(Math.Abs(CDec(.Cells(Me.ColReceivement.Index).Value)), Math.Abs(CDec(.Cells(Me.ColReceivementOriginal.Index).Value)), _
                            IndexAkunAR, True)
                        End If
                    End If
                End With
            Next

            'akun selisih pembayaran
            TblSelisihExchRate.Rows.Clear()
            Me.GetSelisihExchRate(TblSelisihExchRate, DefPrecBase)
            For Each rowSelisihExchRate In TblSelisihExchRate.Rows
                If rowSelisihExchRate.SelisihExchRate <> 0 Then
                    With rowSelisihExchRate
                        'tambahkan bari baru dalam dgv
                        If .SelisihExchRate > 0 Then
                            'masukkan ke dalam debet
                            Me.AddRowItem(.Primary_Akun_SelisihExchRate, 1, Core_Module.PrimaryBaseCurrency, _
                            Math.Abs(.SelisihExchRate), 0, Math.Abs(.SelisihExchRate), 0)
                        Else
                            'masukkan ke dalam kredit
                            Me.AddRowItem(.Primary_Akun_SelisihExchRate, 1, Core_Module.PrimaryBaseCurrency, 0, _
                            Math.Abs(.SelisihExchRate), 0, Math.Abs(.SelisihExchRate))
                        End If
                    End With
                End If
            Next

            'akun difference
            If CDec(Me.TxtDifference.Text) <> 0 Then
                If CDec(Me.TxtDifference.Text) > 0 Then
                    'masuk di kredit
                    Me.AddRowItem(Me.TxtPrimaryAccountDifference.Text, 1, Core_Module.PrimaryBaseCurrency, 0, _
                    Math.Abs(CDec(Me.TxtDifference.Text)), 0, Math.Abs(CDec(Me.TxtDifference.Text)))
                Else
                    'masuk di debet
                    Me.AddRowItem(Me.TxtPrimaryAccountDifference.Text, 1, Core_Module.PrimaryBaseCurrency, _
                    Math.Abs(CDec(Me.TxtDifference.Text)), 0, Math.Abs(CDec(Me.TxtDifference.Text)), 0)
                End If
            End If

            Me.HitungUnbalance()
            Me.SetMulticurrency()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TblSelisihExchRate = Nothing
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

    Private Function GetIndexAkunAR(ByVal vPrimaryAkunAR As Double, ByVal Debet As Boolean, ByVal vExhangeRate As Decimal) As Double
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
    ByVal vCredit As Decimal, ByVal vDebetOriginal As Decimal, ByVal vCreditOriginal As Decimal)
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
                .Cells(Me.ColRemarks.Index).Value = ""
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

    Private Sub HitungDifference()
        Dim TotalARPaid, TotalPayment As Decimal
        Dim rowDGV As DataGridViewRow
        Dim DefPrec As Integer = 0
        Dim TblSelisihExcRate As New DseTransaksiJournal.TblSelisihExchRateDataTable
        Dim rowSelisihExcRate As DseTransaksiJournal.TblSelisihExchRateRow
        Dim SelisihExcRate As Decimal = 0
        Dim rowAkunPayment As DseMasterAkun.tblmaster_akun4Row

        Try
            DefPrec = Me.ObjCur.GetDecimal_DefaultCurrency
            For Each rowDGV In Me.dgvSalesBalance.Rows
                TotalARPaid += CDec(rowDGV.Cells(Me.ColReceivement.Index).Value)
            Next
            Me.TxtARPaid_IDR.Text = Me.ObjCD.DigitGrouping(TotalARPaid, DefPrec)

            If Me.cboPayToAcc.SelectedIndex <> -1 Then
                rowAkunPayment = Me.TBLAkun.DefaultView(Me.cboPayToAcc.SelectedIndex).Row
                If rowAkunPayment.primary_currency = Core_Module.PrimaryBaseCurrency Then
                    TotalPayment = Math.Round(CDec(Me.TxtPaid.Text), DefPrec, MidpointRounding.AwayFromZero)
                Else
                    TotalPayment = Math.Round(CDec(Me.TxtPaid.Text) * CDec(Me.TxtExchangeRate.Text), DefPrec, MidpointRounding.AwayFromZero)
                End If
                Me.TxtPaid_InIDR.Text = Me.ObjCD.DigitGrouping(TotalPayment, DefPrec)
            Else
                TotalPayment = Math.Round(CDec(Me.TxtPaid.Text), DefPrec, MidpointRounding.AwayFromZero)
            End If

            Me.GetSelisihExchRate(TblSelisihExcRate, DefPrec)
            For Each rowSelisihExcRate In TblSelisihExcRate.Rows
                SelisihExcRate += rowSelisihExcRate.SelisihExchRate
            Next
            Me.TxtSelisihExcRate.Text = Me.ObjCD.DigitGrouping(SelisihExcRate, DefPrec)

            Me.TxtDifference.Text = Me.ObjCD.DigitGrouping(TotalPayment - TotalARPaid + SelisihExcRate, DefPrec)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TblSelisihExcRate = Nothing
        End Try
    End Sub

    Private Function ApakahAdaPayment() As Boolean
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvSalesBalance.Rows
                If CDec(rowDGV.Cells(Me.ColReceivementOriginal.Index).Value) <> 0 Then
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetPrimaryCurrencyDokumenAcuan(ByVal vTipeDokumenAcuan As ClsTransaksiSalesBalance.EnumTipeSalesBalance, _
    ByVal vPrimaryDokumenAcuan As Double) As Double
        Try
            If vTipeDokumenAcuan = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic Then
                Return Me.ObjInv.GetPrimaryCurrency(vPrimaryDokumenAcuan)
            ElseIf vTipeDokumenAcuan = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR Then
                Return Me.ObjSR.GetPrimaryCurrency(vPrimaryDokumenAcuan)
            ElseIf vTipeDokumenAcuan = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual Then
                Return Me.ObjInvoiceBalance.GetPrimaryCurrency(vPrimaryDokumenAcuan)
            ElseIf vTipeDokumenAcuan = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR Then
                Me.ObjSRBalance.GetPrimaryCurrency(vPrimaryDokumenAcuan)
            Else
                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Function GetPrimaryCurrencyPayment() As Double
        Dim rowDGV As DataGridViewRow
        Dim PrimaryCurDokumenAcuan As Double = -1

        Try
            For Each rowDGV In Me.dgvSalesBalance.Rows
                If CDec(rowDGV.Cells(Me.ColReceivementOriginal.Index).Value) <> 0 Then
                    If CDec(rowDGV.Cells(Me.ColSalesValue.Index).Value) <> 0 Then
                        'ambil primary currency document asal
                        PrimaryCurDokumenAcuan = Me.GetPrimaryCurrencyDokumenAcuan(rowDGV.Cells(Me.ColTipeDocumentAcuan_Sales.Index).Value, _
                        rowDGV.Cells(Me.ColPrimaryDocumenAcuan_Sales.Index).Value)
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

    Private Function GetCurrencyNonIDR_DGVSalesBalance_Paid() As Double
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvSalesBalance.Rows
                If CDec(row.Cells(Me.ColReceivementOriginal.Index).Value) <> 0 Then
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

    Private Function ApakahAdaCurrency_DGVSalesBalance_Paid(ByVal vPrimaryCurr As Double) As Boolean
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvSalesBalance.Rows
                If CDec(row.Cells(Me.ColReceivementOriginal.Index).Value) <> 0 Then
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
                    For Each rowDGV In Me.dgvSalesBalance.Rows
                        If CDec(rowDGV.Cells(Me.ColReceivementOriginal.Index).Value) <> 0 Then
                            If rowDGV.Cells(Me.ColPrimaryCurrencyBalance.Index).Value <> Core_Module.PrimaryBaseCurrency Then
                                If Me.ApakahAdaCurrency(rowDGV.Cells(Me.ColPrimaryCurrencyBalance.Index).Value) = False Then
                                    IsikanData = True
                                    Exit For
                                End If
                            ElseIf rowDGV.Cells(Me.ColSalesValue.Index).Value = 0 Then
                                TipeTransaksiAsal = rowDGV.Cells(Me.ColTipeDocumentAcuan_Sales.Index).Value
                                PrimaryCurrency = Me.GetPrimaryCurrencyDokumenAcuan(TipeTransaksiAsal, _
                                rowDGV.Cells(Me.ColPrimaryDocumenAcuan_Sales.Index).Value)
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
                            For Each rowDGV In Me.dgvSalesBalance.Rows
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

    'payment bisa diedit jika currency-nya sama atau currencynya akan sama dengan base currency
    Private Function ApakahBolehMasukkanPayment(ByVal vRowIndex As Integer) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim PrimaryCurrencyAcuan As Double = -1

        Try
            rowDGV = Me.dgvSalesBalance.Rows(vRowIndex)

            'jika sama dengan primary base currency maka bisa diedit
            If rowDGV.Cells(Me.ColPrimaryCurrencyBalance.Index).Value = Core_Module.PrimaryBaseCurrency Then
                Return True
            End If

            'jika tidak ada currency lain (non-idr) yang beda dengan yang akan diedit
            PrimaryCurrencyAcuan = rowDGV.Cells(Me.ColPrimaryCurrencyBalance.Index).Value
            For Each rowDGV In Me.dgvSalesBalance.Rows
                If rowDGV.Index <> vRowIndex Then
                    If CDec(rowDGV.Cells(Me.ColReceivementOriginal.Index).Value) <> 0 Then
                        If rowDGV.Cells(Me.ColPrimaryCurrencyBalance.Index).Value <> PrimaryCurrencyAcuan And _
                        rowDGV.Cells(Me.ColPrimaryCurrencyBalance.Index).Value <> Core_Module.PrimaryBaseCurrency Then
                            Return False
                        End If
                    End If
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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

    Private Sub SetGBOEquivalent()
        Dim row As DseMasterAkun.tblmaster_akun4Row
        Dim rowDGV As DataGridViewRow
        Dim Beda As Boolean = False

        Try
            If Me.cboPayToAcc.SelectedIndex = -1 Then
                Me.gboEquivalent.Visible = False
            Else
                row = Me.TBLAkun.DefaultView(Me.cboPayToAcc.SelectedIndex).Row
                For Each rowDGV In Me.dgvSalesBalance.Rows
                    If row.primary_currency <> rowDGV.Cells(Me.ColPrimaryCurrencyBalance.Index).Value Then
                        Beda = True
                        Exit For
                    End If
                Next

                If Beda = False Then
                    Me.gboEquivalent.Visible = False
                Else
                    Me.gboEquivalent.Visible = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

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
                    If Me.ApakahAdaCurrency_DGVSalesBalance_Paid(Core_Module.PrimaryBaseCurrency) = False Then
                        '1. berasal dari item dgv sales balance pure non idr
                        'jika dibayarkan pada akun non idr, berasal dari item sales balance yg non idr
                        Me.TxtPaid.Text = Me.ObjCD.DigitGrouping(Me.GetTotalReceivementOriginalSales, _
                        Me.ObjCur.GetDefaultPrecission(row.primary_currency))
                    Else
                        '2. berasal dari item dgv sales balance campuran
                        'jika dibayarkan pada akun non idr, berasal dari item sales balance yg idr + non idr
                        Me.TxtPaid.Text = Me.ObjCD.DigitGrouping((CDec(Me.TxtARPaid_IDR.Text) - CDec(Me.TxtSelisihExcRate.Text)) / CDec(Me.TxtExchangeRate.Text), Me.ObjCur.GetDefaultPrecission(row.primary_currency))
                    End If
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
                    If PrimaryCurrencyAkun <> Core_Module.PrimaryBaseCurrency And PrimaryCurrencyAkun <> PrimaryCurPayment Then
                        MsgBox("Please select payment account with " + Me.ObjCur.GetCurrencyCode(Core_Module.PrimaryBaseCurrency) + _
                        " or " + Me.ObjCur.GetCurrencyCode(PrimaryCurPayment) + " currency")
                        Exit Sub
                    End If

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
                        Me.cboPayToAcc.SelectedIndex = -1

                        'tambahkan baris baru pada tabel
                        row = Me.TBLAkun.Newtblmaster_akun4Row
                        row.Primary = i.mPrimaryAkun
                        row.primary_currency = Me.ObjAkun.GetPrimaryCurrency(i.mPrimaryAkun)
                        row.Description = Me.ObjAkun.GetDescription_ByPrimaryAcc(i.mPrimaryAkun)
                        row.Account_Code = Me.ObjAkun.GetAccountCode(i.mPrimaryAkun)
                        Me.TBLAkun.Rows.Add(row)

                        'pilih account yang dipilih pada form cari
                        For Each row In Me.TBLAkun.Rows
                            If row.Primary = i.mPrimaryAkun Then
                                Me.cboPayToAcc.SelectedIndex = m
                                Exit For
                            End If
                            m += 1
                        Next
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Try
            If Me.TxtRefNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtRefNo.Text <> "" Then
                Dim i As New FrmTransaksiJournal_Cetak_Sales
                i.PrimaryJurnal = Me.TxtPrimaryJournal.Text
                i.vDGV = Me.dgvSalesBalance
                i.vFrm = Me

                i.ShowDialog()
                i.Dispose()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class