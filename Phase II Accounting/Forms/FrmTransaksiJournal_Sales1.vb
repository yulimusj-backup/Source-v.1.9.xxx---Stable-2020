Option Explicit On
Option Strict Off

Imports Core_Module
Public Class FrmTransaksiJournal_Sales1
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightSeaGreen
    Public TipeJournal As ClsTransaksiJournal.EnumTipeJournal = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal
    Public ObjQA As New ClsTransaksiJournal

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

    Dim Buffer As String = ""
    Dim DefaultPrecision As Integer = 0
    Dim Baru As Boolean
    Dim ObjAkun As New ClsMasterAkun
    Dim ObjCust As New clsMasterCustomer
    Dim ObjJurnalItem As New ClsTransaksiJournal_Item
    Dim ObjQuickJurnal As New ClsTransaksiJournalQuickAdd
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjTaxRate As New ClsTransaksiSalesInvoiceTaxRate

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
                    Me.InsertQuickJurnal(False)
                End If

                For Each rowDGV In Me.dgvSalesBalance.Rows
                    If rowDGV.Cells(Me.ColTipeDocumentAcuan_Sales.Index).Value = vTipeDokumentAcuan And _
                    vPrimaryDokumentAcuan = rowDGV.Cells(Me.ColPrimaryDocumenAcuan_Sales.Index).Value Then
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

    Private Sub FrmTransaksiSalesCycle_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim NullValue As String = ""

        Try
            Me.WindowState = FormWindowState.Maximized
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
                If Me.ObjQA.CariDetailSalesJournal(ij.PrimaryPP, Me.dgvItem, Me, ij.TipeJournal) = True Then
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

    Private Sub AktifkanDetail()
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

            Me.TxtPrimaryJournal.Text = "-1"
            Me.TxtPrimaryCust.Text = "-1"
            Me.TxtUnbalance.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            Me.TxtSalesReceivement.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            Me.TxtSalesReturnPayment.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "

            Me.dgvItem.Rows.Clear()
            Me.dgvSalesBalance.Rows.Clear()
            Me.dgvSalesReturnBalance.Rows.Clear()
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
                .PrimaryPihakAcuan = Me.TxtPrimaryCust.Text
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
            End With
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
        Dim Total1, Total2, Total3 As Double
        Dim PrimaryAkunPiutang As Double = -1
        Dim NilaiPayment As Decimal = 0
        Dim NilaiPayment_AkunPiutang As Decimal = 0

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
                        Me.ObjQA.CariDetailSalesJournal(Me.ObjQA.Primary, Me.dgvItem, Me, Me.TipeJournal)
                        Me.SalinkeField()
                    End If
                    Return False
                End If
            End If

            'receivement/payment tidak boleh nol
            If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
                If Me.GetTotalReceivementSales = 0 Then
                    MsgBox("Please specify receivement data", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            ElseIf Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                If Me.GetTotalPaymentSalesReturn = 0 Then
                    MsgBox("Please specify payment data", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            'biasanya total debet dan total kredit harus sama dengan total receivement
            If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
                Total1 = Me.GetTotalDebet
                Total2 = Me.GetTotalReceivementSales
                Total3 = Me.GetTotalKredit

                If Total2 > 0 Then
                    'sales receivement
                    If Total1 < Total2 Then
                        MsgBox("Total receivement is larger than total debet.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    If Total1 <> Total2 Then
                        If MsgBox("Total receivement is not equal to total debet. Continue?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                            Return False
                        End If
                    End If
                Else
                    'sales return payment
                    If Total1 < Total2 Then
                        MsgBox("Total payment is larger than total debet.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    If Total1 <> Math.Abs(Total2) Then
                        If MsgBox("Total payment is not equal to total debet. Continue?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                            Return False
                        End If
                    End If
                End If
            ElseIf Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.SalesReturnPaymentJournal Then
                Total1 = Me.GetTotalDebet
                Total2 = Me.GetTotalPaymentSalesReturn
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

            'hitung total payment
            Me.TxtSalesReceivement.Text = Me.ObjCD.DigitGrouping(Me.GetTotalReceivementSales)
            'nilai payment harus sama dengan nilai debet pada akun hutang
            If Me.dgvItem.RowCount > 1 Then
                If IsNumeric(Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.SalesJurnal, ClsTransaksiSetting.EnumNamaParameter.Akun_AR)) = True Then
                    PrimaryAkunPiutang = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.SalesJurnal, ClsTransaksiSetting.EnumNamaParameter.Akun_AR)
                    If PrimaryAkunPiutang <> -1 Then
                        If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
                            NilaiPayment = CDec(Me.TxtSalesReceivement.Text)
                        Else
                            NilaiPayment = CDec(Me.TxtSalesReturnPayment.Text)
                        End If

                        For Each row In Me.dgvItem.Rows
                            If IsNothing(row.Cells(Me.ColPrimaryAkun.Index).Value) = False Then
                                If CDbl(row.Cells(Me.ColPrimaryAkun.Index).Value) = PrimaryAkunPiutang Then
                                    If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
                                        NilaiPayment_AkunPiutang = CDec(row.Cells(Me.ColKredit.Index).Value) - CDec(row.Cells(Me.ColDebet.Index).Value)
                                    Else
                                        NilaiPayment_AkunPiutang = CDec(row.Cells(Me.ColDebet.Index).Value) - CDec(row.Cells(Me.ColKredit.Index).Value)
                                    End If
                                End If
                            End If
                        Next
                        If NilaiPayment <> NilaiPayment_AkunPiutang Then
                            If NilaiPayment > 0 Then
                                If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
                                    MsgBox("The credit value for '" + Me.ObjAkun.GetDescription_ByPrimaryAcc(PrimaryAkunPiutang) + "' account must be " + Me.ObjCD.DigitGrouping(NilaiPayment), _
                                    MsgBoxStyle.Information, "Information")
                                Else
                                    MsgBox("The debet value for '" + Me.ObjAkun.GetDescription_ByPrimaryAcc(PrimaryAkunPiutang) + "' account must be " + Me.ObjCD.DigitGrouping(NilaiPayment), _
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
                        Me.cboQuickJournal.Focus()

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

            'isikan data balance
            Me.IsikanBalance()
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
                Me.cboQuickJournal.Focus()
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
            If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
                Me.ObjInvoiceBalance.IsiDGVBalance(Me.dgvSalesBalance, Me.TBLSalesBalance, Me, Me.ObjCD, Me.TxtPrimaryCust.Text, Me.TxtPrimaryJournal.Text, Me.DefaultPrecision, Me.TBLSRBalance)
            ElseIf Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.SalesReturnPaymentJournal Then
                Me.ObjSRBalance.IsiDGVBalance(Me.dgvSalesReturnBalance, Me.TBLSRBalance, Me, Me.ObjCD, Me.TxtPrimaryCust.Text, Me.TxtPrimaryJournal.Text, Me.DefaultPrecision)
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
                    Me.ObjQA.CariDetailSalesJournal(Me.ObjQA.Primary, Me.dgvItem, Me, TipeJournal)
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
                        If Me.ObjQA.Hapus_SalesJurnal(Me.TxtPrimaryJournal.Text, True, UserCode, Me.TipeJournal) = True Then
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
                If ObjQA.ApakahDataBenar(Baru) = True Then
                    If Me.ApakahDataBenar() = True Then
                        'konfirmasi disimpan atau tidak
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")

                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
                                hasil_save = Me.ObjQA.SimpanBaruSalesJournal(Me.dgvItem, Me.UserCode, Me, TipeJournal, Me.dgvSalesBalance)
                            Else
                                hasil_save = Me.ObjQA.SimpanBaruSalesJournal(Me.dgvItem, Me.UserCode, Me, TipeJournal, Me.dgvSalesReturnBalance)
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
                            If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
                                hasil_save = Me.ObjQA.SimpanPerubahanSalesJournal(Me.dgvItem, Me.UserCode, Me, TipeJournal, Me.dgvSalesBalance)
                            Else
                                hasil_save = Me.ObjQA.SimpanPerubahanSalesJournal(Me.dgvItem, Me.UserCode, Me, TipeJournal, Me.dgvSalesReturnBalance)
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
                    ''Me.KosongkanItem()
                    If Baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.ObjQA.CariDetailSalesJournal(Me.TxtPrimaryJournal.Text, Me.dgvItem, Me, TipeJournal)
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
                                    hasil_res = Me.ObjQA.SimpanBaruSalesJournal(Me.dgvItem, Me.UserCode, Me, TipeJournal, Me.dgvSalesBalance)
                                ElseIf Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchaseReturnReceivement Then
                                    hasil_res = Me.ObjQA.SimpanBaruSalesJournal(Me.dgvItem, Me.UserCode, Me, TipeJournal, Me.dgvSalesReturnBalance)
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
                                    hasil_res = Me.ObjQA.SimpanPerubahanSalesJournal(Me.dgvItem, Me.UserCode, Me, TipeJournal, Me.dgvSalesBalance)
                                ElseIf Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchaseReturnReceivement Then
                                    hasil_res = Me.ObjQA.SimpanPerubahanSalesJournal(Me.dgvItem, Me.UserCode, Me, TipeJournal, Me.dgvSalesReturnBalance)
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
            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex)) = True _
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
            'yang boleh di edit adalah
            If e.ColumnIndex = Me.ColReceivement.Index Then
                Me.Buffer = Me.dgvSalesBalance.Item(e.ColumnIndex, e.RowIndex).Value
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
                    Me.SetFullPayment(e.RowIndex)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetFullPayment(ByVal vRowIndex As Integer)
        Dim row As DataGridViewRow

        Try
            'hitung ulang pembyaran sampai lunas
            row = Me.dgvSalesBalance.Rows(vRowIndex)
            row.Cells(Me.ColReceivement.Index).Value = _
            Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColTotalSales_VAT_Value.Index).Value) - _
            CDec(row.Cells(Me.ColReceived.Index).Value), Me.DefaultPrecision)
            row.Cells(Me.ColBalanceSales.Index).Value = 0
            'update data sales receivement
            Me.TxtSalesReceivement.Text = Me.ObjCD.DigitGrouping(Me.GetTotalReceivementSales(), Me.DefaultPrecision)
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
                    row.Cells(Me.ColDebet.Index).Value = Me.TxtSalesReceivement.Text
                ElseIf CDec(row.Cells(Me.ColKredit.Index).Value) <> 0 Then
                    row.Cells(Me.ColKredit.Index).Value = Me.TxtSalesReceivement.Text
                End If
                row = Me.dgvItem.Rows(1)
                If CDec(row.Cells(Me.ColDebet.Index).Value) <> 0 Then
                    row.Cells(Me.ColDebet.Index).Value = Me.TxtSalesReceivement.Text
                ElseIf CDec(row.Cells(Me.ColKredit.Index).Value) <> 0 Then
                    row.Cells(Me.ColKredit.Index).Value = Me.TxtSalesReceivement.Text
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSalesBalance_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSalesBalance.CellEndEdit
        Dim Balance As Decimal = 0

        Try
            'yang boleh di edit adalah
            If e.ColumnIndex = Me.ColReceivement.Index Then
                If IsNothing(Me.dgvSalesBalance.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    'langsung di balikin ke nilai semula
                    Me.dgvSalesBalance.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Else
                    ''cek debet
                    If e.ColumnIndex = Me.ColReceivement.Index Then
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
                        Me.ObjCD.DigitGrouping(Me.dgvSalesBalance.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefaultPrecision)

                        'pastikan nilai balance tidak menjadi negatif
                        Balance = CDec(Me.dgvSalesBalance.Item(Me.ColSalesValue.Index, e.RowIndex).Value) + _
                        CDec(Me.dgvSalesBalance.Item(Me.ColVatSalesValue.Index, e.RowIndex).Value) + _
                        CDec(Me.dgvSalesBalance.Item(Me.ColDutystamp.Index, e.RowIndex).Value) - _
                        CDec(Me.dgvSalesBalance.Item(Me.ColReceived.Index, e.RowIndex).Value)

                        If CDec(Me.dgvSalesBalance.Item(Me.ColTotalSales_VAT_Value.Index, _
                        e.RowIndex).Value) > 0 Then
                            If Balance - CDec(Me.dgvSalesBalance.Item(e.ColumnIndex, e.RowIndex).Value) < 0 Then
                                MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(Balance, Me.DefaultPrecision) + ".", MsgBoxStyle.Information, _
                                "Information")
                                Me.dgvSalesBalance.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                        Else
                            If Balance - CDec(Me.dgvSalesBalance.Item(e.ColumnIndex, e.RowIndex).Value) > 0 Then
                                MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(Balance, Me.DefaultPrecision) + ".", MsgBoxStyle.Information, _
                                "Information")
                                Me.dgvSalesBalance.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                        End If

                        'hitung ulang nilai balance
                        Me.dgvSalesBalance.Item(Me.ColBalanceSales.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Balance - _
                        CDec(Me.dgvSalesBalance.Item(e.ColumnIndex, e.RowIndex).Value), Me.DefaultPrecision)
                        'digit grouping
                        Me.dgvSalesBalance.Item(e.ColumnIndex, e.RowIndex).Value = _
                        ObjCD.DigitGrouping(Me.dgvSalesBalance.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefaultPrecision)
                        'update data sales receivement
                        Me.TxtSalesReceivement.Text = Me.ObjCD.DigitGrouping(Me.GetTotalReceivementSales(), Me.DefaultPrecision)
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
            If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
                Me.ObjQuickJurnal.IsiDGVQuickAdd(Me.TipeJournal, Me.dgvItem, Me.GetTotalReceivementSales, _
                Me.cboQuickJournal.Text, Me.ColAccount.Index, Me.ColDescription.Index, Me.ColPrimaryAkun.Index, _
                Me.ColDebet.Index, Me.ColKredit.Index, Me.ColRemarks.Index, Me.ObjCD, Me.DefaultPrecision, vConfirm)
            ElseIf Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.SalesReturnPaymentJournal Then
                Me.ObjQuickJurnal.IsiDGVQuickAdd(Me.TipeJournal, Me.dgvItem, Me.GetTotalPaymentSalesReturn, _
                Me.cboQuickJournal.Text, Me.ColAccount.Index, Me.ColDescription.Index, Me.ColPrimaryAkun.Index, _
                Me.ColDebet.Index, Me.ColKredit.Index, Me.ColRemarks.Index, Me.ObjCD, Me.DefaultPrecision, vConfirm)
            End If
            Me.HitungUnbalance()
            Me.TxtRemarks.Text = Me.ObjQuickJurnal.GetDesc_ByCode(Me.cboQuickJournal.Text) + " " + Me.TxtCustCode.Text + "-" + Me.TxtCustName.Text

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

            Me.SetMulticurrency()
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

    Private Sub cboQuickJournal_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboQuickJournal.SelectedIndexChanged

    End Sub
End Class