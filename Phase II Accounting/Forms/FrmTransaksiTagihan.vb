Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiTagihan
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public QuickInput As Boolean = False
    Public ShowStopProcess As Boolean = False
    Public StopProcess As Boolean = False

    Dim DefaultPrecision As Integer = 0
    Dim baru As Boolean
    Dim ObjPP As New ClsTransaksiTagihan
    Dim ObjPPItem As New ClsTransaksiTagihanItem
    Dim buffer As String = ""
    Dim ObjCD As New ClsCekData
    Dim ObjCust As New clsMasterCustomer
    Dim ObjInv As New ClsTransaksiSalesInvoice
    Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering
    Dim SuratTagihanOtomatis As Boolean = False
    Dim StartLocation As New Point(0, 0)

    Public DirectlyPrint As Boolean = False
    Public PrinterName As String = ""


    Public Sub PrintToReal(ByVal vPrinterName As String)
        Dim ij As New FrmTransaksiTagihan_Cetak
        Try
            If Me.TxtPrimary.Text <> "-1" And Me.TxtPrimary.Text <> "" Then
                If CDbl(Me.TxtPrimary.Text) >= 0 Then
                    ObjPP.CetakTransaksiTagihan(ij.crvReport, Me.TxtPrimary.Text, Me.TxtPrimaryCust.Text, True, vPrinterName)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij = Nothing
        End Try
    End Sub

    Public Function PrintAndExport(ByVal vPrimaryAccountEmail As Double) As Boolean
        Dim ij As New FrmTransaksiTagihan_Cetak
        Try
            If Me.TxtPrimary.Text <> "-1" And Me.TxtPrimary.Text <> "" Then
                If CDbl(Me.TxtPrimary.Text) >= 0 Then
                    ObjPP.CetakTransaksiTagihan(ij.crvReport, Me.TxtPrimary.Text, Me.TxtPrimaryCust.Text, False, "", True, vPrimaryAccountEmail)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij = Nothing
        End Try
    End Function

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim ij As New FrmTransaksiTagihan_Cetak
        Try
            If Me.TxtPrimary.Text <> "-1" And Me.TxtPrimary.Text <> "" Then
                If CDbl(Me.TxtPrimary.Text) >= 0 Then
                    ObjPP.CetakTransaksiTagihan(ij.crvReport, Me.TxtPrimary.Text, Me.TxtPrimaryCust.Text)
                    ij.ShowDialog()
                End If
            Else
                Me.btnFind.PerformClick()
                If CDbl(Me.TxtPrimary.Text) >= 0 Then
                    ObjPP.CetakTransaksiTagihan(ij.crvReport, Me.TxtPrimary.Text, Me.TxtPrimaryCust.Text)
                    ij.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij = Nothing
        End Try
    End Sub

    Private Sub FrmTransaksiDeliveryMonitoring_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.ObjPP = New ClsTransaksiTagihan
            Me.baru = True

            If Me.QuickInput = False Then
                Me.KosongkanField()
                Me.TxtNoTrans.Clear()
                Me.NonAktifkanDetail()
                Me.StartLocation = New Point(Me.Location)
            End If

            If Me.ShowStopProcess = True Then
                Me.btnStopProcess.Visible = True
                Me.btnStopProcess.Enabled = True
                Me.prgAll.Visible = True
                Me.lblProgress.Visible = True
            End If
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpDate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.BackgroundWorker1.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.ObjNumbering.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO) = True Then
                    If Me.baru = False Then
                        If Me.ObjPP.TanggalTransaksi.Month = Me.dtpDate.Value.Month And Me.ObjPP.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                            'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                            Me.TxtNoTrans.Text = Me.ObjPP.Notransaksi
                        Else
                            'jika pindah bulan maka akan ganti no
                            Me.TxtNoTrans.Text = Me.ObjPP.GetNoBaru(Me.dtpDate.Value)
                        End If
                    Else
                        Me.TxtNoTrans.Text = Me.ObjPP.GetNoBaru(Me.dtpDate.Value)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimary.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboInput.Controls, Me.TxtPrimary.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimary.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimary.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.TxtPrimary.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.TxtPrimary.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Me.TxtPrimary.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Me.TxtPrimary.BackColor, False)

            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.gboInput.Controls, True, True, True, True, True, True, True, True)
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

            If Me.dgvItem.Rows.Count = 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If

            Me.llbCustAdd.Enabled = True
            Me.dtpInvoiceDueDate.Enabled = False
            Me.dtpInvoiceDate.Enabled = False
            ''Me.txtPayment.BackColor = Me.txtReadOnly.BackColor

            If Me.ObjNumbering.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.CollectionLetter) = True Then
                Me.TxtNoTrans.ReadOnly = True
                Me.TxtNoTrans.BackColor = Me.txtReadOnly.BackColor
                Me.TxtNoTrans.TabStop = False
                Me.dtpDate.Focus()
            Else
                Me.TxtNoTrans.ReadOnly = False
                Me.TxtNoTrans.BackColor = Me.TxtPrimary.BackColor
                Me.TxtNoTrans.TabStop = True
                Me.TxtNoTrans.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.gboInput.Controls, True, False, True, False, True, False, True, False)
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

            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.gboInput.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Me.BackColor, True)

            Me.llbCustAdd.Enabled = False
            Me.dtpInvoiceDueDate.Enabled = False
            Me.dtpInvoiceDate.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.KosongkanItem()
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer3.Panel1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer3.Panel2.Controls, True, True, True)
            Me.TxtPrimary.Text = "-1"
            Me.TxtNoTrans.Clear()
            Me.TxtPrimaryCust.Text = "-1"
            Me.TxtCustAddIdentifier.Text = "-1"
            Me.cboAdd.DataSource = Nothing
            Me.cboAdd.Items.Clear()

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "

            Me.dgvItem.Rows.Clear()
            Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanItem()
        Try
            ModCekData.KosongkanControl(Me.gboInput.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, True)
            Me.TxtRemarksItem.Text = ""
            Me.TxtAmount.Text = "0"
            Me.TxtCurrItem.Text = ""
            Me.TxtPrimaryInvoice.Text = "-1"
            Me.dtpInvoiceDate.Value = Now
            Me.dtpInvoiceDueDate.Value = Now
            Me.TxtInvoiceNumber.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinkeObjek()
        Try
            With ObjPP
                .Notransaksi = Me.TxtNoTrans.Text
                .TanggalTransaksi = Me.dtpDate.Value
                .Subject = Me.TxtSubject.Text
                .SignedBy = Me.TxtSignedBy.Text
                .SignedPosition = Me.TxtPosition.Text
                .Attn = Me.TxtAttn.Text
                .Identifier = Me.TxtCustAddIdentifier.Text
                .PrimaryCust = Me.TxtPrimaryCust.Text
                .Remarks = Me.TxtRemark.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinkeField()
        Try
            With ObjPP
                Me.TxtPrimary.Text = .Primary
                Me.TxtNoTrans.Text = .Notransaksi
                Me.dtpDate.Value = .TanggalTransaksi
                Me.TxtSubject.Text = .Subject
                Me.TxtRemark.Text = .Remarks
                Me.TxtCustAddIdentifier.Text = .Identifier
                Me.TxtAttn.Text = .Attn
                Me.TxtSignedBy.Text = .SignedBy
                Me.TxtPosition.Text = .SignedPosition
                Me.TxtPrimaryCust.Text = .PrimaryCust
                Me.TxtCustCode.Text = ObjCust.GetCustomerCode(Me.TxtPrimaryCust.Text)
                Me.TxtCustName.Text = ObjCust.GetCustName(Me.TxtPrimaryCust.Text)
                Me.ObjCust.IsiCboAddress(Me.cboAdd, Me.TxtPrimaryCust.Text)
                Me.cboAdd.Text = Me.ObjCust.GetKodeAddress(Me.TxtCustAddIdentifier.Text)

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim Dummy As String = ""

        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("Ref Number")
            KoleksiTextBox.Add(Me.TxtNoTrans)
            KoleksiNamaData.Add("Customer Code")
            KoleksiTextBox.Add(Me.TxtCustCode)
            KoleksiNamaData.Add("Subject")
            KoleksiTextBox.Add(Me.TxtSubject)
            KoleksiNamaData.Add("Attention")
            KoleksiTextBox.Add(Me.TxtAttn)
            KoleksiNamaData.Add("Signed By")
            KoleksiTextBox.Add(Me.TxtSignedBy)
            KoleksiNamaData.Add("Position")
            KoleksiTextBox.Add(Me.TxtPosition)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data default precision harus non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'Cek DGVItem
            If Me.dgvItem.Rows.Count <= 0 Then
                MsgBox("Please input item", MsgBoxStyle.Information, "Information")
                Me.TxtInvoiceNumber.Focus()
                Return False
            End If

            'cek perubahan nomor transaksi
            ''If Me.baru = False Then
            ''If Me.ObjPP.TanggalTransaksi.Month = Me.dtpTanggal.Value.Month And _
            ''Me.ObjPP.TanggalTransaksi.Year = Me.dtpTanggal.Value.Year Then
            ''Me.TxtNoTrans.Text = Me.ObjPP.RefNumber
            ''Else
            ''If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
            ''MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
            ''Me.dtpTanggal.Value = Me.ObjPP.TanggalTransaksi
            ''Me.TxtNoTrans.Text = Me.ObjPP.RefNumber
            ''Return False
            ''End If
            ''End If
            ''End If

            'cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.ObjPP.ApakahSudahAda(Me.TxtNoTrans.Text) = True Then
                    MsgBox("Duplication in ref number, please change your number", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtNoTrans.Focus()
                    Return False
                End If
            Else
                Dim TipeCodeLama As String = Me.ObjPP.GetNoTransaksi(Me.TxtPrimary.Text)
                If TipeCodeLama <> Me.TxtNoTrans.Text Then
                    If Me.ObjPP.ApakahSudahAda(Me.TxtNoTrans.Text) = True Then
                        MsgBox("Duplication in ref number, please change your number", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtNoTrans.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.ObjPP.GetTimestamp(Me.TxtPrimary.Text)
                If Me.ObjPP.Timestamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjPP.CariDetail(Me.ObjPP.Primary, Me.dgvItem, Me)
                        Me.SalinkeField()
                    End If
                    Return False
                End If
            End If

            If Me.QuickInput = False Then
                If Me.TxtRemark.Text = "" Then
                    MsgBox("Remarks data is empty.", MsgBoxStyle.Information, "information")
                End If
            End If

            'If Me.QuickInput = False Then
            Dummy = ObjPP.CekInvoiceNumber(Me.dgvItem, Me.TxtPrimary.Text, Me)
            If Dummy <> "" Then
                If Me.QuickInput = True Then
                    Me.Hide()
                    Me.Location = Me.StartLocation
                    Me.Show()
                End If
                If MsgBox("The following invoice number have been recorded on the following collection letter(s):" _
                + vbCrLf _
                + Dummy + "." + vbCrLf + " Continue save?", MsgBoxStyle.YesNo, "Confirmation") <> MsgBoxResult.Yes Then
                    If Me.QuickInput = True Then
                        Me.Hide()
                    End If
                    Return False
                End If
                If Me.QuickInput = True Then
                    Me.Hide()
                End If
            End If
            'End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Function

    Private Function ApakahDataItemBenar() As Boolean
        Try
            If Me.TxtCustCode.Text = "" Then
                MsgBox("Please input data customer", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If Me.TxtInvoiceNumber.Text = "" Then
                MsgBox("Please insert invoice number", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If ObjPP.CekApakahNoInvoicePunyaCustomerNya(Me.TxtCustCode.Text, _
            Me.TxtInvoiceNumber.Text) = False Then
                MsgBox("Invoice number " + Me.TxtInvoiceNumber.Text + " isn't created by customer code " + Me.TxtCustCode.Text + _
                ", please adjusted invoice by customer code ", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub txtRemarks_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.Enter
        Try
            Me.buffer = Me.TxtRemark.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtAmount.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtAmount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtAmount.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtAmount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtRemark.ReadOnly = False Then
                    Me.TxtRemark.BackColor = Me.BackColorAll
                Else
                    TxtRemark.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtRemark.ReadOnly = False Then
                    Me.TxtRemark.BackColor = Me.TxtPrimary.BackColor
                Else
                    TxtRemark.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarkItem_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarksItem.GotFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtRemarksItem.ReadOnly = False Then
                    Me.TxtRemarksItem.BackColor = Me.BackColorAll
                Else
                    TxtRemarksItem.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarkItem_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarksItem.LostFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtRemarksItem.ReadOnly = False Then
                    Me.TxtRemarksItem.BackColor = Me.TxtPrimary.BackColor
                Else
                    TxtRemarksItem.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txtkodetujuan_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNoTrans.GotFocus
        Try
            If Me.TxtNoTrans.Enabled = True Then
                If Me.TxtNoTrans.ReadOnly = False Then
                    Me.TxtNoTrans.BackColor = Me.BackColorAll
                Else
                    TxtNoTrans.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txtkodetujuan_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNoTrans.LostFocus
        Try
            If Me.TxtNoTrans.Enabled = True Then
                If Me.TxtNoTrans.ReadOnly = False Then
                    Me.TxtNoTrans.BackColor = Me.TxtPrimary.BackColor
                Else
                    TxtNoTrans.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txtpayment_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSubject.GotFocus
        Try
            If Me.TxtSubject.Enabled = True Then
                If Me.TxtSubject.ReadOnly = False Then
                    Me.TxtSubject.BackColor = Me.BackColorAll
                Else
                    TxtSubject.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txtpayment_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSubject.LostFocus
        Try
            If Me.TxtSubject.Enabled = True Then
                If Me.TxtSubject.ReadOnly = False Then
                    Me.TxtSubject.BackColor = Me.TxtPrimary.BackColor
                Else
                    TxtSubject.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '' Focus No Transaksi
    Private Sub TxtNama_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAttn.GotFocus
        Try
            If Me.TxtAttn.Enabled = True Then
                If Me.TxtAttn.ReadOnly = False Then
                    Me.TxtAttn.BackColor = Me.BackColorAll
                Else
                    TxtAttn.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNama_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAttn.LostFocus
        Try
            If Me.TxtAttn.Enabled = True Then
                If Me.TxtAttn.ReadOnly = False Then
                    Me.TxtAttn.BackColor = Me.TxtPrimary.BackColor
                Else
                    TxtAttn.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txtamount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustCode.GotFocus
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

    Private Sub Txtamount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustCode.LostFocus
        Try
            If Me.TxtCustCode.Enabled = True Then
                If Me.TxtCustCode.ReadOnly = False Then
                    Me.TxtCustCode.BackColor = Me.TxtPrimary.BackColor
                Else
                    TxtCustCode.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txtdestination_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSignedBy.GotFocus
        Try
            If Me.TxtSignedBy.Enabled = True Then
                If Me.TxtSignedBy.ReadOnly = False Then
                    Me.TxtSignedBy.BackColor = Me.BackColorAll
                Else
                    TxtSignedBy.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txtdestination_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSignedBy.LostFocus
        Try
            If Me.TxtSignedBy.Enabled = True Then
                If Me.TxtSignedBy.ReadOnly = False Then
                    Me.TxtSignedBy.BackColor = Me.TxtPrimary.BackColor
                Else
                    TxtSignedBy.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txtdocumentnumber_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPosition.GotFocus
        Try
            If Me.TxtPosition.Enabled = True Then
                If Me.TxtPosition.ReadOnly = False Then
                    Me.TxtPosition.BackColor = Me.BackColorAll
                Else
                    TxtPosition.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txtdocumentnumber_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPosition.LostFocus
        Try
            If Me.TxtPosition.Enabled = True Then
                If Me.TxtPosition.ReadOnly = False Then
                    Me.TxtPosition.BackColor = Me.TxtPrimary.BackColor
                Else
                    TxtPosition.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txtcurramount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtInvoiceNumber.GotFocus
        Try
            If Me.TxtInvoiceNumber.Enabled = True Then
                If Me.TxtInvoiceNumber.ReadOnly = False Then
                    Me.TxtInvoiceNumber.BackColor = Me.BackColorAll
                Else
                    TxtInvoiceNumber.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txtcurramount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtInvoiceNumber.LostFocus
        Try
            If Me.TxtInvoiceNumber.Enabled = True Then
                If Me.TxtInvoiceNumber.ReadOnly = False Then
                    Me.TxtInvoiceNumber.BackColor = Me.TxtPrimary.BackColor
                Else
                    TxtInvoiceNumber.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''Button dan event
    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmTransaksiTagihan_Cari
        Try
            ij.BackColorAll = Me.BackColorAll
            ij.ShowDialog()
            If Not ij.PrimaryPP = -1 Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                If Me.ObjPP.CariDetail(ij.PrimaryPP, Me.dgvItem, Me) = True Then
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

    Public Sub NewTagihan()
        Try
            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()

            'ga standar
            Me.TxtCustCode.ReadOnly = False
            Me.TxtCustCode.BackColor = Me.TxtPrimaryCust.BackColor
            Me.btnFindCust.Enabled = True
            Me.TxtCustCode.TabStop = True

            Me.ObjPP.ResetObject(Me.dgvItem)
            Me.TxtNoTrans.Text = Me.ObjPP.GetNoBaru(Me.dtpDate.Value)

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

            Me.SuratTagihanOtomatis = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()

            'ga standar
            Me.TxtCustCode.ReadOnly = False
            Me.TxtCustCode.BackColor = Me.TxtPrimaryCust.BackColor
            Me.btnFindCust.Enabled = True
            Me.TxtCustCode.TabStop = True

            Me.ObjPP.ResetObject(Me.dgvItem)
            Me.TxtNoTrans.Text = Me.ObjPP.GetNoBaru(Me.dtpDate.Value)

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
            If Me.TxtNoTrans.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtNoTrans.Text <> "" Then
                If Me.ObjPP.ApakahBolehEdit(Me.ObjPP.Primary, UserCode, True, True) = True Then
                    Me.baru = False
                    Me.ObjPP.CariDetail(Me.ObjPP.Primary, Me.dgvItem, Me)
                    Me.SalinkeField()
                    Me.AktifkanDetail()

                    'ga standar
                    If Me.dgvItem.RowCount > 0 Then
                        Me.TxtCustCode.ReadOnly = True
                        Me.TxtCustCode.BackColor = Me.txtReadOnly.BackColor
                        Me.btnFindCust.Enabled = False
                        Me.TxtCustCode.TabStop = False
                    End If

                    Me.dgvItem.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtNoTrans.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtNoTrans.Text <> "" Then
                If Me.ObjPP.ApakahBolehHapus(Me.TxtPrimary.Text, True, UserCode, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjPP.Hapus(Me.TxtPrimary.Text, UserCode, True) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                            Me.NonAktifkanDetail()
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
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinkeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjPP.SimpanBaru(Me.dgvItem, Me.UserCode, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinkeField()
                            Me.NonAktifkanDetail()
                            If Me.DirectlyPrint = True Then
                                Me.PrintToReal(Me.PrinterName)
                                Me.Close()
                            End If
                            Me.btnFind.Focus()
                            If Me.SuratTagihanOtomatis = True Then
                                Me.Close()
                            End If
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ObjPP.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinkeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
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

    Public Sub SimpanBaru()
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinkeObjek()
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                If Me.ObjPP.SimpanBaru(Me.dgvItem, Me.UserCode, Me) Then
                    Me.SalinkeField()
                    Me.NonAktifkanDetail()
                    If Me.DirectlyPrint = True Then
                        Me.PrintToReal(Me.PrinterName)
                        Me.Close()
                    End If
                    Me.btnFind.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtRemark.Text = "" And Me.dgvItem.RowCount = 0 And _
            Me.TxtNoTrans.Text = "" And Me.TxtRemarksItem.Text = "" Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
                If Me.SuratTagihanOtomatis = True Then
                    Me.Close()
                End If
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Me.KosongkanItem()
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.ObjPP.CariDetail(Me.TxtPrimary.Text, Me.dgvItem, Me)
                        Me.SalinkeField()
                    End If

                    If Me.SuratTagihanOtomatis = True Then
                        Me.Close()
                    End If
                Else
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
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.BringToFront()
                Dim hasil As MsgBoxResult

                If Me.QuickInput = True Then
                    hasil = MsgBoxResult.No
                Else
                    hasil = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                End If

                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinkeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjPP.SimpanBaru(Me.dgvItem, Me.UserCode, Me) = True Then
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
                            If Me.ObjPP.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me) = True Then
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''Pengaturan DGV Item
    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProcess_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Try
            Me.dgvItem.TabStop = True
            If Me.TxtRemark.Enabled = True Then
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProcess_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItem.RowsRemoved
        Try
            If Me.dgvItem.Rows.Count = 0 Then
                Me.dgvItem.TabStop = False
                If Me.TxtRemark.Enabled = True Then
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
            If Me.TxtRemark.Enabled = True Then
                If Me.dgvItem.RowCount = 0 Then
                    Me.TxtCustCode.ReadOnly = False
                    Me.TxtCustCode.BackColor = Me.TxtPrimaryCust.BackColor
                    Me.btnFindCust.Enabled = True
                    Me.TxtCustCode.TabStop = True
                End If
            End If
            Me.ObjCD.SetNumber(Me.dgvItem)
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
                    Me.ObjCD.SetNumber(Me.dgvItem)
                    If row.Displayed = False Then
                        Me.ShowRowProcess(row, True)
                    End If
                End If

                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
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
                    Me.ObjCD.SetNumber(Me.dgvItem)
                    Me.ShowRowProcess(row, False)
                End If

                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProcess_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.SelectionChanged
        Try
            If Me.TxtRemark.Enabled = True Then
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
                        Me.cboAdd.Focus()
                        Me.ObjCust.IsiCboAddress(Me.cboAdd, Me.TxtPrimaryCust.Text)
                        If Me.cboAdd.Items.Count > 0 Then
                            Me.cboAdd.SelectedIndex = 0
                        End If
                        IsiDGVOtomatis()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.btnFindCust.PerformClick()
                    End If
                Else
                    Me.btnFindCust.PerformClick()
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
            If Me.TxtCustCode.Enabled = False Or Me.TxtCustCode.ReadOnly = True Then
                Exit Sub
            End If
            If Me.TxtCustCode.Text <> "" Then
                ''If Me.cboTipeTujuan.SelectedIndex = 0 Then
                If ObjCust.ApakahSudahAda(Me.TxtCustCode.Text) = True Then
                    Me.TxtPrimaryCust.Text = ObjCust.GetPrimary(Me.TxtCustCode.Text)
                    Me.TxtCustCode.Text = ObjCust.GetCustomerCode(Me.TxtPrimaryCust.Text)
                    Me.TxtCustName.Text = ObjCust.GetCustName(Me.TxtPrimaryCust.Text)
                    Me.ObjCust.IsiCboAddress(Me.cboAdd, Me.TxtPrimaryCust.Text)
                    If Me.cboAdd.Items.Count > 0 Then
                        Me.cboAdd.SelectedIndex = 0
                    End If
                    IsiDGVOtomatis()
                Else
                    If Me.TxtCustCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtCustCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.TxtPrimaryCust.Text = i.PrimaryCustomer
                        Me.TxtCustCode.Text = ObjCust.GetCustomerCode(Me.TxtPrimaryCust.Text)
                        Me.TxtCustName.Text = ObjCust.GetCustName(Me.TxtPrimaryCust.Text)
                        Me.ObjCust.IsiCboAddress(Me.cboAdd, Me.TxtPrimaryCust.Text)
                        If Me.cboAdd.Items.Count > 0 Then
                            Me.cboAdd.SelectedIndex = 0
                        End If
                        IsiDGVOtomatis()
                    Else
                        Me.TxtCustCode.Clear()
                        Me.TxtCustName.Clear()
                        Me.TxtPrimaryCust.Text = "-1"
                        Me.TxtCustCode.Focus()
                        Me.cboAdd.DataSource = Nothing
                        Me.cboAdd.Items.Clear()
                    End If
                End If
                ''End If
            Else
                Me.TxtCustCode.Clear()
                Me.TxtCustName.Clear()
                Me.TxtPrimaryCust.Text = "-1"
                Me.cboAdd.DataSource = Nothing
                Me.cboAdd.Items.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindCust_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles btnFindCust.Click
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
                Me.ObjCust.IsiCboAddress(Me.cboAdd, Me.TxtPrimaryCust.Text)
                Me.cboAdd.Focus()
                If Me.cboAdd.Items.Count > 0 Then
                    Me.cboAdd.SelectedIndex = 0
                End If
                If Me.dgvItem.RowCount = 0 Then
                    IsiDGVOtomatis()
                End If
            Else
                Me.TxtCustCode.Clear()
                Me.TxtCustName.Clear()
                Me.TxtPrimaryCust.Text = "-1"
                Me.TxtCustCode.Focus()
                Me.cboAdd.DataSource = Nothing
                Me.cboAdd.Items.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtInvoiceNumber_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles TxtInvoiceNumber.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtInvoiceNumber.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjInv.ApakahSudahAda(Me.TxtInvoiceNumber.Text) = True Then
                        Me.TxtPrimaryInvoice.Text = ObjInv.GetPrimary(Me.TxtInvoiceNumber.Text)
                        Me.TxtInvoiceNumber.Text = ObjInv.GetInvoiceNumber(Me.TxtPrimaryInvoice.Text)
                        Me.TxtCurrItem.Text = ObjInv.GetCurrency(Me.TxtPrimaryInvoice.Text)
                        Me.TxtAmount.Text = ObjCD.DigitGrouping(ObjInv.GetAmount(Me.TxtPrimaryInvoice.Text))
                        Me.dtpInvoiceDate.Value = ObjInv.GetInvoiceDate(Me.TxtPrimaryInvoice.Text)
                        Me.dtpInvoiceDueDate.Value = ObjInv.GetDueDate(Me.TxtPrimaryInvoice.Text)
                        Me.TxtRemarksItem.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.btnFindInvoice.PerformClick()
                    End If
                Else
                    Me.btnFindInvoice.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtInvoiceNumber_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtInvoiceNumber.Leave
        Dim ij As New FrmTransaksiTagihan_CariInvoice
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtInvoiceNumber.Text <> "" Then
                ''If Me.cboTipeTujuan.SelectedIndex = 0 Then
                If ObjInv.ApakahSudahAda(Me.TxtInvoiceNumber.Text) = True Then
                    Me.TxtPrimaryInvoice.Text = ObjInv.GetPrimary(Me.TxtInvoiceNumber.Text)
                    Me.TxtInvoiceNumber.Text = ObjInv.GetInvoiceNumber(Me.TxtPrimaryInvoice.Text)
                    Me.TxtCurrItem.Text = ObjInv.GetCurrency(Me.TxtPrimaryInvoice.Text)
                    Me.TxtAmount.Text = ObjCD.DigitGrouping(ObjInv.GetAmount(Me.TxtPrimaryInvoice.Text))
                    Me.dtpInvoiceDate.Value = ObjInv.GetInvoiceDate(Me.TxtPrimaryInvoice.Text)
                    Me.dtpInvoiceDueDate.Value = ObjInv.GetDueDate(Me.TxtPrimaryInvoice.Text)
                Else
                    If Me.TxtInvoiceNumber.Text <> "" Then
                        ij.txtCari.Text = Me.TxtInvoiceNumber.Text
                    End If

                    ij.ShowDialog()
                    If ij.PrimaryTransaksi > 0 Then
                        Me.TxtPrimaryInvoice.Text = ij.PrimaryTransaksi
                        Me.TxtInvoiceNumber.Text = ObjInv.GetInvoiceNumber(Me.TxtPrimaryInvoice.Text)
                        Me.TxtCurrItem.Text = ObjInv.GetCurrency(Me.TxtPrimaryInvoice.Text)
                        Me.TxtAmount.Text = ObjCD.DigitGrouping(ObjInv.GetAmount(Me.TxtPrimaryInvoice.Text))
                        Me.dtpInvoiceDate.Value = ObjInv.GetInvoiceDate(Me.TxtPrimaryInvoice.Text)
                        Me.dtpInvoiceDueDate.Value = ObjInv.GetDueDate(Me.TxtPrimaryInvoice.Text)
                    Else
                        Me.TxtInvoiceNumber.Clear()
                        Me.TxtCurrItem.Clear()
                        Me.TxtAmount.Text = "0"
                        Me.TxtPrimaryInvoice.Text = "-1"
                        Me.dtpInvoiceDate.Value = Now
                        Me.dtpInvoiceDueDate.Value = Now
                        Me.TxtInvoiceNumber.Focus()
                    End If
                End If
                ''End If
            Else
                Me.TxtInvoiceNumber.Clear()
                Me.TxtCurrItem.Clear()
                Me.TxtAmount.Text = "0"
                Me.dtpInvoiceDate.Value = Now
                Me.dtpInvoiceDueDate.Value = Now
                Me.TxtPrimaryInvoice.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub BtnFindInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles btnFindInvoice.Click
        Dim ij As New FrmTransaksiTagihan_CariInvoice
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtInvoiceNumber.Text <> "" Then
                ij.txtCari.Text = Me.TxtInvoiceNumber.Text
            End If

            ij.ShowDialog()
            If ij.PrimaryTransaksi > 0 Then
                Me.TxtPrimaryInvoice.Text = ij.PrimaryTransaksi
                Me.TxtInvoiceNumber.Text = ObjInv.GetInvoiceNumber(Me.TxtPrimaryInvoice.Text)
                Me.TxtCurrItem.Text = ObjInv.GetCurrency(Me.TxtPrimaryInvoice.Text)
                Me.TxtAmount.Text = ObjCD.DigitGrouping(ObjInv.GetAmount(Me.TxtPrimaryInvoice.Text))
                Me.dtpInvoiceDate.Value = ObjInv.GetInvoiceDate(Me.TxtPrimaryInvoice.Text)
                Me.dtpInvoiceDueDate.Value = ObjInv.GetDueDate(Me.TxtPrimaryInvoice.Text)
                Me.TxtRemarksItem.Focus()
            Else
                Me.TxtInvoiceNumber.Clear()
                Me.TxtCurrItem.Clear()
                Me.TxtAmount.Text = "0"
                Me.TxtPrimaryInvoice.Text = "-1"
                Me.dtpInvoiceDate.Value = Now
                Me.dtpInvoiceDueDate.Value = Now
                Me.TxtInvoiceNumber.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    '' Bagian susah
    Private Sub btnInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnINput.Click
        Dim Index1 As Integer = -1
        Dim row As DataGridViewRow
        Dim deliver As String = ""

        Try
            If Me.ApakahDataItemBenar = True Then
                Index1 = Me.GetIndex(Me.TxtInvoiceNumber.Text)

                'cari data yang sama persis
                If Index1 <> -1 Then
                    'tambahkan nilai qty
                    ''row = Me.dgvItem.Rows(Index1)
                    ''row.Cells(Me.ColQty.Index).Value = ObjCd.DigitGrouping(CDbl(row.Cells(Me.ColQty.Index).Value) + _
                    ''CDbl(Me.txtQtyPerBatch.Text))
                    MsgBox("Duplication data item, please change your data item", _
                    MsgBoxStyle.Information, "Information")
                Else
                    'tambahkan baris baru
                    Me.dgvItem.Rows.Add()
                    row = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                    row.Cells(Me.ColInvoiceNumber.Index).Value = Me.TxtInvoiceNumber.Text
                    row.Cells(Me.ColPrimaryInvoice.Index).Value = Me.TxtPrimaryInvoice.Text
                    row.Cells(Me.ColInvoiceDate.Index).Value = Format(Me.dtpInvoiceDate.Value, "dd MMM yyyy")
                    row.Cells(Me.ColDueDate.Index).Value = Format(Me.dtpInvoiceDueDate.Value, "dd MMM yyyy")
                    row.Cells(Me.ColAmount.Index).Value = ObjCD.DigitGrouping(Me.TxtAmount.Text)
                    row.Cells(Me.ColCurrency.Index).Value = Me.TxtCurrItem.Text
                    If String.IsNullOrEmpty(Me.TxtRemarksItem.Text) = True Then
                        row.Cells(Me.ColRemarks.Index).Value = "-"
                    Else
                        row.Cells(Me.ColRemarks.Index).Value = Me.TxtRemarksItem.Text
                    End If

                    'set ulang nomor
                    Me.ObjCD.SetNumber(Me.dgvItem)
                End If

                'kosongkan gbo item
                Me.KosongkanItem()

                'fokuskan ke baris terakhir
                While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While

                If Me.dgvItem.RowCount > 0 Then
                    Me.TxtCustCode.ReadOnly = True
                    Me.TxtCustCode.BackColor = Me.txtReadOnly.BackColor
                    Me.btnFindCust.Enabled = False
                    Me.TxtCustCode.TabStop = False
                End If

                Me.TxtInvoiceNumber.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndex(ByVal vInvoiceNumber As String) As Integer
        Dim row As DataGridViewRow
        Try
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColInvoiceNumber.Index).Value = vInvoiceNumber Then
                    Return row.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvItem_CellBeginEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColRemarks.Index Then
                Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim Index As Integer = -1
        Try
            Dim row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
            'pastikan ada isinya
            With row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If e.ColumnIndex = Me.ColAmount.Index Or e.ColumnIndex = Me.ColAmount.Index Then
                        If IsNumeric(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            If IsNumeric(Me.buffer) = True Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Else
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = 0
                            End If
                        End If
                    End If
                    If e.ColumnIndex <> Me.ColRemarks.Index Then
                        'kalo bukan data remarks maka perlu dicek kebenaran data hasil edit
                        If e.ColumnIndex = Me.ColAmount.Index Then
                            If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, _
                            e.RowIndex).Value) = False Then
                                MsgBox("Please enter positive number", MsgBoxStyle.Information, _
                                "Information")
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.buffer)
                            Else
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString)
                            End If
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboAdd_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboAdd.SelectedIndexChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.cboAdd.SelectedIndex <> -1 And IsNumeric(Me.TxtPrimaryCust.Text) = True Then
                    Me.TxtCustAddIdentifier.Text = Me.ObjCust.GetIdentifierAdd(Me.TxtPrimaryCust.Text, _
                    Me.cboAdd.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbCustAdd_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCustAdd.LinkClicked
        Try
            If Me.TxtPrimaryCust.Text = "" Or Me.TxtPrimaryCust.Text = "-1" Then
                MsgBox("Please specify customer first.", MsgBoxStyle.Information, "Information")
                Me.TxtCustCode.Focus()
                Exit Sub
            End If
            Dim ij As New Core_Module.FrmDetailAddress
            ij.vAddIdentifier = Me.TxtCustAddIdentifier.Text
            ij.PrimaryCust = Me.TxtPrimaryCust.Text
            ij.btnOK.BringToFront()
            ij.AcceptButton = ij.btnOK
            ij.btnExit.TabStop = False
            ij.ShowDialog()
            If ij.DialogResult = Windows.Forms.DialogResult.OK Then
                Me.ObjCust.IsiCboAddress(Me.cboAdd, Me.TxtPrimaryCust.Text)
                Me.cboAdd.Text = Me.ObjCust.GetKodeAddress(ij.vAddIdentifierResult)
            End If
            ij.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub IsiDGVOtomatis()
        Dim i As MsgBoxResult

        Try
            If Me.QuickInput = False Then
                i = MsgBox("Insert data from database?", MsgBoxStyle.YesNo, "Confirm")
            Else
                i = MsgBoxResult.Yes
            End If

            If i = MsgBoxResult.Yes Then
                Me.ObjPP.IsiDGVBelumLunas(Me.dtpDate.Value, Me.dgvItem, Me.TxtPrimaryCust.Text, Me, Me.ObjCD)
            End If

            If Me.dgvItem.RowCount > 0 Then
                Me.TxtCustCode.ReadOnly = True
                Me.TxtCustCode.BackColor = Me.txtReadOnly.BackColor
                Me.btnFindCust.Enabled = False
                Me.TxtCustCode.TabStop = False
            Else
                Me.TxtCustCode.ReadOnly = False
                Me.TxtCustCode.BackColor = Me.TxtPrimaryCust.BackColor
                Me.btnFindCust.Enabled = True
                Me.TxtCustCode.TabStop = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStopProcess_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnStopProcess.Click
        Try
            If MsgBox("Stop creating collection letter?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.StopProcess = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class