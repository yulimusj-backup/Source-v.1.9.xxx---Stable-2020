Option Explicit On
Option Strict Off

Imports Core_Module
Public Class FrmTransaksiPurchaseBalanceManual

    Dim objImport As New CustomSMESS.ClsTrans_ImportAPBalance
    Public UserCode As String = ""
    Public BackCOlorAll As Color = Color.LightSeaGreen

    Dim baru As Boolean
    Dim ObjProcess As ClsTransaksiPurchaseBalance
    Dim buffer As String = ""
    Dim buffer2 As String = ""
    Dim ObjCD As New ClsCekData
    Dim ObjCust As New clsMasterSupplier
    Dim ObjClosing As New ClsTransaksiClosing
    Dim Default_Prec As Integer
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim TBLCurr As New dseMasterCurrency.tblmaster_currencyDataTable
    Dim DeleteHistorical As Boolean = False
    Dim SedangEvaluasi As Boolean = False

    Dim DseAR As New DseMultiAccountTableAdapters.tblisi_cbo_ar_akunTableAdapter
    Dim tblAR As New DseMultiAccount.tblisi_cbo_ar_akunDataTable
    Dim rowAR As DseMultiAccount.tblisi_cbo_ar_akunRow

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryTransaksi.BackColor, False)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, True, True)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.SetExchangeRate()
            Me.cboAccount.Enabled = True
            ' Me.cboAccount.SelectedIndex = 0
            Me.llbImport.Enabled = False

            If Me.baru = True Then
                IsiCBOArAkunByTipePerkiraan1()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, True, False)
            Me.CancelButton = Me.btnExit
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, True)
            Me.cboAccount.Enabled = False
            'Me.cboAccount.SelectedIndex = 0
            Me.llbImport.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Dim dv As DataRowView
        Dim row As dseMasterCurrency.tblmaster_currencyRow
        Dim i As Integer = 0

        Try
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            Me.TxtPrimaryTransaksi.Text = "-1"
            Me.TxtPrimarySupplier.Text = "-1"
            Me.TxtSales.Text = Me.ObjCD.DigitGrouping(0, Me.Default_Prec)
            Me.TxtVAT.Text = Me.ObjCD.DigitGrouping(0, Me.Default_Prec)
            Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(0, Me.Default_Prec)
            Me.TxtPaid.Text = Me.ObjCD.DigitGrouping(0, Me.Default_Prec)
            Me.TxtBalance.Text = Me.ObjCD.DigitGrouping(0, Me.Default_Prec)
            Me.TxtDutyStamp.Text = Me.ObjCD.DigitGrouping(0, Me.Default_Prec)
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.ObjCur.IsiCBO(Me.cboCurrency, Me.TBLCurr)
            Me.cboCurrency.SelectedIndex = -1
            For Each dv In Me.TBLCurr.DefaultView
                row = dv.Row
                If row.Primary = Core_Module.PrimaryBaseCurrency Then
                    Me.cboCurrency.SelectedIndex = i
                    Exit For
                End If
                i += 1
            Next
            Me.TxtPrimaryCurrency.Text = Core_Module.PrimaryBaseCurrency
            Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping("1", Core_Module.DefPrecExchangeRate)
            Me.cboPrimaryAR.DataSource = Nothing
            Me.cboAccount.DataSource = Nothing
            Me.txtPrimaryAccount.Text = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim ConvertionDate As Date

        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("Supplier Code")
            KoleksiTextBox.Add(Me.TxtSuppCode)
            KoleksiNamaData.Add("Purchase Return Number")
            KoleksiTextBox.Add(Me.TxtInvoiceNumber)


            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data default precision harus non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'cek yang tidak boleh = nol
            'If Me.TxtSales.Text = "0" Or Me.TxtSales.Text = "" Then
            '    MsgBox("Please input purchase return", MsgBoxStyle.Information, "Information")
            '    Me.TxtSales.Focus()
            '    Return False
            'End If
            'If Me.TxtVAT.Text = "0" Or Me.TxtVAT.Text = "" Then
            '    If MsgBox("VAT=0?", MsgBoxStyle.YesNo, "Warning") <> MsgBoxResult.Yes Then
            '        Me.TxtVAT.Focus()
            '        Return False
            '    End If
            'End If

            'cek due date
            If Me.Dtpduedate.Value.Date < Me.dtpInvoiceDate.Value.Date Then
                MsgBox("Please input due date <= Purchase return date", MsgBoxStyle.Information, "Information")
                Me.Dtpduedate.Focus()
                Return False
            End If

            'Cek perhitungan balance
            If CDec(Me.TxtTotal.Text) < CDec(Me.TxtPaid.Text) Then
                MsgBox("Please input data paid <= total", MsgBoxStyle.Information, "information")
                Me.TxtPaid.Focus()
                Me.TxtPaid.Text = Me.buffer
                Me.TxtBalance.Text = Me.buffer2
                Return False
            End If

            If Not IsNumeric(Me.txtPrimaryAccount.Text) OrElse Me.txtPrimaryAccount.Text = -1 Then
                MsgBox("Please choose AP Account", MsgBoxStyle.Information, "information")
                Me.cboAccount.Focus()
                Return False
            End If

            'cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.ObjProcess.ApakahSudahAda(Me.TxtInvoiceNumber.Text) = True Then
                    MsgBox("Duplication in Purchase Return number, please change your number", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtInvoiceNumber.Focus()
                    Return False
                End If
            Else
                Dim TipeCodeLama As String = Me.ObjProcess.GetAOPNumberByPrimary(Me.TxtPrimaryTransaksi.Text)
                If TipeCodeLama <> Me.TxtInvoiceNumber.Text Then
                    If Me.ObjProcess.ApakahSudahAda(Me.TxtInvoiceNumber.Text) = True Then
                        MsgBox("Duplication in Purchase Return number, please change your number", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtInvoiceNumber.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.ObjProcess.GetTimeStamp(Me.TxtPrimaryTransaksi.Text)
                If Me.ObjProcess.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjProcess.CariDetail(Me.ObjProcess.Primary)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'kalo tanggal transaksi sama atau lebih besar dari convertion date maka ditolak
            ConvertionDate = Me.ObjClosing.GetConvertionDate
            If ConvertionDate <= Me.dtpInvoiceDate.Value Then
                MsgBox("The convertion date is " + Format(ConvertionDate, "dd MMM yyyy") + _
                " this transaction is for historical data only. Every transaction must be dated before convertion date.", _
                MsgBoxStyle.Information, "Information")
                Me.dtpInvoiceDate.Focus()
                Return False
            End If

            'balance dan exchange rate harus >0
            If CDec(Me.TxtBalance.Text) <= 0 Then
                MsgBox("Please input balance with positive numbers.", MsgBoxStyle.Information, "Information")
                Me.TxtBalance.Focus()
                Return False
            End If
            If CDec(Me.TxtExchangeRate.Text) <= 0 Then
                MsgBox("Please input exchange rate with positive numbers.", MsgBoxStyle.Information, "Information")
                Me.TxtBalance.Focus()
                Return False
            End If

            '2024-04-23 Ilham, balance tidak boleh lebih besar dari 99 Miliar
            If CDec(Me.TxtBalance.Text) > 99000000000 Then
                MsgBox("Please input balance less than 99 Billions.", MsgBoxStyle.Information, "Information")
                Me.TxtBalance.Focus()
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

    Private Sub SalinKeObjek()
        Try
            With Me.ObjProcess
                .Primary_Supplier = Me.TxtPrimarySupplier.Text
                .AOPNumber = Me.TxtInvoiceNumber.Text
                .AOPDate = Me.dtpInvoiceDate.Value
                .Purchase = Me.TxtSales.Text
                .VAT = Me.TxtVAT.Text
                .Total = Me.TxtTotal.Text
                .DutyStamp = Me.TxtDutyStamp.Text
                .DueDate = Me.Dtpduedate.Value
                .Paid = Me.TxtPaid.Text
                .Balance = Me.TxtBalance.Text
                .Remarks = Me.TxtRemarks.Text
                .PrimaryCurrency = Me.TxtPrimaryCurrency.Text
                .ExchangeRate = Me.TxtExchangeRate.Text
                .PrimaryAP = Me.txtPrimaryAccount.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Dim dv As DataRowView
        Dim row As dseMasterCurrency.tblmaster_currencyRow
        Dim i As Integer = 0

        Try
            With Me.ObjProcess
                Me.TxtPrimaryTransaksi.Text = .Primary
                Me.TxtPrimarySupplier.Text = .Primary_Supplier
                Me.TxtSuppCode.Text = ObjCust.GetSupplierCode(Me.TxtPrimarySupplier.Text)
                Me.TxtSuppName.Text = ObjCust.GetSupplierName(Me.TxtPrimarySupplier.Text)
                Me.TxtInvoiceNumber.Text = .AOPNumber
                Me.dtpInvoiceDate.Value = .AOPDate
                Me.TxtRemarks.Text = .Remarks
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
                Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping(.ExchangeRate, Core_Module.DefPrecExchangeRate)
                Me.ObjCur.IsiCBO(Me.cboCurrency, Me.TBLCurr)
                Me.cboCurrency.SelectedIndex = -1
                For Each dv In Me.TBLCurr.DefaultView
                    row = dv.Row
                    If row.Primary = .PrimaryCurrency Then
                        Me.cboCurrency.SelectedIndex = i
                        Exit For
                    End If
                    i += 1
                Next
                Me.TxtPrimaryCurrency.Text = .PrimaryCurrency

                Me.Default_Prec = Me.ObjCur.GetDefaultPrecission(.PrimaryCurrency)
                Me.TxtSales.Text = ObjCD.DigitGrouping(.Purchase, Me.Default_Prec)
                Me.TxtVAT.Text = ObjCD.DigitGrouping(.VAT, Me.Default_Prec)
                Me.TxtDutyStamp.Text = Me.ObjCD.DigitGrouping(.DutyStamp, Me.Default_Prec)
                Me.TxtTotal.Text = ObjCD.DigitGrouping(.Total, Me.Default_Prec)
                Dtpduedate.Value = .DueDate
                Me.TxtPaid.Text = ObjCD.DigitGrouping(.Paid, Me.Default_Prec)
                Me.TxtBalance.Text = ObjCD.DigitGrouping(.Balance, Me.Default_Prec)

                Me.IsiCBOArAkunByTipePerkiraan1()


                If .PrimaryAP > -1 Then
                    Me.txtPrimaryAccount.Text = .PrimaryAP
                    Me.cboPrimaryAR.Text = .PrimaryAP
                Else
                    Me.cboAccount.DataSource = Nothing
                    Me.cboPrimaryAR.DataSource = Nothing
                    Me.txtPrimaryAccount.Text = -1
                End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.ObjProcess = New ClsTransaksiPurchaseBalance
            Me.baru = True
            Me.KosongkanField()
            Me.NonAktifkanDetail()

            Me.Default_Prec = Me.ObjCur.GetDecimal_DefaultCurrency

            'isikan cbo currency
            Me.ObjCur.IsiCBO(Me.cboCurrency, Me.TBLCurr)

            ''@rosadinnn
            Me.llbImport.Visible = False
            'Me.LinkLabel1.Visible = False
            Me.ProgressBar1.Visible = False
            If Me.objImport.ApakahBisaImport = True Then
                Me.llbImport.Visible = True
                'Me.LinkLabel1.Visible = True
                Me.ProgressBar1.Visible = True
            End If
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjProcess.SimpanBaru(Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.NonAktifkanDetail()
                            Me.SalinKeField()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ObjProcess.SimpanPerubahan(Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.NonAktifkanDetail()
                            Me.SalinKeField()
                            Me.btnEdit.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Not Me.TxtInvoiceNumber.Text = "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Me.NonAktifkanDetail()

                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.KosongkanField()
            End If
            Me.NonAktifkanDetail()

            Me.btnNew.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmTransaksiPurchaseBalanceManual_Cari
        Try
            ij.BackColorAll = Me.BackCOlorAll
            ij.ShowDialog()
            If Not ij.mPrimaryDept = -1 Then
                If Me.ObjProcess.CariDetail(ij.mPrimaryDept) = True Then
                    Me.SalinKeField()
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

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Me.baru = True
        Me.KosongkanField()
        Me.AktifkanDetail()
        Me.TxtSuppCode.Focus()
        Me.TxtPrimaryTransaksi.Text = "-1"
        Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
        Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
        Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtInvoiceNumber.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtInvoiceNumber.Text <> "" Then
                If Me.ObjProcess.ApakahBolehEdit(Me.TxtPrimaryTransaksi.Text, Me.UserCode) = True Then
                    Me.baru = False
                    Me.AktifkanDetail()
                    Me.TxtSuppCode.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtInvoiceNumber.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtInvoiceNumber.Text <> "" Then
                'khusus untuk IDR, tidak bisa dihapus, karena digunakan oleh sistem
                If Me.ObjProcess.ApakahBolehHapus(Me.TxtPrimaryTransaksi.Text, Me.UserCode) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ObjProcess.Hapus(Me.TxtPrimaryTransaksi.Text, Me.UserCode) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtInvoiceNumber.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjProcess.SimpanBaru(Me.UserCode) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjProcess.SimpanPerubahan(Me.UserCode) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
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

    Private Sub txtRemarks_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.Enter
        Try
            Me.buffer = Me.TxtRemarks.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRecieved_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaid.Enter
        Try
            Me.buffer = Me.TxtPaid.Text
            Me.buffer2 = Me.TxtBalance.Text

            Me.ObjCD.Number_Enter(Me.TxtPaid, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub txtPrimaryAccount_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrimaryAccount.TextChanged
        Dim PrimaryARAkun As Double = -1

        Try
            If IsNumeric(Me.txtPrimaryAccount.Text) Then
                PrimaryARAkun = CDbl(Me.txtPrimaryAccount.Text)
            End If

            If Me.tblAR.Rows.Count > 0 Then
                For IntA As Integer = 0 To Me.tblAR.Rows.Count - 1
                    Me.rowAR = Me.tblAR.Rows(IntA)
                    With Me.rowAR
                        If .Primary = PrimaryARAkun Then
                            Me.cboPrimaryAR.SelectedIndex = IntA
                            Me.cboAccount.SelectedIndex = IntA
                        End If
                    End With
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSales_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSales.Enter
        Try
            Me.buffer = Me.TxtSales.Text
            Me.buffer2 = Me.TxtTotal.Text

            Me.ObjCD.Number_Enter(Me.TxtSales, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtVAT_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtVAT.Enter
        Try
            Me.buffer = Me.TxtVAT.Text
            Me.buffer2 = Me.TxtTotal.Text

            Me.ObjCD.Number_Enter(Me.TxtVAT, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtWeightBalanceCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtInvoiceNumber.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtInvoiceNumber.Enabled = True Then
                    If Me.TxtInvoiceNumber.ReadOnly = False Then
                        TxtInvoiceNumber.BackColor = Me.BackCOlorAll
                    Else
                        TxtInvoiceNumber.BackColor = Me.txtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtWeightBalanceCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtInvoiceNumber.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtInvoiceNumber.Enabled = True Then
                    If Me.TxtInvoiceNumber.ReadOnly = False Then
                        TxtInvoiceNumber.BackColor = Me.TxtPrimaryTransaksi.BackColor
                    Else
                        TxtInvoiceNumber.BackColor = Me.txtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCustCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSuppCode.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtSuppCode.Enabled = True Then
                    If Me.TxtSuppCode.ReadOnly = False Then
                        TxtSuppCode.BackColor = Me.BackCOlorAll
                    Else
                        TxtSuppCode.BackColor = Me.txtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCustCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSuppCode.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtSuppCode.Enabled = True Then
                    If Me.TxtSuppCode.ReadOnly = False Then
                        TxtSuppCode.BackColor = Me.TxtPrimaryTransaksi.BackColor
                    Else
                        TxtSuppCode.BackColor = Me.txtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSales_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSales.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtSales.Enabled = True Then
                    If Me.TxtSales.ReadOnly = False Then
                        TxtSales.BackColor = Me.BackCOlorAll
                    Else
                        TxtSales.BackColor = Me.txtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSales_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSales.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtSales.Enabled = True Then
                    If Me.TxtSales.ReadOnly = False Then
                        TxtSales.BackColor = Me.TxtPrimaryTransaksi.BackColor
                    Else
                        TxtSales.BackColor = Me.txtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtVAT_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtVAT.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtVAT.Enabled = True Then
                    If Me.TxtVAT.ReadOnly = False Then
                        TxtVAT.BackColor = Me.BackCOlorAll
                    Else
                        TxtVAT.BackColor = Me.txtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtVAT_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtVAT.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtVAT.Enabled = True Then
                    If Me.TxtVAT.ReadOnly = False Then
                        TxtVAT.BackColor = Me.TxtPrimaryTransaksi.BackColor
                    Else
                        TxtVAT.BackColor = Me.txtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReveived_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaid.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtPaid.Enabled = True Then
                    If Me.TxtPaid.ReadOnly = False Then
                        TxtPaid.BackColor = Me.BackCOlorAll
                    Else
                        TxtPaid.BackColor = Me.txtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRecieved_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaid.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtPaid.Enabled = True Then
                    If Me.TxtPaid.ReadOnly = False Then
                        Me.TxtPaid.BackColor = Me.TxtPrimaryTransaksi.BackColor
                    Else
                        Me.TxtPaid.BackColor = Me.txtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.ReadOnly = False Then
                    Me.TxtRemarks.BackColor = Me.BackCOlorAll
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
                    Me.TxtRemarks.BackColor = Me.TxtPrimaryTransaksi.BackColor
                Else
                    TxtRemarks.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindCust_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindSupplier.Click
        Dim i As New Core_Module.FrmCariMasterSupplier
        i.BackColorAll = Me.BackCOlorAll
        Try
            If Me.TxtSuppCode.Text <> "" Then
                i.txtCari.Text = Me.TxtSuppCode.Text
            End If

            i.ShowDialog()
            If i.PrimarySupplier > 0 Then
                Me.TxtPrimarySupplier.Text = i.PrimarySupplier
                Me.TxtSuppCode.Text = Me.ObjCust.GetSupplierCode(Me.TxtPrimarySupplier.Text)
                Me.TxtSuppName.Text = Me.ObjCust.GetSupplierName(Me.TxtPrimarySupplier.Text)
                Me.TxtInvoiceNumber.Focus()
            Else
                Me.TxtSuppCode.Clear()
                Me.TxtSuppName.Clear()
                Me.TxtPrimarySupplier.Text = "-1"
                Me.TxtSuppCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtCustCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtSuppCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtSuppCode.Text <> "" Then
                    If Me.ObjCust.ApakahSudahAda(Me.TxtSuppCode.Text) = True Then
                        Me.TxtPrimarySupplier.Text = Me.ObjCust.GetPrimary(Me.TxtSuppCode.Text)
                        Me.TxtSuppCode.Text = Me.ObjCust.GetSupplierCode(Me.TxtPrimarySupplier.Text)
                        Me.TxtSuppName.Text = Me.ObjCust.GetSupplierName(Me.TxtPrimarySupplier.Text)
                        Me.TxtInvoiceNumber.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.btnFindSupplier.PerformClick()
                    End If
                Else
                    Me.btnFindSupplier.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCustCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtSuppCode.Leave
        Dim i As New Core_Module.FrmCariMasterSupplier
        i.BackColorAll = Me.BackCOlorAll
        Try
            If Me.TxtSuppCode.Text <> "" Then
                If Me.ObjCust.ApakahSudahAda(Me.TxtSuppCode.Text) = True Then
                    Me.TxtPrimarySupplier.Text = Me.ObjCust.GetPrimary(Me.TxtSuppCode.Text)
                    Me.TxtSuppCode.Text = Me.ObjCust.GetSupplierCode(Me.TxtPrimarySupplier.Text)
                    Me.TxtSuppName.Text = Me.ObjCust.GetSupplierName(Me.TxtPrimarySupplier.Text)
                Else
                    If Me.TxtSuppCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtSuppCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimarySupplier > 0 Then
                        Me.TxtPrimarySupplier.Text = i.PrimarySupplier
                        Me.TxtSuppCode.Text = Me.ObjCust.GetSupplierCode(Me.TxtPrimarySupplier.Text)
                        Me.TxtSuppName.Text = Me.ObjCust.GetSupplierName(Me.TxtPrimarySupplier.Text)
                    Else
                        Me.TxtSuppCode.Clear()
                        Me.TxtSuppName.Clear()
                        Me.TxtPrimarySupplier.Text = "-1"
                        Me.TxtInvoiceNumber.Focus()
                    End If
                End If
            Else
                Me.TxtSuppCode.Clear()
                Me.TxtSuppName.Clear()
                Me.TxtPrimarySupplier.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtVAT_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtVAT.Leave
        Try
            'Me.TxtVAT.Text = Me.ObjCD.DigitGrouping(Me.TxtVAT.Text, Me.Default_Prec)
            Me.ObjCD.Number_Leave(Me.TxtVAT, sender, e, Me.Default_Prec, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRecieved_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaid.Leave
        Try
            'Me.TxtPaid.Text = Me.ObjCD.DigitGrouping(Me.TxtPaid.Text, Me.Default_Prec)
            Me.ObjCD.Number_Leave(Me.TxtPaid, sender, e, Me.Default_Prec, Me.ObjCD)
            If CDec(Me.TxtTotal.Text) < CDec(Me.TxtPaid.Text) Then
                MsgBox("Please input data paid <= total", MsgBoxStyle.Information, "information")
                Me.TxtPaid.Focus()
                Me.TxtPaid.Text = Me.buffer
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDutyStamp_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutyStamp.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtDutyStamp, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDutyStamp_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutyStamp.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtDutyStamp.Enabled = True Then
                    If Me.TxtDutyStamp.ReadOnly = False Then
                        TxtDutyStamp.BackColor = Me.BackCOlorAll
                    Else
                        TxtDutyStamp.BackColor = Me.txtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDutyStamp_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutyStamp.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtDutyStamp.Enabled = True Then
                    If Me.TxtDutyStamp.ReadOnly = False Then
                        TxtDutyStamp.BackColor = Me.TxtPrimaryTransaksi.BackColor
                    Else
                        TxtDutyStamp.BackColor = Me.txtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCurrency_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCurrency.SelectedValueChanged
        Dim row As dseMasterCurrency.tblmaster_currencyRow

        Try
            If Me.TxtRemarks.Enabled = True And Me.cboCurrency.Enabled = True Then
                If Me.cboCurrency.SelectedIndex = -1 Then
                    Me.TxtPrimaryCurrency.Text = -1
                    Me.TxtBalance.Text = Me.ObjCD.DigitGrouping(Me.TxtBalance.Text, Me.ObjCur.GetDecimal_DefaultCurrency)
                    Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping("1", Core_Module.DefPrecExchangeRate)
                    Me.IsiCBOArAkunByTipePerkiraan1()
                Else
                    row = Me.TBLCurr.DefaultView(Me.cboCurrency.SelectedIndex).Row
                    Me.TxtPrimaryCurrency.Text = row.Primary
                    Me.TxtBalance.Text = Me.ObjCD.DigitGrouping(Me.TxtBalance.Text, Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text))
                    Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping(Me.ObjCur.GetDefaultExhangeRate(Me.TxtPrimaryCurrency.Text), Core_Module.DefPrecExchangeRate)
                    Me.IsiCBOArAkunByTipePerkiraan1()
                End If
                Me.SetExchangeRate()

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub IsiCBOArAkunByTipePerkiraan1()
        '- isicbo customer
        Dim tp1 As New ClsMasterAkun.EnumTipeAkun
        Dim tp2 As New ClsMasterAkun.EnumTipeAkun2
        tp1 = ClsMasterAkun.EnumTipeAkun.Liabilities
        tp2 = ClsMasterAkun.EnumTipeAkun2.Account_Payable
        Dim objMulti As New ClsMultiAccount
        Try
            'If Me.TxtPrimaryTransaksi.Text = "-1" Then
            Me.cboAccount.DataSource = Nothing
            Me.cboAccount.Items.Clear()
            Me.cboPrimaryAR.DataSource = Nothing
            Me.cboPrimaryAR.Items.Clear()
            objMulti.IsiTBLAPARAcount(tp1, tp2, Me.tblAR, Me.TxtPrimaryCurrency.Text)
            Me.cboAccount.DataSource = Me.tblAR.DefaultView
            Me.cboAccount.DisplayMember = "Account Code"
            Me.cboPrimaryAR.DataSource = Me.tblAR.DefaultView
            Me.cboPrimaryAR.DisplayMember = "Primary"
            Me.txtPrimaryAccount.Text = Me.cboPrimaryAR.Text
            ' End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetExchangeRate()
        Try
            If Me.TxtPrimaryCurrency.Text = "" Or Me.TxtPrimaryCurrency.Text = "-1" Then
                Me.TxtExchangeRate.ReadOnly = True
                Me.TxtExchangeRate.BackColor = Me.txtReadOnly.BackColor
                Me.TxtExchangeRate.TabStop = False
            ElseIf Me.TxtPrimaryCurrency.Text = Core_Module.PrimaryBaseCurrency.ToString Then
                Me.TxtExchangeRate.ReadOnly = True
                Me.TxtExchangeRate.BackColor = Me.txtReadOnly.BackColor
                Me.TxtExchangeRate.TabStop = False
            Else
                Me.TxtExchangeRate.ReadOnly = False
                Me.TxtExchangeRate.BackColor = Me.TxtPrimaryCurrency.BackColor
                Me.TxtExchangeRate.TabStop = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExchangeRate.Enter
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.ObjCD.Number_Enter(Me.TxtExchangeRate, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExchangeRate.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtExchangeRate.Enabled = True Then
                    If Me.TxtExchangeRate.ReadOnly = True Then
                        Me.TxtExchangeRate.BackColor = Me.txtReadOnly.BackColor
                    Else
                        Me.TxtExchangeRate.BackColor = Me.BackCOlorAll
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtExchangeRate.KeyDown
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.ObjCD.Number_KeyDown(Me.TxtExchangeRate, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtExchangeRate.KeyUp
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.ObjCD.Number_KeyUp(Me.TxtExchangeRate, sender, e, False)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExchangeRate.Leave
        Try
            If Me.TxtRemarks.Enabled = True Then
                If CDec(Me.TxtExchangeRate.Text) <= 0 Then
                    Me.TxtExchangeRate.Text = 1
                End If
                Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping(Me.TxtExchangeRate.Text, Core_Module.DefPrecExchangeRate)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExchangeRate.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtExchangeRate.Enabled = True Then
                    If Me.TxtExchangeRate.ReadOnly = True Then
                        Me.TxtExchangeRate.BackColor = Me.txtReadOnly.BackColor
                    Else
                        Me.TxtExchangeRate.BackColor = Me.TxtPrimaryCurrency.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBalance_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBalance.Enter
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.ObjCD.Number_Enter(Me.TxtBalance, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBalance_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBalance.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtBalance.Enabled = True Then
                    If Me.TxtBalance.ReadOnly = True Then
                        Me.TxtBalance.BackColor = Me.txtReadOnly.BackColor
                    Else
                        Me.TxtBalance.BackColor = Me.BackCOlorAll
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBalance_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtBalance.KeyDown
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.ObjCD.Number_KeyDown(Me.TxtBalance, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBalance_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtBalance.KeyUp
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.ObjCD.Number_KeyUp(Me.TxtBalance, sender, e, False)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBalance_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBalance.Leave
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.TxtBalance.Text = Me.ObjCD.DigitGrouping(Me.TxtBalance.Text, Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBalance_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBalance.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtBalance.Enabled = True Then
                    If Me.TxtBalance.ReadOnly = True Then
                        Me.TxtBalance.BackColor = Me.txtReadOnly.BackColor
                    Else
                        Me.TxtBalance.BackColor = Me.TxtPrimaryCurrency.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@20151023-0830
    Private Sub llbImport_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbImport.LinkClicked
        Try
            'Me.ImportAP()
            Control.CheckForIllegalCrossThreadCalls = False

            If MsgBox("Start data import?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                Me.ControlBox = False
                Me.btnOK.Visible = False
                Me.btnCancel.Visible = False
                Me.LblProgress.Visible = True
                Me.ProgressBar1.Visible = True
                Me.llbImport.Enabled = False
                Me.LblProgress.Text = ""

                Me.btnFind.Enabled = False
                Me.btnNew.Enabled = False
                Me.btnEdit.Enabled = False
                Me.btnExit.Enabled = False
                Me.btnDelete.Enabled = False

                Me.BackgroundWorker1.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Me.ObjCD.UpdateThreading()
            Me.ImportAP()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ImportAP() As Boolean
        Dim frmImport As New CustomSMESS.FrmTransferBeginingBalance
        Dim dtaAPBalanceCopy As New CustomSMESS.DseTransferBeginingBalanceTableAdapters.tbltrans_balance_purchase_copyTableAdapter
        Dim tblAPBalanceCopy As New CustomSMESS.DseTransferBeginingBalance.tbltrans_balance_purchase_copyDataTable
        Dim rowAPBalanceCopy As CustomSMESS.DseTransferBeginingBalance.tbltrans_balance_purchase_copyRow

        Dim TBL As New DseTransaksiPurchaseBalance.tbltrans_historicalpurchasebalancemanualDataTable
        Dim row As DseTransaksiPurchaseBalance.tbltrans_historicalpurchasebalancemanualRow

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            Me.SedangEvaluasi = True
            Me.BackgroundWorker1.ReportProgress(1)
            While Me.SedangEvaluasi = True
                Threading.Thread.Sleep(1000)
            End While

            If Me.DeleteHistorical = True Then
                Me.ClearHistoricalPurchase()
            End If

            If Me.objImport.TransferAP(Me.ProgressBar1, Me.LblProgress) = False Then
                Exit Function
            End If

            dtaAPBalanceCopy.Fill(tblAPBalanceCopy)

            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = tblAPBalanceCopy.Count

            For Each rowAPBalanceCopy In tblAPBalanceCopy
                With rowAPBalanceCopy
                    Me.ProgressBar1.Value = tblAPBalanceCopy.Rows.IndexOf(rowAPBalanceCopy) + 1
                    Me.LblProgress.Text = "Saving to database, data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " - " + rowAPBalanceCopy.AOP_Number_Manual

                    TBL.Rows.Clear()
                    row = TBL.Newtbltrans_historicalpurchasebalancemanualRow
                    row.Primary = .Primary
                    row.Primary_Supplier = .Primary_Supplier
                    row.AOP_Number = .AOP_Number_Manual
                    row.AOP_Date = .AOP_Date_Manual
                    row.Due_Date = .Due_Date
                    row.Remarks = .Remarks
                    row.Created_By = .Created_By
                    row.Created_Date = .Created_Date
                    row.Modified_By = .Modified_By
                    row.Modified_Date = .Modified_Date
                    row.Timestamp = .Timestamp
                    row.Purchase = .Purchase
                    row.VAT = .Purchase_VAT
                    row.Total = .Balance
                    row.Paid = .Paid
                    row.Balance = .balance_original
                    row.Dutystamp = .Duty_stamp
                    row.primary_currency = .primary_currency
                    row.exchange_rate = .exchange_rate
                    TBL.Addtbltrans_historicalpurchasebalancemanualRow(row)
                    If Me.objImport.ApakahBisaImport = True Then
                        If Me.ObjProcess.CariDetail(TBL, .Primary) = True Then
                            Me.ObjProcess.SimpanBaru(Me.UserCode)
                        End If
                    End If
                End With
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub ClearHistoricalPurchase()
        Dim dtaHistoricalPurchaseBalance As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_historicalpurchasebalancemanualTableAdapter
        Dim tblHistoricalPurchaseBalance As New DseTransaksiPurchaseBalance.tbltrans_historicalpurchasebalancemanualDataTable
        Dim rowHistoricalPurchaseBalance As DseTransaksiPurchaseBalance.tbltrans_historicalpurchasebalancemanualRow

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.ObjCD.UpdateThreading()

            dtaHistoricalPurchaseBalance.Fill(tblHistoricalPurchaseBalance)

            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = tblHistoricalPurchaseBalance.Count
            Me.LblProgress.Text = ""

            Me.ProgressBar1.Visible = True
            Me.LblProgress.Visible = True

            For Each rowHistoricalPurchaseBalance In tblHistoricalPurchaseBalance
                Me.ProgressBar1.Value = tblHistoricalPurchaseBalance.Rows.IndexOf(rowHistoricalPurchaseBalance) + 1
                Me.LblProgress.Text = "Deleting data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " - " + rowHistoricalPurchaseBalance.AOP_Number
                Me.ObjProcess.Hapus(rowHistoricalPurchaseBalance.Primary, Me.UserCode)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = 0
            dtaHistoricalPurchaseBalance.Dispose()
            tblHistoricalPurchaseBalance.Dispose()
        End Try
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Try
            If e.ProgressPercentage = 1 Then
                If MsgBox("Clear data before transfer?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.DeleteHistorical = True
                Else
                    Me.DeleteHistorical = False
                End If
            End If
            Me.SedangEvaluasi = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.ControlBox = True
            Me.btnOK.Visible = True
            Me.btnCancel.Visible = True
            Me.LblProgress.Visible = False
            Me.ProgressBar1.Visible = False
            Me.llbImport.Enabled = True
            Me.ProgressBar1.Value = 0

            Me.btnFind.Enabled = True
            Me.btnNew.Enabled = True
            Me.btnEdit.Enabled = True
            Me.btnExit.Enabled = True
            Me.btnDelete.Enabled = True

            MsgBox("Purchase balance transfer completed.", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboAccount_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboAccount.SelectedValueChanged

        Try
            'Me.txtPrimaryAccount.Visible = True
            If IsNumeric(Me.cboPrimaryAR.Text) AndAlso Me.cboPrimaryAR.Text > -1 Then
                Me.txtPrimaryAccount.Text = Me.cboPrimaryAR.Text
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboPrimaryAR_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPrimaryAR.SelectedValueChanged

        Try
            'Me.txtPrimaryAccount.Visible = True
            If IsNumeric(Me.cboPrimaryAR.Text) AndAlso Me.cboPrimaryAR.Text > -1 Then
                Me.txtPrimaryAccount.Text = Me.cboPrimaryAR.Text
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub txtPrimaryAccount_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrimaryAccount.TextChanged
    '    Dim PrimaryARAkun As Double = -1

    '    Try
    '        If IsNumeric(Me.txtPrimaryAccount.Text) Then
    '            PrimaryARAkun = CDbl(Me.txtPrimaryAccount.Text)
    '        End If

    '        If Me.tblAR.Rows.Count > 0 Then
    '            For IntA As Integer = 0 To Me.tblAR.Rows.Count - 1
    '                Me.rowAR = Me.tblAR.Rows(IntA)
    '                With Me.rowAR
    '                    If .Primary = PrimaryARAkun Then
    '                        Me.cboPrimaryAR.SelectedIndex = IntA
    '                        Me.cboAccount.SelectedIndex = IntA
    '                    End If
    '                End With
    '            Next
    '        End If

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Public Function FromCutOff_PurchBalanceManual(ByVal vUserCode As String, ByVal vPrimSupp As Double, ByVal vAopNumb As String, _
     ByVal vAopDate As Date, ByVal vDueDate As Date, ByVal vBalance As Decimal, ByVal vPrimaryAP As Double) As Boolean
        Try
            Me.ObjProcess = New ClsTransaksiPurchaseBalance

            Me.baru = True

            With Me.ObjProcess
                .Primary_Supplier = vPrimSupp
                .AOPNumber = vAopNumb
                .AOPDate = vAopDate
                .Purchase = 0
                .VAT = 0
                .Total = 0
                .DutyStamp = 0
                .DueDate = vDueDate
                .Paid = 0
                .Balance = vBalance
                .Remarks = ""
                .PrimaryCurrency = 1
                .ExchangeRate = 1
                .PrimaryAP = vPrimaryAP
            End With

            If Me.ObjProcess.SimpanBaru(vUserCode) = False Then
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            Me.ObjProcess = Nothing
        End Try
    End Function
End Class