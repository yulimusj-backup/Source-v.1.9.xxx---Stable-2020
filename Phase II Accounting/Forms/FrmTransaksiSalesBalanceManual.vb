Option Explicit On
Option Strict Off

Imports Core_Module
Public Class FrmTransaksiSalesBalanceManual
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightSeaGreen

    Public Import_CustCode, Import_CustName, Import_InvoiceNo, Import_Curr, Import_ARAccount, Import_Remarks, Import_ARAccountDefault As String
    Public Import_InvDate, Import_DueDate As Date
    Public Import_Amount, Import_ExchangeRate As Decimal

    Dim baru As Boolean
    Dim ObjProcess As ClsTransaksiSalesBalance
    Dim buffer As String = ""
    Dim buffer2 As String = ""
    Dim ObjCD As New ClsCekData
    Dim ObjCust As New clsMasterCustomer
    Dim ObjClosing As New ClsTransaksiClosing
    Dim Def_Prec As Integer = 0
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjAkun As New ClsMasterAkun
    Dim TBLCurr As New dseMasterCurrency.tblmaster_currencyDataTable
    Dim ObjIdentifier As New CustomSMESS.ClsIdentifier

    '-balikin
    Dim objImportAR As New CustomSMESS.ClsTrans_ImportARBalance
    Dim vHapusHistorical As Boolean = False
    Dim vSedangEvaluasi As Boolean = False

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
            Me.llbImport.Enabled = False
            Me.cboAccount.Enabled = True
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
            Me.TxtPrimaryCustomer.Text = "-1"
            Me.TxtSales.Text = Me.ObjCD.DigitGrouping(0, Me.Def_Prec)
            Me.TxtVAT.Text = Me.ObjCD.DigitGrouping(0, Me.Def_Prec)
            Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(0, Me.Def_Prec)
            Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(0, Me.Def_Prec)
            Me.TxtRecieved.Text = Me.ObjCD.DigitGrouping(0, Me.Def_Prec)
            Me.TxtBalance.Text = Me.ObjCD.DigitGrouping(0, Me.Def_Prec)
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

    Private Sub IsiCBOArAkunByTipePerkiraan1()
        '- isicbo customer
        Dim tp1 As New ClsMasterAkun.EnumTipeAkun
        Dim tp2 As New ClsMasterAkun.EnumTipeAkun2
        tp1 = ClsMasterAkun.EnumTipeAkun.Asset
        tp2 = ClsMasterAkun.EnumTipeAkun2.Account_Receivable
        Dim objMulti As New ClsMultiAccount
        Try
            ' If Me.TxtPrimaryTransaksi.Text = "-1" Then
            Me.cboAccount.DataSource = Nothing
            Me.cboAccount.Items.Clear()
            Me.cboPrimaryAR.DataSource = Nothing
            Me.cboPrimaryAR.Items.Clear()
            objMulti.IsiTBLAPARAcount(tp1, tp2, Me.tblAR, Me.TxtPrimaryCurrency.Text)
            'Dim I As New Form5
            'I.DataGridView1.DataSource = Nothing
            'I.DataGridView1.DataSource = Me.tblAR
            'I.ShowDialog()

            Me.cboAccount.DataSource = Me.tblAR.DefaultView
            Me.cboAccount.DisplayMember = "Account Code"
            Me.cboPrimaryAR.DataSource = Me.tblAR.DefaultView
            Me.cboPrimaryAR.DisplayMember = "Primary"
            Me.txtPrimaryAccount.Text = Me.cboPrimaryAR.Text



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
            KoleksiNamaData.Add("Customer Code")
            KoleksiTextBox.Add(Me.TxtCustCode)
            KoleksiNamaData.Add("Invoice Number")
            KoleksiTextBox.Add(Me.TxtInvoiceNumber)


            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data default precision harus non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'cek yang tidak boleh = nol
            'If Me.TxtSales.Text = "0" Or Me.TxtSales.Text = "" Then
            '    MsgBox("Please input sales", MsgBoxStyle.Information, "Information")
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
                MsgBox("Please input due date <= invoice date", MsgBoxStyle.Information, "Information")
                Me.Dtpduedate.Focus()
                Return False
            End If

            'Cek perhitungan balance
            If CDec(Me.TxtTotal.Text) < CDec(Me.TxtRecieved.Text) Then
                MsgBox("Please input data received <= total", MsgBoxStyle.Information, "information")
                Me.TxtRecieved.Focus()
                Me.TxtRecieved.Text = Me.buffer
                Me.TxtBalance.Text = Me.buffer2
                Return False
            End If

            'cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.ObjProcess.ApakahSudahAda(Me.TxtInvoiceNumber.Text) = True Then
                    MsgBox("Duplication in invoice number, please change your number", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtInvoiceNumber.Focus()
                    Return False
                End If
            Else
                Dim TipeCodeLama As String = Me.ObjProcess.GetInvoiceNumberByPrimary(Me.TxtPrimaryTransaksi.Text)
                If TipeCodeLama <> Me.TxtInvoiceNumber.Text Then
                    If Me.ObjProcess.ApakahSudahAda(Me.TxtInvoiceNumber.Text) = True Then
                        MsgBox("Duplication in invoice number, please change your number", _
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

            If Not IsNumeric(Me.txtPrimaryAccount.Text) OrElse Me.txtPrimaryAccount.Text = -1 Then
                MsgBox("Please choose AR Account", MsgBoxStyle.Information, "information")
                Me.cboAccount.Focus()
                Return False
            End If

            'kalo tanggal transaksi sama atau lebih besar dari convertion date maka ditolak
            ConvertionDate = Me.ObjClosing.GetConvertionDate
            If ConvertionDate <= Me.dtpInvoiceDate.Value Then
                If ObjIdentifier.GetIdentifier = "smess_lmb" Then
                    If #1/1/2018# <= Me.dtpInvoiceDate.Value.Date Then
                        MsgBox("The convertion date is " + Format(#1/1/2018#, "dd MMM yyyy") + _
                                           " this transaction is for historical data only. Every transaction must be dated before convertion date.", _
                                           MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                Else
                    MsgBox("The convertion date is " + Format(ConvertionDate, "dd MMM yyyy") + _
                    " this transaction is for historical data only. Every transaction must be dated before convertion date.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'balance dan exchange rate harus >0
            If CDec(Me.TxtBalance.Text) <= 0 Then
                MsgBox("Please input balance with positive numbers.", MsgBoxStyle.Information, "Information")
                Me.TxtBalance.Focus()
                Return False
            End If
            If CDec(Me.TxtExchangeRate.Text) <= 0 Then
                MsgBox("Please input exchange rate with positive numbers.", MsgBoxStyle.Information, "Information")
                Me.TxtExchangeRate.Focus()
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
                .Primary_Customer = Me.TxtPrimaryCustomer.Text
                .Invoice_Number = Me.TxtInvoiceNumber.Text
                .Invoice_Date = Me.dtpInvoiceDate.Value
                .Sales = Me.TxtSales.Text
                .VAT = Me.TxtVAT.Text
                .Dutystamp = Me.TxtDutystamp.Text
                .Total = Me.TxtTotal.Text
                .DueDate = Me.Dtpduedate.Value
                .Recieved = Me.TxtRecieved.Text
                .Balance = Me.TxtBalance.Text
                .Remarks = Me.TxtRemarks.Text
                .PrimaryCurrency = Me.TxtPrimaryCurrency.Text
                .ExchangeRate = Me.TxtExchangeRate.Text
                .PrimaryAccount = Me.txtPrimaryAccount.Text
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
                Me.TxtPrimaryCustomer.Text = .Primary_Customer
                Me.TxtCustCode.Text = ObjCust.GetCustomerCode(Me.TxtPrimaryCustomer.Text)
                Me.TxtCustName.Text = ObjCust.GetCustName(Me.TxtPrimaryCustomer.Text)
                Me.TxtInvoiceNumber.Text = .Invoice_Number
                Me.dtpInvoiceDate.Value = .Invoice_Date
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

                Me.Def_Prec = Me.ObjCur.GetDefaultPrecission(.PrimaryCurrency)
                Me.TxtSales.Text = ObjCD.DigitGrouping(.Sales, Me.Def_Prec)
                Me.TxtVAT.Text = ObjCD.DigitGrouping(.VAT, Me.Def_Prec)
                Me.TxtDutystamp.Text = ObjCD.DigitGrouping(.Dutystamp, Me.Def_Prec)
                Me.TxtTotal.Text = ObjCD.DigitGrouping(.Total, Me.Def_Prec)
                Dtpduedate.Value = .DueDate
                Me.TxtRecieved.Text = ObjCD.DigitGrouping(.Recieved, Me.Def_Prec)
                Me.TxtBalance.Text = ObjCD.DigitGrouping(.Balance, Me.Def_Prec)

                Me.IsiCBOArAkunByTipePerkiraan1()

                If .PrimaryAccount > -1 Then
                    Me.txtPrimaryAccount.Text = .PrimaryAccount
                    Me.cboPrimaryAR.Text = .PrimaryAccount
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
            Me.ObjProcess = New ClsTransaksiSalesBalance
            Me.baru = True
            Me.KosongkanField()
            Me.NonAktifkanDetail()

            Me.Def_Prec = Me.ObjCur.GetDecimal_DefaultCurrency
            'isikan cbo currency
            Me.ObjCur.IsiCBO(Me.cboCurrency, Me.TBLCurr)

            ''@rosadinnn
            ''@20151022-1009
            Me.llbImport.Visible = False
            Me.ProgressBar1.Visible = False
            Me.ShowIcon = False
            '-balikin lagi
            If Me.objImportAR.ApakahBisaImport = True Then
                Me.llbImport.Visible = True
                Me.ProgressBar1.Visible = True
            Else
                If Me.Import_CustCode <> "" Then
                    Me.ImportData()
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ImportData()
        Try
            Me.Location = New Point((Me.Width * -1) - 5, (Me.Height * -1) - 5)

            Me.btnNew.PerformClick()

            'salin ke obj
            With Me.ObjProcess
                .Primary_Customer = Me.ObjCust.GetPrimary(Me.Import_CustCode)
                .Invoice_Number = Me.Import_InvoiceNo
                .Invoice_Date = Me.Import_InvDate
                .Sales = Me.Import_Amount
                .VAT = 0
                .Dutystamp = 0
                .Total = Me.Import_Amount
                .DueDate = Me.Import_DueDate
                .Recieved = 0
                .Balance = Me.Import_Amount
                .Remarks = Me.Import_Remarks
                .PrimaryCurrency = Me.ObjCur.GetPrimary(Me.Import_Curr)
                .ExchangeRate = Me.Import_ExchangeRate
                .PrimaryAccount = Me.ObjAkun.GetPrimary(Me.Import_ARAccount)

                .CreatedBy = Me.UserCode
                .CreatedDate = Now
                .ModifiedBy = Me.UserCode
                .ModifiedDate = Now
                .TimeStamp = Now
            End With

            Me.SalinKeField()

            If Me.ObjProcess.SimpanBaru(Me.UserCode) Then
                Me.NonAktifkanDetail()
                Me.SalinKeField()
                Me.btnFind.Focus()
            End If

            Me.Close()
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
        Dim ij As New FrmTransaksiSalesBalanceManual_Cari
        Try
            ij.BackColorAll = Me.BackColorAll
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
        Try
            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.TxtCustCode.Focus()
            Me.TxtPrimaryTransaksi.Text = "-1"
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
            If Me.TxtInvoiceNumber.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtInvoiceNumber.Text <> "" Then
                If Me.ObjProcess.ApakahBolehEdit(Me.TxtPrimaryTransaksi.Text, Me.UserCode, True, True) = True Then
                    Me.baru = False
                    Me.AktifkanDetail()
                    Me.TxtCustCode.Focus()
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
                If Me.ObjProcess.ApakahBolehHapus(Me.TxtPrimaryTransaksi.Text, Me.UserCode, True, True) = False Then
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


    Private Sub TxtRecieved_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRecieved.Enter
        Try
            Me.buffer = Me.TxtRecieved.Text
            Me.buffer2 = Me.TxtBalance.Text
            Me.ObjCD.Number_Enter(Me.TxtRecieved, sender, e)
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

    Private Sub TxtDutystamp_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutystamp.Enter
        Try
            Me.buffer = Me.TxtDutystamp.Text
            Me.buffer2 = Me.TxtTotal.Text

            Me.ObjCD.Number_Enter(Me.TxtDutystamp, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtWeightBalanceCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtInvoiceNumber.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtInvoiceNumber.Enabled = True Then
                    If Me.TxtInvoiceNumber.ReadOnly = False Then
                        TxtInvoiceNumber.BackColor = Me.BackColorAll
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

    Private Sub TxtCustCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustCode.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtCustCode.Enabled = True Then
                    If Me.TxtCustCode.ReadOnly = False Then
                        TxtCustCode.BackColor = Me.BackColorAll
                    Else
                        TxtCustCode.BackColor = Me.txtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCustCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustCode.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtCustCode.Enabled = True Then
                    If Me.TxtCustCode.ReadOnly = False Then
                        TxtCustCode.BackColor = Me.TxtPrimaryTransaksi.BackColor
                    Else
                        TxtCustCode.BackColor = Me.txtReadOnly.BackColor
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
                        TxtSales.BackColor = Me.BackColorAll
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
                        TxtVAT.BackColor = Me.BackColorAll
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

    Private Sub TxtDutystamp_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutystamp.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtDutystamp.Enabled = True Then
                    If Me.TxtDutystamp.ReadOnly = False Then
                        TxtDutystamp.BackColor = Me.BackColorAll
                    Else
                        TxtDutystamp.BackColor = Me.txtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDutystamp_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutystamp.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtDutystamp.Enabled = True Then
                    If Me.TxtDutystamp.ReadOnly = False Then
                        TxtDutystamp.BackColor = Me.TxtPrimaryTransaksi.BackColor
                    Else
                        TxtDutystamp.BackColor = Me.txtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReveived_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRecieved.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRecieved.Enabled = True Then
                    If Me.TxtRecieved.ReadOnly = False Then
                        TxtRecieved.BackColor = Me.BackColorAll
                    Else
                        TxtRecieved.BackColor = Me.txtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRecieved_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRecieved.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRecieved.Enabled = True Then
                    If Me.TxtRecieved.ReadOnly = False Then
                        Me.TxtRecieved.BackColor = Me.TxtPrimaryTransaksi.BackColor
                    Else
                        Me.TxtRecieved.BackColor = Me.txtReadOnly.BackColor
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
                    Me.TxtRemarks.BackColor = Me.TxtPrimaryTransaksi.BackColor
                Else
                    TxtRemarks.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindCust_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCust.Click
        Dim i As New Core_Module.FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCustCode.Text <> "" Then
                i.txtCari.Text = Me.TxtCustCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryCustomer > 0 Then
                Me.TxtPrimaryCustomer.Text = i.PrimaryCustomer
                Me.TxtCustCode.Text = Me.ObjCust.GetCustomerCode(Me.TxtPrimaryCustomer.Text)
                Me.TxtCustName.Text = Me.ObjCust.GetCustName(Me.TxtPrimaryCustomer.Text)
                Me.TxtInvoiceNumber.Focus()
            Else
                Me.TxtCustCode.Clear()
                Me.TxtCustName.Clear()
                Me.TxtPrimaryCustomer.Text = "-1"
                Me.TxtCustCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtCustCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtCustCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtCustCode.Text <> "" Then
                    If Me.ObjCust.ApakahSudahAda(Me.TxtCustCode.Text) = True Then
                        Me.TxtPrimaryCustomer.Text = Me.ObjCust.GetPrimary(Me.TxtCustCode.Text)
                        Me.TxtCustCode.Text = Me.ObjCust.GetCustomerCode(Me.TxtPrimaryCustomer.Text)
                        Me.TxtCustName.Text = Me.ObjCust.GetCustName(Me.TxtPrimaryCustomer.Text)
                        Me.TxtInvoiceNumber.Focus()
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

    Private Sub txtCustCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtCustCode.Leave
        Dim i As New Core_Module.FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCustCode.Text <> "" Then
                If Me.ObjCust.ApakahSudahAda(Me.TxtCustCode.Text) = True Then
                    Me.TxtPrimaryCustomer.Text = Me.ObjCust.GetPrimary(Me.TxtCustCode.Text)
                    Me.TxtCustCode.Text = Me.ObjCust.GetCustomerCode(Me.TxtPrimaryCustomer.Text)
                    Me.TxtCustName.Text = Me.ObjCust.GetCustName(Me.TxtPrimaryCustomer.Text)
                Else
                    If Me.TxtCustCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtCustCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.TxtPrimaryCustomer.Text = i.PrimaryCustomer
                        Me.TxtCustCode.Text = Me.ObjCust.GetCustomerCode(Me.TxtPrimaryCustomer.Text)
                        Me.TxtCustName.Text = Me.ObjCust.GetCustName(Me.TxtPrimaryCustomer.Text)
                    Else
                        Me.TxtCustCode.Clear()
                        Me.TxtCustName.Clear()
                        Me.TxtPrimaryCustomer.Text = "-1"
                        Me.TxtInvoiceNumber.Focus()
                    End If
                End If
            Else
                Me.TxtCustCode.Clear()
                Me.TxtCustName.Clear()
                Me.TxtPrimaryCustomer.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
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
            ' Me.txtPrimaryAccount.Visible = True
            If IsNumeric(Me.cboPrimaryAR.Text) AndAlso Me.cboPrimaryAR.Text > -1 Then
                Me.txtPrimaryAccount.Text = Me.cboPrimaryAR.Text
            End If

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
                        Me.TxtExchangeRate.BackColor = Me.BackColorAll
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
                        Me.TxtBalance.BackColor = Me.BackColorAll
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

    Private Sub TxtBalance_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtBalance.TextChanged

    End Sub

    Private Sub llbImport_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbImport.LinkClicked
        Try
            'me.ImportAR
            Control.CheckForIllegalCrossThreadCalls = False
            If MsgBox("Start historical sales balance data import?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.LblProgress.Text = ""
                Me.LblProgress.Visible = True
                Me.BackgroundWorker1.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    '-balikin lagi
    Private Sub HapusSemuaHistorical()
        Dim DTA As New DseTransaksiSalesBalanceTableAdapters.tbltrans_historicalsalesTableAdapter
        Dim TBL As New DseTransaksiSalesBalance.tbltrans_historicalsalesDataTable
        Dim row As DseTransaksiSalesBalance.tbltrans_historicalsalesRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
            If TBL.Rows.Count > 0 Then
                Me.ProgressBar1.Value = 0
                Me.ProgressBar1.Maximum = TBL.Rows.Count

                For Each row In TBL.Rows
                    Me.ProgressBar1.Value += 1
                    Me.LblProgress.Text = "Deleting data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " : " + row.Invoice_Number
                    If Me.ObjProcess.ApakahBolehHapus(row.Primary, Me.UserCode, False, False) = True Then
                        Me.ObjProcess.Hapus(row.Primary, Me.UserCode)
                    End If
                Next
            End If

            Me.LblProgress.Text = ""
            Me.ProgressBar1.Value = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    '-balikin lagi
    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        'struktur untuk TblBalanceSalesCopy dimodule custom harus slalu sama dengan struktur tbl balance sales normal
        Dim TblBalanceSalesCopy As New CustomSMESS.DseTransferBeginingBalance.tbltrans_balance_sales_copyDataTable
        Dim TBL As New DseTransaksiSalesBalance.tbltrans_historicalsalesDataTable
        Dim row As DseTransaksiSalesBalance.tbltrans_historicalsalesRow
        Dim RowBalanceSalesCopy As CustomSMESS.DseTransferBeginingBalance.tbltrans_balance_sales_copyRow

        Try
            llbImport.Enabled = False
            btnFind.Enabled = False
            btnNew.Enabled = False
            btnEdit.Enabled = False
            btnDelete.Enabled = False
            btnExit.Enabled = False
            Me.ControlBox = False
            Me.LblProgress.Text = ""
            Me.LblProgress.Visible = True
            Me.btnOK.Visible = False
            Me.btnCancel.Visible = False

            Me.ObjCD.UpdateThreading()

            Me.vSedangEvaluasi = True
            Me.BackgroundWorker1.ReportProgress(0)
            While Me.vSedangEvaluasi = True
                Threading.Thread.Sleep(1000)
            End While
            If Me.vHapusHistorical = True Then
                Me.HapusSemuaHistorical()
            End If

            Me.objImportAR.ImportToDB(Me.ProgressBar1, TblBalanceSalesCopy, Me.LblProgress)
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = TblBalanceSalesCopy.Count

            For Each RowBalanceSalesCopy In TblBalanceSalesCopy
                Me.ProgressBar1.Value = TblBalanceSalesCopy.Rows.IndexOf(RowBalanceSalesCopy) + 1
                Me.LblProgress.Text = "Saving to database - data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " - " + RowBalanceSalesCopy.Invoice_Number_Manual

                TBL.Clear()
                row = TBL.Newtbltrans_historicalsalesRow
                With row
                    .Primary = -1
                    .Primary_Customer = RowBalanceSalesCopy.Primary_Customer
                    .Invoice_Number = RowBalanceSalesCopy.Invoice_Number_Manual
                    .Invoice_Date = RowBalanceSalesCopy.Invoice_Date_Manual
                    .Due_Date = RowBalanceSalesCopy.Due_Date
                    .Remarks = RowBalanceSalesCopy.Remarks
                    .Created_By = RowBalanceSalesCopy.Created_By
                    .Created_Date = RowBalanceSalesCopy.Created_Date
                    .Modified_By = RowBalanceSalesCopy.Modified_By
                    .Modified_Date = RowBalanceSalesCopy.Modified_Date
                    .Timestamp = RowBalanceSalesCopy.Timestamp
                    .Sales = RowBalanceSalesCopy.Sales
                    .VAT = RowBalanceSalesCopy.Sales_VAT
                    .Dutystamp = RowBalanceSalesCopy.Duty_stamp
                    .Total = RowBalanceSalesCopy.Balance
                    .Recieved = RowBalanceSalesCopy.Received
                    .Balance = RowBalanceSalesCopy.Balance
                    .primary_currency = RowBalanceSalesCopy.primary_currency
                    .exchange_rate = RowBalanceSalesCopy.exchange_rate
                End With
                TBL.Rows.Add(row)

                If Me.ObjProcess.CariDetail(TBL) = True Then
                    Me.ObjProcess.SimpanBaru(Me.UserCode)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            TblBalanceSalesCopy = Nothing
            llbImport.Enabled = True
            btnFind.Enabled = True
            btnNew.Enabled = True
            btnEdit.Enabled = True
            btnDelete.Enabled = True
            btnExit.Enabled = True
            Me.ControlBox = True
            Me.LblProgress.Visible = False
            Me.btnOK.Visible = True
            Me.btnCancel.Visible = True
        End Try
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Try
            If MsgBox("Delete all historical sales before proceeding?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.vHapusHistorical = True
            Else
                Me.vSedangEvaluasi = False
            End If
            Me.vSedangEvaluasi = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        MsgBox("Historical sales balance data import completed.", MsgBoxStyle.Information, "Information")
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

    Public Function FromCutOff_SalesBalanceManual(ByVal vUserCode As String, ByVal vPrimCust As Double, ByVal vInvNo As String, _
   ByVal vInvDate As Date, ByVal vDueDate As Date, ByVal vBalance As Decimal, ByVal vPrimaryAR As Double) As Boolean
        Try
            Me.ObjProcess = New ClsTransaksiSalesBalance

            Me.baru = True

            With Me.ObjProcess
                .Primary_Customer = vPrimCust
                .Invoice_Number = vInvNo
                .Invoice_Date = vInvDate
                .Sales = 0
                .VAT = 0
                .Dutystamp = 0
                .Total = 0
                .DueDate = vDueDate
                .Recieved = 0
                .Balance = vBalance
                .Remarks = ""
                .PrimaryCurrency = 1
                .ExchangeRate = 1
                .PrimaryAccount = vPrimaryAR
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

    Private Sub llbImportAR_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbImportAR.LinkClicked
        Dim Frm As New FrmTransaksiSalesBalanceManual_Import

        Try
            Frm.UserCode = Me.UserCode
            Frm.BackCOlorAll = Me.BackColorAll
            Frm.ShowDialog()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Frm.Dispose()
        End Try
    End Sub
End Class