Option Explicit On
Option Strict Off

Imports Core_Module
Public Class FrmTransaksiSalesReturnBalanceManual
    Public UserCode As String = ""
    Public BackCOlorAll As Color = Color.LightSeaGreen

    Dim baru As Boolean
    Dim ObjProcess As ClsTransaksiSalesReturnBalance
    Dim buffer As String = ""
    Dim buffer2 As String = ""
    Dim ObjCD As New ClsCekData
    Dim ObjCust As New clsMasterCustomer
    Dim ObjClosing As New ClsTransaksiClosing
    Dim Def_Prec As Integer = 0
    Dim ObjCur As New clsMasterCurrency
    Dim TBLCurr As New dseMasterCurrency.tblmaster_currencyDataTable
    Dim ObjIdentifier As New CustomSMESS.ClsIdentifier


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
            Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(0, Me.Def_Prec)
            Me.TxtPaid.Text = Me.ObjCD.DigitGrouping(0, Me.Def_Prec)
            Me.TxtBalance.Text = Me.ObjCD.DigitGrouping(0, Me.Def_Prec)
            Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(0, Me.Def_Prec)
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
            'If Me.TxtPrimaryTransaksi.Text = "-1" Then
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
            ' End If

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
            KoleksiNamaData.Add("SR Number")
            KoleksiTextBox.Add(Me.TxtInvoiceNumber)


            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data default precision harus non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'cek yang tidak boleh = nol
            'If Me.TxtSales.Text = "0" Or Me.TxtSales.Text = "" Then
            '    MsgBox("Please input sales return", MsgBoxStyle.Information, "Information")
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
                MsgBox("Please input due date <= SR date", MsgBoxStyle.Information, "Information")
                Me.Dtpduedate.Focus()
                Return False
            End If

            If Not IsNumeric(Me.txtPrimaryAccount.Text) OrElse Me.txtPrimaryAccount.Text = -1 Then
                MsgBox("Please choose AR Account", MsgBoxStyle.Information, "information")
                Me.cboAccount.Focus()
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

            'cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.ObjProcess.ApakahSudahAda(Me.TxtInvoiceNumber.Text) = True Then
                    MsgBox("Duplication in SR number, please change your number", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtInvoiceNumber.Focus()
                    Return False
                End If
            Else
                Dim TipeCodeLama As String = Me.ObjProcess.GetAOSRNumberByPrimary(Me.TxtPrimaryTransaksi.Text)
                If TipeCodeLama <> Me.TxtInvoiceNumber.Text Then
                    If Me.ObjProcess.ApakahSudahAda(Me.TxtInvoiceNumber.Text) = True Then
                        MsgBox("Duplication in SR number, please change your number", _
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

            'balance harus positive
            If CDec(Me.TxtBalance.Text) <= 0 Then
                MsgBox("Please enter positive number for balance data.", MsgBoxStyle.Information, "Information")
                Me.TxtBalance.Focus()
                Return False
            End If
            If CDec(Me.TxtExchangeRate.Text) <= 0 Then
                MsgBox("Please input exchange rate with positive numbers.", MsgBoxStyle.Information, "Information")
                Me.TxtExchangeRate.Focus()
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
                .Primary_Customer = Me.TxtPrimaryCustomer.Text
                .AOSRNumber = Me.TxtInvoiceNumber.Text
                .AOSRDate = Me.dtpInvoiceDate.Value
                .SalesReturn = 0
                .VATReturn = 0
                .Total = 0
                .DueDate = Me.Dtpduedate.Value
                .Paid = 0
                .Balance = Me.TxtBalance.Text
                .Remarks = Me.TxtRemarks.Text
                .Dutystamp = Me.TxtDutystamp.Text

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
                Me.TxtInvoiceNumber.Text = .AOSRNumber
                Me.dtpInvoiceDate.Value = .AOSRDate
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
                Me.TxtSales.Text = ObjCD.DigitGrouping(.SalesReturn, Me.Def_Prec)
                Me.TxtVAT.Text = ObjCD.DigitGrouping(.VATReturn, Me.Def_Prec)
                Me.TxtTotal.Text = ObjCD.DigitGrouping(.Total, Me.Def_Prec)
                Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(.Dutystamp, Me.Def_Prec)
                Dtpduedate.Value = .DueDate
                Me.TxtPaid.Text = ObjCD.DigitGrouping(.Paid, Me.Def_Prec)
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
            Me.ObjProcess = New ClsTransaksiSalesReturnBalance
            Me.baru = True
            Me.KosongkanField()
            Me.NonAktifkanDetail()

            Me.Def_Prec = Me.ObjCur.GetDecimal_DefaultCurrency
            Me.ObjCur.IsiCBO(Me.cboCurrency, Me.TBLCurr)
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
        Dim ij As New FrmTransaksiSalesReturnBalanceManual_Cari
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
        Me.baru = True
        Me.KosongkanField()
        Me.AktifkanDetail()
        Me.TxtCustCode.Focus()
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

            If Me.ObjProcess.ApakahBolehEdit(Me.TxtPrimaryTransaksi.Text, Me.UserCode, True, True) = True Then
                If Me.TxtInvoiceNumber.Text <> "" Then
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
            Me.buffer = Me.txtRemarks.Text
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
            If Me.txtRemarks.Enabled = True Then
                If Me.txtRemarks.ReadOnly = False Then
                    Me.TxtRemarks.BackColor = Me.TxtPrimaryTransaksi.BackColor
                Else
                    txtRemarks.BackColor = Me.txtReadOnly.BackColor
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

    Private Sub TxtDutystamp_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutystamp.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtDutystamp, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDutyStamp_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutystamp.GotFocus
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

    Private Sub TxtDutyStamp_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutystamp.LostFocus
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
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.TxtBalance.Text, Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text))
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

    Private Sub cboAccount_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboAccount.SelectedIndexChanged
        Try
            'Me.txtPrimaryAccount.Visible = True
            If IsNumeric(Me.cboPrimaryAR.Text) AndAlso Me.cboPrimaryAR.Text > -1 Then
                Me.txtPrimaryAccount.Text = Me.cboPrimaryAR.Text
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboPrimaryAR_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPrimaryAR.SelectedIndexChanged
        Try
            ' Me.txtPrimaryAccount.Visible = True
            If IsNumeric(Me.cboPrimaryAR.Text) AndAlso Me.cboPrimaryAR.Text > -1 Then
                Me.txtPrimaryAccount.Text = Me.cboPrimaryAR.Text
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FromCutOff_SalesBalanceReturnManual(ByVal vUserCode As String, ByVal vPrimCust As Double, ByVal vInvNo As String, _
   ByVal vInvDate As Date, ByVal vDueDate As Date, ByVal vBalance As Decimal, ByVal vPrimaryAR As Double) As Boolean
        Try
            Me.ObjProcess = New ClsTransaksiSalesReturnBalance

            Me.baru = True

            With Me.ObjProcess
                .Primary_Customer = vPrimCust
                .AOSRNumber = vInvNo
                .AOSRDate = vInvDate
                .SalesReturn = 0
                .VATReturn = 0
                .Total = 0
                .DueDate = vDueDate
                .Paid = 0
                .Dutystamp = 0
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
End Class