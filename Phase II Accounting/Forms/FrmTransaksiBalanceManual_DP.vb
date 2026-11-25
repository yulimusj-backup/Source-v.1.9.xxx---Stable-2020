Option Explicit On
Option Strict Off

Imports Core_Module
Public Class FrmTransaksiBalanceManual_DP
    Public UserCode As String = ""
    Public BackCOlorAll As Color = Color.LightSeaGreen
    Public vTipeDP As ClsTransaksiBalanceManual_DP.EnumTipeDP = ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_AOP

    Dim baru As Boolean = True
    Dim Obj As ClsTransaksiBalanceManual_DP
    Dim buffer As String = ""
    Dim ObjCD As New ClsCekData
    Dim ObjSupp As New clsMasterSupplier
    Dim ObjCust As New clsMasterCustomer
    Dim ObjCur As New clsMasterCurrency
    Dim ObjClosing As New ClsTransaksiClosing
    Dim ObjPO As New ClsTransaksiPO
    Dim ObjSO As New ClsTransaksiSalesOrder
    Dim ObjIdentifier As New CustomSMESS.ClsIdentifier
    Dim Def_Prec As Integer = 0

    Dim TBLCurr As New dseMasterCurrency.tblmaster_currencyDataTable
    Dim DseAR As New DseMultiAccountTableAdapters.tblisi_cbo_ar_akunTableAdapter
    Dim tblAR As New DseMultiAccount.tblisi_cbo_ar_akunDataTable
    Dim rowAR As DseMultiAccount.tblisi_cbo_ar_akunRow
    Dim DsAccountDP As New AccModule.DseMultiAccountTableAdapters.tblisi_cbo_ar_akunTableAdapter
    Dim tblAccountDP As New AccModule.DseMultiAccount.tblisi_cbo_ar_akunDataTable
    Dim rowAccountDP As AccModule.DseMultiAccount.tblisi_cbo_ar_akunRow

    'transaksi hanya bisa dibuat jika sebelum convertion date
    'pastikan tambah batasan edit/delete jika sudah di refer oleh transaksi jurnal purchase payment/sales receivement
    'pastikan enum terkait dengan beginning dp pada laporan ap-ar dan transaksi purchase payment/sales receivement

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryTransaksi.BackColor, False)
            Me.ObjCD.SetBackColor(Me.Panel1.Controls, Me.TxtPrimaryTransaksi.BackColor, False)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, True, True)
            ModCekData.KendalikanControl(Me.Panel1.Controls, True, True, True, True, True, True, True, True)

            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.cboAccount.Enabled = True

            If Me.baru = True Then
                Me.IsiCBOArAkunByTipePerkiraan1()
                Me.IsiCBOAkunDP()
            End If

            Me.TxtAOPDP_No.Focus()
            Me.TxtAOPDP_No.BackColor = Me.BackCOlorAll
            Me.cboCurrency.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, True, False)

            Me.CancelButton = Me.btnExit
            Me.cboAccount.Enabled = False
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.Panel1.Controls, Me.BackColor, True)
            Me.TxtPrimaryTransaksi.BackColor = Color.White
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
            Me.TxtPrimaryDocAcuan.Text = "-1"
            Me.TxtPrimaryPihakAcuan.Text = "-1"
            Me.TxtPrimaryCurr.Text = Core_Module.PrimaryBaseCurrency
            Me.TxtPrimaryAcc.Text = -1
            Me.TxtPrimaryAccDP.Text = -1

            Me.TxtAOPDP_No.Text = ""
            Me.TxtDocAcuanNo.Text = ""
            Me.TxtPihakAcuanCode.Text = ""
            Me.TxtPihakAcuanDesc.Text = ""
            Me.TxtSuppInvNo.Text = ""
            Me.DtpDate.Value = Now
            Me.DtpDueDate.Value = Now
            Me.cboCurrency.DataSource = Nothing
            Me.TxtBalance.Text = Me.ObjCD.DigitGrouping(0, Me.Def_Prec)
            Me.TxtExchangeRate.Text = ""
            Me.TxtRemarks.Text = ""

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
            Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping("1", Core_Module.DefPrecExchangeRate)
            Me.cboAccount.DataSource = Nothing
            Me.cboPrimaryAccount.DataSource = Nothing
            Me.CboDPAccount.DataSource = Nothing
            Me.cboPrimaryDPAccount.DataSource = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim ConvertionDate As Date
        Dim PihakAcuan As String = ""
        Dim NoTransaksi As String = ""
        Dim NoAcuan As String = ""

        Try
            If Me.vTipeDP = ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_AOP Then
                NoTransaksi = "AOP DP No"
                NoAcuan = "Purchase Order No"
                PihakAcuan = "Supplier"
                KoleksiNamaData.Add("Supplier Invoice No")
                KoleksiTextBox.Add(Me.TxtSuppInvNo)
            Else
                NoTransaksi = "Invoice DP No"
                NoAcuan = "Sales Order No"
                PihakAcuan = "Customer"
            End If

            'cek data yang harus diisi
            KoleksiNamaData.Add(NoTransaksi)
            KoleksiTextBox.Add(Me.TxtAOPDP_No)
            KoleksiNamaData.Add(NoAcuan)
            KoleksiTextBox.Add(Me.TxtDocAcuanNo)
            KoleksiNamaData.Add(PihakAcuan)
            KoleksiTextBox.Add(Me.TxtPihakAcuanCode)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data default precision harus non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'cek due date
            If Me.DtpDueDate.Value.Date < Me.DtpDate.Value.Date Then
                MsgBox("Please input due date <=  " & NoTransaksi & " date", MsgBoxStyle.Information, "Information")
                Me.DtpDueDate.Focus()
                Return False
            End If

            'cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.Obj.ApakahSudahAda(Me.TxtAOPDP_No.Text, Me.vTipeDP) = True Then
                    MsgBox("Duplication in " & NoTransaksi & ", please change your number", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtAOPDP_No.Focus()
                    Return False
                End If
            Else
                Dim TipeCodeLama As String = Me.Obj.GetDPNo(Me.TxtPrimaryTransaksi.Text)
                If TipeCodeLama <> Me.TxtAOPDP_No.Text Then
                    If Me.Obj.ApakahSudahAda(Me.TxtAOPDP_No.Text, Me.vTipeDP) = True Then
                        MsgBox("Duplication in  " & NoTransaksi & ", please change your number", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtAOPDP_No.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.Obj.GetTimeStamp(Me.TxtPrimaryTransaksi.Text)
                If Me.Obj.Timestamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.Obj.CariDetail(Me.Obj.Primary)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'kalo tanggal transaksi sama atau lebih besar dari convertion date maka ditolak
            ConvertionDate = Me.ObjClosing.GetConvertionDate
            If ConvertionDate <= Me.DtpDate.Value Then
                If ObjIdentifier.GetIdentifier = "smess_lmb" Then
                    If #1/1/2018# <= Me.DtpDate.Value.Date Then
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
                MsgBox("Please input balance with positive numbers.", MsgBoxStyle.Information, "Information")
                Me.TxtBalance.Focus()
                Return False
            End If
            If CDec(Me.TxtExchangeRate.Text) <= 0 Then
                MsgBox("Please input exchange rate with positive numbers.", MsgBoxStyle.Information, "Information")
                Me.TxtBalance.Focus()
                Return False
            End If

            'jangan sampai ngelebihin sub total dari document acuan(PO/SO)
            Dim SubTotal_DocAcuan As Decimal = Me.GetSubTotal_DocAcuan(Me.TxtPrimaryDocAcuan.Text, Me.vTipeDP)
            Dim Subtotal_DP As Decimal = Me.GetSubTotal(Me.TxtPrimaryDocAcuan.Text, Me.TxtPrimaryTransaksi.Text, Me.vTipeDP)
            Dim SubTotal_AOP_Inv_DP As Decimal = Me.GetSubTotal_AOP_Inv_DP(Me.TxtPrimaryDocAcuan.Text, Me.vTipeDP)
            If SubTotal_DocAcuan - SubTotal_AOP_Inv_DP - Subtotal_DP - CDec(Me.TxtBalance.Text) < 0 Then
                MsgBox("Maximum downpayment amount is " + Me.cboCurrency.Text + " " + _
                Me.ObjCD.DigitGrouping(SubTotal_DocAcuan - Subtotal_DP - SubTotal_AOP_Inv_DP, Me.Def_Prec), MsgBoxStyle.Information, "Information")

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

    Public Function GetSubTotal_DocAcuan(ByVal vPrimaryDocAcuan As Double, ByVal vTipeDP As ClsTransaksiBalanceManual_DP.EnumTipeDP) As Decimal
        Dim DTA As New DseTransaksiBalanceManual_DPTableAdapters.tbltrans_historicalbalancemanual_dp1TableAdapter
        Dim Subtotal_DocAcuan As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vTipeDP = ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_AOP Then
                Subtotal_DocAcuan = DTA.ScalarQuery_GetSubtotal_PO(vPrimaryDocAcuan)
            Else
                Subtotal_DocAcuan = DTA.ScalarQuery_GetSubtotal_SO(vPrimaryDocAcuan)
            End If

            If Subtotal_DocAcuan.HasValue = True Then
                Return Subtotal_DocAcuan.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSubTotal_AOP_Inv_DP(ByVal vPrimaryDocAcuan As Double, ByVal vTipeDP As ClsTransaksiBalanceManual_DP.EnumTipeDP) As Decimal
        Dim DTA As New DseTransaksiBalanceManual_DPTableAdapters.tbltrans_historicalbalancemanual_dp1TableAdapter
        Dim SubTotal_AOP_Inv_DP As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vTipeDP = ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_AOP Then
                SubTotal_AOP_Inv_DP = DTA.ScalarQuery_GetSubtotal_AOPDP(vPrimaryDocAcuan)
            Else
                SubTotal_AOP_Inv_DP = DTA.ScalarQuery_GetSubtotal_InvDP(vPrimaryDocAcuan)
            End If

            If SubTotal_AOP_Inv_DP.HasValue = True Then
                Return SubTotal_AOP_Inv_DP.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSubTotal(ByVal vPrimaryDocAcuan As Double, ByVal vPrimary As Double, _
    ByVal vTipeDP As ClsTransaksiBalanceManual_DP.EnumTipeDP) As Decimal
        Dim DTA As New DseTransaksiBalanceManual_DPTableAdapters.tbltrans_historicalbalancemanual_dp1TableAdapter
        Dim Subtotal As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Subtotal = DTA.ScalarQuery_GetSubtotalDP(vPrimary, vPrimaryDocAcuan, vTipeDP)
            If Subtotal.HasValue = True Then
                Return Subtotal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Sub IsiCBOArAkunByTipePerkiraan1()
        Dim tp1 As New ClsMasterAkun.EnumTipeAkun
        Dim tp2 As New ClsMasterAkun.EnumTipeAkun2
        Dim objMulti As New ClsMultiAccount

        Try
            If Me.vTipeDP = ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_AOP Then
                tp1 = Core_Module.ClsMasterAkun1.EnumTipeAkun.Liabilities
                tp2 = Core_Module.ClsMasterAkun1.EnumTipeAkun2.Account_Payable
            Else
                tp1 = Core_Module.ClsMasterAkun1.EnumTipeAkun.Asset
                tp2 = Core_Module.ClsMasterAkun1.EnumTipeAkun2.Account_Receivable
            End If

            Me.cboAccount.DataSource = Nothing
            Me.cboAccount.Items.Clear()
            Me.cboPrimaryAccount.DataSource = Nothing
            Me.cboPrimaryAccount.Items.Clear()

            objMulti.IsiTBLAPARAcount(tp1, tp2, Me.tblAR, Me.TxtPrimaryCurr.Text)

            Me.cboAccount.DataSource = Me.tblAR.DefaultView
            Me.cboAccount.DisplayMember = "Account Code"
            Me.cboPrimaryAccount.DataSource = Me.tblAR.DefaultView
            Me.cboPrimaryAccount.DisplayMember = "Primary"
            Me.TxtPrimaryAcc.Text = Me.cboPrimaryAccount.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCBOAkunDP()
        '- isicbo customer
        Dim tp1 As New Core_Module.ClsMasterAkun1.EnumTipeAkun
        Dim tp2 As New Core_Module.ClsMasterAkun1.EnumTipeAkun2
        Dim objMulti As New AccModule.ClsMultiAccount

        Try
            If Me.vTipeDP = ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_AOP Then
                tp1 = Core_Module.ClsMasterAkun1.EnumTipeAkun.Asset
                tp2 = Core_Module.ClsMasterAkun1.EnumTipeAkun2.Other_Receivable
            Else
                tp1 = Core_Module.ClsMasterAkun1.EnumTipeAkun.Liabilities
                tp2 = Core_Module.ClsMasterAkun1.EnumTipeAkun2.Other_Payable
            End If

            Me.CboDPAccount.DataSource = Nothing
            Me.CboDPAccount.Items.Clear()
            Me.cboPrimaryDPAccount.DataSource = Nothing
            Me.cboPrimaryDPAccount.Items.Clear()
            Me.tblAccountDP.Rows.Clear()

            objMulti.IsiTBLAPARAcount(tp1, tp2, Me.tblAccountDP, 1)

            If Me.vTipeDP = ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_Invoice Then
                'tmbahkan akun income
                Dim tblAccountDP1 As New DseMultiAccount.tblisi_cbo_ar_akunDataTable
                Dim rowAccountDP1 As DseMultiAccount.tblisi_cbo_ar_akunRow
                tp1 = Core_Module.ClsMasterAkun1.EnumTipeAkun.Income
                tp2 = Core_Module.ClsMasterAkun1.EnumTipeAkun2.Income

                objMulti.IsiTBLAPARAcount(tp1, tp2, tblAccountDP1, 1)

                For Each rowAccountDP1 In tblAccountDP1.Rows
                    rowAccountDP = Me.tblAccountDP.Newtblisi_cbo_ar_akunRow
                    For hj As Integer = 0 To Me.tblAccountDP.Columns.Count - 1
                        rowAccountDP.Item(hj) = rowAccountDP1.Item(hj)
                    Next
                    Me.tblAccountDP.Rows.Add(rowAccountDP)
                Next
            End If

            Me.CboDPAccount.DataSource = Me.tblAccountDP.DefaultView
            Me.CboDPAccount.DisplayMember = "Account Code"
            Me.cboPrimaryDPAccount.DataSource = Me.tblAccountDP.DefaultView
            Me.cboPrimaryDPAccount.DisplayMember = "Primary"
            Me.TxtPrimaryAccDP.Text = Me.cboPrimaryDPAccount.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeObjek()
        Try
            With Me.Obj
                .TipeDP = Me.vTipeDP
                .Primary = Me.TxtPrimaryTransaksi.Text
                .PrimaryDocAcuan = Me.TxtPrimaryDocAcuan.Text
                .PrimaryPihakAcuan = Me.TxtPrimaryPihakAcuan.Text
                .PrimaryCurrency = Me.TxtPrimaryCurr.Text
                .PrimaryAccount = Me.TxtPrimaryAcc.Text
                .PrimaryAccountDP = Me.TxtPrimaryAccDP.Text

                .DPNo = Me.TxtAOPDP_No.Text
                .InvNo = Me.TxtSuppInvNo.Text
                .DPDate = Me.DtpDate.Value
                .DueDate = Me.DtpDueDate.Value
                .Balance = Me.TxtBalance.Text
                .ExchangeRate = Me.TxtExchangeRate.Text
                .Remarks = Me.TxtRemarks.Text
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
            With Me.Obj
                Me.Def_Prec = Me.ObjCur.GetDefaultPrecission(.PrimaryCurrency)

                Me.TxtPrimaryTransaksi.Text = .Primary
                Me.TxtPrimaryDocAcuan.Text = .PrimaryDocAcuan
                Me.TxtPrimaryPihakAcuan.Text = .PrimaryPihakAcuan
                Me.TxtPrimaryCurr.Text = .PrimaryCurrency
                Me.TxtPrimaryAcc.Text = .PrimaryAccount
                Me.TxtPrimaryAccDP.Text = .PrimaryAccountDP

                Me.TxtAOPDP_No.Text = .DPNo
                If Me.vTipeDP = ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_AOP Then
                    Me.TxtDocAcuanNo.Text = Me.ObjPO.GetNoTransaksiPO(.PrimaryDocAcuan)
                Else
                    Me.TxtDocAcuanNo.Text = Me.ObjSO.GetNoSO(.PrimaryDocAcuan)
                End If
                'Me.TxtPihakAcuanCode.Text = ""
                'Me.TxtPihakAcuanDesc.Text = ""
                Me.GetPihakAcuan(.PrimaryDocAcuan)

                Me.TxtSuppInvNo.Text = .InvNo
                Me.DtpDate.Value = .DPDate
                Me.DtpDueDate.Value = .DueDate
                Me.cboCurrency.DataSource = Nothing
                Me.TxtBalance.Text = Me.ObjCD.DigitGrouping(.Balance, Me.Def_Prec)
                Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping(.ExchangeRate, Core_Module.DefPrecExchangeRate)
                Me.TxtRemarks.Text = .Remarks

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString

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
                Me.TxtPrimaryCurr.Text = .PrimaryCurrency

                Me.IsiCBOArAkunByTipePerkiraan1()
                If .PrimaryAccount > -1 Then
                    Me.TxtPrimaryAcc.Text = .PrimaryAccount
                    Me.cboPrimaryAccount.Text = .PrimaryAccount
                Else
                    Me.cboAccount.DataSource = Nothing
                    Me.cboPrimaryAccount.DataSource = Nothing
                    Me.TxtPrimaryAcc.Text = -1
                End If

                Me.IsiCBOAkunDP()
                If .PrimaryAccountDP > -1 Then
                    Me.cboPrimaryDPAccount.Text = .PrimaryAccountDP
                    Me.TxtPrimaryAccDP.Text = .PrimaryAccountDP
                Else
                    Me.CboDPAccount.DataSource = Nothing
                    Me.cboPrimaryDPAccount.DataSource = Nothing
                    Me.TxtPrimaryAccDP.Text = -1
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiBalanceManual_DP_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtAOPDP_No.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.Obj.SimpanBaru(Me.UserCode) = True Then
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
                            If Me.Obj.SimpanPerubahan(Me.UserCode) = True Then
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

    Private Sub FrmTransaksiPurchaseBalanceManual_DP_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.TxtRemarks.Multiline = True
            If Me.vTipeDP = ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_AOP Then
                Me.Text = "Historical Down Payment Purchase Balance"
                Me.LblDPNo.Text = "AOP DP No"
                Me.LblPONo.Text = "Purchase Order No"
                Me.LblPihakAcuan.Text = "Supplier"
                Me.LblInvNo.Text = "Supplier Invoice No"
                Me.LblAccount.Text = "AP Account"
            Else
                Dim HeigthAwal, HeigthAkhir As Integer
                HeigthAwal = Panel1.Location.Y
                Me.Panel1.Location = New Point(Panel1.Location.X, Panel1.Location.Y - (Panel1.Location.Y - Me.TxtSuppInvNo.Top))
                HeigthAkhir = Panel1.Location.Y
                Me.Panel1.Height += HeigthAwal - HeigthAkhir
                Me.LblInvNo.Visible = False
                Me.TxtSuppInvNo.Visible = False

                Me.Text = "Historical Down Payment Sales Balance"
                Me.LblDPNo.Text = "Invoice DP No"
                Me.LblPONo.Text = "Sales Order No"
                Me.LblPihakAcuan.Text = "Customer"
                Me.LblAccount.Text = "AR Account"
            End If

            Me.Obj = New ClsTransaksiBalanceManual_DP
            Me.baru = True
            Me.KosongkanField()
            Me.NonAktifkanDetail()

            Me.Def_Prec = Me.ObjCur.GetDecimal_DefaultCurrency
            Me.TxtBalance.TextAlign = HorizontalAlignment.Right
            Me.TxtAOPDP_No.CharacterCasing = CharacterCasing.Upper
            Me.TxtExchangeRate.TextAlign = HorizontalAlignment.Right
            'isikan cbo currency
            Me.ObjCur.IsiCBO(Me.cboCurrency, Me.TBLCurr)

            'update
            'Me.lblExcRate.Visible = False
            'Me.TxtExchangeRate.Visible = False
            Me.LblAccountDP.Visible = False
            Me.cboPrimaryDPAccount.Visible = False
            Me.CboDPAccount.Visible = False
            Me.CboDPAccount.TabStop = False

            Me.ShowIcon = False
            Me.TxtDocAcuanNo.CharacterCasing = CharacterCasing.Upper
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindDocAcuan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindDocAcuan.Click
        Dim SetDefault As Boolean = False

        Try
            If Me.vTipeDP = ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_AOP Then
                Dim i As New FrmCariTransaksiPO
                i.BackColorAll = Me.BackCOlorAll

                If Me.TxtDocAcuanNo.Text <> "" Then
                    i.txtCari.Text = Me.TxtDocAcuanNo.Text
                End If

                i.ShowDialog()
                If i.PrimaryPO > 0 Then
                    Me.TxtPrimaryDocAcuan.Text = i.PrimaryPO
                    Me.TxtDocAcuanNo.Text = Me.ObjPO.GetNoTransaksiPO(Me.TxtPrimaryDocAcuan.Text)
                    'isi pihak acuan
                    Me.GetPihakAcuan(Me.TxtPrimaryDocAcuan.Text)
                    Me.TxtSuppInvNo.Focus()
                Else
                    SetDefault = True
                End If

                i.Dispose()
            Else
                Dim i As New FrmCariTransaksiSO
                i.BackColorAll = Me.BackCOlorAll

                If Me.TxtDocAcuanNo.Text <> "" Then
                    i.txtCari.Text = Me.TxtDocAcuanNo.Text
                End If

                i.ShowDialog()
                If i.PrimaryTransaksi > 0 Then
                    Me.TxtPrimaryDocAcuan.Text = i.PrimaryTransaksi
                    Me.TxtDocAcuanNo.Text = Me.ObjSO.GetNoSO(Me.TxtPrimaryDocAcuan.Text)
                    'isi pihak acuan
                    Me.GetPihakAcuan(Me.TxtPrimaryDocAcuan.Text)
                    Me.TxtSuppInvNo.Focus()
                Else
                    SetDefault = True
                End If

                i.Dispose()
            End If

            If Me.Obj.ApakahDocAcuan_SudahDigunakan(Me.TxtPrimaryDocAcuan.Text, Me.TxtPrimaryTransaksi.Text, Me.vTipeDP) = True Then
                SetDefault = True
            End If

            If SetDefault = True Then
                Me.TxtPrimaryDocAcuan.Text = -1
                Me.TxtDocAcuanNo.Text = ""
                Me.TxtPrimaryPihakAcuan.Text = -1
                Me.TxtPihakAcuanCode.Text = ""
                Me.TxtPihakAcuanDesc.Text = ""
                Me.TxtDocAcuanNo.Focus()

                Dim rowCurr As dseMasterCurrency.tblmaster_currencyRow
                For i As Integer = 0 To Me.TBLCurr.Rows.Count - 1
                    Me.cboCurrency.SelectedIndex = i
                    rowCurr = Me.TBLCurr.DefaultView(Me.cboCurrency.SelectedIndex).Row
                    If rowCurr.Primary = Me.ObjCur.GetPrimary("IDR") Then
                        Me.TxtPrimaryCurr.Text = rowCurr.Primary
                        Me.TxtBalance.Text = Me.ObjCD.DigitGrouping(Me.TxtBalance.Text, Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurr.Text))
                        Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping(Me.ObjCur.GetDefaultExhangeRate(Me.TxtPrimaryCurr.Text), Core_Module.DefPrecExchangeRate)
                        Me.IsiCBOArAkunByTipePerkiraan1()
                        Exit For
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDocAcuanNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDocAcuanNo.GotFocus, TxtAOPDP_No.GotFocus, TxtSuppInvNo.GotFocus, TxtBalance.GotFocus, TxtRemarks.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtPihakAcuanCode.BackColor
                        Else
                            Txt.BackColor = Me.BackCOlorAll
                        End If
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
                Me.TxtBalance.Text = Me.ObjCD.DigitGrouping(Me.TxtBalance.Text, Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurr.Text))
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
                Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping(Me.TxtExchangeRate.Text, Core_Module.DefPrecExchangeRate)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarks_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtRemarks.KeyDown
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDocAcuanNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDocAcuanNo.LostFocus, TxtAOPDP_No.LostFocus, TxtSuppInvNo.LostFocus, TxtBalance.LostFocus, TxtRemarks.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtPihakAcuanCode.BackColor
                        Else
                            Txt.BackColor = Me.TxtPrimaryTransaksi.BackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDocAcuanNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDocAcuanNo.KeyDown
        
        Try
            If e.KeyValue = 13 Then
                If Me.TxtDocAcuanNo.Text <> "" Then
                    If Me.vTipeDP = ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_AOP Then
                        If Me.ObjPO.ApakahSudahAda(Me.TxtDocAcuanNo.Text) = True Then
                            Me.TxtPrimaryDocAcuan.Text = Me.ObjPO.GetPrimary(Me.TxtDocAcuanNo.Text)
                            'isi pihak acuan
                            Me.GetPihakAcuan(Me.TxtPrimaryDocAcuan.Text)
                            Me.TxtSuppInvNo.Focus()
                        Else
                            Me.BtnFindDocAcuan.PerformClick()
                        End If
                    Else
                        If Me.ObjSO.ApakahSudahAda(Me.TxtDocAcuanNo.Text) = True Then
                            Me.TxtPrimaryDocAcuan.Text = Me.ObjSO.GetPrimary(Me.TxtDocAcuanNo.Text)
                            'isi pihak acuan
                            Me.GetPihakAcuan(Me.TxtPrimaryDocAcuan.Text)
                            Me.TxtSuppInvNo.Focus()
                        Else
                            Me.BtnFindDocAcuan.PerformClick()
                        End If
                    End If
                Else
                    Me.BtnFindDocAcuan.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDocAcuanNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDocAcuanNo.Leave
        
        Try
            If Me.TxtDocAcuanNo.Text <> "" Then
                If Me.vTipeDP = ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_AOP Then
                    Dim i As New FrmCariTransaksiPO
                    i.BackColorAll = Me.BackCOlorAll

                    If Me.ObjPO.ApakahSudahAda(Me.TxtDocAcuanNo.Text) = True Then
                        Me.TxtPrimaryDocAcuan.Text = Me.ObjPO.GetPrimary(Me.TxtDocAcuanNo.Text)
                        Me.TxtDocAcuanNo.Text = Me.ObjPO.GetNoTransaksiPO(Me.TxtPrimaryDocAcuan.Text)
                        'isi pihak acuan
                        Me.GetPihakAcuan(Me.TxtPrimaryDocAcuan.Text)
                    Else
                        If Me.TxtDocAcuanNo.Text <> "" Then
                            i.txtCari.Text = Me.TxtDocAcuanNo.Text
                        End If

                        i.ShowDialog()
                        If i.PrimaryPO > 0 Then
                            Me.TxtPrimaryDocAcuan.Text = i.PrimaryPO
                            Me.TxtDocAcuanNo.Text = Me.ObjPO.GetNoTransaksiPO(Me.TxtPrimaryDocAcuan.Text)
                            'isi pihak acuan
                            Me.GetPihakAcuan(Me.TxtPrimaryDocAcuan.Text)
                        Else
                            GoTo setDef
                        End If
                    End If

                    i.Dispose()
                Else
                    Dim i As New FrmCariTransaksiSO
                    i.BackColorAll = Me.BackCOlorAll

                    If Me.ObjSO.ApakahSudahAda(Me.TxtDocAcuanNo.Text) = True Then
                        Me.TxtPrimaryDocAcuan.Text = Me.ObjSO.GetPrimary(Me.TxtDocAcuanNo.Text)
                        Me.TxtDocAcuanNo.Text = Me.ObjSO.GetNoSO(Me.TxtPrimaryDocAcuan.Text)
                        'isi pihak acuan
                        Me.GetPihakAcuan(Me.TxtPrimaryDocAcuan.Text)
                    Else
                        If Me.TxtDocAcuanNo.Text <> "" Then
                            i.txtCari.Text = Me.TxtDocAcuanNo.Text
                        End If

                        i.ShowDialog()
                        If i.PrimaryTransaksi > 0 Then
                            Me.TxtPrimaryDocAcuan.Text = i.PrimaryTransaksi
                            Me.TxtDocAcuanNo.Text = Me.ObjSO.GetNoSO(Me.TxtPrimaryDocAcuan.Text)
                            'isi pihak acuan
                            Me.GetPihakAcuan(Me.TxtPrimaryDocAcuan.Text)
                        Else
                            GoTo setDef
                        End If
                    End If

                    i.Dispose()
                End If
            Else
setDef:
                Me.TxtPrimaryDocAcuan.Text = -1
                Me.TxtDocAcuanNo.Text = ""
                Me.TxtPrimaryPihakAcuan.Text = -1
                Me.TxtPihakAcuanCode.Text = ""
                Me.TxtPihakAcuanDesc.Text = ""

                Dim rowCurr As dseMasterCurrency.tblmaster_currencyRow
                For i As Integer = 0 To Me.TBLCurr.Rows.Count - 1
                    Me.cboCurrency.SelectedIndex = i
                    rowCurr = Me.TBLCurr.DefaultView(Me.cboCurrency.SelectedIndex).Row
                    If rowCurr.Primary = Me.ObjCur.GetPrimary("IDR") Then
                        Me.TxtPrimaryCurr.Text = rowCurr.Primary
                        Me.TxtBalance.Text = Me.ObjCD.DigitGrouping(Me.TxtBalance.Text, Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurr.Text))
                        Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping(Me.ObjCur.GetDefaultExhangeRate(Me.TxtPrimaryCurr.Text), Core_Module.DefPrecExchangeRate)
                        Me.IsiCBOArAkunByTipePerkiraan1()
                        Exit For
                    End If
                Next

                Exit Sub
            End If

            If Me.Obj.ApakahDocAcuan_SudahDigunakan(Me.TxtPrimaryDocAcuan.Text, Me.TxtPrimaryTransaksi.Text, Me.vTipeDP) = True Then
                GoTo setDef
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetPihakAcuan(ByVal vPrimaryDocAcuan As Double)
        Dim setDefault As Boolean = False
        Dim PrimaryCurrFromDocAcuan As Double = -1
        Dim rowCurr As dseMasterCurrency.tblmaster_currencyRow

        Try
            If Me.vTipeDP = ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_AOP Then
                Dim dta As New DseTransaksiBalanceManual_DPTableAdapters.tbl_suppTableAdapter
                Dim tbl As New DseTransaksiBalanceManual_DP.tbl_suppDataTable
                Dim row As DseTransaksiBalanceManual_DP.tbl_suppRow

                dta.Connection.ConnectionString = core_module.getconstring
                dta.Fill(tbl, vPrimaryDocAcuan)
                If tbl.Count > 0 Then
                    row = tbl.Rows(0)
                    Me.TxtPrimaryPihakAcuan.Text = row.Primary_Supplier
                    Me.TxtPihakAcuanCode.Text = row.Supplier_Code
                    Me.TxtPihakAcuanDesc.Text = row.Supplier_Name
                    PrimaryCurrFromDocAcuan = row.Primary_Currency
                Else
                    setDefault = True
                End If

                dta.Dispose()
                tbl.Dispose()
            Else
                Dim dta As New DseTransaksiBalanceManual_DPTableAdapters.tbl_custTableAdapter
                Dim tbl As New DseTransaksiBalanceManual_DP.tbl_custDataTable
                Dim row As DseTransaksiBalanceManual_DP.tbl_custRow

                dta.Connection.ConnectionString = core_module.getconstring
                dta.Fill(tbl, vPrimaryDocAcuan)
                If tbl.Count > 0 Then
                    row = tbl.Rows(0)
                    Me.TxtPrimaryPihakAcuan.Text = row.Primary_Cust
                    Me.TxtPihakAcuanCode.Text = row.Customer_Code
                    Me.TxtPihakAcuanDesc.Text = row.Customer_Name
                    PrimaryCurrFromDocAcuan = row.Primary_Currency
                Else
                    setDefault = True
                End If

                dta.Dispose()
                tbl.Dispose()
            End If

            If setDefault = True Then
                Me.TxtPrimaryPihakAcuan.Text = -1
                Me.TxtPihakAcuanCode.Text = ""
                Me.TxtPihakAcuanDesc.Text = ""
            Else
                For i As Integer = 0 To Me.TBLCurr.Rows.Count - 1
                    Me.cboCurrency.SelectedIndex = i
                    rowCurr = Me.TBLCurr.DefaultView(Me.cboCurrency.SelectedIndex).Row
                    If rowCurr.Primary = PrimaryCurrFromDocAcuan Then
                        Me.TxtPrimaryCurr.Text = rowCurr.Primary
                        Me.TxtBalance.Text = Me.ObjCD.DigitGrouping(Me.TxtBalance.Text, Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurr.Text))
                        Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping(Me.ObjCur.GetDefaultExhangeRate(Me.TxtPrimaryCurr.Text), Core_Module.DefPrecExchangeRate)
                        Me.IsiCBOArAkunByTipePerkiraan1()
                        Exit For
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Not Me.TxtAOPDP_No.Text = "" Then
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

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.Obj.SimpanBaru(Me.UserCode) Then
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
                        If Me.Obj.SimpanPerubahan(Me.UserCode) Then
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

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmTransaksiBalanceManual_DP_Cari
        Try
            ij.BackColorAll = Me.BackCOlorAll
            ij.vTipeDP = Me.vTipeDP
            ij.ShowDialog()
            If Not ij.mPrimaryDept = -1 Then
                If Me.Obj.CariDetail(ij.mPrimaryDept) = True Then
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

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Me.baru = True
        Me.KosongkanField()
        Me.AktifkanDetail()
        Me.TxtPrimaryTransaksi.Text = "-1"
        Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
        Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
        Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtAOPDP_No.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtAOPDP_No.Text <> "" Then
                If Me.Obj.ApakahBolehEdit(Me.TxtPrimaryTransaksi.Text, Me.UserCode, Me.vTipeDP) = True Then
                    Me.baru = False
                    Me.AktifkanDetail()
                    Me.TxtPihakAcuanCode.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtAOPDP_No.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtAOPDP_No.Text <> "" Then
                'khusus untuk IDR, tidak bisa dihapus, karena digunakan oleh sistem
                If Me.Obj.ApakahBolehHapus(Me.TxtPrimaryTransaksi.Text, Me.UserCode, Me.vTipeDP) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.Hapus(Me.TxtPrimaryTransaksi.Text, Me.UserCode, Me.vTipeDP) = True Then
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

    Private Sub cboCurrency_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCurrency.SelectedIndexChanged
        Dim row As dseMasterCurrency.tblmaster_currencyRow

        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.cboCurrency.SelectedIndex = -1 Then
                    Me.TxtPrimaryCurr.Text = -1
                    Me.TxtBalance.Text = Me.ObjCD.DigitGrouping(Me.TxtBalance.Text, Me.ObjCur.GetDecimal_DefaultCurrency)
                    Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping("1", Core_Module.DefPrecExchangeRate)
                    Me.IsiCBOArAkunByTipePerkiraan1()
                Else
                    row = Me.TBLCurr.DefaultView(Me.cboCurrency.SelectedIndex).Row
                    Me.TxtPrimaryCurr.Text = row.Primary
                    Me.TxtBalance.Text = Me.ObjCD.DigitGrouping(Me.TxtBalance.Text, Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurr.Text))
                    Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping(Me.ObjCur.GetDefaultExhangeRate(Me.TxtPrimaryCurr.Text), Core_Module.DefPrecExchangeRate)
                    Me.IsiCBOArAkunByTipePerkiraan1()
                End If

                Me.SetExchangeRate()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboAccount_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboAccount.SelectedIndexChanged

        Try
            'Me.txtPrimaryAccount.Visible = True
            If IsNumeric(Me.cboPrimaryAccount.Text) AndAlso Me.cboPrimaryAccount.Text > -1 Then
                Me.TxtPrimaryAcc.Text = Me.cboPrimaryAccount.Text
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboPrimaryAccount_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPrimaryAccount.SelectedValueChanged

        Try
            'Me.txtPrimaryAccount.Visible = True
            If IsNumeric(Me.cboPrimaryAccount.Text) AndAlso Me.cboPrimaryAccount.Text > -1 Then
                Me.TxtPrimaryAcc.Text = Me.cboPrimaryAccount.Text
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboDPAccount_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboDPAccount.SelectedIndexChanged

        Try
            'Me.txtPrimaryAccount.Visible = True
            If IsNumeric(Me.cboPrimaryDPAccount.Text) AndAlso Me.cboPrimaryDPAccount.Text > -1 Then
                Me.TxtPrimaryAccDP.Text = Me.cboPrimaryDPAccount.Text
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboPrimaryDPAccount_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPrimaryDPAccount.SelectedValueChanged

        Try
            'Me.txtPrimaryAccount.Visible = True
            If IsNumeric(Me.cboPrimaryDPAccount.Text) AndAlso Me.cboPrimaryDPAccount.Text > -1 Then
                Me.TxtPrimaryAccDP.Text = Me.cboPrimaryDPAccount.Text
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetExchangeRate()
        Try
            If Me.TxtPrimaryCurr.Text = "" Or Me.TxtPrimaryCurr.Text = "-1" Then
                Me.TxtExchangeRate.ReadOnly = True
                Me.TxtExchangeRate.BackColor = Me.TxtPihakAcuanCode.BackColor
                Me.TxtExchangeRate.TabStop = False
            ElseIf Me.TxtPrimaryCurr.Text = Core_Module.PrimaryBaseCurrency.ToString Then
                Me.TxtExchangeRate.ReadOnly = True
                Me.TxtExchangeRate.BackColor = Me.TxtPihakAcuanCode.BackColor
                Me.TxtExchangeRate.TabStop = False
            Else
                Me.TxtExchangeRate.ReadOnly = False
                Me.TxtExchangeRate.BackColor = Me.TxtPrimaryCurr.BackColor
                Me.TxtExchangeRate.TabStop = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class