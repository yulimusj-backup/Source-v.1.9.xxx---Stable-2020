Option Explicit On
Option Strict Off

Public Class FrmMasterCurrency2
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim baru As Boolean
    Dim objCurrency As clsMasterCurrency
    Dim buffer As String = ""
    Dim ObjCD As New ClsCekData
    Dim ObjLicense As New ClsLicense
    Dim ObjAkun As New ClsMasterAkun1

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryCurrency.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimaryCurrency.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimaryCurrency.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.TxtPrimaryCurrency.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.TxtPrimaryCurrency.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Me.TxtPrimaryCurrency.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Me.TxtPrimaryCurrency.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel1.Controls, Me.TxtPrimaryCurrency.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel2.Controls, Me.TxtPrimaryCurrency.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel1.Controls, Me.TxtPrimaryCurrency.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel2.Controls, Me.TxtPrimaryCurrency.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer6.Panel1.Controls, Me.TxtPrimaryCurrency.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer6.Panel2.Controls, Me.TxtPrimaryCurrency.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer7.Panel1.Controls, Me.TxtPrimaryCurrency.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer7.Panel2.Controls, Me.TxtPrimaryCurrency.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer8.Panel1.Controls, Me.TxtPrimaryCurrency.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer8.Panel2.Controls, Me.TxtPrimaryCurrency.BackColor, False)

            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer4.Panel1.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer4.Panel2.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer5.Panel1.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer5.Panel2.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer6.Panel1.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer6.Panel2.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer7.Panel1.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer7.Panel2.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer8.Panel1.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer8.Panel2.Controls, True, True, True, True, False, False, False, False)

            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None

            If Me.baru = True Then
                Me.gboEdit.Visible = False
                Me.gboNew.Visible = True
            Else
                Me.gboEdit.Visible = True
                Me.gboNew.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer4.Panel1.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer4.Panel2.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer5.Panel1.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer5.Panel2.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer6.Panel1.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer6.Panel2.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer7.Panel1.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer7.Panel2.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer8.Panel1.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer8.Panel2.Controls, True, False, True, False, False, False, False, False)

            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)

            Me.gboNew.Visible = False
            Me.gboEdit.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.SplitContainer3.Panel1.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.SplitContainer3.Panel2.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.SplitContainer4.Panel1.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.SplitContainer4.Panel2.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.SplitContainer5.Panel1.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.SplitContainer5.Panel2.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.SplitContainer6.Panel1.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.SplitContainer6.Panel2.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.SplitContainer7.Panel1.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.SplitContainer7.Panel2.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.SplitContainer8.Panel1.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.SplitContainer8.Panel2.Controls, True, False, False)

            Me.TxtPrimaryCurrency.Text = "-1"
            Me.TxtPrimaryPayableAccount.Text = "-1"
            Me.TxtPrimaryRealizedExchangeAcc.Text = "-1"
            Me.TxtPrimaryReceivableAcc.Text = "-1"
            Me.TxtPrimaryUnrealizedExchangeAcc.Text = "-1"
            Me.TxtPrimaryReceivableAccountHeader.Text = "-1"
            Me.TxtPrimaryPayableAccountHeader.Text = "-1"
            Me.TxtPrimaryRealiazedGainAccount1.Text = "-1"
            Me.TxtPrimaryRealiazedGainAccount1.Text = "-1"

            Me.txtDefaultPrecision.Text = "0"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim TipeAkun2 As ClsMasterAkun1.EnumTipeAkun2 = ClsMasterAkun1.EnumTipeAkun2.Account_Payable
        Dim DefPrecLama As Integer = 0

        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("currency code")
            KoleksiTextBox.Add(Me.TxtCurrencyCode)
            KoleksiNamaData.Add("currency name")
            KoleksiTextBox.Add(Me.TxtCurrencyName)

            If Me.baru = False Then
                KoleksiNamaData.Add("payable account")
                KoleksiTextBox.Add(Me.TxtPayableAccountCode)
                KoleksiNamaData.Add("receivables account")
                KoleksiTextBox.Add(Me.TxtReceivableAcc)
                KoleksiNamaData.Add("unrealized exchange gain/loss account")
                KoleksiTextBox.Add(Me.TxtUnRealizedExchangeGainLossAccount)
                KoleksiNamaData.Add("realized exchange gain/loss account")
                KoleksiTextBox.Add(Me.TxtRealizedExchangeGainLossAccount)
            Else
                KoleksiNamaData.Add("payable account header")
                KoleksiTextBox.Add(Me.TxtPayableAccountHeaderCode)
                KoleksiNamaData.Add("payable account code")
                KoleksiTextBox.Add(Me.TxtPayableAccountHeaderNew2)
                KoleksiNamaData.Add("payable account description")
                KoleksiTextBox.Add(Me.TxtPayableAccountHeaderNewDesc)

                KoleksiNamaData.Add("receivable account header")
                KoleksiTextBox.Add(Me.TxtReceivableAccountHeaderCode)
                KoleksiNamaData.Add("receivable account code")
                KoleksiTextBox.Add(Me.TxtReceivableAccountHeaderNew2)
                KoleksiNamaData.Add("receivable account description")
                KoleksiTextBox.Add(Me.TxtReceivableAccountHeaderNewDesc)

                KoleksiNamaData.Add("unrealized gain/loss account header")
                KoleksiTextBox.Add(Me.TxtUnrealizedGainAccCode1)
                KoleksiNamaData.Add("realized gain/loss account header")
                KoleksiTextBox.Add(Me.TxtRealiazedGainAccountCode1)
            End If

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data default precision harus non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()
            KoleksiNamaData.Add("default precision")
            KoleksiTextBox.Add(Me.txtDefaultPrecision)
            If ModCekData.CekNonNegatif(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.objCurrency.ApakahSudahAda(Me.TxtCurrencyCode.Text) = True Then
                    MsgBox("Duplication in currency code, please change your code", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtCurrencyCode.Focus()
                    Return False
                End If
            Else
                Dim ProductGroupLama As String = Me.objCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
                If ProductGroupLama <> Me.TxtCurrencyCode.Text Then
                    If Me.objCurrency.ApakahSudahAda(Me.TxtCurrencyCode.Text) = True Then
                        MsgBox("Duplication in currency code, please change your code", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtCurrencyCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.objCurrency.GetTimeStamp(Me.TxtPrimaryCurrency.Text)
                If Me.objCurrency.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.objCurrency.CariDetail(Me.objCurrency.Primary)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'maksimal default precision 4
            If CInt(Me.txtDefaultPrecision.Text) > 4 Then
                MsgBox("Maiximum value for default precision is 4.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If Me.baru = False Then
                'untuk receivable account, tipe akun 2 harus account receivable + mata uangnya harus sama
                If Me.ObjAkun.GetTipeAccount2(Me.TxtPrimaryReceivableAcc.Text) <> ClsMasterAkun1.EnumTipeAkun2.Account_Receivable Then
                    MsgBox("Please specify account receivable type for account receivable code.", MsgBoxStyle.Information, "Information")
                    Me.TxtReceivableAcc.Focus()
                    Return False
                End If
                If Me.ObjAkun.GetPrimaryCurrency(Me.TxtPrimaryReceivableAcc.Text) <> Me.TxtPrimaryCurrency.Text Then
                    MsgBox("Please specify account with " + Me.TxtCurrencyCode.Text + " currency for account receivable code.", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtReceivableAcc.Focus()
                    Return False
                End If

                'untuk payable account, tipe akun harus account payable + mata uangnya harus sama
                If Me.ObjAkun.GetTipeAccount2(Me.TxtPrimaryPayableAccount.Text) <> ClsMasterAkun1.EnumTipeAkun2.Account_Payable Then
                    MsgBox("Please specify account payable type for account payable code.", MsgBoxStyle.Information, "Information")
                    Me.TxtPayableAccountCode.Focus()
                    Return False
                End If
                If Me.ObjAkun.GetPrimaryCurrency(Me.TxtPrimaryPayableAccount.Text) <> Me.TxtPrimaryCurrency.Text Then
                    MsgBox("Please specify account with " + Me.TxtCurrencyCode.Text + " currency for account payable code.", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtPayableAccountCode.Focus()
                    Return False
                End If

                'untuk unrealized exchange gain/loss, tipe accountnya harus other expense/other income
                TipeAkun2 = Me.ObjAkun.GetTipeAccount2(Me.TxtPrimaryUnrealizedExchangeAcc.Text)
                If TipeAkun2 <> ClsMasterAkun1.EnumTipeAkun2.Other_Expense And _
                TipeAkun2 <> ClsMasterAkun1.EnumTipeAkun2.Other_Income And _
                TipeAkun2 <> ClsMasterAkun1.EnumTipeAkun2.Expense Then
                    MsgBox("Please specify expense, other expense and other income type for unrealized exchange gain/loss code.", MsgBoxStyle.Information, "Information")
                    Me.TxtUnRealizedExchangeGainLossAccount.Focus()
                    Return False
                End If

                'untuk realized exchange gain/loss, tipe accountnya harus other expense/other income
                TipeAkun2 = Me.ObjAkun.GetTipeAccount2(Me.TxtPrimaryRealizedExchangeAcc.Text)
                If TipeAkun2 <> ClsMasterAkun1.EnumTipeAkun2.Other_Expense And _
                TipeAkun2 <> ClsMasterAkun1.EnumTipeAkun2.Other_Income And _
                TipeAkun2 <> ClsMasterAkun1.EnumTipeAkun2.Expense Then
                    MsgBox("Please specify expense, other expense and other income type for realized exchange gain/loss code.", MsgBoxStyle.Information, "Information")
                    Me.TxtRealizedExchangeGainLossAccount.Focus()
                    Return False
                End If
            Else
                'untuk receivable account, tipe akun 2 harus account receivable + mata uangnya harus sama
                TipeAkun2 = Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryReceivableAccountHeader.Text)
                If TipeAkun2 <> ClsMasterAkun1.EnumTipeAkun.Asset Then
                    MsgBox("Please specify asset type for account receivable code.", MsgBoxStyle.Information, "Information")
                    Me.TxtReceivableAccountHeaderCode.Focus()
                    Return False
                End If

                'untuk payable account, tipe akun harus account payable + mata uangnya harus sama
                TipeAkun2 = Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryPayableAccountHeader.Text)
                If TipeAkun2 <> ClsMasterAkun1.EnumTipeAkun.Liabilities Then
                    MsgBox("Please specify liabilities type for account payable code.", MsgBoxStyle.Information, "Information")
                    Me.TxtPayableAccountHeaderCode.Focus()
                    Return False
                End If

                'untuk unrealized exchange gain/loss, tipe accountnya harus other expense/other income
                TipeAkun2 = Me.ObjAkun.GetTipeAccount2(Me.TxtPrimaryUnrealiazedGainAccount1.Text)
                If TipeAkun2 <> ClsMasterAkun1.EnumTipeAkun2.Other_Expense And _
                TipeAkun2 <> ClsMasterAkun1.EnumTipeAkun2.Other_Income And _
                TipeAkun2 <> ClsMasterAkun1.EnumTipeAkun2.Expense Then
                    MsgBox("Please specify expense, other expense and other income type for unrealized exchange gain/loss code.", MsgBoxStyle.Information, "Information")
                    Me.TxtUnrealizedGainAccCode1.Focus()
                    Return False
                End If

                'untuk realized exchange gain/loss, tipe accountnya harus other expense/other income
                TipeAkun2 = Me.ObjAkun.GetTipeAccount2(Me.TxtPrimaryRealiazedGainAccount1.Text)
                If TipeAkun2 <> ClsMasterAkun1.EnumTipeAkun2.Other_Expense And _
                TipeAkun2 <> ClsMasterAkun1.EnumTipeAkun2.Other_Income And _
                TipeAkun2 <> ClsMasterAkun1.EnumTipeAkun2.Expense Then
                    MsgBox("Please specify expense, other expense and other income type for realized exchange gain/loss code.", MsgBoxStyle.Information, "Information")
                    Me.TxtRealiazedGainAccountCode1.Focus()
                    Return False
                End If
            End If

            'pastikan kode akun belum ada, jika sudah ada maka tidak boleh diinput, khusus untuk input data baru
            If Me.baru = True Then
                If Me.ObjAkun.ApakahSudahAda(Me.TxtPayableAccountHeaderNew1.Text + Me.TxtPayableAccountHeaderNew2.Text) = True Then
                    MsgBox("Duplication in account code data, please specify unique account code.", MsgBoxStyle.Information, "Information")
                    Me.TxtPayableAccountHeaderNew2.Focus()
                    Return False
                End If
                If Me.ObjAkun.ApakahSudahAda(Me.TxtReceivableAccountHeaderNew1.Text + Me.TxtReceivableAccountHeaderNew2.Text) = True Then
                    MsgBox("Duplication in account code data, please specify unique account code.", MsgBoxStyle.Information, "Information")
                    Me.TxtReceivableAccountHeaderNew2.Focus()
                    Return False
                End If

                'pastikan kode akun tidak sama satu dengan yang lainnya
                'akun payable
                If Me.TxtPayableAccountHeaderNew1.Text + Me.TxtPayableAccountHeaderNew2.Text = Me.TxtReceivableAccountHeaderNew1.Text + Me.TxtReceivableAccountHeaderNew2.Text Then
                    MsgBox("Please specify unique account code.", MsgBoxStyle.Information, "Information")
                    Me.TxtPayableAccountHeaderNew1.Focus()
                    Return False
                End If
                'akun receivbales
            End If

            'kalo nilai defaul precisionnya turun, maka harus dicek, apakah ada nilai yang dilanggar, sehingga menyebabkan sistem menjadi tidak stabil.
            'msgbox("belum")
            If Me.baru = False Then
                DefPrecLama = Me.objCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
                If CInt(Me.txtDefaultPrecision.Text) < DefPrecLama Then
                    Dim ij As New FrmMasterCurrency2_CekDefPrec
                    ij.PrimaryCurrency = Me.TxtPrimaryCurrency.Text
                    ij.DefPrecBaru = Me.txtDefaultPrecision.Text
                    If ij.ShowDialog <> Windows.Forms.DialogResult.OK Then
                        Return False
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

    Private Sub SalinKeObjek()
        Try
            With Me.objCurrency
                .CurrencyCode = Me.TxtCurrencyCode.Text
                .Description = Me.TxtCurrencyName.Text
                .DefaultPrecision = Me.txtDefaultPrecision.Text
                .PrimaryPayableAccount = Me.TxtPrimaryPayableAccount.Text
                .PrimaryReceivableAccount = Me.TxtPrimaryReceivableAcc.Text

                If Me.baru = False Then
                    .PrimaryRealizedExchangeGainLossAccount = Me.TxtPrimaryRealizedExchangeAcc.Text
                    .PrimaryUnrealizedExchangeGainLossAccount = Me.TxtPrimaryUnrealizedExchangeAcc.Text
                Else
                    .PrimaryRealizedExchangeGainLossAccount = Me.TxtPrimaryRealiazedGainAccount1.Text
                    .PrimaryUnrealizedExchangeGainLossAccount = Me.TxtPrimaryUnrealiazedGainAccount1.Text
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.objCurrency
                Me.TxtCurrencyCode.Text = .CurrencyCode
                Me.TxtCurrencyName.Text = .Description
                Me.TxtPrimaryCurrency.Text = .Primary
                Me.txtDefaultPrecision.Text = .DefaultPrecision
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString

                Me.TxtPrimaryPayableAccount.Text = .PrimaryPayableAccount
                Me.TxtPayableAccountCode.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryPayableAccount.Text)
                Me.TxtPayableAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryPayableAccount.Text)
                Me.TxtPrimaryReceivableAcc.Text = .PrimaryReceivableAccount
                Me.TxtReceivableAcc.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryReceivableAcc.Text)
                Me.TxtReceivableAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryReceivableAcc.Text)
                Me.TxtPrimaryUnrealizedExchangeAcc.Text = .PrimaryUnrealizedExchangeGainLossAccount
                Me.TxtUnRealizedExchangeGainLossAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryUnrealizedExchangeAcc.Text)
                Me.TxtUnrealizedExchangeGainLossAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryUnrealizedExchangeAcc.Text)
                Me.TxtPrimaryRealizedExchangeAcc.Text = .PrimaryRealizedExchangeGainLossAccount
                Me.TxtRealizedExchangeGainLossAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryRealizedExchangeAcc.Text)
                Me.TxtRealizedExchangeGainLossAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryRealizedExchangeAcc.Text)
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.objCurrency = New clsMasterCurrency
            Me.baru = True
            Me.NonAktifkanDetail()
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
                        If Me.objCurrency.SimpanBaru(Me.UserCode, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.objCurrency.SimpanPerubahan(Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
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
            If Not Me.TxtCurrencyCode.Text = "" And Me.TxtCurrencyName.Text <> "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
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

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.baru = True
            Me.AktifkanDetail()
            Me.KosongkanField()
            Me.TxtCurrencyCode.Focus()
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

            'Untuk data baru, header nya aset untuk akun asset
            Me.TxtPrimaryPayableAccountHeader.Text = 2
            Me.TxtPayableAccountHeaderCode.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryPayableAccountHeader.Text)
            Me.TxtPayableAccountHeaderDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryPayableAccountHeader.Text)
            Me.TxtPrimaryReceivableAccountHeader.Text = 1
            Me.TxtReceivableAccountHeaderCode.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryReceivableAccountHeader.Text)
            Me.TxtReceivableAccountHeaderDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryReceivableAccountHeader.Text)

            Me.SetAccountPayableNew()
            Me.SetAccountReceivablesNew()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmCariMasterCurrency

        Try
            ij.BackColorAll = Me.BackColorAll
            ij.ShowDialog()
            If Not ij.PrimaryCurrency = -1 Then
                If Me.objCurrency.CariDetail(ij.PrimaryCurrency) = True Then
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

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        If Me.TxtCurrencyCode.Text = "" Then
            Me.btnFind.PerformClick()
        End If

        If Me.TxtCurrencyCode.Text <> "" Then
            Me.baru = False
            Me.AktifkanDetail()
            Me.TxtCurrencyCode.Focus()
            'jika IDR, tidak bisa diedit dan didelete
            If Me.TxtCurrencyCode.Text = "IDR" Then
                Me.TxtCurrencyCode.Enabled = False
                Me.TxtCurrencyName.Focus()
            End If
        End If
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtCurrencyCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtCurrencyCode.Text <> "" Then
                'khusus untuk IDR, tidak bisa dihapus, karena digunakan oleh sistem
                If Me.TxtCurrencyCode.Text = "IDR" Or Me.TxtPrimaryCurrency.Text = "1" Then
                    MsgBox("This the system default currency, you cannot delete this data", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                If Me.objCurrency.ApakahBolehHapus(Me.TxtPrimaryCurrency.Text) = False Then
                    'MsgBox("Belum")
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.objCurrency.Hapus(Me.TxtPrimaryCurrency.Text) = True Then
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
            If Me.TxtCurrencyCode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objCurrency.SimpanBaru(Me.UserCode, Me) = True Then
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
                            If Me.objCurrency.SimpanPerubahan(Me.UserCode) = True Then
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

    Private Sub txtDefaultPrecision_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDefaultPrecision.Enter
        Try
            Me.buffer = Me.txtDefaultPrecision.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDefaultPrecision_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDefaultPrecision.Leave
        Try
            If ModCekData.CekNonNegatif(Me.txtDefaultPrecision.Text) = False Then
                MsgBox("Default precision must a non-negative numeric data.", MsgBoxStyle.Information, _
                "Information")
                Me.txtDefaultPrecision.Text = Me.buffer
                Me.txtDefaultPrecision.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPayableAcc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPayableAccountCode.GotFocus
        TxtPayableAccountCode.BackColor = BackColorAll
    End Sub
    Private Sub TxtPayableAcc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPayableAccountCode.LostFocus
        TxtPayableAccountCode.BackColor = Me.TxtPrimaryCurrency.BackColor
    End Sub

    Private Sub TxtReceivableAcc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReceivableAcc.GotFocus
        TxtReceivableAcc.BackColor = BackColorAll
    End Sub
    Private Sub TxtReceivableAcc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReceivableAcc.LostFocus
        TxtReceivableAcc.BackColor = Me.TxtPrimaryCurrency.BackColor
    End Sub

    Private Sub TxtUnRealizedAcc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnRealizedExchangeGainLossAccount.GotFocus
        TxtUnRealizedExchangeGainLossAccount.BackColor = BackColorAll
    End Sub
    Private Sub TxtUnRealizedAcc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnRealizedExchangeGainLossAccount.LostFocus
        TxtUnRealizedExchangeGainLossAccount.BackColor = Me.TxtPrimaryCurrency.BackColor
    End Sub

  

    Private Sub TxtCurrencyCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrencyCode.GotFocus
        TxtCurrencyCode.BackColor = BackColorAll
    End Sub

    Private Sub TxtCurrencyCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrencyCode.LostFocus
        TxtCurrencyCode.BackColor = Me.TxtPrimaryCurrency.BackColor
    End Sub

    Private Sub TxtCurrencyName_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrencyName.GotFocus
        TxtCurrencyName.BackColor = BackColorAll
    End Sub

    Private Sub TxtCurrencyName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrencyName.LostFocus
        TxtCurrencyName.BackColor = Me.TxtPrimaryCurrency.BackColor
    End Sub

    Private Sub txtDefaultPrecision_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDefaultPrecision.GotFocus
        txtDefaultPrecision.BackColor = BackColorAll
    End Sub

    Private Sub txtDefaultPrecision_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDefaultPrecision.LostFocus
        txtDefaultPrecision.BackColor = Me.TxtPrimaryCurrency.BackColor
    End Sub

    Private Sub txtPayableAccountCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPayableAccountCode.KeyDown
        Dim i As New FrmMasterAkun_Cari1

        Try
            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtPayableAccountCode.Text) = True Then
                    Me.TxtPrimaryPayableAccount.Text = Me.ObjAkun.GetPrimary(Me.TxtPayableAccountCode.Text)
                    Me.TxtPayableAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryPayableAccount.Text)
                    Me.TxtReceivableAcc.Focus()
                Else
                    If Me.TxtPayableAccountCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtPayableAccountCode.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryPayableAccount.Text = i.mPrimaryAkun
                        Me.TxtPayableAccountCode.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryPayableAccount.Text)
                        Me.TxtPayableAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryPayableAccount.Text)
                        Me.TxtReceivableAcc.Focus()
                    Else
                        Me.TxtPayableAccountCode.Clear()
                        Me.TxtPayableAccountDesc.Clear()
                        Me.TxtPrimaryPayableAccount.Text = "-1"
                        Me.TxtPayableAccountCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtPayableAccountCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPayableAccountCode.Leave
        Dim i As New FrmMasterAkun_Cari1

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.TxtPayableAccountCode.Text <> "" Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtPayableAccountCode.Text) = True Then
                    Me.TxtPrimaryPayableAccount.Text = Me.ObjAkun.GetPrimary(Me.TxtPayableAccountCode.Text)
                    Me.TxtPayableAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryPayableAccount.Text)
                Else
                    If Me.TxtPayableAccountCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtPayableAccountCode.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryPayableAccount.Text = i.mPrimaryAkun
                        Me.TxtPayableAccountCode.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryPayableAccount.Text)
                        Me.TxtPayableAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryPayableAccount.Text)
                        Me.TxtReceivableAcc.Focus()
                    Else
                        Me.TxtPayableAccountCode.Clear()
                        Me.TxtPayableAccountDesc.Clear()
                        Me.TxtPrimaryPayableAccount.Text = "-1"
                        Me.TxtPayableAccountCode.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryPayableAccount.Text = "-1"
            End If
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
            If Me.TxtPayableAccountCode.Text <> "" Then
                i.txtCari.Text = Me.TxtPayableAccountCode.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.TxtPrimaryPayableAccount.Text = i.mPrimaryAkun
                Me.TxtPayableAccountCode.Text = obj.GetAccountCode(Me.TxtPrimaryPayableAccount.Text)
                Me.TxtPayableAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryPayableAccount.Text)
                Me.TxtReceivableAcc.Focus()
            Else
                Me.TxtPrimaryPayableAccount.Text = -1
                Me.TxtPayableAccountCode.Clear()
                Me.TxtPayableAccountDesc.Clear()
                Me.TxtPayableAccountCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub TxtReceivableAcc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtReceivableAcc.KeyDown
        Dim i As New FrmMasterAkun_Cari1

        Try
            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtReceivableAcc.Text) = True Then
                    Me.TxtPrimaryReceivableAcc.Text = Me.ObjAkun.GetPrimary(Me.TxtReceivableAcc.Text)
                    Me.TxtReceivableAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryReceivableAcc.Text)
                    Me.TxtUnRealizedExchangeGainLossAccount.Focus()
                Else
                    If Me.TxtReceivableAcc.Text <> "" Then
                        i.txtCari.Text = Me.TxtReceivableAcc.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryReceivableAcc.Text = i.mPrimaryAkun
                        Me.TxtReceivableAcc.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryReceivableAcc.Text)
                        Me.TxtReceivableAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryReceivableAcc.Text)
                        Me.TxtUnRealizedExchangeGainLossAccount.Focus()
                    Else
                        Me.TxtReceivableAcc.Clear()
                        Me.TxtReceivableAccDesc.Clear()
                        Me.TxtPrimaryReceivableAcc.Text = "-1"
                        Me.TxtUnRealizedExchangeGainLossAccount.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtReceivableAcc_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReceivableAcc.Leave
        Dim i As New FrmMasterAkun_Cari1

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.TxtReceivableAcc.Text <> "" Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtReceivableAcc.Text) = True Then
                    Me.TxtPrimaryReceivableAcc.Text = Me.ObjAkun.GetPrimary(Me.TxtReceivableAcc.Text)
                    Me.TxtReceivableAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryReceivableAcc.Text)
                Else
                    If Me.TxtReceivableAcc.Text <> "" Then
                        i.txtCari.Text = Me.TxtReceivableAcc.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryReceivableAcc.Text = i.mPrimaryAkun
                        Me.TxtReceivableAcc.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryReceivableAcc.Text)
                        Me.TxtReceivableAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryReceivableAcc.Text)
                        Me.TxtUnRealizedExchangeGainLossAccount.Focus()
                    Else
                        Me.TxtReceivableAcc.Clear()
                        Me.TxtReceivableAccDesc.Clear()
                        Me.TxtPrimaryReceivableAcc.Text = "-1"
                        Me.TxtUnRealizedExchangeGainLossAccount.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryReceivableAcc.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindReceivableAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindReceivableAccount.Click
        Dim i As New FrmMasterAkun_Cari1
        Dim obj As New ClsMasterAkun1

        Try
            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BackColorAll
            If Me.TxtReceivableAcc.Text <> "" Then
                i.txtCari.Text = Me.TxtReceivableAcc.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.TxtPrimaryReceivableAcc.Text = i.mPrimaryAkun
                Me.TxtReceivableAcc.Text = obj.GetAccountCode(Me.TxtPrimaryReceivableAcc.Text)
                Me.TxtReceivableAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryReceivableAcc.Text)
                Me.TxtUnRealizedExchangeGainLossAccount.Focus()
            Else
                Me.TxtPrimaryReceivableAcc.Text = -1
                Me.TxtReceivableAcc.Clear()
                Me.TxtReceivableAccDesc.Clear()
                Me.TxtUnRealizedExchangeGainLossAccount.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub TxtUnRealizedExchangeGainLossAccount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUnRealizedExchangeGainLossAccount.KeyDown
        Dim i As New FrmMasterAkun_Cari1

        Try
            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtUnRealizedExchangeGainLossAccount.Text) = True Then
                    Me.TxtPrimaryUnrealizedExchangeAcc.Text = Me.ObjAkun.GetPrimary(Me.TxtUnRealizedExchangeGainLossAccount.Text)
                    Me.TxtUnrealizedExchangeGainLossAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryUnrealizedExchangeAcc.Text)
                    Me.TxtRealizedExchangeGainLossAccount.Focus()
                Else
                    If Me.TxtUnRealizedExchangeGainLossAccount.Text <> "" Then
                        i.txtCari.Text = Me.TxtUnRealizedExchangeGainLossAccount.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryUnrealizedExchangeAcc.Text = i.mPrimaryAkun
                        Me.TxtUnRealizedExchangeGainLossAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryUnrealizedExchangeAcc.Text)
                        Me.TxtUnrealizedExchangeGainLossAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryUnrealizedExchangeAcc.Text)
                        Me.TxtRealizedExchangeGainLossAccount.Focus()
                    Else
                        Me.TxtUnRealizedExchangeGainLossAccount.Clear()
                        Me.TxtUnrealizedExchangeGainLossAccDesc.Clear()
                        Me.TxtPrimaryUnrealizedExchangeAcc.Text = "-1"
                        Me.TxtRealizedExchangeGainLossAccount.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtUnRealizedExchangeGainLossAccount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnRealizedExchangeGainLossAccount.Leave
        Dim i As New FrmMasterAkun_Cari1

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.TxtUnRealizedExchangeGainLossAccount.Text <> "" Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtUnRealizedExchangeGainLossAccount.Text) = True Then
                    Me.TxtPrimaryUnrealizedExchangeAcc.Text = Me.ObjAkun.GetPrimary(Me.TxtUnRealizedExchangeGainLossAccount.Text)
                    Me.TxtUnrealizedExchangeGainLossAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryUnrealizedExchangeAcc.Text)
                Else
                    If Me.TxtUnRealizedExchangeGainLossAccount.Text <> "" Then
                        i.txtCari.Text = Me.TxtUnRealizedExchangeGainLossAccount.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryUnrealizedExchangeAcc.Text = i.mPrimaryAkun
                        Me.TxtUnRealizedExchangeGainLossAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryUnrealizedExchangeAcc.Text)
                        Me.TxtUnrealizedExchangeGainLossAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryUnrealizedExchangeAcc.Text)
                        Me.TxtRealizedExchangeGainLossAccount.Focus()
                    Else
                        Me.TxtUnRealizedExchangeGainLossAccount.Clear()
                        Me.TxtUnrealizedExchangeGainLossAccDesc.Clear()
                        Me.TxtPrimaryUnrealizedExchangeAcc.Text = "-1"
                        Me.TxtRealizedExchangeGainLossAccount.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryUnrealizedExchangeAcc.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindUnrealizeExchangeGainLossAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindUnrealizeExchangeGainLossAccount.Click
        Dim i As New FrmMasterAkun_Cari1
        Dim obj As New ClsMasterAkun1

        Try
            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BackColorAll
            If Me.TxtUnRealizedExchangeGainLossAccount.Text <> "" Then
                i.txtCari.Text = Me.TxtUnRealizedExchangeGainLossAccount.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.TxtPrimaryUnrealizedExchangeAcc.Text = i.mPrimaryAkun
                Me.TxtUnRealizedExchangeGainLossAccount.Text = obj.GetAccountCode(Me.TxtPrimaryUnrealizedExchangeAcc.Text)
                Me.TxtUnrealizedExchangeGainLossAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryUnrealizedExchangeAcc.Text)
                Me.TxtRealizedExchangeGainLossAccount.Focus()
            Else
                Me.TxtPrimaryUnrealizedExchangeAcc.Text = -1
                Me.TxtUnRealizedExchangeGainLossAccount.Clear()
                Me.TxtUnrealizedExchangeGainLossAccDesc.Clear()
                Me.TxtRealizedExchangeGainLossAccount.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub TxtRealizedExchangeGainLossAccount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtRealizedExchangeGainLossAccount.KeyDown
        Dim i As New FrmMasterAkun_Cari1

        Try
            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtRealizedExchangeGainLossAccount.Text) = True Then
                    Me.TxtPrimaryRealizedExchangeAcc.Text = Me.ObjAkun.GetPrimary(Me.TxtRealizedExchangeGainLossAccount.Text)
                    Me.TxtRealizedExchangeGainLossAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryRealizedExchangeAcc.Text)
                    Me.btnOK.Focus()
                Else
                    If Me.TxtRealizedExchangeGainLossAccount.Text <> "" Then
                        i.txtCari.Text = Me.TxtRealizedExchangeGainLossAccount.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryRealizedExchangeAcc.Text = i.mPrimaryAkun
                        Me.TxtRealizedExchangeGainLossAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryRealizedExchangeAcc.Text)
                        Me.TxtRealizedExchangeGainLossAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryRealizedExchangeAcc.Text)
                        Me.btnOK.Focus()
                    Else
                        Me.TxtRealizedExchangeGainLossAccount.Clear()
                        Me.TxtRealizedExchangeGainLossAccDesc.Clear()
                        Me.TxtPrimaryRealizedExchangeAcc.Text = "-1"
                        Me.btnOK.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtRealizedExchangeGainLossAccount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRealizedExchangeGainLossAccount.Leave
        Dim i As New FrmMasterAkun_Cari1

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.TxtRealizedExchangeGainLossAccount.Text <> "" Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtRealizedExchangeGainLossAccount.Text) = True Then
                    Me.TxtPrimaryRealizedExchangeAcc.Text = Me.ObjAkun.GetPrimary(Me.TxtRealizedExchangeGainLossAccount.Text)
                    Me.TxtRealizedExchangeGainLossAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryRealizedExchangeAcc.Text)
                Else
                    If Me.TxtRealizedExchangeGainLossAccount.Text <> "" Then
                        i.txtCari.Text = Me.TxtRealizedExchangeGainLossAccount.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryRealizedExchangeAcc.Text = i.mPrimaryAkun
                        Me.TxtRealizedExchangeGainLossAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryRealizedExchangeAcc.Text)
                        Me.TxtRealizedExchangeGainLossAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryRealizedExchangeAcc.Text)
                        Me.btnOK.Focus()
                    Else
                        Me.TxtRealizedExchangeGainLossAccount.Clear()
                        Me.TxtRealizedExchangeGainLossAccDesc.Clear()
                        Me.TxtPrimaryRealizedExchangeAcc.Text = "-1"
                        Me.btnOK.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryRealizedExchangeAcc.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtRealizedAcc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRealizedExchangeGainLossAccount.GotFocus
        Try
            If Me.TxtCurrencyCode.Enabled = True Then
                TxtRealizedExchangeGainLossAccount.BackColor = BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRealizedAcc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRealizedExchangeGainLossAccount.LostFocus
        Try
            If Me.TxtCurrencyCode.Enabled = True Then
                TxtRealizedExchangeGainLossAccount.BackColor = Me.TxtPrimaryCurrency.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindRealizeExchangeExchangeGainLossAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindRealizeExchangeExchangeGainLossAccount.Click
        Dim i As New FrmMasterAkun_Cari1
        Dim obj As New ClsMasterAkun1

        Try
            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BackColorAll
            If Me.TxtRealizedExchangeGainLossAccount.Text <> "" Then
                i.txtCari.Text = Me.TxtRealizedExchangeGainLossAccount.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.TxtPrimaryRealizedExchangeAcc.Text = i.mPrimaryAkun
                Me.TxtRealizedExchangeGainLossAccount.Text = obj.GetAccountCode(Me.TxtPrimaryRealizedExchangeAcc.Text)
                Me.TxtRealizedExchangeGainLossAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryRealizedExchangeAcc.Text)
                Me.btnOK.Focus()
            Else
                Me.TxtPrimaryRealizedExchangeAcc.Text = -1
                Me.TxtRealizedExchangeGainLossAccount.Clear()
                Me.TxtRealizedExchangeGainLossAccDesc.Clear()
                Me.btnOK.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub TxtReceivableAccountHeaderCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReceivableAccountHeaderCode.GotFocus
        Try
            If Me.TxtCurrencyCode.Enabled = True Then
                TxtReceivableAccountHeaderCode.BackColor = BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReceivableAccountHeaderCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReceivableAccountHeaderCode.LostFocus
        Try
            If Me.TxtCurrencyCode.Enabled = True Then
                TxtReceivableAccountHeaderCode.BackColor = Me.TxtPrimaryCurrency.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReceivableAccountHeaderCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtReceivableAccountHeaderCode.KeyDown
        Dim i As New FrmMasterAkun_Cari2

        Try
            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_Header(Me.TxtReceivableAccountHeaderCode.Text) = True Then
                    Me.TxtPrimaryReceivableAccountHeader.Text = Me.ObjAkun.GetPrimary(Me.TxtReceivableAccountHeaderCode.Text)
                    Me.TxtReceivableAccountHeaderDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryReceivableAccountHeader.Text)
                    Me.TxtReceivableAccountHeaderNew2.Focus()
                Else
                    If Me.TxtReceivableAccountHeaderCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtReceivableAccountHeaderCode.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryReceivableAccountHeader.Text = i.mPrimaryAkun
                        Me.TxtReceivableAccountHeaderCode.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryReceivableAccountHeader.Text)
                        Me.TxtReceivableAccountHeaderDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryReceivableAccountHeader.Text)
                        Me.TxtReceivableAccountHeaderNew2.Focus()
                    Else
                        Me.TxtReceivableAccountHeaderCode.Clear()
                        Me.TxtReceivableAccountHeaderDesc.Clear()
                        Me.TxtPrimaryReceivableAccountHeader.Text = "-1"
                        Me.TxtReceivableAccountHeaderNew2.Focus()
                    End If
                End If
                Me.SetAccountReceivablesNew()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtReceivableAccountHeaderCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReceivableAccountHeaderCode.Leave
        Dim i As New FrmMasterAkun_Cari2

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.TxtReceivableAccountHeaderCode.Text <> "" Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_Header(Me.TxtReceivableAccountHeaderCode.Text) = True Then
                    Me.TxtPrimaryReceivableAccountHeader.Text = Me.ObjAkun.GetPrimary(Me.TxtReceivableAccountHeaderCode.Text)
                    Me.TxtReceivableAccountHeaderDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryReceivableAccountHeader.Text)
                Else
                    If Me.TxtReceivableAccountHeaderCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtReceivableAccountHeaderCode.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryReceivableAccountHeader.Text = i.mPrimaryAkun
                        Me.TxtReceivableAccountHeaderCode.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryReceivableAccountHeader.Text)
                        Me.TxtReceivableAccountHeaderDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryReceivableAccountHeader.Text)
                        Me.TxtReceivableAccountHeaderNew2.Focus()
                    Else
                        Me.TxtReceivableAccountHeaderCode.Clear()
                        Me.TxtReceivableAccountHeaderDesc.Clear()
                        Me.TxtPrimaryReceivableAccountHeader.Text = "-1"
                        Me.TxtReceivableAccountHeaderNew2.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryReceivableAccountHeader.Text = "-1"
            End If
            Me.SetAccountReceivablesNew()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindReceivableAccountHeader_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindReceivableAccountHeader.Click
        Dim i As New FrmMasterAkun_Cari2
        Dim obj As New ClsMasterAkun1

        Try
            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BackColorAll
            If Me.TxtReceivableAccountHeaderCode.Text <> "" Then
                i.txtCari.Text = Me.TxtReceivableAccountHeaderCode.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.TxtPrimaryReceivableAccountHeader.Text = i.mPrimaryAkun
                Me.TxtReceivableAccountHeaderCode.Text = obj.GetAccountCode(Me.TxtPrimaryReceivableAccountHeader.Text)
                Me.TxtReceivableAccountHeaderDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryReceivableAccountHeader.Text)
                Me.TxtReceivableAccountHeaderNew2.Focus()
            Else
                Me.TxtPrimaryReceivableAccountHeader.Text = -1
                Me.TxtReceivableAccountHeaderCode.Clear()
                Me.TxtReceivableAccountHeaderDesc.Clear()
                Me.TxtReceivableAccountHeaderNew2.Focus()
            End If
            Me.SetAccountReceivablesNew()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub TxtReceivableAccountHeaderNew2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReceivableAccountHeaderNew2.GotFocus
        Try
            If Me.TxtCurrencyCode.Enabled = True Then
                TxtReceivableAccountHeaderNew2.BackColor = BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReceivableAccountHeaderNew2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReceivableAccountHeaderNew2.LostFocus
        Try
            If Me.TxtCurrencyCode.Enabled = True Then
                TxtReceivableAccountHeaderNew2.BackColor = Me.TxtPrimaryCurrency.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReceivableAccountHeaderNewDesc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReceivableAccountHeaderNewDesc.GotFocus
        Try
            If Me.TxtCurrencyCode.Enabled = True Then
                TxtReceivableAccountHeaderNewDesc.BackColor = BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReceivableAccountHeaderNewDesc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReceivableAccountHeaderNewDesc.LostFocus
        Try
            If Me.TxtCurrencyCode.Enabled = True Then
                TxtReceivableAccountHeaderNewDesc.BackColor = Me.TxtPrimaryCurrency.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetAccountReceivablesNew()
        Try
            If Me.TxtPrimaryReceivableAccountHeader.Text = "" Or Me.TxtPrimaryReceivableAccountHeader.Text = "-1" Then
                Me.TxtReceivableAccountHeaderNew1.Text = ""
                Me.TxtReceivableAccountHeaderNew2.Text = ""
                Me.TxtReceivableAccountHeaderNewDesc.Text = ""
            Else
                Me.TxtReceivableAccountHeaderNew1.Text = Me.TxtReceivableAccountHeaderCode.Text.Substring(0, 1)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPayableAccountHeaderCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPayableAccountHeaderCode.GotFocus
        Try
            If Me.TxtCurrencyCode.Enabled = True Then
                TxtPayableAccountHeaderCode.BackColor = BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPayableAccountHeaderCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPayableAccountHeaderCode.LostFocus
        Try
            If Me.TxtCurrencyCode.Enabled = True Then
                TxtPayableAccountHeaderCode.BackColor = Me.TxtPrimaryCurrency.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPayableAccountHeaderCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPayableAccountHeaderCode.KeyDown
        Dim i As New FrmMasterAkun_Cari2

        Try
            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_Header(Me.TxtPayableAccountHeaderCode.Text) = True Then
                    Me.TxtPrimaryPayableAccountHeader.Text = Me.ObjAkun.GetPrimary(Me.TxtPayableAccountHeaderCode.Text)
                    Me.TxtPayableAccountHeaderDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryPayableAccountHeader.Text)
                    Me.TxtPayableAccountHeaderNew2.Focus()
                Else
                    If Me.TxtPayableAccountHeaderCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtPayableAccountHeaderCode.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryPayableAccountHeader.Text = i.mPrimaryAkun
                        Me.TxtPayableAccountHeaderCode.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryPayableAccountHeader.Text)
                        Me.TxtPayableAccountHeaderDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryPayableAccountHeader.Text)
                        Me.TxtPayableAccountHeaderNew2.Focus()
                    Else
                        Me.TxtPayableAccountHeaderCode.Clear()
                        Me.TxtPayableAccountHeaderDesc.Clear()
                        Me.TxtPrimaryPayableAccountHeader.Text = "-1"
                        Me.TxtPayableAccountHeaderNew2.Focus()
                    End If
                End If
                Me.SetAccountPayableNew()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtPayableAccountHeaderCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPayableAccountHeaderCode.Leave
        Dim i As New FrmMasterAkun_Cari2

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.TxtPayableAccountHeaderCode.Text <> "" Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_Header(Me.TxtPayableAccountHeaderCode.Text) = True Then
                    Me.TxtPrimaryPayableAccountHeader.Text = Me.ObjAkun.GetPrimary(Me.TxtPayableAccountHeaderCode.Text)
                    Me.TxtPayableAccountHeaderDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryPayableAccountHeader.Text)
                Else
                    If Me.TxtPayableAccountHeaderCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtPayableAccountHeaderCode.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryPayableAccountHeader.Text = i.mPrimaryAkun
                        Me.TxtPayableAccountHeaderCode.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryPayableAccountHeader.Text)
                        Me.TxtPayableAccountHeaderDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryPayableAccountHeader.Text)
                        Me.TxtPayableAccountHeaderNew2.Focus()
                    Else
                        Me.TxtPayableAccountHeaderCode.Clear()
                        Me.TxtPayableAccountHeaderDesc.Clear()
                        Me.TxtPrimaryPayableAccountHeader.Text = "-1"
                        Me.TxtPayableAccountHeaderNew2.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryPayableAccountHeader.Text = "-1"
            End If
            Me.SetAccountPayableNew()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindPayableAccountHeader_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindPayableAccountHeader.Click
        Dim i As New FrmMasterAkun_Cari2
        Dim obj As New ClsMasterAkun1

        Try
            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BackColorAll
            If Me.TxtPayableAccountHeaderCode.Text <> "" Then
                i.txtCari.Text = Me.TxtPayableAccountHeaderCode.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.TxtPrimaryPayableAccountHeader.Text = i.mPrimaryAkun
                Me.TxtPayableAccountHeaderCode.Text = obj.GetAccountCode(Me.TxtPrimaryPayableAccountHeader.Text)
                Me.TxtPayableAccountHeaderDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryPayableAccountHeader.Text)
                Me.TxtPayableAccountHeaderNew2.Focus()
            Else
                Me.TxtPrimaryPayableAccountHeader.Text = -1
                Me.TxtPayableAccountHeaderCode.Clear()
                Me.TxtPayableAccountHeaderDesc.Clear()
                Me.TxtPayableAccountHeaderNew2.Focus()
            End If
            Me.SetAccountPayableNew()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub TxtPayableAccountHeaderNew2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPayableAccountHeaderNew2.GotFocus
        Try
            If Me.TxtCurrencyCode.Enabled = True Then
                TxtPayableAccountHeaderNew2.BackColor = BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPayableAccountHeaderNew2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPayableAccountHeaderNew2.LostFocus
        Try
            If Me.TxtCurrencyCode.Enabled = True Then
                TxtPayableAccountHeaderNew2.BackColor = Me.TxtPrimaryCurrency.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPayableAccountHeaderNewDesc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPayableAccountHeaderNewDesc.GotFocus
        Try
            If Me.TxtCurrencyCode.Enabled = True Then
                TxtPayableAccountHeaderNewDesc.BackColor = BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPayableAccountHeaderNewDesc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPayableAccountHeaderNewDesc.LostFocus
        Try
            If Me.TxtCurrencyCode.Enabled = True Then
                TxtPayableAccountHeaderNewDesc.BackColor = Me.TxtPrimaryCurrency.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetAccountPayableNew()
        Try
            If Me.TxtPrimaryPayableAccountHeader.Text = "" Or Me.TxtPrimaryPayableAccountHeader.Text = "-1" Then
                Me.TxtPayableAccountHeaderNew1.Text = ""
                Me.TxtPayableAccountHeaderNew2.Text = ""
                Me.TxtPayableAccountHeaderNewDesc.Text = ""
            Else
                Me.TxtPayableAccountHeaderNew1.Text = Me.TxtPayableAccountHeaderCode.Text.Substring(0, 1)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnrealizedGainAccCode1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUnrealizedGainAccCode1.KeyDown
        Dim i As New FrmMasterAkun_Cari1

        Try
            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtUnrealizedGainAccCode1.Text) = True Then
                    Me.TxtPrimaryUnrealiazedGainAccount1.Text = Me.ObjAkun.GetPrimary(Me.TxtUnrealizedGainAccCode1.Text)
                    Me.TxtUnrealizedGainAccDesc1.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryUnrealiazedGainAccount1.Text)
                    Me.TxtRealiazedGainAccountCode1.Focus()
                Else
                    If Me.TxtUnrealizedGainAccCode1.Text <> "" Then
                        i.txtCari.Text = Me.TxtUnrealizedGainAccCode1.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryUnrealiazedGainAccount1.Text = i.mPrimaryAkun
                        Me.TxtUnrealizedGainAccCode1.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryUnrealiazedGainAccount1.Text)
                        Me.TxtUnrealizedGainAccDesc1.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryUnrealiazedGainAccount1.Text)
                        Me.TxtRealiazedGainAccountCode1.Focus()
                    Else
                        Me.TxtUnrealizedGainAccCode1.Clear()
                        Me.TxtUnrealizedGainAccDesc1.Clear()
                        Me.TxtPrimaryUnrealiazedGainAccount1.Text = "-1"
                        Me.TxtRealiazedGainAccountCode1.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtUnrealizedGainAccCode1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnrealizedGainAccCode1.Leave
        Dim i As New FrmMasterAkun_Cari1

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.TxtUnrealizedGainAccCode1.Text <> "" Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtUnrealizedGainAccCode1.Text) = True Then
                    Me.TxtPrimaryUnrealiazedGainAccount1.Text = Me.ObjAkun.GetPrimary(Me.TxtUnrealizedGainAccCode1.Text)
                    Me.TxtUnrealizedGainAccDesc1.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryUnrealiazedGainAccount1.Text)
                Else
                    If Me.TxtUnrealizedGainAccCode1.Text <> "" Then
                        i.txtCari.Text = Me.TxtUnrealizedGainAccCode1.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryUnrealiazedGainAccount1.Text = i.mPrimaryAkun
                        Me.TxtUnrealizedGainAccCode1.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryUnrealiazedGainAccount1.Text)
                        Me.TxtUnrealizedGainAccDesc1.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryUnrealiazedGainAccount1.Text)
                        Me.TxtRealiazedGainAccountCode1.Focus()
                    Else
                        Me.TxtUnrealizedGainAccCode1.Clear()
                        Me.TxtUnrealizedGainAccDesc1.Clear()
                        Me.TxtPrimaryUnrealiazedGainAccount1.Text = "-1"
                        Me.TxtRealiazedGainAccountCode1.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryUnrealiazedGainAccount1.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub BtnFindUnrealizedGainAcc1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindUnrealizedGainAcc1.Click
        Dim i As New FrmMasterAkun_Cari1
        Dim obj As New ClsMasterAkun1

        Try
            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BackColorAll
            If Me.TxtUnrealizedGainAccCode1.Text <> "" Then
                i.txtCari.Text = Me.TxtUnrealizedGainAccCode1.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.TxtPrimaryUnrealiazedGainAccount1.Text = i.mPrimaryAkun
                Me.TxtUnrealizedGainAccCode1.Text = obj.GetAccountCode(Me.TxtPrimaryUnrealiazedGainAccount1.Text)
                Me.TxtUnrealizedGainAccDesc1.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryUnrealiazedGainAccount1.Text)
                Me.TxtRealiazedGainAccountCode1.Focus()
            Else
                Me.TxtPrimaryUnrealiazedGainAccount1.Text = -1
                Me.TxtUnrealizedGainAccCode1.Clear()
                Me.TxtUnrealizedGainAccDesc1.Clear()
                Me.TxtRealiazedGainAccountCode1.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub BtnFindRealiazedGainAccount1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindRealiazedGainAccount1.Click
        Dim i As New FrmMasterAkun_Cari1
        Dim obj As New ClsMasterAkun1

        Try
            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BackColorAll
            If Me.TxtRealiazedGainAccountCode1.Text <> "" Then
                i.txtCari.Text = Me.TxtRealiazedGainAccountCode1.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.TxtPrimaryRealiazedGainAccount1.Text = i.mPrimaryAkun
                Me.TxtRealiazedGainAccountCode1.Text = obj.GetAccountCode(Me.TxtPrimaryRealiazedGainAccount1.Text)
                Me.TxtRealiazedGainAccountDesc1.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryRealiazedGainAccount1.Text)
                Me.btnOK.Focus()
            Else
                Me.TxtPrimaryRealiazedGainAccount1.Text = -1
                Me.TxtRealiazedGainAccountCode1.Clear()
                Me.TxtRealiazedGainAccountDesc1.Clear()
                Me.btnOK.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub GotFocus1(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRealiazedGainAccountCode1.GotFocus, TxtRealiazedGainAccountDesc1.GotFocus, TxtUnrealizedGainAccCode1.GotFocus, TxtUnrealizedGainAccDesc1.GotFocus
        Try
            If Me.TxtCurrencyCode.Enabled = True Then
                Dim txt As TextBox = TryCast(sender, TextBox)
                If txt.ReadOnly = False Then
                    txt.BackColor = Me.BackColorAll
                Else
                    txt.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRealiazedGainAccountCode1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtRealiazedGainAccountCode1.KeyDown
        Dim i As New FrmMasterAkun_Cari1

        Try
            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtRealiazedGainAccountCode1.Text) = True Then
                    Me.TxtPrimaryRealiazedGainAccount1.Text = Me.ObjAkun.GetPrimary(Me.TxtRealiazedGainAccountCode1.Text)
                    Me.TxtRealiazedGainAccountDesc1.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryRealiazedGainAccount1.Text)
                    Me.btnOK.Focus()
                Else
                    If Me.TxtRealiazedGainAccountCode1.Text <> "" Then
                        i.txtCari.Text = Me.TxtRealiazedGainAccountCode1.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryRealiazedGainAccount1.Text = i.mPrimaryAkun
                        Me.TxtRealiazedGainAccountCode1.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryRealiazedGainAccount1.Text)
                        Me.TxtRealiazedGainAccountDesc1.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryRealiazedGainAccount1.Text)
                        Me.btnOK.Focus()
                    Else
                        Me.TxtRealiazedGainAccountCode1.Clear()
                        Me.TxtRealiazedGainAccountDesc1.Clear()
                        Me.TxtPrimaryRealiazedGainAccount1.Text = "-1"
                        Me.btnOK.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtRealiazedGainAccountCode1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRealiazedGainAccountCode1.Leave
        Dim i As New FrmMasterAkun_Cari1

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.TxtRealiazedGainAccountCode1.Text <> "" Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtRealiazedGainAccountCode1.Text) = True Then
                    Me.TxtPrimaryRealiazedGainAccount1.Text = Me.ObjAkun.GetPrimary(Me.TxtRealiazedGainAccountCode1.Text)
                    Me.TxtRealiazedGainAccountDesc1.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryRealiazedGainAccount1.Text)
                Else
                    If Me.TxtRealiazedGainAccountCode1.Text <> "" Then
                        i.txtCari.Text = Me.TxtRealiazedGainAccountCode1.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryRealiazedGainAccount1.Text = i.mPrimaryAkun
                        Me.TxtRealiazedGainAccountCode1.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryRealiazedGainAccount1.Text)
                        Me.TxtRealiazedGainAccountDesc1.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryRealiazedGainAccount1.Text)
                        Me.btnOK.Focus()
                    Else
                        Me.TxtRealiazedGainAccountCode1.Clear()
                        Me.TxtRealiazedGainAccountDesc1.Clear()
                        Me.TxtPrimaryRealiazedGainAccount1.Text = "-1"
                        Me.btnOK.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryRealiazedGainAccount1.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub LostFocus1(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRealiazedGainAccountCode1.LostFocus, TxtRealiazedGainAccountDesc1.LostFocus, TxtUnrealizedGainAccCode1.LostFocus, TxtUnrealizedGainAccDesc1.LostFocus
        Try
            If Me.TxtCurrencyCode.Enabled = True Then
                Dim txt As TextBox = TryCast(sender, TextBox)
                If txt.ReadOnly = False Then
                    txt.BackColor = Me.TxtPrimaryRealiazedGainAccount1.BackColor
                Else
                    txt.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDefaultPrecision_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDefaultPrecision.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.txtDefaultPrecision, sender, e, 6)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDefaultPrecision_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDefaultPrecision.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.txtDefaultPrecision, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class