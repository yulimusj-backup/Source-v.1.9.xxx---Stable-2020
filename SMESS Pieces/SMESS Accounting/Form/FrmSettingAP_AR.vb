Option Explicit On
Option Strict Off

Imports AccModule
Public Class FrmSettingAP_AR
    Public BackColorAll As Color = Color.LightSeaGreen
    Public UserCode As String = ""

    Dim ObjInv As New ClsMasterAkun
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjCD As New Core_Module.ClsCekData

    Dim SettingInvoiceBerubah As Boolean = False
    Dim SettingAOPBerubah As Boolean = False
    Dim SettingDOBerubah As Boolean = False

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmTransaksiSettingJurnal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.TxtAkunAR.Focus()
            Me.SalinKeFieldSettingJurnal()
            Me.SalinKeFieldSettingInvoice()
            Me.SalinKeFieldSettingAOP()
            Me.SalinKeFieldSettingDO()
            Me.Text = "Global Settings"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim fileExists As Boolean

        Try
            'jika menggunakan custom invoice print maka alamatnya tidak boleh kosong dan filenya harus ada
            If Me.chkUseCustomInvoice.Checked = True Then
                'pastikan ada isinya
                Me.TxtCustomInvoiceAddress.Text.Trim()
                If Me.TxtCustomInvoiceAddress.Text = "" Then
                    MsgBox("Please specify valid custom invoice file.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
                'pastikan file tsb valid
                fileExists = My.Computer.FileSystem.FileExists(Me.TxtCustomInvoiceAddress.Text)
                If fileExists = False Then
                    MsgBox("Please specify valid custom invoice file.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If Me.chkUseCustomTaxInvoice.Checked = True Then
                'pastikan ada isinya
                Me.TxtCustomTaxInvoiceAddress.Text.Trim()
                If Me.TxtCustomTaxInvoiceAddress.Text = "" Then
                    MsgBox("Please specify valid custom tax invoice file.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
                'pastikan file tsb valid
                fileExists = My.Computer.FileSystem.FileExists(Me.TxtCustomTaxInvoiceAddress.Text)
                If fileExists = False Then
                    MsgBox("Please specify valid custom tax invoice file.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If Me.chkUseCustomDO.Checked = True Then
                'pastikan ada isinya
                Me.TxtCustomeDOAddress.Text.Trim()
                If Me.TxtCustomeDOAddress.Text = "" Then
                    MsgBox("Please specify valid custom delivery order file.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
                'pastikan file tsb valid
                fileExists = My.Computer.FileSystem.FileExists(Me.TxtCustomeDOAddress.Text)
                If fileExists = False Then
                    MsgBox("Please specify valid custom delivery order file.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            'simpan perubahan ke dalam basis data
            If Me.ApakahDataBenar() = True Then
                'konfirmasi disimpan atau tidak
                hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    'simpan data setting invoice
                    Me.SimpanDataSettingInvoice()
                    'simpan data setting aop
                    Me.SimpanDataSettingAOP()
                    'simpan data jurnal
                    Me.SimpanDataSettingJurnal()
                    'simpan data do
                    Me.SimpanDataSettingDO()

                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Account AR
    Private Sub TxtAkunAR_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles TxtAkunAR.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtAkunAR.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjInv.ApakahSudahAda(Me.TxtAkunAR.Text) = True Then
                        Me.TxtPrimaryAkunAR.Text = CStr(ObjInv.GetPrimary(Me.TxtAkunAR.Text))
                        Me.TxtAkunAR.Text = ObjInv.GetAccountCode(CDbl(Me.TxtPrimaryAkunAR.Text))
                        Me.TxtAkunAP.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.btnFindAR.PerformClick()
                    End If
                Else
                    Me.btnFindAR.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunAR_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtAkunAR.Leave
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtAkunAR.Text <> "" Then
                If ObjInv.ApakahSudahAda(Me.TxtAkunAR.Text) = True Then
                    Me.TxtPrimaryAkunAR.Text = CStr(ObjInv.GetPrimary(Me.TxtAkunAR.Text))
                    Me.TxtAkunAR.Text = ObjInv.GetAccountCode(CDbl(Me.TxtPrimaryAkunAR.Text))
                Else
                    If Me.TxtAkunAR.Text <> "" Then
                        ij.txtCari.Text = Me.TxtAkunAR.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryAkunAR.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtAkunAR.Text = ObjInv.GetAccountCode(CDbl(Me.TxtPrimaryAkunAR.Text))
                    Else
                        Me.TxtAkunAR.Clear()
                        Me.TxtPrimaryAkunAR.Text = "-1"
                        Me.TxtAkunAR.Focus()
                    End If
                End If
            Else
                Me.TxtAkunAR.Clear()
                Me.TxtPrimaryAkunAR.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnFindAR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindAR.Click
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtAkunAR.Text <> "" Then
                ij.txtCari.Text = Me.TxtAkunAR.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAkunAR.Text = CStr(ij.mPrimaryAkun)
                Me.TxtAkunAR.Text = ObjInv.GetAccountCode(CDbl(Me.TxtPrimaryAkunAR.Text))
                Me.TxtAkunAP.Focus()
            Else
                Me.TxtAkunAR.Clear()
                Me.TxtPrimaryAkunAR.Text = "-1"
                Me.TxtAkunAR.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtAkunAP_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
      Handles TxtAkunAP.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtAkunAP.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjInv.ApakahSudahAda(Me.TxtAkunAP.Text) = True Then
                        Me.TxtPrimaryAkunAP.Text = CStr(ObjInv.GetPrimary(Me.TxtAkunAP.Text))
                        Me.TxtAkunAP.Text = ObjInv.GetAccountCode(CDbl(Me.TxtPrimaryAkunAP.Text))
                        Me.TxtAkunAP.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.btnFindAkunAP.PerformClick()
                    End If
                Else
                    Me.btnFindAkunAP.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunAP_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtAkunAP.Leave
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtAkunAP.Text <> "" Then
                ''If Me.cboTipeTujuan.SelectedIndex = 0 Then
                If ObjInv.ApakahSudahAda(Me.TxtAkunAP.Text) = True Then
                    Me.TxtPrimaryAkunAP.Text = CStr(ObjInv.GetPrimary(Me.TxtAkunAP.Text))
                    Me.TxtAkunAP.Text = ObjInv.GetAccountCode(CDbl(Me.TxtPrimaryAkunAP.Text))
                Else
                    If Me.TxtAkunAP.Text <> "" Then
                        ij.txtCari.Text = Me.TxtAkunAP.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryAkunAP.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtAkunAP.Text = ObjInv.GetAccountCode(CDbl(Me.TxtPrimaryAkunAP.Text))
                    Else
                        Me.TxtAkunAP.Clear()
                        Me.TxtPrimaryAkunAP.Text = "-1"
                        Me.TxtAkunAP.Focus()
                    End If
                End If
                ''End If
            Else
                Me.TxtAkunAP.Clear()
                Me.TxtPrimaryAkunAP.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnFindAkunAP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindAkunAP.Click
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtAkunAP.Text <> "" Then
                ij.txtCari.Text = Me.TxtAkunAP.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAkunAP.Text = CStr(ij.mPrimaryAkun)
                Me.TxtAkunAP.Text = ObjInv.GetAccountCode(CDbl(Me.TxtPrimaryAkunAP.Text))
                Me.btnOK.Focus()
            Else
                Me.TxtAkunAP.Clear()
                Me.TxtPrimaryAkunAP.Text = "-1"
                Me.TxtAkunAP.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtAkunAR_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAkunAR.GotFocus
        Try
            Me.TxtAkunAR.BackColor = Core_Module.GetSMESSBackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunAR_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAkunAR.LostFocus
        Try
            Me.TxtAkunAR.BackColor = Color.White
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunAP_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAkunAP.GotFocus
        Try
            Me.TxtAkunAP.BackColor = Core_Module.GetSMESSBackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunAP_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAkunAP.LostFocus
        Try
            Me.TxtAkunAP.BackColor = Color.White
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeFieldSettingInvoice()
        Dim Dummy As String = ""

        Try
            'salin data use tax
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseTax)
            If Dummy = "" Then
                Me.chkTaxInvoice.Checked = False
            Else
                If Dummy = "0" Then
                    Me.chkTaxInvoice.Checked = False
                Else
                    Me.chkTaxInvoice.Checked = True
                    Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                    If Dummy <> "" Then
                        Me.TxtTaxInvoice.Text = Me.ObjCD.DigitGrouping(Dummy, 2)
                    Else
                        Me.TxtTaxInvoice.Text = Me.ObjCD.DigitGrouping(0, 2)
                    End If
                End If
            End If

            'salin data custom invoice print
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.PrintUseCustomInvoice_Bool)
            If Dummy = "" Then
                Me.chkUseCustomInvoice.Checked = True
                Me.chkUseCustomInvoice.Checked = False
            Else
                If Dummy = "0" Then
                    Me.chkUseCustomInvoice.Checked = True
                    Me.chkUseCustomInvoice.Checked = False
                Else
                    Me.chkUseCustomInvoice.Checked = True
                    Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomInvoiceAddress_Str)
                    Me.TxtCustomInvoiceAddress.Text = Dummy
                End If
            End If
            'salin data custom tax invoice print
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.PrintUseCustomTaxInvoice_Bool)
            If Dummy = "" Then
                Me.chkUseCustomTaxInvoice.Checked = True
                Me.chkUseCustomTaxInvoice.Checked = False
            Else
                If Dummy = "0" Then
                    Me.chkUseCustomTaxInvoice.Checked = True
                    Me.chkUseCustomTaxInvoice.Checked = False
                Else
                    Me.chkUseCustomTaxInvoice.Checked = True
                    Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomTaxInvoice_Str)
                    Me.TxtCustomTaxInvoiceAddress.Text = Dummy
                End If
            End If

            Me.TxtSignedByName.Text = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.SignedByName)
            Me.TxtSignedByPosition.Text = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.SignedByJabatan)

            Me.SettingInvoiceBerubah = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeFieldSettingDO()
        Dim Dummy As String = ""

        Try
            'salin data custom invoice print
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.PrintUseCustomDO_Bool)
            If Dummy = "" Then
                Me.chkUseCustomDO.Checked = True
                Me.chkUseCustomDO.Checked = False
            Else
                If Dummy = "0" Then
                    Me.chkUseCustomDO.Checked = True
                    Me.chkUseCustomDO.Checked = False
                Else
                    Me.chkUseCustomDO.Checked = True
                    Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomDO_Str)
                    Me.TxtCustomeDOAddress.Text = Dummy
                End If
            End If

            Me.SettingDOBerubah = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanDataSettingInvoice()
        Try
            If Me.SettingInvoiceBerubah = True Then
                If Me.chkTaxInvoice.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseTax, "1")
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage, Me.TxtTaxInvoice.Text)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseTax, "0")
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage, "0")
                End If

                If Me.chkUseCustomInvoice.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.PrintUseCustomInvoice_Bool, "1")
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomInvoiceAddress_Str, Me.TxtCustomInvoiceAddress.Text)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.PrintUseCustomInvoice_Bool, "0")
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomInvoiceAddress_Str, "")
                End If

                If Me.chkUseCustomTaxInvoice.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.PrintUseCustomTaxInvoice_Bool, "1")
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomTaxInvoice_Str, Me.TxtCustomTaxInvoiceAddress.Text)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.PrintUseCustomTaxInvoice_Bool, "0")
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomTaxInvoice_Str, "")
                End If

                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.SignedByJabatan, Me.TxtSignedByPosition.Text)
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.SignedByName, Me.TxtSignedByName.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanDataSettingDO()
        Try
            If Me.SettingDOBerubah = True Then
                If Me.chkUseCustomDO.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, Core_Module.ClsTransaksiSetting.EnumNamaParameter.PrintUseCustomDO_Bool, "1")
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomDO_Str, Me.TxtCustomeDOAddress.Text)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, Core_Module.ClsTransaksiSetting.EnumNamaParameter.PrintUseCustomDO_Bool, "0")
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomDO_Str, "")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkTaxInvoice_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTaxInvoice.CheckedChanged
        Try
            If Me.chkTaxInvoice.Checked = True Then
                Me.TxtTaxInvoice.ReadOnly = False
                Me.TxtTaxInvoice.TabStop = True
            Else
                Me.TxtTaxInvoice.ReadOnly = True
                Me.TxtTaxInvoice.TabStop = False
                Me.TxtTaxInvoice.Text = "0"
            End If
            Me.SettingInvoiceBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxInvoice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTaxInvoice.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtTaxInvoice, sender, e, 100)
        Me.SettingInvoiceBerubah = True
    End Sub

    Private Sub TxtTaxInvoice_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTaxInvoice.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtTaxInvoice, sender, e)
    End Sub

    Private Sub SalinKeFieldSettingJurnal()
        Dim Dummy As String = ""

        Try
            'salin data AP
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.PurchaseJurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_AP)
            If Dummy = "" Then
                Me.TxtPrimaryAkunAP.Text = "-1"
                Me.TxtAkunAP.Text = ""
            Else
                Me.TxtPrimaryAkunAP.Text = Dummy
                Me.TxtAkunAP.Text = Me.ObjInv.GetAccountCode(CDbl(Me.TxtPrimaryAkunAP.Text))
            End If

            'salin data AR
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SalesJurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_AR)
            If Dummy = "" Then
                Me.TxtPrimaryAkunAR.Text = "-1"
                Me.TxtAkunAR.Text = ""
            Else
                Me.TxtPrimaryAkunAR.Text = Dummy
                Me.TxtAkunAR.Text = Me.ObjInv.GetAccountCode(CDbl(Me.TxtPrimaryAkunAR.Text))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeFieldSettingAOP()
        Dim Dummy As String = ""

        Try
            'salin data use tax
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseTax)
            If Dummy = "" Then
                Me.ChkTaxAOP.Checked = False
            Else
                If Dummy = "0" Then
                    Me.ChkTaxAOP.Checked = False
                Else
                    Me.ChkTaxAOP.Checked = True
                    Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                    If Dummy <> "" Then
                        Me.TxtTaxAOP.Text = Me.ObjCD.DigitGrouping(Dummy, 2)
                    Else
                        Me.TxtTaxAOP.Text = Me.ObjCD.DigitGrouping("0", 2)
                    End If
                End If
            End If
            Me.SettingAOPBerubah = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanDataSettingAOP()
        Try
            If Me.SettingAOPBerubah = True Then
                If Me.ChkTaxAOP.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseTax, "1")
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage, Me.TxtTaxAOP.Text)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseTax, "0")
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage, "0")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanDataSettingJurnal()
        Try
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.PurchaseJurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_AP, Me.TxtPrimaryAkunAP.Text)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SalesJurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_AR, Me.TxtPrimaryAkunAR.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkTaxAOP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkTaxAOP.CheckedChanged
        Try
            If Me.ChkTaxAOP.Checked = True Then
                Me.TxtTaxAOP.ReadOnly = False
                Me.TxtTaxAOP.TabStop = True
            Else
                Me.TxtTaxAOP.ReadOnly = True
                Me.TxtTaxAOP.TabStop = False
                Me.TxtTaxAOP.Text = "0"
            End If
            Me.SettingAOPBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxAOP_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTaxAOP.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtTaxAOP, sender, e, 100)
        Me.SettingAOPBerubah = True
    End Sub

    Private Sub TxtTaxAOP_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTaxAOP.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtTaxAOP, sender, e)
    End Sub

    Private Sub chkUseCustomInvoice_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUseCustomInvoice.CheckedChanged
        Try
            If Me.chkUseCustomInvoice.Checked = True Then
                Me.TxtCustomInvoiceAddress.Enabled = True
                Me.btnFindCustomInvoiceReport.Enabled = True
            Else
                Me.TxtCustomInvoiceAddress.Enabled = False
                Me.btnFindCustomInvoiceReport.Enabled = False
                Me.TxtCustomInvoiceAddress.Clear()
            End If

            Me.SettingInvoiceBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCustomInvoiceAddress_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtCustomInvoiceAddress.TextChanged
        Try
            Me.SettingInvoiceBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindCustomInvoiceReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCustomInvoiceReport.Click
        Try
            Me.OpenFileDialog.Title = "Find Custom Invoice File"
            If Me.OpenFileDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.TxtCustomInvoiceAddress.Text = Me.OpenFileDialog.FileName
            Else
                Me.TxtCustomInvoiceAddress.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkUseCustomTaxInvoice_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUseCustomTaxInvoice.CheckedChanged
        Try
            If Me.chkUseCustomTaxInvoice.Checked = True Then
                Me.TxtCustomTaxInvoiceAddress.Enabled = True
                Me.btnFindCustomTaxInvoice.Enabled = True
            Else
                Me.TxtCustomTaxInvoiceAddress.Enabled = False
                Me.btnFindCustomTaxInvoice.Enabled = False
                Me.TxtCustomTaxInvoiceAddress.Clear()
            End If

            Me.SettingInvoiceBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCustomTaxInvoiceAddress_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtCustomTaxInvoiceAddress.TextChanged
        Try
            Me.SettingInvoiceBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindCustomTaxInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCustomTaxInvoice.Click
        Try
            Me.OpenFileDialog.Title = "Find Custom Tax Invoice File"
            If Me.OpenFileDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.TxtCustomTaxInvoiceAddress.Text = Me.OpenFileDialog.FileName
            Else
                Me.TxtCustomTaxInvoiceAddress.Text = ""
            End If
            Me.SettingInvoiceBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkUseCustomDO_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUseCustomDO.CheckedChanged
        Try
            If Me.chkUseCustomDO.Checked = True Then
                Me.TxtCustomeDOAddress.Enabled = True
                Me.btnFindCustomDOAddress.Enabled = True
            Else
                Me.TxtCustomeDOAddress.Enabled = False
                Me.btnFindCustomDOAddress.Enabled = False
                Me.TxtCustomeDOAddress.Clear()
            End If

            Me.SettingDOBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCustomeDOAddress_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtCustomeDOAddress.TextChanged
        Try
            Me.SettingDOBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindCustomDOAddress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCustomDOAddress.Click
        Try
            Me.OpenFileDialog.Title = "Find Custom Delivery Order File"
            If Me.OpenFileDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.TxtCustomeDOAddress.Text = Me.OpenFileDialog.FileName
            Else
                Me.TxtCustomeDOAddress.Text = ""
            End If
            Me.SettingDOBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSignedByName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtSignedByName.TextChanged
        Try
            Me.SettingInvoiceBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSignedByPosition_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtSignedByPosition.TextChanged
        Try
            Me.SettingInvoiceBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
