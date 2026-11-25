Option Explicit On
Option Strict Off

Public Class FrmTransaksiSalesInvoice_EditNo2
    Dim ObjInv As New ClsTransaksiSalesInvoice
    Dim ObjTax As New ClsTransaksiSalesInvoice_SettingTaxInvoice
    Dim jml As Double = 0
    Dim TBL As New DseTaxInvoiceSetting.tbltrans_salesinvoice_tax_invoice_setting1DataTable
    Dim ObjCD As New Core_Module.ClsCekData

    Public MinLength As Integer = 0
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public StartNumber, EndNumber As Double
    Public Result As String = ""
    Public InvDate As Date = Today
    Public PrimaryInvoice As Double = -1
    Public primaryTax As Double = -1

    Private Sub Txt_E_Faktur_Jalan_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_E_Faktur_Jalan.GotFocus, Txt_E_Faktur_Blok.GotFocus, _
    Txt_E_Faktur_Jalan.GotFocus, Txt_E_Faktur_Kelurahan.GotFocus, Txt_E_Faktur_KodePos.GotFocus, Txt_E_Faktur_No.GotFocus, Txt_E_Faktur_NoTelepon.GotFocus, Txt_E_Faktur_Propinsi.GotFocus, _
    Txt_E_Faktur_RT.GotFocus, Txt_E_Faktur_RW.GotFocus, Txt_E_Faktur_Kabupaten.GotFocus, Txt_E_Faktur_Kecamatan.GotFocus
        Dim Txt As TextBox

        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Txt.ReadOnly = True Then
                    Txt.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Txt.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_E_Faktur_Jalan_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_E_Faktur_Jalan.LostFocus, Txt_E_Faktur_Blok.LostFocus, _
    Txt_E_Faktur_Jalan.LostFocus, Txt_E_Faktur_Kelurahan.LostFocus, Txt_E_Faktur_KodePos.LostFocus, Txt_E_Faktur_No.LostFocus, Txt_E_Faktur_NoTelepon.LostFocus, Txt_E_Faktur_Propinsi.LostFocus, _
    Txt_E_Faktur_RT.LostFocus, Txt_E_Faktur_RW.LostFocus, Txt_E_Faktur_Kabupaten.LostFocus, Txt_E_Faktur_Kecamatan.LostFocus
        Dim Txt As TextBox

        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Txt.ReadOnly = True Then
                    Txt.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Txt.BackColor = Me.MaskedTextBox1.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Dim Dummy As String = ""
        Dim DummyPrimaryInvoice As Double = -1
        Dim row As DseTaxInvoiceSetting.tbltrans_salesinvoice_tax_invoice_setting1Row
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection

        Try
            'selected cbo harus ada isinya
            If Me.cboJenisFaktur.Text = "" Or Me.TxtPrimaryTaxInvoice.Text = "-1" Or Me.TxtPrimaryTaxInvoice.Text = "" Or Me.TBL.Rows.Count = 0 Then
                MsgBox("Please specify tax invoice", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'cek jumlah karakter
            If Me.mtb_Number.Text.Length <> Me.MinLength Then
                MsgBox("Must a number in " + Me.MinLength.ToString + " character.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'cek max dan min
            If CDbl(Me.mtb_Number.Text) < Me.StartNumber Then
                MsgBox("Minimum tax invoice number is: " + Me.StartNumber.ToString, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If CDbl(Me.mtb_Number.Text) > Me.EndNumber Then
                MsgBox("Maximum tax invoice number is: " + Me.EndNumber.ToString, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'cek belum pernah digunakan dalam invoice
            row = Me.TBL.DefaultView(Me.cboJenisFaktur.SelectedIndex).Row
            Dummy = row.prefix + Me.mtb_Number.Text
            DummyPrimaryInvoice = Me.ObjInv.GetPrimaryInvoiceByNoFakturPajak(Dummy)
            If DummyPrimaryInvoice <> -1 Then
                If PrimaryInvoice = -1 Then
                    Result = Dummy
                Else
                    If CDbl(PrimaryInvoice) = DummyPrimaryInvoice Then
                        Result = Dummy
                    Else
                        MsgBox("The tax invoice number " + Dummy + " is used by " + _
                        Me.ObjInv.GetNoInvoice(DummyPrimaryInvoice) + ".", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                End If
            End If
            Result = Dummy

            'pastikan data faktur pajak lengkap
            KoleksiNamaData.Add("nama jalan")
            KoleksiTextBox.Add(Me.Txt_E_Faktur_Jalan)
            KoleksiNamaData.Add("blok")
            KoleksiTextBox.Add(Me.Txt_E_Faktur_Blok)
            KoleksiNamaData.Add("nomor")
            KoleksiTextBox.Add(Me.Txt_E_Faktur_No)
            KoleksiNamaData.Add("RT")
            KoleksiTextBox.Add(Me.Txt_E_Faktur_RT)
            KoleksiNamaData.Add("RW")
            KoleksiTextBox.Add(Me.Txt_E_Faktur_RW)
            KoleksiNamaData.Add("Kelurahan")
            KoleksiTextBox.Add(Me.Txt_E_Faktur_Kelurahan)
            KoleksiNamaData.Add("Kecamatan")
            KoleksiTextBox.Add(Me.Txt_E_Faktur_Kecamatan)
            KoleksiNamaData.Add("Kabupaten")
            KoleksiTextBox.Add(Me.Txt_E_Faktur_Kabupaten)
            KoleksiNamaData.Add("Propinsi")
            KoleksiTextBox.Add(Me.Txt_E_Faktur_Propinsi)
            KoleksiNamaData.Add("Kode Pos")
            KoleksiTextBox.Add(Me.Txt_E_Faktur_KodePos)
            KoleksiNamaData.Add("Telepon")
            KoleksiTextBox.Add(Me.Txt_E_Faktur_NoTelepon)
            If Me.ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Exit Sub
            End If

            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub mtb_Number_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mtb_Number.GotFocus
        Try
            Me.mtb_Number.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub mtb_Number_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mtb_Number.LostFocus
        Try
            Me.mtb_Number.BackColor = Me.MaskedTextBox1.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSalesInvoice_EditNo2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim i As Integer = 0
        Dim row As DseTaxInvoiceSetting.tbltrans_salesinvoice_tax_invoice_setting1Row

        Try
            'isikan cbo berdasarkan invoice date
            Me.ObjTax.IsiCBOFakturPajak(Me.ObjCD.GetFisrtDateOfTheYear(Me.InvDate), Me.TBL, Me.cboJenisFaktur)
            If Me.TBL.Rows.Count > 0 Then
                For j As Integer = 0 To Me.TBL.Rows.Count - 1
                    row = TBL.DefaultView(j).Row
                    If row.primary = Me.primaryTax Then
                        Me.cboJenisFaktur.SelectedIndex = j
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboJenisFaktur_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboJenisFaktur.SelectedValueChanged
        Dim row As DseTaxInvoiceSetting.tbltrans_salesinvoice_tax_invoice_setting1Row

        Try
            If Me.cboJenisFaktur.SelectedIndex = -1 Then
                Me.TxtPrimaryTaxInvoice.Text = "-1"
                Me.StartNumber = 0
                Me.EndNumber = 0
            Else
                row = Me.TBL.DefaultView(Me.cboJenisFaktur.SelectedIndex).Row
                Me.TxtPrimaryTaxInvoice.Text = row.primary
                Me.StartNumber = row.starting_number
                Me.EndNumber = row.ending_number
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class