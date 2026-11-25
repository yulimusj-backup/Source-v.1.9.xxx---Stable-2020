Option Explicit On
Option Strict Off

Imports Core_Module
Public Class FrmTransaksiSalesInvoice_InfoFakturPajak
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public PrimarySalesInvoice As Double = -1
    Public VATType As ClsTransaksiSalesInvoice.EnumTaxType = ClsTransaksiSalesInvoice.EnumTaxType.Tax07
    Public InvoiceDate As Date = #1/1/2025#
    Public InvoiceNo As String = ""
    Public Simpan As Boolean = False
    Public PrimaryAddInfo As Double = -1
    Public PrimaryFacilityStamp As Double = -1

    Private Sub FrmTransaksiSalesInvoice_InfoFakturPajak_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.LoadDataCbo()
            Me.dtpSPPBDate.Value = InvoiceDate
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDataCbo()
        Dim dtaAddInfo As New Core_Module.DseMasterTaxAddInfoReferenceTableAdapters.tblCboTableAdapter
        Dim tblAddInfo As New Core_Module.DseMasterTaxAddInfoReference.tblCboDataTable
        Dim dtaFacilityStamp As New Core_Module.DseMasterTaxFacilityStampReferenceTableAdapters.TblCBOTableAdapter
        Dim tblFacilityStamp As New Core_Module.DseMasterTaxFacilityStampReference.TblCBODataTable
        Dim TypeAddInfo As Integer = 0
        Try
            If VATType = ClsTransaksiSalesInvoice.EnumTaxType.Tax07 Then
                TypeAddInfo = 0
            ElseIf VATType = ClsTransaksiSalesInvoice.EnumTaxType.Tax08 Then
                TypeAddInfo = 1
            Else
                Me.Simpan = False
                Me.Close()
            End If
            dtaAddInfo.FillBy_Type(tblAddInfo, TypeAddInfo)
            dtaFacilityStamp.FillBy_Type(tblFacilityStamp, TypeAddInfo)
            If tblAddInfo.Rows.Count > 0 Then
                Me.cboTaxAddInfo.DataSource = Nothing
                Me.cboTaxAddInfo.Items.Clear()
                Me.cboTaxAddInfo.DataSource = tblAddInfo.DefaultView
                Me.cboTaxAddInfo.ValueMember = "Primary"
                Me.cboTaxAddInfo.DisplayMember = "Viewer"
            End If
            If tblFacilityStamp.Rows.Count > 0 Then
                Me.cboTaxFacilityStamp.DataSource = Nothing
                Me.cboTaxFacilityStamp.Items.Clear()
                Me.cboTaxFacilityStamp.DataSource = tblFacilityStamp.DefaultView
                Me.cboTaxFacilityStamp.ValueMember = "Primary"
                Me.cboTaxFacilityStamp.DisplayMember = "Viewer"
            End If

            'set default ketika 08
            If VATType = ClsTransaksiSalesInvoice.EnumTaxType.Tax08 Then
                If Me.PrimaryAddInfo = -1 Then
                    Me.cboTaxAddInfo.SelectedIndex = 9 'item ke 10
                Else
                    Me.cboTaxAddInfo.SelectedValue = Me.PrimaryAddInfo
                End If
                If Me.PrimaryFacilityStamp = -1 Then
                    Me.cboTaxFacilityStamp.SelectedIndex = 9 'item ke 10
                Else
                    Me.cboTaxFacilityStamp.SelectedValue = Me.PrimaryFacilityStamp
                End If
                Me.TxtSPPBNo.Text = Me.InvoiceNo
                Me.dtpSPPBDate.Value = Me.InvoiceDate
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If apakahDataBenar() Then
                Me.Simpan = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            'konfirmasi batal atau tidak
            hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
            If hasil = MsgBoxResult.Yes Then
                Me.Simpan = False
                Me.Close()
            Else
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function apakahDataBenar() As Boolean
        Try
            'cek sppb no gk boleh kosong
            If String.IsNullOrEmpty(Me.TxtSPPBNo.Text) Then
                MsgBox("SPPB no cannot empty.", MsgBoxStyle.Information, "Information")
                Me.TxtSPPBNo.Focus()
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub TxtSPPBNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSPPBNo.GotFocus
        Try
            If Me.TxtSPPBNo.ReadOnly = True Then
                Me.TxtSPPBNo.BackColor = Me.TxtReadOnly.BackColor
            Else
                Me.TxtSPPBNo.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSPPBNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSPPBNo.LostFocus
        Try
            If Me.TxtSPPBNo.ReadOnly = True Then
                Me.TxtSPPBNo.BackColor = Me.TxtReadOnly.BackColor
            Else
                Me.TxtSPPBNo.BackColor = Me.TextBox1.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class