Public Class FrmTransaksiSalesInvoice_KeteranganTambahanEFaktur
    Public vIndexAcuan As Integer = -1

    Private Sub FrmTransaksiSalesInvoice_KeteranganTambahanEFaktur_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Icon = My.Resources.desktop_icon_3_fw
            Me.ShowIcon = True
            'Me.ShowInTaskbar = False
            If Me.TxtTaxInvoiceSeries.Text = "070" Or _
            Me.TxtTaxInvoiceSeries.Text = "071" Then
                Me.cboKeteranganTambahan.Items.Add("untuk Kawasan Bebas")
                Me.cboKeteranganTambahan.Items.Add("untuk Tempat Penimbunan Berikat")
                Me.cboKeteranganTambahan.Items.Add("untuk Hibah dan Bantuan Luar Negeri")
                Me.cboKeteranganTambahan.Items.Add("untuk Avtur")
                Me.cboKeteranganTambahan.Items.Add("untuk Lainnya")
                Me.cboKeteranganTambahan.Items.Add("untuk Kontraktor Perjanjian Karya Pengusahaan Pertambangan Batubara Generasi I")
            ElseIf Me.TxtTaxInvoiceSeries.Text = "080" Or _
            Me.TxtTaxInvoiceSeries.Text = "081" Then
                Me.cboKeteranganTambahan.Items.Add("untuk BKP dan JKP Tertentu")
                Me.cboKeteranganTambahan.Items.Add("untuk BKP Tertentu yang Bersifat Strategis")
                Me.cboKeteranganTambahan.Items.Add("untuk Jasa Kebandarudaraan")
                Me.cboKeteranganTambahan.Items.Add("untuk Lainnya")
            End If
            If Me.vIndexAcuan <> -1 Then
                If Me.vIndexAcuan <= Me.cboKeteranganTambahan.Items.Count - 1 Then
                    Me.cboKeteranganTambahan.SelectedIndex = Me.vIndexAcuan
                Else
                    Me.cboKeteranganTambahan.SelectedIndex = 0
                End If
            Else
                Me.cboKeteranganTambahan.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub
End Class