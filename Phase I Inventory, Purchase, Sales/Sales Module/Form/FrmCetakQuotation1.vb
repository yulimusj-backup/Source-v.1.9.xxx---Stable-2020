Public Class FrmCetakQuotation1
    Public Primary As Double = -1
    Public QuotationDate As Date = Today
    Public DefaultPrecision As Integer = 0
    Public JabatanProforma As String = ""

    Public FrmParent As FrmTransaksiQuotation

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Me.Close()
            Me.cboLanguage.SelectedIndex = 0
            Me.cboVersion.SelectedIndex = 0
            Me.NonAktifCRV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmCetakQuotation_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.cbocontrol()
            Me.NonAktifCRV()
            Me.BtnPreview.PerformClick()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cbocontrol()
        Try
            Me.cboLanguage.Items.Clear()
            Me.cboVersion.Items.Clear()

            Me.cboLanguage.Items.Add("English")
            Me.cboLanguage.Items.Add("Indonesian")
            Me.cboLanguage.SelectedIndex = 1

            Me.cboVersion.Items.Add("Short")
            Me.cboVersion.Items.Add("Long")
            Me.cboVersion.SelectedIndex = 0

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cetak()
        Dim obj As New CustomSMESS.ClsTransaksiQuotation

        Try
            If Me.cboLanguage.SelectedIndex = 0 And Me.cboVersion.SelectedIndex = 0 Then
                obj.CetakQuotationVersiEnglishPendek(crv1, Me.Primary)
                Me.crv1.Visible = True
                Me.TabControl1.Visible = False
            ElseIf Me.cboLanguage.SelectedIndex = 0 And Me.cboVersion.SelectedIndex = 1 Then
                obj.CetakQuotationVersiEnglishPanjang(crv2, Me.Primary)
                obj.CetakQuotationVersiDetail(CrvDetail, Me.Primary)
                Me.TabControl1.Visible = True
                Me.crv1.Visible = False
            ElseIf Me.cboLanguage.SelectedIndex = 1 And Me.cboVersion.SelectedIndex = 0 Then
                obj.CetakQuotationVersiIndoPendek(crv1, Me.Primary)
                Me.crv1.Visible = True
                Me.TabControl1.Visible = False
            ElseIf Me.cboLanguage.SelectedIndex = 1 And Me.cboVersion.SelectedIndex = 1 Then
                obj.CetakQuotationVersiIndoPanjang(crv2, Me.Primary)
                obj.CetakQuotationVersiDetail(CrvDetail, Me.Primary)
                Me.TabControl1.Visible = True
                Me.crv1.Visible = False
            Else
                Me.NonAktifCRV()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            obj = Nothing
        End Try
    End Sub

    Private Sub BtnPreview_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnPreview.Click

        Dim ObjPIF As New CustomSMESS.ClsTransaksiProformaInvoice

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If Me.rboQuotation.Checked = True Then
                Me.cetak()
            Else
                ObjPIF.CetakProformaInvoice_Custom(Me.crv1, _
                Me.Primary, Me.QuotationDate, Me.DefaultPrecision, Me.JabatanProforma)
                Me.crv2.ReportSource = Nothing
                Me.CrvDetail.ReportSource = Nothing
                Me.TabControl1.Visible = False
                Me.crv1.Visible = True
            End If

            Windows.Forms.Cursor.Current = Cursors.Arrow
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjPIF = Nothing
        End Try
    End Sub

    Private Sub NonAktifCRV()
        Try
            Me.crv1.Visible = False
            Me.TabControl1.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            If Me.TabControl1.SelectedIndex = 0 Then
                Me.crv2.DisplayStatusBar = True
                Me.crv2.DisplayToolbar = True
            Else
                Me.CrvDetail.DisplayToolbar = True
                Me.CrvDetail.DisplayStatusBar = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboQuotation_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboQuotation.CheckedChanged
        Try
            If Me.rboQuotation.Checked = True Then
                Me.cboLanguage.Enabled = True
                Me.cboVersion.Enabled = True
            Else
                Me.cboLanguage.Enabled = False
                Me.cboVersion.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class