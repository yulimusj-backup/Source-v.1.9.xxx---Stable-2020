Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmLaporanKartuAR
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim TipeKartu As ClsTransaksi_Kartu_APAR.EnumJenisKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR

    Dim Obj As clsMasterCustomer
    Dim ObjKartu As ClsTransaksi_Kartu_APAR

    Dim TBL As DseTrans_Kartu_APAR.tbltransaksi_kartu_ar_ap1DataTable
    Dim buffer As String = ""
    Dim report As New crpCetakKartu_AP_AR
    Dim DS As New DseTrans_Kartu_APAR

    Dim Def_Prec As Integer = 0
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim Negatif As Integer = 0
    Dim ObjLicense As New Core_Module.ClsLicense

    Private Sub txtCustCode_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustCode.Enter
        Me.buffer = Me.txtCustCode.Text
    End Sub

    Private Sub txtCustCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustCode.GotFocus
        Me.txtCustCode.BackColor = Me.BackColorAll
    End Sub

    Private Sub txtCustCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustCode.KeyDown
        Dim i As New FrmCariMasterCustomer
        Dim Jml As Long = 0

        Try
            If e.KeyValue = 13 Then
                Jml = Obj.GetJml(Me.txtCustCode.Text)
                If Jml = 1 Then
                    Me.txtPrimaryCust.Text = Obj.GetPrimary(Me.txtCustCode.Text)
                    Me.txtDescription.Text = Obj.GetCustName(Me.txtPrimaryCust.Text)
                    Me.dtpStartDate.Focus()
                    If Me.txtCustCode.Text <> Me.buffer Then
                        Me.dgvStockCard.DataSource = Nothing
                        Me.dgvStockCard.Refresh()
                    End If
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.txtCustCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCustCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.txtPrimaryCust.Text = i.PrimaryCustomer
                        Me.txtCustCode.Text = Obj.GetCustomerCode(Me.txtPrimaryCust.Text)
                        Me.dtpStartDate.Focus()
                        If Me.txtCustCode.Text <> Me.buffer Then
                            Me.dgvStockCard.DataSource = Nothing
                            Me.dgvStockCard.Refresh()
                        End If
                    Else
                        Me.txtCustCode.Clear()
                        Me.txtPrimaryCust.Text = "-1"
                        Me.txtDescription.Clear()
                        Me.txtCustCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtCustCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustCode.Leave
        Dim i As New FrmCariMasterCustomer

        Try
            If Me.txtCustCode.Text <> "" Then
                If Obj.GetJml(Me.txtCustCode.Text) = 1 Then
                    Me.txtPrimaryCust.Text = Obj.GetPrimary(Me.txtCustCode.Text)
                    Me.txtDescription.Text = Obj.GetCustName(Me.txtPrimaryCust.Text)
                    Me.dtpStartDate.Focus()
                    If Me.txtCustCode.Text <> Me.buffer Then
                        Me.dgvStockCard.DataSource = Nothing
                        Me.dgvStockCard.Refresh()
                    End If
                Else
                    If Me.txtCustCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCustCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.txtPrimaryCust.Text = i.PrimaryCustomer
                        Me.txtCustCode.Text = Obj.GetCustomerCode(Me.txtPrimaryCust.Text)
                        Me.dtpStartDate.Focus()
                        If Me.txtCustCode.Text <> Me.buffer Then
                            Me.dgvStockCard.DataSource = Nothing
                            Me.dgvStockCard.Refresh()
                        End If
                    Else
                        Me.txtCustCode.Clear()
                        Me.txtPrimaryCust.Text = "-1"
                        Me.txtDescription.Clear()
                        Me.txtCustCode.Focus()
                        Me.dgvStockCard.DataSource = Nothing
                    End If
                End If
            Else
                Me.txtPrimaryCust.Text = "-1"
                Me.txtDescription.Clear()
                Me.dgvStockCard.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindProduct.Click
        Dim i As New FrmCariMasterCustomer

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.txtCustCode.Text <> "" Then
                i.txtCari.Text = Me.txtCustCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryCustomer > 0 Then
                Me.txtPrimaryCust.Text = i.PrimaryCustomer
                Me.txtCustCode.Text = Obj.GetCustomerCode(Me.txtPrimaryCust.Text)
                Me.dtpStartDate.Focus()
            Else
                Me.txtCustCode.Clear()
                Me.txtPrimaryCust.Text = "-1"
                Me.txtDescription.Clear()
                Me.txtCustCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub FrmTransaksiStockCard_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.Def_Prec = Me.ObjCur.GetDecimal_DefaultCurrency
            Me.Obj = New clsMasterCustomer
            Me.ObjKartu = New ClsTransaksi_Kartu_APAR
            Me.dtpStartDate.Value = Now - TimeSpan.FromDays(30)
            Me.TBL = Me.DS.tbltransaksi_kartu_ar_ap1
            Me.dgvStockCard.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim ij As New FrmLaporanKartuAP_Cetak
        Dim TanggalAcuan As Date = #1/1/1950#
        Dim row1 As DseTrans_Kartu_APAR.tbltransaksi_kartu_ar_ap1Row

        Try
            If Me.txtPrimaryCust.Text <> -1 Then
                If Me.dtpStartDate.Value.Date <= Me.dtpEndDate.Value.Date Then
                    Me.ObjKartu.IsiDGV(Me.txtPrimaryCust.Text, Me.dtpStartDate.Value, Me.dtpEndDate.Value, _
                    Me.dgvStockCard, Me.TBL, Me.TipeKartu, Me.Def_Prec)

                    Me.dgvStockCard.ReadOnly = False

                    Dim row As DataGridViewRow
                    For Each row In Me.dgvStockCard.Rows
                        row.Cells(0).Value = CDate(row.Cells(0).Value).ToShortDateString
                    Next
                    Me.dgvStockCard.ReadOnly = True
                    Me.dgvStockCard.Columns(Me.dgvStockCard.Columns.Count - 1).Visible = False

                    'pilih mana yang keliatan tanggal mana yang ga.
                    For i As Double = 0 To DS.tbltransaksi_kartu_ar_ap1.Rows.Count - 1
                        row1 = DS.tbltransaksi_kartu_ar_ap1.Rows(i)
                        If i = 0 Then
                            'baris pertama selalu keliatan tanggal
                            TanggalAcuan = row1._Date
                            row1.Tanggal_Str = Format(row1._Date, "dd-MMM-yyyy")
                        Else
                            'jika baris kelipatan 49+1, maka selalu keliatan karena baris pertama pada setiap halaman
                            If i - ((i \ 49) * 49) = 0 Then
                                TanggalAcuan = row1._Date
                                row1.Tanggal_Str = Format(row1._Date, "dd-MMM-yyyy")
                            Else
                                'sisanya hanya keliatan jika ada perubahan tanggal
                                If row1._Date <> TanggalAcuan Then
                                    TanggalAcuan = row1._Date
                                    row1.Tanggal_Str = Format(row1._Date, "dd-MMM-yyyy")
                                Else
                                    row1.Tanggal_Str = ""
                                End If
                            End If
                        End If
                    Next

                    'masukkan data ke tampilan cetak
                    Me.report.SetDataSource(DS)
                    Me.report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
                    Me.report.SetParameterValue(report.Parameter_EndDate.ParameterFieldName, Me.dtpEndDate.Value)
                    Me.report.SetParameterValue(report.Parameter_StartDate.ParameterFieldName, Me.dtpStartDate.Value)
                    Me.report.SetParameterValue(report.Parameter_Kode_Nama_PihakAcuan.ParameterFieldName, Me.txtCustCode.Text + " - " + Me.txtDescription.Text)
                    Me.report.SetParameterValue(report.Parameter_TipePihakAcuan.ParameterFieldName, "Customer")
                    Me.report.SetParameterValue(report.Parameter_NamaLaporan.ParameterFieldName, "Account Receivable Card")
                    If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                        Me.Negatif = 1
                    Else
                        Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
                    End If
                    Me.report.SetParameterValue(report.Parameter_Negatif.ParameterFieldName, Me.Negatif)

                    ij.crvReport.ReportSource = Me.report
                    ij.crvReport.Refresh()
                    ij.crvReport.Show()
                    ij.Text = "AR Card"
                    ij.ShowDialog()
                Else
                    MsgBox("Wrong date filter, baginning date must be smaller than end date", _
                    MsgBoxStyle.Information, "Information")
                End If
            Else
                MsgBox("Please specify customer code.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
            ij = Nothing
        End Try
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.dtpStartDate.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpStartDate.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpEndDate.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpEndDate.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If Me.txtPrimaryCust.Text <> -1 Then
                If Me.dtpStartDate.Value.Date <= Me.dtpEndDate.Value.Date Then
                    Me.ObjKartu.IsiDGV(Me.txtPrimaryCust.Text, Me.dtpStartDate.Value, Me.dtpEndDate.Value, _
                    Me.dgvStockCard, Me.TBL, Me.TipeKartu, Me.Def_Prec)

                    Me.dgvStockCard.ReadOnly = False

                    'Dim row As DataGridViewRow
                    'For Each row In Me.dgvStockCard.Rows
                    '    row.Cells(0).Value = CDate(row.Cells(0).Value).ToShortDateString
                    'Next
                    Me.dgvStockCard.Columns(Me.dgvStockCard.Columns.Count - 1).Visible = False

                    Me.dgvStockCard.ReadOnly = True
                Else
                    MsgBox("Wrong date filter, baginning date must be smaller than end date", _
                    MsgBoxStyle.Information, "Information")
                End If
            Else
                MsgBox("Please specify customer code.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
        End Try
    End Sub

    Private Sub txtCustCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustCode.LostFocus
        Me.txtCustCode.BackColor = Me.txtPrimaryCust.BackColor
    End Sub
End Class