Option Explicit On
Option Strict Off
Public Class FrmLaporanPurchaseDP
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim ObjSupplier As New Core_Module.clsMasterSupplier
    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub FrmLaporanPurchaseDP_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.cboDPStatus.SelectedIndex = 0
            Me.cboPaymentStatus.SelectedIndex = 0
            Me.cboRealizationStatus.SelectedIndex = 0
            Me.TxtSupplierCode.CharacterCasing = CharacterCasing.Upper
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboPODate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboPODate.CheckedChanged
        Try
            If Me.rboPODate.Enabled = True Then
                Me.dtpPOFrom.Enabled = True
                Me.dtpPOTo.Enabled = True
                Me.dtpAOPDPFrom.Enabled = False
                Me.dtpAOPDPTo.Enabled = False
                Me.DtpAOPDateFrom.Enabled = False
                Me.DtpAOPDateTo.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSupplierCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSupplierCode.GotFocus
        Me.TxtSupplierCode.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtSupplierCode_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtSupplierCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.ObjSupplier.ApakahSudahAda(Me.TxtSupplierCode.Text) = True Then
                    Me.TxtPrimarySupplier.Text = Me.ObjSupplier.GetPrimary(Me.TxtSupplierCode.Text)
                    Me.TxtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.TxtPrimarySupplier.Text)
                    Me.TxtSupplierName.Text = Me.ObjSupplier.GetSupplierName(Me.TxtPrimarySupplier.Text)
                Else
                    Me.btnFindSupplier.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSupplierCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSupplierCode.Leave
        Dim i As New Core_Module.FrmCariMasterSupplier
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.TxtSupplierCode.Text <> "" Then
                If Me.ObjSupplier.ApakahSudahAda(Me.TxtSupplierCode.Text) = True Then
                    Me.TxtPrimarySupplier.Text = Me.ObjSupplier.GetPrimary(Me.TxtSupplierCode.Text)
                    Me.TxtSupplierName.Text = Me.ObjSupplier.GetSupplierName(Me.TxtPrimarySupplier.Text)
                Else
                    If Me.TxtSupplierCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtSupplierCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimarySupplier > 0 Then
                        Me.TxtPrimarySupplier.Text = i.PrimarySupplier
                        Me.TxtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.TxtPrimarySupplier.Text)
                        Me.TxtSupplierName.Text = Me.ObjSupplier.GetSupplierName(Me.TxtPrimarySupplier.Text)
                    Else
                        Me.TxtSupplierCode.Clear()
                        Me.TxtSupplierName.Clear()
                        Me.TxtPrimarySupplier.Text = "-1"
                    End If
                End If
            Else
                Me.TxtSupplierCode.Clear()
                Me.TxtSupplierName.Clear()
                Me.TxtPrimarySupplier.Text = "-1"
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindSupplier_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindSupplier.Click
        Dim i As New Core_Module.FrmCariMasterSupplier
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtPrimarySupplier.Text <> "" Then
                i.txtCari.Text = Me.TxtSupplierCode.Text
            End If

            i.ShowDialog()
            If i.PrimarySupplier > 0 Then
                Me.TxtPrimarySupplier.Text = i.PrimarySupplier
                Me.TxtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.TxtPrimarySupplier.Text)
                Me.TxtSupplierName.Text = Me.ObjSupplier.GetSupplierName(Me.TxtPrimarySupplier.Text)
            Else
                Me.TxtSupplierCode.Clear()
                Me.TxtSupplierName.Clear()
                Me.TxtPrimarySupplier.Text = "-1"
                Me.TxtSupplierCode.Focus()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtSupplierCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSupplierCode.LostFocus
        Me.TxtSupplierCode.BackColor = Me.TxtPrimarySupplier.BackColor
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFilter.Click
        Try
            If Me.rboAOPDP.Checked = True Then
                If Me.dtpAOPDPFrom.Value.Date > Me.dtpAOPDPTo.Value.Date Then
                    MsgBox("AOP DP period is not valid, please check again.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If Me.RboAOP.Checked = True Then
                If Me.DtpAOPDateFrom.Value.Date > Me.DtpAOPDateTo.Value.Date Then
                    MsgBox("AOP period is not valid, please check again.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If Me.rboPODate.Checked = True Then
                If Me.dtpPOFrom.Value.Date > Me.dtpPOTo.Value.Date Then
                    MsgBox("PO period is not valid, please check again.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            Me.Cetak()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function Cetak() As Boolean
        Dim DS As New DseLaporanPurchaseDP
        Dim tblCetak As DseLaporanPurchaseDP.tblcetakDataTable = DS.tblcetak
        Dim dtaPO As New DseLaporanPurchaseDPTableAdapters.tbltrans_poTableAdapter
        Dim tblPO As New DseLaporanPurchaseDP.tbltrans_poDataTable
        Dim DtaAOPDP As New DseLaporanPurchaseDPTableAdapters.tbltrans_aop_dpTableAdapter
        Dim tblAOPDP As New DseLaporanPurchaseDP.tbltrans_aop_dpDataTable
        Dim PrimaryMin As Double = -1
        Dim PrimaryMax As Double = -1
        Dim rowAOPDP As DseLaporanPurchaseDP.tbltrans_aop_dpRow
        Dim dtaJour As New DseLaporanPurchaseDPTableAdapters.tbltrans_journalTableAdapter
        Dim tblJour As New DseLaporanPurchaseDP.tbltrans_journalDataTable
        Dim rowJour As DseLaporanPurchaseDP.tbltrans_journalRow
        Dim rowCetak As DseLaporanPurchaseDP.tblcetakRow
        Dim dtaAOP As New DseLaporanPurchaseDPTableAdapters.tbltrans_aopTableAdapter
        Dim tblAOP As New DseLaporanPurchaseDP.tbltrans_aopDataTable
        Dim rowAOP As DseLaporanPurchaseDP.tbltrans_aopRow
        Dim PaymentStatus As String = ""
        Dim rowCetak1 As DseLaporanPurchaseDP.tblcetakRow
        Dim tblCetak1 As New DseLaporanPurchaseDP.tblcetakDataTable
        Try
            dtaPO.Connection.ConnectionString = Core_Module.GetConstring
            DtaAOPDP.Connection.ConnectionString = Core_Module.GetConstring
            dtaJour.Connection.ConnectionString = Core_Module.GetConstring

            If Me.cboDPStatus.SelectedIndex = 0 Or Me.cboDPStatus.SelectedIndex = 2 Then
                If Me.rboPODate.Checked = True Then
                    If Me.TxtSupplierCode.Text <> "" Then
                        dtaPO.FillBy_TglSampaiNSupp(tblPO, Me.dtpPOFrom.Value, Me.dtpPOTo.Value, Me.TxtPrimarySupplier.Text)
                        DtaAOPDP.FillBy_TglSampaiNSupp(tblAOPDP, #1/1/1950#, #1/1/2200#, Me.TxtPrimarySupplier.Text)
                    Else
                        dtaPO.Fill(tblPO, Me.dtpPOFrom.Value, Me.dtpPOTo.Value)
                        DtaAOPDP.Fill(tblAOPDP, #1/1/1950#, #1/1/2200#)
                    End If
                End If
            End If

            If Me.cboDPStatus.SelectedIndex = 0 Or Me.cboDPStatus.SelectedIndex = 1 Then
                If Me.rboAOPDP.Checked = False Then
                    If Me.TxtSupplierCode.Text <> "" Then
                        DtaAOPDP.FillBy_TglSampaiNSupp(tblAOPDP, #1/1/1950#, #1/1/2200#, Me.TxtPrimarySupplier.Text)
                    Else
                        DtaAOPDP.Fill(tblAOPDP, #1/1/1950#, #1/1/2200#)
                    End If
                Else
                    If Me.TxtSupplierCode.Text <> "" Then
                        DtaAOPDP.FillBy_TglSampaiNSupp(tblAOPDP, Me.dtpAOPDPFrom.Value, Me.dtpAOPDPTo.Value, Me.TxtPrimarySupplier.Text)
                    Else
                        DtaAOPDP.Fill(tblAOPDP, Me.dtpAOPDPFrom.Value, Me.dtpAOPDPTo.Value)
                    End If
                End If
            End If
            'convert dari number aop ke Supplier_Invoice_No jika user pilih dan jika Supplier_Invoice_No tidak kosong. kalau kosong tetap pakai number aop
            If Me.ChkUse.Checked = True Then
                For Each rowAOPDP In tblAOPDP
                    If String.IsNullOrEmpty(rowAOPDP.Supplier_Invoice_No) = False Then
                        rowAOPDP.No_Transaksi = rowAOPDP.Supplier_Invoice_No
                    End If
                Next
            End If

            Me.InsertToTabel_PO(tblPO, tblCetak)
            Me.InsertToTabel_AOPDP(tblAOPDP, tblCetak)

            If tblAOPDP.Rows.Count > 0 Then
                tblAOPDP.DefaultView.RowFilter = Nothing
                tblAOPDP.DefaultView.Sort = "[Primary] ASC"
                rowAOPDP = tblAOPDP.DefaultView(0).Row
                PrimaryMin = rowAOPDP.Primary
                tblAOPDP.DefaultView.Sort = "[Primary] DESC"
                rowAOPDP = tblAOPDP.DefaultView(0).Row
                PrimaryMax = rowAOPDP.Primary
                If dtaJour.Fill(tblJour, PrimaryMin, PrimaryMax) > 0 Then
                    For Each rowJour In tblJour.Rows
                        tblCetak.DefaultView.RowFilter = "[Primary AOP DP]=" & rowJour.Primary_Document_Acuan
                        If tblCetak.DefaultView.Count > 0 Then
                            rowCetak = tblCetak.DefaultView(0).Row
                            If rowCetak.Journal_Number = "" Then
                                With rowCetak
                                    .Journal_Number = rowJour.No_Journal
                                    .Journal_Date = Format(rowJour.Tanggal_Journal, "dd MMM yy")
                                    .Amount = rowJour.payment_original
                                    .Group_By_Journal = rowJour.No_Journal
                                End With
                            Else
                                rowCetak1 = tblCetak.DefaultView(0).Row
                                rowCetak = tblCetak.NewtblcetakRow
                                With rowCetak
                                    .Journal_Number = rowJour.No_Journal
                                    .Journal_Date = Format(rowJour.Tanggal_Journal, "dd MMM yy")
                                    .Amount = rowJour.payment_original
                                    '.AOP_DP_Date = rowCetak1.AOP_DP_Date
                                    '.AOP_DP_Number = rowCetak1.AOP_DP_Number
                                    .Primary_PO = rowCetak1.Primary_PO
                                    .Primary_AOP_DP = rowCetak1.Primary_AOP_DP
                                    .Group_By_PO = rowCetak1.Group_By_PO
                                    .Group_By_Supplier = rowCetak1.Group_By_Supplier
                                    .Group_By_Journal = rowJour.No_Journal
                                    .Default_Currency = rowCetak1.Default_Currency
                                    .Currency = rowCetak1.Currency
                                End With
                                tblCetak.AddtblcetakRow(rowCetak)
                            End If
                        End If
                    Next
                End If
            End If

            If Me.RboAOP.Checked = True Then
                If Me.TxtSupplierCode.Text <> "" Then
                    dtaAOP.FillBy_TglSampaiNSupp(tblAOP, Me.DtpAOPDateFrom.Value, Me.DtpAOPDateTo.Value, Me.TxtPrimarySupplier.Text)
                Else
                    dtaAOP.Fill(tblAOP, Me.DtpAOPDateFrom.Value, Me.DtpAOPDateTo.Value)
                End If
            Else
                If Me.TxtSupplierCode.Text <> "" Then
                    dtaAOP.FillBy_TglSampaiNSupp(tblAOP, #1/1/1950#, #1/1/2200#, Me.TxtPrimarySupplier.Text)
                Else
                    dtaAOP.Fill(tblAOP, #1/1/1950#, #1/1/2200#)
                End If
            End If

            For i As Integer = 0 To tblAOP.Rows.Count - 1
                rowAOP = tblAOP.Rows(i)
                tblCetak.DefaultView.Sort = "[Group By Supplier] asc, [Group By PO] asc, [Group By Journal] asc"
                tblCetak.DefaultView.RowFilter = "[Primary AOP DP]=" & rowAOP.Primary_AOP_DP
                If tblCetak.DefaultView.Count > 0 Then
                    rowCetak = tblCetak.DefaultView(0).Row
                    With rowCetak
                        If Me.ChkUse.Checked = True Then
                            If String.IsNullOrEmpty(rowAOP.Supplier_Invoice_No) = False Then
                                .AOP_Number = rowAOP.Supplier_Invoice_No
                            Else
                                .AOP_Number = rowAOP.AOP_Number
                            End If
                        Else
                            .AOP_Number = rowAOP.AOP_Number
                        End If
                        .AOP_Date = Format(rowAOP.Tanggal_AOP, "dd MMM yy")
                        .AOP_Amount = Me.ObjCD.DigitGrouping(rowCetak.DP_Amount, rowCetak.Default_Currency, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)  'karna harus dikonfirmasi dulu sementara disamakan dengan dp amount cek aulia rahman.
                        .Group_By_AOP = i
                    End With
                End If
            Next

            tblCetak1 = tblCetak.Copy
            For Each rowCetak1 In tblCetak1.Rows
                tblCetak.DefaultView.RowFilter = "[Primary AOP DP]=" & rowCetak1.Primary_AOP_DP
                For ij As Integer = 0 To tblCetak.DefaultView.Count - 1
                    rowCetak = tblCetak.DefaultView(ij).Row
                    With rowCetak
                        .DP_Amount2 += rowCetak1.DP_Amount
                        .AOP_Amount2 += rowCetak1.AOP_Amount
                    End With
                Next
            Next

            If cboRealizationStatus.SelectedIndex <> 0 Then
                'filter Realization Status
                For i As Integer = tblCetak.Rows.Count - 1 To 0 Step -1
                    rowCetak = tblCetak.Rows(i)
                    If Me.cboRealizationStatus.SelectedIndex = 1 Then
                        If rowCetak.DP_Amount2 <> rowCetak.AOP_Amount2 Then
                            tblCetak.Rows.Remove(rowCetak)
                        End If
                    ElseIf Me.cboRealizationStatus.SelectedIndex = 2 Then
                        If rowCetak.DP_Amount2 = rowCetak.AOP_Amount2 Then
                            tblCetak.Rows.Remove(rowCetak)
                        End If
                    End If
                Next
            End If

            'filter downpayment Status
            For i As Integer = tblCetak.Rows.Count - 1 To 0 Step -1
                rowCetak = tblCetak.Rows(i)
                If Me.cboDPStatus.SelectedIndex = 1 Then
                    If rowCetak.Primary_AOP_DP = 1 Then
                        tblCetak.Rows.Remove(rowCetak)
                    End If
                ElseIf Me.cboDPStatus.SelectedIndex = 2 Then
                    If rowCetak.Primary_AOP_DP <> -1 Then
                        tblCetak.Rows.Remove(rowCetak)
                    End If
                End If
            Next

            For Each rowCetak1 In tblCetak1.Rows
                tblCetak.DefaultView.RowFilter = "[Primary AOP DP]=" & rowCetak1.Primary_AOP_DP & " AND [Primary PO]=" & rowCetak1.Primary_PO
                For ij As Integer = 0 To tblCetak.DefaultView.Count - 1
                    rowCetak = tblCetak.DefaultView(ij).Row
                    With rowCetak
                        .Amount2 += rowCetak1.Amount
                        .DP_Amount3 += rowCetak1.DP_Amount
                    End With
                Next
            Next

            For Each rowCetak1 In tblCetak1.Rows
                tblCetak.DefaultView.RowFilter = "[Primary PO]=" & rowCetak1.Primary_PO
                For ij As Integer = 0 To tblCetak.DefaultView.Count - 1
                    rowCetak = tblCetak.DefaultView(ij).Row
                    With rowCetak
                        .DP_Amount4 += rowCetak1.DP_Amount
                    End With
                Next
            Next

            'filter payment status--------------
            tblCetak1.Rows.Clear()
            For i As Integer = tblCetak.Rows.Count - 1 To 0 Step -1
                rowCetak = tblCetak.Rows(i)
                If rowCetak.Amount_Use_DP > 0 Then
                    If CDec(rowCetak.DP_Amount4) <> CDec(rowCetak.Amount_Use_DP) Then
                        rowCetak.IsRed = 1
                    End If
                End If

                If Me.cboPaymentStatus.SelectedIndex = 1 Then
                    If CDec(rowCetak.DP_Amount3) <> CDec(rowCetak.Amount2) Then
                        tblCetak1.ImportRow(rowCetak)
                        tblCetak.Rows.Remove(rowCetak)
                    End If
                ElseIf Me.cboPaymentStatus.SelectedIndex = 2 Then
                    If CDec(rowCetak.DP_Amount3) = CDec(rowCetak.Amount2) Then
                        tblCetak1.ImportRow(rowCetak)
                        tblCetak.Rows.Remove(rowCetak)
                    End If
                End If
            Next
            For Each rowCetak In tblCetak.Rows
                If rowCetak.PO_Number = "" And rowCetak.DP_Amount > 0 Then
                    tblCetak1.DefaultView.RowFilter = "[Primary PO]=" & rowCetak.Primary_PO & " AND [PO Number]<>''"
                    If tblCetak1.DefaultView.Count > 0 Then
                        rowCetak1 = tblCetak1.DefaultView(0).Row
                        rowCetak.PO_Number = rowCetak1.PO_Number
                        rowCetak.PO_Date = rowCetak1.PO_Date
                    End If
                End If
            Next
            '----------------------

            'filter aopdate
            If Me.RboAOP.Checked = True Then
                'tblCetak.DefaultView.RowFilter = "[AOP Date] < '#" & Format(Me.DtpAOPDateFrom.Value.Date, "dd MMM yyyy") & "#'" ' Or [AOP Number]=''"
                Dim OK As Boolean = False
                tblCetak.DefaultView.RowFilter = ""
                For i As Integer = tblCetak.DefaultView.Count - 1 To 0 Step -1
                    rowCetak = tblCetak.DefaultView(i).Row
                    With rowCetak
                        OK = False
                        If String.IsNullOrEmpty(.AOP_Number) = True Then
                            tblCetak.Rows.Remove(rowCetak)
                            OK = True
                        End If

                        If OK = False Then
                            If CDate(.AOP_Date) < Me.DtpAOPDateFrom.Value.Date Then
                                tblCetak.Rows.Remove(rowCetak)
                                OK = True
                            End If
                        End If

                        If OK = False Then
                            If CDate(.AOP_Date) > Me.DtpAOPDateTo.Value.Date Then
                                tblCetak.Rows.Remove(rowCetak)
                                OK = True
                            End If
                        End If
                    End With
                Next
            End If
            IsikanSubTotal(tblCetak)
            'sorting
            tblCetak1 = tblCetak.Copy
            tblCetak.Clear()
            'tblCetak1.DefaultView.Sort = "[Group By Supplier], [Group By PO],[Group By Journal], [AOP DP Number], [Journal Number]" '[AOP Date], [AOP Number]"
            tblCetak1.DefaultView.Sort = "[Group By Supplier] asc, [Group By PO] asc, [Group By Journal] asc, [Group By AOP] asc"
            For i As Integer = 0 To tblCetak1.DefaultView.Count - 1
                rowCetak1 = tblCetak1.DefaultView(i).Row
                tblCetak.ImportRow(rowCetak1)
            Next
            Dim No As Integer = 0
            For Each rowCetak In tblCetak.Rows
                If rowCetak.PO_Number <> "" Then
                    No += 1
                    rowCetak.No = No
                End If
            Next

            Dim ReportDetail As New CrpLaporanPurchaseDP
            ReportDetail.SetDataSource(DS)
            If Me.RboAOP.Checked = True Then
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Nama_Label.ParameterFieldName, "AOP Date")
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Label_Date.ParameterFieldName, Format(Me.DtpAOPDateFrom.Value, "dd MMM yyyy") & " to " & Format(Me.DtpAOPDateTo.Value, "dd MMM yyyy"))
            End If

            If Me.rboPODate.Checked = True Then
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Nama_Label.ParameterFieldName, "PO Date")
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Label_Date.ParameterFieldName, Format(Me.dtpPOFrom.Value, "dd MMM yyyy") & " to " & Format(Me.dtpPOTo.Value, "dd MMM yyyy"))
            End If

            If Me.rboAOPDP.Checked = True Then
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Nama_Label.ParameterFieldName, "AOP Date")
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Label_Date.ParameterFieldName, Format(Me.dtpAOPDPFrom.Value, "dd MMM yyyy") & " to " & Format(Me.dtpAOPDPTo.Value, "dd MMM yyyy"))
            End If

            If Me.TxtSupplierName.Text <> "" Then
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Supplier.ParameterFieldName, Me.TxtSupplierName.Text)
            Else
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Supplier.ParameterFieldName, "All")
            End If

            If Me.cboPaymentStatus.SelectedIndex = 0 Then
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Payment_Status.ParameterFieldName, "All")
            Else
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Payment_Status.ParameterFieldName, Me.cboPaymentStatus.SelectedItem)
            End If

            If Me.cboDPStatus.SelectedIndex = 0 Then
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Exclude_PO_With_AOP_DP.ParameterFieldName, "All")
            Else
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Exclude_PO_With_AOP_DP.ParameterFieldName, Me.cboDPStatus.SelectedItem)
            End If

            If Me.cboRealizationStatus.SelectedIndex = 0 Then
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Exclude_Realized_DP.ParameterFieldName, "All")
            Else
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Exclude_Realized_DP.ParameterFieldName, Me.cboRealizationStatus.SelectedItem)
            End If

            ReportDetail.SetParameterValue(ReportDetail.Parameter_Exclude_Paid_DP.ParameterFieldName, "")

            Me.crv.ReportSource = ReportDetail
            Me.crv.Show()
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dtaAOP.Dispose()
            DtaAOPDP.Dispose()
            dtaPO.Dispose()
            dtaJour.Dispose()
        End Try
    End Function

    Private Function IsikanSubTotal(ByRef tblCetak As DseLaporanPurchaseDP.tblcetakDataTable) As Boolean
        Dim tblT As New DseLaporanPurchaseDP.tbl1DataTable
        Dim rowT As DseLaporanPurchaseDP.tbl1Row
        Dim tblT2 As New DseLaporanPurchaseDP.tbl2DataTable
        Dim rowT2 As DseLaporanPurchaseDP.tbl2Row
        Dim rowCetak As DseLaporanPurchaseDP.tblcetakRow
        Dim tblDummmy As New DseLaporanPurchaseDP.tblcetakDataTable
        Try
            For Each rowCetak In tblCetak.Rows
                tblT.DefaultView.RowFilter = "[Group]='" & rowCetak.Group_By_Supplier & "' AND [Currency]='" & rowCetak.Currency & "'"
                If tblT.DefaultView.Count > 0 Then
                    rowT = tblT.DefaultView(0).Row
                    With rowT
                        .Sub_DP_Amount += CDec(rowCetak.DP_Amount)
                        .Sub_Amount += CDec(rowCetak.Amount)
                        .Sub_AOP_Amount += CDec(rowCetak.AOP_Amount)
                    End With
                Else
                    rowT = tblT.Newtbl1Row
                    With rowT
                        .Sub_DP_Amount = CDec(rowCetak.DP_Amount)
                        .Sub_Amount = CDec(rowCetak.Amount)
                        .Sub_AOP_Amount = CDec(rowCetak.AOP_Amount)
                        .Group = rowCetak.Group_By_Supplier
                        .Currency = rowCetak.Currency
                        .Default_Currency = rowCetak.Default_Currency
                    End With
                    tblT.Addtbl1Row(rowT)
                End If
            Next

            If tblT.Rows.Count > 0 Then
                'sub total
                For Each rowT In tblT.Rows
                    tblCetak.DefaultView.RowFilter = "[Group By Supplier]='" & rowT.Group & "'"
                    For i As Integer = 0 To tblCetak.DefaultView.Count - 1
                        rowCetak = tblCetak.DefaultView(i).Row
                        If rowT.Sub_DP_Amount > 0 Then
                            rowCetak.Sub_DP_Amount = Me.ObjCD.DigitGrouping(rowT.Sub_DP_Amount, rowT.Default_Currency)
                        End If
                        If rowT.Sub_Amount > 0 Then
                            rowCetak.Sub_Amount = Me.ObjCD.DigitGrouping(rowT.Sub_Amount, rowT.Default_Currency)
                        End If
                        If rowT.Sub_AOP_Amount > 0 Then
                            rowCetak.Sub_AOP_Amount = Me.ObjCD.DigitGrouping(rowT.Sub_AOP_Amount, rowT.Default_Currency)
                        End If
                        rowCetak.Sub_Curr = rowT.Currency
                    Next
                Next

                For Each rowT In tblT.Rows
                    tblT2.DefaultView.RowFilter = "[Currency]='" & rowT.Currency & "'"
                    If tblT2.DefaultView.Count > 0 Then
                        rowT2 = tblT2.DefaultView(0).Row
                        rowT2.Total_DP_Amount = Me.ObjCD.DigitGrouping(rowT2.Total_DP_Amount + rowT.Sub_DP_Amount, rowT.Default_Currency)
                        rowT2.Total_Amount = Me.ObjCD.DigitGrouping(rowT2.Total_Amount + rowT.Sub_Amount, rowT.Default_Currency)
                        rowT2.Total_AOP_Amount = Me.ObjCD.DigitGrouping(rowT2.Total_AOP_Amount + rowT.Sub_AOP_Amount, rowT.Default_Currency)
                    Else
                        rowT2 = tblT2.Newtbl2Row
                        rowT2.Total_DP_Amount = Me.ObjCD.DigitGrouping(rowT.Sub_DP_Amount, rowT.Default_Currency)
                        rowT2.Total_Amount = Me.ObjCD.DigitGrouping(rowT.Sub_Amount, rowT.Default_Currency)
                        rowT2.Total_AOP_Amount = Me.ObjCD.DigitGrouping(rowT.Sub_AOP_Amount, rowT.Default_Currency)
                        rowT2.Currency = rowT.Currency
                        rowT2.Default_Currency = rowT.Default_Currency
                        tblT2.Addtbl2Row(rowT2)
                    End If
                Next

                For Each rowCetak In tblCetak.Rows
                    'total
                    For Each rowT2 In tblT2.Rows
                        If String.IsNullOrEmpty(rowCetak.Total_Curr) = True Then
                            rowCetak.Total_Curr = rowT2.Currency
                        Else
                            rowCetak.Total_Curr &= vbNewLine & rowT2.Currency
                        End If

                        If rowT2.Total_DP_Amount > 0 Then
                            If String.IsNullOrEmpty(rowCetak.Total_DP_Amount) = True Then
                                rowCetak.Total_DP_Amount = Me.ObjCD.DigitGrouping(rowT2.Total_DP_Amount, rowT2.Default_Currency, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                            Else
                                rowCetak.Total_DP_Amount &= vbNewLine & Me.ObjCD.DigitGrouping(rowT2.Total_DP_Amount, rowT2.Default_Currency, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                            End If
                        End If

                        If rowT2.Total_Amount > 0 Then
                            If String.IsNullOrEmpty(rowCetak.Total_Amount) = True Then
                                rowCetak.Total_Amount = Me.ObjCD.DigitGrouping(rowT2.Total_Amount, rowT2.Default_Currency, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                            Else
                                rowCetak.Total_Amount &= vbNewLine & Me.ObjCD.DigitGrouping(rowT2.Total_Amount, rowT2.Default_Currency, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                            End If
                        End If

                        If rowT2.Total_AOP_Amount > 0 Then
                            If String.IsNullOrEmpty(rowCetak.Total_AOP_Amount) = True Then
                                rowCetak.Total_AOP_Amount = Me.ObjCD.DigitGrouping(rowT2.Total_AOP_Amount, rowT2.Default_Currency, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                            Else
                                rowCetak.Total_AOP_Amount &= vbNewLine & Me.ObjCD.DigitGrouping(rowT2.Total_AOP_Amount, rowT2.Default_Currency, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                            End If
                        End If
                    Next
                Next
            End If

           

            'For Each rowCetak In tblCetak.Rows
            '    tblT.DefaultView.RowFilter = "[Currency]='" & rowCetak.Currency & "'"
            '    Dim TotalDP As String = "0"
            '    Dim TotalAmount As String = "0"
            '    Dim TotalAOPAmount As String = "0"
            '    For i As Integer = 0 To tblT.DefaultView.Count - 1
            '        rowT = tblT.DefaultView(i).Row
            '        With rowCetak
            '            TotalDP = Me.ObjCD.DigitGrouping(CDec(TotalDP) + CDec(rowT.Sub_DP_Amount), .Default_Currency)
            '            TotalAmount = Me.ObjCD.DigitGrouping(CDec(TotalAmount) + CDec(rowT.Sub_Amount), .Default_Currency)
            '            TotalAOPAmount = Me.ObjCD.DigitGrouping(CDec(TotalAOPAmount) + CDec(rowT.Sub_AOP_Amount), .Default_Currency)
            '        End With
            '    Next

            '    With rowCetak
            '        .Total_DP_Amount = TotalDP
            '        .Total_Amount = TotalAmount
            '        .Total_AOP_Amount = TotalAOPAmount
            '        .Total_Curr = rowCetak.Currency
            '    End With
            'Next

            'For Each rowCetak In tblCetak.Rows
            '    With rowCetak
            '        If CDec(.Sub_Amount) <= 0 Then
            '            .Sub_Amount = ""
            '        End If
            '        If CDec(.Sub_DP_Amount) <= 0 Then
            '            .Sub_DP_Amount = ""
            '        End If
            '        If CDec(.Sub_AOP_Amount) <= 0 Then
            '            .Sub_AOP_Amount = ""
            '        End If
            '        If CDec(.Total_Amount) <= 0 Then
            '            .Total_Amount = ""
            '        End If
            '        If CDec(.Total_DP_Amount) <= 0 Then
            '            .Total_DP_Amount = ""
            '        End If
            '        If CDec(.Total_AOP_Amount) <= 0 Then
            '            .Total_AOP_Amount = ""
            '        End If
            '    End With
            'Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function InsertToTabel_PO(ByVal tblPO As DseLaporanPurchaseDP.tbltrans_poDataTable, ByVal tblCetak As DseLaporanPurchaseDP.tblcetakDataTable) As Boolean
        Dim rowPO As DseLaporanPurchaseDP.tbltrans_poRow
        Dim rowCetak As DseLaporanPurchaseDP.tblcetakRow
        Try
            For Each rowPO In tblPO.Rows
                rowCetak = tblCetak.NewtblcetakRow
                With rowCetak
                    .PO_Number = rowPO.No_Transaksi
                    .PO_Date = Format(rowPO.Tanggal_Transaksi, "dd MMM yy")
                    .Currency = rowPO.Currency_Code
                    .DP_Amount = rowPO.dp_amount
                    .Group_By_Supplier = rowPO.Supplier_Name
                    .Default_Currency = rowPO.Default_Precision
                    .Primary_PO = rowPO.Primary
                End With
                tblCetak.AddtblcetakRow(rowCetak)
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function InsertToTabel_AOPDP(ByVal tblAOPDP As DseLaporanPurchaseDP.tbltrans_aop_dpDataTable, ByVal tblCetak As DseLaporanPurchaseDP.tblcetakDataTable) As Boolean
        Dim rowAOPDP As DseLaporanPurchaseDP.tbltrans_aop_dpRow
        Dim rowCetak As DseLaporanPurchaseDP.tblcetakRow
        Dim rowCetak2 As DseLaporanPurchaseDP.tblcetakRow
        Dim rowCetak3 As DseLaporanPurchaseDP.tblcetakRow
        Dim List As New List(Of Double)
        Try
            tblAOPDP.DefaultView.RowFilter = ""
            If Me.rboPODate.Checked = True Then
                tblAOPDP.DefaultView.RowFilter = "[Tanggal PO] >= '#" & Format(Me.dtpPOFrom.Value, "dd MMM yy") & "#' AND [Tanggal PO] <= '#" & Format(Me.dtpPOTo.Value, "dd MMM yy") & "#'"
                For i As Integer = 0 To tblAOPDP.DefaultView.Count - 1
                    rowAOPDP = tblAOPDP.DefaultView(i).Row

                    tblCetak.DefaultView.RowFilter = "[Primary PO]=" & rowAOPDP.Primary_PO_4_DP
                    If tblCetak.DefaultView.Count > 0 Then
                        For ij As Integer = 0 To tblCetak.DefaultView.Count - 1
                            rowCetak2 = tblCetak.DefaultView(0).Row
                            rowCetak = tblCetak.DefaultView(ij).Row
                            If List.Contains(rowAOPDP.Primary) = False Then
                                If rowCetak.Primary_AOP_DP = -1 Then
                                    With rowCetak
                                        .AOP_DP_Number = rowAOPDP.No_Transaksi
                                        .AOP_DP_Date = Format(rowAOPDP.Tanggal_Transaksi, "dd MMM yy")
                                        .Primary_AOP_DP = rowAOPDP.Primary
                                        If rowAOPDP.need_dp = 1 Then
                                            .Amount_Use_DP = rowAOPDP.dp_amount
                                        End If

                                        .DP_Amount = rowAOPDP.Total
                                        .Group_By_PO = i
                                        If List.Contains(rowAOPDP.Primary) = False Then
                                            List.Add(rowAOPDP.Primary)
                                        End If
                                    End With
                                Else
                                    rowCetak = tblCetak.NewtblcetakRow
                                    With rowCetak
                                        If rowAOPDP.need_dp = 1 Then
                                            .Amount_Use_DP = rowAOPDP.dp_amount
                                        End If
                                        .AOP_DP_Number = rowAOPDP.No_Transaksi
                                        .AOP_DP_Date = Format(rowAOPDP.Tanggal_Transaksi, "dd MMM yy")
                                        .Primary_AOP_DP = rowAOPDP.Primary
                                        .DP_Amount = rowAOPDP.Total
                                        .Currency = rowAOPDP.Currency_Code
                                        .Group_By_Supplier = rowCetak2.Group_By_Supplier
                                        .Default_Currency = rowCetak2.Default_Currency
                                        .Primary_PO = rowCetak2.Primary_PO
                                        .Group_By_PO = i
                                        If List.Contains(rowAOPDP.Primary) = False Then
                                            List.Add(rowAOPDP.Primary)
                                        End If
                                    End With
                                    tblCetak.AddtblcetakRow(rowCetak)
                                End If
                            End If

                        Next
                    Else
                        rowCetak = tblCetak.NewtblcetakRow
                        With rowCetak
                            If rowAOPDP.need_dp = 1 Then
                                .Amount_Use_DP = rowAOPDP.dp_amount
                            End If
                            .PO_Number = rowAOPDP.PO_Number
                            .PO_Date = Format(rowAOPDP.Tanggal_PO, "dd MMM yy")
                            .Currency = rowAOPDP.Currency_Code
                            .DP_Amount = rowAOPDP.dp_amount
                            .Group_By_Supplier = rowAOPDP.Supplier_Name
                            .Default_Currency = rowAOPDP.Default_Precision
                            .Primary_PO = rowAOPDP.Primary_PO_4_DP
                            .AOP_DP_Number = rowAOPDP.No_Transaksi
                            .AOP_DP_Date = Format(rowAOPDP.Tanggal_Transaksi, "dd MMM yy")
                            .Primary_AOP_DP = rowAOPDP.Primary
                            .DP_Amount = rowAOPDP.Total
                            .Group_By_PO = i
                        End With
                        tblCetak.AddtblcetakRow(rowCetak)
                    End If
                Next

            Else
                For i As Integer = 0 To tblAOPDP.DefaultView.Count - 1
                    rowAOPDP = tblAOPDP.DefaultView(i).Row
                    tblCetak.DefaultView.RowFilter = "[Primary PO]=" & rowAOPDP.Primary_PO_4_DP
                    If tblCetak.DefaultView.Count > 0 Then
                        rowCetak2 = tblCetak.DefaultView(0).Row
                        rowCetak = tblCetak.NewtblcetakRow
                        With rowCetak
                            .AOP_DP_Number = rowAOPDP.No_Transaksi
                            .AOP_DP_Date = Format(rowAOPDP.Tanggal_Transaksi, "dd MMM yy")
                            .Currency = rowAOPDP.Currency_Code
                            .DP_Amount = rowAOPDP.Total
                            .Default_Currency = rowAOPDP.Default_Precision
                            .Primary_AOP_DP = rowAOPDP.Primary
                            .Primary_PO = rowAOPDP.Primary_PO_4_DP
                            .Group_By_Supplier = rowCetak2.Group_By_Supplier
                            .Group_By_PO = i
                            If rowAOPDP.need_dp = 1 Then
                                .Amount_Use_DP = rowAOPDP.dp_amount
                            End If
                        End With
                        tblCetak.AddtblcetakRow(rowCetak)
                    Else
                        rowCetak = tblCetak.NewtblcetakRow
                        With rowCetak
                            .PO_Number = rowAOPDP.PO_Number
                            .PO_Date = rowAOPDP.Tanggal_PO
                            .AOP_DP_Number = rowAOPDP.No_Transaksi
                            .AOP_DP_Date = Format(rowAOPDP.Tanggal_Transaksi, "dd MMM yy")
                            .Currency = rowAOPDP.Currency_Code
                            .DP_Amount = rowAOPDP.Total
                            .Default_Currency = rowAOPDP.Default_Precision
                            .Primary_AOP_DP = rowAOPDP.Primary
                            .Primary_PO = rowAOPDP.Primary_PO_4_DP
                            .Group_By_Supplier = rowAOPDP.Supplier_Name
                            .Group_By_PO = i
                            If rowAOPDP.need_dp = 1 Then
                                .Amount_Use_DP = rowAOPDP.dp_amount
                            End If
                        End With
                        tblCetak.AddtblcetakRow(rowCetak)
                    End If
                Next
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub rboAOPDP_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rboAOPDP.CheckedChanged
        Try
            If Me.rboAOPDP.Enabled = True Then
                Me.dtpPOFrom.Enabled = False
                Me.dtpPOTo.Enabled = False
                Me.dtpAOPDPFrom.Enabled = True
                Me.dtpAOPDPTo.Enabled = True
                Me.DtpAOPDateFrom.Enabled = False
                Me.DtpAOPDateTo.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RboAOP_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RboAOP.CheckedChanged
        Try
            If Me.RboAOP.Enabled = True Then
                Me.dtpPOFrom.Enabled = False
                Me.dtpPOTo.Enabled = False
                Me.dtpAOPDPFrom.Enabled = False
                Me.dtpAOPDPTo.Enabled = False
                Me.DtpAOPDateFrom.Enabled = True
                Me.DtpAOPDateTo.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class