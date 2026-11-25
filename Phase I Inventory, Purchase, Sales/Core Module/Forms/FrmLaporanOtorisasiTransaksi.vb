Option Explicit On
Option Strict Off

Public Class FrmLaporanOtorisasiTransaksi
    Dim DS As New DseLaporanOtorisasiTransaksi
    Dim TBL As DseLaporanOtorisasiTransaksi.tblmaster_tipe_otorisasiDataTable = DS.tblmaster_tipe_otorisasi
    Dim row As DseLaporanOtorisasiTransaksi.tblmaster_tipe_otorisasiRow
    Dim DTA As New DseLaporanOtorisasiTransaksiTableAdapters.tblmaster_tipe_otorisasiTableAdapter
    Dim DTA1 As New DseLaporanOtorisasiTransaksiTableAdapters.DataTable1TableAdapter
    Dim TBL1 As DseLaporanOtorisasiTransaksi.DataTable1DataTable = DS.DataTable1
    Dim row1 As DseLaporanOtorisasiTransaksi.DataTable1Row
    Dim DTASO As New DseLaporanOtorisasiTransaksiTableAdapters.tbltrans_soTableAdapter
    Dim TBLSO As DseLaporanOtorisasiTransaksi.tbltrans_soDataTable = Me.DS.tbltrans_so
    Dim rowSO As DseLaporanOtorisasiTransaksi.tbltrans_soRow
    Dim DTADO As New DseLaporanOtorisasiTransaksiTableAdapters.tbltrans_doTableAdapter
    Dim TBLDO As DseLaporanOtorisasiTransaksi.tbltrans_doDataTable = DS.tbltrans_do
    Dim rowDO As DseLaporanOtorisasiTransaksi.tbltrans_doRow


    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Private Sub IsikanChecklistbox()
        Try
            Me.clbTipeOtorisasi.DataSource = Nothing
            DTA.Fill(Me.TBL)

            clbTipeOtorisasi.DataSource = Me.TBL.DefaultView
            clbTipeOtorisasi.DisplayMember = "nama"

            For i As Integer = 0 To Me.clbTipeOtorisasi.Items.Count - 1
                Me.clbTipeOtorisasi.SetItemCheckState(i, CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub FrmLaporanOtorisasiTransaksi_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            Me.DTADO.Connection.ConnectionString = DTA.Connection.ConnectionString
            Me.DTASO.Connection.ConnectionString = DTA.Connection.ConnectionString
            Me.IsikanChecklistbox()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Dim Filter As String = ""
        Dim PrimaryMax_SO As Double = Double.MinValue
        Dim PrimaryMin_SO As Double = Double.MaxValue
        Dim PrimaryMax_DO As Double = Double.MinValue
        Dim PrimaryMin_DO As Double = Double.MaxValue
        Dim TipeOtorisasi As ClsMasterTipeOtorisasi.EnumTipeOtorisasi = ClsMasterTipeOtorisasi.EnumTipeOtorisasi.DO_Exceed_Credit_Limit
        Dim ListPrimaryChecked As New List(Of Double)
        Dim TBLDummy As New DseLaporanOtorisasiTransaksi.DataTable1DataTable
        Dim rowDummy As DseLaporanOtorisasiTransaksi.DataTable1Row
        Dim Rpt As New CrpLaporanOtorisasiTransaksi

        Try
            Me.TBL1.Rows.Clear()
            Me.DTA1.Fill(Me.TBL1, Me.dtpTo.Value, Me.dtpFrom.Value)
            Me.TxtRequestBy.Text = Me.TxtRequestBy.Text.Trim
            Me.TxtDocumentNo.Text = Me.TxtDocumentNo.Text.Trim
            Me.TxtApprovedBy.Text = Me.TxtApprovedBy.Text.Trim

            'filter berdasarkan tipe transaksi
            If Me.clbTipeOtorisasi.Items.Count <> Me.clbTipeOtorisasi.CheckedItems.Count Then
                For Each DV As DataRowView In Me.clbTipeOtorisasi.CheckedItems
                    Me.row = DV.Row
                    ListPrimaryChecked.Add(Me.row.primary)
                Next
                For ij As Double = Me.TBL1.Rows.Count - 1 To 0 Step -1
                    Me.row1 = Me.TBL1.Rows(ij)
                    If ListPrimaryChecked.Contains(Me.row1.primary_approval_tipe) = False Then
                        Me.TBL1.Rows.RemoveAt(ij)
                    End If
                Next
            End If
            'isikan no document
            For Each row1 In TBL1.Rows
                TipeOtorisasi = row1.primary_approval_tipe
                If TipeOtorisasi = ClsMasterTipeOtorisasi.EnumTipeOtorisasi.DO_Exceed_Credit_Limit Or _
                TipeOtorisasi = ClsMasterTipeOtorisasi.EnumTipeOtorisasi.DO_Overdue_Invoice Or _
                TipeOtorisasi = ClsMasterTipeOtorisasi.EnumTipeOtorisasi.DO_New Or _
                TipeOtorisasi = ClsMasterTipeOtorisasi.EnumTipeOtorisasi.SF_Override_DO Then
                    If row1.primary_dokumen_acuan > PrimaryMax_DO Then
                        PrimaryMax_DO = row1.primary_dokumen_acuan
                    End If
                    If row1.primary_dokumen_acuan < PrimaryMin_DO Then
                        PrimaryMin_DO = row1.primary_dokumen_acuan
                    End If
                Else
                    If row1.primary_dokumen_acuan > PrimaryMax_SO Then
                        PrimaryMax_SO = row1.primary_dokumen_acuan
                    End If
                    If row1.primary_dokumen_acuan < PrimaryMin_SO Then
                        PrimaryMin_SO = row1.primary_dokumen_acuan
                    End If
                End If
                If TipeOtorisasi = ClsMasterTipeOtorisasi.EnumTipeOtorisasi.DO_New Then
                    row1.Tipe_Approval = "DO New"
                ElseIf TipeOtorisasi = ClsMasterTipeOtorisasi.EnumTipeOtorisasi.DO_Exceed_Credit_Limit Then
                    row1.Tipe_Approval = "DO Exceed Credit Limit"
                ElseIf TipeOtorisasi = ClsMasterTipeOtorisasi.EnumTipeOtorisasi.DO_Overdue_Invoice Then
                    row1.Tipe_Approval = "DO Overdue Invoice"
                ElseIf TipeOtorisasi = ClsMasterTipeOtorisasi.EnumTipeOtorisasi.SO_Delete Then
                    row1.Tipe_Approval = "SO Delete"
                ElseIf TipeOtorisasi = ClsMasterTipeOtorisasi.EnumTipeOtorisasi.SO_Edit Then
                    row1.Tipe_Approval = "SO Edit"
                ElseIf TipeOtorisasi = ClsMasterTipeOtorisasi.EnumTipeOtorisasi.SO_Exceed_Credit_Limit Then
                    row1.Tipe_Approval = "SO Exceed Credit Limit"
                ElseIf TipeOtorisasi = ClsMasterTipeOtorisasi.EnumTipeOtorisasi.SO_Overdue_Invoice Then
                    row1.Tipe_Approval = "SO Overdue Invoice"
                ElseIf TipeOtorisasi = ClsMasterTipeOtorisasi.EnumTipeOtorisasi.SF_Override_DO Then
                    row1.Tipe_Approval = "Shipping Forecast Override DO"
                Else
                    row1.Tipe_Approval = ""
                End If
            Next
            If PrimaryMax_SO <> Double.MinValue Then
                Me.DTASO.Fill(TBLSO, PrimaryMax_SO, PrimaryMin_SO)
            End If
            If PrimaryMax_DO <> Double.MinValue Then
                Me.DTADO.Fill(TBLDO, PrimaryMax_DO, PrimaryMin_DO)
            End If
            For Each row1 In TBL1.Rows
                TipeOtorisasi = row1.primary_approval_tipe
                If TipeOtorisasi = ClsMasterTipeOtorisasi.EnumTipeOtorisasi.DO_Exceed_Credit_Limit Or _
                TipeOtorisasi = ClsMasterTipeOtorisasi.EnumTipeOtorisasi.DO_Overdue_Invoice Or _
                TipeOtorisasi = ClsMasterTipeOtorisasi.EnumTipeOtorisasi.DO_New Or _
                TipeOtorisasi = ClsMasterTipeOtorisasi.EnumTipeOtorisasi.SF_Override_DO Then
                    Me.TBLDO.DefaultView.RowFilter = "primary = " + row1.primary_dokumen_acuan.ToString
                    If Me.TBLDO.DefaultView.Count > 0 Then
                        Me.rowDO = Me.TBLDO.DefaultView(0).Row
                        row1.Document_No = Me.rowDO.No_Transaksi
                    Else
                        row1.Document_No = row1.document_no_his
                    End If
                Else
                    Me.TBLSO.DefaultView.RowFilter = "primary = " + row1.primary_dokumen_acuan.ToString
                    If Me.TBLSO.DefaultView.Count > 0 Then
                        Me.rowSO = Me.TBLSO.DefaultView(0).Row
                        row1.Document_No = Me.rowSO.No_Transaksi
                    Else
                        row1.Document_No = row1.document_no_his
                    End If
                End If
            Next

            If Me.TxtRequestBy.Text <> "" Then
                Filter = "[Nama User Approved By] LIKE '" + Me.TxtRequestBy.Text + "'"
            End If
            If Me.TxtDocumentNo.Text <> "" Then
                If Filter = "" Then
                    Filter = "[Document No] LIKE '" + Me.TxtDocumentNo.Text + "'"
                Else
                    Filter = Filter + " AND [Document No] LIKE " + Me.TxtDocumentNo.Text + "'"
                End If
            End If
            If Me.TxtApprovedBy.Text <> "" Then
                If Filter = "" Then
                    Filter = "[Nama User Approved By] LIKE '" + Me.TxtApprovedBy.Text + "'"
                Else
                    Filter = Filter + " AND [Nama User Approved By] LIKE '" + Me.TxtApprovedBy.Text + "'"
                End If
            End If
            Me.TBL1.DefaultView.RowFilter = Filter
            TBLDummy.Rows.Clear()
            For Each dv As DataRowView In Me.TBL1.DefaultView
                Me.row1 = dv.Row
                rowDummy = TBLDummy.NewDataTable1Row
                For ijk As Integer = 0 To TBL1.Columns.Count - 1
                    rowDummy.Item(ijk) = Me.row1.Item(ijk)
                Next
                TBLDummy.Rows.Add(rowDummy)
            Next
            Me.TBL1.Rows.Clear()
            For Each rowDummy In TBLDummy.Rows
                Me.row1 = TBL1.NewDataTable1Row
                For ijk As Integer = 0 To TBL1.Columns.Count - 1
                    Me.row1.Item(ijk) = rowDummy.Item(ijk)
                Next
                TBL1.Rows.Add(Me.row1)
            Next

            'tampilkan pada laporan
            Rpt.SetDataSource(DS)
            Rpt.SetParameterValue(Rpt.Parameter_End_Date.ParameterFieldName, Me.dtpTo.Value)
            Rpt.SetParameterValue(Rpt.Parameter_Start_Date.ParameterFieldName, Me.dtpFrom.Value)

            Me.crv1.ReportSource = Rpt
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryChecked = Nothing
            TBLDummy.Dispose()
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub TextBox1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtApprovedBy.GotFocus, TxtDocumentNo.GotFocus, TxtRequestBy.GotFocus
        Dim Txt As TextBox

        Try
            Txt = TryCast(sender, TextBox)
            If Txt.Enabled = True Then
                If Txt.ReadOnly = False Then
                    Txt.BackColor = Me.BackColorAll
                Else
                    Txt.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtApprovedBy.LostFocus, TxtDocumentNo.LostFocus, TxtRequestBy.LostFocus
        Dim Txt As TextBox

        Try
            Txt = TryCast(sender, TextBox)
            If Txt.Enabled = True Then
                If Txt.ReadOnly = False Then
                    Txt.BackColor = Me.TextBox1.BackColor
                Else
                    Txt.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCekAll.LinkClicked
        For i As Integer = 0 To Me.clbTipeOtorisasi.Items.Count - 1
            Me.clbTipeOtorisasi.SetItemCheckState(i, CheckState.Checked)
        Next
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUncekAll.LinkClicked
        For i As Integer = 0 To Me.clbTipeOtorisasi.Items.Count - 1
            Me.clbTipeOtorisasi.SetItemCheckState(i, CheckState.Unchecked)
        Next
    End Sub
End Class