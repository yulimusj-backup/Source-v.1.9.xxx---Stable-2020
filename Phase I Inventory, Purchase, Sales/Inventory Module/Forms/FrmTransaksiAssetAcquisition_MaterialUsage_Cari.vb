Option Explicit On
Option Strict Off

Public Class FrmTransaksiAssetAcquisition_MaterialUsage_cari
    Dim ObjCD As New Core_Module.ClsCekData

    Public PrimaryTransaksiStockInOut As Double = -1
    Public JenisTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
    Public InOutAdjust As clsTransaksiStockInOut.EnumInOutAdjust = clsTransaksiStockInOut.EnumInOutAdjust.NoTipe
    Public BackColorAll As Color

    Dim Obj As New clsTransaksiStockInOut
    Dim DTA As New dseTransaksiStockInOut_AssetTableAdapters.DataTable12TableAdapter
    Dim TBL As New dseTransaksiStockInOut_Asset.DataTable12DataTable
    Dim Row As dseTransaksiStockInOut_Asset.DataTable12Row

    Private Sub Pilih()
        Try
            If Me.dgvHasil.SelectedRows.Count > 0 Then
                Me.PrimaryTransaksiStockInOut = Me.dgvHasil.SelectedRows(0).Cells(0).Value
            Else
                Me.PrimaryTransaksiStockInOut = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmCariMasterCustomer_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.txtCari.Focus()
            Me.cboCriteria.SelectedIndex = 0
            Me.dgvHasil.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            If Me.txtCari.Text <> "" Then
                Me.btnFind.PerformClick()
            End If

            Me.dgvHasil.Columns(Me.ColDate.Index).DefaultCellStyle.Format = "dd MMM yyyy"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCari_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCari.KeyDown
        Try
            If e.KeyValue = 13 Then
                Me.IsiDgv()
                If Me.dgvHasil.Rows.Count > 0 Then
                    Me.dgvHasil.ClearSelection()
                    Me.dgvHasil.Rows(0).Selected = True
                End If
            ElseIf e.KeyValue = 40 And Me.dgvHasil.RowCount > 0 Then
                Me.dgvHasil.Focus()
                If Me.dgvHasil.Rows.Count > 0 Then
                    Me.dgvHasil.Rows.Item(0).Selected = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCari_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Try
            Me.IsiDgv()
            If Me.dgvHasil.Rows.Count > 0 Then
                Me.dgvHasil.ClearSelection()
                Me.dgvHasil.Rows(0).Selected = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHasil_CellContentDoubleClick(ByVal sender As Object, ByVal e As _
    System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvHasil.CellContentDoubleClick
        Try
            Me.Pilih()
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHasil_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvHasil.KeyDown
        Try
            If e.KeyValue = 13 Then
                Me.Pilih()
                Me.Close()
            ElseIf e.KeyValue = 38 Then
                If Me.dgvHasil.SelectedRows.Count > 0 And Me.dgvHasil.SelectedRows(0).Index = 0 Then
                    Me.txtCari.Focus()
                End If
            ElseIf e.KeyValue = 9 Then
                Me.txtCari.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOk_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            Me.Pilih()
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnKeluar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.PrimaryTransaksiStockInOut = -1
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCari_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCari.TextChanged
        Try
            If Me.Obj.GetJml_Asset(Me.txtCari.Text, Me.cboCriteria.SelectedIndex, _
            Me.dtpDari.Value.Date, Me.dtpSampai.Value.Date, Me.InOutAdjust) < 150 Then
                Me.IsiDgv()
                If Me.dgvHasil.Rows.Count > 0 Then
                    Me.dgvHasil.ClearSelection()
                    Me.dgvHasil.Rows(0).Selected = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCari_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCari.GotFocus
        txtCari.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtCari_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCari.LostFocus
        txtCari.BackColor = Me.txtprimarygroup.BackColor
    End Sub

    Private Sub cboCriteria_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCriteria.SelectedIndexChanged
        Try
            Me.dgvHasil.Refresh()
            If Me.cboCriteria.SelectedIndex = EnumSreachCriteria_Material_Usage._Date Then
                Me.LblFrom.Visible = True
                Me.lblTo.Visible = True
                Me.dtpDari.Visible = True
                Me.dtpSampai.Visible = True
                Me.txtCari.Visible = False
            Else
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpSampai_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpSampai.KeyDown
        If e.KeyValue = 13 Then
            Me.btnFind.PerformClick()
        End If
    End Sub

    Private Sub dtpDari_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpDari.KeyDown
        If e.KeyValue = 13 Then
            Me.btnFind.PerformClick()
        End If
    End Sub

    Private Sub dgvHasil_MouseDoubleClick(ByVal sender As Object, ByVal e _
    As System.Windows.Forms.MouseEventArgs) Handles dgvHasil.MouseDoubleClick
        Try
            If Me.dgvHasil.SelectedRows.Count > 0 Then
                Me.Pilih()
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHasil_RowHeaderMouseDoubleClick(ByVal sender As System.Object, _
    ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvHasil.RowHeaderMouseDoubleClick
        Try
            If Me.dgvHasil.SelectedRows.Count > 0 Then
                Me.Pilih()
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiDgv()
        Dim QueryResult As Double = 0
        Dim StartT As Date = Now
        Dim RowDGV As DataGridViewRow

        Try
            Me.dgvHasil.Rows.Clear()
            Me.dgvHasil.Refresh()
            Me.dgvHasil.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If Me.cboCriteria.SelectedIndex = EnumSreachCriteria_Material_Usage.NoTransaksi Then
                QueryResult = Me.DTA.FillBy_NoTransaksi(Me.TBL, Me.txtCari.Text + "%", Me.JenisTransaksi, Me.InOutAdjust)
            ElseIf Me.cboCriteria.SelectedIndex = EnumSreachCriteria_Material_Usage._Date Then
                QueryResult = Me.DTA.FillBy_Date(Me.TBL, Me.JenisTransaksi, Me.InOutAdjust, Me.dtpDari.Value, Me.dtpSampai.Value)
            ElseIf Me.cboCriteria.SelectedIndex = EnumSreachCriteria_Material_Usage.AssetCode Then
                QueryResult = Me.DTA.FillBy_AssetCode(Me.TBL, Me.JenisTransaksi, Me.InOutAdjust, Me.txtCari.Text + "%")
            ElseIf Me.cboCriteria.SelectedIndex = EnumSreachCriteria_Material_Usage.AssetDescrition Then
                QueryResult = Me.DTA.FillBy_AssetDesc(Me.TBL, Me.JenisTransaksi, Me.InOutAdjust, Me.txtCari.Text + "%")
            ElseIf Me.cboCriteria.SelectedIndex = EnumSreachCriteria_Material_Usage.Remarks Then
                QueryResult = Me.DTA.FillBy_Remark(Me.TBL, Me.JenisTransaksi, Me.InOutAdjust, Me.txtCari.Text + "%")
            End If

            For Each Me.Row In Me.TBL.Rows
                Me.dgvHasil.Rows.Add()
                RowDGV = Me.dgvHasil.Rows(Me.dgvHasil.RowCount - 1)
                RowDGV.Cells(Me.ColNoTransaksi.Index).Value = Me.Row.No_Transaksi
                RowDGV.Cells(Me.ColPrimary.Index).Value = Me.Row.Primary
                RowDGV.Cells(Me.ColDate.Index).Value = Me.Row.Tanggal_Transaksi
                RowDGV.Cells(Me.ColAssetCode.Index).Value = Me.Row.Asset_Code
                RowDGV.Cells(Me.ColAssetDescription.Index).Value = Me.Row.Description
                RowDGV.Cells(Me.ColRemarks.Index).Value = Me.Row.Remarks
            Next

            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.TotalSeconds > 1 Then
                    Me.tslQuery.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalSeconds, 2) + " seconds"
                Else
                    Me.tslQuery.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalMilliseconds, 2, MidpointRounding.AwayFromZero) + " miliseconds"
                End If
            Else
                Me.tslQuery.Text = "No result found"
            End If

            Me.dgvHasil.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class