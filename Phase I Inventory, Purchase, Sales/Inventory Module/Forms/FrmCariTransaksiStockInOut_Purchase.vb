Option Explicit On
Option Strict Off

Public Class FrmCariTransaksiStockInOut_Purchase
    Dim ObjCD As New Core_Module.ClsCekData

    Public PrimaryTransaksiStockInOut As Double = -1
    Public JenisTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut
    Public InOutAdjust As clsTransaksiStockInOut.EnumInOutAdjust = clsTransaksiStockInOut.EnumInOutAdjust.NoTipe
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim Obj As New clsTransaksiStockInOut
    Dim FindCriteria As ModGlobalVariable.EnumSearchCriteria = EnumSearchCriteria.ByKodeCustomer
    Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout6TableAdapter
    Dim TBL As New dseTransaksiStockInOut.tbltransaksi_stockinout6DataTable

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

            Select Case Me.JenisTransaksi
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
                    Exit Sub
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess
                    Exit Sub
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process
                    Exit Sub
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                    Me.Text = "Find Stock In - Purchase"
                    Me.cboCriteria.Items.Add("Transaction No")
                    Me.cboCriteria.Items.Add("Date")
                    Me.cboCriteria.Items.Add("Supplier Code")
                    Me.cboCriteria.Items.Add("Supplier Name")
                    Me.cboCriteria.Items.Add("Product Code No")
                    Me.cboCriteria.Items.Add("Product Description")
                    Me.cboCriteria.Items.Add("Purchase Order No")
                    Me.cboCriteria.Items.Add("Supplier DO Number")
                    Me.cboCriteria.Items.Add("Remarks")

                    Exit Select
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
                    Exit Sub
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
                    Exit Sub
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Process
                    Exit Sub
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn
                    Me.Text = "Find Stock Out - Purchase Return"
                    Me.cboCriteria.Items.Add("Transaction No")
                    Me.cboCriteria.Items.Add("Date")
                    Me.cboCriteria.Items.Add("Supplier Code")
                    Me.cboCriteria.Items.Add("Supplier Name")
                    Me.cboCriteria.Items.Add("Product Code No")
                    Me.cboCriteria.Items.Add("Product Description")
                    Me.cboCriteria.Items.Add("Remarks")
                    Me.cboCriteria.Items.Add("Purchase Return No")
                    Exit Select
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales
                    Exit Sub
            End Select

            Me.cboCriteria.SelectedIndex = 0
            Me.dgvHasil.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            If Me.txtCari.Text <> "" Then
                Me.btnFind.PerformClick()
            End If

            'Me.dgvHasil.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCari_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCari.KeyDown
        Try
            If e.KeyValue = 13 Then
                Me.IsiDatabase()
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
            Me.IsiDatabase()
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
            If Me.Obj.GetJml(Me.JenisTransaksi, Me.txtCari.Text, Me.FindCriteria, _
            Me.dtpDari.Value.Date, Me.dtpSampai.Value.Date) < 150 Then
                Me.IsiDatabase()
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
            If Me.cboCriteria.Text = "Transaction No" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.FindCriteria = EnumSearchCriteria.ByNoTransaksi
            ElseIf Me.cboCriteria.Text = "Date" Then
                Me.LblFrom.Visible = True
                Me.lblTo.Visible = True
                Me.dtpDari.Visible = True
                Me.dtpSampai.Visible = True
                Me.txtCari.Visible = False
                Me.FindCriteria = EnumSearchCriteria.ByTanggalTransaksi
            ElseIf Me.cboCriteria.Text = "Supplier Code" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.FindCriteria = EnumSearchCriteria.ByKodeSupplier
            ElseIf Me.cboCriteria.Text = "Supplier Name" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.FindCriteria = EnumSearchCriteria.ByNamaSupplier
            ElseIf Me.cboCriteria.Text = "Customer Code" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.FindCriteria = EnumSearchCriteria.ByKodeCustomer
            ElseIf Me.cboCriteria.Text = "Customer Name" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.FindCriteria = EnumSearchCriteria.ByNamaCustomer
            ElseIf Me.cboCriteria.Text = "Product Code No" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.FindCriteria = EnumSearchCriteria.ByProductCodeNo
            ElseIf Me.cboCriteria.Text = "Product Description" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.FindCriteria = EnumSearchCriteria.ByProductDesc
            ElseIf Me.cboCriteria.Text = "Purchase Order No" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.FindCriteria = EnumSearchCriteria.ByPONumber
            ElseIf Me.cboCriteria.Text = "Supplier DO Number" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.FindCriteria = EnumSearchCriteria.BySupplier_s_DONumber
            ElseIf Me.cboCriteria.Text = "Remarks" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.FindCriteria = EnumSearchCriteria.ByRemarks
            ElseIf Me.cboCriteria.Text = "Purchase Return No" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.FindCriteria = EnumSearchCriteria.ByPurchaseReturnNo
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

    Private Sub IsiDatabase()
        Dim QueryResult As Double = 0
        Dim StartT As Date = Now

        Try
            Me.dgvHasil.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If Me.FindCriteria = EnumSearchCriteria.ByNoTransaksi Then
                Me.DTA.Fill(Me.TBL, Me.txtCari.Text + "%", Me.JenisTransaksi, Me.InOutAdjust)
                QueryResult = Me.TBL.Rows.Count
            ElseIf Me.FindCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                Me.DTA.FillBy_TglSampaiDari(Me.TBL, Me.dtpSampai.Value, Me.dtpDari.Value, Me.JenisTransaksi, Me.InOutAdjust)
                QueryResult = Me.TBL.Rows.Count
            ElseIf Me.FindCriteria = EnumSearchCriteria.ByKodeSupplier Then
                Me.DTA.FillBy_SupplierCodeLIKE(Me.TBL, Me.txtCari.Text + "%", Me.JenisTransaksi, Me.InOutAdjust)
                QueryResult = Me.TBL.Rows.Count
            ElseIf Me.FindCriteria = EnumSearchCriteria.ByNamaSupplier Then
                Me.DTA.FillBy_SupplierNameLIKE(Me.TBL, Me.JenisTransaksi, Me.InOutAdjust, Me.txtCari.Text + "%")
                QueryResult = Me.TBL.Rows.Count
            ElseIf Me.FindCriteria = EnumSearchCriteria.ByProductCodeNo Then
                Me.DTA.FillBy_ProductCodeLIKE(Me.TBL, Me.JenisTransaksi, Me.txtCari.Text + "%", Me.InOutAdjust)
                QueryResult = Me.TBL.Rows.Count
            ElseIf Me.FindCriteria = EnumSearchCriteria.ByProductDesc Then
                Me.DTA.FillBy_ProductDescLIKE(Me.TBL, Me.JenisTransaksi, Me.txtCari.Text + "%", Me.InOutAdjust)
                QueryResult = Me.TBL.Rows.Count
            ElseIf Me.FindCriteria = EnumSearchCriteria.ByPONumber Then
                Me.DTA.FillBy_PONumberLIKE(Me.TBL, Me.txtCari.Text + "%", Me.InOutAdjust)
                QueryResult = Me.TBL.Rows.Count
            ElseIf Me.FindCriteria = EnumSearchCriteria.BySupplier_s_DONumber Then
                Me.DTA.FillBy_SupplierDONumber(Me.TBL, Me.txtCari.Text + "%", Me.InOutAdjust)
                QueryResult = Me.TBL.Rows.Count
            ElseIf Me.FindCriteria = EnumSearchCriteria.ByRemarks Then
                Me.DTA.FillBy_Remarks(Me.TBL, Me.JenisTransaksi, Me.InOutAdjust, Me.txtCari.Text + "%")
                QueryResult = Me.TBL.Rows.Count
            ElseIf Me.FindCriteria = EnumSearchCriteria.ByPurchaseReturnNo Then
                Me.DTA.FillBy_PurchaseReturnNoLIKE(Me.TBL, Me.JenisTransaksi, Me.InOutAdjust, Me.txtCari.Text + "%")
                QueryResult = Me.TBL.Rows.Count
            End If

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

            Me.dgvHasil.DataSource = Me.TBL.DefaultView
            Me.dgvHasil.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"

            If Me.JenisTransaksi <> clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Then
                Me.dgvHasil.Columns(5).Visible = False
            Else
                Me.dgvHasil.Columns(5).Visible = True
            End If
            Me.dgvHasil.Columns(0).Visible = False
            If Me.JenisTransaksi <> clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn Then
                Me.dgvHasil.Columns(6).Visible = False
            Else
                Me.dgvHasil.Columns(6).Visible = True
                Me.dgvHasil.Columns(6).HeaderText = "Purchase Return No"
            End If

            Me.dgvHasil.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class