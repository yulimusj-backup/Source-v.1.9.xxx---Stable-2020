Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiStockCard
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim Obj As clsMasterProduct
    Dim ObjStockcard As ClsTransaksiStockCard

    Dim TBL As dseTransaksiStockCard.tbltransaksi_stock_card2DataTable
    Dim buffer As String = ""

    Private Sub txtProductCode_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductCode.Enter
        Me.buffer = Me.txtProductCode.Text
    End Sub

    Private Sub txtProductCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductCode.GotFocus
        Me.txtProductCode.BackColor = Me.BackColorAll
    End Sub

    Private Sub txtProductCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtProductCode.KeyDown
        Dim i As New FrmCariMasterProduct
        Dim Jml As Long = 0

        Try
            If e.KeyValue = 13 Then
                Jml = Obj.GetJml(Me.txtProductCode.Text)
                If Jml = 1 Then
                    Me.txtPrimaryProduct.Text = Obj.GetPrimary(Me.txtProductCode.Text)
                    Me.txtPrimaryCategory.Text = Obj.GetPrimaryCategory(Me.txtPrimaryProduct.Text)
                    Me.txtDescription.Text = Obj.GetDescription(Me.txtPrimaryProduct.Text)
                    Me.dtpStartDate.Focus()
                    If Me.txtProductCode.Text <> Me.buffer Then
                        Me.dgvStockCard.DataSource = Nothing
                        Me.dgvStockCard.Refresh()
                    End If
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.txtProductCode.Text <> "" Then
                        i.txtCari.Text = Me.txtProductCode.Text
                    End If

                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.txtPrimaryProduct.Text = i.PrimaryProduct
                        Me.txtProductCode.Text = Obj.GetProductCode(Me.txtPrimaryProduct.Text)
                        Me.dtpStartDate.Focus()
                        If Me.txtProductCode.Text <> Me.buffer Then
                            Me.dgvStockCard.DataSource = Nothing
                            Me.dgvStockCard.Refresh()
                        End If
                    Else
                        Me.txtProductCode.Clear()
                        Me.txtPrimaryProduct.Text = "-1"
                        Me.txtPrimaryCategory.Text = "-1"
                        Me.txtDescription.Clear()
                        Me.txtProductCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtProductCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductCode.Leave
        Dim i As New FrmCariMasterProduct

        Try
            If Me.txtProductCode.Text <> "" Then
                If obj.GetJml(Me.txtProductCode.Text) = 1 Then
                    Me.txtPrimaryProduct.Text = obj.GetPrimary(Me.txtProductCode.Text)
                    Me.txtPrimaryCategory.Text = obj.GetPrimaryCategory(Me.txtPrimaryProduct.Text)
                    Me.txtDescription.Text = obj.GetDescription(Me.txtPrimaryProduct.Text)
                    Me.dtpStartDate.Focus()
                    If Me.txtProductCode.Text <> Me.buffer Then
                        Me.dgvStockCard.DataSource = Nothing
                        Me.dgvStockCard.Refresh()
                    End If
                Else
                    If Me.txtProductCode.Text <> "" Then
                        i.txtCari.Text = Me.txtProductCode.Text
                    End If

                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.txtPrimaryProduct.Text = i.PrimaryProduct
                        Me.txtProductCode.Text = obj.GetProductCode(Me.txtPrimaryProduct.Text)
                        Me.dtpStartDate.Focus()
                        If Me.txtProductCode.Text <> Me.buffer Then
                            Me.dgvStockCard.DataSource = Nothing
                            Me.dgvStockCard.Refresh()
                        End If
                    Else
                        Me.txtProductCode.Clear()
                        Me.txtPrimaryProduct.Text = "-1"
                        Me.txtPrimaryCategory.Text = "-1"
                        Me.txtDescription.Clear()
                        Me.txtProductCode.Focus()
                        Me.dgvStockCard.DataSource = Nothing
                    End If
                End If
            Else
                Me.txtPrimaryProduct.Text = "-1"
                Me.txtDescription.Clear()
                Me.txtPrimaryCategory.Text = "-1"
                Me.dgvStockCard.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindProduct.Click
        Dim i As New frmCariMasterProduct

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.txtProductCode.Text <> "" Then
                i.txtCari.Text = Me.txtProductCode.Text
            End If

            i.BackColorAll = Me.BackColorAll
            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.txtPrimaryProduct.Text = i.PrimaryProduct
                Me.txtProductCode.Text = obj.GetProductCode(Me.txtPrimaryProduct.Text)
                Me.dtpStartDate.Focus()
            Else
                Me.txtProductCode.Clear()
                Me.txtPrimaryProduct.Text = "-1"
                Me.txtPrimaryCategory.Text = "-1"
                Me.txtDescription.Clear()
                Me.txtProductCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub FrmTransaksiStockCard_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.Obj = New clsMasterProduct
            Me.ObjStockcard = New ClsTransaksiStockCard
            Me.dtpStartDate.Value = Now - TimeSpan.FromDays(30)
            Me.TBL = New dseTransaksiStockCard.tbltransaksi_stock_card2DataTable
            Me.dgvStockCard.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim i As New FrmCetakStockCard
        Try
            If Me.txtPrimaryProduct.Text <> -1 Then
                If Me.dtpStartDate.Value.Date <= Me.dtpEndDate.Value.Date Then
                    i.PrimaryProduct = Me.txtPrimaryProduct.Text
                    i.ProductCode = Me.txtProductCode.Text
                    i.StartDate = Me.dtpStartDate.Value
                    i.EndDate = Me.dtpEndDate.Value
                    i.Units = Me.Obj.GetUnitCode(i.PrimaryProduct)
                    i.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Try
            If Me.txtPrimaryProduct.Text <> -1 Then
                If Me.dtpStartDate.Value.Date <= Me.dtpEndDate.Value.Date Then
                    Me.ObjStockcard.PrintDGV(Me.txtPrimaryProduct.Text, Me.dtpStartDate.Value, Me.dtpEndDate.Value, _
                    Me.dgvStockCard, Me.TBL)

                    Me.dgvStockCard.ReadOnly = False

                    Dim row As DataGridViewRow
                    For Each row In Me.dgvStockCard.Rows
                        row.Cells(0).Value = CDate(row.Cells(0).Value).ToShortDateString
                    Next

                    Me.dgvStockCard.ReadOnly = True
                Else
                    MsgBox("Wrong date filter, baginning date must be smaller than end date", _
                    MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
        End Try
    End Sub

    Private Sub txtProductCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductCode.LostFocus
        Me.txtProductCode.BackColor = Me.txtPrimaryCategory.BackColor
    End Sub
End Class