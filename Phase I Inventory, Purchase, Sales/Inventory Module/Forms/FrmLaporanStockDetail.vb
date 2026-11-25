Option Explicit On
Option Strict Off

Imports Core_Module
Public Class FrmLaporanStockDetail
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = ""
    Dim ObjProduct As New clsMasterProduct
    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub FrmLaporanStockDetail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'DseLaporanStockDetail.stock_detail_view' table. You can move, or remove it, as needed.
        'Me.btnRefresh.PerformClick()
    End Sub

    Private Sub txtCodeNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodeNo.GotFocus
        txtCodeNo.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtCodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCodeNo.KeyDown
        Dim i As New frmCariMasterProduct
        Dim Jml As Long = 0

        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjProduct.GetJml(Me.txtCodeNo.Text)
                If Jml = 1 Then
                    Me.TxtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.txtCodeNo.Text)
                    Me.txtPrimaryCategory.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Me.btnRefresh.PerformClick()
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.txtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.txtCodeNo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.txtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                        Me.btnRefresh.PerformClick()
                    Else
                        Me.txtCodeNo.Clear()
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.txtPrimaryCategory.Text = "-1"
                        Me.txtCodeNo.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodeNo.Leave
        Dim i As New frmCariMasterProduct

        Try
            If Me.txtCodeNo.Text <> "" Then
                If Me.ObjProduct.GetJml(Me.txtCodeNo.Text) = 1 Then
                    Me.TxtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.txtCodeNo.Text)
                    Me.txtPrimaryCategory.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Me.btnRefresh.PerformClick()
                Else
                    If Me.txtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.txtCodeNo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.txtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                        Me.btnRefresh.PerformClick()
                    Else
                        Me.txtCodeNo.Clear()
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.txtPrimaryCategory.Text = "-1"
                        Me.txtCodeNo.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.txtDesc.Clear()
                Me.txtPrimaryCategory.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCodeNo.Click
        Dim i As New frmCariMasterProduct

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.txtCodeNo.Text <> "" Then
                i.txtCari.Text = Me.txtCodeNo.Text
            End If
            i.BackColorAll = Me.BackColorAll
            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                Me.txtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                Me.btnRefresh.PerformClick()
            Else
                Me.txtCodeNo.Clear()
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.txtPrimaryCategory.Text = "-1"
                Me.txtCodeNo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        'TODO: This line of code loads data into the 'DseLaporanStockDetail.stock_detail_view' table. You can move, or remove it, as needed.

        Try
            If Me.txtCodeNo.Text <> "" Then
                Me.Stock_detail_viewTableAdapter.FillBy(Me.DseLaporanStockDetail.stock_detail_view, Me.txtCodeNo.Text)
            Else
                Me.Stock_detail_viewTableAdapter.Fill(Me.DseLaporanStockDetail.stock_detail_view)
            End If
            Dim rowds As dseLaporanStockDetail.stock_detail_viewRow
            Dim row1 As DataGridViewRow
            Dim SecNo As Boolean = True
            Dim LotNo As Boolean = False
            Dim Location As Boolean = True
            Dim ExpiryDate As Boolean = True

            For Each rowds In Me.DseLaporanStockDetail.stock_detail_view
                rowds.Quantity = Me.ObjCD.DigitGrouping(rowds.Qty, rowds._Decimal)
                If rowds.Expiry_Date = #1/1/1950# Then
                    rowds.SetExpiry_DateNull()
                End If
                If rowds.Location = "No Location" Then
                    rowds.SetLocationNull()
                End If
                If rowds.Qty_Secondary = -1 Then
                    rowds.SetQty_SecondaryNull()
                End If
            Next
            For Each row1 In Me.Stock_detail_viewDataGridView.Rows
                If IsDBNull(row1.Cells(6).Value) = False Then
                    SecNo = False
                End If
                If CStr(row1.Cells(7).Value) <> "" Then
                    LotNo = True
                End If
                If IsDBNull(row1.Cells(8).Value) = False Then
                    ExpiryDate = False
                End If
                If IsDBNull(row1.Cells(10).Value) = False Then
                    Location = False
                End If
            Next
            If SecNo = True Then
                Me.Stock_detail_viewDataGridView.Columns(6).Visible = False
            Else
                Me.Stock_detail_viewDataGridView.Columns(6).Visible = True
            End If
            If LotNo = True Then
                Me.Stock_detail_viewDataGridView.Columns(7).Visible = True
            Else
                Me.Stock_detail_viewDataGridView.Columns(7).Visible = False
            End If
            If ExpiryDate = True Then
                Me.Stock_detail_viewDataGridView.Columns(8).Visible = False
            Else
                Me.Stock_detail_viewDataGridView.Columns(8).Visible = True
            End If
            If Location = False Then
                Me.Stock_detail_viewDataGridView.Columns(10).Visible = True
            Else
                Me.Stock_detail_viewDataGridView.Columns(10).Visible = False
            End If
            Me.Stock_detail_viewDataGridView.Columns(Me.ColQty.Index).Visible = False
            Me.Stock_detail_viewDataGridView.Columns(Me.Quantity.Index).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.Stock_detail_viewDataGridView.ScrollBars = ScrollBars.Both
            Me.Stock_detail_viewDataGridView.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodeNo.LostFocus
        txtCodeNo.BackColor = Me.txtPrimaryCategory.BackColor
    End Sub
End Class