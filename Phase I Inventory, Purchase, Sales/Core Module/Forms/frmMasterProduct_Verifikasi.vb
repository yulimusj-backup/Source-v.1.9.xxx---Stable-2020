Option Explicit On
Option Strict Off

Public Class frmMasterProduct_Verifikasi
    Public UserCode As String = "USER"
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim DS As New dseMasterProduct
    Dim TBL As New dseMasterProduct.tblmaster_product4DataTable
    Dim DTA As New dseMasterProductTableAdapters.tblmaster_product4TableAdapter
    Dim row As dseMasterProduct.tblmaster_product4Row
    Dim rowDGV As DataGridViewRow
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjVerified As New ClsMasterProduct_Verifikasi

    Private Sub frmMasterProduct_Verifikasi_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            LoadDGV(Me.dgvProduct, Me, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterProduct_Verifikasi_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadDGV(Me.dgvProduct, Me, Me.ObjCD)

        Me.TabControl1.TabPages.Remove(Me.TabPage_Sales)
        Me.TabControl1.TabPages.Remove(Me.TabPage_SalesReturn)
        Me.TabControl1.TabPages.Remove(Me.TabPage_Purchase)
        Me.TabControl1.TabPages.Remove(Me.TabPage_PurchaseReturn)
        Me.TabControl1.TabPages.Remove(Me.TabPage_Cost)
        Me.TabControl1.TabPages.Remove(Me.TabPage_Adjustment)
        Me.TabControl1.TabPages.Remove(Me.TabPage_StockTaking)
        Me.BackColorAll = Core_Module.GetSMESSBackColorAll
        Me.dgvProduct.AlternatingRowsDefaultCellStyle.BackColor = BackColorAll
        Me.ShowIcon = False
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnVerify_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnVerify.Click
        Dim AdaYangDicek As Boolean = False
        Dim ListRowDgv As New List(Of Integer)
        Dim hasil As MsgBoxResult
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For Each rowDgv As DataGridViewRow In Me.dgvProduct.Rows
                With rowDgv
                    If .Cells(Me.ColCheck.Index).Value = Boolean.TrueString Then
                        AdaYangDicek = True
                    End If
                End With
            Next
            If AdaYangDicek = False Then
                MsgBox("Please check at least 1 item.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            hasil = MsgBox("Verify product account?", MsgBoxStyle.YesNo, "Confirm")
            If hasil <> MsgBoxResult.Yes Then
                Exit Sub
            End If

            For Each rowDgv As DataGridViewRow In Me.dgvProduct.Rows
                With rowDgv
                    If .Cells(Me.ColCheck.Index).Value = Boolean.TrueString Then
                        If hasil = MsgBoxResult.Yes Then
                            Me.ObjVerified.SimpanPerubahanVerify(ClsMasterProduct_Verifikasi.EnumVerified.Yes, _
                            Me.UserCode, .Cells(Me.ColPrimary.Index).Value)
                            ListRowDgv.Add(Me.dgvProduct.Rows.IndexOf(rowDgv))
                        End If
                    End If
                End With
            Next

            Me.dgvProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvProduct.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            If ListRowDgv.Count > 0 Then
                Dim ab As Integer
                For i As Integer = ListRowDgv.Count - 1 To 0 Step -1
                    ab = ListRowDgv.Item(i)
                    Me.dgvProduct.Rows.RemoveAt(ab)
                Next
            End If
            Me.dgvProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Me.ObjCD.SetNumber(Me.dgvProduct)
            MsgBox("Data Saved", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub LoadDGV(ByVal vDGV As DataGridView, ByVal i As frmMasterProduct_Verifikasi, ByVal vObjCD As ClsCekData)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vDGV.Rows.Clear()
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'load database
            DTA.Fill(TBL)

            'pindahin ke DGV
            For Each row In TBL.Rows
                vDGV.Rows.Add()
                rowDGV = vDGV.Rows(vDGV.Rows.Count - 1)
                With rowDGV
                    .Cells(i.ColNo.Index).Value = ""
                    .Cells(i.ColPrimary.Index).Value = row.Primary
                    .Cells(i.ColCodeNo.Index).Value = row.Product_Code
                    .Cells(i.ColDesc.Index).Value = row.Description
                    If row.Type = 0 Then
                        .Cells(i.ColTipe.Index).Value = "Goods"
                    Else
                        .Cells(i.ColTipe.Index).Value = "Service"
                    End If
                    .Cells(i.ColInventoryAccCode.Index).Value = row.Inventory_Account_Code
                    .Cells(i.ColInventoryAccDesc.Index).Value = row.Inventory_Account_Desc
                    .Cells(i.ColSalesAccCode.Index).Value = row.Sales_Account_Code
                    .Cells(i.ColSalesAccDesc.Index).Value = row.Sales_Account_Desc
                    .Cells(i.ColSalesReturnCode.Index).Value = row.Sales_Return_Account_Code
                    .Cells(i.ColSalesReturnDesc.Index).Value = row.Sales_Return_Account_Desc
                    .Cells(i.ColCOGSAccCode.Index).Value = row.Cost_Account_Code
                    .Cells(i.ColCOGSAccDesc.Index).Value = row.Cost_Account_Desc
                    .Cells(i.ColAdjusmentAccCode.Index).Value = row.Adjustment_Account_Code
                    .Cells(i.ColAdjusmentAccDesc.Index).Value = row.Adjustment_Account_Desc
                    .Cells(i.ColStockTakingAccCode.Index).Value = row.Stock_Taking_Account_Code
                    .Cells(i.ColStokTakingAccDesc.Index).Value = row.Stock_Taking_Account_Desc
                    .Cells(i.ColCategory.Index).Value = row.Category_Desc
                End With
            Next

            vObjCD.SetNumber(vDGV)
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            vDGV.ClearSelection()
            If vDGV.RowCount > 0 Then
                vDGV.Rows(0).Selected = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DS.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub dgvProduct_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgvProduct.SelectionChanged
        Try
            If Me.dgvProduct.SelectedRows.Count > 0 Then
                Me.TxtInventoryAcc.Text = Me.dgvProduct.SelectedRows(0).Cells(Me.ColInventoryAccCode.Index).Value
                Me.TxtInventoryAccountDesc.Text = Me.dgvProduct.SelectedRows(0).Cells(Me.ColInventoryAccDesc.Index).Value
                Me.TxtSalesAccount.Text = Me.dgvProduct.SelectedRows(0).Cells(Me.ColSalesAccCode.Index).Value
                Me.TxtSalesAccountDesc.Text = Me.dgvProduct.SelectedRows(0).Cells(Me.ColSalesAccDesc.Index).Value
                Me.TxtSalesReturnAccount.Text = Me.dgvProduct.SelectedRows(0).Cells(Me.ColSalesReturnCode.Index).Value
                Me.TxtSalesReturnAccountDesc.Text = Me.dgvProduct.SelectedRows(0).Cells(Me.ColSalesReturnDesc.Index).Value
                Me.TxtCOGSAccount.Text = Me.dgvProduct.SelectedRows(0).Cells(Me.ColCOGSAccCode.Index).Value
                Me.TxtCOGSAccountDesc.Text = Me.dgvProduct.SelectedRows(0).Cells(Me.ColCOGSAccDesc.Index).Value
                Me.TxtAdjustmentAccount.Text = Me.dgvProduct.SelectedRows(0).Cells(Me.ColAdjusmentAccCode.Index).Value
                Me.TxtAdjustmentAccountDesc.Text = Me.dgvProduct.SelectedRows(0).Cells(Me.ColAdjusmentAccDesc.Index).Value
                Me.TxtStockTakingAccount.Text = Me.dgvProduct.SelectedRows(0).Cells(Me.ColStockTakingAccCode.Index).Value
                Me.TxtStockTakingAccountDesc.Text = Me.dgvProduct.SelectedRows(0).Cells(Me.ColStokTakingAccDesc.Index).Value
            Else
                Me.KosongkanAkunDetail()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProduct_CellMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvProduct.CellMouseClick
        Dim hasil As MsgBoxResult
        Try
            With row
                If e.RowIndex >= 0 Then
                    If e.ColumnIndex = Me.ColVerify.Index Then
                        hasil = MsgBox("Verify product account?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            If Me.ObjVerified.SimpanPerubahanVerify(ClsMasterProduct_Verifikasi.EnumVerified.Yes, _
                            Me.UserCode, Me.dgvProduct.Item(Me.ColPrimary.Index, e.RowIndex).Value.ToString) = True Then
                                dgvProduct.Rows.RemoveAt(dgvProduct.CurrentRow.Index)
                                Me.ObjCD.SetNumber(Me.dgvProduct)
                                MsgBox("Data Saved", MsgBoxStyle.Information, "Information")

                                If Me.dgvProduct.RowCount = 0 Then
                                    Me.KosongkanAkunDetail()
                                End If
                            End If
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanAkunDetail()
        Try
            Me.TxtInventoryAcc.Text = ""
            Me.TxtInventoryAccountDesc.Text = ""
            Me.TxtSalesAccount.Text = ""
            Me.TxtSalesAccountDesc.Text = ""
            Me.TxtSalesReturnAccount.Text = ""
            Me.TxtSalesReturnAccountDesc.Text = ""
            Me.TxtCOGSAccount.Text = ""
            Me.TxtCOGSAccountDesc.Text = ""
            Me.TxtAdjustmentAccount.Text = ""
            Me.TxtAdjustmentAccountDesc.Text = ""
            Me.TxtStockTakingAccount.Text = ""
            Me.TxtStockTakingAccountDesc.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProduct_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvProduct.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LblCheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblCheckAll.LinkClicked
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.dgvProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rowDgv As DataGridViewRow In Me.dgvProduct.Rows
                With rowDgv
                    .Cells(Me.ColCheck.Index).Value = Boolean.TrueString
                End With
            Next
            Me.dgvProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub LblUncheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUncheckAll.LinkClicked
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.dgvProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvProduct.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            For Each rowDgv As DataGridViewRow In Me.dgvProduct.Rows
                With rowDgv
                    .Cells(Me.ColCheck.Index).Value = Boolean.FalseString
                End With
            Next
            Me.dgvProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowDetailAccount.CheckedChanged
        Try
            If Me.chkShowDetailAccount.Checked = True Then
                Me.ColAdjusmentAccCode.Visible = True
                Me.ColAdjusmentAccDesc.Visible = True
                Me.ColCOGSAccCode.Visible = True
                Me.ColCOGSAccDesc.Visible = True
                Me.ColInventoryAccCode.Visible = True
                Me.ColInventoryAccDesc.Visible = True
                Me.ColSalesAccCode.Visible = True
                Me.ColSalesAccDesc.Visible = True
                Me.ColSalesReturnCode.Visible = True
                Me.ColSalesReturnDesc.Visible = True
                Me.ColStockTakingAccCode.Visible = True
                Me.ColStokTakingAccDesc.Visible = True
                Me.dgvProduct.MultiSelect = True
                Me.dgvProduct.Height = Me.TabControl1.Location.Y + Me.TabControl1.Size.Height - Me.dgvProduct.Location.Y
            Else
                Me.ColAdjusmentAccCode.Visible = False
                Me.ColAdjusmentAccDesc.Visible = False
                Me.ColCOGSAccCode.Visible = False
                Me.ColCOGSAccDesc.Visible = False
                Me.ColInventoryAccCode.Visible = False
                Me.ColInventoryAccDesc.Visible = False
                Me.ColSalesAccCode.Visible = False
                Me.ColSalesAccDesc.Visible = False
                Me.ColSalesReturnCode.Visible = False
                Me.ColSalesReturnDesc.Visible = False
                Me.ColStockTakingAccCode.Visible = False
                Me.ColStokTakingAccDesc.Visible = False
                Me.dgvProduct.MultiSelect = False
                Me.dgvProduct.Height = Me.TabControl1.Location.Y - Me.dgvProduct.Location.Y - 10
            End If
            Me.TabControl1.SendToBack()
            Me.dgvProduct.BringToFront()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class