Option Explicit On
Option Strict Off

Partial Public Class frmMasterProduct
    Dim ObjProses As New Core_Module.ClsMasterProses
    Dim IndexRemove As Integer = 0

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Dim PrimaryProduct As Double = -1

        Try
            PrimaryProduct = Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, Me.IndexRemove).Value
            Me.dgvItemBOM.Rows.RemoveAt(Me.IndexRemove)
            If PrimaryProduct <> -1 Then
                Me.Update_TreeView_BuangNodeLvl2(PrimaryProduct)
            End If
            Me.IndexRemove = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtProcessCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProcessCode.GotFocus
        Try
            If Me.txtCategoryCode.Enabled = True Then
                If Me.txtProcessCode.ReadOnly = True Then
                    Me.txtProcessCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtProcessCode.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtProcessCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProcessCode.LostFocus
        Try
            If Me.txtCategoryCode.Enabled = True Then
                If Me.txtProcessCode.ReadOnly = True Then
                    Me.txtProcessCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtProcessCode.BackColor = Me.txtPrimaryCategory.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyRef_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyRef.GotFocus
        Try
            If Me.txtCategoryCode.Enabled = True Then
                If Me.TxtQtyRef.ReadOnly = True Then
                    Me.TxtQtyRef.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtQtyRef.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyRef_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyRef.LostFocus
        Try
            If Me.txtCategoryCode.Enabled = True Then
                If Me.TxtQtyRef.ReadOnly = True Then
                    Me.TxtQtyRef.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtQtyRef.BackColor = Me.txtPrimaryCategory.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtGoodPercentage_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtGoodPercentage.GotFocus
        Try
            If Me.txtCategoryCode.Enabled = True Then
                If Me.TxtGoodPercentage.ReadOnly = True Then
                    Me.TxtGoodPercentage.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtGoodPercentage.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtGoodPercentage_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtGoodPercentage.Enter
        Try
            Me.buffer = Me.TxtGoodPercentage.Text
            If Me.txtCategoryCode.Enabled = True Then
                Me.ObjCD.Number_Enter(Me.TxtGoodPercentage, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtGoodPercentage_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtGoodPercentage.Leave
        Try
            If Me.TxtGoodPercentage.Text = Me.buffer Then Exit Sub
            If Me.txtCategoryCode.Enabled = True Then
                Me.ObjCD.Number_Leave(Me.TxtGoodPercentage, sender, e, 2, Me.ObjCD)
                Me.txtTolerance.Text = Me.ObjCD.DigitGrouping(Me.GetTolerancePerc(Me.TxtGoodPercentage.Text), 2)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtGoodPercentage_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtGoodPercentage.LostFocus
        Try
            If Me.txtCategoryCode.Enabled = True Then
                If Me.TxtGoodPercentage.ReadOnly = True Then
                    Me.TxtGoodPercentage.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtGoodPercentage.BackColor = Me.txtPrimaryCategory.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtGoodPercentage_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtGoodPercentage.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtGoodPercentage, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtGoodPercentage_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtGoodPercentage.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtGoodPercentage, sender, e, 100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyRef_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyRef.Enter
        Try
            If Me.txtCategoryCode.Enabled = True Then
                Me.ObjCD.Number_Enter(Me.TxtQtyRef, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyRef_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyRef.Leave
        Try
            If Me.txtCategoryCode.Enabled = True Then
                Me.ObjCD.Number_Leave(Me.TxtQtyRef, sender, e, Me.TxtDefaultPrecision.Text, Me.ObjCD)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyRef_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyRef.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtQtyRef, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyRef_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyRef.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtQtyRef, sender, e, 0)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Kosongkan_BOM()
        Try
            Me.dgvItemBOM.Rows.Clear()
            Me.txtPrimaryProses.Text = "-1"
            Me.txtProcessCode.Text = ""
            Me.TxtQtyRef.Text = Me.ObjCD.DigitGrouping(1, Me.TxtDefaultPrecision.Text)
            Me.TxtProcessDesc.Text = ""
            Me.TxtGoodPercentage.Text = Me.ObjCD.DigitGrouping("100", 2)
            Me.txtTolerance.Text = Me.ObjCD.DigitGrouping("0", 2)
            Me.Update_TreeView()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkShowProses_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowProses.CheckedChanged
        Try
            Me.Update_TreeView()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Update_TreeView()
        Dim Child As TreeNode
        Dim ListPrimary As New List(Of Double)
        Dim TrackExplodeItem As Boolean = True

        Try
            If Me.trvBOM.Nodes(0).Tag <> Me.txtPrimaryProduct.Text Then
                'jika item baru, maka posisi expand = false
                TrackExplodeItem = False
            Else
                'jika update tree view saja, maka catat posisi expand, bila node pertama di expand
                TrackExplodeItem = Me.trvBOM.Nodes(0).IsExpanded
            End If
            'catatkan node yang di explode
            If TrackExplodeItem = True Then
                Me.Update_TreeView_GetExplodeItem(ListPrimary, Me.trvBOM.Nodes(0))
            End If

            'kosongkan node
            Me.trvBOM.Nodes.Clear()

            'update first node
            Me.Update_TreeView_FirstNode()

            'update node level 2
            For Each rowdgv As DataGridViewRow In Me.dgvItemBOM.Rows
                If rowdgv.Cells(Me.ColBOM_PrimaryProduct.Index).Value <> -1 Then
                    If Me.chkShowProses.Checked = True Then
                        Me.trvBOM.Nodes(0).Nodes.Add(Me.txtProcessCode.Text + " : " + rowdgv.Cells(Me.ColBOM_CodeNo.Index).Value + " - " + rowdgv.Cells(Me.ColBOM_Desc.Index).Value)
                    Else
                        Me.trvBOM.Nodes(0).Nodes.Add(rowdgv.Cells(Me.ColBOM_CodeNo.Index).Value + " - " + rowdgv.Cells(Me.ColBOM_Desc.Index).Value)
                    End If
                    Child = Me.trvBOM.Nodes(0).Nodes(Me.trvBOM.Nodes(0).Nodes.Count - 1)
                    Child.Tag = rowdgv.Cells(Me.ColBOM_PrimaryProduct.Index).Value
                End If
            Next

            'update sisa node level selanjutnya
            For Each Child In Me.trvBOM.Nodes(0).Nodes
                Me.Update_TreeView_IsikanChild(Child)
            Next
            'explode node seperti kondisi sebelumnya
            If TrackExplodeItem = True Then
                Me.Update_TreeView_ExplodeItem(ListPrimary, Me.trvBOM.Nodes(0))
                Me.trvBOM.Nodes(0).Expand()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            ListPrimary = Nothing
        End Try
    End Sub

    Private Sub Update_TreeView_ExplodeItem(ByVal vListPrimary As List(Of Double), ByVal vParentNode As TreeNode)
        Dim ChildNode As TreeNode

        Try
            If vParentNode.Nodes.Count > 0 Then
                If vListPrimary.Contains(vParentNode.Tag) = True Then
                    vParentNode.Expand()
                End If

                For Each ChildNode In vParentNode.Nodes
                    If ChildNode.Nodes.Count > 0 Then
                        Me.Update_TreeView_ExplodeItem(vListPrimary, ChildNode)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Update_TreeView_GetExplodeItem(ByVal vListPrimary As List(Of Double), ByVal vParentNode As TreeNode)
        Dim ChildNode As TreeNode

        Try
            If vParentNode.Nodes.Count > 0 Then
                For Each ChildNode In vParentNode.Nodes
                    If ChildNode.IsExpanded = True Then
                        vListPrimary.Add(ChildNode.Tag)
                    End If
                    If ChildNode.Nodes.Count > 0 Then
                        Me.Update_TreeView_GetExplodeItem(vListPrimary, ChildNode)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Update_TreeView_IsikanChild(ByVal vNodeParent As TreeNode)
        Dim Child As TreeNode
        Dim DTAItem As New dseMasterProduct_BOMTableAdapters.TreeViewTableTableAdapter
        Dim TBLItem As New dseMasterProduct_BOM.TreeViewTableDataTable
        Dim rowItem As dseMasterProduct_BOM.TreeViewTableRow

        Try
            'kosongkan node
            vNodeParent.Nodes.Clear()
            'isikan dr database
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Fill(TBLItem, vNodeParent.Tag)
            For Each rowItem In TBLItem.Rows
                If Me.chkShowProses.Checked = True Then
                    vNodeParent.Nodes.Add(rowItem.Proses + " : " + rowItem.RM_Code + " - " + rowItem.RM_Desc)
                Else
                    vNodeParent.Nodes.Add(rowItem.RM_Code + " - " + rowItem.RM_Desc)
                End If
                Child = vNodeParent.Nodes(vNodeParent.Nodes.Count - 1)
                Child.Tag = rowItem.primary_product_raw
            Next
            'isikan anaknya
            For Each Child In vNodeParent.Nodes
                Me.Update_TreeView_IsikanChild(Child)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
        End Try
    End Sub

    Private Sub Update_TreeView_BuangNodeLvl2(ByVal vPrimaryProduct As Double)
        Try
            For Each ChildNode As TreeNode In Me.trvBOM.Nodes(0).Nodes
                If ChildNode.Tag = vPrimaryProduct Then
                    Me.trvBOM.Nodes(0).Nodes.Remove(ChildNode)
                    Exit Sub
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Update_TreeView_TambahNodeLvl2(ByVal vPrimaryProduct As Double, ByVal vCodeNo As String, ByVal vIndex As Integer, ByVal vDesc As String)
        Try
            Dim ChildNode As New TreeNode()
            If Me.chkShowProses.Checked = True Then
                ChildNode.Text = Me.txtProcessCode.Text + " : " + vCodeNo
            Else
                ChildNode.Text = vCodeNo
            End If
            ChildNode.Tag = vPrimaryProduct
            Me.trvBOM.Nodes(0).Nodes.Insert(vIndex, ChildNode)
            'update itemnya
            Me.Update_TreeView_IsikanChild(ChildNode)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'update nama product, pada node pertama (akar)
    Private Sub Update_TreeView_FirstNode()
        Dim Code As String

        Try
            If Me.chkCategoryBindCode.Checked = True Then
                Code = Me.txtCategoryCode.Text + Me.txtProductCode.Text + " - " + Me.txtDescription.Text
            Else
                Code = Me.txtProductCode.Text + " - " + Me.txtDescription.Text
            End If
            If Me.trvBOM.Nodes.Count = 0 Then
                Me.trvBOM.Nodes.Add(Code)
            Else
                Me.trvBOM.Nodes(0).Text = Code
            End If
            Me.trvBOM.Nodes(0).Tag = Me.txtPrimaryProduct.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtProcessCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtProcessCode.KeyDown
        Dim i As New Core_Module.FrmMasterProsesCari
        Dim Obj As New ClsMasterProses

        Try
            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 And Me.txtProcessCode.ReadOnly = False And Me.txtProcessCode.Enabled = True Then
                If Obj.ApakahSudahAda(Me.txtProcessCode.Text) = True Then
                    Me.LoadWU = True
                    Me.txtPrimaryProses.Text = Obj.GetPrimary(Me.txtProcessCode.Text)
                    Me.TxtQtyRef.Focus()
                Else
                    If Me.txtProcessCode.Text <> "" Then
                        i.txtCari.Text = Me.txtProcessCode.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimary > 0 Then
                        Me.LoadWU = True
                        Me.txtPrimaryProses.Text = i.mPrimary
                        Me.txtProcessCode.Text = Obj.GetCode(Me.txtPrimaryProses.Text)
                        Me.TxtQtyRef.Focus()
                    Else
                        Me.txtProcessCode.Clear()
                        Me.txtPrimaryProses.Text = "-1"
                        Me.txtProcessCode.Focus()
                        Me.Kosongkan_BOM()
                    End If
                End If

                Me.TxtProcessDesc.Text = Me.ObjProses.GetDesc(Me.txtPrimaryProses.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            Obj = Nothing
        End Try
    End Sub

    Private Sub txtProcessCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProcessCode.Leave
        Dim i As New Core_Module.FrmMasterProsesCari
        Dim obj As New ClsMasterProses

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.txtProcessCode.Text <> "" Then
                If obj.ApakahSudahAda(Me.txtProcessCode.Text) = True Then
                    Me.LoadWU = True
                    Me.txtPrimaryProses.Text = obj.GetPrimary(Me.txtProcessCode.Text)
                Else
                    If Me.txtProcessCode.Text <> "" Then
                        i.txtCari.Text = Me.txtProcessCode.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimary > 0 Then
                        Me.LoadWU = True
                        Me.txtPrimaryProses.Text = i.mPrimary
                        Me.txtProcessCode.Text = obj.GetCode(Me.txtPrimaryProses.Text)
                        Me.TxtQtyRef.Focus()
                    Else
                        Me.txtProcessCode.Clear()
                        Me.txtPrimaryProses.Text = "-1"
                        Me.Kosongkan_BOM()
                        Me.txtProcessCode.Focus()
                    End If
                End If
            Else
                Me.txtPrimaryProses.Text = "-1"
                Me.Kosongkan_BOM()
            End If
            Me.TxtProcessDesc.Text = Me.ObjProses.GetDesc(Me.txtPrimaryProses.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFindProses_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindProcess.Click
        Dim i As New Core_Module.FrmMasterProsesCari
        Dim obj As New ClsMasterProses

        Try
            i.BackColorAll = Me.BackColorAll
            'i.cboCriteria.SelectedIndex = 0
            If Me.txtProcessCode.Text <> "" Then
                i.txtCari.Text = Me.txtProcessCode.Text
            End If

            i.ShowDialog()
            If i.mPrimary > 0 Then
                Me.LoadWU = True
                Me.txtPrimaryProses.Text = i.mPrimary
                Me.txtProcessCode.Text = obj.GetCode(Me.txtPrimaryProses.Text)
                Me.TxtQtyRef.Focus()
            Else
                Me.txtProcessCode.Clear()
                Me.txtProcessCode.Focus()
                Me.Kosongkan_BOM()
            End If
            Me.TxtProcessDesc.Text = Me.ObjProses.GetDesc(Me.txtPrimaryProses.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub SetTabBOM()
        Try
            If Me.rboGoods.Checked = True Then
                Me.TxtQtyRef.ReadOnly = False
                Me.TxtQtyRef.BackColor = Me.txtPrimaryCategory.BackColor
                Me.txtProcessCode.ReadOnly = False
                Me.txtProcessCode.BackColor = Me.txtPrimaryCategory.BackColor
                Me.TxtGoodPercentage.ReadOnly = False
                Me.TxtGoodPercentage.BackColor = Me.txtPrimaryCategory.BackColor

                If Me.txtCategoryCode.Enabled = True Then
                    Me.btnFindProcess.Enabled = True
                    Me.dgvItemBOM.AllowUserToAddRows = True
                    Me.dgvItemBOM.AllowUserToDeleteRows = True
                    Me.dgvItemBOM.ReadOnly = False
                End If
            Else
                Me.Kosongkan_BOM()
                Me.TxtQtyRef.ReadOnly = True
                Me.TxtQtyRef.BackColor = Me.TxtReadOnly.BackColor
                Me.txtProcessCode.ReadOnly = True
                Me.txtProcessCode.BackColor = Me.TxtReadOnly.BackColor
                Me.TxtGoodPercentage.ReadOnly = True
                Me.TxtGoodPercentage.BackColor = Me.TxtReadOnly.BackColor
                If Me.txtCategoryCode.Enabled = True Then
                    Me.btnFindProcess.Enabled = False
                    Me.dgvItemBOM.AllowUserToAddRows = False
                    Me.dgvItemBOM.AllowUserToDeleteRows = False
                    Me.dgvItemBOM.ReadOnly = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemBOM_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemBOM.CellBeginEdit
        Try
            If Me.txtCategoryCode.Enabled = True Then
                If e.RowIndex = Me.dgvItemBOM.Rows.Count - 1 Then
                    'jika baris terakhir yang bisa diedit: code
                    If e.ColumnIndex <> Me.ColBOM_CodeNo.Index Then
                        e.Cancel = True
                    Else
                        If Me.txtPrimaryProses.Text = "" Or Me.txtPrimaryProses.Text = -1 Then
                            MsgBox("Please specify process code first.", MsgBoxStyle.Information, "Information")
                            Me.txtProcessCode.Focus()
                            e.Cancel = True
                            Exit Sub
                        Else
                            Me.buffer = Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value
                        End If
                    End If
                Else
                    'jika bukan baris terakhir: code no, remarks, qty, good ref
                    If e.ColumnIndex = Me.ColBOM_CodeNo.Index Or _
                    e.ColumnIndex = Me.ColBOM_GoodPercentage.Index Or _
                    e.ColumnIndex = Me.ColBOM_Qty.Index Or _
                    e.ColumnIndex = Me.ColBOM_Remarks.Index Or _
                    e.ColumnIndex = Me.ColBOM_Tolerance.Index Then
                        Me.buffer = Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EditCodeNoDrillDown(ByVal vPrimaryProduct As Double)
        Try
            If Me.objProduct.CariDetail(Me.DGVKompetensi, vPrimaryProduct, Me.dgvPricing, Me.pbxProduct, Me.dgvItemBOM, _
            Me, Me.dgvItem_MultiUnit, Me.DGVAttanchment, Me.DgvWorkUnit) = True Then
                Me.SalinKeField()
                Me.btnEdit.PerformClick()
            Else
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemBOM_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemBOM.CellClick
        Dim rowDGV As DataGridViewRow
        Dim Dummy As Boolean = False

        Try
            If Me.txtProcessCode.Text = "" Then
                Exit Sub
            End If

            If Me.ColBOM_DrillDown.Index = e.ColumnIndex Then
                If Me.txtCategoryCode.Enabled = False Then
                    'pada saat view
                    If e.RowIndex <> Me.dgvItemBOM.Rows.Count - 1 Then
                        'hanya bisa klik item edit
                        Dim ijk As New frmMasterProduct
                        ijk.PrimaryDrillDown = Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value
                        ijk.EditDrillDown = True
                        ijk.ShowDialog()
                        ijk = Nothing
                        Me.UpdateItemBOM(e.RowIndex)
                        Me.Update_TreeView()
                    End If
                Else
                    'pada saat input
                    If e.RowIndex <> Me.dgvItemBOM.Rows.Count - 1 Then
                        'edit item lama
                        Dim ijk As New frmMasterProduct
                        ijk.UserCode = Me.UserCode
                        ijk.PrimaryDrillDown = Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value
                        ijk.EditDrillDown = True
                        ijk.ShowDialog()
                        ijk = Nothing
                        Me.Update_TreeView()
                    Else
                        'tambahkan item baru
                        Dim ijk As New frmMasterProduct
                        ijk.UserCode = Me.UserCode
                        ijk.NewDrillDown = True
                        ijk.ShowDialog()

                        If ijk.txtPrimaryProduct.Text <> "" And ijk.txtPrimaryProduct.Text <> -1 Then
                            If ijk.rboGoods.Checked = True Then
                                'tambahkan item baru tersebut bila dia goods
                                Dummy = Me.dgvItemBOM.AllowUserToAddRows
                                Me.dgvItemBOM.AllowUserToAddRows = False
                                Me.dgvItemBOM.Rows.Add()
                                rowDGV = Me.dgvItemBOM.Rows(Me.dgvItemBOM.Rows.Count - 1)

                                rowDGV.Cells(Me.ColBOM_PrimaryProduct.Index).Value = ijk.txtPrimaryProduct.Text
                                rowDGV.Cells(Me.ColBOM_CodeNo.Index).Value = Me.objProduct.GetProductCode(rowDGV.Cells(Me.ColBOM_PrimaryProduct.Index).Value)
                                rowDGV.Cells(Me.ColBOM_Desc.Index).Value = Me.objProduct.GetDescription(rowDGV.Cells(Me.ColBOM_PrimaryProduct.Index).Value)
                                rowDGV.Cells(Me.ColBOM_DrillDown.Index).Value = "Edit"
                                rowDGV.Cells(Me.ColBOM_Qty.Index).Value = _
                                Me.ObjCD.DigitGrouping(1, Me.objProduct.GetAngkaDecimal(rowDGV.Cells(Me.ColBOM_PrimaryProduct.Index).Value))
                                rowDGV.Cells(Me.ColBOM_GoodPercentage.Index).Value = Me.ObjCD.DigitGrouping(100, 2)
                                rowDGV.Cells(Me.ColBOM_Unit.Index).Value = _
                                Me.objProduct.GetUnitCode(rowDGV.Cells(Me.ColBOM_PrimaryProduct.Index).Value)
                                rowDGV.Cells(Me.ColBOM_Remarks.Index).Value = ""
                                Me.dgvItemBOM.AllowUserToAddRows = Dummy
                            End If
                        End If
                        ijk = Nothing

                        Me.Update_TreeView()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateItemBOM(ByVal vRowIndex As Integer)
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgvItemBOM.Rows(vRowIndex)
            rowDGV.Cells(Me.ColBOM_CodeNo.Index).Value = Me.objProduct.GetProductCode(rowDGV.Cells(Me.ColBOM_PrimaryProduct.Index).Value)
            rowDGV.Cells(Me.ColBOM_Desc.Index).Value = Me.objProduct.GetDescription(rowDGV.Cells(Me.ColBOM_PrimaryProduct.Index).Value)
            rowDGV.Cells(Me.ColBOM_Unit.Index).Value = Me.objProduct.GetUnitCode(rowDGV.Cells(Me.ColBOM_PrimaryProduct.Index).Value)
            rowDGV.Cells(Me.ColBOM_Qty.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColBOM_Qty.Index).Value, _
            Me.objProduct.GetAngkaDecimal(rowDGV.Cells(Me.ColBOM_PrimaryProduct.Index).Value))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexBOM(ByVal vPrimaryCodeNo As Double) As Double
        Try
            For Each rowDGV As DataGridViewRow In Me.dgvItemBOM.Rows
                If rowDGV.Index = Me.dgvItemBOM.Rows.Count - 1 Then
                    Return -1
                Else
                    If rowDGV.Cells(Me.ColBOM_PrimaryProduct.Index).Value = vPrimaryCodeNo Then
                        Return rowDGV.Index
                    End If
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvItemBOM_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemBOM.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvItemBOM.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryCodeNo As Double

        Try
            With row
                If IsNothing(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'bila tidak ada perubahan
                    If Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                        Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If

                    '==Code No
                    If e.ColumnIndex = Me.ColBOM_CodeNo.Index Then
                        Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        If Me.objProduct.ApakahSudahAda(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryCodeNo = Me.objProduct.GetPrimary(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexBOM(PrimaryCodeNo)
                            If index < 0 Then
                                Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value = Me.objProduct.GetPrimary(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.dgvItemBOM.Item(Me.ColBOM_CodeNo.Index, e.RowIndex).Value = Me.objProduct.GetProductCode(PrimaryCodeNo)
                                Me.dgvItemBOM.Item(Me.ColBOM_Desc.Index, e.RowIndex).Value = Me.objProduct.GetDescription(PrimaryCodeNo)
                                Me.dgvItemBOM.Item(Me.ColBOM_DrillDown.Index, e.RowIndex).Value = "Edit"
                                Me.dgvItemBOM.Item(Me.ColBOM_Qty.Index, e.RowIndex).Value = _
                                Me.ObjCD.DigitGrouping(1, Me.objProduct.GetAngkaDecimal(Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, _
                                e.RowIndex).Value))
                                Me.dgvItemBOM.Item(Me.ColBOM_GoodPercentage.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(100, 2)
                                Me.dgvItemBOM.Item(Me.ColBOM_Unit.Index, e.RowIndex).Value = _
                                Me.objProduct.GetUnitCode(Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value)

                                'jika baris ditengah, maka buang node kemudian insert node
                                If e.RowIndex <> Me.dgvItemBOM.Rows.Count - 2 Then
                                    Me.Update_TreeView_BuangNodeLvl2(Me.objProduct.GetPrimary(Me.buffer))
                                End If
                                'tambahkan node baru
                                Me.Update_TreeView_TambahNodeLvl2(Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value, _
                                Me.dgvItemBOM.Item(Me.ColBOM_CodeNo.Index, e.RowIndex).Value, e.RowIndex, Me.dgvItemBOM.Item(Me.ColBOM_Desc.Index, e.RowIndex).Value)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication in product code, please change your product code", MsgBoxStyle.Information, "Information")
                                    Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    If Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                        Me.IndexRemove = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    End If

                                    Exit Sub
                                Else
                                    Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                            End If
                        Else
                            ''diketik tidak ada lalu tampilkan form cari
                            Dim ij As New Core_Module.frmCariMasterProduct
                            ij.TipePencarian = frmCariMasterProduct.enumTipePencarian.Goods_Only
                            ij.txtCari.Text = Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.ShowDialog()
                            index = Me.GetIndexBOM(ij.PrimaryProduct)
                            If ij.PrimaryProduct = -1 Then
                                ''hapuskan
                                If Me.buffer = "" Then
                                    Me.IndexRemove = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                Else
                                    Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                                Exit Sub
                            Else
                                If index < 0 Then
                                    Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value = ij.PrimaryProduct.ToString
                                    Me.dgvItemBOM.Item(Me.ColBOM_CodeNo.Index, e.RowIndex).Value = Me.objProduct.GetProductCode(ij.PrimaryProduct)
                                    Me.dgvItemBOM.Item(Me.ColBOM_Desc.Index, e.RowIndex).Value = Me.objProduct.GetDescription(ij.PrimaryProduct)
                                    Me.dgvItemBOM.Item(Me.ColBOM_DrillDown.Index, e.RowIndex).Value = "Edit"
                                    Me.dgvItemBOM.Item(Me.ColBOM_Qty.Index, e.RowIndex).Value = _
                                    Me.ObjCD.DigitGrouping(1, Me.objProduct.GetAngkaDecimal(Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, _
                                    e.RowIndex).Value))
                                    Me.dgvItemBOM.Item(Me.ColBOM_GoodPercentage.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(100, 2)
                                    Me.dgvItemBOM.Item(Me.ColBOM_Unit.Index, e.RowIndex).Value = _
                                    Me.objProduct.GetUnitCode(Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value)

                                    'jika baris ditengah, maka buang node kemudian insert node
                                    If e.RowIndex <> Me.dgvItemBOM.Rows.Count - 2 Then
                                        Me.Update_TreeView_BuangNodeLvl2(Me.objProduct.GetPrimary(Me.buffer))
                                    End If
                                    'tambahkan node baru
                                    Me.Update_TreeView_TambahNodeLvl2(Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value, _
                                    Me.dgvItemBOM.Item(Me.ColBOM_CodeNo.Index, e.RowIndex).Value, e.RowIndex, Me.dgvItemBOM.Item(Me.ColBOM_Desc.Index, e.RowIndex).Value)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication in product code, please change your product code", MsgBoxStyle.Information, "Information")
                                        Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        If Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                            Me.IndexRemove = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        End If
                                        Exit Sub
                                    Else
                                        Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Exit Sub
                                    End If
                                End If
                            End If
                            ij = Nothing
                        End If
                    ElseIf e.ColumnIndex = Me.ColBOM_Qty.Index Then
                        'harus bilangan positif
                        If Me.ObjCD.CekBilanganPositif(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        'digit grouping
                        Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = _
                        Me.ObjCD.DigitGrouping(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value, _
                        Me.objProduct.GetAngkaDecimal(Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value))
                    ElseIf e.ColumnIndex = Me.ColBOM_GoodPercentage.Index Then
                        'harus bilangan positif
                        If Me.ObjCD.CekBilanganPositif(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        'maximal nilainya 100 %
                        If CDec(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value) > 100 Then
                            Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        'digit grouping
                        Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = _
                        Me.ObjCD.DigitGrouping(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                        Me.dgvItemBOM.Item(Me.ColBOM_Tolerance.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.GetTolerancePerc(Me.dgvItemBOM.Item(Me.ColBOM_GoodPercentage.Index, e.RowIndex).Value), 2)
                    ElseIf e.ColumnIndex = Me.ColBOM_Remarks.Index Then
                        'maksimum length = 250
                        If Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                            Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                        End If
                    ElseIf e.ColumnIndex = Me.ColBOM_Tolerance.Index Then
                        If Me.ObjCD.CekBilanganPositif(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        If CDec(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value) < 0 Then
                            Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                        Me.dgvItemBOM.Item(Me.ColBOM_GoodPercentage.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.GetGoodPercentagePerc(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value), 2)

                    End If
                Else
                    'hapus baris ybs
                    If e.ColumnIndex = Me.ColBOM_CodeNo.Index Then
                        If Me.buffer = "" Then
                            Me.IndexRemove = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        Else
                            Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                        Exit Sub
                    ElseIf e.ColumnIndex = Me.ColBOM_Qty.Index Or e.ColumnIndex = Me.ColBOM_GoodPercentage.Index Then
                        Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    ElseIf e.ColumnIndex = Me.ColBOM_Remarks.Index Then
                        Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        Exit Sub
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemBOM_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemBOM.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            If Me.dgvItemBOM.AllowUserToAddRows = True And Me.txtCategoryCode.Enabled = True Then
                'jika boleh add new row, maka setting nilai awal
                rowDGV = Me.dgvItemBOM.Rows(Me.dgvItemBOM.Rows.Count - 1)
                With rowDGV
                    .Cells(Me.ColBOM_CodeNo.Index).Value = ""
                    .Cells(Me.ColBOM_Desc.Index).Value = ""
                    .Cells(Me.ColBOM_DrillDown.Index).Value = "Add"
                    .Cells(Me.ColBOM_GoodPercentage.Index).Value = 100.0
                    .Cells(Me.ColBOM_PrimaryProduct.Index).Value = -1
                    .Cells(Me.ColBOM_Qty.Index).Value = 0
                    .Cells(Me.ColBOM_Remarks.Index).Value = ""
                    .Cells(Me.ColBOM_Unit.Index).Value = "Unit"
                    .Cells(Me.ColBOM_Tolerance.Index).Value = "0.00"
                End With
                If Me.dgvItemBOM.Rows.Count > 1 Then
                    Me.txtProcessCode.ReadOnly = True
                    Me.txtProcessCode.BackColor = Me.TxtReadOnly.BackColor
                    Me.btnFindProcess.Enabled = False
                Else
                    Me.txtProcessCode.ReadOnly = False
                    Me.txtProcessCode.BackColor = Me.txtPrimaryCategory.BackColor
                    Me.btnFindProcess.Enabled = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemBOM_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItemBOM.RowsRemoved

        Try
            If Me.txtCategoryCode.Enabled = True And Me.dgvItemBOM.AllowUserToAddRows = True Then
                If Me.dgvItemBOM.Rows.Count > 1 Then
                    Me.txtProcessCode.ReadOnly = True
                    Me.txtProcessCode.BackColor = Me.TxtReadOnly.BackColor
                    Me.btnFindProcess.Enabled = False
                Else
                    Me.txtProcessCode.ReadOnly = False
                    Me.txtProcessCode.BackColor = Me.txtPrimaryCategory.BackColor
                    Me.btnFindProcess.Enabled = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterProduct_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Me.Update_TreeView()
    End Sub

    Private Sub dgvItemBOM_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItemBOM.UserDeletedRow
        Try
            Me.Update_TreeView_BuangNodeLvl2(Me.buffer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemBOM_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItemBOM.UserDeletingRow
        Try
            'tidak ada batasan tidak boleh hapus
            Me.buffer = Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.Row.Index).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ContextMenuStrip1_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening
        Try
            If IsNothing(Me.trvBOM.SelectedNode) = True Then
                e.Cancel = True
            Else
                'jika yang di klik dan yang di select tidak sama

                'tidak boleh edit item pertama
                If Me.trvBOM.Nodes(0).Equals(Me.trvBOM.SelectedNode) Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        Dim ChildNode As TreeNode

        Try
            ChildNode = Me.trvBOM.SelectedNode
            'edit item lama
            Dim ijk As New frmMasterProduct
            ijk.PrimaryDrillDown = ChildNode.Tag
            ijk.EditDrillDown = True
            ijk.ShowDialog()
            ijk = Nothing
            Me.Update_TreeView()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub trvBOM_NodeMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles trvBOM.NodeMouseClick

        Try
            If e.Button = Windows.Forms.MouseButtons.Right Then
                Me.trvBOM.SelectedNode = e.Node
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadWorkUnitSettings()
        Dim NextStep As Boolean = True
        Dim objRateSetting As New ClsMasterProductionRate
        Dim primaryProdRate As Double = -1
        Try
            primaryProdRate = objRateSetting.GetProductionRate(Me.txtPrimaryProduct.Text, Me.txtPrimaryProses.Text, Me.txtPrimaryCategory.Text)
            If primaryProdRate <> -1 Then
                If Me.DgvWorkUnit.RowCount > 1 Then
                    If MsgBox("Process data changed, do you want to replace production rate settings?", MsgBoxStyle.YesNo, "Confirmation") <> MsgBoxResult.Yes Then
                        NextStep = False
                    End If
                End If
                If NextStep = False Then Exit Sub
                objRateSetting.UpdateDgvWorkUnit(Me.DgvWorkUnit, Me, primaryProdRate)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Dim LoadWU As Boolean = False
    Private Sub txtPrimaryProses_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrimaryProses.TextChanged
        Try
            If Me.LoadWU = True Then
                If Me.txtPrimaryProses.Text = "-1" Then
                    Exit Sub
                End If
                If Me.txtPrimaryCategory.Text = "-1" Then
                    MsgBox("Please select the category first.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                Me.LoadWorkUnitSettings()
                Me.LoadWU = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTolerance_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTolerance.Enter
        Try
            Me.buffer = Me.txtTolerance.Text
            If Me.txtCategoryCode.Enabled = True Then
                Me.ObjCD.Number_Enter(Me.txtTolerance, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub txtTolerance_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTolerance.GotFocus
        Try
            If Me.txtCategoryCode.Enabled = True Then
                If Me.txtTolerance.ReadOnly = True Then
                    Me.txtTolerance.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtTolerance.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub txtTolerance_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTolerance.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.txtTolerance, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub txtTolerance_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTolerance.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.txtTolerance, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub txtTolerance_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTolerance.Leave
        Try
            If Me.buffer = Me.txtTolerance.Text Then Exit Sub
            If Me.txtCategoryCode.Enabled = True Then
                Me.ObjCD.Number_Leave(Me.txtTolerance, sender, e, 2, Me.ObjCD)
                Me.TxtGoodPercentage.Text = Me.ObjCD.DigitGrouping(Me.GetGoodPercentagePerc(Me.txtTolerance.Text), 2)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub txtTolerance_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTolerance.LostFocus
        Try
            If Me.txtCategoryCode.Enabled = True Then
                If Me.txtTolerance.ReadOnly = True Then
                    Me.txtTolerance.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtTolerance.BackColor = Me.txtPrimaryCategory.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
