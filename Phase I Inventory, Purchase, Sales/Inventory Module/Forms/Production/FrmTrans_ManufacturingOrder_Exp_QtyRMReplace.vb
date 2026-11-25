Option Explicit On
Option Strict Off

Public Class FrmTrans_ManufacturingOrder_Exp_QtyRMReplace
    Public CodeNoRM, DescRM, WONumber, ProcessCode, vIdentifierJO, vIdentifireWO, UnitCode As String
    Public TblItemRM As New DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public Primary_Product_Code_RM As Double = -1
    Public PrimaryMO As Double = -1
    Public ReadOnly_ As Boolean = False
    Public vQtyRequired, vQtyAdditional As Decimal
    Public vPriorityReplaced As Integer = 0

    Dim LangsungClose As Boolean = False
    Dim DeleteIndex As Integer = -1
    Public DefPrec As Integer = 0
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjUnit As New Core_Module.ClsMasterUnit
    Dim Buffer As String = ""
    Dim EraseRowIndex As Integer = -1
    Dim ObjMO As New ClsTrans_ManufacturingOrder

    Private Sub SetPriorityDGV()
        Dim rowDGV As DataGridViewRow

        Try
            If Me.TxtPriority.Text = "" Or IsNumeric(Me.TxtPriority.Text) = False Then
                Me.TxtPriority.Text = "1"
            End If
            For Each rowDGV In Me.DgvItemRMReplace.Rows
                If rowDGV.Index + 1 >= Me.TxtPriority.Text Then
                    rowDGV.Cells(Me.ColNo.Index).Value = rowDGV.Index + 2
                Else
                    rowDGV.Cells(Me.ColNo.Index).Value = rowDGV.Index + 1
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ManufacturingOrder_Exp_QtyRMReplace_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.ReadOnly_ = True Then
                Exit Sub
            End If

            If Me.LangsungClose = False Then
                If MsgBox("Close and discard data changes?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    e.Cancel = True
                Else
                    Me.DialogResult = Windows.Forms.DialogResult.Cancel
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ManufacturingOrder_Exp_QtyRMReplace_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim RowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
        Dim rowDGV As DataGridViewRow

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.DgvItemRMReplace.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvItemRMReplace.AllowUserToAddRows = False
            Me.DgvItemRMReplace.Rows.Clear()

            'load data header
            Me.TxtCodeNo.Text = Me.CodeNoRM
            Me.TxtDesc.Text = Me.DescRM
            Me.TxtWONumber.Text = Me.WONumber
            Me.TxtProcessCode.Text = Me.ProcessCode
            Me.TxtPriority.Text = vPriorityReplaced
            Me.LblUnit1.Text = Me.UnitCode
            Me.LblUnit2.Text = Me.UnitCode
            Me.LblUnit3.Text = Me.UnitCode

            'load dgv bundling
            Me.DefPrec = Me.ObjProduct.GetAngkaDecimal(Me.Primary_Product_Code_RM)
            Me.DgvItemRMReplace.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvItemRMReplace.AllowUserToAddRows = False
            Me.TblItemRM.DefaultView.RowFilter = "[Identifier JO] = '" + Me.vIdentifierJO + "' AND [Identifier WO] = '" + Me.vIdentifireWO + _
            "' AND [primary code rm replaced] = " + Me.Primary_Product_Code_RM.ToString + " AND [is replacement] = 1"
            Me.TblItemRM.DefaultView.Sort = "[use priority]"

            For Each DV As Data.DataRowView In Me.TblItemRM.DefaultView
                RowItemRM = DV.Row
                Me.DgvItemRMReplace.Rows.Add()
                rowDGV = Me.DgvItemRMReplace.Rows(Me.DgvItemRMReplace.Rows.Count - 1)
                rowDGV.Cells(Me.ColNo.Index).Value = RowItemRM.use_priority
                rowDGV.Cells(Me.ColCodeNo.Index).Value = Me.ObjProduct.GetProductCode(RowItemRM.Primary_Product_Code_RM)
                rowDGV.Cells(Me.ColDesc.Index).Value = Me.ObjProduct.GetDescription(RowItemRM.Primary_Product_Code_RM)
                rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = RowItemRM.Primary_Product_Code_RM
                rowDGV.Cells(Me.ColPrimaryUnitReplaced.Index).Value = RowItemRM.primary_unit_replaced
                rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(RowItemRM.Required_Qty, Me.ObjProduct.GetAngkaDecimal(RowItemRM.Primary_Product_Code_RM))
                rowDGV.Cells(Me.ColRatioReplace.Index).Value = RowItemRM.ratio_replaced
                rowDGV.Cells(Me.ColRatioReplacement.Index).Value = RowItemRM.ratio_replacement
                rowDGV.Cells(Me.ColQtyReplaced.Index).Value = Me.ObjCD.DigitGrouping(RowItemRM.qty_replaced, Me.DefPrec)
                rowDGV.Cells(Me.ColRatioDesc.Index).Value = Me.GetRatio(rowDGV.Cells(Me.ColRatioReplacement.Index).Value, _
                rowDGV.Cells(Me.ColRatioReplace.Index).Value, rowDGV.Cells(Me.ColUnit.Index).Value, rowDGV.Cells(Me.ColUnitReplaced.Index).Value, _
                Me.CodeNoRM, rowDGV.Cells(Me.ColCodeNo.Index).Value, Me.DefPrec, Me.ObjProduct.GetAngkaDecimal(rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value))
                rowDGV.Cells(Me.ColUnit.Index).Value = Me.ObjProduct.GetUnitCode(RowItemRM.Primary_Product_Code_RM)
                rowDGV.Cells(Me.ColUnitReplaced.Index).Value = Me.UnitCode
            Next
            'Me.DgvItemRMReplace.Sort(Me.ColCodeNo, System.ComponentModel.ListSortDirection.Ascending)
            Me.DgvItemRMReplace.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DgvItemRMReplace.AllowUserToAddRows = True

            If Me.ReadOnly_ = True Then
                Me.btnOK.Enabled = False
                Me.DgvItemRMReplace.AllowUserToAddRows = False
                Me.DgvItemRMReplace.ReadOnly = True
                Me.DgvItemRMReplace.AllowUserToDeleteRows = False
                Me.btnOK.Visible = False
                Me.TxtPriority.ReadOnly = True
            Else
                Me.btnOK.Enabled = True
                Me.DgvItemRMReplace.AllowUserToAddRows = True
                Me.DgvItemRMReplace.ReadOnly = False
                Me.DgvItemRMReplace.AllowUserToDeleteRows = True
            End If

            Me.UpdateHeaderQty()
            If CDec(Me.TxtPriority.Text) = 0 Then
                Me.TxtPriority.Text = 1
            End If
            Me.SetPriorityDGV()
            'If Me.ReadOnly_ = True Then
            '    Me.DgvItemRMReplace.AllowUserToAddRows = True
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetRatio(ByVal vRatioReplacement As Decimal, ByVal vRatioReplaced As Decimal, _
    ByVal vUnitReplacement As String, ByVal vUnitReplaced As String, _
    ByVal vCodeNoReplaced As String, ByVal vCodeNoReplacement As String, _
    ByVal vDefPrecReplaced As Integer, ByVal vDefPrecReplacement As Integer) As String
        Try
            Return Me.ObjCD.DigitGrouping(vRatioReplaced, vDefPrecReplaced) + " " + vUnitReplaced + " replaced by " + _
            Me.ObjCD.DigitGrouping(vRatioReplacement, vDefPrecReplacement) + " " + vUnitReplacement
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub DgvItemRMAdd_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvItemRMReplace.CellBeginEdit
        Try
            If Me.DgvItemRMReplace.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value = -1 Then
                If e.ColumnIndex = Me.ColCodeNo.Index Then
                    Me.Buffer = Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColCodeNo.Index Or _
                e.ColumnIndex = Me.ColQty.Index Or _
                e.ColumnIndex = Me.ColQtyReplaced.Index Then
                    Me.Buffer = Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItemRMAdd_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItemRMReplace.CellEndEdit
        Try
            If IsNothing(Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            Else
                If e.ColumnIndex = Me.ColCodeNo.Index Then
                    Me.EndEditCodeNo(sender, e)
                    Me.SetPriorityDGV()
                ElseIf e.ColumnIndex = Me.ColQty.Index Then
                    Me.EndEditQty(sender, e)
                ElseIf e.ColumnIndex = Me.ColQtyReplaced.Index Then
                    Me.EndEditQtyReplaced(sender, e)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndex(ByVal vPrimaryRM As Double) As Integer
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.DgvItemRMReplace.Rows
                If rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryRM Then
                    Return rowDGV.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub EndEditCodeNo(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim row As DataGridViewRow = Me.DgvItemRMReplace.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim PrimaryProduct As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Dim TypeAccProd As Integer = 0
        Dim DefPrecReplacement As Integer = 0

        Try
            With row
                If IsNothing(Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    If Me.ObjProduct.ApakahSudahAda(Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        If Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                            Exit Sub
                        End If
                        PrimaryProduct = Me.ObjProduct.GetPrimary(Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value)
                        TypeAccProd = ObjMO.GetTypeAccProduct(PrimaryProduct)
                        If TypeAccProd = 0 Then
                            'pastikan tidak double dengan baris yang lain
                            If Me.GetIndex(PrimaryProduct) <> -1 Then
                                MsgBox("Code no duplication detected, please choose another code no.", MsgBoxStyle.Information, "Information")
                                If Me.Buffer = "" Then
                                    GoTo HapusBaris
                                Else
                                    Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If

                            'pastikan tidak sama dengan kode barang yg di replace
                            If Me.GetIndex(PrimaryProduct) = Me.Primary_Product_Code_RM Then
                                MsgBox("Please specify code no different from code no replaced.", MsgBoxStyle.Information, "Information")
                                If Me.Buffer = "" Then
                                    GoTo HapusBaris
                                Else
                                    Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If

                            DefPrecReplacement = Me.ObjProduct.GetAngkaDecimal(PrimaryProduct)
                            Me.DgvItemRMReplace.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value = PrimaryProduct
                            Me.DgvItemRMReplace.Item(Me.ColCodeNo.Index, e.RowIndex).Value = Me.ObjProduct.GetProductCode(PrimaryProduct)
                            Me.DgvItemRMReplace.Item(Me.ColDesc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(PrimaryProduct)
                            Me.DgvItemRMReplace.Item(Me.ColNo.Index, e.RowIndex).Value = e.RowIndex + 1
                            Me.DgvItemRMReplace.Item(Me.ColPrimaryUnitReplaced.Index, e.RowIndex).Value = Me.Primary_Product_Code_RM
                            Me.DgvItemRMReplace.Item(Me.ColQty.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(0, DefPrecReplacement)
                            Me.DgvItemRMReplace.Item(Me.ColQtyReplaced.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(0, Me.DefPrec)
                            Me.DgvItemRMReplace.Item(Me.ColRatioReplace.Index, e.RowIndex).Value = "1"
                            Me.DgvItemRMReplace.Item(Me.ColRatioReplacement.Index, e.RowIndex).Value = "1"
                            Me.DgvItemRMReplace.Item(Me.ColUnit.Index, e.RowIndex).Value = Me.ObjProduct.GetUnitCode(PrimaryProduct)
                            Me.DgvItemRMReplace.Item(Me.ColUnitReplaced.Index, e.RowIndex).Value = Me.LblUnit1.Text
                            Me.DgvItemRMReplace.Item(Me.ColRatioDesc.Index, e.RowIndex).Value = Me.GetRatio(1, 1, Me.DgvItemRMReplace.Item(Me.ColUnit.Index, e.RowIndex).Value, _
                            Me.LblUnit1.Text, Me.TxtCodeNo.Text, Me.DgvItemRMReplace.Item(Me.ColCodeNo.Index, e.RowIndex).Value, Me.DefPrec, DefPrecReplacement)
                        Else
                            MsgBox("Please specify code no with goods type.", MsgBoxStyle.Information, "Information")
                            GoTo HapusBaris
                        End If
                    Else
                        Dim ij As New Core_Module.frmCariMasterProduct
                        ij.TipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods
                        ij.txtCari.Text = Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.BackColorAll = Me.BackColorAll
                        ij.ShowDialog()
                        TypeAccProd = ObjMO.GetTypeAccProduct(ij.PrimaryProduct)

                        If ij.PrimaryProduct <> -1 Then
                            If TypeAccProd = 0 Then
                                'pastikan tidak double dengan baris yang lain
                                If Me.GetIndex(ij.PrimaryProduct) <> -1 Then
                                    MsgBox("Code no duplication detected, please choose another code no.", MsgBoxStyle.Information, "Information")
                                    If Me.Buffer = "" Then
                                        GoTo HapusBaris
                                    Else
                                        Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        Exit Sub
                                    End If
                                End If

                                'pastikan tidak sama dengan kode barang yg di replace
                                If Me.GetIndex(ij.PrimaryProduct) = Me.Primary_Product_Code_RM Then
                                    MsgBox("Please specify code no different from code no replaced.", MsgBoxStyle.Information, "Information")
                                    If Me.Buffer = "" Then
                                        GoTo HapusBaris
                                    Else
                                        Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        Exit Sub
                                    End If
                                End If

                                DefPrecReplacement = Me.ObjProduct.GetAngkaDecimal(ij.PrimaryProduct)
                                Me.DgvItemRMReplace.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value = ij.PrimaryProduct
                                Me.DgvItemRMReplace.Item(Me.ColCodeNo.Index, e.RowIndex).Value = Me.ObjProduct.GetProductCode(ij.PrimaryProduct)
                                If Me.DgvItemRMReplace.Item(Me.ColCodeNo.Index, e.RowIndex).Value = Me.Buffer Then
                                    Exit Sub
                                End If
                                Me.DgvItemRMReplace.Item(Me.ColDesc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(ij.PrimaryProduct)
                                Me.DgvItemRMReplace.Item(Me.ColNo.Index, e.RowIndex).Value = e.RowIndex + 1
                                Me.DgvItemRMReplace.Item(Me.ColPrimaryUnitReplaced.Index, e.RowIndex).Value = Me.Primary_Product_Code_RM
                                Me.DgvItemRMReplace.Item(Me.ColQty.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(0, DefPrecReplacement)
                                Me.DgvItemRMReplace.Item(Me.ColQtyReplaced.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(0, Me.DefPrec)
                                Me.DgvItemRMReplace.Item(Me.ColRatioReplace.Index, e.RowIndex).Value = "1"
                                Me.DgvItemRMReplace.Item(Me.ColRatioReplacement.Index, e.RowIndex).Value = "1"
                                Me.DgvItemRMReplace.Item(Me.ColUnit.Index, e.RowIndex).Value = Me.ObjProduct.GetUnitCode(ij.PrimaryProduct)
                                Me.DgvItemRMReplace.Item(Me.ColUnitReplaced.Index, e.RowIndex).Value = Me.LblUnit1.Text
                                Me.DgvItemRMReplace.Item(Me.ColRatioDesc.Index, e.RowIndex).Value = Me.GetRatio(1, 1, Me.DgvItemRMReplace.Item(Me.ColUnit.Index, e.RowIndex).Value, _
                                Me.LblUnit1.Text, Me.TxtCodeNo.Text, Me.DgvItemRMReplace.Item(Me.ColCodeNo.Index, e.RowIndex).Value, Me.DefPrec, DefPrecReplacement)

                            Else
                                MsgBox("Please specify code no with goods type.", MsgBoxStyle.Information, "Information")
                                GoTo HapusBaris
                            End If
                        ElseIf ij.PrimaryProduct = -1 Then
HapusBaris:
                            If Me.DgvItemRMReplace.Rows.Count - 1 <> e.RowIndex Then
                                If String.IsNullOrEmpty(Me.Buffer) = True Then
                                    Me.EraseRowIndex = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                Else
                                    Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                End If
                                'Me.ObjCD.SetNumber(Me.DgvItemRMReplace)
                                Exit Sub
                            End If
                            If Me.DgvItemRMReplace.Item(Me.ColCodeNo.Index, e.RowIndex).Value <> "" Then
                                If String.IsNullOrEmpty(Me.Buffer) = True Then
                                    Me.EraseRowIndex = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                Else
                                    Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                End If
                                'Me.ObjCD.SetNumber(Me.DgvItemRMReplace)
                                Exit Sub
                            End If
                        End If
                        ij = Nothing
                    End If
                    'Me.ObjCD.SetNumber(Me.DgvItemRMReplace)
                Else
                    Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditQtyReplaced(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim QtyReplacement As Decimal = 0
        Dim DefPrecReplacement As Integer = 0

        Try
            If IsNothing(Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                'pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please specify positive number.", MsgBoxStyle.Information, "Information")
                    Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                'pastikan nilainya beda dengan buffer
                If CDec(Me.Buffer) = CDec(Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value) Then
                    Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                'digit grouping
                Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value, _
                Me.DefPrec)

                'pastikan nilai qty replacement positif
                'qty replacement = (qty replaced * ratio replacement) / ratio replaced
                DefPrecReplacement = Me.ObjProduct.GetAngkaDecimal(Me.DgvItemRMReplace.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
                QtyReplacement = Math.Round((CDec(Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value) * _
                CDec(Me.DgvItemRMReplace.Item(Me.ColRatioReplacement.Index, e.RowIndex).Value)) / CDec(Me.DgvItemRMReplace.Item(Me.ColRatioReplace.Index, e.RowIndex).Value), _
                DefPrecReplacement, MidpointRounding.AwayFromZero)

                'pastikan qty replacement <= qty replaced
                If QtyReplacement <= 0 Then
                    MsgBox("Qty replacement =0, please specify another number", MsgBoxStyle.Information, "Information")
                    Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                'update nilai qty replace
                Me.DgvItemRMReplace.Item(Me.ColQty.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(QtyReplacement, DefPrecReplacement)

                'update nilai pada txt header
                UpdateHeaderQty()
            Else
                Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditQty(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim QtyReplaced As Decimal = 0
        Dim DefPrecReplacement As Integer = 0

        Try
            If IsNothing(Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                'pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please specify positive number.", MsgBoxStyle.Information, "Information")
                    Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                'pastikan nilainya beda dengan buffer
                If CDec(Me.Buffer) = CDec(Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value) Then
                    Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                'digit grouping
                DefPrecReplacement = Me.ObjProduct.GetAngkaDecimal(Me.DgvItemRMReplace.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
                Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value, _
                DefPrecReplacement)

                'pastikan nilai qty replaced positif
                'qty replaced = (qty replacement * ratio replaced) / ratio replacement
                QtyReplaced = Math.Round((CDec(Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value) * _
                CDec(Me.DgvItemRMReplace.Item(Me.ColRatioReplacement.Index, e.RowIndex).Value)) / CDec(Me.DgvItemRMReplace.Item(Me.ColRatioReplace.Index, e.RowIndex).Value), _
                Me.DefPrec, MidpointRounding.AwayFromZero)

                'pastikan qty replacement <= qty replaced
                If QtyReplaced <= 0 Then
                    MsgBox("Qty replaced = 0, please specify another number", MsgBoxStyle.Information, "Information")
                    Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                'update nilai qty replace
                Me.DgvItemRMReplace.Item(Me.ColQtyReplaced.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(QtyReplaced, Me.DefPrec)

                'update nilai pada txt header
                UpdateHeaderQty()
            Else
                Me.DgvItemRMReplace.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateHeaderQty()
        Dim rowDGV As DataGridViewRow
        Dim QtyReplaced As Decimal = 0

        Try
            Me.TxtQtyRequired.Text = Me.ObjCD.DigitGrouping(Me.vQtyRequired, Me.DefPrec)
            Me.TxtQtyAdditional.Text = Me.ObjCD.DigitGrouping(Me.vQtyAdditional, Me.DefPrec)
            For Each rowDGV In Me.DgvItemRMReplace.Rows
                QtyReplaced += CDec(rowDGV.Cells(Me.ColQtyReplaced.Index).Value)
            Next
            Me.TxtQtyReplaced.Text = Me.ObjCD.DigitGrouping(QtyReplaced, Me.DefPrec)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItemRMAdd_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItemRMReplace.CellContentClick
        Dim rowDGV As DataGridViewRow

        Try
            If e.RowIndex <> -1 Then
                If e.ColumnIndex = Me.ColRatioDesc.Index Then
                    Dim ij As New FrmTrans_ManufacturingOrder_Exp_RatioReplace
                    ij.vDefPrecReplaced = Me.DefPrec
                    ij.vDefPrecReplacement = Me.ObjProduct.GetAngkaDecimal(Me.DgvItemRMReplace.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
                    ij.vPrimaryProductReplaced = Me.Primary_Product_Code_RM
                    ij.vPrimaryProductReplacement = Me.DgvItemRMReplace.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value
                    ij.vQtyReplaced = Me.DgvItemRMReplace.Item(Me.ColQtyReplaced.Index, e.RowIndex).Value
                    ij.vQtyReplacement = Me.DgvItemRMReplace.Item(Me.ColQty.Index, e.RowIndex).Value
                    ij.vRatioReplacement = Me.DgvItemRMReplace.Item(Me.ColRatioReplacement.Index, e.RowIndex).Value
                    ij.vRatioReplaced = Me.DgvItemRMReplace.Item(Me.ColRatioReplace.Index, e.RowIndex).Value
                    ij.vUnitCodeReplaced = Me.DgvItemRMReplace.Item(Me.ColUnitReplaced.Index, e.RowIndex).Value
                    ij.vUnitCodeReplacement = Me.DgvItemRMReplace.Item(Me.ColUnit.Index, e.RowIndex).Value

                    If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                        Me.DgvItemRMReplace.Item(Me.ColQtyReplaced.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(ij.TxtQtyReplaced.Text, ij.vDefPrecReplaced)
                        Me.DgvItemRMReplace.Item(Me.ColQty.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(ij.TxtQtyReplacement.Text, ij.vDefPrecReplacement)
                        Me.DgvItemRMReplace.Item(Me.ColRatioReplacement.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(ij.TxtRatioReplacement.Text, ij.vDefPrecReplacement)
                        Me.DgvItemRMReplace.Item(Me.ColRatioReplace.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(ij.TxtRatioReplaced.Text, ij.vDefPrecReplaced)
                        rowDGV = Me.DgvItemRMReplace.Rows(e.RowIndex)
                        rowDGV.Cells(Me.ColRatioDesc.Index).Value = Me.GetRatio(rowDGV.Cells(Me.ColRatioReplacement.Index).Value, _
                        rowDGV.Cells(Me.ColRatioReplace.Index).Value, rowDGV.Cells(Me.ColUnit.Index).Value, rowDGV.Cells(Me.ColUnitReplaced.Index).Value, _
                        Me.CodeNoRM, rowDGV.Cells(Me.ColCodeNo.Index).Value, Me.DefPrec, Me.ObjProduct.GetAngkaDecimal(rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value))
                        Me.UpdateHeaderQty()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(10)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.DgvItemRMReplace.Rows.RemoveAt(EraseRowIndex)
            'Me.ObjCD.SetNumber(Me.DgvItemRMReplace)
            Me.SetPriorityDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItemRMReplace_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvItemRMReplace.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.DgvItemRMReplace.Rows(Me.DgvItemRMReplace.Rows.Count - 1)

            rowDGV.Cells(Me.ColNo.Index).Value = Me.DgvItemRMReplace.Rows.Count
            rowDGV.Cells(Me.ColCodeNo.Index).Value = ""
            rowDGV.Cells(Me.ColDesc.Index).Value = ""
            rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = -1
            rowDGV.Cells(Me.ColPrimaryUnitReplaced.Index).Value = Me.Primary_Product_Code_RM
            rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(0, 0)
            rowDGV.Cells(Me.ColRatioReplace.Index).Value = 1
            rowDGV.Cells(Me.ColRatioReplacement.Index).Value = 1
            rowDGV.Cells(Me.ColQtyReplaced.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefPrec)
            rowDGV.Cells(Me.ColRatioDesc.Index).Value = ""
            rowDGV.Cells(Me.ColUnit.Index).Value = ""
            rowDGV.Cells(Me.ColUnitReplaced.Index).Value = ""

            'Me.ObjCD.SetNumber(Me.DgvItemRMReplace)
            Me.SetPriorityDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItemRMReplace_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvItemRMReplace.RowsRemoved
        Try
            'Me.ObjCD.SetNumber(Me.DgvItemRMReplace)
            Me.SetPriorityDGV()
            Me.UpdateHeaderQty()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim rowDGV As DataGridViewRow

        Try
            Me.DgvItemRMReplace.ClearSelection()
            For Each rowDGV In Me.DgvItemRMReplace.Rows
                If rowDGV.Index <> Me.DgvItemRMReplace.RowCount - 1 Then
                    If CDec(rowDGV.Cells(Me.ColQty.Index).Value) = 0 Or _
                    CDec(rowDGV.Cells(Me.ColQtyReplaced.Index).Value) = 0 Then
                        rowDGV.Selected = True
                        MsgBox("Row " + (rowDGV.Index + 1).ToString + " has qty = 0, please fix.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                End If
            Next
            'pastikan qty required + qty addition >= qty replacement
            Me.UpdateHeaderQty()
            If CDec(Me.TxtQtyAdditional.Text) + CDec(Me.TxtQtyRequired.Text) < CDec(Me.TxtQtyReplaced.Text) Then
                MsgBox("Qty Required + Qty Additional must be smaller or equals to Qty Replaced.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If MsgBox("Accept data changes?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.LangsungClose = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            If Me.TabControl1.SelectedTab.Name = Me.TabPage_DefaultData.Name Or _
            Me.TabControl1.SelectedTab.Name = Me.TabPage_History.Name Then
                MsgBox("Under construction, please wait until we finalize this.", MsgBoxStyle.Information, "Information")
                Me.TabControl1.SelectedTab = Me.TabPage_Data
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPriority_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPriority.Enter
        Try
            If Me.TxtPriority.ReadOnly = False Then
                Me.ObjCD.Number_Enter(Me.TxtPriority, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPriority_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPriority.KeyDown
        Try
            If Me.TxtPriority.ReadOnly = False Then
                Me.ObjCD.Number_KeyDown(Me.TxtPriority, sender, e, Me.DgvItemRMReplace.RowCount)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPriority_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPriority.KeyUp
        Try
            If Me.TxtPriority.ReadOnly = False Then
                Me.ObjCD.Number_KeyUp(Me.TxtPriority, sender, e, False)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPriority_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPriority.Leave
        Try
            If Me.TxtPriority.ReadOnly = False Then
                If CDec(Me.TxtPriority.Text) = 0 Then
                    Me.TxtPriority.Text = 1
                End If
                Me.SetPriorityDGV()
                Me.ObjCD.Number_Leave(Me.TxtPriority, sender, e, 0, Me.ObjCD)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUp.Click
        Try
            If Me.DgvItemRMReplace.SelectedRows.Count > 0 Then
                Me.btnUp.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.DgvItemRMReplace.SelectedRows(0)
                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.DgvItemRMReplace.Rows.Remove(row)
                    Me.DgvItemRMReplace.Rows.InsertRange(indexbaru, row)
                    Me.DgvItemRMReplace.Focus()
                    Me.DgvItemRMReplace.ClearSelection()
                    row.Selected = True
                    Me.SetPriorityDGV()
                    If row.Displayed = False Then
                        Me.ShowRow(row, True, Me.DgvItemRMReplace)
                    End If
                End If
                SetBtnUpDownEnableDisable
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            If Me.DgvItemRMReplace.SelectedRows.Count > 0 Then
                Me.btnUp.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.DgvItemRMReplace.SelectedRows(0)
                If row.Index <> Me.DgvItemRMReplace.Rows.Count - 1 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.DgvItemRMReplace.Rows.Remove(row)
                    Me.DgvItemRMReplace.Rows.InsertRange(indexbaru, row)
                    Me.DgvItemRMReplace.Focus()
                    Me.DgvItemRMReplace.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.DgvItemRMReplace)
                    Me.ShowRow(row, False, Me.DgvItemRMReplace)
                End If
                SetBtnUpDownEnableDisable()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowRow(ByVal row As DataGridViewRow, ByVal Up As Boolean, ByVal vdGV As DataGridView)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < vdGV.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    vdGV.FirstDisplayedScrollingRowIndex = vdGV.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > vdGV.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While vdGV.Rows(row.Index).Displayed = False
                    vdGV.FirstDisplayedScrollingRowIndex = vdGV.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If vdGV.DisplayedRowCount(True) <> vdGV.DisplayedRowCount(False) Then
                    vdGV.FirstDisplayedScrollingRowIndex = vdGV.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetBtnUpDownEnableDisable()
        Try
            If Me.DgvItemRMReplace.SelectedRows.Count > 0 Then
                Me.btnUp.Enabled = True
                Me.btnDown.Enabled = True
                If Me.DgvItemRMReplace.SelectedRows(0).Index >= Me.DgvItemRMReplace.RowCount - 2 Then
                    Me.btnDown.Enabled = False
                End If
                If Me.DgvItemRMReplace.SelectedRows(0).Index = 0 Then
                    Me.btnUp.Enabled = False
                End If
            Else
                Me.btnUp.Enabled = False
                Me.btnDown.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItemRMReplace_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvItemRMReplace.SelectionChanged
        Try
            Me.SetBtnUpDownEnableDisable()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class