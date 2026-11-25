Option Explicit On
Option Strict Off

Public Class FrmTransaksiStockInAdj_SetAmount
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim ObjClosing As New ClsTransaksiMonthlyClosingItem
    Dim ObjCD As New Core_Module.ClsCekData
    Dim objGroup As New Core_Module.ClsMasterProductGroup
    Dim objCategory As New Core_Module.ClsMasterCategory
    Dim objProduct As New Core_Module.clsMasterProduct
    Dim ObjCur As New Core_Module.clsMasterCurrency 

    Dim Buffer As String = ""
    Dim index As Integer = -1

    Dim ProdGroup As String = "Group"
    Dim ProdBrand As String = "Brand"
    Dim ProdSubBrand As String = "Sub Brand"
    Dim ProdCategory As String = "Category"
    Dim ProdProduct As String = "Product"
    Dim DefPrec As Integer

    Public Enum EnumGroupLevel As Integer
        LvlGroup
        LvlBrand
        LvlSubBrand
        LvlCategory
        LvlProduct
    End Enum

    Public Function GetGroupLevel_Enum(ByVal vGroupLvlString As String) As EnumGroupLevel

        Try
            If vGroupLvlString = ProdGroup Then
                Return EnumGroupLevel.LvlGroup
            ElseIf vGroupLvlString = ProdBrand Then
                Return EnumGroupLevel.LvlBrand
            ElseIf vGroupLvlString = ProdSubBrand Then
                Return EnumGroupLevel.LvlSubBrand
            ElseIf vGroupLvlString = ProdCategory Then
                Return EnumGroupLevel.LvlCategory
            ElseIf vGroupLvlString = ProdProduct Then
                Return EnumGroupLevel.LvlProduct
            End If

            Return EnumGroupLevel.LvlProduct
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetGroupLevel_String(ByVal vGroupLvlEnum As EnumGroupLevel) As String

        Try
            If vGroupLvlEnum = EnumGroupLevel.LvlGroup Then
                Return ProdGroup
            ElseIf vGroupLvlEnum = EnumGroupLevel.LvlBrand Then
                Return ProdBrand
            ElseIf vGroupLvlEnum = EnumGroupLevel.LvlSubBrand Then
                Return ProdSubBrand
            ElseIf vGroupLvlEnum = EnumGroupLevel.LvlCategory Then
                Return ProdCategory
            ElseIf vGroupLvlEnum = EnumGroupLevel.LvlProduct Then
                Return ProdProduct
            End If

            Return ProdProduct
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ProdProduct
        End Try
    End Function

    Private Sub FrmTransaksiStockInAdj_SetAmount_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try 
            Control.CheckForIllegalCrossThreadCalls = False

            If Me.ObjClosing.GetJml > 0 Then
                MsgBox("There are closing already, you can't edit begining balance amount.", MsgBoxStyle.Information, "Information")
                Me.bgwClose.RunWorkerAsync()
                Exit Sub
            End If

            Me.dgvGroup.Rows.Clear()
            Me.dgvGroup.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvProd.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            ProdGroup = Me.GetGroupLevel_String(EnumGroupLevel.LvlGroup)
            ProdBrand = Me.GetGroupLevel_String(EnumGroupLevel.LvlBrand)
            ProdSubBrand = Me.GetGroupLevel_String(EnumGroupLevel.LvlSubBrand)
            ProdCategory = Me.GetGroupLevel_String(EnumGroupLevel.LvlCategory)
            ProdProduct = Me.GetGroupLevel_String(EnumGroupLevel.LvlProduct)

            Me.lblProgess.Text = ""
            Me.DefPrec = Me.ObjCur.GetDecimal_DefaultCurrency
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProd_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvProd.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColUnitPrice.Index Or e.ColumnIndex = Me.ColAmount.Index Or e.ColumnIndex = Me.ColModifiy.Index Then
                Me.Buffer = Me.dgvProd(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProd_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvProd.CellContentClick
        Try
            If e.ColumnIndex = Me.ColDetail.Index And e.RowIndex > -1 Then
                Dim i As New FrmTransaksiStockInAdj_SetAmount_Detail
                i.vDecimalDefault = Me.DefPrec
                i.vPrimaryProduct = Me.dgvProd(Me.ColPrimaryProd.Index, e.RowIndex).Value
                i.BackColorAll = Me.BackColorAll
                i.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProd_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvProd.CellEndEdit
        Try
            If IsNothing(Me.dgvProd(e.ColumnIndex, e.RowIndex).Value) Then
                Me.dgvProd(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            End If
            If String.IsNullOrEmpty(Me.dgvProd(e.ColumnIndex, e.RowIndex).Value) Then
                Me.dgvProd(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColUnitPrice.Index Or e.ColumnIndex = Me.ColAmount.Index Then
                If Not IsNumeric(Me.dgvProd(e.ColumnIndex, e.RowIndex).Value) Then
                    Me.dgvProd(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If
                Me.dgvProd(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvProd(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                If CDec(Me.dgvProd(e.ColumnIndex, e.RowIndex).Value) < 0 Then
                    Me.dgvProd(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If
                If e.ColumnIndex = Me.ColUnitPrice.Index Then
                    Me.dgvProd(Me.ColAmount.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvProd(Me.ColUnitPrice.Index, e.RowIndex).Value) * CDec(Me.dgvProd(Me.ColQty.Index, e.RowIndex).Value), Me.DefPrec, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                End If
                If e.ColumnIndex = Me.ColAmount.Index Then
                    Me.dgvProd(Me.ColUnitPrice.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvProd(Me.ColAmount.Index, e.RowIndex).Value) / CDec(Me.dgvProd(Me.ColQty.Index, e.RowIndex).Value), Me.DefPrec, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                End If
                Me.UpdateUnitPrice(Me.dgvProd(Me.ColPrimaryProd.Index, e.RowIndex).Value, Me.dgvProd(Me.ColAmount.Index, e.RowIndex).Value, Me.dgvProd(Me.ColUnitPrice.Index, e.RowIndex).Value)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function UpdateUnitPrice(ByVal vPrimaryProduct As Double, ByVal vTotalAmount As Decimal, ByVal vUnitPrice As Decimal) As Boolean
        Dim dta As New dseTransaksiStockInOut_SetAmountTableAdapters.DataTable1TableAdapter
        Dim tbl As New dseTransaksiStockInOut_SetAmount.DataTable1DataTable
        Dim row As dseTransaksiStockInOut_SetAmount.DataTable1Row
        Dim AmountBuffer As Decimal = 0
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.FillBy_primary_prod(tbl, vPrimaryProduct) > 0 Then
                For i As Integer = 0 To tbl.Rows.Count - 1
                    row = tbl(i)
                    If i = tbl.Rows.Count - 1 Then
                        dta.UpdateQuery_SetAmount(vTotalAmount - AmountBuffer, vUnitPrice, row.Primary)
                    Else
                        Dim Amount As Decimal = Math.Round(row.Qty * vUnitPrice, Me.DefPrec, MidpointRounding.AwayFromZero)
                        AmountBuffer += Amount
                        dta.UpdateQuery_SetAmount(Amount, vUnitPrice, row.Primary)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BgwProduct_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwProduct.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwProduct_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwProduct.RunWorkerCompleted
        Try
            If Me.index <> Me.dgvGroup.RowCount - 1 Then
                Me.dgvGroup.Rows.RemoveAt(Me.index)
                Me.ObjCD.SetNumber(Me.dgvGroup)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProduct_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvGroup.CellBeginEdit
        Try
            If e.ColumnIndex = Me.colProduct_GroupLevel.Index Or e.ColumnIndex = Me.colProduct_Code.Index Then
                Me.Buffer = Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProduct_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvGroup.CellEndEdit

        Try
            If IsNothing(Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If String.IsNullOrEmpty(Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If IsNothing(Me.Buffer) = True Then
                        Me.index = e.RowIndex
                        Me.BgwProduct.RunWorkerAsync()
                        Exit Sub
                    End If
                    If Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer And Me.dgvGroup(Me.colProduct_Primary.Index, e.RowIndex).Value <> "-1" Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.colProduct_GroupLevel.Index Then
                        cellEndEditGroupLevel(sender, e)
                    End If
                    If e.ColumnIndex = Me.colProduct_Code.Index Then
                        cellEndEditCode(sender, e)
                    End If
                Else
                    If IsNothing(Me.Buffer) = False Then
                        If String.IsNullOrEmpty(Buffer) = False Then
                            Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                    Me.index = e.RowIndex
                    Me.BgwProduct.RunWorkerAsync()
                End If
            Else
                If IsNothing(Me.Buffer) = False Then
                    If String.IsNullOrEmpty(Buffer) = False Then
                        Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If
                Me.index = e.RowIndex
                Me.BgwProduct.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub cellEndEditGroupLevel(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim frmFindGroup As New Core_Module.FrmCariMasterProductGroup
        Dim frmFindCategory As New Core_Module.frmCariMasterCategory
        Dim frmFindProduct As New Core_Module.frmCariMasterProduct
        Try
            Select Case Me.dgvGroup.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value
                Case ProdGroup
                    frmFindGroup.txtCari.Text = Me.dgvGroup.Item(Me.colProduct_Code.Index, e.RowIndex).Value
                    frmFindGroup.ProductGroupLevel = 3
                    frmFindGroup.ShowDialog()
                    If frmFindGroup.PrimaryProductGroup > 0 Then
                        Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                        Me.dgvGroup.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Me.dgvGroup.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Me.index = e.RowIndex
                        Me.BgwProduct.RunWorkerAsync()
                    End If
                Case ProdBrand
                    frmFindGroup.txtCari.Text = Me.dgvGroup.Item(Me.colProduct_Code.Index, e.RowIndex).Value
                    frmFindGroup.ProductGroupLevel = 2
                    frmFindGroup.ShowDialog()
                    If frmFindGroup.PrimaryProductGroup > 0 Then
                        Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                        Me.dgvGroup.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Me.dgvGroup.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.index = e.RowIndex
                        Me.BgwProduct.RunWorkerAsync()
                    End If
                Case ProdSubBrand
                    frmFindGroup.txtCari.Text = Me.dgvGroup.Item(Me.colProduct_Code.Index, e.RowIndex).Value
                    frmFindGroup.ProductGroupLevel = 1
                    frmFindGroup.ShowDialog()
                    If frmFindGroup.PrimaryProductGroup > 0 Then
                        Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                        Me.dgvGroup.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Me.dgvGroup.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Me.index = e.RowIndex
                        Me.BgwProduct.RunWorkerAsync()
                    End If
                Case ProdCategory
                    frmFindCategory.txtCari.Text = Me.dgvGroup.Item(Me.colProduct_Code.Index, e.RowIndex).Value
                    frmFindCategory.ShowDialog()
                    If frmFindCategory.PrimaryCategory > 0 Then
                        Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindCategory.PrimaryCategory
                        Me.dgvGroup.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objCategory.GetCategoryCode(Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Me.dgvGroup.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objCategory.GetDesc(Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                Case ProdProduct
                    frmFindProduct.txtCari.Text = Me.dgvGroup.Item(Me.colProduct_Code.Index, e.RowIndex).Value
                    frmFindProduct.ShowDialog()
                    If frmFindProduct.PrimaryProduct > 0 Then
                        Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindProduct.PrimaryProduct
                        Me.dgvGroup.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objProduct.GetProductCode(Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Me.dgvGroup.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objProduct.GetDescription(Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
            End Select
            If Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value <> -1 Then
                If Me.getIndexProduct(e.RowIndex, Me.dgvGroup.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value, Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value) <> -1 Then
                    MsgBox(Me.dgvGroup.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value & " " & Me.dgvGroup.Item(Me.colProduct_Code.Index, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                    Me.index = e.RowIndex
                    Me.BgwProduct.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cellEndEditCode(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim frmFindGroup As New Core_Module.FrmCariMasterProductGroup
        Dim frmFindCategory As New Core_Module.frmCariMasterCategory
        Dim frmFindProduct As New Core_Module.frmCariMasterProduct
        Try
            Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
            Select Case Me.dgvGroup.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value
                Case ProdGroup
                    If Me.objGroup.ApakahSudahAda(Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value, 3) = True Then
                        Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = objGroup.GetPrimary(Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value, 3)
                        Me.dgvGroup.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindGroup.txtCari.Text = Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindGroup.ProductGroupLevel = 3
                        frmFindGroup.ShowDialog()
                        If frmFindGroup.PrimaryProductGroup > 0 Then
                            Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                            Me.dgvGroup.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                            Me.dgvGroup.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.index = e.RowIndex
                            Me.BgwProduct.RunWorkerAsync()
                        End If
                    End If
                Case ProdBrand
                    If Me.objGroup.ApakahSudahAda(Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value, 2) = True Then
                        Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = objGroup.GetPrimary(Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                        Me.dgvGroup.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindGroup.txtCari.Text = Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindGroup.ProductGroupLevel = 2
                        frmFindGroup.ShowDialog()
                        If frmFindGroup.PrimaryProductGroup > 0 Then
                            Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                            Me.dgvGroup.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                            Me.dgvGroup.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.index = e.RowIndex
                            Me.BgwProduct.RunWorkerAsync()
                        End If
                    End If
                Case ProdSubBrand
                    If Me.objGroup.ApakahSudahAda(Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value, 1) = True Then
                        Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = objGroup.GetPrimary(Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value, 1)
                        Me.dgvGroup.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindGroup.txtCari.Text = Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindGroup.ProductGroupLevel = 1
                        frmFindGroup.ShowDialog()
                        If frmFindGroup.PrimaryProductGroup > 0 Then
                            Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                            Me.dgvGroup.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                            Me.dgvGroup.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.index = e.RowIndex
                            Me.BgwProduct.RunWorkerAsync()
                        End If
                    End If
                Case ProdCategory
                    If Me.objCategory.ApakahSudahAda(Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = objCategory.GetPrimary(Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvGroup.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objCategory.GetBrandDescription(Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindCategory.txtCari.Text = Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindCategory.ShowDialog()
                        If frmFindCategory.PrimaryCategory > 0 Then
                            Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindCategory.PrimaryCategory
                            Me.dgvGroup.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objCategory.GetCategoryCode(Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                            Me.dgvGroup.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objCategory.GetDesc(Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.index = e.RowIndex
                            Me.BgwProduct.RunWorkerAsync()
                        End If
                    End If
                Case ProdProduct
                    If Me.objProduct.ApakahSudahAda(Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = objProduct.GetPrimary(Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvGroup.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objProduct.GetDescription(Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindProduct.txtCari.Text = Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindProduct.ShowDialog()
                        If frmFindProduct.PrimaryProduct > 0 Then
                            Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindProduct.PrimaryProduct
                            Me.dgvGroup.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objProduct.GetProductCode(Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                            Me.dgvGroup.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objProduct.GetDescription(Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.index = e.RowIndex
                            Me.BgwProduct.RunWorkerAsync()
                        End If
                    End If
            End Select
            If Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value <> -1 Then
                If Me.getIndexProduct(e.RowIndex, Me.dgvGroup.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value, Me.dgvGroup.Item(Me.colProduct_Primary.Index, e.RowIndex).Value) <> -1 Then
                    MsgBox(Me.dgvGroup.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value & " " & Me.dgvGroup.Item(Me.colProduct_Code.Index, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                    Me.index = e.RowIndex
                    Me.BgwProduct.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function getIndexProduct(ByVal vRowException As Integer, ByVal vGroupLevel As String, ByVal vPrimary As Double) As Integer
        Try
            For Each dr As DataGridViewRow In Me.dgvGroup.Rows
                With dr
                    If Me.dgvGroup.Rows.IndexOf(dr) <> vRowException Then
                        If dr.Cells(Me.colProduct_Primary.Index).Value <> -1 Then
                            If dr.Cells(Me.colProduct_GroupLevel.Index).Value = vGroupLevel And dr.Cells(Me.colProduct_Primary.Index).Value = vPrimary Then
                                Return Me.dgvGroup.Rows.IndexOf(dr)
                            End If
                        End If
                    End If
                End With
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvProduct_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvGroup.DataError

    End Sub

    Private Sub dgvProduct_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvGroup.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.dgvGroup.Rows(Me.dgvGroup.Rows.Count - 1)
            dr.Cells(Me.colProduct_No.Index).Value = Me.dgvGroup.Rows.Count
            dr.Cells(Me.colProduct_GroupLevel.Index).Value = ProdGroup
            dr.Cells(Me.colProduct_Code.Index).Value = ""
            dr.Cells(Me.colProduct_Desc.Index).Value = ""
            dr.Cells(Me.colProduct_Primary.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProduct_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvGroup.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.dgvGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProduct_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvGroup.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Dim DTA As New dseTransaksiStockInOut_SetAmountTableAdapters.DataTable2TableAdapter
        Dim TBL As New dseTransaksiStockInOut_SetAmount.DataTable2DataTable
        Dim row As dseTransaksiStockInOut_SetAmount.DataTable2Row
        Dim rowDGV As DataGridViewRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'load data
            DTA.Fill(TBL)
            For Each row In TBL.Rows
                row.Unit_Price = Math.Round(row.Amount / row.Qty, Me.DefPrec, MidpointRounding.AwayFromZero)
            Next
            'filter data
            Me.FilterData(TBL)

            'tampilkan data
            TBL.DefaultView.RowFilter = ""
            TBL.DefaultView.Sort = "[Group Code], [Brand Code], [Sub Brand Code], [Category Code], [Product Code]"
            Me.dgvProd.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvProd.Rows.Clear()
            For Each DV As Data.DataRowView In TBL.DefaultView
                row = DV.Row
                Me.dgvProd.Rows.Add()
                rowDGV = Me.dgvProd.Rows(Me.dgvProd.Rows.Count - 1)

                rowDGV.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(row.Amount, Me.DefPrec)
                rowDGV.Cells(Me.ColBrandCode.Index).Value = row.Brand_Code
                rowDGV.Cells(Me.ColBrandDesc.Index).Value = row.Brand_Desc
                rowDGV.Cells(Me.ColCategoryCode.Index).Value = row.Category_Code
                rowDGV.Cells(Me.ColCategoryDesc.Index).Value = row.Category_Desc
                rowDGV.Cells(Me.ColCodeNo.Index).Value = row.Product_Code
                rowDGV.Cells(Me.ColDesc.Index).Value = row.Description
                rowDGV.Cells(Me.ColNo.Index).Value = rowDGV.Index + 1
                rowDGV.Cells(Me.ColPrimaryProd.Index).Value = row.Primary_Product
                rowDGV.Cells(Me.ColProductGroupCode.Index).Value = row.Group_Code
                rowDGV.Cells(Me.ColProductGroupDesc.Index).Value = row.Group_Desc
                rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(row.Qty, row._Decimal)
                rowDGV.Cells(Me.ColSubBrandCode.Index).Value = row.Sub_Brand_Code
                rowDGV.Cells(Me.ColSubBrandDesc.Index).Value = row.Sub_Brand_Desc
                rowDGV.Cells(Me.ColUnit.Index).Value = row.Unit_Code
                rowDGV.Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(row.Unit_Price, Me.DefPrec)
                rowDGV.Cells(Me.ColDetail.Index).Value = "Show"
            Next
            Me.dgvProd.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub FilterData(ByVal vTBL As dseTransaksiStockInOut_SetAmount.DataTable2DataTable)
        Dim rowDGV As DataGridViewRow
        Dim row As dseTransaksiStockInOut_SetAmount.DataTable2Row
        Dim TBLDummy As New dseTransaksiStockInOut_SetAmount.DataTable2DataTable
        Dim Level As EnumGroupLevel = EnumGroupLevel.LvlGroup
        Dim Filter As String = ""
        Dim DV As Data.DataRowView

        Try
            TBLDummy.Rows.Clear()

            'filter berdasarkan group
            Me.dgvGroup.AllowUserToAddRows = False
            If Me.dgvGroup.RowCount <> 0 Then
                For Each rowDGV In Me.dgvGroup.Rows
                    Level = Me.GetGroupLevel_Enum(rowDGV.Cells(Me.colProduct_GroupLevel.Index).Value)
                    If Level = EnumGroupLevel.LvlGroup Then
                        Filter = "[Primary Group] = " + rowDGV.Cells(Me.colProduct_Primary.Index).Value.ToString
                    ElseIf Level = EnumGroupLevel.LvlBrand Then
                        Filter = "[Primary Brand] = " + rowDGV.Cells(Me.colProduct_Primary.Index).Value.ToString
                    ElseIf Level = EnumGroupLevel.LvlSubBrand Then
                        Filter = "[Primary Sub Brand] = " + rowDGV.Cells(Me.colProduct_Primary.Index).Value.ToString
                    ElseIf Level = EnumGroupLevel.LvlCategory Then
                        Filter = "[Primary Category] = " + rowDGV.Cells(Me.colProduct_Primary.Index).Value.ToString
                    ElseIf Level = EnumGroupLevel.LvlProduct Then
                        Filter = "[Primary Product] = " + rowDGV.Cells(Me.colProduct_Primary.Index).Value.ToString
                    End If

                    vTBL.DefaultView.RowFilter = Filter
                    While vTBL.DefaultView.Count > 0
                        row = vTBL.DefaultView(0).Row
                        Me.CopyRow(TBLDummy, row)
                        vTBL.Rows.Remove(row)
                    End While
                Next

                vTBL.Rows.Clear()
                Me.CopyTabel(vTBL, TBLDummy)
            End If

            'filter berdasarkan unit price
            Filter = ""
            TBLDummy.Rows.Clear()
            If CDec(Me.TxtUnitPriceMin.Text) <> 0 Then
                Filter = "[Unit Price] >= " + Me.TxtUnitPriceMin.Text.Replace(",", "")
            End If
            If CDec(Me.TxtUnitPriceMax.Text) <> 0 Then
                If Filter = "" Then
                    Filter = "[Unit Price] <= " + Me.TxtUnitPriceMax.Text.Replace(",", "")
                Else
                    Filter = Filter + " AND [Unit Price] <= " + Me.TxtUnitPriceMax.Text.Replace(",", "")
                End If
            End If
            If Filter <> "" Then
                vTBL.DefaultView.RowFilter = Filter
                While vTBL.DefaultView.Count > 0
                    row = vTBL.DefaultView(0).Row
                    Me.CopyRow(TBLDummy, row)
                    vTBL.Rows.Remove(row)
                End While

                vTBL.Rows.Clear()
                Me.CopyTabel(vTBL, TBLDummy)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLDummy.Dispose()
            Me.dgvGroup.AllowUserToAddRows = True
        End Try
    End Sub

    Private Sub CopyTabel(ByVal vTBLTarget As dseTransaksiStockInOut_SetAmount.DataTable2DataTable, ByVal vTBLSumber As dseTransaksiStockInOut_SetAmount.DataTable2DataTable)
        Dim rowTarget, rowSumber As dseTransaksiStockInOut_SetAmount.DataTable2Row

        Try
            vTBLTarget.Rows.Clear()

            For Each rowSumber In vTBLSumber.Rows
                rowTarget = vTBLTarget.NewDataTable2Row
                For ij As Integer = 0 To vTBLSumber.Columns.Count - 1
                    rowTarget.Item(ij) = rowSumber.Item(ij)
                Next
                vTBLTarget.Rows.Add(rowTarget)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CopyRow(ByVal vTBLTarget As dseTransaksiStockInOut_SetAmount.DataTable2DataTable, ByVal vRowSumber As dseTransaksiStockInOut_SetAmount.DataTable2Row)
        Dim NewRow As dseTransaksiStockInOut_SetAmount.DataTable2Row

        Try
            NewRow = vTBLTarget.NewDataTable2Row
            For ij As Integer = 0 To vTBLTarget.Columns.Count - 1
                NewRow.Item(ij) = vRowSumber.Item(ij)
            Next
            vTBLTarget.Rows.Add(NewRow)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwClose_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwClose.DoWork
        Try
            Threading.Thread.Sleep(500)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwClose_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwClose.RunWorkerCompleted
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Try
            If MsgBox("Start data reset? Warning, this process is irreversible.", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.bgwReset.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwReset_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwReset.DoWork
        Dim UnitPrice As Decimal = 100
        Dim DTA As New dseTransaksiStockInOut_SetAmountTableAdapters.DataTable1TableAdapter
        Dim TBL As New dseTransaksiStockInOut_SetAmount.DataTable1DataTable
        Dim row As dseTransaksiStockInOut_SetAmount.DataTable1Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'load dari db
            DTA.Fill(TBL)

            Me.prg1.Maximum = TBL.Rows.Count
            For Each row In TBL.Rows
                Me.prg1.Value += 1
                Me.lblProgess.Text = "Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString + " - " + row.Product_Code
                DTA.UpdateQuery_SetAmount(Math.Round(row.Qty * UnitPrice, Me.DefPrec, MidpointRounding.AwayFromZero), UnitPrice, row.Primary)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub bgwReset_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwReset.RunWorkerCompleted
        Try
            Me.lblProgess.Text = ""
            Me.prg1.Value = 0
            MsgBox("Data updated", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkProductGroup_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkProductGroup.CheckedChanged
        Try
            Me.dgvProd.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            If Me.chkProductGroup.Checked = True Then
                Me.ColProductGroupCode.Visible = True
                Me.ColProductGroupDesc.Visible = True
            Else
                Me.ColProductGroupCode.Visible = False
                Me.ColProductGroupDesc.Visible = False
            End If
            Me.dgvProd.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkShowBrand_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowBrand.CheckedChanged
        Try
            Me.dgvProd.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            If Me.chkShowBrand.Checked = True Then
                Me.ColBrandCode.Visible = True
                Me.ColBrandDesc.Visible = True
            Else
                Me.ColBrandCode.Visible = False
                Me.ColBrandDesc.Visible = False
            End If
            Me.dgvProd.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkShowSubBrand_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowSubBrand.CheckedChanged
        Try
            Me.dgvProd.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            If Me.chkShowSubBrand.Checked = True Then
                Me.ColSubBrandCode.Visible = True
                Me.ColSubBrandDesc.Visible = True
            Else
                Me.ColSubBrandCode.Visible = False
                Me.ColSubBrandDesc.Visible = False
            End If
            Me.dgvProd.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkShowCategory_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowCategory.CheckedChanged
        Try
            Me.dgvProd.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            If Me.chkShowCategory.Checked = True Then
                Me.ColCategoryCode.Visible = True
                Me.ColCategoryDesc.Visible = True
            Else
                Me.ColCategoryCode.Visible = False
                Me.ColCategoryDesc.Visible = False
            End If
            Me.dgvProd.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPriceMin.GotFocus, TxtUnitPriceMax.GotFocus
        Dim Tx As Windows.Forms.TextBox

        Try
            Tx = TryCast(sender, Windows.Forms.TextBox)
            If IsNothing(Tx) = False Then
                Tx.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPriceMin_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPriceMin.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtUnitPriceMin, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPriceMin_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUnitPriceMin.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtUnitPriceMin, sender, e, 0)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPriceMin_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUnitPriceMin.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtUnitPriceMin, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPriceMin_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPriceMin.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtUnitPriceMin, sender, e, Me.DefPrec, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPriceMin.LostFocus, TxtUnitPriceMax.LostFocus
        Dim Tx As Windows.Forms.TextBox

        Try
            Tx = TryCast(sender, Windows.Forms.TextBox)
            If IsNothing(Tx) = False Then
                Tx.BackColor = Me.TxtRef.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPriceMax_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPriceMax.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtUnitPriceMax, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPriceMax_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUnitPriceMax.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtUnitPriceMax, sender, e, 0)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPriceMax_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUnitPriceMax.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtUnitPriceMax, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPriceMax_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPriceMax.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtUnitPriceMax, sender, e, Me.DefPrec, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProd_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvProd.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvProd)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSet.Click
        Dim i As New FrmTransaksiStockInAdj_SetAmount_Multiple
        Dim RowDgv As DataGridViewRow
        Try
            i.BackColorAll = Me.BackColorAll
            i.DefPrec = Me.DefPrec
            If i.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.Cursor = Cursors.WaitCursor
                Dim count As Integer = 0
                For Each RowDgv In Me.dgvProd.Rows
                    If RowDgv.Cells(Me.ColModifiy.Index).Value = Boolean.TrueString Then
                        count += 1
                        RowDgv.Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(i.TxtUnitPrice.Text, Me.DefPrec, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                        RowDgv.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(CDec(RowDgv.Cells(Me.ColUnitPrice.Index).Value) * CDec(RowDgv.Cells(Me.ColQty.Index).Value), Me.DefPrec, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                        Me.UpdateUnitPrice(RowDgv.Cells(Me.ColPrimaryProd.Index).Value, RowDgv.Cells(Me.ColAmount.Index).Value, RowDgv.Cells(Me.ColUnitPrice.Index).Value)
                    End If
                Next
                If count > 0 Then
                    MsgBox(count.ToString & " product updated.", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub 

    Private Sub LlbCheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbCheckAll.LinkClicked
        Dim RowDgv As DataGridViewRow
        Try
            Me.dgvProd.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each RowDgv In Me.dgvProd.Rows
                RowDgv.Cells(Me.ColModifiy.Index).Value = Boolean.TrueString
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvProd.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        End Try
    End Sub

    Private Sub LlbUncheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbUncheckAll.LinkClicked
        Dim RowDgv As DataGridViewRow
        Try
            Me.dgvProd.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each RowDgv In Me.dgvProd.Rows
                RowDgv.Cells(Me.ColModifiy.Index).Value = Boolean.FalseString
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvProd.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        End Try
    End Sub
End Class