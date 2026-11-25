Option Explicit On
Option Strict Off
Public Class ClsMasterCommission_Allocate
    Dim TblProduct As New DseMasterCommission_Allocate.tblproductDataTable
    Dim TblProduct1 As New DseMasterCommission_Allocate.tbltrans_salesinvoice_itemDataTable
    Dim TblDummy As New DseMasterCommission_Allocate.tbldummy_allocateDataTable
    dim RowDummy as DseMasterCommission_Allocate .tbldummy_allocateRow 

    Public Enum ProductLevel As Integer
        ProductGroup
        Brand
        SubBrand
        Category
        Product
    End Enum

    Public Function GetProductLevelStr(ByVal vProductGroupLevel As ClsMasterCommission_Allocate.ProductLevel) As String
        Try
            If vProductGroupLevel = ProductLevel.ProductGroup Then
                Return "Product Group"
            ElseIf vProductGroupLevel = ProductLevel.Brand Then
                Return "Brand"
            ElseIf vProductGroupLevel = ProductLevel.SubBrand Then
                Return "Sub Brand"
            ElseIf vProductGroupLevel = ProductLevel.Category Then
                Return "Category"
            Else
                Return "Product"
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetProductLevelStr(ByVal vProductGroupLevelStr As String) As ClsMasterCommission_Allocate.ProductLevel
        Try
            If vProductGroupLevelStr = "Product Group" Then
                Return ProductLevel.ProductGroup
            ElseIf vProductGroupLevelStr = "Brand" Then
                Return ProductLevel.Brand
            ElseIf vProductGroupLevelStr = "Sub Brand" Then
                Return ProductLevel.SubBrand
            ElseIf vProductGroupLevelStr = "Category" Then
                Return ProductLevel.Category
            Else
                Return ProductLevel.Product
            End If

            Return ProductLevel.Product
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ProductLevel.Product
        End Try
    End Function

    Public Function GetProductLevel(ByVal vProductGroupLevelStr As String) As ClsMasterCommission_Allocate.ProductLevel
        Try
            If vProductGroupLevelStr = "Product Group" Then
                Return ClsMasterCommission_Allocate.ProductLevel.ProductGroup
            ElseIf vProductGroupLevelStr = "Brand" Then
                Return ClsMasterCommission_Allocate.ProductLevel.Brand
            ElseIf vProductGroupLevelStr = "Sub Brand" Then
                Return ClsMasterCommission_Allocate.ProductLevel.SubBrand
            ElseIf vProductGroupLevelStr = "Category" Then
                Return ClsMasterCommission_Allocate.ProductLevel.Category
            Else
                Return ClsMasterCommission_Allocate.ProductLevel.Product
            End If

            Return ClsMasterCommission_Allocate.ProductLevel.Product
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ClsMasterCommission_Allocate.ProductLevel.Product
        End Try
    End Function

    Public Function Get_TblProduct(ByVal vFrm As FrmMasterCommission) As Boolean
        Dim dta As New DseMasterCommission_AllocateTableAdapters.tblproductTableAdapter
        Dim vRowProduct As DseMasterCommission_Allocate.tblproductRow
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring

            If vFrm.chkCommission.Checked = False And vFrm.chkProduct.Checked = False And vFrm.chkSalesPeriod.Checked = False Then
                dta.Fill(TblProduct)

                If TblProduct.Rows.Count > 0 Then
                    For Each vRowProduct In TblProduct.Rows
                        RowDummy = TblDummy.Newtbldummy_allocateRow
                        With RowDummy
                            .Primary_Product = vRowProduct.Primary
                            .Product_Code = vRowProduct.Product_Code
                            .Description = vRowProduct.Description
                        End With
                        TblDummy.Addtbldummy_allocateRow(RowDummy)
                    Next
                End If

                Me.LoadData_IsExistAllocate()
            Else
                dta.Fill(TblProduct)
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function Get_FilterBy_ProductLevel(ByVal vFrm As FrmMasterCommission) As Boolean
        Dim dgvItem As DataGridView = vFrm.dgvItemProductGroup
        Dim rowDGV As New DataGridViewRow
        Dim vRowProduct As DseMasterCommission_Allocate.tblproductRow
        Try
            If vFrm.chkProduct.Checked = True Then
                dgvItem.AllowUserToAddRows = False
                If dgvItem.RowCount > 0 Then
                    For Each rowDGV In dgvItem.Rows
                        With rowDGV

                            If Me.GetProductLevel(.Cells(vFrm.ColProductGroup_Level.Index).Value) = ProductLevel.Product Then
                                TblProduct.DefaultView.RowFilter = Nothing
                                TblProduct.DefaultView.RowFilter = "[Primary]=" & .Cells(vFrm.ColProductGroup_PrimaryAcuan.Index).Value
                                For Each vr As DataRowView In TblProduct.DefaultView
                                    vRowProduct = vr.Row
                                    TblDummy.DefaultView.RowFilter = Nothing
                                    TblDummy.DefaultView.RowFilter = "[Primary Product]=" & vRowProduct.Primary
                                    If TblDummy.DefaultView.Count = 0 Then
                                        RowDummy = TblDummy.Newtbldummy_allocateRow
                                        With RowDummy
                                            .Primary_Product = vRowProduct.Primary
                                            .Product_Code = vRowProduct.Product_Code
                                            .Description = vRowProduct.Description
                                        End With
                                        TblDummy.Addtbldummy_allocateRow(RowDummy)
                                    End If
                                Next

                            ElseIf Me.GetProductLevel(.Cells(vFrm.ColProductGroup_Level.Index).Value) = ProductLevel.Category Then
                                TblProduct.DefaultView.RowFilter = Nothing
                                TblProduct.DefaultView.RowFilter = "[Primary Category]=" & .Cells(vFrm.ColProductGroup_PrimaryAcuan.Index).Value
                                For Each vr As DataRowView In TblProduct.DefaultView
                                    vRowProduct = vr.Row
                                    TblDummy.DefaultView.RowFilter = Nothing
                                    TblDummy.DefaultView.RowFilter = "[Primary Product]=" & vRowProduct.Primary
                                    If TblDummy.DefaultView.Count = 0 Then
                                        RowDummy = TblDummy.Newtbldummy_allocateRow
                                        With RowDummy
                                            .Primary_Product = vRowProduct.Primary
                                            .Product_Code = vRowProduct.Product_Code
                                            .Description = vRowProduct.Description
                                        End With
                                        TblDummy.Addtbldummy_allocateRow(RowDummy)
                                    End If
                                Next

                            ElseIf Me.GetProductLevel(.Cells(vFrm.ColProductGroup_Level.Index).Value) = ProductLevel.SubBrand Then
                                TblProduct.DefaultView.RowFilter = Nothing
                                TblProduct.DefaultView.RowFilter = "[Primary Sub Brand]=" & .Cells(vFrm.ColProductGroup_PrimaryAcuan.Index).Value
                                For Each vr As DataRowView In TblProduct.DefaultView
                                    vRowProduct = vr.Row
                                    TblDummy.DefaultView.RowFilter = Nothing
                                    TblDummy.DefaultView.RowFilter = "[Primary Product]=" & vRowProduct.Primary
                                    If TblDummy.DefaultView.Count = 0 Then
                                        RowDummy = TblDummy.Newtbldummy_allocateRow
                                        With RowDummy
                                            .Primary_Product = vRowProduct.Primary
                                            .Product_Code = vRowProduct.Product_Code
                                            .Description = vRowProduct.Description
                                        End With
                                        TblDummy.Addtbldummy_allocateRow(RowDummy)
                                    End If
                                Next

                            ElseIf Me.GetProductLevel(.Cells(vFrm.ColProductGroup_Level.Index).Value) = ProductLevel.Brand Then
                                TblProduct.DefaultView.RowFilter = Nothing
                                TblProduct.DefaultView.RowFilter = "[Primary Brand]=" & .Cells(vFrm.ColProductGroup_PrimaryAcuan.Index).Value
                                For Each vr As DataRowView In TblProduct.DefaultView
                                    vRowProduct = vr.Row
                                    TblDummy.DefaultView.RowFilter = Nothing
                                    TblDummy.DefaultView.RowFilter = "[Primary Product]=" & vRowProduct.Primary
                                    If TblDummy.DefaultView.Count = 0 Then
                                        RowDummy = TblDummy.Newtbldummy_allocateRow
                                        With RowDummy
                                            .Primary_Product = vRowProduct.Primary
                                            .Product_Code = vRowProduct.Product_Code
                                            .Description = vRowProduct.Description
                                        End With
                                        TblDummy.Addtbldummy_allocateRow(RowDummy)
                                    End If
                                Next

                            ElseIf Me.GetProductLevel(.Cells(vFrm.ColProductGroup_Level.Index).Value) = ProductLevel.ProductGroup Then
                                TblProduct.DefaultView.RowFilter = Nothing
                                TblProduct.DefaultView.RowFilter = "[Primary Product Group]=" & .Cells(vFrm.ColProductGroup_PrimaryAcuan.Index).Value
                                For Each vr As DataRowView In TblProduct.DefaultView
                                    vRowProduct = vr.Row
                                    TblDummy.DefaultView.RowFilter = Nothing
                                    TblDummy.DefaultView.RowFilter = "[Primary Product]=" & vRowProduct.Primary
                                    If TblDummy.DefaultView.Count = 0 Then
                                        RowDummy = TblDummy.Newtbldummy_allocateRow
                                        With RowDummy
                                            .Primary_Product = vRowProduct.Primary
                                            .Product_Code = vRowProduct.Product_Code
                                            .Description = vRowProduct.Description
                                        End With
                                        TblDummy.Addtbldummy_allocateRow(RowDummy)
                                    End If
                                Next
                            End If
                        End With
                    Next
                End If
                dgvItem.AllowUserToAddRows = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Get_FilterBy_ProductLevel1(ByVal vFrm As FrmMasterCommission) As Boolean
        Dim dgvItem As DataGridView = vFrm.dgvItemProductGroup
        Dim rowDGV As New DataGridViewRow
        Dim vRowProduct As DseMasterCommission_Allocate.tbltrans_salesinvoice_itemRow
        Try
            If vFrm.chkProduct.Checked = True Then
                dgvItem.AllowUserToAddRows = False
                If dgvItem.RowCount > 0 Then
                    For Each rowDGV In dgvItem.Rows
                        With rowDGV

                            If Me.GetProductLevel(.Cells(vFrm.ColProductGroup_Level.Index).Value) = ProductLevel.Product Then
                                TblProduct1.DefaultView.RowFilter = Nothing
                                TblProduct1.DefaultView.RowFilter = "[Primary Product]=" & .Cells(vFrm.ColProductGroup_PrimaryAcuan.Index).Value
                                For Each vr As DataRowView In TblProduct1.DefaultView
                                    vRowProduct = vr.Row
                                    TblDummy.DefaultView.RowFilter = Nothing
                                    TblDummy.DefaultView.RowFilter = "[Primary Product]=" & vRowProduct.Primary_Product
                                    If TblDummy.DefaultView.Count = 0 Then
                                        RowDummy = TblDummy.Newtbldummy_allocateRow
                                        With RowDummy
                                            .Primary_Product = vRowProduct.Primary_Product
                                            .Product_Code = vRowProduct.Product_Code
                                            .Description = vRowProduct.Description
                                        End With
                                        TblDummy.Addtbldummy_allocateRow(RowDummy)
                                    End If
                                Next

                            ElseIf Me.GetProductLevel(.Cells(vFrm.ColProductGroup_Level.Index).Value) = ProductLevel.Category Then
                                TblProduct1.DefaultView.RowFilter = Nothing
                                TblProduct1.DefaultView.RowFilter = "[Primary Category]=" & .Cells(vFrm.ColProductGroup_PrimaryAcuan.Index).Value
                                For Each vr As DataRowView In TblProduct1.DefaultView
                                    vRowProduct = vr.Row
                                    TblDummy.DefaultView.RowFilter = Nothing
                                    TblDummy.DefaultView.RowFilter = "[Primary Product]=" & vRowProduct.Primary_Product
                                    If TblDummy.DefaultView.Count = 0 Then
                                        RowDummy = TblDummy.Newtbldummy_allocateRow
                                        With RowDummy
                                            .Primary_Product = vRowProduct.Primary_Product
                                            .Product_Code = vRowProduct.Product_Code
                                            .Description = vRowProduct.Description
                                        End With
                                        TblDummy.Addtbldummy_allocateRow(RowDummy)
                                    End If
                                Next

                            ElseIf Me.GetProductLevel(.Cells(vFrm.ColProductGroup_Level.Index).Value) = ProductLevel.SubBrand Then
                                TblProduct1.DefaultView.RowFilter = Nothing
                                TblProduct1.DefaultView.RowFilter = "[Primary Sub Brand]=" & .Cells(vFrm.ColProductGroup_PrimaryAcuan.Index).Value
                                For Each vr As DataRowView In TblProduct1.DefaultView
                                    vRowProduct = vr.Row
                                    TblDummy.DefaultView.RowFilter = Nothing
                                    TblDummy.DefaultView.RowFilter = "[Primary Product]=" & vRowProduct.Primary_Product
                                    If TblDummy.DefaultView.Count = 0 Then
                                        RowDummy = TblDummy.Newtbldummy_allocateRow
                                        With RowDummy
                                            .Primary_Product = vRowProduct.Primary_Product
                                            .Product_Code = vRowProduct.Product_Code
                                            .Description = vRowProduct.Description
                                        End With
                                        TblDummy.Addtbldummy_allocateRow(RowDummy)
                                    End If
                                Next

                            ElseIf Me.GetProductLevel(.Cells(vFrm.ColProductGroup_Level.Index).Value) = ProductLevel.Brand Then
                                TblProduct1.DefaultView.RowFilter = Nothing
                                TblProduct1.DefaultView.RowFilter = "[Primary Brand]=" & .Cells(vFrm.ColProductGroup_PrimaryAcuan.Index).Value
                                For Each vr As DataRowView In TblProduct1.DefaultView
                                    vRowProduct = vr.Row
                                    TblDummy.DefaultView.RowFilter = Nothing
                                    TblDummy.DefaultView.RowFilter = "[Primary Product]=" & vRowProduct.Primary_Product
                                    If TblDummy.DefaultView.Count = 0 Then
                                        RowDummy = TblDummy.Newtbldummy_allocateRow
                                        With RowDummy
                                            .Primary_Product = vRowProduct.Primary_Product
                                            .Product_Code = vRowProduct.Product_Code
                                            .Description = vRowProduct.Description
                                        End With
                                        TblDummy.Addtbldummy_allocateRow(RowDummy)
                                    End If
                                Next

                            ElseIf Me.GetProductLevel(.Cells(vFrm.ColProductGroup_Level.Index).Value) = ProductLevel.ProductGroup Then
                                TblProduct1.DefaultView.RowFilter = Nothing
                                TblProduct1.DefaultView.RowFilter = "[Primary Product Group]=" & .Cells(vFrm.ColProductGroup_PrimaryAcuan.Index).Value
                                For Each vr As DataRowView In TblProduct1.DefaultView
                                    vRowProduct = vr.Row
                                    TblDummy.DefaultView.RowFilter = Nothing
                                    TblDummy.DefaultView.RowFilter = "[Primary Product]=" & vRowProduct.Primary_Product
                                    If TblDummy.DefaultView.Count = 0 Then
                                        RowDummy = TblDummy.Newtbldummy_allocateRow
                                        With RowDummy
                                            .Primary_Product = vRowProduct.Primary_Product
                                            .Product_Code = vRowProduct.Product_Code
                                            .Description = vRowProduct.Description
                                        End With
                                        TblDummy.Addtbldummy_allocateRow(RowDummy)
                                    End If
                                Next
                            End If
                        End With
                    Next
                End If
                dgvItem.AllowUserToAddRows = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Get_FilterBy_SalesPeriod(ByVal vFromDate As Date, ByVal vEndDate As Date, ByVal vFrm As FrmMasterCommission) As Boolean
        Dim dta As New DseMasterCommission_AllocateTableAdapters.tbltrans_salesinvoice_itemTableAdapter
        Dim tbl As New DseMasterCommission_Allocate.tbltrans_salesinvoice_itemDataTable
        Dim row As DseMasterCommission_Allocate.tbltrans_salesinvoice_itemRow
        Dim vRowProduct As DseMasterCommission_Allocate.tblproductRow
        Try
            If vFrm.chkSalesPeriod.Checked = True Then
                dta.Connection.ConnectionString = Core_Module.GetConstring
                dta.Fill(TblProduct1, vEndDate, vFromDate)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function Get_FilterBy_SalesCommission(ByVal vFrm As FrmMasterCommission) As Boolean
        Dim dgvItem As DataGridView = vFrm.dgvFilterComm
        Dim rowDGV As New DataGridViewRow
        Try
            Me.LoadData_IsExistAllocate()

            If vFrm.chkCommission.Checked = True Then
                dgvItem.AllowUserToAddRows = False
                If dgvItem.RowCount > 0 Then
                    For Each rowDGV In dgvItem.Rows
                        With rowDGV
                            TblDummy.DefaultView.RowFilter = Nothing
                            TblDummy.DefaultView.RowFilter = "[Primary Commission]=" & .Cells(vFrm.ColFilter_PrimaryCommission.Index).Value
                            For Each vr As DataRowView In TblDummy.DefaultView
                                RowDummy = vr.Row
                                RowDummy.Set_Filter = 1
                            Next
                        End With
                    Next

                    TblDummy.DefaultView.RowFilter = Nothing
                    TblDummy.DefaultView.RowFilter = "[Set Filter]=0"
                    For Each vr As DataRowView In TblDummy.DefaultView
                        RowDummy = vr.Row
                        TblDummy.Rows.Remove(RowDummy)
                    Next
                    'For ij As Integer = TblDummy.Rows.Count - 1 To 0 Step -1
                    '    RowDummy = TblDummy.Rows(ij)
                    '    If RowDummy.Set_Filter = 0 Then
                    '        TblDummy.Rows.Remove(RowDummy)
                    '    End If
                    'Next
                End If
                dgvItem.AllowUserToAddRows = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function LoadData_IsExistAllocate() As Boolean
        Dim dta As New DseMasterCommission_AllocateTableAdapters.tblmaster_sales_commission_allocate1TableAdapter
        Dim tbl As New DseMasterCommission_Allocate.tblmaster_sales_commission_allocate1DataTable
        Dim row As DseMasterCommission_Allocate.tblmaster_sales_commission_allocate1Row
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.Fill(tbl)

            If TblDummy.Rows.Count > 0 Then

                For Each RowDummy In TblDummy.Rows
                    With RowDummy
                        tbl.DefaultView.RowFilter = Nothing
                        tbl.DefaultView.RowFilter = "[primary code no]=" & RowDummy.Primary_Product
                        If tbl.DefaultView.Count > 0 Then
                            row = tbl.DefaultView(0).Row
                            .Primary_Commission = row.Primary_Commission
                            .Commission_Code = row.Commission_Code
                            .Commission_Desc = row.Commission_Desc
                        End If
                    End With
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function IsiTblDummy_Sementara() As Boolean
        Dim vRowProduct As DseMasterCommission_Allocate.tblproductRow
        Try
            For Each vRowProduct In TblProduct
                With vRowProduct
                    RowDummy = TblDummy.Newtbldummy_allocateRow
                    RowDummy.Primary_Product = .Primary
                    RowDummy.Primary_Commission = -1
                    RowDummy.Product_Code = .Product_Code
                    RowDummy.Description = .Description
                    TblDummy.Addtbldummy_allocateRow(RowDummy)
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function IsiTblDummy_Sementara1() As Boolean
        Dim vRowProduct As DseMasterCommission_Allocate.tbltrans_salesinvoice_itemRow
        Try
            For Each vRowProduct In TblProduct1
                With vRowProduct
                    RowDummy = TblDummy.Newtbldummy_allocateRow
                    RowDummy.Primary_Product = .Primary_Product
                    RowDummy.Primary_Commission = -1
                    RowDummy.Product_Code = .Product_Code
                    RowDummy.Description = .Description
                    TblDummy.Addtbldummy_allocateRow(RowDummy)
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function LoadData_ToTblDummy(ByVal vFromDate As Date, ByVal vEndDate As Date, ByVal vFrm As FrmMasterCommission) As Boolean
        Try
            Me.TblDummy.Rows.Clear()
            If vFrm.chkSalesPeriod.Checked = False Then
                Me.Get_TblProduct(vFrm)
                If vFrm.chkProduct.Checked = True Then
                    Get_FilterBy_ProductLevel(vFrm)
                    'Else
                    '    IsiTblDummy_Sementara()
                End If
                Me.Get_FilterBy_SalesCommission(vFrm)

            Else
                Me.Get_FilterBy_SalesPeriod(vFromDate, vEndDate, vFrm)
                If vFrm.chkProduct.Checked = True Then
                    Get_FilterBy_ProductLevel1(vFrm)
                Else
                    IsiTblDummy_Sementara1()
                End If
                Me.Get_FilterBy_SalesCommission(vFrm)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function LoadData_ToDGV(ByVal vFromDate As Date, ByVal vEndDate As Date, ByVal vFrm As FrmMasterCommission, ByVal vObjCD As ClsCekData) As Boolean
        Dim dgvItem As DataGridView = vFrm.dgvAllocate
        Dim row As DseMasterCommission_Allocate.tbldummy_allocateRow
        Dim rowDGV As DataGridViewRow
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            dgvItem.Rows.Clear()
            Me.LoadData_ToTblDummy(vFromDate, vEndDate, vFrm)
            If TblDummy.Rows.Count > 0 Then
                TblDummy.DefaultView.RowFilter = Nothing
                TblDummy.DefaultView.Sort = "[Product Code] ASC"
                For Each vr As DataRowView In TblDummy.DefaultView
                    RowDummy = vr.Row
                    dgvItem.Rows.Add()
                    rowDGV = dgvItem.Rows(dgvItem.Rows.Count - 1)
                    With rowDGV
                        .Cells(vFrm.ColAllocate_PrimaryCodeNo.Index).Value = RowDummy.Primary_Product
                        .Cells(vFrm.ColAllocate_PrimaryCommision.Index).Value = RowDummy.Primary_Commission
                        .Cells(vFrm.ColAllocate_CodeNo.Index).Value = RowDummy.Product_Code
                        .Cells(vFrm.ColAllocate_Desc.Index).Value = RowDummy.Description
                        .Cells(vFrm.ColAllocate_CommisionnCode.Index).Value = RowDummy.Commission_Code
                    End With
                Next
            End If

            vObjCD.SetNumber(dgvItem)
            dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function SimpanItem(ByVal vRowIdx As Integer, ByVal vFrm As FrmMasterCommission) As Boolean
        Dim dta As New DseMasterCommission_AllocateTableAdapters.tblmaster_sales_commission_allocateTableAdapter
        Dim RowDgv As DataGridViewRow = vFrm.dgvAllocate.Rows(vRowIdx)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring

            Me.HapusItem(vRowIdx, vFrm)
            dta.Insert(vFrm.dgvAllocate.Item(vFrm.ColAllocate_PrimaryCodeNo.Index, vRowIdx).Value, vFrm.dgvAllocate.Item(vFrm.ColAllocate_PrimaryCommision.Index, vRowIdx).Value)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function HapusItem(ByVal vRowIdx As Integer, ByVal vFrm As FrmMasterCommission) As Boolean
        Dim dta As New DseMasterCommission_AllocateTableAdapters.tblmaster_sales_commission_allocateTableAdapter
        Dim RowDgv As DataGridViewRow = vFrm.dgvAllocate.Rows(vRowIdx)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.Delete(vFrm.dgvAllocate.Item(vFrm.ColAllocate_PrimaryCodeNo.Index, vRowIdx).Value)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GenerateQuery(ByVal vDGV As DataGridView, ByVal vFrm As FrmMasterCommission, ByVal vPrimaryCommission As Double) As Boolean
        Dim counter As Integer = 0
        Dim MaxCount As Integer = 5
        Dim vDeleteQuery As String = ""
        Dim vInsertQuery As String = ""
        Try
            'vDeleteQuery = "DELETE FROM tblmaster_sales_commission_allocate WHERE"
            vInsertQuery = "INSERT INTO tblmaster_sales_commission_allocate (`primary code no`, `primary commission`) VALUES"
            If vDGV.Rows.Count > 0 Then
                For Each row As DataGridViewRow In vDGV.Rows
                    With row
                        If .Cells(vFrm.ColAllocate_Set.Index).Value = True Then
                            If counter < MaxCount Then
                                vDeleteQuery &= "DELETE FROM tblmaster_sales_commission_allocate WHERE (`primary code no`= '" & row.Cells(vFrm.ColAllocate_PrimaryCodeNo.Index).Value & "');" & vbNewLine
                                vInsertQuery &= "('" & row.Cells(vFrm.ColAllocate_PrimaryCodeNo.Index).Value & "','" & vPrimaryCommission & "'),"
                                counter += 1
                            Else
                                vInsertQuery = vInsertQuery.Substring(0, vInsertQuery.Length - 1)
                                vInsertQuery &= ";" & vbNewLine
                                counter = 1
                                vInsertQuery &= "INSERT INTO tblmaster_sales_commission_allocate (`primary code no`, `primary commission`) VALUES"
                                vInsertQuery &= "('" & row.Cells(vFrm.ColAllocate_PrimaryCodeNo.Index).Value & "','" & vPrimaryCommission & "'),"
                                vDeleteQuery &= "DELETE FROM tblmaster_sales_commission_allocate WHERE (`primary code no`= '" & row.Cells(vFrm.ColAllocate_PrimaryCodeNo.Index).Value & "');" & vbNewLine
                            End If
                        End If
                    End With
                Next

                If counter <= MaxCount Then
                    vInsertQuery = vInsertQuery.Substring(0, vInsertQuery.Length - 1)
                    vInsertQuery &= ";"
                End If

                ExecuteQuery(vDeleteQuery, vInsertQuery)
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub ExecuteQuery(ByVal vDeleteQuery As String, ByVal vInsertQuery As String)
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(Core_Module.GetConstring)
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, vDeleteQuery)
        Try
            Conn.Open()
            script.Query = vDeleteQuery
            script.Execute()
            script.Query = vInsertQuery
            script.Execute()
            Conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function IsiDGV(ByVal vDGV As DataGridView, ByVal vFrm As FrmMasterCommission) As Boolean
        Dim dta As New DseMasterCommission_AllocateTableAdapters.tblmaster_sales_commission_allocate1TableAdapter
        Dim tbl As New DseMasterCommission_Allocate.tblmaster_sales_commission_allocate1DataTable
        Dim row As DseMasterCommission_Allocate.tblmaster_sales_commission_allocate1Row
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.Fill(tbl)
            For Each rowDgv As DataGridViewRow In vDGV.Rows
                With rowDgv
                    If .Cells(vFrm.ColAllocate_Set.Index).Value = True Then
                        tbl.DefaultView.RowFilter = Nothing
                        tbl.DefaultView.RowFilter = "[primary code no]=" & .Cells(vFrm.ColAllocate_PrimaryCodeNo.Index).Value
                        If tbl.DefaultView.Count > 0 Then
                            row = tbl.DefaultView(0).Row
                            '.Cells(vFrm.ColAllocate_Set.Index).Value = False
                            .Cells(vFrm.ColAllocate_PrimaryCommision.Index).Value = row.Primary_Commission
                            .Cells(vFrm.ColAllocate_CommisionnCode.Index).Value = row.Commission_Code
                        End If
                    End If
                End With
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            dta.Dispose()
        End Try
    End Function

End Class
