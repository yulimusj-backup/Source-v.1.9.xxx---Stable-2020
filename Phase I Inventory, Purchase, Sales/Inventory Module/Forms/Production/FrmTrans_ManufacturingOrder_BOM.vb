Option Explicit On
Option Strict Off

Public Class FrmTrans_ManufacturingOrder_BOM
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public vPrimaryProduct As Double = -1
    Public vPrimaryProductJO As Double = -1
    Public vPrimaryParentProduct As Double = -1
    Public vPrimaryMO As Double = -1
    Public vIdentifierJO As String = ""
    Public vIdentifierWO As String = ""
    'Public vIdentifierBOM As String = "-"
    Public vIdentifierBOM_Parent As String = "-"
    Public vProductCode As String = ""
    Public vDefPrec As Integer = 0
    Public vBolehEdit As Boolean = False
    Public vQtyTarget As Decimal = 0
    Public vLevel As Integer = 0
    Public UserCode As String = ""

    Public TblItemBOM_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable
    Public vTblItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable
    Public vTblProsesOrder As DseTrans_ManufacturingOrder.tbltrans_mo_processorderDataTable
    Public vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable
    Public vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable
    Public vTblItemWIP As DseTrans_ManufacturingOrder.tbltrans_mo_item_wipDataTable
    Dim TblBOM_Temp As New DseTrans_ManufacturingOrder.BOM_TempDataTable
    Dim RowBOM_Temp As DseTrans_ManufacturingOrder.BOM_TempRow

    Dim objUnit As New Core_Module.ClsMasterUnit
    Dim ObjMO As New ClsTrans_ManufacturingOrder
    Dim ObjMO_Item As New ClsTrans_ManufacturingOrder_Item
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjProses As New Core_Module.ClsMasterProses
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim IndexRemove As Integer = 0
    Dim Buffer As String = ""
    Dim Buff_PrimaryProduct As Double = -1
    Dim FromOtherButton As Boolean = False
    Dim CancelClose As Boolean = False

    Public Function ApakahDataBenar() As Boolean
        Try
            'product harus memiliki item bom, jadi minimal harus ada 1 baris item di DgvItemBOM
            If Me.vLevel = 0 Then
                If Me.dgvItemBOM.RowCount <= 1 Then
                    MsgBox("Please specify minimal 1 material", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'qty ref harus positive
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyRef.Text) = False Then
                MsgBox("Please enter a positif number for reference qty.", _
                MsgBoxStyle.Information, "Information")

                Me.TxtQtyRef.Focus()
                Return False
            End If

            'good precentage harus positive
            If Me.ObjCD.CekBilanganPositif(Me.TxtGoodPercentage.Text) = False Then
                MsgBox("Please enter a positif number for good percentage.", _
                MsgBoxStyle.Information, "Information")

                Me.TxtQtyRef.Focus()
                Return False
            End If

            'kalo ada processnya, maka harus isikan minimal 1 raw material
            If Me.txtPrimaryProses.Text <> "" And Me.txtPrimaryProses.Text <> "-1" Then
                If Me.dgvItemBOM.Rows.Count <= 1 Then
                    MsgBox("Please specify minimal 1 material", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'semua item bill of material harus goods
            For Each rowdgv As DataGridViewRow In Me.dgvItemBOM.Rows
                If rowdgv.Cells(Me.ColBOM_PrimaryProduct.Index).Value <> -1 Then
                    If Me.ObjProduct.ApakahGoods(Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, rowdgv.Index).Value) = False Then
                        MsgBox("Each code no in bill of material must be goods", MsgBoxStyle.Information, "Information")

                        Me.dgvItemBOM.Focus()
                        Me.dgvItemBOM.ClearSelection()
                        rowdgv.Selected = True

                        Return False
                    End If
                End If
            Next

            ''@aul 18/09/2015
            Dim ListBoM As New List(Of Double)
            Dim ListChild As New List(Of Double)
            If Me.dgvItemBOM.RowCount > 1 Then
                For Each rowDgv As DataGridViewRow In Me.dgvItemBOM.Rows
                    If IsNothing(rowDgv.Cells(Me.ColBOM_PrimaryProduct.Index).Value) = False Then
                        If rowDgv.Cells(Me.ColBOM_PrimaryProduct.Index).Value <> "-1" Then
                            If ListBoM.Contains(rowDgv.Cells(Me.ColBOM_PrimaryProduct.Index).Value.ToString) = False Then
                                ListBoM.Add(CDbl(rowDgv.Cells(Me.ColBOM_PrimaryProduct.Index).Value.ToString))
                            End If
                        End If
                    End If
                Next
            End If
            If ListBoM.Count > 0 Then
                GetChild(Me.vPrimaryProduct, ListChild)
                If ListChild.Count > 0 Then
                    Dim e As String = ""
                    For Each d As Double In ListBoM
                        If ListChild.Contains(d) Then
                            e = ObjProduct.GetProductCode(d)
                            MsgBox("Invalid bill of material data, the current data will cause infinitely long bill of material." & vbCrLf & "Code No : " & e, MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    Next
                End If
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Function ApakahSudahAdaBOM(ByVal vPrimaryProduct As Double, ByVal vPrimaryRM As Double) As Boolean

        Try
            Me.TblItemBOM_.DefaultView.RowFilter = "[Identifier JO] = '" & Me.vIdentifierJO.ToString & _
            "' AND [Primary Product] = " & vPrimaryProduct.ToString & " AND [Primary Product RM] = " & vPrimaryRM
            If Me.TblItemBOM_.DefaultView.Count > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetIdentifierBOM(ByVal vPrimaryProduct As Double, ByVal vPrimaryRM As Double) As String
        Dim rowItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
        Dim IdentifierBOM As String = ""
        Try
            Me.TblItemBOM_.DefaultView.RowFilter = "[Identifier JO] = '" & Me.vIdentifierJO.ToString & _
            "' AND [Primary Product] = " & vPrimaryProduct.ToString & " AND [Primary Product RM] = " & vPrimaryRM
            If Me.TblItemBOM_.Count > 0 Then
                rowItemBOM = Me.TblItemBOM_.DefaultView(0).Row
                IdentifierBOM = rowItemBOM.Identifier_BOM
            End If

            Return IdentifierBOM
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub GetChild(ByVal vPrimaryProduct As Double, ByVal vListChild As List(Of Double))
        Dim DtaBoM As New Core_Module.dseMasterProduct_BOMTableAdapters.tblmaster_product_item_bomTableAdapter
        Dim TblBoM As New Core_Module.dseMasterProduct_BOM.tblmaster_product_item_bomDataTable
        Dim RowBoM As Core_Module.dseMasterProduct_BOM.tblmaster_product_item_bomRow

        Try
            'DtaBoM.Connection.ConnectionString = Core_Module.GetConstring
            DtaBoM.FillBy_PrimaryParent(TblBoM, vPrimaryProduct)

            If vListChild.Contains(vPrimaryProduct) = False Then
                vListChild.Add(vPrimaryProduct)
            End If
            For Each RowBoM In TblBoM
                If vListChild.Contains(RowBoM.primary_product) = False Then
                    vListChild.Add(RowBoM.primary_product)
                End If
                Me.GetChild(RowBoM.primary_product, vListChild)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaBoM.Dispose()
            TblBoM.Dispose()
        End Try
    End Sub

    Private Sub BtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        Dim hasil As MsgBoxResult

        Try
            Me.FromOtherButton = False
            If Me.ApakahDataBenar() = True Then
                hasil = MsgBox("Close form and accept data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    Dim vFrm As New FrmTrans_ManufacturingOrder
                    Me.UpdateAllTblItem()
                    vFrm.Dispose()
                    Me.DialogResult = Windows.Forms.DialogResult.OK
                    Me.FromOtherButton = True
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Dim hasil As MsgBoxResult

        Try
            Me.FromOtherButton = False
            If Me.vBolehEdit = True Then
                hasil = MsgBox("Close form and discard data changes?", MsgBoxStyle.YesNo, "Confirmation")
                If hasil = MsgBoxResult.Yes Then
                    If Me.TblBOM_Temp.Rows.Count > 0 Then
                        For Each Me.RowBOM_Temp In Me.TblBOM_Temp.Rows
                            Me.DeleteTblItemBOM(Me.RowBOM_Temp.Primary_Parent_Product, Me.RowBOM_Temp.Primary_Product, Me.RowBOM_Temp.Identifier_BOM_Parent, Me.TblItemBOM_)
                            'Me.TblBOM_Temp.Rows.Remove(Me.RowBOM_Temp)
                        Next
                    End If
                    Me.FromOtherButton = True
                    Me.Close()
                Else
                    Me.FromOtherButton = True
                    Me.CancelClose = True
                    Exit Sub
                End If
            Else
                Me.Close()
            End If
            
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub UpdateAllTblItem()

        Try
            If vLevel <> 0 Then
                'update tabel item bom
                Me.UpdateTblItemBOM(Me.TblItemBOM_)
            Else
                Me.vTblItemBOM.Rows.Clear()
                Me.UpdateTblItemBOM(Me.TblItemBOM_)
                Me.vTblItemBOM = Me.TblItemBOM_.Copy

                'update tabel item proses order
                Me.UpdateTblProsesOrder_BOM()

                'update hitungan qty di tbl item wo
                Dim f As New FrmTrans_ManufacturingOrder
                f.IsiTblItemWORM(-1, vPrimaryProduct, vQtyTarget, Me.vTblItemWO, Me.vTblItemRM, Me.vTblItemBOM, vIdentifierJO, "-", "-", Me.vTblItemWIP, #1/1/1950#, False)

                'hapus wo dan rm yang product di tbl bomnya sudah tidak ada
                Me.DeleteTblItemWORM(Me.vTblItemBOM, Me.vTblItemWO, Me.vTblItemRM)

                'isi WO number di tbl item WO
                f.IsiWONO_N_Urutan(vIdentifierJO, Me.vTblProsesOrder, Me.vTblItemWO)
                f.Dispose()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub UpdateTblItemBOM(ByVal vTblItemBomDum_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable)
        Dim rowItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
        Dim rowD As DataGridViewRow
        Dim ListPrimaryProduct As New List(Of Double)
        Dim PrimaryProd As Double = -1
        Dim lewat As Boolean = False
        Dim count_ As Integer = 1
        Dim f As New FrmTrans_ManufacturingOrder
        Dim Filter As String = ""

        Try
            For i As Integer = 0 To Me.dgvItemBOM.RowCount - 2
                rowD = Me.dgvItemBOM.Rows(i)
                ListPrimaryProduct.Add(rowD.Cells(Me.ColBOM_PrimaryProduct.Index).Value)

                Filter = "[Primary Product] = " & vPrimaryProduct.ToString _
                & " AND [Primary Product RM] = " & CDbl(rowD.Cells(Me.ColBOM_PrimaryProduct.Index).Value).ToString

                'filter identifier wo untuk mengetahui bahwa product tsb di edit atau baru di tambahkan di tbl item bom
                'vTblItemBomDum_.DefaultView.RowFilter = Filter & " AND [Identifier WO] = '" & Me.vIdentifierWO & "'"
                'If vTblItemBomDum_.DefaultView.Count > 0 Then
                    'filter identifier jo untuk mengubah semua product dan product rm (bila ada product dan product rm yang sama)
                    vTblItemBomDum_.DefaultView.RowFilter = Filter & " AND [Identifier JO] = '" & Me.vIdentifierJO & "'"
                    If vTblItemBomDum_.DefaultView.Count > 0 Then
                        For Each dv As Data.DataRowView In vTblItemBomDum_.DefaultView
                            rowItemBOM = dv.Row
                            With rowItemBOM
                                .Qty_Ref_Process = CDec(Me.TxtQtyRef.Text)
                                .Good_Percentage_Process = CDec(Me.TxtGoodPercentage.Text)
                                .Primary_Process = CDbl(Me.txtPrimaryProses.Text)
                                .Primary_Product_RM = CDbl(rowD.Cells(Me.ColBOM_PrimaryProduct.Index).Value)
                                .Qty_RM = rowD.Cells(Me.ColBOM_Qty.Index).Value
                                .Good_Percentage_RM = CDec(rowD.Cells(Me.ColBOM_GoodPercentage.Index).Value)
                                .Decimal_RM = CInt(ObjProduct.GetAngkaDecimal(.Primary_Product_RM))
                            .Remarks = rowD.Cells(Me.ColBOM_Remarks.Index).Value
                            .production_tolerance = rowD.Cells(Me.ColBOM_Tolerance.Index).Value
                            .tolerance_process = CDec(Me.txtTolerance.Text)
                        End With
                    Next
                End If
                'Else
                '    rowItemBOM = vTblItemBomDum_.Newtbltrans_mo_item_bomRow
                '    With rowItemBOM
                '        .Primary = -1
                '        .Primary_MO = CDbl(Me.vPrimaryMO)
                '        .Identifier_JO = vIdentifierJO
                '        .Identifier_WO = "-1-1"
                '        .Primary_Product = vPrimaryProduct
                '        .Qty_Ref_Process = CDec(Me.TxtQtyRef.Text)
                '        .Good_Percentage_Process = CDec(Me.TxtGoodPercentage.Text)
                '        .Primary_Process = CDbl(Me.txtPrimaryProses.Text)
                '        .Primary_Product_RM = CDbl(rowD.Cells(Me.ColBOM_PrimaryProduct.Index).Value)
                '        .Qty_RM = rowD.Cells(Me.ColBOM_Qty.Index).Value
                '        .Good_Percentage_RM = CDec(rowD.Cells(Me.ColBOM_GoodPercentage.Index).Value)
                '        .Decimal_RM = CInt(ObjProduct.GetAngkaDecimal(.Primary_Product_RM))
                '        .Remarks = rowD.Cells(Me.ColBOM_Remarks.Index).Value
                '        .Primary_Parent_Product = Me.vPrimaryParentProduct
                '        .Identifier_BOM = ObjMO_Item.GetNextIdentifierBOM(vTblItemBomDum_, Me.vPrimaryMO)
                '        .Identifier_BOM_Parent = Me.vIdentifierBOM
                '        .Identifier_WO = "-1-1"
                '    End With
                '    vTblItemBomDum_.Rows.Add(rowItemBOM)

                '    'f.IsiTblItemBOM(rowItemBOM.Primary_Product_RM, Me.vIdentifierJO, vTblItemBomDum_, Me.vPrimaryProduct)
                'End If
            Next

            'hapus item bom
            vTblItemBomDum_.DefaultView.RowFilter = "[Primary Product] = " & vPrimaryProduct.ToString & _
            " AND [Identifier JO] = '" & Me.vIdentifierJO & "'" '" AND [Identifier WO] = '" & Me.vIdentifierWO & "'"
            If vTblItemBomDum_.DefaultView.Count > 0 Then
                For Each dv_1 As Data.DataRowView In vTblItemBomDum_.DefaultView
                    rowItemBOM = dv_1.Row
                    If ListPrimaryProduct.Contains(rowItemBOM.Primary_Product_RM) = False Then
                        'jika primary product di tbl bom tidak ada di list primary product di dgv maka hapus barisnya beserta bom bila memiliki
                        If rowItemBOM.Primary_Product_RM <> -1 Then
                            Me.DeleteTblItemBOM(rowItemBOM.Primary_Product, rowItemBOM.Primary_Product_RM, rowItemBOM.Identifier_BOM_Parent, vTblItemBomDum_)
                        Else
                            'bila product rm tidak memiliki bom maka tidak usah masuk fungsi hapus bom
                            vTblItemBomDum_.Rows.Remove(rowItemBOM)
                        End If
                    End If
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            f.Dispose()
        End Try
    End Sub

    Private Sub UpdateTblProsesOrder_BOM()
        Dim rowItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
        Dim rowProsesOrder As DseTrans_ManufacturingOrder.tbltrans_mo_processorderRow
        Dim urutanMax As Integer = 0
        Dim ListPrimaryProses As New List(Of Double)
        Dim PrimaryProses As Double = -1

        Try
            'list primary proses yang ada di TblItemBOM
            vTblItemBOM.DefaultView.RowFilter = "[identifier jo] = '" & vIdentifierJO & "'"
            If vTblItemBOM.DefaultView.Count > 0 Then
                For Each dv As Data.DataRowView In vTblItemBOM.DefaultView
                    rowItemBOM = dv.Row
                    If ListPrimaryProses.Contains(rowItemBOM.Primary_Process) = False Then
                        ListPrimaryProses.Add(rowItemBOM.Primary_Process)
                    End If
                Next
            End If

            'bandingkan primary proses di TblItemBOM dengan primary proses di TblProsesOrder
            vTblProsesOrder.DefaultView.RowFilter = "[identifier jo] = '" & vIdentifierJO.ToString & "'"
            vTblProsesOrder.DefaultView.Sort = "[urutan] ASC"
            If vTblProsesOrder.DefaultView.Count > 0 Then
                For Each dv As Data.DataRowView In vTblProsesOrder.DefaultView
                    rowProsesOrder = dv.Row
                    PrimaryProses = rowProsesOrder.primary_process
                    urutanMax = rowProsesOrder.urutan
                    If ListPrimaryProses.Contains(rowProsesOrder.primary_process) = False Then
                        vTblProsesOrder.Rows.Remove(rowProsesOrder)
                    End If
                    ListPrimaryProses.Remove(PrimaryProses)
                Next

                If ListPrimaryProses.Count > 0 Then
                    For Each primary_proses As Double In ListPrimaryProses
                        rowProsesOrder = vTblProsesOrder.Newtbltrans_mo_processorderRow
                        With rowProsesOrder
                            .primary = -1
                            .primary_mo = CDbl(vPrimaryMO)
                            .primary_process = primary_proses
                            .primary_product = vPrimaryProductJO
                            .identifier_jo = vIdentifierJO
                            urutanMax += 1
                            .urutan = urutanMax
                        End With
                        vTblProsesOrder.Rows.Add(rowProsesOrder)
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateTblItemWORM_BOM(ByVal vPrimaryProductRemove As List(Of Double), ByVal vPrimaryProductAdd As List(Of Double), _
    ByVal vLevel As Integer, ByVal vQtyTarget As Decimal, ByVal vIdentifierJO As String, Optional ByVal vIdentifierParent As String = "")

        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DeleteTblItemBOM(ByVal vPrimaryProduct_2 As Double, ByVal vPrimaryProductRM_2 As Double, _
    ByVal vIdentifierBOM_parent As String, ByVal vTblItemBom_2 As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable)
        Dim rowItemBom_dum As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
        Dim filter As String = ""

        Try
            If vPrimaryProductRM_2 <> -1 Then
                filter = "[Primary Product] = " & vPrimaryProduct_2.ToString _
                & " AND [Primary Product RM] = " & vPrimaryProductRM_2.ToString _
                & " AND [Identifier BOM Parent] = '" & vIdentifierBOM_parent & "' AND [Identifier JO] = '" & Me.vIdentifierJO & "'"
            Else
                filter = "[Primary Product] = " & vPrimaryProduct_2.ToString _
                & " AND [Identifier BOM Parent] = '" & vIdentifierBOM_parent & "' AND [Identifier JO] = '" & Me.vIdentifierJO & "'"
            End If

            vTblItemBom_2.DefaultView.RowFilter = filter
            If vTblItemBom_2.DefaultView.Count > 0 Then
                For Each dv As Data.DataRowView In vTblItemBom_2.DefaultView
                    rowItemBom_dum = dv.Row
                    If rowItemBom_dum.Primary_Product_RM <> -1 Then
                        DeleteTblItemBOM(rowItemBom_dum.Primary_Product_RM, -1, rowItemBom_dum.Identifier_BOM, vTblItemBom_2)
                        vTblItemBom_2.Rows.Remove(rowItemBom_dum)
                    End If
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DeleteTblItemWORM(ByVal vTblItemBom_2 As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable, _
    ByVal vTblItemWO_2 As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vTblItemRM_2 As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable)
        'Dim rowItemBom_dum As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
        Dim rowItemWO_dum As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim rowItemRM_dum As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow

        Try
            vTblItemWO_2.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO & "'"
            If vTblItemWO_2.DefaultView.Count > 0 Then
                For Each dv As Data.DataRowView In vTblItemWO_2.DefaultView
                    rowItemWO_dum = dv.Row
                    vTblItemBom_2.DefaultView.RowFilter = "[Identifier WO] = '" & rowItemWO_dum.Identifier & "'"
                    If vTblItemBom_2.DefaultView.Count = 0 Then
                        vTblItemWO_2.Rows.Remove(rowItemWO_dum)
                    End If
                Next
            End If

            vTblItemRM_2.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO & "'"
            If vTblItemRM_2.DefaultView.Count > 0 Then
                For Each dv_2 As Data.DataRowView In vTblItemRM_2.DefaultView
                    rowItemRM_dum = dv_2.Row
                    vTblItemBom_2.DefaultView.RowFilter = "[Identifier WO] = '" & rowItemRM_dum.Identifier_WO.ToString _
                    & "' AND [Primary Product RM] = " & rowItemRM_dum.Primary_Product_Code_RM
                    If vTblItemBom_2.DefaultView.Count = 0 Then
                        vTblItemRM_2.Rows.Remove(rowItemRM_dum)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ManufacturingOrder_BOM_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Hasil As MsgBoxResult

        Try
            If vBolehEdit = True Then
                If Me.FromOtherButton = False Then
                    Hasil = MsgBox("Accept data changes before closing form?", MsgBoxStyle.YesNoCancel, "Confirm")
                    If Hasil = MsgBoxResult.Yes Then
                        If Me.ApakahDataBenar = True Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            Dim vFrm As New FrmTrans_ManufacturingOrder
                            Me.UpdateAllTblItem()
                            vFrm.Dispose()
                            Me.DialogResult = Windows.Forms.DialogResult.OK
                            Me.FromOtherButton = True
                            Me.Close()
                        End If
                    ElseIf Hasil = MsgBoxResult.No Then
                        If Me.TblBOM_Temp.Rows.Count > 0 Then
                            For Each Me.RowBOM_Temp In Me.TblBOM_Temp.Rows
                                Me.DeleteTblItemBOM(Me.RowBOM_Temp.Primary_Product, -1, Me.RowBOM_Temp.Primary_Parent_Product, Me.TblItemBOM_)
                                'Me.TblBOM_Temp.Rows.Remove(Me.RowBOM_Temp)
                            Next
                        End If
                        Me.FromOtherButton = True
                        Me.Close()
                    ElseIf Hasil = MsgBoxResult.Cancel Then
                        e.Cancel = True
                    End If
                Else
                    If Me.CancelClose = True Then
                        Me.CancelClose = False
                        e.Cancel = True
                    End If
                    Me.FromOtherButton = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub FrmTrans_ManufacturingOrder_BOM_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ''ok-rosadin-r5b
        Try
            Me.ColBOM_IdentifierBOM.Visible = False
            Me.ColBOM_IdentifierBOM_Parent.Visible = False
            Me.ColBOM_Qty.DefaultCellStyle.NullValue = 0
            Me.dtaBundling.Connection.ConnectionString = Core_Module.GetConstring
            Me.dtaBundling.Fill(Me.tblBundling, Me.vIdentifierJO)
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ManufacturingOrder_BOM_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Dim rowItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow

        Try
            'Me.ColBOM_IdentifierBOM.Visible = True
            'Me.ColBOM_IdentifierBOM_Parent.Visible = True
            Me.dgvItemBOM.ScrollBars = ScrollBars.Both
            Me.llbAddNewProduct.Visible = False
            Me.dgvItemBOM.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            If vBolehEdit = False Then
                Me.txtProcessCode.Enabled = False
                Me.TxtProcessDesc.Enabled = False
                Me.btnFindProcess.Enabled = False
                Me.TxtQtyRef.Enabled = False
                Me.TxtGoodPercentage.Enabled = False
                Me.dgvItemBOM.ReadOnly = True
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Me.BtnOK.Visible = False
                Me.ColBOM_DrillDown.Visible = False
                Me.llbAddNewProduct.Visible = False
                Me.dgvItemBOM.AllowUserToAddRows = False
            End If

            If vLevel = 0 Then
                Me.TblItemBOM_ = Me.vTblItemBOM.Copy

                Me.TblItemBOM_.DefaultView.RowFilter = "[Primary Product] = " & Me.vPrimaryProduct.ToString & " AND [Identifier JO] = '" & Me.vIdentifierJO & "'"
                If Me.TblItemBOM_.DefaultView.Count > 0 Then
                    rowItemBOM = Me.TblItemBOM_.DefaultView(0).Row
                    Me.vIdentifierBOM_Parent = rowItemBOM.Identifier_BOM_Parent
                    Me.vIdentifierWO = rowItemBOM.Identifier_WO
                End If
            Else
                Me.TblItemBOM_.DefaultView.RowFilter = "[Identifier BOM] = '" & Me.vIdentifierBOM_Parent & "'"
                If Me.TblItemBOM_.DefaultView.Count > 0 Then
                    rowItemBOM = Me.TblItemBOM_.DefaultView(0).Row
                    Me.vPrimaryProduct = rowItemBOM.Primary_Product_RM
                End If
                Me.TblItemBOM_.DefaultView.RowFilter = "[Identifier BOM Parent] = '" & Me.vIdentifierBOM_Parent & "'"
                If Me.TblItemBOM_.DefaultView.Count > 0 Then
                    rowItemBOM = Me.TblItemBOM_.DefaultView(0).Row
                    Me.vIdentifierWO = rowItemBOM.Identifier_WO
                End If
            End If

            Me.LblUnitBOM.Text = Me.objUnit.GetUnitCode(Me.ObjProduct.GetPrimaryUnit(Me.vPrimaryProduct))
            Me.LblUnitBOM.Left = Me.Label1.Right - Me.LblUnitBOM.Width
            Me.TxtQtyRef.Width = Me.LblUnitBOM.Left - 5 - Me.TxtQtyRef.Left
            Me.TxtQtyRef.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyRef.Text, Me.ObjProduct.GetAngkaDecimal(Me.vPrimaryProduct))
            Me.ObjMO_Item.IsiDgvItemBOM(vIdentifierJO, Me.vIdentifierBOM_Parent, Me.dgvItemBOM, Me.TblItemBOM_, Me)

            Me.Update_TreeView()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

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

    Private Sub txtProcessCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProcessCode.GotFocus, TxtQtyRef.GotFocus, TxtGoodPercentage.GotFocus
        'Dim Txt As TextBox
        'Try
        '    Txt = TryCast(sender, TextBox)
        '    If IsNothing(Txt) = False Then
        '        If Txt.Enabled = True Then
        '            If Txt.ReadOnly = True Then
        '                Txt.BackColor = Me.TxtProcessDesc.BackColor
        '            Else
        '                Txt.BackColor = Me.BackColorAll
        '            End If
        '        End If
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub txtProcessCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProcessCode.LostFocus, TxtQtyRef.LostFocus, TxtGoodPercentage.LostFocus
        'Dim Txt As TextBox
        'Try
        '    Txt = TryCast(sender, TextBox)
        '    If IsNothing(Txt) = False Then
        '        If Txt.Enabled = True Then
        '            If Txt.ReadOnly = True Then
        '                Txt.BackColor = Me.TxtProcessDesc.BackColor
        '            Else
        '                Txt.BackColor = Me.txtPrimaryProses.BackColor
        '            End If
        '        End If
        '    End If
        '    Txt.Text = Trim(Txt.Text).ToString.Trim.Replace("%", "").Replace("*", "").Replace("'", "").Replace(" ", "")
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub txtProcessCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtProcessCode.KeyDown
        'Dim i As New Core_Module.FrmMasterProsesCari

        'Try
        '    i.BackColorAll = Me.BackColorAll
        '    If e.KeyValue = 13 And Me.txtProcessCode.ReadOnly = False And Me.txtProcessCode.Enabled = True Then
        '        If ObjProses.ApakahSudahAda(Me.txtProcessCode.Text) = True Then
        '            Me.txtPrimaryProses.Text = ObjProses.GetPrimary(Me.txtProcessCode.Text)
        '            Me.TxtQtyRef.Focus()
        '        Else
        '            If Me.txtProcessCode.Text <> "" Then
        '                i.txtCari.Text = Me.txtProcessCode.Text
        '            End If

        '            i.ShowDialog()
        '            If i.mPrimary > 0 Then
        '                Me.txtPrimaryProses.Text = i.mPrimary
        '                Me.txtProcessCode.Text = ObjProses.GetCode(Me.txtPrimaryProses.Text)
        '                Me.TxtQtyRef.Focus()
        '            Else
        '                Me.txtProcessCode.Clear()
        '                Me.txtPrimaryProses.Text = "-1"
        '                Me.txtProcessCode.Focus()
        '                Me.Kosongkan_BOM()
        '            End If
        '        End If

        '        Me.TxtProcessDesc.Text = Me.ObjProses.GetDesc(Me.txtPrimaryProses.Text)
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'Finally
        '    i.Dispose()
        'End Try
    End Sub

    Private Sub txtProcessCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProcessCode.Leave
        'Dim i As New Core_Module.FrmMasterProsesCari

        'Try
        '    i.BackColorAll = Me.BackColorAll
        '    If Me.txtProcessCode.Text <> "" Then
        '        If ObjProses.ApakahSudahAda(Me.txtProcessCode.Text) = True Then
        '            Me.txtPrimaryProses.Text = ObjProses.GetPrimary(Me.txtProcessCode.Text)
        '        Else
        '            If Me.txtProcessCode.Text <> "" Then
        '                i.txtCari.Text = Me.txtProcessCode.Text
        '            End If

        '            i.ShowDialog()
        '            If i.mPrimary > 0 Then
        '                Me.txtPrimaryProses.Text = i.mPrimary
        '                Me.txtProcessCode.Text = ObjProses.GetCode(Me.txtPrimaryProses.Text)
        '                Me.TxtQtyRef.Focus()
        '            Else
        '                Me.txtProcessCode.Clear()
        '                Me.txtPrimaryProses.Text = "-1"
        '                Me.Kosongkan_BOM()
        '                Me.txtProcessCode.Focus()
        '            End If
        '        End If
        '    Else
        '        Me.txtPrimaryProses.Text = "-1"
        '        Me.Kosongkan_BOM()
        '    End If
        '    Me.TxtProcessDesc.Text = Me.ObjProses.GetDesc(Me.txtPrimaryProses.Text)
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'Finally
        '    i.Dispose()
        'End Try
    End Sub

    Private Sub btnFindProses_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindProcess.Click
        Dim i As New Core_Module.FrmMasterProsesCari

        Try
            i.BackColorAll = Me.BackColorAll
            'i.cboCriteria.SelectedIndex = 0
            If Me.txtProcessCode.Text <> "" Then
                i.txtCari.Text = Me.txtProcessCode.Text
            End If

            i.ShowDialog()
            If i.mPrimary > 0 Then
                Me.txtPrimaryProses.Text = i.mPrimary
                Me.txtProcessCode.Text = ObjProses.GetCode(Me.txtPrimaryProses.Text)
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
        End Try
    End Sub

    Private Sub TxtQtyRef_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyRef.Enter
        Try
            'If Me.txtCategoryCode.Enabled = True Then
            Me.ObjCD.Number_Enter(Me.TxtQtyRef, sender, e)
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyRef_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyRef.Leave
        Try
            'If Me.txtCategoryCode.Enabled = True Then
            Me.ObjCD.Number_Leave(Me.TxtQtyRef, sender, e, vDefPrec, Me.ObjCD)
            'End If
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

    Private Sub TxtGoodPercentage_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtGoodPercentage.Enter
        Try
            Me.Buffer = Me.TxtGoodPercentage.Text
            'If Me.txtCategoryCode.Enabled = True Then
            Me.ObjCD.Number_Enter(Me.TxtGoodPercentage, sender, e)
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtGoodPercentage_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtGoodPercentage.Leave
        Try
            If Me.TxtGoodPercentage.Text = Me.Buffer Then Exit Sub
            ''If Me.txtCategoryCode.Enabled = True Then
            Me.ObjCD.Number_Leave(Me.TxtGoodPercentage, sender, e, 2, Me.ObjCD)
            Me.txtTolerance.Text = Me.ObjCD.DigitGrouping(Me.GetTolerancePerc(Me.TxtGoodPercentage.Text), 2)
            ''End If
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


    Private Sub Kosongkan_BOM()
        Try
            'kosongkan item bom
            Me.dgvItemBOM.Rows.Clear()

            Me.txtPrimaryProses.Text = "-1"
            Me.txtProcessCode.Text = ""
            Me.TxtQtyRef.Text = Me.ObjCD.DigitGrouping(1, vDefPrec)
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
        Dim ListIdentifierBOM As New List(Of String)
        Dim TrackExplodeItem As Boolean = True

        Try
            If Me.trvBOM.Nodes(0).Tag <> Me.vIdentifierBOM_Parent Then
                'jika item baru, maka posisi expand = false
                TrackExplodeItem = False
            Else
                TrackExplodeItem = Me.trvBOM.Nodes(0).IsExpanded
            End If
            'catatkan node yang di explode
            If TrackExplodeItem = True Then
                Me.Update_TreeView_GetExplodeItem(ListIdentifierBOM, Me.trvBOM.Nodes(0))
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
                    Child.Tag = rowdgv.Cells(Me.ColBOM_IdentifierBOM.Index).Value
                End If
            Next

            'update sisa node level selanjutnya
            For Each Child In Me.trvBOM.Nodes(0).Nodes
                Me.Update_TreeView_IsikanChild(Child)
            Next
            'explode node seperti kondisi sebelumnya
            If TrackExplodeItem = True Then
                Me.Update_TreeView_ExplodeItem(ListIdentifierBOM, Me.trvBOM.Nodes(0))
                Me.trvBOM.Nodes(0).Expand()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            ListIdentifierBOM = Nothing
        End Try
    End Sub

    Private Sub Update_TreeView_ExplodeItem(ByVal vListIdentifierBOM As List(Of String), ByVal vParentNode As TreeNode)
        Dim ChildNode As TreeNode

        Try
            If vParentNode.Nodes.Count > 0 Then
                If vListIdentifierBOM.Contains(vParentNode.Tag) = True Then
                    vParentNode.Expand()
                End If

                For Each ChildNode In vParentNode.Nodes
                    If ChildNode.Nodes.Count > 0 Then
                        Me.Update_TreeView_ExplodeItem(vListIdentifierBOM, ChildNode)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Update_TreeView_GetExplodeItem(ByVal vListIdentifierBOM As List(Of String), ByVal vParentNode As TreeNode)
        Dim ChildNode As TreeNode

        Try
            If vParentNode.Nodes.Count > 0 Then
                For Each ChildNode In vParentNode.Nodes
                    If ChildNode.IsExpanded = True Then
                        vListIdentifierBOM.Add(ChildNode.Tag)
                    End If
                    If ChildNode.Nodes.Count > 0 Then
                        Me.Update_TreeView_GetExplodeItem(vListIdentifierBOM, ChildNode)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Update_TreeView_IsikanChild(ByVal vNodeParent As TreeNode)
        Dim Child As TreeNode
        Dim DTAItem As New Core_Module.dseMasterProduct_BOMTableAdapters.TreeViewTableTableAdapter
        Dim TBLItem As New Core_Module.dseMasterProduct_BOM.TreeViewTableDataTable
        Dim rowItem As Core_Module.dseMasterProduct_BOM.TreeViewTableRow
        Dim rowItemBom As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
        Dim ListPrimaryProduct As New List(Of Double)

        Try
            'kosongkan node
            vNodeParent.Nodes.Clear()
            'isikan dr database
            If vBolehEdit = True Then
                Me.TblItemBOM_.DefaultView.RowFilter = "[Identifier BOM Parent] = '" & vNodeParent.Tag & "'"
                If Me.TblItemBOM_.DefaultView.Count > 0 Then
                    ListPrimaryProduct.Clear()
                    For Each dv As Data.DataRowView In Me.TblItemBOM_.DefaultView
                        rowItemBom = dv.Row
                        If ListPrimaryProduct.Contains(rowItemBom.Primary_Product_RM) = False Then
                            ListPrimaryProduct.Add(rowItemBom.Primary_Product_RM)
                            If Me.chkShowProses.Checked = True Then
                                vNodeParent.Nodes.Add(Me.ObjProses.GetCode(rowItemBom.Primary_Process) + " : " + Me.ObjProduct.GetProductCode(rowItemBom.Primary_Product_RM) + " - " + Me.ObjProduct.GetDescription(rowItemBom.Primary_Product_RM))
                            Else
                                vNodeParent.Nodes.Add(Me.ObjProduct.GetProductCode(rowItemBom.Primary_Product_RM) + " - " + Me.ObjProduct.GetDescription(rowItemBom.Primary_Product_RM))
                            End If
                            Child = vNodeParent.Nodes(vNodeParent.Nodes.Count - 1)
                            Child.Tag = rowItemBom.Identifier_BOM
                        End If
                    Next
                End If
            Else
                Me.vTblItemBOM.DefaultView.RowFilter = "[Identifier BOM Parent] = '" & vNodeParent.Tag & "'"
                If Me.vTblItemBOM.DefaultView.Count > 0 Then
                    ListPrimaryProduct.Clear()
                    For Each dv As Data.DataRowView In Me.vTblItemBOM.DefaultView
                        rowItemBom = dv.Row
                        If ListPrimaryProduct.Contains(rowItemBom.Primary_Product_RM) = False Then
                            ListPrimaryProduct.Add(rowItemBom.Primary_Product_RM)
                            If Me.chkShowProses.Checked = True Then
                                vNodeParent.Nodes.Add(Me.ObjProses.GetCode(rowItemBom.Primary_Process) + " : " + Me.ObjProduct.GetProductCode(rowItemBom.Primary_Product_RM) + " - " + Me.ObjProduct.GetDescription(rowItemBom.Primary_Product_RM))
                            Else
                                vNodeParent.Nodes.Add(Me.ObjProduct.GetProductCode(rowItemBom.Primary_Product_RM) + " - " + Me.ObjProduct.GetDescription(rowItemBom.Primary_Product_RM))
                            End If
                            Child = vNodeParent.Nodes(vNodeParent.Nodes.Count - 1)
                            Child.Tag = rowItemBom.Identifier_BOM
                        End If
                    Next
                End If
            End If

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

    Private Sub Update_TreeView_BuangNodeLvl2(ByVal vIdentifierBOM As String)
        Try
            For Each ChildNode As TreeNode In Me.trvBOM.Nodes(0).Nodes
                If ChildNode.Tag = vIdentifierBOM Then
                    Me.trvBOM.Nodes(0).Nodes.Remove(ChildNode)
                    Exit Sub
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Update_TreeView_TambahNodeLvl2(ByVal vIdentifierBOM As String, ByVal vCodeNo As String, ByVal vIndex As Integer, ByVal vDesc As String)
        Try
            Dim ChildNode As New TreeNode()
            If Me.chkShowProses.Checked = True Then
                ChildNode.Text = Me.txtProcessCode.Text + " : " + vCodeNo + " - " + vDesc
            Else
                ChildNode.Text = vCodeNo + " - " + vDesc
            End If
            ChildNode.Tag = vIdentifierBOM
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
            Code = vProductCode '+ " - " + ProductDesc
            If Me.trvBOM.Nodes.Count = 0 Then
                Me.trvBOM.Nodes.Add(Code)
            Else
                Me.trvBOM.Nodes(0).Text = Code
            End If
            Me.trvBOM.Nodes(0).Tag = Me.vIdentifierBOM_Parent
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemBOM_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemBOM.CellBeginEdit
        Dim rowItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
        Dim TblItemBOM_2 As New DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable
        Try
            'If Me.txtCategoryCode.Enabled = True Then
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
                        Me.Buffer = Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                End If
            Else
                'jika bukan baris terakhir: code no, remarks, qty, good ref
                'If e.ColumnIndex = Me.ColBOM_CodeNo.Index Or _ -> 2017-12-03(iwan): ada masalah klo edit langsung code no disini: WIP dan RM-nya bermasalah (rm atas barang yang ditimpa tidak hilang, selain itu juga masalah siklik bundling bisa timbul)
                If e.ColumnIndex = Me.ColBOM_GoodPercentage.Index Or _
                e.ColumnIndex = Me.ColBOM_Qty.Index Or _
                e.ColumnIndex = Me.ColBOM_Remarks.Index Or _
                e.ColumnIndex = Me.ColBOM_Tolerance.Index Then

                    TblItemBOM_2 = TblItemBOM_.Copy
                    Me.TblItemBOM_.DefaultView.RowFilter = "[Primary Product] = " & Me.vPrimaryProduct.ToString & _
                    " AND [Primary Product RM] = " & CDbl(Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value).ToString & _
                    " AND [Identifier JO] = '" & Me.vIdentifierJO & "'"
                    If Me.TblItemBOM_.DefaultView.Count > 0 Then
                        For Each dv As Data.DataRowView In Me.TblItemBOM_.DefaultView
                            rowItemBOM = dv.Row
                            If ObjMO.ApakahBolehEditItemBOM(rowItemBOM.Identifier_BOM, rowItemBOM.Identifier_BOM_Parent, TblItemBOM_2) = False Then
                                MsgBox("The following product is reference by one or several bundling" & _
                                ", you can’t delete this data.", MsgBoxStyle.Information, "Information")
                                e.Cancel = True
                                Exit Sub
                            Else
                                Me.Buffer = Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value
                            End If
                        Next
                    End If
                Else
                    e.Cancel = True
                End If
            End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub EditCodeNoDrillDown(ByVal vPrimaryProduct As Double)
    '    'Dim rowItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
    '    'Dim i As New FrmTrans_ManufacturingOrder_BOM

    '    'Try
    '    '    i.vPrimaryProduct = vPrimaryProduct
    '    '    i.vPrimaryProductJO = vPrimaryProductJO
    '    '    i.vIdentifierJO = Me.vIdentifierJO
    '    '    i.vProductCode = Me.vProductCode
    '    '    i.vDefPrec = Me.vDefPrec
    '    '    i.vBolehEdit = Me.vBolehEdit
    '    '    i.vTblItemBOM = Me.vTblItemBOM
    '    '    i.vTblProsesOrder = Me.vTblProsesOrder
    '    '    i.ShowDialog()
    '    '    i.Dispose()

    '    '    'Me.UpdateItemBOM(e.RowIndex)
    '    '    Me.Update_TreeView()

    '    '    'Me.ObjMO_Item.IsiDgvItemBOM(PrimaryProduct, IdentifierJO, Me.dgvItemBOM, Me.vTblItemBOM, Me)
    '    '    'Me.Update_TreeView()

    '    '    'Me.vTblItemBOM.DefaultView.RowFilter = "[Identifier JO] = '" & Me.IdentifierJO & "' And [Primary Product] = " & vPrimaryProduct
    '    '    'If Me.vTblItemBOM.DefaultView.Count > 0 Then
    '    '    '    For Each dv As Data.DataRowView In Me.vTblItemBOM.DefaultView
    '    '    '        rowItemBOM = dv.Row

    '    '    '    Next
    '    '    'End If
    '    '    'If Me.ObjProduct.CariDetail(Me.DGVKompetensi, vPrimaryProduct, Me.dgvPricing, Me.pbxProduct, Me.dgvItemBOM, _
    '    '    'Me, Me.dgvItem_MultiUnit, Me.DGVAttanchment, Me.DgvWorkUnit) = True Then
    '    '    '    Me.SalinKeField()
    '    '    '    Me.btnEdit.PerformClick()
    '    '    'Else
    '    '    '    Me.Close()
    '    '    'End If
    '    'Catch ex As Exception
    '    '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    'End Try
    'End Sub

    Private Sub dgvItemBOM_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemBOM.CellClick
        Dim Dummy As Boolean = False

        Try
            If Me.txtProcessCode.Text = "" Then
                Exit Sub
            End If

            If Me.ColBOM_DrillDown.Index = e.ColumnIndex Then
                'If Me.txtCategoryCode.Enabled = False Then
                'pada saat view
                If e.RowIndex <> Me.dgvItemBOM.Rows.Count - 1 Then
                    'hanya bisa klik item edit
                    Dim i As New FrmTrans_ManufacturingOrder_BOM
                    i.vPrimaryMO = Me.vPrimaryMO
                    i.vPrimaryProductJO = Me.vPrimaryProductJO
                    'i.vIdentifierBOM = Me.dgvItemBOM.Item(Me.ColBOM_IdentifierBOM.Index, e.RowIndex).Value
                    i.vIdentifierBOM_Parent = Me.dgvItemBOM.Item(Me.ColBOM_IdentifierBOM.Index, e.RowIndex).Value
                    i.vPrimaryProduct = Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value
                    i.vPrimaryParentProduct = Me.vPrimaryProduct
                    i.vIdentifierJO = Me.vIdentifierJO
                    i.vProductCode = Me.dgvItemBOM.Item(Me.ColBOM_CodeNo.Index, e.RowIndex).Value & " - " & Me.dgvItemBOM.Item(Me.ColBOM_Desc.Index, e.RowIndex).Value
                    i.vDefPrec = Me.ObjProduct.GetAngkaDecimal(CDbl(Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value))
                    i.vBolehEdit = Me.vBolehEdit
                    i.vLevel = Me.vLevel + 1
                    i.TblItemBOM_ = Me.TblItemBOM_
                    i.ShowDialog()
                    i.Dispose()
                    Me.UpdateItemBOM(e.RowIndex)
                    Me.Update_TreeView()
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
            rowDGV.Cells(Me.ColBOM_CodeNo.Index).Value = Me.ObjProduct.GetProductCode(rowDGV.Cells(Me.ColBOM_PrimaryProduct.Index).Value)
            rowDGV.Cells(Me.ColBOM_Desc.Index).Value = Me.ObjProduct.GetDescription(rowDGV.Cells(Me.ColBOM_PrimaryProduct.Index).Value)
            rowDGV.Cells(Me.ColBOM_Unit.Index).Value = Me.ObjProduct.GetUnitCode(rowDGV.Cells(Me.ColBOM_PrimaryProduct.Index).Value)
            rowDGV.Cells(Me.ColBOM_Qty.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColBOM_Qty.Index).Value, _
            Me.ObjProduct.GetAngkaDecimal(rowDGV.Cells(Me.ColBOM_PrimaryProduct.Index).Value))
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
                    If Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    '==Code No
                    If e.ColumnIndex = Me.ColBOM_CodeNo.Index Then
                        Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        If Me.ObjProduct.ApakahSudahAda(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryCodeNo = Me.ObjProduct.GetPrimary(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexBOM(PrimaryCodeNo)
                            If index < 0 Then
                                Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimary(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.dgvItemBOM.Item(Me.ColBOM_CodeNo.Index, e.RowIndex).Value = Me.ObjProduct.GetProductCode(PrimaryCodeNo)
                                Me.dgvItemBOM.Item(Me.ColBOM_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(PrimaryCodeNo)
                                Me.dgvItemBOM.Item(Me.ColBOM_DrillDown.Index, e.RowIndex).Value = "Edit"
                                Me.dgvItemBOM.Item(Me.ColBOM_Qty.Index, e.RowIndex).Value = _
                                Me.ObjCD.DigitGrouping(1, Me.ObjProduct.GetAngkaDecimal(Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, _
                                e.RowIndex).Value))
                                Me.dgvItemBOM.Item(Me.ColBOM_GoodPercentage.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(100, 2)
                                Me.dgvItemBOM.Item(Me.ColBOM_Unit.Index, e.RowIndex).Value = _
                                Me.ObjProduct.GetUnitCode(Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value)

                                If Me.ApakahSudahAdaBOM(Me.vPrimaryProduct, Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value) = True Then
                                    Me.dgvItemBOM.Item(Me.ColBOM_IdentifierBOM.Index, e.RowIndex).Value = Me.GetIdentifierBOM(Me.vPrimaryProduct, Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value)
                                Else
                                    If Me.dgvItemBOM.Item(Me.ColBOM_IdentifierBOM.Index, e.RowIndex).Value = "-" Then
                                        Me.dgvItemBOM.Item(Me.ColBOM_IdentifierBOM.Index, e.RowIndex).Value = ObjMO_Item.GetNextIdentifierBOM(Me.TblItemBOM_, Me.vPrimaryMO)
                                    End If

                                    Me.TblItemBom_IsiTemporary(Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value, Me.TblItemBOM_, _
                                    Me.dgvItemBOM.Item(Me.ColBOM_IdentifierBOM.Index, e.RowIndex).Value)
                                End If

                                'jika baris ditengah, maka buang node kemudian insert node
                                If e.RowIndex <> Me.dgvItemBOM.Rows.Count - 1 Then
                                    Me.Update_TreeView_BuangNodeLvl2(Me.dgvItemBOM.Item(Me.ColBOM_IdentifierBOM.Index, e.RowIndex).Value)
                                End If

                                'tambahkan node baru
                                Me.Update_TreeView_TambahNodeLvl2(Me.dgvItemBOM.Item(Me.ColBOM_IdentifierBOM.Index, e.RowIndex).Value, _
                                Me.dgvItemBOM.Item(Me.ColBOM_CodeNo.Index, e.RowIndex).Value, e.RowIndex, Me.dgvItemBOM.Item(Me.ColBOM_Desc.Index, e.RowIndex).Value)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication in product code, please change your product code", MsgBoxStyle.Information, "Information")
                                    Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    If Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                        Me.IndexRemove = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    End If

                                    Exit Sub
                                Else
                                    Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                        Else
                            ''diketik tidak ada lalu tampilkan form cari
                            Dim ij As New Core_Module.frmCariMasterProduct
                            ij.TipePencarian = Core_Module.frmCariMasterProduct.enumTipePencarian.Goods_Only
                            ij.txtCari.Text = Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.ShowDialog()
                            index = Me.GetIndexBOM(ij.PrimaryProduct)
                            If ij.PrimaryProduct = -1 Then
                                ''hapuskan
                                If Me.Buffer = "" Then
                                    Me.IndexRemove = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                Else
                                    Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                End If
                                Exit Sub
                            Else
                                If index < 0 Then
                                    Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value = ij.PrimaryProduct.ToString
                                    Me.dgvItemBOM.Item(Me.ColBOM_CodeNo.Index, e.RowIndex).Value = Me.ObjProduct.GetProductCode(ij.PrimaryProduct)
                                    Me.dgvItemBOM.Item(Me.ColBOM_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(ij.PrimaryProduct)
                                    Me.dgvItemBOM.Item(Me.ColBOM_DrillDown.Index, e.RowIndex).Value = "Edit"
                                    Me.dgvItemBOM.Item(Me.ColBOM_Qty.Index, e.RowIndex).Value = _
                                    Me.ObjCD.DigitGrouping(1, Me.ObjProduct.GetAngkaDecimal(Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, _
                                    e.RowIndex).Value))
                                    Me.dgvItemBOM.Item(Me.ColBOM_GoodPercentage.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(100, 2)
                                    Me.dgvItemBOM.Item(Me.ColBOM_Unit.Index, e.RowIndex).Value = _
                                    Me.ObjProduct.GetUnitCode(Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value)

                                    If Me.ApakahSudahAdaBOM(Me.vPrimaryProduct, Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value) = True Then
                                        Me.dgvItemBOM.Item(Me.ColBOM_IdentifierBOM.Index, e.RowIndex).Value = Me.GetIdentifierBOM(Me.vPrimaryProduct, Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value)
                                    Else
                                        If Me.dgvItemBOM.Item(Me.ColBOM_IdentifierBOM.Index, e.RowIndex).Value = "-" Then
                                            Me.dgvItemBOM.Item(Me.ColBOM_IdentifierBOM.Index, e.RowIndex).Value = ObjMO_Item.GetNextIdentifierBOM(Me.TblItemBOM_, Me.vPrimaryMO)
                                        End If

                                        Me.TblItemBom_IsiTemporary(Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value, Me.TblItemBOM_, _
                                        Me.dgvItemBOM.Item(Me.ColBOM_IdentifierBOM.Index, e.RowIndex).Value)
                                    End If

                                    'jika baris ditengah, maka buang node kemudian insert node
                                    If e.RowIndex <> Me.dgvItemBOM.Rows.Count - 2 Then
                                        Me.Update_TreeView_BuangNodeLvl2(Me.dgvItemBOM.Item(Me.ColBOM_IdentifierBOM.Index, e.RowIndex).Value)
                                    End If

                                    'tambahkan node baru
                                    Me.Update_TreeView_TambahNodeLvl2(Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value, _
                                    Me.dgvItemBOM.Item(Me.ColBOM_CodeNo.Index, e.RowIndex).Value, e.RowIndex, Me.dgvItemBOM.Item(Me.ColBOM_Desc.Index, e.RowIndex).Value)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication in product code, please change your product code", MsgBoxStyle.Information, "Information")
                                        Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        If Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                            Me.IndexRemove = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        End If
                                        Exit Sub
                                    Else
                                        Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        Exit Sub
                                    End If
                                End If
                            End If
                            ij = Nothing
                        End If
                    ElseIf e.ColumnIndex = Me.ColBOM_Qty.Index Then
                        'harus bilangan positif
                        If Me.ObjCD.CekBilanganPositif(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        'digit grouping
                        Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = _
                        Me.ObjCD.DigitGrouping(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value, _
                        Me.ObjProduct.GetAngkaDecimal(Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value))
                    ElseIf e.ColumnIndex = Me.ColBOM_GoodPercentage.Index Then
                        'harus bilangan positif
                        If Me.ObjCD.CekBilanganPositif(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        'maximal nilainya 100 %
                        If CDec(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value) > 100 Then
                            Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
                        Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.Replace("'", "").Replace("%", "").Replace("*", "").Trim
                    ElseIf e.ColumnIndex = Me.ColBOM_Tolerance.Index Then
                        If Me.ObjCD.CekBilanganPositif(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        If CDec(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value) < 0 Then
                            Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                        Me.dgvItemBOM.Item(Me.ColBOM_GoodPercentage.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.GetGoodPercentagePerc(Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value), 2)
                    End If
                Else
                    'hapus baris ybs
                    If e.ColumnIndex = Me.ColBOM_CodeNo.Index Then
                        If Me.Buffer = "" Then
                            Me.IndexRemove = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        Else
                            Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        Exit Sub
                    ElseIf e.ColumnIndex = Me.ColBOM_Qty.Index Or e.ColumnIndex = Me.ColBOM_GoodPercentage.Index Then
                        Me.dgvItemBOM.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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

    ''ok-rosadin-r5b
    Private Sub dgvItemBOM_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItemBOM.DataError

    End Sub

    'Private Sub UpdateTblBOM_temp(ByVal vPrimaryProduct As Double)
    '    'Try
    '    '    'update hitungan qty di tbl item wo
    '    '    Dim f As New FrmTrans_ManufacturingOrder
    '    '    If vLevel = 0 Then
    '    '        f.IsiTblItemWORM(-1, vPrimaryProduct, vQtyTarget, Me.TblItemWO_, Me.TblItemRM_, vIdentifierJO, "-", Me.TblItemBOM_, True)
    '    '    Else
    '    '        f.IsiTblItemWORM(-1, vPrimaryProduct, vQtyTarget, Me.vTblItemWO, Me.vTblItemRM, vIdentifierJO, "-", Me.vTblItemBOM, True)
    '    '    End If
    '    '    f.Dispose()
    '    'Catch ex As Exception
    '    '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    'End Try
    'End Sub

    Private Sub dgvItemBOM_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemBOM.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            'jika boleh add new row, maka setting nilai awal
            rowDGV = Me.dgvItemBOM.Rows(Me.dgvItemBOM.Rows.Count - 1)
            With rowDGV
                .Cells(Me.ColBOM_CodeNo.Index).Value = ""
                .Cells(Me.ColBOM_Desc.Index).Value = ""
                .Cells(Me.ColBOM_DrillDown.Index).Value = ""
                .Cells(Me.ColBOM_GoodPercentage.Index).Value = 100.0
                .Cells(Me.ColBOM_PrimaryProduct.Index).Value = -1
                .Cells(Me.ColBOM_Qty.Index).Value = 0
                .Cells(Me.ColBOM_Remarks.Index).Value = ""
                .Cells(Me.ColBOM_Unit.Index).Value = "Unit"
                '.Cells(Me.ColBOM_IdentifierBOM.Index).Value = ObjMO_Item.GetNextIdentifierBOM(Me.TblItemBOM_, Me.vPrimaryMO)
                .Cells(Me.ColBOM_IdentifierBOM.Index).Value = "-"
                .Cells(Me.ColBOM_IdentifierBOM_Parent.Index).Value = Me.vIdentifierBOM_Parent
                .Cells(Me.ColBOM_Tolerance.Index).Value = "0.00"
            End With

            If Me.dgvItemBOM.Rows.Count > 1 Then
                Me.txtProcessCode.ReadOnly = True
                Me.txtProcessCode.BackColor = Me.TxtProcessDesc.BackColor
                Me.btnFindProcess.Enabled = False
            Else
                Me.txtProcessCode.ReadOnly = False
                Me.txtProcessCode.BackColor = Me.txtPrimaryProses.BackColor
                Me.btnFindProcess.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemBOM_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItemBOM.RowsRemoved

        Try
            'If vBolehEdit = True Then
            If Me.dgvItemBOM.Rows.Count > 1 Then
                Me.txtProcessCode.ReadOnly = True
                Me.txtProcessCode.BackColor = Me.TxtProcessDesc.BackColor
                Me.btnFindProcess.Enabled = False
            Else
                Me.txtProcessCode.ReadOnly = False
                Me.txtProcessCode.BackColor = Me.txtPrimaryProses.BackColor
                Me.btnFindProcess.Enabled = True
            End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemBOM_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItemBOM.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvItemBOM)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemBOM_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItemBOM.UserDeletedRow
        Try
            Me.Update_TreeView_BuangNodeLvl2(Me.Buffer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''rosaidn-r5b
    Dim dtaBundling As New DseTrans_ManufacturingOrderTableAdapters.DataTable3TableAdapter
    Dim tblBundling As New DseTrans_ManufacturingOrder.DataTable3DataTable
    Dim rowBundling As DseTrans_ManufacturingOrder.DataTable3Row
    Function ApakahBolehHapusRM(ByVal vIdentifierJO As String, ByVal vIdentifierWO As String, ByVal vPrimaryProductRM As Double, ByVal vIdentifierBoM As String, ByVal vTblItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable) As Boolean
        Dim rowItemBoM_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
        Dim TblItemBOM_2 As New DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable
        Try
            TblItemBOM_2 = vTblItemBOM.Copy
            Me.tblBundling.DefaultView.RowFilter = "[identifier jo] = '" & vIdentifierJO & "' and [identifier wo] = '" & vIdentifierWO & "' and [primary product rm] = " & vPrimaryProductRM.ToString
            If Me.tblBundling.DefaultView.Count > 0 Then
                Return False
            Else
                TblItemBOM_2.DefaultView.RowFilter = "[identifier bom parent] = '" & vIdentifierBoM & "'"
                If TblItemBOM_2.DefaultView.Count > 0 Then
                    For i As Integer = 0 To TblItemBOM_2.DefaultView.Count - 1
                        rowItemBoM_ = TblItemBOM_2.DefaultView(i).Row
                        If Me.ApakahBolehHapusRM(rowItemBoM_.Identifier_JO, rowItemBoM_.Identifier_WO, rowItemBoM_.Primary_Product_RM, rowItemBoM_.Identifier_BOM, vTblItemBOM) = False Then
                            Return False
                        End If
                    Next
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvItemBOM_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItemBOM.UserDeletingRow
        'Dim rowItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
        Dim TblItemBOM_2 As New DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable
        Try

            ''ok-rosadin-r5b
            TblItemBOM_2 = TblItemBOM_.Copy
            If Me.ApakahBolehHapusRM(Me.dgvItemBOM.Item(Me.ColBOM_IdentifierJO.Index, e.Row.Index).Value, _
                    Me.dgvItemBOM.Item(Me.ColBOM_IdentifierWO.Index, e.Row.Index).Value, _
                    Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.Row.Index).Value, _
                    Me.dgvItemBOM.Item(Me.ColBOM_IdentifierBOM.Index, e.Row.Index).Value, TblItemBOM_2) = False Then
                MsgBox("The following product is reference by one or several bundling, you can’t delete this data.", MsgBoxStyle.Information, "Information")
                e.Cancel = True
                Exit Sub
            End If
            'TblItemBOM_2 = TblItemBOM_.Copy
            'Me.TblItemBOM_.DefaultView.RowFilter = "[Primary Product] = " & Me.vPrimaryProduct.ToString & _
            '" AND [Primary Product RM] = " & CDbl(Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.Row.Index).Value).ToString & _
            '" AND [Identifier JO] = '" & Me.vIdentifierJO & "'"
            'If Me.TblItemBOM_.DefaultView.Count > 0 Then
            '    For Each dv As Data.DataRowView In Me.TblItemBOM_.DefaultView
            '        rowItemBOM = dv.Row
            '        If ObjMO.ApakahBolehHapusItemBOM(rowItemBOM.Identifier_BOM, rowItemBOM.Identifier_BOM_Parent, TblItemBOM_2) = False Then
            '            MsgBox("The following product is reference by one or several bundling" & _
            '            ", you can’t delete this data.", MsgBoxStyle.Information, "Information")
            '            e.Cancel = True
            '            Exit Sub
            '        End If
            '    Next
            'End If
            Me.TblBOM_Temp.DefaultView.RowFilter = "[Identifier BOM] = '" & Me.dgvItemBOM.Item(Me.ColBOM_IdentifierBOM.Index, e.Row.Index).Value.ToString _
                            & "' AND [Identifier BOM Parent] = '" & Me.dgvItemBOM.Item(Me.ColBOM_IdentifierBOM_Parent.Index, e.Row.Index).Value & "'"
            If Me.TblBOM_Temp.DefaultView.Count > 0 Then
                While Me.TblBOM_Temp.DefaultView.Count > 0
                    Me.RowBOM_Temp = Me.TblBOM_Temp.DefaultView(0).Row
                    Me.TblBOM_Temp.Rows.Remove(Me.RowBOM_Temp)
                End While
            End If
            Me.Buffer = Me.dgvItemBOM.Item(Me.ColBOM_IdentifierBOM.Index, e.Row.Index).Value
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
            Dim i As New FrmTrans_ManufacturingOrder_BOM
            i.vPrimaryMO = Me.vPrimaryMO
            i.vPrimaryProductJO = Me.vPrimaryProductJO
            'i.vIdentifierBOM = ChildNode.Tag
            i.vIdentifierBOM_Parent = ChildNode.Tag
            i.vIdentifierJO = Me.vIdentifierJO
            'i.vPrimaryProduct = Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, e.RowIndex).Value
            'i.vPrimaryParentProduct = Me.vPrimaryProduct
            i.vProductCode = ChildNode.Text
            i.vDefPrec = Me.vDefPrec
            i.vBolehEdit = Me.vBolehEdit
            i.vLevel = Me.vLevel + 1
            i.TblItemBOM_ = Me.TblItemBOM_
            i.ShowDialog()
            i = Nothing

            Me.Update_TreeView()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub trvBOM_NodeMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles trvBOM.NodeMouseClick
        Try
            If vBolehEdit = True Then
                If e.Button = Windows.Forms.MouseButtons.Right Then
                    Me.ContextMenuStrip1.Show(Me.trvBOM, e.X, e.Y + 10)
                    Me.trvBOM.SelectedNode = e.Node
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' tampilkan row yang diberikan pada parameter
    ''' </summary>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Private Sub ShowRow(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < Me.dgvItemBOM.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    Me.dgvItemBOM.FirstDisplayedScrollingRowIndex = Me.dgvItemBOM.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.dgvItemBOM.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While Me.dgvItemBOM.Rows(row.Index).Displayed = False
                    Me.dgvItemBOM.FirstDisplayedScrollingRowIndex = Me.dgvItemBOM.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If Me.dgvItemBOM.DisplayedRowCount(True) <> Me.dgvItemBOM.DisplayedRowCount(False) Then
                    Me.dgvItemBOM.FirstDisplayedScrollingRowIndex = Me.dgvItemBOM.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUP_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUP.Click
        Try
            If Me.dgvItemBOM.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItemBOM.SelectedRows(0)
                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.dgvItemBOM.Rows.Remove(row)
                    Me.dgvItemBOM.Rows.InsertRange(indexbaru, row)
                    Me.dgvItemBOM.Focus()
                    Me.dgvItemBOM.ClearSelection()
                    row.Selected = True
                    If row.Displayed = False Then
                        Me.ShowRow(row, True)
                    End If
                End If
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
                Me.Update_TreeView()
            End If
            Me.dgvItemBOM.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            If Me.dgvItemBOM.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItemBOM.SelectedRows(0)
                If row.Index <> Me.dgvItemBOM.Rows.Count - 2 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.dgvItemBOM.Rows.Remove(row)
                    Me.dgvItemBOM.Rows.InsertRange(indexbaru, row)
                    Me.dgvItemBOM.Focus()
                    Me.dgvItemBOM.ClearSelection()
                    row.Selected = True
                    Me.ShowRow(row, False)
                End If
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
                Me.Update_TreeView()
            End If
            Me.dgvItemBOM.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAddNewProduct_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAddNewProduct.LinkClicked
        'Dim i As New Form1
        ''Dim i As New Core_Module.frmMasterProduct

        'Try
        '    i.vTbl = Me.TblItemBOM_
        '    i.vIdnJO = Me.vIdentifierJO
        '    i.ShowDialog()
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'Finally
        '    i.Dispose()
        'End Try
    End Sub

    Private Sub TblItemBom_IsiTemporary(ByVal vPrimaryProduct_3 As Double, ByVal vTbl As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable, _
    ByVal vIdentifierBOMParent As String)
        Dim f As New FrmTrans_ManufacturingOrder
        Dim rowItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
        Dim rowD As DataGridViewRow
        Dim IdentifierBOM_Parent As String = ""

        Try
            rowD = Me.dgvItemBOM.Rows(Me.dgvItemBOM.RowCount - 2)
            rowItemBOM = vTbl.Newtbltrans_mo_item_bomRow
            With rowItemBOM
                .Primary = -1
                .Primary_MO = CDbl(Me.vPrimaryMO)
                .Identifier_JO = vIdentifierJO
                .Identifier_WO = "-1-1"
                .Primary_Product = Me.vPrimaryProduct
                .Qty_Ref_Process = CDec(Me.TxtQtyRef.Text)
                .Good_Percentage_Process = CDec(Me.TxtGoodPercentage.Text)
                .Primary_Process = CDbl(Me.txtPrimaryProses.Text)
                .Primary_Product_RM = CDbl(vPrimaryProduct_3)
                .Qty_RM = rowD.Cells(Me.ColBOM_Qty.Index).Value
                .Good_Percentage_RM = CDec(rowD.Cells(Me.ColBOM_GoodPercentage.Index).Value)
                .Decimal_RM = CInt(ObjProduct.GetAngkaDecimal(.Primary_Product_RM))
                .Remarks = rowD.Cells(Me.ColBOM_Remarks.Index).Value
                .Primary_Parent_Product = Me.vPrimaryParentProduct
                .Identifier_BOM = ObjMO_Item.GetNextIdentifierBOM(vTbl, Me.vPrimaryMO)
                IdentifierBOM_Parent = .Identifier_BOM
                .Identifier_BOM_Parent = Me.vIdentifierBOM_Parent
                .Identifier_WO = ""
                ._rm__decimal_ = .Decimal_RM
                .rm_unit_primary = Me.ObjProduct.GetPrimaryUnit(.Primary_Product_RM)
                .rm_code = Me.ObjProduct.GetProductCode(.Primary_Product_RM)
                .rm_desc = Me.ObjProduct.GetDescription(.Primary_Product_RM)
                .rm_unit_code = Me.objUnit.GetUnitCode(.rm_unit_primary)
                .production_tolerance = Math.Round(Me.GetTolerancePerc(.Good_Percentage_RM), 2)
                .tolerance_process = Math.Round(Me.GetTolerancePerc(.Good_Percentage_Process), 2)

            End With
            vTbl.Rows.Add(rowItemBOM)

            f.IsiTblItemBOM(vPrimaryProduct_3, Me.vIdentifierJO, vTbl, Me.vPrimaryProduct, IdentifierBOM_Parent, Me.vPrimaryMO)

            If Me.vIdentifierBOM_Parent = "" Then
AddRow:
                Me.RowBOM_Temp = Me.TblBOM_Temp.NewBOM_TempRow
                Me.RowBOM_Temp.Primary_Product = vPrimaryProduct_3
                Me.RowBOM_Temp.Primary_Parent_Product = Me.vPrimaryProduct
                Me.RowBOM_Temp.Identifier_JO = Me.vIdentifierJO
                Me.RowBOM_Temp.Identifier_BOM = IdentifierBOM_Parent
                Me.RowBOM_Temp.Identifier_BOM_Parent = Me.vIdentifierBOM_Parent
                Me.TblBOM_Temp.Rows.Add(Me.RowBOM_Temp)
            Else
                Me.TblBOM_Temp.DefaultView.RowFilter = "[Primary Product] = " & vPrimaryProduct_3.ToString _
                & " AND [Identifier BOM] = '" & Me.vIdentifierBOM_Parent & "'"
                If Me.TblBOM_Temp.DefaultView.Count = 0 Then
                    GoTo AddRow
                Else
                    'DeleteTblItemBOM( )
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            f.Dispose()
        End Try
    End Sub

    Public Sub TblItemBom_IsiTemporary2(ByVal vTbl As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable, ByVal vPrimaryMO As Double, ByVal vIdentifierJO As String, ByVal vPrimaryProduct As Double, ByVal vPrimaryProcess As Double, ByVal vQtyRef As Decimal, _
    ByVal vGP_Process As Decimal, ByVal vPrimaryProductRM As Double, ByVal vRM_Qty As Decimal, ByVal vGP_RM As Decimal, ByVal vIdentifierWO As String, ByVal vRemarks As String, ByVal vPrimaryParentProduct As Double, _
    ByVal vIdentifierBOM_Parent As String, ByVal vRM_Code As String, ByVal vRM_Desc As String, ByVal vToleranceRM As Decimal, ByVal vToleranceProcess As Decimal)
        Dim f As New FrmTrans_ManufacturingOrder
        Dim rowItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
        Dim IdentifierBOM_Parent As String = ""
        Try
            rowItemBOM = vTbl.Newtbltrans_mo_item_bomRow
            With rowItemBOM
                .Primary = -1
                .Primary_MO = vPrimaryMO
                .Identifier_JO = vIdentifierJO
                .Primary_Product = vPrimaryProduct
                .Primary_Process = vPrimaryProcess
                .Qty_Ref_Process = vQtyRef
                .Good_Percentage_Process = vGP_Process
                .Primary_Product_RM = vPrimaryProductRM
                .Qty_RM = vRM_Qty
                .Good_Percentage_RM = vGP_RM
                .Decimal_RM = ObjProduct.GetAngkaDecimal(.Primary_Product_RM)
                .Identifier_WO = vIdentifierWO
                .Remarks = vRemarks
                .Primary_Parent_Product = vPrimaryParentProduct
                .Identifier_BOM = ObjMO_Item.GetNextIdentifierBOM(vTbl, vPrimaryMO)
                IdentifierBOM_Parent = .Identifier_BOM
                .Identifier_BOM_Parent = vIdentifierBOM_Parent
                ._rm__decimal_ = Me.ObjProduct.GetAngkaDecimal(.Primary_Product_RM)
                .rm_unit_primary = Me.ObjProduct.GetPrimaryUnit(.Primary_Product_RM)
                .rm_code = vRM_Code
                .rm_desc = vRM_Desc
                .rm_unit_code = Me.objUnit.GetUnitCode(.rm_unit_primary)
                .production_tolerance = vToleranceRM
                .tolerance_process = vToleranceProcess
                .WIP_FG_Code = Me.ObjProduct.GetProductCode(.Primary_Product)
                .WIP_FG_Desc = Me.ObjProduct.GetDescription(.Primary_Product)
            End With
            vTbl.Rows.Add(rowItemBOM)

            f.IsiTblItemBOM(vPrimaryProductRM, vIdentifierJO, vTbl, vPrimaryProduct, IdentifierBOM_Parent, vPrimaryMO)

            If Me.vIdentifierBOM_Parent = "" Then
AddRow:
                Me.RowBOM_Temp = Me.TblBOM_Temp.NewBOM_TempRow
                Me.RowBOM_Temp.Primary_Product = vPrimaryProductRM
                Me.RowBOM_Temp.Primary_Parent_Product = vPrimaryProduct
                Me.RowBOM_Temp.Identifier_JO = vIdentifierJO
                Me.RowBOM_Temp.Identifier_BOM = IdentifierBOM_Parent
                Me.RowBOM_Temp.Identifier_BOM_Parent = IdentifierBOM_Parent
                Me.TblBOM_Temp.Rows.Add(Me.RowBOM_Temp)
            Else
                Me.TblBOM_Temp.DefaultView.RowFilter = "[Primary Product] = " & vPrimaryProductRM.ToString _
                & " AND [Identifier BOM] = '" & vIdentifierBOM_Parent & "'"
                If Me.TblBOM_Temp.DefaultView.Count = 0 Then
                    GoTo AddRow
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            f.Dispose()
        End Try
    End Sub

    Private Sub txtTolerance_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTolerance.Enter
        Try
            Me.Buffer = Me.txtTolerance.Text
            Me.ObjCD.Number_Enter(Me.txtTolerance, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub txtTolerance_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTolerance.GotFocus
        Try
            Me.txtTolerance.BackColor = Me.BackColorAll
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
            Me.ObjCD.Number_Leave(Me.txtTolerance, sender, e, 2, Me.ObjCD)
            Me.TxtGoodPercentage.Text = Me.ObjCD.DigitGrouping(Me.GetGoodPercentagePerc(Me.txtTolerance.Text), 2)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub txtTolerance_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTolerance.LostFocus
        Try
            Me.txtTolerance.BackColor = Color.White
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Public Function GetTolerancePerc(ByVal vGoodPercentage As Decimal) As Decimal
        Try
            If vGoodPercentage > 0 Then
                Return ((100 - vGoodPercentage) / vGoodPercentage) * 100
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function
    Public Function GetGoodPercentagePerc(ByVal vTolerance As Decimal) As Decimal
        Try
            Return (100 / (100 + vTolerance)) * 100
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Private Sub BtnBoMTools_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnBoMTools.Click
        Dim frmBoMTools As New Inventory_Module.frmMasterProduct_BOM1
        Try
            frmBoMTools.UserCode = Me.UserCode
            frmBoMTools.BackColorAll = Me.BackColorAll
            frmBoMTools.TblItemBOM_ = Me.TblItemBOM_
            frmBoMTools.vTblItemWO = Me.vTblItemWO
            frmBoMTools.PrimaryMO = Me.vPrimaryMO
            frmBoMTools.IdentifierJO = Me.vIdentifierJO
            frmBoMTools.ShowDialog()

            Me.ObjMO_Item.IsiDgvItemBOM(vIdentifierJO, Me.vIdentifierBOM_Parent, Me.dgvItemBOM, Me.TblItemBOM_, Me)
            Me.Update_TreeView()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class