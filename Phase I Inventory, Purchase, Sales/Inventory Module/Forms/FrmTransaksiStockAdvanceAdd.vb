Option Explicit On
Option Strict Off
Public Class FrmTransaksiStockAdvanceAdd
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public ParentForm1 As EnumRefDoc = EnumRefDoc.Location_Transfer_Source
    Dim ObjLoc As New Core_Module.clsMasterStorageArea
    Dim ObjWH As New Core_Module.clsMasterWarehouse
    'Dim ObjBarcode As New CustomSMESS.ClsTransaksiBarcode
    Public TipeRefDoc As EnumRefDoc = EnumRefDoc.Stock_Bundling_Item_After

    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjWS As New Core_Module.ClsMasterWorkShift
    Dim ObjUnit As New Core_Module.ClsMasterUnit
    Dim ObjLine As New Core_Module.ClsMasterMachineLine
    Dim ObjOperator As New Core_Module.ClsMasterKaryawan
    Dim BufferRefDoc As New EnumRefDoc
    Dim BufferPrimaryRefDoc As Double = -1
    Public InputType As ClsTransaksiBarcode.EnumInputType = ClsTransaksiBarcode.EnumInputType.Add
    Public vTBL As Core_Module.DseBarcodeInput.DataTable1DataTable
    Public vTblDgv As Core_Module.DseBarcodeInput.tbl_qty_itemDataTable
    Public vPrimaryTransaksi As Double = -1
    Dim ProsesUbahCBOrefdoc As Boolean = True
    Dim FromOtherButton As Boolean = False
    Dim CancelClose As Boolean = False
    Dim Buffer As String = ""
    Dim BufferDocRef As String = ""

    Public Enum EnumRefDoc As Integer
        Stock_Bundling_Item_After
        Stock_Bundling_Item_Before
        Stock_In_Adjust
        Stock_In_Purchase
        Stock_In_Sales_Return
        Stock_Out_Adjust
        Stock_Out_Purchase_Return
        Stock_Out_Sales
        Location_Transfer_Source
        Location_Transfer_Target
    End Enum

    Enum EnumMethod As Integer
        Method_Find
        Method_KeyDown
        Method_Leave
    End Enum

    Private Sub FrmTransaksiStockAdvanceAdd_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ObjIdentifier As New CustomSMESS.ClsIdentifier

        Try
            Me.cboRefDoc.SelectedIndex = 0
            Me.ColBatch.Visible = True
            If Me.ParentForm1 = EnumRefDoc.Stock_In_Adjust Or Me.ParentForm1 = EnumRefDoc.Stock_Bundling_Item_After Then
                Me.ColQtyOnHand.Visible = False
                Me.llbAddByQtyOnHand.Visible = False
            End If
            Me.TxtLocationSource.CharacterCasing = CharacterCasing.Upper
            Me.TxtLocationTarget.CharacterCasing = CharacterCasing.Upper
            Me.TxtRefDoc.CharacterCasing = CharacterCasing.Upper
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            If Me.ParentForm1 = EnumRefDoc.Location_Transfer_Source Or _
            Me.ParentForm1 = EnumRefDoc.Location_Transfer_Target Or _
            Me.ParentForm1 = EnumRefDoc.Stock_Bundling_Item_After Or _
            Me.ParentForm1 = EnumRefDoc.Stock_Out_Adjust Or _
            Me.ParentForm1 = EnumRefDoc.Stock_Out_Purchase_Return Or _
            Me.ParentForm1 = EnumRefDoc.Stock_Out_Sales Then
                If Me.UserCode = "USER" And ObjIdentifier.GetIdentifier = "smess_lmb" Then
                    Me.cboRefDoc.Items.Add("Location")
                    Me.cboRefDoc.Items.Add("Warehouse")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjIdentifier = Nothing
        End Try
    End Sub

    Private Sub TxtLocationSource_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtLocationSource.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.GetLocationAndWarehouse(Me.TxtPrimaryLocationSource.Text, Me.TxtLocationSource.Text, Me.TxtLocationSourceDesc.Text, _
                Me.TxtPrimaryWarehouseSource.Text, Me.TxtWarehouseSource.Text, Me.TxtWarehouseSourceDesc.Text, EnumMethod.Method_KeyDown) = True Then
                    If Me.TxtLocationTarget.Enabled = True Then
                        Me.TxtLocationTarget.Focus()
                    Else
                        Me.rboAdd.Focus()
                    End If
                Else
                    Me.TxtLocationSource.Focus()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocationSource_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocationSource.Leave
        Try
            If Me.TxtLocationSource.Text <> "" Then
                If Me.GetLocationAndWarehouse(Me.TxtPrimaryLocationSource.Text, Me.TxtLocationSource.Text, Me.TxtLocationSourceDesc.Text, _
                Me.TxtPrimaryWarehouseSource.Text, Me.TxtWarehouseSource.Text, Me.TxtWarehouseSourceDesc.Text, EnumMethod.Method_Leave) = True Then
                    If Me.TxtLocationTarget.Enabled = True Then
                        'Me.TxtLocationTarget.Focus()
                    Else
                        'Me.rboAdd.Focus()
                    End If
                Else
                    Me.TxtLocationSource.Focus()
                End If
            Else
                Me.TxtPrimaryLocationSource.Text = -1
                Me.TxtLocationSource.Text = ""
                Me.TxtLocationSourceDesc.Text = ""
                Me.TxtPrimaryWarehouseSource.Text = -1
                Me.TxtWarehouseSource.Text = ""
                Me.TxtWarehouseSourceDesc.Text = ""
                If Me.BufferDocRef <> "" Then
                    RemoveAllQtyOnLocationControl()
                End If

            End If

            'If Me.BufferDocRef <> Me.TxtLocationSource.Text Then
            Me.IsiQtyOnHandIsLocationControl()
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindLocationTarget_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindLocationTarget.Click
        Try
            If Me.GetLocationAndWarehouse(Me.TxtPrimaryLocationTarget.Text, Me.TxtLocationTarget.Text, Me.TxtLocationTargetDesc.Text, _
            Me.TxtPrimaryWarehouseTarget.Text, Me.TxtWarehouseTarget.Text, Me.TxtWarehouseTargetDesc.Text, EnumMethod.Method_Find) = True Then
                Me.rboAdd.Focus()
            Else
                Me.TxtLocationTarget.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocationTarget_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtLocationTarget.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.GetLocationAndWarehouse(Me.TxtPrimaryLocationTarget.Text, Me.TxtLocationTarget.Text, Me.TxtLocationTargetDesc.Text, _
                Me.TxtPrimaryWarehouseTarget.Text, Me.TxtWarehouseTarget.Text, Me.TxtWarehouseTargetDesc.Text, EnumMethod.Method_KeyDown) = True Then
                    Me.rboAdd.Focus()
                Else
                    Me.TxtLocationTarget.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocationTarget_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocationTarget.Leave
        Try
            If Me.TxtLocationTarget.Text <> "" Then
                If Me.GetLocationAndWarehouse(Me.TxtPrimaryLocationTarget.Text, Me.TxtLocationTarget.Text, Me.TxtLocationTargetDesc.Text, _
                Me.TxtPrimaryWarehouseTarget.Text, Me.TxtWarehouseTarget.Text, Me.TxtWarehouseTargetDesc.Text, EnumMethod.Method_Leave) = True Then
                    'Me.rboAdd.Focus()
                Else
                    Me.TxtLocationTarget.Focus()
                End If
            Else
                Me.TxtPrimaryLocationTarget.Text = -1
                Me.TxtLocationTarget.Text = ""
                Me.TxtLocationTargetDesc.Text = ""
                Me.TxtPrimaryWarehouseTarget.Text = -1
                Me.TxtWarehouseTarget.Text = ""
                Me.TxtWarehouseTargetDesc.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindLocationSource_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindLocationSource.Click
        Try
            If Me.GetLocationAndWarehouse(Me.TxtPrimaryLocationSource.Text, Me.TxtLocationSource.Text, Me.TxtLocationSourceDesc.Text, _
            Me.TxtPrimaryWarehouseSource.Text, Me.TxtWarehouseSource.Text, Me.TxtWarehouseSourceDesc.Text, EnumMethod.Method_Find) = True Then
                If Me.TxtLocationTarget.Enabled = True Then
                    Me.TxtLocationTarget.Focus()
                Else
                    Me.rboAdd.Focus()
                End If
            Else
                Me.TxtLocationSource.Focus()
            End If

            If Me.TxtLocationSource.Text = "" Then
                RemoveAllQtyOnLocationControl()
            Else
                IsiQtyOnHandIsLocationControl()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetLocationAndWarehouse(ByRef vPrimaryLocation As Double, ByRef vLocationCode As String, ByRef vLocationDesc As String, _
   ByRef vPrimaryWarehouse As Double, ByRef vWarehouseCode As String, ByRef vWarehouseDesc As String, ByVal vMethod As EnumMethod) As Boolean
        Dim i As New Core_Module.FrmCariMasterStorageArea
        Dim GetWarehouse As Boolean = False

        Try
            If vMethod = EnumMethod.Method_Find Then
                If vLocationCode <> "" Then
                    i.txtCari.Text = vLocationCode
                End If

                i.ShowDialog()
                If i.PrimaryStorageArea > 0 Then
                    vPrimaryLocation = i.PrimaryStorageArea
                    vLocationCode = Me.ObjLoc.GetStorageAreaCode(vPrimaryLocation)
                    vLocationDesc = Me.ObjLoc.GetDescription(vPrimaryLocation)
                    GetWarehouse = True
                End If
            Else
                If Me.ObjLoc.ApakahSudahAda(vLocationCode) = True Then
                    vPrimaryLocation = Me.ObjLoc.GetPrimary(vLocationCode)
                    vLocationCode = Me.ObjLoc.GetStorageAreaCode(vPrimaryLocation)
                    vLocationDesc = Me.ObjLoc.GetDescription(vPrimaryLocation)
                    GetWarehouse = True
                Else
                    If vLocationCode <> "" Then
                        i.txtCari.Text = vLocationCode
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea > 0 Then
                        vPrimaryLocation = i.PrimaryStorageArea
                        vLocationCode = Me.ObjLoc.GetStorageAreaCode(vPrimaryLocation)
                        vLocationDesc = Me.ObjLoc.GetDescription(vPrimaryLocation)
                        GetWarehouse = True
                    End If
                End If
            End If

            If GetWarehouse = True Then
                vPrimaryWarehouse = Me.ObjLoc.GetPrimaryWarehouse(vPrimaryLocation)
                vWarehouseCode = Me.ObjWH.GetWarehouseCode(vPrimaryWarehouse)
                vWarehouseDesc = Me.ObjWH.GetDescription(vPrimaryWarehouse)

                Return True
            End If

            vPrimaryLocation = -1
            vLocationCode = ""
            vLocationDesc = ""
            vPrimaryWarehouse = -1
            vWarehouseCode = ""
            vWarehouseDesc = ""

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")

            vPrimaryLocation = -1
            vLocationCode = ""
            vLocationDesc = ""
            vPrimaryWarehouse = -1
            vWarehouseCode = ""
            vWarehouseDesc = ""
            Return False
        Finally
            i.Dispose()
        End Try
    End Function

    Private Sub TxtRefDoc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtRefDoc.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TipeRefDoc = EnumRefDoc.Stock_Bundling_Item_After Then
                    Me.FindByStockBundling(EnumMethod.Method_KeyDown)
                ElseIf Me.TipeRefDoc = EnumRefDoc.Stock_Bundling_Item_Before Then
                    Me.FindByStockBundling(EnumMethod.Method_KeyDown)
                ElseIf Me.TipeRefDoc = EnumRefDoc.Stock_In_Adjust Then
                    Me.FindByInOut(EnumMethod.Method_KeyDown, ClsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust)
                ElseIf Me.TipeRefDoc = EnumRefDoc.Stock_In_Purchase Then
                    Me.FindByInOut(EnumMethod.Method_KeyDown, ClsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase)
                ElseIf Me.TipeRefDoc = EnumRefDoc.Stock_In_Sales_Return Then
                    Me.FindByInOut(EnumMethod.Method_KeyDown, ClsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn)
                ElseIf Me.TipeRefDoc = EnumRefDoc.Stock_Out_Adjust Then
                    Me.FindByInOut(EnumMethod.Method_KeyDown, ClsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust)
                ElseIf Me.TipeRefDoc = EnumRefDoc.Stock_Out_Purchase_Return Then
                    Me.FindByInOut(EnumMethod.Method_KeyDown, ClsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn)
                ElseIf Me.TipeRefDoc = EnumRefDoc.Stock_Out_Sales Then
                    Me.FindByInOut(EnumMethod.Method_KeyDown, ClsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales)
                ElseIf Me.TipeRefDoc = EnumRefDoc.Location_Transfer_Source Then
                    Me.FindByLocationTransfer(EnumMethod.Method_KeyDown)
                ElseIf Me.TipeRefDoc = EnumRefDoc.Location_Transfer_Target Then
                    Me.FindByLocationTransfer(EnumMethod.Method_KeyDown)
                End If


            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRefDoc_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRefDoc.Leave
        Try
            If Me.TxtRefDoc.Text <> "" Then
                If Me.TipeRefDoc = EnumRefDoc.Stock_Bundling_Item_After Then
                    Me.FindByStockBundling(EnumMethod.Method_Leave)
                ElseIf Me.TipeRefDoc = EnumRefDoc.Stock_Bundling_Item_Before Then
                    Me.FindByStockBundling(EnumMethod.Method_Leave)
                ElseIf Me.TipeRefDoc = EnumRefDoc.Stock_In_Adjust Then
                    Me.FindByInOut(EnumMethod.Method_Leave, ClsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust)
                ElseIf Me.TipeRefDoc = EnumRefDoc.Stock_In_Purchase Then
                    Me.FindByInOut(EnumMethod.Method_Leave, ClsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase)
                ElseIf Me.TipeRefDoc = EnumRefDoc.Stock_In_Sales_Return Then
                    Me.FindByInOut(EnumMethod.Method_Leave, ClsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn)
                ElseIf Me.TipeRefDoc = EnumRefDoc.Stock_Out_Adjust Then
                    Me.FindByInOut(EnumMethod.Method_Leave, ClsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust)
                ElseIf Me.TipeRefDoc = EnumRefDoc.Stock_Out_Purchase_Return Then
                    Me.FindByInOut(EnumMethod.Method_Leave, ClsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn)
                ElseIf Me.TipeRefDoc = EnumRefDoc.Stock_Out_Sales Then
                    Me.FindByInOut(EnumMethod.Method_Leave, ClsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales)
                ElseIf Me.TipeRefDoc = EnumRefDoc.Location_Transfer_Source Then
                    Me.FindByLocationTransfer(EnumMethod.Method_Leave)
                ElseIf Me.TipeRefDoc = EnumRefDoc.Location_Transfer_Target Then
                    Me.FindByLocationTransfer(EnumMethod.Method_Leave)
                End If
            Else
                Me.TxtRefDoc.Text = ""
                Me.TxtPrimaryRefDoc.Text = -1
            End If

            If TxtRefDoc.Text = "" Then
                Me.dgvItem.Rows.Clear()
            Else
                If Me.dgvItem.Rows.Count = 0 Then
                    Me.InputToDgv_Proses(Me.TxtPrimaryRefDoc.Text, Me.TipeRefDoc)
                    If Me.TxtLocationSource.Text <> "" Then
                        Me.IsiQtyOnHandIsLocationControl()
                    End If
                    Exit Sub
                End If
                If Me.BufferDocRef <> Me.TxtRefDoc.Text Then
                    Me.InputToDgv_Proses(Me.TxtPrimaryRefDoc.Text, Me.TipeRefDoc)
                    If Me.TxtLocationSource.Text <> "" Then
                        Me.IsiQtyOnHandIsLocationControl()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindRefDoc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindRefDoc.Click
        Try
            If Me.TipeRefDoc = EnumRefDoc.Stock_Bundling_Item_After Then
                Me.FindByStockBundling(EnumMethod.Method_Find)
            ElseIf Me.TipeRefDoc = EnumRefDoc.Stock_Bundling_Item_Before Then
                Me.FindByStockBundling(EnumMethod.Method_Find)
            ElseIf Me.TipeRefDoc = EnumRefDoc.Stock_In_Adjust Then
                Me.FindByInOut(EnumMethod.Method_Find, ClsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust)
            ElseIf Me.TipeRefDoc = EnumRefDoc.Stock_In_Purchase Then
                Me.FindByInOut(EnumMethod.Method_Find, ClsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase)
            ElseIf Me.TipeRefDoc = EnumRefDoc.Stock_In_Sales_Return Then
                Me.FindByInOut(EnumMethod.Method_Find, ClsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn)
            ElseIf Me.TipeRefDoc = EnumRefDoc.Stock_Out_Adjust Then
                Me.FindByInOut(EnumMethod.Method_Find, ClsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust)
            ElseIf Me.TipeRefDoc = EnumRefDoc.Stock_Out_Purchase_Return Then
                Me.FindByInOut(EnumMethod.Method_Find, ClsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn)
            ElseIf Me.TipeRefDoc = EnumRefDoc.Stock_Out_Sales Then
                Me.FindByInOut(EnumMethod.Method_Find, ClsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales)
            ElseIf Me.TipeRefDoc = EnumRefDoc.Location_Transfer_Source Then
                Me.FindByLocationTransfer(EnumMethod.Method_Find)
                'ElseIf Me.TipeRefDoc = EnumRefDoc.Location_Transfer_Target Then
                '    Me.FindByLocationTransfer(EnumMethod.Method_Find)
                'ElseIf Me.TipeRefDoc = EnumRefDoc.Location Then
                '    Me.FindByLocation(EnumMethod.Method_Find)
                'ElseIf Me.TipeRefDoc = EnumRefDoc.Warehouse Then
                '    Me.FindByWarehouse(EnumMethod.Method_Find)
                'ElseIf Me.TipeRefDoc = EnumRefDoc.Delivery_Order Then
                '    Me.FindByDO(EnumMethod.Method_Find)
            End If

            Me.InputToDgv_Proses(Me.TxtPrimaryRefDoc.Text, Me.TipeRefDoc)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function FindByStockBundling(ByVal vMethod As EnumMethod) As Boolean
        Dim i As New FrmCariTransaksiStockBundling
        Dim Obj As New ClsTransaksiStockBundling

        Try
            If vMethod = EnumMethod.Method_Find Then
                If Me.TxtRefDoc.Text <> "" Then
                    i.txtCari.Text = Me.TxtRefDoc.Text
                End If

                i.TipeBundling = ClsTransaksiStockBundling.EnumTipeBundling.Purchase1
                i.ShowDialog()
                If i.PrimaryTransaksiStockBundling > 0 Then
                    Me.TxtPrimaryRefDoc.Text = i.PrimaryTransaksiStockBundling
                    Me.TxtRefDoc.Text = Obj.GetNoTransaksi(Me.TxtPrimaryRefDoc.Text)
                Else
                    Me.TxtRefDoc.Clear()
                    Me.TxtRefDoc.Focus()
                End If
            Else
                If Obj.ApakahSudahAda(Me.TxtRefDoc.Text) = True Then
                    Me.TxtPrimaryRefDoc.Text = Obj.GetPrimary(Me.TxtRefDoc.Text)
                    Me.TxtLocationSource.Focus()
                Else
                    If Me.TxtRefDoc.Text <> "" Then
                        i.txtCari.Text = Me.TxtRefDoc.Text
                    End If

                    i.TipeBundling = ClsTransaksiStockBundling.EnumTipeBundling.Purchase1
                    i.ShowDialog()
                    If i.PrimaryTransaksiStockBundling > 0 Then
                        Me.TxtPrimaryRefDoc.Text = i.PrimaryTransaksiStockBundling
                        Me.TxtRefDoc.Text = Obj.GetNoTransaksi(Me.TxtPrimaryRefDoc.Text)
                    Else
                        Me.TxtRefDoc.Clear()
                        Me.TxtPrimaryRefDoc.Text = "-1"
                        Me.TxtRefDoc.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Function

    Private Function FindByInOut(ByVal vMethod As EnumMethod, ByVal vTipeInOut As clsTransaksiStockInOut.EnumTipeStockInOut) As Boolean
        Dim i As New FrmCariTransaksiStockInOut
        Dim Obj As New clsTransaksiStockInOut

        Try
            If vMethod = EnumMethod.Method_Find Then
                If Me.TxtRefDoc.Text <> "" Then
                    i.txtCari.Text = Me.TxtRefDoc.Text
                End If

                i.JenisTransaksi = vTipeInOut
                i.ShowDialog()
                If i.PrimaryTransaksiStockInOut > 0 Then
                    Me.TxtPrimaryRefDoc.Text = i.PrimaryTransaksiStockInOut
                    Me.TxtRefDoc.Text = Obj.GetNoTransaksi(Me.TxtPrimaryRefDoc.Text)
                Else
                    Me.TxtRefDoc.Clear()
                    Me.TxtRefDoc.Focus()
                End If
            Else
                If Obj.ApakahSudahAda(Me.TxtRefDoc.Text) = True Then
                    Me.TxtPrimaryRefDoc.Text = Obj.GetPrimary(Me.TxtRefDoc.Text)
                    Me.TxtLocationSource.Focus()
                Else
                    If Me.TxtRefDoc.Text <> "" Then
                        i.txtCari.Text = Me.TxtRefDoc.Text
                    End If

                    i.JenisTransaksi = vTipeInOut
                    i.ShowDialog()
                    If i.PrimaryTransaksiStockInOut > 0 Then
                        Me.TxtPrimaryRefDoc.Text = i.PrimaryTransaksiStockInOut
                        Me.TxtRefDoc.Text = Obj.GetNoTransaksi(Me.TxtPrimaryRefDoc.Text)
                    Else
                        Me.TxtRefDoc.Clear()
                        Me.TxtPrimaryRefDoc.Text = "-1"
                        Me.TxtRefDoc.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Function

    Private Function FindByLocationTransfer(ByVal vMethod As EnumMethod) As Boolean
        Dim i As New FrmCariTransaksiStockTransfer
        Dim Obj As New clsTransaksiStockTransfer

        Try
            If vMethod = EnumMethod.Method_Find Then
                If Me.TxtRefDoc.Text <> "" Then
                    i.txtCari.Text = Me.TxtRefDoc.Text
                End If

                i.ShowDialog()
                If i.PrimaryTransaksiStockTransfer > 0 Then
                    Me.TxtPrimaryRefDoc.Text = i.PrimaryTransaksiStockTransfer
                    Me.TxtRefDoc.Text = Obj.GetNoTransaksi(Me.TxtPrimaryRefDoc.Text)
                Else
                    Me.TxtRefDoc.Clear()
                    Me.TxtRefDoc.Focus()
                End If
            Else
                If Obj.ApakahSudahAda(Me.TxtRefDoc.Text) = True Then
                    Me.TxtPrimaryRefDoc.Text = Obj.GetPrimary(Me.TxtRefDoc.Text)
                    Me.TxtLocationSource.Focus()
                Else
                    If Me.TxtRefDoc.Text <> "" Then
                        i.txtCari.Text = Me.TxtRefDoc.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryTransaksiStockTransfer > 0 Then
                        Me.TxtPrimaryRefDoc.Text = i.PrimaryTransaksiStockTransfer
                        Me.TxtRefDoc.Text = Obj.GetNoTransaksi(Me.TxtPrimaryRefDoc.Text)
                    Else
                        Me.TxtRefDoc.Clear()
                        Me.TxtPrimaryRefDoc.Text = "-1"
                        Me.TxtRefDoc.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Function

    Private Function GetPrimaryMin(Optional ByVal vLocationControl As Integer = 0) As Double
        Dim vPrimaryMin As Double = -1
        Try
            For Each row As DataGridViewRow In Me.dgvItem.Rows
                If row.Cells(Me.ColLocationControl.Index).Value = vLocationControl Then
                    If vPrimaryMin = -1 Then
                        vPrimaryMin = row.Cells(Me.ColPrimaryCodeNo.Index).Value
                    Else
                        If CDbl(row.Cells(Me.ColPrimaryCodeNo.Index).Value) < vPrimaryMin Then
                            vPrimaryMin = CDbl(row.Cells(Me.ColPrimaryCodeNo.Index).Value)
                        End If
                    End If
                End If
            Next
            Return vPrimaryMin
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetPrimaryMax(Optional ByVal vLocationControl As Integer = 0) As Double
        Dim vPrimaryMax As Double = -1
        Try
            For Each row As DataGridViewRow In Me.dgvItem.Rows
                If row.Cells(Me.ColLocationControl.Index).Value = vLocationControl Then
                    If vPrimaryMax = -1 Then
                        vPrimaryMax = row.Cells(Me.ColPrimaryCodeNo.Index).Value
                    Else
                        If CDbl(row.Cells(Me.ColPrimaryCodeNo.Index).Value) > vPrimaryMax Then
                            vPrimaryMax = CDbl(row.Cells(Me.ColPrimaryCodeNo.Index).Value)
                        End If
                    End If
                End If
            Next
            Return vPrimaryMax
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetQtyOnHand(ByVal tblProductQty As DseTransaksiStockAdvanceAdd.tblmaster_productqtyDataTable, Optional ByVal LocationControl As Integer = 0) As Boolean
        Dim r As DseTransaksiStockAdvanceAdd.tblmaster_productqtyRow
        Dim ExpiryDate As Date = CDate("1/1/1950")
        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each row As DataGridViewRow In Me.dgvItem.Rows
                If row.Cells(Me.ColExpiredDate.Index).Value = "-" Then
                    ExpiryDate = CDate("1/1/1950")
                Else
                    ExpiryDate = CDate(row.Cells(Me.ColExpiredDate.Index).Value)
                End If
                If row.Cells(Me.ColLocationControl.Index).Value = LocationControl Then
                    tblProductQty.DefaultView.RowFilter = "[Primary Product]='" & row.Cells(Me.ColPrimaryCodeNo.Index).Value & "' AND [Lot No]='" & row.Cells(Me.ColBatch.Index).Value & "'"
                    If tblProductQty.DefaultView.Count > 0 Then
                        r = tblProductQty.DefaultView(0).Row
                        row.Cells(Me.ColQtyOnHand.Index).Value = Me.ObjCD.DigitGrouping(r.Qty, r._Decimal)
                    Else
                        row.Cells(Me.ColQtyOnHand.Index).Value = Me.ObjCD.DigitGrouping(0, 0)
                    End If
                End If
            Next
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function IsiQtyOnHandIsLocationControl() As Boolean
        Dim dtaProductQty As New DseTransaksiStockAdvanceAddTableAdapters.tblmaster_productqtyTableAdapter
        Dim tblProductQty As New DseTransaksiStockAdvanceAdd.tblmaster_productqtyDataTable
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            dtaProductQty.Connection.ConnectionString = Core_Module.GetConstring
            If Me.TxtPrimaryLocationSource.Text <> -1 Then
                dtaProductQty.Fill(tblProductQty, Me.GetPrimaryMin(1), Me.GetPrimaryMax(1), Me.TxtPrimaryLocationSource.Text)
                'Baru (qty on hand + qty tbl item - qty tbl dgv)
                Me.HitungUlangQty(tblProductQty, Me.vTblDgv, Me.TipeRefDoc, Me.vPrimaryTransaksi)
            End If

            GetQtyOnHand(tblProductQty, 1)
            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub InputToDgv_Proses(ByVal vPrimaryTrans As Double, ByVal vRefDoc As EnumRefDoc)
        Dim dtaProductQty As New DseTransaksiStockAdvanceAddTableAdapters.tblmaster_productqtyTableAdapter
        Dim tblProductQty As New DseTransaksiStockAdvanceAdd.tblmaster_productqtyDataTable
        Try
            'If Me.BufferPrimaryRefDoc = vPrimaryTrans Then
            '    Exit Sub
            'End If

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.dgvItem.Rows.Clear()
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dtaProductQty.Connection.ConnectionString = Core_Module.GetConstring

            If vRefDoc = EnumRefDoc.Stock_Bundling_Item_After Then
                Dim dta As New DseTransaksiStockAdvanceAddTableAdapters.tblAfter1TableAdapter
                Dim tbl As New DseTransaksiStockAdvanceAdd.tblAfter1DataTable
                Dim row As DseTransaksiStockAdvanceAdd.tblAfter1Row

                dta.Connection.ConnectionString = Core_Module.GetConstring
                dta.Fill(tbl, vPrimaryTrans)
                For Each row In tbl
                    Me.InputToDgv_Add(vRefDoc, , , row, )
                Next
                dtaProductQty.Fill(tblProductQty, Me.GetPrimaryMin(0), Me.GetPrimaryMax(0), -1)
                GetQtyOnHand(tblProductQty, 0)

            ElseIf vRefDoc = EnumRefDoc.Stock_Bundling_Item_Before Then
                Dim dta As New DseTransaksiStockAdvanceAddTableAdapters.tblB41TableAdapter
                Dim tbl As New DseTransaksiStockAdvanceAdd.tblB41DataTable
                Dim row As DseTransaksiStockAdvanceAdd.tblB41Row
                dta.Connection.ConnectionString = Core_Module.GetConstring
                dta.Fill(tbl, vPrimaryTrans)
                For Each row In tbl
                    Me.InputToDgv_Add(vRefDoc, , row, , )
                Next
                dtaProductQty.Fill(tblProductQty, Me.GetPrimaryMin(0), Me.GetPrimaryMax(0), -1)
                GetQtyOnHand(tblProductQty, 0)

            ElseIf vRefDoc = EnumRefDoc.Stock_In_Adjust Or _
            vRefDoc = EnumRefDoc.Stock_In_Purchase Or _
            vRefDoc = EnumRefDoc.Stock_In_Sales_Return Then
                Dim dta As New DseTransaksiStockAdvanceAddTableAdapters.tblInOut1TableAdapter
                Dim tbl As New DseTransaksiStockAdvanceAdd.tblInOut1DataTable
                Dim row As DseTransaksiStockAdvanceAdd.tblInOut1Row

                dta.Connection.ConnectionString = Core_Module.GetConstring
                dta.Fill(tbl, vPrimaryTrans)
                For Each row In tbl
                    Me.InputToDgv_Add(vRefDoc, row, , , )
                Next
                dtaProductQty.Fill(tblProductQty, Me.GetPrimaryMin(0), Me.GetPrimaryMax(0), -1)
                GetQtyOnHand(tblProductQty, 0)

            ElseIf vRefDoc = EnumRefDoc.Stock_Out_Adjust Or _
            vRefDoc = EnumRefDoc.Stock_Out_Purchase_Return Or _
            vRefDoc = EnumRefDoc.Stock_Out_Sales Then
                Dim dta As New DseTransaksiStockAdvanceAddTableAdapters.tblInOut1TableAdapter
                Dim tbl As New DseTransaksiStockAdvanceAdd.tblInOut1DataTable
                Dim row As DseTransaksiStockAdvanceAdd.tblInOut1Row
                dta.Connection.ConnectionString = Core_Module.GetConstring
                dta.Fill(tbl, vPrimaryTrans)
                For Each row In tbl
                    Me.InputToDgv_Add(vRefDoc, row)
                Next
                dtaProductQty.Fill(tblProductQty, Me.GetPrimaryMin(0), Me.GetPrimaryMax(0), -1)
                GetQtyOnHand(tblProductQty, 0)

            ElseIf vRefDoc = EnumRefDoc.Location_Transfer_Source Or _
            vRefDoc = EnumRefDoc.Location_Transfer_Target Then
                Dim dta As New DseTransaksiStockAdvanceAddTableAdapters.tblTransfer1TableAdapter
                Dim tbl As New DseTransaksiStockAdvanceAdd.tblTransfer1DataTable
                Dim row As DseTransaksiStockAdvanceAdd.tblTransfer1Row
                dta.Connection.ConnectionString = Core_Module.GetConstring
                dta.Fill(tbl, vPrimaryTrans)
                For Each row In tbl
                    Me.InputToDgv_Add(vRefDoc, , , , row)
                Next
                dtaProductQty.Fill(tblProductQty, Me.GetPrimaryMin(0), Me.GetPrimaryMax(0), -1)
                GetQtyOnHand(tblProductQty, 0)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.BufferPrimaryRefDoc = vPrimaryTrans
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Function GetQtyOnHand_NoLocationControl() As Decimal
        Dim Retval As Decimal = 0
        Try

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahGabungBaris(ByVal vPrimaryProd As Double, ByVal vBatch As String, ByVal vQty As Decimal, ByVal vDec As Integer) As Boolean
        Dim rd As DataGridViewRow

        Try
            For i As Integer = 0 To Me.dgvItem.RowCount - 1
                rd = Me.dgvItem.Rows(i)
                If CDbl(rd.Cells(Me.ColPrimaryCodeNo.Index).Value) = vPrimaryProd And rd.Cells(Me.ColBatch.Index).Value = vBatch Then
                    rd.Cells(Me.ColQtyRef.Index).Value = Me.ObjCD.DigitGrouping(CDec(rd.Cells(Me.ColQtyRef.Index).Value) + vQty, vDec)
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub InputToDgv_Add(ByVal vRefDoc As EnumRefDoc, Optional ByVal rowInOut As DseTransaksiStockAdvanceAdd.tblInOut1Row = Nothing, Optional ByVal rowB4 As DseTransaksiStockAdvanceAdd.tblB41Row = Nothing, _
                               Optional ByVal rowAfter As DseTransaksiStockAdvanceAdd.tblAfter1Row = Nothing, Optional ByVal rowTransfer As DseTransaksiStockAdvanceAdd.tblTransfer1Row = Nothing)
        Dim rd As DataGridViewRow

        Try
            If vRefDoc = EnumRefDoc.Stock_Bundling_Item_After Then
                With rowAfter
                    If Me.ApakahGabungBaris(.Primary_Product, .Batch_No, .Qty, ._Decimal) = False Then
                        Me.dgvItem.Rows.Add()
                        rd = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                        rd.Cells(Me.ColNo.Index).Value = Me.dgvItem.RowCount
                        rd.Cells(Me.ColCodeNo.Index).Value = .Product_Code
                        rd.Cells(Me.ColDescription.Index).Value = .Description
                        rd.Cells(Me.ColQtyRef.Index).Value = Me.ObjCD.DigitGrouping(.Qty, ._Decimal)
                        rd.Cells(Me.ColQtyOnHand.Index).Value = Me.ObjCD.DigitGrouping(0, ._Decimal)
                        rd.Cells(Me.ColQtyAdd.Index).Value = 0

                        rd.Cells(Me.ColUnit.Index).Value = .Unit_Code
                        rd.Cells(Me.ColSecondaryQty.Index).Value = ""
                        rd.Cells(Me.ColSecondaryUnit.Index).Value = ""
                        rd.Cells(Me.ColBatch.Index).Value = .Batch_No
                        If .Expiry_Date = CDate("1/1/1950") Then
                            rd.Cells(Me.ColExpiredDate.Index).Value = "-"
                        Else
                            rd.Cells(Me.ColExpiredDate.Index).Value = Format(.Expiry_Date, "dd MMM yyyy")
                        End If
                        rd.Cells(Me.ColLocation.Index).Value = .Location
                        rd.Cells(Me.ColWareHouse.Index).Value = .Warehouse_Code
                        rd.Cells(Me.ColRemark.Index).Value = ""
                        rd.Cells(Me.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        rd.Cells(Me.ColPrimaryLocation.Index).Value = .Primary_Location
                        rd.Cells(Me.ColPrimaryUnits.Index).Value = .Primary_Unit
                        rd.Cells(Me.ColPrimarySecondaryUnits.Index).Value = "-1"
                        rd.Cells(Me.ColLocationControl.Index).Value = .Location_Controlled
                        rd.Cells(Me.ColDecimalDefault.Index).Value = ._Decimal
                        rd.Cells(Me.ColIsService.Index).Value = .Type
                    End If
                End With
            ElseIf vRefDoc = EnumRefDoc.Stock_Bundling_Item_Before Then
                With rowB4
                    If Me.ApakahGabungBaris(.Primary_Product, .Batch_No, .Qty, ._Decimal) = False Then
                        Me.dgvItem.Rows.Add()
                        rd = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                        rd.Cells(Me.ColNo.Index).Value = Me.dgvItem.RowCount
                        rd.Cells(Me.ColCodeNo.Index).Value = .Product_Code
                        rd.Cells(Me.ColDescription.Index).Value = .Description
                        rd.Cells(Me.ColQtyRef.Index).Value = Me.ObjCD.DigitGrouping(.Qty, ._Decimal)
                        rd.Cells(Me.ColQtyOnHand.Index).Value = Me.ObjCD.DigitGrouping(0, ._Decimal)
                        rd.Cells(Me.ColQtyAdd.Index).Value = 0
                        rd.Cells(Me.ColUnit.Index).Value = .Unit_Code
                        rd.Cells(Me.ColSecondaryQty.Index).Value = ""
                        rd.Cells(Me.ColSecondaryUnit.Index).Value = ""
                        rd.Cells(Me.ColBatch.Index).Value = .Batch_No
                        If .Expiry_Date = CDate("1/1/1950") Then
                            rd.Cells(Me.ColExpiredDate.Index).Value = "-"
                        Else
                            rd.Cells(Me.ColExpiredDate.Index).Value = Format(.Expiry_Date, "dd MMM yyyy")
                        End If
                        rd.Cells(Me.ColLocation.Index).Value = .Location
                        rd.Cells(Me.ColWareHouse.Index).Value = .Warehouse_Code
                        rd.Cells(Me.ColRemark.Index).Value = ""
                        rd.Cells(Me.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        rd.Cells(Me.ColPrimaryLocation.Index).Value = .Primary_Storage_Area
                        rd.Cells(Me.ColPrimaryUnits.Index).Value = .Primary_Unit
                        rd.Cells(Me.ColPrimarySecondaryUnits.Index).Value = "-1"
                        rd.Cells(Me.ColLocationControl.Index).Value = .Location_Controlled
                        rd.Cells(Me.ColDecimalDefault.Index).Value = ._Decimal
                        rd.Cells(Me.ColIsService.Index).Value = .Type
                    End If
                End With
            ElseIf vRefDoc = EnumRefDoc.Stock_In_Adjust Or vRefDoc = EnumRefDoc.Stock_In_Purchase Or vRefDoc = EnumRefDoc.Stock_In_Sales_Return Then
                With rowInOut
                    If Me.ApakahGabungBaris(.Primary_Product, .Batch_No, .Qty, ._Decimal) = False Then
                        Me.dgvItem.Rows.Add()
                        rd = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                        rd.Cells(Me.ColNo.Index).Value = Me.dgvItem.RowCount
                        rd.Cells(Me.ColCodeNo.Index).Value = .Product_Code
                        rd.Cells(Me.ColDescription.Index).Value = .Description
                        rd.Cells(Me.ColQtyRef.Index).Value = Me.ObjCD.DigitGrouping(.Qty, ._Decimal)
                        rd.Cells(Me.ColQtyOnHand.Index).Value = Me.ObjCD.DigitGrouping(0, ._Decimal)
                        rd.Cells(Me.ColQtyAdd.Index).Value = 0
                        rd.Cells(Me.ColUnit.Index).Value = .Unit_Code
                        rd.Cells(Me.ColSecondaryQty.Index).Value = ""
                        rd.Cells(Me.ColSecondaryUnit.Index).Value = ""
                        rd.Cells(Me.ColBatch.Index).Value = .Batch_No
                        If .Expiry_Date = CDate("1/1/1950") Then
                            rd.Cells(Me.ColExpiredDate.Index).Value = "-"
                        Else
                            rd.Cells(Me.ColExpiredDate.Index).Value = Format(.Expiry_Date, "dd MMM yyyy")
                        End If
                        rd.Cells(Me.ColLocation.Index).Value = .Location
                        rd.Cells(Me.ColWareHouse.Index).Value = .Warehouse_Code
                        rd.Cells(Me.ColRemark.Index).Value = ""
                        rd.Cells(Me.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        rd.Cells(Me.ColPrimaryLocation.Index).Value = .Primary_Storage_Area
                        rd.Cells(Me.ColPrimaryUnits.Index).Value = .Primary_Unit
                        rd.Cells(Me.ColPrimarySecondaryUnits.Index).Value = "-1"
                        rd.Cells(Me.ColLocationControl.Index).Value = .Location_Controlled
                        rd.Cells(Me.ColDecimalDefault.Index).Value = ._Decimal
                        rd.Cells(Me.ColIsService.Index).Value = .Type
                    End If
                End With
            ElseIf vRefDoc = EnumRefDoc.Stock_Out_Adjust Or vRefDoc = EnumRefDoc.Stock_Out_Purchase_Return Or vRefDoc = EnumRefDoc.Stock_Out_Sales Then
                With rowInOut
                    If Me.ApakahGabungBaris(.Primary_Product, .Batch_No, .Qty, ._Decimal) = False Then
                        Me.dgvItem.Rows.Add()
                        rd = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                        rd.Cells(Me.ColNo.Index).Value = Me.dgvItem.RowCount
                        rd.Cells(Me.ColCodeNo.Index).Value = .Product_Code
                        rd.Cells(Me.ColDescription.Index).Value = .Description
                        rd.Cells(Me.ColQtyRef.Index).Value = Me.ObjCD.DigitGrouping(.Qty, ._Decimal)
                        rd.Cells(Me.ColQtyOnHand.Index).Value = Me.ObjCD.DigitGrouping(0, ._Decimal)
                        rd.Cells(Me.ColQtyAdd.Index).Value = 0
                        rd.Cells(Me.ColUnit.Index).Value = .Unit_Code
                        rd.Cells(Me.ColSecondaryQty.Index).Value = ""
                        rd.Cells(Me.ColSecondaryUnit.Index).Value = ""
                        rd.Cells(Me.ColBatch.Index).Value = .Batch_No
                        If .Expiry_Date = CDate("1/1/1950") Then
                            rd.Cells(Me.ColExpiredDate.Index).Value = "-"
                        Else
                            rd.Cells(Me.ColExpiredDate.Index).Value = Format(.Expiry_Date, "dd MMM yyyy")
                        End If
                        rd.Cells(Me.ColLocation.Index).Value = .Location
                        rd.Cells(Me.ColWareHouse.Index).Value = .Warehouse_Code
                        rd.Cells(Me.ColRemark.Index).Value = ""
                        rd.Cells(Me.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        rd.Cells(Me.ColPrimaryLocation.Index).Value = .Primary_Storage_Area
                        rd.Cells(Me.ColPrimaryUnits.Index).Value = .Primary_Unit
                        rd.Cells(Me.ColPrimarySecondaryUnits.Index).Value = "-1"
                        rd.Cells(Me.ColLocationControl.Index).Value = .Location_Controlled
                        rd.Cells(Me.ColDecimalDefault.Index).Value = ._Decimal
                        rd.Cells(Me.ColIsService.Index).Value = .Type
                    End If
                End With
            ElseIf vRefDoc = EnumRefDoc.Location_Transfer_Source Then
                With rowTransfer
                    If Me.ApakahGabungBaris(.Primary_Product, .Batch_No, .Qty, ._Decimal) = False Then
                        Me.dgvItem.Rows.Add()
                        rd = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                        rd.Cells(Me.ColNo.Index).Value = Me.dgvItem.RowCount
                        rd.Cells(Me.ColCodeNo.Index).Value = .Product_Code
                        rd.Cells(Me.ColDescription.Index).Value = .Description
                        rd.Cells(Me.ColQtyRef.Index).Value = Me.ObjCD.DigitGrouping(.Qty, ._Decimal)
                        rd.Cells(Me.ColQtyOnHand.Index).Value = Me.ObjCD.DigitGrouping(0, ._Decimal)
                        rd.Cells(Me.ColQtyAdd.Index).Value = 0
                        rd.Cells(Me.ColUnit.Index).Value = .Unit_Code
                        rd.Cells(Me.ColSecondaryQty.Index).Value = ""
                        rd.Cells(Me.ColSecondaryUnit.Index).Value = ""
                        rd.Cells(Me.ColBatch.Index).Value = .Batch_No
                        If .Expiry_Date = CDate("1/1/1950") Then
                            rd.Cells(Me.ColExpiredDate.Index).Value = "-"
                        Else
                            rd.Cells(Me.ColExpiredDate.Index).Value = Format(.Expiry_Date, "dd MMM yyyy")
                        End If
                        rd.Cells(Me.ColLocation.Index).Value = .Location
                        rd.Cells(Me.ColWareHouse.Index).Value = .Warehouse
                        rd.Cells(Me.ColRemark.Index).Value = ""
                        rd.Cells(Me.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        rd.Cells(Me.ColPrimaryLocation.Index).Value = .Primary_Storage_Area
                        rd.Cells(Me.ColPrimaryUnits.Index).Value = .Primary_Units
                        rd.Cells(Me.ColPrimarySecondaryUnits.Index).Value = "-1"
                        rd.Cells(Me.ColLocationControl.Index).Value = .Location_Controlled
                        rd.Cells(Me.ColDecimalDefault.Index).Value = ._Decimal
                        rd.Cells(Me.ColIsService.Index).Value = .Type
                    End If
                End With
            ElseIf vRefDoc = EnumRefDoc.Location_Transfer_Target Then
                With rowTransfer
                    If Me.ApakahGabungBaris(.Primary_Product, .Batch_No, .Qty, ._Decimal) = False Then
                        Me.dgvItem.Rows.Add()
                        rd = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                        rd.Cells(Me.ColNo.Index).Value = Me.dgvItem.RowCount
                        rd.Cells(Me.ColCodeNo.Index).Value = .Product_Code
                        rd.Cells(Me.ColDescription.Index).Value = .Description
                        rd.Cells(Me.ColQtyRef.Index).Value = Me.ObjCD.DigitGrouping(.Qty, ._Decimal)
                        rd.Cells(Me.ColQtyOnHand.Index).Value = Me.ObjCD.DigitGrouping(0, ._Decimal)
                        rd.Cells(Me.ColQtyAdd.Index).Value = 0
                        rd.Cells(Me.ColUnit.Index).Value = .Unit_Code
                        rd.Cells(Me.ColSecondaryQty.Index).Value = ""
                        rd.Cells(Me.ColSecondaryUnit.Index).Value = ""
                        rd.Cells(Me.ColBatch.Index).Value = .Batch_No
                        If .Expiry_Date = CDate("1/1/1950") Then
                            rd.Cells(Me.ColExpiredDate.Index).Value = "-"
                        Else
                            rd.Cells(Me.ColExpiredDate.Index).Value = Format(.Expiry_Date, "dd MMM yyyy")
                        End If
                        rd.Cells(Me.ColLocation.Index).Value = .Location
                        rd.Cells(Me.ColWareHouse.Index).Value = .Warehouse
                        rd.Cells(Me.ColRemark.Index).Value = ""
                        rd.Cells(Me.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        rd.Cells(Me.ColPrimaryLocation.Index).Value = .Primary_Storage_Area
                        rd.Cells(Me.ColPrimaryUnits.Index).Value = .Primary_Units
                        rd.Cells(Me.ColPrimarySecondaryUnits.Index).Value = "-1"
                        rd.Cells(Me.ColLocationControl.Index).Value = .Location_Controlled
                        rd.Cells(Me.ColDecimalDefault.Index).Value = ._Decimal
                        rd.Cells(Me.ColIsService.Index).Value = .Type
                    End If
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboRefDoc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRefDoc.GotFocus
        Me.BufferRefDoc = Me.cboRefDoc.SelectedIndex
    End Sub

    Private Sub AturCboTransfer()
        Try
            Me.TxtLocationTarget.Enabled = False
            Me.BtnFindLocationTarget.Enabled = False
            If Me.ParentForm1 = EnumRefDoc.Location_Transfer_Source Then
                Me.TxtLocationTarget.Enabled = True
                Me.BtnFindLocationTarget.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboRefDoc_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRefDoc.SelectedIndexChanged
        Dim ApaKosongkanDgv As Boolean = False
        Try
            Me.AturCboTransfer()
            If ProsesUbahCBOrefdoc = True Then
                If Me.cboRefDoc.SelectedIndex = EnumRefDoc.Stock_Bundling_Item_After Then
                    Me.TipeRefDoc = EnumRefDoc.Stock_Bundling_Item_After
                    ApaKosongkanDgv = True
                ElseIf Me.cboRefDoc.SelectedIndex = EnumRefDoc.Stock_Bundling_Item_Before Then
                    Me.TipeRefDoc = EnumRefDoc.Stock_Bundling_Item_Before
                    ApaKosongkanDgv = True
                ElseIf Me.cboRefDoc.SelectedIndex = EnumRefDoc.Stock_In_Adjust Then
                    Me.TipeRefDoc = EnumRefDoc.Stock_In_Adjust
                    ApaKosongkanDgv = True
                ElseIf Me.cboRefDoc.SelectedIndex = EnumRefDoc.Stock_In_Purchase Then
                    Me.TipeRefDoc = EnumRefDoc.Stock_In_Purchase
                    ApaKosongkanDgv = True
                ElseIf Me.cboRefDoc.SelectedIndex = EnumRefDoc.Stock_In_Sales_Return Then
                    Me.TipeRefDoc = EnumRefDoc.Stock_In_Sales_Return
                    ApaKosongkanDgv = True
                ElseIf Me.cboRefDoc.SelectedIndex = EnumRefDoc.Stock_Out_Adjust Then
                    Me.TipeRefDoc = EnumRefDoc.Stock_Out_Adjust
                    ApaKosongkanDgv = True
                ElseIf Me.cboRefDoc.SelectedIndex = EnumRefDoc.Stock_Out_Purchase_Return Then
                    Me.TipeRefDoc = EnumRefDoc.Stock_Out_Purchase_Return
                    ApaKosongkanDgv = True
                ElseIf Me.cboRefDoc.SelectedIndex = EnumRefDoc.Stock_Out_Sales Then
                    Me.TipeRefDoc = EnumRefDoc.Stock_Out_Sales
                    ApaKosongkanDgv = True
                ElseIf Me.cboRefDoc.SelectedIndex = EnumRefDoc.Location_Transfer_Source Then
                    Me.TipeRefDoc = EnumRefDoc.Location_Transfer_Source
                    ApaKosongkanDgv = True
                End If

                If Me.dgvItem.RowCount > 0 And ApaKosongkanDgv = True Then
                    Dim BufferRefDoc_2 As EnumRefDoc = Me.BufferRefDoc
                    If MsgBox("Change reference document will delete the product item. Continue?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                        Me.TxtPrimaryRefDoc.Text = -1
                        Me.TxtRefDoc.Clear()
                        Me.dgvItem.Rows.Clear()
                    Else
                        ProsesUbahCBOrefdoc = False
                        Me.TipeRefDoc = BufferRefDoc_2
                        Me.cboRefDoc.SelectedIndex = Me.TipeRefDoc
                        Exit Sub
                    End If
                End If
            End If

            ProsesUbahCBOrefdoc = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim Benar As Boolean = False
        Try
            Me.ColBatch.Visible = True
            If String.IsNullOrEmpty(Me.TxtRefDoc.Text) = True Then
                MsgBox("Please specify reference document.", MsgBoxStyle.Information, "Information")
                Me.TxtRefDoc.Focus()
                Exit Sub
            End If
            If String.IsNullOrEmpty(Me.TxtLocationSource.Text) = True Then
                MsgBox("Please specify source location.", MsgBoxStyle.Information, "Information")
                Me.TxtLocationSource.Focus()
                Exit Sub
            End If
            If Me.TxtLocationTarget.Enabled = True Then
                If String.IsNullOrEmpty(Me.TxtLocationTarget.Text) = True Then
                    MsgBox("Please specify target location.", MsgBoxStyle.Information, "Information")
                    Me.TxtLocationTarget.Focus()
                    Exit Sub
                End If
            End If

            For Each ROW As DataGridViewRow In Me.dgvItem.Rows
                If String.IsNullOrEmpty(ROW.Cells(Me.ColQtyAdd.Index).Value) = False Then
                    If ObjCD.CekBilanganPositif(ROW.Cells(Me.ColQtyAdd.Index).Value) = True Then
                        Benar = True
                    End If
                End If

                If ParentForm1 = EnumRefDoc.Stock_Bundling_Item_After Or ParentForm1 = EnumRefDoc.Stock_In_Adjust Or ParentForm1 = EnumRefDoc.Location_Transfer_Target Then
                    If ROW.Cells(Me.ColLocationControl.Index).Value = 1 Then
                        If Me.TxtLocationSource.Text = "" Then
                            MsgBox("Please specify location.", MsgBoxStyle.Information, "Information")
                            Me.TxtLocationSource.Focus()
                            Exit Sub
                        End If
                    End If
                End If

              
            Next

            If Benar = False Then
                MsgBox("Please specify at least 1 item with qty > 0.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If Me.dgvItem.Rows.Count = 0 Then
                MsgBox("No data item detected.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            Me.FromOtherButton = False

            If Me.ParentForm1 = EnumRefDoc.Stock_Bundling_Item_Before Or Me.ParentForm1 = EnumRefDoc.Stock_Out_Adjust Or _
            Me.ParentForm1 = EnumRefDoc.Location_Transfer_Source Then
                For Each rd As DataGridViewRow In Me.dgvItem.Rows
                    If IsNumeric(rd.Cells(Me.ColQtyOnHand.Index).Value) = True And IsNumeric(rd.Cells(Me.ColQtyAdd.Index).Value) = True Then
                        If CDec(rd.Cells(Me.ColQtyOnHand.Index).Value) < CDec(rd.Cells(Me.ColQtyAdd.Index).Value) Then
                            MsgBox("Qty must be smaller than qty on hand.", MsgBoxStyle.Information, "Information")
                            rd.Selected = True
                            Exit Sub
                        End If
                    End If
                Next
            End If

            If MsgBox("Add item data?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                If Me.rboAdd.Checked = True Then
                    Me.InputType = ClsTransaksiBarcode.EnumInputType.Add
                Else
                    Me.InputType = ClsTransaksiBarcode.EnumInputType.Replace
                End If

                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.IsiKeTbl()

                Me.FromOtherButton = True
                Me.CancelClose = False
                Me.Close()
            Else
                Me.FromOtherButton = True
                Me.CancelClose = True
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataItemBenar() As Boolean
        Try
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub IsiKeTbl()
        Dim row As Core_Module.DseBarcodeInput.DataTable1Row
        Dim rd As DataGridViewRow

        Try
            'isi data item barcode pada dgv form barcode input ke table
            For Each rd In Me.dgvItem.Rows
                If rd.Cells(Me.ColQtyAdd.Index).Value > 0 Then
                    row = vTBL.NewDataTable1Row
                    With row
                        .Primary_Product = rd.Cells(Me.ColPrimaryCodeNo.Index).Value
                        .Primary_Location = Me.TxtPrimaryLocationSource.Text
                        .Primary_Location_Target = Me.TxtPrimaryLocationTarget.Text
                        .Primary_Unit = rd.Cells(Me.ColPrimaryUnits.Index).Value
                        .Code = rd.Cells(Me.ColCodeNo.Index).Value
                        .Description = rd.Cells(Me.ColDescription.Index).Value
                        .Qty = rd.Cells(Me.ColQtyAdd.Index).Value
                        .Unit = rd.Cells(Me.ColUnit.Index).Value
                        .Batch = rd.Cells(Me.ColBatch.Index).Value
                        If rd.Cells(Me.ColExpiredDate.Index).Value = "-" Then
                            .Expired_Date = CDate("1/1/1950")
                        Else
                            .Expired_Date = CDate(rd.Cells(Me.ColExpiredDate.Index).Value)
                        End If

                        .Location = Me.TxtLocationSource.Text
                        .Locatian_Target = Me.TxtLocationTarget.Text
                        .Warehouse = Me.TxtWarehouseSource.Text
                        .Warehouse_Target = Me.TxtWarehouseTarget.Text
                        .Remarks = rd.Cells(Me.ColRemark.Index).Value
                        ._Decimal = rd.Cells(Me.ColDecimalDefault.Index).Value
                        .Location_Controlled = rd.Cells(Me.ColLocationControl.Index).Value
                    End With
                    vTBL.Rows.Add(row)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColQtyAdd.Index Then
                If Me.ParentForm1 = EnumRefDoc.Stock_Bundling_Item_After Or Me.ParentForm1 = EnumRefDoc.Stock_Bundling_Item_Before Or _
                    Me.ParentForm1 = EnumRefDoc.Location_Transfer_Source Or Me.ParentForm1 = EnumRefDoc.Location_Transfer_Target Then
                    If Me.dgvItem.Item(Me.ColIsService.Index, e.RowIndex).Value = 1 Then
                        e.Cancel = True
                        MsgBox("For stock transfer or stock bundling transaction, you can only specify goods.")
                        Exit Sub
                    Else
                        Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            Else
                If e.ColumnIndex <> Me.ColRemark.Index Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Try
            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColQtyAdd.Index Then
                    If ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please enter a non negative number.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    Else
                        If ParentForm1 = EnumRefDoc.Stock_Bundling_Item_After Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))

                        ElseIf ParentForm1 = EnumRefDoc.Stock_Bundling_Item_Before Then
                            If CDec(Me.dgvItem.Item(Me.ColQtyOnHand.Index, e.RowIndex).Value) < CDec(Me.dgvItem.Item(Me.ColQtyRef.Index, e.RowIndex).Value) Then
                                If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > CDec(Me.dgvItem.Item(Me.ColQtyOnHand.Index, e.RowIndex).Value) Then
                                    MsgBox("Qty must be smaller than qty on hand.", MsgBoxStyle.Information, "Information")
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            Else
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))
                            End If

                        ElseIf ParentForm1 = EnumRefDoc.Stock_In_Adjust Or _
                               ParentForm1 = EnumRefDoc.Stock_In_Purchase Or _
                               ParentForm1 = EnumRefDoc.Stock_In_Sales_Return Then

                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))

                        ElseIf ParentForm1 = EnumRefDoc.Stock_Out_Adjust Or _
                              ParentForm1 = EnumRefDoc.Stock_Out_Purchase_Return Or _
                              ParentForm1 = EnumRefDoc.Stock_Out_Sales Then

                            If CDec(Me.dgvItem.Item(Me.ColQtyOnHand.Index, e.RowIndex).Value) < CDec(Me.dgvItem.Item(Me.ColQtyRef.Index, e.RowIndex).Value) Then
                                If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > CDec(Me.dgvItem.Item(Me.ColQtyOnHand.Index, e.RowIndex).Value) Then
                                    MsgBox("Qty must be smaller than qty on hand.", MsgBoxStyle.Information, "Information")
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            Else
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))
                            End If

                        ElseIf ParentForm1 = EnumRefDoc.Location_Transfer_Source Then
                            If CDec(Me.dgvItem.Item(Me.ColQtyOnHand.Index, e.RowIndex).Value) < CDec(Me.dgvItem.Item(Me.ColQtyRef.Index, e.RowIndex).Value) Then
                                If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > CDec(Me.dgvItem.Item(Me.ColQtyOnHand.Index, e.RowIndex).Value) Then
                                    MsgBox("Qty must be smaller than qty on hand.", MsgBoxStyle.Information, "Information")
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            Else
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAddAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles llbAddAll.Click
        Dim Hasil As New MsgBoxResult
        Try
            If Me.dgvItem.Rows.Count = 0 Then
                Exit Sub
            End If
            Hasil = MsgBox("Add all item?", MsgBoxStyle.OkCancel, "Confrm")
            If Hasil = MsgBoxResult.Ok Then
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                For Each row As DataGridViewRow In Me.dgvItem.Rows
                    If ParentForm1 = EnumRefDoc.Stock_Bundling_Item_Before Then
                        If row.Cells(Me.ColIsService.Index).Value = 0 Then
                            If CDec(Me.dgvItem.Item(Me.ColQtyOnHand.Index, row.Index).Value) < CDec(Me.dgvItem.Item(Me.ColQtyRef.Index, row.Index).Value) Then
                                row.Cells(Me.ColQtyAdd.Index).Value = Me.dgvItem.Item(Me.ColQtyOnHand.Index, row.Index).Value
                            Else
                                row.Cells(Me.ColQtyAdd.Index).Value = Me.dgvItem.Item(Me.ColQtyRef.Index, row.Index).Value
                            End If
                        End If
                    ElseIf ParentForm1 = EnumRefDoc.Stock_Bundling_Item_After Then
                        If row.Cells(Me.ColIsService.Index).Value = 0 Then
                            row.Cells(Me.ColQtyAdd.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColQtyRef.Index).Value, ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, row.Index).Value))
                        End If
                    ElseIf ParentForm1 = EnumRefDoc.Stock_In_Adjust Or _
                    Me.ParentForm1 = EnumRefDoc.Stock_In_Purchase Or _
                    Me.ParentForm1 = EnumRefDoc.Stock_In_Sales_Return Then
                        row.Cells(Me.ColQtyAdd.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColQtyRef.Index).Value, ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, row.Index).Value))
                    ElseIf ParentForm1 = EnumRefDoc.Stock_Out_Adjust Or _
                           ParentForm1 = EnumRefDoc.Stock_Out_Purchase_Return Or _
                           ParentForm1 = EnumRefDoc.Stock_Out_Sales Then
                        If CDec(Me.dgvItem.Item(Me.ColQtyOnHand.Index, row.Index).Value) < CDec(Me.dgvItem.Item(Me.ColQtyRef.Index, row.Index).Value) Then
                            row.Cells(Me.ColQtyAdd.Index).Value = Me.dgvItem.Item(Me.ColQtyOnHand.Index, row.Index).Value
                        Else
                            row.Cells(Me.ColQtyAdd.Index).Value = Me.dgvItem.Item(Me.ColQtyRef.Index, row.Index).Value
                        End If
                    ElseIf ParentForm1 = EnumRefDoc.Location_Transfer_Source Then
                        If row.Cells(Me.ColIsService.Index).Value = 0 Then
                            If CDec(Me.dgvItem.Item(Me.ColQtyOnHand.Index, row.Index).Value) < CDec(Me.dgvItem.Item(Me.ColQtyRef.Index, row.Index).Value) Then
                                row.Cells(Me.ColQtyAdd.Index).Value = Me.dgvItem.Item(Me.ColQtyOnHand.Index, row.Index).Value
                            Else
                                row.Cells(Me.ColQtyAdd.Index).Value = Me.dgvItem.Item(Me.ColQtyRef.Index, row.Index).Value
                            End If
                        End If
                    End If
                Next
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Windows.Forms.Cursor.Current = Cursors.Default
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbRemoveAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles llbRemoveAll.Click
        Dim Hasil As New MsgBoxResult
        Try
            If Me.dgvItem.Rows.Count = 0 Then
                Exit Sub
            End If
            Hasil = MsgBox("Set qty=0 to all data?", MsgBoxStyle.OkCancel, "Confrm")
            If Hasil = MsgBoxResult.Ok Then
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                For Each row As DataGridViewRow In Me.dgvItem.Rows
                    row.Cells(Me.ColQtyAdd.Index).Value = Me.ObjCD.DigitGrouping(0)
                Next
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.FromOtherButton = False
            If MsgBox("Cancel input and close form?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                Me.FromOtherButton = True
                Me.CancelClose = False
                Me.Close()
            Else
                Me.FromOtherButton = True
                Me.CancelClose = True
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRefDoc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRefDoc.GotFocus
        Try
            Me.TxtRefDoc.BackColor = Me.BackColorAll

            BufferDocRef = TxtRefDoc.Text

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRefDoc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRefDoc.LostFocus
        Try
            Me.TxtRefDoc.BackColor = Me.TxtPrimaryRefDoc.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocationSource_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocationSource.GotFocus
        Try
            Me.TxtLocationSource.BackColor = Me.BackColorAll
            BufferDocRef = TxtLocationSource.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocationSource_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocationSource.LostFocus
        Try
            Me.TxtLocationSource.BackColor = Me.TxtPrimaryRefDoc.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocationTarget_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocationTarget.GotFocus
        Try
            If TxtLocationTarget.Enabled = True Then
                Me.TxtLocationTarget.BackColor = Me.BackColorAll
            End If
            Me.TxtLocationTarget.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocationTarget_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocationTarget.LostFocus
        Try
            If TxtLocationTarget.Enabled = True Then
                Me.TxtLocationTarget.BackColor = Me.TxtPrimaryRefDoc.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RemoveAllQtyOnLocationControl()
        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For Each row As DataGridViewRow In Me.dgvItem.Rows
                If row.Cells(Me.ColLocationControl.Index).Value = 1 Then
                    row.Cells(Me.ColQtyOnHand.Index).Value = 0
                End If
            Next
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAddByQtyOnHand_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAddByQtyOnHand.LinkClicked
        Dim Hasil As New MsgBoxResult

        Try
            If Me.dgvItem.Rows.Count = 0 Then
                Exit Sub
            End If
            Hasil = MsgBox("Add by qty on hand?", MsgBoxStyle.OkCancel, "Confrm")
            If Hasil = MsgBoxResult.Ok Then
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                For Each row As DataGridViewRow In Me.dgvItem.Rows
                    row.Cells(Me.ColQtyAdd.Index).Value = row.Cells(Me.ColQtyOnHand.Index).Value
                Next
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Windows.Forms.Cursor.Current = Cursors.Default
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub HitungUlangQty(ByRef tblQtyOnHand As DseTransaksiStockAdvanceAdd.tblmaster_productqtyDataTable, _
    ByVal vTBLDgv As Core_Module.DseBarcodeInput.tbl_qty_itemDataTable, _
    ByVal vTipeForm As FrmTransaksiStockAdvanceAdd.EnumRefDoc, ByVal vPrimaryTrans As Double)
        Dim rowQtyOnHand As DseTransaksiStockAdvanceAdd.tblmaster_productqtyRow
        Dim vTBLDgvDum As New Core_Module.DseBarcodeInput.tbl_qty_itemDataTable
        Dim rowDgv, rowDgvDum As Core_Module.DseBarcodeInput.tbl_qty_itemRow
        Dim DtaDB As New Core_Module.DseBarcodeInputTableAdapters.tbl_qty_itemTableAdapter
        Dim TblDB As New Core_Module.DseBarcodeInput.tbl_qty_itemDataTable
        Dim rowDB As Core_Module.DseBarcodeInput.tbl_qty_itemRow
        Dim IdxR As Integer = 0

        Try
            For Each rowDgv In vTBLDgv
                rowDgvDum = vTBLDgvDum.Newtbl_qty_itemRow
                For i As Integer = 0 To vTBLDgv.Columns.Count - 1
                    rowDgvDum.Item(i) = rowDgv.Item(i)
                Next
                vTBLDgvDum.Rows.Add(rowDgvDum)
            Next

            'fungsi get qty db - qty dgv
            If vTipeForm = FrmTransaksiStockAdvanceAdd.EnumRefDoc.Location_Transfer_Source Then
                DtaDB.FillBy_StockTransfer(TblDB, vPrimaryTrans)
            ElseIf vTipeForm = FrmTransaksiStockAdvanceAdd.EnumRefDoc.Stock_Bundling_Item_Before Then
                DtaDB.FillBy_BundBef(TblDB, vPrimaryTrans)
            ElseIf vTipeForm = FrmTransaksiStockAdvanceAdd.EnumRefDoc.Stock_Bundling_Item_After Then
                DtaDB.FillBy_BundAft(TblDB, vPrimaryTrans)
            ElseIf vTipeForm = FrmTransaksiStockAdvanceAdd.EnumRefDoc.Stock_In_Adjust Or vTipeForm = FrmTransaksiStockAdvanceAdd.EnumRefDoc.Stock_Out_Adjust Then
                DtaDB.FillBy_InOut(TblDB, vPrimaryTrans)
            End If

            TblDB.DefaultView.RowFilter = ""
            TblDB.DefaultView.Sort = "[Primary Product], [Primary Location], [Batch No]"
            vTBLDgvDum.DefaultView.RowFilter = ""
            vTBLDgvDum.DefaultView.Sort = "[Primary Product], [Primary Location], [Batch No]"

            IdxR = 0
            For i As Integer = 0 To TblDB.DefaultView.Count - 1
                rowDB = TblDB.DefaultView(i).Row

                While IdxR <= vTBLDgvDum.DefaultView.Count - 1
                    rowDgv = vTBLDgvDum.DefaultView(IdxR).Row
                    If rowDB.Primary_Product < rowDgv.Primary_Product Then
                        Exit While
                    Else
                        If rowDB.Primary_Product = rowDgv.Primary_Product And rowDB.Primary_Location = rowDgv.Primary_Location And rowDB.Batch_No = rowDgv.Batch_No Then
                            rowDB.Qty -= rowDgv.Qty
                            vTBLDgvDum.Rows.Remove(rowDgv)
                            Exit While
                        End If
                    End If

                    IdxR += 1
                End While
            Next

            If vTBLDgvDum.Rows.Count > 0 Then
                For i As Integer = 0 To vTBLDgvDum.Rows.Count - 1
                    rowDgv = vTBLDgvDum.Rows(i)
                    rowDB = TblDB.Newtbl_qty_itemRow
                    rowDB.Primary_Product = rowDgv.Primary_Product
                    rowDB.Primary_Location = rowDgv.Primary_Location
                    rowDB.Batch_No = rowDgv.Batch_No
                    rowDB.Qty = (0 - rowDgv.Qty)
                    TblDB.Rows.Add(rowDB)
                Next
            End If
            'tutup fungsi get qty db - qty dgv

            'fungsi set qty on hand + qty db
            tblQtyOnHand.DefaultView.RowFilter = ""
            tblQtyOnHand.DefaultView.Sort = "[Primary Product], [Primary Storage Area], [Lot No]"
            TblDB.DefaultView.RowFilter = ""
            TblDB.DefaultView.Sort = "[Primary Product], [Primary Location], [Batch No]"
            For i As Integer = 0 To tblQtyOnHand.DefaultView.Count - 1
                rowQtyOnHand = tblQtyOnHand.DefaultView(i).Row

                IdxR = 0
                While IdxR <= TblDB.DefaultView.Count - 1
                    rowDB = TblDB.DefaultView(IdxR).Row
                    If rowQtyOnHand.Primary_Product < rowDB.Primary_Product Then
                        Exit While
                    ElseIf rowQtyOnHand.Primary_Product = rowDB.Primary_Product Then
                        If rowQtyOnHand.Primary_Product = rowDB.Primary_Product And rowQtyOnHand.Primary_Storage_Area = rowDB.Primary_Location And rowQtyOnHand.Lot_No = rowDB.Batch_No Then
                            rowQtyOnHand.Qty += rowDB.Qty
                            TblDB.Rows.Remove(rowDB)
                            Exit While
                        End If
                    ElseIf rowQtyOnHand.Primary_Product > rowDB.Primary_Product Then
                        TblDB.Rows.Remove(rowDB)
                        IdxR -= 1
                    End If

                    IdxR += 1
                End While
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class