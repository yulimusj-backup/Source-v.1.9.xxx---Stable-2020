Option Explicit On
Option Strict Off

Public Class FrmLaporanProd_PickList_TransDetail
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Public TblDummyList As DseLaporanProd_PickList.Tbl_DummyListDataTable
    Public TblDetail As DseLaporanProd_PickList.tbldetailDataTable
    Public TblProductQty As DseLaporanProd_PickList.tblmaster_productqtyDataTable
    Public TblPriority As DseLaporanProd_PickList.tblmaster_storage_areaDataTable
    Public TblExclude As DseLaporanProd_PickList.tblmaster_storage_areaDataTable
    Dim TblDummyList_Dum As DseLaporanProd_PickList.Tbl_DummyListDataTable
    Dim TblDetail_Dum As DseLaporanProd_PickList.tbldetailDataTable
    Dim TblProductQty_Dum As DseLaporanProd_PickList.tblmaster_productqtyDataTable
    Dim TblTransDet As New DseLaporanProd_PickList.tbltransferdetailDataTable

    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjPL As New ClsLaporanProd_PickList
    Public PrimaryProduct As Double = -1
    Public PrimaryLocTo As Double = -1
    Public IdentifierTrans As String = ""
    Public ApakahRekap As Boolean = False
    Dim DecimalProd As Integer = 0
    Dim Buffer As String = ""
    Dim LangsungTutup As Boolean = False
    Dim FromOther As Boolean = False

    Private Function ApakahDataBenar() As Boolean
        Dim rowDummyList As DseLaporanProd_PickList.Tbl_DummyListRow
        Dim QtyTransferBef As Decimal = 0
        Dim QtyTransferAft As Decimal = 0

        Try
            'Me.TblDummyList.DefaultView.RowFilter = "[Identifer] = '" & Me.IdentifierTrans & "'"
            'If Me.TblDummyList.DefaultView.Count > 0 Then
            '    rowDummyList = TblDummyList.DefaultView(0).Row
            '    QtyTransferBef = rowDummyList.Qty_Need - rowDummyList.Qty_On_Site
            'End If

            'For Each rowD As DataGridViewRow In Me.DgvTransferDetail.Rows
            '    QtyTransferAft += rowD.Cells(Me.ColTranDet_QtyTrans.Index).Value
            'Next

            'If QtyTransferBef > QtyTransferAft Then
            '    MsgBox("Can't save this data because will cause negatif qty", MsgBoxStyle.Information, "Information")
            '    Return False
            'End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub FrmLaporanProd_PickList_TransDetail_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Hasil As MsgBoxResult

        Try
            If Me.ApakahRekap = False Then
                If Me.FromOther = False Then
                    Hasil = MsgBox("Accept data changes before closing form?", MsgBoxStyle.YesNoCancel, "Confirm")
                    If Hasil = MsgBoxResult.Yes Then
                        If Me.ApakahDataBenar = True Then
                            Me.SimpanTransferDetail()
                            Me.DialogResult = Windows.Forms.DialogResult.OK
                        End If
                    ElseIf Hasil = MsgBoxResult.No Then
                        Me.DialogResult = Windows.Forms.DialogResult.No
                    ElseIf Hasil = MsgBoxResult.Cancel Then
                        Me.DialogResult = Windows.Forms.DialogResult.Cancel
                        e.Cancel = True
                    End If
                Else
                    If Me.LangsungTutup = True Then
                        Me.LangsungTutup = False
                        e.Cancel = True
                    End If
                    Me.FromOther = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanProd_PickList_TransDetail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            If Me.ApakahRekap = True Then
                Me.BtnOK.Visible = False
                Me.BtnCancel.Text = "Close"
            End If
            Me.DgvTransferDetail.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.TblDummyList_Dum = Me.TblDummyList.Copy
            Me.TblDetail_Dum = Me.TblDetail.Copy
            Me.TblProductQty_Dum = Me.TblProductQty.Copy

            Me.ObjPL.IsiDgvTransferDetail(Me.PrimaryProduct, Me.IdentifierTrans, Me.PrimaryLocTo, Me.DgvTransferDetail, _
            Me, Me.TblDummyList_Dum, Me.TblDetail_Dum, Me.TblProductQty_Dum, Me.TblTransDet, Me.ApakahRekap, Me.DecimalProd)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        Try
            Me.FromOther = False
            Me.DgvTransferDetail.EndEdit()
            If Me.ApakahDataBenar() Then
                If MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.DialogResult = Windows.Forms.DialogResult.OK
                    'simpan perubahan yang terjadi
                    Me.SimpanTransferDetail()
                    Me.FromOther = True
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        Try
            If Me.ApakahRekap = False Then
                Me.FromOther = True
                If MsgBox("Close dan discard data changes?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.DialogResult = Windows.Forms.DialogResult.Cancel
                    Me.Close()
                Else
                    Me.LangsungTutup = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvTransferDetail_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvTransferDetail.CellBeginEdit
        Try
            If Me.ApakahRekap = False Then
                If e.ColumnIndex = Me.ColTranDet_QtyTrans.Index Then
                    Me.Buffer = Me.DgvTransferDetail.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvTransferDetail_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvTransferDetail.CellEndEdit

        Try
            If IsNothing(Me.DgvTransferDetail.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColTranDet_QtyTrans.Index Then
                    'pastikan nilai qty > 0
                    If Me.ObjCD.CekNonNegatif(Me.DgvTransferDetail.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please specify positive number.", MsgBoxStyle.Information, "Information")
                        Me.DgvTransferDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    If CDec(Me.DgvTransferDetail.Item(e.ColumnIndex, e.RowIndex).Value) > CDec(Me.DgvTransferDetail.Item(Me.ColTranDet_QtyOnHand.Index, e.RowIndex).Value) Then
                        MsgBox("Maximum value you can enter is " & Me.DgvTransferDetail.Item(Me.ColTranDet_QtyOnHand.Index, e.RowIndex).Value, _
                        MsgBoxStyle.Information, "Information")
                        Me.DgvTransferDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    Me.DgvTransferDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.DgvTransferDetail.Item(e.ColumnIndex, e.RowIndex).Value, Me.DecimalProd)
                    Me.UpdateTblTransferDetail(e.RowIndex, Me.DgvTransferDetail.Item(e.ColumnIndex, e.RowIndex).Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateTblTransferDetail(ByVal vIndex As Integer, ByVal vQtyTransfer As Decimal)
        Dim rowTransDet As DseLaporanProd_PickList.tbltransferdetailRow

        Try
            If Me.TblTransDet.Count >= vIndex Then
                rowTransDet = Me.TblTransDet.Rows(vIndex)
                rowTransDet.Qty_Transfer = vQtyTransfer
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddQty_InTblProductQty(ByVal vPrimaryProduct As Double, ByVal vPrimaryLokasi As Double, ByVal vLokasi As String, _
    ByVal vBatch As String, ByVal vExpDate As Date, ByVal vQty As Decimal, ByVal vQtyPackage As Decimal)
        Dim rowProductQty_Dum As DseLaporanProd_PickList.tblmaster_productqtyRow

        Try
            Me.TblProductQty_Dum.DefaultView.RowFilter = "[Primary Storage Area] = " & vPrimaryLokasi _
            & " AND [Lot No] = '" & vBatch & "' AND [Expiry Date] = #" & vExpDate & "#"
            If Me.TblProductQty_Dum.DefaultView.Count > 0 Then
                rowProductQty_Dum = Me.TblProductQty_Dum.DefaultView(0).Row
                rowProductQty_Dum.Qty_Sisa += vQty
            Else
                rowProductQty_Dum = Me.TblProductQty_Dum.Newtblmaster_productqtyRow
                With rowProductQty_Dum
                    .Primary = -1
                    .Primary_Product = vPrimaryProduct
                    .Qty = vQty
                    .Lot_No = vBatch
                    .Expiry_Date = vExpDate
                    .Primary_Storage_Area = vPrimaryLokasi
                    .Qty_Sisa = vQty
                    .Storage_Area_Code = vLokasi
                    .Is_Qty_On_Site = 1
                    .Urutan_Product_On_Site = 0
                    .Package_Qty = vQtyPackage
                    .Is_Exclude_Qty = 0
                    Me.TblExclude.DefaultView.RowFilter = "[Primary Location] = " & vPrimaryLokasi
                    If Me.TblExclude.DefaultView.Count > 0 Then
                        .Is_Exclude_Qty = 1
                    End If
                End With
                Me.TblProductQty_Dum.Rows.Add(rowProductQty_Dum)

                Me.ObjPL.PrioritasStrictLocation(Me.TblProductQty_Dum, Me.TblPriority, Me.TblExclude, Me.UserCode, False)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanTransferDetail()
        Dim TblDetail_Dum_2 As New DseLaporanProd_PickList.tbldetailDataTable
        Dim TblProductQty_Dum_2 As New DseLaporanProd_PickList.tblmaster_productqtyDataTable
        Dim rowDummyList, rowDummyList_Dum, rowDummyList_Dum_Bef, rowDummyList_Dum_2 As DseLaporanProd_PickList.Tbl_DummyListRow
        Dim rowDetail, rowDetail_Dum, rowDetail_Dum_2 As DseLaporanProd_PickList.tbldetailRow
        Dim rowProductQty, rowProductQty_Dum, rowProductQty_Dum_2 As DseLaporanProd_PickList.tblmaster_productqtyRow
        Dim rowTransDet As DseLaporanProd_PickList.tbltransferdetailRow
        Dim ProsesCalculate As Boolean = False
        Dim QtyTransfer, QtyTransferDetail, QtyMinus_Temp, QtyMinus_Temp_2 As Decimal
        Dim PrimaryLocFrom As Double
        Dim ListPrimaryLoc As New List(Of Double)
        Dim IsLoop As Boolean = False

        Try
            TblDetail_Dum_2 = Me.TblDetail_Dum.Copy
            For i As Integer = Me.TblDetail_Dum.Rows.Count - 1 To 0 Step -1
                rowDetail_Dum = Me.TblDetail_Dum.Rows(i)
                If rowDetail_Dum.Qty = 0 Then
                    Me.TblDetail_Dum.Rows.Remove(rowDetail_Dum)
                End If
            Next

            IsLoop = False
            For i As Integer = 0 To Me.TblDummyList_Dum.Count - 1
                rowDummyList_Dum = Me.TblDummyList_Dum.Rows(i)
                If rowDummyList_Dum.Identifer = Me.IdentifierTrans Then
                    ProsesCalculate = True
                End If

                If ProsesCalculate = True And rowDummyList_Dum.Qty_Sisa > 0 Then
                    QtyTransfer = 0
                    QtyMinus_Temp = 0
                    QtyMinus_Temp_2 = 0
                    PrimaryLocFrom = -1

                    Me.TblTransDet.DefaultView.RowFilter = "[Qty Transfer] > 0"
                    If Me.TblTransDet.DefaultView.Count > 0 Then
                        For Each dv As Data.DataRowView In Me.TblTransDet.DefaultView
                            rowTransDet = dv.Row

                            QtyTransfer = rowTransDet.Qty_Transfer
                            PrimaryLocFrom = rowTransDet.Primary_Location_From

                            '1. update tbldummy
                            rowDummyList_Dum.Qty_Transfer += QtyTransfer
                            rowDummyList_Dum.Qty_Sisa -= QtyTransfer
                            rowDummyList_Dum.Qty_Minus = rowDummyList_Dum.Qty_Transfer - (rowDummyList_Dum.Qty_Need - rowDummyList_Dum.Qty_On_Site)

                            rowTransDet.Qty_Transfer = 0

                            '2. update productqty
                            Me.TblProductQty_Dum.DefaultView.RowFilter = "[Primary Storage Area] = " & PrimaryLocFrom _
                            & " AND [Lot No] = '" & rowTransDet.Batch & "' AND [Expiry Date] = #" & rowTransDet.Exp_Date & "#"
                            If Me.TblProductQty_Dum.DefaultView.Count > 0 Then
                                For Each dv1 As Data.DataRowView In Me.TblProductQty_Dum.DefaultView
                                    rowProductQty_Dum = dv1.Row
                                    QtyTransferDetail = 0
                                    If rowProductQty_Dum.Qty_Sisa > QtyTransfer Then
                                        QtyTransferDetail = QtyTransfer
                                        rowProductQty_Dum.Qty_Sisa -= QtyTransferDetail
                                        QtyTransfer = 0
                                    Else
                                        QtyTransferDetail = rowProductQty_Dum.Qty_Sisa
                                        QtyTransfer -= QtyTransferDetail
                                        rowProductQty_Dum.Qty_Sisa = 0
                                    End If

                                    '3. update tbldetail
                                    Me.ObjPL.TblDetail_AddRow(rowDummyList_Dum, Me.TblDetail_Dum, rowProductQty_Dum, _
                                    QtyTransferDetail, ClsLaporanProd_PickList.EnumGetQty_Transfer.Proses_1, Nothing)
                                    If QtyTransfer = 0 Then
                                        Exit For
                                    End If
                                Next
                            End If

                            If rowDummyList_Dum.Qty_Minus > 0 Then
                                QtyMinus_Temp_2 = rowDummyList_Dum.Qty_Minus - QtyMinus_Temp
                                QtyMinus_Temp = rowDummyList_Dum.Qty_Minus
                                Me.AddQty_InTblProductQty(rowDummyList_Dum.Primary_RM, rowDummyList_Dum.Primary_Location, rowDummyList_Dum.Location, _
                                rowTransDet.Batch, rowTransDet.Exp_Date, QtyMinus_Temp_2, rowDummyList_Dum.Qty_Package)
                            End If
                        Next
                    Else
                        If i > 0 Then
                            rowDummyList_Dum_Bef = Me.TblDummyList_Dum.Rows(i - 1)
                            If rowDummyList_Dum.Primary_RM = rowDummyList_Dum_Bef.Primary_RM And _
                                rowDummyList_Dum.Primary_Location = rowDummyList_Dum_Bef.Primary_Location Then
                                If rowDummyList_Dum_Bef.Qty_Minus < 0 Then
                                    rowDummyList_Dum.Qty_On_Site = rowDummyList_Dum_Bef.Qty_Minus
                                Else
                                    rowDummyList_Dum.Qty_Sisa += rowDummyList_Dum.Qty_On_Site
                                    rowDummyList_Dum.Qty_On_Site = 0
                                End If
                            End If
                        End If

                        'isi qty on site (hanya mempengaruhi tbldummy dan tblproductqty saja)
                        Me.TblProductQty_Dum.DefaultView.RowFilter = "[Primary Storage Area] = " & rowDummyList_Dum.Primary_Location & " AND [Qty Sisa] > 0"
                        If Me.TblProductQty_Dum.DefaultView.Count > 0 Then
                            For Each dv As Data.DataRowView In Me.TblProductQty_Dum.DefaultView
                                rowProductQty_Dum = dv.Row
                                If rowProductQty_Dum.Qty_Sisa > 0 Then
                                    If rowProductQty_Dum.Qty_Sisa <= rowDummyList_Dum.Qty_Sisa Then
                                        rowDummyList_Dum.Qty_On_Site += rowProductQty_Dum.Qty_Sisa
                                        rowDummyList_Dum.Qty_Sisa -= rowProductQty_Dum.Qty_Sisa
                                        rowProductQty_Dum.Qty_Sisa = 0
                                    Else
                                        rowDummyList_Dum.Qty_On_Site += rowDummyList_Dum.Qty_Sisa
                                        rowProductQty_Dum.Qty_Sisa -= rowDummyList_Dum.Qty_Sisa
                                        rowDummyList_Dum.Qty_Sisa = 0
                                    End If
                                End If
                                rowDummyList_Dum.Qty_Minus = rowDummyList_Dum.Qty_Transfer - (rowDummyList_Dum.Qty_Need - rowDummyList_Dum.Qty_On_Site)

                                If rowDummyList_Dum.Qty_Sisa = 0 Then
                                    GoTo SkipTransfer
                                End If
                            Next
                        Else
                            If rowDummyList_Dum.Qty_On_Site >= 0 Then
                                rowDummyList_Dum.Qty_Sisa += rowDummyList_Dum.Qty_On_Site
                                rowDummyList_Dum.Qty_On_Site = 0
                            Else
                                rowDummyList_Dum.Qty_Sisa += rowDummyList_Dum.Qty_On_Site * -1
                            End If
                            rowDummyList_Dum.Qty_Minus = rowDummyList_Dum.Qty_Transfer - (rowDummyList_Dum.Qty_Need - rowDummyList_Dum.Qty_On_Site)
                        End If

                        'cari sesuai urutan prioritas
                        Me.ObjPL.GetQty_InProductQty(TblDummyList_Dum, TblDetail_Dum, Me.TblProductQty_Dum, rowDummyList_Dum, _
                        ClsLaporanProd_PickList.EnumGetQty_Transfer.Proses_1, Nothing)

                        'ambil qty on site dari baris paling bawah tbldummy
                        If rowDummyList_Dum.Qty_Sisa > 0 Then
                            ListPrimaryLoc.Clear()
                            For i3 As Integer = Me.TblDummyList_Dum.Rows.Count - 1 To 0 Step -1
                                rowDummyList_Dum_2 = Me.TblDummyList.Rows(i3)
                                'tidak ambil dari on site yang lokasinya termasuk ke dalam exclude lokasi
                                If rowDummyList_Dum_2.Is_Exclude_Qty = 0 Then
                                    If rowDummyList_Dum_2.Primary_Location <> rowDummyList_Dum.Primary_Location Then
                                        If ListPrimaryLoc.Contains(rowDummyList_Dum_2.Primary_Location) = False Then
                                            ListPrimaryLoc.Add(rowDummyList_Dum_2.Primary_Location)

                                            TblProductQty_Dum_2.Clear()
                                            Me.TblProductQty_Dum.DefaultView.RowFilter = "[Primary Storage Area] = " _
                                            & rowDummyList_Dum_2.Primary_Location & " AND [Qty Sisa] > 0"
                                            If Me.TblProductQty_Dum.DefaultView.Count > 0 Then
                                                TblProductQty_Dum_2.Merge(Me.TblProductQty_Dum.DefaultView.ToTable)
                                            End If

                                            QtyTransferDetail = 0
                                            For Each rowProductQty_Dum_2 In TblProductQty_Dum_2
                                                Me.TblProductQty_Dum.DefaultView.RowFilter = "[Primary Storage Area] = " & rowProductQty_Dum_2.Primary_Storage_Area _
                                                & " AND [Lot No] = '" & rowProductQty_Dum_2.Lot_No & "' AND [Expiry Date] = #" & rowProductQty_Dum_2.Expiry_Date & "#"
                                                If Me.TblProductQty_Dum.DefaultView.Count > 0 Then
                                                    rowProductQty_Dum = Me.TblProductQty_Dum.DefaultView(0).Row
                                                    If rowDummyList_Dum.Qty_Sisa >= rowProductQty_Dum.Qty_Sisa Then
                                                        QtyTransferDetail = rowProductQty_Dum.Qty_Sisa
                                                        rowDummyList_Dum.Qty_Sisa -= QtyTransferDetail
                                                        rowProductQty_Dum.Qty_Sisa = 0
                                                    Else
                                                        QtyTransferDetail = rowDummyList_Dum.Qty_Sisa
                                                        rowProductQty_Dum.Qty_Sisa -= QtyTransferDetail
                                                        rowDummyList_Dum.Qty_Sisa = 0
                                                    End If
                                                    rowDummyList_Dum.Qty_Transfer += QtyTransferDetail
                                                    rowDummyList_Dum.Qty_Minus = rowDummyList_Dum.Qty_Transfer - (rowDummyList_Dum.Qty_Need - rowDummyList_Dum.Qty_On_Site)

                                                    Me.ObjPL.TblDetail_AddRow(rowDummyList_Dum, Me.TblDetail_Dum, rowProductQty_Dum_2, _
                                                    QtyTransferDetail, ClsLaporanProd_PickList.EnumGetQty_Transfer.Proses_1, Nothing)

                                                    Me.AddQty_InTblProductQty(Me.PrimaryProduct, rowDummyList_Dum.Primary_Location, rowDummyList_Dum.Location, rowProductQty_Dum_2.Lot_No, _
                                                    rowProductQty_Dum_2.Expiry_Date, QtyTransferDetail, rowDummyList_Dum.Qty_Package)
                                                End If
                                            Next
                                        End If
                                    End If
                                    If rowDummyList_Dum.Qty_Sisa = 0 Then
                                        Exit For
                                    End If
                                End If
                            Next
                        End If
SkipTransfer:
                    End If
                End If
            Next

            Me.TblDummyList.DefaultView.RowFilter = "[Primary RM] = " & Me.PrimaryProduct
            For Each dv As Data.DataRowView In Me.TblDummyList.DefaultView
                rowDummyList = dv.Row
                Me.TblDummyList_Dum.DefaultView.RowFilter = "[Identifier JO]='" & rowDummyList.Identifier_JO.ToString _
                & "' AND [Identifier WO]='" & rowDummyList.Identifier_WO.ToString & "' AND [Primary Work Unit] = " _
                & rowDummyList.Primary_Work_Unit
                If Me.TblDummyList_Dum.DefaultView.Count > 0 Then
                    rowDummyList_Dum = Me.TblDummyList_Dum.DefaultView(0).Row
                    rowDummyList.Qty_On_Site = rowDummyList_Dum.Qty_On_Site
                    rowDummyList.Qty_Sisa = rowDummyList_Dum.Qty_Sisa
                    rowDummyList.Qty_Transfer = rowDummyList_Dum.Qty_Transfer
                    rowDummyList.Qty_Minus = rowDummyList_Dum.Qty_Minus
                End If
            Next

            Me.TblProductQty.DefaultView.RowFilter = "[Primary Product] = " & Me.PrimaryProduct
            While Me.TblProductQty.DefaultView.Count > 0
                rowProductQty = Me.TblProductQty.DefaultView(0).Row
                Me.TblProductQty.Rows.Remove(rowProductQty)
            End While
            Me.TblProductQty.Merge(Me.TblProductQty_Dum.Copy)

            Me.TblDetail_Dum.DefaultView.RowFilter = "[Qty] <= 0"
            While Me.TblDetail_Dum.DefaultView.Count > 0
                rowDetail_Dum = Me.TblDetail_Dum.DefaultView(0).Row
                Me.TblDetail_Dum.Rows.Remove(rowDetail_Dum)
            End While

            Me.TblDetail.DefaultView.RowFilter = "[Primary Product] = " & Me.PrimaryProduct
            While Me.TblDetail.DefaultView.Count > 0
                rowDetail = Me.TblDetail.DefaultView(0).Row
                Me.TblDetail.Rows.Remove(rowDetail)
            End While
            Me.TblDetail.Merge(Me.TblDetail_Dum.Copy)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class