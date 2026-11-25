Option Explicit On
Option Strict Off
Imports Core_Module

Public Class ClsMast_Diskon_Item
    Public Function GetConstring() As String
        Try
            Return "server=" + My.Settings.server + ";user id=" + My.Settings.user_id + ";Password=" + My.Settings.password + _
            ";persist security info=True;database=" + My.Settings.database_name + ";port=" + My.Settings.port_number + ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "server=localhost;user id=root;Password=123;persist security info=True;database=core_db;port=3307"
        End Try
    End Function

    Public Function IsiDGV(ByVal vPrimary As Double, ByVal vDGVItemProd As DataGridView, _
    ByVal vDGVItemBranch As DataGridView, ByVal i As FrmMast_Discount, ByVal vObjCD As ClsCekData) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItemP As New DseMast_DiskonTableAdapters.DataTable1TableAdapter
        Dim TBLItemP As New DseMast_Diskon.DataTable1DataTable
        Dim rowItemP As DseMast_Diskon.DataTable1Row
        Dim DTAItemB As New DseMast_DiskonTableAdapters.DataTable2TableAdapter
        Dim TBLItemB As New DseMast_Diskon.DataTable2DataTable
        Dim rowItemB As DseMast_Diskon.DataTable2Row
        Try
            'update connection string
            DTAItemP.Connection.ConnectionString = Me.GetConstring
            DTAItemB.Connection.ConnectionString = Me.GetConstring

            'autosize dgv
            vDGVItemProd.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItemProd.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVItemProd.AllowUserToAddRows = False
            vDGVItemBranch.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItemBranch.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVItemBranch.AllowUserToAddRows = False

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan data item
            DTAItemP.Fill(TBLItemP, vPrimary)
            vDGVItemProd.Rows.Clear()
            For Each rowItemP In TBLItemP.Rows
                vDGVItemProd.Rows.Add()
                rowDGV = vDGVItemProd.Rows(vDGVItemProd.Rows.Count - 1)
                With rowDGV
                    .Cells(i.ColCodeNo.Index).Value = rowItemP.Product_Code
                    .Cells(i.ColDesc.Index).Value = rowItemP.Description
                    .Cells(i.ColRemarks.Index).Value = rowItemP.Remarks
                    .Cells(i.ColPrimaryProduct.Index).Value = rowItemP.Primary_Product
                End With
            Next

            DTAItemB.Fill(TBLItemB, vPrimary)
            vDGVItemBranch.Rows.Clear()
            For Each rowItemB In TBLItemB.Rows
                vDGVItemBranch.Rows.Add()
                rowDGV = vDGVItemBranch.Rows(vDGVItemBranch.Rows.Count - 1)
                With rowDGV
                    .Cells(i.ColBranch_Code.Index).Value = rowItemB.Code
                    .Cells(i.ColBranch_Desc.Index).Value = rowItemB.Description
                    .Cells(i.ColRemarks.Index).Value = rowItemB.Remarks
                    .Cells(i.ColBranch_Primary.Index).Value = rowItemB.Primary_Branch
                End With
            Next

            vObjCD.SetNumber(vDGVItemProd)
            vObjCD.SetNumber(vDGVItemBranch)

            'autosize dgv
            vDGVItemProd.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItemProd.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            vDGVItemProd.AllowUserToAddRows = True
            vDGVItemBranch.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItemBranch.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            vDGVItemBranch.AllowUserToAddRows = True

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItemP.Dispose()
            TBLItemP.Dispose()
            DTAItemB.Dispose()
            TBLItemB.Dispose()
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimary As Double) As Boolean
        Dim DTAP As New DseMast_DiskonTableAdapters.tblmaster_discount_itemproductTableAdapter
        Dim DTAB As New DseMast_DiskonTableAdapters.tblmaster_discount_itembranchTableAdapter
        Try
            DTAP.Connection.ConnectionString = Me.GetConstring
            DTAB.Connection.ConnectionString = Me.GetConstring
            DTAP.Delete(vPrimary)
            DTAB.Delete(vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAP.Dispose()
            DTAB.Dispose()
        End Try
    End Function

    Public Function SimpanItem(ByVal vPrimary As Double, ByVal vDGVItemProduct As DataGridView, ByVal vDGVItemBranch As DataGridView, _
    ByVal i As FrmMast_Discount) As Boolean
        Dim DTAItemP As New DseMast_DiskonTableAdapters.tblmaster_discount_itemproductTableAdapter
        Dim DTAItemB As New DseMast_DiskonTableAdapters.tblmaster_discount_itembranchTableAdapter
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New ClsCekData
        Try
            DTAItemP.Connection.ConnectionString = Me.GetConstring
            DTAItemB.Connection.ConnectionString = Me.GetConstring

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            For Each rowDGV In vDGVItemProduct.Rows
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColRemarks.Index).Value) = True Then
                    rowDGV.Cells(i.ColRemarks.Index).Value = ""
                End If
                With rowDGV
                    DTAItemP.Insert(vPrimary, .Cells(i.ColRemarks.Index).Value, _
                    .Cells(i.ColPrimaryProduct.Index).Value)
                End With
            Next

            For Each rowDGV In vDGVItemBranch.Rows
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColBranc_Remarks.Index).Value) = True Then
                    rowDGV.Cells(i.ColBranc_Remarks.Index).Value = ""
                End If
                With rowDGV
                    DTAItemB.Insert(vPrimary, .Cells(i.ColBranc_Remarks.Index).Value, _
                    .Cells(i.ColBranch_Primary.Index).Value)
                End With
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItemP.Dispose()
            DTAItemB.Dispose()
            ObjCD = Nothing
        End Try
    End Function

End Class
