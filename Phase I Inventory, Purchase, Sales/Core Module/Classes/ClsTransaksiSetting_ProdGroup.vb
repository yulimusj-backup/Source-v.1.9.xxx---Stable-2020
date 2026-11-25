Option Explicit On
Option Strict Off

Public Class ClsTransaksiSetting_ProdGroup
    Public Enum EnumTipeTransaksi As Integer
        LaporanPurchaseRequest
    End Enum

    Public Enum EnumGroupLevel As Integer
        ProductGroup
        Brand
        SubBrand
        Category
        Product
    End Enum

    Public Function GetLevel_Str(ByVal vIdx As EnumGroupLevel) As String
        Try
            If vIdx = EnumGroupLevel.ProductGroup Then
                Return "Product Group"
            ElseIf vIdx = EnumGroupLevel.Brand Then
                Return "Brand"
            ElseIf vIdx = EnumGroupLevel.SubBrand Then
                Return "Sub Brand"
            ElseIf vIdx = EnumGroupLevel.Category Then
                Return "Category"
            ElseIf vIdx = EnumGroupLevel.Product Then
                Return "Product"
            End If

            Return "Product"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "Product"
        End Try
    End Function

    Public Function GetLevelInt(ByVal vLevelStr As String) As EnumGroupLevel
        Try
            If vLevelStr = "Product Group" Then
                Return EnumGroupLevel.ProductGroup
            ElseIf vLevelStr = "Brand" Then
                Return EnumGroupLevel.Brand
            ElseIf vLevelStr = "Sub Brand" Then
                Return EnumGroupLevel.SubBrand
            ElseIf vLevelStr = "Category" Then
                Return EnumGroupLevel.Category
            ElseIf vLevelStr = "Product" Then
                Return EnumGroupLevel.Product
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 4
        End Try
    End Function

    Public Sub SimpanData(ByVal vUserCode As String, ByVal vTipeTransaksi As EnumTipeTransaksi, ByVal vDGV As DataGridView, _
    ByVal vIdx_No As Integer, ByVal vIdxLevel As Integer, ByVal vGroupCode As String, ByVal vGroupDesc As String, _
    ByVal vIdxRemarks As Integer, ByVal vIdxPrimaryRef As Integer)
        Dim DTA As New DseTransaksiSetting_ProductGroupTableAdapters.tbltrans_setting_product_groupTableAdapter
        Dim Dummy As Boolean
        Dim rowDGV As DataGridViewRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = vDGV.AllowUserToAddRows
            vDGV.AllowUserToAddRows = False
            DTA.DeleteQuery(vUserCode, CInt(vTipeTransaksi))
            For Each rowDGV In vDGV.Rows
                If rowDGV.Cells(vIdxPrimaryRef).Value <> -1 Then
                    DTA.Insert(Me.GetLevelInt(rowDGV.Cells(vIdxLevel).Value), _
                    rowDGV.Cells(vIdxPrimaryRef).Value, rowDGV.Cells(vIdxRemarks).Value, vUserCode, CInt(vTipeTransaksi))
                End If
            Next
            vDGV.AllowUserToAddRows = Dummy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanDGV(ByVal vUserCode As String, ByVal vTipeTransaksi As EnumTipeTransaksi, ByVal vDGV As DataGridView, _
    ByVal vIdx_No As Integer, ByVal vIdxLevel As Integer, ByVal vIdxGroupCode As Integer, ByVal vIdxGroupDesc As Integer, _
    ByVal vIdxRemarks As Integer, ByVal vIdxPrimaryRef As Integer)
        Dim DTA As New DseTransaksiSetting_ProductGroupTableAdapters.tbltrans_setting_product_groupTableAdapter
        Dim TBL As New DseTransaksiSetting_ProductGroup.tbltrans_setting_product_groupDataTable
        Dim row As DseTransaksiSetting_ProductGroup.tbltrans_setting_product_groupRow
        Dim Dummy As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTACategory As New DseTransaksiSetting_ProductGroupTableAdapters.tblmaster_categoryTableAdapter
        Dim TBLCategory As New DseTransaksiSetting_ProductGroup.tblmaster_categoryDataTable
        Dim rowCategory As DseTransaksiSetting_ProductGroup.tblmaster_categoryRow
        Dim DTAProdGroup As New DseTransaksiSetting_ProductGroupTableAdapters.tblmaster_product_groupTableAdapter
        Dim TBLProdGroup As New DseTransaksiSetting_ProductGroup.tblmaster_product_groupDataTable
        Dim rowProdGroup As DseTransaksiSetting_ProductGroup.tblmaster_product_groupRow
        Dim GroupLevel As EnumGroupLevel = EnumGroupLevel.ProductGroup
        Dim DTAProd As New DseTransaksiSetting_ProductGroupTableAdapters.tblmaster_productTableAdapter
        Dim TBLProd As New DseTransaksiSetting_ProductGroup.tblmaster_productDataTable
        Dim rowProd As DseTransaksiSetting_ProductGroup.tblmaster_productRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTACategory.Connection.ConnectionString = Core_Module.GetConstring
            DTAProdGroup.Connection.ConnectionString = Core_Module.GetConstring
            DTAProd.Connection.ConnectionString = Core_Module.GetConstring

            Dummy = vDGV.AllowUserToAddRows
            vDGV.AllowUserToAddRows = False
            vDGV.Rows.Clear()
            DTA.DeleteQuery1() 'untuk mencegah logic error karena ada baris baris yg primary -1, tapi disimpan di db
            DTA.FillBy_usercode_tipetransaksi(TBL, vUserCode, CInt(vTipeTransaksi))
            For Each row In TBL.Rows
                vDGV.Rows.Add()
                rowDGV = vDGV.Rows(vDGV.Rows.Count - 1)
                rowDGV.Cells(vIdx_No).Value = vDGV.Rows.Count
                rowDGV.Cells(vIdxLevel).Value = Me.GetLevel_Str(row.level)
                rowDGV.Cells(vIdxPrimaryRef).Value = row.primary_acuan
                rowDGV.Cells(vIdxRemarks).Value = row.remarks
                GroupLevel = row.level

                If GroupLevel = EnumGroupLevel.Product Then
                    If TBLProd.Rows.Count = 0 Then
                        DTAProd.Fill(TBLProd)
                    End If
                    TBLProd.DefaultView.RowFilter = "Primary = " + row.primary_acuan.ToString
                    rowProd = TBLProd.DefaultView(0).Row
                    rowDGV.Cells(vIdxGroupCode).Value = rowProd.Product_Code
                    rowDGV.Cells(vIdxGroupDesc).Value = rowProd.Description
                ElseIf GroupLevel = EnumGroupLevel.Category Then
                    If TBLCategory.Rows.Count = 0 Then
                        DTACategory.Fill(TBLCategory)
                    End If
                    TBLCategory.DefaultView.RowFilter = "Primary = " + row.primary_acuan.ToString
                    rowCategory = TBLCategory.DefaultView(0).Row
                    rowDGV.Cells(vIdxGroupCode).Value = rowCategory.Category_Code
                    rowDGV.Cells(vIdxGroupDesc).Value = rowCategory.Description
                Else
                    If TBLProdGroup.Rows.Count = 0 Then
                        DTAProdGroup.Fill(TBLProdGroup)
                    End If
                    TBLProdGroup.DefaultView.RowFilter = "Primary = " + row.primary_acuan.ToString
                    rowProdGroup = TBLProdGroup.DefaultView(0).Row
                    rowDGV.Cells(vIdxGroupCode).Value = rowProdGroup.Group_Code
                    rowDGV.Cells(vIdxGroupDesc).Value = rowProdGroup.Description
                End If
            Next
            vDGV.AllowUserToAddRows = Dummy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTACategory.Dispose()
            TBLCategory.Dispose()
            DTAProdGroup.Dispose()
            TBLProdGroup.Dispose()
            DTAProd.Dispose()
            TBLProd.Dispose()
        End Try
    End Sub

    Public Sub Hapus(ByVal vUserCode As String, ByVal vTipeTransaksi As EnumTipeTransaksi)
        Dim DTA As New DseTransaksiSetting_ProductGroupTableAdapters.tbltrans_setting_product_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.DeleteQuery(vUserCode, CInt(vTipeTransaksi))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub
End Class
