Option Explicit On
Option Strict Off

Public Class ClsABT
    Public Enum EnumGroupLevel As Integer
        LvlGroup
        LvlBrand
        LvlSubBrand
        LvlCategory
        LvlProduct
    End Enum

    Public Function ApakahBenang(ByVal vPrimaryKategory As Double) As Boolean
        Dim DTA As New dseMasterCategoryTableAdapters.tblmaster_category2TableAdapter
        Dim TBL As New dseMasterCategory.tblmaster_category2DataTable
        Dim row As dseMasterCategory.tblmaster_category2Row 

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.FillBy_PrimaryCategory(TBL, vPrimaryKategory)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                If row.Primary_Product_Group = 3 Or _
                row.Primary_Product_Group = 56 Or _
                row.Primary_Product_Group = 57 Or _
                row.Primary_Product_Group = 58 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Sub ApakahKain(ByVal vTBLProd As DseABT.tblmaster_productDataTable, _
    ByVal vTBLListKain As Core_Module.DseABT.tbltrans_custom_product_kainDataTable)
        Dim row As Core_Module.DseABT.tblmaster_productRow
        Dim Level As EnumGroupLevel = EnumGroupLevel.LvlProduct
        Dim row1 As Core_Module.DseABT.tbltrans_custom_product_kainRow

        Try
            For Each row In vTBLProd.Rows
                row.is_kain = 0
                For Each row1 In vTBLListKain.Rows
                    Level = row1.level
                    If Level = EnumGroupLevel.LvlGroup Then
                        If row1.primary_acuan = row.Primary_Product_Group Then
                            row.is_kain = 1
                            Exit For
                        End If
                    ElseIf Level = EnumGroupLevel.LvlBrand Then
                        If row1.primary_acuan = row.Primary_Brand Then
                            row.is_kain = 1
                            Exit For
                        End If
                    ElseIf Level = EnumGroupLevel.LvlSubBrand Then
                        If row1.primary_acuan = row.Primary_Sub_Brand Then
                            row.is_kain = 1
                            Exit For
                        End If
                    ElseIf Level = EnumGroupLevel.LvlCategory Then
                        If row1.primary_acuan = row.Primary_Category Then
                            row.is_kain = 1
                            Exit For
                        End If
                    ElseIf Level = EnumGroupLevel.LvlProduct Then
                        If row1.primary_acuan = row.Primary_Product Then
                            row.is_kain = 1
                            Exit For
                        End If
                    End If
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function ApakahKain(ByVal vPrimaryKategory As Double) As Boolean
        Dim DTA As New dseMasterCategoryTableAdapters.tblmaster_category2TableAdapter
        Dim TBL As New dseMasterCategory.tblmaster_category2DataTable
        Dim row As dseMasterCategory.tblmaster_category2Row
        Dim DTA1 As New DseABTTableAdapters.tbltrans_custom_product_kainTableAdapter
        Dim TBL1 As New DseABT.tbltrans_custom_product_kainDataTable
        Dim row1 As DseABT.tbltrans_custom_product_kainRow
        Dim Level As EnumGroupLevel = EnumGroupLevel.LvlProduct

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString

            DTA.FillBy_PrimaryCategory(TBL, vPrimaryKategory)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                DTA1.Fill(TBL1)

                For Each row1 In TBL1.Rows
                    Level = row1.level
                    If Level = EnumGroupLevel.LvlGroup Then
                        If row1.primary_acuan = row.Primary_Product_Group Then
                            Return True
                        End If
                    ElseIf Level = EnumGroupLevel.LvlBrand Then
                        If row1.primary_acuan = row.Primary_Brand Then
                            Return True
                        End If
                    ElseIf Level = EnumGroupLevel.LvlSubBrand Then
                        If row1.primary_acuan = row.Primary_Sub_Brand Then
                            Return True
                        End If
                    ElseIf Level = EnumGroupLevel.LvlCategory Then
                        If row1.primary_acuan = row.Primary_Category Then
                            Return True
                        End If
                    End If
                Next

                Return False
            Else
                Return False
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTA1.Dispose()
            TBL1.Dispose()
        End Try
    End Function
End Class
