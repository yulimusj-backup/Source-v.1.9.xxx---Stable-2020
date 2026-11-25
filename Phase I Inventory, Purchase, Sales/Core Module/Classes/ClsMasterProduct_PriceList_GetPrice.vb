Option Strict Off
Option Explicit On

Public Class ClsMasterProduct_PriceList_GetPrice
    Public Function GetCustSalesArea(ByVal vPrimary As Double) As Double
        Dim DTA As New DseMasterProduct_EditPriceList_GetPriceTableAdapters.tblmaster_customerTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimarySalesArea(vPrimary)
            If Dummy.HasValue Then
                Return Dummy
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetCustCustomerGroup(ByVal vPrimary As Double) As Double
        Dim DTA As New DseMasterProduct_EditPriceList_GetPriceTableAdapters.tblmaster_customerTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimaryCustomerGroup(vPrimary)
            If Dummy.HasValue Then
                Return Dummy
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetCustSegmentation(ByVal vPrimary As Double) As Double
        Dim DTA As New DseMasterProduct_EditPriceList_GetPriceTableAdapters.tblmaster_customerTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimarySegmentation(vPrimary)
            If Dummy.HasValue Then
                Return Dummy
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetCustType(ByVal vPrimary As Double) As Double
        Dim DTA As New DseMasterProduct_EditPriceList_GetPriceTableAdapters.tblmaster_customerTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimaryType(vPrimary)
            If Dummy.HasValue Then
                Return Dummy
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTipeGroup() As Integer
        Dim DTA As New DseMasterProduct_EditPriceList_GetPriceTableAdapters.tblmaster_product_pricelist_groupTableAdapter
        Dim Dummy As System.Nullable(Of Integer)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetTipeGroup
            If Dummy.HasValue Then
                Return Dummy
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
