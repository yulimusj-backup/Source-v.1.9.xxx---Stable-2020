Option Explicit On
Option Strict Off

Public Class ClsMasterCekBatasan
    Dim ObjIdentifier As New ClsIdentifier

    Public Function GetJmlWeight(ByVal vPrimaryProduct As Double, ByVal vPrimaryUnit As Double, ByVal vPrimaryUnitWeight As Double) As Integer
        Dim dta As New DseMasterCekBatasanTableAdapters.tbltrans_custom_cetak_packinglistitem_optionTableAdapter
        Dim Dummy As Nullable(Of Long)
        Dim Retval As Integer = 0

        Try
            If vPrimaryProduct = -1 Or vPrimaryUnit = -1 Or vPrimaryUnitWeight = -1 Then
                Return Retval
            End If

            If ObjIdentifier.GetIdentifierDB = "smess_lmb" Then 'untuk sementara fungsi cek batasan master hanya untuk smess_lmb 
                dta.Connection.ConnectionString = Core_Module.GetConstring
                Dummy = dta.ScalarQuery_GetJml(vPrimaryProduct, vPrimaryUnit, vPrimaryUnitWeight)
                If Dummy.HasValue = True Then
                    Retval = Dummy.Value
                End If
            End If

            Return Retval
        Catch ex As Exception
            If ex.Message.ToString.StartsWith("Table") <> True And ex.Message.ToString.Contains("doesn't exist") <> True Then
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
            Return Retval
        End Try
    End Function

    Public Function GetMaxDecimalWeight(ByVal vPrimaryProduct As Double, ByVal vPrimaryUnit As Double, ByVal vPrimaryUnitWeight As Double) As Integer
        Dim dta As New DseMasterCekBatasanTableAdapters.tbltrans_custom_cetak_packinglistitem_optionTableAdapter
        Dim Dummy As Nullable(Of Integer)
        Dim Retval As Integer = 0

        Try
            If vPrimaryProduct = -1 Or vPrimaryUnit = -1 Or vPrimaryUnitWeight = -1 Then
                Return Retval
            End If

            If ObjIdentifier.GetIdentifierDB = "smess_lmb" Then 'untuk sementara fungsi cek batasan master hanya untuk smess_lmb 
                dta.Connection.ConnectionString = Core_Module.GetConstring
                Dummy = dta.ScalarQuery_GetMaxDefPrecWeight(vPrimaryProduct, vPrimaryUnit, vPrimaryUnitWeight)
                If Dummy.HasValue = True Then
                    Retval = Dummy.Value
                End If
            End If

            Return Retval
        Catch ex As Exception
            If ex.Message.ToString.StartsWith("Table") <> True And ex.Message.ToString.Contains("doesn't exist") <> True Then
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
            Return Retval
        End Try
    End Function
End Class
