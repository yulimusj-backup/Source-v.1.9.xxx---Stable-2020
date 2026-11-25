Option Explicit On
Option Strict Off

Public Class ClsTransaksiSalesOrder
    Public Sub GetCustCode_CustName_CustPONO_SalesOrderNumber(ByVal vPrimarySO As Double, ByRef CustName As String, _
       ByRef CustCode As String, ByRef CustPONO As String, ByRef SO_Number As String, ByVal CustContact As String)
        Dim DTA As New DseTransaksiSOTableAdapters.DataTable2TableAdapter
        Dim TBL As New DseTransaksiSO.DataTable2DataTable
        Dim row As DseTransaksiSO.DataTable2Row

        Try
            If DTA.FillBy_PrimarySO(TBL, vPrimarySO) > 0 Then
                row = TBL.Rows(0)
                With row
                    CustName = .Customer_Name
                    CustCode = .Customer_Code
                    CustPONO = .Cust_PO_No
                    SO_Number = .SO_Number
                    CustContact = .Contact_Person
                End With
            Else
                CustName = ""
                CustCode = ""
                CustPONO = ""
                SO_Number = ""
                CustContact = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub
End Class
