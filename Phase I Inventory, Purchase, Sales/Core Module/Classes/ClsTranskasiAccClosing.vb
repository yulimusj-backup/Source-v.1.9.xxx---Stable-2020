Public Class ClsTranskasiAccClosing
    Public Function GetAccountingClosingDate() As Date
        Dim DTA As New DseTransaksiAccClosingTableAdapters.tbltrans_accounting_closingTableAdapter

        Try
            If DTA.ScalarQuery_GetClosingDate.HasValue = True Then
                Return DTA.ScalarQuery_GetClosingDate.Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSMESSConvertionDate() As Date
        Dim DTA As New DseTransaksiAccClosingTableAdapters.tbltrans_accounting_closingTableAdapter

        Try
            If DTA.ScalarQuery_GetConvertionDate.HasValue = True Then
                Return DTA.ScalarQuery_GetConvertionDate.Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPayrollConvertionDate() As Date
        Dim DTA As New DseTransaksiAccClosingTableAdapters.tbltrans_accounting_closingTableAdapter
        Dim Dummy As System.Nullable(Of Date)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPayrollConventionDate()
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
