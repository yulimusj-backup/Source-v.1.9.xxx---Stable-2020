Option Explicit On
Option Strict Off

Public Class ClsIdentifier
    Dim Identifier As String = "smess_lmb"

    Public Function GetIdentifier() As String
        Return Me.Identifier
    End Function

    Public Function GetIdentifierDB() As String
        Dim DTA As New DseTransaksiCustom_IdentifierTableAdapters.tbltrans_custom_identifierTableAdapter
        Dim TBL As New DseTransaksiCustom_Identifier.tbltrans_custom_identifierDataTable
        Dim row As DseTransaksiCustom_Identifier.tbltrans_custom_identifierRow
        Dim RetVal As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Fill(TBL) > 0 Then
                row = TBL.Rows(0)
                RetVal = row.identifier
            End If

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return RetVal
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function ApakahIdentifierDBCocok() As Boolean
        Dim DTA As New DseTransaksiCustom_IdentifierTableAdapters.tbltrans_custom_identifierTableAdapter
        Dim TBL As New DseTransaksiCustom_Identifier.tbltrans_custom_identifierDataTable
        Dim row As DseTransaksiCustom_Identifier.tbltrans_custom_identifierRow
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'cek apakah baris ada
            RetVal = DTA.ScalarQuery_GetJmlBaris
            If RetVal.HasValue = True Then
                If RetVal.Value > 0 Then
                    'jika ada apakah cocok dgn db, jika cocok return true
                    DTA.Fill(TBL)
                    row = TBL.Rows(0)
                    If row.identifier <> Me.Identifier Then
                        Return False
                    Else
                        Return True
                    End If
                Else
                    'jika baris tidak ada maka insert baris dan return true
                    DTA.Insert(Me.Identifier)
                    Return True
                End If
            Else
                'jika baris tidak ada maka insert baris dan return true
                DTA.Insert(Me.Identifier)
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function
End Class
