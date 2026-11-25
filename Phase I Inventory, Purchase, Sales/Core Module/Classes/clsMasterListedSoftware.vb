Option Explicit On
Option Strict Off

''' <summary>
''' kelas ini mendata software - software Artha yang diinstall ke dalam sistem
''' </summary>
''' <remarks></remarks>
Public Class clsMasterListedSoftware
    Public Sub SetInventoryLocked(ByVal vLocked As Boolean)
        Dim DTA As New dseMasterListedSoftwareTableAdapters.tbllistedsoftwareTableAdapter
        Dim TBL As New dseMasterListedSoftware.tbllistedsoftwareDataTable
        Dim row As dseMasterListedSoftware.tbllistedsoftwareRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
            row = TBL.Rows(0)
            row.Inventory_Locked = vLocked
            DTA.Update(TBL)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Function IsInventoryLocked() As Boolean
        Dim DTA As New dseMasterListedSoftwareTableAdapters.tbllistedsoftwareTableAdapter
        Dim TBL As New dseMasterListedSoftware.tbllistedsoftwareDataTable
        Dim row As dseMasterListedSoftware.tbllistedsoftwareRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
            row = TBL.Rows(0)
            If row.Inventory_Locked = True Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function IsInventoryInstalled() As Boolean
        Try
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function IsSalesInstalled() As Boolean
        Dim DTA As New dseMasterListedSoftwareTableAdapters.tbllistedsoftwareTableAdapter
        Dim TBL As New dseMasterListedSoftware.tbllistedsoftwareDataTable
        Dim row As dseMasterListedSoftware.tbllistedsoftwareRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
            row = TBL.Rows(0)
            Return CBool(row.Sales)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function IsPurchaseInstalled() As Boolean
        Dim DTA As New dseMasterListedSoftwareTableAdapters.tbllistedsoftwareTableAdapter
        Dim TBL As New dseMasterListedSoftware.tbllistedsoftwareDataTable
        Dim row As dseMasterListedSoftware.tbllistedsoftwareRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
            row = TBL.Rows(0)
            Return CBool(row.Purchase)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function IsAccountingInstalled() As Boolean
        Dim DTA As New dseMasterListedSoftwareTableAdapters.tbllistedsoftwareTableAdapter
        Dim TBL As New dseMasterListedSoftware.tbllistedsoftwareDataTable
        Dim row As dseMasterListedSoftware.tbllistedsoftwareRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
            row = TBL.Rows(0)
            Return CBool(row.Accounting)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function IsHRDInstalled() As Boolean
        Dim DTA As New dseMasterListedSoftwareTableAdapters.tbllistedsoftwareTableAdapter
        Dim TBL As New dseMasterListedSoftware.tbllistedsoftwareDataTable
        Dim row As dseMasterListedSoftware.tbllistedsoftwareRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
            row = TBL.Rows(0)
            Return CBool(row.HRD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function
End Class
