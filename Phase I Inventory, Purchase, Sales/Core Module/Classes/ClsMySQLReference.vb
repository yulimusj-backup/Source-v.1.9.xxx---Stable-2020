Option Explicit On
Option Strict Off

Imports Microsoft.VisualBasic
Imports System.Data
Imports MySql.Data.MySqlClient
Imports System.Configuration

Public Class ClsMySQLReference
    Private objConn As MySqlConnection
    Private objCmd As MySqlCommand
    Private Trans As MySqlTransaction
    Private strConnString As String = Core_Module.GetConstring

    Public Sub New()
        strConnString = Core_Module.GetConstring
    End Sub

    Public Function GetServerTime() As Date
        Dim DTA As New DseMasterSQLReferenceTableAdapters.server_time_viewTableAdapter
        Dim TBL As New DseMasterSQLReference.server_time_viewDataTable
        Dim row As DseMasterSQLReference.server_time_viewRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
            If TBL.Rows.Count = 0 Then
                Return Now
            Else
                row = TBL.Rows(0)
                Return row.Curdate + TimeSpan.FromSeconds(row.Curtime.TotalSeconds)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function QueryDataReader(ByVal strSQL As String) As MySqlDataReader
        Dim dtReader As MySqlDataReader

        Try
            objConn = New MySqlConnection
            With objConn
                .ConnectionString = strConnString
                .Open()
            End With
            objCmd = New MySqlCommand(strSQL, objConn)
            dtReader = objCmd.ExecuteReader()
            Return dtReader '*** Return DataReader ***'  
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Nothing
        Finally
            objConn.Close()
            objConn = Nothing
            objCmd = Nothing
        End Try
    End Function

    Public Function QueryDataSet(ByVal strSQL As String) As DataSet
        Dim ds As New DataSet
        Dim dtAdapter As New MySqlDataAdapter

        Try
            objConn = New MySqlConnection
            With objConn
                .ConnectionString = strConnString
                .Open()
            End With
            objCmd = New MySqlCommand
            With objCmd
                .Connection = objConn
                .CommandText = strSQL
                .CommandType = CommandType.Text
            End With
            dtAdapter.SelectCommand = objCmd
            dtAdapter.Fill(ds)
            Return ds   '*** Return DataSet ***'  
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ds
        Finally
            objConn.Close()
            objConn.Dispose()
            dtAdapter.Dispose()
        End Try
    End Function

    Public Function QueryDataTable(ByVal strSQL As String) As DataTable
        Dim dtAdapter As MySqlDataAdapter
        Dim dt As New DataTable

        Try
            objConn = New MySqlConnection
            With objConn
                .ConnectionString = strConnString
                .Open()
            End With
            dtAdapter = New MySqlDataAdapter(strSQL, objConn)
            dtAdapter.Fill(dt)
            Return dt '*** Return DataTable ***'  
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return dt
        Finally
            objConn.Close()
            objConn.Dispose()
        End Try
    End Function

    Public Function QueryExecuteNonQuery(ByVal strSQL As String) As Boolean
        objConn = New MySqlConnection
        With objConn
            .ConnectionString = strConnString
            .Open()
        End With

        Try
            objCmd = New MySqlCommand()
            With objCmd
                .Connection = objConn
                .CommandType = CommandType.Text
                .CommandText = strSQL
            End With
            objCmd.ExecuteNonQuery()
            Return True '*** Return True ***'  
        Catch ex As Exception
            Return False '*** Return False ***'  
        Finally
            objConn.Close()
            objConn.Dispose()
            objCmd.Dispose()
        End Try
    End Function

    Public Function QueryExecuteScalar(ByVal strSQL As String) As Object
        Dim obj As Object
        objConn = New MySqlConnection(strConnString)
        With objConn
            '.ConnectionString = strConnString
            .Open()
        End With

        Try
            objCmd = New MySqlCommand()
            With objCmd
                .Connection = objConn
                .CommandType = CommandType.Text
                .CommandText = strSQL
            End With
            obj = objCmd.ExecuteScalar()  '*** Return Scalar ***'
            If obj = Nothing Then
                obj = Nothing
            End If
            Return obj
        Catch ex As Exception
            Return Nothing '*** Return Nothing ***'  
        Finally
            objConn.Close()
            objConn = Nothing
        End Try
    End Function

    Public Sub TransStart()
        objConn = New MySqlConnection
        With objConn
            .ConnectionString = strConnString
            .Open()
        End With
        Trans = objConn.BeginTransaction(IsolationLevel.ReadCommitted)
    End Sub

    Public Function TransExecute(ByVal strSQL As String) As Boolean
        objCmd = New MySqlCommand()
        With objCmd
            .Connection = objConn
            .Transaction = Trans
            .CommandType = CommandType.Text
            .CommandText = strSQL
        End With
        objCmd.ExecuteNonQuery()
    End Function

    Public Sub TransRollBack()
        Trans.Rollback()
    End Sub

    Public Sub TransCommit()
        Trans.Commit()
    End Sub

    Public Sub Close()
        objConn.Close()
        objConn = Nothing
    End Sub
End Class
