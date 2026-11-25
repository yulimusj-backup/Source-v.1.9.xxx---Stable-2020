''@rosadinnn
''201508191518
Option Explicit On
Option Strict Off

Imports System.Windows.Forms

Public Class ClsTrans_Attendance_Log

    Public Function DeleteLog(ByVal vFrom As Date, ByVal vTo As Date) As Integer
        Dim DtaSmessLog As New DseTrans_Attendance_LogTableAdapters.tbltrans_attendancelogTableAdapter
        Try
            DtaSmessLog.Connection.ConnectionString = Core_Module.GetConstring
            Return DtaSmessLog.Delete(vFrom, vTo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaSmessLog.Dispose()
        End Try
    End Function

    Public Function DeleteLog(ByVal vFrom As Date, ByVal vTo As Date, ByVal vListPrimaryServer As List(Of Double)) As Integer
        Dim DtaSmessLog As New DseTrans_Attendance_LogTableAdapters.tbltrans_attendancelogTableAdapter
        Dim DeletedCount_ As Integer = 0
        Try
            DtaSmessLog.Connection.ConnectionString = Core_Module.GetConstring
            For Each primarySrv As Double In vListPrimaryServer
                DeletedCount_ += DtaSmessLog.DeleteQueryBy_FromToPrimaryServer(vFrom, vTo, primarySrv)
            Next
            Return DeletedCount_
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaSmessLog.Dispose()
        End Try
    End Function

    Public Function InsertLog(ByVal vFrm As FrmTrans_Attendance_Log, ByVal vDgv As DataGridView, ByVal vUserCode As String) As Integer
        Dim DgvRow As DataGridViewRow
        Dim DtaSmessLog As New DseTrans_Attendance_LogTableAdapters.tbltrans_attendancelogTableAdapter
        Dim InsertedRow As Integer = 0
        Try
            DtaSmessLog.Connection.ConnectionString = Core_Module.GetConstring
            For Each DgvRow In vDgv.Rows
                With DgvRow
                    InsertedRow += DtaSmessLog.Insert(.Cells(vFrm.ColPrimaryEmployee.Index).Value, .Cells(vFrm.ColFingerprintID.Index).Value, _
                            .Cells(vFrm.ColPrimaryServer.Index).Value, .Cells(vFrm.ColTerminalID.Index).Value, CDate(.Cells(vFrm.ColTime.Index).Value), _
                            vUserCode, Now)
                End With
            Next
            Return InsertedRow
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaSmessLog.Dispose()
        End Try
    End Function

    Public Function InsertItemUsingScript(ByVal vFrm As FrmTrans_Attendance_Log, ByVal vDgv As DataGridView, ByVal vUserCode As String) As Integer
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(Core_Module.GetConstring)
        Dim Sql As String = ""
        Dim Script As New MySql.Data.MySqlClient.MySqlScript(Conn, Sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Dim Result As Integer = 0
        Try
            Conn.Open()
            While StartRowIndex < vDgv.Rows.Count
                Sql = Me.GetInsertString_Item(StartRowIndex, BatchQty, vDgv, vFrm, vUserCode)
                Script.Query = Sql
                Result += Script.Execute()
                StartRowIndex += BatchQty
            End While
            Conn.Close()

            Return vDgv.Rows.Count
        Catch ex As Exception

        End Try
    End Function

    Private Function GetInsertString_Item(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vDgv As DataGridView, ByVal vFrm As FrmTrans_Attendance_Log, ByVal vUserCode As String) As String
        Dim retval As String = ""
        Dim row As DataGridViewRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData
        Try
            retval = "INSERT INTO tbltrans_attendancelog (`primary employee`, `fingerprint id`, `primary attendance server`, `terminal id`, `transaction time`, `downloaded by`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vDgv.Rows.Count > vStartRowIndex + i Then
                    row = vDgv.Rows(i + vStartRowIndex)
                    If vDgv.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If
                    With row
                        retval = retval & vbCrLf + "(" & CDbl(.Cells(vFrm.ColPrimaryEmployee.Index).Value.ToString).ToString & ",'" _
                        & .Cells(vFrm.ColFingerprintID.Index).Value.ToString & "'," _
                        & CDbl(.Cells(vFrm.ColPrimaryServer.Index).Value.ToString).ToString & ",'" _
                        & .Cells(vFrm.ColTerminalID.Index).Value.ToString & "','" _
                        & Format(CDate(.Cells(vFrm.ColTime.Index).Value.ToString), "yyyy-MM-dd HH:mm:ss") & "','" _
                        & vUserCode & "')" _
                        & Delimiter
                    End With
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function
End Class
