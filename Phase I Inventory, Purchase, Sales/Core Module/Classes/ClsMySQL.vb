Option Explicit On
Option Strict Off

Imports MySql.Data.MySqlClient
Imports MySql.Data
Imports System.IO

Imports Microsoft.VisualBasic.FileIO
Public Class ClsMySQL
    Dim FrmDia As New Windows.Forms.SaveFileDialog
    Public OnProgressBackUp As Boolean = False

    Public Function BackUpDataBase(ByVal strConn As String, ByVal vFileName As String, _
    Optional ByVal DBServer As String = "localhost", Optional ByVal Database As String = "core_db", _
    Optional ByVal DBUser As String = "root", Optional ByVal DBPass As String = "123", _
    Optional ByVal DBPort As String = "3307", Optional ByVal vFileName2 As String = "", Optional ByVal vBackUpDesc As String = "", _
    Optional ByVal Compress As Boolean = True, Optional ByVal vShutDownAfter As Boolean = False, Optional ByVal vEncrypt As Boolean = True, _
    Optional ByVal vEncryptPass As String = "", Optional ByVal vBackUpMatDash As Boolean = True) As Boolean
        Dim Frm As New FrmBackUpData

        Try
            Me.OnProgressBackUp = True

            If vEncryptPass = "" Then
                vEncryptPass = Core_Module.GetEncryptPass
            End If
            Frm.TopMost = False
            Frm.strConn = strConn
            Frm.vFileName = vFileName
            Frm.DBServer = DBServer
            Frm.Database = Database
            Frm.DBUser = DBUser
            Frm.DBPass = DBPass
            Frm.DBPort = DBPort
            Frm.vFileName2 = vFileName2
            Frm.vBackUpDesc = vBackUpDesc
            Frm.vCompress = Compress
            Frm.vShutdown = vShutDownAfter
            Frm.vEncrypt = vEncrypt
            Frm.vEncryptPass = vEncryptPass
            Frm.vBackUpMatdash = vBackUpMatDash

            Frm.ShowDialog()

            Me.OnProgressBackUp = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Frm = Nothing
        End Try
    End Function

    Public Sub ExecuteSqlScript(ByVal ConStr As String, _
   ByVal vInsertCommandList As List(Of String), ByVal vProgressBar As ProgressBar, ByVal ProgressControl As Windows.Forms.Control)
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection()
        Dim script As New MySql.Data.MySqlClient.MySqlScript()
        Dim vInsertCommand As String

        Try
            'open connection
            Conn.ConnectionString = ConStr
            script.Connection = Conn
            script.Query = ""
            Conn.Open()

            vProgressBar.Value = 0
            vProgressBar.Maximum = vInsertCommandList.Count

            For Each vInsertCommand In vInsertCommandList
                vProgressBar.Value += 1
                ProgressControl.Text = "Insert Data (" + vProgressBar.Value.ToString + " of " + vProgressBar.Maximum.ToString + ")"

                script.Query = vInsertCommand
                'execute script
                script.Execute()
            Next

            Conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Sub

    Public Function BuatSqlScript(ByVal vNamaTabel As String, ByVal conStrCabangUtama As String, Optional ByVal vSelectCommand As String = "") As List(Of String)
        Dim InsertCommand As String = ""
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection()
        Dim TBL As New DataTable
        Dim CMD As New MySql.Data.MySqlClient.MySqlCommand
        Dim Sql As String = ""
        Dim DTA As New MySql.Data.MySqlClient.MySqlDataAdapter
        Dim StartIndex, EndIndex, Index, BatchQty As Double
        Dim RetVal As New List(Of String)
        Dim row As DataRow
        Dim Value As String = ""
        Dim InsertCommand1 As String = ""

        Try
            Conn.ConnectionString = conStrCabangUtama
            InsertCommand = Me.BuatSQLInsertionScript(vNamaTabel, conStrCabangUtama, Conn)
            If vSelectCommand = "" Then
                'fill semua
                Sql = "SELECT * FROM " + vNamaTabel + ";"
            Else
                'fill dengan kondisi
                Sql = vSelectCommand
            End If
            CMD.CommandText = Sql
            CMD.Connection = Conn
            Conn.Open()
            DTA.SelectCommand = CMD

            DTA.Fill(TBL)

            StartIndex = 0
            Index = 0
            BatchQty = 500
            While Index <= TBL.Rows.Count - 1
                'perhitungan untuk looping, per 500 baris
                EndIndex = StartIndex + BatchQty - 1
                StartIndex = EndIndex + 1
                If EndIndex > TBL.Rows.Count - 1 Then
                    EndIndex = TBL.Rows.Count - 1
                End If

                'reset insert command
                InsertCommand1 = InsertCommand

                While Index <= EndIndex
                    row = TBL.Rows(Index)
                    Value = ""

                    For i As Integer = 0 To TBL.Columns.Count - 1
                        If Value = "" Then
                            If IsDBNull(row.Item(i)) = True Then
                                Value = "NULL"
                            Else
                                If row.Item(i).ToString = "" Then
                                    Value = "''"
                                Else
                                    If IsDate(row.Item(i).ToString) = True Then
                                        Value = "'" + Format(CDate(row.Item(i).ToString), "yyyy-MM-dd hh:mm:ss") + "'"
                                    Else
                                        If IsNumeric(row.Item(i)) = False Then
                                            'Cek string
                                            Dim ik As String = ""
                                            ik = Me.MengecekStringKarakterPetik(row.Item(i).ToString)
                                            Value = "'" + ik + "'"
                                        Else
                                            If row.Item(i).ToString = "False" Then
                                                Value = "0"
                                            Else
                                                If row.Item(i).ToString = "True" Then
                                                    Value = "1"
                                                Else
                                                    Value = row.Item(i).ToString
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            If IsDBNull(row.Item(i)) = True Then
                                Value = Value + ", " + "NULL"
                            Else
                                If row.Item(i).ToString = "" Then
                                    Value = Value + ", " + "''"
                                Else
                                    If IsDate(row.Item(i).ToString) = True Then
                                        Value = Value + ", " + "'" + Format(CDate(row.Item(i).ToString), "yyyy-MM-dd hh:mm:ss") + "'"
                                    Else
                                        If IsNumeric(row.Item(i)) = False Then
                                            'Cek string 
                                            Dim iz As String = ""
                                            iz = Me.MengecekStringKarakterPetik(row.Item(i).ToString)
                                            Value = Value + ", " + "'" + iz + "'"
                                        Else
                                            If row.Item(i).ToString = "False" Then
                                                Value = Value + ", " + "0"
                                            Else
                                                If row.Item(i).ToString = "True" Then
                                                    Value = Value + ", " + "1"
                                                Else
                                                    Value = Value + ", " + row.Item(i).ToString
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next

                    InsertCommand1 = InsertCommand1 + vbCrLf + "(" + Value + "),"
                    'pindah ke baris selanjutnya
                    Index += 1
                End While

                InsertCommand1 = InsertCommand1.Substring(0, InsertCommand1.Length - 1)
                InsertCommand1 = InsertCommand1 + ";" + vbCrLf + "COMMIT;"

                RetVal.Add(InsertCommand1)
            End While

            Conn.Close()

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Nothing
        End Try
    End Function

    Private Function BuatSQLInsertionScript(ByVal vNamaTabel As String, ByVal conStr As String, _
   ByVal Conn As MySql.Data.MySqlClient.MySqlConnection) As String
        Dim sql As String = ""
        Dim Reader As MySql.Data.MySqlClient.MySqlDataReader
        Dim Command1 As New MySql.Data.MySqlClient.MySqlCommand()
        Dim TBL As DataTable
        Dim InsertCommand As String = ""
        Dim NamaKolom As String = ""
        Dim Row As DataRow

        Try
            sql = "SELECT * FROM " + vNamaTabel + ";"
            InsertCommand = "INSERT INTO `" + vNamaTabel + "`"
            Command1.CommandText = sql
            Command1.Connection = Conn

            Conn.Open()
            Reader = Command1.ExecuteReader
            TBL = Reader.GetSchemaTable


            For Each Row In TBL.Rows
                If NamaKolom = "" Then
                    NamaKolom = "`" + Row(0).ToString + "`"
                Else
                    NamaKolom = NamaKolom + ", `" + Row(0).ToString + "`"
                End If
            Next

            InsertCommand = InsertCommand + " (" + NamaKolom + ") VALUES "

            Conn.Close()
            TBL.Dispose()

            Return InsertCommand
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            Command1.Dispose()
        End Try
    End Function

    Private Function MengecekStringKarakterPetik(ByVal karakter As String) As String
        Dim Retval As String = ""
        Dim i As Integer = 0
        Dim Index As Double = 0
        Dim KarakterDummy As String = ""
        Dim PenandaString As String = ""
        Dim ListString As New List(Of String)
        Dim ParamString As String = ""
        Try
            If karakter.IndexOf("'") > 0 Then
                While i <= karakter.Length - 1
                    If i = 0 Then
                        KarakterDummy = karakter.Substring(0, karakter.IndexOf("'") + 1)
                        ListString.Add(KarakterDummy + "'")
                        i = KarakterDummy.Length
                    Else
                        PenandaString = karakter.Substring(i)
                        If PenandaString.IndexOf("'") > 0 Then
                            KarakterDummy = PenandaString.Substring(i, PenandaString.IndexOf("'") + 1)
                            ListString.Add(KarakterDummy + "'")
                            i = KarakterDummy.Length
                        Else
                            ListString.Add(karakter.Substring(i))
                            Exit While
                        End If
                    End If
                End While
                For Each ParamString In ListString
                    Retval += ParamString
                Next
            Else
                Retval = karakter
            End If

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ListString = Nothing
        End Try
    End Function
End Class
