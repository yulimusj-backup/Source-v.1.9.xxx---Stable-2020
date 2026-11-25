Option Explicit On
Option Strict Off

Imports MySql.Data.MySqlClient
Imports MySql.Data
Imports System.IO
Imports System.IO.Compression
Imports System.Security.Cryptography
Imports System.Media
Imports System
Imports System.Security

Public Class FrmBackUpData
    Dim ijk As Integer = 0
    Dim ObjEncrypt As New Core_Module.ClsEncryptFile
    Dim ObjZIP As New Core_Module.ClsCompressFile

    Public strConn As String = ""
    Public vFileName As String = ""
    Public vFileName2 As String = ""
    Public DBServer As String = "localhost"
    Public Database As String = "core_db1"
    Public DBUser As String = "root"
    Public DBPass As String = "123"
    Public DBPort As String = "3307"
    Public vBackUpDesc As String = ""
    Public vCompress As Boolean = True
    Public vShutdown As Boolean = False
    Public vEncrypt As Boolean = True
    Public vEncryptPass As String = ""
    Public vBackUpMatdash As Boolean = True

    Dim FinishBackUp As Boolean = False
    Dim BackUpEx As EnumTroubleBackUp = EnumTroubleBackUp.InvalidFolder

    Enum EnumTroubleBackUp As Integer
        InvalidFolder
    End Enum

    Private Sub FrmBackUpData_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Me.BringToFront()
    End Sub

    Private Sub FrmBackUpData_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.lblProgres.Text = ""
        Me.BringToFront()
        Control.CheckForIllegalCrossThreadCalls = False
        Me.BackgroundWorker1.RunWorkerAsync()
    End Sub

    Public Function BackUpDataBase1() As Boolean
        Dim connection As New MySqlConnection(strConn)
        Try
            connection.Open()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function BackUpDataBase() As Boolean
        Dim connection As New MySqlConnection(strConn)
        Dim tablesCommand As MySqlCommand = connection.CreateCommand()
        Dim rowsCommand As MySqlCommand = connection.CreateCommand()
        Dim rowsReader As MySqlDataReader
        Dim tablesReader As MySqlDataReader
        Dim DummyInsert As String = ""
        Dim Kelipatan As Double = 500
        Dim BatasAntara As Double = 0
        Dim tblDummy As New dsFormat.tblFormatDataTable

        Dim strCreateTbl As String = ""
        Dim strInsertTbl As String = ""
        Dim strInsert As String = ""
        Dim strInsert1 As String = ""
        Dim intR As Integer = 0
        Dim intK2 As Integer = 0
        Dim BackUpDesc1 As String = "Database Back Up"
        Dim counter1 As Integer = 0
        Dim fileExists As Boolean
        Dim FolderName1 As String = ""
        Dim FolderName2 As String = ""
        Dim IndexOf As Integer = 0
        Dim bytKey As Byte()
        Dim bytIV As Byte()
        Dim ExtentionLength As Integer = 0
        Dim FileNameDummy As String = ""
        Dim FileNameDummy1 As String = ""
        Dim Lanjut As Boolean = True

        Try
            If vCompress = True Then
                'kalo kompres, pasti encrypt
                ExtentionLength = 6
                If vEncrypt = False Then
                    vEncrypt = True
                    vEncryptPass = Core_Module.GetEncryptPass
                End If
            Else
                ExtentionLength = 4
            End If
            FileNameDummy = Me.vFileName.Substring(0, Me.vFileName.Length - ExtentionLength) + " - " + Format(Now, "dd MMM yyyy - HH_mm") + Me.vFileName.Substring(Me.vFileName.Length - ExtentionLength)
            FileNameDummy1 = Me.vFileName.Substring(0, Me.vFileName.Length - ExtentionLength) + " - " + Format(Now, "dd MMM yyyy - HH_mm") + Me.vFileName.Substring(Me.vFileName.Length - ExtentionLength)

            Me.lblProgres.Text = ""
            'pastikan direktory yang dimau ada
            If vFileName2 <> "" Then
                IndexOf = vFileName.LastIndexOf("\")
                FolderName1 = vFileName.Substring(0, IndexOf)
                IndexOf = vFileName2.LastIndexOf("\")
                FolderName2 = vFileName2.Substring(0, IndexOf)
                If My.Computer.FileSystem.DirectoryExists(FolderName1) = False And _
                My.Computer.FileSystem.DirectoryExists(FolderName2) = False Then
                    BackUpEx = EnumTroubleBackUp.InvalidFolder
                    Return False
                End If
            Else
                IndexOf = vFileName.LastIndexOf("\")
                FolderName1 = vFileName.Substring(0, IndexOf)
                If My.Computer.FileSystem.DirectoryExists(FolderName1) = False Then
                    BackUpEx = EnumTroubleBackUp.InvalidFolder
                    Return False
                End If
            End If


            If Me.vFileName2 <> "" Then
                Me.vFileName2 = Me.vFileName2.Substring(0, Me.vFileName2.Length - ExtentionLength) + " - " + Format(Now, "dd MMM yyyy HH_mm") + Me.vFileName2.Substring(Me.vFileName2.Length - ExtentionLength)
            End If
            'Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'back up di data directory dulu
            Me.vFileName = My.Application.Info.DirectoryPath + "\AutobackUp.sql"
            'hapus file sementara yang sudah ada
            If My.Computer.FileSystem.FileExists(Me.vFileName) = True Then
                Me.lblProgres.Text = "Preparing...."
                My.Computer.FileSystem.DeleteFile(Me.vFileName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If

            'desc back up
            If Me.vBackUpDesc <> "" Then
                BackUpDesc1 = BackUpDesc1 + " :" + Me.vBackUpDesc
            End If
            connection.Open()

            If Database.ToLower = "smess_lmb" Then
                Dim _connection As New MySqlConnection(connection.ConnectionString)
                Dim _command As MySqlCommand
                Try
                    _connection.Open()
                    If _connection.State = ConnectionState.Open Then
                        _command = New MySqlCommand("flush hosts;", _connection)
                        _command.ExecuteNonQuery()
                    End If
                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                Finally
                    If _connection.State = ConnectionState.Open Then
                        _connection.Close()
                    End If
                End Try
            End If

            '  tablesCommand.CommandText = "SHOW FULL TABLES FROM " & Database & " WHERE Table_type <> 'VIEW' And Tables_In_" & Database & " = 'tbltrans_Kartu'"
            tablesCommand.CommandText = "SHOW FULL TABLES FROM " & Database & " WHERE Table_type <> 'VIEW'"
            Dim queryAll As String = ""
            Dim query As String = ""
            Dim query2 As String = ""
            tablesReader = tablesCommand.ExecuteReader()
            Dim tablesArray As New ArrayList()
            While tablesReader.Read()
                tablesArray.Add(tablesReader.GetValue(0).ToString())
            End While
            tablesReader.Close()
            tablesCommand.Dispose()

            tablesCommand.CommandText = "SHOW FULL TABLES FROM " & Database & " WHERE Table_type = 'VIEW'"
            tablesReader = tablesCommand.ExecuteReader()
            Dim tablesArray2 As New ArrayList()
            While tablesReader.Read()
                tablesArray2.Add(tablesReader.GetValue(0).ToString())
            End While
            tablesReader.Close()
            tablesCommand.Dispose()


            Dim intTotalT As Integer = 0
            intTotalT = tablesArray.Count + tablesArray2.Count

            Dim strCreateDb As String = ""
            strCreateDb += "#" + BackUpDesc1 & vbCrLf & _
                     "# ---------------------------------------" & vbCrLf & _
                     "# Host     : " & DBServer & vbCrLf & _
                     "# Port     : " & DBPort & vbCrLf & _
                     "# Database : " & Database & vbCrLf & vbCrLf

            '===========================================================================================================================
            '---------------------------------------------------------------------------------------------------------------------------
            'QUERY untuk DROP & CREATED DATABASE
            '---------------------------------------------------------------------------------------------------------------------------
            strCreateDb += "SET FOREIGN_KEY_CHECKS=0;" & vbCrLf & vbCrLf
            strCreateDb += "DROP DATABASE IF EXISTS `" & Database & "`;" & vbCrLf & vbCrLf
            strCreateDb += "CREATE DATABASE `" & Database & "` " & vbCrLf

            rowsCommand.CommandText = "SHOW VARIABLES "
            rowsReader = rowsCommand.ExecuteReader()
            While rowsReader.Read()
                If rowsReader.GetString("Variable_name").ToString = "character_set_database" Then
                    strCreateDb += "      CHARACTER SET " & rowsReader.GetString("Value").ToString & vbCrLf
                End If
                If rowsReader.GetString("Variable_name").ToString = "collation_database" Then
                    strCreateDb += "      COLLATE " & rowsReader.GetString("Value").ToString & ";" & vbCrLf & vbCrLf
                End If

            End While
            rowsReader.Close()
            rowsCommand.Dispose()

            strCreateDb += "USE `" & Database & "`;" & vbCrLf & vbCrLf

            '---------------------------------------------------------------------------------------------------------------------------
            '===========================================================================================================================

            '----------------------------------------------------------------------------------
            'untuk progres bar
            Me.ProgressBar1.Show()
            Me.ProgressBar1.Maximum = (tablesArray.Count) + (tablesArray2.Count)
            Me.ProgressBar1.Value = 0
            '----------------------------------------------------------------------------------

            'create view
            Dim strCreateView As String = ""
            Dim intI As Integer = 0
            If tablesArray2.Count > 0 Then
                For intI = 0 To tablesArray2.Count - 1
                    '===========================================================================================================================
                    '---------------------------------------------------------------------------------------------------------------------------
                    'QUERY untuk DROP & CREATED TIAP VIEW
                    '---------------------------------------------------------------------------------------------------------------------------
                    'coment table query --------------------------------------------------------------------------------------------------------
                    strCreateView += "#" & vbCrLf & "# definition for the " & tablesArray2(intI) & " VIEW : " & vbCrLf & "#" & vbCrLf & vbCrLf
                    'Begin to create table -----------------------------------------------------------------------------------------------------
                    'drop table query ----------------------------------------------------------------------------------------------------------
                    strCreateView += "DROP VIEW IF EXISTS `" & tablesArray2(intI) & "`;" & vbCrLf & vbCrLf
                    rowsCommand.CommandText = "SHOW CREATE VIEW " & tablesArray2(intI)
                    rowsReader = rowsCommand.ExecuteReader()
                    While rowsReader.Read()
                        'direplace, karea terlalu banyak tulisan database yang tidak perlu
                        strCreateView += rowsReader.GetString("Create view").Replace("`" + Database + "`.", "")
                    End While
                    rowsReader.Close()
                    rowsCommand.Dispose()
                    strCreateView += ";" & vbCrLf & vbCrLf
                    '---------------------------------------------------------------------------------------------------------------------------
                    '===========================================================================================================================
                    Me.ProgressBar1.Value += 1
                    Me.lblProgres.Text = ""
                    Me.lblProgres.Text = "Processing :" & Format(((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100), "##0") & "%"
                Next
            End If

            'create tabel
            Dim intA As Integer = 0
            Me.lblProgres.Text = ""
            For intA = 0 To tablesArray.Count - 1
                'If tablesArray(intA) <> "tbltrans_purchasereturn" Then
                ''===========================================================================================================================
                '---------------------------------------------------------------------------------------------------------------------------
                'QUERY untuk DROP & CREATED TIAP TABLE
                '---------------------------------------------------------------------------------------------------------------------------
                'coment table query --------------------------------------------------------------------------------------------------------
                strCreateTbl += "#" & vbCrLf & "# Structure for the " & tablesArray(intA) & " table : " & vbCrLf & "#" & vbCrLf & vbCrLf
                'Begin to create table -----------------------------------------------------------------------------------------------------
                'drop table query ----------------------------------------------------------------------------------------------------------
                strCreateTbl += "DROP TABLE IF EXISTS `" & tablesArray(intA) & "`;" & vbCrLf & vbCrLf
                rowsCommand.CommandText = "SHOW CREATE TABLE " & tablesArray(intA)
                rowsReader = rowsCommand.ExecuteReader()
                While rowsReader.Read()
                    strCreateTbl += rowsReader.GetString("Create table")
                End While
                rowsReader.Close()
                rowsCommand.Dispose()
                strCreateTbl += ";" & vbCrLf & vbCrLf
            Next

            'atur posisi hasil query
            queryAll = strCreateDb + strCreateTbl + strCreateView

            If File.Exists(vFileName) Then
                File.Delete(vFileName)
            End If

            Dim objWriter As New System.IO.StreamWriter(vFileName)
            objWriter.Write(queryAll)
            objWriter.Close()

            For intA = 0 To tablesArray.Count - 1
                Lanjut = True
                If Me.vBackUpMatdash = False Then
                    If tablesArray(intA).ToString.Contains("tbltrans_material_planning") = True Then
                        Lanjut = False
                    End If
                End If
                ' If tablesArray(intA) = "tblmaster_product" Then
                '---------------------------------------------------------------------------------------------------------------------------
                '===========================================================================================================================


                '===========================================================================================================================
                '---------------------------------------------------------------------------------------------------------------------------
                'QUERY untuk  INSERT TIAP TABLE
                '---------------------------------------------------------------------------------------------------------------------------
                Dim strCount As String = ""
                Dim sqlCount As String = "Select Count(*) From " & tablesArray(intA)
                If Lanjut = True Then
                    strCount = Me.SQLE(sqlCount, 0)
                    Dim decCount As Decimal = 0
                    If IsNumeric(strCount) Then
                        decCount = CDec(strCount)
                    End If


                    Dim dt As New DataTable
                    If decCount < 250000 Then
                        Me.lblProgres.Text = "Processing :" & Format(((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100), "##0") & _
                                       "% (Loading data for " & tablesArray(intA) + ")"
                        Dim sql As String = "SELECT * FROM " & tablesArray(intA) & " Limit 0, 250000 "
                        Dim da As New MySqlDataAdapter(sql, connection)
                        Dim intCol As Integer = 0
                        da.Fill(dt)

                        Dim intK As Integer = 0
                        If dt.Rows.Count > 0 Then
                            '---------------------------------------------------------------------------------------------------------------------------
                            'CODE TAMBAHAN -------------------------------------------------------------------------------------------------------------
                            tblDummy.Rows.Clear()
                            rowsCommand.CommandText = "DESCRIBE " & tablesArray(intA)
                            rowsReader = rowsCommand.ExecuteReader()
                            intCol = 0
                            For intCol = 0 To dt.Columns.Count - 1
                                While rowsReader.Read()
                                    tblDummy.Rows.Add(rowsReader.GetString("Type").ToString.ToUpper, intCol)
                                    Exit While
                                End While
                            Next
                            rowsCommand.Dispose()
                            rowsReader.Close()
                            '---------------------------------------------------------------------------------------------------------------------------
                            '---------------------------------------------------------------------------------------------------------------------------

                            '---------------------------------------------------------------------------------------------------------------------------
                            'coment table query --------------------------------------------------------------------------------------------------------
                            strInsert1 = ""
                            strInsert1 += "#" & vbCrLf & "# data for the " & tablesArray(intA) & " TABLE : " & vbCrLf & "#" & vbCrLf & vbCrLf
                            '---------------------------------------------------------------------------------------------------------------------------
                            '--> DEFINISIKAN TIAP KOLOM
                            strInsert1 += "INSERT INTO `" & tablesArray(intA) & "` ("
                            For intK = 0 To dt.Columns.Count - 1
                                If intK <> 0 Then
                                    strInsert1 += ","
                                End If
                                strInsert1 += "`" & dt.Columns(intK).ToString & "`"
                                If intK = dt.Columns.Count - 1 Then
                                    strInsert1 += ")"
                                End If
                            Next
                            strInsert1 += " VALUES " & vbCrLf
                            '---------------------------------------------------------------------------------------------------------------------------
                            '---------------------------------------------------------------------------------------------------------------------------

                            DummyInsert = strInsert1
                            Kelipatan = 500
                            BatasAntara = 0

                            '--> DEFINISIKAN TIAP ROW
                            BatasAntara += Kelipatan
                            For intR = 0 To dt.Rows.Count - 1
                                strInsert = ""
                                For intK2 = 0 To dt.Columns.Count - 1
                                    If intK2 <> 0 Then
                                        strInsert += ","
                                    End If
                                    strInsert += FormatValue(tblDummy, intK2, dt.Rows(intR)(intK2))
                                Next

                                If intK2 = dt.Columns.Count - 1 Then
                                    strInsert += vbCrLf
                                End If
                                If intR = dt.Rows.Count - 1 Then
                                    strInsert1 += "(" & strInsert & ");" & vbCrLf & vbCrLf
                                    strInsert1 += "COMMIT;" & vbCrLf & vbCrLf & vbCrLf
                                    My.Computer.FileSystem.WriteAllText(vFileName, strInsert1, True, System.Text.Encoding.UTF8)
                                    'My.Computer.FileSystem.WriteAllText("C:\New folder\Test.txt", vbCrLf + strInsert1, True, System.Text.Encoding.UTF8)
                                    'strInsertTbl += strInsert1
                                Else
                                    If intR >= BatasAntara Then
                                        strInsert1 += "(" & strInsert & ");" & vbCrLf & vbCrLf
                                        strInsert1 += "COMMIT;" & vbCrLf & vbCrLf & vbCrLf
                                        BatasAntara += Kelipatan
                                        My.Computer.FileSystem.WriteAllText(vFileName, strInsert1, True, System.Text.Encoding.UTF8)
                                        'My.Computer.FileSystem.WriteAllText("C:\New folder\Test.txt", vbCrLf + strInsert1, True, System.Text.Encoding.UTF8)
                                        'strInsertTbl += strInsert1
                                        strInsert1 = DummyInsert
                                    Else
                                        strInsert1 += "(" & strInsert & ")," & vbCrLf
                                        ' My.Computer.FileSystem.WriteAllText("C:\New folder\Test.txt", vbCrLf + strInsert1, True, System.Text.Encoding.UTF8)
                                    End If
                                End If

                                If intR + 1 > 500 Then
                                    If Math.IEEERemainder((intR + 1), 500) = 0 Then
                                        Me.lblProgres.Text = "Processing :" & Format(((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100), "##0") & _
                                        "% (" & tablesArray(intA) + " - row " + (intR + 1).ToString + " of " + (dt.Rows.Count).ToString + ")"
                                    End If
                                Else
                                    If intR + 1 = 1 Then
                                        Me.lblProgres.Text = "Processing :" & Format(((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100), "##0") & _
                                        "% (" & tablesArray(intA) + " - row " + (intR + 1).ToString + " of " + dt.Rows.Count.ToString + ")"
                                    Else
                                        If Math.IEEERemainder((intR + 1), 25) = 0 Then
                                            Me.lblProgres.Text = "Processing :" & Format(((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100), "##0") & _
                                            "% (" & tablesArray(intA) + " - row " + (intR + 1).ToString + " of " + dt.Rows.Count.ToString + ")"
                                        End If
                                    End If
                                End If
                            Next
                        End If

                        da.Dispose()
                        dt.Clear()
                        dt.Dispose()
                        '---------------------------------------------------------------------------------------------------------------------------
                        '===========================================================================================================================

                    Else
                        Dim decX As Decimal = 0
                        decX = Fix(decCount / 250000)
                        If CDec(decCount / 250000) > Fix(decCount / 250000) Then
                            decX += 1
                        End If

                        For intZ As Integer = 0 To decX - 1
                            Dim sql As String = "SELECT * FROM " & tablesArray(intA) & " Limit " & (intZ * 250000) & ", " & 250000
                            Dim da As New MySqlDataAdapter(sql, connection)
                            Dim intCol As Integer = 0
                            Me.lblProgres.Text = "Processing :" & Format(((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100), "##0") & _
                                           "% (Loading data for " & tablesArray(intA) + ")"
                            da.Fill(dt)

                            Dim intK As Integer = 0
                            If dt.Rows.Count > 0 Then
                                '---------------------------------------------------------------------------------------------------------------------------
                                'CODE TAMBAHAN -------------------------------------------------------------------------------------------------------------
                                tblDummy.Rows.Clear()
                                rowsCommand.CommandText = "DESCRIBE " & tablesArray(intA)
                                rowsReader = rowsCommand.ExecuteReader()
                                intCol = 0
                                For intCol = 0 To dt.Columns.Count - 1
                                    While rowsReader.Read()
                                        tblDummy.Rows.Add(rowsReader.GetString("Type").ToString.ToUpper, intCol)
                                        Exit While
                                    End While
                                Next
                                rowsCommand.Dispose()
                                rowsReader.Close()
                                '---------------------------------------------------------------------------------------------------------------------------
                                '---------------------------------------------------------------------------------------------------------------------------

                                '---------------------------------------------------------------------------------------------------------------------------
                                'coment table query --------------------------------------------------------------------------------------------------------
                                strInsert1 = ""
                                strInsert1 += "#" & vbCrLf & "# data for the " & tablesArray(intA) & " TABLE : " & vbCrLf & "#" & vbCrLf & vbCrLf
                                '---------------------------------------------------------------------------------------------------------------------------
                                '--> DEFINISIKAN TIAP KOLOM
                                strInsert1 += "INSERT INTO `" & tablesArray(intA) & "` ("
                                For intK = 0 To dt.Columns.Count - 1
                                    If intK <> 0 Then
                                        strInsert1 += ","
                                    End If
                                    strInsert1 += "`" & dt.Columns(intK).ToString & "`"
                                    If intK = dt.Columns.Count - 1 Then
                                        strInsert1 += ")"
                                    End If
                                Next
                                strInsert1 += " VALUES " & vbCrLf
                                '---------------------------------------------------------------------------------------------------------------------------
                                '---------------------------------------------------------------------------------------------------------------------------

                                DummyInsert = strInsert1
                                Kelipatan = 500
                                BatasAntara = 0

                                '--> DEFINISIKAN TIAP ROW
                                intR = 0
                                BatasAntara += Kelipatan
                                For intR = 0 To dt.Rows.Count - 1
                                    strInsert = ""
                                    For intK2 = 0 To dt.Columns.Count - 1
                                        If intK2 <> 0 Then
                                            strInsert += ","
                                        End If
                                        strInsert += FormatValue(tblDummy, intK2, dt.Rows(intR)(intK2))
                                    Next

                                    If intK2 = dt.Columns.Count - 1 Then
                                        strInsert += vbCrLf
                                    End If
                                    If intR = dt.Rows.Count - 1 Then
                                        strInsert1 += "(" & strInsert & ");" & vbCrLf & vbCrLf
                                        strInsert1 += "COMMIT;" & vbCrLf & vbCrLf & vbCrLf
                                        My.Computer.FileSystem.WriteAllText(vFileName, strInsert1, True, System.Text.Encoding.UTF8)
                                        'My.Computer.FileSystem.WriteAllText("C:\New folder\Test.txt", vbCrLf + strInsert1, True, System.Text.Encoding.UTF8)
                                        'strInsertTbl += strInsert1
                                    Else
                                        If intR >= BatasAntara Then
                                            strInsert1 += "(" & strInsert & ");" & vbCrLf & vbCrLf
                                            strInsert1 += "COMMIT;" & vbCrLf & vbCrLf & vbCrLf
                                            BatasAntara += Kelipatan
                                            My.Computer.FileSystem.WriteAllText(vFileName, strInsert1, True, System.Text.Encoding.UTF8)
                                            'My.Computer.FileSystem.WriteAllText("C:\New folder\Test.txt", vbCrLf + strInsert1, True, System.Text.Encoding.UTF8)
                                            'strInsertTbl += strInsert1
                                            strInsert1 = DummyInsert
                                        Else
                                            strInsert1 += "(" & strInsert & ")," & vbCrLf
                                            ' My.Computer.FileSystem.WriteAllText("C:\New folder\Test.txt", vbCrLf + strInsert1, True, System.Text.Encoding.UTF8)
                                        End If
                                    End If

                                    'Me.lblProgres.Text = "Processing :" & Format(((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100), "##0") & _
                                    '                    "% (" & tablesArray(intA) + " - row " & (CDec((intZ * 250000) + intR + 1)).ToString & " of " & decCount & ")"

                                    If intR + 1 > 500 Then
                                        If Math.IEEERemainder((intR + 1), 500) = 0 Then
                                            Me.lblProgres.Text = "Processing :" & Format(((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100), "##0") & _
                                                                "% (" & tablesArray(intA) + " - row " & (CDec((intZ * 250000) + intR + 1)).ToString & " of " & decCount & ")"
                                            'Me.lblProgres.Text = "Processing :" & Format(((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100), "##0") & _
                                            '                                      "% (" & tablesArray(intA) + " - row " + (intR + 1).ToString + " of " + (dt.Rows.Count).ToString + ")"
                                        End If
                                    Else
                                        If intR + 1 = 1 Then
                                            Me.lblProgres.Text = "Processing :" & Format(((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100), "##0") & _
                                                               "% (" & tablesArray(intA) + " - row " & (CDec((intZ * 250000) + intR + 1)).ToString & " of " & decCount & ")"
                                            'Me.lblProgres.Text = "Processing :" & Format(((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100), "##0") & _
                                            '"% (" & tablesArray(intA) + " - row " + (intR + 1).ToString + " of " + dt.Rows.Count.ToString + ")"
                                        Else
                                            If Math.IEEERemainder((intR + 1), 25) = 0 Then
                                                'Me.lblProgres.Text = "Processing :" & Format(((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100), "##0") & _
                                                '"% (" & tablesArray(intA) + " - row " + (intR + 1).ToString + " of " + dt.Rows.Count.ToString + ")"
                                                Me.lblProgres.Text = "Processing :" & Format(((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100), "##0") & _
                                                           "% (" & tablesArray(intA) + " - row " & (CDec((intZ * 250000) + intR + 1)).ToString & " of " & decCount & ")"

                                            End If
                                        End If
                                    End If

                                    'If intZ = (decX - 1) Then
                                    'If intR + 1 > 500 Then
                                    '    If Math.IEEERemainder((intR + 1), 500) = 0 Then
                                    '        Me.lblProgres.Text = "Processing :" & Format(((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100), "##0") & _
                                    '        "% (" & tablesArray(intA) + " - row " & (CDec((intZ * 250000) + intR + 1)).ToString & " of " & decCount & ")"
                                    '    End If
                                    'Else
                                    '    If intR + 1 = 1 Then
                                    '        Me.lblProgres.Text = "Processing :" & Format(((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100), "##0") & _
                                    '        "% (" & tablesArray(intA) + " - row " & (CDec((intZ * 250000) + intR + 1)).ToString & " of " & decCount & ")"
                                    '    Else
                                    '        If Math.IEEERemainder((intR + 1), 25) = 0 Then
                                    '            Me.lblProgres.Text = "Processing :" & Format(((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100), "##0") & _
                                    '            "% (" & tablesArray(intA) + " - row " & (CDec((intZ * 250000) + intR + 1)).ToString & " of " & decCount & ")"
                                    '        End If
                                    '    End If
                                    'End If
                                    'Else
                                    '    If intR + 1 > 500 Then
                                    '        If Math.IEEERemainder((intR + 1), 500) = 0 Then
                                    '            Me.lblProgres.Text = "Processing :" & Format(((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100), "##0") & _
                                    '            "% (" & tablesArray(intA) + " - row " & (CDec((intZ * 250000) + intR + 1)).ToString & " of " & decCount & ")"
                                    '        End If
                                    '    Else
                                    '        If intR + 1 = 1 Then
                                    '            Me.lblProgres.Text = "Processing :" & Format(((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100), "##0") & _
                                    '            "% (" & tablesArray(intA) + " - row " & (CDec((intZ * 250000) + intR + 1)).ToString & " of " & decCount & ")"
                                    '        Else
                                    '            If Math.IEEERemainder((intR + 1), 25) = 0 Then
                                    '                Me.lblProgres.Text = "Processing :" & Format(((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100), "##0") & _
                                    '                "% (" & tablesArray(intA) + " - row " & (CDec((intZ * 250000) + intR + 1)).ToString & " of " & decCount & ")"
                                    '            End If
                                    '        End If
                                    '    End If
                                    'End If


                                Next
                            End If

                            da.Dispose()
                            dt.Clear()
                            dt.Dispose()
                            '---------------------------------------------------------------------------------------------------------------------------
                            '===========================================================================================================================
                            ' End If
                            ' End If
                        Next
                        '  Frm.lblProgres.Text = "Proses :" & Format((((intA + 1) / intTotalT) * 100), "##0") & "%"

                        'Me.BackgroundWorker1.ReportProgress((intA / tablesArray.Count) * 100)
                    End If
                End If

                Me.ProgressBar1.Value += 1
                Me.lblProgres.Text = ""
            Next

            'copykan data ke file tujuan
            counter1 = 0
            fileExists = My.Computer.FileSystem.FileExists(FileNameDummy)
            While fileExists = True
                counter1 += 1
                FileNameDummy = FileNameDummy1 + "(" + counter1.ToString + ")"
                fileExists = My.Computer.FileSystem.FileExists(FileNameDummy)
            End While

            'compress file
            If Me.vCompress = True Then
                If My.Computer.FileSystem.FileExists(Me.vFileName.Substring(0, Me.vFileName.Length - ExtentionLength) + ".zip") = True Then
                    My.Computer.FileSystem.DeleteFile(Me.vFileName.Substring(0, Me.vFileName.Length - ExtentionLength) + ".zip", FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                End If
                Me.lblProgres.Text = "Compressing file..."
                Me.ObjZIP.ZipFile(Me.vFileName, Me.vFileName.Substring(0, Me.vFileName.Length - ExtentionLength) + ".zip")
                File.SetAttributes(Me.vFileName.Substring(0, Me.vFileName.Length - ExtentionLength) + ".zip", FileAttributes.Hidden)
                My.Computer.FileSystem.DeleteFile(Me.vFileName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)

                'encrypt
                'Send the password to the CreateKey function.
                bytKey = Me.ObjEncrypt.CreateKey(Me.vEncryptPass)
                'Send the password to the CreateIV function.
                bytIV = Me.ObjEncrypt.CreateIV(Me.vEncryptPass)
                If vEncrypt = True Then
                    Me.lblProgres.Text = "Encrypting file..."
                    If My.Computer.FileSystem.FileExists(Me.vFileName.Substring(0, Me.vFileName.Length - ExtentionLength) + ".smess") = True Then
                        My.Computer.FileSystem.DeleteFile(Me.vFileName.Substring(0, Me.vFileName.Length - ExtentionLength) + ".smess", FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                    End If
                    'Start the encryption.
                    Me.ObjEncrypt.EncryptOrDecryptFile(Me.vFileName.Substring(0, Me.vFileName.Length - ExtentionLength) + ".zip", Me.vFileName.Substring(0, Me.vFileName.Length - ExtentionLength) + ".smess", _
                    bytKey, bytIV, ClsEncryptFile.CryptoAction.ActionEncrypt, Me.ProgressBar1)
                    File.SetAttributes(Me.vFileName.Substring(0, Me.vFileName.Length - ExtentionLength) + ".smess", FileAttributes.Hidden)
                    'hapus file hasil compression
                    If My.Computer.FileSystem.FileExists(Me.vFileName.Substring(0, Me.vFileName.Length - ExtentionLength) + ".zip") = True Then
                        My.Computer.FileSystem.DeleteFile(Me.vFileName.Substring(0, Me.vFileName.Length - ExtentionLength) + ".zip", FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                    End If
                End If

                'copy hasil ke folder tujuan
                Me.lblProgres.Text = "Copying file to: " + FileNameDummy
                My.Computer.FileSystem.CopyFile(Me.vFileName.Substring(0, Me.vFileName.Length - ExtentionLength) + ".smess", FileNameDummy, FileIO.UIOption.OnlyErrorDialogs, FileIO.UICancelOption.DoNothing)
                File.SetAttributes(FileNameDummy, FileAttributes.Normal)
                If Me.vFileName2 <> "" Then
                    Me.lblProgres.Text = "Copying file to: " + Me.vFileName2
                    My.Computer.FileSystem.CopyFile(Me.vFileName.Substring(0, Me.vFileName.Length - ExtentionLength) + ".smess", Me.vFileName2, FileIO.UIOption.OnlyErrorDialogs, FileIO.UICancelOption.DoNothing)
                    File.SetAttributes(Me.vFileName2, FileAttributes.Normal)
                End If
            Else
                Me.lblProgres.Text = "Copying file to: " + FileNameDummy
                My.Computer.FileSystem.CopyFile(Me.vFileName, FileNameDummy, FileIO.UIOption.OnlyErrorDialogs, FileIO.UICancelOption.DoNothing)

                If Me.vFileName2 <> "" Then
                    Me.lblProgres.Text = "Copying file to: " + Me.vFileName2
                    My.Computer.FileSystem.CopyFile(Me.vFileName, Me.vFileName2, FileIO.UIOption.OnlyErrorDialogs, FileIO.UICancelOption.DoNothing)
                End If
                'hapus file sementara yang sudah ada
                My.Computer.FileSystem.DeleteFile(Me.vFileName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If

            'Windows.Forms.Cursor.Current = Cursors.Arrow

            Me.FinishBackUp = True

            'Dim connection1 As New MySqlConnection(strConn)
            'connection1.Open()
            'connection.Open()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function SQLE(ByVal Sql As String, ByVal index As Integer) As String
        Dim conn As MySqlClient.MySqlConnection
        Dim cmd As MySqlClient.MySqlCommand
        Dim Readers As MySqlClient.MySqlDataReader
        conn = New MySqlClient.MySqlConnection(strConn)
        Try
            conn.Open()
            cmd = New MySqlClient.MySqlCommand(Sql, conn)
            Readers = cmd.ExecuteReader(CommandBehavior.SingleRow)
            While Readers.Read()
                Return Readers.GetValue(index)
            End While
            Readers.Close()

            Return ""
        Catch ex As Exception
            MessageBox.Show("ERROR in SQLE!" & Sql, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return ""
        Finally
            conn.Close()
        End Try
    End Function

    Public Function FormatValue(ByVal table As DataTable, ByVal intRow As Integer, ByVal value As Object) As String
        Dim strValue As String = ""
        Dim DummyString As String = ""

        Try
            If IsDBNull(value) Then
                strValue += "NULL"
            ElseIf (table.Rows(intRow)(0).Length >= 6 AndAlso _
               table.Rows(intRow)(0).ToString.Substring(0, 6) = "DOUBLE") Then
                strValue += CDbl(value).ToString
            ElseIf (table.Rows(intRow)(0).Length >= 7 AndAlso _
               table.Rows(intRow)(0).ToString.Substring(0, 7) = "DECIMAL") Then
                strValue += CDec(value).ToString
            ElseIf (table.Rows(intRow)(0).Length >= 7 AndAlso _
               table.Rows(intRow)(0).ToString.Substring(0, 7) = "VARCHAR") Then
                DummyString = value.ToString
                If DummyString <> "" Then
                    If DummyString.Contains("\") = True Then
                        DummyString = DummyString.Replace("\", "\\")
                    End If
                    If DummyString.Contains(vbCrLf) = True Then
                        DummyString = DummyString.Replace(vbCrLf, "\r\n")
                    End If
                    If DummyString.Contains("'") = True Then
                        DummyString = DummyString.Replace("'", "''")
                    End If
                    If DummyString.Contains("Â‚") Then
                        DummyString = DummyString.Replace("Â‚", "")
                    End If
                    If DummyString.Contains("Ã,") Then
                        DummyString = DummyString.Replace("Ã,", "")
                    End If
                    If DummyString.Contains("Ãƒ") Then
                        DummyString = DummyString.Replace("Ãƒ", "")
                    End If
                    If DummyString.Contains("Â") = True Then
                        DummyString = DummyString.Replace("Â", "")
                        'My.Computer.FileSystem.WriteAllText("C:\New folder\Test.txt", vbCrLf + "'" + DummyString + "'", True, System.Text.Encoding.UTF8)
                    End If
                    If DummyString.Contains("Ã") Then
                        DummyString = DummyString.Replace("Ã", "")
                        'My.Computer.FileSystem.WriteAllText("C:\New folder\Test.txt", vbCrLf + "'" + DummyString + "'", True, System.Text.Encoding.UTF8)
                    End If
                    'DummyString = "Turbiquant ®1500IR/1500T Calibration standard Set"
                End If
                strValue += "'" & DummyString & "'"
            ElseIf table.Rows(intRow)(0) = "DATE" Then
                If IsDate(value) Then
                    strValue += "'" & Format(CDate(value), "yyyy-MM-dd") & "'"
                Else
                    strValue += "NULL"
                End If
            ElseIf table.Rows(intRow)(0) = "DATETIME" Then
                If IsDate(value) Then
                    strValue += "'" & Format(CDate(value), "yyyy-MM-dd HH:mm:ss") & "'"
                Else
                    strValue += "NULL"
                End If
            ElseIf table.Rows(intRow)(0) = "TIMESTAMP" Then
                If IsDate(value) Then
                    strValue += "'" & Format(CDate(value), "yyyy-MM-dd HH:mm:ss") & "'"
                Else
                    strValue += "NULL"
                End If
            ElseIf table.Rows(intRow)(0).Length >= 3 AndAlso _
                    table.Rows(intRow)(0).ToString.Substring(0, 3) = "BIT" Then
                If value = 1 Then
                    strValue += "True"
                Else
                    strValue += "False"
                End If
            ElseIf (table.Rows(intRow)(0).Length >= 4 AndAlso _
                   table.Rows(intRow)(0).ToString.Substring(0, 4) = "CHAR") Then
                strValue += "'" & value.ToString & "'"
            ElseIf (table.Rows(intRow)(0).Length >= 8 AndAlso _
                     table.Rows(intRow)(0).ToString.Substring(0, 8) = "LONGTEXT") Then
                strValue += "'" & value.ToString & "'"

            ElseIf (table.Rows(intRow)(0).Length >= 9 AndAlso _
                    table.Rows(intRow)(0).ToString.Substring(0, 9) = "MEDIUMTEXT") Then
                strValue += "'" & value.ToString & "'"

            ElseIf (table.Rows(intRow)(0).Length >= 4 AndAlso _
                    table.Rows(intRow)(0).ToString.Substring(0, 4) = "TEXT") Then
                strValue += "'" & value.ToString & "'"
            ElseIf table.Rows(intRow)(0) = "TIME" Then
                If Not IsDBNull(value) Then
                    strValue += "'" & value.ToString & "'"
                Else
                    strValue += "NULL"
                End If
            ElseIf (table.Rows(intRow)(0).Length >= 8 AndAlso _
        table.Rows(intRow)(0).ToString.Substring(0, 8) = "TINYTEXT") Then
                strValue += "'" & value.ToString & "'"

            ElseIf (table.Rows(intRow)(0).Length >= 7 AndAlso _
         table.Rows(intRow)(0).ToString.Substring(0, 7) = "TINYINT") Then
                If IsNumeric(value) Then
                    strValue += CDbl(value).ToString
                End If
            ElseIf IsNumeric(value) = True Then
                strValue += value.ToString
            Else
                strValue += "'" & value.ToString & "'"
            End If
            Return strValue
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim ij As New System.Globalization.DateTimeFormatInfo
        Dim jk As New System.Globalization.NumberFormatInfo
        Dim ObjCD1 As New Core_Module.ClsCekData

        'My.Application.MinimumSplashScreenDisplayTime = 6000

        'cek ui culture dan keyboard input 
        My.Application.ChangeCulture("en-US")
        My.Application.ChangeUICulture("en-US")

        My.Application.ChangeCulture("en-US")
        My.Application.ChangeUICulture("en-US")
        ij.LongDatePattern = "dddd, MMM dd yyyy"
        ij.ShortDatePattern = "MMM dd yyyy"
        jk.CurrencyDecimalSeparator = "."
        jk.CurrencyGroupSeparator = ","
        My.Application.Culture.NumberFormat = jk
        My.Application.Culture.DateTimeFormat = ij
        ObjCD1.UpdateThreading()
        Me.BackUpDataBase()
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.BackUpDataBase1()
            'Windows.Forms.Cursor.Current = Cursors.Arrow
            SystemSounds.Beep.Play()
            If Me.FinishBackUp = True Then
                'Me.ProgressBar1.Style = ProgressBarStyle.Continuous
                Me.lblProgres.Text = "Back Up Finish"


                If Me.vShutdown = True Then
                    System.Diagnostics.Process.Start("ShutDown", "-f -s")
                    Me.Close()
                    Exit Sub
                Else
                    Me.btnClose.Visible = True
                    Me.btnClose.Focus()
                    Me.Timer1.Interval = 1000
                    Me.Timer1.Enabled = True
                    Me.ijk = 0
                End If
            Else
                If Me.BackUpEx = EnumTroubleBackUp.InvalidFolder Then
                    Me.lblProgres.Text = "Back up interupted. Invalid back up location, please make sure you have access to back up location."
                Else
                    Me.lblProgres.Text = "Back up interupted. Please contact your vendor if you see this message."
                End If

                Me.btnClose.Visible = True
                Me.btnClose.Focus()
                Me.Timer1.Interval = 1000
                Me.Timer1.Enabled = True
                Me.ijk = 0
            End If
            'Windows.Forms.Cursor.Current = Cursors.Arrow
            'Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Try
            Me.ijk += 1

            If Me.ijk >= 30 Then
                Me.Close()
            Else
                Me.btnClose.Text = "Close (" + (30 - ijk).ToString + ")"
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class