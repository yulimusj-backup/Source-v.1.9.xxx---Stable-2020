Option Explicit On
Option Strict Off

Imports System.IO
Imports MySql.Data.MySqlClient
Imports MySql.Data
Imports System.Media
Imports ICSharpCode.SharpZipLib.Core
Imports ICSharpCode.SharpZipLib.Zip

Public Class FrmBackUpData2
    Dim ijk As Integer = 0
    Dim ObjEncrypt As New ClsEncryptFile
    Dim ObjZip As New Core_Module.ClsCompressFile

    Public vConStr As String = ""
    Public vFileName As String = ""
    Public vCompress As Boolean = False
    Public vExcryptPass As String = Core_Module.GetEncryptPass
    Public vUpdateStructure As Boolean = True
    Public vUpdateData As Boolean = True
    Public vRestoreMatDash As Boolean = True

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim ij As New System.Globalization.DateTimeFormatInfo
        Dim jk As New System.Globalization.NumberFormatInfo

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
        Me.RestoreDatabase(Me.vConStr, Me.vFileName, Me.vRestoreMatDash)
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Windows.Forms.Cursor.Current = Cursors.Arrow
            SystemSounds.Beep.Play()
            Me.btnClose.Visible = True
            Me.btnClose.Focus()
            Me.Timer1.Interval = 1000
            If Me.TxtErrorLog.Text = "" Then
                Me.Timer1.Enabled = True
            End If
            Me.ijk = 0

            Windows.Forms.Cursor.Current = Cursors.Arrow
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
            Me.TextBox1.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmBackUpData2_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Me.BringToFront()
    End Sub

    Private Sub FrmBackUpData2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.lblProgres.Text = ""
        Me.BringToFront()
        Control.CheckForIllegalCrossThreadCalls = False
        Me.BackgroundWorker1.RunWorkerAsync()

        ''tanpa multi thread
        'Dim ij As New System.Globalization.DateTimeFormatInfo
        'Dim jk As New System.Globalization.NumberFormatInfo
        'My.Application.ChangeCulture("en-US")
        'My.Application.ChangeUICulture("en-US")
        'My.Application.ChangeCulture("en-US")
        'My.Application.ChangeUICulture("en-US")
        'ij.LongDatePattern = "dddd, MMM dd yyyy"
        'ij.ShortDatePattern = "MMM dd yyyy"
        'jk.CurrencyDecimalSeparator = "."
        'jk.CurrencyGroupSeparator = ","
        'My.Application.Culture.NumberFormat = jk
        'My.Application.Culture.DateTimeFormat = ij
        'Me.RestoreDatabase(Me.vConStr, Me.vFileName)
        'Windows.Forms.Cursor.Current = Cursors.Arrow
        'SystemSounds.Beep.Play()
        'Me.btnClose.Visible = True
        'Me.btnClose.Focus()
        'Me.Timer1.Interval = 1000
        'Me.Timer1.Enabled = True
        'Me.ijk = 0
        'Windows.Forms.Cursor.Current = Cursors.Arrow

        Me.TextBox1.Focus()
    End Sub

    'Private Sub ExtractZipFile(ByVal archiveFilenameIn As String, ByVal password As String, ByVal outFolder As String)
    '    Dim zf As ZipFile = Nothing

    '    Try
    '        Dim fs As FileStream = File.OpenRead(archiveFilenameIn)
    '        zf = New ZipFile(fs)
    '        'If password = "" Then
    '        '    password = "password"
    '        'End If

    '        For Each Zipentry As ZipEntry In zf
    '            If Zipentry.IsFile = True Then
    '                Dim entryFileName As String = Zipentry.Name
    '                'to remove the folder from the entry:- entryFileName = Path.GetFileName(entryFileName);
    '                'Optionally match entrynames against a selection list here to skip as desired.
    '                'The unpacked length is available in the zipEntry.Size property.
    '                Dim buffer As Byte() = New Byte(4096)
    '                Dim zipStream As Stream = zf.GetInputStream(Zipentry)

    '                'byte[] buffer = new byte[4096];     // 4K is optimum

    '                'Manipulate the output filename here as desired.
    '        String fullZipToPath = Path.Combine(outFolder, entryFileName);
    '        string directoryName = Path.GetDirectoryName(fullZipToPath);
    '                If (directoryName.Length > 0) Then
    '            Directory.CreateDirectory(directoryName);

    '                    'Unzip file in buffered chunks. This is just as fast as unpacking to a buffer the full size
    '                    'of the file, but does not waste memory.
    '                    'The "using" will close the stream even if an exception occurs.
    '        using (FileStream streamWriter = File.Create(fullZipToPath)) {
    '            StreamUtils.Copy(zipStream, streamWriter, buffer);

    '            End If
    '        Next
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        If IsNothing(zf) = False Then
    '            zf.IsStreamOwner = True
    '            zf.Close()
    '        End If
    '    End Try
    'End Sub

    Private Function GetTargetDatabase(ByVal vConstring As String) As String
        '"server=server;user id=root;Password=#SMeSS_p4ss*897%#;persist security info=True;database=core_db;port=3307;default command timeout=19000000"
        Dim Retval As String = ""
        Dim StartIndex As Integer = 0
        Dim EndIndex As Integer = 0
        Try
            StartIndex = vConstring.IndexOf("database=") + 9
            EndIndex = vConstring.IndexOf(";", StartIndex)
            Retval = vConStr.Substring(StartIndex, EndIndex - StartIndex)
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function RestoreDatabase(ByVal vConstring As String, ByVal vFileName As String, Optional ByVal vRestoreWithMatDash As Boolean = True) As Boolean
        Dim Attempt As Integer = 0
        Dim ErrorString As New List(Of String)
        Dim DummyString As String = ""
        Dim conStr As String = vConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim SetOfString As String = ""
        Dim ReadingSetOfString As Boolean = False
        Dim MaxLine As Double = 0
        Dim CurrentLine As Double = 0
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ListQueryPending As New List(Of String)
        Dim QueryPending As String = ""
        Dim QueryPendingExe As Double = 0
        Dim ij As Double = 0
        Dim bytKey As Byte()
        Dim bytIV As Byte()
        Dim FileNameBaru As String = ""
        Dim FileNameHasilExtract As String = ""
        Dim TargetDBName As String = ""
        Dim SourceDBName As String = ""

        Try
            Me.TxtErrorLog.Text = ""

            TargetDBName = Me.GetTargetDatabase(vConstring)

            If Me.vCompress = True Then
                'Send the password to the CreateKey function.
                bytKey = Me.ObjEncrypt.CreateKey(Me.vExcryptPass)
                'Send the password to the CreateIV function.
                bytIV = Me.ObjEncrypt.CreateIV(Me.vExcryptPass)

                '1. decrypt file
                FileNameBaru = vFileName.Substring(0, Me.vFileName.Length - 6) + ".zip"
                Me.lblProgres.Text = "Decrypting file, please wait..."
                If My.Computer.FileSystem.FileExists(FileNameBaru) = True Then
                    My.Computer.FileSystem.DeleteFile(FileNameBaru, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                End If
                Me.ObjEncrypt.EncryptOrDecryptFile(vFileName, FileNameBaru, bytKey, _
                bytIV, ClsEncryptFile.CryptoAction.ActionDecrypt, Me.ProgressBar1)
                'set hidden
                File.SetAttributes(FileNameBaru, FileAttributes.Hidden)

                '2. decompress file
                Me.lblProgres.Text = "Decompressing file, please wait..."
                FileNameHasilExtract = vFileName.Substring(0, Me.vFileName.Length - 6) + ".sql"
                If My.Computer.FileSystem.FileExists(FileNameHasilExtract) = True Then
                    My.Computer.FileSystem.DeleteFile(FileNameHasilExtract, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                End If
                Me.ObjZip.UnZipFile(FileNameBaru, FileNameHasilExtract)
                vFileName = FileNameHasilExtract
                'set hidden
                File.SetAttributes(vFileName, FileAttributes.Hidden)
            End If

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.lblProgres.Text = "Parsing file, please wait..."
            Using parser As New FileIO.TextFieldParser(vFileName)
                While Not parser.EndOfData
                    MaxLine += 1
                    parser.ReadLine()
                End While
            End Using
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = MaxLine

            'jalankan struktur
            'Conn.Open() -> dipindahkan setelah tahu nama db

            Using parser As New FileIO.TextFieldParser(vFileName, System.Text.Encoding.UTF8, True)
                While Not parser.EndOfData
                    Me.ProgressBar1.Value += 1
                    If Me.ProgressBar1.Value < 250 Then
                        Me.lblProgres.Text = "Restoring database " + ObjCD.DigitGrouping(Math.Round((Me.ProgressBar1.Value / MaxLine) * 100, 2, MidpointRounding.AwayFromZero), 2) + _
                        "% - Data " + ObjCD.DigitGrouping(Me.ProgressBar1.Value) + " of " + ObjCD.DigitGrouping(Me.ProgressBar1.Maximum)
                    Else
                        If Math.IEEERemainder(Me.ProgressBar1.Value, 500) = 0 Then
                            Me.lblProgres.Text = "Restoring database " + ObjCD.DigitGrouping(Math.Round((Me.ProgressBar1.Value / MaxLine) * 100, 2, MidpointRounding.AwayFromZero), 2) + _
                            "% - Data " + ObjCD.DigitGrouping(Me.ProgressBar1.Value) + " of " + ObjCD.DigitGrouping(Me.ProgressBar1.Maximum)
                        ElseIf Me.ProgressBar1.Value = Me.ProgressBar1.Maximum Then
                            Me.lblProgres.Text = "Restoring database " + ObjCD.DigitGrouping(Math.Round((Me.ProgressBar1.Value / MaxLine) * 100, 2, MidpointRounding.AwayFromZero), 2) + _
                            "% - Data " + ObjCD.DigitGrouping(Me.ProgressBar1.Value) + " of " + ObjCD.DigitGrouping(Me.ProgressBar1.Maximum)
                        End If
                    End If

                    ' Read in the fields for the current line
                    DummyString = parser.ReadLine
                    ' Add code here to use data in fields variable.
                    DummyString = DummyString.Trim

                    If ReadingSetOfString = False Then
                        If DummyString = "SET FOREIGN_KEY_CHECKS=0;" Then
                            'MsgBox("Execute: " + DummyString)
                            Me.ExecuteQuery(Conn, script, DummyString)
                        ElseIf DummyString.StartsWith("# Database :") = True Then
                            SourceDBName = DummyString.Substring(13)
                            If SourceDBName = TargetDBName Then
                                Conn.Open()
                            Else
                                Conn.ConnectionString = Conn.ConnectionString.Replace(TargetDBName, SourceDBName)
                                Conn.Open()
                            End If
                        ElseIf DummyString.StartsWith("DROP DATABASE IF EXISTS") = True Then
                            'MsgBox("Execute: " + DummyString)
                            If Me.vUpdateStructure = True Then
                                Me.ExecuteQuery(Conn, script, DummyString.Replace(SourceDBName, TargetDBName))
                            End If
                        ElseIf DummyString.StartsWith("USE `") = True Then
                            'MsgBox("Execute: " + DummyString)
                            Me.ExecuteQuery(Conn, script, DummyString.Replace(SourceDBName, TargetDBName))
                        ElseIf DummyString = "#" = True Then
                            'diabaikan
                        ElseIf DummyString.StartsWith("CREATE DATABASE ") = True Then
                            SetOfString = ""
                            ReadingSetOfString = True
                            SetOfString = DummyString
                        ElseIf DummyString.StartsWith("# Structure for ") = True Then
                            SetOfString = ""
                            ReadingSetOfString = True
                        ElseIf DummyString.StartsWith("CREATE TABLE ") = True Then
                            SetOfString = DummyString
                            ReadingSetOfString = True
                        ElseIf DummyString.StartsWith("CREATE ALGORITHM") = True Then
                            If Me.vUpdateStructure = True Then
                                If Me.ExecuteQuery(Conn, script, DummyString, False) = False Then
                                    ListQueryPending.Add(DummyString)
                                End If
                            End If
                        ElseIf DummyString.StartsWith("# definition for") = True Then
                            SetOfString = ""
                            ReadingSetOfString = True
                        ElseIf DummyString.StartsWith("# data for the") Then
                            If vUpdateStructure = True Then
                                ij = 0
                                While ListQueryPending.Count > 0
                                    QueryPendingExe += 1
                                    If ij < ListQueryPending.Count Then
                                        QueryPending = ListQueryPending(ij)
                                    Else
                                        ij = 0
                                        QueryPendingExe = 0
                                        QueryPending = ListQueryPending(ij)
                                    End If
                                    If Me.ExecuteQuery(Conn, script, QueryPending) = True Then
                                        ListQueryPending.RemoveAt(ij)
                                    Else
                                        ij += 1
                                    End If
                                    If QueryPendingExe > 500 Then
                                        MsgBox("Recursive table structure", MsgBoxStyle.Information, "Information")
                                        Me.Close()
                                        Exit Function
                                    End If
                                End While
                            End If

                            SetOfString = ""
                            ReadingSetOfString = True
                        End If
                    Else
                        If DummyString.EndsWith(";") Then
                            SetOfString = SetOfString + vbCrLf + DummyString
                            ReadingSetOfString = False
                            'MsgBox("Execute: " + SetOfString)
                            If SetOfString.StartsWith("CREATE TABLE ") = True Or SetOfString.StartsWith("CREATE ALGORITHM") = True Then
                                If SetOfString.StartsWith("CREATE TABLE ") = True Then
                                    If vUpdateStructure = True Then
                                        If Me.ExecuteQuery(Conn, script, SetOfString) = False Then
                                            ListQueryPending.Add(SetOfString)
                                        End If
                                    End If
                                ElseIf SetOfString.StartsWith("CREATE ALGORITHM") = True Then
                                    If vUpdateStructure = True Then
                                        If Me.ExecuteQuery(Conn, script, SetOfString, False) = False Then
                                            ListQueryPending.Add(SetOfString)
                                        End If
                                    End If
                                End If
                            Else
                                If SetOfString.StartsWith("CREATE DATABASE") = True Then
                                    If vUpdateStructure = True Then
                                        Me.ExecuteQuery(Conn, script, SetOfString.Replace(SourceDBName, TargetDBName))
                                    End If
                                    Conn.Close()
                                    Conn.ConnectionString = Conn.ConnectionString.Replace(SourceDBName, TargetDBName)
                                    Conn.Open()
                                    Me.ExecuteQuery(Conn, script, "SET FOREIGN_KEY_CHECKS=0;")
                                ElseIf SetOfString.Contains("DROP TABLE IF EXISTS") = True Then
                                    If vUpdateStructure = True Then
                                        Me.ExecuteQuery(Conn, script, SetOfString)
                                    End If
                                ElseIf SetOfString.Contains("DROP VIEW IF EXISTS") = True Then
                                    If vUpdateStructure = True Then
                                        Me.ExecuteQuery(Conn, script, SetOfString)
                                    End If
                                ElseIf SetOfString.Contains("INSERT INTO") = True Then
                                    If vUpdateData = True Then
                                        If vRestoreWithMatDash = True Then
                                            Me.ExecuteQuery(Conn, script, SetOfString)
                                        Else
                                            If SetOfString.Contains("INSERT INTO `tbltrans_material_planning") = False Then
                                                Me.ExecuteQuery(Conn, script, SetOfString)
                                                'Else
                                                'MsgBox("")
                                            End If
                                        End If
                                    Else
                                        Exit While
                                    End If
                                Else
                                    Me.ExecuteQuery(Conn, script, SetOfString)
                                End If
                            End If
                        ElseIf DummyString = "#" = True Then
                            'diabaikan
                        ElseIf DummyString.Trim.StartsWith("CHARACTER SET") = True Then
                            'MsgBox("ting")
                            If DummyString.Trim.Remove(0, 14) <> "latin1" Then
                                Conn.Close()
                                Conn.ConnectionString = Conn.ConnectionString + ";character set=" + DummyString.Trim.Remove(0, 14)
                                Conn.Open()
                            End If
                            SetOfString = SetOfString + vbCrLf + DummyString
                        Else
                            SetOfString = SetOfString + vbCrLf + DummyString
                        End If
                    End If
                End While
            End Using
            Conn.Close()

            If Me.vCompress = True Then
                '3. delete file hasil extract
                If My.Computer.FileSystem.FileExists(FileNameHasilExtract) = True Then
                    My.Computer.FileSystem.DeleteFile(FileNameHasilExtract, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                End If
                'delete hasil decrypt
                If My.Computer.FileSystem.FileExists(FileNameBaru) = True Then
                    My.Computer.FileSystem.DeleteFile(FileNameBaru, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
            ObjCD = Nothing
        End Try
    End Function

    Private Function ExecuteQuery(ByVal vConn As MySql.Data.MySqlClient.MySqlConnection, ByVal vScript As MySql.Data.MySqlClient.MySqlScript, _
    ByVal vsql As String, Optional ByVal vWriteLog As Boolean = True) As Boolean
        Dim DummyIndex As Double = 0
        Dim Dummy As String = ""

        Try
            vsql = vsql.Trim
            'open connection
            'vConn.Open()
            Dummy = vsql
            If Dummy.StartsWith("CREATE TABLE ") = True Then
                DummyIndex = Dummy.IndexOf("`")
                DummyIndex = Dummy.IndexOf("`", CInt(DummyIndex + 1))
                If DummyIndex <> -1 Then
                    Dummy = Dummy.Substring(0, DummyIndex)
                    Dummy = Dummy.Replace("`", "")
                End If
                Me.TextBox1.Text = Dummy
            ElseIf Dummy.StartsWith("CREATE ALGORITHM") Then
                DummyIndex = Dummy.IndexOf("VIEW")
                Dummy = Dummy.Substring(DummyIndex)
                DummyIndex = Dummy.IndexOf("`")
                DummyIndex = Dummy.IndexOf("`", CInt(DummyIndex + 1))
                Dummy = Dummy.Substring(0, DummyIndex)
                Dummy = Dummy.Replace("`", "")
                Me.TextBox1.Text = "Create " + Dummy
            ElseIf Dummy.StartsWith("INSERT INTO") = True Then
                DummyIndex = Dummy.IndexOf("`")
                DummyIndex = Dummy.IndexOf("`", CInt(DummyIndex + 1))
                If DummyIndex <> -1 Then
                    Dummy = Dummy.Substring(0, DummyIndex)
                    Dummy = Dummy.Replace("`", "")
                End If
                Me.TextBox1.Text = Dummy
            Else
                Me.TextBox1.Text = Dummy
            End If

            vScript.Query = vsql
            'execute script
            vScript.Execute()
            'vConn.Close()
            Return True
        Catch ex As Exception
            If vWriteLog = True Then
                If Me.TxtErrorLog.Text.Contains(vScript.Query.Substring(0, 50)) = False Then
                    If Me.TxtErrorLog.Text = "" Then
                        Me.TxtErrorLog.Text = "....." + ex.Message + vbCrLf + vScript.Query.Substring(0, 50) + "..." + vbCrLf
                    Else
                        Me.TxtErrorLog.Text = Me.TxtErrorLog.Text + vbCrLf + "....." + ex.Message + vbCrLf + vScript.Query.Substring(0, 50) + "..." + vbCrLf
                    End If
                End If
            End If
            ' MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function
End Class