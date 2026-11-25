Option Explicit On
Option Strict Off

Imports System.Net
Imports System.IO
Imports System.Environment

Public Class FrmDownload
    Dim whereToSave As String
    Dim Dta As New DseAutoUpdateTableAdapters.tbl_setting_updateTableAdapter
    Dim ObjUpdate As New ClsAutoUpdate
    'Dim ObjURL As New Clsweb
    'Dim ObjDBVersion As New ClsDBVersion
    Dim ObjAutoUpdate As New ClsAutoUpdate
    Dim PanggilFungsiBuat As EnumFungsi
    Dim Fase As EnumProgress = EnumProgress.Start
    Dim Identifier As String = ""

    Private Enum EnumProgress As Integer
        Start
        CheckingTargetDirectory
        CopyFile
        DecryptFolder
        Update
        DeleteBufferFile
        StartSMESS
    End Enum

    Public Enum EnumFungsi As Integer
        _download
        _UpdateViaServer
    End Enum

    Public Sub DecryptFile(ByVal vFileName As String)
        'decrypt: decrypt -> decompress
        Dim ObjZIP As New ClsCompressFile
        Dim ObjEncrypt As New ClsEncryptFile
        Dim FileName As String = ""
        Dim FileNameHasilDeCompress As String = ""
        Dim FileNameHasilDecrypt As String = ""
        Dim bytKey As Byte()
        Dim bytIV As Byte()
        Dim vEncryptPass As String = GetEncryptPass_Installer()
        Try
            'pilih file *.smess
            FileName = vFileName
            'decrypt 
            bytKey = ObjEncrypt.CreateKey(vEncryptPass)
            'Send the password to the CreateIV function.
            bytIV = ObjEncrypt.CreateIV(vEncryptPass)
            FileNameHasilDecrypt = FileName.Substring(0, FileName.Length - 6) + ".zip"
            ObjEncrypt.EncryptOrDecryptFile(FileName, FileNameHasilDecrypt, bytKey, bytIV, ClsEncryptFile.CryptoAction.ActionDecrypt, Me.ProgressBar1)
            Me.ProgressBar1.Value = 0
            Me.LblSize.Text = "Decrypt Smess Version"
            'de-compress
            FileNameHasilDeCompress = FileName.Substring(0, FileName.Length - 6) + ".txt"
            ObjZIP.UnZipFile(FileNameHasilDecrypt, FileNameHasilDeCompress)
            'hapus file hasil compress
            If My.Computer.FileSystem.FileExists(FileNameHasilDecrypt) = True Then
                My.Computer.FileSystem.DeleteFile(FileNameHasilDecrypt, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
            'Set Progress Bar default
            Me.ProgressBar1.Value = 0
            Me.LblSize.Text = ""
            'MsgBox("Finish", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjZIP = Nothing
            ObjEncrypt = Nothing
        End Try
    End Sub

    Private Sub DecryptFolder(ByVal vFileName As String)
        'decrypt: decrypt -> decompress
        Dim ObjZIP As New ClsCompressFile
        Dim ObjEncrypt As New ClsEncryptFile
        Dim FileName As String = ""
        Dim FolderNameHasilDeCompress As String = ""
        Dim FileNameHasilDecrypt As String = ""
        Dim bytKey As Byte()
        Dim bytIV As Byte()
        Dim vEncryptPass As String = GetEncryptPass_Installer()
        Try
            FileName = vFileName
            'decrypt 
            bytKey = ObjEncrypt.CreateKey(vEncryptPass)
            'Send the password to the CreateIV function.
            bytIV = ObjEncrypt.CreateIV(vEncryptPass)
            FileNameHasilDecrypt = FileName.Substring(0, FileName.Length - 6) + ".zip"
            ObjEncrypt.EncryptOrDecryptFile(FileName, FileNameHasilDecrypt, bytKey, bytIV, ClsEncryptFile.CryptoAction.ActionDecrypt, Me.ProgressBar1)
            Me.ProgressBar1.Value = 0
            'Me.LblSize.Text = "Decrypt Smess Update"
            'de-compress
            FolderNameHasilDeCompress = FileName.Substring(0, FileName.LastIndexOf("\"))
            ObjZIP.UnZipFolder(FileNameHasilDecrypt, GetEncryptPass_Installer2, FolderNameHasilDeCompress)
            'hapus file hasil decrypt
            If My.Computer.FileSystem.FileExists(FileNameHasilDecrypt) = True Then
                My.Computer.FileSystem.DeleteFile(FileNameHasilDecrypt, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
            'Set Progress Bar Value
            Me.ProgressBar1.Value = 0
            Me.LblSize.Text = ""
            'MsgBox("Finish", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjZIP = Nothing
            ObjEncrypt = Nothing
        End Try
    End Sub

    Private Function ConnectionInternet(ByVal vEnumFungsi As EnumFungsi) As Boolean
        Dim ServerPath As String = ""
        Try
            If vEnumFungsi = EnumFungsi._download Then
                If My.Computer.Network.Ping("www.google.com") = False Then
                    MsgBox("No internet connection.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
                'If Me.txtFileName.Text <> "" Then
                '    MsgBox("Please insert valid URL for download", MsgBoxStyle.Information, "Information")
                '    Return False
                'End If
                If Me.ObjAutoUpdate.GetServerPath() = "" Then
                    MsgBox("Please setting server path.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
                If System.IO.Directory.Exists(ServerPath) = False Then
                    MsgBox("Server path doesn't exists.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If vEnumFungsi = EnumFungsi._UpdateViaServer Then
                ServerPath = Me.ObjAutoUpdate.GetServerPath()
                If ServerPath = "" Then
                    MsgBox("Please setting server path.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
                If System.IO.Directory.Exists(ServerPath) = False Then
                    MsgBox("Server path doesn't exists.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
                If System.IO.File.Exists(ServerPath & "\" & "SMESS_Version.smess") = False Then
                    MsgBox("File SMESS_Version.smess doesn't exists at server path.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
                If System.IO.File.Exists(ServerPath & "\" & "SMESS_Update.smess") = False Then
                    MsgBox("File SMESS_Update.smess doesn't exists at server path.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Me.Close()
        End Try
    End Function

    'Public Sub RunProsesDownload(ByVal vEnumFungsi As EnumFungsi)
    '    Dim proc As New System.Diagnostics.Process()
    '    Dim ServerPath As String = Me.ObjAutoUpdate.GetServerPath()
    '    Dim fileExists As Boolean

    '    Try
    '        For Each p As Process In Process.GetProcesses
    '            If p.ProcessName = Me.Dta.ScalarQuery_GetProductName.ToString Then
    '                p.Kill()

    '                'klo version lebih besar maka download file smess + update
    '                'set tampilan
    '                Threading.Thread.Sleep(2000)
    '                Me.ProgressBar1.Value = 0
    '                Me.LblTypeFile.Text = ""
    '                Me.LblSize.Text = "Checking for download folder"
    '                Me.CekDirectoryDownload(My.Application.Info.DirectoryPath & "\Download")

    '                'copy File via server
    '                'set tampilan
    '                Threading.Thread.Sleep(2000)
    '                Me.ProgressBar1.Value = 0
    '                Me.LblTypeFile.Text = ""
    '                Me.LblSize.Text = "Finding source file from LAN folder"
    '                System.IO.File.Copy(ServerPath & "\" & "SMESS_Update.smess" _
    '                                   , My.Application.Info.DirectoryPath & "\Download" & "\" & "SMESS_Update.smess", True)

    '                'decrypt folder
    '                Threading.Thread.Sleep(2000)
    '                Me.DecryptFolder(My.Application.Info.DirectoryPath & "\Download\" & "SMESS_Update.smess")

    '                'update -> copy timpa
    '                Threading.Thread.Sleep(2000)
    '                Me.UpdateData(My.Application.Info.DirectoryPath & "\Download", My.Application.Info.DirectoryPath, "Software_SMESS_Update")

    '                'update version
    '                Threading.Thread.Sleep(2000)
    '                fileExists = My.Computer.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\" & "SMESS_Version_Current.smess")
    '                If fileExists = True Then
    '                    'hapus versi lama
    '                    My.Computer.FileSystem.DeleteFile(My.Application.Info.DirectoryPath & "\" & "SMESS_Version_Current.smess", _
    '                    FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
    '                End If
    '                fileExists = My.Computer.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\" & "SMESS_Version.smess")
    '                If fileExists = True Then
    '                    'hapus versi lama
    '                    My.Computer.FileSystem.DeleteFile(My.Application.Info.DirectoryPath & "\" & "SMESS_Version.smess", _
    '                    FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
    '                End If
    '                fileExists = My.Computer.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\" & "SMESS_Version.txt")
    '                If fileExists = True Then
    '                    'hapus versi lama
    '                    My.Computer.FileSystem.DeleteFile(My.Application.Info.DirectoryPath & "\" & "SMESS_Version.txt", _
    '                    FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
    '                End If
    '                System.IO.File.Copy(ServerPath & "\" & "SMESS_Version.smess" _
    '                                   , My.Application.Info.DirectoryPath & "\" & "SMESS_Version_Current.smess", True)

    '                'Hapus directory download
    '                Threading.Thread.Sleep(2000)
    '                System.IO.Directory.Delete(My.Application.Info.DirectoryPath & "\Download", True)
    '                MsgBox("SMESS update finished.", MsgBoxStyle.Information, "Information")

    '                'nyalain lagi smess-nya
    '                Threading.Thread.Sleep(360)
    '                proc = Process.Start(My.Application.Info.DirectoryPath & "\SMESS.exe")
    '                Me.Close()
    '                Exit For
    '            End If
    '        Next
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Me.Close()
    '    Finally
    '        Me.BackgroundWorker2.CancelAsync()
    '    End Try
    'End Sub

    Private Sub SetIdentifier()
        Dim DTA As New DseTransaksiCustom_IdentifierTableAdapters.tbltrans_custom_identifierTableAdapter
        Dim TBL As New DseTransaksiCustom_Identifier.tbltrans_custom_identifierDataTable
        Dim row As DseTransaksiCustom_Identifier.tbltrans_custom_identifierRow

        Try
            DTA.Fill(TBL)
            row = TBL.Rows(0)
            Me.Identifier = row.identifier
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Sub RunProsesDownload()
        Dim ServerPath As String = Me.ObjAutoUpdate.GetServerPath()
        Dim fileExists As Boolean

        Try
            Me.SetIdentifier()
            fileExists = My.Computer.FileSystem.FileExists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current" + "_" + Me.Identifier + ".smess")
            If fileExists = True Then
                System.IO.File.SetAttributes(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current" + "_" + Me.Identifier + ".smess", FileAttributes.Normal)
            End If
            Me.Label3.Text = ""
            Me.Label1.Text = ""
            Me.Label4.Text = ""
            Me.LblSize.Text = ""
            Me.LblTypeFile.Text = ""
            Me.LinkLabel1.Text = ""

            For Each p As Process In Process.GetProcesses
                If p.ProcessName = Me.Dta.ScalarQuery_GetProductName.ToString Then
                    p.Kill()
                    Exit For
                End If
            Next

            Me.ProgressBar1.Value = 0
            Me.LblTypeFile.Text = ""
            Me.LblSize.Text = "Checking for download folder"
            Me.CekDirectoryDownload(My.Application.Info.DirectoryPath & "\Download")

            Me.ProgressBar1.Value = 0
            Me.LblTypeFile.Text = ""
            Me.LblSize.Text = "Finding source file from LAN folder"
            System.IO.File.Copy(ServerPath & "\" & "SMESS_Update.smess" _
                               , My.Application.Info.DirectoryPath & "\Download" & "\" & "SMESS_Update.smess", True)

            Me.LblSize.Text = "Decrypting update files"
            Me.DecryptFolder(My.Application.Info.DirectoryPath & "\Download\" & "SMESS_Update.smess")

            Me.LblSize.Text = "Updating data"
            Me.UpdateData(My.Application.Info.DirectoryPath & "\Download", My.Application.Info.DirectoryPath, "SMESS Software Update")

            Me.LblSize.Text = "Updating version"
            fileExists = My.Computer.FileSystem.FileExists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current" + "_" + Me.Identifier + ".smess")
            If fileExists = True Then
                'hapus versi lama
                My.Computer.FileSystem.DeleteFile(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current" + "_" + Me.Identifier + ".smess", _
                FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
            fileExists = My.Computer.FileSystem.FileExists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version.smess")
            If fileExists = True Then
                'hapus versi lama
                My.Computer.FileSystem.DeleteFile(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version.smess", _
                FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
            fileExists = My.Computer.FileSystem.FileExists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version.txt")
            If fileExists = True Then
                'hapus versi lama
                My.Computer.FileSystem.DeleteFile(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version.txt", _
                FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
            fileExists = My.Computer.FileSystem.FileExists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current" + "_" + Me.Identifier + ".txt")
            If fileExists = True Then
                'hapus versi lama
                My.Computer.FileSystem.DeleteFile(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current" + "_" + Me.Identifier + ".txt", _
                FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
            System.IO.File.Copy(ServerPath & "\" & "SMESS_Version.smess" _
                               , Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current" + "_" + Me.Identifier + ".smess", True)
            Me.LblSize.Text = "Deleting buffer file"
            System.IO.Directory.Delete(My.Application.Info.DirectoryPath & "\Download", True)
            System.IO.File.SetAttributes(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current" + "_" + Me.Identifier + ".smess", FileAttributes.Hidden)

            Me.LblSize.Text = "Starting SMESS"

            Dim proc As New System.Diagnostics.Process()
            proc = Process.Start(My.Application.Info.DirectoryPath & "\SMESS.exe")
            Me.Close()
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Dim proc As New System.Diagnostics.Process()
            proc = Process.Start(My.Application.Info.DirectoryPath & "\SMESS.exe")
            Me.Close()
        Finally
            Me.BackgroundWorker2.CancelAsync()
        End Try
    End Sub

    Private Sub btnDownload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDownload.Click
        Try
            Me.PanggilFungsiBuat = EnumFungsi._download
            'If Me.ConnectionInternet(Me.PanggilFungsiBuat) = True Then
            Me.BackgroundWorker1.RunWorkerAsync()
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Me.Close()
        End Try
    End Sub

    Private Sub BtnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnUpdate.Click
        Try
            Me.PanggilFungsiBuat = EnumFungsi._UpdateViaServer
            'If Me.ConnectionInternet(Me.PanggilFungsiBuat) = True Then
            Me.BackgroundWorker1.RunWorkerAsync()
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Me.Close()
        End Try
    End Sub

    Private Function ApakahBolehUpdateData(ByVal vProcName As String) As Boolean
        Try
            If vProcName = "Download Module.exe" Then
                Return False
            End If
            If vProcName = "Client Update Module.exe" Then
                Return False
            End If
            If vProcName = "Setting Module.dll" Then
                Return False
            End If
            'If vProcName = "CustomSMESS.dll" Then
            '    Return False
            'End If
            If vProcName = "ICSharpCode.SharpZipLib.dll" Then
                Return False
            End If
            If vProcName = "MySql.Data.dll" Then
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Me.Close()
        End Try
    End Function

    Private Sub UpdateData(ByVal vDownloadPath As String, ByVal vProgramPath As String, ByVal vKeterangan As String)
        Dim ListFileDownload As String() = System.IO.Directory.GetFiles(vDownloadPath)
        'Dim UsedPath As String = "C:\Program Files\Artha Jina Mandala Nusantara\SMESS\Download"
        Dim Extension As String = ""
        Dim NamaFile1 As String()

        Try
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = ListFileDownload.Length
            Me.LblSize.Text = vKeterangan
            'vProgramPath = UsedPath
            'ambil semua list file yang ada di file instalasi
            For i As Integer = 1 To ListFileDownload.Length
                NamaFile1 = ListFileDownload(i - 1).Split("\")

                Dim NamaFile As String() = ListFileDownload(i - 1).Split("\")
                Me.LblSize.Text = "Updating data " + i.ToString + " of " + ListFileDownload.Length.ToString + " : " + NamaFile(NamaFile.Length - 1).ToString

                Extension = System.IO.Path.GetExtension(vDownloadPath & "\" & NamaFile(NamaFile.Length - 1))
                If Extension = ".dll" Or Extension = ".exe" Or Extension = ".config" Then
                    If Me.ApakahBolehUpdateData(NamaFile(NamaFile.Length - 1).ToString) = True Then
                        System.IO.File.Copy(vDownloadPath & "\" & NamaFile(NamaFile.Length - 1) _
                                           , vProgramPath & "\" & NamaFile(NamaFile.Length - 1), True)
                    End If
                End If
                'Me.ProgressBar1.Value += 1
            Next i
            'set tampilan
            Me.ProgressBar1.Value = 0
            Me.LblTypeFile.Text = ""
            'Me.LblSize.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Me.Close()
        End Try
    End Sub

    Private Sub CekDirectoryDownload(ByVal vPath As String)
        Try
            If System.IO.Directory.Exists(vPath) = False Then
                System.IO.Directory.CreateDirectory(vPath)
            Else
                System.IO.Directory.Delete(vPath, True)
                System.IO.Directory.CreateDirectory(vPath)
            End If
            If System.IO.Directory.Exists(vPath) = False Then
                MsgBox("Unable to write on program files folder.", MsgBoxStyle.Information, "Information")
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Me.Close()
        End Try
    End Sub

    Private Function ApakahVersionBerbeda(ByVal vLastVersion As String, ByVal vNewVersion As String) As Boolean
        Dim ParamLast As String() = vLastVersion.Split(".")
        Dim ParamNew As String() = vNewVersion.Split(".")
        Dim LengthVersion As Integer = vNewVersion.Split(".").Length
        Try
            For k As Integer = 0 To LengthVersion - 1
                If CInt(ParamNew(k)) > CInt(ParamLast(k)) Then
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Me.Close()
        End Try
    End Function

    Private Sub Download(ByVal vSourceURL As String, ByVal vWhereToSave As String, ByVal vKeterangan As String)
        Dim theRequest As HttpWebRequest = WebRequest.Create(vSourceURL)
        Dim theResponse As HttpWebResponse = theRequest.GetResponse
        Dim length As Long = theResponse.ContentLength
        If System.IO.File.Exists(vWhereToSave) = True Then
            System.IO.File.Delete(vWhereToSave)
        End If
        Dim writeStream As New IO.FileStream(vWhereToSave, IO.FileMode.Create)
        Try
            Me.ProgressBar1.Maximum = length + 1
            Me.LblSize.Text = vKeterangan
            While Me.ProgressBar1.Value < length
                Dim readBytes(4095) As Byte
                Dim bytesread As Integer = theResponse.GetResponseStream.Read(readBytes, 0, 4096)
                writeStream.Write(readBytes, 0, bytesread)
                Me.ProgressBar1.Value += bytesread
                Me.LblSize.Text = Me.ProgressBar1.Value.ToString & " kb / " & Me.ProgressBar1.Maximum & " kb"
            End While

            theResponse.GetResponseStream.Close()
            writeStream.Close()

            'set tampilan
            Me.ProgressBar1.Value = 0
            Me.LblTypeFile.Text = ""
            Me.LblSize.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Me.Close()
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Me.RunProsesDownload()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Me.Close()
        End Try
    End Sub

    Private Sub FrmDownload_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Control.CheckForIllegalCrossThreadCalls = False
            Me.txtFileName.Text = ""
            Me.ProgressBar1.Value = 0
            Me.LblTypeFile.Text = ""
            Me.LblSize.Text = ""
            Me.LblSize.Text = "Starting update process"
            Me.PanggilFungsiBuat = EnumFungsi._UpdateViaServer
            Me.BackgroundWorker1.RunWorkerAsync() '-> ga stabil klo pake ini
            'Me.RunProsesDownload()

            'Me.BackgroundWorker2.RunWorkerAsync()
            'Me.Fase = EnumProgress.Start
            'Me.Timer1.Enabled = True
            'Me.RunProsesDownload(Me.PanggilFungsiBuat)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Me.Close()
        End Try
    End Sub

    'Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
    '    Try
    '        Dim ij As New FrmAutoUpdate_Check
    '        ij.Show()

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Dim ServerPath As String = Me.ObjAutoUpdate.GetServerPath()
        Dim fileExists As Boolean

        Try
            Me.Label3.Text = ""
            If Me.Fase = EnumProgress.Start Then
                For Each p As Process In Process.GetProcesses
                    If p.ProcessName = Me.Dta.ScalarQuery_GetProductName.ToString Then
                        p.Kill()
                        Me.Fase = EnumProgress.CheckingTargetDirectory
                    End If
                Next
                'Me.Close()
            ElseIf Me.Fase = EnumProgress.CheckingTargetDirectory Then
                Me.Timer1.Enabled = False
                Me.ProgressBar1.Value = 0
                Me.LblTypeFile.Text = ""
                Me.LblSize.Text = "Checking for download folder"
                Me.CekDirectoryDownload(My.Application.Info.DirectoryPath & "\Download")
                Me.Fase = EnumProgress.CopyFile
                Me.Timer1.Enabled = True
            ElseIf Me.Fase = EnumProgress.CopyFile Then
                Me.Timer1.Enabled = False
                Me.ProgressBar1.Value = 0
                Me.LblTypeFile.Text = ""
                Me.LblSize.Text = "Finding source file from LAN folder"
                System.IO.File.Copy(ServerPath & "\" & "SMESS_Update.smess" _
                                   , My.Application.Info.DirectoryPath & "\Download" & "\" & "SMESS_Update.smess", True)
                Me.Fase = EnumProgress.DecryptFolder
                Me.Timer1.Enabled = True
            ElseIf Me.Fase = EnumProgress.DecryptFolder Then
                Me.Timer1.Enabled = False
                Me.LblSize.Text = "Decrypting update files"
                Me.DecryptFolder(My.Application.Info.DirectoryPath & "\Download\" & "SMESS_Update.smess")
                Me.Fase = EnumProgress.Update
                Me.Timer1.Enabled = True
            ElseIf Me.Fase = EnumProgress.Update Then
                Me.Timer1.Enabled = False
                Me.LblSize.Text = "Updating data"
                Me.UpdateData(My.Application.Info.DirectoryPath & "\Download", My.Application.Info.DirectoryPath, "Software_SMESS_Update")
                Me.LblSize.Text = "Updating version"
                fileExists = My.Computer.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\" & "SMESS_Version_Current.smess")
                If fileExists = True Then
                    'hapus versi lama
                    My.Computer.FileSystem.DeleteFile(My.Application.Info.DirectoryPath & "\" & "SMESS_Version_Current.smess", _
                    FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                End If
                fileExists = My.Computer.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\" & "SMESS_Version.smess")
                If fileExists = True Then
                    'hapus versi lama
                    My.Computer.FileSystem.DeleteFile(My.Application.Info.DirectoryPath & "\" & "SMESS_Version.smess", _
                    FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                End If
                fileExists = My.Computer.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\" & "SMESS_Version.txt")
                If fileExists = True Then
                    'hapus versi lama
                    My.Computer.FileSystem.DeleteFile(My.Application.Info.DirectoryPath & "\" & "SMESS_Version.txt", _
                    FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                End If
                System.IO.File.Copy(ServerPath & "\" & "SMESS_Version.smess" _
                                   , My.Application.Info.DirectoryPath & "\" & "SMESS_Version_Current.smess", True)
                Me.Fase = EnumProgress.DeleteBufferFile
                Me.Timer1.Enabled = True
            ElseIf Me.Fase = EnumProgress.DeleteBufferFile Then
                Me.LblSize.Text = "Deleting buffer file"
                Me.Timer1.Enabled = False
                System.IO.Directory.Delete(My.Application.Info.DirectoryPath & "\Download", True)
                Me.Fase = EnumProgress.StartSMESS
                Me.Timer1.Enabled = True
            ElseIf Me.Fase = EnumProgress.StartSMESS Then
                Me.LblSize.Text = "Starting SMESS"
                Me.Timer1.Enabled = False
                Dim proc As New System.Diagnostics.Process()
                proc = Process.Start(My.Application.Info.DirectoryPath & "\SMESS.exe")
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.BackgroundWorker2.CancelAsync()
        End Try
    End Sub
End Class
