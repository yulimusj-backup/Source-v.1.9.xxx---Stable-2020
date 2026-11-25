Option Explicit On
Option Strict Off

Imports System.Net
Imports System.IO
Imports System.Environment

Public Class FrmDownload
    Dim whereToSave As String
    Dim Dta As New DseAutoUpdateTableAdapters.tbl_setting_updateTableAdapter
    Dim ObjUpdate As New ClsAutoUpdate
    Dim ObjURL As New CustomSMESS.ClsUpdate
    'Dim ObjDBVersion As New ClsDBVersion
    Dim ObjAutoUpdate As New ClsAutoUpdate
    Dim PanggilFungsiBuat As EnumFungsi
    Dim Succeed As Boolean = False
    Dim Proceed As Boolean = True

    Public Progres As EnumProgress = EnumProgress.Start
    'Public Download_SMESS As Integer = 0
    'Public Download_ClientUpdate As Integer = 0

    Public Enum enumProgress As Integer
        Start
        CheckFolder
        Download
        CopyFile
        DeleteFile
        CancelledByUser
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
            Me.LblTypeFile.Text = "Decrypt Smess Version"
            'de-compress
            FileNameHasilDeCompress = FileName.Substring(0, FileName.Length - 6) + ".txt"
            ObjZIP.UnZipFile(FileNameHasilDecrypt, FileNameHasilDeCompress)
            'hapus file hasil compress
            If My.Computer.FileSystem.FileExists(FileNameHasilDecrypt) = True Then
                My.Computer.FileSystem.DeleteFile(FileNameHasilDecrypt, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
            'Set Progress Bar default
            Me.ProgressBar1.Value = 0
            Me.LblTypeFile.Text = ""
            'MsgBox("Finish", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjZIP = Nothing
            ObjEncrypt = Nothing
            Me.Close()
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
            Me.LblTypeFile.Text = "Decrypt Smess Update"
            'de-compress
            FolderNameHasilDeCompress = FileName.Substring(0, FileName.LastIndexOf("\"))
            ObjZIP.UnZipFolder(FileNameHasilDecrypt, GetEncryptPass_Installer2, FolderNameHasilDeCompress)
            'hapus file hasil decrypt
            If My.Computer.FileSystem.FileExists(FileNameHasilDecrypt) = True Then
                My.Computer.FileSystem.DeleteFile(FileNameHasilDecrypt, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
            'Set Progress Bar Value
            Me.ProgressBar1.Value = 0
            Me.LblTypeFile.Text = ""
            'MsgBox("Finish", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjZIP = Nothing
            ObjEncrypt = Nothing
            Me.Close()
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

    Public Sub RunProsesDownload(ByVal vEnumFungsi As EnumFungsi)
        Dim ServerPath As String = Me.ObjAutoUpdate.GetServerPath()

        Try
            Me.LblTypeFile.Text = ""
            Me.LblTypeFile.Text = "Please wait, while we contact our server."
            If System.IO.Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Download") = True Then
                System.IO.Directory.Delete(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Download", True)
            End If

            'klo version lebih besar maka download file smess + update
            Me.Progres = enumProgress.CheckFolder
            If Me.Proceed = True Then
                Me.CekDirectoryDownload(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Download")
            Else
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Me.Proceed = False
            End If

            'download smess
            If Me.Proceed = True Then
                Me.Progres = enumProgress.Download
                'If Me.Download_SMESS = 1 Then
                Me.Download(Me.ObjURL.GetLinkDownload_SMESS, Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Download\" & "SMESS_Update.smess", "Downloading SMESS_Update.smess")
                Me.Download(Me.ObjURL.GetFileVersionTxt_SMESS, Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Download\" & "SMESS_Version.smess", "Downloading SMESS_Version.smess")
                'End If
                'If Me.Download_ClientUpdate = 1 Then
                '    Me.Download(Me.ObjURL.GetLinkDownload_SMESS, Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Download\" & "SMESS_Update.smess", "Downloading SMESS_Update.smess")
                '    Me.Download(Me.ObjURL.GetFileVersionTxt_SMESS, Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Download\" & "SMESS_Version.smess", "Downloading SMESS_Version.smess")
                'End If
            End If

            'decrypt folder
            If Me.Proceed = False Then
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
                'Me.Close()
                'Exit Sub
            End If
            If Me.Proceed = True Then
                Me.btnCancel.Enabled = False
                Me.Progres = enumProgress.CopyFile
                Me.LblTypeFile.Text = "Copying required files to target folder..."
                'If Me.Download_SMESS = 1 Then
                System.IO.File.Copy(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Download" & "\" & "SMESS_Update.smess" _
                                               , ServerPath & "\" & "SMESS_Update.smess", True)
                System.IO.File.Copy(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Download" & "\" & "SMESS_Version.smess" _
                                   , ServerPath & "\" & "SMESS_Version.smess", True)
                'End If
                'If Me.Download_ClientUpdate = 1 Then
                '    System.IO.File.Copy(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Download" & "\" & "Client_Update.smess" _
                '                                                       , ServerPath & "\" & "Client_Update.smess", True)
                '    System.IO.File.Copy(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Download" & "\" & "Client_Update_Version.smess" _
                '                       , ServerPath & "\" & "Client_Update_Version.smess", True)
                'End If
            End If

            'Hapus directory download
            If Me.Proceed = False Then
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
                'Me.Close()
            End If
            Me.Progres = enumProgress.DeleteFile
            Me.LblTypeFile.Text = "Deleting required files to target folder..."
            If My.Computer.FileSystem.DirectoryExists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Download") = True Then
                System.IO.Directory.Delete(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Download", True)
            End If

            If Me.Proceed = True Then
                Me.Succeed = True
            Else
                Me.Succeed = False
            End If
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
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
        Dim Extension As String = ""

        Try
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = ListFileDownload.Length
            Me.LblTypeFile.Text = vKeterangan
            'ambil semua list file yang ada di file instalasi
            For i As Integer = 1 To ListFileDownload.Length
                Dim NamaFile As String() = ListFileDownload(i - 1).Split("\")
                Extension = System.IO.Path.GetExtension(vDownloadPath & "\" & NamaFile(NamaFile.Length - 1))
                If Extension = ".dll" Or Extension = ".exe" Or Extension = ".config" Then
                    If Me.ApakahBolehUpdateData(NamaFile(NamaFile.Length - 1).ToString) = True Then
                        System.IO.File.Copy(vDownloadPath & "\" & NamaFile(NamaFile.Length - 1) _
                                           , vProgramPath & "\" & NamaFile(NamaFile.Length - 1), True)
                    End If
                End If
                Me.ProgressBar1.Value += 1
                Me.LblTypeFile.Text = Me.ProgressBar1.Value.ToString & "file(s) of " & Me.ProgressBar1.Maximum & " file(s)"
            Next i
            'set tampilan
            Me.ProgressBar1.Value = 0
            Me.LblTypeFile.Text = ""
            Me.LblTypeFile.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Me.Close()
        End Try
    End Sub

    Private Sub CekDirectoryDownload(ByVal vPath As String)
        Try
            Me.ProgressBar1.Value = 0
            Me.LblTypeFile.Text = ""
            Me.LblTypeFile.Text = "Checking target folder"
            If System.IO.Directory.Exists(vPath) = False Then
                System.IO.Directory.CreateDirectory(vPath)
            Else
                System.IO.Directory.Delete(vPath, True)
                System.IO.Directory.CreateDirectory(vPath)
            End If
            Me.ProgressBar1.Value = 0
            Me.LblTypeFile.Text = ""
            Me.LblTypeFile.Text = ""

            If System.IO.Directory.Exists(vPath) = False Then
                MsgBox("Unable to write on target folder.", MsgBoxStyle.Information, "Information")
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
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
        'time out 2 menit
        theRequest.Timeout = 120000
        theRequest.ReadWriteTimeout = 120000

        Dim theResponse As HttpWebResponse = theRequest.GetResponse
        Dim length As Long = theResponse.ContentLength
        If System.IO.File.Exists(vWhereToSave) = True Then
            System.IO.File.Delete(vWhereToSave)
        End If
        Dim writeStream As New IO.FileStream(vWhereToSave, IO.FileMode.Create)
        Dim ObjCD As New Core_Module.ClsCekData
        'Dim Fail, MaxFail As Integer
        Dim Lama As Double = -1
        Dim Baru As Double = -1
        Dim Pembagi As Double = 1024 ^ 2
        Dim FileSize As String = ""

        Try
            Me.ProgressBar1.Maximum = length + 1
            Me.LblTypeFile.Text = vKeterangan
            'Me.LblTypeFile.Text = "File Size: "
            Me.LblSize.Left = Me.LblTypeFile.Left + Me.LblTypeFile.Width + 5
            'MaxFail = 3

            If Me.ProgressBar1.Maximum / Pembagi > 100 Then
                Pembagi *= 10
            End If
            FileSize = Math.Round(Me.ProgressBar1.Maximum / Pembagi, 2, MidpointRounding.AwayFromZero)

            While Me.ProgressBar1.Value < length
                If Me.Proceed = False Then
                    Me.DialogResult = Windows.Forms.DialogResult.Cancel
                    theResponse.GetResponseStream.Close()
                    writeStream.Close()
                    Exit Sub
                End If

                Dim readBytes(4095) As Byte
                Dim bytesread As Integer = theResponse.GetResponseStream.Read(readBytes, 0, 4096)

                writeStream.Write(readBytes, 0, bytesread)
                Me.ProgressBar1.Value += bytesread
                Me.LblTypeFile.Text = vKeterangan + " - File size: " + FileSize + " MB - " + Math.Round(CDec(CDec(Me.ProgressBar1.Value) * CDec(100) / CDec(Me.ProgressBar1.Maximum)), 2, MidpointRounding.AwayFromZero).ToString + "%"

            End While

            theResponse.GetResponseStream.Close()
            writeStream.Close()

            'set tampilan
            Me.ProgressBar1.Value = 0
            Me.LblTypeFile.Text = ""
            Me.LblTypeFile.Text = ""
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        Finally
            ObjCD = Nothing
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            'If Me.Download_SMESS = 1 Or Me.Download_ClientUpdate = 1 Then
            Me.RunProsesDownload(Me.PanggilFungsiBuat)
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Me.Close()
        End Try
    End Sub

    Private Sub FrmDownload_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.Label1.Text = ""
            Me.LblTypeFile.Text = ""
            Me.LblTypeFile.Text = ""
            Me.LblTypeFile.Text = ""
            Me.Label4.Text = ""

            Me.btnCancel.Top = 43
            Me.ProgressBar1.Top = Me.btnCancel.Top

            Me.txtFileName.Top = 13
            Me.Label1.Top = Me.txtFileName.Top
            Me.Label3.Top = Me.txtFileName.Top
            Me.LblSize.Top = Me.txtFileName.Top
            Me.LblTypeFile.Top = Me.txtFileName.Top
            Me.Label4.Top = Me.txtFileName.Top

            Control.CheckForIllegalCrossThreadCalls = False
            Me.txtFileName.Text = ""
            Me.PanggilFungsiBuat = EnumFungsi._download
            Me.BackgroundWorker1.RunWorkerAsync() '-> ga stabil klo pake ini
            'Me.RunProsesDownload(Me.PanggilFungsiBuat)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Me.Close()
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.Succeed = True Then
                'If Me.Download_ClientUpdate = 1 Then
                '    Me.LblTypeFile.Text = "Updating Client Update Module."
                '    Me.RunProsesUpdate_ClientUpdate 
                'End If

                Me.LblTypeFile.Text = "Download Completed."
                MsgBox("Download Completed.", MsgBoxStyle.Information, "Information")
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            Else
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Me.Close()
            End If
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Public Sub RunProsesUpdate_ClientUpdate()
    '    Dim ServerPath As String = Me.ObjAutoUpdate.GetServerPath()
    '    Dim fileExists As Boolean

    '    Try
    '        System.IO.File.SetAttributes(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update_Current.smess", FileAttributes.Normal)
    '        Me.Label3.Text = ""
    '        Me.Label1.Text = ""
    '        Me.Label4.Text = ""
    '        Me.LblSize.Text = ""
    '        Me.LblTypeFile.Text = ""
    '        Me.LinkLabel1.Text = ""

    '        Me.ProgressBar1.Value = 0
    '        Me.LblTypeFile.Text = ""
    '        Me.LblTypeFile.Text = "Updating Client Update Module - Checking for download folder"
    '        Me.CekDirectoryDownload(Environment.GetFolderPath(SpecialFolder.MyDocuments) & "\Download")

    '        Me.ProgressBar1.Value = 0
    '        Me.LblTypeFile.Text = ""
    '        Me.LblTypeFile.Text = "Updating Client Update Module - Finding source file from LAN folder"
    '        System.IO.File.Copy(ServerPath & "\" & "Client_Update.smess" _
    '                           , Environment.GetFolderPath(SpecialFolder.MyDocuments) & "\Download" & "\" & "Client_Update.smess", True)

    '        Me.LblTypeFile.Text = "Updating Client Update Module - Decrypting update files"
    '        Me.DecryptFolder(Environment.GetFolderPath(SpecialFolder.MyDocuments) & "\Download\" & "Client_Update.smess")

    '        Me.LblTypeFile.Text = "Updating Client Update Module - Updating data"
    '        Me.UpdateData(Environment.GetFolderPath(SpecialFolder.MyDocuments) & "\Download", My.Application.Info.DirectoryPath, "Updating Client Update Module")

    '        Me.LblTypeFile.Text = "Updating Client Update Module - Updating version"
    '        fileExists = My.Computer.FileSystem.FileExists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update_Version_Current.smess")
    '        If fileExists = True Then
    '            'hapus versi lama
    '            My.Computer.FileSystem.DeleteFile(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update_Version_Current.smess", _
    '            FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
    '        End If
    '        fileExists = My.Computer.FileSystem.FileExists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update_Version.smess")
    '        If fileExists = True Then
    '            'hapus versi lama
    '            My.Computer.FileSystem.DeleteFile(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update_Version.smess", _
    '            FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
    '        End If

    '        fileExists = My.Computer.FileSystem.FileExists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update_Version.txt")
    '        If fileExists = True Then
    '            'hapus versi lama
    '            My.Computer.FileSystem.DeleteFile(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update_Version.txt", _
    '            FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
    '        End If
    '        fileExists = My.Computer.FileSystem.FileExists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update_Version_Current.txt")
    '        If fileExists = True Then
    '            'hapus versi lama
    '            My.Computer.FileSystem.DeleteFile(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update_Version_Current.txt", _
    '            FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
    '        End If
    '        System.IO.File.Copy(ServerPath & "\" & "Client_Update_Version.smess" _
    '                           , Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update_Version_Current.smess", True)
    '        Me.LblTypeFile.Text = "Deleting buffer file"
    '        System.IO.Directory.Delete(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Download", True)
    '        System.IO.File.SetAttributes(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update_Version_Current.smess", FileAttributes.Hidden)
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Me.Close()
    '    Finally
    '        Me.BackgroundWorker2.CancelAsync()
    '    End Try
    'End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click

        Try
            Me.LblTypeFile.Text = "Cancelling process"
            Me.Proceed = False
            Me.Progres = enumProgress.CancelledByUser
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        End Try
    End Sub
End Class
