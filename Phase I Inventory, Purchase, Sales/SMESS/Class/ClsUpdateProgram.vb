Option Explicit On
Option Strict Off

Imports System.Net
Imports System.IO
Imports System.Environment
'Imports Setting_Module

Public Class ClsUpdateProgram
    Dim ObjAutoUpdate As New ClsAutoUpdate
    Dim ObjUpdate As New ClsAutoUpdate
    Dim ObjURL As New CustomSMESS.ClsUpdate
    'Dim ObjDBVersion As New ClsDBVersion
    Dim mProgressBar1 As New System.Windows.Forms.ProgressBar

    Public Enum ReasonOfFailure As Integer
        ServerPathDoesntExist

    End Enum

    Public Enum EnumFungsi As Integer
        _download
        _UpdateViaServer
    End Enum

    Private Sub CekDirectoryDownload(ByVal vPath As String)
        Try
            If System.IO.Directory.Exists(vPath) = False Then
                System.IO.Directory.CreateDirectory(vPath)
            Else
                System.IO.Directory.Delete(vPath, True)
                System.IO.Directory.CreateDirectory(vPath)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function ApakahVersionBerbeda(ByVal vLastVersion As String, ByVal vNewVersion As String) As Boolean
        Dim ParamLast As String() = vLastVersion.Split(".")
        Dim ParamNew As String() = vNewVersion.Split(".")
        Dim LengthVersion As Integer = vNewVersion.Split(".").Length

        Try
            'MsgBox("Current version: " + vLastVersion)
            'MsgBox("New version: " + vNewVersion)

            For k As Integer = 0 To LengthVersion - 1
                If CInt(ParamNew(k)) > CInt(ParamLast(k)) Then
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub Download(ByVal vSourceURL As String, ByVal vWhereToSave As String)
        Dim theRequest As HttpWebRequest = WebRequest.Create(vSourceURL)
        Dim theResponse As HttpWebResponse = theRequest.GetResponse
        Dim length As Long = theResponse.ContentLength
        If System.IO.File.Exists(vWhereToSave) = True Then
            System.IO.File.Delete(vWhereToSave)
        End If
        Dim writeStream As New IO.FileStream(vWhereToSave, IO.FileMode.Create)
        Try
            mProgressBar1.Maximum = length
            While mProgressBar1.Value < length
                Dim readBytes(4095) As Byte
                Dim bytesread As Integer = theResponse.GetResponseStream.Read(readBytes, 0, 4096)
                writeStream.Write(readBytes, 0, bytesread)
                mProgressBar1.Value += bytesread
                'Me.LblSize.Text = mProgressBar1.Value.ToString & " kb / " & mProgressBar1.Maximum & " kb"
            End While

            theResponse.GetResponseStream.Close()
            writeStream.Close()

            'set tampilan
            mProgressBar1.Value = 0
            'Me.LblSize.Text = "0 kb / 0 kb"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Public Function ConnectionInternet_ClientUpdate(ByVal vEnumFungsi As EnumFungsi, ByVal vShowMsg As Boolean) As Boolean
    '    Dim ServerPath As String = Me.ObjAutoUpdate.GetServerPath()
    '    Try
    '        If vEnumFungsi = EnumFungsi._download Then
    '            'If My.Computer.Network.Ping("www.google.com") = False Then
    '            '    MsgBox("No internet connection.", MsgBoxStyle.Information, "Information")
    '            '    Return False
    '            'End If
    '            If IsLinkConnectionAvailable("https://www.google.com") = False Then
    '                If vShowMsg = True Then
    '                    MsgBox("No internet connection.", MsgBoxStyle.Information, "Information")
    '                End If
    '                Return False
    '            End If
    '            If Me.ObjAutoUpdate.GetServerPath() = "" Then
    '                If vShowMsg = True Then
    '                    MsgBox("Please setting server path.", MsgBoxStyle.Information, "Information")
    '                End If
    '                Return False
    '            End If
    '            If System.IO.Directory.Exists(ServerPath) = False Then
    '                If vShowMsg = True Then
    '                    MsgBox("Server path doesn't exists.", MsgBoxStyle.Information, "Information")
    '                End If
    '                Return False
    '            End If
    '            'Cek berkaitan dengan Link Downloadnya bisa di akses
    '            If IsLinkConnectionAvailable(Core_Module.GetDownloadLink_ClientUpdateModule_Version) = False Then
    '                If vShowMsg = True Then
    '                    MsgBox("Link download version access denied, please contact your vendor.", MsgBoxStyle.Information, "Information")
    '                End If
    '                Return False
    '            End If
    '            If IsLinkConnectionAvailable(Core_Module.GetDownloadLink_ClientUpdateModule) = False Then
    '                If vShowMsg = True Then
    '                    MsgBox("Link download file access denied, please contact your vendor.", MsgBoxStyle.Information, "Information")
    '                End If
    '                Return False
    '            End If
    '        End If

    '        If vEnumFungsi = EnumFungsi._UpdateViaServer Then
    '            ServerPath = Me.ObjAutoUpdate.GetServerPath()
    '            If ServerPath = "" Then
    '                If vShowMsg = True Then
    '                    MsgBox("Please setting server path.", MsgBoxStyle.Information, "Information")
    '                End If
    '                Return False
    '            End If
    '            If System.IO.Directory.Exists(ServerPath) = False Then
    '                If vShowMsg = True Then
    '                    MsgBox("Server path doesn't exists.", MsgBoxStyle.Information, "Information")
    '                End If
    '                Return False
    '            End If
    '            If System.IO.File.Exists(ServerPath & "\" & "Client_Update_Version.smess") = False Then
    '                If vShowMsg = True Then
    '                    MsgBox("File Client_Update_Version.smess doesn't exists at server path.", MsgBoxStyle.Information, "Information")
    '                End If
    '                Return False
    '            End If
    '            If System.IO.File.Exists(ServerPath & "\" & "Client_Update.smess") = False Then
    '                If vShowMsg = True Then
    '                    MsgBox("File Client_Update.smess doesn't exists at server path.", MsgBoxStyle.Information, "Information")
    '                End If
    '                Return False
    '            End If
    '        End If

    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    Public Function ConnectionInternet(ByVal vEnumFungsi As EnumFungsi, ByVal vShowMsg As Boolean) As Boolean
        Dim ServerPath As String = Me.ObjAutoUpdate.GetServerPath()
        Try
            If vEnumFungsi = EnumFungsi._download Then
                'If My.Computer.Network.Ping("www.google.com") = False Then
                '    MsgBox("No internet connection.", MsgBoxStyle.Information, "Information")
                '    Return False
                'End If
                If IsLinkConnectionAvailable("https://www.google.com") = False Then
                    If vShowMsg = True Then
                        MsgBox("No internet connection.", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
                If Me.ObjAutoUpdate.GetServerPath() = "" Then
                    If vShowMsg = True Then
                        MsgBox("Please setting server path.", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
                If System.IO.Directory.Exists(ServerPath) = False Then
                    If vShowMsg = True Then
                        MsgBox("Server path doesn't exists.", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
                'Cek berkaitan dengan Link Downloadnya bisa di akses
                If IsLinkConnectionAvailable(Me.ObjURL.GetFileVersionTxt_SMESS()) = False Then
                    If vShowMsg = True Then
                        MsgBox("Link download version access denied, please contact your vendor.", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
                If IsLinkConnectionAvailable(Me.ObjURL.GetLinkDownload_SMESS()) = False Then
                    If vShowMsg = True Then
                        MsgBox("Link download file access denied, please contact your vendor.", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            If vEnumFungsi = EnumFungsi._UpdateViaServer Then
                ServerPath = Me.ObjAutoUpdate.GetServerPath()
                If ServerPath = "" Then
                    If vShowMsg = True Then
                        MsgBox("Please setting server path.", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
                If System.IO.Directory.Exists(ServerPath) = False Then
                    If vShowMsg = True Then
                        MsgBox("Server path doesn't exists.", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
                If System.IO.File.Exists(ServerPath & "\" & "SMESS_Version.smess") = False Then
                    If vShowMsg = True Then
                        MsgBox("File SMESS_Version.smess doesn't exists at server path.", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
                If System.IO.File.Exists(ServerPath & "\" & "SMESS_Update.smess") = False Then
                    If vShowMsg = True Then
                        MsgBox("File SMESS_Update.smess doesn't exists at server path.", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function IsLinkConnectionAvailable(ByVal vLink As String) As Boolean
        Dim objUrl As New System.Uri(vLink)
        Dim objWebReq As System.Net.WebRequest
        objWebReq = System.Net.WebRequest.Create(objUrl)
        objWebReq.Timeout = 120000
        'objWebReq.ReadWriteTimeout = 120000

        Dim objresp As System.Net.WebResponse
        Try
            objresp = objWebReq.GetResponse
            objresp.Close()
            objresp = Nothing
            Return True
        Catch ex As Exception
            objresp = Nothing
            objWebReq = Nothing
            Return False
        End Try
    End Function

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
            ObjEncrypt.EncryptOrDecryptFile(FileName, FileNameHasilDecrypt, bytKey, bytIV, ClsEncryptFile.CryptoAction.ActionDecrypt, mProgressBar1)
            'de-compress
            FileNameHasilDeCompress = FileName.Substring(0, FileName.Length - 6) + ".txt"
            ObjZIP.UnZipFile(FileNameHasilDecrypt, FileNameHasilDeCompress)
            'hapus file hasil compress
            If My.Computer.FileSystem.FileExists(FileNameHasilDecrypt) = True Then
                My.Computer.FileSystem.DeleteFile(FileNameHasilDecrypt, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
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
            ObjEncrypt.EncryptOrDecryptFile(FileName, FileNameHasilDecrypt, bytKey, bytIV, ClsEncryptFile.CryptoAction.ActionDecrypt, mProgressBar1)
            'de-compress
            FolderNameHasilDeCompress = FileName.Substring(0, FileName.LastIndexOf("\"))
            ObjZIP.UnZipFolder(FileNameHasilDecrypt, GetEncryptPass_Installer2, FolderNameHasilDeCompress)
            'hapus file hasil decrypt
            If My.Computer.FileSystem.FileExists(FileNameHasilDecrypt) = True Then
                My.Computer.FileSystem.DeleteFile(FileNameHasilDecrypt, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
            'MsgBox("Finish", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjZIP = Nothing
            ObjEncrypt = Nothing
        End Try
    End Sub

    'Public Function RunProsesDownloadUpdate_ClientUpdate(ByVal vEnumFungsi As ClsUpdateProgram.EnumFungsi, ByVal vShowMsg As Boolean) As Boolean
    '    Dim NewVersionTxt As String = "0.0.00"
    '    Dim LastVersionTxt As String = "0.0.00"
    '    Dim ServerPath As String = Me.ObjAutoUpdate.GetServerPath()
    '    Dim ObjVersion As New ClsVersionControl

    '    Try
    '        Windows.Forms.Cursor.Current = Cursors.WaitCursor
    '        If Me.ConnectionInternet(vEnumFungsi, vShowMsg) = False Then
    '            Return False
    '        End If
    '        If vEnumFungsi = EnumFungsi._download Then
    '            'download
    '            Me.Download(Me.ObjURL.GetFileVersionTxt_SMESS, Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update_Version.smess")
    '        ElseIf vEnumFungsi = EnumFungsi._UpdateViaServer Then
    '            If ServerPath = "" Then
    '                If vShowMsg = True Then
    '                    MsgBox("There is no update available.", MsgBoxStyle.Information, "Information")
    '                End If
    '                Return False
    '            End If
    '            'copy File via server
    '            If System.IO.File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update_Current.smess") = True Then
    '                System.IO.File.SetAttributes(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update_Current.smess", FileAttributes.Normal)
    '            End If

    '            System.IO.File.Copy(ServerPath & "\" & "Client_Update.smess" _
    '                               , Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update.smess", True)
    '        End If

    '        'decrypt file
    '        Me.DecryptFile(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update.smess")

    '        'bandingin version
    '        NewVersionTxt = My.Computer.FileSystem.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update.txt")
    '        LastVersionTxt = ObjVersion.GetVersion

    '        If Me.ApakahVersionBerbeda(LastVersionTxt, NewVersionTxt) = False Then
    '            System.IO.File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update.txt")
    '            System.IO.File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update.smess")
    '            If vShowMsg = True Then
    '                MsgBox("Application is up to date", MsgBoxStyle.Information, "Information")
    '            End If
    '            Return False
    '        Else
    '            Return True
    '        End If
    '        Windows.Forms.Cursor.Current = Cursors.Arrow
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        ObjVersion = Nothing
    '    End Try
    'End Function

    Public Function RunProsesDownloadUpdate(ByVal vEnumFungsi As ClsUpdateProgram.EnumFungsi, ByVal vShowMsg As Boolean) As Boolean
        Dim NewVersionTxt As String = "0.0.00"
        Dim LastVersionTxt As String = "0.0.00"
        Dim ServerPath As String = Me.ObjAutoUpdate.GetServerPath()
        Dim ObjVersion As New ClsVersionControl

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If Me.ConnectionInternet(vEnumFungsi, vShowMsg) = False Then
                Return False
            End If
            If vEnumFungsi = EnumFungsi._download Then
                'download
                Me.Download(Me.ObjURL.GetFileVersionTxt_SMESS, Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version.smess")
            ElseIf vEnumFungsi = EnumFungsi._UpdateViaServer Then
                If ServerPath = "" Then
                    If vShowMsg = True Then
                        MsgBox("There is no update available.", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
                'copy File via server
                If System.IO.File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current.smess") = True Then
                    System.IO.File.SetAttributes(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current.smess", FileAttributes.Normal)
                End If

                System.IO.File.Copy(ServerPath & "\" & "SMESS_Version.smess" _
                                   , Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version.smess", True)
            End If

            'decrypt file
            Me.DecryptFile(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version.smess")
            'MsgBox("Checking for: " + Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version.txt")
            'If My.Computer.FileSystem.FileExists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version.txt") = False Then
            '    MsgBox(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version.txt doesnot exits.", MsgBoxStyle.Information, "Information")
            'End If
            'bandingin version
            NewVersionTxt = My.Computer.FileSystem.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version.txt")
            ''NewVersionTxt = "1.7.00"
            LastVersionTxt = ObjVersion.GetVersion
            'MsgBox("Please cek file")
            If Me.ApakahVersionBerbeda(LastVersionTxt, NewVersionTxt) = False Then
                System.IO.File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version.txt")
                System.IO.File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version.smess")
                If vShowMsg = True Then
                    If MsgBox("We detect your smess is up to date, update anyway?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Return True
                    Else
                        Return False
                    End If
                End If
                Return False
            Else
                Return True
            End If
            Windows.Forms.Cursor.Current = Cursors.Arrow
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            ObjVersion = Nothing
        End Try
    End Function

#Region "BackUpSource"
    Public Function RunProsesDownloadUpdate1(ByVal vEnumFungsi As ClsUpdateProgram.EnumFungsi) As Boolean
        Dim NewVersionTxt As String = "0.0.00"
        Dim LastVersionTxt As String = "0.0.00"
        Dim ServerPath As String = Me.ObjAutoUpdate.GetServerPath()
        Dim ObjVersion As New ClsVersionControl

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If Me.ConnectionInternet(vEnumFungsi, True) = False Then
                Return False
            End If
            'Cek Directory Local
            Me.CekDirectoryDownload(My.Application.Info.DirectoryPath & "\Download")
            If vEnumFungsi = EnumFungsi._download Then
                'download
                Me.Download(Me.ObjURL.GetFileVersionTxt_SMESS, My.Application.Info.DirectoryPath & "\Download\" & "SMESS_Version.smess")
            ElseIf vEnumFungsi = EnumFungsi._UpdateViaServer Then
                'copy File via server
                System.IO.File.Copy(ServerPath & "\" & "SMESS_Version.smess" _
                                   , My.Application.Info.DirectoryPath & "\Download" & "\" & "SMESS_Version.smess", True)
            End If
            'decrypt file
            Me.DecryptFile(My.Application.Info.DirectoryPath & "\Download\" & "SMESS_Version.smess")
            'bandingin version
            'NewVersionTxt = My.Computer.FileSystem.ReadAllText(Me.DefaultServer & "\Download\" & "SMESS_Version.txt")
            NewVersionTxt = My.Computer.FileSystem.ReadAllText(My.Application.Info.DirectoryPath & "\Download\" & "SMESS_Version.txt")
            LastVersionTxt = ObjVersion.GetVersion
            If Me.ApakahVersionBerbeda(LastVersionTxt, NewVersionTxt) = False Then
                NewVersionTxt = "1.7.00"
                System.IO.Directory.Delete(My.Application.Info.DirectoryPath & "\Download", True)
                'System.IO.Directory.Delete(Me.DefaultServer & "\Download", True)
                MsgBox("Application is up to date", MsgBoxStyle.Information, "Information")
                Return False
            Else
                Return True
            End If
            Windows.Forms.Cursor.Current = Cursors.Arrow
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjVersion = Nothing
        End Try
    End Function
#End Region


    Public Sub SetTipeDownload(ByVal vTipeFungsi As ClsUpdateProgram.EnumFungsi)

    End Sub
End Class
