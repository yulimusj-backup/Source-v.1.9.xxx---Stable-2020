Imports Core_Module
Imports System.Environment

Public Class FrmUpdatePrototipe
    'Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
    '    Dim ObjAuto As New Download_Module.ClsUpdateProgram
    '    Dim Hasil As MsgBoxResult
    '    Try
    '        If ObjAuto.RunProsesDownloadUpdate(Download_Module.ClsUpdateProgram.EnumFungsi._UpdateViaServer, True) = True Then
    '            Hasil = MsgBox("Update ?", MsgBoxStyle.YesNo, "Confirm")
    '            If Hasil = MsgBoxResult.Yes Then
    '                Dim proc As New System.Diagnostics.Process()
    '                proc = Process.Start(CurrentDirectory & "\Client Update Module.exe")
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ObjAuto = Nothing
    '    End Try
    'End Sub

    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
    '    Dim ObjAuto As New Download_Module.ClsUpdateProgram
    '    Dim Hasil As MsgBoxResult

    '    Try
    '        If ObjAuto.RunProsesDownloadUpdate(Download_Module.ClsUpdateProgram.EnumFungsi._download, True) = True Then
    '            If ObjAuto.ConnectionInternet(Download_Module.ClsUpdateProgram.EnumFungsi._download, True) = True Then
    '                Hasil = MsgBox("Download ?", MsgBoxStyle.YesNo, "Confirm")
    '                If Hasil = MsgBoxResult.Yes Then
    '                    Dim proc As New System.Diagnostics.Process()
    '                    proc = Process.Start(CurrentDirectory & "\Download Module2.exe")
    '                End If
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ObjAuto = Nothing
    '    End Try
    'End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        'encrypt: compress -> encrypt
        Dim ObjZIP As New Core_Module.ClsCompressFile
        Dim ObjEncrypt As New Core_Module.ClsEncryptFile
        Dim FileName As String = ""
        Dim FileNameHasilCompress As String = ""
        Dim FileNameHasilEncryt As String = ""
        Dim bytKey As Byte()
        Dim bytIV As Byte()
        Dim vEncryptPass As String = Core_Module.GetEncryptPass_Installer

        Try
            'pilih file *.sql
            Me.OpenFileDialog1.Title = "Please specify file to encrypt"
            Me.OpenFileDialog1.Filter = "text files (*.txt)|*.txt"
            Me.OpenFileDialog1.FilterIndex = 1
            Me.OpenFileDialog1.DefaultExt = "txt"
            Me.OpenFileDialog1.CheckFileExists = True
            Me.OpenFileDialog1.Multiselect = False
            Me.OpenFileDialog1.FileName = ""

            If Me.OpenFileDialog1.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If
            If Me.OpenFileDialog1.FileName = "" Then
                Exit Sub
            End If

            FileName = Me.OpenFileDialog1.FileName
            'compress
            FileNameHasilCompress = FileName.Substring(0, FileName.Length - 4) + ".zip"
            ObjZIP.ZipFile(FileName, FileNameHasilCompress)
            'encrypt + rename
            bytKey = ObjEncrypt.CreateKey(vEncryptPass)
            'Send the password to the CreateIV function.
            bytIV = ObjEncrypt.CreateIV(vEncryptPass)
            FileNameHasilEncryt = FileName.Substring(0, FileName.Length - 4) + ".smess"
            ObjEncrypt.EncryptOrDecryptFile(FileNameHasilCompress, FileNameHasilEncryt, bytKey, bytIV, ClsEncryptFile.CryptoAction.ActionEncrypt, Me.ProgressBar1)
            Me.ProgressBar1.Value = 0
            'hapus file hasil compress
            If My.Computer.FileSystem.FileExists(FileNameHasilCompress) = True Then
                My.Computer.FileSystem.DeleteFile(FileNameHasilCompress, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
            MsgBox("Finish", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjZIP = Nothing
            ObjEncrypt = Nothing
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        'decrypt: decrypt -> decompress
        Dim ObjZIP As New Core_Module.ClsCompressFile
        Dim ObjEncrypt As New Core_Module.ClsEncryptFile
        Dim FileName As String = ""
        Dim FileNameHasilDeCompress As String = ""
        Dim FileNameHasilDecrypt As String = ""
        Dim bytKey As Byte()
        Dim bytIV As Byte()
        Dim vEncryptPass As String = Core_Module.GetEncryptPass_Installer

        Try
            'pilih file *.smess
            Me.OpenFileDialog1.Title = "Please specify file to decrypt"
            Me.OpenFileDialog1.Filter = "smess files (*.smess)|*.smess"
            Me.OpenFileDialog1.FilterIndex = 1
            Me.OpenFileDialog1.DefaultExt = "smess"
            Me.OpenFileDialog1.CheckFileExists = True
            Me.OpenFileDialog1.Multiselect = False
            Me.OpenFileDialog1.FileName = ""

            If Me.OpenFileDialog1.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If
            If Me.OpenFileDialog1.FileName = "" Then
                Exit Sub
            End If

            FileName = Me.OpenFileDialog1.FileName
            'decrypt 
            bytKey = ObjEncrypt.CreateKey(vEncryptPass)
            'Send the password to the CreateIV function.
            bytIV = ObjEncrypt.CreateIV(vEncryptPass)
            FileNameHasilDecrypt = FileName.Substring(0, FileName.Length - 6) + ".zip"
            ObjEncrypt.EncryptOrDecryptFile(FileName, FileNameHasilDecrypt, bytKey, bytIV, ClsEncryptFile.CryptoAction.ActionDecrypt, Me.ProgressBar1)
            Me.ProgressBar1.Value = 0
            'de-compress
            FileNameHasilDeCompress = FileName.Substring(0, FileName.Length - 6) + ".txt"
            ObjZIP.UnZipFile(FileNameHasilDecrypt, FileNameHasilDeCompress)
            'hapus file hasil compress
            If My.Computer.FileSystem.FileExists(FileNameHasilDecrypt) = True Then
                My.Computer.FileSystem.DeleteFile(FileNameHasilDecrypt, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
            MsgBox("Finish", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjZIP = Nothing
            ObjEncrypt = Nothing
        End Try
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        'encrypt: compress -> encrypt
        Dim ObjZIP As New Core_Module.ClsCompressFile
        Dim ObjEncrypt As New Core_Module.ClsEncryptFile
        Dim FileName As String = ""
        Dim FileNameHasilCompress As String = ""
        Dim FileNameHasilEncryt As String = ""
        Dim bytKey As Byte()
        Dim bytIV As Byte()
        Dim vEncryptPass As String = Core_Module.GetEncryptPass_Installer

        Try
            'pilih file *.sql
            Me.OpenFileDialog1.Title = "Please specify file to encrypt"
            Me.OpenFileDialog1.Filter = "smess custom (*.dll)|*.dll"
            Me.OpenFileDialog1.FilterIndex = 1
            Me.OpenFileDialog1.DefaultExt = "dll"
            Me.OpenFileDialog1.CheckFileExists = True
            Me.OpenFileDialog1.Multiselect = False
            Me.OpenFileDialog1.FileName = ""

            If Me.OpenFileDialog1.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If
            If Me.OpenFileDialog1.FileName = "" Then
                Exit Sub
            End If

            FileName = Me.OpenFileDialog1.FileName
            'compress
            FileNameHasilCompress = FileName.Substring(0, FileName.Length - 4) + ".zip"
            ObjZIP.ZipFile(FileName, FileNameHasilCompress)
            'encrypt + rename
            bytKey = ObjEncrypt.CreateKey(vEncryptPass)
            'Send the password to the CreateIV function.
            bytIV = ObjEncrypt.CreateIV(vEncryptPass)
            FileNameHasilEncryt = FileName.Substring(0, FileName.Length - 4) + ".smess"
            ObjEncrypt.EncryptOrDecryptFile(FileNameHasilCompress, FileNameHasilEncryt, bytKey, bytIV, ClsEncryptFile.CryptoAction.ActionEncrypt, Me.ProgressBar1)
            Me.ProgressBar1.Value = 0
            'hapus file hasil compress
            If My.Computer.FileSystem.FileExists(FileNameHasilCompress) = True Then
                My.Computer.FileSystem.DeleteFile(FileNameHasilCompress, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
            MsgBox("Finish", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjZIP = Nothing
            ObjEncrypt = Nothing
        End Try
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        'decrypt: decrypt -> decompress
        Dim ObjZIP As New Core_Module.ClsCompressFile
        Dim ObjEncrypt As New Core_Module.ClsEncryptFile
        Dim FileName As String = ""
        Dim FileNameHasilDeCompress As String = ""
        Dim FileNameHasilDecrypt As String = ""
        Dim bytKey As Byte()
        Dim bytIV As Byte()
        Dim vEncryptPass As String = Core_Module.GetEncryptPass_Installer

        Try
            'pilih file *.smess
            Me.OpenFileDialog1.Title = "Please specify file to decrypt"
            Me.OpenFileDialog1.Filter = "smess files (*.smess)|*.smess"
            Me.OpenFileDialog1.FilterIndex = 1
            Me.OpenFileDialog1.DefaultExt = "smess"
            Me.OpenFileDialog1.CheckFileExists = True
            Me.OpenFileDialog1.Multiselect = False
            Me.OpenFileDialog1.FileName = ""

            If Me.OpenFileDialog1.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If
            If Me.OpenFileDialog1.FileName = "" Then
                Exit Sub
            End If

            FileName = Me.OpenFileDialog1.FileName
            'decrypt 
            bytKey = ObjEncrypt.CreateKey(vEncryptPass)
            'Send the password to the CreateIV function.
            bytIV = ObjEncrypt.CreateIV(vEncryptPass)
            FileNameHasilDecrypt = FileName.Substring(0, FileName.Length - 6) + ".zip"
            ObjEncrypt.EncryptOrDecryptFile(FileName, FileNameHasilDecrypt, bytKey, bytIV, ClsEncryptFile.CryptoAction.ActionDecrypt, Me.ProgressBar1)
            Me.ProgressBar1.Value = 0
            'de-compress
            FileNameHasilDeCompress = FileName.Substring(0, FileName.Length - 6) + ".dll"
            ObjZIP.UnZipFile(FileNameHasilDecrypt, FileNameHasilDeCompress)
            'hapus file hasil compress
            If My.Computer.FileSystem.FileExists(FileNameHasilDecrypt) = True Then
                My.Computer.FileSystem.DeleteFile(FileNameHasilDecrypt, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
            MsgBox("Finish", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjZIP = Nothing
            ObjEncrypt = Nothing
        End Try
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        'decrypt: decrypt -> decompress
        Dim ObjZIP As New Core_Module.ClsCompressFile
        Dim ObjEncrypt As New Core_Module.ClsEncryptFile
        Dim FileName As String = ""
        Dim FolderNameHasilDeCompress As String = ""
        Dim FileNameHasilDecrypt As String = ""
        Dim bytKey As Byte()
        Dim bytIV As Byte()
        Dim vEncryptPass As String = Core_Module.GetEncryptPass_Installer

        Try
            'pilih file *.smess
            Me.OpenFileDialog1.Title = "Please specify file to decrypt"
            Me.OpenFileDialog1.Filter = "smess files (*.smess)|*.smess"
            Me.OpenFileDialog1.FilterIndex = 1
            Me.OpenFileDialog1.DefaultExt = "smess"
            Me.OpenFileDialog1.CheckFileExists = True
            Me.OpenFileDialog1.Multiselect = False
            Me.OpenFileDialog1.FileName = ""

            If Me.OpenFileDialog1.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If
            If Me.OpenFileDialog1.FileName = "" Then
                Exit Sub
            End If

            FileName = Me.OpenFileDialog1.FileName
            'decrypt 
            bytKey = ObjEncrypt.CreateKey(vEncryptPass)
            'Send the password to the CreateIV function.
            bytIV = ObjEncrypt.CreateIV(vEncryptPass)
            FileNameHasilDecrypt = FileName.Substring(0, FileName.Length - 6) + ".zip"
            ObjEncrypt.EncryptOrDecryptFile(FileName, FileNameHasilDecrypt, bytKey, bytIV, ClsEncryptFile.CryptoAction.ActionDecrypt, Me.ProgressBar1)
            Me.ProgressBar1.Value = 0
            'de-compress
            FolderNameHasilDeCompress = FileName.Substring(0, FileName.LastIndexOf("\"))
            ObjZIP.UnZipFolder(FileNameHasilDecrypt, Core_Module.GetEncryptPass_Installer2, FolderNameHasilDeCompress)
            'hapus file hasil decrypt
            If My.Computer.FileSystem.FileExists(FileNameHasilDecrypt) = True Then
                My.Computer.FileSystem.DeleteFile(FileNameHasilDecrypt, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
            MsgBox("Finish", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjZIP = Nothing
            ObjEncrypt = Nothing
        End Try
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        'encrypt: compress -> encrypt
        Dim ObjZIP As New Core_Module.ClsCompressFile
        Dim ObjEncrypt As New Core_Module.ClsEncryptFile
        Dim FolderPath As String = ""
        Dim FilePathHasilCompress As String = ""
        Dim FilePathHasilEncryt As String = ""
        Dim bytKey As Byte()
        Dim bytIV As Byte()
        Dim vEncryptPass As String = Core_Module.GetEncryptPass_Installer
        Dim fileExists As Boolean

        Try
            'pilih folder yang mau di kompress
            Me.FolderBrowserDialog1.Description = "Please specify folder to encrypt"

            If Me.FolderBrowserDialog1.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If
            If Me.FolderBrowserDialog1.SelectedPath = "" Then
                Exit Sub
            End If
            FolderPath = Me.FolderBrowserDialog1.SelectedPath

            'compress
            FilePathHasilCompress = FolderPath.Substring(0, FolderPath.LastIndexOf("\")) + "\SMESS_Update.zip"
            fileExists = My.Computer.FileSystem.FileExists(FilePathHasilCompress)
            If fileExists = True Then
                My.Computer.FileSystem.DeleteFile(FilePathHasilCompress, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
            ObjZIP.ZipFolder(FilePathHasilCompress, Core_Module.GetEncryptPass_Installer2, FolderPath)

            'encrypt + rename
            bytKey = ObjEncrypt.CreateKey(vEncryptPass)
            'Send the password to the CreateIV function.
            bytIV = ObjEncrypt.CreateIV(vEncryptPass)
            FilePathHasilEncryt = FolderPath.Substring(0, FolderPath.LastIndexOf("\")) + "\SMESS_Update.smess"
            fileExists = My.Computer.FileSystem.FileExists(FilePathHasilEncryt)
            If fileExists = True Then
                My.Computer.FileSystem.DeleteFile(FilePathHasilEncryt, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
            ObjEncrypt.EncryptOrDecryptFile(FilePathHasilCompress, FilePathHasilEncryt, bytKey, bytIV, ClsEncryptFile.CryptoAction.ActionEncrypt, Me.ProgressBar1)
            Me.ProgressBar1.Value = 0

            'hapus file hasil compress
            If My.Computer.FileSystem.FileExists(FilePathHasilCompress) = True Then
                My.Computer.FileSystem.DeleteFile(FilePathHasilCompress, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If

            MsgBox("Finish", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjZIP = Nothing
            ObjEncrypt = Nothing
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim i As New FrmSettingUpdate
        Try
            i.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub FrmUpdatePrototipe_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'encrypt: compress -> encrypt
        Dim ObjZIP As New Core_Module.ClsCompressFile
        Dim ObjEncrypt As New Core_Module.ClsEncryptFile
        Dim FolderPath As String = ""
        Dim FilePathHasilCompress As String = ""
        Dim FilePathHasilEncryt As String = ""
        Dim bytKey As Byte()
        Dim bytIV As Byte()
        Dim vEncryptPass As String = Core_Module.GetEncryptPass_Installer
        Dim fileExists As Boolean

        Try
            'pilih folder yang mau di kompress
            Me.FolderBrowserDialog1.Description = "Please specify folder to encrypt"

            If Me.FolderBrowserDialog1.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If
            If Me.FolderBrowserDialog1.SelectedPath = "" Then
                Exit Sub
            End If
            FolderPath = Me.FolderBrowserDialog1.SelectedPath

            'compress
            FilePathHasilCompress = FolderPath.Substring(0, FolderPath.LastIndexOf("\")) + "\Client_Update.zip"
            fileExists = My.Computer.FileSystem.FileExists(FilePathHasilCompress)
            If fileExists = True Then
                My.Computer.FileSystem.DeleteFile(FilePathHasilCompress, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
            ObjZIP.ZipFolder(FilePathHasilCompress, Core_Module.GetEncryptPass_Installer2, FolderPath)

            'encrypt + rename
            bytKey = ObjEncrypt.CreateKey(vEncryptPass)
            'Send the password to the CreateIV function.
            bytIV = ObjEncrypt.CreateIV(vEncryptPass)
            FilePathHasilEncryt = FolderPath.Substring(0, FolderPath.LastIndexOf("\")) + "\Client_Update.smess"
            fileExists = My.Computer.FileSystem.FileExists(FilePathHasilEncryt)
            If fileExists = True Then
                My.Computer.FileSystem.DeleteFile(FilePathHasilEncryt, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
            ObjEncrypt.EncryptOrDecryptFile(FilePathHasilCompress, FilePathHasilEncryt, bytKey, bytIV, ClsEncryptFile.CryptoAction.ActionEncrypt, Me.ProgressBar1)
            Me.ProgressBar1.Value = 0

            'hapus file hasil compress
            If My.Computer.FileSystem.FileExists(FilePathHasilCompress) = True Then
                My.Computer.FileSystem.DeleteFile(FilePathHasilCompress, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If

            MsgBox("Finish", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjZIP = Nothing
            ObjEncrypt = Nothing
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'decrypt: decrypt -> decompress
        Dim ObjZIP As New Core_Module.ClsCompressFile
        Dim ObjEncrypt As New Core_Module.ClsEncryptFile
        Dim FileName As String = ""
        Dim FolderNameHasilDeCompress As String = ""
        Dim FileNameHasilDecrypt As String = ""
        Dim bytKey As Byte()
        Dim bytIV As Byte()
        Dim vEncryptPass As String = Core_Module.GetEncryptPass_Installer

        Try
            'pilih file *.smess
            Me.OpenFileDialog1.Title = "Please specify file to decrypt"
            Me.OpenFileDialog1.Filter = "smess files (*.smess)|*.smess"
            Me.OpenFileDialog1.FilterIndex = 1
            Me.OpenFileDialog1.DefaultExt = "smess"
            Me.OpenFileDialog1.CheckFileExists = True
            Me.OpenFileDialog1.Multiselect = False
            Me.OpenFileDialog1.FileName = ""

            If Me.OpenFileDialog1.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If
            If Me.OpenFileDialog1.FileName = "" Then
                Exit Sub
            End If

            FileName = Me.OpenFileDialog1.FileName
            'decrypt 
            bytKey = ObjEncrypt.CreateKey(vEncryptPass)
            'Send the password to the CreateIV function.
            bytIV = ObjEncrypt.CreateIV(vEncryptPass)
            FileNameHasilDecrypt = FileName.Substring(0, FileName.Length - 6) + ".zip"
            ObjEncrypt.EncryptOrDecryptFile(FileName, FileNameHasilDecrypt, bytKey, bytIV, ClsEncryptFile.CryptoAction.ActionDecrypt, Me.ProgressBar1)
            Me.ProgressBar1.Value = 0
            'de-compress
            FolderNameHasilDeCompress = FileName.Substring(0, FileName.LastIndexOf("\"))
            ObjZIP.UnZipFolder(FileNameHasilDecrypt, Core_Module.GetEncryptPass_Installer2, FolderNameHasilDeCompress)
            'hapus file hasil decrypt
            If My.Computer.FileSystem.FileExists(FileNameHasilDecrypt) = True Then
                My.Computer.FileSystem.DeleteFile(FileNameHasilDecrypt, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
            MsgBox("Finish", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjZIP = Nothing
            ObjEncrypt = Nothing
        End Try
    End Sub
End Class