Option Explicit On
Option Strict Off

Imports System.IO
Imports System.IO.Compression
Imports ICSharpCode.SharpZipLib.Core
Imports ICSharpCode.SharpZipLib.Zip

Public Class ClsVersionControl
    Dim vEncryptPass As String = "*#%Simple83_999%$***##@@"

    'Public Function GetVersion_ClientUpdate(Optional ByVal vDefaultVersion As String = "") As String
    '    Dim RetVal As String = ""
    '    'Dim LastVersionTxt As String = "0.0.00"
    '    Dim fileExists As Boolean
    '    Dim ObjEncrypt As New ClsEncryptFile

    '    Try
    '        Windows.Forms.Cursor.Current = Cursors.WaitCursor
    '        'cek apakah file exist
    '        fileExists = My.Computer.FileSystem.FileExists(Environment.GetFolderPath(Environment.SpecialFolder.Personal) & "\" & "Client_Update_Version_Current.smess")
    '        If fileExists = False Then
    '            'MsgBox("Returning Default version")
    '            Return vDefaultVersion
    '        End If

    '        'decrypt file
    '        System.IO.File.SetAttributes(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update_Version_Current.smess", FileAttributes.Normal)
    '        Me.DecryptFile(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update_Version_Current.smess")

    '        'get version
    '        RetVal = My.Computer.FileSystem.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update_Version_Current.txt")

    '        'hapus file sementara
    '        System.IO.File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update_Version_Current.txt")
    '        System.IO.File.SetAttributes(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update_Version_Current.smess", FileAttributes.Hidden)

    '        Return RetVal

    '        Windows.Forms.Cursor.Current = Cursors.Arrow
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return vDefaultVersion
    '    End Try
    'End Function

    'Public Sub SetVersion_ClientUpdate(ByVal vNewVersion As String)
    '    Dim RetVal As String = ""
    '    'Dim LastVersionTxt As String = "0.0.00"
    '    Dim fileExists As Boolean
    '    Dim ObjEncrypt As New ClsEncryptFile

    '    Try
    '        'MsgBox("Set version")
    '        Windows.Forms.Cursor.Current = Cursors.WaitCursor
    '        'cek apakah file exist
    '        fileExists = My.Computer.FileSystem.FileExists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update_Version_Current.smess")
    '        If fileExists = True Then
    '            'hapus versi lama
    '            System.IO.File.SetAttributes(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update_Version_Current.smess", FileAttributes.Normal)
    '            My.Computer.FileSystem.DeleteFile(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update_Version_Current.smess", _
    '            FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
    '        End If

    '        'update ke versi baru
    '        My.Computer.FileSystem.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update_Version_Current.txt", vNewVersion, False)
    '        'encrypt
    '        Me.EncryptFileVersion(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update_Version_Current.txt")
    '        'hapus file txt
    '        My.Computer.FileSystem.DeleteFile(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update_Version_Current.txt", _
    '        FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
    '        System.IO.File.SetAttributes(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "Client_Update_Version_Current.smess", FileAttributes.Hidden)
    '        Windows.Forms.Cursor.Current = Cursors.Arrow
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Public Function GetVersion(Optional ByVal vDefaultVersion As String = "") As String
        Dim RetVal As String = ""
        'Dim LastVersionTxt As String = "0.0.00"
        Dim fileExists As Boolean
        Dim ObjEncrypt As New ClsEncryptFile
        Dim ObjIdentifier As New CustomSMESS.ClsIdentifier

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'cek apakah file exist
            fileExists = My.Computer.FileSystem.FileExists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current" + "_" + ObjIdentifier.GetIdentifier + ".smess")
            If fileExists = False Then
                'MsgBox("Returning Default version")
                Return vDefaultVersion
            End If

            'decrypt file
            System.IO.File.SetAttributes(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current" + "_" + ObjIdentifier.GetIdentifier + ".smess", FileAttributes.Normal)
            Me.DecryptFile(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current" + "_" + ObjIdentifier.GetIdentifier + ".smess")
            'MsgBox("Checking for: " + Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current.txt")
            'If My.Computer.FileSystem.FileExists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current.txt") = False Then
            '    MsgBox(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current.txt doesnot exits.", MsgBoxStyle.Information, "Information")
            'End If
            'get version
            RetVal = My.Computer.FileSystem.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current" + "_" + ObjIdentifier.GetIdentifier + ".txt")
            'hapus file sementara
            System.IO.File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current" + "_" + ObjIdentifier.GetIdentifier + ".txt")
            System.IO.File.SetAttributes(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current" + "_" + ObjIdentifier.GetIdentifier + ".smess", FileAttributes.Hidden)

            Return RetVal

            Windows.Forms.Cursor.Current = Cursors.Arrow
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return vDefaultVersion
        Finally
            ObjIdentifier = Nothing
        End Try
    End Function

    Public Sub SetVersion(ByVal vNewVersion As String)
        Dim RetVal As String = ""
        'Dim LastVersionTxt As String = "0.0.00"
        Dim fileExists As Boolean
        Dim ObjEncrypt As New ClsEncryptFile
        Dim ObjIdentifier As New CustomSMESS.ClsIdentifier

        Try
            'MsgBox("Set version")
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'cek apakah file exist
            fileExists = My.Computer.FileSystem.FileExists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current" + "_" + ObjIdentifier.GetIdentifier + ".smess")
            If fileExists = True Then
                'hapus versi lama
                System.IO.File.SetAttributes(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current" + "_" + ObjIdentifier.GetIdentifier + ".smess", FileAttributes.Normal)
                My.Computer.FileSystem.DeleteFile(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current" + "_" + ObjIdentifier.GetIdentifier + ".smess", _
                FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If

            'update ke versi baru
            My.Computer.FileSystem.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current" + "_" + ObjIdentifier.GetIdentifier + ".txt", vNewVersion, False)
            'encrypt
            Me.EncryptFileVersion(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current" + "_" + ObjIdentifier.GetIdentifier + ".txt")
            'hapus file txt
            My.Computer.FileSystem.DeleteFile(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current" + "_" + ObjIdentifier.GetIdentifier + ".txt", _
            FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            System.IO.File.SetAttributes(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current" + "_" + ObjIdentifier.GetIdentifier + ".smess", FileAttributes.Hidden)
            Windows.Forms.Cursor.Current = Cursors.Arrow

            'MsgBox("Update ke my doc: " + vNewVersion, MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjIdentifier = Nothing
            ObjEncrypt = Nothing
        End Try
    End Sub

    Private Sub EncryptFileVersion(ByVal FileName As String)
        'Dim ObjZIP As New Core_Module.ClsCompressFile
        Dim ObjEncrypt As New ClsEncryptFile
        Dim FileNameHasilCompress As String = ""
        Dim FileNameHasilEncryt As String = ""
        Dim bytKey As Byte()
        Dim bytIV As Byte()
        'Dim vEncryptPass As String = "*#%Simple83_999%$***##@@" 'Core_Module.GetEncryptPass_Installer
        Dim prg As New Windows.Forms.ProgressBar

        Try
            'compress
            FileNameHasilCompress = FileName.Substring(0, FileName.Length - 4) + ".zip"
            Me.ZipFile(FileName, FileNameHasilCompress)
            'encrypt + rename
            bytKey = ObjEncrypt.CreateKey(vEncryptPass)
            'Send the password to the CreateIV function.
            bytIV = ObjEncrypt.CreateIV(vEncryptPass)
            FileNameHasilEncryt = FileName.Substring(0, FileName.Length - 4) + ".smess"
            ObjEncrypt.EncryptOrDecryptFile(FileNameHasilCompress, FileNameHasilEncryt, bytKey, bytIV, ClsEncryptFile.CryptoAction.ActionEncrypt, prg)

            'hapus file hasil compress
            If My.Computer.FileSystem.FileExists(FileNameHasilCompress) = True Then
                My.Computer.FileSystem.DeleteFile(FileNameHasilCompress, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjEncrypt = Nothing
            prg = Nothing
        End Try
    End Sub

    Private Sub DecryptFile(ByVal FileName As String)
        'decrypt: decrypt -> decompress
        'Dim ObjZIP As New Core_Module.ClsCompressFile
        Dim ObjEncrypt As New ClsEncryptFile
        Dim FileNameHasilDeCompress As String = ""
        Dim FileNameHasilDecrypt As String = ""
        Dim bytKey As Byte()
        Dim bytIV As Byte()
        'Dim vEncryptPass As String = Core_Module.GetEncryptPass_Installer
        Dim prg As New Windows.Forms.ProgressBar

        Try
            'decrypt 
            bytKey = ObjEncrypt.CreateKey(vEncryptPass)
            'Send the password to the CreateIV function.
            bytIV = ObjEncrypt.CreateIV(vEncryptPass)
            FileNameHasilDecrypt = FileName.Substring(0, FileName.Length - 6) + ".zip"
            ObjEncrypt.EncryptOrDecryptFile(FileName, FileNameHasilDecrypt, bytKey, bytIV, ClsEncryptFile.CryptoAction.ActionDecrypt, prg)
            prg.Value = 0
            'de-compress
            FileNameHasilDeCompress = FileName.Substring(0, FileName.Length - 6) + ".txt"
            Me.UnZipFile(FileNameHasilDecrypt, FileNameHasilDeCompress)
            'hapus file hasil compress
            If My.Computer.FileSystem.FileExists(FileNameHasilDecrypt) = True Then
                My.Computer.FileSystem.DeleteFile(FileNameHasilDecrypt, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
            'MsgBox("Finish", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'ObjZIP = Nothing
            ObjEncrypt = Nothing
            prg = Nothing
        End Try
    End Sub

    Private Function UnZipFile(ByVal SourceZipPath As String, ByVal ResultFileName As String) As Boolean
        Try
            Dim fileStreamIn As FileStream = New FileStream(SourceZipPath, FileMode.Open, FileAccess.Read)
            Dim fileStreamOut As FileStream
            Dim zf As ZipFile = New ZipFile(fileStreamIn)
            Dim Size As Integer
            Dim buffer(4096) As Byte
            zf.Password = ""
            Dim Zentry As ZipEntry = zf.EntryByIndex(0)

            If (Zentry Is Nothing) Then
                Debug.Print("not found in Zip")
                Return False
                Exit Function
            End If

            Dim fstr As ICSharpCode.SharpZipLib.Zip.Compression.Streams.InflaterInputStream
            fstr = zf.GetInputStream(Zentry)

            fileStreamOut = New FileStream(ResultFileName, FileMode.Create, FileAccess.Write)

            Do
                Size = fstr.Read(buffer, 0, buffer.Length)
                fileStreamOut.Write(buffer, 0, Size)
            Loop While (Size > 0)

            fstr.Close()
            fileStreamOut.Close()
            fileStreamIn.Close()
            zf = Nothing

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub ZipFile(ByVal source As String, ByVal dest As String)
        Try
            Dim entry As New ICSharpCode.SharpZipLib.Zip.ZipEntry(Path.GetFileName(source))
            Dim fi As New FileInfo(source)
            entry.ExternalFileAttributes = fi.Attributes
            entry.Size = fi.Length

            Dim input As FileStream = File.OpenRead(source)
            Dim output As New ICSharpCode.SharpZipLib.Zip.ZipOutputStream(File.Create(dest))

            output.PutNextEntry(entry)

            Dim buffer(8191) As Byte
            Dim len As Integer
            Do
                len = input.Read(buffer, 0, buffer.Length)
                If len > 0 Then
                    output.Write(buffer, 0, len)
                End If
            Loop Until len = 0

            output.Close()
            input.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
