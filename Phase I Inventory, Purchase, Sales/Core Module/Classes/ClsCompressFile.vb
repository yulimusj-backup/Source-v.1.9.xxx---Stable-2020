Option Explicit On
Option Strict Off

Imports System.IO
Imports System.IO.Compression
Imports ICSharpCode.SharpZipLib.Core
Imports ICSharpCode.SharpZipLib.Zip

Public Class ClsCompressFile
    Public Function UnZipFile(ByVal SourceZipPath As String, ByVal ResultFileName As String) As Boolean
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

    Public Sub ZipFile(ByVal source As String, ByVal dest As String)
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

    Public Sub ZipFolder(ByVal outPathname As String, ByVal password As String, ByVal folderName As String)

        Try
            Dim fsOut As FileStream = File.Create(outPathname)
            Dim zipStream As New ZipOutputStream(fsOut)

            zipStream.SetLevel(3)       '0-9, 9 being the highest level of compression
            zipStream.Password = password   ' optional. Null is the same as not setting.

            ' This setting will strip the leading part of the folder path in the entries, to
            ' make the entries relative to the starting folder.
            ' To include the full path for each entry up to the drive root, assign folderOffset = 0.
            Dim folderOffset As Integer = 1

            CompressFolder(folderName, zipStream, folderOffset)

            zipStream.IsStreamOwner = True
            ' Makes the Close also Close the underlying stream
            zipStream.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ' Recurses down the folder structure
    '
    Private Sub CompressFolder(ByVal path As String, ByVal zipStream As ZipOutputStream, ByVal folderOffset As Integer)
        Dim DummyFile As String = ""
        Try
            Dim files As String() = Directory.GetFiles(path)

            For Each filename As String In files
                DummyFile = filename
                Dim fi As New FileInfo(filename)

                Dim entryName As String = filename.Substring(folderOffset)  ' Makes the name in zip based on the folder
                entryName = ZipEntry.CleanName(entryName)       ' Removes drive from name and fixes slash direction
                Dim newEntry As New ZipEntry(entryName)
                newEntry.DateTime = fi.LastWriteTime            ' Note the zip format stores 2 second granularity

                ' Specifying the AESKeySize triggers AES encryption. Allowable values are 0 (off), 128 or 256.
                '   newEntry.AESKeySize = 256;

                ' To permit the zip to be unpacked by built-in extractor in WinXP and Server2003, WinZip 8, Java, and other older code,
                ' you need to do one of the following: Specify UseZip64.Off, or set the Size.
                ' If the file may be bigger than 4GB, or you do not need WinXP built-in compatibility, you do not need either,
                ' but the zip will be in Zip64 format which not all utilities can understand.
                '   zipStream.UseZip64 = UseZip64.Off;
                newEntry.Size = fi.Length

                zipStream.PutNextEntry(newEntry)

                ' Zip the file in buffered chunks
                ' the "using" will close the stream even if an exception occurs
                Dim buffer As Byte() = New Byte(4095) {}
                Using streamReader As FileStream = File.OpenRead(filename)
                    StreamUtils.Copy(streamReader, zipStream, buffer)
                End Using
                zipStream.CloseEntry()
            Next
            Dim folders As String() = Directory.GetDirectories(path)
            For Each folder As String In folders
                CompressFolder(folder, zipStream, folderOffset)
            Next
        Catch ex As Exception
            MsgBox(ex.Message + DummyFile, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub UnZipFolder(ByVal archiveFilenameIn As String, ByVal password As String, ByVal outFolder As String)
        Dim zf As ZipFile = Nothing

        Try
            Dim fs As FileStream = File.OpenRead(archiveFilenameIn)
            zf = New ZipFile(fs)
            If Not [String].IsNullOrEmpty(password) Then    ' AES encrypted entries are handled automatically
                zf.Password = password
            End If
            For Each zipEntry As ZipEntry In zf
                If Not zipEntry.IsFile Then     ' Ignore directories
                    Continue For
                End If
                Dim entryFileName As [String] = zipEntry.Name.Substring(zipEntry.Name.LastIndexOf("/") + 1)
                ' to remove the folder from the entry:- entryFileName = Path.GetFileName(entryFileName);
                ' Optionally match entrynames against a selection list here to skip as desired.
                ' The unpacked length is available in the zipEntry.Size property.

                Dim buffer As Byte() = New Byte(4095) {}    ' 4K is optimum
                Dim zipStream As Stream = zf.GetInputStream(zipEntry)

                ' Manipulate the output filename here as desired.
                Dim fullZipToPath As [String] = Path.Combine(outFolder, entryFileName)
                Dim directoryName As String = Path.GetDirectoryName(fullZipToPath)
                If directoryName.Length > 0 Then
                    Directory.CreateDirectory(directoryName)
                End If

                ' Unzip file in buffered chunks. This is just as fast as unpacking to a buffer the full size
                ' of the file, but does not waste memory.
                ' The "Using" will close the stream even if an exception occurs.
                Using streamWriter As FileStream = File.Create(fullZipToPath)
                    StreamUtils.Copy(zipStream, streamWriter, buffer)
                End Using
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            If zf IsNot Nothing Then
                zf.IsStreamOwner = True     ' Makes close also shut the underlying stream
                ' Ensure we release resources
                zf.Close()
            End If
        End Try
    End Sub
End Class
