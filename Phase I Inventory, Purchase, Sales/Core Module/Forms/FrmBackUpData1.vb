Public Class FrmBackUpData1
    Dim ObjDB As New ClsMySQL

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ObjBackUp As New Core_Module.ClsMySQL

        FrmDia.Title = "Please specify where to save back up file"
        If Me.chkCompress.Checked = False Then
            FrmDia.Filter = "*.sql|sql script"
            FrmDia.FilterIndex = 0
            FrmDia.DefaultExt = ".sql"
        Else
            FrmDia.Filter = "*.smess|smess back up"
            FrmDia.FilterIndex = 0
            FrmDia.DefaultExt = ".smess"
        End If
        FrmDia.OverwritePrompt = True
        If FrmDia.ShowDialog <> DialogResult.OK Then
            Exit Sub
        End If
        If FrmDia.FileName = "" Then
            Exit Sub
        End If

        If Me.chkExcludeMatdash.Checked = True Then
            ObjBackUp.BackUpDataBase(Me.GetConstring, FrmDia.FileName, Me.txthost.Text, Me.txtdatabase.Text, _
            Me.txtuserid.Text, Me.txtpassw.Text, Me.txtport.Text, "", "SMESS Auto Back Up", Me.chkCompress.Checked, Me.CheckBox1.Checked, Me.chkCompress.Checked, "", False)
        Else
            ObjBackUp.BackUpDataBase(Me.GetConstring, FrmDia.FileName, Me.txthost.Text, Me.txtdatabase.Text, _
            Me.txtuserid.Text, Me.txtpassw.Text, Me.txtport.Text, "", "SMESS Auto Back Up", Me.chkCompress.Checked, Me.CheckBox1.Checked, Me.chkCompress.Checked, "", True)
        End If
        ObjBackUp = Nothing
    End Sub

    Private Function GetConstring() As String
        Return "server=" + Me.txthost.Text + ";user id=" + Me.txtuserid.Text + ";Password=" + Me.txtpassw.Text + _
        ";persist security info=True;database=" + Me.txtdatabase.Text + ";port=" + Me.txtport.Text + ";default command timeout=19000000"
    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Me.OpenFileDialog1.Title = "Please specify database file to restore"
            If Me.chkCompress.Checked = False Then
                Me.OpenFileDialog1.Filter = "sql files (*.sql)|*.sql"
                Me.OpenFileDialog1.FilterIndex = 1
                Me.OpenFileDialog1.DefaultExt = "sql"
                Me.OpenFileDialog1.CheckFileExists = True
                Me.OpenFileDialog1.Multiselect = False
            Else
                Me.OpenFileDialog1.Filter = "smess files (*.smess)|*.smess"
                Me.OpenFileDialog1.FilterIndex = 1
                Me.OpenFileDialog1.DefaultExt = "smess"
                Me.OpenFileDialog1.CheckFileExists = True
                Me.OpenFileDialog1.Multiselect = False
            End If
            Me.OpenFileDialog1.FileName = ""

            If Me.OpenFileDialog1.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If
            If Me.OpenFileDialog1.FileName = "" Then
                Exit Sub
            End If

            Dim FrmRestore As New Core_Module.FrmBackUpData2
            FrmRestore.vFileName = Me.OpenFileDialog1.FileName
            FrmRestore.vConStr = Me.GetConstring
            FrmRestore.vCompress = Me.chkCompress.Checked
            If Me.chkExcludeMatdash.Checked = True Then
                FrmRestore.vRestoreMatDash = False
            Else
                FrmRestore.vRestoreMatDash = True
            End If

            FrmRestore.ShowDialog()

            FrmDia.FileName = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmBackUpData1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.txthost.Text = "localhost"
            Me.txtdatabase.Text = "core_db"
            Me.txtpassw.Text = "123"
            Me.txtport.Text = "3307"
            Me.txtuserid.Text = "root"
            Me.Text = "SMESS My SQL Back Up & Restore"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEncrypt.Click
        'encrypt: compress -> encrypt
        Dim ObjZIP As New Core_Module.ClsCompressFile
        Dim ObjEncrypt As New Core_Module.ClsEncryptFile
        Dim FileName As String = ""
        Dim FileNameHasilCompress As String = ""
        Dim FileNameHasilEncryt As String = ""
        Dim bytKey As Byte()
        Dim bytIV As Byte()
        Dim vEncryptPass As String = Core_Module.GetEncryptPass

        Try
            'pilih file *.sql
            Me.OpenFileDialog1.Title = "Please specify file to encrypt"
            Me.OpenFileDialog1.Filter = "sql files (*.sql)|*.sql"
            Me.OpenFileDialog1.FilterIndex = 1
            Me.OpenFileDialog1.DefaultExt = "sql"
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

    Private Sub btnDecrypt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDecrypt.Click
        'decrypt: decrypt -> decompress
        Dim ObjZIP As New Core_Module.ClsCompressFile
        Dim ObjEncrypt As New Core_Module.ClsEncryptFile
        Dim FileName As String = ""
        Dim FileNameHasilDeCompress As String = ""
        Dim FileNameHasilDecrypt As String = ""
        Dim bytKey As Byte()
        Dim bytIV As Byte()
        Dim vEncryptPass As String = Core_Module.GetEncryptPass

        Try
            'pilih file *.smess
            Me.OpenFileDialog1.Title = "Please specify file to decrypt"
            Me.OpenFileDialog1.Filter = "smess files (*.smess)|*.smess|all files (*.*)|*.*"
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
            FileNameHasilDeCompress = FileName.Substring(0, FileName.Length - 6) + ".sql"
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

    Private Sub btnRestructure_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRestructure.Click
        Try
            Me.OpenFileDialog1.Title = "Please specify database structure file to restore"
            If Me.chkCompress.Checked = False Then
                Me.OpenFileDialog1.Filter = "sql files (*.sql)|*.sql"
                Me.OpenFileDialog1.FilterIndex = 1
                Me.OpenFileDialog1.DefaultExt = "sql"
                Me.OpenFileDialog1.CheckFileExists = True
                Me.OpenFileDialog1.Multiselect = False
            Else
                Me.OpenFileDialog1.Filter = "smess files (*.smess)|*.smess"
                Me.OpenFileDialog1.FilterIndex = 1
                Me.OpenFileDialog1.DefaultExt = "smess"
                Me.OpenFileDialog1.CheckFileExists = True
                Me.OpenFileDialog1.Multiselect = False
            End If
            Me.OpenFileDialog1.FileName = ""

            If Me.OpenFileDialog1.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If
            If Me.OpenFileDialog1.FileName = "" Then
                Exit Sub
            End If

            Dim FrmRestore As New Core_Module.FrmBackUpData2
            FrmRestore.vFileName = Me.OpenFileDialog1.FileName
            FrmRestore.vConStr = Me.GetConstring
            FrmRestore.vCompress = Me.chkCompress.Checked
            FrmRestore.vUpdateData = False
            FrmRestore.vUpdateStructure = True
            FrmRestore.ShowDialog()

            FrmDia.FileName = ""

            If MsgBox("Load custom struture too?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.OpenFileDialog1.Title = "Please specify database custom structure file to restore"
                If Me.chkCompress.Checked = False Then
                    Me.OpenFileDialog1.Filter = "sql files (*.sql)|*.sql"
                    Me.OpenFileDialog1.FilterIndex = 1
                    Me.OpenFileDialog1.DefaultExt = "sql"
                    Me.OpenFileDialog1.CheckFileExists = True
                    Me.OpenFileDialog1.Multiselect = False
                Else
                    Me.OpenFileDialog1.Filter = "smess files (*.smess)|*.smess"
                    Me.OpenFileDialog1.FilterIndex = 1
                    Me.OpenFileDialog1.DefaultExt = "smess"
                    Me.OpenFileDialog1.CheckFileExists = True
                    Me.OpenFileDialog1.Multiselect = False
                End If
                Me.OpenFileDialog1.FileName = ""


                If Me.OpenFileDialog1.ShowDialog <> DialogResult.OK Then
                    Exit Sub
                End If
                If Me.OpenFileDialog1.FileName = "" Then
                    Exit Sub
                End If

                Dim FrmRestore3 As New Core_Module.FrmBackUpData2
                FrmRestore3.vFileName = Me.OpenFileDialog1.FileName
                FrmRestore3.vConStr = Me.GetConstring
                FrmRestore3.vCompress = Me.chkCompress.Checked
                FrmRestore3.vUpdateData = False
                FrmRestore3.vUpdateStructure = True
                FrmRestore3.ShowDialog()

                FrmDia.FileName = ""
            End If


            Me.OpenFileDialog1.Title = "Please specify database data file to restore"
            If Me.chkCompress.Checked = False Then
                Me.OpenFileDialog1.Filter = "sql files (*.sql)|*.sql"
                Me.OpenFileDialog1.FilterIndex = 1
                Me.OpenFileDialog1.DefaultExt = "sql"
                Me.OpenFileDialog1.CheckFileExists = True
                Me.OpenFileDialog1.Multiselect = False
            Else
                Me.OpenFileDialog1.Filter = "smess files (*.smess)|*.smess"
                Me.OpenFileDialog1.FilterIndex = 1
                Me.OpenFileDialog1.DefaultExt = "smess"
                Me.OpenFileDialog1.CheckFileExists = True
                Me.OpenFileDialog1.Multiselect = False
            End If
            Me.OpenFileDialog1.FileName = ""

            If Me.OpenFileDialog1.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If
            If Me.OpenFileDialog1.FileName = "" Then
                Exit Sub
            End If

            Dim FrmRestore2 As New Core_Module.FrmBackUpData2
            FrmRestore2.vFileName = Me.OpenFileDialog1.FileName
            FrmRestore2.vConStr = Me.GetConstring
            FrmRestore2.vCompress = Me.chkCompress.Checked
            FrmRestore2.vUpdateStructure = False
            FrmRestore2.vUpdateData = True
            If Me.chkExcludeMatdash.Checked = True Then
                FrmRestore2.vRestoreMatDash = False
            Else
                FrmRestore2.vRestoreMatDash = True
            End If
            FrmRestore2.ShowDialog()

            FrmDia.FileName = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class