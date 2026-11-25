Option Explicit On
Option Strict Off

Imports System.IO
Public Class FrmLicence_Add

    Public UserCode As String = ""
    Public BackColorAll As Color = Color.MistyRose
    Public IsBlackList As Integer = 0
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjLicense As New ClsLicense
    Dim Buffer As String = ""

    Private Sub FrmLicence_Add_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Me.IsBlackList = 1 Then
                Me.Text = "Req. License Blacklist"
                Me.ColReplaceFrom.Visible = True
                Me.ColAddLicence.Visible = False
                Me.LlbCheckAll.Visible = False
                Me.LlbUncheckAll.Visible = False
                Me.LblCount.Visible = False
                Me.ColIsBlackList.Visible = True
            Else
                Me.Text = "Req. License"
                Me.ColReplaceFrom.Visible = False
                Me.ColReplaceFromCompName.Visible = False
                Me.ColIsBlackList.Visible = False
            End If
            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvItem.AllowUserToAddRows = False
            Me.DgvItem.AllowUserToDeleteRows = False
            If Me.IsBlackList = 0 Then
                Me.DgvItem.AllowUserToDeleteRows = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnExport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnExport.Click
        Try

            If Me.IsBlackList = 0 Then
                Me.GetJmlLicense()
                If LblCount.Text = "Number of license : -" Then
                    MsgBox("please select an item first, minimum 1 item.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                Me.ExportToCSV()
            Else
                Me.ExportToCSV_Blacklist()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnLoad_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnLoad.Click
        Try
            If Me.IsBlackList = 0 Then
                Me.IsikanDGV()
            Else
                Me.IsikanDGV_Blacklist()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function IsikanDGV() As Boolean
        Dim DTA As New DseLicence_AddTableAdapters.tbltrans_version_controlTableAdapter
        Dim TBL As New DseLicence_Add.tbltrans_version_controlDataTable
        Dim Row As DseLicence_Add.tbltrans_version_controlRow
        Dim rowD As DataGridViewRow
        Dim ObjVer As New Clstrans_VersionControl
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Fill(TBL) > 0 Then
                Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                Me.DgvItem.Rows.Clear()
                For Each Row In TBL.Rows
                    If String.IsNullOrEmpty(Row.new_license_expiry) = False And Row.is_blacklist = 0 Then
                        rowD = Me.DgvItem.Rows(Me.DgvItem.Rows.Add)
                        With rowD
                            .Cells(Me.ColPrimary.Index).Value = Row.primary

                            .Cells(Me.ColComputerName.Index).Value = Row.computer_name
                            .Cells(Me.ColProductKey.Index).Value = Row.product_key
                            .Cells(Me.ColLicenceKey.Index).Value = ObjVer.GetDecrypt(Row.new_license_expiry)
                            .Cells(Me.ColExpiryDate.Index).Value = Format(Me.ObjLicense.GetExpiryDate(.Cells(Me.ColLicenceKey.Index).Value), "dd/MMM/yyyy")
                            .Cells(Me.ColLastUserLogin.Index).Value = Row.last_user_login
                            .Cells(Me.ColLastLoginTimestamp.Index).Value = Format(Row.last_login_timestamp, "dd/MMM/yyyy HH.mm")
                        End With
                    End If
                Next
            End If

            Me.ColExpiryDate.DefaultCellStyle.Format = "dd/MMM/yyyy"
            Me.ColLastLoginTimestamp.DefaultCellStyle.Format = "dd/MMM/yyyy HH.mm"
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True

            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function IsikanDGV_Blacklist() As Boolean
        Dim DTA As New DseLicence_AddTableAdapters.tbltrans_version_controlTableAdapter
        Dim TBL As New DseLicence_Add.tbltrans_version_controlDataTable
        Dim Row As DseLicence_Add.tbltrans_version_controlRow
        Dim rowD As DataGridViewRow
        Dim ObjVer As New Clstrans_VersionControl
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Fill(TBL) > 0 Then
                Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                Me.DgvItem.Rows.Clear()
                For Each Row In TBL.Rows
                    If String.IsNullOrEmpty(Row.new_license_expiry) = False Then
                        rowD = Me.DgvItem.Rows(Me.DgvItem.Rows.Add)
                        With rowD
                            .Cells(Me.ColPrimary.Index).Value = Row.primary

                            If Row.is_blacklist = 0 Then
                                .Cells(Me.ColIsBlackList.Index).Value = Boolean.FalseString
                            Else
                                .Cells(Me.ColIsBlackList.Index).Value = Boolean.TrueString
                            End If
                            .Cells(Me.ColComputerName.Index).Value = Row.computer_name
                            .Cells(Me.ColProductKey.Index).Value = Row.product_key
                            .Cells(Me.ColLicenceKey.Index).Value = ObjVer.GetDecrypt(Row.new_license_expiry)
                            .Cells(Me.ColExpiryDate.Index).Value = Format(Me.ObjLicense.GetExpiryDate(.Cells(Me.ColLicenceKey.Index).Value), "dd/MMM/yyyy")
                            .Cells(Me.ColLastUserLogin.Index).Value = Row.last_user_login
                            .Cells(Me.ColLastLoginTimestamp.Index).Value = Format(Row.last_login_timestamp, "dd/MMM/yyyy HH.mm")
                            .Cells(Me.ColReplaceFrom.Index).Value = Row.lencense_replace
                        End With
                    End If
                Next
            End If

            For Each rowD In Me.DgvItem.Rows
                With rowD
                    .Cells(Me.ColReplaceFromCompName.Index).Value = Me.GetCompNameReplace(rowD.Index, .Cells(Me.ColReplaceFrom.Index).Value)
                End With
            Next

            Me.ColExpiryDate.DefaultCellStyle.Format = "dd/MMM/yyyy"
            Me.ColLastLoginTimestamp.DefaultCellStyle.Format = "dd/MMM/yyyy HH.mm"
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True

            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ExportToCSV() As Boolean
        Dim StringKosong As String = ""
        Dim LicensiKeyEncrypt As String = ""
        Dim ObjVer As New Clstrans_VersionControl
        Dim ObjEncryptFile As New ClsEncryptFile
        Dim bytKey As Byte()
        Dim bytIV As Byte()
        Dim ExtentionLength As Integer = 0
        Dim ProgressBar1 As New ProgressBar
        Try
            bytKey = ObjEncryptFile.CreateKey("?qt$UG2#@&=eBHRF%Aw6VL$2L7^GRQ@D7%@EsRz=GM@yvyUM%@6=NhdX?@-#jVRAD5QFw8Va6g_WuS-wWX8eq7Wj$?hWhEFVfX5qe-d*Ud$p?DydxKY--VnrQAmb_X5PtwnJ@gHd6rj#m&!*tc-5SUmgsW7GtrqaD_Q$v%e*#q2uG%xzwe^Z#Zu-%h$jqx=H+by=K8?#VqH@T9FTFWze#5VHjkuX#F_ht@VA%!6nBn?^xcEJ@_2D@wZkMpvD*$du")
            'Send the password to the CreateIV function.
            bytIV = ObjEncryptFile.CreateIV("?qt$UG2#@&=eBHRF%Aw6VL$2L7^GRQ@D7%@EsRz=GM@yvyUM%@6=NhdX?@-#jVRAD5QFw8Va6g_WuS-wWX8eq7Wj$?hWhEFVfX5qe-d*Ud$p?DydxKY--VnrQAmb_X5PtwnJ@gHd6rj#m&!*tc-5SUmgsW7GtrqaD_Q$v%e*#q2uG%xzwe^Z#Zu-%h$jqx=H+by=K8?#VqH@T9FTFWze#5VHjkuX#F_ht@VA%!6nBn?^xcEJ@_2D@wZkMpvD*$du")

            Dim ValTxt As String = "COMPUTER_NAME;PRODUCT_KEY;OLD_LICENSE;NEW_LICENSE;LAST_LOGIN_USER;LAST_LOGIN_TIMESTAMP;IS_GENERATE"
            'ValTxt = ValTxt + vbCrLf + "OF;KODE_OBJEK;NAMA;HARGA_SATUAN;JUMLAH_BARANG;HARGA_TOTAL;DISKON;DPP;PPN;TARIF_PPNBM;PPNBM"
            If Me.DgvItem.Rows.Count > 0 Then
                If Me.SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    For Each rowD As DataGridViewRow In Me.DgvItem.Rows
                        With rowD
                            If String.IsNullOrEmpty(.Cells(Me.ColComputerName.Index).Value) = False And .Cells(Me.ColAddLicence.Index).Value = Boolean.TrueString Then
                                If String.IsNullOrEmpty(.Cells(Me.ColLicenceKey.Index).Value) = False Then
                                    LicensiKeyEncrypt = ObjVer.GetEncrypt(.Cells(Me.ColLicenceKey.Index).Value)
                                End If
                                ValTxt = ValTxt + vbNewLine + .Cells(Me.ColComputerName.Index).Value.ToString + ";" + .Cells(Me.ColProductKey.Index).Value.ToString + ";" + LicensiKeyEncrypt.ToString + ";" + _
                                                         StringKosong + ";" + .Cells(Me.ColLastUserLogin.Index).Value.ToString + ";" + _
                                                      .Cells(Me.ColLastLoginTimestamp.Index).Value + ";" + "0"
                            End If
                        End With
                    Next

                    My.Computer.FileSystem.WriteAllText(Me.SaveFileDialog1.FileName, ValTxt, False)
                    File.SetAttributes(Me.SaveFileDialog1.FileName, FileAttributes.Hidden)

                End If
            End If

            If String.IsNullOrEmpty(Me.SaveFileDialog1.FileName) = False Then
                ObjEncryptFile.EncryptOrDecryptFile(Me.SaveFileDialog1.FileName.Substring(0, Me.SaveFileDialog1.FileName.Length), Me.SaveFileDialog1.FileName.Substring(0, Me.SaveFileDialog1.FileName.Length - 4) + ".smess", _
              bytKey, bytIV, ClsEncryptFile.CryptoAction.ActionEncrypt, ProgressBar1)
                MsgBox("Finish created file", MsgBoxStyle.Information, "Information")

                If My.Computer.FileSystem.FileExists(Me.SaveFileDialog1.FileName.Substring(0, Me.SaveFileDialog1.FileName.Length - 6) + ".csv") = True Then
                    My.Computer.FileSystem.DeleteFile(Me.SaveFileDialog1.FileName.Substring(0, Me.SaveFileDialog1.FileName.Length - 6) + ".csv", FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Private Function ExportToCSV_Blacklist() As Boolean
    '    Dim StringKosong As String = ""
    '    Dim LicensiKeyEncrypt As String = ""
    '    Dim ObjVer As New Clstrans_VersionControl
    '    Dim isBlackList As String = "0"
    '    Dim ApakahAdaBlakclistDiHapus As Boolean = False
    '    Dim ListProductKey As New List(Of String)
    '    Try
    '        For Each rowD As DataGridViewRow In Me.DgvItem.Rows
    '            With rowD
    '                If String.IsNullOrEmpty(.Cells(Me.ColProductKey.Index).Value) = False Then
    '                    If ObjVer.ApakahKompIsBlacklist(.Cells(Me.ColProductKey.Index).Value) = True Then
    '                        If rowD.Cells(Me.ColIsBlackList.Index).Value = False Then
    '                            ApakahAdaBlakclistDiHapus = True
    '                            If ListProductKey.Contains(.Cells(Me.ColProductKey.Index).Value) = False Then
    '                                ListProductKey.Add(.Cells(Me.ColProductKey.Index).Value)
    '                            End If
    '                        End If
    '                    End If
    '                End If
    '            End With
    '        Next


    '        If ApakahAdaBlakclistDiHapus = True Then
    '            Dim ValTxt As String = "COMPUTER_NAME;PRODUCT_KEY;OLD_LICENSE;NEW_LICENSE;LAST_LOGIN_USER;LAST_LOGIN_TIMESTAMP;IS_BLACKLIST;TO;IS_GENERATE"
    '            'ValTxt = ValTxt + vbCrLf + "OF;KODE_OBJEK;NAMA;HARGA_SATUAN;JUMLAH_BARANG;HARGA_TOTAL;DISKON;DPP;PPN;TARIF_PPNBM;PPNBM"
    '            If Me.DgvItem.Rows.Count > 0 Then
    '                If Me.SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
    '                    For Each rowD As DataGridViewRow In Me.DgvItem.Rows
    '                        With rowD
    '                            If String.IsNullOrEmpty(.Cells(Me.ColComputerName.Index).Value) = False And ListProductKey.Contains(.Cells(Me.ColProductKey.Index).Value) = True Then
    '                                If String.IsNullOrEmpty(.Cells(Me.ColLicenceKey.Index).Value) = False Then
    '                                    LicensiKeyEncrypt = ObjVer.GetEncrypt(.Cells(Me.ColLicenceKey.Index).Value)
    '                                End If
    '                                ValTxt = ValTxt + vbNewLine + .Cells(Me.ColComputerName.Index).Value.ToString + ";" + .Cells(Me.ColProductKey.Index).Value.ToString + ";" + LicensiKeyEncrypt.ToString + ";" + _
    '                                                         StringKosong + ";" + .Cells(Me.ColLastUserLogin.Index).Value.ToString + ";" + _
    '                                                      .Cells(Me.ColLastLoginTimestamp.Index).Value + ";" + "0" + ";" + StringKosong + ";" + "0"
    '                            End If
    '                        End With
    '                    Next

    '                    My.Computer.FileSystem.WriteAllText(Me.SaveFileDialog1.FileName, ValTxt, False)
    '                    MsgBox("Csv file created.", MsgBoxStyle.Information, "Information")
    '                End If
    '            End If


    '        Else
    '            Dim ValTxt As String = "COMPUTER_NAME;PRODUCT_KEY;OLD_LICENSE;NEW_LICENSE;LAST_LOGIN_USER;LAST_LOGIN_TIMESTAMP;IS_BLACKLIST;TO;IS_GENERATE"
    '            'ValTxt = ValTxt + vbCrLf + "OF;KODE_OBJEK;NAMA;HARGA_SATUAN;JUMLAH_BARANG;HARGA_TOTAL;DISKON;DPP;PPN;TARIF_PPNBM;PPNBM"
    '            If Me.DgvItem.Rows.Count > 0 Then
    '                If Me.SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
    '                    For Each rowD As DataGridViewRow In Me.DgvItem.Rows
    '                        With rowD
    '                            If String.IsNullOrEmpty(.Cells(Me.ColComputerName.Index).Value) = False Then
    '                                If String.IsNullOrEmpty(.Cells(Me.ColLicenceKey.Index).Value) = False Then
    '                                    LicensiKeyEncrypt = ObjVer.GetEncrypt(.Cells(Me.ColLicenceKey.Index).Value)
    '                                End If
    '                                If rowD.Cells(Me.ColIsBlackList.Index).Value = True Then
    '                                    isBlackList = "1"
    '                                Else
    '                                    isBlackList = "0"
    '                                End If
    '                                ValTxt = ValTxt + vbNewLine + .Cells(Me.ColComputerName.Index).Value.ToString + ";" + .Cells(Me.ColProductKey.Index).Value.ToString + ";" + LicensiKeyEncrypt.ToString + ";" + _
    '                                                         StringKosong + ";" + .Cells(Me.ColLastUserLogin.Index).Value.ToString + ";" + _
    '                                                      .Cells(Me.ColLastLoginTimestamp.Index).Value + ";" + isBlackList + ";" + .Cells(Me.ColReplaceFrom.Index).Value.ToString + ";" + "0"
    '                            End If
    '                        End With
    '                    Next

    '                    My.Computer.FileSystem.WriteAllText(Me.SaveFileDialog1.FileName, ValTxt, False)
    '                    MsgBox("Csv file created.", MsgBoxStyle.Information, "Information")
    '                End If
    '            End If
    '        End If


    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    Private Function ExportToCSV_Blacklist() As Boolean
        Dim StringKosong As String = ""
        Dim LicensiKeyEncrypt As String = ""
        Dim ObjVer As New Clstrans_VersionControl
        Dim isBlackList As String = "0"
        Dim ApakahAdaBlakclistDiHapus As Boolean = False
        Dim ListProductKey As New List(Of String)
        Dim ObjEncryptFile As New ClsEncryptFile
        Dim bytKey As Byte()
        Dim bytIV As Byte()
        Dim ExtentionLength As Integer = 0
        Dim ProgressBar1 As New ProgressBar
        Try
            bytKey = ObjEncryptFile.CreateKey("?qt$UG2#@&=eBHRF%Aw6VL$2L7^GRQ@D7%@EsRz=GM@yvyUM%@6=NhdX?@-#jVRAD5QFw8Va6g_WuS-wWX8eq7Wj$?hWhEFVfX5qe-d*Ud$p?DydxKY--VnrQAmb_X5PtwnJ@gHd6rj#m&!*tc-5SUmgsW7GtrqaD_Q$v%e*#q2uG%xzwe^Z#Zu-%h$jqx=H+by=K8?#VqH@T9FTFWze#5VHjkuX#F_ht@VA%!6nBn?^xcEJ@_2D@wZkMpvD*$du")
            'Send the password to the CreateIV function.
            bytIV = ObjEncryptFile.CreateIV("?qt$UG2#@&=eBHRF%Aw6VL$2L7^GRQ@D7%@EsRz=GM@yvyUM%@6=NhdX?@-#jVRAD5QFw8Va6g_WuS-wWX8eq7Wj$?hWhEFVfX5qe-d*Ud$p?DydxKY--VnrQAmb_X5PtwnJ@gHd6rj#m&!*tc-5SUmgsW7GtrqaD_Q$v%e*#q2uG%xzwe^Z#Zu-%h$jqx=H+by=K8?#VqH@T9FTFWze#5VHjkuX#F_ht@VA%!6nBn?^xcEJ@_2D@wZkMpvD*$du")

            Dim ValTxt As String = "COMPUTER_NAME;PRODUCT_KEY;OLD_LICENSE;NEW_LICENSE;LAST_LOGIN_USER;LAST_LOGIN_TIMESTAMP;IS_BLACKLIST;FROM;IS_GENERATE"
            If Me.SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                For Each rowD As DataGridViewRow In Me.DgvItem.Rows
                    With rowD
                        If String.IsNullOrEmpty(.Cells(Me.ColProductKey.Index).Value) = False Then
                            If String.IsNullOrEmpty(.Cells(Me.ColLicenceKey.Index).Value) = False Then
                                LicensiKeyEncrypt = ObjVer.GetEncrypt(.Cells(Me.ColLicenceKey.Index).Value)
                            End If

                            If ObjVer.ApakahKompIsBlacklist(.Cells(Me.ColProductKey.Index).Value) = True Then
                                If rowD.Cells(Me.ColIsBlackList.Index).Value = False Then
                                    'jika masuk sini berarti ada perubahaan 
                                    'yaitu delete blacklist
                                    'dan cek apakah berubah jadi perpindah blacklist tidak
                                    
                                    ValTxt = ValTxt + vbNewLine + .Cells(Me.ColComputerName.Index).Value.ToString + ";" + .Cells(Me.ColProductKey.Index).Value.ToString + ";" + LicensiKeyEncrypt.ToString + ";" + _
                                                                                                StringKosong + ";" + .Cells(Me.ColLastUserLogin.Index).Value.ToString + ";" + _
                                                                                             .Cells(Me.ColLastLoginTimestamp.Index).Value + ";" + "0" + ";" + .Cells(Me.ColReplaceFrom.Index).Value + ";" + "0"

                                Else
                                    'masuk sini tidak ada perubahaan jadi abaikan
                                End If
                            Else
                                If rowD.Cells(Me.ColIsBlackList.Index).Value = False Then
                                    'masuk sini jika ada perpindahan blacklist replace from
                                    Dim RepLaceF As String = ObjVer.GetLicenseReplace(.Cells(Me.ColProductKey.Index).Value)
                                    If RepLaceF <> .Cells(Me.ColReplaceFrom.Index).Value Then
                                        'jika terjadi perubahaan 
                                        ValTxt = ValTxt + vbNewLine + .Cells(Me.ColComputerName.Index).Value.ToString + ";" + .Cells(Me.ColProductKey.Index).Value.ToString + ";" + LicensiKeyEncrypt.ToString + ";" + _
                                                                                                    StringKosong + ";" + .Cells(Me.ColLastUserLogin.Index).Value.ToString + ";" + _
                                                                                                 .Cells(Me.ColLastLoginTimestamp.Index).Value + ";" + "0" + ";" + .Cells(Me.ColReplaceFrom.Index).Value.ToString + ";" + "0"
                                    End If

                                Else
                                    'masuk sini jika da penambahan blacklist
                                    If rowD.Cells(Me.ColIsBlackList.Index).Value = True Then
                                        isBlackList = "1"
                                    Else
                                        isBlackList = "0"
                                    End If

                                    ValTxt = ValTxt + vbNewLine + .Cells(Me.ColComputerName.Index).Value.ToString + ";" + .Cells(Me.ColProductKey.Index).Value.ToString + ";" + LicensiKeyEncrypt.ToString + ";" + _
                                                             StringKosong + ";" + .Cells(Me.ColLastUserLogin.Index).Value.ToString + ";" + _
                                                          .Cells(Me.ColLastLoginTimestamp.Index).Value + ";" + isBlackList + ";" + .Cells(Me.ColReplaceFrom.Index).Value.ToString + ";" + "0"

                                End If
                            End If
                        End If
                    End With
                Next
                My.Computer.FileSystem.WriteAllText(Me.SaveFileDialog1.FileName, ValTxt, False)
                File.SetAttributes(Me.SaveFileDialog1.FileName, FileAttributes.Hidden)
            End If
            If String.IsNullOrEmpty(Me.SaveFileDialog1.FileName) = False Then
                ObjEncryptFile.EncryptOrDecryptFile(Me.SaveFileDialog1.FileName.Substring(0, Me.SaveFileDialog1.FileName.Length), Me.SaveFileDialog1.FileName.Substring(0, Me.SaveFileDialog1.FileName.Length - 4) + ".smess", _
                        bytKey, bytIV, ClsEncryptFile.CryptoAction.ActionEncrypt, ProgressBar1)
                MsgBox("Finish created file", MsgBoxStyle.Information, "Information")

                If My.Computer.FileSystem.FileExists(Me.SaveFileDialog1.FileName.Substring(0, Me.SaveFileDialog1.FileName.Length - 6) + ".csv") = True Then
                    My.Computer.FileSystem.DeleteFile(Me.SaveFileDialog1.FileName.Substring(0, Me.SaveFileDialog1.FileName.Length - 6) + ".csv", FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvItem.CellBeginEdit
        Try
            If Me.IsBlackList = 0 Then
                If e.ColumnIndex <> Me.ColAddLicence.Index Then
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColReplaceFrom.Index Then
                    If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColReplaceFrom.Index, e.RowIndex).Value) = False Then
                        ReplaceToLama = Me.DgvItem.Item(Me.ColReplaceFrom.Index, e.RowIndex).Value
                    End If

                    If Me.DgvItem.Item(Me.ColIsBlackList.Index, e.RowIndex).Value = True Then
                        e.Cancel = True
                    End If
                    'If Me.ApakahBlackList(e.RowIndex, Me.DgvItem.Item(Me.ColProductKey.Index, e.RowIndex).Value) = False Then
                    '    Buffer = Me.DgvItem.Item(Me.ColReplaceFrom.Index, e.RowIndex).Value
                    'Else
                    '    e.Cancel = True
                    'End If
                ElseIf e.ColumnIndex = Me.ColIsBlackList.Index Then
                    'If Me.DgvItem.Item(Me.ColIsBlackList.Index, e.RowIndex).Value = False Then
                    '    e.Cancel = True
                    'End If
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItem.CellEndEdit
        Dim ObjVer As New Clstrans_VersionControl
        Try
            If Me.IsBlackList = 0 Then
                If e.ColumnIndex = Me.ColAddLicence.Index Then
                    GetJmlLicense()
                End If
            Else

                If e.ColumnIndex = Me.ColReplaceFrom.Index Then
                    If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColReplaceFrom.Index, e.RowIndex).Value) = False Then
                        ' cek dulu kedb jika ia blacklist maka tidak bisa dipindahkan
                        If ObjVer.ApakahKompIsBlacklist(Me.DgvItem.Item(Me.ColReplaceFrom.Index, e.RowIndex).Value) = True Then
                            MsgBox("Please change product key this product key is blacklist.", MsgBoxStyle.Information, "Information")
                            Me.DgvItem.Item(Me.ColReplaceFrom.Index, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        If String.IsNullOrEmpty(ReplaceToLama) = False Then
                            If ReplaceToLama <> Me.DgvItem.Item(Me.ColReplaceFrom.Index, e.RowIndex).Value Then
                                Me.HapusBlacklist2(e.RowIndex, ReplaceToLama)
                            End If
                        End If
                        If Me.ApakahProductKeyBenar(e.RowIndex, Me.DgvItem.Item(Me.ColReplaceFrom.Index, e.RowIndex).Value) = False Then
                            MsgBox("Invalid data, Please change your Product Key license.", MsgBoxStyle.Information, "Information")
                            Me.DgvItem.Item(Me.ColReplaceFrom.Index, e.RowIndex).Value = Me.Buffer
                            Me.DgvItem.Item(Me.ColReplaceFromCompName.Index, e.RowIndex).Value = Me.Buffer
                            Me.HapusBlacklist2(e.RowIndex, ReplaceToLama)
                            Exit Sub
                        End If

                        If Me.ApakahBlackList(e.RowIndex, Me.DgvItem.Item(Me.ColReplaceFrom.Index, e.RowIndex).Value) = True Then
                            MsgBox("Duplication data license, please change your product key.", MsgBoxStyle.Information, "Information")
                            Me.DgvItem.Item(Me.ColReplaceFrom.Index, e.RowIndex).Value = Me.Buffer
                            Me.DgvItem.Item(Me.ColReplaceFromCompName.Index, e.RowIndex).Value = Me.Buffer
                            Me.HapusBlacklist2(e.RowIndex, ReplaceToLama)
                            Exit Sub
                        End If

                        'If Me.ApakahBlackList2(e.RowIndex, Me.DgvItem.Item(Me.ColReplaceFrom.Index, e.RowIndex).Value) = True Then
                        '    MsgBox("Duplication data license, please change your product key.", MsgBoxStyle.Information, "Information")
                        '    Me.DgvItem.Item(Me.ColReplaceFrom.Index, e.RowIndex).Value = Me.Buffer
                        '    Exit Sub
                        'End If
                        If Me.DgvItem.Item(Me.ColProductKey.Index, e.RowIndex).Value <> Me.DgvItem.Item(Me.ColReplaceFrom.Index, e.RowIndex).Value Then
                            Me.DgvItem.Item(Me.ColReplaceFromCompName.Index, e.RowIndex).Value = Me.GetCompNameReplace(e.RowIndex, Me.DgvItem.Item(Me.ColReplaceFrom.Index, e.RowIndex).Value)
                        Else
                            Me.DgvItem.Item(Me.ColReplaceFrom.Index, e.RowIndex).Value = Me.Buffer
                            Me.DgvItem.Item(Me.ColReplaceFromCompName.Index, e.RowIndex).Value = Me.Buffer
                            Me.HapusBlacklist2(e.RowIndex, ReplaceToLama)
                        End If
                    Else
                        Me.DgvItem.Item(Me.ColReplaceFromCompName.Index, e.RowIndex).Value = Me.Buffer
                        Me.HapusBlacklist2(e.RowIndex, ReplaceToLama)
                    End If

                ElseIf e.ColumnIndex = Me.ColIsBlackList.Index Then
                    If Me.DgvItem.Item(Me.ColIsBlackList.Index, e.RowIndex).Value = False Then
                        Me.HapusBlacklist(e.RowIndex, Me.DgvItem.Item(Me.ColProductKey.Index, e.RowIndex).Value)
                        Exit Sub
                    End If
                End If


            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub DeleteBlacklist()
    '    Try
    '        For Each row As DataGridViewRow In Me.DgvItem.Rows

    '        Next
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Function ApakahAdaReplaceTo(ByVal vProductkey As String) As Boolean
    '    Try
    '        For Each row As DataGridViewRow In Me.DgvItem.Rows

    '        Next
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    Private Sub GetJmlLicense()
        Try
            Dim Jml As Integer = 0
            For Each row As DataGridViewRow In Me.DgvItem.Rows
                If row.Cells(Me.ColAddLicence.Index).Value = Boolean.TrueString Then
                    Jml += 1
                End If
            Next
            If Jml <> 0 Then
                Me.LblCount.Text = "Number of license : " & Jml
            Else
                Me.LblCount.Text = "Number of license : -"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbCheckAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LlbCheckAll.Click
        Try
            For Each row As DataGridViewRow In Me.DgvItem.Rows
                row.Cells(Me.ColAddLicence.Index).Value = Boolean.TrueString
            Next
            Me.GetJmlLicense()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbUncheckAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LlbUncheckAll.Click
        Try
            For Each row As DataGridViewRow In Me.DgvItem.Rows
                row.Cells(Me.ColAddLicence.Index).Value = Boolean.FalseString
            Next
            Me.GetJmlLicense()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub DgvItem_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItem.CellContentClick

    End Sub

    Private Sub DgvItem_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItem.CellClick
        Try
            'GetJmlLicense()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetCompNameReplace(ByVal vRowIndex As Integer, ByVal vProductKey As String) As String

        Try
            For Each rowD As DataGridViewRow In Me.DgvItem.Rows
                If vRowIndex <> rowD.Index Then
                    If String.IsNullOrEmpty(vProductKey) = False Then
                        If vProductKey = rowD.Cells(Me.ColProductKey.Index).Value Then
                            rowD.Cells(Me.ColIsBlackList.Index).Value = Boolean.TrueString
                            Return rowD.Cells(Me.ColComputerName.Index).Value
                        End If
                    End If
                End If
            Next
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function HapusBlacklist(ByVal vRowIndex As Integer, ByVal vProductKey As String) As Boolean
        Try
            For Each rowD As DataGridViewRow In Me.DgvItem.Rows
                If vRowIndex <> rowD.Index Then
                    If String.IsNullOrEmpty(vProductKey) = False Then
                        If vProductKey = rowD.Cells(Me.ColReplaceFrom.Index).Value Then
                            rowD.Cells(Me.ColReplaceFrom.Index).Value = ""
                            rowD.Cells(Me.ColReplaceFromCompName.Index).Value = ""
                            Return True
                        End If
                    End If
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Dim ReplaceToLama As String = ""
    Private Function HapusBlacklist2(ByVal vRowIndex As Integer, ByVal vReplaceTo As String) As Boolean
        Try
            For Each rowD As DataGridViewRow In Me.DgvItem.Rows
                If vRowIndex <> rowD.Index Then
                    If String.IsNullOrEmpty(vReplaceTo) = False Then
                        If vReplaceTo = rowD.Cells(Me.ColProductKey.Index).Value Then
                            rowD.Cells(Me.ColIsBlackList.Index).Value = Boolean.FalseString
                            Return True
                        End If
                    End If
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahProductKeyBenar(ByVal vRowIndex As Integer, ByVal vProductKey As String) As Boolean
        Try
            For Each rowD As DataGridViewRow In Me.DgvItem.Rows
                If vRowIndex <> rowD.Index Then
                    If String.IsNullOrEmpty(vProductKey) = False Then
                        If vProductKey = rowD.Cells(Me.ColProductKey.Index).Value Then
                            Return True
                        End If
                    End If
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahBlackList(ByVal vRowIndex As Integer, ByVal vProductKey As String) As Boolean
        Try
            For Each rowD As DataGridViewRow In Me.DgvItem.Rows
                If vRowIndex <> rowD.Index Then
                    If String.IsNullOrEmpty(vProductKey) = False Then
                        If vProductKey = rowD.Cells(Me.ColReplaceFrom.Index).Value Then
                            Return True
                        End If
                    End If
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahBlackList2(ByVal vRowIndex As Integer, ByVal vProductKeyReplaceTo As String) As Boolean
        Try
            For Each rowD As DataGridViewRow In Me.DgvItem.Rows
                If vRowIndex <> rowD.Index Then
                    If String.IsNullOrEmpty(vProductKeyReplaceTo) = False Then
                        If vProductKeyReplaceTo = rowD.Cells(Me.ColProductKey.Index).Value And rowD.Cells(Me.ColIsBlackList.Index).Value = False Then
                            'If String.IsNullOrEmpty(rowD.Cells(Me.ColReplaceFrom.Index).Value) = False Then
                            Return True
                            'End If
                        End If
                    End If
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class