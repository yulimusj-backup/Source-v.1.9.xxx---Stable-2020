Option Explicit On
Option Strict Off

Imports System.IO

Public Class FrmLicence_Input
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.MistyRose
    Public IsBlackList As Integer = 0
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjLicense As New ClsLicense
    Dim Tbl As New DseLicense_Manager.DataTable2DataTable
    Dim Row As DseLicense_Manager.DataTable2Row
    Dim FileExtStr As String = ""
    Dim bytKey As Byte()
    Dim bytIV As Byte()
    Dim ObjEncryptFile As New ClsEncryptFile


    Private Sub FrmLicence_Input_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            bytKey = ObjEncryptFile.CreateKey("?qt$UG2#@&=eBHRF%Aw6VL$2L7^GRQ@D7%@EsRz=GM@yvyUM%@6=NhdX?@-#jVRAD5QFw8Va6g_WuS-wWX8eq7Wj$?hWhEFVfX5qe-d*Ud$p?DydxKY--VnrQAmb_X5PtwnJ@gHd6rj#m&!*tc-5SUmgsW7GtrqaD_Q$v%e*#q2uG%xzwe^Z#Zu-%h$jqx=H+by=K8?#VqH@T9FTFWze#5VHjkuX#F_ht@VA%!6nBn?^xcEJ@_2D@wZkMpvD*$du")
            bytIV = ObjEncryptFile.CreateIV("?qt$UG2#@&=eBHRF%Aw6VL$2L7^GRQ@D7%@EsRz=GM@yvyUM%@6=NhdX?@-#jVRAD5QFw8Va6g_WuS-wWX8eq7Wj$?hWhEFVfX5qe-d*Ud$p?DydxKY--VnrQAmb_X5PtwnJ@gHd6rj#m&!*tc-5SUmgsW7GtrqaD_Q$v%e*#q2uG%xzwe^Z#Zu-%h$jqx=H+by=K8?#VqH@T9FTFWze#5VHjkuX#F_ht@VA%!6nBn?^xcEJ@_2D@wZkMpvD*$du")

            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItem.AllowUserToAddRows = False
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True
            If Me.IsBlackList = 1 Then
                Me.ColIsBlackList.Visible = True
                'Me.ColAddLicence.Visible = False
                Me.Text &= " BlackList"
            Else
                'Me.ColAddLicence.Visible = True
                Me.ColIsBlackList.Visible = False
                Me.ColReplaceFrom.Visible = False
                Me.ColReplaceFromCompName.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function Openfile(ByRef vFileExtStr As String) As Boolean
        Dim FileExtStr As String = ""

        Try
            '1. cari file csv format infinity
            Me.OpenFileDialog1.Title = "Please specify file"
            Me.OpenFileDialog1.Filter = "smess files (*.smess)|*.smess"
            'Me.OpenFileDialog1.Filter = "Excel Worksheets (*.xlsx)|*.xlsx|excel files (*.xls)|*.xls|csv files (*.csv)|*.csv"
            Me.OpenFileDialog1.FilterIndex = 1
            Me.OpenFileDialog1.DefaultExt = "smess"
            Me.OpenFileDialog1.CheckFileExists = True
            Me.OpenFileDialog1.Multiselect = False
            Me.OpenFileDialog1.FileName = ""
            If Me.OpenFileDialog1.ShowDialog <> Windows.Forms.DialogResult.OK Then
                Return False
            End If
            If Me.OpenFileDialog1.FileName = "" Then
                Return False
            End If

            Me.txtFilePath.Text = Me.OpenFileDialog1.FileName
            vFileExtStr = System.IO.Path.GetExtension(Me.OpenFileDialog1.FileName)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BtnImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnImport.Click
        Try
            If Me.Openfile(FileExtStr) = False Then
                Me.txtFilePath.Text = ""
                Exit Sub
            End If

            If Me.IsBlackList = 0 Then
                If FileExtStr.StartsWith(".smess") Then
                    Me.ShowDgv()
                End If
            Else
                If FileExtStr.StartsWith(".smess") Then
                    Me.ShowDgv_Blacklist()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPath() As String
        Dim ObjEncryptFile As New ClsEncryptFile
        Dim ExtentionLength As Integer = 0
        Dim ProgressBar1 As New ProgressBar

        Try
            If String.IsNullOrEmpty(Me.OpenFileDialog1.FileName) = False Then
                ObjEncryptFile.EncryptOrDecryptFile(Me.OpenFileDialog1.FileName, Me.OpenFileDialog1.FileName.Substring(0, Me.OpenFileDialog1.FileName.Length - 6) + ".csv", _
                         bytKey, bytIV, ClsEncryptFile.CryptoAction.ActionDecrypt, ProgressBar1)
                File.SetAttributes(Me.OpenFileDialog1.FileName.Substring(0, Me.OpenFileDialog1.FileName.Length - 6) + ".csv", FileAttributes.Hidden)
                Return Me.OpenFileDialog1.FileName.Substring(0, Me.OpenFileDialog1.FileName.Length - 6) + ".csv"
            End If
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub ShowDgv()
        Dim SR As System.IO.StreamReader = New System.IO.StreamReader(Me.GetPath)
        Dim line As String = SR.ReadLine()
        line = line.Replace(";", ",")
        Dim iCount As String() = line.Split(","c)
        Dim ICount2 As String() = line.Split(";"c)
        Dim strArray As String()
        Dim j As New Clstrans_VersionControl
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.Tbl.Rows.Clear()
            If iCount.Length > ICount2.Length Then
                strArray = line.Split(","c)
            Else
                strArray = line.Split(";"c)
            End If


            If iCount.Length <> 7 Then
                MsgBox("Please change your file csv, this file is data blacklist.", MsgBoxStyle.Information, "Information")
                Tbl.Rows.Clear()
                Me.txtFilePath.Clear()
                Exit Sub
            End If

            Do
                line = SR.ReadLine
                If Not line = String.Empty Then
                    line = line.Replace(";", ",")
                    Dim SplitValue As String = ","
                    If iCount.Length > ICount2.Length Then
                        strArray = line.Split(","c)
                    Else
                        strArray = line.Split(";"c)
                    End If

                    'cek data csv harus bukan blacklist.csv
                    'If strArray(6) = "1" Then
                    '    'ini blacklist
                    '    MsgBox("Please change your file csv, this file is data blacklist.", MsgBoxStyle.Information, "Information")
                    '    Tbl.Rows.Clear()
                    '    Me.txtFilePath.Clear()
                    '    Exit Sub
                    'End If

                    If strArray(6) = "1" Then 'maksudnya yang sudah ada license baru 
                        Row = Tbl.NewRow
                        With Row
                            .computer_name = strArray(0)
                            .product_key = strArray(1)
                            .license_key = j.GetDecrypt(strArray(2))
                            .expiry_date = Me.GetExpiryDate(.license_key)
                            .new_license_key = j.GetDecrypt(strArray(3))
                            .new_expiry_date = Me.GetExpiryDate(.new_license_key)
                            .last_login_user = strArray(4)
                            .last_login_timestamp = strArray(5)
                            .is_generate = strArray(6)
                        End With
                        Tbl.Rows.Add(Row)
                    End If
                Else
                    Exit Do
                End If
            Loop

            If Tbl.Rows.Count > 0 Then
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                Me.dgvItem.Rows.Clear()
                For Each Me.Row In Me.Tbl.Rows
                    Dim RowD As DataGridViewRow
                    RowD = Me.dgvItem.Rows(Me.dgvItem.Rows.Add)
                    With RowD
                        .Cells(Me.ColComputerName.Index).Value = Row.computer_name
                        .Cells(Me.ColProductKey.Index).Value = Row.product_key
                        .Cells(Me.ColLicenceKey.Index).Value = Row.license_key
                        .Cells(Me.ColExpiryDate.Index).Value = Row.expiry_date
                        .Cells(Me.ColNewLicenceKey.Index).Value = Row.new_license_key
                        .Cells(Me.ColNewExpiryDate.Index).Value = Format(CDate(Row.new_expiry_date), "dd/MMM/yyyy")
                        .Cells(Me.ColLastLoginUser.Index).Value = Row.last_login_user
                        .Cells(Me.ColLastLoginTimestamp.Index).Value = Row.last_login_timestamp
                    End With
                Next
                Me.ColExpiryDate.DefaultCellStyle.Format = "dd/MMM/yyyy"
                Me.ColNewExpiryDate.DefaultCellStyle.Format = "dd/MMM/yyyy"
                Me.ColLastLoginTimestamp.DefaultCellStyle.Format = "dd/MMM/yyyy HH.mm"
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            End If

            SR.Close()
            If My.Computer.FileSystem.FileExists(Me.GetPath) = True Then
                My.Computer.FileSystem.DeleteFile(Me.GetPath, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnExecute_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnExecute.Click
        Try
            If IsBlackList = 0 Then
                Me.Execute()
            Else
                Me.Execute_BlackList()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function Execute() As Boolean
        Dim DTA As New Core_Module.DseTrans_Version_ControlTableAdapters.tbltrans_version_controlTableAdapter
        Dim ObjVer As New Clstrans_VersionControl
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            For Each rowD As DataGridViewRow In Me.dgvItem.Rows
                With rowD
                    If String.IsNullOrEmpty(.Cells(Me.ColProductKey.Index).Value) = False And String.IsNullOrEmpty(.Cells(Me.ColNewLicenceKey.Index).Value) = False Then
                        ObjVer.UpdateNewLicenseKey(.Cells(Me.ColNewLicenceKey.Index).Value, .Cells(Me.ColProductKey.Index).Value)
                        ObjVer.UpdateOldLicenseKey(.Cells(Me.ColNewExpiryDate.Index).Value, .Cells(Me.ColProductKey.Index).Value)
                        ObjVer.UpdateBlacklist("0", .Cells(Me.ColProductKey.Index).Value)
                    End If
                End With
            Next
            MsgBox("Finish", MsgBoxStyle.Information, "Information")
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetExpiryDate(ByVal vLicenseKey As String) As Date
        Dim d1, d2, m1, m2, y1, y2, y3, y4, MonthStr As String
        Dim a As Date = Today

        Try
            d1 = vLicenseKey.Substring(24, 1)
            d2 = vLicenseKey.Substring(0, 1)
            m1 = vLicenseKey.Substring(6, 1)
            m2 = vLicenseKey.Substring(12, 1)
            y1 = vLicenseKey.Substring(9, 1)
            y2 = vLicenseKey.Substring(21, 1)
            y3 = vLicenseKey.Substring(15, 1)
            y4 = vLicenseKey.Substring(18, 1)

            If m1 + m2 = "01" Then
                MonthStr = "Jan"
            ElseIf m1 + m2 = "02" Then
                MonthStr = "Feb"
            ElseIf m1 + m2 = "03" Then
                MonthStr = "Mar"
            ElseIf m1 + m2 = "04" Then
                MonthStr = "Apr"
            ElseIf m1 + m2 = "05" Then
                MonthStr = "May"
            ElseIf m1 + m2 = "06" Then
                MonthStr = "Jun"
            ElseIf m1 + m2 = "07" Then
                MonthStr = "Jul"
            ElseIf m1 + m2 = "08" Then
                MonthStr = "Aug"
            ElseIf m1 + m2 = "09" Then
                MonthStr = "Sep"
            ElseIf m1 + m2 = "10" Then
                MonthStr = "Oct"
            ElseIf m1 + m2 = "11" Then
                MonthStr = "Nov"
            Else
                MonthStr = "Dec"
            End If

            Return CDate(d1 + d2 + " " + MonthStr + " " + y1 + y2 + y3 + y4)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'khusus blacklist
    Private Function Execute_BlackList() As Boolean
        Dim DTA As New Core_Module.DseTrans_Version_ControlTableAdapters.tbltrans_version_controlTableAdapter
        Dim ObjVer As New Clstrans_VersionControl
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            For Each rowD As DataGridViewRow In Me.dgvItem.Rows
                With rowD
                    Dim IsBlacklist As Integer = 0

                    If String.IsNullOrEmpty(.Cells(Me.ColProductKey.Index).Value) = False Then
                        If String.IsNullOrEmpty(.Cells(Me.ColNewLicenceKey.Index).Value) = False Then
                            ObjVer.UpdateNewLicenseKey(.Cells(Me.ColNewLicenceKey.Index).Value, .Cells(Me.ColProductKey.Index).Value)
                            ObjVer.UpdateOldLicenseKey(.Cells(Me.ColNewExpiryDate.Index).Value, .Cells(Me.ColProductKey.Index).Value)
                            ObjVer.UpdateLicenseReplace(.Cells(Me.ColReplaceFrom.Index).Value, .Cells(Me.ColProductKey.Index).Value)
                            If .Cells(Me.ColIsBlackList.Index).Value = True Then
                                ObjVer.UpdateBlacklist("1", .Cells(Me.ColProductKey.Index).Value)
                            Else
                                ObjVer.UpdateBlacklist("0", .Cells(Me.ColProductKey.Index).Value)
                            End If
                        Else
                            If String.IsNullOrEmpty(.Cells(Me.ColReplaceFrom.Index).Value) = True Then
                                If .Cells(Me.ColIsBlackList.Index).Value = True Then
                                    ObjVer.UpdateBlacklist("1", .Cells(Me.ColProductKey.Index).Value)
                                Else
                                    ObjVer.UpdateBlacklist("0", .Cells(Me.ColProductKey.Index).Value)
                                End If
                            End If
                        End If
                    End If
                End With
            Next
            MsgBox("Finish", MsgBoxStyle.Information, "Information")
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub ShowDgv_Blacklist()
        Dim SR As System.IO.StreamReader = New System.IO.StreamReader(Me.GetPath)
        Dim line As String = SR.ReadLine()
        line = line.Replace(";", ",")
        Dim iCount As String() = line.Split(","c)
        Dim ICount2 As String() = line.Split(";"c)
        Dim strArray As String()
        Dim j As New Clstrans_VersionControl
        Dim ApakahFileCsvBenar As Boolean = False
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.Tbl.Rows.Clear()
            If iCount.Length > ICount2.Length Then
                strArray = line.Split(","c)
            Else
                strArray = line.Split(";"c)
            End If

            If iCount.Length <> 9 Then
                MsgBox("Please change your file csv, this file cannot data blacklist.", MsgBoxStyle.Information, "Information")
                Tbl.Rows.Clear()
                Me.txtFilePath.Clear()
                Exit Sub
            End If

            Do
                line = SR.ReadLine
                If Not line = String.Empty Then
                    line = line.Replace(";", ",")
                    Dim SplitValue As String = ","
                    If iCount.Length > ICount2.Length Then
                        strArray = line.Split(","c)
                    Else
                        strArray = line.Split(";"c)
                    End If

                    'If strArray(8) = "1" Then 'maksudnya yang sudah ada license baru 
                    Row = Tbl.NewRow
                    With Row
                        .computer_name = strArray(0)
                        .product_key = strArray(1)
                        .license_key = j.GetDecrypt(strArray(2))
                        .expiry_date = Me.GetExpiryDate(.license_key)
                        If String.IsNullOrEmpty(strArray(3)) = False Then
                            .new_license_key = j.GetDecrypt(strArray(3))
                            .new_expiry_date = Me.GetExpiryDate(.new_license_key)
                        Else
                            .new_license_key = ""
                            .new_expiry_date = ""
                        End If
                        .last_login_user = strArray(4)
                        .last_login_timestamp = strArray(5)
                        .is_blacklist = strArray(6)
                        .replace_to = strArray(7)
                        .is_generate = strArray(8)
                    End With
                    Tbl.Rows.Add(Row)
                    'End If
                Else
                Exit Do
                End If
            Loop

            If Tbl.Rows.Count > 0 Then
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                Me.dgvItem.Rows.Clear()
                For Each Me.Row In Me.Tbl.Rows
                    Dim RowD As DataGridViewRow
                    RowD = Me.dgvItem.Rows(Me.dgvItem.Rows.Add)
                    With RowD
                        .Cells(Me.ColComputerName.Index).Value = Row.computer_name
                        If Row.is_blacklist = "1" Then
                            .Cells(Me.ColIsBlackList.Index).Value = Boolean.TrueString
                        Else
                            .Cells(Me.ColIsBlackList.Index).Value = Boolean.FalseString
                        End If
                        .Cells(Me.ColProductKey.Index).Value = Row.product_key
                        .Cells(Me.ColLicenceKey.Index).Value = Row.license_key
                        .Cells(Me.ColExpiryDate.Index).Value = Row.expiry_date
                        .Cells(Me.ColNewLicenceKey.Index).Value = Row.new_license_key
                        .Cells(Me.ColNewExpiryDate.Index).Value = Row.new_expiry_date
                        .Cells(Me.ColLastLoginUser.Index).Value = Row.last_login_user
                        .Cells(Me.ColLastLoginTimestamp.Index).Value = Row.last_login_timestamp
                        .Cells(Me.ColReplaceFrom.Index).Value = Row.replace_to
                    End With
                Next

                For Each RowD As DataGridViewRow In Me.dgvItem.Rows
                    With RowD
                        .Cells(Me.ColReplaceFromCompName.Index).Value = Me.GetCompNameReplace(RowD.Index, .Cells(Me.ColReplaceFrom.Index).Value)
                    End With
                Next

                Me.ColExpiryDate.DefaultCellStyle.Format = "dd/MMM/yyyy"
                Me.ColLastLoginTimestamp.DefaultCellStyle.Format = "dd/MMM/yyyy HH.mm"
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            End If
            SR.Close()
            If My.Computer.FileSystem.FileExists(Me.GetPath) = True Then
                My.Computer.FileSystem.DeleteFile(Me.GetPath, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
            Windows.Forms.Cursor.Current = Cursors.Default
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
            Return ""
        End Try
    End Function
End Class