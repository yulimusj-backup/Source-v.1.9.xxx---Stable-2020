Option Explicit On
Option Strict Off

Public Class FrmMasterBackUpDB
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim baru As Boolean
    Dim objBackUpDB As ClsMasterBackUpDB
    Dim ObjCD As New ClsCekData
    Dim ObjLicense As New ClsLicense
    Dim buffer As String = ""
    Dim ObjUser As New Core_Module.ClsUser

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimary.BackColor, False)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None

            If Me.cboInterval.SelectedIndex = ClsMasterBackUpDB.EnumInterval.Daily Then
                Me.cboDayOfWeek.Enabled = False
            Else
                Me.cboDayOfWeek.Enabled = True
            End If
            Me.chlShutdown.Enabled = True
            Me.chkReminderClosing.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.Controls, True, True, True, True, True, True, True, True)
            Me.CancelButton = Me.btnExit
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.chlShutdown.Enabled = False
            Me.chkReminderClosing.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, False, False)
            Me.TxtPrimary.Text = "-1"
            Me.TxtPrimaryUserID.Text = "-1"
            Me.TxtBackUpTime.Text = "00.00"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.chlShutdown.Checked = False
            Me.chkReminderClosing.Checked = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection

        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("code")
            KoleksiTextBox.Add(Me.TxtCode)
            KoleksiNamaData.Add("description") '--> description, tidak boleh memiliki *, :
            KoleksiTextBox.Add(Me.TxtDesc)
            KoleksiNamaData.Add("user code")
            KoleksiTextBox.Add(Me.TxtUserID)
            KoleksiNamaData.Add("location 1")
            KoleksiTextBox.Add(Me.TxtLocation1)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            If Me.TxtDesc.Text.Contains(":") = True Or Me.TxtDesc.Text.Contains("?") = True Or Me.TxtDesc.Text.Contains("|") = True _
            Or Me.TxtDesc.Text.Contains("<") = True Or Me.TxtDesc.Text.Contains(">") = True Or Me.TxtDesc.Text.Contains("*") = True _
            Or Me.TxtDesc.Text.Contains("""") = True Then
                MsgBox("Invalid character, please avoid using: :, *, <, >, |, "", and ? for description.", MsgBoxStyle.Information, "Information")
                Me.TxtDesc.Focus()
                Return False
            End If

            'karena pengecekan dilakukan 1 jam 1x, maka minimum back up time adalah 02.15
            If Me.ObjCD.Convert_ToMinutes(Me.TxtBackUpTime.Text) < 135 Then
                MsgBox("Minimum back up time is: 02.15.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.objBackUpDB.ApakahSudahAda(Me.TxtCode.Text) = True Then
                    MsgBox("Duplication in code, please change your code", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtCode.Focus()
                    Return False
                End If
            Else
                Dim ProductGroupLama As String = Me.objBackUpDB.GetCode(Me.TxtPrimary.Text)
                If ProductGroupLama <> Me.TxtCode.Text Then
                    If Me.objBackUpDB.ApakahSudahAda(Me.TxtCode.Text) = True Then
                        MsgBox("Duplication in code, please change your code", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.objBackUpDB.GetTimeStamp(Me.TxtPrimary.Text)
                If Me.objBackUpDB.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.objBackUpDB.CariDetail(Me.objBackUpDB.Primary)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'pastikan location 1 dan location 2 ada foldernya
            Me.TxtLocation1.Text = Me.TxtLocation1.Text.Trim
            Me.TxtLocation2.Text = Me.TxtLocation2.Text.Trim
            If Me.TxtLocation2.Text <> "" And Me.TxtLocation1.Text = "" Then
                Me.TxtLocation1.Text = Me.TxtLocation2.Text
                Me.TxtLocation2.Text = ""
            End If
            Dim folderExists As Boolean
            folderExists = My.Computer.FileSystem.DirectoryExists(Me.TxtLocation1.Text)
            If folderExists = False Then
                MsgBox("Please specify valid folder.", MsgBoxStyle.Information, "Information")
                Me.TxtLocation1.Focus()
                Return False
            End If
            If Me.TxtLocation2.Text <> "" Then
                folderExists = My.Computer.FileSystem.DirectoryExists(Me.TxtLocation2.Text)
                If folderExists = False Then
                    MsgBox("Please specify valid folder.", MsgBoxStyle.Information, "Information")
                    Me.TxtLocation2.Focus()
                    Return False
                End If
                'location 1<> location 2
                If Me.TxtLocation1.Text = Me.TxtLocation2.Text Then
                    MsgBox("Please specify different back up location", MsgBoxStyle.Information, "Information")
                    Me.TxtLocation2.Focus()
                    Return False
                End If
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            With Me.objBackUpDB
                .Code = Me.TxtCode.Text
                .Description = Me.TxtDesc.Text
                .BackUpTime = Me.ObjCD.Convert_ToMinutes(Me.TxtBackUpTime.Text)
                .DayOfWeek = Me.cboDayOfWeek.SelectedIndex
                .Interval = Me.cboInterval.SelectedIndex
                .Location1 = Me.TxtLocation1.Text
                .Location2 = Me.TxtLocation2.Text
                .PrimaryUser = Me.TxtPrimaryUserID.Text
                If Me.chlShutdown.Checked = False Then
                    .ShutdownAfter = 0
                Else
                    .ShutdownAfter = 1
                End If
                If Me.chkReminderClosing.Checked = True Then
                    .RemindOnClose = 1
                Else
                    .RemindOnClose = 0
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.objBackUpDB
                Me.TxtCode.Text = .Code
                Me.TxtDesc.Text = .Description
                Me.TxtPrimary.Text = .Primary

                Me.TxtBackUpTime.Text = Me.ObjCD.Convert_ToString(.BackUpTime)
                Me.cboInterval.SelectedIndex = .Interval
                If .Interval = ClsMasterBackUpDB.EnumInterval.Weekly Then
                    Me.cboDayOfWeek.SelectedIndex = .DayOfWeek
                Else
                    Me.cboDayOfWeek.SelectedIndex = Me.cboDayOfWeek.Items.Count - 1
                End If
                Me.TxtLocation1.Text = .Location1
                Me.TxtLocation2.Text = .Location2
                Me.TxtPrimaryUserID.Text = .PrimaryUser
                If .ShutdownAfter = 1 Then
                    Me.chlShutdown.Checked = True
                Else
                    Me.chlShutdown.Checked = False
                End If
                Me.TxtUserID.Text = Me.ObjUser.GetUserCode(.PrimaryUser)
                If .RemindOnClose = 1 Then
                    Me.chkReminderClosing.Checked = True
                Else
                    Me.chkReminderClosing.Checked = False
                End If

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.objBackUpDB = New ClsMasterBackUpDB
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.cboInterval.SelectedIndex = 0
            Me.cboDayOfWeek.SelectedIndex = Me.cboDayOfWeek.Items.Count - 1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objBackUpDB.SimpanBaru(Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.objBackUpDB.SimpanPerubahan(Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Not Me.TxtCode.Text = "" And Me.TxtDesc.Text <> "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.KosongkanField()
            End If
            Me.NonAktifkanDetail()
            Me.btnNew.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.baru = True
            Me.AktifkanDetail()
            Me.KosongkanField()
            Me.TxtCode.Focus()
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmMasterBackUpDB_Cari

        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.Primary = -1 Then
                If Me.objBackUpDB.CariDetail(hj.Primary) = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        If Me.TxtCode.Text = "" Then
            Me.btnFind.PerformClick()
        End If

        If Me.TxtCode.Text <> "" Then
            Me.baru = False
            Me.AktifkanDetail()
            Me.TxtCode.Focus()
        End If
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtCode.Text <> "" Then
                If Me.objBackUpDB.ApakahBolehHapus(Me.TxtPrimary.Text) = False Then
                    MsgBox("You cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.objBackUpDB.Hapus(Me.TxtPrimary.Text) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtCode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objBackUpDB.SimpanBaru(Me.UserCode) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objBackUpDB.SimpanPerubahan(Me.UserCode) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            Else
                                e.Cancel = True
                            End If
                        End If
                    Else
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUserID_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUserID.GotFocus
        TxtUserID.BackColor = BackColorAll
    End Sub

    Private Sub TxtUserID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUserID.LostFocus
        TxtUserID.BackColor = Me.TxtPrimary.BackColor
    End Sub

    Private Sub TxtLocation1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocation1.GotFocus
        TxtLocation1.BackColor = BackColorAll
    End Sub

    Private Sub TxtLocation1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtLocation1.KeyDown
        Try
            If e.KeyValue = 13 Then
                FrmDia.SelectedPath = ""
                FrmDia.ShowDialog()
                If FrmDia.SelectedPath <> "" Then
                    Me.TxtLocation1.Text = FrmDia.SelectedPath
                    Me.TxtLocation2.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocation1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocation1.Leave
        Try
            Me.TxtLocation1.Text = Me.TxtLocation1.Text.Trim
            If Me.TxtLocation1.Text <> "" Then
                Dim folderExists As Boolean
                folderExists = My.Computer.FileSystem.DirectoryExists(Me.TxtLocation1.Text)
                If folderExists = False Then
                    FrmDia.SelectedPath = ""
                    FrmDia.ShowDialog()
                    If FrmDia.SelectedPath <> "" Then
                        Me.TxtLocation1.Text = FrmDia.SelectedPath
                    End If

                    folderExists = My.Computer.FileSystem.DirectoryExists(Me.TxtLocation1.Text)
                    If folderExists = False Then
                        Me.TxtLocation1.Clear()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocation1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocation1.LostFocus
        TxtLocation1.BackColor = Me.TxtPrimary.BackColor
    End Sub

    Private Sub TxtLocation2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocation2.GotFocus
        TxtLocation2.BackColor = BackColorAll
    End Sub

    Private Sub TxtLocation2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtLocation2.KeyDown
        Try
            If e.KeyValue = 13 Then
                FrmDia.SelectedPath = ""
                FrmDia.ShowDialog()
                If FrmDia.SelectedPath <> "" Then
                    Me.TxtLocation2.Text = FrmDia.SelectedPath
                    Me.cboInterval.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocation2_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocation2.Leave
        Try
            Me.TxtLocation2.Text = Me.TxtLocation2.Text.Trim
            If Me.TxtLocation2.Text <> "" Then
                Dim folderExists As Boolean
                folderExists = My.Computer.FileSystem.DirectoryExists(Me.TxtLocation2.Text)

                If folderExists = False Then
                    FrmDia.SelectedPath = ""
                    FrmDia.ShowDialog()
                    If FrmDia.SelectedPath <> "" Then
                        Me.TxtLocation2.Text = FrmDia.SelectedPath
                    End If
                End If

                folderExists = My.Computer.FileSystem.DirectoryExists(Me.TxtLocation2.Text)
                If folderExists = False Then
                    Me.TxtLocation2.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocation2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocation2.LostFocus
        TxtLocation2.BackColor = Me.TxtPrimary.BackColor
    End Sub

    Private Sub TxtCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCode.GotFocus
        TxtCode.BackColor = BackColorAll
    End Sub

    Private Sub TxtCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCode.LostFocus
        TxtCode.BackColor = Me.TxtPrimary.BackColor
    End Sub

    Private Sub TxtDesc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDesc.GotFocus
        TxtDesc.BackColor = BackColorAll
    End Sub

    Private Sub TxtDesc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDesc.LostFocus
        TxtDesc.BackColor = Me.TxtPrimary.BackColor
    End Sub

    Private Sub TxtBackUpTime_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBackUpTime.Enter
        Try
            Me.buffer = Me.TxtBackUpTime.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBackUpTime_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBackUpTime.GotFocus
        Me.TxtBackUpTime.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtBackUpTime_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBackUpTime.Leave
        Try
            If Me.ObjCD.Is_Time_String(Me.TxtBackUpTime.Text) = False Then
                Me.TxtBackUpTime.Text = Me.buffer
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBackUpTime_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBackUpTime.LostFocus
        Me.TxtBackUpTime.BackColor = Me.TxtPrimary.BackColor
    End Sub

    Private Sub TxtUserID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUserID.KeyDown
        Dim i As New FrmCariMasterUser
        Dim Obj As New ClsUser

        Try
            'i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Obj.ApakahSudahAda(Me.TxtUserID.Text) = True Then
                    Me.TxtPrimaryUserID.Text = Obj.GetPrimary(Me.TxtUserID.Text)
                    Me.TxtDesc.Focus()
                Else
                    If Me.TxtUserID.Text <> "" Then
                        i.TxtCari.Text = Me.TxtUserID.Text
                    End If

                    i.ShowDialog()
                    If i.Kodeuser <> "" Then
                        Me.TxtUserID.Text = i.Kodeuser
                        Me.TxtPrimaryUserID.Text = Obj.GetPrimary(i.Kodeuser)
                        Me.TxtDesc.Focus()
                    Else
                        Me.TxtUserID.Clear()
                        Me.TxtPrimaryUserID.Text = "-1"
                        Me.TxtUserID.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            Obj = Nothing
        End Try
    End Sub

    Private Sub TxtUserID_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUserID.Leave
        Dim i As New FrmCariMasterUser
        Dim obj As New ClsUser

        Try
            If Me.TxtUserID.Text <> "" Then
                If obj.ApakahSudahAda(Me.TxtUserID.Text) = True Then
                    Me.TxtPrimaryUserID.Text = obj.GetPrimary(Me.TxtUserID.Text)
                Else
                    If Me.TxtUserID.Text <> "" Then
                        i.TxtCari.Text = Me.TxtUserID.Text
                    End If

                    i.ShowDialog()
                    If i.Kodeuser <> "" Then
                        Me.TxtUserID.Text = i.Kodeuser
                        Me.TxtPrimaryUserID.Text = obj.GetPrimary(i.Kodeuser)
                    Else
                        Me.TxtUserID.Clear()
                        Me.TxtPrimaryUserID.Text = "-1"
                        Me.TxtUserID.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryUserID.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFindUnit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindUserID.Click
        Dim i As New FrmCariMasterUser
        Dim obj As New ClsUser

        Try
            'i.BackColorAll = Me.BackColorAll
            If Me.TxtUserID.Text <> "" Then
                i.TxtCari.Text = Me.TxtUserID.Text
            End If

            i.ShowDialog()
            If i.Kodeuser <> "" Then
                Me.TxtPrimaryUserID.Text = obj.GetPrimary(i.Kodeuser)
                Me.TxtUserID.Text = i.Kodeuser
                Me.TxtDesc.Focus()
            Else
                Me.TxtUserID.Clear()
                Me.TxtUserID.Focus()
                Me.TxtPrimaryUserID.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFindLocation1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindLocation1.Click
        Try
            FrmDia.SelectedPath = ""
            FrmDia.ShowDialog()
            If FrmDia.SelectedPath <> "" Then
                Me.TxtLocation1.Text = FrmDia.SelectedPath
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindLocation2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindLocation2.Click
        Try
            FrmDia.SelectedPath = ""
            FrmDia.ShowDialog()
            If FrmDia.SelectedPath <> "" Then
                Me.TxtLocation2.Text = FrmDia.SelectedPath
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetCboDayOfWeek()
        Try
            If Me.TxtCode.Enabled = True Then
                If Me.cboInterval.SelectedIndex = ClsMasterBackUpDB.EnumInterval.Daily Then
                    Me.cboDayOfWeek.Enabled = False
                    Me.cboDayOfWeek.SelectedIndex = Me.cboDayOfWeek.Items.Count - 1
                Else
                    Me.cboDayOfWeek.Enabled = True
                    Me.cboDayOfWeek.SelectedIndex = 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboInterval_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboInterval.SelectedIndexChanged
        Try
            Me.SetCboDayOfWeek()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboDayOfWeek_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDayOfWeek.Leave
        Try
            If Me.TxtCode.Enabled = True Then
                If Me.cboInterval.SelectedIndex = 1 And Me.cboDayOfWeek.SelectedIndex = Me.cboDayOfWeek.Items.Count - 1 Then
                    MsgBox("Please specify 1 valid day for weekly back up.", MsgBoxStyle.Information, "Information")
                    Me.cboDayOfWeek.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboDayOfWeek_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDayOfWeek.SelectedIndexChanged

    End Sub
End Class