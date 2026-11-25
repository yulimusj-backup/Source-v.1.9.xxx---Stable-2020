Option Explicit On
Option Strict Off

Imports Core_Module, System.Drawing, System.Windows.Forms
Public Class FrmMasterPayrollAccess
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim TBL_SF As New DseMasterPayrollAkses.tblmaster_salary_formula1DataTable
    Dim TBL_User As New DseMasterPayrollAkses.tbluserDataTable

    Dim Obj As New ClsMasterPayrollAccess
    Dim ObjItem As New ClsMasterPayrollAccess_Item
    Dim ObjCD As New Core_Module.ClsCekData
    Dim Baru As Boolean = False
    Dim Buffer As String = ""
    Dim Load_1 As Boolean = False

    Private Sub FrmTransaksiTicketing_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Obj = New ClsMasterPayrollAccess
            Me.Baru = True
            Me.KosongkanField()
            Me.NonAktifkanDetail()
            Me.Load_1 = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.ObjCD.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            Me.ObjCD.KosongkanControl(Me.GroupBox1.Controls, True, True, False)
            Me.ObjCD.KosongkanControl(Me.GroupBox2.Controls, True, True, False)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, False)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, False)
            Me.TxtPrimaryJob.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.clbSalary.DataSource = Nothing
            Me.clbUser.DataSource = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.GroupBox1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.GroupBox2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.Controls, True, True, True, True, True, True, True, True)
            Me.CancelButton = Me.btnExit
            Me.BackColorDefault()
            Me.clbSalary.Enabled = False
            Me.clbUser.Enabled = False
            Me.LblCK_SalaryFormula.Enabled = False
            Me.LblUC_SalaryFormula.Enabled = False
            Me.LblCK_Users.Enabled = False
            Me.LblUC_Users.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.GroupBox1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.GroupBox2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.BackColorReset()
            Me.clbSalary.Enabled = True
            Me.clbUser.Enabled = True
            Me.LblCK_SalaryFormula.Enabled = True
            Me.LblUC_SalaryFormula.Enabled = True
            Me.LblCK_Users.Enabled = True
            Me.LblUC_Users.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorReset()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.GroupBox1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.GroupBox2.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Color.White, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorDefault()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.GroupBox1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.GroupBox2.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtReadOnly.BackColor, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Kontrol Button
    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If Baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.Obj.SimpanBaru(Me.UserCode, Me.ObjCD, Me.TxtPrimaryJob.Text, Me) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.SimpanPerubahan(Me.UserCode, Me.ObjCD, Me.TxtPrimaryJob.Text, Me) = True Then
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

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Not Me.TxtJobCode.Text = "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Baru = True Then
                        Me.KosongkanField()
                        Me.TxtJobCode.Clear()
                    Else
                        Me.Obj.CariDetail(Me.TxtPrimaryJob.Text, Me.TBL_SF, Me.TBL_User, Me)
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
            Me.Baru = True
            Obj.ResetObject(Me.clbSalary, Me.clbUser)
            Me.AktifkanDetail()
            Me.KosongkanField()
            Me.ObjItem.IsiListSalaryFormula(Me.clbSalary, Me.TBL_SF)
            Me.ObjItem.IsiListBoxUser(Me.clbUser, Me.TBL_User)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.TxtJobCode.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ijk As New FrmMasterPayrollAccessCari
        Try
            'Me.BackgroundWorker1.RunWorkerAsync()
            ijk.BackColorAll = Me.BackColorAll
            ijk.cboCriteria.Items.Clear()
            ijk.txtCari.Text = ""
            ijk.dgvHasil.DataSource = ""
            ijk.ShowDialog()
            If Not ijk.mPrimary = -1 Then
                If Me.Obj.CariDetail(ijk.mPrimary, Me.TBL_SF, Me.TBL_User, Me) = True Then
                    If Me.Load_1 = True Then
                        Me.Obj.CariDetail(ijk.mPrimary, Me.TBL_SF, Me.TBL_User, Me)
                        Me.Load_1 = False
                    End If
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
            ijk.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtJobCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtJobCode.Text <> "" Then
                If Me.Obj.ApakahBolehEdit(Me.TxtPrimaryJob.Text, Me.UserCode, True, True, Me) = True Then
                    Me.Baru = False
                    Me.AktifkanDetail()
                    If Me.TxtRemarks.Enabled = True Then
                        Me.TxtJobCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtJobCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtJobCode.Text <> "" Then
                If Me.Obj.ApakahBolehHapus(Me.TxtPrimaryJob.Text, Me.UserCode, True, True, Me) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.Hapus(Me.UserCode, Me.ObjCD, Me.TxtPrimaryJob.Text, _
                         Me, True) = True Then
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
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtJobCode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If Baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.Obj.SimpanBaru(Me.UserCode, Me.ObjCD, Me.TxtPrimaryJob.Text, _
                            Me) = True Then
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
                            If Me.Obj.SimpanPerubahan(Me.UserCode, Me.ObjCD, Me.TxtPrimaryJob.Text, Me) = True Then
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

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim Ketemu As Boolean = False
        Try
            '1. cek data yang harus diisi
            KoleksiNamaData.Add("Code")
            KoleksiTextBox.Add(Me.TxtJobCode)
            KoleksiNamaData.Add("Description")
            KoleksiTextBox.Add(Me.TxtJobName)

            If Me.ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            '1. harus ada data item
            If Me.clbSalary.CheckedItems.Count <= 0 Then
                MsgBox("Please choose salary formula.", MsgBoxStyle.Information, "Information")
                Me.clbSalary.Focus()
                Return False
            End If
            If Me.clbUser.CheckedItems.Count <= 0 Then
                MsgBox("Please choose user.", MsgBoxStyle.Information, "Information")
                Me.clbUser.Focus()
                Return False
            End If
            If Me.ObjCD.CekDataKosong(Me.TxtJobCode.Text, False) = False Then
                MsgBox("Data can not contain the following characters: ' , % , or * .", MsgBoxStyle.Information, "Information")
                Me.TxtJobCode.Focus()
                Return False
            End If
            If Me.ObjCD.CekDataKosong(Me.TxtJobName.Text, False) = False Then
                MsgBox("Data can not contain the following characters: ' , % , or * .", MsgBoxStyle.Information, "Information")
                Me.TxtJobName.Focus()
                Return False
            End If
            If Me.ObjCD.CekDataKosong(Me.TxtRemarks.Text, False) = False Then
                MsgBox("Data can not contain the following characters: ' , % , or * .", MsgBoxStyle.Information, "Information")
                Me.TxtRemarks.Focus()
                Return False
            End If

            '2. cek data yang salah, double dalam kode
            If Me.Baru = True Then
                If Me.Obj.ApakahSudahAda(Me.TxtJobCode.Text) = True Then
                    MsgBox("Duplication in Document no, please change your code", MsgBoxStyle.Information, "Information")
                    Me.TxtJobCode.Focus()
                    Return False
                End If
            Else
                Dim TipeCodeLama As String = Me.Obj.GetCode(Me.TxtPrimaryJob.Text)
                If TipeCodeLama <> Me.TxtJobCode.Text Then
                    If Me.Obj.ApakahSudahAda(Me.TxtJobCode.Text) = True Then
                        MsgBox("Duplication in Document no, please change your code", MsgBoxStyle.Information, "Information")
                        Me.TxtJobCode.Focus()
                        Return False
                    End If
                End If
                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.Obj.GetTimestamp(Me.TxtPrimaryJob.Text)
                If Me.Obj.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have modified this data, reload latest data?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.Obj.CariDetail(Me.Obj.Primary, Me.TBL_SF, Me.TBL_User, Me)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'sampe sini
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
            With Me.Obj
                .Code = Me.TxtJobCode.Text
                .Description = Me.TxtJobName.Text
                .Remarks = Me.TxtRemarks.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.Obj
                Me.TxtPrimaryJob.Text = .Primary
                Me.TxtJobCode.Text = .Code
                Me.TxtJobName.Text = .Description
                Me.TxtRemarks.Text = .Remarks
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LblCK_SalaryFormula_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LblCK_SalaryFormula.Click
        Try
            For i As Integer = 0 To Me.clbSalary.Items.Count - 1
                Me.clbSalary.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LblUC_SalaryFormula_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LblUC_SalaryFormula.Click
        Try
            For i As Integer = 0 To Me.clbSalary.Items.Count - 1
                 Me.clbSalary.SetItemCheckState(i, Windows.Forms.CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LblCK_Users_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LblCK_Users.Click
        Try
            For i As Integer = 0 To Me.clbUser.Items.Count - 1
                Me.clbUser.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LblUC_Users_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LblUC_Users.Click
        Try
            For i As Integer = 0 To Me.clbUser.Items.Count - 1
                Me.clbUser.SetItemCheckState(i, Windows.Forms.CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtJobCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJobCode.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtJobCode.Enabled = True Then
                    If Me.TxtJobCode.ReadOnly = False Then
                        TxtJobCode.BackColor = Me.BackColorAll
                    Else
                        TxtJobCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtJobCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJobCode.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtJobCode.Enabled = True Then
                    If Me.TxtJobCode.ReadOnly = False Then
                        TxtJobCode.BackColor = Me.TxtPrimaryJob.BackColor
                    Else
                        TxtJobCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtJobName_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJobName.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtJobName.Enabled = True Then
                    If Me.TxtJobName.ReadOnly = False Then
                        TxtJobName.BackColor = Me.BackColorAll
                    Else
                        TxtJobName.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtJobName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJobName.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtJobName.Enabled = True Then
                    If Me.TxtJobName.ReadOnly = False Then
                        TxtJobName.BackColor = Me.TxtPrimaryJob.BackColor
                    Else
                        TxtJobName.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarks_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.Enabled = True Then
                    If Me.TxtRemarks.ReadOnly = False Then
                        TxtRemarks.BackColor = Me.BackColorAll
                    Else
                        TxtRemarks.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarks_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.Enabled = True Then
                    If Me.TxtRemarks.ReadOnly = False Then
                        TxtRemarks.BackColor = Me.TxtPrimaryJob.BackColor
                    Else
                        TxtRemarks.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Dim ijk As New FrmMasterPayrollAccessCari
        Try
            ijk.BackColorAll = Me.BackColorAll
            ijk.cboCriteria.Items.Clear()
            ijk.txtCari.Text = ""
            ijk.dgvHasil.DataSource = ""
            ijk.ShowDialog()
            If Not ijk.mPrimary = -1 Then
                If Me.Obj.CariDetail(ijk.mPrimary, Me.TBL_SF, Me.TBL_User, Me) = True Then
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
            ijk.Dispose()
        End Try
    End Sub
End Class