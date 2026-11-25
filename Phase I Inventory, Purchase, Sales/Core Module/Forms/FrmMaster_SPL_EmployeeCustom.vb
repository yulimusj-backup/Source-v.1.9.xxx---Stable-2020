Option Explicit On
Option Strict Off

Public Class FrmMaster_SPL_EmployeeCustom
    Public UserCode As String = "USER"
    Public PrimaryEmployee As Double = -1
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim baru As Boolean
    Dim obj As New ClsMaster_SPL_EmployeeCustom
    Dim objEmployee As New ClsMasterKaryawan

    Dim ObjCD As New Core_Module.ClsCekData
    Dim buffer As String = ""
    Dim EraseRowIndex As Integer = 0

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimary.BackColor, False)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.DgvItem.ReadOnly = False
            Me.DgvItem.AllowUserToAddRows = True
            Me.DgvItem.AllowUserToDeleteRows = True


            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.DgvItem.AllowUserToAddRows = False
            Me.DgvItem.AllowUserToDeleteRows = False
            Me.DgvItem.ReadOnly = True

            Me.CancelButton = Me.btnExit
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, False, False)
            Me.obj.ResetObject(Me.DgvItem)

            Me.TxtPrimary.Text = "-1"
            Me.TxtCode.Text = ""
            Me.TxtDesc.Text = ""
            Me.TxtRemarks.Text = ""

            Me.tslModifiedBy.Text = "ModifiedBy: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.tslCreatedBy.Text = "Created  By: "
            Me.tslCreatedDate.Text = "Created Date: "

            Me.DgvItem.Rows.Clear()
            Me.buffer = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("Code")
            KoleksiTextBox.Add(Me.TxtCode)
            KoleksiNamaData.Add("Description")
            KoleksiTextBox.Add(Me.TxtDesc)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.obj.ApakahSudahAda(Me.TxtCode.Text) = True Then
                    MsgBox("Duplication in code, please change your code", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtCode.Focus()
                    Return False
                End If
            Else
                Dim ProductGroupLama As String = Me.obj.GetCode(Me.TxtPrimary.Text)
                If ProductGroupLama <> Me.TxtCode.Text Then
                    If Me.obj.ApakahSudahAda(Me.TxtCode.Text) = True Then
                        MsgBox("Duplication in code, please change your code", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.obj.GetTimestamp(Me.TxtPrimary.Text)
                If Me.obj.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.obj.CariDetail(Me.obj.Primary, Me.DgvItem, Me.ObjCD, Me)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'Item tidak boleh kosong 
            If Me.DgvItem.Rows.Count = 1 Then
                MsgBox("Please input item employee", MsgBoxStyle.Information, "Information")
                Return False
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
            With Me.obj
                .Code = Me.TxtCode.Text
                .Desc = Me.TxtDesc.Text
                .Remarks = Me.TxtRemarks.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()

        Try
            With Me.obj
                Me.TxtCode.Text = .Code
                Me.TxtDesc.Text = .Desc
                Me.TxtRemarks.Text = .Remarks
                Me.TxtPrimary.Text = .Primary

                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Text = "Overtime Order Employee Group"
            'Me.obj = New ClsMasterWorkUnit
            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.baru = True
            Me.NonAktifkanDetail()
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
                        If Me.obj.SimpanBaru(Me.UserCode, Me.TxtPrimary.Text, Me.ObjCD, Me.DgvItem, Me) = True Then
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
                        If Me.obj.SimpanPerubahan(Me.UserCode, Me.TxtPrimary.Text, Me.ObjCD, Me.DgvItem, Me) = True Then
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
                        Me.obj.CariDetail(Me.TxtPrimary.Text, Me.DgvItem, Me.ObjCD, Me)
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
            Me.baru = True
            Me.AktifkanDetail()
            Me.KosongkanField()

            Me.TxtCode.Focus()
            Me.tslModifiedBy.Text = "Created By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslCreatedBy.Text = "Modified By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmMaster_SPL_EmployeeCustom_Cari
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.PrimaryTransaksi = -1 Then
                If Me.obj.CariDetail(hj.PrimaryTransaksi, Me.DgvItem, Me.ObjCD, Me) = True Then
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
        Try
            If Me.TxtCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtCode.Text <> "" Then
                Me.baru = False
                Me.AktifkanDetail()
                Me.TxtCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtCode.Text <> "" Then
                hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.obj.Hapus(Me.TxtPrimary.Text, Me.UserCode, Me) = True Then
                        Me.KosongkanField()
                        Me.btnNew.Focus()
                        MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
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
                            If Me.obj.SimpanBaru(Me.UserCode, Me.TxtPrimary.Text, Me.ObjCD, Me.DgvItem, Me) = True Then
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
                            If Me.obj.SimpanPerubahan(Me.UserCode, Me.TxtPrimary.Text, Me.ObjCD, Me.DgvItem, Me) = True Then
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

    Private Sub TxtCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCode.GotFocus, TxtDesc.GotFocus, _
    TxtRemarks.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.BackColorAll
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarks_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.Leave
        Me.TxtRemarks.Text = Me.TxtRemarks.Text.ToString.Replace("'", "").Replace("*", "").Replace("%", "")
    End Sub

    Private Sub TxtCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCode.LostFocus, TxtDesc.LostFocus, _
        TxtRemarks.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.TxtPrimary.BackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvItem.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColEmpCode.Index Then
                Me.buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemEmployee(ByVal vPrimaryEmployee As Double, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvItem.Rows.Count - 2
                RowDGV = Me.DgvItem.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColPrimaryEmp.Index).Value = vPrimaryEmployee Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItem.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvItem.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryAcuan As Double = -1
        Dim RowIdx As Integer = e.RowIndex

        Try
            With row
                If IsNothing(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColEmpCode.Index Then
                        'FUNGSI COL EMPLOYEE CODE
                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        If Me.objEmployee.ApakahSudahAda(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryAcuan = Me.objEmployee.GetPrimaryByCode(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemEmployee(PrimaryAcuan, e.RowIndex)
                            If index < 0 Then
                                If Me.objEmployee.ApakahAktif(PrimaryAcuan) = False Then
                                    MsgBox("Employee not active", MsgBoxStyle.Information, "Information")
                                    GoTo DeleteRows
                                End If

                                'jika employee ditambahkan maka isi dgv sch setup, dgv sch detail, dgv approval
                                Me.DgvItem.Item(Me.ColPrimaryEmp.Index, e.RowIndex).Value = PrimaryAcuan
                                Me.DgvItem.Item(Me.ColEmpCode.Index, e.RowIndex).Value = Me.objEmployee.GetCodeByPrimary(PrimaryAcuan)
                                Me.DgvItem.Item(Me.ColEmpName.Index, e.RowIndex).Value = Me.objEmployee.GetFullNameByPrimary(PrimaryAcuan)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication employee code, please input another code.", MsgBoxStyle.Information, "Information")
                                    GoTo DeleteRows
                                End If
                            End If
                        Else
                            Dim ij As New FrmMasterKaryawanCari
                            ij.txtCari.Text = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryEmploye <> -1 Then
                                index = Me.GetIndexItemEmployee(ij.PrimaryEmploye, e.RowIndex)
                                If index < 0 Then
                                    If Me.objEmployee.ApakahAktif(ij.PrimaryEmploye) = False Then
                                        MsgBox("Employee not active", MsgBoxStyle.Information, "Information")
                                        GoTo DeleteRows
                                    End If

                                    'jika employee ditambahkan maka isi dgv sch setup, dgv sch detail, dgv approval
                                    Me.DgvItem.Item(Me.ColPrimaryEmp.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                    Me.DgvItem.Item(Me.ColEmpCode.Index, e.RowIndex).Value = Me.objEmployee.GetCodeByPrimary(ij.PrimaryEmploye)
                                    Me.DgvItem.Item(Me.ColEmpName.Index, e.RowIndex).Value = Me.objEmployee.GetFullNameByPrimary(ij.PrimaryEmploye)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication employee code, please input another code.", MsgBoxStyle.Information, "Information")
                                        GoTo DeleteRows
                                    Else
                                        If Me.objEmployee.ApakahAktif(ij.PrimaryEmploye) = False Then
                                            MsgBox("Employee not active", MsgBoxStyle.Information, "Information")
                                            GoTo DeleteRows
                                        End If

                                        'jika employee ditambahkan maka isi dgv sch setup, dgv sch detail, dgv approval
                                        Me.DgvItem.Item(Me.ColPrimaryEmp.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                        Me.DgvItem.Item(Me.ColEmpCode.Index, e.RowIndex).Value = Me.objEmployee.GetCodeByPrimary(ij.PrimaryEmploye)
                                        Me.DgvItem.Item(Me.ColEmpName.Index, e.RowIndex).Value = Me.objEmployee.GetFullNameByPrimary(ij.PrimaryEmploye)
                                    End If
                                End If
                            ElseIf ij.PrimaryEmploye = -1 Then
                                GoTo DeleteRows
                            End If

                            ij = Nothing
                        End If
                    ElseIf e.ColumnIndex = Me.ColRemarks.Index Then
                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Replace("'", "").Replace("*", "").Replace("%", "")
                    End If
                Else
DeleteRows:
                    If String.IsNullOrEmpty(Me.buffer) = True Then
                        If Me.DgvItem.Rows.Count - 1 <> e.RowIndex Then
                            Me.EraseRowIndex = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If
                    Else
                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvItem.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            If Me.TxtDesc.Enabled = True Then
                rowDGV = Me.DgvItem.Rows(Me.DgvItem.RowCount - 1)
                With rowDGV
                    .Cells(Me.ColNo.Index).Value = Me.DgvItem.RowCount
                    .Cells(Me.ColEmpCode.Index).Value = ""
                    .Cells(Me.ColEmpName.Index).Value = ""
                    .Cells(Me.ColRemarks.Index).Value = ""
                    .Cells(Me.ColPrimaryEmp.Index).Value = "-1"
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvItem.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DgvItem.UserDeletedRow
        Try
            'panggil fungsi set number
            Me.ObjCD.SetNumber(Me.DgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.DgvItem.Rows.RemoveAt(Me.EraseRowIndex)
            Me.ObjCD.SetNumber(Me.DgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class