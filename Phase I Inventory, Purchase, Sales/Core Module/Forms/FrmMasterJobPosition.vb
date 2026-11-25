Option Explicit On
Option Strict Off

Public Class FrmMasterJobPosition
    Public UserCode As String = "USER"
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim baru As Boolean
    Dim Obj As New ClsMasterJobPosition
    Dim buffer As String = ""
    Dim ObjCD As New ClsCekData
    Dim IndexRemove As Integer = -1
    Dim ObjKompetensi As New ClsMasterKompetensi
    Dim ObjKaryawan As New ClsMasterKaryawan

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.BackColorReset()

            Me.dgvCompetence.ReadOnly = False
            Me.dgvCompetence.AllowUserToAddRows = True
            Me.dgvCompetence.AllowUserToDeleteRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.Controls, True, True, True, True, True, True, True, True)
            Me.CancelButton = Me.btnExit

            Me.dgvCompetence.ReadOnly = True
            Me.dgvCompetence.AllowUserToAddRows = False
            Me.dgvCompetence.AllowUserToDeleteRows = False

            Me.BackColorDefault()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorReset()
        Try
            Me.TxtCode.BackColor = Me.TxtPrimary.BackColor
            Me.TxtDesc.BackColor = Me.TxtPrimary.BackColor
            Me.TxtRemarks.BackColor = Me.TxtPrimary.BackColor
            Me.TxtJobDesc.BackColor = Me.TxtPrimary.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorDefault()
        Try
            Me.TxtCode.BackColor = Me.TxtReadOnly.BackColor
            Me.TxtDesc.BackColor = Me.TxtReadOnly.BackColor
            Me.TxtRemarks.BackColor = Me.TxtReadOnly.BackColor
            Me.TxtJobDesc.BackColor = Me.TxtReadOnly.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.ObjCD.KosongkanControl(Me.gboDetail.Controls, True, False, False)
            Me.TxtPrimary.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "

            Me.dgvCompetence.Rows.Clear()
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

            If Me.ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data default precision harus non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.Obj.ApakahSudahAda(Me.TxtCode.Text) = True Then
                    MsgBox("Code duplication, please change code.", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtCode.Focus()
                    Return False
                End If
            Else
                Dim TipeCodeLama As String = Me.Obj.GetCode(Me.TxtPrimary.Text)
                If TipeCodeLama <> Me.TxtCode.Text Then
                    If Me.Obj.ApakahSudahAda(Me.TxtCode.Text) = True Then
                        MsgBox("Code duplication, please change code.", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.Obj.GetTimestamp(Me.TxtPrimary.Text)
                If Me.Obj.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user changed this data. Reload from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.Obj.CariDetail(Me.Obj.Primary, Me.dgvCompetence, Me)
                        Me.SalinKeField()
                    End If
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
            With Me.Obj
                .Code = Me.TxtCode.Text
                .Desc = Me.TxtDesc.Text
                .Remarks = Me.TxtRemarks.Text
                .JobDesc = Me.TxtJobDesc.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.Obj
                Me.TxtPrimary.Text = .Primary
                Me.TxtCode.Text = .Code
                Me.TxtDesc.Text = .Desc
                Me.TxtRemarks.Text = .Remarks
                Me.TxtJobDesc.Text = .JobDesc

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
            Me.Obj = New ClsMasterJobPosition
            Me.baru = True
            Me.KosongkanField()
            Me.NonAktifkanDetail()
            Me.dgvCompetence.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
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
                    hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.Obj.SimpanBaru(Me.UserCode, Me.dgvCompetence, Me) Then
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
                        If Me.Obj.SimpanPerubahan(Me.UserCode, Me.dgvCompetence, Me) Then
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
            If Not Me.TxtCode.Text = "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Konfirmasi")
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
            Me.baru = True
            Me.AktifkanDetail()
            Me.KosongkanField()
            Me.TxtPrimary.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

            Me.TxtDesc.Focus()
            Me.TxtCode.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ijk As New FrmMasterJobPosition_Cari

        Try
            ijk.BackColorAll = Me.BackColorAll
            ijk.cboCriteria.Items.Clear()
            ijk.txtCari.Text = ""
            ijk.dgvHasil.DataSource = ""
            ijk.mStatusAktif = True
            ijk.mSearchAll = True
            ijk.ShowDialog()

            If Not ijk.mPrimary = -1 Then
                If Me.Obj.CariDetail(ijk.mPrimary, Me.dgvCompetence, Me) = True Then
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
                'khusus untuk IDR, tidak bisa dihapus, karena digunakan oleh sistem
                If Me.Obj.ApakahBolehHapus(Me.TxtPrimary.Text, Me.UserCode, True, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.Hapus(Me.TxtPrimary.Text, Me.UserCode, True, True) = True Then
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
                            If Me.Obj.SimpanBaru(Me.UserCode, Me.dgvCompetence, Me) = True Then
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
                            If Me.Obj.SimpanPerubahan(Me.UserCode, Me.dgvCompetence, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Infomation")
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

    Private Sub TxtJobDesc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJobDesc.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtJobDesc.Enabled = True Then
                    If Me.TxtJobDesc.ReadOnly = False Then
                        TxtJobDesc.BackColor = Me.BackColorAll
                    Else
                        TxtJobDesc.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtJobDesc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJobDesc.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtJobDesc.Enabled = True Then
                    If Me.TxtJobDesc.ReadOnly = False Then
                        TxtJobDesc.BackColor = Me.TxtPrimary.BackColor
                    Else
                        TxtJobDesc.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCode.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtCode.Enabled = True Then
                    If Me.TxtCode.ReadOnly = False Then
                        TxtCode.BackColor = Me.BackColorAll
                    Else
                        TxtCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCode.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtCode.Enabled = True Then
                    If Me.TxtCode.ReadOnly = False Then
                        TxtCode.BackColor = Me.TxtPrimary.BackColor
                    Else
                        TxtCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNama_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDesc.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtDesc.Enabled = True Then
                    If Me.TxtDesc.ReadOnly = False Then
                        TxtDesc.BackColor = Me.BackColorAll
                    Else
                        TxtDesc.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDesc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDesc.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtDesc.Enabled = True Then
                    If Me.TxtDesc.ReadOnly = False Then
                        TxtDesc.BackColor = Me.TxtPrimary.BackColor
                    Else
                        TxtDesc.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.ReadOnly = False Then
                    Me.TxtRemarks.BackColor = Me.BackColorAll
                Else
                    TxtRemarks.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.ReadOnly = False Then
                    Me.TxtRemarks.BackColor = Me.TxtPrimary.BackColor
                Else
                    TxtRemarks.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCompetence_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvCompetence.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            If Me.TxtRemarks.Enabled = True Then
                rowDGV = Me.dgvCompetence.Rows(Me.dgvCompetence.Rows.Count - 1)

                rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvCompetence.Rows.Count
                rowDGV.Cells(Me.ColCode.Index).Value = ""
                rowDGV.Cells(Me.ColDefaultValue.Index).Value = ""
                rowDGV.Cells(Me.ColDesc.Index).Value = ""
                rowDGV.Cells(Me.ColPeriodicUpdate.Index).Value = Boolean.FalseString
                rowDGV.Cells(Me.ColPrimaryCompetence.Index).Value = -1
                rowDGV.Cells(Me.ColRemarks.Index).Value = ""
                rowDGV.Cells(Me.ColValue.Index).Value = ""
                rowDGV.Cells(Me.ColValueType.Index).Value = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCompetence_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvCompetence.CellBeginEdit
        Try
            If Me.TxtRemarks.Enabled = True Then
                If e.RowIndex = Me.dgvCompetence.Rows.Count - 1 Then
                    'jika baris terakhir yang bisa diedit: code
                    If e.ColumnIndex <> Me.ColCode.Index Then
                        e.Cancel = True
                    Else
                        Me.buffer = Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    'jika bukan baris terakhir: code no, remarks, periodic update, value
                    If e.ColumnIndex = Me.ColCode.Index Or _
                    e.ColumnIndex = Me.ColPeriodicUpdate.Index Or _
                    e.ColumnIndex = Me.ColRemarks.Index Or _
                    e.ColumnIndex = Me.ColValue.Index Then
                        'jika sudah ada yang merujuk pada master karyawan master kompetensi, maka tidak bisa diedit
                        If e.ColumnIndex = Me.ColPeriodicUpdate.Index Then
                            If Me.ObjKaryawan.GetJmlJobPositionKompetensi(Me.dgvCompetence.Item(Me.ColPrimaryCompetence.Index, e.RowIndex).Value, Me.TxtPrimary.Text) _
                            > 0 Then
                                MsgBox("There are employees using this data, you cannot edit this data", MsgBoxStyle.Information, "Information")
                                e.Cancel = True
                            End If
                        End If

                        Me.buffer = Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexDGV(ByVal vPrimaryKompetensi As Double) As Double
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvCompetence.Rows
                If rowDGV.Index <> Me.dgvCompetence.Rows.Count - 1 Then
                    If rowDGV.Cells(Me.ColPrimaryCompetence.Index).Value = vPrimaryKompetensi Then
                        Return rowDGV.Index
                    End If
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub dgvCompetence_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCompetence.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvCompetence.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryKompetensi As Double
        Dim DataType As ClsMasterKompetensi.EnumDataType = ClsMasterKompetensi.EnumDataType.Alfabetic

        Try
            With row
                If IsNothing(Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'bila tidak ada perubahan
                    If Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                        Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If

                    '==Code No
                    If e.ColumnIndex = Me.ColCode.Index Then
                        Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        If Me.ObjKompetensi.ApakahSudahAda(Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryKompetensi = Me.ObjKompetensi.GetPrimary(Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexDGV(PrimaryKompetensi)

                            If index < 0 Then
                                Me.ObjKompetensi.CariDetail(Me.ObjKompetensi.GetPrimary(Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value))
                                Me.SalinkeDGV(e, Me.ObjKompetensi)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication in competence code, please change your competence code", MsgBoxStyle.Information, "Information")
                                    Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    If Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                        Me.IndexRemove = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    End If

                                    Exit Sub
                                Else
                                    Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                            End If
                        Else
                            ''diketik tidak ada lalu tampilkan form cari
                            Dim ij As New Core_Module.FrmMasterKompetensi_Cari
                            ij.txtCari.Text = Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            index = Me.GetIndexDGV(ij.mPrimary)
                            If ij.mPrimary = -1 Then
                                ''hapuskan
                                If Me.buffer = "" Then
                                    Me.IndexRemove = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                Else
                                    Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                                Exit Sub
                            Else
                                If index < 0 Then
                                    Me.ObjKompetensi.CariDetail(ij.mPrimary)
                                    Me.SalinkeDGV(e, Me.ObjKompetensi)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication in competence code, please change your competence code", MsgBoxStyle.Information, "Information")
                                        Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        If Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                            Me.IndexRemove = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        End If
                                        Exit Sub
                                    Else
                                        Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Exit Sub
                                    End If
                                End If
                            End If
                            ij = Nothing
                        End If
                    ElseIf e.ColumnIndex = Me.ColValue.Index Then
                        DataType = Me.dgvCompetence.Item(Me.ColValueType.Index, e.RowIndex).Value
                        If DataType = ClsMasterKompetensi.EnumDataType.Alfabetic Then
                            'upper case
                            Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                            'pastikan nilainya benar
                            If Me.ObjKompetensi.ApakahValid_Aphabetic(Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                        ElseIf DataType = ClsMasterKompetensi.EnumDataType.Number Then
                            'harus numeric
                            If IsNumeric(Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                            'digit grouping
                            Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.ObjCD.DigitGrouping(Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value, _
                            Me.ObjKompetensi.GetDefPrec(Me.dgvCompetence.Item(Me.ColPrimaryCompetence.Index, e.RowIndex).Value))
                        ElseIf DataType = ClsMasterKompetensi.EnumDataType.Tanggal Then
                            'harus tanggal
                            If IsDate(Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                        ElseIf DataType = ClsMasterKompetensi.EnumDataType.Text Then
                            'ditrim
                            Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                        ElseIf DataType = ClsMasterKompetensi.EnumDataType.YN Then
                            'pilihannya hanya yes/no saja
                            If Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToLower <> "yes" And _
                            Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToLower <> "no" Then
                                If IsDate(Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                    Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                            End If
                            'format
                            If Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToLower <> "yes" Then
                                Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value = "Yes"
                            Else
                                Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value = "No"
                            End If
                        End If
                    ElseIf e.ColumnIndex = Me.ColRemarks.Index Then
                        'maksimum length = 250
                        If Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                            Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                        End If
                    End If
                Else
                    'hapus baris ybs
                    If e.ColumnIndex = Me.ColCode.Index Then
                        If Me.buffer = "" Then
                            Me.IndexRemove = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        Else
                            Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                        Exit Sub
                    ElseIf e.ColumnIndex = Me.ColPeriodicUpdate.Index Then
                        Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    ElseIf e.ColumnIndex = Me.ColValue.Index Then
                        If Me.dgvCompetence.Item(Me.ColValueType.Index, e.RowIndex).Value = ClsMasterKompetensi.EnumDataType.Text Then
                            Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value = ""
                            Exit Sub
                        Else
                            Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = Me.ColRemarks.Index Then
                        Me.dgvCompetence.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        Exit Sub
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinkeDGV(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, ByVal vObj As ClsMasterKompetensi)
        Try
            Me.dgvCompetence.Item(Me.ColPrimaryCompetence.Index, e.RowIndex).Value = Me.ObjKompetensi.Primary
            Me.dgvCompetence.Item(Me.ColCode.Index, e.RowIndex).Value = Me.ObjKompetensi.DataCode
            Me.dgvCompetence.Item(Me.ColDesc.Index, e.RowIndex).Value = Me.ObjKompetensi.Description
            Me.dgvCompetence.Item(Me.ColPeriodicUpdate.Index, e.RowIndex).Value = Boolean.FalseString
            Me.dgvCompetence.Item(Me.ColRemarks.Index, e.RowIndex).Value = ""
            Me.dgvCompetence.Item(Me.ColValueType.Index, e.RowIndex).Value = CInt(Me.ObjKompetensi.DataType)

            If Me.ObjKompetensi.DataType = ClsMasterKompetensi.EnumDataType.Text Then
                Me.dgvCompetence.Item(Me.ColValue.Index, e.RowIndex).Value = Me.ObjKompetensi.Value
                Me.dgvCompetence.Item(Me.ColDefaultValue.Index, e.RowIndex).Value = Me.ObjKompetensi.Value
            ElseIf Me.ObjKompetensi.DataType = ClsMasterKompetensi.EnumDataType.Alfabetic Then
                If Me.ObjKompetensi.Value = ClsMasterKompetensi.Aphabetic.A Then
                    Me.dgvCompetence.Item(Me.ColValue.Index, e.RowIndex).Value = "A"
                    Me.dgvCompetence.Item(Me.ColDefaultValue.Index, e.RowIndex).Value = "A"
                ElseIf Me.ObjKompetensi.Value = ClsMasterKompetensi.Aphabetic.B Then
                    Me.dgvCompetence.Item(Me.ColValue.Index, e.RowIndex).Value = "B"
                    Me.dgvCompetence.Item(Me.ColDefaultValue.Index, e.RowIndex).Value = "B"
                ElseIf Me.ObjKompetensi.Value = ClsMasterKompetensi.Aphabetic.C Then
                    Me.dgvCompetence.Item(Me.ColValue.Index, e.RowIndex).Value = "C"
                    Me.dgvCompetence.Item(Me.ColDefaultValue.Index, e.RowIndex).Value = "C"
                ElseIf Me.ObjKompetensi.Value = ClsMasterKompetensi.Aphabetic.D Then
                    Me.dgvCompetence.Item(Me.ColValue.Index, e.RowIndex).Value = "D"
                    Me.dgvCompetence.Item(Me.ColDefaultValue.Index, e.RowIndex).Value = "D"
                ElseIf Me.ObjKompetensi.Value = ClsMasterKompetensi.Aphabetic.E Then
                    Me.dgvCompetence.Item(Me.ColValue.Index, e.RowIndex).Value = "E"
                    Me.dgvCompetence.Item(Me.ColDefaultValue.Index, e.RowIndex).Value = "E"
                ElseIf Me.ObjKompetensi.Value = ClsMasterKompetensi.Aphabetic.F Then
                    Me.dgvCompetence.Item(Me.ColValue.Index, e.RowIndex).Value = "F"
                    Me.dgvCompetence.Item(Me.ColDefaultValue.Index, e.RowIndex).Value = "F"
                End If
            ElseIf Me.ObjKompetensi.DataType = ClsMasterKompetensi.EnumDataType.Number Then
                Me.dgvCompetence.Item(Me.ColValue.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.ObjKompetensi.Value, Me.ObjKompetensi.DefaultPrecision)
                Me.dgvCompetence.Item(Me.ColDefaultValue.Index, e.RowIndex).Value = Me.dgvCompetence.Item(Me.ColValue.Index, e.RowIndex).Value
            ElseIf Me.ObjKompetensi.DataType = ClsMasterKompetensi.EnumDataType.Tanggal Then
                Me.dgvCompetence.Item(Me.ColValue.Index, e.RowIndex).Value = Format(CDate(Me.ObjKompetensi.Value), "dd MMM yyyy")
                Me.dgvCompetence.Item(Me.ColDefaultValue.Index, e.RowIndex).Value = Me.dgvCompetence.Item(Me.ColValue.Index, e.RowIndex).Value
            ElseIf Me.ObjKompetensi.DataType = ClsMasterKompetensi.EnumDataType.YN Then
                If Me.ObjKompetensi.Value = 0 Then
                    Me.dgvCompetence.Item(Me.ColValue.Index, e.RowIndex).Value = "No"
                Else
                    Me.dgvCompetence.Item(Me.ColValue.Index, e.RowIndex).Value = "Yes"
                End If
                Me.dgvCompetence.Item(Me.ColDefaultValue.Index, e.RowIndex).Value = Me.dgvCompetence.Item(Me.ColValue.Index, e.RowIndex).Value
            End If

            Me.ObjCD.SetNumber(Me.dgvCompetence)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCompetence_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvCompetence.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvCompetence)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCompetence_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvCompetence.UserDeletingRow
        Try
            'tidak ada batasan tidak boleh hapus
            Me.buffer = Me.dgvCompetence.Item(Me.ColPrimaryCompetence.Index, e.Row.Index).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.dgvCompetence.Rows.RemoveAt(Me.IndexRemove)
            Me.ObjCD.SetNumber(Me.dgvCompetence)
            Me.IndexRemove = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCompetence_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCompetence.CellContentClick

    End Sub
End Class