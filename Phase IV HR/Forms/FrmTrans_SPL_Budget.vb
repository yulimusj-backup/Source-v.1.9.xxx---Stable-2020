Option Explicit On
Option Strict Off

Imports Core_Module, System.Drawing, System.Windows.Forms, System.Threading
Public Class FrmTrans_SPL_Budget
    Public UserCode As String = ""
    Public BackColorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll

    Dim baru As Boolean
    Dim objDepartement As New Core_Module.ClsMasterDepartement
    Dim ObjCD As New Core_Module.ClsCekData
    Dim Obj As New ClsTrans_SPL_Budget
    Dim ObjLicense As New ClsLicense
    Dim ObjUser As New Core_Module.ClsUser

    Dim ObjAntrian As New ClsAntrian
    Dim UserLain, NamaKomputerLain As String
    Dim mClose As Boolean = False
    Dim _ok As Boolean = False
    Dim e As System.Windows.Forms.DataGridViewCellEventArgs
    Dim Buffer As String = ""
    Dim EraseRowIndex As Integer = 0
    Dim BufferTransDate As Date = #1/1/1950#

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.txtPrimary.BackColor, False)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, False, False)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, False, False)
            Core_Module.ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, False, False)
            Core_Module.ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.dgvItem.AllowUserToAddRows = True
            Me.dgvItem.ReadOnly = False

            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None

            Me.dgvItem.TabStop = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, False, False)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, False, False)
            Core_Module.ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, False, False)
            Core_Module.ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.dgvItem.AllowUserToAddRows = False
            Me.dgvItem.ReadOnly = True

            Me.CancelButton = Me.btnExit
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.BackColor, False)

            If Me.dgvItem.RowCount = 0 Then
                Me.dgvItem.TabStop = False
            Else
                Me.dgvItem.TabStop = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Core_Module.ModCekData.KosongkanControl(Me.gboDetail.Controls, True, False, False)
            Core_Module.ModCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, False, False)
            Core_Module.ModCekData.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, False, False)

            Me.txtPrimary.Text = "-1"
            Me.txtDocumentNo.Text = ""
            Me.dtpDate.Value = Now
            Me.dtpPeriod.Value = Now
            Me.txtReadonly.Text = ""
            Me.dgvItem.Rows.Clear()

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim DocNoBuffer As String = ""

        Try
            'tidak boleh kosong
            KoleksiNamaData.Add("Document No")
            KoleksiTextBox.Add(Me.txtDocumentNo)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'tidak boleh double
            If Me.baru = True Then
                If Me.Obj.ApakahSudahAda(Me.txtDocumentNo.Text) = True Then
                    DocNoBuffer = Me.txtDocumentNo.Text
                    Me.txtDocumentNo.Text = Me.Obj.GetNoBaru(Me.dtpDate.Value, Me.txtPrimary.Text)
                    'MsgBox("Duplication in Document no, please change your code", MsgBoxStyle.Information, "Information")
                    'Me.TxtDocumentNo.Focus()
                    'Return False
                End If
            Else
                Dim CodeLama As String = Me.Obj.GetNoTransaksiByPrimary(Me.txtPrimary.Text)
                If CodeLama <> Me.txtDocumentNo.Text Then
                    If Me.Obj.ApakahSudahAda(Me.txtDocumentNo.Text) = True Then
                        DocNoBuffer = Me.txtDocumentNo.Text
                        Me.txtDocumentNo.Text = Me.Obj.GetNoBaru(Me.dtpDate.Value, Me.txtPrimary.Text)
                        'MsgBox("Duplication in Document no, please change your code", MsgBoxStyle.Information, "Information")
                        'Me.TxtDocumentNo.Focus()
                        'Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses(edit)
                Dim OldTimeStamp As Date = Me.Obj.GetTimeStampByPrimary(Me.txtPrimary.Text)
                If Me.Obj.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.Obj.CariDetail(Me.txtPrimary.Text, Me, Me.dgvItem)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'harus ada minimal 1 item 
            If Me.dgvItem.RowCount - 1 < 1 Then
                MsgBox("Please input item", MsgBoxStyle.Information, "Information")
                Me.dgvItem.Rows(0).Selected = True
                Return False
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            With Me.Obj
                .NoTransaksi = Me.txtDocumentNo.Text
                .Urutan = Me.txtDocumentNo.Text.Substring(0, 6)
                .Date_ = Me.dtpDate.Value
                .Period = Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpPeriod.Value)
                .Remarks = Me.txtRemarks.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.Obj
                Me.txtPrimary.Text = .Primary
                Me.txtDocumentNo.Text = .NoTransaksi
                Me.dtpDate.Value = .Date_
                Me.dtpPeriod.Value = .Period
                Me.txtRemarks.Text = .Remarks
              
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_SPL_Budget_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            'fungsi batasan hanya satu user yang dapat membuka form pada satu waktu
            'If Me.mClose = True Then
            '    MsgBox("Can't access salary formula transaction because user code : " + Me.UserLain.ToString + _
            '    " at computer name-ip address : " + Me.NamaKomputerLain.ToString + " is using this transaction.", _
            '    MsgBoxStyle.Information, "Information")
            'End If
            'Me.ObjAntrian.EndAntrianTunggal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_SPL_Budget_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'fungsi batasan hanya satu user yang dapat membuka form pada satu waktu
            'System.Windows.Forms.Form.CheckForIllegalCrossThreadCalls = False
            'UserLain = ""
            'NamaKomputerLain = ""
            'If Me.ObjAntrian.StartAntrianTunggal(ClsAntrian.EnumTipeAntrian.AntrianTunggalSalaryFormula, Me.UserCode, UserLain, NamaKomputerLain) = False Then
            '    Me.mClose = True
            '    Me.BackgroundWorker_CloseForm.RunWorkerAsync()
            'End If

            Me.NonAktifkanDetail()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.dgvItem.MultiSelect = False

            Me.dtpDate.Left = Me.SplitContainer1.Panel1.Width - Me.dtpDate.Width
            Me.txtDocumentNo.Width = Me.dtpDate.Left - 5 - Me.txtDocumentNo.Left
            Me.dtpPeriod.Width = Me.SplitContainer1.Panel1.Width - Me.dtpPeriod.Left
            Me.txtRemarks.Width = Me.SplitContainer1.Panel2.Width - Me.txtRemarks.Left
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult

        Try
            Me.Cursor = Windows.Forms.Cursors.WaitCursor
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    hasil = MsgBox("Save ?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.SimpanBaru(Me.UserCode, Me.dgvItem, Me) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save ?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.SimpanPerubahan(Me.UserCode, Me.dgvItem, Me) Then
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
        Finally
            Me.Cursor = Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult

        Try
            If Me.txtDocumentNo.Text <> "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.NonAktifkanDetail()
                        Me.KosongkanField()
                    Else
                        Me.NonAktifkanDetail()
                        Me.Obj.CariDetail(Me.txtPrimary.Text, Me, Me.dgvItem)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.NonAktifkanDetail()
                Me.KosongkanField()
            End If

            Me.btnNew.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmTrans_SPL_BudgetCari

        Try
            hj.BackColorAll = Me.BackColorAll
            hj.Frm = Me
            hj.ShowDialog()
            If Not hj.PrimaryTransaksi = -1 Then
                If Me.Obj.CariDetail(hj.PrimaryTransaksi, Me, Me.dgvItem) = True Then
                    Me.baru = False
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                    Me.NonAktifkanDetail()
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

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_SPL_Budget_New") = False Then
                MsgBox("You don't have permission to perform this action", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.baru = True

            Me.KosongkanField()
            Me.AktifkanDetail()

            Me.txtDocumentNo.Text = Me.Obj.GetNoBaru(Me.dtpDate.Value, Me.txtPrimary.Text)

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.dtpDate.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click

        Try
            If Me.txtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtDocumentNo.Text <> "" Then
                If Me.Obj.ApakahBolehedit(Me.UserCode, Me.txtPrimary.Text) = True Then
                    Me.baru = False
                    Me.Obj.CariDetail(Me.txtPrimary.Text, Me, Me.dgvItem)
                    Me.AktifkanDetail()
                    Me.dtpDate.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult

        Try
            If Me.txtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtDocumentNo.Text <> "" Then
                If Me.Obj.ApakahBolehHapus(Me.UserCode, Me.txtPrimary.Text) = True Then
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.Hapus(Me.txtPrimary.Text, Me.UserCode, Me) = True Then
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

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim FrmO As New Core_Module.FrmMasterUser_Otorisasi
        Dim FrmCetak As New FrmTrans_SPL_Budget_Cetak

        Try
            If Me.txtPrimary.Text = "" Or Me.txtPrimary.Text = -1 Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtPrimary.Text <> "" And Me.txtPrimary.Text <> -1 Then
                If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_SPL_Budget_Cetak") = False Then
                    MsgBox("You don't have permission to perform this action", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                If Me.Obj.Cetak(Me.txtPrimary.Text, FrmCetak.crv) = True Then
                    FrmCetak.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmO.Dispose()
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub FrmTrans_SPL_Budget_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Try
            Me.Cursor = Windows.Forms.Cursors.WaitCursor
            If Me.mClose = False Then
                If _ok = False Then
                    If Me.txtRemarks.Enabled = True Then
                        Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            If Me.ApakahDataBenar() = True Then
                                Me.SalinKeObjek()
                                If baru = True Then
                                    'simpan baru
                                    If Me.Obj.SimpanBaru(Me.UserCode, Me.dgvItem, Me) = True Then
                                        MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                        Me.SalinKeField()
                                        Me.NonAktifkanDetail()
                                        Me.btnFind.Focus()
                                        Me._ok = True
                                        Me.Close()
                                    Else
                                        e.Cancel = True
                                    End If
                                Else
                                    'simpan perubahan
                                    If Me.Obj.SimpanPerubahan(Me.UserCode, Me.dgvItem, Me) = True Then
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
                            'ElseIf hasil = Windows.Forms.DialogResult.No Then
                            '    Me.BackgroundWorker_CloseForm.RunWorkerAsync()
                        ElseIf hasil = MsgBoxResult.Cancel Then
                            e.Cancel = True
                        End If
                    End If
                End If
            Else
                Me.BackgroundWorker_CloseForm.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub FrmTrans_SPL_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColDeptCode.Index Then
                If e.ColumnIndex = Me.ColNo.Index Or e.ColumnIndex = Me.ColDeptDesc.Index Or e.ColumnIndex = Me.ColTotal.Index Then
                    e.Cancel = True
                Else
                    If String.IsNullOrEmpty(Me.dgvItem.Item(Me.ColDeptCode.Index, e.RowIndex).Value) = False Then
                        Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                End If
            Else
                Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemDept(ByVal vPrimaryDept As Double, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.dgvItem.Rows.Count - 2
                RowDGV = Me.dgvItem.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColDeptPrimary.Index).Value = vPrimaryDept Then
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

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim PrimaryDept As Double = -1

        Try
            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColDeptCode.Index Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    If Me.objDepartement.ApakahSudahAda(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        PrimaryDept = Me.objDepartement.GetPrimary(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                        index = Me.GetIndexItemDept(PrimaryDept, e.RowIndex)
                        If index < 0 Then
                            Me.dgvItem.Item(Me.ColDeptPrimary.Index, e.RowIndex).Value = PrimaryDept
                            Me.dgvItem.Item(Me.ColDeptDesc.Index, e.RowIndex).Value = Me.objDepartement.GetDeptDesc(PrimaryDept)
                        Else
                            If e.RowIndex <> index Then
                                MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                GoTo DeleteRows
                            End If
                        End If
                    Else
                        Dim ij As New FrmMasterDepartementCari
                        ij.txtCari.Text = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.BackColorAll = Me.BackColorAll
                        ij.ShowDialog()
                        If ij.PrimaryDepartement <> -1 Then
                            index = Me.GetIndexItemDept(ij.PrimaryDepartement, e.RowIndex)
                            If index < 0 Then
                                Me.dgvItem.Item(Me.ColDeptCode.Index, e.RowIndex).Value = Me.objDepartement.GetDeptCode(ij.PrimaryDepartement)
                                Me.dgvItem.Item(Me.ColDeptDesc.Index, e.RowIndex).Value = Me.objDepartement.GetDeptDesc(ij.PrimaryDepartement)
                                Me.dgvItem.Item(Me.ColDeptPrimary.Index, e.RowIndex).Value = ij.PrimaryDepartement
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    GoTo DeleteRows
                                Else
                                    Me.dgvItem.Item(Me.ColDeptCode.Index, e.RowIndex).Value = Me.objDepartement.GetDeptCode(ij.PrimaryDepartement)
                                    Me.dgvItem.Item(Me.ColDeptDesc.Index, e.RowIndex).Value = Me.objDepartement.GetDeptDesc(ij.PrimaryDepartement)
                                    Me.dgvItem.Item(Me.ColDeptPrimary.Index, e.RowIndex).Value = ij.PrimaryDepartement
                                End If
                            End If
                        ElseIf ij.PrimaryDepartement = -1 Then
                            GoTo DeleteRows
                        End If

                        ij = Nothing
                    End If

                    Dim PeriodVal As Date = Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpPeriod.Value)
                    If Me.Obj.ApakahDepartmentSudahDigunakan(Me.txtPrimary.Text, Me.dgvItem.Item(Me.ColDeptPrimary.Index, e.RowIndex).Value, PeriodVal) = True Then
                        MsgBox("Department (" & Me.dgvItem.Item(Me.ColDeptCode.Index, e.RowIndex).Value & " - " & Me.dgvItem.Item(Me.ColDeptDesc.Index, e.RowIndex).Value & ") has been used in period", MsgBoxStyle.Information, "Information")
                        GoTo DeleteRows
                    End If
                Else
                    If e.ColumnIndex <> Me.ColRemarks.Index Then
                        If Me.ObjCD.Is_Time_String(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, True) = False Then
                            MsgBox("Please input hh.mm format", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        Dim ConvertValToTime As Double = Me.ObjCD.Convert_ToMinutes(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.Convert_ToString(ConvertValToTime)

                        Me.HitungTotal(row)
                    Else
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Replace("'", "").Replace("*", "").Replace("%", "")
                    End If
                End If
            Else
DeleteRows:
                If String.IsNullOrEmpty(Me.Buffer) = True Then
                    If Me.dgvItem.Rows.Count - 1 <> e.RowIndex Then
                        Me.EraseRowIndex = e.RowIndex
                        Me.BackgroundWorker_DeleteRows.RunWorkerAsync()
                    End If
                Else
                    If e.ColumnIndex <> Me.ColRemarks.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Else
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Dim row As DataGridViewRow = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)

        Try
            row.Cells(Me.ColDeptPrimary.Index).Value = -1
            row.Cells(Me.ColNo.Index).Value = Me.dgvItem.RowCount
            row.Cells(Me.ColDeptCode.Index).Value = ""
            row.Cells(Me.ColDeptDesc.Index).Value = ""
            row.Cells(Me.ColJan.Index).Value = "00.00"
            row.Cells(Me.ColFeb.Index).Value = "00.00"
            row.Cells(Me.ColMar.Index).Value = "00.00"
            row.Cells(Me.ColApr.Index).Value = "00.00"
            row.Cells(Me.ColMay.Index).Value = "00.00"
            row.Cells(Me.ColJun.Index).Value = "00.00"
            row.Cells(Me.ColJul.Index).Value = "00.00"
            row.Cells(Me.ColAug.Index).Value = "00.00"
            row.Cells(Me.ColSept.Index).Value = "00.00"
            row.Cells(Me.ColOkt.Index).Value = "00.00"
            row.Cells(Me.ColNov.Index).Value = "00.00"
            row.Cells(Me.ColDec.Index).Value = "00.00"
            row.Cells(Me.ColTotal.Index).Value = "00.00"
            row.Cells(Me.ColRemarks.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvItem, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvItem, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        'batasan jika hapus baris item
    End Sub

    Private Sub BackgroundWorker_CloseForm_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_CloseForm.DoWork
        Try
            Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker_CloseForm_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker_CloseForm.RunWorkerCompleted
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker_DeleteRows_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_DeleteRows.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker_DeleteRows_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker_DeleteRows.RunWorkerCompleted
        Try
            Me.dgvItem.Rows.RemoveAt(Me.EraseRowIndex)
            Me.ObjCD.SetNumber(Me.dgvItem)
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
        Try
            If Me.Baru = True Then
                Me.txtDocumentNo.Text = Me.Obj.GetNoBaru(Me.dtpDate.Value, Me.txtPrimary.Text)
            Else
                Dim TransDate As Date = Me.Obj.GetTanggalTransaksi(Me.txtPrimary.Text)
                If Me.dtpDate.Value.Month = TransDate.Month And _
                Me.dtpDate.Value.Year = TransDate.Year Then
                    Me.txtDocumentNo.Text = Me.Obj.GetNoTransaksiByPrimary(Me.txtPrimary.Text)
                Else
                    If Me.dtpDate.Value.Month <> Me.Obj.Date_.Month Or _
                    Me.dtpDate.Value.Year <> Me.Obj.Date_.Year Then
                        Me.txtDocumentNo.Text = Me.Obj.GetNoBaru(Me.dtpDate.Value, Me.txtPrimary.Text)
                    Else
                        Me.txtDocumentNo.Text = Me.Obj.GetNoTransaksiByPrimary(Me.txtPrimary.Text)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpDate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.Enter
        Try
            Me.BufferTransDate = Me.dtpDate.Value.Date
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpDate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.Leave
        Try
            If Me.txtRemarks.Enabled = True Then
                Me.BackgroundWorker1.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpPeriod_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpPeriod.Leave
        Dim PeriodVal As Date = Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpPeriod.Value)

        Try
            'dalam satu periode hanya bisa input satu department saja.
            For Each row As DataGridViewRow In Me.dgvItem.Rows
                If row.Cells(Me.ColDeptPrimary.Index).Value <> -1 And String.IsNullOrEmpty(row.Cells(Me.ColDeptPrimary.Index).Value) = False Then
                    If Me.Obj.ApakahDepartmentSudahDigunakan(Me.txtPrimary.Text, row.Cells(Me.ColDeptPrimary.Index).Value, PeriodVal) = True Then
                        MsgBox("Department (" & row.Cells(Me.ColDeptCode.Index).Value & " - " & row.Cells(Me.ColDeptDesc.Index).Value & ") has been used in period", MsgBoxStyle.Information, "Information")
                        Me.dtpPeriod.Focus()
                        Exit Sub
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtRemarks_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemarks.GotFocus
        Me.txtRemarks.BackColor = Me.BackColorAll
    End Sub

    Private Sub txtRemarks_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemarks.Leave
        Me.txtRemarks.Text = Me.txtRemarks.Text.ToString.Replace("'", "").Replace("*", "").Replace("%", "")
    End Sub

    Private Sub txtRemarks_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemarks.LostFocus
        Me.txtRemarks.BackColor = Me.txtPrimary.BackColor
    End Sub

    Private Sub LlbAdvance_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbAdvance.LinkClicked
        If Me.btnOK.Enabled = True Then
            Me.ContextMenuStrip1.Show(Me.LlbAdvance, 0, 20)
        End If
    End Sub

    Private Sub AddAllDeptToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddAllDept.Click
        Try
            Me.GetDepartmentData_ByAdd(ClsTrans_SPL_Budget.EnumCariData.AllDept, -1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CopyFromOthDataToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyFromOthData.Click
        Dim Frm As New FrmTrans_SPL_BudgetCari

        Try
            Frm.BackColorAll = Me.BackColorAll
            Frm.ShowDialog()
            If Frm.PrimaryTransaksi <> -1 Then
                Me.GetDepartmentData_ByAdd(ClsTrans_SPL_Budget.EnumCariData.Other, Frm.PrimaryTransaksi)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetDepartmentData_ByAdd(ByVal vCariBerdasakan As ClsTrans_SPL_Budget.EnumCariData, ByVal vPrimaryBudget As Double)
        Dim dta As New DseTrans_SPL_BudgetTableAdapters.tblmaster_departemenTableAdapter
        Dim tbl As New DseTrans_SPL_Budget.tblmaster_departemenDataTable
        Dim row As DseTrans_SPL_Budget.tblmaster_departemenRow

        Dim Dta1 As New DseTrans_SPL_BudgetTableAdapters.tbltrans_spl_budget_itemTableAdapter
        Dim Tbl1 As New DseTrans_SPL_Budget.tbltrans_spl_budget_itemDataTable
        Dim row1 As DseTrans_SPL_Budget.tbltrans_spl_budget_itemRow

        Dim PeriodVal As Date = Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpPeriod.Value)
        Dim rowdgv As DataGridViewRow
        Dim FrmDialog As New Core_Module.FrmAddReplace
        Dim InputKeList As Boolean = False
        'Dim ObjKaryawan As New ClsMasterKaryawan
        Dim ListMsg_1 As String = ""
        Dim hasil As MsgBoxResult
        Dim CopyOTDuration As Boolean = False

        Try
            FrmDialog.ShowDialog()
            If FrmDialog.DialogResult = DialogResult.OK Then 'jika add 
            ElseIf FrmDialog.DialogResult = DialogResult.Yes Then 'jika replace
                Me.dgvItem.Rows.Clear()
            ElseIf FrmDialog.DialogResult = DialogResult.Cancel Then 'jika cancel
                Exit Sub
            Else
                Exit Sub
            End If

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            ''filter data
            If vCariBerdasakan = ClsTrans_SPL_Budget.EnumCariData.AllDept Then
                dta.Connection.ConnectionString = Core_Module.GetConstring
                dta.Fill(tbl)
            ElseIf vCariBerdasakan = ClsTrans_SPL_Budget.EnumCariData.Other Then
                hasil = MsgBox("Copy overtime duration?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    CopyOTDuration = True
                End If

                Dta1.Connection.ConnectionString = Core_Module.GetConstring
                Dta1.Fill(Tbl1, vPrimaryBudget)
            End If

            ''pindahkan ke dgv
            If vCariBerdasakan = ClsTrans_SPL_Budget.EnumCariData.AllDept Then
                For i As Integer = tbl.Rows.Count - 1 To 0 Step -1
                    row = tbl.Rows(i)
                    'cek data apakah sudah ada di dgv
                    'jika replace lewati pengecekan
                    InputKeList = True
                    If FrmDialog.DialogResult = DialogResult.OK Then 'jika add 
                        For Each rowdgv In Me.dgvItem.Rows
                            If rowdgv.Cells(Me.ColDeptPrimary.Index).Value = row.Primary Then
                                tbl.Rows.Remove(row)
                                InputKeList = False
                                Exit For
                            End If
                        Next
                    End If
                Next

                For Each row In tbl
                    InputKeList = True
                    If Me.Obj.ApakahDepartmentSudahDigunakan(Me.txtPrimary.Text, row.Primary, PeriodVal) = True Then
                        If ListMsg_1 = "" Then
                            ListMsg_1 &= "Department : " & row.Description
                        Else
                            ListMsg_1 &= ", " & row.Description
                        End If
                        InputKeList = False
                    End If

                    If InputKeList = True Then
                        Me.dgvItem.Rows.Add()
                        If Me.dgvItem.Rows.Count > 1 Then
                            rowdgv = Me.dgvItem.Rows(Me.dgvItem.RowCount - 2)
                        Else
                            rowdgv = Me.dgvItem.Rows(Me.dgvItem.RowCount)
                        End If

                        rowdgv.Cells(Me.ColNo.Index).Value = rowdgv.Index + 1
                        rowdgv.Cells(Me.ColDeptPrimary.Index).Value = row.Primary
                        rowdgv.Cells(Me.ColDeptCode.Index).Value = row.Dept_Code
                        rowdgv.Cells(Me.ColDeptDesc.Index).Value = row.Description
                        rowdgv.Cells(Me.ColJan.Index).Value = "00.00"
                        rowdgv.Cells(Me.ColFeb.Index).Value = "00.00"
                        rowdgv.Cells(Me.ColMar.Index).Value = "00.00"
                        rowdgv.Cells(Me.ColApr.Index).Value = "00.00"
                        rowdgv.Cells(Me.ColMay.Index).Value = "00.00"
                        rowdgv.Cells(Me.ColJun.Index).Value = "00.00"
                        rowdgv.Cells(Me.ColJul.Index).Value = "00.00"
                        rowdgv.Cells(Me.ColAug.Index).Value = "00.00"
                        rowdgv.Cells(Me.ColSept.Index).Value = "00.00"
                        rowdgv.Cells(Me.ColOkt.Index).Value = "00.00"
                        rowdgv.Cells(Me.ColNov.Index).Value = "00.00"
                        rowdgv.Cells(Me.ColDec.Index).Value = "00.00"
                        rowdgv.Cells(Me.ColTotal.Index).Value = "00.00"
                        rowdgv.Cells(Me.ColRemarks.Index).Value = ""
                    End If
                Next
            ElseIf vCariBerdasakan = ClsTrans_SPL_Budget.EnumCariData.Other Then
                'berdasarkan data spl yang sudah ada
                For i As Integer = Tbl1.Rows.Count - 1 To 0 Step -1
                    row1 = Tbl1.Rows(i)
                    'cek data apakah sudah ada di dgv
                    'jika replace lewati pengecekan
                    InputKeList = True
                    If FrmDialog.DialogResult = DialogResult.OK Then 'jika add 
                        For Each rowdgv In Me.dgvItem.Rows
                            If rowdgv.Cells(Me.ColDeptPrimary.Index).Value = row1.Primary_Department Then
                                Tbl1.Rows.Remove(row1)
                                InputKeList = False
                                Exit For
                            End If
                        Next
                    End If
                Next

                For Each row1 In Tbl1
                    InputKeList = True
                    If Me.Obj.ApakahDepartmentSudahDigunakan(Me.txtPrimary.Text, row1.Primary_Department, PeriodVal) = True Then
                        If ListMsg_1 = "" Then
                            ListMsg_1 &= "Department : " & row1.Dept_Desc
                        Else
                            ListMsg_1 &= ", " & row1.Dept_Desc
                        End If
                        InputKeList = False
                    End If

                    If InputKeList = True Then
                        Me.dgvItem.Rows.Add()
                        If Me.dgvItem.Rows.Count > 1 Then
                            rowdgv = Me.dgvItem.Rows(Me.dgvItem.RowCount - 2)
                        Else
                            rowdgv = Me.dgvItem.Rows(Me.dgvItem.RowCount)
                        End If

                        rowdgv.Cells(Me.ColNo.Index).Value = rowdgv.Index + 1
                        rowdgv.Cells(Me.ColDeptPrimary.Index).Value = row1.Primary_Department
                        rowdgv.Cells(Me.ColDeptCode.Index).Value = row1.Dept_Code
                        rowdgv.Cells(Me.ColDeptDesc.Index).Value = row1.Dept_Desc
                        rowdgv.Cells(Me.ColRemarks.Index).Value = row1.Remarks

                        If CopyOTDuration = True Then
                            rowdgv.Cells(Me.ColJan.Index).Value = Me.ObjCD.Convert_ToString(row1.Jan)
                            rowdgv.Cells(Me.ColFeb.Index).Value = Me.ObjCD.Convert_ToString(row1.Feb)
                            rowdgv.Cells(Me.ColMar.Index).Value = Me.ObjCD.Convert_ToString(row1.Mar)
                            rowdgv.Cells(Me.ColApr.Index).Value = Me.ObjCD.Convert_ToString(row1.Apr)
                            rowdgv.Cells(Me.ColMay.Index).Value = Me.ObjCD.Convert_ToString(row1.May)
                            rowdgv.Cells(Me.ColJun.Index).Value = Me.ObjCD.Convert_ToString(row1.Jun)
                            rowdgv.Cells(Me.ColJul.Index).Value = Me.ObjCD.Convert_ToString(row1.Jul)
                            rowdgv.Cells(Me.ColAug.Index).Value = Me.ObjCD.Convert_ToString(row1.Aug)
                            rowdgv.Cells(Me.ColSept.Index).Value = Me.ObjCD.Convert_ToString(row1.Sep)
                            rowdgv.Cells(Me.ColOkt.Index).Value = Me.ObjCD.Convert_ToString(row1.Oct)
                            rowdgv.Cells(Me.ColNov.Index).Value = Me.ObjCD.Convert_ToString(row1.Nov)
                            rowdgv.Cells(Me.ColDec.Index).Value = Me.ObjCD.Convert_ToString(row1.Dec)
                        Else
                            rowdgv.Cells(Me.ColJan.Index).Value = "00.00"
                            rowdgv.Cells(Me.ColFeb.Index).Value = "00.00"
                            rowdgv.Cells(Me.ColMar.Index).Value = "00.00"
                            rowdgv.Cells(Me.ColApr.Index).Value = "00.00"
                            rowdgv.Cells(Me.ColMay.Index).Value = "00.00"
                            rowdgv.Cells(Me.ColJun.Index).Value = "00.00"
                            rowdgv.Cells(Me.ColJul.Index).Value = "00.00"
                            rowdgv.Cells(Me.ColAug.Index).Value = "00.00"
                            rowdgv.Cells(Me.ColSept.Index).Value = "00.00"
                            rowdgv.Cells(Me.ColOkt.Index).Value = "00.00"
                            rowdgv.Cells(Me.ColNov.Index).Value = "00.00"
                            rowdgv.Cells(Me.ColDec.Index).Value = "00.00"
                            rowdgv.Cells(Me.ColRemarks.Index).Value = ""
                        End If

                        Me.HitungTotal(rowdgv)
                    End If
                Next
            End If

            If ListMsg_1 <> "" Then
                MsgBox("Department has been used in period." & vbCrLf & ListMsg_1, MsgBoxStyle.Information, "Information")
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmDialog.Dispose()
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Public Function HitungTotal(ByVal Row As DataGridViewRow) As Double
        Dim TotalOT As Double = 0

        Try
            For i As Integer = Me.ColJan.Index To Me.ColDec.Index
                TotalOT += Me.ObjCD.Convert_ToMinutes(Row.Cells(i).Value)
            Next

            Row.Cells(Me.ColTotal.Index).Value = Me.ObjCD.Convert_ToString(TotalOT)
            Return TotalOT
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function
End Class