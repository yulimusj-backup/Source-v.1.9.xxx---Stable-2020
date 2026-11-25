Option Explicit On
Option Strict Off

Public Class FrmTrans_SchDownTime

    Public userCode As String = "USER"
    Public backColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim baru As Boolean = True
    Dim buffer As Object = Nothing
    Dim objCd As New Core_Module.ClsCekData
    Dim obj As New ClsTrans_SchDownTime
    Dim objWorkUnit As New Core_Module.ClsMasterWorkUnit
    Dim rowToDelete As Integer = -1

    Dim objNumbering As New Core_Module.ClsMasterTransactionNumbering

    Private Sub AktifkanDetail()
        Try
            Me.objCd.SetBackColor(Me.gboDetail.Controls, Me.txtPrimary.BackColor, False)
            Me.objCd.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.txtPrimary.BackColor, False)
            Me.objCd.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.txtPrimary.BackColor, False)
            Me.objCd.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            Me.objCd.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.objCd.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            Me.objCd.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.dgvItem.ReadOnly = False
            Me.dgvItem.AllowUserToAddRows = True
            Me.dgvItem.AllowUserToDeleteRows = True
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.objCd.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            Me.objCd.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.objCd.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.objCd.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.objCd.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.objCd.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, False)
            Me.objCd.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.BackColor, False)
            Me.dgvItem.ReadOnly = True
            Me.dgvItem.AllowUserToAddRows = False
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.CancelButton = Me.btnExit
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.objCd.KosongkanControl(Me.gboDetail.Controls, True, True, True)
            Me.objCd.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, True)
            Me.objCd.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, True)
            Me.dgvItem.Rows.Clear()
            Me.txtPrimary.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Try
            If Me.baru = True Then
                If Me.obj.ApakahSudahAda(Me.txtDocumentNo.Text) = True Then
                    If Me.objNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Schedule_Downtime) = True Then
                        Me.txtDocumentNo.Text = Me.obj.GetNoBaru(Me.dtpDate.Value.Date)
                    Else
                        MsgBox("Document No: " & Me.txtDocumentNo.Text & " already exist.", MsgBoxStyle.Information, "Information")
                        Me.txtDocumentNo.Focus()
                        Return False
                    End If
                End If
            Else
                Dim noLama As String = Me.obj.GetdocumentNo(Me.txtPrimary.Text)
                If noLama <> Me.txtDocumentNo.Text Then
                    If Me.obj.ApakahSudahAda(Me.txtDocumentNo.Text) = True Then
                        If Me.objNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Schedule_Downtime) = True Then
                            Me.txtDocumentNo.Text = Me.obj.GetNoBaru(Me.dtpDate.Value.Date)
                        Else
                            MsgBox("Document No: " & Me.txtDocumentNo.Text & " already exist.", MsgBoxStyle.Information, "Information")
                            Me.txtDocumentNo.Focus()
                            Return False
                        End If
                    End If
                End If
                Dim oldTimeStamp As Date = Me.obj.GetTimestamp(Me.txtPrimary.Text)
                If Me.obj.Timestamp <> oldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.obj.CariDetail(Me.obj.Primary, Me.dgvItem, Me)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If
            If Me.txtDocumentNo.Text.Trim.Length = 0 Then
                MsgBox("Document No cannot be an empty string.", MsgBoxStyle.Information, "Information")
                Me.txtDocumentNo.Focus()
                Return False
            End If
            If Me.txtReason.Text.Trim.Length = 0 Then
                MsgBox("Reason cannot be an empty string.", MsgBoxStyle.Information, "Information")
                Me.txtReason.Focus()
                Return False
            End If
            If (Me.dgvItem.AllowUserToAddRows = True And Me.dgvItem.Rows.Count <= 1) Or _
                    (Me.dgvItem.AllowUserToAddRows = True And Me.dgvItem.Rows.Count <= 1) Then
                If Me.dgvItem.Rows.Count = 1 Then
                    MsgBox("Input at least 1 item.", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Focus()
                    Return False
                End If
            End If
            If Me.ApakahDataItemBenar = False Then
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Function

    Private Function ApakahDataItemBenar() As Boolean
        Dim r As DataGridViewRow
        Dim listPrimaryWorkUnit As New List(Of Integer)
        Dim listInvalidNumber As String = ""
        Dim listInvalidDate As New List(Of Integer)
        Dim listConflictTime As String = ""
        Try
            Me.dgvItem.AllowUserToAddRows = False
            For Each r In Me.dgvItem.Rows
                If r.Cells(Me.colWU_Primary.Index).Value = "-1" And String.IsNullOrEmpty(r.Cells(Me.colWU_Code.Index).Value) Then

                    ''work unit wajib diisi
                    If listPrimaryWorkUnit.Contains(r.Cells(Me.colWU_No.Index).Value) = False Then
                        listPrimaryWorkUnit.Add(r.Cells(Me.colWU_No.Index).Value)
                    End If

                Else

                    ''start and end date must be valid
                    If IsDate(r.Cells(Me.colWU_StartDate.Index).Value) = False Or IsDate(r.Cells(Me.colWU_EndDate.Index).Value) = False Then
                        If listInvalidDate.Contains(r.Cells(Me.colWU_No.Index).Value) = False Then
                            listInvalidDate.Add(r.Cells(Me.colWU_No.Index).Value)
                        End If
                    End If

                    ''check is conflict schedule with others item in the same transaction
                    listConflictTime += Me.getListConflictSchedule(r.Index, r.Cells(Me.colWU_Primary.Index).Value, r.Cells(Me.colWU_Code.Index).Value, _
                        CDate(r.Cells(Me.colWU_StartDate.Index).Value & " " & r.Cells(Me.colWU_StartTime.Index).Value.ToString.Replace(".", ":") & ":00"), _
                        CDate(r.Cells(Me.colWU_EndDate.Index).Value & " " & r.Cells(Me.colWU_EndTime.Index).Value.ToString.Replace(".", ":") & ":00"))

                    ''cek apakah bentrok dg item-item di transaksi lain pada DB

                End If
            Next
            Me.dgvItem.AllowUserToAddRows = True

            If listPrimaryWorkUnit.Count > 0 Then
                For Each no As Integer In listPrimaryWorkUnit
                    If listInvalidNumber = "" Then
                        listInvalidNumber = no.ToString
                    Else
                        listInvalidNumber += ", " & no.ToString
                    End If
                Next
                MsgBox("Please input work unit for the following rows: " & vbNewLine & listInvalidNumber & ".", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If listInvalidDate.Count > 0 Then
                For Each no As Integer In listInvalidDate
                    If listInvalidNumber = "" Then
                        listInvalidNumber = no.ToString
                    Else
                        listInvalidNumber += ", " & no.ToString
                    End If
                Next
                MsgBox("Please fix schedule datetime for the following rows: " & vbNewLine & listInvalidNumber & ".", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If listConflictTime <> "" Then
                MsgBox("Conflict schedule: " & listConflictTime, MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally

        End Try
    End Function

    Private Function getListConflictSchedule(ByVal vExcludeRow As Integer, ByVal vPrimaryWorkUnit As Double, ByVal vWorkUnitCode As String, ByVal vStartTime As Date, ByVal vEndTime As Date) As String
        Dim val As String = ""
        Dim otherStartTime, otherEndTime As Date
        Try
            For Each row As DataGridViewRow In Me.dgvItem.Rows
                If row.Index > vExcludeRow And vPrimaryWorkUnit.ToString = row.Cells(Me.colWU_Primary.Index).Value.ToString Then

                    otherStartTime = CDate(row.Cells(Me.colWU_StartDate.Index).Value & " " & row.Cells(Me.colWU_StartTime.Index).Value.ToString.Replace(".", ":") & ":00")
                    otherEndTime = CDate(row.Cells(Me.colWU_EndDate.Index).Value & " " & row.Cells(Me.colWU_EndTime.Index).Value.ToString.Replace(".", ":") & ":00")

                    If (vStartTime < otherStartTime And vEndTime >= otherStartTime) Or _
                        (vStartTime = otherStartTime) Or _
                        (vStartTime > otherStartTime And vStartTime <= otherEndTime) Then
                        val += vbNewLine & "Work Unit " & vWorkUnitCode & ", row " & (vExcludeRow + 1).ToString & " with row " & (row.Index + 1).ToString & "."
                    End If

                End If
            Next
            Return val
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            With Me.obj
                .Primary = Me.txtPrimary.Text
                .DocumentNo = Me.txtDocumentNo.Text
                .Date_ = Me.dtpDate.Value.Date
                .Reason = Me.txtReason.Text
                .Remarks = Me.txtRemarks.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.obj
                Me.txtPrimary.Text = .Primary
                Me.txtDocumentNo.Text = .DocumentNo
                Me.dtpDate.Value = .Date_
                Me.txtReason.Text = .Reason
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

    Private Sub FrmTransaksiDeliveryOperation_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.txtRemarks.Enabled = True And (Me.txtReason.Text.Trim <> "" Or Me.txtRemarks.Text.Trim <> "" Or _
                    (Me.dgvItem.Rows.Count > 1 And Me.dgvItem.AllowUserToAddRows = True) Or (Me.dgvItem.Rows.Count > 0 _
                    And Me.dgvItem.AllowUserToAddRows = False)) Then
                Dim hasil As MsgBoxResult = MsgBox("Save data before close?", MsgBoxStyle.YesNoCancel, "Confirmation")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.obj.SimpanBaru(Me.userCode, Me.dgvItem, Me) = True Then
                                MsgBox("Data saved.", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.obj.SimpanPerubahan(Me.userCode, Me.dgvItem, Me) = True Then
                                MsgBox("Date change saved.", MsgBoxStyle.Information, "Information")
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

    Private Sub FrmTransaksiDeliveryOperation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If Me.objNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Schedule_Downtime) = True Then
                Me.txtDocumentNo.ReadOnly = True
            Else
                Me.txtDocumentNo.ReadOnly = False
            End If
            FrmTrans_SchDownTime.CheckForIllegalCrossThreadCalls = False
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.backColorAll
            Me.baru = True
            Me.NonAktifkanDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() Then
                Me.SalinKeObjek()
                If Me.obj.ApakahDataBenar(Me.txtPrimary.Text, Me.dgvItem, Me, Me.baru) Then
                    If baru = True Then
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirmation")
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.obj.SimpanBaru(Me.userCode, Me.dgvItem, Me) Then
                                MsgBox("Data saved.", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            End If
                        End If
                    Else
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirmation")
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.obj.SimpanPerubahan(Me.userCode, Me.dgvItem, Me) Then
                                MsgBox("Data change saved.", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtRemarks.Enabled = True And (Me.txtReason.Text.Trim <> "" Or Me.txtRemarks.Text.Trim <> "" Or _
                    (Me.dgvItem.Rows.Count > 1 And Me.dgvItem.AllowUserToAddRows = True) Or (Me.dgvItem.Rows.Count > 0 _
                    And Me.dgvItem.AllowUserToAddRows = False)) Then
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirmation")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        If obj.CariDetail(Me.txtPrimary.Text, Me.dgvItem, Me) = True Then
                            Me.SalinKeField()
                        End If
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
            Me.txtDocumentNo.Focus()
            Me.txtDocumentNo.Text = Me.obj.GetNoBaru(Me.dtpDate.Value.Date)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmTrans_SchDownTime_Cari
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.primary = -1 Then
                If Me.obj.CariDetail(hj.primary, Me.dgvItem, Me) = True Then
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
            If Me.txtPrimary.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtPrimary.Text <> "-1" Then
                If Me.obj.ApakahBolehEdit(Me.txtPrimary.Text, Me.UserCode) = False Then
                    Me.btnFind.Focus()
                Else
                    If Me.obj.CariDetail(Me.txtPrimary.Text, Me.dgvItem, Me) = True Then
                        Me.SalinKeField()
                        Me.baru = False
                        Me.AktifkanDetail()
                        Me.txtDocumentNo.Focus()
                    Else
                        MsgBox("The specified data not found, please find the respected data.", MsgBoxStyle.Information, "Information")
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtPrimary.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtPrimary.Text <> "-1" Then
                If Me.obj.ApakahBolehHapus(Me.txtPrimary.Text, Me.userCode, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Are you sure?", MsgBoxStyle.YesNo, "Confirmation")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.obj.Hapus(Me.txtPrimary.Text, Me.userCode, Me.dgvItem, Me, True) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted.", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub
 
    Private Sub txtKode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtReason.GotFocus, txtRemarks.GotFocus, txtDocumentNo.GotFocus
        Dim txt As TextBox
        Try
            txt = TryCast(sender, TextBox)
            If txt.ReadOnly = False Then
                txt.BackColor = backColorAll
            Else
                txt.BackColor = Me.txtReadonly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtKode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtReason.Leave, txtRemarks.Leave, txtDocumentNo.Leave
        Dim txt As TextBox
        Try
            txt = TryCast(sender, TextBox)
            txt.Text = txt.Text.Replace("'", "").Replace("#", "").Replace("%", "").Replace("*", "").Trim
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtKode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtReason.LostFocus, txtRemarks.LostFocus, txtDocumentNo.LostFocus
        Dim txt As TextBox
        Try
            txt = TryCast(sender, TextBox)
            If txt.ReadOnly = False And txt.Enabled = True Then
                txt.BackColor = Me.txtPrimary.BackColor
            Else
                txt.BackColor = Me.txtReadonly.BackColor
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            If Me.objNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Schedule_Downtime) = True Then
                Me.BackgroundWorker1.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(10)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.baru = False Then
                    If Me.obj.Date_.Year = Me.dtpDate.Value.Year And Me.obj.Date_.Month = Me.dtpDate.Value.Month Then
                        Me.txtDocumentNo.Text = Me.obj.DocumentNo
                    Else
                        Me.txtDocumentNo.Text = Me.obj.GetNoBaru(Me.dtpDate.Value.Date)
                    End If
                Else
                    Me.txtDocumentNo.Text = Me.obj.GetNoBaru(Me.dtpDate.Value.Date)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex = Me.colWU_Code.Index Or e.ColumnIndex = Me.colWU_StartDate.Index Or e.ColumnIndex = Me.colWU_StartTime.Index Or _
                    e.ColumnIndex = Me.colWU_EndDate.Index Or e.ColumnIndex = Me.colWU_EndTime.Index Or e.ColumnIndex = Me.colWU_Remarks.Index Then

                ''tidak isi yang lain sebelum isi work unit
                If e.ColumnIndex <> Me.colWU_Code.Index And Me.dgvItem.Item(Me.colWU_Primary.Index, e.RowIndex).Value = "-1" Then
                    If Me.charKeyDown <> Keys.D0 And Me.charKeyDown <> Keys.NumPad0 Then
                        MsgBox("Input work unit first.", MsgBoxStyle.Information, "Information")
                    Else
                        Me.charKeyDown = Keys.A
                    End If
                    e.Cancel = True
                    Exit Sub
                End If

                ''hanya betulkan start time jika start date sudah diisi
                If e.ColumnIndex = Me.colWU_StartTime.Index And IsDate(Me.dgvItem.Item(Me.colWU_StartDate.Index, e.RowIndex).Value) = False Then
                    If Me.charKeyDown <> Keys.D0 And Me.charKeyDown <> Keys.NumPad0 Then
                        MsgBox("Define start date first.", MsgBoxStyle.Information, "Information")
                    Else
                        Me.charKeyDown = Keys.A
                    End If
                    e.Cancel = True
                    Exit Sub
                End If

                ''hanya betulkan end time jika end date sudah diisi
                If e.ColumnIndex = Me.colWU_EndTime.Index And IsDate(Me.dgvItem.Item(Me.colWU_EndDate.Index, e.RowIndex).Value) = False Then
                    If Me.charKeyDown <> Keys.D0 And Me.charKeyDown <> Keys.NumPad0 Then
                        MsgBox("Define end date first.", MsgBoxStyle.Information, "Information")
                    Else
                        Me.charKeyDown = Keys.A
                    End If
                    e.Cancel = True
                    Exit Sub
                End If

                ''hanya mulai wu code jika yang lain sudah komplit
                If IsOthersValid(e.RowIndex) = False Then
                    If Me.charKeyDown <> Keys.D0 And Me.charKeyDown <> Keys.NumPad0 Then
                        MsgBox("Please complete start date and end date for rows above.", MsgBoxStyle.Information, "Information")
                    Else
                        Me.charKeyDown = Keys.A
                    End If
                    e.Cancel = True
                    Exit Sub
                End If
                Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function IsOthersValid(ByVal vCurrentRow As Integer) As Boolean
        Try
            For Each r As DataGridViewRow In Me.dgvItem.Rows
                If r.Index < vCurrentRow Then
                    If r.Cells(Me.colWU_Primary.Index).Value <> "-1" Then
                        If IsDate(r.Cells(Me.colWU_StartDate.Index).Value) = False Or IsDate(r.Cells(Me.colWU_EndDate.Index).Value) = False Then
                            Return False
                        End If
                    End If
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function 

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim frmWU As New Core_Module.FrmMasterWorkUnit_Cari
        Try
            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If String.IsNullOrEmpty(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.colWU_Code.Index Then
                        If Me.objWorkUnit.ApakahSudahAda(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, True) = True Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                            Me.dgvItem.Item(Me.colWU_Primary.Index, e.RowIndex).Value = Me.objWorkUnit.GetPrimary(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                            Me.dgvItem.Item(Me.colWU_Desc.Index, e.RowIndex).Value = Me.objWorkUnit.GetDesc(Me.dgvItem.Item(Me.colWU_Primary.Index, e.RowIndex).Value)
                        Else
                            frmWU.txtCari.Text = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString
                            frmWU.ShowDialog()
                            If frmWU.PrimaryWorkUnit > 0 Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.objWorkUnit.GetCode(frmWU.PrimaryWorkUnit)
                                Me.dgvItem.Item(Me.colWU_Primary.Index, e.RowIndex).Value = frmWU.PrimaryWorkUnit
                                Me.dgvItem.Item(Me.colWU_Desc.Index, e.RowIndex).Value = Me.objWorkUnit.GetDesc(frmWU.PrimaryWorkUnit)
                            Else
                                If String.IsNullOrEmpty(Me.buffer) Then
                                    Me.rowToDelete = e.RowIndex
                                    Me.BackgroundWorker2.RunWorkerAsync()
                                Else
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                            End If
                        End If
                    End If
                    If e.ColumnIndex = Me.colWU_StartDate.Index Then
                        If IsDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        ''jika sudah ada end date, tidak boleh melebihi tanggal end
                        If IsDate(Me.dgvItem.Item(Me.colWU_EndDate.Index, e.RowIndex).Value) Then
                            If CDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > CDate(Me.dgvItem.Item(Me.colWU_EndDate.Index, e.RowIndex).Value) Then
                                MsgBox("Start date cannot greater than end date.", MsgBoxStyle.Information, "Information")
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                        End If
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")
                    End If
                    If e.ColumnIndex = Me.colWU_EndDate.Index Then
                        If IsDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        ''jika sudah ada start date, tidak boleh kurang dari tanggal start
                        If IsDate(Me.dgvItem.Item(Me.colWU_StartDate.Index, e.RowIndex).Value) Then
                            If CDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) < CDate(Me.dgvItem.Item(Me.colWU_StartDate.Index, e.RowIndex).Value) Then
                                MsgBox("End date cannot less than start date.", MsgBoxStyle.Information, "Information")
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                        End If
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")
                    End If
                    If e.ColumnIndex = Me.colWU_StartTime.Index Then
                        If Me.objCd.Is_Time_String(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        If Me.objCd.Convert_ToMinutes(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > 1439 Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        ''jika sudah ada start date dan end date data haus valid
                        If IsDate(Me.dgvItem.Item(Me.colWU_EndDate.Index, e.RowIndex).Value) And IsDate(Me.dgvItem.Item(Me.colWU_StartDate.Index, e.RowIndex).Value) Then
                            If CDate(Me.dgvItem.Item(Me.colWU_StartDate.Index, e.RowIndex).Value.ToString & " " & Me.dgvItem.Item(Me.colWU_StartTime.Index, e.RowIndex).Value.ToString.Replace(".", ":") & ":00") >= _
                                         CDate(Me.dgvItem.Item(Me.colWU_EndDate.Index, e.RowIndex).Value.ToString & " " & Me.dgvItem.Item(Me.colWU_EndTime.Index, e.RowIndex).Value.ToString.Replace(".", ":") & ":00") Then
                                MsgBox("Start datetime cannot greater or equals than end datetime.", MsgBoxStyle.Information, "Information")
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                        End If
                    End If
                    If e.ColumnIndex = Me.colWU_EndTime.Index Then
                        If Me.objCd.Is_Time_String(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        If Me.objCd.Convert_ToMinutes(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > 1439 Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        ''jika sudah ada start date dan end date, data harus valid
                        If IsDate(Me.dgvItem.Item(Me.colWU_StartDate.Index, e.RowIndex).Value) And IsDate(Me.dgvItem.Item(Me.colWU_EndDate.Index, e.RowIndex).Value) Then
                            If CDate(Me.dgvItem.Item(Me.colWU_StartDate.Index, e.RowIndex).Value.ToString & " " & Me.dgvItem.Item(Me.colWU_StartTime.Index, e.RowIndex).Value.ToString.Replace(".", ":") & ":00") >= _
                                         CDate(Me.dgvItem.Item(Me.colWU_EndDate.Index, e.RowIndex).Value.ToString & " " & Me.dgvItem.Item(Me.colWU_EndTime.Index, e.RowIndex).Value.ToString.Replace(".", ":") & ":00") Then
                                MsgBox("End datetime cannot less or equals than start datetime.", MsgBoxStyle.Information, "Information")
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                        End If
                    End If
                    If e.ColumnIndex = Me.colWU_Remarks.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Replace("'", "").Replace("%", "").Replace("*", "").Trim
                    End If
                Else
                    If e.ColumnIndex <> Me.colWU_Remarks.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Else
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    End If
                End If
            Else
                If e.ColumnIndex <> Me.colWU_Remarks.Index Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Dim charKeyDown As System.Windows.Forms.Keys = Keys.A
    Private Sub dgvItem_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvItem.KeyDown
        Try
            Me.charKeyDown = e.KeyCode
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Dim r As DataGridViewRow
        Try
            r = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
            r.Cells(Me.colWU_No.Index).Value = Me.dgvItem.Rows.Count
            r.Cells(Me.colWU_Primary.Index).Value = -1
            r.Cells(Me.colWU_Remarks.Index).Value = ""
            r.Cells(Me.colWU_StartTime.Index).Value = "00.00"
            r.Cells(Me.colWU_EndTime.Index).Value = "23.59"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItem.RowsRemoved
        Try
            Me.objCd.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            Me.objCd.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(10)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.dgvItem.Rows.RemoveAt(Me.rowToDelete)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbWorkUnit_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbWorkUnit.LinkClicked
        Try
            If Me.btnOK.Enabled = True Then
                Me.ContextMenuStrip1.Show(Me.LlbWorkUnit, (Me.ContextMenuStrip1.Width - Me.LlbWorkUnit.Width) * -1, (Me.ContextMenuStrip1.Height) * -1)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ApplyToAllWokUnitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ApplyToAllWokUnitToolStripMenuItem.Click
        Try
            If MsgBox("Are you sure to add all work unit?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                Me.ApplyBatchWorkUnit(-1)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ApplyToWorkUnitGroupToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ApplyToWorkUnitGroupToolStripMenuItem.Click
        Dim f As New Core_Module.FrmMasterWorkUnitGroup_Cari
        Dim objG As New Core_Module.ClsMasterWorkUnitGroup
        Try
            f.ShowDialog()
            If f.PrimaryWorkUnitGroup > 0 Then
                If MsgBox("Are you sure to add all work unit of " & objG.GetDesc(f.PrimaryWorkUnitGroup) & " group?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                    Me.ApplyBatchWorkUnit(f.PrimaryWorkUnitGroup)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ApplyBatchWorkUnit(ByVal vPrimaryWUGroup As Double)
        Dim d As New DseTrans_SchDownTimeTableAdapters.tblmaster_workunitTableAdapter
        Dim t As New DseTrans_SchDownTime.tblmaster_workunitDataTable
        Dim r As DseTrans_SchDownTime.tblmaster_workunitRow
        Dim countAdded As Integer = 0
        Dim dr As DataGridViewRow
        Dim forceAdd As Boolean = False
        Dim confirmed As Boolean = False
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            Me.Cursor = Cursors.WaitCursor
            If vPrimaryWUGroup = -1 Then
                d.Fill(t)
            Else
                d.FillBy_Group(t, vPrimaryWUGroup)
            End If
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AllowUserToAddRows = False
            For Each r In t
                If IsExistWorkUnit(r.primary) = True And confirmed = False Then
                    If MsgBox("Some work unit already exist, do you want to add it too?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                        forceAdd = True
                    End If
                    confirmed = True
                End If
                If (Me.IsExistWorkUnit(r.primary) = False And forceAdd = False) Or (forceAdd = True) Then
                    Me.dgvItem.Rows.Add()
                    dr = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    dr.Cells(Me.colWU_Code.Index).Value = r.work_unit_code
                    dr.Cells(Me.colWU_Desc.Index).Value = r.description
                    dr.Cells(Me.colWU_EndTime.Index).Value = "23.59"
                    dr.Cells(Me.colWU_No.Index).Value = Me.dgvItem.Rows.Count
                    dr.Cells(Me.colWU_Primary.Index).Value = r.primary
                    dr.Cells(Me.colWU_Remarks.Index).Value = ""
                    dr.Cells(Me.colWU_StartTime.Index).Value = "00.00"
                    countAdded += 1
                End If
            Next
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItem.AllowUserToAddRows = True
            If countAdded > 0 Then
                MsgBox("There is " & countAdded.ToString & " work unit added to item list, please fix the schedule datetime.", MsgBoxStyle.Information, "Information")
            Else
                MsgBox("There is no work unit added, cause all of work unit already exist.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Function IsExistWorkUnit(ByVal vPrimaryWorkUnit As Double) As Boolean
        Try
            For Each row As DataGridViewRow In Me.dgvItem.Rows
                If row.Cells(Me.colWU_Primary.Index).Value = vPrimaryWorkUnit Then
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function
End Class
