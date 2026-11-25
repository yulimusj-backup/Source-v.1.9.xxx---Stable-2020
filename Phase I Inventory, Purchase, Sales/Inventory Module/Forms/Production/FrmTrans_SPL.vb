Option Explicit On
Option Strict Off

Public Class FrmTrans_SPL
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Public mTblLog As New DseTrans_SPL.tbltrans_spl_item_logDataTable
    Public mRowLog As DseTrans_SPL.tbltrans_spl_item_logRow

   
    Dim objCD As New Core_Module.ClsCekData
    Dim objSPL As New ClsTrans_SPL
    Dim objWS As New Core_Module.ClsMasterWorkShift
    Dim objWU As New Core_Module.ClsMasterWorkUnit
    Dim EraseRowIndex As Integer = 0
    Dim baru As Boolean = False
    Dim bufferWS As String = ""
    Dim bufferWU As String = ""

    Dim dgvToRemove As Integer = 0
    Dim duration As Double
    Dim delaybeforebreak As Double



    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbWorkUnit.LinkClicked
        Try
            Me.ContextMenuStrip1.Show(Me.LlbWorkUnit, 0, 20)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try

            'Me.objCD.SetBackColor(Me.GroupBox4.Controls, Me.txtPrimary.BackColor, False)
            Me.objCD.SetBackColor(Me.gboDetail.Controls, Me.txtPrimary.BackColor, False)
            Me.objCD.SetBackColor(Me.gboItem.Controls, Me.txtPrimary.BackColor, False)

            Core_Module.ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, True, True, True)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, True, True, True, True)
            Core_Module.ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, True, True)

          
            Me.LlbWorkUnit.Enabled = True

            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.dgvWorkshift.AllowUserToAddRows = True
            Me.dgvWorkUnit.AllowUserToAddRows = True
            Me.dgvWorkshift.AllowUserToDeleteRows = True
            Me.dgvWorkUnit.AllowUserToDeleteRows = True
            Me.dgvWorkshift.ReadOnly = False
            Me.dgvWorkUnit.ReadOnly = False

            Me.txtDocumentNo.ReadOnly = True
            Me.txtDocumentNo.BackColor = Me.txtReadonly.BackColor

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Core_Module.ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False)
            Core_Module.ModCekData.KendalikanControl(Me.Controls, False, False, True, True, True, False, True, False)

            
            Me.LlbWorkUnit.Enabled = False

            Me.CancelButton = Me.btnExit
            Me.dgvWorkshift.AllowUserToAddRows = False
            Me.dgvWorkUnit.AllowUserToAddRows = False
            Me.dgvWorkshift.AllowUserToDeleteRows = False
            Me.dgvWorkUnit.AllowUserToDeleteRows = False

            Me.dgvWorkshift.ReadOnly = True
            Me.dgvWorkUnit.ReadOnly = True

            Me.dgvWorkshift.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvWorkUnit.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.objCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, True)
            Me.objCD.SetBackColor(Me.gboItem.Controls, Me.BackColor, True)
            'Me.dgvWorkshift.Rows.Clear()
            'Me.dgvWorkUnit.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Core_Module.ModCekData.KosongkanControl(Me.gboDetail.Controls, True, False, False)

            Me.txtPrimary.Text = "-1"

            Me.txtDocumentNo.Clear()
            Me.txtJobDesc.Clear()
            Me.dtpDate.Value = Today
            Me.DtpFrom.Value = Today
            Me.DtpTo.Value = Today
            Me.txtRemarks.Text = ""

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "

            Me.dgvWorkshift.ReadOnly = True
            Me.dgvWorkUnit.ReadOnly = True
            Me.dgvWorkshift.AllowUserToDeleteRows = False
            Me.dgvWorkUnit.AllowUserToDeleteRows = False

            Me.dgvWorkshift.Rows.Clear()
            Me.dgvWorkUnit.Rows.Clear()
            
            If baru = False Then
                Me.dgvWorkshift.AllowUserToAddRows = False
                Me.dgvWorkUnit.AllowUserToAddRows = False
            End If
           
            Me.mTblLog.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim rowDGV As DataGridViewRow
      
        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("Document No")
            KoleksiTextBox.Add(Me.txtDocumentNo)
            KoleksiNamaData.Add("Reason")
            KoleksiTextBox.Add(Me.txtJobDesc)
            If Core_Module.ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            If Me.dgvWorkshift.Rows.Count <= 0 Then
                MsgBox("Please input item Work Shift", MsgBoxStyle.Information, "Information")
                Me.dgvWorkshift.Focus()
                Return False
            End If

            If Me.dgvWorkUnit.Rows.Count <= 0 Then
                MsgBox("Please input item Work Unit", MsgBoxStyle.Information, "Information")
                Me.dgvWorkUnit.Focus()
                Return False
            End If

            'cek data yang salah, double dalam kode
            If Me.baru = False Then
                If Me.objSPL.GetTanggal(Me.objSPL.Primary).Month = Me.dtpDate.Value.Month And _
                Me.objSPL.GetTanggal(Me.objSPL.Primary).Year = Me.dtpDate.Value.Year Then
                    Me.txtDocumentNo.Text = Me.objSPL.NoDocument
                Else
                    If MsgBox("You just changed transaction date, this will change your transaction " _
                                    + "number and there's possibility your transaction no cannot be change back. " _
                                    + "Continue and accept this change?", _
                                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                        Me.dtpDate.Value = Me.objSPL.GetTanggal(Me.txtPrimary.Text)
                        Me.txtDocumentNo.Text = Me.objSPL.GetNoDokumen(Me.txtPrimary.Text)
                        Return False
                    End If
                End If
            End If

            If Me.baru = True Then

                If Me.objSPL.ApakahSudahAda(Me.txtDocumentNo.Text) = True Then
                    MsgBox("Duplication in No Dokumen, please change your code", _
                    MsgBoxStyle.Information, "Information")
                    Me.txtDocumentNo.Focus()
                    Return False
                End If
            Else
                Dim TipeCodeLama As String = Me.objSPL.GetNoDokumen(Me.txtPrimary.Text)
                If TipeCodeLama <> Me.txtDocumentNo.Text Then
                    If Me.objSPL.ApakahSudahAda(Me.txtDocumentNo.Text) = True Then
                        MsgBox("Duplication in No Dokumen, please change your code", _
                                            MsgBoxStyle.Information, "Information")
                        Me.txtDocumentNo.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.objSPL.GetTimeStamp(Me.txtPrimary.Text)
                If Me.objSPL.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.objSPL.CariDetail(Me.objSPL.Primary, Me.dgvWorkshift, Me.dgvWorkUnit, Me)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If
            If Me.DtpFrom.Value > Me.DtpTo.Value Then
                MsgBox("Strarting period mustbe smaller than ending period", MsgBoxStyle.Information, "Information")
                Return False
            End If
            'cek ada workshift dan tidak ada yg unchek pada EOT dan OT

            rowDGV = Me.dgvWorkshift.Rows(0)
            If rowDGV.Cells(Me.colDesc.Index).Value = "" Then
                MsgBox("There must be one Work Shift", MsgBoxStyle.Information, "Information")
                Return False
            Else


                For ij As Integer = 0 To Me.dgvWorkshift.Rows.Count - 2
                    rowDGV = Me.dgvWorkshift.Rows(ij)

                    If rowDGV.Cells(Me.colEarlyOT.Index).Value = False And rowDGV.Cells(Me.colOT.Index).Value = False Then
                        MsgBox("Please check again column Work Shift Early OT Duration or OT Duration ", MsgBoxStyle.Information, "Information")
                        Return False
                    Else
                        'cek bila sudah cheklis tapi durasi EOT 0
                        If rowDGV.Cells(Me.colEarlyOT.Index).Value = True Then
                            If rowDGV.Cells(Me.colEarlyOTDuration.Index).Value = "00.00" Then
                                MsgBox("Duration can't be empty", MsgBoxStyle.Information, "Information")
                                Return False
                            Else
                                'cek bila ada Early break
                                If rowDGV.Cells(Me.ColEarlyBreak.Index).Value <> "00.00" Then
                                    If rowDGV.Cells(Me.ColDelayBeforeEarlyBreak.Index).Value = "00.00" Then
                                        MsgBox("Early Delay Before Break can't be empty", MsgBoxStyle.Information, "Information")
                                        Return False
                                    Else
                                        'cek bila DBB >= Durasi
                                        duration = CDbl(rowDGV.Cells(Me.colEarlyOTDuration.Index).Value)
                                        delaybeforebreak = CDbl(rowDGV.Cells(Me.ColDelayBeforeEarlyBreak.Index).Value)
                                        If delaybeforebreak >= duration Then
                                            MsgBox("Early Delay Before Break shouldn't be greater equal to the Early OT Duration", MsgBoxStyle.Information, "Information")
                                            'Me.bufferWS = Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value
                                            Return False
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        'cek bila sudah cheklis tapi durasi OT 0
                        If rowDGV.Cells(Me.colOT.Index).Value = True Then
                            If rowDGV.Cells(Me.colOTDuration.Index).Value = "00.00" Then
                                MsgBox("OT Duration can not be empty", MsgBoxStyle.Information, "Information")
                                Return False
                            Else
                                'cek bila ada break
                                If rowDGV.Cells(Me.ColBreak.Index).Value <> "00.00" Then
                                    If rowDGV.Cells(Me.ColDelayBeforeBreak.Index).Value = "00.00" Then
                                        MsgBox("Delay Before Break can't be empty", MsgBoxStyle.Information, "Information")
                                        Return False
                                    Else
                                        'cek bila DBB >= Durasi
                                        duration = CDbl(rowDGV.Cells(Me.colOTDuration.Index).Value)
                                        delaybeforebreak = CDbl(rowDGV.Cells(Me.ColDelayBeforeBreak.Index).Value)
                                        If delaybeforebreak >= duration Then
                                            MsgBox("Delay Before Break shouldn't be greater equal to the OT Duration", MsgBoxStyle.Information, "Information")
                                            'Me.bufferWS = Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value
                                            Return False
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
                End If

            'cek minimal ada satu workunit
            For ij As Integer = 0 To Me.dgvWorkUnit.Rows.Count - 1
                rowDGV = Me.dgvWorkUnit.Rows(ij)
                If rowDGV.Cells(Me.colWU_Desc.Index).Value = "" Then
                    MsgBox("There must be one Work Unit", MsgBoxStyle.Information, "Information")
                    Return False
                Else
                    Exit For
                End If
            Next

            ' cek bila terjadi duplikasi SPL pada workunit yang sama
            If baru = True Then
                If objSPL.ApakahSPLSudahdibuat(Me.dgvWorkUnit, Me) = False Then
                    Return False
                End If
            End If

            'buat log
            Me.objSPL.CreateLog(Me.dgvWorkUnit, Me)


            If Me.mTblLog.Count = 0 Then
                MsgBox("Please check again Allocated Workshift", MsgBoxStyle.Information, "Information")
                Return False
            End If
            'cek allocated workunit
            If objSPL.ApakahSudahDiallocate(Me.dgvWorkUnit, Me) = False Then
                Return False
            End If

            'cek sudah ada bundling
            If Me.baru = False Then
                'If objSPL.GetStatusAuth(Me.txtPrimary.Text) = True Then
                '    MsgBox("SPL can't be changed or deleted because it's already authorize", MsgBoxStyle.Information, "Information")
                '    Return False
                'End If

                If objSPL.apakahSudahBundling(Me.dgvWorkUnit, Me) = False Then
                    MsgBox("SPL can't be changed or deleted because it is used in the transaction bundling", MsgBoxStyle.Information, "Information")
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
            With Me.objSPL
                .NoDocument = Me.txtDocumentNo.Text
                .Tanggal = Me.dtpDate.Value
                .JobDesc = Me.txtJobDesc.Text
                .Remarks = Me.txtRemarks.Text
                .FromDate = Me.DtpFrom.Value
                .ToDate = Me.DtpTo.Value
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub SalinKeField()
        Try
            With Me.objSPL

                Me.txtPrimary.Text = .Primary
                Me.txtDocumentNo.Text = .NoDocument
                Me.dtpDate.Value = .Tanggal
                Me.txtJobDesc.Text = .JobDesc
                Me.txtRemarks.Text = .Remarks
                Me.DtpFrom.Value = .FromDate
                Me.DtpTo.Value = .ToDate

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmTrans_SPL_Cari
        Try
            ij.BackColorAll = Me.BackColorAll
            ij.ShowDialog()
            If Not ij.PrimaryPP = -1 Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                Me.KosongkanField()
                If Me.objSPL.CariDetail(ij.PrimaryPP, Me.dgvWorkshift, Me.dgvWorkUnit, Me) = True Then
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
            ij.Dispose()
        End Try
    End Sub

    Private Sub FrmTrans_SPL_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.txtDocumentNo.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objSPL.SimpanBaru(Me.txtDocumentNo.Text, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SPL, _
                            Me.dtpDate.Value, Me.dgvWorkshift, Me.dgvWorkUnit, UserCode, Me, mTblLog) = True Then
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
                            If Me.objSPL.SimpanPerubahan(Me.txtDocumentNo.Text, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SPL, _
                            Me.dtpDate.Value, Me.dgvWorkshift, Me.dgvWorkUnit, UserCode, Me, mTblLog) = True Then
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

    Private Sub FrmTrans_SPL_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            Me.baru = True
            Me.NonAktifkanDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()

            Me.dtpDate.Focus()
            Me.objSPL.ResetObject(Me.dgvWorkshift)
            Me.objSPL.ResetObject(Me.dgvWorkUnit)
            Me.txtDocumentNo.Text = Me.objSPL.GetNoBaru(Me.dtpDate.Value)

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDocumentNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDocumentNo.GotFocus, txtJobDesc.GotFocus, _
    txtRemarks.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.txtReadonly.BackColor
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

    Private Sub txtDocumentNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDocumentNo.LostFocus, txtJobDesc.LostFocus, _
    txtRemarks.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.txtReadonly.BackColor
                        Else
                            Txt.BackColor = Me.txtPrimary.BackColor
                        End If
                    End If
                End If
            End If
            txtJobDesc.Text = txtJobDesc.Text.ToString.Replace("'", "").Replace("#", "").Replace("%", "").Replace("*", "")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtRemarks.Text = "" And Me.dgvWorkshift.RowCount = 0 And _
            Me.txtDocumentNo.Text = "" And Me.txtJobDesc.Text = "" And Me.dgvWorkUnit.RowCount = 0 Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
            Else
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        baru = False
                        Me.KosongkanField()
                    Else
                        Me.KosongkanField()
                        Me.objSPL.CariDetail(Me.objSPL.Primary, Me.dgvWorkshift, Me.dgvWorkUnit, Me)
                        Me.SalinKeField()

                    End If
                Else
                    Exit Sub
                End If
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.baru = False Then
                    If Me.objSPL.Tanggal.Month = Me.dtpDate.Value.Month And _
                    Me.objSPL.Tanggal.Year = Me.dtpDate.Value.Year Then
                        Me.txtDocumentNo.Text = Me.objSPL.NoDocument
                    Else
                        Me.txtDocumentNo.Text = Me.objSPL.GetNoBaru(Me.dtpDate.Value)
                    End If
                Else
                    Me.txtDocumentNo.Text = Me.objSPL.GetNoBaru(Me.dtpDate.Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            Me.SalinKeObjek()
            If baru = True Then
                If objSPL.ApakahDataBenar(baru) = True Then
                    If Me.ApakahDataBenar() = True Then
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            baru = False
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objSPL.SimpanBaru(Me.txtDocumentNo.Text, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SPL, _
                            Me.dtpDate.Value, Me.dgvWorkshift, Me.dgvWorkUnit, Me.UserCode, Me, mTblLog) Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            End If
                        End If
                    End If
                End If
            Else
                If Me.ApakahDataBenar = True Then
                    If objSPL.ApakahBolehEdit(Me.txtPrimary.Text, UserCode, True, True) = True Then
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objSPL.SimpanPerubahan(Me.txtDocumentNo.Text, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SPL, _
                            Me.dtpDate.Value, Me.dgvWorkshift, Me.dgvWorkUnit, Me.UserCode, Me, mTblLog) Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
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
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If
            
            If Me.txtDocumentNo.Text <> "" Then
                If Me.objSPL.ApakahBolehHapus(Me.txtPrimary.Text, True, UserCode, True) = False Then
                    Me.btnFind.Focus()
                Else
                    If objSPL.GetStatusAuth(txtPrimary.Text) = True Then
                        MsgBox("SPL can't be changed or deleted because it's already authorize", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    Else
                        hasil = MsgBox("Delete this data ?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objSPL.Hapus(Me.txtPrimary.Text, UserCode, True) = True Then
                                Me.KosongkanField()
                                Me.btnNew.Focus()
                                MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.txtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtDocumentNo.Text <> "" Then
                If Me.objSPL.ApakahBolehEdit(Me.objSPL.Primary, UserCode, True, True) = True Then
                    'Me.KosongkanField()
                    Me.baru = False
                    Me.KosongkanField()
                    Me.objSPL.CariDetail(Me.objSPL.Primary, Me.dgvWorkshift, Me.dgvWorkUnit, Me)
                    Me.SalinKeField()
                    Me.AktifkanDetail()

                    Me.dgvWorkshift.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWorkshift_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvWorkshift.CellBeginEdit
        Try

            If e.ColumnIndex = Me.colNo.Index Or e.ColumnIndex = Me.colDesc.Index Then
                e.Cancel = True
            Else
                If e.ColumnIndex <> Me.colShiftCode.Index Then
                    If Me.dgvWorkshift.Item(Me.colPrimaryWorkshift.Index, e.RowIndex).Value = "-1" Then
                        e.Cancel = True
                    End If
                    If e.ColumnIndex = Me.colShiftCode.Index Or e.ColumnIndex = Me.colDesc.Index Or e.ColumnIndex = Me.colEarlyOTDuration.Index Or e.ColumnIndex = Me.ColEarlyBreak.Index Or e.ColumnIndex = Me.colEarlyDelayDuration.Index Or e.ColumnIndex = Me.ColDelayBeforeEarlyBreak.Index Then
                        If Me.dgvWorkshift.Item(Me.colEarlyOT.Index, e.RowIndex).Value = False Then
                            e.Cancel = True
                        Else
                            Me.bufferWS = Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value
                        End If
                    End If
                    If e.ColumnIndex = Me.colShiftCode.Index Or e.ColumnIndex = Me.colDesc.Index Or e.ColumnIndex = Me.colOTDuration.Index Or e.ColumnIndex = Me.ColBreak.Index Or e.ColumnIndex = Me.colDelayDuration.Index Or e.ColumnIndex = Me.ColDelayBeforeBreak.Index Then
                        If Me.dgvWorkshift.Item(Me.colOT.Index, e.RowIndex).Value = False Then
                            e.Cancel = True
                        Else
                            Me.bufferWS = Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value
                        End If
                    End If
                Else
                    Me.bufferWS = Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWorkshift_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvWorkshift.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvWorkshift.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryWS As Double = -1
        Dim RowIdx As Integer = e.RowIndex

        Try
            With row

                If IsNothing(Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value) = False Then

                    If e.ColumnIndex = Me.colShiftCode.Index Then
                        Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.objWS.ApakahSudahAda(Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value, Core_Module.ClsMasterWorkShift.EnumTipeWorkshift.Production) = True Then
                            PrimaryWS = Me.objWS.GetPrimary(Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value, Core_Module.ClsMasterWorkShift.EnumTipeWorkshift.Production)
                            index = Me.GetIndexItemWS(PrimaryWS, e.RowIndex)
                            If index < 0 Then
                                Me.dgvWorkshift.Item(Me.colPrimaryWorkshift.Index, e.RowIndex).Value = PrimaryWS
                                Me.dgvWorkshift.Item(Me.colShiftCode.Index, e.RowIndex).Value = Me.objWS.GetCodeByPrimary(PrimaryWS)
                                Me.dgvWorkshift.Item(Me.colDesc.Index, e.RowIndex).Value = Me.objWS.GetDescription(PrimaryWS)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvWorkshift.Item(Me.colShiftCode.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.bufferWS) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.dgvToRemove = 0
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.bufferWS
                                    End If
                                    Me.objCD.SetNumber(Me.dgvWorkshift)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmMasterWorkShift_Cari(Core_Module.ClsMasterWorkShift.EnumTipeWorkshift.Production)
                            ij.txtCari.Text = Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryWorkshift <> -1 Then

                                index = Me.GetIndexItemWS(ij.PrimaryWorkshift, e.RowIndex)
                                If index < 0 Then
                                    Me.dgvWorkshift.Item(Me.colPrimaryWorkshift.Index, e.RowIndex).Value = ij.PrimaryWorkshift
                                    Me.dgvWorkshift.Item(Me.colShiftCode.Index, e.RowIndex).Value = Me.objWS.GetCodeByPrimary(ij.PrimaryWorkshift)
                                    Me.dgvWorkshift.Item(Me.colDesc.Index, e.RowIndex).Value = Me.objWS.GetDescription(ij.PrimaryWorkshift)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.bufferWS) = True Then
                                            Me.dgvToRemove = 0
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.bufferWS
                                        End If
                                        Me.objCD.SetNumber(Me.dgvWorkshift)
                                        Exit Sub
                                    Else
                                        Me.dgvWorkshift.Item(Me.colPrimaryWorkshift.Index, e.RowIndex).Value = ij.PrimaryWorkshift
                                        Me.dgvWorkshift.Item(Me.colShiftCode.Index, e.RowIndex).Value = Me.objWS.GetCodeByPrimary(ij.PrimaryWorkshift)
                                        Me.dgvWorkshift.Item(Me.colDesc.Index, e.RowIndex).Value = Me.objWS.GetDescription(ij.PrimaryWorkshift)

                                    End If
                                End If
                            ElseIf ij.PrimaryWorkshift = -1 Then
                                If Me.dgvWorkshift.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.bufferWS) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.dgvToRemove = 0
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.bufferWS
                                    End If
                                    Me.objCD.SetNumber(Me.dgvWorkshift)
                                    Exit Sub
                                End If
                                If Me.dgvWorkshift.Item(Me.colShiftCode.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.bufferWS) = True Then
                                        Me.dgvToRemove = 0
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.bufferWU
                                    End If
                                    Me.objCD.SetNumber(Me.dgvWorkshift)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.objCD.SetNumber(Me.dgvWorkshift)
                    ElseIf e.ColumnIndex = Me.colShiftCode.Index Then
                        Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.bufferWS
                    ElseIf e.ColumnIndex = Me.colEarlyOT.Index Then
                        If Me.dgvWorkshift.Item(Me.colEarlyOT.Index, e.RowIndex).Value = False Then
                            Me.dgvWorkshift.Item(Me.colEarlyOTDuration.Index, e.RowIndex).Value = "00.00"
                            Me.dgvWorkshift.Item(Me.colEarlyDelayDuration.Index, e.RowIndex).Value = "00.00"
                            Me.dgvWorkshift.Item(Me.ColEarlyBreak.Index, e.RowIndex).Value = "00.00"
                            Me.dgvWorkshift.Item(Me.ColDelayBeforeEarlyBreak.Index, e.RowIndex).Value = "00.00"
                        End If
                    ElseIf e.ColumnIndex = Me.colOT.Index Then
                        If Me.dgvWorkshift.Item(Me.colOT.Index, e.RowIndex).Value = False Then
                            Me.dgvWorkshift.Item(Me.colOTDuration.Index, e.RowIndex).Value = "00.00"
                            Me.dgvWorkshift.Item(Me.colDelayDuration.Index, e.RowIndex).Value = "00.00"
                            Me.dgvWorkshift.Item(Me.ColBreak.Index, e.RowIndex).Value = "00.00"
                            Me.dgvWorkshift.Item(Me.ColDelayBeforeBreak.Index, e.RowIndex).Value = "00.00"
                        End If
                    End If
                    If e.ColumnIndex = Me.colEarlyOTDuration.Index Then
                        If objCD.Is_Time_String(Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Format should be time string: hh.mm", MsgBoxStyle.Information, "Information")
                            Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.bufferWS
                        End If
                    End If
                    If e.ColumnIndex = Me.colEarlyDelayDuration.Index Then
                        If objCD.Is_Time_String(Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Format should be time string: hh.mm", MsgBoxStyle.Information, "Information")
                            Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.bufferWS
                        End If
                    End If
                    If e.ColumnIndex = Me.ColEarlyBreak.Index Then
                        If objCD.Is_Time_String(Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Format should be time string: hh.mm", MsgBoxStyle.Information, "Information")
                            Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.bufferWS
                        End If
                    End If
                    If e.ColumnIndex = Me.ColDelayBeforeEarlyBreak.Index Then
                        If objCD.Is_Time_String(Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Format should be time string: hh.mm", MsgBoxStyle.Information, "Information")
                            Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.bufferWS
                        End If
                        duration = CDbl(Me.dgvWorkshift.Item(Me.colEarlyOTDuration.Index, e.RowIndex).Value)
                        delaybeforebreak = CDbl(Me.dgvWorkshift.Item(Me.ColDelayBeforeEarlyBreak.Index, e.RowIndex).Value)
                        If delaybeforebreak >= duration Then
                            MsgBox("Early Delay Before Break shouldn't be greater than Early OT Duration", MsgBoxStyle.Information, "Information")
                            Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.bufferWS
                        End If
                    End If
                    If e.ColumnIndex = Me.colOTDuration.Index Then
                        If objCD.Is_Time_String(Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Format should be time string: hh.mm", MsgBoxStyle.Information, "Information")
                            Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.bufferWS
                        End If
                    End If

                    If e.ColumnIndex = Me.colDelayDuration.Index Then
                        If objCD.Is_Time_String(Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Format should be time string: hh.mm", MsgBoxStyle.Information, "Information")
                            Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.bufferWS
                        End If
                    End If
                    If e.ColumnIndex = Me.ColBreak.Index Then
                        If objCD.Is_Time_String(Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Format should be time string: hh.mm", MsgBoxStyle.Information, "Information")
                            Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.bufferWS
                        End If
                    End If
                    If e.ColumnIndex = Me.ColDelayBeforeBreak.Index Then
                        If objCD.Is_Time_String(Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Format should be time string: hh.mm", MsgBoxStyle.Information, "Information")
                            Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.bufferWS
                        End If
                        duration = CDbl(Me.dgvWorkshift.Item(Me.colOTDuration.Index, e.RowIndex).Value)
                        delaybeforebreak = CDbl(Me.dgvWorkshift.Item(Me.ColDelayBeforeBreak.Index, e.RowIndex).Value)
                        If delaybeforebreak >= duration Then
                            MsgBox("Delay Before Break shouldn't be greater than OT Duration", MsgBoxStyle.Information, "Information")
                            Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.bufferWS
                        End If
                    End If

                Else
                    Me.dgvWorkshift.Item(e.ColumnIndex, e.RowIndex).Value = Me.bufferWS
                    If String.IsNullOrEmpty(Me.bufferWS) = True Then
                        Me.dgvToRemove = 0
                        Me.EraseRowIndex = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                    Else
                        Me.dgvWorkshift.Item(Me.colEarlyOTDuration.Index, e.RowIndex).Value = Me.bufferWS
                    End If
                    Me.objCD.SetNumber(Me.dgvWorkshift)
                    Exit Sub
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemWS(ByVal vPrimary As Double, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            'pastikan tidak ada code yang sama
            For ij As Integer = 0 To Me.dgvWorkshift.Rows.Count - 2
                RowDGV = Me.dgvWorkshift.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.colPrimaryWorkshift.Index).Value = vPrimary Then
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
    Public Function GetIndexItemWU(ByVal vPrimary As Double, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Try
            'pastikan tidak ada code yang sama
            For ij As Integer = 0 To Me.dgvWorkUnit.Rows.Count - 1
                RowDGV = Me.dgvWorkUnit.Rows(ij)
                If RowDGV.Cells(Me.colWU_Primary.Index).Value <> -1 Then
                    If vRowIdxFocus <> ij Then
                        If RowDGV.Cells(Me.colWU_Primary.Index).Value = vPrimary Then
                            Return Me.dgvWorkUnit.Rows.IndexOf(RowDGV)
                        End If
                    End If
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.dgvToRemove = 0 Then
                Me.dgvWorkshift.Rows.RemoveAt(EraseRowIndex)
                Me.objCD.SetNumber(Me.dgvWorkshift)
            Else
                Me.dgvWorkUnit.Rows.RemoveAt(EraseRowIndex)
                Me.objCD.SetNumber(Me.dgvWorkUnit)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub dgvWorkshift_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvWorkshift.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            Me.dgvWorkshift.TabStop = True
            rowDGV = Me.dgvWorkshift.Rows(Me.dgvWorkshift.Rows.Count - 1)
            With rowDGV
                .Cells(Me.colNo.Index).Value = rowDGV.Index + 1
                .Cells(Me.colShiftCode.Index).Value = ""
                .Cells(Me.colDesc.Index).Value = ""
                .Cells(Me.colEarlyOT.Index).Value = False
                .Cells(Me.colEarlyOTDuration.Index).Value = "00.00"
                .Cells(Me.colEarlyDelayDuration.Index).Value = "00.00"
                .Cells(Me.ColEarlyBreak.Index).Value = "00.00"
                .Cells(Me.ColDelayBeforeEarlyBreak.Index).Value = "00.00"
                .Cells(Me.colOT.Index).Value = False
                .Cells(Me.colOTDuration.Index).Value = "00.00"
                .Cells(Me.colDelayDuration.Index).Value = "00.00"
                .Cells(Me.ColBreak.Index).Value = "00.00"
                .Cells(Me.ColDelayBeforeBreak.Index).Value = "00.00"
                .Cells(Me.colRemarks.Index).Value = ""
                .Cells(Me.colPrimaryWorkshift.Index).Value = "-1"
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            rowDGV = Nothing
        End Try
    End Sub

    Private Sub dgvWorkshift_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvWorkshift.Sorted
        Me.objCD.SetNumber(Me.dgvWorkshift)
    End Sub

    Private Sub dgvWorkshift_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvWorkshift.UserDeletedRow
        Me.objCD.SetNumber(Me.dgvWorkshift)
    End Sub

    Private Sub dgvWorkshift_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvWorkshift.DataError

    End Sub

    Private Sub dgvWorkUnit_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvWorkUnit.CellBeginEdit
        Try
            If Me.dgvWorkUnit.Item(Me.colWU_Primary.Index, e.RowIndex).Value <> -1 Then
                If e.ColumnIndex = Me.colWU_Code.Index Then
                    Me.bufferWU = Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.colWU_Code.Index Then
                    Me.bufferWU = Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWorkUnit_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvWorkUnit.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvWorkUnit.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryWU As Double = -1
        Dim RowIdx As Integer = e.RowIndex

        Try
            With row
                If IsNothing(Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value) = False Then

                    If e.ColumnIndex = Me.colWU_Code.Index Then
                        Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.objWU.ApakahSudahAda(Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryWU = Me.objWU.GetPrimary(Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemWU(PrimaryWU, e.RowIndex)
                            If index < 0 Then
                                Me.dgvWorkUnit.Item(Me.colWU_Primary.Index, e.RowIndex).Value = PrimaryWU
                                Me.dgvWorkUnit.Item(Me.colWU_Code.Index, e.RowIndex).Value = Me.objWU.GetCode(PrimaryWU)
                                Me.dgvWorkUnit.Item(Me.colWU_Desc.Index, e.RowIndex).Value = Me.objWU.GetDesc(PrimaryWU)

                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvWorkUnit.Item(Me.colWU_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.bufferWU) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.dgvToRemove = 1
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.bufferWU
                                    End If
                                    Me.objCD.SetNumber(Me.dgvWorkUnit)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmMasterWorkUnit_Cari
                            ij.txtCari.Text = Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryWorkUnit <> -1 Then

                                index = Me.GetIndexItemWU(ij.PrimaryWorkUnit, e.RowIndex)
                                If index < 0 Then
                                    Me.dgvWorkUnit.Item(Me.colWU_Primary.Index, e.RowIndex).Value = ij.PrimaryWorkUnit
                                    Me.dgvWorkUnit.Item(Me.colWU_Code.Index, e.RowIndex).Value = Me.objWU.GetCode(ij.PrimaryWorkUnit)
                                    Me.dgvWorkUnit.Item(Me.colWU_Desc.Index, e.RowIndex).Value = Me.objWU.GetDesc(ij.PrimaryWorkUnit)

                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.bufferWU) = True Then
                                            Me.dgvToRemove = 1
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.bufferWU
                                        End If
                                        Me.objCD.SetNumber(Me.dgvWorkUnit)
                                        Exit Sub
                                    Else
                                        Me.dgvWorkUnit.Item(Me.colWU_Primary.Index, e.RowIndex).Value = PrimaryWU
                                        Me.dgvWorkUnit.Item(Me.colWU_Code.Index, e.RowIndex).Value = Me.objWU.GetCode(PrimaryWU)
                                        Me.dgvWorkUnit.Item(Me.colWU_Desc.Index, e.RowIndex).Value = Me.objWU.GetDesc(PrimaryWU)

                                    End If
                                End If
                            ElseIf ij.PrimaryWorkUnit = -1 Then
                                If Me.dgvWorkUnit.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.bufferWU) = True Then
                                        Me.EraseRowIndex = e.RowIndex

                                        Me.dgvToRemove = 1
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.bufferWU
                                    End If
                                    Me.objCD.SetNumber(Me.dgvWorkUnit)
                                    Exit Sub
                                End If
                                If Me.dgvWorkUnit.Item(Me.colWU_Code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.bufferWU) = True Then
                                        Me.dgvToRemove = 1
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.bufferWU
                                    End If
                                    Me.objCD.SetNumber(Me.dgvWorkUnit)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.objCD.SetNumber(Me.dgvWorkUnit)
                    End If
                Else
                    Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.bufferWU
                    If String.IsNullOrEmpty(Me.bufferWU) = True Then
                        Me.dgvToRemove = 1
                        Me.EraseRowIndex = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                    Else
                        Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.bufferWU
                    End If
                    Me.objCD.SetNumber(Me.dgvWorkUnit)
                    Exit Sub
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWorkUnit_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvWorkUnit.DataError

    End Sub


    Private Sub dgvWorkUnit_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvWorkUnit.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvWorkUnit.TabStop = True

            rowDGV = Me.dgvWorkUnit.Rows(Me.dgvWorkUnit.RowCount - 1)
            With rowDGV
                .Cells(Me.colNo.Index).Value = rowDGV.Index + 1
                .Cells(Me.colWU_Code.Index).Value = ""
                .Cells(Me.colWU_Desc.Index).Value = ""
                .Cells(Me.colWU_Primary.Index).Value = "-1"
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            rowDGV = Nothing
        End Try
    End Sub

    Private Sub dgvWorkUnit_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvWorkUnit.Sorted
        Me.objCD.SetNumber(Me.dgvWorkUnit)
    End Sub

    Private Sub dgvWorkUnit_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvWorkUnit.UserDeletedRow
        Me.objCD.SetNumber(Me.dgvWorkUnit)
    End Sub

    Private Sub ApplyToAllWokUnitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ApplyToAllWokUnitToolStripMenuItem.Click
        Try
            Me.objSPL.ApplyToAllWokUnit(Me.dgvWorkUnit, Me)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
        

    End Sub

    Private Sub ApplyToWorkUnitGroupToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ApplyToWorkUnitGroupToolStripMenuItem.Click

        Try
            Me.objSPL.ApplyToWorkUnitGroup(Me.dgvWorkUnit, Me)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
       
    End Sub

    

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim frmCetakSPL As New FrmTrans_SPL_Cetak
        Dim frmSPLCetakOption As New FrmTrans_SPL_Cetak_Option
        Try
            If Me.txtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtDocumentNo.Text <> "" Then
                Me.objSPL.CetakanSPL(Me.txtPrimary.Text, frmCetakSPL.crvSPL)
                If objSPL.print = True Then
                    frmCetakSPL.ShowDialog()
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
        
    End Sub
End Class