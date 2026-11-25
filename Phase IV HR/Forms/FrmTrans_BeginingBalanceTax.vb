Option Explicit On
Option Strict Off

Imports System.Windows.Forms, System.Drawing, Core_Module
Public Class FrmTrans_BeginingBalanceTax
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public TBLSetting As New DseTrans_Tax_Beginning_Balance.TblSettingDetailDataTable

    Dim Obj As New ClsTrans_Tax_Beginning_Balance
    Dim ObjItem As New ClsTrans_Tax_Beginning_Balance_Item
    Dim ObjCurr As New clsMasterCurrency
    Dim ObjCD As New ClsCekData
    Dim ObjKary As New Core_Module.ClsMasterKaryawan
    Dim ObjClosing As New AccModule.ClsTransaksiClosing
    Dim ObjTax As New HR.ClsMasterSettingPajak
    Dim DefPrec As Double = 0
    Dim Baru As Boolean
    Dim Buffer As String
    Dim BufferStatus As String
    Dim BufferNPWPNo As String
    Dim BufferCategoryTax As String
    Dim EraseRowIndex As Integer

    Public Enum EnumCriteria As Integer
        TK0
        TK1
        TK2
        TK3
        K0
        K1
        K2
        K3
    End Enum

    Private Sub FrmTrans_BeginingBalanceTax_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim RowSetting As DseTrans_Tax_Beginning_Balance.TblSettingDetailRow
        Try
            Me.Obj = New ClsTrans_Tax_Beginning_Balance
            Me.Baru = True
            Me.KosongkanField()
            Me.NonAktifkanDetail()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            RowSetting = Me.TBLSetting.NewRow
            With RowSetting
                .NIK = 0
                .Dept = 0
                .Position = 0
            End With
            Me.TBLSetting.Rows.Add(RowSetting) 
            Me.ShowIcon = False

            Me.TxtRemarks.Width = Me.SplitContainer1.Panel2.Width - 5 - Me.TxtRemarks.Left
            'Me.TxtDocumentNo.Width = Me.SplitContainer1.Panel1.Width - 5 - Me.TxtDocumentNo.Left
            Me.DtpDate.Width = Me.SplitContainer1.Panel1.Width - 5 - Me.DtpDate.Left
            Me.dtpPeriod.Width = Me.SplitContainer1.Panel1.Width - 5 - Me.dtpPeriod.Left

            For Each Col1 As DataGridViewColumn In Me.dgvItem.Columns
                If Col1.Index = Me.ColGaji.Index Or Col1.Index = Me.ColTunjanganPPH.Index Then 'Or _
                    'Col1.Index = Me.ColTunjanganLainnya.Index Or Col1.Index = Me.ColHonor.Index Or _
                    'Col1.Index = Me.ColPremi.Index Or Col1.Index = Me.ColTantiem.Index Or _
                    'Col1.Index = Me.ColIuranPensiun.Index Or _
                    'Col1.Index = Me.ColNatura.Index Then
                    'do nothing
                Else
                    Col1.Visible = False
                End If
            Next

            Me.ColTunjanganPPH.HeaderText = "PPh"
            Me.ColGaji.HeaderText = "Penghasilan Netto"

            Me.ColNoNPWP.Visible = True
            Me.ColNIK.Visible = True
            Me.ColNamaKaryawan.Visible = True
            Me.ColEmp_No.Visible = True
            Me.ColDepartement.Visible = True
            Me.ColKodeKaryawan.Visible = True
            Me.ColNIK.Visible = True
            Me.ColPosition.Visible = True
            Me.ColStatus.Visible = True
            Me.ColRemarks.Visible = True

            Me.MinimizeBox = True
            Me.SettingVisibleColumn()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim ij As New FrmTrans_Tax_Beginning_Cetak
        Try
            If Me.TxtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtDocumentNo.Text <> "" Then
                ij.mPrimary = Me.TxtPrimary.Text
                ij.mDef_Prec = Me.DefPrec
                ij.ShowDialog()
                ij.Dispose()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.BackColorReset()
            'set dgv
            Me.dgvItem.ReadOnly = False
            Me.dgvItem.AllowUserToDeleteRows = True
            Me.dgvItem.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.Controls, True, True, True, True, True, True, True, True)
            Me.CancelButton = Me.btnExit
            Me.BackColorDefault()
            Me.dgvItem.ReadOnly = True
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.AllowUserToAddRows = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorReset()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Color.White, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorDefault()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtReadOnly.BackColor, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.ObjCD.KosongkanControl(Me.gboDetail.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, True)
            Me.TxtPrimary.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.dgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim CoventionDate As Date = #1/1/1950#
        Try
            Me.dgvItem.AllowUserToAddRows = False

            '1. cek data yang harus diisi
            KoleksiNamaData.Add("Document No")
            KoleksiTextBox.Add(Me.TxtDocumentNo)

            If Me.ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            '2. harus ada data item
            If Me.dgvItem.Rows.Count <= 0 Then
                MsgBox("Please input item data", MsgBoxStyle.Information, "Information")
                Me.dgvItem.AllowUserToAddRows = True
                Return False
            End If


            '3. cek perubahan nomor transaksi
            If Me.Baru = False Then
                If Me.Obj.TanggalTransaksi.Month = Me.DtpDate.Value.Month And _
                Me.Obj.TanggalTransaksi.Year = Me.DtpDate.Value.Year Then
                    Me.TxtDocumentNo.Text = Me.Obj.DocumentNo
                Else
                    Dim NoTransaksiLama As String = Me.Obj.GetDocumentNo(Me.TxtPrimary.Text)
                    If NoTransaksiLama <> Me.TxtDocumentNo.Text Then
                        If Obj.ApakahSudahAda(Me.TxtDocumentNo.Text) = False Then
                            If MsgBox("You just changed transaction date, this will change your transaction " _
                                      + "number." _
                                      + " Continue and accept this change?", _
                                      MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                                Me.DtpDate.Value = Me.Obj.TanggalTransaksi
                                Me.TxtDocumentNo.Text = Me.Obj.DocumentNo
                                Me.dgvItem.AllowUserToAddRows = True
                                Return False
                            End If
                        End If
                    End If
                End If
            End If

            '4. cek data yang salah, double dalam kode
            If Me.Baru = True Then
                If Me.Obj.ApakahSudahAda(Me.TxtDocumentNo.Text) = True Then
                    Me.TxtDocumentNo.Text = Obj.GetNoBaru(Me.DtpDate.Value, Me.ObjCD)
                End If
            Else
                Dim TipeCodeLama As String = Me.Obj.GetDocumentNo(Me.TxtPrimary.Text)
                If TipeCodeLama <> Me.TxtDocumentNo.Text Then
                    If Me.Obj.ApakahSudahAda(Me.TxtDocumentNo.Text) = True Then
                        Me.TxtDocumentNo.Text = Obj.GetNoBaru(Me.DtpDate.Value, Me.ObjCD)
                    End If
                End If
                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.Obj.GetTimestamp(Me.TxtPrimary.Text)
                If Me.Obj.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have modified this data, reload latest data?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.Obj.CariDetail(Me.Obj.Primary, ObjCD, Me.dgvItem, Me.DefPrec, Me)
                        Me.SalinKeField()
                    End If
                    Me.dgvItem.AllowUserToAddRows = True
                    Return False
                End If
            End If

            If String.IsNullOrEmpty(Me.TxtRemarks.Text) = False Then
                If Me.ObjCD.CekDataKosong(Me.TxtRemarks.Text, False) = False Then
                    MsgBox("Data cannot contains character : ' , % , or * .", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.AllowUserToAddRows = True
                    Return False
                End If
            End If

            CoventionDate = ObjClosing.GetPayrollConvertionDate()
            If CoventionDate.Year > Me.dtpPeriod.Value.Year Then
                MsgBox("Please make transaction period on " + CoventionDate.Month.ToString + CoventionDate.Year.ToString, MsgBoxStyle.Information, "Information")
                Me.dgvItem.AllowUserToAddRows = True
                Return False
            End If

            If Me.dtpPeriod.Value.Month = 12 Then
                MsgBox("Tax beginning balance in December is invalid.", MsgBoxStyle.Information, "Information")
                Me.dgvItem.AllowUserToAddRows = True
                Return False
            End If

            If Me.dtpPeriod.Value.Year <> CoventionDate.Year Then
                If MsgBox("Payroll convertion date is: " + CoventionDate.Year.ToString + ", continue anyway?", MsgBoxStyle.YesNo, "Confirmation") <> MsgBoxResult.Yes Then
                    Me.dgvItem.AllowUserToAddRows = True
                    Return False
                End If
                'MsgBox("Please make transaction period  " + CoventionDate.Year.ToString, MsgBoxStyle.Information, "Information")
            End If

            'If Me.dtpPeriod.Value.Month = 12 Then
            '    If Me.dtpPeriod.Value.Year <> CoventionDate.Year Then
            '        'MsgBox("Maximum period is november " + CoventionDate.Year.ToString, MsgBoxStyle.Information, "Information")
            '        MsgBox("You can not enter the data outside of " + CoventionDate.Year.ToString + _
            '        " and maximum period is november " + CoventionDate.Year.ToString, MsgBoxStyle.Information, "Information")
            '        Me.dgvItem.AllowUserToAddRows = True
            '        Return False
            '    End If
            'Else
            '    If Me.dtpPeriod.Value.Year <> CoventionDate.Year Then
            '        MsgBox("You can not enter the data outside of " + CoventionDate.Year.ToString, MsgBoxStyle.Information, "Information")
            '        Me.dgvItem.AllowUserToAddRows = True
            '        Return False
            '    End If
            'End If

            Dim StringKet As String = Me.Obj.Cek2TransBeginningBalance(Me.Baru, Me.dtpPeriod.Value, Me.ObjCD, Me)
            If String.IsNullOrEmpty(StringKet) = False Then
                MsgBox("The following employee(s) have tax beginning balance calculation: " + vbNewLine + StringKet, MsgBoxStyle.Information, "Information")
                Me.dgvItem.AllowUserToAddRows = True
                Return False
            End If

            'kalo sampe sini return true
            Me.dgvItem.AllowUserToAddRows = True
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
                .DocumentNo = Me.TxtDocumentNo.Text
                .TanggalTransaksi = Me.DtpDate.Value
                .Period = "#" + Me.dtpPeriod.Value.Date.Month.ToString + "/1/" + Me.dtpPeriod.Value.Year.ToString + "#"
                .Remarks = Me.TxtRemarks.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.Obj
                Me.TxtPrimary.Text = .Primary
                Me.TxtDocumentNo.Text = .DocumentNo
                Me.DtpDate.Value = .TanggalTransaksi
                Me.dtpPeriod.Value = .Period
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
                        If Me.Obj.SimpanBaru(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, Me.dgvItem, _
                        Me) = True Then
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
                        If Me.Obj.SimpanPerubahan(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, Me.dgvItem, _
                        Me) = True Then
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
            If Not Me.TxtDocumentNo.Text = "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Baru = True Then
                        Me.KosongkanField()
                        Me.TxtDocumentNo.Clear()
                    Else
                        Me.Obj.CariDetail(Me.TxtPrimary.Text, Me.ObjCD, _
                        Me.dgvItem, Me.DefPrec, Me)
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
            Obj.ResetObject(Me.dgvItem)
            Me.AktifkanDetail()
            Me.KosongkanField()
            Me.TxtDocumentNo.Text = Me.Obj.GetNoBaru(Me.DtpDate.Value, Me.ObjCD)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.DtpDate.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ijk As New FrmTrans_BeginingBalanceTax_Cari
        Try
            ijk.BackColorAll = Me.BackColorAll
            ijk.cboCriteria.Items.Clear()
            ijk.txtCari.Text = ""
            ijk.dgvHasil.DataSource = ""
            ijk.ShowDialog()
            If Not ijk.mPrimary = -1 Then
                If Me.Obj.CariDetail(ijk.mPrimary, Me.ObjCD, _
                Me.dgvItem, Me.DefPrec, Me) = True Then
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
            If Me.TxtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtDocumentNo.Text <> "" Then
                If Me.Obj.ApakahBolehEdit(Me.TxtPrimary.Text, Me.UserCode, True, True) = True Then
                    Me.Baru = False
                    Me.AktifkanDetail()
                    Me.DtpDate.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtDocumentNo.Text <> "" Then
                If Me.Obj.ApakahBolehHapus(Me.TxtPrimary.Text, Me.UserCode, True, True, Me) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.Hapus(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, Me.dgvItem, Me, True) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.AllowUserToAddRows = False
                        End If
                    Else
                        Me.dgvItem.AllowUserToAddRows = False
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
            If Me.TxtDocumentNo.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If Baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.Obj.SimpanBaru(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, _
                            Me.dgvItem, Me) = True Then
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
                            If Me.Obj.SimpanPerubahan(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, _
                            Me.dgvItem, Me) = True Then
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

    ''GetNoBaru
    Private Sub dtpDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpDate.ValueChanged
        Try
            Me.BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColKodeKaryawan.Index Or _
                e.ColumnIndex = Me.ColRemarks.Index Or e.ColumnIndex = Me.ColGaji.Index Or _
                e.ColumnIndex = Me.ColTunjanganPPH.Index Or e.ColumnIndex = Me.ColTunjanganLainnya.Index Or _
                e.ColumnIndex = Me.ColHonor.Index Or e.ColumnIndex = Me.ColPremi.Index Or _
                e.ColumnIndex = Me.ColNatura.Index Or e.ColumnIndex = Me.ColTantiem.Index Or _
                e.ColumnIndex = Me.ColBiayaJabatan.Index Or e.ColumnIndex = Me.ColIuranPensiun.Index Or _
                e.ColumnIndex = Me.ColStatus.Index Or e.ColumnIndex = Me.ColPenghasilBruto.Index Then
                If IsNothing(Me.dgvItem.Item(Me.ColPrimaryKaryawan.Index, e.RowIndex).Value) = True Then
                    Me.dgvItem.Item(Me.ColPrimaryKaryawan.Index, e.RowIndex).Value = "-1"
                End If
                If Me.dgvItem.Item(Me.ColPrimaryKaryawan.Index, e.RowIndex).Value.ToString = "" _
                Or Me.dgvItem.Item(Me.ColPrimaryKaryawan.Index, e.RowIndex).Value.ToString = "-1" Then
                    If e.ColumnIndex = Me.ColKodeKaryawan.Index Then
                        Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        Me.BufferStatus = Me.dgvItem.Item(Me.ColStatus.Index, e.RowIndex).Value
                        Me.BufferNPWPNo = Me.dgvItem.Item(Me.ColNoNPWP.Index, e.RowIndex).Value
                        Me.BufferCategoryTax = Me.dgvItem.Item(Me.ColStatus.Index, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                Else
                    Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    Me.BufferStatus = Me.dgvItem.Item(Me.ColStatus.Index, e.RowIndex).Value
                    Me.BufferNPWPNo = Me.dgvItem.Item(Me.ColNoNPWP.Index, e.RowIndex).Value
                    Me.BufferCategoryTax = Me.dgvItem.Item(Me.ColStatus.Index, e.RowIndex).Value
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemInput(ByVal vPrimaryEmployee As Double) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColPrimaryKaryawan.Index).Value = vPrimaryEmployee Then
                    Return rowDGV.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub DGVItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim Row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
        Dim Index As Integer = -1
        '-------Batasan setelah edit------------
        Dim DTA As New DseTrans_Tax_Beginning_BalanceTableAdapters.tbltrans_tax_beginning_balanceTableAdapter
        Dim TBLNew As New DseTrans_Tax_Beginning_Balance.tbltrans_tax_beginning_balanceDataTable
        Dim TBLEdit As New DseTrans_Tax_Beginning_Balance.tbltrans_tax_beginning_balanceDataTable
        Dim RowTbl As DseTrans_Tax_Beginning_Balance.tbltrans_tax_beginning_balanceRow
        'buffer
        Dim Buffer_Primary As String = "-1"
        Dim CboDgv As DataGridViewComboBoxCell
        Try
            'Me.dgvItem.AllowUserToAddRows = False
            'pastikan ada isinya
            With Row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColGaji.Index Or e.ColumnIndex = Me.ColTunjanganPPH.Index Or _
                    e.ColumnIndex = Me.ColTunjanganLainnya.Index Or e.ColumnIndex = Me.ColHonor.Index Or _
                    e.ColumnIndex = Me.ColPremi.Index Or e.ColumnIndex = Me.ColTantiem.Index Or _
                    e.ColumnIndex = Me.ColBiayaJabatan.Index Or e.ColumnIndex = Me.ColIuranPensiun.Index Or _
                    e.ColumnIndex = Me.ColNatura.Index Then
                        If IsNumeric(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Please input numeric data", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec)
                            'Me.dgvItem.AllowUserToAddRows = True
                            Exit Sub
                        End If
                        If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Please input positive number", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec)
                            'Me.dgvItem.AllowUserToAddRows = True
                            Exit Sub
                        End If
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec)
                    End If
                    If e.ColumnIndex = Me.ColKodeKaryawan.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjKary.ApakahSudahAda(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            Index = Me.GetIndexItemInput(Me.ObjKary.GetPrimaryByCode(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value))
                            If Index < 0 Then
                                Me.dgvItem.Item(Me.ColPrimaryKaryawan.Index, e.RowIndex).Value = Me.ObjKary.GetPrimaryByCode(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.dgvItem.Item(Me.ColNamaKaryawan.Index, e.RowIndex).Value = Me.ObjKary.GetFullNameByPrimary(Me.dgvItem.Item(Me.ColPrimaryKaryawan.Index, e.RowIndex).Value)
                                Me.dgvItem.Item(Me.ColNIK.Index, e.RowIndex).Value = Me.ObjKary.GetNIK(Me.dgvItem.Item(Me.ColPrimaryKaryawan.Index, e.RowIndex).Value)
                                CboDgv = Row.Cells(Me.ColStatus.Index)
                                Row.Cells(Me.ColStatus.Index).Value = CboDgv.Items.Item(Me.ObjKary.GetNPWPCategory(Row.Cells(Me.ColPrimaryKaryawan.Index).Value))
                                Row.Cells(Me.ColNoNPWP.Index).Value = Me.ObjKary.GetNoNPWP(Row.Cells(Me.ColPrimaryKaryawan.Index).Value)
                                Row.Cells(Me.ColDepartement.Index).Value = Me.ObjKary.GetDeptDescription(Row.Cells(Me.ColPrimaryKaryawan.Index).Value)
                                Row.Cells(Me.ColPosition.Index).Value = Me.ObjKary.GetPositionDescription(Row.Cells(Me.ColPrimaryKaryawan.Index).Value)
                            Else
                                If e.RowIndex <> Index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvItem.Item(Me.ColKodeKaryawan.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.dgvItem.AllowUserToAddRows = True
                                    Me.ObjCD.SetNumber(Me.dgvItem)
                                    Exit Sub
                                End If
                            End If
                        Else
                            '2. kode akun yang diketik tidak ada --> tampilkan form cari
                            Dim ij As New Core_Module.FrmMasterKaryawanCari
                            ij.txtCari.Text = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryEmploye <> -1 Then
                                '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                                Index = Me.GetIndexItemInput(ij.PrimaryEmploye)
                                If Index < 0 Then
                                    Me.dgvItem.Item(Me.ColKodeKaryawan.Index, e.RowIndex).Value = Me.ObjKary.GetCodeByPrimary(ij.PrimaryEmploye)
                                    Me.dgvItem.Item(Me.ColNamaKaryawan.Index, e.RowIndex).Value = Me.ObjKary.GetFullNameByPrimary(ij.PrimaryEmploye)
                                    Me.dgvItem.Item(Me.ColNIK.Index, e.RowIndex).Value = Me.ObjKary.GetNIK(ij.PrimaryEmploye)
                                    Me.dgvItem.Item(Me.ColPrimaryKaryawan.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                    CboDgv = Row.Cells(Me.ColStatus.Index)
                                    Row.Cells(Me.ColStatus.Index).Value = CboDgv.Items.Item(Me.ObjKary.GetNPWPCategory(Row.Cells(Me.ColPrimaryKaryawan.Index).Value))
                                    Row.Cells(Me.ColNoNPWP.Index).Value = Me.ObjKary.GetNoNPWP(Row.Cells(Me.ColPrimaryKaryawan.Index).Value)
                                    Row.Cells(Me.ColDepartement.Index).Value = Me.ObjKary.GetDeptDescription(Row.Cells(Me.ColPrimaryKaryawan.Index).Value)
                                    Row.Cells(Me.ColPosition.Index).Value = Me.ObjKary.GetPositionDescription(Row.Cells(Me.ColPrimaryKaryawan.Index).Value)
                                Else
                                    If e.RowIndex <> Index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.dgvItem.AllowUserToAddRows = True
                                        Me.ObjCD.SetNumber(Me.dgvItem)
                                        Exit Sub
                                    Else
                                        Me.dgvItem.Item(Me.ColKodeKaryawan.Index, e.RowIndex).Value = Me.ObjKary.GetCodeByPrimary(ij.PrimaryEmploye)
                                        Me.dgvItem.Item(Me.ColNamaKaryawan.Index, e.RowIndex).Value = Me.ObjKary.GetFullNameByPrimary(ij.PrimaryEmploye)
                                        Me.dgvItem.Item(Me.ColNIK.Index, e.RowIndex).Value = Me.ObjKary.GetNIK(ij.PrimaryEmploye)
                                        Me.dgvItem.Item(Me.ColPrimaryKaryawan.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                        CboDgv = Row.Cells(Me.ColStatus.Index)
                                        Row.Cells(Me.ColStatus.Index).Value = CboDgv.Items.Item(Me.ObjKary.GetNPWPCategory(Row.Cells(Me.ColPrimaryKaryawan.Index).Value))
                                        Row.Cells(Me.ColNoNPWP.Index).Value = Me.ObjKary.GetNoNPWP(Row.Cells(Me.ColPrimaryKaryawan.Index).Value)
                                        Row.Cells(Me.ColDepartement.Index).Value = Me.ObjKary.GetDeptDescription(Row.Cells(Me.ColPrimaryKaryawan.Index).Value)
                                        Row.Cells(Me.ColPosition.Index).Value = Me.ObjKary.GetPositionDescription(Row.Cells(Me.ColPrimaryKaryawan.Index).Value)
                                    End If
                                End If
                            ElseIf ij.PrimaryEmploye = -1 Then
                                If Me.dgvItem.Rows.Count - 1 <> e.RowIndex Then
                                    Me.EraseRowIndex = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.dgvItem)
                        '------------Batasan Edit---------------
                        DTA.Connection.ConnectionString = Core_Module.GetConstring
                        If Me.Baru = True Then
                            DTA.FillBy_GetALLDataByPeriod(TBLNew, Row.Cells(Me.ColPrimaryKaryawan.Index).Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value))
                            If TBLNew.Rows.Count > 0 Then
                                RowTbl = TBLNew.Rows(0)
                                MsgBox("The employee " + Row.Cells(Me.ColKodeKaryawan.Index).Value.ToString + "-" + Row.Cells(Me.ColNamaKaryawan.Index).Value.ToString + " has tax beginning calculation for " + _
                                Format(RowTbl.Period.Date, "MMM") + " " + Format(RowTbl.Period, "yyyy") + " on " + RowTbl.Document_No.ToString, MsgBoxStyle.Information, "Information")
                                If String.IsNullOrEmpty(Me.Buffer) = True Then
                                    Me.EraseRowIndex = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                Else
                                    Me.dgvItem.Item(Me.ColKodeKaryawan.Index, e.RowIndex).Value = Me.Buffer
                                    Me.dgvItem.Item(Me.ColPrimaryKaryawan.Index, e.RowIndex).Value = Me.ObjKary.GetPrimaryByCode(Me.dgvItem.Item(Me.ColKodeKaryawan.Index, e.RowIndex).Value)
                                    'CboDgv = Row.Cells(Me.ColStatus.Index)
                                    'Row.Cells(Me.ColStatus.Index).Value = CboDgv.Items.Item(Me.ObjKary.GetNPWPCategory(Row.Cells(Me.ColPrimaryKaryawan.Index).Value))
                                    Row.Cells(Me.ColStatus.Index).Value = Me.BufferStatus
                                    Row.Cells(Me.ColNoNPWP.Index).Value = Me.ObjKary.GetNoNPWP(Row.Cells(Me.ColPrimaryKaryawan.Index).Value)
                                    Row.Cells(Me.ColDepartement.Index).Value = Me.ObjKary.GetDeptDescription(Row.Cells(Me.ColPrimaryKaryawan.Index).Value)
                                    Row.Cells(Me.ColPosition.Index).Value = Me.ObjKary.GetPositionDescription(Row.Cells(Me.ColPrimaryKaryawan.Index).Value)
                                    Row.Cells(Me.ColNamaKaryawan.Index).Value = Me.ObjKary.GetFullNameByPrimary(Row.Cells(Me.ColPrimaryKaryawan.Index).Value)
                                End If
                                Exit Sub
                            End If
                        Else
                            DTA.FillBy_GetALLDataByPeriodEdit(TBLEdit, Row.Cells(Me.ColPrimaryKaryawan.Index).Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value), Me.TxtPrimary.Text)
                            If TBLEdit.Rows.Count > 0 Then
                                RowTbl = TBLEdit.Rows(0)
                                MsgBox("The employee " + Row.Cells(Me.ColKodeKaryawan.Index).Value.ToString + "-" + Row.Cells(Me.ColNamaKaryawan.Index).Value.ToString + " has tax beginning calculation for " + _
                                Format(RowTbl.Period, "MMM") + " " + Format(RowTbl.Period, "yyyy") + " on " + RowTbl.Document_No.ToString, MsgBoxStyle.Information, "Information")
                                If String.IsNullOrEmpty(Me.Buffer) = True Then
                                    Me.EraseRowIndex = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                Else
                                    Me.dgvItem.Item(Me.ColKodeKaryawan.Index, e.RowIndex).Value = Me.Buffer
                                    Me.dgvItem.Item(Me.ColPrimaryKaryawan.Index, e.RowIndex).Value = Me.ObjKary.GetPrimaryByCode(Me.dgvItem.Item(Me.ColKodeKaryawan.Index, e.RowIndex).Value)
                                    'CboDgv = Row.Cells(Me.ColStatus.Index)
                                    'Row.Cells(Me.ColStatus.Index).Value = CboDgv.Items.Item(Me.ObjKary.GetNPWPCategory(Row.Cells(Me.ColPrimaryKaryawan.Index).Value))
                                    Row.Cells(Me.ColStatus.Index).Value = Me.BufferStatus
                                    Row.Cells(Me.ColNoNPWP.Index).Value = Me.ObjKary.GetNoNPWP(Row.Cells(Me.ColPrimaryKaryawan.Index).Value)
                                    Row.Cells(Me.ColDepartement.Index).Value = Me.ObjKary.GetDeptDescription(Row.Cells(Me.ColPrimaryKaryawan.Index).Value)
                                    Row.Cells(Me.ColPosition.Index).Value = Me.ObjKary.GetPositionDescription(Row.Cells(Me.ColPrimaryKaryawan.Index).Value)
                                    Row.Cells(Me.ColNamaKaryawan.Index).Value = Me.ObjKary.GetFullNameByPrimary(Row.Cells(Me.ColPrimaryKaryawan.Index).Value)
                                End If
                                Exit Sub
                            End If
                        End If
                        If Me.ApakahBolehEdit() = False Then
                            If IsNothing(Me.Buffer) = True Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                                Exit Sub
                            End If
                            Me.dgvItem.Item(Me.ColKodeKaryawan.Index, e.RowIndex).Value = Me.Buffer
                            Me.dgvItem.Item(Me.ColPrimaryKaryawan.Index, e.RowIndex).Value = Me.ObjKary.GetPrimaryByCode(Me.dgvItem.Item(Me.ColKodeKaryawan.Index, e.RowIndex).Value)
                            CboDgv = Row.Cells(Me.ColStatus.Index)
                            'Row.Cells(Me.ColStatus.Index).Value = CboDgv.Items.Item(Me.ObjKary.GetNPWPCategory(Row.Cells(Me.ColPrimaryKaryawan.Index).Value))
                            'Row.Cells(Me.ColNoNPWP.Index).Value = Me.ObjKary.GetNoNPWP(Row.Cells(Me.ColPrimaryKaryawan.Index).Value)
                            Row.Cells(Me.ColStatus.Index).Value = Me.BufferStatus
                            Row.Cells(Me.ColNoNPWP.Index).Value = Me.BufferNPWPNo
                            Row.Cells(Me.ColDepartement.Index).Value = Me.ObjKary.GetDeptDescription(Row.Cells(Me.ColPrimaryKaryawan.Index).Value)
                            Row.Cells(Me.ColPosition.Index).Value = Me.ObjKary.GetPositionDescription(Row.Cells(Me.ColPrimaryKaryawan.Index).Value)
                            Row.Cells(Me.ColNamaKaryawan.Index).Value = Me.ObjKary.GetFullNameByPrimary(Row.Cells(Me.ColPrimaryKaryawan.Index).Value)
                            Exit Sub
                        End If
                        '-----
                    End If
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        If Me.ObjCD.CekDataKosong(Row.Cells(Me.ColRemarks.Index).Value.ToString, False) = False Then
                            MsgBox("Data cannot contains character : ' , % , atau * .", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            'Me.dgvItem.AllowUserToAddRows = True
                            Exit Sub
                        End If
                    End If
                    'Hitungan PPh 21 
                    'If e.ColumnIndex = Me.ColGaji.Index Or e.ColumnIndex = Me.ColTunjanganPPH.Index Or _
                    'e.ColumnIndex = Me.ColTunjanganLainnya.Index Or e.ColumnIndex = Me.ColHonor.Index Or _
                    'e.ColumnIndex = Me.ColPremi.Index Or e.ColumnIndex = Me.ColTantiem.Index Or _
                    'e.ColumnIndex = Me.ColBiayaJabatan.Index Or e.ColumnIndex = Me.ColNatura.Index Then
                    '    Row.Cells(Me.ColPenghasilBruto.Index).Value = Me.ObjCD.DigitGrouping(Me.PerhitunganJumlahPenghasilanBruto(Row.Cells(Me.ColGaji.Index).Value, _
                    '    Row.Cells(Me.ColTunjanganPPH.Index).Value, Row.Cells(Me.ColTunjanganLainnya.Index).Value, Row.Cells(Me.ColHonor.Index).Value, _
                    '    Row.Cells(Me.ColPremi.Index).Value, Row.Cells(Me.ColNatura.Index).Value, Row.Cells(Me.ColTantiem.Index).Value), Me.DefPrec)
                    '    '9. Biaya Jabatan
                    '    If CDec(Row.Cells(Me.ColBiayaJabatan.Index).Value) = 0 Then
                    '        Dim CostJabatan As Decimal = 0
                    '        CostJabatan = 0.05 * CDec(Row.Cells(Me.ColPenghasilBruto.Index).Value)
                    '        If CostJabatan < 500000 Then
                    '            Me.dgvItem.Item(Me.ColBiayaJabatan.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(CostJabatan, Me.DefPrec)
                    '        Else
                    '            Row.Cells(Me.ColBiayaJabatan.Index).Value = Me.ObjCD.DigitGrouping(500000, Me.DefPrec)
                    '        End If
                    '    End If
                    '    Me.HitungDariNomor11Sampai20(Row)
                    'End If
                    'If e.ColumnIndex = Me.ColBiayaJabatan.Index Then
                    '    Dim LimaPersen As Decimal = 0
                    '    If CDec(Row.Cells(Me.ColBiayaJabatan.Index).Value) < 500000 Then
                    '        LimaPersen = 0.05 * CDec(Row.Cells(Me.ColPenghasilBruto.Index).Value)
                    '        If CDec(Me.ObjCD.DigitGrouping(LimaPersen, Me.DefPrec)) <> CDec(Row.Cells(Me.ColBiayaJabatan.Index).Value) Then
                    '            Dim HasilMsg_1 As MsgBoxResult
                    '            HasilMsg_1 = MsgBox("Biaya jabatan should be " + Me.ObjCD.DigitGrouping(LimaPersen, Me.DefPrec) + " , do you want to continue?", MsgBoxStyle.YesNo, "Confirm")
                    '            If HasilMsg_1 = MsgBoxResult.No Then
                    '                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec)
                    '                Exit Sub
                    '            Else
                    '                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec)
                    '                Me.HitungDariNomor11Sampai20(Row)
                    '            End If
                    '        End If
                    '    Else
                    '        Row.Cells(Me.ColBiayaJabatan.Index).Value = Me.ObjCD.DigitGrouping(500000, Me.DefPrec)
                    '        Me.HitungDariNomor11Sampai20(Row)
                    '    End If
                    'End If
                    'If e.ColumnIndex = Me.ColIuranPensiun.Index Or e.ColumnIndex = Me.ColStatus.Index Then
                    '    Me.HitungDariNomor11Sampai20(Row)
                    'End If
                Else
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                    If Me.dgvItem.Item(Me.ColPrimaryKaryawan.Index, e.RowIndex).Value = "-1" Then
                        Me.EraseRowIndex = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                    End If
                End If
            End With
            'Me.dgvItem.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBLNew.Dispose()
            TBLEdit.Dispose()
        End Try
    End Sub

    Private Sub HitungDariNomor11Sampai20(ByVal Row As DataGridViewRow)
        Dim CBODGV_PTKP As DataGridViewComboBoxCell
        Dim vColumn As DataGridViewColumn
        Try
            '11. Jumlah Pengurang
            Row.Cells(Me.ColJumlahPengurangan.Index).Value = Me.ObjCD.DigitGrouping(CDec(Row.Cells(Me.ColBiayaJabatan.Index).Value) + CDec(Row.Cells(Me.ColIuranPensiun.Index).Value), Me.DefPrec)
            '12. Jumlah Penghasilan Netto
            Row.Cells(Me.ColJumlahPenghasilanNetto.Index).Value = Me.ObjCD.DigitGrouping(CDec(Row.Cells(Me.ColPenghasilBruto.Index).Value) - CDec(Row.Cells(Me.ColJumlahPengurangan.Index).Value), Me.DefPrec)
            '13. Penghasilan masa pajak sebelum
            Row.Cells(Me.ColJumlahPenghasilNettoMasaPajakSebelum.Index).Value = "0"
            '14. Penghasilan Netto untuk PPh
            Row.Cells(Me.ColJumlahPenghasilanNettoUntukPPH.Index).Value = Me.ObjCD.DigitGrouping((CDec(Row.Cells(Me.ColJumlahPenghasilanNetto.Index).Value) * 12), Me.DefPrec)
            '15. PTKP
            Dim TipeTax As Integer = 0
            CBODGV_PTKP = Row.Cells(Me.ColStatus.Index)
            If CBODGV_PTKP.Items.IndexOf(Row.Cells(Me.ColStatus.Index).Value) = FrmTrans_BeginingBalanceTax.EnumCriteria.TK0 Then
                Row.Cells(Me.ColPTKP.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjTax.GetPTKPPerTahun(FrmTrans_BeginingBalanceTax.EnumCriteria.TK0), Me.DefPrec)
            ElseIf CBODGV_PTKP.Items.IndexOf(Row.Cells(Me.ColStatus.Index).Value) = FrmTrans_BeginingBalanceTax.EnumCriteria.TK1 Then
                Row.Cells(Me.ColPTKP.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjTax.GetPTKPPerTahun(FrmTrans_BeginingBalanceTax.EnumCriteria.TK1), Me.DefPrec)
            ElseIf CBODGV_PTKP.Items.IndexOf(Row.Cells(Me.ColStatus.Index).Value) = FrmTrans_BeginingBalanceTax.EnumCriteria.TK2 Then
                Row.Cells(Me.ColPTKP.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjTax.GetPTKPPerTahun(FrmTrans_BeginingBalanceTax.EnumCriteria.TK2), Me.DefPrec)
            ElseIf CBODGV_PTKP.Items.IndexOf(Row.Cells(Me.ColStatus.Index).Value) = FrmTrans_BeginingBalanceTax.EnumCriteria.TK3 Then
                Row.Cells(Me.ColPTKP.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjTax.GetPTKPPerTahun(FrmTrans_BeginingBalanceTax.EnumCriteria.TK3), Me.DefPrec)
            ElseIf CBODGV_PTKP.Items.IndexOf(Row.Cells(Me.ColStatus.Index).Value) = FrmTrans_BeginingBalanceTax.EnumCriteria.K0 Then
                Row.Cells(Me.ColPTKP.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjTax.GetPTKPPerTahun(FrmTrans_BeginingBalanceTax.EnumCriteria.K0), Me.DefPrec)
            ElseIf CBODGV_PTKP.Items.IndexOf(Row.Cells(Me.ColStatus.Index).Value) = FrmTrans_BeginingBalanceTax.EnumCriteria.K1 Then
                Row.Cells(Me.ColPTKP.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjTax.GetPTKPPerTahun(FrmTrans_BeginingBalanceTax.EnumCriteria.K1), Me.DefPrec)
            ElseIf CBODGV_PTKP.Items.IndexOf(Row.Cells(Me.ColStatus.Index).Value) = FrmTrans_BeginingBalanceTax.EnumCriteria.K2 Then
                Row.Cells(Me.ColPTKP.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjTax.GetPTKPPerTahun(FrmTrans_BeginingBalanceTax.EnumCriteria.K2), Me.DefPrec)
            ElseIf CBODGV_PTKP.Items.IndexOf(Row.Cells(Me.ColStatus.Index).Value) = FrmTrans_BeginingBalanceTax.EnumCriteria.K3 Then
                Row.Cells(Me.ColPTKP.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjTax.GetPTKPPerTahun(FrmTrans_BeginingBalanceTax.EnumCriteria.K3), Me.DefPrec)
            End If
            '16. Penghasilan Kena Pajak
            Row.Cells(Me.ColPenghasilanKenaPajak.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(CDec(Row.Cells(Me.ColJumlahPenghasilanNettoUntukPPH.Index).Value) - CDec(Row.Cells(Me.ColPTKP.Index).Value), _
            Me.DefPrec, ClsCekData.EnumTipeRounding.Down, 1000, Me.ObjCD), 0, ClsCekData.EnumTipeRounding.Down)
            '17. PPh ps 21 atas PKP
            Row.Cells(Me.ColPPhPs21AtasPKP.Index).Value = Me.ObjCD.DigitGrouping(CDec(Me.Obj.HitungPKP(Row.Cells(Me.ColPenghasilanKenaPajak.Index).Value)), Me.DefPrec)
            '18. PPh Ps 21 yang telah dipotong masa pajak sebelum
            Row.Cells(Me.ColPPHps21TelahDipotong.Index).Value = "0"
            '19. PPh Ps 21 terhitung 
            Row.Cells(Me.CollPPHTerutang.Index).Value = Me.ObjCD.DigitGrouping(CDec(Row.Cells(Me.ColPPhPs21AtasPKP.Index).Value) / 12, Me.DefPrec)
            '20. PPh Ps 21 dan pasal 26
            Row.Cells(Me.ColPPhTelahDipotong.Index).Value = "0"
            '21. Cek jika angka negative, maka format = 0
            If Me.dgvItem.Columns.Count > 0 Then
                Me.dgvItem.AllowUserToAddRows = False
                For Each Row In Me.dgvItem.Rows
                    For ij As Integer = 0 To Me.dgvItem.Columns.Count - 1
                        vColumn = Me.dgvItem.Columns(ij)
                        If ij <> Me.ColEmp_No.Index Or ij <> Me.ColKodeKaryawan.Index Or ij <> Me.ColNamaKaryawan.Index _
                        Or ij <> Me.ColStatus.Index Or ij <> Me.ColNoNPWP.Index Or ij <> Me.ColNIK.Index _
                        Or ij <> Me.ColDepartement.Index Or ij <> Me.ColPosition.Index _
                        Or ij <> Me.ColRemarks.Index Or ij <> Me.ColPrimaryKaryawan.Index Then
                            If IsNumeric(Me.dgvItem.Item(ij, Row.Index).Value) = True Then
                                If CDec(Me.dgvItem.Item(ij, Row.Index).Value) <= 0 Then
                                    Me.dgvItem.Item(ij, Row.Index).Value = "0"
                                End If
                            End If
                        End If
                    Next
                Next
                Me.dgvItem.AllowUserToAddRows = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function PerhitunganJumlahPenghasilanBruto(ByVal nmr1 As Decimal, ByVal nmr2 As Decimal, ByVal nmr3 As Decimal, _
    ByVal nmr4 As Decimal, ByVal nmr5 As Decimal, ByVal nmr6 As Decimal, ByVal nmr7 As Decimal) As Decimal
        Dim Retval As Decimal = 0
        Try
            Retval = nmr1 + nmr2 + nmr3 + nmr4 + nmr5 + nmr6 + nmr7
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlSalaryByTanggalSalary(ByVal vEndDate As Date, ByVal vStartDate As Date) As Double
        Dim DTA As New DseTrans_Tax_Beginning_BalanceTableAdapters.tbltrans_salaryTableAdapter
        Dim Dummy As New System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJmlByTanggalDari_Sampai(vEndDate, vStartDate)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function ApakahBolehEdit() As Boolean
        Dim DTA As New DseTrans_Tax_Beginning_BalanceTableAdapters.tbltrans_salaryTableAdapter
        Dim TBL As New DseTrans_Tax_Beginning_Balance.tbltrans_salaryDataTable
        Dim RowTbl As DseTrans_Tax_Beginning_Balance.tbltrans_salaryRow
        Dim InfoMsg As String = ""
        Dim ij As Integer = 1
        Try
            If Me.GetJmlSalaryByTanggalSalary(Me.ObjCD.GetLastDateOfTheYear(Me.dtpPeriod.Value), _
                Me.ObjCD.GetFirstDateofTheMonth(Me.ObjCD.GetLastDateOfTheYear(Me.dtpPeriod.Value))) > 0 Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                DTA.FillBy_TanggalDariSampai(TBL, Me.ObjCD.GetLastDateOfTheYear(Me.dtpPeriod.Value), _
                Me.ObjCD.GetFirstDateofTheMonth(Me.ObjCD.GetLastDateOfTheYear(Me.dtpPeriod.Value)))
                If TBL.Rows.Count > 0 Then
                    For Each RowTbl In TBL.Rows
                        If InfoMsg = "" Then
                            InfoMsg = ij.ToString + ". " + RowTbl.No_Transaksi.ToString + "."
                        Else
                            InfoMsg += vbCrLf + ij.ToString + ". " + RowTbl.No_Transaksi.ToString + "."
                        End If
                        ij += 1
                    Next
                End If
                If MsgBox("The following employee have salary transaction in December " + Format(Me.dtpPeriod.Value, "yyyy") + " :" + _
                vbCrLf + InfoMsg + vbCrLf + "Continue anyway?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Private Sub dgvItem_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItem.DataError
        Try
            'do nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded

        Try
            For Each col1 As DataGridViewColumn In Me.dgvItem.Columns
                If col1.Index = Me.ColStatus.Index Then
                   Me.dgvItem.Item(col1.Index, Me.dgvItem.Rows.Count - 1).Value = ""
                Else
                    Me.dgvItem.Item(col1.Index, Me.dgvItem.Rows.Count - 1).Value = "0"
                End If
            Next
            Me.dgvItem.Item(Me.ColPrimaryKaryawan.Index, Me.dgvItem.Rows.Count - 1).Value = "-1"
            Me.dgvItem.Item(Me.ColNoNPWP.Index, Me.dgvItem.Rows.Count - 1).Value = ""
            Me.dgvItem.Item(Me.ColNIK.Index, Me.dgvItem.Rows.Count - 1).Value = ""
            Me.dgvItem.Item(Me.ColNamaKaryawan.Index, Me.dgvItem.Rows.Count - 1).Value = ""
            Me.dgvItem.Item(Me.ColEmp_No.Index, Me.dgvItem.Rows.Count - 1).Value = Me.dgvItem.Rows.Count.ToString
            Me.dgvItem.Item(Me.ColDepartement.Index, Me.dgvItem.Rows.Count - 1).Value = ""
            Me.dgvItem.Item(Me.ColKodeKaryawan.Index, Me.dgvItem.Rows.Count - 1).Value = ""
            Me.dgvItem.Item(Me.ColNIK.Index, Me.dgvItem.Rows.Count - 1).Value = ""
            Me.dgvItem.Item(Me.ColPosition.Index, Me.dgvItem.Rows.Count - 1).Value = ""
            Me.dgvItem.Item(Me.ColStatus.Index, Me.dgvItem.Rows.Count - 1).Value = ""
            Me.dgvItem.Item(Me.ColRemarks.Index, Me.dgvItem.Rows.Count - 1).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Dim Row As DataGridViewRow = Me.dgvItem.Rows(e.Row.Index)
        Dim DTA As New DseTrans_Tax_Beginning_BalanceTableAdapters.tbltrans_salaryTableAdapter
        Dim TBL As New DseTrans_Tax_Beginning_Balance.tbltrans_salaryDataTable
        Dim RowTbl As DseTrans_Tax_Beginning_Balance.tbltrans_salaryRow
        Dim InfoMsg As String = ""
        Dim ij As Integer = 1
        Try
            If Me.GetJmlSalaryByTanggalSalary(Me.ObjCD.GetLastDateOfTheYear(Me.dtpPeriod.Value), _
                Me.ObjCD.GetFirstDateofTheMonth(Me.ObjCD.GetLastDateOfTheYear(Me.dtpPeriod.Value))) > 0 Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                DTA.FillBy_TanggalDariSampai(TBL, Me.ObjCD.GetLastDateOfTheYear(Me.dtpPeriod.Value), _
                Me.ObjCD.GetFirstDateofTheMonth(Me.ObjCD.GetLastDateOfTheYear(Me.dtpPeriod.Value)))
                If TBL.Rows.Count > 0 Then
                    For Each RowTbl In TBL.Rows
                        If InfoMsg = "" Then
                            InfoMsg = ij.ToString + ". " + RowTbl.No_Transaksi.ToString + "."
                        Else
                            InfoMsg += vbCrLf + ij.ToString + ". " + RowTbl.No_Transaksi.ToString + "."
                        End If
                        ij += 1
                    Next
                End If
                If MsgBox("The following employee have salary transaction in December " + Format(Me.dtpPeriod.Value, "yyyy") + " :" + _
                vbCrLf + InfoMsg + vbCrLf + "Continue anyway?", MsgBoxStyle.YesNo, "Information") <> MsgBoxResult.Yes Then
                    e.Cancel = True
                    Exit Sub
                End If
               
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    'Back Color
    Private Sub TxtRemarks_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.ReadOnly = False Then
                    Me.TxtRemarks.BackColor = Me.BackColorAll
                Else
                    Me.TxtRemarks.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarks_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.ReadOnly = False Then
                    Me.TxtRemarks.BackColor = Me.TxtPrimary.BackColor
                Else
                    Me.TxtRemarks.BackColor = Me.TxtReadOnly.BackColor
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
        Try
            Me.dgvItem.AllowUserToAddRows = False
            If Me.dgvItem.Rows.Count > Me.EraseRowIndex Then
                Me.dgvItem.Rows.RemoveAt(Me.EraseRowIndex)
                Me.ObjCD.SetNumber(Me.dgvItem)
            End If
            Me.dgvItem.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.Baru = True Then
                    Me.TxtDocumentNo.Text = Me.Obj.GetNoBaru(Me.DtpDate.Value, Me.ObjCD)
                Else
                    If Me.DtpDate.Value.Month = Obj.GetTanggalTransaksi(Me.TxtPrimary.Text).Month And _
                    Me.DtpDate.Value.Year = Obj.GetTanggalTransaksi(Me.TxtPrimary.Text).Year Then
                        Me.TxtDocumentNo.Text = Me.Obj.GetDocumentNo(Me.TxtPrimary.Text)
                    Else
                        If Me.DtpDate.Value.Month <> Me.Obj.TanggalTransaksi.Month Or _
                        Me.DtpDate.Value.Year <> Me.Obj.TanggalTransaksi.Year Then
                            Me.TxtDocumentNo.Text = Me.Obj.GetNoBaru(Me.DtpDate.Value, Me.ObjCD)
                        Else
                            Me.TxtDocumentNo.Text = Me.Obj.GetDocumentNo(Me.TxtPrimary.Text)
                        End If
                    End If
                End If
                Me.ObjCD.SetNumber(Me.dgvItem)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SettingVisibleColumn()
        Dim RowSetting As DseTrans_Tax_Beginning_Balance.TblSettingDetailRow
        Try
            RowSetting = Me.TBLSetting.Rows(0)
            If RowSetting.Dept = 0 Then
                Me.dgvItem.Columns.Item(Me.ColDepartement.Index).Visible = False
            Else
                Me.dgvItem.Columns.Item(Me.ColDepartement.Index).Visible = True
            End If
            If RowSetting.NIK = 0 Then
                Me.dgvItem.Columns.Item(Me.ColNIK.Index).Visible = False
            Else
                Me.dgvItem.Columns.Item(Me.ColNIK.Index).Visible = True
            End If
            If RowSetting.Position = 0 Then
                Me.dgvItem.Columns.Item(Me.ColPosition.Index).Visible = False
            Else
                Me.dgvItem.Columns.Item(Me.ColPosition.Index).Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LblDetailSetting_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblDetailSetting.LinkClicked
        Dim Ijk As New FrmTrans_BeginingBalanceTax_Detail

        Try
            Ijk.TBL = Me.TBLSetting
            Ijk.ShowDialog()
            Me.SettingVisibleColumn()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Ijk.Dispose()
        End Try
    End Sub
End Class