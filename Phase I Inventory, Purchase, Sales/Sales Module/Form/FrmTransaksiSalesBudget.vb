Option Explicit On
Option Strict Off

Imports Core_Module
Public Class FrmTransaksiSalesBudget
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public StrictEdit As Boolean = True
    Public StrictDelete As Boolean = True

    Dim baru As Boolean
    Dim objSB As New ClsTransaksiSalesBudget
    Dim ObjCD As New ClsCekData
    Dim ObjProduct As New clsMasterProduct
    Dim ObjCurrency As New Core_Module.clsMasterCurrency
    Dim ObjUser As New Core_Module.ClsUser
    Dim objSBItem As New ClsTransaksiSalesBudget_Item
    Dim ObjLicense As New ClsLicense
    Dim ObjSalesP As New Core_Module.ClsMasterSalesPerson
    Dim ObjSalesPGroup As New Core_Module.ClsMasterSalesPersonGroup
    Dim DefaultPrecision As Integer = 0
    Dim buffer As String = ""
    Dim RemoveIndex As Integer = -1

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.txtPrimary.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.txtPrimary.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.txtPrimary.BackColor, False)
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)

            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            If Me.dgvItem.Rows.Count = 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
            Me.dgvItem.AllowUserToDeleteRows = True
            Me.dgvItem.ReadOnly = False
            Me.dgvItem.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True
            Me.dgvItem.AllowUserToAddRows = False

            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtReadOnly.BackColor, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, True)
            Me.txtPrimary.Text = "-1"
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
        Dim vPrimaryAcuan As Double = -1
        Dim rowDGV As DataGridViewRow
        Dim ListSalesPName As String = ""
        Dim DummyInt As Integer = 0

        Try
            'msgbox("belum")
            'data salah bila:
            '1. tidak ada 1 pun item 
            If Me.dgvItem.RowCount = 0 Then
                MsgBox("Please specify at least 1 item.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '2. sudah ada data yang sama period dan tipe budgetnya
            If Me.baru = True Then
                If Me.objSB.GetJml(Me.cboTipe.SelectedIndex, Me.dtpPeriod.Value) > 0 Then
                    MsgBox("Data duplication, the transaction: " + Me.objSB.GetNoTransaksi(Me.objSB.GetPrimary(Me.cboTipe.SelectedIndex, Me.dtpPeriod.Value)) + _
                    " is also " + Me.cboTipe.Text + " data. You cannot save this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                vPrimaryAcuan = Me.objSB.GetPrimary(Me.cboTipe.SelectedIndex, Me.dtpPeriod.Value, Me.txtPrimary.Text)
                If vPrimaryAcuan <> -1 Then
                    MsgBox("Data duplication, the transaction: " + Me.objSB.GetNoTransaksi(vPrimaryAcuan) + _
                    " is also " + Me.cboTipe.Text + " data. You cannot save this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            '3. batasan untuk no transaksi
            If Me.baru = True Then
                If Me.objSB.ApakahSudahAda(Me.txtDocNo.Text) = True Then
                    Me.txtDocNo.Text = Me.objSB.GetNoBaru(Me.dtpDate.Value)
                End If
            Else
                If Me.objSB.GetPrimary(Me.txtDocNo.Text) <> Me.txtPrimary.Text Then
                    Me.txtDocNo.Text = Me.objSB.GetNoBaru(Me.dtpDate.Value)
                End If
            End If

            '4. untuk setiap baris, minimal ada 1 bulan yang tidak nol
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Index <> Me.dgvItem.Rows.Count - 1 Then
                    If CDec(rowDGV.Cells(Me.ColJan.Index).Value) = 0 And CDec(rowDGV.Cells(Me.ColFeb.Index).Value) = 0 And _
                    CDec(rowDGV.Cells(Me.ColMar.Index).Value) = 0 And CDec(rowDGV.Cells(Me.ColApr.Index).Value) = 0 And _
                    CDec(rowDGV.Cells(Me.ColMay.Index).Value) = 0 And CDec(rowDGV.Cells(Me.ColJun.Index).Value) = 0 And _
                    CDec(rowDGV.Cells(Me.ColJul.Index).Value) = 0 And CDec(rowDGV.Cells(Me.ColAug.Index).Value) = 0 And _
                    CDec(rowDGV.Cells(Me.ColSep.Index).Value) = 0 And CDec(rowDGV.Cells(Me.ColOct.Index).Value) = 0 And _
                    CDec(rowDGV.Cells(Me.ColNov.Index).Value) = 0 And CDec(rowDGV.Cells(Me.ColDec.Index).Value) = 0 Then
                        DummyInt += 1
                        If ListSalesPName = "" Then
                            ListSalesPName = DummyInt.ToString + "." + rowDGV.Cells(Me.ColSalesPersonName.Index).Value.ToString
                        Else
                            ListSalesPName = ListSalesPName + vbCrLf + DummyInt.ToString + "." + rowDGV.Cells(Me.ColSalesPersonName.Index).Value.ToString
                        End If
                    End If
                End If
            Next
            If ListSalesPName <> "" Then
                MsgBox("The following sales person(s) has invalid data: " + vbCrLf + ListSalesPName, MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.objSB
                .DocumentNo = Me.txtDocNo.Text
                .TipeBudget = Me.cboTipe.SelectedIndex
                .Remarks = Me.txtRemarks.Text
                .TanggalTransaksi = Me.dtpDate.Value
                .Period = Me.dtpPeriod.Value
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.objSB
                Me.txtPrimary.Text = .Primary
                Me.txtDocNo.Text = .DocumentNo
                Me.cboTipe.SelectedIndex = CInt(.TipeBudget)
                Me.txtRemarks.Text = .Remarks
                Me.dtpDate.Value = .TanggalTransaksi
                Me.dtpPeriod.Value = .Period
                Me.HitungTotal()

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
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Control.CheckForIllegalCrossThreadCalls = False
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
            Me.dtpPeriod.Value = Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpPeriod.Value)
            Me.DefaultPrecision = Me.ObjCurrency.GetDecimal_DefaultCurrency
            Control.CheckForIllegalCrossThreadCalls = False
            Me.dtpPeriod.Value = Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpPeriod.Value)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOk.Click
        Dim hasil As MsgBoxResult

        Try
            If Me.ApakahDataBenar() = True Then
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Me.SalinKeObjek()
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        Me.dgvItem.AllowUserToAddRows = False
                        If Me.objSB.SimpanBaru(Me.dgvItem, Me.UserCode, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                            Exit Sub
                        End If
                        Me.dgvItem.AllowUserToAddRows = True
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Me.SalinKeObjek()
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        Me.dgvItem.AllowUserToAddRows = False
                        If Me.objSB.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                            Exit Sub
                        End If
                        Me.dgvItem.AllowUserToAddRows = False
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult

        Try
            If Me.txtRemarks.Text = "" And Me.dgvItem.RowCount = 1 Then
                Me.KosongkanField()
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.objSB.CariDetail(Me.txtPrimary.Text, Me.dgvItem, Me)
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

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()

            Me.objSB.ResetObject(Me.dgvItem)
            Me.txtDocNo.Text = Me.objSB.GetNoBaru(Me.dtpDate.Value)
            Me.dtpPeriod.Value = Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpPeriod.Value)

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmTransaksiSalesBudget_cari
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.PrimaryTransaksi = -1 Then
                If Me.objSB.CariDetail(hj.PrimaryTransaksi, Me.dgvItem, Me) _
                = True Then
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
            If Me.txtDocNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtDocNo.Text <> "" Then
                If Me.objSB.ApakahBolehEdit(Me.objSB.Primary, Me.StrictEdit, Me.UserCode) = True Then
                    Me.baru = False
                    Me.objSB.CariDetail(Me.objSB.Primary, Me.dgvItem, Me)
                    Me.SalinKeField()
                    Me.AktifkanDetail()
                    Me.dgvItem.Focus()
                    'nilai min juga harus diatas tanggal pengkuncian data terakhir
                    'msgbox("belum")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult

        Try
            If Me.txtDocNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtDocNo.Text <> "" Then
                If Me.objSB.ApakahBolehHapus(Me.objSB.Primary, Me.StrictDelete, Me.UserCode) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objSB.Hapus(Me.txtPrimary.Text, Me.StrictDelete, Me.UserCode) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
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

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.txtRemarks.Enabled = True Then
                Me.BringToFront()
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objSB.SimpanBaru(Me.dgvItem, Me.UserCode, Me) = True Then
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
                            If Me.objSB.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me) = True Then
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

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColSalesPersonCode.Index Then
                If Me.dgvItem.Rows.Count - 1 = e.RowIndex Then
                    'tidak bisa edit baris terakhir klo bulan kolom sales person code
                    e.Cancel = True
                    Exit Sub
                Else
                    If e.ColumnIndex = Me.ColRemarks.Index Or e.ColumnIndex = Me.ColJan.Index Or _
                    e.ColumnIndex = Me.ColFeb.Index Or e.ColumnIndex = Me.ColMar.Index Or _
                    e.ColumnIndex = Me.ColApr.Index Or e.ColumnIndex = Me.ColMay.Index Or _
                    e.ColumnIndex = Me.ColJun.Index Or e.ColumnIndex = Me.ColJul.Index Or _
                    e.ColumnIndex = Me.ColAug.Index Or e.ColumnIndex = Me.ColSep.Index Or _
                    e.ColumnIndex = Me.ColOct.Index Or e.ColumnIndex = Me.ColNov.Index Or _
                    e.ColumnIndex = Me.ColDec.Index Then
                        'yang bisa diedit, kolom remarks + jan-dec
                        Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                        Exit Sub
                    End If
                End If
            Else
                'col sales person code
                If e.ColumnIndex = Me.ColSalesPersonCode.Index Then
                    Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungSubTotal(ByVal vRowIndex As Integer)
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgvItem.Rows(vRowIndex)
            rowDGV.Cells(Me.ColTotal.Index).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvItem.Item(Me.ColJan.Index, vRowIndex).Value) + _
            CDec(Me.dgvItem.Item(Me.ColFeb.Index, vRowIndex).Value) + _
            CDec(Me.dgvItem.Item(Me.ColMar.Index, vRowIndex).Value) + _
            CDec(Me.dgvItem.Item(Me.ColApr.Index, vRowIndex).Value) + _
            CDec(Me.dgvItem.Item(Me.ColMay.Index, vRowIndex).Value) + _
            CDec(Me.dgvItem.Item(Me.ColJun.Index, vRowIndex).Value) + _
            CDec(Me.dgvItem.Item(Me.ColJul.Index, vRowIndex).Value) + _
            CDec(Me.dgvItem.Item(Me.ColAug.Index, vRowIndex).Value) + _
            CDec(Me.dgvItem.Item(Me.ColSep.Index, vRowIndex).Value) + _
            CDec(Me.dgvItem.Item(Me.ColOct.Index, vRowIndex).Value) + _
            CDec(Me.dgvItem.Item(Me.ColNov.Index, vRowIndex).Value) + _
            CDec(Me.dgvItem.Item(Me.ColDec.Index, vRowIndex).Value), Me.DefaultPrecision)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungTotal()
        Dim rowDGV As DataGridViewRow
        Dim FinalIndex As Integer = 0
        Dim RetVal As Decimal = 0

        Try
            If Me.dgvItem.AllowUserToAddRows = True Then
                FinalIndex = Me.dgvItem.Rows.Count - 2
            Else
                FinalIndex = Me.dgvItem.Rows.Count - 1
            End If

            For i As Integer = 0 To FinalIndex
                rowDGV = Me.dgvItem.Rows(i)
                RetVal += rowDGV.Cells(Me.ColTotal.Index).Value
            Next

            Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(RetVal, Me.DefaultPrecision)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CopyBaris(ByVal vIndexAcuan As Integer, ByVal vIndexTarget As Integer)
        Dim rowTarget, rowAcuan As DataGridViewRow

        Try
            rowTarget = Me.dgvItem.Rows(vIndexTarget)
            rowAcuan = Me.dgvItem.Rows(vIndexAcuan)

            With rowTarget
                rowAcuan.Cells(Me.ColJan.Index).Value = rowTarget.Cells(Me.ColJan.Index).Value
                rowAcuan.Cells(Me.ColFeb.Index).Value = rowTarget.Cells(Me.ColFeb.Index).Value
                rowAcuan.Cells(Me.ColMar.Index).Value = rowTarget.Cells(Me.ColMar.Index).Value
                rowAcuan.Cells(Me.ColApr.Index).Value = rowTarget.Cells(Me.ColApr.Index).Value
                rowAcuan.Cells(Me.ColMay.Index).Value = rowTarget.Cells(Me.ColMay.Index).Value
                rowAcuan.Cells(Me.ColJun.Index).Value = rowTarget.Cells(Me.ColJun.Index).Value
                rowAcuan.Cells(Me.ColJul.Index).Value = rowTarget.Cells(Me.ColJul.Index).Value
                rowAcuan.Cells(Me.ColAug.Index).Value = rowTarget.Cells(Me.ColAug.Index).Value
                rowAcuan.Cells(Me.ColSep.Index).Value = rowTarget.Cells(Me.ColSep.Index).Value
                rowAcuan.Cells(Me.ColOct.Index).Value = rowTarget.Cells(Me.ColOct.Index).Value
                rowAcuan.Cells(Me.ColNov.Index).Value = rowTarget.Cells(Me.ColNov.Index).Value
                rowAcuan.Cells(Me.ColDec.Index).Value = rowTarget.Cells(Me.ColDec.Index).Value
                rowAcuan.Cells(Me.ColTotal.Index).Value = rowTarget.Cells(Me.ColTotal.Index).Value
            End With

            Me.HitungTotal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True _
            Or Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColRemarks.Index Then
                    'Remarks
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Exit Sub
                ElseIf e.ColumnIndex = Me.ColSalesPersonCode.Index Then
                    'kalo kode sales person maka hapus satu bari semuanya
                    If e.RowIndex <> Me.dgvItem.RowCount - 1 Then
                        Me.RemoveIndex = e.RowIndex
                    End If
                    Exit Sub
                ElseIf e.ColumnIndex = Me.ColJan.Index Or _
                e.ColumnIndex = Me.ColFeb.Index Or e.ColumnIndex = Me.ColMar.Index Or _
                e.ColumnIndex = Me.ColApr.Index Or e.ColumnIndex = Me.ColMay.Index Or _
                e.ColumnIndex = Me.ColJun.Index Or e.ColumnIndex = Me.ColJul.Index Or _
                e.ColumnIndex = Me.ColAug.Index Or e.ColumnIndex = Me.ColSep.Index Or _
                e.ColumnIndex = Me.ColOct.Index Or e.ColumnIndex = Me.ColNov.Index Or _
                e.ColumnIndex = Me.ColDec.Index Then
                    'untuk kolom debet kredit, langsung di balikin ke nilai semula
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
            Else
                ''cek budget jan - dec
                If e.ColumnIndex = Me.ColJan.Index Or _
                e.ColumnIndex = Me.ColFeb.Index Or e.ColumnIndex = Me.ColMar.Index Or _
                e.ColumnIndex = Me.ColApr.Index Or e.ColumnIndex = Me.ColMay.Index Or _
                e.ColumnIndex = Me.ColJun.Index Or e.ColumnIndex = Me.ColJul.Index Or _
                e.ColumnIndex = Me.ColAug.Index Or e.ColumnIndex = Me.ColSep.Index Or _
                e.ColumnIndex = Me.ColOct.Index Or e.ColumnIndex = Me.ColNov.Index Or _
                e.ColumnIndex = Me.ColDec.Index Then
                    If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, _
                            e.RowIndex).Value) = False Then
                        MsgBox("Please enter positive number", MsgBoxStyle.Information, _
                        "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefaultPrecision)
                    'hitung sub total
                    Me.HitungSubTotal(e.RowIndex)
                    'hitung total
                    Me.HitungTotal()
                    'beres
                    Exit Sub
                ElseIf e.ColumnIndex = Me.ColSalesPersonCode.Index Then
                    ''data sales person code
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    'kemungkinan
                    If Me.ObjSalesP.ApakahSudahAda(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        '1. kode sales person yang diketik ada --> load data deskripsi + primary + group sales person
                        Me.dgvItem.Item(Me.ColPrimarySalesPerson.Index, e.RowIndex).Value = Me.ObjSalesP.GetPrimary(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvItem.Item(Me.ColSalesPersonName.Index, e.RowIndex).Value = Me.ObjSalesP.GetName(Me.dgvItem.Item(Me.ColPrimarySalesPerson.Index, e.RowIndex).Value)
                        Me.dgvItem.Item(Me.ColSalesPGroup.Index, e.RowIndex).Value = Me.ObjSalesPGroup.GetDesc(Me.ObjSalesP.GetPrimaryGroup(Me.dgvItem.Item(Me.ColPrimarySalesPerson.Index, e.RowIndex).Value))
                        'copy baris data atasnya
                        If e.RowIndex = Me.dgvItem.Rows.Count - 2 And e.RowIndex <> 0 Then
                            Me.CopyBaris(e.RowIndex, e.RowIndex - 1)
                        End If
                    Else
                        '2. kode akun yang diketik tidak ada --> tampilkan form cari
                        Dim ij As New Core_Module.FrmMasterSalesPerson_Cari
                        ij.txtCari.Text = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.PrimarySalesPerson <> -1 Then
                            '2.a. kode akun yang dicari ada --> load data kode + deskripsi + primary + nama group
                            Me.dgvItem.Item(Me.ColPrimarySalesPerson.Index, e.RowIndex).Value = ij.PrimarySalesPerson
                            Me.dgvItem.Item(Me.ColSalesPersonName.Index, e.RowIndex).Value = Me.ObjSalesP.GetName(Me.dgvItem.Item(Me.ColPrimarySalesPerson.Index, e.RowIndex).Value)
                            Me.dgvItem.Item(Me.ColSalesPGroup.Index, e.RowIndex).Value = Me.ObjSalesPGroup.GetDesc(Me.ObjSalesP.GetPrimaryGroup(Me.dgvItem.Item(Me.ColPrimarySalesPerson.Index, e.RowIndex).Value))
                            Me.dgvItem.Item(Me.ColSalesPersonCode.Index, e.RowIndex).Value = Me.ObjSalesP.GetCode(ij.PrimarySalesPerson)
                            'copy baris data atasnya
                            If e.RowIndex = Me.dgvItem.Rows.Count - 2 And e.RowIndex <> 0 Then
                                Me.CopyBaris(e.RowIndex, e.RowIndex - 1)
                            End If
                        ElseIf ij.PrimarySalesPerson = -1 Then
                            '2.b. kode akun yang dicari tidak ada --> kembalikan ke kondisi buffer
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                Me.RemoveIndex = e.RowIndex
                                Me.BackgroundWorker2.RunWorkerAsync()
                                Exit Sub
                            End If
                        End If
                        ij = Nothing
                    End If
                ElseIf e.ColumnIndex = Me.ColRemarks.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                        If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("%") Or _
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("*") Or _
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("#") Or _
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("'") Then
                            MsgBox("Please avoid using the following character: %, *, ', #", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsAdded(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Dim rowDgv As DataGridViewRow

        Try
            Me.dgvItem.TabStop = True
            If Me.txtRemarks.Enabled = True Then
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If

            If Me.txtRemarks.Enabled = True Then
                If e.RowIndex = Me.dgvItem.RowCount - 1 Then
                    rowDgv = Me.dgvItem.Rows(e.RowIndex)
                    With rowDgv
                        .Cells(Me.ColJan.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                        .Cells(Me.ColFeb.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                        .Cells(Me.ColMar.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                        .Cells(Me.ColApr.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                        .Cells(Me.ColMay.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                        .Cells(Me.ColJun.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                        .Cells(Me.ColJul.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                        .Cells(Me.ColAug.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                        .Cells(Me.ColSep.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                        .Cells(Me.ColOct.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                        .Cells(Me.ColNov.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                        .Cells(Me.ColDec.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                        .Cells(Me.ColNo.Index).Value = Me.dgvItem.RowCount
                        .Cells(Me.ColPrimarySalesPerson.Index).Value = -1
                        .Cells(Me.ColRemarks.Index).Value = ""
                        .Cells(Me.ColSalesPersonCode.Index).Value = ""
                        .Cells(Me.ColSalesPersonName.Index).Value = ""
                        .Cells(Me.ColSalesPGroup.Index).Value = ""
                        .Cells(Me.ColTotal.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                    End With
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItem.RowsRemoved
        Try
            If Me.dgvItem.Rows.Count = 0 Then
                Me.dgvItem.TabStop = False
                If Me.txtRemarks.Enabled = True Then
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
            End If

            If Me.txtRemarks.Enabled = True Then
                Me.HitungTotal()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.SelectionChanged
        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.dgvItem.SelectedRows.Count > 0 Then
                    Me.btnUP.Enabled = True
                    Me.btnDown.Enabled = True
                Else
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Me.ObjCD.SetNumber(Me.dgvItem)
    End Sub

    ''' <summary>
    ''' tampilkan row yang diberikan pada parameter
    ''' </summary>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Private Sub ShowRow(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < Me.dgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.dgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While Me.dgvItem.Rows(row.Index).Displayed = False
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False) Then
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUP.Click
        Try
            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItem.SelectedRows(0)
                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.dgvItem.Rows.Remove(row)
                    Me.dgvItem.Rows.InsertRange(indexbaru, row)
                    Me.dgvItem.Focus()
                    Me.dgvItem.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItem)
                    If row.Displayed = False Then
                        Me.ShowRow(row, True)
                    End If
                End If
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItem.SelectedRows(0)
                If row.Index <> Me.dgvItem.Rows.Count - 1 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.dgvItem.Rows.Remove(row)
                    Me.dgvItem.Rows.InsertRange(indexbaru, row)
                    Me.dgvItem.Focus()
                    Me.dgvItem.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItem)
                    Me.ShowRow(row, False)
                End If
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
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

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemarks.GotFocus
        Try
            Me.txtRemarks.BackColor = BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemarks.LostFocus
        Try
            Me.txtRemarks.BackColor = Me.txtPrimary.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDocNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDocNo.GotFocus
        If Me.txtDocNo.ReadOnly = False And Me.txtDocNo.Enabled = True Then
            txtDocNo.BackColor = BackColorAll
        End If
    End Sub

    Private Sub txtDocNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDocNo.LostFocus
        If Me.txtDocNo.ReadOnly = False And Me.txtDocNo.Enabled = True Then
            txtDocNo.BackColor = Me.txtPrimary.BackColor
        End If
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPrint.Click
        Dim DS As New DseTransaksiSalesBudget
        Dim rpt As New CrpCetakTransaksiSalesBudget
        Dim FrmCetak As New FrmTransaksiSalesBudget_Cetak

        Try
            If Me.txtDocNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtDocNo.Text <> "" Then

                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                Me.objSB.Cetak(Me.txtPrimary.Text, FrmCetak, rpt, DS)
                FrmCetak.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            rpt.Dispose()
            FrmCetak.Dispose()
        End Try
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            Me.BackgroundWorker1.RunWorkerAsync()
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
            If Me.txtRemarks.Enabled = True Then
                If Me.baru = False Then
                    If Me.objSB.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                        'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                        Me.txtDocNo.Text = Me.objSB.DocumentNo
                    Else
                        'jika pindah bulan maka akan ganti no
                        Me.txtDocNo.Text = Me.objSB.GetNoBaru(Me.dtpDate.Value)
                    End If
                Else
                    Me.txtDocNo.Text = Me.objSB.GetNoBaru(Me.dtpDate.Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(10)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            If Me.RemoveIndex <> -1 Then
                Me.dgvItem.Rows.RemoveAt(Me.RemoveIndex)
                Me.RemoveIndex = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim MinRows As Integer = 0

        Try
            If Me.dgvItem.AllowUserToAddRows = True Then
                MinRows = 2
            Else
                MinRows = 1
            End If
            If Me.dgvItem.Rows.Count >= MinRows Then
                Dim ij As New FrmTransaksiSalesBudget_DetailBudget
                ij.RefForm = Me
                ij.DefPrec = Me.DefaultPrecision
                ij.ShowDialog()
                ij = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class