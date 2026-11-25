Option Explicit On
Option Strict Off

Imports Core_Module
Public Class FrmMasterAkun
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = "USER"
    Public ObjAkun As New ClsMasterAkun

    Dim ObjCD As New ClsCekData
    Dim baru As Boolean
    Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim Def_Prec As Integer = 0
    Dim LevelAkunMax As Integer = 8

    Private Sub FrmMasterAkun_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        'Try
        '    Me.Def_Prec = Me.ObjCur.GetDecimal_DefaultCurrency
        '    ObjAkun.IsiDGVItem(Me.dgvAkun, Me, Me.Def_Prec)
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub FrmMasterAkun_Header_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ''Dim ij As New System.Drawing.Font("Microsoft Sans Serif", 8, FontStyle.Bold)
        ''Me.dgvAkun.Rows(0).Cells(0).Style.Font = ij
        Dim ij As New FrmMasterAkun_Detail

        Try
            Me.Def_Prec = Me.ObjCur.GetDecimal_DefaultCurrency
            Me.dgvAkun.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            ij.UserCode = Me.UserCode
            ij.BackColorAll = Me.BackColorAll
            ij.Baru = True
            'load data dari database dan salin ke DGV
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            ObjAkun.IsiDGVItem(Me.dgvAkun, Me, Me.Def_Prec)
            'aktifkan detail
            ij.AktifkanDetail()
            Me.dgvAkun.ClearSelection()
            Me.ShowIcon = False
            Me.MinimizeBox = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Dim FrmParent As Windows.Forms.Form = Me.MdiParent
        Dim RowDgv As DataGridViewRow
        Dim counter As Double = 0
        Dim counter2 As Double = 0

        Try
            'dapetin yg lagi disorot turunkan kode
            If Me.dgvAkun.SelectedRows.Count > 0 Then
                RowDgv = Me.dgvAkun.SelectedRows(0)
                Dim KodeWarisan As String = Me.dgvAkun.SelectedRows(0).Cells(Me.ColAccount.Index).Value
                ''Coding Form show
                Dim ij As FrmMasterAkun_Detail
                Dim j As Windows.Forms.Form
                For Each j In FrmParent.MdiChildren
                    ij = TryCast(j, FrmMasterAkun_Detail)
                    If IsNothing(ij) = False Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                Next

                ' Increment the caption counter.
                Dim i As New FrmMasterAkun_Detail
                i.UserCode = Me.UserCode
                i.Def_Prec = Me.Def_Prec
                ' Set the caption to be unique.
                i.BackColorAll = Core_Module.GetSMESSBackColorAll

                ''sesuaikan kondisi awal
                With i
                    .Baru = True
                    .KosongkanField()
                    .AktifkanDetail()
                    .TxtKodeAkun1.TabStop = True
                    .tslCreatedBy.Text = "Created By: " + Me.UserCode.ToString
                    .tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
                    .tslModifiedBy.Text = "Modified By: " + Me.UserCode.ToString
                    .tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
                    .ObjAkun.TipePerkiraan = RowDgv.Cells(Me.ColTipePerkiraan.Index).Value
                    'jika baru maka default currencynya adalah IDR
                    .ObjAkun.PrimaryCurrency = 1
                    .ObjAkun.BeginningBalance = 0
                    .ObjAkun.BeginningBalanceOriginalCurrency = 0
                    .ObjAkun.TipeAkun2 = Me.ObjAkun.GetTipeAkun2Default(.ObjAkun.TipePerkiraan)
                    .ObjAkun.Header = 0

                    If RowDgv.Cells(Me.ColLevel.Index).Value = 1 Then
                        .ObjAkun.LevelPerkiraan = 2
                    Else
                        If RowDgv.Cells(Me.ColHeader.Index).Value = 1 Then
                            .ObjAkun.LevelPerkiraan = RowDgv.Cells(Me.ColLevel.Index).Value + 1
                        Else
                            .ObjAkun.LevelPerkiraan = RowDgv.Cells(Me.ColLevel.Index).Value
                        End If
                    End If

                    .ObjAkun.SaldoNormal = Me.ObjAkun.GetSaldoNormal(RowDgv.Cells(Me.ColPrimaryAkun.Index).Value)
                    .ObjAkun.Urutan = RowDgv.Cells(Me.ColUrutan.Index).Value + 1
                    If RowDgv.Cells(Me.ColLevel.Index).Value = 1 Then
                        .ObjAkun.PrimaryLv1 = RowDgv.Cells(Me.ColPrimaryAkun.Index).Value
                        .ObjAkun.PrimaryLv2 = -1
                        .ObjAkun.PrimaryLv3 = -1
                        .ObjAkun.PrimaryLvl4 = -1
                        .ObjAkun.PrimaryLvl5 = -1
                        .ObjAkun.PrimaryLvl6 = -1
                        .ObjAkun.PrimaryLvl7 = -1
                    Else
                        If RowDgv.Cells(Me.ColHeader.Index).Value = 1 Then
                            .ObjAkun.PrimaryLv1 = RowDgv.Cells(Me.ColPrimaryLvl1.Index).Value
                            If RowDgv.Cells(Me.ColLevel.Index).Value = 2 Then
                                .ObjAkun.PrimaryLv2 = RowDgv.Cells(Me.ColPrimaryAkun.Index).Value
                                .ObjAkun.PrimaryLv3 = -1
                                .ObjAkun.PrimaryLvl4 = -1
                                .ObjAkun.PrimaryLvl5 = -1
                                .ObjAkun.PrimaryLvl6 = -1
                                .ObjAkun.PrimaryLvl7 = -1
                            ElseIf RowDgv.Cells(Me.ColLevel.Index).Value = 3 Then
                                .ObjAkun.PrimaryLv2 = RowDgv.Cells(Me.ColPrimaryLvl2.Index).Value
                                .ObjAkun.PrimaryLv3 = RowDgv.Cells(Me.ColPrimaryAkun.Index).Value
                                .ObjAkun.PrimaryLvl4 = -1
                                .ObjAkun.PrimaryLvl5 = -1
                                .ObjAkun.PrimaryLvl6 = -1
                                .ObjAkun.PrimaryLvl7 = -1
                            ElseIf RowDgv.Cells(Me.ColLevel.Index).Value = 4 Then
                                .ObjAkun.PrimaryLv2 = RowDgv.Cells(Me.ColPrimaryLvl2.Index).Value
                                .ObjAkun.PrimaryLv3 = RowDgv.Cells(Me.ColPrimaryLvl3.Index).Value
                                .ObjAkun.PrimaryLvl4 = RowDgv.Cells(Me.ColPrimaryAkun.Index).Value
                                .ObjAkun.PrimaryLvl5 = -1
                                .ObjAkun.PrimaryLvl6 = -1
                                .ObjAkun.PrimaryLvl7 = -1
                            ElseIf RowDgv.Cells(Me.ColLevel.Index).Value = 5 Then
                                .ObjAkun.PrimaryLv2 = RowDgv.Cells(Me.ColPrimaryLvl2.Index).Value
                                .ObjAkun.PrimaryLv3 = RowDgv.Cells(Me.ColPrimaryLvl3.Index).Value
                                .ObjAkun.PrimaryLvl4 = RowDgv.Cells(Me.ColPrimaryLvl4.Index).Value
                                .ObjAkun.PrimaryLvl5 = RowDgv.Cells(Me.ColPrimaryAkun.Index).Value
                                .ObjAkun.PrimaryLvl6 = -1
                                .ObjAkun.PrimaryLvl7 = -1
                            ElseIf RowDgv.Cells(Me.ColLevel.Index).Value = 6 Then
                                .ObjAkun.PrimaryLv2 = RowDgv.Cells(Me.ColPrimaryLvl2.Index).Value
                                .ObjAkun.PrimaryLv3 = RowDgv.Cells(Me.ColPrimaryLvl3.Index).Value
                                .ObjAkun.PrimaryLvl4 = RowDgv.Cells(Me.ColPrimaryLvl4.Index).Value
                                .ObjAkun.PrimaryLvl5 = RowDgv.Cells(Me.ColPrimaryLvl5.Index).Value
                                .ObjAkun.PrimaryLvl6 = RowDgv.Cells(Me.ColPrimaryAkun.Index).Value
                                .ObjAkun.PrimaryLvl7 = -1
                            ElseIf RowDgv.Cells(Me.ColLevel.Index).Value = 7 Then
                                .ObjAkun.PrimaryLv2 = RowDgv.Cells(Me.ColPrimaryLvl2.Index).Value
                                .ObjAkun.PrimaryLv3 = RowDgv.Cells(Me.ColPrimaryLvl3.Index).Value
                                .ObjAkun.PrimaryLvl4 = RowDgv.Cells(Me.ColPrimaryLvl4.Index).Value
                                .ObjAkun.PrimaryLvl5 = RowDgv.Cells(Me.ColPrimaryLvl5.Index).Value
                                .ObjAkun.PrimaryLvl6 = RowDgv.Cells(Me.ColPrimaryLvl6.Index).Value
                                .ObjAkun.PrimaryLvl7 = RowDgv.Cells(Me.ColPrimaryAkun.Index).Value
                            End If
                        Else
                            .ObjAkun.PrimaryLv1 = RowDgv.Cells(Me.ColPrimaryLvl1.Index).Value
                            .ObjAkun.PrimaryLv2 = RowDgv.Cells(Me.ColPrimaryLvl2.Index).Value
                            .ObjAkun.PrimaryLv3 = RowDgv.Cells(Me.ColPrimaryLvl3.Index).Value
                            .ObjAkun.PrimaryLvl4 = RowDgv.Cells(Me.ColPrimaryLvl4.Index).Value
                            .ObjAkun.PrimaryLvl5 = RowDgv.Cells(Me.ColPrimaryLvl5.Index).Value
                            .ObjAkun.PrimaryLvl6 = RowDgv.Cells(Me.ColPrimaryLvl6.Index).Value
                            .ObjAkun.PrimaryLvl7 = RowDgv.Cells(Me.ColPrimaryLvl7.Index).Value
                        End If
                    End If
                End With

                If Me.dgvAkun.SelectedRows(0).Cells(Me.ColLevel.Index).Value = 1 Then
                    i.TxtKodeAkun1.Text = KodeWarisan.Substring(0, 1)
                ElseIf Me.dgvAkun.SelectedRows(0).Cells(Me.ColLevel.Index).Value = 2 Then
                    i.TxtKodeAkun1.Text = KodeWarisan.Substring(5, 1)
                ElseIf Me.dgvAkun.SelectedRows(0).Cells(Me.ColLevel.Index).Value = 3 Then
                    i.TxtKodeAkun1.Text = KodeWarisan.Substring(10, 1)
                ElseIf Me.dgvAkun.SelectedRows(0).Cells(Me.ColLevel.Index).Value = 4 Then
                    i.TxtKodeAkun1.Text = KodeWarisan.Substring(15, 1)
                ElseIf Me.dgvAkun.SelectedRows(0).Cells(Me.ColLevel.Index).Value = 5 Then
                    i.TxtKodeAkun1.Text = KodeWarisan.Substring(20, 1)
                ElseIf Me.dgvAkun.SelectedRows(0).Cells(Me.ColLevel.Index).Value = 6 Then
                    i.TxtKodeAkun1.Text = KodeWarisan.Substring(25, 1)
                ElseIf Me.dgvAkun.SelectedRows(0).Cells(Me.ColLevel.Index).Value = 7 Then
                    i.TxtKodeAkun1.Text = KodeWarisan.Substring(30, 1)
                ElseIf Me.dgvAkun.SelectedRows(0).Cells(Me.ColLevel.Index).Value = 8 Then
                    i.TxtKodeAkun1.Text = KodeWarisan.Substring(35, 1)
                End If

                i.TxtKodeAkun1.Enabled = False
                i.TxtAccountNo.Focus()

                'jika akun header level 7 akan bikin anak, maka anaknya tidak boleh header juga
                If RowDgv.Cells(Me.ColHeader.Index).Value = 1 And RowDgv.Cells(Me.ColLevel.Index).Value = Me.LevelAkunMax - 1 Then
                    i.chkHeaderAccount.Checked = False
                    i.chkHeaderAccount.Enabled = False
                End If
                'jika akun non header level 8 akan bikin anak maka anaknya pasti bukan header
                If RowDgv.Cells(Me.ColHeader.Index).Value = 0 And RowDgv.Cells(Me.ColLevel.Index).Value = Me.LevelAkunMax Then
                    i.chkHeaderAccount.Checked = False
                    i.chkHeaderAccount.Enabled = False
                End If

                i.Show()
                i.SalinKeField()
                i.Hide()

                If i.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    counter = RowDgv.Index + 1
                    counter2 = RowDgv.Index + 1

                    If RowDgv.Displayed = True Then
                        While RowDgv.Displayed = True
                            counter += 1
                            If counter >= Me.dgvAkun.RowCount Then
                                counter = Me.dgvAkun.RowCount - 1
                                Exit While
                            End If
                            RowDgv = Me.dgvAkun.Rows(counter)
                        End While
                        counter -= 1
                    End If
                    Me.ObjAkun.IsiDGVItem(Me.dgvAkun, Me, Me.Def_Prec)

                    'tampilkan baris yang diinginkan
                    If counter <= Me.dgvAkun.Rows.Count - 1 Then
                        RowDgv = Me.dgvAkun.Rows(counter)
                        If RowDgv.Displayed = False Then
                            Me.ShowRow(RowDgv, False)
                        End If
                        Me.dgvAkun.ClearSelection()
                        RowDgv = Me.dgvAkun.Rows(counter2)
                        RowDgv.Selected = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Dim ij As New FrmMasterAkun_Detail
        Dim HeaderAktif As Boolean = False
        Dim rowDGV As DataGridViewRow
        Dim counter As Double = 0
        Dim counter2 As Double = 0

        Try
            If Me.dgvAkun.SelectedRows.Count > 0 Then
                ij.Def_Prec = Me.Def_Prec
                ij.UserCode = Me.UserCode
                ij.BackColorAll = Me.BackColorAll
                ij.Baru = False
                ' caridata
                If Me.ObjAkun.ApakahBolehEdit(Me.dgvAkun.SelectedRows(0) _
                .Cells(Me.ColPrimaryAkun.Index).Value) = True Then
                    rowDGV = Me.dgvAkun.SelectedRows(0)
                    ij.ObjAkun.CariDetail(Me.dgvAkun.SelectedRows(0).Cells(Me.ColPrimaryAkun.Index).Value)
                    'salin data
                    ij.Show()
                    ij.SalinKeField()
                    'atur object
                    ij.NonAktifkanDetail()
                    'aktifkan detail
                    ij.TxtAccountNo.Enabled = True
                    'ij.chkActiveAccount.Enabled = True
                    'ij.chkHeaderAccount.Enabled = True
                    ij.TxtDesc.Enabled = True
                    ij.btnOK.Enabled = True
                    ij.TxtRemarks.Enabled = True
                    ij.btnCancel.Enabled = True

                    If Me.dgvAkun.SelectedRows(0).Cells(Me.ColLevel.Index).Value <> 1 Then
                        'data berkaitan perubahan header ke non header
                        'A. Jika akun lvl 4 maka pasti non header
                        If Me.dgvAkun.SelectedRows(0).Cells(Me.ColLevel.Index).Value = Me.LevelAkunMax Then
                            HeaderAktif = False
                            ij.chkHeaderAccount.Enabled = False
                        Else
                            'B. Jika bukan lvl 4
                            If Me.dgvAkun.SelectedRows(0).Cells(Me.ColHeader.Index).Value = 0 Then
                                'jika non header
                                '1. status perubahan non header menjadi header, bisa dilakukan bila:
                                '   a. belum ada jurnal yang merujuk ke akun tsb
                                Dim ObjJurnal As New ClsTransaksiJournal
                                If ObjJurnal.GetJml_ByPrimaryAkun(Me.dgvAkun.SelectedRows(0).Cells(Me.ColPrimaryAkun.Index).Value) > 0 Then
                                    HeaderAktif = False
                                    ij.chkHeaderAccount.Enabled = False
                                End If
                                ObjJurnal = Nothing
                            Else
                                'jika header
                                '2. status perubahan header menjadi non header bisa dilakukan jika:
                                '   a. tidak ada anak yang merujuk ke akun ini
                                If Me.ObjAkun.GetJmlAnak(Me.dgvAkun.SelectedRows(0).Cells(Me.ColPrimaryAkun.Index).Value) > 0 Then
                                    HeaderAktif = False
                                    ij.chkHeaderAccount.Enabled = False
                                End If
                            End If
                        End If
                    Else
                        ij.chkActiveAccount.Enabled = False
                        ij.chkHeaderAccount.Enabled = False
                    End If

                    ij.Hide()
                    If ij.ShowDialog() = Windows.Forms.DialogResult.OK Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        counter = rowDGV.Index
                        counter2 = rowDGV.Index
                        While rowDGV.Displayed = True And counter + 1 < Me.dgvAkun.Rows.Count - 1
                            counter += 1
                            If counter = Me.dgvAkun.RowCount Then
                                counter = Me.dgvAkun.RowCount - 1
                                Exit While
                            End If
                            rowDGV = Me.dgvAkun.Rows(counter)
                        End While
                        counter -= 1

                        Me.ObjAkun.IsiDGVItem(Me.dgvAkun, Me, Me.Def_Prec)

                        'tampilkan baris yang diinginkan
                        If counter <= Me.dgvAkun.Rows.Count - 1 Then
                            rowDGV = Me.dgvAkun.Rows(counter)
                            If rowDGV.Displayed = False Then
                                Me.ShowRow(rowDGV, False)
                            End If
                            Me.dgvAkun.ClearSelection()
                            rowDGV = Me.dgvAkun.Rows(counter2)
                            rowDGV.Selected = True
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Dim counter, counter2 As Double
        Dim rowDGV As DataGridViewRow

        Try
            If Me.dgvAkun.SelectedRows.Count > 0 Then
                If Me.ObjAkun.ApakahBolehHapus(Me.dgvAkun.SelectedRows(0).Cells(Me.ColPrimaryAkun.Index).Value) = False Then
                    Me.btnDelete.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjAkun.Hapus(Me.dgvAkun.SelectedRows(0).Cells(Me.ColPrimaryAkun.Index).Value) = True Then
                            rowDGV = Me.dgvAkun.SelectedRows(0)
                            counter = rowDGV.Index
                            counter2 = rowDGV.Index

                            If counter >= Me.dgvAkun.RowCount Then
                                counter = Me.dgvAkun.RowCount - 1
                                counter2 = Me.dgvAkun.RowCount - 1
                            Else
                                While rowDGV.Displayed = True And counter + 1 < Me.dgvAkun.Rows.Count - 1
                                    counter += 1
                                    If counter = Me.dgvAkun.RowCount Then
                                        counter = Me.dgvAkun.RowCount - 1
                                        Exit While
                                    End If
                                    rowDGV = Me.dgvAkun.Rows(counter)
                                End While
                                counter -= 1
                            End If

                            Me.ObjAkun.IsiDGVItem(Me.dgvAkun, Me, Me.Def_Prec)

                            'tampilkan baris yang diinginkan
                            If counter <= Me.dgvAkun.Rows.Count - 1 Then
                                rowDGV = Me.dgvAkun.Rows(counter)
                                If rowDGV.Displayed = False Then
                                    Me.ShowRow(rowDGV, False)
                                End If
                                Me.dgvAkun.ClearSelection()
                                If counter2 >= Me.dgvAkun.RowCount Then
                                    counter2 = Me.dgvAkun.RowCount - 1
                                End If
                                rowDGV = Me.dgvAkun.Rows(counter2)
                                ' Me.dgvAkun.Focus()
                                rowDGV.Selected = True

                            End If

                            ' Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")

                            'Me.dgvAkun.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.Controls, False, True, True, True, False, True, False, True)
            Me.CancelButton = Me.btnClose
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            If Me.dgvAkun.Rows.Count = 0 Then
                Me.btnUp.Enabled = False
                Me.btnDown.Enabled = False
                Me.BtnRight.Enabled = False
                Me.btnLeft.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.CancelButton = Me.btnClose
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAkun_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvAkun.CellDoubleClick
        Try
            If Me.dgvAkun.SelectedRows.Count > 0 Then
                'If Me.dgvAkun.SelectedRows(0).Cells(Me.ColLevel.Index).Value <> 1 Then
                Me.btnEdit.PerformClick()
                'End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAkun_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvAkun.ColumnHeaderMouseClick
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.ObjAkun.IsiDGVItem(Me.dgvAkun, Me, Me.Def_Prec)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPrint.Click
        Dim i As New FrmMasterAkun_Cetak

        Try
            If Me.dgvAkun.Rows.Count > 0 Then
                Me.ObjAkun.CetakMasterAkun(i.crvReport)
                i.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub SetUrutan()
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvAkun.Rows
                rowDGV.Cells(Me.ColUrutan.Index).Value = rowDGV.Index
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Set_Down_Arrow()
        Dim rowDGV As DataGridViewRow
        Dim Level, RowIndex As Integer
        Dim Ketemu As Boolean = False

        Try
            rowDGV = Me.dgvAkun.Rows(Me.dgvAkun.SelectedCells(0).RowIndex)
            Level = rowDGV.Cells(Me.ColLevel.Index).Value
            RowIndex = rowDGV.Index
            If RowIndex = 0 Then
                Me.btnDown.Enabled = False
                Exit Sub
            End If

            '1. Akun level 1 tidak boleh turun
            If rowDGV.Cells(Me.ColLevel.Index).Value = 1 Then
                Me.btnDown.Enabled = False
                Exit Sub
            End If

            '2. jika bukan paling bawah maka boleh
            If rowDGV.Index = Me.dgvAkun.Rows.Count - 1 Then
                Me.btnDown.Enabled = False
                Exit Sub
            End If

            '3. asal bawahnya bukan akun level 1, maka boleh
            rowDGV = Me.dgvAkun.Rows(RowIndex + 1)
            If rowDGV.Cells(Me.ColLevel.Index).Value = 1 Then
                Me.btnDown.Enabled = False
                Exit Sub
            End If

            '4. khusus untuk header, kalo bawah dari anak paling bawah bukan akun lvl 1 maka boleh
            For i As Integer = Me.dgvAkun.SelectedRows(0).Index + 1 To Me.dgvAkun.RowCount - 1
                rowDGV = Me.dgvAkun.Rows(i)
                If rowDGV.Cells(Me.ColLevel.Index).Value <= Me.dgvAkun.SelectedRows(0).Cells(Me.ColLevel.Index).Value Then
                    If rowDGV.Cells(Me.ColLevel.Index).Value = 1 Then
                        Me.btnDown.Enabled = False
                        Exit Sub
                    End If
                    Exit For
                Else
                    '5. Jika dia ada header yang paling bawah, juga tidak boleh
                    If rowDGV.Index = Me.dgvAkun.Rows.Count - 1 Then
                        Me.btnDown.Enabled = False
                        Exit Sub
                    End If
                End If
            Next

            Me.btnDown.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Set_Left_Arrow()
        Dim rowDGV As DataGridViewRow
        Dim Level, RowIndex As Integer
        Dim Ketemu As Boolean = False

        Try
            rowDGV = Me.dgvAkun.Rows(Me.dgvAkun.SelectedCells(0).RowIndex)
            Level = rowDGV.Cells(Me.ColLevel.Index).Value
            RowIndex = rowDGV.Index
            If RowIndex = 0 Then
                Me.btnLeft.Enabled = False
                Exit Sub
            End If

            '1. Akun level 1 & 2 tidak boleh geser kiri
            If rowDGV.Cells(Me.ColLevel.Index).Value = 1 Or rowDGV.Cells(Me.ColLevel.Index).Value = 2 Then
                Me.btnLeft.Enabled = False
                Exit Sub
            End If

            Me.btnLeft.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Set_Right_Arrow()
        Dim rowDGV As DataGridViewRow
        Dim Level, RowIndex, Header_Int As Integer
        Dim Ketemu As Boolean = False

        Try
            rowDGV = Me.dgvAkun.Rows(Me.dgvAkun.SelectedCells(0).RowIndex)
            Level = rowDGV.Cells(Me.ColLevel.Index).Value
            Header_Int = rowDGV.Cells(Me.ColHeader.Index).Value
            RowIndex = rowDGV.Index
            If RowIndex = 0 Then
                Me.BtnRight.Enabled = False
                Exit Sub
            End If

            '1. Akun level 1 tidak boleh geser kanan
            If rowDGV.Cells(Me.ColLevel.Index).Value = 1 Then
                Me.BtnRight.Enabled = False
                Exit Sub
            End If

            '2. Jika dilacak ke-atas ada akun header yang levelnya 
            '   sama dengan level saat ini maka boleh, jika tidak, tidak boleh
            '    pelacakan tidak boleh melewati akun level 1
            Level = Me.dgvAkun.SelectedRows(0).Cells(Me.ColLevel.Index).Value
            'jika level akun tepat diatasnya sama atau lebih kecil maka tidak bisa
            If Me.dgvAkun.Rows(Me.dgvAkun.SelectedRows(0).Index - 1).Cells(Me.ColLevel.Index).Value <= Level Then
                rowDGV = Me.dgvAkun.Rows(Me.dgvAkun.SelectedRows(0).Index - 1)
                If rowDGV.Cells(Me.ColLevel.Index).Value = Level And rowDGV.Cells(Me.ColHeader.Index).Value = 1 Then
                    'ok boleh geser kanan
                    'jika akkunnya adalah akun lvl 3 dan merupakan akun header, maka tidak bisa geser kanan lagi
                    If Level = Me.LevelAkunMax - 1 And Header_Int = 1 Then
                        Me.BtnRight.Enabled = False
                        Exit Sub
                    End If
                Else
                    Me.BtnRight.Enabled = False
                    Exit Sub
                End If
            Else
                'jika level akun ditasnya lebih besar maka cari terus ke atasnya
                For i As Integer = Me.dgvAkun.SelectedRows(0).Index - 1 To 0 Step -1
                    rowDGV = Me.dgvAkun.Rows(i)
                    If rowDGV.Cells(Me.ColLevel.Index).Value = 1 Then
                        Exit For
                    Else
                        If rowDGV.Cells(Me.ColLevel.Index).Value = Level And rowDGV.Cells(Me.ColHeader.Index).Value = 1 Then
                            Ketemu = True
                            Exit For
                        End If
                    End If
                Next
                If Ketemu = False Then
                    Me.BtnRight.Enabled = False
                    Exit Sub
                End If
            End If

            '3. jika akun header
            '   a. Jika anaknya ada yang level akun max, maka tidak boleh
            If Me.dgvAkun.SelectedRows(0).Cells(Me.ColHeader.Index).Value = 1 Then
                If Me.GetLevelAnakMaksimum(Me.dgvAkun.SelectedRows(0).Index) = Me.LevelAkunMax Then
                    Me.BtnRight.Enabled = False
                    Exit Sub
                End If
            End If

            Me.BtnRight.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Set_Up_Arrow()
        Dim rowDGV As DataGridViewRow
        Dim Level, RowIndex As Integer
        Dim Ketemu As Boolean = False

        Try
            rowDGV = Me.dgvAkun.Rows(Me.dgvAkun.SelectedCells(0).RowIndex)
            Level = rowDGV.Cells(Me.ColLevel.Index).Value
            RowIndex = rowDGV.Index
            If RowIndex = 0 Then
                Me.btnUp.Enabled = False
            Else
                '1. Akun level 1 tidak boleh naik
                If rowDGV.Cells(Me.ColLevel.Index).Value = 1 Then
                    Me.btnUp.Enabled = False
                    Exit Sub
                End If

                '2. Jika bukan akun header
                'asal atasnya bukan akun level 1, maka boleh
                rowDGV = Me.dgvAkun.Rows(RowIndex - 1)
                If rowDGV.Cells(Me.ColLevel.Index).Value = 1 Then
                    Me.btnUp.Enabled = False
                    Exit Sub
                End If

                Me.btnUp.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAkun_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvAkun.KeyDown
        Try
            If e.KeyCode = Keys.Delete Then
                Me.btnDelete.PerformClick()
                'Me.dgvAkun.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAkun_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvAkun.SelectionChanged
        Try
            If Me.dgvAkun.SelectedRows.Count <> 0 And Me.dgvAkun.RowCount > 8 Then
                If Me.dgvAkun.SelectedRows.Count > 0 Then
                    Me.Set_Up_Arrow()
                    Me.Set_Down_Arrow()
                    Me.Set_Left_Arrow()
                    Me.Set_Right_Arrow()
                Else
                    Me.btnUp.Enabled = False
                    Me.btnDown.Enabled = False
                    Me.BtnRight.Enabled = False
                    Me.btnLeft.Enabled = False
                End If
            Else
                Me.btnUp.Enabled = False
                Me.btnDown.Enabled = False
                Me.BtnRight.Enabled = False
                Me.btnLeft.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUp.Click
        Try
            If Me.dgvAkun.SelectedRows(0).Cells(Me.ColHeader.Index).Value = 0 Then
                Me.Up_One_Row()
                '2. Level
                Me.UpdateLevel()
            Else
                Me.Up_One_Row_Header()
            End If

            'sesuaikan nilai
            '1. Urutan 
            Me.SetUrutan()
            'saldo header
            'Me.UpdateSaldoHeader()
            'cek kembali apakah boleh naik
            Me.Set_Up_Arrow()
            Me.Set_Down_Arrow()
            Me.Set_Left_Arrow()
            Me.Set_Right_Arrow()
            'update data
            Me.UpdateData2(Me.dgvAkun.SelectedRows(0).Cells(Me.ColTipePerkiraan.Index).Value)
            Me.ShowRow(Me.dgvAkun.SelectedRows(0), True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateSaldoHeader_Unused()
        Dim rowdgv As DataGridViewRow
        Me.dgvAkun.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        Me.dgvAkun.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

        Try
            For Each rowdgv In Me.dgvAkun.Rows
                If rowdgv.Cells(Me.ColHeader.Index).Value = 1 Then
                    If rowdgv.Cells(Me.ColSaldoNormal.Index).Value = 1 Then
                        rowdgv.Cells(Me.ColKredit.Index).Value = Me.ObjCD.DigitGrouping(Me.GetSaldoHeader(rowdgv.Index, _
                        rowdgv.Cells(Me.ColSaldoNormal.Index).Value, rowdgv.Cells(Me.ColLevel.Index).Value))
                    Else
                        rowdgv.Cells(Me.ColDebet.Index).Value = Me.ObjCD.DigitGrouping(Me.GetSaldoHeader(rowdgv.Index, _
                        rowdgv.Cells(Me.ColSaldoNormal.Index).Value, rowdgv.Cells(Me.ColLevel.Index).Value))
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvAkun.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvAkun.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        End Try
    End Sub

    Private Function GetSaldoHeader(ByVal vRowIndex As Double, ByVal vDebetKredit As Integer, ByVal vLevel As Integer) As Double
        Dim TargetIndex, CurrentIndex As Integer
        Dim rowdgv As DataGridViewRow
        Dim RetVal As Double = 0

        Try
            CurrentIndex = vRowIndex + 1
            While CurrentIndex <= Me.dgvAkun.Rows.Count - 1
                rowdgv = Me.dgvAkun.Rows(CurrentIndex)
                If rowdgv.Cells(Me.ColLevel.Index).Value <= vLevel Then
                    TargetIndex = rowdgv.Index - 1
                    Exit While
                ElseIf rowdgv.Index = Me.dgvAkun.Rows.Count - 1 Then
                    TargetIndex = rowdgv.Index
                    Exit While
                End If
                CurrentIndex += 1
            End While

            For i As Integer = vRowIndex + 1 To TargetIndex
                rowdgv = Me.dgvAkun.Rows(i)
                If rowdgv.Cells(Me.ColHeader.Index).Value = 1 And rowdgv.Cells(Me.ColLevel.Index).Value = vLevel + 1 Then
                    RetVal += Me.GetSaldoHeader(i, vDebetKredit, rowdgv.Cells(Me.ColLevel.Index).Value)
                ElseIf rowdgv.Cells(Me.ColLevel.Index).Value = vLevel + 1 Then
                    If vDebetKredit = 0 Then
                        RetVal += CDbl(rowdgv.Cells(Me.ColDebet.Index).Value)
                    Else
                        RetVal += CDbl(rowdgv.Cells(Me.ColKredit.Index).Value)
                    End If
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetLevelAnakMaksimum(ByVal vHeaderRowIndex As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Lvl As Integer = Me.dgvAkun.Rows(vHeaderRowIndex).Cells(Me.ColLevel.Index).Value
        Dim LvlMax As Integer = Lvl

        Try
            For i As Integer = vHeaderRowIndex + 1 To Me.dgvAkun.RowCount - 1
                RowDGV = Me.dgvAkun.Rows(i)
                If RowDGV.Cells(Me.ColLevel.Index).Value <= Lvl Then
                    Exit For
                End If
                If RowDGV.Cells(Me.ColLevel.Index).Value > LvlMax Then
                    LvlMax = RowDGV.Cells(Me.ColLevel.Index).Value
                End If
            Next

            Return LvlMax
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub UpdateLevel(ByVal vRowIndex As Integer)
        Dim rowDGV, rowDGV_Upper As DataGridViewRow
        Dim Level As Integer
        Dim Space As String = "     "
        Dim Selisih As Integer = 0


        Try
            rowDGV = Me.dgvAkun.Rows(vRowIndex)
            rowDGV_Upper = Me.dgvAkun.Rows(vRowIndex - 1)
            If rowDGV.Cells(Me.ColHeader.Index).Value = 0 Then
                If rowDGV_Upper.Cells(Me.ColHeader.Index).Value = 1 Then
                    'level = level upper + 1
                    rowDGV.Cells(Me.ColLevel.Index).Value = rowDGV_Upper.Cells(Me.ColLevel.Index).Value + 1
                Else
                    'level = level upper
                    rowDGV.Cells(Me.ColLevel.Index).Value = rowDGV_Upper.Cells(Me.ColLevel.Index).Value
                End If
            Else
                Selisih = Me.GetLevelAnakMaksimum(rowDGV.Index) - rowDGV.Cells(Me.ColLevel.Index).Value
                If rowDGV_Upper.Cells(Me.ColHeader.Index).Value = 1 Then
                    If rowDGV_Upper.Cells(Me.ColLevel.Index).Value + Selisih + 1 <= Me.LevelAkunMax Then
                        rowDGV.Cells(Me.ColLevel.Index).Value = rowDGV_Upper.Cells(Me.ColLevel.Index).Value + 1
                    End If
                Else
                    If rowDGV_Upper.Cells(Me.ColLevel.Index).Value + Selisih <= Me.LevelAkunMax Then
                        rowDGV.Cells(Me.ColLevel.Index).Value = rowDGV_Upper.Cells(Me.ColLevel.Index).Value
                    End If
                End If
            End If

            'update data tampilan
            rowDGV.Cells(Me.ColAccount.Index).Value = Trim(rowDGV.Cells(Me.ColAccount.Index).Value.ToString)
            If rowDGV.Cells(Me.ColLevel.Index).Value = 2 Then
                rowDGV.Cells(Me.ColAccount.Index).Value = Space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString
            ElseIf rowDGV.Cells(Me.ColLevel.Index).Value = 3 Then
                rowDGV.Cells(Me.ColAccount.Index).Value = Space + Space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString
            ElseIf rowDGV.Cells(Me.ColLevel.Index).Value = 4 Then
                rowDGV.Cells(Me.ColAccount.Index).Value = Space + Space + Space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString
            ElseIf rowDGV.Cells(Me.ColLevel.Index).Value = 5 Then
                rowDGV.Cells(Me.ColAccount.Index).Value = Space + Space + Space + Space + Space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString
            ElseIf rowDGV.Cells(Me.ColLevel.Index).Value = 6 Then
                rowDGV.Cells(Me.ColAccount.Index).Value = Space + Space + Space + Space + Space + Space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString
            ElseIf rowDGV.Cells(Me.ColLevel.Index).Value = 7 Then
                rowDGV.Cells(Me.ColAccount.Index).Value = Space + Space + Space + Space + Space + Space + Space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString
            ElseIf rowDGV.Cells(Me.ColLevel.Index).Value = 8 Then
                rowDGV.Cells(Me.ColAccount.Index).Value = Space + Space + Space + Space + Space + Space + Space + Space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString
            End If

            'update data primary lvl 1, lvl 2, lvl 3,lvl 4, lvl 5, lvl 6, lvl 7
            Level = rowDGV.Cells(Me.ColLevel.Index).Value
            If Level = 2 Then
                rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
            ElseIf Level = 3 Then
                rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
            ElseIf Level = 4 Then
                rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
            ElseIf Level = 5 Then
                rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = Me.GetPrimary(4, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
            ElseIf Level = 6 Then
                rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = Me.GetPrimary(4, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = Me.GetPrimary(5, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
            ElseIf Level = 7 Then
                rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = Me.GetPrimary(4, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = Me.GetPrimary(5, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = Me.GetPrimary(6, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
            ElseIf Level = 8 Then
                rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = Me.GetPrimary(4, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = Me.GetPrimary(5, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = Me.GetPrimary(6, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = Me.GetPrimary(7, rowDGV.Index - 1)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateLevel()
        Dim rowDGV, rowDGV_Upper As DataGridViewRow
        Dim Level As Integer
        Dim Space As String = "     "
        Dim Selisih As Integer = 0

        Try
            rowDGV = Me.dgvAkun.SelectedRows(0)
            rowDGV_Upper = Me.dgvAkun.Rows(Me.dgvAkun.SelectedRows(0).Index - 1)
            If rowDGV.Cells(Me.ColHeader.Index).Value = 0 Then
                If rowDGV_Upper.Cells(Me.ColHeader.Index).Value = 1 Then
                    'level = level upper + 1
                    rowDGV.Cells(Me.ColLevel.Index).Value = rowDGV_Upper.Cells(Me.ColLevel.Index).Value + 1
                Else
                    'level = level upper
                    rowDGV.Cells(Me.ColLevel.Index).Value = rowDGV_Upper.Cells(Me.ColLevel.Index).Value
                End If
            Else
                Selisih = Me.GetLevelAnakMaksimum(rowDGV.Index)
                If rowDGV_Upper.Cells(Me.ColHeader.Index).Value = 1 Then
                    If rowDGV_Upper.Cells(Me.ColLevel.Index).Value + Selisih + 1 <= Me.LevelAkunMax Then
                        rowDGV.Cells(Me.ColLevel.Index).Value = rowDGV_Upper.Cells(Me.ColLevel.Index).Value + 1
                    End If
                Else
                    If rowDGV_Upper.Cells(Me.ColLevel.Index).Value + Selisih <= Me.LevelAkunMax Then
                        rowDGV.Cells(Me.ColLevel.Index).Value = rowDGV_Upper.Cells(Me.ColLevel.Index).Value
                    End If
                End If
            End If

            'update data tampilan
            rowDGV.Cells(Me.ColAccount.Index).Value = Trim(rowDGV.Cells(Me.ColAccount.Index).Value.ToString)
            If rowDGV.Cells(Me.ColLevel.Index).Value = 2 Then
                rowDGV.Cells(Me.ColAccount.Index).Value = Space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString
            ElseIf rowDGV.Cells(Me.ColLevel.Index).Value = 3 Then
                rowDGV.Cells(Me.ColAccount.Index).Value = Space + Space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString
            ElseIf rowDGV.Cells(Me.ColLevel.Index).Value = 4 Then
                rowDGV.Cells(Me.ColAccount.Index).Value = Space + Space + Space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString
            ElseIf rowDGV.Cells(Me.ColLevel.Index).Value = 5 Then
                rowDGV.Cells(Me.ColAccount.Index).Value = Space + Space + Space + Space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString
            ElseIf rowDGV.Cells(Me.ColLevel.Index).Value = 6 Then
                rowDGV.Cells(Me.ColAccount.Index).Value = Space + Space + Space + Space + Space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString
            ElseIf rowDGV.Cells(Me.ColLevel.Index).Value = 7 Then
                rowDGV.Cells(Me.ColAccount.Index).Value = Space + Space + Space + Space + Space + Space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString
            ElseIf rowDGV.Cells(Me.ColLevel.Index).Value = 8 Then
                rowDGV.Cells(Me.ColAccount.Index).Value = Space + Space + Space + Space + Space + Space + Space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString
            End If

            'update data primary lvl 1, lvl 2, lvl 3
            Level = rowDGV.Cells(Me.ColLevel.Index).Value
            If Level = 2 Then
                rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
            ElseIf Level = 3 Then
                rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
            ElseIf Level = 4 Then
                rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
            ElseIf Level = 5 Then
                rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = Me.GetPrimary(4, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
            ElseIf Level = 6 Then
                rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = Me.GetPrimary(4, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = Me.GetPrimary(5, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
            ElseIf Level = 7 Then
                rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = Me.GetPrimary(4, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = Me.GetPrimary(5, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = Me.GetPrimary(6, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
            ElseIf Level = 8 Then
                rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = Me.GetPrimary(4, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = Me.GetPrimary(5, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = Me.GetPrimary(6, rowDGV.Index - 1)
                rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = Me.GetPrimary(7, rowDGV.Index - 1)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPrimary(ByVal vTargetLvl As Integer, ByVal vStartingRow As Double) As Double
        Dim rowDGV As DataGridViewRow
        Dim RetVal As Double = -1

        Try
            rowDGV = Me.dgvAkun.Rows(vStartingRow)
            While rowDGV.Index >= 0
                If rowDGV.Cells(Me.ColLevel.Index).Value = vTargetLvl Then
                    RetVal = rowDGV.Cells(Me.ColPrimaryAkun.Index).Value
                    Exit While
                End If
                If rowDGV.Index - 1 < 0 Then
                    Exit While
                Else
                    rowDGV = Me.dgvAkun.Rows(rowDGV.Index - 1)
                End If
            End While

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub Up_One_Row_Header()
        Dim TargetIndex, StartIndex As Integer
        Dim rowdgv As DataGridViewRow
        Dim Lvl, x As Integer
        Dim pTipeAkun As Integer = 0

        Try
            pTipeAkun = Me.dgvAkun.SelectedRows(0).Cells(Me.ColTipePerkiraan.Index).Value
            Lvl = Me.dgvAkun.SelectedRows(0).Cells(Me.ColLevel.Index).Value
            'target index adalah index dari akun pertama yang levelnya sama ke arah atas
            For i As Integer = Me.dgvAkun.SelectedRows(0).Index - 1 To 1 Step -1
                rowdgv = Me.dgvAkun.Rows(i)
                If rowdgv.Cells(Me.ColLevel.Index).Value <= Lvl Then
                    TargetIndex = rowdgv.Index
                    Exit For
                End If
            Next
            'TargetIndex = Me.dgvAkun.SelectedRows(0).Index - 1

            'start index adalah anak, termasuk anak dari anak, dengan index terakhir yang levelnya sama atau lebih kecil
            'dari level baris yang bersangkutan
            For i As Integer = Me.dgvAkun.SelectedRows(0).Index + 1 To Me.dgvAkun.RowCount - 1
                rowdgv = Me.dgvAkun.Rows(i)
                If rowdgv.Cells(Me.ColLevel.Index).Value <= Lvl Then
                    StartIndex = rowdgv.Index - 1
                    Exit For
                ElseIf rowdgv.Index = Me.dgvAkun.Rows.Count - 1 Then
                    'mulai dari baris yang paling bawah
                    StartIndex = rowdgv.Index
                End If
            Next

            'pindahkan baris ke target index, mulai dari start index
            x = StartIndex - Me.dgvAkun.SelectedRows(0).Index + 1
            For i As Integer = 1 To x
                Dim row As DataGridViewRow = Me.dgvAkun.Rows(StartIndex)
                If row.Index <> 0 Then
                    Dim indexbaru As Integer = TargetIndex
                    Me.dgvAkun.Rows.Remove(row)
                    Me.dgvAkun.Rows.InsertRange(indexbaru, row)
                End If
            Next

            Me.dgvAkun.ClearSelection()
            rowdgv = Me.dgvAkun.Rows(TargetIndex)
            rowdgv.Selected = True

            For i As Integer = 0 To x - 1
                Dim row As DataGridViewRow = Me.dgvAkun.Rows(TargetIndex + i)
                Me.UpdateLevel(row.Index)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Up_One_Row()
        Try
            If Me.dgvAkun.SelectedRows.Count > 0 Then
                Dim row As DataGridViewRow = Me.dgvAkun.SelectedRows(0)
                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.dgvAkun.Rows.Remove(row)
                    Me.dgvAkun.Rows.InsertRange(indexbaru, row)
                    Me.dgvAkun.Focus()
                    Me.dgvAkun.ClearSelection()
                    row.Selected = True
                    If row.Displayed = False Then
                        Me.ShowRow(row, True)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' tampilkan row yang diberikan pada parameter
    ''' </summary>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Private Sub ShowRow(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < Me.dgvAkun.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    Me.dgvAkun.FirstDisplayedScrollingRowIndex = Me.dgvAkun.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.dgvAkun.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While Me.dgvAkun.Rows(row.Index).Displayed = False
                    Me.dgvAkun.FirstDisplayedScrollingRowIndex = Me.dgvAkun.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If Me.dgvAkun.DisplayedRowCount(True) <> Me.dgvAkun.DisplayedRowCount(False) Then
                    Me.dgvAkun.FirstDisplayedScrollingRowIndex = Me.dgvAkun.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            If Me.dgvAkun.SelectedRows(0).Cells(Me.ColHeader.Index).Value = 0 Then
                Me.Down_One_Row()
                '2. Level
                Me.UpdateLevel()
            Else
                Me.Down_One_Row_Header()
            End If

            'sesuaikan nilai
            '1. Urutan 
            Me.SetUrutan()
            'saldo header
            'Me.UpdateSaldoHeader()
            'cek kembali apakah boleh naik
            Me.Set_Up_Arrow()
            Me.Set_Down_Arrow()
            Me.Set_Left_Arrow()
            Me.Set_Right_Arrow()
            'update data
            Me.UpdateData2(Me.dgvAkun.SelectedRows(0).Cells(Me.ColTipePerkiraan.Index).Value)
            Me.ShowRow(Me.dgvAkun.SelectedRows(0), False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Down_One_Row()
        Try
            If Me.dgvAkun.SelectedRows.Count > 0 Then
                Dim row As DataGridViewRow = Me.dgvAkun.SelectedRows(0)
                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.dgvAkun.Rows.Remove(row)
                    Me.dgvAkun.Rows.InsertRange(indexbaru, row)
                    Me.dgvAkun.Focus()
                    Me.dgvAkun.ClearSelection()
                    row.Selected = True
                    If row.Displayed = False Then
                        Me.ShowRow(row, True)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Down_One_Row_Header()
        Dim TargetIndex, StartIndex As Integer
        Dim rowdgv, rowdgv_Dummy As DataGridViewRow
        Dim Lvl, x As Integer

        Try
            Lvl = Me.dgvAkun.SelectedRows(0).Cells(Me.ColLevel.Index).Value
            'target index adalah index dari akun pertama yang levelnya sama ke arah bawah
            For i As Integer = Me.dgvAkun.SelectedRows(0).Index + 1 To Me.dgvAkun.Rows.Count - 1
                rowdgv = Me.dgvAkun.Rows(i)
                If rowdgv.Cells(Me.ColLevel.Index).Value <= Lvl Then
                    TargetIndex = rowdgv.Index

                    'jika dibatasi oleh header maka turunkan dengan melewati 1 set header itu
                    If rowdgv.Cells(Me.ColHeader.Index).Value = 1 Then
                        For j As Integer = rowdgv.Index + 1 To Me.dgvAkun.Rows.Count - 1
                            rowdgv_Dummy = Me.dgvAkun.Rows(j)
                            If rowdgv_Dummy.Cells(Me.ColLevel.Index).Value <= Lvl Then
                                TargetIndex = rowdgv_Dummy.Index - 1
                                Exit For
                            End If
                        Next
                    End If
                    Exit For
                End If
            Next

            'start index adalah anak, termasuk anak dari anak, dengan index terakhir yang levelnya sama atau lebih kecil
            'dari level baris yang bersangkutan
            For i As Integer = Me.dgvAkun.SelectedRows(0).Index + 1 To Me.dgvAkun.RowCount - 1
                rowdgv = Me.dgvAkun.Rows(i)
                If rowdgv.Cells(Me.ColLevel.Index).Value <= Lvl Then
                    StartIndex = rowdgv.Index - 1
                    Exit For
                End If
            Next

            'pindahkan baris ke target index, mulai dari start index
            x = StartIndex - Me.dgvAkun.SelectedRows(0).Index + 1
            StartIndex = Me.dgvAkun.SelectedRows(0).Index
            For i As Integer = 1 To x
                Dim row As DataGridViewRow = Me.dgvAkun.Rows(StartIndex)
                If row.Index <> 0 Then
                    Dim indexbaru As Integer = TargetIndex
                    Me.dgvAkun.Rows.Remove(row)
                    Me.dgvAkun.Rows.InsertRange(indexbaru, row)
                End If
            Next

            Me.dgvAkun.ClearSelection()
            rowdgv = Me.dgvAkun.Rows(TargetIndex - x + 1)
            rowdgv.Selected = True

            'tidak perlu melakukan update level karena memang tidak ada yang berubah

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnRight_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnRight.Click
        Try
            If Me.dgvAkun.SelectedRows(0).Cells(Me.ColHeader.Index).Value = 0 Then
                Me.Right_One_Row()
            Else
                Me.Right_One_Row_Header()
            End If

            'saldo header
            'Me.UpdateSaldoHeader()
            'cek kembali apakah boleh geser kanan
            Me.Set_Up_Arrow()
            Me.Set_Down_Arrow()
            Me.Set_Left_Arrow()
            Me.Set_Right_Arrow()
            'update data
            Me.UpdateData2(Me.dgvAkun.SelectedRows(0).Cells(Me.ColTipePerkiraan.Index).Value)
            Me.ShowRow(Me.dgvAkun.SelectedRows(0), False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub Right_One_Row()
        Dim space As String = "     "
        Dim Level As Integer = 1

        Try
            If Me.dgvAkun.SelectedRows.Count > 0 Then
                Dim rowDGV As DataGridViewRow = Me.dgvAkun.SelectedRows(0)
                rowDGV.Cells(Me.ColLevel.Index).Value = CDbl(rowDGV.Cells(Me.ColLevel.Index).Value) + 1
                rowDGV.Cells(Me.ColAccount.Index).Value = space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString

                'update data primary lvl 1, lvl 2, lvl 3
                Level = rowDGV.Cells(Me.ColLevel.Index).Value
                If Level = 2 Then
                    rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
                ElseIf Level = 3 Then
                    rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
                ElseIf Level = 4 Then
                    rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
                ElseIf Level = 5 Then
                    rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = Me.GetPrimary(4, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
                ElseIf Level = 6 Then
                    rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = Me.GetPrimary(4, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = Me.GetPrimary(5, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
                ElseIf Level = 7 Then
                    rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = Me.GetPrimary(4, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = Me.GetPrimary(5, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = Me.GetPrimary(6, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
                ElseIf Level = 8 Then
                    rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = Me.GetPrimary(4, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = Me.GetPrimary(5, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = Me.GetPrimary(6, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = Me.GetPrimary(7, rowDGV.Index - 1)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Right_One_Row_Header()
        Dim TargetIndex, StartIndex, Level As Integer
        Dim rowdgv As DataGridViewRow
        Dim Lvl As Integer
        Dim Space As String = "     "

        Try
            Lvl = Me.dgvAkun.SelectedRows(0).Cells(Me.ColLevel.Index).Value
            'target index adalah index dari akun pertama yang levelnya sama ke arah bawah
            For i As Integer = Me.dgvAkun.SelectedRows(0).Index + 1 To Me.dgvAkun.Rows.Count - 1
                rowdgv = Me.dgvAkun.Rows(i)
                If rowdgv.Cells(Me.ColLevel.Index).Value <= Lvl Then
                    TargetIndex = rowdgv.Index
                    Exit For
                End If
            Next

            'start index adalah index dirinya sendiri
            StartIndex = Me.dgvAkun.SelectedRows(0).Index

            'tambahkan 1 level, mulai dari start index
            For i As Integer = StartIndex To TargetIndex - 1
                rowdgv = Me.dgvAkun.Rows(i)
                rowdgv.Cells(Me.ColLevel.Index).Value = CDbl(rowdgv.Cells(Me.ColLevel.Index).Value) + 1
                rowdgv.Cells(Me.ColAccount.Index).Value = Space + rowdgv.Cells(Me.ColAccount.Index).Value.ToString

                'update data primary lvl 1, lvl 2, lvl 3
                Level = rowdgv.Cells(Me.ColLevel.Index).Value
                If Level = 2 Then
                    rowdgv.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl2.Index).Value = -1
                    rowdgv.Cells(Me.ColPrimaryLvl3.Index).Value = -1
                    rowdgv.Cells(Me.ColPrimaryLvl4.Index).Value = -1
                    rowdgv.Cells(Me.ColPrimaryLvl5.Index).Value = -1
                    rowdgv.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                    rowdgv.Cells(Me.ColPrimaryLvl7.Index).Value = -1
                ElseIf Level = 3 Then
                    rowdgv.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl3.Index).Value = -1
                    rowdgv.Cells(Me.ColPrimaryLvl4.Index).Value = -1
                    rowdgv.Cells(Me.ColPrimaryLvl5.Index).Value = -1
                    rowdgv.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                    rowdgv.Cells(Me.ColPrimaryLvl7.Index).Value = -1
                ElseIf Level = 4 Then
                    rowdgv.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl4.Index).Value = -1
                    rowdgv.Cells(Me.ColPrimaryLvl5.Index).Value = -1
                    rowdgv.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                    rowdgv.Cells(Me.ColPrimaryLvl7.Index).Value = -1
                ElseIf Level = 5 Then
                    rowdgv.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl4.Index).Value = Me.GetPrimary(4, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl5.Index).Value = -1
                    rowdgv.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                    rowdgv.Cells(Me.ColPrimaryLvl7.Index).Value = -1
                ElseIf Level = 6 Then
                    rowdgv.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl4.Index).Value = Me.GetPrimary(4, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl5.Index).Value = Me.GetPrimary(5, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                    rowdgv.Cells(Me.ColPrimaryLvl7.Index).Value = -1
                ElseIf Level = 7 Then
                    rowdgv.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl4.Index).Value = Me.GetPrimary(4, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl5.Index).Value = Me.GetPrimary(5, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl6.Index).Value = Me.GetPrimary(6, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl7.Index).Value = -1
                ElseIf Level = 8 Then
                    rowdgv.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl4.Index).Value = Me.GetPrimary(4, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl5.Index).Value = Me.GetPrimary(5, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl6.Index).Value = Me.GetPrimary(6, rowdgv.Index - 1)
                    rowdgv.Cells(Me.ColPrimaryLvl7.Index).Value = Me.GetPrimary(7, rowdgv.Index - 1)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' update data tampilan, urutan, primary lvl 1, primary lvl 2, primary lvl 3, level
    ''' digunakan uleh button up dan down, yang diupdate adalah 1 set tipe akun
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub UpdateData2(ByVal vTipeAkun As Integer)
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akun2TableAdapter
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvAkun.Rows
                If rowDGV.Cells(Me.ColTipePerkiraan.Index).Value = vTipeAkun Then
                    With rowDGV
                        DTA.UpdateQuery1(.Cells(Me.ColLevel.Index).Value, .Cells(Me.ColAccount.Index).Value, _
                        .Cells(Me.ColUrutan.Index).Value, .Cells(Me.ColPrimaryLvl1.Index).Value, _
                        .Cells(Me.ColPrimaryLvl2.Index).Value, .Cells(Me.ColPrimaryLvl3.Index).Value, _
                        .Cells(Me.ColPrimaryLvl4.Index).Value, .Cells(Me.ColPrimaryLvl5.Index).Value, _
                        .Cells(Me.ColPrimaryLvl6.Index).Value, .Cells(Me.ColPrimaryLvl7.Index).Value, _
                        .Cells(Me.ColPrimaryAkun.Index).Value)
                    End With
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' melakukan update data urutan ke dalam basis data
    ''' </summary>
    ''' <param name="vTipeAkun"></param>
    ''' <remarks></remarks>
    Private Sub UpdateUrutan(ByVal vTipeAkun As Integer)
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akun2TableAdapter
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvAkun.Rows
                If rowDGV.Cells(Me.ColTipePerkiraan.Index).Value = vTipeAkun Then
                    With rowDGV
                        DTA.UpdateQuery1(.Cells(Me.ColLevel.Index).Value, .Cells(Me.ColAccount.Index).Value, _
                        .Cells(Me.ColUrutan.Index).Value, .Cells(Me.ColPrimaryLvl1.Index).Value, _
                        .Cells(Me.ColPrimaryLvl2.Index).Value, .Cells(Me.ColPrimaryLvl3.Index).Value, _
                        .Cells(Me.ColPrimaryLvl4.Index).Value, .Cells(Me.ColPrimaryLvl5.Index).Value, _
                        .Cells(Me.ColPrimaryLvl6.Index).Value, .Cells(Me.ColPrimaryLvl7.Index).Value, _
                        .Cells(Me.ColPrimaryAkun.Index).Value)
                    End With
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub btnLeft_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLeft.Click
        Try
            If Me.dgvAkun.SelectedRows(0).Cells(Me.ColHeader.Index).Value = 0 Then
                Me.Left_One_Row()
            Else
                Me.Left_One_Row_Header()
            End If

            'saldo header
            'Me.UpdateSaldoHeader()
            'cek kembali apakah boleh geser kanan
            Me.Set_Up_Arrow()
            Me.Set_Down_Arrow()
            Me.Set_Left_Arrow()
            Me.Set_Right_Arrow()
            'update data
            Me.UpdateData2(Me.dgvAkun.SelectedRows(0).Cells(Me.ColTipePerkiraan.Index).Value)
            Me.ShowRow(Me.dgvAkun.SelectedRows(0), False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Left_One_Row_Header()
        Dim space As String = "     "
        Dim Level As Integer = 1
        Dim Lvl As Integer = 1
        Dim TargetIndex As Integer = 0
        Dim StartIndex, StartIndex2, X As Integer
        Dim rowDGV_Dummy, rowDGV As DataGridViewRow

        Try
            If Me.dgvAkun.SelectedRows.Count > 0 Then
                'tentukan target index dan start index untuk update date level dan tampilan
                Lvl = Me.dgvAkun.SelectedRows(0).Cells(Me.ColLevel.Index).Value
                'target index adalah index dari akun pertama yang levelnya sama ke arah bawah
                For i As Integer = Me.dgvAkun.SelectedRows(0).Index + 1 To Me.dgvAkun.Rows.Count - 1
                    rowDGV = Me.dgvAkun.Rows(i)
                    If rowDGV.Cells(Me.ColLevel.Index).Value <= Lvl Then
                        TargetIndex = rowDGV.Index - 1
                        Exit For
                    End If
                Next
                StartIndex = Me.dgvAkun.SelectedRows(0).Index

                'menentukan posisi start baris yang dipindahkan 
                For i As Integer = Me.dgvAkun.SelectedRows(0).Index + 1 To Me.dgvAkun.RowCount - 1
                    rowDGV = Me.dgvAkun.Rows(i)
                    If rowDGV.Cells(Me.ColLevel.Index).Value <= Lvl Then
                        StartIndex2 = rowDGV.Index - 1
                        Exit For
                    End If
                Next

                For i As Integer = StartIndex To TargetIndex
                    rowDGV = Me.dgvAkun.Rows(i)
                    rowDGV.Cells(Me.ColLevel.Index).Value = CDbl(rowDGV.Cells(Me.ColLevel.Index).Value) - 1
                    'update data primary lvl 1, lvl 2, lvl 3 dan tampilan
                    Level = rowDGV.Cells(Me.ColLevel.Index).Value
                    If Level = 2 Then
                        rowDGV.Cells(Me.ColAccount.Index).Value = space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString.Trim
                        rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = -1
                        rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = -1
                        rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = -1
                        rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = -1
                        rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                        rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
                    ElseIf Level = 3 Then
                        rowDGV.Cells(Me.ColAccount.Index).Value = space + space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString.Trim
                        rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = -1
                        rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = -1
                        rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = -1
                        rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                        rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
                    ElseIf Level = 4 Then
                        rowDGV.Cells(Me.ColAccount.Index).Value = space + space + space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString.Trim
                        rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = -1
                        rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = -1
                        rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                        rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
                    ElseIf Level = 5 Then
                        rowDGV.Cells(Me.ColAccount.Index).Value = space + space + space + space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString.Trim
                        rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = Me.GetPrimary(4, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = -1
                        rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                        rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
                    ElseIf Level = 6 Then
                        rowDGV.Cells(Me.ColAccount.Index).Value = space + space + space + space + space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString.Trim
                        rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = Me.GetPrimary(4, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = Me.GetPrimary(5, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                        rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
                    ElseIf Level = 7 Then
                        rowDGV.Cells(Me.ColAccount.Index).Value = space + space + space + space + space + space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString.Trim
                        rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = Me.GetPrimary(4, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = Me.GetPrimary(5, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = Me.GetPrimary(6, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
                    ElseIf Level = 8 Then
                        rowDGV.Cells(Me.ColAccount.Index).Value = space + space + space + space + space + space + space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString.Trim
                        rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = Me.GetPrimary(4, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = Me.GetPrimary(5, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = Me.GetPrimary(6, rowDGV.Index - 1)
                        rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = Me.GetPrimary(7, rowDGV.Index - 1)
                    End If
                Next

                'pindahkan data secara keseluruhan jika memang dibutuhkan
                'dipindahkan mulai dari baris yang paling bawah
                'StartIndex = TargetIndex
                'pindahkan baris sesuai dengan posisi level yang baru
                'target index adalah, akun pertama dengan level akun yang sama dengan level baru dari akun ybs ke arah bawah
                'dalam hal ketemu akun level 1, maka target index = (index akun lvl 1) - 1
                'dalam hak ketemu sampe paling bawah maka target akunnya adalah yang paling bawah
                rowDGV = Me.dgvAkun.SelectedRows(0)
                Level = rowDGV.Cells(Me.ColLevel.Index).Value
                Lvl = rowDGV.Cells(Me.ColLevel.Index).Value
                For i As Integer = rowDGV.Index + 1 To Me.dgvAkun.Rows.Count - 1
                    rowDGV_Dummy = Me.dgvAkun.Rows(i)
                    If rowDGV_Dummy.Cells(Me.ColLevel.Index).Value = 1 Then
                        TargetIndex = rowDGV_Dummy.Index - 1
                        Exit For
                    ElseIf rowDGV_Dummy.Index = Me.dgvAkun.Rows.Count - 1 Then
                        TargetIndex = Me.dgvAkun.Rows.Count - 1
                        Exit For
                    Else
                        If rowDGV_Dummy.Cells(Me.ColLevel.Index).Value <= Level Then
                            TargetIndex = rowDGV_Dummy.Index - 1
                            Exit For
                        End If
                    End If
                Next

                If TargetIndex <> Me.dgvAkun.SelectedRows(0).Index Then
                    X = StartIndex2 - Me.dgvAkun.SelectedRows(0).Index + 1
                    StartIndex2 = Me.dgvAkun.SelectedRows(0).Index
                    For i As Integer = 1 To X
                        Dim row As DataGridViewRow = Me.dgvAkun.Rows(StartIndex2)
                        If row.Index <> 0 Then
                            Dim indexbaru As Integer = TargetIndex
                            Me.dgvAkun.Rows.Remove(row)
                            Me.dgvAkun.Rows.InsertRange(indexbaru, row)
                        End If
                    Next

                    Me.dgvAkun.ClearSelection()
                    rowDGV = Me.dgvAkun.Rows(TargetIndex - X + 1)
                    rowDGV.Selected = True

                    'update data urutan
                    Me.SetUrutan()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Left_One_Row()
        Dim space As String = "     "
        Dim Level As Integer = 1
        Dim TargetIndex As Integer = 0
        Dim rowDGV_Dummy As DataGridViewRow

        Try
            If Me.dgvAkun.SelectedRows.Count > 0 Then
                Dim rowDGV As DataGridViewRow = Me.dgvAkun.SelectedRows(0)
                rowDGV.Cells(Me.ColLevel.Index).Value = CDbl(rowDGV.Cells(Me.ColLevel.Index).Value) - 1

                'update data primary lvl 1, lvl 2, lvl 3, lvl4, lvl5, lvl6, vl 7 dan tampilan
                Level = rowDGV.Cells(Me.ColLevel.Index).Value
                'pindahkan baris sesuai dengan posisi level yang baru
                'target index adalah, akun pertama dengan level akun yang sama dengan level baru dari akun ybs ke arah bawah
                'dalam hal ketemu akun level 1, maka target index = (index akun lvl 1) - 1
                'dalam hak ketemu sampe paling bawah maka target akunnya adalah yang paling bawah

                If rowDGV.Index <> Me.dgvAkun.Rows.Count - 1 Then
                    For i As Integer = rowDGV.Index + 1 To Me.dgvAkun.Rows.Count - 1
                        rowDGV_Dummy = Me.dgvAkun.Rows(i)
                        If rowDGV_Dummy.Cells(Me.ColLevel.Index).Value = 1 Then
                            TargetIndex = rowDGV_Dummy.Index - 1
                            Exit For
                        ElseIf rowDGV_Dummy.Index = Me.dgvAkun.Rows.Count - 1 Then
                            TargetIndex = Me.dgvAkun.Rows.Count - 1
                            Exit For
                        Else
                            If rowDGV_Dummy.Cells(Me.ColLevel.Index).Value <= Level Then
                                TargetIndex = rowDGV_Dummy.Index - 1
                                Exit For
                            End If
                        End If
                    Next
                Else
                    TargetIndex = rowDGV.Index
                End If
                If TargetIndex <> rowDGV.Index Then
                    Me.dgvAkun.Rows.Remove(rowDGV)
                    Me.dgvAkun.Rows.InsertRange(TargetIndex, rowDGV)
                    Me.dgvAkun.Focus()
                    Me.dgvAkun.ClearSelection()
                    rowDGV.Selected = True
                    If rowDGV.Displayed = False Then
                        Me.ShowRow(rowDGV, True)
                    End If
                    'update data urutan
                    Me.SetUrutan()
                End If

                If Level = 2 Then
                    rowDGV.Cells(Me.ColAccount.Index).Value = space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString.Trim
                    rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
                ElseIf Level = 3 Then
                    rowDGV.Cells(Me.ColAccount.Index).Value = space + space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString.Trim
                    rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
                ElseIf Level = 4 Then
                    rowDGV.Cells(Me.ColAccount.Index).Value = space + space + space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString.Trim
                    rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
                ElseIf Level = 5 Then
                    rowDGV.Cells(Me.ColAccount.Index).Value = space + space + space + space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString.Trim
                    rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = Me.GetPrimary(4, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
                ElseIf Level = 6 Then
                    rowDGV.Cells(Me.ColAccount.Index).Value = space + space + space + space + space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString.Trim
                    rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = Me.GetPrimary(4, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = Me.GetPrimary(5, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = -1
                    rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
                ElseIf Level = 7 Then
                    rowDGV.Cells(Me.ColAccount.Index).Value = space + space + space + space + space + space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString.Trim
                    rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = Me.GetPrimary(4, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = Me.GetPrimary(5, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = Me.GetPrimary(6, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = -1
                ElseIf Level = 8 Then
                    rowDGV.Cells(Me.ColAccount.Index).Value = space + space + space + space + space + space + space + rowDGV.Cells(Me.ColAccount.Index).Value.ToString.Trim
                    rowDGV.Cells(Me.ColPrimaryLvl1.Index).Value = Me.GetPrimary(1, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl2.Index).Value = Me.GetPrimary(2, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl3.Index).Value = Me.GetPrimary(3, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl4.Index).Value = Me.GetPrimary(4, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl5.Index).Value = Me.GetPrimary(5, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl6.Index).Value = Me.GetPrimary(6, rowDGV.Index - 1)
                    rowDGV.Cells(Me.ColPrimaryLvl7.Index).Value = Me.GetPrimary(7, rowDGV.Index - 1)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnBudget_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBudget.Click
        Dim FrmParent As Windows.Forms.Form = Me.MdiParent
        Try
            Dim ij As FrmMasterBudget
            Dim j As Windows.Forms.Form
            For Each j In FrmParent.MdiChildren
                ij = TryCast(j, FrmMasterBudget)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next
            ' Increment the caption counter.
            Dim i As New FrmMasterBudget
            i.UserCode = Me.UserCode
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = FrmParent
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
