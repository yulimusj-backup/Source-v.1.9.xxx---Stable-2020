Option Explicit On
Option Strict Off
Imports Core_Module

Public Class FrmMast_Discount_Main
    Public UserCode As String = "USER"
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Public Bahasa As Core_Module.ModGlobalVariable.EnumBahasa = Core_Module.EnumBahasa.Inggris

    Dim baru As Boolean
    Dim buffer As String
    Dim ObjCD As New ClsCekData
    Dim Obj As New ClsMast_Diskon
    Dim ObjItem As New ClsMast_Diskon_Item

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    ''Up Down
    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvMain.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvMain)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsAdded(ByVal sender As Object, ByVal e As _
    System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvMain.RowsAdded
        Try
            Me.dgvMain.TabStop = True
            'If Me.TxtRemarks.Enabled = True Then
            Me.btnUp.Enabled = True
            Me.btnDown.Enabled = True
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsRemoved(ByVal sender As Object, ByVal e As _
    System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvMain.RowsRemoved
        Try
            If Me.dgvMain.Rows.Count = 0 Then
                Me.dgvMain.TabStop = False
                'If Me.TxtRemarks.Enabled = True Then
                Me.btnUp.Enabled = False
                Me.btnDown.Enabled = False
                'End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e _
    As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvMain.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvMain)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles dgvMain.SelectionChanged
        Try
            'If Me.TxtRemarks.Enabled = True Then
            If Me.dgvMain.SelectedRows.Count > 0 Then
                Me.btnUp.Enabled = True
                Me.btnDown.Enabled = True
            Else
                Me.btnUp.Enabled = False
                Me.btnDown.Enabled = False
            End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' tampilkan row yang diberikan pada parameter
    ''' </summary>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Private Sub ShowRowProcess(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < Me.dgvMain.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    Me.dgvMain.FirstDisplayedScrollingRowIndex = Me.dgvMain.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.dgvMain.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While Me.dgvMain.Rows(row.Index).Displayed = False
                    Me.dgvMain.FirstDisplayedScrollingRowIndex = Me.dgvMain.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If Me.dgvMain.DisplayedRowCount(True) <> Me.dgvMain.DisplayedRowCount(False) Then
                    Me.dgvMain.FirstDisplayedScrollingRowIndex = Me.dgvMain.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUp.Click
        Try
            If Me.dgvMain.SelectedRows.Count > 0 Then
                Me.btnUp.Enabled = False
                Me.btnDown.Enabled = False

                Dim row As DataGridViewRow = Me.dgvMain.SelectedRows(0)
                If row.Index <> 0 Then
                    Dim indexlama = row.Index
                    Dim indexbaru = row.Index - 1
                    Me.dgvMain.Rows.Remove(row)
                    Me.dgvMain.Rows.InsertRange(indexbaru, row)
                    Me.dgvMain.Focus()
                    Me.dgvMain.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvMain)
                    If row.Displayed = False Then
                        Me.ShowRowProcess(row, True)
                    End If
                    'update urutan diskon
                    Obj.UpdateUrutanDiskonSatuBandingSatu(indexlama, indexbaru, Me.dgvMain, Me)
                End If

                Me.btnUp.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            If Me.dgvMain.SelectedRows.Count > 0 Then
                Me.btnUp.Enabled = False
                Me.btnDown.Enabled = False

                Dim row As DataGridViewRow = Me.dgvMain.SelectedRows(0)
                If row.Index <> Me.dgvMain.Rows.Count - 1 Then
                    Dim indexlama = row.Index
                    Dim indexbaru = row.Index + 1
                    Me.dgvMain.Rows.Remove(row)
                    Me.dgvMain.Rows.InsertRange(indexbaru, row)
                    Me.dgvMain.Focus()
                    Me.dgvMain.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvMain)
                    Me.ShowRowProcess(row, False)
                    'update urutan diskon
                    Obj.UpdateUrutanDiskonSatuBandingSatu(indexlama, indexbaru, Me.dgvMain, Me)
                End If

                Me.btnUp.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMast_Discount_Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.dgvMain.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Obj.LoadDataMain(Me.dgvMain, Me)
            Me.Text = "Master Discount"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        Dim FrmParent As Windows.Forms.Form = Me.MdiParent
        Try
            'dapet in lagi disorot turunkan kode
            ''RowDgv = Me.dgvMain.SelectedRows(0)
            ''Coding Form show
            Dim ij As FrmMast_Discount
            Dim j As Windows.Forms.Form
            For Each j In FrmParent.MdiChildren
                ij = TryCast(j, FrmMast_Discount)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New FrmMast_Discount
            i.UserCode = Me.UserCode
            ' Set the caption to be unique.
            i.BackColorAll = Color.LightGoldenrodYellow
            ''sesuaikan kondisi awal
            With i
                .Baru = True
                .KosongkanField()
                .AktifkanDetail()
                i.UserCode = Me.UserCode
                i.BackColorAll = Me.BackColorAll
                i.Bahasa = Me.Bahasa
                .tslCreatedBy.Text = "Created By: " + Me.UserCode.ToString
                .tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
                .tslModifiedBy.Text = "Modified By: " + Me.UserCode.ToString
                .tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
                i.dgvItemProd.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
                i.dgvBranch.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
                i.TxtUrutanDiskon.Text = Obj.GetMAXUrutan + 1
                i.TxtCode.Focus()
                .TxtCode.BackColor = Color.LightGoldenrodYellow
            End With
            i.ShowDialog()
            If i.Ok = True Then
                'Load Data
                Obj.LoadDataMain(Me.dgvMain, Me)
                i.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Dim ij As New FrmMast_Discount
        Dim rowDGV As DataGridViewRow
        Dim FrmParent As Windows.Forms.Form = Me.MdiParent
        Try
            'dapet in lagi disorot turunkan kode
            If Me.dgvMain.SelectedRows.Count > 0 Then
                'Sesuaikan Data
                ij.UserCode = Me.UserCode
                ij.BackColorAll = Me.BackColorAll
                ij.Baru = False
                ij.Bahasa = Me.Bahasa
                ' caridata
                If Me.Obj.ApakahBolehEdit(Me.dgvMain.SelectedRows(0) _
                .Cells(Me.ColPrimary.Index).Value, Me.UserCode, True, True, Me.Bahasa) = True Then
                    rowDGV = Me.dgvMain.SelectedRows(0)
                    ij.Obj.CariDetail(Me.dgvMain.SelectedRows(0).Cells(Me.ColPrimary.Index).Value, Me.ObjCD, ij.dgvItemProd, ij.dgvBranch, ij)
                    'salin data
                    ij.SalinKeField()
                    'atur object
                    ij.NonAktifkanDetail()
                    'aktifkan detail
                    ij.AktifkanDetail()
                    ij.TxtCode.Focus()
                    ij.TxtCode.BackColor = Color.LightGoldenrodYellow
                    ''Coding Form show
                    ij.ShowDialog()
                End If
                'load data
                Obj.LoadDataMain(Me.dgvMain, Me)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Dim rowDGV As DataGridViewRow
        Dim ij As New FrmMast_Discount
        Try
            If Me.dgvMain.SelectedRows.Count > 0 Then
                If Me.Obj.ApakahBolehHapus(Me.dgvMain.SelectedRows(0).Cells(Me.ColPrimary.Index).Value, Me.UserCode, _
                 True, True, Me.Bahasa, ij) = False Then
                    Me.btnDelete.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        Obj.UpdateUrutanDiskonDeleting(Me.dgvMain.SelectedRows(0).Cells(Me.ColNo.Index).Value, _
                        Me.dgvMain.SelectedRows(0).Cells(Me.ColPrimary.Index).Value, Me.dgvMain, Me)
                        If Me.Obj.Hapus(Me.dgvMain.SelectedRows(0).Cells(Me.ColPrimary.Index).Value, Me.UserCode, True, True, Me.Bahasa, ij) = True Then
                            rowDGV = Me.dgvMain.SelectedRows(0)
                            'load data
                            Obj.LoadDataMain(Me.dgvMain, Me)
                            Me.btnAdd.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvMain_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvMain.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColMultiDiscount.Index Then
                Me.buffer = Me.dgvMain.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvMain_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvMain.CellEndEdit
        Dim Row As DataGridViewRow = Me.dgvMain.Rows(e.RowIndex)
        Dim Multidiskon As Double = -1
        Try
            Multidiskon = Obj.GetMultidiskon(Me.dgvMain.Item(Me.ColPrimary.Index, e.RowIndex).Value)
            If Multidiskon = 1 Then
                Obj.UpdateMultidiskon(Row.Cells(Me.ColPrimary.Index).Value, 0)
            Else
                Obj.UpdateMultidiskon(Row.Cells(Me.ColPrimary.Index).Value, 1)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class