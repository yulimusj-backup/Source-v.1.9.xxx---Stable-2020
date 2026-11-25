Option Explicit On
Option Strict Off

Public Class FrmTrans_ProductionSchedule_NewSch
    Dim ObjMOItem As New ClsTrans_ManufacturingOrder_Item
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjSetting As New Core_Module.ClsTransaksiSettingUser
    Dim ObjProd As New Core_Module.clsMasterProduct
    Dim ObjMO As New ClsTrans_ManufacturingOrder
    Public UserCode As String = "USER"
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Private Sub FrmTrans_ProductionSchedule_NewSch_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.ObjMOItem.IsiDGVJadwalBaru(Me.dgvItem, Me)
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.SetKolomVisible()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetKolomVisible()
        Dim ShowJO, ShowMO, ShowCodeNo, ShowDesc, ShowQtyTarget As String

        Try
            ShowJO = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch_ShowJO, Me.UserCode)
            ShowMO = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch_ShowMO, Me.UserCode)
            ShowCodeNo = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch_ShowCodeNo, Me.UserCode)
            ShowDesc = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch_ShowDesc, Me.UserCode)
            ShowQtyTarget = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch_ShowQtyTarget, Me.UserCode)
            If ShowJO = "" And ShowMO = "" And ShowCodeNo = "" And ShowDesc = "" And ShowQtyTarget = "" Then
                'SET KE DEFAULT VALUE
                ShowJO = "0"
                ShowMO = "1"
                ShowCodeNo = "1"
                ShowDesc = "1"
                ShowQtyTarget = "1"
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch_ShowJO, ShowJO, Me.UserCode)
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch_ShowMO, ShowMO, Me.UserCode)
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch_ShowCodeNo, ShowCodeNo, Me.UserCode)
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch_ShowDesc, ShowDesc, Me.UserCode)
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch_ShowQtyTarget, ShowQtyTarget, Me.UserCode)
            End If
            If ShowJO = "1" Then
                Me.ColJO.Visible = True
            Else
                Me.ColJO.Visible = False
            End If
            If ShowMO = "1" Then
                Me.ColMO.Visible = True
            Else
                Me.ColMO.Visible = False
            End If
            If ShowCodeNo = "1" Then
                Me.ColCodeNo.Visible = True
            Else
                Me.ColCodeNo.Visible = False
            End If
            If ShowDesc = "1" Then
                Me.ColDesc.Visible = True
            Else
                Me.ColDesc.Visible = False
            End If
            If ShowQtyTarget = "1" Then
                Me.ColQtyTarget.Visible = True
            Else
                Me.ColQtyTarget.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

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

    Private Sub btnUP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUp.Click
        Try
            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.btnUp.Enabled = False
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
                Me.btnUp.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.btnUp.Enabled = False
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
                Me.btnUp.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColSchedule.Index Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbCheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCheckAll.LinkClicked
        Dim rowDGV As DataGridViewRow

        Try
            If MsgBox("Check all data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                For Each rowDGV In Me.dgvItem.Rows
                    rowDGV.Cells(Me.ColSchedule.Index).Value = Boolean.TrueString
                Next
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUncheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUncheckAll.LinkClicked
        Dim rowDGV As DataGridViewRow

        Try
            If MsgBox("Uncheck all data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                For Each rowDGV In Me.dgvItem.Rows
                    rowDGV.Cells(Me.ColSchedule.Index).Value = Boolean.FalseString
                Next
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbColumnsSettings_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbColumnsSettings.LinkClicked
        Dim ij As New FrmTrans_ProductionSchedule_NewSch_Setting

        Try
            ij.UserCode = Me.UserCode

            If ij.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Me.SetKolomVisible()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If MsgBox("Cancel scheduling?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim CurrentJO, TotalJO As Double

        Try
            'pastikan untuk setiap item yang diklik, setiap productnya memiliki work unit
            If Me.CekWorkUnit = False Then
                Exit Sub
            End If

            CurrentJO = 0
            TotalJO = 0
            For Each rowDGV As DataGridViewRow In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColSchedule.Index).Value = Boolean.TrueString Then
                    TotalJO += 1
                End If
            Next

            If TotalJO > 0 Then
                If MsgBox("Start alocating schedule?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                    For Each rowDGV As DataGridViewRow In Me.dgvItem.Rows
                        If rowDGV.Cells(Me.ColSchedule.Index).Value = Boolean.TrueString Then
                            CurrentJO += 1
                            Dim ij As New FrmTrans_ProductionSchedule_NewSch2
                            ij.UserCode = Me.UserCode
                            ij.BackColorAll = Me.BackColorAll
                            ij.IdentifierJO = rowDGV.Cells(Me.ColIdentifierJO.Index).Value
                            ij.TotalJO = TotalJO
                            ij.CurrentJO = CurrentJO
                            ij.ShowDialog()
                            If ij.Stopped = True Then
                                ij.Dispose()
                                Exit For
                            End If
                            ij.Dispose()
                        End If
                    Next

                    'load ulang data
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    Me.ObjMOItem.IsiDGVJadwalBaru(Me.dgvItem, Me)
                    Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
                    Me.SetKolomVisible()

                    If Me.dgvItem.RowCount = 0 Then
                        Me.Close()
                    End If
                End If
            Else
                MsgBox("Please specify at least 1 job order to schedule.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function CekWorkUnit() As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTADummy As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_wo3TableAdapter
        Dim TBLDummy As New DseTrans_ManufacturingOrder.tbltrans_mo_item_wo3DataTable
        Dim TBLDummy1 As New DseTrans_ManufacturingOrder.tbltrans_mo_item_wo3DataTable
        Dim rowDummy, rowDummy1 As DseTrans_ManufacturingOrder.tbltrans_mo_item_wo3Row
        Dim Frm1 As New FrmTrans_ProductionSchedule_WUExc

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColSchedule.Index).Value = Boolean.TrueString Then
                    TBLDummy.Rows.Clear()
                    DTADummy.Fill(TBLDummy, rowDGV.Cells(Me.ColIdentifierJO.Index).Value, rowDGV.Cells(Me.ColPrimaryMO.Index).Value)

                    For Each rowDummy In TBLDummy.Rows
                        If rowDummy.Jml <= 0 Then
                            TBLDummy1.DefaultView.RowFilter = "[Primary Product] = " + rowDummy.Primary_Product.ToString
                            If TBLDummy1.DefaultView.Count = 0 Then
                                rowDummy1 = TBLDummy1.Newtbltrans_mo_item_wo3Row
                                For ij As Integer = 0 To TBLDummy1.Columns.Count - 1
                                    rowDummy1.Item(ij) = rowDummy.Item(ij)
                                Next
                                TBLDummy1.Rows.Add(rowDummy1)
                            End If
                        End If
                    Next
                End If
            Next

            If TBLDummy1.Rows.Count > 0 Then
                Frm1.TBLDummy = TBLDummy1
                Frm1.ShowDialog()

                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTADummy.Dispose()
            TBLDummy.Dispose()
            TBLDummy1.Dispose()
            Frm1.Dispose()
        End Try
    End Function

    'TR// Ini update perihal Penambahan Filter

    Private Sub btnFilter_FindCodeNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter_FindCodeNo.Click
        Dim i As New Core_Module.frmCariMasterProduct

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtFilter_CodeNo.Text <> "" Then
                i.txtCari.Text = Me.TxtFilter_CodeNo.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.TxtFilter_CodeNo.Text = Me.ObjProd.GetProductCode(i.PrimaryProduct)
                Me.LoadUlang()
            Else
                Me.TxtFilter_CodeNo.Clear()
                Me.TxtFilter_CodeNo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtFilter_MONo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtFilter_MONo.KeyDown
        Try
            Me.Cursor = Cursors.WaitCursor
            If e.KeyCode = 13 Then
                If Me.TxtFilter_MONo.Text <> "" Then
                    If Me.ObjMO.Apakahsudahada(Me.TxtFilter_MONo.Text) = True Then
                        Me.TxtFilter_PrimaryMO.Text = Me.ObjMO.GetPrimary(Me.TxtFilter_MONo.Text)
                        Me.TxtFilter_MONo.Text = Me.ObjMO.GetNoTransaksi(Me.TxtFilter_PrimaryMO.Text)
                    Else
                        Me.btnFilter_FindMO.PerformClick()
                    End If
                Else
                    Me.TxtFilter_PrimaryMO.Text = "-1"
                    Me.TxtFilter_MONo.Clear()
                End If

                Me.LoadUlang()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub TxtFilter_CodeNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFilter_CodeNo.GotFocus, TxtFilter_Desc.GotFocus, _
    TxtFilter_MONo.GotFocus
        Dim txt As TextBox = TryCast(sender, TextBox)
        txt.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtFilter_CodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtFilter_CodeNo.KeyDown
        Try
            If e.KeyValue = 13 Then
                Me.LoadUlang()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtFilter_Desc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtFilter_Desc.KeyDown
        Try
            If e.KeyValue = 13 Then
                Me.LoadUlang()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtFilter_CodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFilter_CodeNo.LostFocus, TxtFilter_Desc.LostFocus, _
    TxtFilter_MONo.LostFocus
        Dim txt As TextBox = TryCast(sender, TextBox)
        txt.BackColor = Color.White
    End Sub

    Private Sub TxtFilter_MONo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFilter_MONo.Leave
        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.TxtFilter_MONo.Text <> "" Then
                If Me.ObjMO.Apakahsudahada(Me.TxtFilter_MONo.Text) = True Then
                    Me.TxtFilter_PrimaryMO.Text = Me.ObjMO.GetPrimary(Me.TxtFilter_MONo.Text)
                    Me.TxtFilter_MONo.Text = Me.ObjMO.GetNoTransaksi(Me.TxtFilter_PrimaryMO.Text)
                Else
                    Me.btnFilter_FindMO.PerformClick()
                End If
            Else
                Me.TxtFilter_PrimaryMO.Text = "-1"
                Me.TxtFilter_MONo.Clear()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btnFilter_FindMO_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFilter_FindMO.Click
        Dim frm As New FrmTrans_ManufacturingOrder_Cari

        Try
            If Me.TxtFilter_MONo.Text <> "" Then
                frm.txtCari.Text = Me.TxtFilter_MONo.Text
            End If

            frm.ShowDialog()
            If frm.Primary > 0 Then
                Me.TxtFilter_PrimaryMO.Text = frm.Primary
                Me.TxtFilter_MONo.Text = Me.ObjMO.GetNoTransaksi(frm.Primary)
            Else
                Me.TxtFilter_PrimaryMO.Text = "-1"
                Me.TxtFilter_MONo.Clear()
            End If

            Me.LoadUlang()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            frm.Dispose()
        End Try
    End Sub


    Private Sub LoadUlang()
        Try
            Me.ObjMOItem.IsiDGVJadwalBaru(Me.dgvItem, Me)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    
    Private Sub llbMP_Filter_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbMP_Filter.LinkClicked
        Try
            LoadUlang()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class