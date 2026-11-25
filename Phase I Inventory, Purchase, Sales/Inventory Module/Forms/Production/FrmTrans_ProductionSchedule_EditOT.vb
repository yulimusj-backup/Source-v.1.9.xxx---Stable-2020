Option Explicit On
Option Strict Off

Public Class FrmTrans_work_unitionSchedule_EditOT
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public IsEdited As Boolean = False

    Dim ObjSPL As New ClsTrans_SPL
    Dim ObjAntrian As New Core_Module.ClsAntrian
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjWUGroup As New Core_Module.ClsMasterWorkUnitGroup
    Dim ObjWU As New Core_Module.ClsMasterWorkUnit
    Dim ObjSch As New ClsTrans_Schedule
    Dim buffer1 As String = ""
    Dim DummyIndex As Integer = -1
    Dim UserLain, NamaKomputerLain As String
    Dim mClose As Boolean = False

    Public Enum EnumGroup As Integer
        WorkUnitGroup
        WorkUnit
    End Enum

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSetOvertime.Click
        Dim ij As New FrmTrans_ProductionSchedule_EditOT2

        Try
            ij.UserCode = Me.UserCode
            ij.BackColorAll = Me.BackColorAll
            ij.IsEdited = Me.IsEdited
            If ij.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Me.IsEdited = ij.IsEdited
                Me.btnFilter.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Try
            Me.LoadDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDGV()
        Dim DTA1 As New DseTrans_OT_ItemTableAdapters.tbltrans_overtime_item1TableAdapter
        Dim TBL1 As New DseTrans_OT_Item.tbltrans_overtime_item1DataTable
        Dim row1, row2 As DseTrans_OT_Item.tbltrans_overtime_item1Row
        Dim rowDGV As DataGridViewRow
        Dim CboDgv As Windows.Forms.DataGridViewComboBoxCell
        Dim LevelWU As EnumGroup = EnumGroup.WorkUnit
        Dim ListPrimarySPL As New List(Of Double)
        Dim NoSPL As String = ""
        Dim ShowEndDate As Boolean = False
        Dim Ketemu As Boolean = False
        Dim nb As Integer = 1
        Dim mn As Integer = 0

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If Me.TxtSPLNo.Text <> "" Then
                DTA1.FillBy_primary_spl(TBL1, Me.TxtPrimarySPL.Text)
            Else
                DTA1.FillBy_Tgl(TBL1, CDate(Format(Me.dtpEndDate.Value, "dd MMM yyyy") + " 23:59:00"), Me.dtpStartDate.Value)
            End If


            If Me.dgvFilterWU.Rows.Count > 1 Then
                For Each rowDGV In Me.dgvFilterWU.Rows
                    If rowDGV.Index <> Me.dgvFilterWU.Rows.Count - 1 Then
                        CboDgv = Me.dgvFilterWU.Item(Me.ColFilterWU_Level.Index, rowDGV.Index)
                        LevelWU = CboDgv.Items.IndexOf(Me.dgvFilterWU.Item(Me.ColFilterWU_Level.Index, rowDGV.Index).Value)

                        If LevelWU = EnumGroup.WorkUnitGroup Then
                            'filter berdasarkan primary work unit group
                            For Each row1 In TBL1.Rows
                                If row1.primary_work_unit_group = rowDGV.Cells(Me.ColFilterWU_Primary.Index).Value Then
                                    row1.load_to_dgv = 1
                                End If
                            Next
                        ElseIf LevelWU = EnumGroup.WorkUnit Then
                            'filter berdasarkan primary work unit
                            For Each row1 In TBL1.Rows
                                If row1.primary_workunit = rowDGV.Cells(Me.ColFilterWU_Primary.Index).Value Then
                                    row1.load_to_dgv = 1
                                End If
                            Next
                        End If
                    End If
                Next
            Else
                For Each row1 In TBL1.Rows
                    row1.load_to_dgv = 1
                Next
            End If
            'buang baris yang tidak digunakan
            If TBL1.Rows.Count > 0 Then
                For vb As Integer = TBL1.Rows.Count - 1 To 0 Step 1
                    row1 = TBL1.Rows(vb)
                    If row1.load_to_dgv <> 1 Then
                        TBL1.Rows.RemoveAt(vb)
                    End If
                Next
            End If

            'isikan no spl
            For Each row1 In TBL1.Rows
                If row1.primary_spl <> -1 Then
                    If ListPrimarySPL.Contains(row1.primary_spl) = True Then
                        ListPrimarySPL.Add(row1.primary_spl)
                    End If
                End If
            Next
            If ListPrimarySPL.Count > 0 Then
                For Each PrimarySPL As Double In ListPrimarySPL
                    NoSPL = Me.ObjSPL.GetNoDokumen(PrimarySPL)
                    TBL1.DefaultView.RowFilter = "primary spl = " + PrimarySPL.ToString
                    For Each DV As Data.DataRowView In TBL1.DefaultView
                        row1 = DV.Row
                        row1.spl_no = NoSPL
                    Next
                Next
            End If

            'hanya tampilkan yang sesuai saja filternya
            Me.dgvOT.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvOT.Rows.Clear()
            For Each row1 In TBL1.Rows
                If row1.load_to_dgv = 1 Then
                    Me.dgvOT.Rows.Add()
                    rowDGV = Me.dgvOT.Rows(Me.dgvOT.Rows.Count - 1)
                    rowDGV.Cells(Me.ColEndDate.Index).Value = Format(row1._end, "dd MMM yyyy")
                    rowDGV.Cells(Me.ColEndTime.Index).Value = Format(row1._end, "HH.mm")
                    rowDGV.Cells(Me.ColPrimarySPL.Index).Value = row1.primary_spl
                    rowDGV.Cells(Me.ColRemarks.Index).Value = ""
                    rowDGV.Cells(Me.ColSPLNo.Index).Value = row1.spl_no
                    rowDGV.Cells(Me.ColStartDate.Index).Value = Format(row1.start, "dd MMM yyyy")
                    rowDGV.Cells(Me.ColStartTime.Index).Value = Format(row1.start, "HH.mm")
                    rowDGV.Cells(Me.ColWorkUnitCode.Index).Value = row1.work_unit_code
                    rowDGV.Cells(Me.ColWorkUnitDesc.Index).Value = row1.description
                    rowDGV.Cells(Me.ColPrimary.Index).Value = row1.primary
                    rowDGV.Cells(Me.ColPrimaryWU.Index).Value = row1.primary_workunit

                    If row1.start.Date <> row1._end.Date Then
                        ShowEndDate = True
                    End If
                End If
            Next

            'gabungkan baris yang sama
            If ShowEndDate = True Then
                Me.ColEndDate.Visible = True
            Else
                Me.ColEndDate.Visible = False
            End If

            Me.dgvOT.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Dispose()
            ListPrimarySPL = Nothing
        End Try
    End Sub

    Private Sub FrmTrans_work_unitionSchedule_EditOT_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            If Me.mClose = True Then
                MsgBox("Can't access overtime transaction because user code : " + Me.UserLain.ToString + _
                " at computer name-ip address : " + Me.NamaKomputerLain.ToString + " is using this transaction.", _
                MsgBoxStyle.Information, "Information")
            End If
            Me.ObjAntrian.EndAntrianTunggal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_work_unitionSchedule_EditOT_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.lblFirmPlanDate.Text = Format(Me.ObjSch.GetFirmPlan, "dd MMM yyyy")

            Me.dgvFilterWU.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvOT.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            UserLain = ""
            NamaKomputerLain = ""
            If Me.ObjAntrian.StartAntrianTunggal(Core_Module.ClsAntrian.EnumTipeAntrian.AntrianTunggalSchedulingOT, Me.UserCode, UserLain, NamaKomputerLain) = False Then
                Me.mClose = True
                Me.BackgroundWorker_CloseForm.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker_CloseForm_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_CloseForm.DoWork
        Try
            Threading.Thread.Sleep(1)
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

    Private Sub dgvFilterWU_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvFilterWU.CellBeginEdit
        Dim rowDGV As DataGridViewRow

        Try
            'yang bisa diedit:
            '1. jika baris baru
            If e.RowIndex = Me.dgvFilterWU.Rows.Count - 1 Then
                If e.ColumnIndex = Me.ColFilterWU_Level.Index Then
                    'pastikan baris diatasnya sudah beres dahulu
                    If e.RowIndex > 0 Then
                        rowDGV = Me.dgvFilterWU.Rows(e.RowIndex - 1)
                        If rowDGV.Cells(Me.ColFilterWU_Code.Index).Value.ToString.Trim = "" Then
                            MsgBox("Please specify " + rowDGV.Cells(Me.ColFilterWU_Level.Index).Value.ToString + _
                            " code on the previous row.", MsgBoxStyle.Information, "Information")
                            Me.dgvFilterWU.ClearSelection()
                            rowDGV.Selected = True
                            e.Cancel = True
                            Exit Sub
                        Else
                            Me.buffer1 = Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value
                        End If
                    Else
                        Me.buffer1 = Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    e.Cancel = True
                End If
            Else
                If Me.dgvFilterWU.Item(Me.ColFilterWU_Primary.Index, e.RowIndex).Value = -1 Then
                    '2. jika bukan baris baru dan primary =-1
                    'yang bisa diedit hanya code work_unit group dan level
                    If e.ColumnIndex = Me.ColFilterWU_Code.Index Or e.ColumnIndex = Me.ColFilterWU_Level.Index Then
                        Me.buffer1 = Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                Else
                    '3. jika bukan baris baru dan primary <> -1
                    'yang bisa diedit: code, remarks
                    If e.ColumnIndex = Me.ColFilterWU_Code.Index Then
                        Me.buffer1 = Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Try
            If Me.DummyIndex <> -1 Then
                Me.dgvFilterWU.Rows.RemoveAt(Me.DummyIndex)
                DummyIndex = -1
                Me.ObjCD.SetNumber(Me.dgvFilterWU)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvFilterWU_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvFilterWU.CellEndEdit
        Dim Level As EnumGroup = EnumGroup.WorkUnit
        Dim CboDgv As Windows.Forms.DataGridViewComboBoxCell

        Try
            If IsNothing(Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value) = True _
                       Or Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColFilterWU_Code.Index Then
                    'kalo kode work_unit group dihapus, maka hapus satu bari semuanya
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                    Exit Sub
                End If
            Else
                'level
                If e.ColumnIndex = Me.ColFilterWU_Level.Index Then
                    'do nothing

                ElseIf e.ColumnIndex = Me.ColFilterWU_Code.Index Then
                    Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
                    If Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If
                    'cek ke master
                    CboDgv = Me.dgvFilterWU.Item(Me.ColFilterWU_Level.Index, e.RowIndex)
                    Level = CboDgv.Items.IndexOf(Me.dgvFilterWU.Item(Me.ColFilterWU_Level.Index, e.RowIndex).Value)

                    If Level = EnumGroup.WorkUnitGroup Then
                        Me.EndEditWorkUnitGroup(e)
                    Else
                        Me.EndEditwork_unit(e)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditWorkUnitGroup(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If Me.ObjWUGroup.Apakahsudahada(Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                '1. kode work unit group ada  --> load data nama + primary
                Me.dgvFilterWU.Item(Me.ColFilterWU_Primary.Index, e.RowIndex).Value = Me.ObjWUGroup.GetPrimary(Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value)
                Me.dgvFilterWU.Item(Me.ColFilterWU_Desc.Index, e.RowIndex).Value = Me.ObjWUGroup.GetDesc(Me.dgvFilterWU.Item(Me.ColFilterWU_Primary.Index, e.RowIndex).Value)
            Else
                '2. kode work unit group yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.FrmMasterWorkUnitGroup_Cari
                ij.txtCari.Text = Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryWorkUnitGroup <> -1 Then
                    '2.a. work unit group yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvFilterWU.Item(Me.ColFilterWU_Primary.Index, e.RowIndex).Value = ij.PrimaryWorkUnitGroup
                    Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjWUGroup.GetCode(ij.PrimaryWorkUnitGroup)
                    Me.dgvFilterWU.Item(Me.ColFilterWU_Desc.Index, e.RowIndex).Value = Me.ObjWUGroup.GetDesc(ij.PrimaryWorkUnitGroup)
                ElseIf ij.PrimaryWorkUnitGroup = -1 Then
                    '2.b. kode work unit group yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer1
                    If Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If

            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemWorkUnit(Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvFilterWU.Item(Me.ColFilterWU_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvFilterWU.Item(Me.ColFilterWU_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If buffer1 = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer1
                    Me.dgvFilterWU.Item(Me.ColFilterWU_Primary.Index, e.RowIndex).Value = Me.ObjWUGroup.GetPrimary(Me.buffer1)
                    Me.dgvFilterWU.Item(Me.ColFilterWU_Desc.Index, e.RowIndex).Value = Me.ObjWUGroup.GetDesc(Me.dgvFilterWU.Item(Me.ColFilterWU_Primary.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditwork_unit(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If Me.ObjWU.ApakahSudahAda(Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                '1. kode work_unit ada  --> load data nama + primary
                Me.dgvFilterWU.Item(Me.ColFilterWU_Primary.Index, e.RowIndex).Value = Me.ObjWU.GetPrimary(Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value)
                Me.dgvFilterWU.Item(Me.ColFilterWU_Desc.Index, e.RowIndex).Value = Me.ObjWU.GetDesc(Me.dgvFilterWU.Item(Me.ColFilterWU_Primary.Index, e.RowIndex).Value)
            Else
                '2. kode work_unit yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.FrmMasterWorkUnit_Cari
                ij.txtCari.Text = Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryWorkUnit <> -1 Then
                    '2.a. work_unit yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvFilterWU.Item(Me.ColFilterWU_Primary.Index, e.RowIndex).Value = ij.PrimaryWorkUnit
                    Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjWU.GetCode(ij.PrimaryWorkUnit)
                    Me.dgvFilterWU.Item(Me.ColFilterWU_Desc.Index, e.RowIndex).Value = Me.ObjWU.GetDesc(ij.PrimaryWorkUnit)
                ElseIf ij.PrimaryWorkUnit = -1 Then
                    '2.b. kode work_unit yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer1
                    If Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If

            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemWorkUnit(Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvFilterWU.Item(Me.ColFilterWU_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvFilterWU.Item(Me.ColFilterWU_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If buffer1 = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer1
                    Me.dgvFilterWU.Item(Me.ColFilterWU_Primary.Index, e.RowIndex).Value = Me.ObjWU.GetPrimary(Me.buffer1)
                    Me.dgvFilterWU.Item(Me.ColFilterWU_Desc.Index, e.RowIndex).Value = Me.ObjWU.GetDesc(Me.dgvFilterWU.Item(Me.ColFilterWU_Primary.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndex_ItemWorkUnit(ByVal vCode As String, ByVal vRefIndex As Integer, ByVal vLevelStr As String) As Integer
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvFilterWU.Rows
                If rowDGV.Cells(Me.ColFilterWU_Code.Index).Value = vCode And _
                rowDGV.Index <> vRefIndex And _
                rowDGV.Cells(Me.ColFilterWU_Level.Index).Value = vLevelStr Then
                    Return rowDGV.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub dgvFilterWU_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvFilterWU.DataError
        'do nothing -> dibuat untuk handle data error event, jangan dihapuskan
    End Sub

    Private Sub dgvFilterWU_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvFilterWU.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgvFilterWU.Rows(Me.dgvFilterWU.Rows.Count - 1)
            rowDGV.Cells(Me.ColFilterWU_Code.Index).Value = ""
            rowDGV.Cells(Me.ColFilterWU_Desc.Index).Value = ""
            rowDGV.Cells(Me.ColFilterWU_Level.Index).Value = "Work Unit Group"
            rowDGV.Cells(Me.ColFilterWU_No.Index).Value = rowDGV.Index + 1
            rowDGV.Cells(Me.ColFilterWU_Primary.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvFilterWU_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvFilterWU.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvFilterWU)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindSPL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindSPL.Click
        Dim FrmCariGroup As New FrmTrans_SPL_Cari
        Dim ObjGroup As New ClsTrans_SPL

        Try
            FrmCariGroup.BackColorAll = Me.BackColorAll
            If Me.TxtPrimarySPL.Text = "-1" And Me.TxtSPLNo.Text = "" Then
                FrmCariGroup.ShowDialog()
            Else
                FrmCariGroup.txtCari.Text = Me.TxtSPLNo.Text
                FrmCariGroup.ShowDialog()
            End If
            If CDbl(FrmCariGroup.PrimaryPP) > 0 Then
                Me.TxtPrimarySPL.Text = FrmCariGroup.PrimaryPP
                Me.TxtSPLNo.Text = ObjGroup.GetNoDokumen(FrmCariGroup.PrimaryPP)
                Me.btnAddBySPL.Focus()
            Else
                Me.TxtPrimarySPL.Text = -1
                Me.TxtSPLNo.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmCariGroup.Dispose()
        End Try
    End Sub

    Private Sub TxtSPLNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSPLNo.GotFocus
        Try
            If Me.TxtSPLNo.Enabled = True Then
                If Me.TxtSPLNo.ReadOnly = True Then
                    Me.TxtSPLNo.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtSPLNo.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSPLNo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtSPLNo.KeyDown
        Dim ObjGroup As New ClsTrans_SPL
        Try
            If e.KeyValue = 13 Then
                If ObjGroup.ApakahSudahAda(Me.TxtSPLNo.Text) = True Then
                    Me.TxtPrimarySPL.Text = ObjGroup.GetPrimary(Me.TxtSPLNo.Text)
                    Me.btnAddBySPL.Focus()
                Else
                    Me.btnFindSPL.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSPLNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSPLNo.Leave
        Dim ObjGroup As New ClsTrans_SPL
        Try
            If Me.TxtSPLNo.Text <> "" Then
                If ObjGroup.ApakahSudahAda(Me.TxtSPLNo.Text) = True Then
                    Me.TxtPrimarySPL.Text = ObjGroup.GetPrimary(Me.TxtSPLNo.Text)
                Else
                    Me.btnFindSPL.PerformClick()
                End If
            Else
                Me.TxtPrimarySPL.Text = -1
                Me.TxtSPLNo.Text = ""
            End If
            Me.TxtSPLNo.BackColor = Me.TxtPrimarySPL.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSPLNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSPLNo.LostFocus
        Try
            If Me.TxtSPLNo.Enabled = True Then
                If Me.TxtSPLNo.ReadOnly = True Then
                    Me.TxtSPLNo.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtSPLNo.BackColor = Me.TxtPrimarySPL.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnRemoveOvertime_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemoveOvertime.Click
        Dim Ketemu As Boolean = False
        Dim rowDGV As DataGridViewRow
        Dim TglFirmPlan As Date = Today
        Dim DTA As New DseTrans_OT_ItemTableAdapters.tbltrans_overtime_itemTableAdapter

        Try
            'cek apakah ada yang perlu di delete
            For Each rowDGV In Me.dgvOT.Rows
                If rowDGV.Cells(Me.ColDelete.Index).Value = Boolean.TrueString Then
                    Ketemu = True
                    Exit For
                End If
            Next

            If Ketemu = True Then
                'pastikan yang didelete tidak bertanggal sebelum atau sama dengan firm plan
                TglFirmPlan = Me.ObjSch.GetFirmPlan
                Ketemu = False
                For Each rowDGV In Me.dgvOT.Rows
                    If rowDGV.Cells(Me.ColDelete.Index).Value = Boolean.TrueString Then
                        If CDate(rowDGV.Cells(Me.ColStartDate.Index).Value).Date <= TglFirmPlan Then
                            Ketemu = True
                            Exit For
                        End If
                    End If
                Next
                If Ketemu = True Then
                    MsgBox("Can't delete item dated <= " + Format(TglFirmPlan, "dd MMM yyyy") + ".", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                'delete sch bar
                If MsgBox("Delete overtime data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Dim RowWorkunit As DseTrans_OT_Item.DataTable1Row
                    Dim TblWorkunit As New DseTrans_OT_Item.DataTable1DataTable
                    Dim vTBLSchBar As New DseTrans_Schedule.tbl_sch_barDataTable
                    Dim vTBLSpaceBeforeDeleteOT As New DseTrans_Schedule.tbl_sch_barDataTable
                    Dim objSchOT As New ClsTrans_Schedule_OT
                    Dim TBLUpdate As New DseTrans_OT_Item.DataTable2DataTable

                    For Each rowDGV In Me.dgvOT.Rows
                        If rowDGV.Cells(Me.ColDelete.Index).Value = Boolean.TrueString Then
                            TblWorkunit.DefaultView.RowFilter = "[Primary Workunit] = " & rowDGV.Cells(Me.ColPrimaryWU.Index).Value.ToString
                            If TblWorkunit.DefaultView.Count = 0 Then
                                RowWorkunit = TblWorkunit.NewDataTable1Row
                                RowWorkunit.Primary_Workunit = rowDGV.Cells(Me.ColPrimaryWU.Index).Value
                                RowWorkunit.Minimum_Start_Deleted_OT = CDate(rowDGV.Cells(Me.ColStartDate.Index).Value & " " & rowDGV.Cells(Me.ColStartTime.Index).Value.ToString.Replace(".", ":"))
                                TblWorkunit.Rows.Add(RowWorkunit)
                            Else
                                RowWorkunit = TblWorkunit.DefaultView(0).Row
                                If RowWorkunit.Minimum_Start_Deleted_OT > CDate(rowDGV.Cells(Me.ColStartDate.Index).Value & " " & rowDGV.Cells(Me.ColStartTime.Index).Value.ToString.Replace(".", ":")) Then
                                    RowWorkunit.Minimum_Start_Deleted_OT = CDate(rowDGV.Cells(Me.ColStartDate.Index).Value & " " & rowDGV.Cells(Me.ColStartTime.Index).Value.ToString.Replace(".", ":"))
                                End If
                            End If
                        End If
                    Next

                    ObjSch.LoadScheduleBar_WOItemOnly(#1/1/1950#, #12/31/2200#, vTBLSchBar)
                    ObjSch.LoadScheduleBar_TanpaScheduledItem(ObjSch.GetFirmPlan, #12/31/2200#, vTBLSpaceBeforeDeleteOT)
                    For Each RowWorkunit In TblWorkunit.Rows
                        objSchOT.RecalculateAfterDeleteOT(RowWorkunit.Primary_Workunit, RowWorkunit.Minimum_Start_Deleted_OT, vTBLSpaceBeforeDeleteOT, vTBLSchBar, Me.dgvOT, TBLUpdate)
                    Next

                    If Not objSchOT.EnoughAllocation(TBLUpdate) Then
                        Exit Sub
                    End If
                    If objSchOT.AnyClosedJO(TBLUpdate) Then
                        Exit Sub
                    End If

                    Dim message As String = ""
                    If objSchOT.AnyProducedJO(TBLUpdate, message) Then
                        MsgBox("Cannot remove overtime, there are log transaction for the following JO: " & vbNewLine & message & ".", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If

                    objSchOT.UpdateScheduleChange(TBLUpdate)

                    DTA.Connection.ConnectionString = Core_Module.GetConstring
                    For Each rowDGV In Me.dgvOT.Rows
                        If rowDGV.Cells(Me.ColDelete.Index).Value = Boolean.TrueString Then
                            DTA.Delete(rowDGV.Cells(Me.ColPrimary.Index).Value)
                        End If
                    Next

                    Me.LoadDGV()
                    Me.IsEdited = True
                    MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub dgvOT_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvOT.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColEndTime.Index Or e.ColumnIndex = Me.ColStartTime.Index Or e.ColumnIndex = Me.ColDelete.Index Then
                Me.buffer1 = Me.dgvOT.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvOT_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvOT.CellEndEdit
        Dim StartTime, EndTime As Date
        Dim DTAWS As New DseTrans_OT_ItemTableAdapters.tbltrans_mo_workshift_allocateTableAdapter
        Dim TBLWS As New DseTrans_OT_Item.tbltrans_mo_workshift_allocateDataTable
        Dim rowWS, rowWS_Next As DseTrans_OT_Item.tbltrans_mo_workshift_allocateRow
        Dim DTAWS1 As New DseTrans_OT_ItemTableAdapters.tblmaster_workshift_itemTableAdapter
        Dim TBLWS1 As New DseTrans_OT_Item.tblmaster_workshift_itemDataTable
        Dim rowWS1 As DseTrans_OT_Item.tblmaster_workshift_itemRow
        Dim DTAOT As New DseTrans_OT_ItemTableAdapters.tbltrans_overtime_itemTableAdapter
        Dim TBLOT As New DseTrans_OT_Item.tbltrans_overtime_itemDataTable
        Dim rowOT As DseTrans_OT_Item.tbltrans_overtime_itemRow
        Dim Ketemu As Boolean = True
        Dim WSStart, WSEnd As Date
        Dim mn, jk As Double
        Dim TipeWorkDays As Core_Module.ClsMasterWorkShift.TipeWorkdays = Core_Module.ClsMasterWorkShift.TipeWorkdays.Holiday
        Dim DTA4 As New DseTrans_OT_ItemTableAdapters.tblmaster_workshift_itemholidayTableAdapter
        Dim TBL4 As New DseTrans_OT_Item.tblmaster_workshift_itemholidayDataTable

        Try
            If e.ColumnIndex = Me.ColDelete.Index Then
                Exit Sub
            End If

            DTAWS.Connection.ConnectionString = Core_Module.GetConstring
            DTAWS1.Connection.ConnectionString = DTAWS.Connection.ConnectionString
            DTAOT.Connection.ConnectionString = DTAWS.Connection.ConnectionString
            DTA4.Connection.ConnectionString = DTAWS.Connection.ConnectionString

            If IsNothing(Me.dgvOT.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                Me.dgvOT.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer1
                Exit Sub
            Else
                If e.ColumnIndex = Me.ColEndDate.Index Or e.ColumnIndex = Me.ColStartDate.Index Then
                    'cek tipe data
                    If IsDate(Me.dgvOT.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.dgvOT.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer1
                        Exit Sub
                    End If
                    'format tanggal
                    Me.dgvOT.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.dgvOT.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")
                ElseIf e.ColumnIndex = Me.ColEndTime.Index Or e.ColumnIndex = Me.ColStartTime.Index Then
                    'cek tipe data
                    If Me.ObjCD.Is_Time_String(Me.dgvOT.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.dgvOT.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer1
                        Exit Sub
                    End If
                End If

                'cek data awal < data akhir
                StartTime = CDate(Me.dgvOT.Item(Me.ColStartDate.Index, e.RowIndex).Value) + TimeSpan.FromMinutes(Me.ObjCD.Convert_ToMinutes(Me.dgvOT.Item(Me.ColStartTime.Index, e.RowIndex).Value))
                EndTime = CDate(Me.dgvOT.Item(Me.ColEndDate.Index, e.RowIndex).Value) + TimeSpan.FromMinutes(Me.ObjCD.Convert_ToMinutes(Me.dgvOT.Item(Me.ColEndTime.Index, e.RowIndex).Value))
                If StartTime >= EndTime Then
                    MsgBox("Please make sure start time < end time.", MsgBoxStyle.Information, "Information")
                    Me.dgvOT.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer1
                    Exit Sub
                End If

                'start date dan end date harus ada workshift allocatenya
                DTAWS.FillBy_Primary_WorkUnit(TBLWS, EndTime.Date + TimeSpan.FromDays(1), StartTime.Date - TimeSpan.FromDays(1), _
                Me.dgvOT.Item(Me.ColPrimaryWU.Index, e.RowIndex).Value)
                Ketemu = False
                For Each rowWS In TBLWS.Rows
                    If rowWS._date = StartTime.Date Then
                        Ketemu = True
                        Exit For
                    End If
                Next
                If Ketemu = False Then
                    MsgBox("Please allocate workshift until " + Format(EndTime, "dd MMM yyyy") + ".", MsgBoxStyle.Information, "Information")
                    Me.dgvOT.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer1
                    Exit Sub
                End If
                Ketemu = False
                For Each rowWS In TBLWS.Rows
                    If rowWS._date = EndTime.Date Then
                        Ketemu = True
                        Exit For
                    End If
                Next
                If Ketemu = False Then
                    MsgBox("Please allocate workshift until " + Format(EndTime, "dd MMM yyyy") + ".", MsgBoxStyle.Information, "Information")
                    Me.dgvOT.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer1
                    Exit Sub
                End If

                'cek tidak berbenturan dengan work shift
                Ketemu = False

                For Each rowWS In TBLWS.Rows
                    rowWS.day_of_week = rowWS._date.DayOfWeek
                Next

                mn = 0
                DTA4.Fill(TBL4, CDate(Me.dgvOT.Item(Me.ColStartDate.Index, e.RowIndex).Value) - TimeSpan.FromDays(1))
                For Each rowWS In TBLWS.Rows
                    If rowWS.sch_in = -1 Then
                        DTAWS1.FillBy1(TBLWS1, rowWS.day_of_week, rowWS.primary_workshift)
                        rowWS1 = TBLWS1.Rows(0)

                        rowWS.sch_in = rowWS1.schedule_in
                        rowWS.sch_out = rowWS1.schedule_out
                        rowWS.tipe_workdays = rowWS1.tipe_workdays

                        'copy ke semua baris bawahnya
                        jk = mn + 1
                        While jk <= TBLWS.Rows.Count - 1
                            rowWS_Next = TBLWS.Rows(jk)
                            If rowWS_Next.primary_workshift = rowWS.primary_workshift And rowWS_Next.day_of_week = rowWS.day_of_week Then
                                rowWS_Next.sch_in = rowWS.sch_in
                                rowWS_Next.sch_out = rowWS.sch_out
                                rowWS_Next.tipe_workdays = rowWS.tipe_workdays
                            End If
                            jk += 1
                        End While

                        
                    End If
                    mn += 1
                Next
                For Each rowWS In TBLWS.Rows
                    'cek hari libur
                    TBL4.DefaultView.RowFilter = "[primary transaksi] = " + rowWS.primary_workshift.ToString + " AND date = #" + Format(rowWS._date, "MM/dd/yyyy") + "#"
                    If TBL4.DefaultView.Count > 0 Then
                        'holiday
                        rowWS.tipe_workdays = Core_Module.ClsMasterWorkShift.TipeWorkdays.Holiday
                    End If
                Next
                For Each rowWS In TBLWS.Rows
                    TipeWorkDays = rowWS.tipe_workdays
                    If TipeWorkDays = Core_Module.ClsMasterWorkShift.TipeWorkdays.Weekday Then
                        If rowWS.sch_in < rowWS.sch_out Then
                            WSStart = rowWS._date + TimeSpan.FromMinutes(rowWS.sch_in)
                            WSEnd = rowWS._date + TimeSpan.FromMinutes(rowWS.sch_out)
                        Else
                            WSStart = rowWS._date + TimeSpan.FromMinutes(rowWS.sch_in)
                            WSEnd = rowWS._date + TimeSpan.FromDays(1) + TimeSpan.FromMinutes(rowWS.sch_out)
                        End If
                        If (StartTime >= WSEnd And EndTime > WSEnd) Or _
                        (EndTime <= WSStart And StartTime < WSStart) Then
                        Else
                            Ketemu = True
                            Exit For
                        End If
                    End If
                Next
                If Ketemu = True Then
                    MsgBox("Conflicting with reguler work time, please fix.", MsgBoxStyle.Information, "Information")
                    Me.dgvOT.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer1
                    Exit Sub
                End If

                'tidak berbenturan dengan overtime lainnya
                Ketemu = False
                DTAOT.FillBy1(TBLOT, Me.dgvOT.Item(Me.ColPrimaryWU.Index, e.RowIndex).Value, StartTime - TimeSpan.FromDays(1))
                For Each rowOT In TBLOT.Rows
                    If (StartTime >= rowOT._end And EndTime > rowOT._end) Or _
                    (EndTime <= rowOT.start And StartTime < rowOT.start) Then
                    Else
                        If rowOT.primary <> Me.dgvOT.Item(Me.ColPrimary.Index, e.RowIndex).Value Then
                            Ketemu = True
                            Exit For
                        End If
                    End If
                Next
                If Ketemu = True Then
                    MsgBox("Conflicting with another overtime, please fix.", MsgBoxStyle.Information, "Information")
                    Me.dgvOT.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer1
                    Exit Sub
                End If

                'sampai sini, simpan ke dalam database
                DTAOT.Update(StartTime, EndTime, Me.dgvOT.Item(Me.ColPrimary.Index, e.RowIndex).Value)

                'MsgBox("belum coding pergeseran item.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAWS.Dispose()
            TBLWS.Dispose()
            DTAOT.Dispose()
            TBLOT.Dispose()
            DTAWS1.Dispose()
            TBLWS1.Dispose()
            DTA4.Dispose()
            TBL4.Dispose()
        End Try
    End Sub

    Private Sub llbCheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCheckAll.LinkClicked
        Dim rowDGV As DataGridViewRow
        Dim FirmPlanDate As Date = Today

        Try
            FirmPlanDate = Me.ObjSch.GetFirmPlan

            For Each rowDGV In Me.dgvOT.Rows
                If CDate(rowDGV.Cells(Me.ColStartDate.Index).Value).Date > FirmPlanDate Then
                    rowDGV.Cells(Me.ColDelete.Index).Value = Boolean.TrueString
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUncheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUncheckAll.LinkClicked
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvOT.Rows
                rowDGV.Cells(Me.ColDelete.Index).Value = Boolean.FalseString
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpEndDate_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpEndDate.Enter
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpEndDate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpEndDate.Leave
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        Me.Close()
    End Sub
End Class