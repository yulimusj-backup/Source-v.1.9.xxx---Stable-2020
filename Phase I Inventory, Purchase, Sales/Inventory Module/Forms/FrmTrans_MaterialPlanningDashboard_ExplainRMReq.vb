Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_ExplainRMReq
    Public vPrimaryProduct As Double = -1
    Public vDT_ItemRMReq As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqDataTable
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = ""
    Public CodeNo, UnitCode As String
    Public AdaPerubahan As Boolean = False
    Public ReadOnly1 As Boolean = False

    Dim RowItemRMReq As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqRow
    Dim DefPrec As Integer = 0
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser
    Dim LoadData_CheckChange As Boolean = False
    Dim TBLProses As New Core_Module.DseMasterProses.tblmaster_proses1DataTable
    Dim ObjProses As New Core_Module.ClsMasterProses
    Dim Buffer As String = ""
    Dim ObjMO As New ClsTrans_ManufacturingOrder_Item
    Dim ObjMO1 As New ClsTrans_ManufacturingOrder
    Dim ObjSch As New ClsTrans_Schedule
    Dim ObjUser As New Core_Module.ClsUser

    Private Sub LoadDGV_Group_By_MO_JO()
        Dim DV As Data.DataRowView
        Dim rowDGV As DataGridViewRow
        Dim row, rowBefore As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqRow
        Dim Dummy As String = ""
        Dim Idx As Integer = 0

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.dgvDataRMReq.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvDataRMReq.Rows.Clear()
            Me.ColRMReq_WONo.Visible = False
            Me.ColRMReq_Process.Visible = False
            Me.ColRMReq_Process.Visible = False

            'UnitCode = Me.ObjProduct.GetUnitCode(Me.vPrimaryProduct)

            Me.ColRMReq_MODate.DefaultCellStyle.Format = "dd MMM yyyy"
            Me.DefPrec = Me.ObjProduct.GetAngkaDecimal(Me.vPrimaryProduct)
            Me.ColRMReq_BalanceQty.DefaultCellStyle.Format = "N" + Me.DefPrec.ToString
            Me.ColRMReq_SpareQty.DefaultCellStyle.Format = Me.ColRMReq_BalanceQty.DefaultCellStyle.Format
            Me.ColRMReq_TotalQty.DefaultCellStyle.Format = Me.ColRMReq_BalanceQty.DefaultCellStyle.Format

            Me.vDT_ItemRMReq.DefaultView.RowFilter = "[Primary Product] = " + Me.vPrimaryProduct.ToString
            Me.vDT_ItemRMReq.DefaultView.Sort = "[mo date rec], [mo no rec], [jo no rec], [wo no rec]"

            For Each DV In Me.vDT_ItemRMReq.DefaultView
                row = DV.Row

                If Idx = 0 Then
                    Me.dgvDataRMReq.Rows.Add()

                    rowDGV = Me.dgvDataRMReq.Rows(Me.dgvDataRMReq.Rows.Count - 1)
                    rowDGV.Cells(Me.ColRMReq_No.Index).Value = Me.dgvDataRMReq.RowCount
                    rowDGV.Cells(Me.ColRMReq_BalanceQty.Index).Value = row.qty
                    rowDGV.Cells(Me.ColRMReq_JONo.Index).Value = row.jo_no_rec
                    rowDGV.Cells(Me.ColRMReq_MODate.Index).Value = row.mo_date_rec
                    rowDGV.Cells(Me.ColRMReq_MONo.Index).Value = row.mo_no_rec
                    rowDGV.Cells(Me.ColRMReq_WONo.Index).Value = row.wo_no_rec
                    rowDGV.Cells(Me.ColRMReq_UnitCode.Index).Value = UnitCode
                    rowDGV.Cells(Me.ColRMReq_IdentifierJO.Index).Value = row.identifier_jo_rec
                    rowDGV.Cells(Me.ColRMReq_IdentifierWO.Index).Value = row.identifier_wo_rec
                    rowDGV.Cells(Me.ColRMReq_SpareQty.Index).Value = row.MP_Spare_Qty
                    rowDGV.Cells(Me.ColRMReq_TotalQty.Index).Value = row.MP_Spare_Qty + row.qty
                Else
                    rowBefore = Me.vDT_ItemRMReq.DefaultView(Idx - 1).Row
                    If rowBefore.jo_no_rec = row.jo_no_rec And rowBefore.mo_no_rec = row.mo_no_rec Then
                        'hanya tambahkan qty saja
                        rowDGV = Me.dgvDataRMReq.Rows(Me.dgvDataRMReq.RowCount - 1)
                        rowDGV.Cells(Me.ColRMReq_BalanceQty.Index).Value = CDec(rowDGV.Cells(Me.ColRMReq_BalanceQty.Index).Value) + row.qty
                        rowDGV.Cells(Me.ColRMReq_SpareQty.Index).Value = CDec(rowDGV.Cells(Me.ColRMReq_SpareQty.Index).Value) + row.MP_Spare_Qty
                        rowDGV.Cells(Me.ColRMReq_TotalQty.Index).Value = CDec(rowDGV.Cells(Me.ColRMReq_TotalQty.Index).Value) + row.qty + row.MP_Spare_Qty
                    Else
                        Me.dgvDataRMReq.Rows.Add()
                        rowDGV = Me.dgvDataRMReq.Rows(Me.dgvDataRMReq.Rows.Count - 1)
                        rowDGV.Cells(Me.ColRMReq_No.Index).Value = Me.dgvDataRMReq.RowCount
                        rowDGV.Cells(Me.ColRMReq_BalanceQty.Index).Value = row.qty
                        rowDGV.Cells(Me.ColRMReq_JONo.Index).Value = row.jo_no_rec
                        rowDGV.Cells(Me.ColRMReq_MODate.Index).Value = row.mo_date_rec
                        rowDGV.Cells(Me.ColRMReq_MONo.Index).Value = row.mo_no_rec
                        rowDGV.Cells(Me.ColRMReq_WONo.Index).Value = row.wo_no_rec
                        rowDGV.Cells(Me.ColRMReq_UnitCode.Index).Value = UnitCode
                        rowDGV.Cells(Me.ColRMReq_IdentifierJO.Index).Value = row.identifier_jo_rec
                        rowDGV.Cells(Me.ColRMReq_IdentifierWO.Index).Value = row.identifier_wo_rec
                        rowDGV.Cells(Me.ColRMReq_SpareQty.Index).Value = row.MP_Spare_Qty
                        rowDGV.Cells(Me.ColRMReq_TotalQty.Index).Value = row.MP_Spare_Qty + row.qty
                    End If
                End If

                Idx += 1
            Next

            Me.dgvDataRMReq.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvDataRMReq.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDGVDetail()
        Dim DV As Data.DataRowView
        Dim rowDGV As DataGridViewRow
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqRow

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.dgvDataRMReq.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvDataRMReq.Rows.Clear()
            Me.ColRMReq_WONo.Visible = True
            Me.ColRMReq_Process.Visible = True
            Me.ColRMReq_IsWIP.Visible = True
            UnitCode = Me.ObjProduct.GetUnitCode(Me.vPrimaryProduct)
            Me.ObjProses.FillTabelProses(Me.TBLProses)

            Me.ColRMReq_MODate.DefaultCellStyle.Format = "dd MMM yyyy"
            Me.DefPrec = Me.ObjProduct.GetAngkaDecimal(Me.vPrimaryProduct)
            Me.ColRMReq_BalanceQty.DefaultCellStyle.Format = "N" + Me.DefPrec.ToString
            Me.ColRMReq_SpareQty.DefaultCellStyle.Format = Me.ColRMReq_BalanceQty.DefaultCellStyle.Format
            Me.ColRMReq_TotalQty.DefaultCellStyle.Format = Me.ColRMReq_BalanceQty.DefaultCellStyle.Format

            Me.vDT_ItemRMReq.DefaultView.RowFilter = "[Primary Product] = " + Me.vPrimaryProduct.ToString
            Me.vDT_ItemRMReq.DefaultView.Sort = "[mo date rec], [mo no rec], [jo no rec], [wo no rec]"

            For Each DV In Me.vDT_ItemRMReq.DefaultView
                row = DV.Row
                Me.dgvDataRMReq.Rows.Add()

                rowDGV = Me.dgvDataRMReq.Rows(Me.dgvDataRMReq.Rows.Count - 1)
                rowDGV.Cells(Me.ColRMReq_No.Index).Value = Me.dgvDataRMReq.RowCount
                rowDGV.Cells(Me.ColRMReq_BalanceQty.Index).Value = row.qty
                rowDGV.Cells(Me.ColRMReq_SpareQty.Index).Value = row.MP_Spare_Qty
                rowDGV.Cells(Me.ColRMReq_TotalQty.Index).Value = row.qty + row.MP_Spare_Qty
                rowDGV.Cells(Me.ColRMReq_JONo.Index).Value = row.jo_no_rec
                rowDGV.Cells(Me.ColRMReq_MODate.Index).Value = row.mo_date_rec
                rowDGV.Cells(Me.ColRMReq_MONo.Index).Value = row.mo_no_rec
                rowDGV.Cells(Me.ColRMReq_WONo.Index).Value = row.wo_no_rec
                rowDGV.Cells(Me.ColRMReq_UnitCode.Index).Value = UnitCode
                rowDGV.Cells(Me.ColRMReq_IdentifierJO.Index).Value = row.identifier_jo_rec
                rowDGV.Cells(Me.ColRMReq_IdentifierWO.Index).Value = row.identifier_wo_rec
                rowDGV.Cells(Me.ColRMReq_Process.Index).Value = Me.GetProses(row.primary_proses)
                rowDGV.Cells(Me.ColRMReq_RequiredDate.Index).Value = Format(row.required_date, "dd MMM yyyy")
                If row.is_wip = 1 Then
                    rowDGV.Cells(Me.ColRMReq_IsWIP.Index).Value = Boolean.TrueString
                Else
                    rowDGV.Cells(Me.ColRMReq_IsWIP.Index).Value = Boolean.FalseString
                End If
            Next

            Me.dgvDataRMReq.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetProses(ByVal vPrimaryProses As Double) As String
        Dim rowProses As Core_Module.DseMasterProses.tblmaster_proses1Row
        Try
            For Each rowProses In Me.TBLProses.Rows
                If rowProses.Primary = vPrimaryProses Then
                    Return rowProses.Code
                End If
            Next

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub FrmTrans_MaterialPlanningDashboard_ExplainRMReq_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.chkShowWONo.Checked = True Then
                Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanning, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanning_ShowWONo, "1", Me.UserCode)
            Else
                Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanning, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanning_ShowWONo, "0", Me.UserCode)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_ExplainMOB_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Dummy As String = ""
        Try
            Me.dgvDataRMReq.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.Text = Me.Text + " - " + Me.ObjProduct.GetProductCode(Me.vPrimaryProduct)
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanning, _
                       Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanning_ShowWONo, Me.UserCode)
            If Dummy = "1" Then
                Me.chkShowWONo.Checked = True
                Me.ColRMReq_RequiredDate.Visible = True
                Me.llbSetReqDate.Visible = True
                Me.LoadDGVDetail()
            Else
                Me.chkShowWONo.Checked = False
                Me.LoadDGV_Group_By_MO_JO()
                Me.ColRMReq_RequiredDate.Visible = False
                Me.llbSetReqDate.Visible = False
            End If
            LoadData_CheckChange = True
            If Me.ReadOnly1 = True Then
                Me.llbSetReqDate.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkShowWONo_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowWONo.CheckedChanged
        Try
            If LoadData_CheckChange = True Then
                If Me.chkShowWONo.Checked = True Then
                    Me.LoadDGVDetail()
                    Me.ColRMReq_RequiredDate.Visible = True
                    Me.llbSetReqDate.Visible = True
                Else
                    Me.LoadDGV_Group_By_MO_JO()
                    Me.ColRMReq_RequiredDate.Visible = False
                    Me.llbSetReqDate.Visible = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvData_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvDataRMReq.CellBeginEdit
        Try
            If Me.chkShowWONo.Checked = True And Me.ReadOnly1 = False Then
                If e.ColumnIndex = Me.ColRMReq_TotalQty.Index Or e.ColumnIndex = Me.ColRMReq_SpareQty.Index Then
                    If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_MP_EditSpareQty") = True Then
                        Me.Buffer = Me.dgvDataRMReq.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        MsgBox("You don't have access to edit this data.", MsgBoxStyle.Information, "Information")
                        e.Cancel = True
                    End If
                ElseIf e.ColumnIndex = Me.ColRMReq_RequiredDate.Index And Me.chkShowWONo.Checked = True Then
                    If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_MP_EditSpareQty") = True Then
                        Me.Buffer = Me.dgvDataRMReq.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        MsgBox("You don't have access to edit this data.", MsgBoxStyle.Information, "Information")
                        e.Cancel = True
                    End If
                Else
                    e.Cancel = True
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvData_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDataRMReq.CellEndEdit
        Dim TglMO, TglJadwal As Date
        Dim PrimaryMO As Double = -1
        Dim Dummy As String = ""

        Try
            If IsNothing(Me.dgvDataRMReq.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                Me.dgvDataRMReq.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            Else
                If e.ColumnIndex = Me.ColRMReq_RequiredDate.Index Then
                    'pastikan tanggal
                    If IsDate(Me.dgvDataRMReq.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please specify a date value.", MsgBoxStyle.Information, "Information")
                        Me.dgvDataRMReq.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    'pastikan tanggalnya lebih kecil dari tanggal mo
                    Dummy = Me.dgvDataRMReq.Item(Me.ColRMReq_IdentifierJO.Index, e.RowIndex).Value.ToString
                    PrimaryMO = Dummy.Substring(0, Dummy.IndexOf("-"))
                    TglMO = Me.ObjMO1.GetDate(PrimaryMO).Date
                    TglJadwal = Me.ObjSch.GetStartTime(Me.dgvDataRMReq.Item(Me.ColRMReq_IdentifierJO.Index, e.RowIndex).Value, _
                    Me.dgvDataRMReq.Item(Me.ColRMReq_IdentifierWO.Index, e.RowIndex).Value).Date
                    If TglJadwal <> #1/1/1950# Then
                        If TglJadwal <= TglMO Then
                            TglMO = TglJadwal - TimeSpan.FromDays(1) 'minus 1 hari, karena asumsi rm harus digudang paling lambat 1 hari sebelum produksi dimulai
                        End If
                    End If

                    'jika sudah terjadwal, pastikan tanggal lebih kecil dari tanggal mulai jadwal
                    If CDate(Me.dgvDataRMReq.Item(e.ColumnIndex, e.RowIndex).Value) < TglMO Then
                        MsgBox("Please specify a date value with minimum value: " + Format(TglMO, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.dgvDataRMReq.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    'format ulang tampilannya
                    Me.dgvDataRMReq.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.dgvDataRMReq.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")

                    'simpan ke db
                    Me.ObjMO.UpdateRequiredDate(Me.dgvDataRMReq.Item(Me.ColRMReq_IdentifierJO.Index, e.RowIndex).Value, Me.dgvDataRMReq.Item(Me.ColRMReq_IdentifierWO.Index, e.RowIndex).Value, Me.vPrimaryProduct, _
                    Me.dgvDataRMReq.Item(Me.ColRMReq_RequiredDate.Index, e.RowIndex).Value)

                    'update ke row item pr ob
                    Me.vDT_ItemRMReq.DefaultView.RowFilter = "[primary product] = " + Me.vPrimaryProduct.ToString + _
                                       " AND [identifier jo rec] = '" + Me.dgvDataRMReq.Item(Me.ColRMReq_IdentifierJO.Index, e.RowIndex).Value.ToString + "'" + _
                                       " AND [identifier wo rec] = '" + Me.dgvDataRMReq.Item(Me.ColRMReq_IdentifierWO.Index, e.RowIndex).Value.ToString + "'"
                    If Me.vDT_ItemRMReq.DefaultView.Count > 0 Then
                        RowItemRMReq = Me.vDT_ItemRMReq.DefaultView(0).Row
                        RowItemRMReq.required_date = CDate(Me.dgvDataRMReq.Item(Me.ColRMReq_RequiredDate.Index, e.RowIndex).Value)
                    Else
                        MsgBox("Logic error: row not found. Please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    End If
                ElseIf e.ColumnIndex = Me.ColRMReq_SpareQty.Index Or e.ColumnIndex = Me.ColRMReq_TotalQty.Index Then
                    'pastikan angka
                    If IsNumeric(Me.dgvDataRMReq.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please specify a number.", MsgBoxStyle.Information, "Information")
                        Me.dgvDataRMReq.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    'digit grouping
                    Me.dgvDataRMReq.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvDataRMReq.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec)

                    If e.ColumnIndex = Me.ColRMReq_SpareQty.Index Then
                        'pastikan non negatif
                        If Me.ObjCD.CekNonNegatif(Me.dgvDataRMReq.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Please specify a non negative number.", MsgBoxStyle.Information, "Information")
                            Me.dgvDataRMReq.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If

                        'hitung ulang qty total
                        Me.dgvDataRMReq.Item(Me.ColRMReq_TotalQty.Index, e.RowIndex).Value = CDec(Me.dgvDataRMReq.Item(e.ColumnIndex, e.RowIndex).Value) + _
                        CDec(Me.dgvDataRMReq.Item(Me.ColRMReq_BalanceQty.Index, e.RowIndex).Value)
                    ElseIf e.ColumnIndex = Me.ColRMReq_TotalQty.Index Then
                        'pastikan positif
                        If Me.ObjCD.CekBilanganPositif(Me.dgvDataRMReq.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Please specify a positive number.", MsgBoxStyle.Information, "Information")
                            Me.dgvDataRMReq.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If

                        'pastikan lebih besar atau sama dengan kolom qty
                        If CDec(Me.dgvDataRMReq.Item(Me.ColRMReq_BalanceQty.Index, e.RowIndex).Value) > CDec(Me.dgvDataRMReq.Item(Me.ColRMReq_TotalQty.Index, e.RowIndex).Value) Then
                            MsgBox("Please specify a number larger than " + Me.ObjCD.DigitGrouping(Me.dgvDataRMReq.Item(Me.ColRMReq_BalanceQty.Index, e.RowIndex).Value.ToString, Me.DefPrec) + " " + Me.dgvDataRMReq.Item(Me.ColRMReq_UnitCode.Index, e.RowIndex).Value.ToString + ".", MsgBoxStyle.Information, "Information")
                            Me.dgvDataRMReq.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If

                        'hitung spare
                        Me.dgvDataRMReq.Item(Me.ColRMReq_SpareQty.Index, e.RowIndex).Value = CDec(Me.dgvDataRMReq.Item(e.ColumnIndex, e.RowIndex).Value) - _
                                          CDec(Me.dgvDataRMReq.Item(Me.ColRMReq_BalanceQty.Index, e.RowIndex).Value)
                    End If
                    'simpan ke database
                    Me.ObjMO.UpdateQtySpare(Me.dgvDataRMReq.Item(Me.ColRMReq_IdentifierJO.Index, e.RowIndex).Value, Me.dgvDataRMReq.Item(Me.ColRMReq_IdentifierWO.Index, e.RowIndex).Value, Me.vPrimaryProduct, _
                    Me.dgvDataRMReq.Item(Me.ColRMReq_SpareQty.Index, e.RowIndex).Value)
                    'update ke row item pr ob
                    Me.vDT_ItemRMReq.DefaultView.RowFilter = "[primary product] = " + Me.vPrimaryProduct.ToString + _
                    " AND [identifier jo rec] = '" + Me.dgvDataRMReq.Item(Me.ColRMReq_IdentifierJO.Index, e.RowIndex).Value.ToString + "'" + _
                    " AND [identifier wo rec] = '" + Me.dgvDataRMReq.Item(Me.ColRMReq_IdentifierWO.Index, e.RowIndex).Value.ToString + "'"
                    If Me.vDT_ItemRMReq.DefaultView.Count > 0 Then
                        RowItemRMReq = Me.vDT_ItemRMReq.DefaultView(0).Row
                        RowItemRMReq.MP_Spare_Qty = CDec(Me.dgvDataRMReq.Item(Me.ColRMReq_SpareQty.Index, e.RowIndex).Value)
                    Else
                        MsgBox("Logic error: row not found. Please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    End If
                End If

                Me.AdaPerubahan = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbSetReqDate.LinkClicked
        Dim ij As New FrmTrans_MaterialPlanningDashboard_ExplainRMReq_SetReqDate
        Dim rowAcuan As DseTrans_ManufacturingOrder.DataTable7Row
        Dim ListJONumber As New List(Of String)
        Dim rowDGV As DataGridViewRow

        Try
            If Me.ReadOnly1 = False Then
                ij.vPrimaryRM = Me.vPrimaryProduct
                ij.TBLAcuan.Rows.Clear()
                For Each rowDGV In Me.dgvDataRMReq.Rows
                    If ListJONumber.Contains(rowDGV.Cells(Me.ColRMReq_JONo.Index).Value) = False Then
                        rowAcuan = ij.TBLAcuan.NewDataTable7Row
                        rowAcuan.Identifier_JO = rowDGV.Cells(Me.ColRMReq_IdentifierJO.Index).Value
                        rowAcuan.JO_No = rowDGV.Cells(Me.ColRMReq_JONo.Index).Value
                        ListJONumber.Add(rowAcuan.JO_No)
                        ij.TBLAcuan.Rows.Add(rowAcuan)
                    End If
                Next

                If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                    For Each rowAcuan In ij.TBLAcuan.Rows
                        If rowAcuan.JO_No = ij.cboJONumber.Text Then
                            Me.ObjMO.UpdateRequiredDate(rowAcuan.Identifier_JO, Me.vPrimaryProduct, ij.dtpRequiredDate.Value)

                            'update ke row item pr ob
                            Me.vDT_ItemRMReq.DefaultView.RowFilter = "[primary product] = " + Me.vPrimaryProduct.ToString + _
                            " AND [identifier jo rec] = '" + rowAcuan.Identifier_JO + "'"
                            If Me.vDT_ItemRMReq.DefaultView.Count > 0 Then
                                For Each DV As Data.DataRowView In Me.vDT_ItemRMReq.DefaultView
                                    RowItemRMReq = DV.Row
                                    RowItemRMReq.required_date = ij.dtpRequiredDate.Value.Date
                                Next
                            Else
                                MsgBox("Logic error: row not found. Please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                            End If

                            Me.AdaPerubahan = True

                            Me.LoadDGVDetail()

                            Exit For
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class