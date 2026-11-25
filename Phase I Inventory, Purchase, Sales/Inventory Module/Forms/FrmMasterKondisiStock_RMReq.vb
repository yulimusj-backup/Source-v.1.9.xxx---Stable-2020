Option Explicit On
Option Strict Off

Partial Public Class FrmMasterKondisiStock
    Dim LoadData_CheckChange As Boolean = False
    Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser

    Private Sub UpdateSetting_ChkRMReq()
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

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim Dummy As String = ""
        Try
            Me.dgvDataRMReq.AlternatingRowsDefaultCellStyle.BackColor = Me.BackcolorAll
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanning, _
                       Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanning_ShowWONo, Me.UserCode)
            If Me.chkShowWONo.Checked = True Then
                Me.LoadDGVDetail()
            Else
                Me.LoadDGV_Group_By_MO_JO()
            End If
            Me.dgvDataRMReq.BringToFront()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanDataRMReq(ByVal vTBLRM As DseTransaksiMaterialPlanning2.tbltrans_mo_item_rmDataTable)
        Dim DTARM As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_mo_item_rmTableAdapter
        Dim rowRM As DseTransaksiMaterialPlanning2.tbltrans_mo_item_rmRow
        Dim DTAWIP As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_mo_item_wipTableAdapter
        Dim TBLWIP As New DseTransaksiMaterialPlanning2.tbltrans_mo_item_wipDataTable
        Dim rowWIP As DseTransaksiMaterialPlanning2.tbltrans_mo_item_wipRow

        Try
            DTARM.Connection.ConnectionString = Core_Module.GetConstring
            DTAWIP.Connection.ConnectionString = Core_Module.GetConstring

            'isikan data rm
            DTARM.FillBy_PrimaryProduct(vTBLRM, Me.txtPrimaryCodeNo.Text)
            For Each rowRM In vTBLRM.Rows
                rowRM.is_wip = 0
            Next

            'isikan data wip
            DTAWIP.FillBy_PrimaryProduct(TBLWIP, Me.txtPrimaryCodeNo.Text)
            'pindahkan data wip ke rm
            For Each rowWIP In TBLWIP.Rows
                rowRM = vTBLRM.Newtbltrans_mo_item_rmRow
                With rowRM
                    .Balance_Qty = rowWIP.Qty_Balance
                    .Identifier_JO = rowWIP.Identifier_JO
                    .Identifier_WO = rowWIP.Identifier_WO
                    .is_wip = 1
                    .JO_Number = rowWIP.JO_Number
                    .MO_Date = rowWIP._Date
                    .MO_Number = rowWIP.MO_Number
                    .Primary_MO = rowWIP.Primary_MO
                    .Primary_Process = rowWIP.Primary_Process
                    .Primary_Product = rowWIP.Primary_Product_Code_WIP
                    .WO_Number = rowWIP.WO_Number
                End With
                vTBLRM.Rows.Add(rowRM)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDGV_Group_By_MO_JO()
        Dim DV As Data.DataRowView
        Dim rowDGV As DataGridViewRow
        Dim row, rowBefore As DseTransaksiMaterialPlanning2.tbltrans_mo_item_rmRow
        Dim UnitCode As String = ""
        Dim Dummy As String = ""
        Dim Idx As Integer = 0
        Dim mDefPrec As Integer = 0
        Dim TBLRM As New DseTransaksiMaterialPlanning2.tbltrans_mo_item_rmDataTable

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            Me.dgvDataRMReq.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvDataRMReq.Rows.Clear()
            Me.ColRMReq_WONo.Visible = False
            Me.ColRMReq_Process.Visible = False
            Me.ColRMReq_Process.Visible = False

            UnitCode = Me.lblUnit1.Text

            Me.ColRMReq_MODate.DefaultCellStyle.Format = "dd MMM yyyy"
            mDefPrec = Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text)
            Me.ColRMReq_BalanceQty.DefaultCellStyle.Format = "N" + mDefPrec.ToString

            Me.IsikanDataRMReq(TBLRM)
            TBLRM.DefaultView.Sort = "[MO Date], [MO Number], [JO Number], [WO Number]"

            For Each DV In TBLRM.DefaultView
                row = DV.Row

                If Idx = 0 Then
                    Me.dgvDataRMReq.Rows.Add()

                    rowDGV = Me.dgvDataRMReq.Rows(Me.dgvDataRMReq.Rows.Count - 1)
                    rowDGV.Cells(Me.ColRMReq_No.Index).Value = Me.dgvDataRMReq.RowCount
                    rowDGV.Cells(Me.ColRMReq_BalanceQty.Index).Value = row.Balance_Qty
                    rowDGV.Cells(Me.ColRMReq_JONo.Index).Value = row.JO_Number
                    rowDGV.Cells(Me.ColRMReq_MODate.Index).Value = row.MO_Date
                    rowDGV.Cells(Me.ColRMReq_MONo.Index).Value = row.MO_Number
                    rowDGV.Cells(Me.ColRMReq_WONo.Index).Value = row.WO_Number
                    rowDGV.Cells(Me.ColRMReq_UnitCode.Index).Value = UnitCode
                    rowDGV.Cells(Me.ColRMReq_IdentifierJO.Index).Value = row.Identifier_JO
                    rowDGV.Cells(Me.ColRMReq_IdentifierWO.Index).Value = row.Identifier_WO
                Else
                    rowBefore = TBLRM.DefaultView(Idx - 1).Row
                    If rowBefore.JO_Number = row.JO_Number And rowBefore.MO_Number = row.MO_Number Then
                        'hanya tambahkan qty saja
                        rowDGV = Me.dgvDataRMReq.Rows(Me.dgvDataRMReq.RowCount - 1)
                        rowDGV.Cells(Me.ColRMReq_BalanceQty.Index).Value = CDec(rowDGV.Cells(Me.ColRMReq_BalanceQty.Index).Value) + row.Balance_Qty
                    Else
                        Me.dgvDataRMReq.Rows.Add()
                        rowDGV = Me.dgvDataRMReq.Rows(Me.dgvDataRMReq.Rows.Count - 1)
                        rowDGV.Cells(Me.ColRMReq_No.Index).Value = Me.dgvDataRMReq.RowCount
                        rowDGV.Cells(Me.ColRMReq_BalanceQty.Index).Value = row.Balance_Qty
                        rowDGV.Cells(Me.ColRMReq_JONo.Index).Value = row.JO_Number
                        rowDGV.Cells(Me.ColRMReq_MODate.Index).Value = row.MO_Date
                        rowDGV.Cells(Me.ColRMReq_MONo.Index).Value = row.MO_Number
                        rowDGV.Cells(Me.ColRMReq_WONo.Index).Value = row.WO_Number
                        rowDGV.Cells(Me.ColRMReq_UnitCode.Index).Value = UnitCode
                        rowDGV.Cells(Me.ColRMReq_IdentifierJO.Index).Value = row.Identifier_JO
                        rowDGV.Cells(Me.ColRMReq_IdentifierWO.Index).Value = row.Identifier_WO
                    End If
                End If

                Idx += 1
            Next

            Me.dgvDataRMReq.AlternatingRowsDefaultCellStyle.BackColor = Me.BackcolorAll
            Me.dgvDataRMReq.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDGVDetail()
        Dim DV As Data.DataRowView
        Dim rowDGV As DataGridViewRow
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_mo_item_rmRow
        Dim UnitCode As String = ""
        Dim mDefPrec As Integer = 0
        Dim TBLRM As New DseTransaksiMaterialPlanning2.tbltrans_mo_item_rmDataTable

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.dgvDataRMReq.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvDataRMReq.Rows.Clear()
            Me.ColRMReq_WONo.Visible = True
            Me.ColRMReq_Process.Visible = True
            Me.ColRMReq_IsWIP.Visible = True
            UnitCode = Me.lblUnit1.Text
            Me.ObjProses.FillTabelProses(Me.TBLProses)

            Me.ColRMReq_MODate.DefaultCellStyle.Format = "dd MMM yyyy"
            mDefPrec = Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text)
            Me.ColRMReq_BalanceQty.DefaultCellStyle.Format = "N" + mDefPrec.ToString

            Me.IsikanDataRMReq(TBLRM)
            TBLRM.DefaultView.Sort = "[MO Date], [MO Number], [JO Number], [WO Number]"

            For Each DV In TBLRM.DefaultView
                row = DV.Row
                Me.dgvDataRMReq.Rows.Add()

                rowDGV = Me.dgvDataRMReq.Rows(Me.dgvDataRMReq.Rows.Count - 1)
                rowDGV.Cells(Me.ColRMReq_No.Index).Value = Me.dgvDataRMReq.RowCount
                rowDGV.Cells(Me.ColRMReq_BalanceQty.Index).Value = row.Balance_Qty
                rowDGV.Cells(Me.ColRMReq_JONo.Index).Value = row.JO_Number
                rowDGV.Cells(Me.ColRMReq_MODate.Index).Value = row.MO_Date
                rowDGV.Cells(Me.ColRMReq_MONo.Index).Value = row.MO_Number
                rowDGV.Cells(Me.ColRMReq_WONo.Index).Value = row.WO_Number
                rowDGV.Cells(Me.ColRMReq_UnitCode.Index).Value = UnitCode
                rowDGV.Cells(Me.ColRMReq_IdentifierJO.Index).Value = row.Identifier_JO
                rowDGV.Cells(Me.ColRMReq_IdentifierWO.Index).Value = row.Identifier_WO
                rowDGV.Cells(Me.ColRMReq_Process.Index).Value = Me.GetProses(row.Primary_Process)
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

    Private Sub chkShowWONo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowWONo.CheckedChanged
        Try
            If LoadData_CheckChange = True Then
                If Me.chkShowWONo.Checked = True Then
                    Me.LoadDGVDetail()
                Else
                    Me.LoadDGV_Group_By_MO_JO()
                End If
                Me.dgvDataRMReq.BringToFront()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
