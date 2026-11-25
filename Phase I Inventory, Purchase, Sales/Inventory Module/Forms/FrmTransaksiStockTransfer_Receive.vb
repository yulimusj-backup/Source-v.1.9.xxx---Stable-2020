Option Explicit On
Option Strict Off

Public Class FrmTransaksiStockTransfer_Receive
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = ""

    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjWarehouse As New Core_Module.clsMasterWarehouse
    Dim objLocation As New Core_Module.clsMasterStorageArea
    Dim ObjStockTransfer As New clsTransaksiStockTransfer
    Dim ObjStockTransferItem As New clsTransaksiStockTransferItem
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjCL As New ClsTransaksiMonthlyClosingItem

    Dim Buffer As String = ""
    Dim DataGridViewToDelete As DataGridView = Nothing
    Dim DataGridViewRowToDelete As Integer = -1
    Dim ApakahPunyaFullAkses As Boolean = False

    Private Sub GetHakAkses()
        Try
            'jika user tidak memiliki full akses terhadap form receive maka user hanya bisa buat dan hapus transaksi yang dibuat olehnya saja.
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockTransfer_Receive_FullAkses") = False Then
                Me.ApakahPunyaFullAkses = False
            Else
                ApakahPunyaFullAkses = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Dim PrimaryUser As Double = -1
    Enum EnumWHLocation As Integer
        Warehouse
        Location
    End Enum

    Sub GetFilterWarehouse()
        Dim dta As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransfer_setting_warehouseTableAdapter
        Dim tbl As New dseTransaksiStockTransfer.tbltransaksi_stocktransfer_setting_warehouseDataTable
        Dim row As dseTransaksiStockTransfer.tbltransaksi_stocktransfer_setting_warehouseRow
        Dim dtaWH As New dseTransaksiStockTransferTableAdapters.tblmaster_warehouseTableAdapter
        Dim tblWH As New dseTransaksiStockTransfer.tblmaster_warehouseDataTable
        Dim rowWH As dseTransaksiStockTransfer.tblmaster_warehouseRow
        Dim rd As DataGridViewRow

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dtaWH.Connection.ConnectionString = Core_Module.GetConstring
            dta.FillBy_PrimaryUser(tbl, Me.PrimaryUser)
            dtaWH.Fill(tblWH)

            'level=0 (warehouse)
            'level=1(Location)
            tbl.DefaultView.RowFilter = "[level] = 0"
            tbl.DefaultView.Sort = "[primary acuan]"
            tblWH.DefaultView.RowFilter = ""
            tblWH.DefaultView.Sort = "[primary]"
            'get warehouse code + desc
            For i As Integer = 0 To tbl.DefaultView.Count - 1
                row = tbl.DefaultView(i).Row
                With row
                    While tblWH.DefaultView.Count > 0
                        rowWH = tblWH.DefaultView(0).Row
                        If .primary_acuan < rowWH.Primary Then
                            Exit While
                        ElseIf .primary_acuan = rowWH.Primary Then
                            .location_code = rowWH.Warehouse_Code
                            .location_desc = rowWH.Description
                            Exit While
                        ElseIf .primary_acuan > rowWH.Primary Then
                            tblWH.Rows.Remove(rowWH)
                        End If
                    End While
                End With
            Next

            'add ke dgv
            Me.dgvFilter.AllowUserToAddRows = False
            Me.dgvFilter.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For i As Integer = 0 To tbl.Rows.Count - 1
                row = tbl.Rows(i)
                With row
                    rd = Me.dgvFilter.Rows(Me.dgvFilter.Rows.Add)
                    rd.Cells(Me.ColFilter_Primary.Index).Value = .primary_acuan
                    rd.Cells(Me.ColFilter_Level.Index).Value = Me.GetWarehouseString(.level)
                    rd.Cells(Me.ColFilter_Code.Index).Value = .location_code
                    rd.Cells(Me.ColFilter_Desc.Index).Value = .location_desc
                End With
            Next
            Me.dgvFilter.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            If Me.dgvFilter.AllowUserToAddRows = False Then
                Me.dgvFilter.AllowUserToAddRows = True
            End If
        End Try
    End Sub

    Sub SetFilterWarehouse()
        Dim dta As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransfer_setting_warehouseTableAdapter
        Dim Level As Integer = 0
        Dim PrimaryAcuan As Double = -1

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Me.dgvFilter.AllowUserToAddRows = False
            dta.Delete(Me.PrimaryUser)
            For Each rd As DataGridViewRow In Me.dgvFilter.Rows
                Level = Me.GetWarehouseEnum(rd.Cells(Me.ColFilter_Level.Index).Value)
                If String.IsNullOrEmpty(rd.Cells(Me.ColFilter_Primary.Index).Value) = False Then
                    If rd.Cells(Me.ColFilter_Primary.Index).Value <> -1 Then
                        PrimaryAcuan = CDbl(rd.Cells(Me.ColFilter_Primary.Index).Value)
                        dta.Insert(Me.PrimaryUser, Level, PrimaryAcuan)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            If Me.dgvFilter.AllowUserToAddRows = False Then
                Me.dgvFilter.AllowUserToAddRows = True
            End If
        End Try
    End Sub

    Private Function GetWarehouseString(ByVal vEnum As EnumWHLocation) As String

        Try
            If vEnum = EnumWHLocation.Warehouse Then
                Return "Warehouse"
            Else
                Return "Location"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetWarehouseEnum(ByVal vCode As String) As EnumWHLocation

        Try
            If vCode = "Warehouse" Then
                Return EnumWHLocation.Warehouse
            Else
                Return EnumWHLocation.Location
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub FrmTransaksiStockTransfer_Receive_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Try
            Me.SetFilterWarehouse()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiStockTransfer_Receive_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.PrimaryUser = Me.ObjUser.GetPrimary(Me.UserCode)
            Me.GetFilterWarehouse()

            Me.dgvFilter.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvHeader.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.dgvHeader.ReadOnly = False
            Me.GetHakAkses()

            If Me.ApakahPunyaFullAkses = False Then
                Me.BtnReceived.Enabled = False
                Me.BtnReceivedCancel.Enabled = False
                Me.ColModify.Visible = False
                Me.ColReceived.Visible = False
                Me.ColCancelReceived.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvFilter_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvFilter.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColFilter_Desc.Index Then
                e.Cancel = True
                Exit Sub
            Else
                ''tak boleh ganti level jika sudah ada code
                If e.ColumnIndex = Me.ColFilter_Level.Index Then
                    If Not String.IsNullOrEmpty(Me.dgvFilter.Item(Me.ColFilter_Code.Index, e.RowIndex).Value) Then
                        MsgBox("Cannot change level if code is defined.", MsgBoxStyle.Information, "Information")
                        e.Cancel = True
                        Exit Sub
                    End If
                End If

                ''tak boleh edit yang lain jika code belum diisi
                If e.ColumnIndex <> Me.ColFilter_Code.Index And e.ColumnIndex <> Me.ColFilter_Level.Index Then
                    If String.IsNullOrEmpty(Me.dgvFilter.Item(Me.ColFilter_Code.Index, e.RowIndex).Value) Then
                        MsgBox("Input code first.", MsgBoxStyle.Information, "Informatio")
                        e.Cancel = True
                        Exit Sub
                    End If
                End If

                If e.RowIndex = (Me.dgvFilter.RowCount - 1) Then
                    If e.RowIndex > 0 Then
                        If String.IsNullOrEmpty(Me.dgvFilter.Item(Me.ColFilter_Code.Index, e.RowIndex - 1).Value) = True Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If
                End If

                Me.Buffer = Me.dgvFilter.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub dgvFilter_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvFilter.CellEndEdit
        Try
            If Not IsNothing(Me.dgvFilter.Item(e.ColumnIndex, e.RowIndex).Value) Then
                If Not String.IsNullOrEmpty(Me.dgvFilter.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim) Then
                    If e.ColumnIndex = Me.ColFilter_Level.Index Then

                    End If
                    If e.ColumnIndex = Me.ColFilter_Code.Index Then
                        Me.dgvFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvFilter.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.Replace("'", "").Replace("*", "").Trim.ToUpper
                        If Not IsNothing(Me.Buffer) Then
                            If Me.dgvFilter.Item(e.ColumnIndex, e.RowIndex).Value.ToString = Me.Buffer.ToString Then
                                Exit Sub
                            End If
                        End If
                        Me.CellEndEdit_Code(sender, e)
                    End If
                Else
                    Me.DataGridViewRowDelete(sender, e, Me.dgvFilter)
                End If
            Else
                Me.DataGridViewRowDelete(sender, e, Me.dgvFilter)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Function IsExistWarehouseLocation(ByVal vWarehouse As Boolean, ByVal vCode As String, ByVal vCurrentRow As Integer) As Boolean
        Try
            Me.dgvFilter.AllowUserToAddRows = False
            For Each DataGridViewRow_ As DataGridViewRow In Me.dgvFilter.Rows
                If DataGridViewRow_.Index <> vCurrentRow Then
                    If vWarehouse Then
                        If DataGridViewRow_.Cells(Me.ColFilter_Level.Index).Value = "Warehouse" And _
                                DataGridViewRow_.Cells(Me.ColFilter_Code.Index).Value = vCode Then
                            Return True
                        End If
                    Else
                        If DataGridViewRow_.Cells(Me.ColFilter_Level.Index).Value = "Location" And _
                                DataGridViewRow_.Cells(Me.ColFilter_Code.Index).Value = vCode Then
                            Return True
                        End If
                    End If
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            Me.dgvFilter.AllowUserToAddRows = True
        End Try
    End Function

    Sub CellEndEdit_Code(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim frmFindWarehouse As New Core_Module.frmCariMasterWarehouse
        Dim frmFindLocation As New Core_Module.FrmCariMasterStorageArea
        Try
            If Me.dgvFilter.Item(Me.ColFilter_Level.Index, e.RowIndex).Value = "Warehouse" Then
                If Me.ObjWarehouse.ApakahSudahAda(Me.dgvFilter.Item(Me.ColFilter_Code.Index, e.RowIndex).Value) = False Then
                    frmFindWarehouse.txtCari.Text = Me.dgvFilter.Item(Me.ColFilter_Code.Index, e.RowIndex).Value
                    frmFindWarehouse.ShowDialog()
                    If frmFindWarehouse.PrimaryWarehouse <> -1 Then
                        Me.dgvFilter(e.ColumnIndex, e.RowIndex).Value = Me.ObjWarehouse.GetWarehouseCode(frmFindWarehouse.PrimaryWarehouse)
                    Else
                        Me.DataGridViewRowDelete(sender, e, Me.dgvFilter)
                        Exit Sub
                    End If
                End If
                If Not String.IsNullOrEmpty(Me.dgvFilter.Item(e.ColumnIndex, e.RowIndex).Value) Then
                    If Me.IsExistWarehouseLocation(True, Me.dgvFilter.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex) Then
                        MsgBox(Me.dgvFilter.Item(Me.ColFilter_Level.Index, e.RowIndex).Value & " " & Me.dgvFilter.Item(e.ColumnIndex, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                        Me.DataGridViewRowDelete(sender, e, Me.dgvFilter)
                        Exit Sub
                    End If
                    Me.dgvFilter.Item(Me.ColFilter_Primary.Index, e.RowIndex).Value = Me.ObjWarehouse.GetPrimary(Me.dgvFilter.Item(e.ColumnIndex, e.RowIndex).Value)
                    Me.dgvFilter.Item(Me.ColFilter_Desc.Index, e.RowIndex).Value = Me.ObjWarehouse.GetDescription(Me.dgvFilter.Item(Me.ColFilter_Primary.Index, e.RowIndex).Value)
                End If
            ElseIf Me.dgvFilter.Item(Me.ColFilter_Level.Index, e.RowIndex).Value = "Location" Then
                If Me.objLocation.ApakahSudahAda(Me.dgvFilter.Item(Me.ColFilter_Code.Index, e.RowIndex).Value, Core_Module.clsMasterStorageArea.enumTransit.NotTransit) = False Then
                    frmFindLocation.txtCari.Text = Me.dgvFilter.Item(Me.ColFilter_Code.Index, e.RowIndex).Value
                    frmFindLocation.PencarianTransit = Core_Module.clsMasterStorageArea.enumTransit.NotTransit
                    frmFindLocation.ShowDialog()
                    If frmFindLocation.PrimaryStorageArea <> -1 Then
                        Me.dgvFilter(e.ColumnIndex, e.RowIndex).Value = Me.objLocation.GetStorageAreaCode(frmFindLocation.PrimaryStorageArea)
                    Else
                        Me.DataGridViewRowDelete(sender, e, Me.dgvFilter)
                        Exit Sub
                    End If
                End If
                If Not String.IsNullOrEmpty(Me.dgvFilter.Item(e.ColumnIndex, e.RowIndex).Value) Then
                    If Me.IsExistWarehouseLocation(False, Me.dgvFilter.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex) Then
                        MsgBox(Me.dgvFilter.Item(Me.ColFilter_Level.Index, e.RowIndex).Value & " " & Me.dgvFilter.Item(e.ColumnIndex, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                        Me.DataGridViewRowDelete(sender, e, Me.dgvFilter)
                        Exit Sub
                    End If
                    Me.dgvFilter.Item(Me.ColFilter_Primary.Index, e.RowIndex).Value = Me.objLocation.GetPrimary(Me.dgvFilter.Item(e.ColumnIndex, e.RowIndex).Value)
                    Me.dgvFilter.Item(Me.ColFilter_Desc.Index, e.RowIndex).Value = Me.objLocation.GetDescription(Me.dgvFilter.Item(Me.ColFilter_Primary.Index, e.RowIndex).Value)
                End If
            Else
                MsgBox("Logic Error: unspecified ComboBoxCell text.", MsgBoxStyle.Critical, "Error")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvFilter.AllowUserToAddRows = True
        End Try
    End Sub

    Sub DataGridViewRowDelete(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, ByVal vDataGridView As DataGridView)
        Try
            If Me.Buffer = Nothing Then
                Me.DataGridViewRowToDelete = e.RowIndex
                Me.DataGridViewToDelete = vDataGridView
                Me.BackgroundWorker1.RunWorkerAsync()
            Else
                vDataGridView.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub dgvFilter_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvFilter.DataError

    End Sub

    Private Sub dgvFilter_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvFilter.RowsAdded
        Dim rd As New DataGridViewRow

        Try
            rd = Me.dgvFilter.Rows(Me.dgvFilter.RowCount - 1)
            rd.Cells(Me.ColFilter_Level.Index).Value = "Warehouse"
            rd.Cells(Me.ColFilter_Primary.Index).Value = -1
            rd.Cells(Me.ColFilter_Code.Index).Value = ""
            rd.Cells(Me.ColFilter_Desc.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvFilter_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvFilter.Sorted
        Try
            'Me.ObjCD.SetNumber(Me.dgvFilter)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvFilter_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvFilter.UserDeletedRow
        Try
            'Me.ObjCD.SetNumber(Me.dgvFilter)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHeader_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvHeader.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColModify.Index Then

            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHeader_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvHeader.CellClick

        Try
            'If e.RowIndex = -1 Then
            '    Exit Sub
            'End If
            'If e.ColumnIndex = Me.ColReceivedDate.Index Or e.ColumnIndex = Me.ColReceivedOn.Index Then
            '    If Me.dgvHeader.Item(e.ColumnIndex, e.RowIndex).Value = Nothing Then
            '        Me.dgvHeader.Item(e.ColumnIndex, e.RowIndex).Style.SelectionForeColor = Color.CornflowerBlue
            '    Else
            '        Me.dgvHeader.Item(e.ColumnIndex, e.RowIndex).Style.SelectionForeColor = Me.dgvHeader.Item(Me.ColPrimaryDoc.Index, e.RowIndex).Style.BackColor
            '    End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHeader_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvHeader.CellContentClick
        Dim i As New FrmTransaksiStockTransfer
        Dim hasil As MsgBoxResult
        Dim ListPrimaryReceiveModify As New List(Of Double)

        Try
            If e.ColumnIndex = Me.ColReceived.Index Then
                'If Me.dgvHeader.Item(Me.ColStatus.Index, e.RowIndex).Value = "Received" Then
                '    MsgBox("Can't be received because transaction stock transfer has been received.", MsgBoxStyle.Information, "Information")
                '    Exit Sub
                'End If
                If Me.BatasanReceivedOrCancel(True, ListPrimaryReceiveModify, False, e.RowIndex) = False Then
                    Exit Sub
                End If
                hasil = MsgBox("Received?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    i.UserCode = Me.UserCode
                    i.ApakahTransaksiReceived = True
                    i.PrimaryRefDoc = Me.dgvHeader.Item(Me.ColPrimaryDoc.Index, e.RowIndex).Value
                    i.NoRefDocument = Me.dgvHeader.Item(Me.ColDocumentNo.Index, e.RowIndex).Value

                    i.LblProgress.Text = "Data 1 of 1"  '+ " - " + RowEmp.Emp_Name
                    i.ProgressBar1.Maximum = 1
                    i.ProgressBar1.Value = 1
                    i.ShowDialog()
                    i.Dispose()
                    Me.btnFilter.PerformClick()
                End If
            ElseIf e.ColumnIndex = Me.ColEdit.Index Then
                'hide
            ElseIf e.ColumnIndex = Me.ColCancelReceived.Index Then
                'If Me.dgvHeader.Item(Me.ColStatus.Index, e.RowIndex).Value <> "Received" Then
                '    MsgBox("Can't be cancel received because transaction stock transfer has not been received.", MsgBoxStyle.Information, "Information")
                '    Exit Sub
                'End If
                If Me.BatasanReceivedOrCancel(False, ListPrimaryReceiveModify, False, e.RowIndex) = False Then
                    Exit Sub
                End If
                hasil = MsgBox("Cancel received?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Me.Cursor = Cursors.WaitCursor
                    If i.DeleteTransaksiReceived(Me.dgvHeader.Item(Me.ColPrimaryReceiveDoc.Index, e.RowIndex).Value, _
                    Me.dgvHeader.Item(Me.ColPrimaryDoc.Index, e.RowIndex).Value, Me.UserCode) = True Then
                        MsgBox("Cancel received succeed", MsgBoxStyle.Information, "Information")
                        Me.btnFilter.PerformClick()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub dgvHeader_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvHeader.DataError

    End Sub

    Private Sub dgvHeader_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvHeader.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvHeader)
            For Each rd As DataGridViewRow In Me.dgvHeader.Rows
                If rd.Cells(Me.ColReceivedDate.Index).Value = "1/1/1950" Or String.IsNullOrEmpty(rd.Cells(Me.ColReceivedDate.Index).Value) = True Then
                    If (rd.Index) Mod 2 = 0 Then
                        rd.Cells(Me.ColReceivedDate.Index).Style.ForeColor = Color.White
                        rd.Cells(Me.ColReceivedOn.Index).Style.ForeColor = Color.White
                    Else
                        rd.Cells(Me.ColReceivedDate.Index).Style.ForeColor = Me.BackColorAll
                        rd.Cells(Me.ColReceivedOn.Index).Style.ForeColor = Me.BackColorAll
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbCheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCheckAll.LinkClicked
        Try
            Me.dgvHeader.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rd As DataGridViewRow In Me.dgvHeader.Rows
                rd.Cells(Me.ColModify.Index).Value = True
            Next
            Me.dgvHeader.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUncheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUncheckAll.LinkClicked
        Try
            Me.dgvHeader.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rd As DataGridViewRow In Me.dgvHeader.Rows
                rd.Cells(Me.ColModify.Index).Value = False
            Next
            Me.dgvHeader.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Dim dta As New dseTransaksiStockTransferTableAdapters.tbl_receivedTableAdapter
        Dim tbl As New dseTransaksiStockTransfer.tbl_receivedDataTable
        Dim row As dseTransaksiStockTransfer.tbl_receivedRow
        Dim OnTransit_ As Integer = 1
        Dim Received_ As Integer = 2
        Dim ListPrimaryUnique As New List(Of Double)
        Dim rd As DataGridViewRow

        Try
            Me.Cursor = Cursors.WaitCursor
            Me.dgvHeader.Rows.Clear()
            Me.dgvHeader.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            If Me.dtpFrom.Value.Date > Me.dtpTo.Value.Date Then
                MsgBox("Starting period must be smaller than ending period", MsgBoxStyle.Information, "Information")
                Me.dtpFrom.Focus()
                Exit Sub
            End If

            If Me.chkFilter_OnTransit.Checked = False And Me.chkFilter_Received.Checked = False Then
                MsgBox("Please check on transit or received", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If Me.chkFilter_OnTransit.Checked = False And Me.chkFilter_Received.Checked = True Then
                OnTransit_ = 2
                Received_ = 2
            ElseIf Me.chkFilter_OnTransit.Checked = True And Me.chkFilter_Received.Checked = False Then
                OnTransit_ = 1
                Received_ = 1
            ElseIf Me.chkFilter_OnTransit.Checked = True And Me.chkFilter_Received.Checked = True Then
                OnTransit_ = 1
                Received_ = 2
            Else
                OnTransit_ = 3
                Received_ = 3
            End If

            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.Fill(tbl, OnTransit_, Received_, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date)
            'filter dgv
            tbl.DefaultView.RowFilter = GetFilter()
            For Each dv As Data.DataRowView In tbl.DefaultView
                row = dv.Row
                With row
                    If ListPrimaryUnique.Contains(.Primary) = False Then
                        ListPrimaryUnique.Add(.Primary)
                        Me.dgvHeader.Rows.Add()
                        rd = Me.dgvHeader.Rows(Me.dgvHeader.RowCount - 1)
                        rd.Cells(Me.ColPrimaryDoc.Index).Value = .Primary
                        rd.Cells(Me.ColPrimaryReceiveDoc.Index).Value = .Primary_Received
                        rd.Cells(Me.ColNo.Index).Value = rd.Index + 1
                        rd.Cells(Me.ColDocumentNo.Index).Value = .Document_No
                        rd.Cells(Me.ColDate.Index).Value = Format(._Date, "dd MMM yyyy")
                        rd.Cells(Me.ColRemarks.Index).Value = .Remarks
                        rd.Cells(Me.ColCreatedBy.Index).Value = .Created_By
                        If .Status = clsTransaksiStockTransfer.EnumTransitStatus.No_Transit Then
                            rd.Cells(Me.ColStatus.Index).Value = "No Transit"
                        ElseIf .Status = clsTransaksiStockTransfer.EnumTransitStatus.On_Transit Then
                            rd.Cells(Me.ColStatus.Index).Value = "On Transit"
                        ElseIf .Status = clsTransaksiStockTransfer.EnumTransitStatus.Finished_Transit Then
                            rd.Cells(Me.ColStatus.Index).Value = "Received"
                        End If

                        If .Status = clsTransaksiStockTransfer.EnumTransitStatus.Finished_Transit Then
                            rd.Cells(Me.ColReceivedNo.Index).Value = .Received_No
                            rd.Cells(Me.ColReceivedDate.Index).Value = .Received_Date
                            rd.Cells(Me.ColReceivedBy.Index).Value = .Received_By
                            rd.Cells(Me.ColReceivedOn.Index).Value = .Received_Timestamp
                        Else
                            If (rd.Index) Mod 2 = 0 Then
                                rd.Cells(Me.ColReceivedDate.Index).Style.ForeColor = Color.White
                                rd.Cells(Me.ColReceivedOn.Index).Style.ForeColor = Color.White
                            Else
                                rd.Cells(Me.ColReceivedDate.Index).Style.ForeColor = Me.BackColorAll
                                rd.Cells(Me.ColReceivedOn.Index).Style.ForeColor = Me.BackColorAll
                            End If
                            If rd.Cells(Me.ColReceivedDate.Index).Value = Nothing Then
                                rd.Cells(Me.ColReceivedDate.Index).Style.SelectionForeColor = Color.CornflowerBlue
                                rd.Cells(Me.ColReceivedOn.Index).Style.SelectionForeColor = Color.CornflowerBlue
                            Else
                                rd.Cells(Me.ColReceivedDate.Index).Style.SelectionForeColor = rd.Cells(Me.ColPrimaryDoc.Index).Style.BackColor
                                rd.Cells(Me.ColReceivedOn.Index).Style.SelectionForeColor = rd.Cells(Me.ColPrimaryDoc.Index).Style.BackColor
                            End If
                            'rd.Cells(Me.ColReceivedNo.Index).Value = ""
                            'rd.Cells(Me.ColReceivedDate.Index).Value = ""
                            'rd.Cells(Me.ColReceivedBy.Index).Value = ""
                            'rd.Cells(Me.ColReceivedOn.Index).Value = ""
                        End If

                        rd.Cells(Me.ColModify.Index).Value = False
                        'rd.Cells(Me.ColReceived.Index).Value = ""
                        'rd.Cells(Me.ColEdit.Index).Value = ""
                        'rd.Cells(Me.ColCancelReceived.Index).Value = ""
                    End If
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvHeader.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Function GetFilter() As String
        Dim FilterDgv As String = ""
        Dim apakahfilterrefdoc As Boolean = False
        Dim apakahfilterdgv As Boolean = False

        Try
            If String.IsNullOrEmpty(Me.TxtDocNoFilter.Text) = False Then
                FilterDgv = "[Document No] LIKE '" & Me.TxtDocNoFilter.Text & "%'"
                'FilterDgv = "[Received No] LIKE '" & Me.TxtDocNoFilter.Text & "%'"
                apakahfilterrefdoc = True
            End If

            For Each rd As DataGridViewRow In Me.dgvFilter.Rows
                If rd.Cells(Me.ColFilter_Primary.Index).Value <> -1 Then
                    apakahfilterdgv = True
                    If rd.Cells(Me.ColFilter_Level.Index).Value = "Warehouse" Then
                        If FilterDgv = "" Then
                            FilterDgv = "[Primary Warehouse Transit] = " & CDbl(rd.Cells(Me.ColFilter_Primary.Index).Value)
                        Else
                            If rd.Index = 0 Then
                                FilterDgv &= " AND ([Primary Warehouse Transit] = " & CDbl(rd.Cells(Me.ColFilter_Primary.Index).Value)
                            Else
                                FilterDgv &= " OR [Primary Warehouse Transit] = " & CDbl(rd.Cells(Me.ColFilter_Primary.Index).Value)
                            End If
                        End If
                    ElseIf rd.Cells(Me.ColFilter_Level.Index).Value = "Location" Then
                        If FilterDgv = "" Then
                            FilterDgv = "[primary transit area] = " & CDbl(rd.Cells(Me.ColFilter_Primary.Index).Value)
                        Else
                            If rd.Index = 0 Then
                                FilterDgv &= " AND ([primary transit area] = " & CDbl(rd.Cells(Me.ColFilter_Primary.Index).Value)
                            Else
                                FilterDgv &= " OR [primary transit area] = " & CDbl(rd.Cells(Me.ColFilter_Primary.Index).Value)
                            End If
                        End If
                    End If
                End If
            Next

            If apakahfilterrefdoc = True And apakahfilterdgv = True Then
                FilterDgv &= ")"
            End If
            ''jika user tidak memiliki full akses terhadap form receive maka user hanya bisa buat dan hapus transaksi yang dibuat olehnya saja.
            'If Me.ApakahPunyaFullAkses = False Then
            '    If FilterDgv = "" Then
            '        FilterDgv = "[Created By] = '" & Me.UserCode & "'"
            '    Else
            '        FilterDgv &= " AND [Created By] = '" & Me.UserCode & "'"
            '    End If
            'End If

            Return FilterDgv
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub BtnReceived_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnReceived.Click
        Dim CountModif As Integer = 0
        Dim TotalCountModif As Integer = 0
        Dim AdaPerubahan As Boolean = False
        Dim hasil As MsgBoxResult
        Dim ListPrimaryReceiveModify As New List(Of Double)

        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.BatasanReceivedOrCancel(True, ListPrimaryReceiveModify) = False Then
                'Exit Sub
            End If

            If ListPrimaryReceiveModify.Count > 0 Then
                hasil = MsgBox("Received?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    For Each rd As DataGridViewRow In Me.dgvHeader.Rows
                        If rd.Cells(Me.ColModify.Index).Value = True Then
                            If ListPrimaryReceiveModify.Contains(rd.Cells(Me.ColPrimaryDoc.Index).Value) = True Then
                                Dim i As New FrmTransaksiStockTransfer
                                i.UserCode = Me.UserCode
                                i.ApakahTransaksiReceived = True
                                i.PrimaryRefDoc = rd.Cells(Me.ColPrimaryDoc.Index).Value
                                i.NoRefDocument = rd.Cells(Me.ColDocumentNo.Index).Value
                                CountModif += 1
                                i.LblProgress.Text = "Data " + CountModif.ToString + " of " + ListPrimaryReceiveModify.Count.ToString '+ " - " + RowEmp.Emp_Name
                                i.ProgressBar1.Maximum = ListPrimaryReceiveModify.Count
                                i.ProgressBar1.Value = CountModif
                                If i.ShowDialog() = Windows.Forms.DialogResult.Abort Then
                                    MsgBox("Process stopped by user.", MsgBoxStyle.Information, "Information")
                                    Exit For
                                End If
                                i.Dispose()
                                AdaPerubahan = True
                            End If
                        End If
                    Next

                    If AdaPerubahan = True Then
                        Me.btnFilter.PerformClick()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub BtnReceivedCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnReceivedCancel.Click
        Dim TotalCountModif As Integer = 0
        Dim CountModif As Integer = 0
        Dim hasil As MsgBoxResult
        Dim AdaPerubahan As Boolean = False
        Dim ListPrimaryReceiveModify As New List(Of Double)
        Dim i As New FrmTransaksiStockTransfer

        Try
            Me.Cursor = Cursors.WaitCursor
            'Batasan 
            'tidak bisa cancel jika mengakibatkan minus di location to/destination
            'tidak bisa cancel jika sudah monthly closing (by date stock transfer received)
            If Me.BatasanReceivedOrCancel(False, ListPrimaryReceiveModify) = False Then
                'Exit Sub
            End If

            If ListPrimaryReceiveModify.Count > 0 Then
                hasil = MsgBox("Cancel received?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    For Each rd As DataGridViewRow In Me.dgvHeader.Rows
                        If rd.Cells(Me.ColModify.Index).Value = True Then
                            If ListPrimaryReceiveModify.Contains(rd.Cells(Me.ColPrimaryReceiveDoc.Index).Value) = True Then
                                If i.DeleteTransaksiReceived(rd.Cells(Me.ColPrimaryReceiveDoc.Index).Value, rd.Cells(Me.ColPrimaryDoc.Index).Value, Me.UserCode) = True Then
                                    AdaPerubahan = True
                                End If
                            End If
                        End If
                    Next

                    If AdaPerubahan = True Then
                        MsgBox("Cancel received succeed", MsgBoxStyle.Information, "Information")
                        Me.btnFilter.PerformClick()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function BatasanReceivedOrCancel(ByVal vApakahReceived As Boolean, _
    ByRef ListPrimaryReceiveModify As List(Of Double), Optional ByVal vCekAllRowsDgv As Boolean = True, _
    Optional ByVal vIdxDgv As Integer = -1) As Boolean
        Dim DTA As New dseTransaksiStockTransferTableAdapters.DataTable3TableAdapter
        Dim TBL As New dseTransaksiStockTransfer.DataTable3DataTable

        Dim TblReason As New dseTransaksiStockTransfer.tbl_cancel_receivedDataTable
        Dim ListBatasan As String = ""
        Dim TotalCeklis As Integer = 0
        Dim RowAwal As Integer = 0
        Dim RowAkhir As Integer = 0
        Dim rd As DataGridViewRow
        Dim ApakahCekModify As Boolean = True

        Try
            DTA.Connection.ConnectionString = core_module.getconstring

            If vCekAllRowsDgv = True Then
                RowAwal = 0
                RowAkhir = Me.dgvHeader.RowCount - 1
            Else
                If vIdxDgv <> -1 Then
                    RowAwal = vIdxDgv
                    RowAkhir = vIdxDgv
                    ApakahCekModify = False
                Else
                    Return True
                End If
            End If

            For rIdx As Integer = RowAwal To RowAkhir
                rd = Me.dgvHeader.Rows(rIdx)
                If ApakahCekModify = False Then
                    GoTo skipModify
                End If
                If rd.Cells(Me.ColModify.Index).Value = True Then
skipModify:
                    TotalCeklis += 1
                    If vApakahReceived = True Then
                        'received
                        If rd.Cells(Me.ColStatus.Index).Value = "Received" Then
                            Me.AddTblReason(TblReason, rd.Cells(Me.ColDocumentNo.Index).Value, "Transaction stock transfer has been received.")
                        Else
                            If ListPrimaryReceiveModify.Contains(rd.Cells(Me.ColPrimaryDoc.Index).Value) = False Then
                                ListPrimaryReceiveModify.Add(rd.Cells(Me.ColPrimaryDoc.Index).Value)
                            End If
                        End If
                    Else
                        'cancel received
                        If rd.Cells(Me.ColStatus.Index).Value <> "Received" Then
                            Me.AddTblReason(TblReason, rd.Cells(Me.ColDocumentNo.Index).Value, "Transaction stock transfer has not been received.")
                        Else
                            If Me.ApakahBolehCancelReceive(rd.Cells(Me.ColPrimaryReceiveDoc.Index).Value, rd.Cells(Me.ColDocumentNo.Index).Value, _
                            Me.UserCode, TblReason, DTA, TBL) = True Then
                                If ListPrimaryReceiveModify.Contains(rd.Cells(Me.ColPrimaryReceiveDoc.Index).Value) = False Then
                                    ListPrimaryReceiveModify.Add(rd.Cells(Me.ColPrimaryReceiveDoc.Index).Value)
                                End If
                            End If
                        End If
                    End If
                End If
            Next

            If TotalCeklis = 0 Then
                If Me.dgvHeader.RowCount = 0 Then
                    MsgBox("No data can be processed", MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("Please check minimum 1 document.", MsgBoxStyle.Information, "Information")
                End If

                Return False
            End If

            If TblReason.Count > 0 Then
                Dim frm As New FrmTransaksiStockTransfer_Receive_Msgbox
                frm.BackColorAll = Me.BackColorAll
                frm.TblReason = TblReason
                If vApakahReceived = True Then
                    frm.Text = "The following document list can't be received"
                Else
                    frm.Text = "The following document list can't be cancel received"
                End If
                frm.ShowDialog()
                frm.Dispose()

                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahBolehCancelReceive(ByVal vPrimaryReceivedDoc As Double, ByVal vNoDocument As String, ByVal vUser As String, _
    ByVal vTblReason As dseTransaksiStockTransfer.tbl_cancel_receivedDataTable, _
    ByVal DTA As dseTransaksiStockTransferTableAdapters.DataTable3TableAdapter, _
    ByVal TBL As dseTransaksiStockTransfer.DataTable3DataTable) As Boolean
        Dim row As dseTransaksiStockTransfer.DataTable3Row
        Dim TanggalStockInOut As Date = #1/1/1950#
        Dim ProductCode As String = ""

        Try
            '1. jika ada product dalam stock transfer ini yang mengalami perubahan kontrol
            DTA.Fill(TBL, vPrimaryReceivedDoc)
            If TBL.Rows.Count > 0 Then
                For Each row In TBL.Rows
                    If ProductCode = "" Then
                        ProductCode = row.Product_Code
                    Else
                        ProductCode = ProductCode + ", " + row.Product_Code
                    End If
                Next

                Me.AddTblReason(vTblReason, vNoDocument, "Code no " + ProductCode + " have modified it's control state. You cannot delete this data")
                Return False
            End If

            '2. jika sudah diposting/tanggal transaksi < tanggal kunci terakhir
            TanggalStockInOut = Me.ObjStockTransfer.GetTanggalTransaksi(vPrimaryReceivedDoc)
            If Me.ObjCD.GetLasDateOfTheMonth(Me.ObjCL.GetLastClosingPeriod()) >= TanggalStockInOut Then
                Me.AddTblReason(vTblReason, vNoDocument, "The period " + Format(TanggalStockInOut, "MMM yyyy") + " has been closed, you cannot delete this data.")
                Return False
            End If

            '3. jika menghapus transaksi ini akan mengakibatkan data product menjadi minus
            If Me.ObjStockTransferItem.ApakahHapusMengakibatkanMinus(vPrimaryReceivedDoc) = True Then
                Me.AddTblReason(vTblReason, vNoDocument, "Deleting this data will cause minus in stock quotatity, process cancelled.")
                Return False
            End If

            '4. Jika strict edit nyala maka hanya user yang bersangkutan yang boleh melakukan proses edit data
            'If Me.ApakahPunyaFullAkses = False Then
            Dim CreatedBy As String = Me.ObjStockTransfer.GetCreatedBy(vPrimaryReceivedDoc)
            If vUser <> CreatedBy Then
                Me.AddTblReason(vTblReason, vNoDocument, "This transaction can only be delete by: " + CreatedBy)
                Return False
            End If
            'End If

            '5. tidak bisa receive jika sudah dibuat transaksi receive
            Dim vPrimaryStockTransfer As Double = -1
            If Me.ObjStockTransfer.ApakahSudahDibuatReceive(vPrimaryReceivedDoc, vPrimaryStockTransfer) = True Then
                Me.AddTblReason(vTblReason, vNoDocument, "Transaction stock transfer has been receceived")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub AddTblReason(ByRef vTbl As dseTransaksiStockTransfer.tbl_cancel_receivedDataTable, _
    ByVal vNoDoc As String, ByVal vReason As String)
        Dim row As dseTransaksiStockTransfer.tbl_cancel_receivedRow

        Try
            row = vTbl.Newtbl_cancel_receivedRow
            row.No_Doc = vNoDoc
            row.Reason = vReason
            vTbl.Rows.Add(row)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDocNoFilter_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDocNoFilter.GotFocus
        Me.TxtDocNoFilter.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtDocNoFilter_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDocNoFilter.Leave
        Try
            Me.TxtDocNoFilter.Text = Me.TxtDocNoFilter.Text.ToString.Replace("'", "").Replace("*", "").Trim
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDocNoFilter_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDocNoFilter.LostFocus
        Me.TxtDocNoFilter.BackColor = Color.White
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Threading.Thread.Sleep(10)
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.DataGridViewRowToDelete > -1 And Me.DataGridViewRowToDelete < Me.DataGridViewToDelete.Rows.Count - 1 Then
                Me.DataGridViewToDelete.Rows.RemoveAt(Me.DataGridViewRowToDelete)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class