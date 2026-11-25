Option Explicit On
Option Strict Off

Imports Core_Module

Partial Public Class FrmSettingAP_AR
    Dim tblUser As New Core_Module.DseUserLocationPrivilegeSetting.tbl_setting_user_location_privilegeDataTable
    Dim rowUser As Core_Module.DseUserLocationPrivilegeSetting.tbl_setting_user_location_privilegeRow
    Dim tblLocation As New Core_Module.DseUserLocationPrivilegeSetting.tbl_setting_user_location_privilege_itemDataTable
    Dim rowLocation As Core_Module.DseUserLocationPrivilegeSetting.tbl_setting_user_location_privilege_itemRow
    Dim UserCodeDeleted As String
    Dim PrimaryRefPartyDeleted As Double
    Dim LevelDeleted As Integer
    Dim DataGridViewToDelete As DataGridView = Nothing
    Dim DataGridViewRowToDelete As Integer = -1
    Dim objLocation As New Core_Module.clsMasterStorageArea
    Dim objWarehouse As New Core_Module.clsMasterWarehouse

    Public Enum EnumTipeTransaksiInOut As Integer
        STOCK_IN_ADJUST '
        STOCK_IN_PURCAHSE '
        STOCK_IN_SALES_RETURN '
        STOCK_TRANSFER_IN
        STOCK_BUNDLING_AFTER
        STOCK_OUT_ADJUST
        STOCK_OUT_SALES '
        STOCK_OUT_PURCHASE_RETURN '
        STOCK_TANSFER_OUT
        STOCK_BUNDLING_BEFORE
    End Enum

    Private Sub SalinKeField_LocationAccess()
        Dim DefaultUserNotIncludeInList As String = ""
        Try
            Core_Module.SetConstring(Core_Module.GetConstring)

            Me.dgvUser.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvLocation.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            DefaultUserNotIncludeInList = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Location_Privileges, ClsTransaksiSetting.EnumNamaParameter.Default_User_Not_IncludeInList)
            If DefaultUserNotIncludeInList = "" Then
                DefaultUserNotIncludeInList = 0
                Me.ObjSetting.SetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Location_Privileges, ClsTransaksiSetting.EnumNamaParameter.Default_User_Not_IncludeInList, DefaultUserNotIncludeInList)
            End If
            Me.cboDefaultUserNotInList.SelectedIndex = CInt(DefaultUserNotIncludeInList)
            Me.tblLocation.Columns.Add("code", System.Type.GetType("System.String"))
            Me.tblLocation.Columns("code").DefaultValue = ""
            Me.tblLocation.Columns.Add("desc", System.Type.GetType("System.String"))
            Me.tblLocation.Columns("desc").DefaultValue = ""
            Me.LoadData()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub LoadData()
        Dim dataGridViewRow_ As DataGridViewRow

        Dim dtaWarehouse As New DseUserLocationPrivilegeSettingTableAdapters.DataTable1TableAdapter
        Dim tblWarehouse As New DseUserLocationPrivilegeSetting.DataTable1DataTable
        Dim rowWarehouse As DseUserLocationPrivilegeSetting.DataTable1Row

        Dim dtaStorageArea As New DseUserLocationPrivilegeSettingTableAdapters.DataTable2TableAdapter
        Dim tblStorageArea As New DseUserLocationPrivilegeSetting.DataTable2DataTable
        Dim rowStorageArea As DseUserLocationPrivilegeSetting.DataTable2Row

        Dim dtaUser As New DseUserLocationPrivilegeSettingTableAdapters.tbl_setting_user_location_privilegeTableAdapter
        Dim dtaLocation As New DseUserLocationPrivilegeSettingTableAdapters.tbl_setting_user_location_privilege_itemTableAdapter

        Try
            'dtaWarehouse.Connection.ConnectionString = Core_Module.GetConstring
            'dtaStorageArea.Connection.ConnectionString = Core_Module.GetConstring
            'dtaUser.Connection.ConnectionString = Core_Module.GetConstring
            'dtaLocation.Connection.ConnectionString = Core_Module.GetConstring
            dtaWarehouse.Fill(tblWarehouse)
            dtaStorageArea.Fill(tblStorageArea)
            dtaUser.Fill(Me.tblUser)
            dtaLocation.Fill(Me.tblLocation)
            For Each rowLocation In tblLocation
                If rowLocation.location_level = 0 Then
                    tblWarehouse.DefaultView.RowFilter = "[Primary] = " & rowLocation.primary_location.ToString
                    If tblWarehouse.DefaultView.Count > 0 Then
                        rowWarehouse = tblWarehouse.DefaultView(0).Row
                        rowLocation("code") = rowWarehouse.Warehouse_Code
                        rowLocation("desc") = rowWarehouse.Description
                    End If
                Else
                    tblStorageArea.DefaultView.RowFilter = "[Primary] = " & rowLocation.primary_location.ToString
                    If tblStorageArea.DefaultView.Count > 0 Then
                        rowStorageArea = tblStorageArea.DefaultView(0).Row
                        rowLocation("code") = rowStorageArea.Storage_Area_Code
                        rowLocation("desc") = rowStorageArea.Description
                    End If
                End If
            Next
            Me.dgvUser.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvUser.AllowUserToAddRows = False
            For Each rowUser In tblUser
                dataGridViewRow_ = Me.dgvUser.Rows(Me.dgvUser.Rows.Add())
                dataGridViewRow_.Cells(Me.colUser_Code.Index).Value = rowUser.usercode
                If rowUser.default_access = 0 Then
                    dataGridViewRow_.Cells(Me.colUser_DefaultSetting.Index).Value = "Can Access Location"
                Else
                    dataGridViewRow_.Cells(Me.colUser_DefaultSetting.Index).Value = "Cannot Access Location"
                End If
                dataGridViewRow_.Cells(Me.colUser_Name.Index).Value = Me.Objuser.GetNama(rowUser.usercode)
                dataGridViewRow_.Cells(Me.colUser_No.Index).Value = Me.dgvUser.RowCount
                dataGridViewRow_.Cells(Me.ColUser_Primary.Index).Value = Me.Objuser.GetPrimary(rowUser.usercode)
            Next
            If Me.dgvUser.Rows.Count > 0 Then
                Me.dgvUser.ClearSelection()
                Me.dgvUser.Rows(0).Cells(Me.colUser_Code.Index).Selected = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvUser.AllowUserToAddRows = True
            Me.dgvUser.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub dgvUser_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvUser.CellBeginEdit
        Try
            If e.ColumnIndex = Me.colUser_Code.Index Or e.ColumnIndex = Me.colUser_DefaultSetting.Index Then
                Dim i As Integer = Me.GetInvalidDataRowIndex(Me.dgvUser, e.RowIndex, Me.colUser_Code.Index)
                If i <> -1 Then
                    MsgBox("Please complete data at row " & i + 1 & ".", MsgBoxStyle.Information, "Information")
                    e.Cancel = True
                    Exit Sub
                End If
                If e.ColumnIndex = Me.colUser_DefaultSetting.Index Then
                    If String.IsNullOrEmpty(Me.dgvUser.Item(Me.colUser_Code.Index, e.RowIndex).Value) Then
                        MsgBox("Input code first.", MsgBoxStyle.Information, "Information")
                        e.Cancel = True
                        Exit Sub
                    End If
                End If
                Me.Buffer = Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub dgvUser_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvUser.CellEndEdit
        Dim frmFindUser As New Core_Module.FrmCariMasterUser
        Try
            If Not IsNothing(Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value) Then
                If Not String.IsNullOrEmpty(Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim) Then
                    If e.ColumnIndex = Me.colUser_Code.Index Then
                        Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.Replace("'", "").Replace("*", "").Trim.ToUpper
                        If Not IsNothing(Me.Buffer) Then
                            If Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value.ToString = Me.Buffer.ToString Then
                                Exit Sub
                            End If
                        End If
                        If Me.ObjUser.ApakahSudahAda(Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = False Then
                            frmFindUser.TxtCari.Text = Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value.ToString
                            frmFindUser.ShowDialog()
                            If Not String.IsNullOrEmpty(frmFindUser.Kodeuser) Then
                                Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value = frmFindUser.Kodeuser.ToUpper
                            Else
                                Me.DataGridViewRowDelete(sender, e, Me.dgvUser)
                                Exit Sub
                            End If
                        End If
                        If Not String.IsNullOrEmpty(Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value) Then
                            If Me.IsExistUser(Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex) Then
                                MsgBox("User " & Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                                Me.DataGridViewRowDelete(sender, e, Me.dgvUser)
                                Exit Sub
                            End If
                            Me.dgvUser.Item(Me.colUser_Name.Index, e.RowIndex).Value = Me.ObjUser.GetNama(Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value)
                            Me.dgvUser.Item(Me.ColUser_Primary.Index, e.RowIndex).Value = Me.ObjUser.GetPrimary(Me.dgvUser.Item(Me.colUser_Code.Index, e.RowIndex).Value)
                            ''==
                            For i As Integer = Me.tblUser.Rows.Count - 1 To 0 Step -1
                                Me.rowUser = Me.tblUser.Rows(i)
                                If Me.rowUser.usercode = Me.Buffer.ToString Then
                                    Me.tblUser.Rows.RemoveAt(i)
                                End If
                            Next
                            Me.rowUser = Me.tblUser.Newtbl_setting_user_location_privilegeRow
                            Me.rowUser.primary = Me.dgvUser.Item(Me.ColUser_Primary.Index, e.RowIndex).Value
                            Me.rowUser.usercode = Me.dgvUser.Item(Me.colUser_Code.Index, e.RowIndex).Value
                            If Me.dgvUser.Item(Me.colUser_DefaultSetting.Index, e.RowIndex).Value = "Can Access Location" Then
                                Me.rowUser.default_access = 0
                            ElseIf Me.dgvUser.Item(Me.colUser_DefaultSetting.Index, e.RowIndex).Value = "Cannot Access Location" Then
                                Me.rowUser.default_access = 1
                            Else
                                Me.rowUser.default_access = 0
                                MsgBox("Logic Error: unspecified ComboBoxCell text.", MsgBoxStyle.Critical, "Error")
                            End If
                            Me.tblUser.Rows.Add(Me.rowUser)
                            For i As Integer = Me.tblLocation.Rows.Count - 1 To 0 Step -1
                                Me.rowLocation = Me.tblLocation.Rows(i)
                                If Me.rowLocation.usercode.ToLower = Me.Buffer.ToString.ToLower Then
                                    Me.rowLocation.usercode = Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value.ToString
                                End If
                            Next
                            ''==
                        End If
                    End If
                    If e.ColumnIndex = Me.colUser_DefaultSetting.Index Then
                        ''==
                        Me.tblUser.DefaultView.RowFilter = "[usercode] = '" & Me.dgvUser.Item(Me.colUser_Code.Index, e.RowIndex).Value & "'"
                        If Me.tblUser.DefaultView.Count > 0 Then
                            Me.rowUser = Me.tblUser.DefaultView(0).Row
                            If Me.dgvUser.Item(Me.colUser_DefaultSetting.Index, e.RowIndex).Value = "Can Access Location" Then
                                Me.rowUser.default_access = 0
                            ElseIf Me.dgvUser.Item(Me.colUser_DefaultSetting.Index, e.RowIndex).Value = "Cannot Access Location" Then
                                Me.rowUser.default_access = 1
                            Else
                                Me.rowUser.default_access = 0
                                MsgBox("Logic Error: unspecified ComboBoxCell text.", MsgBoxStyle.Critical, "Error")
                            End If
                        End If
                        ''==
                    End If
                Else
                    Me.DataGridViewRowDelete(sender, e, Me.dgvUser)
                End If
            Else
                Me.DataGridViewRowDelete(sender, e, Me.dgvUser)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub dgvUser_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvUser.DataError

    End Sub

    Private Sub dgvUser_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvUser.RowsAdded
        Dim DataGridViewRow_ As DataGridViewRow
        Try
            If Me.dgvUser.AllowUserToAddRows Then
                DataGridViewRow_ = Me.dgvUser.Rows(Me.dgvUser.RowCount - 1)
                DataGridViewRow_.Cells(Me.colUser_Code.Index).Value = String.Empty
                DataGridViewRow_.Cells(Me.colUser_DefaultSetting.Index).Value = "Can Access Location"
                DataGridViewRow_.Cells(Me.colUser_Name.Index).Value = String.Empty
                DataGridViewRow_.Cells(Me.colUser_No.Index).Value = Me.dgvUser.RowCount.ToString
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub dgvUser_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvUser.RowsRemoved
        Try
            Me.objCd.SetNumber(Me.dgvUser)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub dgvUser_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvUser.SelectionChanged
        Dim dataGridViewRow As DataGridViewRow
        Try
            Me.dgvLocation.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvLocation.AllowUserToAddRows = False
            Me.dgvLocation.Rows.Clear()
            If Me.dgvUser.SelectedCells.Count <= 0 Then
                Exit Sub
            End If
            If String.IsNullOrEmpty(Me.dgvUser.Item(Me.colUser_Code.Index, Me.dgvUser.SelectedCells(0).RowIndex).Value) Then
                Exit Sub
            End If
            Me.tblLocation.DefaultView.RowFilter = "[usercode] = '" & Me.dgvUser.Item(Me.colUser_Code.Index, Me.dgvUser.SelectedCells(0).RowIndex).Value.ToString & "'"
            If Me.tblLocation.DefaultView.Count > 0 Then
                For i As Integer = 0 To Me.tblLocation.DefaultView.Count - 1
                    Me.rowLocation = Me.tblLocation.DefaultView(i).Row
                    dataGridViewRow = Me.dgvLocation.Rows(Me.dgvLocation.Rows.Add())
                    With dataGridViewRow
                        .Cells(Me.colLocation_Code.Index).Value = Me.rowLocation("code")
                        .Cells(Me.colLocation_Desc.Index).Value = Me.rowLocation("desc")
                        If Me.rowLocation.location_level = 0 Then
                            .Cells(Me.colLocation_Level.Index).Value = "Warehouse"
                        Else
                            .Cells(Me.colLocation_Level.Index).Value = "Location"
                        End If
                        .Cells(Me.colLocation_No.Index).Value = Me.dgvLocation.Rows.Count
                        .Cells(Me.ColLocation_StockBundlingAfter.Index).Value = Me.rowLocation.stock_bundling_after * -1
                        .Cells(Me.ColLocation_StockBundlingBefore.Index).Value = Me.rowLocation.stock_bundling_before * -1
                        .Cells(Me.ColLocation_StockInAdjust.Index).Value = Me.rowLocation.stock_in_adjust * -1
                        .Cells(Me.colLocation_StockInPurchase.Index).Value = Me.rowLocation.stock_in_purchase * -1
                        .Cells(Me.colLocation_StockInSalesReturn.Index).Value = Me.rowLocation.stock_in_sales_return * -1
                        .Cells(Me.ColLocation_StockOutAdjust.Index).Value = Me.rowLocation.stock_out_adjust * -1
                        .Cells(Me.ColLocation_StockOutPurcahseReturn.Index).Value = Me.rowLocation.stock_out_purchase_return * -1
                        .Cells(Me.ColLocation_StockOutSales.Index).Value = Me.rowLocation.stock_out_sales * -1
                        .Cells(Me.ColLocation_StockTransferIn.Index).Value = Me.rowLocation.stock_transfer_in * -1
                        .Cells(Me.ColLocation_StockTransferOut.Index).Value = Me.rowLocation.stock_transfer_out * -1
                        .Cells(Me.colLocationPrimaryRefParty.Index).Value = Me.rowLocation.primary_location
                    End With
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvLocation.AllowUserToAddRows = True
            Me.dgvLocation.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        End Try
    End Sub

    Private Sub dgvUser_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvUser.Sorted
        Try
            Me.objCd.SetNumber(Me.dgvUser)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub dgvUser_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvUser.UserDeletedRow
        Try
            If Not String.IsNullOrEmpty(Me.UserCodeDeleted) Then
                ''==
                For i As Integer = Me.tblUser.Rows.Count - 1 To 0 Step -1
                    Me.rowUser = Me.tblUser.Rows(i)
                    If Me.rowUser.usercode = Me.UserCodeDeleted Then
                        Me.tblUser.Rows.RemoveAt(i)
                    End If
                Next
                For i As Integer = Me.tblLocation.Rows.Count - 1 To 0 Step -1
                    Me.rowLocation = Me.tblLocation.Rows(i)
                    If Me.rowLocation.usercode = Me.UserCodeDeleted Then
                        Me.tblLocation.Rows.RemoveAt(i)
                    End If
                Next
                ''==
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub dgvUser_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvUser.UserDeletingRow
        Try
            Me.UserCodeDeleted = Me.dgvUser.Item(Me.colUser_Code.Index, e.Row.Index).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub dgvLocation_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvLocation.CellBeginEdit
        Try
            If e.ColumnIndex = Me.colLocation_No.Index Or e.ColumnIndex = Me.colLocation_Desc.Index Then
                e.Cancel = True
                Exit Sub
            Else
                ''bisa input jika ada user yg dipilih
                If String.IsNullOrEmpty(Me.dgvUser.Item(Me.colUser_Code.Index, Me.dgvUser.SelectedCells(0).RowIndex).Value) Then
                    MsgBox("Select user first.", MsgBoxStyle.Information, "Information")
                    e.Cancel = True
                    Exit Sub
                End If

                ''boleh input lagi jika diatasnya sudah isi code
                Dim i As Integer = Me.GetInvalidDataRowIndex(Me.dgvLocation, e.RowIndex, Me.colLocation_Code.Index)
                If i <> -1 Then
                    MsgBox("Please complete data at row " & i + 1 & ".", MsgBoxStyle.Information, "Information")
                    e.Cancel = True
                    Exit Sub
                End If

                ''boleh edit yang lain asal sudah pilih level
                If e.ColumnIndex <> Me.colLocation_Level.Index Then
                    If String.IsNullOrEmpty(Me.dgvLocation.Item(Me.colLocation_Level.Index, e.RowIndex).Value) Then
                        MsgBox("Select level first.", MsgBoxStyle.Information, "Information")
                        e.Cancel = True
                        Exit Sub
                    End If
                End If

                ''tak boleh ganti level jika sudah ada code
                If e.ColumnIndex = Me.colLocation_Level.Index Then
                    If Not String.IsNullOrEmpty(Me.dgvLocation.Item(Me.colLocation_Code.Index, e.RowIndex).Value) Then
                        MsgBox("Cannot change level if code is defined.", MsgBoxStyle.Information, "Information")
                        e.Cancel = True
                        Exit Sub
                    End If
                End If

                ''tak boleh edit yang lain jika code belum diisi
                If e.ColumnIndex <> Me.colLocation_Code.Index And e.ColumnIndex <> Me.colLocation_Level.Index Then
                    If String.IsNullOrEmpty(Me.dgvLocation.Item(Me.colLocation_Code.Index, e.RowIndex).Value) Then
                        MsgBox("Input code first.", MsgBoxStyle.Information, "Informatio")
                        e.Cancel = True
                        Exit Sub
                    End If
                End If

                Me.Buffer = Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub dgvLocation_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvLocation.CellEndEdit
        Try
            If Not IsNothing(Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value) Then
                If Not String.IsNullOrEmpty(Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim) Then
                    If e.ColumnIndex = Me.colLocation_Level.Index Then

                    End If
                    If e.ColumnIndex = Me.colLocation_Code.Index Then
                        Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.Replace("'", "").Replace("*", "").Trim.ToUpper
                        If Not IsNothing(Me.Buffer) Then
                            If Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value.ToString = Me.Buffer.ToString Then
                                Exit Sub
                            End If
                        End If
                        Me.CellEndEdit_Code(sender, e)
                    End If
                    If e.ColumnIndex = Me.ColLocation_StockBundlingAfter.Index Or e.ColumnIndex = Me.ColLocation_StockBundlingBefore.Index Or _
                            e.ColumnIndex = Me.ColLocation_StockInAdjust.Index Or e.ColumnIndex = Me.colLocation_StockInPurchase.Index Or _
                            e.ColumnIndex = Me.colLocation_StockInSalesReturn.Index Or e.ColumnIndex = Me.ColLocation_StockOutAdjust.Index Or _
                            e.ColumnIndex = Me.ColLocation_StockOutPurcahseReturn.Index Or e.ColumnIndex = Me.ColLocation_StockOutSales.Index Or _
                            e.ColumnIndex = Me.ColLocation_StockTransferIn.Index Or e.ColumnIndex = Me.ColLocation_StockTransferOut.Index Then
                        Me.CellEndEdit_LocationAccess(sender, e)
                    End If
                Else
                    Me.DataGridViewRowDelete(sender, e, Me.dgvLocation)
                End If
            Else
                Me.DataGridViewRowDelete(sender, e, Me.dgvLocation)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub dgvLocation_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvLocation.DataError

    End Sub

    Private Sub dgvLocation_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvLocation.RowsAdded
        Dim DataGridViewRow_ As DataGridViewRow
        Try
            If Me.dgvLocation.AllowUserToAddRows Then
                DataGridViewRow_ = Me.dgvLocation.Rows(Me.dgvLocation.RowCount - 1)
                DataGridViewRow_.Cells(Me.colLocation_Code.Index).Value = String.Empty
                DataGridViewRow_.Cells(Me.colLocation_Desc.Index).Value = String.Empty
                DataGridViewRow_.Cells(Me.colLocation_Level.Index).Value = ""
                DataGridViewRow_.Cells(Me.colLocation_No.Index).Value = Me.dgvLocation.RowCount
                DataGridViewRow_.Cells(Me.ColLocation_StockBundlingAfter.Index).Value = True
                DataGridViewRow_.Cells(Me.ColLocation_StockBundlingBefore.Index).Value = True
                DataGridViewRow_.Cells(Me.ColLocation_StockInAdjust.Index).Value = True
                DataGridViewRow_.Cells(Me.colLocation_StockInPurchase.Index).Value = True
                DataGridViewRow_.Cells(Me.colLocation_StockInSalesReturn.Index).Value = True
                DataGridViewRow_.Cells(Me.ColLocation_StockOutAdjust.Index).Value = True
                DataGridViewRow_.Cells(Me.ColLocation_StockOutPurcahseReturn.Index).Value = True
                DataGridViewRow_.Cells(Me.ColLocation_StockOutSales.Index).Value = True
                DataGridViewRow_.Cells(Me.ColLocation_StockTransferIn.Index).Value = True
                DataGridViewRow_.Cells(Me.ColLocation_StockTransferOut.Index).Value = True
                DataGridViewRow_.Cells(Me.colLocationPrimaryRefParty.Index).Value = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub dgvLocation_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvLocation.RowsRemoved
        Try
            Me.objCd.SetNumber(Me.dgvLocation)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub dgvLocation_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvLocation.Sorted
        Try
            Me.objCd.SetNumber(Me.dgvLocation)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub dgvLocation_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvLocation.UserDeletedRow
        Try
            For i As Integer = Me.tblLocation.Rows.Count - 1 To 0 Step -1
                Me.rowLocation = Me.tblLocation.Rows(i)
                If Me.rowLocation.location_level = Me.LevelDeleted And Me.rowLocation.primary_location = Me.PrimaryRefPartyDeleted Then
                    Me.tblLocation.Rows.RemoveAt(i)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub dgvLocation_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvLocation.UserDeletingRow
        Try
            Me.PrimaryRefPartyDeleted = Me.dgvLocation.Item(Me.colLocationPrimaryRefParty.Index, e.Row.Index).Value
            Me.LevelDeleted = 0
            If Me.dgvLocation.Item(Me.colLocation_Level.Index, e.Row.Index).Value = "Location" Then
                Me.LevelDeleted = 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            If Me.DataGridViewRowToDelete > -1 And Me.DataGridViewRowToDelete < Me.DataGridViewToDelete.Rows.Count - 1 Then
                Me.DataGridViewToDelete.Rows.RemoveAt(Me.DataGridViewRowToDelete)
                Me.SettinglocationAccessBerubah = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Sub DataGridViewRowDelete(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, ByVal vDataGridView As DataGridView)
        Try
            If Me.Buffer = Nothing Then
                Me.DataGridViewRowToDelete = e.RowIndex
                Me.DataGridViewToDelete = vDataGridView
                Me.BackgroundWorker2.RunWorkerAsync()
            Else
                vDataGridView.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Function IsExistUser(ByVal vUserCode As String, ByVal vCurrentRow As Integer) As Boolean
        Try
            Me.dgvUser.AllowUserToAddRows = False
            For Each DataGridViewRow_ As DataGridViewRow In Me.dgvUser.Rows
                If Not String.IsNullOrEmpty(DataGridViewRow_.Cells(Me.colUser_Code.Index).Value) Then
                    If vUserCode = DataGridViewRow_.Cells(Me.colUser_Code.Index).Value And DataGridViewRow_.Index <> vCurrentRow Then
                        Return True
                    End If
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            Me.dgvUser.AllowUserToAddRows = True
        End Try
    End Function

    Sub CellEndEdit_Code(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim frmFindWarehouse As New Core_Module.frmCariMasterWarehouse
        Dim frmFindLocation As New Core_Module.FrmCariMasterStorageArea
        Try
            If Me.dgvLocation.Item(Me.colLocation_Level.Index, e.RowIndex).Value = "Warehouse" Then
                If Me.objWarehouse.ApakahSudahAda(Me.dgvLocation.Item(Me.colLocation_Code.Index, e.RowIndex).Value) = False Then
                    frmFindWarehouse.txtCari.Text = Me.dgvLocation.Item(Me.colLocation_Code.Index, e.RowIndex).Value
                    frmFindWarehouse.ShowDialog()
                    If frmFindWarehouse.PrimaryWarehouse <> -1 Then
                        Me.dgvLocation(e.ColumnIndex, e.RowIndex).Value = Me.objWarehouse.GetWarehouseCode(frmFindWarehouse.PrimaryWarehouse)
                    Else
                        Me.DataGridViewRowDelete(sender, e, Me.dgvLocation)
                        Exit Sub
                    End If
                End If
                If Not String.IsNullOrEmpty(Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value) Then
                    If Me.IsExistWarehouseLocation(True, Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex) Then
                        MsgBox(Me.dgvLocation.Item(Me.colLocation_Level.Index, e.RowIndex).Value & " " & Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                        Me.DataGridViewRowDelete(sender, e, Me.dgvLocation)
                        Exit Sub
                    End If
                    Me.dgvLocation.Item(Me.colLocationPrimaryRefParty.Index, e.RowIndex).Value = Me.objWarehouse.GetPrimary(Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value)
                    Me.dgvLocation.Item(Me.colLocation_Desc.Index, e.RowIndex).Value = Me.objWarehouse.GetDescription(Me.dgvLocation.Item(Me.colLocationPrimaryRefParty.Index, e.RowIndex).Value)
                End If
            ElseIf Me.dgvLocation.Item(Me.colLocation_Level.Index, e.RowIndex).Value = "Location" Then
                If Me.objLocation.ApakahSudahAda(Me.dgvLocation.Item(Me.colLocation_Code.Index, e.RowIndex).Value) = False Then
                    frmFindLocation.txtCari.Text = Me.dgvLocation.Item(Me.colLocation_Code.Index, e.RowIndex).Value
                    frmFindLocation.ShowDialog()
                    If frmFindLocation.PrimaryStorageArea <> -1 Then
                        Me.dgvLocation(e.ColumnIndex, e.RowIndex).Value = Me.objLocation.GetStorageAreaCode(frmFindLocation.PrimaryStorageArea)
                    Else
                        Me.DataGridViewRowDelete(sender, e, Me.dgvLocation)
                        Exit Sub
                    End If
                End If
                If Not String.IsNullOrEmpty(Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value) Then
                    If Me.IsExistWarehouseLocation(False, Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex) Then
                        MsgBox(Me.dgvLocation.Item(Me.colLocation_Level.Index, e.RowIndex).Value & " " & Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                        Me.DataGridViewRowDelete(sender, e, Me.dgvLocation)
                        Exit Sub
                    End If
                    Me.dgvLocation.Item(Me.colLocationPrimaryRefParty.Index, e.RowIndex).Value = Me.objLocation.GetPrimary(Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value)
                    Me.dgvLocation.Item(Me.colLocation_Desc.Index, e.RowIndex).Value = Me.objLocation.GetDescription(Me.dgvLocation.Item(Me.colLocationPrimaryRefParty.Index, e.RowIndex).Value)
                End If
            Else
                MsgBox("Logic Error: unspecified ComboBoxCell text.", MsgBoxStyle.Critical, "Error")
            End If

            ''==
            For i As Integer = Me.tblLocation.Rows.Count - 1 To 0 Step -1
                Me.rowLocation = Me.tblLocation.Rows(i)
                If Me.rowLocation.usercode = Me.dgvUser.Item(Me.colUser_Code.Index, Me.dgvUser.SelectedCells(0).RowIndex).Value Then
                    Me.tblLocation.Rows.RemoveAt(i)
                End If
            Next
            Me.dgvLocation.AllowUserToAddRows = False
            For Each DataGridViewRow_ As DataGridViewRow In Me.dgvLocation.Rows
                With DataGridViewRow_
                    Me.rowLocation = Me.tblLocation.Newtbl_setting_user_location_privilege_itemRow
                    Me.rowLocation.primary = -1
                    Me.rowLocation.usercode = Me.dgvUser.Item(Me.colUser_Code.Index, Me.dgvUser.SelectedCells(0).RowIndex).Value
                    If DataGridViewRow_.Cells(Me.colLocation_Level.Index).Value = "Warehouse" Then
                        Me.rowLocation.location_level = 0
                    ElseIf DataGridViewRow_.Cells(Me.colLocation_Level.Index).Value = "Location" Then
                        Me.rowLocation.location_level = 1
                    Else
                        MsgBox("Logic Error: unspecified ComboBoxCell Text.", MsgBoxStyle.Critical, "Error")
                    End If
                    Me.rowLocation.primary_location = .Cells(Me.colLocationPrimaryRefParty.Index).Value
                    Me.rowLocation("code") = .Cells(e.ColumnIndex).Value
                    Me.rowLocation("desc") = .Cells(Me.colLocation_Desc.Index).Value
                    Me.tblLocation.Rows.Add(Me.rowLocation)
                End With
            Next
            ''==
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvLocation.AllowUserToAddRows = True
        End Try
    End Sub

    Sub CellEndEdit_LocationAccess(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim level As Integer = 0
        Try

            If Me.dgvLocation.Item(Me.colLocation_Level.Index, e.RowIndex).Value = "Location" Then
                level = 1
            End If
            Me.tblLocation.DefaultView.RowFilter = "[usercode] = '" & Me.dgvUser.Item(Me.colUser_Code.Index, Me.dgvUser.SelectedCells(0).RowIndex).Value & "' AND [location level] = " & level.ToString & " AND [primary location] = " & Me.dgvLocation.Item(Me.colLocationPrimaryRefParty.Index, e.RowIndex).Value.ToString
            If Me.tblLocation.DefaultView.Count > 0 Then
                Me.rowLocation = Me.tblLocation.DefaultView(0).Row
                Me.rowLocation.stock_in_purchase = Me.dgvLocation.Item(Me.colLocation_StockInPurchase.Index, e.RowIndex).Value * -1
                Me.rowLocation.stock_in_sales_return = Me.dgvLocation.Item(Me.colLocation_StockInSalesReturn.Index, e.RowIndex).Value * -1
                Me.rowLocation.stock_in_adjust = Me.dgvLocation.Item(Me.ColLocation_StockInAdjust.Index, e.RowIndex).Value * -1
                Me.rowLocation.stock_out_sales = Me.dgvLocation.Item(Me.ColLocation_StockOutSales.Index, e.RowIndex).Value * -1
                Me.rowLocation.stock_out_purchase_return = Me.dgvLocation.Item(Me.ColLocation_StockOutPurcahseReturn.Index, e.RowIndex).Value * -1
                Me.rowLocation.stock_out_adjust = Me.dgvLocation.Item(Me.ColLocation_StockOutAdjust.Index, e.RowIndex).Value * -1
                Me.rowLocation.stock_transfer_out = Me.dgvLocation.Item(Me.ColLocation_StockTransferOut.Index, e.RowIndex).Value * -1
                Me.rowLocation.stock_transfer_in = Me.dgvLocation.Item(Me.ColLocation_StockTransferIn.Index, e.RowIndex).Value * -1
                Me.rowLocation.stock_bundling_before = Me.dgvLocation.Item(Me.ColLocation_StockBundlingBefore.Index, e.RowIndex).Value * -1
                Me.rowLocation.stock_bundling_after = Me.dgvLocation.Item(Me.ColLocation_StockBundlingAfter.Index, e.RowIndex).Value * -1
                Me.rowLocation.production_log_after = Me.rowLocation.stock_bundling_after
                Me.rowLocation.production_log_after = Me.rowLocation.stock_bundling_before
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Function IsExistWarehouseLocation(ByVal vWarehouse As Boolean, ByVal vCode As String, ByVal vCurrentRow As Integer) As Boolean
        Try
            Me.dgvLocation.AllowUserToAddRows = False
            For Each DataGridViewRow_ As DataGridViewRow In Me.dgvLocation.Rows
                If DataGridViewRow_.Index <> vCurrentRow Then
                    If vWarehouse Then
                        If DataGridViewRow_.Cells(Me.colLocation_Level.Index).Value = "Warehouse" And _
                                DataGridViewRow_.Cells(Me.colLocation_Code.Index).Value = vCode Then
                            Return True
                        End If
                    Else
                        If DataGridViewRow_.Cells(Me.colLocation_Level.Index).Value = "Location" And _
                                DataGridViewRow_.Cells(Me.colLocation_Code.Index).Value = vCode Then
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
            Me.dgvLocation.AllowUserToAddRows = True
        End Try
    End Function

    Function GetInvalidDataRowIndex(ByVal vDataGridView As DataGridView, ByVal vCurrentRowIndx As Integer, ByVal vColumnIndexToCheck As Integer) As Integer
        Try
            'vDataGridView.AllowUserToAddRows = False
            For Each dataGridViewRow As DataGridViewRow In vDataGridView.Rows
                If dataGridViewRow.Index <> vCurrentRowIndx And dataGridViewRow.Index <> vDataGridView.Rows.Count - 1 Then
                    If String.IsNullOrEmpty(dataGridViewRow.Cells(vColumnIndexToCheck).Value) Then
                        Return dataGridViewRow.Index
                    End If
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            'vDataGridView.AllowUserToAddRows = True
        End Try
    End Function

    Public Function HasAccessPrivileges(ByVal vUserCode As String, ByVal vListOfPrimaryLocation As List(Of Double), ByVal vTipeTransInOut As FrmUserLocationPrivilegeSetting.EnumTipeTransaksiInOut, _
           Optional ByVal vShowMsgBox As Boolean = True) As Boolean
        Dim dtaUser As New DseUserLocationPrivilegeSettingTableAdapters.tbl_setting_user_location_privilegeTableAdapter
        Dim dtaLocation As New DseUserLocationPrivilegeSettingTableAdapters.tbl_setting_user_location_privilege_itemTableAdapter
        Dim listOfInvalidLocation As New List(Of String)
        Dim invalidLocation As String = ""
        Try
            'dtaUser.Connection.ConnectionString = Core_Module.GetConstring
            'dtaLocation.Connection.ConnectionString = Core_Module.GetConstring
            dtaUser.FillBy_UserCode(Me.tblUser, vUserCode)
            dtaLocation.FillBy_UserCode(Me.tblLocation, vUserCode)
            For Each primaryLocation As Double In vListOfPrimaryLocation
                If Me.HasAccessPrivileges(vUserCode, primaryLocation, vTipeTransInOut) = False Then
                    Dim locationCode As String = objLocation.GetStorageAreaCode(primaryLocation)
                    If Not listOfInvalidLocation.Contains(locationCode) Then listOfInvalidLocation.Add(locationCode)
                End If
            Next
            If listOfInvalidLocation.Count > 0 Then
                Dim i As Integer = 0
                For Each location As String In listOfInvalidLocation
                    i += 1
                    If i > 10 Then
                        invalidLocation = invalidLocation & vbNewLine & "..."
                        Exit For
                    End If
                    invalidLocation = invalidLocation & vbNewLine & location
                Next
                If vShowMsgBox Then
                    MsgBox("You don't have access to the following location for this transaction:" & invalidLocation, MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Function

    Private Function HasAccessPrivileges(ByVal vUserCode As String, ByVal vPrimaryLocation As Double, _
            ByVal vTipeTransInOut As FrmUserLocationPrivilegeSetting.EnumTipeTransaksiInOut) As Boolean

        Dim DefaultUserNotIncludeInList As String = ""
        Try
            If vPrimaryLocation = -1 Then
                Return True
            End If

            If tblUser.Rows.Count = 0 Then
                DefaultUserNotIncludeInList = Me.objSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Location_Privileges, ClsTransaksiSetting.EnumNamaParameter.Default_User_Not_IncludeInList)
                If DefaultUserNotIncludeInList = "" Or DefaultUserNotIncludeInList = "0" Then
                    Return True
                Else
                    'If vShowMsgBox Then
                    '    MsgBox("You don't have access to any location/warehouse.", MsgBoxStyle.Information, "Information")
                    'End If
                    Return False
                End If
            End If

            tblLocation.DefaultView.RowFilter = "[location level] = 1 and [primary location] = " & vPrimaryLocation.ToString
            If tblLocation.DefaultView.Count > 0 Then
                rowLocation = tblLocation.DefaultView(0).Row
                Dim False_ As Boolean = False
                If vTipeTransInOut = EnumTipeTransaksiInOut.STOCK_BUNDLING_AFTER Then
                    If rowLocation.stock_bundling_after = 0 Then False_ = True
                End If
                If vTipeTransInOut = EnumTipeTransaksiInOut.STOCK_BUNDLING_BEFORE Then
                    If rowLocation.stock_bundling_before = 0 Then False_ = True
                End If
                If vTipeTransInOut = EnumTipeTransaksiInOut.STOCK_IN_ADJUST Then
                    If rowLocation.stock_in_adjust = 0 Then False_ = True
                End If
                If vTipeTransInOut = EnumTipeTransaksiInOut.STOCK_IN_PURCAHSE Then
                    If rowLocation.stock_in_purchase = 0 Then False_ = True
                End If
                If vTipeTransInOut = EnumTipeTransaksiInOut.STOCK_IN_SALES_RETURN Then
                    If rowLocation.stock_in_sales_return = 0 Then False_ = True
                End If
                If vTipeTransInOut = EnumTipeTransaksiInOut.STOCK_OUT_ADJUST Then
                    If rowLocation.stock_out_adjust = 0 Then False_ = True
                End If
                If vTipeTransInOut = EnumTipeTransaksiInOut.STOCK_OUT_PURCHASE_RETURN Then
                    If rowLocation.stock_out_purchase_return = 0 Then False_ = True
                End If
                If vTipeTransInOut = EnumTipeTransaksiInOut.STOCK_OUT_SALES Then
                    If rowLocation.stock_out_sales = 0 Then False_ = True
                End If
                If vTipeTransInOut = EnumTipeTransaksiInOut.STOCK_TANSFER_OUT Then
                    If rowLocation.stock_transfer_out = 0 Then False_ = True
                End If
                If vTipeTransInOut = EnumTipeTransaksiInOut.STOCK_TRANSFER_IN Then
                    If rowLocation.stock_transfer_in = 0 Then False_ = True
                End If
                If False_ Then
                    'If vShowMsgBox Then
                    '    MsgBox("You don't have access to location " & Me.objLocation.GetStorageAreaCode(vPrimaryLocation) & " for this transaction.", MsgBoxStyle.Information, "Information")
                    'End If
                    Return False
                End If
            Else
                tblLocation.DefaultView.RowFilter = "[location level] = 0 and [primary location] = " & Me.objWarehouse.GetPrimary(Me.objLocation.GetWarehouseCode(vPrimaryLocation))
                If tblLocation.DefaultView.Count > 0 Then
                    rowLocation = tblLocation.DefaultView(0).Row
                    Dim False_ As Boolean = False
                    If vTipeTransInOut = EnumTipeTransaksiInOut.STOCK_BUNDLING_AFTER Then
                        If rowLocation.stock_bundling_after = 0 Then False_ = True
                    End If
                    If vTipeTransInOut = EnumTipeTransaksiInOut.STOCK_BUNDLING_BEFORE Then
                        If rowLocation.stock_bundling_before = 0 Then False_ = True
                    End If
                    If vTipeTransInOut = EnumTipeTransaksiInOut.STOCK_IN_ADJUST Then
                        If rowLocation.stock_in_adjust = 0 Then False_ = True
                    End If
                    If vTipeTransInOut = EnumTipeTransaksiInOut.STOCK_IN_PURCAHSE Then
                        If rowLocation.stock_in_purchase = 0 Then False_ = True
                    End If
                    If vTipeTransInOut = EnumTipeTransaksiInOut.STOCK_IN_SALES_RETURN Then
                        If rowLocation.stock_in_sales_return = 0 Then False_ = True
                    End If
                    If vTipeTransInOut = EnumTipeTransaksiInOut.STOCK_OUT_ADJUST Then
                        If rowLocation.stock_out_adjust = 0 Then False_ = True
                    End If
                    If vTipeTransInOut = EnumTipeTransaksiInOut.STOCK_OUT_PURCHASE_RETURN Then
                        If rowLocation.stock_out_purchase_return = 0 Then False_ = True
                    End If
                    If vTipeTransInOut = EnumTipeTransaksiInOut.STOCK_OUT_SALES Then
                        If rowLocation.stock_out_sales = 0 Then False_ = True
                    End If
                    If vTipeTransInOut = EnumTipeTransaksiInOut.STOCK_TANSFER_OUT Then
                        If rowLocation.stock_transfer_out = 0 Then False_ = True
                    End If
                    If vTipeTransInOut = EnumTipeTransaksiInOut.STOCK_TRANSFER_IN Then
                        If rowLocation.stock_transfer_in = 0 Then False_ = True
                    End If
                    If False_ Then
                        'If vShowMsgBox Then
                        '    MsgBox("You don't have access to warehouse " & Me.objLocation.GetWarehouseCode(vPrimaryLocation) & " for this transaction.", MsgBoxStyle.Information, "Information")
                        'End If
                        Return False
                    End If
                Else
                    rowUser = tblUser.Rows(0)
                    If rowUser.default_access = 1 Then
                        'If vShowMsgBox Then
                        '    MsgBox("You don't have access to location " & Me.objLocation.GetStorageAreaCode(vPrimaryLocation) & ".", MsgBoxStyle.Information, "Information")
                        'End If
                        Return False
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SimpanDataSetting_LocationAccess()
        Dim dtaHeader As New DseUserLocationPrivilegeSettingTableAdapters.tbl_setting_user_location_privilegeTableAdapter
        Dim dtaItem As New DseUserLocationPrivilegeSettingTableAdapters.tbl_setting_user_location_privilege_itemTableAdapter

        Try
            If Me.SettinglocationAccessBerubah = True Then
                Me.ObjSetting.SetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Location_Privileges, ClsTransaksiSetting.EnumNamaParameter.Default_User_Not_IncludeInList, Me.cboDefaultUserNotInList.SelectedIndex)
                'dtaHeader.Connection.ConnectionString = Core_Module.GetConstring
                'dtaItem.Connection.ConnectionString = Core_Module.GetConstring
                dtaHeader.DeleteQuery_All()
                dtaItem.DeleteQuery_All()
                For Each Me.rowUser In Me.tblUser
                    dtaHeader.InsertQuery(Me.rowUser.usercode, Me.rowUser.default_access)
                Next
                For Each Me.rowLocation In Me.tblLocation
                    dtaItem.InsertQuery(Me.rowLocation.usercode, Me.rowLocation.location_level, Me.rowLocation.primary_location, Me.rowLocation.stock_in_purchase, _
                            Me.rowLocation.stock_in_sales_return, Me.rowLocation.stock_in_adjust, Me.rowLocation.stock_out_sales, Me.rowLocation.stock_out_purchase_return, _
                            Me.rowLocation.stock_out_adjust, Me.rowLocation.stock_transfer_out, Me.rowLocation.stock_transfer_in, Me.rowLocation.stock_bundling_before, _
                            Me.rowLocation.stock_bundling_after, Me.rowLocation.production_log_before, Me.rowLocation.production_log_after)
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboDefaultUserNotInList_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDefaultUserNotInList.SelectedIndexChanged
        Try
            Me.SettinglocationAccessBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
