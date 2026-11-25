Option Explicit On
Option Strict Off

Public Class FrmLaporanInventoryMovement
    Public UserCode As String = ""
    Dim TBLBrand As New Core_Module.dseMasterProductGroup.tblmaster_product_group2DataTable
    Dim DTA As New dseLaporanInventoryMovementTableAdapters.inventory_movement_viewTableAdapter
    Dim DTA2 As New dseLaporanInventoryMovementTableAdapters.tblmaster_productTableAdapter

    Dim DS As New dseLaporanInventoryMovement
    Dim TBL As dseLaporanInventoryMovement.inventory_movement_viewDataTable = DS.inventory_movement_view
    Dim TBL2 As dseLaporanInventoryMovement.tblmaster_productDataTable = DS.tblmaster_product

    Dim report As New crpInventoryMovementC
    Dim row As dseLaporanInventoryMovement.inventory_movement_viewRow
    Dim row1 As dseLaporanInventoryMovement.tblmaster_productRow
    Dim SudahLoadProduct As Boolean = False
    Dim buffer As String = ""

    Dim objInvMove As New ClsLaporanInventoryMovement

    'aul
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Private LevelGroup As EnumLevel
    Private LevelIndex As EnumLevelIndex

    Private objGroup As New Core_Module.ClsMasterProductGroup
    Private ObjCategory As New Core_Module.ClsMasterCategory
    Private ObjProduct As New Core_Module.clsMasterProduct
    Private ObjCD As New Core_Module.ClsCekData

    Private EraseRowIndexProduct As Integer = 0
    Dim BufferCbo As Integer
    Dim ObjLicense As New Core_Module.ClsLicense

    'aul
#Region "Enum saja"
    Public Enum EnumLevel As Integer
        Category
        SubBrand
        Brand
        ProductGroup
    End Enum

    Public Enum EnumLevelIndex As Integer
        Index01
        Index02
        Index03
    End Enum

    Public Enum EnumPilih
        No
        Yes
    End Enum

    Public Enum JenisProductGroup As Integer
        Product_Group
        Brand
        Sub_Brand
        Category
        Product
    End Enum
#End Region

    'aul
    Public Sub SalinKeObjek()
        Try
            With objInvMove
                .User = UserCode
                If Me.CboGrouping.SelectedIndex = 0 Then
                    .Grouping = 0
                ElseIf Me.CboGrouping.SelectedIndex = 1 Then
                    .Grouping = 1
                ElseIf Me.CboGrouping.SelectedIndex = 2 Then
                    .Grouping = 2
                End If

                If Me.CboGroup1.SelectedIndex = 0 Then
                    .Group1 = 0
                ElseIf Me.CboGroup1.SelectedIndex = 1 Then
                    .Group1 = 1
                ElseIf Me.CboGroup1.SelectedIndex = 2 Then
                    .Group1 = 2
                ElseIf Me.CboGroup1.SelectedIndex = 3 Then
                    .Group1 = 3
                End If

                If Me.CboGroup2.SelectedIndex = 0 Then
                    .Group2 = 0
                ElseIf Me.CboGroup2.SelectedIndex = 1 Then
                    .Group2 = 1
                ElseIf Me.CboGroup2.SelectedIndex = 2 Then
                    .Group2 = 2
                ElseIf Me.CboGroup2.SelectedIndex = 3 Then
                    .Group2 = 3
                End If

                'GroupLevel
                If Me.CboPilihGroupLevel.SelectedIndex = 0 Then
                    .GroupLevel = 0
                ElseIf Me.CboPilihGroupLevel.SelectedIndex = 1 Then
                    .GroupLevel = 1
                ElseIf Me.CboPilihGroupLevel.SelectedIndex = 2 Then
                    .GroupLevel = 2
                ElseIf Me.CboPilihGroupLevel.SelectedIndex = 3 Then
                    .GroupLevel = 3
                ElseIf Me.CboPilihGroupLevel.SelectedIndex = 4 Then
                    .GroupLevel = 4
                End If
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeField()
        Try
            With objInvMove

                Me.TxtPrimary.Text = .Primary
                If .Grouping = 0 Then
                    Me.CboGrouping.SelectedIndex = 0
                ElseIf .Grouping = 1 Then
                    Me.CboGrouping.SelectedIndex = 1
                ElseIf .Grouping = 2 Then
                    Me.CboGrouping.SelectedIndex = 2
                End If

                If .Group1 = 0 Then
                    Me.CboGroup1.SelectedIndex = 0
                ElseIf .Group1 = 1 Then
                    Me.CboGroup1.SelectedIndex = 1
                ElseIf .Group1 = 2 Then
                    Me.CboGroup1.SelectedIndex = 2
                ElseIf .Group1 = 3 Then
                    Me.CboGroup1.SelectedIndex = 3
                End If

                If .Group2 = 0 Then
                    Me.CboGroup2.SelectedIndex = 0
                ElseIf .Group2 = 1 Then
                    Me.CboGroup2.SelectedIndex = 1
                ElseIf .Group2 = 2 Then
                    Me.CboGroup2.SelectedIndex = 2
                ElseIf .Group2 = 3 Then
                    Me.CboGroup2.SelectedIndex = 3
                End If

                'grouplevel
                If .GroupLevel = 0 Then
                    Me.CboPilihGroupLevel.SelectedIndex = 0
                ElseIf .GroupLevel = 1 Then
                    Me.CboPilihGroupLevel.SelectedIndex = 1
                ElseIf .GroupLevel = 2 Then
                    Me.CboPilihGroupLevel.SelectedIndex = 2
                ElseIf .GroupLevel = 3 Then
                    Me.CboPilihGroupLevel.SelectedIndex = 3
                ElseIf .GroupLevel = 4 Then
                    Me.CboPilihGroupLevel.SelectedIndex = 4
                End If

                objInvMove.IsiDGVProductGroup(Me, objInvMove.Primary, Me.DgvGroup)
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'aul
    Private Sub CboPilihGroupLevel_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboPilihGroupLevel.SelectedIndexChanged
        Try
            If Me.CboPilihGroupLevel.SelectedIndex = 0 Then
                LevelGroup = EnumLevel.ProductGroup
                LevelIndex = EnumLevelIndex.Index01
                Me.DgvGroup.Rows.Clear()
                Me.DgvGroup.AllowUserToAddRows = True
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = 1 Then
                LevelGroup = EnumLevel.Brand
                LevelIndex = EnumLevelIndex.Index01
                Me.DgvGroup.Rows.Clear()
                Me.DgvGroup.AllowUserToAddRows = True
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = 2 Then
                LevelGroup = EnumLevel.SubBrand
                LevelIndex = EnumLevelIndex.Index01
                Me.DgvGroup.Rows.Clear()
                Me.DgvGroup.AllowUserToAddRows = True
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = 3 Then
                LevelGroup = EnumLevel.Category
                LevelIndex = EnumLevelIndex.Index02
                Me.DgvGroup.Rows.Clear()
                Me.DgvGroup.AllowUserToAddRows = True
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = 4 Then
                LevelIndex = EnumLevelIndex.Index03
                Me.DgvGroup.Rows.Clear()
                Me.DgvGroup.AllowUserToAddRows = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Dim counter As Double = 0
        Dim Filter As String = ""
        Dim PrimaryAcuan As Double = -1
        Dim rowDgv As DataGridViewRow
        Dim TypeScope As Integer = -1
        Dim parGroup1 As String = ""
        Dim parGroup2 As String = ""
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.dtpStartDate.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpStartDate.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpEndDate.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpEndDate.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If Me.ApakahBolehCetak = False Then
                Exit Sub
            End If

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.TBL.Rows.Clear()

            If Me.dtpStartDate.Value.Date <= Me.dtpEndDate.Value.Date Then
                'isikan data
                If Me.chkGoods.Checked = True And Me.chkService.Checked = False Then
                    TypeScope = 0
                ElseIf Me.chkService.Checked = True And Me.chkGoods.Checked = False Then
                    TypeScope = 1
                End If
                If Me.CboGrouping.SelectedIndex = 0 Or Me.CboGrouping.SelectedIndex = 1 Then
                    If Me.CboGroup1.SelectedIndex = 0 Then
                        parGroup1 = "Product Group"
                    ElseIf Me.CboGroup1.SelectedIndex = 1 Then
                        parGroup1 = "Brand"
                    ElseIf Me.CboGroup1.SelectedIndex = 2 Then
                        parGroup1 = "Sub Brand"
                    ElseIf Me.CboGroup1.SelectedIndex = 3 Then
                        parGroup1 = "Category"
                    End If
                    If Me.CboGrouping.SelectedIndex = 0 Then
                        If Me.CboGroup2.SelectedIndex = 0 Then
                            parGroup2 = "Product Group"
                        ElseIf Me.CboGroup2.SelectedIndex = 1 Then
                            parGroup2 = "Brand"
                        ElseIf Me.CboGroup2.SelectedIndex = 2 Then
                            parGroup2 = "Sub Brand"
                        ElseIf Me.CboGroup2.SelectedIndex = 3 Then
                            parGroup2 = "Category"
                        End If
                    Else
                        parGroup2 = ""
                    End If
                Else
                    parGroup1 = ""
                    parGroup2 = ""
                End If

                'filter dgv product group
                If Me.DgvGroup.RowCount > 1 Then
                    For i As Integer = 0 To Me.DgvGroup.RowCount - 2
                        rowDgv = Me.DgvGroup.Rows(i)
                        PrimaryAcuan = rowDgv.Cells(Me.ColGroup_Primary.Index).Value
                        DTA.ClearBeforeFill = False
                        If Me.chkGoods.Checked = True And Me.chkService.Checked = True Then
                            If LevelGroup = EnumLevel.ProductGroup And LevelIndex = EnumLevelIndex.Index01 Then
                                Me.DTA.FillBy_PrimaryGroup(TBL, PrimaryAcuan, Me.dtpEndDate.Value, Me.dtpStartDate.Value)
                            ElseIf LevelGroup = EnumLevel.Brand And LevelIndex = EnumLevelIndex.Index01 Then
                                Me.DTA.FillBy_PrimaryBrand(TBL, PrimaryAcuan, Me.dtpEndDate.Value, Me.dtpStartDate.Value)
                            ElseIf LevelGroup = EnumLevel.SubBrand And LevelIndex = EnumLevelIndex.Index01 Then
                                Me.DTA.FillBy_PrimarySubBrand(TBL, PrimaryAcuan, Me.dtpEndDate.Value, Me.dtpStartDate.Value)
                            ElseIf LevelGroup = EnumLevel.Category And LevelIndex = EnumLevelIndex.Index02 Then
                                Me.DTA.FillBy_PrimaryCategory(TBL, PrimaryAcuan, Me.dtpEndDate.Value, Me.dtpStartDate.Value)
                            ElseIf LevelIndex = EnumLevelIndex.Index03 Then
                                Me.DTA.FillBy_PrimaryProduct(TBL, PrimaryAcuan, Me.dtpEndDate.Value, Me.dtpStartDate.Value)
                            End If
                        ElseIf Me.chkGoods.Checked = True Or Me.chkService.Checked = True Then
                            If LevelGroup = EnumLevel.ProductGroup And LevelIndex = EnumLevelIndex.Index01 Then
                                Me.DTA.FillBy_PrimaryGroupNType(TBL, PrimaryAcuan, TypeScope, Me.dtpEndDate.Value, Me.dtpStartDate.Value)
                            ElseIf LevelGroup = EnumLevel.Brand And LevelIndex = EnumLevelIndex.Index01 Then
                                Me.DTA.FillBy_PrimaryBrandNType(TBL, PrimaryAcuan, TypeScope, Me.dtpEndDate.Value, Me.dtpStartDate.Value)
                            ElseIf LevelGroup = EnumLevel.SubBrand And LevelIndex = EnumLevelIndex.Index01 Then
                                Me.DTA.FillBy_PrimarySubBrandNType(TBL, PrimaryAcuan, TypeScope, Me.dtpEndDate.Value, Me.dtpStartDate.Value)
                            ElseIf LevelGroup = EnumLevel.Category And LevelIndex = EnumLevelIndex.Index02 Then
                                Me.DTA.FillBy_PrimaryCategoryNType(TBL, PrimaryAcuan, TypeScope, Me.dtpEndDate.Value, Me.dtpStartDate.Value)
                            ElseIf LevelIndex = EnumLevelIndex.Index03 Then
                                Me.DTA.FillBy_PrimaryProductNType(TBL, PrimaryAcuan, TypeScope, Me.dtpEndDate.Value, Me.dtpStartDate.Value)
                            End If
                        Else
                            'MsgBox("Please check type product", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If
                    Next
                ElseIf Me.DgvGroup.RowCount = 1 Then
                    If Me.chkGoods.Checked = True And Me.chkService.Checked = True Then
                        Me.DTA.Fill(TBL, Me.dtpEndDate.Value, Me.dtpStartDate.Value)
                    ElseIf Me.chkGoods.Checked = True Or Me.chkService.Checked = True Then
                        Me.DTA.Fill_ByType(TBL, TypeScope, Me.dtpEndDate.Value, Me.dtpStartDate.Value)
                    End If
                End If
                'tutup filter dgv product group 

                If Me.CboGrouping.SelectedIndex = 0 Or Me.CboGrouping.SelectedIndex = 1 Then
                    For Each row In TBL
                        With row
                            If Me.CboGroup1.SelectedIndex = 0 Then
                                .Group_1 = .Group_Code
                                .Group_1_Desc = .Group_Desc
                            ElseIf Me.CboGroup1.SelectedIndex = 1 Then
                                .Group_1 = .Brand_Code
                                .Group_1_Desc = .Brand_Desc
                            ElseIf Me.CboGroup1.SelectedIndex = 2 Then
                                .Group_1 = .Sub_Brand_Code
                                .Group_1_Desc = .Sub_Brand_Desc
                            ElseIf Me.CboGroup1.SelectedIndex = 3 Then
                                .Group_1 = .Category_Code
                                .Group_1_Desc = .Category_Description
                            End If
                            If Me.CboGrouping.SelectedIndex = 0 Then
                                If Me.CboGroup2.SelectedIndex = 0 Then
                                    .Group_2 = .Group_Code
                                    .Group_2_Desc = .Group_Desc
                                ElseIf Me.CboGroup2.SelectedIndex = 1 Then
                                    .Group_2 = .Brand_Code
                                    .Group_2_Desc = .Brand_Desc
                                ElseIf Me.CboGroup2.SelectedIndex = 2 Then
                                    .Group_2 = .Sub_Brand_Code
                                    .Group_2_Desc = .Sub_Brand_Desc
                                ElseIf Me.CboGroup2.SelectedIndex = 3 Then
                                    .Group_2 = .Category_Code
                                    .Group_2_Desc = .Category_Description
                                End If
                            End If
                        End With
                    Next
                End If

                If Me.SudahLoadProduct = False Then
                    DTA2.Fill(TBL2, Me.dtpEndDate.Value)
                End If

                While counter < TBL.Rows.Count
                    row = TBL.Rows(counter)
                    Filter = "Primary =" + row.Product_Primary.ToString
                    TBL2.DefaultView.RowFilter = Filter
                    If TBL2.DefaultView.Count > 0 Then
                        row1 = TBL2.DefaultView(0).Row
                        row.Ending_Balance = row1.End_Balance
                        row.Beginning_Balance = row.Ending_Balance - row.Moving_In + row._Moving__Out_
                    Else
                        row.Ending_Balance = 0
                        row.Beginning_Balance = 0
                    End If

                    If row.Beginning_Balance = 0 And row.Moving_In = 0 And row._Moving__Out_ = 0 And _
                    row.Ending_Balance = 0 Then
                        TBL.Rows.Remove(row)
                    Else
                        counter += 1
                    End If
                End While

                Dim Crv As CrystalDecisions.Windows.Forms.CrystalReportViewer = Me.CrvLaporan
                'tampilkan report
                report.SetDataSource(DS)
                report.SetParameterValue(report.Parameter_paramGroup1.ParameterFieldName, parGroup1)
                report.SetParameterValue(report.Parameter_paramGroup2.ParameterFieldName, parGroup2)
                report.SetParameterValue(0, Me.dtpStartDate.Value)
                report.SetParameterValue(1, Me.dtpEndDate.Value)

                Crv.ReportSource = report
                Crv.ShowRefreshButton = False
                Crv.Show()
            Else
                MsgBox("Wrong date filter, baginning date must be smaller than end date", _
                MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub dtpEndDate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpEndDate.Enter
        Me.buffer = Format(Me.dtpEndDate.Value, "dd MMM yyyy")
    End Sub

    Private Sub dtpEndDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpEndDate.ValueChanged
        Try
            If Me.dtpEndDate.Value <> CDate(Me.buffer) Then
                Me.SudahLoadProduct = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim FrmCetak1 As New FrmCetakLaporanInventoryMovement
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.dtpStartDate.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpStartDate.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpEndDate.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpEndDate.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If Me.dgvStockCard.Rows.Count > 0 Then
                Dim Crv As CrystalDecisions.Windows.Forms.CrystalReportViewer = FrmCetak1.crvStockCard
                'tampilkan report
                report.SetDataSource(DS)
                report.SetParameterValue(0, Me.dtpStartDate.Value)
                report.SetParameterValue(1, Me.dtpEndDate.Value)

                Crv.ReportSource = report
                Crv.ShowRefreshButton = False
                Crv.Show()

                FrmCetak1.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmCetak1.Dispose()
        End Try
    End Sub

    Private Function ApakahBolehCetak() As Boolean
        Try
            'If Me.clbBrand.CheckedItems.Count <= 0 Then
            '    MsgBox("Please select at least 1 Brand.", MsgBoxStyle.Information, "Information")
            '    Me.clbBrand.Focus()
            '    Return False
            'End If
            Me.BufferCbo = Me.CboGroup2.SelectedIndex
            If CboSetting() = False Then
                Me.CboGroup2.SelectedIndex = Me.BufferCbo
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub FrmLaporanInventoryMovement_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Me.SalinKeObjek()
            objInvMove.Simpan(Me.DgvGroup, Me, UserCode)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanInventoryMovement_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Dim OBjBrand As New Core_Module.ClsMasterProductGroup
        Try
            Me.ColGroup_Primary.Visible = True
            If Me.objInvMove.CariDetail(UserCode, Me) = True Then
                Me.SalinKeField()
            Else
                Me.CboPilihGroupLevel.SelectedIndex = 0
                Me.CboGrouping.SelectedIndex = 0
                Me.CboGroup1.SelectedIndex = 0
                Me.CboGroup2.SelectedIndex = 1
            End If

            Me.dtpStartDate.Value = Today - TimeSpan.FromDays(30)
            Me.ColGroup_Primary.Visible = False
            'OBjBrand.IsiListBoxBrand(Me.clbBrand, Me.TBLBrand)
            'Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
            Me.DgvGroup.AlternatingRowsDefaultCellStyle.BackColor = BackColorAll
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'OBjBrand = Nothing
        End Try
    End Sub

    'Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
    '    Try
    '        For i As Integer = 0 To Me.clbBrand.Items.Count - 1
    '            Me.clbBrand.SetItemCheckState(i, CheckState.Checked)
    '        Next
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
    '    Try
    '        For i As Integer = 0 To Me.clbBrand.Items.Count - 1
    '            Me.clbBrand.SetItemCheckState(i, CheckState.Unchecked)
    '        Next
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'aul

#Region "Dgv Product Group"
    Private Sub DgvProductGroup_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvGroup.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColGroup_Code.Index Then
                    '''set header fungsi digunakan ketika edit code 
                    'Me.DgvGroup.Columns(Me.ColGroup_Code.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvGroup.Columns(Me.ColGroup_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    '''simpan nilai awal di buffer
                    Me.buffer = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColGroup_Code.Index Then
                    '''set header fungsi digunakan ketika edit code 
                    'Me.DgvGroup.Columns(Me.ColGroup_Code.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvGroup.Columns(Me.ColGroup_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.buffer = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvGroup_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvGroup.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvGroup.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryGroup As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Dim _Group As Boolean = True
        Try
            With row
                If IsNothing(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColGroup_Code.Index Then
                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper

                        If Me.LevelIndex = EnumLevelIndex.Index01 Then
                            If Me.objGroup.ApakahSudahAda(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value, Me.LevelGroup) = True Then
                                PrimaryGroup = Me.objGroup.GetPrimary(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value, Me.LevelGroup)
                                index = Me.GetIndexItemGroup(PrimaryGroup, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = PrimaryGroup
                                    Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.objGroup.GetGroupCode(PrimaryGroup)
                                    Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.objGroup.GetGroupDesc(PrimaryGroup)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = ""
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseRowIndexProduct = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                                        Exit Sub
                                    End If
                                End If
                            Else
                                Dim ij As New Core_Module.FrmCariMasterProductGroup
                                ij.ProductGroupLevel = Me.LevelGroup
                                ij.txtCari.Text = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value
                                ij.BackColorAll = Me.BackColorAll
                                ij.ShowDialog()
                                If ij.PrimaryProductGroup <> -1 Then

                                    index = Me.GetIndexItemGroup(ij.PrimaryProductGroup, e.RowIndex)
                                    If index < 0 Then
                                        Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = ij.PrimaryProductGroup
                                        Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.objGroup.GetGroupCode(ij.PrimaryProductGroup)
                                        Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.objGroup.GetGroupDesc(ij.PrimaryProductGroup)
                                    Else
                                        If e.RowIndex <> index Then
                                            MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                                Me.EraseRowIndexProduct = e.RowIndex
                                                Me.BackgroundWorker2.RunWorkerAsync()
                                            Else
                                                Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                            End If
                                            'Me.ObjCD.SetNumber(Me.dgvGroup)
                                            Exit Sub
                                        Else
                                            Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = ij.PrimaryProductGroup
                                            Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.objGroup.GetGroupCode(ij.PrimaryProductGroup)
                                            Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.objGroup.GetGroupDesc(ij.PrimaryProductGroup)
                                        End If
                                    End If
                                ElseIf ij.PrimaryProductGroup = -1 Then
                                    If Me.DgvGroup.Rows.Count - 1 <> e.RowIndex Then
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseRowIndexProduct = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                                        Exit Sub
                                    End If
                                    If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseRowIndexProduct = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                                        Exit Sub
                                    End If
                                End If
                                ij = Nothing
                            End If
                            'Me.ObjCD.SetNumber(Me.dgvGroup)
                        ElseIf Me.LevelIndex = EnumLevelIndex.Index02 Then
                            EndEditCategory(sender, e)
                        ElseIf Me.LevelIndex = EnumLevelIndex.Index03 Then
                            EndEditProduct(sender, e)
                        End If
                    End If
                Else
                    If Me.DgvGroup.Rows.Count - 1 <> e.RowIndex Then
                        If String.IsNullOrEmpty(Me.buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                        Exit Sub
                    End If
                    If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                        If String.IsNullOrEmpty(Me.buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                        Exit Sub
                    End If
                End If
                'Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'Me.DgvGroup.Columns(Me.ColGroup_Code.Index).SortMode = DataGridViewColumnSortMode.Automatic
            'Me.DgvGroup.Columns(Me.ColGroup_Desc.Index).SortMode = DataGridViewColumnSortMode.Automatic
        End Try
    End Sub

    Private Sub EndEditCategory(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim row As DataGridViewRow = Me.DgvGroup.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryGroup As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Dim LvlGroup As Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.ProductGroup
        Try
            With row
                If IsNothing(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColGroup_Code.Index Then
                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjCategory.ApakahSudahAda(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value, LvlGroup) = True Then
                            PrimaryGroup = Me.ObjCategory.GetPrimary(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemGroup(PrimaryGroup, e.RowIndex)
                            If index < 0 Then
                                Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = PrimaryGroup
                                Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.ObjCategory.GetCategoryCode(PrimaryGroup)
                                Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(PrimaryGroup)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvGroup)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.frmCariMasterCategory
                            ij.txtCari.Text = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryCategory <> -1 Then

                                index = Me.GetIndexItemGroup(ij.PrimaryCategory, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = ij.PrimaryCategory
                                    Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.ObjCategory.GetCategoryCode(ij.PrimaryCategory)
                                    Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(ij.PrimaryCategory)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseRowIndexProduct = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                                        Exit Sub
                                    Else
                                        Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = ij.PrimaryCategory
                                        Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.ObjCategory.GetCategoryCode(ij.PrimaryCategory)
                                        Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(ij.PrimaryCategory)
                                    End If
                                End If
                            ElseIf ij.PrimaryCategory = -1 Then
                                If Me.DgvGroup.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvGroup)
                                    Exit Sub
                                End If
                                If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvGroup)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                    End If
                Else
                    If Me.DgvGroup.Rows.Count - 1 <> e.RowIndex Then
                        If String.IsNullOrEmpty(Me.buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                        Exit Sub
                    End If
                    If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                        If String.IsNullOrEmpty(Me.buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                        Exit Sub
                    End If
                End If
                'Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            'Finally
            '    Me.DgvGroup.Columns(Me.ColGroup_Code.Index).SortMode = DataGridViewColumnSortMode.Automatic
            '    Me.DgvGroup.Columns(Me.ColGroup_Desc.Index).SortMode = DataGridViewColumnSortMode.Automatic
        End Try
    End Sub

    Private Sub EndEditProduct(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim row As DataGridViewRow = Me.DgvGroup.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryGroup As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Dim LvlGroup As Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.ProductGroup
        Try
            With row
                If IsNothing(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColGroup_Code.Index Then
                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        If Me.ObjProduct.ApakahSudahAda(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryGroup = Me.ObjProduct.GetPrimary(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemGroup(PrimaryGroup, e.RowIndex)
                            If index < 0 Then
                                Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = PrimaryGroup
                                Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.ObjProduct.GetProductCode(PrimaryGroup)
                                Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(PrimaryGroup)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvGroup)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.frmCariMasterProduct
                            ij.txtCari.Text = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryProduct <> -1 Then

                                index = Me.GetIndexItemGroup(ij.PrimaryProduct, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = ij.PrimaryProduct
                                    Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.ObjProduct.GetProductCode(ij.PrimaryProduct)
                                    Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(ij.PrimaryProduct)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseRowIndexProduct = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                                        Exit Sub
                                    Else
                                        Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = ij.PrimaryProduct
                                        Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.ObjProduct.GetProductCode(ij.PrimaryProduct)
                                        Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(ij.PrimaryProduct)
                                    End If
                                End If
                            ElseIf ij.PrimaryProduct = -1 Then
                                If Me.DgvGroup.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvGroup)
                                    Exit Sub
                                End If
                                If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvGroup)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                    End If
                Else
                    If Me.DgvGroup.Rows.Count - 1 <> e.RowIndex Then
                        If String.IsNullOrEmpty(Me.buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                        Exit Sub
                    End If
                    If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                        If String.IsNullOrEmpty(Me.buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                        Exit Sub
                    End If
                End If
                'Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            'Finally
            '    Me.DgvGroup.Columns(Me.ColGroup_Code.Index).SortMode = DataGridViewColumnSortMode.Automatic
            '    Me.DgvGroup.Columns(Me.ColGroup_Desc.Index).SortMode = DataGridViewColumnSortMode.Automatic
        End Try
    End Sub

    Private Function GetIndexItemGroup(ByVal vPrimary As Double, _
                    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            '''pastikan tidak ada code yang sama
            For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                RowDGV = Me.DgvGroup.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColGroup_Primary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BackgroundWorker4_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker4_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.DgvGroup.Rows.RemoveAt(EraseRowIndexProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

    Private Sub CboGrouping_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboGrouping.SelectedIndexChanged
        Try
            If Me.CboGrouping.SelectedIndex = 0 Then
                Me.CboGroup1.Enabled = True
                Me.CboGroup2.Enabled = True
            ElseIf Me.CboGrouping.SelectedIndex = 1 Then
                Me.CboGroup1.Enabled = True
                Me.CboGroup2.Enabled = False
            ElseIf Me.CboGrouping.SelectedIndex = 2 Then
                Me.CboGroup1.Enabled = False
                Me.CboGroup2.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
        End Try
    End Sub

    Private Sub CboGroup1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboGroup1.GotFocus
        Try
            Me.BufferCbo = Me.CboGroup1.SelectedIndex
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
        End Try
    End Sub

    Private Sub CboGroup2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboGroup2.GotFocus
        Try
            Me.BufferCbo = Me.CboGroup1.SelectedIndex
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
        End Try
    End Sub

    Private Function CboSetting() As Boolean
        Try
            If Me.CboGroup1.SelectedIndex <> -1 And Me.CboGroup2.SelectedIndex <> -1 Then
                If Me.CboGroup1.SelectedIndex >= Me.CboGroup2.SelectedIndex And Me.CboGroup2.Enabled = True Then
                    MsgBox("Level group 2 must be smaller than group 1", MsgBoxStyle.Information, "Information")
                    Return False
                Else
                    Return True
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
        End Try
    End Function

End Class