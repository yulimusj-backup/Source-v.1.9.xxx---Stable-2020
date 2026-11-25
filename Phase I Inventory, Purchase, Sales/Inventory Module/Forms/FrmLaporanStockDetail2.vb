Option Explicit On
Option Strict Off

Imports System.Data

Public Class FrmLaporanStockDetail2

    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public LoadForMonthlyClosing As Boolean = False
    Public DaysForLoadPartialStockCard As Integer = 30
    Public Scope_Index As EnumScope = EnumScope.Both

    Private GroupBy_Index As EnumGroupBy
    Private LevelGroup As EnumLevel
    Private LevelIndex As EnumLevelIndex
    Private ExpiryMin As EnumExpiryMinDate
    Private ExpiryMax As EnumExpiryMaxDate
    Private ShowSubTotal As EnumShowSubTotal
    Private ShowGrandTotal As EnumShowGrandTotal
    Private ShowStockAsOff As EnumShowStockAsOff

    Private Buffer As String = ""
    Private objGroup As New Core_Module.ClsMasterProductGroup
    Private ObjCategory As New Core_Module.ClsMasterCategory
    Private ObjProduct As New Core_Module.clsMasterProduct
    Private ObjLocation As New Core_Module.clsMasterStorageArea
    Private ObjCD As New Core_Module.ClsCekData

    Private EraseRowindexLocation As Integer = 0
    Private EraseRowIndexProduct As Integer = 0
    Private EraseRowIndexBatchNo As Integer = 0

    Dim ReportDetail2 As New CrpLaporanStockDetail2_A
    Dim ReportDetail3 As New CrpLaporanStockDetail2_C
    Dim ReportDetail1 As New CrpLaporanStockDetail2_B

    Dim ExportToBeginingBalance As Boolean = False
    Dim PrimaryStockInOut As Double = 40161
    Dim TargetConstring As String = "server=localhost;user id=root;Password=#SMeSS_p4ss*897%#;persist security info=True;database=smess_cmsi_temporary;port=3307;default command timeout=19000000"
    Dim ObjLicense As New Core_Module.ClsLicense

    Dim DS As New DseLaporanStockDetail2
    Dim TBLCetak As DseLaporanStockDetail2.TblCetakDataTable = DS.TblCetak

    Public Enum EnumGroupBy
        No_Grouping
        By_Location
        By_Product_Group
        By_Brand
        By_Sub_Brand
        Group_By_Category
    End Enum

    Public Enum EnumScope
        Goods_Only
        Service_Only
        Both
    End Enum

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

    Public Enum EnumExpiryMinDate
        No
        Yes
    End Enum

    Public Enum EnumExpiryMaxDate
        No
        Yes
    End Enum

    Public Enum EnumShowSubTotal
        No
        Yes
    End Enum

    Public Enum EnumShowGrandTotal
        No
        Yes
    End Enum

    Public Enum EnumShowStockAsOff
        No
        Yes
    End Enum

    Public Enum EnumPilih
        No
        Yes
    End Enum

    Private Sub FrmLaporanStockDetail2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.CboFormat.SelectedIndex = 0
            Me.CboScope.SelectedIndex = 0
            Me.CboPilihGroupLevel.SelectedIndex = 0
            Me.DtpMinExpiryDate.Value = Now
            Me.DtpMaxExpiryDate.Value = Now
            Me.DtpMaxExpiryDate.Enabled = False
            Me.DtpMinExpiryDate.Enabled = False
            Me.DtpStockasOff.Enabled = False
            Me.DgvLocation.AllowUserToAddRows = True
            Me.DgvGroup.AllowUserToAddRows = True
            Me.DgvBatchNo.AllowUserToAddRows = True
            Me.DgvLocation.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvGroup.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvBatchNo.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.CancelButton = Me.btnClose
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

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

    Private Sub CboFormat_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboFormat.SelectedIndexChanged
        Try
            If CboFormat.SelectedIndex = 0 Then
                Me.GroupBy_Index = EnumGroupBy.No_Grouping
                Me.ChkShowSubTotal.Checked = False
            ElseIf CboFormat.SelectedIndex = 1 Then
                Me.GroupBy_Index = EnumGroupBy.By_Location
            ElseIf CboFormat.SelectedIndex = 2 Then
                Me.GroupBy_Index = EnumGroupBy.By_Product_Group
            ElseIf CboFormat.SelectedIndex = 3 Then
                Me.GroupBy_Index = EnumGroupBy.By_Brand
            ElseIf CboFormat.SelectedIndex = 4 Then
                Me.GroupBy_Index = EnumGroupBy.By_Sub_Brand
            ElseIf CboFormat.SelectedIndex = 5 Then
                Me.GroupBy_Index = EnumGroupBy.Group_By_Category
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboScope_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboScope.SelectedIndexChanged
        Try
            If Me.CboScope.SelectedIndex = 0 Then
                Me.Scope_Index = EnumScope.Goods_Only
            ElseIf Me.CboScope.SelectedIndex = 1 Then
                Me.Scope_Index = EnumScope.Service_Only
            ElseIf Me.CboScope.SelectedIndex = 2 Then
                Me.Scope_Index = EnumScope.Both
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkMaxExpiryDate_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkMaxExpiryDate.CheckedChanged
        Try
            If Me.ChkMaxExpiryDate.Checked = False Then
                Me.ExpiryMax = EnumExpiryMaxDate.No
            ElseIf Me.ChkMaxExpiryDate.Checked = True Then
                Me.ExpiryMax = EnumExpiryMaxDate.Yes
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkMaxExpiryDate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkMaxExpiryDate.Click
        Try
            If Me.ChkMaxExpiryDate.Checked = True Then
                Me.DtpMaxExpiryDate.Enabled = True
            Else
                Me.DtpMaxExpiryDate.Enabled = False
                Me.DtpMaxExpiryDate.Value = Now
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkMinExpiryDate_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkMinExpiryDate.CheckedChanged
        Try
            If Me.ChkMinExpiryDate.Checked = False Then
                Me.ExpiryMin = EnumExpiryMinDate.No
            ElseIf Me.ChkMinExpiryDate.Checked = True Then
                Me.ExpiryMin = EnumExpiryMinDate.Yes
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkMinExpiryDate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkMinExpiryDate.Click
        Try
            If Me.ChkMinExpiryDate.Checked = True Then
                Me.DtpMinExpiryDate.Enabled = True
            Else
                Me.DtpMinExpiryDate.Enabled = False
                Me.DtpMinExpiryDate.Value = Now
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkShowGrandTotal_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkShowGrandTotal.CheckedChanged
        Try
            If Me.ChkShowGrandTotal.Checked = False Then
                Me.ShowGrandTotal = EnumShowGrandTotal.No
            ElseIf Me.ChkShowGrandTotal.Checked = True Then
                Me.ShowGrandTotal = EnumShowGrandTotal.Yes
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkShowSubTotal_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkShowSubTotal.CheckedChanged
        Try
            If Me.GroupBy_Index <> EnumGroupBy.No_Grouping Then
                If Me.ChkShowSubTotal.Checked = False Then
                    Me.ShowSubTotal = EnumShowSubTotal.No
                ElseIf Me.ChkShowSubTotal.Checked = True Then
                    Me.ShowSubTotal = EnumShowSubTotal.Yes
                End If
            Else
                Me.ChkShowSubTotal.Checked = False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvLocation_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvLocation.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColLocation_Code.Index Then
                    '''set header fungsi digunakan ketika edit code 
                    Me.DgvLocation.Columns(Me.ColLocation_Code.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.DgvLocation.Columns(Me.ColLocation_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    '''simpan nilai awal di buffer
                    Me.Buffer = Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColLocation_Code.Index Then
                    '''set header fungsi digunakan ketika edit code 
                    Me.DgvLocation.Columns(Me.ColLocation_Code.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.DgvLocation.Columns(Me.ColLocation_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.Buffer = Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvLocation_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvLocation.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvLocation.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryLocation As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColLocation_Code.Index Then
                        Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjLocation.ApakahSudahAda(Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value, Core_Module.clsMasterStorageArea.enumTransit.All) = True Then
                            PrimaryLocation = Me.ObjLocation.GetPrimary(Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemLocation(PrimaryLocation, e.RowIndex)
                            If index < 0 Then
                                Me.DgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = PrimaryLocation
                                Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = Me.ObjLocation.GetStorageAreaCode(PrimaryLocation)
                                Me.DgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = Me.ObjLocation.GetDescription(PrimaryLocation)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowindexLocation = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvLocation)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmCariMasterStorageArea
                            ij.PencarianTransit = Core_Module.clsMasterStorageArea.enumTransit.All
                            ij.txtCari.Text = Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryStorageArea <> -1 Then

                                index = Me.GetIndexItemLocation(ij.PrimaryStorageArea, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = ij.PrimaryStorageArea
                                    Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = Me.ObjLocation.GetStorageAreaCode(ij.PrimaryStorageArea)
                                    Me.DgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = Me.ObjLocation.GetDescription(ij.PrimaryStorageArea)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowindexLocation = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvLocation)
                                        Exit Sub
                                    Else
                                        Me.DgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = ij.PrimaryStorageArea
                                        Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = Me.ObjLocation.GetStorageAreaCode(ij.PrimaryStorageArea)
                                        Me.DgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = Me.ObjLocation.GetDescription(ij.PrimaryStorageArea)
                                    End If
                                End If
                            ElseIf ij.PrimaryStorageArea = -1 Then
                                If Me.DgvLocation.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowindexLocation = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvLocation)
                                    Exit Sub
                                End If
                                If Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowindexLocation = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvLocation)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvLocation)
                    End If
                Else
                    If Me.DgvLocation.Rows.Count - 1 <> e.RowIndex Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowindexLocation = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        Else
                            Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvLocation)
                        Exit Sub
                    End If
                    If Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value <> "" Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowindexLocation = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        Else
                            Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvLocation)
                        Exit Sub
                    End If
                End If
                'Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DgvLocation.Columns(Me.ColLocation_Code.Index).SortMode = DataGridViewColumnSortMode.Automatic
            Me.DgvLocation.Columns(Me.ColLocation_Desc.Index).SortMode = DataGridViewColumnSortMode.Automatic
        End Try
    End Sub

    Private Function GetIndexItemLocation(ByVal vPrimary As Double, _
          ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            '''pastikan tidak ada code yang sama
            For ij As Integer = 0 To Me.DgvLocation.Rows.Count - 2
                RowDGV = Me.DgvLocation.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColLocation_Primary.Index).Value = vPrimary Then
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

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.DgvLocation.Rows.RemoveAt(EraseRowindexLocation)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvProductGroup_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvGroup.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColGroup_Code.Index Then
                    '''set header fungsi digunakan ketika edit code 
                    'Me.DgvGroup.Columns(Me.ColGroup_Code.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvGroup.Columns(Me.ColGroup_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    '''simpan nilai awal di buffer
                    Me.Buffer = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColGroup_Code.Index Then
                    '''set header fungsi digunakan ketika edit code 
                    'Me.DgvGroup.Columns(Me.ColGroup_Code.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvGroup.Columns(Me.ColGroup_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.Buffer = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value
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
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndexProduct = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
                                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                                Me.EraseRowIndexProduct = e.RowIndex
                                                Me.BackgroundWorker2.RunWorkerAsync()
                                            Else
                                                Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndexProduct = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                                        Exit Sub
                                    End If
                                    If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndexProduct = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                        Exit Sub
                    End If
                    If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndexProduct = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvGroup)
                                    Exit Sub
                                End If
                                If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                        Exit Sub
                    End If
                    If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
            Me.DgvGroup.Columns(Me.ColGroup_Code.Index).SortMode = DataGridViewColumnSortMode.Automatic
            Me.DgvGroup.Columns(Me.ColGroup_Desc.Index).SortMode = DataGridViewColumnSortMode.Automatic
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
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndexProduct = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvGroup)
                                    Exit Sub
                                End If
                                If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                        Exit Sub
                    End If
                    If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
            Me.DgvGroup.Columns(Me.ColGroup_Code.Index).SortMode = DataGridViewColumnSortMode.Automatic
            Me.DgvGroup.Columns(Me.ColGroup_Desc.Index).SortMode = DataGridViewColumnSortMode.Automatic
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

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.DgvGroup.Rows.RemoveAt(EraseRowIndexProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvBatchNo_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvBatchNo.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColBatch_BatchNo.Index Then
                Me.Buffer = Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvBatchNo_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvBatchNo.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvBatchNo.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim BatchNo As String = ""
        Dim RowIdx As Integer = e.RowIndex
        Dim data As String
        Try
            With row
                If IsNothing(Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColBatch_BatchNo.Index Then
                        'set ke huruf besar dan hapus spasi
                        Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim

                        If Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value.ToString.StartsWith("'") = True Then
                            If Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 0 Then
                                Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(1, Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length - 1)
                            End If
                        End If

                        If Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value.ToString.EndsWith("'") = True Then
                            If Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 0 Then
                                Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length - 2)
                            End If
                        End If

                        index = Me.GetIndexItemBatchNo(Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value, RowIdx)
                        If index >= 0 Then
                            If e.RowIndex <> index Then
                                MsgBox("Duplication Batch No, please input another Batch No.", MsgBoxStyle.Information, "Information")
                                Me.DgvBatchNo.Item(Me.ColBatch_BatchNo.Index, e.RowIndex).Value = ""
                                If String.IsNullOrEmpty(Me.Buffer) = True Then
                                    Me.EraseRowIndexBatchNo = e.RowIndex
                                    Me.BackgroundWorker3.RunWorkerAsync()
                                Else
                                    Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                End If
                                Exit Sub
                            End If
                        End If
                    End If
                Else

                    If Me.DgvBatchNo.Rows.Count - 1 <> e.RowIndex Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexBatchNo = e.RowIndex
                            Me.BackgroundWorker3.RunWorkerAsync()
                        Else
                            Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        Exit Sub
                    End If

                    If Me.DgvBatchNo.Item(Me.ColBatch_BatchNo.Index, e.RowIndex).Value <> "" Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexBatchNo = e.RowIndex
                            Me.BackgroundWorker3.RunWorkerAsync()
                        Else
                            Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        Exit Sub
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemBatchNo(ByVal vBatchNo As String, _
         ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            '''pastikan tidak ada code yang sama
            For ij As Integer = 0 To Me.DgvBatchNo.Rows.Count - 2
                RowDGV = Me.DgvBatchNo.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColBatch_BatchNo.Index).Value = vBatchNo Then
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

    Private Sub BackgroundWorker3_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Try
            Me.DgvBatchNo.Rows.RemoveAt(EraseRowIndexBatchNo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DisableControls(ByVal vDisable As Boolean)
        Try
            If vDisable = True Then
                CboFormat.Enabled = False
                btnFilter.Enabled = False
                btnClose.Enabled = False
                ChkShowSubTotal.Enabled = False
                ChkShowGrandTotal.Enabled = False
                ChkMinExpiryDate.Enabled = False
                DtpMinExpiryDate.Enabled = False
                ChkMaxExpiryDate.Enabled = False
                DtpMaxExpiryDate.Enabled = False
                ChkStockAsOFF.Enabled = False
                DtpStockasOff.Enabled = False
                chkUseWarehouse.Enabled = False
                Me.DgvBatchNo.Enabled = False
                Me.DgvGroup.Enabled = False
                Me.DgvLocation.Enabled = False
                Me.CboPilihGroupLevel.Enabled = False
                crv1.Enabled = False
            Else
                CboFormat.Enabled = True
                btnFilter.Enabled = True
                btnClose.Enabled = True
                ChkShowSubTotal.Enabled = True
                ChkShowGrandTotal.Enabled = True
                ChkMinExpiryDate.Enabled = True
                DtpMinExpiryDate.Enabled = True
                ChkMaxExpiryDate.Enabled = True
                DtpMaxExpiryDate.Enabled = True
                ChkStockAsOFF.Enabled = True
                DtpStockasOff.Enabled = True
                chkUseWarehouse.Enabled = True
                Me.DgvBatchNo.Enabled = True
                Me.DgvGroup.Enabled = True
                Me.CboPilihGroupLevel.Enabled = True
                Me.DgvLocation.Enabled = True
                crv1.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub LoadPartialLocationTransfer(ByVal vTblLTF As DseLaporanStockDetail2.DataTable6DataTable, ByVal vPeriodFrom As Date)
        Dim vRowLTF As DseLaporanStockDetail2.DataTable6Row

        Dim idx0 As Integer = 0
        Dim idx1 As Integer = 0
        Dim primaryProductMin As Double = 9999999999
        Dim primaryProductMax As Double = 0
        Dim primaryLocationMin As Double = 9999999999
        Dim primaryLocationMax As Double = 0
        Dim dtaProduct As New DseLaporanStockDetail2TableAdapters.ProductDataTableAdapter
        Dim tblProduct As New DseLaporanStockDetail2.ProductDataDataTable
        Dim rowProduct As DseLaporanStockDetail2.ProductDataRow
        Dim dtaLocation As New DseLaporanStockDetail2TableAdapters.LocationDataTableAdapter
        Dim tblLocation As New DseLaporanStockDetail2.LocationDataDataTable
        Dim rowLocation As DseLaporanStockDetail2.LocationDataRow

        Dim dtaLTF_lite As New DseLaporanStockDetail2TableAdapters.DataTable6_liteTableAdapter
        Dim tblLTF_lite As New DseLaporanStockDetail2.DataTable6_liteDataTable
        Dim rowLTF_lite As DseLaporanStockDetail2.DataTable6_liteRow
        Try
            dtaLTF_lite.Connection.ConnectionString = Core_Module.GetConstring
            dtaProduct.Connection.ConnectionString = Core_Module.GetConstring
            dtaLocation.Connection.ConnectionString = Core_Module.GetConstring
            While vPeriodFrom.Date <= Now.Date
                Dim vPeriodTo As Date = vPeriodFrom.AddDays(DaysForLoadPartialStockCard)
                If vPeriodTo >= Now.Date Then
                    vPeriodTo = #1/1/2200#
                End If
                Me.Label4.Text = "Loading Stock Transfer: " & Format(vPeriodFrom, "dd MMM yyyy") & " to " & Format(vPeriodFrom.AddDays(DaysForLoadPartialStockCard), "dd MMM yyyy")
                dtaLTF_lite.Fill(tblLTF_lite, vPeriodFrom, vPeriodTo)
                tblLTF_lite.DefaultView.RowFilter = ""
                tblLTF_lite.DefaultView.Sort = "[Primary Product], [Primary Storage Area], [Primary Storage Area II], [Lot No]"
                vTblLTF.DefaultView.RowFilter = ""
                vTblLTF.DefaultView.Sort = "[Primary Product], [Primary Storage Area], [Primary Storage Area II], [Lot No]"
                idx0 = 0
                idx1 = 0
                While idx0 < tblLTF_lite.DefaultView.Count
                    rowLTF_lite = tblLTF_lite.DefaultView(idx0).Row
                    If primaryProductMin > rowLTF_lite.Primary_Product Then
                        primaryProductMin = rowLTF_lite.Primary_Product
                    End If
                    If primaryProductMax < rowLTF_lite.Primary_Product Then
                        primaryProductMax = rowLTF_lite.Primary_Product
                    End If
                    If primaryLocationMin > rowLTF_lite.Primary_Storage_Area Then
                        primaryLocationMin = rowLTF_lite.Primary_Storage_Area
                    End If
                    If primaryLocationMax < rowLTF_lite.Primary_Storage_Area Then
                        primaryLocationMax = rowLTF_lite.Primary_Storage_Area
                    End If
                    If primaryLocationMin > rowLTF_lite.Primary_Storage_Area_II Then
                        primaryLocationMin = rowLTF_lite.Primary_Storage_Area_II
                    End If
                    If primaryLocationMax < rowLTF_lite.Primary_Storage_Area_II Then
                        primaryLocationMax = rowLTF_lite.Primary_Storage_Area_II
                    End If
                    Dim found As Boolean = False
                    While idx1 < vTblLTF.DefaultView.Count
                        vRowLTF = vTblLTF.DefaultView(idx1).Row
                        If vRowLTF.Primary_Product = rowLTF_lite.Primary_Product Then
                            If vRowLTF.Primary_Storage_Area = rowLTF_lite.Primary_Storage_Area Then
                                If vRowLTF.Primary_Storage_Area_II = rowLTF_lite.Primary_Storage_Area_II Then
                                    If vRowLTF.Lot_No = rowLTF_lite.Lot_No Then
                                        vRowLTF.Qty += rowLTF_lite.Qty
                                        found = True
                                        Exit While
                                    End If
                                ElseIf vRowLTF.Primary_Storage_Area_II > rowLTF_lite.Primary_Storage_Area_II Then
                                    Exit While
                                End If
                            ElseIf vRowLTF.Primary_Storage_Area > rowLTF_lite.Primary_Storage_Area Then
                                Exit While
                            End If
                        ElseIf vRowLTF.Primary_Product > rowLTF_lite.Primary_Product Then
                            Exit While
                        End If
                        idx1 += 1
                    End While
                    If Not found Then
                        vRowLTF = vTblLTF.NewDataTable6Row
                        vRowLTF.Primary_Product = rowLTF_lite.Primary_Product
                        vRowLTF.Lot_No = rowLTF_lite.Lot_No
                        vRowLTF.Expiry_Date = rowLTF_lite.Expiry_Date
                        vRowLTF.Primary_Storage_Area = rowLTF_lite.Primary_Storage_Area
                        vRowLTF.Primary_Storage_Area_II = rowLTF_lite.Primary_Storage_Area_II
                        vRowLTF.Qty = rowLTF_lite.Qty
                        vTblLTF.Rows.Add(vRowLTF)
                    End If
                    idx0 += 1
                End While
                vPeriodFrom = vPeriodFrom.AddDays(DaysForLoadPartialStockCard + 1)
            End While

            If Me.LoadForMonthlyClosing = False Then
                dtaProduct.Fill(tblProduct, primaryProductMin, primaryProductMax)
                idx0 = 0
                idx1 = 0
                vTblLTF.DefaultView.RowFilter = ""
                vTblLTF.DefaultView.Sort = "[Primary Product]"
                tblProduct.DefaultView.RowFilter = ""
                tblProduct.DefaultView.Sort = "[Product Primary]"
                While idx0 < vTblLTF.DefaultView.Count
                    vRowLTF = vTblLTF.DefaultView(idx0).Row
                    Me.Label4.Text = "Update product data " & idx0 + 1 & " of " & vTblLTF.DefaultView.Count
                    While idx1 < tblProduct.DefaultView.Count
                        rowProduct = tblProduct.DefaultView(idx1).Row
                        If rowProduct.Product_Primary < vRowLTF.Primary_Product Then
                            idx1 += 1
                        ElseIf rowProduct.Product_Primary = vRowLTF.Primary_Product Then
                            With vRowLTF
                                .Product_Code = rowProduct.Product_Code
                                .Description = rowProduct.Product_Desc
                                ._Decimal = rowProduct._Decimal
                                .Batch_No_Controlled = rowProduct.Batch_No_Controlled
                                .Location_Controlled = rowProduct.Location_Controlled
                                .Expiry_Date_Controlled = rowProduct.Expiry_Date_Controlled
                                .Type = rowProduct.Type
                                .Primary_Unit = rowProduct.Primary_Unit
                                .Unit_Code = rowProduct.Unit_Code
                                .Categorey_Code = rowProduct.Categorey_Code
                                .Category_Desc = rowProduct.Categrory_Desc
                                .Sub_Brand_Code = rowProduct.Sub_Brand_Code
                                .Sub_Brand_Desc = rowProduct.Sub_Brand_Desc
                                .Brand_Code = rowProduct.Brand_Code
                                .Brand_Desc = rowProduct.Brand_Desc
                                .Group_Code = rowProduct.Group_Code
                                .Group_Desc = rowProduct.Group_Desc
                                .Category_Primary = rowProduct.Category_Primary
                                .Categorey_Code = rowProduct.Categorey_Code
                                .Sub_Brand_Primary = rowProduct.Sub_Brand_Primary
                                .Brand_Primary = rowProduct.Brand_Primary
                                .Group_Primary = rowProduct.Group_Primary
                            End With
                            Exit While
                        ElseIf rowProduct.Product_Primary > vRowLTF.Primary_Product Then
                            Exit While
                        End If
                    End While
                    idx0 += 1
                End While

                dtaLocation.Fill(tblLocation, primaryLocationMin, primaryLocationMax)
                idx0 = 0
                idx1 = 0
                vTblLTF.DefaultView.RowFilter = ""
                vTblLTF.DefaultView.Sort = "[Primary Storage Area]"
                tblLocation.DefaultView.RowFilter = ""
                tblLocation.DefaultView.Sort = "[Primary Storage Area]"
                While idx0 < vTblLTF.DefaultView.Count
                    vRowLTF = vTblLTF.DefaultView(idx0).Row
                    Me.Label4.Text = "Update location data " & idx1 + 1 & " of " & vTblLTF.DefaultView.Count
                    While idx1 < tblLocation.DefaultView.Count
                        rowLocation = tblLocation.DefaultView(idx1).Row
                        If rowLocation.Primary_Storage_Area < vRowLTF.Primary_Storage_Area Then
                            idx1 += 1
                        ElseIf rowLocation.Primary_Storage_Area = vRowLTF.Primary_Storage_Area Then
                            With vRowLTF
                                .Storage_Area_Code = rowLocation.Storage_Area_Code
                                .Storage_Area_Desc = rowLocation.Storage_Area_Desc
                                .Warehouse_Code = rowLocation.Warehouse_Code
                                .Warehouse_Desc = rowLocation.Warehouse_Desc
                            End With
                            Exit While
                        ElseIf rowLocation.Primary_Storage_Area > vRowLTF.Primary_Storage_Area Then
                            Exit While
                        End If
                    End While
                    idx0 += 1
                End While

                idx0 = 0
                idx1 = 0
                vTblLTF.DefaultView.RowFilter = ""
                vTblLTF.DefaultView.Sort = "[Primary Storage Area II]"
                While idx0 < vTblLTF.DefaultView.Count
                    vRowLTF = vTblLTF.DefaultView(idx0).Row
                    Me.Label4.Text = "Update location II data " & idx1 + 1 & " of " & vTblLTF.DefaultView.Count
                    While idx1 < tblLocation.DefaultView.Count
                        rowLocation = tblLocation.DefaultView(idx1).Row
                        If rowLocation.Primary_Storage_Area < vRowLTF.Primary_Storage_Area_II Then
                            idx1 += 1
                        ElseIf rowLocation.Primary_Storage_Area = vRowLTF.Primary_Storage_Area_II Then
                            With vRowLTF
                                .Storage_Area_Code_II = rowLocation.Storage_Area_Code
                                .Storage_Area_Desc_II = rowLocation.Storage_Area_Desc
                                .Warehouse_Code_II = rowLocation.Warehouse_Code
                                .Warehouse_Desc_II = rowLocation.Warehouse_Desc
                            End With
                            Exit While
                        ElseIf rowLocation.Primary_Storage_Area > vRowLTF.Primary_Storage_Area_II Then
                            Exit While
                        End If
                    End While
                    idx0 += 1
                End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' Load data stock card dari mulai tanggal yang ditentukan sampai dengan 1-1-2200 per jumlah hari yang ditentukan (Me.DaysForLoadPartialStockCard)
    ''' Load sebagian untuk menghindari exception out of memory pada customer yang datanya sangat banyak.
    ''' </summary>
    Sub LoadPartialStockCard(ByVal vTblSC As DseLaporanStockDetail2.DataTable5DataTable, ByVal vPeriodFrom As Date, ByVal vIntStockTaking As Integer)
        Dim vRowSC As DseLaporanStockDetail2.DataTable5Row

        Dim dtaSC_lite As New DseLaporanStockDetail2TableAdapters.DataTable5_liteTableAdapter
        Dim tblSC_lite As New DseLaporanStockDetail2.DataTable5_liteDataTable
        Dim rowSC_lite As DseLaporanStockDetail2.DataTable5_liteRow

        Dim idx0 As Integer = 0
        Dim idx1 As Integer = 0
        Dim primaryProductMin As Double = 9999999999
        Dim primaryProductMax As Double = 0
        Dim primaryLocationMin As Double = 9999999999
        Dim primaryLocationMax As Double = 0
        Dim dtaProduct As New DseLaporanStockDetail2TableAdapters.ProductDataTableAdapter
        Dim tblProduct As New DseLaporanStockDetail2.ProductDataDataTable
        Dim rowProduct As DseLaporanStockDetail2.ProductDataRow
        Dim dtaLocation As New DseLaporanStockDetail2TableAdapters.LocationDataTableAdapter
        Dim tblLocation As New DseLaporanStockDetail2.LocationDataDataTable
        Dim rowLocation As DseLaporanStockDetail2.LocationDataRow
        Try
            dtaSC_lite.Connection.ConnectionString = Core_Module.GetConstring
            dtaProduct.Connection.ConnectionString = Core_Module.GetConstring
            dtaLocation.Connection.ConnectionString = Core_Module.GetConstring 

            While vPeriodFrom.Date <= Now.Date
                Dim vPeriodTo As Date = vPeriodFrom.AddDays(DaysForLoadPartialStockCard)
                If vPeriodTo >= Now.Date Then
                    vPeriodTo = #1/1/2200#
                End If
                Me.Label4.Text = "Loading Stock Card: " & Format(vPeriodFrom, "dd MMM yyyy") & " to " & Format(vPeriodFrom.AddDays(DaysForLoadPartialStockCard), "dd MMM yyyy")
                dtaSC_lite.FillBy_Period(tblSC_lite, vPeriodFrom, vPeriodTo, vIntStockTaking)
                tblSC_lite.DefaultView.RowFilter = ""
                tblSC_lite.DefaultView.Sort = "[Primary Product], [Primary Storage Area], [Lot No]"
                vTblSC.DefaultView.RowFilter = ""
                vTblSC.DefaultView.Sort = "[Primary Product], [Primary Storage Area], [Lot No]"
                idx0 = 0
                idx1 = 0
                While idx0 < tblSC_lite.DefaultView.Count
                    rowSC_lite = tblSC_lite.DefaultView(idx0).Row
                    If primaryProductMin > rowSC_lite.Primary_Product Then
                        primaryProductMin = rowSC_lite.Primary_Product
                    End If
                    If primaryProductMax < rowSC_lite.Primary_Product Then
                        primaryProductMax = rowSC_lite.Primary_Product
                    End If
                    If primaryLocationMin > rowSC_lite.Primary_Storage_Area Then
                        primaryLocationMin = rowSC_lite.Primary_Storage_Area
                    End If
                    If primaryLocationMax < rowSC_lite.Primary_Storage_Area Then
                        primaryLocationMax = rowSC_lite.Primary_Storage_Area
                    End If
                    Dim found As Boolean = False
                    While idx1 < vTblSC.DefaultView.Count
                        vRowSC = vTblSC.DefaultView(idx1).Row
                        If vRowSC.Primary_Product = rowSC_lite.Primary_Product Then
                            If vRowSC.Primary_Storage_Area = rowSC_lite.Primary_Storage_Area Then
                                If vRowSC.Lot_No = rowSC_lite.Lot_No Then
                                    vRowSC.Total_In += rowSC_lite.Total_In
                                    vRowSC._Total__Out_ += rowSC_lite._Total__Out_
                                    found = True
                                    Exit While
                                End If
                            ElseIf vRowSC.Primary_Storage_Area > rowSC_lite.Primary_Storage_Area Then
                                Exit While
                            End If
                        ElseIf vRowSC.Primary_Product > rowSC_lite.Primary_Product Then
                            Exit While
                        End If
                        idx1 += 1
                    End While
                    If Not found Then
                        vRowSC = vTblSC.NewDataTable5Row
                        vRowSC.Primary_Product = rowSC_lite.Primary_Product
                        vRowSC.Lot_No = rowSC_lite.Lot_No
                        vRowSC.Expiry_Date = rowSC_lite.Expiry
                        vRowSC.Primary_Storage_Area = rowSC_lite.Primary_Storage_Area
                        vRowSC.Total_In = rowSC_lite.Total_In
                        vRowSC._Total__Out_ = rowSC_lite._Total__Out_
                        vTblSC.Rows.Add(vRowSC)
                    End If
                    idx0 += 1
                End While
                vPeriodFrom = vPeriodFrom.AddDays(DaysForLoadPartialStockCard + 1)
            End While

            If Me.LoadForMonthlyClosing = False Then
                dtaProduct.Fill(tblProduct, primaryProductMin, primaryProductMax)
                idx0 = 0
                idx1 = 0
                vTblSC.DefaultView.RowFilter = ""
                vTblSC.DefaultView.Sort = "[Primary Product]"
                tblProduct.DefaultView.RowFilter = ""
                tblProduct.DefaultView.Sort = "[Product Primary]"
                While idx0 < vTblSC.DefaultView.Count
                    vRowSC = vTblSC.DefaultView(idx0).Row
                    Me.Label4.Text = "Update product data " & idx0 + 1 & " of " & vTblSC.DefaultView.Count
                    While idx1 < tblProduct.DefaultView.Count
                        rowProduct = tblProduct.DefaultView(idx1).Row
                        If rowProduct.Product_Primary < vRowSC.Primary_Product Then
                            idx1 += 1
                        ElseIf rowProduct.Product_Primary = vRowSC.Primary_Product Then
                            With vRowSC
                                .Product_Code = rowProduct.Product_Code
                                .Description = rowProduct.Product_Desc
                                ._Decimal = rowProduct._Decimal
                                .Batch_No_Controlled = rowProduct.Batch_No_Controlled
                                .Location_Controlled = rowProduct.Location_Controlled
                                .Expiry_Date_Controlled = rowProduct.Expiry_Date_Controlled
                                .Type = rowProduct.Type
                                .Primary_Unit = rowProduct.Primary_Unit
                                .Unit_Code = rowProduct.Unit_Code
                                .Category_Primary = rowProduct.Category_Primary
                                .Categorey_Code = rowProduct.Categorey_Code
                                .Categrory_Desc = rowProduct.Categrory_Desc
                                .Sub_Brand_Primary = rowProduct.Sub_Brand_Primary
                                .Sub_Brand_Code = rowProduct.Sub_Brand_Code
                                .Sub_Brand_Desc = rowProduct.Sub_Brand_Desc
                                .Brand_Primary = rowProduct.Brand_Primary
                                .Brand_Code = rowProduct.Brand_Code
                                .Brand_Desc = rowProduct.Brand_Desc
                                .Group_Primary = rowProduct.Group_Primary
                                .Group_Code = rowProduct.Group_Code
                                .Group_Desc = rowProduct.Group_Desc
                            End With
                            Exit While
                        ElseIf rowProduct.Product_Primary > vRowSC.Primary_Product Then
                            Exit While
                        End If
                    End While
                    idx0 += 1
                End While

                dtaLocation.Fill(tblLocation, primaryLocationMin, primaryLocationMax)
                idx0 = 0
                idx1 = 0
                vTblSC.DefaultView.RowFilter = ""
                vTblSC.DefaultView.Sort = "[Primary Storage Area]"
                tblLocation.DefaultView.RowFilter = ""
                tblLocation.DefaultView.Sort = "[Primary Storage Area]"
                While idx0 < vTblSC.DefaultView.Count
                    vRowSC = vTblSC.DefaultView(idx0).Row
                    Me.Label4.Text = "Update location data " & idx1 + 1 & " of " & vTblSC.DefaultView.Count
                    While idx1 < tblLocation.DefaultView.Count
                        rowLocation = tblLocation.DefaultView(idx1).Row
                        If rowLocation.Primary_Storage_Area < vRowSC.Primary_Storage_Area Then
                            idx1 += 1
                        ElseIf rowLocation.Primary_Storage_Area = vRowSC.Primary_Storage_Area Then
                            With vRowSC
                                .Storage_Area_Code = rowLocation.Storage_Area_Code
                                .Storage_Area_Desc = rowLocation.Storage_Area_Desc
                                .Warehouse_Code = rowLocation.Warehouse_Code
                                .Warehouse_Desc = rowLocation.Warehouse_Desc
                            End With
                            Exit While
                        ElseIf rowLocation.Primary_Storage_Area > vRowSC.Primary_Storage_Area Then
                            Exit While
                        End If
                    End While
                    idx0 += 1
                End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function CetakanStockDetail() As Boolean
        Dim DTA As New DseLaporanStockDetail2TableAdapters.DataTable1TableAdapter
        Dim TBL As DseLaporanStockDetail2.DataTable1DataTable = DS.DataTable1
        Dim Row As DseLaporanStockDetail2.DataTable1Row

        Me.TBLCetak.Clear()
        Dim ROWCetak, rowDummy As DseLaporanStockDetail2.TblCetakRow
        Dim tblDummy As New DseLaporanStockDetail2.TblCetakDataTable

        Dim ParamPerGrouping As String = ""
        Dim MinDate As Date
        Dim MaxDate As Date
        Dim PrimaryMax As Double = -1
        Dim PrimaryMin As Double = -1
        Dim Filter01 As String = ""
        Dim Filter01_II As String = ""
        Dim Filter02 As String = ""
        Dim Filter03 As String = ""
        Dim Filter04 As String = ""
        Dim ParamGrandTotal As String = ""
        Dim DV1 As DataRowView
        Dim DTAKonv As New DseLaporanStockDetail2TableAdapters.TblKonversiTableAdapter
        Dim TBLKonv As New DseLaporanStockDetail2.TblKonversiDataTable
        Dim rowKonv As DseLaporanStockDetail2.TblKonversiRow
        Dim ListPrimaryMinus As New List(Of Double)
        Dim IdxAcuan As Integer = 0

        Dim idx0 As Integer = 0
        Dim idx1 As Integer = 0
        Dim Idx As Double = 0

        Try
            Label4.Text = "Initializing..."
            DisableControls(True)
            Prg1.Maximum = 100
            Prg1.Value = 10

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAKonv.Connection.ConnectionString = Core_Module.GetConstring

            TBLCetak.Rows.Clear()
            MaxDate = Me.DtpMaxExpiryDate.Value
            MinDate = Me.DtpMinExpiryDate.Value
            If ExpiryMax = EnumExpiryMaxDate.No And ExpiryMin = EnumExpiryMinDate.No Then
                DTA.Fill(TBL)
            End If
            TBL.DefaultView.Sort = "Primary"
            Filter01 = "[Primary] = [Primary]"
            Filter01_II = "[Primary] = [Primary]"

            If ExpiryMax = EnumExpiryMaxDate.Yes Or ExpiryMin = EnumExpiryMinDate.Yes Then
                Filter01 &= " AND ([Expiry Date Controlled] = 1)"
                Filter01_II &= " AND ([Expiry Date Controlled] = 1)"
                Filter04 = "[Expiry Date Controlled] = 1"

                If ExpiryMax = EnumExpiryMaxDate.Yes And ExpiryMin = EnumExpiryMinDate.Yes Then
                    DTA.FillBy_MaxAndMinExpiryDate(TBL, MinDate.Date, MaxDate.Date, -1)
                    Filter01 &= " AND ([Expiry Date] >= #" & Format(Me.DtpMinExpiryDate.Value.Date, "MM/dd/yyyy") & "# And [Expiry Date] <= #" & Format(Me.DtpMaxExpiryDate.Value.Date, "MM/dd/yyyy") & "# )"
                    Filter01_II &= " AND ([Expiry Date] >= #" & Format(Me.DtpMinExpiryDate.Value.Date, "MM/dd/yyyy") & "# And [Expiry Date] <= #" & Format(Me.DtpMaxExpiryDate.Value.Date, "MM/dd/yyyy") & "# )"
                    Filter04 &= " AND ([Expiry Date] >= #" & Format(Me.DtpMinExpiryDate.Value.Date, "MM/dd/yyyy") & "# And [Expiry Date] <= #" & Format(Me.DtpMaxExpiryDate.Value.Date, "MM/dd/yyyy") & "# )"
                ElseIf ExpiryMax = EnumExpiryMaxDate.Yes And ExpiryMin = EnumExpiryMinDate.No Then
                    DTA.FillBy_ExpiryMaxDate(TBL, MaxDate.Date, -1)
                    Filter01 &= " AND ([Expiry Date] >= #" & #1/1/1950# & "# And [Expiry Date] <= #" & Format(Me.DtpMaxExpiryDate.Value.Date, "MM/dd/yyyy") & "# )"
                    Filter01_II &= " AND ([Expiry Date] >= #" & #1/1/1950# & "# And [Expiry Date] <= #" & Format(Me.DtpMaxExpiryDate.Value.Date, "MM/dd/yyyy") & "# )"
                    Filter04 &= " AND ([Expiry Date] >= #" & #1/1/1950# & "# And [Expiry Date] <= #" & Format(Me.DtpMaxExpiryDate.Value.Date, "MM/dd/yyyy") & "# )"
                ElseIf ExpiryMin = EnumExpiryMinDate.Yes And ExpiryMax = EnumExpiryMaxDate.No Then
                    DTA.FillBy_ExpiryMinDate(TBL, MinDate.Date, -1)
                    Filter01 &= " AND ([Expiry Date] >= #" & Format(Me.DtpMinExpiryDate.Value.Date, "MM/dd/yyyy") & "# And [Expiry Date] <= #" & #1/1/2100# & "# )"
                    Filter01_II &= " AND ([Expiry Date] >= #" & Format(Me.DtpMinExpiryDate.Value.Date, "MM/dd/yyyy") & "# And [Expiry Date] <= #" & #1/1/2100# & "# )"
                    Filter04 &= " AND ([Expiry Date] >= #" & Format(Me.DtpMinExpiryDate.Value.Date, "MM/dd/yyyy") & "# And [Expiry Date] <= #" & #1/1/2100# & "# )"
                ElseIf ExpiryMax = EnumExpiryMaxDate.No And ExpiryMin = EnumExpiryMinDate.No Then
                End If
            End If

            '2. filter berdasarkan scope
            If TBL.DefaultView.Count > 0 Then
                If Me.Scope_Index = EnumScope.Goods_Only Then
                    Filter01 &= " AND [Type] = " & EnumScope.Goods_Only
                    Filter01_II &= " AND [Type] = " & EnumScope.Goods_Only
                    If Filter04 = "" Then
                        Filter04 = "[Type] = " & EnumScope.Goods_Only
                    Else
                        Filter04 &= " AND [Type] = " & EnumScope.Goods_Only
                    End If
                ElseIf Me.Scope_Index = EnumScope.Service_Only Then
                    Filter01 &= " AND [Type] = " & EnumScope.Service_Only
                    Filter01_II &= " AND [Type] = " & EnumScope.Service_Only
                    If Filter04 = "" Then
                        Filter04 = "[Type] = " & EnumScope.Service_Only
                    Else
                        Filter04 &= " AND [Type] = " & EnumScope.Service_Only
                    End If
                End If
            End If

            '3 filter berdasarkan lokasi
            Me.DgvLocation.AllowUserToAddRows = False
            For i As Integer = 0 To DgvLocation.Rows.Count - 1
                If DgvLocation.Item(Me.ColLocation_Code.Index, i).Value <> "" Then
                    If DgvLocation.Rows.Count > 1 Then
                        If i > 0 Then
                            Filter01 &= " OR [Primary Storage Area] = " & DgvLocation.Item(Me.ColLocation_Primary.Index, i).Value & ""
                            Filter01_II &= " OR ([Primary Storage Area] = " & DgvLocation.Item(Me.ColLocation_Primary.Index, i).Value & " OR [Primary Storage Area II] = " & DgvLocation.Item(Me.ColLocation_Primary.Index, i).Value & ")"
                            Filter04 &= " OR [Primary Location] = " & DgvLocation.Item(Me.ColLocation_Primary.Index, i).Value & ""
                        Else
                            Filter01 &= " AND ([Primary Storage Area] = " & DgvLocation.Item(Me.ColLocation_Primary.Index, i).Value & ""
                            Filter01_II &= " AND (([Primary Storage Area] = " & DgvLocation.Item(Me.ColLocation_Primary.Index, i).Value & " OR [Primary Storage Area II] = " & DgvLocation.Item(Me.ColLocation_Primary.Index, i).Value & ")"
                            If Filter04 = "" Then
                                Filter04 = "([Primary Location] = " & DgvLocation.Item(Me.ColLocation_Primary.Index, i).Value & ""
                            Else
                                Filter04 &= " AND ([Primary Location] = " & DgvLocation.Item(Me.ColLocation_Primary.Index, i).Value & ""
                            End If
                        End If
                    Else
                        Filter01 &= " AND [Primary Storage Area] = " & DgvLocation.Item(Me.ColLocation_Primary.Index, i).Value & ""
                        Filter01_II &= " AND ([Primary Storage Area] = " & DgvLocation.Item(Me.ColLocation_Primary.Index, i).Value & " OR [Primary Storage Area II] = " & DgvLocation.Item(Me.ColLocation_Primary.Index, i).Value & ")"
                        If Filter04 = "" Then
                            Filter04 = "[Primary Location] = " & DgvLocation.Item(Me.ColLocation_Primary.Index, i).Value & ""
                        Else
                            Filter04 &= " AND [Primary Location] = " & DgvLocation.Item(Me.ColLocation_Primary.Index, i).Value & ""
                        End If
                    End If
                End If
                If DgvLocation.Rows.Count > 1 Then
                    If i = DgvLocation.Rows.Count - 1 Then
                        Filter01 &= ")"
                        Filter01_II &= ")"
                        Filter04 &= ")"
                    End If
                End If
            Next
            Me.DgvLocation.AllowUserToAddRows = True

            '4. filter berdasarkan Product Group
            Me.DgvGroup.AllowUserToAddRows = False
            For i As Integer = 0 To DgvGroup.Rows.Count - 1
                If DgvGroup.Item(Me.ColGroup_Code.Index, i).Value <> "" Then
                    If DgvGroup.Rows.Count > 1 Then
                        If i > 0 Then
                            Filter01 &= Me.FilterProductOr() & DgvGroup.Item(Me.ColGroup_Primary.Index, i).Value & ""
                            Filter01_II &= Me.FilterProductOr() & DgvGroup.Item(Me.ColGroup_Primary.Index, i).Value & ""
                            Filter04 &= Me.FilterProductOr() & DgvGroup.Item(Me.ColGroup_Primary.Index, i).Value & ""
                        Else
                            Filter01 &= Me.FilterProductAnd(True) & DgvGroup.Item(Me.ColGroup_Primary.Index, i).Value & ""
                            Filter01_II &= Me.FilterProductAnd(True) & DgvGroup.Item(Me.ColGroup_Primary.Index, i).Value & ""
                            If Filter04 = "" Then
                                Filter04 &= Me.FilterProductAnd(True).Remove(0, 5) & DgvGroup.Item(Me.ColGroup_Primary.Index, i).Value & ""
                            Else
                                Filter04 &= Me.FilterProductAnd(True) & DgvGroup.Item(Me.ColGroup_Primary.Index, i).Value & ""
                            End If
                        End If
                    Else
                        Filter01 &= Me.FilterProductAnd(False) & DgvGroup.Item(Me.ColGroup_Primary.Index, i).Value & ""
                        Filter01_II &= Me.FilterProductAnd(False) & DgvGroup.Item(Me.ColGroup_Primary.Index, i).Value & ""
                        If Filter04 = "" Then
                            Filter04 &= Me.FilterProductAnd(False).Remove(0, 5) & DgvGroup.Item(Me.ColGroup_Primary.Index, i).Value & ""
                        Else
                            Filter04 &= Me.FilterProductAnd(False) & DgvGroup.Item(Me.ColGroup_Primary.Index, i).Value & ""
                        End If
                    End If
                End If
                If DgvGroup.Rows.Count > 1 Then
                    If i = DgvGroup.Rows.Count - 1 Then
                        Filter01 &= ")"
                        Filter01_II &= ")"
                        Filter04 &= ")"
                    End If
                End If
            Next
            Me.DgvGroup.AllowUserToAddRows = True

            '5. filter berdasarkan batch no
            Me.DgvBatchNo.AllowUserToAddRows = False
            For i As Integer = 0 To DgvBatchNo.Rows.Count - 1
                If DgvBatchNo.Item(Me.ColBatch_BatchNo.Index, i).Value <> "" Then
                    If DgvBatchNo.Rows.Count > 1 Then
                        If i > 0 Then
                            Filter01 &= " Or [Lot No] LIKE '" & DgvBatchNo.Item(Me.ColBatch_BatchNo.Index, i).Value & "%'"
                            Filter01_II &= " Or [Lot No] LIKE '" & DgvBatchNo.Item(Me.ColBatch_BatchNo.Index, i).Value & "%'"
                            Filter04 &= " Or [Batch No] LIKE '" & DgvBatchNo.Item(Me.ColBatch_BatchNo.Index, i).Value & "%'"
                        Else
                            Filter01 &= " AND ([Lot No] LIKE '" & DgvBatchNo.Item(Me.ColBatch_BatchNo.Index, i).Value & "%'"
                            Filter01_II &= " AND ([Lot No] LIKE '" & DgvBatchNo.Item(Me.ColBatch_BatchNo.Index, i).Value & "%'"
                            If Filter04 = "" Then
                                Filter04 = "([Batch No] LIKE '" & DgvBatchNo.Item(Me.ColBatch_BatchNo.Index, i).Value & "%'"
                            Else
                                Filter04 &= " AND ([Batch No] LIKE '" & DgvBatchNo.Item(Me.ColBatch_BatchNo.Index, i).Value & "%'"
                            End If
                        End If
                    Else
                        Filter01 &= " AND [Lot No] LIKE '" & DgvBatchNo.Item(Me.ColBatch_BatchNo.Index, i).Value & "%'"
                        Filter01_II &= " AND [Lot No] LIKE '" & DgvBatchNo.Item(Me.ColBatch_BatchNo.Index, i).Value & "%'"
                        If Filter04 = "" Then
                            Filter04 = "[Batch No] LIKE '" & DgvBatchNo.Item(Me.ColBatch_BatchNo.Index, i).Value & "%'"
                        Else
                            Filter04 &= " AND [Batch No] LIKE '" & DgvBatchNo.Item(Me.ColBatch_BatchNo.Index, i).Value & "%'"
                        End If
                    End If
                End If
                If DgvBatchNo.Rows.Count > 1 Then
                    If i = DgvBatchNo.Rows.Count - 1 Then
                        Filter01 &= ")"
                        Filter01_II &= ")"
                        Filter04 &= ")"
                    End If
                End If
            Next
            Me.DgvBatchNo.AllowUserToAddRows = True

            TBL.DefaultView.RowFilter = Nothing
            TBL.DefaultView.RowFilter = Filter01

            ''================================================================================================================================================
            If Me.ShowStockAsOff = EnumPilih.Yes Then
                Dim TBL_ As New DseLaporanStockDetail2.DataTable1DataTable
                Dim Row_, Row_1 As DseLaporanStockDetail2.DataTable1Row
                Dim tblSC As New DseLaporanStockDetail2.DataTable5DataTable
                Dim rowSC As DseLaporanStockDetail2.DataTable5Row
                Dim dtaLTF As New DseLaporanStockDetail2TableAdapters.DataTable6TableAdapter
                Dim tblLTF As New DseLaporanStockDetail2.DataTable6DataTable
                Dim rowLTF As DseLaporanStockDetail2.DataTable6Row
                Dim DTA_ST As New DseLaporanStockDetail2TableAdapters.DataTable7TableAdapter
                Dim TBL_ST As New DseLaporanStockDetail2.DataTable7DataTable
                Dim row_ST As DseLaporanStockDetail2.DataTable7Row

                If Me.LoadForMonthlyClosing = True Then
                    For i As Integer = TBL_.Columns.Count - 1 To 0 Step -1
                        If TBL_.Columns.Item(i).ColumnName = "Primary Product" Or _
                                TBL_.Columns.Item(i).ColumnName = "Lot No" Or _
                                TBL_.Columns.Item(i).ColumnName = "Expiry Date" Or _
                                TBL_.Columns.Item(i).ColumnName = "Primary Storage Area" Or _
                                TBL_.Columns.Item(i).ColumnName = "Qty" Then
                        Else
                            TBL_.Columns.RemoveAt(i)
                        End If
                    Next
                    For i As Integer = TBL.Columns.Count - 1 To 0 Step -1
                        If TBL.Columns.Item(i).ColumnName = "Primary Product" Or _
                                TBL.Columns.Item(i).ColumnName = "Lot No" Or _
                                TBL.Columns.Item(i).ColumnName = "Expiry Date" Or _
                                TBL.Columns.Item(i).ColumnName = "Primary Storage Area" Or _
                                TBL.Columns.Item(i).ColumnName = "Qty" Then
                        Else
                            TBL.Columns.RemoveAt(i)
                        End If
                    Next
                End If

                For Each DV1 In TBL.DefaultView
                    Row = DV1.Row
                    Row_ = TBL_.NewDataTable1Row
                    For i As Integer = 0 To TBL.Columns.Count - 1
                        Row_.Item(i) = Row.Item(i)
                    Next
                    TBL_.Rows.Add(Row_)
                Next
                TBL.Rows.Clear()

                DTA_ST.Connection.ConnectionString = Core_Module.GetConstring
                Me.LoadPartialStockCard(tblSC, Me.DtpStockasOff.Value.Date.AddDays(1), CInt(ClsTransaksiStockCard.EnumJenisTransaksiAsal.Stock_Taking))
                Prg1.Maximum = tblSC.DefaultView.Count
                Prg1.Value = 0
                Prg1.Refresh()
                tblSC.DefaultView.RowFilter = Filter01
                tblSC.DefaultView.Sort = "[Primary Product], [Primary Storage Area], [Lot No]"
                TBL_.DefaultView.RowFilter = ""
                TBL_.DefaultView.Sort = "[Primary Product], [Primary Storage Area], [Lot No]"
                While idx0 < tblSC.DefaultView.Count
                    rowSC = tblSC.DefaultView(idx0).Row
                    Prg1.Value += 1
                    Label4.Text = "Stock Card calculation... " & Prg1.Value.ToString & " of " & Prg1.Maximum.ToString & "."
                    Dim found As Boolean = False
                    While idx1 < TBL_.DefaultView.Count
                        Row_ = TBL_.DefaultView(idx1).Row
                        If Row_.Primary_Product = rowSC.Primary_Product Then
                            If Row_.Primary_Storage_Area = rowSC.Primary_Storage_Area Then
                                If Row_.Lot_No = rowSC.Lot_No Then
                                    Row_.Qty = Row_.Qty + rowSC._Total__Out_ - rowSC.Total_In
                                    If Row_.Qty = 0 Then
                                        TBL_.Rows.Remove(Row_)
                                    End If
                                    found = True
                                    Exit While
                                End If
                            ElseIf Row_.Primary_Storage_Area > rowSC.Primary_Storage_Area Then
                                Exit While
                            End If
                        ElseIf Row_.Primary_Product > rowSC.Primary_Product Then
                            Exit While
                        End If
                        idx1 += 1
                    End While
                    If Not found Then
                        Row_ = TBL_.NewDataTable1Row
                        Row_.Primary_Product = rowSC.Primary_Product
                        Row_.Lot_No = rowSC.Lot_No
                        Row_.Expiry_Date = rowSC.Expiry_Date
                        Row_.Primary_Storage_Area = rowSC.Primary_Storage_Area
                        Row_.Qty = rowSC._Total__Out_ - rowSC.Total_In
                        If Me.LoadForMonthlyClosing = False Then
                            Row_._Decimal = rowSC._Decimal
                            Row_.Batch_No_Controlled = rowSC.Batch_No_Controlled
                            Row_.Location_Controlled = rowSC.Location_Controlled
                            Row_.Expiry_Date_Controlled = rowSC.Expiry_Date_Controlled
                            Row_.Type = rowSC.Type
                            Row_.Primary_Unit = rowSC.Primary_Unit
                            Row_.Category_Primary = rowSC.Category_Primary
                            Row_.Sub_Brand_Primary = rowSC.Sub_Brand_Primary
                            Row_.Brand_Primary = rowSC.Brand_Primary
                            Row_.Group_Primary = rowSC.Group_Primary
                            Row_.Qty_Secondary = -1
                            Row_.Product_Code = rowSC.Product_Code
                            Row_.Description = rowSC.Description
                            Row_.Unit_Code = rowSC.Unit_Code
                            Row_.Categorey_Code = rowSC.Categorey_Code
                            Row_.Category_Desc = rowSC.Categrory_Desc
                            Row_.Sub_Brand_Code = rowSC.Sub_Brand_Code
                            Row_.Sub_Brand_Desc = rowSC.Sub_Brand_Desc
                            Row_.Brand_Code = rowSC.Brand_Code
                            Row_.Brand_Desc = rowSC.Brand_Desc
                            Row_.Group_Code = rowSC.Group_Code
                            Row_.Group_Desc = rowSC.Group_Desc
                            Row_.Storage_Area_Code = rowSC.Storage_Area_Code
                            Row_.Storage_Area_Desc = rowSC.Storage_Area_Desc
                            Row_.Warehouse_Code = rowSC.Warehouse_Code
                            Row_.Warehouse_Desc = rowSC.Warehouse_Desc
                        End If
                        If Row_.Qty <> 0 Then
                            TBL_.Rows.Add(Row_)
                        End If
                    End If
                    idx0 += 1
                End While
                tblSC.Rows.Clear()

                Me.LoadPartialLocationTransfer(tblLTF, Me.DtpStockasOff.Value.Date.AddDays(1))
                Prg1.Maximum = tblLTF.DefaultView.Count
                Prg1.Value = 0
                Prg1.Refresh()
                idx0 = 0
                idx1 = 0
                tblLTF.DefaultView.RowFilter = Filter01_II
                tblLTF.DefaultView.Sort = "[Primary Product], [Primary Storage Area II], [Lot No]"
                TBL_.DefaultView.RowFilter = ""
                TBL_.DefaultView.Sort = "[Primary Product], [Primary Storage Area], [Lot No]"
                While idx0 < tblLTF.DefaultView.Count
                    rowLTF = tblLTF.DefaultView(idx0).Row
                    Prg1.Value += 1
                    Label4.Text = "Location Transfer calculation... " & Prg1.Value.ToString & " of " & Prg1.Maximum.ToString & "."
                    Dim found As Boolean = False
                    While idx1 < TBL_.DefaultView.Count
                        Row_ = TBL_.DefaultView(idx1).Row
                        If Row_.Primary_Product = rowLTF.Primary_Product Then
                            If Row_.Primary_Storage_Area = rowLTF.Primary_Storage_Area_II Then
                                If Row_.Lot_No = rowLTF.Lot_No Then
                                    Row_.Qty = Row_.Qty - rowLTF.Qty
                                    If Row_.Qty = 0 Then
                                        TBL_.Rows.Remove(Row_)
                                    End If
                                    found = True
                                    Exit While
                                End If
                            ElseIf Row_.Primary_Storage_Area > rowLTF.Primary_Storage_Area_II Then
                                Exit While
                            End If
                        ElseIf Row_.Primary_Product > rowLTF.Primary_Product Then
                            Exit While
                        End If
                        idx1 += 1
                    End While
                    If Not found Then
                        Row_ = TBL_.NewDataTable1Row
                        Row_.Primary_Product = rowLTF.Primary_Product
                        Row_.Lot_No = rowLTF.Lot_No
                        Row_.Expiry_Date = rowLTF.Expiry_Date
                        Row_.Primary_Storage_Area = rowLTF.Primary_Storage_Area_II
                        Row_.Qty = rowLTF.Qty * -1
                        If Me.LoadForMonthlyClosing = False Then
                            Row_._Decimal = rowLTF._Decimal
                            Row_.Batch_No_Controlled = rowLTF.Batch_No_Controlled
                            Row_.Location_Controlled = rowLTF.Location_Controlled
                            Row_.Expiry_Date_Controlled = rowLTF.Expiry_Date_Controlled
                            Row_.Type = rowLTF.Type
                            Row_.Primary_Unit = rowLTF.Primary_Unit
                            Row_.Category_Primary = rowLTF.Category_Primary
                            Row_.Sub_Brand_Primary = rowLTF.Sub_Brand_Primary
                            Row_.Brand_Primary = rowLTF.Brand_Primary
                            Row_.Group_Primary = rowLTF.Group_Primary
                            Row_.Qty_Secondary = -1
                            Row_.Product_Code = rowLTF.Product_Code
                            Row_.Description = rowLTF.Description
                            Row_.Unit_Code = rowLTF.Unit_Code
                            Row_.Categorey_Code = rowLTF.Categorey_Code
                            Row_.Category_Desc = rowLTF.Category_Desc
                            Row_.Sub_Brand_Code = rowLTF.Sub_Brand_Code
                            Row_.Sub_Brand_Desc = rowLTF.Sub_Brand_Desc
                            Row_.Brand_Code = rowLTF.Brand_Code
                            Row_.Brand_Desc = rowLTF.Brand_Desc
                            Row_.Group_Code = rowLTF.Group_Code
                            Row_.Group_Desc = rowLTF.Group_Desc
                            Row_.Storage_Area_Code = rowLTF.Storage_Area_Code_II
                            Row_.Storage_Area_Desc = rowLTF.Storage_Area_Desc_II
                            Row_.Warehouse_Code = rowLTF.Warehouse_Code_II
                            Row_.Warehouse_Desc = rowLTF.Warehouse_Desc_II
                        End If
                        If Row_.Qty <> 0 Then
                            TBL_.Rows.Add(Row_)
                        End If
                    End If
                    idx0 += 1
                End While

                Prg1.Value = 0
                Prg1.Refresh()
                idx0 = 0
                idx1 = 0
                tblLTF.DefaultView.RowFilter = Filter01_II
                tblLTF.DefaultView.Sort = "[Primary Product], [Primary Storage Area], [Lot No]"
                TBL_.DefaultView.RowFilter = ""
                TBL_.DefaultView.Sort = "[Primary Product], [Primary Storage Area], [Lot No]"
                While idx0 < tblLTF.DefaultView.Count
                    rowLTF = tblLTF.DefaultView(idx0).Row
                    Prg1.Value += 1
                    Label4.Text = "Location Transfer calculation II ... " & Prg1.Value.ToString & " of " & Prg1.Maximum.ToString & "."
                    Dim found As Boolean = False
                    While idx1 < TBL_.DefaultView.Count
                        Row_ = TBL_.DefaultView(idx1).Row
                        If Row_.Primary_Product = rowLTF.Primary_Product Then
                            If Row_.Primary_Storage_Area = rowLTF.Primary_Storage_Area Then
                                If Row_.Lot_No = rowLTF.Lot_No Then
                                    Row_.Qty = Row_.Qty + rowLTF.Qty
                                    If Row_.Qty = 0 Then
                                        TBL_.Rows.Remove(Row_)
                                    End If
                                    found = True
                                    Exit While
                                End If
                            ElseIf Row_.Primary_Storage_Area > rowLTF.Primary_Storage_Area Then
                                Exit While
                            End If
                        ElseIf Row_.Primary_Product > rowLTF.Primary_Product Then
                            Exit While
                        End If
                        idx1 += 1
                    End While
                    If Not found Then
                        Row_ = TBL_.NewDataTable1Row
                        Row_.Primary_Product = rowLTF.Primary_Product
                        Row_.Lot_No = rowLTF.Lot_No
                        Row_.Expiry_Date = rowLTF.Expiry_Date
                        Row_.Primary_Storage_Area = rowLTF.Primary_Storage_Area
                        Row_.Qty = rowLTF.Qty
                        If Me.LoadForMonthlyClosing = False Then
                            Row_._Decimal = rowLTF._Decimal
                            Row_.Batch_No_Controlled = rowLTF.Batch_No_Controlled
                            Row_.Location_Controlled = rowLTF.Location_Controlled
                            Row_.Expiry_Date_Controlled = rowLTF.Expiry_Date_Controlled
                            Row_.Type = rowLTF.Type
                            Row_.Primary_Unit = rowLTF.Primary_Unit
                            Row_.Category_Primary = rowLTF.Category_Primary
                            Row_.Sub_Brand_Primary = rowLTF.Sub_Brand_Primary
                            Row_.Brand_Primary = rowLTF.Brand_Primary
                            Row_.Group_Primary = rowLTF.Group_Primary
                            Row_.Qty_Secondary = -1
                            Row_.Product_Code = rowLTF.Product_Code
                            Row_.Description = rowLTF.Description
                            Row_.Unit_Code = rowLTF.Unit_Code
                            Row_.Categorey_Code = rowLTF.Categorey_Code
                            Row_.Category_Desc = rowLTF.Category_Desc
                            Row_.Sub_Brand_Code = rowLTF.Sub_Brand_Code
                            Row_.Sub_Brand_Desc = rowLTF.Sub_Brand_Desc
                            Row_.Brand_Code = rowLTF.Brand_Code
                            Row_.Brand_Desc = rowLTF.Brand_Desc
                            Row_.Group_Code = rowLTF.Group_Code
                            Row_.Group_Desc = rowLTF.Group_Desc
                            Row_.Storage_Area_Code = rowLTF.Storage_Area_Code
                            Row_.Storage_Area_Desc = rowLTF.Storage_Area_Desc
                            Row_.Warehouse_Code = rowLTF.Warehouse_Code
                            Row_.Warehouse_Desc = rowLTF.Warehouse_Desc
                        End If
                        If Row_.Qty <> 0 Then
                            TBL_.Rows.Add(Row_)
                        End If
                    End If
                    idx0 += 1
                End While
                tblLTF.Rows.Clear()

                DTA_ST.Fill(TBL_ST, Me.DtpStockasOff.Value.Date.AddDays(1))
                TBL_ST.DefaultView.RowFilter = Filter04
                Prg1.Value = 0
                Prg1.Maximum = TBL_ST.DefaultView.Count
                Prg1.Refresh()
                For Each DV1 In TBL_ST.DefaultView
                    row_ST = DV1.Row
                    Prg1.Value += 1
                    Label4.Text = "Stock Taking calculation... " & Prg1.Value.ToString & " of " & Prg1.Maximum.ToString & "."
                    TBL_.DefaultView.RowFilter = "[Primary Product] = " & row_ST.Primary_Product.ToString & " AND [Lot No] = '" & Me.ObjCD.FormatSpecialChar_DataFilter(row_ST.Batch_No) & "' AND [Primary Storage Area] = " & row_ST.Primary_Location.ToString
                    If TBL_.DefaultView.Count > 0 Then
                        Row_ = TBL_.DefaultView(0).Row
                        Row_.Qty = Row_.Qty - row_ST.Qty
                        If Row_.Qty = 0 Then
                            TBL_.Rows.Remove(Row_)
                        End If
                    Else
                        Row_ = TBL_.NewDataTable1Row
                        Row_.Primary_Product = row_ST.Primary_Product
                        Row_.Lot_No = row_ST.Batch_No
                        Row_.Expiry_Date = row_ST.Expiry_Date
                        Row_.Primary_Storage_Area = row_ST.Primary_Location
                        Row_.Qty = row_ST.Qty * (-1)

                        If Me.LoadForMonthlyClosing = False Then
                            Row_.Product_Code = row_ST.Product_Code
                            Row_.Description = row_ST.Description
                            Row_._Decimal = row_ST._Decimal
                            Row_.Batch_No_Controlled = row_ST.Batch_No_Controlled
                            Row_.Location_Controlled = row_ST.Location_Controlled
                            Row_.Expiry_Date_Controlled = row_ST.Expiry_Date_Controlled
                            Row_.Type = row_ST.Type
                            Row_.Primary_Unit = row_ST.Primary_Unit
                            Row_.Unit_Code = row_ST.Unit_Code
                            Row_.Category_Primary = row_ST.Category_Primary
                            Row_.Categorey_Code = row_ST.Category_Code
                            Row_.Category_Desc = row_ST.Category_Desc
                            Row_.Sub_Brand_Primary = row_ST.Sub_Brand_Primary
                            Row_.Sub_Brand_Code = row_ST.Sub_Brand_Code
                            Row_.Sub_Brand_Desc = row_ST.Sub_Brand_Desc
                            Row_.Brand_Primary = row_ST.Brand_Primary
                            Row_.Brand_Code = row_ST.Brand_Code
                            Row_.Brand_Desc = row_ST.Brand_Desc
                            Row_.Group_Primary = row_ST.Group_Primary
                            Row_.Group_Code = row_ST.Product_Group_Code
                            Row_.Group_Desc = row_ST.Product_Group_Desc
                            Row_.Storage_Area_Code = row_ST.Storage_Area_Code
                            Row_.Storage_Area_Desc = row_ST.Storage_Area_Desc
                            Row_.Qty_Secondary = -1
                            Row_.Warehouse_Code = row_ST.Warehouse_Code
                            Row_.Warehouse_Desc = row_ST.Warehouse_Desc
                        End If
                        If Row_.Qty <> 0 Then
                            TBL_.Rows.Add(Row_)
                        End If
                    End If
                Next
                TBL_ST.Rows.Clear()
                Label4.Text = "Validating...."

                Dim listOfPrimaryLocation As New List(Of Double)
                For Each dgvRow As DataGridViewRow In Me.DgvLocation.Rows
                    If IsNothing(dgvRow.Cells(Me.ColLocation_Primary.Index).Value) = False Then
                        If listOfPrimaryLocation.Contains(dgvRow.Cells(Me.ColLocation_Primary.Index).Value) = False Then
                            listOfPrimaryLocation.Add(dgvRow.Cells(Me.ColLocation_Primary.Index).Value)
                        End If
                    End If
                Next

                ''cek minus
                'For Each Row_ In TBL_
                '    If Row_.Qty < 0 Then
                '        If ListPrimaryMinus.Contains(Row_.Primary_Product) = False Then
                '            ListPrimaryMinus.Add(Row_.Primary_Product)
                '        End If
                '    End If
                'Next
                'For Each PrimaryProduct As Double In ListPrimaryMinus
                '    TBL_.DefaultView.RowFilter = "[Primary Product] = " + PrimaryProduct.ToString
                '    TBL_.DefaultView.Sort = "Qty ASC"

                '    IdxAcuan = 0
                '    Row_ = TBL_.DefaultView(IdxAcuan).Row
                '    While TBL_.DefaultView.Count > 1 And IdxAcuan <= TBL_.DefaultView.Count - 1
                '        Row_ = TBL_.DefaultView(IdxAcuan).Row
                '        Row_1 = TBL_.DefaultView(TBL_.DefaultView.Count - 1).Row
                '        If Row_1.Qty < 0 Then
                '            Exit While
                '        ElseIf Row_.Qty > 0 Then
                '            Exit While
                '        Else
                '            If Row_.Qty + Row_1.Qty = 0 Then
                '                TBL_.Rows.Remove(Row_)
                '                TBL_.Rows.Remove(Row_1)
                '            ElseIf Row_.Qty + Row_1.Qty > 0 Then
                '                Row_1.Qty += Row_.Qty
                '                TBL_.Rows.Remove(Row_)
                '            ElseIf Row_.Qty + Row_1.Qty < 0 Then
                '                Row_.Qty += Row_1.Qty
                '                TBL_.Rows.Remove(Row_1)
                '            End If
                '        End If
                '    End While
                'Next

                If Me.LoadForMonthlyClosing = True Then
                    For Each Row_ In TBL_
                        ROWCetak = TBLCetak.NewTblCetakRow
                        ROWCetak.Primary_Product = Row_.Primary_Product
                        ROWCetak.Lot_No = Row_.Lot_No
                        ROWCetak.Expiry_Date_Date = Row_.Expiry_Date
                        ROWCetak.Primary_Location = Row_.Primary_Storage_Area
                        ROWCetak.QtyDec = Row_.Qty
                        TBLCetak.Rows.Add(ROWCetak)
                    Next
                    Exit Function
                End If

                If Me.UserCode = "USER" And Me.ExportToBeginingBalance = True Then
                    If MsgBox("Insert as begining balance?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.InsertBeginingBalance(TBL_)
                    End If
                End If

                TBL.Rows.Clear()
                For Each Row_ In TBL_
                    If listOfPrimaryLocation.Count > 0 Then
                        If listOfPrimaryLocation.Contains(Row_.Primary_Storage_Area) = True Then
                            Row = TBL.NewDataTable1Row
                            For i As Integer = 0 To TBL.Columns.Count - 1
                                Row.Item(i) = Row_.Item(i)
                            Next
                            TBL.Rows.Add(Row)
                        End If
                    Else
                        Row = TBL.NewDataTable1Row
                        For i As Integer = 0 To TBL.Columns.Count - 1
                            Row.Item(i) = Row_.Item(i)
                        Next
                        TBL.Rows.Add(Row)
                    End If
                Next
                TBL_.Rows.Clear()
                TBL.DefaultView.RowFilter = ""

            End If
            ''================================================================================================================================================

            If TBL.DefaultView.Count > 0 Then
                For Each DV1 In TBL.DefaultView
                    Row = DV1.Row
                    ROWCetak = TBLCetak.NewTblCetakRow
                    With Row
                        ROWCetak._Decimal = ._Decimal
                        ROWCetak.QtyDec = .Qty
                        ROWCetak.Units_Code = .Unit_Code
                        ROWCetak.Primary = .Primary
                        ROWCetak.Primary_Product = .Primary_Product
                        ROWCetak.Primary_Unit = .Primary_Unit
                        ROWCetak.Product_Code = .Product_Code
                        ROWCetak.Description = .Description
                        ROWCetak.Brand_Code = .Brand_Code
                        ROWCetak.SubBrand_Code = .Sub_Brand_Code
                        ROWCetak.Category_Code = .Categorey_Code
                        ROWCetak.Group_Code = .Group_Code

                        If .Batch_No_Controlled = 1 Then
                            'kondisi dimana batch double dalam satu product 
                            ROWCetak.Multi_Batch = "Yes"
                        Else
                            ROWCetak.Multi_Batch = ""
                        End If
                        If .Location_Controlled = FrmLaporanStockDetail2.EnumPilih.No Then
                            ROWCetak.Storage_Area_Code = "-"
                            ROWCetak.Primary_Location = -1
                        Else
                            ROWCetak.Storage_Area_Code = .Storage_Area_Code
                            ROWCetak.Primary_Location = .Primary_Storage_Area
                        End If
                        If .Batch_No_Controlled = FrmLaporanStockDetail2.EnumPilih.No Then
                            ROWCetak.Lot_No = "-"
                        Else
                            If Row.Lot_No = "" Then
                                ROWCetak.Lot_No = "-"
                            Else
                                ROWCetak.Lot_No = .Lot_No
                            End If
                        End If
                        If .Expiry_Date_Controlled = FrmLaporanStockDetail2.EnumPilih.No Then
                            ROWCetak.Expiry_Date = "-"
                        Else
                            If Row.Expiry_Date = CDate(#1/1/1950#) Then
                                ROWCetak.Expiry_Date = "-"
                            Else
                                ROWCetak.Expiry_Date = Format(.Expiry_Date, "dd MMM yyyy")
                            End If
                        End If
                        ROWCetak.Expiry_Date_Date = .Expiry_Date
                        ROWCetak.Qty = Me.ObjCD.DigitGrouping(.Qty, ._Decimal) & " " & .Unit_Code

                        '6. set grouping
                        If GroupBy_Index = EnumGroupBy.No_Grouping Then
                            ROWCetak.Group_By = ""
                        ElseIf GroupBy_Index = EnumGroupBy.By_Location Then
                            If Me.chkUseWarehouse.Checked = True Then
                                If ROWCetak.Storage_Area_Code <> "-" Or ROWCetak.Storage_Area_Code <> "" Then
                                    ROWCetak.Group_By = .Warehouse_Code & "-" & .Warehouse_Desc
                                    ParamPerGrouping = "Location : "
                                Else
                                    ParamPerGrouping = "Location : "
                                    ROWCetak.Storage_Area_Code = ""
                                End If
                            Else
                                If ROWCetak.Storage_Area_Code <> "-" Or ROWCetak.Storage_Area_Code <> "" Then
                                    ROWCetak.Group_By = .Storage_Area_Code & "-" & .Storage_Area_Desc
                                    ParamPerGrouping = "Location : "
                                Else
                                    ParamPerGrouping = "Location : "
                                    ROWCetak.Storage_Area_Code = ""
                                End If
                            End If
                        ElseIf GroupBy_Index = EnumGroupBy.By_Product_Group Then
                            If .Group_Code <> "" Then
                                ROWCetak.Group_By = .Group_Code & "-" & .Group_Desc
                                ParamPerGrouping = "Product Group : "
                            Else
                                ParamPerGrouping = "Product Group : "
                                ROWCetak.Group_By = ""
                            End If
                        ElseIf GroupBy_Index = EnumGroupBy.By_Brand Then
                            If .Brand_Code <> "" Then
                                ROWCetak.Group_By = .Brand_Code & "-" & .Brand_Desc
                                ParamPerGrouping = "Brand : "
                            Else
                                ParamPerGrouping = "Brand : "
                                ROWCetak.Group_By = ""
                            End If
                        ElseIf GroupBy_Index = EnumGroupBy.By_Sub_Brand Then
                            If .Sub_Brand_Code <> "" Then
                                ROWCetak.Group_By = .Sub_Brand_Code & "-" & .Sub_Brand_Desc
                                ParamPerGrouping = "Sub Brand : "
                            Else
                                ParamPerGrouping = "Sub Brand : "
                                ROWCetak.Group_By = ""
                            End If
                        ElseIf GroupBy_Index = EnumGroupBy.Group_By_Category Then
                            If .Categorey_Code <> "" Then
                                ROWCetak.Group_By = .Categorey_Code & "-" & .Category_Desc
                                ParamPerGrouping = "Category : "
                            Else
                                ParamPerGrouping = "Category : "
                                ROWCetak.Group_By = ""
                            End If
                        End If
                    End With
                    TBLCetak.AddTblCetakRow(ROWCetak)
                Next
            End If

            'ubah ke warehouse klo dibutuhkan
            If Me.chkUseWarehouse.Checked = True Then
                Me.GantiKeWarehouse()

                'tblDummy.Rows.Clear()
                'DTAKonv.Fill(TBLKonv)
                ''pindahkan ke tabel dummy sambil pindahkan menggunakan warehouse code
                'For Each ROWCetak In TBLCetak.Rows
                '    TBLKonv.DefaultView.RowFilter = "[Storage Area Code] = '" + ROWCetak.Storage_Area_Code + "'"
                '    If TBLKonv.DefaultView.Count = 0 Then
                '        'tidak kontrol lokasi
                '        tblDummy.DefaultView.RowFilter = "Storage_Area_Code = 'n/a' AND [Primary_Product] = " + ROWCetak.Primary_Product.ToString + _
                '        " AND Lot_No = '" + ROWCetak.Lot_No + "' AND Expiry_Date = '" + ROWCetak.Expiry_Date + "' AND [Group By] = '" + ROWCetak.Group_By + "'"
                '        If tblDummy.DefaultView.Count = 0 Then
                '            'tambahkan baris baru
                '            With ROWCetak
                '                rowDummy = tblDummy.NewTblCetakRow

                '                rowDummy._Decimal = ._Decimal
                '                rowDummy.Units_Code = .Units_Code
                '                rowDummy.Primary = tblDummy.Rows.Count + 1
                '                rowDummy.Primary_Product = .Primary_Product
                '                rowDummy.Primary_Unit = .Primary_Unit
                '                rowDummy.Product_Code = .Product_Code
                '                rowDummy.Description = .Description
                '                rowDummy.Brand_Code = .Brand_Code
                '                rowDummy.SubBrand_Code = .SubBrand_Code
                '                rowDummy.Category_Code = .Category_Code
                '                rowDummy.Group_Code = .Group_Code
                '                rowDummy.Multi_Batch = ROWCetak.Multi_Batch
                '                rowDummy.Storage_Area_Code = "n/a"
                '                rowDummy.Lot_No = .Lot_No
                '                rowDummy.Expiry_Date = .Expiry_Date
                '                rowDummy.Group_By = .Group_By
                '                rowDummy.Group_By = .Group_By
                '                rowDummy.Sub_Total_per_Group = .Sub_Total_per_Group
                '                rowDummy.Grand_Total = .Grand_Total
                '                rowDummy.Qty = .Qty
                '                rowDummy.QtyDec = .QtyDec
                '                rowDummy.Expiry_Date_Date = .Expiry_Date_Date

                '                tblDummy.Rows.Add(rowDummy)
                '            End With
                '        Else
                '            'tambahkan qty saja
                '            rowDummy = tblDummy.DefaultView(0).Row
                '            rowDummy.QtyDec += ROWCetak.QtyDec
                '            rowDummy.Qty = Me.ObjCD.DigitGrouping(rowDummy.QtyDec, rowDummy._Decimal) & " " & rowDummy.Units_Code

                '        End If
                '    Else
                '        rowKonv = TBLKonv.DefaultView(0).Row
                '        tblDummy.DefaultView.RowFilter = "Storage_Area_Code = '" + rowKonv.Warehouse_Code + "' AND [Primary_Product] = " + ROWCetak.Primary_Product.ToString + _
                '        " AND Lot_No = '" + ROWCetak.Lot_No + "' AND Expiry_Date = '" + ROWCetak.Expiry_Date + "' AND [Group By] = '" + ROWCetak.Group_By + "'"

                '        If tblDummy.DefaultView.Count = 0 Then
                '            'tambahkan baris baru
                '            With ROWCetak
                '                rowDummy = tblDummy.NewTblCetakRow
                '                rowDummy._Decimal = ._Decimal
                '                rowDummy.Units_Code = .Units_Code
                '                rowDummy.Primary = tblDummy.Rows.Count + 1
                '                rowDummy.Primary_Product = .Primary_Product
                '                rowDummy.Primary_Unit = .Primary_Unit
                '                rowDummy.Product_Code = .Product_Code
                '                rowDummy.Description = .Description
                '                rowDummy.Brand_Code = .Brand_Code
                '                rowDummy.SubBrand_Code = .SubBrand_Code
                '                rowDummy.Category_Code = .Category_Code
                '                rowDummy.Group_Code = .Group_Code
                '                rowDummy.Multi_Batch = ROWCetak.Multi_Batch
                '                rowDummy.Storage_Area_Code = rowKonv.Warehouse_Code
                '                rowDummy.Lot_No = .Lot_No
                '                rowDummy.Expiry_Date = .Expiry_Date
                '                rowDummy.Group_By = .Group_By
                '                rowDummy.Group_By = .Group_By
                '                rowDummy.Sub_Total_per_Group = .Sub_Total_per_Group
                '                rowDummy.Grand_Total = .Grand_Total
                '                rowDummy.Qty = .Qty
                '                rowDummy.QtyDec = .QtyDec
                '                rowDummy.Expiry_Date_Date = .Expiry_Date_Date

                '                tblDummy.Rows.Add(rowDummy)
                '            End With
                '        Else
                '            'tambahkan qty saja
                '            rowDummy = tblDummy.DefaultView(0).Row
                '            rowDummy.QtyDec += ROWCetak.QtyDec
                '            rowDummy.Qty = Me.ObjCD.DigitGrouping(rowDummy.QtyDec, rowDummy._Decimal) & " " & rowDummy.Units_Code

                '        End If
                '    End If
                'Next

                ''pindahkan kembali ke table cetak
                'TBLCetak.Rows.Clear()
                'For Each rowDummy In tblDummy.Rows
                '    ROWCetak = TBLCetak.NewTblCetakRow
                '    For ij As Integer = 0 To TBLCetak.Columns.Count - 1
                '        ROWCetak.Item(ij) = rowDummy.Item(ij)
                '    Next
                '    TBLCetak.Rows.Add(ROWCetak)
                'Next
            End If

            'merge qty yang sama
            Me.TBLCetak.DefaultView.RowFilter = ""
            Me.TBLCetak.DefaultView.Sort = "[Primary_Product], Lot_No, Expiry_Date, Storage_Area_Code, [Group By]"
            If Me.TBLCetak.DefaultView.Count > 1 Then
                Idx = 1

                Me.Prg1.Value = 0
                Me.Prg1.Maximum = Me.TBLCetak.DefaultView.Count
                While Idx <= Me.TBLCetak.DefaultView.Count - 1
                    Me.Prg1.Value += 1
                    Me.Label4.Text = "Merging Qty..."

                    ROWCetak = Me.TBLCetak.DefaultView(Idx).Row
                    rowDummy = Me.TBLCetak.DefaultView(Idx - 1).Row

                    If rowDummy.Storage_Area_Code = ROWCetak.Storage_Area_Code And _
                    rowDummy.Primary_Product = ROWCetak.Primary_Product And _
                    rowDummy.Lot_No = ROWCetak.Lot_No And _
                    rowDummy.Group_By = ROWCetak.Group_By Then
                        rowDummy.QtyDec += ROWCetak.QtyDec
                        rowDummy.Qty = Me.ObjCD.DigitGrouping(rowDummy.QtyDec, rowDummy._Decimal) & " " & rowDummy.Units_Code
                        TBLCetak.Rows.Remove(ROWCetak)
                    Else
                        Idx += 1
                    End If
                End While
            End If

            'buang qty nol
            Idx = 0
            Me.Prg1.Value = 0
            Me.Prg1.Maximum = Me.TBLCetak.DefaultView.Count
            While Idx <= Me.TBLCetak.DefaultView.Count - 1
                Me.Prg1.Value += 1
                Me.Label4.Text = "Checking Qty..."

                ROWCetak = Me.TBLCetak.DefaultView(Idx).Row
                If ROWCetak.QtyDec = 0 Then
                    Me.TBLCetak.Rows.Remove(ROWCetak)
                Else
                    Idx += 1
                End If
            End While

            'Sub total PerGroup
            Dim TblSubTotal As New DseLaporanStockDetail2.DataTable2DataTable
            Dim RowSubtotal As DseLaporanStockDetail2.DataTable2Row
            For Each ROWCetak In TBLCetak
                With ROWCetak
                    TblSubTotal.DefaultView.RowFilter = "[group] = '" & .Group_By & "' and [units] = '" & .Units_Code & "'"
                    If TblSubTotal.DefaultView.Count > 0 Then
                        RowSubtotal = TblSubTotal.DefaultView(0).Row
                        RowSubtotal.qty += .QtyDec
                        If RowSubtotal.max_default_precision < ._Decimal Then
                            RowSubtotal.max_default_precision = ._Decimal
                        End If
                    Else
                        RowSubtotal = TblSubTotal.NewDataTable2Row
                        RowSubtotal.group = .Group_By
                        RowSubtotal.units = .Units_Code
                        RowSubtotal.qty = .QtyDec
                        RowSubtotal.max_default_precision = ._Decimal
                        TblSubTotal.Rows.Add(RowSubtotal)
                    End If
                End With
            Next

            For Each ROWCetak In TBLCetak
                With ROWCetak
                    TblSubTotal.DefaultView.RowFilter = "[group] = '" & .Group_By & "'"
                    If TblSubTotal.DefaultView.Count > 0 Then
                        For i As Integer = 0 To TblSubTotal.DefaultView.Count - 1
                            RowSubtotal = TblSubTotal.DefaultView(i).Row
                            If i < TblSubTotal.DefaultView.Count - 1 Then
                                .Sub_Total_per_Group += ObjCD.DigitGrouping(RowSubtotal.qty, RowSubtotal.max_default_precision) & " " & RowSubtotal.units & vbNewLine
                            Else
                                .Sub_Total_per_Group += ObjCD.DigitGrouping(RowSubtotal.qty, RowSubtotal.max_default_precision) & " " & RowSubtotal.units
                            End If
                        Next
                    End If
                End With
            Next

            '----
            'GrandTotal
            Dim TblSubTotal2 As New DseLaporanStockDetail2.DataTable3DataTable
            Dim RowSubtotal2 As DseLaporanStockDetail2.DataTable3Row
            For Each ROWCetak In TBLCetak
                With ROWCetak
                    TblSubTotal2.DefaultView.RowFilter = "[units] = '" & .Units_Code & "'"
                    If TblSubTotal2.DefaultView.Count > 0 Then
                        RowSubtotal2 = TblSubTotal2.DefaultView(0).Row
                        RowSubtotal2.qty += .QtyDec
                        If RowSubtotal2.max_default_precision < ._Decimal Then
                            RowSubtotal2.max_default_precision = ._Decimal
                        End If
                    Else
                        RowSubtotal2 = TblSubTotal2.NewDataTable3Row
                        RowSubtotal2.units = .Units_Code
                        RowSubtotal2.qty = .QtyDec
                        RowSubtotal2.max_default_precision = ._Decimal
                        TblSubTotal2.Rows.Add(RowSubtotal2)
                    End If
                End With
            Next

            TblSubTotal2.DefaultView.RowFilter = Nothing
            'TblSubTotal2.DefaultView.RowFilter = "[units] = '" & .Units_Code & "'"
            If TblSubTotal2.DefaultView.Count > 0 Then
                For i As Integer = 0 To TblSubTotal2.DefaultView.Count - 1
                    RowSubtotal2 = TblSubTotal2.DefaultView(i).Row
                    If i < TblSubTotal2.DefaultView.Count - 1 Then
                        ParamGrandTotal += ObjCD.DigitGrouping(RowSubtotal2.qty, RowSubtotal2.max_default_precision) & " " & RowSubtotal2.units & vbNewLine
                    Else
                        ParamGrandTotal += ObjCD.DigitGrouping(RowSubtotal2.qty, RowSubtotal2.max_default_precision) & " " & RowSubtotal2.units
                    End If
                Next
            End If

            GetSorting(tblDummy, TBLCetak)
            Dim ParamShowGrandTotal As Boolean
            Dim ParamShowSubtotal As Boolean
            If Me.ShowSubTotal = EnumShowSubTotal.Yes Then
                ParamShowSubtotal = True
            End If
            If Me.ShowGrandTotal = EnumShowGrandTotal.Yes Then
                ParamShowGrandTotal = True
            End If

            If Me.GroupBy_Index = EnumGroupBy.By_Location Then
                ReportDetail1.SetDataSource(DS)
                ReportDetail1.SetParameterValue(ReportDetail1.Parameter_StrPerGroup.ParameterFieldName, ParamPerGrouping)
                ReportDetail1.SetParameterValue(ReportDetail1.Parameter_Showsubtotal.ParameterFieldName, ParamShowSubtotal)
                ReportDetail1.SetParameterValue(ReportDetail1.Parameter_ShowGrandtotal.ParameterFieldName, ParamShowGrandTotal)
                ReportDetail1.SetParameterValue(ReportDetail1.Parameter_GrandTotal.ParameterFieldName, ParamGrandTotal)
            ElseIf Me.GroupBy_Index = EnumGroupBy.By_Product_Group Or Me.GroupBy_Index = EnumGroupBy.By_Brand Or Me.GroupBy_Index = EnumGroupBy.By_Sub_Brand Or Me.GroupBy_Index = EnumGroupBy.Group_By_Category Then
                ReportDetail2.SetDataSource(DS)
                ReportDetail2.SetParameterValue(ReportDetail2.Parameter_StrGrouping.ParameterFieldName, ParamPerGrouping)
                ReportDetail2.SetParameterValue(ReportDetail2.Parameter_ShowSubTotal.ParameterFieldName, ParamShowSubtotal)
                ReportDetail2.SetParameterValue(ReportDetail2.Parameter_ShowGrandTotal.ParameterFieldName, ParamShowGrandTotal)
                ReportDetail2.SetParameterValue(ReportDetail2.Parameter_GrandTotal.ParameterFieldName, ParamGrandTotal)
            ElseIf Me.GroupBy_Index = EnumGroupBy.No_Grouping Then
                ReportDetail3.SetDataSource(DS)
                ReportDetail3.SetParameterValue(ReportDetail3.Parameter_ShowGrandTotal.ParameterFieldName, ParamShowGrandTotal)
                ReportDetail3.SetParameterValue(ReportDetail3.Parameter_GrandTotal.ParameterFieldName, ParamGrandTotal)
            End If

            Label4.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
            TBL.Dispose()
            DTAKonv.Dispose()
            TBLKonv.Dispose()
            DisableControls(False)
        End Try
    End Function

    Private Sub GantiKeWarehouse()
        Dim ROWCetak, rowDummy As DseLaporanStockDetail2.TblCetakRow
        Dim DTAKonv As New DseLaporanStockDetail2TableAdapters.TblKonversiTableAdapter
        Dim TBLKonv As New DseLaporanStockDetail2.TblKonversiDataTable
        Dim rowKonv As DseLaporanStockDetail2.TblKonversiRow
        Dim ListLocation As New List(Of String)
        Dim Idx As New Double
        Dim DV As DataRowView

        Try
            DTAKonv.Connection.ConnectionString = Core_Module.GetConstring
            DTAKonv.Fill(TBLKonv)

            Me.TBLCetak.DefaultView.RowFilter = ""
            Me.TBLCetak.DefaultView.Sort = "Storage_Area_Code"
            Idx = 0
            While Idx <= Me.TBLCetak.DefaultView.Count - 1
                ROWCetak = Me.TBLCetak.DefaultView(Idx).Row

                If Idx = 0 Then
                    ListLocation.Add(ROWCetak.Storage_Area_Code)
                Else
                    rowDummy = Me.TBLCetak.DefaultView(Idx - 1).Row
                    If rowDummy.Storage_Area_Code <> ROWCetak.Storage_Area_Code Then
                        ListLocation.Add(ROWCetak.Storage_Area_Code)
                    End If
                End If

                Idx += 1
            End While

            Me.Prg1.Value = 0
            Me.Prg1.Maximum = ListLocation.Count
            For Each LocationCode As String In ListLocation
                Me.Prg1.Value += 1
                If Me.Prg1.Value < 50 Then
                    Label4.Text = "Updating warehouse data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString
                Else
                    If Math.IEEERemainder(Me.Prg1.Value, 50) = 0 Then
                        Label4.Text = "Updating warehouse data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString
                    ElseIf Me.Prg1.Value = Me.Prg1.Maximum Then
                        Label4.Text = "Updating warehouse data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString
                    End If
                End If

                TBLKonv.DefaultView.RowFilter = "[Storage Area Code] = '" + LocationCode + "'"
                Me.TBLCetak.DefaultView.RowFilter = "Storage_Area_Code = '" + LocationCode + "'"

                If TBLKonv.DefaultView.Count = 0 Then
                    For Each DV In TBLCetak.DefaultView
                        ROWCetak = DV.Row
                        ROWCetak.Storage_Area_Code = "n/a"
                    Next
                Else
                    rowKonv = TBLKonv.DefaultView(0).Row
                    For Each DV In TBLCetak.DefaultView
                        ROWCetak = DV.Row
                        ROWCetak.Storage_Area_Code = rowKonv.Warehouse_Code
                    Next
                End If
            Next

          
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAKonv.Dispose()
            TBLKonv.Dispose()
        End Try
    End Sub

    Private Sub InsertBeginingBalance(ByVal TBLCetak As DseLaporanStockDetail2.DataTable1DataTable)
        Dim DTAItem As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim rowCetak As DseLaporanStockDetail2.DataTable1Row

        Try
            DTAItem.Connection.ConnectionString = Me.TargetConstring

            DTAItem.DeleteQuery(Me.PrimaryStockInOut)
            For Each rowCetak In TBLCetak.Rows
                With rowCetak
                    DTAItem.Insert_Multi(Me.PrimaryStockInOut, .Primary_Product, .Qty, .Primary_Unit, -1, -1, .Lot_No, _
                    .Expiry_Date, .Primary_Storage_Area, -1, -1, "", "", .Qty, .Primary_Unit, 1, 1)
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetSorting(ByVal tblDummy As DseLaporanStockDetail2.TblCetakDataTable, _
    ByVal TBLCetak As DseLaporanStockDetail2.TblCetakDataTable) As Boolean
        Dim rowDummy As DseLaporanStockDetail2.TblCetakRow
        Dim Row As DseLaporanStockDetail2.TblCetakRow
        Dim Sort_ As String = ""
        Try
            tblDummy.Rows.Clear()

            If Me.GroupBy_Index = EnumGroupBy.No_Grouping Then
                Sort_ = "[Product_Code] ASC, [Expiry Date Date] ASC, [Lot_No] ASC, [Storage_Area_Code] ASC"
            ElseIf Me.GroupBy_Index = EnumGroupBy.By_Location Then
                Sort_ = "[Storage_Area_Code] ASC,[Product_Code] ASC, [Expiry Date Date] ASC"
            ElseIf Me.GroupBy_Index = EnumGroupBy.By_Product_Group Then
                Sort_ = "[Group_Code] ASC,[Product_Code] ASC, [Expiry Date Date] ASC, [Lot_No] ASC,[Storage_Area_Code] ASC"
            ElseIf Me.GroupBy_Index = EnumGroupBy.By_Brand Then
                Sort_ = "[Brand_Code] ASC,[Product_Code] ASC, [Expiry Date Date] ASC, [Lot_No] ASC,[Storage_Area_Code] ASC"
            ElseIf Me.GroupBy_Index = EnumGroupBy.By_Sub_Brand Then
                Sort_ = "[SubBrand_Code] ASC,[Product_Code] ASC, [Expiry Date Date] ASC, [Lot_No] ASC,[Storage_Area_Code] ASC"
            ElseIf Me.GroupBy_Index = EnumGroupBy.Group_By_Category Then
                Sort_ = "[Category_Code] ASC,[Product_Code] ASC, [Expiry Date Date] ASC, [Lot_No] ASC,[Storage_Area_Code] ASC"
            End If

            For Each Row In TBLCetak
                rowDummy = tblDummy.NewTblCetakRow
                For i As Integer = 0 To TBLCetak.Columns.Count - 1
                    rowDummy.Item(i) = Row.Item(i)
                Next
                tblDummy.Rows.Add(rowDummy)
            Next

            TBLCetak.Rows.Clear()
            tblDummy.DefaultView.Sort = Nothing
            tblDummy.DefaultView.Sort = Sort_
            tblDummy.DefaultView.RowFilter = ""
            For i As Integer = 0 To tblDummy.DefaultView.Count - 1
                rowDummy = tblDummy.DefaultView(i).Row
                Row = TBLCetak.NewTblCetakRow
                For i2 As Integer = 0 To tblDummy.Columns.Count - 1
                    Row.Item(i2) = rowDummy.Item(i2)
                Next
                TBLCetak.Rows.Add(Row)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Dim from As Date
    Private Sub btnFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.DtpStockasOff.Enabled = True Then
                If EndLicenseDate.Date <= Me.DtpStockasOff.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.DtpMaxExpiryDate.Enabled = True Then
                If EndLicenseDate.Date <= Me.DtpMaxExpiryDate.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.DtpMinExpiryDate.Enabled = True Then
                If EndLicenseDate.Date <= Me.DtpMinExpiryDate.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            from = Now
            If Me.ExpiryMin = EnumExpiryMinDate.Yes And Me.ExpiryMax = EnumExpiryMaxDate.Yes Then
                If Me.DtpMinExpiryDate.Value.Date > Me.DtpMaxExpiryDate.Value.Date Then
                    MsgBox("Starting min expiry date must be smaller than ending max expiry date.", MsgBoxStyle.Information, "Information")
                    Me.DtpMinExpiryDate.Focus()
                    Exit Sub
                End If
            End If
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            Me.Prg1.Visible = True
            Me.ControlBox = False
            BackgroundWorker4.RunWorkerAsync()
            'CetakanStockDetail()
            'Me.Prg1.Visible = False
            'Me.ControlBox = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Public Function FilterProductOr() As String
        Dim Retval As String = ""
        Try
            If LevelIndex = EnumLevelIndex.Index01 And LevelGroup = EnumLevel.ProductGroup Then
                Retval = " Or [Group Primary] = "
            ElseIf LevelIndex = EnumLevelIndex.Index01 And LevelGroup = EnumLevel.Brand Then
                Retval = " Or [Brand Primary] = "
            ElseIf LevelIndex = EnumLevelIndex.Index01 And LevelGroup = EnumLevel.SubBrand Then
                Retval = " Or [Sub Brand Primary] = "
            ElseIf LevelIndex = EnumLevelIndex.Index02 And LevelGroup = EnumLevel.Category Then
                Retval = " Or [Category Primary] = "
            ElseIf LevelIndex = EnumLevelIndex.Index03 Then
                Retval = " Or [Primary Product] = "
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function FilterProductAnd(Optional ByVal Include As Boolean = False) As String
        Dim Retval As String = ""
        Try
            If Include = False Then
                If LevelIndex = EnumLevelIndex.Index01 And LevelGroup = EnumLevel.ProductGroup Then
                    Retval = " And [Group Primary] = "
                ElseIf LevelIndex = EnumLevelIndex.Index01 And LevelGroup = EnumLevel.Brand Then
                    Retval = " And [Brand Primary] = "
                ElseIf LevelIndex = EnumLevelIndex.Index01 And LevelGroup = EnumLevel.SubBrand Then
                    Retval = " And [Sub Brand Primary] = "
                ElseIf LevelIndex = EnumLevelIndex.Index02 And LevelGroup = EnumLevel.Category Then
                    Retval = " And [Category Primary] = "
                ElseIf LevelIndex = EnumLevelIndex.Index03 Then
                    Retval = " And [Primary Product] = "
                End If
            Else
                If LevelIndex = EnumLevelIndex.Index01 And LevelGroup = EnumLevel.ProductGroup Then
                    Retval = " And ([Group Primary] = "
                ElseIf LevelIndex = EnumLevelIndex.Index01 And LevelGroup = EnumLevel.Brand Then
                    Retval = " And ([Brand Primary] = "
                ElseIf LevelIndex = EnumLevelIndex.Index01 And LevelGroup = EnumLevel.SubBrand Then
                    Retval = " And ([Sub Brand Primary] = "
                ElseIf LevelIndex = EnumLevelIndex.Index02 And LevelGroup = EnumLevel.Category Then
                    Retval = " And ([Category Primary] = "
                ElseIf LevelIndex = EnumLevelIndex.Index03 Then
                    Retval = " And ([Primary Product] = "
                End If
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub ChkStockAsOFF_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkStockAsOFF.CheckedChanged
        Try
            If Me.ChkStockAsOFF.Checked = False Then
                Me.ShowStockAsOff = EnumShowStockAsOff.No
            ElseIf Me.ChkStockAsOFF.Checked = True Then
                Me.ShowStockAsOff = EnumShowStockAsOff.Yes
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkStockAsOFF_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkStockAsOFF.Click
        Try
            If Me.ChkStockAsOFF.Checked = False Then
                Me.DtpStockasOff.Enabled = False
                Me.DtpStockasOff.Value = Now
            ElseIf Me.ChkStockAsOFF.Checked = True Then
                Me.DtpStockasOff.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetItemStart(ByVal TBL As DseLaporanStockDetail2.DataTable1DataTable, _
   ByVal vTanggalTransaksi As Date)

        Dim DTA As New dseLaporanStockTakingTableAdapters.tbltrans_stock_taking_itemstart1TableAdapter

        Dim DTAInOut As New DseLaporanStockDetail2TableAdapters.Tbl_Transaksi_StockInOutTableAdapter
        Dim TBLInOut As New DseLaporanStockDetail2.Tbl_Transaksi_StockInOutDataTable
        Dim RowInOut As DseLaporanStockDetail2.Tbl_Transaksi_StockInOutRow
        Dim TipeInOut As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust

        Dim DTABefore As New DseLaporanStockDetail2TableAdapters.Tbl_BeforeTableAdapter
        Dim TBLBefore As New DseLaporanStockDetail2.Tbl_BeforeDataTable
        Dim rowBefore As DseLaporanStockDetail2.Tbl_BeforeRow
        Dim DTAAfter As New DseLaporanStockDetail2TableAdapters.Tbl_AfterTableAdapter
        Dim TBLAfter As New DseLaporanStockDetail2.Tbl_AfterDataTable
        Dim rowAfter As DseLaporanStockDetail2.Tbl_AfterRow

        Dim DTATransfer As New DseLaporanStockDetail2TableAdapters.Tbl_StockTransferTableAdapter
        Dim TBLTransfer As New DseLaporanStockDetail2.Tbl_StockTransferDataTable
        Dim rowTransfer As DseLaporanStockDetail2.Tbl_StockTransferRow

        Dim StartDate As Date = vTanggalTransaksi.AddDays(1)
        Dim TipeProductGoods As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods
        Dim DescErr As String = ""

        Try
            'start date termasuk, jadi semua transaksi stock pada tanggal stock taking dianggap dilakukan
            'setelah proses stock taking
            'hanya memperhitungkan item2 goods
            'stock in out
            DTAInOut.Fill(TBLInOut, #1/1/2100#, StartDate)
            For Each RowInOut In TBLInOut.Rows
                TipeInOut = RowInOut.Tipe_Transaksi
                With RowInOut
                    If TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust Or _
                    TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess Or _
                    TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process Or _
                    TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Or _
                    TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Then
                        'kurangkan item
                        Me.KurangkanItem(TBL, .Primary_Product, .Qty, .Secondary_Qty, .Batch_No, _
                        .Expiry_Date, .Primary_Storage_Area, _
                        .Primary_Category, .Category_Code, .Category_Desc, _
                        .Primary_SubBrand, .SubBrand_Code, .SubBrand_Desc, _
                        .Primary_Brand, .Brand_Code, .Brand_Desc, _
                        .Primary_Product_Group, .Product_Group_Code, .Group_Desc, .Batch_No_Controlled, .Expiry_Date_Controlled, _
                        .Product_Code, .Description, .Storage_Area_Code, .Storage_Area_Desc, .Type, .Unit_Code, ._Decimal, .Primary_Units)
                    ElseIf TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust Or _
                        TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Process Or _
                        TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn Or _
                        TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales Then
                        'tambahkan item
                        Me.TambahkanItem(TBL, .Primary_Product, .Qty, .Secondary_Qty, .Batch_No, _
                        .Expiry_Date, .Primary_Storage_Area, _
                        .Primary_Category, .Category_Code, .Category_Desc, _
                        .Primary_SubBrand, .SubBrand_Code, .SubBrand_Desc, _
                        .Primary_Brand, .Brand_Code, .Brand_Desc, _
                        .Primary_Product_Group, .Product_Group_Code, .Group_Desc, .Batch_No_Controlled, .Expiry_Date_Controlled, _
                        .Product_Code, .Description, .Storage_Area_Code, .Storage_Area_Desc, .Type, .Unit_Code, ._Decimal, .Primary_Units)
                    End If
                End With
            Next

            DTABefore.Fill(TBLBefore, StartDate, #1/1/2100#)
            For Each rowBefore In TBLBefore.Rows
                With rowBefore
                    Me.TambahkanItem(TBL, .Primary_Product, .Qty, .Secondary_Qty, .Batch_No, _
                        .Expiry_Date, .Primary_Storage_Area, _
                        .Primary_Category, .Category_Code, .Desc_Category, _
                        .Primary_SubBrand, .Code_SubBrand, .Desc_SubBrand, _
                        .Primary_Brand, .Brand_Code, .Desc_Brand, _
                        .Primary_Group, .Product_Group_Code, .Desc_Product_Group, .Batch_No_Controlled, .Expiry_Date_Controlled, _
                        .Product_Code, .Description, .Storage_Area_Code, .Storage_Area_Desc, .Type, .Unit_Code, ._Decimal, .Primary_Units)
                End With
            Next

            DTAAfter.Fill(TBLAfter, StartDate, #1/1/2100#)
            For Each rowAfter In TBLAfter.Rows
                With rowAfter
                    Me.KurangkanItem(TBL, .Primary_Product, .Qty, .Secondary_Qty, .Batch_No, _
                        .Expiry_Date, .Primary_Storage_Area, _
                        .Primary_Category, .Category_Code, .Desc_Category, _
                        .Primary_SubBrand, .Code_SubBrand, .Desc_SubBrand, _
                        .Primary_Brand, .Brand_Code, .Desc_Brand, _
                        .Primary_Group, .Product_Group_Code, .Desc_Product_Group, .Batch_No_Controlled, .Expiry_Date_Controlled, _
                        .Product_Code, .Description, .Storage_Area_Code, .Storage_Area_Desc, .Type, .Unit_Code, ._Decimal, .Primary_Units)
                End With
            Next

            'stock transfer-pasti goods, jadi tidak dibatasi lagi.
            DTATransfer.Fill(TBLTransfer, StartDate, #1/1/2100#)
            For Each rowTransfer In TBLTransfer.Rows
                With rowTransfer
                    'kurangkan jumlah dari lokasi tujuan
                    Me.KurangkanItem(TBL, .Primary_Product, .Qty, .Secondary_Qty, .Batch_No, _
                        .Expiry_Date, .primary_destination_area, _
                        .Primary_Category, .Category_Code, .Category_Desc, _
                        .Primary_SubBrand, .SubBrand_Code, .SubBrand_Desc, _
                        .Primary_Brand, .Brand_Code, .Brand_Desc, _
                        .Primary_Product_Group, .Product_Group_Code, .Product_Group_Desc, .Batch_No_Controlled, .Expiry_Date_Controlled, _
                        .Product_Code, .Description, .Storage_Area_Code, .Storage_Area_Desc, .Type, .Unit_Code, ._Decimal, .Primary_Units)
                    'tambahkan jumlah dari lokasi asal
                    Me.TambahkanItem(TBL, .Primary_Product, .Qty, .Secondary_Qty, .Batch_No, _
                        .Expiry_Date, .Primary_Storage_Area, _
                        .Primary_Category, .Category_Code, .Category_Desc, _
                        .Primary_SubBrand, .SubBrand_Code, .SubBrand_Desc, _
                        .Primary_Brand, .Brand_Code, .Brand_Desc, _
                        .Primary_Product_Group, .Product_Group_Code, .Product_Group_Desc, .Batch_No_Controlled, .Expiry_Date_Controlled, _
                        .Product_Code, .Description, .Storage_Area_Code, .Storage_Area_Desc, .Type, .Unit_Code, ._Decimal, .Primary_Units)
                End With
            Next
        Catch ex As Exception
            MsgBox(DescErr)
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAInOut.Dispose()
            TBLInOut.Dispose()
            DTABefore.Dispose()
            TBLBefore.Dispose()
            DTAAfter.Dispose()
            TBLAfter.Dispose()
            DTATransfer.Dispose()
            TBLTransfer.Dispose()
        End Try
    End Sub

    Private Sub KurangkanItem(ByVal TBL As DseLaporanStockDetail2.DataTable1DataTable, _
                ByVal vPrimaryProduct As Double, ByVal vQty As Double, ByVal vQtySecondary As Double, _
                ByVal vLotNo As String, ByVal vExpiryDate As Date, ByVal vPrimaryStorageArea As Double, _
                ByVal vPrimaryCategory_ As Double, ByVal vCategory_Code_ As String, ByVal vCategory_Desc_ As String, _
                ByVal vPrimarySubBrand_ As Double, ByVal vSubBrand_Code_ As String, ByVal vSubBrand_Desc_ As String, _
                ByVal vPrimaryBrand_ As Double, ByVal vBrand_Code_ As String, ByVal vBrand_Desc_ As String, _
                ByVal vPrimaryProduct_Group_ As Double, ByVal vProduct_Group_Code_ As String, ByVal vProduct_Group_Desc_ As String, _
                ByVal vBatch_No_Controlled As Integer, ByVal vExpiry_Date_Controlled As Integer, _
                ByVal vProduct_Code As String, ByVal vDescription As String, ByVal vStorage_Area_Code As String, ByVal vStorage_Area_Desc As String, _
                ByVal vType As Integer, ByVal vUnit_Code As String, ByVal vDecimal As Decimal, ByVal vPrimary_Unit As Double)

        Dim filter As String = ""
        Dim rowInOut As DseLaporanStockDetail2.DataTable1Row

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + _
            " AND [Qty Secondary] = " + vQtySecondary.ToString + " AND [Primary Storage Area] = " + _
            vPrimaryStorageArea.ToString + " AND [Expiry Date] = #" + vExpiryDate.ToShortDateString + _
            "# AND [Lot No] = '" + vLotNo + "'"

            TBL.DefaultView.RowFilter = filter
            If TBL.DefaultView.Count = 0 Then
                'tambahkan baris baru
                rowInOut = TBL.Rows(TBL.Rows.Count - 1)
                filter = rowInOut.Primary
                rowInOut = TBL.NewDataTable1Row
                With rowInOut
                    .Primary_Unit = vPrimary_Unit
                    .Product_Code = vProduct_Code
                    .Description = vDescription
                    .Storage_Area_Code = vStorage_Area_Code
                    .Storage_Area_Desc = vStorage_Area_Desc
                    .Unit_Code = vUnit_Code
                    ._Decimal = vDecimal
                    .Type = vType
                    .Expiry_Date_Controlled = vBatch_No_Controlled
                    .Batch_No_Controlled = vBatch_No_Controlled
                    .Expiry_Date = vExpiryDate
                    If vPrimaryStorageArea <> -1 Then
                        .Location_Controlled = 1
                    Else
                        .Location_Controlled = 0
                    End If
                    .Lot_No = vLotNo
                    .Primary = CDbl(filter) + 1
                    .Primary_Product = vPrimaryProduct
                    .Primary_Storage_Area = vPrimaryStorageArea
                    .Qty = vQty * (-1)
                    .Qty_Secondary = vQtySecondary
                    .Category_Primary = vPrimaryCategory_
                    .Categorey_Code = vCategory_Code_
                    .Category_Desc = vCategory_Desc_
                    .Sub_Brand_Code = vSubBrand_Code_
                    .Sub_Brand_Desc = vSubBrand_Desc_
                    .Sub_Brand_Primary = vPrimarySubBrand_
                    .Brand_Code = vBrand_Code_
                    .Brand_Desc = vBrand_Desc_
                    .Brand_Primary = vPrimaryBrand_
                    .Group_Code = vProduct_Group_Code_
                    .Group_Desc = vProduct_Group_Desc_
                    .Group_Primary = vPrimaryProduct_Group_
                    .Warehouse_Code = ""
                    .Warehouse_Desc = ""
                End With
                TBL.Rows.Add(rowInOut)
            Else
                'kurangkan nilai qty
                rowInOut = TBL.DefaultView(0).Row

                Dim QtyDec As Decimal = Me.ObjCD.DigitGrouping(rowInOut.Qty, rowInOut._Decimal)
                Dim QtyDec2 As Decimal = Me.ObjCD.DigitGrouping(vQty, vDecimal)
                QtyDec -= QtyDec2
                rowInOut.Qty = QtyDec
                If rowInOut.Qty = 0 Then
                    TBL.Rows.Remove(rowInOut)
                End If

                'rowInOut.Qty -= vQty
                'If rowInOut.Qty = 0 Then
                '    TBL.Rows.Remove(rowInOut)
                'End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TambahkanItem(ByVal TBL As DseLaporanStockDetail2.DataTable1DataTable, _
                ByVal vPrimaryProduct As Double, ByVal vQty As Double, ByVal vQtySecondary As Double, _
                ByVal vLotNo As String, ByVal vExpiryDate As Date, ByVal vPrimaryStorageArea As Double, _
                ByVal vPrimaryCategory_ As Double, ByVal vCategory_Code_ As String, ByVal vCategory_Desc_ As String, _
                ByVal vPrimarySubBrand_ As Double, ByVal vSubBrand_Code_ As String, ByVal vSubBrand_Desc_ As String, _
                ByVal vPrimaryBrand_ As Double, ByVal vBrand_Code_ As String, ByVal vBrand_Desc_ As String, _
                ByVal vPrimaryProduct_Group_ As Double, ByVal vProduct_Group_Code_ As String, ByVal vProduct_Group_Desc_ As String, _
                ByVal vBatch_No_Controlled As Integer, ByVal vExpiry_Date_Controlled As Integer, _
                ByVal vProduct_Code As String, ByVal vDescription As String, ByVal vStorage_Area_Code As String, ByVal vStorage_Area_Desc As String, _
                ByVal vType As Integer, ByVal vUnit_Code As String, ByVal vDecimal As Decimal, ByVal vPrimary_Unit As Double)


        Dim filter As String = ""
        Dim rowInOut As DseLaporanStockDetail2.DataTable1Row

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + _
            " AND [Qty Secondary] = " + vQtySecondary.ToString + " AND [Primary Storage Area] = " + _
            vPrimaryStorageArea.ToString + " AND [Expiry Date] = #" + vExpiryDate.ToShortDateString + _
            "# AND [Lot No] = '" + vLotNo + "'"

            TBL.DefaultView.RowFilter = filter
            If TBL.DefaultView.Count = 0 Then
                'tambahkan baris baru
                If TBL.Rows.Count > 0 Then
                    rowInOut = TBL.Rows(TBL.Rows.Count - 1)
                    filter = rowInOut.Primary
                Else
                    filter = "0"
                End If

                rowInOut = TBL.NewDataTable1Row
                With rowInOut
                    .Primary_Unit = vPrimary_Unit
                    .Product_Code = vProduct_Code
                    .Description = vDescription
                    .Storage_Area_Code = vStorage_Area_Code
                    .Storage_Area_Desc = vStorage_Area_Desc
                    ._Decimal = vDecimal
                    .Unit_Code = vUnit_Code
                    .Type = vType
                    .Expiry_Date_Controlled = vExpiry_Date_Controlled
                    .Batch_No_Controlled = vBatch_No_Controlled

                    .Expiry_Date = vExpiryDate
                    If vPrimaryStorageArea <> -1 Then
                        .Location_Controlled = 1
                    Else
                        .Location_Controlled = 0
                    End If
                    .Lot_No = vLotNo

                    .Primary = CDbl(filter) + 1
                    .Primary_Product = vPrimaryProduct
                    .Primary_Storage_Area = vPrimaryStorageArea
                    .Qty = vQty
                    .Qty_Secondary = vQtySecondary

                    .Category_Primary = vPrimaryCategory_
                    .Categorey_Code = vCategory_Code_
                    .Category_Desc = vCategory_Desc_
                    .Sub_Brand_Code = vSubBrand_Code_
                    .Sub_Brand_Desc = vSubBrand_Desc_
                    .Sub_Brand_Primary = vPrimarySubBrand_
                    .Brand_Code = vBrand_Code_
                    .Brand_Desc = vBrand_Desc_
                    .Brand_Primary = vPrimaryBrand_
                    .Group_Code = vProduct_Group_Code_
                    .Group_Desc = vProduct_Group_Desc_
                    .Group_Primary = vPrimaryProduct_Group_
                    .Warehouse_Code = ""
                    .Warehouse_Desc = ""
                End With
                TBL.Rows.Add(rowInOut)
            Else
                rowInOut = TBL.DefaultView(0).Row
                Dim QtyDec As Decimal = Me.ObjCD.DigitGrouping(rowInOut.Qty, rowInOut._Decimal)
                Dim QtyDec2 As Decimal = Me.ObjCD.DigitGrouping(vQty, vDecimal)
                QtyDec += QtyDec2
                rowInOut.Qty = QtyDec
                If rowInOut.Qty = 0 Then
                    TBL.Rows.Remove(rowInOut)
                End If

                ''tambahkan nilai qty
                'rowInOut = TBL.DefaultView(0).Row
                'rowInOut.Qty += vQty
                'If rowInOut.Qty = 0 Then
                '    TBL.Rows.Remove(rowInOut)
                'End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker4_DoWork1(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker4.DoWork
        Try
            ObjCD.UpdateThreading()
            Control.CheckForIllegalCrossThreadCalls = False
            CetakanStockDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker4_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker4.RunWorkerCompleted
        Try
            Prg1.Visible = False
            If GroupBy_Index = EnumGroupBy.By_Location Then
                crv1.ReportSource = ReportDetail1
            ElseIf GroupBy_Index = EnumGroupBy.By_Product_Group Or GroupBy_Index = EnumGroupBy.By_Brand Or GroupBy_Index = EnumGroupBy.By_Sub_Brand Or GroupBy_Index = EnumGroupBy.Group_By_Category Then
                crv1.ReportSource = ReportDetail2
            ElseIf GroupBy_Index = EnumGroupBy.No_Grouping Then
                crv1.ReportSource = ReportDetail3
            End If
            crv1.Show()
            Windows.Forms.Cursor.Current = Cursors.Arrow

            Me.ControlBox = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FromCutOff_BeginingBalance(ByRef vTbl As DseLaporanStockDetail2.TblCetakDataTable, ByVal vDtpCutOffPeriod As Date) As Boolean
        Try
            Me.ChkStockAsOFF.Checked = True
            Me.DtpStockasOff.Value = vDtpCutOffPeriod
            Me.CetakanStockDetail()
            vTbl = Me.TBLCetak.Copy
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function
End Class