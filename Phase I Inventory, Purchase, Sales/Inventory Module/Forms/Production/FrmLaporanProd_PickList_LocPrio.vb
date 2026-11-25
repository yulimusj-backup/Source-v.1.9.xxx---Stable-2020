Option Explicit On
Option Strict Off

Public Class FrmLaporanProd_PickList_LocPrio
    Public UserCode As String = "" 'JGN LUPA DISI DIMENU UTAMA KARNA SAVE DAN HAPUS BY USERCODE
    Public BackAllColor As Color = Core_Module.GetSMESSBackColorAll
    Dim ObjL As New Core_Module.clsMasterStorageArea
    Dim ObjW As New Core_Module.clsMasterWarehouse
    Dim ObjCD As New Core_Module.ClsCekData
    Dim Obj As New ClsLaporanProd_PickList_LocPrio
    Dim Buffer As String = ""
    Dim index As Integer = -1
    Dim _Exit As Boolean = False

    Public Enum EnumTipe
        _PriorityLocation
        _ExcludeLocation
    End Enum

    Public Enum EnumLevel
        _Lvl_Location
        _Lvl_Warehouse
    End Enum

    Private Sub FrmLaporanProd_PickList_LocPrio_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If _Exit = False Then
                If MsgBox("Exit without save?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanProd_PickList_LocPrio_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.dgvLocation.AlternatingRowsDefaultCellStyle.BackColor = Me.BackAllColor
            Me.dgvExcludeLocation.AlternatingRowsDefaultCellStyle.BackColor = Me.BackAllColor
            Me.dgvLocation.Rows.Clear()
            Me.dgvExcludeLocation.Rows.Clear()
            Me.Obj.IsiDataDgv(Me.dgvLocation, Me.dgvExcludeLocation, Me.UserCode, Me)
            Me.btnUP.Enabled = False
            Me.btnDown.Enabled = False
            Me.BtnUp1.Enabled = False
            Me.BtnDown1.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvLocation_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvLocation.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColLocation_Code.Index Or e.ColumnIndex = Me.ColLocation_Group.Index Then
                If e.ColumnIndex = Me.ColLocation_Code.Index Then
                    If IsNothing(Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    End If
                End If
                Me.Buffer = Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvLocation_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvLocation.CellEndEdit
        Try
            If IsNothing(Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If String.IsNullOrEmpty(Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If IsNothing(Me.Buffer) = True Then
                        Me.index = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                        Exit Sub
                    End If
                    If Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColLocation_Group.Index Then
                        cellEndEditGroupLevel(sender, e)
                    End If
                    If e.ColumnIndex = Me.ColLocation_Code.Index Then
                        cellEndEditCode(sender, e)
                    End If
                Else
                    If IsNothing(Me.Buffer) = False Then
                        If String.IsNullOrEmpty(Buffer) = False Then
                            Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                    Me.index = e.RowIndex
                    Me.BackgroundWorker1.RunWorkerAsync()
                End If
            Else
                Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
                'If IsNothing(Me.Buffer) = False Then
                '    If String.IsNullOrEmpty(Buffer) = False Then
                '        Me.index = e.RowIndex
                '        Me.BackgroundWorker1.RunWorkerAsync()
                '        Exit Sub
                '    End If
                'Else
                '    Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                '    Exit Sub
                'End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cellEndEditGroupLevel(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim frmFindLocation As New Core_Module.FrmCariMasterStorageArea
        Dim frmFindWarehouse As New Core_Module.frmCariMasterWarehouse
        Try
            Select Case Me.dgvLocation.Item(Me.ColLocation_Group.Index, e.RowIndex).Value
                Case "Location"
                    frmFindLocation.txtCari.Text = Me.dgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value
                    frmFindLocation.ShowDialog()
                    If frmFindLocation.PrimaryStorageArea > 0 Then
                        Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = frmFindLocation.PrimaryStorageArea
                        Me.dgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = ObjL.GetStorageAreaCode(Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value)
                        Me.dgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = ObjL.GetDescription(Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value)
                        Me.dgvLocation.Item(Me.ColLocation_PrimaryW.Index, e.RowIndex).Value = Obj.GetPrimaryWarehouseByPirmaryLocation(Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Me.index = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                        Exit Sub
                    End If
                Case "Warehouse"
                    frmFindWarehouse.txtCari.Text = Me.dgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value
                    frmFindWarehouse.ShowDialog()
                    If frmFindWarehouse.PrimaryWarehouse > 0 Then
                        Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = frmFindWarehouse.PrimaryWarehouse
                        Me.dgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = ObjW.GetWarehouseCode(Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value)
                        Me.dgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = ObjW.GetDescription(Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value)
                        Me.dgvLocation.Item(Me.ColLocation_PrimaryW.Index, e.RowIndex).Value = Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value
                    Else
                        Me.index = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                        Exit Sub
                    End If
            End Select

            If Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value <> -1 Then
                If Me.getIndex(e.RowIndex, Me.dgvLocation.Item(Me.ColLocation_Group.Index, e.RowIndex).Value, Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value) <> -1 Then
                    MsgBox(Me.dgvLocation.Item(Me.ColLocation_Group.Index, e.RowIndex).Value & " " & Me.dgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                    Me.index = e.RowIndex
                    Me.BackgroundWorker1.RunWorkerAsync()
                    Exit Sub
                End If
            End If

            If Me.ApakahAdaDiExcludeLocation(Me.dgvLocation.Item(Me.ColLocation_Group.Index, e.RowIndex).Value, Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value, _
                   Me.dgvLocation.Item(Me.ColLocation_PrimaryW.Index, e.RowIndex).Value) = True Then
                MsgBox("Already exist in exclude location.", MsgBoxStyle.Information, "Information")
                Me.index = e.RowIndex
                Me.BackgroundWorker1.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function getIndex(ByVal vRowException As Integer, ByVal vGroupLevel As String, ByVal vPrimary As Double) As Integer
        Try
            For Each dr As DataGridViewRow In Me.dgvLocation.Rows
                With dr
                    If Me.dgvLocation.Rows.IndexOf(dr) <> vRowException Then
                        If dr.Cells(Me.ColLocation_Primary.Index).Value <> -1 Then
                            If dr.Cells(Me.ColLocation_Group.Index).Value = vGroupLevel And dr.Cells(Me.ColLocation_Primary.Index).Value = vPrimary Then
                                Return Me.dgvLocation.Rows.IndexOf(dr)
                            End If
                        End If
                    End If
                End With
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function getIndex2(ByVal vRowException As Integer, ByVal vGroupLevel As String, ByVal vPrimary As Double) As Integer
        Try
            For Each dr As DataGridViewRow In Me.dgvExcludeLocation.Rows
                With dr
                    If Me.dgvExcludeLocation.Rows.IndexOf(dr) <> vRowException Then
                        If dr.Cells(Me.ColE_Primary.Index).Value <> -1 Then
                            If dr.Cells(Me.ColE_Group.Index).Value = vGroupLevel And dr.Cells(Me.ColE_Primary.Index).Value = vPrimary Then
                                Return Me.dgvExcludeLocation.Rows.IndexOf(dr)
                            End If
                        End If
                    End If
                End With
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub cellEndEditCode(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim frmFindLocation As New Core_Module.FrmCariMasterStorageArea
        Dim frmFindWarehouse As New Core_Module.frmCariMasterWarehouse
        Try
            Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
            Select Case Me.dgvLocation.Item(Me.ColLocation_Group.Index, e.RowIndex).Value
                Case "Location Level"
LocLvl:
                    If Me.ObjL.ApakahSudahAda(Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = ObjL.GetPrimary(Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = ObjL.GetDescription(Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value)
                        Me.dgvLocation.Item(Me.ColLocation_PrimaryW.Index, e.RowIndex).Value = Obj.GetPrimaryWarehouseByPirmaryLocation(Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindLocation.txtCari.Text = Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindLocation.ShowDialog()
                        If frmFindLocation.PrimaryStorageArea > 0 Then
                            Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = frmFindLocation.PrimaryStorageArea
                            Me.dgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = ObjL.GetStorageAreaCode(Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value)
                            Me.dgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = ObjL.GetDescription(Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value)
                            Me.dgvLocation.Item(Me.ColLocation_PrimaryW.Index, e.RowIndex).Value = Obj.GetPrimaryWarehouseByPirmaryLocation(Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.index = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If
                    End If
                Case "Location"
                    GoTo LocLvl
                Case "Warehouse"
                    If Me.ObjW.ApakahSudahAda(Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = ObjW.GetPrimary(Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = ObjW.GetDescription(Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value)
                        Me.dgvLocation.Item(Me.ColLocation_PrimaryW.Index, e.RowIndex).Value = Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value
                    Else
                        frmFindWarehouse.txtCari.Text = Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindWarehouse.ShowDialog()
                        If frmFindWarehouse.PrimaryWarehouse > 0 Then
                            Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = frmFindWarehouse.PrimaryWarehouse
                            Me.dgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = ObjW.GetWarehouseCode(Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value)
                            Me.dgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = ObjW.GetDescription(Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value)
                            Me.dgvLocation.Item(Me.ColLocation_PrimaryW.Index, e.RowIndex).Value = Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.index = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If
                    End If
            End Select

            If Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value <> -1 Then
                If Me.getIndex(e.RowIndex, Me.dgvLocation.Item(Me.ColLocation_Group.Index, e.RowIndex).Value, Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value) <> -1 Then
                    MsgBox(Me.dgvLocation.Item(Me.ColLocation_Group.Index, e.RowIndex).Value & " " & Me.dgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                    Me.index = e.RowIndex
                    Me.BackgroundWorker1.RunWorkerAsync()
                    Exit Sub
                End If
            End If

            If Me.ApakahAdaDiExcludeLocation(Me.dgvLocation.Item(Me.ColLocation_Group.Index, e.RowIndex).Value, Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value, _
                    Me.dgvLocation.Item(Me.ColLocation_PrimaryW.Index, e.RowIndex).Value) = True Then
                MsgBox("Already exist in exclude location.", MsgBoxStyle.Information, "Information")
                Me.index = e.RowIndex
                Me.BackgroundWorker1.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.dgvLocation.Rows.RemoveAt(Me.index)
            Me.ObjCD.SetNumber(Me.dgvLocation)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvLocation_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvLocation.DataError

    End Sub

    Private Sub dgvLocation_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvLocation.RowsAdded
        Dim dr As DataGridViewRow
        Try
            If Me.dgvLocation.AllowUserToAddRows = True Then
                dr = Me.dgvLocation.Rows(Me.dgvLocation.Rows.Count - 1)
                dr.Cells(Me.ColLocation_No.Index).Value = Me.dgvLocation.Rows.Count
                dr.Cells(Me.ColLocation_Group.Index).Value = "Location Level"
                dr.Cells(Me.ColLocation_Desc.Index).Value = ""
                dr.Cells(Me.ColLocation_Urutan.Index).Value = Me.dgvLocation.Rows.Count
                dr.Cells(Me.ColLocation_Primary.Index).Value = -1
                dr.Cells(Me.ColLocation_PrimaryW.Index).Value = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvLocation_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvLocation.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.dgvLocation)
            Me.btnUP.Enabled = False
            Me.btnDown.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvLocation_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvLocation.SelectionChanged
        Try
            If Me.dgvLocation.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            Else
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvLocation_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvLocation.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvLocation)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cellEndEditGroupLevel1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim frmFindLocation As New Core_Module.FrmCariMasterStorageArea
        Dim frmFindWarehouse As New Core_Module.frmCariMasterWarehouse
        Try
            Select Case Me.dgvExcludeLocation.Item(Me.ColE_Group.Index, e.RowIndex).Value
                Case "Location"
                    frmFindLocation.txtCari.Text = Me.dgvExcludeLocation.Item(Me.ColE_Code.Index, e.RowIndex).Value
                    frmFindLocation.ShowDialog()
                    If frmFindLocation.PrimaryStorageArea > 0 Then
                        Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value = frmFindLocation.PrimaryStorageArea
                        Me.dgvExcludeLocation.Item(Me.ColE_Code.Index, e.RowIndex).Value = ObjL.GetStorageAreaCode(Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value)
                        Me.dgvExcludeLocation.Item(Me.ColE_Desc.Index, e.RowIndex).Value = ObjL.GetDescription(Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value)
                        Me.dgvExcludeLocation.Item(Me.ColE_PrimaryW.Index, e.RowIndex).Value = Obj.GetPrimaryWarehouseByPirmaryLocation(Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.dgvExcludeLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Me.index = e.RowIndex
                        Me.BackgroundWorker2.RunWorkerAsync()
                    End If
                Case "Warehouse"
                    frmFindWarehouse.txtCari.Text = Me.dgvExcludeLocation.Item(Me.ColE_Code.Index, e.RowIndex).Value
                    frmFindWarehouse.ShowDialog()
                    If frmFindWarehouse.PrimaryWarehouse > 0 Then
                        Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value = frmFindWarehouse.PrimaryWarehouse
                        Me.dgvExcludeLocation.Item(Me.ColE_Code.Index, e.RowIndex).Value = ObjW.GetWarehouseCode(Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value)
                        Me.dgvExcludeLocation.Item(Me.ColE_Desc.Index, e.RowIndex).Value = ObjW.GetDescription(Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value)
                        Me.dgvExcludeLocation.Item(Me.ColE_PrimaryW.Index, e.RowIndex).Value = Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value
                    Else
                        Me.index = e.RowIndex
                        Me.BackgroundWorker2.RunWorkerAsync()
                    End If
            End Select

            If Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value <> -1 Then
                If Me.getIndex2(e.RowIndex, Me.dgvExcludeLocation.Item(Me.ColE_Group.Index, e.RowIndex).Value, Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value) <> -1 Then
                    MsgBox(Me.dgvExcludeLocation.Item(Me.ColE_Group.Index, e.RowIndex).Value & " " & Me.dgvExcludeLocation.Item(Me.ColE_Code.Index, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                    Me.index = e.RowIndex
                    Me.BackgroundWorker2.RunWorkerAsync()
                    Exit Sub
                End If
            End If

            If Me.ApakahAdaDiLocationPriority(Me.dgvExcludeLocation.Item(Me.ColE_Group.Index, e.RowIndex).Value, Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value, _
                    Me.dgvExcludeLocation.Item(Me.ColE_PrimaryW.Index, e.RowIndex).Value) = True Then
                MsgBox("Already exist in location priority.", MsgBoxStyle.Information, "Information")
                Me.index = e.RowIndex
                Me.BackgroundWorker2.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cellEndEditCode1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim frmFindLocation As New Core_Module.FrmCariMasterStorageArea
        Dim frmFindWarehouse As New Core_Module.frmCariMasterWarehouse
        Try
            Me.dgvExcludeLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvExcludeLocation.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
            Select Case Me.dgvExcludeLocation.Item(Me.ColE_Group.Index, e.RowIndex).Value
                Case "Location Level"
LocLvl:
                    If Me.ObjL.ApakahSudahAda(Me.dgvExcludeLocation.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value = ObjL.GetPrimary(Me.dgvExcludeLocation.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvExcludeLocation.Item(Me.ColE_Desc.Index, e.RowIndex).Value = ObjL.GetDescription(Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value)
                        Me.dgvExcludeLocation.Item(Me.ColE_PrimaryW.Index, e.RowIndex).Value = Obj.GetPrimaryWarehouseByPirmaryLocation(Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindLocation.txtCari.Text = Me.dgvExcludeLocation.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindLocation.ShowDialog()
                        If frmFindLocation.PrimaryStorageArea > 0 Then
                            Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value = frmFindLocation.PrimaryStorageArea
                            Me.dgvExcludeLocation.Item(Me.ColE_Code.Index, e.RowIndex).Value = ObjL.GetStorageAreaCode(Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value)
                            Me.dgvExcludeLocation.Item(Me.ColE_Desc.Index, e.RowIndex).Value = ObjL.GetDescription(Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value)
                            Me.dgvExcludeLocation.Item(Me.ColE_PrimaryW.Index, e.RowIndex).Value = Obj.GetPrimaryWarehouseByPirmaryLocation(Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvExcludeLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.index = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        End If
                    End If
                Case "Location"
                    GoTo LocLvl
                Case "Warehouse"
                    If Me.ObjW.ApakahSudahAda(Me.dgvExcludeLocation.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value = ObjW.GetPrimary(Me.dgvExcludeLocation.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvExcludeLocation.Item(Me.ColE_Desc.Index, e.RowIndex).Value = ObjW.GetDescription(Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value)
                        Me.dgvExcludeLocation.Item(Me.ColE_PrimaryW.Index, e.RowIndex).Value = Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value
                    Else
                        frmFindWarehouse.txtCari.Text = Me.dgvExcludeLocation.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindWarehouse.ShowDialog()
                        If frmFindWarehouse.PrimaryWarehouse > 0 Then
                            Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value = frmFindWarehouse.PrimaryWarehouse
                            Me.dgvExcludeLocation.Item(Me.ColE_Code.Index, e.RowIndex).Value = ObjW.GetWarehouseCode(Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value)
                            Me.dgvExcludeLocation.Item(Me.ColE_Desc.Index, e.RowIndex).Value = ObjW.GetDescription(Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value)
                            Me.dgvExcludeLocation.Item(Me.ColE_PrimaryW.Index, e.RowIndex).Value = Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvExcludeLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.index = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        End If
                    End If
            End Select

            If Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value <> -1 Then
                If Me.getIndex2(e.RowIndex, Me.dgvExcludeLocation.Item(Me.ColE_Group.Index, e.RowIndex).Value, Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value) <> -1 Then
                    MsgBox(Me.dgvExcludeLocation.Item(Me.ColE_Group.Index, e.RowIndex).Value & " " & Me.dgvExcludeLocation.Item(Me.ColE_Code.Index, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                    Me.index = e.RowIndex
                    Me.BackgroundWorker2.RunWorkerAsync()
                    Exit Sub
                End If
            End If

            If Me.ApakahAdaDiLocationPriority(Me.dgvExcludeLocation.Item(Me.ColE_Group.Index, e.RowIndex).Value, Me.dgvExcludeLocation.Item(Me.ColE_Primary.Index, e.RowIndex).Value, _
                    Me.dgvExcludeLocation.Item(Me.ColE_PrimaryW.Index, e.RowIndex).Value) = True Then
                MsgBox("Already exist in location priority.", MsgBoxStyle.Information, "Information")
                Me.index = e.RowIndex
                Me.BackgroundWorker2.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvExcludeLocation_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvExcludeLocation.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColE_Code.Index Or e.ColumnIndex = Me.ColE_Group.Index Then
                If e.ColumnIndex = Me.ColE_Code.Index Then
                    If IsNothing(Me.dgvExcludeLocation.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvExcludeLocation.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    End If
                End If
                Me.Buffer = Me.dgvExcludeLocation.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvExcludeLocation_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvExcludeLocation.CellEndEdit
        Try
            If IsNothing(Me.dgvExcludeLocation.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If String.IsNullOrEmpty(Me.dgvExcludeLocation.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If IsNothing(Me.Buffer) = True Then
                        Me.index = e.RowIndex
                        Me.BackgroundWorker2.RunWorkerAsync()
                        Exit Sub
                    End If
                    If Me.dgvExcludeLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    
                    If e.ColumnIndex = Me.ColE_Group.Index Then
                        cellEndEditGroupLevel1(sender, e)
                    End If
                    If e.ColumnIndex = Me.ColE_Code.Index Then
                        cellEndEditCode1(sender, e)
                    End If
                Else
                    If IsNothing(Me.Buffer) = False Then
                        If String.IsNullOrEmpty(Buffer) = False Then
                            Me.dgvExcludeLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                    Me.index = e.RowIndex
                    Me.BackgroundWorker2.RunWorkerAsync()
                End If
            Else
                Me.dgvExcludeLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvExcludeLocation_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvExcludeLocation.DataError

    End Sub

    Private Sub dgvExcludeLocation_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvExcludeLocation.RowsAdded
        Dim dr As DataGridViewRow
        Try
            If Me.dgvExcludeLocation.AllowUserToAddRows = True Then
                dr = Me.dgvExcludeLocation.Rows(Me.dgvExcludeLocation.Rows.Count - 1)
                dr.Cells(Me.ColE_No.Index).Value = Me.dgvExcludeLocation.Rows.Count
                dr.Cells(Me.ColE_Group.Index).Value = "Location Level"
                dr.Cells(Me.ColE_Desc.Index).Value = ""
                dr.Cells(Me.ColE_Primary.Index).Value = -1
                dr.Cells(Me.ColE_PrimaryW.Index).Value = -1
                dr.Cells(Me.ColE_Urutan.Index).Value = Me.dgvExcludeLocation.Rows.Count
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvExcludeLocation_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvExcludeLocation.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.dgvExcludeLocation)
            Me.BtnUp1.Enabled = False
            Me.BtnDown1.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvExcludeLocation_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvExcludeLocation.SelectionChanged
        Try
            If Me.dgvExcludeLocation.SelectedRows.Count > 0 Then
                Me.BtnUp1.Enabled = True
                Me.BtnDown1.Enabled = True
            Else
                Me.BtnUp1.Enabled = False
                Me.BtnDown1.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvExcludeLocation_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvExcludeLocation.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvExcludeLocation)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(10)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            If Me.dgvExcludeLocation.RowCount > 1 Then
                Me.dgvExcludeLocation.Rows.RemoveAt(Me.index)
                Me.ObjCD.SetNumber(Me.dgvExcludeLocation)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahAdaDiExcludeLocation(ByVal vGroupLevelStr As String, ByVal vPrimary As Double, ByVal vPrimaryW As Double) As Boolean
        Try
            If GetDgvExcludeLocation(vGroupLevelStr, vPrimary, vPrimaryW) = 1 Then
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahAdaDiLocationPriority(ByVal vGroupLevelStr As String, ByVal vPrimary As Double, ByVal vPrimaryW As Double) As Boolean
        Try
            If GetDgvLocation(vGroupLevelStr, vPrimary, vPrimaryW) = 1 Then
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetDgvExcludeLocation(ByVal vGroupLevelStr As String, ByVal vPrimary As Double, ByVal vPrimaryW As Double) As Integer
        Try
            For Each dr As DataGridViewRow In Me.dgvExcludeLocation.Rows
                With dr
                    If vPrimary = -1 Then
                        Return 0
                    End If

                    If dr.Cells(Me.ColE_PrimaryW.Index).Value = vPrimaryW Then
                        Return 1
                    End If

                    If vGroupLevelStr = "Location" Or vGroupLevelStr = "Location level" Then
                        If dr.Cells(Me.ColE_Primary.Index).Value = vPrimary Then
                            Return 1
                        End If
                    End If
                End With
            Next
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetDgvLocation(ByVal vGroupLevelStr As String, ByVal vPrimary As Double, ByVal vPrimaryW As Double) As Integer
        Try
            For Each dr As DataGridViewRow In Me.dgvLocation.Rows
                With dr
                    If vPrimary = -1 Then
                        Return 0
                    End If

                    If dr.Cells(Me.ColLocation_PrimaryW.Index).Value = vPrimaryW Then
                        Return 1
                    End If

                    If vGroupLevelStr = "Location" Or vGroupLevelStr = "Location level" Then
                        If dr.Cells(Me.ColLocation_Primary.Index).Value = vPrimary Then
                            Return 1
                        End If
                    End If
                End With
            Next
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If MsgBox("Save detail item?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.Obj.HapusData(Me.UserCode)
                Me.Obj.SimpanData(Me.dgvLocation, Me.dgvExcludeLocation, Me, Me.UserCode)
                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                Me.DialogResult = Windows.Forms.DialogResult.OK
                _Exit = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowRowProcess(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < Me.dgvLocation.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    Me.dgvLocation.FirstDisplayedScrollingRowIndex = Me.dgvLocation.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.dgvLocation.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While Me.dgvLocation.Rows(row.Index).Displayed = False
                    Me.dgvLocation.FirstDisplayedScrollingRowIndex = Me.dgvLocation.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If Me.dgvLocation.DisplayedRowCount(True) <> Me.dgvLocation.DisplayedRowCount(False) Then
                    Me.dgvLocation.FirstDisplayedScrollingRowIndex = Me.dgvLocation.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUP.Click
        Try

            Me.dgvLocation.AllowUserToAddRows = False

            If Me.dgvLocation.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvLocation.SelectedRows(0)

                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.dgvLocation.Rows.Remove(row)
                    Me.dgvLocation.Rows.InsertRange(indexbaru, row)
                    Me.dgvLocation.Focus()
                    Me.dgvLocation.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvLocation)
                    If row.Displayed = False Then
                        Me.ShowRowProcess(row, True)
                    End If
                End If

                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
            Me.dgvLocation.AllowUserToAddRows = True
            Me.dgvLocation.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvLocation.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            Me.dgvLocation.AllowUserToAddRows = False
           
            If Me.dgvLocation.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False

                Dim row As DataGridViewRow = Me.dgvLocation.SelectedRows(0)
                If row.Index <> Me.dgvLocation.Rows.Count - 1 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.dgvLocation.Rows.Remove(row)
                    Me.dgvLocation.Rows.InsertRange(indexbaru, row)
                    Me.dgvLocation.Focus()
                    Me.dgvLocation.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvLocation)
                    Me.ShowRowProcess(row, False)
                End If

                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If

            Me.dgvLocation.AllowUserToAddRows = True
            Me.dgvLocation.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvLocation.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowRowProcess1(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < Me.dgvExcludeLocation.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    Me.dgvExcludeLocation.FirstDisplayedScrollingRowIndex = Me.dgvExcludeLocation.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.dgvExcludeLocation.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While Me.dgvExcludeLocation.Rows(row.Index).Displayed = False
                    Me.dgvExcludeLocation.FirstDisplayedScrollingRowIndex = Me.dgvExcludeLocation.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If Me.dgvExcludeLocation.DisplayedRowCount(True) <> Me.dgvExcludeLocation.DisplayedRowCount(False) Then
                    Me.dgvExcludeLocation.FirstDisplayedScrollingRowIndex = Me.dgvExcludeLocation.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnUp1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnUp1.Click
        Try

            Me.dgvExcludeLocation.AllowUserToAddRows = False

            If Me.dgvExcludeLocation.SelectedRows.Count > 0 Then
                Me.BtnUp1.Enabled = False
                Me.BtnDown1.Enabled = False
                Dim row As DataGridViewRow = Me.dgvExcludeLocation.SelectedRows(0)

                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.dgvExcludeLocation.Rows.Remove(row)
                    Me.dgvExcludeLocation.Rows.InsertRange(indexbaru, row)
                    Me.dgvExcludeLocation.Focus()
                    Me.dgvExcludeLocation.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvExcludeLocation)
                    If row.Displayed = False Then
                        Me.ShowRowProcess1(row, True)
                    End If
                End If

                Me.BtnUp1.Enabled = True
                Me.BtnDown1.Enabled = True
            End If
            Me.dgvExcludeLocation.AllowUserToAddRows = True
            Me.dgvExcludeLocation.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvExcludeLocation.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnDown1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDown1.Click
        Try
            Me.dgvExcludeLocation.AllowUserToAddRows = False

            If Me.dgvExcludeLocation.SelectedRows.Count > 0 Then
                Me.BtnUp1.Enabled = False
                Me.BtnDown1.Enabled = False

                Dim row As DataGridViewRow = Me.dgvExcludeLocation.SelectedRows(0)
                If row.Index <> Me.dgvExcludeLocation.Rows.Count - 1 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.dgvExcludeLocation.Rows.Remove(row)
                    Me.dgvExcludeLocation.Rows.InsertRange(indexbaru, row)
                    Me.dgvExcludeLocation.Focus()
                    Me.dgvExcludeLocation.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvExcludeLocation)
                    Me.ShowRowProcess1(row, False)
                End If

                Me.BtnUp1.Enabled = True
                Me.BtnDown1.Enabled = True
            End If

            Me.dgvExcludeLocation.AllowUserToAddRows = True
            Me.dgvExcludeLocation.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvExcludeLocation.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class