Option Explicit On
Option Strict Off
Public Class FrmMasterWorkUnit_AddSetupTime
    Public UserCode As String = "USER"
    Public PrimaryEmployee As Double = -1
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim baru As Boolean
    Dim buffer As String = ""
    Dim ObjCD As New Core_Module.ClsCekData
    Dim objSetup As New Core_Module.ClsMasterSetupTime
    Dim objProd As New Core_Module.clsMasterProduct
    Dim objWU As New Core_Module.ClsMasterWorkUnit

    Public TBL As New DseMasterWorkUnit_addSetupTime.Datatable_1DataTable
    Public row As DseMasterWorkUnit_addSetupTime.Datatable_1Row
    Dim EraseRowIndex As Integer = 0

    Dim ListOfIndex As New List(Of Integer)


    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmMasterWorkUnit_AddSetupTime_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim objItem As New Core_Module.ClsMasterWorkUnitItem
        Try

            Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNo, "Confirm")
            If hasil = MsgBoxResult.Yes Then
                If Me.ApakahDataBenar() = True Then
                    objItem.SimpanItemSetup2(Me.dgvSetupTime, Me)
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

    Private Sub FrmMasterWorkUnit_AddSetupTime_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            ' PuraPura()
            IsiKeDgv()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub PuraPura()
    '    Dim DTA As New DseMasterWorkUnit_addSetupTimeTableAdapters.tblmaster_workunit_item_setuptimeTableAdapter
    '    Dim tbl1 As New DseMasterWorkUnit_addSetupTime.tblmaster_workunit_item_setuptimeDataTable
    '    Dim row1 As DseMasterWorkUnit_addSetupTime.tblmaster_workunit_item_setuptimeRow
    '    Dim objProses As New Core_Module.ClsMasterProses

    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA.Fill(tbl1)

    '        If tbl1.Rows.Count > 0 Then
    '            For Each row1 In tbl1.Rows
    '                With row1
    '                    Me.row = Me.TBL.NewDatatable_1Row
    '                    Me.row.Primary = .Primary
    '                    Me.row.Primary_Product_1 = .Primary_Product_1
    '                    Me.row.Primary_Product_2 = .Primary_Product_2
    '                    Me.row.Primary_Process_1 = .Primary_Process_1
    '                    Me.row.Primary_Process_2 = .Primary_Process_2
    '                    Me.row.Primary_Setup_Time = .Primary_Setup_Time
    '                    Me.row.Setup_Time_Code = Me.objSetup.GetCode(.Primary_Setup_Time)
    '                    Me.row.Primary_Work_Unit = .Primary_Work_Unit
    '                    Me.row.Process_Code_1 = objProses.GetCode(.Primary_Process_1)
    '                    Me.row.Process_Code_2 = objProses.GetCode(.Primary_Process_2)
    '                    Me.row.Product_Code_1 = Me.objProd.GetProductCode(.Primary_Product_1)
    '                    Me.row.Product_Code_2 = Me.objProd.GetProductCode(.Primary_Product_2)
    '                    Me.row.Duration = .Duration
    '                    Me.TBL.Rows.Add(Me.row)
    '                End With
    '            Next
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub IsiKeDgv()
        Dim rowDgv As New DataGridViewRow
        Try
            dgvSetupTime.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvSetupTime.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Me.dgvSetupTime.Rows.Clear()
            If Me.TBL.Rows.Count > 0 Then
                For Each Me.row In Me.TBL.Rows
                    With Me.row
                        Me.dgvSetupTime.Rows.Add()
                        rowDgv = Me.dgvSetupTime.Rows(Me.dgvSetupTime.Rows.Count - 1)
                        rowDgv.Cells(Me.ColSetupTime_WorkUnit.Index).Value = .Primary_Setup_Time
                        rowDgv.Cells(Me.ColSetupTime_PrimarySetupTime.Index).Value = .Primary_Setup_Time
                        rowDgv.Cells(Me.ColSetupTime_Code.Index).Value = .Setup_Time_Code
                        If IsNumeric(.Duration) Then
                            rowDgv.Cells(Me.ColSetupTime_SetupTimeDuration.Index).Value = Me.ObjCD.Convert_ToString(.Duration)
                        Else
                            rowDgv.Cells(Me.ColSetupTime_SetupTimeDuration.Index).Value = Me.ObjCD.Convert_ToString(0)
                        End If
                        rowDgv.Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value = .Primary_Process_1
                        rowDgv.Cells(Me.ColSetupTime_PrimaryProses2.Index).Value = .Primary_Process_2
                        rowDgv.Cells(Me.ColSetupTime_ProcessCode1.Index).Value = .Process_Code_1
                        rowDgv.Cells(Me.ColSetupTime_ProcessCode2.Index).Value = .Process_Code_2
                        rowDgv.Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value = .Primary_Product_1
                        rowDgv.Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value = .Primary_Product_2
                        rowDgv.Cells(Me.ColSetupTime_CodeNo1.Index).Value = .Product_Code_1
                        rowDgv.Cells(Me.ColSetupTime_CodeNo2.Index).Value = .Product_Code_2
                        rowDgv.Cells(Me.ColSetupTime_Desc1.Index).Value = objProd.GetDescription(.Primary_Product_1)
                        rowDgv.Cells(Me.ColSetupTime_Desc2.Index).Value = objProd.GetDescription(.Primary_Product_2)
                        rowDgv.Cells(Me.ColSetupTime_PrimaryWorkUnit.Index).Value = .Primary_Work_Unit
                        rowDgv.Cells(Me.ColSetupTime_WorkUnit.Index).Value = Me.objWU.GetCode(.Primary_Work_Unit)
                        rowDgv.Cells(Me.ColSetupTime_Remarks.Index).Value = .Remarks
                        rowDgv.Cells(Me.ColSetupTime_Primary.Index).Value = .Primary
                    End With
                Next
            End If

            Me.ObjCD.SetNumber(Me.dgvSetupTime)

            dgvSetupTime.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgvSetupTime.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub dgvSetuptime_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvSetupTime.CellBeginEdit

        Try

            If e.ColumnIndex = Me.ColSetupTime_Code.Index Then
                If String.IsNullOrEmpty(Me.dgvSetupTime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = False Then
                    Me.buffer = Me.dgvSetupTime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                Else
                    Me.buffer = ""
                End If
            ElseIf e.ColumnIndex = Me.ColSetupTime_SetupTimeDuration.Index Then
                If Me.ObjCD.Is_Time_String(Me.dgvSetupTime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = True Then
                    Me.buffer = Me.dgvSetupTime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                Else
                    Me.buffer = Me.ObjCD.Convert_ToString(0)
                End If
            ElseIf e.ColumnIndex = Me.ColSetupTime_Set.Index Then
                If Me.dgvSetupTime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Boolean.TrueString Then
                    Me.buffer = Boolean.TrueString
                Else
                    Me.buffer = Boolean.FalseString
                End If
            Else
                e.Cancel = True
            End If

            'If e.ColumnIndex = Me.ColSetupTime_ProcessCode1.Index Then
            '    If String.IsNullOrEmpty(Me.dgvSetupTime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = False Then
            '        Me.buffer = Me.dgvSetupTime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
            '    Else
            '        Me.buffer = ""
            '    End If
            'Else
            '    If String.IsNullOrEmpty(Me.dgvSetupTime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode1.Index).Value) = False Then
            '        If e.ColumnIndex = Me.ColSetupTime_ProcessCode1.Index Then
            '            If String.IsNullOrEmpty(Me.dgvSetupTime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = False Then
            '                Me.buffer = Me.dgvSetupTime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
            '            Else
            '                Me.buffer = ""
            '            End If
            '        ElseIf e.ColumnIndex = Me.ColSetupTime_CodeNo1.Index Then
            '            If String.IsNullOrEmpty(Me.dgvSetupTime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = False Then
            '                Me.buffer = Me.dgvSetupTime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
            '            Else
            '                Me.buffer = ""
            '            End If
            '        ElseIf e.ColumnIndex = Me.ColSetupTime_ProcessCode2.Index Then
            '            If String.IsNullOrEmpty(Me.dgvSetupTime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = False Then
            '                Me.buffer = Me.dgvSetupTime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
            '            Else
            '                Me.buffer = ""
            '            End If
            '        ElseIf e.ColumnIndex = Me.ColSetupTime_CodeNo2.Index Then
            '            If String.IsNullOrEmpty(Me.dgvSetupTime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = False Then
            '                Me.buffer = Me.dgvSetupTime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
            '            Else
            '                Me.buffer = ""
            '            End If
            '        ElseIf e.ColumnIndex = Me.ColSetupTime_Code.Index Then
            '            If String.IsNullOrEmpty(Me.dgvSetupTime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = False Then
            '                Me.buffer = Me.dgvSetupTime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
            '            Else
            '                Me.buffer = ""
            '            End If
            '        ElseIf e.ColumnIndex = Me.ColSetupTime_SetupTimeDuration.Index Then
            '            If Me.ObjCD.Is_Time_String(Me.dgvSetupTime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = True Then
            '                Me.buffer = Me.dgvSetupTime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
            '            Else
            '                Me.buffer = ""
            '            End If
            '        ElseIf e.ColumnIndex = Me.ColSetupTime_Remarks.Index Then
            '            If String.IsNullOrEmpty(Me.dgvSetupTime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = False Then
            '                Me.buffer = Me.dgvSetupTime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
            '            Else
            '                Me.buffer = Me.ObjCD.Convert_ToString(0)
            '            End If
            '        ElseIf e.ColumnIndex = Me.ColSetupTime_WorkUnit.Index Then
            '            If String.IsNullOrEmpty(Me.dgvSetupTime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = False Then
            '                Me.buffer = Me.dgvSetupTime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
            '            Else
            '                Me.buffer = ""
            '            End If
            '        ElseIf e.ColumnIndex = Me.ColSetupTime_Set.Index Then
            '            If Me.dgvSetupTime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Boolean.TrueString Then
            '                Me.buffer = Boolean.TrueString
            '            Else
            '                Me.buffer = Boolean.FalseString
            '            End If
            '        Else
            '            e.Cancel = True
            '        End If
            '    Else
            '        e.Cancel = True
            '    End If
            'End If




        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try


    End Sub

    Private Sub dgvSetupTime_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSetupTime.CellClick

        Try


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSetuptime_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSetupTime.CellEndEdit
        Dim objPro As New Core_Module.ClsMasterProses
        Dim objCodeNo As New Core_Module.clsMasterProduct
        Dim IJ As New Core_Module.FrmMasterProsesCari
        Dim IK As New Core_Module.frmCariMasterProduct
        Dim IL As New Core_Module.FrmMasterSetupTime_Cari
        Try
            If e.ColumnIndex = Me.ColSetupTime_ProcessCode1.Index Then
                If Not String.IsNullOrEmpty(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode1.Index).Value) Then
                    If objPro.ApakahSudahAda(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode1.Index).Value) = True Then
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode1.Index).Value = _
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode1.Index).Value.ToString.ToUpper
                        If Me.ApakahDataMengulang(e.RowIndex, Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value, _
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value, _
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value, _
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value) = False Then
                            Dim primary As Double = -1
                            primary = objPro.GetPrimary(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode1.Index).Value)
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value = primary
                        Else
                            MsgBox("double input", MsgBoxStyle.Information, "Information")
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                If Me.dgvSetuptime.Rows.Count - 1 <> e.RowIndex Then
                                    Me.EraseRowIndex = e.RowIndex
                                    Me.BackgroundWorker3.RunWorkerAsync()
                                End If
                            Else
                                Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value = objPro.GetPrimary(Me.buffer)
                            End If
                        End If
                    Else
                        IJ.txtCari.Text = Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode1.Index).Value
                        IJ.ShowDialog()
                        If IJ.mPrimary > -1 Then
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value = IJ.mPrimary
                            If Me.ApakahDataMengulang(e.RowIndex, Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value, _
                               Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value, _
                               Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value, _
                               Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value) = False Then
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode1.Index).Value = objPro.GetCode(IJ.mPrimary).ToString.ToUpper
                            Else
                                MsgBox("double input", MsgBoxStyle.Information, "Information")
                                If String.IsNullOrEmpty(Me.buffer) = True Then
                                    If Me.dgvSetuptime.Rows.Count - 1 <> e.RowIndex Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker3.RunWorkerAsync()
                                    End If
                                Else
                                    Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value = objPro.GetPrimary(Me.buffer)
                                End If
                            End If
                        Else
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                If Me.dgvSetuptime.Rows.Count - 1 <> e.RowIndex Then
                                    Me.EraseRowIndex = e.RowIndex
                                    Me.BackgroundWorker3.RunWorkerAsync()
                                End If
                            Else
                                Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value = objPro.GetPrimary(Me.buffer)
                            End If
                        End If
                    End If
                Else
                    If String.IsNullOrEmpty(Me.buffer) = True Then
                        If Me.dgvSetuptime.Rows.Count - 1 <> e.RowIndex Then
                            Me.EraseRowIndex = e.RowIndex
                            Me.BackgroundWorker3.RunWorkerAsync()
                        End If
                    Else
                        Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value = objPro.GetPrimary(Me.buffer)
                    End If
                End If
            End If


            If e.ColumnIndex = Me.ColSetupTime_ProcessCode2.Index Then
                If Not String.IsNullOrEmpty(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode2.Index).Value) Then
                    If objPro.ApakahSudahAda(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode2.Index).Value) = True Then
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode2.Index).Value = Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode2.Index).Value.ToString.ToUpper
                        If Me.ApakahDataMengulang(e.RowIndex, Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value, _
                                                 Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value, _
                                                 Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value, _
                                                 Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value) = False Then

                            Dim primary As Double = -1
                            primary = objPro.GetPrimary(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode2.Index).Value)
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value = primary
                        Else
                            MsgBox("double input", MsgBoxStyle.Information, "Information")
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value = "-1"
                            Else
                                Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value = objPro.GetPrimary(Me.buffer)
                            End If
                        End If
                    Else
                        IJ.txtCari.Text = Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode2.Index).Value
                        IJ.ShowDialog()
                        If IJ.mPrimary > -1 Then
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value = IJ.mPrimary
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode2.Index).Value = objPro.GetCode(IJ.mPrimary).ToString.ToUpper
                        Else
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value = "-1"
                            Else
                                Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value = objPro.GetPrimary(Me.buffer)
                            End If
                        End If
                    End If
                Else
                    If String.IsNullOrEmpty(Me.buffer) = True Then
                        Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value = "-1"
                    Else
                        Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value = objPro.GetPrimary(Me.buffer)
                    End If
                End If
            End If






            If e.ColumnIndex = Me.ColSetupTime_CodeNo1.Index Then
                If Not String.IsNullOrEmpty(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value) Then
                    If objCodeNo.ApakahSudahAda(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value) = True Then
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value = Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value.ToString.ToUpper
                        If Me.ApakahDataMengulang(e.RowIndex, Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value, _
                           Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value, _
                           Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value, _
                           Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value) = False Then
                            Dim primary As Double = -1
                            primary = objCodeNo.GetPrimary(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value)
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value = primary
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc1.Index).Value = objCodeNo.GetDescription(primary)
                        Else
                            MsgBox("double input", MsgBoxStyle.Information, "Information")
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value = ""
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value = -1
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc1.Index).Value = ""
                            Else
                                Dim primary As Double = -1
                                primary = objCodeNo.GetPrimary(Me.buffer)
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value = Me.buffer
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value = primary
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc1.Index).Value = objCodeNo.GetDescription(primary)
                            End If
                        End If
                    Else
                        IK.txtCari.Text = Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value
                        IK.ShowDialog()
                        If IK.PrimaryProduct > -1 Then
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value = IK.PrimaryProduct
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value = objCodeNo.GetProductCode(IK.PrimaryProduct).ToString.ToUpper
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc1.Index).Value = objCodeNo.GetDescription(IK.PrimaryProduct)
                        Else
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value = ""
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value = -1
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc1.Index).Value = ""
                            Else
                                Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Dim PrimaryProd As Double = -1
                                PrimaryProd = objCodeNo.GetPrimary(Me.buffer)
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value = PrimaryProd
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value = Me.buffer
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc1.Index).Value = objCodeNo.GetDescription(PrimaryProd)
                            End If
                        End If
                    End If

                Else
                    If String.IsNullOrEmpty(Me.buffer) = True Then
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value = ""
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value = -1
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc1.Index).Value = ""
                    Else
                        Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Dim PrimaryProd As Double = -1
                        PrimaryProd = objCodeNo.GetPrimary(Me.buffer)
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value = PrimaryProd
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value = Me.buffer
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc1.Index).Value = objCodeNo.GetDescription(PrimaryProd)
                    End If
                End If
            End If







            If e.ColumnIndex = Me.ColSetupTime_CodeNo2.Index Then
                If Not String.IsNullOrEmpty(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value) Then
                    If objCodeNo.ApakahSudahAda(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value) = True Then
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value = Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value.ToString.ToUpper
                        Dim primary As Double = -1
                        primary = objCodeNo.GetPrimary(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value)
                        If Me.ApakahDataMengulang(e.RowIndex, Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value, _
                       Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value, _
                       Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value, _
                       primary) = False Then
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value = primary
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc2.Index).Value = objCodeNo.GetDescription(primary)
                        Else
                            MsgBox("double input", MsgBoxStyle.Information, "Information")
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value = ""
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value = -1
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc2.Index).Value = ""
                            Else
                                Dim primaryDet As Double = -1
                                primaryDet = objCodeNo.GetPrimary(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value)
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value = primaryDet
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc2.Index).Value = objCodeNo.GetDescription(primaryDet)
                            End If
                        End If
                    Else

                        IK.txtCari.Text = Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value
                        IK.ShowDialog()
                        If IK.PrimaryProduct > -1 Then
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value = IK.PrimaryProduct
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value = objCodeNo.GetProductCode(IK.PrimaryProduct).ToString.ToUpper
                            If Me.ApakahDataMengulang(e.RowIndex, Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value, _
                                               Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value, _
                                               Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value, _
                                               Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value) = False Then
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc2.Index).Value = objCodeNo.GetDescription(IK.PrimaryProduct)
                            Else
                                MsgBox("double input", MsgBoxStyle.Information, "Information")
                                If String.IsNullOrEmpty(Me.buffer) = True Then
                                    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value = ""
                                    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value = -1
                                    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc2.Index).Value = ""
                                Else
                                    Dim primaryDet As Double = -1
                                    primaryDet = objCodeNo.GetPrimary(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value)
                                    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value = primaryDet
                                    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc2.Index).Value = objCodeNo.GetDescription(primaryDet)
                                End If
                            End If
                        Else
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value = ""
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value = -1
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc2.Index).Value = ""
                            Else
                                Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Dim PrimaryProd As Double = -1
                                PrimaryProd = objCodeNo.GetPrimary(Me.buffer)
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value = PrimaryProd
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value = Me.buffer
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc2.Index).Value = objCodeNo.GetDescription(PrimaryProd)
                            End If
                        End If
                    End If
                Else
                    If String.IsNullOrEmpty(Me.buffer) = True Then
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value = ""
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value = -1
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc2.Index).Value = ""
                    Else
                        Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Dim PrimaryProd As Double = -1
                        PrimaryProd = objCodeNo.GetPrimary(Me.buffer)
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value = PrimaryProd
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value = Me.buffer
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc2.Index).Value = objCodeNo.GetDescription(PrimaryProd)
                    End If
                End If
            End If




            If e.ColumnIndex = Me.ColSetupTime_Code.Index Then
                If Not String.IsNullOrEmpty(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Code.Index).Value) Then
                    If Me.objSetup.ApakahSudahAda(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Code.Index).Value) = True Then
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Code.Index).Value = Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Code.Index).Value.ToString.ToUpper
                        Dim primary As Double = -1
                        primary = objSetup.GetPrimary(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Code.Index).Value)
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimarySetupTime.Index).Value = primary
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_SetupTimeDuration.Index).Value = Me.ObjCD.Convert_ToString(objSetup.GetDuration(primary))
                    Else
                        IL.txtCari.Text = Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Code.Index).Value
                        IL.ShowDialog()
                        If IK.PrimaryProduct > -1 Then
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimarySetupTime.Index).Value = IL.Primary
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Code.Index).Value = Me.objSetup.GetCode(IL.Primary).ToString.ToUpper
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_SetupTimeDuration.Index).Value = Me.ObjCD.Convert_ToString(objSetup.GetDuration(IL.Primary))
                        Else
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimarySetupTime.Index).Value = -1
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Code.Index).Value = ""
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_SetupTimeDuration.Index).Value = ""

                            Else
                                Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Dim PrimaryProd As Double = -1
                                PrimaryProd = objSetup.GetPrimary(Me.buffer)
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimarySetupTime.Index).Value = PrimaryProd
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Code.Index).Value = Me.buffer
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_SetupTimeDuration.Index).Value = Me.ObjCD.Convert_ToString(objSetup.GetDuration(PrimaryProd))
                            End If
                        End If
                    End If
                Else
                    If String.IsNullOrEmpty(Me.buffer) = True Then
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimarySetupTime.Index).Value = -1
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Code.Index).Value = ""
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_SetupTimeDuration.Index).Value = ""

                    Else
                        Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Dim PrimaryProd As Double = -1
                        PrimaryProd = objSetup.GetPrimary(Me.buffer)
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimarySetupTime.Index).Value = PrimaryProd
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Code.Index).Value = Me.buffer
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_SetupTimeDuration.Index).Value = Me.ObjCD.Convert_ToString(objSetup.GetDuration(PrimaryProd))
                    End If
                End If
            End If

            If e.ColumnIndex = Me.ColSetupTime_SetupTimeDuration.Index Then
                If Me.ObjCD.Is_Time_String(Me.dgvSetupTime.Rows(e.RowIndex).Cells(Me.ColSetupTime_SetupTimeDuration.Index).Value) = False Then
                    MsgBox("fill setuptime data only with time type data, please correct your data", MsgBoxStyle.Information, "Information")
                    If Me.ObjCD.Is_Time_String(Me.buffer) = True Then
                        Me.dgvSetupTime.Rows(e.RowIndex).Cells(Me.ColSetupTime_SetupTimeDuration.Index).Value = Me.buffer
                    Else
                        Me.dgvSetupTime.Rows(e.RowIndex).Cells(Me.ColSetupTime_SetupTimeDuration.Index).Value = Me.ObjCD.Convert_ToString(0)
                    End If
                End If
            End If


            'If Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value <> "-1" AndAlso _
            ' Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value <> "-1" AndAlso _
            ' Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimarySetupTime.Index).Value <> "-1" AndAlso _
            ' Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value <> "-1" AndAlso _
            ' Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value <> "-1" Then

            '    If Me.ApakahDataMengulang(e.RowIndex, Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value, _
            '    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value, _
            '    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value, _
            '    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value, _
            '    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimarySetupTime.Index).Value) = True Then
            '        MsgBox("Double input ", MsgBoxStyle.Information, "information")

            '    End If
            'End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objPro = Nothing
            objCodeNo = Nothing
            IJ.Dispose()
            IK.Dispose()
            IL.Dispose()
        End Try
    End Sub

    Public Function ApakahDataMengulang(ByVal vIndex As Integer, ByVal vPrimaryProses1 As Double, ByVal vPrimaryProses2 As Double, _
                                        ByVal vPrimaryProduct1 As Double, ByVal vprimaryProduct2 As Double) As Boolean
        Dim rowDGV As New DataGridViewRow

        Try
            If Me.dgvSetuptime.RowCount > 1 Then
                For intA As Integer = 0 To Me.dgvSetuptime.Rows.Count - 2
                    rowDGV = Me.dgvSetuptime.Rows(intA)
                    If intA <> vIndex Then
                        If vPrimaryProses1 = rowDGV.Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value AndAlso _
                        vPrimaryProses2 = rowDGV.Cells(Me.ColSetupTime_PrimaryProses2.Index).Value AndAlso _
                        vPrimaryProduct1 = rowDGV.Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value AndAlso _
                        vprimaryProduct2 = rowDGV.Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value Then
                            Return True
                        End If
                    End If
                Next
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function



    Private Sub dgvSetuptime_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSetupTime.Sorted

        Try
            Me.ObjCD.SetNumber(Me.dgvSetuptime)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try


    End Sub

    Private Sub dgvSetuptime_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvSetupTime.UserDeletedRow

        Try
            Me.ObjCD.SetNumber(Me.dgvSetuptime)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try


    End Sub

    Private Sub BackgroundWorker3_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Try
            Me.dgvSetuptime.Rows.RemoveAt(Me.EraseRowIndex)
            Me.ObjCD.SetNumber(Me.dgvSetuptime)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSet.Click
        Dim rowDGV As New DataGridViewRow
        Dim I As New Core_Module.FrmMasterSetupTime_Cari
        Try
            Me.ListOfIndex.Clear()

            If Me.dgvSetupTime.RowCount > 0 Then
                For intA As Integer = 0 To Me.dgvSetupTime.Rows.Count - 1
                    rowDGV = Me.dgvSetupTime.Rows(intA)
                    If rowDGV.Cells(Me.ColSetupTime_Set.Index).Value = Boolean.TrueString Then
                        If Me.ListOfIndex.Contains(intA) = False Then
                            Me.ListOfIndex.Add(intA)
                        End If
                    End If
                Next


                If Me.ListOfIndex.Count > 0 Then
                    I.ShowDialog()
                    If I.Primary > -1 Then
                        For intX As Integer = 0 To Me.ListOfIndex.Count - 1
                            rowDGV = Me.dgvSetupTime.Rows(Me.ListOfIndex(intX))
                            rowDGV.Cells(Me.ColSetupTime_PrimarySetupTime.Index).Value = I.Primary
                            rowDGV.Cells(Me.ColSetupTime_Code.Index).Value = Me.objSetup.GetCode(I.Primary)
                            rowDGV.Cells(Me.ColSetupTime_SetupTimeDuration.Index).Value = Me.ObjCD.Convert_ToString(Me.objSetup.GetDuration(I.Primary))
                        Next
                    End If
                Else
                    MsgBox("Please choose at lease 1 item. ", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            Else
                MsgBox("Please choose at lease 1 item. ", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If



        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Public Function ApakahDataBenar()
        Dim ListX As New List(Of Integer)

        Try
            If Me.dgvSetupTime.Rows.Count > 0 Then
                For intS As Integer = 0 To Me.dgvSetupTime.Rows.Count - 1
                    If Me.dgvSetupTime.Rows(intS).Cells(Me.ColSetupTime_PrimarySetupTime.Index).Value = -1 Then
                        ListX.Add(intS)
                    End If
                Next
            End If

            If ListX.Count > 0 Then
                Dim barisKe As String = ""
                For intX As Integer = 0 To ListX.Count - 1
                    If barisKe = "" Then
                        barisKe &= ListX.Item(intX) + 1
                    Else
                        barisKe &= ", " & ListX.Item(intX) + 1
                    End If
                Next

                If barisKe <> "" Then
                    MsgBox("There are data without setup time on row " & barisKe & ", please correct your data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If


            Return True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try


    End Function

    Private Sub LinkLabel1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblCheckAll.Click

        Try
            If Me.dgvSetupTime.Rows.Count > 0 Then
                Dim rowDGV As New DataGridViewRow
                For Each rowDGV In Me.dgvSetupTime.Rows
                    rowDGV.Cells(Me.ColSetupTime_Set.Index).Value = Boolean.TrueString
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblUncheckedAll.Click

        Try
            If Me.dgvSetupTime.Rows.Count > 0 Then
                Dim rowDGV As New DataGridViewRow
                For Each rowDGV In Me.dgvSetupTime.Rows
                    rowDGV.Cells(Me.ColSetupTime_Set.Index).Value = Boolean.FalseString
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try


    End Sub
End Class