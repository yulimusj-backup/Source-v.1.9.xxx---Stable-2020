Option Explicit On
Option Strict Off
Imports System.Data
Public Class FrmLaporanProd_List
    Public BAckColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public Usercode As String = ""

    Private Buffer As String = ""
    Private EraseRowIndexProses As Integer = 0
    Private EraseRowIndexWorkUnit As Integer = 0
    Private EraseRowIndexWorkShift As Integer = 0
    Private EraseRowIndexProduct As Integer = 0
    Private EraseRowIndexCategory As Integer = 0

    Private ObjCD As New Core_Module.ClsCekData
    Private ObjProcess As New Core_Module.ClsMasterProses
    Private ObjWorkShift As New Core_Module.ClsMasterWorkShift
    Private ObjWorkUnit As New Core_Module.ClsMasterWorkUnit
    Private objGroup As New Core_Module.ClsMasterProductGroup
    Private ObjCategory As New Core_Module.ClsMasterCategory
    Private ObjProduct As New Core_Module.clsMasterProduct
    Private ObjLapProd As New ClsLaporan_Prod_List
    Private ObjLaporanProdOpt As New FrmLaporanProd_List_Opt

    Dim Format_Index As EnumFormat
    Dim Format2_Index As EnumFormat2
    Dim GroupBy_Index As EnumGroupBy
    Dim LvlGroup As Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.ProductGroup
    Dim TBL As New DseLaporanProd_List.tbltrans_setting_report_stockbundling1DataTable
    Dim ObjLicense As New Core_Module.ClsLicense

#Region "Enum saja"
    Private Enum EnumFormat As Integer
        Recap
        Recap_w_group
        Detail
    End Enum

    Private Enum EnumFormat2 As Integer
        Input
        Output
        Both
    End Enum

    Private Enum EnumGroupBy As Integer
        _Date
        Document_No
        Work_Unit
        Work_Shift
        Process
    End Enum

    Private Enum EnumPilih
        No
        Yes
    End Enum
#End Region

#Region "Atur Tampilan"
    Private Sub cboFormat_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFormat.SelectedIndexChanged
        Try
            If Me.cboFormat.SelectedIndex = 0 Then
                Me.Format_Index = EnumFormat.Recap
                Me.dgvSort.Rows.Clear()
                If Me.dgvSort.Rows.Count = 0 Then
                    Me.AddSorting()
                End If
            ElseIf Me.cboFormat.SelectedIndex = 1 Then
                Me.Format_Index = EnumFormat.Recap_w_group
                Me.dgvSort.Rows.Clear()
                If Me.dgvSort.Rows.Count = 0 Then
                    Me.AddSorting()
                End If
            ElseIf Me.cboFormat.SelectedIndex = 2 Then
                Me.Format_Index = EnumFormat.Detail
                Me.dgvSort.Rows.Clear()
                If Me.dgvSort.Rows.Count = 0 Then
                    Me.AddSorting()
                End If
            End If
            If Me.Format_Index = EnumFormat.Recap_w_group Then
                Me.cboGroupBy.Enabled = True
                Me.cboGroupBy.SelectedIndex = 0
                If Me.GroupBy_Index = EnumGroupBy.Document_No Then
                    Me.dgvSort.Rows.Clear()
                Else
                    If Me.dgvSort.Rows.Count = 0 Then
                        Me.AddSorting()
                    End If
                End If
            Else
                Me.cboGroupBy.SelectedIndex = -1
                Me.cboGroupBy.Enabled = False
                Me.cboFormat2.Enabled = True
                If Me.dgvSort.Rows.Count = 0 Then
                    Me.AddSorting()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboFormat2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboFormat2.SelectedIndexChanged
        Try
            If Me.cboFormat2.SelectedIndex = 0 Then
                Me.Format2_Index = EnumFormat2.Input
            ElseIf Me.cboFormat2.SelectedIndex = 1 Then
                Me.Format2_Index = EnumFormat2.Output
            ElseIf Me.cboFormat2.SelectedIndex = 2 Then
                Me.Format2_Index = EnumFormat2.Both
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboGroupBy_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboGroupBy.SelectedIndexChanged
        Try
            If Me.cboGroupBy.SelectedIndex = 0 Then
                Me.GroupBy_Index = EnumGroupBy._Date
            ElseIf Me.cboGroupBy.SelectedIndex = 1 Then
                Me.GroupBy_Index = EnumGroupBy.Document_No
            ElseIf Me.cboGroupBy.SelectedIndex = 2 Then
                Me.GroupBy_Index = EnumGroupBy.Process
            ElseIf Me.cboGroupBy.SelectedIndex = 3 Then
                Me.GroupBy_Index = EnumGroupBy.Work_Unit
            ElseIf Me.cboGroupBy.SelectedIndex = 4 Then
                Me.GroupBy_Index = EnumGroupBy.Work_Shift
            End If
            If Me.GroupBy_Index = EnumGroupBy.Document_No Then
                Me.dgvSort.Rows.Clear()
            Else
                If Me.dgvSort.Rows.Count = 0 Then
                    Me.AddSorting()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            If Me.TabControl1.SelectedIndex = 1 Then
                Me.crv1.DisplayGroupTree = False
                Me.crv1.DisplayStatusBar = True
                Me.crv1.DisplayToolbar = True
                Me.crv1.ShowGroupTreeButton = False
                Me.crv1.ShowExportButton = True
                Me.crv1.ShowRefreshButton = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddSorting()
        Dim rowDGV As DataGridViewRow
        Try
            If Me.Format_Index <> EnumFormat.Detail Then
                'add sorting
                Me.dgvSort.Rows.Add()
                rowDGV = Me.dgvSort.Rows(Me.dgvSort.Rows.Count - 1)
                rowDGV.Cells(Me.ColSort_Primary.Index).Value = "-1"
                rowDGV.Cells(Me.ColSort_Desc.Index).Value = "Code No"
                rowDGV.Cells(Me.ColSort_Direction.Index).Value = "Ascending"
                rowDGV.Cells(Me.ColSort_Order.Index).Value = "1"
                Me.dgvSort.Rows.Add()
                rowDGV = Me.dgvSort.Rows(Me.dgvSort.Rows.Count - 1)
                rowDGV.Cells(Me.ColSort_Primary.Index).Value = "-1"
                rowDGV.Cells(Me.ColSort_Desc.Index).Value = "Process"
                rowDGV.Cells(Me.ColSort_Direction.Index).Value = "Ascending"
                rowDGV.Cells(Me.ColSort_Order.Index).Value = "2"
                Me.dgvSort.Rows.Add()
                rowDGV = Me.dgvSort.Rows(Me.dgvSort.Rows.Count - 1)
                rowDGV.Cells(Me.ColSort_Primary.Index).Value = "-1"
                rowDGV.Cells(Me.ColSort_Desc.Index).Value = "Work Unit"
                rowDGV.Cells(Me.ColSort_Direction.Index).Value = "Ascending"
                rowDGV.Cells(Me.ColSort_Order.Index).Value = "3"
                Me.dgvSort.Rows.Add()
                rowDGV = Me.dgvSort.Rows(Me.dgvSort.Rows.Count - 1)
                rowDGV.Cells(Me.ColSort_Primary.Index).Value = "-1"
                rowDGV.Cells(Me.ColSort_Desc.Index).Value = "Workshift"
                rowDGV.Cells(Me.ColSort_Direction.Index).Value = "Ascending"
                rowDGV.Cells(Me.ColSort_Order.Index).Value = "4"
            Else
                'add sorting
                Me.dgvSort.Rows.Add()
                rowDGV = Me.dgvSort.Rows(Me.dgvSort.Rows.Count - 1)
                rowDGV.Cells(Me.ColSort_Primary.Index).Value = "-1"
                rowDGV.Cells(Me.ColSort_Desc.Index).Value = "Date"
                rowDGV.Cells(Me.ColSort_Direction.Index).Value = "Ascending"
                rowDGV.Cells(Me.ColSort_Order.Index).Value = "1"
                Me.dgvSort.Rows.Add()
                rowDGV = Me.dgvSort.Rows(Me.dgvSort.Rows.Count - 1)
                rowDGV.Cells(Me.ColSort_Primary.Index).Value = "-1"
                rowDGV.Cells(Me.ColSort_Desc.Index).Value = "Document No"
                rowDGV.Cells(Me.ColSort_Direction.Index).Value = "Ascending"
                rowDGV.Cells(Me.ColSort_Order.Index).Value = "2"
                Me.dgvSort.Rows.Add()
                rowDGV = Me.dgvSort.Rows(Me.dgvSort.Rows.Count - 1)
                rowDGV.Cells(Me.ColSort_Primary.Index).Value = "-1"
                rowDGV.Cells(Me.ColSort_Desc.Index).Value = "Code No"
                rowDGV.Cells(Me.ColSort_Direction.Index).Value = "Ascending"
                rowDGV.Cells(Me.ColSort_Order.Index).Value = "3"
                Me.dgvSort.Rows.Add()
                rowDGV = Me.dgvSort.Rows(Me.dgvSort.Rows.Count - 1)
                rowDGV.Cells(Me.ColSort_Primary.Index).Value = "-1"
                rowDGV.Cells(Me.ColSort_Desc.Index).Value = "Process"
                rowDGV.Cells(Me.ColSort_Direction.Index).Value = "Ascending"
                rowDGV.Cells(Me.ColSort_Order.Index).Value = "4"
                Me.dgvSort.Rows.Add()
                rowDGV = Me.dgvSort.Rows(Me.dgvSort.Rows.Count - 1)
                rowDGV.Cells(Me.ColSort_Primary.Index).Value = "-1"
                rowDGV.Cells(Me.ColSort_Desc.Index).Value = "Work Unit"
                rowDGV.Cells(Me.ColSort_Direction.Index).Value = "Ascending"
                rowDGV.Cells(Me.ColSort_Order.Index).Value = "5"
                Me.dgvSort.Rows.Add()
                rowDGV = Me.dgvSort.Rows(Me.dgvSort.Rows.Count - 1)
                rowDGV.Cells(Me.ColSort_Primary.Index).Value = "-1"
                rowDGV.Cells(Me.ColSort_Desc.Index).Value = "Workshift"
                rowDGV.Cells(Me.ColSort_Direction.Index).Value = "Ascending"
                rowDGV.Cells(Me.ColSort_Order.Index).Value = "6"
            End If
          
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanProd_List_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.IsiCbo()
            'Me.AddSorting()
            'add group by
            Me.cboGroupBy.Items.Clear()
            Me.cboGroupBy.Items.Add("Date")
            Me.cboGroupBy.Items.Add("Document No")
            Me.cboGroupBy.Items.Add("Process")
            Me.cboGroupBy.Items.Add("Work Unit")
            Me.cboGroupBy.Items.Add("Work Shift")
            Me.Reset()
            Me.cboFormat.SelectedIndex = 0
            Me.cboFormat2.SelectedIndex = 0
            Me.CboPilihGroupLevel.Enabled = False
            Me.DgvGroup.AllowUserToAddRows = False
            Me.dgvProcess.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            Me.dgvSort.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            Me.DgvWorkShift.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            Me.dgvWorkUnit.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            Me.DgvGroup.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            Me.CancelButton = Me.BtnClose
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None

            'atur positioning
            Me.GbFilter.Dock = DockStyle.Fill
            Me.SplitContainer2.Width = Me.SplitContainer1.Panel1.Width - Me.SplitContainer2.Left
            Me.SplitContainer2.Height = Me.SplitContainer1.Panel1.Height - Me.SplitContainer2.Top
            Me.GbSort.Width = Me.SplitContainer4.Panel1.Width - Me.GbSort.Left
            'Me.GroupBox2.Width = Me.SplitContainer2.Panel1.Width - Me.GroupBox2.Left
            Me.GroupBox2.Dock = DockStyle.Fill
            Me.GroupBox4.Dock = DockStyle.Fill
            Me.GroupBox5.Dock = DockStyle.Fill
            Me.GroupBox7.Dock = DockStyle.Fill
            Me.SplitContainer3.Dock = DockStyle.Fill
            Me.dtpTo.Width = Me.SplitContainer1.Panel1.Width - Me.dtpTo.Left
            Me.dtpFrom.Width = Me.dtpTo.Width
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkWorkUnit_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkWorkUnit.CheckedChanged
        Try
            If Me.chkWorkUnit.Checked = False Then
                Me.dgvWorkUnit.Rows.Clear()
                Me.dgvWorkUnit.AllowUserToAddRows = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkWorkShift_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkWorkShift.CheckedChanged
        Try
            If Me.ChkWorkShift.Checked = False Then
                Me.DgvWorkShift.Rows.Clear()
                Me.DgvWorkShift.AllowUserToAddRows = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkProcess_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkProcess.CheckedChanged
        Try
            If Me.chkProcess.Checked = False Then
                Me.dgvProcess.Rows.Clear()
                Me.dgvProcess.AllowUserToAddRows = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkProductGroup_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkProductGroup.CheckedChanged
        Try
            If Me.ChkProductGroup.Checked = False Then
                Me.DgvGroup.Rows.Clear()
                Me.DgvGroup.AllowUserToAddRows = False
                Me.CboPilihGroupLevel.SelectedIndex = 0
                Me.CboPilihGroupLevel.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkProductGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkProductGroup.Click
        Try
            Me.DgvGroup.Rows.Clear()
            If Me.ChkProductGroup.Checked = True Then
                Me.DgvGroup.AllowUserToAddRows = True
                Me.CboPilihGroupLevel.Enabled = True
                Me.CboPilihGroupLevel.SelectedIndex = 0
                Me.DgvGroup.ReadOnly = False
            Else
                Me.DgvGroup.AllowUserToAddRows = False
                Me.CboPilihGroupLevel.Enabled = False
                Me.CboPilihGroupLevel.SelectedIndex = 0
                Me.DgvGroup.ReadOnly = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkWorkUnit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkWorkUnit.Click
        Try
            If Me.chkWorkUnit.Checked = True Then
                Me.dgvWorkUnit.ReadOnly = False
                Me.dgvWorkUnit.AllowUserToAddRows = True
            Else
                Me.dgvWorkUnit.ReadOnly = True
                Me.dgvWorkUnit.Rows.Clear()
                Me.dgvWorkUnit.AllowUserToAddRows = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkWorkShift_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkWorkShift.Click
        Try
            If Me.ChkWorkShift.Checked = True Then
                Me.DgvWorkShift.ReadOnly = False
                Me.DgvWorkShift.AllowUserToAddRows = True
            Else
                Me.DgvWorkShift.ReadOnly = True
                Me.DgvWorkShift.Rows.Clear()
                Me.DgvWorkShift.AllowUserToAddRows = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkProcess_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkProcess.Click
        Try
            If Me.chkProcess.Checked = True Then
                Me.dgvProcess.ReadOnly = False
                Me.dgvProcess.AllowUserToAddRows = True
            Else
                Me.dgvProcess.ReadOnly = True
                Me.dgvProcess.Rows.Clear()
                Me.dgvProcess.AllowUserToAddRows = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboPilihGroupLevel_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboPilihGroupLevel.SelectedIndexChanged
        Try
            Me.DgvGroup.Rows.Clear()
            If Me.CboPilihGroupLevel.SelectedIndex = 0 Then
                LvlGroup = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.ProductGroup
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = 1 Then
                LvlGroup = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.Brand
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = 2 Then
                LvlGroup = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.SubBrand
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = 3 Then
                LvlGroup = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.Category
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCbo()
        Dim DTA As New DseLaporanProd_ListTableAdapters.tbltrans_setting_report_stockbundling1TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Me.ObjLapProd.IsiTBL(TBL)
            Me.cboSavedReport.DataSource = Me.TBL.DefaultView
            Me.TBL.DefaultView.Sort = "Timestamp"
            cboSavedReport.DisplayMember = "Report Desc"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub
#End Region

#Region "DGV Sort"

    Private Sub dgvSort_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvSort.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.dgvSort.Item(Me.ColGroup_Code.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColSort_Order.Index Then
                    Me.Buffer = Me.dgvSort.Item(e.ColumnIndex, e.RowIndex).Value
                    'Else
                    '    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColSort_Order.Index Then
                    Me.Buffer = Me.dgvSort.Item(e.ColumnIndex, e.RowIndex).Value
                    'Else
                    '    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSort_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSort.CellEndEdit
        Try
            Dim row As DataGridViewRow
            Dim i As String = Me.dgvSort.Item(Me.ColSort_Order.Index, e.RowIndex).Value
            Dim RowIdx As Integer = e.RowIndex + 1
            Dim ijk As Integer = 1
            If IsNothing(Me.dgvSort.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColSort_Order.Index Then
                    If IsNumeric(Me.dgvSort.Item(Me.ColSort_Order.Index, e.RowIndex).Value) = False Then
                        Me.dgvSort.Item(Me.ColSort_Order.Index, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    If Me.Format_Index <> EnumFormat.Detail Then
                        If CDec(Me.dgvSort.Item(Me.ColSort_Order.Index, e.RowIndex).Value) > 4 Then
                            Me.dgvSort.Item(Me.ColSort_Order.Index, e.RowIndex).Value = Me.Buffer
                        Else
                            For Each row In dgvSort.Rows
                                If ijk <> RowIdx Then
                                    If Me.dgvSort.Item(Me.ColSort_Order.Index, e.RowIndex).Value = row.Cells(Me.ColSort_Order.Index).Value Then
                                        row.Cells(Me.ColSort_Order.Index).Value = Me.Buffer
                                    End If
                                End If
                                ijk += 1
                            Next
                        End If
                    Else
                        If CDec(Me.dgvSort.Item(Me.ColSort_Order.Index, e.RowIndex).Value) > 6 Then
                            Me.dgvSort.Item(Me.ColSort_Order.Index, e.RowIndex).Value = Me.Buffer
                        Else
                            For Each row In dgvSort.Rows
                                If ijk <> RowIdx Then
                                    If Me.dgvSort.Item(Me.ColSort_Order.Index, e.RowIndex).Value = row.Cells(Me.ColSort_Order.Index).Value Then
                                        row.Cells(Me.ColSort_Order.Index).Value = Me.Buffer
                                    End If
                                End If
                                ijk += 1
                            Next
                        End If
                    End If
                End If
            Else
                Me.dgvSort.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

#End Region

#Region "DGV Proses"
    Private Sub dgvProcess_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvProcess.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.dgvProcess.Item(Me.ColProcess_Code.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColProcess_Code.Index Then
                    '''set header fungsi digunakan ketika edit code 
                    Me.dgvProcess.Columns(Me.ColProcess_Code.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.dgvProcess.Columns(Me.ColProcess_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    '''simpan nilai awal di buffer
                    Me.Buffer = Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColProcess_Code.Index Then
                    '''set header fungsi digunakan ketika edit code 
                    Me.dgvProcess.Columns(Me.ColProcess_Code.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.dgvProcess.Columns(Me.ColProcess_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.Buffer = Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemProcess(ByVal vPrimary As Double, _
           ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            '''pastikan tidak ada code yang sama
            For ij As Integer = 0 To Me.dgvProcess.Rows.Count - 2
                RowDGV = Me.dgvProcess.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColProcess_PrimaryOriginal.Index).Value = vPrimary Then
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

    Private Sub dgvProcess_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvProcess.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvProcess.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryProcess As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColProcess_Code.Index Then
                        Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjProcess.ApakahSudahAda(Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryProcess = Me.ObjProcess.GetPrimary(Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemProcess(PrimaryProcess, e.RowIndex)
                            If index < 0 Then
                                Me.dgvProcess.Item(Me.ColProcess_Primary.Index, e.RowIndex).Value = "-1"
                                Me.dgvProcess.Item(Me.ColProcess_PrimaryOriginal.Index, e.RowIndex).Value = PrimaryProcess
                                Me.dgvProcess.Item(Me.ColProcess_Code.Index, e.RowIndex).Value = Me.ObjProcess.GetCode(PrimaryProcess)
                                Me.dgvProcess.Item(Me.ColProcess_Desc.Index, e.RowIndex).Value = Me.ObjProcess.GetDesc(PrimaryProcess)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvProcess.Item(Me.ColProcess_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexProses = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvProcess)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmMasterProsesCari
                            ij.txtCari.Text = Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BAckColorAll
                            ij.ShowDialog()
                            If ij.mPrimary <> -1 Then

                                index = Me.GetIndexItemProcess(ij.mPrimary, e.RowIndex)
                                If index < 0 Then
                                    Me.dgvProcess.Item(Me.ColProcess_Primary.Index, e.RowIndex).Value = "-1"
                                    Me.dgvProcess.Item(Me.ColProcess_PrimaryOriginal.Index, e.RowIndex).Value = ij.mPrimary
                                    Me.dgvProcess.Item(Me.ColProcess_Code.Index, e.RowIndex).Value = Me.ObjProcess.GetCode(ij.mPrimary)
                                    Me.dgvProcess.Item(Me.ColProcess_Desc.Index, e.RowIndex).Value = Me.ObjProcess.GetDesc(ij.mPrimary)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndexProses = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvProcess)
                                        Exit Sub
                                    Else
                                        Me.dgvProcess.Item(Me.ColProcess_Primary.Index, e.RowIndex).Value = "-1"
                                        Me.dgvProcess.Item(Me.ColProcess_PrimaryOriginal.Index, e.RowIndex).Value = ij.mPrimary
                                        Me.dgvProcess.Item(Me.ColProcess_Code.Index, e.RowIndex).Value = Me.ObjProcess.GetCode(ij.mPrimary)
                                        Me.dgvProcess.Item(Me.ColProcess_Desc.Index, e.RowIndex).Value = Me.ObjProcess.GetDesc(ij.mPrimary)
                                    End If
                                End If
                            ElseIf ij.mPrimary = -1 Then
                                If Me.dgvProcess.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexProses = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvProcess)
                                    Exit Sub
                                End If
                                If Me.dgvProcess.Item(Me.ColProcess_Code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexProses = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvProcess)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvProcess)
                    End If
                Else
                    If Me.dgvProcess.Rows.Count - 1 <> e.RowIndex Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexProses = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        Else
                            Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvProcess)
                        Exit Sub
                    End If
                    If Me.dgvProcess.Item(Me.ColProcess_Code.Index, e.RowIndex).Value <> "" Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexProses = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        Else
                            Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvProcess)
                        Exit Sub
                    End If
                End If
                'Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvProcess.Columns(Me.ColProcess_Code.Index).SortMode = DataGridViewColumnSortMode.Automatic
            Me.dgvProcess.Columns(Me.ColProcess_Desc.Index).SortMode = DataGridViewColumnSortMode.Automatic
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.dgvProcess.Rows.RemoveAt(EraseRowIndexProses)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "DGV Work Shift"
    Private Sub DgvWorkShift_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvWorkShift.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.DgvWorkShift.Item(Me.ColWorkShift_Code.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColWorkShift_Code.Index Then
                    '''set header fungsi digunakan ketika edit code 
                    Me.DgvWorkShift.Columns(Me.ColWorkShift_Code.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.DgvWorkShift.Columns(Me.ColWorkShift_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    '''simpan nilai awal di buffer
                    Me.Buffer = Me.DgvWorkShift.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColWorkShift_Code.Index Then
                    '''set header fungsi digunakan ketika edit code 
                    Me.DgvWorkShift.Columns(Me.ColWorkShift_Code.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.DgvWorkShift.Columns(Me.ColWorkShift_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.Buffer = Me.DgvWorkShift.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemWorkShift(ByVal vPrimary As Double, _
              ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            '''pastikan tidak ada code yang sama
            For ij As Integer = 0 To Me.DgvWorkShift.Rows.Count - 2
                RowDGV = Me.DgvWorkShift.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColWorkShift_PrimaryOriginal.Index).Value = vPrimary Then
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

    Private Sub DgvWorkShift_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvWorkShift.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvWorkShift.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryWorkShift As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.DgvWorkShift.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColWorkShift_Code.Index Then
                        Me.DgvWorkShift.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvWorkShift.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjWorkShift.ApakahSudahAda(Me.DgvWorkShift.Item(e.ColumnIndex, e.RowIndex).Value, Core_Module.ClsMasterWorkShift.EnumTipeWorkshift.Production) = True Then
                            PrimaryWorkShift = Me.ObjWorkShift.GetPrimary(Me.DgvWorkShift.Item(e.ColumnIndex, e.RowIndex).Value, Core_Module.ClsMasterWorkShift.EnumTipeWorkshift.Production)
                            index = Me.GetIndexItemWorkShift(PrimaryWorkShift, e.RowIndex)
                            If index < 0 Then
                                Me.DgvWorkShift.Item(Me.ColWorkShift_Primary.Index, e.RowIndex).Value = "-1"
                                Me.DgvWorkShift.Item(Me.ColWorkShift_PrimaryOriginal.Index, e.RowIndex).Value = PrimaryWorkShift
                                Me.DgvWorkShift.Item(Me.ColWorkShift_Code.Index, e.RowIndex).Value = Me.ObjWorkShift.GetCodeByPrimary(PrimaryWorkShift)
                                Me.DgvWorkShift.Item(Me.ColWorkShift_Desc.Index, e.RowIndex).Value = Me.ObjWorkShift.GetDescription(PrimaryWorkShift)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvWorkShift.Item(Me.ColWorkShift_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexWorkShift = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvWorkShift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvWorkShift)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmMasterWorkShift_Cari(Core_Module.ClsMasterWorkShift.EnumTipeWorkshift.Production)
                            ij.txtCari.Text = Me.DgvWorkShift.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BAckColorAll
                            ij.ShowDialog()
                            If ij.PrimaryWorkshift <> -1 Then
                                index = Me.GetIndexItemWorkShift(ij.PrimaryWorkshift, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvWorkShift.Item(Me.ColWorkShift_Primary.Index, e.RowIndex).Value = "-1"
                                    Me.DgvWorkShift.Item(Me.ColWorkShift_PrimaryOriginal.Index, e.RowIndex).Value = ij.PrimaryWorkshift
                                    Me.DgvWorkShift.Item(Me.ColWorkShift_Code.Index, e.RowIndex).Value = Me.ObjWorkShift.GetCodeByPrimary(ij.PrimaryWorkshift)
                                    Me.DgvWorkShift.Item(Me.ColWorkShift_Desc.Index, e.RowIndex).Value = Me.ObjWorkShift.GetDescription(ij.PrimaryWorkshift)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndexWorkShift = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.DgvWorkShift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvWorkShift)
                                        Exit Sub
                                    Else
                                        Me.DgvWorkShift.Item(Me.ColWorkShift_Primary.Index, e.RowIndex).Value = "-1"
                                        Me.DgvWorkShift.Item(Me.ColWorkShift_PrimaryOriginal.Index, e.RowIndex).Value = ij.PrimaryWorkshift
                                        Me.DgvWorkShift.Item(Me.ColWorkShift_Code.Index, e.RowIndex).Value = Me.ObjWorkShift.GetCodeByPrimary(ij.PrimaryWorkshift)
                                        Me.DgvWorkShift.Item(Me.ColWorkShift_Desc.Index, e.RowIndex).Value = Me.ObjWorkShift.GetDescription(ij.PrimaryWorkshift)
                                    End If
                                End If
                            ElseIf ij.PrimaryWorkshift = -1 Then
                                If Me.DgvWorkShift.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexWorkShift = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvWorkShift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvWorkShift)
                                    Exit Sub
                                End If
                                If Me.DgvWorkShift.Item(Me.ColWorkShift_Code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexWorkShift = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvWorkShift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvWorkShift)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvWorkShift)
                    End If
                Else
                    If Me.DgvWorkShift.Rows.Count - 1 <> e.RowIndex Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexWorkShift = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvWorkShift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvWorkShift)
                        Exit Sub
                    End If
                    If Me.DgvWorkShift.Item(Me.ColWorkShift_Code.Index, e.RowIndex).Value <> "" Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexWorkShift = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvWorkShift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvWorkShift)
                        Exit Sub
                    End If
                End If
                'Me.dgvWorkShift.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DgvWorkShift.Columns(Me.ColWorkShift_Code.Index).SortMode = DataGridViewColumnSortMode.Automatic
            Me.DgvWorkShift.Columns(Me.ColWorkShift_Desc.Index).SortMode = DataGridViewColumnSortMode.Automatic
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.DgvWorkShift.Rows.RemoveAt(EraseRowIndexWorkShift)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "DGV Work Unit"
    Private Sub dgvWorkUnit_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvWorkUnit.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.dgvWorkUnit.Item(Me.ColWorkUnit_Code.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColWorkUnit_Code.Index Then
                    '''set header fungsi digunakan ketika edit code 
                    Me.dgvWorkUnit.Columns(Me.ColWorkUnit_Code.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.dgvWorkUnit.Columns(Me.ColWorkUnit_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    '''simpan nilai awal di buffer
                    Me.Buffer = Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColWorkUnit_Code.Index Then
                    '''set header fungsi digunakan ketika edit code 
                    Me.dgvWorkUnit.Columns(Me.ColWorkUnit_Code.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.dgvWorkUnit.Columns(Me.ColWorkUnit_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.Buffer = Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemWorkUnit(ByVal vPrimary As Double, _
                 ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            '''pastikan tidak ada code yang sama
            For ij As Integer = 0 To Me.dgvWorkUnit.Rows.Count - 2
                RowDGV = Me.dgvWorkUnit.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColWorkUnit_PrimaryOriginal.Index).Value = vPrimary Then
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

    Private Sub dgvWorkUnit_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvWorkUnit.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvWorkUnit.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryWorkUnit As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColWorkUnit_Code.Index Then
                        Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjWorkUnit.ApakahSudahAda(Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryWorkUnit = Me.ObjWorkUnit.GetPrimary(Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemWorkUnit(PrimaryWorkUnit, e.RowIndex)
                            If index < 0 Then
                                Me.dgvWorkUnit.Item(Me.ColWorkUnit_Primary.Index, e.RowIndex).Value = "-1"
                                Me.dgvWorkUnit.Item(Me.ColWorkUnit_PrimaryOriginal.Index, e.RowIndex).Value = PrimaryWorkUnit
                                Me.dgvWorkUnit.Item(Me.ColWorkUnit_Code.Index, e.RowIndex).Value = Me.ObjWorkUnit.GetCode(PrimaryWorkUnit)
                                Me.dgvWorkUnit.Item(Me.ColWorkUnit_Desc.Index, e.RowIndex).Value = Me.ObjWorkUnit.GetDesc(PrimaryWorkUnit)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvWorkUnit.Item(Me.ColWorkUnit_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexWorkUnit = e.RowIndex
                                        Me.BackgroundWorker3.RunWorkerAsync()
                                    Else
                                        Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvWorkUnit)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmMasterWorkUnit_Cari
                            ij.txtCari.Text = Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BAckColorAll
                            ij.Status = "Only Active"
                            ij.ShowDialog()
                            If ij.PrimaryWorkUnit <> -1 Then

                                index = Me.GetIndexItemWorkUnit(ij.PrimaryWorkUnit, e.RowIndex)
                                If index < 0 Then
                                    Me.dgvWorkUnit.Item(Me.ColWorkUnit_Primary.Index, e.RowIndex).Value = "-1"
                                    Me.dgvWorkUnit.Item(Me.ColWorkUnit_PrimaryOriginal.Index, e.RowIndex).Value = ij.PrimaryWorkUnit
                                    Me.dgvWorkUnit.Item(Me.ColWorkUnit_Code.Index, e.RowIndex).Value = Me.ObjWorkUnit.GetCode(ij.PrimaryWorkUnit)
                                    Me.dgvWorkUnit.Item(Me.ColWorkUnit_Desc.Index, e.RowIndex).Value = Me.ObjWorkUnit.GetDesc(ij.PrimaryWorkUnit)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndexWorkUnit = e.RowIndex
                                            Me.BackgroundWorker3.RunWorkerAsync()
                                        Else
                                            Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvWorkUnit)
                                        Exit Sub
                                    Else
                                        Me.dgvWorkUnit.Item(Me.ColWorkUnit_Primary.Index, e.RowIndex).Value = "-1"
                                        Me.dgvWorkUnit.Item(Me.ColWorkUnit_PrimaryOriginal.Index, e.RowIndex).Value = ij.PrimaryWorkUnit
                                        Me.dgvWorkUnit.Item(Me.ColWorkUnit_Code.Index, e.RowIndex).Value = Me.ObjWorkUnit.GetCode(ij.PrimaryWorkUnit)
                                        Me.dgvWorkUnit.Item(Me.ColWorkUnit_Desc.Index, e.RowIndex).Value = Me.ObjWorkUnit.GetDesc(ij.PrimaryWorkUnit)
                                    End If
                                End If
                            ElseIf ij.PrimaryWorkUnit = -1 Then
                                If Me.dgvWorkUnit.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexWorkUnit = e.RowIndex
                                        Me.BackgroundWorker3.RunWorkerAsync()
                                    Else
                                        Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvWorkUnit)
                                    Exit Sub
                                End If
                                If Me.dgvWorkUnit.Item(Me.ColWorkUnit_Code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexWorkUnit = e.RowIndex
                                        Me.BackgroundWorker3.RunWorkerAsync()
                                    Else
                                        Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvWorkUnit)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvWorkUnit)
                    End If
                Else
                    If Me.dgvWorkUnit.Rows.Count - 1 <> e.RowIndex Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexWorkUnit = e.RowIndex
                            Me.BackgroundWorker3.RunWorkerAsync()
                        Else
                            Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvWorkUnit)
                        Exit Sub
                    End If
                    If Me.dgvWorkUnit.Item(Me.ColWorkUnit_Code.Index, e.RowIndex).Value <> "" Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexWorkUnit = e.RowIndex
                            Me.BackgroundWorker3.RunWorkerAsync()
                        Else
                            Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvWorkUnit)
                        Exit Sub
                    End If
                End If
                'Me.dgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvWorkUnit.Columns(Me.ColWorkUnit_Code.Index).SortMode = DataGridViewColumnSortMode.Automatic
            Me.dgvWorkUnit.Columns(Me.ColWorkUnit_Desc.Index).SortMode = DataGridViewColumnSortMode.Automatic
        End Try
    End Sub

    Private Sub BackgroundWorker3_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Try
            Me.dgvWorkUnit.Rows.RemoveAt(EraseRowIndexWorkUnit)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "DGV Product Group"

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
                        'kemungkinan
                        If CboPilihGroupLevel.SelectedIndex = 0 Then
                            LvlGroup = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.ProductGroup
                            _Group = True
                        ElseIf CboPilihGroupLevel.SelectedIndex = 1 Then
                            LvlGroup = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.Brand
                            _Group = True
                        ElseIf CboPilihGroupLevel.SelectedIndex = 2 Then
                            LvlGroup = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.SubBrand
                            _Group = True
                        Else
                            LvlGroup = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.Category
                            _Group = False
                        End If
                        If _Group = True Then
                            If Me.objGroup.ApakahSudahAda(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value, LvlGroup) = True Then
                                PrimaryGroup = Me.objGroup.GetPrimary(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value, LvlGroup)
                                index = Me.GetIndexItemGroup(PrimaryGroup, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = "-1"
                                    Me.DgvGroup.Item(Me.ColGroup_PrimaryOriginal.Index, e.RowIndex).Value = PrimaryGroup
                                    Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.objGroup.GetGroupCode(PrimaryGroup)
                                    Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.objGroup.GetGroupDesc(PrimaryGroup)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = ""
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndexProduct = e.RowIndex
                                            Me.BackgroundWorker4.RunWorkerAsync()
                                        Else
                                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                                        Exit Sub
                                    End If
                                End If
                            Else
                                Dim ij As New Core_Module.FrmCariMasterProductGroup
                                ij.ProductGroupLevel = LvlGroup
                                ij.txtCari.Text = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value
                                ij.BackColorAll = Me.BAckColorAll
                                ij.ShowDialog()
                                If ij.PrimaryProductGroup <> -1 Then

                                    index = Me.GetIndexItemGroup(ij.PrimaryProductGroup, e.RowIndex)
                                    If index < 0 Then
                                        Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = "-1"
                                        Me.DgvGroup.Item(Me.ColGroup_PrimaryOriginal.Index, e.RowIndex).Value = ij.PrimaryProductGroup
                                        Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.objGroup.GetGroupCode(ij.PrimaryProductGroup)
                                        Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.objGroup.GetGroupDesc(ij.PrimaryProductGroup)
                                    Else
                                        If e.RowIndex <> index Then
                                            MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                                Me.EraseRowIndexProduct = e.RowIndex
                                                Me.BackgroundWorker4.RunWorkerAsync()
                                            Else
                                                Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                            End If
                                            'Me.ObjCD.SetNumber(Me.dgvGroup)
                                            Exit Sub
                                        Else
                                            Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = "-1"
                                            Me.DgvGroup.Item(Me.ColGroup_PrimaryOriginal.Index, e.RowIndex).Value = ij.PrimaryProductGroup
                                            Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.objGroup.GetGroupCode(ij.PrimaryProductGroup)
                                            Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.objGroup.GetGroupDesc(ij.PrimaryProductGroup)
                                        End If
                                    End If
                                ElseIf ij.PrimaryProductGroup = -1 Then
                                    If Me.DgvGroup.Rows.Count - 1 <> e.RowIndex Then
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndexProduct = e.RowIndex
                                            Me.BackgroundWorker4.RunWorkerAsync()
                                        Else
                                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                                        Exit Sub
                                    End If
                                    If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndexProduct = e.RowIndex
                                            Me.BackgroundWorker4.RunWorkerAsync()
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
                        Else
                            EndEditCategory(sender, e)
                        End If

                    End If
                Else
                    If Me.DgvGroup.Rows.Count - 1 <> e.RowIndex Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker4.RunWorkerAsync()
                        Else
                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                        Exit Sub
                    End If
                    If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker4.RunWorkerAsync()
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
                                Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = "-1"
                                Me.DgvGroup.Item(Me.ColGroup_PrimaryOriginal.Index, e.RowIndex).Value = PrimaryGroup
                                Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.ObjCategory.GetCategoryCode(PrimaryGroup)
                                Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(PrimaryGroup)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker4.RunWorkerAsync()
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
                            ij.BackColorAll = Me.BAckColorAll
                            ij.ShowDialog()
                            If ij.PrimaryCategory <> -1 Then

                                index = Me.GetIndexItemGroup(ij.PrimaryCategory, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = "-1"
                                    Me.DgvGroup.Item(Me.ColGroup_PrimaryOriginal.Index, e.RowIndex).Value = ij.PrimaryCategory
                                    Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.ObjCategory.GetCategoryCode(ij.PrimaryCategory)
                                    Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(ij.PrimaryCategory)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndexProduct = e.RowIndex
                                            Me.BackgroundWorker4.RunWorkerAsync()
                                        Else
                                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                                        Exit Sub
                                    Else
                                        Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = "-1"
                                        Me.DgvGroup.Item(Me.ColGroup_PrimaryOriginal.Index, e.RowIndex).Value = ij.PrimaryCategory
                                        Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.ObjCategory.GetCategoryCode(ij.PrimaryCategory)
                                        Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(ij.PrimaryCategory)
                                    End If
                                End If
                            ElseIf ij.PrimaryCategory = -1 Then
                                If Me.DgvGroup.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker4.RunWorkerAsync()
                                    Else
                                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvGroup)
                                    Exit Sub
                                End If
                                If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker4.RunWorkerAsync()
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
                            Me.BackgroundWorker4.RunWorkerAsync()
                        Else
                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                        Exit Sub
                    End If
                    If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker4.RunWorkerAsync()
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
                    If RowDGV.Cells(Me.ColGroup_PrimaryOriginal.Index).Value = vPrimary Then
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

    Private Sub BackgroundWorker4_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker4.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker4_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker4.RunWorkerCompleted
        Try
            Me.DgvGroup.Rows.RemoveAt(EraseRowIndexProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "Atur Simpannya Disini"
    Public Sub SalinKeObjek(ByVal ij As FrmLaporanProd_List_Opt)
        Try
            With ObjLapProd
                .Desc = ij.TxtDesc.Text
                .FromDate = Format(Me.dtpFrom.Value, "dd MMM yyyy")
                .ToDate = Format(Me.dtpTo.Value, "dd MMM yyyy")
                'GroupLevel
                If Me.CboPilihGroupLevel.SelectedIndex = 0 Then
                    .GroupLevel = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.ProductGroup
                ElseIf Me.CboPilihGroupLevel.SelectedIndex = 1 Then
                    .GroupBy = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.Brand
                ElseIf Me.CboPilihGroupLevel.SelectedIndex = 2 Then
                    .GroupBy = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.SubBrand
                ElseIf Me.CboPilihGroupLevel.SelectedIndex = 3 Then
                    .GroupBy = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.Category
                End If
                'Format 1
                If Me.cboFormat.SelectedIndex = 0 Then
                    .Format1 = ClsLaporan_Prod_List.EnumFormat.Recap
                ElseIf Me.cboFormat.SelectedIndex = 1 Then
                    .Format1 = ClsLaporan_Prod_List.EnumFormat.Recap_w_group
                ElseIf Me.cboFormat.SelectedIndex = 2 Then
                    .Format1 = ClsLaporan_Prod_List.EnumFormat.Detail
                End If
                'Format 2
                If Me.cboFormat2.SelectedIndex = 0 Then
                    .Format2 = ClsLaporan_Prod_List.EnumFormat2.Input
                ElseIf Me.cboFormat2.SelectedIndex = 1 Then
                    .Format2 = ClsLaporan_Prod_List.EnumFormat2.Output
                ElseIf Me.cboFormat2.SelectedIndex = 2 Then
                    .Format2 = ClsLaporan_Prod_List.EnumFormat2.Both
                End If
                'Format By Group
                If Me.cboGroupBy.SelectedIndex = -1 Then
                    .GroupBy = -1
                ElseIf Me.cboGroupBy.SelectedIndex = 0 Then
                    .GroupBy = ClsLaporan_Prod_List.EnumGroupBy._Date
                ElseIf Me.cboGroupBy.SelectedIndex = 1 Then
                    .GroupBy = ClsLaporan_Prod_List.EnumGroupBy._Date
                ElseIf Me.cboGroupBy.SelectedIndex = 2 Then
                    .GroupBy = ClsLaporan_Prod_List.EnumGroupBy.Work_Unit
                ElseIf Me.cboGroupBy.SelectedIndex = 3 Then
                    .GroupBy = ClsLaporan_Prod_List.EnumGroupBy.Work_Shift
                ElseIf Me.cboGroupBy.SelectedIndex = 4 Then
                    .GroupBy = ClsLaporan_Prod_List.EnumGroupBy.Process
                End If
                ' chknya
                If Me.ChkShowSubTotal.Checked = False Then
                    .ShowsubTotal = ClsLaporan_Prod_List.EnumShowSubTotal.No
                Else
                    .ShowsubTotal = ClsLaporan_Prod_List.EnumShowSubTotal.Yes
                End If

                If Me.chkProcess.Checked = False Then
                    .FilterByProcess = ClsLaporan_Prod_List.EnumFilter.No
                Else
                    .FilterByProcess = ClsLaporan_Prod_List.EnumFilter.YEs
                End If
                If Me.chkWorkUnit.Checked = False Then
                    .FilterByWorkUnit = ClsLaporan_Prod_List.EnumFilter.No
                Else
                    .FilterByWorkUnit = ClsLaporan_Prod_List.EnumFilter.YEs
                End If
                If Me.ChkWorkShift.Checked = False Then
                    .FilterByWorkshift = ClsLaporan_Prod_List.EnumFilter.No
                Else
                    .FilterByWorkshift = ClsLaporan_Prod_List.EnumFilter.YEs
                End If
                If Me.ChkProductGroup.Checked = False Then
                    .FilterByProductGroup = ClsLaporan_Prod_List.EnumFilter.No
                Else
                    .FilterByProductGroup = ClsLaporan_Prod_List.EnumFilter.YEs
                End If
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeField()
        Try
            With ObjLapProd

                Me.TxtPrimary.Text = .Primary
                Me.dtpFrom.Value = .FromDate
                Me.dtpTo.Value = .ToDate

                'Format1
                If .Format1 = ClsLaporan_Prod_List.EnumFormat.Recap Then
                    Me.cboFormat.SelectedIndex = 0
                ElseIf .Format1 = ClsLaporan_Prod_List.EnumFormat.Recap_w_group Then
                    Me.cboFormat.SelectedIndex = 1
                ElseIf .Format1 = ClsLaporan_Prod_List.EnumFormat.Detail Then
                    Me.cboFormat.SelectedIndex = 2
                End If

                'Format2
                If .Format2 = ClsLaporan_Prod_List.EnumFormat2.Input Then
                    Me.cboFormat2.SelectedIndex = 0
                ElseIf .Format2 = ClsLaporan_Prod_List.EnumFormat2.Output Then
                    Me.cboFormat2.SelectedIndex = 1
                ElseIf .Format2 = ClsLaporan_Prod_List.EnumFormat2.Both Then
                    Me.cboFormat2.SelectedIndex = 2
                End If

                'Group By
                If .GroupBy = -1 Then
                    Me.cboGroupBy.SelectedIndex = -1
                ElseIf .GroupBy = ClsLaporan_Prod_List.EnumGroupBy._Date Then
                    Me.cboGroupBy.SelectedIndex = 0
                ElseIf .GroupBy = ClsLaporan_Prod_List.EnumGroupBy.Document_No Then
                    Me.cboGroupBy.SelectedIndex = 1
                ElseIf .GroupBy = ClsLaporan_Prod_List.EnumGroupBy.Process Then
                    Me.cboGroupBy.SelectedIndex = 2
                ElseIf .GroupBy = ClsLaporan_Prod_List.EnumGroupBy.Work_Unit Then
                    Me.cboGroupBy.SelectedIndex = 3
                ElseIf .GroupBy = ClsLaporan_Prod_List.EnumGroupBy.Work_Shift Then
                    Me.cboGroupBy.SelectedIndex = 4
                End If

                'grouplevel
                If .GroupLevel = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.ProductGroup Then
                    Me.CboPilihGroupLevel.SelectedIndex = 0
                ElseIf .GroupLevel = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.Brand Then
                    Me.CboPilihGroupLevel.SelectedIndex = 1
                ElseIf .GroupLevel = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.SubBrand Then
                    Me.CboPilihGroupLevel.SelectedIndex = 2
                ElseIf .GroupLevel = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.Category Then
                    Me.CboPilihGroupLevel.SelectedIndex = 3
                End If

                'chekbox
                If .ShowsubTotal = ClsLaporan_Prod_List.EnumShowSubTotal.No Then
                    Me.ChkShowSubTotal.Checked = False
                Else
                    Me.ChkShowSubTotal.Checked = True
                End If

                If .FilterByProcess = ClsLaporan_Prod_List.EnumFilter.No Then
                    Me.chkProcess.Checked = False
                Else
                    Me.chkProcess.Checked = True
                End If

                If .FilterByWorkUnit = ClsLaporan_Prod_List.EnumFilter.No Then
                    Me.chkWorkUnit.Checked = False
                Else
                    Me.chkWorkUnit.Checked = True
                End If

                If .FilterByWorkshift = ClsLaporan_Prod_List.EnumFilter.No Then
                    Me.ChkWorkShift.Checked = False
                Else
                    Me.ChkWorkShift.Checked = True
                End If

                If .FilterByProductGroup = ClsLaporan_Prod_List.EnumFilter.No Then
                    Me.ChkProductGroup.Checked = False
                Else
                    Me.ChkProductGroup.Checked = True
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim ij As New FrmLaporanProd_List_Opt
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.dtpFrom.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpFrom.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpTo.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpTo.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If Me.cboSavedReport.Text <> "" Then ij.TxtDesc.Text = Me.cboSavedReport.Text
            If ApakahDataBenar(False) = False Then Exit Sub
            If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                If Me.ApakahDataBenar(ij.TxtDesc.Text) = True Then
                    Me.SalinKeObjek(ij)
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    If Me.cboSavedReport.SelectedIndex = 0 And Me.TxtPrimary.Text = "-1" Then
                        If Me.ObjLapProd.Simpan(Me.dgvSort, Me.DgvGroup, Me.dgvWorkUnit, Me.DgvWorkShift, Me.dgvProcess, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.IsiCbo()
                            If Me.cboSavedReport.Items.Count > 0 Then
                                Me.cboSavedReport.SelectedIndex = Me.cboSavedReport.Items.Count - 1
                            End If
                        End If
                    Else
                        If Me.ObjLapProd.SimpanPerubahan(Me.dgvSort, Me.DgvGroup, Me.dgvWorkUnit, Me.DgvWorkShift, Me.dgvProcess, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.IsiCbo()
                            If Me.cboSavedReport.Items.Count > 0 Then
                                Me.cboSavedReport.SelectedIndex = Me.cboSavedReport.Items.Count - 1
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub Reset()
        Try
            Me.TxtPrimary.Text = "-1"
            Me.dtpFrom.Value = Now
            Me.dtpTo.Value = Now
            Me.chkProcess.Checked = False
            Me.ChkShowSubTotal.Checked = False
            Me.ChkProductGroup.Checked = False
            Me.ChkWorkShift.Checked = False
            Me.chkWorkUnit.Checked = False
            Me.CboPilihGroupLevel.Enabled = False
            Me.CboPilihGroupLevel.SelectedIndex = 0
            Me.DgvGroup.Rows.Clear()
            Me.dgvProcess.Rows.Clear()
            Me.dgvSort.Rows.Clear()
            Me.DgvWorkShift.Rows.Clear()
            Me.dgvWorkUnit.Rows.Clear()

            Me.DgvGroup.AllowUserToAddRows = False
            Me.dgvProcess.AllowUserToAddRows = False
            Me.DgvWorkShift.AllowUserToAddRows = False
            Me.dgvWorkUnit.AllowUserToAddRows = False
            Me.AddSorting()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function ApakahDataBenar(ByVal VDesc As String, Optional ByVal TanpaDesc As Boolean = False) As Boolean
        Dim Row As DataGridViewRow
        Try
            '0. 
            If Me.TxtPrimary.Text = "-1" And Me.cboSavedReport.SelectedIndex = 0 Then
                If Me.ObjLapProd.ApakahSudahada(Me.cboSavedReport.Text) = True Then
                    MsgBox("Duplication in description, please change your description", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            '1. cek cbo format1
            If Me.cboFormat.SelectedIndex = -1 Then
                Return False
            End If
            '2. cek cbo format2
            If Me.cboFormat2.SelectedIndex = -1 Then
                Return False
            End If

            If TanpaDesc = False Then
                If Trim(VDesc) = "" Then
                    MsgBox("Description cannot empty string", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If Me.chkProcess.Checked = True Then
                If dgvProcess.RowCount = 1 Then
                    For Each Row In dgvProcess.Rows
                        If Row.Cells(Me.ColProcess_Code.Index).Value = "" Then
                            MsgBox("Please input item in filter process", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    Next
                End If
            End If

            If Me.ChkProductGroup.Checked = True Then
                If DgvGroup.RowCount = 1 Then
                    For Each Row In DgvGroup.Rows
                        If Row.Cells(Me.ColGroup_Code.Index).Value = "" Then
                            MsgBox("Please input item in filter product group", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    Next
                End If
            End If

            If Me.ChkWorkShift.Checked = True Then
                If DgvWorkShift.RowCount = 1 Then
                    For Each Row In DgvWorkShift.Rows
                        If Row.Cells(Me.ColWorkShift_Code.Index).Value = "" Then
                            MsgBox("Please input item in filter workshift", MsgBoxStyle.Information, "Information")

                            Return False
                        End If
                    Next
                End If
            End If

            If Me.chkWorkUnit.Checked = True Then
                If dgvWorkUnit.RowCount = 1 Then
                    For Each Row In dgvWorkUnit.Rows
                        If Row.Cells(Me.ColWorkUnit_Code.Index).Value = "" Then
                            MsgBox("Please input item in filter workunit", MsgBoxStyle.Information, "Information")

                            Return False
                        End If
                    Next
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub cboSavedReport_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSavedReport.SelectedValueChanged
        Try
            If Me.cboSavedReport.Text <> "" Then
                Me.TxtPrimary.Text = ObjLapProd.GetPrimary(Me.cboSavedReport.Text)
                Me.ObjLapProd.CariDetail(Me.dgvSort, Me.DgvGroup, Me.dgvWorkUnit, Me.DgvWorkShift, Me.dgvProcess, Me.TxtPrimary.Text, Me)
                Me.SalinKeField()
                If Me.ChkProductGroup.Checked = True And Me.DgvGroup.RowCount > 0 Then
                    Me.CboPilihGroupLevel.Enabled = True
                End If
            Else
                Me.TxtPrimary.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "Atur Cetakannya Disini"
    Private Sub btnShow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.dtpFrom.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpFrom.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpTo.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpTo.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If Me.dtpFrom.Value.Date > Me.dtpTo.Value.Date Then
                MsgBox("Starting period must be smaller than ending period", MsgBoxStyle.Information, "Information")
                Me.dtpFrom.Focus()
                Exit Sub
            End If

            If Format_Index = EnumFormat.Recap_w_group Then
                If Me.GroupBy_Index = EnumGroupBy.Document_No Then
                    Me.CetakStockBundling_Recap_with_Group()
                Else
                    Me.CetakStockBundling_Recap()
                End If
            ElseIf Format_Index = EnumFormat.Recap Then
                Me.CetakStockBundling_Recap()
            Else
                Me.CetakStockBundling_Detail()
            End If
            Me.TabControl1.SelectedTab = TabPage2
            Me.TabPage2.Select()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function CetakStockBundling_Recap() As Boolean
        Dim ReportDetail As New CrpLaporan_Prod_List_Recap
        Dim DS As New DseLaporanProd_List
        Dim DTACetak_Header As New DseLaporanProd_ListTableAdapters.TblCetak_HeaderTableAdapter
        Dim TblCetak_Header As DseLaporanProd_List.TblCetak_HeaderDataTable = DS.TblCetak_Header
        Dim RowCetak_Header As DseLaporanProd_List.TblCetak_HeaderRow

        Dim DTACetak_Item_Before As New DseLaporanProd_ListTableAdapters.TblCetakItem_BeforeTableAdapter
        Dim TblCetak_Item_Before As DseLaporanProd_List.TblCetakItem_BeforeDataTable = DS.TblCetakItem_Before
        Dim RowCetak_Item_Before As DseLaporanProd_List.TblCetakItem_BeforeRow

        Dim DTACetak_Item_After As New DseLaporanProd_ListTableAdapters.TblCetakItem_AfterTableAdapter
        Dim TblCetak_Item_After As DseLaporanProd_List.TblCetakItem_AfterDataTable = DS.TblCetakItem_After
        Dim RowCetak_Item_After As DseLaporanProd_List.TblCetakItem_AfterRow

        Dim TBLHeader As DseLaporanProd_List.TBLHeaderDataTable = DS.TBLHeader
        Dim ROWHeader As DseLaporanProd_List.TBLHeaderRow

        Dim TBLItem_Before As DseLaporanProd_List.TBLItem_BeforeDataTable = DS.TBLItem_Before
        Dim TBLItem_Before_Dummy As New DseLaporanProd_List.TBLItem_BeforeDataTable
        Dim rowItem_Before As DseLaporanProd_List.TBLItem_BeforeRow
        Dim rowItem_Before_Dummy As DseLaporanProd_List.TBLItem_BeforeRow

        Dim TBLItem_After As DseLaporanProd_List.TBLItem_AfterDataTable = DS.TBLItem_After
        Dim TBLItem_After_Dummy As New DseLaporanProd_List.TBLItem_AfterDataTable
        Dim rowItem_After As DseLaporanProd_List.TBLItem_AfterRow
        Dim rowItem_After_Dummy As DseLaporanProd_List.TBLItem_AfterRow


        Dim paramStartDate As Date
        Dim paramEndDate As Date
        Dim ParamInput As String = ""
        Dim ParamOutput As String = ""
        Dim ParamFormat2 As String = ""
        Dim ParamSubTotalBefore As String = ""
        Dim ParamSubTotalAfter As String = ""
        Dim SubTotalBefore As String = ""
        Dim SubTotalAfter As String = ""

        Dim PrimaryMax As Double = -1
        Dim PrimaryMin As Double = -1
        Dim PrimaryTransaksi As Double = -1
        Dim index As Integer = 0
        Dim Filter1 As String = ""
        Dim Filter2 As String = ""
        Dim Filter3 As String = ""
        Dim DV1 As DataRowView
        Dim DV2 As DataRowView
        Dim row As New DataGridViewRow
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTACetak_Header.Connection.ConnectionString = Core_Module.GetConstring
            DTACetak_Item_After.Connection.ConnectionString = DTACetak_Header.Connection.ConnectionString
            DTACetak_Item_Before.Connection.ConnectionString = DTACetak_Header.Connection.ConnectionString
            TBLHeader.Rows.Clear()
            TBLItem_Before.Rows.Clear()
            TBLItem_After.Rows.Clear()
            paramStartDate = dtpFrom.Value.Date
            paramEndDate = dtpTo.Value.Date

            '1. Fill semua berdasarkan tanggal dari sampai
            DTACetak_Header.FillBy_StartEndDate1(TblCetak_Header, dtpFrom.Value, dtpTo.Value)
            TblCetak_Header.DefaultView.Sort = Nothing
            TblCetak_Header.DefaultView.Sort = "Primary"
            If TblCetak_Header.DefaultView.Count > 0 Then
                RowCetak_Header = TblCetak_Header.DefaultView(0).Row
                PrimaryMin = RowCetak_Header.Primary
                RowCetak_Header = TblCetak_Header.DefaultView(TblCetak_Header.DefaultView.Count - 1).Row
                PrimaryMax = RowCetak_Header.Primary
            End If
            '2. Fill semua item berdasarkan primaryTransaksi Max and Min
            DTACetak_Item_Before.FillBy_PrimaryMinAndMax(TblCetak_Item_Before, PrimaryMin, PrimaryMax)
            DTACetak_Item_After.FillBy_PrimaryMinAndMax(TblCetak_Item_After, PrimaryMin, PrimaryMax)
            '3. Filter dari header
            Filter1 = "[Primary] = [Primary]"

            '3.1 Filter Dgv Process
            If Me.chkProcess.Checked = True Then
                Filter1 &= Me.GetFilter(Me.dgvProcess, "primary process", Me.ColProcess_PrimaryOriginal)
            End If

            '3.2 Filter Dgv Work Unit
            If Me.chkWorkUnit.Checked = True Then
                Filter1 &= Me.GetFilter(Me.dgvWorkUnit, "primary work unit", Me.ColWorkUnit_PrimaryOriginal)
            End If
            '3.3 Filter Dgv Work Shift
            If Me.ChkWorkShift.Checked = True Then
                Filter1 &= Me.GetFilter(Me.DgvWorkShift, "primary workshift", Me.ColWorkShift_PrimaryOriginal)
            End If
            '4 tentukan input dan output yg akan ditampilkan
            If Me.Format2_Index = EnumFormat2.Input Then
                ParamInput = "BEFORE"
                ParamOutput = ""
            ElseIf Me.Format2_Index = EnumFormat2.Output Then
                ParamInput = ""
                ParamOutput = "AFTER"
            Else
                ParamInput = "BEFORE"
                ParamOutput = "AFTER"
            End If
            '5. isi header dan sorting sesuai yang dipilih
            TblCetak_Header.DefaultView.RowFilter = Filter1
            If TblCetak_Header.DefaultView.Count > 0 Then
                For Each DV1 In TblCetak_Header.DefaultView
                    RowCetak_Header = DV1.Row
                    ROWHeader = TBLHeader.NewTBLHeaderRow
                    With RowCetak_Header
                        ROWHeader.Primary = .Primary
                        ROWHeader.Tanggal_Transaksi = Format(.Tanggal_Transaksi, "dd MMM yyyy").ToString
                        ROWHeader.No_Transaksi = .No_Transaksi
                        ROWHeader.WorkUnit_Code = .work_unit_code
                        ROWHeader.WorkShift_Code = .Workshift_Code
                        ROWHeader.Process_Code = .Process_Code
                        'planned
                        ROWHeader.setup_time_planned = .setup_time_planned
                        ROWHeader.work_hour_planned = .work_hour_planned
                        ROWHeader.start_time_planned = .start_time_planned
                        ROWHeader.end_time_planned = .end_time_planned
                        ROWHeader.prod__index_planned = .prod__index_planned
                        'actual
                        ROWHeader.setup_time_actual = .setup_time_actual
                        ROWHeader.work_hour_actual = .work_hour_actual
                        ROWHeader.start_time_actual = .start_time_actual
                        ROWHeader.prod__index_actual = .prod__index_actual
                        ROWHeader.end_time_actual = .end_time_actual
                        ROWHeader.down_time_actual = .down_time_actual
                    End With
                    TBLHeader.AddTBLHeaderRow(ROWHeader)
                    '6. Filter item berdasarkan transaksi dengan headernya
                    Filter2 = "[Primary Transaksi] = " & RowCetak_Header.Primary.ToString
                    Filter3 = "[Primary Transaksi] = " & RowCetak_Header.Primary.ToString
                    '5. Filter item berdasarkan DGV product
                    'ini before
                    If Me.ChkProductGroup.Checked = True Then
                        Filter2 &= Me.GetFilter(Me.DgvGroup, Me.FilterProduct, Me.ColGroup_PrimaryOriginal)
                        Filter3 &= Me.GetFilter(Me.DgvGroup, Me.FilterProduct, Me.ColGroup_PrimaryOriginal)
                    End If

                    'isi item before
                    TblCetak_Item_Before.DefaultView.RowFilter = Nothing
                    TblCetak_Item_Before.DefaultView.RowFilter = Filter2
                    If TblCetak_Item_Before.DefaultView.Count > 0 Then
                        For Each DV2 In TblCetak_Item_Before.DefaultView
                            RowCetak_Item_Before = DV2.Row
                            rowItem_Before = TBLItem_Before.NewTBLItem_BeforeRow
                            With RowCetak_Item_Before
                                rowItem_Before.work_unit_code = ROWHeader.WorkUnit_Code
                                rowItem_Before.workshift_code = ROWHeader.WorkShift_Code
                                rowItem_Before.process_code = ROWHeader.Process_Code
                                rowItem_Before.Primary_Transaksi = .Primary_Transaksi
                                rowItem_Before.Primary = .Primary
                                rowItem_Before.Primary_Units = .Primary_Units
                                rowItem_Before.Primary_Product = .Primary_Product
                                rowItem_Before.Product_Code = .Product_Code
                                rowItem_Before.Description = .Description

                                If Me.rbo_both.Checked = True Then
                                    rowItem_Before.code_no_and_desc = .Product_Code & "-" & .Description
                                ElseIf Me.rbo_code_no_only.Checked = True Then
                                    rowItem_Before.code_no_and_desc = .Product_Code
                                ElseIf Me.rbo_desc_only.Checked = True Then
                                    rowItem_Before.code_no_and_desc = .Description
                                End If

                                rowItem_Before.Qty = Me.ObjCD.DigitGrouping(.Qty, RowCetak_Item_Before._Decimal)
                                rowItem_Before.Unit_Code = .Unit_Code
                                If .Location_Controlled = EnumPilih.No Then
                                    rowItem_Before.Location = "-"
                                Else
                                    rowItem_Before.Location = .Storage_Area_Code
                                End If
                                If .Batch_No_Controlled = EnumPilih.No Then
                                    rowItem_Before.Batch_No = "-"
                                Else
                                    RowCetak_Item_Before.Batch_No = .Batch_No
                                End If
                                If .Expiry_Date_Controlled = EnumPilih.No Then
                                    rowItem_Before.Expiry_Date = "-"
                                Else
                                    rowItem_Before.Expiry_Date = Format(.Expiry_Date, "dd MMM yyyy")
                                End If
                                rowItem_Before.Remarks = .Remarks

                                If Format_Index = EnumFormat.Recap_w_group Then
                                    If GroupBy_Index = EnumGroupBy.Process Then
                                        rowItem_Before.Group_By = ROWHeader.Process_Code
                                    ElseIf GroupBy_Index = EnumGroupBy.Work_Unit Then
                                        rowItem_Before.Group_By = ROWHeader.WorkUnit_Code
                                    ElseIf GroupBy_Index = EnumGroupBy.Work_Shift Then
                                        rowItem_Before.Group_By = ROWHeader.WorkShift_Code
                                    ElseIf GroupBy_Index = EnumGroupBy._Date Then
                                        rowItem_Before.Group_By = ROWHeader.Tanggal_Transaksi
                                    End If
                                End If
                            End With
                            TBLItem_Before.AddTBLItem_BeforeRow(rowItem_Before)
                        Next
                    End If

                    'isi item after
                    TblCetak_Item_After.DefaultView.RowFilter = Nothing
                    TblCetak_Item_After.DefaultView.RowFilter = Filter3
                    If TblCetak_Item_After.DefaultView.Count > 0 Then
                        For Each DV2 In TblCetak_Item_After.DefaultView
                            RowCetak_Item_After = DV2.Row
                            rowItem_After = TBLItem_After.NewTBLItem_AfterRow
                            With RowCetak_Item_After
                                rowItem_After.work_unit_code = ROWHeader.WorkUnit_Code
                                rowItem_After.workshift_code = ROWHeader.WorkShift_Code
                                rowItem_After.process_code = ROWHeader.Process_Code
                                rowItem_After.Primary_Transaksi = .Primary_Transaksi
                                rowItem_After.Primary = .Primary
                                rowItem_After.Primary_Product = .Primary_Product
                                rowItem_After.Primary_Units = .Primary_Units
                                rowItem_After.Product_Code = .Product_Code
                                rowItem_After.Description = .Description

                                If Me.rbo_both.Checked = True Then
                                    rowItem_After.code_no_and_desc = .Product_Code & "-" & .Description
                                ElseIf Me.rbo_code_no_only.Checked = True Then
                                    rowItem_After.code_no_and_desc = .Product_Code
                                ElseIf Me.rbo_desc_only.Checked = True Then
                                    rowItem_After.code_no_and_desc = .Description
                                End If

                                rowItem_After.Qty = Me.ObjCD.DigitGrouping(.Qty, RowCetak_Item_After._Decimal)
                                rowItem_After.Unit_Code = .Unit_Code

                                If .Location_Controlled = EnumPilih.No Then
                                    rowItem_After.Location = "-"
                                Else
                                    rowItem_After.Location = .Storage_Area_Code
                                End If
                                If .Batch_No_Controlled = EnumPilih.No Then
                                    rowItem_After.Batch_No = "-"
                                Else
                                    RowCetak_Item_After.Batch_No = .Batch_No
                                End If
                                If .Expiry_Date_Controlled = False Then
                                    rowItem_After.Expiry_Date = "-"
                                Else
                                    rowItem_After.Expiry_Date = Format(.Expiry_Date, "dd MMM yyyy")
                                End If

                                rowItem_After.Remarks = .Remarks
                                If Format_Index = EnumFormat.Recap_w_group Then
                                    If GroupBy_Index = EnumGroupBy.Process Then
                                        rowItem_After.Group_By = ROWHeader.Process_Code
                                    ElseIf GroupBy_Index = EnumGroupBy.Work_Unit Then
                                        rowItem_After.Group_By = ROWHeader.WorkUnit_Code
                                    ElseIf GroupBy_Index = EnumGroupBy.Work_Shift Then
                                        rowItem_After.Group_By = ROWHeader.WorkShift_Code
                                    ElseIf GroupBy_Index = EnumGroupBy._Date Then
                                        rowItem_After.Group_By = ROWHeader.Tanggal_Transaksi
                                    End If
                                End If
                            End With
                            TBLItem_After.AddTBLItem_AfterRow(rowItem_After)
                        Next
                    End If
                Next
            End If

            If Format_Index <> EnumFormat.Recap_w_group Then
                'sorting_not_group
                Me.GetSorting_Before(TBLItem_Before, TBLItem_Before_Dummy)
                Me.GetSorting_After(TBLItem_After, TBLItem_After_Dummy)
            Else
                'sorting_bygroup
                Me.GetSortingByGroup_Before(TBLItem_Before, TBLItem_Before_Dummy)
                Me.GetSortingByGroup_After(TBLItem_After, TBLItem_After_Dummy)
            End If

            'hitung subtotal
            Dim param_showsubtotal As New Boolean
            Dim tbl_subtotal As New DseLaporanProd_List.DataTable1DataTable
            If ChkShowSubTotal.Checked = False Then
                param_showsubtotal = False
            Else
                param_showsubtotal = True
            End If

            If Format_Index = EnumFormat.Recap_w_group Then
                Me.getSubtotalBefore1_perGroup(TBLItem_Before, tbl_subtotal)
                Me.getSubtotalAfter1_perGroup(TBLItem_After, tbl_subtotal)
            End If
            ParamSubTotalBefore = Me.getSubtotal1_Before(TBLItem_Before, tbl_subtotal, SubTotalBefore)
            ParamSubTotalAfter = Me.getSubtotal1_After(TBLItem_After, tbl_subtotal, SubTotalAfter)

            ReportDetail.SetDataSource(DS)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_StartDate.ParameterFieldName, paramStartDate)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_EndDate.ParameterFieldName, paramEndDate)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_Input.ParameterFieldName, ParamInput)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_SubReport1_Output.ParameterFieldName, ParamOutput)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_Sub_Total_Before.ParameterFieldName, ParamSubTotalBefore)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_SubReport1_Sub_Total_After.ParameterFieldName, ParamSubTotalAfter)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_show_subtotal_before.ParameterFieldName, param_showsubtotal)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_SubReport1_show_subtotal_after.ParameterFieldName, param_showsubtotal)
            Me.crv1.ReportSource = ReportDetail
            Me.crv1.Show()
            TabPage2.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTACetak_Header.Dispose()
            DTACetak_Item_After.Dispose()
            DTACetak_Item_Before.Dispose()
            TBLHeader.Dispose()
            TBLItem_After.Dispose()
            TBLItem_Before.Dispose()
        End Try
    End Function

    Public Function CetakStockBundling_Recap_with_Group() As Boolean
        Dim ReportDetail As New CrpLaporan_Prod_List_Recap_Group_with_Doc_No
        Dim DS As New DseLaporanProd_List
        Dim DTACetak_Header As New DseLaporanProd_ListTableAdapters.TblCetak_HeaderTableAdapter
        Dim TblCetak_Header As DseLaporanProd_List.TblCetak_HeaderDataTable = DS.TblCetak_Header
        Dim RowCetak_Header As DseLaporanProd_List.TblCetak_HeaderRow

        Dim DTACetak_Item_Before As New DseLaporanProd_ListTableAdapters.TblCetakItem_BeforeTableAdapter
        Dim TblCetak_Item_Before As DseLaporanProd_List.TblCetakItem_BeforeDataTable = DS.TblCetakItem_Before
        Dim RowCetak_Item_Before As DseLaporanProd_List.TblCetakItem_BeforeRow

        Dim DTACetak_Item_After As New DseLaporanProd_ListTableAdapters.TblCetakItem_AfterTableAdapter
        Dim TblCetak_Item_After As DseLaporanProd_List.TblCetakItem_AfterDataTable = DS.TblCetakItem_After
        Dim RowCetak_Item_After As DseLaporanProd_List.TblCetakItem_AfterRow

        Dim TBLHeader As DseLaporanProd_List.TBLHeaderDataTable = DS.TBLHeader
        Dim ROWHeader As DseLaporanProd_List.TBLHeaderRow

        Dim TBLItem1 As DseLaporanProd_List.TblItem1DataTable = DS.TblItem1
        Dim ROWItem1 As DseLaporanProd_List.TblItem1Row


        Dim paramStartDate As Date
        Dim paramEndDate As Date
        Dim ParamInput As String = ""
        Dim ParamOutput As String = ""
        Dim ParamFormat2 As String = ""
        Dim ParamSubtotal_Before As String = ""
        Dim ParamSubtotal_After As String = ""
        Dim SubTotalBefore As String = ""
        Dim SubTotalAfter As String = ""

        Dim PrimaryMax As Double = -1
        Dim PrimaryMin As Double = -1
        Dim PrimaryTransaksi As Double = -1
        Dim index As Integer = 0
        Dim Filter1 As String = ""
        Dim Filter2 As String = ""
        Dim Filter3 As String = ""
        Dim DV1 As DataRowView
        Dim DV2 As DataRowView
        Dim row As New DataGridViewRow
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTACetak_Header.Connection.ConnectionString = Core_Module.GetConstring
            DTACetak_Item_After.Connection.ConnectionString = DTACetak_Header.Connection.ConnectionString
            DTACetak_Item_Before.Connection.ConnectionString = DTACetak_Header.Connection.ConnectionString
            TBLHeader.Rows.Clear()
            TBLItem1.Rows.Clear()
            paramStartDate = dtpFrom.Value.Date
            paramEndDate = dtpTo.Value.Date

            '1. Fill semua berdasarkan tanggal dari sampai
            DTACetak_Header.FillBy_StartEndDate1(TblCetak_Header, dtpFrom.Value, dtpTo.Value)
            TblCetak_Header.DefaultView.Sort = Nothing
            TblCetak_Header.DefaultView.Sort = "Primary"
            If TblCetak_Header.DefaultView.Count > 0 Then
                RowCetak_Header = TblCetak_Header.DefaultView(0).Row
                PrimaryMin = RowCetak_Header.Primary
                RowCetak_Header = TblCetak_Header.DefaultView(TblCetak_Header.DefaultView.Count - 1).Row
                PrimaryMax = RowCetak_Header.Primary
            End If
            '2. Fill semua item berdasarkan primaryTransaksi Max and Min
            DTACetak_Item_Before.FillBy_PrimaryMinAndMax(TblCetak_Item_Before, PrimaryMin, PrimaryMax)
            DTACetak_Item_After.FillBy_PrimaryMinAndMax(TblCetak_Item_After, PrimaryMin, PrimaryMax)
            '3. Filter dari header
            Filter1 = "[Primary] = [Primary]"

            '3.1 Filter Dgv Process
            If Me.chkProcess.Checked = True Then
                Filter1 &= Me.GetFilter(Me.dgvProcess, "primary process", Me.ColProcess_PrimaryOriginal)
            End If

            '3.2 Filter Dgv Work Unit
            If Me.chkWorkUnit.Checked Then
                Filter1 &= Me.GetFilter(Me.dgvWorkUnit, "primary work unit", Me.ColWorkUnit_PrimaryOriginal)
            End If

            '3.3 Filter Dgv Work Shift
            If Me.ChkWorkShift.Checked = True Then
                Filter1 &= Me.GetFilter(Me.DgvWorkShift, "primary workshift", Me.ColWorkShift_PrimaryOriginal)
            End If

            '4 tentukan input dan output yg akan ditampilkan
            If Me.Format2_Index = EnumFormat2.Input Then
                ParamInput = "BEFORE"
                ParamOutput = ""
            ElseIf Me.Format2_Index = EnumFormat2.Output Then
                ParamInput = ""
                ParamOutput = "AFTER"
            Else
                ParamInput = "BEFORE"
                ParamOutput = "AFTER"
            End If
            '5. isi header dan sorting sesuai yang dipilih
            TblCetak_Header.DefaultView.Sort = Nothing
            TblCetak_Header.DefaultView.Sort = "Tanggal Transaksi, [No Transaksi]"
            TblCetak_Header.DefaultView.RowFilter = Nothing
            TblCetak_Header.DefaultView.RowFilter = Filter1
            If TblCetak_Header.DefaultView.Count > 0 Then
                For Each DV1 In TblCetak_Header.DefaultView
                    RowCetak_Header = DV1.Row
                    ROWHeader = TBLHeader.NewTBLHeaderRow
                    With RowCetak_Header
                        ROWHeader.Primary = .Primary
                        ROWHeader.Tanggal_Transaksi = Format(.Tanggal_Transaksi, "dd MMM yyyy").ToString
                        ROWHeader.No_Transaksi = .No_Transaksi
                        ROWHeader.WorkUnit_Code = .work_unit_code
                        ROWHeader.WorkShift_Code = .Workshift_Code
                        ROWHeader.Process_Code = .Process_Code
                        'planned
                        ROWHeader.setup_time_planned = .setup_time_planned
                        ROWHeader.work_hour_planned = .work_hour_planned
                        ROWHeader.start_time_planned = .start_time_planned
                        ROWHeader.end_time_planned = .end_time_planned
                        ROWHeader.prod__index_planned = .prod__index_planned
                        'actual
                        ROWHeader.setup_time_actual = .setup_time_actual
                        ROWHeader.work_hour_actual = .work_hour_actual
                        ROWHeader.start_time_actual = .start_time_actual
                        ROWHeader.prod__index_actual = .prod__index_actual
                        ROWHeader.end_time_actual = .end_time_actual
                        ROWHeader.down_time_actual = .down_time_actual
                    End With
                    TBLHeader.AddTBLHeaderRow(ROWHeader)
                    '6. Filter item berdasarkan transaksi dengan headernya
                    Filter2 = "[Primary Transaksi] = " & RowCetak_Header.Primary.ToString
                    Filter3 = "[Primary Transaksi] = " & RowCetak_Header.Primary.ToString
                    '5. Filter item berdasarkan DGV product
                    'ini before
                    If Me.ChkProductGroup.Checked = True Then
                        Filter2 &= Me.GetFilter(Me.DgvGroup, Me.FilterProduct, Me.ColGroup_PrimaryOriginal)
                        Filter3 &= Me.GetFilter(Me.DgvGroup, Me.FilterProduct, Me.ColGroup_PrimaryOriginal)
                    End If

                    'tentukan mana yg mau ditampilin
                    If Me.Format2_Index = EnumFormat2.Input Then
                        '1.kalo input saja
                        'isi item before
                        TblCetak_Item_Before.DefaultView.RowFilter = Nothing
                        TblCetak_Item_Before.DefaultView.RowFilter = Filter2
                        If TblCetak_Item_Before.DefaultView.Count > 0 Then
                            For Each DV2 In TblCetak_Item_Before.DefaultView
                                RowCetak_Item_Before = DV2.Row
                                ROWItem1 = TBLItem1.NewTblItem1Row
                                With RowCetak_Item_Before
                                    ROWItem1.Primary_Transaksi = .Primary_Transaksi
                                    ROWItem1.Primary = .Primary
                                    ROWItem1.Primary_Product = .Primary_Product
                                    ROWItem1.Primary_Units = .Primary_Units
                                    ROWItem1.Product_Code = .Product_Code
                                    ROWItem1.Description = .Description

                                    If Me.rbo_both.Checked = True Then
                                        ROWItem1.code_no_and_desc = .Product_Code & "-" & .Description
                                    ElseIf Me.rbo_code_no_only.Checked = True Then
                                        ROWItem1.code_no_and_desc = .Product_Code
                                    ElseIf Me.rbo_desc_only.Checked = True Then
                                        ROWItem1.code_no_and_desc = .Description
                                    End If

                                    ROWItem1.Qty = Me.ObjCD.DigitGrouping(.Qty, RowCetak_Item_Before._Decimal)
                                    ROWItem1.QtyDec = Me.ObjCD.DigitGrouping(.Qty, RowCetak_Item_Before._Decimal)
                                    ROWItem1.Unit_Code = .Unit_Code
                                    If .Location_Controlled = EnumPilih.No Then
                                        ROWItem1.Location = "-"
                                    Else
                                        ROWItem1.Location = .Storage_Area_Code
                                    End If
                                    If .Batch_No_Controlled = EnumPilih.No Then
                                        ROWItem1.Batch_No = "-"
                                    Else
                                        ROWItem1.Batch_No = .Batch_No
                                    End If
                                    If .Expiry_Date_Controlled = EnumPilih.No Then
                                        ROWItem1.Expiry_Date = "-"
                                    Else
                                        ROWItem1.Expiry_Date = Format(.Expiry_Date, "dd MMM yyyy")
                                    End If
                                    ROWItem1.Faktor_Pemberat = "-"
                                    ROWItem1.Remarks = .Remarks
                                    ROWItem1.Group_Before_After = "BEFORE"
                                    If Format_Index = EnumFormat.Recap_w_group Then
                                        If GroupBy_Index = EnumGroupBy.Document_No Then
                                            ROWItem1.Group_By = ROWHeader.No_Transaksi
                                        End If
                                    End If
                                End With
                                TBLItem1.AddTblItem1Row(ROWItem1)
                            Next
                        End If
                    ElseIf Me.Format2_Index = EnumFormat2.Output Then
                        '2. kalo output saja
                        'isi item after
                        TblCetak_Item_After.DefaultView.RowFilter = Nothing
                        TblCetak_Item_After.DefaultView.RowFilter = Filter3
                        If TblCetak_Item_After.DefaultView.Count > 0 Then
                            For Each DV2 In TblCetak_Item_After.DefaultView
                                RowCetak_Item_After = DV2.Row
                                ROWItem1 = TBLItem1.NewTblItem1Row
                                With RowCetak_Item_After
                                    ROWItem1.Primary_Transaksi = .Primary_Transaksi
                                    ROWItem1.Primary = .Primary
                                    ROWItem1.Primary_Product = .Primary_Product
                                    ROWItem1.Primary_Units = .Primary_Units
                                    ROWItem1.Product_Code = .Product_Code
                                    ROWItem1.Description = .Description

                                    If Me.rbo_both.Checked = True Then
                                        ROWItem1.code_no_and_desc = .Product_Code & "-" & .Description
                                    ElseIf Me.rbo_code_no_only.Checked = True Then
                                        ROWItem1.code_no_and_desc = .Product_Code
                                    ElseIf Me.rbo_desc_only.Checked = True Then
                                        ROWItem1.code_no_and_desc = .Description
                                    End If

                                    ROWItem1.Qty = Me.ObjCD.DigitGrouping(.Qty, RowCetak_Item_After._Decimal)
                                    ROWItem1.QtyDec = Me.ObjCD.DigitGrouping(.Qty, RowCetak_Item_After._Decimal)
                                    ROWItem1.Unit_Code = .Unit_Code

                                    If .Location_Controlled = EnumPilih.No Then
                                        ROWItem1.Location = "-"
                                    Else
                                        ROWItem1.Location = .Storage_Area_Code
                                    End If
                                    If .Batch_No_Controlled = EnumPilih.No Then
                                        ROWItem1.Batch_No = "-"
                                    Else
                                        ROWItem1.Batch_No = .Batch_No
                                    End If
                                    If .Expiry_Date_Controlled = False Then
                                        ROWItem1.Expiry_Date = "-"
                                    Else
                                        ROWItem1.Expiry_Date = Format(.Expiry_Date, "dd MMM yyyy")
                                    End If

                                    ROWItem1.Faktor_Pemberat = .Faktor_Pemberat
                                    ROWItem1.Remarks = .Remarks
                                    ROWItem1.Group_Before_After = "AFTER"
                                    If Format_Index = EnumFormat.Recap_w_group Then
                                        If GroupBy_Index = EnumGroupBy.Document_No Then
                                            ROWItem1.Group_By = RowCetak_Header.No_Transaksi
                                        End If
                                    End If
                                End With
                                TBLItem1.AddTblItem1Row(ROWItem1)
                            Next
                        End If
                    Else

                        'kalo keduanya
                        'isi item before
                        TblCetak_Item_Before.DefaultView.RowFilter = Nothing
                        TblCetak_Item_Before.DefaultView.RowFilter = Filter2
                        If TblCetak_Item_Before.DefaultView.Count > 0 Then
                            For Each DV2 In TblCetak_Item_Before.DefaultView
                                RowCetak_Item_Before = DV2.Row
                                ROWItem1 = TBLItem1.NewTblItem1Row
                                With RowCetak_Item_Before
                                    ROWItem1.Primary_Transaksi = .Primary_Transaksi
                                    ROWItem1.Primary = .Primary
                                    ROWItem1.Primary_Product = .Primary_Product
                                    ROWItem1.Primary_Units = .Primary_Units
                                    ROWItem1.Product_Code = .Product_Code
                                    ROWItem1.Description = .Description

                                    If Me.rbo_both.Checked = True Then
                                        ROWItem1.code_no_and_desc = .Product_Code & "-" & .Description
                                    ElseIf Me.rbo_code_no_only.Checked = True Then
                                        ROWItem1.code_no_and_desc = .Product_Code
                                    ElseIf Me.rbo_desc_only.Checked = True Then
                                        ROWItem1.code_no_and_desc = .Description
                                    End If

                                    ROWItem1.Qty = Me.ObjCD.DigitGrouping(.Qty, RowCetak_Item_Before._Decimal)
                                    ROWItem1.QtyDec = Me.ObjCD.DigitGrouping(.Qty, RowCetak_Item_Before._Decimal)
                                    ROWItem1.Unit_Code = .Unit_Code

                                    If .Location_Controlled = EnumPilih.No Then
                                        ROWItem1.Location = "-"
                                    Else
                                        ROWItem1.Location = .Storage_Area_Code
                                    End If
                                    If .Batch_No_Controlled = EnumPilih.No Then
                                        ROWItem1.Batch_No = "-"
                                    Else
                                        ROWItem1.Batch_No = .Batch_No
                                    End If
                                    If .Expiry_Date_Controlled = False Then
                                        ROWItem1.Expiry_Date = "-"
                                    Else
                                        ROWItem1.Expiry_Date = Format(.Expiry_Date, "dd MMM yyyy")
                                    End If

                                    ROWItem1.Faktor_Pemberat = "-"
                                    ROWItem1.Remarks = .Remarks
                                    ROWItem1.Group_Before_After = "BEFORE"
                                    If Format_Index = EnumFormat.Recap_w_group Then
                                        If GroupBy_Index = EnumGroupBy.Document_No Then
                                            ROWItem1.Group_By = ROWHeader.No_Transaksi
                                        End If
                                    End If
                                End With
                                TBLItem1.AddTblItem1Row(ROWItem1)
                            Next
                        End If

                        'isi item after
                        TblCetak_Item_After.DefaultView.RowFilter = Nothing
                        TblCetak_Item_After.DefaultView.RowFilter = Filter3
                        If TblCetak_Item_After.DefaultView.Count > 0 Then
                            For Each DV2 In TblCetak_Item_After.DefaultView
                                RowCetak_Item_After = DV2.Row
                                ROWItem1 = TBLItem1.NewTblItem1Row
                                With RowCetak_Item_After
                                    ROWItem1.Primary_Transaksi = .Primary_Transaksi
                                    ROWItem1.Primary = .Primary
                                    ROWItem1.Primary_Product = .Primary_Product
                                    ROWItem1.Primary_Units = .Primary_Units
                                    ROWItem1.Product_Code = .Product_Code
                                    ROWItem1.Description = .Description

                                    If Me.rbo_both.Checked = True Then
                                        ROWItem1.code_no_and_desc = .Product_Code & "-" & .Description
                                    ElseIf Me.rbo_code_no_only.Checked = True Then
                                        ROWItem1.code_no_and_desc = .Product_Code
                                    ElseIf Me.rbo_desc_only.Checked = True Then
                                        ROWItem1.code_no_and_desc = .Description
                                    End If

                                    ROWItem1.Qty = Me.ObjCD.DigitGrouping(.Qty, RowCetak_Item_After._Decimal)
                                    ROWItem1.QtyDec = Me.ObjCD.DigitGrouping(.Qty, RowCetak_Item_After._Decimal)
                                    ROWItem1.Unit_Code = .Unit_Code

                                    If .Location_Controlled = EnumPilih.No Then
                                        ROWItem1.Location = "-"
                                    Else
                                        ROWItem1.Location = .Storage_Area_Code
                                    End If
                                    If .Batch_No_Controlled = EnumPilih.No Then
                                        ROWItem1.Batch_No = "-"
                                    Else
                                        ROWItem1.Batch_No = .Batch_No
                                    End If
                                    If .Expiry_Date_Controlled = False Then
                                        ROWItem1.Expiry_Date = "-"
                                    Else
                                        ROWItem1.Expiry_Date = Format(.Expiry_Date, "dd MMM yyyy")
                                    End If

                                    ROWItem1.Faktor_Pemberat = .Faktor_Pemberat
                                    ROWItem1.Remarks = .Remarks
                                    ROWItem1.Group_Before_After = "AFTER"

                                    If Format_Index = EnumFormat.Recap_w_group Then
                                        If GroupBy_Index = EnumGroupBy.Document_No Then
                                            ROWItem1.Group_By = RowCetak_Header.No_Transaksi
                                        End If
                                    End If
                                End With
                                TBLItem1.AddTblItem1Row(ROWItem1)
                            Next
                        End If
                    End If
                Next
            End If

            'hitung subtotal
            Dim tblsubtotal As New DseLaporanProd_List.DataTable1DataTable
            Dim param_showsubtotal As New Boolean
            If Me.ChkShowSubTotal.Checked = False Then
                param_showsubtotal = False
            Else
                param_showsubtotal = True
                ParamSubtotal_Before = Me.getSubtotal2_Before(TBLItem1, tblsubtotal, SubTotalBefore)
                ParamSubtotal_After = Me.getSubtotal2_After(TBLItem1, tblsubtotal, SubTotalAfter)
            End If

            ReportDetail.SetDataSource(DS)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_StartDate.ParameterFieldName, paramStartDate)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_EndDate.ParameterFieldName, paramEndDate)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_Input.ParameterFieldName, ParamInput)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_output.ParameterFieldName, ParamOutput)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_SubtotalBefore.ParameterFieldName, ParamSubtotal_Before)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_SubtotalAfter.ParameterFieldName, ParamSubtotal_After)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_show_subtotal.ParameterFieldName, param_showsubtotal)
            Me.crv1.ReportSource = ReportDetail
            Me.crv1.Show()
            TabPage2.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTACetak_Header.Dispose()
            DTACetak_Item_After.Dispose()
            DTACetak_Item_Before.Dispose()
            TBLHeader.Dispose()
        End Try
    End Function

    Public Function CetakStockBundling_Detail() As Boolean
        Dim ReportDetail As New CrpLaporan_Prod_List_Detail
        Dim DS As New DseLaporanProd_List
        Dim DTACetak_Header As New DseLaporanProd_ListTableAdapters.TblCetak_HeaderTableAdapter
        Dim TblCetak_Header As DseLaporanProd_List.TblCetak_HeaderDataTable = DS.TblCetak_Header
        Dim RowCetak_Header As DseLaporanProd_List.TblCetak_HeaderRow

        Dim DTACetak_Item_Before As New DseLaporanProd_ListTableAdapters.TblCetakItem_BeforeTableAdapter
        Dim TblCetak_Item_Before As DseLaporanProd_List.TblCetakItem_BeforeDataTable = DS.TblCetakItem_Before
        Dim RowCetak_Item_Before As DseLaporanProd_List.TblCetakItem_BeforeRow

        Dim DTACetak_Item_After As New DseLaporanProd_ListTableAdapters.TblCetakItem_AfterTableAdapter
        Dim TblCetak_Item_After As DseLaporanProd_List.TblCetakItem_AfterDataTable = DS.TblCetakItem_After
        Dim RowCetak_Item_After As DseLaporanProd_List.TblCetakItem_AfterRow

        Dim TBLHeader As DseLaporanProd_List.TBLHeaderDataTable = DS.TBLHeader
        Dim ROWHeader As DseLaporanProd_List.TBLHeaderRow

        Dim TBLItem_Before As DseLaporanProd_List.TBLItem_BeforeDataTable = DS.TBLItem_Before
        Dim rowItem_Before As DseLaporanProd_List.TBLItem_BeforeRow

        Dim TBLItem_After As DseLaporanProd_List.TBLItem_AfterDataTable = DS.TBLItem_After
        Dim rowItem_After As DseLaporanProd_List.TBLItem_AfterRow

        Dim TBLItem_Before_Dummy As New DseLaporanProd_List.TBLItem_BeforeDataTable
        Dim rowItem_Before_Dummy As DseLaporanProd_List.TBLItem_BeforeRow

        Dim TBLItem_After_Dummy As New DseLaporanProd_List.TBLItem_AfterDataTable
        Dim rowItem_After_Dummy As DseLaporanProd_List.TBLItem_AfterRow



        Dim paramStartDate As Date
        Dim paramEndDate As Date
        Dim ParamInput As String = ""
        Dim ParamOutput As String = ""
        Dim ParamFormat2 As String = ""
        Dim ParamSubTotalBefore As String = ""
        Dim ParamSubTotalAfter As String = ""
        Dim SubTotalBefore As String = ""
        Dim SubTotalAfter As String = ""
        Dim PrimaryMax As Double = -1
        Dim PrimaryMin As Double = -1
        Dim PrimaryTransaksi As Double = -1
        Dim index As Integer = 0
        Dim Filter1 As String = ""
        Dim Filter2 As String = ""
        Dim Filter3 As String = ""
        Dim DV1 As DataRowView
        Dim DV2 As DataRowView
        Dim row As New DataGridViewRow
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTACetak_Header.Connection.ConnectionString = Core_Module.GetConstring
            DTACetak_Item_After.Connection.ConnectionString = DTACetak_Header.Connection.ConnectionString
            DTACetak_Item_Before.Connection.ConnectionString = DTACetak_Header.Connection.ConnectionString
            TBLHeader.Rows.Clear()
            TBLItem_Before.Rows.Clear()
            TBLItem_After.Rows.Clear()
            paramStartDate = dtpFrom.Value.Date
            paramEndDate = dtpTo.Value.Date

            '1. Fill semua berdasarkan tanggal dari sampai
            DTACetak_Header.FillBy_StartEndDate1(TblCetak_Header, dtpFrom.Value, dtpTo.Value)
            TblCetak_Header.DefaultView.Sort = Nothing
            TblCetak_Header.DefaultView.Sort = "Primary"
            If TblCetak_Header.DefaultView.Count > 0 Then
                RowCetak_Header = TblCetak_Header.DefaultView(0).Row
                PrimaryMin = RowCetak_Header.Primary
                RowCetak_Header = TblCetak_Header.DefaultView(TblCetak_Header.DefaultView.Count - 1).Row
                PrimaryMax = RowCetak_Header.Primary
            End If
            '2. Fill semua item berdasarkan primaryTransaksi Max and Min
            DTACetak_Item_Before.FillBy_PrimaryMinAndMax(TblCetak_Item_Before, PrimaryMin, PrimaryMax)
            DTACetak_Item_After.FillBy_PrimaryMinAndMax(TblCetak_Item_After, PrimaryMin, PrimaryMax)
            '3. Filter dari header
            Filter1 = "[Primary] = [Primary]"

            '3.1 Filter Dgv Process
            If Me.chkProcess.Checked = True Then
                Filter1 &= Me.GetFilter(Me.dgvProcess, "primary process", Me.ColProcess_PrimaryOriginal)
            End If

            '3.2 Filter Dgv Work Unit
            If Me.chkWorkUnit.Checked = True Then
                Filter1 &= Me.GetFilter(Me.dgvWorkUnit, "primary work unit", Me.ColWorkUnit_PrimaryOriginal)
            End If

            '3.3 Filter Dgv Work Shift
            If Me.ChkWorkShift.Checked = True Then
                Filter1 &= Me.GetFilter(Me.DgvWorkShift, "primary workshift", Me.ColWorkShift_PrimaryOriginal)
            End If

            '4 tentukan input dan output yg akan ditampilkan
            If Me.Format2_Index = EnumFormat2.Input Then
                ParamInput = "BEFORE"
                ParamOutput = ""
            ElseIf Me.Format2_Index = EnumFormat2.Output Then
                ParamInput = ""
                ParamOutput = "AFTER"
            ElseIf Me.Format2_Index = EnumFormat2.Both Then
                ParamInput = "BEFORE"
                ParamOutput = "AFTER"
            End If
            '5. isi header dan sorting sesuai yang dipilih
            TblCetak_Header.DefaultView.RowFilter = Nothing
            TblCetak_Header.DefaultView.RowFilter = Filter1
            If TblCetak_Header.DefaultView.Count > 0 Then
                For Each DV1 In TblCetak_Header.DefaultView
                    RowCetak_Header = DV1.Row
                    ROWHeader = TBLHeader.NewTBLHeaderRow
                    With RowCetak_Header
                        ROWHeader.Primary = .Primary
                        ROWHeader.Tanggal_Transaksi = Format(.Tanggal_Transaksi, "dd MMM yyyy").ToString
                        ROWHeader.No_Transaksi = .No_Transaksi
                        ROWHeader.WorkUnit_Code = .work_unit_code
                        ROWHeader.WorkShift_Code = .Workshift_Code
                        ROWHeader.Process_Code = .Process_Code
                        'planned
                        ROWHeader.setup_time_planned = .setup_time_planned
                        ROWHeader.work_hour_planned = .work_hour_planned
                        ROWHeader.start_time_planned = .start_time_planned
                        ROWHeader.end_time_planned = .end_time_planned
                        ROWHeader.prod__index_planned = .prod__index_planned
                        'actual
                        ROWHeader.setup_time_actual = .setup_time_actual
                        ROWHeader.work_hour_actual = .work_hour_actual
                        ROWHeader.start_time_actual = .start_time_actual
                        ROWHeader.prod__index_actual = .prod__index_actual
                        ROWHeader.end_time_actual = .end_time_actual
                        ROWHeader.down_time_actual = .down_time_actual
                    End With
                    TBLHeader.AddTBLHeaderRow(ROWHeader)
                    '6. Filter item berdasarkan transaksi dengan headernya
                    Filter2 = "[Primary Transaksi] = " & RowCetak_Header.Primary.ToString
                    Filter3 = "[Primary Transaksi] = " & RowCetak_Header.Primary.ToString
                    '5. Filter item berdasarkan DGV product
                    If Me.ChkProductGroup.Checked Then
                        Filter2 &= Me.GetFilter(Me.DgvGroup, Me.FilterProduct, Me.ColGroup_PrimaryOriginal)
                        Filter3 &= Me.GetFilter(Me.DgvGroup, Me.FilterProduct, Me.ColGroup_PrimaryOriginal)
                    End If
                    'isi item before
                    TblCetak_Item_Before.DefaultView.RowFilter = Nothing
                    TblCetak_Item_Before.DefaultView.RowFilter = Filter2
                    If TblCetak_Item_Before.DefaultView.Count > 0 Then
                        For Each DV2 In TblCetak_Item_Before.DefaultView
                            RowCetak_Item_Before = DV2.Row
                            rowItem_Before = TBLItem_Before.NewTBLItem_BeforeRow
                            With RowCetak_Item_Before
                                rowItem_Before._date = ROWHeader.Tanggal_Transaksi
                                rowItem_Before.no_transaksi = ROWHeader.No_Transaksi
                                rowItem_Before.work_unit_code = ROWHeader.WorkUnit_Code
                                rowItem_Before.workshift_code = ROWHeader.WorkShift_Code
                                rowItem_Before.process_code = ROWHeader.Process_Code
                                rowItem_Before.Primary_Transaksi = .Primary_Transaksi
                                rowItem_Before.Primary = .Primary
                                rowItem_Before.Primary_Product = .Primary_Product
                                rowItem_Before.Primary_Units = .Primary_Units
                                rowItem_Before.Product_Code = .Product_Code
                                rowItem_Before.Description = .Description

                                If Me.rbo_both.Checked = True Then
                                    rowItem_Before.code_no_and_desc = .Product_Code & "-" & .Description
                                ElseIf Me.rbo_code_no_only.Checked = True Then
                                    rowItem_Before.code_no_and_desc = .Product_Code
                                ElseIf Me.rbo_desc_only.Checked = True Then
                                    rowItem_Before.code_no_and_desc = .Description
                                End If

                                rowItem_Before.Qty = Me.ObjCD.DigitGrouping(.Qty, RowCetak_Item_Before._Decimal)
                                rowItem_Before.Unit_Code = .Unit_Code

                                If .Location_Controlled = EnumPilih.No Then
                                    rowItem_Before.Location = "-"
                                Else
                                    rowItem_Before.Location = .Storage_Area_Code
                                End If
                                If .Batch_No_Controlled = EnumPilih.No Then
                                    rowItem_Before.Batch_No = "-"
                                Else
                                    RowCetak_Item_Before.Batch_No = .Batch_No
                                End If
                                If .Expiry_Date_Controlled = EnumPilih.No Then
                                    rowItem_Before.Expiry_Date = "-"
                                Else
                                    rowItem_Before.Expiry_Date = Format(.Expiry_Date, "dd MMM yyyy")
                                End If

                                rowItem_Before.Remarks = .Remarks
                                If Format_Index = EnumFormat.Recap_w_group Then
                                    If GroupBy_Index = EnumGroupBy.Process Then
                                        rowItem_Before.Group_By = ROWHeader.Process_Code
                                    ElseIf GroupBy_Index = EnumGroupBy.Work_Unit Then
                                        rowItem_Before.Group_By = ROWHeader.WorkUnit_Code
                                    ElseIf GroupBy_Index = EnumGroupBy.Work_Shift Then
                                        rowItem_Before.Group_By = ROWHeader.WorkShift_Code
                                    ElseIf GroupBy_Index = EnumGroupBy._Date Then
                                        rowItem_Before.Group_By = ROWHeader.Tanggal_Transaksi
                                    End If
                                End If
                            End With
                            TBLItem_Before.AddTBLItem_BeforeRow(rowItem_Before)
                        Next
                    End If

                    'isi item after
                    TblCetak_Item_After.DefaultView.RowFilter = Nothing
                    TblCetak_Item_After.DefaultView.RowFilter = Filter3
                    If TblCetak_Item_After.DefaultView.Count > 0 Then
                        For Each DV2 In TblCetak_Item_After.DefaultView
                            RowCetak_Item_After = DV2.Row
                            rowItem_After = TBLItem_After.NewTBLItem_AfterRow
                            With RowCetak_Item_After
                                rowItem_After._date = ROWHeader.Tanggal_Transaksi
                                rowItem_After.no_transaksi = ROWHeader.No_Transaksi
                                rowItem_After.work_unit_code = ROWHeader.WorkUnit_Code
                                rowItem_After.workshift_code = ROWHeader.WorkShift_Code
                                rowItem_After.process_code = ROWHeader.Process_Code
                                rowItem_After.Primary_Transaksi = .Primary_Transaksi
                                rowItem_After.Primary = .Primary
                                rowItem_After.Primary_Product = .Primary_Product
                                rowItem_After.Primary_Units = .Primary_Units
                                rowItem_After.Product_Code = .Product_Code
                                rowItem_After.Description = .Description

                                If Me.rbo_both.Checked = True Then
                                    rowItem_After.code_no_and_desc = .Product_Code & "-" & .Description
                                ElseIf Me.rbo_code_no_only.Checked = True Then
                                    rowItem_After.code_no_and_desc = .Product_Code
                                ElseIf Me.rbo_desc_only.Checked = True Then
                                    rowItem_After.code_no_and_desc = .Description
                                End If

                                rowItem_After.Qty = Me.ObjCD.DigitGrouping(.Qty, RowCetak_Item_After._Decimal)
                                rowItem_After.Unit_Code = .Unit_Code

                                If .Location_Controlled = EnumPilih.No Then
                                    rowItem_After.Location = "-"
                                Else
                                    rowItem_After.Location = .Storage_Area_Code
                                End If
                                If .Batch_No_Controlled = EnumPilih.No Then
                                    rowItem_After.Batch_No = "-"
                                Else
                                    RowCetak_Item_After.Batch_No = .Batch_No
                                End If
                                If .Expiry_Date_Controlled = False Then
                                    rowItem_After.Expiry_Date = "-"
                                Else
                                    rowItem_After.Expiry_Date = Format(.Expiry_Date, "dd MMM yyyy")
                                End If

                                rowItem_After.Remarks = .Remarks
                                If Format_Index = EnumFormat.Recap_w_group Then
                                    If GroupBy_Index = EnumGroupBy.Process Then
                                        rowItem_After.Group_By = ROWHeader.Process_Code
                                    ElseIf GroupBy_Index = EnumGroupBy.Work_Unit Then
                                        rowItem_After.Group_By = ROWHeader.WorkUnit_Code
                                    ElseIf GroupBy_Index = EnumGroupBy.Work_Shift Then
                                        rowItem_After.Group_By = ROWHeader.WorkUnit_Code
                                    ElseIf GroupBy_Index = EnumGroupBy._Date Then
                                        rowItem_After.Group_By = ROWHeader.Tanggal_Transaksi
                                    End If
                                End If

                            End With
                            TBLItem_After.AddTBLItem_AfterRow(rowItem_After)
                        Next
                    End If
                Next
            End If

            'sorting not group
            Me.GetSorting_Before(TBLItem_Before, TBLItem_Before_Dummy)
            Me.GetSorting_After(TBLItem_After, TBLItem_After_Dummy)

            'hitung sub total
            Dim param_showsubtotal As New Boolean
            Dim tbl_subtotal As New DseLaporanProd_List.DataTable1DataTable
            If ChkShowSubTotal.Checked = False Then
                param_showsubtotal = False
            Else
                param_showsubtotal = True
                ParamSubTotalBefore = Me.getSubtotal1_Before(TBLItem_Before, tbl_subtotal, SubTotalBefore)
                ParamSubTotalAfter = Me.getSubtotal1_After(TBLItem_After, tbl_subtotal, SubTotalAfter)
            End If
            ''Dim idx As Integer = 0
            ''For Each rowItem_Before In TBLItem_Before
            ''    rowItem_Before.Index_ = TBLItem_Before.Rows.IndexOf(rowItem_Before) + 1
            ''    idx = rowItem_Before.Index_
            ''Next

            ReportDetail.SetDataSource(DS)
            ''ReportDetail.SetParameterValue(ReportDetail.Parameter_LastIndex_.ParameterFieldName, idx)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_StartDate.ParameterFieldName, paramStartDate)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_EndDate.ParameterFieldName, paramEndDate)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_Input.ParameterFieldName, ParamInput)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_SubReport1_Output.ParameterFieldName, ParamOutput)

            ReportDetail.SetParameterValue(ReportDetail.Parameter_Sub_Total_Before.ParameterFieldName, ParamSubTotalBefore)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_SubReport1_Sub_Total_After.ParameterFieldName, ParamSubTotalAfter)

            ReportDetail.SetParameterValue(ReportDetail.Parameter_show_subtotal.ParameterFieldName, param_showsubtotal)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_SubReport1_show_subtotal_after.ParameterFieldName, param_showsubtotal)
            Me.crv1.ReportSource = ReportDetail
            Me.crv1.Show()
            TabPage2.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTACetak_Header.Dispose()
            DTACetak_Item_After.Dispose()
            DTACetak_Item_Before.Dispose()
            TBLHeader.Dispose()
            TBLItem_After.Dispose()
            TBLItem_Before.Dispose()
        End Try
    End Function

    Public Function getSubtotal1_Before(ByVal TBLItem_Before As DseLaporanProd_List.TBLItem_BeforeDataTable, _
                                    ByVal tblSubTotal As DseLaporanProd_List.DataTable1DataTable, ByVal subTotalBefore As String) As String
        Dim rowItem_Before As DseLaporanProd_List.TBLItem_BeforeRow
        Dim rowSubTotal As DseLaporanProd_List.DataTable1Row
        Try
            For Each rowItem_Before In TBLItem_Before
                With rowItem_Before
                    tblSubTotal.DefaultView.RowFilter = "[Primary Unit] = " & .Primary_Units.ToString & "AND [Before or After] = 'before'"
                    If tblSubTotal.DefaultView.Count > 0 Then
                        rowSubTotal = tblSubTotal.DefaultView(0).Row
                        rowSubTotal.Qty += .Qty
                        If rowSubTotal.Max_DefPrec < ._Decimal Then
                            rowSubTotal.Max_DefPrec = ._Decimal
                        End If
                    Else
                        rowSubTotal = tblSubTotal.NewDataTable1Row
                        rowSubTotal.Primary_Unit = .Primary_Units
                        rowSubTotal.Unit_Code = .Unit_Code
                        rowSubTotal.Before_or_After = "before"
                        rowSubTotal.Qty = .Qty
                        rowSubTotal.Group = ""
                        rowSubTotal.Max_DefPrec = ._Decimal
                        tblSubTotal.Rows.Add(rowSubTotal)
                    End If
                End With
            Next
            tblSubTotal.DefaultView.RowFilter = "[Before or After] = 'before'"
            For i As Integer = 0 To tblSubTotal.DefaultView.Count - 1
                rowSubTotal = tblSubTotal.DefaultView(i).Row
                With rowSubTotal
                    If i < tblSubTotal.DefaultView.Count - 1 Then
                        subTotalBefore += ObjCD.DigitGrouping(.Qty, .Max_DefPrec) & " " & .Unit_Code & vbNewLine
                    Else
                        subTotalBefore += ObjCD.DigitGrouping(.Qty, .Max_DefPrec) & " " & .Unit_Code
                    End If
                End With
            Next
            Return subTotalBefore
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function getSubtotal1_After(ByVal TBLItem_After As DseLaporanProd_List.TBLItem_AfterDataTable, _
                               ByVal tblSubTotal As DseLaporanProd_List.DataTable1DataTable, ByVal subTotalAfter As String) As String
        Dim rowItem_After As DseLaporanProd_List.TBLItem_AfterRow
        Dim rowSubTotal As DseLaporanProd_List.DataTable1Row
        Try
            For Each rowItem_After In TBLItem_After
                With rowItem_After
                    tblSubTotal.DefaultView.RowFilter = "[Primary Unit] = " & .Primary_Units.ToString & " AND [Before or After] = 'after'"
                    If tblSubTotal.DefaultView.Count > 0 Then
                        rowSubTotal = tblSubTotal.DefaultView(0).Row
                        rowSubTotal.Qty += .Qty
                        If rowSubTotal.Max_DefPrec < ._Decimal Then
                            rowSubTotal.Max_DefPrec = ._Decimal
                        End If
                    Else
                        rowSubTotal = tblSubTotal.NewDataTable1Row
                        rowSubTotal.Primary_Unit = .Primary_Units
                        rowSubTotal.Unit_Code = .Unit_Code
                        rowSubTotal.Before_or_After = "after"
                        rowSubTotal.Qty = .Qty
                        rowSubTotal.Max_DefPrec = ._Decimal
                        rowSubTotal.Group = ""
                        tblSubTotal.Rows.Add(rowSubTotal)
                    End If
                End With
            Next
            tblSubTotal.DefaultView.RowFilter = "[Before or After] = 'after'"
            For i As Integer = 0 To tblSubTotal.DefaultView.Count - 1
                rowSubTotal = tblSubTotal.DefaultView(i).Row
                With rowSubTotal
                    If i < tblSubTotal.DefaultView.Count - 1 Then
                        subTotalAfter += ObjCD.DigitGrouping(.Qty, .Max_DefPrec) & " " & .Unit_Code & vbNewLine
                    Else
                        subTotalAfter += ObjCD.DigitGrouping(.Qty, .Max_DefPrec) & " " & .Unit_Code
                    End If
                End With
            Next
            Return subTotalAfter
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function getSubtotalBefore1_perGroup(ByVal TBLItem_Before As DseLaporanProd_List.TBLItem_BeforeDataTable, _
                                ByVal tblSubTotal As DseLaporanProd_List.DataTable1DataTable) As Boolean
        Dim rowItem_Before As DseLaporanProd_List.TBLItem_BeforeRow
        Dim rowSubTotal As DseLaporanProd_List.DataTable1Row
        Try
            For Each rowItem_Before In TBLItem_Before
                With rowItem_Before
                    tblSubTotal.DefaultView.RowFilter = "[Primary Unit] = " & .Primary_Units.ToString & "AND [Before or After] = 'before'" & "AND [Group] =" & "'" & .Group_By & "'"
                    If tblSubTotal.DefaultView.Count > 0 Then
                        rowSubTotal = tblSubTotal.DefaultView(0).Row
                        rowSubTotal.Qty += .Qty
                        If rowSubTotal.Max_DefPrec < ._Decimal Then
                            rowSubTotal.Max_DefPrec = ._Decimal
                        End If
                    Else
                        rowSubTotal = tblSubTotal.NewDataTable1Row
                        rowSubTotal.Primary_Unit = .Primary_Units
                        rowSubTotal.Unit_Code = .Unit_Code
                        rowSubTotal.Before_or_After = "before"
                        rowSubTotal.Qty = .Qty
                        rowSubTotal.Group = .Group_By
                        rowSubTotal.Max_DefPrec = ._Decimal
                        tblSubTotal.Rows.Add(rowSubTotal)
                    End If
                End With
            Next
            For Each rowItem_Before In TBLItem_Before
                With rowItem_Before
                    tblSubTotal.DefaultView.RowFilter = Nothing
                    tblSubTotal.DefaultView.RowFilter = "[Before or After] = 'before'" & "AND [Group] =" & "'" & .Group_By & "'"
                    For i As Integer = 0 To tblSubTotal.DefaultView.Count - 1
                        rowSubTotal = tblSubTotal.DefaultView(i).Row
                        With rowSubTotal
                            If i < tblSubTotal.DefaultView.Count - 1 Then
                                rowItem_Before.SubTotal_Pergroup += ObjCD.DigitGrouping(.Qty, .Max_DefPrec) & " " & .Unit_Code & vbNewLine
                            Else
                                rowItem_Before.SubTotal_Pergroup += ObjCD.DigitGrouping(.Qty, .Max_DefPrec) & " " & .Unit_Code
                            End If
                        End With
                    Next
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function getSubtotalAfter1_perGroup(ByVal TBLItem_After As DseLaporanProd_List.TBLItem_AfterDataTable, _
                              ByVal tblSubTotal As DseLaporanProd_List.DataTable1DataTable) As Boolean
        Dim rowItem_After As DseLaporanProd_List.TBLItem_AfterRow
        Dim rowSubTotal As DseLaporanProd_List.DataTable1Row
        Try
            For Each rowItem_After In TBLItem_After
                With rowItem_After
                    tblSubTotal.DefaultView.RowFilter = "[Primary Unit] = " & .Primary_Units.ToString & "AND [Before or After] = 'after'" & "AND [Group] =" & "'" & .Group_By & "'"
                    If tblSubTotal.DefaultView.Count > 0 Then
                        rowSubTotal = tblSubTotal.DefaultView(0).Row
                        rowSubTotal.Qty += .Qty
                        If rowSubTotal.Max_DefPrec < ._Decimal Then
                            rowSubTotal.Max_DefPrec = ._Decimal
                        End If
                    Else
                        rowSubTotal = tblSubTotal.NewDataTable1Row
                        rowSubTotal.Primary_Unit = .Primary_Units
                        rowSubTotal.Unit_Code = .Unit_Code
                        rowSubTotal.Before_or_After = "after"
                        rowSubTotal.Qty = .Qty
                        rowSubTotal.Max_DefPrec = ._Decimal
                        rowSubTotal.Group = .Group_By
                        tblSubTotal.Rows.Add(rowSubTotal)
                    End If
                End With
            Next
            For Each rowItem_After In TBLItem_After
                With rowItem_After
                    tblSubTotal.DefaultView.RowFilter = Nothing
                    tblSubTotal.DefaultView.RowFilter = "[Before or After] = 'after'" & "AND [Group] =" & "'" & .Group_By & "'"
                    For i As Integer = 0 To tblSubTotal.DefaultView.Count - 1
                        rowSubTotal = tblSubTotal.DefaultView(i).Row
                        With rowSubTotal
                            If i < tblSubTotal.DefaultView.Count - 1 Then
                                rowItem_After.SubTotal_Pergroup += ObjCD.DigitGrouping(.Qty, .Max_DefPrec) & " " & .Unit_Code & vbNewLine
                            Else
                                rowItem_After.SubTotal_Pergroup += ObjCD.DigitGrouping(.Qty, .Max_DefPrec) & " " & .Unit_Code
                            End If
                        End With
                    Next
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function FilterProduct() As String
        Dim Retval As String = ""
        Try
            If LvlGroup = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.ProductGroup Then
                Retval = "Primary Group"
            ElseIf LvlGroup = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.Brand Then
                Retval = "Primary Brand"
            ElseIf LvlGroup = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.SubBrand Then
                Retval = "Primary SubBrand"
            ElseIf LvlGroup = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.Category Then
                Retval = "Primary Category"
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function getSubtotal2_Before(ByVal TBLItem As DseLaporanProd_List.TblItem1DataTable, _
                                ByVal tblSubTotal As DseLaporanProd_List.DataTable1DataTable, ByVal subTotalBefore As String) As String
        Dim rowItem_Before As DseLaporanProd_List.TblItem1Row
        Dim rowSubTotal As DseLaporanProd_List.DataTable1Row
        Try
            TBLItem.DefaultView.RowFilter = "[Group Before After] = 'BEFORE'"
            For abad As Integer = 0 To TBLItem.DefaultView.Count - 1
                rowItem_Before = TBLItem.DefaultView(abad).Row
                With rowItem_Before
                    tblSubTotal.DefaultView.RowFilter = "[Primary Unit] = " & .Primary_Units.ToString & " AND [Before or After] = 'BEFORE'"
                    If tblSubTotal.DefaultView.Count > 0 Then
                        rowSubTotal = tblSubTotal.DefaultView(0).Row
                        rowSubTotal.Qty += .QtyDec
                        If rowSubTotal.Max_DefPrec < ._Decimal Then
                            rowSubTotal.Max_DefPrec = ._Decimal
                        End If
                    Else
                        rowSubTotal = tblSubTotal.NewDataTable1Row
                        rowSubTotal.Primary_Unit = .Primary_Units
                        rowSubTotal.Unit_Code = .Unit_Code
                        rowSubTotal.Before_or_After = "BEFORE"
                        rowSubTotal.Qty = .QtyDec
                        rowSubTotal.Group = .Group_Before_After
                        rowSubTotal.Max_DefPrec = ._Decimal
                        tblSubTotal.Rows.Add(rowSubTotal)
                    End If
                End With
            Next
            tblSubTotal.DefaultView.RowFilter = "[Before or After] = 'BEFORE'"
            For i As Integer = 0 To tblSubTotal.DefaultView.Count - 1
                rowSubTotal = tblSubTotal.DefaultView(i).Row
                With rowSubTotal
                    If i < tblSubTotal.DefaultView.Count - 1 Then
                        subTotalBefore += ObjCD.DigitGrouping(.Qty, .Max_DefPrec) & " " & .Unit_Code & vbNewLine
                    Else
                        subTotalBefore += ObjCD.DigitGrouping(.Qty, .Max_DefPrec) & " " & .Unit_Code
                    End If
                End With
            Next
            Return subTotalBefore
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function getSubtotal2_After(ByVal TBLItem As DseLaporanProd_List.TblItem1DataTable, _
                                ByVal tblSubTotal As DseLaporanProd_List.DataTable1DataTable, ByVal subTotalAfter As String) As String
        Dim rowItem_After As DseLaporanProd_List.TblItem1Row
        Dim rowSubTotal As DseLaporanProd_List.DataTable1Row
        Try
            TBLItem.DefaultView.RowFilter = "[Group Before After] = 'AFTER'"
            For abad As Integer = 0 To TBLItem.DefaultView.Count - 1
                rowItem_After = TBLItem.DefaultView(abad).Row
                If TBLItem.DefaultView.Count > 0 Then
                    With rowItem_After
                        tblSubTotal.DefaultView.RowFilter = "[Primary Unit] = " & .Primary_Units.ToString & " AND [Before or After] = 'AFTER'"
                        If tblSubTotal.DefaultView.Count > 0 Then
                            rowSubTotal = tblSubTotal.DefaultView(0).Row
                            rowSubTotal.Qty += .QtyDec
                            If rowSubTotal.Max_DefPrec < ._Decimal Then
                                rowSubTotal.Max_DefPrec = ._Decimal
                            End If
                        Else
                            rowSubTotal = tblSubTotal.NewDataTable1Row
                            rowSubTotal.Primary_Unit = .Primary_Units
                            rowSubTotal.Unit_Code = .Unit_Code
                            rowSubTotal.Before_or_After = "AFTER"
                            rowSubTotal.Qty = .QtyDec
                            rowSubTotal.Group = ""
                            rowSubTotal.Max_DefPrec = ._Decimal
                            tblSubTotal.Rows.Add(rowSubTotal)
                        End If
                    End With
                End If
            Next

            tblSubTotal.DefaultView.RowFilter = "[Before or After] = 'AFTER'"
            For i As Integer = 0 To tblSubTotal.DefaultView.Count - 1
                rowSubTotal = tblSubTotal.DefaultView(i).Row
                With rowSubTotal
                    If i < tblSubTotal.DefaultView.Count - 1 Then
                        subTotalAfter += ObjCD.DigitGrouping(.Qty, .Max_DefPrec) & " " & .Unit_Code & vbNewLine
                    Else
                        subTotalAfter += ObjCD.DigitGrouping(.Qty, .Max_DefPrec) & " " & .Unit_Code
                    End If
                End With
            Next
            Return subTotalAfter
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetSortingByGroup_Before(ByVal TBLItem_Before As DseLaporanProd_List.TBLItem_BeforeDataTable, _
    ByVal TBLItem_Before_Dummy As DseLaporanProd_List.TBLItem_BeforeDataTable) As Boolean

        Dim rowItem_Before As DseLaporanProd_List.TBLItem_BeforeRow
        Dim rowItem_Before_Dummy As DseLaporanProd_List.TBLItem_BeforeRow
        Try
            ''@ABAD BADRUDIN
            Dim SortType_ As String = "ASC"
            For Each RowDgv As DataGridViewRow In Me.dgvSort.Rows
                With RowDgv
                    If .Cells(Me.ColSort_Desc.Index).Value = Me.cboGroupBy.Text Then
                        If .Cells(Me.ColSort_Direction.Index).Value = "Descending" Then
                            SortType_ = "DESC"
                        End If
                    End If
                End With
            Next

            Dim Sort_ As String = ""
            Sort_ += "[Group By] " & SortType_ & ", "
            For i As Integer = 1 To 4
                For Each rowDgv As DataGridViewRow In Me.dgvSort.Rows
                    With rowDgv
                        If .Cells(Me.ColSort_Order.Index).Value = i.ToString Then
                            If (Me.cboGroupBy.Text = "Process Code" And .Cells(Me.ColSort_Desc.Index).Value = "Process") And _
                            (Me.cboGroupBy.Text = "Work Unit" And .Cells(Me.ColSort_Desc.Index).Value = "Work Unit") And _
                            (Me.cboGroupBy.Text = "Workshift" And .Cells(Me.ColSort_Desc.Index).Value = "Workshift") Then
                            Else
                                If .Cells(Me.ColSort_Direction.Index).Value = "Descending" Then
                                    SortType_ = "DESC"
                                Else
                                    SortType_ = "ASC"
                                End If
                                Select Case .Cells(Me.ColSort_Desc.Index).Value
                                    Case "Process"
                                        Sort_ += "[Process Code] " & SortType_ & ", "
                                        Exit Select
                                    Case "Work Unit"
                                        Sort_ += "[work unit code] " & SortType_ & ", "
                                        Exit Select
                                    Case "Workshift"
                                        Sort_ += "[workshift code] " & SortType_ & ", "
                                        Exit Select
                                    Case "Code No"
                                        Sort_ += "[Product Code] " & SortType_ & ", "
                                        Exit Select
                                End Select
                            End If
                        End If
                    End With
                Next
            Next

            Sort_ = Sort_.Substring(0, Sort_.Length - 2)
            For Each rowItem_Before In TBLItem_Before
                rowItem_Before_Dummy = TBLItem_Before_Dummy.NewTBLItem_BeforeRow
                For i As Integer = 0 To TBLItem_Before.Columns.Count - 1
                    rowItem_Before_Dummy.Item(i) = rowItem_Before.Item(i)
                Next
                TBLItem_Before_Dummy.Rows.Add(rowItem_Before_Dummy)
            Next

            TBLItem_Before.Rows.Clear()
            TBLItem_Before_Dummy.DefaultView.Sort = Sort_
            For i As Integer = 0 To TBLItem_Before_Dummy.DefaultView.Count - 1
                rowItem_Before_Dummy = TBLItem_Before_Dummy.DefaultView(i).Row
                rowItem_Before = TBLItem_Before.NewTBLItem_BeforeRow
                For i2 As Integer = 0 To TBLItem_Before_Dummy.Columns.Count - 1
                    rowItem_Before.Item(i2) = rowItem_Before_Dummy.Item(i2)
                Next
                TBLItem_Before.Rows.Add(rowItem_Before)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetSortingByGroup_After(ByVal TBLItem_After As DseLaporanProd_List.TBLItem_AfterDataTable, _
   ByVal TBLItem_After_Dummy As DseLaporanProd_List.TBLItem_AfterDataTable) As Boolean

        Dim rowItem_After As DseLaporanProd_List.TBLItem_AfterRow
        Dim rowItem_After_Dummy As DseLaporanProd_List.TBLItem_AfterRow
        Try
            ''@ABAD BADRUDIN
            Dim SortType_ As String = "ASC"
            For Each RowDgv As DataGridViewRow In Me.dgvSort.Rows
                With RowDgv
                    If .Cells(Me.ColSort_Desc.Index).Value = Me.cboGroupBy.Text Then
                        If .Cells(Me.ColSort_Direction.Index).Value = "Descending" Then
                            SortType_ = "DESC"
                        End If
                    End If
                End With
            Next

            Dim Sort_ As String = ""
            Sort_ += "[Group By] " & SortType_ & ", "
            For i As Integer = 1 To 4
                For Each rowDgv As DataGridViewRow In Me.dgvSort.Rows
                    With rowDgv
                        If .Cells(Me.ColSort_Order.Index).Value = i.ToString Then
                            If (Me.cboGroupBy.Text = "Process Code" And .Cells(Me.ColSort_Desc.Index).Value = "Process") And _
                            (Me.cboGroupBy.Text = "Work Unit" And .Cells(Me.ColSort_Desc.Index).Value = "Work Unit") And _
                            (Me.cboGroupBy.Text = "Workshift" And .Cells(Me.ColSort_Desc.Index).Value = "Workshift") Then
                            Else
                                If .Cells(Me.ColSort_Direction.Index).Value = "Descending" Then
                                    SortType_ = "DESC"
                                Else
                                    SortType_ = "ASC"
                                End If
                                Select Case .Cells(Me.ColSort_Desc.Index).Value
                                    Case "Process"
                                        Sort_ += "[Process Code] " & SortType_ & ", "
                                        Exit Select
                                    Case "Work Unit"
                                        Sort_ += "[work unit code] " & SortType_ & ", "
                                        Exit Select
                                    Case "Workshift"
                                        Sort_ += "[workshift code] " & SortType_ & ", "
                                        Exit Select
                                    Case "Code No"
                                        Sort_ += "[Product Code] " & SortType_ & ", "
                                        Exit Select
                                End Select
                            End If
                        End If
                    End With
                Next
            Next

            Sort_ = Sort_.Substring(0, Sort_.Length - 2)
            For Each rowItem_After In TBLItem_After
                rowItem_After_Dummy = TBLItem_After_Dummy.NewTBLItem_AfterRow
                For i As Integer = 0 To TBLItem_After.Columns.Count - 1
                    rowItem_After_Dummy.Item(i) = rowItem_After.Item(i)
                Next
                TBLItem_After_Dummy.Rows.Add(rowItem_After_Dummy)
            Next

            TBLItem_After.Rows.Clear()
            TBLItem_After_Dummy.DefaultView.Sort = Sort_
            For i As Integer = 0 To TBLItem_After_Dummy.DefaultView.Count - 1
                rowItem_After_Dummy = TBLItem_After_Dummy.DefaultView(i).Row
                rowItem_After = TBLItem_After.NewTBLItem_AfterRow
                For i2 As Integer = 0 To TBLItem_After_Dummy.Columns.Count - 1
                    rowItem_After.Item(i2) = rowItem_After_Dummy.Item(i2)
                Next
                TBLItem_After.Rows.Add(rowItem_After)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetSorting_Before(ByVal TBLItem_Before As DseLaporanProd_List.TBLItem_BeforeDataTable, _
        ByVal TBLItem_Before_Dummy As DseLaporanProd_List.TBLItem_BeforeDataTable) As Boolean
        Dim rowItem_Before As DseLaporanProd_List.TBLItem_BeforeRow
        Dim rowItem_Before_Dummy As DseLaporanProd_List.TBLItem_BeforeRow
        Try
            ''@ABAD BADRUDIN
            Dim SortType_ As String = "ASC"
            Dim Sort_ As String = ""
            For i As Integer = 1 To 4
                For Each rowDgv As DataGridViewRow In Me.dgvSort.Rows
                    With rowDgv
                        If .Cells(Me.ColSort_Order.Index).Value = i.ToString Then
                            If .Cells(Me.ColSort_Direction.Index).Value = "Descending" Then
                                SortType_ = "DESC"
                            Else
                                SortType_ = "ASC"
                            End If

                            If Me.Format_Index <> EnumFormat.Detail Then
                                Select Case .Cells(Me.ColSort_Desc.Index).Value
                                    Case "Process"
                                        Sort_ += "[Process Code] " & SortType_ & ", "
                                        Exit Select
                                    Case "Work Unit"
                                        Sort_ += "[work unit code] " & SortType_ & ", "
                                        Exit Select
                                    Case "Workshift"
                                        Sort_ += "[workshift code] " & SortType_ & ", "
                                        Exit Select
                                    Case "Code No"
                                        Sort_ += "[Product Code] " & SortType_ & ", "
                                        Exit Select
                                End Select
                            Else
                                Select Case .Cells(Me.ColSort_Desc.Index).Value
                                    Case "Date"
                                        Sort_ += "[date] " & SortType_ & ", "
                                        Exit Select
                                    Case "Document No"
                                        Sort_ += "[no_transaksi] " & SortType_ & ", "
                                        Exit Select
                                    Case "Process"
                                        Sort_ += "[Process Code] " & SortType_ & ", "
                                        Exit Select
                                    Case "Work Unit"
                                        Sort_ += "[work unit code] " & SortType_ & ", "
                                        Exit Select
                                    Case "Workshift"
                                        Sort_ += "[workshift code] " & SortType_ & ", "
                                        Exit Select
                                    Case "Code No"
                                        Sort_ += "[Product Code] " & SortType_ & ", "
                                        Exit Select
                                End Select
                            End If
                        End If
                    End With
                Next
            Next

            Sort_ = Sort_.Substring(0, Sort_.Length - 2)
            For Each rowItem_Before In TBLItem_Before
                rowItem_Before_Dummy = TBLItem_Before_Dummy.NewTBLItem_BeforeRow
                For i As Integer = 0 To TBLItem_Before.Columns.Count - 1
                    rowItem_Before_Dummy.Item(i) = rowItem_Before.Item(i)
                Next
                TBLItem_Before_Dummy.Rows.Add(rowItem_Before_Dummy)
            Next

            TBLItem_Before.Rows.Clear()
            TBLItem_Before_Dummy.DefaultView.Sort = Sort_
            For i As Integer = 0 To TBLItem_Before_Dummy.DefaultView.Count - 1
                rowItem_Before_Dummy = TBLItem_Before_Dummy.DefaultView(i).Row
                rowItem_Before = TBLItem_Before.NewTBLItem_BeforeRow
                For i2 As Integer = 0 To TBLItem_Before_Dummy.Columns.Count - 1
                    rowItem_Before.Item(i2) = rowItem_Before_Dummy.Item(i2)
                Next
                TBLItem_Before.Rows.Add(rowItem_Before)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetSorting_After(ByVal TBLItem_After As DseLaporanProd_List.TBLItem_AfterDataTable, _
      ByVal TBLItem_After_Dummy As DseLaporanProd_List.TBLItem_AfterDataTable) As Boolean
        Dim rowItem_After As DseLaporanProd_List.TBLItem_AfterRow
        Dim rowItem_After_Dummy As DseLaporanProd_List.TBLItem_AfterRow
        Try
            ''@ABAD BADRUDIN
            Dim SortType_ As String = "ASC"
            Dim Sort_ As String = ""
            For i As Integer = 1 To 4
                For Each rowDgv As DataGridViewRow In Me.dgvSort.Rows
                    With rowDgv
                        If .Cells(Me.ColSort_Order.Index).Value = i.ToString Then
                            If .Cells(Me.ColSort_Direction.Index).Value = "Descending" Then
                                SortType_ = "DESC"
                            Else
                                SortType_ = "ASC"
                            End If

                            If Me.Format_Index <> EnumFormat.Detail Then
                                Select Case .Cells(Me.ColSort_Desc.Index).Value
                                    Case "Process"
                                        Sort_ += "[Process Code] " & SortType_ & ", "
                                        Exit Select
                                    Case "Work Unit"
                                        Sort_ += "[work unit code] " & SortType_ & ", "
                                        Exit Select
                                    Case "Workshift"
                                        Sort_ += "[workshift code] " & SortType_ & ", "
                                        Exit Select
                                    Case "Code No"
                                        Sort_ += "[Product Code] " & SortType_ & ", "
                                        Exit Select
                                End Select
                            Else
                                Select Case .Cells(Me.ColSort_Desc.Index).Value
                                    Case "Date"
                                        Sort_ += "[date] " & SortType_ & ", "
                                        Exit Select
                                    Case "Document No"
                                        Sort_ += "[no_transaksi] " & SortType_ & ", "
                                        Exit Select
                                    Case "Process"
                                        Sort_ += "[Process Code] " & SortType_ & ", "
                                        Exit Select
                                    Case "Work Unit"
                                        Sort_ += "[work unit code] " & SortType_ & ", "
                                        Exit Select
                                    Case "Workshift"
                                        Sort_ += "[workshift code] " & SortType_ & ", "
                                        Exit Select
                                    Case "Code No"
                                        Sort_ += "[Product Code] " & SortType_ & ", "
                                        Exit Select
                                End Select
                            End If
                        End If
                    End With
                Next
            Next

            Sort_ = Sort_.Substring(0, Sort_.Length - 2)
            For Each rowItem_After In TBLItem_After
                rowItem_After_Dummy = TBLItem_After_Dummy.NewTBLItem_AfterRow
                For i As Integer = 0 To TBLItem_After.Columns.Count - 1
                    rowItem_After_Dummy.Item(i) = rowItem_After.Item(i)
                Next
                TBLItem_After_Dummy.Rows.Add(rowItem_After_Dummy)
            Next

            TBLItem_After.Rows.Clear()
            TBLItem_After_Dummy.DefaultView.Sort = Sort_
            For i As Integer = 0 To TBLItem_After_Dummy.DefaultView.Count - 1
                rowItem_After_Dummy = TBLItem_After_Dummy.DefaultView(i).Row
                rowItem_After = TBLItem_After.NewTBLItem_AfterRow
                For i2 As Integer = 0 To TBLItem_After_Dummy.Columns.Count - 1
                    rowItem_After.Item(i2) = rowItem_After_Dummy.Item(i2)
                Next
                TBLItem_After.Rows.Add(rowItem_After)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetFilter(ByVal vDgvItem As DataGridView, ByVal Primary_Filter As String, ByVal vCol_Primary As DataGridViewTextBoxColumn) As String
        Dim Retval_ As String = ""
        Dim row_ As DataGridViewRow
        Try
            vDgvItem.AllowUserToAddRows = False
            If vDgvItem.RowCount <> 0 Then
                If vDgvItem.RowCount > 1 Then
                    For Each row_ In vDgvItem.Rows
                        If row_.Index = 0 Then
                            Retval_ &= " AND (" & "[" & Primary_Filter & "] =" & row_.Cells(vCol_Primary.Index).Value
                        Else
                            If row_.Index = vDgvItem.RowCount - 1 Then
                                Retval_ &= " OR " & "[" & Primary_Filter & "] =" & row_.Cells(vCol_Primary.Index).Value & ")"
                            Else
                                Retval_ &= " OR " & "[" & Primary_Filter & "] =" & row_.Cells(vCol_Primary.Index).Value
                            End If
                        End If
                    Next
                Else
                    Retval_ &= " AND " & "[" & Primary_Filter & "] =" & vDgvItem.Item(vCol_Primary.Index, 0).Value
                End If
            End If
            vDgvItem.AllowUserToAddRows = True
            Return Retval_
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function
#End Region

End Class